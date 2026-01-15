(async () => {
  /***********************
   * FULL CANVAS ROSTER SCRIPT
   * - Finds pending students
   * - Opens kebab → Edit Sections
   * - Removes alpha-starting section (e.g., "SPE ...")
   * - Saves
   * - Reopens → Edit Sections
   * - Types section name
   * - Forces dropdown open (Alt+ArrowDown / ArrowDown)
   * - Selects matching option from listbox
   * - Saves
   * - Continues + scrolls for more rows
   ***********************/

  /***********************
   * CONFIG
   ***********************/
  const DRY_RUN = false;      // true = do everything except click Save
  const SCROLL_PAUSE_MS = 900;

  /***********************
   * UTILITIES
   ***********************/
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  async function waitFor(getter, { timeout = 15000, interval = 50 } = {}) {
    const t0 = Date.now();
    while (Date.now() - t0 < timeout) {
      const v = getter();
      if (v) return v;
      await sleep(interval);
    }
    throw new Error("waitFor timeout");
  }

  async function waitGone(selector, { timeout = 15000 } = {}) {
    const t0 = Date.now();
    while (Date.now() - t0 < timeout) {
      if (!document.querySelector(selector)) return;
      await sleep(50);
    }
    throw new Error("waitGone timeout: " + selector);
  }

  function isVisible(el) {
    return !!el && el.offsetParent !== null;
  }

  function norm(s) {
    return (s || "").trim().replace(/\s+/g, " ");
  }

  function userClick(el) {
    // More reliable than el.click() for Canvas/React controls
    el.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent("mousedown",   { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent("mouseup",     { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent("click",       { bubbles: true, cancelable: true, view: window }));
  }

  function sendKey(el, { key, code, keyCode, which, altKey = false }) {
    const evtInit = {
      key,
      code,
      bubbles: true,
      cancelable: true,
      composed: true,
      altKey,
    };

    const down = new KeyboardEvent("keydown", evtInit);
    Object.defineProperty(down, "keyCode", { get: () => keyCode });
    Object.defineProperty(down, "which", { get: () => which });

    const up = new KeyboardEvent("keyup", evtInit);
    Object.defineProperty(up, "keyCode", { get: () => keyCode });
    Object.defineProperty(up, "which", { get: () => which });

    el.dispatchEvent(down);
    el.dispatchEvent(up);
  }

  /***********************
   * SELECTORS (grounded in your provided HTML)
   ***********************/
  const ROW_SEL = "tr.rosterUser";
  const PENDING_SEL = 'span.label.label-info[title*="not yet accepted the invitation"]';

  const KEBAB_SEL = "a.al-trigger.al-trigger-gray";
  const ROW_MENU_EDIT_SECTIONS_SEL = "ul.al-options a[data-event='editSections']";
  const POPUP_MENU_EDIT_SECTIONS_SEL = "ul.al-options[role='menu'] a[data-event='editSections']";

  const MODAL_SEL = "span[role='dialog'][aria-label='Edit Sections']";
  const INPUT_SEL = "input[data-testid='section-input']";
  const SAVE_SEL = "button[data-testid='save-button']";
  const USER_SECTIONS_LI_SEL = "#user_sections li";
  const REMOVE_BTN_SEL = "span[data-testid^='remove-section-'] button";

  /***********************
   * ROW HELPERS
   ***********************/
  function rowIsPending(row) {
    return !!row.querySelector(PENDING_SEL);
  }

  function alphaSectionFromRow(row) {
    // Picks first section cell text that starts with a letter (e.g., "SPE ...")
    const secs = [...row.querySelectorAll("td[data-testid='section-column-cell'] .section")]
      .map((n) => norm(n.textContent))
      .filter(Boolean);
    return secs.find((s) => /^[A-Za-z]/.test(s)) || null;
  }

  /***********************
   * MODAL / MENU ACTIONS
   ***********************/
  async function openEditSectionsModal(row) {
    const kebab = row.querySelector(KEBAB_SEL);
    if (!kebab) throw new Error("Kebab trigger not found");
    kebab.scrollIntoView({ block: "center" });
    userClick(kebab);
    await sleep(160);

    // Prefer the row's inline menu item, fallback to popup menu item
    const rowEdit = row.querySelector(ROW_MENU_EDIT_SECTIONS_SEL);
    if (rowEdit && isVisible(rowEdit)) {
      userClick(rowEdit);
    } else {
      const popupEdit = [...document.querySelectorAll(POPUP_MENU_EDIT_SECTIONS_SEL)].find(isVisible);
      if (!popupEdit) throw new Error("Edit Sections menu item not found (row or popup)");
      userClick(popupEdit);
    }

    const modal = await waitFor(() => document.querySelector(MODAL_SEL), { timeout: 15000 });

    // Kick focus out of any leftover menus (prevents weird aria-hidden focus warnings)
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    await sleep(80);

    return modal;
  }

  function modalHasSection(modal, sectionBase) {
    const want = sectionBase.toLowerCase();
    return [...modal.querySelectorAll(USER_SECTIONS_LI_SEL)].some((li) =>
      (li.textContent || "").toLowerCase().includes(want)
    );
  }

  async function removeSectionInModal(modal, sectionBase) {
    const want = sectionBase.toLowerCase();
    const li = [...modal.querySelectorAll(USER_SECTIONS_LI_SEL)].find((x) =>
      (x.textContent || "").toLowerCase().includes(want)
    );
    if (!li) return false;

    const btn = li.querySelector(REMOVE_BTN_SEL);
    if (!btn) throw new Error("Remove button not found for section: " + sectionBase);

    userClick(btn);
    await sleep(220);
    return true;
  }

  function getListboxForInput(input) {
    // input has aria-describedby="Selectable___X-description" → listbox id is "Selectable___X-list"
    const desc = input.getAttribute("aria-describedby");
    if (!desc) return null;
    const listId = desc.replace("-description", "-list");
    return document.getElementById(listId);
  }

  function optionLabel(optionEl) {
    // Section name appears inside option text
    return norm(optionEl.innerText || optionEl.textContent);
  }

  async function forceOpenSectionDropdown(input) {
    // Ensure the element is truly active
    input.focus();
    userClick(input);
    await sleep(90);

    // Many comboboxes open on Alt+ArrowDown
    for (let i = 0; i < 4; i++) {
      sendKey(input, { key: "ArrowDown", code: "ArrowDown", keyCode: 40, which: 40, altKey: true });
      await sleep(140);
      if (input.getAttribute("aria-expanded") === "true") return;
    }

    // Fallback: plain ArrowDown
    for (let i = 0; i < 4; i++) {
      sendKey(input, { key: "ArrowDown", code: "ArrowDown", keyCode: 40, which: 40, altKey: false });
      await sleep(140);
      if (input.getAttribute("aria-expanded") === "true") return;
    }
  }

  async function selectSectionFromDropdown(modal, sectionBase) {
    const input = modal.querySelector(INPUT_SEL);
    if (!input) throw new Error("Section input not found in modal");

    // Type the section name
    input.focus();
    input.value = "";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    await sleep(90);

    input.value = sectionBase;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    await sleep(180);

    // Force dropdown to appear (your manual behavior)
    await forceOpenSectionDropdown(input);

    // Wait for the listbox tied to this input to exist
    const listbox = await waitFor(() => {
      const lb = getListboxForInput(input);
      return lb && lb.getAttribute("role") === "listbox" ? lb : null;
    }, { timeout: 15000 });

    const options = [...listbox.querySelectorAll("span[role='option']")];
    if (!options.length) throw new Error("Listbox opened but has no options");

    const want = sectionBase.toLowerCase();

    // Find best match
    let target =
      options.find((o) => optionLabel(o).toLowerCase() === want) ||
      options.find((o) => optionLabel(o).toLowerCase().startsWith(want)) ||
      options.find((o) => optionLabel(o).toLowerCase().includes(want)) ||
      null;

    // If multiple similar names exist, pick shortest containing match as a heuristic
    if (!target) {
      const key = want.split(" ")[0]; // e.g., "spe"
      target =
        options
          .map((o) => ({ o, t: optionLabel(o).toLowerCase() }))
          .filter((x) => x.t.includes(key))
          .sort((a, b) => a.t.length - b.t.length)[0]?.o || options[0];
    }

    target.scrollIntoView({ block: "nearest" });

    // Important: realistic click
    userClick(target);
    await sleep(350);

    if (!modalHasSection(modal, sectionBase)) {
      // Helpful debugging output
      const sample = options.slice(0, 8).map(o => optionLabel(o)).join(" | ");
      throw new Error(
        `Selected an option but "${sectionBase}" did not appear in #user_sections. ` +
        `Top options were: ${sample}`
      );
    }
  }

  async function saveAndClose(modal) {
    if (DRY_RUN) {
      console.log("[DRY_RUN] Would click Save");
      // Optionally close modal manually if you want; DRY_RUN leaves it open
      return;
    }

    const save = modal.querySelector(SAVE_SEL);
    if (!save) throw new Error("Save button not found");
    userClick(save);

    await waitGone(MODAL_SEL, { timeout: 20000 });
    await sleep(220);
  }

  /***********************
   * PER-ROW PROCESS
   ***********************/
  async function processPendingRow(row) {
    const section = alphaSectionFromRow(row);
    if (!section) {
      console.warn("Pending row, but no alpha section found; skipping.");
      return;
    }

    console.log("Processing pending row. Section:", section);

    // 1) Remove section
    let modal = await openEditSectionsModal(row);
    await removeSectionInModal(modal, section);
    await saveAndClose(modal);

    // 2) Re-add section (must select from dropdown)
    modal = await openEditSectionsModal(row);
    await selectSectionFromDropdown(modal, section);
    await saveAndClose(modal);

    console.log("Done row:", section);
  }

  /***********************
   * MAIN LOOP: SCAN + SCROLL
   ***********************/
  const seen = new WeakSet();

  for (let loops = 0; loops < 500; loops++) {
    const rows = [...document.querySelectorAll(ROW_SEL)];
    let didWorkThisPass = false;

    for (const row of rows) {
      if (seen.has(row)) continue;
      seen.add(row);

      if (rowIsPending(row)) {
        didWorkThisPass = true;
        try {
          await processPendingRow(row);
        } catch (e) {
          console.error("Failed row:", e);
        }
      }
    }

    // Scroll to load more rows
    const before = document.documentElement.scrollHeight;
    window.scrollBy(0, Math.round(window.innerHeight * 0.85));
    await sleep(SCROLL_PAUSE_MS);
    const after = document.documentElement.scrollHeight;

    const atBottom = window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 2;
    if (after === before && atBottom && !didWorkThisPass) break;
  }

  console.log("All done.");
})();
