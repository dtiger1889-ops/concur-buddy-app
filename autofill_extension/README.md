# Concur Buddy Autofill — Chrome extension (MVP v0.1)

Fills SAP Concur's New Expense form from Concur Buddy's staged data. Human-in-the-loop by
design: you pick the expense, you click Fill, you review, **you** click Save Expense — the
extension never saves, submits, or acts on its own. Works because it runs inside your
already-logged-in Concur session (no passwords, no SSO/Duo handling anywhere).

## Install (once per PC — both home and work PC passed the policy probe)
1. Pull the repo (or copy this `autofill_extension/` folder).
2. `chrome://extensions` → toggle **Developer mode** → **Load unpacked** → select this folder.
3. **After every repo update:** `chrome://extensions` → ↻ **Reload** on the extension card.
   The popup header shows the loaded version — confirm it matches `manifest.json`.

## How it works (v0.3 architecture — matters for debugging)
No content script. The popup injects the fill function **on demand into the page's MAIN
world** via `chrome.scripting.executeScript` — the exact mechanism verified live against a
a real Concur expense. v0.1/0.2 used a content script instead; that ran in
Chrome's isolated world and wasn't injected into already-open tabs, which made Fill a silent
no-op — the bug fixed by this design. The Concur tab never needs reloading.

## Use — the P-card flow (the normal case)
Corporate-card purchases ALREADY exist in Concur (they arrive in **Available Expenses** a few
days after purchase, pre-filled with date/vendor/amount/payment type — all locked). The work
Concur Buddy automates is the coding + receipt, which Concur only allows once the expense is
in a report:
1. In Concur Buddy: stage the expense the day you buy (that's the whole point — Concur Buddy is
   your immediate record while Concur lags), attach the receipt file, mark **Ready to file**.
2. When the card transactions appear in Concur: put them in a report (checkbox in Available
   Expenses → Move to report, or in-report **Add Expense ▾ → Select from Available Expenses**).
3. Export from Concur Buddy (**More ▾ → Export for Autofill…**; receipts ride along — 📎 in the
   popup list). Open the card expense in the report (its edit form).
4. Extension icon → **Fill** on the matching expense. It fills (v0.4): the **Expense Type
   dropdown** (searched by code — done FIRST since changing type re-renders the form), the
   editable text fields (business purpose, business name, comment), the checkboxes, the
   **"Is this a Vendor Invoice?" dropdown** (No unless staged as an invoice), a best-effort
   **City** pick, and **drops the embedded receipt into the upload panel** (a real upload
   starts; confirm it renders — Remove + Available Receipts → 🗑 undoes a wrong one).
   Card-locked fields (date, vendor, amount) are skipped and reported as "locked"; fields you
   left blank in Concur Buddy are reported as "empty in Concur Buddy" rather than silently skipped.
5. Review, set whatever it lists as still yours (attendees, oddball dropdowns), **Save
   Expense**, then **Mark Filed** back in Concur Buddy.

## Use — the personal-card flow (occasional; e.g. a delegate's out-of-pocket reimbursement)
Same, except the expense doesn't come from the card feed: **Add Expense ▾ → Manually Create
Expense** → pick the type (shown next to each expense in the popup) → Fill works on the blank
New Expense form too, including date/vendor/amount since nothing is locked.

## When Concur updates its UI and Fill stops finding fields
The popup will say "Not found on page: …". All selectors live in **`selectors.js`** (one file, loaded by the popup).
Re-capture by opening a Concur expense form, inspecting the field you need in your browser's DevTools, and
reading its `data-nuiexp="field-…"` attribute (Concur's own stable test id). Update the matching entry in
`selectors.js` and reload the extension.

## Verified live (2026-07-05, on a real card expense in an open report)
- The React-safe fill sticks (char counters update = component state, not just pixels).
- Receipt drop via DataTransfer + change event **actually uploads** and renders in the viewer.
  (A trivially small/blank test image was rejected with "failed to upload" — real receipt files
  are fine.) Remove moves a receipt to Available Receipts; the pool's 🗑 deletes it for good.
- The edit form prompts Save & Continue / Continue Without Saving on unsaved changes; the
  New Expense form discards silently.

## Roadmap (not yet built)
- Auto-matching: read the report/Available Expenses rows and pair them with staged expenses by
  vendor + amount + date, instead of the human picking from the popup list.
- Native-messaging bridge: pull staged expenses straight from Concur Buddy's DB and write
  "Filled" status back (see `Apps/resources/browser_autofill_bridge_pattern.md`, Tier 2).
- Per-type field maps (airfare: ticket number, class of service — already captured).
