// Concur selector map — the ONE file to maintain when Concur updates its UI.
// data-nuiexp is Concur's own test attribute (stable across tenants) — prefer it over ids or CSS paths.
// When Concur drifts, re-capture from a live expense form's field attributes and update this map (see README).

const CB_TEXT_FIELDS = {
  transactionDate:   'input[data-nuiexp="field-transactionDate"]',
  businessPurpose:   'input[data-nuiexp="field-businessPurpose"]',
  vendorName:        'input[data-nuiexp="field-vendorName"]',
  custom34:          'input[data-nuiexp="field-custom34"]',        // Business Name
  transactionAmount: 'input[data-nuiexp="field-transactionAmount"]',
  comment:           'textarea[data-nuiexp="field-comment"]'
};

const CB_CHECKBOXES = {
  isPersonalExpense: 'input[data-nuiexp="field-isPersonalExpense"]', // Personal Expense (do not reimburse)
  custom3:           'input[data-nuiexp="field-custom3"]'            // Missing Receipt Acknowledgment (expense form)
};

// Presence of this selector = "we are on an expense form" (New Expense AND the edit form of an
// existing/card-feed expense share the same fields — verified live 2026-07-05).
const CB_FORM_SENTINEL = CB_TEXT_FIELDS.transactionDate;

// Concur's receipt-upload file input (right-hand Receipt panel; accepts .png .jpg .jpeg .pdf .tif .tiff).
const CB_RECEIPT_INPUT = 'input[type="file"]';

// Search-combobox widgets (click to open → filtered option list → click option).
// `sel` = stable data-nuiexp root; `label` = fallback lookup by the field's visible label text
// (the Expense Type combobox ships without a data-nuiexp attribute).
const CB_COMBOS = {
  expenseType: { label: 'Expense Type' },                    // filled with "<code>-<name>", matched by code
  custom17:    { sel: '[data-nuiexp="field-custom17"]' },    // Is this a Vendor Invoice? → Yes/No
  location:    { sel: '[data-nuiexp="field-location"]' }     // City (server-searched list; best-effort)
};
