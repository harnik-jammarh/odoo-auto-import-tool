// =============================================================================
// Odoo Auto-Import — engine (ported from the Chrome extension's popup.js)
// Same schemas / fuzzy matching / cleaning / upload logic as the extension.
// The ONLY behavioural change: odooCall() now hits our own /api/odoo-proxy
// server route (server-to-server, no CORS issue) instead of fetching Odoo
// directly from the browser. Everything else — field aliases, dedupe keys,
// GST tax resolution, attribute parsing, account auto-creation, relocation,
// the "smart" duplicate-transaction check — is unchanged.
// =============================================================================
import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// Odoo schema library (unchanged from the extension)
// ---------------------------------------------------------------------------
export const ODOO_SCHEMAS = {
  contacts: {
    label: "Contacts", model: "res.partner", icon: "👤",
    note: "Import screen: Contacts app -> list view -> Action (gear icon) -> Import records.",
    dedupeKeys: [["email"], ["name"]],
    fields: [
      { name: "name", label: "Name", aliases: ["name", "full name", "contact name", "party name", "customer name", "vendor name"], required: true },
      { name: "email", label: "Email", aliases: ["email", "email address", "e-mail"] },
      { name: "phone", label: "Phone", aliases: ["phone", "phone number", "telephone", "landline"] },
      { name: "mobile", label: "Mobile", aliases: ["mobile", "mobile no", "mobile number", "cell", "contact no"] },
      { name: "street", label: "Street", aliases: ["street", "address", "address line 1", "addr1"] },
      { name: "street2", label: "Street 2", aliases: ["street2", "address line 2", "addr2"] },
      { name: "city", label: "City", aliases: ["city", "town"] },
      { name: "zip", label: "Zip", aliases: ["zip", "pincode", "postal code", "pin code"] },
      { name: "state_id", label: "State", aliases: ["state", "province"], relation: true, relModel: "res.country.state", autoCreate: false },
      { name: "country_id", label: "Country", aliases: ["country"], relation: true, relModel: "res.country", autoCreate: false },
      { name: "vat", label: "Tax ID / GSTIN", aliases: ["vat", "gstin", "gst no", "tax id", "pan"] },
      { name: "website", label: "Website", aliases: ["website", "web", "url"] },
      { name: "is_company", label: "Is a company", aliases: ["is company", "company", "type"], boolean: true },
      { name: "parent_id", label: "Related Company", aliases: ["parent company", "related company", "company name"], relation: true, relModel: "res.partner", autoCreate: true, autoCreateExtra: { is_company: true } },
      { name: "category_id", label: "Tags", aliases: ["tags", "category", "contact tags"], relation: true, relModel: "res.partner.category", autoCreate: true, isTagField: true },
      { name: "ref", label: "Reference", aliases: ["reference", "customer ref", "old id", "legacy id"] },
    ],
  },
  inventory: {
    label: "Inventory", model: "product.template", icon: "📦",
    note: "Import screen: Inventory/Sales app -> Products -> Action -> Import records. Opening stock qty is set separately via an Inventory Adjustment, not on the product import. This module now also covers the product's Attributes & Variants, Sales, and Purchase tabs — map the extra columns below like any other field.",
    dedupeKeys: [["default_code"], ["barcode"], ["name"]],
    fields: [
      { name: "name", label: "Product Name", aliases: ["name", "product name", "item name", "description"], required: true },
      { name: "default_code", label: "Internal Reference", aliases: ["sku", "internal reference", "item code", "product code", "ref"] },
      { name: "barcode", label: "Barcode", aliases: ["barcode", "ean", "upc"] },
      { name: "list_price", label: "Sales Price", aliases: ["sales price", "selling price", "mrp", "rate", "price"], number: true },
      { name: "standard_price", label: "Cost", aliases: ["cost", "cost price", "purchase price", "buying price"], number: true },
      { name: "categ_id", label: "Product Category", aliases: ["category", "product category", "group"], relation: true, relModel: "product.category", autoCreate: true },
      { name: "type", label: "Product Type", aliases: ["type", "product type", "goods/service"], selection: "productType" },
      { name: "uom_id", label: "Unit of Measure", aliases: ["uom", "unit", "unit of measure"], relation: true, relModel: "uom.uom", autoCreate: false },
      { name: "uom_po_id", label: "Purchase UoM", aliases: ["purchase uom", "purchase unit"], relation: true, relModel: "uom.uom", autoCreate: false },
      { name: "sale_ok", label: "Can be Sold", aliases: ["can be sold", "sales product"], boolean: true },
      { name: "purchase_ok", label: "Can be Purchased", aliases: ["can be purchased", "purchase product"], boolean: true },
      { name: "is_storable", label: "Track Inventory", aliases: ["track inventory", "by quantity", "storable", "manage stock"], boolean: true },
      { name: "attribute_line_ids", label: "Attributes & Variants", aliases: ["attributes", "attribute", "variants", "attribute values", "attributes & variants"], isAttributeLine: true },
      { name: "description_sale", label: "Quotation Description", aliases: ["quotation description", "sales description", "sale description", "sales note", "description for customers"] },
      { name: "invoice_policy", label: "Invoicing Policy", aliases: ["invoicing policy", "invoice policy"], selection: "invoicePolicy" },
      { name: "product_tag_ids", label: "Tags", aliases: ["tags", "product tags", "extra info tags", "sales tags"], relation: true, relModel: "product.tag", autoCreate: true, isTagField: true },
      { name: "description_purchase", label: "Purchase Description", aliases: ["purchase description", "purchase note", "description for vendors"] },
      { name: "purchase_method", label: "Control Policy", aliases: ["control policy", "purchase method", "bill control"], selection: "purchaseMethod" },
      { name: "seller_ids/partner_id", label: "Vendor", aliases: ["vendor", "vendor name", "supplier", "supplier name"], relation: true, relModel: "res.partner", autoCreate: true, isVendorLine: true },
      { name: "seller_ids/price", label: "Vendor Price", aliases: ["vendor price", "unit price", "supplier price", "purchase price from vendor"], number: true, isVendorLine: true },
      { name: "seller_ids/min_qty", label: "Vendor Quantity", aliases: ["quantity", "vendor quantity", "minimum quantity"], number: true, isVendorLine: true },
      { name: "seller_ids/delay", label: "Lead Time", aliases: ["lead time", "delivery lead time", "vendor lead time"], number: true, isVendorLine: true },
      { name: "weight", label: "Weight", aliases: ["weight"], number: true },
      { name: "volume", label: "Volume", aliases: ["volume"], number: true },
      { name: "qty_on_hand", label: "On Hand Quantity", aliases: ["on hand qty", "on hand quantity", "quantity on hand", "opening stock", "stock qty", "quantity"], number: true, isQuantity: true },
      { name: "relocate_from", label: "Relocate From", aliases: ["from location", "source location", "relocate from", "move from"], isRelocateFrom: true },
      { name: "relocate_to", label: "Relocate To", aliases: ["to location", "destination location", "relocate to", "move to"], isRelocateTo: true },
      { name: "relocate_qty", label: "Relocate Quantity (optional — blank = full qty)", aliases: ["relocate qty", "relocate quantity", "qty to move", "quantity to move", "move quantity"], number: true, isRelocateQty: true },
      { name: "property_account_income_id", label: "Income Account", aliases: ["income account", "sales account"], relation: true, relModel: "account.account", autoCreate: false },
      { name: "property_account_expense_id", label: "Expense Account", aliases: ["expense account", "purchase account"], relation: true, relModel: "account.account", autoCreate: false },
    ],
  },
  accounting: {
    label: "Accounting", model: "account.move", icon: "🧮",
    note: "Maps to Journal Entries / Invoices / Bills. Journal and Account must already exist in Odoo (auto-creating these could corrupt your books) — set them up once in Accounting first if missing.",
    dedupeKeys: [["partner_id", "ref"]],
    fields: [
      { name: "partner_id", label: "Partner", aliases: ["party", "party name", "partner", "customer/vendor", "customer", "vendor"], required: true, relation: true, relModel: "res.partner", autoCreate: true },
      { name: "journal_id", label: "Journal", aliases: ["journal", "voucher type"], required: true, relation: true, relModel: "account.journal", autoCreate: false },
      { name: "invoice_date", label: "Invoice/Bill Date", aliases: ["date", "invoice date", "bill date", "transaction date", "entry date"], date: true },
      { name: "ref", label: "Reference", aliases: ["reference", "voucher no", "bill no", "invoice no"] },
      { name: "invoice_line_ids/product_id", label: "Line / Product", aliases: ["product", "item", "particulars"], relation: true, relModel: "product.product", autoCreate: false, isLine: true },
      { name: "invoice_line_ids/quantity", label: "Line / Quantity", aliases: ["qty", "quantity"], number: true, isLine: true },
      { name: "invoice_line_ids/price_unit", label: "Line / Unit Price", aliases: ["unit price", "rate", "amount"], number: true, isLine: true },
      { name: "invoice_line_ids/account_id", label: "Line / Account", aliases: ["account", "ledger", "account code", "account name"], relation: true, relModel: "account.account", autoCreate: false, isLine: true },
    ],
  },
  openingBalance: {
    label: "Opening Balances", model: "account.move", icon: "💰",
    note: "For a raw debit/credit journal entry sheet like the Debtors/Creditors/Balance Sheet opening balance templates — one row per journal line, with every line for the same entry sharing the same Journal Entry Reference value. If an Account name isn't found in your Chart of Accounts, it's created automatically — review newly created accounts afterward. Not for invoices/bills (use Accounting for those).",
    dedupeKeys: [["ref"]],
    defaults: { move_type: "entry" },
    fields: [
      { name: "__group", label: "Journal Entry Reference", aliases: ["journal entry", "journal entry reference", "entry reference", "move reference"], required: true, isGroupKey: true },
      { name: "__row_id", label: "Row ID (not imported)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "date", label: "Date", aliases: ["date"], date: true },
      { name: "journal_id", label: "Journal", aliases: ["journal"], relation: true, relModel: "account.journal", autoCreate: false },
      { name: "line_ids/account_id", label: "Line / Account", aliases: ["account", "ledger", "account code", "account name"], relation: true, relModel: "account.account", isLine: true, isAccountField: true, isLineRequired: true },
      { name: "line_ids/partner_id", label: "Line / Partner", aliases: ["partner", "customer/vendor", "customer", "vendor"], relation: true, relModel: "res.partner", autoCreate: true, isLine: true },
      { name: "line_ids/name", label: "Line / Label", aliases: ["label", "description", "narration"], isLine: true },
      { name: "line_ids/debit", label: "Line / Debit", aliases: ["debit"], number: true, isLine: true, isDebitField: true },
      { name: "line_ids/credit", label: "Line / Credit", aliases: ["credit"], number: true, isLine: true, isCreditField: true },
    ],
  },
  sales: {
    label: "Sales", model: "sale.order", icon: "🛒",
    note: "Import screen: Sales app -> Orders -> Action -> Import records. Multi-line orders: rows below the first with a blank Customer are treated as extra lines of the same order. Your sheet's own order number/reference is kept on the order as the Customer/Order Reference field and is used to detect a re-import.",
    dedupeKeys: [["partner_id", "client_order_ref"]],
    fields: [
      { name: "partner_id", label: "Customer", aliases: ["customer", "customer name", "party", "buyer"], required: true, relation: true, relModel: "res.partner", autoCreate: true },
      { name: "date_order", label: "Order Date", aliases: ["order date", "date", "invoice date", "sale date"], date: true },
      { name: "order_line/product_id", label: "Line / Product", aliases: ["product", "item", "product name"], relation: true, relModel: "product.product", autoCreate: true, confirmCreate: true, isLine: true },
      { name: "order_line/product_uom_qty", label: "Line / Quantity", aliases: ["qty", "quantity", "units"], number: true, isLine: true },
      { name: "order_line/price_unit", label: "Line / Unit Price", aliases: ["unit price", "rate", "price", "selling price"], number: true, isLine: true },
      { name: "order_line/tax_ids", label: "Line / Taxes", aliases: ["taxes", "tax", "gst", "gst rate", "order lines/taxes"], isLine: true, isGstTaxField: "sale" },
      { name: "client_order_ref", label: "Order Reference (your file's number)", aliases: ["order no", "reference", "order reference", "customer reference", "po number", "customer po"] },
    ],
  },
  purchaseOrders: {
    label: "Purchase Orders", model: "purchase.order", icon: "📦",
    note: "Import screen: Purchase app -> Orders -> Action -> Import records. Multi-line orders: rows below the first with a blank Vendor are treated as extra lines of the same order. Type Of Purchase maps to the Studio field (values must read exactly \"RM Purchase\", \"Store Purchase\", or \"Service Purchase\").",
    dedupeKeys: [["partner_id", "partner_ref"]],
    fields: [
      { name: "partner_id", label: "Vendor", aliases: ["vendor", "vendor name", "supplier", "supplier name"], required: true, relation: true, relModel: "res.partner", autoCreate: true },
      { name: "date_order", label: "Order Date", aliases: ["order date", "date", "bill date", "purchase date"], date: true },
      { name: "order_line/product_id", label: "Line / Product", aliases: ["product", "item", "product name"], relation: true, relModel: "product.product", autoCreate: true, confirmCreate: true, isLine: true },
      { name: "order_line/product_qty", label: "Line / Quantity", aliases: ["qty", "quantity", "units"], number: true, isLine: true },
      { name: "order_line/price_unit", label: "Line / Unit Price", aliases: ["unit price", "rate", "price", "purchase price"], number: true, isLine: true },
      { name: "order_line/tax_ids", label: "Line / Taxes", aliases: ["taxes", "tax", "gst", "gst rate", "order lines/taxes"], isLine: true, isGstTaxField: "purchase" },
      { name: "partner_ref", label: "Order Reference (your file's number)", aliases: ["order no", "reference", "order reference", "vendor reference", "po number", "vendor po"] },
      { name: "x_studio_type_of_purchase", label: "Purchase Type", aliases: ["type of purchase", "purchase type", "type"], selection: "purchaseType" },
    ],
  },
  leads: {
    label: "Leads", model: "crm.lead", icon: "🎯",
    note: "Import screen: CRM app -> Leads -> Action -> Import records. If Leads isn't enabled as a separate menu, this creates records straight into Pipeline as leads.",
    dedupeKeys: [["email_from", "contact_name"], ["email_from", "partner_name"], ["name"]],
    fields: [
      { name: "name", label: "Name", aliases: ["name", "title", "opportunity", "lead name", "subject", "lead title"], required: true },
      { name: "partner_name", label: "Company Name", aliases: ["company name", "company", "organization", "organisation"] },
      { name: "contact_name", label: "Contact Name", aliases: ["contact name", "person name", "full name"] },
      { name: "email_from", label: "Email", aliases: ["email", "email address", "e-mail"] },
      { name: "function", label: "Job Position", aliases: ["job position", "designation", "position", "job title"] },
      { name: "phone", label: "Phone", aliases: ["phone", "phone number", "telephone", "landline"] },
      { name: "mobile", label: "Mobile", aliases: ["mobile", "mobile no", "mobile number", "cell"] },
      { name: "street", label: "Street", aliases: ["street", "address", "address line 1", "addr1"] },
      { name: "street2", label: "Street2", aliases: ["street2", "street 2", "address line 2", "addr2"] },
      { name: "city", label: "City", aliases: ["city", "town"] },
      { name: "state_id", label: "State", aliases: ["state", "province"], relation: true, relModel: "res.country.state", autoCreate: false },
      { name: "zip", label: "Zip", aliases: ["zip", "pincode", "postal code", "pin code"] },
      { name: "country_id", label: "Country", aliases: ["country"], relation: true, relModel: "res.country", autoCreate: false },
      { name: "website", label: "Website", aliases: ["website", "web", "url"] },
      { name: "description", label: "Notes", aliases: ["notes", "description", "remarks", "comments"] },
    ],
  },
  physicalInventory: {
    label: "Physical Inventory", model: "stock.quant", icon: "🧮",
    note: "Import screen: Inventory app -> Operations -> Physical Inventory -> gear icon -> Import records (or just use this module directly, without touching Odoo's own wizard). Each row sets the on-hand quantity for one Product at one Location, exactly like a Physical Inventory count + Apply. Location and Product must already exist in Odoo — neither is auto-created here, since a stock count should always point at real, already-decided locations/products (map a different column, or fix the name, if one isn't found). This module writes straight to stock.quant and always forces the true total at that Location to equal your Counted Quantity — including summing up and correcting for any duplicate quant rows Odoo may already have for that same Product+Location — so a re-run (or a product that already has more than one quant row there) can never silently double the count.",
    dedupeKeys: null,
    fields: [
      { name: "location_id", label: "Location", aliases: ["location", "location name", "warehouse location", "bin", "stock location"], required: true, relation: true, relModel: "stock.location", autoCreate: false, isLocationRelation: true },
      { name: "product_id", label: "Product", aliases: ["product", "product name", "item", "item name"], required: true, relation: true, relModel: "product.product", autoCreate: false },
      { name: "counted_qty", label: "Counted Quantity", aliases: ["counted quantity", "quantity", "qty", "counted qty", "count", "on hand", "on hand quantity"], required: true, number: true, isPhysicalInventoryQty: true },
    ],
  },
};


const PRODUCT_TYPE_MAP = { goods: "consu", product: "consu", "physical product": "consu", consu: "consu", service: "service", services: "service" };
const INVOICE_POLICY_MAP = { order: "order", "ordered quantities": "order", ordered: "order", delivery: "delivery", "delivered quantities": "delivery", delivered: "delivery" };
const PURCHASE_METHOD_MAP = { receive: "receive", "received quantities": "receive", received: "receive", order: "purchase", "ordered quantities": "purchase", ordered: "purchase", purchase: "purchase" };
const PURCHASE_TYPE_MAP = {
  "rm purchase": "RM Purchase", rm: "RM Purchase", "raw material": "RM Purchase", "raw material purchase": "RM Purchase",
  "store purchase": "Store Purchase", store: "Store Purchase", stock: "Store Purchase", "stock purchase": "Store Purchase",
  "service purchase": "Service Purchase", service: "Service Purchase", services: "Service Purchase",
};

// ---------------------------------------------------------------------------
// Fuzzy matching (pure functions, unchanged)
// ---------------------------------------------------------------------------
function normalize(s) { return String(s || "").toLowerCase().trim().replace(/[_\-.]+/g, " ").replace(/\s+/g, " "); }
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
  return dp[m][n];
}
function similarity(a, b) {
  const na = normalize(a), nb = normalize(b);
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  if (na.includes(nb) || nb.includes(na)) return 0.85;
  const dist = levenshtein(na, nb);
  return Math.max(0, 1 - dist / Math.max(na.length, nb.length));
}
function bestFieldScore(header, field) {
  let best = similarity(header, field.label);
  for (const alias of field.aliases) best = Math.max(best, similarity(header, alias));
  return best;
}
export function mapForModule(headers, moduleKey) {
  const schema = ODOO_SCHEMAS[moduleKey];
  const used = new Set();
  return headers.map((h) => {
    let best = { field: null, score: 0 };
    for (const field of schema.fields) {
      if (used.has(field.name)) continue;
      const score = bestFieldScore(h, field);
      if (score > best.score) best = { field, score };
    }
    if (best.field && best.score >= 0.55) used.add(best.field.name);
    return { header: h, field: best.score >= 0.55 ? best.field : null, confidence: best.score };
  });
}
export function analyzeSheet(headers) {
  let bestModule = null, bestModuleScore = -1, bestMapping = null;
  for (const key of Object.keys(ODOO_SCHEMAS)) {
    const mapping = mapForModule(headers, key);
    const schema = ODOO_SCHEMAS[key];
    const avg = mapping.reduce((s, m) => s + m.confidence, 0) / headers.length;
    const requiredCovered = schema.fields.filter((f) => f.required).every((f) => mapping.some((m) => m.field?.name === f.name));
    const moduleScore = avg + (requiredCovered ? 0.15 : 0);
    if (moduleScore > bestModuleScore) { bestModuleScore = moduleScore; bestModule = key; bestMapping = mapping; }
  }
  return { moduleKey: bestModule, moduleScore: bestModuleScore, mapping: bestMapping };
}
export function confidenceColor(score) {
  if (score >= 0.8) return "#0F6E56";
  if (score >= 0.55) return "#854F0B";
  return "#A32D2D";
}

// ---------------------------------------------------------------------------
// Value cleaning helpers (unchanged)
// ---------------------------------------------------------------------------
function cleanNumber(val) {
  if (typeof val === "number") return val;
  const cleaned = String(val).replace(/[₹$,\s]/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}
function cleanBoolean(val) {
  if (typeof val === "boolean") return val;
  const s = String(val).trim().toLowerCase();
  return ["yes", "true", "1", "y"].includes(s);
}
const MONTH_NAMES = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
function cleanDate(val) {
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
    return val;
  }
  if (val instanceof Date && !isNaN(val)) {
    return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, "0")}-${String(val.getDate()).padStart(2, "0")}`;
  }
  const s = String(val || "").trim();
  if (!s) return val;
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[ T].*)?$/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})[\s\-\/]+([A-Za-z]{3,9})[\s\-\/,]+(\d{4})$/);
  if (m && MONTH_NAMES[m[2].slice(0, 3).toLowerCase()]) {
    return `${m[3]}-${String(MONTH_NAMES[m[2].slice(0, 3).toLowerCase()]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;
  }
  m = s.match(/^([A-Za-z]{3,9})[\s\-\/]+(\d{1,2}),?[\s\-\/]+(\d{4})$/);
  if (m && MONTH_NAMES[m[1].slice(0, 3).toLowerCase()]) {
    return `${m[3]}-${String(MONTH_NAMES[m[1].slice(0, 3).toLowerCase()]).padStart(2, "0")}-${String(m[2]).padStart(2, "0")}`;
  }
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let [, a, b, y] = m;
    a = parseInt(a, 10); b = parseInt(b, 10);
    if (y.length === 2) y = (parseInt(y, 10) < 50 ? "20" : "19") + y;
    let day, month;
    if (a > 12 && b <= 12) { day = a; month = b; }
    else if (b > 12 && a <= 12) { day = b; month = a; }
    else { day = a; month = b; }
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${y}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }
  return val;
}
function cleanSelection(val, map) { const key = String(val).trim().toLowerCase(); return map[key] || val; }
function cleanText(val) { return typeof val === "string" ? val.trim().replace(/\s+/g, " ") : val; }
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
function isValidEmail(val) { return EMAIL_RE.test(String(val).trim()); }
function isRowEmpty(row, headers) { return headers.every((h) => row[h] === "" || row[h] == null || String(row[h]).trim() === ""); }

// ---------------------------------------------------------------------------
// Account auto-classification (pure, unchanged)
// ---------------------------------------------------------------------------
const PRE_SIDE_RULES = [
  [/opening balance equity|profit\s+(or|and)\s+loss appropriation|\bsuspense\b/i, "equity"],
  [/\breserve\b/i, "equity_unaffected"],
  [/share capital|capital account|equity share/i, "equity"],
  [/profit\s*&?\s*loss|p\s*&\s*l\b|retained earning/i, "equity_unaffected"],
  [/overdraft|\bo\.?d\.?\s*a\/?c\b|\bbank\s*od\b|\bcash credit\b|\bcc a\/?c\b/i, "liability_current"],
  [/^sales\b|^revenue\b|^income\b/i, "income"],
  [/^purchase\b|^expense\b|^cost of/i, "expense"],
];
const ASSET_SIDE_RULES = [
  [/\bdebtors?\b/i, "asset_receivable"],
  [/(trade|customer).*receivable|receivable.*(trade|customer)/i, "asset_receivable"],
  [/\breceivable\b/i, "asset_current"],
  [/\bcash in hand\b|\bcash\b/i, "asset_cash"],
  [/\bbank\b/i, "asset_cash"],
  [/prepaid/i, "asset_prepayments"],
  [/deposit/i, "asset_non_current"],
  [/depreciation|depriciation/i, "asset_fixed"],
  [/machine|m\/c\b|plant\s*&|plant\s+mach|equipment|furniture|\bbuilding\b|\bvehicle\b|\bcar\b|\bactiva\b|\bscooter\b|\bmotorcycle\b|\bbike\b|\btruck\b|\btempo\b|\btractor\b|extruder|shredder|kneader|pulveriser|pulverizer|\bdryer\b|drayer|compressor|\btower\b|installation|tubewell|\bsolar\b|hydrant|\blift\b|\bcomputer\b|\bland\b|c\.?w\.?i\.?p\.?|rou assets?|lab equipment/i, "asset_fixed"],
  [/raw material|finished good|work in progress|\bstores\b|\bspares\b/i, "asset_current"],
  [/mat credit|tcs on purchase|margin money/i, "asset_current"],
];
const ASSET_DEFAULT = "asset_current";
const LIABILITY_SIDE_RULES = [
  [/\bpayable\b|\bcreditors?\b/i, "liability_payable"],
  [/credit card/i, "liability_credit_card"],
  [/\bloan\b|term loan|lease liability/i, "liability_non_current"],
  [/deposit/i, "liability_non_current"],
  [/deferred.*tax.*liab|provision/i, "liability_non_current"],
  [/deferred.*grant|govt\.?\s*grant/i, "liability_non_current"],
];
const LIABILITY_DEFAULT = "liability_current";
function classifyAccountType(name, hasDebit, hasCredit) {
  for (const [pattern, type] of PRE_SIDE_RULES) if (pattern.test(name)) return type;
  const side = hasDebit && !hasCredit ? "asset" : hasCredit && !hasDebit ? "liability" : null;
  if (side === "asset") {
    for (const [pattern, type] of ASSET_SIDE_RULES) if (pattern.test(name)) return type;
    return ASSET_DEFAULT;
  }
  if (side === "liability") {
    for (const [pattern, type] of LIABILITY_SIDE_RULES) if (pattern.test(name)) return type;
    return LIABILITY_DEFAULT;
  }
  for (const [pattern, type] of [...ASSET_SIDE_RULES, ...LIABILITY_SIDE_RULES]) if (pattern.test(name)) return type;
  return ASSET_DEFAULT;
}
const ACCOUNT_CODE_BASE = {
  asset_receivable: 139000, asset_cash: 112000, asset_current: 130000, asset_non_current: 180000,
  asset_prepayments: 140000, asset_fixed: 150000, liability_payable: 400000, liability_credit_card: 410000,
  liability_current: 200000, liability_non_current: 250000, equity: 300000, equity_unaffected: 350000,
  income: 500000, income_other: 550000, expense: 600000, expense_depreciation: 650000, expense_direct_cost: 610000,
  off_balance: 900000,
};

// ---------------------------------------------------------------------------
// File parsing — unchanged, still 100% client-side (SheetJS)
// ---------------------------------------------------------------------------
export function parseWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const sheets = wb.SheetNames.map((name) => {
          const ws = wb.Sheets[name];
          const rawJson = XLSX.utils.sheet_to_json(ws, { defval: "" });
          if (!rawJson.length) return null;
          const headers = Object.keys(rawJson[0]);
          const json = rawJson.filter((row) => !isRowEmpty(row, headers));
          if (!json.length) return null;
          return { name, headers, rows: json, analysis: analyzeSheet(headers), skipDuplicates: true, duplicateAction: "smart", stockLocationId: null };
        }).filter(Boolean);
        resolve(sheets);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Could not read the file"));
    reader.readAsBinaryString(file);
  });
}

// ---------------------------------------------------------------------------
// Engine factory — one instance per active Odoo connection. All network
// calls go through /api/odoo-proxy (server-side), never directly to Odoo.
// ---------------------------------------------------------------------------
export function createEngine(connection) {
  let uid = null;
  const relationCache = new Map();
  const accountCache = new Map();
  const accountCodeCounters = {};
  const gstTaxCache = new Map();
  const attributeCache = new Map();
  const attributeValueCache = new Map();
  const locationResolveCache = new Map();
  const quantLocationCache = new Map();
  const defaultLocationParentCache = { checked: false, id: false };
  let internalPickingTypeCache = null;

  async function odooCall(service, method, args) {
    const res = await fetch("/api/odoo-proxy", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: connection.url, service, method, args }),
    });
    const data = await res.json();
    if (data.error) {
      const msg = data.error.data?.message || data.error.message || "Odoo RPC error";
      throw new Error(msg);
    }
    return data.result;
  }
  async function odooAuthenticate() {
    uid = await odooCall("common", "authenticate", [connection.db, connection.username, connection.apiKey, {}]);
    return uid;
  }
  async function odooExecute(model, method, args, kwargs = {}) {
    return odooCall("object", "execute_kw", [connection.db, uid, connection.apiKey, model, method, args, kwargs]);
  }
  async function withRetry(fn, retries = 1) {
    try { return await fn(); }
    catch (e) {
      const looksTransient = e instanceof TypeError || /network|fetch|timeout/i.test(e.message || "");
      if (retries > 0 && looksTransient) { await new Promise((r) => setTimeout(r, 600)); return withRetry(fn, retries - 1); }
      throw e;
    }
  }

  async function checkConnection() {
    if (!connection.url || !connection.db || !connection.username || !connection.apiKey) return false;
    try { await odooAuthenticate(); return !!uid; } catch (e) { return false; }
  }

  async function loadStockLocations() {
    try {
      return await odooExecute("stock.location", "search_read", [[["usage", "=", "internal"]]], { fields: ["id", "display_name"], order: "complete_name" });
    } catch (e) { return []; }
  }

  async function findDuplicate(model, dedupeKeys, values) {
    if (!dedupeKeys) return null;
    for (const key of dedupeKeys) {
      const usable = key.every((f) => values[f] !== undefined && values[f] !== "" && values[f] !== false && values[f] !== null);
      if (!usable) continue;
      try {
        const domain = key.map((f) => [f, "=", values[f]]);
        const ids = await odooExecute(model, "search", [domain], { limit: 1 });
        if (ids.length) return { id: ids[0], matchedOn: key };
      } catch (e) { continue; }
    }
    return null;
  }

  async function resolveRelation(field, value, allowCreate) {
    if (value === null || value === undefined || value === "") return { id: false, created: false, note: null };
    const cacheKey = `${field.relModel}::${String(value).toLowerCase()}`;
    if (!field.confirmCreate && relationCache.has(cacheKey)) return relationCache.get(cacheKey);
    const canCreate = field.confirmCreate ? allowCreate === true : field.autoCreate && connection.autoCreateSafe;
    let result;
    try {
      let ids = await odooExecute(field.relModel, "search", [[["name", "=", String(value)]]], { limit: 1 });
      if (!ids.length) ids = await odooExecute(field.relModel, "search", [[["name", "ilike", String(value)]]], { limit: 1 });
      if (ids.length) {
        result = { id: ids[0], created: false, note: null };
      } else if (canCreate) {
        const newId = await odooExecute(field.relModel, "create", [{ name: String(value), ...(field.autoCreateExtra || {}) }]);
        result = { id: newId, created: true, note: `Created new ${field.relModel} "${value}"` };
      } else if (field.confirmCreate) {
        result = { id: false, created: false, note: `"${value}" wasn't created — new-record creation for this import wasn't confirmed. Re-run and confirm, or create it in Odoo first.` };
      } else {
        result = { id: false, created: false, note: `"${value}" not found in ${field.relModel} — create it in Odoo first, or map a different column.` };
      }
    } catch (e) {
      result = { id: false, created: false, note: `Lookup failed for "${value}": ${e.message}` };
    }
    if (!field.confirmCreate) relationCache.set(cacheKey, result);
    return result;
  }

  async function nextFreeAccountCode(accountType) {
    let code = accountCodeCounters[accountType] ?? (ACCOUNT_CODE_BASE[accountType] || 190000);
    for (let attempts = 0; attempts < 500; attempts++) {
      const count = await odooExecute("account.account", "search_count", [[["code", "=", String(code)]]]);
      if (!count) { accountCodeCounters[accountType] = code + 10; return String(code); }
      code += 10;
    }
    return `AUTO${Math.floor(Math.random() * 1e6)}`;
  }

  async function resolveOrCreateAccount(name, hasDebit, hasCredit, notes) {
    if (!name) return { id: false, created: false, note: null };
    const cacheKey = String(name).toLowerCase();
    if (accountCache.has(cacheKey)) {
      const cached = accountCache.get(cacheKey);
      if (cached.note) notes.push(cached.note);
      return cached;
    }
    let result;
    try {
      let ids = await odooExecute("account.account", "search", [[["name", "=", String(name)]]], { limit: 1 });
      if (!ids.length) ids = await odooExecute("account.account", "search", [[["name", "ilike", String(name)]]], { limit: 1 });
      if (ids.length) {
        result = { id: ids[0], created: false, note: null };
      } else {
        const accountType = classifyAccountType(String(name), hasDebit, hasCredit);
        const code = await nextFreeAccountCode(accountType);
        const newId = await odooExecute("account.account", "create", [{ name: String(name), code, account_type: accountType }]);
        result = { id: newId, created: true, note: `Created new account "${name}" — code ${code}, type "${accountType}" (auto-detected from the name). Review it in Chart of Accounts.` };
      }
    } catch (e) {
      result = { id: false, created: false, note: `Account lookup/creation failed for "${name}": ${e.message}` };
    }
    accountCache.set(cacheKey, result);
    if (result.note) notes.push(result.note);
    return result;
  }

  async function resolveGstTaxIds(cellText, notes, kind) {
    const text = String(cellText || "").trim();
    if (!text) return null;
    const suffix = kind === "purchase" ? "P" : "S";
    const key = `${kind}::${text.toLowerCase()}`;
    if (gstTaxCache.has(key)) return gstTaxCache.get(key);
    let taxName = null;
    const igstMatch = text.match(/([\d.]+)\s*%\s*igst/i) || text.match(/igst[^0-9]*([\d.]+)\s*%/i);
    if (igstMatch) {
      taxName = `${igstMatch[1]}% IGST ${suffix}`;
    } else {
      const pcts = [...text.matchAll(/([\d.]+)\s*%/g)].map((m) => parseFloat(m[1])).filter((n) => !isNaN(n));
      if (pcts.length) {
        const total = pcts.reduce((a, b) => a + b, 0);
        taxName = `${Number.isInteger(total) ? total : total.toFixed(2)}% GST ${suffix}`;
      }
    }
    let result = null;
    if (taxName) {
      try {
        let ids = await odooExecute("account.tax", "search", [[["name", "=", taxName], ["type_tax_use", "=", kind]]], { limit: 1 });
        if (!ids.length) ids = await odooExecute("account.tax", "search", [[["name", "=ilike", taxName], ["type_tax_use", "=", kind]]], { limit: 1 });
        if (ids.length) {
          result = [[6, 0, [ids[0]]]];
        } else {
          notes.push(`Could not find a ${kind} tax named "${taxName}" (from "${text}") — that line was imported without a tax; set it manually or add a matching tax in Accounting > Configuration > Taxes.`);
        }
      } catch (e) {
        notes.push(`Tax lookup failed for "${text}": ${e.message}`);
      }
    }
    gstTaxCache.set(key, result);
    return result;
  }

  async function resolveAttributeValue(attributeId, valueName) {
    const key = `${attributeId}::${valueName.toLowerCase()}`;
    if (attributeValueCache.has(key)) return attributeValueCache.get(key);
    let ids = await odooExecute("product.attribute.value", "search", [[["attribute_id", "=", attributeId], ["name", "=", valueName]]], { limit: 1 });
    if (!ids.length) ids = await odooExecute("product.attribute.value", "search", [[["attribute_id", "=", attributeId], ["name", "ilike", valueName]]], { limit: 1 });
    let id;
    if (ids.length) id = ids[0];
    else id = await odooExecute("product.attribute.value", "create", [{ name: valueName, attribute_id: attributeId }]);
    attributeValueCache.set(key, id);
    return id;
  }

  async function parseAttributeString(raw, notes) {
    const groups = String(raw).split(";").map((g) => g.trim()).filter(Boolean);
    const lines = [];
    for (const group of groups) {
      const [attrPart, valuesPart] = group.split(":");
      const attrName = (attrPart || "").trim();
      if (!attrName || !valuesPart) { notes.push(`Could not parse attribute group "${group}" — expected format "Attribute:Value1,Value2".`); continue; }
      const valueNames = valuesPart.split(",").map((v) => v.trim()).filter(Boolean);
      if (!valueNames.length) continue;
      try {
        let attrId = attributeCache.get(attrName.toLowerCase());
        if (attrId === undefined) {
          let ids = await odooExecute("product.attribute", "search", [[["name", "=", attrName]]], { limit: 1 });
          if (!ids.length) ids = await odooExecute("product.attribute", "search", [[["name", "ilike", attrName]]], { limit: 1 });
          if (ids.length) attrId = ids[0];
          else { attrId = await odooExecute("product.attribute", "create", [{ name: attrName }]); notes.push(`Created new attribute "${attrName}".`); }
          attributeCache.set(attrName.toLowerCase(), attrId);
        }
        const valueIds = [];
        for (const vName of valueNames) valueIds.push(await resolveAttributeValue(attrId, vName));
        if (valueIds.length) lines.push([0, 0, { attribute_id: attrId, value_ids: [[6, 0, valueIds]] }]);
      } catch (e) {
        notes.push(`Failed to resolve attribute "${attrName}": ${e.message}`);
      }
    }
    return lines;
  }

  async function applyOnHandQuantity(templateId, qty, locationId, notes) {
    try {
      const productIds = await odooExecute("product.product", "search", [[["product_tmpl_id", "=", templateId]]], { limit: 1 });
      if (!productIds.length) { notes.push(`Could not set on-hand quantity for product #${templateId} — no variant found.`); return; }
      const productId = productIds[0];
      const quantIds = await odooExecute("stock.quant", "search", [[["product_id", "=", productId], ["location_id", "=", locationId]]], { limit: 1 });
      let quantId;
      if (quantIds.length) { quantId = quantIds[0]; await odooExecute("stock.quant", "write", [[quantId], { inventory_quantity: qty }]); }
      else quantId = await odooExecute("stock.quant", "create", [{ product_id: productId, location_id: locationId, inventory_quantity: qty }]);
      await odooExecute("stock.quant", "action_apply_inventory", [[quantId]]);
    } catch (e) {
      notes.push(`On-hand quantity update failed: ${e.message}`);
    }
  }

  async function getDefaultLocationParent() {
    if (defaultLocationParentCache.checked) return defaultLocationParentCache.id;
    defaultLocationParentCache.checked = true;
    try {
      const whs = await odooExecute("stock.warehouse", "search_read", [[]], { fields: ["lot_stock_id"], limit: 1 });
      defaultLocationParentCache.id = whs.length && whs[0].lot_stock_id ? whs[0].lot_stock_id[0] : false;
    } catch (e) { defaultLocationParentCache.id = false; }
    return defaultLocationParentCache.id;
  }

  async function resolveOrCreateLocation(name, notes) {
    if (!name) return { id: false, note: null };
    const key = String(name).trim().toLowerCase();
    if (locationResolveCache.has(key)) {
      const cached = locationResolveCache.get(key);
      if (cached.note) notes.push(cached.note);
      return cached;
    }
    let result;
    try {
      let ids = await odooExecute("stock.location", "search", [[["complete_name", "=", String(name)], ["usage", "=", "internal"]]], { limit: 1 });
      if (!ids.length) ids = await odooExecute("stock.location", "search", [[["name", "=", String(name)], ["usage", "=", "internal"]]], { limit: 1 });
      if (!ids.length) ids = await odooExecute("stock.location", "search", [[["complete_name", "ilike", String(name)], ["usage", "=", "internal"]]], { limit: 1 });
      if (ids.length) {
        result = { id: ids[0], note: null };
      } else {
        const parentId = await getDefaultLocationParent();
        const createVals = { name: String(name), usage: "internal" };
        if (parentId) createVals.location_id = parentId;
        const newId = await odooExecute("stock.location", "create", [createVals]);
        result = { id: newId, note: `Created new location "${name}"${parentId ? "" : " at the root of your warehouse locations — no default warehouse stock location was found to nest it under, you may want to move it in Odoo"}.` };
      }
    } catch (e) {
      result = { id: false, note: `Location lookup/creation failed for "${name}": ${e.message}` };
    }
    locationResolveCache.set(key, result);
    if (result.note) notes.push(result.note);
    return result;
  }

  async function getInternalPickingType() {
    if (internalPickingTypeCache !== null) return internalPickingTypeCache;
    try {
      const types = await odooExecute("stock.picking.type", "search_read", [[["code", "=", "internal"]]], { fields: ["id"], limit: 1 });
      internalPickingTypeCache = types.length ? types[0].id : false;
    } catch (e) { internalPickingTypeCache = false; }
    return internalPickingTypeCache;
  }

  async function currentQuantAt(productId, locationId) {
    const quants = await odooExecute("stock.quant", "search_read", [[["product_id", "=", productId], ["location_id", "=", locationId]]], { fields: ["quantity"], limit: 1 });
    return quants.length ? quants[0].quantity : 0;
  }

  async function applyRelocation(templateId, fromName, toName, qty, notes) {
    if (!fromName || !toName) return;
    try {
      const productIds = await odooExecute("product.product", "search", [[["product_tmpl_id", "=", templateId]]], { limit: 1 });
      if (!productIds.length) { notes.push(`Could not relocate stock for product #${templateId} — no variant found.`); return; }
      const productId = productIds[0];
      const from = await resolveOrCreateLocation(fromName, notes);
      const to = await resolveOrCreateLocation(toName, notes);
      if (from.id === false || to.id === false) { notes.push(`Relocation skipped for product #${templateId} — could not resolve the From/To location.`); return; }
      if (from.id === to.id) { notes.push(`Relocation skipped for product #${templateId} — From and To location are the same ("${fromName}").`); return; }
      const pickingTypeId = await getInternalPickingType();
      if (!pickingTypeId) { notes.push(`Relocation skipped for product #${templateId} — no Internal Transfers operation type found in this database.`); return; }
      const fromQty = await currentQuantAt(productId, from.id);
      let moveQty = qty;
      if (moveQty === null || moveQty === undefined || moveQty === "") moveQty = fromQty;
      if (!moveQty || moveQty <= 0) { notes.push(`Relocation skipped for product #${templateId} — nothing on hand at "${fromName}" to move.`); return; }
      if (fromQty < moveQty) notes.push(`Relocating ${moveQty} of product #${templateId} from "${fromName}", but only ${fromQty} was on hand there — that location's on-hand quantity will go negative.`);
      const pickingId = await odooExecute("stock.picking", "create", [{
        picking_type_id: pickingTypeId, location_id: from.id, location_dest_id: to.id,
        move_ids: [[0, 0, { product_id: productId, product_uom_qty: moveQty, location_id: from.id, location_dest_id: to.id }]],
      }]);
      await odooExecute("stock.picking", "action_confirm", [[pickingId]]);
      try { await odooExecute("stock.picking", "action_assign", [[pickingId]]); } catch (e) {}
      const moveIds = await odooExecute("stock.move", "search", [[["picking_id", "=", pickingId]]]);
      for (const moveId of moveIds) {
        try { await odooExecute("stock.move", "write", [[moveId], { quantity: moveQty, picked: true }]); }
        catch (e) { await odooExecute("stock.move", "write", [[moveId], { quantity_done: moveQty }]); }
      }
      await odooExecute("stock.picking", "button_validate", [[pickingId]]);
      notes.push(`Relocated ${moveQty} of product #${templateId} from "${fromName}" to "${toName}" (Internal Transfer).`);
    } catch (e) {
      notes.push(`Relocation failed for product #${templateId}: ${e.message}`);
    }
  }

  // Physical Inventory module — resolve a Location by name WITHOUT
  // auto-creating it (unlike Relocate's resolveOrCreateLocation above): a
  // stock count should always point at a real, already-decided location, so a
  // typo or missing location is reported and that row is skipped rather than
  // silently creating a new location. Matches on the full path
  // ("WH/Pre-Production/ALPHA - I") first, then falls back to just the last
  // segment name, then a fuzzy ilike on the path.
  async function resolveLocationForQuant(name, notes) {
    if (!name) return { id: false, note: null };
    const key = String(name).trim().toLowerCase();
    if (quantLocationCache.has(key)) {
      const cached = quantLocationCache.get(key);
      if (cached.note) notes.push(cached.note);
      return cached;
    }
    let result;
    try {
      let ids = await odooExecute("stock.location", "search", [[["complete_name", "=", String(name)], ["usage", "=", "internal"]]], { limit: 1 });
      if (!ids.length) ids = await odooExecute("stock.location", "search", [[["name", "=", String(name)], ["usage", "=", "internal"]]], { limit: 1 });
      if (!ids.length) ids = await odooExecute("stock.location", "search", [[["complete_name", "ilike", String(name)], ["usage", "=", "internal"]]], { limit: 1 });
      if (ids.length) {
        result = { id: ids[0], note: null };
      } else {
        result = { id: false, note: `Location "${name}" not found — check the spelling/path, or create it in Odoo first (Physical Inventory doesn't auto-create locations).` };
      }
    } catch (e) {
      result = { id: false, note: `Location lookup failed for "${name}": ${e.message}` };
    }
    quantLocationCache.set(key, result);
    if (result.note) notes.push(result.note);
    return result;
  }

  // Physical Inventory module — set the true on-hand quantity for one
  // Product+Location to exactly `targetQty`. This is the corrective technique
  // worked out by hand after a real import doubled every row's quantity: Odoo
  // can already have (or create) more than one stock.quant row for the same
  // product+location, and just writing inventory_quantity onto ONE of them
  // adds on top of the others instead of replacing the total. So this always
  // reads every quant row for that product+location first, and sets the
  // primary row's inventory_quantity to (target - sum of every OTHER row),
  // forcing the true combined total to land exactly on target — never double,
  // no matter how many quant rows already exist underneath.
  async function applyPhysicalInventoryCount(productId, locationId, targetQty, notes) {
    try {
      const quants = await odooExecute("stock.quant", "search_read", [[["product_id", "=", productId], ["location_id", "=", locationId]]], { fields: ["id", "quantity"] });
      if (!quants.length) {
        const newId = await odooExecute("stock.quant", "create", [{ product_id: productId, location_id: locationId, inventory_quantity: targetQty }]);
        await odooExecute("stock.quant", "action_apply_inventory", [[newId]]);
        return;
      }
      const primary = quants[0];
      const othersTotal = quants.slice(1).reduce((s, q) => s + q.quantity, 0);
      await odooExecute("stock.quant", "write", [[primary.id], { inventory_quantity: targetQty - othersTotal }]);
      await odooExecute("stock.quant", "action_apply_inventory", [[primary.id]]);
    } catch (e) {
      notes.push(`Physical inventory count failed for product #${productId} at location #${locationId}: ${e.message}`);
    }
  }

  async function findProductTransactions(templateId) {
    const productIds = await odooExecute("product.product", "search", [[["product_tmpl_id", "=", templateId]]]);
    if (!productIds.length) return [];
    const checks = [
      { model: "stock.move", domain: [["product_id", "in", productIds], ["state", "!=", "cancel"], ["location_id.usage", "!=", "inventory"], ["location_dest_id.usage", "!=", "inventory"]], label: "stock moves" },
      { model: "sale.order.line", domain: [["product_id", "in", productIds]], label: "sales orders" },
      { model: "purchase.order.line", domain: [["product_id", "in", productIds]], label: "purchase orders" },
      { model: "account.move.line", domain: [["product_id", "in", productIds], ["parent_state", "!=", "cancel"]], label: "invoices/bills" },
    ];
    const reasons = [];
    for (const check of checks) {
      try { const count = await odooExecute(check.model, "search_count", [check.domain]); if (count > 0) reasons.push(check.label); }
      catch (e) { continue; }
    }
    return reasons;
  }

  // -------------------------------------------------------------------------
  // Upload — same grouping / resolution / dedupe logic as the extension.
  // Instead of a module-level uploadState + render(), progress is reported
  // via onProgress(partialState), and the function returns the final state.
  // Two-phase flow for confirmCreate fields (Sales/PO line products):
  //   1st call -> may return {status:'confirm-create', missingProducts}
  //   caller sets sheet.autoCreateProducts = true and calls again to proceed.
  // -------------------------------------------------------------------------
  async function uploadSheetToOdoo(sheet, onProgress) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const schema = ODOO_SCHEMAS[sheet.analysis.moduleKey];
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const skipDuplicates = sheet.skipDuplicates !== false;
    const duplicateAction = sheet.duplicateAction || "smart";

    let usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));
    const requiredLineFields = mappedFields.filter((m) => m.field.isLineRequired);
    if (requiredLineFields.length) usableRows = usableRows.filter((row) => requiredLineFields.every((m) => row[m.header] !== "" && row[m.header] != null));

    const debitHeader = mappedFields.find((m) => m.field.isDebitField)?.header;
    const creditHeader = mappedFields.find((m) => m.field.isCreditField)?.header;

    const confirmFields = mappedFields.filter((m) => m.field.confirmCreate);
    if (confirmFields.length && sheet.autoCreateProducts !== true) {
      emit({ status: "checking" });
      const missing = [];
      for (const m of confirmFields) {
        const names = [...new Set(usableRows.map((row) => row[m.header]).filter((v) => v !== "" && v != null).map((v) => cleanText(String(v))))];
        if (!names.length) continue;
        try {
          const foundExact = await odooExecute(m.field.relModel, "search_read", [[["name", "in", names]]], { fields: ["name"] });
          const foundNames = new Set(foundExact.map((r) => String(r.name).toLowerCase()));
          const stillMissing = names.filter((n) => !foundNames.has(n.toLowerCase()));
          for (const name of stillMissing) {
            const ids = await odooExecute(m.field.relModel, "search", [[["name", "ilike", name]]], { limit: 1 });
            if (!ids.length) missing.push(name);
          }
        } catch (e) { missing.push(...names); }
      }
      if (missing.length) return emit({ status: "confirm-create", missingProducts: [...new Set(missing)].sort() });
    }

    const groupKeyField = mappedFields.find((m) => m.field.isGroupKey);
    const groups = [];
    if (groupKeyField) {
      let lastKey;
      for (const row of usableRows) {
        const key = row[groupKeyField.header];
        if (!groups.length || key !== lastKey) groups.push([row]); else groups[groups.length - 1].push(row);
        lastKey = key;
      }
    } else {
      const headerFieldDefs = mappedFields.filter((m) => !m.field.isLine && !m.field.isVendorLine);
      const rowIsContinuation = (row) => headerFieldDefs.length > 0 && headerFieldDefs.every((m) => row[m.header] === "" || row[m.header] == null);
      for (const row of usableRows) {
        if (!groups.length || !rowIsContinuation(row)) groups.push([row]); else groups[groups.length - 1].push(row);
      }
    }

    emit({ status: "uploading", progress: 0, total: groups.length });
    const createdIds = []; const errors = []; const notes = []; let duplicateCount = 0;

    for (let g = 0; g < groups.length; g++) {
      const group = groups[g];
      const values = { ...(schema.defaults || {}) };
      const linesByField = {};
      let pendingQty = null;
      let pendingRelocate = { from: null, to: null, qty: null };
      let pendingCountedQty = null; // Physical Inventory module: Counted Quantity for this row
      try {
        for (const row of group) {
          const lineAccumulator = {};
          for (const m of mappedFields) {
            const raw = row[m.header];
            if (raw === "" || raw == null) continue;
            const field = m.field;
            if (field.isGroupKey) { values.ref = cleanText(String(raw)); continue; }
            if (field.isIgnored) continue;
            if (field.isQuantity) { pendingQty = cleanNumber(raw); continue; }
            if (field.isRelocateFrom) { pendingRelocate.from = cleanText(String(raw)); continue; }
            if (field.isRelocateTo) { pendingRelocate.to = cleanText(String(raw)); continue; }
            if (field.isRelocateQty) { pendingRelocate.qty = cleanNumber(raw); continue; }
            // Physical Inventory module — Counted Quantity isn't sent to
            // create()/write(); it's applied afterward via
            // applyPhysicalInventoryCount so duplicate quant rows can't
            // silently double the total (see that function's comment).
            if (field.isPhysicalInventoryQty) { pendingCountedQty = cleanNumber(raw); continue; }
            // Physical Inventory module — Location resolves against the full
            // location path WITHOUT auto-creating (see resolveLocationForQuant).
            if (field.isLocationRelation) {
              const resolved = await resolveLocationForQuant(cleanText(String(raw)), notes);
              if (resolved.id === false) continue; // skip field, don't fail the row for it
              values[field.name] = resolved.id;
              continue;
            }
            if (field.isAccountField) {
              const hasDebit = debitHeader ? row[debitHeader] !== "" && row[debitHeader] != null : false;
              const hasCredit = creditHeader ? row[creditHeader] !== "" && row[creditHeader] != null : false;
              const resolved = await resolveOrCreateAccount(cleanText(String(raw)), hasDebit, hasCredit, notes);
              if (resolved.id === false) continue;
              const [lineField, subField] = field.name.split("/");
              lineAccumulator[lineField] = lineAccumulator[lineField] || {};
              lineAccumulator[lineField][subField] = resolved.id;
              continue;
            }
            if (field.isAttributeLine) {
              const lines = await parseAttributeString(raw, notes);
              if (lines.length) values[field.name] = (values[field.name] || []).concat(lines);
              continue;
            }
            if (field.isGstTaxField) {
              const taxCmd = await resolveGstTaxIds(raw, notes, field.isGstTaxField);
              if (taxCmd) {
                const [lineField, subField] = field.name.split("/");
                lineAccumulator[lineField] = lineAccumulator[lineField] || {};
                lineAccumulator[lineField][subField] = taxCmd;
              }
              continue;
            }
            if (field.isTagField) {
              const tagNames = String(raw).split(",").map((t) => t.trim()).filter(Boolean);
              const tagIds = [];
              for (const tName of tagNames) {
                const resolved = await resolveRelation(field, tName);
                if (resolved.note) notes.push(resolved.note);
                if (resolved.id !== false) tagIds.push(resolved.id);
              }
              if (tagIds.length) values[field.name] = [[6, 0, tagIds]];
              continue;
            }
            let val = raw;
            if (field.number) val = cleanNumber(val);
            else if (field.boolean) val = cleanBoolean(val);
            else if (field.date) val = cleanDate(val);
            else if (field.selection === "productType") val = cleanSelection(val, PRODUCT_TYPE_MAP);
            else if (field.selection === "invoicePolicy") val = cleanSelection(val, INVOICE_POLICY_MAP);
            else if (field.selection === "purchaseMethod") val = cleanSelection(val, PURCHASE_METHOD_MAP);
            else if (field.selection === "purchaseType") val = cleanSelection(val, PURCHASE_TYPE_MAP);
            else if (typeof val === "string") val = cleanText(val);
            if (field.name === "email" && val && !isValidEmail(val)) { notes.push(`Skipped invalid email "${val}" — row still created without it.`); continue; }
            if (field.relation) {
              const resolved = await resolveRelation(field, val, field.confirmCreate ? sheet.autoCreateProducts === true : undefined);
              if (resolved.note) notes.push(resolved.note);
              if (resolved.id === false) continue;
              val = resolved.id;
            }
            if (field.isLine || field.isVendorLine) {
              const [lineField, subField] = field.name.split("/");
              lineAccumulator[lineField] = lineAccumulator[lineField] || {};
              lineAccumulator[lineField][subField] = val;
            } else {
              values[field.name] = val;
            }
          }
          for (const [lineField, lineVals] of Object.entries(lineAccumulator)) {
            if (Object.keys(lineVals).length) { linesByField[lineField] = linesByField[lineField] || []; linesByField[lineField].push(lineVals); }
          }
        }
        for (const [lineField, lines] of Object.entries(linesByField)) values[lineField] = lines.map((l) => [0, 0, l]);

        // Physical Inventory module — never goes through the generic
        // create()/dedupe/write() path below (a plain create() on stock.quant
        // is exactly the mechanism that produced the doubling bug this module
        // exists to avoid). Instead, once Location + Product + Counted
        // Quantity are all resolved for this row, hand off to
        // applyPhysicalInventoryCount, then move to the next row.
        if (schema.model === "stock.quant") {
          if (values.location_id && values.product_id && pendingCountedQty !== null) {
            await applyPhysicalInventoryCount(values.product_id, values.location_id, pendingCountedQty, notes);
            createdIds.push(true); // count this row as a success for the summary
          } else {
            errors.push({ row: g, error: "Missing Location, Product, or Counted Quantity for this row.", data: values });
          }
          emit({ status: "uploading", progress: g + 1, total: groups.length });
          continue;
        }

        if (skipDuplicates) {
          const dup = await findDuplicate(schema.model, schema.dedupeKeys, values);
          if (dup) {
            duplicateCount++;
            let effectiveAction = duplicateAction; let blockedReasons = [];
            if (duplicateAction === "smart" && schema.model === "product.template") {
              blockedReasons = await findProductTransactions(dup.id);
              effectiveAction = blockedReasons.length ? "skip" : "update";
            } else if (duplicateAction === "smart") {
              effectiveAction = "update";
            }
            if (effectiveAction === "update") {
              try {
                const writeValues = { ...values };
                for (const lineField of Object.keys(linesByField)) writeValues[lineField] = [[5, 0, 0], ...values[lineField]];
                await withRetry(() => odooExecute(schema.model, "write", [[dup.id], writeValues]));
                notes.push(`Row ${g + 1}: updated existing record #${dup.id} (matched on ${dup.matchedOn.join(" + ")}).`);
              } catch (e) {
                errors.push({ row: g, error: `Update failed for existing record #${dup.id}: ${e.message}`, data: values });
              }
            } else if (blockedReasons.length) {
              notes.push(`Row ${g + 1}: skipped — record #${dup.id} already has ${blockedReasons.join(", ")}, so it wasn't overwritten. Update it manually in Odoo if you're sure.`);
            } else {
              notes.push(`Row ${g + 1}: skipped — already exists as record #${dup.id} (matched on ${dup.matchedOn.join(" + ")}).`);
            }
            if (pendingQty !== null && sheet.stockLocationId && schema.model === "product.template") await applyOnHandQuantity(dup.id, pendingQty, sheet.stockLocationId, notes);
            if (schema.model === "product.template") await applyRelocation(dup.id, pendingRelocate.from, pendingRelocate.to, pendingRelocate.qty, notes);
            emit({ status: "uploading", progress: g + 1, total: groups.length });
            continue;
          }
        }

        const newId = await withRetry(() => odooExecute(schema.model, "create", [values]));
        createdIds.push(newId);
        if (pendingQty !== null && sheet.stockLocationId && schema.model === "product.template") await applyOnHandQuantity(newId, pendingQty, sheet.stockLocationId, notes);
        if (schema.model === "product.template") await applyRelocation(newId, pendingRelocate.from, pendingRelocate.to, pendingRelocate.qty, notes);
      } catch (e) {
        errors.push({ row: g, error: e.message, data: values });
      }
      emit({ status: "uploading", progress: g + 1, total: groups.length });
    }

    return emit({
      status: "done",
      result: {
        total: groups.length, created_count: createdIds.length, duplicate_count: duplicateCount,
        duplicate_action: duplicateAction, failed_count: errors.length, errors, notes: [...new Set(notes)],
      },
    });
  }

  return {
    odooAuthenticate,
    odooExecute,
    checkConnection,
    loadStockLocations,
    uploadSheetToOdoo,
    get uid() { return uid; },
  };
}
