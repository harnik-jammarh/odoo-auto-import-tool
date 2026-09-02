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
    note: "Import screen: Inventory/Sales app -> Products -> Action -> Import records. Opening stock qty is set separately via an Inventory Adjustment, not on the product import. This module now also covers the product's Attributes & Variants, Sales, and Purchase tabs — map the extra columns below like any other field. Images: put a direct image URL (https://...) in the Main Image column and it's downloaded and set as the product's photo; for eCommerce Extra Media, put multiple URLs in one cell separated by commas/semicolons and each one is added as a separate image in the product's Extra Media gallery. Google Drive share links (drive.google.com/file/d/.../view) work too — the file just needs to be shared as \"Anyone with the link\", since this app can only read what's public. A Drive FOLDER link (drive.google.com/drive/folders/...) works only in the Extra Images column and only once you've added a Google Drive API Key on this database's connection (edit the database, see the Google Drive API Key field) — it then pulls in every image inside that folder automatically; without a key, share each image's own file link instead and list them comma-separated. Pictures actually pasted/inserted into the sheet's cells (not typed as links) also work — they're picked up automatically as \"Embedded Main Image\"/\"Embedded Extra Images\" and need no URL at all, as long as one picture sits in the same row as its product. Local file paths don't work here — the images need to already be reachable at a URL. Re-uploading the same file adds the Extra Media images again rather than replacing them, so avoid re-running a sheet that already succeeded for that column.",
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
      // --- Images ---
      // Main product photo (General Information tab) — one URL, downloaded
      // and set as image_1920 directly on the product.template create/write
      // values (see the isProductImage handling in uploadSheetToOdoo).
      { name: "image_1920", label: "Main Image (URL)", aliases: ["image", "image url", "product image", "main image", "photo", "picture"], isProductImage: true },
      // eCommerce "Extra Media" gallery — a product sold online can have
      // several extra photos beyond the main one (product.image records
      // linked to the template). One cell can list multiple URLs separated
      // by a comma, semicolon, or pipe; each becomes its own Extra Media
      // image, applied after the product exists (see applyExtraImages).
      { name: "extra_images", label: "eCommerce Extra Images (URLs, comma-separated)", aliases: ["extra images", "extra media", "gallery images", "ecommerce images", "ecommerce media", "additional images", "media images", "other images"], isExtraImages: true },
    ],
  },
  accounting: {
    label: "Journal Entries (Invoices/Bills)", model: "account.move", icon: "🧮",
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
    label: "Opening Balances (Balance Sheet / Debtors & Creditors)", model: "account.move", icon: "💰",
    note: "For a raw debit/credit journal entry sheet like the Debtors/Creditors/Balance Sheet opening balance templates — one row per journal line, with every line for the same entry sharing the same Journal Entry Reference value. If any Account name in the sheet isn't found in your Chart of Accounts, nothing is created until you say so: before the upload runs, you're shown every such account with a type picker (Receivable, Payable, Current Assets, Current Liabilities, Fixed Assets, Bank and Cash, Equity, Income, Expenses, etc.), pre-selected with a guess based on which side (Debit/Credit) the row's amount is on and keywords in the name — review or change each one, then continue. Not for invoices/bills (use Journal Entries (Invoices/Bills) for those).",
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
  vendorBills: {
    label: "Vendor Bills", model: "account.move", icon: "🧾",
    isVendorBillsModule: true,
    note: "Import screen: Accounting/Purchase app -> Vendor Bills -> Action -> Import records (or just use this module directly). If this bill has a Purchase Order, map a PO Reference column matching exactly what's in that PO's own Vendor Reference field (the same field the Purchase Orders module writes your file's PO number into) — the bill's Vendor and Purchase Order link are both pulled from that matched PO automatically, you never need a separate Vendor column, and Line / Unit Price and Line / Taxes become optional (pulled from the PO's own rate/tax for that product). If more than one PO shares the same Vendor Reference, that row is stopped rather than guessing which one — map the PO Name column too (the PO's own Odoo number, e.g. P00088) to point at the exact one. If this bill has NO Purchase Order at all — a phone bill, courier charge, or any other cash/direct expense — just leave the PO Reference column out entirely (or don't map one): the bill still imports fine as a standalone bill, matched only by the Vendor column you do map. (If you map a PO Reference column and it fails to match, that's treated as a real error rather than silently skipped, since you clearly intended a link — but leaving PO Reference unmapped is a completely normal, supported way to import a bill with no PO.) Line / Product or Expense Account accepts either: a real product, or a Chart of Accounts entry like \"Telephone & Telex\" for a line that isn't a billed product at all — tried as a product first, then as an account, auto-creating a new expense account only if neither matches. Note is kept on the bill as an internal note — useful for recording how/when a cash bill was settled. The bill's own Line / Bill Quantity is what gets billed, never copied from the PO's own line quantities. Every bill is created as a draft in Odoo — nothing is posted automatically, so review and post afterward. Multi-line bills: rows below the first with a blank Vendor are treated as extra lines of the same bill. Optional, on top of all that: map TDS Section (and, if you don't want the bill's own Untaxed Amount used as the base, TDS Base Amount / TDS Rate %) on a bill's row to have this bill deducted for TDS right after it's created — same calculation and account handling as the standalone TDS Entry module, applied automatically, bill posted at the end. Separately, map Extra Line Item — Product or Expense Account (with Extra Line Item / Quantity, Amount, Taxes, Description as needed) to add one extra charge line (freight, loading, packing, etc.) that isn't itself a PO line — this one is added but the bill is left in Draft either way so you can review it, same as the standalone Bill Additional Lines module. Both are entirely optional per row; a bill with neither column filled in imports exactly as it always did.",
    defaults: { move_type: "in_invoice" },
    // NOTE: intentionally does NOT include ["invoice_origin", "partner_id"].
    // It's completely normal for several separate bills to point at the same
    // PO (partial deliveries, multiple shipments, etc.) — matching on the PO
    // link alone would treat the 2nd and 3rd bill against one PO as
    // "duplicates" of the 1st and silently overwrite it instead of creating
    // them. Bill Reference (your vendor's own invoice number) is the only
    // signal that's actually meant to be unique per bill.
    dedupeKeys: [["partner_id", "ref"]],
    fields: [
      { name: "__row_id", label: "Row ID (not imported — your own tracking key, if the sheet has one)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "partner_id", label: "Vendor (only needed for a bill with no PO — a PO-linked bill pulls it from the PO automatically)", aliases: ["vendor", "vendor name", "supplier", "supplier name"], relation: true, relModel: "res.partner", autoCreate: true },
      { name: "__po_ref", label: "PO Reference (must match the Purchase Order's Vendor Reference — leave the column out entirely for a bill with no PO)", aliases: ["po reference", "purchase order reference", "po ref", "order reference", "vendor reference", "po number", "purchase order"], isPoReferenceField: true },
      { name: "__po_name", label: "PO Name (Odoo's own reference, e.g. P00088 — only needed if two POs share the same Vendor Reference)", aliases: ["po name", "po number (odoo)", "odoo po reference", "po odoo number", "purchase order number"], isPoNameField: true },
      { name: "invoice_date", label: "Bill Date", aliases: ["bill date", "invoice date", "date"], date: true },
      { name: "invoice_date_due", label: "Due Date", aliases: ["due date", "payment due date"], date: true },
      { name: "ref", label: "Bill Reference (your vendor's own invoice/bill number)", aliases: ["bill reference", "bill no", "vendor bill no", "vendor invoice no", "invoice number"] },
      { name: "invoice_line_ids/product_id", label: "Line / Product or Expense Account", aliases: ["product", "item", "product name", "product or expense account", "account", "expense account"], isLine: true, isProductOrAccountField: true },
      { name: "invoice_line_ids/quantity", label: "Line / Bill Quantity", aliases: ["qty", "quantity", "bill qty", "billed quantity", "units"], number: true, isLine: true },
      { name: "invoice_line_ids/price_unit", label: "Line / Unit Price (optional — leave blank to use the PO's own rate for that product, if this bill has a PO)", aliases: ["unit price", "rate", "price", "purchase price"], number: true, isLine: true },
      { name: "invoice_line_ids/tax_ids", label: "Line / Taxes (optional — leave blank to use the PO's own tax for that product, if this bill has a PO)", aliases: ["taxes", "tax", "gst", "gst rate"], isLine: true, isGstTaxField: "purchase" },
      { name: "narration", label: "Note (kept on the bill as an internal note — e.g. how/when a cash bill was settled)", aliases: ["note", "notes", "reason", "remarks"] },
      { name: "__tds_section", label: "TDS Section (optional — leave blank for a bill with no TDS)", aliases: ["tds section", "section", "section code", "tds", "u/s"], isTdsSectionField: true },
      { name: "__tds_base_amount", label: "TDS Base Amount (optional — leave blank to use this bill's own Untaxed Amount)", aliases: ["tds base amount", "base amount", "taxable amount", "tds base"], number: true, isTdsBaseField: true },
      { name: "__tds_rate", label: "TDS Rate % (optional override — leave blank to use the standard rate for the Section)", aliases: ["tds rate", "tds %", "tds percentage"], number: true, isTdsRateField: true },
      { name: "__extra_line_item", label: "Extra Line Item — Product or Expense Account (optional, e.g. \"Freight & Cartage\")", aliases: ["extra line item", "extra charge", "additional line", "additional charge", "freight"], isExtraLineField: "product" },
      { name: "__extra_line_qty", label: "Extra Line Item / Quantity (optional — defaults to 1)", aliases: ["extra line quantity", "extra charge quantity", "additional quantity"], number: true, isExtraLineField: "quantity" },
      { name: "__extra_line_price", label: "Extra Line Item / Amount", aliases: ["extra line amount", "extra charge amount", "additional amount", "freight amount", "freight charges"], number: true, isExtraLineField: "price" },
      { name: "__extra_line_tax", label: "Extra Line Item / Taxes (optional)", aliases: ["extra line taxes", "extra charge tax", "additional tax"], isExtraLineField: "tax" },
      { name: "__extra_line_desc", label: "Extra Line Item / Description (optional — defaults to the product/account name)", aliases: ["extra line description", "extra charge description", "additional description"], isExtraLineField: "desc" },
    ],
  },
  salesInvoices: {
    label: "Sales Invoices", model: "account.move", icon: "🧾",
    isSalesInvoicesModule: true,
    note: "Import screen: Accounting/Sales app -> Invoices -> Action -> Import records (or just use this module directly). Each invoice is matched to its Sales Order by the SO Reference column — this must match exactly what's in that order's own Order Reference field (the same field the Sales module writes your file's order number into), and the invoice's Customer and Sales Order link are both pulled from that matched order automatically — you never need a separate Customer column here. If you don't map an SO Reference column at all (or leave it blank on a row), the order is auto-computed from the Invoice Reference instead — this only works when your own invoice number is the same value already sitting in that order's Order Reference field; otherwise map a real SO Reference column. If more than one order happens to share the same Order Reference, that row is stopped rather than guessing which one — map the SO Name column too (the order's own Odoo number, e.g. S00088, which is always unique) to point at the exact one. The invoice's own Line / Invoice Quantity is what gets invoiced, and it can be less than, more than, or different from the order's own quantity on purpose (a partial shipment, an extra charge, etc.) — this module always creates the invoice exactly with the quantity your sheet says, it never copies the order's own line quantities. Line / Unit Price and Line / Taxes are optional: since an invoice is really just converting the order, leave them blank on a line and the rate and tax are pulled straight from that same product's line on the matched order — only fill them in if this particular invoice genuinely charges a different rate or tax than the order did. If an SO Reference doesn't match any real order, that row is skipped with a clear reason rather than creating a customer-less invoice. Every invoice is created as a draft in Odoo, same as importing through Odoo's own screen — nothing is posted/validated automatically, so review and post them in Odoo afterward. Multi-line invoices: rows below the first with a blank SO Reference are treated as extra lines of the same invoice. Optional, on top of all that: map TDS Section (and, if you don't want the invoice's own Untaxed Amount used as the base, TDS Base Amount / TDS Rate %) on an invoice's row to have TDS the customer deducted recorded against it right after it's created, via Odoo's own native \"TDS Entry\" wizard — the invoice is posted automatically once TDS is applied. Also optional: map TCS Section (and TCS Rate % to override) to add the matching Tax Collected at Source tax onto every line of the invoice, same mechanism as the Line / Taxes GST column (Odoo's own l10n_in_tax_type = \"tcs\" taxes) rather than a wizard — this can be used with or without TDS. And optional: Extra Line Item / Quantity / Amount / Taxes / Description columns, the sales-side mirror of the Vendor Bills module's own Extra Line Item columns, for a charge billed alongside the SO'd goods/services (freight recovered, packing, etc.) that isn't itself an SO line. All three are entirely optional; an invoice with none of these columns filled in imports exactly as it always did.",
    defaults: { move_type: "out_invoice" },
    // NOTE: intentionally does NOT include ["invoice_origin", "partner_id"].
    // It's completely normal for several separate invoices to point at the
    // same Sales Order (partial shipments, staged billing, etc.) — matching
    // on the SO link alone would treat the 2nd and 3rd invoice against one
    // order as "duplicates" of the 1st and silently overwrite it instead of
    // creating them. Invoice Reference (your own invoice number) is the
    // only signal that's actually meant to be unique per invoice.
    dedupeKeys: [["partner_id", "ref"]],
    fields: [
      { name: "__row_id", label: "Row ID (not imported — your own tracking key, if the sheet has one)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "partner_id", label: "Customer (only needed for an invoice with no Sales Order — an SO-linked invoice pulls it from the order automatically)", aliases: ["customer", "customer name", "party", "buyer"], relation: true, relModel: "res.partner", autoCreate: true },
      { name: "__so_ref", label: "SO Reference (must match the Sales Order's Order Reference — leave the column out entirely to auto-compute the order from Invoice Reference instead)", aliases: ["so reference", "sales order reference", "so ref", "order reference", "customer reference", "so number", "sales order"], isSoReferenceField: true },
      { name: "__so_name", label: "SO Name (Odoo's own reference, e.g. S00088 — only needed if two orders share the same Order Reference)", aliases: ["so name", "so number (odoo)", "odoo so reference", "so odoo number", "sales order number"], isSoNameField: true },
      { name: "invoice_date", label: "Invoice Date", aliases: ["invoice date", "bill date", "date"], date: true },
      { name: "invoice_date_due", label: "Due Date", aliases: ["due date", "payment due date"], date: true },
      { name: "ref", label: "Invoice Reference (your own invoice number)", aliases: ["invoice reference", "invoice no", "invoice number", "reference"] },
      { name: "invoice_line_ids/product_id", label: "Line / Product", aliases: ["product", "item", "product name"], relation: true, relModel: "product.product", autoCreate: false, isLine: true },
      { name: "invoice_line_ids/quantity", label: "Line / Invoice Quantity", aliases: ["qty", "quantity", "invoice qty", "invoiced quantity", "units"], number: true, isLine: true },
      { name: "invoice_line_ids/price_unit", label: "Line / Unit Price (optional — leave blank to use the order's own rate for that product)", aliases: ["unit price", "rate", "price", "selling price"], number: true, isLine: true },
      { name: "invoice_line_ids/tax_ids", label: "Line / Taxes (optional — leave blank to use the order's own tax for that product)", aliases: ["taxes", "tax", "gst", "gst rate"], isLine: true, isGstTaxField: "sale" },
      { name: "__tds_section", label: "TDS Section (optional — leave blank for an invoice with no TDS)", aliases: ["tds section", "section", "section code", "tds", "u/s"], isTdsSectionField: true },
      { name: "__tds_base_amount", label: "TDS Base Amount (optional — leave blank to use this invoice's own Untaxed Amount)", aliases: ["tds base amount", "base amount", "taxable amount", "tds base"], number: true, isTdsBaseField: true },
      { name: "__tds_rate", label: "TDS Rate % (optional override — leave blank to use the standard rate for the Section)", aliases: ["tds rate", "tds %", "tds percentage"], number: true, isTdsRateField: true },
      { name: "__tcs_section", label: "TCS Section (optional — leave blank for an invoice with no TCS, e.g. \"206C(1H)\" or a keyword from the tax name)", aliases: ["tcs section", "tcs", "tcs code", "tcs section code"], isTcsSectionField: true },
      { name: "__tcs_rate", label: "TCS Rate % (optional override — leave blank to use the standard rate for the Section)", aliases: ["tcs rate", "tcs %", "tcs percentage"], number: true, isTcsRateField: true },
      { name: "__extra_line_item", label: "Extra Line Item — Product or Income Account (optional, e.g. \"Freight Recovered\")", aliases: ["extra line item", "extra charge", "additional line", "additional charge", "freight recovered"], isExtraLineField: "product" },
      { name: "__extra_line_qty", label: "Extra Line Item / Quantity (optional — defaults to 1)", aliases: ["extra line quantity", "extra charge quantity", "additional quantity"], number: true, isExtraLineField: "quantity" },
      { name: "__extra_line_price", label: "Extra Line Item / Amount", aliases: ["extra line amount", "extra charge amount", "additional amount"], number: true, isExtraLineField: "price" },
      { name: "__extra_line_tax", label: "Extra Line Item / Taxes (optional)", aliases: ["extra line taxes", "extra charge tax", "additional tax"], isExtraLineField: "tax" },
      { name: "__extra_line_desc", label: "Extra Line Item / Description (optional — defaults to the product/account name)", aliases: ["extra line description", "extra charge description", "additional description"], isExtraLineField: "desc" },
    ],
  },
  creditNotesDebtors: {
    label: "Credit Notes (Debtors)", model: "account.move", icon: "↩️",
    isCreditNoteModule: true,
    note: "For crediting a customer's balance — e.g. a short delivery, a return, or a pricing correction against an invoice already raised. Reversed Entry / Origin Invoice is matched two ways at once: the invoice's own Odoo number (e.g. INV/2026/00042) OR its Customer Reference field — in practice, the Customer Reference is usually what you actually have on hand, since an invoice created from a Sales Order automatically carries that order's own Order Reference forward into its Customer Reference field. Either one works; put in whichever number you know. If you already know it, the credit note links to it directly and its Line / Unit Price and Line / Taxes become optional (leave them blank to pull the invoice's own rate and tax for that product). If you don't have a matching number — or none exists yet — leave that column blank; the note still imports as a standalone Customer Credit Note, it just won't show as linked to a specific invoice in Odoo. Reason is kept on the note as an internal note — worth filling in, since it's usually the only record of WHY the note was issued (a short delivery, a rate correction, etc.). Every note is created as a draft, same as Odoo's own screen — review and post them afterward. Multi-line notes: rows below the first with a blank Customer are treated as extra lines of the same note.",
    defaults: { move_type: "out_refund" },
    dedupeKeys: [["partner_id", "ref"]],
    fields: [
      { name: "ref", label: "Note Reference (your own ID for this note, optional — used to avoid re-creating it if you import the same file twice)", aliases: ["id", "note reference", "note id", "reference", "credit note no"] },
      { name: "partner_id", label: "Customer", aliases: ["customer", "customer name", "party", "buyer"], required: true, relation: true, relModel: "res.partner", autoCreate: true },
      { name: "invoice_date", label: "Credit Note Date", aliases: ["credit note date", "date", "note date"], date: true },
      { name: "__origin_ref", label: "Reversed Entry / Origin Invoice (matches either that Sales Invoice's own Odoo number OR its Customer Reference field — leave blank if none exists yet)", aliases: ["reversed entry", "origin invoice", "reversed entry / origin invoice", "origin", "invoice reference", "customer reference"], isReversalRefField: true },
      { name: "invoice_line_ids/product_id", label: "Line / Product", aliases: ["product", "item", "credit note lines/product"], relation: true, relModel: "product.product", autoCreate: false, isLine: true },
      { name: "invoice_line_ids/quantity", label: "Line / Quantity", aliases: ["qty", "quantity", "credit note lines/quantity"], number: true, isLine: true },
      { name: "invoice_line_ids/price_unit", label: "Line / Unit Price (optional — leave blank to use the origin invoice's own rate for that product)", aliases: ["unit price", "rate", "price", "credit note lines/unit price"], number: true, isLine: true },
      { name: "invoice_line_ids/tax_ids", label: "Line / Taxes (optional — leave blank to use the origin invoice's own tax for that product)", aliases: ["taxes", "tax", "gst", "credit note lines/taxes"], isLine: true, isGstTaxField: "sale" },
      { name: "narration", label: "Reason (kept on the note as an internal note — your record of why it was issued)", aliases: ["reason", "narration", "remarks", "note"] },
    ],
  },
  debitNotesCreditors: {
    label: "Debit Notes (Creditors)", model: "account.move", icon: "↩️",
    isDebitNoteModule: true,
    note: "For reducing what you owe a vendor — e.g. a short receipt, a freight/expense claim back against them, or a pricing correction against a bill already booked. Despite the \"Debit Note\" name (that's the accounting term for crediting a Creditors/vendor balance), this creates an Odoo Vendor Credit Note (not a new Vendor Bill) — it reduces the vendor's balance, exactly matching what a Debit Note against a vendor is supposed to do. Reversed Entry / Origin Bill is matched two ways at once: the bill's own Odoo number OR its Vendor Reference field — in practice, the Vendor Reference is usually what you actually have on hand, since it's the same field the Vendor Bills module matches on. Either one works; put in whichever number you know. If you already know it, the note links to it directly and Line / Unit Price and Line / Taxes become optional (leave them blank to pull the bill's own rate and tax for that product). If you don't have a matching number — or none exists — leave that column blank; the note still imports as a standalone Vendor Credit Note. Line / Product or Expense Account accepts either: a real product (matched the same way as any other module), or a Chart of Accounts entry like \"Freight & Cartage\" for a line that isn't a billed product at all — tried as a product first, then as an account, auto-creating a new expense/income account only if neither matches (same behavior as the Opening Balances module's Account column). Reason is kept on the note as an internal note — usually the only record of why it was issued. Every note is created as a draft — review and post afterward. Multi-line notes: rows below the first with a blank Vendor are treated as extra lines of the same note.",
    defaults: { move_type: "in_refund" },
    dedupeKeys: [["partner_id", "ref"]],
    fields: [
      { name: "ref", label: "Note Reference (your own ID for this note, optional — used to avoid re-creating it if you import the same file twice)", aliases: ["id", "note reference", "note id", "reference", "debit note no"] },
      { name: "partner_id", label: "Vendor", aliases: ["vendor", "vendor name", "supplier", "supplier name"], required: true, relation: true, relModel: "res.partner", autoCreate: true },
      { name: "invoice_date", label: "Note Date", aliases: ["credit note date", "debit note date", "date", "note date"], date: true },
      { name: "__origin_ref", label: "Reversed Entry / Origin Bill (matches either that Vendor Bill's own Odoo number OR its Vendor Reference field — leave blank if none exists)", aliases: ["reversed entry", "origin bill", "reversed entry / origin bill", "origin", "bill reference", "vendor reference"], isReversalRefField: true },
      { name: "invoice_line_ids/product_id", label: "Line / Product or Expense Account", aliases: ["product or expense account", "product", "account", "expense account", "credit note lines/product or expense account"], isLine: true, isProductOrAccountField: true },
      { name: "invoice_line_ids/quantity", label: "Line / Quantity", aliases: ["qty", "quantity", "credit note lines/quantity"], number: true, isLine: true },
      { name: "invoice_line_ids/price_unit", label: "Line / Unit Price (optional — leave blank to use the origin bill's own rate for that product)", aliases: ["unit price", "rate", "price", "credit note lines/unit price"], number: true, isLine: true },
      { name: "invoice_line_ids/tax_ids", label: "Line / Taxes (optional — leave blank to use the origin bill's own tax for that product)", aliases: ["taxes", "tax", "gst", "credit note lines/taxes"], isLine: true, isGstTaxField: "purchase" },
      { name: "narration", label: "Reason (kept on the note as an internal note — your record of why it was issued)", aliases: ["reason", "narration", "remarks", "note"] },
    ],
  },
  tdsEntry: {
    label: "TDS Entry (Purchase & Sales)", model: "account.move", icon: "🧮",
    isTdsEntry: true,
    note: "For deducting TDS against a Vendor Bill OR a Customer Invoice that's already in Odoo (posted or draft) — works for both TDS types, and you don't need to say which: map Bill/Invoice Reference to either the Vendor Bill's Bill Reference or the Customer Invoice's Invoice Reference value (whichever record actually exists), and the module reads that record's own type to know which side it is. Runs Odoo's own native \"TDS Entry\" wizard against the record — the same one you'd get by opening it yourself and clicking \"TDS Entry\" — so it posts against whatever real account.tax record matches your TDS Section (filtered to Purchase-side or Sales-side taxes automatically), with the correct accounts and tax-grid tagging, not a custom line. If two records share the same reference, also map Bill/Invoice Name (Odoo's own number, e.g. BILL/2026/00042 or INV/2026/00042) or Vendor/Customer to pick the exact one. TDS Section can be the old Income-tax Act section code (194C, 194J, 194Q, ...) or the new-law section code/keyword — whatever matches a tax name or description in your own Configuration > Taxes; map a TDS Rate % column too if you want to override the matched tax's rate for a row, or to disambiguate when more than one tax matches the Section text. Base Amount is the amount TDS is calculated on; leave it blank to use the record's own Untaxed Amount automatically. For each row: the record is posted first if it wasn't already, then the TDS Entry wizard is created and applied — no manual step needed. If the TDS tax, section rate, or record match can't be resolved, that row is left alone and reported as an error rather than guessed at.",
    fields: [
      { name: "__row_id", label: "Row ID (not imported)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "__bill_ref", label: "Bill/Invoice Reference (matches the Vendor Bill's Bill Reference OR the Customer Invoice's Invoice Reference field)", aliases: ["bill reference", "invoice reference", "bill no", "vendor bill no", "vendor invoice no", "invoice number", "bill number"], required: true, isBillReferenceField: true },
      { name: "__bill_name", label: "Bill/Invoice Name (Odoo's own number, e.g. BILL/2026/00042 or INV/2026/00042 — only needed if two records share the same Reference)", aliases: ["bill name", "invoice name", "bill number (odoo)", "odoo bill number", "move name"], isBillNameField: true },
      { name: "partner_id", label: "Vendor/Customer (optional — only needed to disambiguate if two records share the same Reference)", aliases: ["vendor", "vendor name", "supplier", "supplier name", "customer", "customer name"], relation: true, relModel: "res.partner", autoCreate: false },
      { name: "__tds_section", label: "TDS Section", aliases: ["section", "tds section", "section code", "tds", "u/s"], required: true, isTdsSectionField: true },
      { name: "__base_amount", label: "Base Amount (optional — leave blank to use the record's own Untaxed Amount)", aliases: ["base amount", "taxable amount", "tds base", "basic amount", "untaxed amount"], number: true, isTdsBaseField: true },
      { name: "__tds_rate", label: "TDS Rate % (optional override — leave blank to use the standard rate for the Section)", aliases: ["tds rate", "rate", "percentage", "%", "tds %"], number: true, isTdsRateField: true },
    ],
  },
  billAdditionalLines: {
    label: "Bill Additional Lines", model: "account.move", icon: "➕",
    isBillAdditionalLines: true,
    note: "For adding extra line items (freight, loading charges, other cash expenses billed alongside the goods, etc.) onto a Vendor Bill that's already in Odoo — the usual lines come from the PO, but this covers anything billed on top of that. Map Bill Reference the same way the TDS Entry module does (map Bill Name and/or Vendor too if the reference alone is ambiguous). Multi-line bills: rows below the first with a blank Bill Reference are treated as extra lines of the same bill, same convention as the Vendor Bills module. Line / Product or Expense Account accepts either a real product or a Chart of Accounts entry like \"Freight & Cartage\" (auto-creating a new expense account if neither matches). For each Bill Reference group: the Bill is reset to Draft (if it was posted) and the new line(s) are added — it's then deliberately left in Draft, not re-posted, so you can review the addition before posting it yourself.",
    fields: [
      { name: "__row_id", label: "Row ID (not imported)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "__bill_ref", label: "Bill Reference (matches the Vendor Bill's own Bill Reference field)", aliases: ["bill reference", "bill no", "vendor bill no", "vendor invoice no", "invoice number", "bill number"], required: true, isGroupKey: true, isBillReferenceField: true },
      { name: "__bill_name", label: "Bill Name (Odoo's own number, e.g. BILL/2026/00042 — only needed if two bills share the same Bill Reference)", aliases: ["bill name", "bill number (odoo)", "odoo bill number", "move name"], isBillNameField: true },
      { name: "partner_id", label: "Vendor (optional — only needed to disambiguate if two bills share the same Bill Reference)", aliases: ["vendor", "vendor name", "supplier", "supplier name"], relation: true, relModel: "res.partner", autoCreate: false },
      { name: "invoice_line_ids/product_id", label: "Line / Product or Expense Account", aliases: ["product", "item", "product name", "product or expense account", "account", "expense account"], isLine: true, isProductOrAccountField: true },
      { name: "invoice_line_ids/quantity", label: "Line / Quantity", aliases: ["qty", "quantity", "units"], number: true, isLine: true },
      { name: "invoice_line_ids/price_unit", label: "Line / Unit Price", aliases: ["unit price", "rate", "price", "amount"], number: true, isLine: true },
      { name: "invoice_line_ids/tax_ids", label: "Line / Taxes (optional)", aliases: ["taxes", "tax", "gst", "gst rate"], isLine: true, isGstTaxField: "purchase" },
      { name: "invoice_line_ids/name", label: "Line / Description (optional — defaults to the product/account name)", aliases: ["description", "narration", "particulars", "remarks"], isLine: true },
    ],
  },
  invoiceAdditionalLines: {
    label: "Invoice Additional Lines", model: "account.move", icon: "➕",
    isInvoiceAdditionalLines: true,
    note: "Sales-side mirror of Bill Additional Lines — for adding extra line items (freight, packing, other charges billed alongside the goods/services, etc.) onto a Customer Invoice that's already in Odoo. Map Invoice Reference the same way the TDS Entry module does (map Invoice Name and/or Customer too if the reference alone is ambiguous). Multi-line invoices: rows below the first with a blank Invoice Reference are treated as extra lines of the same invoice, same convention as the Sales Invoices module. Line / Product or Income Account accepts either a real product or a Chart of Accounts entry like \"Freight Recovered\" (auto-creating a new income account if neither matches). For each Invoice Reference group: the invoice is reset to Draft (if it was posted) and the new line(s) are added — it's then deliberately left in Draft, not re-posted, so you can review the addition before posting it yourself.",
    fields: [
      { name: "__row_id", label: "Row ID (not imported)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "__bill_ref", label: "Invoice Reference (matches the Customer Invoice's own Invoice Reference field)", aliases: ["invoice reference", "invoice no", "invoice number", "customer invoice no", "bill reference", "reference"], required: true, isGroupKey: true, isBillReferenceField: true },
      { name: "__bill_name", label: "Invoice Name (Odoo's own number, e.g. INV/2026/00042 — only needed if two invoices share the same Invoice Reference)", aliases: ["invoice name", "invoice number (odoo)", "odoo invoice number", "move name"], isBillNameField: true },
      { name: "partner_id", label: "Customer (optional — only needed to disambiguate if two invoices share the same Invoice Reference)", aliases: ["customer", "customer name", "buyer", "buyer name"], relation: true, relModel: "res.partner", autoCreate: false },
      { name: "invoice_line_ids/product_id", label: "Line / Product or Income Account", aliases: ["product", "item", "product name", "product or income account", "account", "income account"], isLine: true, isProductOrAccountField: true },
      { name: "invoice_line_ids/quantity", label: "Line / Quantity", aliases: ["qty", "quantity", "units"], number: true, isLine: true },
      { name: "invoice_line_ids/price_unit", label: "Line / Unit Price", aliases: ["unit price", "rate", "price", "amount"], number: true, isLine: true },
      { name: "invoice_line_ids/tax_ids", label: "Line / Taxes (optional)", aliases: ["taxes", "tax", "gst", "gst rate"], isLine: true, isGstTaxField: "sale" },
      { name: "invoice_line_ids/name", label: "Line / Description (optional — defaults to the product/account name)", aliases: ["description", "narration", "particulars", "remarks"], isLine: true },
    ],
  },
  expenseJv: {
    label: "Expense JV", model: "account.move", icon: "🧾",
    note: "For an Expenses JV register — admin/corporate expense journal vouchers (professional fees, rent, freight, telephone, travelling, depository/compliance fees, labour charges, etc.), NOT trade purchases (those already have their own RM/Store/Service PO+Bill modules — don't use this for those). Built around 10 durable rules learned from a real batch that looked successful but wasn't (see below) — this always imports as a plain Journal Entry, NEVER as a vendor Bill, no matter how simple a voucher looks: a Bill only accepts invoice/tax lines, and a TDS line, an offsetting adjustment, or a Round Off line is silently dropped on one with no error — which is exactly what happened to an entire batch before this rule existed. Every entry posts to one explicit Miscellaneous/General journal (never a default) — set one up in Accounting first if you don't have exactly one, since more than one is treated as too ambiguous to guess between and blocks the whole sheet (or just pick the journal explicitly below, if this connection lists more than one). One row per account line of a voucher: Doc No., Doc Ref, Date, and Due Date are filled only on the first line, exactly like Cash Book's Voucher Reference blank-continuation convention — extra lines below just repeat the same voucher by leaving Doc No. blank. Particulars is the account name, optionally prefixed \"By\" (debit-side) or \"To\" (credit-side) — that prefix is read automatically when a line's own Debit/Credit cells are blank, and is always stripped before any account lookup (it's redundant with the actual amount columns, not part of the account name). Every account — Particulars line, TDS line, Round Off line, all of it — is resolved through an explicit mapping table stored inside this Odoo database itself (a System Parameter, editable from the results screen or Settings, Technical, System Parameters — not a live \"search the Chart of Accounts for something similar\" step, and not hardcoded in the app): text not yet in that table blocks the whole voucher rather than guessing, so a new account only ever needs mapping once per database, never a code change, and never a silent mis-post. A Particulars line reading \"SUNDRY CREDITORS\"/\"TRADE PAYABLES\" is always resolved by account TYPE (account_type = liability_payable), never by name — this is what stops two same-named accounts of different types (a real Payables control account and a stray same-named expense account) from ever being confused for each other — and only when that line actually names a Subledger/vendor; without one it's blocked, not guessed at. Due Date lands on that Sundry Creditors line's own due date (date_maturity), not as a header field, since a plain Journal Entry has no native due-date field the way an invoice does. Nothing is created for a voucher that doesn't self-balance (Total Debit = Total Credit across all of its lines, blank TDS/Round Off cells counted as 0) — that voucher is reported as a row error instead, checked before anything is submitted. Every entry is built as one single line_ids list and created in one call, never a header-first-then-append-lines sequence (Odoo rejects an unbalanced intermediate write even when the finished entry balances). Before creating anything, this checks for an already-imported entry using Doc No. (this sheet's own row identifier, e.g. \"EJ EX001\") against this journal — never the vendor's own Bill No./Doc Ref, since that's free text the vendor controls and different vendors commonly reuse the same short reference. After creating an entry, it's read back from Odoo to confirm the journal, line count, and total actually match what was submitted, rather than just trusting the create call succeeded. One Doc No. can carry several Subledger lines (e.g. a telecom bill split by phone number/department) — each stays its own line rather than being merged or split into separate entries. Narration is passed through exactly as given, including where the source report's own line-wrap has visibly garbled it — no reordering/cleanup heuristics are applied, since a \"fix\" that works for the common pattern risks corrupting the entries that don't follow it, and narration never affects any balance or posting anyway. This module does NOT auto-link a Credit/Debit Note to an Expense JV vendor just because the vendor name matches on both registers — check that manually via Doc Ref/amount. Every entry this module creates is a draft — review and post in Odoo afterward.",
    isExpenseJv: true,
    fields: [
      { name: "__row_id", label: "Row ID (not imported)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "__group", label: "Doc No.", aliases: ["doc no", "doc no.", "document no", "voucher no", "voucher number"], required: true, isGroupKey: true },
      { name: "__doc_ref", label: "Doc Ref (the vendor's own bill/invoice number — for your own reference only, never used for dedup/matching)", aliases: ["doc ref", "doc ref.", "bill ref", "bill reference", "vendor bill no", "reference"], isExpenseDocRefField: true },
      { name: "date", label: "Date", aliases: ["date"], date: true },
      { name: "__due_date", label: "Due Date (applied to the Sundry Creditors line's own due date, not a header field)", aliases: ["due date"], isExpenseDueDateField: true },
      { name: "__particulars", label: "Particulars (account name, optionally \"By\"/\"To\" prefixed)", aliases: ["particulars", "account", "ledger", "head"], required: true, isExpenseParticularsField: true },
      { name: "__subledger", label: "Subledger (the actual vendor/payee, when this line has one)", aliases: ["subledger", "party", "payee"], isExpenseSubledgerField: true },
      { name: "__debit", label: "Debit", aliases: ["debit"], number: true, isExpenseDebitField: true },
      { name: "__credit", label: "Credit", aliases: ["credit"], number: true, isExpenseCreditField: true },
      { name: "narration", label: "Narration (passed through exactly as given, never reordered/cleaned up)", aliases: ["narration", "note", "notes", "remarks"], isExpenseNarrationField: true },
    ],
  },
  cashBook: {
    label: "Cash Book", model: "account.move", icon: "💵",
    note: "For a Cash Book / Cash Payment-Receipt register — one row per voucher (Voucher Reference, Date, Type, Head (Account), Payee (Subledger), Amount, Narration). Type must read \"CP\"/\"Cash Paid\" or \"CR\"/\"Cash Received\". Each row becomes a two-line journal entry: the Head account for the row's Amount, balanced automatically against your Cash journal (found by looking for the one journal of Journal Type \"Cash\" in your Chart of Accounts — set one up in Accounting first if you don't have one) — you never map a separate Cash/Journal column. A Head of \"SUNDRY CREDITORS\"/\"SUNDRY DEBTORS\" means this voucher is settling a specific vendor bill/customer invoice — map the Payee column with the vendor/customer name, and if there's exactly one open bill/invoice on that partner for exactly this Amount (or, failing that, exactly one combination of 2-3 of that partner's open bills/invoices that sums to it — e.g. one payment settling two separate small bills at once), a real payment is registered against it/them automatically (Odoo's own bill-then-payment mechanism, as one combined Group Payment when it's more than one), so it comes back already reconciled — no manual step in Odoo afterward. If the match isn't unambiguous (nothing sums to this Amount, or more than one bill/combination does), or the payment registration itself fails, that row falls back to posting the amount to your real Payables/Receivables control account (never auto-created) instead — check the notes after upload for which rows needed that fallback, and reconcile those manually in Odoo (Accounting -> the account itself -> match open items). Any other Head is resolved against the Chart of Accounts the same way a Vendor Bill's Product/Expense Account line is (auto-creating a new expense account if nothing matches). Every plain journal entry is created as a draft — review and post in Odoo afterward; an auto-registered payment posts and reconciles immediately, same as clicking Register Payment yourself.",
    defaults: { move_type: "entry" },
    dedupeKeys: [["ref"]],
    isCashBook: true,
    fields: [
      { name: "__row_id", label: "Row ID (not imported)", aliases: ["id", "row id", "line id"], isIgnored: true },
      { name: "__group", label: "Voucher Reference", aliases: ["voucher reference", "voucher no", "voucher number", "voucher"], required: true, isGroupKey: true },
      { name: "date", label: "Date", aliases: ["date"], date: true },
      { name: "__cash_type", label: "Type (CP = Cash Paid, CR = Cash Received)", aliases: ["type", "voucher type"], required: true, isCashTypeField: true },
      { name: "__cash_head", label: "Head (Account)", aliases: ["head", "head (account)", "account", "ledger"], required: true, isCashHeadField: true },
      { name: "__cash_payee", label: "Payee (Subledger)", aliases: ["payee", "payee (subledger)", "subledger", "party"], isCashPayeeField: true },
      { name: "__cash_amount", label: "Amount", aliases: ["amount"], required: true, isCashAmountField: true },
      { name: "narration", label: "Narration", aliases: ["narration", "note", "notes", "remarks"] },
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

// Snapshot of every module's original `fields` array, taken once at module
// load — refreshCustomFields() resets ODOO_SCHEMAS[key].fields back to this
// before re-appending discovered custom fields, so repeated
// connects/refreshes never duplicate-append the same custom field twice.
const BASE_SCHEMA_FIELDS = {};
for (const key of Object.keys(ODOO_SCHEMAS)) {
  BASE_SCHEMA_FIELDS[key] = ODOO_SCHEMAS[key].fields.slice();
}


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
// Embedded pictures — a sheet where photos were literally dropped/pasted
// into cells (Excel's Insert > Picture, anchored over a cell) rather than
// typed as a URL. Those pictures live inside the .xlsx file itself (an
// .xlsx is a zip of XML + media files) — SheetJS's free build never reads
// them, so sheet_to_json() sees only the text columns and the pictures are
// invisible to it. This does its own minimal unzip + OOXML-drawing parse to
// find every embedded picture, work out which row it's anchored to, and
// turn it into base64 the URL-based path already knows how to hand to Odoo.
// All of this runs entirely client-side against the arrayBuffer already in
// the browser from the user's own file upload — no network fetch involved,
// so it's unaffected by CORS. Ported as-is from the extension's popup.js
// (functions inflateRaw/readZipEntries/readZipEntryData/resolveRelPath/
// relsPathFor/extractEmbeddedImagesBySheet, popup.js ~1786-1958).
//
// inflateRaw uses the browser-native DecompressionStream("deflate-raw") API
// — this is a zero-dependency Web API (supported in current Chrome, Edge,
// Firefox, Safari), not a bundled JS library, so no new package.json
// dependency was needed to port this.
// ---------------------------------------------------------------------------
function arrayBufferToBase64(buf) {
  const bytes = new Uint8Array(buf);
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function inflateRaw(bytes) {
  const ds = new DecompressionStream("deflate-raw");
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function readZipEntries(buf) {
  const view = new DataView(buf);
  const bytes = new Uint8Array(buf);
  let eocdOffset = -1;
  const searchFloor = Math.max(0, bytes.length - 65557);
  for (let i = bytes.length - 22; i >= searchFloor; i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocdOffset = i; break; }
  }
  if (eocdOffset === -1) throw new Error("not a valid .xlsx (zip end-of-directory not found)");
  const cdCount = view.getUint16(eocdOffset + 10, true);
  const cdOffset = view.getUint32(eocdOffset + 16, true);
  const entries = [];
  let ptr = cdOffset;
  for (let i = 0; i < cdCount; i++) {
    if (view.getUint32(ptr, true) !== 0x02014b50) break;
    const compMethod = view.getUint16(ptr + 10, true);
    const compSize = view.getUint32(ptr + 20, true);
    const nameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    const localHeaderOffset = view.getUint32(ptr + 42, true);
    const name = new TextDecoder("utf-8").decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen));
    entries.push({ name, compMethod, compSize, localHeaderOffset });
    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

async function readZipEntryData(buf, entry) {
  const view = new DataView(buf);
  const bytes = new Uint8Array(buf);
  const lh = entry.localHeaderOffset;
  if (view.getUint32(lh, true) !== 0x04034b50) throw new Error("corrupt zip local header");
  const nameLen = view.getUint16(lh + 26, true);
  const extraLen = view.getUint16(lh + 28, true);
  const dataStart = lh + 30 + nameLen + extraLen;
  const raw = bytes.subarray(dataStart, dataStart + entry.compSize);
  if (entry.compMethod === 0) return raw;
  if (entry.compMethod === 8) return inflateRaw(raw);
  throw new Error(`unsupported zip compression method ${entry.compMethod}`);
}

// Resolves an OOXML relationship Target ("../drawings/drawing1.xml") against
// the directory of the file that referenced it ("xl/worksheets").
function resolveRelPath(baseDir, target) {
  const parts = baseDir.split("/");
  for (const part of target.split("/")) {
    if (part === "..") parts.pop();
    else if (part !== ".") parts.push(part);
  }
  return parts.join("/");
}
function relsPathFor(path) {
  const parts = path.split("/");
  const file = parts.pop();
  return [...parts, "_rels", `${file}.rels`].join("/");
}

// Returns { [sheetName]: { [absoluteRowIndex0Based]: [{col, dataUri}, ...] } }
async function extractEmbeddedImagesBySheet(arrayBuffer, sheetNames) {
  let entries;
  try {
    entries = readZipEntries(arrayBuffer);
  } catch (e) {
    return {}; // not a real zip (e.g. a .xls or .csv) — nothing to extract, not an error
  }
  const byName = {};
  for (const e of entries) byName[e.name] = e;
  async function readText(path) {
    const e = byName[path];
    if (!e) return null;
    try {
      return new TextDecoder("utf-8").decode(await readZipEntryData(arrayBuffer, e));
    } catch (err) {
      return null;
    }
  }
  const parseXml = (text) => new DOMParser().parseFromString(text, "application/xml");

  const workbookXml = await readText("xl/workbook.xml");
  const workbookRels = await readText("xl/_rels/workbook.xml.rels");
  if (!workbookXml || !workbookRels) return {};

  const relMap = {};
  for (const rel of Array.from(parseXml(workbookRels).getElementsByTagName("Relationship"))) {
    relMap[rel.getAttribute("Id")] = rel.getAttribute("Target");
  }
  const sheetPathByName = {};
  for (const sheetEl of Array.from(parseXml(workbookXml).getElementsByTagName("sheet"))) {
    const rid = sheetEl.getAttribute("r:id");
    const target = relMap[rid];
    if (target) sheetPathByName[sheetEl.getAttribute("name")] = resolveRelPath("xl", target);
  }

  const result = {};
  for (const sheetName of sheetNames) {
    const wsPath = sheetPathByName[sheetName];
    if (!wsPath) continue;
    const wsXml = await readText(wsPath);
    if (!wsXml) continue;
    const drawingEl = parseXml(wsXml).getElementsByTagName("drawing")[0];
    if (!drawingEl) continue; // this sheet has no embedded pictures at all
    const drawingRid = drawingEl.getAttribute("r:id");

    const wsRelsXml = await readText(relsPathFor(wsPath));
    if (!wsRelsXml) continue;
    let drawingTarget = null;
    for (const rel of Array.from(parseXml(wsRelsXml).getElementsByTagName("Relationship"))) {
      if (rel.getAttribute("Id") === drawingRid) drawingTarget = rel.getAttribute("Target");
    }
    if (!drawingTarget) continue;
    const drawingPath = resolveRelPath(wsPath.substring(0, wsPath.lastIndexOf("/")), drawingTarget);

    const drawingXml = await readText(drawingPath);
    if (!drawingXml) continue;
    const drawingRelsXml = await readText(relsPathFor(drawingPath));
    const mediaByRid = {};
    if (drawingRelsXml) {
      for (const rel of Array.from(parseXml(drawingRelsXml).getElementsByTagName("Relationship"))) {
        mediaByRid[rel.getAttribute("Id")] = resolveRelPath(drawingPath.substring(0, drawingPath.lastIndexOf("/")), rel.getAttribute("Target"));
      }
    }

    const drawingDoc = parseXml(drawingXml);
    const anchors = [
      ...Array.from(drawingDoc.getElementsByTagName("xdr:twoCellAnchor")),
      ...Array.from(drawingDoc.getElementsByTagName("xdr:oneCellAnchor")),
    ];
    const byRow = {};
    for (const anchor of anchors) {
      const fromEl = anchor.getElementsByTagName("xdr:from")[0];
      const rowEl = fromEl && fromEl.getElementsByTagName("xdr:row")[0];
      if (!rowEl) continue;
      const row = parseInt(rowEl.textContent, 10);
      const colEl = fromEl.getElementsByTagName("xdr:col")[0];
      const col = colEl ? parseInt(colEl.textContent, 10) : 0;
      const blip = anchor.getElementsByTagName("a:blip")[0];
      const embedRid = blip && blip.getAttribute("r:embed");
      const mediaPath = embedRid && mediaByRid[embedRid];
      const mediaEntry = mediaPath && byName[mediaPath];
      if (!mediaEntry) continue;
      const ext = mediaPath.split(".").pop().toLowerCase();
      const mime = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" }[ext];
      if (!mime) continue; // vector formats (emf/wmf) some Excel versions embed aren't real photos — skip rather than guess
      let dataUri;
      try {
        const bytes = await readZipEntryData(arrayBuffer, mediaEntry);
        dataUri = `data:${mime};base64,${arrayBufferToBase64(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))}`;
      } catch (e) {
        continue;
      }
      byRow[row] = byRow[row] || [];
      byRow[row].push({ col, dataUri });
    }
    for (const row of Object.keys(byRow)) byRow[row].sort((a, b) => a.col - b.col);
    result[sheetName] = byRow;
  }
  return result;
}

function extractDriveFileId(url) {
  let m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return null;
}
// Drive exposes the same file through several different endpoints, and
// which ones actually return raw image bytes (vs. an HTML page, vs. a plain
// 403) varies by file and seemingly by Google's mood — "uc?export=download"
// alone throws a bare 403 for plenty of otherwise-public files. So
// urlToBase64 (below, inside createEngine) tries several known-working
// shapes in order, via the server-side image proxy, and uses whichever one
// actually returns an image, instead of betting on just one.
function driveCandidateUrls(id) {
  return [
    `https://drive.google.com/thumbnail?id=${id}&sz=w2000`,
    `https://lh3.googleusercontent.com/d/${id}=s2000`,
    `https://drive.google.com/uc?export=download&id=${id}`,
    `https://drive.google.com/uc?id=${id}`,
  ];
}
// A Google Drive FOLDER link (drive.google.com/drive/folders/FOLDER_ID) is
// not a file — it's a page that lists files, and reading that listing
// normally requires being signed into a Google account, which this app
// never is. The one way around that without OAuth is Google's Drive API v3
// "files.list", which — for a folder shared as "Anyone with the link" —
// works with just an API key (no sign-in). If the user has set one on this
// connection (Google Drive API Key field), a folder link expands into every
// image file inside it; if not, this reports exactly what's needed instead
// of silently failing.
function extractDriveFolderId(url) {
  const m = url.match(/\/drive\/folders\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

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

// Human-readable label per account_type, for the "what type should this new
// account be?" pre-flight prompt (see the isAccountField pre-flight check in
// uploadSheetToOdoo). Same key set as ACCOUNT_CODE_BASE, in the order
// Odoo's own Chart of Accounts form groups them.
export const ACCOUNT_TYPE_LABELS = {
  asset_receivable: "Receivable (Asset)",
  asset_cash: "Bank and Cash (Asset)",
  asset_current: "Current Assets",
  asset_non_current: "Non-current Assets",
  asset_prepayments: "Prepayments (Asset)",
  asset_fixed: "Fixed Assets",
  liability_payable: "Payable (Liability)",
  liability_credit_card: "Credit Card (Liability)",
  liability_current: "Current Liabilities",
  liability_non_current: "Non-current Liabilities",
  equity: "Equity",
  equity_unaffected: "Current Year Earnings (Equity)",
  income: "Income",
  income_other: "Other Income",
  expense: "Expenses",
  expense_depreciation: "Depreciation",
  expense_direct_cost: "Cost of Revenue",
  off_balance: "Off-Balance Sheet",
};

// ---------------------------------------------------------------------------
// File parsing — unchanged, still 100% client-side (SheetJS)
// ---------------------------------------------------------------------------
// Shared by parseWorkbookFile (a real uploaded .xlsx/.xls/.csv) and
// fetchGoogleSheet (a fetched Google Sheet, read as CSV text) — turns a
// parsed SheetJS workbook into this app's own `sheets` array shape, with
// the same embedded-image handling, blank-row filtering, and column
// analysis either path goes through equally.
function buildSheetsFromWorkbook(wb, embeddedBySheet = {}) {
  return wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const rawJson = XLSX.utils.sheet_to_json(ws, { defval: "" });
    if (!rawJson.length) return null;

    // Merge in any embedded pictures for this sheet BEFORE computing
    // the final header list. Each picture is written into the SAME
    // cell it's visually anchored over (e.g. a blank "Item Image"
    // column) — not a separate new column — because a sheet with
    // pasted-in pictures almost always already has a text column
    // named something like "Image"/"Photo" sitting right where the
    // pictures are (just empty, since the picture floats over the
    // cell rather than being cell content). Filling the original
    // cell in place means whatever field the sheet's own image
    // column maps to now actually has content. A picture with no
    // recognizable header at its column (or more than one picture
    // stacked in one row) falls back to synthetic "Embedded Main/
    // Extra Image" columns.
    const byRow = embeddedBySheet[name] || {};
    if (Object.keys(byRow).length) {
      const range = XLSX.utils.decode_range(ws["!ref"]);
      const headerRowAbs = range.s.r;
      const headerByCol = {};
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: headerRowAbs, c })];
        if (cell && cell.v !== undefined && cell.v !== "") headerByCol[c] = String(cell.v);
      }
      let usedMain = false, usedExtra = false;
      rawJson.forEach((row, idx) => {
        const imgs = byRow[headerRowAbs + 1 + idx];
        if (!imgs || !imgs.length) return;
        const extras = [];
        for (const img of imgs) {
          const headerName = headerByCol[img.col];
          if (headerName && (row[headerName] === "" || row[headerName] == null)) {
            row[headerName] = img.dataUri;
          } else {
            extras.push(img.dataUri);
          }
        }
        if (extras.length) {
          if (!usedMain && row["Embedded Main Image"] === undefined) {
            row["Embedded Main Image"] = extras.shift();
            usedMain = true;
          }
          if (extras.length) {
            row["Embedded Extra Images"] = extras.join(",");
            usedExtra = true;
          }
        }
      });
      if (usedMain) rawJson.forEach((row) => { if (row["Embedded Main Image"] === undefined) row["Embedded Main Image"] = ""; });
      if (usedExtra) rawJson.forEach((row) => { if (row["Embedded Extra Images"] === undefined) row["Embedded Extra Images"] = ""; });
    }

    const headers = Object.keys(rawJson[0]);
    const json = rawJson.filter((row) => !isRowEmpty(row, headers));
    if (!json.length) return null;
    return { name, headers, rows: json, analysis: analyzeSheet(headers), skipDuplicates: true, duplicateAction: "smart", stockLocationId: null };
  }).filter(Boolean);
}

export function parseWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target.result;
        const wb = XLSX.read(arrayBuffer, { type: "array" });

        // Best-effort: a sheet with pictures pasted into cells (rather than
        // typed as URLs) gets those pictures pulled out and attached to the
        // right row here — see extractEmbeddedImagesBySheet's comment above.
        // If this fails for any reason (unusual file structure, an older
        // browser without DecompressionStream("deflate-raw"), etc.) the
        // import still proceeds normally with whatever text columns are
        // there — this is never allowed to block the rest of the upload.
        let embeddedBySheet = {};
        try {
          embeddedBySheet = await extractEmbeddedImagesBySheet(arrayBuffer, wb.SheetNames);
        } catch (err) {
          console.error("Embedded image extraction failed:", err);
        }

        resolve(buildSheetsFromWorkbook(wb, embeddedBySheet));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Could not read the file"));
    reader.readAsArrayBuffer(file);
  });
}

function extractGoogleSheetId(url) {
  const m = String(url || "").match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : null;
}
function extractGoogleSheetGid(url) {
  const m = String(url || "").match(/[?#&]gid=(\d+)/);
  return m ? m[1] : null;
}

// Google Sheets import — fetches a publicly-shared sheet as CSV through
// Google's own export endpoint (no OAuth) and feeds it through the exact
// same column-mapping/upload pipeline as an uploaded file. The sheet's
// sharing must be "Anyone with the link" (Viewer is enough) — a private
// sheet's export endpoint redirects to a Google sign-in page instead of CSV
// data, which is caught below and reported as a clear error rather than a
// silent empty/garbled import.
//
// Unlike the Chrome extension version of this tool (privileged context,
// immune to CORS), a browser tab on this app's own domain cannot fetch()
// docs.google.com directly — so this goes through the same server-side
// proxy pattern as /api/image-proxy (see that file's own comment for the
// full reasoning), via its "googleSheet" action.
//
// Limitation: Google's export endpoint returns ONE tab per request (the
// gid in the URL, or the first/default tab if no gid is present) — a
// multi-tab workbook needs its tabs imported one at a time (click the tab
// in Google Sheets first, then copy the address bar — its gid changes per
// tab — and import that URL as its own pass).
export async function fetchGoogleSheet(url) {
  const sheetId = extractGoogleSheetId(url);
  if (!sheetId) {
    throw new Error('that doesn\'t look like a Google Sheets link — open the sheet, copy the URL from the address bar (it should contain "/spreadsheets/d/..."), and paste that.');
  }
  const gid = extractGoogleSheetGid(url);
  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv${gid ? `&gid=${gid}` : ""}`;

  const res = await fetch("/api/image-proxy", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action: "googleSheet", url: exportUrl }),
  });
  const data = await res.json();
  if (data.error) {
    throw new Error(`couldn't fetch this sheet (${data.error}) — make sure it's shared as "Anyone with the link" (Viewer access is enough), not just to specific people.`);
  }
  const csvText = data.csvText || "";
  if (/^\s*<!doctype html/i.test(csvText) || /^\s*<html/i.test(csvText)) {
    throw new Error('got a sign-in/permission page back instead of your sheet\'s data — it needs to be shared as "Anyone with the link" first.');
  }

  const wb = XLSX.read(csvText, { type: "string" });
  const built = buildSheetsFromWorkbook(wb, {});
  if (!built.length) throw new Error("the sheet came back empty, or only has a header row with no data below it — check you copied the right tab's link.");
  return { sheets: built, fileName: `Google Sheet (${gid ? `tab gid=${gid}` : "first tab"})` };
}

// ---------------------------------------------------------------------------
// Engine factory — one instance per active Odoo connection. All network
// calls go through /api/odoo-proxy (server-side), never directly to Odoo.
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Expense JV / Cash Book shared constants — see the expenseJv schema's own
// `note` for full background. Kept at module scope (not inside createEngine)
// since they're pure text patterns / a fixed System Parameter key, not
// per-connection state.
// ---------------------------------------------------------------------------
const SUNDRY_CREDITORS_RE = /sundry\s*creditors|trade\s*payables?/i;
const SUNDRY_DEBTORS_RE = /sundry\s*debtors|trade\s*receivables?/i;
const ROUND_OFF_RE = /round\s*off/i;
const EXPENSE_JV_MAP_PARAM_KEY = "odoo_auto_import.expense_jv_account_map";
const EXPENSE_JV_ROUND_OFF_KEYS = { credit: "round off (credit)", debit: "round off (debit)" };

// Maps one discovered Odoo custom (x_...) field to a schema field
// descriptor, reusing the same flag vocabulary every built-in field uses.
function buildCustomFieldDescriptor(name, info) {
  const label = info.string || name;
  const niceName = name.replace(/^x_studio_/, "").replace(/^x_/, "").replace(/_/g, " ");
  const aliases = [...new Set([label, niceName, name].map((s) => String(s).toLowerCase()).filter(Boolean))];
  const base = { name, label: `${label} (custom)`, aliases, custom: true };
  switch (info.type) {
    case "char":
    case "text":
    case "html":
      return base;
    case "integer":
    case "float":
    case "monetary":
      return { ...base, number: true };
    case "boolean":
      return { ...base, boolean: true };
    case "date":
    case "datetime":
      return { ...base, date: true };
    case "many2one":
      return info.relation ? { ...base, relation: true, relModel: info.relation, autoCreate: false } : null;
    case "many2many":
      return info.relation ? { ...base, relation: true, relModel: info.relation, autoCreate: false, isTagField: true } : null;
    case "selection": {
      if (!Array.isArray(info.selection)) return null;
      const map = {};
      for (const [value, selLabel] of info.selection) {
        map[String(value).toLowerCase()] = value;
        map[String(selLabel).toLowerCase()] = value;
      }
      return { ...base, dynamicSelectionMap: map };
    }
    default:
      return null; // one2many, binary, reference, etc. aren't sensibly importable from a flat sheet row
  }
}

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
  const expenseAccountCache = new Map();
  const tdsTaxCache = new Map();
  const tcsTaxCache = new Map();
  const controlAccountCache = new Map(); // account_type -> {id, created, note}
  const expenseJvCodeAccountCache = new Map(); // account code -> {id} | {id:false, note}
  let expenseJvMappingCache = null;      // { "particulars text": "account code", ... } for the active database
  let expenseJvMappingParamId = null;    // ir.config_parameter record id, once known, so saves can `write` instead of re-`search`
  let expenseJvPayableAccountResult;     // undefined = not yet resolved this session

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

  // ---------------------------------------------------------------------------
  // Image import — Main Image / Extra Media URLs (incl. Google Drive links
  // and Drive folder links) and embedded pictures (handled separately, see
  // extractEmbeddedImagesBySheet above — those arrive here already as a
  // data: URI, no network needed).
  //
  // CORS investigation (done before writing any of this): the extension
  // fetches Drive/image URLs directly from its own privileged context
  // (manifest.json host_permissions: "https://*/*", "http://*/*"), which
  // bypasses CORS entirely — a browser extension isn't a normal web page.
  // A page served from this app's own domain has no such privilege: a
  // browser fetch() straight to drive.google.com / lh3.googleusercontent.com
  // / an arbitrary image host is blocked by the browser's CORS check, since
  // none of those origins send back an Access-Control-Allow-Origin header
  // for this app's domain (confirmed — this is the same reason the existing
  // /api/odoo-proxy route exists for Odoo itself). So every network fetch
  // of image bytes below goes through a new server route, /api/image-proxy
  // (server-to-server calls aren't subject to CORS at all, same fix as the
  // Odoo proxy). Only the local, already-in-memory embedded-picture bytes
  // (extractEmbeddedImagesBySheet, module-level above) skip the proxy,
  // since those never touch the network.
  //
  // Google's own Drive API v3 (googleapis.com/drive/v3/files) does appear to
  // allow direct browser CORS calls with just an API key in many cases, but
  // this routes the folder-listing call through the same server proxy too
  // rather than relying on that — one code path, no special-casing, and no
  // risk of a future Drive API CORS-policy change silently breaking folder
  // imports.
  async function fetchImageProxy(body) {
    const res = await fetch("/api/image-proxy", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`);
    return data;
  }

  async function urlToBase64(rawUrl) {
    const trimmed = String(rawUrl).trim();
    // An embedded picture (see extractEmbeddedImagesBySheet) arrives here as
    // a data: URI, not a real URL — its bytes are already sitting right
    // there in the string, so pull the base64 out directly instead of
    // round-tripping it through the proxy for nothing.
    const dataUriMatch = trimmed.match(/^data:[^;]+;base64,(.*)$/s);
    if (dataUriMatch) return dataUriMatch[1];
    const isDrive = /drive\.google\.com|docs\.google\.com/i.test(trimmed);
    if (isDrive && /\/drive\/folders\//.test(trimmed)) {
      throw new Error('this is a Google Drive FOLDER link, which holds multiple images — a single-image field can only take one file link. Use the eCommerce Extra Images column for a whole folder (it expands a folder link into every image inside it), or pick one specific image\'s own share link for this field.');
    }
    const driveId = isDrive ? extractDriveFileId(trimmed) : null;
    const candidates = driveId ? driveCandidateUrls(driveId) : [trimmed];

    let lastError;
    for (const url of candidates) {
      try {
        const { base64 } = await fetchImageProxy({ action: "fetch", url });
        return base64;
      } catch (e) {
        lastError = e;
      }
    }
    if (driveId) {
      throw new Error(`couldn't fetch this Google Drive file through any known link format (last error: ${lastError.message}) — double check it's shared as "Anyone with the link", not just to specific people`);
    }
    throw new Error(lastError.message);
  }

  async function listDriveFolderImages(folderId) {
    if (!connection.driveApiKey) {
      throw new Error('this is a Google Drive FOLDER link — to import every image inside it automatically, add a "Google Drive API Key" on this database\'s connection (edit the database). Without one, share each image file individually instead (open the image, Share, Copy link) and list those file links in the cell separated by commas.');
    }
    const files = [];
    let pageToken = "";
    for (let page = 0; page < 10; page++) { // 10 pages x 100 = up to 1000 images, far beyond any real product gallery
      const { files: pageFiles, nextPageToken, error } = await fetchImageProxy({ action: "driveList", folderId, apiKey: connection.driveApiKey, pageToken });
      if (error) throw new Error(`Drive API folder lookup failed: ${error} — check the API key is valid and the Drive API is enabled for it in Google Cloud Console`);
      files.push(...(pageFiles || []));
      if (!nextPageToken) break;
      pageToken = nextPageToken;
    }
    if (!files.length) {
      throw new Error("no images found in this Drive folder — check it's shared as \"Anyone with the link\" and actually contains image files");
    }
    return files; // [{id, name}, ...]
  }

  // eCommerce "Extra Media" images — each URL in the cell becomes its own
  // product.image record linked to the template (product_tmpl_id), which is
  // exactly what Odoo's own "Extra Media" gallery on the product form
  // displays. A cell can mix plain URLs, individual Drive file links, AND
  // Drive folder links in one go — a folder link expands into one
  // product.image per image file found inside it (see listDriveFolderImages
  // above). Applied after the product template exists (create or matched
  // duplicate), same timing as On Hand Quantity/Relocate.
  async function applyExtraImages(templateId, urls, notes) {
    let i = 0;
    for (const url of urls) {
      const folderId = /drive\.google\.com/i.test(url) ? extractDriveFolderId(url) : null;
      const isFolder = folderId && /\/drive\/folders\//.test(url);
      if (isFolder) {
        try {
          const files = await listDriveFolderImages(folderId);
          for (const f of files) {
            i++;
            try {
              const base64 = await urlToBase64(`https://drive.google.com/uc?id=${f.id}`);
              await odooExecute("product.image", "create", [{ product_tmpl_id: templateId, name: f.name || `Image ${i}`, image_1920: base64 }]);
            } catch (e) {
              notes.push(`Could not add Extra Media image "${f.name}" from Drive folder for product #${templateId}: ${e.message}`);
            }
          }
        } catch (e) {
          notes.push(`Could not read Drive folder for product #${templateId}: ${e.message}`);
        }
        continue;
      }
      i++;
      try {
        const base64 = await urlToBase64(url);
        await odooExecute("product.image", "create", [{ product_tmpl_id: templateId, name: `Image ${i}`, image_1920: base64 }]);
      } catch (e) {
        notes.push(`Could not add Extra Media image from "${url}" for product #${templateId}: ${e.message}`);
      }
    }
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

  // `explicitType` (optional): an account_type the user picked themselves
  // for this exact name, from the "what type should this new account be?"
  // pre-flight prompt (see the isAccountField pre-flight check in
  // uploadSheetToOdoo). Used verbatim instead of classifyAccountType's
  // keyword guess — only consulted on the create path, a name that already
  // matches an existing account is reused as-is regardless.
  async function resolveOrCreateAccount(name, hasDebit, hasCredit, notes, explicitType) {
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
        const accountType = explicitType || classifyAccountType(String(name), hasDebit, hasCredit);
        const code = await nextFreeAccountCode(accountType);
        const newId = await odooExecute("account.account", "create", [{ name: String(name), code, account_type: accountType }]);
        result = { id: newId, created: true, note: `Created new account "${name}" — code ${code}, type "${accountType}"${explicitType ? " (you picked this)" : " (auto-detected from the name)"}. Review it in Chart of Accounts.` };
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

  // Resolves a line item's "Product or Expense/Income Account" column
  // (Bill/Invoice Additional Lines) against a real product first, falling
  // back to a Chart of Accounts entry (auto-creating a new expense/income
  // account, detected from keywords in the name, if neither matches).
  async function resolveOrCreateExpenseAccount(name, notes) {
    if (!name) return { id: false, created: false, note: null };
    const cacheKey = String(name).toLowerCase();
    if (expenseAccountCache.has(cacheKey)) {
      const cached = expenseAccountCache.get(cacheKey);
      if (cached.note) notes.push(cached.note);
      return cached;
    }
    // Safety net: never mint a duplicate Payables/Receivables control
    // account just because a Product-or-Account text cell happened to read
    // "Sundry Creditors"/"Sundry Debtors" — route those to the real
    // account_type-resolved control account instead (see Expense JV's
    // resolveExpenseJvLineAccount, which hits this same class of bug).
    if (SUNDRY_CREDITORS_RE.test(String(name))) {
      const result = await resolveControlAccount("liability_payable", "Payables/Creditors", notes);
      expenseAccountCache.set(cacheKey, result);
      return result;
    }
    if (SUNDRY_DEBTORS_RE.test(String(name))) {
      const result = await resolveControlAccount("asset_receivable", "Receivables/Debtors", notes);
      expenseAccountCache.set(cacheKey, result);
      return result;
    }
    let result;
    try {
      let ids = await odooExecute("account.account", "search", [[["name", "=", String(name)]]], { limit: 1 });
      if (!ids.length) ids = await odooExecute("account.account", "search", [[["name", "ilike", String(name)]]], { limit: 1 });
      if (ids.length) {
        result = { id: ids[0], created: false, note: null };
      } else {
        const accountType = /\b(income|recovery|recovered|discount received)\b/i.test(String(name)) ? "income" : "expense";
        const code = await nextFreeAccountCode(accountType);
        const newId = await odooExecute("account.account", "create", [{ name: String(name), code, account_type: accountType }]);
        result = { id: newId, created: true, note: `Created new account "${name}" — code ${code}, type "${accountType}" (auto-detected from the name). Review it in Chart of Accounts.` };
      }
    } catch (e) {
      result = { id: false, created: false, note: `Account lookup/creation failed for "${name}": ${e.message}` };
    }
    expenseAccountCache.set(cacheKey, result);
    if (result.note) notes.push(result.note);
    return result;
  }

  // Resolves a TDS Section (+ optional rate override) to a real account.tax
  // record, filtered to Purchase-side or Sales-side TDS taxes automatically.
  async function resolveTdsTax(sectionRaw, rateOverrideRaw, side, notes) {
    const raw = cleanText(String(sectionRaw || ""));
    if (!raw) return { id: false, error: "TDS Section is required." };
    const l10nType = side === "sales" ? "tds_sale" : "tds_purchase";
    const cacheKey = `${l10nType}::${raw.toLowerCase()}::${rateOverrideRaw ?? ""}`;
    if (tdsTaxCache.has(cacheKey)) return tdsTaxCache.get(cacheKey);

    const rateOverride = rateOverrideRaw !== null && rateOverrideRaw !== undefined && rateOverrideRaw !== "" ? cleanNumber(rateOverrideRaw) : null;
    let result;
    try {
      const baseDomain = [["l10n_in_tax_type", "=", l10nType], ["active", "=", true]];
      const withRate = (domain) => (rateOverride ? [...domain, ["amount", "=", -Math.abs(rateOverride)]] : domain);

      let candidates = await odooExecute("account.tax", "search_read", [withRate([...baseDomain, ["name", "ilike", raw]]), ["id", "name", "amount"]]);
      if (!candidates.length) candidates = await odooExecute("account.tax", "search_read", [withRate([...baseDomain, ["description", "ilike", raw]]), ["id", "name", "amount"]]);
      if (!candidates.length && rateOverride) candidates = await odooExecute("account.tax", "search_read", [withRate(baseDomain), ["id", "name", "amount"]]);

      if (!candidates.length) {
        result = { id: false, error: `No ${side === "sales" ? "TDS Sale" : "TDS Purchase"} tax found matching "${raw}"${rateOverride ? ` at ${rateOverride}%` : ""}. Check Accounting > Configuration > Taxes for the exact section/name, or map a TDS Rate % column to narrow it down.` };
      } else if (candidates.length > 1) {
        result = { id: false, error: `${candidates.length} ${side === "sales" ? "TDS Sale" : "TDS Purchase"} taxes match "${raw}": ${candidates.map((c) => c.name).join(", ")} — make the TDS Section value more specific (or map a TDS Rate % column) to pick one.` };
      } else {
        const tax = candidates[0];
        result = { id: tax.id, name: tax.name, rate: Math.abs(tax.amount), section: raw };
      }
    } catch (e) {
      result = { id: false, error: `TDS tax lookup failed for "${raw}": ${e.message}` };
    }
    tdsTaxCache.set(cacheKey, result);
    return result;
  }

  // Finds the account.tax record a TCS Section value refers to — same
  // matching strategy as resolveTdsTax (Tax Name, then Description, each
  // optionally narrowed by a rate override), but scoped to Odoo's own
  // l10n_in_tax_type = "tcs" instead of tds_sale/tds_purchase. TCS is always
  // sales-side (the seller collects it from the buyer), so there's no
  // "side" parameter here.
  async function resolveTcsTax(sectionRaw, rateOverrideRaw, notes) {
    const raw = cleanText(String(sectionRaw || ""));
    if (!raw) return { id: false, error: "TCS Section is required." };
    const cacheKey = `${raw.toLowerCase()}::${rateOverrideRaw ?? ""}`;
    if (tcsTaxCache.has(cacheKey)) return tcsTaxCache.get(cacheKey);

    const rateOverride = rateOverrideRaw !== null && rateOverrideRaw !== undefined && rateOverrideRaw !== "" ? cleanNumber(rateOverrideRaw) : null;
    let result;
    try {
      const baseDomain = [["l10n_in_tax_type", "=", "tcs"], ["active", "=", true]];
      const withRate = (domain) => (rateOverride ? [...domain, ["amount", "=", Math.abs(rateOverride)]] : domain);

      let candidates = await odooExecute("account.tax", "search_read", [withRate([...baseDomain, ["name", "ilike", raw]]), ["id", "name", "amount"]]);
      if (!candidates.length) candidates = await odooExecute("account.tax", "search_read", [withRate([...baseDomain, ["description", "ilike", raw]]), ["id", "name", "amount"]]);
      if (!candidates.length && rateOverride) candidates = await odooExecute("account.tax", "search_read", [withRate(baseDomain), ["id", "name", "amount"]]);

      if (!candidates.length) {
        result = { id: false, error: `No TCS tax found matching "${raw}"${rateOverride ? ` at ${rateOverride}%` : ""}. Check Accounting > Configuration > Taxes for the exact section/name, or map a TCS Rate % column to narrow it down.` };
      } else if (candidates.length > 1) {
        result = { id: false, error: `${candidates.length} TCS taxes match "${raw}": ${candidates.map((c) => c.name).join(", ")} — make the TCS Section value more specific (or map a TCS Rate % column) to pick one.` };
      } else {
        const tax = candidates[0];
        result = { id: tax.id, name: tax.name, rate: Math.abs(tax.amount), section: raw };
      }
    } catch (e) {
      result = { id: false, error: `TCS tax lookup failed for "${raw}": ${e.message}` };
    }
    tcsTaxCache.set(cacheKey, result);
    return result;
  }

  // Adds the resolved TCS tax onto every product/account line of an
  // invoice — TCS is a real tax charged on the sale (like GST), not a
  // withholding entry, so it's just appended to each line's own tax_ids
  // (command [4, id, 0], on top of whatever GST tax that line already has)
  // rather than run through a wizard. Must happen while the invoice is
  // still editable (before any TDS step posts it). Returns the tax name on
  // success, or null (with a note pushed) if the section couldn't be
  // resolved.
  async function applyTcsToInvoiceLines(moveId, moveName, tcs, notes) {
    const taxInfo = await resolveTcsTax(tcs.section, tcs.rate, notes);
    if (taxInfo.id === false) {
      notes.push(`"${moveName}": ${taxInfo.error} — TCS was not applied.`);
      return null;
    }
    try {
      const lines = await odooExecute("account.move.line", "search_read", [[["move_id", "=", moveId], ["display_type", "=", "product"]], ["id"]]);
      if (!lines.length) {
        notes.push(`"${moveName}": no invoice lines found to apply TCS "${taxInfo.name}" to.`);
        return null;
      }
      const lineCommands = lines.map((l) => [1, l.id, { tax_ids: [[4, taxInfo.id]] }]);
      await withRetry(() => odooExecute("account.move", "write", [[moveId], { invoice_line_ids: lineCommands }]));
      return taxInfo.name;
    } catch (e) {
      notes.push(`"${moveName}": applying TCS "${taxInfo.name}" failed — ${e.message}.`);
      return null;
    }
  }

  // Creates and posts the native TDS withholding entry against an already-
  // Posted Vendor Bill or Customer Invoice — the RPC equivalent of opening
  // the record, clicking "TDS Entry", filling in Section/Base Amount, and
  // clicking "Apply TDS". moveId MUST already be Posted; callers post first.
  async function applyNativeTdsWithhold(moveId, moveName, amountUntaxed, tds, side, notes) {
    const taxInfo = await resolveTdsTax(tds.section, tds.rate, side, notes);
    if (taxInfo.id === false) {
      notes.push(`"${moveName}": ${taxInfo.error} — TDS was not applied.`);
      return null;
    }
    const baseAmount = tds.base != null && tds.base > 0 ? tds.base : amountUntaxed;
    const tdsAmount = Math.round(baseAmount * (taxInfo.rate / 100) * 100) / 100;
    if (!tdsAmount || tdsAmount <= 0) {
      notes.push(`"${moveName}": calculated TDS amount is zero for "${taxInfo.name}" (${taxInfo.rate}% on base ${baseAmount}) — TDS was not applied.`);
      return null;
    }
    try {
      const context = { active_model: "account.move", active_id: moveId, active_ids: [moveId] };
      const wizardId = await odooExecute("l10n_in.withhold.wizard", "create", [{
        related_move_id: moveId, tax_id: taxInfo.id, base: baseAmount, amount: tdsAmount, tds_deduction: "normal",
      }], { context });
      await withRetry(() => odooExecute("l10n_in.withhold.wizard", "action_create_and_post_withhold", [[wizardId]], { context }));
      return { amount: tdsAmount, section: taxInfo.section, taxName: taxInfo.name };
    } catch (e) {
      notes.push(`"${moveName}": native TDS Entry ("${taxInfo.name}") failed — ${e.message}. TDS was not applied.`);
      return null;
    }
  }

  // Finds the single Vendor Bill or Customer Invoice a TDS Entry / Bill or
  // Invoice Additional Lines row refers to. Matches on Bill/Invoice
  // Reference ("ref" — same field Vendor Bills/Sales Invoices dedupe on),
  // narrowed by Bill/Invoice Name and/or Vendor/Customer if more than one
  // record shares that reference. Searches both purchase and sales move
  // types at once — side is read off whichever one actually matches.
  async function findTargetBillOrInvoice(refRaw, nameRaw, partnerId, notes) {
    const ref = cleanText(String(refRaw ?? ""));
    if (!ref) return { id: false, error: "Bill/Invoice Reference is required to locate the record." };
    const domain = [["move_type", "in", ["in_invoice", "in_refund", "out_invoice", "out_refund"]], ["ref", "=", ref]];
    if (partnerId) domain.push(["partner_id", "=", partnerId]);
    let ids = await odooExecute("account.move", "search", [domain]);
    if (!ids.length) {
      return { id: false, error: `No Vendor Bill or Customer Invoice found with Reference "${ref}"${partnerId ? " for that Vendor/Customer" : ""}.` };
    }
    if (ids.length > 1) {
      const name = cleanText(String(nameRaw ?? ""));
      if (name) {
        const narrowed = await odooExecute("account.move", "search", [[...domain, ["name", "=", name]]]);
        if (narrowed.length) ids = narrowed;
      }
    }
    if (ids.length > 1) {
      return { id: false, error: `${ids.length} records match Reference "${ref}" — map Bill/Invoice Name (or Vendor/Customer) to pick the exact one.` };
    }
    const [move] = await odooExecute("account.move", "read", [ids, ["name", "state", "amount_untaxed", "partner_id", "move_type"]]);
    return { id: move.id, bill: move, side: move.move_type.startsWith("out_") ? "sales" : "purchase" };
  }

  // Resets a move to Draft only if it isn't already there.
  async function resetBillToDraft(billId, currentState, notes) {
    if (currentState === "draft") return;
    await withRetry(() => odooExecute("account.move", "button_draft", [[billId]]));
    notes.push(`Record was in "${currentState}" — reset to Draft to apply the change.`);
  }

  // TDS Entry module — deducts TDS against a Vendor Bill or Customer
  // Invoice that already exists in Odoo. Reports progress the same way
  // uploadSheetToOdoo does (via onProgress), and returns the same
  // {status:'done', result:{...}} shape.
  async function processTdsEntrySheet(sheet, onProgress) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const billRefM = mappedFields.find((m) => m.field.isBillReferenceField);
    const billNameM = mappedFields.find((m) => m.field.isBillNameField);
    const partnerM = mappedFields.find((m) => m.field.name === "partner_id");
    const sectionM = mappedFields.find((m) => m.field.isTdsSectionField);
    const baseM = mappedFields.find((m) => m.field.isTdsBaseField);
    const rateM = mappedFields.find((m) => m.field.isTdsRateField);

    const usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));

    if (!billRefM || !sectionM) {
      return emit({ status: "done", result: { total: 0, created_count: 0, duplicate_count: 0, failed_count: 1, errors: [{ row: 0, error: "Map both Bill Reference and TDS Section columns before importing.", data: {} }], notes: [] } });
    }

    emit({ status: "uploading", progress: 0, total: usableRows.length });

    const notes = [];
    const errors = [];
    let successCount = 0;

    for (let i = 0; i < usableRows.length; i++) {
      const row = usableRows[i];
      try {
        const billRefRaw = row[billRefM.header];
        const sectionRaw = row[sectionM.header];
        if (!billRefRaw || !sectionRaw) {
          errors.push({ row: i, error: "Bill Reference and TDS Section are both required for this row.", data: row });
          continue;
        }

        let partnerId = null;
        if (partnerM && row[partnerM.header]) {
          const p = await resolveRelation(partnerM.field, row[partnerM.header], false);
          if (p.id) partnerId = p.id;
        }

        const lookup = await findTargetBillOrInvoice(billRefRaw, billNameM ? row[billNameM.header] : "", partnerId, notes);
        if (lookup.id === false) {
          errors.push({ row: i, error: lookup.error, data: row });
          continue;
        }
        const bill = lookup.bill;

        const baseAmountOverride = baseM && row[baseM.header] !== "" && row[baseM.header] != null ? cleanNumber(row[baseM.header]) : null;
        const rateOverride = rateM ? row[rateM.header] : null;

        if (bill.state !== "posted") {
          await withRetry(() => odooExecute("account.move", "action_post", [[bill.id]]));
        }
        const [posted] = await odooExecute("account.move", "read", [[bill.id], ["name", "amount_untaxed"]]);
        const result = await applyNativeTdsWithhold(bill.id, posted.name, posted.amount_untaxed, { section: sectionRaw, base: baseAmountOverride, rate: rateOverride }, lookup.side, notes);
        if (!result) {
          errors.push({ row: i, error: `Could not apply TDS to "${bill.name}" — see the note above for why.`, data: row });
          continue;
        }

        successCount++;
        notes.push(`${lookup.side === "sales" ? "Invoice" : "Bill"} "${bill.name}": TDS Entry posted via "${result.taxName}" — ${result.amount}.`);
      } catch (e) {
        errors.push({ row: i, error: e.message, data: row });
      }
      emit({ status: "uploading", progress: i + 1, total: usableRows.length });
    }

    return emit({ status: "done", result: { total: usableRows.length, created_count: successCount, duplicate_count: 0, failed_count: errors.length, errors, notes } });
  }

  // Shared by Bill Additional Lines (side="purchase", against Vendor Bills)
  // and Invoice Additional Lines (side="sales", against Customer Invoices).
  async function processAdditionalLinesSheet(sheet, onProgress, side) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const recordLabel = side === "sales" ? "Invoice" : "Bill";
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const billRefM = mappedFields.find((m) => m.field.isBillReferenceField);
    const billNameM = mappedFields.find((m) => m.field.isBillNameField);
    const partnerM = mappedFields.find((m) => m.field.name === "partner_id");
    const productM = mappedFields.find((m) => m.field.isProductOrAccountField);
    const qtyM = mappedFields.find((m) => m.field.name === "invoice_line_ids/quantity");
    const priceM = mappedFields.find((m) => m.field.name === "invoice_line_ids/price_unit");
    const taxM = mappedFields.find((m) => m.field.isGstTaxField);
    const descM = mappedFields.find((m) => m.field.name === "invoice_line_ids/name");

    const usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));

    if (!billRefM || !productM) {
      return emit({ status: "done", result: { total: 0, created_count: 0, duplicate_count: 0, failed_count: 1, errors: [{ row: 0, error: `Map both ${recordLabel} Reference and Line / Product or ${side === "sales" ? "Income" : "Expense"} Account columns before importing.`, data: {} }], notes: [] } });
    }

    // Blank-continuation grouping: a row with no Reference is an extra line
    // of the most recently referenced record above it.
    const groups = [];
    let current = null;
    for (let i = 0; i < usableRows.length; i++) {
      const row = usableRows[i];
      const refVal = row[billRefM.header];
      if (refVal !== "" && refVal != null) {
        current = { billRefRaw: refVal, billNameRaw: billNameM ? row[billNameM.header] : "", partnerRaw: partnerM ? row[partnerM.header] : "", firstRow: i, rows: [row] };
        groups.push(current);
      } else if (current) {
        current.rows.push(row);
      } else {
        groups.push({ billRefRaw: "", firstRow: i, rows: [row] });
      }
    }

    emit({ status: "uploading", progress: 0, total: groups.length });

    const notes = [];
    const errors = [];
    let successCount = 0;

    for (let g = 0; g < groups.length; g++) {
      const group = groups[g];
      try {
        if (!group.billRefRaw) {
          errors.push({ row: group.firstRow, error: `No ${recordLabel} Reference given for this row (and no earlier row in the sheet to carry it forward from).`, data: group.rows[0] });
          continue;
        }

        let partnerId = null;
        if (partnerM && group.partnerRaw) {
          const p = await resolveRelation(partnerM.field, group.partnerRaw, false);
          if (p.id) partnerId = p.id;
        }

        const lookup = await findTargetBillOrInvoice(group.billRefRaw, group.billNameRaw, partnerId, notes);
        if (lookup.id === false) {
          errors.push({ row: group.firstRow, error: lookup.error, data: group.rows[0] });
          continue;
        }
        const bill = lookup.bill;

        const lineCommands = [];
        const lineErrors = [];
        for (const row of group.rows) {
          const productText = productM ? cleanText(String(row[productM.header] ?? "")) : "";
          if (!productText) continue;

          const lineVals = { quantity: 1, price_unit: 0 };
          let prodIds = await odooExecute("product.product", "search", [[["name", "=", productText]]], { limit: 1 });
          if (!prodIds.length) prodIds = await odooExecute("product.product", "search", [[["name", "ilike", productText]]], { limit: 1 });
          if (prodIds.length) {
            lineVals.product_id = prodIds[0];
            lineVals.name = productText;
          } else {
            const acct = await resolveOrCreateExpenseAccount(productText, notes);
            if (acct.id === false) { lineErrors.push(`"${productText}": ${acct.note || "could not resolve as a product or account"}`); continue; }
            lineVals.account_id = acct.id;
            lineVals.name = productText;
          }
          if (descM && row[descM.header]) lineVals.name = cleanText(String(row[descM.header]));
          if (qtyM && row[qtyM.header] !== "" && row[qtyM.header] != null) lineVals.quantity = cleanNumber(row[qtyM.header]);
          if (priceM && row[priceM.header] !== "" && row[priceM.header] != null) lineVals.price_unit = cleanNumber(row[priceM.header]);
          if (taxM && row[taxM.header]) {
            const taxCmd = await resolveGstTaxIds(row[taxM.header], notes, side === "sales" ? "sale" : "purchase");
            if (taxCmd) lineVals.tax_ids = taxCmd;
          }
          lineCommands.push([0, 0, lineVals]);
        }

        if (!lineCommands.length) {
          errors.push({ row: group.firstRow, error: lineErrors.length ? `No usable line items for ${recordLabel} "${bill.name}": ${lineErrors.join("; ")}` : `No usable line items found for ${recordLabel} "${bill.name}".`, data: group.rows[0] });
          continue;
        }

        const originalState = bill.state;
        await resetBillToDraft(bill.id, originalState, notes);
        await withRetry(() => odooExecute("account.move", "write", [[bill.id], { invoice_line_ids: lineCommands }]));

        successCount++;
        notes.push(`${recordLabel} "${bill.name}": added ${lineCommands.length} extra line item(s)${originalState !== "draft" ? " (reset from Posted to Draft — left in Draft for you to review and post)" : ""}.${lineErrors.length ? ` Some lines were skipped: ${lineErrors.join("; ")}` : ""}`);
      } catch (e) {
        errors.push({ row: group.firstRow, error: e.message, data: group.rows[0] });
      }
      emit({ status: "uploading", progress: g + 1, total: groups.length });
    }

    return emit({ status: "done", result: { total: groups.length, created_count: successCount, duplicate_count: 0, failed_count: errors.length, errors, notes } });
  }

  // -------------------------------------------------------------------------
  // Vendor Bills / Sales Invoices / Credit Notes (Debtors) / Debit Notes
  // (Creditors) — dedicated processors (Batch 2). These four modules were
  // special-cased inline inside the extension's single giant
  // uploadSheetToOdoo (popup.js ~2841-3964). Ported here as their own
  // functions instead (per HANDOFF.md's recommendation (b)) so the existing
  // generic create/dedupe/write loop below — still used by contacts,
  // inventory, accounting, sales, purchaseOrders, openingBalance, leads,
  // physicalInventory — doesn't need any new per-field branches bolted on.
  // Shared helpers below (buildInvoiceLikeGroup, groupByBlankContinuation,
  // applyOrderLink, applyReversalLink, findDuplicateAndUpsert,
  // applyVendorBillTdsAndExtras, applyMoveTdsOnly) are reused by all four.
  // -------------------------------------------------------------------------

  // Same "blank continuation row" grouping the generic loop already uses for
  // accounting/sales/purchaseOrders (no isGroupKey on any of these four
  // schemas) — a row where every non-line mapped field is blank is treated
  // as an extra line of the record above it.
  function groupByBlankContinuation(usableRows, mappedFields) {
    // TDS/TCS Section/Base/Rate and Extra Line Item columns are deliberately
    // excluded here — they apply to the whole bill/invoice, not a specific
    // line, but it's entirely normal for a sheet to repeat the same
    // TDS/TCS Section value on every product row of a multi-line record
    // (rather than only the first) — that must still count as a
    // continuation row, not the start of a new record.
    const headerFieldDefs = mappedFields.filter((m) => !m.field.isLine && !m.field.isVendorLine
      && !m.field.isTdsSectionField && !m.field.isTdsBaseField && !m.field.isTdsRateField
      && !m.field.isTcsSectionField && !m.field.isTcsRateField && !m.field.isExtraLineField);
    const rowIsContinuation = (row) => headerFieldDefs.length > 0 && headerFieldDefs.every((m) => row[m.header] === "" || row[m.header] == null);
    const groups = [];
    for (const row of usableRows) {
      if (!groups.length || !rowIsContinuation(row)) groups.push([row]);
      else groups[groups.length - 1].push(row);
    }
    return groups;
  }

  // Builds `values`/`linesByField` for one group of rows (a Vendor Bill,
  // Sales Invoice, or Credit/Debit Note and its continuation lines), plus
  // whatever these four modules' own special columns (PO/SO Reference+Name,
  // Reversed Entry/Origin, TDS Section/Base/Rate, Extra Line Item) resolved
  // to for the group — mirrors the per-row field loop inside the
  // extension's giant uploadSheetToOdoo (popup.js ~3047-3305), minus the
  // Cash Book/Expense JV-only branches (Batch 3, not reachable for these
  // four schemas' field lists).
  async function buildInvoiceLikeGroup(group, mappedFields, schema, notes) {
    const values = { ...(schema.defaults || {}) };
    const linesByField = {};
    let pendingPoRef = null, pendingPoName = null;
    let pendingTdsSection = null, pendingTdsBase = null, pendingTdsRate = null;
    let pendingTcsSection = null, pendingTcsRate = null;
    let pendingExtraLine = null;
    let pendingOriginRef = null;

    for (const row of group) {
      const lineAccumulator = {};
      for (const m of mappedFields) {
        const raw = row[m.header];
        if (raw === "" || raw == null) continue;
        const field = m.field;

        if (field.isIgnored) continue;
        if (field.isPoReferenceField || field.isSoReferenceField) { pendingPoRef = cleanText(String(raw)); continue; }
        if (field.isPoNameField || field.isSoNameField) { pendingPoName = cleanText(String(raw)); continue; }
        if (field.isReversalRefField) { pendingOriginRef = cleanText(String(raw)); continue; }
        if (field.isTdsSectionField) { pendingTdsSection = cleanText(String(raw)); continue; }
        if (field.isTdsBaseField) { pendingTdsBase = cleanNumber(raw); continue; }
        if (field.isTdsRateField) { pendingTdsRate = cleanNumber(raw); continue; }
        if (field.isTcsSectionField) { pendingTcsSection = cleanText(String(raw)); continue; }
        if (field.isTcsRateField) { pendingTcsRate = cleanNumber(raw); continue; }
        if (field.isExtraLineField) {
          pendingExtraLine = pendingExtraLine || {};
          pendingExtraLine[field.isExtraLineField] = field.isExtraLineField === "quantity" || field.isExtraLineField === "price" ? cleanNumber(raw) : cleanText(String(raw));
          continue;
        }

        // Line / Product or Expense (or Income) Account — tries a real
        // product first, falls back to the Chart of Accounts (see
        // resolveOrCreateExpenseAccount), same as Bill/Invoice Additional
        // Lines' own productM handling above.
        if (field.isProductOrAccountField) {
          const text = cleanText(String(raw));
          const [lineField, subField] = field.name.split("/");
          lineAccumulator[lineField] = lineAccumulator[lineField] || {};
          if (!text) continue;
          let prodIds = await odooExecute("product.product", "search", [[["name", "=", text]]], { limit: 1 });
          if (!prodIds.length) prodIds = await odooExecute("product.product", "search", [[["name", "ilike", text]]], { limit: 1 });
          if (prodIds.length) {
            lineAccumulator[lineField][subField] = prodIds[0];
          } else {
            const acct = await resolveOrCreateExpenseAccount(text, notes);
            if (acct.id !== false) {
              lineAccumulator[lineField].account_id = acct.id;
              lineAccumulator[lineField].name = text;
            } else {
              notes.push(`Could not resolve "${text}" as a product or an account — that line was skipped.`);
            }
          }
          continue;
        }

        if (field.relation) {
          const resolved = await resolveRelation(field, raw);
          if (resolved.note) notes.push(resolved.note);
          if (resolved.id === false) continue;
          if (field.isLine) {
            const [lineField, subField] = field.name.split("/");
            lineAccumulator[lineField] = lineAccumulator[lineField] || {};
            lineAccumulator[lineField][subField] = resolved.id;
          } else {
            values[field.name] = resolved.id;
          }
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

        let val = raw;
        if (field.number) val = cleanNumber(val);
        else if (field.date) val = cleanDate(val);
        else if (typeof val === "string") val = cleanText(val);
        if (field.isLine) {
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

    return { values, linesByField, pendingPoRef, pendingPoName, pendingTdsSection, pendingTdsBase, pendingTdsRate, pendingTcsSection, pendingTcsRate, pendingExtraLine, pendingOriginRef };
  }

  // Vendor Bills / Sales Invoices — links the bill/invoice back to its
  // Purchase/Sales Order (matched on partner_ref/client_order_ref, with an
  // optional Name column to disambiguate), pulling Vendor/Customer and any
  // blank Line / Unit Price / Taxes from the matched order's own lines.
  // Ported from popup.js ~3641-3761 ("Vendor Bills / Sales Invoices modules"
  // comment block) — logic unchanged, just parameterized by isVendorBill
  // instead of branching on values.move_type.
  async function applyOrderLink(values, isVendorBill, pendingPoRef, pendingPoName, hasOrderRefColumnMapped, notes) {
    const orderModel = isVendorBill ? "purchase.order" : "sale.order";
    const orderLineModel = isVendorBill ? "purchase.order.line" : "sale.order.line";
    const refField = isVendorBill ? "partner_ref" : "client_order_ref";
    const orderNoun = isVendorBill ? "Purchase Order" : "Sales Order";
    const refNoun = isVendorBill ? "Vendor Reference" : "Order Reference";
    const refColumnNoun = isVendorBill ? "PO Reference" : "SO Reference";
    const nameColumnNoun = isVendorBill ? "PO Name" : "SO Name";
    const ownRefNoun = isVendorBill ? "Bill Reference" : "Invoice Reference";
    const exampleName = isVendorBill ? "P00088" : "S00088";

    let refSource = refColumnNoun;
    let poRef = pendingPoRef;
    if (!poRef && values.ref) { poRef = values.ref; refSource = ownRefNoun; }

    if (poRef !== null && poRef !== undefined && poRef !== "") {
      const domain = [[refField, "=", poRef]];
      if (pendingPoName) domain.push(["name", "=", pendingPoName]);
      let orderIds = await odooExecute(orderModel, "search", [domain]);
      if (!orderIds.length && !pendingPoName) {
        orderIds = await odooExecute(orderModel, "search", [[[refField, "ilike", poRef]]]);
      }
      if (!orderIds.length) {
        if (!hasOrderRefColumnMapped) {
          notes.push(`No ${orderNoun} found matching "${poRef}" — created as a standalone ${isVendorBill ? "bill" : "invoice"} with no order link (no ${refColumnNoun} column was mapped for this sheet, so this wasn't necessarily expected to match).`);
        } else {
          throw new Error(`No ${orderNoun} found with ${refNoun} "${poRef}"${pendingPoName ? ` and ${nameColumnNoun} "${pendingPoName}"` : ""} (matched via ${refSource}) — check it matches that order's ${refNoun} field exactly.${refSource === ownRefNoun ? ` No ${refColumnNoun} column was mapped for this row, so the ${ownRefNoun} was used to look it up instead — add a ${refColumnNoun} column if that number doesn't also appear as the ${refNoun} on your orders.` : ""}`);
        }
      } else if (orderIds.length > 1) {
        if (!hasOrderRefColumnMapped) {
          notes.push(`"${poRef}" matched ${orderIds.length} different ${orderNoun.toLowerCase()}s — left this row unlinked to any order rather than guessing (no ${refColumnNoun} column was mapped for this sheet, so set the link manually in Odoo if it should have one).`);
        } else {
          throw new Error(`${orderIds.length} different ${orderNoun}s share the ${refNoun} "${poRef}" (matched via ${refSource}) — map a "${nameColumnNoun}" column with each row's exact Odoo order number (e.g. ${exampleName}, shown on the order itself) so this row can tell them apart.`);
        }
      } else {
        const [order] = await odooExecute(orderModel, "read", [orderIds, ["name", "partner_id", "order_line"]]);
        values.partner_id = order.partner_id[0];
        values.invoice_origin = order.name;
        if (refSource === ownRefNoun) {
          notes.push(`Row matched to ${order.name} using the ${ownRefNoun} "${poRef}" — no ${refColumnNoun} column was mapped for this row.`);
        }
        const lineField = "invoice_line_ids";
        if (order.order_line && order.order_line.length && values[lineField]) {
          const orderLines = await odooExecute(orderLineModel, "read", [order.order_line, ["product_id", "price_unit", "tax_ids"]]);
          const orderLineByProduct = {};
          for (const ol of orderLines) if (ol.product_id) orderLineByProduct[ol.product_id[0]] = ol;
          for (const cmd of values[lineField]) {
            const l = cmd[2];
            const orderLine = l.product_id != null ? orderLineByProduct[l.product_id] : null;
            if (!orderLine) continue;
            if (isVendorBill) { if (l.purchase_line_id == null) l.purchase_line_id = orderLine.id; }
            else { if (l.sale_line_ids == null) l.sale_line_ids = [[6, 0, [orderLine.id]]]; }
            if (l.price_unit == null) l.price_unit = orderLine.price_unit;
            if (l.tax_ids == null && orderLine.tax_ids && orderLine.tax_ids.length) l.tax_ids = [[6, 0, orderLine.tax_ids]];
          }
        }
      }
    } else if (hasOrderRefColumnMapped) {
      throw new Error(`Could not find a ${orderNoun} for this ${isVendorBill ? "bill" : "invoice"} — map a ${refColumnNoun} column, or a ${ownRefNoun} that also matches an order's ${refNoun}, on the first line.`);
    }
    // else: no PO/SO Reference column mapped and no Bill/Invoice Reference
    // to fall back to either — proceed as a standalone bill/invoice using
    // the explicitly mapped Vendor/Customer.
  }

  // Credit Notes (Debtors) / Debit Notes (Creditors) — links the note back
  // to the invoice/bill it reverses, when Reversed Entry / Origin was
  // mapped and matches (checked against both the origin's own Odoo "name"
  // and its "ref" field). Optional — unlike applyOrderLink above, a blank
  // or unmatched reference is not an error, just a standalone note. Ported
  // from popup.js ~3763-3811.
  async function applyReversalLink(values, isCustomerNote, pendingOriginRef, notes) {
    const originMoveType = isCustomerNote ? "out_invoice" : "in_invoice";
    const originNoun = isCustomerNote ? "Invoice" : "Bill";
    if (!pendingOriginRef) return;
    const domain = [["move_type", "=", originMoveType], "|", ["name", "=", pendingOriginRef], ["ref", "=", pendingOriginRef]];
    if (values.partner_id) domain.push(["partner_id", "=", values.partner_id]);
    const originIds = await odooExecute("account.move", "search", [domain]);
    if (originIds.length === 1) {
      values.reversed_entry_id = originIds[0];
      const [origin] = await odooExecute("account.move", "read", [originIds, ["invoice_line_ids"]]);
      if (origin.invoice_line_ids && origin.invoice_line_ids.length && values.invoice_line_ids) {
        const originLines = await odooExecute("account.move.line", "read", [origin.invoice_line_ids, ["product_id", "price_unit", "tax_ids"]]);
        const originLineByProduct = {};
        for (const ol of originLines) if (ol.product_id) originLineByProduct[ol.product_id[0]] = ol;
        for (const cmd of values.invoice_line_ids) {
          const l = cmd[2];
          const originLine = l.product_id != null ? originLineByProduct[l.product_id] : null;
          if (!originLine) continue;
          if (l.price_unit == null) l.price_unit = originLine.price_unit;
          if (l.tax_ids == null && originLine.tax_ids && originLine.tax_ids.length) l.tax_ids = [[6, 0, originLine.tax_ids]];
        }
      }
    } else if (originIds.length > 1) {
      notes.push(`"${pendingOriginRef}" matched ${originIds.length} different ${originNoun.toLowerCase()}s for this party — left the Reversed Entry link blank; set it manually on the note in Odoo.`);
    } else {
      notes.push(`Could not find ${originNoun} "${pendingOriginRef}" for this party — created as a standalone ${isCustomerNote ? "credit" : "debit"} note with the Reversed Entry link left blank.`);
    }
  }

  // Vendor Bills module: applies the optional TDS-at-import and Extra Line
  // Item columns to a bill right after it's created (or updated, if this
  // was a duplicate row) — same logic as the standalone TDS Entry / Bill
  // Additional Lines modules, just triggered inline instead of as a
  // separate later import. No-op if neither set of columns was actually
  // filled in for this row. Ported from popup.js ~2548-2610.
  async function applyVendorBillTdsAndExtras(billId, tds, extraLine, notes, errors, rowIndex) {
    const hasTds = !!(tds && tds.section);
    const hasExtraLine = !!(extraLine && (extraLine.product || extraLine.price));
    if (!hasTds && !hasExtraLine) return;

    try {
      const [bill] = await odooExecute("account.move", "read", [[billId], ["name", "state", "amount_untaxed"]]);
      const lineCommands = [];

      if (hasExtraLine) {
        const productText = cleanText(String(extraLine.product || ""));
        const lineVals = { quantity: extraLine.quantity || 1, price_unit: extraLine.price || 0 };
        if (productText) {
          let prodIds = await odooExecute("product.product", "search", [[["name", "=", productText]]], { limit: 1 });
          if (!prodIds.length) prodIds = await odooExecute("product.product", "search", [[["name", "ilike", productText]]], { limit: 1 });
          if (prodIds.length) {
            lineVals.product_id = prodIds[0];
            lineVals.name = productText;
          } else {
            const acct = await resolveOrCreateExpenseAccount(productText, notes);
            if (acct.id !== false) { lineVals.account_id = acct.id; lineVals.name = productText; }
            else notes.push(`Bill "${bill.name}": extra line item "${productText}" could not be resolved as a product or account — skipped.`);
          }
        }
        if (extraLine.desc) lineVals.name = extraLine.desc;
        if (extraLine.tax) {
          const taxCmd = await resolveGstTaxIds(extraLine.tax, notes, "purchase");
          if (taxCmd) lineVals.tax_ids = taxCmd;
        }
        if (lineVals.product_id || lineVals.account_id) lineCommands.push([0, 0, lineVals]);
      }

      if (lineCommands.length) {
        if (bill.state !== "draft") await resetBillToDraft(billId, bill.state, notes);
        await withRetry(() => odooExecute("account.move", "write", [[billId], { invoice_line_ids: lineCommands }]));
      }

      if (!hasTds && !lineCommands.length) return;

      if (!hasTds) {
        notes.push(`Bill "${bill.name}": extra line item added — left in Draft for you to review and post.`);
        return;
      }

      const [freshBill] = await odooExecute("account.move", "read", [[billId], ["name", "state", "amount_untaxed"]]);
      if (freshBill.state !== "posted") {
        await withRetry(() => odooExecute("account.move", "action_post", [[billId]]));
      }

      const [postedBill] = await odooExecute("account.move", "read", [[billId], ["name", "amount_untaxed"]]);
      const result = await applyNativeTdsWithhold(billId, postedBill.name, postedBill.amount_untaxed, tds, "purchase", notes);
      if (result) {
        notes.push(`Bill "${bill.name}": TDS Entry posted via "${result.taxName}" — ${result.amount}${lineCommands.length ? " (extra line item also added)" : ""}.`);
      }
    } catch (e) {
      errors.push({ row: rowIndex, error: `TDS/extra-line step failed for bill #${billId}: ${e.message}`, data: {} });
    }
  }

  // Sales Invoices module: applies the optional TDS-at-import, Extra Line
  // Item, and TCS-at-import columns to a Customer Invoice right after it's
  // created (or updated). TCS is applied first (it's a real tax and needs
  // the invoice to still be editable); TDS posts the invoice and runs
  // after, same ordering as the extension's applyMoveTdsAndExtras.
  async function applyMoveTdsOnly(moveId, tds, side, notes, errors, rowIndex, extraLine, tcs) {
    const hasTds = !!(tds && tds.section);
    const hasExtraLine = !!(extraLine && (extraLine.product || extraLine.price));
    const hasTcs = side === "sales" && !!(tcs && tcs.section);
    if (!hasTds && !hasExtraLine && !hasTcs) return;

    try {
      const [move] = await odooExecute("account.move", "read", [[moveId], ["name", "state", "amount_untaxed"]]);
      const lineCommands = [];

      if (hasExtraLine) {
        const productText = cleanText(String(extraLine.product || ""));
        const lineVals = { quantity: extraLine.quantity || 1, price_unit: extraLine.price || 0 };
        if (productText) {
          let prodIds = await odooExecute("product.product", "search", [[["name", "=", productText]]], { limit: 1 });
          if (!prodIds.length) prodIds = await odooExecute("product.product", "search", [[["name", "ilike", productText]]], { limit: 1 });
          if (prodIds.length) {
            lineVals.product_id = prodIds[0];
            lineVals.name = productText;
          } else {
            const acct = await resolveOrCreateExpenseAccount(productText, notes);
            if (acct.id !== false) { lineVals.account_id = acct.id; lineVals.name = productText; }
            else notes.push(`Invoice "${move.name}": extra line item "${productText}" could not be resolved as a product or account — skipped.`);
          }
        }
        if (extraLine.desc) lineVals.name = extraLine.desc;
        if (extraLine.tax) {
          const taxCmd = await resolveGstTaxIds(extraLine.tax, notes, "sale");
          if (taxCmd) lineVals.tax_ids = taxCmd;
        }
        if (lineVals.product_id || lineVals.account_id) lineCommands.push([0, 0, lineVals]);
      }

      if (lineCommands.length) {
        if (move.state !== "draft") await resetBillToDraft(moveId, move.state, notes);
        await withRetry(() => odooExecute("account.move", "write", [[moveId], { invoice_line_ids: lineCommands }]));
      }

      // TCS is a real tax on the invoice itself, so it needs to be applied
      // while still editable — before any TDS step below posts it.
      let tcsApplied = null;
      if (hasTcs) tcsApplied = await applyTcsToInvoiceLines(moveId, move.name, tcs, notes);

      if (!hasTds && !lineCommands.length && !hasTcs) return;

      if (!hasTds) {
        const parts = [];
        if (lineCommands.length) parts.push("extra line item added");
        if (tcsApplied) parts.push(`TCS "${tcsApplied}" added`);
        if (parts.length) notes.push(`Invoice "${move.name}": ${parts.join(", ")} — left in Draft for you to review and post.`);
        return;
      }

      const [fresh] = await odooExecute("account.move", "read", [[moveId], ["name", "state", "amount_untaxed"]]);
      if (fresh.state !== "posted") {
        await withRetry(() => odooExecute("account.move", "action_post", [[moveId]]));
      }
      const [posted] = await odooExecute("account.move", "read", [[moveId], ["name", "amount_untaxed"]]);
      const result = await applyNativeTdsWithhold(moveId, posted.name, posted.amount_untaxed, tds, side, notes);
      if (result) {
        const extras = [lineCommands.length && "extra line item also added", tcsApplied && `TCS "${tcsApplied}" also added`].filter(Boolean).join(", ");
        notes.push(`"${move.name}": TDS Entry posted via "${result.taxName}" — ${result.amount}${extras ? ` (${extras})` : ""}.`);
      }
    } catch (e) {
      errors.push({ row: rowIndex, error: `TDS/TCS/extra-line step failed for move #${moveId}: ${e.message}`, data: {} });
    }
  }

  // Shared dedupe-then-create-or-update step, same behavior as the generic
  // loop's own duplicate handling (popup.js ~3842-3908 / the existing
  // generic uploadSheetToOdoo below) — "smart" always resolves to "update"
  // for account.move (no transaction-lock check like product.template has).
  // Returns { id, isDuplicate } so the caller can apply Vendor
  // Bills/Sales Invoices' inline TDS/extras against whichever id (new or
  // existing) ends up holding the row's data.
  async function findDuplicateAndUpsert(schema, values, linesByField, skipDuplicates, duplicateAction, notes, errors, groupIndex) {
    if (skipDuplicates) {
      const dup = await findDuplicate(schema.model, schema.dedupeKeys, values);
      if (dup) {
        const effectiveAction = duplicateAction === "smart" ? "update" : duplicateAction;
        if (effectiveAction === "update") {
          try {
            const writeValues = { ...values };
            for (const lineField of Object.keys(linesByField)) writeValues[lineField] = [[5, 0, 0], ...values[lineField]];
            await withRetry(() => odooExecute(schema.model, "write", [[dup.id], writeValues]));
            notes.push(`Row ${groupIndex + 1}: updated existing record #${dup.id} (matched on ${dup.matchedOn.join(" + ")}).`);
          } catch (e) {
            errors.push({ row: groupIndex, error: `Update failed for existing record #${dup.id}: ${e.message}`, data: values });
          }
        } else {
          notes.push(`Row ${groupIndex + 1}: skipped — already exists as record #${dup.id} (matched on ${dup.matchedOn.join(" + ")}).`);
        }
        return { id: dup.id, isDuplicate: true };
      }
    }
    const newId = await withRetry(() => odooExecute(schema.model, "create", [values]));
    return { id: newId, isDuplicate: false };
  }

  // Vendor Bills — dedicated processor (Batch 2). Follows the same
  // emit()-based progress pattern as processTdsEntrySheet/
  // processAdditionalLinesSheet above.
  async function processVendorBillsSheet(sheet, onProgress) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const schema = ODOO_SCHEMAS.vendorBills;
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const hasOrderRefColumnMapped = mappedFields.some((m) => m.field.isPoReferenceField);
    const skipDuplicates = sheet.skipDuplicates !== false;
    const duplicateAction = sheet.duplicateAction || "smart";

    const usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));
    const groups = groupByBlankContinuation(usableRows, mappedFields);

    emit({ status: "uploading", progress: 0, total: groups.length });
    const createdIds = []; const errors = []; const notes = []; let duplicateCount = 0;

    for (let g = 0; g < groups.length; g++) {
      try {
        const { values, linesByField, pendingPoRef, pendingPoName, pendingTdsSection, pendingTdsBase, pendingTdsRate, pendingExtraLine } =
          await buildInvoiceLikeGroup(groups[g], mappedFields, schema, notes);

        await applyOrderLink(values, true, pendingPoRef, pendingPoName, hasOrderRefColumnMapped, notes);

        const upsert = await findDuplicateAndUpsert(schema, values, linesByField, skipDuplicates, duplicateAction, notes, errors, g);
        if (upsert.isDuplicate) duplicateCount++; else createdIds.push(upsert.id);

        await applyVendorBillTdsAndExtras(upsert.id, { section: pendingTdsSection, base: pendingTdsBase, rate: pendingTdsRate }, pendingExtraLine, notes, errors, g);
      } catch (e) {
        errors.push({ row: g, error: e.message, data: groups[g][0] });
      }
      emit({ status: "uploading", progress: g + 1, total: groups.length });
    }

    return emit({
      status: "done",
      result: { total: groups.length, created_count: createdIds.length, duplicate_count: duplicateCount, duplicate_action: duplicateAction, failed_count: errors.length, errors, notes: [...new Set(notes)] },
    });
  }

  // Sales Invoices — dedicated processor (Batch 2).
  async function processSalesInvoicesSheet(sheet, onProgress) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const schema = ODOO_SCHEMAS.salesInvoices;
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const hasOrderRefColumnMapped = mappedFields.some((m) => m.field.isSoReferenceField);
    const skipDuplicates = sheet.skipDuplicates !== false;
    const duplicateAction = sheet.duplicateAction || "smart";

    const usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));
    const groups = groupByBlankContinuation(usableRows, mappedFields);

    emit({ status: "uploading", progress: 0, total: groups.length });
    const createdIds = []; const errors = []; const notes = []; let duplicateCount = 0;

    for (let g = 0; g < groups.length; g++) {
      try {
        const { values, linesByField, pendingPoRef, pendingPoName, pendingTdsSection, pendingTdsBase, pendingTdsRate, pendingTcsSection, pendingTcsRate, pendingExtraLine } =
          await buildInvoiceLikeGroup(groups[g], mappedFields, schema, notes);

        await applyOrderLink(values, false, pendingPoRef, pendingPoName, hasOrderRefColumnMapped, notes);

        const upsert = await findDuplicateAndUpsert(schema, values, linesByField, skipDuplicates, duplicateAction, notes, errors, g);
        if (upsert.isDuplicate) duplicateCount++; else createdIds.push(upsert.id);

        await applyMoveTdsOnly(upsert.id, { section: pendingTdsSection, base: pendingTdsBase, rate: pendingTdsRate }, "sales", notes, errors, g, pendingExtraLine, { section: pendingTcsSection, rate: pendingTcsRate });
      } catch (e) {
        errors.push({ row: g, error: e.message, data: groups[g][0] });
      }
      emit({ status: "uploading", progress: g + 1, total: groups.length });
    }

    return emit({
      status: "done",
      result: { total: groups.length, created_count: createdIds.length, duplicate_count: duplicateCount, duplicate_action: duplicateAction, failed_count: errors.length, errors, notes: [...new Set(notes)] },
    });
  }

  // Credit Notes (Debtors) / Debit Notes (Creditors) — shared processor,
  // side="sales" for Credit Notes (out_refund, against Customer Invoices)
  // and side="purchase" for Debit Notes (in_refund, against Vendor Bills).
  // No PO/SO linking (neither schema has those fields) — only the optional
  // Reversed Entry / Origin link.
  async function processCreditNoteSheet(sheet, onProgress, side) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const schema = ODOO_SCHEMAS[side === "sales" ? "creditNotesDebtors" : "debitNotesCreditors"];
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const skipDuplicates = sheet.skipDuplicates !== false;
    const duplicateAction = sheet.duplicateAction || "smart";

    const usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));
    const groups = groupByBlankContinuation(usableRows, mappedFields);

    emit({ status: "uploading", progress: 0, total: groups.length });
    const createdIds = []; const errors = []; const notes = []; let duplicateCount = 0;

    for (let g = 0; g < groups.length; g++) {
      try {
        const { values, linesByField, pendingOriginRef } = await buildInvoiceLikeGroup(groups[g], mappedFields, schema, notes);

        await applyReversalLink(values, side === "sales", pendingOriginRef, notes);

        const upsert = await findDuplicateAndUpsert(schema, values, linesByField, skipDuplicates, duplicateAction, notes, errors, g);
        if (upsert.isDuplicate) duplicateCount++; else createdIds.push(upsert.id);
      } catch (e) {
        errors.push({ row: g, error: e.message, data: groups[g][0] });
      }
      emit({ status: "uploading", progress: g + 1, total: groups.length });
    }

    return emit({
      status: "done",
      result: { total: groups.length, created_count: createdIds.length, duplicate_count: duplicateCount, duplicate_action: duplicateAction, failed_count: errors.length, errors, notes: [...new Set(notes)] },
    });
  }

  // Resolves a real control account (e.g. Payables/Receivables) by
  // account_type — NEVER by name/creation. This is what stops a
  // same-named-but-wrong-type account from ever being confused with the
  // real control account (the exact bug a same-named stray "SUNDRY
  // CREDITORS" expense account caused in production once).
  async function resolveControlAccount(accountType, label, notes) {
    if (controlAccountCache.has(accountType)) {
      const cached = controlAccountCache.get(accountType);
      if (cached.note) notes.push(cached.note);
      return cached;
    }
    let result;
    try {
      const ids = await odooExecute("account.account", "search", [[["account_type", "=", accountType]]], { limit: 1, order: "code" });
      if (ids.length) {
        result = { id: ids[0], created: false, note: null };
      } else {
        result = { id: false, created: false, note: `Could not find a ${label} control account in your Chart of Accounts — that line was skipped. Set one up in Accounting first.` };
      }
    } catch (e) {
      result = { id: false, created: false, note: `${label} control account lookup failed: ${e.message}` };
    }
    controlAccountCache.set(accountType, result);
    if (result.note) notes.push(result.note);
    return result;
  }

  async function resolveCashBookHeadAccount(name, notes) {
    const text = cleanText(String(name));
    if (SUNDRY_CREDITORS_RE.test(text)) return resolveControlAccount("liability_payable", "Payables/Creditors", notes);
    if (SUNDRY_DEBTORS_RE.test(text)) return resolveControlAccount("asset_receivable", "Receivables/Debtors", notes);
    return resolveOrCreateExpenseAccount(text, notes);
  }

  // The one Cash journal to post a Cash Book import against — found once
  // per upload. Deliberately does NOT auto-create one: which account backs
  // a Cash journal is a real accounting decision the user should set up
  // once themselves.
  async function findCashJournal() {
    const ids = await odooExecute("account.journal", "search", [[["type", "=", "cash"]]], { order: "id" });
    if (!ids.length) return { id: false, error: `No Cash journal found in your Chart of Accounts. Go to Accounting -> Configuration -> Journals and create one (Type: Cash) first.` };
    if (ids.length > 1) return { id: ids[0], note: `Found ${ids.length} Cash journals — used the first one (id ${ids[0]}). If that's the wrong one, post these entries to the right journal manually in Odoo.` };
    return { id: ids[0] };
  }

  // All distinct subsets of `arr` with exactly `size` elements — used to
  // search for a combination of open bills that sums to a Cash Book row's
  // amount. `arr` is capped at 12 candidates by the caller.
  function combinationsOfSize(arr, size) {
    const results = [];
    function helper(start, combo) {
      if (combo.length === size) { results.push(combo.slice()); return; }
      for (let i = start; i < arr.length; i++) {
        combo.push(arr[i]);
        helper(i + 1, combo);
        combo.pop();
      }
    }
    helper(0, []);
    return results;
  }

  // Cash Book: for a "SUNDRY CREDITORS"/"SUNDRY DEBTORS" row that names a
  // Payee, try to find the ONE open bill/invoice (or one unambiguous 2-3
  // combination) on that partner summing to exactly this amount, and
  // register a real payment against it — instead of just posting a plain
  // line to the control account. Only auto-registers when the match is
  // unambiguous; anything less certain falls back to plain posting.
  async function tryCashBookAutoPayment(partnerId, amount, isReceived, cashBookJournalId, notes) {
    const moveType = isReceived ? "out_invoice" : "in_invoice";
    const noun = isReceived ? "customer invoice" : "vendor bill";
    try {
      const candidates = await odooExecute("account.move", "search_read", [[
        ["partner_id", "=", partnerId],
        ["move_type", "=", moveType],
        ["state", "=", "posted"],
        ["payment_state", "not in", ["paid", "reversed", "in_payment"]],
      ]], { fields: ["id", "name", "amount_residual"] });
      const isMatch = (sum) => Math.abs(sum - amount) < 0.01;

      let matches = candidates.filter((c) => isMatch(c.amount_residual));

      if (!matches.length && candidates.length >= 2 && candidates.length <= 12) {
        const foundCombos = [];
        for (let size = 2; size <= 3 && size <= candidates.length && foundCombos.length <= 1; size++) {
          for (const combo of combinationsOfSize(candidates, size)) {
            if (isMatch(combo.reduce((s, c) => s + c.amount_residual, 0))) foundCombos.push(combo);
            if (foundCombos.length > 1) break;
          }
        }
        if (foundCombos.length === 1) matches = foundCombos[0];
        else if (foundCombos.length > 1) return { registered: false, note: `More than one combination of open ${noun}s for this partner sums to ₹${amount} — couldn't tell which set was meant, so this was posted to the control account instead; register the right combined payment manually in Odoo.` };
      }

      if (!matches.length) return { registered: false, note: `No open ${noun}(s) found for this partner summing to ₹${amount} exactly — posted to the control account instead; reconcile manually against the right bill/invoice in Odoo once you find it (its amount may not match this row exactly, e.g. a partial payment).` };

      const matchIds = matches.map((c) => c.id);
      const label = matches.map((c) => c.name).join(", ");
      const action = await odooExecute("account.move", "action_register_payment", [matchIds]);
      const wizIds = await odooExecute("account.payment.register", "create", [[{}]], { context: action.context });
      const wizId = Array.isArray(wizIds) ? wizIds[0] : wizIds;
      await odooExecute("account.payment.register", "write", [[wizId], { journal_id: cashBookJournalId, group_payment: matchIds.length > 1 }]);
      await odooExecute("account.payment.register", "action_create_payments", [[wizId]], { context: action.context });
      return { registered: true, note: matchIds.length > 1 ? `Registered one combined payment of ₹${amount} across ${matchIds.length} open ${noun}s (${label}) — reconciled automatically, no manual step needed in Odoo.` : `Registered a payment of ₹${amount} against ${label} — reconciled automatically, no manual step needed in Odoo.` };
    } catch (e) {
      return { registered: false, note: `Couldn't auto-register a payment for this row (${e.message}) — posted to the control account instead; reconcile manually in Odoo.` };
    }
  }

  // Cash Book processor — one two-line journal entry per voucher (Head
  // account vs. the Cash journal's own account), built from the
  // blank-continuation "__group" (Voucher Reference) convention.
  async function processCashBookSheet(sheet, onProgress) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const groupM = mappedFields.find((m) => m.field.isGroupKey);
    const dateM = mappedFields.find((m) => m.field.name === "date");
    const typeM = mappedFields.find((m) => m.field.isCashTypeField);
    const headM = mappedFields.find((m) => m.field.isCashHeadField);
    const payeeM = mappedFields.find((m) => m.field.isCashPayeeField);
    const amountM = mappedFields.find((m) => m.field.isCashAmountField);
    const narrationM = mappedFields.find((m) => m.field.name === "narration");

    if (!groupM || !typeM || !headM || !amountM) {
      return emit({ status: "done", result: { total: 0, created_count: 0, duplicate_count: 0, failed_count: 1, errors: [{ row: 0, error: "Map Voucher Reference, Type, Head (Account), and Amount before importing.", data: {} }], notes: [] } });
    }

    const cashJournal = await findCashJournal();
    if (cashJournal.id === false) {
      return emit({ status: "done", result: { total: 0, created_count: 0, duplicate_count: 0, failed_count: 1, errors: [{ row: 0, error: cashJournal.error, data: {} }], notes: [] } });
    }
    const cashBookJournalId = cashJournal.id;
    const notes = [];
    if (cashJournal.note) notes.push(cashJournal.note);

    const usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));
    const groups = [];
    let lastKey;
    for (const row of usableRows) {
      const key = row[groupM.header];
      const isBlankKey = key === "" || key == null;
      if (!groups.length || (!isBlankKey && key !== lastKey)) groups.push([row]);
      else groups[groups.length - 1].push(row);
      if (!isBlankKey) lastKey = key;
    }

    emit({ status: "uploading", progress: 0, total: groups.length });

    const errors = [];
    let successCount = 0;
    let duplicateCount = 0;

    for (let g = 0; g < groups.length; g++) {
      const group = groups[g];
      const voucherRef = cleanText(String(group[0][groupM.header] ?? ""));
      try {
        if (!voucherRef) {
          errors.push({ row: g, error: "Voucher Reference is required.", data: group[0] });
          continue;
        }

        // Same dedupe convention as the rest of the app: skip if a move
        // already exists in this journal with this ref.
        const existing = await odooExecute("account.move", "search", [[["journal_id", "=", cashBookJournalId], ["ref", "=", voucherRef]]], { limit: 1 });
        if (existing.length) { duplicateCount++; continue; }

        const date = dateM ? cleanDate(group[0][dateM.header]) : null;
        const narrationParts = [];
        const lineCommands = [];
        let cashDelta = 0;
        let rowsTotal = 0, rowsDiverted = 0;

        for (const row of group) {
          const typeRaw = typeM ? cleanText(String(row[typeM.header] ?? "")).toUpperCase() : "";
          const headRaw = headM ? cleanText(String(row[headM.header] ?? "")) : "";
          const payeeRaw = payeeM ? cleanText(String(row[payeeM.header] ?? "")) : "";
          const amount = amountM && row[amountM.header] !== "" && row[amountM.header] != null ? cleanNumber(row[amountM.header]) : null;
          if (narrationM && row[narrationM.header]) narrationParts.push(cleanText(String(row[narrationM.header])));
          if (!headRaw && amount == null) continue;

          const isReceived = typeRaw === "CR" || /received/i.test(typeRaw);
          const isPaid = typeRaw === "CP" || /paid/i.test(typeRaw);
          if (!isReceived && !isPaid) notes.push(`Voucher "${voucherRef}", Head "${headRaw}": Type "${typeRaw}" wasn't recognized as CP/Cash Paid or CR/Cash Received — treated as CP (Cash Paid).`);

          const amt = amount || 0;
          rowsTotal++;
          const isControlAccountRow = SUNDRY_CREDITORS_RE.test(headRaw) || SUNDRY_DEBTORS_RE.test(headRaw);

          let diverted = false;
          if (isControlAccountRow && payeeRaw) {
            const payeeResolved = await resolveRelation({ relModel: "res.partner", autoCreate: true }, payeeRaw);
            if (payeeResolved.note) notes.push(payeeResolved.note);
            if (payeeResolved.id !== false) {
              const attempt = await tryCashBookAutoPayment(payeeResolved.id, amt, isReceived, cashBookJournalId, notes);
              if (attempt.note) notes.push(`Voucher "${voucherRef}", Head "${headRaw}", Payee "${payeeRaw}": ${attempt.note}`);
              if (attempt.registered) { diverted = true; rowsDiverted++; }
            }
          }

          if (!diverted) {
            const head = await resolveCashBookHeadAccount(headRaw, notes);
            if (head.id !== false) {
              const line = { account_id: head.id, name: payeeRaw || headRaw || "" };
              if (payeeRaw) {
                const payeeResolved = await resolveRelation({ relModel: "res.partner", autoCreate: true }, payeeRaw);
                if (payeeResolved.id !== false) line.partner_id = payeeResolved.id;
              } else if (isControlAccountRow) {
                notes.push(`Voucher "${voucherRef}", Head "${headRaw}": no Payee mapped — posted to the control account with no subledger partner, so you'll need to pick the partner yourself when reconciling this against the open bill/invoice in Odoo.`);
              }
              if (isReceived) { line.credit = amt; cashDelta -= amt; }
              else { line.debit = amt; cashDelta += amt; }
              lineCommands.push([0, 0, line]);
            }
          }
        }

        if (rowsTotal > 0 && rowsDiverted === rowsTotal) {
          // Every row in this voucher was fully handled by an
          // auto-registered payment — no plain journal entry needed at all.
          successCount++;
          notes.push(`Voucher "${voucherRef}": every row was settled via an auto-registered payment — no separate Journal Entry was needed.`);
          continue;
        }

        if (Math.round(cashDelta * 100) !== 0) {
          if (!uploadSheetToOdoo._cashAccountCache) uploadSheetToOdoo._cashAccountCache = new Map();
          let cashAccountId = uploadSheetToOdoo._cashAccountCache.get(cashBookJournalId);
          if (cashAccountId === undefined) {
            const [journal] = await odooExecute("account.journal", "read", [[cashBookJournalId], ["default_account_id"]]);
            cashAccountId = journal.default_account_id ? journal.default_account_id[0] : false;
            uploadSheetToOdoo._cashAccountCache.set(cashBookJournalId, cashAccountId);
          }
          if (cashAccountId) {
            const cashLine = { account_id: cashAccountId, name: voucherRef };
            if (cashDelta > 0) cashLine.credit = cashDelta; else cashLine.debit = -cashDelta;
            lineCommands.push([0, 0, cashLine]);
          } else {
            notes.push(`Voucher "${voucherRef}": the Cash journal has no account configured — the balancing Cash line was skipped, so this entry won't be balanced. Fix the journal's account in Odoo and re-import.`);
          }
        }

        if (!lineCommands.length) {
          errors.push({ row: g, error: `Voucher "${voucherRef}": no usable lines resolved — nothing was created.`, data: group[0] });
          continue;
        }

        const moveVals = { move_type: "entry", journal_id: cashBookJournalId, ref: voucherRef, date, narration: narrationParts.join(" ") || undefined, line_ids: lineCommands };
        const newId = await withRetry(() => odooExecute("account.move", "create", [moveVals]));
        successCount++;
        notes.push(`Voucher "${voucherRef}": created draft Journal Entry #${newId} — ${lineCommands.length} line(s).`);
      } catch (e) {
        errors.push({ row: g, error: e.message, data: group[0] });
      }
      emit({ status: "uploading", progress: g + 1, total: groups.length });
    }

    return emit({ status: "done", result: { total: groups.length, created_count: successCount, duplicate_count: duplicateCount, failed_count: errors.length, errors, notes } });
  }

  // -------------------------------------------------------------------------
  // Expense JV helpers — see the expenseJv schema's own `note` for the full
  // background and the 10 hard rules. All account resolution goes through
  // an explicit mapping table persisted in the connected Odoo database
  // itself (an ir.config_parameter, NOT the web app's own Supabase DB), so
  // it travels with the database rather than with any one user/session.
  // -------------------------------------------------------------------------

  // A Particulars cell like "By Rent Office Premises" or "To TDS PAYABLE"
  // carries a "By"/"To" prefix that's the only reliable Dr/Cr signal when
  // the row's own Debit/Credit cells are blank — always stripped before any
  // account lookup, since it's never itself part of the account name.
  function parseByToPrefix(particulars) {
    const text = String(particulars || "").trim();
    const m = text.match(/^(by|to)\b[\s:.\-]*(.*)$/i);
    if (m && m[2].trim()) return { cleaned: m[2].trim(), side: m[1].toLowerCase() === "by" ? "debit" : "credit" };
    return { cleaned: text, side: null };
  }

  async function loadExpenseJvMapping(notes) {
    if (expenseJvMappingCache) return expenseJvMappingCache;
    try {
      const rows = await odooExecute("ir.config_parameter", "search_read", [[["key", "=", EXPENSE_JV_MAP_PARAM_KEY]]], { fields: ["id", "value"], limit: 1 });
      if (rows.length) {
        expenseJvMappingParamId = rows[0].id;
        try {
          expenseJvMappingCache = JSON.parse(rows[0].value || "{}");
        } catch {
          expenseJvMappingCache = {};
          notes.push(`Expense JV: this database's account mapping table (System Parameter "${EXPENSE_JV_MAP_PARAM_KEY}") contains invalid JSON — treated as empty this run. Check it in Settings → Technical → System Parameters.`);
        }
      } else {
        expenseJvMappingCache = {};
      }
    } catch (e) {
      expenseJvMappingCache = {};
      notes.push(`Expense JV: couldn't read this database's account mapping table (${e.message}) — treated as empty this run; every line will need mapping.`);
    }
    return expenseJvMappingCache;
  }

  async function saveExpenseJvMapping(label, code, notes) {
    const key = String(label || "").trim().toLowerCase();
    if (!key || !code) return false;
    const map = expenseJvMappingCache || (expenseJvMappingCache = {});
    map[key] = String(code).trim();
    const value = JSON.stringify(map);
    try {
      if (expenseJvMappingParamId) {
        await odooExecute("ir.config_parameter", "write", [[expenseJvMappingParamId], { value }]);
      } else {
        const existing = await odooExecute("ir.config_parameter", "search", [[["key", "=", EXPENSE_JV_MAP_PARAM_KEY]]], { limit: 1 });
        if (existing.length) {
          expenseJvMappingParamId = existing[0];
          await odooExecute("ir.config_parameter", "write", [[expenseJvMappingParamId], { value }]);
        } else {
          expenseJvMappingParamId = await odooExecute("ir.config_parameter", "create", [{ key: EXPENSE_JV_MAP_PARAM_KEY, value }]);
        }
      }
      return true;
    } catch (e) {
      if (notes) notes.push(`Expense JV: mapped "${label}" for this session, but saving it back to the database failed (${e.message}) — it will need remapping next run.`);
      return false;
    }
  }

  // Purely informational — suggests Chart-of-Accounts candidates alongside
  // a block message, so a human can pick the right code quickly. NEVER
  // auto-applies a suggestion (Rule 5: unmapped text always blocks).
  async function suggestExpenseJvAccountCandidates(label) {
    const words = String(label || "").replace(/[().,]/g, " ").split(/\s+/).filter((w) => w.length > 3);
    if (!words.length) return [];
    try {
      const terms = words.slice(0, 3);
      const orPrefix = Array(Math.max(0, terms.length - 1)).fill("|");
      const nameConds = terms.map((w) => ["name", "ilike", w]);
      const rows = await odooExecute("account.account", "search_read", [orPrefix.concat(nameConds)], { fields: ["id", "code", "name", "account_type"], limit: 5 });
      return rows.map((r) => ({ code: r.code, name: r.name, type: r.account_type }));
    } catch {
      return [];
    }
  }

  async function resolveExpenseJvAccountByCode(code, label, notes) {
    if (expenseJvCodeAccountCache.has(code)) {
      const cached = expenseJvCodeAccountCache.get(code);
      if (cached.note) notes.push(cached.note);
      return cached;
    }
    let result;
    try {
      const ids = await odooExecute("account.account", "search", [[["code", "=", code]]], { limit: 2 });
      if (ids.length === 1) {
        result = { id: ids[0] };
      } else if (ids.length === 0) {
        result = { id: false, note: `Expense JV mapping table points "${label}" at account code ${code}, but no account with that code exists in this database — blocked rather than guessing. Fix the mapping table or the Chart of Accounts, then re-run.` };
      } else {
        result = { id: false, note: `Expense JV mapping table points "${label}" at account code ${code}, but more than one account in this database now shares that code — blocked; a human needs to sort out the duplicate code first.` };
      }
    } catch (e) {
      result = { id: false, note: `Account code ${code} lookup failed: ${e.message}` };
    }
    expenseJvCodeAccountCache.set(code, result);
    if (result.note) notes.push(result.note);
    return result;
  }

  // Rule 6 — the real Payables/Creditors control account, resolved by
  // account_type (never by name).
  async function resolveExpenseJvPayableAccount(notes) {
    if (expenseJvPayableAccountResult !== undefined) {
      if (expenseJvPayableAccountResult.note) notes.push(expenseJvPayableAccountResult.note);
      return expenseJvPayableAccountResult;
    }
    let result;
    try {
      const ids = await odooExecute("account.account", "search", [[["account_type", "=", "liability_payable"]]], { limit: 2, order: "code" });
      if (ids.length === 1) {
        result = { id: ids[0] };
      } else if (ids.length === 0) {
        result = { id: false, note: `No account with account_type "liability_payable" exists in this database — a Sundry Creditors line was blocked. Set up your real Payables control account first.` };
      } else {
        result = { id: ids[0], note: `More than one account has account_type "liability_payable" — used the first one (id ${ids[0]}, by code). If that's the wrong one, fix this Sundry Creditors line manually.` };
      }
    } catch (e) {
      result = { id: false, note: `Payables control account lookup failed: ${e.message}` };
    }
    expenseJvPayableAccountResult = result;
    if (result.note) notes.push(result.note);
    return result;
  }

  // Rules 4/5/6 combined: resolve one Expense JV line's account. `cleaned`
  // already has its By/To prefix stripped.
  async function resolveExpenseJvLineAccount(cleaned, hasSubledger, side, notes, unresolved) {
    if (SUNDRY_CREDITORS_RE.test(cleaned)) {
      if (!hasSubledger) {
        notes.push(`Expense JV: a "${cleaned}" line had no Subledger/Vendor name on it — blocked rather than guessing which vendor this payable belongs to.`);
        return { id: false, blocked: true };
      }
      const acct = await resolveExpenseJvPayableAccount(notes);
      return acct.id === false ? { id: false, blocked: true } : { id: acct.id, blocked: false };
    }

    const map = await loadExpenseJvMapping(notes);

    if (ROUND_OFF_RE.test(cleaned)) {
      const mapKey = side === "credit" ? EXPENSE_JV_ROUND_OFF_KEYS.credit : EXPENSE_JV_ROUND_OFF_KEYS.debit;
      const code = map[mapKey];
      const label = `Round Off (${side || "unknown side"})`;
      if (!code) {
        notes.push(`Expense JV: no mapping table entry for "${mapKey}" — blocked rather than guessed. Map it (with the correct account code) and re-run.`);
        if (unresolved) unresolved.push({ label: mapKey, suggestions: await suggestExpenseJvAccountCandidates("round off") });
        return { id: false, blocked: true };
      }
      const acct = await resolveExpenseJvAccountByCode(code, label, notes);
      return acct.id === false ? { id: false, blocked: true } : { id: acct.id, blocked: false };
    }

    const key = cleaned.trim().toLowerCase();
    const code = map[key];
    if (!code) {
      const suggestions = await suggestExpenseJvAccountCandidates(cleaned);
      const suggestText = suggestions.length ? ` Closest Chart-of-Accounts matches found (not applied — pick one): ${suggestions.map((s) => `${s.code} ${s.name} (${s.type})`).join("; ")}.` : "";
      notes.push(`Expense JV: no mapping table entry for account text "${cleaned}" — blocked rather than guessed. Add "${cleaned}" to this database's mapping table (with the correct account code) and re-run.${suggestText}`);
      if (unresolved) unresolved.push({ label: cleaned, suggestions });
      return { id: false, blocked: true };
    }
    const acct = await resolveExpenseJvAccountByCode(code, cleaned, notes);
    return acct.id === false ? { id: false, blocked: true } : { id: acct.id, blocked: false };
  }

  // Rule 1 — the one Miscellaneous/General journal every Expense JV entry
  // posts to, resolved once per upload (or overridden by the user's own
  // explicit sheet.expenseJvJournalId pick — see processExpenseJvSheet).
  async function findExpenseJvJournal() {
    const miscByCode = await odooExecute("account.journal", "search", [[["type", "=", "general"], ["code", "=", "MISC"]]], { limit: 2 });
    if (miscByCode.length === 1) return { id: miscByCode[0] };

    const general = await odooExecute("account.journal", "search_read", [[["type", "=", "general"]]], { fields: ["id", "name", "code"], order: "id" });
    if (!general.length) return { id: false, error: `No Miscellaneous/General journal found in your Chart of Accounts. Go to Accounting -> Configuration -> Journals and create one (Type: Miscellaneous, Code: MISC) first — Expense JV always posts there explicitly, never to a default or guessed journal.` };

    const byName = general.filter((j) => /miscellaneous/i.test(j.name || ""));
    if (byName.length === 1) return { id: byName[0].id, note: `No journal coded "MISC" was found — used "${byName[0].name}" (matched by name) instead. Recommended: set that journal's own Code to "MISC" in Odoo so this is unambiguous next time.` };
    if (general.length === 1) return { id: general[0].id, note: `Used the only type-"Miscellaneous/General" journal in this database: "${general[0].name}". Recommended: set its Code to "MISC" in Odoo so this is unambiguous next time.` };

    const list = general.map((j) => `"${j.name}" (code ${j.code || "—"}, id ${j.id})`).join(", ");
    return { id: false, error: `Found ${general.length} Miscellaneous/General-type journals and none is coded "MISC" or clearly named "Miscellaneous": ${list}. Either set the real Miscellaneous Operations journal's Code to "MISC" in Odoo, or pick the journal explicitly for this sheet below, then re-run.` };
  }

  // Rule 2 — idempotency keyed on Doc No, scoped to this exact journal +
  // move_type "entry".
  async function findExpenseJvExistingByDocNo(docNo, journalId, notes) {
    const ref = cleanText(String(docNo || ""));
    if (!ref) return null;
    try {
      const ids = await odooExecute("account.move", "search", [[["move_type", "=", "entry"], ["journal_id", "=", journalId], ["ref", "=", ref]]], { limit: 1 });
      return ids.length ? ids[0] : null;
    } catch (e) {
      notes.push(`Expense JV: idempotency check failed for Doc No "${docNo}" (${e.message}) — treated as new; check manually for an existing entry before relying on this row.`);
      return null;
    }
  }

  // Rule 2a — a hit here is never silently skipped/overridden; it's
  // surfaced for a human decision (see processExpenseJvSheet).
  async function findExpenseJvPossibleDuplicate(partnerId, docRef, amount, notes) {
    if (!partnerId) return null;
    try {
      const candidates = await odooExecute("account.move", "search_read", [[["partner_id", "=", partnerId]]], { fields: ["id", "name", "ref", "move_type", "state", "journal_id", "amount_total"], limit: 2000 });
      if (!candidates.length) return null;
      const normRef = cleanText(String(docRef || "")).toLowerCase();
      const refHits = normRef ? candidates.filter((c) => c.ref && cleanText(String(c.ref)).toLowerCase() === normRef) : [];
      const amountHits = (amount && amount > 0) ? candidates.filter((c) => Math.abs((c.amount_total || 0) - amount) < 0.01) : [];
      const hit = refHits[0] || amountHits[0];
      if (!hit) return null;
      return { id: hit.id, name: hit.name || `#${hit.id}`, journal: (hit.journal_id && hit.journal_id[1]) || "unknown journal", moveType: hit.move_type, state: hit.state, matchedOn: refHits.length ? "reference" : "amount" };
    } catch (e) {
      notes.push(`Expense JV: the cross-journal duplicate check (Rule 2a) failed for this voucher (${e.message}) — proceeded without it; double-check manually that this vendor/amount isn't already booked elsewhere in Odoo.`);
      return null;
    }
  }

  async function loadJournals() {
    return odooExecute("account.journal", "search_read", [[]], { fields: ["id", "name", "code", "type"], order: "name" });
  }

  // Discovers Studio/custom (x_...) fields on every model used across
  // ODOO_SCHEMAS and appends them as mappable columns. Resets each
  // schema's fields back to BASE_SCHEMA_FIELDS first so a repeated
  // connect/refresh never appends the same custom field twice.
  async function refreshCustomFields() {
    for (const key of Object.keys(ODOO_SCHEMAS)) ODOO_SCHEMAS[key].fields = BASE_SCHEMA_FIELDS[key].slice();

    const modelsToSchemaKeys = {};
    for (const key of Object.keys(ODOO_SCHEMAS)) {
      const model = ODOO_SCHEMAS[key].model;
      (modelsToSchemaKeys[model] = modelsToSchemaKeys[model] || []).push(key);
    }

    let foundAny = false;
    for (const [model, keys] of Object.entries(modelsToSchemaKeys)) {
      try {
        const fieldsInfo = await odooExecute(model, "fields_get", [], { attributes: ["string", "type", "relation", "selection", "readonly"] });
        const customNames = Object.keys(fieldsInfo).filter((n) => n.startsWith("x_") && !fieldsInfo[n].readonly);
        for (const name of customNames) {
          const desc = buildCustomFieldDescriptor(name, fieldsInfo[name]);
          if (!desc) continue;
          for (const key of keys) {
            if (ODOO_SCHEMAS[key].fields.some((f) => f.name === name)) continue;
            ODOO_SCHEMAS[key].fields.push(desc);
            foundAny = true;
          }
        }
      } catch (e) {
        // One model's custom-field discovery failing (e.g. the connected
        // user lacks access to that model) shouldn't block the rest.
      }
    }
    return foundAny;
  }

  // Expense JV processor.
  async function processExpenseJvSheet(sheet, onProgress) {
    const emit = (s) => { if (onProgress) onProgress(s); return s; };
    const mappedFields = sheet.analysis.mapping.filter((m) => m.field);
    const groupM = mappedFields.find((m) => m.field.isGroupKey);
    const docRefM = mappedFields.find((m) => m.field.isExpenseDocRefField);
    const dateM = mappedFields.find((m) => m.field.name === "date");
    const dueDateM = mappedFields.find((m) => m.field.isExpenseDueDateField);
    const particularsM = mappedFields.find((m) => m.field.isExpenseParticularsField);
    const subledgerM = mappedFields.find((m) => m.field.isExpenseSubledgerField);
    const debitM = mappedFields.find((m) => m.field.isExpenseDebitField);
    const creditM = mappedFields.find((m) => m.field.isExpenseCreditField);
    const narrationM = mappedFields.find((m) => m.field.isExpenseNarrationField);

    if (!groupM || !particularsM) {
      return emit({ status: "done", result: { total: 0, created_count: 0, duplicate_count: 0, failed_count: 1, errors: [{ row: 0, error: "Map both Doc No. and Particulars before importing.", data: {} }], notes: [] } });
    }

    const notes = [];
    let journalId = sheet.expenseJvJournalId || null;
    if (journalId) {
      notes.push(`Posted every entry in this sheet to the journal chosen for it (id ${journalId}).`);
    } else {
      const jvJournal = await findExpenseJvJournal();
      if (jvJournal.id === false) {
        return emit({ status: "done", result: { total: 0, created_count: 0, duplicate_count: 0, failed_count: 1, errors: [{ row: 0, error: jvJournal.error, data: {} }], notes: [] } });
      }
      journalId = jvJournal.id;
      if (jvJournal.note) notes.push(jvJournal.note);
    }

    const usableRows = sheet.rows.filter((row) => !isRowEmpty(row, sheet.headers));
    const groups = [];
    let lastKey;
    for (const row of usableRows) {
      const key = row[groupM.header];
      const isBlankKey = key === "" || key == null;
      if (!groups.length || (!isBlankKey && key !== lastKey)) groups.push([row]);
      else groups[groups.length - 1].push(row);
      if (!isBlankKey) lastKey = key;
    }

    emit({ status: "uploading", progress: 0, total: groups.length });

    const errors = [];
    const existingMatches = [];
    const possibleDuplicates = [];
    const blockedRows = [];
    const unresolvedMappingLabels = [];
    let successCount = 0;
    let duplicateCount = 0;

    for (let g = 0; g < groups.length; g++) {
      const group = groups[g];
      const docNo = cleanText(String(group[0][groupM.header] ?? ""));
      const voucherLabel = docNo || `row ${g + 1}`;
      try {
        const date = dateM ? cleanDate(group[0][dateM.header]) : null;
        let docRef = null, dueDate = null;
        const narrationParts = [];
        const expenseJvRows = [];

        for (const row of group) {
          if (docRefM && row[docRefM.header]) docRef = cleanText(String(row[docRefM.header]));
          if (dueDateM && row[dueDateM.header]) { const t = cleanDate(row[dueDateM.header]); if (t) dueDate = t; }
          if (narrationM && row[narrationM.header]) narrationParts.push(cleanText(String(row[narrationM.header])));
          const particularsRaw = particularsM ? row[particularsM.header] : "";
          const debit = debitM && row[debitM.header] !== "" && row[debitM.header] != null ? cleanNumber(row[debitM.header]) : 0;
          const credit = creditM && row[creditM.header] !== "" && row[creditM.header] != null ? cleanNumber(row[creditM.header]) : 0;
          const subledger = subledgerM && row[subledgerM.header] ? cleanText(String(row[subledgerM.header])) : "";
          if (!particularsRaw && !debit && !credit) continue;
          const { cleaned, side } = parseByToPrefix(particularsRaw || "");
          expenseJvRows.push({ raw: particularsRaw || "", cleaned, side, subledger, debit, credit });
        }

        if (!docNo) {
          errors.push({ row: g, error: "Doc No. is required for this voucher.", data: group[0] });
          continue;
        }
        if (!expenseJvRows.length) {
          blockedRows.push({ row: g, voucher: voucherLabel, reason: "no Particulars/Debit/Credit lines were found for this Doc No." });
          continue;
        }

        const totalDebit = expenseJvRows.reduce((s, l) => s + l.debit, 0);
        const totalCredit = expenseJvRows.reduce((s, l) => s + l.credit, 0);
        if (Math.round((totalDebit - totalCredit) * 100) !== 0) {
          blockedRows.push({ row: g, voucher: voucherLabel, reason: `doesn't balance: Total Debit ₹${totalDebit.toFixed(2)} vs Total Credit ₹${totalCredit.toFixed(2)} — check the source rows (a missing line, or a TDS/Round Off amount that didn't carry through). Nothing was created.` });
          continue;
        }

        const existingId = await findExpenseJvExistingByDocNo(docNo, journalId, notes);
        if (existingId) {
          duplicateCount++;
          existingMatches.push({ row: g, voucher: voucherLabel, moveId: existingId, detail: `already imported as Journal Entry #${existingId} in this journal` });
          continue;
        }

        const sundryLinesForDup = expenseJvRows.filter((l) => SUNDRY_CREDITORS_RE.test(l.cleaned));
        let vendorNameForDup = (sundryLinesForDup.find((l) => l.subledger) || {}).subledger || "";
        if (!vendorNameForDup) {
          const anyWithSubledger = expenseJvRows.find((l) => l.subledger);
          if (anyWithSubledger) vendorNameForDup = anyWithSubledger.subledger;
        }
        let dupPartnerId = null;
        if (vendorNameForDup) {
          const resolved = await resolveRelation({ relModel: "res.partner", autoCreate: false }, vendorNameForDup);
          if (resolved.id !== false) dupPartnerId = resolved.id;
        }
        if (dupPartnerId) {
          const netSundryAmount = sundryLinesForDup.reduce((s, l) => s + l.credit - l.debit, 0);
          const dupAmount = Math.round(netSundryAmount * 100) !== 0 ? netSundryAmount : totalDebit;
          const possibleDup = await findExpenseJvPossibleDuplicate(dupPartnerId, docRef, dupAmount, notes);
          if (possibleDup) {
            possibleDuplicates.push({ row: g, voucher: voucherLabel, vendor: vendorNameForDup, matchedOn: possibleDup.matchedOn, detail: `possible duplicate of ${possibleDup.name} (id ${possibleDup.id}) in journal "${possibleDup.journal}", ${possibleDup.moveType}/${possibleDup.state}, matched by ${possibleDup.matchedOn} — needs a human decision: skip this voucher if that record is the same expense already correctly booked, or delete that record first (if it's a leftover broken draft) and re-run this import.` });
            continue;
          }
        }

        const lineCmds = [];
        let blocked = false;
        for (const l of expenseJvRows) {
          const hasSubledger = !!l.subledger;
          const resolved = await resolveExpenseJvLineAccount(l.cleaned, hasSubledger, l.side, notes, unresolvedMappingLabels);
          if (resolved.blocked) { blocked = true; continue; }
          const lineVals = { account_id: resolved.id, name: l.raw || l.cleaned };
          if (l.debit) lineVals.debit = l.debit;
          if (l.credit) lineVals.credit = l.credit;
          if (l.subledger) {
            const payeeResolved = await resolveRelation({ relModel: "res.partner", autoCreate: true }, l.subledger);
            if (payeeResolved.note) notes.push(payeeResolved.note);
            if (payeeResolved.id !== false) lineVals.partner_id = payeeResolved.id;
          }
          if (dueDate && SUNDRY_CREDITORS_RE.test(l.cleaned)) lineVals.date_maturity = dueDate;
          lineCmds.push([0, 0, lineVals]);
        }

        if (blocked) {
          blockedRows.push({ row: g, voucher: voucherLabel, reason: "one or more lines couldn't be resolved to a mapped account (see the notes above for which text needs adding to the mapping table). Nothing was created for this voucher." });
          continue;
        }
        if (!lineCmds.length) {
          blockedRows.push({ row: g, voucher: voucherLabel, reason: "no line could be resolved into a Journal Entry line. Nothing was created." });
          continue;
        }

        const jeValues = { move_type: "entry", journal_id: journalId, ref: docNo, date, narration: narrationParts.join(" ") || undefined, line_ids: lineCmds };

        try {
          const newJeId = await withRetry(() => odooExecute("account.move", "create", [jeValues]));
          successCount++;
          try {
            const [readBack] = await odooExecute("account.move", "read", [[newJeId], ["line_ids", "journal_id", "amount_total"]]);
            const journalOk = readBack && Array.isArray(readBack.journal_id) && readBack.journal_id[0] === journalId;
            const lineCountOk = readBack && Array.isArray(readBack.line_ids) && readBack.line_ids.length === lineCmds.length;
            if (!readBack || !journalOk || !lineCountOk) {
              notes.push(`Voucher "${voucherLabel}": created Journal Entry #${newJeId}, but the read-back check found a mismatch (${!journalOk ? "wrong/missing journal" : `expected ${lineCmds.length} line(s), found ${readBack ? readBack.line_ids.length : "none"}`}) — review this entry manually before trusting it.`);
            } else {
              notes.push(`Voucher "${voucherLabel}"${docRef ? ` (Doc Ref "${docRef}")` : ""}: created draft Journal Entry #${newJeId} — ${lineCmds.length} line(s), verified by reading it back.`);
            }
          } catch (e) {
            notes.push(`Voucher "${voucherLabel}": created Journal Entry #${newJeId}, but the read-back verification itself failed (${e.message}) — check this entry manually.`);
          }
        } catch (e) {
          errors.push({ row: g, error: `Voucher "${voucherLabel}": ${e.message}`, data: group[0] });
          continue;
        }
      } catch (e) {
        errors.push({ row: g, error: e.message, data: group[0] });
      }
      emit({ status: "uploading", progress: g + 1, total: groups.length });
    }

    return emit({ status: "done", result: { total: groups.length, created_count: successCount, duplicate_count: duplicateCount, failed_count: errors.length, errors, notes, existingMatches, possibleDuplicates, blockedRows, unresolvedMappingLabels } });
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

    // These three modules act against records that already exist in Odoo
    // (they don't create/dedupe/write a fresh record the generic way below),
    // so they're routed to their own dedicated processors up front.
    if (schema.isTdsEntry) return processTdsEntrySheet(sheet, onProgress);
    if (schema.isBillAdditionalLines) return processAdditionalLinesSheet(sheet, onProgress, "purchase");
    if (schema.isInvoiceAdditionalLines) return processAdditionalLinesSheet(sheet, onProgress, "sales");
    if (schema.isVendorBillsModule) return processVendorBillsSheet(sheet, onProgress);
    if (schema.isSalesInvoicesModule) return processSalesInvoicesSheet(sheet, onProgress);
    if (schema.isCreditNoteModule) return processCreditNoteSheet(sheet, onProgress, "sales");
    if (schema.isDebitNoteModule) return processCreditNoteSheet(sheet, onProgress, "purchase");
    if (schema.isExpenseJv) return processExpenseJvSheet(sheet, onProgress);
    if (schema.isCashBook) return processCashBookSheet(sheet, onProgress);

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

    // Opening Balances (and anything else with an isAccountField column): an
    // Account name that doesn't already exist in the Chart of Accounts is
    // about to be auto-created. Rather than only guessing its type from
    // keywords (asset vs. liability from which side its rows landed on,
    // then the specific bucket from words in the name), ask once per sheet
    // — same "stop and show a list before creating anything" pattern as the
    // product-confirmation step above, just with a type picker per name
    // instead of a plain yes/no. sheet.accountTypeChoices, once set (even
    // to {} when nothing needed asking), means this sheet has already been
    // through this check — skip re-asking on a retry/re-click.
    const accountFields = mappedFields.filter((m) => m.field.isAccountField);
    if (accountFields.length && !sheet.accountTypeChoices) {
      emit({ status: "checking" });

      const missingAccounts = [];
      for (const m of accountFields) {
        const names = [...new Set(usableRows.map((row) => cleanText(String(row[m.header] ?? ""))).filter(Boolean))];
        if (!names.length) continue;
        try {
          const foundExact = await odooExecute("account.account", "search_read", [[["name", "in", names]]], { fields: ["name"] });
          const foundNames = new Set(foundExact.map((r) => String(r.name).toLowerCase()));
          const stillMissing = names.filter((n) => !foundNames.has(n.toLowerCase()));
          for (const name of stillMissing) {
            const ids = await odooExecute("account.account", "search", [[["name", "ilike", name]]], { limit: 1 });
            if (ids.length) continue;
            // Suggest a starting selection from the same signal
            // classifyAccountType would use (which column this name's own
            // rows put their amount in) — the user confirms or overrides
            // it, nothing is applied without their say-so.
            const rowsForName = usableRows.filter((row) => cleanText(String(row[m.header] ?? "")) === name);
            const hasDebit = debitHeader ? rowsForName.some((row) => row[debitHeader] !== "" && row[debitHeader] != null) : false;
            const hasCredit = creditHeader ? rowsForName.some((row) => row[creditHeader] !== "" && row[creditHeader] != null) : false;
            missingAccounts.push({ name, suggested: classifyAccountType(name, hasDebit, hasCredit) });
          }
        } catch (e) {
          for (const name of names) missingAccounts.push({ name, suggested: classifyAccountType(name, false, false) });
        }
      }

      if (missingAccounts.length) {
        missingAccounts.sort((a, b) => a.name.localeCompare(b.name));
        return emit({ status: "confirm-account-type", missingAccounts });
      }
      sheet.accountTypeChoices = {}; // nothing needed asking about — mark this sheet as checked
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
      let pendingExtraImageUrls = [];
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
              const cleanedAccountName = cleanText(String(raw));
              // Whatever type the user picked for this exact name in the
              // pre-flight "what type should this be?" prompt above
              // (sheet.accountTypeChoices) — falls back to the keyword
              // guess only if that prompt found nothing missing to ask
              // about.
              const explicitType = sheet.accountTypeChoices ? sheet.accountTypeChoices[cleanedAccountName.toLowerCase()] : null;
              const resolved = await resolveOrCreateAccount(cleanedAccountName, hasDebit, hasCredit, notes, explicitType);
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
            // Main product image — downloads the URL (or decodes the local
            // data: URI for an embedded picture) and puts the base64
            // content straight into the create()/write() values, same as
            // any other product.template field (image_1920 accepts it
            // directly).
            if (field.isProductImage) {
              const imgUrl = cleanText(String(raw));
              if (imgUrl) {
                try {
                  values.image_1920 = await urlToBase64(imgUrl);
                } catch (e) {
                  notes.push(`Could not load main image from "${imgUrl}": ${e.message}`);
                }
              }
              continue;
            }
            // eCommerce Extra Media — one or more URLs in the cell, split on
            // comma/semicolon/pipe. product.image records need the
            // product's id, which doesn't exist yet mid-row, so just
            // capture the list here and apply it after create/dedupe below
            // (see applyExtraImages).
            if (field.isExtraImages) {
              pendingExtraImageUrls = String(raw).split(/[;,|]/).map((u) => u.trim()).filter(Boolean);
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
            if (schema.model === "product.template") {
              await applyRelocation(dup.id, pendingRelocate.from, pendingRelocate.to, pendingRelocate.qty, notes);
              if (pendingExtraImageUrls.length) await applyExtraImages(dup.id, pendingExtraImageUrls, notes);
            }
            emit({ status: "uploading", progress: g + 1, total: groups.length });
            continue;
          }
        }

        const newId = await withRetry(() => odooExecute(schema.model, "create", [values]));
        createdIds.push(newId);
        if (pendingQty !== null && sheet.stockLocationId && schema.model === "product.template") await applyOnHandQuantity(newId, pendingQty, sheet.stockLocationId, notes);
        if (schema.model === "product.template") {
          await applyRelocation(newId, pendingRelocate.from, pendingRelocate.to, pendingRelocate.qty, notes);
          if (pendingExtraImageUrls.length) await applyExtraImages(newId, pendingExtraImageUrls, notes);
        }
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
    processTdsEntrySheet,
    processAdditionalLinesSheet,
    processVendorBillsSheet,
    processSalesInvoicesSheet,
    processCreditNoteSheet,
    processExpenseJvSheet,
    processCashBookSheet,
    loadJournals,
    refreshCustomFields,
    loadExpenseJvMapping,
    saveExpenseJvMapping,
    suggestExpenseJvAccountCandidates,
    get uid() { return uid; },
  };
}
