// ============================================================================
// emptyCompanyData.js  (rebuilt against the current app - July 2026 pass)
//
// Replaces seedData() as the starting point for a BRAND NEW COMPANY signup.
// It keeps the chart of accounts (required scaffolding - every posting
// function in the app references specific account IDs) but contains zero
// business activity: no demo transactions, customers, invoices, bills,
// inventory, employees, assets, leases, sales/purchase documents, time
// sheets, production records, locations, departments, or budgets.
//
// This file was regenerated from scratch against the CURRENT ces.jsx rather
// than patched forward from an older version, since the data model has
// grown substantially since the first pass (inventory types with their own
// GL accounts, Opening Balance Equity, locations/departments, budget
// tracking, time sheets, production records, and more). Every account below
// was copied verbatim from seedData()'s own accounts array, then verified
// field-by-field against the real app output before shipping - see
// DEPLOYMENT-GUIDE.md for the verification command.
// ============================================================================

function chartOfAccountsSkeleton() {
  return [
    { id: "1000", code: "1000", name: "Main Operating Account", type: "asset", subtype: "bank" },
    { id: "1010", code: "1010", name: "Reserve Account", type: "asset", subtype: "bank" },
    { id: "1050", code: "1050", name: "Petty Cash", type: "asset", subtype: "current" },
    { id: "1100", code: "1100", name: "Accounts Receivable", type: "asset", subtype: "current" },
    { id: "1200", code: "1200", name: "Finished Goods Inventory", type: "asset", subtype: "current", category: "Inventory" },
    { id: "1201", code: "1201", name: "Raw Materials Inventory", type: "asset", subtype: "current", category: "Inventory" },
    { id: "1202", code: "1202", name: "Work In Progress Inventory", type: "asset", subtype: "current", category: "Inventory" },
    { id: "1203", code: "1203", name: "Consumables & Supplies Inventory", type: "asset", subtype: "current", category: "Inventory" },
    { id: "1300", code: "1300", name: "Fixed Assets", type: "asset", subtype: "fixed" },
    { id: "1310", code: "1310", name: "Accumulated Depreciation", type: "asset", subtype: "fixed", normal: "credit", contra: true },
    { id: "2000", code: "2000", name: "Accounts Payable", type: "liability", subtype: "current" },
    { id: "2100", code: "2100", name: "VAT Payable", type: "liability", subtype: "current" },
    { id: "2200", code: "2200", name: "WHT Payable", type: "liability", subtype: "current" },
    { id: "2210", code: "2210", name: "PAYE Payable", type: "liability", subtype: "current" },
    { id: "2220", code: "2220", name: "Pension Payable", type: "liability", subtype: "current" },
    { id: "2230", code: "2230", name: "NHF Payable", type: "liability", subtype: "current" },
    { id: "2240", code: "2240", name: "Other Payroll Deductions Payable", type: "liability", subtype: "current" },
    { id: "2250", code: "2250", name: "Current Tax Payable", type: "liability", subtype: "current", category: "Current Liabilities" },
    { id: "2260", code: "2260", name: "Deferred Tax Liability", type: "liability", subtype: "noncurrent", category: "Deferred Tax Liabilities" },
    { id: "1450", code: "1450", name: "Deferred Tax Asset", type: "asset", subtype: "fixed", category: "Other Assets" },
    { id: "1320", code: "1320", name: "Right-of-Use Assets", type: "asset", category: "Fixed Assets" },
    { id: "1330", code: "1330", name: "Accumulated Depreciation - ROU Assets", type: "asset", contra: true, normal: "credit", category: "Fixed Assets" },
    { id: "2270", code: "2270", name: "Lease Liability", type: "liability", category: "Non-Current Liabilities" },
    { id: "2300", code: "2300", name: "Stamp Duty Payable", type: "liability", subtype: "current" },
    { id: "3000", code: "3000", name: "Owner's Equity", type: "equity" },
    { id: "4000", code: "4000", name: "Sales Revenue", type: "revenue" },
    { id: "4100", code: "4100", name: "Other Income", type: "revenue" },
    { id: "4200", code: "4200", name: "Sales Returns & Allowances", type: "revenue", contra: true, normal: "debit" },
    { id: "5000", code: "5000", name: "Cost of Goods Sold", type: "expense", subtype: "cogs" },
    { id: "5010", code: "5010", name: "Direct Materials", type: "expense", subtype: "cogs" },
    { id: "5020", code: "5020", name: "Direct Labour", type: "expense", subtype: "cogs" },
    { id: "5030", code: "5030", name: "Carriage / Freight Inwards", type: "expense", subtype: "cogs" },
    { id: "5040", code: "5040", name: "Production Overheads", type: "expense", subtype: "cogs" },
    { id: "5100", code: "5100", name: "Rent Expense", type: "expense" },
    { id: "5200", code: "5200", name: "Utilities Expense", type: "expense" },
    { id: "5300", code: "5300", name: "Payroll Expense", type: "expense" },
    { id: "5400", code: "5400", name: "Office Supplies", type: "expense" },
    { id: "5500", code: "5500", name: "Marketing Expense", type: "expense" },
    { id: "5600", code: "5600", name: "Bank Fees", type: "expense" },
    { id: "5700", code: "5700", name: "Depreciation Expense", type: "expense" },
    { id: "5900", code: "5900", name: "Purchase Returns & Allowances", type: "expense", subtype: "cogs", contra: true, normal: "credit" },
    { id: "5850", code: "5850", name: "Income Tax Expense", type: "expense", category: "Other Expenses" },
    { id: "5650", code: "5650", name: "Lease Interest Expense", type: "expense", category: "Finance Costs" },
    { id: "5750", code: "5750", name: "Inventory Write-down", type: "expense", subtype: "cogs", category: "Cost of Goods Sold (COGS)" },
    { id: "5950", code: "5950", name: "Bad Debt Expense (ECL)", type: "expense", category: "Other Expenses" },
    { id: "1150", code: "1150", name: "Allowance for Doubtful Accounts", type: "asset", contra: true, normal: "credit", category: "Accounts Receivable" },
    { id: "2290", code: "2290", name: "Deferred Revenue", type: "liability", category: "Current Liabilities" },
    { id: "2296", code: "2296", name: "Goods Received Not Invoiced", type: "liability", category: "Current Liabilities" },
    { id: "3900", code: "3900", name: "Opening Balance Equity", type: "equity", category: "Opening Balance Equity" },
    { id: "6010", code: "6010", name: "Foreign Exchange Gain / (Loss)", type: "revenue", subtype: "other", category: "Other Income" },
    { id: "5760", code: "5760", name: "Impairment Loss", type: "expense", category: "Other Expenses" },
    { id: "5770", code: "5770", name: "Provision Expense", type: "expense", category: "Other Expenses" },
    { id: "5780", code: "5780", name: "Revaluation Loss", type: "expense", category: "Other Expenses" },
    { id: "2295", code: "2295", name: "Provisions", type: "liability", category: "Non-Current Liabilities" },
    { id: "3200", code: "3200", name: "Revaluation Surplus", type: "equity", category: "Reserves" },
    { id: "6000", code: "6000", name: "Realized Gains & Losses", type: "revenue", subtype: "other" },
  ];
}

function defaultSettingsForNewCompany(companyName) {
  return {
    companyName: companyName || "My Company",
    currencyCode: "NGN",
    mode: "light",
    accent: "#2563EB",
    reportOptions: { showCodes: false, hideZeroLines: true },
    userName: "",
    accountingBasis: "accrual",
    payrollCountry: "NG",
    industry: "general",
    country: "NG",
    corporateTaxRate: 30,
    eclRates: { "Current": 0.5, "1-30 days": 1, "31-60 days": 5, "61-90 days": 15, "90+ days": 40 },
    lockDate: null,
  };
}

function subtypeForCategory(type, category) {
  if (type === "asset") {
    if (category === "Bank") return "bank";
    if (["Current Assets", "Cash", "Accounts Receivable"].includes(category)) return "current";
    return "fixed";
  }
  if (type === "liability") return ["Current Liabilities", "Accounts Payable", "Other Current Liabilities"].includes(category) ? "current" : "noncurrent";
  if (type === "revenue") return category === "Other Income" ? "other" : undefined;
  if (type === "expense") return category === "Cost of Goods Sold (COGS)" ? "cogs" : undefined;
  return undefined;
}
function categoryFromLegacy(a) {
  if (a.category) return a.category;
  if (a.type === "asset") {
    if (a.subtype === "bank") return "Bank";
    if (a.subtype === "fixed") return "Fixed Assets";
    if (a.id === "1100") return "Accounts Receivable";
    if ((a.name || "").toLowerCase().includes("cash")) return "Cash";
    return "Current Assets";
  }
  if (a.type === "liability") return a.id === "2000" ? "Accounts Payable" : "Current Liabilities";
  if (a.type === "equity") return "Owner's Equity";
  if (a.type === "revenue") return a.subtype === "other" ? "Other Income" : "Operating Income";
  if (a.type === "expense") {
    if (a.subtype === "cogs") return "Cost of Goods Sold (COGS)";
    if (a.id === "5700") return "Depreciation & Amortization";
    if (a.id === "5600") return "Finance Costs";
    return "Operating Expenses";
  }
  return "Other";
}
function normalizeAccount(a) {
  const category = categoryFromLegacy(a);
  const subtype = subtypeForCategory(a.type, category);
  return { parentId: null, description: "", taxAccountId: "", currency: "", ...a, category, subtype: subtype !== undefined ? subtype : a.subtype, status: a.status || "active" };
}

export function emptyCompanyData(companyName) {
  return {
    settings: defaultSettingsForNewCompany(companyName),
    accounts: chartOfAccountsSkeleton().map(normalizeAccount),
    banks: [],
    transactions: [],
    invoices: [],
    bills: [],
    expenses: [],
    inventory: [],
    inventoryLots: [],
    fixedAssets: [],
    payments: [],
    taxGroups: [],
    projects: [],
    locations: [],
    departments: [],
    budgets: {},
    budgetAccounts: [],
    favoriteReports: ["pl", "ar-aging", "inventory-summary"],
    bankFeed: [],
    categoryRules: [],
    bin: [],
    employees: [],
    payrollRuns: [],
    recurringJournals: [],
    reconciliations: [],
    auditLog: [],
    taxProvisions: [],
    leases: [],
    deferredRevenueSchedules: [],
    provisions: [],
    salesOrders: [],
    salesReceipts: [],
    creditNotes: [],
    salesReturns: [],
    purchaseOrders: [],
    purchaseReceipts: [],
    vendorCredits: [],
    timesheets: [],
    productionRecords: [],
    nextProductionNum: 1,
    openingBalances: { asOfDate: new Date().toISOString().slice(0, 10), accountAmounts: {}, customerBalances: [], vendorBalances: [], posted: false, postedDate: null },
    nextInvoiceNum: 1001,
    nextBillNum: 2001,
    nextSalesOrderNum: 1,
    nextSalesReceiptNum: 1,
    nextCreditNoteNum: 1,
    nextSalesReturnNum: 1,
    nextPurchaseOrderNum: 1,
    nextPurchaseReceiptNum: 1,
    nextVendorCreditNum: 1,
  };
}
