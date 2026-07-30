import React, { useState, useEffect, useMemo, useCallback, createContext, useContext } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  LineChart, Line, AreaChart, Area, BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from "recharts";

/* =================================================================
   CAPITAL EDGE STELLAR - multi-bank accounting, modeling & tax engine
================================================================= */

const FONT_IMPORT = `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
`;

const ACCOUNT_TYPES = ["asset", "liability", "equity", "revenue", "expense"];
const TYPE_LABELS = { asset: "Assets", liability: "Liabilities", equity: "Equity", revenue: "Income", expense: "Expenses" };
// Standard subcategories per primary type (displayed on the COA and used to
// derive each account's behavior in reports and ratios).
const ACCOUNT_CATEGORIES = {
  asset: ["Current Assets", "Non-Current Assets", "Cash", "Bank", "Accounts Receivable", "Fixed Assets", "Intangible Assets", "Other Assets"],
  liability: ["Current Liabilities", "Non-Current Liabilities", "Accounts Payable", "Deferred Tax Liabilities", "Other Current Liabilities", "Other Liabilities"],
  equity: ["Share Capital", "Retained Earnings", "Reserves", "Owner's Equity", "Other Equity Accounts"],
  revenue: ["Operating Income", "Other Income"],
  expense: ["Cost of Goods Sold (COGS)", "Operating Expenses", "Administrative Expenses", "Selling & Distribution Expenses", "Finance Costs", "Depreciation & Amortization", "Other Expenses"],
};
// Behavioral subtype derived from the display category - this is what the
// engine (liquidity ratios, P&L sections, bank logic) keys off.
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
// Accounts the posting engine references directly - protected from deletion.
const SYSTEM_ACCOUNT_IDS = ["1100", "1200", "1201", "1202", "1203", "1300", "1310", "2000", "3000", "4000", "4200", "5000", "5700", "5900", "6000"];
// Which GL account holds each inventory type's value - the single lookup
// every posting function routes through, so adding a type never means
// hunting down "1200" hardcoded in a dozen places again.
const INVENTORY_TYPES = [
  { id: "raw_material", label: "Raw Material", accountId: "1201" },
  { id: "wip", label: "Work In Progress", accountId: "1202" },
  { id: "finished_good", label: "Finished Goods", accountId: "1200" },
  { id: "consumable", label: "Consumables / Supplies", accountId: "1203" },
];
function inventoryAccountForType(type) {
  return INVENTORY_TYPES.find(t => t.id === type)?.accountId || "1200";
}
function inventoryAccountForItem(data, itemId) {
  const item = data.inventory.find(i => i.id === itemId);
  return inventoryAccountForType(item?.inventoryType);
}
// Which inventory types actually come up for a given line of business -
// used to default and reorder the type picker so the right option is
// usually already selected, without ever hiding the other three.
const INDUSTRY_INVENTORY_TYPES = {
  general: ["finished_good", "raw_material", "wip", "consumable"],
  retail: ["finished_good", "consumable"],
  services: ["consumable"],
  manufacturing: ["raw_material", "wip", "finished_good", "consumable"],
  construction: ["raw_material", "wip", "finished_good"],
  hospitality: ["raw_material", "consumable", "finished_good"],
  healthcare: ["consumable", "finished_good"],
  technology: ["finished_good", "consumable"],
  agriculture: ["raw_material", "finished_good", "consumable"],
  nonprofit: ["consumable", "finished_good"],
  logistics: ["consumable", "finished_good"],
};
function inventoryTypesForIndustry(industry) {
  const order = INDUSTRY_INVENTORY_TYPES[industry] || INDUSTRY_INVENTORY_TYPES.general;
  return [...order, ...INVENTORY_TYPES.map(t => t.id).filter(id => !order.includes(id))].map(id => INVENTORY_TYPES.find(t => t.id === id));
}
// Smallest unused numeric account code >= start, stepping by `step`. Checks
// both codes and ids so a new account can never collide with an existing one.
/* ---------- Nigeria statutory payroll engine (PITA, Finance Act bands) ----------
   Gross is annualized; CRA = higher of NGN200,000 or 1% of gross, plus 20% of
   gross. Employee pension 8% and NHF 2.5% (both optional per employee) are
   tax-deductible. PAYE bands: 7% / 11% / 15% / 19% / 21% / 24%; minimum tax
   1% of gross when taxable income is nil. Rates current as of the 2023 Acts -
   verify against current FIRS/state IRS guidance before filing. */
function computeNigeriaPayroll(emp) {
  const grossM = Number(emp.grossMonthly) || 0;
  const grossA = grossM * 12;
  const pensionA = emp.pension !== false ? grossA * 0.08 : 0;
  const nhfA = emp.nhf === true ? grossA * 0.025 : 0;
  const cra = Math.max(200000, grossA * 0.01) + grossA * 0.20;
  const taxableA = Math.max(0, grossA - cra - pensionA - nhfA);
  const bands = [[300000, 0.07], [300000, 0.11], [500000, 0.15], [500000, 0.19], [1600000, 0.21], [Infinity, 0.24]];
  let remaining = taxableA, payeA = 0;
  for (const [width, rate] of bands) { if (remaining <= 0) break; const slice = Math.min(remaining, width); payeA += slice * rate; remaining -= slice; }
  if (taxableA <= 0 && grossA > 0) payeA = grossA * 0.01; // minimum tax
  const custom = (emp.deductions || []).reduce((s, dd) => s + (Number(dd.amount) || 0), 0);
  const payeM = payeA / 12, pensionM = pensionA / 12, nhfM = nhfA / 12;
  return { grossM, payeM, pensionM, nhfM, customM: custom, netM: grossM - payeM - pensionM - nhfM - custom, craA: cra, taxableA };
}
// Advance a date by one recurrence period.
function advanceDate(iso, frequency) {
  const d = new Date(iso);
  if (frequency === "weekly") { d.setDate(d.getDate() + 7); return localDateToISO(d); }
  const day = d.getDate();
  let months = frequency === "quarterly" ? 3 : frequency === "yearly" ? 12 : 1;
  if (frequency === "yearly") {
    const y = d.getFullYear() + 1, m = d.getMonth();
    const lastDay = new Date(y, m + 1, 0).getDate();
    return localDateToISO(new Date(y, m, Math.min(d.getDate(), lastDay)));
  }
  // Clamp to the target month's last day instead of overflowing (e.g. Jan 31 + 1mo -> Feb 28/29, not Mar 3).
  const targetMonthFirst = new Date(d.getFullYear(), d.getMonth() + months, 1);
  const lastDayOfTargetMonth = new Date(targetMonthFirst.getFullYear(), targetMonthFirst.getMonth() + 1, 0).getDate();
  targetMonthFirst.setDate(Math.min(day, lastDayOfTargetMonth));
  return localDateToISO(targetMonthFirst);
}
// Materialize every recurring journal occurrence that has come due (catch-up
// on load), posting each with its scheduled date. Returns the updated payload
// and how many entries were posted.
function processRecurringJournals(data) {
  let next = data; let posted = 0;
  const today = todayStr();
  const updated = (next.recurringJournals || []).map(r => {
    if (!r.active) return r;
    let cursor = r.nextDate;
    let rr = r;
    let guard = 0;
    while (cursor && cursor <= today && (!r.endDate || cursor <= r.endDate) && guard < 120) {
      const txn = buildTxn(`${r.memo} (recurring)`, cursor, r.lines.map(l => ({ ...l })), "recurring");
      if (!txn) break;
      next = { ...next, transactions: [...next.transactions, txn] };
      posted++;
      cursor = advanceDate(cursor, r.frequency);
      rr = { ...rr, nextDate: cursor, lastRun: today };
      guard++;
    }
    if (rr.endDate && rr.nextDate > rr.endDate) rr = { ...rr, active: false };
    return rr;
  });
  return { data: { ...next, recurringJournals: updated }, posted };
}
function nextFreeCode(accounts, start = 1020, step = 10) {
  const taken = new Set(accounts.flatMap(a => [String(a.code), String(a.id)]));
  let c = start;
  while (taken.has(String(c))) c += step;
  return String(c);
}
function activeAccounts(data) { return data.accounts.filter(a => a.status !== "inactive"); }
const DEBIT_NORMAL = { asset: true, liability: false, equity: false, revenue: false, expense: true };
const uid = (p) => `${p}_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
// Formats a calendar date as YYYY-MM-DD with no timezone conversion - pure
// string building from already-known year/month(0-idx)/day numbers.
function isoDate(y, m, d) { return `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`; }
// Reads the LOCAL calendar date off a Date object (getFullYear/getMonth/
// getDate are local by definition) and formats it directly - critically,
// this never routes through toISOString(), which converts to UTC first and
// silently rolls the date back by one in any positive-UTC-offset timezone
// (e.g. Lagos, UTC+1): local midnight Jan 1 becomes UTC 23:00 Dec 31, so
// .toISOString().slice(0,10) would wrongly report "Dec 31". That was the
// cause of "This Year" / "This Month" starting a day early.
function localDateToISO(d) { return isoDate(d.getFullYear(), d.getMonth(), d.getDate()); }
// For dates decoded by the xlsx library's cellDates option: SheetJS's serial
// -> Date conversion can drift by up to ~1 minute due to floating-point
// rounding in its date math, which can land the timestamp just before a UTC
// day boundary and shift the apparent calendar day. Snapping to the nearest
// UTC midnight before reading it back corrects that drift; using UTC getters
// keeps the result independent of the viewer's timezone.
function excelDateToISO(d) { return new Date(Math.round(d.getTime() / 86400000) * 86400000).toISOString().slice(0, 10); }
const todayStr = () => localDateToISO(new Date());
const fmtDate = (d) => new Date(d).toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
const monthKey = (d) => { const dt = new Date(d); return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}`; };

const CURRENCIES = [
  { code: "NGN", symbol: "₦", locale: "en-NG" },
  { code: "USD", symbol: "$", locale: "en-US" },
  { code: "GBP", symbol: "£", locale: "en-GB" },
  { code: "EUR", symbol: "€", locale: "de-DE" },
  { code: "GHS", symbol: "₵", locale: "en-GH" },
  { code: "KES", symbol: "KSh", locale: "en-KE" },
  { code: "ZAR", symbol: "R", locale: "en-ZA" },
  { code: "XOF", symbol: "CFA", locale: "fr-SN" },
];

const THEME_BASE = {
  // Premium enterprise SaaS look (Linear / Stripe / Fluent): an all-white work
  // surface, hairline borders, layered soft shadows, blue + emerald accents.
  light: {
    mode: "light", ink: "#0F1728",
    paper: "#F8FAFC", panel: "#FFFFFF", panel2: "#F4F6FA",
    border: "#E6EAF1", text: "#101828", muted: "#667085",
    emerald: "#12A66F", amber: "#D9911F", rose: "#DC4C5E",
    sidebarBg: "#FFFFFF", sidebarBorder: "#EAECF2", sidebarText: "#344054", sidebarMuted: "#98A2B3",
    shadowSm: "0 1px 2px rgba(16,24,40,0.05)",
    shadowMd: "0 1px 3px rgba(16,24,40,0.06), 0 2px 8px rgba(16,24,40,0.05)",
    shadowLg: "0 4px 12px rgba(16,24,40,0.08), 0 12px 32px rgba(16,24,40,0.07)",
  },
  dark: {
    mode: "dark", ink: "#0B1220",
    paper: "#0C1220", panel: "#131B2C", panel2: "#101827",
    border: "#243049", text: "#E7ECF7", muted: "#8B97B8",
    emerald: "#34C283", amber: "#D89A4B", rose: "#E0616D",
    sidebarBg: "#0F1626", sidebarBorder: "#1E2A42", sidebarText: "#C7D0E0", sidebarMuted: "#6B7A99",
    shadowSm: "0 1px 2px rgba(0,0,0,0.3)",
    shadowMd: "0 2px 8px rgba(0,0,0,0.35)",
    shadowLg: "0 8px 28px rgba(0,0,0,0.45)",
  },
};
function buildTheme(settings) {
  const base = THEME_BASE[settings.mode] || THEME_BASE.light;
  const accent = settings.accent || "#2563EB";
  return { ...base, accent, accentDeep: shade(accent, -22), accentSoft: shade(accent, settings.mode === "dark" ? -34 : 90) };
}
function shade(hex, pct) {
  try {
    const n = parseInt(hex.replace("#", ""), 16);
    let r = (n >> 16) + Math.round(2.55 * pct);
    let g = ((n >> 8) & 0xff) + Math.round(2.55 * pct);
    let b = (n & 0xff) + Math.round(2.55 * pct);
    r = Math.max(0, Math.min(255, r)); g = Math.max(0, Math.min(255, g)); b = Math.max(0, Math.min(255, b));
    return "#" + [r, g, b].map(v => v.toString(16).padStart(2, "0")).join("");
  } catch { return hex; }
}
// Used to tint hover glows and 3D-effect shadows with whatever accent color
// is chosen in Settings, instead of a fixed color.
function hexToRgba(hex, alpha) {
  try {
    const n = parseInt(hex.replace("#", ""), 16);
    return `rgba(${(n >> 16) & 0xff}, ${(n >> 8) & 0xff}, ${n & 0xff}, ${alpha})`;
  } catch { return `rgba(37, 99, 235, ${alpha})`; }
}

const UIContext = createContext(null);
const useUI = () => useContext(UIContext);

/* ---------------- default settings & seed data ---------------- */
function defaultSettings() {
  return { companyName: "Capital Edge Stellar", currencyCode: "NGN", mode: "light", accent: "#2563EB", reportOptions: { showCodes: false, hideZeroLines: true }, userName: "Admin", accountingBasis: "accrual", payrollCountry: "NG", industry: "general", country: "NG", corporateTaxRate: 30, eclRates: { "Current": 0.5, "1-30 days": 1, "31-60 days": 5, "61-90 days": 15, "90+ days": 40 }, lockDate: null };
}

function seedData() {
  const accounts = [
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
  const banks = [
    { id: "bank_main", name: "First Capital Bank - Operating", accountNumber: "0021 4477 190", accountId: "1000" },
    { id: "bank_reserve", name: "Union Trust Bank - Reserve", accountNumber: "0089 1123 004", accountId: "1010" },
  ];
  const mk = (offsetDays) => { const d = new Date(); d.setDate(d.getDate() - offsetDays); return localDateToISO(d); };
  const transactions = [
    { id: uid("txn"), date: mk(90), memo: "Owner capital injection", source: "manual", lines: [{ accountId: "1000", debit: 15000000, credit: 0 }, { accountId: "3000", debit: 0, credit: 15000000 }] },
    { id: uid("txn"), date: mk(88), memo: "Transfer to Reserve Account", source: "transfer", lines: [{ accountId: "1010", debit: 3000000, credit: 0 }, { accountId: "1000", debit: 0, credit: 3000000 }] },
    { id: uid("txn"), date: mk(85), memo: "Petty cash top-up", source: "transfer", lines: [{ accountId: "1050", debit: 50000, credit: 0 }, { accountId: "1000", debit: 0, credit: 50000 }] },
    { id: uid("txn"), date: mk(80), memo: "Invoice INV-1001 - Acme Corp", source: "invoice", lines: [{ accountId: "1100", debit: 462500, credit: 0 }, { accountId: "2100", debit: 37500, credit: 0 }, { accountId: "4000", debit: 0, credit: 500000 }] },
    { id: uid("txn"), date: mk(75), memo: "Rent - August", source: "expense", lines: [{ accountId: "5100", debit: 120000, credit: 0 }, { accountId: "1000", debit: 0, credit: 120000 }] },
    { id: uid("txn"), date: mk(70), memo: "Payment received INV-1001", source: "payment", lines: [{ accountId: "1000", debit: 462500, credit: 0 }, { accountId: "1100", debit: 0, credit: 462500 }] },
    { id: uid("txn"), date: mk(60), memo: "Utilities - August", source: "expense", lines: [{ accountId: "5200", debit: 24000, credit: 0 }, { accountId: "1000", debit: 0, credit: 24000 }] },
    { id: uid("txn"), date: mk(55), memo: "Fixed asset purchase - Delivery Van", source: "fixed-asset", lines: [{ accountId: "1300", debit: 2400000, credit: 0 }, { accountId: "1000", debit: 0, credit: 2400000 }] },
    { id: uid("txn"), date: mk(50), memo: "Bill from Nova Supplies (inventory)", source: "bill", docId: "BILL-2001", lines: [{ accountId: "1200", debit: 818000, credit: 0 }, { accountId: "2000", debit: 0, credit: 777100 }, { accountId: "2200", debit: 0, credit: 40900 }] },
    { id: uid("txn"), date: mk(45), memo: "Payroll - September", source: "expense", lines: [{ accountId: "5300", debit: 420000, credit: 0 }, { accountId: "1000", debit: 0, credit: 420000 }] },
    { id: uid("txn"), date: mk(30), memo: "Rent - September", source: "expense", lines: [{ accountId: "5100", debit: 120000, credit: 0 }, { accountId: "1000", debit: 0, credit: 120000 }] },
    { id: uid("txn"), date: mk(20), memo: "Marketing - social ads", source: "expense", lines: [{ accountId: "5500", debit: 65000, credit: 0 }, { accountId: "1010", debit: 0, credit: 65000 }] },
    { id: uid("txn"), date: mk(12), memo: "Invoice INV-1002 - Northgate Retail", source: "invoice", lines: [{ accountId: "1100", debit: 856000, credit: 0 }, { accountId: "4000", debit: 0, credit: 800000 }, { accountId: "2100", debit: 0, credit: 56000 }] },
    { id: uid("txn"), date: mk(6), memo: "Utilities - September", source: "expense", lines: [{ accountId: "5200", debit: 26000, credit: 0 }, { accountId: "1000", debit: 0, credit: 26000 }] },
    { id: uid("txn"), date: mk(2), memo: "Bank service fee", source: "expense", lines: [{ accountId: "5600", debit: 3500, credit: 0 }, { accountId: "1000", debit: 0, credit: 3500 }] },
    { id: uid("txn"), date: mk(1), memo: "Depreciation - Delivery Van", source: "depreciation", lines: [{ accountId: "5700", debit: 40000, credit: 0 }, { accountId: "1310", debit: 0, credit: 40000 }] },
  ];
  const invoices = [
    { id: "INV-1001", customer: "Acme Corp", date: mk(80), dueDate: mk(50), items: [{ desc: "Consulting - August", qty: 1, price: 500000, inventoryId: null }], taxes: [{ id: uid("tax"), name: "VAT", rate: 7.5, mode: "percent", effect: "deduct", accountId: "2100" }], status: "paid", amountPaid: 462500, projectId: null },
    { id: "INV-1002", customer: "Northgate Retail", date: mk(12), dueDate: mk(-2), items: [{ desc: "Product batch - Q3", qty: 40, price: 20000, inventoryId: null }], taxes: [{ id: uid("tax"), name: "VAT", rate: 7, mode: "percent", effect: "add", accountId: "2100" }], status: "sent", amountPaid: 0, projectId: "seed_project" },
  ];
  const bills = [
    { id: "BILL-2001", vendor: "Nova Supplies", date: mk(50), dueDate: mk(20), items: [{ desc: "Stock - assorted units", qty: 200, price: 4000, inventoryId: "inv_widget" }, { desc: "Stock - protective cases", qty: 12, price: 1500, inventoryId: "inv_case" }], taxes: [{ id: uid("tax"), name: "WHT", rate: 5, mode: "percent", effect: "deduct", accountId: "2200" }], status: "unpaid", amountPaid: 0, projectId: "seed_project" },
  ];
  const inventory = [
    { id: "inv_widget", sku: "WDG-001", name: "Standard Widget", qty: 200, unitCost: 4000, salePrice: 6500, reorderLevel: 50, nrv: 6500, inventoryType: "finished_good" },
    { id: "inv_case", sku: "CASE-014", name: "Protective Case", qty: 12, unitCost: 1500, salePrice: 2800, reorderLevel: 20, nrv: 1200, inventoryType: "finished_good" },
  ];
  const inventoryLots = [
    { id: uid("lot"), itemId: "inv_widget", date: mk(50), qty: 200, remainingQty: 200, unitCost: 4000, sourceDocId: "BILL-2001" },
    { id: uid("lot"), itemId: "inv_case", date: mk(50), qty: 12, remainingQty: 12, unitCost: 1500, sourceDocId: "BILL-2001" },
  ];
  const fixedAssets = [
    { id: uid("fa"), name: "Delivery Van", category: "Vehicles", purchaseDate: mk(55), cost: 2400000, salvageValue: 200000, usefulLifeMonths: 60, accumulatedDepreciation: 40000, lastDepreciationMonth: monthKey(mk(1)), taxRatePct: 25, accumulatedTaxDepreciation: 60000 },
  ];
  const leaseRouCost = 9681240; // PV of 36 monthly payments of 350,000 at 18% annual discount rate
  const leases = [
    { id: uid("lease"), name: "Office Lease - Victoria Island", lessor: "Landmark Properties", startDate: mk(2), termMonths: 36, paymentAmount: 350000, discountRate: 18, bankId: "bank_main", rouCost: leaseRouCost, liabilityBalance: leaseRouCost, accumulatedROUDep: 0, periodsRun: 0, lastRunMonth: null, status: "active" },
  ];
  const leaseCommencementTxn = { id: uid("txn"), date: mk(2), memo: "Lease commencement - Office Lease - Victoria Island", source: "lease", meta: { leaseId: leases[0].id, kind: "commencement" }, lines: [{ accountId: "1320", debit: leaseRouCost, credit: 0 }, { accountId: "2270", debit: 0, credit: leaseRouCost }] };
  const payments = [
    { id: uid("pay"), date: mk(70), type: "received", amount: 462500, bankId: "bank_main", relatedType: "invoice", relatedId: "INV-1001", memo: "Payment received INV-1001" },
  ];
  const taxGroups = [
    { id: "grp_vat_wht", name: "VAT + WHT (parallel)", components: [
      { id: uid("c"), name: "VAT", rate: 7.5, mode: "percent", effect: "deduct", accountId: "2100" },
      { id: uid("c"), name: "WHT", rate: 5, mode: "percent", effect: "deduct", accountId: "2200" },
    ] },
  ];
  const projects = [
    { id: "seed_project", name: "Northgate Rollout", client: "Northgate Retail", budget: 900000 },
  ];
  const categoryRules = [
    { id: uid("rule"), keyword: "utilities", accountId: "5200", direction: "out", bankId: "bank_main" },
  ];
  return { settings: defaultSettings(), accounts: accounts.map(normalizeAccount), banks, transactions: [...transactions, leaseCommencementTxn], invoices, bills, expenses: [], inventory, inventoryLots, fixedAssets, payments, taxGroups, projects, locations: [], departments: [], budgets: {}, budgetAccounts: [], favoriteReports: ["pl", "ar-aging", "inventory-summary"], bankFeed: [], categoryRules, bin: [], employees: [], payrollRuns: [], recurringJournals: [], reconciliations: [], auditLog: [], taxProvisions: [], leases, deferredRevenueSchedules: [], provisions: [], salesOrders: [], salesReceipts: [], creditNotes: [], salesReturns: [], purchaseOrders: [], purchaseReceipts: [], vendorCredits: [], timesheets: [], productionRecords: [], nextProductionNum: 1, openingBalances: { asOfDate: todayStr(), accountAmounts: {}, customerBalances: [], vendorBalances: [], posted: false, postedDate: null }, nextInvoiceNum: 1003, nextBillNum: 2002, nextSalesOrderNum: 1, nextSalesReceiptNum: 1, nextCreditNoteNum: 1, nextSalesReturnNum: 1, nextPurchaseOrderNum: 1, nextPurchaseReceiptNum: 1, nextVendorCreditNum: 1 };
}

// Upgrade a previously saved payload so older data works with new features:
// fills in newly added collections and injects any GL accounts (by code)
// that didn't exist when the data was first saved.
function migrateData(saved) {
  const seed = seedData();
  const merged = { ...seed, ...saved };
  // Repair a specific unbalanced demo transaction shipped in an earlier
  // version (WHT line had debit/credit reversed and a stale amount, and
  // separately never included the Protective Case's value even though that
  // item exists in inventory) that would otherwise make the Cash Flow
  // Statement fail to reconcile and understate inventory on the Balance Sheet.
  merged.transactions = (merged.transactions || []).map(t => {
    if (t.memo === "Bill from Nova Supplies (inventory)" && (t.lines.some(l => l.accountId === "2200" && l.debit === 37200) || t.lines.some(l => l.accountId === "1200" && l.debit === 800000))) {
      return { ...t, docId: "BILL-2001", lines: [{ accountId: "1200", debit: 818000, credit: 0 }, { accountId: "2000", debit: 0, credit: 777100 }, { accountId: "2200", debit: 0, credit: 40900 }] };
    }
    return t;
  });
  merged.settings = { ...seed.settings, ...(saved.settings || {}), reportOptions: { ...seed.settings.reportOptions, ...((saved.settings || {}).reportOptions || {}) } };
  merged.inventoryLots = saved.inventoryLots || (saved.inventory || []).filter(i => i.qty > 0).map(i => ({ id: uid("lot"), itemId: i.id, date: todayStr(), qty: i.qty, remainingQty: i.qty, unitCost: i.unitCost }));
  merged.projects = saved.projects || [];
  merged.budgets = saved.budgets || {};
  merged.favoriteReports = saved.favoriteReports || ["pl", "ar-aging", "inventory-summary"];
  merged.bankFeed = saved.bankFeed || [];
  merged.categoryRules = (saved.categoryRules || []).map(r => r.bankId ? r : { ...r, bankId: (merged.banks[0] || {}).id });
  merged.bin = saved.bin || [];
  merged.employees = (saved.employees || []).map(e => ({ deductions: [], pension: true, nhf: false, ...e }));
  merged.fixedAssets = (merged.fixedAssets || []).map(a => ({ taxRatePct: a.taxRatePct ?? 25, accumulatedTaxDepreciation: a.accumulatedTaxDepreciation ?? a.accumulatedDepreciation ?? 0, accumulatedImpairment: a.accumulatedImpairment ?? 0, valuationModel: a.valuationModel || "cost", revaluationSurplus: a.revaluationSurplus ?? 0, ...a }));
  merged.taxProvisions = saved.taxProvisions || [];
  merged.leases = saved.leases || [];
  merged.deferredRevenueSchedules = saved.deferredRevenueSchedules || [];
  merged.provisions = saved.provisions || [];
  merged.salesOrders = saved.salesOrders || [];
  merged.salesReceipts = saved.salesReceipts || [];
  merged.creditNotes = saved.creditNotes || [];
  merged.salesReturns = saved.salesReturns || [];
  merged.nextSalesOrderNum = saved.nextSalesOrderNum || 1;
  merged.nextSalesReceiptNum = saved.nextSalesReceiptNum || 1;
  merged.nextCreditNoteNum = saved.nextCreditNoteNum || 1;
  merged.nextSalesReturnNum = saved.nextSalesReturnNum || 1;
  merged.purchaseOrders = saved.purchaseOrders || [];
  merged.purchaseReceipts = saved.purchaseReceipts || [];
  merged.vendorCredits = saved.vendorCredits || [];
  merged.nextPurchaseOrderNum = saved.nextPurchaseOrderNum || 1;
  merged.nextPurchaseReceiptNum = saved.nextPurchaseReceiptNum || 1;
  merged.nextVendorCreditNum = saved.nextVendorCreditNum || 1;
  merged.openingBalances = saved.openingBalances || { asOfDate: todayStr(), accountAmounts: {}, customerBalances: [], vendorBalances: [], posted: false, postedDate: null };
  if (merged.settings.corporateTaxRate === undefined) merged.settings.corporateTaxRate = 30;
  if (!merged.settings.eclRates) merged.settings.eclRates = { "Current": 0.5, "1-30 days": 1, "31-60 days": 5, "61-90 days": 15, "90+ days": 40 };
  if (merged.settings.lockDate === undefined) merged.settings.lockDate = null;
  merged.timesheets = saved.timesheets || [];
  merged.locations = saved.locations || [];
  merged.departments = saved.departments || [];
  merged.productionRecords = saved.productionRecords || [];
  merged.nextProductionNum = saved.nextProductionNum || 1;
  merged.inventory = (merged.inventory || []).map(i => i.inventoryType ? i : { ...i, inventoryType: "finished_good" });
  // Old shape was flat: { accountId: monthlyAmount }, applied to every month
  // uniformly. New shape is per-period: { "2026-07": { accountId: amount },
  // default: { accountId: amount } } - a period key looks like YYYY-MM, so
  // any key that doesn't match that pattern must be a leftover flat entry
  // from before, which gets folded into "default" (used as a fallback for
  // any month without its own override) rather than losing what was there.
  {
    const rawBudgets = merged.budgets || {};
    const isPeriodKey = (k) => /^\d{4}-\d{2}$/.test(k) || k === "default";
    const looksLegacyFlat = Object.keys(rawBudgets).some(k => !isPeriodKey(k));
    if (looksLegacyFlat) {
      const legacyEntries = Object.fromEntries(Object.entries(rawBudgets).filter(([k]) => !isPeriodKey(k)));
      const alreadyPeriodShaped = Object.fromEntries(Object.entries(rawBudgets).filter(([k]) => isPeriodKey(k)));
      merged.budgets = { ...alreadyPeriodShaped, default: { ...legacyEntries, ...(alreadyPeriodShaped.default || {}) } };
    } else if (!rawBudgets.default) {
      merged.budgets = { ...rawBudgets, default: {} };
    }
  }
  // The curated "in the budget" list is new - for anyone with existing
  // budget figures from before this existed, auto-include every account
  // that already has a nonzero value anywhere, so nothing they'd already
  // set disappears from view.
  if (saved.budgetAccounts === undefined) {
    const inferred = new Set();
    Object.values(merged.budgets).forEach(byAccount => Object.entries(byAccount).forEach(([accountId, amt]) => { if (Number(amt)) inferred.add(accountId); }));
    merged.budgetAccounts = [...inferred];
  }
  merged.inventory = (merged.inventory || []).map(i => ({ nrv: i.salePrice, ...i }));
  merged.payrollRuns = saved.payrollRuns || [];
  merged.recurringJournals = saved.recurringJournals || [];
  merged.reconciliations = saved.reconciliations || [];
  merged.auditLog = saved.auditLog || [];
  if (!merged.settings.userName) merged.settings.userName = "Admin";
  if (!merged.settings.accountingBasis) merged.settings.accountingBasis = "accrual";
  if (!merged.settings.industry) merged.settings.industry = "general";
  if (!merged.settings.country) merged.settings.country = "NG";
  merged.accounts = (merged.accounts || []).map(a => (["5000", "5900"].includes(a.code) && !a.subtype) ? { ...a, subtype: "cogs" } : a);
  // Fixes a sign-convention bug: these two contra accounts were missing the
  // explicit normal-balance override that every other contra account here
  // has (1310, 1330, 1150) - without it, their balances were silently ADDING
  // to revenue/expense totals instead of reducing them.
  merged.accounts = merged.accounts.map(a => a.code === "4200" && !a.normal ? { ...a, normal: "debit" } : a.code === "5900" && !a.normal ? { ...a, normal: "credit" } : a);
  const existingCodes = new Set((merged.accounts || []).map(a => a.code));
  seed.accounts.forEach(a => { if (!existingCodes.has(a.code)) merged.accounts = [...merged.accounts, a]; });
  merged.accounts = merged.accounts.map(normalizeAccount);
  // Repair data saved while the old bank-numbering bug was live: if two
  // account records ended up sharing an id, keep the first and drop clones.
  const seenIds = new Set();
  merged.accounts = merged.accounts.filter(a => { if (seenIds.has(a.id)) return false; seenIds.add(a.id); return true; });
  // Every bank must point at a real, unique account; recreate any missing one.
  merged.banks = (merged.banks || []).map(b => {
    if (merged.accounts.some(a => a.id === b.accountId)) return b;
    const code = nextFreeCode(merged.accounts, 1020, 10);
    merged.accounts = [...merged.accounts, normalizeAccount({ id: code, code, name: b.name, type: "asset", category: "Bank", status: "active" })];
    return { ...b, accountId: code };
  });
  return merged;
}

const INDUSTRIES = [
  { id: "general", name: "General / Other" },
  { id: "retail", name: "Retail & Trading" },
  { id: "services", name: "Professional Services & Consulting" },
  { id: "manufacturing", name: "Manufacturing & Production" },
  { id: "construction", name: "Construction & Real Estate" },
  { id: "hospitality", name: "Hospitality, Restaurants & Food" },
  { id: "healthcare", name: "Healthcare & Medical" },
  { id: "technology", name: "Technology & SaaS" },
  { id: "agriculture", name: "Agriculture & Agribusiness" },
  { id: "nonprofit", name: "Nonprofit & NGO" },
  { id: "logistics", name: "Logistics & Transportation" },
];
const COUNTRIES = [
  { code: "NG", name: "Nigeria", taxName: "VAT" },
  { code: "US", name: "United States", taxName: "Sales Tax" },
  { code: "GB", name: "United Kingdom", taxName: "VAT" },
  { code: "ZA", name: "South Africa", taxName: "VAT" },
  { code: "KE", name: "Kenya", taxName: "VAT" },
  { code: "GH", name: "Ghana", taxName: "VAT" },
  { code: "IN", name: "India", taxName: "GST" },
  { code: "CA", name: "Canada", taxName: "GST/HST" },
  { code: "AU", name: "Australia", taxName: "GST" },
  { code: "AE", name: "United Arab Emirates", taxName: "VAT" },
  { code: "DE", name: "Germany", taxName: "VAT" },
  { code: "OTHER", name: "Other / Not listed", taxName: "Tax" },
];
// Additive suggestions only - never duplicates of the base seeded accounts.
// Each entry: [name, type, category]. Codes are assigned automatically at
// add-time via nextFreeCode so they never collide with existing accounts.
const INDUSTRY_COA_SUGGESTIONS = {
  retail: [
    ["Sales Discounts & Allowances", "revenue", "Operating Income"],
    ["Store Supplies Expense", "expense", "Operating Expenses"],
    ["Point of Sale & Card Processing Fees", "expense", "Operating Expenses"],
    ["Inventory Shrinkage / Stock Loss", "expense", "Other Expenses"],
    ["Gift Card & Store Credit Liability", "liability", "Other Current Liabilities"],
  ],
  services: [
    ["Unbilled Revenue / Work in Progress", "asset", "Current Assets"],
    ["Deferred Revenue", "liability", "Current Liabilities"],
    ["Subcontractor Expense", "expense", "Cost of Goods Sold (COGS)"],
    ["Professional Indemnity Insurance", "expense", "Operating Expenses"],
    ["Software & Subscription Expense", "expense", "Administrative Expenses"],
    ["Reimbursable Client Expenses", "asset", "Current Assets"],
  ],
  manufacturing: [
    ["Raw Materials Inventory", "asset", "Current Assets"],
    ["Work-in-Progress Inventory", "asset", "Current Assets"],
    ["Finished Goods Inventory", "asset", "Current Assets"],
    ["Factory Rent & Utilities", "expense", "Cost of Goods Sold (COGS)"],
    ["Machinery Maintenance & Repairs", "expense", "Cost of Goods Sold (COGS)"],
    ["Scrap & Wastage", "expense", "Other Expenses"],
  ],
  construction: [
    ["Retention Receivable", "asset", "Accounts Receivable"],
    ["Retention Payable", "liability", "Accounts Payable"],
    ["Construction Work-in-Progress", "asset", "Current Assets"],
    ["Subcontractor Costs", "expense", "Cost of Goods Sold (COGS)"],
    ["Equipment Rental", "expense", "Cost of Goods Sold (COGS)"],
    ["Permits & Regulatory Fees", "expense", "Administrative Expenses"],
  ],
  hospitality: [
    ["Food & Beverage Inventory", "asset", "Current Assets"],
    ["Kitchen Supplies Expense", "expense", "Cost of Goods Sold (COGS)"],
    ["Service Charge Payable", "liability", "Other Current Liabilities"],
    ["Tips & Gratuities Payable", "liability", "Other Current Liabilities"],
    ["Laundry & Linen Expense", "expense", "Operating Expenses"],
    ["Licensing & Health Permits", "expense", "Administrative Expenses"],
  ],
  healthcare: [
    ["Medical Supplies Inventory", "asset", "Current Assets"],
    ["Insurance Claims Receivable", "asset", "Accounts Receivable"],
    ["Medical Equipment", "asset", "Fixed Assets"],
    ["Malpractice Insurance", "expense", "Operating Expenses"],
    ["Laboratory & Diagnostic Costs", "expense", "Cost of Goods Sold (COGS)"],
    ["Regulatory & Licensing Fees", "expense", "Administrative Expenses"],
  ],
  technology: [
    ["Deferred Revenue", "liability", "Current Liabilities"],
    ["Capitalized Software Development", "asset", "Intangible Assets"],
    ["Cloud Hosting & Infrastructure", "expense", "Cost of Goods Sold (COGS)"],
    ["Customer Acquisition / Marketing", "expense", "Selling & Distribution Expenses"],
    ["Research & Development", "expense", "Operating Expenses"],
    ["Software Amortization", "expense", "Depreciation & Amortization"],
  ],
  agriculture: [
    ["Livestock Inventory", "asset", "Current Assets"],
    ["Crop Inventory", "asset", "Current Assets"],
    ["Farm Equipment", "asset", "Fixed Assets"],
    ["Seeds, Fertilizer & Agrochemicals", "expense", "Cost of Goods Sold (COGS)"],
    ["Livestock Feed", "expense", "Cost of Goods Sold (COGS)"],
    ["Irrigation & Land Maintenance", "expense", "Operating Expenses"],
  ],
  nonprofit: [
    ["Restricted Fund Balance", "equity", "Reserves"],
    ["Unrestricted Fund Balance", "equity", "Owner's Equity"],
    ["Grant Revenue", "revenue", "Operating Income"],
    ["Donations & Contributions", "revenue", "Operating Income"],
    ["Program Expenses", "expense", "Operating Expenses"],
    ["Fundraising Expense", "expense", "Selling & Distribution Expenses"],
  ],
  logistics: [
    ["Fleet Vehicles", "asset", "Fixed Assets"],
    ["Fuel Expense", "expense", "Cost of Goods Sold (COGS)"],
    ["Vehicle Maintenance", "expense", "Cost of Goods Sold (COGS)"],
    ["Driver Wages", "expense", "Cost of Goods Sold (COGS)"],
    ["Toll & Logistics Fees", "expense", "Operating Expenses"],
    ["Freight Revenue", "revenue", "Operating Income"],
  ],
  general: [],
};
// Statutory payable accounts commonly required by each country's payroll and
// tax regime. Note: only Nigeria's PAYE calculation is automated in Payroll
// today - these accounts for other countries give you somewhere correct to
// post to, but you compute and enter those figures yourself for now.
const COUNTRY_COA_SUGGESTIONS = {
  NG: [["Industrial Training Fund (ITF) Payable", "liability", "Other Current Liabilities"], ["NSITF Payable", "liability", "Other Current Liabilities"]],
  US: [["Federal Income Tax Withholding Payable", "liability", "Other Current Liabilities"], ["FICA / Social Security Payable", "liability", "Other Current Liabilities"], ["State Unemployment Tax Payable", "liability", "Other Current Liabilities"]],
  GB: [["PAYE / NIC Payable", "liability", "Other Current Liabilities"], ["Corporation Tax Payable", "liability", "Other Current Liabilities"]],
  ZA: [["PAYE Payable", "liability", "Other Current Liabilities"], ["UIF Payable", "liability", "Other Current Liabilities"], ["SDL Payable", "liability", "Other Current Liabilities"]],
  KE: [["PAYE Payable", "liability", "Other Current Liabilities"], ["NHIF Payable", "liability", "Other Current Liabilities"], ["NSSF Payable", "liability", "Other Current Liabilities"]],
  GH: [["PAYE Payable", "liability", "Other Current Liabilities"], ["SSNIT Payable", "liability", "Other Current Liabilities"]],
  IN: [["GST Payable", "liability", "Other Current Liabilities"], ["TDS Payable", "liability", "Other Current Liabilities"], ["Provident Fund Payable", "liability", "Other Current Liabilities"]],
  CA: [["GST/HST Payable", "liability", "Other Current Liabilities"], ["CPP Payable", "liability", "Other Current Liabilities"], ["EI Payable", "liability", "Other Current Liabilities"]],
  AU: [["GST Payable", "liability", "Other Current Liabilities"], ["PAYG Withholding Payable", "liability", "Other Current Liabilities"], ["Superannuation Payable", "liability", "Other Current Liabilities"]],
  AE: [["VAT Payable", "liability", "Other Current Liabilities"], ["WPS Payable", "liability", "Other Current Liabilities"]],
  DE: [["Umsatzsteuer (VAT) Payable", "liability", "Other Current Liabilities"], ["Lohnsteuer Payable", "liability", "Other Current Liabilities"], ["Sozialversicherung Payable", "liability", "Other Current Liabilities"]],
  OTHER: [],
};
const TAX_PRESETS = [
  { name: "VAT", rate: 7.5, mode: "percent", effect: "deduct", accountId: "2100" },
  { name: "WHT", rate: 5, mode: "percent", effect: "deduct", accountId: "2200" },
  { name: "WHT (10%)", rate: 10, mode: "percent", effect: "deduct", accountId: "2200" },
  { name: "Stamp Duty", rate: 50, mode: "fixed", effect: "deduct", accountId: "2300" },
];
const CATEGORY_RULES = [
  { re: /rent|lease|property/i, accountId: "5100" },
  { re: /electric|water|utilit|gas co|internet|isp/i, accountId: "5200" },
  { re: /payroll|salary|gusto|adp|wages/i, accountId: "5300" },
  { re: /staples|office|depot|supplies/i, accountId: "5400" },
  { re: /ads|marketing|meta|google ads|facebook/i, accountId: "5500" },
  { re: /bank fee|service charge|nsf|overdraft/i, accountId: "5600" },
  { re: /invoice|client payment|deposit|revenue|sale/i, accountId: "4000" },
];
function guessCategory(desc) { const hit = CATEGORY_RULES.find((r) => r.re.test(desc || "")); return hit ? hit.accountId : null; }

/* ---------------- cascading tax engine ---------------- */
function applyCascadingTaxes(subtotal, taxes) {
  let running = Number(subtotal) || 0;
  const steps = (taxes || []).map((t) => {
    const before = running;
    const amt = t.mode === "fixed" ? Number(t.rate) || 0 : before * ((Number(t.rate) || 0) / 100);
    running = t.effect === "deduct" ? running - amt : running + amt;
    return { ...t, amount: amt, before, after: running };
  });
  return { steps, finalAmount: running, subtotal: Number(subtotal) || 0 };
}
// docType: 'invoice' -> deduct=debit(recoverable/withheld), add=credit(payable)
//          'bill'    -> deduct=credit(withheld payable to authority), add=debit(recoverable input tax)
function taxJournalLines(steps, docType) {
  return steps.filter(s => Math.abs(s.amount) > 0.004).map((s) => {
    const isDebit = docType === "invoice" ? s.effect === "deduct" : s.effect === "add";
    return { accountId: s.accountId, debit: isDebit ? s.amount : 0, credit: isDebit ? 0 : s.amount };
  });
}
// Group tax: two or more taxes calculated in parallel on the SAME base amount
// (e.g. VAT and WHT both computed on 500,000, then combined) rather than
// cascading on top of each other's running balance.
function computeGroupForBase(base, group) {
  const components = (group?.components || []).map((c) => {
    const amt = c.mode === "fixed" ? Number(c.rate) || 0 : (Number(base) || 0) * ((Number(c.rate) || 0) / 100);
    return { ...c, amount: amt };
  });
  const netAdj = components.reduce((s, c) => s + (c.effect === "deduct" ? -c.amount : c.amount), 0);
  return { components, netAdj };
}
// Full document total: line items (each optionally taxed by its own group,
// in parallel) -> subtotal after line taxes -> document-level cascading
// taxes (e.g. a stamp duty applied to the whole invoice) -> final amount.
function computeDocTotals(doc, taxGroups) {
  const lineCalcs = (doc.items || []).map((it) => {
    const base = (Number(it.qty) || 0) * (Number(it.price) || 0);
    const group = it.taxComponents ? { components: it.taxComponents } : (taxGroups || []).find((g) => g.id === it.taxGroupId);
    const { components, netAdj } = computeGroupForBase(base, group);
    return { item: it, base, components, netAdj, lineTotal: base + netAdj };
  });
  const subtotal = lineCalcs.reduce((s, l) => s + l.base, 0);
  const lineTaxAdj = lineCalcs.reduce((s, l) => s + l.netAdj, 0);
  const afterLineTax = subtotal + lineTaxAdj;
  const cascade = applyCascadingTaxes(afterLineTax, doc.taxes || []);
  return { lineCalcs, subtotal, lineTaxAdj, afterLineTax, cascadeSteps: cascade.steps, finalAmount: cascade.finalAmount };
}

/* ---------------- ledger engine ---------------- */
function acctNormal(acc) { return acc.normal ? acc.normal === "debit" : DEBIT_NORMAL[acc.type]; }
function computeBalances(accounts, transactions, upTo = null) {
  const bal = {}; accounts.forEach((a) => (bal[a.id] = 0));
  transactions.forEach((t) => {
    if (upTo && t.date > upTo) return;
    t.lines.forEach((l) => {
      const acc = accounts.find((a) => a.id === l.accountId); if (!acc) return;
      bal[acc.id] += acctNormal(acc) ? (l.debit - l.credit) : (l.credit - l.debit);
    });
  });
  return bal;
}
function periodMovement(accounts, transactions, from, to) {
  const bal = {}; accounts.forEach((a) => (bal[a.id] = 0));
  transactions.forEach((t) => {
    if (t.date < from || t.date > to) return;
    t.lines.forEach((l) => {
      const acc = accounts.find((a) => a.id === l.accountId); if (!acc) return;
      bal[acc.id] += acctNormal(acc) ? (l.debit - l.credit) : (l.credit - l.debit);
    });
  });
  return bal;
}
function sumByType(accounts, balances, type, subtype = null) {
  return accounts.filter(a => a.type === type && (!subtype || a.subtype === subtype))
    .reduce((s, a) => s + (a.contra ? -balances[a.id] : balances[a.id]), 0);
}
// Consume oldest lots first (FIFO) for a given item - used only for the
// informational FIFO costing report; the general ledger itself uses
// weighted-average costing (simpler, and what most small businesses run).
function consumeFIFO(lots, itemId, qty) {
  let remaining = qty; const consumed = []; let fifoCost = 0;
  const nextLots = lots.map(l => ({ ...l }));
  nextLots.filter(l => l.itemId === itemId && l.remainingQty > 0).sort((a, b) => a.date.localeCompare(b.date)).forEach(l => {
    if (remaining <= 0) return;
    const take = Math.min(l.remainingQty, remaining);
    l.remainingQty -= take; remaining -= take; fifoCost += take * l.unitCost;
    consumed.push({ lotId: l.id, qty: take, unitCost: l.unitCost });
  });
  return { lots: nextLots, consumed, fifoCost };
}
// Reverse of consumeFIFO - puts qty back into the newest partially-consumed
// lots first (last-consumed, first-restored), used when an invoice is edited.
function restoreFIFO(lots, itemId, qty) {
  let remaining = qty;
  const nextLots = lots.map(l => ({ ...l }));
  nextLots.filter(l => l.itemId === itemId && l.remainingQty < l.qty).sort((a, b) => b.date.localeCompare(a.date)).forEach(l => {
    if (remaining <= 0) return;
    const give = Math.min(l.qty - l.remainingQty, remaining);
    l.remainingQty += give; remaining -= give;
  });
  return nextLots;
}
// Remove the journal entry a document originally posted. Newer entries carry
// docId; older ones are matched by memo + source, deliberately leaving
// payment/refund entries (which mention the doc id in their memos) untouched.
// Delete a single journal entry with the right approach per source:
// documents point back to their source, feed lines are uncategorized instead,
// depreciation rolls the asset register back (when it carries per-asset meta).
function deleteTransactionFromData(d, txn, notify) {
  if (txn.source === "invoice" || txn.source === "bill") { notify("This entry belongs to a document - delete the invoice/bill itself"); return d; }
  if (txn.source === "fixed-asset") { notify("This entry belongs to a fixed asset - dispose of the asset in the Fixed Asset Register"); return d; }
  let next = { ...d, transactions: d.transactions.filter(t => t.id !== txn.id) };
  if (txn.feedId) { // bank-import: return the feed line to uncategorized
    next = { ...next, bankFeed: next.bankFeed.map(f => f.id === txn.feedId ? { ...f, status: "uncategorized", txnId: null, accountId: null } : f) };
  }
  if (txn.source === "depreciation" && txn.meta?.perAsset) { // roll the register back
    next = { ...next, fixedAssets: next.fixedAssets.map(a => { const m = txn.meta.perAsset.find(x => x.id === a.id); return m ? { ...a, accumulatedDepreciation: Math.max(0, a.accumulatedDepreciation - m.amt), lastDepreciationMonth: null } : a; }) };
  }
  if (txn.source === "payment" || txn.source === "refund") { // roll back document paid amounts via the linked payment record
    const pay = d.payments.find(p => p.txnId === txn.id);
    if (pay) {
      next = { ...next,
        payments: next.payments.filter(p => p.id !== pay.id),
        invoices: pay.relatedType === "invoice" ? next.invoices.map(i => i.id === pay.relatedId ? { ...i, amountPaid: Math.max(0, (i.amountPaid || 0) - pay.amount) } : i) : next.invoices,
        bills: pay.relatedType === "bill" ? next.bills.map(b => b.id === pay.relatedId ? { ...b, amountPaid: Math.max(0, (b.amountPaid || 0) - pay.amount) } : b) : next.bills,
      };
    } else notify("Entry deleted - note: this older payment had no link to its record, so any invoice/bill paid amount was not adjusted");
  }
  return next;
}
function stripDocTransactions(transactions, id) {
  return transactions.filter(t => !(t.docId === id || ((t.source === "invoice" || t.source === "bill") && (t.memo || "").includes(id))));
}
function netIncomeAllTime(data, upTo = null) { const bal = computeBalances(data.accounts, data.transactions, upTo); return sumByType(data.accounts, bal, "revenue") - sumByType(data.accounts, bal, "expense"); }
// Build a balanced transaction object, or null if debits ≠ credits.
function buildTxn(memo, date, lines, source = "manual", docId = null) {
  const debit = lines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
  const credit = lines.reduce((s, l) => s + (Number(l.credit) || 0), 0);
  if (Math.round(debit * 100) !== Math.round(credit * 100)) return null;
  return { id: uid("txn"), date, memo, source, docId, lines };
}
function agingBucket(days) {
  if (days <= 0) return "Current"; if (days <= 30) return "1-30 days"; if (days <= 60) return "31-60 days"; if (days <= 90) return "61-90 days"; return "90+ days";
}
const AGING_BUCKETS = ["Current", "1-30 days", "31-60 days", "61-90 days", "90+ days"];

/* =================================== App root =================================== */
// Compares two arrays of {id,...} records and reports what changed.
function diffCollection(prevArr, nextArr) {
  const a = prevArr || [], b = nextArr || [];
  if (a === b) return null;
  const prevMap = new Map(a.map(x => [x.id, x]));
  const nextMap = new Map(b.map(x => [x.id, x]));
  const added = [...nextMap.keys()].filter(id => !prevMap.has(id));
  const removed = [...prevMap.keys()].filter(id => !nextMap.has(id));
  const changed = [...nextMap.keys()].filter(id => prevMap.has(id) && JSON.stringify(prevMap.get(id)) !== JSON.stringify(nextMap.get(id)));
  if (added.length === 0 && removed.length === 0 && changed.length === 0) return null;
  return { added, removed, changed, prevMap, nextMap };
}
// Produces one plain-English line per meaningful change between two data
// payloads. This is the sole source of audit-log entries - see the setData
// wrapper in App, which calls this on every state update.
function describeMutations(prev, next) {
  const out = [];
  const txnDiff = diffCollection(prev.transactions, next.transactions);
  if (txnDiff) {
    txnDiff.added.forEach(id => { const t = txnDiff.nextMap.get(id); const amt = t.lines.reduce((s, l) => s + l.debit, 0); out.push(`Posted journal entry "${t.memo}" (${t.source}, ${amt.toLocaleString()})`); });
    txnDiff.removed.forEach(id => { const t = txnDiff.prevMap.get(id); out.push(`Deleted journal entry "${t.memo}" (${t.source})`); });
    txnDiff.changed.forEach(id => { const t = txnDiff.nextMap.get(id); out.push(`Edited journal entry "${t.memo}"`); });
  }
  const collections = [
    ["invoices", "invoice", x => x.id], ["bills", "bill", x => x.id],
    ["accounts", "account", x => `${x.code} ${x.name}`], ["banks", "bank", x => x.name],
    ["employees", "employee", x => x.name], ["payrollRuns", "payroll run", x => x.month],
    ["recurringJournals", "recurring journal", x => x.memo], ["projects", "project", x => x.name],
    ["inventory", "inventory item", x => x.name], ["fixedAssets", "fixed asset", x => x.name],
    ["taxGroups", "tax group", x => x.name], ["categoryRules", "bank rule", x => x.keyword],
    ["expenses", "expense", x => x.vendor], ["payments", "payment", x => x.relatedId || x.refundTo || "payment"],
    ["reconciliations", "reconciliation", x => x.bankId],
  ];
  collections.forEach(([key, label, nameFn]) => {
    const diff = diffCollection(prev[key], next[key]);
    if (!diff) return;
    diff.added.forEach(id => out.push(`Created ${label} - ${nameFn(diff.nextMap.get(id))}`));
    diff.removed.forEach(id => out.push(`Deleted ${label} - ${nameFn(diff.prevMap.get(id))}`));
    diff.changed.forEach(id => out.push(`Updated ${label} - ${nameFn(diff.nextMap.get(id))}`));
  });
  const feedDiff = diffCollection(prev.bankFeed, next.bankFeed);
  if (feedDiff) {
    feedDiff.changed.forEach(id => {
      const before = feedDiff.prevMap.get(id), after = feedDiff.nextMap.get(id);
      if (before.status !== after.status) out.push(`Bank line "${after.desc}" ${after.status === "categorized" ? "categorized" : "uncategorized"}`);
    });
    if (feedDiff.added.length) out.push(`Imported ${feedDiff.added.length} bank statement line(s)`);
    if (feedDiff.removed.length) out.push(`Removed ${feedDiff.removed.length} bank feed line(s) to bin`);
  }
  if (JSON.stringify(prev.settings) !== JSON.stringify(next.settings)) {
    if (prev.settings.accountingBasis !== next.settings.accountingBasis) out.push(`Changed accounting basis to ${next.settings.accountingBasis}`);
    else if (prev.settings.companyName !== next.settings.companyName) out.push(`Renamed company to "${next.settings.companyName}"`);
    else if (prev.settings.logo !== next.settings.logo) out.push(next.settings.logo ? "Uploaded company logo" : "Removed company logo");
    else out.push("Updated settings");
  }
  return out;
}
// Shows every transaction touching one or more accounts, within a date range
// or up to a point-in-time date - the click-through behind any report line.
// Groups by transaction (not just the matching line) so the full double
// entry is visible, not just the half that happened to touch this account.
function DrilldownModal({ data, spec, onClose }) {
  const { styles, fmt, theme } = useUI();
  const { accountIds, label, range, asOf } = spec;
  const idSet = new Set(accountIds);
  const inRange = (date) => asOf ? date <= asOf : range ? (date >= range.from && date <= range.to) : true;
  const matches = data.transactions.filter(t => inRange(t.date) && t.lines.some(l => idSet.has(l.accountId)));
  const sorted = [...matches].sort((a, b) => b.date.localeCompare(a.date));
  // Sign each line the same way the report itself did: by that account's own
  // normal balance (debit- or credit-increasing), flipped for contra
  // accounts - a raw debit-minus-credit would show revenue drill-downs as
  // negative, since revenue is credit-normal.
  const signedAmount = (t) => t.lines.filter(l => idSet.has(l.accountId)).reduce((s, l) => {
    const acc = data.accounts.find(a => a.id === l.accountId);
    const raw = acctNormal(acc) ? l.debit - l.credit : l.credit - l.debit;
    return s + (acc && acc.contra ? -raw : raw);
  }, 0);
  const total = matches.reduce((s, t) => s + signedAmount(t), 0);
  const accountNames = accountIds.map(id => { const a = data.accounts.find(x => x.id === id); return a ? `${a.code} ${a.name}` : id; });

  return (
    <div className="no-print" style={{ position: "fixed", inset: 0, zIndex: 100, background: "rgba(11,23,48,0.55)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div style={{ background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 12, padding: "22px 24px", maxWidth: 720, width: "100%", maxHeight: "80vh", overflowY: "auto", boxShadow: "0 16px 48px rgba(16,24,40,0.3)" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
          <div>
            <div style={{ fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontSize: 17, fontWeight: 600, color: theme.text }}>{label}</div>
            <div style={{ fontSize: 11.5, color: theme.muted, marginTop: 2 }}>{accountNames.join(", ")}{asOf ? ` \u00b7 up to ${fmtDate(asOf)}` : range ? ` \u00b7 ${fmtDate(range.from)} to ${fmtDate(range.to)}` : ""}</div>
          </div>
          <button style={styles.btnGhost} onClick={onClose}>Close</button>
        </div>
        <div style={{ marginTop: 14 }}>
          {sorted.length === 0 ? <div style={{ fontSize: 13, color: theme.muted, padding: "16px 0" }}>No transactions in this period.</div> : (
            <table style={styles.table}>
              <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={styles.th}>Source</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
              <tbody>{sorted.map(t => {
                const amt = signedAmount(t);
                return (
                  <tr key={t.id}>
                    <td style={styles.tdMono}>{fmtDate(t.date)}</td>
                    <td style={styles.td}>{t.memo}</td>
                    <td style={{ ...styles.td, fontSize: 12, color: theme.muted }}>{t.source}</td>
                    <td style={{ ...styles.tdMono, textAlign: "right", color: amt >= 0 ? theme.text : theme.rose }}>{fmt(amt)}</td>
                  </tr>
                );
              })}</tbody>
              <tfoot><tr><td colSpan={3} style={{ ...styles.td, fontWeight: 700 }}>Total</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(total)}</td></tr></tfoot>
            </table>
          )}
        </div>
        <div style={{ fontSize: 11, color: theme.muted, marginTop: 10 }}>Amounts follow each account's own normal balance (a positive figure is an increase). {matches.length} transaction(s) shown - open the Journal tab for a full ledger view.</div>
      </div>
    </div>
  );
}
export default function App({ companyData, onCompanyDataChange, onSignOut }) {
  const [data, rawSetData] = useState(null);
  // Every mutation to `data` goes through this wrapper, which diffs the
  // before/after state and appends one audit-log entry describing what
  // changed, who did it, and when. Because it sits at the single point all
  // state updates pass through, no action anywhere in the app can bypass
  // being logged - this is deliberately structural rather than something
  // sprinkled per-button, so it can't be missed on a new feature later.
  const setData = useCallback((updater) => {
    rawSetData((prev) => {
      if (!prev) return typeof updater === "function" ? updater(prev) : updater;
      const next = typeof updater === "function" ? updater(prev) : updater;
      if (next === prev) return next;
      const who = (next.settings || prev.settings)?.userName || "Admin";
      if (next.__auditLabel) {
        const { __auditLabel, ...clean } = next;
        return { ...clean, auditLog: [...(clean.auditLog || prev.auditLog || []), { id: uid("log"), time: new Date().toISOString(), user: who, action: __auditLabel }].slice(-2000) };
      }
      const actions = describeMutations(prev, next);
      if (actions.length === 0) return next;
      const entries = actions.map(action => ({ id: uid("log"), time: new Date().toISOString(), user: who, action }));
      return { ...next, auditLog: [...(next.auditLog || prev.auditLog || []), ...entries].slice(-2000) };
    });
  }, []);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [toast, setToast] = useState(null);

  useEffect(() => { (async () => {
    try {
      const loaded = migrateData(companyData);
      const { data: withRecurring, posted } = processRecurringJournals(loaded);
      setData(withRecurring);
      if (posted > 0) setTimeout(() => setToast(`${posted} recurring journal entr${posted === 1 ? "y" : "ies"} posted automatically`), 600);
    }
    catch (e) { console.error("Failed to load company data:", e); }
    setLoading(false);
  })(); }, []);

  useEffect(() => {
    if (!data) return;
    const t = setTimeout(() => { try { onCompanyDataChange(data); } catch {} }, 400);
    return () => clearTimeout(t);
  }, [data]);

  const notify = useCallback((msg) => { setToast(msg); setTimeout(() => setToast(null), 2800); }, []);

  const postTransaction = useCallback((memo, date, lines, source = "manual", docId = null) => {
    const txn = buildTxn(memo, date, lines, source, docId);
    if (!txn) { notify("Entry not balanced - debits must equal credits"); return false; }
    setData((d) => ({ ...d, transactions: [...d.transactions, txn] }));
    return true;
  }, [notify]);

  const theme = useMemo(() => buildTheme(data?.settings || defaultSettings()), [data?.settings]);
  const currency = useMemo(() => {
    const c = CURRENCIES.find(c => c.code === data?.settings?.currencyCode) || CURRENCIES[0];
    return { ...c, symbol: data?.settings?.currencySymbol || c.symbol };
  }, [data?.settings]);
  const fmt = useCallback((n) => {
    const v = Number(n) || 0;
    return (v < 0 ? "-" : "") + currency.symbol + Math.abs(v).toLocaleString(currency.locale, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }, [currency]);
  const styles = useMemo(() => makeStyles(theme), [theme]);
  // In-app confirmation dialog. window.confirm is blocked inside the sandboxed
  // artifact iframe (it silently returns false), so destructive actions use
  // this promise-based modal instead.
  const [confirmState, setConfirmState] = useState(null);
  const [drilldown, setDrilldown] = useState(null);
  // Drives the cursor-following sheen + tilt on .ces-card/.ces-bubble
  // surfaces and the specular highlight + tilt on .ces-logo badges.
  // Throttled to one update per animation frame (not one per raw mousemove
  // event, which fires far more often than the screen can repaint) - that
  // over-firing was the main reason the tilt looked blurry before, since
  // the element was constantly being nudged to a new angle and never
  // settling into a frame the browser could rasterize cleanly.
  useEffect(() => {
    let pending = null;
    const onMove = (e) => {
      if (pending) return;
      pending = requestAnimationFrame(() => {
        pending = null;
        const card = e.target.closest && e.target.closest(".ces-card");
        if (card) {
          const r = card.getBoundingClientRect();
          const px = (e.clientX - r.left) / r.width, py = (e.clientY - r.top) / r.height;
          card.style.setProperty("--mx", `${e.clientX - r.left}px`);
          card.style.setProperty("--my", `${e.clientY - r.top}px`);
          card.style.setProperty("--rx", `${(0.5 - py) * 2.5}deg`);
          card.style.setProperty("--ry", `${(px - 0.5) * 2.5}deg`);
        }
        const logo = e.target.closest && e.target.closest(".ces-logo");
        if (logo) {
          const r = logo.getBoundingClientRect();
          const px = (e.clientX - r.left) / r.width, py = (e.clientY - r.top) / r.height;
          logo.style.setProperty("--rx", `${(0.5 - py) * 4}deg`);
          logo.style.setProperty("--ry", `${(px - 0.5) * 4}deg`);
          logo.style.setProperty("--lx", `${px * 100}%`);
          logo.style.setProperty("--ly", `${py * 100}%`);
        }
      });
    };
    document.addEventListener("mousemove", onMove);
    return () => { document.removeEventListener("mousemove", onMove); if (pending) cancelAnimationFrame(pending); };
  }, []);
  const confirm = useCallback((message) => new Promise(resolve => setConfirmState({ message, resolve })), []);
  const answerConfirm = (ok) => { confirmState?.resolve(ok); setConfirmState(null); };
  // Opens a drill-down showing every ledger line for one or more accounts
  // within a range (or up to a point-in-time date) - the click-through
  // behind every report line, so "why is this number what it is" is always
  // one click away.
  const openDrilldown = useCallback((spec) => setDrilldown(spec), []);
  const ui = useMemo(() => ({ theme, styles, fmt, currency, confirm, openDrilldown }), [theme, styles, fmt, currency, confirm, openDrilldown]);

  if (loading || !data) {
    return <div style={{ minHeight: 640, display: "flex", alignItems: "center", justifyContent: "center", background: "#0B1730", color: "#8B97B8", fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em" }}>
      <style>{FONT_IMPORT}</style>Opening the books…
    </div>;
  }

  const balances = computeBalances(data.accounts, data.transactions);

  return (
    <UIContext.Provider value={ui}>
      <div style={styles.app}>
        <style>{FONT_IMPORT + globalCss(theme)}</style>
        <Sidebar tab={tab} setTab={setTab} data={data} />
        <main style={styles.main}>
          {tab === "dashboard" && <Dashboard data={data} balances={balances} />}
          {tab === "accounts" && <ChartOfAccounts data={data} balances={balances} setData={setData} notify={notify} />}
          {tab === "journal" && <Journal data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "banks" && <Banks data={data} setData={setData} balances={balances} notify={notify} postTransaction={postTransaction} />}
          {tab === "invoices" && <Invoices data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "sales-orders" && <SalesOrders data={data} setData={setData} notify={notify} />}
          {tab === "sales-receipts" && <SalesReceipts data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "credit-notes" && <CreditNotes data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "sales-returns" && <SalesReturns data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "bills" && <Bills data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "payments" && <Payments data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "expenses" && <Expenses data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "recurring-expenses" && <RecurringExpenses data={data} setData={setData} notify={notify} />}
          {tab === "purchase-orders" && <PurchaseOrders data={data} setData={setData} notify={notify} />}
          {tab === "purchase-receipts" && <PurchaseReceipts data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "vendor-credits" && <VendorCredits data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "inventory" && <Inventory data={data} setData={setData} notify={notify} />}
          {tab === "assets" && <FixedAssets data={data} setData={setData} postTransaction={postTransaction} notify={notify} />}
          {tab === "timesheet" && <TimeSheet data={data} setData={setData} notify={notify} />}
          {tab === "reports" && <Reports data={data} balances={balances} setData={setData} notify={notify} />}
          {tab === "forecast" && <Forecast data={data} balances={balances} />}
          {tab === "customers" && <Customers data={data} setData={setData} notify={notify} />}
          {tab === "vendors" && <Vendors data={data} />}
          {tab === "payroll" && <Payroll data={data} setData={setData} notify={notify} />}
          {tab === "projects" && <ProjectsPage data={data} setData={setData} notify={notify} />}
          {tab === "tax" && <TaxPage data={data} setData={setData} notify={notify} />}
          {tab === "provisions" && <ProvisionsPage data={data} setData={setData} notify={notify} />}
          {tab === "budgets" && <BudgetsPage data={data} setData={setData} notify={notify} />}
          {tab === "bulk-update" && <BulkUpdate data={data} setData={setData} notify={notify} />}
          {tab === "transaction-locking" && <TransactionLocking data={data} setData={setData} notify={notify} />}
          {tab === "presentation" && <Presentation data={data} balances={balances} />}
          {tab === "audit" && <AuditTrail data={data} />}
          {tab === "settings" && <Settings data={data} setData={setData} notify={notify} />}
        </main>
        {toast && <div className="no-print" style={styles.toast}>{toast}</div>}
        {drilldown && <DrilldownModal data={data} spec={drilldown} onClose={() => setDrilldown(null)} />}
        {confirmState && (
          <div className="no-print" style={{ position: "fixed", inset: 0, zIndex: 100, background: "rgba(11,23,48,0.55)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
            <div style={{ background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 12, padding: "22px 24px", maxWidth: 460, width: "100%", boxShadow: "0 16px 48px rgba(16,24,40,0.3)" }}>
              <div style={{ fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontSize: 17, fontWeight: 600, color: theme.text }}>Please confirm</div>
              <div style={{ fontSize: 13.5, color: theme.text, marginTop: 10, lineHeight: 1.55 }}>{confirmState.message}</div>
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 18 }}>
                <button style={styles.btnGhost} onClick={() => answerConfirm(false)}>Cancel</button>
                <button style={{ ...styles.btnPrimary, background: theme.rose, borderColor: theme.rose }} onClick={() => answerConfirm(true)}>Yes, continue</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </UIContext.Provider>
  );
}

/* ---------------- nav & layout ---------------- */
const NAV = [
  { id: "dashboard", label: "Dashboard", icon: "◈", group: "Overview" },
  { id: "banks", label: "Banking", icon: "▣", group: "Money" },
  { id: "payments", label: "Payments", icon: "⇄", group: "Money" },
  { id: "customers", label: "Customers", icon: "◉", group: "Sales" },
  { id: "invoices", label: "Invoicing", icon: "▧", group: "Sales" },
  { id: "sales-orders", label: "Sales Orders", icon: "▤", group: "Sales" },
  { id: "sales-receipts", label: "Sales Receipts", icon: "▥", group: "Sales" },
  { id: "credit-notes", label: "Credit Notes", icon: "◇", group: "Sales" },
  { id: "sales-returns", label: "Sales Returns", icon: "\u21a9", group: "Sales" },
  { id: "vendors", label: "Vendors", icon: "◎", group: "Purchases" },
  { id: "bills", label: "Bills", icon: "▨", group: "Purchases" },
  { id: "expenses", label: "Expenses", icon: "▽", group: "Purchases" },
  { id: "recurring-expenses", label: "Recurring Expenses", icon: "↻", group: "Purchases" },
  { id: "purchase-orders", label: "Purchase Orders", icon: "▤", group: "Purchases" },
  { id: "purchase-receipts", label: "Purchase Receipts", icon: "▥", group: "Purchases" },
  { id: "vendor-credits", label: "Vendor Credits", icon: "◇", group: "Purchases" },
  { id: "accounts", label: "Chart of Accounts", icon: "▤", group: "Accounting" },
  { id: "journal", label: "Journal Entries", icon: "≡", group: "Accounting" },
  { id: "tax", label: "Tax", icon: "§", group: "Accounting" },
  { id: "provisions", label: "Provisions", icon: "\u26a0", group: "Accounting" },
  { id: "budgets", label: "Budgets", icon: "\ud83d\udcca", group: "Accounting" },
  { id: "bulk-update", label: "Bulk Update", icon: "\u2261", group: "Accounting" },
  { id: "transaction-locking", label: "Transaction Locking", icon: "\ud83d\udd12", group: "Accounting" },
  { id: "payroll", label: "Payroll", icon: "◫", group: "Operations" },
  { id: "projects", label: "Projects", icon: "◪", group: "Operations" },
  { id: "inventory", label: "Inventory", icon: "▥", group: "Operations" },
  { id: "assets", label: "Fixed Assets", icon: "▦", group: "Operations" },
  { id: "timesheet", label: "Time Sheet", icon: "\u23f1", group: "Operations" },
  { id: "reports", label: "Reports", icon: "◧", group: "Insights" },
  { id: "forecast", label: "Forecast", icon: "∿", group: "Insights" },
  { id: "presentation", label: "Presentation", icon: "▶", group: "Insights" },
  { id: "audit", label: "Audit Trail", icon: "◉", group: "System" },
  { id: "settings", label: "Settings", icon: "⚙", group: "System" },
];
function Sidebar({ tab, setTab, data }) {
  const { theme, styles } = useUI();
  const [collapsed, setCollapsed] = useState(false);
  const [closedGroups, setClosedGroups] = useState(new Set());
  const toggleGroup = (g) => setClosedGroups(prev => { const s = new Set(prev); s.has(g) ? s.delete(g) : s.add(g); return s; });
  const groups = [...new Set(NAV.map(n => n.group))];
  return (
    <aside className="no-print" style={{ ...styles.sidebar, ...(collapsed ? { width: 62, padding: "20px 8px" } : {}) }}>
      <div style={{ ...styles.brand, ...(collapsed ? { justifyContent: "center" } : {}) }}>
        {data.settings.logo
          ? <div className="ces-logo" style={{ width: 34, height: 34, padding: 3, flexShrink: 0 }}><img src={data.settings.logo} alt="logo" style={{ width: "100%", height: "100%", objectFit: "contain" }} /></div>
          : <div style={styles.brandMark}>CE</div>}
        {!collapsed && <div>
          <div style={styles.brandName}>{data.settings.companyName}</div>
          <div style={styles.brandSub}>books, modeled</div>
        </div>}
      </div>
      <button className="nav-item" onClick={() => setCollapsed(c => !c)} title={collapsed ? "Expand sidebar" : "Collapse sidebar"}
        style={{ ...styles.navItem, marginTop: 14, justifyContent: collapsed ? "center" : "flex-start", color: theme.sidebarMuted }}>
        <span style={{ fontSize: 13 }}>{collapsed ? "»" : "«"}</span>{!collapsed && <span style={{ fontSize: 12 }}>Collapse</span>}
      </button>
      <nav style={{ marginTop: 8, flex: 1, overflowY: "auto" }}>
        {groups.map(g => {
          const items = NAV.filter(n => n.group === g);
          const closed = closedGroups.has(g);
          const groupHasActive = items.some(n => n.id === tab);
          return (
            <React.Fragment key={g}>
              {!collapsed && (
                <button onClick={() => toggleGroup(g)} style={{ ...styles.navGroup, display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%", background: "none", border: "none", cursor: "pointer", textAlign: "left" }}>
                  <span>{g}</span>
                  <span style={{ fontSize: 9, opacity: 0.8 }}>{closed ? "▸" : "▾"}</span>
                </button>
              )}
              {(collapsed || !closed) && items.map(n => {
                const active = tab === n.id;
                return (
                  <button key={n.id} className="nav-item" onClick={() => setTab(n.id)} title={n.label}
                    style={{ ...styles.navItem, ...(active ? styles.navItemActive : {}), ...(collapsed ? { justifyContent: "center", padding: "9px 0" } : {}) }}>
                    {!collapsed && <span style={{ width: 3, height: 16, borderRadius: 2, background: active ? theme.accent : "transparent", marginRight: 2 }} />}
                    <span style={{ opacity: active ? 1 : 0.6, width: 16, display: "inline-block", fontSize: collapsed ? 14 : 12, textAlign: "center" }}>{n.icon}</span>
                    {!collapsed && n.label}
                  </button>
                );
              })}
              {!collapsed && closed && groupHasActive && <div style={{ fontSize: 10, color: theme.accent, padding: "0 10px 4px" }}>· current page in this group</div>}
            </React.Fragment>
          );
        })}
      </nav>
      {!collapsed && (
        <div style={styles.sidebarFoot}>
          <button className="nav-item" onClick={onSignOut} style={{ ...styles.navItem, width: "100%", justifyContent: "flex-start", color: theme.sidebarMuted }}>
            <span style={{ fontSize: 13 }}>⏻</span><span style={{ fontSize: 12 }}>Sign out</span>
          </button>
        </div>
      )}
    </aside>
  );
}
function PageHeader({ eyebrow, title, sub, action }) {
  const { theme } = useUI();
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 22, flexWrap: "wrap", gap: 12 }}>
      <div>
        <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase", color: theme.muted }}>{eyebrow}</div>
        <div style={{ fontFamily: "Inter, sans-serif", letterSpacing: "-0.02em", fontSize: 24, fontWeight: 700, color: theme.text, marginTop: 2 }}>{title}</div>
        {sub && <div style={{ fontSize: 13, color: theme.muted, marginTop: 2 }}>{sub}</div>}
      </div>
      {action}
    </div>
  );
}
function Kpi({ label, value, tone }) {
  const { theme, styles } = useUI();
  const c = tone === "emerald" ? theme.emerald : tone === "rose" ? theme.rose : tone === "amber" ? theme.amber : theme.accent;
  return (
    <div className="ces-card ces-bubble" style={{ ...styles.kpi, borderTop: `3px solid ${c}` }}>
      <div style={{ fontSize: 11, color: theme.muted, textTransform: "uppercase", letterSpacing: "0.05em" }}>{label}</div>
      <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 21, fontWeight: 600, color: theme.text, marginTop: 6 }}>{value}</div>
    </div>
  );
}

/* =================================== Dashboard =================================== */
function Dashboard({ data, balances }) {
  const { theme, styles, fmt } = useUI();
  const bankTotal = data.banks.reduce((s, b) => s + (balances[b.accountId] || 0), 0);
  const ar = balances["1100"] || 0, ap = balances["2000"] || 0;

  // This-year performance for the headline KPIs
  const { from: yFrom, to: yTo } = rangeToDates("year");
  const mvYear = periodMovement(data.accounts, data.transactions, yFrom, yTo);
  const revenueYr = data.accounts.filter(a => a.type === "revenue").reduce((s, a) => s + (a.contra ? -(mvYear[a.id] || 0) : (mvYear[a.id] || 0)), 0);
  const expenseYr = data.accounts.filter(a => a.type === "expense").reduce((s, a) => s + (a.contra ? -(mvYear[a.id] || 0) : (mvYear[a.id] || 0)), 0);
  const profitYr = revenueYr - expenseYr;

  const months = last6Months();
  const series = months.map(({ label, from, to }) => {
    const mv = periodMovement(data.accounts, data.transactions, from, to);
    const rev = sumByType(data.accounts, mv, "revenue"), exp = sumByType(data.accounts, mv, "expense");
    return { label, revenue: Math.round(rev), expenses: Math.round(exp), net: Math.round(rev - exp) };
  });

  const overdue = data.invoices.filter(i => i.status !== "paid" && i.dueDate && i.dueDate < todayStr());
  const recent = [...data.transactions].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 7);

  // Upcoming statutory filings, derived from live tax-payable balances. Due
  // dates follow the FIRS convention (21st of the following month); verify
  // against your actual filing calendar.
  const nextMonth = new Date(); nextMonth.setMonth(nextMonth.getMonth() + 1);
  const due21 = localDateToISO(new Date(nextMonth.getFullYear(), nextMonth.getMonth(), 21));
  const taxFilings = data.accounts
    .filter(a => a.type === "liability" && /vat|wht|tax|paye|duty/i.test(a.name) && Math.abs(balances[a.id] || 0) > 0.5)
    .map(a => ({ name: a.name, amount: balances[a.id] || 0, due: due21 }));

  // Rule-based insights from the live books
  const insights = [];
  const margin = revenueYr ? profitYr / revenueYr : 0;
  if (revenueYr > 0) insights.push({ tone: margin >= 0.15 ? "emerald" : margin >= 0 ? "amber" : "rose", text: `Net margin this year is ${(margin * 100).toFixed(1)}% on ${fmt(revenueYr)} revenue.` });
  if (overdue.length > 0) insights.push({ tone: "rose", text: `${overdue.length} invoice${overdue.length > 1 ? "s are" : " is"} past due, worth ${fmt(overdue.reduce((s, i) => s + computeDocTotals(i, data.taxGroups).finalAmount - (i.amountPaid || 0), 0))} in collectible cash.` });
  const lowStock = data.inventory.filter(i => i.qty <= i.reorderLevel && i.reorderLevel > 0);
  if (lowStock.length > 0) insights.push({ tone: "amber", text: `${lowStock.length} inventory item${lowStock.length > 1 ? "s are" : " is"} at or below reorder level.` });
  const pendingFeed = data.bankFeed.filter(f => f.status === "uncategorized").length;
  if (pendingFeed > 0) insights.push({ tone: "amber", text: `${pendingFeed} bank feed line${pendingFeed > 1 ? "s" : ""} awaiting categorization - balances exclude them until posted.` });
  if (ap > bankTotal) insights.push({ tone: "rose", text: `Payables (${fmt(ap)}) exceed cash on hand (${fmt(bankTotal)}) - watch near-term liquidity.` });
  if (insights.length === 0) insights.push({ tone: "emerald", text: "No exceptions detected. Books are current and liquidity is comfortable." });

  const insightDot = (tone) => ({ emerald: theme.emerald, amber: theme.amber, rose: theme.rose }[tone] || theme.muted);

  return (
    <div>
      <PageHeader eyebrow="Overview" title="Dashboard" sub={fmtDate(new Date())} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 14, marginBottom: 20 }}>
        <Kpi label="Revenue (this year)" value={fmt(revenueYr)} tone="emerald" />
        <Kpi label="Expenses (this year)" value={fmt(expenseYr)} tone="rose" />
        <Kpi label="Profit (this year)" value={fmt(profitYr)} tone={profitYr >= 0 ? "emerald" : "rose"} />
        <Kpi label="Cash balance" value={fmt(bankTotal)} tone="emerald" />
        <Kpi label="Accounts receivable" value={fmt(ar)} />
        <Kpi label="Accounts payable" value={fmt(ap)} tone="amber" />
      </div>

      <div className="ces-card ces-bubble" style={styles.cardWide}>
        <div style={styles.cardTitle}>Revenue vs. expenses - last 6 months</div>
        <ResponsiveContainer width="100%" height={260}>
          <AreaChart data={series} margin={{ left: 0, right: 12, top: 10 }}>
            <defs>
              <linearGradient id="rev" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={theme.emerald} stopOpacity={0.35} /><stop offset="100%" stopColor={theme.emerald} stopOpacity={0} /></linearGradient>
              <linearGradient id="exp" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={theme.rose} stopOpacity={0.3} /><stop offset="100%" stopColor={theme.rose} stopOpacity={0} /></linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="2 4" stroke={theme.border} vertical={false} />
            <XAxis dataKey="label" tick={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, fill: theme.muted }} axisLine={{ stroke: theme.border }} tickLine={false} />
            <YAxis tick={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, fill: theme.muted }} axisLine={false} tickLine={false} tickFormatter={(v) => `${v / 1000}k`} />
            <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 12, background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 6, color: theme.text }} />
            <Area type="monotone" dataKey="revenue" stroke={theme.emerald} fill="url(#rev)" strokeWidth={2} />
            <Area type="monotone" dataKey="expenses" stroke={theme.rose} fill="url(#exp)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.35fr 1fr", gap: 20, marginTop: 20 }}>
        <div className="ces-card ces-bubble" style={{ ...styles.card, marginBottom: 0 }}>
          <div style={styles.cardTitle}>Recent transactions</div>
          {recent.length === 0 ? <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>Nothing posted yet.</div> : recent.map(t => {
            const amt = t.lines.reduce((s, l) => s + l.debit, 0);
            return (
              <div key={t.id} style={styles.attentionRow}>
                <div><div style={{ fontWeight: 500, fontSize: 13.5 }}>{t.memo}</div><div style={{ fontSize: 12, color: theme.muted }}>{fmtDate(t.date)} · {t.source}</div></div>
                <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13 }}>{fmt(amt)}</div>
              </div>
            );
          })}
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
          <div className="ces-card ces-bubble" style={{ ...styles.card, marginBottom: 0 }}>
            <div style={styles.cardTitle}>Upcoming tax filings</div>
            {taxFilings.length === 0 ? <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No tax balances awaiting remittance.</div> : taxFilings.map((t, i) => (
              <div key={i} style={styles.attentionRow}>
                <div><div style={{ fontWeight: 500, fontSize: 13.5 }}>{t.name}</div><div style={{ fontSize: 12, color: theme.muted }}>Due {fmtDate(t.due)}</div></div>
                <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13, color: theme.amber }}>{fmt(t.amount)}</div>
              </div>
            ))}
            {taxFilings.length > 0 && <div style={{ fontSize: 11, color: theme.muted, marginTop: 8 }}>Due dates assume the 21st-of-following-month convention - confirm against your filing calendar.</div>}
          </div>
          <div className="ces-card ces-bubble" style={{ ...styles.card, marginBottom: 0 }}>
            <div style={styles.cardTitle}>Bank balances</div>
            {data.banks.map(b => (
              <div key={b.id} style={styles.attentionRow}>
                <div><div style={{ fontWeight: 500, fontSize: 13.5 }}>{b.name}</div><div style={{ fontSize: 12, color: theme.muted }}>{b.accountNumber}</div></div>
                <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13, color: theme.emerald }}>{fmt(balances[b.accountId] || 0)}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginTop: 20 }}>
        <div className="ces-card ces-bubble" style={styles.card}>
          <div style={styles.cardTitle}>Financial insights</div>
          {insights.map((ins, i) => (
            <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "8px 0", borderBottom: i < insights.length - 1 ? `1px solid ${theme.border}` : "none" }}>
              <span style={{ width: 8, height: 8, borderRadius: 4, background: insightDot(ins.tone), marginTop: 5, flexShrink: 0 }} />
              <span style={{ fontSize: 13.5, lineHeight: 1.55 }}>{ins.text}</span>
            </div>
          ))}
        </div>
        <div className="ces-card ces-bubble" style={styles.card}>
          <div style={styles.cardTitle}>Needs attention</div>
          {overdue.length === 0 ? <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No overdue invoices. Books are current.</div> :
            overdue.map(i => (
              <div key={i.id} style={styles.attentionRow}>
                <div><div style={{ fontWeight: 500, fontSize: 13.5 }}>{i.id} · {i.customer}</div><div style={{ fontSize: 12, color: theme.muted }}>Due {fmtDate(i.dueDate)}</div></div>
                <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13, color: theme.rose }}>{fmt(computeDocTotals(i, data.taxGroups).finalAmount - (i.amountPaid || 0))}</div>
              </div>
            ))}
        </div>
      </div>
    </div>
  );
}
// IAS 21: monetary items held in a foreign currency are retranslated at the
// closing rate at each reporting date, with the difference recognized as an
// FX gain or loss. Since this system posts every transaction in the
// functional currency (it doesn't track a separate foreign-currency amount
// per transaction), revaluation works from a manually-confirmed foreign
// currency balance - typically read off the actual bank/creditor statement -
// rather than being derived automatically from posted transactions.
function computeFXRevaluation(account, glBalance, fcBalance, newRate) {
  const revaluedValue = fcBalance * newRate;
  const balanceChange = revaluedValue - glBalance; // change needed to the account's own balance
  const isDebitNormal = account.type === "asset" ? !account.contra : Boolean(account.contra);
  // A growing debit-normal (asset) balance is a gain; a growing credit-normal
  // (liability) balance means owing more, which is a loss - opposite signs.
  const pnlImpact = isDebitNormal ? balanceChange : -balanceChange;
  return { fcBalance, glBalance, revaluedValue, balanceChange, isDebitNormal, pnlImpact };
}
function FXRevaluationPanel({ data, balances, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const fxAccounts = data.accounts.filter(a => a.currency && a.status !== "inactive");
  const [drafts, setDrafts] = useState({});
  const draftFor = (a) => drafts[a.id] || { fcBalance: a.fcBalance ?? "", rate: a.fxRate ?? "" };
  const setDraft = (id, patch) => setDrafts(d => ({ ...d, [id]: { ...draftFor({ id }), ...patch } }));

  const postRevaluation = async (a) => {
    const draft = draftFor(a);
    const fcBalance = Number(draft.fcBalance) || 0, rate = Number(draft.rate) || 0;
    if (!rate) return notify("Enter this period's spot rate");
    const bal = balances[a.id] || 0;
    const r = computeFXRevaluation(a, bal, fcBalance, rate);
    if (Math.abs(r.balanceChange) < 0.5) {
      setData(d => ({ ...d, accounts: d.accounts.map(x => x.id === a.id ? { ...x, fcBalance, fxRate: rate } : x) }));
      return notify("No revaluation needed - already at the new rate");
    }
    if (!(await confirm(`Revalue ${a.name}? ${fmt(fcBalance)} ${a.currency} at ${rate} = ${fmt(r.revaluedValue)}, an FX ${r.pnlImpact >= 0 ? "gain" : "loss"} of ${fmt(Math.abs(r.pnlImpact))}.`))) return;
    const accountLine = { accountId: a.id, debit: r.isDebitNormal ? Math.max(0, r.balanceChange) : Math.max(0, -r.balanceChange), credit: r.isDebitNormal ? Math.max(0, -r.balanceChange) : Math.max(0, r.balanceChange) };
    const pnlLine = { accountId: "6010", debit: r.pnlImpact < 0 ? -r.pnlImpact : 0, credit: r.pnlImpact > 0 ? r.pnlImpact : 0 };
    const txn = buildTxn(`FX revaluation - ${a.name}`, todayStr(), [accountLine, pnlLine], "manual");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], accounts: d.accounts.map(x => x.id === a.id ? { ...x, fcBalance, fxRate: rate } : x) }));
    notify(`FX ${r.pnlImpact >= 0 ? "gain" : "loss"} of ${fmt(Math.abs(r.pnlImpact))} posted`);
  };

  if (fxAccounts.length === 0) return null;
  return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Foreign currency revaluation</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>For each account holding a foreign currency balance, confirm its actual foreign-currency amount (from your bank or vendor statement) and this period's spot rate - the ledger balance is revalued to match, with the difference posted as an FX gain or loss.</div>
      <div style={{ overflowX: "auto" }}>
      <table style={{ ...styles.table, marginTop: 10 }}>
        <thead><tr><th style={styles.th}>Account</th><th style={styles.th}>Currency</th><th style={{ ...styles.th, textAlign: "right" }}>FC balance</th><th style={{ ...styles.th, textAlign: "right" }}>Spot rate</th><th style={{ ...styles.th, textAlign: "right" }}>Ledger balance</th><th style={{ ...styles.th, textAlign: "right" }}>Revalued</th><th style={styles.th}></th></tr></thead>
        <tbody>{fxAccounts.map(a => {
          const draft = draftFor(a);
          const bal = balances[a.id] || 0;
          const r = computeFXRevaluation(a, bal, Number(draft.fcBalance) || 0, Number(draft.rate) || 0);
          return (
            <tr key={a.id}>
              <td style={styles.td}>{a.name}</td><td style={styles.td}>{a.currency}</td>
              <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={draft.fcBalance} onChange={e => setDraft(a.id, { fcBalance: e.target.value })} /></td>
              <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 80, textAlign: "right" }} value={draft.rate} onChange={e => setDraft(a.id, { rate: e.target.value })} /></td>
              <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(bal)}</td>
              <td style={{ ...styles.tdMono, textAlign: "right", color: Math.abs(r.pnlImpact) > 0.5 ? (r.pnlImpact > 0 ? theme.emerald : theme.rose) : theme.text }}>{Number(draft.rate) ? fmt(r.revaluedValue) : "-"}</td>
              <td style={styles.td}><button style={styles.btnGhost} onClick={() => postRevaluation(a)}>Revalue</button></td>
            </tr>
          );
        })}</tbody>
      </table>
      </div>
    </div>
  );
}
function ChartOfAccounts({ data, balances, setData, notify }) {
  const { theme, styles, fmt, openDrilldown } = useUI();
  const blankForm = () => ({ code: "", name: "", type: "expense", category: "Operating Expenses", parentId: "", description: "", taxAccountId: "", currency: "" });
  const [form, setForm] = useState(blankForm());
  const [editingId, setEditingId] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [mergingId, setMergingId] = useState(null);
  const [mergeTargetId, setMergeTargetId] = useState("");
  const [showInactive, setShowInactive] = useState(false);

  const taxAccounts = data.accounts.filter(a => a.type === "liability");
  const hasTransactions = (id) => data.transactions.some(t => t.lines.some(l => l.accountId === id));
  const hasChildren = (id) => data.accounts.some(a => a.parentId === id);
  const isBankLinked = (id) => data.banks.some(b => b.accountId === id);
  const isSystem = (id) => SYSTEM_ACCOUNT_IDS.includes(id);
  const rollup = (id) => (balances[id] || 0) + data.accounts.filter(a => a.parentId === id).reduce((s, c) => s + rollup(c.id), 0);
  const allDescendantIds = (id) => data.accounts.filter(a => a.parentId === id).flatMap(c => [c.id, ...allDescendantIds(c.id)]);

  const startCreate = (parent = null) => {
    setEditingId(null);
    setForm(parent
      ? { ...blankForm(), type: parent.type, category: parent.category, parentId: parent.id, code: "" }
      : blankForm());
    setShowForm(true);
  };
  const startEdit = (a) => {
    setEditingId(editingId === a.id ? null : a.id);
    setForm({ code: a.code, name: a.name, type: a.type, category: a.category, parentId: a.parentId || "", description: a.description || "", taxAccountId: a.taxAccountId || "", currency: a.currency || "" });
  };

  const saveAccount = () => {
    if (!form.code || !form.name) return notify("Enter a code and name");
    if (data.accounts.some(a => a.code === form.code && a.id !== editingId)) return notify("That account code already exists");
    if (editingId && form.parentId === editingId) return notify("An account can't be its own parent");
    const subtype = subtypeForCategory(form.type, form.category);
    if (editingId) {
      const existing = data.accounts.find(a => a.id === editingId);
      if (isSystem(editingId) && existing.type !== form.type) return notify("This is a system account - its type can't change, but you can rename or reclassify it");
      const wasBank = existing.subtype === "bank" || isBankLinked(editingId);
      const willBeBank = subtype === "bank";
      if (wasBank && !willBeBank && isBankLinked(editingId)) return notify("This account backs a bank in the Banking tab - remove that bank first before reclassifying it");
      setData(d => ({
        ...d,
        accounts: d.accounts.map(a => a.id === editingId ? { ...a, code: form.code, name: form.name, type: isSystem(editingId) ? a.type : form.type, category: form.category, subtype, parentId: form.parentId || null, description: form.description, taxAccountId: form.taxAccountId, currency: form.currency } : a),
        // Reclassified to Bank and not yet in the Banking tab: create its bank record.
        banks: (willBeBank && !d.banks.some(b => b.accountId === editingId))
          ? [...d.banks, { id: uid("bank"), name: form.name, accountNumber: "", accountId: editingId }]
          : d.banks.map(b => b.accountId === editingId ? { ...b, name: form.name } : b),
      }));
      notify(willBeBank ? "Account updated - available in the Banking tab" : "Account updated");
    } else {
      const isBank = subtype === "bank";
      setData(d => ({
        ...d,
        accounts: [...d.accounts, normalizeAccount({ id: form.code, code: form.code, name: form.name, type: form.type, category: form.category, subtype, parentId: form.parentId || null, description: form.description, taxAccountId: form.taxAccountId, currency: form.currency, status: "active" })],
        // A Bank-category account is a bank: it appears in the Banking tab too.
        banks: isBank ? [...d.banks, { id: uid("bank"), name: form.name, accountNumber: "", accountId: form.code }] : d.banks,
      }));
      notify(isBank ? "Bank account created - it now appears in both the Chart of Accounts and the Banking tab" : "Account created");
    }
    setShowForm(false); setEditingId(null); setForm(blankForm());
  };

  const toggleStatus = (a) => {
    if (isBankLinked(a.id) && a.status !== "inactive") return notify("This account is linked to a bank - remove the bank first");
    setData(d => ({ ...d, accounts: d.accounts.map(x => x.id === a.id ? { ...x, status: x.status === "inactive" ? "active" : "inactive" } : x) }));
    notify(a.status === "inactive" ? "Account reactivated" : "Account deactivated - it stays on past reports but is hidden from new postings");
  };

  const deleteAccount = (a) => {
    if (isSystem(a.id)) return notify("System accounts can't be deleted - deactivate instead");
    if (isBankLinked(a.id)) return notify("This account is linked to a bank and can't be deleted");
    if (hasTransactions(a.id)) return notify("This account has transactions - merge it into another account or deactivate it instead");
    if (hasChildren(a.id)) return notify("This account has sub-accounts - move or delete them first");
    setData(d => ({ ...d, accounts: d.accounts.filter(x => x.id !== a.id) }));
    notify(`${a.name} deleted`);
  };

  // Merge: every reference to the source account is rewritten to the target,
  // sub-accounts are re-parented, then the source is removed. Balances and
  // history are preserved - they now live under the target.
  const confirmMerge = (source) => {
    const target = data.accounts.find(a => a.id === mergeTargetId);
    if (!target) return notify("Choose an account to merge into");
    if (target.type !== source.type) return notify("Accounts can only merge within the same type");
    if (isSystem(source.id) || isBankLinked(source.id)) return notify("System and bank-linked accounts can't be merged away");
    setData(d => ({
      ...d,
      transactions: d.transactions.map(t => ({ ...t, lines: t.lines.map(l => l.accountId === source.id ? { ...l, accountId: target.id } : l) })),
      accounts: d.accounts.filter(a => a.id !== source.id).map(a => a.parentId === source.id ? { ...a, parentId: target.id } : a),
      expenses: d.expenses.map(e => e.category === source.id ? { ...e, category: target.id } : e),
      categoryRules: d.categoryRules.map(r => r.accountId === source.id ? { ...r, accountId: target.id } : r),
      bankFeed: d.bankFeed.map(f => f.accountId === source.id ? { ...f, accountId: target.id } : f),
      taxGroups: d.taxGroups.map(g => ({ ...g, components: g.components.map(c => c.accountId === source.id ? { ...c, accountId: target.id } : c) })),
      budgets: Object.fromEntries(Object.entries(d.budgets).map(([period, byAccount]) => [period, Object.fromEntries(Object.entries(byAccount).map(([k, v]) => [k === source.id ? target.id : k, v]))])),
    }));
    setMergingId(null); setMergeTargetId("");
    notify(`${source.name} merged into ${target.name} - all history moved`);
  };

  // Hierarchy renderer: roots of a category first, children indented beneath.
  const renderTree = (accounts, parentId, depth) => accounts
    .filter(a => (a.parentId || null) === parentId)
    .sort((x, y) => x.code.localeCompare(y.code))
    .flatMap(a => {
      const inactive = a.status === "inactive";
      if (inactive && !showInactive) return [];
      const merging = mergingId === a.id;
      const kids = hasChildren(a.id);
      const row = (
        <React.Fragment key={a.id}>
          <tr style={{ opacity: inactive ? 0.5 : 1 }}>
            <td style={styles.tdMono}>{a.code}</td>
            <td style={styles.td}>
              <span
                onClick={() => openDrilldown({ accountIds: kids ? [a.id, ...allDescendantIds(a.id)] : [a.id], label: `${a.code} ${a.name}`, asOf: todayStr() })}
                className="drill-row" style={{ paddingLeft: depth * 18, cursor: "pointer" }}
              >{depth > 0 && <span style={{ color: theme.muted }}>└ </span>}{a.name}</span>
              {inactive && <span style={{ ...styles.pillAmberSm, marginLeft: 8 }}>inactive</span>}
              {isSystem(a.id) && <span style={{ marginLeft: 8, fontSize: 10.5, color: theme.muted }}>system</span>}
              {a.currency && <span style={{ marginLeft: 8, fontSize: 10.5, color: theme.muted }}>{a.currency}</span>}
              {a.description && <div style={{ fontSize: 11.5, color: theme.muted, paddingLeft: depth * 18, marginTop: 2 }}>{a.description}</div>}
            </td>
            <td style={{ ...styles.tdMono, textAlign: "right", color: rollup(a.id) < 0 ? theme.rose : theme.text }}>{fmt(rollup(a.id))}{kids && <span style={{ fontSize: 10, color: theme.muted }}> Σ</span>}</td>
            <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
              <button style={styles.iconBtn} onClick={() => startEdit(a)} title="Edit">✎</button>
              <button style={styles.iconBtn} onClick={() => startCreate(a)} title="Add sub-account">+sub</button>
              <button style={styles.iconBtn} onClick={() => { setMergingId(merging ? null : a.id); setMergeTargetId(""); }} title="Merge into another account">⇄</button>
              <button style={styles.iconBtn} onClick={() => toggleStatus(a)} title={inactive ? "Reactivate" : "Deactivate"}>{inactive ? "▶" : "⏸"}</button>
              <button style={styles.iconBtn} onClick={() => deleteAccount(a)} title="Delete">🗑</button>
            </td>
          </tr>
          {merging && (
            <tr><td colSpan={4} style={{ ...styles.td, background: theme.panel2 }}>
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <span style={{ fontSize: 13 }}>Merge <strong>{a.name}</strong> and all its history into:</span>
                <select style={styles.inputSmall} value={mergeTargetId} onChange={e => setMergeTargetId(e.target.value)}>
                  <option value="">Select target account</option>
                  {data.accounts.filter(x => x.type === a.type && x.id !== a.id && x.status !== "inactive").map(x => <option key={x.id} value={x.id}>{x.code} · {x.name}</option>)}
                </select>
                <button style={styles.btnPrimary} onClick={() => confirmMerge(a)}>Merge</button>
              </div>
            </td></tr>
          )}
          {editingId === a.id && (
            <tr><td colSpan={4} style={{ ...styles.td, background: theme.panel2 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: theme.muted, marginBottom: 8 }}>Editing {a.code} · {a.name}</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 10 }}>
                <label style={{ fontSize: 12 }}>Account code<input style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.code} onChange={e => setForm({ ...form, code: e.target.value })} /></label>
                <label style={{ fontSize: 12 }}>Account name<input style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></label>
                <label style={{ fontSize: 12 }}>Type<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.type} onChange={e => { const t = e.target.value; setForm({ ...form, type: t, category: ACCOUNT_CATEGORIES[t][0], parentId: "" }); }} disabled={isSystem(a.id)}>
                  {ACCOUNT_TYPES.map(t => <option key={t} value={t}>{TYPE_LABELS[t]}</option>)}</select></label>
                <label style={{ fontSize: 12 }}>Category<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>
                  {ACCOUNT_CATEGORIES[form.type].map(c => <option key={c} value={c}>{c}</option>)}</select></label>
                <label style={{ fontSize: 12 }}>Parent account (optional)<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.parentId} onChange={e => setForm({ ...form, parentId: e.target.value })}>
                  <option value="">None (top level)</option>
                  {data.accounts.filter(x => x.type === form.type && x.id !== a.id && x.status !== "inactive").map(x => <option key={x.id} value={x.id}>{x.code} · {x.name}</option>)}</select></label>
                <label style={{ fontSize: 12 }}>Tax mapping (optional)<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.taxAccountId} onChange={e => setForm({ ...form, taxAccountId: e.target.value })}>
                  <option value="">None</option>{taxAccounts.map(x => <option key={x.id} value={x.id}>{x.code} · {x.name}</option>)}</select></label>
                <label style={{ fontSize: 12 }}>Currency (optional)<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.currency} onChange={e => setForm({ ...form, currency: e.target.value })}>
                  <option value="">Base currency</option>{CURRENCIES.map(c => <option key={c.code} value={c.code}>{c.code} ({c.symbol})</option>)}</select></label>
              </div>
              <label style={{ fontSize: 12, display: "block", marginTop: 10 }}>Description<textarea style={{ ...styles.input, width: "100%", marginTop: 4, minHeight: 54, resize: "vertical", fontFamily: "Inter, sans-serif" }} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></label>
              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <button style={styles.btnPrimary} onClick={saveAccount}>Save changes</button>
                <button style={styles.btnGhost} onClick={() => setEditingId(null)}>Cancel</button>
              </div>
            </td></tr>
          )}
        </React.Fragment>
      );
      return [row, ...renderTree(accounts, a.id, depth + 1)];
    });

  return (
    <div>
      <PageHeader eyebrow="Setup" title="Chart of Accounts" sub={`${data.accounts.length} accounts · ${data.accounts.filter(a => a.status === "inactive").length} inactive`} action={
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
            <input type="checkbox" checked={showInactive} onChange={e => setShowInactive(e.target.checked)} /> Show inactive
          </label>
          <button style={styles.btnPrimary} onClick={() => showForm ? (setShowForm(false), setEditingId(null)) : startCreate()}>{showForm ? "Cancel" : "+ New account"}</button>
        </div>
      } />

      {showForm && !editingId && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>New account</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 10, marginTop: 12 }}>
            <label style={{ fontSize: 12 }}>Account code<input style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.code} onChange={e => setForm({ ...form, code: e.target.value })} placeholder="e.g. 5150" /></label>
            <label style={{ fontSize: 12 }}>Account name<input style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="e.g. Vehicle Running Costs" /></label>
            <label style={{ fontSize: 12 }}>Type<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.type} onChange={e => { const t = e.target.value; setForm({ ...form, type: t, category: ACCOUNT_CATEGORIES[t][0], parentId: "" }); }} disabled={editingId && isSystem(editingId)}>
              {ACCOUNT_TYPES.map(t => <option key={t} value={t}>{TYPE_LABELS[t]}</option>)}</select></label>
            <label style={{ fontSize: 12 }}>Category<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>
              {ACCOUNT_CATEGORIES[form.type].map(c => <option key={c} value={c}>{c}</option>)}</select></label>
            <label style={{ fontSize: 12 }}>Parent account (optional)<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.parentId} onChange={e => setForm({ ...form, parentId: e.target.value })}>
              <option value="">None (top level)</option>
              {data.accounts.filter(a => a.type === form.type && a.id !== editingId && a.status !== "inactive").map(a => <option key={a.id} value={a.id}>{a.code} · {a.name}</option>)}</select></label>
            <label style={{ fontSize: 12 }}>Tax mapping (optional)<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.taxAccountId} onChange={e => setForm({ ...form, taxAccountId: e.target.value })}>
              <option value="">None</option>{taxAccounts.map(a => <option key={a.id} value={a.id}>{a.code} · {a.name}</option>)}</select></label>
            <label style={{ fontSize: 12 }}>Currency (optional)<select style={{ ...styles.input, width: "100%", marginTop: 4 }} value={form.currency} onChange={e => setForm({ ...form, currency: e.target.value })}>
              <option value="">Base currency</option>{CURRENCIES.map(c => <option key={c.code} value={c.code}>{c.code} ({c.symbol})</option>)}</select></label>
          </div>
          <label style={{ fontSize: 12, display: "block", marginTop: 10 }}>Description<textarea style={{ ...styles.input, width: "100%", marginTop: 4, minHeight: 54, resize: "vertical", fontFamily: "Inter, sans-serif" }} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} placeholder="What is recorded in this account" /></label>
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={saveAccount}>{editingId ? "Save changes" : "Create account"}</button>
        </div>
      )}

      {ACCOUNT_TYPES.map(type => {
        const typeAccounts = data.accounts.filter(a => a.type === type);
        const categories = ACCOUNT_CATEGORIES[type].filter(c => typeAccounts.some(a => a.category === c && (showInactive || a.status !== "inactive")));
        if (categories.length === 0) return null;
        return (
          <div key={type} style={styles.card}>
            <div style={styles.cardTitle}>{TYPE_LABELS[type]}</div>
            {categories.map(cat => {
              const catAccounts = typeAccounts.filter(a => a.category === cat);
              return (
                <div key={cat} style={{ marginTop: 14 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.muted, marginBottom: 4 }}>{cat}</div>
                  <table style={styles.table}>
                    <thead><tr><th style={{ ...styles.th, width: 70 }}>Code</th><th style={styles.th}>Account</th><th style={{ ...styles.th, textAlign: "right", width: 150 }}>Balance</th><th style={{ ...styles.th, width: 170 }}>Actions</th></tr></thead>
                    <tbody>{renderTree(catAccounts, null, 0)}</tbody>
                  </table>
                </div>
              );
            })}
          </div>
        );
      })}
      <FXRevaluationPanel data={data} balances={balances} setData={setData} notify={notify} />
      <div style={{ fontSize: 12, color: theme.muted }}>Σ marks a parent whose balance includes its sub-accounts. Accounts with transactions can be merged or deactivated but not deleted; system and bank-linked accounts are protected. Deactivated accounts stay on historical reports but disappear from posting screens.</div>
    </div>
  );
}

// Books are "closed" up to a single date (settings.lockDate) - anything on
// or before it can no longer be deleted or edited directly, the standard
// way accounting software prevents a reported/filed period from silently
// changing later. Fixing a mistake in a locked period means posting a new,
// dated correcting entry instead - never editing history.
function isLocked(date, data) {
  const lockDate = data.settings.lockDate;
  return Boolean(lockDate) && date <= lockDate;
}
function Journal({ data, setData, postTransaction, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [memo, setMemo] = useState(""); const [date, setDate] = useState(todayStr());
  const [recur, setRecur] = useState({ on: false, frequency: "monthly", endDate: "" });
  const [lines, setLines] = useState([{ accountId: "1000", debit: "", credit: "" }, { accountId: "4000", debit: "", credit: "" }]);
  const [viewingId, setViewingId] = useState(null);
  const [editingTxnId, setEditingTxnId] = useState(null);
  const updateLine = (i, field, val) => setLines(ls => ls.map((l, idx) => idx === i ? { ...l, [field]: val } : l));
  const addLine = () => setLines(ls => [...ls, { accountId: data.accounts[0]?.id, debit: "", credit: "" }]);
  const removeLine = (i) => setLines(ls => ls.filter((_, idx) => idx !== i));
  const totalDebit = lines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
  const totalCredit = lines.reduce((s, l) => s + (Number(l.credit) || 0), 0);
  const balanced = Math.abs(totalDebit - totalCredit) < 0.005 && totalDebit > 0;

  // Direct editing is limited to hand-made entries; entries generated by an
  // invoice, bill, payment, bank import etc. must be adjusted at their source
  // or cancelled with a reversing entry, so records never disagree.
  const isEditable = (t) => t.source === "manual";

  const resetForm = () => { setMemo(""); setDate(todayStr()); setLines([{ accountId: "1000", debit: "", credit: "" }, { accountId: "4000", debit: "", credit: "" }]); setEditingTxnId(null); };

  const submit = () => {
    if (!memo) return notify("Add a memo describing this entry");
    if (isLocked(date, data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - choose a later date, or use a reversing entry to adjust a locked period`);
    const cleanLines = lines.map(l => ({ accountId: l.accountId, debit: Number(l.debit) || 0, credit: Number(l.credit) || 0 }));
    if (editingTxnId) {
      const orig = data.transactions.find(t => t.id === editingTxnId);
      const txn = buildTxn(memo, date, cleanLines, orig?.source || "manual", orig?.docId || null);
      if (!txn) return notify("Entry not balanced");
      setData(d => ({ ...d, transactions: d.transactions.map(t => t.id === editingTxnId ? { ...txn, id: t.id } : t) }));
      notify("Entry updated"); resetForm();
    } else {
      const ok = postTransaction(memo, date, cleanLines);
      if (!ok) return;
      if (recur.on) {
        const nextDate = advanceDate(date, recur.frequency);
        setData(d => ({ ...d, recurringJournals: [...d.recurringJournals, { id: uid("rec"), memo, lines: cleanLines, frequency: recur.frequency, startDate: date, nextDate, endDate: recur.endDate || null, active: true }] }));
        notify(`Entry posted and scheduled ${recur.frequency} - next on ${fmtDate(nextDate)}`);
      } else notify("Entry posted");
      resetForm(); setRecur({ on: false, frequency: "monthly", endDate: "" });
    }
  };

  const toggleRecurring = (r) => setData(d => ({ ...d, recurringJournals: d.recurringJournals.map(x => x.id === r.id ? { ...x, active: !x.active } : x) }));
  const deleteRecurring = async (r) => {
    if (!(await confirm(`Delete the recurring schedule "${r.memo}"? Entries it has already posted stay on the ledger.`))) return;
    setData(d => ({ ...d, recurringJournals: d.recurringJournals.filter(x => x.id !== r.id) }));
    notify("Recurring schedule deleted");
  };
  const runDueNow = () => {
    const { data: next, posted } = processRecurringJournals(data);
    if (posted === 0) return notify("Nothing due - all recurring entries are up to date");
    setData(next);
    notify(`${posted} recurring entr${posted === 1 ? "y" : "ies"} posted`);
  };

  const startEdit = (t) => {
    setMemo(t.memo); setDate(t.date);
    setLines(t.lines.map(l => ({ accountId: l.accountId, debit: l.debit || "", credit: l.credit || "" })));
    setEditingTxnId(t.id); setViewingId(null);
    window.scrollTo && window.scrollTo({ top: 0, behavior: "smooth" });
  };

  // Post an equal-and-opposite entry - the accountant's way to adjust
  // anything, including entries generated by documents.
  const reverseEntry = (t) => {
    const rev = t.lines.map(l => ({ accountId: l.accountId, debit: l.credit, credit: l.debit }));
    const ok = postTransaction(`REVERSAL of: ${t.memo}`, todayStr(), rev, "reversal");
    if (ok) { notify("Reversing entry posted"); setViewingId(null); }
  };

  const sorted = [...data.transactions].sort((a, b) => b.date.localeCompare(a.date));
  const viewing = viewingId ? data.transactions.find(t => t.id === viewingId) : null;

  if (viewing) {
    const editable = isEditable(viewing);
    const locked = isLocked(viewing.date, data);
    return (
      <div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 8 }}>
          <button style={styles.btnGhost} onClick={() => setViewingId(null)}>← Back to journal</button>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={styles.btnGhost} onClick={() => reverseEntry(viewing)}>⇄ Post reversing entry</button>
            <button style={{ ...styles.btnGhost, color: locked ? theme.muted : theme.rose, cursor: locked ? "not-allowed" : "pointer" }} disabled={locked} onClick={async () => { if (!(await confirm("Delete this journal entry? This removes it from the ledger permanently."))) return; setData(d => deleteTransactionFromData(d, viewing, notify)); setViewingId(null); notify("Entry deleted"); }}>🗑 Delete</button>
            {editable && !locked && <button style={styles.btnPrimary} onClick={() => startEdit(viewing)}>✎ Edit entry</button>}
          </div>
        </div>
        <PageHeader eyebrow="Journal entry" title={viewing.memo} sub={`${fmtDate(viewing.date)} · source: ${viewing.source}${viewing.docId ? ` · ${viewing.docId}` : ""}`} />
        {locked && <div style={{ fontSize: 12.5, color: theme.amber, marginBottom: 14 }}>🔒 This entry falls on or before the books-locked date ({fmtDate(data.settings.lockDate)}) and can't be edited or deleted. Post a reversing entry above to correct it instead - that keeps the locked period untouched.</div>}
        {!editable && !locked && <div style={{ fontSize: 12.5, color: theme.amber, marginBottom: 14 }}>This entry was generated by {viewing.docId ? <strong>{viewing.docId}</strong> : `a ${viewing.source}`} - edit it from its source screen so the document and ledger stay in sync, or post a reversing entry here to cancel its effect.</div>}
        <div className="ces-card" style={styles.card}>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Account</th><th style={{ ...styles.th, textAlign: "right" }}>Debit</th><th style={{ ...styles.th, textAlign: "right" }}>Credit</th></tr></thead>
            <tbody>{viewing.lines.map((l, i) => { const acc = data.accounts.find(a => a.id === l.accountId); return (
              <tr key={i}><td style={styles.td}>{acc ? `${acc.code} · ${acc.name}` : l.accountId}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{l.debit ? fmt(l.debit) : ""}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{l.credit ? fmt(l.credit) : ""}</td></tr>
            ); })}</tbody>
            <tfoot><tr><td style={{ ...styles.td, fontWeight: 700 }}>Totals</td>
              <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(viewing.lines.reduce((s, l) => s + l.debit, 0))}</td>
              <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(viewing.lines.reduce((s, l) => s + l.credit, 0))}</td></tr></tfoot>
          </table>
        </div>
      </div>
    );
  }

  return (
    <div>
      <PageHeader eyebrow="Books" title="Journal & Ledger" sub={`${data.transactions.length} entries`} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>{editingTxnId ? "Edit journal entry" : "New journal entry"}</div>
        {editingTxnId && <div style={{ fontSize: 12, color: theme.amber, marginTop: 4 }}>Editing replaces the original entry in the ledger. <button style={{ ...styles.iconBtn, textDecoration: "underline" }} onClick={resetForm}>Cancel edit</button></div>}
        <div style={{ display: "flex", gap: 10, marginTop: 10 }}>
          <input type="date" style={styles.input} value={date} onChange={e => setDate(e.target.value)} />
          <input style={{ ...styles.input, flex: 1 }} placeholder="Memo" value={memo} onChange={e => setMemo(e.target.value)} />
        </div>
        <table style={{ ...styles.table, marginTop: 12 }}>
          <thead><tr><th style={styles.th}>Account</th><th style={{ ...styles.th, textAlign: "right" }}>Debit</th><th style={{ ...styles.th, textAlign: "right" }}>Credit</th><th></th></tr></thead>
          <tbody>{lines.map((l, i) => (
            <tr key={i}>
              <td style={styles.td}><select style={styles.inputSmall} value={l.accountId} onChange={e => updateLine(i, "accountId", e.target.value)}>{activeAccounts(data).map(a => <option key={a.id} value={a.id}>{a.code} · {a.name}</option>)}</select></td>
              <td style={{ ...styles.td, textAlign: "right" }}><input style={{ ...styles.inputSmall, width: 110, textAlign: "right" }} type="number" value={l.debit} onChange={e => updateLine(i, "debit", e.target.value)} /></td>
              <td style={{ ...styles.td, textAlign: "right" }}><input style={{ ...styles.inputSmall, width: 110, textAlign: "right" }} type="number" value={l.credit} onChange={e => updateLine(i, "credit", e.target.value)} /></td>
              <td><button style={styles.iconBtn} onClick={() => removeLine(i)}>✕</button></td>
            </tr>
          ))}</tbody>
          <tfoot><tr><td style={styles.td}><button style={styles.btnGhost} onClick={addLine}>+ line</button></td>
            <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(totalDebit)}</td>
            <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(totalCredit)}</td><td></td></tr></tfoot>
        </table>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 10 }}>
          <span style={{ fontSize: 12, color: balanced ? theme.emerald : theme.rose, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{balanced ? "✓ balanced" : "not balanced"}</span>
          {!editingTxnId && (
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13 }}>
              <input type="checkbox" checked={recur.on} onChange={e => setRecur({ ...recur, on: e.target.checked })} /> Repeat this entry
            </label>
          )}
          {recur.on && !editingTxnId && <>
            <select style={styles.inputSmall} value={recur.frequency} onChange={e => setRecur({ ...recur, frequency: e.target.value })}>
              <option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="quarterly">Quarterly</option><option value="yearly">Yearly</option>
            </select>
            <label style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 5 }}>until <input type="date" style={styles.inputSmall} value={recur.endDate} onChange={e => setRecur({ ...recur, endDate: e.target.value })} /></label>
          </>}
          <button style={styles.btnPrimary} onClick={submit} disabled={!balanced}>{editingTxnId ? "Save changes" : recur.on ? "Post & schedule" : "Post entry"}</button>
        </div>
      </div>
      {data.recurringJournals.length > 0 && (
        <div className="ces-card" style={styles.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <div style={styles.cardTitle}>Recurring journals</div>
            <button style={styles.btnGhost} onClick={runDueNow}>Run due now</button>
          </div>
          <div style={{ overflowX: "auto" }}>
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}>Memo</th><th style={styles.th}>Frequency</th><th style={styles.th}>Next run</th><th style={styles.th}>Ends</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
            <tbody>{data.recurringJournals.map(r => (
              <tr key={r.id} style={{ opacity: r.active ? 1 : 0.55 }}>
                <td style={styles.td}>{r.memo}</td>
                <td style={styles.td}>{r.frequency}</td>
                <td style={styles.tdMono}>{fmtDate(r.nextDate)}</td>
                <td style={styles.tdMono}>{r.endDate ? fmtDate(r.endDate) : "Open-ended"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.lines.reduce((s, l) => s + l.debit, 0))}</td>
                <td style={styles.td}><span style={{ ...styles.pill, ...(r.active ? styles.pillGreen : styles.pillAmber) }}>{r.active ? "active" : "paused"}</span></td>
                <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                  <button style={styles.iconBtn} onClick={() => toggleRecurring(r)}>{r.active ? "\u23f8 pause" : "\u25b6 resume"}</button>{" "}
                  <button style={styles.iconBtn} onClick={() => deleteRecurring(r)}>🗑</button>
                </td>
              </tr>
            ))}</tbody>
          </table>
          </div>
          <div style={{ fontSize: 12, color: theme.muted, marginTop: 8 }}>Due occurrences post automatically when the app opens, each dated on its scheduled day.</div>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Recent entries</div>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={styles.th}>Source</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
          <tbody>{sorted.slice(0, 40).map(t => (
            <tr key={t.id} onClick={() => setViewingId(t.id)} style={{ cursor: "pointer" }}>
              <td style={styles.tdMono}>{fmtDate(t.date)}</td><td style={styles.td}>{isLocked(t.date, data) && <span title="Locked" style={{ marginRight: 5 }}>🔒</span>}{t.memo}</td>
              <td style={{ ...styles.td, fontSize: 12, color: theme.muted }}>{t.source}</td>
              <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t.lines.reduce((s, l) => s + l.debit, 0))}</td></tr>
          ))}</tbody>
        </table>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>Click an entry to see its debit/credit lines, edit it (manual and imported entries), or post a reversing adjustment.</div>
      </div>
    </div>
  );
}

/* =================================== Banks =================================== */
function Banks({ data, setData, balances, notify, postTransaction }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [asOf, setAsOf] = useState(todayStr());
  const [viewingBankId, setViewingBankId] = useState(null);
  const [form, setForm] = useState({ name: "", accountNumber: "", opening: "", code: "" });
  const asOfBalances = useMemo(() => computeBalances(data.accounts, data.transactions, asOf), [data, asOf]);

  const addBank = () => {
    if (!form.name) return notify("Name the bank account");
    if (form.code && data.accounts.some(a => String(a.code) === String(form.code) || String(a.id) === String(form.code))) return notify(`Account code ${form.code} is already taken - pick another or leave it blank to auto-assign`);
    setData(d => {
      // Code generation checks every existing account (including inactive and
      // non-bank accounts like Petty Cash), so two records can never share a
      // ledger account - that was making one bank's activity show in all.
      const code = form.code ? String(form.code) : nextFreeCode(d.accounts, 1020, 10);
      const account = normalizeAccount({ id: code, code, name: form.name, type: "asset", category: "Bank", status: "active" });
      const opening = Number(form.opening) || 0;
      const openTxn = opening > 0 ? buildTxn(`Opening balance - ${form.name}`, todayStr(), [{ accountId: code, debit: opening, credit: 0 }, { accountId: "3000", debit: 0, credit: opening }], "manual") : null;
      return {
        ...d,
        accounts: [...d.accounts, account],
        banks: [...d.banks, { id: uid("bank"), name: form.name, accountNumber: form.accountNumber, accountId: code }],
        transactions: openTxn ? [...d.transactions, openTxn] : d.transactions,
      };
    });
    setForm({ name: "", accountNumber: "", opening: "", code: "" });
    notify("Bank account added - it appears in the Chart of Accounts under Assets > Bank");
  };

  const deleteBank = async (b, e) => {
    e.stopPropagation();
    const bal = balances[b.accountId] || 0;
    if (Math.abs(bal) > 0.5) return notify(`${b.name} still has a balance of ${fmt(bal)} - transfer it out or delete/uncategorize its transactions first`);
    const pend = data.bankFeed.filter(f => f.bankId === b.id).length;
    if (pend > 0) return notify(`${b.name} has ${pend} feed line(s) - delete or move them first`);
    if (!(await confirm(`Remove ${b.name}? Its ledger account is deactivated (history stays on reports) and its rules are removed.`))) return;
    setData(d => ({
      ...d,
      banks: d.banks.filter(x => x.id !== b.id),
      categoryRules: d.categoryRules.filter(r => r.bankId !== b.id),
      accounts: d.accounts.map(a => a.id === b.accountId ? { ...a, status: "inactive" } : a),
    }));
    notify(`${b.name} removed`);
  };

  const viewingBank = viewingBankId ? data.banks.find(b => b.id === viewingBankId) : null;
  if (viewingBank) return <BankTransactions bank={viewingBank} data={data} balances={balances} setData={setData} notify={notify} onBack={() => setViewingBankId(null)} />;

  const thisMonthStart = localDateToISO(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
  const monthFlows = (accountId) => {
    const lines = data.transactions.filter(t => t.date >= thisMonthStart).flatMap(t => t.lines.filter(l => l.accountId === accountId));
    return { inflow: lines.reduce((s, l) => s + l.debit, 0), outflow: lines.reduce((s, l) => s + l.credit, 0) };
  };

  return (
    <div>
      <PageHeader eyebrow="Books" title="Banks" sub={`${data.banks.length} accounts connected`} action={
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ fontSize: 12, color: theme.muted, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>balance as of</span>
          <input type="date" style={styles.input} value={asOf} onChange={e => setAsOf(e.target.value)} />
        </div>
      } />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(290px, 1fr))", gap: 16, marginBottom: 20 }}>
        {data.banks.map(b => {
          const flows = monthFlows(b.accountId);
          const pending = data.bankFeed.filter(f => f.status === "uncategorized" && f.bankId === b.id).length;
          return (
            <div className="ces-card" key={b.id} style={{ ...styles.card, marginBottom: 0, cursor: "pointer", borderTop: `3px solid ${theme.accent}`, display: "flex", flexDirection: "column" }} onClick={() => setViewingBankId(b.id)}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <div style={{ fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontSize: 17, fontWeight: 600 }}>{b.name}</div>
                  <div style={{ fontSize: 11, color: theme.muted, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", marginTop: 2 }}>{b.accountNumber || "-"} · GL {b.accountId}</div>
                </div>
                <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                  {pending > 0 && <span style={{ ...styles.pill, ...styles.pillAmber }}>{pending} to categorize</span>}
                  <div style={{ width: 34, height: 34, borderRadius: 8, background: theme.accentSoft, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>🏦</div>
                  <button style={styles.iconBtn} title="Remove bank" onClick={(e) => deleteBank(b, e)}>🗑</button>
                </div>
              </div>
              <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 25, fontWeight: 600, color: theme.emerald, marginTop: 12 }}>{fmt(asOfBalances[b.accountId] || 0)}</div>
              <div style={{ fontSize: 11, color: theme.muted, marginTop: 2 }}>as of {fmtDate(asOf)}</div>
              <div style={{ display: "flex", gap: 18, marginTop: 12, fontSize: 12 }}>
                <div><span style={{ color: theme.muted }}>In this month </span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", color: theme.emerald }}>{fmt(flows.inflow)}</span></div>
                <div><span style={{ color: theme.muted }}>Out </span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", color: theme.rose }}>{fmt(flows.outflow)}</span></div>
              </div>
              <BankMiniTrend data={data} accountId={b.accountId} />
              <div style={{ fontSize: 12, color: theme.accent, fontWeight: 600, marginTop: "auto", paddingTop: 8 }}>Transactions · Import · Rules →</div>
            </div>
          );
        })}
      </div>

      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Connect a bank account</div>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 180 }} placeholder="Bank & account name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
          <input style={styles.input} placeholder="Account number" value={form.accountNumber} onChange={e => setForm({ ...form, accountNumber: e.target.value })} />
          <input type="number" style={styles.input} placeholder="Opening balance (optional)" value={form.opening} onChange={e => setForm({ ...form, opening: e.target.value })} />
          <input style={{ ...styles.input, width: 150 }} placeholder="GL code (auto)" value={form.code} onChange={e => setForm({ ...form, code: e.target.value })} title="Ledger account code - leave blank to auto-assign" />
          <button style={styles.btnPrimary} onClick={addBank}>Add bank</button>
        </div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 8 }}>Adding a bank creates its ledger account automatically under Assets &gt; Bank in the Chart of Accounts. Leave GL code blank to auto-assign the next free code.</div>
      </div>

      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Balances by month</div>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Month end</th>{data.banks.map(b => <th key={b.id} style={{ ...styles.th, textAlign: "right" }}>{b.name}</th>)}</tr></thead>
          <tbody>
            {last6Months().map(m => {
              const bal = computeBalances(data.accounts, data.transactions, m.to);
              return <tr key={m.label}><td style={styles.tdMono}>{fmtDate(m.to)}</td>{data.banks.map(b => <td key={b.id} style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(bal[b.accountId] || 0)}</td>)}</tr>;
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
// Per-bank statement view: every ledger line touching this bank's account,
// with a running balance, money in/out totals, and Zoho-style matching -
// unmatched imported receipts/payments can be matched to open invoices/bills,
// which re-books the entry against AR/AP and marks the document paid.
function BankTransactions({ bank, data, balances, setData, notify, onBack }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [range, setRange] = useState("all");
  const [customFrom, setCustomFrom] = useState(localDateToISO(new Date(Date.now() - 13 * 86400000)));
  const [customTo, setCustomTo] = useState(todayStr());
  const [subTab, setSubTab] = useState("transactions");
  const [matchingTxnId, setMatchingTxnId] = useState(null);
  const [selectedDocId, setSelectedDocId] = useState("");
  const [selectedIds, setSelectedIds] = useState(new Set());
  const effRange = range === "custom" ? { from: customFrom, to: customTo } : range;
  const { from, to } = rangeToDates(effRange);

  const all = data.transactions
    .flatMap(t => t.lines.filter(l => l.accountId === bank.accountId).map(l => ({ ...l, date: t.date, memo: t.memo, source: t.source, txnId: t.id, matchedDocId: t.matchedDocId })))
    .sort((a, b) => a.date.localeCompare(b.date));
  let running = 0;
  const withRunning = all.map(l => { running += l.debit - l.credit; return { ...l, running }; });
  const visible = withRunning.filter(l => l.date >= from && l.date <= to);
  const moneyIn = visible.reduce((s, l) => s + l.debit, 0);
  const moneyOut = visible.reduce((s, l) => s + l.credit, 0);

  const openInvoices = data.invoices.filter(i => (i.amountPaid || 0) < computeDocTotals(i, data.taxGroups).finalAmount - 0.5);
  const openBills = data.bills.filter(b => (b.amountPaid || 0) < computeDocTotals(b, data.taxGroups).finalAmount - 0.5);

  // A line can be matched if it came from a bank import and hasn't been
  // matched yet. Money in matches invoices; money out matches bills.
  const matchable = (l) => l.source === "bank-import" && !l.matchedDocId;

  const startMatch = (l) => {
    setMatchingTxnId(l.txnId);
    const candidates = l.debit > 0 ? openInvoices : openBills;
    setSelectedDocId(candidates[0]?.id || "");
  };

  const confirmMatch = (l) => {
    if (!selectedDocId) return notify("Choose a document to match against");
    const amt = l.debit > 0 ? l.debit : l.credit;
    const isInvoice = l.debit > 0;
    setData(d => ({
      ...d,
      // Re-book the imported entry: the bank side stays, the category side is
      // replaced with AR (for receipts) or AP (for payments made).
      transactions: d.transactions.map(t => t.id !== l.txnId ? t : {
        ...t,
        matchedDocId: selectedDocId,
        memo: `${t.memo} (matched ${selectedDocId})`,
        lines: isInvoice
          ? [{ accountId: bank.accountId, debit: amt, credit: 0 }, { accountId: "1100", debit: 0, credit: amt }]
          : [{ accountId: "2000", debit: amt, credit: 0 }, { accountId: bank.accountId, debit: 0, credit: amt }],
      }),
      invoices: isInvoice ? d.invoices.map(i => i.id === selectedDocId ? { ...i, amountPaid: (i.amountPaid || 0) + amt } : i) : d.invoices,
      bills: !isInvoice ? d.bills.map(b => b.id === selectedDocId ? { ...b, amountPaid: (b.amountPaid || 0) + amt } : b) : d.bills,
      payments: [...d.payments, { id: uid("pay"), date: l.date, type: isInvoice ? "received" : "paid", amount: amt, bankId: bank.id, relatedType: isInvoice ? "invoice" : "bill", relatedId: selectedDocId, memo: `Matched from bank: ${l.memo}`, txnId: l.txnId }],
    }));
    setMatchingTxnId(null); setSelectedDocId("");
    notify(`Matched to ${selectedDocId} - ledger re-booked against ${isInvoice ? "receivables" : "payables"}`);
  };

  const pendingCount = data.bankFeed.filter(f => f.status === "uncategorized" && f.bankId === bank.id).length;

  const toggleSelect = (txnId) => setSelectedIds(prev => { const s = new Set(prev); s.has(txnId) ? s.delete(txnId) : s.add(txnId); return s; });
  const visibleTxnIds = [...new Set(visible.map(l => l.txnId))];
  const allSelected = visibleTxnIds.length > 0 && visibleTxnIds.every(id => selectedIds.has(id));
  const toggleSelectAll = () => setSelectedIds(allSelected ? new Set() : new Set(visibleTxnIds));

  // Bulk uncategorize: bank-import entries (not matched to documents) drop
  // their ledger posting and return to the uncategorized feed.
  const bulkUncategorize = async () => {
    const targets = data.transactions.filter(t => selectedIds.has(t.id) && t.source === "bank-import");
    if (targets.length === 0) return notify("None of the selected lines are categorized bank imports");
    const matched = targets.filter(t => t.matchedDocId);
    if (matched.length > 0) return notify(`${matched.length} selected line(s) are matched to invoices/bills - delete the linked payment first`);
    if (!(await confirm(`Uncategorize ${targets.length} line(s)? Their ledger entries are removed and they return to the feed.`))) return;
    setData(d => ({
      ...d,
      transactions: d.transactions.filter(t => !targets.some(x => x.id === t.id)),
      bankFeed: d.bankFeed.map(f => targets.some(t => t.feedId === f.id) ? { ...f, status: "uncategorized", txnId: null, accountId: null } : f),
    }));
    setSelectedIds(new Set());
    notify(`${targets.length} line(s) uncategorized`);
  };

  // Bulk delete: each entry goes through the same per-source rules as a single
  // delete (documents blocked, feed lines returned + deletable, payments roll
  // back, depreciation rolls the register back).
  const bulkDelete = async () => {
    const targets = data.transactions.filter(t => selectedIds.has(t.id));
    if (targets.length === 0) return notify("Nothing selected");
    const blocked = targets.filter(t => t.source === "invoice" || t.source === "bill" || t.source === "fixed-asset" || t.matchedDocId);
    const deletable = targets.filter(t => !blocked.some(b => b.id === t.id));
    if (deletable.length === 0) return notify("Selected entries belong to documents or matched items - delete those from their source instead");
    if (!(await confirm(`Delete ${deletable.length} entr${deletable.length === 1 ? "y" : "ies"} from the ledger?${blocked.length ? ` (${blocked.length} document-linked line(s) will be skipped.)` : ""} Bank-import lines return to the feed where they can then be deleted to the bin.`))) return;
    setData(d => {
      let next = d;
      deletable.forEach(t => { next = deleteTransactionFromData(next, t, () => {}); });
      return next;
    });
    setSelectedIds(new Set());
    notify(`${deletable.length} deleted${blocked.length ? ` · ${blocked.length} skipped (document-linked)` : ""}`);
  };

  return (
    <div>
      <div style={{ marginBottom: 12 }}><button style={styles.btnGhost} onClick={onBack}>← All banks</button></div>
      <PageHeader eyebrow={bank.accountNumber || "Bank account"} title={bank.name} sub={`Current balance ${fmt(balances[bank.accountId] || 0)}`} action={subTab === "transactions" ? <RangePicker range={range} setRange={setRange} customFrom={customFrom} setCustomFrom={setCustomFrom} customTo={customTo} setCustomTo={setCustomTo} /> : null} />
      <div style={{ display: "flex", gap: 6, marginBottom: 16 }}>
        {[["transactions", "Transactions"], ["reconcile", "Reconcile"], ["feed", `Import & Rules${pendingCount ? ` (${pendingCount} pending)` : ""}`]].map(([id, label]) => (
          <button key={id} style={{ ...styles.btnGhost, ...(subTab === id ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }} onClick={() => setSubTab(id)}>{label}</button>
        ))}
      </div>
      {subTab === "reconcile" && <BankReconcile bank={bank} data={data} balances={balances} setData={setData} notify={notify} />}
      {subTab === "feed" && <BankFeedPanel bank={bank} data={data} setData={setData} notify={notify} />}
      {subTab === "transactions" && <>
      {pendingCount > 0 && <div className="ces-card" style={{ ...styles.card, borderLeft: `3px solid ${theme.amber}`, padding: "12px 16px", fontSize: 13 }}>⚠ {pendingCount} uploaded line{pendingCount > 1 ? "s" : ""} for this bank {pendingCount > 1 ? "are" : "is"} still <strong>uncategorized</strong> and not reflected in this balance - categorize {pendingCount > 1 ? "them" : "it"} in the Import & Rules tab above.</div>}
      <div style={styles.kpiRow}>
        <Kpi label="Money in (range)" value={fmt(moneyIn)} tone="emerald" />
        <Kpi label="Money out (range)" value={fmt(moneyOut)} tone="rose" />
        <Kpi label="Net movement" value={fmt(moneyIn - moneyOut)} tone={moneyIn - moneyOut >= 0 ? "emerald" : "rose"} />
        <Kpi label="Awaiting match" value={String(visible.filter(matchable).length)} tone="amber" />
      </div>
      <div className="ces-card" style={styles.card}>
        {selectedIds.size > 0 && (
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
            <span style={{ fontSize: 13, fontWeight: 600 }}>{selectedIds.size} selected</span>
            <button style={styles.btnGhost} onClick={bulkUncategorize}>↩ Uncategorize selected</button>
            <button style={{ ...styles.btnGhost, color: theme.rose }} onClick={bulkDelete}>🗑 Delete selected</button>
            <button style={styles.iconBtn} onClick={() => setSelectedIds(new Set())}>clear</button>
          </div>
        )}
        <div style={{ overflowX: "auto" }}>
        <table style={styles.table}>
          <thead><tr><th style={{ ...styles.th, width: 30 }}><input type="checkbox" checked={allSelected} onChange={toggleSelectAll} title="Select all visible" /></th><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={styles.th}>Source</th><th style={{ ...styles.th, textAlign: "right" }}>Money in</th><th style={{ ...styles.th, textAlign: "right" }}>Money out</th><th style={{ ...styles.th, textAlign: "right" }}>Balance</th><th style={styles.th}>Matching</th></tr></thead>
          <tbody>{[...visible].reverse().map((l, i) => {
            const isMatching = matchingTxnId === l.txnId;
            const candidates = l.debit > 0 ? openInvoices : openBills;
            return (
              <React.Fragment key={i}>
                <tr>
                  <td style={styles.td}><input type="checkbox" checked={selectedIds.has(l.txnId)} onChange={() => toggleSelect(l.txnId)} /></td>
                  <td style={styles.tdMono}>{fmtDate(l.date)}</td><td style={styles.td}>{l.memo}</td>
                  <td style={{ ...styles.td, fontSize: 12, color: theme.muted }}>{l.source}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right", color: theme.emerald }}>{l.debit ? fmt(l.debit) : ""}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right", color: theme.rose }}>{l.credit ? fmt(l.credit) : ""}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(l.running)}</td>
                  <td style={styles.td}>
                    {l.matchedDocId ? <span style={{ ...styles.pill, ...styles.pillGreen }}>✓ {l.matchedDocId}</span>
                      : matchable(l) ? <button style={styles.btnGhost} onClick={(e) => { e.stopPropagation(); isMatching ? setMatchingTxnId(null) : startMatch(l); }}>{isMatching ? "Cancel" : "Match"}</button>
                      : <span style={{ fontSize: 11, color: theme.muted }}>-</span>}
                  </td>
                </tr>
                {isMatching && (
                  <tr>
                    <td colSpan={8} style={{ ...styles.td, background: theme.panel2 }}>
                      {candidates.length === 0 ? (
                        <span style={{ fontSize: 13, color: theme.muted }}>No open {l.debit > 0 ? "invoices" : "bills"} to match against.</span>
                      ) : (
                        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                          <span style={{ fontSize: 13 }}>Match this {l.debit > 0 ? "receipt" : "payment"} of <strong style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(l.debit || l.credit)}</strong> to:</span>
                          <select style={styles.inputSmall} value={selectedDocId} onChange={e => setSelectedDocId(e.target.value)}>
                            {candidates.map(dcc => { const t = computeDocTotals(dcc, data.taxGroups).finalAmount - (dcc.amountPaid || 0); return <option key={dcc.id} value={dcc.id}>{dcc.id} · {dcc.customer || dcc.vendor} · {fmt(t)} outstanding</option>; })}
                          </select>
                          <button style={styles.btnPrimary} onClick={() => confirmMatch(l)}>Confirm match</button>
                        </div>
                      )}
                    </td>
                  </tr>
                )}
              </React.Fragment>
            );
          })}</tbody>
        </table>
        </div>
        {visible.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No transactions for this bank in the selected range.</div>}
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>Lines imported from bank statements can be matched to open invoices (money in) or bills (money out). Matching re-books the entry against receivables/payables and updates the document's paid amount.</div>
      </div>
      </>}
    </div>
  );
}
// Standard bank reconciliation: enter the statement's ending balance and
// date, tick off which ledger lines have cleared, and the difference must
// reach zero before you can finish. Completed reconciliations lock their
// lines so a later reconciliation starts clean and history stays intact.
function BankReconcile({ bank, data, balances, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const draft = data.reconciliations.find(r => r.bankId === bank.id && !r.completed);
  const history = data.reconciliations.filter(r => r.bankId === bank.id && r.completed).sort((a, b) => b.date.localeCompare(a.date));
  const [form, setForm] = useState({ date: todayStr(), statementBalance: "" });

  const startReconciliation = () => {
    if (form.statementBalance === "") return notify("Enter the statement's ending balance");
    setData(d => ({ ...d, reconciliations: [...d.reconciliations, { id: uid("rec"), bankId: bank.id, date: form.date, statementBalance: Number(form.statementBalance), clearedTxnIds: [], completed: false }] }));
    notify("Reconciliation started - tick off cleared items below");
  };
  const cancelReconciliation = async () => {
    if (!(await confirm("Discard this reconciliation in progress? No ledger entries are affected."))) return;
    setData(d => ({ ...d, reconciliations: d.reconciliations.filter(r => r.id !== draft.id) }));
  };

  const lines = data.transactions
    .filter(t => t.date <= (draft?.date || todayStr()))
    .flatMap(t => t.lines.filter(l => l.accountId === bank.accountId).map(l => ({ ...l, date: t.date, memo: t.memo, source: t.source, txnId: t.id })))
    .filter(l => !data.reconciliations.some(r => r.completed && r.bankId === bank.id && r.clearedTxnIds.includes(l.txnId))) // already reconciled elsewhere
    .sort((a, b) => a.date.localeCompare(b.date));

  const toggleCleared = (txnId) => setData(d => ({ ...d, reconciliations: d.reconciliations.map(r => r.id === draft.id ? { ...r, clearedTxnIds: r.clearedTxnIds.includes(txnId) ? r.clearedTxnIds.filter(x => x !== txnId) : [...r.clearedTxnIds, txnId] } : r) }));
  const markAllCleared = () => setData(d => ({ ...d, reconciliations: d.reconciliations.map(r => r.id === draft.id ? { ...r, clearedTxnIds: [...new Set(lines.map(l => l.txnId))] } : r) }));

  const clearedBalance = draft ? lines.filter(l => draft.clearedTxnIds.includes(l.txnId)).reduce((s, l) => s + l.debit - l.credit, 0) : 0;
  const difference = draft ? draft.statementBalance - clearedBalance : 0;
  const balanced = draft && Math.abs(difference) < 0.5;

  const finish = async () => {
    if (!balanced) return notify("The difference must be zero before finishing");
    if (!(await confirm(`Finish reconciling ${bank.name} as of ${fmtDate(draft.date)}? The ${draft.clearedTxnIds.length} cleared line(s) will be locked to this reconciliation.`))) return;
    setData(d => ({ ...d, reconciliations: d.reconciliations.map(r => r.id === draft.id ? { ...r, completed: true, completedAt: todayStr() } : r) }));
    notify("Reconciliation completed");
  };

  return (
    <div>
      {!draft ? (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Start a reconciliation</div>
          <div style={{ fontSize: 12.5, color: theme.muted, marginTop: 4 }}>Enter the ending balance and date from your bank statement, then tick off every line that appears on it.</div>
          <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <input type="number" style={{ ...styles.input, width: 180 }} placeholder="Statement ending balance" value={form.statementBalance} onChange={e => setForm({ ...form, statementBalance: e.target.value })} />
            <button style={styles.btnPrimary} onClick={startReconciliation}>Start reconciliation</button>
          </div>
        </div>
      ) : (
        <div className="ces-card" style={styles.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <div style={styles.cardTitle}>Reconciling as of {fmtDate(draft.date)}</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button style={styles.btnGhost} onClick={markAllCleared}>Mark all cleared</button>
              <button style={styles.btnGhost} onClick={cancelReconciliation}>Discard</button>
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginTop: 12 }}>
            <Kpi label="Statement balance" value={fmt(draft.statementBalance)} />
            <Kpi label="Cleared balance" value={fmt(clearedBalance)} />
            <Kpi label="Difference" value={fmt(difference)} tone={balanced ? "emerald" : "rose"} />
          </div>
          <table style={{ ...styles.table, marginTop: 14 }}>
            <thead><tr><th style={{ ...styles.th, width: 30 }}></th><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
            <tbody>{lines.map((l, i) => (
              <tr key={i} style={{ cursor: "pointer" }} onClick={() => toggleCleared(l.txnId)}>
                <td style={styles.td}><input type="checkbox" checked={draft.clearedTxnIds.includes(l.txnId)} onChange={() => toggleCleared(l.txnId)} onClick={e => e.stopPropagation()} /></td>
                <td style={styles.tdMono}>{fmtDate(l.date)}</td>
                <td style={styles.td}>{l.memo}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", color: l.debit ? theme.emerald : theme.rose }}>{fmt(l.debit - l.credit)}</td>
              </tr>
            ))}</tbody>
          </table>
          {lines.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No unreconciled lines on or before this date.</div>}
          <button style={{ ...styles.btnPrimary, marginTop: 14, opacity: balanced ? 1 : 0.5 }} disabled={!balanced} onClick={finish}>Finish reconciliation</button>
        </div>
      )}
      {history.length > 0 && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Reconciliation history</div>
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}>As of</th><th style={{ ...styles.th, textAlign: "right" }}>Statement balance</th><th style={{ ...styles.th, textAlign: "right" }}>Lines cleared</th><th style={styles.th}>Completed</th></tr></thead>
            <tbody>{history.map(r => (
              <tr key={r.id}><td style={styles.tdMono}>{fmtDate(r.date)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.statementBalance)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.clearedTxnIds.length}</td><td style={styles.tdMono}>{fmtDate(r.completedAt)}</td></tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}
function BankMiniTrend({ data, accountId }) {
  const { theme } = useUI();
  const points = last6Months().map(m => computeBalances(data.accounts, data.transactions, m.to)[accountId] || 0);
  const max = Math.max(...points.map(Math.abs), 1);
  return (
    <svg viewBox="0 0 220 46" width="100%" height="46" style={{ marginTop: 10 }}>
      <polyline fill="none" stroke={theme.accent} strokeWidth="2"
        points={points.map((p, i) => `${(i / 5) * 210 + 5},${40 - (Math.abs(p) / max) * 32}`).join(" ")} />
    </svg>
  );
}

/* =================================== Tax Stack (shared) =================================== */
function TaxStack({ taxes, setTaxes, subtotal, docType, accounts, label }) {
  const { styles, fmt, theme } = useUI();
  const taxAccounts = accounts.filter(a => (a.type === "liability" || a.type === "asset") && a.status !== "inactive");
  const update = (i, field, val) => setTaxes(ts => ts.map((t, idx) => idx === i ? { ...t, [field]: val } : t));
  const remove = (i) => setTaxes(ts => ts.filter((_, idx) => idx !== i));
  const move = (i, dir) => setTaxes(ts => { const arr = [...ts]; const j = i + dir; if (j < 0 || j >= arr.length) return ts; [arr[i], arr[j]] = [arr[j], arr[i]]; return arr; });
  const addPreset = (preset) => setTaxes(ts => [...ts, { ...preset, id: uid("tax") }]);
  const { steps, finalAmount } = applyCascadingTaxes(subtotal, taxes);

  return (
    <div style={{ marginTop: 14 }}>
      <div style={{ fontSize: 12, fontWeight: 600, color: theme.muted, textTransform: "uppercase", letterSpacing: "0.04em" }}>{label || "Cascading taxes"}</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 2 }}>Each tax here is calculated on the running amount left after the tax before it. Choose add (adds to total) or deduct (withheld from total).</div>
      <table style={{ ...styles.table, marginTop: 10 }}>
        <thead><tr><th style={styles.th}>Tax</th><th style={{ ...styles.th, textAlign: "right" }}>Rate</th><th style={styles.th}>Mode</th><th style={styles.th}>Effect</th><th style={styles.th}>Posts to</th><th></th></tr></thead>
        <tbody>
          {taxes.map((t, i) => (
            <tr key={t.id}>
              <td style={styles.td}><input style={{ ...styles.inputSmall, width: 110 }} value={t.name} onChange={e => update(i, "name", e.target.value)} /></td>
              <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={t.rate} onChange={e => update(i, "rate", e.target.value)} /></td>
              <td style={styles.td}><select style={styles.inputSmall} value={t.mode} onChange={e => update(i, "mode", e.target.value)}><option value="percent">%</option><option value="fixed">flat</option></select></td>
              <td style={styles.td}><select style={styles.inputSmall} value={t.effect} onChange={e => update(i, "effect", e.target.value)}><option value="deduct">deduct</option><option value="add">add</option></select></td>
              <td style={styles.td}><select style={styles.inputSmall} value={t.accountId} onChange={e => update(i, "accountId", e.target.value)}>{taxAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select></td>
              <td style={styles.td}>
                <button style={styles.iconBtn} onClick={() => move(i, -1)}>↑</button>
                <button style={styles.iconBtn} onClick={() => move(i, 1)}>↓</button>
                <button style={styles.iconBtn} onClick={() => remove(i)}>✕</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
        {TAX_PRESETS.map((p, i) => <button key={i} style={styles.btnGhost} onClick={() => addPreset(p)}>+ {p.name}</button>)}
        <button style={styles.btnGhost} onClick={() => addPreset({ name: "Custom", rate: 0, mode: "percent", effect: "deduct", accountId: taxAccounts[0]?.id })}>+ Custom tax</button>
      </div>

      <div className="ces-card" style={{ ...styles.card, background: theme.panel2, marginTop: 12, padding: "12px 16px" }}>
        <RowLine label="Subtotal" value={subtotal} bold />
        {steps.map((s, i) => <RowLine key={i} label={`${s.name} (${s.mode === "fixed" ? fmt(s.rate) : s.rate + "%"} ${s.effect})`} value={s.effect === "deduct" ? -s.amount : s.amount} indent tone={s.effect === "deduct" ? "rose" : "emerald"} />)}
        <RowLine label="Final amount" value={finalAmount} bold divider />
      </div>
    </div>
  );
}
function RowLine({ label, value, bold, indent, divider, tone, onClick }) {
  const { theme, fmt } = useUI();
  const c = tone === "rose" ? theme.rose : tone === "emerald" ? theme.emerald : theme.text;
  return (
    <div
      onClick={onClick}
      className={onClick ? "drill-row" : undefined}
      style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderTop: divider ? `1px solid ${theme.border}` : "none", marginTop: divider ? 4 : 0, cursor: onClick ? "pointer" : "default", borderRadius: 4 }}
    >
      <span style={{ paddingLeft: indent ? 14 : 0, fontWeight: bold ? 600 : 400, fontSize: 13 }}>{label}</span>
      {value !== undefined && value !== null && <span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13, fontWeight: bold ? 600 : 400, color: c }}>{fmt(value)}</span>}
    </div>
  );
}

/* =================================== Invoices =================================== */
// Shared helpers for creating/editing invoices & bills ----------------------
function invoiceJournalLines(form, data) {
  const t = computeDocTotals(form, data.taxGroups);
  const cogsByAccount = {};
  form.items.filter(it => it.inventoryId).forEach(it => {
    const inv = data.inventory.find(x => x.id === it.inventoryId);
    if (!inv) return;
    const acct = inventoryAccountForItem(data, it.inventoryId);
    cogsByAccount[acct] = (cogsByAccount[acct] || 0) + inv.unitCost * Number(it.qty || 0);
  });
  const totalCogs = Object.values(cogsByAccount).reduce((s, v) => s + v, 0);
  return [
    { accountId: "1100", debit: t.finalAmount, credit: 0 },
    { accountId: "4000", debit: 0, credit: t.subtotal },
    ...t.lineCalcs.flatMap(l => taxJournalLines(l.components, "invoice")),
    ...taxJournalLines(t.cascadeSteps, "invoice"),
    ...(totalCogs > 0 ? [{ accountId: "5000", debit: totalCogs, credit: 0 }, ...Object.entries(cogsByAccount).map(([acct, amt]) => ({ accountId: acct, debit: 0, credit: amt }))] : []),
  ];
}
function snapshotTaxItems(items, taxGroups) {
  return items.map(it => it.taxGroupId ? { ...it, taxComponents: (taxGroups.find(g => g.id === it.taxGroupId)?.components || []).map(c => ({ ...c })) } : { ...it, taxComponents: undefined });
}

// IFRS 15: reclassifies an already-posted invoice's revenue into Deferred
// Revenue, then recognizes it back into Sales Revenue straight-line over the
// service period. Deliberately built as a reclassification on top of a
// normal invoice (rather than changing how invoices post) so the invoice
// creation flow itself never needs to change.
function DeferredRevenuePanel({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const blank = () => ({ invoiceId: "", months: 12, startDate: todayStr() });
  const [form, setForm] = useState(blank());
  const currentMonth = monthKey(todayStr());

  const deferredInvoiceIds = new Set(data.deferredRevenueSchedules.filter(s => s.status !== "cancelled").map(s => s.invoiceId));
  const eligibleInvoices = data.invoices.filter(inv => !deferredInvoiceIds.has(inv.id));

  const createSchedule = () => {
    const inv = data.invoices.find(i => i.id === form.invoiceId);
    if (!inv) return notify("Choose an invoice to defer");
    if (!Number(form.months) || Number(form.months) < 1) return notify("Enter how many months the service runs for");
    const docTxn = data.transactions.find(t => t.docId === inv.id && t.source === "invoice") || data.transactions.find(t => t.source === "invoice" && (t.memo || "").includes(inv.id));
    if (!docTxn) return notify("Could not find this invoice's original journal entry");
    const revenueAmount = docTxn.lines.filter(l => { const a = data.accounts.find(x => x.id === l.accountId); return a && a.type === "revenue" && a.subtype !== "other"; }).reduce((s, l) => s + (l.credit - l.debit), 0);
    if (revenueAmount <= 0.5) return notify("This invoice has no recognizable revenue to defer");
    const txn = buildTxn(`Defer revenue - ${inv.id}`, form.startDate, [{ accountId: "4000", debit: revenueAmount, credit: 0 }, { accountId: "2290", debit: 0, credit: revenueAmount }], "deferred-revenue");
    if (!txn) return notify("Entry not balanced");
    const scheduleId = uid("defrev");
    txn.meta = { scheduleId, kind: "defer" };
    setData(d => ({
      ...d,
      transactions: [...d.transactions, txn],
      deferredRevenueSchedules: [...d.deferredRevenueSchedules, { id: scheduleId, invoiceId: inv.id, customer: inv.customer, totalAmount: revenueAmount, startDate: form.startDate, months: Number(form.months), recognizedAmount: 0, periodsRun: 0, lastRunMonth: null, status: "active" }],
    }));
    setForm(blank());
    notify(`${fmt(revenueAmount)} moved to Deferred Revenue - will recognize over ${form.months} months`);
  };

  const dueSchedules = data.deferredRevenueSchedules.filter(s => s.status === "active" && s.lastRunMonth !== currentMonth && s.recognizedAmount < s.totalAmount - 0.5);
  const runDue = () => {
    if (dueSchedules.length === 0) return notify("No deferred revenue due for recognition this month");
    let working = data;
    dueSchedules.forEach(s => {
      const { dueAmount: rawDue } = computeDeferredRevenueMovement(s);
      const dueAmount = Math.round(rawDue);
      if (dueAmount <= 0.5) return;
      const txn = buildTxn(`Revenue recognition ${s.periodsRun + 1}/${s.months} - ${s.customer}`, todayStr(), [{ accountId: "2290", debit: dueAmount, credit: 0 }, { accountId: "4000", debit: 0, credit: dueAmount }], "deferred-revenue");
      if (!txn) return;
      txn.meta = { scheduleId: s.id, kind: "recognition" };
      const newRecognized = s.recognizedAmount + dueAmount;
      working = {
        ...working,
        transactions: [...working.transactions, txn],
        deferredRevenueSchedules: working.deferredRevenueSchedules.map(x => x.id === s.id ? { ...x, recognizedAmount: newRecognized, periodsRun: x.periodsRun + 1, lastRunMonth: currentMonth, status: newRecognized >= x.totalAmount - 0.5 ? "completed" : "active" } : x),
      };
    });
    setData(working);
    notify(`Recognized revenue for ${dueSchedules.length} schedule(s)`);
  };

  const deleteSchedule = async (s) => {
    if (!(await confirm(`Delete this deferred revenue schedule for ${s.customer}? All ${s.periodsRun + 1} of its journal entries (the original deferral plus every recognition run) are removed - the invoice's revenue effectively reverts to being recognized immediately, as it was before deferring.`))) return;
    setData(d => ({ ...d, deferredRevenueSchedules: d.deferredRevenueSchedules.filter(x => x.id !== s.id), transactions: d.transactions.filter(t => !(t.meta && t.meta.scheduleId === s.id)) }));
    notify("Schedule and all its journal entries deleted");
  };

  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <div style={styles.cardTitle}>Deferred revenue</div>
        {dueSchedules.length > 0 && <button style={styles.btnPrimary} onClick={runDue}>Recognize {dueSchedules.length} due</button>}
      </div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>For invoices that actually cover a multi-month service (subscriptions, retainers), move the revenue into Deferred Revenue and recognize it straight-line over the service period instead of all at once on the invoice date.</div>
      {data.deferredRevenueSchedules.length > 0 && (
        <div style={{ overflowX: "auto" }}>
        <table style={{ ...styles.table, marginTop: 10 }}>
          <thead><tr><th style={styles.th}>Invoice</th><th style={styles.th}>Customer</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={{ ...styles.th, textAlign: "right" }}>Recognized</th><th style={{ ...styles.th, textAlign: "right" }}>Remaining</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
          <tbody>{data.deferredRevenueSchedules.map(s => (
            <tr key={s.id}><td style={styles.tdMono}>{s.invoiceId}</td><td style={styles.td}>{s.customer}</td>
              <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.totalAmount)}</td>
              <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.recognizedAmount)}</td>
              <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.totalAmount - s.recognizedAmount)}</td>
              <td style={styles.td}><span style={{ ...styles.pill, ...(s.status === "completed" ? styles.pillGreen : styles.pillAmber) }}>{s.status} ({s.periodsRun}/{s.months}mo)</span></td>
              <td style={styles.td}><button style={styles.iconBtn} onClick={() => deleteSchedule(s)}>🗑</button></td></tr>
          ))}</tbody>
        </table>
        </div>
      )}
      <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
        <select style={{ ...styles.input, flex: 1, minWidth: 200 }} value={form.invoiceId} onChange={e => setForm({ ...form, invoiceId: e.target.value })}>
          <option value="">Choose an invoice to defer...</option>
          {eligibleInvoices.map(inv => <option key={inv.id} value={inv.id}>{inv.id} - {inv.customer}</option>)}
        </select>
        <input type="date" style={styles.input} value={form.startDate} onChange={e => setForm({ ...form, startDate: e.target.value })} />
        <input type="number" style={{ ...styles.input, width: 130 }} placeholder="Service months" value={form.months} onChange={e => setForm({ ...form, months: e.target.value })} />
        <button style={styles.btnPrimary} onClick={createSchedule}>Defer this invoice's revenue</button>
      </div>
    </div>
  );
}
function Invoices({ data, setData, postTransaction, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const blank = () => ({ customer: "", date: todayStr(), dueDate: "", projectId: "", salesperson: "", items: [{ desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }], taxes: [] });
  const [form, setForm] = useState(blank());

  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }] }));
  const totals = computeDocTotals(form, data.taxGroups);

  const startEdit = (inv) => {
    if (inv.locked) return notify("This invoice is locked - unlock it first to edit");
    setForm({ customer: inv.customer, date: inv.date, dueDate: inv.dueDate || "", projectId: inv.projectId || "", salesperson: inv.salesperson || "", items: inv.items.map(it => ({ ...it, taxComponents: undefined })), taxes: (inv.taxes || []).map(t => ({ ...t })) });
    setEditingId(inv.id); setViewingId(null); setShowForm(true);
  };

  const saveInvoice = () => {
    if (!form.customer || form.items.some(it => !it.desc)) return notify("Add a customer and describe each line item");
    const isEdit = Boolean(editingId);
    const id = isEdit ? editingId : `INV-${data.nextInvoiceNum}`;
    let working = { ...data };
    // If editing: remove the old journal entry and put the old stock back first,
    // so the new posting starts from a clean slate.
    if (isEdit) {
      const old = data.invoices.find(i => i.id === id);
      working = { ...working, transactions: stripDocTransactions(working.transactions, id) };
      let restoredLots = working.inventoryLots;
      (old?.items || []).forEach(it => {
        if (!it.inventoryId) return;
        working = adjustInventory(working, it.inventoryId, Number(it.qty || 0));
        restoredLots = restoreFIFO(restoredLots, it.inventoryId, Number(it.qty || 0));
      });
      working = { ...working, inventoryLots: restoredLots };
    }
    const lines = invoiceJournalLines(form, working);
    const txn = buildTxn(`Invoice ${id} - ${form.customer}`, form.date, lines, "invoice", id);
    if (!txn) return notify("Entry not balanced - check tax accounts");
    working = { ...working, transactions: [...working.transactions, txn] };
    const snapshotItems = snapshotTaxItems(form.items, data.taxGroups);
    const docBase = { ...form, items: snapshotItems, id };
    working = isEdit
      ? { ...working, invoices: working.invoices.map(i => i.id === id ? { ...i, ...docBase } : i) }
      : { ...working, invoices: [...working.invoices, { ...docBase, status: "sent", amountPaid: 0, locked: false }], nextInvoiceNum: working.nextInvoiceNum + 1 };
    let lots = working.inventoryLots;
    form.items.forEach(it => {
      if (!it.inventoryId) return;
      working = adjustInventory(working, it.inventoryId, -Number(it.qty || 0));
      const r = consumeFIFO(lots, it.inventoryId, Number(it.qty || 0)); lots = r.lots;
    });
    working = { ...working, inventoryLots: lots };
    setData(working);
    setForm(blank()); setShowForm(false); setEditingId(null);
    notify(isEdit ? `${id} updated and reposted to the ledger` : `${id} created and posted to the ledger`);
  };

  const toggleLock = (id) => setData(d => ({ ...d, invoices: d.invoices.map(i => i.id === id ? { ...i, locked: !i.locked } : i) }));

  // Delete an invoice with the right approach: journal reversed, inventory
  // restored (weighted-avg qty and FIFO lots). Payments must be deleted first
  // so no cash entry is left pointing at a document that no longer exists.
  const deleteInvoice = async (inv) => {
    if (inv.locked) return notify("This invoice is locked - unlock it first");
    if (isLocked(inv.date, data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - this invoice falls within that period and can't be deleted`);
    if ((inv.amountPaid || 0) > 0) return notify("This invoice has payments recorded against it - delete those payments (Payments tab) before deleting the invoice");
    if (!(await confirm(`Delete ${inv.id}? Its journal entry is removed and any inventory it shipped is restored. This can't be undone.`))) return;
    setData(d => {
      let next = { ...d, transactions: stripDocTransactions(d.transactions, inv.id) };
      let lots = next.inventoryLots;
      (inv.items || []).forEach(it => {
        if (!it.inventoryId) return;
        next = adjustInventory(next, it.inventoryId, Number(it.qty || 0));
        lots = restoreFIFO(lots, it.inventoryId, Number(it.qty || 0));
      });
      return { ...next, inventoryLots: lots, invoices: next.invoices.filter(i => i.id !== inv.id) };
    });
    setViewingId(null);
    notify(`${inv.id} deleted - ledger and stock restored`);
  };

  const sorted = [...data.invoices].sort((a, b) => b.date.localeCompare(a.date));
  const viewing = viewingId ? data.invoices.find(i => i.id === viewingId) : null;
  if (viewing) return <DocumentDetail doc={viewing} docType="invoice" data={data} onBack={() => setViewingId(null)} onEdit={() => startEdit(viewing)} onToggleLock={() => toggleLock(viewing.id)} onDelete={() => deleteInvoice(viewing)} />;

  return (
    <div>
      <PageHeader eyebrow="Sales" title="Invoices" sub={`${data.invoices.length} total`} action={<button style={styles.btnPrimary} onClick={() => { setShowForm(s => !s); setEditingId(null); setForm(blank()); }}>{showForm ? "Cancel" : "+ New invoice"}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>{editingId ? `Edit ${editingId}` : "New invoice"}</div>
          {editingId && <div style={{ fontSize: 12, color: theme.amber, marginTop: 4 }}>Saving will remove the original journal entry and stock movement, then repost with the new values. Payments already received stay attached.</div>}
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Customer name" value={form.customer} onChange={e => setForm({ ...form, customer: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <input type="date" style={styles.input} value={form.dueDate} onChange={e => setForm({ ...form, dueDate: e.target.value })} placeholder="Due date" />
            <select style={styles.input} value={form.projectId} onChange={e => setForm({ ...form, projectId: e.target.value })}><option value="">No project</option>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
            <input style={{ ...styles.input, width: 160 }} placeholder="Salesperson (optional)" value={form.salesperson || ""} onChange={e => setForm({ ...form, salesperson: e.target.value })} />
          </div>
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Description</th><th style={styles.th}>Inventory item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Price</th><th style={styles.th}>Tax group</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: "100%" }} value={it.desc} onChange={e => updateItem(i, "desc", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.salePrice); } }}>
                  <option value="">None</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name} ({inv.qty} in stock)</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.taxGroupId} onChange={e => updateItem(i, "taxGroupId", e.target.value)}>
                  <option value="">No tax group</option>{data.taxGroups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}</select></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(totals.lineCalcs[i]?.lineTotal || 0)}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <LineTaxSummary lineCalcs={totals.lineCalcs} subtotal={totals.subtotal} afterLineTax={totals.afterLineTax} />
          <TaxStack taxes={form.taxes} setTaxes={(fn) => setForm(f => ({ ...f, taxes: typeof fn === "function" ? fn(f.taxes) : fn }))} subtotal={totals.afterLineTax} docType="invoice" accounts={data.accounts} label="Document-level taxes (applied to the invoice total)" />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={saveInvoice}>{editingId ? "Save changes & repost" : "Create & post to ledger"}</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Invoice</th><th style={styles.th}>Customer</th><th style={styles.th}>Due</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={{ ...styles.th, textAlign: "right" }}>Paid</th><th style={styles.th}>Status</th></tr></thead>
          <tbody>{sorted.map(inv => {
            const total = computeDocTotals(inv, data.taxGroups).finalAmount;
            const overdue = inv.status !== "paid" && inv.dueDate && inv.dueDate < todayStr();
            const status = inv.amountPaid >= total - 0.5 ? "paid" : inv.amountPaid > 0 ? "partial" : overdue ? "overdue" : "sent";
            return (
              <tr key={inv.id} onClick={() => setViewingId(inv.id)} style={{ cursor: "pointer" }}>
                <td style={styles.tdMono}>{inv.locked ? "🔒 " : ""}{inv.id}</td><td style={styles.td}>{inv.customer}</td><td style={styles.tdMono}>{inv.dueDate ? fmtDate(inv.dueDate) : "-"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(total)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(inv.amountPaid || 0)}</td>
                <td style={styles.td}><span style={{ ...styles.pill, ...(status === "paid" ? styles.pillGreen : status === "overdue" ? styles.pillRose : styles.pillAmber) }}>{status}</span></td></tr>
            );
          })}</tbody>
        </table>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>Click an invoice to view, edit, lock or print it. Record collections from the Payments tab.</div>
      </div>
      <DeferredRevenuePanel data={data} setData={setData} notify={notify} />
    </div>
  );
}
// Printable, client-ready document view - shared by invoices and bills.
function DocumentDetail({ doc, docType, data, onBack, onEdit, onToggleLock, onDelete }) {
  const { theme, styles, fmt } = useUI();
  const totals = computeDocTotals(doc, data.taxGroups);
  const isInvoice = docType === "invoice";
  const party = isInvoice ? doc.customer : doc.vendor;
  const overdue = doc.status !== "paid" && doc.dueDate && doc.dueDate < todayStr();
  const status = doc.amountPaid >= totals.finalAmount - 0.5 ? "paid" : doc.amountPaid > 0 ? "partial" : overdue ? "overdue" : (isInvoice ? "sent" : "unpaid");
  const logo = data.settings.logo;
  return (
    <div>
      <div className="no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 8 }}>
        <button style={styles.btnGhost} onClick={onBack}>← Back to {isInvoice ? "invoices" : "bills"}</button>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={styles.btnGhost} onClick={onToggleLock}>{doc.locked ? "🔓 Unlock" : "🔒 Lock"}</button>
          <button style={{ ...styles.btnGhost, color: theme.rose, opacity: doc.locked ? 0.5 : 1 }} onClick={onDelete}>🗑 Delete</button>
          <button style={{ ...styles.btnGhost, opacity: doc.locked ? 0.5 : 1 }} onClick={onEdit}>✎ Edit</button>
          <button style={styles.btnPrimary} onClick={() => window.print()}>Print / Save as PDF</button>
        </div>
      </div>
      {doc.locked && <div className="no-print" style={{ fontSize: 12, color: theme.amber, marginBottom: 12, textAlign: "center" }}>This {isInvoice ? "invoice" : "bill"} is locked - unlock it to make changes.</div>}
      <div className="print-area" style={{ background: "#FFFFFF", color: "#1B2559", border: `1px solid ${theme.border}`, borderRadius: 12, padding: "40px 44px", maxWidth: 780, margin: "0 auto", boxShadow: "0 1px 3px rgba(16,24,40,0.08)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${theme.accent}`, paddingBottom: 20 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            {logo && <div className="ces-logo" style={{ padding: 8 }}><img src={logo} alt="logo" style={{ height: 46, width: "auto", maxWidth: 120, objectFit: "contain" }} /></div>}
            <div>
              <div style={{ fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontSize: 24, fontWeight: 600, color: theme.ink }}>{data.settings.companyName}</div>
              <div style={{ fontSize: 12, color: "#6B7690", marginTop: 4 }}>{isInvoice ? "Invoice" : "Bill"}</div>
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 20, fontWeight: 600 }}>{doc.id}</div>
            <span style={{ ...styles.pill, ...(status === "paid" ? styles.pillGreen : status === "overdue" ? styles.pillRose : styles.pillAmber), marginTop: 6, display: "inline-block" }}>{status}</span>
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 24 }}>
          <div>
            <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#6B7690" }}>{isInvoice ? "Billed to" : "Vendor"}</div>
            <div style={{ fontSize: 15, fontWeight: 600, marginTop: 4 }}>{party}</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 12, color: "#6B7690" }}>{isInvoice ? "Invoice" : "Bill"} date <span style={{ color: "#1B2559", fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmtDate(doc.date)}</span></div>
            <div style={{ fontSize: 12, color: "#6B7690", marginTop: 3 }}>Due date <span style={{ color: "#1B2559", fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{doc.dueDate ? fmtDate(doc.dueDate) : "-"}</span></div>
          </div>
        </div>
        <table style={{ width: "100%", marginTop: 28, borderCollapse: "collapse", fontSize: 13.5 }}>
          <thead><tr style={{ borderBottom: "2px solid #1B2559" }}>
            <th style={{ textAlign: "left", padding: "6px 4px", fontSize: 11, textTransform: "uppercase", color: "#6B7690" }}>Description</th>
            <th style={{ textAlign: "right", padding: "6px 4px", fontSize: 11, textTransform: "uppercase", color: "#6B7690" }}>Qty</th>
            <th style={{ textAlign: "right", padding: "6px 4px", fontSize: 11, textTransform: "uppercase", color: "#6B7690" }}>{isInvoice ? "Price" : "Cost"}</th>
            <th style={{ textAlign: "right", padding: "6px 4px", fontSize: 11, textTransform: "uppercase", color: "#6B7690" }}>Amount</th>
          </tr></thead>
          <tbody>{totals.lineCalcs.map((l, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #E3E8F0" }}>
              <td style={{ padding: "8px 4px" }}>{l.item.desc}</td>
              <td style={{ padding: "8px 4px", textAlign: "right", fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{l.item.qty}</td>
              <td style={{ padding: "8px 4px", textAlign: "right", fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(l.item.price)}</td>
              <td style={{ padding: "8px 4px", textAlign: "right", fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(l.lineTotal)}</td>
            </tr>
          ))}</tbody>
        </table>
        <div style={{ marginTop: 18, marginLeft: "auto", width: 260 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}><span>Subtotal</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(totals.subtotal)}</span></div>
          {totals.lineCalcs.flatMap(l => l.components).map((c, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: "#6B7690", padding: "2px 0" }}><span>{c.name}</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{c.effect === "deduct" ? "−" : "+"}{fmt(c.amount)}</span></div>
          ))}
          {totals.cascadeSteps.map((s, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: "#6B7690", padding: "2px 0" }}><span>{s.name}</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{s.effect === "deduct" ? "−" : "+"}{fmt(s.amount)}</span></div>
          ))}
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 16, fontWeight: 700, borderTop: "2px solid #1B2559", marginTop: 6, paddingTop: 8 }}><span>{isInvoice ? "Total due" : "Total payable"}</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(totals.finalAmount)}</span></div>
          {doc.amountPaid > 0 && <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: theme.emerald, marginTop: 6 }}><span>Paid</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(doc.amountPaid)}</span></div>}
        </div>
        <div style={{ marginTop: 36, fontSize: 11, color: "#9AA5C0", textAlign: "center", borderTop: "1px solid #E3E8F0", paddingTop: 14 }}>{isInvoice ? "Thank you for your business - " : ""}generated by {data.settings.companyName}</div>
      </div>
    </div>
  );
}
function adjustInventory(data, itemId, deltaQty, newUnitCost = null) {
  return { ...data, inventory: data.inventory.map(it => {
    if (it.id !== itemId) return it;
    if (newUnitCost !== null && deltaQty > 0) {
      const totalCost = it.qty * it.unitCost + deltaQty * newUnitCost;
      const newQty = it.qty + deltaQty;
      return { ...it, qty: newQty, unitCost: newQty > 0 ? totalCost / newQty : newUnitCost };
    }
    return { ...it, qty: Math.max(0, it.qty + deltaQty) };
  }) };
}
// Shows how much each tax GROUP added/deducted per line item, combined,
// before any document-level cascading tax is applied on top.
function LineTaxSummary({ lineCalcs, subtotal, afterLineTax }) {
  const { theme } = useUI();
  const totalsByName = {};
  lineCalcs.forEach(l => l.components.forEach(c => { totalsByName[c.name] = (totalsByName[c.name] || 0) + (c.effect === "deduct" ? -c.amount : c.amount); }));
  const names = Object.keys(totalsByName);
  if (names.length === 0) return null;
  return (
    <div style={{ marginTop: 14 }}>
      <div style={{ fontSize: 12, fontWeight: 600, color: theme.muted, textTransform: "uppercase", letterSpacing: "0.04em" }}>Line-item group taxes</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 2 }}>Every tax in a line's group is calculated on that line's own amount, independently of each other, then combined - not cascaded.</div>
      <div style={{ marginTop: 8 }}>
        <RowLine label="Line items subtotal" value={subtotal} bold />
        {names.map(n => <RowLine key={n} label={n} value={totalsByName[n]} indent tone={totalsByName[n] < 0 ? "rose" : "emerald"} />)}
        <RowLine label="Subtotal after line taxes" value={afterLineTax} bold divider />
      </div>
    </div>
  );
}

/* =================================== Bills =================================== */
// Sales receipt: identical to an invoice's posting except the debit side is
// cash received immediately instead of Accounts Receivable - there's no
// unpaid state to track.
function salesReceiptJournalLines(form, data, bankId) {
  const t = computeDocTotals(form, data.taxGroups);
  const bank = data.banks.find(b => b.id === bankId);
  const cogsByAccount = {};
  form.items.filter(it => it.inventoryId).forEach(it => {
    const inv = data.inventory.find(x => x.id === it.inventoryId);
    if (!inv) return;
    const acct = inventoryAccountForItem(data, it.inventoryId);
    cogsByAccount[acct] = (cogsByAccount[acct] || 0) + inv.unitCost * Number(it.qty || 0);
  });
  const totalCogs = Object.values(cogsByAccount).reduce((s, v) => s + v, 0);
  return [
    { accountId: bank ? bank.accountId : "1000", debit: t.finalAmount, credit: 0 },
    { accountId: "4000", debit: 0, credit: t.subtotal },
    ...t.lineCalcs.flatMap(l => taxJournalLines(l.components, "invoice")),
    ...taxJournalLines(t.cascadeSteps, "invoice"),
    ...(totalCogs > 0 ? [{ accountId: "5000", debit: totalCogs, credit: 0 }, ...Object.entries(cogsByAccount).map(([acct, amt]) => ({ accountId: acct, debit: 0, credit: amt }))] : []),
  ];
}
// Credit notes and sales returns share this: the exact mirror image of
// invoiceJournalLines. Reusing taxJournalLines with docType "bill" instead
// of "invoice" on the same tax steps always produces the reversed debit/
// credit direction (the two docTypes are defined as opposites of each other
// for every tax effect), so this doesn't need to hand-flip anything. When
// items reference inventory (a physical return), the goods are restocked at
// current cost and the matching COGS is reversed; a pure credit note (no
// inventory items) skips that part entirely.
function creditNoteJournalLines(form, data, refundBankId) {
  const t = computeDocTotals(form, data.taxGroups);
  const restockByAccount = {};
  form.items.filter(it => it.inventoryId).forEach(it => {
    const inv = data.inventory.find(x => x.id === it.inventoryId);
    if (!inv) return;
    const acct = inventoryAccountForItem(data, it.inventoryId);
    restockByAccount[acct] = (restockByAccount[acct] || 0) + inv.unitCost * Number(it.qty || 0);
  });
  const totalRestock = Object.values(restockByAccount).reduce((s, v) => s + v, 0);
  const bank = refundBankId ? data.banks.find(b => b.id === refundBankId) : null;
  const creditAccountId = bank ? bank.accountId : "1100";
  return [
    { accountId: "4200", debit: t.subtotal, credit: 0 },
    { accountId: creditAccountId, debit: 0, credit: t.finalAmount },
    ...t.lineCalcs.flatMap(l => taxJournalLines(l.components, "bill")),
    ...taxJournalLines(t.cascadeSteps, "bill"),
    ...(totalRestock > 0 ? [...Object.entries(restockByAccount).map(([acct, amt]) => ({ accountId: acct, debit: amt, credit: 0 })), { accountId: "5000", debit: 0, credit: totalRestock }] : []),
  ];
}
/* =================================== Sales Orders =================================== */
// A customer's order before it becomes a real sale - purely a tracking
// document with no ledger impact until it's converted to an invoice.
function SalesOrders({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const blank = () => ({ customer: "", date: todayStr(), expectedDate: "", projectId: "", items: [{ desc: "", qty: 1, price: 0, inventoryId: "" }] });
  const [form, setForm] = useState(blank());
  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "" }] }));
  const total = form.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0);

  const createOrder = () => {
    if (!form.customer || form.items.some(it => !it.desc)) return notify("Add a customer and describe each line item");
    const id = `SO-${data.nextSalesOrderNum}`;
    setData(d => ({ ...d, salesOrders: [...d.salesOrders, { ...form, id, status: "open" }], nextSalesOrderNum: d.nextSalesOrderNum + 1 }));
    setForm(blank()); setShowForm(false);
    notify(`${id} created - no ledger impact until converted to an invoice`);
  };

  const convertToInvoice = async (so) => {
    if (!(await confirm(`Convert ${so.id} to an invoice for ${so.customer}? This posts a real invoice with the same line items to the ledger.`))) return;
    const invForm = { customer: so.customer, date: todayStr(), dueDate: "", projectId: so.projectId, items: so.items, taxes: [] };
    const lines = invoiceJournalLines(invForm, data);
    const id = `INV-${data.nextInvoiceNum}`;
    const txn = buildTxn(`Invoice ${id} - ${so.customer}`, invForm.date, lines, "invoice", id);
    if (!txn) return notify("Entry not balanced");
    let newData = {
      ...data,
      transactions: [...data.transactions, txn],
      invoices: [...data.invoices, { ...invForm, items: snapshotTaxItems(invForm.items, data.taxGroups), id, status: "sent", amountPaid: 0, locked: false }],
      nextInvoiceNum: data.nextInvoiceNum + 1,
      salesOrders: data.salesOrders.map(x => x.id === so.id ? { ...x, status: "fulfilled" } : x),
    };
    let lots = newData.inventoryLots;
    so.items.forEach(it => {
      if (!it.inventoryId) return;
      newData = adjustInventory(newData, it.inventoryId, -Number(it.qty || 0));
      const r = consumeFIFO(lots, it.inventoryId, Number(it.qty || 0));
      lots = r.lots;
    });
    setData({ ...newData, inventoryLots: lots });
    notify(`${id} created from ${so.id} and posted to the ledger`);
  };

  const cancelOrder = async (so) => {
    if (!(await confirm(`Cancel ${so.id}? It stays on record but won't be available to convert.`))) return;
    setData(d => ({ ...d, salesOrders: d.salesOrders.map(x => x.id === so.id ? { ...x, status: "cancelled" } : x) }));
    notify("Sales order cancelled");
  };

  const deleteOrder = async (so) => {
    if (!(await confirm(`Delete ${so.id} entirely? It never posted to the ledger, so this is a clean removal.`))) return;
    setData(d => ({ ...d, salesOrders: d.salesOrders.filter(x => x.id !== so.id) }));
    notify("Sales order deleted");
  };

  const sorted = [...data.salesOrders].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div>
      <PageHeader eyebrow="Sales" title="Sales Orders" sub={`${data.salesOrders.length} total - no ledger impact until converted`} action={<button style={styles.btnPrimary} onClick={() => setShowForm(s => !s)}>{showForm ? "Cancel" : "+ New sales order"}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>New sales order</div>
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Customer name" value={form.customer} onChange={e => setForm({ ...form, customer: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <input type="date" style={styles.input} value={form.expectedDate} onChange={e => setForm({ ...form, expectedDate: e.target.value })} placeholder="Expected fulfillment" />
            <select style={styles.input} value={form.projectId} onChange={e => setForm({ ...form, projectId: e.target.value })}><option value="">No project</option>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
          </div>
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Description</th><th style={styles.th}>Inventory item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Price</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: "100%" }} value={it.desc} onChange={e => updateItem(i, "desc", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.salePrice); } }}>
                  <option value="">None</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name} ({inv.qty} in stock)</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(Number(it.qty || 0) * Number(it.price || 0))}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <RowLine label="Order total" value={total} bold divider />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={createOrder}>Create sales order</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Order</th><th style={styles.th}>Customer</th><th style={styles.th}>Expected</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(so => {
            const t = so.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0);
            return (
              <tr key={so.id}><td style={styles.tdMono}>{so.id}</td><td style={styles.td}>{so.customer}</td><td style={styles.tdMono}>{so.expectedDate ? fmtDate(so.expectedDate) : "-"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td>
                <td style={styles.td}><span style={{ ...styles.pill, ...(so.status === "fulfilled" ? styles.pillGreen : so.status === "cancelled" ? styles.pillRose : styles.pillAmber) }}>{so.status}</span></td>
                <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                  {so.status === "open" && <><button style={styles.iconBtn} onClick={() => convertToInvoice(so)}>Convert to invoice</button>{" "}<button style={styles.iconBtn} onClick={() => cancelOrder(so)}>Cancel</button>{" "}</>}
                  <button style={styles.iconBtn} onClick={() => deleteOrder(so)}>🗑</button>
                </td></tr>
            );
          })}</tbody>
        </table>
        {sorted.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No sales orders yet.</div>}
      </div>
    </div>
  );
}

/* =================================== Sales Receipts =================================== */
// A sale paid for on the spot - posts to the ledger immediately (cash in,
// revenue recognized), with no unpaid/AR stage to track.
function SalesReceipts({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const blank = () => ({ customer: "", date: todayStr(), bankId: data.banks[0]?.id, projectId: "", items: [{ desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }], taxes: [] });
  const [form, setForm] = useState(blank());
  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }] }));
  const totals = computeDocTotals(form, data.taxGroups);

  const createReceipt = () => {
    if (!form.customer || form.items.some(it => !it.desc)) return notify("Add a customer and describe each line item");
    const bank = data.banks.find(b => b.id === form.bankId);
    if (!bank) return notify("Choose which bank received the payment");
    const id = `SR-${data.nextSalesReceiptNum}`;
    const lines = salesReceiptJournalLines(form, data, form.bankId);
    const txn = buildTxn(`Sales receipt ${id} - ${form.customer}`, form.date, lines, "sales-receipt", id);
    if (!txn) return notify("Entry not balanced");
    let newData = { ...data, transactions: [...data.transactions, txn], salesReceipts: [...data.salesReceipts, { ...form, items: snapshotTaxItems(form.items, data.taxGroups), id }], nextSalesReceiptNum: data.nextSalesReceiptNum + 1 };
    let lots = newData.inventoryLots;
    form.items.forEach(it => {
      if (!it.inventoryId) return;
      newData = adjustInventory(newData, it.inventoryId, -Number(it.qty || 0));
      const r = consumeFIFO(lots, it.inventoryId, Number(it.qty || 0));
      lots = r.lots;
    });
    setData({ ...newData, inventoryLots: lots });
    setForm(blank()); setShowForm(false);
    notify(`${id} recorded and posted to the ledger`);
  };

  const deleteReceipt = async (sr) => {
    if (!(await confirm(`Delete ${sr.id}? Its journal entry is removed and any inventory it shipped is restored. This can't be undone.`))) return;
    setData(d => {
      let next = { ...d, transactions: stripDocTransactions(d.transactions, sr.id) };
      let lots = next.inventoryLots;
      (sr.items || []).forEach(it => {
        if (!it.inventoryId) return;
        next = adjustInventory(next, it.inventoryId, Number(it.qty || 0));
        lots = restoreFIFO(lots, it.inventoryId, Number(it.qty || 0));
      });
      return { ...next, inventoryLots: lots, salesReceipts: next.salesReceipts.filter(x => x.id !== sr.id) };
    });
    setViewingId(null);
    notify(`${sr.id} deleted - ledger and stock restored`);
  };

  const sorted = [...data.salesReceipts].sort((a, b) => b.date.localeCompare(a.date));
  const viewing = viewingId ? data.salesReceipts.find(i => i.id === viewingId) : null;
  if (viewing) return <DocumentDetail doc={viewing} docType="invoice" data={data} onBack={() => setViewingId(null)} onEdit={() => {}} onToggleLock={() => {}} onDelete={() => deleteReceipt(viewing)} />;

  return (
    <div>
      <PageHeader eyebrow="Sales" title="Sales Receipts" sub={`${data.salesReceipts.length} total - paid in full at the point of sale`} action={<button style={styles.btnPrimary} onClick={() => setShowForm(s => !s)}>{showForm ? "Cancel" : "+ New sales receipt"}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>New sales receipt</div>
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Customer name" value={form.customer} onChange={e => setForm({ ...form, customer: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <select style={styles.input} value={form.bankId} onChange={e => setForm({ ...form, bankId: e.target.value })}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
            <select style={styles.input} value={form.projectId} onChange={e => setForm({ ...form, projectId: e.target.value })}><option value="">No project</option>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
          </div>
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Description</th><th style={styles.th}>Inventory item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Price</th><th style={styles.th}>Tax group</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: "100%" }} value={it.desc} onChange={e => updateItem(i, "desc", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.salePrice); } }}>
                  <option value="">None</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name} ({inv.qty} in stock)</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.taxGroupId} onChange={e => updateItem(i, "taxGroupId", e.target.value)}>
                  <option value="">No tax group</option>{data.taxGroups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}</select></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(totals.lineCalcs[i]?.lineTotal || 0)}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <LineTaxSummary lineCalcs={totals.lineCalcs} subtotal={totals.subtotal} afterLineTax={totals.afterLineTax} />
          <RowLine label="Total received" value={totals.finalAmount} bold divider />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={createReceipt}>Record & post to ledger</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Receipt</th><th style={styles.th}>Customer</th><th style={styles.th}>Bank</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(sr => {
            const t = computeDocTotals(sr, data.taxGroups).finalAmount;
            return (
              <tr key={sr.id} onClick={() => setViewingId(sr.id)} style={{ cursor: "pointer" }}>
                <td style={styles.tdMono}>{sr.id}</td><td style={styles.td}>{sr.customer}</td><td style={styles.td}>{data.banks.find(b => b.id === sr.bankId)?.name || "-"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td>
                <td style={styles.td} onClick={e => e.stopPropagation()}><button style={styles.iconBtn} onClick={() => deleteReceipt(sr)}>🗑</button></td></tr>
            );
          })}</tbody>
        </table>
        {sorted.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No sales receipts yet.</div>}
      </div>
    </div>
  );
}

/* =================================== Credit Notes & Sales Returns =================================== */
// Shared implementation - a credit note is a pure financial adjustment; a
// sales return is the same thing with inventory line items, which restocks
// goods and reverses their cost of sale. allowInventory switches between them.
function CreditNoteLike({ data, setData, notify, allowInventory, docPrefix, counterKey, collectionKey, title, eyebrowSub }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const openInvoices = data.invoices.filter(i => (i.amountPaid || 0) < computeDocTotals(i, data.taxGroups).finalAmount - 0.5);
  const blank = () => ({ customer: "", date: todayStr(), invoiceId: "", refundBankId: "", reason: "", items: [{ desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }], taxes: [] });
  const [form, setForm] = useState(blank());
  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }] }));
  const totals = computeDocTotals(form, data.taxGroups);

  const create = () => {
    if (!form.customer || form.items.some(it => !it.desc)) return notify("Add a customer and describe each line item");
    if (allowInventory && form.items.some(it => !it.inventoryId)) return notify("Every line must reference an inventory item, since this restocks goods");
    const id = `${docPrefix}-${data[counterKey]}`;
    const lines = creditNoteJournalLines(form, data, form.refundBankId || null);
    const txn = buildTxn(`${title} ${id} - ${form.customer}${form.invoiceId ? ` (applied to ${form.invoiceId})` : ""}`, form.date, lines, allowInventory ? "sales-return" : "credit-note", id);
    if (!txn) return notify("Entry not balanced");
    let newData = {
      ...data,
      transactions: [...data.transactions, txn],
      [collectionKey]: [...data[collectionKey], { ...form, items: snapshotTaxItems(form.items, data.taxGroups), id, status: "issued" }],
      [counterKey]: data[counterKey] + 1,
    };
    if (form.invoiceId) newData.invoices = newData.invoices.map(i => i.id === form.invoiceId ? { ...i, amountPaid: (i.amountPaid || 0) + totals.finalAmount } : i);
    let lots = newData.inventoryLots;
    form.items.forEach(it => {
      if (!it.inventoryId) return;
      newData = adjustInventory(newData, it.inventoryId, Number(it.qty || 0), Number(it.price || 0));
      lots = [...lots, { id: uid("lot"), itemId: it.inventoryId, date: form.date, qty: Number(it.qty || 0), remainingQty: Number(it.qty || 0), unitCost: Number(it.price || 0), sourceDocId: id }];
    });
    setData({ ...newData, inventoryLots: lots });
    setForm(blank()); setShowForm(false);
    notify(`${id} issued${form.refundBankId ? " and refunded" : form.invoiceId ? ` and applied to ${form.invoiceId}` : ""}`);
  };

  const deleteDoc = async (doc) => {
    if (!(await confirm(`Delete ${doc.id}? Its journal entry is removed${allowInventory ? " and any restocked inventory is reversed" : ""}. This can't be undone.`))) return;
    setData(d => {
      let next = { ...d, transactions: stripDocTransactions(d.transactions, doc.id) };
      if (allowInventory) {
        next = { ...next, inventoryLots: next.inventoryLots.filter(l => l.sourceDocId !== doc.id) };
        (doc.items || []).forEach(it => { if (it.inventoryId) next = adjustInventory(next, it.inventoryId, -Number(it.qty || 0)); });
      }
      if (doc.invoiceId) {
        const t = computeDocTotals(doc, d.taxGroups).finalAmount;
        next = { ...next, invoices: next.invoices.map(i => i.id === doc.invoiceId ? { ...i, amountPaid: Math.max(0, (i.amountPaid || 0) - t) } : i) };
      }
      return { ...next, [collectionKey]: next[collectionKey].filter(x => x.id !== doc.id) };
    });
    setViewingId(null);
    notify(`${doc.id} deleted`);
  };

  const sorted = [...data[collectionKey]].sort((a, b) => b.date.localeCompare(a.date));
  const viewing = viewingId ? data[collectionKey].find(i => i.id === viewingId) : null;
  if (viewing) return <DocumentDetail doc={viewing} docType="invoice" data={data} onBack={() => setViewingId(null)} onEdit={() => {}} onToggleLock={() => {}} onDelete={() => deleteDoc(viewing)} />;

  return (
    <div>
      <PageHeader eyebrow="Sales" title={title} sub={`${data[collectionKey].length} total - ${eyebrowSub}`} action={<button style={styles.btnPrimary} onClick={() => setShowForm(s => !s)}>{showForm ? "Cancel" : `+ New ${title.toLowerCase().slice(0, -1)}`}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>New {title.toLowerCase().slice(0, -1)}</div>
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Customer name" value={form.customer} onChange={e => setForm({ ...form, customer: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <select style={styles.input} value={form.invoiceId} onChange={e => setForm({ ...form, invoiceId: e.target.value })}>
              <option value="">Not applied to an invoice</option>{openInvoices.map(inv => <option key={inv.id} value={inv.id}>Apply to {inv.id}</option>)}
            </select>
            <select style={styles.input} value={form.refundBankId} onChange={e => setForm({ ...form, refundBankId: e.target.value })}>
              <option value="">Credit to Accounts Receivable</option>{data.banks.map(b => <option key={b.id} value={b.id}>Refund from {b.name}</option>)}
            </select>
          </div>
          <input style={{ ...styles.input, width: "100%", marginTop: 10 }} placeholder="Reason (optional)" value={form.reason} onChange={e => setForm({ ...form, reason: e.target.value })} />
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Description</th><th style={styles.th}>Inventory item{allowInventory ? " (required)" : ""}</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Price</th><th style={styles.th}>Tax group</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: "100%" }} value={it.desc} onChange={e => updateItem(i, "desc", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.unitCost); } }}>
                  <option value="">{allowInventory ? "Choose item" : "None"}</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name}</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.taxGroupId} onChange={e => updateItem(i, "taxGroupId", e.target.value)}>
                  <option value="">No tax group</option>{data.taxGroups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}</select></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(totals.lineCalcs[i]?.lineTotal || 0)}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <LineTaxSummary lineCalcs={totals.lineCalcs} subtotal={totals.subtotal} afterLineTax={totals.afterLineTax} />
          <RowLine label="Total credit" value={totals.finalAmount} bold divider />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={create}>Issue {title.toLowerCase().slice(0, -1)}</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>{title.slice(0, -1)}</th><th style={styles.th}>Customer</th><th style={styles.th}>Applied to</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(doc => {
            const t = computeDocTotals(doc, data.taxGroups).finalAmount;
            return (
              <tr key={doc.id} onClick={() => setViewingId(doc.id)} style={{ cursor: "pointer" }}>
                <td style={styles.tdMono}>{doc.id}</td><td style={styles.td}>{doc.customer}</td><td style={styles.td}>{doc.invoiceId || (doc.refundBankId ? `Refunded via ${data.banks.find(b => b.id === doc.refundBankId)?.name}` : "Open credit")}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td>
                <td style={styles.td} onClick={e => e.stopPropagation()}><button style={styles.iconBtn} onClick={() => deleteDoc(doc)}>🗑</button></td></tr>
            );
          })}</tbody>
        </table>
        {sorted.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>None issued yet.</div>}
      </div>
    </div>
  );
}
function CreditNotes({ data, setData, notify }) {
  return <CreditNoteLike data={data} setData={setData} notify={notify} allowInventory={false} docPrefix="CN" counterKey="nextCreditNoteNum" collectionKey="creditNotes" title="Credit Notes" eyebrowSub="billing adjustments and goodwill credits" />;
}
function SalesReturns({ data, setData, notify }) {
  return <CreditNoteLike data={data} setData={setData} notify={notify} allowInventory={true} docPrefix="RTN" counterKey="nextSalesReturnNum" collectionKey="salesReturns" title="Sales Returns" eyebrowSub="physical goods returned, restocked at cost" />;
}

function billJournalLines(form, t, data) {
  const invLines = form.items.filter(it => it.inventoryId);
  const nonInvSubtotal = form.items.filter(it => !it.inventoryId).reduce((s, it) => s + it.qty * it.price, 0);
  const invByAccount = {};
  invLines.forEach(it => { const acct = inventoryAccountForItem(data, it.inventoryId); invByAccount[acct] = (invByAccount[acct] || 0) + it.qty * it.price; });
  return [
    ...Object.entries(invByAccount).map(([acct, amt]) => ({ accountId: acct, debit: amt, credit: 0 })),
    ...(nonInvSubtotal > 0 ? [{ accountId: form.expenseAccountId, debit: nonInvSubtotal, credit: 0 }] : []),
    { accountId: "2000", debit: 0, credit: t.finalAmount },
    ...t.lineCalcs.flatMap(l => taxJournalLines(l.components, "bill")),
    ...taxJournalLines(t.cascadeSteps, "bill"),
  ];
}
// The exact mirror image of billJournalLines: inventory lines are removed
// from stock instead of added (nothing was ever sold, so there's no COGS to
// reverse - the goods just go back), the non-inventory portion reduces
// Purchase Returns & Allowances instead of an expense, and Accounts Payable
// is debited (reducing what's owed) instead of credited - or, if the vendor
// refunded cash instead, a bank account is debited in its place. Reusing
// taxJournalLines with "invoice" instead of "bill" on the same steps always
// produces the reversed direction, the same trick creditNoteJournalLines
// uses to mirror invoiceJournalLines.
function vendorCreditJournalLines(form, data, refundBankId) {
  const t = computeDocTotals(form, data.taxGroups);
  const invLines = form.items.filter(it => it.inventoryId);
  const nonInvSubtotal = form.items.filter(it => !it.inventoryId).reduce((s, it) => s + it.qty * it.price, 0);
  const invByAccount = {};
  invLines.forEach(it => { const acct = inventoryAccountForItem(data, it.inventoryId); invByAccount[acct] = (invByAccount[acct] || 0) + it.qty * it.price; });
  const bank = refundBankId ? data.banks.find(b => b.id === refundBankId) : null;
  const debitAccountId = bank ? bank.accountId : "2000";
  return [
    ...Object.entries(invByAccount).map(([acct, amt]) => ({ accountId: acct, debit: 0, credit: amt })),
    ...(nonInvSubtotal > 0 ? [{ accountId: "5900", debit: 0, credit: nonInvSubtotal }] : []),
    { accountId: debitAccountId, debit: t.finalAmount, credit: 0 },
    ...t.lineCalcs.flatMap(l => taxJournalLines(l.components, "invoice")),
    ...taxJournalLines(t.cascadeSteps, "invoice"),
  ];
}
function Bills({ data, setData, postTransaction, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const expenseAccounts = activeAccounts(data).filter(a => a.type === "expense");
  const blank = () => ({ vendor: "", date: todayStr(), dueDate: "", projectId: "", expenseAccountId: expenseAccounts[0]?.id, items: [{ desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }], taxes: [] });
  const [form, setForm] = useState(blank());

  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }] }));
  const totals = computeDocTotals(form, data.taxGroups);

  const startEdit = (bill) => {
    if (bill.locked) return notify("This bill is locked - unlock it first to edit");
    setForm({ vendor: bill.vendor, date: bill.date, dueDate: bill.dueDate || "", projectId: bill.projectId || "", expenseAccountId: bill.expenseAccountId || expenseAccounts[0]?.id, items: bill.items.map(it => ({ ...it, taxComponents: undefined })), taxes: (bill.taxes || []).map(t => ({ ...t })) });
    setEditingId(bill.id); setViewingId(null); setShowForm(true);
  };

  const saveBill = () => {
    if (!form.vendor || form.items.some(it => !it.desc)) return notify("Add a vendor and describe each line item");
    const isEdit = Boolean(editingId);
    const id = isEdit ? editingId : `BILL-${data.nextBillNum}`;
    let working = { ...data };
    // If editing: remove old journal, remove the stock lots this bill created,
    // and take the old received qty back out of inventory at the lot's cost.
    if (isEdit) {
      const old = data.bills.find(b => b.id === id);
      working = { ...working, transactions: stripDocTransactions(working.transactions, id) };
      working = { ...working, inventoryLots: working.inventoryLots.filter(l => l.sourceDocId !== id) };
      (old?.items || []).forEach(it => { if (it.inventoryId) working = adjustInventory(working, it.inventoryId, -Number(it.qty || 0)); });
    }
    const t = computeDocTotals(form, working.taxGroups);
    const txn = buildTxn(`Bill ${id} - ${form.vendor}`, form.date, billJournalLines(form, t, working), "bill", id);
    if (!txn) return notify("Entry not balanced - check tax accounts");
    working = { ...working, transactions: [...working.transactions, txn] };
    const snapshotItems = snapshotTaxItems(form.items, data.taxGroups);
    const docBase = { ...form, items: snapshotItems, id };
    working = isEdit
      ? { ...working, bills: working.bills.map(b => b.id === id ? { ...b, ...docBase } : b) }
      : { ...working, bills: [...working.bills, { ...docBase, status: "unpaid", amountPaid: 0, locked: false }], nextBillNum: working.nextBillNum + 1 };
    const invLines = form.items.filter(it => it.inventoryId);
    invLines.forEach(it => { working = adjustInventory(working, it.inventoryId, Number(it.qty || 0), Number(it.price || 0)); });
    working = { ...working, inventoryLots: [...working.inventoryLots, ...invLines.map(it => ({ id: uid("lot"), itemId: it.inventoryId, date: form.date, qty: Number(it.qty || 0), remainingQty: Number(it.qty || 0), unitCost: Number(it.price || 0), sourceDocId: id }))] };
    setData(working);
    setForm(blank()); setShowForm(false); setEditingId(null);
    notify(isEdit ? `${id} updated and reposted to the ledger` : `${id} created and posted to the ledger`);
  };

  const toggleLock = (id) => setData(d => ({ ...d, bills: d.bills.map(b => b.id === id ? { ...b, locked: !b.locked } : b) }));

  // Delete a bill: journal reversed, its stock lots removed, and the received
  // quantity taken back out of inventory. Payments must be deleted first.
  const deleteBill = async (bill) => {
    if (bill.locked) return notify("This bill is locked - unlock it first");
    if (isLocked(bill.date, data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - this bill falls within that period and can't be deleted`);
    if ((bill.amountPaid || 0) > 0) return notify("This bill has payments recorded against it - delete those payments (Payments tab) before deleting the bill");
    if (!(await confirm(`Delete ${bill.id}? Its journal entry and the stock it received are removed. This can't be undone.`))) return;
    setData(d => {
      let next = { ...d, transactions: stripDocTransactions(d.transactions, bill.id) };
      next = { ...next, inventoryLots: next.inventoryLots.filter(l => l.sourceDocId !== bill.id) };
      (bill.items || []).forEach(it => { if (it.inventoryId) next = adjustInventory(next, it.inventoryId, -Number(it.qty || 0)); });
      return { ...next, bills: next.bills.filter(b => b.id !== bill.id) };
    });
    setViewingId(null);
    notify(`${bill.id} deleted - ledger and stock reversed`);
  };

  const sorted = [...data.bills].sort((a, b) => b.date.localeCompare(a.date));
  const viewing = viewingId ? data.bills.find(b => b.id === viewingId) : null;
  if (viewing) return <DocumentDetail doc={viewing} docType="bill" data={data} onBack={() => setViewingId(null)} onEdit={() => startEdit(viewing)} onToggleLock={() => toggleLock(viewing.id)} onDelete={() => deleteBill(viewing)} />;

  return (
    <div>
      <PageHeader eyebrow="Purchases" title="Bills" sub={`${data.bills.length} total`} action={<button style={styles.btnPrimary} onClick={() => { setShowForm(s => !s); setEditingId(null); setForm(blank()); }}>{showForm ? "Cancel" : "+ New bill"}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>{editingId ? `Edit ${editingId}` : "New bill"}</div>
          {editingId && <div style={{ fontSize: 12, color: theme.amber, marginTop: 4 }}>Saving will remove the original journal entry and stock lots, then repost with the new values. Payments already made stay attached.</div>}
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Vendor name" value={form.vendor} onChange={e => setForm({ ...form, vendor: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <input type="date" style={styles.input} value={form.dueDate} onChange={e => setForm({ ...form, dueDate: e.target.value })} placeholder="Due date" />
            <select style={styles.input} value={form.expenseAccountId} onChange={e => setForm({ ...form, expenseAccountId: e.target.value })}>{expenseAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
            <select style={styles.input} value={form.projectId} onChange={e => setForm({ ...form, projectId: e.target.value })}><option value="">No project</option>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
          </div>
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Description</th><th style={styles.th}>Inventory item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={styles.th}>Tax group</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: "100%" }} value={it.desc} onChange={e => updateItem(i, "desc", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.unitCost); } }}>
                  <option value="">None (expense line)</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name}</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.taxGroupId} onChange={e => updateItem(i, "taxGroupId", e.target.value)}>
                  <option value="">No tax group</option>{data.taxGroups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}</select></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(totals.lineCalcs[i]?.lineTotal || 0)}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <LineTaxSummary lineCalcs={totals.lineCalcs} subtotal={totals.subtotal} afterLineTax={totals.afterLineTax} />
          <TaxStack taxes={form.taxes} setTaxes={(fn) => setForm(f => ({ ...f, taxes: typeof fn === "function" ? fn(f.taxes) : fn }))} subtotal={totals.afterLineTax} docType="bill" accounts={data.accounts} label="Document-level taxes (applied to the bill total)" />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={saveBill}>{editingId ? "Save changes & repost" : "Create & post to ledger"}</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Bill</th><th style={styles.th}>Vendor</th><th style={styles.th}>Due</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={{ ...styles.th, textAlign: "right" }}>Paid</th><th style={styles.th}>Status</th></tr></thead>
          <tbody>{sorted.map(b => {
            const total = computeDocTotals(b, data.taxGroups).finalAmount;
            const overdue = b.status !== "paid" && b.dueDate && b.dueDate < todayStr();
            const status = b.amountPaid >= total - 0.5 ? "paid" : b.amountPaid > 0 ? "partial" : overdue ? "overdue" : "unpaid";
            return (
              <tr key={b.id} onClick={() => setViewingId(b.id)} style={{ cursor: "pointer" }}>
                <td style={styles.tdMono}>{b.locked ? "🔒 " : ""}{b.id}</td><td style={styles.td}>{b.vendor}</td><td style={styles.tdMono}>{b.dueDate ? fmtDate(b.dueDate) : "-"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(total)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(b.amountPaid || 0)}</td>
                <td style={styles.td}><span style={{ ...styles.pill, ...(status === "paid" ? styles.pillGreen : status === "overdue" ? styles.pillRose : styles.pillAmber) }}>{status}</span></td></tr>
            );
          })}</tbody>
        </table>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>Click a bill to view, edit, lock or print it. Settle it from the Payments tab.</div>
      </div>
    </div>
  );
}

/* =================================== Payments =================================== */
function Payments({ data, setData, postTransaction, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const openInvoices = data.invoices.filter(i => (i.amountPaid || 0) < computeDocTotals(i, data.taxGroups).finalAmount - 0.5);
  const openBills = data.bills.filter(b => (b.amountPaid || 0) < computeDocTotals(b, data.taxGroups).finalAmount - 0.5);
  const [form, setForm] = useState({ date: todayStr(), type: "received", bankId: data.banks[0]?.id, relatedType: "invoice", relatedId: openInvoices[0]?.id || "", refundTo: "", amount: "", memo: "", otherAccountId: "4100" });

  const submit = () => {
    const bank = data.banks.find(b => b.id === form.bankId);
    if (!bank || !Number(form.amount)) return notify("Choose a bank and enter an amount");
    const amt = Number(form.amount);
    let lines = [];
    if (form.type === "received") lines = [{ accountId: bank.accountId, debit: amt, credit: 0 }, { accountId: form.relatedType === "invoice" ? "1100" : "4100", debit: 0, credit: amt }];
    else if (form.type === "paid") lines = [{ accountId: form.relatedType === "bill" ? "2000" : form.otherAccountId, debit: amt, credit: 0 }, { accountId: bank.accountId, debit: 0, credit: amt }];
    else if (form.type === "refund_to_customer") lines = [{ accountId: "4200", debit: amt, credit: 0 }, { accountId: bank.accountId, debit: 0, credit: amt }];
    else if (form.type === "refund_from_vendor") lines = [{ accountId: bank.accountId, debit: amt, credit: 0 }, { accountId: "5900", debit: 0, credit: amt }];
    const memoDefault = { received: "Payment received", paid: "Payment made", refund_to_customer: "Refund to customer", refund_from_vendor: "Refund from vendor" }[form.type];
    const memo = form.memo || `${memoDefault} - ${form.relatedId || form.refundTo || "general"}`;
    const txn = buildTxn(memo, form.date, lines, form.type.startsWith("refund") ? "refund" : "payment");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({
      ...d,
      transactions: [...d.transactions, txn],
      payments: [...d.payments, { ...form, id: uid("pay"), amount: amt, txnId: txn.id }],
      invoices: form.relatedType === "invoice" && form.type === "received" ? d.invoices.map(i => i.id === form.relatedId ? { ...i, amountPaid: (i.amountPaid || 0) + amt } : i) : d.invoices,
      bills: form.relatedType === "bill" && form.type === "paid" ? d.bills.map(b => b.id === form.relatedId ? { ...b, amountPaid: (b.amountPaid || 0) + amt } : b) : d.bills,
    }));
    setForm({ ...form, amount: "", memo: "", refundTo: "" });
    notify("Recorded");
  };

  const sorted = [...data.payments].sort((a, b) => b.date.localeCompare(a.date));
  const typeLabel = { received: "In", paid: "Out", refund_to_customer: "Refund out", refund_from_vendor: "Refund in" };
  return (
    <div>
      <PageHeader eyebrow="Cash movement" title="Payments" sub={`${data.payments.length} recorded`} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Record a payment or refund</div>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
          <select style={styles.input} value={form.type} onChange={e => setForm({ ...form, type: e.target.value, relatedType: e.target.value === "received" ? "invoice" : e.target.value === "paid" ? "bill" : "other" })}>
            <option value="received">Money in</option><option value="paid">Money out</option>
            <option value="refund_to_customer">Refund to customer</option><option value="refund_from_vendor">Refund from vendor</option>
          </select>
          <select style={styles.input} value={form.bankId} onChange={e => setForm({ ...form, bankId: e.target.value })}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
          {(form.type === "received" || form.type === "paid") && <select style={styles.input} value={form.relatedType} onChange={e => setForm({ ...form, relatedType: e.target.value, relatedId: "" })}>
            {form.type === "received" ? <><option value="invoice">Against invoice</option><option value="other">Other income</option></> : <><option value="bill">Against bill</option><option value="other">Other expense</option></>}
          </select>}
          {form.relatedType === "invoice" && form.type === "received" && <select style={styles.input} value={form.relatedId} onChange={e => setForm({ ...form, relatedId: e.target.value })}>{openInvoices.map(i => <option key={i.id} value={i.id}>{i.id} · {i.customer}</option>)}</select>}
          {form.relatedType === "bill" && form.type === "paid" && <select style={styles.input} value={form.relatedId} onChange={e => setForm({ ...form, relatedId: e.target.value })}>{openBills.map(b => <option key={b.id} value={b.id}>{b.id} · {b.vendor}</option>)}</select>}
          {(form.type === "refund_to_customer" || form.type === "refund_from_vendor") && <input style={styles.input} placeholder={form.type === "refund_to_customer" ? "Customer name" : "Vendor name"} value={form.refundTo} onChange={e => setForm({ ...form, refundTo: e.target.value })} />}
          <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 130 }} placeholder="Amount" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Memo (optional)" value={form.memo} onChange={e => setForm({ ...form, memo: e.target.value })} />
          <button style={styles.btnPrimary} onClick={submit}>Record</button>
        </div>
      </div>
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Type</th><th style={styles.th}>Bank</th><th style={styles.th}>Related to</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th></th></tr></thead>
          <tbody>{sorted.map(p => (
            <tr key={p.id}><td style={styles.tdMono}>{fmtDate(p.date)}</td><td style={styles.td}>{typeLabel[p.type] || p.type}</td>
              <td style={styles.td}>{data.banks.find(b => b.id === p.bankId)?.name || "-"}</td><td style={styles.td}>{p.relatedId || p.refundTo || "general"}</td>
              <td style={{ ...styles.tdMono, textAlign: "right", color: (p.type === "received" || p.type === "refund_from_vendor") ? theme.emerald : theme.rose }}>{fmt(p.amount)}</td>
              <td style={styles.td}><button style={styles.iconBtn} onClick={async () => {
                if (!(await confirm("Delete this payment? Its ledger entry is removed and any invoice/bill it settled goes back to unpaid."))) return;
                // Locate the ledger entry: by stored link, or for older
                // records by source + date + bank-line amount.
                const bank = data.banks.find(b => b.id === p.bankId);
                const txn = (p.txnId && data.transactions.find(t => t.id === p.txnId))
                  || data.transactions.find(t => (t.source === "payment" || t.source === "refund" || t.source === "bank-import") && t.date === p.date &&
                      t.lines.some(l => bank && l.accountId === bank.accountId && Math.abs((l.debit || l.credit) - p.amount) < 0.005) &&
                      (!p.relatedId || (t.memo || "").includes(p.relatedId)));
                setData(d => {
                  let next = { ...d, payments: d.payments.filter(x => x.id !== p.id) };
                  if (txn) {
                    next = { ...next, transactions: next.transactions.filter(t => t.id !== txn.id) };
                    // A matched bank line returns to the uncategorized feed -
                    // the money still exists on the statement.
                    if (txn.feedId) next = { ...next, bankFeed: next.bankFeed.map(f => f.id === txn.feedId ? { ...f, status: "uncategorized", txnId: null, accountId: null } : f) };
                  }
                  if (p.relatedType === "invoice") next = { ...next, invoices: next.invoices.map(i => i.id === p.relatedId ? { ...i, amountPaid: Math.max(0, (i.amountPaid || 0) - p.amount) } : i) };
                  if (p.relatedType === "bill") next = { ...next, bills: next.bills.map(b => b.id === p.relatedId ? { ...b, amountPaid: Math.max(0, (b.amountPaid || 0) - p.amount) } : b) };
                  return next;
                });
                notify(txn ? "Payment deleted and its journal entry removed" : "Payment record deleted - its journal entry couldn't be located; check the Journal");
              }}>🗑</button></td></tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  );
}

/* =================================== Quick Expenses =================================== */
// A friendlier, purpose-built view over the same recurring-journal engine
// used by the Journal tab's "Repeat this entry" - scoped to the simple
// Dr expense / Cr bank shape a recurring expense (rent, subscriptions,
// retainers) actually needs, so setting one up doesn't require the general
// journal form.
function RecurringExpenses({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const expenseAccounts = activeAccounts(data).filter(a => a.type === "expense");
  const blank = () => ({ vendor: "", category: expenseAccounts[0]?.id, bankId: data.banks[0]?.id, amount: "", frequency: "monthly", startDate: todayStr(), endDate: "" });
  const [form, setForm] = useState(blank());
  const mine = data.recurringJournals.filter(r => r.meta === "expense");

  const create = () => {
    if (!form.vendor || !Number(form.amount)) return notify("Name the vendor and enter an amount");
    const bank = data.banks.find(b => b.id === form.bankId);
    if (!bank) return notify("Choose which bank pays it");
    const amount = Number(form.amount);
    const lines = [{ accountId: form.category, debit: amount, credit: 0 }, { accountId: bank.accountId, debit: 0, credit: amount }];
    const txn = buildTxn(form.vendor, form.startDate, lines, "expense");
    if (!txn) return notify("Entry not balanced");
    const nextDate = advanceDate(form.startDate, form.frequency);
    setData(d => ({
      ...d,
      transactions: [...d.transactions, txn],
      expenses: [...d.expenses, { vendor: form.vendor, date: form.startDate, category: form.category, bankId: form.bankId, amount, deductible: true, memo: "", id: uid("exp"), txnId: txn.id }],
      recurringJournals: [...d.recurringJournals, { id: uid("rec"), memo: form.vendor, lines, frequency: form.frequency, startDate: form.startDate, nextDate, endDate: form.endDate || null, active: true, meta: "expense" }],
    }));
    setForm(blank());
    notify(`Recorded and scheduled ${form.frequency} - next on ${fmtDate(nextDate)}`);
  };

  const toggle = (r) => setData(d => ({ ...d, recurringJournals: d.recurringJournals.map(x => x.id === r.id ? { ...x, active: !x.active } : x) }));
  const remove = async (r) => {
    if (!(await confirm(`Delete the recurring expense "${r.memo}"? Entries it has already posted stay on the ledger.`))) return;
    setData(d => ({ ...d, recurringJournals: d.recurringJournals.filter(x => x.id !== r.id) }));
    notify("Recurring expense deleted");
  };

  return (
    <div>
      <PageHeader eyebrow="Purchases" title="Recurring Expenses" sub={`${mine.length} scheduled - rent, subscriptions, retainers`} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Schedule a recurring expense</div>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Vendor / description" value={form.vendor} onChange={e => setForm({ ...form, vendor: e.target.value })} />
          <select style={styles.input} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{expenseAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
          <select style={styles.input} value={form.bankId} onChange={e => setForm({ ...form, bankId: e.target.value })}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
          <input type="number" style={{ ...styles.input, width: 130 }} placeholder="Amount" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
          <select style={styles.input} value={form.frequency} onChange={e => setForm({ ...form, frequency: e.target.value })}>
            <option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="quarterly">Quarterly</option><option value="yearly">Yearly</option>
          </select>
          <input type="date" style={styles.input} value={form.startDate} onChange={e => setForm({ ...form, startDate: e.target.value })} />
          <input type="date" style={styles.input} value={form.endDate} onChange={e => setForm({ ...form, endDate: e.target.value })} placeholder="Ends (optional)" />
          <button style={styles.btnPrimary} onClick={create}>Record & schedule</button>
        </div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 8 }}>Posts the first occurrence immediately, then repeats automatically when the app opens on or after each due date.</div>
      </div>
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Vendor</th><th style={styles.th}>Frequency</th><th style={styles.th}>Next run</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
          <tbody>{mine.map(r => (
            <tr key={r.id} style={{ opacity: r.active ? 1 : 0.55 }}>
              <td style={styles.td}>{r.memo}</td><td style={styles.td}>{r.frequency}</td><td style={styles.tdMono}>{fmtDate(r.nextDate)}</td>
              <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.lines.reduce((s, l) => s + l.debit, 0))}</td>
              <td style={styles.td}><span style={{ ...styles.pill, ...(r.active ? styles.pillGreen : styles.pillAmber) }}>{r.active ? "active" : "paused"}</span></td>
              <td style={{ ...styles.td, whiteSpace: "nowrap" }}><button style={styles.iconBtn} onClick={() => toggle(r)}>{r.active ? "\u23f8 pause" : "\u25b6 resume"}</button>{" "}<button style={styles.iconBtn} onClick={() => remove(r)}>🗑</button></td>
            </tr>
          ))}</tbody>
        </table>
        {mine.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No recurring expenses scheduled yet.</div>}
      </div>
    </div>
  );
}

/* =================================== Purchase Orders =================================== */
// A commitment to buy from a vendor before the goods or bill arrive - no
// ledger impact until it's converted to a real bill.
function PurchaseOrders({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const expenseAccounts = activeAccounts(data).filter(a => a.type === "expense");
  const blank = () => ({ vendor: "", date: todayStr(), expectedDate: "", expenseAccountId: expenseAccounts[0]?.id, projectId: "", items: [{ desc: "", qty: 1, price: 0, inventoryId: "" }] });
  const [form, setForm] = useState(blank());
  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "" }] }));
  const total = form.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0);

  const createOrder = () => {
    if (!form.vendor || form.items.some(it => !it.desc)) return notify("Add a vendor and describe each line item");
    const id = `PO-${data.nextPurchaseOrderNum}`;
    setData(d => ({ ...d, purchaseOrders: [...d.purchaseOrders, { ...form, id, status: "open" }], nextPurchaseOrderNum: d.nextPurchaseOrderNum + 1 }));
    setForm(blank()); setShowForm(false);
    notify(`${id} created - no ledger impact until converted to a bill`);
  };

  const convertToBill = async (po) => {
    if (!(await confirm(`Convert ${po.id} to a bill from ${po.vendor}? This posts a real bill with the same line items to the ledger.`))) return;
    const billForm = { vendor: po.vendor, date: todayStr(), dueDate: "", expenseAccountId: po.expenseAccountId, projectId: po.projectId, items: po.items, taxes: [] };
    const t = computeDocTotals(billForm, data.taxGroups);
    const lines = billJournalLines(billForm, t, data);
    const id = `BILL-${data.nextBillNum}`;
    const txn = buildTxn(`Bill ${id} - ${po.vendor}`, billForm.date, lines, "bill", id);
    if (!txn) return notify("Entry not balanced");
    let newData = {
      ...data,
      transactions: [...data.transactions, txn],
      bills: [...data.bills, { ...billForm, items: snapshotTaxItems(billForm.items, data.taxGroups), id, status: "unpaid", amountPaid: 0, locked: false }],
      nextBillNum: data.nextBillNum + 1,
      purchaseOrders: data.purchaseOrders.map(x => x.id === po.id ? { ...x, status: "fulfilled" } : x),
    };
    let lots = newData.inventoryLots;
    po.items.forEach(it => {
      if (!it.inventoryId) return;
      newData = adjustInventory(newData, it.inventoryId, Number(it.qty || 0), Number(it.price || 0));
      lots = [...lots, { id: uid("lot"), itemId: it.inventoryId, date: billForm.date, qty: Number(it.qty || 0), remainingQty: Number(it.qty || 0), unitCost: Number(it.price || 0), sourceDocId: id }];
    });
    setData({ ...newData, inventoryLots: lots });
    notify(`${id} created from ${po.id} and posted to the ledger`);
  };

  const cancelOrder = async (po) => {
    if (!(await confirm(`Cancel ${po.id}? It stays on record but won't be available to convert.`))) return;
    setData(d => ({ ...d, purchaseOrders: d.purchaseOrders.map(x => x.id === po.id ? { ...x, status: "cancelled" } : x) }));
    notify("Purchase order cancelled");
  };

  const deleteOrder = async (po) => {
    if (!(await confirm(`Delete ${po.id} entirely? It never posted to the ledger, so this is a clean removal.`))) return;
    setData(d => ({ ...d, purchaseOrders: d.purchaseOrders.filter(x => x.id !== po.id) }));
    notify("Purchase order deleted");
  };

  const sorted = [...data.purchaseOrders].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div>
      <PageHeader eyebrow="Purchases" title="Purchase Orders" sub={`${data.purchaseOrders.length} total - no ledger impact until converted`} action={<button style={styles.btnPrimary} onClick={() => setShowForm(s => !s)}>{showForm ? "Cancel" : "+ New purchase order"}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>New purchase order</div>
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Vendor name" value={form.vendor} onChange={e => setForm({ ...form, vendor: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <input type="date" style={styles.input} value={form.expectedDate} onChange={e => setForm({ ...form, expectedDate: e.target.value })} placeholder="Expected delivery" />
            <select style={styles.input} value={form.expenseAccountId} onChange={e => setForm({ ...form, expenseAccountId: e.target.value })}>{expenseAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
          </div>
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Description</th><th style={styles.th}>Inventory item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Price</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: "100%" }} value={it.desc} onChange={e => updateItem(i, "desc", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.unitCost); } }}>
                  <option value="">None</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name}</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(Number(it.qty || 0) * Number(it.price || 0))}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <RowLine label="Order total" value={total} bold divider />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={createOrder}>Create purchase order</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Order</th><th style={styles.th}>Vendor</th><th style={styles.th}>Expected</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(po => {
            const t = po.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0);
            return (
              <tr key={po.id}><td style={styles.tdMono}>{po.id}</td><td style={styles.td}>{po.vendor}</td><td style={styles.tdMono}>{po.expectedDate ? fmtDate(po.expectedDate) : "-"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td>
                <td style={styles.td}><span style={{ ...styles.pill, ...(po.status === "fulfilled" ? styles.pillGreen : po.status === "cancelled" ? styles.pillRose : styles.pillAmber) }}>{po.status}</span></td>
                <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                  {po.status === "open" && <><button style={styles.iconBtn} onClick={() => convertToBill(po)}>Convert to bill</button>{" "}<button style={styles.iconBtn} onClick={() => cancelOrder(po)}>Cancel</button>{" "}</>}
                  <button style={styles.iconBtn} onClick={() => deleteOrder(po)}>🗑</button>
                </td></tr>
            );
          })}</tbody>
        </table>
        {sorted.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No purchase orders yet.</div>}
      </div>
    </div>
  );
}

/* =================================== Purchase Receipts =================================== */
// Goods physically received before the vendor's bill arrives - stocks the
// inventory immediately against a holding liability (Goods Received Not
// Invoiced), then reclassifies to a formal Accounts Payable once the actual
// bill shows up. A self-contained lifecycle: delete reverses whichever stage
// it's in cleanly, without touching the Bills page's own data or logic.
function PurchaseReceipts({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const blank = () => ({ vendor: "", date: todayStr(), items: [{ desc: "", qty: 1, price: 0, inventoryId: "" }] });
  const [form, setForm] = useState(blank());
  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "" }] }));
  const total = form.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0);

  const createReceipt = () => {
    if (!form.vendor || form.items.some(it => !it.inventoryId)) return notify("Add a vendor, and every line must reference an inventory item");
    const id = `PR-${data.nextPurchaseReceiptNum}`;
    const amt = total;
    const byAccount = {};
    form.items.forEach(it => { const acct = inventoryAccountForItem(data, it.inventoryId); byAccount[acct] = (byAccount[acct] || 0) + Number(it.qty || 0) * Number(it.price || 0); });
    const txn = buildTxn(`Goods received ${id} - ${form.vendor}`, form.date, [...Object.entries(byAccount).map(([acct, a]) => ({ accountId: acct, debit: a, credit: 0 })), { accountId: "2296", debit: 0, credit: amt }], "purchase-receipt", id);
    if (!txn) return notify("Entry not balanced");
    let newData = { ...data, transactions: [...data.transactions, txn], purchaseReceipts: [...data.purchaseReceipts, { ...form, id, status: "pending" }], nextPurchaseReceiptNum: data.nextPurchaseReceiptNum + 1 };
    let lots = newData.inventoryLots;
    form.items.forEach(it => {
      newData = adjustInventory(newData, it.inventoryId, Number(it.qty || 0), Number(it.price || 0));
      lots = [...lots, { id: uid("lot"), itemId: it.inventoryId, date: form.date, qty: Number(it.qty || 0), remainingQty: Number(it.qty || 0), unitCost: Number(it.price || 0), sourceDocId: id }];
    });
    setData({ ...newData, inventoryLots: lots });
    setForm(blank()); setShowForm(false);
    notify(`${id} recorded - goods received, awaiting the vendor's bill`);
  };

  const markBilled = async (pr) => {
    if (!(await confirm(`Mark ${pr.id} as billed? Moves its value from "Goods Received Not Invoiced" to Accounts Payable now that the vendor's invoice has arrived. Inventory isn't touched again - it was already received.`))) return;
    const amt = pr.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0);
    const txn = buildTxn(`Bill matched to ${pr.id} - ${pr.vendor}`, todayStr(), [{ accountId: "2296", debit: amt, credit: 0 }, { accountId: "2000", debit: 0, credit: amt }], "purchase-receipt", pr.id + "-bill");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], purchaseReceipts: d.purchaseReceipts.map(x => x.id === pr.id ? { ...x, status: "billed" } : x) }));
    notify(`${pr.id} matched to Accounts Payable`);
  };

  const deleteReceipt = async (pr) => {
    if (!(await confirm(`Delete ${pr.id}? Its journal entries are removed and the received inventory is reversed. This can't be undone.`))) return;
    setData(d => {
      let next = { ...d, transactions: d.transactions.filter(t => !(t.docId === pr.id || t.docId === pr.id + "-bill")), inventoryLots: d.inventoryLots.filter(l => l.sourceDocId !== pr.id) };
      pr.items.forEach(it => { next = adjustInventory(next, it.inventoryId, -Number(it.qty || 0)); });
      return { ...next, purchaseReceipts: next.purchaseReceipts.filter(x => x.id !== pr.id) };
    });
    notify(`${pr.id} deleted - ledger and stock reversed`);
  };

  const sorted = [...data.purchaseReceipts].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div>
      <PageHeader eyebrow="Purchases" title="Purchase Receipts" sub={`${data.purchaseReceipts.length} total - goods received ahead of the vendor's bill`} action={<button style={styles.btnPrimary} onClick={() => setShowForm(s => !s)}>{showForm ? "Cancel" : "+ New purchase receipt"}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>New purchase receipt</div>
          <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Records goods arriving before the vendor's invoice does - stock goes up immediately, held against a temporary liability until you match it to the actual bill.</div>
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Vendor name" value={form.vendor} onChange={e => setForm({ ...form, vendor: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
          </div>
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Inventory item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty received</th><th style={{ ...styles.th, textAlign: "right" }}>Unit cost</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><select style={{ ...styles.inputSmall, width: "100%" }} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.unitCost); } }}>
                  <option value="">Choose item</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name}</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(Number(it.qty || 0) * Number(it.price || 0))}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <RowLine label="Total received" value={total} bold divider />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={createReceipt}>Record receipt</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Receipt</th><th style={styles.th}>Vendor</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(pr => {
            const t = pr.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0);
            return (
              <tr key={pr.id}><td style={styles.tdMono}>{pr.id}</td><td style={styles.td}>{pr.vendor}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td>
                <td style={styles.td}><span style={{ ...styles.pill, ...(pr.status === "billed" ? styles.pillGreen : styles.pillAmber) }}>{pr.status === "billed" ? "Billed" : "Awaiting bill"}</span></td>
                <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                  {pr.status === "pending" && <button style={styles.iconBtn} onClick={() => markBilled(pr)}>Mark as billed</button>}{" "}
                  <button style={styles.iconBtn} onClick={() => deleteReceipt(pr)}>🗑</button>
                </td></tr>
            );
          })}</tbody>
        </table>
        {sorted.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No purchase receipts yet.</div>}
        <div style={{ fontSize: 11, color: theme.muted, marginTop: 10 }}>Receipts marked as billed post straight to Accounts Payable rather than creating a separate Bill record, so they won't appear individually on the Payables Aging report - the balance is still correct on the Balance Sheet.</div>
      </div>
    </div>
  );
}

/* =================================== Vendor Credits =================================== */
// The purchase-side mirror of Credit Notes - a credit issued BY a vendor,
// reducing what you owe them (or refunded to you in cash).
function VendorCredits({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const [viewingId, setViewingId] = useState(null);
  const openBills = data.bills.filter(b => (b.amountPaid || 0) < computeDocTotals(b, data.taxGroups).finalAmount - 0.5);
  const blank = () => ({ vendor: "", date: todayStr(), billId: "", refundBankId: "", reason: "", items: [{ desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }], taxes: [] });
  const [form, setForm] = useState(blank());
  const updateItem = (i, field, val) => setForm(f => ({ ...f, items: f.items.map((it, idx) => idx === i ? { ...it, [field]: val } : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { desc: "", qty: 1, price: 0, inventoryId: "", taxGroupId: "" }] }));
  const totals = computeDocTotals(form, data.taxGroups);

  const create = () => {
    if (!form.vendor || form.items.some(it => !it.desc)) return notify("Add a vendor and describe each line item");
    const id = `VC-${data.nextVendorCreditNum}`;
    const lines = vendorCreditJournalLines(form, data, form.refundBankId || null);
    const txn = buildTxn(`Vendor credit ${id} - ${form.vendor}${form.billId ? ` (applied to ${form.billId})` : ""}`, form.date, lines, "vendor-credit", id);
    if (!txn) return notify("Entry not balanced");
    let newData = {
      ...data,
      transactions: [...data.transactions, txn],
      vendorCredits: [...data.vendorCredits, { ...form, items: snapshotTaxItems(form.items, data.taxGroups), id, status: "issued" }],
      nextVendorCreditNum: data.nextVendorCreditNum + 1,
    };
    if (form.billId) newData.bills = newData.bills.map(b => b.id === form.billId ? { ...b, amountPaid: (b.amountPaid || 0) + totals.finalAmount } : b);
    let lots = newData.inventoryLots;
    form.items.forEach(it => {
      if (!it.inventoryId) return;
      newData = adjustInventory(newData, it.inventoryId, -Number(it.qty || 0));
      const r = consumeFIFO(lots, it.inventoryId, Number(it.qty || 0));
      lots = r.lots;
    });
    setData({ ...newData, inventoryLots: lots });
    setForm(blank()); setShowForm(false);
    notify(`${id} issued${form.refundBankId ? " and refunded" : form.billId ? ` and applied to ${form.billId}` : ""}`);
  };

  const deleteDoc = async (doc) => {
    if (!(await confirm(`Delete ${doc.id}? Its journal entry is removed and any returned inventory is restored. This can't be undone.`))) return;
    setData(d => {
      let next = { ...d, transactions: stripDocTransactions(d.transactions, doc.id) };
      let lots = next.inventoryLots;
      (doc.items || []).forEach(it => {
        if (!it.inventoryId) return;
        next = adjustInventory(next, it.inventoryId, Number(it.qty || 0));
        lots = restoreFIFO(lots, it.inventoryId, Number(it.qty || 0));
      });
      if (doc.billId) {
        const t = computeDocTotals(doc, d.taxGroups).finalAmount;
        next = { ...next, bills: next.bills.map(b => b.id === doc.billId ? { ...b, amountPaid: Math.max(0, (b.amountPaid || 0) - t) } : b) };
      }
      return { ...next, inventoryLots: lots, vendorCredits: next.vendorCredits.filter(x => x.id !== doc.id) };
    });
    setViewingId(null);
    notify(`${doc.id} deleted`);
  };

  const sorted = [...data.vendorCredits].sort((a, b) => b.date.localeCompare(a.date));
  const viewing = viewingId ? data.vendorCredits.find(i => i.id === viewingId) : null;
  if (viewing) return <DocumentDetail doc={viewing} docType="bill" data={data} onBack={() => setViewingId(null)} onEdit={() => {}} onToggleLock={() => {}} onDelete={() => deleteDoc(viewing)} />;

  return (
    <div>
      <PageHeader eyebrow="Purchases" title="Vendor Credits" sub={`${data.vendorCredits.length} total - credits received from vendors`} action={<button style={styles.btnPrimary} onClick={() => setShowForm(s => !s)}>{showForm ? "Cancel" : "+ New vendor credit"}</button>} />
      {showForm && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>New vendor credit</div>
          <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
            <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Vendor name" value={form.vendor} onChange={e => setForm({ ...form, vendor: e.target.value })} />
            <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            <select style={styles.input} value={form.billId} onChange={e => setForm({ ...form, billId: e.target.value })}>
              <option value="">Not applied to a bill</option>{openBills.map(b => <option key={b.id} value={b.id}>Apply to {b.id}</option>)}
            </select>
            <select style={styles.input} value={form.refundBankId} onChange={e => setForm({ ...form, refundBankId: e.target.value })}>
              <option value="">Credit to Accounts Payable</option>{data.banks.map(b => <option key={b.id} value={b.id}>Refunded to {b.name}</option>)}
            </select>
          </div>
          <input style={{ ...styles.input, width: "100%", marginTop: 10 }} placeholder="Reason (optional)" value={form.reason} onChange={e => setForm({ ...form, reason: e.target.value })} />
          <table style={{ ...styles.table, marginTop: 12 }}>
            <thead><tr><th style={styles.th}>Description</th><th style={styles.th}>Inventory item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Price</th><th style={styles.th}>Tax group</th><th style={{ ...styles.th, textAlign: "right" }}>Line total</th></tr></thead>
            <tbody>{form.items.map((it, i) => (
              <tr key={i}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: "100%" }} value={it.desc} onChange={e => updateItem(i, "desc", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.inventoryId} onChange={e => { const inv = data.inventory.find(x => x.id === e.target.value); updateItem(i, "inventoryId", e.target.value); if (inv) { updateItem(i, "desc", inv.name); updateItem(i, "price", inv.unitCost); } }}>
                  <option value="">None</option>{data.inventory.map(inv => <option key={inv.id} value={inv.id}>{inv.name}</option>)}</select></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={it.qty} onChange={e => updateItem(i, "qty", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={it.price} onChange={e => updateItem(i, "price", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={it.taxGroupId} onChange={e => updateItem(i, "taxGroupId", e.target.value)}>
                  <option value="">No tax group</option>{data.taxGroups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}</select></td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(totals.lineCalcs[i]?.lineTotal || 0)}</td>
              </tr>
            ))}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addItem}>+ line item</button>
          <LineTaxSummary lineCalcs={totals.lineCalcs} subtotal={totals.subtotal} afterLineTax={totals.afterLineTax} />
          <RowLine label="Total credit" value={totals.finalAmount} bold divider />
          <button style={{ ...styles.btnPrimary, marginTop: 12 }} onClick={create}>Issue vendor credit</button>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Credit</th><th style={styles.th}>Vendor</th><th style={styles.th}>Applied to</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(doc => {
            const t = computeDocTotals(doc, data.taxGroups).finalAmount;
            return (
              <tr key={doc.id} onClick={() => setViewingId(doc.id)} style={{ cursor: "pointer" }}>
                <td style={styles.tdMono}>{doc.id}</td><td style={styles.td}>{doc.vendor}</td><td style={styles.td}>{doc.billId || (doc.refundBankId ? `Refunded via ${data.banks.find(b => b.id === doc.refundBankId)?.name}` : "Open credit")}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td>
                <td style={styles.td} onClick={e => e.stopPropagation()}><button style={styles.iconBtn} onClick={() => deleteDoc(doc)}>🗑</button></td></tr>
            );
          })}</tbody>
        </table>
        {sorted.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>None issued yet.</div>}
      </div>
    </div>
  );
}
function Expenses({ data, setData, postTransaction, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const expenseAccounts = activeAccounts(data).filter(a => a.type === "expense");
  const [form, setForm] = useState({ vendor: "", date: todayStr(), category: expenseAccounts[0]?.id, bankId: data.banks[0]?.id, projectId: "", amount: "", deductible: true, memo: "" });
  const submit = () => {
    if (!form.vendor || !form.amount) return notify("Add a vendor and amount");
    const bank = data.banks.find(b => b.id === form.bankId);
    const txn = buildTxn(form.memo || form.vendor, form.date, [{ accountId: form.category, debit: Number(form.amount), credit: 0 }, { accountId: bank?.accountId || "1000", debit: 0, credit: Number(form.amount) }], "expense");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], expenses: [...d.expenses, { ...form, id: uid("exp"), amount: Number(form.amount), txnId: txn.id }] }));
    setForm({ ...form, vendor: "", amount: "", memo: "" }); notify("Expense recorded");
  };
  // Deleting an expense removes its journal entry too. Older expense records
  // (before entries were linked) fall back to matching by source, date,
  // category and amount so their entry is still found and removed.
  const findExpenseTxn = (e) => {
    if (e.txnId) return data.transactions.find(t => t.id === e.txnId);
    return data.transactions.find(t => t.source === "expense" && t.date === e.date &&
      t.lines.some(l => l.accountId === e.category && Math.abs(l.debit - e.amount) < 0.005));
  };
  const deleteExpense = async (e) => {
    const txn = findExpenseTxn(e);
    if (isLocked(e.date, data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - this expense falls within that period and can't be deleted`);
    if (!(await confirm(`Delete this ${fmt(e.amount)} expense (${e.vendor})? ${txn ? "Its journal entry is removed with it." : "Warning: its journal entry couldn't be located automatically - check the Journal after deleting."} This can't be undone.`))) return;
    setData(d => ({
      ...d,
      expenses: d.expenses.filter(x => x.id !== e.id),
      transactions: txn ? d.transactions.filter(t => t.id !== txn.id) : d.transactions,
    }));
    notify(txn ? "Expense and its journal entry deleted" : "Expense record deleted");
  };
  const byCategory = expenseAccounts.map(a => ({ name: a.name, value: data.expenses.filter(e => e.category === a.id).reduce((s, e) => s + e.amount, 0) })).filter(c => c.value > 0);
  const sorted = [...data.expenses].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div>
      <PageHeader eyebrow="Spend" title="Quick Expenses" sub="For simple, single-line, tax-free spend" />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Record expense</div>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Vendor" value={form.vendor} onChange={e => setForm({ ...form, vendor: e.target.value })} />
          <input type="date" style={styles.input} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
          <select style={styles.input} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{expenseAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
          <select style={styles.input} value={form.bankId} onChange={e => setForm({ ...form, bankId: e.target.value })}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
          <select style={styles.input} value={form.projectId} onChange={e => setForm({ ...form, projectId: e.target.value })}><option value="">No project</option>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
          <input type="number" style={{ ...styles.input, width: 110 }} placeholder="Amount" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
          <button style={styles.btnPrimary} onClick={submit}>Save</button>
        </div>
      </div>
      {byCategory.length > 0 && (
        <div className="ces-card" style={styles.cardWide}>
          <div style={styles.cardTitle}>Spend by category</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={byCategory} layout="vertical" margin={{ left: 20 }}>
              <CartesianGrid strokeDasharray="2 4" stroke={theme.border} horizontal={false} />
              <XAxis type="number" tick={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, fill: theme.muted }} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="name" width={140} tick={{ fontFamily: "Inter", fontSize: 12, fill: theme.text }} axisLine={false} tickLine={false} />
              <Tooltip formatter={v => fmt(v)} contentStyle={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 12, background: theme.panel }} />
              <Bar dataKey="value" fill={theme.amber} radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Vendor</th><th style={styles.th}>Category</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(e => (<tr key={e.id}><td style={styles.tdMono}>{fmtDate(e.date)}</td><td style={styles.td}>{e.vendor}</td><td style={styles.td}>{data.accounts.find(a => a.id === e.category)?.name || "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(e.amount)}</td><td style={styles.td}><button style={styles.iconBtn} onClick={() => deleteExpense(e)}>🗑</button></td></tr>))}</tbody>
        </table>
      </div>
    </div>
  );
}

/* =================================== Inventory =================================== */
// IAS 2: writes inventory down to net realizable value where NRV has fallen
// below cost, or reverses a prior write-down (capped at original cost) if
// NRV has since recovered.
// Turns raw materials (or consumables) into a finished good or WIP item -
// the shawarma-from-bread-and-chicken, cement-and-blocks-into-a-building
// case. Consumed items' value moves out of their inventory account into
// whatever account the produced item belongs to, so the GL always reflects
// where value actually sits, not just where it started.
function ProductionPanel({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const producibleItems = data.inventory.filter(i => ["finished_good", "wip"].includes(i.inventoryType || "finished_good"));
  const consumableItems = data.inventory.filter(i => ["raw_material", "consumable", "wip"].includes(i.inventoryType));
  const blank = () => ({ producedItemId: producibleItems[0]?.id || "", producedQty: "", date: todayStr(), consumedLines: [{ itemId: consumableItems[0]?.id || "", qty: "" }] });
  const [form, setForm] = useState(blank());

  const updateLine = (i, field, val) => setForm(f => ({ ...f, consumedLines: f.consumedLines.map((l, idx) => idx === i ? { ...l, [field]: val } : l) }));
  const addLine = () => setForm(f => ({ ...f, consumedLines: [...f.consumedLines, { itemId: consumableItems[0]?.id || "", qty: "" }] }));
  const removeLine = (i) => setForm(f => ({ ...f, consumedLines: f.consumedLines.filter((_, idx) => idx !== i) }));

  const costByAccount = {};
  let totalCost = 0;
  form.consumedLines.forEach(l => {
    const item = data.inventory.find(x => x.id === l.itemId);
    if (!item || !Number(l.qty)) return;
    const cost = item.unitCost * Number(l.qty);
    totalCost += cost;
    const acct = inventoryAccountForItem(data, l.itemId);
    costByAccount[acct] = (costByAccount[acct] || 0) + cost;
  });
  const producedItem = data.inventory.find(i => i.id === form.producedItemId);
  const unitCostProduced = Number(form.producedQty) > 0 ? totalCost / Number(form.producedQty) : 0;

  // Exactly undoes one production record: removes its transaction, restores
  // the consumed materials (quantity and FIFO lots), and backs the produced
  // item's quantity and weighted-average cost out using the same math
  // adjustInventory used going forward, just in reverse.
  const reverseRecord = (d, record) => {
    let next = { ...d, transactions: d.transactions.filter(t => t.docId !== record.id) };
    let lots = next.inventoryLots;
    record.consumedLines.forEach(l => {
      next = { ...next, inventory: next.inventory.map(i => i.id === l.itemId ? { ...i, qty: i.qty + Number(l.qty) } : i) };
      lots = restoreFIFO(lots, l.itemId, Number(l.qty));
    });
    next = { ...next, inventory: next.inventory.map(i => {
      if (i.id !== record.producedItemId) return i;
      const newQty = i.qty - record.producedQty;
      if (newQty <= 0.0005) return { ...i, qty: Math.max(0, newQty), unitCost: 0 };
      if (record.totalCost > 0) {
        const newTotalValue = i.qty * i.unitCost - record.totalCost;
        return { ...i, qty: newQty, unitCost: newTotalValue / newQty };
      }
      return { ...i, qty: newQty };
    }) };
    lots = lots.filter(l => !(l.sourceDocId === record.id && l.itemId === record.producedItemId));
    return { ...next, inventoryLots: lots };
  };

  const startEdit = (record) => {
    if (isLocked(record.date, data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - this production run falls within that period and can't be edited`);
    setForm({ producedItemId: record.producedItemId, producedQty: String(record.producedQty), date: record.date, consumedLines: record.consumedLines.map(l => ({ itemId: l.itemId, qty: String(l.qty) })) });
    setEditingId(record.id);
    setShowForm(true);
  };
  const cancelEdit = () => { setEditingId(null); setForm(blank()); setShowForm(false); };

  const deleteRecord = async (record) => {
    if (isLocked(record.date, data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - this production run falls within that period and can't be deleted`);
    if (!(await confirm(`Delete this production run? ${record.producedQty} × ${data.inventory.find(i => i.id === record.producedItemId)?.name || "item"} and the materials it used are reversed.`))) return;
    setData(d => {
      const reversed = reverseRecord(d, record);
      return { ...reversed, productionRecords: reversed.productionRecords.filter(r => r.id !== record.id) };
    });
    notify("Production run deleted and reversed");
  };

  const record = async () => {
    if (!form.producedItemId || !Number(form.producedQty)) return notify("Choose what's being produced and how much");
    const validLines = form.consumedLines.filter(l => l.itemId && Number(l.qty) > 0);
    if (validLines.length === 0) return notify("Add at least one raw material or consumable that's being used up");
    if (editingId && isLocked(form.date, data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - choose a later date`);

    // Editing: work from a version of the data with the old record's
    // effects already undone, so re-checking stock and re-posting starts
    // from a clean slate rather than double-counting the original run.
    const baseData = editingId ? reverseRecord(data, data.productionRecords.find(r => r.id === editingId)) : data;
    const shortages = validLines.map(l => { const item = baseData.inventory.find(x => x.id === l.itemId); return item && Number(l.qty) > item.qty ? `${item.name} (have ${item.qty}, using ${l.qty})` : null; }).filter(Boolean);
    const confirmMsg = shortages.length > 0
      ? `This takes ${shortages.join(", ")} below zero on hand - your recorded stock may be behind what you actually have. ${editingId ? "Save anyway" : "Record it anyway"}?`
      : `${editingId ? "Save changes to" : "Record"} production of ${form.producedQty} × ${producedItem?.name}, consuming ${validLines.length} item(s)${totalCost > 0 ? ` worth ${fmt(totalCost)} total` : ""}?`;
    if (!(await confirm(confirmMsg))) return;

    const costByAcct2 = {};
    let totalCost2 = 0;
    validLines.forEach(l => { const item = baseData.inventory.find(x => x.id === l.itemId); if (!item) return; const cost = item.unitCost * Number(l.qty); totalCost2 += cost; const acct = inventoryAccountForItem(baseData, l.itemId); costByAcct2[acct] = (costByAcct2[acct] || 0) + cost; });
    const producedAccount = inventoryAccountForItem(baseData, form.producedItemId);
    const id = editingId || `PROD-${baseData.nextProductionNum}`;
    let newData = baseData;
    if (totalCost2 > 0) {
      const lines = [
        { accountId: producedAccount, debit: totalCost2, credit: 0 },
        ...Object.entries(costByAcct2).map(([acct, amt]) => ({ accountId: acct, debit: 0, credit: amt })),
      ];
      const txn = buildTxn(`Production - ${form.producedQty} x ${producedItem?.name}`, form.date, lines, "production", id);
      if (!txn) return notify("Entry not balanced");
      newData = { ...newData, transactions: [...newData.transactions, txn] };
    }

    const unitCostProduced2 = totalCost2 / Number(form.producedQty);
    let lots = newData.inventoryLots;
    validLines.forEach(l => {
      newData = { ...newData, inventory: newData.inventory.map(i => i.id === l.itemId ? { ...i, qty: i.qty - Number(l.qty) } : i) };
      lots = consumeFIFO(lots, l.itemId, Number(l.qty)).lots;
    });
    newData = adjustInventory(newData, form.producedItemId, Number(form.producedQty), totalCost2 > 0 ? unitCostProduced2 : null);
    lots = [...lots, { id: uid("lot"), itemId: form.producedItemId, date: form.date, qty: Number(form.producedQty), remainingQty: Number(form.producedQty), unitCost: totalCost2 > 0 ? unitCostProduced2 : (baseData.inventory.find(i => i.id === form.producedItemId)?.unitCost || 0), sourceDocId: id }];

    const newRecord = { id, date: form.date, producedItemId: form.producedItemId, producedQty: Number(form.producedQty), consumedLines: validLines, totalCost: totalCost2 };
    setData({
      ...newData, inventoryLots: lots,
      productionRecords: editingId ? newData.productionRecords.map(r => r.id === editingId ? newRecord : r) : [...newData.productionRecords, newRecord],
      nextProductionNum: editingId ? newData.nextProductionNum : newData.nextProductionNum + 1,
    });
    setForm(blank()); setShowForm(false); setEditingId(null);
    notify(editingId ? "Production run updated" : `Recorded: ${form.producedQty} × ${producedItem?.name} produced, ${fmt(totalCost2)} of materials used`);
  };

  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <div>
          <div style={styles.cardTitle}>Record usage / production<InfoTip text="Turn raw materials or consumables into a finished item - bread and chicken into shawarma, cement and blocks into a unit under construction." /></div>
        </div>
        <button style={styles.btnPrimary} onClick={() => showForm ? cancelEdit() : setShowForm(true)}>{showForm ? "Cancel" : "+ Record production"}</button>
      </div>
      {showForm && (
        <div style={{ marginTop: 14 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
            <label style={{ fontSize: 12 }}>Producing
              <select style={{ ...styles.input, display: "block", marginTop: 4, minWidth: 200 }} value={form.producedItemId} onChange={e => setForm({ ...form, producedItemId: e.target.value })}>
                {producibleItems.length === 0 && <option value="">No finished good / WIP items yet</option>}
                {producibleItems.map(i => <option key={i.id} value={i.id}>{i.name}</option>)}
              </select>
            </label>
            <label style={{ fontSize: 12 }}>Quantity produced
              <input type="number" style={{ ...styles.input, display: "block", marginTop: 4, width: 120 }} value={form.producedQty} onChange={e => setForm({ ...form, producedQty: e.target.value })} />
            </label>
            <label style={{ fontSize: 12 }}>Date
              <input type="date" style={{ ...styles.input, display: "block", marginTop: 4 }} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
            </label>
          </div>

          <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginTop: 16, marginBottom: 6 }}>Materials used</div>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>On hand</th><th style={{ ...styles.th, textAlign: "right" }}>Qty used</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={styles.th}></th></tr></thead>
            <tbody>{form.consumedLines.map((l, i) => {
              const item = data.inventory.find(x => x.id === l.itemId);
              return (
                <tr key={i}>
                  <td style={styles.td}><select style={styles.inputSmall} value={l.itemId} onChange={e => updateLine(i, "itemId", e.target.value)}>
                    {consumableItems.length === 0 && <option value="">No raw material / consumable items yet</option>}
                    {consumableItems.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}
                  </select></td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{item?.qty ?? "-"}</td>
                  <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 80, textAlign: "right" }} value={l.qty} onChange={e => updateLine(i, "qty", e.target.value)} /></td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{item && l.qty ? fmt(item.unitCost * Number(l.qty)) : "-"}</td>
                  <td style={styles.td}><button style={styles.iconBtn} onClick={() => removeLine(i)}>🗑</button></td>
                </tr>
              );
            })}</tbody>
          </table>
          <button style={styles.btnGhost} onClick={addLine}>+ material</button>

          <div style={{ marginTop: 14, padding: "10px 14px", background: theme.mode === "dark" ? "rgba(255,255,255,0.04)" : "rgba(0,0,0,0.02)", borderRadius: 10 }}>
            <RowLine label="Total material cost" value={totalCost} bold />
            <div style={{ fontSize: 12, color: theme.muted, marginTop: 2 }}>{totalCost > 0 ? `Unit cost of ${producedItem?.name || "produced item"}: ${form.producedQty ? fmt(unitCostProduced) : "-"} per unit` : "No cost entered on the materials used yet - quantities will update, but no value moves in the ledger until you set unit costs in Inventory."}</div>
          </div>
          <button style={{ ...styles.btnPrimary, marginTop: 14 }} onClick={record}>{editingId ? "Save changes" : "Record production"}</button>
        </div>
      )}
      {data.productionRecords.length > 0 && (
        <div style={{ marginTop: 18 }}>
          <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>History</div>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Produced</th><th style={styles.th}>Materials used</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={styles.th}></th></tr></thead>
            <tbody>{[...data.productionRecords].sort((a, b) => b.date.localeCompare(a.date)).map(r => {
              const locked = isLocked(r.date, data);
              const producedName = data.inventory.find(i => i.id === r.producedItemId)?.name || "(deleted item)";
              const materialsLabel = r.consumedLines.map(l => `${l.qty} × ${data.inventory.find(i => i.id === l.itemId)?.name || "(deleted item)"}`).join(", ");
              return (
                <tr key={r.id}>
                  <td style={styles.tdMono}>{locked && <span title="Locked" style={{ marginRight: 5 }}>🔒</span>}{fmtDate(r.date)}</td>
                  <td style={styles.td}>{r.producedQty} × {producedName}</td>
                  <td style={{ ...styles.td, fontSize: 12 }}>{materialsLabel}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.totalCost)}</td>
                  <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                    <button style={{ ...styles.iconBtn, color: locked ? theme.muted : theme.text, cursor: locked ? "not-allowed" : "pointer" }} disabled={locked} onClick={() => startEdit(r)}>✎</button>{" "}
                    <button style={{ ...styles.iconBtn, color: locked ? theme.muted : theme.rose, cursor: locked ? "not-allowed" : "pointer" }} disabled={locked} onClick={() => deleteRecord(r)}>🗑</button>
                  </td>
                </tr>
              );
            })}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}
function NRVPanel({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const rows = computeNRVCheck(data);
  const totalWriteDown = rows.reduce((s, r) => s + r.writeDownAmount, 0);
  const totalReversal = rows.reduce((s, r) => s + r.reversalAmount, 0);
  const net = totalWriteDown - totalReversal;

  const post = async () => {
    if (rows.length === 0) return notify("No items need a write-down or reversal - net realizable value covers cost everywhere");
    if (!(await confirm(`Post inventory NRV adjustment? ${totalWriteDown > 0.5 ? `${fmt(totalWriteDown)} write-down` : ""}${totalWriteDown > 0.5 && totalReversal > 0.5 ? " and " : ""}${totalReversal > 0.5 ? `${fmt(totalReversal)} reversal` : ""} across ${rows.length} item(s).`))) return;
    const netByAccount = {};
    rows.forEach(r => { const acct = inventoryAccountForItem(data, r.id); netByAccount[acct] = (netByAccount[acct] || 0) + (r.writeDownAmount - r.reversalAmount); });
    const lines = [];
    let expenseNet = 0;
    Object.entries(netByAccount).forEach(([acct, net]) => {
      if (net > 0.5) { lines.push({ accountId: acct, debit: 0, credit: net }); expenseNet += net; }
      else if (net < -0.5) { lines.push({ accountId: acct, debit: -net, credit: 0 }); expenseNet += net; }
    });
    if (Math.abs(expenseNet) > 0.5) lines.push(expenseNet > 0 ? { accountId: "5750", debit: expenseNet, credit: 0 } : { accountId: "5750", debit: 0, credit: -expenseNet });
    const txn = lines.length ? buildTxn("Inventory NRV adjustment", todayStr(), lines, "manual") : null;
    if (lines.length && !txn) return notify("Entry not balanced");
    setData(d => ({
      ...d,
      transactions: txn ? [...d.transactions, txn] : d.transactions,
      inventory: d.inventory.map(item => {
        const r = rows.find(x => x.id === item.id);
        if (!r) return item;
        return { ...item, unitCost: r.newUnitCost, costBeforeNRV: item.costBeforeNRV ?? r.cost };
      }),
    }));
    notify("Inventory adjusted to net realizable value");
  };

  if (rows.length === 0) return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Net realizable value check</div>
      <div style={{ fontSize: 13, color: theme.muted, marginTop: 8 }}>Every item's net realizable value covers its cost - no write-down needed. Set an item's NRV below its unit cost in the table above if it's damaged, obsolete, or its selling price has dropped.</div>
    </div>
  );
  return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Net realizable value check</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Inventory is carried at the lower of cost and net realizable value. Items below are written down to NRV; a later recovery in NRV reverses the write-down, capped at the original cost.</div>
      <table style={{ ...styles.table, marginTop: 10 }}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={{ ...styles.th, textAlign: "right" }}>NRV</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
        <tbody>{rows.map(r => (
          <tr key={r.id}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.cost)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.nrv)}</td>
            <td style={{ ...styles.tdMono, textAlign: "right", color: r.writeDownAmount > 0 ? theme.rose : theme.emerald }}>{r.writeDownAmount > 0 ? `-${fmt(r.writeDownAmount)}` : `+${fmt(r.reversalAmount)}`}</td></tr>
        ))}</tbody>
      </table>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginTop: 14 }}>
        <Kpi label="Total write-down" value={fmt(totalWriteDown)} tone="rose" />
        <Kpi label="Total reversal" value={fmt(totalReversal)} tone="emerald" />
        <Kpi label="Net entry" value={fmt(net)} tone={net > 0 ? "rose" : "emerald"} />
      </div>
      <button style={{ ...styles.btnPrimary, marginTop: 14 }} onClick={post}>Post NRV adjustment</button>
    </div>
  );
}
function Inventory({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const orderedTypes = inventoryTypesForIndustry(data.settings.industry);
  const blank = () => ({ sku: "", name: "", qty: "", unitCost: "", salePrice: "", reorderLevel: "", inventoryType: orderedTypes[0].id, locationId: "", departmentId: "" });
  const [form, setForm] = useState(blank());
  const addItem = () => {
    if (!form.name) return notify("Name the item");
    setData(d => ({ ...d, inventory: [...d.inventory, { id: uid("inv"), sku: form.sku, name: form.name, qty: Number(form.qty) || 0, unitCost: Number(form.unitCost) || 0, salePrice: Number(form.salePrice) || 0, reorderLevel: Number(form.reorderLevel) || 0, inventoryType: form.inventoryType, locationId: form.locationId, departmentId: form.departmentId }] }));
    setForm(blank());
    notify("Item added to inventory");
  };
  // Delete an inventory item the right approach: if stock remains, its value
  // is written off the books (Dr loss / Cr Inventory) so the balance sheet
  // stays true; then the item and its lots are removed.
  const deleteItem = async (i) => {
    const value = i.qty * i.unitCost;
    const msg = i.qty > 0
      ? `${i.name} still has ${i.qty} on hand worth ${fmt(value)}. Deleting will post a stock write-off to the ledger and remove the item. Continue?`
      : `Delete ${i.name} from inventory? Its lot history is removed too.`;
    if (!(await confirm(msg))) return;
    setData(d => {
      let next = d;
      if (i.qty > 0 && value > 0.005) {
        const txn = buildTxn(`Inventory write-off - ${i.name}`, todayStr(), [{ accountId: "6000", debit: value, credit: 0 }, { accountId: inventoryAccountForItem(d, i.id), debit: 0, credit: value }], "manual");
        if (txn) next = { ...next, transactions: [...next.transactions, txn] };
      }
      return { ...next, inventory: next.inventory.filter(x => x.id !== i.id), inventoryLots: next.inventoryLots.filter(l => l.itemId !== i.id) };
    });
    notify(i.qty > 0 ? `${i.name} deleted - ${fmt(value)} written off` : `${i.name} deleted`);
  };

  const totalValue = data.inventory.reduce((s, i) => s + i.qty * i.unitCost, 0);
  return (
    <div>
      <PageHeader eyebrow="Assets" title="Inventory" sub={`${data.inventory.length} items · ${fmt(totalValue)} on hand`} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Add item</div>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
          <input style={styles.input} placeholder="SKU" value={form.sku} onChange={e => setForm({ ...form, sku: e.target.value })} />
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Item name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
          <select style={styles.input} value={form.inventoryType} onChange={e => setForm({ ...form, inventoryType: e.target.value })}>
            {orderedTypes.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
          </select>
          <input type="number" style={{ ...styles.input, width: 90 }} placeholder="Qty" value={form.qty} onChange={e => setForm({ ...form, qty: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 110 }} placeholder="Unit cost" value={form.unitCost} onChange={e => setForm({ ...form, unitCost: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 110 }} placeholder="Sale price" value={form.salePrice} onChange={e => setForm({ ...form, salePrice: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 110 }} placeholder="Reorder level" value={form.reorderLevel} onChange={e => setForm({ ...form, reorderLevel: e.target.value })} />
          <select style={styles.input} value={form.locationId} onChange={e => setForm({ ...form, locationId: e.target.value })}><option value="">No location</option>{data.locations.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select>
          <select style={styles.input} value={form.departmentId} onChange={e => setForm({ ...form, departmentId: e.target.value })}><option value="">No department</option>{data.departments.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}</select>
          <button style={styles.btnPrimary} onClick={addItem}>Add</button>
        </div>
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={{ overflowX: "auto" }}>
        <table style={{ ...styles.table, minWidth: 1280 }}>
          <thead><tr><th style={styles.th}>SKU</th><th style={styles.th}>Item</th><th style={styles.th}>Type</th><th style={styles.th}>Location</th><th style={styles.th}>Department</th><th style={{ ...styles.th, textAlign: "right" }}>On hand</th><th style={{ ...styles.th, textAlign: "right" }}>Unit cost</th><th style={{ ...styles.th, textAlign: "right" }}>Sale price</th><th style={{ ...styles.th, textAlign: "right" }}>Net realizable value</th><th style={{ ...styles.th, textAlign: "right" }}>Reorder level</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th><th style={styles.th}></th></tr></thead>
          <tbody>{data.inventory.map(i => (
            <InventoryRow key={i.id} item={i} data={data} setData={setData} notify={notify} deleteItem={deleteItem} />
          ))}</tbody>
        </table>
        </div>
      </div>
      <div style={{ fontSize: 12, color: theme.muted, display: "flex", alignItems: "center" }}>How editing works here<InfoTip text={'SKU, name, sale price, and reorder level are administrative details - edit them freely, anytime. Quantity and unit cost carry real accounting value, so changing either posts a proper adjusting entry once you confirm it, the same way a stock count correction or cost correction would in any accounting system.'} /></div>
      <ProductionPanel data={data} setData={setData} notify={notify} />
      <NRVPanel data={data} setData={setData} notify={notify} />
    </div>
  );
}
// A single inventory row with two editing tiers: plain fields (SKU, name,
// sale price, reorder level) commit immediately on blur since they carry no
// accounting value. Qty and unit cost stage locally as you type and only
// post - as a real, confirmed adjusting journal entry - once you click
// Apply, so nothing silently drifts from the ledger. Locked periods block
// the adjustment the same way they'd block any other dated entry.
function InventoryRow({ item: i, data, setData, notify, deleteItem }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(null);

  const startEdit = () => {
    setDraft({ sku: i.sku, name: i.name, inventoryType: i.inventoryType || "finished_good", qty: String(i.qty), unitCost: String(i.unitCost), salePrice: String(i.salePrice), nrv: String(i.nrv ?? i.salePrice), reorderLevel: String(i.reorderLevel), locationId: i.locationId || "", departmentId: i.departmentId || "" });
    setEditing(true);
  };
  const cancelEdit = () => { setDraft(null); setEditing(false); };
  const setDraftField = (field, val) => setDraft(d => ({ ...d, [field]: val }));

  const save = async () => {
    const pendingQty = Number(draft.qty) || 0, pendingCost = Number(draft.unitCost) || 0;
    const oldAccount = inventoryAccountForItem(data, i.id);
    const newAccount = inventoryAccountForType(draft.inventoryType);
    const oldValue = i.qty * i.unitCost, newValue = pendingQty * pendingCost;
    const valueChanged = Math.abs(newValue - oldValue) > 0.005 || newAccount !== oldAccount;

    if (valueChanged && isLocked(todayStr(), data)) return notify(`Books are locked through ${fmtDate(data.settings.lockDate)} - today's date falls within that period, so this change can't be posted right now`);
    if (valueChanged) {
      const parts = [];
      if (newAccount !== oldAccount) parts.push(`reclassifies it to ${INVENTORY_TYPES.find(t => t.id === draft.inventoryType)?.label}`);
      if (Math.abs(pendingQty * pendingCost - i.qty * i.unitCost) > 0.005 || newAccount === oldAccount) parts.push(`changes its value by ${newValue - oldValue >= 0 ? "+" : ""}${fmt(newValue - oldValue)}`);
      if (!(await confirm(`Save changes to ${draft.name}? This ${parts.join(" and ")}, posting a correcting entry to the ledger.`))) return;
    }

    setData(d => {
      let next = { ...d, inventory: d.inventory.map(x => x.id === i.id ? { ...x, sku: draft.sku, name: draft.name, inventoryType: draft.inventoryType, qty: pendingQty, unitCost: pendingCost, salePrice: Number(draft.salePrice) || 0, nrv: Number(draft.nrv) || 0, reorderLevel: Number(draft.reorderLevel) || 0, locationId: draft.locationId, departmentId: draft.departmentId } : x) };
      if (valueChanged) {
        const lines = [];
        if (newAccount !== oldAccount) {
          lines.push({ accountId: newAccount, debit: newValue, credit: 0 }, { accountId: oldAccount, debit: 0, credit: oldValue });
          const resid = newValue - oldValue;
          if (Math.abs(resid) > 0.005) lines.push(resid > 0 ? { accountId: "6000", debit: 0, credit: resid } : { accountId: "6000", debit: -resid, credit: 0 });
        } else {
          const diff = newValue - oldValue;
          lines.push(diff > 0 ? { accountId: oldAccount, debit: diff, credit: 0 } : { accountId: oldAccount, debit: 0, credit: -diff });
          lines.push(diff > 0 ? { accountId: "6000", debit: 0, credit: diff } : { accountId: "6000", debit: -diff, credit: 0 });
        }
        const txn = buildTxn(`Stock adjustment - ${draft.name}`, todayStr(), lines, "manual");
        if (txn) next = { ...next, transactions: [...next.transactions, txn] };
      }
      return next;
    });
    setDraft(null); setEditing(false);
    notify(`${draft.name} updated`);
  };

  if (!editing) {
    return (
      <tr>
        <td style={styles.tdMono}>{i.sku}</td>
        <td style={styles.td}>{i.name}</td>
        <td style={styles.td}>{INVENTORY_TYPES.find(t => t.id === (i.inventoryType || "finished_good"))?.label}</td>
        <td style={styles.td}>{data.locations.find(l => l.id === i.locationId)?.name || "-"}</td>
        <td style={styles.td}>{data.departments.find(x => x.id === i.departmentId)?.name || "-"}</td>
        <td style={{ ...styles.tdMono, textAlign: "right", color: i.qty <= i.reorderLevel ? theme.rose : theme.text }}>{i.qty}</td>
        <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.unitCost)}</td>
        <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.salePrice)}</td>
        <td style={{ ...styles.tdMono, textAlign: "right", color: (i.nrv ?? i.salePrice) < i.unitCost ? theme.rose : theme.text }}>{fmt(i.nrv ?? i.salePrice)}</td>
        <td style={{ ...styles.tdMono, textAlign: "right" }}>{i.reorderLevel}</td>
        <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.qty * i.unitCost)}</td>
        <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
          {i.qty <= i.reorderLevel && <span style={styles.pillAmberSm}>reorder</span>}{" "}
          <button style={styles.iconBtn} onClick={startEdit}>✎</button>{" "}
          <button style={styles.iconBtn} onClick={() => deleteItem(i)}>🗑</button>
        </td>
      </tr>
    );
  }

  const pendingQty = Number(draft.qty) || 0;
  return (
    <tr style={{ background: theme.mode === "dark" ? "rgba(255,255,255,0.03)" : "rgba(0,0,0,0.015)" }}>
      <td style={styles.td}><input style={{ ...styles.inputSmall, width: 80 }} value={draft.sku} onChange={e => setDraftField("sku", e.target.value)} /></td>
      <td style={styles.td}><input style={{ ...styles.inputSmall, width: 140 }} value={draft.name} onChange={e => setDraftField("name", e.target.value)} /></td>
      <td style={styles.td}><select style={styles.inputSmall} value={draft.inventoryType} onChange={e => setDraftField("inventoryType", e.target.value)}>{INVENTORY_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}</select></td>
      <td style={styles.td}><select style={styles.inputSmall} value={draft.locationId} onChange={e => setDraftField("locationId", e.target.value)}><option value="">No location</option>{data.locations.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select></td>
      <td style={styles.td}><select style={styles.inputSmall} value={draft.departmentId} onChange={e => setDraftField("departmentId", e.target.value)}><option value="">No department</option>{data.departments.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}</select></td>
      <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 80, textAlign: "right", color: pendingQty <= Number(draft.reorderLevel) ? theme.rose : theme.text }} value={draft.qty} onChange={e => setDraftField("qty", e.target.value)} /></td>
      <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={draft.unitCost} onChange={e => setDraftField("unitCost", e.target.value)} /></td>
      <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={draft.salePrice} onChange={e => setDraftField("salePrice", e.target.value)} /></td>
      <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 100, textAlign: "right" }} value={draft.nrv} onChange={e => setDraftField("nrv", e.target.value)} /></td>
      <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 80, textAlign: "right" }} value={draft.reorderLevel} onChange={e => setDraftField("reorderLevel", e.target.value)} /></td>
      <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(pendingQty * (Number(draft.unitCost) || 0))}</td>
      <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
        <button style={{ ...styles.iconBtn, color: theme.accent }} onClick={save}>✓ Save</button>{" "}
        <button style={styles.iconBtn} onClick={cancelEdit}>Cancel</button>
      </td>
    </tr>
  );
}

/* =================================== Fixed Asset Register =================================== */
function FixedAssets({ data, setData, postTransaction, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const blank = () => ({ name: "", category: "", purchaseDate: todayStr(), cost: "", salvageValue: "", usefulLifeMonths: 36, taxRatePct: 25, bankId: data.banks[0]?.id, financed: "bank", locationId: "", departmentId: "" });
  const [form, setForm] = useState(blank());

  const addAsset = () => {
    if (!form.name || !form.cost) return notify("Name the asset and enter its cost");
    const bank = data.banks.find(b => b.id === form.bankId);
    const creditAccount = form.financed === "bank" ? bank?.accountId : "2000";
    const ok = postTransaction(`Fixed asset purchase - ${form.name}`, form.purchaseDate, [{ accountId: "1300", debit: Number(form.cost), credit: 0 }, { accountId: creditAccount, debit: 0, credit: Number(form.cost) }], "fixed-asset");
    if (!ok) return;
    setData(d => ({ ...d, fixedAssets: [...d.fixedAssets, { id: uid("fa"), name: form.name, category: form.category, purchaseDate: form.purchaseDate, cost: Number(form.cost), salvageValue: Number(form.salvageValue) || 0, usefulLifeMonths: Number(form.usefulLifeMonths) || 36, accumulatedDepreciation: 0, taxRatePct: Number(form.taxRatePct) || 25, accumulatedTaxDepreciation: 0, lastDepreciationMonth: null, locationId: form.locationId, departmentId: form.departmentId }] }));
    setForm(blank()); notify("Asset added to register");
  };

  const currentMonth = monthKey(todayStr());
  const runDepreciation = () => {
    const due = data.fixedAssets.filter(a => a.lastDepreciationMonth !== currentMonth && a.accumulatedDepreciation < a.cost - a.salvageValue - 0.5);
    if (due.length === 0) return notify("Nothing due for depreciation this month");
    const perAsset = due.map(a => ({ id: a.id, amt: Math.min((a.cost - a.salvageValue) / a.usefulLifeMonths, a.cost - a.salvageValue - a.accumulatedDepreciation) }));
    const total = perAsset.reduce((s, x) => s + x.amt, 0);
    const txn = buildTxn(`Depreciation - ${currentMonth}`, todayStr(), [{ accountId: "5700", debit: Math.round(total), credit: 0 }, { accountId: "1310", debit: 0, credit: Math.round(total) }], "depreciation");
    if (!txn) return notify("Entry not balanced");
    txn.meta = { perAsset }; // lets a deleted depreciation entry roll the register back
    setData(d => ({
      ...d,
      transactions: [...d.transactions, txn],
      fixedAssets: d.fixedAssets.map(a => {
        const m = perAsset.find(x => x.id === a.id);
        if (!m) return a;
        // Tax depreciation runs in parallel at the asset's own capital
        // allowance rate - a tax-only figure (not posted to the GL itself),
        // whose gap against book depreciation is what creates deferred tax.
        const taxCap = a.cost - a.salvageValue - (a.accumulatedTaxDepreciation || 0);
        const taxAmt = Math.max(0, Math.min(((a.cost - a.salvageValue) * (a.taxRatePct || 25) / 100) / 12, taxCap));
        return { ...a, accumulatedDepreciation: a.accumulatedDepreciation + m.amt, accumulatedTaxDepreciation: (a.accumulatedTaxDepreciation || 0) + taxAmt, lastDepreciationMonth: currentMonth };
      }),
    }));
    notify(`Posted ${fmt(total)} depreciation across ${due.length} asset(s)`);
  };

  // Dispose of an asset the proper way: write off cost and accumulated
  // depreciation, post any remaining net book value as a realized loss, then
  // remove it from the register.
  const disposeAsset = async (a) => {
    const nbv = a.cost - a.accumulatedDepreciation;
    if (!(await confirm(`Dispose of ${a.name}? Cost and accumulated depreciation are written off${nbv > 0.5 ? ` and the remaining book value of ${fmt(nbv)} is posted as a loss on disposal` : ""}. This can't be undone.`))) return;
    const lines = [
      { accountId: "1310", debit: a.accumulatedDepreciation, credit: 0 },
      ...(nbv > 0.005 ? [{ accountId: "6000", debit: nbv, credit: 0 }] : []),
      { accountId: "1300", debit: 0, credit: a.cost },
    ];
    const ok = postTransaction(`Disposal - ${a.name}`, todayStr(), lines, "fixed-asset");
    if (!ok) return;
    setData(d => ({ ...d, fixedAssets: d.fixedAssets.filter(x => x.id !== a.id) }));
    notify(`${a.name} disposed - write-off posted to the ledger`);
  };

  // IAS 36: if an asset's recoverable amount (the higher of fair value less
  // costs to sell, and value in use - both estimated by the user) is below
  // its carrying amount, write it down to the recoverable amount. Folded
  // into the same Accumulated Depreciation balance so the net book value
  // formula used everywhere else in the app (cost - accumulatedDepreciation)
  // stays correct without needing a second contra-account; a separate
  // accumulatedImpairment figure is kept purely for disclosure.
  const [impairingId, setImpairingId] = useState(null);
  const [recoverableAmount, setRecoverableAmount] = useState("");
  const impairAsset = async (a) => {
    const nbv = a.cost - a.accumulatedDepreciation;
    const recoverable = Number(recoverableAmount);
    if (!recoverable || recoverable < 0) return notify("Enter the asset's recoverable amount");
    if (recoverable >= nbv) { setImpairingId(null); return notify("Recoverable amount is at or above carrying value - no impairment needed"); }
    const loss = nbv - recoverable;
    if (!(await confirm(`Impair ${a.name}? Carrying value ${fmt(nbv)} written down to its recoverable amount of ${fmt(recoverable)} - a loss of ${fmt(loss)}.`))) return;
    const ok = postTransaction(`Impairment - ${a.name}`, todayStr(), [{ accountId: "5760", debit: loss, credit: 0 }, { accountId: "1310", debit: 0, credit: loss }], "manual");
    if (!ok) return;
    setData(d => ({ ...d, fixedAssets: d.fixedAssets.map(x => x.id === a.id ? { ...x, accumulatedDepreciation: x.accumulatedDepreciation + loss, accumulatedImpairment: (x.accumulatedImpairment || 0) + loss } : x) }));
    setImpairingId(null); setRecoverableAmount("");
    notify(`Impairment loss of ${fmt(loss)} posted`);
  };

  // IAS 16 revaluation model: an increase is recognized in equity (a
  // Revaluation Surplus), not P&L, since it isn't a realized gain - only a
  // decrease is expensed, and only the portion exceeding any surplus already
  // held for this specific asset (a decrease first erodes that asset's own
  // surplus before touching P&L). Uses the "deemed cost" approach: after
  // revaluing, cost resets to the new fair value and accumulated
  // depreciation resets to zero, so future depreciation runs on the new
  // valuation over the asset's remaining useful life.
  const [revaluingId, setRevaluingId] = useState(null);
  const [fairValue, setFairValue] = useState("");
  const revalueAsset = async (a) => {
    const nbv = a.cost - a.accumulatedDepreciation;
    const fv = Number(fairValue);
    if (!fv || fv < 0) return notify("Enter the asset's fair value");
    const diff = fv - nbv;
    if (Math.abs(diff) < 0.5) { setRevaluingId(null); return notify("Fair value matches carrying value - nothing to revalue"); }
    let lines, newSurplus = a.revaluationSurplus || 0;
    if (diff > 0) {
      lines = [{ accountId: "1300", debit: diff, credit: 0 }, { accountId: "3200", debit: 0, credit: diff }];
      newSurplus += diff;
    } else {
      const decrease = -diff;
      const surplusUsed = Math.min(a.revaluationSurplus || 0, decrease);
      const pnlLoss = decrease - surplusUsed;
      lines = [
        ...(surplusUsed > 0.005 ? [{ accountId: "3200", debit: surplusUsed, credit: 0 }] : []),
        ...(pnlLoss > 0.005 ? [{ accountId: "5780", debit: pnlLoss, credit: 0 }] : []),
        { accountId: "1300", debit: 0, credit: decrease },
      ];
      newSurplus -= surplusUsed;
    }
    if (!(await confirm(`Revalue ${a.name} from ${fmt(nbv)} to ${fmt(fv)}? ${diff > 0 ? `The ${fmt(diff)} increase is recognized in equity (Revaluation Surplus).` : `The ${fmt(-diff)} decrease is ${(a.revaluationSurplus || 0) > 0.5 ? "applied against this asset's revaluation surplus first, with any excess expensed to P&L." : "expensed to P&L."}`} Cost and accumulated depreciation reset to reflect the new valuation.`))) return;
    const ok = postTransaction(`Revaluation - ${a.name}`, todayStr(), lines, "manual");
    if (!ok) return;
    setData(d => ({ ...d, fixedAssets: d.fixedAssets.map(x => x.id === a.id ? { ...x, cost: fv, accumulatedDepreciation: 0, valuationModel: "revaluation", revaluationSurplus: newSurplus } : x) }));
    setRevaluingId(null); setFairValue("");
    notify(diff > 0 ? `Revalued up by ${fmt(diff)} - recognized in equity` : `Revalued down by ${fmt(-diff)}`);
  };

  const totalNBV = data.fixedAssets.reduce((s, a) => s + (a.cost - a.accumulatedDepreciation), 0);
  return (
    <div>
      <PageHeader eyebrow="Assets" title="Fixed Asset Register" sub={`${data.fixedAssets.length} assets · net book value ${fmt(totalNBV)}`} action={<button style={styles.btnPrimary} onClick={runDepreciation}>Run depreciation - {currentMonth}</button>} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Add fixed asset</div>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Asset name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
          <input style={styles.input} placeholder="Category" value={form.category} onChange={e => setForm({ ...form, category: e.target.value })} />
          <input type="date" style={styles.input} value={form.purchaseDate} onChange={e => setForm({ ...form, purchaseDate: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 120 }} placeholder="Cost" value={form.cost} onChange={e => setForm({ ...form, cost: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 120 }} placeholder="Salvage value" value={form.salvageValue} onChange={e => setForm({ ...form, salvageValue: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 130 }} placeholder="Useful life (months)" value={form.usefulLifeMonths} onChange={e => setForm({ ...form, usefulLifeMonths: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 150 }} placeholder="Tax rate % (capital allowance)" value={form.taxRatePct} onChange={e => setForm({ ...form, taxRatePct: e.target.value })} title="Annual capital allowance rate used for deferred tax" />
          <select style={styles.input} value={form.financed} onChange={e => setForm({ ...form, financed: e.target.value })}><option value="bank">Paid from bank</option><option value="credit">On credit (AP)</option></select>
          {form.financed === "bank" && <select style={styles.input} value={form.bankId} onChange={e => setForm({ ...form, bankId: e.target.value })}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>}
          <select style={styles.input} value={form.locationId} onChange={e => setForm({ ...form, locationId: e.target.value })}><option value="">No location</option>{data.locations.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select>
          <select style={styles.input} value={form.departmentId} onChange={e => setForm({ ...form, departmentId: e.target.value })}><option value="">No department</option>{data.departments.map(x => <option key={x.id} value={x.id}>{x.name}</option>)}</select>
          <button style={styles.btnPrimary} onClick={addAsset}>Add asset</button>
        </div>
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={{ overflowX: "auto" }}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Asset</th><th style={styles.th}>Category</th><th style={styles.th}>Location</th><th style={styles.th}>Department</th><th style={styles.th}>Purchased</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={{ ...styles.th, textAlign: "right" }}>Acc. depreciation</th><th style={{ ...styles.th, textAlign: "right" }}>Net book value</th><th style={{ ...styles.th, textAlign: "right" }}>Monthly dep.</th><th style={styles.th}></th></tr></thead>
          <tbody>{data.fixedAssets.map(a => (
            <React.Fragment key={a.id}>
              <tr><td style={styles.td}>{a.name}{a.accumulatedImpairment > 0.5 && <span style={{ ...styles.pillAmberSm, marginLeft: 6 }}>impaired</span>}{a.valuationModel === "revaluation" && <span style={{ ...styles.pillAmberSm, marginLeft: 6 }}>revalued</span>}</td><td style={styles.td}>{a.category || "-"}</td><td style={styles.td}>{data.locations.find(l => l.id === a.locationId)?.name || "-"}</td><td style={styles.td}>{data.departments.find(x => x.id === a.departmentId)?.name || "-"}</td><td style={styles.tdMono}>{fmtDate(a.purchaseDate)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(a.cost)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(a.accumulatedDepreciation)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(a.cost - a.accumulatedDepreciation)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", color: theme.muted }}>{fmt((a.cost - a.salvageValue) / a.usefulLifeMonths)}</td>
                <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                  <button style={styles.iconBtn} onClick={() => { setRevaluingId(revaluingId === a.id ? null : a.id); setFairValue(""); }}>Revalue</button>{" "}
                  <button style={styles.iconBtn} onClick={() => { setImpairingId(impairingId === a.id ? null : a.id); setRecoverableAmount(""); }}>Impair</button>{" "}
                  <button style={styles.iconBtn} onClick={() => disposeAsset(a)}>🗑 Dispose</button>
                </td></tr>
              {impairingId === a.id && (
                <tr><td colSpan={10} style={{ ...styles.td, background: theme.panel2 }}>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                    <span style={{ fontSize: 13 }}>Carrying value <strong style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(a.cost - a.accumulatedDepreciation)}</strong> - recoverable amount:</span>
                    <input type="number" style={{ ...styles.inputSmall, width: 140 }} placeholder="Recoverable amount" value={recoverableAmount} onChange={e => setRecoverableAmount(e.target.value)} />
                    <button style={styles.btnPrimary} onClick={() => impairAsset(a)}>Test & post impairment</button>
                  </div>
                </td></tr>
              )}
              {revaluingId === a.id && (
                <tr><td colSpan={10} style={{ ...styles.td, background: theme.panel2 }}>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                    <span style={{ fontSize: 13 }}>Carrying value <strong style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmt(a.cost - a.accumulatedDepreciation)}</strong>{(a.revaluationSurplus || 0) > 0.5 && <span style={{ color: theme.muted }}> (surplus held: {fmt(a.revaluationSurplus)})</span>} - fair value:</span>
                    <input type="number" style={{ ...styles.inputSmall, width: 140 }} placeholder="Fair value" value={fairValue} onChange={e => setFairValue(e.target.value)} />
                    <button style={styles.btnPrimary} onClick={() => revalueAsset(a)}>Post revaluation</button>
                  </div>
                </td></tr>
              )}
            </React.Fragment>
          ))}</tbody>
        </table>
        </div>
      </div>
      <LeaseRegister data={data} setData={setData} notify={notify} />
    </div>
  );
}

// IFRS 16 lease register: at commencement, capitalizes the right-of-use
// asset and lease liability at the present value of fixed payments. Each
// "Run due periods" posts the interest unwind, principal repayment, and the
// straight-line ROU depreciation together in one balanced entry.
function LeaseRegister({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const blank = () => ({ name: "", lessor: "", startDate: todayStr(), termMonths: 36, paymentAmount: "", discountRate: 18, bankId: data.banks[0]?.id });
  const [form, setForm] = useState(blank());
  const [openLease, setOpenLease] = useState(null);
  const currentMonth = monthKey(todayStr());

  const addLease = () => {
    if (!form.name || !form.paymentAmount || !Number(form.termMonths)) return notify("Name the lease, and enter its term and payment amount");
    const rouCost = leasePV(Number(form.paymentAmount), Number(form.termMonths), Number(form.discountRate) || 0);
    const txn = buildTxn(`Lease commencement - ${form.name}`, form.startDate, [{ accountId: "1320", debit: Math.round(rouCost), credit: 0 }, { accountId: "2270", debit: 0, credit: Math.round(rouCost) }], "lease");
    if (!txn) return notify("Entry not balanced");
    const leaseId = uid("lease");
    txn.meta = { leaseId, kind: "commencement" };
    setData(d => ({
      ...d,
      transactions: [...d.transactions, txn],
      leases: [...d.leases, { id: leaseId, name: form.name, lessor: form.lessor, startDate: form.startDate, termMonths: Number(form.termMonths), paymentAmount: Number(form.paymentAmount), discountRate: Number(form.discountRate) || 0, bankId: form.bankId, rouCost: Math.round(rouCost), liabilityBalance: Math.round(rouCost), accumulatedROUDep: 0, periodsRun: 0, lastRunMonth: null, status: "active" }],
    }));
    setForm(blank());
    notify(`Lease capitalized: ROU asset and liability recognized at ${fmt(rouCost)}`);
  };

  const dueLeases = data.leases.filter(l => l.status === "active" && l.lastRunMonth !== currentMonth && l.periodsRun < l.termMonths);
  const runDuePeriods = () => {
    if (dueLeases.length === 0) return notify("No leases due for this month");
    let working = data;
    dueLeases.forEach(lease => {
      const bank = working.banks.find(b => b.id === lease.bankId);
      if (!bank) return;
      const r = (lease.discountRate || 0) / 100 / 12;
      const interest = lease.liabilityBalance * r;
      const principal = lease.paymentAmount - interest;
      const monthlyDep = lease.rouCost / lease.termMonths;
      const txn = buildTxn(`Lease period ${lease.periodsRun + 1}/${lease.termMonths} - ${lease.name}`, todayStr(), [
        { accountId: "5650", debit: Math.round(interest), credit: 0 },
        { accountId: "2270", debit: Math.round(principal), credit: 0 },
        { accountId: bank.accountId, debit: 0, credit: Math.round(interest + principal) },
        { accountId: "5700", debit: Math.round(monthlyDep), credit: 0 },
        { accountId: "1330", debit: 0, credit: Math.round(monthlyDep) },
      ], "lease");
      if (!txn) return;
      txn.meta = { leaseId: lease.id, kind: "period" };
      working = {
        ...working,
        transactions: [...working.transactions, txn],
        leases: working.leases.map(l => l.id === lease.id ? { ...l, liabilityBalance: Math.max(0, l.liabilityBalance - principal), accumulatedROUDep: l.accumulatedROUDep + monthlyDep, periodsRun: l.periodsRun + 1, lastRunMonth: currentMonth } : l),
      };
    });
    setData(working);
    notify(`Posted ${dueLeases.length} lease period(s): interest, principal repayment and ROU depreciation`);
  };

  const terminateLease = async (lease) => {
    const nbv = lease.rouCost - lease.accumulatedROUDep;
    if (!(await confirm(`Terminate "${lease.name}" early? The remaining ROU asset (${fmt(nbv)}) and lease liability (${fmt(lease.liabilityBalance)}) are fully derecognized, with any difference posted as a gain/loss.`))) return;
    const gainLoss = lease.liabilityBalance - nbv; // liability released minus asset written off
    const lines = [
      { accountId: "1330", debit: lease.accumulatedROUDep, credit: 0 },
      { accountId: "2270", debit: lease.liabilityBalance, credit: 0 },
      { accountId: "1320", debit: 0, credit: lease.rouCost },
      ...(Math.abs(gainLoss) > 0.5 ? [{ accountId: "6000", debit: gainLoss < 0 ? -gainLoss : 0, credit: gainLoss > 0 ? gainLoss : 0 }] : []),
    ];
    const txn = buildTxn(`Lease termination - ${lease.name}`, todayStr(), lines, "lease");
    if (!txn) return notify("Entry not balanced");
    txn.meta = { leaseId: lease.id, kind: "termination" };
    setData(d => ({ ...d, transactions: [...d.transactions, txn], leases: d.leases.map(l => l.id === lease.id ? { ...l, status: "terminated" } : l) }));
    notify("Lease terminated - ROU asset and liability derecognized");
  };

  const deleteLease = async (lease) => {
    if (!(await confirm(`Delete "${lease.name}" entirely? All ${lease.periodsRun + 1} of its journal entries (commencement plus every period run) are removed. This can't be undone.`))) return;
    setData(d => ({ ...d, leases: d.leases.filter(l => l.id !== lease.id), transactions: d.transactions.filter(t => !(t.meta && t.meta.leaseId === lease.id)) }));
    notify("Lease and all its journal entries deleted");
  };

  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <div style={styles.cardTitle}>Leases</div>
        {dueLeases.length > 0 && <button style={styles.btnPrimary} onClick={runDuePeriods}>Run {dueLeases.length} due period(s)</button>}
      </div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>At commencement, the lease is capitalized as a right-of-use asset and lease liability at the present value of fixed payments. Each period splits the payment into interest (finance cost) and principal, and depreciates the ROU asset straight-line over the term.</div>

      {data.leases.length > 0 && (
        <table style={{ ...styles.table, marginTop: 10 }}>
          <thead><tr><th style={styles.th}>Lease</th><th style={styles.th}>Term</th><th style={{ ...styles.th, textAlign: "right" }}>ROU asset (NBV)</th><th style={{ ...styles.th, textAlign: "right" }}>Lease liability</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
          <tbody>{data.leases.map(l => {
            const nbv = l.rouCost - l.accumulatedROUDep;
            const split = leaseLiabilitySplit(l);
            return (
              <React.Fragment key={l.id}>
                <tr onClick={() => setOpenLease(openLease === l.id ? null : l.id)} style={{ cursor: "pointer" }}>
                  <td style={{ ...styles.td, fontWeight: 500 }}>{openLease === l.id ? "\u25be" : "\u25b8"} {l.name}</td>
                  <td style={styles.td}>{l.periodsRun}/{l.termMonths} mo</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(nbv)}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(l.liabilityBalance)}</td>
                  <td style={styles.td}><span style={{ ...styles.pill, ...(l.status === "active" ? styles.pillGreen : styles.pillAmber) }}>{l.status}</span></td>
                  <td style={{ ...styles.td, whiteSpace: "nowrap" }} onClick={e => e.stopPropagation()}>
                    {l.status === "active" && <button style={styles.iconBtn} onClick={() => terminateLease(l)}>Terminate</button>}{" "}
                    <button style={styles.iconBtn} onClick={() => deleteLease(l)}>🗑</button>
                  </td>
                </tr>
                {openLease === l.id && (
                  <tr><td colSpan={6} style={{ ...styles.td, background: theme.panel2 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, padding: "6px 4px 12px" }}>
                      <Kpi label="Lessor" value={l.lessor || "-"} />
                      <Kpi label="Monthly payment" value={fmt(l.paymentAmount)} />
                      <Kpi label="Discount rate" value={`${l.discountRate}%`} />
                      <Kpi label="Current portion (12mo)" value={fmt(split.current)} tone="amber" />
                      <Kpi label="Non-current portion" value={fmt(split.nonCurrent)} />
                    </div>
                    <div style={{ fontSize: 11, color: theme.muted, padding: "0 4px 8px" }}>Current/non-current split is shown for balance sheet analysis - this system posts the full liability to a single Lease Liability account rather than actively reclassifying between two accounts each period.</div>
                  </td></tr>
                )}
              </React.Fragment>
            );
          })}</tbody>
        </table>
      )}

      <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
        <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Lease name (e.g. Office Lease)" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
        <input style={styles.input} placeholder="Lessor (optional)" value={form.lessor} onChange={e => setForm({ ...form, lessor: e.target.value })} />
        <input type="date" style={styles.input} value={form.startDate} onChange={e => setForm({ ...form, startDate: e.target.value })} />
        <input type="number" style={{ ...styles.input, width: 110 }} placeholder="Term (months)" value={form.termMonths} onChange={e => setForm({ ...form, termMonths: e.target.value })} />
        <input type="number" style={{ ...styles.input, width: 140 }} placeholder="Monthly payment" value={form.paymentAmount} onChange={e => setForm({ ...form, paymentAmount: e.target.value })} />
        <input type="number" style={{ ...styles.input, width: 140 }} placeholder="Discount rate %" value={form.discountRate} onChange={e => setForm({ ...form, discountRate: e.target.value })} title="Incremental borrowing rate, annual %" />
        <select style={styles.input} value={form.bankId} onChange={e => setForm({ ...form, bankId: e.target.value })}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
        <button style={styles.btnPrimary} onClick={addLease}>Capitalize lease</button>
      </div>
    </div>
  );
}

/* =================================== Bank Feed / Import =================================== */
// Uploads land in a staging feed as UNCATEGORIZED - they have no effect on the
// ledger, financial statements, or bank balances until categorized. User rules
// can auto-categorize on upload; categorized lines can be uncategorized (which
// removes their ledger entry) and only uncategorized lines can be deleted,
// moving them to a bin for reference.
function matchUserRule(rules, desc, amount, bankId) {
  const d = (desc || "").toLowerCase();
  const dir = amount >= 0 ? "in" : "out";
  const hit = (rules || []).find(r => r.bankId === bankId && r.keyword && d.includes(r.keyword.toLowerCase()) && (r.direction === "both" || r.direction === dir));
  return hit || null;
}
function categorizeFeedItem(d, item, accountId) {
  const bank = d.banks.find(b => b.id === item.bankId);
  if (!bank || !accountId) return d;
  const amt = Math.abs(item.amount);
  const lines = item.amount >= 0
    ? [{ accountId: bank.accountId, debit: amt, credit: 0 }, { accountId, debit: 0, credit: amt }]
    : [{ accountId, debit: amt, credit: 0 }, { accountId: bank.accountId, debit: 0, credit: amt }];
  const txn = buildTxn(item.desc, item.date, lines, "bank-import", null);
  if (!txn) return d;
  txn.feedId = item.id;
  return {
    ...d,
    transactions: [...d.transactions, txn],
    bankFeed: d.bankFeed.map(f => f.id === item.id ? { ...f, status: "categorized", txnId: txn.id, accountId } : f),
  };
}
// Parse statement dates robustly. Handles: JS Date objects (from Excel with
// cellDates), Excel serial numbers, ISO yyyy-mm-dd, month names ("04-Mar-2025",
// "Mar 4, 2025"), and numeric d/m/y styles - where the chosen format decides
// how ambiguous values like 03/04/2025 are read. Returns null when unparseable
// (never silently substitutes today's date).
const MONTHS3 = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
function toISO(y, m, d) {
  y = Number(y); m = Number(m); d = Number(d);
  if (y < 100) y += y > 50 ? 1900 : 2000;
  if (!y || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null; // e.g. 31/02
  return dt.toISOString().slice(0, 10);
}
function parseStatementDate(raw, fmt) {
  if (raw === null || raw === undefined || raw === "") return null;
  if (raw instanceof Date && !isNaN(raw.getTime())) return excelDateToISO(raw);
  // Excel serial number (days since 1899-12-30)
  const asNum = typeof raw === "number" ? raw : (/^\d{4,5}(\.\d+)?$/.test(String(raw).trim()) ? Number(raw) : NaN);
  if (!isNaN(asNum) && asNum > 20000 && asNum < 80000) {
    const dt = new Date(Math.round((asNum - 25569) * 86400 * 1000));
    return isNaN(dt.getTime()) ? null : dt.toISOString().slice(0, 10);
  }
  const s = String(raw).trim();
  let m;
  // ISO: 2025-03-04 (also 2025/03/04, 2025.03.04)
  if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/))) return toISO(m[1], m[2], m[3]);
  // Month name styles: 04-Mar-2025 / 4 March 2025 / Mar 4, 2025
  if ((m = s.match(/^(\d{1,2})[\s\-/]([A-Za-z]{3,})[\s\-/,]+(\d{2,4})$/))) {
    const mo = MONTHS3[m[2].slice(0, 3).toLowerCase()]; return mo ? toISO(m[3], mo, m[1]) : null;
  }
  if ((m = s.match(/^([A-Za-z]{3,})[\s\-/]+(\d{1,2})[\s,\-/]+(\d{2,4})$/))) {
    const mo = MONTHS3[m[1].slice(0, 3).toLowerCase()]; return mo ? toISO(m[3], mo, m[2]) : null;
  }
  // Numeric d/m/y or m/d/y: 03/04/2025, 3-4-25, 03.04.2025
  if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/))) {
    const [_, a, b, y] = m;
    if (fmt === "MDY") return toISO(y, a, b);
    if (fmt === "DMY") return toISO(y, b, a) ? toISO(y, b, a) : null; // day=a month=b
    // auto: unambiguous when one part exceeds 12; ambiguous defaults to DD/MM
    if (Number(a) > 12) return toISO(y, b, a);
    if (Number(b) > 12) return toISO(y, a, b);
    return toISO(y, b, a);
  }
  return null;
}
function BankFeedPanel({ bank, data, setData, notify }) {
  const { styles, fmt, theme } = useUI();
  const [fileName, setFileName] = useState("");
  const bankId = bank.id;
  const [autoApply, setAutoApply] = useState(true);
  const [dateFormat, setDateFormat] = useState("auto");
  const [preview, setPreview] = useState(null); // { fileName, rows: [{rawDate, desc, amount}] }
  const [ruleForm, setRuleForm] = useState({ keyword: "", accountId: "5200", direction: "out" });
  const [feedTab, setFeedTab] = useState("uncategorized");
  const [pendingCategory, setPendingCategory] = useState({});
  const catAccounts = activeAccounts(data).filter(a => a.type === "expense" || a.type === "revenue");

  const handleFile = async (file) => {
    setFileName(file.name);
    const ext = file.name.split(".").pop().toLowerCase();
    let parsed = [];
    if (ext === "csv") { const text = await file.text(); parsed = Papa.parse(text, { header: true, skipEmptyLines: true }).data; }
    else if (ext === "xlsx" || ext === "xls") { const buf = await file.arrayBuffer(); const wb = XLSX.read(buf, { type: "array", cellDates: true }); parsed = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }); }
    else if (ext === "pdf") { notify("PDF statements need OCR/text extraction on a server - this browser demo supports CSV and Excel exports from your bank instead."); return; }
    else { notify("Unsupported file type. Use CSV or Excel."); return; }

    const rows = parsed.map((r) => {
      const keys = Object.keys(r);
      const dateKey = keys.find(k => /date/i.test(k));
      const descKey = keys.find(k => /desc|memo|narrat|detail/i.test(k)) || keys.find(k => typeof r[k] === "string");
      const amountKey = keys.find(k => /amount|value/i.test(k));
      const debitKey = keys.find(k => /debit|withdraw/i.test(k));
      const creditKey = keys.find(k => /credit|deposit/i.test(k));
      let amount = 0;
      if (amountKey) amount = Number(String(r[amountKey]).replace(/[^0-9.-]/g, "")) || 0;
      else if (debitKey || creditKey) amount = (Number(String(r[creditKey] || 0).replace(/[^0-9.-]/g, "")) || 0) - (Number(String(r[debitKey] || 0).replace(/[^0-9.-]/g, "")) || 0);
      const desc = descKey ? String(r[descKey]) : "Imported transaction";
      return { rawDate: dateKey ? r[dateKey] : "", desc, amount };
    }).filter(r => r.amount !== 0);

    if (rows.length === 0) return notify("No usable rows found in that file");
    setPreview({ fileName: file.name, rows });
  };

  // Live re-parse against the chosen format so the preview always shows
  // exactly the dates that will be staged.
  const previewParsed = preview ? preview.rows.map(r => ({ ...r, date: parseStatementDate(r.rawDate, dateFormat) })) : [];
  const invalidCount = previewParsed.filter(r => !r.date).length;

  const confirmStage = () => {
    const good = previewParsed.filter(r => r.date);
    if (good.length === 0) return notify("No rows have a valid date - try a different date format");
    const items = good.map(r => ({ id: uid("feed"), bankId, date: r.date, desc: r.desc, amount: r.amount, status: "uncategorized", importedAt: todayStr(), fileName: preview.fileName }));
    setData(d => {
      let next = { ...d, bankFeed: [...d.bankFeed, ...items] };
      let autoCount = 0;
      if (autoApply) {
        items.forEach(item => {
          const rule = matchUserRule(next.categoryRules, item.desc, item.amount, item.bankId);
          if (rule) { next = categorizeFeedItem(next, item, rule.accountId); autoCount++; }
        });
      }
      setTimeout(() => notify(`${items.length} lines staged as uncategorized${invalidCount ? ` · ${invalidCount} skipped (bad date)` : ""}${autoCount ? ` · ${autoCount} auto-categorized by your rules` : ""}`), 0);
      return next;
    });
    setPreview(null); setFileName("");
  };

  const categorize = (item) => {
    const accountId = pendingCategory[item.id] || matchUserRule(data.categoryRules, item.desc, item.amount, item.bankId)?.accountId || (item.amount < 0 ? guessCategory(item.desc) : "4100");
    if (!accountId) return notify("Choose a category first");
    setData(d => categorizeFeedItem(d, item, accountId));
    notify("Categorized and posted to the ledger");
  };

  const uncategorize = (item) => {
    const txn = data.transactions.find(t => t.id === item.txnId);
    if (txn?.matchedDocId) return notify(`This line is matched to ${txn.matchedDocId} - its effect on that document must be reversed before uncategorizing`);
    setData(d => ({
      ...d,
      transactions: d.transactions.filter(t => t.id !== item.txnId),
      bankFeed: d.bankFeed.map(f => f.id === item.id ? { ...f, status: "uncategorized", txnId: null, accountId: null } : f),
    }));
    notify("Uncategorized - the ledger entry was removed");
  };

  const deleteItem = (item) => {
    if (item.status !== "uncategorized") return notify("Uncategorize this line first, then delete it");
    setData(d => ({
      ...d,
      bankFeed: d.bankFeed.filter(f => f.id !== item.id),
      bin: [...d.bin, { ...item, deletedAt: todayStr() }],
    }));
    notify("Moved to bin");
  };

  const restoreItem = (item) => {
    setData(d => ({ ...d, bin: d.bin.filter(f => f.id !== item.id), bankFeed: [...d.bankFeed, { ...item, deletedAt: undefined }] }));
    notify("Restored to uncategorized");
  };

  const addRule = () => {
    if (!ruleForm.keyword) return notify("Enter a keyword for the rule");
    setData(d => ({ ...d, categoryRules: [...d.categoryRules, { ...ruleForm, id: uid("rule"), bankId }] }));
    setRuleForm({ keyword: "", accountId: "5200", direction: "out" });
    notify("Rule added");
  };
  const removeRule = (id) => setData(d => ({ ...d, categoryRules: d.categoryRules.filter(r => r.id !== id) }));

  const runRulesNow = () => {
    setData(d => {
      let next = d; let count = 0;
      d.bankFeed.filter(f => f.status === "uncategorized" && f.bankId === bankId).forEach(item => {
        const rule = matchUserRule(next.categoryRules, item.desc, item.amount, item.bankId);
        if (rule) { next = categorizeFeedItem(next, item, rule.accountId); count++; }
      });
      setTimeout(() => notify(count ? `${count} line(s) auto-categorized` : "No uncategorized lines matched your rules"), 0);
      return next;
    });
  };

  const uncategorizedItems = data.bankFeed.filter(f => f.status === "uncategorized" && f.bankId === bankId).sort((a, b) => b.date.localeCompare(a.date));
  const categorizedItems = data.bankFeed.filter(f => f.status === "categorized" && f.bankId === bankId).sort((a, b) => b.date.localeCompare(a.date));
  const bankRules = data.categoryRules.filter(r => r.bankId === bankId);
  const bankBin = data.bin.filter(f => f.bankId === bankId);

  return (
    <div>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Upload a statement into {bank.name}</div>
        <div style={{ fontSize: 13, color: theme.muted, marginTop: 4 }}>Uploaded lines are staged as <strong>uncategorized</strong> - they have no effect on your financial statements or this bank's balance until you categorize them (or a rule does).</div>
        <div style={{ display: "flex", gap: 14, marginTop: 10, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13 }}>
            <span style={{ color: theme.muted }}>Date format in file:</span>
            <select style={styles.input} value={dateFormat} onChange={e => setDateFormat(e.target.value)}>
              <option value="auto">Auto-detect (ambiguous → DD/MM/YYYY)</option>
              <option value="DMY">DD/MM/YYYY (e.g. 31/05/2026)</option>
              <option value="MDY">MM/DD/YYYY (e.g. 05/31/2026)</option>
            </select>
          </label>
          <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13 }}>
            <input type="checkbox" checked={autoApply} onChange={e => setAutoApply(e.target.checked)} /> Auto-categorize on upload when a rule matches
          </label>
        </div>
        <div style={{ fontSize: 11.5, color: theme.muted, marginTop: 6 }}>Also understood automatically: 2026-05-31 (ISO), 31-May-2026, May 31 2026, and Excel date cells. Rows whose date can't be read are flagged in the preview and skipped - never silently given today's date.</div>
        <div style={styles.dropzone} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}>
          <input id="file-input" type="file" accept=".csv,.xlsx,.xls,.pdf" style={{ display: "none" }} onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />
          <label htmlFor="file-input" style={{ cursor: "pointer" }}><div style={{ fontSize: 28 }}>⇩</div><div style={{ marginTop: 6 }}>{fileName || "Drop a CSV or Excel file, or click to browse"}</div></label>
        </div>
      </div>

      {preview && (
        <div className="ces-card" style={styles.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <div style={styles.cardTitle}>Check the dates before staging - {preview.fileName}</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button style={styles.btnGhost} onClick={() => { setPreview(null); setFileName(""); }}>Cancel</button>
              <button style={styles.btnPrimary} onClick={confirmStage}>Stage {previewParsed.length - invalidCount} line{previewParsed.length - invalidCount === 1 ? "" : "s"}</button>
            </div>
          </div>
          <div style={{ fontSize: 12.5, color: invalidCount ? theme.rose : theme.muted, marginTop: 4 }}>
            {invalidCount ? `${invalidCount} row(s) have unreadable dates and will be skipped - try switching the date format above; the preview updates instantly.` : "All dates parsed. If they look wrong (day and month swapped), switch the date format above - the preview updates instantly."}
          </div>
          <div style={{ maxHeight: 320, overflowY: "auto", marginTop: 10 }}>
            <table style={styles.table}>
              <thead><tr><th style={styles.th}>In your file</th><th style={styles.th}>Will be recorded as</th><th style={styles.th}>Description</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
              <tbody>{previewParsed.map((r, i) => (
                <tr key={i}>
                  <td style={styles.tdMono}>{r.rawDate instanceof Date ? excelDateToISO(r.rawDate) : String(r.rawDate) || "-"}</td>
                  <td style={{ ...styles.tdMono, color: r.date ? theme.emerald : theme.rose, fontWeight: 600 }}>{r.date ? fmtDate(r.date) : "✕ unreadable"}</td>
                  <td style={styles.td}>{r.desc}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right", color: r.amount < 0 ? theme.rose : theme.emerald }}>{fmt(r.amount)}</td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        </div>
      )}

      <div className="ces-card" style={styles.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <div style={styles.cardTitle}>Categorization rules - {bank.name}</div>
          <button style={styles.btnGhost} onClick={runRulesNow}>▶ Run rules on uncategorized now</button>
        </div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Rules belong to this bank only - other banks keep their own rules. If a line's description contains the keyword (and the money direction matches), the rule categorizes it automatically.</div>
        {bankRules.length > 0 && (
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}>Keyword</th><th style={styles.th}>Direction</th><th style={styles.th}>Categorize as</th><th></th></tr></thead>
            <tbody>{bankRules.map(r => (
              <tr key={r.id}><td style={styles.tdMono}>"{r.keyword}"</td><td style={styles.td}>{r.direction === "in" ? "Money in" : r.direction === "out" ? "Money out" : "Both"}</td>
                <td style={styles.td}>{data.accounts.find(a => a.id === r.accountId)?.name}</td>
                <td style={styles.td}><button style={styles.iconBtn} onClick={() => removeRule(r.id)}>✕</button></td></tr>
            ))}</tbody>
          </table>
        )}
        <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 140 }} placeholder="Keyword (e.g. rent, MTN, salary)" value={ruleForm.keyword} onChange={e => setRuleForm({ ...ruleForm, keyword: e.target.value })} />
          <select style={styles.input} value={ruleForm.direction} onChange={e => setRuleForm({ ...ruleForm, direction: e.target.value })}><option value="out">Money out</option><option value="in">Money in</option><option value="both">Both</option></select>
          <select style={styles.input} value={ruleForm.accountId} onChange={e => setRuleForm({ ...ruleForm, accountId: e.target.value })}>{catAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
          <button style={styles.btnPrimary} onClick={addRule}>Add rule</button>
        </div>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {[["uncategorized", `Uncategorized (${uncategorizedItems.length})`], ["categorized", `Categorized (${categorizedItems.length})`], ["bin", `Bin (${bankBin.length})`]].map(([id, label]) => (
          <button key={id} style={{ ...styles.btnGhost, ...(feedTab === id ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }} onClick={() => setFeedTab(id)}>{label}</button>
        ))}
      </div>

      {feedTab === "uncategorized" && (
        <div className="ces-card" style={styles.card}>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Description</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}>Category</th><th style={styles.th}>Actions</th></tr></thead>
            <tbody>{uncategorizedItems.map(item => {
              const rule = matchUserRule(data.categoryRules, item.desc, item.amount, item.bankId);
              const suggested = pendingCategory[item.id] || rule?.accountId || (item.amount < 0 ? guessCategory(item.desc) : "4100") || "";
              return (
                <tr key={item.id}>
                  <td style={styles.tdMono}>{fmtDate(item.date)}</td>
                  <td style={styles.td}>{item.desc}{rule && <span style={{ ...styles.pillAmberSm, marginLeft: 8 }}>rule: "{rule.keyword}"</span>}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right", color: item.amount < 0 ? theme.rose : theme.emerald }}>{fmt(item.amount)}</td>
                  <td style={styles.td}><select style={styles.inputSmall} value={suggested} onChange={e => setPendingCategory(pc => ({ ...pc, [item.id]: e.target.value }))}>
                    <option value="">Select category</option>{catAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select></td>
                  <td style={styles.td}>
                    <button style={styles.btnGhost} onClick={() => categorize(item)}>Categorize</button>{" "}
                    <button style={styles.iconBtn} onClick={() => deleteItem(item)}>🗑 Delete</button>
                  </td>
                </tr>
              );
            })}</tbody>
          </table>
          {uncategorizedItems.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>Nothing awaiting categorization. Upload a statement above.</div>}
        </div>
      )}

      {feedTab === "categorized" && (
        <div className="ces-card" style={styles.card}>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Description</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}>Categorized as</th><th style={styles.th}>Actions</th></tr></thead>
            <tbody>{categorizedItems.map(item => (
              <tr key={item.id}>
                <td style={styles.tdMono}>{fmtDate(item.date)}</td>
                <td style={styles.td}>{item.desc}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", color: item.amount < 0 ? theme.rose : theme.emerald }}>{fmt(item.amount)}</td>
                <td style={styles.td}><span style={{ ...styles.pill, ...styles.pillGreen }}>{data.accounts.find(a => a.id === item.accountId)?.name || "posted"}</span></td>
                <td style={styles.td}><button style={styles.btnGhost} onClick={() => uncategorize(item)}>↩ Uncategorize</button></td>
              </tr>
            ))}</tbody>
          </table>
          {categorizedItems.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No categorized feed lines yet.</div>}
          <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>Uncategorizing removes the line's ledger entry and returns it to the uncategorized list - it must be uncategorized before it can be deleted. Lines already matched to an invoice or bill can't be uncategorized until that match is reversed.</div>
        </div>
      )}

      {feedTab === "bin" && (
        <div className="ces-card" style={styles.card}>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Deleted</th><th style={styles.th}>Original date</th><th style={styles.th}>Description</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}></th></tr></thead>
            <tbody>{[...bankBin].sort((a, b) => (b.deletedAt || "").localeCompare(a.deletedAt || "")).map(item => (
              <tr key={item.id}>
                <td style={styles.tdMono}>{fmtDate(item.deletedAt)}</td>
                <td style={styles.tdMono}>{fmtDate(item.date)}</td>
                <td style={styles.td}>{item.desc}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(item.amount)}</td>
                <td style={styles.td}><button style={styles.btnGhost} onClick={() => restoreItem(item)}>↩ Restore</button></td>
              </tr>
            ))}</tbody>
          </table>
          {bankBin.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>The bin is empty. Deleted uncategorized lines are kept here for reference.</div>}
        </div>
      )}
    </div>
  );
}

/* =================================== Reports =================================== */
// Rolling six calendar months (oldest first): label plus from/to date bounds,
// with the current month running only up to today.
function last6Months() {
  const out = [];
  const now = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const from = localDateToISO(d);
    const endOfMonth = localDateToISO(new Date(d.getFullYear(), d.getMonth() + 1, 0));
    const to = i === 0 ? todayStr() : endOfMonth;
    out.push({ label: d.toLocaleDateString(undefined, { month: "short" }), from, to });
  }
  return out;
}
function rangeToDates(range) {
  // Custom range: { from, to } picked by the user.
  if (range && typeof range === "object") return { from: range.from || "2000-01-01", to: range.to || todayStr() };
  const today = new Date();
  if (range === "month") // 1st of this month → today
    return { from: localDateToISO(new Date(today.getFullYear(), today.getMonth(), 1)), to: todayStr() };
  if (range === "lastmonth") { // 1st → last day of the previous month
    const from = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const to = new Date(today.getFullYear(), today.getMonth(), 0);
    return { from: localDateToISO(from), to: localDateToISO(to) };
  }
  if (range === "quarter") // last 3 months → today
    return { from: localDateToISO(new Date(today.getFullYear(), today.getMonth() - 2, 1)), to: todayStr() };
  if (range === "year") // full calendar year: Jan 1 → Dec 31
    return { from: `${today.getFullYear()}-01-01`, to: `${today.getFullYear()}-12-31` };
  return { from: "2000-01-01", to: todayStr() }; // all time
}
// IAS 1 requires at least one comparative period on every primary statement.
// Given a range, returns the immediately preceding period of equal length
// (or null when there's no meaningful "prior" - e.g. All time).
// Returns the period N steps before the given range, using the same
// granularity as that range (N months back for "month", N years back for
// "year", N quarter-length windows for "quarter", etc.). N=1 is the
// immediately preceding period; up to 20 supported.
function priorRangeN(range, n) {
  if (!n || n < 1) return null;
  if (range && typeof range === "object") {
    const days = Math.round((new Date(range.to) - new Date(range.from)) / 86400000) + 1;
    const to = new Date(range.from); to.setDate(to.getDate() - 1 - (n - 1) * days);
    const from = new Date(to); from.setDate(from.getDate() - days + 1);
    return { from: localDateToISO(from), to: localDateToISO(to) };
  }
  const d = new Date();
  if (range === "year") { const y = d.getFullYear() - n; return { from: `${y}-01-01`, to: `${y}-12-31` }; }
  if (range === "month") return { from: localDateToISO(new Date(d.getFullYear(), d.getMonth() - n, 1)), to: localDateToISO(new Date(d.getFullYear(), d.getMonth() - n + 1, 0)) };
  if (range === "lastmonth") return { from: localDateToISO(new Date(d.getFullYear(), d.getMonth() - 1 - n, 1)), to: localDateToISO(new Date(d.getFullYear(), d.getMonth() - n, 0)) };
  if (range === "quarter") return { from: localDateToISO(new Date(d.getFullYear(), d.getMonth() - 2 - n * 3, 1)), to: localDateToISO(new Date(d.getFullYear(), d.getMonth() + 1 - n * 3, 0)) };
  return null; // "all" has no meaningful prior
}
// Same date, N years earlier - the standard comparative for a point-in-time
// statement like the Balance Sheet.
function priorAsOfN(asOf, n) { const d = new Date(asOf); d.setFullYear(d.getFullYear() - n); return localDateToISO(d); }
// Human period name for a range: "July 2026" for a full month, "Jan 2026 -
// Dec 2026" for a full year or any multi-month span, or exact dates for a
// short custom window - no jargon, just the calendar period being shown.
function periodLabel(range) {
  const { from, to } = rangeToDates(range);
  const fromD = new Date(from), toD = new Date(to);
  const sameMonth = fromD.getFullYear() === toD.getFullYear() && fromD.getMonth() === toD.getMonth();
  if (sameMonth) return fromD.toLocaleDateString(undefined, { month: "long", year: "numeric" });
  return `${fromD.toLocaleDateString(undefined, { month: "short", year: "numeric" })} - ${toD.toLocaleDateString(undefined, { month: "short", year: "numeric" })}`;
}
// Renders a line with current-period, prior-period, and variance columns.
// Falls back to a single-column RowLine when no prior value is supplied.
// Builds N consecutive periods of the same granularity as `range`, oldest
// first, ending with the current period - e.g. range="month", n=4 while
// viewing July gives [April, May, June, July]. This is a rolling trend
// window, not a single "vs one period back" comparison.
function computeTrendRanges(range, n) {
  const out = [];
  for (let i = Math.max(1, n) - 1; i >= 0; i--) out.push(i === 0 ? range : priorRangeN(range, i));
  return out;
}
// N consecutive month-end dates ending at asOf, for point-in-time statements
// (Balance Sheet) that don't have a "range" to step - e.g. asOf=Jul 16 2026,
// n=4 gives month-end dates for Apr, May, Jun, and asOf itself.
function computeTrendDates(asOf, n) {
  const out = []; const d = new Date(asOf);
  for (let i = Math.max(1, n) - 1; i >= 0; i--) {
    if (i === 0) { out.push(asOf); continue; }
    out.push(localDateToISO(new Date(d.getFullYear(), d.getMonth() - i + 1, 0)));
  }
  return out;
}
// Average period-over-period percentage change across a series of values -
// the arithmetic mean of each consecutive step's growth rate.
function avgChangePct(values) {
  const steps = [];
  for (let i = 1; i < values.length; i++) { const a = values[i - 1], b = values[i]; if (a) steps.push((b - a) / Math.abs(a)); }
  if (steps.length === 0) return null;
  return (steps.reduce((s, v) => s + v, 0) / steps.length) * 100;
}
// One line item across every period in the trend, with a trailing average
// period-over-period change column.
function TrendRow({ label, values, periodCount, bold, indent, divider, tone, onClick }) {
  const { theme, fmt } = useUI();
  const n = periodCount || (values ? values.length : 1);
  const avg = values && values.length > 1 ? avgChangePct(values) : null;
  const c = tone === "rose" ? theme.rose : tone === "emerald" ? theme.emerald : theme.text;
  return (
    <div onClick={onClick} className={onClick ? "drill-row" : undefined} style={{ display: "grid", gridTemplateColumns: `1fr repeat(${n}, 96px) 84px`, gap: 6, alignItems: "baseline", padding: "4px 0", borderTop: divider ? `1px solid ${theme.border}` : "none", marginTop: divider ? 4 : 0, cursor: onClick ? "pointer" : "default" }}>
      <span style={{ paddingLeft: indent ? 14 : 0, fontWeight: bold ? 600 : 400, fontSize: 12.5 }}>{label}</span>
      {Array.from({ length: n }, (_, i) => <span key={i} style={{ textAlign: "right", fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 12.5, fontWeight: bold ? 600 : 400, color: c }}>{values && values[i] !== undefined ? fmt(values[i]) : ""}</span>)}
      <span style={{ textAlign: "right", fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, color: avg === null ? theme.muted : avg >= 0 ? theme.emerald : theme.rose }}>{avg === null ? "" : `${avg >= 0 ? "+" : ""}${avg.toFixed(1)}%`}</span>
    </div>
  );
}
function TrendHeader({ labels }) {
  const { theme } = useUI();
  return (
    <div style={{ display: "grid", gridTemplateColumns: `1fr repeat(${labels.length}, 96px) 84px`, gap: 6, marginBottom: 4 }}>
      <span></span>
      {labels.map((l, i) => <span key={i} style={{ textAlign: "right", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em", color: theme.muted }}>{l}</span>)}
      <span style={{ textAlign: "right", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em", color: theme.muted }}>Avg chg</span>
    </div>
  );
}
// Table/Chart toggle for any trend-capable report, plus the line chart itself.
function TrendViewSwitch({ view, setView }) {
  const { styles, theme } = useUI();
  return (
    <div className="no-print" style={{ display: "flex", gap: 6, marginBottom: 12 }}>
      {["table", "chart"].map(v => (
        <button key={v} style={{ ...styles.btnGhost, ...(view === v ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }} onClick={() => setView(v)}>{v === "table" ? "Table" : "Chart"}</button>
      ))}
    </div>
  );
}
function TrendChart({ labels, series, colors }) {
  const { theme, fmt } = useUI();
  const chartData = labels.map((label, i) => { const row = { label }; series.forEach(s => { row[s.key] = s.values[i]; }); return row; });
  return (
    <ResponsiveContainer width="100%" height={300}>
      <LineChart data={chartData} margin={{ left: 0, right: 12, top: 10 }}>
        <CartesianGrid strokeDasharray="2 4" stroke={theme.border} vertical={false} />
        <XAxis dataKey="label" tick={{ fontFamily: "Inter, sans-serif", fontSize: 11, fill: theme.muted }} axisLine={{ stroke: theme.border }} tickLine={false} />
        <YAxis tick={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, fill: theme.muted }} axisLine={false} tickLine={false} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
        <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontFamily: "Inter, sans-serif", fontSize: 12, background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 6, color: theme.text }} />
        <Legend wrapperStyle={{ fontSize: 12 }} />
        {series.map((s, i) => <Line key={s.key} type="monotone" dataKey={s.key} name={s.label} stroke={colors[i % colors.length]} strokeWidth={2} dot={{ r: 3 }} />)}
      </LineChart>
    </ResponsiveContainer>
  );
}
const RANGE_OPTIONS = [
  ["month", "This month"], ["lastmonth", "Last month"], ["quarter", "Last 3 months"],
  ["year", "This year"], ["all", "All time"], ["custom", "Custom range…"],
];
function RangePicker({ range, setRange, customFrom, setCustomFrom, customTo, setCustomTo }) {
  const { styles } = useUI();
  return (
    <div className="no-print" style={{ marginBottom: 16, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
      <select style={styles.input} value={range} onChange={e => setRange(e.target.value)}>
        {RANGE_OPTIONS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
      </select>
      {range === "custom" && <>
        <input type="date" style={styles.input} value={customFrom} onChange={e => setCustomFrom(e.target.value)} />
        <span style={{ fontSize: 12 }}>to</span>
        <input type="date" style={styles.input} value={customTo} onChange={e => setCustomTo(e.target.value)} />
      </>}
    </div>
  );
}

const REPORT_CATALOG = [
  { id: "pl", period: "range", name: "Profit & Loss", cat: "Financial Statements", desc: "Revenue, expenses and net income for a period" },
  { id: "bs", period: "asof", name: "Balance Sheet", cat: "Financial Statements", desc: "Assets, liabilities and equity as of a date" },
  { id: "cf", period: "range", name: "Cash Flow", cat: "Financial Statements", desc: "Cash in and out across all connected banks" },
  { id: "equity-movement", period: "range", name: "Movement in Equity", cat: "Financial Statements", desc: "Opening equity, contributions, net income, closing equity" },
  { id: "trial-balance", period: "asof", name: "Trial Balance", cat: "Financial Statements", desc: "Every account's debit/credit balance as of a date" },
  { id: "general-ledger", period: "range", name: "General Ledger", cat: "Financial Statements", desc: "Every posted transaction, grouped by account" },
  { id: "account-transactions", period: "range", name: "Account Transactions", cat: "Financial Statements", desc: "Drill into one account's activity with a running balance" },
  { id: "account-type-summary", period: "today", name: "Account Type Summary", cat: "Financial Statements", desc: "Totals grouped by account type and subtype" },
  { id: "ar-aging", period: "today", name: "Receivables Aging", cat: "Receivables & Payables", desc: "Outstanding customer invoices by age bucket" },
  { id: "ap-aging", period: "today", name: "Payables Aging", cat: "Receivables & Payables", desc: "Outstanding vendor bills by age bucket" },
  { id: "sales-by-item", period: "alltime", name: "Sales by Item", cat: "Sales", desc: "Revenue and quantity sold per item or description" },
  { id: "sales-by-customer", period: "alltime", name: "Sales by Customer", cat: "Sales", desc: "Revenue and invoice count per customer" },
  { id: "sales-by-salesperson", period: "alltime", name: "Sales by Salesperson", cat: "Sales", desc: "Revenue and invoice count per salesperson" },
  { id: "profit-by-item", period: "alltime", name: "Profit by Item", cat: "Sales", desc: "Revenue, cost and margin per item sold" },
  { id: "order-fulfillment-item", period: "alltime", name: "Order Fulfillment by Item", cat: "Sales", desc: "Ordered vs fulfilled quantity per item across sales orders" },
  { id: "sales-return-history", period: "alltime", name: "Sales Return History", cat: "Sales", desc: "Every sales return, with customer and value" },
  { id: "refund-history", period: "alltime", name: "Refund History", cat: "Sales", desc: "Refunds issued to customers or received from vendors" },
  { id: "expenses-by-category", period: "range", name: "Expenses by Category", cat: "Purchases", desc: "Spend across every expense account for a period" },
  { id: "expenses-by-project", period: "range", name: "Expenses by Project", cat: "Purchases", desc: "Spend grouped by project for a period" },
  { id: "purchase-by-vendor", period: "alltime", name: "Purchase by Vendor", cat: "Purchases", desc: "Bill count and total spend per vendor" },
  { id: "purchase-by-item", period: "alltime", name: "Purchase by Item", cat: "Purchases", desc: "Quantity purchased and spend per item" },
  { id: "goods-received-history", period: "alltime", name: "Goods Received History", cat: "Purchases", desc: "Every purchase receipt, billed or awaiting a bill" },
  { id: "goods-received-by-item", period: "alltime", name: "Goods Received by Item", cat: "Purchases", desc: "Quantity and value received per item" },
  { id: "customer-balance-summary", period: "today", name: "Customer Balance Summary", cat: "Receivables & Payables", desc: "Invoiced, paid and outstanding per customer" },
  { id: "receivables-details", period: "today", name: "Receivables Details", cat: "Receivables & Payables", desc: "Every open invoice, with days overdue" },
  { id: "vendor-balance-summary", period: "today", name: "Vendor Balance Summary", cat: "Receivables & Payables", desc: "Billed, paid and outstanding per vendor" },
  { id: "payments-received", period: "range", name: "Payments Received", cat: "Receivables & Payables", desc: "Every payment received in a period" },
  { id: "time-to-get-paid", period: "alltime", name: "Time to Get Paid", cat: "Receivables & Payables", desc: "Average days between invoice date and payment" },
  { id: "credit-note-details", period: "alltime", name: "Credit Note Details", cat: "Receivables & Payables", desc: "Every credit note issued, with reason and application" },
  { id: "recurring-transactions", period: "alltime", name: "Recurring Transactions", cat: "Receivables & Payables", desc: "Every recurring schedule - journals, expenses and bills" },
  { id: "committed-stock", period: "today", name: "Committed Stock", cat: "Inventory", desc: "On-hand, committed and available quantity per item" },
  { id: "inventory-movement-history", period: "range", name: "Inventory Movement History", cat: "Inventory", desc: "Stock received, by lot, over a period" },
  { id: "inventory-turnover-item", period: "range", name: "Inventory Turnover by Item", cat: "Inventory", desc: "How many times each item's stock turned over in a period" },
  { id: "inventory-turnover-amount", period: "range", name: "Inventory Turnover by Amount", cat: "Inventory", desc: "Whole-inventory turnover ratio and days on hand" },
  { id: "reconciliation-status", period: "alltime", name: "Reconciliation Status", cat: "Tax", desc: "Every bank reconciliation and whether it's completed" },
  { id: "unrealized-gain-loss", period: "today", name: "Unrealized Gain/Loss", cat: "Tax", desc: "Preview FX gain/loss on foreign-currency accounts, not yet posted" },
  { id: "timesheet-details", period: "range", name: "Timesheet Details", cat: "Projects", desc: "Every logged time entry, billable status and invoicing" },
  { id: "project-details", period: "alltime", name: "Project Details", cat: "Projects", desc: "Every invoice, bill, expense and time entry for one project" },
  { id: "project-cost-summary", period: "alltime", name: "Project Cost Summary", cat: "Projects", desc: "Cost broken down by bills, expenses and billable time" },
  { id: "journal-report", period: "range", name: "Journal Report", cat: "Financial Statements", desc: "Every entry in chronological order, full detail" },
  { id: "inventory-summary", period: "today", name: "Inventory Summary", cat: "Inventory", desc: "On-hand quantity, cost and value per item" },
  { id: "fixed-assets-by-location", period: "today", name: "Fixed Assets by Location", cat: "Inventory", desc: "Asset count, cost and net book value per location" },
  { id: "fixed-assets-by-department", period: "today", name: "Fixed Assets by Department", cat: "Inventory", desc: "Asset count, cost and net book value per department" },
  { id: "inventory-by-location", period: "today", name: "Inventory by Location", cat: "Inventory", desc: "Item count, quantity and value per location" },
  { id: "inventory-by-department", period: "today", name: "Inventory by Department", cat: "Inventory", desc: "Item count, quantity and value per department" },
  { id: "inventory-aging", period: "today", name: "Inventory Aging", cat: "Inventory", desc: "How long current stock has been held, by lot" },
  { id: "fifo-lots", period: "today", name: "FIFO Cost Lot Tracking", cat: "Inventory", desc: "Purchase lots and FIFO consumption order per item" },
  { id: "weighted-avg", period: "today", name: "Weighted Average Summary", cat: "Inventory", desc: "Current weighted-average unit cost per item (used by the ledger)" },
  { id: "abc", period: "today", name: "ABC Classification", cat: "Inventory", desc: "Items ranked by value contribution - A / B / C" },
  { id: "tax-summary", period: "range", name: "Tax Summary", cat: "Tax", desc: "Tax charged and withheld by tax name, across invoices & bills" },
  { id: "project-summary", period: "alltime", name: "Project Summary", cat: "Projects", desc: "Revenue, cost and profit per project" },
  { id: "project-performance", period: "alltime", name: "Project Performance", cat: "Projects", desc: "Margin and budget usage per project" },
  { id: "budget-vs-actual", period: "month", name: "Budget vs Actual", cat: "Performance", desc: "Compare a month's spend/income to your budget" },
  { id: "ratios", period: "today", name: "Business Performance Ratios", cat: "Performance", desc: "Profitability and liquidity ratios" },
  { id: "realized-gl", period: "alltime", name: "Realized Gains & Losses", cat: "Performance", desc: "Posted gain/loss transactions over time" },
];

function Reports({ data, balances, setData, notify }) {
  const { styles, theme } = useUI();
  const [activeReport, setActiveReport] = useState(null);
  const [range, setRange] = useState("month");
  const [customFrom, setCustomFrom] = useState(localDateToISO(new Date(Date.now() - 13 * 86400000)));
  const [customTo, setCustomTo] = useState(todayStr());
  const [asOf, setAsOf] = useState(todayStr());
  const [month, setMonth] = useState(monthKey(todayStr()));
  const [showOptions, setShowOptions] = useState(false);
  const [periodsBack, setPeriodsBack] = useState(4);
  const toggleFav = (id) => setData(d => ({ ...d, favoriteReports: d.favoriteReports.includes(id) ? d.favoriteReports.filter(x => x !== id) : [...d.favoriteReports, id] }));

  // Resolve the range once: reports and the letterhead both receive either the
  // named range or the custom { from, to } object.
  const effRange = range === "custom" ? { from: customFrom, to: customTo } : range;

  if (activeReport) {
    const meta = REPORT_CATALOG.find(r => r.id === activeReport);
    const isFav = data.favoriteReports.includes(activeReport);
    return (
      <div>
        <div className="no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 8 }}>
          <button style={styles.btnGhost} onClick={() => setActiveReport(null)}>← All reports</button>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            {meta.period === "range" && <>
              <select style={styles.input} value={range} onChange={e => setRange(e.target.value)}>
                {RANGE_OPTIONS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
              {range === "custom" && <>
                <input type="date" style={styles.input} value={customFrom} onChange={e => setCustomFrom(e.target.value)} />
                <span style={{ fontSize: 12 }}>to</span>
                <input type="date" style={styles.input} value={customTo} onChange={e => setCustomTo(e.target.value)} />
              </>}
            </>}
            {meta.period === "asof" && <input type="date" style={styles.input} value={asOf} onChange={e => setAsOf(e.target.value)} />}
            {meta.period === "month" && <input type="month" style={styles.input} value={month} onChange={e => setMonth(e.target.value)} />}
            <button style={{ ...styles.btnGhost, color: isFav ? theme.amber : theme.text }} onClick={() => toggleFav(activeReport)}>{isFav ? "★ Favorited" : "☆ Add to favorites"}</button>
            <button style={styles.btnPrimary} onClick={() => window.print()}>Export / Print PDF</button>
            <div style={{ position: "relative" }}>
              <button style={{ ...styles.btnGhost, ...(showOptions ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }} onClick={() => setShowOptions(s => !s)} title="Report template options">⚙</button>
              {showOptions && <div onClick={() => setShowOptions(false)} style={{ position: "fixed", inset: 0, zIndex: 15 }} />}
              {showOptions && (
                <div style={{ position: "absolute", right: 0, top: "110%", zIndex: 20, background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 8, padding: "12px 14px", boxShadow: "0 8px 24px rgba(16,24,40,0.16)", minWidth: 230 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.muted, marginBottom: 8 }}>Report template</div>
                  <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, padding: "4px 0", cursor: "pointer" }}>
                    <input type="checkbox" checked={Boolean(data.settings.reportOptions?.showCodes)} onChange={e => setData(d => ({ ...d, settings: { ...d.settings, reportOptions: { ...d.settings.reportOptions, showCodes: e.target.checked } } }))} /> Show account codes
                  </label>
                  <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, padding: "4px 0", cursor: "pointer" }}>
                    <input type="checkbox" checked={data.settings.reportOptions?.hideZeroLines !== false} onChange={e => setData(d => ({ ...d, settings: { ...d.settings, reportOptions: { ...d.settings.reportOptions, hideZeroLines: e.target.checked } } }))} /> Hide lines with no activity
                  </label>
                  <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, padding: "4px 0", cursor: "pointer" }}>
                    <input type="checkbox" checked={Boolean(data.settings.reportOptions?.compareEnabled)} onChange={e => setData(d => ({ ...d, settings: { ...d.settings, reportOptions: { ...d.settings.reportOptions, compareEnabled: e.target.checked } } }))} /> Show period trend
                  </label>
                  {Boolean(data.settings.reportOptions?.compareEnabled) && (
                    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, padding: "4px 0 4px 21px" }}>
                      Number of periods
                      <select style={{ ...styles.inputSmall, marginLeft: "auto" }} value={periodsBack} onChange={e => setPeriodsBack(Number(e.target.value))}>
                        {Array.from({ length: 20 }, (_, i) => i + 1).map(n => <option key={n} value={n}>{n}</option>)}
                      </select>
                    </label>
                  )}
                  <div style={{ fontSize: 11, color: theme.muted, marginTop: 6 }}>Applies to all reports and saved automatically.</div>
                </div>
              )}
            </div>
          </div>
        </div>
        <div className="print-area">
          <ReportDocHeader data={data} meta={meta} range={effRange} asOf={asOf} month={month} />
          <ReportRenderer id={activeReport} data={data} balances={balances} setData={setData} notify={notify} range={effRange} asOf={asOf} month={month} periodsBack={periodsBack} onNavigate={setActiveReport} />
        </div>
      </div>
    );
  }

  const favorites = REPORT_CATALOG.filter(r => data.favoriteReports.includes(r.id));
  const categories = [...new Set(REPORT_CATALOG.map(r => r.cat))];
  return (
    <div>
      <PageHeader eyebrow="Analysis" title="Reports" sub={`${REPORT_CATALOG.length} reports available`} />
      {favorites.length > 0 && (
        <div style={{ marginBottom: 26 }}>
          <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.muted, marginBottom: 10 }}>★ Favorites</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px,1fr))", gap: 12 }}>
            {favorites.map(r => <ReportCard key={r.id} r={r} isFav onOpen={() => setActiveReport(r.id)} onToggleFav={() => toggleFav(r.id)} />)}
          </div>
        </div>
      )}
      {categories.map(cat => (
        <div key={cat} style={{ marginBottom: 26 }}>
          <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.muted, marginBottom: 10 }}>{cat}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px,1fr))", gap: 12 }}>
            {REPORT_CATALOG.filter(r => r.cat === cat).map(r => <ReportCard key={r.id} r={r} isFav={data.favoriteReports.includes(r.id)} onOpen={() => setActiveReport(r.id)} onToggleFav={() => toggleFav(r.id)} />)}
          </div>
        </div>
      ))}
    </div>
  );
}
// Formal letterhead shown at the top of every opened report - centered like a
// published financial statement, with the reporting period beneath the title.
function reportPeriodLabel(meta, range, asOf, month) {
  if (meta.period === "asof") return `As at ${fmtDate(asOf)}`;
  if (meta.period === "today") return `As at ${fmtDate(new Date())}`;
  if (meta.period === "alltime") return "All time to date";
  if (meta.period === "month") {
    const [y, m] = month.split("-").map(Number);
    return `For the month of ${new Date(y, m - 1, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" })}`;
  }
  if (range && typeof range === "object") return `For the period ${fmtDate(range.from)} to ${fmtDate(range.to)}`;
  if (range === "all") return "All time to date";
  if (range === "lastmonth") {
    const d = new Date(); const lm = new Date(d.getFullYear(), d.getMonth() - 1, 1);
    return `For the month of ${lm.toLocaleDateString(undefined, { month: "long", year: "numeric" })}`;
  }
  if (range === "year") return `For the year ending ${fmtDate(`${new Date().getFullYear()}-12-31`)}`;
  const { from, to } = rangeToDates(range);
  return `For the period ${fmtDate(from)} to ${fmtDate(to)}`;
}
// Report template options (persisted in settings): hide lines with no
// activity/balance, and optionally prefix account codes to line labels.
function reportOpts(data) { return { showCodes: false, hideZeroLines: true, compareEnabled: false, ...(data.settings.reportOptions || {}) }; }
function acctLabel(a, opts) { return opts.showCodes ? `${a.code} · ${a.name}` : a.name; }
function keepLine(value, opts) { return !opts.hideZeroLines || Math.abs(value) > 0.004; }
const STATEMENT_TITLES = { pl: "Statement of Profit or Loss", bs: "Statement of Financial Position", cf: "Statement of Cash Flows", "equity-movement": "Statement of Changes in Equity" };
function ReportDocHeader({ data, meta, range, asOf, month }) {
  const { theme } = useUI();
  const title = STATEMENT_TITLES[meta.id] || meta.name;
  return (
    <div style={{ borderBottom: `3px solid ${theme.accent}`, paddingBottom: 18, marginBottom: 24, textAlign: "center" }}>
      {data.settings.logo && <div className="ces-logo" style={{ padding: 9, display: "inline-flex", marginBottom: 10 }}><img src={data.settings.logo} alt="logo" style={{ height: 48, width: "auto", maxWidth: 130, objectFit: "contain" }} /></div>}
      <div style={{ fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontSize: 24, fontWeight: 600, color: theme.text }}>{data.settings.companyName}</div>
      <div style={{ fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontSize: 16, fontWeight: 500, color: theme.text, marginTop: 4 }}>{title}</div>
      <div style={{ fontSize: 12, color: theme.muted, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", marginTop: 4 }}>{reportPeriodLabel(meta, range, asOf, month)}</div>
    </div>
  );
}
function ReportCard({ r, isFav, onOpen, onToggleFav }) {
  const { styles, theme } = useUI();
  return (
    <div className="ces-card" style={{ ...styles.card, marginBottom: 0, cursor: "pointer", padding: "14px 16px" }} onClick={onOpen}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <div style={{ fontWeight: 600, fontSize: 13.5 }}>{r.name}</div>
        <button onClick={(e) => { e.stopPropagation(); onToggleFav(); }} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 15, color: isFav ? theme.amber : theme.muted, flexShrink: 0 }}>{isFav ? "★" : "☆"}</button>
      </div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4, lineHeight: 1.4 }}>{r.desc}</div>
    </div>
  );
}
function ReportRenderer({ id, data, balances, setData, notify, range, asOf, month, periodsBack, onNavigate }) {
  switch (id) {
    case "pl": return <ReportPL data={data} range={range} periodsBack={periodsBack} />;
    case "bs": return <ReportBS data={data} balances={balances} asOf={asOf} periodsBack={periodsBack} />;
    case "cf": return <ReportCF data={data} range={range} periodsBack={periodsBack} onNavigate={onNavigate} />;
    case "equity-movement": return <ReportEquityMovement data={data} range={range} periodsBack={periodsBack} />;
    case "trial-balance": return <ReportTrialBalance data={data} asOf={asOf} />;
    case "general-ledger": return <ReportGeneralLedger data={data} range={range} />;
    case "account-transactions": return <ReportAccountTransactions data={data} range={range} />;
    case "account-type-summary": return <ReportAccountTypeSummary data={data} balances={balances} />;
    case "journal-report": return <ReportJournalReport data={data} range={range} />;
    case "ar-aging": return <ReportAging data={data} kind="ar" />;
    case "ap-aging": return <ReportAging data={data} kind="ap" />;
    case "customer-balance-summary": return <ReportCustomerBalanceSummary data={data} />;
    case "receivables-details": return <ReportReceivablesDetails data={data} />;
    case "vendor-balance-summary": return <ReportVendorBalanceSummary data={data} />;
    case "payments-received": return <ReportPaymentsReceived data={data} range={range} />;
    case "time-to-get-paid": return <ReportTimeToGetPaid data={data} />;
    case "credit-note-details": return <ReportCreditNoteDetails data={data} />;
    case "recurring-transactions": return <ReportRecurringTransactions data={data} />;
    case "sales-by-item": return <ReportSalesByItem data={data} />;
    case "sales-by-customer": return <ReportSalesByCustomer data={data} />;
    case "sales-by-salesperson": return <ReportSalesBySalesperson data={data} />;
    case "profit-by-item": return <ReportProfitByItem data={data} />;
    case "order-fulfillment-item": return <ReportOrderFulfillment data={data} />;
    case "sales-return-history": return <ReportSalesReturnHistory data={data} />;
    case "refund-history": return <ReportRefundHistory data={data} />;
    case "expenses-by-category": return <ReportExpensesByCategory data={data} range={range} />;
    case "expenses-by-project": return <ReportExpensesByProject data={data} range={range} />;
    case "purchase-by-vendor": return <ReportPurchaseByVendor data={data} />;
    case "purchase-by-item": return <ReportPurchaseByItem data={data} />;
    case "goods-received-history": return <ReportGoodsReceivedHistory data={data} />;
    case "goods-received-by-item": return <ReportGoodsReceivedByItem data={data} />;
    case "inventory-summary": return <ReportInventorySummary data={data} />;
    case "fixed-assets-by-location": return <ReportFixedAssetsByLocation data={data} />;
    case "fixed-assets-by-department": return <ReportFixedAssetsByDepartment data={data} />;
    case "inventory-by-location": return <ReportInventoryByLocation data={data} />;
    case "inventory-by-department": return <ReportInventoryByDepartment data={data} />;
    case "inventory-aging": return <ReportInventoryAging data={data} />;
    case "committed-stock": return <ReportCommittedStock data={data} />;
    case "inventory-movement-history": return <ReportInventoryMovementHistory data={data} range={range} />;
    case "inventory-turnover-item": return <ReportInventoryTurnoverByItem data={data} range={range} />;
    case "inventory-turnover-amount": return <ReportInventoryTurnoverByAmount data={data} range={range} />;
    case "fifo-lots": return <ReportFIFO data={data} />;
    case "weighted-avg": return <ReportWeightedAvg data={data} />;
    case "abc": return <ReportABC data={data} />;
    case "tax-summary": return <ReportTaxSummary data={data} range={range} />;
    case "reconciliation-status": return <ReportReconciliationStatus data={data} />;
    case "unrealized-gain-loss": return <ReportUnrealizedGainLoss data={data} balances={balances} />;
    case "timesheet-details": return <ReportTimesheetDetails data={data} range={range} />;
    case "project-details": return <ReportProjectDetails data={data} />;
    case "project-cost-summary": return <ReportProjectCostSummary data={data} />;
    case "project-summary": return <ReportProjectSummary data={data} />;
    case "project-performance": return <ReportProjectPerformance data={data} />;
    case "budget-vs-actual": return <ReportBudgetVsActual data={data} balances={balances} month={month} />;
    case "ratios": return <ReportRatiosCard data={data} balances={balances} />;
    case "realized-gl": return <ReportRealizedGL data={data} />;
    default: return null;
  }
}

// Cash-basis P&L movement: revenue/expense from invoices and bills is moved
// from the invoice/bill date to the date(s) cash actually changed hands,
// pro-rated by how much of the document each payment settled. Entries that
// already represent real cash movement at posting (quick expenses, payroll,
// manual entries, bank imports, refunds) are counted on their own date.
// Non-cash entries (depreciation, disposal write-offs) are excluded, matching
// standard cash-basis treatment.
function cashBasisMovement(data, from, to) {
  const mv = {};
  data.accounts.forEach(a => { mv[a.id] = 0; });
  const add = (accountId, amount) => { if (accountId in mv) mv[accountId] += amount; };

  data.transactions.forEach(t => {
    if (t.source === "invoice" || t.source === "bill" || t.source === "depreciation") return; // handled below or excluded
    if (t.date < from || t.date > to) return;
    t.lines.forEach(l => {
      if (l.accountId === "6000") return; // non-cash gain/loss adjustment
      const acc = data.accounts.find(a => a.id === l.accountId);
      if (!acc || (acc.type !== "revenue" && acc.type !== "expense")) return;
      add(l.accountId, acctNormal(acc) ? (l.debit - l.credit) : (l.credit - l.debit));
    });
  });

  const allocate = (docs, relatedType, paymentType) => {
    docs.forEach(doc => {
      const total = computeDocTotals(doc, data.taxGroups).finalAmount;
      if (total <= 0) return;
      const docTxn = data.transactions.find(t => t.docId === doc.id && (t.source === "invoice" || t.source === "bill"))
        || data.transactions.find(t => (t.source === "invoice" || t.source === "bill") && (t.memo || "").includes(doc.id));
      if (!docTxn) return;
      const lineMoves = docTxn.lines
        .map(l => { const acc = data.accounts.find(a => a.id === l.accountId); return acc && (acc.type === "revenue" || acc.type === "expense") ? { accountId: l.accountId, amt: acctNormal(acc) ? (l.debit - l.credit) : (l.credit - l.debit) } : null; })
        .filter(Boolean);
      data.payments.filter(p => p.relatedType === relatedType && p.relatedId === doc.id && p.type === paymentType && p.date >= from && p.date <= to)
        .forEach(p => { const frac = p.amount / total; lineMoves.forEach(({ accountId, amt }) => add(accountId, amt * frac)); });
    });
  };
  allocate(data.invoices, "invoice", "received");
  allocate(data.bills, "bill", "paid");
  return mv;
}
function ReportPL({ data, range, periodsBack }) {
  const { styles, theme, openDrilldown } = useUI();
  const opts = reportOpts(data);
  const cashBasis = data.settings.accountingBasis === "cash";
  const [view, setView] = useState("table");
  const n = opts.compareEnabled ? Math.max(1, periodsBack || 1) : 1;
  const trendRanges = computeTrendRanges(range, n);
  const movementFor = (r) => { const { from, to } = rangeToDates(r); return cashBasis ? cashBasisMovement(data, from, to) : periodMovement(data.accounts, data.transactions, from, to); };
  const mvList = trendRanges.map(movementFor);
  const mv = mvList[mvList.length - 1]; // current period, used for "which lines to show"
  const curRange = rangeToDates(trendRanges[trendRanges.length - 1]);
  const drill = (accountIds, label) => openDrilldown({ accountIds, label, range: curRange });
  const lineVal = (a, m) => a.contra ? -(m[a.id] || 0) : (m[a.id] || 0);
  const revenue = data.accounts.filter(a => a.type === "revenue" && a.subtype !== "other");
  const cogs = data.accounts.filter(a => a.type === "expense" && a.subtype === "cogs");
  const expense = data.accounts.filter(a => a.type === "expense" && a.subtype !== "cogs");
  const nonOp = data.accounts.filter(a => a.subtype === "other");
  const anyPeriodKeeps = (a) => mvList.some(m => keepLine(lineVal(a, m), opts));
  const revenueShown = revenue.filter(anyPeriodKeeps);
  const cogsShown = cogs.filter(anyPeriodKeeps);
  const expenseShown = expense.filter(anyPeriodKeeps);
  const nonOpShown = nonOp.filter(a => mvList.some(m => keepLine(m[a.id] || 0, opts)));
  const totalRevList = mvList.map(m => revenue.reduce((s, a) => s + lineVal(a, m), 0));
  const totalCogsList = mvList.map(m => cogs.reduce((s, a) => s + lineVal(a, m), 0));
  const totalExpList = mvList.map(m => expense.reduce((s, a) => s + lineVal(a, m), 0));
  const totalNonOpList = mvList.map(m => nonOp.reduce((s, a) => s + (m[a.id] || 0), 0));
  const grossProfitList = totalRevList.map((r, i) => r - totalCogsList[i]);
  const netIncomeList = grossProfitList.map((g, i) => g - totalExpList[i] + totalNonOpList[i]);
  const labels = trendRanges.map(periodLabel);
  const showCogsBlock = cogsShown.length > 0 || !opts.hideZeroLines || totalCogsList.some(v => Math.abs(v) > 0.004);
  const T = (props) => <TrendRow {...props} />;
  return (
    <div>
      {n > 1 && <TrendViewSwitch view={view} setView={setView} />}
      <div className="ces-card" style={styles.card}>
        {view === "table" ? <>
          {n > 1 && <TrendHeader labels={labels} />}
          <T label="Revenue" periodCount={n} bold />
          {revenueShown.map(a => <T key={a.id} label={acctLabel(a, opts)} values={mvList.map(m => lineVal(a, m))} indent onClick={() => drill([a.id], acctLabel(a, opts))} />)}
          <T label="Total revenue" values={totalRevList} bold divider />
          {showCogsBlock && <>
            <T label="Cost of sales" periodCount={n} bold />
            {cogsShown.map(a => <T key={a.id} label={acctLabel(a, opts)} values={mvList.map(m => lineVal(a, m))} indent onClick={() => drill([a.id], acctLabel(a, opts))} />)}
            <T label="Total cost of sales" values={totalCogsList} bold divider />
            <T label="Gross profit" values={grossProfitList} bold divider tone={grossProfitList[grossProfitList.length - 1] >= 0 ? "emerald" : "rose"} />
          </>}
          <T label="Operating expenses" periodCount={n} bold />
          {expenseShown.map(a => <T key={a.id} label={acctLabel(a, opts)} values={mvList.map(m => lineVal(a, m))} indent onClick={() => drill([a.id], acctLabel(a, opts))} />)}
          <T label="Total operating expenses" values={totalExpList} bold divider />
          {nonOpShown.length > 0 && <><T label="Non-operating" periodCount={n} bold />{nonOpShown.map(a => <T key={a.id} label={acctLabel(a, opts)} values={mvList.map(m => m[a.id] || 0)} indent onClick={() => drill([a.id], acctLabel(a, opts))} />)}</>}
          <T label="Net income" values={netIncomeList} bold tone={netIncomeList[netIncomeList.length - 1] >= 0 ? "emerald" : "rose"} divider />
        </> : (
          <TrendChart labels={labels} colors={[theme.accent, theme.rose, theme.emerald]}
            series={[{ key: "rev", label: "Revenue", values: totalRevList }, { key: "exp", label: "Expenses", values: totalExpList.map((v, i) => v + totalCogsList[i]) }, { key: "net", label: "Net income", values: netIncomeList }]} />
        )}
        {cashBasis && <div style={{ fontSize: 11.5, color: theme.muted, marginTop: 10 }}>Shown on a cash basis: revenue and expenses are recognized when cash was actually received or paid, not when invoiced or billed. Depreciation and asset write-offs are non-cash and excluded. Switch back to accrual in Settings.</div>}
      </div>
    </div>
  );
}

function ReportBS({ data, balances, asOf, periodsBack }) {
  const { styles, theme, openDrilldown } = useUI();
  const opts = reportOpts(data);
  const [view, setView] = useState("table");
  const n = opts.compareEnabled ? Math.max(1, periodsBack || 1) : 1;
  const dates = computeTrendDates(asOf, n);
  const balList = dates.map(d => d === todayStr() ? balances : computeBalances(data.accounts, data.transactions, d));
  const bal = balList[balList.length - 1];
  const curAsOf = dates[dates.length - 1];
  const drill = (accountIds, label) => openDrilldown({ accountIds, label, asOf: curAsOf });
  const bsVal = (a, b) => a.contra ? -(b[a.id] || 0) : (b[a.id] || 0);
  const assets = data.accounts.filter(a => a.type === "asset" && balList.some(b => keepLine(bsVal(a, b), opts)));
  const liabilities = data.accounts.filter(a => a.type === "liability" && balList.some(b => keepLine(b[a.id] || 0, opts)));
  const totalAssetsList = balList.map(b => sumByType(data.accounts, b, "asset"));
  const totalLiabList = balList.map(b => sumByType(data.accounts, b, "liability"));
  const totalEquityList = dates.map((d, i) => sumByType(data.accounts, balList[i], "equity") + netIncomeAllTime(data, d));
  const labels = dates.map(d => n > 1 ? fmtDate(d) : fmtDate(d));
  const T = (props) => <TrendRow {...props} />;
  const totalAssets = totalAssetsList[totalAssetsList.length - 1], totalLiab = totalLiabList[totalLiabList.length - 1], totalEquity = totalEquityList[totalEquityList.length - 1];
  return (
    <div>
      {n > 1 && <TrendViewSwitch view={view} setView={setView} />}
      <div className="ces-card" style={styles.card}>
        {view === "table" ? <>
          {n > 1 && <TrendHeader labels={labels} />}
          <T label="Assets" periodCount={n} bold />
          {assets.map(a => <T key={a.id} label={acctLabel(a, opts)} values={balList.map(b => bsVal(a, b))} indent onClick={() => drill([a.id], acctLabel(a, opts))} />)}
          <T label="Total assets" values={totalAssetsList} bold divider />
          <T label="Liabilities" periodCount={n} bold />
          {liabilities.map(a => <T key={a.id} label={acctLabel(a, opts)} values={balList.map(b => b[a.id] || 0)} indent onClick={() => drill([a.id], acctLabel(a, opts))} />)}
          <T label="Total liabilities" values={totalLiabList} bold divider />
          <T label="Equity (incl. retained earnings)" values={totalEquityList} bold divider />
        </> : (
          <TrendChart labels={labels} colors={[theme.accent, theme.rose, theme.emerald]}
            series={[{ key: "assets", label: "Total assets", values: totalAssetsList }, { key: "liab", label: "Total liabilities", values: totalLiabList }, { key: "eq", label: "Total equity", values: totalEquityList }]} />
        )}
        <div style={{ fontSize: 11, color: theme.muted, marginTop: 6, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{Math.abs(totalAssets - (totalLiab + totalEquity)) < 0.5 ? "\u2713 balanced: assets = liabilities + equity" : "\u26a0 out of balance - check entries"}</div>
      </div>
    </div>
  );
}

// Pure indirect-method cash flow calculation for one period, factored out so
// it can run once per period in a trend.
// IAS 12 deferred tax: for each fixed asset, the gap between its book net
// book value and its tax written-down value is a temporary difference. A
// positive gap (book NBV > tax WDV - the common case when capital allowances
// run faster than book depreciation) is a taxable temporary difference and
// creates a Deferred Tax Liability; a negative gap creates a Deferred Tax
// Asset. Presented net (one account or the other), which IAS 12 permits when
// there's a legally enforceable right to offset within the same tax
// jurisdiction - the normal case for a single-entity SME.
// IFRS 16: present value of a lease's remaining fixed payments at a given
// monthly discount rate - used both to recognize the right-of-use asset and
// lease liability at commencement, and to build the full amortization table.
function leasePV(paymentAmount, termMonths, annualRatePct) {
  const r = (annualRatePct || 0) / 100 / 12;
  if (r === 0) return paymentAmount * termMonths;
  return paymentAmount * (1 - Math.pow(1 + r, -termMonths)) / r;
}
// Full period-by-period amortization schedule: each row splits the fixed
// payment into interest (unwinding the liability at the discount rate) and
// principal (reducing the liability), alongside straight-line ROU
// depreciation over the lease term.
function computeLeaseSchedule(lease) {
  const r = (lease.discountRate || 0) / 100 / 12;
  const monthlyDep = lease.rouCost / lease.termMonths;
  const rows = [];
  let balance = lease.rouCost;
  for (let i = 1; i <= lease.termMonths; i++) {
    const interest = balance * r;
    const principal = lease.paymentAmount - interest;
    const opening = balance;
    balance = Math.max(0, balance - principal);
    rows.push({ period: i, opening, interest, principal, payment: lease.paymentAmount, closing: balance, depreciation: monthlyDep });
  }
  return rows;
}
// Current portion (principal due within 12 months) vs non-current portion of
// the remaining lease liability, as of how many periods have already run -
// shown as analysis since this system posts to a single Lease Liability
// account rather than actively reclassifying between two GL accounts.
// IAS 2: inventory is carried at the lower of cost and net realizable
// value. Where NRV has fallen below the weighted-average cost, the item is
// written down to NRV, expensed within cost of sales. The original cost is
// remembered (costBeforeNRV) so a later recovery in NRV can be reversed -
// but never above that original cost, per IAS 2's reversal cap.
function computeNRVCheck(data) {
  return data.inventory.map(item => {
    const nrv = item.nrv ?? item.salePrice ?? item.unitCost;
    const cost = item.unitCost;
    const original = item.costBeforeNRV ?? cost;
    const writeDownPerUnit = Math.max(0, cost - nrv);
    const reversalPerUnit = item.costBeforeNRV ? Math.max(0, Math.min(nrv, original) - cost) : 0;
    return { id: item.id, name: item.name, qty: item.qty, cost, nrv, original, writeDownAmount: writeDownPerUnit * item.qty, reversalAmount: reversalPerUnit * item.qty, newUnitCost: writeDownPerUnit > 0 ? nrv : reversalPerUnit > 0 ? Math.min(nrv, original) : cost };
  }).filter(r => r.writeDownAmount > 0.5 || r.reversalAmount > 0.5);
}
// IFRS 9 simplified approach for trade receivables: a provision matrix
// applying a loss rate per aging bucket to outstanding invoice balances.
// IFRS 15: how much of a deferred revenue schedule is due to be recognized
// this run, straight-line over its service period.
function computeDeferredRevenueMovement(schedule) {
  const monthlyAmount = schedule.totalAmount / schedule.months;
  const remaining = Math.max(0, schedule.totalAmount - schedule.recognizedAmount);
  return { monthlyAmount, remaining, dueAmount: Math.min(monthlyAmount, remaining) };
}
function computeECL(data) {
  const rates = data.settings.eclRates || {};
  const byBucket = {};
  AGING_BUCKETS.forEach(b => { byBucket[b] = 0; });
  data.invoices.forEach(inv => {
    const total = computeDocTotals(inv, data.taxGroups).finalAmount;
    const outstanding = total - (inv.amountPaid || 0);
    if (outstanding <= 0.5) return;
    const days = inv.dueDate ? Math.floor((new Date(todayStr()) - new Date(inv.dueDate)) / 86400000) : 0;
    byBucket[agingBucket(days)] += outstanding;
  });
  const rows = AGING_BUCKETS.map(b => ({ bucket: b, outstanding: byBucket[b], rate: rates[b] ?? 0, provision: byBucket[b] * (rates[b] ?? 0) / 100 }));
  const totalOutstanding = rows.reduce((s, r) => s + r.outstanding, 0);
  const targetProvision = rows.reduce((s, r) => s + r.provision, 0);
  return { rows, totalOutstanding, targetProvision };
}
function leaseLiabilitySplit(lease) {
  const schedule = computeLeaseSchedule(lease);
  const remaining = schedule.slice(lease.periodsRun || 0);
  const next12 = remaining.slice(0, 12);
  const current = next12.reduce((s, r) => s + r.principal, 0);
  return { current: Math.min(current, lease.liabilityBalance), nonCurrent: Math.max(0, lease.liabilityBalance - current) };
}
function computeDeferredTax(data) {

  const rate = (data.settings.corporateTaxRate || 30) / 100;
  const perAsset = data.fixedAssets.map(a => {
    const bookNBV = a.cost - a.accumulatedDepreciation;
    const taxWDV = a.cost - (a.accumulatedTaxDepreciation || 0);
    return { id: a.id, name: a.name, bookNBV, taxWDV, tempDiff: bookNBV - taxWDV };
  });
  const totalTempDiff = perAsset.reduce((s, a) => s + a.tempDiff, 0);
  const targetNet = totalTempDiff * rate; // >0 net DTL, <0 net DTA
  const bal = computeBalances(data.accounts, data.transactions);
  const recognizedDTL = bal["2260"] || 0, recognizedDTA = bal["1450"] || 0;
  const desiredDTL = Math.max(0, targetNet), desiredDTA = Math.max(0, -targetNet);
  const dDTL = desiredDTL - recognizedDTL, dDTA = desiredDTA - recognizedDTA;
  const movement = dDTL - dDTA; // net change in tax expense this adjustment represents
  return { perAsset, totalTempDiff, rate, targetNet, recognizedDTL, recognizedDTA, recognizedNet: recognizedDTL - recognizedDTA, dDTL, dDTA, movement };
}
// A self-balancing journal moving DTL/DTA to their target balances, with
// Income Tax Expense as the plug - the balance identity holds by
// construction regardless of which direction each account is moving.
function deferredTaxAdjustmentLines(dt) {
  const lines = [];
  if (Math.abs(dt.dDTL) > 0.5) lines.push({ accountId: "2260", debit: dt.dDTL < 0 ? -dt.dDTL : 0, credit: dt.dDTL > 0 ? dt.dDTL : 0 });
  if (Math.abs(dt.dDTA) > 0.5) lines.push({ accountId: "1450", debit: dt.dDTA > 0 ? dt.dDTA : 0, credit: dt.dDTA < 0 ? -dt.dDTA : 0 });
  if (Math.abs(dt.movement) > 0.5) lines.push({ accountId: "5850", debit: dt.movement > 0 ? dt.movement : 0, credit: dt.movement < 0 ? -dt.movement : 0 });
  return lines;
}
// Current tax provision for a period: taxable profit reverses the accounting
// (book) depreciation charge and substitutes the tax capital allowance
// instead, since only the latter is tax-deductible. Tax depreciation for the
// period is estimated pro-rata from each asset's annual capital allowance
// rate, since (unlike book depreciation) it isn't posted to the ledger.
function computeCurrentTax(data, range) {
  const { from, to } = rangeToDates(range);
  const mv = periodMovement(data.accounts, data.transactions, from, to);
  const accountingProfit = sumByType(data.accounts, mv, "revenue") - sumByType(data.accounts, mv, "expense");
  const bookDepreciation = mv["5700"] || 0;
  const taxDepreciation = data.fixedAssets.reduce((s, a) => {
    const activeFrom = a.purchaseDate > from ? a.purchaseDate : from;
    if (activeFrom > to) return s; // not yet acquired during this period
    const activeMonths = Math.max(0, (new Date(to) - new Date(activeFrom)) / (30.44 * 86400000));
    const cap = a.cost - a.salvageValue - (a.accumulatedTaxDepreciation || 0);
    return s + Math.max(0, Math.min(((a.cost - a.salvageValue) * (a.taxRatePct || 25) / 100 / 12) * activeMonths, cap));
  }, 0);
  const taxableProfit = accountingProfit + bookDepreciation - taxDepreciation;
  const rate = (data.settings.corporateTaxRate || 30) / 100;
  const currentTax = Math.max(0, taxableProfit) * rate;
  return { accountingProfit, bookDepreciation, taxDepreciation, taxableProfit, rate, currentTax };
}
function computeCashFlow(data, range) {
  const { from, to } = rangeToDates(range);
  const dayBefore = new Date(from); dayBefore.setDate(dayBefore.getDate() - 1);
  const openBal = computeBalances(data.accounts, data.transactions, localDateToISO(dayBefore));
  const closeBal = computeBalances(data.accounts, data.transactions, to);
  const mv = periodMovement(data.accounts, data.transactions, from, to);
  const delta = (id) => (closeBal[id] || 0) - (openBal[id] || 0);
  const sumDelta = (pred) => data.accounts.filter(pred).reduce((s, a) => s + delta(a.id), 0);

  const netIncome = sumByType(data.accounts, mv, "revenue") - sumByType(data.accounts, mv, "expense");
  const depreciation = mv["5700"] || 0;
  const impairment = mv["5760"] || 0; // non-cash, add back - same treatment as depreciation
  const revaluationLoss = mv["5780"] || 0; // non-cash, add back - same treatment as impairment
  // The Allowance for Doubtful Accounts is a contra-asset stored with a
  // credit-normal balance - a plain delta sweep of "current assets" would
  // treat a bigger provision as if receivables themselves had grown
  // (subtracting it, as if cash were tied up), when the opposite is true:
  // it's a non-cash expense that needs the same add-back as depreciation.
  const badDebtProvisionDelta = delta("1150");
  const gainsLosses = -(mv["6000"] || 0);
  // Deferred tax movements are entirely non-cash - the portion of income tax
  // expense that was deferred rather than paid needs the same add-back
  // treatment as depreciation, otherwise cash from operations understates
  // real performance by a purely bookkeeping tax provision.
  const deferredTaxDelta = delta("2260") - delta("1450");
  const dAR = delta("1100"), dInventory = ["1200", "1201", "1202", "1203"].reduce((s, id) => s + delta(id), 0);
  const dOtherCurrentAssets = sumDelta(a => a.type === "asset" && a.subtype === "current" && !["1100", "1200", "1201", "1202", "1203", "1150"].includes(a.id) && !(a.name || "").toLowerCase().includes("cash"));
  const dAP = delta("2000");
  const dOtherCurrentLiab = sumDelta(a => a.type === "liability" && a.subtype === "current" && a.id !== "2000");
  // Provisions: every movement in this account pairs with either an expense
  // line (recognizing, increasing, or releasing one - all non-cash) or a
  // cash line (utilizing one against a real payment). A plain delta
  // add-back correctly self-cancels the non-cash cases against their
  // matching expense movement in netIncome, and correctly leaves the cash
  // case as a real operating outflow, without needing to inspect which kind
  // of transaction caused it.
  const provisionDelta = delta("2295");
  const cfOperating = netIncome + depreciation + impairment + revaluationLoss + badDebtProvisionDelta + deferredTaxDelta + provisionDelta + gainsLosses - dAR - dInventory - dOtherCurrentAssets + dAP + dOtherCurrentLiab;

  // Investing: computed directly from transactions touching fixed/intangible
  // asset accounts, not from balance deltas (which would conflate the
  // non-cash depreciation charge, already added back above, with real cash
  // purchases and with credit-financed purchases that have zero cash effect).
  let investingOut = 0, investingIn = 0;
  data.transactions.forEach(t => {
    if (t.date < from || t.date > to) return;
    t.lines.forEach(l => {
      const acc = data.accounts.find(a => a.id === l.accountId);
      if (!acc || acc.subtype !== "fixed" || acc.contra) return;
      if (l.debit > l.credit) {
        const cashLine = t.lines.find(x => { const xa = data.accounts.find(a => a.id === x.accountId); return xa && (xa.subtype === "bank" || (xa.name || "").toLowerCase().includes("cash")); });
        if (cashLine) investingOut += Math.min(l.debit, cashLine.credit);
      }
    });
  });
  const cfInvest = investingIn - investingOut;

  // Repayment of lease liabilities: like fixed-asset purchases, this can't
  // be read off the account's balance delta, since the same account also
  // moves for a non-cash reason (lease commencement, capitalizing the
  // right-of-use asset). Only the portion of any decrease actually paired
  // with a real bank payment counts as a financing cash outflow.
  let leaseRepayment = 0;
  data.transactions.forEach(t => {
    if (t.date < from || t.date > to) return;
    const leaseLine = t.lines.find(l => l.accountId === "2270");
    if (!leaseLine || leaseLine.debit <= leaseLine.credit) return; // only paydowns (debit) count
    const cashLine = t.lines.find(x => { const xa = data.accounts.find(a => a.id === x.accountId); return xa && (xa.subtype === "bank" || (xa.name || "").toLowerCase().includes("cash")); });
    if (cashLine) leaseRepayment += Math.min(leaseLine.debit, cashLine.credit);
  });

  const dEquity = sumDelta(a => a.type === "equity" && a.id !== "3200" && a.id !== "3900");
  // Deferred tax liability, lease liability, and provisions are excluded
  // here - each is handled above via its own non-cash add-back or actual
  // cash-payment scan, since a plain balance delta would otherwise count
  // their non-cash origination as real financing cash. Revaluation Surplus
  // (excluded from dEquity above) is a non-cash OCI movement that never
  // touches cash either.
  const dNonCurrentLiab = sumDelta(a => a.type === "liability" && a.subtype !== "current" && a.id !== "2260" && a.id !== "2270" && a.id !== "2295");
  const cfFinancing = dEquity + dNonCurrentLiab - leaseRepayment;

  const cashAccounts = data.accounts.filter(a => a.subtype === "bank" || (a.type === "asset" && (a.name || "").toLowerCase().includes("cash")));
  const actualCashChange = cashAccounts.reduce((s, a) => s + delta(a.id), 0);
  const computedChange = cfOperating + cfInvest + cfFinancing;
  const openingCash = cashAccounts.reduce((s, a) => s + (openBal[a.id] || 0), 0);

  return { netIncome, depreciation, impairment, revaluationLoss, badDebtProvisionDelta, deferredTaxDelta, provisionDelta, gainsLosses, dAR, dInventory, dOtherCurrentAssets, dAP, dOtherCurrentLiab, cfOperating, investingOut, investingIn, cfInvest, dEquity, dNonCurrentLiab, leaseRepayment, cfFinancing, actualCashChange, computedChange, openingCash, reconciles: Math.abs(computedChange - actualCashChange) < 1 };
}
function ReportCF({ data, range, periodsBack, onNavigate }) {
  const { styles, theme, openDrilldown } = useUI();
  const opts = reportOpts(data);
  const [view, setView] = useState("table");
  const n = opts.compareEnabled ? Math.max(1, periodsBack || 1) : 1;
  const trendRanges = computeTrendRanges(range, n);
  const cList = trendRanges.map(r => computeCashFlow(data, r));
  const c = cList[cList.length - 1];
  const labels = trendRanges.map(periodLabel);
  const T = (props) => <TrendRow {...props} />;
  const curRange = rangeToDates(trendRanges[trendRanges.length - 1]);
  const drill = (accountIds, label) => openDrilldown({ accountIds, label, range: curRange });

  // Single period: full detailed statement, as before. Multiple periods: a
  // summarized trend (headline totals per period) - the full line-item
  // detail repeated across many periods would be unreadably wide.
  if (n === 1) {
    return (
      <div>
        <div className="ces-card" style={styles.card}>
          <RowLine label="Cash flows from operating activities" bold />
          <RowLine label="Net income" value={c.netIncome} indent onClick={() => onNavigate && onNavigate("pl")} />
          <RowLine label="Add back: depreciation & amortization" value={c.depreciation} indent onClick={() => drill(["5700"], "Depreciation & amortization")} />
          {Math.abs(c.deferredTaxDelta) > 0.5 && <RowLine label={c.deferredTaxDelta >= 0 ? "Add back: deferred tax (non-cash)" : "Less: deferred tax (non-cash)"} value={c.deferredTaxDelta} indent />}
          {c.impairment > 0.5 && <RowLine label="Add back: impairment losses" value={c.impairment} indent />}
          {Math.abs(c.badDebtProvisionDelta) > 0.5 && <RowLine label={c.badDebtProvisionDelta >= 0 ? "Add back: bad debt provision" : "Less: bad debt provision released"} value={c.badDebtProvisionDelta} indent />}
          {c.revaluationLoss > 0.5 && <RowLine label="Add back: revaluation losses" value={c.revaluationLoss} indent />}
          {Math.abs(c.provisionDelta) > 0.5 && <RowLine label={c.provisionDelta >= 0 ? "Add back: provisions recognized" : "Less: provisions paid out / released"} value={c.provisionDelta} indent />}
          {Math.abs(c.gainsLosses) > 0.5 && <RowLine label={c.gainsLosses >= 0 ? "Add back: loss on disposal of assets" : "Less: gain on disposal of assets"} value={c.gainsLosses} indent />}
          <RowLine label={c.dAR >= 0 ? "(Increase) in accounts receivable" : "Decrease in accounts receivable"} value={-c.dAR} indent onClick={() => drill(["1100"], "Accounts Receivable")} />
          <RowLine label={c.dInventory >= 0 ? "(Increase) in inventory" : "Decrease in inventory"} value={-c.dInventory} indent onClick={() => drill(["1200", "1201", "1202", "1203"], "Inventory")} />
          {Math.abs(c.dOtherCurrentAssets) > 0.5 && <RowLine label={c.dOtherCurrentAssets >= 0 ? "(Increase) in other current assets" : "Decrease in other current assets"} value={-c.dOtherCurrentAssets} indent />}
          <RowLine label={c.dAP >= 0 ? "Increase in accounts payable" : "(Decrease) in accounts payable"} value={c.dAP} indent onClick={() => drill(["2000"], "Accounts Payable")} />
          {Math.abs(c.dOtherCurrentLiab) > 0.5 && <RowLine label={c.dOtherCurrentLiab >= 0 ? "Increase in other current liabilities" : "(Decrease) in other current liabilities"} value={c.dOtherCurrentLiab} indent />}
          <RowLine label="Net cash from operating activities" value={c.cfOperating} bold divider tone={c.cfOperating >= 0 ? "emerald" : "rose"} />
          <div style={{ height: 14 }} />
          <RowLine label="Cash flows from investing activities" bold />
          <RowLine label={c.cfInvest <= 0 ? "Purchase of fixed & intangible assets (cash-funded)" : "Net proceeds from disposals of fixed assets"} value={c.cfInvest} indent />
          <RowLine label="Net cash used in investing activities" value={c.cfInvest} bold divider tone={c.cfInvest >= 0 ? "emerald" : "rose"} />
          <div style={{ height: 14 }} />
          <RowLine label="Cash flows from financing activities" bold />
          {Math.abs(c.dEquity) > 0.5 && <RowLine label={c.dEquity >= 0 ? "Owner contributions / capital raised" : "Owner withdrawals"} value={c.dEquity} indent />}
          {Math.abs(c.dNonCurrentLiab) > 0.5 && <RowLine label={c.dNonCurrentLiab >= 0 ? "Proceeds from borrowings" : "Repayment of borrowings"} value={c.dNonCurrentLiab} indent />}
          {c.leaseRepayment > 0.5 && <RowLine label="Repayment of lease liabilities" value={-c.leaseRepayment} indent />}
          {Math.abs(c.dEquity) <= 0.5 && Math.abs(c.dNonCurrentLiab) <= 0.5 && c.leaseRepayment <= 0.5 && <RowLine label="No financing activity in the period" value={0} indent />}
          <RowLine label="Net cash from financing activities" value={c.cfFinancing} bold divider tone={c.cfFinancing >= 0 ? "emerald" : "rose"} />
          <div style={{ height: 14 }} />
          <RowLine label="Net increase / (decrease) in cash" value={c.computedChange} bold divider />
          <RowLine label="Cash & equivalents at start of period" value={c.openingCash} indent />
          <RowLine label="Cash & equivalents at end of period" value={c.openingCash + c.actualCashChange} bold divider />
          <div style={{ fontSize: 11.5, color: c.reconciles ? theme.muted : theme.rose, marginTop: 8, fontVariantNumeric: "tabular-nums" }}>
            {c.reconciles ? "\u2713 Statement reconciles to the actual movement on cash and bank accounts." : `\u26a0 Indirect-method figure differs from actual cash movement by ${Math.round(c.computedChange - c.actualCashChange)} - usually caused by entries posted directly between cash and non-standard accounts.`}
          </div>
        </div>
      </div>
    );
  }

  const opList = cList.map(x => x.cfOperating), investList = cList.map(x => x.cfInvest), finList = cList.map(x => x.cfFinancing), netList = cList.map(x => x.computedChange);
  const closingList = cList.map(x => x.openingCash + x.actualCashChange);
  return (
    <div>
      <TrendViewSwitch view={view} setView={setView} />
      <div className="ces-card" style={styles.card}>
        {view === "table" ? <>
          <TrendHeader labels={labels} />
          <T label="Net cash from operating activities" values={opList} bold tone={opList[opList.length - 1] >= 0 ? "emerald" : "rose"} />
          <T label="Net cash used in investing activities" values={investList} bold tone={investList[investList.length - 1] >= 0 ? "emerald" : "rose"} />
          <T label="Net cash from financing activities" values={finList} bold tone={finList[finList.length - 1] >= 0 ? "emerald" : "rose"} />
          <T label="Net increase / (decrease) in cash" values={netList} bold divider />
          <T label="Cash & equivalents at end of period" values={closingList} bold divider />
        </> : (
          <TrendChart labels={labels} colors={[theme.accent, theme.rose, theme.emerald, theme.amber]}
            series={[{ key: "op", label: "Operating", values: opList }, { key: "inv", label: "Investing", values: investList }, { key: "fin", label: "Financing", values: finList }, { key: "net", label: "Net change", values: netList }]} />
        )}
        {!cList.every(x => x.reconciles) && <div style={{ fontSize: 11.5, color: theme.rose, marginTop: 8 }}>⚠ One or more periods shown do not fully reconcile to the actual cash movement - open a single-period view for the detailed breakdown.</div>}
      </div>
    </div>
  );
}

// Per-component equity roll-forward for one period: separates "net income
// for the period" from "other equity movements" (contributions, drawings,
// transfers) per category. Retained Earnings always appears as a component
// even with no dedicated account, since undistributed profit accumulates
// there implicitly in this app's model (no formal year-end closing entry).
function computeEquityMovement(data, range) {
  const { from, to } = rangeToDates(range);
  const dayBefore = new Date(from); dayBefore.setDate(dayBefore.getDate() - 1);
  const openStr = localDateToISO(dayBefore);
  const openBal = computeBalances(data.accounts, data.transactions, openStr);
  const mv = periodMovement(data.accounts, data.transactions, from, to);
  const equityAccounts = data.accounts.filter(a => a.type === "equity");
  const categories = [...new Set(equityAccounts.map(a => a.category || "Other Equity Accounts"))];
  if (!categories.includes("Retained Earnings")) categories.push("Retained Earnings");

  const open = {}, netIncomeRow = {}, other = {}, close = {};
  categories.forEach(cat => {
    const accts = equityAccounts.filter(a => (a.category || "Other Equity Accounts") === cat);
    const acctOpen = accts.reduce((s, a) => s + (openBal[a.id] || 0), 0);
    const acctMove = accts.reduce((s, a) => s + (mv[a.id] || 0), 0);
    if (cat === "Retained Earnings") {
      open[cat] = acctOpen + netIncomeAllTime(data, openStr);
      netIncomeRow[cat] = sumByType(data.accounts, mv, "revenue") - sumByType(data.accounts, mv, "expense");
      other[cat] = acctMove;
    } else {
      open[cat] = acctOpen;
      netIncomeRow[cat] = 0;
      other[cat] = acctMove;
    }
    close[cat] = open[cat] + netIncomeRow[cat] + other[cat];
  });
  const totalOf = (row) => categories.reduce((s, cat) => s + row[cat], 0);
  return { categories, open, netIncomeRow, other, close, totals: { open: totalOf(open), netIncomeRow: totalOf(netIncomeRow), other: totalOf(other), close: totalOf(close) } };
}
function EquityMatrix({ m, title }) {
  const { styles, fmt, theme } = useUI();
  return (
    <div style={{ marginBottom: 18 }}>
      {title && <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 8 }}>{title}</div>}
      <table style={styles.table}>
        <thead><tr><th style={styles.th}></th>{m.categories.map(cat => <th key={cat} style={{ ...styles.th, textAlign: "right" }}>{cat}</th>)}<th style={{ ...styles.th, textAlign: "right" }}>Total</th></tr></thead>
        <tbody>
          <tr><td style={{ ...styles.td, fontWeight: 600 }}>Balance at start of period</td>{m.categories.map(cat => <td key={cat} style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(m.open[cat])}</td>)}<td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(m.totals.open)}</td></tr>
          <tr><td style={styles.td}>Net income for the period</td>{m.categories.map(cat => <td key={cat} style={{ ...styles.tdMono, textAlign: "right" }}>{m.netIncomeRow[cat] ? fmt(m.netIncomeRow[cat]) : "-"}</td>)}<td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(m.totals.netIncomeRow)}</td></tr>
          <tr><td style={styles.td}>Other equity movements <span style={{ fontSize: 10.5, color: theme.muted }}>(contributions, drawings, transfers)</span></td>{m.categories.map(cat => <td key={cat} style={{ ...styles.tdMono, textAlign: "right" }}>{m.other[cat] ? fmt(m.other[cat]) : "-"}</td>)}<td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(m.totals.other)}</td></tr>
          <tr style={{ borderTop: `2px solid ${theme.border}` }}><td style={{ ...styles.td, fontWeight: 700 }}>Balance at end of period</td>{m.categories.map(cat => <td key={cat} style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(m.close[cat])}</td>)}<td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(m.totals.close)}</td></tr>
        </tbody>
      </table>
    </div>
  );
}
function ReportEquityMovement({ data, range, periodsBack }) {
  const { styles, theme } = useUI();
  const opts = reportOpts(data);
  const [view, setView] = useState("table");
  const n = opts.compareEnabled ? Math.max(1, periodsBack || 1) : 1;
  const trendRanges = computeTrendRanges(range, n);
  const mList = trendRanges.map(r => computeEquityMovement(data, r));
  const m = mList[mList.length - 1];
  const labels = trendRanges.map(periodLabel);
  if (n === 1) {
    return (
      <div>
        <div className="ces-card" style={styles.card}>
          <EquityMatrix m={m} title={null} />
          <div style={{ fontSize: 11, color: theme.muted }}>Components reflect the Chart of Accounts categories under Equity. Retained Earnings includes undistributed net income even without a dedicated account, since no formal year-end closing entry is posted in this system. Other Comprehensive Income (revaluation surpluses, FX translation reserves) is not yet tracked separately.</div>
        </div>
      </div>
    );
  }
  const closingList = mList.map(x => x.totals.close);
  const netIncomeList = mList.map(x => x.totals.netIncomeRow);
  return (
    <div>
      <TrendViewSwitch view={view} setView={setView} />
      <div className="ces-card" style={styles.card}>
        {view === "table" ? <>
          <TrendHeader labels={labels} />
          <TrendRow label="Net income for the period" values={netIncomeList} indent />
          <TrendRow label="Total equity at end of period" values={closingList} bold divider />
        </> : (
          <TrendChart labels={labels} colors={[theme.accent, theme.emerald]}
            series={[{ key: "eq", label: "Total equity", values: closingList }, { key: "ni", label: "Net income", values: netIncomeList }]} />
        )}
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 8 }}>Detail for {labels[labels.length - 1]}</div>
        <EquityMatrix m={m} title={null} />
        <div style={{ fontSize: 11, color: theme.muted }}>Components reflect the Chart of Accounts categories under Equity. Retained Earnings includes undistributed net income even without a dedicated account, since no formal year-end closing entry is posted in this system.</div>
      </div>
    </div>
  );
}
function ReportTrialBalance({ data, asOf }) {
  const { styles, fmt, theme, openDrilldown } = useUI();
  const bal = computeBalances(data.accounts, data.transactions, asOf);
  const rows = data.accounts.map(a => {
    const isDebitNormal = a.normal ? a.normal === "debit" : DEBIT_NORMAL[a.type];
    const raw = bal[a.id] || 0;
    const debit = isDebitNormal && raw > 0 ? raw : (!isDebitNormal && raw < 0 ? -raw : 0);
    const credit = !isDebitNormal && raw > 0 ? raw : (isDebitNormal && raw < 0 ? -raw : 0);
    return { a, debit, credit };
  }).filter(r => Math.abs(r.debit) + Math.abs(r.credit) > 0.005);
  const totalDebit = rows.reduce((s, r) => s + r.debit, 0), totalCredit = rows.reduce((s, r) => s + r.credit, 0);
  return (
    <div>
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Code</th><th style={styles.th}>Account</th><th style={{ ...styles.th, textAlign: "right" }}>Debit</th><th style={{ ...styles.th, textAlign: "right" }}>Credit</th></tr></thead>
          <tbody>{rows.map(r => <tr key={r.a.id} className="drill-row" style={{ cursor: "pointer" }} onClick={() => openDrilldown({ accountIds: [r.a.id], label: `${r.a.code} ${r.a.name}`, asOf })}><td style={styles.tdMono}>{r.a.code}</td><td style={styles.td}>{r.a.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.debit ? fmt(r.debit) : ""}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.credit ? fmt(r.credit) : ""}</td></tr>)}</tbody>
          <tfoot><tr><td colSpan={2} style={{ ...styles.td, fontWeight: 700 }}>Totals</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(totalDebit)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(totalCredit)}</td></tr></tfoot>
        </table>
        <div style={{ fontSize: 11, color: theme.muted, marginTop: 8, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{Math.abs(totalDebit - totalCredit) < 0.5 ? "✓ debits = credits" : "⚠ out of balance"}</div>
      </div>
    </div>
  );
}
function ReportGeneralLedger({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const activeAccounts = data.accounts.filter(a => data.transactions.some(t => t.date >= from && t.date <= to && t.lines.some(l => l.accountId === a.id)));
  return (
    <div>
      {activeAccounts.length === 0 && <div style={{ color: theme.muted, fontSize: 13 }}>No activity in this period.</div>}
      {activeAccounts.map(a => {
        const lines = data.transactions.filter(t => t.date >= from && t.date <= to).flatMap(t => t.lines.filter(l => l.accountId === a.id).map(l => ({ ...l, date: t.date, memo: t.memo }))).sort((x, y) => x.date.localeCompare(y.date));
        let running = 0;
        const isDebitNormal = a.normal ? a.normal === "debit" : DEBIT_NORMAL[a.type];
        return (
          <div key={a.id} style={styles.card}>
            <div style={styles.cardTitle}>{a.code} · {a.name}</div>
            <table style={styles.table}>
              <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={{ ...styles.th, textAlign: "right" }}>Debit</th><th style={{ ...styles.th, textAlign: "right" }}>Credit</th><th style={{ ...styles.th, textAlign: "right" }}>Balance</th></tr></thead>
              <tbody>{lines.map((l, i) => { running += isDebitNormal ? (l.debit - l.credit) : (l.credit - l.debit); return (
                <tr key={i}><td style={styles.tdMono}>{fmtDate(l.date)}</td><td style={styles.td}>{l.memo}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{l.debit ? fmt(l.debit) : ""}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{l.credit ? fmt(l.credit) : ""}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(running)}</td></tr>
              ); })}</tbody>
            </table>
          </div>
        );
      })}
    </div>
  );
}
function ReportAccountTransactions({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const [accountId, setAccountId] = useState(data.accounts[0]?.id);
  const { from, to } = rangeToDates(range);
  const account = data.accounts.find(a => a.id === accountId);
  const isDebitNormal = account ? (account.normal ? account.normal === "debit" : DEBIT_NORMAL[account.type]) : true;
  const lines = data.transactions.filter(t => t.date >= from && t.date <= to).flatMap(t => t.lines.filter(l => l.accountId === accountId).map(l => ({ ...l, date: t.date, memo: t.memo, source: t.source }))).sort((x, y) => x.date.localeCompare(y.date));
  let running = 0;
  return (
    <div>
      <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <select style={styles.input} value={accountId} onChange={e => setAccountId(e.target.value)}>{data.accounts.map(a => <option key={a.id} value={a.id}>{a.code} · {a.name}</option>)}</select>
      </div>
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={styles.th}>Source</th><th style={{ ...styles.th, textAlign: "right" }}>Debit</th><th style={{ ...styles.th, textAlign: "right" }}>Credit</th><th style={{ ...styles.th, textAlign: "right" }}>Balance</th></tr></thead>
          <tbody>{lines.map((l, i) => { running += isDebitNormal ? (l.debit - l.credit) : (l.credit - l.debit); return (
            <tr key={i}><td style={styles.tdMono}>{fmtDate(l.date)}</td><td style={styles.td}>{l.memo}</td><td style={{ ...styles.td, fontSize: 12, color: theme.muted }}>{l.source}</td>
              <td style={{ ...styles.tdMono, textAlign: "right" }}>{l.debit ? fmt(l.debit) : ""}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{l.credit ? fmt(l.credit) : ""}</td>
              <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(running)}</td></tr>
          ); })}</tbody>
        </table>
        {lines.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No activity for this account in range.</div>}
      </div>
    </div>
  );
}
function ReportAccountTypeSummary({ data, balances }) {
  const { styles } = useUI();
  const opts = reportOpts(data);
  return (
    <div className="ces-card" style={styles.card}>
      {ACCOUNT_TYPES.map(type => {
        const accts = data.accounts.filter(a => a.type === type);
        const subtypes = [...new Set(accts.map(a => a.subtype || "general"))];
        const total = sumByType(data.accounts, balances, type);
        const rows = subtypes.map(st => {
          const subAccts = accts.filter(a => (a.subtype || "general") === st);
          const subTotal = subAccts.reduce((s, a) => s + (a.contra ? -(balances[a.id] || 0) : (balances[a.id] || 0)), 0);
          return { st, subTotal };
        }).filter(r => keepLine(r.subTotal, opts));
        if (opts.hideZeroLines && rows.length === 0 && Math.abs(total) < 0.004) return null;
        return (
          <div key={type} style={{ marginBottom: 14 }}>
            <RowLine label={type.charAt(0).toUpperCase() + type.slice(1)} value={total} bold />
            {rows.map(r => <RowLine key={r.st} label={r.st} value={r.subTotal} indent tone="amber" />)}
          </div>
        );
      })}
    </div>
  );
}
function ReportAging({ data, kind }) {
  const { styles, fmt, theme } = useUI();
  const docs = kind === "ar" ? data.invoices : data.bills;
  const partyKey = kind === "ar" ? "customer" : "vendor";
  const rows = docs.map(d => {
    const total = computeDocTotals(d, data.taxGroups).finalAmount;
    const outstanding = total - (d.amountPaid || 0);
    if (outstanding <= 0.5) return null;
    const days = d.dueDate ? Math.floor((new Date(todayStr()) - new Date(d.dueDate)) / 86400000) : 0;
    return { id: d.id, party: d[partyKey], outstanding, bucket: agingBucket(days) };
  }).filter(Boolean);
  const byParty = {};
  rows.forEach(r => { if (!byParty[r.party]) byParty[r.party] = { "Current": 0, "1-30 days": 0, "31-60 days": 0, "61-90 days": 0, "90+ days": 0, total: 0 }; byParty[r.party][r.bucket] += r.outstanding; byParty[r.party].total += r.outstanding; });
  const parties = Object.keys(byParty);
  const grandTotal = rows.reduce((s, r) => s + r.outstanding, 0);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>{kind === "ar" ? "Customer" : "Vendor"}</th>{AGING_BUCKETS.map(b => <th key={b} style={{ ...styles.th, textAlign: "right" }}>{b}</th>)}<th style={{ ...styles.th, textAlign: "right" }}>Total</th></tr></thead>
        <tbody>{parties.map(p => (
          <tr key={p}><td style={styles.td}>{p}</td>{AGING_BUCKETS.map(b => <td key={b} style={{ ...styles.tdMono, textAlign: "right" }}>{byParty[p][b] ? fmt(byParty[p][b]) : "-"}</td>)}<td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(byParty[p].total)}</td></tr>
        ))}</tbody>
        {parties.length > 0 && <tfoot><tr><td style={{ ...styles.td, fontWeight: 700 }}>Total</td>{AGING_BUCKETS.map(b => <td key={b} style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(parties.reduce((s, p) => s + byParty[p][b], 0))}</td>)}<td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(grandTotal)}</td></tr></tfoot>}
      </table>
      {parties.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>Nothing outstanding - all {kind === "ar" ? "invoices" : "bills"} are settled.</div>}
    </div>
  );
}
function ReportSalesByItem({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.invoices.forEach(inv => inv.items.forEach(it => { const key = it.desc || "Unnamed"; if (!totals[key]) totals[key] = { qty: 0, revenue: 0 }; totals[key].qty += Number(it.qty || 0); totals[key].revenue += Number(it.qty || 0) * Number(it.price || 0); }));
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.revenue - a.revenue);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty sold</th><th style={{ ...styles.th, textAlign: "right" }}>Revenue</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.revenue)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No invoices yet.</div>}
    </div>
  );
}
function ReportSalesByCustomer({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.invoices.forEach(inv => { const t = computeDocTotals(inv, data.taxGroups).finalAmount; if (!totals[inv.customer]) totals[inv.customer] = { count: 0, revenue: 0 }; totals[inv.customer].count++; totals[inv.customer].revenue += t; });
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.revenue - a.revenue);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Customer</th><th style={{ ...styles.th, textAlign: "right" }}>Invoices</th><th style={{ ...styles.th, textAlign: "right" }}>Total billed</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.count}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.revenue)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No invoices yet.</div>}
    </div>
  );
}
function ReportRefundHistory({ data }) {
  const { styles, fmt, theme } = useUI();
  const refunds = data.payments.filter(p => p.type === "refund_to_customer" || p.type === "refund_from_vendor").sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Direction</th><th style={styles.th}>Party</th><th style={styles.th}>Bank</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
        <tbody>{refunds.map(p => (
          <tr key={p.id}><td style={styles.tdMono}>{fmtDate(p.date)}</td><td style={styles.td}>{p.type === "refund_to_customer" ? "To customer" : "From vendor"}</td><td style={styles.td}>{p.refundTo || "-"}</td><td style={styles.td}>{data.banks.find(b => b.id === p.bankId)?.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(p.amount)}</td></tr>
        ))}</tbody>
      </table>
      {refunds.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No refunds recorded yet - record one from the Payments tab.</div>}
    </div>
  );
}
function ReportOrderFulfillment({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.salesOrders.forEach(so => so.items.forEach(it => {
    if (!it.inventoryId) return;
    const inv = data.inventory.find(x => x.id === it.inventoryId);
    const key = inv ? inv.name : it.desc || "Unnamed";
    if (!totals[key]) totals[key] = { ordered: 0, fulfilled: 0 };
    totals[key].ordered += Number(it.qty || 0);
    if (so.status === "fulfilled") totals[key].fulfilled += Number(it.qty || 0);
  }));
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v, open: v.ordered - v.fulfilled })).sort((a, b) => b.ordered - a.ordered);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Ordered</th><th style={{ ...styles.th, textAlign: "right" }}>Fulfilled</th><th style={{ ...styles.th, textAlign: "right" }}>Open</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.ordered}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.fulfilled}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.open > 0 ? theme.amber : theme.text }}>{r.open}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No inventory-linked sales orders yet.</div>}
    </div>
  );
}
function ReportSalesReturnHistory({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = [...data.salesReturns].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Return</th><th style={styles.th}>Customer</th><th style={styles.th}>Applied to</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.tdMono}>{fmtDate(r.date)}</td><td style={styles.tdMono}>{r.id}</td><td style={styles.td}>{r.customer}</td><td style={styles.td}>{r.invoiceId || (r.refundBankId ? "Refunded" : "Open credit")}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(computeDocTotals(r, data.taxGroups).finalAmount)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No sales returns recorded yet.</div>}
    </div>
  );
}
function ReportSalesBySalesperson({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.invoices.forEach(inv => { const key = inv.salesperson || "Unassigned"; const t = computeDocTotals(inv, data.taxGroups).finalAmount; if (!totals[key]) totals[key] = { count: 0, revenue: 0 }; totals[key].count++; totals[key].revenue += t; });
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.revenue - a.revenue);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Salesperson</th><th style={{ ...styles.th, textAlign: "right" }}>Invoices</th><th style={{ ...styles.th, textAlign: "right" }}>Revenue</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.count}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.revenue)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No invoices yet. Add a salesperson name when creating an invoice to track this.</div>}
    </div>
  );
}
function ReportProfitByItem({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.invoices.forEach(inv => inv.items.forEach(it => {
    if (!it.inventoryId) return;
    const invItem = data.inventory.find(x => x.id === it.inventoryId);
    const key = invItem ? invItem.name : it.desc || "Unnamed";
    if (!totals[key]) totals[key] = { qty: 0, revenue: 0, cost: 0 };
    totals[key].qty += Number(it.qty || 0);
    totals[key].revenue += Number(it.qty || 0) * Number(it.price || 0);
    totals[key].cost += Number(it.qty || 0) * (invItem ? invItem.unitCost : 0);
  }));
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v, profit: v.revenue - v.cost, margin: v.revenue ? (v.revenue - v.cost) / v.revenue : 0 })).sort((a, b) => b.profit - a.profit);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty sold</th><th style={{ ...styles.th, textAlign: "right" }}>Revenue</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={{ ...styles.th, textAlign: "right" }}>Profit</th><th style={{ ...styles.th, textAlign: "right" }}>Margin</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.revenue)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.cost)}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.profit >= 0 ? theme.emerald : theme.rose, fontWeight: 600 }}>{fmt(r.profit)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{(r.margin * 100).toFixed(1)}%</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No inventory-linked sales yet.</div>}
      <div style={{ fontSize: 11, color: theme.muted, marginTop: 10 }}>Cost uses each item's current unit cost, since the app doesn't retain a per-line historical cost snapshot separately from the ledger - for items whose cost has changed recently, this is an approximation of past sales.</div>
    </div>
  );
}
function ReportCommittedStock({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = data.inventory.map(item => {
    const committed = data.salesOrders.filter(so => so.status === "open").reduce((s, so) => s + so.items.filter(it => it.inventoryId === item.id).reduce((s2, it) => s2 + Number(it.qty || 0), 0), 0);
    return { ...item, committed, available: item.qty - committed };
  }).filter(r => r.qty > 0 || r.committed > 0);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>On hand</th><th style={{ ...styles.th, textAlign: "right" }}>Committed</th><th style={{ ...styles.th, textAlign: "right" }}>Available</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.committed > 0 ? theme.amber : theme.text }}>{r.committed}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.available < 0 ? theme.rose : theme.text, fontWeight: 600 }}>{r.available}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No inventory items yet.</div>}
      <div style={{ fontSize: 11, color: theme.muted, marginTop: 10 }}>Committed = quantity on open (not yet fulfilled) sales orders. Available = on hand minus committed.</div>
    </div>
  );
}
function ReportInventoryMovementHistory({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const rows = data.inventoryLots.filter(l => l.date >= from && l.date <= to).map(l => ({ ...l, item: data.inventory.find(i => i.id === l.itemId) })).sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty received</th><th style={{ ...styles.th, textAlign: "right" }}>Remaining</th><th style={{ ...styles.th, textAlign: "right" }}>Unit cost</th><th style={styles.th}>Source</th></tr></thead>
        <tbody>{rows.map(l => <tr key={l.id}><td style={styles.tdMono}>{fmtDate(l.date)}</td><td style={styles.td}>{l.item ? l.item.name : l.itemId}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{l.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{l.remainingQty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(l.unitCost)}</td><td style={styles.td}>{l.sourceDocId || "-"}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No stock received in this period.</div>}
    </div>
  );
}
function ReportInventoryTurnoverByItem({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const rows = data.inventory.map(item => {
    const qtySold = data.invoices.filter(i => i.date >= from && i.date <= to).reduce((s, inv) => s + inv.items.filter(it => it.inventoryId === item.id).reduce((s2, it) => s2 + Number(it.qty || 0), 0), 0);
    const cogs = qtySold * item.unitCost;
    const avgInventoryValue = item.qty * item.unitCost; // current value used as the average, since historical per-item balances aren't separately retained
    const turnover = avgInventoryValue > 0 ? cogs / avgInventoryValue : null;
    return { ...item, qtySold, cogs, avgInventoryValue, turnover };
  }).filter(r => r.qtySold > 0 || r.qty > 0);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty sold</th><th style={{ ...styles.th, textAlign: "right" }}>COGS</th><th style={{ ...styles.th, textAlign: "right" }}>Avg. inventory value</th><th style={{ ...styles.th, textAlign: "right" }}>Turnover</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qtySold}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.cogs)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.avgInventoryValue)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{r.turnover === null ? "-" : r.turnover.toFixed(2) + "x"}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No inventory activity yet.</div>}
      <div style={{ fontSize: 11, color: theme.muted, marginTop: 10 }}>Turnover = cost of goods sold in the period ÷ current inventory value, used as a stand-in for average holding since per-item balances aren't tracked at every prior date.</div>
    </div>
  );
}
function ReportInventoryTurnoverByAmount({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const dayBefore = new Date(from); dayBefore.setDate(dayBefore.getDate() - 1);
  const openBal = computeBalances(data.accounts, data.transactions, localDateToISO(dayBefore));
  const closeBal = computeBalances(data.accounts, data.transactions, to);
  const mv = periodMovement(data.accounts, data.transactions, from, to);
  const cogs = sumByType(data.accounts, mv, "expense", "cogs");
  const invAccounts = ["1200", "1201", "1202", "1203"];
  const openInv = invAccounts.reduce((s, id) => s + (openBal[id] || 0), 0);
  const closeInv = invAccounts.reduce((s, id) => s + (closeBal[id] || 0), 0);
  const avgInventory = (openInv + closeInv) / 2;
  const turnover = avgInventory > 0 ? cogs / avgInventory : null;
  const daysInPeriod = Math.round((new Date(to) - new Date(from)) / 86400000) + 1;
  const daysOnHand = turnover ? daysInPeriod / turnover : null;
  return (
    <div className="ces-card" style={styles.card}>
      <RowLine label="Cost of goods sold (period)" value={cogs} />
      <RowLine label="Opening inventory value" value={openInv} indent />
      <RowLine label="Closing inventory value" value={closeInv} indent />
      <RowLine label="Average inventory value" value={avgInventory} bold divider />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginTop: 14 }}>
        <Kpi label="Inventory turnover" value={turnover === null ? "-" : turnover.toFixed(2) + "x"} />
        <Kpi label="Avg. days inventory held" value={daysOnHand === null ? "-" : Math.round(daysOnHand) + " days"} tone="amber" />
      </div>
    </div>
  );
}
function ReportExpensesByCategory({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const mv = periodMovement(data.accounts, data.transactions, from, to);
  const opts = reportOpts(data);
  const expense = data.accounts.filter(a => a.type === "expense").map(a => ({ name: acctLabel(a, opts), value: a.contra ? -(mv[a.id] || 0) : (mv[a.id] || 0) })).filter(c => Math.abs(c.value) > 0.5).sort((a, b) => b.value - a.value);
  const total = expense.reduce((s, c) => s + c.value, 0);
  return (
    <div>
      <div className="ces-card" style={styles.card}>
        {expense.length > 0 && <ResponsiveContainer width="100%" height={Math.max(160, expense.length * 32)}>
          <BarChart data={expense} layout="vertical" margin={{ left: 20 }}>
            <CartesianGrid strokeDasharray="2 4" stroke={theme.border} horizontal={false} />
            <XAxis type="number" tick={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, fill: theme.muted }} axisLine={false} tickLine={false} />
            <YAxis type="category" dataKey="name" width={160} tick={{ fontFamily: "Inter", fontSize: 12, fill: theme.text }} axisLine={false} tickLine={false} />
            <Tooltip formatter={v => fmt(v)} contentStyle={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 12, background: theme.panel }} />
            <Bar dataKey="value" fill={theme.amber} radius={[0, 4, 4, 0]} />
          </BarChart>
        </ResponsiveContainer>}
        <div style={{ marginTop: 12 }}>{expense.map(c => <RowLine key={c.name} label={c.name} value={c.value} />)}<RowLine label="Total" value={total} bold divider /></div>
        {expense.length === 0 && <div style={{ color: theme.muted, fontSize: 13 }}>No expenses in this period.</div>}
      </div>
    </div>
  );
}
function ReportCustomerBalanceSummary({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.invoices.forEach(inv => {
    const t = computeDocTotals(inv, data.taxGroups).finalAmount;
    if (!totals[inv.customer]) totals[inv.customer] = { invoiced: 0, paid: 0 };
    totals[inv.customer].invoiced += t;
    totals[inv.customer].paid += inv.amountPaid || 0;
  });
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v, balance: v.invoiced - v.paid })).filter(r => Math.abs(r.balance) > 0.5).sort((a, b) => b.balance - a.balance);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Customer</th><th style={{ ...styles.th, textAlign: "right" }}>Invoiced</th><th style={{ ...styles.th, textAlign: "right" }}>Paid</th><th style={{ ...styles.th, textAlign: "right" }}>Balance</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.invoiced)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.paid)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.balance)}</td></tr>)}</tbody>
        <tfoot><tr><td style={{ ...styles.td, fontWeight: 700 }}>Total</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(rows.reduce((s, r) => s + r.invoiced, 0))}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(rows.reduce((s, r) => s + r.paid, 0))}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(rows.reduce((s, r) => s + r.balance, 0))}</td></tr></tfoot>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No outstanding customer balances.</div>}
    </div>
  );
}
function ReportReceivablesDetails({ data }) {
  const { styles, fmt, theme } = useUI();
  const today = todayStr();
  const rows = data.invoices.map(inv => {
    const total = computeDocTotals(inv, data.taxGroups).finalAmount;
    const balance = total - (inv.amountPaid || 0);
    const daysOverdue = inv.dueDate ? Math.floor((new Date(today) - new Date(inv.dueDate)) / 86400000) : 0;
    return { ...inv, total, balance, daysOverdue };
  }).filter(r => r.balance > 0.5).sort((a, b) => b.daysOverdue - a.daysOverdue);
  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ overflowX: "auto" }}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Invoice</th><th style={styles.th}>Customer</th><th style={styles.th}>Date</th><th style={styles.th}>Due date</th><th style={{ ...styles.th, textAlign: "right" }}>Days overdue</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th><th style={{ ...styles.th, textAlign: "right" }}>Balance</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.tdMono}>{r.id}</td><td style={styles.td}>{r.customer}</td><td style={styles.tdMono}>{fmtDate(r.date)}</td><td style={styles.tdMono}>{r.dueDate ? fmtDate(r.dueDate) : "-"}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.daysOverdue > 0 ? theme.rose : theme.text }}>{r.daysOverdue > 0 ? r.daysOverdue : "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.total)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.balance)}</td></tr>)}</tbody>
      </table>
      </div>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>Nothing outstanding.</div>}
    </div>
  );
}
function ReportPaymentsReceived({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const rows = data.payments.filter(p => p.type === "received" && p.date >= from && p.date <= to).sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Invoice</th><th style={styles.th}>Bank</th><th style={styles.th}>Memo</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
        <tbody>{rows.map(p => <tr key={p.id}><td style={styles.tdMono}>{fmtDate(p.date)}</td><td style={styles.td}>{p.relatedId || "-"}</td><td style={styles.td}>{data.banks.find(b => b.id === p.bankId)?.name || "-"}</td><td style={styles.td}>{p.memo || "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(p.amount)}</td></tr>)}</tbody>
        <tfoot><tr><td colSpan={4} style={{ ...styles.td, fontWeight: 700 }}>Total</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(rows.reduce((s, p) => s + p.amount, 0))}</td></tr></tfoot>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No payments received in this period.</div>}
    </div>
  );
}
function ReportTimeToGetPaid({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = data.invoices.filter(inv => inv.status === "paid").map(inv => {
    const relatedPayments = data.payments.filter(p => p.relatedType === "invoice" && p.relatedId === inv.id);
    const paidDate = relatedPayments.length ? relatedPayments.map(p => p.date).sort().slice(-1)[0] : null;
    const days = paidDate ? Math.round((new Date(paidDate) - new Date(inv.date)) / 86400000) : null;
    return { ...inv, paidDate, days };
  }).filter(r => r.days !== null);
  const avg = rows.length ? rows.reduce((s, r) => s + r.days, 0) / rows.length : null;
  const byCustomer = {};
  rows.forEach(r => { if (!byCustomer[r.customer]) byCustomer[r.customer] = []; byCustomer[r.customer].push(r.days); });
  const customerRows = Object.entries(byCustomer).map(([name, arr]) => ({ name, avgDays: arr.reduce((s, v) => s + v, 0) / arr.length, count: arr.length })).sort((a, b) => b.avgDays - a.avgDays);
  return (
    <div>
      <div className="ces-card" style={styles.card}>
        <Kpi label="Average days to get paid" value={avg === null ? "-" : avg.toFixed(1) + " days"} tone={avg !== null && avg > 30 ? "amber" : "emerald"} />
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>By customer</div>
        <table style={{ ...styles.table, marginTop: 10 }}>
          <thead><tr><th style={styles.th}>Customer</th><th style={{ ...styles.th, textAlign: "right" }}>Invoices</th><th style={{ ...styles.th, textAlign: "right" }}>Avg. days to pay</th></tr></thead>
          <tbody>{customerRows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.count}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.avgDays.toFixed(1)}</td></tr>)}</tbody>
        </table>
        {customerRows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No fully paid invoices with a matched payment yet.</div>}
      </div>
    </div>
  );
}
function ReportCreditNoteDetails({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = [...data.creditNotes].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Credit note</th><th style={styles.th}>Customer</th><th style={styles.th}>Reason</th><th style={styles.th}>Applied to</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.tdMono}>{fmtDate(r.date)}</td><td style={styles.tdMono}>{r.id}</td><td style={styles.td}>{r.customer}</td><td style={styles.td}>{r.reason || "-"}</td><td style={styles.td}>{r.invoiceId || (r.refundBankId ? "Refunded" : "Open credit")}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(computeDocTotals(r, data.taxGroups).finalAmount)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No credit notes issued yet.</div>}
    </div>
  );
}
function ReportVendorBalanceSummary({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.bills.forEach(bill => {
    const t = computeDocTotals(bill, data.taxGroups).finalAmount;
    if (!totals[bill.vendor]) totals[bill.vendor] = { billed: 0, paid: 0 };
    totals[bill.vendor].billed += t;
    totals[bill.vendor].paid += bill.amountPaid || 0;
  });
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v, balance: v.billed - v.paid })).filter(r => Math.abs(r.balance) > 0.5).sort((a, b) => b.balance - a.balance);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Vendor</th><th style={{ ...styles.th, textAlign: "right" }}>Billed</th><th style={{ ...styles.th, textAlign: "right" }}>Paid</th><th style={{ ...styles.th, textAlign: "right" }}>Balance</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.billed)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.paid)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.balance)}</td></tr>)}</tbody>
        <tfoot><tr><td style={{ ...styles.td, fontWeight: 700 }}>Total</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(rows.reduce((s, r) => s + r.billed, 0))}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(rows.reduce((s, r) => s + r.paid, 0))}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(rows.reduce((s, r) => s + r.balance, 0))}</td></tr></tfoot>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No outstanding vendor balances.</div>}
    </div>
  );
}
function ReportRecurringTransactions({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = [...data.recurringJournals].sort((a, b) => (a.nextDate || "").localeCompare(b.nextDate || ""));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Memo</th><th style={styles.th}>Frequency</th><th style={styles.th}>Next due</th><th style={styles.th}>End date</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}>Status</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.td}>{r.memo}{r.kind ? <span style={{ fontSize: 10.5, color: theme.muted, marginLeft: 6 }}>({r.kind})</span> : null}</td><td style={styles.td}>{r.frequency}</td><td style={styles.tdMono}>{r.nextDate ? fmtDate(r.nextDate) : "-"}</td><td style={styles.tdMono}>{r.endDate ? fmtDate(r.endDate) : "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.lines.reduce((s, l) => s + l.debit, 0))}</td><td style={styles.td}><span style={{ ...styles.pill, ...(r.active ? styles.pillGreen : styles.pillAmber) }}>{r.active ? "active" : "paused"}</span></td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No recurring schedules yet - set one up from the Journal or Recurring Expenses.</div>}
    </div>
  );
}
function ReportPurchaseByVendor({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.bills.forEach(bill => { const t = computeDocTotals(bill, data.taxGroups).finalAmount; if (!totals[bill.vendor]) totals[bill.vendor] = { count: 0, spend: 0 }; totals[bill.vendor].count++; totals[bill.vendor].spend += t; });
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.spend - a.spend);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Vendor</th><th style={{ ...styles.th, textAlign: "right" }}>Bills</th><th style={{ ...styles.th, textAlign: "right" }}>Total spend</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.count}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.spend)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No bills yet.</div>}
    </div>
  );
}
function ReportPurchaseByItem({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.bills.forEach(bill => bill.items.forEach(it => { const key = it.desc || "Unnamed"; if (!totals[key]) totals[key] = { qty: 0, spend: 0 }; totals[key].qty += Number(it.qty || 0); totals[key].spend += Number(it.qty || 0) * Number(it.price || 0); }));
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.spend - a.spend);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty purchased</th><th style={{ ...styles.th, textAlign: "right" }}>Spend</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.spend)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No bills yet.</div>}
    </div>
  );
}
function ReportGoodsReceivedHistory({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = [...data.purchaseReceipts].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Receipt</th><th style={styles.th}>Vendor</th><th style={styles.th}>Status</th><th style={{ ...styles.th, textAlign: "right" }}>Total</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.tdMono}>{fmtDate(r.date)}</td><td style={styles.tdMono}>{r.id}</td><td style={styles.td}>{r.vendor}</td><td style={styles.td}><span style={{ ...styles.pill, ...(r.status === "billed" ? styles.pillGreen : styles.pillAmber) }}>{r.status === "billed" ? "Billed" : "Awaiting bill"}</span></td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.items.reduce((s, it) => s + Number(it.qty || 0) * Number(it.price || 0), 0))}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No purchase receipts yet.</div>}
    </div>
  );
}
function ReportGoodsReceivedByItem({ data }) {
  const { styles, fmt, theme } = useUI();
  const totals = {};
  data.purchaseReceipts.forEach(r => r.items.forEach(it => { const inv = data.inventory.find(x => x.id === it.inventoryId); const key = inv ? inv.name : it.desc || "Unnamed"; if (!totals[key]) totals[key] = { qty: 0, value: 0 }; totals[key].qty += Number(it.qty || 0); totals[key].value += Number(it.qty || 0) * Number(it.price || 0); }));
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.value - a.value);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty received</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.value)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No purchase receipts yet.</div>}
    </div>
  );
}
function ReportFixedAssetsByLocation({ data }) {
  const { styles, fmt, theme } = useUI();
  const groups = {};
  data.fixedAssets.forEach(a => {
    const key = data.locations.find(l => l.id === a.locationId)?.name || "Unassigned";
    if (!groups[key]) groups[key] = { count: 0, cost: 0, nbv: 0 };
    groups[key].count++;
    groups[key].cost += a.cost;
    groups[key].nbv += a.cost - a.accumulatedDepreciation;
  });
  const rows = Object.entries(groups).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.nbv - a.nbv);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Location</th><th style={{ ...styles.th, textAlign: "right" }}>Assets</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={{ ...styles.th, textAlign: "right" }}>Net book value</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.count}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.cost)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.nbv)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No fixed assets yet.</div>}
    </div>
  );
}
function ReportFixedAssetsByDepartment({ data }) {
  const { styles, fmt, theme } = useUI();
  const groups = {};
  data.fixedAssets.forEach(a => {
    const key = data.departments.find(x => x.id === a.departmentId)?.name || "Unassigned";
    if (!groups[key]) groups[key] = { count: 0, cost: 0, nbv: 0 };
    groups[key].count++;
    groups[key].cost += a.cost;
    groups[key].nbv += a.cost - a.accumulatedDepreciation;
  });
  const rows = Object.entries(groups).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.nbv - a.nbv);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Department</th><th style={{ ...styles.th, textAlign: "right" }}>Assets</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={{ ...styles.th, textAlign: "right" }}>Net book value</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.count}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.cost)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.nbv)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No fixed assets yet.</div>}
    </div>
  );
}
function ReportInventoryByLocation({ data }) {
  const { styles, fmt, theme } = useUI();
  const groups = {};
  data.inventory.forEach(i => {
    const key = data.locations.find(l => l.id === i.locationId)?.name || "Unassigned";
    if (!groups[key]) groups[key] = { items: 0, qty: 0, value: 0 };
    groups[key].items++;
    groups[key].qty += i.qty;
    groups[key].value += i.qty * i.unitCost;
  });
  const rows = Object.entries(groups).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.value - a.value);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Location</th><th style={{ ...styles.th, textAlign: "right" }}>Items</th><th style={{ ...styles.th, textAlign: "right" }}>Total qty</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.items}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.value)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No inventory items yet.</div>}
    </div>
  );
}
function ReportInventoryByDepartment({ data }) {
  const { styles, fmt, theme } = useUI();
  const groups = {};
  data.inventory.forEach(i => {
    const key = data.departments.find(x => x.id === i.departmentId)?.name || "Unassigned";
    if (!groups[key]) groups[key] = { items: 0, qty: 0, value: 0 };
    groups[key].items++;
    groups[key].qty += i.qty;
    groups[key].value += i.qty * i.unitCost;
  });
  const rows = Object.entries(groups).map(([name, v]) => ({ name, ...v })).sort((a, b) => b.value - a.value);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Department</th><th style={{ ...styles.th, textAlign: "right" }}>Items</th><th style={{ ...styles.th, textAlign: "right" }}>Total qty</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.items}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{r.qty}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.value)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No inventory items yet.</div>}
    </div>
  );
}
function ReportInventorySummary({ data }) {
  const { styles, fmt, theme } = useUI();
  const totalValue = data.inventory.reduce((s, i) => s + i.qty * i.unitCost, 0);
  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ fontSize: 13, color: theme.muted, marginBottom: 10 }}>Total inventory value: <strong style={{ color: theme.text }}>{fmt(totalValue)}</strong></div>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>SKU</th><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>On hand</th><th style={{ ...styles.th, textAlign: "right" }}>Unit cost</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th></tr></thead>
        <tbody>{data.inventory.map(i => (<tr key={i.id}><td style={styles.tdMono}>{i.sku}</td><td style={styles.td}>{i.name}</td><td style={{ ...styles.tdMono, textAlign: "right", color: i.qty <= i.reorderLevel ? theme.rose : theme.text }}>{i.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.unitCost)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.qty * i.unitCost)}</td></tr>))}</tbody>
      </table>
    </div>
  );
}
function ReportInventoryAging({ data }) {
  const { styles, fmt, theme } = useUI();
  const openLots = data.inventoryLots.filter(l => l.remainingQty > 0.001).sort((a, b) => a.date.localeCompare(b.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={styles.th}>Received</th><th style={{ ...styles.th, textAlign: "right" }}>Days held</th><th style={{ ...styles.th, textAlign: "right" }}>Qty remaining</th><th style={{ ...styles.th, textAlign: "right" }}>Unit cost</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th></tr></thead>
        <tbody>{openLots.map(l => { const item = data.inventory.find(i => i.id === l.itemId); const days = Math.floor((new Date(todayStr()) - new Date(l.date)) / 86400000);
          return <tr key={l.id}><td style={styles.td}>{item?.name || l.itemId}</td><td style={styles.tdMono}>{fmtDate(l.date)}</td><td style={{ ...styles.tdMono, textAlign: "right", color: days > 90 ? theme.rose : theme.text }}>{days}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{l.remainingQty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(l.unitCost)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(l.remainingQty * l.unitCost)}</td></tr>; })}</tbody>
      </table>
      {openLots.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No open stock lots.</div>}
    </div>
  );
}
function ReportFIFO({ data }) {
  const { styles, fmt, theme } = useUI();
  return (
    <div>
      {data.inventory.map(item => {
        const lots = data.inventoryLots.filter(l => l.itemId === item.id).sort((a, b) => a.date.localeCompare(b.date));
        if (lots.length === 0) return null;
        return (
          <div key={item.id} style={styles.card}>
            <div style={styles.cardTitle}>{item.name} <span style={{ fontWeight: 400, fontSize: 12, color: theme.muted }}>({item.sku})</span></div>
            <table style={styles.table}>
              <thead><tr><th style={styles.th}>Lot received</th><th style={{ ...styles.th, textAlign: "right" }}>Qty received</th><th style={{ ...styles.th, textAlign: "right" }}>Remaining (FIFO)</th><th style={{ ...styles.th, textAlign: "right" }}>Unit cost</th></tr></thead>
              <tbody>{lots.map(l => <tr key={l.id}><td style={styles.tdMono}>{fmtDate(l.date)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{l.qty}</td><td style={{ ...styles.tdMono, textAlign: "right", color: l.remainingQty < l.qty ? theme.amber : theme.text }}>{l.remainingQty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(l.unitCost)}</td></tr>)}</tbody>
            </table>
          </div>
        );
      })}
      <div style={{ fontSize: 12, color: theme.muted }}>This is an informational FIFO view of purchase lots and consumption order. The general ledger itself values inventory and COGS using weighted-average costing.</div>
    </div>
  );
}
function ReportWeightedAvg({ data }) {
  const { styles, fmt } = useUI();
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Qty on hand</th><th style={{ ...styles.th, textAlign: "right" }}>Weighted avg. unit cost</th><th style={{ ...styles.th, textAlign: "right" }}>Total value</th></tr></thead>
        <tbody>{data.inventory.map(i => <tr key={i.id}><td style={styles.td}>{i.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{i.qty}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.unitCost)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.qty * i.unitCost)}</td></tr>)}</tbody>
      </table>
    </div>
  );
}
function ReportABC({ data }) {
  const { styles, fmt, theme } = useUI();
  const items = data.inventory.map(i => ({ ...i, value: i.qty * i.unitCost })).sort((a, b) => b.value - a.value);
  const totalValue = items.reduce((s, i) => s + i.value, 0);
  let cum = 0;
  const rows = items.map(i => { cum += i.value; const cumPct = totalValue ? cum / totalValue : 0; const cls = cumPct <= 0.8 ? "A" : cumPct <= 0.95 ? "B" : "C"; return { ...i, cumPct, cls }; });
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Item</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th><th style={{ ...styles.th, textAlign: "right" }}>Cumulative %</th><th style={styles.th}>Class</th></tr></thead>
        <tbody>{rows.map(r => (<tr key={r.id}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.value)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{(r.cumPct * 100).toFixed(1)}%</td><td style={styles.td}><span style={{ ...styles.pill, ...(r.cls === "A" ? styles.pillGreen : r.cls === "B" ? styles.pillAmber : styles.pillRose) }}>{r.cls}</span></td></tr>))}</tbody>
      </table>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>A = top 80% of inventory value, B = next 15%, C = remaining 5%.</div>
    </div>
  );
}
function ReportTaxSummary({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const totals = {};
  const add = (name, effect, amount, accountId) => { if (!totals[name]) totals[name] = { added: 0, deducted: 0, accountId }; if (effect === "deduct") totals[name].deducted += amount; else totals[name].added += amount; };
  [...data.invoices, ...data.bills].filter(d => d.date >= from && d.date <= to).forEach(d => {
    const t = computeDocTotals(d, data.taxGroups);
    t.lineCalcs.forEach(l => l.components.forEach(c => add(c.name, c.effect, c.amount, c.accountId)));
    t.cascadeSteps.forEach(s => add(s.name, s.effect, s.amount, s.accountId));
  });
  const rows = Object.entries(totals).map(([name, v]) => ({ name, ...v, net: v.added - v.deducted }));
  return (
    <div>
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Tax</th><th style={{ ...styles.th, textAlign: "right" }}>Added (charged)</th><th style={{ ...styles.th, textAlign: "right" }}>Deducted (withheld)</th><th style={{ ...styles.th, textAlign: "right" }}>Net</th><th style={styles.th}>GL account</th></tr></thead>
          <tbody>{rows.map(r => (<tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.added)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.deducted)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.net)}</td><td style={styles.td}>{data.accounts.find(a => a.id === r.accountId)?.name}</td></tr>))}</tbody>
        </table>
        {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No taxes recorded in this range.</div>}
      </div>
    </div>
  );
}
function ReportExpensesByProject({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const totals = {};
  data.expenses.filter(e => e.date >= from && e.date <= to).forEach(e => { const key = e.projectId ? (data.projects.find(p => p.id === e.projectId)?.name || "Unknown project") : "No project"; if (!totals[key]) totals[key] = 0; totals[key] += e.amount; });
  const rows = Object.entries(totals).map(([name, amount]) => ({ name, amount })).sort((a, b) => b.amount - a.amount);
  const total = rows.reduce((s, r) => s + r.amount, 0);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Project</th><th style={{ ...styles.th, textAlign: "right" }}>Expenses</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.amount)}</td></tr>)}</tbody>
        <tfoot><tr><td style={{ ...styles.td, fontWeight: 700 }}>Total</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(total)}</td></tr></tfoot>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No expenses in this period.</div>}
    </div>
  );
}
function ReportReconciliationStatus({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = [...data.reconciliations].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Bank</th><th style={{ ...styles.th, textAlign: "right" }}>Statement balance</th><th style={{ ...styles.th, textAlign: "right" }}>Cleared items</th><th style={styles.th}>Status</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.tdMono}>{fmtDate(r.date)}</td><td style={styles.td}>{data.banks.find(b => b.id === r.bankId)?.name || "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.statementBalance)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{(r.clearedTxnIds || []).length}</td><td style={styles.td}><span style={{ ...styles.pill, ...(r.completed ? styles.pillGreen : styles.pillAmber) }}>{r.completed ? "Completed" : "In progress"}</span></td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No reconciliations started yet - begin one from a bank's Reconcile tab.</div>}
    </div>
  );
}
function ReportTimesheetDetails({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const rows = data.timesheets.filter(ts => ts.date >= from && ts.date <= to).sort((a, b) => b.date.localeCompare(a.date));
  const totalHours = rows.reduce((s, ts) => s + ts.hours, 0);
  const totalValue = rows.filter(ts => ts.billable).reduce((s, ts) => s + ts.hours * ts.rate, 0);
  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ overflowX: "auto" }}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Employee</th><th style={styles.th}>Project</th><th style={styles.th}>Description</th><th style={{ ...styles.th, textAlign: "right" }}>Hours</th><th style={styles.th}>Billable</th><th style={styles.th}>Status</th></tr></thead>
        <tbody>{rows.map(ts => (
          <tr key={ts.id}><td style={styles.tdMono}>{fmtDate(ts.date)}</td>
            <td style={styles.td}>{ts.employeeId ? (data.employees.find(e => e.id === ts.employeeId)?.name || "-") : ts.employeeName || "-"}</td>
            <td style={styles.td}>{ts.projectId ? (data.projects.find(p => p.id === ts.projectId)?.name || "-") : "-"}</td>
            <td style={styles.td}>{ts.description || "-"}</td>
            <td style={{ ...styles.tdMono, textAlign: "right" }}>{ts.hours}</td>
            <td style={styles.td}>{ts.billable ? fmt(ts.rate) + "/hr" : "-"}</td>
            <td style={styles.td}>{!ts.billable ? "-" : ts.invoiced ? "Invoiced" : "Unbilled"}</td></tr>
        ))}</tbody>
        <tfoot><tr><td colSpan={4} style={{ ...styles.td, fontWeight: 700 }}>Total</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{totalHours.toFixed(1)}</td><td colSpan={2} style={{ ...styles.tdMono, fontWeight: 700 }}>{fmt(totalValue)} billable</td></tr></tfoot>
      </table>
      </div>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No time logged in this period.</div>}
    </div>
  );
}
function ReportProjectDetails({ data }) {
  const { styles, fmt, theme } = useUI();
  const [projectId, setProjectId] = useState(data.projects[0]?.id || "");
  const project = data.projects.find(p => p.id === projectId);
  if (!project) return <div className="ces-card" style={styles.card}><div style={{ color: theme.muted, fontSize: 13 }}>No projects yet - add one in Settings.</div></div>;
  const invoices = data.invoices.filter(i => i.projectId === projectId);
  const bills = data.bills.filter(b => b.projectId === projectId);
  const expenses = data.expenses.filter(e => e.projectId === projectId);
  const timesheets = data.timesheets.filter(ts => ts.projectId === projectId);
  return (
    <div>
      <div className="no-print" style={{ marginBottom: 14 }}>
        <select style={styles.input} value={projectId} onChange={e => setProjectId(e.target.value)}>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Invoices ({invoices.length})</div>
        <table style={{ ...styles.table, marginTop: 8 }}><tbody>{invoices.map(i => <tr key={i.id}><td style={styles.tdMono}>{i.id}</td><td style={styles.td}>{fmtDate(i.date)}</td><td style={styles.td}>{i.customer}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(computeDocTotals(i, data.taxGroups).finalAmount)}</td></tr>)}</tbody></table>
        {invoices.length === 0 && <div style={{ fontSize: 12, color: theme.muted, marginTop: 6 }}>None.</div>}
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Bills ({bills.length})</div>
        <table style={{ ...styles.table, marginTop: 8 }}><tbody>{bills.map(b => <tr key={b.id}><td style={styles.tdMono}>{b.id}</td><td style={styles.td}>{fmtDate(b.date)}</td><td style={styles.td}>{b.vendor}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(computeDocTotals(b, data.taxGroups).finalAmount)}</td></tr>)}</tbody></table>
        {bills.length === 0 && <div style={{ fontSize: 12, color: theme.muted, marginTop: 6 }}>None.</div>}
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Expenses ({expenses.length})</div>
        <table style={{ ...styles.table, marginTop: 8 }}><tbody>{expenses.map(e => <tr key={e.id}><td style={styles.td}>{fmtDate(e.date)}</td><td style={styles.td}>{e.vendor}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(e.amount)}</td></tr>)}</tbody></table>
        {expenses.length === 0 && <div style={{ fontSize: 12, color: theme.muted, marginTop: 6 }}>None.</div>}
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Time logged ({timesheets.reduce((s, ts) => s + ts.hours, 0).toFixed(1)} hours)</div>
        <table style={{ ...styles.table, marginTop: 8 }}><tbody>{timesheets.map(ts => <tr key={ts.id}><td style={styles.td}>{fmtDate(ts.date)}</td><td style={styles.td}>{ts.employeeId ? (data.employees.find(e => e.id === ts.employeeId)?.name || "-") : ts.employeeName}</td><td style={styles.td}>{ts.description || "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{ts.hours}h</td></tr>)}</tbody></table>
        {timesheets.length === 0 && <div style={{ fontSize: 12, color: theme.muted, marginTop: 6 }}>None.</div>}
      </div>
    </div>
  );
}
function ReportProjectCostSummary({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = data.projects.map(p => {
    const billCost = data.bills.filter(b => b.projectId === p.id).reduce((s, b) => s + computeDocTotals(b, data.taxGroups).subtotal, 0);
    const expenseCost = data.expenses.filter(e => e.projectId === p.id).reduce((s, e) => s + e.amount, 0);
    const timeCost = data.timesheets.filter(ts => ts.projectId === p.id && ts.billable).reduce((s, ts) => s + ts.hours * ts.rate, 0);
    return { ...p, billCost, expenseCost, timeCost, total: billCost + expenseCost + timeCost };
  }).filter(r => r.total > 0);
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Project</th><th style={{ ...styles.th, textAlign: "right" }}>Bills</th><th style={{ ...styles.th, textAlign: "right" }}>Expenses</th><th style={{ ...styles.th, textAlign: "right" }}>Billable time (at rate)</th><th style={{ ...styles.th, textAlign: "right" }}>Total cost</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.billCost)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.expenseCost)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.timeCost)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(r.total)}</td></tr>)}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No project costs recorded yet.</div>}
    </div>
  );
}
function ReportJournalReport({ data, range }) {
  const { styles, fmt, theme } = useUI();
  const { from, to } = rangeToDates(range);
  const rows = data.transactions.filter(t => t.date >= from && t.date <= to).sort((a, b) => a.date.localeCompare(b.date));
  const total = rows.reduce((s, t) => s + t.lines.reduce((s2, l) => s2 + l.debit, 0), 0);
  return (
    <div className="ces-card" style={styles.card}>
      {rows.map(t => (
        <div key={t.id} style={{ borderBottom: `1px solid ${theme.border}`, padding: "8px 0" }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
            <span><strong style={{ fontFamily: "Inter, sans-serif" }}>{fmtDate(t.date)}</strong> · {t.memo} <span style={{ color: theme.muted }}>({t.source})</span></span>
          </div>
          <table style={{ ...styles.table, marginTop: 4 }}><tbody>{t.lines.map((l, i) => { const acc = data.accounts.find(a => a.id === l.accountId); return (
            <tr key={i}><td style={{ ...styles.td, paddingLeft: 14, fontSize: 12 }}>{acc ? `${acc.code} ${acc.name}` : l.accountId}</td>
              <td style={{ ...styles.tdMono, textAlign: "right", fontSize: 12 }}>{l.debit ? fmt(l.debit) : ""}</td>
              <td style={{ ...styles.tdMono, textAlign: "right", fontSize: 12 }}>{l.credit ? fmt(l.credit) : ""}</td></tr>
          ); })}</tbody></table>
        </div>
      ))}
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13 }}>No entries in this period.</div>}
      {rows.length > 0 && <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>{rows.length} entries totaling {fmt(total)} in debits.</div>}
    </div>
  );
}
function ReportUnrealizedGainLoss({ data, balances }) {
  const { styles, fmt, theme } = useUI();
  const fxAccounts = data.accounts.filter(a => a.currency && a.status !== "inactive" && (a.fcBalance || a.fxRate));
  const rows = fxAccounts.map(a => {
    const bal = balances[a.id] || 0;
    const r = computeFXRevaluation(a, bal, a.fcBalance || 0, a.fxRate || 0);
    return { ...a, glBalance: bal, revaluedValue: r.revaluedValue, pnlImpact: r.pnlImpact };
  }).filter(r => Math.abs(r.pnlImpact) > 0.5);
  const total = rows.reduce((s, r) => s + r.pnlImpact, 0);
  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ fontSize: 12, color: theme.muted, marginBottom: 10 }}>Shows the gain or loss that would be recognized if each foreign-currency account were revalued today, using the last foreign-currency balance and rate recorded for it (see Chart of Accounts → Foreign currency revaluation). Nothing here is posted - it's a preview.</div>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Account</th><th style={styles.th}>Currency</th><th style={{ ...styles.th, textAlign: "right" }}>Ledger balance</th><th style={{ ...styles.th, textAlign: "right" }}>Revalued</th><th style={{ ...styles.th, textAlign: "right" }}>Unrealized gain/(loss)</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td style={styles.td}>{r.name}</td><td style={styles.td}>{r.currency}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.glBalance)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.revaluedValue)}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.pnlImpact >= 0 ? theme.emerald : theme.rose, fontWeight: 600 }}>{fmt(r.pnlImpact)}</td></tr>)}</tbody>
        <tfoot><tr><td colSpan={4} style={{ ...styles.td, fontWeight: 700 }}>Total</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700, color: total >= 0 ? theme.emerald : theme.rose }}>{fmt(total)}</td></tr></tfoot>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No foreign-currency accounts with a recorded balance and rate yet.</div>}
    </div>
  );
}
function ReportProjectSummary({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = data.projects.map(p => {
    const revenue = data.invoices.filter(i => i.projectId === p.id).reduce((s, i) => s + computeDocTotals(i, data.taxGroups).subtotal, 0);
    const cost = data.bills.filter(b => b.projectId === p.id).reduce((s, b) => s + computeDocTotals(b, data.taxGroups).subtotal, 0) + data.expenses.filter(e => e.projectId === p.id).reduce((s, e) => s + e.amount, 0);
    return { ...p, revenue, cost, profit: revenue - cost };
  });
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Project</th><th style={styles.th}>Client</th><th style={{ ...styles.th, textAlign: "right" }}>Revenue</th><th style={{ ...styles.th, textAlign: "right" }}>Cost</th><th style={{ ...styles.th, textAlign: "right" }}>Profit</th></tr></thead>
        <tbody>{rows.map(r => (<tr key={r.id}><td style={styles.td}>{r.name}</td><td style={styles.td}>{r.client}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.revenue)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.cost)}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.profit >= 0 ? theme.emerald : theme.rose, fontWeight: 600 }}>{fmt(r.profit)}</td></tr>))}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No projects yet - add one in Settings.</div>}
    </div>
  );
}
function ReportProjectPerformance({ data }) {
  const { styles, fmt, theme } = useUI();
  const rows = data.projects.map(p => {
    const revenue = data.invoices.filter(i => i.projectId === p.id).reduce((s, i) => s + computeDocTotals(i, data.taxGroups).subtotal, 0);
    const cost = data.bills.filter(b => b.projectId === p.id).reduce((s, b) => s + computeDocTotals(b, data.taxGroups).subtotal, 0) + data.expenses.filter(e => e.projectId === p.id).reduce((s, e) => s + e.amount, 0);
    const margin = revenue ? (revenue - cost) / revenue : 0;
    const budgetUsedPct = p.budget ? cost / p.budget : 0;
    return { ...p, revenue, cost, margin, budgetUsedPct };
  });
  return (
    <div className="ces-card" style={styles.card}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Project</th><th style={{ ...styles.th, textAlign: "right" }}>Margin</th><th style={{ ...styles.th, textAlign: "right" }}>Budget</th><th style={{ ...styles.th, textAlign: "right" }}>Spent</th><th style={{ ...styles.th, textAlign: "right" }}>Budget used</th></tr></thead>
        <tbody>{rows.map(r => (<tr key={r.id}><td style={styles.td}>{r.name}</td><td style={{ ...styles.tdMono, textAlign: "right", color: r.margin >= 0 ? theme.emerald : theme.rose }}>{(r.margin * 100).toFixed(1)}%</td>
          <td style={{ ...styles.tdMono, textAlign: "right" }}>{r.budget ? fmt(r.budget) : "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.cost)}</td>
          <td style={{ ...styles.tdMono, textAlign: "right", color: r.budgetUsedPct > 1 ? theme.rose : r.budgetUsedPct > 0.85 ? theme.amber : theme.emerald }}>{r.budget ? (r.budgetUsedPct * 100).toFixed(0) + "%" : "-"}</td></tr>))}</tbody>
      </table>
      {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No projects yet - add one in Settings.</div>}
    </div>
  );
}
// Looks up a budget for one account in one month: a month-specific override
// if one's been set, otherwise the account's default budget, otherwise 0.
function budgetFor(data, month, accountId) {
  return data.budgets?.[month]?.[accountId] ?? data.budgets?.default?.[accountId] ?? 0;
}
function ReportBudgetVsActual({ data, balances, month }) {
  const { styles, fmt, theme } = useUI();
  const [varianceView, setVarianceView] = useState("amount");
  const [y, m] = month.split("-").map(Number);
  const from = localDateToISO(new Date(y, m - 1, 1)), to = localDateToISO(new Date(y, m, 0));
  const mv = periodMovement(data.accounts, data.transactions, from, to);
  const accts = data.budgetAccounts.map(id => data.accounts.find(a => a.id === id)).filter(Boolean);
  return (
    <div>
      <div className="ces-card" style={styles.card}>
        <div className="no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontSize: 12, color: theme.muted }}>{new Date(y, m - 1, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" })} - a read-only comparison. Set or change budget amounts under Accounting → Budgets; changes there are reflected here automatically.</div>
          <div style={{ display: "flex", gap: 6 }}>
            {[["amount", "Amount"], ["percent", "%"]].map(([v, l]) => (
              <button key={v} style={{ ...styles.btnGhost, ...(varianceView === v ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }} onClick={() => setVarianceView(v)}>{l}</button>
            ))}
          </div>
        </div>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Account</th><th style={{ ...styles.th, textAlign: "right" }}>Budget</th><th style={{ ...styles.th, textAlign: "right" }}>Actual</th><th style={{ ...styles.th, textAlign: "right" }}>Variance</th></tr></thead>
          <tbody>{accts.map(a => {
            const actual = a.contra ? -(mv[a.id] || 0) : (mv[a.id] || 0);
            const budget = budgetFor(data, month, a.id);
            // Favorable variance is always positive: above budget for
            // revenue (sold more than expected), below budget for expenses
            // (spent less than expected) - the standard budgeting convention.
            const variance = a.type === "expense" ? budget - actual : actual - budget;
            const variancePct = budget !== 0 ? (variance / Math.abs(budget)) * 100 : (variance !== 0 ? 100 : 0);
            const badVariance = variance < 0;
            return (
              <tr key={a.id}><td style={styles.td}>{a.name}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(budget)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(actual)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", color: badVariance ? theme.rose : theme.emerald }}>{varianceView === "amount" ? `${variance >= 0 ? "+" : ""}${fmt(variance)}` : `${variancePct >= 0 ? "+" : ""}${variancePct.toFixed(1)}%`}</td></tr>
            );
          })}</tbody>
        </table>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>Enter a monthly budget per account - it's saved automatically and compared against actual postings for the selected month. Positive variance is always favorable: above budget for revenue, below budget for expenses.</div>
      </div>
    </div>
  );
}
function computeRatios(data, balances) {
  const revenue = sumByType(data.accounts, balances, "revenue", null);
  const operatingRevenue = data.accounts.filter(a => a.type === "revenue" && a.subtype !== "other").reduce((s, a) => s + (a.contra ? -(balances[a.id] || 0) : (balances[a.id] || 0)), 0);
  const cogs = data.accounts.filter(a => a.type === "expense" && a.subtype === "cogs").reduce((s, a) => s + (a.contra ? -(balances[a.id] || 0) : (balances[a.id] || 0)), 0);
  const totalExpense = sumByType(data.accounts, balances, "expense");
  const netIncome = revenue - totalExpense;
  const grossProfit = operatingRevenue - cogs;
  const currentAssets = sumByType(data.accounts, balances, "asset", "current") + data.accounts.filter(a => a.subtype === "bank").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const cashAndBank = data.accounts.filter(a => a.subtype === "bank").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const inventoryBal = ["1200", "1201", "1202", "1203"].reduce((s, id) => s + (balances[id] || 0), 0);
  const currentLiabilities = sumByType(data.accounts, balances, "liability", "current");
  const totalAssets = sumByType(data.accounts, balances, "asset");
  const totalEquity = sumByType(data.accounts, balances, "equity") + netIncomeAllTime(data);
  return {
    grossMargin: operatingRevenue ? grossProfit / operatingRevenue : 0,
    netMargin: revenue ? netIncome / revenue : 0,
    roa: totalAssets ? netIncome / totalAssets : 0,
    roe: totalEquity ? netIncome / totalEquity : 0,
    currentRatio: currentLiabilities ? currentAssets / currentLiabilities : Infinity,
    quickRatio: currentLiabilities ? (currentAssets - inventoryBal) / currentLiabilities : Infinity,
    cashRatio: currentLiabilities ? cashAndBank / currentLiabilities : Infinity,
    workingCapital: currentAssets - currentLiabilities,
  };
}
function ReportRatiosCard({ data, balances }) {
  const { styles, fmt } = useUI();
  const ratios = computeRatios(data, balances);
  const pct = (v) => isFinite(v) ? `${(v * 100).toFixed(1)}%` : "-";
  const ratioTone = (v, good) => !isFinite(v) ? "emerald" : v >= good ? "emerald" : v >= good * 0.6 ? "amber" : "rose";
  return (
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Profitability ratios</div>
        <RatioRow label="Gross profit margin" value={pct(ratios.grossMargin)} tone={ratioTone(ratios.grossMargin, 0.4)} note="Gross profit ÷ revenue" />
        <RatioRow label="Net profit margin" value={pct(ratios.netMargin)} tone={ratioTone(ratios.netMargin, 0.15)} note="Net income ÷ revenue" />
        <RatioRow label="Return on assets (ROA)" value={pct(ratios.roa)} tone={ratioTone(ratios.roa, 0.08)} note="Net income ÷ total assets" />
        <RatioRow label="Return on equity (ROE)" value={pct(ratios.roe)} tone={ratioTone(ratios.roe, 0.15)} note="Net income ÷ total equity" />
      </div>
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Liquidity ratios</div>
        <RatioRow label="Current ratio" value={isFinite(ratios.currentRatio) ? ratios.currentRatio.toFixed(2) : "-"} tone={ratioTone(ratios.currentRatio, 1.5)} note="Current assets ÷ current liabilities" />
        <RatioRow label="Quick ratio" value={isFinite(ratios.quickRatio) ? ratios.quickRatio.toFixed(2) : "-"} tone={ratioTone(ratios.quickRatio, 1)} note="(Current assets − inventory) ÷ current liabilities" />
        <RatioRow label="Cash ratio" value={isFinite(ratios.cashRatio) ? ratios.cashRatio.toFixed(2) : "-"} tone={ratioTone(ratios.cashRatio, 0.5)} note="Cash & bank ÷ current liabilities" />
        <RatioRow label="Working capital" value={fmt(ratios.workingCapital)} tone={ratios.workingCapital >= 0 ? "emerald" : "rose"} note="Current assets − current liabilities" />
      </div>
    </div>
  );
}
function ReportRealizedGL({ data }) {
  const { styles, fmt, theme } = useUI();
  const lines = data.transactions.flatMap(t => t.lines.filter(l => l.accountId === "6000").map(l => ({ ...l, date: t.date, memo: t.memo }))).sort((a, b) => b.date.localeCompare(a.date));
  const cum = lines.reduce((s, l) => s + (l.credit - l.debit), 0);
  return (
    <div className="ces-card" style={styles.card}>
      <RowLine label="Cumulative realized gain / (loss)" value={cum} bold tone={cum >= 0 ? "emerald" : "rose"} />
      <table style={{ ...styles.table, marginTop: 12 }}>
        <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
        <tbody>{lines.map((l, i) => <tr key={i}><td style={styles.tdMono}>{fmtDate(l.date)}</td><td style={styles.td}>{l.memo}</td><td style={{ ...styles.tdMono, textAlign: "right", color: (l.credit - l.debit) >= 0 ? theme.emerald : theme.rose }}>{fmt(l.credit - l.debit)}</td></tr>)}</tbody>
      </table>
      {lines.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No realized gains or losses posted yet. Post one via the Journal tab using the "Realized Gains & Losses" account - credit for a gain, debit for a loss.</div>}
    </div>
  );
}

/* =================================== Forecast + Ratios =================================== */
function Forecast({ data, balances }) {
  const { styles, fmt, theme } = useUI();
  const [horizon, setHorizon] = useState(3);
  const monthly = useMemo(() => aggregateMonthly(data), [data]);
  const projection = useMemo(() => projectForward(monthly, horizon), [monthly, horizon]);
  const combined = [...monthly.map(m => ({ ...m, kind: "actual" })), ...projection.map(m => ({ ...m, kind: "forecast" }))];
  const growthRate = monthly.length >= 2 ? ((monthly[monthly.length - 1].revenue - monthly[0].revenue) / Math.max(monthly[0].revenue, 1)) / Math.max(monthly.length - 1, 1) : 0;
  const runwayMonths = estimateRunway(data, monthly);

  const ratios = computeRatios(data, balances);
  const pct = (v) => isFinite(v) ? `${(v * 100).toFixed(1)}%` : "-";
  const ratioTone = (v, good) => !isFinite(v) ? "emerald" : v >= good ? "emerald" : v >= good * 0.6 ? "amber" : "rose";

  return (
    <div>
      <PageHeader eyebrow="Planning" title="Modeling & Forecast" sub="Trend-based projection, profitability and liquidity ratios" action={
        <select style={styles.input} value={horizon} onChange={e => setHorizon(Number(e.target.value))}><option value={3}>3 month horizon</option><option value={6}>6 month horizon</option><option value={12}>12 month horizon</option></select>
      } />
      <div style={styles.kpiRow}>
        <Kpi label="Avg. monthly revenue growth" value={pct(growthRate)} tone={growthRate >= 0 ? "emerald" : "rose"} />
        <Kpi label="Projected net income, next month" value={fmt(projection[0]?.net || 0)} tone={(projection[0]?.net || 0) >= 0 ? "emerald" : "rose"} />
        <Kpi label="Cash runway" value={runwayMonths === Infinity ? "profitable" : `${runwayMonths.toFixed(1)} mo`} tone={runwayMonths === Infinity || runwayMonths > 6 ? "emerald" : runwayMonths > 2 ? "amber" : "rose"} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Profitability ratios</div>
          <RatioRow label="Gross profit margin" value={pct(ratios.grossMargin)} tone={ratioTone(ratios.grossMargin, 0.4)} note="Gross profit ÷ revenue" />
          <RatioRow label="Net profit margin" value={pct(ratios.netMargin)} tone={ratioTone(ratios.netMargin, 0.15)} note="Net income ÷ revenue" />
          <RatioRow label="Return on assets (ROA)" value={pct(ratios.roa)} tone={ratioTone(ratios.roa, 0.08)} note="Net income ÷ total assets" />
          <RatioRow label="Return on equity (ROE)" value={pct(ratios.roe)} tone={ratioTone(ratios.roe, 0.15)} note="Net income ÷ total equity" />
        </div>
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Liquidity ratios</div>
          <RatioRow label="Current ratio" value={isFinite(ratios.currentRatio) ? ratios.currentRatio.toFixed(2) : "-"} tone={ratioTone(ratios.currentRatio, 1.5)} note="Current assets ÷ current liabilities" />
          <RatioRow label="Quick ratio" value={isFinite(ratios.quickRatio) ? ratios.quickRatio.toFixed(2) : "-"} tone={ratioTone(ratios.quickRatio, 1)} note="(Current assets − inventory) ÷ current liabilities" />
          <RatioRow label="Cash ratio" value={isFinite(ratios.cashRatio) ? ratios.cashRatio.toFixed(2) : "-"} tone={ratioTone(ratios.cashRatio, 0.5)} note="Cash & bank ÷ current liabilities" />
          <RatioRow label="Working capital" value={fmt(ratios.workingCapital)} tone={ratios.workingCapital >= 0 ? "emerald" : "rose"} note="Current assets − current liabilities" />
        </div>
      </div>

      <div className="ces-card" style={styles.cardWide}>
        <div style={styles.cardTitle}>Revenue & expense trend with {horizon}-month projection</div>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={combined} margin={{ left: 0, right: 12, top: 10 }}>
            <CartesianGrid strokeDasharray="2 4" stroke={theme.border} vertical={false} />
            <XAxis dataKey="label" tick={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, fill: theme.muted }} axisLine={{ stroke: theme.border }} tickLine={false} />
            <YAxis tick={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 11, fill: theme.muted }} axisLine={false} tickLine={false} tickFormatter={v => `${v / 1000}k`} />
            <Tooltip formatter={v => fmt(v)} contentStyle={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 12, background: theme.panel }} />
            <Legend wrapperStyle={{ fontFamily: "Inter", fontSize: 12 }} />
            <Line type="monotone" dataKey="revenue" stroke={theme.emerald} strokeWidth={2} dot={{ r: 3 }} name="Revenue" />
            <Line type="monotone" dataKey="expenses" stroke={theme.rose} strokeWidth={2} dot={{ r: 3 }} name="Expenses" />
            <Line type="monotone" dataKey="net" stroke={theme.accent} strokeWidth={1.5} strokeDasharray="4 3" dot={false} name="Net" />
          </LineChart>
        </ResponsiveContainer>
        <div style={{ fontSize: 11, color: theme.muted, marginTop: 8, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>Solid lines are actuals from posted transactions; the model projects forward using a linear trend fit to recent months. Treat this as directional, not a guarantee.</div>
      </div>
    </div>
  );
}
function RatioRow({ label, value, tone, note }) {
  const { theme } = useUI();
  const c = tone === "emerald" ? theme.emerald : tone === "amber" ? theme.amber : theme.rose;
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${theme.border}` }}>
      <div><div style={{ fontSize: 13.5 }}>{label}</div><div style={{ fontSize: 11, color: theme.muted }}>{note}</div></div>
      <div style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 15, fontWeight: 600, color: c }}>{value}</div>
    </div>
  );
}
function aggregateMonthly(data) {
  const keys = {};
  data.transactions.forEach(t => {
    const k = monthKey(t.date); if (!keys[k]) keys[k] = { revenue: 0, expenses: 0 };
    t.lines.forEach(l => { const acc = data.accounts.find(a => a.id === l.accountId); if (!acc) return;
      if (acc.type === "revenue") keys[k].revenue += (l.credit - l.debit);
      if (acc.type === "expense") keys[k].expenses += (l.debit - l.credit); });
  });
  return Object.keys(keys).sort().map(k => { const [y, m] = k.split("-"); const label = new Date(Number(y), Number(m) - 1, 1).toLocaleDateString(undefined, { month: "short", year: "2-digit" });
    return { key: k, label, revenue: Math.round(keys[k].revenue), expenses: Math.round(keys[k].expenses), net: Math.round(keys[k].revenue - keys[k].expenses) }; });
}
function linearFit(values) {
  const n = values.length; if (n < 2) return { slope: 0, intercept: values[0] || 0 };
  const xs = values.map((_, i) => i); const xMean = xs.reduce((a, b) => a + b, 0) / n; const yMean = values.reduce((a, b) => a + b, 0) / n;
  let num = 0, den = 0; xs.forEach((x, i) => { num += (x - xMean) * (values[i] - yMean); den += (x - xMean) ** 2; });
  const slope = den === 0 ? 0 : num / den; return { slope, intercept: yMean - slope * xMean };
}
function projectForward(monthly, horizon) {
  const recent = monthly.slice(-6); const revFit = linearFit(recent.map(m => m.revenue)); const expFit = linearFit(recent.map(m => m.expenses));
  const startIdx = recent.length; const lastKey = monthly[monthly.length - 1]?.key || monthKey(new Date());
  const [y, m] = lastKey.split("-").map(Number); const out = [];
  for (let i = 1; i <= horizon; i++) { const d = new Date(y, m - 1 + i, 1);
    const rev = Math.max(0, Math.round(revFit.slope * (startIdx + i - 1) + revFit.intercept));
    const exp = Math.max(0, Math.round(expFit.slope * (startIdx + i - 1) + expFit.intercept));
    out.push({ key: monthKey(d), label: d.toLocaleDateString(undefined, { month: "short", year: "2-digit" }), revenue: rev, expenses: exp, net: rev - exp }); }
  return out;
}
function estimateRunway(data, monthly) {
  const bal = computeBalances(data.accounts, data.transactions);
  const cash = data.banks.reduce((s, b) => s + (bal[b.accountId] || 0), 0);
  const recent = monthly.slice(-3); const avgNet = recent.length ? recent.reduce((s, m) => s + m.net, 0) / recent.length : 0;
  if (avgNet >= 0) return Infinity; return cash / Math.abs(avgNet);
}

/* =================================== Settings =================================== */
// Read an uploaded logo, downscale it to a small PNG data-URL (max 240px on
// the long edge) so it stores compactly, and save it into settings.
function handleLogoFile(file, setData, notify) {
  const reader = new FileReader();
  reader.onload = () => {
    const img = new Image();
    img.onload = () => {
      const MAX = 640; // stored at 640px so it stays crisp in the large Settings/report badge
      const scale = Math.min(1, MAX / Math.max(img.width, img.height));
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(img.width * scale));
      canvas.height = Math.max(1, Math.round(img.height * scale));
      canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
      const dataUrl = canvas.toDataURL("image/png");
      setData(d => ({ ...d, settings: { ...d.settings, logo: dataUrl } }));
      notify("Logo saved");
    };
    img.onerror = () => notify("Couldn't read that image - try a PNG or JPG");
    img.src = reader.result;
  };
  reader.onerror = () => notify("Couldn't read that file");
  reader.readAsDataURL(file);
}
// Read-only, searchable, filterable log of every recorded action, its time,
// and who performed it. Entries are appended automatically by the setData
// wrapper in App - nothing here writes to the log, only displays it.
function AuditTrail({ data }) {
  const { styles, theme } = useUI();
  const [q, setQ] = useState("");
  const [user, setUser] = useState("all");
  const [range, setRange] = useState("all");
  const { from, to } = rangeToDates(range);
  const users = [...new Set(data.auditLog.map(l => l.user))];
  const rows = [...data.auditLog].sort((a, b) => b.time.localeCompare(a.time)).filter(l => {
    const d = l.time.slice(0, 10);
    if (d < from || d > to) return false;
    if (user !== "all" && l.user !== user) return false;
    if (q && !l.action.toLowerCase().includes(q.toLowerCase())) return false;
    return true;
  });
  const fmtTime = (iso) => { const d = new Date(iso); return `${fmtDate(localDateToISO(d))} ${d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })}`; };
  const iconFor = (action) => action.startsWith("Deleted") || action.startsWith("Removed") || action.includes("cleared") || action.includes("reset") ? "🗑"
    : action.startsWith("Created") || action.startsWith("Posted") || action.startsWith("Imported") ? "\u2795"
    : action.startsWith("Updated") || action.startsWith("Edited") || action.startsWith("Renamed") || action.startsWith("Changed") ? "\u270e"
    : action.includes("categorized") ? "\ud83c\udff7" : "\u2022";

  return (
    <div>
      <PageHeader eyebrow="System" title="Audit Trail" sub={`${data.auditLog.length} recorded actions`} />
      <div className="ces-card" style={{ ...styles.card, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
        <input style={{ ...styles.input, flex: 1, minWidth: 200 }} placeholder="Search actions..." value={q} onChange={e => setQ(e.target.value)} />
        <select style={styles.input} value={user} onChange={e => setUser(e.target.value)}>
          <option value="all">All users</option>{users.map(u => <option key={u} value={u}>{u}</option>)}
        </select>
        <select style={styles.input} value={range} onChange={e => setRange(e.target.value)}>
          {RANGE_OPTIONS.filter(([v]) => v !== "custom").map(([v, l]) => <option key={v} value={v}>{l}</option>)}
        </select>
      </div>
      <div className="ces-card" style={styles.card}>
        {rows.length === 0 ? <div style={{ fontSize: 13, color: theme.muted }}>No actions match these filters.</div> : (
          <table style={styles.table}>
            <thead><tr><th style={{ ...styles.th, width: 150 }}>Time</th><th style={{ ...styles.th, width: 110 }}>User</th><th style={styles.th}>Action</th></tr></thead>
            <tbody>{rows.map(l => (
              <tr key={l.id}>
                <td style={styles.tdMono}>{fmtTime(l.time)}</td>
                <td style={styles.td}>{l.user}</td>
                <td style={styles.td}><span style={{ marginRight: 6 }}>{iconFor(l.action)}</span>{l.action}</td>
              </tr>
            ))}</tbody>
          </table>
        )}
      </div>
      <div style={{ fontSize: 12, color: theme.muted }}>Every create, edit, delete, post, and reset anywhere in the app is recorded here automatically with a timestamp and the acting user - it cannot be turned off or edited from within the app. The user name is set in Settings; on a shared instance, ask each person to set their own name before working.</div>
    </div>
  );
}
// Shows accounts recommended for the chosen industry + country that aren't
// already in the Chart of Accounts, with checkboxes to pick which to add.
// Codes are assigned automatically (nextFreeCode) so nothing can collide.
function SuggestedAccounts({ data, setData, notify, industry, country }) {
  const { styles, theme } = useUI();
  const existingNames = new Set(data.accounts.map(a => a.name.trim().toLowerCase()));
  const suggestions = [
    ...(INDUSTRY_COA_SUGGESTIONS[industry] || []),
    ...(COUNTRY_COA_SUGGESTIONS[country] || []),
  ].filter(([name]) => !existingNames.has(name.trim().toLowerCase()));
  const [checked, setChecked] = useState(new Set(suggestions.map(([name]) => name)));
  useEffect(() => { setChecked(new Set(suggestions.map(([name]) => name))); }, [industry, country]); // eslint-disable-line

  const toggle = (name) => setChecked(prev => { const s = new Set(prev); s.has(name) ? s.delete(name) : s.add(name); return s; });
  const addSelected = () => {
    const toAdd = suggestions.filter(([name]) => checked.has(name));
    if (toAdd.length === 0) return notify("Nothing selected to add");
    setData(d => {
      let accounts = d.accounts;
      const codeStart = { asset: 1400, liability: 2300, equity: 3100, revenue: 4300, expense: 5800 };
      toAdd.forEach(([name, type, category]) => {
        const code = nextFreeCode(accounts, codeStart[type] || 1900, 10);
        accounts = [...accounts, normalizeAccount({ id: code, code, name, type, category, status: "active" })];
      });
      return { ...d, accounts, settings: { ...d.settings, industry, country } };
    });
    notify(`${toAdd.length} account(s) added to the Chart of Accounts`);
  };

  if (suggestions.length === 0) {
    return <div style={{ fontSize: 12.5, color: theme.muted, marginTop: 12 }}>No additional accounts suggested - either this is General/Other, or your Chart of Accounts already has everything typically needed here.</div>;
  }
  return (
    <div style={{ marginTop: 16, borderTop: `1px solid ${theme.border}`, paddingTop: 14 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted }}>Suggested accounts ({suggestions.length})</div>
        <button style={styles.btnPrimary} onClick={addSelected}>Add selected to Chart of Accounts</button>
      </div>
      <div style={{ marginTop: 10, display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 6 }}>
        {suggestions.map(([name, type, category]) => (
          <label key={name} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13, cursor: "pointer" }}>
            <input type="checkbox" checked={checked.has(name)} onChange={() => toggle(name)} />
            <span>{name}</span>
            <span style={{ fontSize: 10.5, color: theme.muted }}>{TYPE_LABELS[type]} · {category}</span>
          </label>
        ))}
      </div>
      <div style={{ fontSize: 11, color: theme.muted, marginTop: 10 }}>Adding accounts here saves your industry and country to Business profile too. You can always create, rename or remove accounts yourself in the Chart of Accounts tab afterward.</div>
    </div>
  );
}
// A collapsible section for Settings - closed by default so the page reads
// as a clean list of what's configurable, rather than every field at once.
// `plain` skips the extra card wrapper for children that already render
// their own (TaxGroupsManager, ProjectsManager, etc. are shared with other
// pages, so they're left untouched rather than reworked just for this).
// Penetration pricing: deliberately well under the market rates these were
// first benchmarked against (QuickBooks/Xero/Zoho run $19-275/mo with no
// free tier below Zoho's) - a genuine Free plan plus paid tiers priced
// roughly half of comparable competitor tiers, to win share on price while
// still spanning the same feature range from basic invoicing up through
// IFRS-standard deferred tax, leases and provisions.
const FREE_PLAN = { id: "free", name: "Free", monthly: 0, includedUsers: 1, tagline: "Invoicing and expense tracking, no cost", features: ["Invoicing & billing", "Expense tracking", "Basic reports", "1 user"], includedAddOns: [] };
const SUBSCRIPTION_PLANS = [
  { id: "starter", name: "Starter", monthly: 9, includedUsers: 1, tagline: "Freelancers and solo founders", features: ["Invoicing & billing", "Bank reconciliation", "Expense tracking", "Core financial reports"], includedAddOns: [] },
  { id: "growth", name: "Growth", monthly: 25, includedUsers: 3, tagline: "Small businesses finding their footing", features: ["Everything in Starter", "Inventory & fixed assets", "Multi-currency & FX revaluation", "Projects & budgets", "25+ reports"], includedAddOns: [] },
  { id: "professional", name: "Professional", monthly: 49, includedUsers: 5, tagline: "Growing teams with real complexity", features: ["Everything in Growth", "Payroll", "Sales & purchase orders", "Time sheets", "50+ reports"], includedAddOns: ["payroll"] },
  { id: "enterprise", name: "Enterprise", monthly: 89, includedUsers: 10, tagline: "Full IFRS-standard compliance & controls", features: ["Everything in Professional", "Deferred tax, leases & provisions", "Transaction locking & bulk update", "Priority support"], includedAddOns: ["payroll", "tax-pack", "priority-support"] },
];
const SUBSCRIPTION_ADDONS = [
  { id: "extra-user", name: "Additional user", price: 4, unit: "per user / month" },
  { id: "payroll", name: "Payroll module", price: 12, unit: "per month" },
  { id: "tax-pack", name: "Advanced tax & compliance pack", price: 18, unit: "per month", desc: "Deferred tax, leases, provisions, revaluation, expected credit loss" },
  { id: "priority-support", name: "Priority support", price: 9, unit: "per month" },
  { id: "api-access", name: "API access", price: 15, unit: "per month" },
];
// 6-month discount is deliberately exactly half of the annual discount.
const BILLING_FREQUENCIES = [
  { id: "monthly", label: "Monthly", months: 1, discount: 0 },
  { id: "quarterly", label: "Quarterly", months: 3, discount: 0.05 },
  { id: "semiannual", label: "6 Months", months: 6, discount: 0.10 },
  { id: "annual", label: "Annually", months: 12, discount: 0.20 },
];
const TRIAL_DAYS = 30;
function computeSubscriptionPrice(plan, frequencyId, extraUsers, addOnIds) {
  const freq = BILLING_FREQUENCIES.find(f => f.id === frequencyId) || BILLING_FREQUENCIES[0];
  const includedAddOns = new Set(plan.includedAddOns || []);
  const extraUserPrice = SUBSCRIPTION_ADDONS.find(a => a.id === "extra-user").price;
  const extraUsersMonthly = Math.max(0, extraUsers) * extraUserPrice;
  const addOnMonthly = addOnIds.filter(id => id !== "extra-user" && !includedAddOns.has(id)).reduce((s, id) => s + (SUBSCRIPTION_ADDONS.find(a => a.id === id)?.price || 0), 0);
  const baseMonthly = plan.monthly + extraUsersMonthly + addOnMonthly;
  const fullPriceForPeriod = baseMonthly * freq.months;
  const discountAmount = fullPriceForPeriod * freq.discount;
  const totalForPeriod = fullPriceForPeriod - discountAmount;
  const effectiveMonthly = totalForPeriod / freq.months;
  return { baseMonthly, freq, fullPriceForPeriod, discountAmount, totalForPeriod, effectiveMonthly, extraUsersMonthly, addOnMonthly };
}
// No backend means no server-side cron to expire a trial - instead, this
// gets checked live from the stored start date every time it's read.
// Anyone still "trialing" past the cutoff without ever confirming is
// treated as expired (and gets reverted to Free) - not silently left on a
// paid plan they never actually subscribed to.
function subscriptionStatus(subscription, asOfDate) {
  if (!subscription || !subscription.planId || subscription.planId === "free") return { effectivePlanId: "free", status: "free", daysLeft: null };
  if (subscription.status === "active") return { effectivePlanId: subscription.planId, status: "active", daysLeft: null };
  if (subscription.status === "trialing") {
    const daysLeft = Math.ceil((new Date(subscription.trialEndDate) - new Date(asOfDate)) / 86400000);
    if (daysLeft < 0) return { effectivePlanId: "free", status: "expired", daysLeft: 0 };
    return { effectivePlanId: subscription.planId, status: "trialing", daysLeft };
  }
  return { effectivePlanId: "free", status: "free", daysLeft: null };
}
// A small lightbulb that reveals a helper note only when someone asks for
// it, instead of a paragraph of gray text sitting on the page permanently.
function InfoTip({ text, width = 260 }) {
  const { theme } = useUI();
  const [open, setOpen] = useState(false);
  return (
    <span style={{ position: "relative", display: "inline-block", marginLeft: 6, verticalAlign: "middle" }}>
      <button
        type="button"
        onClick={(e) => { e.stopPropagation(); setOpen(o => !o); }}
        onMouseEnter={() => setOpen(true)}
        onMouseLeave={() => setOpen(false)}
        style={{ background: "none", border: "none", cursor: "pointer", fontSize: 13, padding: 0, lineHeight: 1, opacity: 0.75 }}
        aria-label="More info"
      >💡</button>
      {open && (
        <div style={{ position: "absolute", zIndex: 50, top: "130%", left: 0, width, background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 12, lineHeight: 1.45, color: theme.text, boxShadow: theme.shadowMd, fontWeight: 400 }}>
          {text}
        </div>
      )}
    </span>
  );
}
function SettingsSection({ title, subtitle, defaultOpen = false, plain = false, children }) {
  const { styles, theme } = useUI();
  const [open, setOpen] = useState(defaultOpen);
  const header = (
    <div
      className="ces-card" onClick={() => setOpen(o => !o)}
      style={{ ...styles.card, marginBottom: 0, padding: "16px 20px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}
    >
      <div>
        <div style={{ fontFamily: "Inter, sans-serif", fontWeight: 600, fontSize: 14, color: theme.text }}>{title}</div>
        {subtitle && <div style={{ fontSize: 12, color: theme.muted, marginTop: 2 }}>{subtitle}</div>}
      </div>
      <span style={{ fontSize: 11, color: theme.muted, transition: "transform 0.2s ease", display: "inline-block", transform: open ? "rotate(180deg)" : "rotate(0deg)" }}>▾</span>
    </div>
  );
  if (plain) return <div style={{ marginBottom: 12 }}>{header}{open && <div style={{ marginTop: 10 }}>{children}</div>}</div>;
  return (
    <div style={{ marginBottom: 12 }}>
      {header}
      {open && <div className="ces-card" style={{ ...styles.card, marginTop: 10 }}>{children}</div>}
    </div>
  );
}
function Settings({ data, setData, notify }) {
  const { styles, theme } = useUI();
  const [local, setLocal] = useState(data.settings);
  const save = () => { setData(d => ({ ...d, settings: local })); notify("Settings saved"); };
  const accentSwatches = ["#2E5AAC", "#0E7FA6", "#6D5BD0", "#17A673", "#D98A1F", "#D6455A"];
  return (
    <div>
      <PageHeader eyebrow="Configuration" title="Settings" sub="Branding, theme and currency" />

      <SettingsSection title="Company" subtitle="Name and logo">
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap", alignItems: "center" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 220 }} value={local.companyName} onChange={e => setLocal({ ...local, companyName: e.target.value })} placeholder="Company name" />
        </div>
        <div style={{ display: "flex", gap: 14, marginTop: 14, alignItems: "center", flexWrap: "wrap" }}>
          {data.settings.logo
            ? <div className="ces-logo" style={{ padding: 14 }}><img src={data.settings.logo} alt="Company logo" style={{ height: 84, width: "auto", maxWidth: 220, objectFit: "contain" }} /></div>
            : <div style={{ fontSize: 12, color: theme.muted }}>No logo uploaded yet.</div>}
          <input id="logo-input" type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" style={{ display: "none" }} onChange={e => { const f = e.target.files?.[0]; if (f) handleLogoFile(f, setData, notify); e.target.value = ""; }} />
          <label htmlFor="logo-input" style={{ ...styles.btnGhost, display: "inline-block", cursor: "pointer" }}>{data.settings.logo ? "Replace logo" : "Upload logo"}</label>
          {data.settings.logo && <button style={styles.btnGhost} onClick={() => { setData(d => ({ ...d, settings: { ...d.settings, logo: null } })); notify("Logo removed"); }}>Remove</button>}
          <span style={{ fontSize: 11, color: theme.muted }}>PNG/JPG/SVG · shown in the sidebar, on invoices & bills, and on report headers</span>
        </div>
      </SettingsSection>

      <SettingsSection title="Workspace" subtitle="Your name and accounting basis">
        <div style={{ display: "flex", gap: 14, marginTop: 10, flexWrap: "wrap", alignItems: "flex-start" }}>
          <label style={{ fontSize: 12 }}>Your name (attributed in the audit trail)
            <input style={{ ...styles.input, display: "block", marginTop: 4, width: 200 }} value={local.userName} onChange={e => setLocal({ ...local, userName: e.target.value })} placeholder="Admin" />
          </label>
          <label style={{ fontSize: 12 }}>Accounting basis
            <select style={{ ...styles.input, display: "block", marginTop: 4, width: 200 }} value={local.accountingBasis} onChange={e => setLocal({ ...local, accountingBasis: e.target.value })}>
              <option value="accrual">Accrual basis</option>
              <option value="cash">Cash basis</option>
            </select>
          </label>
        </div>
        <div style={{ fontSize: 11.5, color: theme.muted, marginTop: 8, maxWidth: 520 }}>
          Accrual basis recognizes revenue and expenses when invoiced or billed, regardless of payment - this is what every report currently shows. Cash basis recognizes them only when cash actually moves; reports will filter to paid invoices, paid bills, and settled payments once you switch.
        </div>
      </SettingsSection>

      <SettingsSection title="Business profile" subtitle="Industry and country - drives suggested chart of accounts">
        <div style={{ fontSize: 12.5, color: theme.muted, marginTop: 4 }}>Tell the system what kind of business this is and where it operates, and it will suggest chart-of-accounts entries suited to that industry and country's statutory obligations - you choose which to add.</div>
        <div style={{ display: "flex", gap: 14, marginTop: 12, flexWrap: "wrap" }}>
          <label style={{ fontSize: 12 }}>Industry
            <select style={{ ...styles.input, display: "block", marginTop: 4, width: 240 }} value={local.industry} onChange={e => setLocal({ ...local, industry: e.target.value })}>
              {INDUSTRIES.map(i => <option key={i.id} value={i.id}>{i.name}</option>)}
            </select>
          </label>
          <label style={{ fontSize: 12 }}>Country
            <select style={{ ...styles.input, display: "block", marginTop: 4, width: 220 }} value={local.country} onChange={e => setLocal({ ...local, country: e.target.value })}>
              {COUNTRIES.map(c => <option key={c.code} value={c.code}>{c.name}</option>)}
            </select>
          </label>
        </div>
        <SuggestedAccounts data={data} setData={setData} notify={notify} industry={local.industry} country={local.country} />
      </SettingsSection>

      <SettingsSection title="Currency" subtitle={`${local.currencyCode} - ${CURRENCIES.find(c => c.code === local.currencyCode)?.symbol || ""}`}>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap", alignItems: "center" }}>
          <select style={styles.input} value={local.currencyCode} onChange={e => setLocal({ ...local, currencyCode: e.target.value, currencySymbol: undefined })}>
            {CURRENCIES.map(c => <option key={c.code} value={c.code}>{c.code} ({c.symbol})</option>)}
          </select>
          <input style={{ ...styles.input, width: 90 }} placeholder="Custom symbol (optional)" value={local.currencySymbol || ""} onChange={e => setLocal({ ...local, currencySymbol: e.target.value })} />
        </div>
      </SettingsSection>

      <SettingsSection title="Theme" subtitle="Light or dark, and accent color">
        <div style={{ display: "flex", gap: 16, marginTop: 10, alignItems: "center", flexWrap: "wrap" }}>
          <div style={{ display: "flex", gap: 6 }}>
            <button style={{ ...styles.btnGhost, ...(local.mode === "dark" ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }} onClick={() => setLocal({ ...local, mode: "dark" })}>Dark blue</button>
            <button style={{ ...styles.btnGhost, ...(local.mode === "light" ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }} onClick={() => setLocal({ ...local, mode: "light" })}>Light</button>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <span style={{ fontSize: 12, color: theme.muted }}>Accent</span>
            {accentSwatches.map(c => <button key={c} onClick={() => setLocal({ ...local, accent: c })} style={{ width: 24, height: 24, borderRadius: "50%", background: c, border: local.accent === c ? `2px solid ${theme.text}` : "1px solid transparent", cursor: "pointer" }} />)}
            <input type="color" value={local.accent} onChange={e => setLocal({ ...local, accent: e.target.value })} style={{ width: 30, height: 26, border: "none", background: "none", cursor: "pointer" }} />
          </div>
        </div>
      </SettingsSection>

      <SettingsSection title="Tax groups" subtitle="Tax rates used on invoices and bills" plain>
        <TaxGroupsManager data={data} setData={setData} notify={notify} />
      </SettingsSection>

      <SettingsSection title="Projects" subtitle="Tag invoices, bills and expenses to a project" plain>
        <ProjectsManager data={data} setData={setData} notify={notify} />
      </SettingsSection>

      <SettingsSection title="Locations and departments" subtitle="Tag fixed assets and inventory by location and department" plain>
        <LocationsAndDepartments data={data} setData={setData} notify={notify} />
      </SettingsSection>

      <SettingsSection title="Subscription" subtitle="Compare plans, add-ons and billing frequency" plain>
        <SubscriptionPanel data={data} setData={setData} notify={notify} />
      </SettingsSection>
      <SettingsSection title="Opening balances" subtitle="Set starting balances when migrating from another system" plain>
        <OpeningBalancesPanel data={data} setData={setData} notify={notify} />
      </SettingsSection>

      <SettingsSection title="Data management" subtitle="Start fresh or factory reset" plain>
        <DataManagement data={data} setData={setData} notify={notify} />
      </SettingsSection>

      <button style={styles.btnPrimary} onClick={save}>Save settings</button>
    </div>
  );
}
// "Start afresh" tools. Start fresh clears every transaction and document but
// keeps your setup (company, banks, chart of accounts, rules, tax groups).
// Factory reset returns the whole app to the original demo state.
// Builds one consolidated, self-balancing journal entry from every entered
// account opening balance (everything except AR and AP, which are tracked
// per-customer/per-vendor instead so aging reports work correctly). Each
// account gets its own Debit and Credit box - the person fills in whichever
// side matches how they think of that balance, with no need to know an
// account's "normal" side. Whatever doesn't balance on its own is plugged
// against Opening Balance Equity, the standard clearing account for this.
function computeOpeningBalanceEntry(data) {
  const ob = data.openingBalances;
  const lines = [];
  let totalDebit = 0, totalCredit = 0;
  Object.entries(ob.accountAmounts).forEach(([accountId, amt]) => {
    const debit = Number(amt?.debit) || 0;
    const credit = Number(amt?.credit) || 0;
    totalDebit += debit; totalCredit += credit;
    const net = debit - credit;
    if (Math.abs(net) < 0.005) return;
    if (net > 0) lines.push({ accountId, debit: net, credit: 0 });
    else lines.push({ accountId, debit: 0, credit: -net });
  });
  const plug = totalDebit - totalCredit;
  if (Math.abs(plug) > 0.005) {
    if (plug > 0) lines.push({ accountId: "3900", debit: 0, credit: plug });
    else lines.push({ accountId: "3900", debit: -plug, credit: 0 });
  }
  return { lines, totalDebit, totalCredit, plug };
}
// Zoho-style Opening Balances: set a starting point for every account when
// migrating from another system. Accounts Receivable and Accounts Payable
// are deliberately excluded from the general list - they're entered per
// customer and per vendor instead, each becoming a real opening invoice or
// bill with its own due date, so AR/AP aging reports work correctly from
// day one rather than showing one unexplained lump sum.
// A plan configurator, not a payment processor - this browser-based app has
// no billing backend, so selecting a plan here saves your choice to
// settings and shows exactly what it would cost, ready for when a real
// payment provider is wired in behind hosting.
function SubscriptionPanel({ data, setData, notify }) {
  const { styles, theme, confirm } = useUI();
  const sub = data.settings.subscription || null;
  const status = subscriptionStatus(sub, todayStr());

  // No backend cron can expire a trial for us - so the moment this panel is
  // actually looked at past the trial's end date, formally revert to Free
  // rather than leaving a stale "trialing" record that no longer reflects
  // what's true.
  useEffect(() => {
    if (status.status === "expired") {
      setData(d => ({ ...d, settings: { ...d.settings, subscription: { planId: "free", status: "active", frequency: "monthly", extraUsers: 0, addOns: [] } } }));
      notify("Your free trial ended without confirming - you've been moved back to the Free plan");
    }
  }, [status.status]);

  const [planId, setPlanId] = useState(status.effectivePlanId !== "free" ? status.effectivePlanId : "growth");
  const [frequency, setFrequency] = useState(sub?.frequency || "monthly");
  const [extraUsers, setExtraUsers] = useState(sub?.extraUsers || 0);
  const [addOns, setAddOns] = useState(sub?.addOns || []);
  const [currencyCode, setCurrencyCode] = useState("USD");
  const [rate, setRate] = useState(1);
  const [rateSource, setRateSource] = useState(null); // "live" | "unavailable" | null (not fetched yet)
  const [rateAsOf, setRateAsOf] = useState(null);
  const [fetchingRate, setFetchingRate] = useState(false);

  const fetchLiveRate = async (code) => {
    if (code === "USD") { setRate(1); setRateSource(null); return; }
    setFetchingRate(true); setRateSource(null);
    try {
      const res = await fetch(`https://api.frankfurter.dev/v1/latest?base=USD&symbols=${code}`);
      const json = await res.json();
      const value = json?.rates?.[code];
      if (value) { setRate(value); setRateSource("live"); setRateAsOf(json.date || todayStr()); }
      else { setRateSource("unavailable"); }
    } catch {
      setRateSource("unavailable");
    } finally {
      setFetchingRate(false);
    }
  };
  useEffect(() => { fetchLiveRate(currencyCode); }, [currencyCode]);

  const allPlans = [FREE_PLAN, ...SUBSCRIPTION_PLANS];
  const plan = allPlans.find(p => p.id === planId) || FREE_PLAN;
  const price = computeSubscriptionPrice(plan, frequency, extraUsers, addOns);
  const currency = CURRENCIES.find(c => c.code === currencyCode) || CURRENCIES[1];
  const fmtPrice = (usdAmount) => `${currency.symbol}${(usdAmount * (currencyCode === "USD" ? 1 : rate)).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const includedAddOns = new Set(plan.includedAddOns || []);

  const isCurrentConfig = sub?.planId === planId && sub?.frequency === frequency && sub?.extraUsers === extraUsers && JSON.stringify([...(sub?.addOns || [])].sort()) === JSON.stringify([...addOns].sort());
  const isTrialingThis = status.status === "trialing" && isCurrentConfig;
  const isActiveThis = status.status === "active" && isCurrentConfig && sub?.planId !== "free";
  const isFreeSelected = planId === "free";

  const toggleAddOn = (id) => setAddOns(a => a.includes(id) ? a.filter(x => x !== id) : [...a, id]);

  const switchToFree = async () => {
    if (!(await confirm("Switch to the Free plan? Any active paid subscription or trial is cancelled immediately."))) return;
    setData(d => ({ ...d, settings: { ...d.settings, subscription: { planId: "free", status: "active", frequency: "monthly", extraUsers: 0, addOns: [] } } }));
    notify("Switched to the Free plan");
  };
  const startTrial = async () => {
    if (!(await confirm(`Start a ${TRIAL_DAYS}-day free trial of ${plan.name}? No charge today. If you don't confirm the subscription before the trial ends, you're moved back to the Free plan automatically.`))) return;
    const trialStartDate = todayStr();
    const end = new Date(trialStartDate); end.setDate(end.getDate() + TRIAL_DAYS);
    setData(d => ({ ...d, settings: { ...d.settings, subscription: { planId, frequency, extraUsers, addOns, status: "trialing", trialStartDate, trialEndDate: localDateToISO(end) } } }));
    notify(`${TRIAL_DAYS}-day free trial started`);
  };
  const confirmNow = async () => {
    if (!(await confirm(`Confirm ${plan.name} now at ${fmtPrice(price.totalForPeriod)} billed ${price.freq.label.toLowerCase()}? This ends the trial early and starts billing. No card is actually charged here, since this app has no payment backend yet - your selection is simply recorded as confirmed.`))) return;
    setData(d => ({ ...d, settings: { ...d.settings, subscription: { planId, frequency, extraUsers, addOns, status: "active" } } }));
    notify(`${plan.name} confirmed`);
  };
  const cancelTrial = async () => {
    if (!(await confirm("Cancel your trial and return to the Free plan now?"))) return;
    setData(d => ({ ...d, settings: { ...d.settings, subscription: { planId: "free", status: "active", frequency: "monthly", extraUsers: 0, addOns: [] } } }));
    notify("Trial cancelled - back on the Free plan");
  };

  return (
    <div>
      <div style={{ fontSize: 12.5, color: theme.muted, marginBottom: 14, display: "flex", alignItems: "center" }}>Plans and pricing<InfoTip text={`Introductory pricing - meaningfully below comparable accounting platforms while we build a customer base. Every paid plan offers an optional ${TRIAL_DAYS}-day free trial, or you can subscribe immediately - your choice. A trial left unconfirmed reverts automatically to Free.`} /></div>

      {status.status === "trialing" && (
        <div className="ces-card" style={{ ...styles.card, borderLeft: `3px solid ${theme.amber}` }}>
          <div style={{ fontSize: 13 }}>You're on a <strong>{TRIAL_DAYS}-day free trial</strong> of {allPlans.find(p => p.id === sub.planId)?.name} - <strong>{status.daysLeft} day{status.daysLeft === 1 ? "" : "s"} left</strong>. Confirm below to keep it after the trial ends, or it reverts to Free automatically.</div>
        </div>
      )}

      <div className="ces-card" style={styles.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <div style={styles.cardTitle}>Currency</div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <select style={styles.inputSmall} value={currencyCode} onChange={e => setCurrencyCode(e.target.value)}>
              {[{ code: "USD", symbol: "$" }, ...CURRENCIES.filter(c => c.code !== "USD")].map(c => <option key={c.code} value={c.code}>{c.code}</option>)}
            </select>
            {currencyCode !== "USD" && <>
              <span style={{ fontSize: 12, color: theme.muted }}>1 USD =</span>
              <input type="number" style={{ ...styles.inputSmall, width: 100 }} value={rate} onChange={e => { setRate(Number(e.target.value) || 1); setRateSource(rateSource === "live" ? "edited" : rateSource); }} />
              <span style={{ fontSize: 12, color: theme.muted }}>{currencyCode}</span>
              <button style={styles.iconBtn} onClick={() => fetchLiveRate(currencyCode)} disabled={fetchingRate}>{fetchingRate ? "..." : "↻"}</button>
            </>}
          </div>
        </div>
        {currencyCode !== "USD" && (
          <div style={{ fontSize: 11, color: rateSource === "unavailable" ? theme.amber : theme.muted, marginTop: 6 }}>
            {fetchingRate && "Fetching the current rate…"}
            {!fetchingRate && rateSource === "live" && `Live rate as of ${rateAsOf}, sourced from the European Central Bank via Frankfurter. Plans are billed in USD - this converts the display only.`}
            {!fetchingRate && rateSource === "edited" && `Rate adjusted manually. Plans are billed in USD - this converts the display only.`}
            {!fetchingRate && rateSource === "unavailable" && <>
              Couldn't fetch a live rate for {currencyCode} from the free rate source. Look it up at{" "}
              <a href={`https://www.xe.com/currencyconverter/convert/?Amount=1&From=USD&To=${currencyCode}`} target="_blank" rel="noopener noreferrer" style={{ color: theme.accent }}>xe.com</a>{" "}
              and enter it above.
            </>}
          </div>
        )}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 14, marginBottom: 20 }}>
        {allPlans.map(p => (
          <div key={p.id} className="ces-card" onClick={() => setPlanId(p.id)} style={{ ...styles.card, cursor: "pointer", marginBottom: 0, border: p.id === planId ? `2px solid ${theme.accent}` : styles.card.border, position: "relative" }}>
            <div style={{ fontFamily: "Inter, sans-serif", fontWeight: 700, fontSize: 16, color: theme.text }}>{p.name}</div>
            <div style={{ fontSize: 11.5, color: theme.muted, marginTop: 2, minHeight: 30 }}>{p.tagline}</div>
            <div style={{ marginTop: 10, display: "flex", alignItems: "baseline", gap: 4 }}>
              <span style={{ fontFamily: "Inter, sans-serif", fontWeight: 700, fontSize: 24, color: theme.text }}>{p.monthly === 0 ? fmtPrice(0) : fmtPrice(p.monthly)}</span>
              <span style={{ fontSize: 12, color: theme.muted }}>/mo</span>
            </div>
            <div style={{ fontSize: 11.5, color: theme.muted, marginTop: 2 }}>{p.includedUsers} user{p.includedUsers > 1 ? "s" : ""} included</div>
            <ul style={{ margin: "10px 0 0", paddingLeft: 16, fontSize: 11.5, color: theme.text }}>{p.features.map(f => <li key={f} style={{ marginBottom: 3 }}>{f}</li>)}</ul>
            {p.id === planId && <div style={{ position: "absolute", top: 12, right: 12, color: theme.accent, fontSize: 13 }}>✓</div>}
          </div>
        ))}
      </div>

      {!isFreeSelected && <>
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Billing frequency</div>
          <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
            {BILLING_FREQUENCIES.map(f => (
              <button key={f.id} onClick={() => setFrequency(f.id)} style={{ ...styles.btnGhost, ...(frequency === f.id ? { background: theme.accent, color: "#fff", borderColor: theme.accent } : {}) }}>
                {f.label}{f.discount > 0 && <span style={{ marginLeft: 5, fontSize: 10.5, opacity: 0.85 }}>save {(f.discount * 100).toFixed(0)}%</span>}
              </button>
            ))}
          </div>
        </div>

        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Add-ons</div>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 10, padding: "8px 0", borderBottom: `1px solid ${theme.border}` }}>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 13, fontWeight: 500 }}>Additional users</div>
              <div style={{ fontSize: 11, color: theme.muted }}>{plan.includedUsers} included with {plan.name} · {fmtPrice(SUBSCRIPTION_ADDONS.find(a => a.id === "extra-user").price * price.freq.months * (1 - price.freq.discount))} per extra user, billed {price.freq.label.toLowerCase()}{price.freq.months > 1 ? ` (${fmtPrice(SUBSCRIPTION_ADDONS.find(a => a.id === "extra-user").price)}/mo)` : ""}</div>
            </div>
            <button style={styles.iconBtn} onClick={() => setExtraUsers(n => Math.max(0, n - 1))}>−</button>
            <span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 14, minWidth: 20, textAlign: "center" }}>{extraUsers}</span>
            <button style={styles.iconBtn} onClick={() => setExtraUsers(n => n + 1)}>+</button>
          </div>
          {SUBSCRIPTION_ADDONS.filter(a => a.id !== "extra-user").map(a => {
            const included = includedAddOns.has(a.id);
            const periodPrice = a.price * price.freq.months * (1 - price.freq.discount);
            return (
              <label key={a.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 0", borderBottom: `1px solid ${theme.border}`, cursor: included ? "default" : "pointer" }}>
                <input type="checkbox" checked={included || addOns.includes(a.id)} disabled={included} onChange={() => toggleAddOn(a.id)} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 500 }}>{a.name}{included && <span style={{ marginLeft: 6, fontSize: 10.5, color: theme.emerald }}>included with {plan.name}</span>}</div>
                  {a.desc && <div style={{ fontSize: 11, color: theme.muted }}>{a.desc}</div>}
                </div>
                <div style={{ fontSize: 12.5, color: theme.muted, fontFamily: "Inter, sans-serif", textAlign: "right" }}>
                  {included ? "—" : price.freq.months === 1 ? `${fmtPrice(a.price)} ${a.unit}` : <>{fmtPrice(periodPrice)} billed {price.freq.label.toLowerCase()}<br /><span style={{ fontSize: 10.5 }}>({fmtPrice(a.price)}/mo)</span></>}
                </div>
              </label>
            );
          })}
        </div>

        <div className="ces-card" style={{ ...styles.card, borderLeft: `3px solid ${theme.accent}` }}>
          <div style={styles.cardTitle}>Summary</div>
          <RowLine label={`${plan.name} plan`} value={undefined} bold />
          <div style={{ fontSize: 12.5, color: theme.muted, marginTop: -6, marginBottom: 6 }}>{fmtPrice(plan.monthly)}/mo base{extraUsers > 0 ? ` + ${extraUsers} extra user${extraUsers > 1 ? "s" : ""} (${fmtPrice(price.extraUsersMonthly)}/mo)` : ""}{price.addOnMonthly > 0 ? ` + add-ons (${fmtPrice(price.addOnMonthly)}/mo)` : ""}</div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}><span>Billed {price.freq.label.toLowerCase()} ({price.freq.months} month{price.freq.months > 1 ? "s" : ""})</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmtPrice(price.fullPriceForPeriod)}</span></div>
          {price.discountAmount > 0 && <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0", color: theme.emerald }}><span>Discount ({(price.freq.discount * 100).toFixed(0)}%)</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>-{fmtPrice(price.discountAmount)}</span></div>}
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700, padding: "8px 0", borderTop: `1px solid ${theme.border}`, marginTop: 6 }}><span>Total per period</span><span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" }}>{fmtPrice(price.totalForPeriod)}</span></div>
          <div style={{ fontSize: 11.5, color: theme.muted, marginTop: 2 }}>Equivalent to {fmtPrice(price.effectiveMonthly)}/month</div>

          {isActiveThis ? (
            <button style={{ ...styles.btnPrimary, marginTop: 14 }} disabled>This is your current plan</button>
          ) : isTrialingThis ? (
            <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
              <button style={styles.btnPrimary} onClick={confirmNow}>Confirm & subscribe now</button>
              <button style={{ ...styles.btnGhost, color: theme.rose, borderColor: theme.rose }} onClick={cancelTrial}>Cancel trial</button>
            </div>
          ) : (
            <button style={{ ...styles.btnPrimary, marginTop: 14 }} onClick={startTrial}>Start {TRIAL_DAYS}-day free trial</button>
          )}
        </div>
      </>}
      {isFreeSelected && sub?.planId !== "free" && (
        <div className="ces-card" style={styles.card}>
          <button style={styles.btnPrimary} onClick={switchToFree}>Switch to Free</button>
        </div>
      )}
      {isFreeSelected && (!sub || sub.planId === "free") && (
        <div className="ces-card" style={styles.card}><div style={{ fontSize: 13, color: theme.muted }}>You're on the Free plan.</div></div>
      )}
    </div>
  );
}
function OpeningBalancesPanel({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const ob = data.openingBalances;
  const [draft, setDraft] = useState(ob.accountAmounts);
  const [asOfDate, setAsOfDate] = useState(ob.asOfDate);
  const [custForm, setCustForm] = useState({ customer: "", amount: "", dueDate: "" });
  const [vendForm, setVendForm] = useState({ vendor: "", amount: "", dueDate: "" });

  const eligible = data.accounts.filter(a => a.status !== "inactive" && !["1100", "2000", "3900"].includes(a.id));
  const groups = [
    ["Assets", eligible.filter(a => a.type === "asset")],
    ["Liabilities", eligible.filter(a => a.type === "liability")],
    ["Equity", eligible.filter(a => a.type === "equity")],
    ["Income", eligible.filter(a => a.type === "revenue")],
    ["Expense", eligible.filter(a => a.type === "expense")],
  ];
  const setAmt = (id, side, val) => setDraft(d => ({ ...d, [id]: { ...d[id], [side]: val } }));
  const previewData = { ...data, openingBalances: { ...ob, accountAmounts: draft } };
  const preview = computeOpeningBalanceEntry(previewData);

  const addCustomerRow = () => {
    if (!custForm.customer || !Number(custForm.amount)) return notify("Name the customer and enter an amount");
    setData(d => ({ ...d, openingBalances: { ...d.openingBalances, customerBalances: [...d.openingBalances.customerBalances, { id: uid("obc"), customer: custForm.customer, amount: Number(custForm.amount), dueDate: custForm.dueDate || asOfDate } ] } }));
    setCustForm({ customer: "", amount: "", dueDate: "" });
  };
  const addVendorRow = () => {
    if (!vendForm.vendor || !Number(vendForm.amount)) return notify("Name the vendor and enter an amount");
    setData(d => ({ ...d, openingBalances: { ...d.openingBalances, vendorBalances: [...d.openingBalances.vendorBalances, { id: uid("obv"), vendor: vendForm.vendor, amount: Number(vendForm.amount), dueDate: vendForm.dueDate || asOfDate } ] } }));
    setVendForm({ vendor: "", amount: "", dueDate: "" });
  };
  const removeCustomerRow = (id) => setData(d => ({ ...d, openingBalances: { ...d.openingBalances, customerBalances: d.openingBalances.customerBalances.filter(x => x.id !== id) } }));
  const removeVendorRow = (id) => setData(d => ({ ...d, openingBalances: { ...d.openingBalances, vendorBalances: d.openingBalances.vendorBalances.filter(x => x.id !== id) } }));

  const post = async () => {
    if (!(await confirm(`Post opening balances as of ${fmtDate(asOfDate)}? ${Math.abs(preview.plug) > 0.5 ? `${fmt(Math.abs(preview.plug))} will be plugged to Opening Balance Equity to keep the books balanced.` : "The entered amounts already balance exactly."} This creates real ledger entries${ob.customerBalances.length || ob.vendorBalances.length ? ", plus an opening invoice per customer and an opening bill per vendor" : ""}.`))) return;
    let newData = { ...data, openingBalances: { ...ob, asOfDate, accountAmounts: draft, posted: true, postedDate: todayStr() } };
    if (preview.lines.length > 0) {
      const txn = buildTxn("Opening balances", asOfDate, preview.lines, "opening-balance", "OB-MAIN");
      if (txn) newData = { ...newData, transactions: [...newData.transactions, txn] };
    }
    ob.customerBalances.forEach((c, i) => {
      const id = `OB-CUST-${i + 1}`;
      const txn = buildTxn(`Opening balance - ${c.customer}`, asOfDate, [{ accountId: "1100", debit: c.amount, credit: 0 }, { accountId: "3900", debit: 0, credit: c.amount }], "opening-balance", id);
      if (!txn) return;
      newData = {
        ...newData,
        transactions: [...newData.transactions, txn],
        invoices: [...newData.invoices, { id, customer: c.customer, date: asOfDate, dueDate: c.dueDate, items: [{ desc: "Opening balance", qty: 1, price: c.amount, inventoryId: "" }], taxes: [], status: "sent", amountPaid: 0, locked: false, isOpeningBalance: true }],
      };
    });
    ob.vendorBalances.forEach((v, i) => {
      const id = `OB-VEND-${i + 1}`;
      const txn = buildTxn(`Opening balance - ${v.vendor}`, asOfDate, [{ accountId: "3900", debit: v.amount, credit: 0 }, { accountId: "2000", debit: 0, credit: v.amount }], "opening-balance", id);
      if (!txn) return;
      newData = {
        ...newData,
        transactions: [...newData.transactions, txn],
        bills: [...newData.bills, { id, vendor: v.vendor, date: asOfDate, dueDate: v.dueDate, expenseAccountId: "", items: [{ desc: "Opening balance", qty: 1, price: v.amount, inventoryId: "" }], taxes: [], status: "unpaid", amountPaid: 0, locked: false, isOpeningBalance: true }],
      };
    });
    setData(newData);
    notify("Opening balances posted");
  };

  const unpost = async () => {
    if (!(await confirm("Unpost opening balances? All the journal entries and opening invoices/bills they created are removed, so you can edit the amounts and post again."))) return;
    setData(d => ({
      ...d,
      transactions: d.transactions.filter(t => t.source !== "opening-balance"),
      invoices: d.invoices.filter(i => !i.isOpeningBalance),
      bills: d.bills.filter(b => !b.isOpeningBalance),
      openingBalances: { ...d.openingBalances, posted: false, postedDate: null },
    }));
    notify("Opening balances unposted - amounts are still here to edit");
  };

  if (ob.posted) {
    return (
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Opening balances</div>
        <div style={{ fontSize: 13, color: theme.text, marginTop: 8 }}>Posted as of {fmtDate(ob.asOfDate)}{ob.postedDate ? ` (on ${fmtDate(ob.postedDate)})` : ""}.</div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>{Object.values(ob.accountAmounts).filter(v => Number(v?.debit) || Number(v?.credit)).length} account balance(s), {ob.customerBalances.length} customer opening invoice(s), {ob.vendorBalances.length} vendor opening bill(s).</div>
        <button style={{ ...styles.btnGhost, marginTop: 12 }} onClick={unpost}>Unpost & edit</button>
      </div>
    );
  }

  return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Opening balances<InfoTip width={300} text='Migrating from another system? Set where every account stood as of a specific date. Enter each balance in whichever column matches it - Debit or Credit - no need to work out which side is "normal" for that account. Receivables and payables are entered per customer and per vendor below, so aging reports work correctly from day one.' /></div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 12 }}>
        <span style={{ fontSize: 13 }}>As of date:</span>
        <input type="date" style={styles.input} value={asOfDate} onChange={e => setAsOfDate(e.target.value)} />
      </div>

      {groups.map(([label, accts]) => accts.length > 0 && (
        <div key={label} style={{ marginTop: 16 }}>
          <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>{label}</div>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}></th><th style={{ ...styles.th, textAlign: "right" }}>Debit</th><th style={{ ...styles.th, textAlign: "right" }}>Credit</th></tr></thead>
            <tbody>{accts.map(a => (
              <tr key={a.id}><td style={{ ...styles.td, width: "50%" }}>{a.code} {a.name}</td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 130, textAlign: "right" }} value={draft[a.id]?.debit || ""} placeholder="0" onChange={e => setAmt(a.id, "debit", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 130, textAlign: "right" }} value={draft[a.id]?.credit || ""} placeholder="0" onChange={e => setAmt(a.id, "credit", e.target.value)} /></td></tr>
            ))}</tbody>
          </table>
        </div>
      ))}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginTop: 16 }}>
        <Kpi label="Total debits" value={fmt(preview.totalDebit)} />
        <Kpi label="Total credits" value={fmt(preview.totalCredit)} />
        <Kpi label="To Opening Balance Equity" value={fmt(Math.abs(preview.plug))} tone={Math.abs(preview.plug) > 0.5 ? "amber" : "emerald"} />
      </div>

      <div style={{ marginTop: 20 }}>
        <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>Customer opening balances (Accounts Receivable)</div>
        {ob.customerBalances.length > 0 && (
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Customer</th><th style={{ ...styles.th, textAlign: "right" }}>Amount owed</th><th style={styles.th}>Due date</th><th style={styles.th}></th></tr></thead>
            <tbody>{ob.customerBalances.map(c => (
              <tr key={c.id}><td style={styles.td}>{c.customer}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(c.amount)}</td><td style={styles.tdMono}>{fmtDate(c.dueDate)}</td>
                <td style={styles.td}><button style={styles.iconBtn} onClick={() => removeCustomerRow(c.id)}>🗑</button></td></tr>
            ))}</tbody>
          </table>
        )}
        <div style={{ display: "flex", gap: 10, marginTop: 8, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Customer name" value={custForm.customer} onChange={e => setCustForm({ ...custForm, customer: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 140 }} placeholder="Amount owed" value={custForm.amount} onChange={e => setCustForm({ ...custForm, amount: e.target.value })} />
          <input type="date" style={styles.input} value={custForm.dueDate} onChange={e => setCustForm({ ...custForm, dueDate: e.target.value })} placeholder="Due date" />
          <button style={styles.btnGhost} onClick={addCustomerRow}>+ Add customer</button>
        </div>
      </div>

      <div style={{ marginTop: 20 }}>
        <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>Vendor opening balances (Accounts Payable)</div>
        {ob.vendorBalances.length > 0 && (
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Vendor</th><th style={{ ...styles.th, textAlign: "right" }}>Amount owed</th><th style={styles.th}>Due date</th><th style={styles.th}></th></tr></thead>
            <tbody>{ob.vendorBalances.map(v => (
              <tr key={v.id}><td style={styles.td}>{v.vendor}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(v.amount)}</td><td style={styles.tdMono}>{fmtDate(v.dueDate)}</td>
                <td style={styles.td}><button style={styles.iconBtn} onClick={() => removeVendorRow(v.id)}>🗑</button></td></tr>
            ))}</tbody>
          </table>
        )}
        <div style={{ display: "flex", gap: 10, marginTop: 8, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Vendor name" value={vendForm.vendor} onChange={e => setVendForm({ ...vendForm, vendor: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 140 }} placeholder="Amount owed" value={vendForm.amount} onChange={e => setVendForm({ ...vendForm, amount: e.target.value })} />
          <input type="date" style={styles.input} value={vendForm.dueDate} onChange={e => setVendForm({ ...vendForm, dueDate: e.target.value })} placeholder="Due date" />
          <button style={styles.btnGhost} onClick={addVendorRow}>+ Add vendor</button>
        </div>
      </div>

      <button style={{ ...styles.btnPrimary, marginTop: 20 }} onClick={post}>Post opening balances</button>
      <div style={{ fontSize: 11, color: theme.muted, marginTop: 8 }}>Set the "as of" date to just before you start using this system day-to-day - reports for periods after that date will be unaffected by this one-time setup entry.</div>
    </div>
  );
}
function DataManagement({ data, setData, notify }) {
  const { styles, theme, confirm } = useUI();
  const startFresh = async () => {
    if (!(await confirm("Start fresh? ALL transactions, invoices, bills, expenses, payments, sales and purchase documents, time sheets, production records, payroll runs, reconciliations, and fixed assets are permanently deleted. Your company settings, banks, chart of accounts, rules, projects, locations, departments, recurring schedules and tax groups are kept. This cannot be undone."))) return;
    if (!(await confirm("Are you absolutely sure? There is no way to recover the deleted records."))) return;
    setData(d => ({
      ...d,
      transactions: [], invoices: [], bills: [], expenses: [], payments: [],
      bankFeed: [], bin: [], inventoryLots: [], fixedAssets: [], budgets: {},
      inventory: d.inventory.map(i => ({ ...i, qty: 0 })),
      salesOrders: [], salesReceipts: [], creditNotes: [], salesReturns: [],
      purchaseOrders: [], purchaseReceipts: [], vendorCredits: [],
      timesheets: [], productionRecords: [], nextProductionNum: 1,
      payrollRuns: [], reconciliations: [], auditLog: [], taxProvisions: [],
      leases: [], deferredRevenueSchedules: [], provisions: [],
      openingBalances: { asOfDate: todayStr(), accountAmounts: {}, customerBalances: [], vendorBalances: [], posted: false, postedDate: null },
      nextInvoiceNum: 1001, nextBillNum: 2001, nextSalesOrderNum: 1, nextSalesReceiptNum: 1,
      nextCreditNoteNum: 1, nextSalesReturnNum: 1, nextPurchaseOrderNum: 1, nextPurchaseReceiptNum: 1, nextVendorCreditNum: 1,
      __auditLabel: "Start Fresh - cleared all transactions and documents",
    }));
    notify("All transactions cleared - you're starting fresh");
  };
  const factoryReset = async () => {
    if (!(await confirm("Factory reset? EVERYTHING - transactions, documents, banks, rules and settings you've customized - is permanently deleted. The chart of accounts resets to the standard starting list. This cannot be undone."))) return;
    const empty = seedData(); // used only for its structural chart-of-accounts skeleton, never for demo transactions
    setData({
      settings: { ...empty.settings, companyName: data.settings.companyName },
      accounts: empty.accounts,
      banks: [], transactions: [], invoices: [], bills: [], expenses: [], inventory: [], inventoryLots: [],
      fixedAssets: [], payments: [], taxGroups: [], projects: [], locations: [], departments: [],
      budgets: {}, budgetAccounts: [], favoriteReports: empty.favoriteReports,
      bankFeed: [], categoryRules: [], bin: [], employees: [], payrollRuns: [], recurringJournals: [],
      reconciliations: [], auditLog: [], taxProvisions: [], leases: [], deferredRevenueSchedules: [], provisions: [],
      salesOrders: [], salesReceipts: [], creditNotes: [], salesReturns: [], purchaseOrders: [], purchaseReceipts: [],
      vendorCredits: [], timesheets: [], productionRecords: [], nextProductionNum: 1,
      openingBalances: { asOfDate: todayStr(), accountAmounts: {}, customerBalances: [], vendorBalances: [], posted: false, postedDate: null },
      nextInvoiceNum: 1001, nextBillNum: 2001, nextSalesOrderNum: 1, nextSalesReceiptNum: 1, nextCreditNoteNum: 1,
      nextSalesReturnNum: 1, nextPurchaseOrderNum: 1, nextPurchaseReceiptNum: 1, nextVendorCreditNum: 1,
      __auditLabel: "Factory reset - cleared to an empty company",
    });
    notify("Reset to an empty company");
  };
  return (
    <div className="ces-card" style={{ ...styles.card, borderLeft: `3px solid ${theme.rose}` }}>
      <div style={styles.cardTitle}>Data management</div>
      <div style={{ fontSize: 12.5, color: theme.muted, marginTop: 4 }}>Both actions are permanent. "Start fresh" keeps your setup (company, banks, chart of accounts, rules, tax groups, projects, locations, departments, inventory items at zero stock) and deletes every transaction, document and activity record. "Factory reset" clears everything, including your setup, back to a blank company.</div>
      <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
        <button style={{ ...styles.btnGhost, color: theme.rose, borderColor: theme.rose }} onClick={startFresh}>Start fresh - delete all transactions</button>
        <button style={styles.btnGhost} onClick={factoryReset}>Factory reset - clear to a blank company</button>
      </div>
    </div>
  );
}
// Two simple named lists - tag fixed assets and inventory to a location
// and/or department, so "what do we have at the Lagos office" or "what's
// assigned to Operations" is a real, answerable question rather than
// something you have to reconstruct by memory.
function LocationsAndDepartments({ data, setData, notify }) {
  const { styles, theme } = useUI();
  const [locName, setLocName] = useState("");
  const [deptName, setDeptName] = useState("");
  const addLocation = () => {
    if (!locName.trim()) return notify("Name the location");
    setData(d => ({ ...d, locations: [...d.locations, { id: uid("loc"), name: locName.trim() }] }));
    setLocName("");
  };
  const addDepartment = () => {
    if (!deptName.trim()) return notify("Name the department");
    setData(d => ({ ...d, departments: [...d.departments, { id: uid("dept"), name: deptName.trim() }] }));
    setDeptName("");
  };
  const removeLocation = (id) => setData(d => ({ ...d, locations: d.locations.filter(l => l.id !== id) }));
  const removeDepartment = (id) => setData(d => ({ ...d, departments: d.departments.filter(x => x.id !== id) }));

  return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Locations and departments</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Tag fixed assets and inventory to a location and a department, so you can see what's where and whose it is.</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 20, marginTop: 14 }}>
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>Locations</div>
          {data.locations.map(l => (
            <div key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: `1px solid ${theme.border}` }}>
              <span style={{ fontSize: 13 }}>{l.name}</span>
              <button style={styles.iconBtn} onClick={() => removeLocation(l.id)}>✕</button>
            </div>
          ))}
          {data.locations.length === 0 && <div style={{ fontSize: 12, color: theme.muted, padding: "4px 0" }}>None yet.</div>}
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <input style={{ ...styles.input, flex: 1 }} placeholder="e.g. Lagos Office" value={locName} onChange={e => setLocName(e.target.value)} />
            <button style={styles.btnGhost} onClick={addLocation}>Add</button>
          </div>
        </div>
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>Departments</div>
          {data.departments.map(x => (
            <div key={x.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: `1px solid ${theme.border}` }}>
              <span style={{ fontSize: 13 }}>{x.name}</span>
              <button style={styles.iconBtn} onClick={() => removeDepartment(x.id)}>✕</button>
            </div>
          ))}
          {data.departments.length === 0 && <div style={{ fontSize: 12, color: theme.muted, padding: "4px 0" }}>None yet.</div>}
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <input style={{ ...styles.input, flex: 1 }} placeholder="e.g. Operations" value={deptName} onChange={e => setDeptName(e.target.value)} />
            <button style={styles.btnGhost} onClick={addDepartment}>Add</button>
          </div>
        </div>
      </div>
    </div>
  );
}
function ProjectsManager({ data, setData, notify }) {
  const { styles, fmt, theme } = useUI();
  const [form, setForm] = useState({ name: "", client: "", budget: "" });
  const addProject = () => {
    if (!form.name) return notify("Name the project");
    setData(d => ({ ...d, projects: [...d.projects, { id: uid("proj"), name: form.name, client: form.client, budget: Number(form.budget) || 0 }] }));
    setForm({ name: "", client: "", budget: "" });
    notify("Project added - tag it on invoices, bills and expenses");
  };
  const removeProject = (id) => setData(d => ({ ...d, projects: d.projects.filter(p => p.id !== id) }));
  return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Projects</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Tag invoices, bills and quick expenses to a project to track its revenue, cost and margin in the Reports tab.</div>
      {data.projects.length > 0 && (
        <table style={{ ...styles.table, marginTop: 12 }}>
          <thead><tr><th style={styles.th}>Project</th><th style={styles.th}>Client</th><th style={{ ...styles.th, textAlign: "right" }}>Budget</th><th></th></tr></thead>
          <tbody>{data.projects.map(p => (<tr key={p.id}><td style={styles.td}>{p.name}</td><td style={styles.td}>{p.client || "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{p.budget ? fmt(p.budget) : "-"}</td><td style={styles.td}><button style={styles.iconBtn} onClick={() => removeProject(p.id)}>✕ remove</button></td></tr>))}</tbody>
        </table>
      )}
      <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
        <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Project name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
        <input style={styles.input} placeholder="Client (optional)" value={form.client} onChange={e => setForm({ ...form, client: e.target.value })} />
        <input type="number" style={{ ...styles.input, width: 130 }} placeholder="Budget (optional)" value={form.budget} onChange={e => setForm({ ...form, budget: e.target.value })} />
        <button style={styles.btnPrimary} onClick={addProject}>Add project</button>
      </div>
    </div>
  );
}
function TaxGroupsManager({ data, setData, notify }) {
  const { styles, theme } = useUI();
  const taxAccounts = data.accounts.filter(a => a.type === "liability" || a.type === "asset");
  const blankGroup = () => ({ name: "", components: [{ id: uid("c"), name: "VAT", rate: 7.5, mode: "percent", effect: "deduct", accountId: "2100" }, { id: uid("c"), name: "WHT", rate: 5, mode: "percent", effect: "deduct", accountId: "2200" }] });
  const [draft, setDraft] = useState(blankGroup());
  const [showNew, setShowNew] = useState(false);

  const updateComp = (i, field, val) => setDraft(d => ({ ...d, components: d.components.map((c, idx) => idx === i ? { ...c, [field]: val } : c) }));
  const addComp = () => setDraft(d => ({ ...d, components: [...d.components, { id: uid("c"), name: "", rate: 0, mode: "percent", effect: "deduct", accountId: taxAccounts[0]?.id }] }));
  const removeComp = (i) => setDraft(d => ({ ...d, components: d.components.filter((_, idx) => idx !== i) }));

  const saveGroup = () => {
    if (!draft.name || draft.components.some(c => !c.name)) return notify("Name the group and every tax inside it");
    setData(d => ({ ...d, taxGroups: [...d.taxGroups, { id: uid("grp"), name: draft.name, components: draft.components }] }));
    setDraft(blankGroup()); setShowNew(false);
    notify("Tax group saved - it's now selectable on any invoice or bill line item");
  };
  const removeGroup = (id) => setData(d => ({ ...d, taxGroups: d.taxGroups.filter(g => g.id !== id) }));

  return (
    <div className="ces-card" style={styles.card}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={styles.cardTitle}>Tax groups</div>
        <button style={styles.btnGhost} onClick={() => setShowNew(s => !s)}>{showNew ? "Cancel" : "+ New group"}</button>
      </div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>
        A group bundles two or more taxes that are each calculated on the SAME line-item amount in parallel - e.g. VAT and WHT both computed on 500,000 independently, then combined - rather than cascading on top of one another. Groups appear as a "Tax group" choice on each invoice/bill line item.
      </div>

      {data.taxGroups.length > 0 && (
        <div style={{ marginTop: 12 }}>
          {data.taxGroups.map(g => (
            <div key={g.id} style={{ border: `1px solid ${theme.border}`, borderRadius: 8, padding: "10px 14px", marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ fontWeight: 600, fontSize: 13.5 }}>{g.name}</div>
                <button style={styles.iconBtn} onClick={() => removeGroup(g.id)}>✕ remove</button>
              </div>
              <div style={{ fontSize: 12.5, color: theme.muted, marginTop: 4 }}>
                {g.components.map(c => `${c.name} (${c.mode === "fixed" ? "flat" : c.rate + "%"}, ${c.effect})`).join("  ·  ")}
              </div>
            </div>
          ))}
        </div>
      )}

      {showNew && (
        <div style={{ marginTop: 14, background: theme.panel2, border: `1px solid ${theme.border}`, borderRadius: 8, padding: 14 }}>
          <input style={{ ...styles.input, width: "100%" }} placeholder="Group name (e.g. VAT + WHT)" value={draft.name} onChange={e => setDraft({ ...draft, name: e.target.value })} />
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}>Tax</th><th style={{ ...styles.th, textAlign: "right" }}>Rate</th><th style={styles.th}>Mode</th><th style={styles.th}>Effect</th><th style={styles.th}>Posts to</th><th></th></tr></thead>
            <tbody>{draft.components.map((c, i) => (
              <tr key={c.id}>
                <td style={styles.td}><input style={{ ...styles.inputSmall, width: 100 }} value={c.name} onChange={e => updateComp(i, "name", e.target.value)} /></td>
                <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={c.rate} onChange={e => updateComp(i, "rate", e.target.value)} /></td>
                <td style={styles.td}><select style={styles.inputSmall} value={c.mode} onChange={e => updateComp(i, "mode", e.target.value)}><option value="percent">%</option><option value="fixed">flat</option></select></td>
                <td style={styles.td}><select style={styles.inputSmall} value={c.effect} onChange={e => updateComp(i, "effect", e.target.value)}><option value="deduct">deduct</option><option value="add">add</option></select></td>
                <td style={styles.td}><select style={styles.inputSmall} value={c.accountId} onChange={e => updateComp(i, "accountId", e.target.value)}>{taxAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select></td>
                <td style={styles.td}><button style={styles.iconBtn} onClick={() => removeComp(i)}>✕</button></td>
              </tr>
            ))}</tbody>
          </table>
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <button style={styles.btnGhost} onClick={addComp}>+ tax to this group</button>
            <button style={styles.btnPrimary} onClick={saveGroup}>Save group</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* =================================== styles =================================== */
function globalCss(theme) {
  return `
* { box-sizing: border-box; }
html { -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; text-rendering: optimizeLegibility; }
body { font-size: 14px; line-height: 1.5; }
input, select, button, textarea { font-family: Inter, sans-serif; font-size: 13.5px; }
h1,h2,h3 { letter-spacing: -0.01em; }
input:focus, select:focus { outline: 2px solid ${theme.accent}; outline-offset: 1px; }
button:focus-visible { outline: 2px solid ${theme.accent}; outline-offset: 2px; }
table { border-collapse: collapse; width: 100%; }
tbody tr { transition: background 0.12s ease; }
tbody tr:hover { background: ${theme.panel2}; }
.nav-item:hover { background: ${theme.mode === "dark" ? "rgba(255,255,255,0.06)" : theme.panel2} !important; }
.drill-row:hover { background: ${hexToRgba(theme.accent, theme.mode === "dark" ? 0.12 : 0.06)}; }
/* Sidebar nav items lift and tilt slightly toward the cursor on hover.
   Kept the rotation angle small (was 16px translateZ / large scale before,
   which is what caused the blur) and added backface-visibility + font
   smoothing hints, which is what actually keeps text and icons crisp
   through a 3D transform - the 3D effect itself was never the problem. */
.nav-item {
  transform: perspective(700px) translateZ(0) scale(1);
  transition: transform 0.3s cubic-bezier(0.22, 1, 0.36, 1), background 0.15s ease, box-shadow 0.35s ease;
  backface-visibility: hidden; -webkit-font-smoothing: antialiased; will-change: transform;
}
.nav-item:hover {
  transform: perspective(700px) translateZ(6px) scale(1.02);
  transition: transform 0.15s cubic-bezier(0.22, 1, 0.36, 1), background 0.1s ease, box-shadow 0.15s ease;
  box-shadow: 0 5px 14px ${hexToRgba(theme.accent, 0.22)};
  position: relative; z-index: 5;
}
/* Cursor spotlight sheen + gentle lift on cards (Linear-style) */
.ces-card {
  transition: box-shadow 0.18s ease, transform 0.18s ease, border-color 0.18s ease;
}
.ces-card:hover { transform: translateY(-1px); box-shadow: ${theme.shadowLg}; border-color: ${shade(theme.accent, 55)} !important; }
/* Dashboard summary items (KPIs, insight cards) tilt toward the cursor in
   3D on hover. Same hardening as the sidebar: smaller angles than the
   original, backface-visibility hidden, font smoothing on, and the
   rotation value itself is now throttled to one update per animation
   frame (see the mousemove handler) instead of firing on every raw
   mouse event, which was the main source of the blur/jank before. */
.ces-bubble {
  transform: perspective(900px) rotateX(0deg) rotateY(0deg) scale(1) translateZ(0);
  transition: transform 0.3s cubic-bezier(0.22, 1, 0.36, 1), box-shadow 0.35s ease, border-color 0.18s ease;
  backface-visibility: hidden; -webkit-font-smoothing: antialiased; will-change: transform;
}
.ces-bubble:hover {
  transform: perspective(900px) rotateX(var(--rx, 0deg)) rotateY(var(--ry, 0deg)) scale(1.015) translateZ(0) translateY(-3px);
  transition: transform 0.12s cubic-bezier(0.22, 1, 0.36, 1), box-shadow 0.2s ease, border-color 0.18s ease;
  box-shadow: 0 8px 20px ${hexToRgba(theme.accent, 0.14)}, ${theme.shadowLg};
  border-color: ${shade(theme.accent, 55)} !important;
}
.ces-card::after {
  content: ""; position: absolute; inset: 0; border-radius: inherit; pointer-events: none;
  opacity: 0; transition: opacity 0.25s ease;
  background: radial-gradient(300px circle at var(--mx, 50%) var(--my, 50%), ${hexToRgba(theme.accent, theme.mode === "dark" ? 0.08 : 0.045)}, transparent 60%);
}
.ces-card:hover::after { opacity: 1; }
.ces-logo {
  position: relative; display: inline-flex; align-items: center; justify-content: center;
  background: linear-gradient(155deg, #ffffff, #eef1f8);
  border-radius: 16px; overflow: hidden; isolation: isolate;
  box-shadow: 0 1px 1px rgba(16,24,40,0.06), 0 4px 10px rgba(16,24,40,0.10), 0 14px 28px rgba(16,24,40,0.14), inset 0 1px 0 rgba(255,255,255,0.9);
  transform: perspective(700px) rotateX(0deg) rotateY(0deg);
  transition: transform 0.25s ease-out, box-shadow 0.2s ease;
  backface-visibility: hidden; -webkit-font-smoothing: antialiased; will-change: transform;
}
.ces-logo:hover {
  transform: perspective(700px) rotateX(var(--rx, 0deg)) rotateY(var(--ry, 0deg));
  transition: transform 0.1s ease-out, box-shadow 0.2s ease;
  box-shadow: 0 2px 2px rgba(16,24,40,0.08), 0 8px 18px rgba(16,24,40,0.14), 0 24px 44px rgba(16,24,40,0.20), inset 0 1px 0 rgba(255,255,255,0.9);
}
.ces-logo::before {
  content: ""; position: absolute; inset: 0; z-index: 2; pointer-events: none; opacity: 0; transition: opacity 0.2s ease;
  background: radial-gradient(100px circle at var(--lx, 30%) var(--ly, 20%), rgba(255,255,255,0.5), transparent 60%);
  mix-blend-mode: overlay;
}
.ces-logo:hover::before { opacity: 1; }
.ces-logo img { position: relative; z-index: 1; display: block; image-rendering: -webkit-optimize-contrast; }
button { transition: opacity 0.12s ease, background 0.12s ease, box-shadow 0.12s ease; }
button:hover { opacity: 0.92; }
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-thumb { background: ${theme.border}; border-radius: 4px; }
option { color: #000; }
@media print {
  body * { visibility: hidden; }
  .print-area, .print-area * { visibility: visible; }
  .print-area { position: absolute; top: 0; left: 0; width: 100%; box-shadow: none !important; border: none !important; background: #fff !important; }
  .print-area * { color: #1B2559 !important; background: transparent !important; }
  .print-area img { visibility: visible !important; }
  .no-print { display: none !important; }
}
`;
}
/* =================================== Customers & Vendors =================================== */
// Derived directories: parties are read from documents, so the list is always
// consistent with what has actually been billed.
function PartyDirectory({ title, eyebrow, rows, docLabel, data, renderDocs }) {
  const { styles, fmt, theme } = useUI();
  const [openParty, setOpenParty] = useState(null);
  return (
    <div>
      <PageHeader eyebrow={eyebrow} title={title} sub={`${rows.length} on record`} />
      <div className="ces-card" style={styles.card}>
        <table style={styles.table}>
          <thead><tr><th style={styles.th}>Name</th><th style={{ ...styles.th, textAlign: "right" }}>{docLabel}</th><th style={{ ...styles.th, textAlign: "right" }}>Total value</th><th style={{ ...styles.th, textAlign: "right" }}>Outstanding</th></tr></thead>
          <tbody>{rows.map(r => (
            <React.Fragment key={r.name}>
              <tr onClick={() => setOpenParty(openParty === r.name ? null : r.name)} style={{ cursor: "pointer" }}>
                <td style={{ ...styles.td, fontWeight: 600 }}>{r.name}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{r.count}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.total)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", color: r.outstanding > 0.5 ? theme.amber : theme.emerald }}>{fmt(r.outstanding)}</td>
              </tr>
              {openParty === r.name && <tr><td colSpan={4} style={{ ...styles.td, background: theme.panel2 }}>{renderDocs(r.name)}</td></tr>}
            </React.Fragment>
          ))}</tbody>
        </table>
        {rows.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>Nothing here yet - parties appear automatically as you raise documents.</div>}
      </div>
    </div>
  );
}
// IFRS 9 simplified approach: a provision matrix applying a loss rate per
// aging bucket to outstanding receivables, posted as a movement against the
// current Allowance for Doubtful Accounts balance.
function ECLPanel({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const ecl = computeECL(data);
  const balances = computeBalances(data.accounts, data.transactions);
  const recognized = balances["1150"] || 0;
  const movement = ecl.targetProvision - recognized;

  const setRate = (bucket, val) => setData(d => ({ ...d, settings: { ...d.settings, eclRates: { ...d.settings.eclRates, [bucket]: Number(val) || 0 } } }));

  const post = async () => {
    if (Math.abs(movement) < 0.5) return notify("Provision is already at the target level - nothing to post");
    if (!(await confirm(`Post an expected credit loss ${movement > 0 ? "increase" : "decrease"} of ${fmt(Math.abs(movement))}? Adjusts the Allowance for Doubtful Accounts to match the provision matrix below.`))) return;
    const lines = movement > 0
      ? [{ accountId: "5950", debit: movement, credit: 0 }, { accountId: "1150", debit: 0, credit: movement }]
      : [{ accountId: "1150", debit: -movement, credit: 0 }, { accountId: "5950", debit: 0, credit: -movement }];
    const txn = buildTxn("Expected credit loss provisioning", todayStr(), lines, "manual");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn] }));
    notify("ECL provision posted");
  };

  return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Expected credit loss provision</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>A loss rate applied to each aging bucket of outstanding invoices, using the simplified approach for trade receivables. Adjust the rates to match your own collection experience.</div>
      <table style={{ ...styles.table, marginTop: 10 }}>
        <thead><tr><th style={styles.th}>Bucket</th><th style={{ ...styles.th, textAlign: "right" }}>Outstanding</th><th style={{ ...styles.th, textAlign: "right" }}>Loss rate</th><th style={{ ...styles.th, textAlign: "right" }}>Provision</th></tr></thead>
        <tbody>{ecl.rows.map(r => (
          <tr key={r.bucket}><td style={styles.td}>{r.bucket}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.outstanding)}</td>
            <td style={{ ...styles.td, textAlign: "right" }}><input type="number" style={{ ...styles.inputSmall, width: 70, textAlign: "right" }} value={r.rate} onChange={e => setRate(r.bucket, e.target.value)} />%</td>
            <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.provision)}</td></tr>
        ))}</tbody>
      </table>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginTop: 14 }}>
        <Kpi label="Total outstanding" value={fmt(ecl.totalOutstanding)} />
        <Kpi label="Target provision" value={fmt(ecl.targetProvision)} tone="amber" />
        <Kpi label="Currently recognized" value={fmt(recognized)} />
        <Kpi label="Adjustment needed" value={fmt(movement)} tone={Math.abs(movement) < 0.5 ? "emerald" : "amber"} />
      </div>
      <button style={{ ...styles.btnPrimary, marginTop: 14 }} onClick={post}>Post ECL adjustment</button>
    </div>
  );
}
function Customers({ data, setData, notify }) {
  const { styles, fmt } = useUI();
  const byName = {};
  data.invoices.forEach(i => {
    const t = computeDocTotals(i, data.taxGroups).finalAmount;
    if (!byName[i.customer]) byName[i.customer] = { name: i.customer, count: 0, total: 0, outstanding: 0 };
    byName[i.customer].count++; byName[i.customer].total += t; byName[i.customer].outstanding += Math.max(0, t - (i.amountPaid || 0));
  });
  const rows = Object.values(byName).sort((a, b) => b.total - a.total);
  const renderDocs = (name) => (
    <table style={styles.table}><tbody>
      {data.invoices.filter(i => i.customer === name).map(i => {
        const t = computeDocTotals(i, data.taxGroups).finalAmount;
        return <tr key={i.id}><td style={styles.tdMono}>{i.id}</td><td style={styles.tdMono}>{fmtDate(i.date)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(i.amountPaid || 0)} paid</td></tr>;
      })}
    </tbody></table>
  );
  return (
    <div>
      <PartyDirectory title="Customers" eyebrow="Sales" rows={rows} docLabel="Invoices" data={data} renderDocs={renderDocs} />
      <ECLPanel data={data} setData={setData} notify={notify} />
    </div>
  );
}
function Vendors({ data }) {
  const { styles, fmt } = useUI();
  const byName = {};
  data.bills.forEach(b => {
    const t = computeDocTotals(b, data.taxGroups).finalAmount;
    if (!byName[b.vendor]) byName[b.vendor] = { name: b.vendor, count: 0, total: 0, outstanding: 0 };
    byName[b.vendor].count++; byName[b.vendor].total += t; byName[b.vendor].outstanding += Math.max(0, t - (b.amountPaid || 0));
  });
  data.expenses.forEach(e => {
    if (!byName[e.vendor]) byName[e.vendor] = { name: e.vendor, count: 0, total: 0, outstanding: 0 };
    byName[e.vendor].count++; byName[e.vendor].total += e.amount;
  });
  const rows = Object.values(byName).sort((a, b) => b.total - a.total);
  const renderDocs = (name) => (
    <table style={styles.table}><tbody>
      {data.bills.filter(b => b.vendor === name).map(b => {
        const t = computeDocTotals(b, data.taxGroups).finalAmount;
        return <tr key={b.id}><td style={styles.tdMono}>{b.id}</td><td style={styles.tdMono}>{fmtDate(b.date)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(t)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(b.amountPaid || 0)} paid</td></tr>;
      })}
      {data.expenses.filter(e => e.vendor === name).map(e => (
        <tr key={e.id}><td style={styles.tdMono}>expense</td><td style={styles.tdMono}>{fmtDate(e.date)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(e.amount)}</td><td style={styles.td}></td></tr>
      ))}
    </tbody></table>
  );
  return <PartyDirectory title="Vendors" eyebrow="Purchases" rows={rows} docLabel="Bills / expenses" data={data} renderDocs={renderDocs} />;
}

/* =================================== Payroll =================================== */
function Payroll({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [form, setForm] = useState({ name: "", role: "", grossMonthly: "", bankId: data.banks[0]?.id });
  const [openEmp, setOpenEmp] = useState(null);
  const [openRun, setOpenRun] = useState(null);
  const [dedForm, setDedForm] = useState({ name: "", amount: "" });
  const currentMonth = monthKey(todayStr());

  const addEmployee = () => {
    if (!form.name || !Number(form.grossMonthly)) return notify("Enter a name and gross monthly pay");
    setData(d => ({ ...d, employees: [...d.employees, { id: uid("emp"), name: form.name, role: form.role, grossMonthly: Number(form.grossMonthly), pension: true, nhf: false, deductions: [] }] }));
    setForm({ ...form, name: "", role: "", grossMonthly: "" });
    notify("Employee added - statutory deductions apply automatically");
  };
  const removeEmployee = async (emp) => {
    if (!(await confirm(`Remove ${emp.name} from payroll? Past payroll runs stay on the ledger.`))) return;
    setData(d => ({ ...d, employees: d.employees.filter(e => e.id !== emp.id) }));
    notify("Removed");
  };
  const patchEmp = (id, patch) => setData(d => ({ ...d, employees: d.employees.map(e => e.id === id ? { ...e, ...patch } : e) }));
  const addDeduction = (emp) => {
    if (!dedForm.name || !Number(dedForm.amount)) return notify("Name the deduction and enter a monthly amount");
    patchEmp(emp.id, { deductions: [...(emp.deductions || []), { id: uid("ded"), name: dedForm.name, amount: Number(dedForm.amount) }] });
    setDedForm({ name: "", amount: "" });
    notify("Deduction added");
  };
  const removeDeduction = (emp, ded) => patchEmp(emp.id, { deductions: emp.deductions.filter(x => x.id !== ded.id) });

  const calcs = data.employees.map(e => ({ emp: e, c: computeNigeriaPayroll(e) }));
  const totals = calcs.reduce((s, { c }) => ({ gross: s.gross + c.grossM, paye: s.paye + c.payeM, pension: s.pension + c.pensionM, nhf: s.nhf + c.nhfM, custom: s.custom + c.customM, net: s.net + c.netM }), { gross: 0, paye: 0, pension: 0, nhf: 0, custom: 0, net: 0 });
  const alreadyRun = data.payrollRuns.some(r => r.month === currentMonth);

  const runPayroll = async () => {
    if (data.employees.length === 0) return notify("Add employees first");
    if (alreadyRun) return notify(`Payroll for ${currentMonth} has already been run`);
    const bank = data.banks.find(b => b.id === form.bankId);
    if (!bank) return notify("Choose the bank to pay from");
    if (!(await confirm(`Run ${currentMonth} payroll? Gross ${fmt(totals.gross)}, net pay ${fmt(totals.net)} from ${bank.name}; PAYE ${fmt(totals.paye)}, pension ${fmt(totals.pension)}${totals.nhf > 0.5 ? `, NHF ${fmt(totals.nhf)}` : ""}${totals.custom > 0.5 ? `, other deductions ${fmt(totals.custom)}` : ""} accrue as payables.`))) return;
    const r2 = (x) => Math.round(x * 100) / 100;
    const lines = [
      { accountId: "5300", debit: r2(totals.gross), credit: 0 },
      { accountId: bank.accountId, debit: 0, credit: r2(totals.net) },
      ...(totals.paye > 0.005 ? [{ accountId: "2210", debit: 0, credit: r2(totals.paye) }] : []),
      ...(totals.pension > 0.005 ? [{ accountId: "2220", debit: 0, credit: r2(totals.pension) }] : []),
      ...(totals.nhf > 0.005 ? [{ accountId: "2230", debit: 0, credit: r2(totals.nhf) }] : []),
      ...(totals.custom > 0.005 ? [{ accountId: "2240", debit: 0, credit: r2(totals.custom) }] : []),
    ];
    // Guard rounding: force balance by adjusting net (bank) line
    const dr = lines.reduce((s, l) => s + l.debit, 0), cr = lines.reduce((s, l) => s + l.credit, 0);
    if (Math.abs(dr - cr) > 0.001) lines[1].credit = r2(lines[1].credit + (dr - cr));
    const txn = buildTxn(`Payroll - ${currentMonth}`, todayStr(), lines, "payroll");
    if (!txn) return notify("Entry not balanced");
    const snapshot = calcs.map(({ emp, c }) => ({ name: emp.name, role: emp.role, gross: c.grossM, paye: c.payeM, pension: c.pensionM, nhf: c.nhfM, custom: c.customM, net: c.netM, deductions: (emp.deductions || []).map(x => ({ ...x })) }));
    setData(d => ({ ...d, transactions: [...d.transactions, txn], payrollRuns: [...d.payrollRuns, { id: uid("run"), month: currentMonth, date: todayStr(), gross: totals.gross, paye: totals.paye, pension: totals.pension, nhf: totals.nhf, custom: totals.custom, net: totals.net, total: totals.gross, employeeCount: d.employees.length, txnId: txn.id, bankId: bank.id, snapshot }] }));
    notify(`Payroll posted: net ${fmt(totals.net)} paid, statutory deductions accrued`);
  };
  const deleteRun = async (run) => {
    if (!(await confirm(`Delete the ${run.month} payroll run? Its journal entry (gross ${fmt(run.gross || run.total)}) is removed.`))) return;
    setData(d => ({ ...d, payrollRuns: d.payrollRuns.filter(r => r.id !== run.id), transactions: d.transactions.filter(t => t.id !== run.txnId) }));
    notify("Payroll run deleted and its ledger entry removed");
  };

  return (
    <div>
      <PageHeader eyebrow="People" title="Payroll" sub={`${data.employees.length} employees \u00b7 Nigeria PAYE (PITA bands)`} action={
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <select style={styles.input} value={form.bankId} onChange={e => setForm({ ...form, bankId: e.target.value })}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
          <button style={{ ...styles.btnPrimary, opacity: alreadyRun ? 0.5 : 1 }} onClick={runPayroll}>{alreadyRun ? `${currentMonth} run \u2713` : `Run ${currentMonth} payroll`}</button>
        </div>
      } />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 18 }}>
        <Kpi label="Total gross" value={fmt(totals.gross)} />
        <Kpi label="PAYE (monthly)" value={fmt(totals.paye)} tone="amber" />
        <Kpi label="Pension 8%" value={fmt(totals.pension)} tone="amber" />
        <Kpi label="Other deductions" value={fmt(totals.nhf + totals.custom)} tone="amber" />
        <Kpi label="Net pay" value={fmt(totals.net)} tone="emerald" />
      </div>

      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Employees</div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Net pay is computed automatically: gross less pension (8%), NHF (2.5%, optional), PAYE per the graduated PITA bands with the consolidated relief allowance, and any custom deductions. Click an employee to view and adjust the breakdown.</div>
        <div style={{ overflowX: "auto" }}>
        <table style={{ ...styles.table, marginTop: 10 }}>
          <thead><tr><th style={styles.th}>Employee</th><th style={styles.th}>Role</th><th style={{ ...styles.th, textAlign: "right" }}>Gross</th><th style={{ ...styles.th, textAlign: "right" }}>PAYE</th><th style={{ ...styles.th, textAlign: "right" }}>Pension</th><th style={{ ...styles.th, textAlign: "right" }}>Other</th><th style={{ ...styles.th, textAlign: "right" }}>Net pay</th><th style={styles.th}></th></tr></thead>
          <tbody>{calcs.map(({ emp, c }) => (
            <React.Fragment key={emp.id}>
              <tr onClick={() => setOpenEmp(openEmp === emp.id ? null : emp.id)} style={{ cursor: "pointer" }}>
                <td style={{ ...styles.td, fontWeight: 500 }}>{openEmp === emp.id ? "\u25be" : "\u25b8"} {emp.name}</td>
                <td style={styles.td}>{emp.role || "-"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(c.grossM)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", color: theme.amber }}>{fmt(c.payeM)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(c.pensionM)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(c.nhfM + c.customM)}</td>
                <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600, color: theme.emerald }}>{fmt(c.netM)}</td>
                <td style={styles.td}><button style={styles.iconBtn} onClick={(e) => { e.stopPropagation(); removeEmployee(emp); }}>🗑</button></td>
              </tr>
              {openEmp === emp.id && (
                <tr><td colSpan={8} style={{ ...styles.td, background: theme.panel2 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, padding: "6px 4px" }}>
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>Statutory deductions</div>
                      <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, padding: "3px 0" }}>
                        <input type="checkbox" checked={emp.pension !== false} onChange={e => patchEmp(emp.id, { pension: e.target.checked })} /> Pension (employee 8%) - {fmt(c.pensionM)}/mo
                      </label>
                      <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, padding: "3px 0" }}>
                        <input type="checkbox" checked={emp.nhf === true} onChange={e => patchEmp(emp.id, { nhf: e.target.checked })} /> NHF (2.5%) - {fmt(c.nhfM)}/mo
                      </label>
                      <div style={{ fontSize: 12.5, marginTop: 8, color: theme.muted }}>Consolidated relief: {fmt(c.craA)} / year · Taxable income: {fmt(c.taxableA)} / year</div>
                      <div style={{ fontSize: 13, marginTop: 4 }}>PAYE: <strong style={{ fontVariantNumeric: "tabular-nums" }}>{fmt(c.payeM)}</strong> per month ({fmt(c.payeM * 12)} annually)</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: theme.muted, marginBottom: 6 }}>Custom deductions</div>
                      {(emp.deductions || []).length === 0 && <div style={{ fontSize: 12.5, color: theme.muted }}>None - add loan repayments, union dues, cooperative contributions, etc.</div>}
                      {(emp.deductions || []).map(dd => (
                        <div key={dd.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, padding: "3px 0" }}>
                          <span>{dd.name}</span>
                          <span style={{ display: "flex", gap: 8, alignItems: "center" }}><span style={{ fontVariantNumeric: "tabular-nums" }}>{fmt(dd.amount)}/mo</span><button style={styles.iconBtn} onClick={() => removeDeduction(emp, dd)}>✕</button></span>
                        </div>
                      ))}
                      <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                        <input style={{ ...styles.inputSmall, flex: 1 }} placeholder="Deduction name" value={dedForm.name} onChange={e => setDedForm({ ...dedForm, name: e.target.value })} />
                        <input type="number" style={{ ...styles.inputSmall, width: 110 }} placeholder="Monthly amt" value={dedForm.amount} onChange={e => setDedForm({ ...dedForm, amount: e.target.value })} />
                        <button style={styles.btnGhost} onClick={() => addDeduction(emp)}>Add</button>
                      </div>
                      <div style={{ fontSize: 13, marginTop: 10, borderTop: `1px solid ${theme.border}`, paddingTop: 8 }}>Net pay: <strong style={{ color: theme.emerald, fontVariantNumeric: "tabular-nums" }}>{fmt(c.netM)}</strong> per month</div>
                    </div>
                  </div>
                </td></tr>
              )}
            </React.Fragment>
          ))}</tbody>
        </table>
        </div>
        <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 160 }} placeholder="Employee name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
          <input style={styles.input} placeholder="Role (optional)" value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} />
          <input type="number" style={{ ...styles.input, width: 160 }} placeholder="Gross monthly pay" value={form.grossMonthly} onChange={e => setForm({ ...form, grossMonthly: e.target.value })} />
          <button style={styles.btnPrimary} onClick={addEmployee}>Add employee</button>
        </div>
      </div>

      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Payroll runs</div>
        {data.payrollRuns.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 8 }}>No runs yet.</div>}
        {[...data.payrollRuns].sort((a, b) => b.month.localeCompare(a.month)).map(run => (
          <React.Fragment key={run.id}>
            <div onClick={() => setOpenRun(openRun === run.id ? null : run.id)} style={{ ...styles.attentionRow, cursor: "pointer" }}>
              <div><div style={{ fontWeight: 500, fontSize: 13.5 }}>{openRun === run.id ? "\u25be" : "\u25b8"} {run.month} \u00b7 {run.employeeCount} employee(s)</div><div style={{ fontSize: 12, color: theme.muted }}>Posted {fmtDate(run.date)}</div></div>
              <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                <span style={{ fontVariantNumeric: "tabular-nums", fontSize: 13 }}>Net {fmt(run.net ?? run.total)}</span>
                <button style={styles.iconBtn} onClick={(e) => { e.stopPropagation(); deleteRun(run); }}>🗑</button>
              </div>
            </div>
            {openRun === run.id && (
              <div style={{ background: theme.panel2, borderRadius: 8, padding: "10px 14px", marginBottom: 8 }}>
                {run.snapshot ? (
                  <div style={{ overflowX: "auto" }}>
                  <table style={styles.table}>
                    <thead><tr><th style={styles.th}>Employee</th><th style={{ ...styles.th, textAlign: "right" }}>Gross</th><th style={{ ...styles.th, textAlign: "right" }}>PAYE</th><th style={{ ...styles.th, textAlign: "right" }}>Pension</th><th style={{ ...styles.th, textAlign: "right" }}>NHF</th><th style={{ ...styles.th, textAlign: "right" }}>Other</th><th style={{ ...styles.th, textAlign: "right" }}>Net</th></tr></thead>
                    <tbody>{run.snapshot.map((s, i) => (
                      <tr key={i}><td style={styles.td}>{s.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.gross)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.paye)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.pension)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.nhf)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(s.custom)}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(s.net)}</td></tr>
                    ))}</tbody>
                  </table>
                  </div>
                ) : <div style={{ fontSize: 12.5, color: theme.muted }}>This run predates per-employee snapshots - only totals were recorded.</div>}
              </div>
            )}
          </React.Fragment>
        ))}
      </div>
      <div style={{ fontSize: 12, color: theme.muted }}>Statutory rates follow the Nigerian PITA bands and Finance Act reliefs as of 2023; confirm current rates with FIRS/your state IRS before filing. Running payroll posts gross to Payroll Expense, net to the bank, and each deduction to its own payable account (PAYE 2210, Pension 2220, NHF 2230, Other 2240) - remit them from the Payments tab.</div>
    </div>
  );
}
function ProjectsPage({ data, setData, notify }) {
  return (
    <div>
      <PageHeader eyebrow="Operations" title="Projects" sub={`${data.projects.length} projects`} />
      <ProjectsManager data={data} setData={setData} notify={notify} />
      <ReportProjectSummary data={data} />
    </div>
  );
}

/* =================================== Tax page =================================== */
// IAS 12 deferred tax and current tax provisioning. Both actions post real
// journal entries and log to data.taxProvisions for a visible, deletable
// history - deleting a run reverses its ledger entry, matching every other
// "run" feature in the app (payroll, depreciation).
function DeferredTaxPanel({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [range, setRange] = useState("year");
  const dt = computeDeferredTax(data);
  const ct = computeCurrentTax(data, range);
  const rate = data.settings.corporateTaxRate ?? 30;

  const postDeferred = async () => {
    if (Math.abs(dt.movement) < 0.5) return notify("Deferred tax is already at its target balance - nothing to post");
    if (!(await confirm(`Post a deferred tax adjustment of ${fmt(Math.abs(dt.movement))} (${dt.movement > 0 ? "increase" : "decrease"} in net deferred tax liability)? This moves Deferred Tax Liability/Asset to match the current temporary differences on your fixed assets.`))) return;
    const lines = deferredTaxAdjustmentLines(dt);
    const txn = buildTxn(`Deferred tax adjustment`, todayStr(), lines, "manual");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], taxProvisions: [...d.taxProvisions, { id: uid("taxp"), type: "deferred", date: todayStr(), amount: dt.movement, targetNet: dt.targetNet, txnId: txn.id }] }));
    notify("Deferred tax adjustment posted");
  };
  const postCurrent = async () => {
    if (ct.currentTax <= 0.5) return notify("No current tax due for this period at a taxable profit of " + fmt(ct.taxableProfit));
    if (!(await confirm(`Post current tax provision of ${fmt(ct.currentTax)} for ${periodLabel(range)}? Posts Dr Income Tax Expense / Cr Current Tax Payable.`))) return;
    const txn = buildTxn(`Current tax provision - ${periodLabel(range)}`, todayStr(), [{ accountId: "5850", debit: Math.round(ct.currentTax), credit: 0 }, { accountId: "2250", debit: 0, credit: Math.round(ct.currentTax) }], "manual");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], taxProvisions: [...d.taxProvisions, { id: uid("taxp"), type: "current", date: todayStr(), amount: ct.currentTax, period: periodLabel(range), txnId: txn.id }] }));
    notify("Current tax provision posted");
  };
  const deleteRun = async (r) => {
    if (!(await confirm(`Delete this ${r.type} tax provision of ${fmt(r.amount)}? Its journal entry is removed.`))) return;
    setData(d => ({ ...d, taxProvisions: d.taxProvisions.filter(x => x.id !== r.id), transactions: d.transactions.filter(t => t.id !== r.txnId) }));
    notify("Provision deleted and its journal entry removed");
  };

  return (
    <>
      <div className="ces-card" style={styles.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <div style={styles.cardTitle}>Corporate income tax rate</div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={{ ...styles.input, width: 90 }} value={rate} onChange={e => setData(d => ({ ...d, settings: { ...d.settings, corporateTaxRate: Number(e.target.value) || 0 } }))} />
            <span style={{ fontSize: 13, color: theme.muted }}>%</span>
          </div>
        </div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Used for both the current tax provision and deferred tax calculations below. Nigeria's large-company CIT rate is 30% - check the tiered rate that applies to your business.</div>
      </div>

      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Deferred tax</div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>The gap between each fixed asset's book value and its tax written-down value (using the tax rate set on each asset) is a temporary difference. Positive differences create a deferred tax liability; negative differences create a deferred tax asset.</div>
        {dt.perAsset.length === 0 ? <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No fixed assets on the register.</div> : (
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}>Asset</th><th style={{ ...styles.th, textAlign: "right" }}>Book NBV</th><th style={{ ...styles.th, textAlign: "right" }}>Tax WDV</th><th style={{ ...styles.th, textAlign: "right" }}>Temp. difference</th></tr></thead>
            <tbody>{dt.perAsset.map(a => (
              <tr key={a.id}><td style={styles.td}>{a.name}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(a.bookNBV)}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(a.taxWDV)}</td><td style={{ ...styles.tdMono, textAlign: "right", color: a.tempDiff >= 0 ? theme.amber : theme.emerald }}>{fmt(a.tempDiff)}</td></tr>
            ))}</tbody>
          </table>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginTop: 14 }}>
          <Kpi label="Total temp. difference" value={fmt(dt.totalTempDiff)} />
          <Kpi label={dt.targetNet >= 0 ? "Target deferred tax liability" : "Target deferred tax asset"} value={fmt(Math.abs(dt.targetNet))} tone="amber" />
          <Kpi label="Currently recognized (net)" value={fmt(dt.recognizedNet)} />
          <Kpi label="Adjustment needed" value={fmt(dt.movement)} tone={Math.abs(dt.movement) < 0.5 ? "emerald" : "amber"} />
        </div>
        <button style={{ ...styles.btnPrimary, marginTop: 14 }} onClick={postDeferred}>Post deferred tax adjustment</button>
      </div>

      <div className="ces-card" style={styles.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <div style={styles.cardTitle}>Current tax provision</div>
          <select style={styles.input} value={range} onChange={e => setRange(e.target.value)}>{RANGE_OPTIONS.filter(([v]) => v !== "custom" && v !== "all").map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select>
        </div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Taxable profit reverses book depreciation (not tax-deductible) and substitutes each asset's tax capital allowance instead. For an in-progress period, tax depreciation is estimated through to the period end.</div>
        <div style={{ marginTop: 10 }}>
          <RowLine label="Accounting profit" value={ct.accountingProfit} />
          <RowLine label="Add back: book depreciation" value={ct.bookDepreciation} indent />
          <RowLine label="Less: tax capital allowances" value={-ct.taxDepreciation} indent />
          <RowLine label="Taxable profit" value={ct.taxableProfit} bold divider />
          <RowLine label={`Current tax @ ${(ct.rate * 100).toFixed(0)}%`} value={ct.currentTax} bold tone="amber" divider />
        </div>
        <button style={{ ...styles.btnPrimary, marginTop: 14 }} onClick={postCurrent}>Post current tax provision</button>
      </div>

      {data.taxProvisions.length > 0 && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Tax provision history</div>
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Type</th><th style={styles.th}>Period</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th><th style={styles.th}></th></tr></thead>
            <tbody>{[...data.taxProvisions].sort((a, b) => b.date.localeCompare(a.date)).map(r => (
              <tr key={r.id}><td style={styles.tdMono}>{fmtDate(r.date)}</td><td style={styles.td}>{r.type === "deferred" ? "Deferred" : "Current"}</td><td style={styles.td}>{r.period || "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(r.amount)}</td><td style={styles.td}><button style={styles.iconBtn} onClick={() => deleteRun(r)}>🗑</button></td></tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </>
  );
}
// A present obligation from a past event, where an outflow is probable and
// the amount can be reliably estimated, gets recognized as a provision -
// legal claims, warranty obligations, restructuring costs, and the like.
// Recognizing or increasing one is a non-cash expense; utilizing one is a
// real cash payment against it; releasing one reverses it when the
// obligation no longer applies.
// Lets you close the books up to a date - a standard control preventing a
// reported/filed period from being silently altered later. Shows exactly
// how many existing transactions would fall inside the new lock before you
// commit to it, since this affects delete/edit permissions app-wide.
// A whole year's budget in one grid - one row per revenue/expense account,
// one editable column per month, plus a "Default" column that fills any
// month without its own override. This is the actual place to plan a
// period-by-period budget; the Budget vs Actual report edits the same
// underlying data, just one month at a time while you're looking at it.
// Upload a budget from CSV/Excel - same shape as the bank statement import:
// parse, auto-detect columns, preview and flag anything unmatched, then
// commit only once confirmed. Expects one row per account with a "Default"
// column and/or one column per month (Jan, Feb, ... or 1-12) for the
// selected year - the same wide layout as the Budgets grid itself, so a
// spreadsheet exported from there (or filled in from scratch) round-trips.
const MONTH_NAMES = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
function BudgetUploadPanel({ data, setData, notify, year, onDone }) {
  const { styles, theme } = useUI();
  const [fileName, setFileName] = useState("");
  const [preview, setPreview] = useState(null); // { rows: [{ raw, account, matched, values: {default, "01":..,"12":..} }] }

  const downloadTemplate = () => {
    const accts = data.accounts.filter(a => (a.type === "expense" || a.type === "revenue") && a.status !== "inactive");
    const header = ["Account", "Default", ...MONTH_NAMES.map(m => m[0].toUpperCase() + m.slice(1))];
    const rows = accts.map(a => [a.name, budgetFor(data, "default", a.id) || "", ...Array.from({ length: 12 }, (_, i) => budgetFor(data, `${year}-${String(i + 1).padStart(2, "0")}`, a.id) || "")]);
    const csv = [header, ...rows].map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = `budget-template-${year}.csv`; a.click();
    URL.revokeObjectURL(url);
  };

  const handleFile = async (file) => {
    setFileName(file.name);
    const ext = file.name.split(".").pop().toLowerCase();
    let parsed = [];
    if (ext === "csv") { const text = await file.text(); parsed = Papa.parse(text, { header: true, skipEmptyLines: true }).data; }
    else if (ext === "xlsx" || ext === "xls") { const buf = await file.arrayBuffer(); const wb = XLSX.read(buf, { type: "array" }); parsed = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }); }
    else { notify("Unsupported file type. Use CSV or Excel."); return; }
    if (parsed.length === 0) return notify("No rows found in that file");

    const keys = Object.keys(parsed[0]);
    const accountKey = keys.find(k => /^account$|^name$/i.test(k)) || keys[0];
    const defaultKey = keys.find(k => /default/i.test(k));
    const monthKeyFor = (mIdx) => keys.find(k => {
      const norm = k.trim().toLowerCase();
      return norm === MONTH_NAMES[mIdx] || norm === String(mIdx + 1) || norm.startsWith(MONTH_NAMES[mIdx]);
    });
    const monthKeys = MONTH_NAMES.map((_, i) => monthKeyFor(i));
    const numOf = (v) => { const n = Number(String(v).replace(/[^0-9.-]/g, "")); return isNaN(n) ? 0 : n; };

    const rows = parsed.map(r => {
      const raw = String(r[accountKey] || "").trim();
      if (!raw) return null;
      const matched = data.accounts.find(a => a.name.toLowerCase() === raw.toLowerCase()) || data.accounts.find(a => a.code === raw);
      const hasDefault = defaultKey && r[defaultKey] !== undefined && r[defaultKey] !== "";
      const values = { default: hasDefault ? numOf(r[defaultKey]) : undefined };
      monthKeys.forEach((k, i) => { if (k && r[k] !== undefined && r[k] !== "") values[String(i + 1).padStart(2, "0")] = numOf(r[k]); });
      return { raw, matched, values };
    }).filter(Boolean);

    if (rows.length === 0) return notify("No usable rows found - make sure there's an Account column");
    setPreview(rows);
  };

  const matchedCount = preview ? preview.filter(r => r.matched).length : 0;
  const unmatchedCount = preview ? preview.length - matchedCount : 0;

  const confirmImport = () => {
    const good = preview.filter(r => r.matched);
    if (good.length === 0) return notify("None of the account names matched your chart of accounts");
    setData(d => {
      const next = { ...d.budgets };
      good.forEach(r => {
        Object.entries(r.values).forEach(([key, val]) => {
          if (val === undefined) return;
          const periodKey = key === "default" ? "default" : `${year}-${key}`;
          next[periodKey] = { ...(next[periodKey] || {}), [r.matched.id]: val };
        });
      });
      const newIds = good.map(r => r.matched.id).filter(id => !d.budgetAccounts.includes(id));
      return { ...d, budgets: next, budgetAccounts: [...d.budgetAccounts, ...newIds] };
    });
    notify(`Imported budgets for ${good.length} account(s)${unmatchedCount ? ` · ${unmatchedCount} row(s) skipped (no matching account)` : ""}`);
    setPreview(null); setFileName(""); onDone && onDone();
  };

  return (
    <div className="ces-card" style={styles.card}>
      <div style={styles.cardTitle}>Upload a budget</div>
      <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Import a CSV or Excel file with one row per account: an "Account" column matched against your chart of accounts by name, an optional "Default" column, and a column per month (Jan, Feb... or 1-12) for {year}. Not sure of the layout? Download a template pre-filled from your current budget.</div>
      <div style={{ display: "flex", gap: 10, marginTop: 12, alignItems: "center", flexWrap: "wrap" }}>
        <button style={styles.btnGhost} onClick={downloadTemplate}>Download template ({year})</button>
        <input id="budget-upload-input" type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f); e.target.value = ""; }} />
        <label htmlFor="budget-upload-input" style={{ ...styles.btnPrimary, display: "inline-block", cursor: "pointer" }}>Choose file to upload</label>
        {fileName && <span style={{ fontSize: 12, color: theme.muted }}>{fileName}</span>}
      </div>

      {preview && (
        <div style={{ marginTop: 16 }}>
          <div style={{ fontSize: 12.5, color: theme.muted, marginBottom: 8 }}>{matchedCount} account(s) matched{unmatchedCount ? `, ${unmatchedCount} not found in your chart of accounts (shown in red - these rows will be skipped)` : ""}.</div>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Account (from file)</th><th style={styles.th}>Matched to</th><th style={{ ...styles.th, textAlign: "right" }}>Default</th><th style={{ ...styles.th, textAlign: "right" }}>Months set</th></tr></thead>
            <tbody>{preview.map((r, i) => (
              <tr key={i}><td style={{ ...styles.td, color: r.matched ? theme.text : theme.rose }}>{r.raw}</td>
                <td style={styles.td}>{r.matched ? r.matched.name : "No match"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{r.values.default !== undefined ? r.values.default : "-"}</td>
                <td style={{ ...styles.tdMono, textAlign: "right" }}>{Object.keys(r.values).filter(k => k !== "default").length}</td></tr>
            ))}</tbody>
          </table>
          <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
            <button style={styles.btnPrimary} onClick={confirmImport}>Import {matchedCount} account(s)</button>
            <button style={styles.btnGhost} onClick={() => { setPreview(null); setFileName(""); }}>Cancel</button>
          </div>
        </div>
      )}
    </div>
  );
}
function BudgetsPage({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [year, setYear] = useState(new Date().getFullYear());
  const [showUpload, setShowUpload] = useState(false);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(null); // { accountIds, budgets } - untouched until Save
  const [addAccountId, setAddAccountId] = useState("");
  const months = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, "0"));
  const monthLabel = (mm) => new Date(2000, Number(mm) - 1, 1).toLocaleDateString(undefined, { month: "short" });

  // While editing, everything reads from and writes to the local draft;
  // Save is the only thing that commits it back to real data. Locked
  // (view) mode always reflects the real, saved data.
  const source = editing ? draft : data;
  const accts = source.budgetAccounts.map(id => data.accounts.find(a => a.id === id)).filter(Boolean);
  const eligibleToAdd = data.accounts.filter(a => (a.type === "expense" || a.type === "revenue") && a.status !== "inactive" && !source.budgetAccounts.includes(a.id));
  const lookup = (periodKey, accountId) => source.budgets[periodKey]?.[accountId] ?? source.budgets.default?.[accountId] ?? 0;

  const startEdit = () => { setDraft({ budgets: JSON.parse(JSON.stringify(data.budgets)), budgetAccounts: [...data.budgetAccounts] }); setEditing(true); setShowUpload(false); };
  const cancelEdit = async () => { if (!(await confirm("Discard your changes since you last saved?"))) return; setDraft(null); setEditing(false); };
  const saveEdit = () => { setData(d => ({ ...d, budgets: draft.budgets, budgetAccounts: draft.budgetAccounts })); setDraft(null); setEditing(false); notify("Budget saved"); };

  // Setting the Default is a broadcast: it overwrites all 12 months of the
  // year currently being viewed, on top of updating the stored default
  // itself - not just a fallback used when a month happens to be blank.
  // Editing one month afterward only changes that month; going back and
  // changing Default again re-broadcasts and overwrites all 12 once more.
  const setDefault = (accountId, val) => {
    const amount = Number(val) || 0;
    setDraft(d => {
      const next = { ...d.budgets, default: { ...(d.budgets.default || {}), [accountId]: amount } };
      months.forEach(mm => { const key = `${year}-${mm}`; next[key] = { ...(next[key] || {}), [accountId]: amount }; });
      return { ...d, budgets: next };
    });
  };
  const setMonth = (accountId, mm, val) => {
    const key = `${year}-${mm}`;
    setDraft(d => ({ ...d, budgets: { ...d.budgets, [key]: { ...(d.budgets[key] || {}), [accountId]: Number(val) || 0 } } }));
  };
  const addAccount = () => { if (!addAccountId) return; setDraft(d => ({ ...d, budgetAccounts: [...d.budgetAccounts, addAccountId] })); setAddAccountId(""); };
  const removeAccount = (accountId) => setDraft(d => ({ ...d, budgetAccounts: d.budgetAccounts.filter(id => id !== accountId) }));

  const rowTotal = (accountId) => months.reduce((s, mm) => s + lookup(`${year}-${mm}`, accountId), 0);
  const colTotal = (mm) => accts.reduce((s, a) => s + (a.type === "expense" ? lookup(`${year}-${mm}`, a.id) : 0), 0);
  const bg = theme.mode === "dark" ? "rgba(255,255,255,0.04)" : "rgba(0,0,0,0.02)";

  return (
    <div>
      <PageHeader eyebrow="Accounting" title="Budgets" sub={editing ? "Editing - changes apply once you save" : "Locked - press Edit to make changes"} action={
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {!editing && <button style={styles.btnGhost} onClick={() => setShowUpload(s => !s)}>{showUpload ? "Hide upload" : "⇧ Upload budget"}</button>}
          <button style={styles.btnGhost} onClick={() => setYear(y => y - 1)}>←</button>
          <span style={{ fontFamily: "Inter, sans-serif", fontWeight: 600, fontSize: 15 }}>{year}</span>
          <button style={styles.btnGhost} onClick={() => setYear(y => y + 1)}>→</button>
          {editing ? <>
            <button style={styles.btnGhost} onClick={cancelEdit}>Cancel</button>
            <button style={styles.btnPrimary} onClick={saveEdit}>Save</button>
          </> : <button style={styles.btnPrimary} onClick={startEdit}>✎ Edit</button>}
        </div>
      } />
      {showUpload && !editing && <BudgetUploadPanel data={data} setData={setData} notify={notify} year={year} onDone={() => setShowUpload(false)} />}
      <div className="ces-card" style={styles.card}>
        <div style={{ fontSize: 12, color: theme.muted, marginBottom: 10, display: "flex", alignItems: "center" }}>How Default works<InfoTip width={300} text={`Changing "Default" sets all 12 months of ${year} to that amount at once - edit any single month afterward to make it differ, like a bigger marketing push in December. Changing Default again resets all 12 months, including any you'd customized individually.`} /></div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ ...styles.table, minWidth: 1100 }}>
            <thead><tr>
              <th style={styles.th}>Account</th>
              <th style={{ ...styles.th, textAlign: "right", background: bg }}>Default</th>
              {months.map(mm => <th key={mm} style={{ ...styles.th, textAlign: "right" }}>{monthLabel(mm)}</th>)}
              <th style={{ ...styles.th, textAlign: "right" }}>Year total</th>
              {editing && <th style={styles.th}></th>}
            </tr></thead>
            <tbody>{accts.map(a => (
              <tr key={a.id}>
                <td style={styles.td}>{a.name}</td>
                <td style={{ ...styles.td, textAlign: "right", background: bg }}>
                  {editing
                    ? <input type="number" style={{ ...styles.inputSmall, width: 90, textAlign: "right" }} value={source.budgets.default?.[a.id] || ""} placeholder="0" onChange={e => setDefault(a.id, e.target.value)} />
                    : <span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13 }}>{fmt(source.budgets.default?.[a.id] || 0)}</span>}
                </td>
                {months.map(mm => {
                  const key = `${year}-${mm}`;
                  const val = lookup(key, a.id);
                  return (
                    <td key={mm} style={{ ...styles.td, textAlign: "right" }}>
                      {editing
                        ? <input type="number" style={{ ...styles.inputSmall, width: 80, textAlign: "right" }} value={val || ""} placeholder="0" onChange={e => setMonth(a.id, mm, e.target.value)} />
                        : <span style={{ fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13 }}>{fmt(val)}</span>}
                    </td>
                  );
                })}
                <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(rowTotal(a.id))}</td>
                {editing && <td style={styles.td}><button style={styles.iconBtn} onClick={() => removeAccount(a.id)}>🗑</button></td>}
              </tr>
            ))}</tbody>
            <tfoot><tr>
              <td style={{ ...styles.td, fontWeight: 700 }}>Total expense budget</td><td></td>
              {months.map(mm => <td key={mm} style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(colTotal(mm))}</td>)}
              <td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 700 }}>{fmt(months.reduce((s, mm) => s + colTotal(mm), 0))}</td>
              {editing && <td></td>}
            </tr></tfoot>
          </table>
        </div>
        {accts.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No accounts in the budget yet{editing ? " - add one below." : ". Press Edit to add some."}</div>}
        {editing && (
          <div style={{ display: "flex", gap: 10, marginTop: 14, alignItems: "center", flexWrap: "wrap" }}>
            <select style={styles.input} value={addAccountId} onChange={e => setAddAccountId(e.target.value)}>
              <option value="">Choose an account to add...</option>{eligibleToAdd.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
            <button style={styles.btnGhost} onClick={addAccount}>+ Add to budget</button>
          </div>
        )}
      </div>
    </div>
  );
}
function TransactionLocking({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [draftDate, setDraftDate] = useState(data.settings.lockDate || "");
  const currentLock = data.settings.lockDate;
  const affectedCount = draftDate ? data.transactions.filter(t => t.date <= draftDate).length : 0;
  const currentlyLockedCount = currentLock ? data.transactions.filter(t => t.date <= currentLock).length : 0;

  const applyLock = async () => {
    if (!draftDate) return notify("Choose a date first");
    if (!(await confirm(`Lock the books through ${fmtDate(draftDate)}? ${affectedCount} existing transaction(s) on or before that date will no longer be editable or deletable anywhere in the app - correcting one after this means posting a new, dated reversing entry instead.`))) return;
    setData(d => ({ ...d, settings: { ...d.settings, lockDate: draftDate } }));
    notify(`Books locked through ${fmtDate(draftDate)}`);
  };
  const removeLock = async () => {
    if (!(await confirm("Remove the lock date entirely? Every transaction becomes editable and deletable again."))) return;
    setData(d => ({ ...d, settings: { ...d.settings, lockDate: null } }));
    setDraftDate("");
    notify("Lock removed");
  };

  return (
    <div>
      <PageHeader eyebrow="Accounting" title="Transaction Locking" sub={currentLock ? `Books locked through ${fmtDate(currentLock)}` : "Books are fully open - nothing is locked"} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>{currentLock ? "Update the lock date" : "Lock the books"}<InfoTip width={300} text="Once a period is reported on or filed with a regulator, locking it stops anyone from quietly changing the numbers behind it - anywhere in the app: the Journal, Invoices, Bills, Expenses. Fixing a mistake in a locked period is still possible, just done the right way: post a new, dated correcting or reversing entry rather than editing history." /></div>
        {currentLock && <div style={{ fontSize: 13, color: theme.text, marginTop: 12 }}>Currently locked through <strong>{fmtDate(currentLock)}</strong> - {currentlyLockedCount} transaction(s) affected.</div>}
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
          <input type="date" style={styles.input} value={draftDate} onChange={e => setDraftDate(e.target.value)} />
          <button style={styles.btnPrimary} onClick={applyLock}>{currentLock ? "Update lock date" : "Lock through this date"}</button>
          {currentLock && <button style={{ ...styles.btnGhost, color: theme.rose, borderColor: theme.rose }} onClick={removeLock}>Remove lock</button>}
        </div>
        {draftDate && <div style={{ fontSize: 12, color: theme.muted, marginTop: 10 }}>{affectedCount} transaction(s) are on or before {fmtDate(draftDate)} and would become locked.</div>}
      </div>
    </div>
  );
}
// Reclassifies many transaction lines from one account to another at once -
// the common year-end cleanup of "these 30 things got coded to the wrong
// account." Only relabels the account on matching lines; amounts and
// balance are untouched, so nothing can go out of balance from this.
// Locked transactions are excluded automatically - reclassifying inside a
// closed period would defeat the point of locking it.
function BulkUpdate({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const [sourceId, setSourceId] = useState(data.accounts[0]?.id || "");
  const [targetId, setTargetId] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(() => new Set());

  const candidates = data.transactions
    .filter(t => t.lines.some(l => l.accountId === sourceId))
    .filter(t => !from || t.date >= from)
    .filter(t => !to || t.date <= to)
    .filter(t => !search || (t.memo || "").toLowerCase().includes(search.toLowerCase()))
    .map(t => ({ t, locked: isLocked(t.date, data) }))
    .sort((a, b) => b.t.date.localeCompare(a.t.date));
  const selectableIds = candidates.filter(c => !c.locked).map(c => c.t.id);

  const toggle = (id) => setSelected(s => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toggleAll = () => setSelected(s => s.size === selectableIds.length ? new Set() : new Set(selectableIds));

  const apply = async () => {
    if (!targetId || targetId === sourceId) return notify("Choose a different target account");
    if (selected.size === 0) return notify("Select at least one transaction");
    const sourceAcc = data.accounts.find(a => a.id === sourceId), targetAcc = data.accounts.find(a => a.id === targetId);
    if (!(await confirm(`Reclassify ${selected.size} transaction line(s) from ${sourceAcc?.name} to ${targetAcc?.name}? Amounts stay exactly the same - only the account changes. This can't be undone automatically (though you can always run another bulk update back).`))) return;
    setData(d => ({
      ...d,
      transactions: d.transactions.map(t => selected.has(t.id)
        ? { ...t, lines: t.lines.map(l => l.accountId === sourceId ? { ...l, accountId: targetId } : l) }
        : t),
      __auditLabel: `Bulk update - reclassified ${selected.size} transaction(s) from ${sourceAcc?.name} to ${targetAcc?.name}`,
    }));
    setSelected(new Set());
    notify(`${selected.size} transaction(s) reclassified`);
  };

  return (
    <div>
      <PageHeader eyebrow="Accounting" title="Bulk Update" sub="Reclassify many transactions between accounts at once" />
      <div className="ces-card" style={styles.card}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <label style={{ fontSize: 12 }}>From account
            <select style={{ ...styles.input, display: "block", marginTop: 4, minWidth: 200 }} value={sourceId} onChange={e => { setSourceId(e.target.value); setSelected(new Set()); }}>
              {data.accounts.map(a => <option key={a.id} value={a.id}>{a.code} {a.name}</option>)}
            </select>
          </label>
          <label style={{ fontSize: 12 }}>To account
            <select style={{ ...styles.input, display: "block", marginTop: 4, minWidth: 200 }} value={targetId} onChange={e => setTargetId(e.target.value)}>
              <option value="">Choose target...</option>{data.accounts.filter(a => a.id !== sourceId).map(a => <option key={a.id} value={a.id}>{a.code} {a.name}</option>)}
            </select>
          </label>
          <label style={{ fontSize: 12 }}>From date
            <input type="date" style={{ ...styles.input, display: "block", marginTop: 4 }} value={from} onChange={e => setFrom(e.target.value)} />
          </label>
          <label style={{ fontSize: 12 }}>To date
            <input type="date" style={{ ...styles.input, display: "block", marginTop: 4 }} value={to} onChange={e => setTo(e.target.value)} />
          </label>
          <label style={{ fontSize: 12, flex: 1, minWidth: 160 }}>Search memo
            <input style={{ ...styles.input, display: "block", marginTop: 4, width: "100%" }} value={search} onChange={e => setSearch(e.target.value)} placeholder="Optional" />
          </label>
        </div>
      </div>

      <div className="ces-card" style={styles.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <div style={styles.cardTitle}>{candidates.length} matching transaction(s)</div>
          {selected.size > 0 && <button style={styles.btnPrimary} onClick={apply}>Reclassify {selected.size} selected</button>}
        </div>
        {candidates.length === 0 ? <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No transactions touch this account for the current filters.</div> : (
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}><input type="checkbox" checked={selected.size > 0 && selected.size === selectableIds.length} onChange={toggleAll} /></th><th style={styles.th}>Date</th><th style={styles.th}>Memo</th><th style={styles.th}>Source</th><th style={{ ...styles.th, textAlign: "right" }}>Amount</th></tr></thead>
            <tbody>{candidates.map(({ t, locked }) => {
              const line = t.lines.find(l => l.accountId === sourceId);
              const amt = line ? (line.debit || line.credit) : 0;
              return (
                <tr key={t.id} style={{ opacity: locked ? 0.5 : 1 }}>
                  <td style={styles.td}><input type="checkbox" disabled={locked} checked={selected.has(t.id)} onChange={() => toggle(t.id)} /></td>
                  <td style={styles.tdMono}>{fmtDate(t.date)}</td>
                  <td style={styles.td}>{locked && <span title="Locked - excluded" style={{ marginRight: 5 }}>🔒</span>}{t.memo}</td>
                  <td style={{ ...styles.td, fontSize: 12, color: theme.muted }}>{t.source}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(amt)}</td>
                </tr>
              );
            })}</tbody>
          </table>
        )}
        <div style={{ fontSize: 11, color: theme.muted, marginTop: 10, display: "flex", alignItems: "center" }}>How reclassifying works<InfoTip text="Locked transactions (🔒) are excluded automatically and can't be selected. Only the account on each line changes - debit and credit amounts stay exactly as they were." /></div>
      </div>
    </div>
  );
}
// Logs hours against an employee and (optionally) a project, tags each
// entry billable or not, and lets you turn a project's unbilled billable
// hours straight into a real invoice - one line per entry, using the same
// posting path (invoiceJournalLines) every other invoice in the app uses.
function TimeSheet({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const blank = () => ({ employeeId: "", employeeName: "", projectId: "", date: todayStr(), hours: "", description: "", billable: false, rate: "" });
  const [form, setForm] = useState(blank());
  const [filterEmployee, setFilterEmployee] = useState("");
  const [filterProject, setFilterProject] = useState("");
  const [billingProjectId, setBillingProjectId] = useState("");

  const employeeLabel = (ts) => ts.employeeId ? (data.employees.find(e => e.id === ts.employeeId)?.name || "(removed employee)") : ts.employeeName || "Unassigned";
  const projectLabel = (ts) => ts.projectId ? (data.projects.find(p => p.id === ts.projectId)?.name || "(removed project)") : "No project";

  const addEntry = () => {
    if (!Number(form.hours)) return notify("Enter the hours worked");
    if (!form.employeeId && !form.employeeName) return notify("Choose an employee or type a name");
    setData(d => ({ ...d, timesheets: [...d.timesheets, { id: uid("ts"), employeeId: form.employeeId, employeeName: form.employeeId ? "" : form.employeeName, projectId: form.projectId, date: form.date, hours: Number(form.hours), description: form.description, billable: form.billable, rate: form.billable ? Number(form.rate) || 0 : 0, invoiced: false, invoiceId: null }] }));
    setForm(blank());
    notify("Time entry logged");
  };
  const deleteEntry = async (ts) => {
    if (ts.invoiced) return notify("This entry has already been invoiced - delete or credit the invoice first if it needs to change");
    if (!(await confirm("Delete this time entry?"))) return;
    setData(d => ({ ...d, timesheets: d.timesheets.filter(x => x.id !== ts.id) }));
  };

  const filtered = data.timesheets.filter(ts => (!filterEmployee || ts.employeeId === filterEmployee) && (!filterProject || ts.projectId === filterProject));
  const sorted = [...filtered].sort((a, b) => b.date.localeCompare(a.date));
  const totalHours = filtered.reduce((s, ts) => s + ts.hours, 0);
  const billableHours = filtered.filter(ts => ts.billable).reduce((s, ts) => s + ts.hours, 0);
  const unbilledValue = filtered.filter(ts => ts.billable && !ts.invoiced).reduce((s, ts) => s + ts.hours * ts.rate, 0);

  const unbilledForProject = (projectId) => data.timesheets.filter(ts => ts.projectId === projectId && ts.billable && !ts.invoiced);
  const billProject = async (project) => {
    const entries = unbilledForProject(project.id);
    if (entries.length === 0) return notify("No unbilled billable time for this project");
    if (!project.client) return notify("This project has no client set - add one in Projects first");
    const total = entries.reduce((s, ts) => s + ts.hours * ts.rate, 0);
    if (!(await confirm(`Create an invoice for ${project.client} covering ${entries.length} time entr${entries.length === 1 ? "y" : "ies"} (${fmt(total)})?`))) return;
    const invForm = { customer: project.client, date: todayStr(), dueDate: "", projectId: project.id, items: entries.map(ts => ({ desc: `${ts.description || "Time"} - ${employeeLabel(ts)} (${ts.hours}h @ ${fmt(ts.rate)})`, qty: 1, price: ts.hours * ts.rate, inventoryId: "" })), taxes: [] };
    const lines = invoiceJournalLines(invForm, data);
    const id = `INV-${data.nextInvoiceNum}`;
    const txn = buildTxn(`Invoice ${id} - ${project.client}`, invForm.date, lines, "invoice", id);
    if (!txn) return notify("Entry not balanced");
    setData(d => ({
      ...d,
      transactions: [...d.transactions, txn],
      invoices: [...d.invoices, { ...invForm, items: snapshotTaxItems(invForm.items, d.taxGroups), id, status: "sent", amountPaid: 0, locked: false }],
      nextInvoiceNum: d.nextInvoiceNum + 1,
      timesheets: d.timesheets.map(ts => entries.some(e => e.id === ts.id) ? { ...ts, invoiced: true, invoiceId: id } : ts),
    }));
    notify(`${id} created from ${entries.length} time entr${entries.length === 1 ? "y" : "ies"}`);
  };

  const billableProjects = data.projects.filter(p => unbilledForProject(p.id).length > 0);

  return (
    <div>
      <PageHeader eyebrow="Operations" title="Time Sheet" sub={`${data.timesheets.length} entries logged`} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Log time</div>
        <div style={{ display: "flex", gap: 10, marginTop: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          {data.employees.length > 0 ? (
            <label style={{ fontSize: 12 }}>Employee
              <select style={{ ...styles.input, display: "block", marginTop: 4, minWidth: 160 }} value={form.employeeId} onChange={e => setForm({ ...form, employeeId: e.target.value, employeeName: "" })}>
                <option value="">Choose or type below...</option>{data.employees.map(e => <option key={e.id} value={e.id}>{e.name}</option>)}
              </select>
            </label>
          ) : null}
          {!form.employeeId && (
            <label style={{ fontSize: 12 }}>{data.employees.length > 0 ? "or type a name" : "Employee name"}
              <input style={{ ...styles.input, display: "block", marginTop: 4, minWidth: 160 }} value={form.employeeName} onChange={e => setForm({ ...form, employeeName: e.target.value })} placeholder="Name" />
            </label>
          )}
          <label style={{ fontSize: 12 }}>Project
            <select style={{ ...styles.input, display: "block", marginTop: 4, minWidth: 160 }} value={form.projectId} onChange={e => setForm({ ...form, projectId: e.target.value })}>
              <option value="">No project</option>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </label>
          <label style={{ fontSize: 12 }}>Date
            <input type="date" style={{ ...styles.input, display: "block", marginTop: 4 }} value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
          </label>
          <label style={{ fontSize: 12 }}>Hours
            <input type="number" style={{ ...styles.input, display: "block", marginTop: 4, width: 90 }} value={form.hours} onChange={e => setForm({ ...form, hours: e.target.value })} />
          </label>
          <label style={{ fontSize: 12, flex: 1, minWidth: 160 }}>Description
            <input style={{ ...styles.input, display: "block", marginTop: 4, width: "100%" }} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} placeholder="What was worked on" />
          </label>
        </div>
        <div style={{ display: "flex", gap: 14, marginTop: 12, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, cursor: "pointer" }}>
            <input type="checkbox" checked={form.billable} onChange={e => setForm({ ...form, billable: e.target.checked })} /> Billable
          </label>
          {form.billable && (
            <label style={{ fontSize: 12 }}>Rate per hour
              <input type="number" style={{ ...styles.input, display: "block", marginTop: 4, width: 120 }} value={form.rate} onChange={e => setForm({ ...form, rate: e.target.value })} />
            </label>
          )}
          <button style={styles.btnPrimary} onClick={addEntry}>Log time</button>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 16 }}>
        <Kpi label="Total hours" value={totalHours.toFixed(1)} />
        <Kpi label="Billable hours" value={billableHours.toFixed(1)} tone="emerald" />
        <Kpi label="Unbilled value" value={fmt(unbilledValue)} tone="amber" />
      </div>

      {billableProjects.length > 0 && (
        <div className="ces-card" style={styles.card}>
          <div style={styles.cardTitle}>Ready to bill</div>
          <table style={{ ...styles.table, marginTop: 10 }}>
            <thead><tr><th style={styles.th}>Project</th><th style={styles.th}>Client</th><th style={{ ...styles.th, textAlign: "right" }}>Unbilled hours</th><th style={{ ...styles.th, textAlign: "right" }}>Value</th><th style={styles.th}></th></tr></thead>
            <tbody>{billableProjects.map(p => {
              const entries = unbilledForProject(p.id);
              const hrs = entries.reduce((s, ts) => s + ts.hours, 0), val = entries.reduce((s, ts) => s + ts.hours * ts.rate, 0);
              return (
                <tr key={p.id}><td style={styles.td}>{p.name}</td><td style={styles.td}>{p.client || "-"}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{hrs.toFixed(1)}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(val)}</td>
                  <td style={styles.td}><button style={styles.btnGhost} onClick={() => billProject(p)}>Create invoice</button></td></tr>
              );
            })}</tbody>
          </table>
        </div>
      )}

      <div className="ces-card" style={styles.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <div style={styles.cardTitle}>Entries</div>
          <div style={{ display: "flex", gap: 8 }}>
            <select style={styles.inputSmall} value={filterEmployee} onChange={e => setFilterEmployee(e.target.value)}><option value="">All employees</option>{data.employees.map(e => <option key={e.id} value={e.id}>{e.name}</option>)}</select>
            <select style={styles.inputSmall} value={filterProject} onChange={e => setFilterProject(e.target.value)}><option value="">All projects</option>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
        <table style={{ ...styles.table, marginTop: 10 }}>
          <thead><tr><th style={styles.th}>Date</th><th style={styles.th}>Employee</th><th style={styles.th}>Project</th><th style={styles.th}>Description</th><th style={{ ...styles.th, textAlign: "right" }}>Hours</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
          <tbody>{sorted.map(ts => (
            <tr key={ts.id}><td style={styles.tdMono}>{fmtDate(ts.date)}</td><td style={styles.td}>{employeeLabel(ts)}</td><td style={styles.td}>{projectLabel(ts)}</td>
              <td style={styles.td}>{ts.description || "-"}</td><td style={{ ...styles.tdMono, textAlign: "right" }}>{ts.hours}</td>
              <td style={styles.td}>{!ts.billable ? <span style={{ fontSize: 11, color: theme.muted }}>non-billable</span> : ts.invoiced ? <span style={{ ...styles.pill, ...styles.pillGreen }}>invoiced ({ts.invoiceId})</span> : <span style={{ ...styles.pill, ...styles.pillAmber }}>unbilled</span>}</td>
              <td style={styles.td}><button style={styles.iconBtn} onClick={() => deleteEntry(ts)}>🗑</button></td></tr>
          ))}</tbody>
        </table>
        </div>
        {sorted.length === 0 && <div style={{ fontSize: 13, color: theme.muted, marginTop: 10 }}>No time entries logged yet.</div>}
      </div>
    </div>
  );
}
function ProvisionsPage({ data, setData, notify }) {
  const { styles, fmt, theme, confirm } = useUI();
  const blank = () => ({ name: "", category: "Legal", amount: "", description: "" });
  const [form, setForm] = useState(blank());
  const [actioningId, setActioningId] = useState(null);
  const [actionAmount, setActionAmount] = useState("");
  const [actionBankId, setActionBankId] = useState(data.banks[0]?.id);

  const activeProvisions = data.provisions.filter(p => p.status === "active");
  const totalProvisions = activeProvisions.reduce((s, p) => s + p.amount, 0);

  const recognize = () => {
    if (!form.name || !Number(form.amount)) return notify("Name the provision and enter an amount");
    const amount = Number(form.amount);
    const txn = buildTxn(`Provision recognized - ${form.name}`, todayStr(), [{ accountId: "5770", debit: amount, credit: 0 }, { accountId: "2295", debit: 0, credit: amount }], "manual");
    if (!txn) return notify("Entry not balanced");
    const id = uid("prov");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], provisions: [...d.provisions, { id, name: form.name, category: form.category, description: form.description, amount, recognizedDate: todayStr(), status: "active" }] }));
    setForm(blank());
    notify(`Provision recognized: ${fmt(amount)}`);
  };

  const openAction = (p, kind) => { setActioningId(actioningId === p.id + kind ? null : p.id + kind); setActionAmount(""); };

  const increase = async (p) => {
    const amt = Number(actionAmount);
    if (!amt) return notify("Enter an amount");
    if (!(await confirm(`Increase the provision for "${p.name}" by ${fmt(amt)}?`))) return;
    const txn = buildTxn(`Provision increased - ${p.name}`, todayStr(), [{ accountId: "5770", debit: amt, credit: 0 }, { accountId: "2295", debit: 0, credit: amt }], "manual");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], provisions: d.provisions.map(x => x.id === p.id ? { ...x, amount: x.amount + amt } : x) }));
    setActioningId(null); notify("Provision increased");
  };

  const utilize = async (p) => {
    const amt = Number(actionAmount);
    if (!amt || amt > p.amount + 0.5) return notify(`Enter an amount up to the provision balance of ${fmt(p.amount)}`);
    const bank = data.banks.find(b => b.id === actionBankId);
    if (!bank) return notify("Choose which bank paid it");
    if (!(await confirm(`Pay out ${fmt(amt)} against the provision for "${p.name}" from ${bank.name}?`))) return;
    const txn = buildTxn(`Provision utilized - ${p.name}`, todayStr(), [{ accountId: "2295", debit: amt, credit: 0 }, { accountId: bank.accountId, debit: 0, credit: amt }], "manual");
    if (!txn) return notify("Entry not balanced");
    const newAmount = p.amount - amt;
    setData(d => ({ ...d, transactions: [...d.transactions, txn], provisions: d.provisions.map(x => x.id === p.id ? { ...x, amount: newAmount, status: newAmount <= 0.5 ? "utilized" : "active" } : x) }));
    setActioningId(null); notify(`${fmt(amt)} paid out against the provision`);
  };

  const release = async (p) => {
    if (!(await confirm(`Release the entire remaining provision for "${p.name}" (${fmt(p.amount)})? Use this when the obligation no longer applies - it reverses the expense.`))) return;
    const txn = buildTxn(`Provision released - ${p.name}`, todayStr(), [{ accountId: "2295", debit: p.amount, credit: 0 }, { accountId: "5770", debit: 0, credit: p.amount }], "manual");
    if (!txn) return notify("Entry not balanced");
    setData(d => ({ ...d, transactions: [...d.transactions, txn], provisions: d.provisions.map(x => x.id === p.id ? { ...x, amount: 0, status: "released" } : x) }));
    notify("Provision released back to income");
  };

  const deleteProvision = async (p) => {
    if (!(await confirm(`Delete "${p.name}" entirely? All of its journal entries (recognition, any increases, utilizations, or release) are removed. This can't be undone.`))) return;
    setData(d => ({ ...d, provisions: d.provisions.filter(x => x.id !== p.id), transactions: d.transactions.filter(t => !(t.memo || "").endsWith(`- ${p.name}`)) }));
    notify("Provision and its journal entries deleted");
  };

  return (
    <div>
      <PageHeader eyebrow="Accounting" title="Provisions" sub={`${activeProvisions.length} active - ${fmt(totalProvisions)} provisioned`} />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Recognize a provision</div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Recognize one when there's a present obligation from a past event, an outflow is probable, and the amount can be reasonably estimated - a legal claim, a warranty commitment, a restructuring cost.</div>
        <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
          <input style={{ ...styles.input, flex: 1, minWidth: 180 }} placeholder="Provision name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
          <select style={styles.input} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>
            <option>Legal</option><option>Warranty</option><option>Restructuring</option><option>Other</option>
          </select>
          <input type="number" style={{ ...styles.input, width: 150 }} placeholder="Amount" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
          <button style={styles.btnPrimary} onClick={recognize}>Recognize</button>
        </div>
        <input style={{ ...styles.input, width: "100%", marginTop: 10 }} placeholder="Description (optional)" value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} />
      </div>

      {data.provisions.length > 0 && (
        <div className="ces-card" style={styles.card}>
          <table style={styles.table}>
            <thead><tr><th style={styles.th}>Provision</th><th style={styles.th}>Category</th><th style={{ ...styles.th, textAlign: "right" }}>Balance</th><th style={styles.th}>Status</th><th style={styles.th}></th></tr></thead>
            <tbody>{data.provisions.map(p => (
              <React.Fragment key={p.id}>
                <tr>
                  <td style={styles.td}>{p.name}{p.description && <div style={{ fontSize: 11, color: theme.muted }}>{p.description}</div>}</td>
                  <td style={styles.td}>{p.category}</td>
                  <td style={{ ...styles.tdMono, textAlign: "right" }}>{fmt(p.amount)}</td>
                  <td style={styles.td}><span style={{ ...styles.pill, ...(p.status === "active" ? styles.pillAmber : styles.pillGreen) }}>{p.status}</span></td>
                  <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                    {p.status === "active" && <>
                      <button style={styles.iconBtn} onClick={() => openAction(p, "inc")}>Increase</button>{" "}
                      <button style={styles.iconBtn} onClick={() => openAction(p, "util")}>Utilize</button>{" "}
                      <button style={styles.iconBtn} onClick={() => release(p)}>Release</button>{" "}
                    </>}
                    <button style={styles.iconBtn} onClick={() => deleteProvision(p)}>🗑</button>
                  </td>
                </tr>
                {actioningId === p.id + "inc" && (
                  <tr><td colSpan={5} style={{ ...styles.td, background: theme.panel2 }}>
                    <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                      <span style={{ fontSize: 13 }}>Increase by:</span>
                      <input type="number" style={{ ...styles.inputSmall, width: 140 }} value={actionAmount} onChange={e => setActionAmount(e.target.value)} />
                      <button style={styles.btnPrimary} onClick={() => increase(p)}>Confirm increase</button>
                    </div>
                  </td></tr>
                )}
                {actioningId === p.id + "util" && (
                  <tr><td colSpan={5} style={{ ...styles.td, background: theme.panel2 }}>
                    <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                      <span style={{ fontSize: 13 }}>Pay out (up to {fmt(p.amount)}):</span>
                      <input type="number" style={{ ...styles.inputSmall, width: 140 }} value={actionAmount} onChange={e => setActionAmount(e.target.value)} />
                      <span style={{ fontSize: 13 }}>from</span>
                      <select style={styles.inputSmall} value={actionBankId} onChange={e => setActionBankId(e.target.value)}>{data.banks.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
                      <button style={styles.btnPrimary} onClick={() => utilize(p)}>Confirm payment</button>
                    </div>
                  </td></tr>
                )}
              </React.Fragment>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}
function TaxPage({ data, setData, notify }) {
  const { styles, fmt, theme } = useUI();
  const balances = computeBalances(data.accounts, data.transactions);
  const taxAccounts = data.accounts.filter(a => a.type === "liability" && /tax|vat|wht|duty|paye/i.test(a.name));
  const nextMonth = new Date(); nextMonth.setMonth(nextMonth.getMonth() + 1);
  const dueLabel = `due ${new Date(nextMonth.getFullYear(), nextMonth.getMonth(), 21).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" })}`;
  return (
    <div>
      <PageHeader eyebrow="Accounting" title="Tax" sub="Balances, groups and filings" />
      <div className="ces-card" style={styles.card}>
        <div style={styles.cardTitle}>Upcoming filings</div>
        <div style={{ fontSize: 12, color: theme.muted, marginTop: 4 }}>Estimated from the current balance of each tax liability account. Filing dates shown assume the common 21st-of-following-month deadline - confirm against your jurisdiction's calendar.</div>
        <table style={{ ...styles.table, marginTop: 10 }}>
          <thead><tr><th style={styles.th}>Tax account</th><th style={{ ...styles.th, textAlign: "right" }}>Balance payable</th><th style={styles.th}>Indicative deadline</th></tr></thead>
          <tbody>{taxAccounts.map(a => (
            <tr key={a.id}><td style={styles.td}>{a.name}</td><td style={{ ...styles.tdMono, textAlign: "right", fontWeight: 600 }}>{fmt(balances[a.id] || 0)}</td>
              <td style={{ ...styles.td, color: (balances[a.id] || 0) > 0.5 ? theme.amber : theme.muted }}>{(balances[a.id] || 0) > 0.5 ? dueLabel : "nothing due"}</td></tr>
          ))}</tbody>
        </table>
        {taxAccounts.length === 0 && <div style={{ color: theme.muted, fontSize: 13, marginTop: 10 }}>No tax liability accounts found.</div>}
      </div>
      <TaxGroupsManager data={data} setData={setData} notify={notify} />
      <DeferredTaxPanel data={data} setData={setData} notify={notify} />
      <ReportTaxSummary data={data} range="year" />
    </div>
  );
}

/* =================================== Presentation =================================== */
// Board-pack mode: large-type slides built live from the books, for walking a
// leadership team or board through the numbers.
function Presentation({ data, balances }) {
  const { styles, fmt, theme } = useUI();
  const [slide, setSlide] = useState(0);
  const year = new Date().getFullYear();
  const { from, to } = rangeToDates("year");
  const mv = periodMovement(data.accounts, data.transactions, from, to);
  const rev = data.accounts.filter(a => a.type === "revenue").reduce((s, a) => s + (a.contra ? -(mv[a.id] || 0) : (mv[a.id] || 0)), 0);
  const exp = data.accounts.filter(a => a.type === "expense").reduce((s, a) => s + (a.contra ? -(mv[a.id] || 0) : (mv[a.id] || 0)), 0);
  const cash = data.banks.reduce((s, b) => s + (balances[b.accountId] || 0), 0);
  const ar = balances["1100"] || 0, ap = balances["2000"] || 0;
  const topCustomers = Object.entries(data.invoices.reduce((acc, i) => { const t = computeDocTotals(i, data.taxGroups).finalAmount; acc[i.customer] = (acc[i.customer] || 0) + t; return acc; }, {})).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const Big = ({ label, value, tone }) => (
    <div style={{ textAlign: "center" }}>
      <div style={{ fontSize: 14, textTransform: "uppercase", letterSpacing: "0.08em", color: theme.muted }}>{label}</div>
      <div style={{ fontSize: 44, fontWeight: 700, letterSpacing: "-0.02em", fontVariantNumeric: "tabular-nums", color: tone || theme.text, marginTop: 6 }}>{value}</div>
    </div>
  );
  const slides = [
    { title: "", body: (
      <div style={{ textAlign: "center", padding: "60px 0" }}>
        {data.settings.logo && <div className="ces-logo" style={{ padding: 14, display: "inline-flex", marginBottom: 20 }}><img src={data.settings.logo} alt="logo" style={{ height: 72, objectFit: "contain" }} /></div>}
        <div style={{ fontSize: 42, fontWeight: 700, letterSpacing: "-0.02em" }}>{data.settings.companyName}</div>
        <div style={{ fontSize: 17, color: theme.muted, marginTop: 10 }}>Financial review · {year}</div>
      </div>
    ) },
    { title: "Headline numbers", body: (
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 40, padding: "40px 0" }}>
        <Big label="Revenue (YTD)" value={fmt(rev)} tone={theme.emerald} />
        <Big label="Expenses (YTD)" value={fmt(exp)} tone={theme.rose} />
        <Big label="Net income" value={fmt(rev - exp)} tone={rev - exp >= 0 ? theme.emerald : theme.rose} />
      </div>
    ) },
    { title: "Liquidity", body: (
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 40, padding: "40px 0" }}>
        <Big label="Cash & bank" value={fmt(cash)} />
        <Big label="Receivables" value={fmt(ar)} tone={theme.amber} />
        <Big label="Payables" value={fmt(ap)} tone={theme.rose} />
      </div>
    ) },
    { title: "Top customers by billing", body: (
      <div style={{ maxWidth: 560, margin: "30px auto" }}>
        {topCustomers.map(([name, total], i) => (
          <div key={name} style={{ display: "flex", justifyContent: "space-between", padding: "14px 0", borderBottom: `1px solid ${theme.border}`, fontSize: 19 }}>
            <span><span style={{ color: theme.muted, marginRight: 12 }}>{i + 1}.</span>{name}</span>
            <span style={{ fontVariantNumeric: "tabular-nums", fontWeight: 600 }}>{fmt(total)}</span>
          </div>
        ))}
        {topCustomers.length === 0 && <div style={{ color: theme.muted, textAlign: "center" }}>No invoices yet.</div>}
      </div>
    ) },
  ];
  return (
    <div>
      <PageHeader eyebrow="Insights" title="Presentation" sub="Live board-pack built from your books" action={
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <button style={styles.btnGhost} onClick={() => setSlide(s => Math.max(0, s - 1))}>← Prev</button>
          <span style={{ fontSize: 12.5, color: theme.muted, fontVariantNumeric: "tabular-nums" }}>{slide + 1} / {slides.length}</span>
          <button style={styles.btnGhost} onClick={() => setSlide(s => Math.min(slides.length - 1, s + 1))}>Next →</button>
          <button style={styles.btnPrimary} onClick={() => window.print()}>Export slide</button>
        </div>
      } />
      <div className="print-area">
        <div className="ces-card" style={{ ...styles.card, minHeight: 440, display: "flex", flexDirection: "column", justifyContent: "center" }}>
          {slides[slide].title && <div style={{ fontSize: 13, textTransform: "uppercase", letterSpacing: "0.08em", color: theme.accent, fontWeight: 700, textAlign: "center" }}>{slides[slide].title}</div>}
          {slides[slide].body}
        </div>
      </div>
      <div style={{ fontSize: 12, color: theme.muted }}>Figures are computed live from the ledger - what you present is always what the books say.</div>
    </div>
  );
}

function makeStyles(theme) {
  const isDark = theme.mode === "dark";
  return {
    app: { display: "flex", height: "100vh", overflow: "hidden", background: theme.paper, fontFamily: "Inter, sans-serif", color: theme.text },
    sidebar: { width: 244, flexShrink: 0, height: "100vh", overflowY: "auto", background: theme.sidebarBg, borderRight: `1px solid ${theme.sidebarBorder}`, color: theme.sidebarText, padding: "20px 12px", display: "flex", flexDirection: "column", transition: "width 0.18s ease, padding 0.18s ease" },
    brand: { display: "flex", alignItems: "center", gap: 10, padding: "2px 8px 16px", borderBottom: `1px solid ${theme.sidebarBorder}` },
    brandMark: { width: 32, height: 32, borderRadius: 9, background: `linear-gradient(135deg, ${theme.accent}, ${shade(theme.accent, -18)})`, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontWeight: 600, fontSize: 13, color: "#fff", boxShadow: `0 2px 6px ${theme.accent}40` },
    brandName: { fontFamily: "Inter, sans-serif", letterSpacing: "-0.01em", fontSize: 14.5, fontWeight: 600, lineHeight: 1.25, color: theme.text },
    brandSub: { fontSize: 10.5, color: theme.sidebarMuted, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums" },
    navGroup: { fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.06em", color: theme.sidebarMuted, padding: "14px 10px 5px", fontWeight: 600 },
    navItem: { display: "flex", alignItems: "center", gap: 8, width: "100%", textAlign: "left", background: "transparent", border: "none", color: theme.sidebarText, padding: "7px 9px", borderRadius: 8, fontSize: 13, cursor: "pointer", marginBottom: 1, fontWeight: 500 },
    navItemActive: { background: theme.accentSoft, color: theme.accent, fontWeight: 600 },
    sidebarFoot: { paddingTop: 14, borderTop: `1px solid ${theme.sidebarBorder}` },
    main: { flex: 1, minWidth: 0, height: "100vh", padding: "28px 36px", overflowY: "auto" },
    kpiRow: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px,1fr))", gap: 16, marginBottom: 22 },
    kpi: { background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 12, padding: "16px 18px", boxShadow: theme.shadowSm, position: "relative", overflow: "hidden" },
    card: { background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 12, padding: "20px 22px", marginBottom: 20, boxShadow: theme.shadowMd, position: "relative", overflow: "hidden" },
    cardWide: { background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 12, padding: "20px 22px", marginBottom: 20, boxShadow: theme.shadowMd, position: "relative", overflow: "hidden" },
    cardTitle: { fontFamily: "Inter, sans-serif", fontSize: 14.5, fontWeight: 600, color: theme.text, letterSpacing: "-0.005em" },
    table: { width: "100%", fontSize: 13.5 },
    th: { textAlign: "left", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.muted, background: theme.panel2, borderBottom: `1px solid ${theme.border}`, padding: "9px 8px", fontWeight: 700 },
    td: { padding: "9px 8px", borderBottom: `1px solid ${theme.border}` },
    tdMono: { padding: "9px 8px", borderBottom: `1px solid ${theme.border}`, fontFamily: "Inter, sans-serif", fontVariantNumeric: "tabular-nums", fontSize: 13 },
    input: { border: `1px solid ${theme.border}`, borderRadius: 7, padding: "8px 10px", fontSize: 13.5, background: isDark ? theme.panel2 : "#FFFFFF", color: theme.text },
    inputSmall: { border: `1px solid ${theme.border}`, borderRadius: 6, padding: "5px 8px", fontSize: 13, background: isDark ? theme.panel2 : "#FFFFFF", color: theme.text },
    btnPrimary: { background: `linear-gradient(180deg, ${shade(theme.accent, 6)}, ${theme.accent})`, color: "#fff", border: `1px solid ${shade(theme.accent, -10)}`, borderRadius: 8, padding: "8px 16px", fontSize: 13.5, fontWeight: 600, cursor: "pointer", boxShadow: `0 1px 2px rgba(16,24,40,0.1), inset 0 1px 0 rgba(255,255,255,0.15)` },
    btnGhost: { background: isDark ? "transparent" : "#FFFFFF", color: theme.text, border: `1px solid ${theme.border}`, borderRadius: 8, padding: "7px 12px", fontSize: 13, cursor: "pointer", boxShadow: theme.shadowSm },
    iconBtn: { background: "transparent", border: "none", color: theme.muted, cursor: "pointer", fontSize: 13, marginRight: 4 },
    pill: { fontSize: 10.5, padding: "3px 9px", borderRadius: 20, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em" },
    pillGreen: { background: isDark ? "rgba(52,194,131,0.15)" : "#E4F7EF", color: theme.emerald },
    pillAmber: { background: isDark ? "rgba(216,154,75,0.18)" : "#FBEFDC", color: theme.amber },
    pillRose: { background: isDark ? "rgba(224,97,109,0.18)" : "#FBE4E7", color: theme.rose },
    pillAmberSm: { fontSize: 10, background: isDark ? "rgba(216,154,75,0.18)" : "#FBEFDC", color: theme.amber, padding: "2px 7px", borderRadius: 20 },
    attentionRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${theme.border}` },
    dropzone: { border: `1.5px dashed ${theme.border}`, borderRadius: 10, padding: "30px 20px", textAlign: "center", color: theme.muted, marginTop: 14, background: theme.panel2 },
    toast: { position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: theme.ink, color: "#fff", padding: "10px 18px", borderRadius: 8, fontSize: 13.5, boxShadow: "0 6px 20px rgba(0,0,0,0.3)" },
  };
}
