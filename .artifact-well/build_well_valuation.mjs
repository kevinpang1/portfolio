import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outDir = path.resolve("..", "outputs/01a06864-0723-7a51-a935-b11fb1d3179e");
await fs.mkdir(outDir, { recursive: true });

const wb = Workbook.create();
wb.comments.setSelf({ displayName: "Kevin" });

const summary = wb.worksheets.add("Summary");
const sources = wb.worksheets.add("Source Data");
const assumptions = wb.worksheets.add("Assumptions");
const historical = wb.worksheets.add("Historical");
const operating = wb.worksheets.add("Operating Model");
const nav = wb.worksheets.add("NAV");
const dcf = wb.worksheets.add("AFFO DCF");
const multiples = wb.worksheets.add("Multiples");
const valuation = wb.worksheets.add("Valuation");
const checks = wb.worksheets.add("Checks");

const COLORS = {
  navy: "#0B1F33",
  navy2: "#173A5E",
  teal: "#0F8B8D",
  tealLight: "#DCEFF0",
  blueLight: "#D9EAF7",
  input: "#FFF2CC",
  forecast: "#EAF2F8",
  gray: "#F3F4F6",
  gray2: "#D1D5DB",
  green: "#008000",
  greenFill: "#E2F0D9",
  redFill: "#FCE4D6",
  amberFill: "#FFF2CC",
  white: "#FFFFFF",
  black: "#000000",
  blue: "#0000FF",
  red: "#C00000",
};

const fmt = {
  money: '$#,##0.0;[Red]($#,##0.0);-',
  money0: '$#,##0;[Red]($#,##0);-',
  perShare: '$0.00;[Red]($0.00);-',
  pct: '0.0%;[Red](0.0%);-',
  pct1: '0.0%',
  multiple: '0.0x;[Red](0.0x);-',
  number: '#,##0.0;[Red](#,##0.0);-',
  number0: '#,##0;[Red](#,##0);-',
};

function title(sheet, range, text) {
  sheet.getRange(range).merge();
  sheet.getRange(range.split(":")[0]).values = [[text]];
  sheet.getRange(range).format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 18 },
    verticalAlignment: "center",
  };
  sheet.getRange(range).format.rowHeight = 32;
}

function section(sheet, range, text) {
  sheet.getRange(range).merge();
  sheet.getRange(range.split(":")[0]).values = [[text]];
  sheet.getRange(range).format = {
    fill: COLORS.navy2,
    font: { bold: true, color: COLORS.white },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: COLORS.navy2 },
  };
  sheet.getRange(range).format.rowHeight = 22;
}

function header(range) {
  range.format = {
    fill: COLORS.teal,
    font: { bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#B7CED0" },
  };
}

function total(range) {
  range.format = {
    font: { bold: true, color: COLORS.black },
    borders: { top: { style: "thin", color: COLORS.black } },
  };
}

function input(range, numberFormat) {
  range.format = {
    fill: COLORS.input,
    font: { color: COLORS.blue },
    horizontalAlignment: "right",
    ...(numberFormat ? { numberFormat } : {}),
  };
}

function crossSheet(range, numberFormat) {
  range.format = {
    font: { color: COLORS.green },
    horizontalAlignment: "right",
    ...(numberFormat ? { numberFormat } : {}),
  };
}

function body(range, numberFormat) {
  range.format = {
    font: { color: COLORS.black },
    horizontalAlignment: "right",
    ...(numberFormat ? { numberFormat } : {}),
  };
}

function sourceComment(sheet, cell, label, url) {
  wb.comments.addThread({ cell: sheet.getRange(cell) }, `${label}\nSource: ${url}`);
}

for (const sheet of wb.worksheets.items) {
  sheet.showGridLines = false;
}

// ---------------- Source Data ----------------
title(sources, "A1:F1", "Welltower Valuation Model — Source Data & Audit Trail");
sources.getRange("A2:F2").merge();
sources.getRange("A2").values = [["USD millions except per-share data, ratios, and operating metrics. Market data as of 2026-09-02; company data through 2Q26."]];
sources.getRange("A2:F2").format = { font: { italic: true, color: "#4B5563" }, wrapText: true };

section(sources, "A4:F4", "Market data and 2026 guidance");
sources.getRange("A5:F5").values = [["Metric", "Value", "Units", "Period / as of", "Source ID", "Notes"]];
header(sources.getRange("A5:F5"));
const marketRows = [
  ["Market price", 238.59, "$ / share", "2026-09-02 close", "S1", "Official close reported by FinancialContent"],
  ["Normalized FFO / share — low", 6.36, "$ / share", "FY2026 guidance", "S2", "Company guidance"],
  ["Normalized FFO / share — high", 6.44, "$ / share", "FY2026 guidance", "S2", "Company guidance"],
  ["Normalized FFO — low", 4713, "$mm", "FY2026 guidance", "S2", "Company guidance reconciliation"],
  ["Normalized FFO — high", 4773, "$mm", "FY2026 guidance", "S2", "Company guidance reconciliation"],
  ["Straight-line rent and lease amortization", -341, "$mm", "FY2026 guidance", "S2", "AFFO proxy adjustment"],
  ["Non-cash interest expense", 50, "$mm", "FY2026 guidance", "S2", "AFFO proxy adjustment"],
  ["Recurring capex, TIs and LCs", -465, "$mm", "FY2026 guidance", "S2", "AFFO proxy adjustment"],
  ["Fully diluted shares for NAV", 740.776, "mm shares", "2026-06-30", "S3", "Includes common, OP, exchangeable units and dilutive notes"],
  ["Net obligations", 17232.814, "$mm", "2026-06-30", "S3", "Company NAV components"],
  ["Net debt", 16121.380, "$mm", "2026-06-30", "S3", "Debt less cash and restricted cash"],
  ["Annualized adjusted EBITDA", 5390.052, "$mm", "2026-06-30", "S3", "Company leverage reconciliation"],
  ["Net debt / adjusted EBITDA", 2.99, "x", "2026-06-30", "S3", "Company-reported ratio"],
  ["Risk-free rate", 0.0475, "%", "2026-09-01", "S4", "Damodaran US dollar risk-free rate"],
  ["Implied equity risk premium", 0.0414, "%", "2026-09-01", "S4", "Trailing 12-month adjusted-payout ERP"],
  ["2Q26 normalized FFO / share", 1.60, "$ / share", "2Q26", "S2", "Quarterly reported result"],
  ["2Q26 SHO same-store NOI growth", 0.205, "%", "2Q26 YoY", "S2", "Reported result"],
  ["2Q26 total portfolio SSNOI growth", 0.155, "%", "2Q26 YoY", "S2", "Reported result"],
  ["2Q26 SHO occupancy", 0.888, "%", "2026-06-30", "S3", "Stable portfolio occupancy"],
  ["2Q26 SHO margin", 0.321, "%", "2Q26", "S5", "Business update"],
];
sources.getRange("A6:F25").values = marketRows;
sources.getRange("B6:B21").format.numberFormat = fmt.number;
sources.getRange("B19:B20").format.numberFormat = fmt.pct;
sources.getRange("B22:B25").format.numberFormat = fmt.pct;

section(sources, "A27:F27", "Historical financial and cash-flow inputs");
sources.getRange("A28:E28").values = [["Metric", 2022, 2023, 2024, 2025]];
header(sources.getRange("A28:E28"));
sources.getRange("A29:E37").values = [
  ["Total revenue", 5860.615, 6637.995, 7991.118, 10838.034],
  ["Property operating expense", 3558.770, 3947.776, 4830.211, 6488.081],
  ["Normalized FFO", 1558.270, 1885.544, 2626.757, 3591.666],
  ["Average diluted shares", 465.158, 518.701, 608.750, 679.521],
  ["Normalized FFO / share", 3.35, 3.64, 4.32, 5.29],
  ["Straight-line rent and lease amortization", -106.496, -135.356, -156.460, -221.708],
  ["Non-cash interest expense", 21.805, 27.252, 44.335, 51.230],
  ["Recurring capex, TIs and LCs", -179.133, -199.359, -286.613, -370.693],
  ["Dividends / share", 2.44, 2.44, 2.56, 2.82],
];
sources.getRange("B29:E37").format.numberFormat = fmt.number;

section(sources, "A39:F39", "2Q26 NAV components");
sources.getRange("A40:F40").values = [["Metric", "Value", "Units", "Period / as of", "Source ID", "Notes"]];
header(sources.getRange("A40:F40"));
sources.getRange("A41:F53").values = [
  ["SHO annualized in-place NOI", 3402.972, "$mm", "2Q26 annualized", "S3", "170,828 units"],
  ["SH-NNN annualized in-place NOI", 623.896, "$mm", "2Q26 annualized", "S3", "28,649 units"],
  ["Outpatient Medical annualized in-place NOI", 116.848, "$mm", "2Q26 annualized", "S3", "4.13mm square feet"],
  ["LT / PAC annualized in-place NOI", 721.952, "$mm", "2Q26 annualized", "S3", "47,408 beds"],
  ["Incremental stabilized NOI", 152.032, "$mm", "2026-06-30", "S3", "From unstabilized SHO properties"],
  ["Land parcels", 637.295, "$mm", "2026-06-30", "S3", "Book / carrying basis"],
  ["Real estate loans receivable", 4341.954, "$mm", "2026-06-30", "S3", "Net of credit allowances"],
  ["Non-real-estate loans receivable", 230.250, "$mm", "2026-06-30", "S3", "Net of credit allowances"],
  ["JV real estate loans receivable", 225.578, "$mm", "2026-06-30", "S3", "Partners' share of WELL loans"],
  ["Expected property dispositions", 798.541, "$mm", "Next 12 months", "S3", "Expected proceeds"],
  ["Committed development balances", 1469.573, "$mm", "2026-06-30", "S3", "Current plus unfunded commitments"],
  ["Net obligations", 17232.814, "$mm", "2026-06-30", "S3", "Debt and other obligations less cash"],
  ["Fully diluted shares", 740.776, "mm shares", "2026-06-30", "S3", "NAV share count"],
];
sources.getRange("B41:B53").format.numberFormat = fmt.number;

section(sources, "A55:F55", "Source URLs");
sources.getRange("A56:F56").values = [["Source ID", "Source", "URL", "Used for", "Accessed", "Caveat"]];
header(sources.getRange("A56:F56"));
sources.getRange("A57:F64").values = [
  ["S1", "FinancialContent", "https://markets.financialcontent.com/stocks/quote/detailedquote?Symbol=NY%3AWELL", "Market price", new Date("2026-09-03T00:00:00Z"), "Market-data source; close differs slightly across delayed aggregators"],
  ["S2", "Welltower 2Q26 earnings release", "https://welltower.com/investors/press-release-details/?id=808", "2026 guidance, FFO, SSNOI, leverage", new Date("2026-09-03T00:00:00Z"), "Company non-GAAP measures"],
  ["S3", "Welltower 2Q26 supplemental", "https://welltower.com/wp-content/uploads/2026/07/2Q26-Supplement-99.2-FINAL.pdf", "NAV, NOI, debt, shares, portfolio metrics", new Date("2026-09-03T00:00:00Z"), "Pro rata data; may differ from GAAP consolidation"],
  ["S4", "Damodaran data", "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm", "Risk-free rate and ERP", new Date("2026-09-03T00:00:00Z"), "Inputs are market-sensitive"],
  ["S5", "Welltower 2Q26 business update", "https://welltower.com/wp-content/uploads/2026/07/Business-Update-2Q26_vFF.pdf", "Operating KPIs and strategic context", new Date("2026-09-03T00:00:00Z"), "Forward-looking statements"],
  ["S6", "Welltower 2025 Form 10-K", "https://www.sec.gov/Archives/edgar/data/766704/000076670426000010/well-20251231.htm", "2023-2025 GAAP historicals", new Date("2026-09-03T00:00:00Z"), "Audited GAAP financials"],
  ["S7", "Welltower 2023 Form 10-K", "https://www.sec.gov/Archives/edgar/data/766704/000076670424000008/well-20231231.htm", "2022-2023 GAAP historicals", new Date("2026-09-03T00:00:00Z"), "Audited GAAP financials"],
  ["S8", "Welltower annual earnings releases", "https://welltower.com/investors/financial-summary/", "2022-2025 FFO and cash adjustments", new Date("2026-09-03T00:00:00Z"), "Company non-GAAP measures"],
];
sources.getRange("E57:E64").format.numberFormat = "yyyy-mm-dd";
sources.getRange("A57:F64").format.wrapText = true;

// Source comments on every major input block.
for (let r = 6; r <= 25; r++) sourceComment(sources, `B${r}`, `Source ${marketRows[r - 6][4]}: ${marketRows[r - 6][0]}`, r === 6 ? "https://markets.financialcontent.com/stocks/quote/detailedquote?Symbol=NY%3AWELL" : (r === 19 || r === 20 ? "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm" : "https://welltower.com/investors/press-release-details/?id=808"));
for (let c of ["B", "C", "D", "E"]) {
  for (let r = 29; r <= 37; r++) sourceComment(sources, `${c}${r}`, "Historical company financial input", r <= 30 ? (c === "B" ? "https://www.sec.gov/Archives/edgar/data/766704/000076670424000008/well-20231231.htm" : "https://www.sec.gov/Archives/edgar/data/766704/000076670426000010/well-20251231.htm") : "https://welltower.com/investors/financial-summary/");
}
for (let r = 41; r <= 53; r++) sourceComment(sources, `B${r}`, "Welltower 2Q26 supplemental NAV input", "https://welltower.com/wp-content/uploads/2026/07/2Q26-Supplement-99.2-FINAL.pdf");

// ---------------- Assumptions ----------------
title(assumptions, "A1:E1", "Assumptions & Scenario Control");
assumptions.getRange("A2:E2").merge();
assumptions.getRange("A2").values = [["Yellow cells with blue font are editable. Base case is selected for the operating model; all valuation tabs show Bear / Base / Bull simultaneously."]];
assumptions.getRange("A2:E2").format = { fill: COLORS.gray, font: { italic: true }, wrapText: true };
assumptions.getRange("A3:B3").values = [["Selected scenario", "Base"]];
assumptions.getRange("B3").dataValidation = { rule: { type: "list", values: ["Bear", "Base", "Bull"] } };
input(assumptions.getRange("B3"));
assumptions.getRange("A5:E5").values = [["Driver", "Units", "Bear", "Base", "Bull"]];
header(assumptions.getRange("A5:E5"));
const scenarioRows = [
  ["2026E normalized FFO / share", "$ / share", 6.36, 6.40, 6.44],
  ["2027 FFO / share growth", "%", 0.07, 0.10, 0.13],
  ["2028 FFO / share growth", "%", 0.06, 0.09, 0.11],
  ["2029 FFO / share growth", "%", 0.05, 0.08, 0.10],
  ["2030 FFO / share growth", "%", 0.04, 0.07, 0.08],
  ["2031 FFO / share growth", "%", 0.03, 0.06, 0.07],
  ["AFFO proxy conversion", "% of NFFO", 0.82, 3987 / 4743, 0.86],
  ["Diluted share growth", "%", 0.02, 0.01, 0.00],
  ["Dividend / share growth", "%", 0.04, 0.07, 0.10],
  ["Risk-free rate", "%", 0.0475, 0.0475, 0.0475],
  ["Equity beta", "x", 1.10, 1.00, 0.90],
  ["Equity risk premium", "%", 0.0414, 0.0414, 0.0414],
  ["Cost of equity", "%", null, null, null],
  ["Terminal growth", "%", 0.030, 0.040, 0.045],
  ["2027E target P / FFO", "x", 30.0, 34.0, 38.0],
  ["2027E target AFFO yield", "%", 0.0325, 0.0275, 0.0235],
  ["Dividend payout ceiling", "% of AFFO", 0.70, 0.65, 0.60],
];
assumptions.getRange("A6:E22").values = scenarioRows;
assumptions.getRange("C18:E18").formulas = [["=C15+C16*C17", "=D15+D16*D17", "=E15+E16*E17"]];
input(assumptions.getRange("C6:E17"));
body(assumptions.getRange("C18:E18"), fmt.pct);
input(assumptions.getRange("C19:E22"));
assumptions.getRange("C6:E6").format.numberFormat = fmt.perShare;
assumptions.getRange("C7:E15").format.numberFormat = fmt.pct;
assumptions.getRange("C16:E16").format.numberFormat = fmt.multiple;
assumptions.getRange("C17:E19").format.numberFormat = fmt.pct;
assumptions.getRange("C20:E20").format.numberFormat = fmt.multiple;
assumptions.getRange("C21:E22").format.numberFormat = fmt.pct;
sourceComment(assumptions, "C6", "Bear 2026 FFO equals company guidance low end", "https://welltower.com/investors/press-release-details/?id=808");
sourceComment(assumptions, "D6", "Base 2026 FFO equals company guidance midpoint", "https://welltower.com/investors/press-release-details/?id=808");
sourceComment(assumptions, "E6", "Bull 2026 FFO equals company guidance high end", "https://welltower.com/investors/press-release-details/?id=808");
sourceComment(assumptions, "D12", "Base AFFO conversion derived from 2026 guidance FFO less disclosed cash adjustments", "https://welltower.com/investors/press-release-details/?id=808");

section(assumptions, "A24:E24", "NAV cap-rate assumptions");
assumptions.getRange("A25:E25").values = [["Segment", "Units", "Bear", "Base", "Bull"]];
header(assumptions.getRange("A25:E25"));
assumptions.getRange("A26:E30").values = [
  ["Seniors Housing Operating", "%", 0.0550, 0.0475, 0.0400],
  ["Seniors Housing Triple-net", "%", 0.0600, 0.0525, 0.0450],
  ["Outpatient Medical", "%", 0.0650, 0.0575, 0.0500],
  ["Long-Term / Post-Acute Care", "%", 0.0800, 0.0700, 0.0600],
  ["Incremental stabilized NOI", "%", 0.0625, 0.0525, 0.0450],
];
input(assumptions.getRange("C26:E30"), fmt.pct);

section(assumptions, "A32:E32", "Other-asset valuation haircuts");
assumptions.getRange("A33:E33").values = [["Asset", "Units", "Bear", "Base", "Bull"]];
header(assumptions.getRange("A33:E33"));
assumptions.getRange("A34:E39").values = [
  ["Land parcels", "% of carrying value", 0.70, 0.80, 0.90],
  ["Real estate loans", "% of carrying value", 0.95, 1.00, 1.00],
  ["Non-real-estate loans", "% of carrying value", 0.80, 0.90, 1.00],
  ["JV real estate loans", "% of carrying value", 0.90, 1.00, 1.00],
  ["Expected dispositions", "% of expected proceeds", 0.95, 1.00, 1.00],
  ["Development balances", "% of committed balances", 0.90, 1.00, 1.10],
];
input(assumptions.getRange("C34:E39"), fmt.pct);

section(assumptions, "A41:E41", "Composite valuation weights");
assumptions.getRange("A42:E42").values = [["Method", "Units", "Bear", "Base", "Bull"]];
header(assumptions.getRange("A42:E42"));
assumptions.getRange("A43:E46").values = [
  ["P / FFO", "%", 0.40, 0.40, 0.40],
  ["AFFO yield", "%", 0.15, 0.15, 0.15],
  ["AFFO DCF", "%", 0.20, 0.20, 0.20],
  ["NAV", "%", 0.25, 0.25, 0.25],
];
input(assumptions.getRange("C43:E46"), fmt.pct);
assumptions.getRange("A47:E47").values = [["Weight check", "", null, null, null]];
assumptions.getRange("C47:E47").formulas = [["=SUM(C43:C46)", "=SUM(D43:D46)", "=SUM(E43:E46)"]];
total(assumptions.getRange("A47:E47"));
body(assumptions.getRange("C47:E47"), fmt.pct);

section(assumptions, "A49:E49", "Price-target conventions");
assumptions.getRange("A50:B50").values = [["Convention", "Value"]];
header(assumptions.getRange("A50:B50"));
assumptions.getRange("A51:A54").values = [["Valuation date"],["Target date"],["Target horizon (years)"],["Target horizon (months)"]];
assumptions.getRange("B51:B52").values = [[new Date("2026-09-02T00:00:00Z")],[new Date("2027-12-31T00:00:00Z")]];
input(assumptions.getRange("B51:B52"), "yyyy-mm-dd");
assumptions.getRange("B53").formulas = [["=('Assumptions'!B52-'Assumptions'!B51+1)/365"]];
assumptions.getRange("B54").formulas = [["='Assumptions'!B53*12"]];
body(assumptions.getRange("B53"), '0.00 "yrs"');
body(assumptions.getRange("B54"), '0.0 "months"');

// ---------------- Historical ----------------
title(historical, "A1:F1", "Historical Financial Profile & AFFO Proxy");
historical.getRange("A2:F2").merge();
historical.getRange("A2").values = [["AFFO proxy = normalized FFO + straight-line rent/lease amortization + non-cash interest expense + recurring capex/TIs/LCs. WELL does not report AFFO; this model-derived proxy is explicitly labeled."]];
historical.getRange("A2:F2").format = { fill: COLORS.gray, font: { italic: true }, wrapText: true };
historical.getRange("A4:F4").values = [["USD millions except per-share data", new Date("2022-12-31"), new Date("2023-12-31"), new Date("2024-12-31"), new Date("2025-12-31"), new Date("2026-12-31")]];
header(historical.getRange("A4:F4"));
historical.getRange("B4:F4").format.numberFormat = 'yyyy"A"';
historical.getRange("F4").format.numberFormat = 'yyyy"E"';
historical.getRange("A5:A18").values = [
  ["Total revenue"], ["Property operating expense"], ["NOI proxy"], ["NOI margin"], [""],
  ["Normalized FFO"], ["Average diluted shares"], ["Normalized FFO / share"],
  ["Straight-line rent / lease amortization"], ["Non-cash interest expense"], ["Recurring capex, TIs and LCs"],
  ["AFFO proxy"], ["AFFO proxy / share"], ["AFFO conversion"]
];
historical.getRange("B5:E6").formulas = [
  ["='Source Data'!B29", "='Source Data'!C29", "='Source Data'!D29", "='Source Data'!E29"],
  ["='Source Data'!B30", "='Source Data'!C30", "='Source Data'!D30", "='Source Data'!E30"],
];
historical.getRange("B7:E7").formulas = [["=B5-B6", "=C5-C6", "=D5-D6", "=E5-E6"]];
historical.getRange("B8:E8").formulas = [["=B7/B5", "=C7/C5", "=D7/D5", "=E7/E5"]];
historical.getRange("B10:E15").formulas = [
  ["='Source Data'!B31", "='Source Data'!C31", "='Source Data'!D31", "='Source Data'!E31"],
  ["='Source Data'!B32", "='Source Data'!C32", "='Source Data'!D32", "='Source Data'!E32"],
  ["='Source Data'!B33", "='Source Data'!C33", "='Source Data'!D33", "='Source Data'!E33"],
  ["='Source Data'!B34", "='Source Data'!C34", "='Source Data'!D34", "='Source Data'!E34"],
  ["='Source Data'!B35", "='Source Data'!C35", "='Source Data'!D35", "='Source Data'!E35"],
  ["='Source Data'!B36", "='Source Data'!C36", "='Source Data'!D36", "='Source Data'!E36"],
];
historical.getRange("B16:E16").formulas = [["=SUM(B10,B13:B15)", "=SUM(C10,C13:C15)", "=SUM(D10,D13:D15)", "=SUM(E10,E13:E15)"]];
historical.getRange("B17:E17").formulas = [["=B16/B11", "=C16/C11", "=D16/D11", "=E16/E11"]];
historical.getRange("B18:E18").formulas = [["=B16/B10", "=C16/C10", "=D16/D10", "=E16/E10"]];
historical.getRange("F10:F15").formulas = [
  ["=AVERAGE('Source Data'!B9:B10)"], ["=F10/F12"], ["=AVERAGE('Source Data'!B7:B8)"],
  ["='Source Data'!B11"], ["='Source Data'!B12"], ["='Source Data'!B13"]
];
historical.getRange("F16").formulas = [["=SUM(F10,F13:F15)"]];
historical.getRange("F17").formulas = [["=F16/F11"]];
historical.getRange("F18").formulas = [["=F16/F10"]];
historical.getRange("B5:F18").format.numberFormat = fmt.money;
historical.getRange("B8:F8").format.numberFormat = fmt.pct;
historical.getRange("B11:F11").format.numberFormat = fmt.number;
historical.getRange("B12:F12").format.numberFormat = fmt.perShare;
historical.getRange("B17:F17").format.numberFormat = fmt.perShare;
historical.getRange("B18:F18").format.numberFormat = fmt.pct;
crossSheet(historical.getRange("B5:E6"), fmt.money);
crossSheet(historical.getRange("B10:F15"), fmt.money);
historical.getRange("F5:F8").format = { fill: COLORS.gray, font: { italic: true, color: "#6B7280" } };
historical.getRange("F5:F8").values = [["n/a"], ["n/a"], ["n/a"], ["n/a"]];
total(historical.getRange("A7:F8"));
total(historical.getRange("A16:F18"));

// ---------------- Operating Model ----------------
title(operating, "A1:K1", "Operating Model — Selected Scenario");
operating.getRange("A2:K2").merge();
operating.getRange("A2").formulas = [["=\"Selected scenario: \"&'Assumptions'!B3&\" | Forecast is per-share-led and explicitly includes dilution.\""]];
crossSheet(operating.getRange("A2:K2"));
const years = [2022,2023,2024,2025,2026,2027,2028,2029,2030,2031];
operating.getRange("A4:K4").values = [["USD millions except per-share data", ...years.map(y => new Date(`${y}-12-31`))]];
header(operating.getRange("A4:K4"));
operating.getRange("B4:E4").format.numberFormat = 'yyyy"A"';
operating.getRange("F4:K4").format.numberFormat = 'yyyy"E"';
operating.getRange("A5:A13").values = [
  ["Normalized FFO / share"], ["Growth"], ["AFFO conversion"], ["AFFO proxy / share"],
  ["Average diluted shares"], ["Normalized FFO"], ["AFFO proxy"], ["Dividend / share"], ["AFFO payout ratio"]
];
operating.getRange("B5:E5").formulas = [["='Historical'!B12", "='Historical'!C12", "='Historical'!D12", "='Historical'!E12"]];
operating.getRange("B6:E6").formulas = [["", "=C5/B5-1", "=D5/C5-1", "=E5/D5-1"]];
operating.getRange("B7:E7").formulas = [["='Historical'!B18", "='Historical'!C18", "='Historical'!D18", "='Historical'!E18"]];
operating.getRange("B8:E8").formulas = [["='Historical'!B17", "='Historical'!C17", "='Historical'!D17", "='Historical'!E17"]];
operating.getRange("B9:E9").formulas = [["='Historical'!B11", "='Historical'!C11", "='Historical'!D11", "='Historical'!E11"]];
operating.getRange("B10:E10").formulas = [["='Historical'!B10", "='Historical'!C10", "='Historical'!D10", "='Historical'!E10"]];
operating.getRange("B11:E11").formulas = [["='Historical'!B16", "='Historical'!C16", "='Historical'!D16", "='Historical'!E16"]];
operating.getRange("B12:E12").formulas = [["='Source Data'!B37", "='Source Data'!C37", "='Source Data'!D37", "='Source Data'!E37"]];
operating.getRange("B13:E13").formulas = [["=B12/B8", "=C12/C8", "=D12/D8", "=E12/E8"]];
const pick = (bearCell, baseCell, bullCell) => `=IF('Assumptions'!$B$3=\"Bear\",'Assumptions'!${bearCell},IF('Assumptions'!$B$3=\"Bull\",'Assumptions'!${bullCell},'Assumptions'!${baseCell}))`;
operating.getRange("F5").formulas = [[pick("C6","D6","E6")]];
for (let i = 0; i < 5; i++) {
  const col = String.fromCharCode("G".charCodeAt(0) + i);
  const prev = String.fromCharCode(col.charCodeAt(0) - 1);
  const assumptionRow = 7 + i;
  operating.getRange(`${col}5`).formulas = [[`=${prev}5*(1+${pick(`C${assumptionRow}`,`D${assumptionRow}`,`E${assumptionRow}`).slice(1)})`]];
}
operating.getRange("F6:K6").formulas = [["=F5/E5-1", "=G5/F5-1", "=H5/G5-1", "=I5/H5-1", "=J5/I5-1", "=K5/J5-1"]];
for (const col of ["F","G","H","I","J","K"]) operating.getRange(`${col}7`).formulas = [[pick("C12","D12","E12")]];
operating.getRange("F8:K8").formulas = [["=F5*F7", "=G5*G7", "=H5*H7", "=I5*I7", "=J5*J7", "=K5*K7"]];
operating.getRange("F9").formulas = [["=AVERAGE('Source Data'!B9:B10)/F5"]];
for (const [col, prev] of [["G","F"],["H","G"],["I","H"],["J","I"],["K","J"]]) operating.getRange(`${col}9`).formulas = [[`=${prev}9*(1+${pick("C13","D13","E13").slice(1)})`]];
operating.getRange("F10:K10").formulas = [["=F5*F9", "=G5*G9", "=H5*H9", "=I5*I9", "=J5*J9", "=K5*K9"]];
operating.getRange("F11:K11").formulas = [["=F8*F9", "=G8*G9", "=H8*H9", "=I8*I9", "=J8*J9", "=K8*K9"]];
operating.getRange("F12").values = [[3.29]];
input(operating.getRange("F12"), fmt.perShare);
for (const [col, prev] of [["G","F"],["H","G"],["I","H"],["J","I"],["K","J"]]) operating.getRange(`${col}12`).formulas = [[`=MIN(${prev}12*(1+${pick("C14","D14","E14").slice(1)}),${col}8*${pick("C22","D22","E22").slice(1)})`]];
operating.getRange("F13:K13").formulas = [["=F12/F8", "=G12/G8", "=H12/H8", "=I12/I8", "=J12/J8", "=K12/K8"]];
operating.getRange("B5:K5").format.numberFormat = fmt.perShare;
operating.getRange("B6:K7").format.numberFormat = fmt.pct;
operating.getRange("B8:K8").format.numberFormat = fmt.perShare;
operating.getRange("B9:K11").format.numberFormat = fmt.money;
operating.getRange("B12:K12").format.numberFormat = fmt.perShare;
operating.getRange("B13:K13").format.numberFormat = fmt.pct;
operating.getRange("F4:K13").format.fill = COLORS.forecast;
crossSheet(operating.getRange("B5:E5"), fmt.perShare);
crossSheet(operating.getRange("B7:E12"));
total(operating.getRange("A8:K8"));
total(operating.getRange("A10:K11"));
total(operating.getRange("A13:K13"));

// ---------------- NAV ----------------
title(nav, "A1:H1", "Net Asset Value — Segment Cap-Rate Method");
nav.getRange("A2:H2").merge();
nav.getRange("A2").values = [["Spot NAV at the 2Q26 data date. The Valuation tab rolls this present value to the selected target date at the scenario cost of equity, net of projected dividends."]];
nav.getRange("A2:H2").format = { fill: COLORS.gray, font: { italic: true }, wrapText: true };
nav.getRange("A4:H4").values = [["Real estate NOI", "NOI ($mm)", "Bear cap", "Bear value", "Base cap", "Base value", "Bull cap", "Bull value"]];
header(nav.getRange("A4:H4"));
nav.getRange("A5:A9").values = [["Seniors Housing Operating"],["Seniors Housing Triple-net"],["Outpatient Medical"],["Long-Term / Post-Acute Care"],["Incremental stabilized NOI"]];
nav.getRange("B5:B9").formulas = [["='Source Data'!B41"],["='Source Data'!B42"],["='Source Data'!B43"],["='Source Data'!B44"],["='Source Data'!B45"]];
nav.getRange("C5:C9").formulas = [["='Assumptions'!C26"],["='Assumptions'!C27"],["='Assumptions'!C28"],["='Assumptions'!C29"],["='Assumptions'!C30"]];
nav.getRange("E5:E9").formulas = [["='Assumptions'!D26"],["='Assumptions'!D27"],["='Assumptions'!D28"],["='Assumptions'!D29"],["='Assumptions'!D30"]];
nav.getRange("G5:G9").formulas = [["='Assumptions'!E26"],["='Assumptions'!E27"],["='Assumptions'!E28"],["='Assumptions'!E29"],["='Assumptions'!E30"]];
for (let r = 5; r <= 9; r++) nav.getRange(`D${r}:H${r}`).formulas = [[`=B${r}/C${r}`, null, `=B${r}/E${r}`, null, `=B${r}/G${r}`]];
nav.getRange("A10:H10").values = [["Total real estate value", null, null, null, null, null, null, null]];
nav.getRange("D10").formulas = [["=SUM(D5:D9)"]];
nav.getRange("F10").formulas = [["=SUM(F5:F9)"]];
nav.getRange("H10").formulas = [["=SUM(H5:H9)"]];
total(nav.getRange("A10:H10"));

section(nav, "A12:H12", "Other assets");
nav.getRange("A13:H13").values = [["Other asset", "Carrying value", "Bear factor", "Bear value", "Base factor", "Base value", "Bull factor", "Bull value"]];
header(nav.getRange("A13:H13"));
nav.getRange("A14:A19").values = [["Land parcels"],["Real estate loans"],["Non-real-estate loans"],["JV real estate loans"],["Expected dispositions"],["Development balances"]];
nav.getRange("B14:B19").formulas = [["='Source Data'!B46"],["='Source Data'!B47"],["='Source Data'!B48"],["='Source Data'!B49"],["='Source Data'!B50"],["='Source Data'!B51"]];
nav.getRange("C14:C19").formulas = [["='Assumptions'!C34"],["='Assumptions'!C35"],["='Assumptions'!C36"],["='Assumptions'!C37"],["='Assumptions'!C38"],["='Assumptions'!C39"]];
nav.getRange("E14:E19").formulas = [["='Assumptions'!D34"],["='Assumptions'!D35"],["='Assumptions'!D36"],["='Assumptions'!D37"],["='Assumptions'!D38"],["='Assumptions'!D39"]];
nav.getRange("G14:G19").formulas = [["='Assumptions'!E34"],["='Assumptions'!E35"],["='Assumptions'!E36"],["='Assumptions'!E37"],["='Assumptions'!E38"],["='Assumptions'!E39"]];
for (let r = 14; r <= 19; r++) nav.getRange(`D${r}:H${r}`).formulas = [[`=B${r}*C${r}`, null, `=B${r}*E${r}`, null, `=B${r}*G${r}`]];
nav.getRange("A20:H20").values = [["Total other assets", null, null, null, null, null, null, null]];
nav.getRange("D20").formulas = [["=SUM(D14:D19)"]];
nav.getRange("F20").formulas = [["=SUM(F14:F19)"]];
nav.getRange("H20").formulas = [["=SUM(H14:H19)"]];
total(nav.getRange("A20:H20"));

section(nav, "A22:H22", "Equity value bridge");
nav.getRange("A23:A27").values = [["Gross asset value"],["Less: net obligations"],["Equity NAV"],["Fully diluted shares"],["NAV / share"]];
nav.getRange("D23").formulas = [["=SUM(D10,D20)"]];
nav.getRange("F23").formulas = [["=SUM(F10,F20)"]];
nav.getRange("H23").formulas = [["=SUM(H10,H20)"]];
nav.getRange("D24:H24").formulas = [["='Source Data'!B52", null, "='Source Data'!B52", null, "='Source Data'!B52"]];
nav.getRange("D25").formulas = [["=D23-D24"]];
nav.getRange("F25").formulas = [["=F23-F24"]];
nav.getRange("H25").formulas = [["=H23-H24"]];
nav.getRange("D26:H26").formulas = [["='Source Data'!B53", null, "='Source Data'!B53", null, "='Source Data'!B53"]];
nav.getRange("D27").formulas = [["=D25/D26"]];
nav.getRange("F27").formulas = [["=F25/F26"]];
nav.getRange("H27").formulas = [["=H25/H26"]];
total(nav.getRange("A25:H25"));
total(nav.getRange("A27:H27"));
nav.getRange("B5:B9").format.numberFormat = fmt.money;
nav.getRange("C5:C9").format.numberFormat = fmt.pct;
nav.getRange("E5:E9").format.numberFormat = fmt.pct;
nav.getRange("G5:G9").format.numberFormat = fmt.pct;
nav.getRange("D5:D25").format.numberFormat = fmt.money;
nav.getRange("F5:F25").format.numberFormat = fmt.money;
nav.getRange("H5:H25").format.numberFormat = fmt.money;
nav.getRange("D26").format.numberFormat = fmt.number;
nav.getRange("F26").format.numberFormat = fmt.number;
nav.getRange("H26").format.numberFormat = fmt.number;
nav.getRange("C14:C19").format.numberFormat = fmt.pct;
nav.getRange("E14:E19").format.numberFormat = fmt.pct;
nav.getRange("G14:G19").format.numberFormat = fmt.pct;
nav.getRange("D27:H27").format.numberFormat = fmt.perShare;

// ---------------- AFFO DCF ----------------
title(dcf, "A1:H1", "AFFO Proxy DCF — Equity Cash Flow Method");
dcf.getRange("A2:H2").merge();
dcf.getRange("A2").values = [["Present value as of the valuation date using discounted AFFO proxy per share and a Gordon-growth terminal value. The Valuation tab rolls this value to the target date net of projected dividends."]];
dcf.getRange("A2:H2").format = { fill: COLORS.gray, font: { italic: true }, wrapText: true };

function dcfBlock(start, scenarioName, scenarioCol) {
  section(dcf, `A${start}:H${start}`, `${scenarioName} case`);
  dcf.getRange(`A${start+1}:H${start+1}`).values = [["Metric", 2027, 2028, 2029, 2030, 2031, "", "Terminal / Output"]];
  header(dcf.getRange(`A${start+1}:H${start+1}`));
  dcf.getRange(`A${start+2}:A${start+9}`).values = [["Normalized FFO / share"],["AFFO conversion"],["AFFO proxy / share"],["Discount factor"],["PV of AFFO"],["Terminal growth"],["Cost of equity"],["DCF value / share"]];
  const growthRows = [7,8,9,10,11];
  const cols = ["B","C","D","E","F"];
  for (let i=0;i<5;i++) {
    const col = cols[i];
    if (i===0) dcf.getRange(`${col}${start+2}`).formulas = [[`='Assumptions'!${scenarioCol}6*(1+'Assumptions'!${scenarioCol}${growthRows[i]})`]];
    else dcf.getRange(`${col}${start+2}`).formulas = [[`=${cols[i-1]}${start+2}*(1+'Assumptions'!${scenarioCol}${growthRows[i]})`]];
    dcf.getRange(`${col}${start+3}`).formulas = [[`='Assumptions'!${scenarioCol}12`]];
    dcf.getRange(`${col}${start+4}`).formulas = [[`=${col}${start+2}*${col}${start+3}`]];
    dcf.getRange(`${col}${start+5}`).formulas = [[`=1/(1+'Assumptions'!${scenarioCol}18)^${i+1}`]];
    dcf.getRange(`${col}${start+6}`).formulas = [[`=${col}${start+4}*${col}${start+5}`]];
  }
  dcf.getRange(`H${start+7}`).formulas = [[`='Assumptions'!${scenarioCol}19`]];
  dcf.getRange(`H${start+8}`).formulas = [[`='Assumptions'!${scenarioCol}18`]];
  dcf.getRange(`H${start+9}`).formulas = [[`=SUM(B${start+6}:F${start+6})+(F${start+4}*(1+H${start+7})/(H${start+8}-H${start+7}))*F${start+5}`]];
  dcf.getRange(`B${start+2}:F${start+2}`).format.numberFormat = fmt.perShare;
  dcf.getRange(`B${start+3}:F${start+3}`).format.numberFormat = fmt.pct;
  dcf.getRange(`B${start+4}:F${start+4}`).format.numberFormat = fmt.perShare;
  dcf.getRange(`B${start+5}:F${start+5}`).format.numberFormat = "0.000x";
  dcf.getRange(`B${start+6}:F${start+6}`).format.numberFormat = fmt.perShare;
  dcf.getRange(`H${start+7}:H${start+8}`).format.numberFormat = fmt.pct;
  dcf.getRange(`H${start+9}`).format.numberFormat = fmt.perShare;
  total(dcf.getRange(`A${start+9}:H${start+9}`));
}

dcfBlock(4, "Bear", "C");
dcfBlock(16, "Base", "D");
dcfBlock(28, "Bull", "E");

section(dcf, "A40:G40", "Base-case DCF sensitivity — cost of equity vs terminal growth");
dcf.getRange("A41:G41").values = [["Cost of equity", "", 0.030, 0.035, 0.040, 0.045, 0.050]];
header(dcf.getRange("A41:G41"));
dcf.getRange("A42:A46").values = [[0.080],[0.085],[0.090],[0.095],[0.100]];
dcf.getRange("A42:A46").format.numberFormat = fmt.pct;
dcf.getRange("C41:G41").format.numberFormat = fmt.pct;
for (let r=42;r<=46;r++) {
  for (let c=3;c<=7;c++) {
    const col = String.fromCharCode(64+c);
    dcf.getRange(`${col}${r}`).formulas = [[`=SUM($B$20/(1+$A${r})^1,$C$20/(1+$A${r})^2,$D$20/(1+$A${r})^3,$E$20/(1+$A${r})^4,$F$20/(1+$A${r})^5)+($F$20*(1+${col}$41)/(($A${r}-${col}$41)*(1+$A${r})^5))`]];
  }
}
dcf.getRange("C42:G46").format.numberFormat = fmt.perShare;
dcf.getRange("C42:G46").conditionalFormats.add("colorScale", { colors: ["#F8696B", "#FFEB84", "#63BE7B"] });

// ---------------- Multiples ----------------
title(multiples, "A1:E1", "Forward Multiples & Yield Price Targets");
multiples.getRange("A2:E2").merge();
multiples.getRange("A2").formulas = [["=\"Target-date values use 2027E earnings and the selected \"&TEXT('Assumptions'!B52,\"yyyy-mm-dd\")&\" target date. Multiples and AFFO yields are explicit assumptions, not third-party consensus.\""]];
multiples.getRange("A2:E2").format = { fill: COLORS.gray, font: { italic: true }, wrapText: true };
multiples.getRange("A4:E4").values = [["Metric", "Units", "Bear", "Base", "Bull"]];
header(multiples.getRange("A4:E4"));
multiples.getRange("A5:A16").values = [["2026E normalized FFO / share"],["2027 growth"],["2027E normalized FFO / share"],["Target P / FFO"],["Target price / share — P / FFO"],[""],["AFFO conversion"],["2027E AFFO proxy / share"],["Target AFFO yield"],["Target price / share — AFFO yield"],[""],["Current market price"]];
multiples.getRange("C5:E5").formulas = [["='Assumptions'!C6","='Assumptions'!D6","='Assumptions'!E6"]];
multiples.getRange("C6:E6").formulas = [["='Assumptions'!C7","='Assumptions'!D7","='Assumptions'!E7"]];
multiples.getRange("C7:E7").formulas = [["=C5*(1+C6)","=D5*(1+D6)","=E5*(1+E6)"]];
multiples.getRange("C8:E8").formulas = [["='Assumptions'!C20","='Assumptions'!D20","='Assumptions'!E20"]];
multiples.getRange("C9:E9").formulas = [["=C7*C8","=D7*D8","=E7*E8"]];
multiples.getRange("C11:E11").formulas = [["='Assumptions'!C12","='Assumptions'!D12","='Assumptions'!E12"]];
multiples.getRange("C12:E12").formulas = [["=C7*C11","=D7*D11","=E7*E11"]];
multiples.getRange("C13:E13").formulas = [["='Assumptions'!C21","='Assumptions'!D21","='Assumptions'!E21"]];
multiples.getRange("C14:E14").formulas = [["=C12/C13","=D12/D13","=E12/E13"]];
multiples.getRange("C16:E16").formulas = [["='Source Data'!B6","='Source Data'!B6","='Source Data'!B6"]];
multiples.getRange("A18:E18").values = [["Current-price implied metrics", "Units", "Bear", "Base", "Bull"]];
header(multiples.getRange("A18:E18"));
multiples.getRange("A19:A22").values = [["Current P / 2026E FFO"],["Current P / 2027E FFO"],["2026E AFFO yield"],["2027E AFFO yield"]];
multiples.getRange("C19:E19").formulas = [["=C16/C5","=D16/D5","=E16/E5"]];
multiples.getRange("C20:E20").formulas = [["=C16/C7","=D16/D7","=E16/E7"]];
multiples.getRange("C21:E21").formulas = [["=C5*C11/C16","=D5*D11/D16","=E5*E11/E16"]];
multiples.getRange("C22:E22").formulas = [["=C12/C16","=D12/D16","=E12/E16"]];
multiples.getRange("C5:E7").format.numberFormat = fmt.perShare;
multiples.getRange("C6:E6").format.numberFormat = fmt.pct;
multiples.getRange("C8:E8").format.numberFormat = fmt.multiple;
multiples.getRange("C9:E9").format.numberFormat = fmt.perShare;
multiples.getRange("C11:E11").format.numberFormat = fmt.pct;
multiples.getRange("C12:E12").format.numberFormat = fmt.perShare;
multiples.getRange("C13:E13").format.numberFormat = fmt.pct;
multiples.getRange("C14:E16").format.numberFormat = fmt.perShare;
multiples.getRange("C19:E20").format.numberFormat = fmt.multiple;
multiples.getRange("C21:E22").format.numberFormat = fmt.pct;
total(multiples.getRange("A9:E9"));
total(multiples.getRange("A14:E14"));

section(multiples, "A24:G24", "Target-price sensitivity — 2027E FFO / share vs P / FFO");
multiples.getRange("A25:G25").values = [["2027E FFO / share", "", 30, 32, 34, 36, 38]];
header(multiples.getRange("A25:G25"));
multiples.getRange("A26:A30").values = [[6.60],[6.80],[7.00],[7.20],[7.40]];
multiples.getRange("A26:A30").format.numberFormat = fmt.perShare;
multiples.getRange("C25:G25").format.numberFormat = fmt.multiple;
for (let r=26;r<=30;r++) for (let c=3;c<=7;c++) {
  const col=String.fromCharCode(64+c);
  multiples.getRange(`${col}${r}`).formulas = [[`=$A${r}*${col}$25`]];
}
multiples.getRange("C26:G30").format.numberFormat = fmt.perShare;
multiples.getRange("C26:G30").conditionalFormats.add("colorScale", { colors: ["#F8696B", "#FFEB84", "#63BE7B"] });

// ---------------- Valuation ----------------
title(valuation, "A1:E1", "12–18 Month Price Target Synthesis");
valuation.getRange("A2:E2").merge();
valuation.getRange("A2").formulas = [["=\"Target date: \"&TEXT('Assumptions'!B52,\"yyyy-mm-dd\")&\". P / FFO and AFFO-yield methods use 2027E earnings; present-value DCF and NAV are rolled forward at the scenario cost of equity, net of projected dividends.\""]];
valuation.getRange("A2:E2").format = { fill: COLORS.gray, font: { italic: true }, wrapText: true };
valuation.getRange("A4:E4").values = [["Target-price method", "Weight", "Bear", "Base", "Bull"]];
header(valuation.getRange("A4:E4"));
valuation.getRange("A5:A8").values = [["2027E P / FFO"],["2027E AFFO yield"],["AFFO DCF rolled to target"],["NAV rolled to target"]];
valuation.getRange("C5:E5").formulas = [["='Multiples'!C9","='Multiples'!D9","='Multiples'!E9"]];
valuation.getRange("C6:E6").formulas = [["='Multiples'!C14","='Multiples'!D14","='Multiples'!E14"]];
valuation.getRange("C7:E7").formulas = [["=C30","=D30","=E30"]];
valuation.getRange("C8:E8").formulas = [["=C31","=D31","=E31"]];
valuation.getRange("B5:B8").formulas = [["='Assumptions'!D43"],["='Assumptions'!D44"],["='Assumptions'!D45"],["='Assumptions'!D46"]];
valuation.getRange("A10:E10").values = [["Composite price target / share", null, null, null, null]];
valuation.getRange("C10:E10").formulas = [["=SUMPRODUCT(C5:C8,'Assumptions'!C43:C46)","=SUMPRODUCT(D5:D8,'Assumptions'!D43:D46)","=SUMPRODUCT(E5:E8,'Assumptions'!E43:E46)"]];
valuation.getRange("A11:E11").values = [["Current market price", null, null, null, null]];
valuation.getRange("C11:E11").formulas = [["='Source Data'!B6","='Source Data'!B6","='Source Data'!B6"]];
valuation.getRange("A12:E12").values = [["Price appreciation / (downside)", null, null, null, null]];
valuation.getRange("C12:E12").formulas = [["=C10/C11-1","=D10/D11-1","=E10/E11-1"]];
valuation.getRange("A13:E13").values = [["Forecast dividends through target", null, null, null, null]];
valuation.getRange("C13:E13").formulas = [["=C29","=D29","=E29"]];
valuation.getRange("A14:E14").values = [["Total shareholder return", null, null, null, null]];
valuation.getRange("C14:E14").formulas = [["=(C10+C13)/C11-1","=(D10+D13)/D11-1","=(E10+E13)/E11-1"]];
valuation.getRange("A16:E16").values = [["Scenario price-target range", null, null, null, null]];
valuation.getRange("C16:E16").formulas = [["=C10","=D10","=E10"]];

section(valuation, "A18:E18", "Target-date roll-forward audit");
valuation.getRange("A19:A31").values = [
  ["DCF present value / share"],["NAV present value / share"],["Cost of equity"],["Target horizon (years)"],
  ["2026 annual dividend / share"],["2027 annual dividend / share"],["2028 annual dividend / share"],
  ["2026 horizon fraction"],["2027 horizon fraction"],["2028 horizon fraction"],
  ["Forecast dividends through target"],["DCF target-date value / share"],["NAV target-date value / share"]
];
valuation.getRange("C19:E19").formulas = [["='AFFO DCF'!H13","='AFFO DCF'!H25","='AFFO DCF'!H37"]];
valuation.getRange("C20:E20").formulas = [["='NAV'!D27","='NAV'!F27","='NAV'!H27"]];
valuation.getRange("C21:E21").formulas = [["='Assumptions'!C18","='Assumptions'!D18","='Assumptions'!E18"]];
valuation.getRange("C22:E22").formulas = [["='Assumptions'!$B$53","='Assumptions'!$B$53","='Assumptions'!$B$53"]];
valuation.getRange("C23:E23").formulas = [["='Operating Model'!$F$12","='Operating Model'!$F$12","='Operating Model'!$F$12"]];
for (const col of ["C","D","E"]) {
  valuation.getRange(`${col}24`).formulas = [[`=MIN(${col}23*(1+'Assumptions'!${col}14),('Assumptions'!${col}6*(1+'Assumptions'!${col}7))*'Assumptions'!${col}12*'Assumptions'!${col}22)`]];
  valuation.getRange(`${col}25`).formulas = [[`=MIN(${col}24*(1+'Assumptions'!${col}14),('Assumptions'!${col}6*(1+'Assumptions'!${col}7)*(1+'Assumptions'!${col}8))*'Assumptions'!${col}12*'Assumptions'!${col}22)`]];
}
valuation.getRange("C26:E26").formulas = [["=MAX(0,(MIN('Assumptions'!$B$52,DATE(2026,12,31))-MAX('Assumptions'!$B$51,DATE(2026,1,1))+1)/365)","=C26","=C26"]];
valuation.getRange("C27:E27").formulas = [["=MAX(0,(MIN('Assumptions'!$B$52,DATE(2027,12,31))-MAX('Assumptions'!$B$51,DATE(2027,1,1))+1)/365)","=C27","=C27"]];
valuation.getRange("C28:E28").formulas = [["=MAX(0,(MIN('Assumptions'!$B$52,DATE(2028,12,31))-MAX('Assumptions'!$B$51,DATE(2028,1,1))+1)/366)","=C28","=C28"]];
valuation.getRange("C29:E29").formulas = [["=SUMPRODUCT(C23:C25,C26:C28)","=SUMPRODUCT(D23:D25,D26:D28)","=SUMPRODUCT(E23:E25,E26:E28)"]];
valuation.getRange("C30:E30").formulas = [["=C19*(1+C21)^C22-C29","=D19*(1+D21)^D22-D29","=E19*(1+E21)^E22-E29"]];
valuation.getRange("C31:E31").formulas = [["=C20*(1+C21)^C22-C29","=D20*(1+D21)^D22-D29","=E20*(1+E21)^E22-E29"]];

valuation.getRange("B5:B8").format.numberFormat = fmt.pct;
valuation.getRange("C5:E11").format.numberFormat = fmt.perShare;
valuation.getRange("C12:E12").format.numberFormat = fmt.pct;
valuation.getRange("C13:E13").format.numberFormat = fmt.perShare;
valuation.getRange("C14:E14").format.numberFormat = fmt.pct;
valuation.getRange("C16:E16").format.numberFormat = fmt.perShare;
valuation.getRange("C19:E20").format.numberFormat = fmt.perShare;
valuation.getRange("C21:E21").format.numberFormat = fmt.pct;
valuation.getRange("C22:E22").format.numberFormat = '0.00 "yrs"';
valuation.getRange("C23:E25").format.numberFormat = fmt.perShare;
valuation.getRange("C26:E28").format.numberFormat = fmt.pct;
valuation.getRange("C29:E31").format.numberFormat = fmt.perShare;
total(valuation.getRange("A10:E10"));
total(valuation.getRange("A14:E14"));
total(valuation.getRange("A16:E16"));
total(valuation.getRange("A29:E31"));

// ---------------- Checks ----------------
title(checks, "A1:G1", "Model Checks");
checks.getRange("A3:G3").values = [["Check", "Actual", "Expected", "Difference", "Tolerance", "Status", "Notes / fix"]];
header(checks.getRange("A3:G3"));
checks.getRange("A4:A15").values = [
  ["2026 FFO midpoint"],["2026 implied shares"],["2026 AFFO bridge"],["NAV equity bridge — Base"],["DCF PV bridge — Base"],["Valuation scenario order"],["Assumption weight total — Base"],["Cap rates positive"],["Cost of equity > terminal growth"],["Operating AFFO tie"],["Target horizon 12–18 months"],["Base DCF target roll-forward"]
];
checks.getRange("B4:C15").formulas = [
  ["='Assumptions'!D6", "=AVERAGE('Source Data'!B7:B8)"],
  ["='Operating Model'!F9", "=AVERAGE('Source Data'!B9:B10)/'Assumptions'!D6"],
  ["='Historical'!F16", "=SUM('Historical'!F10,'Historical'!F13:F15)"],
  ["='NAV'!F25", "='NAV'!F23-'NAV'!F24"],
  ["='AFFO DCF'!H25", "=SUM('AFFO DCF'!B22:F22)+('AFFO DCF'!F20*(1+'AFFO DCF'!H23)/('AFFO DCF'!H24-'AFFO DCF'!H23))*'AFFO DCF'!F21"],
  ["=IF(AND('Valuation'!C10<='Valuation'!D10,'Valuation'!D10<='Valuation'!E10),1,0)", "=1"],
  ["='Assumptions'!D47", "=1"],
  ["=MIN('Assumptions'!C26:E30)", "=0"],
  ["=MIN('Assumptions'!C18-'Assumptions'!C19,'Assumptions'!D18-'Assumptions'!D19,'Assumptions'!E18-'Assumptions'!E19)", "=0"],
  ["='Operating Model'!F11", "='Historical'!F16"],
  ["=IF(AND('Assumptions'!B54>=12,'Assumptions'!B54<=18),1,0)", "=1"],
  ["='Valuation'!D30", "='Valuation'!D19*(1+'Valuation'!D21)^'Valuation'!D22-'Valuation'!D29"],
];
for (let r=4;r<=15;r++) checks.getRange(`D${r}`).formulas = [[`=B${r}-C${r}`]];
checks.getRange("E4:E15").values = [[0.0001],[0.0001],[0.0001],[0.0001],[0.0001],[0],[0.0001],[0],[0],[0.0001],[0],[0.0001]];
checks.getRange("F4:F8").formulas = [["=IF(ABS(D4)<=E4,\"OK\",\"FAIL\")"],["=IF(ABS(D5)<=E5,\"OK\",\"FAIL\")"],["=IF(ABS(D6)<=E6,\"OK\",\"FAIL\")"],["=IF(ABS(D7)<=E7,\"OK\",\"FAIL\")"],["=IF(ABS(D8)<=E8,\"OK\",\"FAIL\")"]];
checks.getRange("F9:F12").formulas = [["=IF(B9=C9,\"OK\",\"FAIL\")"],["=IF(ABS(D10)<=E10,\"OK\",\"FAIL\")"],["=IF(B11>C11,\"OK\",\"FAIL\")"],["=IF(B12>C12,\"OK\",\"FAIL\")"]];
checks.getRange("F13").formulas = [["=IF(ABS(D13)<=E13,\"OK\",\"FAIL\")"]];
checks.getRange("F14:F15").formulas = [["=IF(B14=C14,\"OK\",\"FAIL\")"],["=IF(ABS(D15)<=E15,\"OK\",\"FAIL\")"]];
checks.getRange("G4:G15").values = [
  ["Guidance midpoint should equal average of low/high"],["Shares derive from guidance dollars divided by per-share guidance"],["AFFO proxy must tie to disclosed adjustments"],["Gross assets less obligations"],["PV of annual AFFO plus discounted terminal value"],["Bear <= Base <= Bull"],["Weights must sum to 100%"],["All cap rates must exceed zero"],["Prevents invalid Gordon-growth terminal value"],["2026E operating AFFO must tie to the disclosed bridge"],["Selected target date must be 12 to 18 months after valuation date"],["Present DCF value compounded to target date less projected dividends"]
];
checks.getRange("A17:F17").values = [["Overall model status", null, null, null, null, null]];
checks.getRange("F17").formulas = [["=IF(COUNTIF(F4:F15,\"FAIL\")=0,\"OK\",\"REVIEW\")"]];
total(checks.getRange("A17:G17"));
checks.getRange("B4:E15").format.numberFormat = "0.0000";
checks.getRange("F4:F17").conditionalFormats.add("containsText", { text: "OK", format: { fill: COLORS.greenFill, font: { color: "#006100", bold: true } } });
checks.getRange("F4:F17").conditionalFormats.add("containsText", { text: "FAIL", format: { fill: COLORS.redFill, font: { color: COLORS.red, bold: true } } });
checks.getRange("F17").conditionalFormats.add("containsText", { text: "REVIEW", format: { fill: COLORS.amberFill, font: { color: "#9C6500", bold: true } } });

// ---------------- Summary ----------------
title(summary, "A1:H1", "Welltower (NYSE: WELL) — 12–18 Month Price Target Dashboard");
summary.getRange("A2:H2").merge();
summary.getRange("A2").formulas = [["=\"Market close: \"&TEXT('Assumptions'!B51,\"yyyy-mm-dd\")&\" | Target date: \"&TEXT('Assumptions'!B52,\"yyyy-mm-dd\")&\" | \"&TEXT('Assumptions'!B54,\"0.0\")&\"-month horizon | Company data through 2Q26\""]];
summary.getRange("A2:H2").format = { fill: COLORS.gray, font: { italic: true, color: "#4B5563" } };

section(summary, "A4:H4", "Price-target snapshot");
summary.getRange("A5:H5").values = [["Market price", null, "Target date", null, "Base price target", null, "Base total return", null]];
summary.getRange("A6:H6").formulas = [[null,"='Source Data'!B6",null,"='Assumptions'!B52",null,"='Valuation'!D10",null,"='Valuation'!D14"]];
summary.getRange("A5:H5").format = { fill: COLORS.tealLight, font: { bold: true, color: COLORS.navy }, horizontalAlignment: "center" };
summary.getRange("B6").format.numberFormat = fmt.perShare;
summary.getRange("D6").format.numberFormat = "yyyy-mm-dd";
summary.getRange("F6").format.numberFormat = fmt.perShare;
summary.getRange("H6").format.numberFormat = fmt.pct;
summary.getRange("B6:H6").format = { fill: COLORS.white, font: { bold: true, size: 14, color: COLORS.green }, horizontalAlignment: "center", borders: { preset: "outside", style: "thin", color: COLORS.gray2 } };

section(summary, "A8:H8", "12–18 month price-target range");
summary.getRange("A9:D9").values = [["", "Bear", "Base", "Bull"]];
header(summary.getRange("A9:D9"));
summary.getRange("A10:A14").values = [["Composite price target / share"],["Price appreciation / (downside)"],["Forecast dividends"],["Total shareholder return"],["Current price"]];
summary.getRange("B10:D10").formulas = [["='Valuation'!C10","='Valuation'!D10","='Valuation'!E10"]];
summary.getRange("B11:D11").formulas = [["='Valuation'!C12","='Valuation'!D12","='Valuation'!E12"]];
summary.getRange("B12:D12").formulas = [["='Valuation'!C13","='Valuation'!D13","='Valuation'!E13"]];
summary.getRange("B13:D13").formulas = [["='Valuation'!C14","='Valuation'!D14","='Valuation'!E14"]];
summary.getRange("B14:D14").formulas = [["='Source Data'!B6","='Source Data'!B6","='Source Data'!B6"]];
summary.getRange("B10:D10").format.numberFormat = fmt.perShare;
summary.getRange("B11:D11").format.numberFormat = fmt.pct;
summary.getRange("B12:D12").format.numberFormat = fmt.perShare;
summary.getRange("B13:D13").format.numberFormat = fmt.pct;
summary.getRange("B14:D14").format.numberFormat = fmt.perShare;
total(summary.getRange("A10:D10"));
total(summary.getRange("A13:D14"));

summary.getRange("F9:H9").values = [["Implied market metrics", "Value", "Interpretation"]];
header(summary.getRange("F9:H9"));
summary.getRange("F10:F14").values = [["2026E P / FFO"],["2027E P / FFO"],["2026E AFFO yield"],["Net debt / EBITDA"],["Uniform implied cap rate"]];
summary.getRange("G10").formulas = [["='Multiples'!D19"]];
summary.getRange("G11").formulas = [["='Multiples'!D20"]];
summary.getRange("G12").formulas = [["='Multiples'!D21"]];
summary.getRange("G13").formulas = [["='Source Data'!B18"]];
summary.getRange("G14").formulas = [["=SUM('Source Data'!B41:B45)/(('Source Data'!B6*'Source Data'!B53)+'Source Data'!B52-SUM('NAV'!F20))"]];
summary.getRange("H10:H14").values = [["Premium multiple"],["Still premium after base growth"],["Low cash-flow yield"],["Low leverage"],["Cap rate implied by market equity value"]];
summary.getRange("G10:G11").format.numberFormat = fmt.multiple;
summary.getRange("G12").format.numberFormat = fmt.pct;
summary.getRange("G13").format.numberFormat = fmt.multiple;
summary.getRange("G14").format.numberFormat = fmt.pct;
summary.getRange("H10:H14").format.wrapText = true;

section(summary, "A16:H16", "Investment read-through");
summary.getRange("A17:A20").values = [["Base-case conclusion"],["Core upside drivers"],["Core downside risks"],["How to use the model"]];
for (let r=17;r<=20;r++) summary.getRange(`B${r}:H${r}`).merge();
summary.getRange("B17").formulas = [["=\"The base price target is \"&TEXT('Valuation'!D10,\"$0.00\")&\", implying \"&TEXT('Valuation'!D12,\"0.0%\")&\" price return and \"&TEXT('Valuation'!D14,\"0.0%\")&\" total return through \"&TEXT('Assumptions'!B52,\"yyyy-mm-dd\")&\".\""]];
summary.getRange("B18").values = [["Continued SHO occupancy gains, RevPOR ahead of ExpPOR, Welltower Business System margin expansion, accretive acquisitions, and low leverage."]];
summary.getRange("B19").values = [["Cap-rate normalization, higher required returns, slowing SHO same-store NOI, acquisition integration, dilution, tenant/operator stress, and weaker capital markets."]];
summary.getRange("B20").values = [["Change the target date, scenario assumptions, FFO growth, AFFO conversion, target multiples/yields, cap rates, and weights on the Assumptions tab. Review Checks before relying on outputs."]];
for (let r=17;r<=20;r++) {
  summary.getRange(`A${r}`).format = { fill: COLORS.tealLight, font: { bold: true, color: COLORS.navy }, verticalAlignment: "top" };
  summary.getRange(`B${r}:H${r}`).format = { wrapText: true, verticalAlignment: "top" };
  summary.getRange(`A${r}:H${r}`).format.rowHeight = 34;
}

summary.getRange("A23:C23").values = [["Year", "Normalized FFO / share", "AFFO proxy / share"]];
header(summary.getRange("A23:C23"));
for (let i=0;i<10;i++) {
  const r=24+i;
  const col=String.fromCharCode("B".charCodeAt(0)+i);
  summary.getRange(`A${r}:C${r}`).formulas = [[`=TEXT('Operating Model'!${col}4,\"yyyy\")`,`='Operating Model'!${col}5`,`='Operating Model'!${col}8`]];
}
summary.getRange("A24:A33").format.numberFormat = "@";
summary.getRange("B24:C33").format.numberFormat = fmt.perShare;
const trend = summary.charts.add("line", summary.getRange("A23:C33"));
trend.title = "Per-share cash earnings trajectory";
trend.hasLegend = true;
trend.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
trend.yAxis = { numberFormatCode: "$0.00" };
trend.setPosition("J4", "Q18");

summary.getRange("E23:F23").values = [["Base target-price method", "Target / share"]];
header(summary.getRange("E23:F23"));
summary.getRange("E24:E27").values = [["2027E P / FFO"],["2027E AFFO yield"],["AFFO DCF rolled forward"],["NAV rolled forward"]];
summary.getRange("F24:F27").formulas = [["='Valuation'!D5"],["='Valuation'!D6"],["='Valuation'!D7"],["='Valuation'!D8"]];
summary.getRange("F24:F27").format.numberFormat = fmt.perShare;
const bars = summary.charts.add("bar", summary.getRange("E23:F27"));
bars.title = "Base-case target methods ($ / share)";
bars.hasLegend = false;
bars.yAxis = { numberFormatCode: "$0" };
bars.setPosition("J20", "Q33");

summary.getRange("A35:H35").merge();
summary.getRange("A35").formulas = [["=\"Model status: \"&'Checks'!F17"]];
summary.getRange("A35:H35").format = { fill: COLORS.navy, font: { bold: true, color: COLORS.white }, horizontalAlignment: "center" };

// ---------------- Global formatting ----------------
const widths = {
  "Summary": [230,105,105,105,155,200,95,250,25,95,95,95,95,95,95,95,95],
  "Source Data": [320,105,400,150,80,340],
  "Assumptions": [255,185,95,95,95],
  "Historical": [270,105,105,105,105,105],
  "Operating Model": [245,92,92,92,92,92,92,92,92,92,92],
  "NAV": [245,105,90,120,90,120,90,120],
  "AFFO DCF": [210,95,95,95,95,95,95,140],
  "Multiples": [260,115,100,100,100,100,100],
  "Valuation": [220,90,110,110,110],
  "Checks": [230,110,110,110,90,90,310],
};
for (const sheet of wb.worksheets.items) {
  const ws = widths[sheet.name];
  if (ws) ws.forEach((w,i)=>sheet.getRangeByIndexes(0,i,1,1).format.columnWidthPx=w);
  sheet.freezePanes.freezeRows(sheet.name === "Summary" ? 2 : 4);
  sheet.getUsedRange().format.font = { name: "Aptos", size: 10 };
  sheet.getUsedRange().format.verticalAlignment = "center";
}
sources.getRange("F6:F64").format.wrapText = true;
sources.getRange("C57:C64").format.wrapText = true;
sources.getRange("A57:F64").format.rowHeight = 42;
assumptions.getRange("A2:E2").format.rowHeight = 34;
historical.getRange("A2:F2").format.rowHeight = 42;
nav.getRange("A2:H2").format.rowHeight = 36;
dcf.getRange("A2:H2").format.rowHeight = 42;
multiples.getRange("A2:E2").format.rowHeight = 34;
valuation.getRange("A2:E2").format.rowHeight = 48;
checks.getRange("G4:G15").format.wrapText = true;
checks.getRange("A4:G15").format.rowHeight = 28;
operating.getRange("A2:K2").format = { fill: COLORS.gray, font: { italic: true, color: COLORS.green }, horizontalAlignment: "left" };

// Final checks and compact inspection before export.
const inspectSummary = await wb.inspect({ kind: "table", range: "Summary!A1:H20", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 10, maxChars: 14000 });
console.log(inspectSummary.ndjson);
const inspectValuation = await wb.inspect({ kind: "table", range: "Valuation!A1:E31", include: "values,formulas", tableMaxRows: 35, tableMaxCols: 8, maxChars: 18000 });
console.log(inspectValuation.ndjson);
const errorScan = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!", options: { useRegex: true, maxResults: 300 }, summary: "final formula error scan", maxChars: 10000 });
console.log(errorScan.ndjson);
const checkInspect = await wb.inspect({ kind: "table", range: "Checks!A3:G17", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 8, maxChars: 14000 });
console.log(checkInspect.ndjson);

for (const sheet of wb.worksheets.items) {
  const preview = await wb.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
  const safe = sheet.name.toLowerCase().replace(/[^a-z0-9]+/g,"-");
  await fs.writeFile(path.join(outDir, `preview-${safe}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const exported = await SpreadsheetFile.exportXlsx(wb);
await exported.save(path.join(outDir, "WELL_Valuation_Model_2026-09-03.xlsx"));
console.log(path.join(outDir, "WELL_Valuation_Model_2026-09-03.xlsx"));
