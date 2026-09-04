import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "../outputs/01a06864-0723-7a51-a935-b11fb1d3179e/WELL_Valuation_Model_2026-09-03.xlsx";
const input = await FileBlob.load(workbookPath);
const wb = await SpreadsheetFile.importXlsx(input);

for (const [range, rows, cols] of [
  ["Assumptions!A5:E54", 60, 8],
  ["NAV!A4:H27", 30, 10],
  ["AFFO DCF!A40:G46", 10, 10],
  ["Operating Model!A4:K13", 15, 12],
  ["Valuation!A4:E31", 35, 8],
  ["Checks!A3:G17", 20, 8],
]) {
  const result = await wb.inspect({ kind: "table", range, include: "values,formulas", tableMaxRows: rows, tableMaxCols: cols, maxChars: 14000 });
  console.log(result.ndjson);
}

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!",
  options: { useRegex: true, maxResults: 300 },
  summary: "post-export formula error scan",
  maxChars: 10000,
});
console.log(errors.ndjson);

const trace = await wb.trace("Valuation!D10");
console.log(String(trace.ndjson ?? trace).slice(0, 10000));
