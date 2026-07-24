import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "C:/trial3";
const weatherPath = `${root}/HKO_Rain_Forecast_Log.xlsx`;
const calendarPath = `${root}/school_lunch_2026-07-25_to_2026-08-23.xlsx`;
const outputPath = `${root}/Lunchbox_Quantity_Forecast.xlsx`;
const outputSheetName = "Lunchbox Forecast";

const dateHeaders = ["Forecast_Target_Date", "Date"];
const studentHeaders = ["Student_Count", "Number of students eating school lunch"];

function hongKongTargetDate() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Hong_Kong", year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.filter((p) => p.type !== "literal").map((p) => [p.type, p.value]));
  return new Date(Date.UTC(+values.year, +values.month - 1, +values.day + 7)).toISOString().slice(0, 10);
}

function normalizeDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && Number.isFinite(value)) {
    // Excel's 1900-date system, preserving the calendar day even when a time fraction is present.
    return new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000).toISOString().slice(0, 10);
  }
  if (typeof value !== "string") return null;
  const iso = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
  if (!iso) return null;
  const candidate = `${iso[1]}-${iso[2]}-${iso[3]}`;
  const date = new Date(`${candidate}T00:00:00Z`);
  return Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== candidate ? null : candidate;
}

function columnIndex(headers, accepted) {
  return headers.findIndex((header) => accepted.includes(String(header ?? "").trim()));
}

async function importWorkbook(path) {
  return SpreadsheetFile.importXlsx(await FileBlob.load(path));
}

function extractRows(workbook, acceptedValueHeaders, label) {
  const sheet = workbook.worksheets.getItemAt(0);
  const used = sheet.getUsedRange(true);
  const rows = used?.values ?? [];
  if (!rows.length) throw new Error(`${label} workbook is empty.`);
  const dateIndex = columnIndex(rows[0], dateHeaders);
  const valueIndex = columnIndex(rows[0], acceptedValueHeaders);
  if (dateIndex < 0 || valueIndex < 0) throw new Error(`${label} workbook does not contain its required date and value columns.`);
  return rows.slice(1).map((row) => ({ date: normalizeDate(row[dateIndex]), value: row[valueIndex] }));
}

async function loadOutputWorkbook() {
  try {
    const workbook = await importWorkbook(outputPath);
    const sheet = workbook.worksheets.getItem(outputSheetName);
    const used = sheet.getUsedRange(true);
    const rows = used?.values ?? [];
    if (rows.length === 0 || rows[0].length !== 2 || rows[0][0] !== "Forecast_Target_Date" || rows[0][1] !== "Recommended_Lunchboxes") {
      throw new Error("Output workbook does not have exactly the required two columns.");
    }
    return { workbook, sheet, rows };
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add(outputSheetName);
    sheet.getRange("A1:B1").values = [["Forecast_Target_Date", "Recommended_Lunchboxes"]];
    sheet.getRange("A1:B1").format = { fill: "#1F4E78", font: { bold: true, color: "#FFFFFF" } };
    sheet.getRange("A1:B1").format.borders = { preset: "outside", style: "thin", color: "#1F1F1F" };
    sheet.getRange("A:A").format.columnWidth = 24;
    sheet.getRange("B:B").format.columnWidth = 28;
    sheet.freezePanes.freezeRows(1);
    sheet.showGridLines = false;
    return { workbook, sheet, rows: [["Forecast_Target_Date", "Recommended_Lunchboxes"]] };
  }
}

async function main() {
  const targetDate = hongKongTargetDate();
  const [weatherWorkbook, calendarWorkbook, output] = await Promise.all([
    importWorkbook(weatherPath), importWorkbook(calendarPath), loadOutputWorkbook(),
  ]);
  const weatherRows = extractRows(weatherWorkbook, ["Rain_Status"], "Weather log");
  const calendarRows = extractRows(calendarWorkbook, studentHeaders, "Calendar");
  const calendarMatches = calendarRows.filter((row) => row.date === targetDate);
  if (calendarMatches.length !== 1 || !Number.isInteger(calendarMatches[0]?.value) || calendarMatches[0].value < 0) {
    throw new Error(`Calendar issue for ${targetDate}: expected exactly one non-negative integer Student_Count; no output record was changed.`);
  }
  const studentCount = calendarMatches[0].value;
  const weatherMatches = weatherRows.filter((row) => row.date === targetDate);
  const rainStatus = weatherMatches.length === 1 ? String(weatherMatches[0].value ?? "").trim() : "[ERROR]";
  const isEstimated = !["[R]", "[NR]"].includes(rainStatus);
  const recommended = studentCount === 0 ? 0 : Math.ceil(studentCount * (rainStatus === "[R]" ? 0.75 : 0.65));

  const outputMatches = output.rows.slice(1)
    .map((row, index) => ({ rowNumber: index + 2, date: normalizeDate(row[0]) }))
    .filter((row) => row.date === targetDate);
  if (outputMatches.length > 1) throw new Error(`Output integrity issue: multiple records already exist for ${targetDate}; no records were changed.`);
  if (outputMatches.length === 1) {
    output.sheet.getRange(`A${outputMatches[0].rowNumber}:B${outputMatches[0].rowNumber}`).values = [[targetDate, recommended]];
  } else {
    const rowNumber = output.rows.length + 1;
    output.sheet.getRange(`A${rowNumber}:B${rowNumber}`).values = [[targetDate, recommended]];
  }
  const dataEnd = outputMatches.length === 1 ? output.rows.length : output.rows.length + 1;
  output.sheet.getRange(`A2:A${dataEnd}`).format.numberFormat = "@";
  output.sheet.getRange(`B2:B${dataEnd}`).format.numberFormat = "0";
  const xlsx = await SpreadsheetFile.exportXlsx(output.workbook);
  await xlsx.save(outputPath);
  console.log(JSON.stringify({ targetDate, studentCount, rainStatus, recommendedLunchboxes: recommended, estimated: isEstimated, action: outputMatches.length ? "updated" : "appended" }));
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
