import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workbookPath = "C:/trial3/HKO_Rain_Forecast_Log.xlsx";
const sheetName = "Rain Forecast Log";
const apiUrl = "https://data.weather.gov.hk/weatherAPI/opendata/weather.php?dataType=fnd&lang=en";
const rainTerms = [
  "drizzle", "light rain", "moderate rain", "heavy rain", "torrential rain",
  "showers", "squally showers", "thunder showers", "thunderstorm", "thunder",
  "rain patches", "rain bands", "rainfall", "rainy", "occasional rain",
  "scattered showers", "isolated showers",
];

function hongKongDatePlusSeven() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Hong_Kong",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const value = Object.fromEntries(parts.filter(({ type }) => type !== "literal").map(({ type, value }) => [type, value]));
  const date = new Date(Date.UTC(Number(value.year), Number(value.month) - 1, Number(value.day) + 7));
  return date.toISOString().slice(0, 10);
}

async function loadWorkbook() {
  try {
    await fs.access(workbookPath);
    const input = await FileBlob.load(workbookPath);
    return await SpreadsheetFile.importXlsx(input);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add(sheetName);
    sheet.getRange("A1:B1").values = [["Forecast_Target_Date", "Rain_Status"]];
    sheet.getRange("A1:B1").format = {
      fill: "#1F4E78",
      font: { bold: true, color: "#FFFFFF" },
    };
    sheet.getRange("A:B").format.columnWidth = 24;
    sheet.freezePanes.freezeRows(1);
    sheet.showGridLines = false;
    return workbook;
  }
}

async function main() {
  const targetDate = hongKongDatePlusSeven();
  const workbook = await loadWorkbook();
  const sheet = workbook.worksheets.getItem(sheetName);
  const used = sheet.getUsedRange(true);
  const rows = used ? used.values : [];
  if (rows.length === 0 || rows[0][0] !== "Forecast_Target_Date" || rows[0][1] !== "Rain_Status") {
    throw new Error("The workbook does not have the required Rain Forecast Log headers.");
  }
  const existingIndex = rows.slice(1).findIndex((row) => String(row[0] ?? "") === targetDate);
  if (existingIndex >= 0 && process.argv[2] !== "--correct-current-run") {
    console.log(JSON.stringify({ targetDate, action: "existing record prevented duplication" }));
    return;
  }

  if (existingIndex >= 0) {
    const status = process.argv[3];
    const forecastText = process.argv[4];
    if (rows[existingIndex + 1][1] !== "[ERROR]" || !["[R]", "[NR]"].includes(status) || typeof forecastText !== "string") {
      throw new Error("Current-run correction is only permitted for a just-created [ERROR] record.");
    }
    sheet.getRange(`B${existingIndex + 2}`).values = [[status]];
    const xlsx = await SpreadsheetFile.exportXlsx(workbook);
    await xlsx.save(workbookPath);
    console.log(JSON.stringify({ targetDate, status, forecastText, action: "current-run error corrected" }));
    return;
  }

  let status = "[ERROR]";
  let forecastText = "";
  let reason = "";
  if (process.argv[2] === "--append-classified") {
    status = process.argv[3];
    forecastText = process.argv[4];
    if (!["[R]", "[NR]", "[ERROR]"].includes(status) || typeof forecastText !== "string") {
      throw new Error("Classified append requires a valid status and forecast text or error reason.");
    }
  } else {
    try {
      const response = await fetch(apiUrl);
      if (!response.ok) throw new Error(`HKO API returned HTTP ${response.status}`);
      const data = await response.json();
      if (!Array.isArray(data.weatherForecast)) throw new Error("HKO API response did not contain weatherForecast.");
      const target = data.weatherForecast.find((entry) => entry.forecastDate === targetDate.replaceAll("-", ""));
      if (!target || typeof target.forecastWeather !== "string") throw new Error(`No forecast entry exists for ${targetDate}.`);
      forecastText = target.forecastWeather;
      const lower = forecastText.toLowerCase();
      status = rainTerms.some((term) => lower.includes(term)) ? "[R]" : "[NR]";
    } catch (error) {
      reason = error instanceof Error ? error.message : String(error);
    }
  }

  const appendRow = rows.length + 1;
  sheet.getRange(`A${appendRow}:B${appendRow}`).values = [[targetDate, status]];
  sheet.getRange(`A${appendRow}`).format.numberFormat = "@";
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(workbookPath);
  console.log(JSON.stringify({
    targetDate,
    status,
    forecastText: forecastText || null,
    errorReason: reason || null,
    action: "new Excel row appended",
  }));
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
