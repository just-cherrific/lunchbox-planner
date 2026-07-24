import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("C:/trial3/HKO_Rain_Forecast_Log.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);
const check = await workbook.inspect({
  kind: "table",
  range: "Rain Forecast Log!A1:B3",
  include: "values,formulas",
  tableMaxRows: 5,
  tableMaxCols: 2,
});
console.log(check.ndjson);
const preview = await workbook.render({ sheetName: "Rain Forecast Log", range: "A1:B3", scale: 2, format: "png" });
await fs.writeFile("C:/trial3/HKO_Rain_Forecast_Log_preview.png", new Uint8Array(await preview.arrayBuffer()));
