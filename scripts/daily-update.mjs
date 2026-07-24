import { writeFile } from "node:fs/promises";

await writeFile(
  new URL("../public/daily-status.json", import.meta.url),
  `${JSON.stringify({ updatedAt: new Date().toISOString() }, null, 2)}\n`,
);
