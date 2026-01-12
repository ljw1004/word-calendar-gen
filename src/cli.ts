import * as fs from "fs";
import { generateCalendarDocx } from "./calendar";

function printUsage(): never {
  console.error(`Usage: word-calendar-gen --start YYYY-M --months N [-o output.docx]

Options:
  --start YYYY-M    Start month (e.g., 2025-6 for June 2025)
  --months N        Number of months to generate
  -o FILE           Output file (default: calendar.docx)

Example:
  word-calendar-gen --start 2025-6 --months 3 -o summer.docx`);
  process.exit(1);
}

function parseArgs(args: string[]): { startYear: number; startMonth: number; months: number; outputPath: string } {
  let start: string | undefined;
  let months: string | undefined;
  let outputPath = "calendar.docx";

  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--start" && args[i + 1]) {
      start = args[++i];
    } else if (args[i] === "--months" && args[i + 1]) {
      months = args[++i];
    } else if (args[i] === "-o" && args[i + 1]) {
      outputPath = args[++i];
    }
  }

  if (!start || !months) {
    printUsage();
  }

  const startMatch = start.match(/^(\d{4})-(1|2|3|4|5|6|7|8|9|10|11|12)$/);
  if (!startMatch) {
    console.error(`Error: Invalid start format "${start}". Expected YYYY-M (e.g., 2025-6)`);
    process.exit(1);
  }

  const monthsMatch = months.match(/^(\d+)$/);
  if (!monthsMatch) {
    console.error(`Error: Invalid months "${months}". Expected a number.`);
    process.exit(1);
  }

  return {
    startYear: parseInt(startMatch[1], 10),
    startMonth: parseInt(startMatch[2], 10),
    months: parseInt(monthsMatch[1], 10),
    outputPath,
  };
}

async function main() {
  const { startYear, startMonth, months, outputPath } = parseArgs(process.argv.slice(2));

  const buffer = await generateCalendarDocx(startYear, startMonth, months);
  fs.writeFileSync(outputPath, buffer);
  console.log(`Generated: ${outputPath}`);
}

main().catch(console.error);
