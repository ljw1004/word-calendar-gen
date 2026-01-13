import * as fs from "fs";
import { generateCalendarDocx, Packer } from "./calendar";

const args = process.argv.slice(2);
const startMatch = args[0]?.match(/^(\d{4})-(1|2|3|4|5|6|7|8|9|10|11|12)$/);
const monthsMatch = args[1]?.match(/^(\d+)$/);
const outputPath = args[2];

if (!startMatch || !monthsMatch || !outputPath) {
  console.error(`Usage: node cli.js YYYY-M MONTHS OUTPUT.docx

Example:
  node cli.js 2025-6 3 summer.docx`);
  process.exit(1);
}

const doc = generateCalendarDocx(parseInt(startMatch[1], 10), parseInt(startMatch[2], 10), parseInt(monthsMatch[1], 10));
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Generated: ${outputPath}`);
}).catch(console.error);
