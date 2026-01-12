import * as fs from "fs";
import { generateCalendarDocx } from "./calendar";

async function main() {
  const args = process.argv.slice(2);
  let outputPath = "calendar.docx";

  for (let i = 0; i < args.length; i++) {
    if (args[i] === "-o" && args[i + 1]) {
      outputPath = args[i + 1];
      i++;
    }
  }

  const buffer = await generateCalendarDocx();
  fs.writeFileSync(outputPath, buffer);
  console.log(`Generated: ${outputPath}`);
}

main().catch(console.error);
