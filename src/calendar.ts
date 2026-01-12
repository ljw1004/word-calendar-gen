import { Document, Packer, Paragraph, TextRun } from "docx";

export async function generateCalendarDocx(startYear: number, startMonth: number, months: number): Promise<Uint8Array> {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            children: [new TextRun(`Calendar: ${startYear}-${String(startMonth).padStart(2, "0")} for ${months} months`)],
          }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
}
