import { Document, Packer, Paragraph, TextRun } from "docx";

export async function generateCalendarDocx(): Promise<Uint8Array> {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            children: [new TextRun("Hello World")],
          }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
}
