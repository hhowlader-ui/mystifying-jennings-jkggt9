import { PDFDocument } from "pdf-lib";

export const getPdfPageCount = async (file: File): Promise<number> => {
  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await PDFDocument.load(arrayBuffer);
  return pdfDoc.getPageCount();
};

export async function* splitPdf(
  file: File,
  pagesPerChunk: number
): AsyncGenerator<File> {
  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await PDFDocument.load(arrayBuffer);
  const totalPages = pdfDoc.getPageCount();

  for (let i = 0; i < totalPages; i += pagesPerChunk) {
    const subDoc = await PDFDocument.create();
    const range: number[] = [];

    for (let j = 0; j < pagesPerChunk && i + j < totalPages; j++) {
      range.push(i + j);
    }

    const copiedPages = await subDoc.copyPages(pdfDoc, range);
    copiedPages.forEach((page) => subDoc.addPage(page));

    const pdfBytes = await subDoc.save();
    const chunkName = `${file.name.replace(".pdf", "")}_part_${
      Math.floor(i / pagesPerChunk) + 1
    }.pdf`;

    yield new File([pdfBytes], chunkName, { type: "application/pdf" });
  }
}
