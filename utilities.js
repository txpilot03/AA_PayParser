import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

export class Utilities {
  
  // Function to convert a date string in the format "MM/DD/YYYY" to a Date object
  convertStringToDate(date) {
    const parts = date.split("/");
    if (parts.length === 3) {
      return new Date(`${parts[2]},${parts[0]},${parts[1]}`);
    } else {
      throw new Error("Invalid date format");
    }
  }

  // Function to remove non-numeric characters from a string like commas or negative signs
  removeNonNumericChars(str) {
    return String(str).replace(/[^0-9.-]/g, "");
  }

  // Function to extract text from a PDF buffer
  async extractTextFromPDF(fileBuffer) {
    try {
      const pdfDocument = await pdfjsLib.getDocument({ data: fileBuffer })
        .promise;
      let text = "";
      const numPages = pdfDocument.numPages;

      for (let pageNum = 1; pageNum <= numPages; pageNum++) {
        const page = await pdfDocument.getPage(pageNum);
        const content = await page.getTextContent();
        text += content.items.map((item) => item.str).join(" ");
      }

      return text;
    } catch (error) {
      console.error("Error extracting text from PDF:", error);
      throw new Error("Failed to extract text from PDF.");
    }
  }
}
