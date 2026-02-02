import { DateTime } from 'luxon';
import PDFParser from 'pdf2json';

export class Utilities {
  
  // Function to convert a date string in the format "MM/DD/YYYY" to a Date object
  convertStringToDate(date) {
    const parts = date.split("/");
    if (parts.length === 3) {
      const date = DateTime.fromObject({
        day: parseInt(parts[1], 10), 
        month: parseInt(parts[0], 10), 
        year: parseInt(parts[2], 10)}, { zone: 'utc' });
      return date.toISODate();
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
    return new Promise((resolve, reject) => {
      try {
        // Convert ArrayBuffer to Buffer if needed (from Electron IPC)
        let buffer = fileBuffer;
        if (fileBuffer instanceof ArrayBuffer) {
          buffer = Buffer.from(fileBuffer);
        } else if (fileBuffer.buffer instanceof ArrayBuffer) {
          // Handle Uint8Array
          buffer = Buffer.from(fileBuffer.buffer);
        }
        
        const pdfParser = new PDFParser();
        
        pdfParser.on('pdfParser_dataError', errData => {
          reject(new Error(errData.parserError));
        });
        
        pdfParser.on('pdfParser_dataReady', pdfData => {
          try {
            let text = "";
            
            // Extract all text from all pages
            pdfData.Pages.forEach(page => {
              page.Texts.forEach(textItem => {
                // Decode and join all text runs
                textItem.R.forEach(run => {
                  text += decodeURIComponent(run.T) + " ";
                });
              });
            });
            
            resolve(text);
          } catch (error) {
            reject(error);
          }
        });
        
        pdfParser.parseBuffer(buffer);
        
      } catch (error) {
        console.error("Error extracting text from PDF:", error);
        reject(new Error("Failed to extract text from PDF."));
      }
    });
  }
}
