/**
 * EXAMPLE: Updated worksheet-gen.js methods to work with hybrid parser
 * 
 * This shows how to add a new method that accepts parsed data objects
 * instead of raw text strings, eliminating the need for regex parsing
 */

import { TextParser } from "../parse-text.js";
import { DateTime } from "luxon";
import { Utilities } from "../utilities.js";

export class WorkSheetGenerator {
  util = new Utilities();
  textParser = new TextParser();

  /**
   * NEW METHOD: Add data to workbook from already-parsed data object
   * This bypasses all the regex parsing and uses structured data directly
   */
  async addDataToWorkbookFromParsed(workbook, parsedData) {
    const removeNonNumericChars = this.util.removeNonNumericChars;

    // ✅ NO MORE REGEX PARSING! Data is already structured
    const date = this.util.convertStringToDate(parsedData.payPeriod);
    const worksheetName = DateTime.fromISO(date.toISOString()).toFormat("ddMMMyyyy");

    // Check if sheet already exists
    if (workbook.getWorksheet(worksheetName)) {
      throw new Error(`A worksheet with the name "${worksheetName}" already exists.`);
    }

    const worksheet = workbook.addWorksheet(worksheetName);

    // Style worksheet columns and cells (same as before)
    for (let col of ["A", "B", "C", "D", "E", "F"]) {
      const column = worksheet.getColumn(col);
      if (col === "A") {
        column.width = 30;
      } else {
        column.width = 25;
      }
      column.alignment = { vertical: "middle", horizontal: "center" };
      column.font = { name: "Arial", size: 12 };
    }

    // Set background color and font for specific cells
    for (let cell of [
      "A4", "B4", "C4", "D4", "E4", "A1", "A2", "B2", "C2",
      "A6", "B6", "C6", "D6", "A7", "B7", "C7", "D7", "E6", "E7", "F7",
      "E14", "F14", "E19", "F19",
      "A32", "B32", "C32", "D32", "E32", "F32",
      "A34", "B34", "C34", "D34", "E34", "F34",
      "E36", "F36"
    ]) {
      worksheet.getCell(cell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD3D3D3" }
      };
      worksheet.getCell(cell).font = { bold: true };
    }

    // ========================================================================
    // POPULATE WORKSHEET WITH PARSED DATA
    // ========================================================================

    // Header Information
    worksheet.getCell("A1").value = "Pay Period";
    worksheet.getCell("B1").value = parsedData.payPeriod;
    
    worksheet.getCell("A2").value = "Seniority Year";
    worksheet.getCell("B2").value = parsedData.seniorityYear;
    worksheet.getCell("C2").value = parsedData.group;
    
    worksheet.getCell("A3").value = "Hourly Rate";
    worksheet.getCell("B3").value = `$${parsedData.hourlyRate}`;

    // Earnings Header
    worksheet.getCell("A4").value = "Earnings";
    worksheet.getCell("B4").value = "Rate";
    worksheet.getCell("C4").value = "Hours";
    worksheet.getCell("D4").value = "Current";
    worksheet.getCell("E4").value = "YTD";

    // ✅ Populate Earnings Data - Direct access, no regex needed!
    let row = 5;
    
    // Crew Advance
    if (parsedData.crewAdvanceCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Crew Advance";
      worksheet.getCell(`B${row}`).value = parsedData.crewAdvanceRate;
      worksheet.getCell(`C${row}`).value = parsedData.crewAdvanceHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.crewAdvanceCurrent);
      row++;
    }

    // Pilot Exp D Taxable
    if (parsedData.pltExpDtaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp D Taxable";
      worksheet.getCell(`B${row}`).value = parsedData.pltExpDtaxRate;
      worksheet.getCell(`C${row}`).value = parsedData.pltExpDtaxHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.pltExpDtaxCurrent);
      row++;
    }

    // Pilot Exp ADJ Taxable
    if (parsedData.pltExpADJtaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp ADJ Taxable";
      worksheet.getCell(`B${row}`).value = parsedData.pltExpADJtaxRate;
      worksheet.getCell(`C${row}`).value = parsedData.pltExpADJtaxHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.pltExpADJtaxCurrent);
      row++;
    }

    // Pilot Exp I Taxable
    if (parsedData.pltExpItaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp I Taxable";
      worksheet.getCell(`B${row}`).value = parsedData.pltExpItaxRate;
      worksheet.getCell(`C${row}`).value = parsedData.pltExpItaxHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.pltExpItaxCurrent);
      row++;
    }

    // AAG Profit Sharing
    if (parsedData.AagProfitSharingCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "AAG Profit Sharing";
      worksheet.getCell(`B${row}`).value = parsedData.AagProfitSharingRate;
      worksheet.getCell(`C${row}`).value = parsedData.AagProfitSharingHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.AagProfitSharingCurrent);
      row++;
    }

    // Pilot Exp D Non-Taxable
    if (parsedData.pltExpDnonTaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp D Non-Taxable";
      worksheet.getCell(`B${row}`).value = parsedData.pltExpDnonTaxRate;
      worksheet.getCell(`C${row}`).value = parsedData.pltExpDnonTaxHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.pltExpDnonTaxCurrent);
      row++;
    }

    // Pilot Exp ADJ Non-Taxable
    if (parsedData.pltExpAdjnonTaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp ADJ Non-Taxable";
      worksheet.getCell(`B${row}`).value = parsedData.pltExpAdjnonTaxRate;
      worksheet.getCell(`C${row}`).value = parsedData.pltExpAdjnonTaxHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.pltExpAdjnonTaxCurrent);
      row++;
    }

    // Pilot Exp I Non-Taxable
    if (parsedData.pltExpInonTaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp I Non-Taxable";
      worksheet.getCell(`B${row}`).value = parsedData.pltExpInonTaxRate;
      worksheet.getCell(`C${row}`).value = parsedData.pltExpInonTaxHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.pltExpInonTaxCurrent);
      row++;
    }

    // Operational Pay
    if (parsedData.operationalPayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Operational Pay";
      worksheet.getCell(`B${row}`).value = parsedData.operationalPayRate;
      worksheet.getCell(`C${row}`).value = parsedData.operationalPayHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.operationalPayCurrent);
      row++;
    }

    // Flight Training Pay
    if (parsedData.fltTrainingPayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Flight Training Pay";
      worksheet.getCell(`B${row}`).value = parsedData.fltTrainingPayRate;
      worksheet.getCell(`C${row}`).value = parsedData.fltTrainingPayHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.fltTrainingPayCurrent);
      row++;
    }

    // ... Continue for all other earnings fields ...
    // I'll show a few more examples:

    // Sick Pay
    if (parsedData.sickPayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Sick Pay";
      worksheet.getCell(`B${row}`).value = parsedData.sickPayRate;
      worksheet.getCell(`C${row}`).value = parsedData.sickPayHours;
      worksheet.getCell(`D${row}`).value = removeNonNumericChars(parsedData.sickPayCurrent);
      row++;
    }

    // ========================================================================
    // DEDUCTIONS SECTION
    // ========================================================================
    
    row = 32; // Start deductions at row 32 (adjust as needed)
    
    worksheet.getCell("A32").value = "Deductions";
    worksheet.getCell("B32").value = "Current";
    worksheet.getCell("C32").value = "YTD";

    row = 33;

    // Pre-Tax Deductions
    if (parsedData.medicalCoverage !== "0") {
      worksheet.getCell(`A${row}`).value = "Medical Coverage";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.medicalCoverage);
      row++;
    }

    if (parsedData.dentalCoverage !== "0") {
      worksheet.getCell(`A${row}`).value = "Dental Coverage";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.dentalCoverage);
      row++;
    }

    if (parsedData.visionCoverage !== "0") {
      worksheet.getCell(`A${row}`).value = "Vision Coverage";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.visionCoverage);
      row++;
    }

    if (parsedData._401k !== "0") {
      worksheet.getCell(`A${row}`).value = "401k";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData._401k);
      row++;
    }

    // Taxes
    if (parsedData.withholdingTax !== "0") {
      worksheet.getCell(`A${row}`).value = "Withholding Tax";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.withholdingTax);
      row++;
    }

    if (parsedData.socialSecurityTax !== "0") {
      worksheet.getCell(`A${row}`).value = "Social Security Tax";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.socialSecurityTax);
      row++;
    }

    if (parsedData.medicareTax !== "0") {
      worksheet.getCell(`A${row}`).value = "Medicare Tax";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.medicareTax);
      row++;
    }

    // After-Tax Deductions
    if (parsedData.roth401k !== "0") {
      worksheet.getCell(`A${row}`).value = "Roth 401k";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.roth401k);
      row++;
    }

    if (parsedData.unionDues !== "0") {
      worksheet.getCell(`A${row}`).value = "Union Dues - APA";
      worksheet.getCell(`B${row}`).value = removeNonNumericChars(parsedData.unionDues);
      row++;
    }

    // ========================================================================
    // SUMMARY SECTION
    // ========================================================================
    
    worksheet.getCell("E6").value = "Summary";
    worksheet.getCell("E7").value = "Current";
    worksheet.getCell("F7").value = "YTD";

    worksheet.getCell("E8").value = "Gross";
    worksheet.getCell("F8").value = removeNonNumericChars(parsedData.gross);

    worksheet.getCell("E9").value = "Pre-Tax Deduct";
    worksheet.getCell("F9").value = removeNonNumericChars(parsedData.preTaxDeduct);

    worksheet.getCell("E10").value = "Taxes";
    worksheet.getCell("F10").value = removeNonNumericChars(parsedData.taxes);

    worksheet.getCell("E11").value = "After Tax Deduct";
    worksheet.getCell("F11").value = removeNonNumericChars(parsedData.afterTaxDeduct);

    worksheet.getCell("E12").value = "Net Pay";
    worksheet.getCell("F12").value = removeNonNumericChars(parsedData.netPay);

    return worksheet;
  }

  /**
   * ORIGINAL METHOD: Keep for backward compatibility
   * This still uses regex parsing via TextParser
   */
  async addDataToWorkbook(workbook, extractedText) {
    // ... original implementation stays the same ...
    // This allows gradual migration
  }
}

// ========================================================================
// USAGE EXAMPLE IN MAIN.JS
// ========================================================================

/*
// In your main.js:

import { HybridPDFParser } from './hybrid-parser.js';
const hybridParser = new HybridPDFParser();

ipcMain.handle("parse-pdf", async (event, pdfFile, pdfFileBuffer, outputPath, outputFileName) => {
  try {
    // Parse PDF with hybrid parser
    const parsedData = await hybridParser.parse(pdfFileBuffer);
    const legacyData = hybridParser.convertToLegacyFormat(parsedData);

    // Create/load workbook
    const workbook = new Workbook();
    
    // ✅ Use new method that accepts parsed data
    await workSheetGenerator.addDataToWorkbookFromParsed(workbook, legacyData);
    
    // Rest of your code...
    await deductTotalsWS.getAllDeductions(workbook);
    await earningTotalsWS.getAllEarnings(workbook);
    await workbook.xlsx.writeFile(outputFilePath);
    
    return outputFilePath;
  } catch (error) {
    throw new Error(`Failed to parse PDF file: ${error.message}`);
  }
});
*/
