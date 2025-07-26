import { app, BrowserWindow, ipcMain, dialog } from "electron";
import pkg from "exceljs";
const { Workbook } = pkg;
import path from "node:path";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { DateTime} from "luxon";
import fs from "fs";

let window: any;

function createWindow() {
  window = new BrowserWindow({
    width: 800,
    height: 500,
    webPreferences: {
      nodeIntegration: true,
      preload: path.join(process.cwd(), "preload.js"),
      contextIsolation: true,
    },
  });

  window.loadFile("index.html");
}

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

// IPC to handle the file reading
ipcMain.handle("parse-pdf", async (event, pdfFile, excelFile ) => {
  try {
    if (!pdfFile) {
      throw new Error("PDF file buffer is undefined or null");
    }

    const extractedText = await extractTextFromPDF(pdfFile);
    // let extractedText = "";

    // const loadingTask = pdfjsLib.getDocument({ data: pdfFile });
    // const pdf = await loadingTask.promise;
    // for (let i = 1; i <= pdf.numPages; i++) {
    //   const page = await pdf.getPage(i);
    //   const content = await page.getTextContent();
    //   extractedText += content.items
    //     .map((item) => ('str' in item ? item.str : ''))
    //     .join(" ");
    // }

    let outputFilePath = path.join(app.getPath('downloads'), 'My_AA_Pay.xlsx');

    // If Excel file buffer exists, process the existing file, else create a new file
    if (excelFile) {
      // If the Excel file exists, open and modify it
      const workbook = new Workbook();
      await workbook.xlsx.load(excelFile);

      // Add new worksheet and populate data
      await addDataToWorkbook(workbook, extractedText);
      
      // Save the modified workbook to the same file path
      await workbook.xlsx.writeFile(outputFilePath);
    } else {
      // If no Excel file is provided, create a new workbook and add data
      const newWorkbook = new Workbook();
      
      // Add the first sheet with parsed data
      await addDataToWorkbook(newWorkbook, extractedText);
      
      // Save the new workbook
      await newWorkbook.xlsx.writeFile(outputFilePath);
    }

    return outputFilePath;  // Return the file path for the saved Excel file
  } catch (error) {
    console.error("Error during PDF parsing:", error);
    throw new Error(`Failed to parse PDF file: ${error.message}`);
  }
});



async function addDataToWorkbook(workbook: any, extractedText: string) {
  const infoData = parsePDFInfoData(extractedText);
    const summaryData = parsePDFSummaryData(extractedText);
    const earningsTopData = parsePDFEarningsTopData(extractedText);
    const earningsBottomData = parsePDFEarningsBottomData(extractedText);
    const preTaxDeductionsData = parsePDFPreTaxDeductionsData(extractedText);
    const taxDeductionsData = parsePDFTaxDeductionsData(extractedText);
    const afterTaxDeductionsData = parsePDFAfterTaxDeductionsData(extractedText);
    const companyContributionsData = parseCompanyContributionsData(extractedText);
    const taxableEarningsData = parsePDFTaxableEarningsData(extractedText);

    const date = convertStringToDate(infoData.payPeriod);
    const worksheetName = DateTime.fromISO(date.toISOString()).toFormat("MMddyyyy");

    // Check if the sheet already exists
    if (workbook.getWorksheet(worksheetName)) {
      throw new Error(`A worksheet with the name "${worksheetName}" already exists.`);
    }

    const worksheet = workbook.addWorksheet(worksheetName);

    // Style worksheet columns and cells
    for (let col of ["A", "B", "C", "D", "E", "F"]) {
      const column = worksheet.getColumn(col);
      if (col === "A") {
        column.width = 30; // Set width for column A
      } else {
        column.width = 25; // Set width for columns B to F
      }
      column.alignment = { vertical: "middle", horizontal: "center" }; // Center align text
      column.font = { name: "Arial", size: 12 }; // Set font for
    }

    // Set background color and font for specific cells
    for (let cell of [
      "A4","B4","C4","D4","E4","A1","A2","B2","C2","A6","B6","C6","D6","A7",
      "B7","C7","D7", "E6", "E7", "F7","E14","F14","E19","F19","A32","B32",
      "C32","D32","E32","F32","A34","B34","C34", "D34", "E34", "F34","E36","F36"
    ]) {
      worksheet.getCell(cell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFBFBFBF" },
      };
      worksheet.getCell(cell).font = {
        bold: true,
      };
    }

    for (let cell of ["A6","A7","B7","C7","D7","E6","E7","F7","E14","F14",
      "E19","F19","A32","B32","C32","D32","E32","F32","A34","D34","E34","F34","E36","F36"
    ]) {
      worksheet.getCell(cell).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }

    for (let cell of ["A8", "A9", "A10", "A11", "A12", "A13", "A14", "A15", "A16",
      "A17", "A18", "A19", "A20", "A21", "A22", "A23", "A24", "A25", "A26", "A27", 
      "A28", "A29", "A30", "A31", "A35", "A36", "A37","E35","E37"
    ]) {
      worksheet.getCell(cell).alignment = { vertical: "middle", horizontal: "left" };
    }

    // Fill in the worksheet with the extracted data
    //? Header row
    worksheet.getCell("A1").value = "Pay Period";
    worksheet.getCell("B1").value = date;
    worksheet.getCell("B1").numFmt = "MMM dd, yyyy";

    worksheet.getCell("A2").value = "Seniority Year";
    worksheet.getCell("A3").value = Number(infoData.seniorityYear);

    worksheet.getCell("B2").value = "Group";
    worksheet.getCell("B3").value = infoData.group;

    worksheet.getCell("C2").value = "Hourly Rate";
    worksheet.getCell("C3").value = Number(removeNonNumericChars(infoData.hourlyRate));
    worksheet.getCell("C3").numFmt = "$#,##0.00";

    //? Summary
    worksheet.getCell("A4").value = "Gross Earnings";
    worksheet.getCell("B4").value = "Pre-Tax Deduction";
    worksheet.getCell("C4").value = "Taxes";
    worksheet.getCell("D4").value = "After-Tax Deduction";
    worksheet.getCell("E4").value = "Net Pay";

    worksheet.getCell("A5").value = Number(removeNonNumericChars(summaryData.gross));
    worksheet.getCell("A5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("B5").value = Number(removeNonNumericChars(summaryData.preTaxDeduct));
    worksheet.getCell("B5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("C5").value = Number(removeNonNumericChars(summaryData.taxes));
    worksheet.getCell("C5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("D5").value = Number(removeNonNumericChars(summaryData.afterTaxDeduct));
    worksheet.getCell("D5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E5").value = Number(removeNonNumericChars(summaryData.netPay));
    worksheet.getCell("E5").numFmt = "$#,##0.00"; // Format as currency

    //? EARNINGS
    worksheet.getCell("A6").value = "EARNINGS";
    worksheet.mergeCells("A6:D6");

    //* Earnings Top
    // Header row
    worksheet.getCell("A7").value = "Earnings";
    worksheet.getCell("B7").value = "Rate";
    worksheet.getCell("C7").value = "Hours";
    worksheet.getCell("D7").value = "Current";

    worksheet.getCell("A8").value = "Crew Advance";
    worksheet.getCell("B8").value = Number(earningsTopData.crewAdvanceRate);
    worksheet.getCell("C8").value = Number(earningsTopData.crewAdvanceHours);
    worksheet.getCell("D8").value = Number(removeNonNumericChars(earningsTopData.crewAdvanceCurrent));
    console.log("Crew Advance Current:", removeNonNumericChars(earningsTopData.crewAdvanceCurrent));
    worksheet.getCell("D8").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A9").value = "PLT EXP D Taxable";
    worksheet.getCell("B9").value = Number(earningsTopData.pltExpDtaxRate);
    worksheet.getCell("C9").value = Number(earningsTopData.pltExpDtaxHours);
    worksheet.getCell("D9").value = Number(removeNonNumericChars(earningsTopData.pltExpDtaxCurrent));
    worksheet.getCell("D9").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("A10").value = "PLT EXP ADJ Taxable";
    worksheet.getCell("B10").value = Number(earningsTopData.pltExpADJtaxRate);
    worksheet.getCell("C10").value = Number(earningsTopData.pltExpADJtaxHours);
    worksheet.getCell("D10").value = Number(removeNonNumericChars(earningsTopData.pltExpADJtaxCurrent));
    worksheet.getCell("D10").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A11").value = "PLT EXP I Taxable";
    worksheet.getCell("B11").value = Number(earningsTopData.pltExpItaxRate);
    worksheet.getCell("C11").value = Number(earningsTopData.pltExpItaxHours);
    worksheet.getCell("D11").value = Number(removeNonNumericChars(earningsTopData.pltExpItaxCurrent));
    worksheet.getCell("D11").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A12").value = "AAG Profit Sharing";
    worksheet.getCell("B12").value = Number(earningsTopData.AagProfitSharingRate);
    worksheet.getCell("C12").value = Number(earningsTopData.AagProfitSharingHours);
    worksheet.getCell("D12").value = Number(removeNonNumericChars(earningsTopData.AagProfitSharingCurrent));
    worksheet.getCell("D12").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A13").value = "PLT EXP D Non-Taxable";
    worksheet.getCell("B13").value = Number(earningsTopData.pltExpDnonTaxRate);
    worksheet.getCell("C13").value = Number(earningsTopData.pltExpDnonTaxHours);
    worksheet.getCell("D13").value = Number(removeNonNumericChars(earningsTopData.pltExpDnonTaxCurrent));
    worksheet.getCell("D13").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A14").value = "PLT EXP ADJ Non-Taxable";
    worksheet.getCell("B14").value = Number(earningsTopData.pltExpAdjnonTaxRate);
    worksheet.getCell("C14").value = Number(earningsTopData.pltExpAdjnonTaxHours);
    worksheet.getCell("D14").value = Number(removeNonNumericChars(earningsTopData.pltExpAdjnonTaxCurrent));
    worksheet.getCell("D14").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A15").value = "PLT EXP I Non-Taxable";
    worksheet.getCell("B15").value = Number(earningsTopData.pltExpInonTaxRate);
    worksheet.getCell("C15").value = Number(earningsTopData.pltExpInonTaxHours);
    worksheet.getCell("D15").value = Number(removeNonNumericChars(earningsTopData.pltExpInonTaxCurrent));
    worksheet.getCell("D15").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("A16").value = "Earnings Total";
    worksheet.getCell("A16").font = { bold: true };
    worksheet.getCell("D16").value = { formula: 'SUM(D8:D15)', result: 7 };
    worksheet.getCell("D16").font = { bold: true };
    worksheet.getCell("D16").numFmt = "$#,##0.00"; // Format as currency
    

    //* Earnings Bottom
    worksheet.getCell("A18").value = "Operational Pay";
    worksheet.getCell("B18").value = Number(earningsBottomData.operationalPayRate);
    worksheet.getCell("C18").value = Number(earningsBottomData.operationalPayHours);
    worksheet.getCell("D18").value = Number(removeNonNumericChars(earningsBottomData.operationalPayCurrent));
    worksheet.getCell("D18").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A19").value = "Flight Training Pay";
    worksheet.getCell("B19").value = Number(earningsBottomData.fltTrainingPayRate);
    worksheet.getCell("C19").value = Number(earningsBottomData.fltTrainingPayHours);
    worksheet.getCell("D19").value = Number(removeNonNumericChars(earningsBottomData.fltTrainingPayCurrent));
    worksheet.getCell("D19").numFmt = "$#,##0.00"; // Format as currency  
    worksheet.getCell("A20").value = "Sit Time";
    worksheet.getCell("B20").value = Number(earningsBottomData.sitTimeRate);
    worksheet.getCell("C20").value = Number(earningsBottomData.sitTimeHours);
    worksheet.getCell("D20").value = Number(removeNonNumericChars(earningsBottomData.sitTimeCurrent));
    worksheet.getCell("D20").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A21").value = "Pay Above Guarantee (RSV)";
    worksheet.getCell("B21").value = Number(earningsBottomData.payAbvGuaranteeRsvRate);
    worksheet.getCell("C21").value = Number(earningsBottomData.payAbvGuaranteeRsvHours);
    worksheet.getCell("D21").value = Number(removeNonNumericChars(earningsBottomData.payAbvGuaranteeRsvCurrent));
    worksheet.getCell("D21").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A22").value = "RA Prem";
    worksheet.getCell("B22").value = Number(earningsBottomData.raPremRate);
    worksheet.getCell("C22").value = Number(earningsBottomData.raPremHours);
    worksheet.getCell("D22").value = Number(removeNonNumericChars(earningsBottomData.raPremCurrent));
    worksheet.getCell("D22").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A23").value = "Min Guarantee Adj";
    worksheet.getCell("B23").value = Number(earningsBottomData.minGuaranteeAdjRate);
    worksheet.getCell("C23").value = Number(earningsBottomData.minGuaranteeAdjHours);
    worksheet.getCell("D23").value = Number(removeNonNumericChars(earningsBottomData.minGuaranteeAdjCurrent));
    worksheet.getCell("D23").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A24").value = "Intl Override";
    worksheet.getCell("B24").value = Number(earningsBottomData.intlOverrideRate);
    worksheet.getCell("C24").value = Number(earningsBottomData.intlOverrideHours);
    worksheet.getCell("D24").value = Number(removeNonNumericChars(earningsBottomData.intlOverrideCurrent));
    worksheet.getCell("D24").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A25").value = "Distance Learning";
    worksheet.getCell("B25").value = Number(earningsBottomData.distanceLearningRate);
    worksheet.getCell("C25").value = Number(earningsBottomData.distanceLearningHours);
    worksheet.getCell("D25").value = Number(removeNonNumericChars(earningsBottomData.distanceLearningCurrent));
    worksheet.getCell("D25").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A26").value = "Union PD Union Leave";
    worksheet.getCell("B26").value = Number(earningsBottomData.unionPdLeaveRate);
    worksheet.getCell("C26").value = Number(earningsBottomData.unionPdLeaveHours);
    worksheet.getCell("D26").value = Number(removeNonNumericChars(earningsBottomData.unionPdLeaveCurrent));
    worksheet.getCell("D26").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A27").value = "Prem Incentive Pay";
    worksheet.getCell("B27").value = Number(earningsBottomData.premIncentivePayRate);
    worksheet.getCell("C27").value = Number(earningsBottomData.premIncentivePayHours);
    worksheet.getCell("D27").value = Number(removeNonNumericChars(earningsBottomData.premIncentivePayCurrent));
    worksheet.getCell("D27").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A28").value = "Flight Vacation Pay";
    worksheet.getCell("B28").value = Number(earningsBottomData.fltVacationPayRate);
    worksheet.getCell("C28").value = Number(earningsBottomData.fltVacationPayHours);
    worksheet.getCell("D28").value = Number(removeNonNumericChars(earningsBottomData.fltVacationPayCurrent));
    worksheet.getCell("D28").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A29").value = "Sick Pay";
    worksheet.getCell("B29").value = Number(earningsBottomData.sickPayRate);
    worksheet.getCell("C29").value = Number(earningsBottomData.sickPayHours);
    worksheet.getCell("D29").value = Number(removeNonNumericChars(earningsBottomData.sickPayCurrent));
    worksheet.getCell("D29").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A30").value = "Prior Year Vacation Pay Out";
    worksheet.getCell("B30").value = Number(earningsBottomData.priorYearVacPayoutRate);
    worksheet.getCell("C30").value = Number(earningsBottomData.priorYearVacPayoutHours);
    worksheet.getCell("D30").value = Number(earningsBottomData.priorYearVacPayoutCurrent);
    worksheet.getCell("D30").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("A31").value = "Earnings Sub Total";
    worksheet.getCell("A31").font = { bold: true };
    worksheet.getCell("D31").value = { formula: 'SUM(D18:D30)', result: 7 };
    worksheet.getCell("D31").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("D31").font = { bold: true };


    //? Deductions
    worksheet.getCell("E6").value = "DEDUCTIONS";
    worksheet.mergeCells("E6:F6");
    for (let cell of ["E8", "E9", "E10", "E11", "E12", "E13", "E15", "E16", "E17", "E18", "E20", "E21", "E22", "E23", "E24", "E25"]) {
      worksheet.getCell(cell).alignment = { vertical: "middle", horizontal: "left" };
    }
    //* Pre-Tax Deductions
    worksheet.getCell("E7").value = "Pre-Tax Deductions";
    worksheet.getCell("F7").value = "Current";
    worksheet.getCell("E8").value = "Medical Coverage";
    worksheet.getCell("F8").value = Number(removeNonNumericChars(preTaxDeductionsData.medicalCoverage));
    worksheet.getCell("F8").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E9").value = "Dental Coverage";
    worksheet.getCell("F9").value = Number(removeNonNumericChars(preTaxDeductionsData.dentalCoverage));
    worksheet.getCell("F9").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E10").value = "Vision Coverage";
    worksheet.getCell("F10").value = Number(removeNonNumericChars(preTaxDeductionsData.visionCoverage));
    worksheet.getCell("F10").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E11").value = "Accident Ins Pre-tax";
    worksheet.getCell("F11").value = Number(removeNonNumericChars(preTaxDeductionsData.accidentInsPreTax));
    worksheet.getCell("F11").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E12").value = "401k";
    worksheet.getCell("F12").value = Number(removeNonNumericChars(preTaxDeductionsData._401k));
    worksheet.getCell("F12").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("E13").value = "Pre-Tax Deductions Total";
    worksheet.getCell("E13").font = { bold: true };
    worksheet.getCell("F13").value = { formula: 'SUM(F8:F12)', result: 7 };
    worksheet.getCell("F13").font = { bold: true };
    worksheet.getCell("F13").numFmt = "$#,##0.00"; // Format as currency

    //* Taxes
    worksheet.getCell("E14").value = "Taxes";
    worksheet.getCell("F14").value = "Current";
    worksheet.getCell("E15").value = "Withholding Tax";
    worksheet.getCell("F15").value = Number(removeNonNumericChars(taxDeductionsData.withholdingTax));
    worksheet.getCell("F15").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E16").value = "Social Security Tax";
    worksheet.getCell("F16").value = Number(removeNonNumericChars(taxDeductionsData.socialSecurityTax));
    worksheet.getCell("F17").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E17").value = "Medicare Tax";
    worksheet.getCell("F17").value = Number(removeNonNumericChars(taxDeductionsData.medicareTax));
    worksheet.getCell("F17").numFmt = "$#,##0.00"; // Format as currency
    
    worksheet.getCell("E18").value = "Taxes Total";
    worksheet.getCell("E18").font = { bold: true };
    worksheet.getCell("F18").value = { formula: 'SUM(F15:F17)', result: 7 };
    worksheet.getCell("F18").font = { bold: true };
    worksheet.getCell("F18").numFmt = "$#,##0.00"; // Format as currency

    //* After-Tax Deductions
    worksheet.getCell("E19").value = "After-Tax Deductions";
    worksheet.getCell("F19").value = "Current";
    worksheet.getCell("E20").value = "Employee Life";
    worksheet.getCell("F20").value = Number(removeNonNumericChars(afterTaxDeductionsData.employeeLife));
    worksheet.getCell("F20").numFmt = "$#,##0.00"; 
    worksheet.getCell("E21").value = "Dental Discount Plan";
    worksheet.getCell("F21").value = Number(removeNonNumericChars(afterTaxDeductionsData.dentalDiscountPlan));
    worksheet.getCell("F21").numFmt = "$#,##0.00"; 
    worksheet.getCell("E22").value = "Roth 401k";
    worksheet.getCell("F22").value = Number(removeNonNumericChars(afterTaxDeductionsData.roth401k));
    worksheet.getCell("F22").numFmt = "$#,##0.00"; 
    worksheet.getCell("E23").value = "PAC - APA";
    worksheet.getCell("F23").value = Number(removeNonNumericChars(afterTaxDeductionsData.pacAPA));
    worksheet.getCell("F23").numFmt = "$#,##0.00"; 
    worksheet.getCell("E24").value = "Union Dues - APA";
    worksheet.getCell("F24").value = Number(removeNonNumericChars(afterTaxDeductionsData.unionDues));
    worksheet.getCell("F24").numFmt = "$#,##0.00"; 

    worksheet.getCell("E25").value = "After-Tax Deductions Total";
    worksheet.getCell("E25").font = { bold: true };
    worksheet.getCell("F25").value = { formula: 'SUM(F20:F24)', result: 7 };
    worksheet.getCell("F25").font = { bold: true };
    worksheet.getCell("F25").numFmt = "$#,##0.00"; 

    worksheet.mergeCells("E26:F31");

    //? Taxable Earnings
    worksheet.getCell("A34").value = "Taxable Earnings - Federal Taxes";
    worksheet.mergeCells("A34:C34");
    worksheet.getCell("D34").value = "Current";

    worksheet.getCell("A35").value = "Withholding Tax";
    worksheet.mergeCells("A35:C35");
    worksheet.getCell("D35").value = Number(removeNonNumericChars(taxableEarningsData.withHoldingTaxEarnings));
    worksheet.getCell("D35").numFmt = "$#,##0.00";

    worksheet.getCell("A36").value = "Social Security Tax";
    worksheet.mergeCells("A36:C36");
    worksheet.getCell("D36").value = Number(removeNonNumericChars(taxableEarningsData.socialSecurityTaxEarnings));
    worksheet.getCell("D36").numFmt = "$#,##0.00";

    worksheet.getCell("A37").value = "Medicare Tax";
    worksheet.mergeCells("A37:C37");
    worksheet.getCell("D37").value = Number(removeNonNumericChars(taxableEarningsData.medicareTaxEarnings));
    worksheet.getCell("D37").numFmt = "$#,##0.00";

    //? Overpayments
    worksheet.getCell("A32").value = "Overpayments";
    worksheet.getCell("B32").value = "Moved A/R";
    worksheet.getCell("C32").value = "Original Balance";
    worksheet.getCell("D32").value = "Current Recovery";
    worksheet.getCell("E32").value = "Total Recovery";
    worksheet.getCell("F32").value = "Overpayments Total";

    worksheet.getCell("A33").value = "NOT TRACKED - FOR FUTURE USE";
    worksheet.mergeCells("A33:F33");



    //? Additional Information
    worksheet.getCell("E34").value = "Additional Information";
    worksheet.getCell("F34").value = "Current";

    worksheet.getCell("E35").value = "401k Company Contrib.";
    worksheet.getCell("F35").value = Number(removeNonNumericChars(companyContributionsData._401kCompanyContribution));
    worksheet.getCell("F35").numFmt = "$#,##0.00";

    worksheet.getCell("E36").value = "Imputed Income";
    worksheet.getCell("F36").value = "Current";
    worksheet.getCell("E37").value = "Group Term Life";
    worksheet.getCell("F37").value = Number(removeNonNumericChars(companyContributionsData.groupTermLife));
    worksheet.getCell("F37").numFmt = "$#,##0.00"; //

    worksheet.getCell("A40").value = extractedText; // Store the full extracted text in A40


    return worksheet;
}

async function addTotalsToWorkbook(workbook: any) {

  let medicalCoverage = [];
  workbook.eachSheet( (sheet) => {
    const totalMedicalCoverage = sheet.getCell("F8").value;
    console.log("Total Medical Coverage:", totalMedicalCoverage);
  })
}

function parsePDFInfoData(extractedText: string) {
  const payPeriodRegex = /1-800-447-2000\s*([\d]+(?:\/\d+)+)/i;
  const seniorityYearRegex = /Effective\s*([\d]+)/i;
  const groupRegex = /Effective\s*\d+\s+([A-Z]{2})/i;
  const hourlyRate = /Effective\s*\d+\s+([A-Z]{2})\s+(\$\d+\.\d{2})/i;

  const payPeriodMatch = extractedText.match(payPeriodRegex);
  const seniorityYearMatch = extractedText.match(seniorityYearRegex);
  const groupMatch = extractedText.match(groupRegex);
  const hourlyRateMatch = extractedText.match(hourlyRate);

  return {
    payPeriod: payPeriodMatch ? payPeriodMatch[1] : "",
    seniorityYear: seniorityYearMatch ? seniorityYearMatch[1] : "0",
    group: groupMatch ? groupMatch[1] : "",
    hourlyRate: hourlyRateMatch ? hourlyRateMatch[2] : "0",
  };
}

function parsePDFSummaryData(extractedText: string) {
  const grossRegex = /Current\s*([\d,]+(?:\.\d+)?)/i;
  const preTaxDeductRegex =
    /Current\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const taxesRegex =
    /Current\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const afterTaxDeductRegex =
    /Current\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const netPayRegex =
    /Current\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;

  const grossMatch = extractedText.match(grossRegex);
  const preTaxDeductMatch = extractedText.match(preTaxDeductRegex);
  const taxesMatch = extractedText.match(taxesRegex);
  const afterTaxDeductMatch = extractedText.match(afterTaxDeductRegex);
  const netPayMatch = extractedText.match(netPayRegex);
  
  return {
    gross: grossMatch ? grossMatch[1] : "0",
    preTaxDeduct: preTaxDeductMatch ? preTaxDeductMatch[2] : "0",
    taxes: taxesMatch ? taxesMatch[3] : "0",
    afterTaxDeduct: afterTaxDeductMatch ? afterTaxDeductMatch[4] : "0",
    netPay: netPayMatch ? netPayMatch[5] : "0",
  };
}

function parsePDFEarningsTopData(extractedText: string) {
  const crewAdvanceRegex = /Advance\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*((?:-\d+[\d,]+(?:\.\d+)?|\d+[\d,]+(?:\.\d+)?))/i
  ///Advance\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*((?:\-\d+[\d,]+(?:\.\d+)?))/i;
  const pltExpDtaxRegex = /D\s*Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const pltExpADJtax = /ADJ\s*Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const pltExpItax = /I\s*Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const AagProfitSharingRegex = /Profit\s*Sharing\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const pltExpDnonTaxRegex = /D\s*Non-Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const pltExpAdjnonTaxRegex = /ADJ\s*Non-Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*((?:\-\d+[\d,]+(?:\.\d+)?))/i;
  const pltExpInonTaxRegex = /I\s*Non-Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;

  const crewAdvanceMatch = extractedText.match(crewAdvanceRegex);
  const pltExpDtaxMatch = extractedText.match(pltExpDtaxRegex);
  const pltExpADJtaxMatch = extractedText.match(pltExpADJtax);
  const pltExpItaxMatch = extractedText.match(pltExpItax);
  const AagProfitSharingMatch = extractedText.match(AagProfitSharingRegex);
  const pltExpDnonTaxMatch = extractedText.match(pltExpDnonTaxRegex);
  const pltExpAdjnonTaxMatch = extractedText.match(pltExpAdjnonTaxRegex);
  const pltExpInonTaxMatch = extractedText.match(pltExpInonTaxRegex);

  return {
    crewAdvanceRate: crewAdvanceMatch ? crewAdvanceMatch[1] : "0",
    crewAdvanceHours: crewAdvanceMatch ? crewAdvanceMatch[2] : "0",
    crewAdvanceCurrent: crewAdvanceMatch ? crewAdvanceMatch[3] : "0",
    pltExpDtaxRate: pltExpDtaxMatch ? pltExpDtaxMatch[1] : "0",
    pltExpDtaxHours: pltExpDtaxMatch ? pltExpDtaxMatch[2] : "0",
    pltExpDtaxCurrent: pltExpDtaxMatch ? pltExpDtaxMatch[3] : "0",
    pltExpADJtaxRate: pltExpADJtaxMatch ? pltExpADJtaxMatch[1] : "0",
    pltExpADJtaxHours: pltExpADJtaxMatch ? pltExpADJtaxMatch[2] : "0",
    pltExpADJtaxCurrent: pltExpADJtaxMatch ? pltExpADJtaxMatch[3] : "0",
    pltExpItaxRate: pltExpItaxMatch ? pltExpItaxMatch[1] : "0",
    pltExpItaxHours: pltExpItaxMatch ? pltExpItaxMatch[2] : "0",
    pltExpItaxCurrent: pltExpItaxMatch ? pltExpItaxMatch[3] : "0",
    AagProfitSharingRate: AagProfitSharingMatch ? AagProfitSharingMatch[1] : "0",
    AagProfitSharingHours: AagProfitSharingMatch ? AagProfitSharingMatch[2] : "0",
    AagProfitSharingCurrent: AagProfitSharingMatch ? AagProfitSharingMatch[3] : "0",
    pltExpDnonTaxRate: pltExpDnonTaxMatch ? pltExpDnonTaxMatch[1] : "0",
    pltExpDnonTaxHours: pltExpDnonTaxMatch ? pltExpDnonTaxMatch[2] : "0",
    pltExpDnonTaxCurrent: pltExpDnonTaxMatch ? pltExpDnonTaxMatch[3] : "0",
    pltExpAdjnonTaxRate: pltExpAdjnonTaxMatch ? pltExpAdjnonTaxMatch[1] : "0",
    pltExpAdjnonTaxHours: pltExpAdjnonTaxMatch ? pltExpAdjnonTaxMatch[2] : "0",
    pltExpAdjnonTaxCurrent: pltExpAdjnonTaxMatch ? pltExpAdjnonTaxMatch[3] : "0",
    pltExpInonTaxRate: pltExpInonTaxMatch ? pltExpInonTaxMatch[1] : "0",
    pltExpInonTaxHours: pltExpInonTaxMatch ? pltExpInonTaxMatch[2] : "0",
    pltExpInonTaxCurrent: pltExpInonTaxMatch ? pltExpInonTaxMatch[3] : "0"
  };
}

function parsePDFEarningsBottomData(extractedText: string) {
  const operationalPayRegex = /Operational\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const fltTrainingPayRegex = /Flight\s*Training\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const sitTime = /Sit\s*Time\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const payAbvGuaranteeRsvRegex = /Pay\s*Above\s*Guar\s*\(RSV\)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const raPrem = /RA\s*PREM\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const minGuaranteeAdjRegex = /Min\s*Guarantee\s*Adj\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const intlOverrideRegex = /Intl\s*Override\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const distanceLearningRegex = /Distance\s*Learning\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const unionPdLeaveRegex = /Union\s*PD\s*Union\s*Leave\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const premIncentivePayRegex = /Prem\s*Incentive\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const fltVacationPayRegex = /Flight\s*Vacation\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const sickPayRegex = /Sick\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
  const priorYearVacPayoutRegex = /Prior\s*Year\s*Vacation\s*Pay\s*Out\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;

  const operationalPayMatch = extractedText.match(operationalPayRegex);
  const fltTrainingPayMatch = extractedText.match(fltTrainingPayRegex);
  const sitTimeMatch = extractedText.match(sitTime);
  const payAbvGuaranteeRsvMatch = extractedText.match(payAbvGuaranteeRsvRegex);
  const raPremMatch = extractedText.match(raPrem);
  const minGuaranteeAdjMatch = extractedText.match(minGuaranteeAdjRegex);
  const intlOverrideMatch = extractedText.match(intlOverrideRegex);
  const distanceLearningMatch = extractedText.match(distanceLearningRegex);
  const unionPdLeaveMatch = extractedText.match(unionPdLeaveRegex);
  const premIncentivePayMatch = extractedText.match(premIncentivePayRegex);
  const fltVacationPayMatch = extractedText.match(fltVacationPayRegex);
  const sickPayMatch = extractedText.match(sickPayRegex);
  const priorYearVacPayoutMatch = extractedText.match(priorYearVacPayoutRegex);

  return {
    operationalPayRate: operationalPayMatch ? operationalPayMatch[1] : "0",
    operationalPayHours: operationalPayMatch ? operationalPayMatch[2] : "0",
    operationalPayCurrent: operationalPayMatch ? operationalPayMatch[3] : "0",
    fltTrainingPayRate: fltTrainingPayMatch ? fltTrainingPayMatch[1] : "0",
    fltTrainingPayHours: fltTrainingPayMatch ? fltTrainingPayMatch[2] : "0",
    fltTrainingPayCurrent: fltTrainingPayMatch ? fltTrainingPayMatch[3] : "0",
    sitTimeRate: sitTimeMatch ? sitTimeMatch[1] : "0",
    sitTimeHours: sitTimeMatch ? sitTimeMatch[2] : "0",
    sitTimeCurrent: sitTimeMatch ? sitTimeMatch[3] : "0",
    payAbvGuaranteeRsvRate: payAbvGuaranteeRsvMatch ? payAbvGuaranteeRsvMatch[1] : "0",
    payAbvGuaranteeRsvHours: payAbvGuaranteeRsvMatch ? payAbvGuaranteeRsvMatch[2] : "0",
    payAbvGuaranteeRsvCurrent: payAbvGuaranteeRsvMatch ? payAbvGuaranteeRsvMatch[3] : "0",
    raPremRate: raPremMatch ? raPremMatch[1] : "0",
    raPremHours: raPremMatch ? raPremMatch[2] : "0",
    raPremCurrent: raPremMatch ? raPremMatch[3] : "0",
    minGuaranteeAdjRate: minGuaranteeAdjMatch ? minGuaranteeAdjMatch[1] : "0",
    minGuaranteeAdjHours: minGuaranteeAdjMatch ? minGuaranteeAdjMatch[2] : "0",
    minGuaranteeAdjCurrent: minGuaranteeAdjMatch ? minGuaranteeAdjMatch[3] : "0",
    intlOverrideRate: intlOverrideMatch ? intlOverrideMatch[1] : "0",
    intlOverrideHours: intlOverrideMatch ? intlOverrideMatch[2] : "0",
    intlOverrideCurrent: intlOverrideMatch ? intlOverrideMatch[3] : "0",
    distanceLearningRate: distanceLearningMatch ? distanceLearningMatch[1] : "0",
    distanceLearningHours: distanceLearningMatch ? distanceLearningMatch[2] : "0",
    distanceLearningCurrent: distanceLearningMatch ? distanceLearningMatch[3] : "0",
    unionPdLeaveRate: unionPdLeaveMatch ? unionPdLeaveMatch[1] : "0",
    unionPdLeaveHours: unionPdLeaveMatch ? unionPdLeaveMatch[2] : "0",
    unionPdLeaveCurrent: unionPdLeaveMatch ? unionPdLeaveMatch[3] : "0",
    premIncentivePayRate: premIncentivePayMatch ? premIncentivePayMatch[1] : "0",
    premIncentivePayHours: premIncentivePayMatch ? premIncentivePayMatch[2] : "0",
    premIncentivePayCurrent: premIncentivePayMatch ? premIncentivePayMatch[3] : "0",
    fltVacationPayRate: fltVacationPayMatch ? fltVacationPayMatch[1] : "0",
    fltVacationPayHours: fltVacationPayMatch ? fltVacationPayMatch[2] : "0",
    fltVacationPayCurrent: fltVacationPayMatch ? fltVacationPayMatch[3] : "0",
    sickPayRate: sickPayMatch ? sickPayMatch[1] : "0",
    sickPayHours: sickPayMatch ? sickPayMatch[2] : "0",
    sickPayCurrent: sickPayMatch ? sickPayMatch[3] : "0",
    priorYearVacPayoutRate: priorYearVacPayoutMatch ? priorYearVacPayoutMatch[1] : "0",
    priorYearVacPayoutHours: priorYearVacPayoutMatch ? priorYearVacPayoutMatch[2] : "0",
    priorYearVacPayoutCurrent: priorYearVacPayoutMatch ? priorYearVacPayoutMatch[3] : "0"
  }
}

function parsePDFPreTaxDeductionsData(extractedText: string) {
  const medicalCoverageRegex = /Medical\s*Coverage\s*([\d,]+(?:\.\d+)?)/i;
  const dentalCoverageRegex = /Dental\s*Coverage\s*([\d,]+(?:\.\d+)?)/i;
  const visionCoverageRegex = /Vision\s*Coverage\s*([\d,]+(?:\.\d+)?)/i;
  const accidentInsPreTaxRegex = /Accident\s*Ins\s*Pre-tax\s*([\d,]+(?:\.\d+)?)/i;
  const _401kRegex = /401k\s*([\d,]+(?:\.\d+)?)/i;

  const medicalCoverageMatch = extractedText.match(medicalCoverageRegex);
  const dentalCoverageMatch = extractedText.match(dentalCoverageRegex);
  const visionCoverageMatch = extractedText.match(visionCoverageRegex);
  const accidentInsPreTaxMatch = extractedText.match(accidentInsPreTaxRegex);
  const _401kMatch = extractedText.match(_401kRegex);
  return {
    medicalCoverage: medicalCoverageMatch ? medicalCoverageMatch[1] : "0",
    dentalCoverage: dentalCoverageMatch ? dentalCoverageMatch[1] : "0",
    visionCoverage: visionCoverageMatch ? visionCoverageMatch[1] : "0",
    accidentInsPreTax: accidentInsPreTaxMatch ? accidentInsPreTaxMatch[1] : "0",
    _401k: _401kMatch ? _401kMatch[1] : "0"
  };
}

function parsePDFTaxDeductionsData(extractedText: string) {
  const withholdingTaxRegex = /Withholding\s*Tax\s*([\d,]+(?:\.\d+)?)/i;
  const socialSecurityTaxRegex = /Social\s*Security\s*Tax\s*([\d,]+(?:\.\d+)?)/i;
  const medicareTaxRegex = /Medicare\s*Tax\s*([\d,]+(?:\.\d+)?)/i;
  
  const withholdingTaxMatch = extractedText.match(withholdingTaxRegex);
  const socialSecurityTaxMatch = extractedText.match(socialSecurityTaxRegex);
  const medicareTaxMatch = extractedText.match(medicareTaxRegex);
  return {
    withholdingTax: withholdingTaxMatch ? withholdingTaxMatch[1] : "0",
    socialSecurityTax: socialSecurityTaxMatch ? socialSecurityTaxMatch[1] : "0",
    medicareTax: medicareTaxMatch ? medicareTaxMatch[1] : "0"
  };
}

function parsePDFAfterTaxDeductionsData(extractedText: string) {
  const employeeLifeRegex = /Employee\s*Life\s*([\d,]+(?:\.\d+)?)/i;
  const detalDiscountPlanRegex = /Dental\s*Discount\s*Plan\s*([\d,]+(?:\.\d+)?)/i;
  const roth401kRegex = /Roth\s*401k\s*([\d,]+(?:\.\d+)?)/i;
  const pacAPARegex = /PAC\s*-\s*APA\s*([\d,]+(?:\.\d+)?)/i;
  const unionDuesRegex = /Union\s*Dues\s*-\s*APA\s*([\d,]+(?:\.\d+)?)/i;

  const employeeLifeMatch = extractedText.match(employeeLifeRegex);
  const dentalDiscountPlanMatch = extractedText.match(detalDiscountPlanRegex);
  const roth401kMatch = extractedText.match(roth401kRegex);
  const pacAPAMatch = extractedText.match(pacAPARegex);
  const unionDuesMatch = extractedText.match(unionDuesRegex);

  return {
    employeeLife: employeeLifeMatch ? employeeLifeMatch[1] : "0",
    dentalDiscountPlan: dentalDiscountPlanMatch ? dentalDiscountPlanMatch[1] : "0",
    roth401k: roth401kMatch ? roth401kMatch[1] : "0",
    pacAPA: pacAPAMatch ? pacAPAMatch[1] : "0",
    unionDues: unionDuesMatch ? unionDuesMatch[1] : "0"
  };
}

function parseCompanyContributionsData(extractedText: string) {
  const _401kCompanyContributionRegex = /401k\s*Company\s*Contrib\.+\s*([\d,]+(?:\.\d+)?)/i;
  const groupTermLifeRegex = /Group\s*Term\s*Life\s*([\d,]+(?:\.\d+)?)/i;

  const _401kCompanyContributionMatch = extractedText.match(_401kCompanyContributionRegex);
  const groupTermLifeMatch = extractedText.match(groupTermLifeRegex);

  return {
    _401kCompanyContribution: _401kCompanyContributionMatch ? _401kCompanyContributionMatch[1] : "0",
    groupTermLife: groupTermLifeMatch ? groupTermLifeMatch[1] : "0"
  };
}

function parsePDFTaxableEarningsData(extractedText: string) {
  const withHoldingTaxRegex =/Withholding\s*Tax\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)/g;
  const socialSecurityTaxRegex = /Social\s*Security\s*Tax\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)/g;
  const medicareTaxRegex = /Medicare\s*Tax\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)/g;

  const withHoldingTaxMatches = [...extractedText.matchAll(withHoldingTaxRegex)];
  const socialSecurityTaxMatches = [...extractedText.matchAll(socialSecurityTaxRegex)];
  const medicareTaxMatches = [...extractedText.matchAll(medicareTaxRegex)];
  const withHoldingTaxMatch = withHoldingTaxMatches[1] ? withHoldingTaxMatches[1][1] : "0";
  const socialSecurityTaxMatch = socialSecurityTaxMatches[1] ? socialSecurityTaxMatches[1][1] : "0";
  const medicareTaxMatch = medicareTaxMatches[1] ? medicareTaxMatches[1][1] : "0";

  return {
    withHoldingTaxEarnings: withHoldingTaxMatch,
    socialSecurityTaxEarnings: socialSecurityTaxMatch,
    medicareTaxEarnings: medicareTaxMatch
  };
}

// Function to convert a date string in the format "MM/DD/YYYY" to a Date object
function convertStringToDate(date: string) {
  const parts = date.split("/");
  if (parts.length === 3) {
    return new Date(`${parts[2]},${parts[0]},${parts[1]}`);
  } else {
    throw new Error("Invalid date format");
  }
}

// Function to remove non-numeric characters from a string like commas or negative signs
function removeNonNumericChars(str: string) {
  return String(str).replace(/[^0-9.-]/g, "");
}

// Function to extract text from a PDF buffer
async function extractTextFromPDF(fileBuffer: ArrayBuffer): Promise<string> {
  try {
    const pdfDocument = await pdfjsLib.getDocument({ data: fileBuffer }).promise;
    let text = '';
    const numPages = pdfDocument.numPages;

    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const page = await pdfDocument.getPage(pageNum);
      const content = await page.getTextContent();
      text += content.items.map((item: any) => item.str).join(' ');
    }

    return text;
  } catch (error) {
    console.error("Error extracting text from PDF:", error);
    throw new Error("Failed to extract text from PDF.");
  }
}
