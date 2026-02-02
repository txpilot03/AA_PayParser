import { app, BrowserWindow, ipcMain, dialog } from "electron";
import pkg from "exceljs";
import path, { parse } from "node:path";
import { DateTime} from "luxon";
import fs from "fs";
import { WorkSheetGenerator } from './worksheet-gen.js';
import { Utilities } from "./utilities.js";
import { DeductTotalsWorksheet } from './deduct-tot-ws.js';
import { EarningsTotalsWorksheet } from './earn-tot-ws.js';
import { HybridPDFParser } from './parser.js';
//import fs from "fs";
import PDFParser from "pdf2json";

let window;
const { Workbook } = pkg;
const hybridParser = new HybridPDFParser();

const workSheetGenerator = new WorkSheetGenerator();
const util = new Utilities();
const deductTotalsWS = new DeductTotalsWorksheet();
const earningTotalsWS = new EarningsTotalsWorksheet();


function createWindow() {
  window = new BrowserWindow({
    width: 800,
    height: 600,
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

ipcMain.handle("show-open-dialog", async () => {
  const result = await dialog.showOpenDialog(window, {
    title: "Open File",
    defaultPath: app.getPath("documents"),
    filters: [
      { name: "Excel Files", extensions: ["xlsx"] },
      { name: "All Files", extensions: ["*"] },
    ],
    properties: ["openFile"]
  });
  return result.filePaths && result.filePaths[0] ? result.filePaths[0] : null;
});

// IPC to handle the file reading
ipcMain.handle("parse-pdf", async (event, pdfFile, pdfFileBuffer, outputPath, outputFileName ) => {
  console.log('=== PDF PARSE DEBUG ===');
  console.log('Buffer type:', pdfFileBuffer.constructor.name);
  console.log('Is ArrayBuffer:', pdfFileBuffer instanceof ArrayBuffer);
  console.log('Is Buffer:', Buffer.isBuffer(pdfFileBuffer));
  console.log('Buffer length:', pdfFileBuffer.byteLength || pdfFileBuffer.length);
  console.log('=======================');

  try {
    if (!pdfFileBuffer) {
      throw new Error("PDF file buffer is undefined or null");
    }

    // Use hybrid parser to extract structured data
    const parsedData = await hybridParser.parse(pdfFileBuffer);
    const legacyData = hybridParser.convertToLegacyFormat(parsedData);
    console.log('Parsed Data: ', parsedData);

    let outputFilePath;
    let excelFile = null;
    if (outputPath) {
      outputFilePath = outputPath;
      const fileContent = fs.readFileSync(outputPath);
      excelFile = fileContent ? Promise.resolve(fileContent) : Promise.resolve(null);
    } else {
      outputFilePath = path.join(app.getPath('downloads'), `${outputFileName}.xlsx` ?? 'My_AA_Pay.xlsx');
    }

    // If Excel file buffer exists, process the existing file, else create a new file
    if (excelFile) {
      // If the Excel file exists, open and modify it
      const workbook = new Workbook();
      await workbook.xlsx.load(excelFile);

      // Add new worksheet and populate data using hybrid parser
      await workSheetGenerator.addDataToWorkbookFromParsed(workbook, legacyData);
      await deductTotalsWS.getAllDeductions(workbook);
      await earningTotalsWS.getAllEarnings(workbook);

      const sortedSheetNames = workbook.worksheets
        .map(ws => ws.name)
        .sort((a, b) => {
          // Parse ddMMMyyyy to Date objects for comparison
          const parseDate = str => DateTime.fromFormat(str, "yyyy-MM-dd", { zone: 'utc' });
          console.log('DATE A: ', parseDate(a), 'DATE B: ', parseDate(b));
          const dateA = parseDate(a);
          const dateB = parseDate(b);
          return dateA.toMillis() - dateB.toMillis();
        });

      sortedSheetNames.forEach((name, idx) => {
        workbook.worksheets.find(ws => ws.name === name).order = idx;
      });

      await workbook.xlsx.writeFile(outputFilePath);
    } else {
      const newWorkbook = new Workbook();
      await workSheetGenerator.addDataToWorkbookFromParsed(newWorkbook, legacyData);
      await deductTotalsWS.getAllDeductions(newWorkbook);
      await earningTotalsWS.getAllEarnings(newWorkbook);
      await newWorkbook.xlsx.writeFile(outputFilePath);
    }

    return outputFilePath;  // Return the file path for the saved Excel file
  } catch (error) {
    console.error('Error parsing PDF')
    throw new Error(`Failed to parse PDF file: ${error.message}`);
  }
});

// ========================================================================
// Test function to see hybrid parser output
// ========================================================================
async function testHybridParser(pdfFilePath) {
  try {
    // Read PDF file
    const pdfBuffer = fs.readFileSync(pdfFilePath);
    console.log('Hybrid Buffer');
    // Parse with hybrid parser
    const hybridParser = new HybridPDFParser();
    const parsedData = await hybridParser.parse(pdfBuffer);
    
    // Display results in a readable format
    // console.log('\n=== HEADER INFO ===');
    // console.log('Pay Period:', parsedData.header.payPeriod);
    // console.log('Seniority Year:', parsedData.header.seniorityYear);
    // console.log('Group:', parsedData.header.group);
    // console.log('Hourly Rate:', parsedData.header.hourlyRate);
    
    // console.log('\n=== EARNINGS ===');
    // console.log('Operational Pay:', parsedData.earnings.operationalPay);
    // console.log('Flight Training Pay:', parsedData.earnings.fltTrainingPay);
    // console.log('Sick Pay:', parsedData.earnings.sickPay);
    
    // console.log('\n=== DEDUCTIONS ===');
    // console.log('Medical Coverage:', parsedData.deductions.preTax.medicalCoverage);
    // console.log('401k:', parsedData.deductions.preTax._401k);
    // console.log('Withholding Tax:', parsedData.deductions.taxes.withholdingTax);
    
    // console.log('\n=== SUMMARY ===');
    // console.log('Gross:', parsedData.summary.gross);
    // console.log('Net Pay:', parsedData.summary.netPay);
    
    // // Show ALL earnings that were found
    // console.log('\n=== ALL EARNINGS DATA ===');
    // for (const [key, value] of Object.entries(parsedData.earnings)) {
    //   if (value.current !== '0') {
    //     console.log(`${key}:`, value);
    //   }
    // }
    
    return parsedData;
  } catch (error) {
    console.error('Test failed:', error);
    throw error;
  }
}

// To run test: testHybridParser('/path/to/your/paystub.pdf');
