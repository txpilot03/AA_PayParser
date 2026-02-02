/**
 * EXAMPLE: How to integrate the Hybrid Parser with your existing code
 * 
 * This shows two approaches:
 * 1. Direct replacement - use hybrid parser instead of regex parser
 * 2. Gradual migration - use hybrid parser alongside existing parser
 */

import { app, BrowserWindow, ipcMain, dialog } from "electron";
import pkg from "exceljs";
import path from "node:path";
import { DateTime } from "luxon";
import fs from "fs";
import { WorkSheetGenerator } from '../worksheet-gen.js';
import { Utilities } from "../utilities.js";
import { DeductTotalsWorksheet } from '../deduct-tot-ws.js';
import { EarningsTotalsWorksheet } from '../earn-tot-ws.js';
import { HybridPDFParser } from '../parser.js';  // NEW IMPORT

let window;
const { Workbook } = pkg;
const workSheetGenerator = new WorkSheetGenerator();
const util = new Utilities();
const deductTotalsWS = new DeductTotalsWorksheet();
const earningTotalsWS = new EarningsTotalsWorksheet();
const hybridParser = new HybridPDFParser();  // NEW INSTANCE

/* ... createWindow and app setup code stays the same ... */

// ========================================================================
// APPROACH 1: Direct Replacement - Use Hybrid Parser Only
// ========================================================================
ipcMain.handle("parse-pdf", async (event, pdfFile, pdfFileBuffer, outputPath, outputFileName) => {
  console.log('PDF FILE: ', pdfFile);

  try {
    if (!pdfFileBuffer) {
      throw new Error("PDF file buffer is undefined or null");
    }

    // ✅ NEW: Use hybrid parser instead of extracting text and using regex
    const parsedData = await hybridParser.parse(pdfFileBuffer);
    
    // Convert to legacy format for compatibility with existing worksheet generator
    const legacyData = hybridParser.convertToLegacyFormat(parsedData);

    // Log the parsed data to see what we got
    console.log('Parsed Data:', JSON.stringify(parsedData, null, 2));
    console.log('Legacy Format:', legacyData);

    let outputFilePath;
    let excelFile = null;
    if (outputPath) {
      outputFilePath = outputPath;
      const fileContent = fs.readFileSync(outputPath);
      excelFile = fileContent ? Promise.resolve(fileContent) : Promise.resolve(null);
    } else {
      outputFilePath = path.join(app.getPath('downloads'), `${outputFileName}.xlsx` ?? 'My_AA_Pay.xlsx');
    }

    // Process workbook
    if (excelFile) {
      const workbook = new Workbook();
      await workbook.xlsx.load(excelFile);

      // ✅ MODIFIED: Pass parsed data to worksheet generator
      // You'll need to update WorkSheetGenerator to accept parsed object instead of raw text
      await workSheetGenerator.addDataToWorkbookFromParsed(workbook, legacyData);
      await deductTotalsWS.getAllDeductions(workbook);
      await earningTotalsWS.getAllEarnings(workbook);

      const sortedSheetNames = workbook.worksheets
        .map(ws => ws.name)
        .sort((a, b) => {
          const parseDate = str => DateTime.fromFormat(str, "ddMMMyyyy");
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

    return outputFilePath;
  } catch (error) {
    console.error('Error parsing PDF:', error);
    throw new Error(`Failed to parse PDF file: ${error.message}`);
  }
});


// ========================================================================
// APPROACH 2: Gradual Migration - Use Both Parsers for Comparison
// ========================================================================
ipcMain.handle("parse-pdf-with-comparison", async (event, pdfFile, pdfFileBuffer, outputPath, outputFileName) => {
  try {
    if (!pdfFileBuffer) {
      throw new Error("PDF file buffer is undefined or null");
    }

    // OLD WAY: Extract text and use regex
    const extractedText = await util.extractTextFromPDF(pdfFileBuffer);
    
    // NEW WAY: Use hybrid parser
    const parsedData = await hybridParser.parse(pdfFileBuffer);
    const legacyData = hybridParser.convertToLegacyFormat(parsedData);

    // COMPARISON: Log both results to see differences
    console.log('========================================');
    console.log('HYBRID PARSER RESULTS:');
    console.log('Pay Period:', legacyData.payPeriod);
    console.log('Seniority Year:', legacyData.seniorityYear);
    console.log('Operational Pay Rate:', legacyData.operationalPayRate);
    console.log('Operational Pay Hours:', legacyData.operationalPayHours);
    console.log('Operational Pay Current:', legacyData.operationalPayCurrent);
    console.log('Gross:', legacyData.gross);
    console.log('Net Pay:', legacyData.netPay);
    console.log('========================================');

    // For now, use hybrid parsed data
    let outputFilePath;
    let excelFile = null;
    if (outputPath) {
      outputFilePath = outputPath;
      const fileContent = fs.readFileSync(outputPath);
      excelFile = fileContent ? Promise.resolve(fileContent) : Promise.resolve(null);
    } else {
      outputFilePath = path.join(app.getPath('downloads'), `${outputFileName}.xlsx` ?? 'My_AA_Pay.xlsx');
    }

    if (excelFile) {
      const workbook = new Workbook();
      await workbook.xlsx.load(excelFile);
      await workSheetGenerator.addDataToWorkbookFromParsed(workbook, legacyData);
      await deductTotalsWS.getAllDeductions(workbook);
      await earningTotalsWS.getAllEarnings(workbook);

      const sortedSheetNames = workbook.worksheets
        .map(ws => ws.name)
        .sort((a, b) => {
          const parseDate = str => DateTime.fromFormat(str, "ddMMMyyyy");
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

    return outputFilePath;
  } catch (error) {
    console.error('Error parsing PDF:', error);
    throw new Error(`Failed to parse PDF file: ${error.message}`);
  }
});


// ========================================================================
// EXAMPLE: Simple test function to see hybrid parser output
// ========================================================================
async function testHybridParser(pdfFilePath) {
  try {
    // Read PDF file
    const pdfBuffer = fs.readFileSync(pdfFilePath);
    
    // Parse with hybrid parser
    const hybridParser = new HybridPDFParser();
    const parsedData = await hybridParser.parse(pdfBuffer);
    
    // Display results in a readable format
    console.log('\n=== HEADER INFO ===');
    console.log('Pay Period:', parsedData.header.payPeriod);
    console.log('Seniority Year:', parsedData.header.seniorityYear);
    console.log('Group:', parsedData.header.group);
    console.log('Hourly Rate:', parsedData.header.hourlyRate);
    
    console.log('\n=== EARNINGS ===');
    console.log('Operational Pay:', parsedData.earnings.operationalPay);
    console.log('Flight Training Pay:', parsedData.earnings.fltTrainingPay);
    console.log('Sick Pay:', parsedData.earnings.sickPay);
    
    console.log('\n=== DEDUCTIONS ===');
    console.log('Medical Coverage:', parsedData.deductions.preTax.medicalCoverage);
    console.log('401k:', parsedData.deductions.preTax._401k);
    console.log('Withholding Tax:', parsedData.deductions.taxes.withholdingTax);
    
    console.log('\n=== SUMMARY ===');
    console.log('Gross:', parsedData.summary.gross);
    console.log('Net Pay:', parsedData.summary.netPay);
    
    // Show ALL earnings that were found
    console.log('\n=== ALL EARNINGS DATA ===');
    for (const [key, value] of Object.entries(parsedData.earnings)) {
      if (value.current !== '0') {
        console.log(`${key}:`, value);
      }
    }
    
    return parsedData;
  } catch (error) {
    console.error('Test failed:', error);
    throw error;
  }
}

// To run test: testHybridParser('/path/to/your/paystub.pdf');
