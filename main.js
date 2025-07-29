import { app, BrowserWindow, ipcMain, dialog } from "electron";
import pkg from "exceljs";
import path from "node:path";
import { DateTime} from "luxon";
import fs from "fs";
import { WorkSheetGenerator } from './worksheet-gen.js';
import { Utilities } from "./utilities.js";
import { DeductTotalsWorksheet } from './deduct-tot-ws.js';

let window;
const { Workbook } = pkg;
const workSheetGenerator = new WorkSheetGenerator();
const util = new Utilities();
const deductTotalsWS = new DeductTotalsWorksheet();


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
ipcMain.handle("parse-pdf", async (event, pdfFile, outputPath, outputFileName ) => {

  
  try {
    if (!pdfFile) {
      throw new Error("PDF file buffer is undefined or null");
    }

    const extractedText = await util.extractTextFromPDF(pdfFile);

    let outputFilePath;
    let excelFile = null;
    if (outputPath) {
      outputFilePath = outputPath;
      //const outputFileName = path.basename(outputFilePath);
      const fileContent = fs.readFileSync(outputPath);
      excelFile = fileContent ? Promise.resolve(fileContent) : Promise.resolve(null);
    }
    else {
      outputFilePath = path.join(app.getPath('downloads'), `${outputFileName}.xlsx` ?? 'My_AA_Pay.xlsx');
    }

    // If Excel file buffer exists, process the existing file, else create a new file
    if (excelFile) {
      // If the Excel file exists, open and modify it
      const workbook = new Workbook();
      await workbook.xlsx.load(excelFile);

      // Add new worksheet and populate data
      await workSheetGenerator.addDataToWorkbook(workbook, extractedText);
      await deductTotalsWS.getAllDeductions(workbook);
      const sortedSheetNames = workbook.worksheets
        .map(ws => ws.name)
        //.filter(name => name !== "Totals")
        .sort((a, b) => {
          // Parse ddMMMyyyy to Date objects for comparison
          const parseDate = str => DateTime.fromFormat(str, "ddMMMyyyy");
          const dateA = parseDate(a);
          const dateB = parseDate(b);
          return dateA.toMillis() - dateB.toMillis();
        });

      sortedSheetNames.forEach((name, idx) => {
        workbook.worksheets.find(ws => ws.name === name).order = idx;
      });
      // Save the modified workbook to the same file path
      await workbook.xlsx.writeFile(outputFilePath);
    } else {
      const newWorkbook = new Workbook();
      
      // Add the first sheet with parsed data
      await workSheetGenerator.addDataToWorkbook(newWorkbook, extractedText);
      await deductTotalsWS.getAllDeductions(newWorkbook);
      // Save the new workbook
      await newWorkbook.xlsx.writeFile(outputFilePath);
    }

    return outputFilePath;  // Return the file path for the saved Excel file
  } catch (error) {
    throw new Error(`Failed to parse PDF file: ${error.message}`);
  }
});

