import { TextParser } from "./parse-text.js";
import { DateTime} from "luxon";
import { Utilities } from "./utilities.js";


export class WorkSheetGenerator {

  util = new Utilities();
  textParser = new TextParser();

  /**
   * Add data to workbook from already-parsed data object
   * This bypasses all the regex parsing and uses structured data directly
   */
  async addDataToWorkbookFromParsed(workbook, parsedData) {
    const removeNonNumericChars = this.util.removeNonNumericChars;

    const worksheetName = this.util.convertStringToDate(parsedData.regularPayRoll);

    // Check if sheet already exists
    if (workbook.getWorksheet(worksheetName)) {
      throw new Error(`A worksheet with the name "${worksheetName}" already exists.`);
    }

    const worksheet = workbook.addWorksheet(worksheetName);

    // Style worksheet columns and cells
    for (let col of ["A", "B", "C", "D", "E", "F"]) {
      const column = worksheet.getColumn(col);
      if (col === "A") {
        column.width = 30;
      } else {
        column.width = 25;
      }
      column.alignment = { vertical: "middle", horizontal: "center" };
      //column.font = { name: "Arial", size: 12 };
    }

    // Set background color and font for specific cells
    for (let cell of [
      "A4", "B4", "C1", "C4", "D2", "D4", "E4", "F4", "A1", "A2", "B2", "C2",
      "B6", "C6", "D6", "A34", "B34", "C34", "D34", "E34", 
      "F34", "E36", "F36"
    ]) {
      worksheet.getCell(cell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD3D3D3" }
      };
      worksheet.getCell(cell).font = { bold: true };
    }

    //Green background:
    for (let cell of ["A6", "A7","B7", "C7", "D7"]) {
      worksheet.getCell(cell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF00FF00" }
      };
      worksheet.getCell(cell).font = { bold: true };
    }

    //Red background:
    for (let cell of ["E6", "E7", "F7", "E14", "F14", "E19", 
      "F19" 
    ]) {
      worksheet.getCell(cell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF0000" }
      };
      worksheet.getCell(cell).font = { bold: true, color: { argb: "FFFFFFFF" } };
    }

    // Set currency format for specific cells
    for (let cell of ["C3", "B5", "C5", "D5", "E5", "F5", "D8", "D9", "D10", "D11", "D12", 
      "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D20", "D21", "D22", "D23", "D24", 
      "D25", "D26", "D27", "D28", "D29", "D30", "D31", "D32", "D33", "F8", "F9", "F10", "F11", "F12",
      "F13", "F15", "F16", "F17", "F18", "F20", "F21", "F22", "F23", "F24", "F25", "F26", 
      "F27", "F28", "F29", "F30", "F31", "D35", "D36", "D37", "F35", "F37"
    ]) {
      worksheet.getCell(cell).numFmt = "$#,##0.00";
    }

    // Set left alignment for label cells
    for (let cell of [
      "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A15", 
      "A16", "A17", "A18", "A19", "A20", "A21", "A22", "A23", 
      "A24", "A25", "A26", "A27", "A28", "A29", "A30", "A31",
      "A32", "A33", "E8", "E9", "E10", "E11", "E12", "E13", "E15", "E16",
      "E17", "E18", "E20", "E21", "E22", "E23", "E24", "E25",
      "E26", "E27", "E28", "E29", "E30", "E31", "A35", "A36",
      "A37", "E35", "E37"
    ]) {
      worksheet.getCell(cell).alignment = {
        vertical: "middle",
        horizontal: "left",
      };
      worksheet.getCell(cell).font = { bold: true };
    }

    // Set right alignment for numeric cells
    for (let cell of [
      "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18",
      "D19", "D20", "D21", "D22", "D23", "D24", "D25", "D26", "D27", "D28", "D29",
      "D30", "D31", "D32", "D33", "F8", "F9", "F10", "F11", "F12", "F13", "F15", "F16", "F17",
      "F18", "F20", "F21", "F22", "F23", "F24", "F25", "F26", "F27", "F28", "F29",
      "F30", "F31"
    ]) {
      worksheet.getCell(cell).alignment = {
        vertical: "middle",
        horizontal: "right",
      };
    }

    // Set border for additional information section
    for (let rowNum of [34, 35, 36, 37]) {
      for (let col of ["A", "B", "C", "D", "E", "F"]) {
        worksheet.getCell(`${col}${rowNum}`).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      }
    }

    // set top border for earnings subtotal rows
    for (let col of ["A", "B", "C", "D"]) {
      worksheet.getCell(`${col}16`).border = {
        top: { style: "medium" }
      };
    }

    for (let col of ["A", "B", "C", "D"]) {
      worksheet.getCell(`${col}31`).border = {
        top: { style: "medium" }
      };
    }

    // Set border for Earnings and deductions section headers
    for (let rowNum of [6, 7, 14, 19]) {
      if (rowNum === 14 || rowNum === 19) {
        for (let col of ["A", "E", "F"]) {
          if (col === "A") {
            worksheet.getCell(`${col}${rowNum}`).border = {
              left: { style: "thin" },
            };
          } else {
            worksheet.getCell(`${col}${rowNum}`).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" }
            };
          };
        } 
      } else {
        for (let col of ["A", "B", "C", "D", "E", "F"]) {
          worksheet.getCell(`${col}${rowNum}`).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
          };
        };
      }
    }

    for (let rowNum of [8,9,10,11,12,13,15,16,17,18,20,21,22,23,24,25,26,27,28,29,30,31,32,33]) {
      for (let col of ["A", "E", "F"]) {
        if (col === "E" || col === "A") {
          worksheet.getCell(`${col}${rowNum}`).border = {
            left: { style: "thin" },
          };
        } else {
          worksheet.getCell(`${col}${rowNum}`).border = {
            right: { style: "thin" },
          };
        }
      }
    }


    // ========================================================================
    // HEADER & SUMMARY INFORMATION
    // ========================================================================

    // Header Information
    worksheet.getCell("A1").value = "Regular Pay";
    worksheet.getCell("B1").value = parsedData.regularPayRoll;
    worksheet.getCell("C1").value = "Pay Period";
    worksheet.getCell("D1").value = parsedData.payPeriod;
    
    worksheet.getCell("A2").value = "Seniority Year";
    worksheet.getCell("A3").value = Number(parsedData.seniorityYear);
    worksheet.getCell("A3").numFmt = "0";

    worksheet.getCell("B2").value = "Group"
    worksheet.getCell("B3").value = parsedData.group;
    
    worksheet.getCell("C2").value = "Hourly Rate";
    worksheet.getCell("C3").value = Number(
      removeNonNumericChars(parsedData.hourlyRate)
    );
    

    worksheet.getCell("D2").value = "Effective Period"
    worksheet.getCell("D3").value = parsedData.effectivePeriod;

    // Earnings Header
    worksheet.getCell("A4").value = "Summary";
    worksheet.getCell("A5").value = "Current";
    worksheet.getCell("B4").value = "Gross";
    worksheet.getCell("B5").value = Number(
      removeNonNumericChars(parsedData.gross));
    worksheet.getCell("C4").value = "Pre-Tax Deduct";
    worksheet.getCell("C5").value = Number(
      removeNonNumericChars(parsedData.preTaxDeduct));
    worksheet.getCell("D4").value = "Taxes";
    worksheet.getCell("D5").value = Number(
      removeNonNumericChars(parsedData.taxes));
    worksheet.getCell("E4").value = "After-Tax Deduct";
    worksheet.getCell("E5").value = Number(
      removeNonNumericChars(parsedData.afterTaxDeduct));
    worksheet.getCell("F4").value = "Net Pay";
    worksheet.getCell("F5").value = Number(
      removeNonNumericChars(parsedData.netPay));

    worksheet.mergeCells("A6:D6");
    worksheet.mergeCells("E6:F6");
    worksheet.getCell("A6").value = "EARNINGS";
    worksheet.getCell("E6").value = "DEDUCTIONS";


    // ========================================================================
    // EARNINGS SECTION PART 1
    // ========================================================================

    worksheet.getCell("A7").value = "Earnings";
    worksheet.getCell("B7").value = "Rate";
    worksheet.getCell("C7").value = "Hours";
    worksheet.getCell("D7").value = "Current";
    worksheet.getCell("E7").value = "Pre-Tax Deduct";
    worksheet.getCell("F7").value = "Current";

    // ✅ Populate Earnings Data - Direct access, no regex needed!
    let row = 8;
    
    // Crew Advance
    if (parsedData.crewAdvanceCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Crew Advance";
      worksheet.getCell(`B${row}`).value = Number(parsedData.crewAdvanceRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.crewAdvanceHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.crewAdvanceCurrent));
      row++;
    }

    // Pilot Exp D Taxable
    if (parsedData.pltExpDtaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp D Taxable";
      worksheet.getCell(`B${row}`).value = Number(parsedData.pltExpDtaxRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.pltExpDtaxHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.pltExpDtaxCurrent));
      row++;
    }

    // Pilot Exp ADJ Taxable
    if (parsedData.pltExpADJtaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp ADJ Taxable";
      worksheet.getCell(`B${row}`).value = Number(parsedData.pltExpADJtaxRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.pltExpADJtaxHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.pltExpADJtaxCurrent));
      row++;
    }

    // Pilot Exp I Taxable
    if (parsedData.pltExpItaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp I Taxable";
      worksheet.getCell(`B${row}`).value = Number(parsedData.pltExpItaxRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.pltExpItaxHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.pltExpItaxCurrent));
      row++;
    }

    // AAG Profit Sharing
    if (parsedData.AagProfitSharingCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "AAG Profit Sharing";
      worksheet.getCell(`B${row}`).value = Number(parsedData.AagProfitSharingRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.AagProfitSharingHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.AagProfitSharingCurrent));
      row++;
    }

    // Pilot Exp D Non-Taxable
    if (parsedData.pltExpDnonTaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp D Non-Taxable";
      worksheet.getCell(`B${row}`).value = Number(parsedData.pltExpDnonTaxRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.pltExpDnonTaxHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.pltExpDnonTaxCurrent));
      row++;
    }

    // Pilot Exp ADJ Non-Taxable
    if (parsedData.pltExpAdjnonTaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp ADJ Non-Taxable";
      worksheet.getCell(`B${row}`).value = Number(parsedData.pltExpAdjnonTaxRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.pltExpAdjnonTaxHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.pltExpAdjnonTaxCurrent));
      row++;
    }

    // Pilot Exp I Non-Taxable
    if (parsedData.pltExpInonTaxCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pilot Exp I Non-Taxable";
      worksheet.getCell(`B${row}`).value = Number(parsedData.pltExpInonTaxRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.pltExpInonTaxHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.pltExpInonTaxCurrent));
      row++;
    }

    row = 16;
    // Earnings Part 1 Total
    worksheet.getCell(`A${row}`).value = "Earnings Subtotal";
    worksheet.getCell(`A${row}`).font = { bold: true };
    worksheet.getCell(`D${row}`).value = { formula: `SUM(D8:D15)`, result: 0 };
    worksheet.getCell(`D${row}`).font = { bold: true };



    // ========================================================================
    // EARNINGS SECTION PART 2
    // ========================================================================

    row = 18; // Move to row 18 for next earnings

    // Operational Pay
    if (parsedData.operationalPayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Operational Pay";
      worksheet.getCell(`B${row}`).value = Number(parsedData.operationalPayRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.operationalPayHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.operationalPayCurrent));
      row++;
    }

    // Flight Training Pay
    if (parsedData.fltTrainingPayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Flight Training Pay";
      worksheet.getCell(`B${row}`).value = Number(parsedData.fltTrainingPayRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.fltTrainingPayHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.fltTrainingPayCurrent));
      row++;
    }

    // Sit Time
    if (parsedData.sitTimeCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Sit Time";
      worksheet.getCell(`B${row}`).value = Number(parsedData.sitTimeRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.sitTimeHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.sitTimeCurrent));
      row++;
    }

    // Pay Above Guarantee (RSV)
    if (parsedData.payAbvGuaranteeRsvCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Pay Above Guarantee (RSV)";
      worksheet.getCell(`B${row}`).value = Number(parsedData.payAbvGuaranteeRsvRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.payAbvGuaranteeRsvHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.payAbvGuaranteeRsvCurrent));
      row++;
    }

    // RA Prem
    if (parsedData.raPremCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "RA Prem";
      worksheet.getCell(`B${row}`).value = Number(parsedData.raPremRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.raPremHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.raPremCurrent));
      row++;
    }

    // Min Guarantee Adj
    if (parsedData.minGuaranteeAdjCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Min Guarantee Adj";
      worksheet.getCell(`B${row}`).value = Number(parsedData.minGuaranteeAdjRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.minGuaranteeAdjHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.minGuaranteeAdjCurrent));
      row++;
    }

    // Intl Override
    if (parsedData.intlOverrideCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Intl Override";
      worksheet.getCell(`B${row}`).value = Number(parsedData.intlOverrideRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.intlOverrideHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.intlOverrideCurrent));
      row++;
    }

    // Distance Learning
    if (parsedData.distanceLearningCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Distance Learning";
      worksheet.getCell(`B${row}`).value = Number(parsedData.distanceLearningRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.distanceLearningHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.distanceLearningCurrent));
      row++;
    }

    // Union PD Union Leave
    if (parsedData.unionPdLeaveCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Union PD Union Leave";
      worksheet.getCell(`B${row}`).value = Number(parsedData.unionPdLeaveRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.unionPdLeaveHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.unionPdLeaveCurrent));
      row++;
    }

    // Prem Incentive Pay
    if (parsedData.premIncentivePayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Prem Incentive Pay";
      worksheet.getCell(`B${row}`).value = Number(parsedData.premIncentivePayRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.premIncentivePayHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.premIncentivePayCurrent));
      row++;
    }

    // Flight Vacation Pay
    if (parsedData.fltVacationPayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Flight Vacation Pay";
      worksheet.getCell(`B${row}`).value = Number(parsedData.fltVacationPayRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.fltVacationPayHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.fltVacationPayCurrent));
      row++;
    }

    // Prior Year Vacation Pay Out
    if (parsedData.priorYearVacPayoutCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Prior Year Vacation Pay Out";
      worksheet.getCell(`B${row}`).value = Number(parsedData.priorYearVacPayoutRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.priorYearVacPayoutHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.priorYearVacPayoutCurrent));
      row++;
    }

    // Sick Pay
    if (parsedData.sickPayCurrent !== "0") {
      worksheet.getCell(`A${row}`).value = "Sick Pay";
      worksheet.getCell(`B${row}`).value = Number(parsedData.sickPayRate);
      worksheet.getCell(`B${row}`).numFmt = "0.00";
      worksheet.getCell(`C${row}`).value = Number(parsedData.sickPayHours);
      worksheet.getCell(`C${row}`).numFmt = "0.00";
      worksheet.getCell(`D${row}`).value = Number(
        removeNonNumericChars(parsedData.sickPayCurrent));
      row++;
    }

    row = 31;
     // Earnings Part 2 Total
    worksheet.getCell(`A${row}`).value = "Earnings Subtotal";
    worksheet.getCell(`A${row}`).font = { bold: true };
    worksheet.getCell(`C${row}`).value = { formula: `SUM(C18:C${row - 1})`, result: 0 };
    worksheet.getCell(`C${row}`).numFmt = "0.00";
    worksheet.getCell(`C${row}`).font = { bold: true };
    worksheet.getCell(`D${row}`).value = { formula: `SUM(D18:D${row - 1})`, result: 0 };
    worksheet.getCell(`D${row}`).font = { bold: true };

    worksheet.getCell(`A${row + 2}`).value = "Earnings Total";
    worksheet.getCell(`A${row + 2}`).font = { bold: true };
    worksheet.getCell(`D${row + 2}`).value = Number(
      removeNonNumericChars(parsedData.earningsTotalCurrent));
    worksheet.getCell(`D${row + 2}`).font = { bold: true };

    // ========================================================================
    // DEDUCTIONS SECTION
    // ========================================================================
    
    row = 8; // Start deductions at row 8
    
    worksheet.getCell("E14").value = "Taxes";
    worksheet.getCell("F14").value = "Current";

    worksheet.getCell("E19").value = "After-Tax Deduct";
    worksheet.getCell("F19").value = "Current";


    // Pre-Tax Deductions
    if (parsedData.medicalCoverage !== "0") {
      worksheet.getCell(`E${row}`).value = "Medical Coverage";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.medicalCoverage));
      row++;
    }

    if (parsedData.dentalCoverage !== "0") {
      worksheet.getCell(`E${row}`).value = "Dental Coverage";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.dentalCoverage));
      row++;
    }

    if (parsedData.visionCoverage !== "0") {
      worksheet.getCell(`E${row}`).value = "Vision Coverage";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.visionCoverage));
      row++;
    }

    if (parsedData.accidentInsPreTax !== "0") {
      worksheet.getCell(`E${row}`).value = "Accident Ins - Pre-Tax";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.accidentInsPreTax));
      row++;
    }

    if (parsedData._401k !== "0") {
      worksheet.getCell(`E${row}`).value = "401k";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData._401k));
      row++;
    }

    // Pre-Tax Deduction Total
    worksheet.getCell(`E${row}`).value = "Total";
    worksheet.getCell(`E${row}`).font = { bold: true };
    worksheet.getCell(`F${row}`).value = { formula: `SUM(F8:F${row - 1})`, result: 0 };
    worksheet.getCell(`F${row}`).font = { bold: true };

    row = 15; // Move to row 15 for taxes

    // Taxes
    if (parsedData.withholdingTax !== "0") {
      worksheet.getCell(`E${row}`).value = "Withholding Tax";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.withholdingTax));
      row++;
    }

    if (parsedData.socialSecurityTax !== "0") {
      worksheet.getCell(`E${row}`).value = "Social Security Tax";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.socialSecurityTax));
      row++;
    }

    if (parsedData.medicareTax !== "0") {
      worksheet.getCell(`E${row}`).value = "Medicare Tax";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.medicareTax));
      row++;
    }

    // Tax Total
    worksheet.getCell(`E${row}`).value = "Total";
    worksheet.getCell(`E${row}`).font = { bold: true };
    worksheet.getCell(`F${row}`).value = { formula: `SUM(F15:F${row - 1})`, result: 0 };
    worksheet.getCell(`F${row}`).font = { bold: true };

    row = 20; // Move to row 20 for after-tax deductions

    // After-Tax Deductions
    if (parsedData.employeeLife !== "0") {
      worksheet.getCell(`E${row}`).value = "Employee Life";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.employeeLife));
      row++;
    }

    if (parsedData.dentalDiscountPlan !== "0") {
      worksheet.getCell(`E${row}`).value = "Dental Discount Plan";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.dentalDiscountPlan));
      row++;
    }

    
    if (parsedData.pacAPA !== "0") {
      worksheet.getCell(`E${row}`).value = "PAC - APA";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.pacAPA));
      row++;
    }
      
    if (parsedData.unionDues !== "0") {
      worksheet.getCell(`E${row}`).value = "Union Dues - APA";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.unionDues));
      row++;
    }
      
    if (parsedData.roth401k !== "0") {
      worksheet.getCell(`E${row}`).value = "Roth 401k";
      worksheet.getCell(`F${row}`).value = Number(
        removeNonNumericChars(parsedData.roth401k));
      row++;
    }
    // After-Tax Deduction Total
    worksheet.getCell(`E${row}`).value = "Total";
    worksheet.getCell(`E${row}`).font = { bold: true };
    worksheet.getCell(`F${row}`).value = { formula: `SUM(F20:F${row - 1})`, result: 0 };
    worksheet.getCell(`F${row}`).font = { bold: true };

    // ========================================================================
    // Additional Info Section
    // ========================================================================
    
    worksheet.getCell("A34").value = "Taxable Earnings";
    worksheet.getCell("D34").value = "Current";

    worksheet.getCell("A35").value = "Withholding Tax";
    worksheet.mergeCells("A35:C35");
    worksheet.getCell("D35").value = Number(
      removeNonNumericChars(parsedData.withHoldingTaxEarnings));

    worksheet.getCell("A36").value = "Social Security Tax";
    worksheet.mergeCells("A36:C36");
    worksheet.getCell("D36").value = Number(
      removeNonNumericChars(parsedData.socialSecurityTaxEarnings));

    worksheet.getCell("A37").value = "Medicare Tax";
    worksheet.mergeCells("A37:C37");
    worksheet.getCell("D37").value = Number(
      removeNonNumericChars(parsedData.medicareTaxEarnings));

    worksheet.getCell("E34").value = "Additional Info";
    worksheet.getCell("F34").value = "Current";

    worksheet.getCell("E35").value = "401k Co. Contrib";
    worksheet.getCell("F35").value = Number(
      removeNonNumericChars(parsedData._401kCompanyContribution));
    

    worksheet.getCell("E36").value = "Imputed Income";
    worksheet.getCell("F36").value = "Current";
    worksheet.getCell("E37").value = "Group Term Life";
    worksheet.getCell("F37").value = Number(
      removeNonNumericChars(parsedData.groupTermLife));


    return worksheet;
  }


  // async addDataToWorkbook(workbook, extractedText) {
  //   const removeNonNumericChars = this.util.removeNonNumericChars;

  //   const infoData = this.textParser.parsePDFInfoData(extractedText);
  //   const summaryData = this.textParser.parsePDFSummaryData(extractedText);
  //   const earningsTopData =
  //     this.textParser.parsePDFEarningsTopData(extractedText);
  //   const earningsBottomData =
  //     this.textParser.parsePDFEarningsBottomData(extractedText);
  //   const preTaxDeductionsData =
  //     this.textParser.parsePDFPreTaxDeductionsData(extractedText);
  //   const taxDeductionsData =
  //     this.textParser.parsePDFTaxDeductionsData(extractedText);
  //   const afterTaxDeductionsData =
  //     this.textParser.parsePDFAfterTaxDeductionsData(extractedText);
  //   const companyContributionsData =
  //     this.textParser.parseCompanyContributionsData(extractedText);
  //   const taxableEarningsData =
  //     this.textParser.parsePDFTaxableEarningsData(extractedText);

  //   const date = this.util.convertStringToDate(infoData.payPeriod);
  //   const worksheetName = DateTime.fromISO(date.toISOString()).toFormat(
  //     "ddMMMyyyy"
  //   );

  //   // Check if the sheet already exists
  //   if (workbook.getWorksheet(worksheetName)) {
  //     throw new Error(
  //       `A worksheet with the name "${worksheetName}" already exists.`
  //     );
  //   }

  //   const worksheet = workbook.addWorksheet(worksheetName);

  //   // Style worksheet columns and cells
  //   for (let col of ["A", "B", "C", "D", "E", "F"]) {
  //     const column = worksheet.getColumn(col);
  //     if (col === "A") {
  //       column.width = 30; // Set width for column A
  //     } else {
  //       column.width = 25; // Set width for columns B to F
  //     }
  //     column.alignment = { vertical: "middle", horizontal: "center" }; // Center align text
  //     column.font = { name: "Arial", size: 12 }; // Set font for
  //   }

  //   // Set background color and font for specific cells
  //   for (let cell of [
  //     "A4",
  //     "B4",
  //     "C4",
  //     "D4",
  //     "E4",
  //     "A1",
  //     "A2",
  //     "B2",
  //     "C2",
  //     "A6",
  //     "B6",
  //     "C6",
  //     "D6",
  //     "A7",
  //     "B7",
  //     "C7",
  //     "D7",
  //     "E6",
  //     "E7",
  //     "F7",
  //     "E14",
  //     "F14",
  //     "E19",
  //     "F19",
  //     "A32",
  //     "B32",
  //     "C32",
  //     "D32",
  //     "E32",
  //     "F32",
  //     "A34",
  //     "B34",
  //     "C34",
  //     "D34",
  //     "E34",
  //     "F34",
  //     "E36",
  //     "F36",
  //   ]) {
  //     worksheet.getCell(cell).fill = {
  //       type: "pattern",
  //       pattern: "solid",
  //       fgColor: { argb: "FFBFBFBF" },
  //     };
  //     worksheet.getCell(cell).font = {
  //       bold: true,
  //     };
  //   }

  //   for (let cell of [
  //     "A6",
  //     "A7",
  //     "B7",
  //     "C7",
  //     "D7",
  //     "E6",
  //     "E7",
  //     "F7",
  //     "E14",
  //     "F14",
  //     "E19",
  //     "F19",
  //     "A32",
  //     "B32",
  //     "C32",
  //     "D32",
  //     "E32",
  //     "F32",
  //     "A34",
  //     "D34",
  //     "E34",
  //     "F34",
  //     "E36",
  //     "F36",
  //   ]) {
  //     worksheet.getCell(cell).border = {
  //       top: { style: "thin" },
  //       left: { style: "thin" },
  //       bottom: { style: "thin" },
  //       right: { style: "thin" },
  //     };
  //   }

  //   for (let cell of [
  //     "A8",
  //     "A9",
  //     "A10",
  //     "A11",
  //     "A12",
  //     "A13",
  //     "A14",
  //     "A15",
  //     "A16",
  //     "A17",
  //     "A18",
  //     "A19",
  //     "A20",
  //     "A21",
  //     "A22",
  //     "A23",
  //     "A24",
  //     "A25",
  //     "A26",
  //     "A27",
  //     "A28",
  //     "A29",
  //     "A30",
  //     "A31",
  //     "A35",
  //     "A36",
  //     "A37",
  //     "E35",
  //     "E37",
  //   ]) {
  //     worksheet.getCell(cell).alignment = {
  //       vertical: "middle",
  //       horizontal: "left",
  //     };
  //   }

  //   // Fill in the worksheet with the extracted data
  //   //? Header row
  //   worksheet.getCell("A1").value = "Pay Period";
  //   worksheet.getCell("B1").value = date;
  //   worksheet.getCell("B1").numFmt = "MMM dd, yyyy";

  //   worksheet.getCell("A2").value = "Seniority Year";
  //   worksheet.getCell("A3").value = Number(infoData.seniorityYear);

  //   worksheet.getCell("B2").value = "Group";
  //   worksheet.getCell("B3").value = infoData.group;

  //   worksheet.getCell("C2").value = "Hourly Rate";
  //   worksheet.getCell("C3").value = Number(
  //     removeNonNumericChars(infoData.hourlyRate)
  //   );
  //   worksheet.getCell("C3").numFmt = "$#,##0.00";

  //   //? Summary
  //   worksheet.getCell("A4").value = "Gross Earnings";
  //   worksheet.getCell("B4").value = "Pre-Tax Deduction";
  //   worksheet.getCell("C4").value = "Taxes";
  //   worksheet.getCell("D4").value = "After-Tax Deduction";
  //   worksheet.getCell("E4").value = "Net Pay";

  //   worksheet.getCell("A5").value = Number(
  //     removeNonNumericChars(summaryData.gross)
  //   );
  //   worksheet.getCell("A5").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("B5").value = Number(
  //     removeNonNumericChars(summaryData.preTaxDeduct)
  //   );
  //   worksheet.getCell("B5").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("C5").value = Number(
  //     removeNonNumericChars(summaryData.taxes)
  //   );
  //   worksheet.getCell("C5").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("D5").value = Number(
  //     removeNonNumericChars(summaryData.afterTaxDeduct)
  //   );
  //   worksheet.getCell("D5").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("E5").value = Number(
  //     removeNonNumericChars(summaryData.netPay)
  //   );
  //   worksheet.getCell("E5").numFmt = "$#,##0.00"; // Format as currency

  //   //? EARNINGS
  //   worksheet.getCell("A6").value = "EARNINGS";
  //   worksheet.mergeCells("A6:D6");

  //   //* Earnings Top
  //   // Header row
  //   worksheet.getCell("A7").value = "Earnings";
  //   worksheet.getCell("B7").value = "Rate";
  //   worksheet.getCell("C7").value = "Hours";
  //   worksheet.getCell("D7").value = "Current";

  //   worksheet.getCell("A8").value = "Crew Advance";
  //   worksheet.getCell("B8").value = Number(earningsTopData.crewAdvanceRate);
  //   worksheet.getCell("C8").value = Number(earningsTopData.crewAdvanceHours);
  //   worksheet.getCell("D8").value = Number(
  //     removeNonNumericChars(earningsTopData.crewAdvanceCurrent)
  //   );
  //   worksheet.getCell("D8").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A9").value = "PLT EXP D Taxable";
  //   worksheet.getCell("B9").value = Number(earningsTopData.pltExpDtaxRate);
  //   worksheet.getCell("C9").value = Number(earningsTopData.pltExpDtaxHours);
  //   worksheet.getCell("D9").value = Number(
  //     removeNonNumericChars(earningsTopData.pltExpDtaxCurrent)
  //   );
  //   worksheet.getCell("D9").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A10").value = "PLT EXP ADJ Taxable";
  //   worksheet.getCell("B10").value = Number(earningsTopData.pltExpADJtaxRate);
  //   worksheet.getCell("C10").value = Number(earningsTopData.pltExpADJtaxHours);
  //   worksheet.getCell("D10").value = Number(
  //     removeNonNumericChars(earningsTopData.pltExpADJtaxCurrent)
  //   );
  //   worksheet.getCell("D10").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A11").value = "PLT EXP I Taxable";
  //   worksheet.getCell("B11").value = Number(earningsTopData.pltExpItaxRate);
  //   worksheet.getCell("C11").value = Number(earningsTopData.pltExpItaxHours);
  //   worksheet.getCell("D11").value = Number(
  //     removeNonNumericChars(earningsTopData.pltExpItaxCurrent)
  //   );
  //   worksheet.getCell("D11").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A12").value = "AAG Profit Sharing";
  //   worksheet.getCell("B12").value = Number(
  //     earningsTopData.AagProfitSharingRate
  //   );
  //   worksheet.getCell("C12").value = Number(
  //     earningsTopData.AagProfitSharingHours
  //   );
  //   worksheet.getCell("D12").value = Number(
  //     removeNonNumericChars(earningsTopData.AagProfitSharingCurrent)
  //   );
  //   worksheet.getCell("D12").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A13").value = "PLT EXP D Non-Taxable";
  //   worksheet.getCell("B13").value = Number(earningsTopData.pltExpDnonTaxRate);
  //   worksheet.getCell("C13").value = Number(earningsTopData.pltExpDnonTaxHours);
  //   worksheet.getCell("D13").value = Number(
  //     removeNonNumericChars(earningsTopData.pltExpDnonTaxCurrent)
  //   );
  //   worksheet.getCell("D13").numFmt = "$#,##0.00"; // Format as currency
    
  //   worksheet.getCell("A14").value = "PLT EXP ADJ Non-Taxable";
  //   worksheet.getCell("B14").value = Number(
  //     earningsTopData.pltExpAdjnonTaxRate
  //   );
  //   worksheet.getCell("C14").value = Number(
  //     earningsTopData.pltExpAdjnonTaxHours
  //   );
  //   worksheet.getCell("D14").value = Number(
  //     removeNonNumericChars(earningsTopData.pltExpAdjnonTaxCurrent)
  //   );
  //   worksheet.getCell("D14").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A15").value = "PLT EXP I Non-Taxable";
  //   worksheet.getCell("B15").value = Number(earningsTopData.pltExpInonTaxRate);
  //   worksheet.getCell("C15").value = Number(earningsTopData.pltExpInonTaxHours);
  //   worksheet.getCell("D15").value = Number(
  //     removeNonNumericChars(earningsTopData.pltExpInonTaxCurrent)
  //   );
  //   worksheet.getCell("D15").numFmt = "$#,##0.00"; // Format as currency

  //   worksheet.getCell("A16").value = "Earnings Total";
  //   worksheet.getCell("A16").font = { bold: true };
  //   worksheet.getCell("D16").value = { formula: "SUM(D8:D15)", result: 7 };
  //   worksheet.getCell("D16").font = { bold: true };
  //   worksheet.getCell("D16").numFmt = "$#,##0.00"; // Format as currency

  //   //* Earnings Bottom
  //   worksheet.getCell("A18").value = "Operational Pay";
  //   worksheet.getCell("B18").value = Number(
  //     earningsBottomData.operationalPayRate
  //   );
  //   worksheet.getCell("C18").value = Number(
  //     earningsBottomData.operationalPayHours
  //   );
  //   worksheet.getCell("D18").value = Number(
  //     removeNonNumericChars(earningsBottomData.operationalPayCurrent)
  //   );
  //   worksheet.getCell("D18").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A19").value = "Flight Training Pay";
  //   worksheet.getCell("B19").value = Number(
  //     earningsBottomData.fltTrainingPayRate
  //   );
  //   worksheet.getCell("C19").value = Number(
  //     earningsBottomData.fltTrainingPayHours
  //   );
  //   worksheet.getCell("D19").value = Number(
  //     removeNonNumericChars(earningsBottomData.fltTrainingPayCurrent)
  //   );
  //   worksheet.getCell("D19").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A20").value = "Sit Time";
  //   worksheet.getCell("B20").value = Number(earningsBottomData.sitTimeRate);
  //   worksheet.getCell("C20").value = Number(earningsBottomData.sitTimeHours);
  //   worksheet.getCell("D20").value = Number(
  //     removeNonNumericChars(earningsBottomData.sitTimeCurrent)
  //   );
  //   worksheet.getCell("D20").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A21").value = "Pay Above Guarantee (RSV)";
  //   worksheet.getCell("B21").value = Number(
  //     earningsBottomData.payAbvGuaranteeRsvRate
  //   );
  //   worksheet.getCell("C21").value = Number(
  //     earningsBottomData.payAbvGuaranteeRsvHours
  //   );
  //   worksheet.getCell("D21").value = Number(
  //     removeNonNumericChars(earningsBottomData.payAbvGuaranteeRsvCurrent)
  //   );
  //   worksheet.getCell("D21").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A22").value = "RA Prem";
  //   worksheet.getCell("B22").value = Number(earningsBottomData.raPremRate);
  //   worksheet.getCell("C22").value = Number(earningsBottomData.raPremHours);
  //   worksheet.getCell("D22").value = Number(
  //     removeNonNumericChars(earningsBottomData.raPremCurrent)
  //   );
  //   worksheet.getCell("D22").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A23").value = "Min Guarantee Adj";
  //   worksheet.getCell("B23").value = Number(
  //     earningsBottomData.minGuaranteeAdjRate
  //   );
  //   worksheet.getCell("C23").value = Number(
  //     earningsBottomData.minGuaranteeAdjHours
  //   );
  //   worksheet.getCell("D23").value = Number(
  //     removeNonNumericChars(earningsBottomData.minGuaranteeAdjCurrent)
  //   );
  //   worksheet.getCell("D23").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A24").value = "Intl Override";
  //   worksheet.getCell("B24").value = Number(
  //     earningsBottomData.intlOverrideRate
  //   );
  //   worksheet.getCell("C24").value = Number(
  //     earningsBottomData.intlOverrideHours
  //   );
  //   worksheet.getCell("D24").value = Number(
  //     removeNonNumericChars(earningsBottomData.intlOverrideCurrent)
  //   );
  //   worksheet.getCell("D24").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A25").value = "Distance Learning";
  //   worksheet.getCell("B25").value = Number(
  //     earningsBottomData.distanceLearningRate
  //   );
  //   worksheet.getCell("C25").value = Number(
  //     earningsBottomData.distanceLearningHours
  //   );
  //   worksheet.getCell("D25").value = Number(
  //     removeNonNumericChars(earningsBottomData.distanceLearningCurrent)
  //   );
  //   worksheet.getCell("D25").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A26").value = "Union PD Union Leave";
  //   worksheet.getCell("B26").value = Number(
  //     earningsBottomData.unionPdLeaveRate
  //   );
  //   worksheet.getCell("C26").value = Number(
  //     earningsBottomData.unionPdLeaveHours
  //   );
  //   worksheet.getCell("D26").value = Number(
  //     removeNonNumericChars(earningsBottomData.unionPdLeaveCurrent)
  //   );
  //   worksheet.getCell("D26").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A27").value = "Prem Incentive Pay";
  //   worksheet.getCell("B27").value = Number(
  //     earningsBottomData.premIncentivePayRate
  //   );
  //   worksheet.getCell("C27").value = Number(
  //     earningsBottomData.premIncentivePayHours
  //   );
  //   worksheet.getCell("D27").value = Number(
  //     removeNonNumericChars(earningsBottomData.premIncentivePayCurrent)
  //   );
  //   worksheet.getCell("D27").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A28").value = "Flight Vacation Pay";
  //   worksheet.getCell("B28").value = Number(
  //     earningsBottomData.fltVacationPayRate
  //   );
  //   worksheet.getCell("C28").value = Number(
  //     earningsBottomData.fltVacationPayHours
  //   );
  //   worksheet.getCell("D28").value = Number(
  //     removeNonNumericChars(earningsBottomData.fltVacationPayCurrent)
  //   );
  //   worksheet.getCell("D28").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A29").value = "Sick Pay";
  //   worksheet.getCell("B29").value = Number(earningsBottomData.sickPayRate);
  //   worksheet.getCell("C29").value = Number(earningsBottomData.sickPayHours);
  //   worksheet.getCell("D29").value = Number(
  //     removeNonNumericChars(earningsBottomData.sickPayCurrent)
  //   );
  //   worksheet.getCell("D29").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("A30").value = "Prior Year Vacation Pay Out";
  //   worksheet.getCell("B30").value = Number(
  //     earningsBottomData.priorYearVacPayoutRate
  //   );
  //   worksheet.getCell("C30").value = Number(
  //     earningsBottomData.priorYearVacPayoutHours
  //   );
  //   worksheet.getCell("D30").value = Number(
  //     earningsBottomData.priorYearVacPayoutCurrent
  //   );
  //   worksheet.getCell("D30").numFmt = "$#,##0.00"; // Format as currency

  //   worksheet.getCell("A31").value = "Earnings Sub Total";
  //   worksheet.getCell("A31").font = { bold: true };
  //   worksheet.getCell("D31").value = { formula: "SUM(D18:D30)", result: 7 };
  //   worksheet.getCell("D31").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("D31").font = { bold: true };

  //   //? Deductions
  //   worksheet.getCell("E6").value = "DEDUCTIONS";
  //   worksheet.mergeCells("E6:F6");
  //   for (let cell of [
  //     "E8",
  //     "E9",
  //     "E10",
  //     "E11",
  //     "E12",
  //     "E13",
  //     "E15",
  //     "E16",
  //     "E17",
  //     "E18",
  //     "E20",
  //     "E21",
  //     "E22",
  //     "E23",
  //     "E24",
  //     "E25",
  //   ]) {
  //     worksheet.getCell(cell).alignment = {
  //       vertical: "middle",
  //       horizontal: "left",
  //     };
  //   }
  //   //? Pre-Tax Deductions
  //   worksheet.getCell("E7").value = "Pre-Tax Deductions";
  //   worksheet.getCell("F7").value = "Current";
  //   worksheet.getCell("E8").value = "Medical Coverage";
  //   worksheet.getCell("F8").value = Number(
  //     removeNonNumericChars(preTaxDeductionsData.medicalCoverage)
  //   );
  //   worksheet.getCell("F8").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("E9").value = "Dental Coverage";
  //   worksheet.getCell("F9").value = Number(
  //     removeNonNumericChars(preTaxDeductionsData.dentalCoverage)
  //   );
  //   worksheet.getCell("F9").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("E10").value = "Vision Coverage";
  //   worksheet.getCell("F10").value = Number(
  //     removeNonNumericChars(preTaxDeductionsData.visionCoverage)
  //   );
  //   worksheet.getCell("F10").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("E11").value = "Accident Ins Pre-tax";
  //   worksheet.getCell("F11").value = Number(
  //     removeNonNumericChars(preTaxDeductionsData.accidentInsPreTax)
  //   );
  //   worksheet.getCell("F11").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("E12").value = "401k";
  //   worksheet.getCell("F12").value = Number(
  //     removeNonNumericChars(preTaxDeductionsData._401k)
  //   );
  //   worksheet.getCell("F12").numFmt = "$#,##0.00"; // Format as currency

  //   worksheet.getCell("E13").value = "Pre-Tax Deductions Total";
  //   worksheet.getCell("E13").font = { bold: true };
  //   worksheet.getCell("F13").value = { formula: "SUM(F8:F12)", result: 7 };
  //   worksheet.getCell("F13").font = { bold: true };
  //   worksheet.getCell("F13").numFmt = "$#,##0.00"; // Format as currency

  //   //? Taxes
  //   worksheet.getCell("E14").value = "Taxes";
  //   worksheet.getCell("F14").value = "Current";
  //   worksheet.getCell("E15").value = "Withholding Tax";
  //   worksheet.getCell("F15").value = Number(
  //     removeNonNumericChars(taxDeductionsData.withholdingTax)
  //   );
  //   worksheet.getCell("F15").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("E16").value = "Social Security Tax";
  //   worksheet.getCell("F16").value = Number(
  //     removeNonNumericChars(taxDeductionsData.socialSecurityTax)
  //   );
  //   worksheet.getCell("F17").numFmt = "$#,##0.00"; // Format as currency
  //   worksheet.getCell("E17").value = "Medicare Tax";
  //   worksheet.getCell("F17").value = Number(
  //     removeNonNumericChars(taxDeductionsData.medicareTax)
  //   );
  //   worksheet.getCell("F17").numFmt = "$#,##0.00"; // Format as currency

  //   worksheet.getCell("E18").value = "Taxes Total";
  //   worksheet.getCell("E18").font = { bold: true };
  //   worksheet.getCell("F18").value = { formula: "SUM(F15:F17)", result: 7 };
  //   worksheet.getCell("F18").font = { bold: true };
  //   worksheet.getCell("F18").numFmt = "$#,##0.00"; // Format as currency

  //   //? After-Tax Deductions
  //   worksheet.getCell("E19").value = "After-Tax Deductions";
  //   worksheet.getCell("F19").value = "Current";
  //   worksheet.getCell("E20").value = "Employee Life";
  //   worksheet.getCell("F20").value = Number(
  //     removeNonNumericChars(afterTaxDeductionsData.employeeLife)
  //   );
  //   worksheet.getCell("F20").numFmt = "$#,##0.00";
  //   worksheet.getCell("E21").value = "Dental Discount Plan";
  //   worksheet.getCell("F21").value = Number(
  //     removeNonNumericChars(afterTaxDeductionsData.dentalDiscountPlan)
  //   );
  //   worksheet.getCell("F21").numFmt = "$#,##0.00";
  //   worksheet.getCell("E22").value = "Roth 401k";
  //   worksheet.getCell("F22").value = Number(
  //     removeNonNumericChars(afterTaxDeductionsData.roth401k)
  //   );
  //   worksheet.getCell("F22").numFmt = "$#,##0.00";
  //   worksheet.getCell("E23").value = "PAC - APA";
  //   worksheet.getCell("F23").value = Number(
  //     removeNonNumericChars(afterTaxDeductionsData.pacAPA)
  //   );
  //   worksheet.getCell("F23").numFmt = "$#,##0.00";
  //   worksheet.getCell("E24").value = "Union Dues - APA";
  //   worksheet.getCell("F24").value = Number(
  //     removeNonNumericChars(afterTaxDeductionsData.unionDues)
  //   );
  //   worksheet.getCell("F24").numFmt = "$#,##0.00";

  //   worksheet.getCell("E25").value = "After-Tax Deductions Total";
  //   worksheet.getCell("E25").font = { bold: true };
  //   worksheet.getCell("F25").value = { formula: "SUM(F20:F24)", result: 7 };
  //   worksheet.getCell("F25").font = { bold: true };
  //   worksheet.getCell("F25").numFmt = "$#,##0.00";

  //   worksheet.mergeCells("E26:F31");

  //   //? Taxable Earnings
  //   worksheet.getCell("A34").value = "Taxable Earnings - Federal Taxes";
  //   worksheet.mergeCells("A34:C34");
  //   worksheet.getCell("D34").value = "Current";

  //   worksheet.getCell("A35").value = "Withholding Tax";
  //   worksheet.mergeCells("A35:C35");
  //   worksheet.getCell("D35").value = Number(
  //     removeNonNumericChars(taxableEarningsData.withHoldingTaxEarnings)
  //   );
  //   worksheet.getCell("D35").numFmt = "$#,##0.00";

  //   worksheet.getCell("A36").value = "Social Security Tax";
  //   worksheet.mergeCells("A36:C36");
  //   worksheet.getCell("D36").value = Number(
  //     removeNonNumericChars(taxableEarningsData.socialSecurityTaxEarnings)
  //   );
  //   worksheet.getCell("D36").numFmt = "$#,##0.00";

  //   worksheet.getCell("A37").value = "Medicare Tax";
  //   worksheet.mergeCells("A37:C37");
  //   worksheet.getCell("D37").value = Number(
  //     removeNonNumericChars(taxableEarningsData.medicareTaxEarnings)
  //   );
  //   worksheet.getCell("D37").numFmt = "$#,##0.00";

  //   //? Overpayments
  //   worksheet.getCell("A32").value = "Overpayments";
  //   worksheet.getCell("B32").value = "Moved A/R";
  //   worksheet.getCell("C32").value = "Original Balance";
  //   worksheet.getCell("D32").value = "Current Recovery";
  //   worksheet.getCell("E32").value = "Total Recovery";
  //   worksheet.getCell("F32").value = "Overpayments Total";

  //   worksheet.getCell("A33").value = "NOT TRACKED - FOR FUTURE USE";
  //   worksheet.mergeCells("A33:F33");

  //   //? Additional Information
  //   worksheet.getCell("E34").value = "Additional Information";
  //   worksheet.getCell("F34").value = "Current";

  //   worksheet.getCell("E35").value = "401k Company Contrib.";
  //   worksheet.getCell("F35").value = Number(
  //     removeNonNumericChars(companyContributionsData._401kCompanyContribution)
  //   );
  //   worksheet.getCell("F35").numFmt = "$#,##0.00";

  //   worksheet.getCell("E36").value = "Imputed Income";
  //   worksheet.getCell("F36").value = "Current";
  //   worksheet.getCell("E37").value = "Group Term Life";
  //   worksheet.getCell("F37").value = Number(
  //     removeNonNumericChars(companyContributionsData.groupTermLife)
  //   );
  //   worksheet.getCell("F37").numFmt = "$#,##0.00"; //

  //   worksheet.getCell("A40").value = extractedText; 
  //   worksheet.getCell("A40").hidden = true;// Store the full extracted text in A40

  //   return worksheet;
  // }

}
