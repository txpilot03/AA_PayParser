import { TextParser } from "./parse-text.js";
import { DateTime} from "luxon";
import { Utilities } from "./utilities.js";


export class WorkSheetGenerator {

  util = new Utilities();
  textParser = new TextParser();

  // constructor() {
  //   this.textParser = new TextParser();
  //   this.util = new Utilities();
  // }

  async addDataToWorkbook(workbook, extractedText) {
    const removeNonNumericChars = this.util.removeNonNumericChars;

    const infoData = this.textParser.parsePDFInfoData(extractedText);
    const summaryData = this.textParser.parsePDFSummaryData(extractedText);
    const earningsTopData =
      this.textParser.parsePDFEarningsTopData(extractedText);
    const earningsBottomData =
      this.textParser.parsePDFEarningsBottomData(extractedText);
    const preTaxDeductionsData =
      this.textParser.parsePDFPreTaxDeductionsData(extractedText);
    const taxDeductionsData =
      this.textParser.parsePDFTaxDeductionsData(extractedText);
    const afterTaxDeductionsData =
      this.textParser.parsePDFAfterTaxDeductionsData(extractedText);
    const companyContributionsData =
      this.textParser.parseCompanyContributionsData(extractedText);
    const taxableEarningsData =
      this.textParser.parsePDFTaxableEarningsData(extractedText);

    const date = this.util.convertStringToDate(infoData.payPeriod);
    const worksheetName = DateTime.fromISO(date.toISOString()).toFormat(
      "ddMMMyyyy"
    );

    // Check if the sheet already exists
    if (workbook.getWorksheet(worksheetName)) {
      throw new Error(
        `A worksheet with the name "${worksheetName}" already exists.`
      );
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
      "A4",
      "B4",
      "C4",
      "D4",
      "E4",
      "A1",
      "A2",
      "B2",
      "C2",
      "A6",
      "B6",
      "C6",
      "D6",
      "A7",
      "B7",
      "C7",
      "D7",
      "E6",
      "E7",
      "F7",
      "E14",
      "F14",
      "E19",
      "F19",
      "A32",
      "B32",
      "C32",
      "D32",
      "E32",
      "F32",
      "A34",
      "B34",
      "C34",
      "D34",
      "E34",
      "F34",
      "E36",
      "F36",
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

    for (let cell of [
      "A6",
      "A7",
      "B7",
      "C7",
      "D7",
      "E6",
      "E7",
      "F7",
      "E14",
      "F14",
      "E19",
      "F19",
      "A32",
      "B32",
      "C32",
      "D32",
      "E32",
      "F32",
      "A34",
      "D34",
      "E34",
      "F34",
      "E36",
      "F36",
    ]) {
      worksheet.getCell(cell).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }

    for (let cell of [
      "A8",
      "A9",
      "A10",
      "A11",
      "A12",
      "A13",
      "A14",
      "A15",
      "A16",
      "A17",
      "A18",
      "A19",
      "A20",
      "A21",
      "A22",
      "A23",
      "A24",
      "A25",
      "A26",
      "A27",
      "A28",
      "A29",
      "A30",
      "A31",
      "A35",
      "A36",
      "A37",
      "E35",
      "E37",
    ]) {
      worksheet.getCell(cell).alignment = {
        vertical: "middle",
        horizontal: "left",
      };
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
    worksheet.getCell("C3").value = Number(
      removeNonNumericChars(infoData.hourlyRate)
    );
    worksheet.getCell("C3").numFmt = "$#,##0.00";

    //? Summary
    worksheet.getCell("A4").value = "Gross Earnings";
    worksheet.getCell("B4").value = "Pre-Tax Deduction";
    worksheet.getCell("C4").value = "Taxes";
    worksheet.getCell("D4").value = "After-Tax Deduction";
    worksheet.getCell("E4").value = "Net Pay";

    worksheet.getCell("A5").value = Number(
      removeNonNumericChars(summaryData.gross)
    );
    worksheet.getCell("A5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("B5").value = Number(
      removeNonNumericChars(summaryData.preTaxDeduct)
    );
    worksheet.getCell("B5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("C5").value = Number(
      removeNonNumericChars(summaryData.taxes)
    );
    worksheet.getCell("C5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("D5").value = Number(
      removeNonNumericChars(summaryData.afterTaxDeduct)
    );
    worksheet.getCell("D5").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E5").value = Number(
      removeNonNumericChars(summaryData.netPay)
    );
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
    worksheet.getCell("D8").value = Number(
      removeNonNumericChars(earningsTopData.crewAdvanceCurrent)
    );
    worksheet.getCell("D8").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A9").value = "PLT EXP D Taxable";
    worksheet.getCell("B9").value = Number(earningsTopData.pltExpDtaxRate);
    worksheet.getCell("C9").value = Number(earningsTopData.pltExpDtaxHours);
    worksheet.getCell("D9").value = Number(
      removeNonNumericChars(earningsTopData.pltExpDtaxCurrent)
    );
    worksheet.getCell("D9").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A10").value = "PLT EXP ADJ Taxable";
    worksheet.getCell("B10").value = Number(earningsTopData.pltExpADJtaxRate);
    worksheet.getCell("C10").value = Number(earningsTopData.pltExpADJtaxHours);
    worksheet.getCell("D10").value = Number(
      removeNonNumericChars(earningsTopData.pltExpADJtaxCurrent)
    );
    worksheet.getCell("D10").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A11").value = "PLT EXP I Taxable";
    worksheet.getCell("B11").value = Number(earningsTopData.pltExpItaxRate);
    worksheet.getCell("C11").value = Number(earningsTopData.pltExpItaxHours);
    worksheet.getCell("D11").value = Number(
      removeNonNumericChars(earningsTopData.pltExpItaxCurrent)
    );
    worksheet.getCell("D11").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A12").value = "AAG Profit Sharing";
    worksheet.getCell("B12").value = Number(
      earningsTopData.AagProfitSharingRate
    );
    worksheet.getCell("C12").value = Number(
      earningsTopData.AagProfitSharingHours
    );
    worksheet.getCell("D12").value = Number(
      removeNonNumericChars(earningsTopData.AagProfitSharingCurrent)
    );
    worksheet.getCell("D12").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A13").value = "PLT EXP D Non-Taxable";
    worksheet.getCell("B13").value = Number(earningsTopData.pltExpDnonTaxRate);
    worksheet.getCell("C13").value = Number(earningsTopData.pltExpDnonTaxHours);
    worksheet.getCell("D13").value = Number(
      removeNonNumericChars(earningsTopData.pltExpDnonTaxCurrent)
    );
    worksheet.getCell("D13").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A14").value = "PLT EXP ADJ Non-Taxable";
    worksheet.getCell("B14").value = Number(
      earningsTopData.pltExpAdjnonTaxRate
    );
    worksheet.getCell("C14").value = Number(
      earningsTopData.pltExpAdjnonTaxHours
    );
    worksheet.getCell("D14").value = Number(
      removeNonNumericChars(earningsTopData.pltExpAdjnonTaxCurrent)
    );
    worksheet.getCell("D14").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A15").value = "PLT EXP I Non-Taxable";
    worksheet.getCell("B15").value = Number(earningsTopData.pltExpInonTaxRate);
    worksheet.getCell("C15").value = Number(earningsTopData.pltExpInonTaxHours);
    worksheet.getCell("D15").value = Number(
      removeNonNumericChars(earningsTopData.pltExpInonTaxCurrent)
    );
    worksheet.getCell("D15").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("A16").value = "Earnings Total";
    worksheet.getCell("A16").font = { bold: true };
    worksheet.getCell("D16").value = { formula: "SUM(D8:D15)", result: 7 };
    worksheet.getCell("D16").font = { bold: true };
    worksheet.getCell("D16").numFmt = "$#,##0.00"; // Format as currency

    //* Earnings Bottom
    worksheet.getCell("A18").value = "Operational Pay";
    worksheet.getCell("B18").value = Number(
      earningsBottomData.operationalPayRate
    );
    worksheet.getCell("C18").value = Number(
      earningsBottomData.operationalPayHours
    );
    worksheet.getCell("D18").value = Number(
      removeNonNumericChars(earningsBottomData.operationalPayCurrent)
    );
    worksheet.getCell("D18").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A19").value = "Flight Training Pay";
    worksheet.getCell("B19").value = Number(
      earningsBottomData.fltTrainingPayRate
    );
    worksheet.getCell("C19").value = Number(
      earningsBottomData.fltTrainingPayHours
    );
    worksheet.getCell("D19").value = Number(
      removeNonNumericChars(earningsBottomData.fltTrainingPayCurrent)
    );
    worksheet.getCell("D19").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A20").value = "Sit Time";
    worksheet.getCell("B20").value = Number(earningsBottomData.sitTimeRate);
    worksheet.getCell("C20").value = Number(earningsBottomData.sitTimeHours);
    worksheet.getCell("D20").value = Number(
      removeNonNumericChars(earningsBottomData.sitTimeCurrent)
    );
    worksheet.getCell("D20").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A21").value = "Pay Above Guarantee (RSV)";
    worksheet.getCell("B21").value = Number(
      earningsBottomData.payAbvGuaranteeRsvRate
    );
    worksheet.getCell("C21").value = Number(
      earningsBottomData.payAbvGuaranteeRsvHours
    );
    worksheet.getCell("D21").value = Number(
      removeNonNumericChars(earningsBottomData.payAbvGuaranteeRsvCurrent)
    );
    worksheet.getCell("D21").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A22").value = "RA Prem";
    worksheet.getCell("B22").value = Number(earningsBottomData.raPremRate);
    worksheet.getCell("C22").value = Number(earningsBottomData.raPremHours);
    worksheet.getCell("D22").value = Number(
      removeNonNumericChars(earningsBottomData.raPremCurrent)
    );
    worksheet.getCell("D22").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A23").value = "Min Guarantee Adj";
    worksheet.getCell("B23").value = Number(
      earningsBottomData.minGuaranteeAdjRate
    );
    worksheet.getCell("C23").value = Number(
      earningsBottomData.minGuaranteeAdjHours
    );
    worksheet.getCell("D23").value = Number(
      removeNonNumericChars(earningsBottomData.minGuaranteeAdjCurrent)
    );
    worksheet.getCell("D23").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A24").value = "Intl Override";
    worksheet.getCell("B24").value = Number(
      earningsBottomData.intlOverrideRate
    );
    worksheet.getCell("C24").value = Number(
      earningsBottomData.intlOverrideHours
    );
    worksheet.getCell("D24").value = Number(
      removeNonNumericChars(earningsBottomData.intlOverrideCurrent)
    );
    worksheet.getCell("D24").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A25").value = "Distance Learning";
    worksheet.getCell("B25").value = Number(
      earningsBottomData.distanceLearningRate
    );
    worksheet.getCell("C25").value = Number(
      earningsBottomData.distanceLearningHours
    );
    worksheet.getCell("D25").value = Number(
      removeNonNumericChars(earningsBottomData.distanceLearningCurrent)
    );
    worksheet.getCell("D25").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A26").value = "Union PD Union Leave";
    worksheet.getCell("B26").value = Number(
      earningsBottomData.unionPdLeaveRate
    );
    worksheet.getCell("C26").value = Number(
      earningsBottomData.unionPdLeaveHours
    );
    worksheet.getCell("D26").value = Number(
      removeNonNumericChars(earningsBottomData.unionPdLeaveCurrent)
    );
    worksheet.getCell("D26").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A27").value = "Prem Incentive Pay";
    worksheet.getCell("B27").value = Number(
      earningsBottomData.premIncentivePayRate
    );
    worksheet.getCell("C27").value = Number(
      earningsBottomData.premIncentivePayHours
    );
    worksheet.getCell("D27").value = Number(
      removeNonNumericChars(earningsBottomData.premIncentivePayCurrent)
    );
    worksheet.getCell("D27").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A28").value = "Flight Vacation Pay";
    worksheet.getCell("B28").value = Number(
      earningsBottomData.fltVacationPayRate
    );
    worksheet.getCell("C28").value = Number(
      earningsBottomData.fltVacationPayHours
    );
    worksheet.getCell("D28").value = Number(
      removeNonNumericChars(earningsBottomData.fltVacationPayCurrent)
    );
    worksheet.getCell("D28").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A29").value = "Sick Pay";
    worksheet.getCell("B29").value = Number(earningsBottomData.sickPayRate);
    worksheet.getCell("C29").value = Number(earningsBottomData.sickPayHours);
    worksheet.getCell("D29").value = Number(
      removeNonNumericChars(earningsBottomData.sickPayCurrent)
    );
    worksheet.getCell("D29").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("A30").value = "Prior Year Vacation Pay Out";
    worksheet.getCell("B30").value = Number(
      earningsBottomData.priorYearVacPayoutRate
    );
    worksheet.getCell("C30").value = Number(
      earningsBottomData.priorYearVacPayoutHours
    );
    worksheet.getCell("D30").value = Number(
      earningsBottomData.priorYearVacPayoutCurrent
    );
    worksheet.getCell("D30").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("A31").value = "Earnings Sub Total";
    worksheet.getCell("A31").font = { bold: true };
    worksheet.getCell("D31").value = { formula: "SUM(D18:D30)", result: 7 };
    worksheet.getCell("D31").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("D31").font = { bold: true };

    //? Deductions
    worksheet.getCell("E6").value = "DEDUCTIONS";
    worksheet.mergeCells("E6:F6");
    for (let cell of [
      "E8",
      "E9",
      "E10",
      "E11",
      "E12",
      "E13",
      "E15",
      "E16",
      "E17",
      "E18",
      "E20",
      "E21",
      "E22",
      "E23",
      "E24",
      "E25",
    ]) {
      worksheet.getCell(cell).alignment = {
        vertical: "middle",
        horizontal: "left",
      };
    }
    //* Pre-Tax Deductions
    worksheet.getCell("E7").value = "Pre-Tax Deductions";
    worksheet.getCell("F7").value = "Current";
    worksheet.getCell("E8").value = "Medical Coverage";
    worksheet.getCell("F8").value = Number(
      removeNonNumericChars(preTaxDeductionsData.medicalCoverage)
    );
    worksheet.getCell("F8").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E9").value = "Dental Coverage";
    worksheet.getCell("F9").value = Number(
      removeNonNumericChars(preTaxDeductionsData.dentalCoverage)
    );
    worksheet.getCell("F9").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E10").value = "Vision Coverage";
    worksheet.getCell("F10").value = Number(
      removeNonNumericChars(preTaxDeductionsData.visionCoverage)
    );
    worksheet.getCell("F10").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E11").value = "Accident Ins Pre-tax";
    worksheet.getCell("F11").value = Number(
      removeNonNumericChars(preTaxDeductionsData.accidentInsPreTax)
    );
    worksheet.getCell("F11").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E12").value = "401k";
    worksheet.getCell("F12").value = Number(
      removeNonNumericChars(preTaxDeductionsData._401k)
    );
    worksheet.getCell("F12").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("E13").value = "Pre-Tax Deductions Total";
    worksheet.getCell("E13").font = { bold: true };
    worksheet.getCell("F13").value = { formula: "SUM(F8:F12)", result: 7 };
    worksheet.getCell("F13").font = { bold: true };
    worksheet.getCell("F13").numFmt = "$#,##0.00"; // Format as currency

    //* Taxes
    worksheet.getCell("E14").value = "Taxes";
    worksheet.getCell("F14").value = "Current";
    worksheet.getCell("E15").value = "Withholding Tax";
    worksheet.getCell("F15").value = Number(
      removeNonNumericChars(taxDeductionsData.withholdingTax)
    );
    worksheet.getCell("F15").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E16").value = "Social Security Tax";
    worksheet.getCell("F16").value = Number(
      removeNonNumericChars(taxDeductionsData.socialSecurityTax)
    );
    worksheet.getCell("F17").numFmt = "$#,##0.00"; // Format as currency
    worksheet.getCell("E17").value = "Medicare Tax";
    worksheet.getCell("F17").value = Number(
      removeNonNumericChars(taxDeductionsData.medicareTax)
    );
    worksheet.getCell("F17").numFmt = "$#,##0.00"; // Format as currency

    worksheet.getCell("E18").value = "Taxes Total";
    worksheet.getCell("E18").font = { bold: true };
    worksheet.getCell("F18").value = { formula: "SUM(F15:F17)", result: 7 };
    worksheet.getCell("F18").font = { bold: true };
    worksheet.getCell("F18").numFmt = "$#,##0.00"; // Format as currency

    //* After-Tax Deductions
    worksheet.getCell("E19").value = "After-Tax Deductions";
    worksheet.getCell("F19").value = "Current";
    worksheet.getCell("E20").value = "Employee Life";
    worksheet.getCell("F20").value = Number(
      removeNonNumericChars(afterTaxDeductionsData.employeeLife)
    );
    worksheet.getCell("F20").numFmt = "$#,##0.00";
    worksheet.getCell("E21").value = "Dental Discount Plan";
    worksheet.getCell("F21").value = Number(
      removeNonNumericChars(afterTaxDeductionsData.dentalDiscountPlan)
    );
    worksheet.getCell("F21").numFmt = "$#,##0.00";
    worksheet.getCell("E22").value = "Roth 401k";
    worksheet.getCell("F22").value = Number(
      removeNonNumericChars(afterTaxDeductionsData.roth401k)
    );
    worksheet.getCell("F22").numFmt = "$#,##0.00";
    worksheet.getCell("E23").value = "PAC - APA";
    worksheet.getCell("F23").value = Number(
      removeNonNumericChars(afterTaxDeductionsData.pacAPA)
    );
    worksheet.getCell("F23").numFmt = "$#,##0.00";
    worksheet.getCell("E24").value = "Union Dues - APA";
    worksheet.getCell("F24").value = Number(
      removeNonNumericChars(afterTaxDeductionsData.unionDues)
    );
    worksheet.getCell("F24").numFmt = "$#,##0.00";

    worksheet.getCell("E25").value = "After-Tax Deductions Total";
    worksheet.getCell("E25").font = { bold: true };
    worksheet.getCell("F25").value = { formula: "SUM(F20:F24)", result: 7 };
    worksheet.getCell("F25").font = { bold: true };
    worksheet.getCell("F25").numFmt = "$#,##0.00";

    worksheet.mergeCells("E26:F31");

    //? Taxable Earnings
    worksheet.getCell("A34").value = "Taxable Earnings - Federal Taxes";
    worksheet.mergeCells("A34:C34");
    worksheet.getCell("D34").value = "Current";

    worksheet.getCell("A35").value = "Withholding Tax";
    worksheet.mergeCells("A35:C35");
    worksheet.getCell("D35").value = Number(
      removeNonNumericChars(taxableEarningsData.withHoldingTaxEarnings)
    );
    worksheet.getCell("D35").numFmt = "$#,##0.00";

    worksheet.getCell("A36").value = "Social Security Tax";
    worksheet.mergeCells("A36:C36");
    worksheet.getCell("D36").value = Number(
      removeNonNumericChars(taxableEarningsData.socialSecurityTaxEarnings)
    );
    worksheet.getCell("D36").numFmt = "$#,##0.00";

    worksheet.getCell("A37").value = "Medicare Tax";
    worksheet.mergeCells("A37:C37");
    worksheet.getCell("D37").value = Number(
      removeNonNumericChars(taxableEarningsData.medicareTaxEarnings)
    );
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
    worksheet.getCell("F35").value = Number(
      removeNonNumericChars(companyContributionsData._401kCompanyContribution)
    );
    worksheet.getCell("F35").numFmt = "$#,##0.00";

    worksheet.getCell("E36").value = "Imputed Income";
    worksheet.getCell("F36").value = "Current";
    worksheet.getCell("E37").value = "Group Term Life";
    worksheet.getCell("F37").value = Number(
      removeNonNumericChars(companyContributionsData.groupTermLife)
    );
    worksheet.getCell("F37").numFmt = "$#,##0.00"; //

    worksheet.getCell("A40").value = extractedText; // Store the full extracted text in A40

    return worksheet;
  }

  // async addTotalsToWorkbook(workbook) {
  //   // Create a new worksheet for totals
  //   let worksheetTotals = workbook.getWorksheet("Totals");
  //   if (!worksheetTotals) {
  //     worksheetTotals = workbook.addWorksheet("Totals");
  //   }

  //   const medicalCoverage = this.getMedicalCoverage(workbook);
  //   worksheetTotals.getCell("A1").value = "Medical Coverage Totals";
  //   worksheetTotals.getCell("A1").font = { bold: true, size: 14 };
  //   worksheetTotals.getCell("A2").value = "Date";
  //   worksheetTotals.getCell("B2").value = "Total Medical Coverage";
  //   worksheetTotals.getCell("A2").font = { bold: true };
  //   worksheetTotals.getCell("B2").font = { bold: true };
  //   for (let i = 0; i < medicalCoverage.length; i++) {
  //     const row = i + 3; // Start from row 3
  //     worksheetTotals.getCell(`A${row}`).value = medicalCoverage[i].date;
  //     worksheetTotals.getCell(`B${row}`).value =
  //       medicalCoverage[i].totalMedicalCoverage;
  //     worksheetTotals.getCell(`B${row}`).numFmt = "$#,##0.00"; // Format as currency
  //     worksheetTotals.getCell((`A${row + 1}`)).value = "Total";
  //     worksheetTotals.getCell((`B${row + 1}`)).value = {
  //       formula: `SUM(B3:B${row})`, result: 0,
  //     };
  //     worksheetTotals.getCell((`B${row + 1}`)).numFmt = "$#,##0.00"; // Format as currency
  //     worksheetTotals.getCell((`B${row + 1}`)).font = { bold: true };
  //   }

  //   return workbook;
  // }

  // getMedicalCoverage(workbook) {
  //   let medicalCoverage = [];
  //   workbook.eachSheet((sheet) => {
  //     const totalMedicalCoverage = sheet.getCell("F8").value;
  //     const sheetDate = sheet.getCell("B1").value;
  //     if (totalMedicalCoverage >= 0 && sheetDate) {
  //       medicalCoverage.push({
  //         date: sheetDate,
  //         totalMedicalCoverage:
  //           totalMedicalCoverage.result || totalMedicalCoverage,
  //       });
  //     }
  //   });

  //   return medicalCoverage;
  // }
}
