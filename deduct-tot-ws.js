export class DeductTotalsWorksheet {

  async getAllDeductions(workbook) {
    // Create a new worksheet for totals
    let worksheet = workbook.getWorksheet("Deduction Totals");
    if (!worksheet) {
      worksheet = workbook.addWorksheet("Deduction Totals");
    }

    for (let col of ["B","G"]) {
      worksheet.getColumn(col).width = 20
    }
    await this.addPreTaxDeductionsToWorkbook(workbook, worksheet);
    await this.addMedicalDeductionsToWorkbook(workbook, worksheet);
    await this.addDentalDeductionsToWorkbook(workbook, worksheet);

    return workbook;

  }

  async addPreTaxDeductionsToWorkbook(workbook, worksheet) {
    
    const preTaxDeductions = this.getPreTaxDeductionsData(workbook);
    worksheet.getCell("A1").value = "Total Pre-Tax Deductions";
    worksheet.getCell("A1").font = { bold: true, size: 14 };
    worksheet.getCell("A2").value = "Date";
    worksheet.getCell("B2").value = "Amount";
    worksheet.getCell("A2").font = { bold: true };
    worksheet.getCell("B2").font = { bold: true };
    worksheet.getCell("B2").alignment = { horizontal: 'right' };

    for (let i = 0; i < preTaxDeductions.length; i++) {
      const row = i + 3; // Start from row 3
      worksheet.getCell(`A${row}`).value = preTaxDeductions[i].date;
      worksheet.getCell(`B${row}`).value =
        preTaxDeductions[i].deductionAmount;
      worksheet.getCell(`B${row}`).numFmt = "$#,##0.00"; // Format as currency
      worksheet.getCell((`A${row + 1}`)).value = "Total";
      worksheet.getCell((`B${row + 1}`)).value = {
        formula: `SUM(B3:B${row})`, result: 0,
      };
      worksheet.getCell((`A${row + 1}`)).font = { bold: true };
      worksheet.getCell((`B${row + 1}`)).numFmt = "$#,##0.00"; // Format as currency
      worksheet.getCell((`B${row + 1}`)).font = { bold: true };
    }

    return workbook;
  }

  getPreTaxDeductionsData(workbook) {
    let preTaxDeductions = [];
    workbook.eachSheet((sheet) => {
      const deductionAmount = sheet.getCell("B5").value;
      const sheetDate = sheet.getCell("B1").value;
      if (deductionAmount >= 0 && sheetDate) {
        preTaxDeductions.push({
          date: sheetDate,
          deductionAmount: deductionAmount.result || deductionAmount,
        });
      }
    });

    return preTaxDeductions;
  }

  async addPostTaxDeductionsToWorkbook(workbook, worksheet) {}

  async addAfterTaxDeductionsToWorkbook(workbook, worksheet) {}

  async addMedicalDeductionsToWorkbook(workbook, worksheet) {

    const medicalCoverage = this.getMedicalCoverageData(workbook);
    worksheet.getCell("F1").value = "Medical Coverage Totals";
    worksheet.getCell("F1").font = { bold: true, size: 14 };
    worksheet.getCell("F2").value = "Date";
    worksheet.getCell("F2").font = { bold: true };
    worksheet.getCell("G2").value = "Amount";
    worksheet.getCell("G2").font = { bold: true };
    worksheet.getCell("G2").alignment = { horizontal: 'right' };

    for (let i = 0; i < medicalCoverage.length; i++) {
      const row = i + 3; // Start from row 3
      worksheet.getCell(`F${row}`).value = medicalCoverage[i].date;
      worksheet.getCell(`G${row}`).value = medicalCoverage[i].totalMedicalCoverage;
      worksheet.getCell(`G${row}`).numFmt = "$#,##0.00"; 

      worksheet.getCell((`F${row + 1}`)).value = "Total";
      worksheet.getCell((`F${row + 1}`)).font = { bold: true };
      
      worksheet.getCell((`G${row + 1}`)).value = { formula: `SUM(G3:G${row})`, result: 0 };
      worksheet.getCell((`G${row + 1}`)).numFmt = "$#,##0.00";
      worksheet.getCell((`G${row + 1}`)).font = { bold: true };
      
    }

    return workbook;
  }

  getMedicalCoverageData(workbook) {
    let medicalCoverage = [];
    workbook.eachSheet((sheet) => {
      const totalMedicalCoverage = sheet.getCell("F8").value;
      const sheetDate = sheet.getCell("B1").value;
      if (totalMedicalCoverage >= 0 && sheetDate) {
        medicalCoverage.push({
          date: sheetDate,
          totalMedicalCoverage:
            totalMedicalCoverage.result || totalMedicalCoverage,
        });
      }
    });

    return medicalCoverage;
  }

  addDentalDeductionsToWorkbook(workbook, worksheet) {

    const dentalCoverage = this.getDentalCoverageData(workbook);
    worksheet.getCell("F27").value = "Dental Coverage Totals";
    worksheet.getCell("F27").font = { bold: true, size: 14 };
    worksheet.getCell("F28").value = "Date";
    worksheet.getCell("G28").value = "Amount";
    worksheet.getCell("F28").font = { bold: true };
    worksheet.getCell("G28").font = { bold: true };
    worksheet.getCell("G28").alignment = { horizontal: 'right' };

    for (let i = 0; i < dentalCoverage.length; i++) {
      const row = i + 29; // Start from row 29
      worksheet.getCell(`F${row}`).value = dentalCoverage[i].date;
      worksheet.getCell(`G${row}`).value =
        dentalCoverage[i].totalDentalCoverage;
      worksheet.getCell(`G${row}`).numFmt = "$#,##0.00"; // Format as currency
      worksheet.getCell((`F${row + 1}`)).value = "Total";
      worksheet.getCell((`G${row + 1}`)).value = {
        formula: `SUM(G29:G${row})`, result: 0,
      };
      worksheet.getCell((`F${row + 1}`)).font = { bold: true };
      worksheet.getCell((`G${row + 1}`)).numFmt = "$#,##0.00"; // Format as currency
      worksheet.getCell((`G${row + 1}`)).font = { bold: true };
    }

    return workbook;
  }

  getDentalCoverageData(workbook) {
    let dentalCoverage = [];
    workbook.eachSheet((sheet) => {
      const totalDentalCoverage = sheet.getCell("F9").value;
      const sheetDate = sheet.getCell("B1").value;
      if (totalDentalCoverage >= 0 && sheetDate) {
        dentalCoverage.push({
          date: sheetDate,
          totalDentalCoverage:
            totalDentalCoverage.result || totalDentalCoverage,
        });
      }
    });

    return dentalCoverage;
  }
}
