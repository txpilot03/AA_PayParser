export class DeductTotalsWorksheet {

  async getAllDeductions(workbook) {
    // Create a new worksheet for totals
    let worksheet = workbook.getWorksheet("Deduction Totals");
    if (!worksheet) {
      worksheet = workbook.addWorksheet("Deduction Totals");
    } else {
      workbook.removeWorksheet(worksheet.id);
      worksheet = workbook.addWorksheet("Deduction Totals");
    }

    for (let col of ["A","B","C","D","E","F","G"]) {
      worksheet.getColumn(col).width = 20
    }

    for (let cell of ["A1","A2","B2","C2","D2","E2","F2","G2",
      "A31","A32","B32","C32","D32","E32","A59","A60","B60",
      "C60","D60","E60","F60","G60"
    ]) {
      worksheet.getCell(cell).alignment = { horizontal: 'center' };
      if (cell === "A1" || cell === "A31" || cell === "A59") {
        worksheet.getCell(cell).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFEE943' },
        };
        worksheet.getCell(cell).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
      else {
        worksheet.getCell(cell).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFAEAEB2' },
        };
      }
    }

    worksheet.mergeCells("A1:G1");
    worksheet.mergeCells("A31:E31");
    worksheet.mergeCells("A59:G59");
    worksheet.getRow(1).height = 30;
    worksheet.getRow(31).height = 30;
    worksheet.getRow(59).height = 30;

    await this.addPreTaxDeductionsToWorkbook(workbook, worksheet);
    await this.addTaxDeductionsToWorkbook(workbook, worksheet);
    await this.addAfterTaxDeductionsToWorkbook(workbook, worksheet);

    return workbook;
  }

  //! Pre-Tax Deductions
  async addPreTaxDeductionsToWorkbook(workbook, worksheet) {
    const preTaxDeductions = this.getPreTaxDeductionsData(workbook);
    worksheet.getCell("A1").value = "Total Pre-Tax Deductions";
    worksheet.getCell("A1").font = { bold: true, size: 14 };
    worksheet.getCell("A1").alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell("A2").value = "Date";
    worksheet.getCell("A2").font = { bold: true };
    worksheet.getCell("B2").value = "Medical";
    worksheet.getCell("B2").font = { bold: true };
    worksheet.getCell("C2").value = "Dental";
    worksheet.getCell("C2").font = { bold: true };
    worksheet.getCell("D2").value = "Vision";
    worksheet.getCell("D2").font = { bold: true };
    worksheet.getCell("E2").value = "Accident Ins";
    worksheet.getCell("E2").font = { bold: true };
    worksheet.getCell("F2").value = "401(k)";
    worksheet.getCell("F2").font = { bold: true };
    worksheet.getCell("G2").value = "Total";
    worksheet.getCell("G2").font = { bold: true };

    for (let i = 0; i < preTaxDeductions.length; i++) {
      const row = i + 3;
      worksheet.getCell(`A${row}`).value = preTaxDeductions[i].date;
      worksheet.getCell(`A${row}`).alignment = { horizontal: 'center' };
      worksheet.getCell(`A${row}`).numFmt = "ddMMMyyyy";
      worksheet.getCell(`A${row}`).font = { bold: true };
      worksheet.getCell(`B${row}`).value = preTaxDeductions[i].medicalCoverage;
      worksheet.getCell(`B${row}`).numFmt = "$#,##0.00"; 
      worksheet.getCell(`C${row}`).value = preTaxDeductions[i].dentalCoverage;
      worksheet.getCell(`C${row}`).numFmt = "$#,##0.00"; 
      worksheet.getCell(`D${row}`).value = preTaxDeductions[i].visionCoverage;
      worksheet.getCell(`D${row}`).numFmt = "$#,##0.00"; 
      worksheet.getCell(`E${row}`).value = preTaxDeductions[i].accidentInsPreTax;
      worksheet.getCell(`E${row}`).numFmt = "$#,##0.00"; 
      worksheet.getCell(`F${row}`).value = preTaxDeductions[i]._401kDeduction;
      worksheet.getCell(`F${row}`).numFmt = "$#,##0.00"; 
      worksheet.getCell(`G${row}`).value = {
        formula: `SUM(B${row}:F${row})`,
        result: 0
      };
      worksheet.getCell(`G${row}`).numFmt = "$#,##0.00"; 
      worksheet.getCell(`G${row}`).font = { bold: true };
    }

    const row = preTaxDeductions.length + 3;
    worksheet.getCell(`A${row}`).value = "Total YTD";
    worksheet.getCell(`A${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`A${row}`).alignment = { horizontal: 'left' };
    worksheet.getCell(`A${row}`).border = {
      top: { style: 'thick' },
    };
    // Medical Totals
    worksheet.getCell(`B${row}`).value = { formula: `SUM(B3:B${row - 1})`, result: 0 };
    worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`B${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`B${row}`).border = {
      top: { style: 'thick' },
    };
    // Dental Totals
    worksheet.getCell(`C${row}`).value = { formula: `SUM(C3:C${row - 1})`, result: 0 };
    worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`C${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`C${row}`).border = {
      top: { style: 'thick' },
    };

    // Vision Totals
    worksheet.getCell(`D${row}`).value = { formula: `SUM(D3:D${row - 1})`, result: 0 };
    worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`D${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`D${row}`).border = {
      top: { style: 'thick' },
    };
    // Accident Ins Pre-tax Totals
    worksheet.getCell(`E${row}`).value = { formula: `SUM(E3:E${row - 1})`, result: 0 };
    worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`E${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`E${row}`).border = {
      top: { style: 'thick' },
    };
    // 401(k) Totals
    worksheet.getCell(`F${row}`).value = { formula: `SUM(F3:F${row - 1})`, result: 0 };
    worksheet.getCell(`F${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`F${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`F${row}`).border = {
      top: { style: 'thick' },
    };
    // Totals of Totals Row
    worksheet.getCell(`G${row}`).value = { formula: `SUM(G3:G${row - 1})`, result: 0 };
    worksheet.getCell(`G${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`G${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`G${row}`).border = {
      top: { style: 'thick' },
    };

    return workbook;
  }

  getPreTaxDeductionsData(workbook) {
    let preTaxDeductions = [];
    workbook.eachSheet((sheet) => {
      if (sheet.name !== "Deduction Totals" && sheet.name !== "Earning Totals") {
        const sheetDate = sheet.getCell("B1").value;
        const medicalCoverage = sheet.getCell("F8").value;
        const dentalCoverage = sheet.getCell("F9").value;
        const visionCoverage = sheet.getCell("F10").value;
        const accidentInsPreTax = sheet.getCell("F11").value;
        const _401kDeduction = sheet.getCell("F12").value;
        if (sheetDate) {
          preTaxDeductions.push({
            date: sheetDate,
            medicalCoverage: medicalCoverage || 0,
            dentalCoverage: dentalCoverage || 0,
            visionCoverage: visionCoverage || 0,
            accidentInsPreTax: accidentInsPreTax || 0,
            _401kDeduction: _401kDeduction || 0,
          });
        }
      }
    });

    return preTaxDeductions;
  }

  //! Tax Deductions
  async addTaxDeductionsToWorkbook(workbook, worksheet) {
    const postTaxDeductions = this.getTaxDeductionsData(workbook);
    worksheet.getCell("A31").value = "Total Tax Deductions";
    worksheet.getCell("A31").font = { bold: true, size: 14 };
    worksheet.getCell("A31").alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell("A32").value = "Date";
    worksheet.getCell("A32").font = { bold: true };
    worksheet.getCell("B32").value = "Withholding";
    worksheet.getCell("B32").font = { bold: true };
    worksheet.getCell("C32").value = "Social Security";
    worksheet.getCell("C32").font = { bold: true };
    worksheet.getCell("D32").value = "Medicare";
    worksheet.getCell("D32").font = { bold: true };
    worksheet.getCell("E32").value = "Total";
    worksheet.getCell("E32").font = { bold: true };

    for (let i = 0; i < postTaxDeductions.length; i++) {
      const row = i + 33;
      worksheet.getCell(`A${row}`).value = postTaxDeductions[i].date;
      worksheet.getCell(`A${row}`).alignment = { horizontal: 'center' };
      worksheet.getCell(`A${row}`).numFmt = "ddMMMyyyy";
      worksheet.getCell(`A${row}`).font = { bold: true };
      worksheet.getCell(`B${row}`).value = postTaxDeductions[i].withholding;
      worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`C${row}`).value = postTaxDeductions[i].socialSecurity;
      worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`D${row}`).value = postTaxDeductions[i].medicare;
      worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`E${row}`).value = { formula: `SUM(B${row}:D${row})`, result: 0 };
      worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
    }

    const row = postTaxDeductions.length + 33;
    worksheet.getCell(`A${row}`).value = "Total YTD";
    worksheet.getCell(`A${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`A${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`A${row}`).alignment = { horizontal: 'left' };
    worksheet.getCell(`A${row}`).border = {
      top: { style: 'thick' },
    };
    // Withholding Totals
    worksheet.getCell(`B${row}`).value = { formula: `SUM(B33:B${row - 1})`, result: 0 };
    worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`B${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`B${row}`).border = {
      top: { style: 'thick' },
    };
    // Social Security Totals
    worksheet.getCell(`C${row}`).value = { formula: `SUM(C33:C${row - 1})`, result: 0 };
    worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`C${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`C${row}`).border = {
      top: { style: 'thick' },
    };
    // Medicare Totals
    worksheet.getCell(`D${row}`).value = { formula: `SUM(D33:D${row - 1})`, result: 0 };
    worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`D${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`D${row}`).border = {
      top: { style: 'thick' },
    };
    // Total of Totals Row
    worksheet.getCell(`E${row}`).value = { formula: `SUM(E33:E${row - 1})`, result: 0 };
    worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`E${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`E${row}`).border = {
      top: { style: 'thick' },
    };

    return workbook;
  }

  getTaxDeductionsData(workbook) {
    let postTaxDeductions = [];
    workbook.eachSheet((sheet) => {
      if (sheet.name !== "Deduction Totals" && sheet.name !== "Earning Totals") {
        const sheetDate = sheet.getCell("B1").value;
        const withholding = sheet.getCell("F15").value;
        const socialSecurity = sheet.getCell("F16").value;
        const medicare = sheet.getCell("F17").value;
        if (sheetDate) {
          postTaxDeductions.push({
            date: sheetDate,
            withholding: withholding || 0,
            socialSecurity: socialSecurity || 0,
            medicare: medicare || 0,
          });
        }
      }
    });

    return postTaxDeductions;
  }

  //! After-Tax Deductions
  async addAfterTaxDeductionsToWorkbook(workbook, worksheet) {
    const afterTaxDeductions = this.getAfterTaxDeductionsData(workbook);
    worksheet.getCell("A59").value = "Total After-Tax Deductions";
    worksheet.getCell("A59").font = { bold: true, size: 14 };
    worksheet.getCell("A59").alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell("A60").value = "Date";
    worksheet.getCell("A60").font = { bold: true };
    worksheet.getCell("B60").value = "Employee Life";
    worksheet.getCell("B60").font = { bold: true };
    worksheet.getCell("C60").value = "Dental Discount";
    worksheet.getCell("C60").font = { bold: true };
    worksheet.getCell("D60").value = "Roth 401(k)";
    worksheet.getCell("D60").font = { bold: true };
    worksheet.getCell("E60").value = "PAC-APA";
    worksheet.getCell("E60").font = { bold: true };
    worksheet.getCell("F60").value = "Union Dues";
    worksheet.getCell("F60").font = { bold: true };
    worksheet.getCell("G60").value = "Total";
    worksheet.getCell("G60").font = { bold: true };

    for (let i = 0; i < afterTaxDeductions.length; i++) {
      const row = i + 61;
      worksheet.getCell(`A${row}`).value = afterTaxDeductions[i].date;
      worksheet.getCell(`A${row}`).alignment = { horizontal: 'center' };
      worksheet.getCell(`A${row}`).numFmt = "ddMMMyyyy";
      worksheet.getCell(`A${row}`).font = { bold: true };
      worksheet.getCell(`B${row}`).value = afterTaxDeductions[i].employeeLife;
      worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`C${row}`).value = afterTaxDeductions[i].dentalDiscountPlan;
      worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`D${row}`).value = afterTaxDeductions[i].roth401k;
      worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`E${row}`).value = afterTaxDeductions[i].pacAPA;
      worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`F${row}`).value = afterTaxDeductions[i].unionDues;
      worksheet.getCell(`F${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`G${row}`).value = {
        formula: `SUM(B${row}:F${row})`,
        result: 0
      };
      worksheet.getCell(`G${row}`).numFmt = "$#,##0.00"; 
      worksheet.getCell(`G${row}`).font = { bold: true };
    }

    const row = afterTaxDeductions.length + 61;
    worksheet.getCell(`A${row}`).value = "Total YTD";
    worksheet.getCell(`A${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`A${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`A${row}`).alignment = { horizontal: 'left' };
    worksheet.getCell(`A${row}`).border = {
      top: { style: 'thick' },
    };
    // Employee Life Totals
    worksheet.getCell(`B${row}`).value = { formula: `SUM(B61:B${row - 1})`, result: 0 };
    worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`B${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`B${row}`).border = {
      top: { style: 'thick' },
    };
    // Dental Discount Plan Totals
    worksheet.getCell(`C${row}`).value = { formula: `SUM(C61:C${row - 1})`, result: 0 };
    worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`C${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`C${row}`).border = {
      top: { style: 'thick' },
    };
    // Roth 401(k) Totals
    worksheet.getCell(`D${row}`).value = { formula: `SUM(D61:D${row - 1})`, result: 0 };
    worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`D${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`D${row}`).border = {
      top: { style: 'thick' },
    };
    // PAC-APA Totals
    worksheet.getCell(`E${row}`).value = { formula: `SUM(E61:E${row - 1})`, result: 0 };
    worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`E${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`E${row}`).border = {
      top: { style: 'thick' },
    };
    // Union Dues Totals
    worksheet.getCell(`F${row}`).value = { formula: `SUM(F61:F${row - 1})`, result: 0 };
    worksheet.getCell(`F${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`F${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`F${row}`).border = {
      top: { style: 'thick' },
    };
    // Totals of Totals Row
    worksheet.getCell(`G${row}`).value = { formula: `SUM(G61:G${row - 1})`, result: 0 };
    worksheet.getCell(`G${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`G${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`G${row}`).border = {
      top: { style: 'thick' },
    };

  }

  getAfterTaxDeductionsData(workbook) {
    let afterTaxDeductions = [];
    workbook.eachSheet((sheet) => {
      if (sheet.name !== "Deduction Totals" && sheet.name !== "Earning Totals") {
        const sheetDate = sheet.getCell("B1").value;
        const employeeLife = sheet.getCell("F20").value;
        const dentalDiscountPlan = sheet.getCell("F21").value;
        const roth401k = sheet.getCell("F22").value;
        const pacAPA = sheet.getCell("F23").value;
        const unionDues = sheet.getCell("F24").value;
        if (sheetDate) {
          afterTaxDeductions.push({
            date: sheetDate,
            employeeLife: employeeLife || 0,
            dentalDiscountPlan: dentalDiscountPlan || 0,
            roth401k: roth401k || 0,
            pacAPA: pacAPA || 0,
            unionDues: unionDues || 0,
          });
        }
      }
    });

    return afterTaxDeductions;
  }

}
