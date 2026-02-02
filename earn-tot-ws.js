import { sum } from 'pdf-lib';

export class EarningsTotalsWorksheet {

  async getAllEarnings(workbook) {
    // Create a new worksheet for totals
    let worksheet = workbook.getWorksheet("Earning Totals");
    if (!worksheet) {
      worksheet = workbook.addWorksheet("Earning Totals");
    } else {
      workbook.removeWorksheet(worksheet.id);
      worksheet = workbook.addWorksheet("Earning Totals");
    }

    for (let col of ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"]) {
      worksheet.getColumn(col).width = 20
    }

    for (let cell of ["A1","A2","B2","C2","D2","E2","F2","G2",
      "H2","I2","J2","A31","A32","B32","C32","D32","E32","F32",
      "G32","H32","I32","J32","K32","L32","M32","N32","O32","L1",
      "L2","M2","N2","O2"
    ]) {
      worksheet.getCell(cell).alignment = { horizontal: 'center' };
      if (cell === "A1" || cell === "A31" || cell === "L1") {
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

    worksheet.mergeCells("A1:J1");
    worksheet.mergeCells("A31:O31");
    worksheet.mergeCells("L1:O1");
    worksheet.getRow(1).height = 30;
    worksheet.getRow(31).height = 30;

    this.addEarningsTop(workbook, worksheet);
    this.addEarningsBottom(workbook, worksheet);
    this.addTotalEarnings(workbook, worksheet);

    return workbook;
  }

  async addEarningsTop(workbook, worksheet) {
    const earningsTop = this.getEarningsTopData(workbook);
    worksheet.getCell("A1").value = "Earnings Pay Month Totals";
    worksheet.getCell("A1").font = { bold: true, size: 14 };
    worksheet.getCell("A1").alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell("A2").value = "Date";
    worksheet.getCell("A2").font = { bold: true };
    worksheet.getCell("B2").value = "Crew Advance";
    worksheet.getCell("B2").font = { bold: true };
    worksheet.getCell("C2").value = "PLT Exp D Taxable";
    worksheet.getCell("C2").font = { bold: true };
    worksheet.getCell("D2").value = "PLT Exp Adj Taxable";
    worksheet.getCell("D2").font = { bold: true };
    worksheet.getCell("E2").value = "PLT Exp I Taxable";
    worksheet.getCell("E2").font = { bold: true };
    worksheet.getCell("F2").value = "AAG Profit Sharing";
    worksheet.getCell("F2").font = { bold: true };
    worksheet.getCell("G2").value = "PLT Exp D Non Taxable";
    worksheet.getCell("G2").font = { bold: true };
    worksheet.getCell("H2").value = "PLT Exp Adj Non Taxable";
    worksheet.getCell("H2").font = { bold: true };
    worksheet.getCell("I2").value = "PLT Exp I Non Taxable";
    worksheet.getCell("I2").font = { bold: true };
    worksheet.getCell("J2").value = "Total";
    worksheet.getCell("J2").font = { bold: true };

    for (let i = 0; i < earningsTop.length; i++) {
      const row = i + 3;
      worksheet.getCell(`A${row}`).value = earningsTop[i].date;
      worksheet.getCell(`A${row}`).alignment = { horizontal: 'center' };
      worksheet.getCell(`A${row}`).numFmt = "ddMMMyyyy";
      worksheet.getCell(`A${row}`).font = { bold: true };
      worksheet.getCell(`B${row}`).value = earningsTop[i].crewAdvance;
      worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`C${row}`).value = earningsTop[i].pltExpDTaxable;
      worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`D${row}`).value = earningsTop[i].pltExpAdjTaxable;
      worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`E${row}`).value = earningsTop[i].pltExpITaxable;
      worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`F${row}`).value = earningsTop[i].aagProfitSharing;
      worksheet.getCell(`F${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`G${row}`).value = earningsTop[i].pltExpDNonTaxable;
      worksheet.getCell(`G${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`H${row}`).value = earningsTop[i].pltExpAdjNonTaxable;
      worksheet.getCell(`H${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`I${row}`).value = earningsTop[i].pltExpINonTaxable;
      worksheet.getCell(`I${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`J${row}`).value = {
        formula: `SUM(B${row}:I${row})`,
        result: 0
      };
      worksheet.getCell(`J${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`J${row}`).font = { bold: true };
    }

    // Add totals row
    const row = earningsTop.length + 3;
    worksheet.getCell(`A${row}`).value = "Totals";
    worksheet.getCell(`A${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`A${row}`).alignment = { horizontal: 'left' };
    worksheet.getCell(`A${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`B${row}`).value = { formula: `SUM(B3:B${row - 1})`, result: 0 };
    worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`B${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`B${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`C${row}`).value = { formula: `SUM(C3:C${row - 1})`, result: 0 };
    worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`C${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`C${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`D${row}`).value = { formula: `SUM(D3:D${row - 1})`, result: 0 };
    worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`D${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`D${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`E${row}`).value = { formula: `SUM(E3:E${row - 1})`, result: 0 };
    worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`E${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`E${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`F${row}`).value = { formula: `SUM(F3:F${row - 1})`, result: 0 };
    worksheet.getCell(`F${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`F${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`F${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`G${row}`).value = { formula: `SUM(G3:G${row - 1})`, result: 0 };
    worksheet.getCell(`G${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`G${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`G${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`H${row}`).value = { formula: `SUM(H3:H${row - 1})`, result: 0 };
    worksheet.getCell(`H${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`H${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`H${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`I${row}`).value = { formula: `SUM(I3:I${row - 1})`, result: 0 };
    worksheet.getCell(`I${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`I${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`I${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`J${row}`).value = { formula: `SUM(J3:J${row - 1})`, result: 0 };
    worksheet.getCell(`J${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`J${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`J${row}`).border = {
      top: { style: 'thick' },
    };
  }

  getEarningsTopData(workbook) {
    let earningsTop = [];
    workbook.eachSheet((sheet) => {
      if (sheet.name !== "Deduction Totals" && sheet.name !== "Earning Totals") {
        const sheetDate = sheet.getCell("B1").value;
        const crewAdvance = sheet.getCell("D8").value;
        const pltExpDTaxable = sheet.getCell("D9").value;
        const pltExpAdjTaxable = sheet.getCell("D10").value;
        const pltExpITaxable = sheet.getCell("D11").value;
        const aagProfitSharing = sheet.getCell("D12").value;
        const pltExpDNonTaxable = sheet.getCell("D13").value;
        const pltExpAdjNonTaxable = sheet.getCell("D14").value;
        const pltExpINonTaxable = sheet.getCell("D15").value;
        if (sheetDate) {
          earningsTop.push({
            date: sheetDate,
            crewAdvance: crewAdvance || 0,
            pltExpDTaxable: pltExpDTaxable || 0,
            pltExpAdjTaxable: pltExpAdjTaxable || 0,
            pltExpITaxable: pltExpITaxable || 0,
            aagProfitSharing: aagProfitSharing || 0,
            pltExpDNonTaxable: pltExpDNonTaxable || 0,
            pltExpAdjNonTaxable: pltExpAdjNonTaxable || 0,
            pltExpINonTaxable: pltExpINonTaxable || 0
          });
        }

      }
    });

    return earningsTop;
  }

  async addEarningsBottom(workbook, worksheet) {
    let earningsBottom = this.getEarningsBottomData(workbook);
    worksheet.getCell("A31").value = "Earnings Pay Period Totals";
    worksheet.getCell("A31").font = { bold: true, size: 14 };
    worksheet.getCell("A31").alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell("A32").value = "Date";
    worksheet.getCell("A32").font = { bold: true };
    worksheet.getCell("B32").value = "Operational Pay";
    worksheet.getCell("B32").font = { bold: true };
    worksheet.getCell("C32").value = "FLT Training Pay";
    worksheet.getCell("C32").font = { bold: true };
    worksheet.getCell("D32").value = "SIT Time";
    worksheet.getCell("D32").font = { bold: true };
    worksheet.getCell("E32").value = "Pay Above Guarantee";
    worksheet.getCell("E32").font = { bold: true };
    worksheet.getCell("F32").value = "RA Prem";
    worksheet.getCell("F32").font = { bold: true };
    worksheet.getCell("G32").value = "Min Guar Adj";
    worksheet.getCell("G32").font = { bold: true };
    worksheet.getCell("H32").value = "Intl Override";
    worksheet.getCell("H32").font = { bold: true };
    worksheet.getCell("I32").value = "Dist Learning";
    worksheet.getCell("I32").font = { bold: true };
    worksheet.getCell("J32").value = "Union Paid Leave";
    worksheet.getCell("J32").font = { bold: true };
    worksheet.getCell("K32").value = "Prem Incentive Pay";
    worksheet.getCell("K32").font = { bold: true };
    worksheet.getCell("L32").value = "FLT Vacation Pay";
    worksheet.getCell("L32").font = { bold: true };
    worksheet.getCell("M32").value = "Sick Pay";
    worksheet.getCell("M32").font = { bold: true };
    worksheet.getCell("N32").value = "Prior Yr Vac Pay Out";
    worksheet.getCell("N32").font = { bold: true };
    worksheet.getCell("O32").value = "Total";
    worksheet.getCell("O32").font = { bold: true };

    for (let i = 0; i < earningsBottom.length; i++) {
      const row = i + 33;
      worksheet.getCell(`A${row}`).value = earningsBottom[i].date;
      worksheet.getCell(`A${row}`).alignment = { horizontal: 'center' };
      worksheet.getCell(`A${row}`).numFmt = "ddMMMyyyy";
      worksheet.getCell(`A${row}`).font = { bold: true };
      worksheet.getCell(`B${row}`).value = earningsBottom[i].operationalPay;
      worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`C${row}`).value = earningsBottom[i].fltTrainingPay;
      worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`D${row}`).value = earningsBottom[i].sitTime;
      worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`E${row}`).value = earningsBottom[i].payAboveGuarantee;
      worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`F${row}`).value = earningsBottom[i].raPrem;
      worksheet.getCell(`F${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`G${row}`).value = earningsBottom[i].minGuarAdj;
      worksheet.getCell(`G${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`H${row}`).value = earningsBottom[i].intlOverride;
      worksheet.getCell(`H${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`I${row}`).value = earningsBottom[i].distLearning;
      worksheet.getCell(`I${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`J${row}`).value = earningsBottom[i].unionPaidLeave;
      worksheet.getCell(`J${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`K${row}`).value = earningsBottom[i].premIncentivePay;
      worksheet.getCell(`K${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`L${row}`).value = earningsBottom[i].fltVacationPay;
      worksheet.getCell(`L${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`M${row}`).value = earningsBottom[i].sickPay;
      worksheet.getCell(`M${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`N${row}`).value = earningsBottom[i].priorYrVacPayOut;
      worksheet.getCell(`N${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`O${row}`).value = {
        formula: `SUM(B${row}:N${row})`,
        result: 0
      };
      worksheet.getCell(`O${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`O${row}`).font = { bold: true };
    }

    const row = earningsBottom.length + 33;
    worksheet.getCell(`A${row}`).value = "Total YTD";
    worksheet.getCell(`A${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`A${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`A${row}`).alignment = { horizontal: 'left' };
    worksheet.getCell(`A${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`B${row}`).value = { formula: `SUM(B33:B${row - 1})`, result: 0 };
    worksheet.getCell(`B${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`B${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`B${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`C${row}`).value = { formula: `SUM(C33:C${row - 1})`, result: 0 };
    worksheet.getCell(`C${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`C${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`C${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`D${row}`).value = { formula: `SUM(D33:D${row - 1})`, result: 0 };
    worksheet.getCell(`D${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`D${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`D${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`E${row}`).value = { formula: `SUM(E33:E${row - 1})`, result: 0 };
    worksheet.getCell(`E${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`E${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`E${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`F${row}`).value = { formula: `SUM(F33:F${row - 1})`, result: 0 };
    worksheet.getCell(`F${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`F${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`F${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`G${row}`).value = { formula: `SUM(G33:G${row - 1})`, result: 0 };
    worksheet.getCell(`G${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`G${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`G${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`H${row}`).value = { formula: `SUM(H33:H${row - 1})`, result: 0 };
    worksheet.getCell(`H${row }`).numFmt = "$#,##0.00";
    worksheet.getCell(`H${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`H${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`I${row}`).value = { formula: `SUM(I33:I${row - 1})`, result: 0 };
    worksheet.getCell(`I${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`I${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`I${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`J${row}`).value = { formula: `SUM(J33:J${row - 1})`, result: 0 };
    worksheet.getCell(`J${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`J${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`J${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`K${row}`).value = { formula: `SUM(K33:K${row - 1})`, result: 0 };
    worksheet.getCell(`K${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`K${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`K${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`L${row}`).value = { formula: `SUM(L33:L${row - 1})`, result: 0 };
    worksheet.getCell(`L${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`L${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`L${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`M${row}`).value = { formula: `SUM(M33:M${row - 1})`, result: 0 };
    worksheet.getCell(`M${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`M${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`M${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`N${row}`).value = { formula: `SUM(N33:N${row - 1})`, result: 0 };
    worksheet.getCell(`N${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`N${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`N${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`O${row}`).value = { formula: `SUM(O33:O${row - 1})`, result: 0 };
    worksheet.getCell(`O${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`O${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`O${row}`).border = {
      top: { style: 'thick' },
    };
  }

  getEarningsBottomData(workbook) {
    let earningsBottom = [];
    workbook.eachSheet((sheet) => {
      if (sheet.name !== "Deduction Totals" && sheet.name !== "Earning Totals") {
        const sheetDate = sheet.getCell("B1").value;
        const operationalPay = sheet.getCell("D18").value;
        const fltTrainingPay = sheet.getCell("D19").value;
        const sitTime = sheet.getCell("D20").value;
        const payAboveGuarantee = sheet.getCell("D21").value;
        const raPrem = sheet.getCell("D22").value;
        const minGuarAdj = sheet.getCell("D23").value;
        const intlOverride = sheet.getCell("D24").value;
        const distLearning = sheet.getCell("D25").value;
        const unionPaidLeave = sheet.getCell("D26").value;
        const premIncentivePay = sheet.getCell("D27").value;
        const fltVacationPay = sheet.getCell("D28").value;
        const sickPay = sheet.getCell("D29").value;
        const priorYrVacPayOut = sheet.getCell("D30").value;
        if (sheetDate) {
          earningsBottom.push({
            date: sheetDate,
            operationalPay: operationalPay || 0,
            fltTrainingPay: fltTrainingPay || 0,
            sitTime: sitTime || 0,
            payAboveGuarantee: payAboveGuarantee || 0,
            raPrem: raPrem || 0,
            minGuarAdj: minGuarAdj || 0,
            intlOverride: intlOverride || 0,
            distLearning: distLearning || 0,
            unionPaidLeave: unionPaidLeave || 0,
            premIncentivePay: premIncentivePay || 0,
            fltVacationPay: fltVacationPay || 0,
            sickPay: sickPay || 0,
            priorYrVacPayOut: priorYrVacPayOut || 0
          });
        }

      }
    });

    return earningsBottom;
  }

  async addTotalEarnings(workbook, worksheet) {
    let earningsTop = this.getEarningsTopData(workbook);
    let earningsBottom = this.getEarningsBottomData(workbook);

    worksheet.getCell("L1").value = "Earnings Totals";
    worksheet.getCell("L1").font = { bold: true, size: 14 };
    worksheet.getCell("L1").alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell("L2").value = "Date";
    worksheet.getCell("L2").font = { bold: true };
    worksheet.getCell("M2").value = "Pay Month";
    worksheet.getCell("M2").font = { bold: true };
    worksheet.getCell("N2").value = "Pay Period";
    worksheet.getCell("N2").font = { bold: true };
    worksheet.getCell("O2").value = "Earnings Total";
    worksheet.getCell("O2").font = { bold: true };

    for (let i = 0; i < earningsTop.length; i++) {
      const row = i + 3;
      
      // Sum the current entry's values (excluding date)
      const sumEarningsTop = Object.entries(earningsTop[i])
        .filter(([key, _]) => key !== 'date')
        .reduce((sum, [_, value]) => sum + value, 0);
      
      const sumEarningsBottom = Object.entries(earningsBottom[i])
        .filter(([key, _]) => key !== 'date')
        .reduce((sum, [_, value]) => sum + value, 0);

      worksheet.getCell(`L${row}`).value = earningsTop[i].date;
      worksheet.getCell(`L${row}`).alignment = { horizontal: 'center' };
      worksheet.getCell(`L${row}`).numFmt = "ddMMMyyyy";
      worksheet.getCell(`L${row}`).font = { bold: true };
      worksheet.getCell(`M${row}`).value = sumEarningsTop;
      worksheet.getCell(`M${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`N${row}`).value = sumEarningsBottom;
      worksheet.getCell(`N${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`O${row}`).value = {
        formula: `SUM(M${row}:N${row})`,
        result: 0
      };
      worksheet.getCell(`O${row}`).numFmt = "$#,##0.00";
      worksheet.getCell(`O${row}`).font = { bold: true };
    }

    // Add totals row
    const row = earningsTop.length + 3;
    worksheet.getCell(`L${row}`).value = "Totals";
    worksheet.getCell(`L${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`L${row}`).alignment = { horizontal: 'left' };
    worksheet.getCell(`L${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`M${row}`).value = { formula: `SUM(M3:M${row - 1})`, result: 0 };
    worksheet.getCell(`M${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`M${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`M${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`N${row}`).value = { formula: `SUM(N3:N${row - 1})`, result: 0 };
    worksheet.getCell(`N${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`N${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`N${row}`).border = {
      top: { style: 'thick' },
    };
    worksheet.getCell(`O${row}`).value = { formula: `SUM(O3:O${row - 1})`, result: 0 };
    worksheet.getCell(`O${row}`).numFmt = "$#,##0.00";
    worksheet.getCell(`O${row}`).font = { bold: true, size: 14 };
    worksheet.getCell(`O${row}`).border = {
      top: { style: 'thick' },
    };
  }


    // sumEarningsTop = earningsTop.reduce((acc, entry) => {
    //   return acc + Object.entries(entry)
    //   .filter(([key, _]) => key !== 'date')
    //   .reduce((sum, [_, value]) => sum + value, 0);
    // }, 0);
    // sumEarningsBottom = earningsBottom.reduce((acc, entry) => {
    //   return acc + Object.entries(entry)
    //   .filter(([key, _]) => key !== 'date')
    //   .reduce((sum, [_, value]) => sum + value, 0);
    // }, 0);
}