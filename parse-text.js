export class TextParser {
  
  parsePDFInfoData(extractedText) {
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

  parsePDFSummaryData(extractedText) {
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

  parsePDFEarningsTopData(extractedText) {
    const crewAdvanceRegex =
      /Advance\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*((?:-\d+[\d,]+(?:\.\d+)?|\d+[\d,]+(?:\.\d+)?))/i;
    const pltExpDtaxRegex =
      /D\s*Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const pltExpADJtax =
      /ADJ\s*Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const pltExpItax =
      /I\s*Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const AagProfitSharingRegex =
      /Profit\s*Sharing\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const pltExpDnonTaxRegex =
      /D\s*Non-Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const pltExpAdjnonTaxRegex =
      /ADJ\s*Non-Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*((?:-[\d,]+(?:\.\d+)?)|([\d,]+(?:\.\d+)?))/i;
    const pltExpInonTaxRegex =
      /I\s*Non-Taxable\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;

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
      AagProfitSharingRate: AagProfitSharingMatch
        ? AagProfitSharingMatch[1]
        : "0",
      AagProfitSharingHours: AagProfitSharingMatch
        ? AagProfitSharingMatch[2]
        : "0",
      AagProfitSharingCurrent: AagProfitSharingMatch
        ? AagProfitSharingMatch[3]
        : "0",
      pltExpDnonTaxRate: pltExpDnonTaxMatch ? pltExpDnonTaxMatch[1] : "0",
      pltExpDnonTaxHours: pltExpDnonTaxMatch ? pltExpDnonTaxMatch[2] : "0",
      pltExpDnonTaxCurrent: pltExpDnonTaxMatch ? pltExpDnonTaxMatch[3] : "0",
      pltExpAdjnonTaxRate: pltExpAdjnonTaxMatch ? pltExpAdjnonTaxMatch[1] : "0",
      pltExpAdjnonTaxHours: pltExpAdjnonTaxMatch
        ? pltExpAdjnonTaxMatch[2]
        : "0",
      pltExpAdjnonTaxCurrent: pltExpAdjnonTaxMatch
        ? pltExpAdjnonTaxMatch[3]
        : "0",
      pltExpInonTaxRate: pltExpInonTaxMatch ? pltExpInonTaxMatch[1] : "0",
      pltExpInonTaxHours: pltExpInonTaxMatch ? pltExpInonTaxMatch[2] : "0",
      pltExpInonTaxCurrent: pltExpInonTaxMatch ? pltExpInonTaxMatch[3] : "0",
    };
  }

  parsePDFEarningsBottomData(extractedText) {
    const operationalPayRegex =
      /Operational\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const fltTrainingPayRegex =
      /Flight\s*Training\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const sitTime =
      /Sit\s*Time\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const payAbvGuaranteeRsvRegex =
      /Pay\s*Above\s*Guar\s*\(RSV\)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const raPrem =
      /RA\s*PREM\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const minGuaranteeAdjRegex =
      /Min\s*Guarantee\s*Adj\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const intlOverrideRegex =
      /Intl\s*Override\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const distanceLearningRegex =
      /Distance\s*Learning\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const unionPdLeaveRegex =
      /Union\s*PD\s*Union\s*Leave\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const premIncentivePayRegex =
      /Prem\s*Incentive\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const fltVacationPayRegex =
      /Flight\s*Vacation\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const sickPayRegex =
      /Sick\s*Pay\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;
    const priorYearVacPayoutRegex =
      /Prior\s*Year\s*Vacation\s*Pay\s*Out\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)\s*([\d,]+(?:\.\d+)?)/i;

    const operationalPayMatch = extractedText.match(operationalPayRegex);
    const fltTrainingPayMatch = extractedText.match(fltTrainingPayRegex);
    const sitTimeMatch = extractedText.match(sitTime);
    const payAbvGuaranteeRsvMatch = extractedText.match(
      payAbvGuaranteeRsvRegex
    );
    const raPremMatch = extractedText.match(raPrem);
    const minGuaranteeAdjMatch = extractedText.match(minGuaranteeAdjRegex);
    const intlOverrideMatch = extractedText.match(intlOverrideRegex);
    const distanceLearningMatch = extractedText.match(distanceLearningRegex);
    const unionPdLeaveMatch = extractedText.match(unionPdLeaveRegex);
    const premIncentivePayMatch = extractedText.match(premIncentivePayRegex);
    const fltVacationPayMatch = extractedText.match(fltVacationPayRegex);
    const sickPayMatch = extractedText.match(sickPayRegex);
    const priorYearVacPayoutMatch = extractedText.match(
      priorYearVacPayoutRegex
    );

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
      payAbvGuaranteeRsvRate: payAbvGuaranteeRsvMatch
        ? payAbvGuaranteeRsvMatch[1]
        : "0",
      payAbvGuaranteeRsvHours: payAbvGuaranteeRsvMatch
        ? payAbvGuaranteeRsvMatch[2]
        : "0",
      payAbvGuaranteeRsvCurrent: payAbvGuaranteeRsvMatch
        ? payAbvGuaranteeRsvMatch[3]
        : "0",
      raPremRate: raPremMatch ? raPremMatch[1] : "0",
      raPremHours: raPremMatch ? raPremMatch[2] : "0",
      raPremCurrent: raPremMatch ? raPremMatch[3] : "0",
      minGuaranteeAdjRate: minGuaranteeAdjMatch ? minGuaranteeAdjMatch[1] : "0",
      minGuaranteeAdjHours: minGuaranteeAdjMatch
        ? minGuaranteeAdjMatch[2]
        : "0",
      minGuaranteeAdjCurrent: minGuaranteeAdjMatch
        ? minGuaranteeAdjMatch[3]
        : "0",
      intlOverrideRate: intlOverrideMatch ? intlOverrideMatch[1] : "0",
      intlOverrideHours: intlOverrideMatch ? intlOverrideMatch[2] : "0",
      intlOverrideCurrent: intlOverrideMatch ? intlOverrideMatch[3] : "0",
      distanceLearningRate: distanceLearningMatch
        ? distanceLearningMatch[1]
        : "0",
      distanceLearningHours: distanceLearningMatch
        ? distanceLearningMatch[2]
        : "0",
      distanceLearningCurrent: distanceLearningMatch
        ? distanceLearningMatch[3]
        : "0",
      unionPdLeaveRate: unionPdLeaveMatch ? unionPdLeaveMatch[1] : "0",
      unionPdLeaveHours: unionPdLeaveMatch ? unionPdLeaveMatch[2] : "0",
      unionPdLeaveCurrent: unionPdLeaveMatch ? unionPdLeaveMatch[3] : "0",
      premIncentivePayRate: premIncentivePayMatch
        ? premIncentivePayMatch[1]
        : "0",
      premIncentivePayHours: premIncentivePayMatch
        ? premIncentivePayMatch[2]
        : "0",
      premIncentivePayCurrent: premIncentivePayMatch
        ? premIncentivePayMatch[3]
        : "0",
      fltVacationPayRate: fltVacationPayMatch ? fltVacationPayMatch[1] : "0",
      fltVacationPayHours: fltVacationPayMatch ? fltVacationPayMatch[2] : "0",
      fltVacationPayCurrent: fltVacationPayMatch ? fltVacationPayMatch[3] : "0",
      sickPayRate: sickPayMatch ? sickPayMatch[1] : "0",
      sickPayHours: sickPayMatch ? sickPayMatch[2] : "0",
      sickPayCurrent: sickPayMatch ? sickPayMatch[3] : "0",
      priorYearVacPayoutRate: priorYearVacPayoutMatch
        ? priorYearVacPayoutMatch[1]
        : "0",
      priorYearVacPayoutHours: priorYearVacPayoutMatch
        ? priorYearVacPayoutMatch[2]
        : "0",
      priorYearVacPayoutCurrent: priorYearVacPayoutMatch
        ? priorYearVacPayoutMatch[3]
        : "0",
    };
  }

  parsePDFPreTaxDeductionsData(extractedText) {
    const medicalCoverageRegex = /Medical\s*Coverage\s*([\d,]+(?:\.\d+)?)/i;
    const dentalCoverageRegex = /Dental\s*Coverage\s*([\d,]+(?:\.\d+)?)/i;
    const visionCoverageRegex = /Vision\s*Coverage\s*([\d,]+(?:\.\d+)?)/i;
    const accidentInsPreTaxRegex =
      /Accident\s*Ins\s*Pre-tax\s*([\d,]+(?:\.\d+)?)/i;
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
      accidentInsPreTax: accidentInsPreTaxMatch
        ? accidentInsPreTaxMatch[1]
        : "0",
      _401k: _401kMatch ? _401kMatch[1] : "0",
    };
  }

  parsePDFTaxDeductionsData(extractedText) {
    const withholdingTaxRegex = /Withholding\s*Tax\s*([\d,]+(?:\.\d+)?)/i;
    const socialSecurityTaxRegex =
      /Social\s*Security\s*Tax\s*([\d,]+(?:\.\d+)?)/i;
    const medicareTaxRegex = /Medicare\s*Tax\s*([\d,]+(?:\.\d+)?)/i;

    const withholdingTaxMatch = extractedText.match(withholdingTaxRegex);
    const socialSecurityTaxMatch = extractedText.match(socialSecurityTaxRegex);
    const medicareTaxMatch = extractedText.match(medicareTaxRegex);
    return {
      withholdingTax: withholdingTaxMatch ? withholdingTaxMatch[1] : "0",
      socialSecurityTax: socialSecurityTaxMatch
        ? socialSecurityTaxMatch[1]
        : "0",
      medicareTax: medicareTaxMatch ? medicareTaxMatch[1] : "0",
    };
  }

  parsePDFAfterTaxDeductionsData(extractedText) {
    const employeeLifeRegex = /Employee\s*Life\s*([\d,]+(?:\.\d+)?)/i;
    const detalDiscountPlanRegex =
      /Dental\s*Discount\s*Plan\s*([\d,]+(?:\.\d+)?)/i;
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
      dentalDiscountPlan: dentalDiscountPlanMatch
        ? dentalDiscountPlanMatch[1]
        : "0",
      roth401k: roth401kMatch ? roth401kMatch[1] : "0",
      pacAPA: pacAPAMatch ? pacAPAMatch[1] : "0",
      unionDues: unionDuesMatch ? unionDuesMatch[1] : "0",
    };
  }

  parseCompanyContributionsData(extractedText) {
    const _401kCompanyContributionRegex =
      /401k\s*Company\s*Contrib\.+\s*([\d,]+(?:\.\d+)?)/i;
    const groupTermLifeRegex = /Group\s*Term\s*Life\s*([\d,]+(?:\.\d+)?)/i;

    const _401kCompanyContributionMatch = extractedText.match(
      _401kCompanyContributionRegex
    );
    const groupTermLifeMatch = extractedText.match(groupTermLifeRegex);

    return {
      _401kCompanyContribution: _401kCompanyContributionMatch
        ? _401kCompanyContributionMatch[1]
        : "0",
      groupTermLife: groupTermLifeMatch ? groupTermLifeMatch[1] : "0",
    };
  }

  parsePDFTaxableEarningsData(extractedText) {
    const withHoldingTaxRegex =
      /Withholding\s*Tax\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)/g;
    const socialSecurityTaxRegex =
      /Social\s*Security\s*Tax\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)/g;
    const medicareTaxRegex = /Medicare\s*Tax\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)/g;

    const withHoldingTaxMatches = [
      ...extractedText.matchAll(withHoldingTaxRegex),
    ];
    const socialSecurityTaxMatches = [
      ...extractedText.matchAll(socialSecurityTaxRegex),
    ];
    const medicareTaxMatches = [...extractedText.matchAll(medicareTaxRegex)];
    const withHoldingTaxMatch = withHoldingTaxMatches[1]
      ? withHoldingTaxMatches[1][1]
      : "0";
    const socialSecurityTaxMatch = socialSecurityTaxMatches[1]
      ? socialSecurityTaxMatches[1][1]
      : "0";
    const medicareTaxMatch = medicareTaxMatches[1]
      ? medicareTaxMatches[1][1]
      : "0";

    return {
      withHoldingTaxEarnings: withHoldingTaxMatch,
      socialSecurityTaxEarnings: socialSecurityTaxMatch,
      medicareTaxEarnings: medicareTaxMatch,
    };
  }
}
