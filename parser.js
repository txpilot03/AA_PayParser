import PDFParser from 'pdf2json';

/**
 * Hybrid PDF Parser - Combines position-based extraction with smart parsing
 * This approach is more reliable than pure regex and easier to maintain
 * 
 * Note: Uses pdf2json for Node.js compatibility (already in your dependencies)
 */
export class HybridPDFParser {
  constructor() {
    // State tracking for multi-line parsing
    this.previousLine = null;
    this.waitingForHeaderValues = false;
    
    // Track which taxes have been parsed (to avoid second occurrence)
    this.taxesParsed = {
      withholdingTax: false,
      socialSecurityTax: false,
      medicareTax: false
    };
    
    // Track summary parsing state
    this.summaryState = {
      seenHeader: false,
      seenCurrent: false
    };
    
    // Define known field patterns for classification
    this.fieldPatterns = {
      // Header info patterns
      payPeriod: ['1-800-447-2000', 'Pay Period'],
      seniorityYear: ['Effective', 'Seniority'],
      group: ['Effective', 'Group'],
      hourlyRate: ['Effective', 'Rate'],
      effectivePeriod: ['Effective', 'EffectivePeriod'],
      
      // Earnings patterns
      earnings: [
        'Crew Advance', 'Operational Pay', 'Flight Training Pay',
        'Sit Time', 'Pay Above Guar', 'RA PREM', 'Min Guarantee Adj',
        'Intl Override', 'Distance Learning', 'Union PD Union Leave',
        'Prem Incentive Pay', 'Flight Vacation Pay', 'Sick Pay',
        'Prior Year Vacation Pay Out', 'Pilot Exp', 'Taxable', 'Non-Taxable',
        'AAG Profit Sharing'
      ],
      
      // Deductions patterns
      preTaxDeductions: [
        'Medical Coverage', 'Dental Coverage', 'Vision Coverage',
        'Accident Ins Pre-tax', '401k'
      ],
      
      taxes: [
        'Withholding Tax', 'Social Security Tax', 'Medicare Tax'
      ],
      
      afterTaxDeductions: [
        'Employee Life', 'Dental Discount Plan', 'Roth 401k',
        'PAC - APA', 'Union Dues - APA'
      ],
      
      companyContributions: [
        '401k Company Contrib', 'Group Term Life'
      ],
      
      summary: [
        'Gross', 'Pre-Tax Deduct', 'Taxes', 'After Tax Deduct', 'Net Pay'
      ]
    };
  }

  /**
   * Main parsing method - returns structured data
   */
  async parse(fileBuffer) {
    console.log(fileBuffer);
    const structuredContent = await this.extractStructuredContent(fileBuffer);
    
    const data = {
      header: {
        regularPayRoll: '',
        payPeriod: '',
        seniorityYear: '0',
        group: '',
        hourlyRate: '0',
        effectivePeriod: ''
      },
      earnings: {
        crewAdvance: { rate: '0', hours: '0', current: '0' },
        pltExpDtax: { rate: '0', hours: '0', current: '0' },
        pltExpADJtax: { rate: '0', hours: '0', current: '0' },
        pltExpItax: { rate: '0', hours: '0', current: '0' },
        AagProfitSharing: { rate: '0', hours: '0', current: '0' },
        pltExpDnonTax: { rate: '0', hours: '0', current: '0' },
        pltExpAdjnonTax: { rate: '0', hours: '0', current: '0' },
        pltExpInonTax: { rate: '0', hours: '0', current: '0' },
        operationalPay: { rate: '0', hours: '0', current: '0' },
        fltTrainingPay: { rate: '0', hours: '0', current: '0' },
        sitTime: { rate: '0', hours: '0', current: '0' },
        payAbvGuaranteeRsv: { rate: '0', hours: '0', current: '0' },
        raPrem: { rate: '0', hours: '0', current: '0' },
        minGuaranteeAdj: { rate: '0', hours: '0', current: '0' },
        intlOverride: { rate: '0', hours: '0', current: '0' },
        distanceLearning: { rate: '0', hours: '0', current: '0' },
        unionPdLeave: { rate: '0', hours: '0', current: '0' },
        premIncentivePay: { rate: '0', hours: '0', current: '0' },
        fltVacationPay: { rate: '0', hours: '0', current: '0' },
        sickPay: { rate: '0', hours: '0', current: '0' },
        priorYearVacPayout: { rate: '0', hours: '0', current: '0' },
        earningsTotal: { current: '0' }
      },
      deductions: {
        preTax: {
          medicalCoverage: '0',
          dentalCoverage: '0',
          visionCoverage: '0',
          accidentInsPreTax: '0',
          _401k: '0'
        },
        taxes: {
          withholdingTax: '0',
          socialSecurityTax: '0',
          medicareTax: '0'
        },
        afterTax: {
          employeeLife: '0',
          dentalDiscountPlan: '0',
          roth401k: '0',
          pacAPA: '0',
          unionDues: '0'
        }
      },
      companyContributions: {
        _401kCompanyContribution: '0',
        groupTermLife: '0'
      },
      taxableEarnings: {
        withHoldingTaxEarnings: '0',
        socialSecurityTaxEarnings: '0',
        medicareTaxEarnings: '0'
      },
      summary: {
        gross: '0',
        preTaxDeduct: '0',
        taxes: '0',
        afterTaxDeduct: '0',
        netPay: '0'
      }
    };

    // Process only page 1 (first page, index 0)
    if (structuredContent.length > 0) {
      const page = structuredContent[0];
      this.previousLine = null;
      this.waitingForHeaderValues = false;
      
      // Reset tax parsing flags
      this.taxesCurrentParsed = {
        withholdingTax: false,
        socialSecurityTax: false,
        medicareTax: false
      };      
      // Reset summary parsing state
      this.summaryState = {
        seenHeader: false,
        seenCurrent: false
      };
      this.taxesTotalParsed = {
        withholdingTax: false,
        socialSecurityTax: false,
        medicareTax: false
      }
      
      for (let i = 0; i < page.length; i++) {
        const line = page[i];
        const nextLine = i < page.length - 1 ? page[i + 1] : null;
        this.classifyAndParseLine(line, nextLine, data);
        this.previousLine = line;
      }
    }
    return data;
  }

  /**
   * Extract PDF content with position information
   * Uses pdf2json for Node.js compatibility
   */
  async extractStructuredContent(fileBuffer) {
    return new Promise((resolve, reject) => {
      // Convert ArrayBuffer to Buffer if needed (from Electron IPC)
      let buffer = fileBuffer;
      if (fileBuffer instanceof ArrayBuffer) {
        buffer = Buffer.from(fileBuffer);
      } else if (fileBuffer.buffer instanceof ArrayBuffer) {
        // Handle Uint8Array
        buffer = Buffer.from(fileBuffer.buffer);
      }
      
      const pdfParser = new PDFParser();
      
      pdfParser.on('pdfParser_dataError', errData => {
        reject(new Error(errData.parserError));
      });
      
      pdfParser.on('pdfParser_dataReady', pdfData => {
        try {
          const pages = [];
          
          // Process each page
          pdfData.Pages.forEach(page => {
            const lineMap = new Map();
            
            // Group texts by Y position to create lines
            page.Texts.forEach(text => {
              const y = Math.round(text.y * 100); // Round Y to group nearby texts
              const x = text.x;
              
              if (!lineMap.has(y)) {
                lineMap.set(y, []);
              }
              
              // Decode text
              const decodedText = decodeURIComponent(text.R[0].T);
              
              lineMap.get(y).push({
                x: x,
                text: decodedText,
                width: text.w
              });
            });
            
            // Convert map to sorted lines array
            const pageLines = Array.from(lineMap.entries())
              .sort((a, b) => a[0] - b[0]) // Sort by Y position (top to bottom)
              .map(([y, items]) => ({
                y: y,
                items: items.sort((a, b) => a.x - b.x) // Sort items left to right
              }));
            
            pages.push(pageLines);
          });
          
          resolve(pages);
          
        } catch (error) {
          reject(error);
        }
      });
      
      // Parse the buffer (now guaranteed to be a Node.js Buffer)
      pdfParser.parseBuffer(buffer);
    });
  }

  /**
   * Group text items into lines based on Y position
   */
  groupTextIntoLines(items) {
    const lineThreshold = 2; // Y position tolerance
    const lines = [];

    items.forEach(item => {
      const y = Math.round(item.transform[5]);
      const x = item.transform[4];

      // Find existing line or create new one
      let line = lines.find(l => Math.abs(l.y - y) < lineThreshold);
      if (!line) {
        line = { y, items: [] };
        lines.push(line);
      }

      line.items.push({
        x,
        text: item.str,
        width: item.width
      });
    });

    // Sort lines top to bottom
    lines.sort((a, b) => b.y - a.y);

    // Sort items in each line left to right
    lines.forEach(line => {
      line.items.sort((a, b) => a.x - b.x);
    });

    return lines;
  }

  /**
   * Classify line and extract data
   * Now supports multi-line format where labels are on one line and values on the next
   */
  classifyAndParseLine(line, nextLine, data) {
    const lineText = line.items.map(item => item.text).join(' ');
    //console.log('Processing Line:', lineText);
    
    // Check for multi-line header format: "Rate: Sen Grp Rate Effective"
    if (lineText.includes('Rate:') && lineText.includes('Sen') && lineText.includes('Grp')) {
      this.waitingForHeaderValues = true;
      return; // Values are on the next line
    }
    // Check for multi-line Pay Period format
    if (lineText.includes('Pay') && lineText.includes('Period') && nextLine) {
      const nextLineText = nextLine.items.map(item => item.text).join(' ');
      //console.log('Pay Period Line:', nextLineText);
      
      // Look for date range pattern like "12/17/2024 - 12/31/2024"
      const dateRangeMatch = nextLineText.match(/([\d]{1,2}\/[\d]{1,2}\/[\d]{4})\s*-\s*([\d]{1,2}\/[\d]{1,2}\/[\d]{4})/);
      if (dateRangeMatch) {
      data.header.payPeriod = `${dateRangeMatch[1]} - ${dateRangeMatch[2]}`;
      //console.log('Found Pay Period:', data.header.payPeriod);
      }
      return;
    }
    
    // If we're waiting for header values, parse them from the current line
    if (this.waitingForHeaderValues && nextLine) {
      this.waitingForHeaderValues = false;
      // Extract seniority, group, and rate from the NEXT line
      const nextLineText = nextLine.items.map(item => item.text).join(' ');
      //console.log('Header Values Line:', nextLineText);
      
      const values = nextLineText.trim().split(/\s+/);
      //console.log('Header Values:', values);
      
      // Try to find numeric seniority (usually 2 digits)
      let senIdx = values.findIndex(v => /^\d{2}$/.test(v));
      if (senIdx !== -1) {
        data.header.seniorityYear = values[senIdx];
      }
      
      // Try to find group (usually Roman numerals or letters like II, IV, I, etc.)
      let grpIdx = values.findIndex(v => /^[IVX]+$|^[A-Z]{1,2}$/.test(v) && v !== values[senIdx]);
      if (grpIdx !== -1) {
        data.header.group = values[grpIdx];
      }
      
      // Try to find rate (usually has $ or decimal)
      let rateIdx = values.findIndex(v => /\$|^\d+\.\d{2}$/.test(v));
      if (rateIdx !== -1) {
        data.header.hourlyRate = values[rateIdx].replace('$', '').replace(',', '');
      }

      // Try to find effective period (date with slashes)
      let epIdx = values.findIndex(v => v.includes('/'));
      if (epIdx !== -1) {
        data.header.effectivePeriod = values[epIdx];
      }

      return; 
    }
    
    //! Original single-line format for backward compatibility
    // if (lineText.includes('Effective') && /Effective\s*(\d+)/.test(lineText)) {
    //   const match = lineText.match(/Effective\s*(\d+)\s+([A-Z]{2})\s+\$?([\d,]+\.[\d]{2})/);
    //   if (match) {
    //     data.header.seniorityYear = match[1];
    //     data.header.group = match[2];
    //     data.header.hourlyRate = match[3];
    //   }
    // }

    if (lineText.includes('Regular') && lineText.includes('Payroll')) {
      const match = lineText.match(/([\d]{1,2}\/[\d]{1,2}\/[\d]{4})/);
      if (match) data.header.regularPayRoll = match[0];
    }

    // Earnings - Crew Advance
    if (lineText.includes('Crew') && lineText.includes('Advance')) {
      const values = this.extractNumericValues(line);
      if (values.length >= 3) {
        data.earnings.crewAdvance = {
          rate: values[0],
          hours: values[1],
          current: values[2]
        };
      }
    }

    // Earnings - PLT EXP variations
    // Check Non-Taxable FIRST (more specific), then Taxable (less specific)
    // Check ADJ before D/I (since ADJ contains D)
    if (lineText.includes('PLT') && lineText.includes('EXP')) {
      const values = this.extractNumericValues(line);      
      // Non-Taxable entries (check these first)
      if (lineText.includes('Non-Taxable')) {
        if (lineText.includes('ADJ')) {
          if (values.length >= 3) {
            data.earnings.pltExpAdjnonTax = { rate: values[0], hours: values[1], current: values[2] };
          }
        } else if (lineText.includes(' D ') || lineText.match(/\bD\b/)) {
          if (values.length >= 3) {
            data.earnings.pltExpDnonTax = { rate: values[0], hours: values[1], current: values[2] };
          }
        } else if (lineText.includes(' I ') || lineText.match(/\bI\b/)) {
          if (values.length >= 3) {
            data.earnings.pltExpInonTax = { rate: values[0], hours: values[1], current: values[2] };
          }
        }
      } 
      // Taxable entries (without "Non-")
      else if (lineText.includes('Taxable')) {
        if (lineText.includes('ADJ')) {
          if (values.length >= 3) {
            data.earnings.pltExpADJtax = { rate: values[0], hours: values[1], current: values[2] };
          }
        } else if (lineText.includes(' D ') || lineText.match(/\bD\b/)) {
          if (values.length >= 3) {
            data.earnings.pltExpDtax = { rate: values[0], hours: values[1], current: values[2] };
          }
        } else if (lineText.includes(' I ') || lineText.match(/\bI\b/)) {
          if (values.length >= 3) {
            data.earnings.pltExpItax = { rate: values[0], hours: values[1], current: values[2] };
          }
        }
      }
    }

    // AAG Profit Sharing
    if (lineText.includes('AAG') && lineText.includes('Profit Sharing')) {
      const values = this.extractNumericValues(line);
      if (values.length >= 3) {
        data.earnings.AagProfitSharing = { rate: values[0], hours: values[1], current: values[2] };
      }
    }

    // Other earnings items
    this.parseEarningsLine(lineText, line, data.earnings);

    // Deductions
    this.parseDeductionsLine(lineText, line, data);

    // Summary
    this.parseSummaryLine(lineText, line, data.summary);
  }

  /**
   * Parse earnings lines
   */
  parseEarningsLine(lineText, line, earnings) {
    const earningsMap = [
      { pattern: 'Operational Pay', key: 'operationalPay' },
      { pattern: 'Flight Training Pay', key: 'fltTrainingPay' },
      { pattern: 'Sit Time', key: 'sitTime' },
      { pattern: 'Pay Above Guar', key: 'payAbvGuaranteeRsv' },
      { pattern: 'RA PREM', key: 'raPrem' },
      { pattern: 'Min Guarantee Adj', key: 'minGuaranteeAdj' },
      { pattern: 'Intl Override', key: 'intlOverride' },
      { pattern: 'Distance Learning', key: 'distanceLearning' },
      { pattern: 'Union Pd Union Leave', key: 'unionPdLeave' },
      { pattern: 'Prem Incentive Pay', key: 'premIncentivePay' },
      { pattern: 'Flight Vacation Pay', key: 'fltVacationPay' },
      { pattern: 'Sick Pay', key: 'sickPay' },
      { pattern: 'Prior Year VC Pay Out', key: 'priorYearVacPayout' },
      { pattern: 'Earnings Total', key: 'earningsTotal' }
    ];

    for (const item of earningsMap) {
      if (lineText.includes(item.pattern)) {
        let values = this.extractNumericValues(line);
        
        // Special case: Pay Above Guar (RSV) - values might be on previous line
        if (item.pattern === 'Pay Above Guar' && values.length < 3 && this.previousLine) {
          const prevValues = this.extractNumericValues(this.previousLine);
          if (prevValues.length >= 3) {
            values = prevValues;
          }
        }

        if (item.pattern === 'Earnings Total') {
          if (values.length >= 1) {
            earnings.earningsTotal = {
              current: values[0]
            };
          }
        }
        
        if (values.length >= 3) {
          earnings[item.key] = {
            rate: values[0],
            hours: values[1],
            current: values[2]
          };
        }
        break;
      }
    }
  }

  /**
   * Parse deductions lines
   */
  parseDeductionsLine(lineText, line, data) {
    // Company Contributions - Check FIRST (most specific)
    if (lineText.includes('401k Company Contrib')) {
      const values = this.extractNumericValues(line);
      if (values.length > 0) {
        data.companyContributions._401kCompanyContribution = values[0];
      }
      return; // Exit early to avoid matching other 401k patterns
    }
    
    if (lineText.includes('Group Term Life')) {
      const values = this.extractNumericValues(line);
      if (values.length > 0) {
        data.companyContributions.groupTermLife = values[0];
      }
      return;
    }

    // Pre-tax deductions
    const preTaxMap = [
      { pattern: 'Medical Coverage', key: 'medicalCoverage' },
      { pattern: 'Dental Coverage', key: 'dentalCoverage' },
      { pattern: 'Vision Coverage', key: 'visionCoverage' },
      { pattern: 'Accident Ins Pre-tax', key: 'accidentInsPreTax' },
      { pattern: '401k', key: '_401k' } // Generic 401k (not Roth, not Company)
    ];

    for (const item of preTaxMap) {
      const matches = typeof item.pattern === 'string' 
        ? lineText.includes(item.pattern) && !lineText.includes('Roth') && !lineText.includes('Company')
        : item.pattern.test(lineText);
      
      if (matches) {
        const values = this.extractNumericValues(line);
        if (values.length > 0) {
          data.deductions.preTax[item.key] = values[0];
        }
        return;
      }
    }

    // Taxes - Capture both Current (deduction) and Total (earnings)
    const taxMap = [
      { pattern: 'Withholding Tax', deductionKey: 'withholdingTax', earningsKey: 'withHoldingTaxEarnings' },
      { pattern: 'EE Social Security Tax', deductionKey: 'socialSecurityTax', earningsKey: 'socialSecurityTaxEarnings' },
      { pattern: 'EE Medicare Tax', deductionKey: 'medicareTax', earningsKey: 'medicareTaxEarnings' }
    ];

    for (const item of taxMap) {
      if (lineText.includes(item.pattern)) {
      const values = this.extractNumericValues(line);
      
      // First occurrence: Current deduction (single value)
      if (!this.taxesCurrentParsed[item.deductionKey] && values.length > 0) {
        data.deductions.taxes[item.deductionKey] = values[0];
        this.taxesCurrentParsed[item.deductionKey] = true;
        return;
      }
      
      // Second occurrence: Total earnings (single value)
      if (!this.taxesTotalParsed[item.deductionKey] && values.length > 0) {
        data.taxableEarnings[item.earningsKey] = values[0];
        this.taxesTotalParsed[item.deductionKey] = true;
        return;
      }
      }
    }

    // After-tax deductions - Check Roth 401k before other patterns
    const afterTaxMap = [
      { pattern: 'Roth 401k', key: 'roth401k' },
      { pattern: 'Employee Life', key: 'employeeLife' },
      { pattern: 'Dental Discount Plan', key: 'dentalDiscountPlan' },
      { pattern: 'PAC - APA', key: 'pacAPA' },
      { pattern: 'Union Dues - APA', key: 'unionDues' }
    ];

    for (const item of afterTaxMap) {
      if (lineText.includes(item.pattern)) {
        const values = this.extractNumericValues(line);
        if (values.length > 0) {
          data.deductions.afterTax[item.key] = values[0];
        }
        return; 
      }
    }
  }

  /**
   * Parse summary lines
   * Handles multi-line format:
   * Line 1: "Summary Gross Earnings  Pre-Tax Deduction Taxes After-Tax Deduction  Net Pay"
   * Line 2: "Current"
   * Line 3: "9,410.73 660.03 2,145.11 483.06 6,122.53"
   */
  parseSummaryLine(lineText, line, summary) {
    // Check for summary header line
    if (lineText.includes('Summary') && lineText.includes('Gross') && lineText.includes('Net Pay')) {
      this.summaryState.seenHeader = true;
      this.summaryState.seenCurrent = false;
      return;
    }
    
    // Check for "Current" line after header
    if (this.summaryState.seenHeader && lineText.includes('Current') && !lineText.includes('Employer')) {
      this.summaryState.seenCurrent = true;
      return;
    }
    
    // Check for values line after "Current"
    if (this.summaryState.seenCurrent) {
      const values = this.extractNumericValues(line);
      
      // Values format: Gross, Pre-Tax Deduct, Taxes, After Tax Deduct, Net Pay
      if (values.length >= 5) {
        summary.gross = values[0];
        summary.preTaxDeduct = values[1];
        summary.taxes = values[2];
        summary.afterTaxDeduct = values[3];
        summary.netPay = values[4];
        // Reset state after parsing
        this.summaryState.seenHeader = false;
        this.summaryState.seenCurrent = false;
      }
    }
  }

  /**
   * Extract all numeric values from a line (including currency)
   */
  extractNumericValues(line) {
    const values = [];
    
    for (const item of line.items) {
      // Match numbers with optional commas, decimals, dollar signs, and negative signs
      const text = item.text.replace(/\s/g, '');
      const match = text.match(/^(-)?\$?([\d,]+\.?\d*)$/);
      if (match) {
        // Preserve negative sign if present
        const negativeSign = match[1] || '';
        values.push(negativeSign + match[2]);
      }
    }
    
    return values;
  }

  /**
   * Convert hybrid parsed data to match your existing format
   */
  convertToLegacyFormat(hybridData) {
    return {
      // Info Data
      regularPayRoll: hybridData.header.regularPayRoll,
      payPeriod: hybridData.header.payPeriod,
      seniorityYear: hybridData.header.seniorityYear,
      group: hybridData.header.group,
      hourlyRate: hybridData.header.hourlyRate,
      effectivePeriod: hybridData.header.effectivePeriod,

      // Earnings Top Data
      crewAdvanceRate: hybridData.earnings.crewAdvance.rate,
      crewAdvanceHours: hybridData.earnings.crewAdvance.hours,
      crewAdvanceCurrent: hybridData.earnings.crewAdvance.current,
      pltExpDtaxRate: hybridData.earnings.pltExpDtax.rate,
      pltExpDtaxHours: hybridData.earnings.pltExpDtax.hours,
      pltExpDtaxCurrent: hybridData.earnings.pltExpDtax.current,
      pltExpADJtaxRate: hybridData.earnings.pltExpADJtax.rate,
      pltExpADJtaxHours: hybridData.earnings.pltExpADJtax.hours,
      pltExpADJtaxCurrent: hybridData.earnings.pltExpADJtax.current,
      pltExpItaxRate: hybridData.earnings.pltExpItax.rate,
      pltExpItaxHours: hybridData.earnings.pltExpItax.hours,
      pltExpItaxCurrent: hybridData.earnings.pltExpItax.current,
      AagProfitSharingRate: hybridData.earnings.AagProfitSharing.rate,
      AagProfitSharingHours: hybridData.earnings.AagProfitSharing.hours,
      AagProfitSharingCurrent: hybridData.earnings.AagProfitSharing.current,
      pltExpDnonTaxRate: hybridData.earnings.pltExpDnonTax.rate,
      pltExpDnonTaxHours: hybridData.earnings.pltExpDnonTax.hours,
      pltExpDnonTaxCurrent: hybridData.earnings.pltExpDnonTax.current,
      pltExpAdjnonTaxRate: hybridData.earnings.pltExpAdjnonTax.rate,
      pltExpAdjnonTaxHours: hybridData.earnings.pltExpAdjnonTax.hours,
      pltExpAdjnonTaxCurrent: hybridData.earnings.pltExpAdjnonTax.current,
      pltExpInonTaxRate: hybridData.earnings.pltExpInonTax.rate,
      pltExpInonTaxHours: hybridData.earnings.pltExpInonTax.hours,
      pltExpInonTaxCurrent: hybridData.earnings.pltExpInonTax.current,

      // Earnings Bottom Data
      operationalPayRate: hybridData.earnings.operationalPay.rate,
      operationalPayHours: hybridData.earnings.operationalPay.hours,
      operationalPayCurrent: hybridData.earnings.operationalPay.current,
      fltTrainingPayRate: hybridData.earnings.fltTrainingPay.rate,
      fltTrainingPayHours: hybridData.earnings.fltTrainingPay.hours,
      fltTrainingPayCurrent: hybridData.earnings.fltTrainingPay.current,
      sitTimeRate: hybridData.earnings.sitTime.rate,
      sitTimeHours: hybridData.earnings.sitTime.hours,
      sitTimeCurrent: hybridData.earnings.sitTime.current,
      payAbvGuaranteeRsvRate: hybridData.earnings.payAbvGuaranteeRsv.rate,
      payAbvGuaranteeRsvHours: hybridData.earnings.payAbvGuaranteeRsv.hours,
      payAbvGuaranteeRsvCurrent: hybridData.earnings.payAbvGuaranteeRsv.current,
      raPremRate: hybridData.earnings.raPrem.rate,
      raPremHours: hybridData.earnings.raPrem.hours,
      raPremCurrent: hybridData.earnings.raPrem.current,
      minGuaranteeAdjRate: hybridData.earnings.minGuaranteeAdj.rate,
      minGuaranteeAdjHours: hybridData.earnings.minGuaranteeAdj.hours,
      minGuaranteeAdjCurrent: hybridData.earnings.minGuaranteeAdj.current,
      intlOverrideRate: hybridData.earnings.intlOverride.rate,
      intlOverrideHours: hybridData.earnings.intlOverride.hours,
      intlOverrideCurrent: hybridData.earnings.intlOverride.current,
      distanceLearningRate: hybridData.earnings.distanceLearning.rate,
      distanceLearningHours: hybridData.earnings.distanceLearning.hours,
      distanceLearningCurrent: hybridData.earnings.distanceLearning.current,
      unionPdLeaveRate: hybridData.earnings.unionPdLeave.rate,
      unionPdLeaveHours: hybridData.earnings.unionPdLeave.hours,
      unionPdLeaveCurrent: hybridData.earnings.unionPdLeave.current,
      premIncentivePayRate: hybridData.earnings.premIncentivePay.rate,
      premIncentivePayHours: hybridData.earnings.premIncentivePay.hours,
      premIncentivePayCurrent: hybridData.earnings.premIncentivePay.current,
      fltVacationPayRate: hybridData.earnings.fltVacationPay.rate,
      fltVacationPayHours: hybridData.earnings.fltVacationPay.hours,
      fltVacationPayCurrent: hybridData.earnings.fltVacationPay.current,
      sickPayRate: hybridData.earnings.sickPay.rate,
      sickPayHours: hybridData.earnings.sickPay.hours,
      sickPayCurrent: hybridData.earnings.sickPay.current,
      priorYearVacPayoutRate: hybridData.earnings.priorYearVacPayout.rate,
      priorYearVacPayoutHours: hybridData.earnings.priorYearVacPayout.hours,
      priorYearVacPayoutCurrent: hybridData.earnings.priorYearVacPayout.current,
      earningsTotalCurrent: hybridData.earnings.earningsTotal.current,

      // Pre-tax deductions
      medicalCoverage: hybridData.deductions.preTax.medicalCoverage,
      dentalCoverage: hybridData.deductions.preTax.dentalCoverage,
      visionCoverage: hybridData.deductions.preTax.visionCoverage,
      accidentInsPreTax: hybridData.deductions.preTax.accidentInsPreTax,
      _401k: hybridData.deductions.preTax._401k,

      // Taxes
      withholdingTax: hybridData.deductions.taxes.withholdingTax,
      socialSecurityTax: hybridData.deductions.taxes.socialSecurityTax,
      medicareTax: hybridData.deductions.taxes.medicareTax,

      // After-tax deductions
      employeeLife: hybridData.deductions.afterTax.employeeLife,
      dentalDiscountPlan: hybridData.deductions.afterTax.dentalDiscountPlan,
      roth401k: hybridData.deductions.afterTax.roth401k,
      pacAPA: hybridData.deductions.afterTax.pacAPA,
      unionDues: hybridData.deductions.afterTax.unionDues,

      // Company contributions
      _401kCompanyContribution: hybridData.companyContributions._401kCompanyContribution,
      groupTermLife: hybridData.companyContributions.groupTermLife,

      // Taxable earnings
      withHoldingTaxEarnings: hybridData.taxableEarnings.withHoldingTaxEarnings,
      socialSecurityTaxEarnings: hybridData.taxableEarnings.socialSecurityTaxEarnings,
      medicareTaxEarnings: hybridData.taxableEarnings.medicareTaxEarnings,

      // Summary
      gross: hybridData.summary.gross,
      preTaxDeduct: hybridData.summary.preTaxDeduct,
      taxes: hybridData.summary.taxes,
      afterTaxDeduct: hybridData.summary.afterTaxDeduct,
      netPay: hybridData.summary.netPay
    };
  }

  
}
