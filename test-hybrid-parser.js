/**
 * QUICK START: Testing the Hybrid Parser
 * 
 * Run this script to test the hybrid parser with your PDF files
 * Usage: node test-hybrid-parser.js /path/to/your/paystub.pdf
 */

import fs from 'fs';
import { HybridPDFParser } from './hybrid-parser.js';

async function testParser(pdfPath) {
  console.log('🚀 Starting Hybrid Parser Test...\n');
  console.log(`📄 PDF File: ${pdfPath}\n`);
  
  try {
    // Read the PDF file
    const pdfBuffer = fs.readFileSync(pdfPath);
    console.log('✅ PDF file loaded successfully\n');
    
    // Create parser and parse
    const parser = new HybridPDFParser();
    console.log('⚙️  Parsing PDF with Hybrid Parser...\n');
    
    const parsedData = await parser.parse(pdfBuffer);
    console.log('✅ Parsing completed!\n');
    
    // Display results
    console.log('═══════════════════════════════════════════════════════════');
    console.log('                    PARSED DATA RESULTS');
    console.log('═══════════════════════════════════════════════════════════\n');
    
    // Header Info
    console.log('📋 HEADER INFORMATION:');
    console.log('   Pay Period:', parsedData.header.payPeriod || '(not found)');
    console.log('   Seniority Year:', parsedData.header.seniorityYear);
    console.log('   Group:', parsedData.header.group || '(not found)');
    console.log('   Hourly Rate: $' + parsedData.header.hourlyRate);
    console.log('');
    
    // Earnings with values
    console.log('💰 EARNINGS (non-zero values):');
    let earningsCount = 0;
    for (const [key, value] of Object.entries(parsedData.earnings)) {
      if (value.current && value.current !== '0') {
        console.log(`   ${formatFieldName(key)}:`);
        console.log(`      Rate: ${value.rate}, Hours: ${value.hours}, Current: ${value.current}`);
        earningsCount++;
      }
    }
    if (earningsCount === 0) {
      console.log('   (No earnings data found)');
    }
    console.log('');
    
    // Deductions
    console.log('💳 PRE-TAX DEDUCTIONS:');
    let preTaxCount = 0;
    for (const [key, value] of Object.entries(parsedData.deductions.preTax)) {
      if (value && value !== '0') {
        console.log(`   ${formatFieldName(key)}: ${value}`);
        preTaxCount++;
      }
    }
    if (preTaxCount === 0) {
      console.log('   (No pre-tax deductions found)');
    }
    console.log('');
    
    console.log('📊 TAXES:');
    let taxCount = 0;
    for (const [key, value] of Object.entries(parsedData.deductions.taxes)) {
      if (value && value !== '0') {
        console.log(`   ${formatFieldName(key)}: ${value}`);
        taxCount++;
      }
    }
    if (taxCount === 0) {
      console.log('   (No tax data found)');
    }
    console.log('');
    
    console.log('💵 AFTER-TAX DEDUCTIONS:');
    let afterTaxCount = 0;
    for (const [key, value] of Object.entries(parsedData.deductions.afterTax)) {
      if (value && value !== '0') {
        console.log(`   ${formatFieldName(key)}: ${value}`);
        afterTaxCount++;
      }
    }
    if (afterTaxCount === 0) {
      console.log('   (No after-tax deductions found)');
    }
    console.log('');
    
    // Summary
    console.log('📈 SUMMARY:');
    console.log(`   Gross: ${parsedData.summary.gross}`);
    console.log(`   Pre-Tax Deduct: ${parsedData.summary.preTaxDeduct}`);
    console.log(`   Taxes: ${parsedData.summary.taxes}`);
    console.log(`   After Tax Deduct: ${parsedData.summary.afterTaxDeduct}`);
    console.log(`   Net Pay: ${parsedData.summary.netPay}`);
    console.log('');
    
    console.log('═══════════════════════════════════════════════════════════\n');
    
    // Test legacy format conversion
    console.log('🔄 Testing Legacy Format Conversion...\n');
    const legacyData = parser.convertToLegacyFormat(parsedData);
    console.log('✅ Legacy format conversion successful!');
    console.log('   Sample fields:');
    console.log(`   - operationalPayRate: ${legacyData.operationalPayRate}`);
    console.log(`   - operationalPayHours: ${legacyData.operationalPayHours}`);
    console.log(`   - operationalPayCurrent: ${legacyData.operationalPayCurrent}`);
    console.log(`   - gross: ${legacyData.gross}`);
    console.log(`   - netPay: ${legacyData.netPay}`);
    console.log('');
    
    // Export to JSON for inspection
    const outputPath = pdfPath.replace('.pdf', '_parsed.json');
    fs.writeFileSync(outputPath, JSON.stringify(parsedData, null, 2));
    console.log(`📁 Full parsed data exported to: ${outputPath}\n`);
    
    console.log('✨ Test completed successfully!\n');
    
    return parsedData;
    
  } catch (error) {
    console.error('❌ Error:', error.message);
    console.error('\nStack trace:', error.stack);
    process.exit(1);
  }
}

// Helper function to format field names nicely
function formatFieldName(fieldName) {
  return fieldName
    .replace(/([A-Z])/g, ' $1')
    .replace(/^./, str => str.toUpperCase())
    .trim();
}

// Main execution
const pdfPath = process.argv[2];

if (!pdfPath) {
  console.log('Usage: node test-hybrid-parser.js <path-to-pdf-file>');
  console.log('Example: node test-hybrid-parser.js ~/Downloads/paystub.pdf');
  process.exit(1);
}

if (!fs.existsSync(pdfPath)) {
  console.error(`Error: File not found: ${pdfPath}`);
  process.exit(1);
}

testParser(pdfPath).catch(error => {
  console.error('Unhandled error:', error);
  process.exit(1);
});
