/**
 * 08_hoursDemographic.gs
 * Hours Demographic Report - Gender and hours analysis for current employees
 */

/**
 * Generates demographic breakdown by gender and contract hours
 * @param {Array} data - Employee data array
 * @param {Object} options - Options object (not used but kept for consistency)
 * @returns {Array} 2D array with report data
 */
function generateHoursDemographicReport(data, options = {}) {
  try {
    console.log('📊 Generating Hours Demographic Report...');
    
    // Filter for current employees only
    const currentEmployees = data.filter(emp => {
      const status = (emp.Status || '').toLowerCase().trim();
      return status === 'current';
    });
    
    console.log(`Found ${currentEmployees.length} current employees to analyze`);
    
    // Initialize counters
    const demographics = {
      malePartTime: 0,    // Male < 30 hours
      maleFullTime: 0,    // Male ≥ 30 hours
      femalePartTime: 0,  // Female < 30 hours
      femaleFullTime: 0   // Female ≥ 30 hours
    };
    
    let invalidGender = 0;
    let invalidHours = 0;
    let processed = 0;
    
    // Analyze each employee
    currentEmployees.forEach(emp => {
      const gender = (emp.Gender || '').toLowerCase().trim();
      const contractHours = parseFloat(emp.ContractHours || '0');
      
      // Validate data
      if (!gender || (gender !== 'male' && gender !== 'female')) {
        invalidGender++;
        return;
      }
      
      if (isNaN(contractHours) || contractHours < 0) {
        invalidHours++;
        return;
      }
      
      // Categorize
      const isPartTime = contractHours < 30;
      const isMale = gender === 'male';
      
      if (isMale && isPartTime) {
        demographics.malePartTime++;
      } else if (isMale && !isPartTime) {
        demographics.maleFullTime++;
      } else if (!isMale && isPartTime) {
        demographics.femalePartTime++;
      } else {
        demographics.femaleFullTime++;
      }
      
      processed++;
    });
    
    console.log(`Processed: ${processed}, Invalid gender: ${invalidGender}, Invalid hours: ${invalidHours}`);
    
    // Calculate totals
    const totalMale = demographics.malePartTime + demographics.maleFullTime;
    const totalFemale = demographics.femalePartTime + demographics.femaleFullTime;
    const totalPartTime = demographics.malePartTime + demographics.femalePartTime;
    const totalFullTime = demographics.maleFullTime + demographics.femaleFullTime;
    const grandTotal = processed;
    
    // Build report output
    const reportDate = new Date().toLocaleDateString('en-GB');
    const titleRow = ['MONTHLY HOURS DEMOGRAPHIC REPORT', `Report Date: ${reportDate}`, '', '', ''];
    const summaryText = `${grandTotal} current employees analyzed (${invalidGender + invalidHours} excluded)`;
    const summaryRow = ['SUMMARY', summaryText, '', '', ''];
    const blankRow = ['', '', '', '', ''];
    
    // Main data table
    const headerRow = ['Category', 'Part Time (<30 hrs)', 'Full Time (≥30 hrs)', 'Total', ''];
    const maleRow = ['Male', demographics.malePartTime, demographics.maleFullTime, totalMale, ''];
    const femaleRow = ['Female', demographics.femalePartTime, demographics.femaleFullTime, totalFemale, ''];
    const totalRow = ['TOTAL', totalPartTime, totalFullTime, grandTotal, ''];
    
    const blank2Row = ['', '', '', '', ''];
    
    // Detailed breakdown
    const detailHeaderRow = ['Detailed Breakdown', 'Count', 'Percentage', '', ''];
    const detail1 = ['Male Part Time (<30 hrs)', demographics.malePartTime, `${((demographics.malePartTime / grandTotal) * 100).toFixed(1)}%`, '', ''];
    const detail2 = ['Male Full Time (≥30 hrs)', demographics.maleFullTime, `${((demographics.maleFullTime / grandTotal) * 100).toFixed(1)}%`, '', ''];
    const detail3 = ['Female Part Time (<30 hrs)', demographics.femalePartTime, `${((demographics.femalePartTime / grandTotal) * 100).toFixed(1)}%`, '', ''];
    const detail4 = ['Female Full Time (≥30 hrs)', demographics.femaleFullTime, `${((demographics.femaleFullTime / grandTotal) * 100).toFixed(1)}%`, '', ''];
    
    // Data quality notes
    const blank3Row = ['', '', '', '', ''];
    const qualityHeaderRow = ['Data Quality Notes', '', '', '', ''];
    const qualityNote1 = [`Employees with invalid gender: ${invalidGender}`, '', '', '', ''];
    const qualityNote2 = [`Employees with invalid hours: ${invalidHours}`, '', '', '', ''];
    const qualityNote3 = [`Total processed: ${processed} of ${currentEmployees.length}`, '', '', '', ''];
    
    console.log(`✅ Hours Demographic report generated successfully`);
    console.log(`Matrix: Male PT:${demographics.malePartTime}, Male FT:${demographics.maleFullTime}, Female PT:${demographics.femalePartTime}, Female FT:${demographics.femaleFullTime}`);
    
    return [
      titleRow,
      summaryRow,
      blankRow,
      headerRow,
      maleRow,
      femaleRow,
      totalRow,
      blank2Row,
      detailHeaderRow,
      detail1,
      detail2,
      detail3,
      detail4,
      blank3Row,
      qualityHeaderRow,
      qualityNote1,
      qualityNote2,
      qualityNote3
    ];
    
  } catch (error) {
    console.error(`Error generating hours demographic report: ${error.message}`);
    throw new Error('Failed to generate hours demographic report: ' + error.message);
  }
}

/**
 * Test function for the hours demographic report
 */
function testHoursDemographicReport() {
  console.log('🧪 Testing Hours Demographic Report...');
  
  try {
    const employees = getAllEmployees();
    console.log(`Loaded ${employees.length} total employees`);
    
    const output = generateHoursDemographicReport(employees);
    console.log(`Generated report with ${output.length} rows`);
    
    // Show key results
    console.log('\nKey Results:');
    output.slice(3, 7).forEach((row, index) => {
      if (index === 0) {
        console.log(`Headers: ${row.join(' | ')}`);
      } else {
        console.log(`${row[0]}: PT=${row[1]}, FT=${row[2]}, Total=${row[3]}`);
      }
    });
    
    // Create test file
    const testFile = writeReportToStandaloneWorkbook(
      'hoursDemographic', 
      output, 
      Session.getEffectiveUser().getEmail()
    );
    
    console.log(`✅ Test report created: ${testFile.getUrl()}`);
    
    return {
      success: true,
      totalEmployees: employees.length,
      reportRows: output.length,
      url: testFile.getUrl()
    };
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Quick analysis of gender and hours data quality
 */
function analyzeHoursDemographicDataQuality() {
  console.log('🔍 Analyzing data quality for Hours Demographic Report...');
  
  try {
    const employees = getAllEmployees();
    const currentEmployees = employees.filter(emp => (emp.Status || '').toLowerCase().trim() === 'current');
    
    console.log(`Total employees: ${employees.length}`);
    console.log(`Current employees: ${currentEmployees.length}`);
    
    // Analyze gender data
    const genderStats = {};
    const hourStats = { valid: 0, invalid: 0, zero: 0 };
    
    currentEmployees.forEach(emp => {
      const gender = (emp.Gender || '').toLowerCase().trim();
      const hours = parseFloat(emp.ContractHours || '0');
      
      // Gender analysis
      if (!gender) {
        genderStats['blank'] = (genderStats['blank'] || 0) + 1;
      } else {
        genderStats[gender] = (genderStats[gender] || 0) + 1;
      }
      
      // Hours analysis
      if (isNaN(hours) || hours < 0) {
        hourStats.invalid++;
      } else if (hours === 0) {
        hourStats.zero++;
      } else {
        hourStats.valid++;
      }
    });
    
    console.log('\n📊 Gender Distribution:');
    Object.entries(genderStats).forEach(([gender, count]) => {
      const pct = ((count / currentEmployees.length) * 100).toFixed(1);
      console.log(`   ${gender}: ${count} (${pct}%)`);
    });
    
    console.log('\n⏱️ Contract Hours Distribution:');
    console.log(`   Valid hours: ${hourStats.valid}`);
    console.log(`   Zero hours: ${hourStats.zero}`);
    console.log(`   Invalid hours: ${hourStats.invalid}`);
    
    const dataQuality = ((hourStats.valid / currentEmployees.length) * 100).toFixed(1);
    console.log(`\n📈 Data Quality Score: ${dataQuality}% usable records`);
    
    return {
      totalCurrent: currentEmployees.length,
      genderStats: genderStats,
      hourStats: hourStats,
      dataQuality: dataQuality
    };
    
  } catch (error) {
    console.error('❌ Analysis failed:', error.message);
    return { error: error.message };
  }
}
