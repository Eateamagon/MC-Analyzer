/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                          EXPORT & EMAIL FUNCTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Generates a Google Sheet with analysis results for a specific teacher
 */
function createTeacherSheet(assessmentId, teacherEmail) {
  const user = getCurrentUser();
  if (!user || user.role === CONFIG.roles.TEACHER) {
    throw new Error('Admin access required');
  }
  
  const results = getAnalysisResults(assessmentId);
  const assessmentData = getAssessmentData(assessmentId);
  
  // Create new spreadsheet
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const teacherName = teacherEmail.split('@')[0];
  const newSS = SpreadsheetApp.create(`${teacherName} - ${assessmentData.assessment.name} - ${timestamp}`);
  
  // Create analysis sheets
  createAnalysisSheets(newSS, results, teacherName);
  
  // Remove default sheet
  const defaultSheet = newSS.getSheetByName('Sheet1');
  if (defaultSheet && newSS.getSheets().length > 1) {
    newSS.deleteSheet(defaultSheet);
  }
  
  // Move to appropriate folder
  const file = DriveApp.getFileById(newSS.getId());
  const folder = getOrCreateFolder(teacherName, assessmentData.assessment.subject || 'General');
  file.moveTo(folder);
  
  return {
    url: newSS.getUrl(),
    name: newSS.getName()
  };
}

/**
 * Creates analysis sheets in a spreadsheet
 */
function createAnalysisSheets(ss, results, teacherFilter) {
  const colors = CONFIG.colors;
  
  // Teacher Summary Sheet
  if (results.teacherSummaries) {
    const summaries = teacherFilter 
      ? results.teacherSummaries.filter(ts => ts.teacher.toLowerCase().includes(teacherFilter.toLowerCase()))
      : results.teacherSummaries;
    
    summaries.forEach(ts => {
      const sh = ss.insertSheet(`${ts.teacher.substring(0, 15)}-Summary`);
      
      // Header
      sh.getRange('A1:G1').merge()
        .setValue(`Analysis for ${ts.teacher}`)
        .setBackground('#00a14b')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setFontSize(14);
      
      sh.getRange('A2:G2').merge()
        .setValue(`Students: ${ts.studentCount}   Average: ${ts.average}%`)
        .setHorizontalAlignment('center')
        .setFontStyle('italic');
      
      let row = 4;
      
      // Growth Section
      if (ts.growthSOLs.length > 0) {
        sh.getRange(row, 1, 1, 7).merge()
          .setValue(`AREAS OF GROWTH (<${results.settings.growth || 50}%)`)
          .setBackground('#f4c7c3')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        row++;
        
        sh.getRange(row, 1, 1, 7)
          .setValues([['SOL', 'Correct', 'Total', '% Mastery', 'Red', 'Yellow', 'Green']])
          .setFontWeight('bold')
          .setBackground('#efefef');
        row++;
        
        ts.growthSOLs.forEach(sol => {
          sh.getRange(row, 1, 1, 7).setValues([[
            sol.sol, sol.correct, sol.total, `${sol.percentage}%`,
            sol.redCount, sol.yellowCount, sol.greenCount
          ]]);
          row++;
        });
        row++;
      }
      
      // Monitor Section
      if (ts.monitorSOLs.length > 0) {
        sh.getRange(row, 1, 1, 7).merge()
          .setValue(`AREAS TO MONITOR (${results.settings.growth || 50}% - ${results.settings.strength || 75}%)`)
          .setBackground('#fce8b2')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        row++;
        
        sh.getRange(row, 1, 1, 7)
          .setValues([['SOL', 'Correct', 'Total', '% Mastery', 'Red', 'Yellow', 'Green']])
          .setFontWeight('bold')
          .setBackground('#efefef');
        row++;
        
        ts.monitorSOLs.forEach(sol => {
          sh.getRange(row, 1, 1, 7).setValues([[
            sol.sol, sol.correct, sol.total, `${sol.percentage}%`,
            sol.redCount, sol.yellowCount, sol.greenCount
          ]]);
          row++;
        });
        row++;
      }
      
      // Strength Section
      if (ts.strengthSOLs.length > 0) {
        sh.getRange(row, 1, 1, 7).merge()
          .setValue(`AREAS OF STRENGTH (>=${results.settings.strength || 75}%)`)
          .setBackground('#b7e1cd')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        row++;
        
        sh.getRange(row, 1, 1, 7)
          .setValues([['SOL', 'Correct', 'Total', '% Mastery', 'Red', 'Yellow', 'Green']])
          .setFontWeight('bold')
          .setBackground('#efefef');
        row++;
        
        ts.strengthSOLs.forEach(sol => {
          sh.getRange(row, 1, 1, 7).setValues([[
            sol.sol, sol.correct, sol.total, `${sol.percentage}%`,
            sol.redCount, sol.yellowCount, sol.greenCount
          ]]);
          row++;
        });
      }
      
      sh.autoResizeColumns(1, 7);
    });
  }
  
  // Item Analysis Sheet
  if (results.itemAnalysis) {
    Object.keys(results.itemAnalysis).forEach(teacher => {
      if (teacherFilter && !teacher.toLowerCase().includes(teacherFilter.toLowerCase())) return;
      
      const analysis = results.itemAnalysis[teacher];
      const sh = ss.insertSheet(`${teacher.substring(0, 12)}-ItemAnalysis`);
      
      sh.getRange('A1:H1').merge()
        .setValue(`Item Analysis - ${teacher}`)
        .setBackground('#00a14b')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setFontSize(14);
      
      sh.getRange(3, 1, 1, 8)
        .setValues([['Question', 'SOL', 'Correct Answer', '% Correct', 'Top Wrong', 'Count', 'All Distractors', 'Alert']])
        .setFontWeight('bold')
        .setBackground('#e8f0fe');
      
      analysis.forEach((item, idx) => {
        const distractorStr = item.allDistractors.map(d => `${d.answer}:${d.count}`).join(' | ');
        sh.getRange(4 + idx, 1, 1, 8).setValues([[
          item.question,
          item.sol,
          item.correctAnswer,
          item.pctCorrect,
          item.topWrongAnswer,
          item.topWrongCount || '-',
          distractorStr || 'None',
          item.reviewFlag ? '⚠️ Review' : ''
        ]]);
        
        if (item.reviewFlag) {
          sh.getRange(4 + idx, 8).setBackground('#fce8e6').setFontColor('#c5221f');
        }
      });
      
      sh.autoResizeColumns(1, 8);
    });
  }
  
  // Small Groups Sheet
  if (results.smallGroups && results.smallGroups.length > 0) {
    const sh = ss.insertSheet('Small Groups');
    
    sh.getRange('A1:F1').merge()
      .setValue('Small Group Interventions')
      .setBackground('#00a14b')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setFontSize(14);
    
    sh.getRange(3, 1, 1, 6)
      .setValues([['Teacher', 'Period', 'Group', 'Students', 'Shared Weak SOLs', 'Count']])
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
    
    results.smallGroups.forEach((group, idx) => {
      if (teacherFilter && !group.teacher.toLowerCase().includes(teacherFilter.toLowerCase())) return;
      
      sh.getRange(4 + idx, 1, 1, 6).setValues([[
        group.teacher,
        group.period,
        `Group ${group.groupNumber}`,
        group.students.join(', '),
        group.sharedWeakSOLs.join(', '),
        group.studentCount
      ]]);
      
      if (group.isUnknownPeriod) {
        sh.getRange(4 + idx, 2).setBackground('#f4c7c3');
      }
    });
    
    sh.autoResizeColumns(1, 6);
  }
}

/**
 * Gets or creates folder structure for organizing files
 */
function getOrCreateFolder(teacherName, subject) {
  teacherName = teacherName.trim().replace(/[\/\\:*?"<>|]/g, '-');
  subject = subject.trim().replace(/[\/\\:*?"<>|]/g, '-');
  
  // Get or create root folder
  let rootFolders = DriveApp.getFoldersByName('Mastery Connect Analysis');
  let rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder('Mastery Connect Analysis');
  
  // Get or create subject folder
  let subjectFolders = rootFolder.getFoldersByName(subject);
  let subjectFolder = subjectFolders.hasNext() ? subjectFolders.next() : rootFolder.createFolder(subject);
  
  // Get or create teacher folder
  let teacherFolders = subjectFolder.getFoldersByName(teacherName);
  return teacherFolders.hasNext() ? teacherFolders.next() : subjectFolder.createFolder(teacherName);
}

/**
 * Generates Excel download data
 */
function generateExcelDownload(assessmentId, options) {
  const results = getAnalysisResults(assessmentId);
  const assessmentData = getAssessmentData(assessmentId);
  
  // Create temporary spreadsheet
  const tempSS = SpreadsheetApp.create('temp_export_' + Date.now());
  
  try {
    createAnalysisSheets(tempSS, results, options.teacher || null);
    
    // Remove default sheet
    const defaultSheet = tempSS.getSheetByName('Sheet1');
    if (defaultSheet && tempSS.getSheets().length > 1) {
      tempSS.deleteSheet(defaultSheet);
    }
    
    // Get file as Excel
    const url = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=xlsx';
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    const blob = response.getBlob();
    blob.setName(`${assessmentData.assessment.name}_Analysis.xlsx`);
    
    // Convert to base64 for download
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return {
      success: true,
      data: base64,
      filename: blob.getName(),
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };
  } finally {
    // Clean up temp file
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  }
}

/**
 * Sends analysis results via email
 */
function sendAnalysisEmail(email, results) {
  const subject = `Mastery Connect Analysis Complete: ${results.assessmentName}`;
  
  // Build email body
  let body = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: #00a14b; color: white; padding: 20px; text-align: center;">
        <h1 style="margin: 0;">Mastery Connect Analyzer</h1>
      </div>
      <div style="padding: 20px; background: #f8f9fa;">
        <h2>Analysis Complete!</h2>
        <p>Your analysis for <strong>${results.assessmentName}</strong> has been completed.</p>
        
        <h3>Summary</h3>
        <table style="width: 100%; border-collapse: collapse;">
          <tr style="background: #e8f0fe;">
            <td style="padding: 10px; border: 1px solid #ddd;"><strong>Teachers Analyzed</strong></td>
            <td style="padding: 10px; border: 1px solid #ddd;">${results.teacherSummaries.length}</td>
          </tr>
  `;
  
  if (results.districtSummary) {
    body += `
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd;"><strong>Total Students</strong></td>
            <td style="padding: 10px; border: 1px solid #ddd;">${results.districtSummary.totalStudents}</td>
          </tr>
          <tr style="background: #e8f0fe;">
            <td style="padding: 10px; border: 1px solid #ddd;"><strong>Overall Average</strong></td>
            <td style="padding: 10px; border: 1px solid #ddd;">${results.districtSummary.overallAverage}%</td>
          </tr>
    `;
  }
  
  body += `
        </table>
        
        <p style="margin-top: 20px;">
          <a href="${ScriptApp.getService().getUrl()}" 
             style="background: #00a14b; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">
            View Full Results
          </a>
        </p>
      </div>
      <div style="padding: 10px; text-align: center; color: #666; font-size: 12px;">
        Mastery Connect Analyzer - Waynesboro Public Schools
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: body
  });
}

/**
 * Creates separate files for all teachers (admin function)
 */
function createAllTeacherFiles(assessmentId) {
  const user = getCurrentUser();
  if (!user || user.role === CONFIG.roles.TEACHER) {
    throw new Error('Admin access required');
  }
  
  const results = getAnalysisResults(assessmentId);
  const createdFiles = [];
  
  results.teacherSummaries.forEach(ts => {
    try {
      const file = createTeacherSheet(assessmentId, ts.teacher);
      createdFiles.push({
        teacher: ts.teacher,
        url: file.url,
        name: file.name,
        success: true
      });
    } catch (e) {
      createdFiles.push({
        teacher: ts.teacher,
        error: e.message,
        success: false
      });
    }
  });
  
  return createdFiles;
}

/**
 * Exports comparison results as a PDF.
 * Creates a temporary spreadsheet with summary, teacher, and student sheets.
 */
function exportComparisonPDF(assessmentIdA, assessmentIdB) {
  var comparison = compareAssessments(assessmentIdA, assessmentIdB);
  var tempSS = SpreadsheetApp.create('temp_compare_pdf_' + Date.now());

  try {
    // Summary sheet
    var sh = tempSS.getSheets()[0];
    sh.setName('Summary');
    sh.getRange('A1:F1').merge()
      .setValue('Assessment Comparison')
      .setBackground('#7C3AED').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(14);
    sh.getRange('A3').setValue('Assessment A:');
    sh.getRange('B3').setValue(comparison.assessmentA.name);
    sh.getRange('C3').setValue('Average: ' + comparison.assessmentA.average + '%');
    sh.getRange('D3').setValue(comparison.assessmentA.studentCount + ' students');
    sh.getRange('A4').setValue('Assessment B:');
    sh.getRange('B4').setValue(comparison.assessmentB.name);
    sh.getRange('C4').setValue('Average: ' + comparison.assessmentB.average + '%');
    sh.getRange('D4').setValue(comparison.assessmentB.studentCount + ' students');
    sh.getRange('A5').setValue('Overall Change:');
    sh.getRange('B5').setValue((comparison.overallDelta > 0 ? '+' : '') + comparison.overallDelta + '%')
      .setFontWeight('bold');

    // Teacher sheet
    var tSh = tempSS.insertSheet('By Teacher');
    tSh.getRange(1, 1, 1, 6)
      .setValues([['Teacher', 'Students A', 'Students B', 'Avg A', 'Avg B', 'Change']])
      .setFontWeight('bold').setBackground('#e8f0fe');
    comparison.teacherComparisons.forEach(function(tc, idx) {
      tSh.getRange(2 + idx, 1, 1, 6).setValues([[
        tc.teacher, tc.studentsA, tc.studentsB,
        tc.avgA !== null ? tc.avgA + '%' : '-',
        tc.avgB !== null ? tc.avgB + '%' : '-',
        tc.avgDelta !== null ? (tc.avgDelta > 0 ? '+' : '') + tc.avgDelta + '%' : '-'
      ]]);
    });
    tSh.autoResizeColumns(1, 6);

    // Student sheet
    if (comparison.studentComparisons.length > 0) {
      var sSh = tempSS.insertSheet('By Student');
      sSh.getRange(1, 1, 1, 5)
        .setValues([['Student', 'Teacher', comparison.assessmentA.name, comparison.assessmentB.name, 'Change']])
        .setFontWeight('bold').setBackground('#e8f0fe');
      comparison.studentComparisons.forEach(function(sc, idx) {
        sSh.getRange(2 + idx, 1, 1, 5).setValues([[
          sc.name, sc.teacher, sc.pctA + '%', sc.pctB + '%',
          (sc.delta > 0 ? '+' : '') + sc.delta + '%'
        ]]);
      });
      sSh.autoResizeColumns(1, 5);
    }

    sh.autoResizeColumns(1, 6);

    var url = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=pdf&portrait=false&size=letter&fitw=true';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    var blob = response.getBlob();
    var filename = 'Comparison_' + comparison.assessmentA.name + '_vs_' + comparison.assessmentB.name + '.pdf';
    blob.setName(filename.replace(/[^a-zA-Z0-9._\-]/g, '_'));
    var base64 = Utilities.base64Encode(blob.getBytes());

    return { success: true, data: base64, filename: blob.getName(), mimeType: 'application/pdf' };
  } finally {
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  }
}

/**
 * Generates PDF report
 */
function generatePDFReport(assessmentId, options) {
  const results = getAnalysisResults(assessmentId);
  const assessmentData = getAssessmentData(assessmentId);
  
  // Create temporary spreadsheet
  const tempSS = SpreadsheetApp.create('temp_pdf_' + Date.now());
  
  try {
    createAnalysisSheets(tempSS, results, options.teacher || null);
    
    // Remove default sheet
    const defaultSheet = tempSS.getSheetByName('Sheet1');
    if (defaultSheet && tempSS.getSheets().length > 1) {
      tempSS.deleteSheet(defaultSheet);
    }
    
    // Get file as PDF
    const url = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=pdf&portrait=false&size=letter&fitw=true';
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    const blob = response.getBlob();
    blob.setName(`${assessmentData.assessment.name}_Analysis.pdf`);
    
    // Convert to base64 for download
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return {
      success: true,
      data: base64,
      filename: blob.getName(),
      mimeType: 'application/pdf'
    };
  } finally {
    // Clean up temp file
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  }
}
