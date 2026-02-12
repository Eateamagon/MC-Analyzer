/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                          ADMIN FUNCTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Gets all users (admin function)
 */
function getAllUsers() {
  const user = getCurrentUser();
  if (!user || user.role === CONFIG.roles.TEACHER) {
    throw new Error('Admin access required');
  }
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.users);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const userObj = {};
    headers.forEach((h, idx) => {
      userObj[h] = data[i][idx];
    });
    
    // Filter by school for school admins
    if (user.role === CONFIG.roles.SCHOOL_ADMIN) {
      if (userObj.school === user.school) {
        users.push(userObj);
      }
    } else {
      users.push(userObj);
    }
  }
  
  return users;
}

/**
 * Gets admin list from settings
 */
function getAdminList() {
  const user = getCurrentUser();
  if (!user || user.role !== CONFIG.roles.DIVISION_ADMIN) {
    throw new Error('Division admin access required');
  }
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.settings);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { division: [], school: [] };
  }
  
  const data = sheet.getDataRange().getValues();
  const division = [];
  const school = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'division_admin') {
      division.push({ email: data[i][1], school: data[i][2] || '' });
    } else if (data[i][0] === 'school_admin') {
      school.push({ email: data[i][1], school: data[i][2] || '' });
    }
  }
  
  return { division, school };
}

/**
 * Adds an admin email to the settings
 */
function addAdminEmail(type, email, school) {
  const user = getCurrentUser();
  if (!user || user.role !== CONFIG.roles.DIVISION_ADMIN) {
    throw new Error('Division admin access required');
  }
  
  if (!email || !email.includes('@')) {
    throw new Error('Invalid email address');
  }
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.settings);
  
  // Check if already exists
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === type && data[i][1].toLowerCase() === email.toLowerCase()) {
      throw new Error('This email is already an admin');
    }
  }
  
  // Add to settings
  sheet.appendRow([type, email.toLowerCase(), school || '']);
  
  // Update user role if they exist
  const userSheet = ss.getSheetByName(CONFIG.sheets.users);
  const userData = userSheet.getDataRange().getValues();
  const userHeaders = userData[0];
  const emailCol = userHeaders.indexOf('email');
  const roleCol = userHeaders.indexOf('role');
  const schoolCol = userHeaders.indexOf('school');
  
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][emailCol] && userData[i][emailCol].toLowerCase() === email.toLowerCase()) {
      userSheet.getRange(i + 1, roleCol + 1).setValue(type);
      if (school && type === 'school_admin') {
        userSheet.getRange(i + 1, schoolCol + 1).setValue(school);
      }
      break;
    }
  }
  
  return { success: true };
}

/**
 * Removes an admin email from settings
 */
function removeAdminEmail(type, email) {
  const user = getCurrentUser();
  if (!user || user.role !== CONFIG.roles.DIVISION_ADMIN) {
    throw new Error('Division admin access required');
  }
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.settings);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === type && data[i][1].toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      
      // Update user role to teacher
      const userSheet = ss.getSheetByName(CONFIG.sheets.users);
      const userData = userSheet.getDataRange().getValues();
      const userHeaders = userData[0];
      const emailCol = userHeaders.indexOf('email');
      const roleCol = userHeaders.indexOf('role');
      
      for (let j = 1; j < userData.length; j++) {
        if (userData[j][emailCol] && userData[j][emailCol].toLowerCase() === email.toLowerCase()) {
          userSheet.getRange(j + 1, roleCol + 1).setValue(CONFIG.roles.TEACHER);
          break;
        }
      }
      
      return { success: true };
    }
  }
  
  throw new Error('Admin not found');
}

/**
 * Deletes an assessment and all related data
 */
function deleteAssessment(assessmentId) {
  const user = getCurrentUser();
  if (!user || user.isNewUser) {
    throw new Error('User not authenticated');
  }
  
  const ss = getOrCreateDatabase();
  
  // Check permissions
  const assessmentSheet = ss.getSheetByName(CONFIG.sheets.assessments);
  const assessmentData = assessmentSheet.getDataRange().getValues();
  const assessmentHeaders = assessmentData[0];
  const idCol = assessmentHeaders.indexOf('id');
  const teacherCol = assessmentHeaders.indexOf('teacher_email');
  
  let assessmentRow = -1;
  for (let i = 1; i < assessmentData.length; i++) {
    if (assessmentData[i][idCol] === assessmentId) {
      // Check permission
      if (user.role === CONFIG.roles.TEACHER && assessmentData[i][teacherCol] !== user.email) {
        throw new Error('Permission denied');
      }
      assessmentRow = i + 1;
      break;
    }
  }
  
  if (assessmentRow === -1) {
    throw new Error('Assessment not found');
  }
  
  // Delete from assessments
  assessmentSheet.deleteRow(assessmentRow);
  
  // Delete from raw data
  deleteRelatedRows(ss, CONFIG.sheets.rawData, 'assessment_id', assessmentId);
  
  // Delete from class periods
  deleteRelatedRows(ss, CONFIG.sheets.classPeriods, 'assessment_id', assessmentId);
  
  // Delete from analysis results
  deleteRelatedRows(ss, CONFIG.sheets.analysisResults, 'assessment_id', assessmentId);
  
  return { success: true };
}

/**
 * Helper to delete rows matching a value
 */
function deleteRelatedRows(ss, sheetName, columnName, value) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = headers.indexOf(columnName);
  
  if (col === -1) return;
  
  // Find rows to delete (from bottom to top)
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][col] === value) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // Delete from bottom to top
  rowsToDelete.reverse().forEach(row => {
    sheet.deleteRow(row);
  });
}

/**
 * Updates a user's role (admin function)
 */
function updateUserRole(email, newRole, newSchool) {
  const user = getCurrentUser();
  if (!user || user.role !== CONFIG.roles.DIVISION_ADMIN) {
    throw new Error('Division admin access required');
  }
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.users);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('email');
  const roleCol = headers.indexOf('role');
  const schoolCol = headers.indexOf('school');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] && data[i][emailCol].toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, roleCol + 1).setValue(newRole);
      if (newSchool) {
        sheet.getRange(i + 1, schoolCol + 1).setValue(newSchool);
      }
      return { success: true };
    }
  }
  
  throw new Error('User not found');
}

/**
 * Gets division-wide statistics
 */
function getDivisionStats() {
  const user = getCurrentUser();
  if (!user || user.role === CONFIG.roles.TEACHER) {
    throw new Error('Admin access required');
  }
  
  const ss = getOrCreateDatabase();
  
  // Get assessments
  const assessmentSheet = ss.getSheetByName(CONFIG.sheets.assessments);
  const assessmentData = assessmentSheet.getDataRange().getValues();
  const assessmentHeaders = assessmentData[0];
  
  const stats = {
    totalAssessments: 0,
    totalStudents: 0,
    bySchool: {},
    bySubject: {},
    byMonth: {}
  };
  
  for (let i = 1; i < assessmentData.length; i++) {
    const row = {};
    assessmentHeaders.forEach((h, idx) => {
      row[h] = assessmentData[i][idx];
    });
    
    // Filter by school for school admins
    if (user.role === CONFIG.roles.SCHOOL_ADMIN && row.school !== user.school) {
      continue;
    }
    
    stats.totalAssessments++;
    stats.totalStudents += row.student_count || 0;
    
    // By school
    const school = row.school || 'Unassigned';
    if (!stats.bySchool[school]) {
      stats.bySchool[school] = { assessments: 0, students: 0 };
    }
    stats.bySchool[school].assessments++;
    stats.bySchool[school].students += row.student_count || 0;
    
    // By subject
    if (row.subject) {
      if (!stats.bySubject[row.subject]) {
        stats.bySubject[row.subject] = 0;
      }
      stats.bySubject[row.subject]++;
    }
    
    // By month
    if (row.date_uploaded) {
      const date = new Date(row.date_uploaded);
      const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      if (!stats.byMonth[monthKey]) {
        stats.byMonth[monthKey] = 0;
      }
      stats.byMonth[monthKey]++;
    }
  }
  
  return stats;
}

/**
 * Gets default settings
 */
function getDefaultSettings() {
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.settings);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'default_settings') {
      try {
        return JSON.parse(data[i][1]);
      } catch (e) {
        return CONFIG.defaults;
      }
    }
  }
  
  return CONFIG.defaults;
}

/**
 * Saves default settings
 */
function saveDefaultSettings(settings) {
  const user = getCurrentUser();
  if (!user || user.role !== CONFIG.roles.DIVISION_ADMIN) {
    throw new Error('Division admin access required');
  }
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.settings);
  const data = sheet.getDataRange().getValues();
  
  // Find existing row or add new
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'default_settings') {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(settings));
      return { success: true };
    }
  }
  
  // Add new row
  sheet.appendRow(['default_settings', JSON.stringify(settings), '']);
  
  return { success: true };
}

/**
 * Gets unique teacher names for a given assessment (from ClassPeriods).
 * Used by the admin Rename tab.
 */
function getTeachersForAssessment(assessmentId) {
  var user = getCurrentUser();
  if (!user || user.role === CONFIG.roles.TEACHER) {
    throw new Error('Admin access required');
  }

  var normalId = String(assessmentId).trim();
  var ss = getOrCreateDatabase();
  var cpSheet = ss.getSheetByName(CONFIG.sheets.classPeriods);
  if (!cpSheet || cpSheet.getLastRow() < 2) return [];

  var data = cpSheet.getDataRange().getValues();
  var teachers = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] != null && String(data[i][0]).trim() === normalId) {
      var name = String(data[i][3] || '').trim();
      if (name) teachers[name] = true;
    }
  }
  return Object.keys(teachers).sort();
}

/**
 * Exports all data for a school (admin function)
 */
function exportSchoolData(school) {
  const user = getCurrentUser();
  if (!user || user.role === CONFIG.roles.TEACHER) {
    throw new Error('Admin access required');
  }
  
  if (user.role === CONFIG.roles.SCHOOL_ADMIN && user.school !== school) {
    throw new Error('Access denied to this school');
  }
  
  // Create export spreadsheet
  const ss = getOrCreateDatabase();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss");
  const exportSS = SpreadsheetApp.create(`${school} - Data Export - ${timestamp}`);
  
  // Copy relevant data
  const sheets = [CONFIG.sheets.assessments, CONFIG.sheets.classPeriods, CONFIG.sheets.analysisResults];
  
  sheets.forEach(sheetName => {
    const sourceSheet = ss.getSheetByName(sheetName);
    if (!sourceSheet) return;
    
    const data = sourceSheet.getDataRange().getValues();
    const headers = data[0];
    const schoolCol = headers.indexOf('school');
    const teacherCol = headers.indexOf('teacher_email') !== -1 ? headers.indexOf('teacher_email') : headers.indexOf('teacher');
    
    // Filter data
    const filteredData = [headers];
    for (let i = 1; i < data.length; i++) {
      if (schoolCol !== -1 && data[i][schoolCol] === school) {
        filteredData.push(data[i]);
      }
    }
    
    // Create sheet and add data
    const newSheet = exportSS.insertSheet(sheetName);
    if (filteredData.length > 0) {
      newSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    }
  });
  
  // Remove default sheet
  const defaultSheet = exportSS.getSheetByName('Sheet1');
  if (defaultSheet && exportSS.getSheets().length > 1) {
    exportSS.deleteSheet(defaultSheet);
  }
  
  return {
    url: exportSS.getUrl(),
    name: exportSS.getName()
  };
}
