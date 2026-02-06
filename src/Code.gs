/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                    MASTERYCONNECT ANALYZER WEB APP v1.0
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * A comprehensive web application for analyzing Mastery Connect benchmark data
 * 
 * Features:
 *   ✓ Role-based access (Teacher, School Admin, Division Admin)
 *   ✓ CSV file upload with drag-and-drop
 *   ✓ Class period assignment
 *   ✓ Item Analysis with distractor reports
 *   ✓ Small Group Interventions
 *   ✓ Visual Heatmaps
 *   ✓ District Summary (Admin only)
 *   ✓ Historical data storage
 *   ✓ Download reports (Excel/PDF)
 *   ✓ Email notifications
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
//                            CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════

const CONFIG = {
  // Google Sheet IDs (will be set after initial setup)
  sheets: {
    settings: 'Settings',
    users: 'Users',
    assessments: 'Assessments',
    rawData: 'RawData',
    classPeriods: 'ClassPeriods',
    analysisResults: 'AnalysisResults'
  },

  // Schools in the division
  schools: [
    'Kate Collins Middle School'
  ],

  // Color Palette - KCMS Purple & Gold Theme (Dark)
  colors: {
    primary: '#7C3AED',       // KCMS Purple
    primaryDark: '#6D28D9',   // Darker purple
    primaryLight: '#8B5CF6',  // Lighter purple
    gold: '#EAB308',          // KCMS Gold accent
    background: '#151225',    // Dark background
    surface: '#1e1a32',       // Card/surface background
    surfaceLight: '#28243f',  // Lighter surface
    text: '#ffffff',          // Primary text
    textMuted: '#a0aec0',     // Muted text
    growth: '#ef4444',        // Red - needs growth
    monitor: '#f59e0b',       // Yellow - monitor
    strength: '#10b981',      // Green - strength
    error: '#dc2626',         // Error red
    success: '#059669'        // Success green
  },
  
  // Default Thresholds
  defaults: {
    growth: 50,      // Below this % = needs growth
    strength: 75,    // Above this % = strength
    sgThreshold: 70, // Small group weakness threshold
    sgMax: 5,        // Max students per small group
    sgMin: 2         // Min students per small group
  },
  
  // Role definitions
  roles: {
    TEACHER: 'teacher',
    SCHOOL_ADMIN: 'school_admin',
    DIVISION_ADMIN: 'division_admin'
  }
};

// ═══════════════════════════════════════════════════════════════════════════
//                          WEB APP ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Serves the main web app HTML
 */
function doGet(e) {
  try {
    const user = getCurrentUser();
    
    // If no user or new user, show login/registration
    if (!user || user.isNewUser) {
      const loginTemplate = HtmlService.createTemplateFromFile('Login');
      loginTemplate.user = user;
      loginTemplate.config = CONFIG;
      loginTemplate.scriptUrl = ScriptApp.getService().getUrl();
      loginTemplate.urlParams = (e && e.parameter) ? e.parameter : {};
      return loginTemplate.evaluate()
        .setTitle('KCMS Benchmark Analyzer - Login')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // Safely get page parameter
    const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'dashboard';
    let template;
  
  switch (page) {
    case 'upload':
      template = HtmlService.createTemplateFromFile('Upload');
      break;
    case 'compare':
      template = HtmlService.createTemplateFromFile('Compare');
      break;
    case 'analysis':
      template = HtmlService.createTemplateFromFile('Analysis');
      // Pass assessment ID directly as a simple string
      template.assessmentId = (e && e.parameter && e.parameter.id) ? e.parameter.id : '';
      break;
    case 'analysis_test':
      // Debug page for testing
      template = HtmlService.createTemplateFromFile('Analysis_Test');
      template.assessmentId = (e && e.parameter && e.parameter.id) ? e.parameter.id : '';
      break;
    case 'history':
      template = HtmlService.createTemplateFromFile('History');
      break;
    case 'admin':
      if (user.role === CONFIG.roles.DIVISION_ADMIN || user.role === CONFIG.roles.SCHOOL_ADMIN) {
        template = HtmlService.createTemplateFromFile('Admin');
      } else {
        template = HtmlService.createTemplateFromFile('Dashboard');
      }
      break;
    default:
      template = HtmlService.createTemplateFromFile('Dashboard');
  }
  
  template.user = user;
  template.config = CONFIG;
  template.scriptUrl = ScriptApp.getService().getUrl();
  template.urlParams = (e && e.parameter) ? e.parameter : {};
  // Ensure assessmentId is always defined (mainly used by Analysis page)
  if (!template.assessmentId) {
    template.assessmentId = '';
  }
  
  return template.evaluate()
    .setTitle('KCMS Benchmark Analyzer')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
  } catch (error) {
    // Return error page if something goes wrong
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family: Arial; padding: 40px; background: #1a1a2e; color: white;">' +
      '<h1 style="color: #ef4444;">Error Loading Page</h1>' +
      '<p>Something went wrong: ' + error.message + '</p>' +
      '<p><a href="' + ScriptApp.getService().getUrl() + '" style="color: #7C3AED;">Return to Dashboard</a></p>' +
      '</body></html>'
    ).setTitle('Error - KCMS Benchmark Analyzer');
  }
}

/**
 * Handles POST requests (for larger data submissions)
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch (action) {
      case 'uploadCSV':
        return ContentService.createTextOutput(JSON.stringify(processCSVUpload(data)))
          .setMimeType(ContentService.MimeType.JSON);
      default:
        return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
          .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Include HTML files (for templating)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ═══════════════════════════════════════════════════════════════════════════
//                      AUTHENTICATION & USER MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Gets the current user's email and role
 */
function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    
    if (!email) {
      return null;
    }
    
    // Check if user exists in our system
    const userData = getUserData(email);
    
    if (!userData) {
      // New user - return basic info for registration
      return {
        email: email,
        name: email.split('@')[0],
        role: null,
        school: null,
        settings: CONFIG.defaults,
        isNewUser: true
      };
    }
    
    return userData;
  } catch (error) {
    Logger.log('getCurrentUser error: ' + error.message);
    return null;
  }
}

/**
 * Gets user data from the Users sheet
 */
function getUserData(email) {
  try {
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.users);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('email');
    
    if (emailCol === -1) {
      return null;
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] && data[i][emailCol].toString().toLowerCase() === email.toLowerCase()) {
        let settings = CONFIG.defaults;
        try {
          const settingsStr = data[i][headers.indexOf('settings')];
          if (settingsStr && settingsStr !== '') {
            settings = JSON.parse(settingsStr);
          }
        } catch (e) {
          settings = CONFIG.defaults;
        }
        
        return {
          email: data[i][emailCol],
          name: data[i][headers.indexOf('name')] || email.split('@')[0],
          role: data[i][headers.indexOf('role')] || CONFIG.roles.TEACHER,
          school: data[i][headers.indexOf('school')] || '',
          settings: settings,
          isNewUser: false
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('getUserData error: ' + error.message);
    return null;
  }
}

/**
 * Registers a new user or updates existing user
 */
function registerUser(name, school) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Not authenticated');
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.users);
  
  // Check if user already exists
  const existingUser = getUserData(email);
  if (existingUser && !existingUser.isNewUser) {
    throw new Error('User already registered');
  }
  
  // Determine role - check if admin
  const role = checkIfAdmin(email) || CONFIG.roles.TEACHER;
  
  // Add user
  sheet.appendRow([
    email,
    name,
    role,
    school,
    JSON.stringify(CONFIG.defaults),
    new Date().toISOString()
  ]);
  
  return {
    email: email,
    name: name,
    role: role,
    school: school,
    settings: CONFIG.defaults,
    isNewUser: false
  };
}

/**
 * Checks if email is in admin list
 */
function checkIfAdmin(email) {
  const ss = getOrCreateDatabase();
  const settingsSheet = ss.getSheetByName(CONFIG.sheets.settings);
  
  if (!settingsSheet) return null;
  
  const data = settingsSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'division_admin' && data[i][1].toLowerCase() === email.toLowerCase()) {
      return CONFIG.roles.DIVISION_ADMIN;
    }
    if (data[i][0] === 'school_admin' && data[i][1].toLowerCase() === email.toLowerCase()) {
      return CONFIG.roles.SCHOOL_ADMIN;
    }
  }
  
  return null;
}

/**
 * Updates user settings
 */
function updateUserSettings(settings) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Not authenticated');
  
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.users);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('email');
  const settingsCol = headers.indexOf('settings');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] && data[i][emailCol].toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, settingsCol + 1).setValue(JSON.stringify(settings));
      return true;
    }
  }
  
  throw new Error('User not found');
}

// ═══════════════════════════════════════════════════════════════════════════
//                          DATABASE MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Gets or creates the main database spreadsheet
 */
function getOrCreateDatabase() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('DATABASE_ID');
  
  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {
      // Database not found, create new
    }
  }
  
  // Create new database
  const ss = SpreadsheetApp.create('KCMS Benchmark Analyzer Database');
  ssId = ss.getId();
  props.setProperty('DATABASE_ID', ssId);
  
  // Initialize sheets
  initializeDatabase(ss);
  
  return ss;
}

/**
 * Initializes the database structure
 */
function initializeDatabase(ss) {
  // Settings sheet
  let sheet = ss.getSheetByName('Sheet1');
  if (sheet) {
    sheet.setName(CONFIG.sheets.settings);
  } else {
    sheet = ss.insertSheet(CONFIG.sheets.settings);
  }
  sheet.getRange(1, 1, 1, 3).setValues([['type', 'email', 'school']]);
  
  // Users sheet
  sheet = ss.insertSheet(CONFIG.sheets.users);
  sheet.getRange(1, 1, 1, 6).setValues([['email', 'name', 'role', 'school', 'settings', 'created_at']]);
  
  // Assessments sheet
  sheet = ss.insertSheet(CONFIG.sheets.assessments);
  sheet.getRange(1, 1, 1, 12).setValues([[
    'id', 'name', 'school', 'teacher_email', 'subject', 'grade_level', 
    'assessment_type', 'date_administered', 'date_uploaded', 'student_count', 
    'question_count', 'status'
  ]]);
  
  // Raw Data sheet
  sheet = ss.insertSheet(CONFIG.sheets.rawData);
  sheet.getRange(1, 1, 1, 4).setValues([['assessment_id', 'row_index', 'row_type', 'data']]);
  
  // Class Periods sheet
  sheet = ss.insertSheet(CONFIG.sheets.classPeriods);
  sheet.getRange(1, 1, 1, 5).setValues([['assessment_id', 'student_id', 'student_name', 'teacher', 'period']]);
  
  // Analysis Results sheet
  sheet = ss.insertSheet(CONFIG.sheets.analysisResults);
  sheet.getRange(1, 1, 1, 6).setValues([['assessment_id', 'analysis_type', 'teacher', 'created_at', 'settings', 'data']]);
}

/**
 * Gets the database spreadsheet ID for client-side use
 */
function getDatabaseId() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('DATABASE_ID');
}

// ═══════════════════════════════════════════════════════════════════════════
//                          CSV PROCESSING
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Gets existing assessments for the current user (for duplicate checking on Upload page)
 */
function getUserExistingAssessments() {
  try {
    const user = getCurrentUser();
    if (!user) return [];
    
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.assessments);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const assessments = [];
    const idCol = headers.indexOf('id');
    const nameCol = headers.indexOf('name');
    const teacherCol = headers.indexOf('teacher_email');
    const dateUploadedCol = headers.indexOf('date_uploaded');
    const studentCountCol = headers.indexOf('student_count');
    const subjectCol = headers.indexOf('subject');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idCol]) continue;
      
      // Only return assessments for this teacher (or all for admins)
      if (user.role === CONFIG.roles.TEACHER && row[teacherCol] !== user.email) {
        continue;
      }
      
      assessments.push({
        id: row[idCol],
        name: row[nameCol],
        dateUploaded: row[dateUploadedCol],
        studentCount: row[studentCountCol],
        subject: row[subjectCol]
      });
    }
    
    return assessments;
  } catch (error) {
    Logger.log('getUserExistingAssessments error: ' + error.message);
    return [];
  }
}

/**
 * Checks for duplicate assessment before upload
 * Returns existing assessment if found, null otherwise
 */
function checkForDuplicateAssessment(metadata, csvContent) {
  try {
    const user = getCurrentUser();
    if (!user) return null;
    
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.assessments);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return null; // No existing assessments
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.indexOf('id');
    const nameCol = headers.indexOf('name');
    const teacherCol = headers.indexOf('teacher_email');
    const dateCol = headers.indexOf('date_administered');
    const subjectCol = headers.indexOf('subject');
    const studentCountCol = headers.indexOf('student_count');
    
    // Parse CSV to get student count for matching
    const rows = Utilities.parseCsv(csvContent);
    const uploadStudentCount = rows.length - 2;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idCol]) continue;
      
      // Match criteria: same name + same teacher + same date + same student count
      const nameMatch = row[nameCol] === metadata.name;
      const teacherMatch = row[teacherCol] === user.email;
      const dateMatch = row[dateCol] === metadata.dateAdministered;
      const countMatch = row[studentCountCol] === uploadStudentCount;
      
      // If name, teacher, and either date or student count match, it's likely a duplicate
      if (nameMatch && teacherMatch && (dateMatch || countMatch)) {
        return {
          isDuplicate: true,
          existingId: row[idCol],
          existingName: row[nameCol],
          existingDate: row[headers.indexOf('date_uploaded')],
          message: `An assessment named "${row[nameCol]}" was already uploaded on ${new Date(row[headers.indexOf('date_uploaded')]).toLocaleDateString()}.`
        };
      }
    }
    
    return null; // No duplicate found
  } catch (error) {
    Logger.log('checkForDuplicateAssessment error: ' + error.message);
    return null;
  }
}

/**
 * Processes uploaded CSV data
 */
function processCSVUpload(uploadData) {
  const user = getCurrentUser();
  if (!user || user.isNewUser) {
    throw new Error('User not authenticated or not registered');
  }
  
  const csvContent = uploadData.csvContent;
  const metadata = uploadData.metadata;
  
  // Check for duplicate first
  const duplicate = checkForDuplicateAssessment(metadata, csvContent);
  if (duplicate) {
    return {
      success: false,
      isDuplicate: true,
      existingId: duplicate.existingId,
      existingName: duplicate.existingName,
      message: duplicate.message
    };
  }
  
  // Parse CSV
  const rows = Utilities.parseCsv(csvContent);
  
  if (rows.length < 3) {
    throw new Error('CSV must have at least 3 rows (SOL row, header row, and data rows)');
  }
  
  // Validate required columns
  const header = rows[1];
  const requiredCols = ['teacher', 'first_name', 'last_name', 'percentage', 'student_id'];
  const missingCols = requiredCols.filter(col => {
    var lowerCol = col.toLowerCase();
    return !header.some(function(h) {
      if (!h) return false;
      var lowerH = h.toString().toLowerCase();
      // Match exact name or with underscores/spaces swapped
      return lowerH === lowerCol || lowerH === lowerCol.replace(/_/g, ' ') || lowerH.replace(/ /g, '_') === lowerCol;
    });
  });

  if (missingCols.length > 0) {
    throw new Error('Missing required columns: ' + missingCols.join(', '));
  }

  // Case-insensitive column finder (matching AnalysisFunctions.gs findColIndex)
  function findCol(colName) {
    var lowerName = colName.toLowerCase();
    var idx = header.findIndex(function(h) { return h && h.toString().toLowerCase() === lowerName; });
    if (idx !== -1) return idx;
    var altName = lowerName.indexOf('_') !== -1 ? lowerName.replace(/_/g, ' ') : lowerName.replace(/ /g, '_');
    return header.findIndex(function(h) { return h && h.toString().toLowerCase() === altName; });
  }

  // Find teacher column to filter data for the current teacher
  const teacherColIdx = findCol('teacher');
  
  // Filter rows to only include the current teacher's students (for teachers)
  // Admins can upload all data
  let filteredRows = rows;
  if (user.role === CONFIG.roles.TEACHER && teacherColIdx !== -1) {
    // Keep header rows (0 and 1) plus only student rows for this teacher
    const teacherName = user.name || user.email.split('@')[0];
    filteredRows = [rows[0], rows[1]]; // SOL row and header row
    
    for (let i = 2; i < rows.length; i++) {
      const rowTeacher = rows[i][teacherColIdx];
      // Match by teacher name (case-insensitive, partial match)
      if (rowTeacher && (
        rowTeacher.toLowerCase().includes(teacherName.toLowerCase()) ||
        teacherName.toLowerCase().includes(rowTeacher.toLowerCase().split(',')[0])
      )) {
        filteredRows.push(rows[i]);
      }
    }
    
    // If no matching students found, use all rows (might be different naming)
    if (filteredRows.length <= 2) {
      filteredRows = rows;
    }
  }
  
  // Generate assessment ID
  const assessmentId = Utilities.getUuid();
  
  // Store in database
  const ss = getOrCreateDatabase();
  
  // Store assessment metadata
  const assessmentSheet = ss.getSheetByName(CONFIG.sheets.assessments);
  const studentCount = filteredRows.length - 2;
  const pctColIdx = findCol('percentage');
  const questionCount = pctColIdx !== -1 ? Math.floor((header.length - pctColIdx - 1) / 2) : 0;
  
  assessmentSheet.appendRow([
    assessmentId,
    metadata.name || 'Unnamed Assessment',
    metadata.school || user.school,
    user.email,
    metadata.subject || '',
    metadata.gradeLevel || '',
    metadata.assessmentType || '',
    metadata.dateAdministered || '',
    new Date().toISOString(),
    studentCount,
    questionCount,
    'pending_periods'
  ]);
  
  // Store raw data
  const rawDataSheet = ss.getSheetByName(CONFIG.sheets.rawData);
  
  // Store SOL row
  rawDataSheet.appendRow([assessmentId, 0, 'sol', JSON.stringify(filteredRows[0])]);
  
  // Store header row
  rawDataSheet.appendRow([assessmentId, 1, 'header', JSON.stringify(filteredRows[1])]);
  
  // Store student rows
  for (let i = 2; i < filteredRows.length; i++) {
    rawDataSheet.appendRow([assessmentId, i, 'student', JSON.stringify(filteredRows[i])]);
  }
  
  // Prepare class period data
  const classPeriodSheet = ss.getSheetByName(CONFIG.sheets.classPeriods);
  const firstNameCol = findCol('first_name');
  const lastNameCol = findCol('last_name');
  const studentIdCol = findCol('student_id');
  
  const students = [];
  for (let i = 2; i < filteredRows.length; i++) {
    const row = filteredRows[i];
    const studentName = `${row[firstNameCol] || ''} ${row[lastNameCol] || ''}`.trim();
    
    students.push({
      studentId: row[studentIdCol],
      studentName: studentName,
      teacher: row[teacherColIdx],
      period: ''
    });
    
    // Add to class periods sheet
    classPeriodSheet.appendRow([
      assessmentId,
      row[studentIdCol],
      studentName,
      row[teacherColIdx],
      ''
    ]);
  }
  
  return {
    success: true,
    assessmentId: assessmentId,
    studentCount: studentCount,
    questionCount: questionCount,
    students: students,
    message: `Successfully uploaded ${studentCount} students with ${questionCount} questions.`
  };
}

/**
 * Updates class periods for an assessment
 */
function updateClassPeriods(assessmentId, periodData) {
  const user = getCurrentUser();
  if (!user || user.isNewUser) {
    throw new Error('User not authenticated');
  }

  const normalId = assessmentId.toString().trim();
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.classPeriods);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf('assessment_id');
  const studentIdCol = headers.indexOf('student_id');
  const periodCol = headers.indexOf('period');

  // Create lookup for period data
  const periodMap = {};
  periodData.forEach(p => {
    periodMap[p.studentId] = p.period;
  });

  // Update periods
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === normalId && periodMap[data[i][studentIdCol]] !== undefined) {
      sheet.getRange(i + 1, periodCol + 1).setValue(periodMap[data[i][studentIdCol]]);
    }
  }

  // Update assessment status
  const assessmentSheet = ss.getSheetByName(CONFIG.sheets.assessments);
  const assessmentData = assessmentSheet.getDataRange().getValues();
  const assessmentHeaders = assessmentData[0];
  const assessmentIdCol = assessmentHeaders.indexOf('id');
  const statusCol = assessmentHeaders.indexOf('status');

  for (let i = 1; i < assessmentData.length; i++) {
    if (String(assessmentData[i][assessmentIdCol]).trim() === normalId) {
      assessmentSheet.getRange(i + 1, statusCol + 1).setValue('ready');
      break;
    }
  }
  
  return { success: true, message: 'Class periods updated successfully' };
}

/**
 * Gets assessment data for a specific assessment
 */
function getAssessmentData(assessmentId) {
  try {
    const user = getCurrentUser();
    if (!user || user.isNewUser) {
      throw new Error('User not authenticated');
    }

    if (!assessmentId) {
      throw new Error('No assessment ID provided');
    }

    // Normalize the ID for comparison (Sheets can coerce types)
    const normalId = String(assessmentId).trim();

    const ss = getOrCreateDatabase();

    // Get assessment metadata
    const assessmentSheet = ss.getSheetByName(CONFIG.sheets.assessments);
    if (!assessmentSheet || assessmentSheet.getLastRow() < 2) {
      throw new Error('Assessments sheet is empty or missing');
    }

    const assessmentData = assessmentSheet.getDataRange().getValues();
    const assessmentHeaders = assessmentData[0];

    // Find the ID column robustly - try exact, then case-insensitive, then column 0
    var idColIdx = assessmentHeaders.indexOf('id');
    if (idColIdx === -1) {
      for (var hi = 0; hi < assessmentHeaders.length; hi++) {
        if (String(assessmentHeaders[hi]).toLowerCase().trim() === 'id') {
          idColIdx = hi;
          break;
        }
      }
    }
    if (idColIdx === -1) {
      idColIdx = 0; // assume first column is the ID
    }

    let assessment = null;

    for (let i = 1; i < assessmentData.length; i++) {
      var cellVal = assessmentData[i][idColIdx];
      if (cellVal != null && String(cellVal).trim() === normalId) {
        assessment = {};
        assessmentHeaders.forEach((h, idx) => {
          var val = assessmentData[i][idx];
          // Convert Date objects to ISO strings so google.script.run serializes safely
          if (val instanceof Date) {
            val = val.toISOString();
          }
          assessment[String(h).trim()] = val;
        });
        break;
      }
    }

    if (!assessment) {
      // Diagnostic info to help debug
      var rowCount = assessmentData.length - 1;
      var sampleIds = [];
      for (var di = 1; di < Math.min(assessmentData.length, 4); di++) {
        sampleIds.push(String(assessmentData[di][idColIdx]).substring(0, 12));
      }
      throw new Error(
        'Assessment not found. ID="' + normalId.substring(0, 12) + '..." ' +
        '(col=' + idColIdx + ', rows=' + rowCount +
        ', samples=[' + sampleIds.join(', ') + '...])'
      );
    }

    // Check permissions - teachers can only see their own assessments
    if (user.role === CONFIG.roles.TEACHER && assessment.teacher_email !== user.email) {
      throw new Error('Access denied - you can only view your own assessments');
    }

    // Get raw data
    const rawDataSheet = ss.getSheetByName(CONFIG.sheets.rawData);
    if (!rawDataSheet || rawDataSheet.getLastRow() < 2) {
      return {
        assessment: assessment,
        solRow: [],
        header: [],
        students: [],
        periodMap: {}
      };
    }

    const rawData = rawDataSheet.getDataRange().getValues();

    let solRow = [];
    let header = [];
    const students = [];

    for (let i = 1; i < rawData.length; i++) {
      if (rawData[i][0] != null && String(rawData[i][0]).trim() === normalId) {
        const rowType = rawData[i][2];
        try {
          const data = JSON.parse(rawData[i][3]);

          if (rowType === 'sol') solRow = data;
          else if (rowType === 'header') header = data;
          else if (rowType === 'student') students.push(data);
        } catch (parseError) {
          Logger.log('Error parsing row ' + i + ': ' + parseError.message);
        }
      }
    }

    // Get class periods
    const periodSheet = ss.getSheetByName(CONFIG.sheets.classPeriods);
    const periodMap = {};

    if (periodSheet && periodSheet.getLastRow() > 1) {
      const periodData = periodSheet.getDataRange().getValues();

      for (let i = 1; i < periodData.length; i++) {
        if (periodData[i][0] != null && String(periodData[i][0]).trim() === normalId) {
          periodMap[periodData[i][1]] = periodData[i][4]; // student_id -> period
        }
      }
    }

    return {
      assessment: assessment,
      solRow: solRow,
      header: header,
      students: students,
      periodMap: periodMap
    };
  } catch (error) {
    Logger.log('getAssessmentData error: ' + error.message);
    throw error;
  }
}

/**
 * Gets list of assessments for current user
 */
function getAssessmentList() {
  try {
    const user = getCurrentUser();
    if (!user || user.isNewUser) {
      return []; // Return empty array instead of throwing
    }
    
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.assessments);
    
    if (!sheet) {
      return []; // No assessments sheet yet
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return []; // Only header row, no data
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const assessments = [];
    
    for (let i = 1; i < data.length; i++) {
      // Skip empty rows
      if (!data[i][0]) continue;
      
      const assessment = {};
      headers.forEach((h, idx) => {
        assessment[h] = data[i][idx];
      });
      
      // Filter based on role
      if (user.role === CONFIG.roles.TEACHER) {
        if (assessment.teacher_email === user.email) {
          assessments.push(assessment);
        }
      } else if (user.role === CONFIG.roles.SCHOOL_ADMIN) {
        if (assessment.school === user.school) {
          assessments.push(assessment);
        }
      } else {
        // Division admin sees all
        assessments.push(assessment);
      }
    }
    
    // Sort by date uploaded, newest first
    assessments.sort((a, b) => new Date(b.date_uploaded) - new Date(a.date_uploaded));
    
    return assessments;
  } catch (error) {
    Logger.log('getAssessmentList error: ' + error.message);
    return []; // Return empty array on error
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//                          HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Gets schools list
 */
function getSchoolsList() {
  return CONFIG.schools;
}

/**
 * Gets configuration for client-side use
 */
function getClientConfig() {
  return {
    schools: CONFIG.schools,
    colors: CONFIG.colors,
    defaults: CONFIG.defaults,
    roles: CONFIG.roles
  };
}
