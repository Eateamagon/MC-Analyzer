/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                          ANALYSIS FUNCTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Runs the complete analysis for an assessment
 */
function runAnalysis(assessmentId, options) {
  const user = getCurrentUser();
  if (!user || user.isNewUser) {
    throw new Error('User not authenticated');
  }
  
  // Get assessment data
  const assessmentData = getAssessmentData(assessmentId);
  const { assessment, solRow, header, students, periodMap } = assessmentData;
  
  // Get user settings
  const settings = user.settings || CONFIG.defaults;
  
  // Parse questions from header
  const questions = parseQuestions(header, solRow);
  
  // Get column indices
  const teacherCol = header.indexOf('teacher');
  const firstNameCol = header.indexOf('first_name');
  const lastNameCol = header.indexOf('last_name');
  const studentIdCol = header.indexOf('student_id');
  const pctCol = header.indexOf('percentage');
  
  const results = {
    assessmentId: assessmentId,
    assessmentName: assessment.name,
    timestamp: new Date().toISOString(),
    settings: settings,
    options: options,
    itemAnalysis: null,
    smallGroups: null,
    heatmaps: null,
    districtSummary: null,
    teacherSummaries: []
  };
  
  // Get unique teachers
  const teachers = [...new Set(students.map(r => r[teacherCol]))].sort();
  
  // Run Item Analysis
  if (options.itemAnalysis) {
    results.itemAnalysis = {};
    teachers.forEach(teacher => {
      const teacherStudents = students.filter(r => r[teacherCol] === teacher);
      results.itemAnalysis[teacher] = generateItemAnalysis(teacherStudents, questions, pctCol);
    });
  }
  
  // Run Small Groups
  if (options.smallGroups) {
    results.smallGroups = generateSmallGroups(
      students, questions, teacherCol, firstNameCol, lastNameCol, 
      studentIdCol, settings, periodMap
    );
  }
  
  // Run Heatmaps
  if (options.heatmaps) {
    results.heatmaps = {
      studentHeatmap: generateStudentHeatmap(students, questions, teacherCol, firstNameCol, lastNameCol),
      solHeatmap: generateSOLHeatmap(students, questions, teacherCol)
    };
  }
  
  // Generate Teacher Summaries
  teachers.forEach(teacher => {
    const teacherStudents = students.filter(r => r[teacherCol] === teacher);
    const summary = generateTeacherSummary(teacher, teacherStudents, questions, firstNameCol, lastNameCol, settings);
    results.teacherSummaries.push(summary);
  });
  
  // Generate District Summary (for admins)
  if (user.role !== CONFIG.roles.TEACHER) {
    results.districtSummary = generateDistrictSummary(results.teacherSummaries);
  }
  
  // Store results
  storeAnalysisResults(assessmentId, results);
  
  // Send email notification if requested
  if (options.sendEmail) {
    sendAnalysisEmail(user.email, results);
  }
  
  return results;
}

/**
 * Parses questions from header row
 */
function parseQuestions(header, solRow) {
  const pctCol = header.indexOf('percentage');
  const firstQ = pctCol + 1;
  const qCount = Math.floor((header.length - firstQ) / 2);
  const questions = [];
  
  for (let i = 0; i < qCount; i++) {
    const col = firstQ + (i * 2);
    if (col >= header.length) break;
    
    questions.push({
      number: i + 1,
      sol: solRow[col] || `Q${i + 1}`,
      answer: header[col],
      answerColIndex: col,
      scoreColIndex: col + 1
    });
  }
  
  return questions;
}

/**
 * Generates Item Analysis with distractor report
 */
function generateItemAnalysis(students, questions, pctCol) {
  // Sort students to identify high performers (top 25%)
  const sortedStudents = [...students].sort((a, b) => b[pctCol] - a[pctCol]);
  const cutoff = Math.ceil(sortedStudents.length * 0.25);
  const highPerformers = sortedStudents.slice(0, cutoff);
  
  const analysis = [];
  
  questions.forEach(q => {
    let correct = 0;
    let correctAnswer = '';
    const answerCounts = {};
    
    students.forEach(s => {
      const ans = s[q.answerColIndex];
      const score = s[q.scoreColIndex];
      
      if (score == 1) {
        correct++;
        if (!correctAnswer && ans) correctAnswer = ans;
      }
      
      if (ans && score != 1) {
        answerCounts[ans] = (answerCounts[ans] || 0) + 1;
      }
    });
    
    const pctCorrect = students.length > 0 ? Math.round((correct / students.length) * 100) : 0;
    
    // Find most common wrong answer
    let topWrong = '-';
    let topCount = 0;
    const sortedWrong = Object.entries(answerCounts).sort((a, b) => b[1] - a[1]);
    
    if (sortedWrong.length > 0) {
      topWrong = sortedWrong[0][0];
      topCount = sortedWrong[0][1];
    }
    
    // All distractors
    const allDistractors = sortedWrong.map(([ans, count]) => ({ answer: ans, count: count }));
    
    // Check if high performers struggled
    let hpWrong = 0;
    highPerformers.forEach(hp => {
      if (hp[q.scoreColIndex] != 1) hpWrong++;
    });
    
    const reviewFlag = highPerformers.length > 0 && (hpWrong / highPerformers.length) > 0.5;
    
    analysis.push({
      question: `Q${q.number}`,
      sol: q.sol,
      correctAnswer: correctAnswer || '-',
      pctCorrect: pctCorrect,
      topWrongAnswer: topWrong,
      topWrongCount: topCount,
      allDistractors: allDistractors,
      reviewFlag: reviewFlag
    });
  });
  
  return analysis;
}

/**
 * Generates Small Group Interventions
 */
function generateSmallGroups(students, questions, teacherCol, firstNameCol, lastNameCol, studentIdCol, settings, periodMap) {
  const teachers = [...new Set(students.map(r => r[teacherCol]))].sort();
  const groups = [];
  
  teachers.forEach(teacher => {
    const teacherStudents = students.filter(r => r[teacherCol] === teacher);
    
    // Group by period
    const byPeriod = {};
    teacherStudents.forEach(r => {
      const period = (periodMap && periodMap[r[studentIdCol]]) ? periodMap[r[studentIdCol]] : 'Unknown';
      if (!byPeriod[period]) byPeriod[period] = [];
      byPeriod[period].push(r);
    });
    
    Object.keys(byPeriod).sort().forEach(period => {
      const profiles = [];
      
      byPeriod[period].forEach(r => {
        const name = `${r[firstNameCol] || ''} ${r[lastNameCol] || ''}`.trim();
        const solStats = {};
        
        questions.forEach(q => {
          const sc = r[q.scoreColIndex];
          if (!solStats[q.sol]) solStats[q.sol] = { c: 0, t: 0 };
          if (sc == 1 || sc == '1') { solStats[q.sol].c++; solStats[q.sol].t++; }
          else if (sc == 0 || sc == '0') { solStats[q.sol].t++; }
        });
        
        const weak = [];
        Object.keys(solStats).forEach(s => {
          if (solStats[s].t > 0 && ((solStats[s].c / solStats[s].t) * 100) < settings.sgThreshold) {
            weak.push(s);
          }
        });
        
        if (weak.length > 0) {
          profiles.push({ name: name, sols: weak, id: r[studentIdCol] });
        }
      });
      
      // Form groups using the algorithm
      const periodGroups = formGroups(profiles, settings);
      
      periodGroups.forEach((g, idx) => {
        groups.push({
          teacher: teacher,
          period: period,
          groupNumber: idx + 1,
          students: g.students.map(s => s.name),
          sharedWeakSOLs: g.shared,
          studentCount: g.students.length,
          isUnknownPeriod: period === 'Unknown'
        });
      });
    });
  });
  
  return groups;
}

/**
 * Group formation algorithm
 */
function formGroups(list, settings) {
  const max = settings.sgMax || 5;
  const min = settings.sgMin || 2;
  const rem = [...list].sort((a, b) => b.sols.length - a.sols.length);
  const groups = [];
  
  while (rem.length > 0) {
    const core = rem.shift();
    let grp = [core];
    let shared = [...core.sols];
    
    for (let i = rem.length - 1; i >= 0; i--) {
      const c = rem[i];
      const ov = c.sols.filter(s => shared.includes(s));
      if (ov.length >= Math.ceil(shared.length * 0.5)) {
        grp.push(c);
        shared = shared.filter(s => c.sols.includes(s));
        rem.splice(i, 1);
      }
    }
    
    if (grp.length < min) {
      for (let i = rem.length - 1; i >= 0 && grp.length < min; i--) {
        const c = rem[i];
        if (c.sols.some(s => shared.includes(s))) {
          grp.push(c);
          rem.splice(i, 1);
        }
      }
    }
    
    if (grp.length > max) {
      const keep = grp.splice(0, max);
      groups.push({ students: keep, shared: shared });
      rem.unshift(...grp);
    } else {
      groups.push({ students: grp, shared: shared });
    }
  }
  
  return groups;
}

/**
 * Generates Student Heatmap data
 */
function generateStudentHeatmap(students, questions, teacherCol, firstNameCol, lastNameCol) {
  const data = [];
  
  students.forEach(r => {
    const teacher = r[teacherCol];
    const name = `${r[firstNameCol] || ''} ${r[lastNameCol] || ''}`.trim();
    let correct = 0;
    let total = 0;
    
    const questionScores = questions.map(q => {
      const sc = r[q.scoreColIndex];
      if (sc == 1 || sc == '1') { correct++; total++; return 1; }
      if (sc == 0 || sc == '0') { total++; return 0; }
      return null;
    });
    
    const pct = total > 0 ? Math.round((correct / total) * 100) : 0;
    
    data.push({
      teacher: teacher,
      student: name,
      scorePct: pct,
      questionScores: questionScores
    });
  });
  
  // Calculate averages for each question
  const averages = questions.map((q, idx) => {
    let qCorrect = 0;
    let qTotal = 0;
    students.forEach(r => {
      const sc = r[q.scoreColIndex];
      if (sc == 1 || sc == '1') { qCorrect++; qTotal++; }
      if (sc == 0 || sc == '0') { qTotal++; }
    });
    return qTotal > 0 ? Math.round((qCorrect / qTotal) * 100) : 0;
  });
  
  return {
    headers: ['Teacher', 'Student', 'Score %', ...questions.map(q => `Q${q.number} (${q.sol})`)],
    rows: data,
    averages: averages
  };
}

/**
 * Generates SOL Heatmap data
 */
function generateSOLHeatmap(students, questions, teacherCol) {
  const sols = [...new Set(questions.map(q => q.sol))].sort();
  const stats = {};
  const global = {};
  
  sols.forEach(s => global[s] = { c: 0, t: 0 });
  
  students.forEach(r => {
    const t = r[teacherCol];
    if (!t) return;
    if (!stats[t]) {
      stats[t] = {};
      sols.forEach(s => stats[t][s] = { c: 0, t: 0 });
    }
    
    questions.forEach(q => {
      const sc = r[q.scoreColIndex];
      if (sc == 1 || sc == '1') { stats[t][q.sol].c++; global[q.sol].c++; }
      if (sc == 1 || sc == 0 || sc == '1' || sc == '0') { stats[t][q.sol].t++; global[q.sol].t++; }
    });
  });
  
  const rows = Object.keys(stats).map(teacher => {
    const solPcts = sols.map(s => stats[teacher][s].t ? Math.round((stats[teacher][s].c / stats[teacher][s].t) * 100) : null);
    return {
      teacher: teacher,
      solPercentages: solPcts
    };
  });
  
  const averages = sols.map(s => global[s].t ? Math.round((global[s].c / global[s].t) * 100) : null);
  
  return {
    headers: ['Teacher', ...sols],
    rows: rows,
    averages: averages
  };
}

/**
 * Generates Teacher Summary
 */
function generateTeacherSummary(teacher, students, questions, firstNameCol, lastNameCol, settings) {
  const low = settings.growth || 50;
  const high = settings.strength || 75;
  
  const solStats = {};
  let totalC = 0;
  let totalA = 0;
  
  questions.forEach(q => {
    if (!solStats[q.sol]) solStats[q.sol] = { c: 0, t: 0, red: 0, yel: 0, grn: 0 };
  });
  
  const studentSolMap = {};
  
  students.forEach((s, idx) => {
    studentSolMap[idx] = {};
    questions.forEach(q => {
      if (!studentSolMap[idx][q.sol]) studentSolMap[idx][q.sol] = { c: 0, t: 0 };
      
      const sc = s[q.scoreColIndex];
      if (sc == 1 || sc == '1') { solStats[q.sol].c++; totalC++; }
      if (sc == 1 || sc == 0 || sc == '1' || sc == '0') { solStats[q.sol].t++; totalA++; }
      
      if (sc == 1 || sc == '1') studentSolMap[idx][q.sol].c++;
      if (sc == 1 || sc == 0 || sc == '1' || sc == '0') studentSolMap[idx][q.sol].t++;
    });
  });
  
  // Count students in each category per SOL
  Object.keys(studentSolMap).forEach(idx => {
    const map = studentSolMap[idx];
    Object.keys(map).forEach(sol => {
      if (map[sol].t > 0) {
        const pct = (map[sol].c / map[sol].t) * 100;
        if (pct >= high) solStats[sol].grn++;
        else if (pct >= low) solStats[sol].yel++;
        else solStats[sol].red++;
      }
    });
  });
  
  const avg = totalA > 0 ? Math.round((totalC / totalA) * 100) : 0;
  
  const data = Object.keys(solStats).map(k => ({
    sol: k,
    correct: solStats[k].c,
    total: solStats[k].t,
    percentage: solStats[k].t ? Math.round((solStats[k].c / solStats[k].t) * 100) : 0,
    redCount: solStats[k].red,
    yellowCount: solStats[k].yel,
    greenCount: solStats[k].grn
  })).sort((a, b) => a.percentage - b.percentage);
  
  // Categorize SOLs
  const growthSOLs = data.filter(d => d.percentage < low);
  const monitorSOLs = data.filter(d => d.percentage >= low && d.percentage < high);
  const strengthSOLs = data.filter(d => d.percentage >= high);
  
  return {
    teacher: teacher,
    studentCount: students.length,
    average: avg,
    bestSOL: data.length > 0 ? `${data[data.length - 1].sol} (${data[data.length - 1].percentage}%)` : '-',
    worstSOL: data.length > 0 ? `${data[0].sol} (${data[0].percentage}%)` : '-',
    growthSOLs: growthSOLs,
    monitorSOLs: monitorSOLs,
    strengthSOLs: strengthSOLs,
    allSOLs: data
  };
}

/**
 * Generates District Summary
 */
function generateDistrictSummary(teacherSummaries) {
  const summary = teacherSummaries.map(ts => ({
    teacher: ts.teacher,
    studentCount: ts.studentCount,
    average: ts.average,
    bestSOL: ts.bestSOL,
    worstSOL: ts.worstSOL
  })).sort((a, b) => b.average - a.average);
  
  const totalStudents = summary.reduce((acc, s) => acc + s.studentCount, 0);
  const overallAverage = summary.length > 0 
    ? Math.round(summary.reduce((acc, s) => acc + s.average, 0) / summary.length) 
    : 0;
  
  return {
    teachers: summary,
    totalStudents: totalStudents,
    overallAverage: overallAverage,
    teacherCount: summary.length
  };
}

/**
 * Stores analysis results in the database
 */
function storeAnalysisResults(assessmentId, results) {
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);
  
  // Store each type of analysis separately
  const timestamp = new Date().toISOString();
  
  if (results.itemAnalysis) {
    sheet.appendRow([
      assessmentId,
      'itemAnalysis',
      JSON.stringify(Object.keys(results.itemAnalysis)),
      timestamp,
      JSON.stringify(results.settings),
      JSON.stringify(results.itemAnalysis)
    ]);
  }
  
  if (results.smallGroups) {
    sheet.appendRow([
      assessmentId,
      'smallGroups',
      '',
      timestamp,
      JSON.stringify(results.settings),
      JSON.stringify(results.smallGroups)
    ]);
  }
  
  if (results.heatmaps) {
    sheet.appendRow([
      assessmentId,
      'heatmaps',
      '',
      timestamp,
      JSON.stringify(results.settings),
      JSON.stringify(results.heatmaps)
    ]);
  }
  
  if (results.districtSummary) {
    sheet.appendRow([
      assessmentId,
      'districtSummary',
      '',
      timestamp,
      JSON.stringify(results.settings),
      JSON.stringify(results.districtSummary)
    ]);
  }
  
  sheet.appendRow([
    assessmentId,
    'teacherSummaries',
    '',
    timestamp,
    JSON.stringify(results.settings),
    JSON.stringify(results.teacherSummaries)
  ]);
}

/**
 * Gets stored analysis results
 */
function getAnalysisResults(assessmentId) {
  try {
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return {}; // No results yet
    }
    
    const data = sheet.getDataRange().getValues();
    
    const results = {};
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === assessmentId) {
        const type = data[i][1];
        try {
          results[type] = JSON.parse(data[i][5] || '{}');
          results.timestamp = data[i][3];
          results.settings = JSON.parse(data[i][4] || '{}');
        } catch (parseError) {
          Logger.log('Error parsing results row ' + i + ': ' + parseError.message);
        }
      }
    }
    
    return results;
  } catch (error) {
    Logger.log('getAnalysisResults error: ' + error.message);
    return {};
  }
}

/**
 * Checks if re-run should overwrite or create new version
 */
function checkExistingAnalysis(assessmentId) {
  try {
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { exists: false };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === assessmentId) {
        return {
          exists: true,
          timestamp: data[i][3]
        };
      }
    }
    
    return { exists: false };
  } catch (error) {
    Logger.log('checkExistingAnalysis error: ' + error.message);
    return { exists: false };
  }
}

/**
 * Deletes existing analysis results for re-run
 */
function deleteExistingAnalysis(assessmentId) {
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);
  const data = sheet.getDataRange().getValues();
  
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === assessmentId) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // Delete from bottom to top to preserve row indices
  rowsToDelete.reverse().forEach(row => {
    sheet.deleteRow(row);
  });
}
