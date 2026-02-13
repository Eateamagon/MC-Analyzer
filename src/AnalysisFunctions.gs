/**
 * ═══════════════════════════════════════════════════════════════════════════
 *                          ANALYSIS FUNCTIONS
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Case-insensitive column index finder for CSV headers.
 * Handles variations in Mastery Connect CSV exports (e.g., "Teacher", "teacher", "TEACHER").
 */
function findColIndex(header, colName) {
  var lowerName = colName.toLowerCase();
  // First try exact case-insensitive match
  for (var i = 0; i < header.length; i++) {
    if (header[i] && header[i].toString().toLowerCase() === lowerName) {
      return i;
    }
  }
  // Fallback: try matching with underscores replaced by spaces and vice versa
  var altName = lowerName.indexOf('_') !== -1 ? lowerName.replace(/_/g, ' ') : lowerName.replace(/ /g, '_');
  for (var i = 0; i < header.length; i++) {
    if (header[i] && header[i].toString().toLowerCase() === altName) {
      return i;
    }
  }
  return -1;
}

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
  const { assessment, solRow, header, periodMap, solMetadata } = assessmentData;
  var students = assessmentData.students;
  const dataFormat = (assessment && assessment.data_format) || 'mastery_connect';

  // Get user settings
  const settings = user.settings || CONFIG.defaults;

  // Parse questions from header
  const questions = parseQuestions(header, solRow);

  // Get column indices (case-insensitive)
  const teacherCol = findColIndex(header, 'teacher');
  const firstNameCol = findColIndex(header, 'first_name');
  const lastNameCol = findColIndex(header, 'last_name');
  const studentIdCol = findColIndex(header, 'student_id');
  const pctCol = findColIndex(header, 'percentage');
  
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
  
  // Filter out students with no valid test data (all scores missing/empty)
  const validStudents = students.filter(function(r) {
    // Check if student has at least one answered question (score of 0 or 1)
    return questions.some(function(q) {
      var sc = r[q.scoreColIndex];
      return sc === 1 || sc === 0 || sc === '1' || sc === '0';
    });
  });

  Logger.log('Filtered students: ' + students.length + ' -> ' + validStudents.length + ' (removed ' + (students.length - validStudents.length) + ' with no test data)');
  students = validStudents;

  // Get unique teachers, then filter out teachers with zero students
  const teacherCounts = {};
  students.forEach(function(r) {
    var t = r[teacherCol];
    teacherCounts[t] = (teacherCounts[t] || 0) + 1;
  });
  const teachers = Object.keys(teacherCounts).filter(function(t) {
    return teacherCounts[t] > 0;
  }).sort();

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

  // SOL format-specific analysis
  results.dataFormat = dataFormat;
  if (dataFormat === 'sol_with_sol' && solMetadata) {
    // SOL With SOL: SOL Analysis, Item Analysis by description, Reporting Categories, Remediation
    if (options.solAnalysis !== false) {
      results.solAnalysis = generateSOLAnalysis(students, questions, solMetadata, teacherCol);
    }
    if (options.solItemAnalysis !== false) {
      results.solItemAnalysis = generateSOLItemAnalysis(students, questions, solMetadata);
    }
    if (options.reportingCategories !== false) {
      results.reportingCategories = generateReportingCategoryAnalysis(students, questions, solMetadata);
    }
    if (options.remediationGroups !== false) {
      results.remediationGroups = generateSmallGroups(
        students, questions, teacherCol, firstNameCol, lastNameCol,
        studentIdCol, settings, periodMap
      );
    }
  } else if (dataFormat === 'sol_without_sol' && solMetadata) {
    // SOL Without SOL: Scaled Scores, Item Difficulty Analysis
    if (options.scaledScores !== false) {
      results.scaledScores = generateScaledScoreAnalysis(
        students, questions, solMetadata, teacherCol, firstNameCol, lastNameCol, studentIdCol
      );
    }
    if (options.itemDifficulty !== false) {
      results.itemDifficulty = generateItemDifficultyAnalysis(
        students, questions, solMetadata, teacherCol, firstNameCol, lastNameCol
      );
    }
  }

  // Store results
  storeAnalysisResults(assessmentId, results);

  // Mark assessment as ready (in case it was still pending_periods)
  try {
    const ss = getOrCreateDatabase();
    const aSheet = ss.getSheetByName(CONFIG.sheets.assessments);
    if (aSheet) {
      const aData = aSheet.getDataRange().getValues();
      const aHeaders = aData[0];
      const idCol = aHeaders.indexOf('id');
      const statusCol = aHeaders.indexOf('status');
      const normalId = String(assessmentId).trim();
      for (let i = 1; i < aData.length; i++) {
        if (aData[i][idCol] != null && String(aData[i][idCol]).trim() === normalId) {
          if (aData[i][statusCol] !== 'ready') {
            aSheet.getRange(i + 1, statusCol + 1).setValue('ready');
          }
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('Could not update assessment status: ' + e.message);
  }
  
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
  const pctCol = findColIndex(header, 'percentage');
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
    
    // All distractors - filter out placeholder dashes (SOL format has no answer choices)
    const allDistractors = sortedWrong
      .map(([ans, count]) => ({ answer: ans, count: count }))
      .filter(d => d.answer !== '-' && d.answer !== '');

    // Update top wrong from filtered distractors
    if (allDistractors.length > 0) {
      topWrong = allDistractors[0].answer;
      topCount = allDistractors[0].count;
    } else if (sortedWrong.length > 0 && sortedWrong[0][0] === '-') {
      // SOL format: no meaningful distractor data
      topWrong = '-';
      topCount = 0;
    }

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

  // Students already filtered by runAnalysis, but guard against direct calls
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
  // Filter out teachers with zero students
  const summary = teacherSummaries
    .filter(ts => ts.studentCount > 0)
    .map(ts => ({
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

// ═══════════════════════════════════════════════════════════════════════════
//                   SOL FORMAT-SPECIFIC ANALYSIS FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * SOL Analysis: correct/incorrect breakdown per SOL code.
 * Used for "SOL With SOL" format.
 */
function generateSOLAnalysis(students, questions, solMetadata, teacherCol) {
  var sols = {};
  // Group questions by SOL code
  questions.forEach(function(q) {
    if (!sols[q.sol]) {
      sols[q.sol] = { sol: q.sol, questions: [], correct: 0, incorrect: 0, total: 0 };
    }
    sols[q.sol].questions.push(q);
  });

  // Tally correct/incorrect for each SOL across all students
  students.forEach(function(student) {
    Object.keys(sols).forEach(function(sol) {
      sols[sol].questions.forEach(function(q) {
        var sc = student[q.scoreColIndex];
        if (sc == 1 || sc === '1') { sols[sol].correct++; sols[sol].total++; }
        else if (sc == 0 || sc === '0') { sols[sol].incorrect++; sols[sol].total++; }
      });
    });
  });

  // Build result array sorted by SOL
  var result = Object.keys(sols).sort().map(function(sol) {
    var s = sols[sol];
    return {
      sol: sol,
      correct: s.correct,
      incorrect: s.incorrect,
      total: s.total,
      pctCorrect: s.total > 0 ? Math.round((s.correct / s.total) * 100) : 0,
      questionCount: s.questions.length
    };
  });

  return result;
}

/**
 * SOL Item Analysis: per-item breakdown using Item Description.
 * Used for "SOL With SOL" format.
 */
function generateSOLItemAnalysis(students, questions, solMetadata) {
  var descriptions = solMetadata.itemDescriptions || [];
  var difficulties = solMetadata.itemDifficulties || [];

  var items = questions.map(function(q, idx) {
    var correct = 0;
    var incorrect = 0;
    students.forEach(function(student) {
      var sc = student[q.scoreColIndex];
      if (sc == 1 || sc === '1') correct++;
      else if (sc == 0 || sc === '0') incorrect++;
    });
    var total = correct + incorrect;
    return {
      question: 'Q' + q.number,
      sol: q.sol,
      description: descriptions[idx] || '',
      difficulty: difficulties[idx] || '',
      correct: correct,
      incorrect: incorrect,
      total: total,
      pctCorrect: total > 0 ? Math.round((correct / total) * 100) : 0
    };
  });

  return items;
}

/**
 * Reporting Category Analysis: correct/incorrect/grand total per category.
 * Used for "SOL With SOL" format.
 */
function generateReportingCategoryAnalysis(students, questions, solMetadata) {
  var categories = {};
  var reportingCats = solMetadata.reportingCategories || [];

  // Group questions by reporting category
  questions.forEach(function(q, idx) {
    var cat = reportingCats[idx] || 'Unknown';
    if (!categories[cat]) {
      categories[cat] = { category: cat, questions: [], correct: 0, incorrect: 0, total: 0 };
    }
    categories[cat].questions.push(q);
  });

  // Tally correct/incorrect
  students.forEach(function(student) {
    Object.keys(categories).forEach(function(cat) {
      categories[cat].questions.forEach(function(q) {
        var sc = student[q.scoreColIndex];
        if (sc == 1 || sc === '1') { categories[cat].correct++; categories[cat].total++; }
        else if (sc == 0 || sc === '0') { categories[cat].incorrect++; categories[cat].total++; }
      });
    });
  });

  var result = Object.keys(categories).sort().map(function(cat) {
    var c = categories[cat];
    return {
      category: c.category,
      correct: c.correct,
      incorrect: c.incorrect,
      total: c.total,
      pctCorrect: c.total > 0 ? Math.round((c.correct / c.total) * 100) : 0,
      questionCount: c.questions.length
    };
  });

  // Add grand total row
  var grandCorrect = result.reduce(function(s, r) { return s + r.correct; }, 0);
  var grandIncorrect = result.reduce(function(s, r) { return s + r.incorrect; }, 0);
  var grandTotal = grandCorrect + grandIncorrect;

  return {
    categories: result,
    grandTotal: {
      correct: grandCorrect,
      incorrect: grandIncorrect,
      total: grandTotal,
      pctCorrect: grandTotal > 0 ? Math.round((grandCorrect / grandTotal) * 100) : 0
    }
  };
}

/**
 * Scaled Score per Student table.
 * Used for "SOL Without SOL" format.
 */
function generateScaledScoreAnalysis(students, questions, solMetadata, teacherCol, firstNameCol, lastNameCol, studentIdCol) {
  var scaledScores = solMetadata.scaledScores || {};

  var result = students.map(function(student) {
    var name = ((student[firstNameCol] || '') + ' ' + (student[lastNameCol] || '')).trim();
    var teacher = student[teacherCol] || '';
    var sid = student[studentIdCol] || '';

    // Calculate correct/total from the MC data
    var correct = 0;
    var total = 0;
    questions.forEach(function(q) {
      var sc = student[q.scoreColIndex];
      if (sc == 1 || sc === '1') { correct++; total++; }
      else if (sc == 0 || sc === '0') { total++; }
    });

    return {
      student: name,
      teacher: teacher,
      studentId: sid,
      scaledScore: scaledScores[sid] || '',
      correct: correct,
      total: total,
      pctCorrect: total > 0 ? Math.round((correct / total) * 100) : 0
    };
  });

  // Sort by teacher, then by name
  result.sort(function(a, b) {
    if (a.teacher !== b.teacher) return a.teacher.localeCompare(b.teacher);
    return a.student.localeCompare(b.student);
  });

  return result;
}

/**
 * Item Difficulty Analysis: per-student and total performance by difficulty (H/M/L).
 * Used for "SOL Without SOL" format.
 */
function generateItemDifficultyAnalysis(students, questions, solMetadata, teacherCol, firstNameCol, lastNameCol) {
  var difficulties = solMetadata.itemDifficulties || [];

  // Group questions by difficulty level
  var diffGroups = {};
  questions.forEach(function(q, idx) {
    var diff = (difficulties[idx] || 'Unknown').toUpperCase();
    if (!diffGroups[diff]) diffGroups[diff] = [];
    diffGroups[diff].push(q);
  });

  var diffLevels = Object.keys(diffGroups).sort();

  // Per-student breakdown
  var perStudent = students.map(function(student) {
    var name = ((student[firstNameCol] || '') + ' ' + (student[lastNameCol] || '')).trim();
    var teacher = student[teacherCol] || '';

    var byDiff = {};
    diffLevels.forEach(function(diff) {
      var correct = 0;
      var total = 0;
      diffGroups[diff].forEach(function(q) {
        var sc = student[q.scoreColIndex];
        if (sc == 1 || sc === '1') { correct++; total++; }
        else if (sc == 0 || sc === '0') { total++; }
      });
      byDiff[diff] = {
        correct: correct,
        total: total,
        pctCorrect: total > 0 ? Math.round((correct / total) * 100) : 0
      };
    });

    return {
      student: name,
      teacher: teacher,
      byDifficulty: byDiff
    };
  });

  // Sort by teacher then name
  perStudent.sort(function(a, b) {
    if (a.teacher !== b.teacher) return a.teacher.localeCompare(b.teacher);
    return a.student.localeCompare(b.student);
  });

  // Total breakdown by difficulty
  var totals = {};
  diffLevels.forEach(function(diff) {
    var correct = 0;
    var total = 0;
    students.forEach(function(student) {
      diffGroups[diff].forEach(function(q) {
        var sc = student[q.scoreColIndex];
        if (sc == 1 || sc === '1') { correct++; total++; }
        else if (sc == 0 || sc === '0') { total++; }
      });
    });
    totals[diff] = {
      correct: correct,
      total: total,
      pctCorrect: total > 0 ? Math.round((correct / total) * 100) : 0,
      questionCount: diffGroups[diff].length
    };
  });

  return {
    difficultyLevels: diffLevels,
    perStudent: perStudent,
    totals: totals,
    questionCountByDifficulty: diffLevels.map(function(d) { return { difficulty: d, count: diffGroups[d].length }; })
  };
}

/**
 * Stores analysis results in the database
 */
function storeAnalysisResults(assessmentId, results) {
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);
  const timestamp = new Date().toISOString();

  // Consolidate all analysis types into a single JSON object in one row
  const allData = {
    itemAnalysis: results.itemAnalysis || null,
    smallGroups: results.smallGroups || null,
    heatmaps: results.heatmaps || null,
    districtSummary: results.districtSummary || null,
    teacherSummaries: results.teacherSummaries || [],
    // SOL format-specific results
    dataFormat: results.dataFormat || null,
    solAnalysis: results.solAnalysis || null,
    solItemAnalysis: results.solItemAnalysis || null,
    reportingCategories: results.reportingCategories || null,
    remediationGroups: results.remediationGroups || null,
    scaledScores: results.scaledScores || null,
    itemDifficulty: results.itemDifficulty || null
  };

  // Single row write instead of 5 separate appendRow calls
  sheet.appendRow([
    assessmentId,
    'all',        // consolidated type marker
    '',
    timestamp,
    JSON.stringify(results.settings || {}),
    JSON.stringify(allData)
  ]);
}

/**
 * Gets stored analysis results
 */
function getAnalysisResults(assessmentId) {
  try {
    const normalId = assessmentId.toString().trim();
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);

    if (!sheet || sheet.getLastRow() < 2) {
      return {}; // No results yet
    }

    const data = sheet.getDataRange().getValues();
    const results = {};

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === normalId) {
        const type = data[i][1];
        try {
          if (type === 'all') {
            // New consolidated format: single row with all analysis types
            const allData = JSON.parse(data[i][5] || '{}');
            if (allData.itemAnalysis) results.itemAnalysis = allData.itemAnalysis;
            if (allData.smallGroups) results.smallGroups = allData.smallGroups;
            if (allData.heatmaps) results.heatmaps = allData.heatmaps;
            if (allData.districtSummary) results.districtSummary = allData.districtSummary;
            if (allData.teacherSummaries) results.teacherSummaries = allData.teacherSummaries;
            // SOL format-specific results
            if (allData.dataFormat) results.dataFormat = allData.dataFormat;
            if (allData.solAnalysis) results.solAnalysis = allData.solAnalysis;
            if (allData.solItemAnalysis) results.solItemAnalysis = allData.solItemAnalysis;
            if (allData.reportingCategories) results.reportingCategories = allData.reportingCategories;
            if (allData.remediationGroups) results.remediationGroups = allData.remediationGroups;
            if (allData.scaledScores) results.scaledScores = allData.scaledScores;
            if (allData.itemDifficulty) results.itemDifficulty = allData.itemDifficulty;
          } else {
            // Legacy format: one row per analysis type
            results[type] = JSON.parse(data[i][5] || '{}');
          }
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
    const normalId = assessmentId.toString().trim();
    const ss = getOrCreateDatabase();
    const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);

    if (!sheet || sheet.getLastRow() < 2) {
      return { exists: false };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === normalId) {
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
  const normalId = assessmentId.toString().trim();
  const ss = getOrCreateDatabase();
  const sheet = ss.getSheetByName(CONFIG.sheets.analysisResults);
  const data = sheet.getDataRange().getValues();

  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === normalId) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // Delete from bottom to top to preserve row indices
  rowsToDelete.reverse().forEach(row => {
    sheet.deleteRow(row);
  });
}

/**
 * Generates small groups from multiple assessment data sources.
 * Merges per-student SOL profiles across assessments, then runs the
 * group-formation algorithm.  Students are matched by student_id.
 *
 * @param {string[]} assessmentIds  - Array of assessment IDs to pull data from
 * @return {Object} { groups: [...], sources: [...] }
 */
function generateMultiSourceSmallGroups(assessmentIds) {
  var user = getCurrentUser();
  if (!user || user.isNewUser) throw new Error('Not authenticated');
  var settings = user.settings || CONFIG.defaults;

  // Collect per-student profiles across all assessments.
  // Key = student_id, value = { name, teacher, sols: { solName: {c, t} }, periods: {} }
  var studentProfiles = {};
  var sourceNames = [];

  assessmentIds.forEach(function(assessmentId) {
    var data = getAssessmentData(assessmentId);
    sourceNames.push(data.assessment.name || assessmentId);

    var header = data.header;
    var solRow = data.solRow;
    var students = data.students;
    var periodMap = data.periodMap || {};
    var questions = parseQuestions(header, solRow);
    var teacherCol = findColIndex(header, 'teacher');
    var firstNameCol = findColIndex(header, 'first_name');
    var lastNameCol = findColIndex(header, 'last_name');
    var studentIdCol = findColIndex(header, 'student_id');

    // Filter out students with no valid test data
    var validStudents = students.filter(function(r) {
      return questions.some(function(q) {
        var sc = r[q.scoreColIndex];
        return sc === 1 || sc === 0 || sc === '1' || sc === '0';
      });
    });

    validStudents.forEach(function(r) {
      var sid = r[studentIdCol];
      if (!sid) return;
      var sKey = String(sid).trim();
      var name = ((r[firstNameCol] || '') + ' ' + (r[lastNameCol] || '')).trim();
      var teacher = r[teacherCol] || 'Unknown';
      var period = (periodMap[sKey]) ? periodMap[sKey] : 'Unknown';

      if (!studentProfiles[sKey]) {
        studentProfiles[sKey] = {
          id: sKey,
          name: name,
          teacher: teacher,
          period: period,
          sols: {}
        };
      }
      // Keep most recent teacher/name (last assessment wins for metadata)
      studentProfiles[sKey].name = name;
      studentProfiles[sKey].teacher = teacher;
      if (period !== 'Unknown') studentProfiles[sKey].period = period;

      // Accumulate SOL scores across assessments
      questions.forEach(function(q) {
        var sc = r[q.scoreColIndex];
        if (!studentProfiles[sKey].sols[q.sol]) {
          studentProfiles[sKey].sols[q.sol] = { c: 0, t: 0 };
        }
        if (sc == 1 || sc == '1') {
          studentProfiles[sKey].sols[q.sol].c++;
          studentProfiles[sKey].sols[q.sol].t++;
        } else if (sc == 0 || sc == '0') {
          studentProfiles[sKey].sols[q.sol].t++;
        }
      });
    });
  });

  // Convert profiles to the format the grouping algorithm expects,
  // grouped by teacher then period.
  var byTeacher = {};
  Object.keys(studentProfiles).forEach(function(sKey) {
    var p = studentProfiles[sKey];
    if (!byTeacher[p.teacher]) byTeacher[p.teacher] = {};
    if (!byTeacher[p.teacher][p.period]) byTeacher[p.teacher][p.period] = [];

    // Determine weak SOLs (below threshold across combined data)
    var weak = [];
    Object.keys(p.sols).forEach(function(sol) {
      var s = p.sols[sol];
      if (s.t > 0 && ((s.c / s.t) * 100) < (settings.sgThreshold || 70)) {
        weak.push(sol);
      }
    });

    if (weak.length > 0) {
      byTeacher[p.teacher][p.period].push({
        name: p.name,
        sols: weak,
        id: p.id
      });
    }
  });

  // Run group formation per teacher per period
  var groups = [];
  Object.keys(byTeacher).sort().forEach(function(teacher) {
    Object.keys(byTeacher[teacher]).sort().forEach(function(period) {
      var profiles = byTeacher[teacher][period];
      var periodGroups = formGroups(profiles, settings);
      periodGroups.forEach(function(g, idx) {
        groups.push({
          teacher: teacher,
          period: period,
          groupNumber: idx + 1,
          students: g.students.map(function(s) { return s.name; }),
          sharedWeakSOLs: g.shared,
          studentCount: g.students.length,
          isUnknownPeriod: period === 'Unknown'
        });
      });
    });
  });

  return {
    groups: groups,
    sources: sourceNames,
    studentCount: Object.keys(studentProfiles).length
  };
}

/**
 * Compares two assessments and returns a comparison report.
 * Each assessment must already be uploaded. Runs analysis on both
 * (if not already cached) and produces per-teacher & per-SOL deltas.
 */
function compareAssessments(assessmentIdA, assessmentIdB) {
  var user = getCurrentUser();
  if (!user || user.isNewUser) throw new Error('User not authenticated');

  // Load both datasets
  var dataA = getAssessmentData(assessmentIdA);
  var dataB = getAssessmentData(assessmentIdB);
  var settings = user.settings || CONFIG.defaults;

  var questionsA = parseQuestions(dataA.header, dataA.solRow);
  var questionsB = parseQuestions(dataB.header, dataB.solRow);

  // Filter out students with no valid test data
  function hasValidScores(row, questions) {
    return questions.some(function(q) {
      var sc = row[q.scoreColIndex];
      return sc === 1 || sc === 0 || sc === '1' || sc === '0';
    });
  }
  dataA.students = dataA.students.filter(function(r) { return hasValidScores(r, questionsA); });
  dataB.students = dataB.students.filter(function(r) { return hasValidScores(r, questionsB); });

  var teacherColA = findColIndex(dataA.header, 'teacher');
  var teacherColB = findColIndex(dataB.header, 'teacher');
  var pctColA = findColIndex(dataA.header, 'percentage');
  var pctColB = findColIndex(dataB.header, 'percentage');
  var studentIdColA = findColIndex(dataA.header, 'student_id');
  var studentIdColB = findColIndex(dataB.header, 'student_id');
  var firstNameColA = findColIndex(dataA.header, 'first_name');
  var lastNameColA = findColIndex(dataA.header, 'last_name');
  var firstNameColB = findColIndex(dataB.header, 'first_name');
  var lastNameColB = findColIndex(dataB.header, 'last_name');

  // ── Per-teacher SOL comparison ──
  function buildTeacherSOLStats(students, questions, teacherCol) {
    var stats = {};
    students.forEach(function(r) {
      var t = r[teacherCol] || 'Unknown';
      if (!stats[t]) stats[t] = {};
      questions.forEach(function(q) {
        if (!stats[t][q.sol]) stats[t][q.sol] = { c: 0, t: 0 };
        var sc = r[q.scoreColIndex];
        if (sc == 1 || sc == '1') { stats[t][q.sol].c++; stats[t][q.sol].t++; }
        else if (sc == 0 || sc == '0') { stats[t][q.sol].t++; }
      });
    });
    // convert to percentages
    var result = {};
    Object.keys(stats).forEach(function(t) {
      result[t] = {};
      Object.keys(stats[t]).forEach(function(sol) {
        var s = stats[t][sol];
        result[t][sol] = s.t > 0 ? Math.round((s.c / s.t) * 100) : null;
      });
    });
    return result;
  }

  var solStatsA = buildTeacherSOLStats(dataA.students, questionsA, teacherColA);
  var solStatsB = buildTeacherSOLStats(dataB.students, questionsB, teacherColB);

  // All teachers across both assessments
  var allTeachers = Object.keys(solStatsA);
  Object.keys(solStatsB).forEach(function(t) {
    if (allTeachers.indexOf(t) === -1) allTeachers.push(t);
  });
  allTeachers.sort();

  // All SOLs across both
  var allSOLs = [];
  [solStatsA, solStatsB].forEach(function(ss) {
    Object.keys(ss).forEach(function(t) {
      Object.keys(ss[t]).forEach(function(sol) {
        if (allSOLs.indexOf(sol) === -1) allSOLs.push(sol);
      });
    });
  });
  allSOLs.sort();

  // Build teacher comparison rows
  var teacherComparisons = allTeachers.map(function(teacher) {
    var solDeltas = allSOLs.map(function(sol) {
      var pctA = (solStatsA[teacher] && solStatsA[teacher][sol] != null) ? solStatsA[teacher][sol] : null;
      var pctB = (solStatsB[teacher] && solStatsB[teacher][sol] != null) ? solStatsB[teacher][sol] : null;
      var delta = (pctA !== null && pctB !== null) ? pctB - pctA : null;
      return { sol: sol, pctA: pctA, pctB: pctB, delta: delta };
    });

    // Overall averages
    var countA = 0, sumA = 0, countB = 0, sumB = 0;
    dataA.students.forEach(function(r) {
      if (r[teacherColA] === teacher && r[pctColA] != null && r[pctColA] !== '') {
        sumA += parseFloat(r[pctColA]); countA++;
      }
    });
    dataB.students.forEach(function(r) {
      if (r[teacherColB] === teacher && r[pctColB] != null && r[pctColB] !== '') {
        sumB += parseFloat(r[pctColB]); countB++;
      }
    });
    var avgA = countA > 0 ? Math.round(sumA / countA) : null;
    var avgB = countB > 0 ? Math.round(sumB / countB) : null;
    var avgDelta = (avgA !== null && avgB !== null) ? avgB - avgA : null;

    return {
      teacher: teacher,
      studentsA: countA,
      studentsB: countB,
      avgA: avgA,
      avgB: avgB,
      avgDelta: avgDelta,
      solDeltas: solDeltas
    };
  });

  // ── Per-student comparison (matched by student_id) ──
  var studentMapA = {};
  dataA.students.forEach(function(r) {
    var sid = r[studentIdColA];
    if (sid) {
      var name = ((r[firstNameColA] || '') + ' ' + (r[lastNameColA] || '')).trim();
      var teacher = r[teacherColA] || '';
      var pct = parseFloat(r[pctColA]) || 0;
      studentMapA[sid] = { name: name, teacher: teacher, pct: pct };
    }
  });
  var studentMapB = {};
  dataB.students.forEach(function(r) {
    var sid = r[studentIdColB];
    if (sid) {
      var name = ((r[firstNameColB] || '') + ' ' + (r[lastNameColB] || '')).trim();
      var teacher = r[teacherColB] || '';
      var pct = parseFloat(r[pctColB]) || 0;
      studentMapB[sid] = { name: name, teacher: teacher, pct: pct };
    }
  });

  var studentComparisons = [];
  // Students present in both assessments
  Object.keys(studentMapA).forEach(function(sid) {
    if (studentMapB[sid]) {
      studentComparisons.push({
        studentId: sid,
        name: studentMapA[sid].name || studentMapB[sid].name,
        teacher: studentMapB[sid].teacher || studentMapA[sid].teacher,
        pctA: studentMapA[sid].pct,
        pctB: studentMapB[sid].pct,
        delta: Math.round(studentMapB[sid].pct - studentMapA[sid].pct)
      });
    }
  });
  studentComparisons.sort(function(a, b) { return a.delta - b.delta; });

  // ── Overall summary ──
  var totalA = dataA.students.length;
  var totalB = dataB.students.length;
  var overallSumA = 0, overallCountA = 0;
  dataA.students.forEach(function(r) {
    var v = parseFloat(r[pctColA]);
    if (!isNaN(v)) { overallSumA += v; overallCountA++; }
  });
  var overallSumB = 0, overallCountB = 0;
  dataB.students.forEach(function(r) {
    var v = parseFloat(r[pctColB]);
    if (!isNaN(v)) { overallSumB += v; overallCountB++; }
  });
  var overallAvgA = overallCountA > 0 ? Math.round(overallSumA / overallCountA) : 0;
  var overallAvgB = overallCountB > 0 ? Math.round(overallSumB / overallCountB) : 0;

  return {
    assessmentA: { id: assessmentIdA, name: dataA.assessment.name, studentCount: totalA, average: overallAvgA },
    assessmentB: { id: assessmentIdB, name: dataB.assessment.name, studentCount: totalB, average: overallAvgB },
    overallDelta: overallAvgB - overallAvgA,
    sols: allSOLs,
    teacherComparisons: teacherComparisons,
    studentComparisons: studentComparisons,
    matchedStudentCount: studentComparisons.length
  };
}
