/**
 * ═══════════════════════════════════════════════════════════════
 *  UNIFIED CENTRAL API — Google Apps Script Web App (v3.0)
 *  Supports both Smart Attendance PWA and Academic File PWA.
 *  Proxies access to college sheets via the sheetId parameter.
 * ═══════════════════════════════════════════════════════════════
 */

var _ssCache = {};
function _getSpreadsheet(sheetId) {
  if (!sheetId) {
    throw new Error("Missing sheetId parameter");
  }
  if (!_ssCache[sheetId]) {
    _ssCache[sheetId] = SpreadsheetApp.openById(sheetId);
  }
  return _ssCache[sheetId];
}

/**
 * Main GET entry point - merges routes for Attendance and Academic PWAs
 */
function doGet(e) {
  try {
    var action = e.parameter.action;
    var sheetId = e.parameter.sheetId; // Master config sheet ID
    var result;

    switch (action) {
      // ── Attendance & Common Routes ──
      case 'getTeachers': 
        result = getTeachers(sheetId); 
        break;
      case 'getSubjects': 
        result = getSubjects(e.parameter.teacher, sheetId); 
        break;
      case 'getStudents': 
        result = getStudents(e.parameter.sheet, e.parameter.batch, sheetId); 
        break;
      case 'getAttendanceLimit': 
        result = getAttendanceLimit(sheetId); 
        break;
      case 'getAttendance': 
        result = getAttendance(e.parameter.code, e.parameter.year, e.parameter.date, e.parameter.outputSheetId, sheetId); 
        break;
      case 'getSyllabus':
        result = getSyllabus(e.parameter.link, e.parameter.code, sheetId);
        break;
      case 'getConfig':
      case 'getAllData': 
        result = getAllData(sheetId); 
        break;

      // ── Academic File Routes ──
      case 'getTeachingPlan':
        result = getTeachingPlan(e.parameter.code, e.parameter.teacher, sheetId);
        break;
      case 'syncTeachingPlan':
        result = syncTeachingPlan(e.parameter.code, e.parameter.teacher, sheetId);
        break;
      case 'getAcademicSchedule':
        result = getAcademicSchedule(sheetId);
        break;

      default: 
        result = { error: 'Unknown GET action: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Main POST entry point - merges post routes for Attendance and Academic PWAs
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || (e.parameter && e.parameter.action);
    var sheetId = data.sheetId || (e.parameter && e.parameter.sheetId);
    var result;

    switch (action) {
      // ── Attendance POSTs ──
      case 'saveAttendance': 
        result = saveAttendance(data.records, data.outputSheetId, data.collegeName, data.managementName, sheetId); 
        break;

      // ── Academic File POSTs ──
      case 'saveRemark':
        result = saveRemark(data.code, data.rowIndex, data.remark, sheetId);
        break;
      case 'addCustomSyllabusTopic':
        result = addCustomSyllabusTopic(data, sheetId);
        break;

      default:
        result = { error: 'Unknown POST action: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ═══════════════════════════════════════════════════════════════
   COMMON / UTILS FUNCTIONS
   ═══════════════════════════════════════════════════════════════ */

/**
 * Map the columns of the "subjects" tab by scanning header names.
 * Falls back to the classic fixed positions (A=code ... H=pin) only
 * when a header cannot be matched, so old sheets keep working.
 */
function _mapSubjectCols(headers) {
  var H = headers.map(function(h) { return String(h).toLowerCase().trim(); });
  var used = {};
  function find(keywords, fallback) {
    for (var k = 0; k < keywords.length; k++) {
      for (var c = 0; c < H.length; c++) {
        if (!used[c] && H[c] === keywords[k]) { used[c] = true; return c; }
      }
    }
    for (var k = 0; k < keywords.length; k++) {
      for (var c = 0; c < H.length; c++) {
        if (!used[c] && H[c] && H[c].indexOf(keywords[k]) !== -1) { used[c] = true; return c; }
      }
    }
    used[fallback] = true;
    return fallback;
  }
  // Specific labels first so partial matches can't steal their columns
  // (e.g. "faculty name" must resolve to faculty before "name" is searched).
  return {
    code: find(['subject code', 'code'], 0),
    faculty: find(['faculty', 'teacher'], 6),
    pin: find(['pin', 'password'], 7),
    semester: find(['semester', 'sem'], 4),
    year: find(['year', 'class'], 2),
    program: find(['program', 'course'], 3),
    type: find(['type'], 5),
    name: find(['subject name', 'subject', 'name'], 1)
  };
}

/**
 * Smart Subject Code Parser
 * Parses strings like "BP702P (A)", "BP702P(A)", "BP 702T", "BP701T (IMA)", "BP702P"
 */
function _parseSubjectCode(code, typeHint, nameHint) {
  var raw = String(code || '').trim();
  if (!raw) {
    return {
      raw: '',
      baseCode: '',
      cleanBaseCode: '',
      cleanFullCode: '',
      batch: '',
      isPractical: false
    };
  }

  // Extract batch inside brackets if present, e.g. "BP702P (A)" -> batch = "A"
  var batch = '';
  var bracketMatch = raw.match(/\(([^)]+)\)/);
  if (bracketMatch && bracketMatch[1]) {
    batch = bracketMatch[1].trim();
  }

  // Base code stripped of parenthetical text and extra spaces
  var baseCode = raw.replace(/\s*\([^)]*\)/g, '').trim();
  // Alphanumeric clean base code without spaces or dashes, e.g. "BP 702P" -> "BP702P"
  var cleanBaseCode = baseCode.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  var cleanFullCode = raw.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

  // Practical determination
  var typeStr = String(typeHint || '').toLowerCase();
  var nameStr = String(nameHint || '').toLowerCase();
  var codeUpper = cleanBaseCode;

  var isPractical = false;
  if (typeStr.indexOf('practical') !== -1 || typeStr.indexOf('lab') !== -1 || typeStr === 'pr' || typeStr === 'p') {
    isPractical = true;
  } else if (nameStr.indexOf('practical') !== -1 || nameStr.indexOf('lab') !== -1) {
    isPractical = true;
  } else if (raw.toLowerCase().indexOf('practical') !== -1 || raw.toLowerCase().indexOf('lab') !== -1) {
    isPractical = true;
  } else {
    // Check code ending with P (e.g. BP702P, BP106P, etc.)
    if (/.*?\d+P$/i.test(codeUpper) || codeUpper.endsWith('P')) {
      isPractical = true;
    }
  }

  return {
    raw: raw,
    baseCode: baseCode,
    cleanBaseCode: cleanBaseCode,
    cleanFullCode: cleanFullCode,
    batch: batch,
    isPractical: isPractical
  };
}

/**
 * Fail-Proof Sheet Search Algorithm
 * Finds the best sheet tab matching subject code (handles "BP702P (A)", "BP 702P", "BP702P", etc.)
 */
function _findSheetByCode(ss, inputCode) {
  if (!ss || !inputCode) return null;

  var parsedInput = _parseSubjectCode(inputCode);
  var sheets = ss.getSheets();
  if (!sheets || sheets.length === 0) return null;

  var bestSheet = null;
  var maxScore = -1;

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName().trim();
    var parsedSheet = _parseSubjectCode(sheetName);
    var score = 0;

    // 1. Exact string match (case-insensitive)
    if (sheetName.toLowerCase() === parsedInput.raw.toLowerCase()) {
      score = 100;
    }
    // 2. Clean full code match (e.g., "BP 702P (A)" == "BP702P(A)")
    else if (parsedSheet.cleanFullCode === parsedInput.cleanFullCode) {
      score = 90;
    }
    // 3. Exact clean base code match with matching batch
    else if (parsedSheet.cleanBaseCode === parsedInput.cleanBaseCode && parsedInput.batch && parsedSheet.batch && parsedInput.batch.toLowerCase() === parsedSheet.batch.toLowerCase()) {
      score = 85;
    }
    // 4. Exact clean base code match (e.g. input "BP702P (A)" matches sheet "BP702P" or "BP 702P")
    else if (parsedSheet.cleanBaseCode === parsedInput.cleanBaseCode) {
      score = 80;
    }
    // 5. Sheet name contains clean base code and batch keyword
    else if (parsedSheet.cleanBaseCode.indexOf(parsedInput.cleanBaseCode) !== -1 || parsedInput.cleanBaseCode.indexOf(parsedSheet.cleanBaseCode) !== -1) {
      if (parsedInput.batch && sheetName.toLowerCase().indexOf(parsedInput.batch.toLowerCase()) !== -1) {
        score = 75;
      } else {
        score = 70;
      }
    }
    // 6. Sheet name starts with clean base code
    else if (sheetName.toUpperCase().replace(/[^A-Z0-9]/g, '').indexOf(parsedInput.cleanBaseCode) === 0) {
      score = 60;
    }

    if (score > maxScore) {
      maxScore = score;
      bestSheet = sheet;
    }
  }

  // If score is strong enough, return best match
  if (bestSheet && maxScore >= 60) {
    return bestSheet;
  }

  // Fallback: Check Priority 2 keyword match if sheet has "syllabus" or "teaching plan" or "plan"
  for (var i = 0; i < sheets.length; i++) {
    var nameLower = sheets[i].getName().trim().toLowerCase();
    if (looksLikeSubjectCode(nameLower) && _parseSubjectCode(nameLower).cleanBaseCode !== parsedInput.cleanBaseCode) {
      continue; // Skip different subject code sheet
    }
    if (nameLower.indexOf("syllabus") !== -1 || nameLower.indexOf("teaching plan") !== -1 || nameLower.indexOf("plan") !== -1) {
      return sheets[i];
    }
  }

  // Fallback: First sheet if it doesn't look like a completely different subject code
  if (sheets[0]) {
    var firstName = sheets[0].getName().trim();
    if (looksLikeSubjectCode(firstName) && _parseSubjectCode(firstName).cleanBaseCode !== parsedInput.cleanBaseCode) {
      return null;
    }
    return sheets[0];
  }

  return null;
}

function getTeachers(sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects');
  if (!ws) return { success: false, error: 'Sheet "subjects" not found' };
  var data = ws.getDataRange().getValues(), map = {};
  var cols = _mapSubjectCols(data[0] || []);
  for (var i = 1; i < data.length; i++) {
    var fStr = String(data[i][cols.faculty]).trim(), pStr = String(data[i][cols.pin]).trim();
    if (fStr && fStr !== 'undefined') {
      var fs = fStr.split(','), ps = pStr.split(',');
      for (var f = 0; f < fs.length; f++) {
        var n = fs[f].trim(), p = (ps[f] && ps[f].trim()) || ps[0].trim();
        if (n) {
          if (!map[n]) map[n] = p;
          else if (map[n].split(',').indexOf(p) === -1) map[n] += ',' + p;
        }
      }
    }
  }
  var res = []; for (var k in map) res.push({ name: k, pin: map[k] });
  return { success: true, teachers: res };
}

function getSubjects(teacher, sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects');
  if (!ws) return { success: false };
  var data = ws.getDataRange().getValues(), res = [];
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var cols = _mapSubjectCols(data[0] || []);

  var teachingPlanIdx = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h.indexOf('teaching plan') !== -1 || h.indexOf('syllabus') !== -1) {
      teachingPlanIdx = c;
      break;
    }
  }

  for (var i = 1; i < data.length; i++) {
    var fs = String(data[i][cols.faculty]).toLowerCase().split(',').map(function(x){return x.trim()});
    if (fs.indexOf(teacher.toLowerCase()) !== -1) {
      var sCode = String(data[i][cols.code]).trim();
      var sName = String(data[i][cols.name]).trim();
      var sType = String(data[i][cols.type]).trim();
      var parsedCode = _parseSubjectCode(sCode, sType, sName);
      if (parsedCode.isPractical && (!sType || sType.toLowerCase() === 'theory' || sType === '')) {
        sType = 'Practical';
      }
      var subObj = { code: sCode, name: sName, year: String(data[i][cols.year]).trim(), program: String(data[i][cols.program]).trim(), semester: String(data[i][cols.semester]).trim(), type: sType };
      subObj.teachingPlanLink = (teachingPlanIdx !== -1) ? String(data[i][teachingPlanIdx]).trim() : '';
      res.push(subObj);
    }
  }
  // Fallback: fill empty teachingPlanLink from master config sheet
  var globalLink = '';
  for (var i = 0; i < res.length; i++) {
    if (!res[i].teachingPlanLink) {
      if (!globalLink) globalLink = getGlobalTeachingPlanLink(sheetId);
      if (globalLink) res[i].teachingPlanLink = globalLink;
    }
  }
  return { success: true, subjects: res };
}

function getStudents(sheet, batch, sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName(sheet);
  if (!ws) return { success: false };
  var data = ws.getDataRange().getValues(), res = [];
  // Locate columns by header name; fall back to A=roll, B=name, C=batch
  var H = (data[0] || []).map(function(h) { return String(h).toLowerCase().trim(); });
  var rollCol = 0, nameCol = 1, batchCol = 2;
  for (var c = 0; c < H.length; c++) {
    if (H[c].indexOf('roll') !== -1) rollCol = c;
    else if (H[c].indexOf('name') !== -1) nameCol = c;
    else if (H[c].indexOf('batch') !== -1) batchCol = c;
  }
  for (var i = 1; i < data.length; i++) {
    var r = data[i][rollCol], n = String(data[i][nameCol]).trim(), b = String(data[i][batchCol] || '').trim();
    if (!r && !n) continue;
    if (batch && b !== batch) continue;
    res.push({ rollNo: r, name: n, batch: b });
  }
  return { success: true, students: res, sheet: sheet };
}

function getAttendanceLimit(sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects');
  var data = ws ? ws.getDataRange().getValues() : [], limit = 75;
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (String(data[i][j]).toLowerCase().indexOf('attendance limit') !== -1 && j + 1 < data[i].length) {
        var v = Number(data[i][j + 1]); if (!isNaN(v) && v > 0) limit = v; break;
      }
    }
  }
  return { success: true, limit: limit };
}

function getAllData(sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects'), subs = [], config = { collegeName: '', managementName: '' };
  if (ws) {
    var data = ws.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var cols = _mapSubjectCols(data[0] || []);

    var teachingPlanIdx = -1;
    for (var c = 0; c < headers.length; c++) {
      var h = headers[c];
      if (h.indexOf('teaching plan') !== -1 || h.indexOf('syllabus') !== -1) {
        teachingPlanIdx = c;
        break;
      }
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cols.code]).trim()) {
        var sCode = String(data[i][cols.code]).trim();
        var sName = String(data[i][cols.name]).trim();
        var sType = String(data[i][cols.type]).trim();
        var parsedCode = _parseSubjectCode(sCode, sType, sName);
        if (parsedCode.isPractical && (!sType || sType.toLowerCase() === 'theory' || sType === '')) {
          sType = 'Practical';
        }
        var subObj = {
          code: sCode,
          name: sName,
          year: String(data[i][cols.year]).trim(),
          program: String(data[i][cols.program]).trim(),
          semester: String(data[i][cols.semester]).trim(),
          type: sType,
          faculty: String(data[i][cols.faculty]).trim()
        };
        subObj.teachingPlanLink = (teachingPlanIdx !== -1) ? String(data[i][teachingPlanIdx]).trim() : '';
        subs.push(subObj);
      }
    }
    // Fallback: fill empty teachingPlanLink from master config sheet
    var globalLink = '';
    for (var i = 0; i < subs.length; i++) {
      if (!subs[i].teachingPlanLink) {
        if (!globalLink) globalLink = getGlobalTeachingPlanLink(sheetId);
        if (globalLink) subs[i].teachingPlanLink = globalLink;
      }
    }
    var cs = ss.getSheetByName('client sheet') || ss.getSheetByName('subjects');
    if (cs) {
      var cd = cs.getDataRange().getValues(), keys = ['college name', 'management name'];
      for (var r = 0; r < cd.length; r++) {
        for (var c = 0; c < cd[r].length; c++) {
          var v = String(cd[r][c]).trim().toLowerCase();
          for (var k = 0; k < keys.length; k++) {
            if (v.indexOf(keys[k]) !== -1) {
              var f = '';
              for (var n = c + 1; n < cd[r].length; n++) { var nv = String(cd[r][n]).trim(); if (nv !== '' && ['link','name','text'].indexOf(nv.toLowerCase()) === -1) { f = nv; break; } }
              if (f === '' && r + 1 < cd.length) f = String(cd[r+1][c]).trim();
              if (f) { if (keys[k] === 'college name') config.collegeName = f; else config.managementName = f; }
            }
          }
        }
      }
    }
  }
  return { success: !!ws, teachers: getTeachers(sheetId).teachers || [], subjects: subs, attendanceLimit: getAttendanceLimit(sheetId).limit || 75, config: config };
}

function getOutputSheetId(sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects');
  var data = ws ? ws.getDataRange().getValues() : [];
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      var cellVal = String(data[i][j]).trim().toLowerCase();
      if (cellVal === 'output excel link' || cellVal.indexOf('output sheet') !== -1 || cellVal.indexOf('output excel') !== -1 || cellVal.indexOf('output link') !== -1) {
         var f = '';
         for (var n = j + 1; n < data[i].length; n++) { var nv = String(data[i][n]).trim(); if (nv !== '' && ['link','name','text'].indexOf(nv.toLowerCase()) === -1) { f = nv; break; } }
         if (f === '' && i + 1 < data.length) f = String(data[i+1][j]).trim();
         if (f) { var m = f.match(/\/d\/(.*?)(\/|$)/); if (m && m[1]) return m[1]; }
      }
    }
  }
  return '';
}

function getTargetSheetIds(code, sheetId) {
  var teachingPlanId = '';
  var outputSheetId = '';

  // 1. Try to resolve links directly from the shared Master Config Sheet row
  try {
    var MASTER_CONFIG_SHEET_ID = "1p3WoC2s-YYqn9ekqkQ72banxAAd-ujlDoFYpv4fkXmk";
    var masterSs = SpreadsheetApp.openById(MASTER_CONFIG_SHEET_ID);
    var masterWs = masterSs.getSheetByName("smart attendance client sheet") || masterSs.getSheets()[0];
    if (masterWs) {
      var data = masterWs.getDataRange().getValues();
      // Look at the headers in row 3 (index 2) to find output link and teaching plan link
      var headers = data[2] || data[0]; 
      var inputCol = -1, outputCol = -1, tpCol = -1;
      
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]).toLowerCase().trim();
        if (h.indexOf('input sheet id') !== -1 || h.indexOf('input link') !== -1) inputCol = c;
        if (h.indexOf('output link') !== -1 || h.indexOf('output sheet') !== -1 || h.indexOf('output excel') !== -1) outputCol = c;
        if (h.indexOf('teaching plan link') !== -1 || h.indexOf('teaching plan') !== -1 || h.indexOf('syllabus') !== -1) tpCol = c;
      }
      
      // Default fallback column indexes if header strings didn't hit
      if (inputCol === -1) inputCol = 4; // Col E
      if (outputCol === -1) outputCol = 5; // Col F
      if (tpCol === -1) tpCol = 6; // Col G
      
      for (var r = 3; r < data.length; r++) {
        var row = data[r];
        var rowInputId = String(row[inputCol] || '').trim();
        if (rowInputId === sheetId || (sheetId && rowInputId.indexOf(sheetId) !== -1) || (rowInputId && sheetId.indexOf(rowInputId) !== -1)) {
          
          // Extract output link ID
          var outVal = (outputCol !== -1 && outputCol < row.length) ? String(row[outputCol] || '').trim() : '';
          if (outVal) {
            var m = outVal.match(/\/d\/(.*?)(\/|$)/);
            outputSheetId = m ? m[1] : outVal;
          }
          
          // Extract teaching plan link ID
          var tpVal = (tpCol !== -1 && tpCol < row.length) ? String(row[tpCol] || '').trim() : '';
          if (tpVal) {
            var m = tpVal.match(/\/d\/(.*?)(\/|$)/);
            teachingPlanId = m ? m[1] : tpVal;
          }
          break;
        }
      }
    }
  } catch(err) {
    Logger.log("Error looking up from master config sheet: " + err.message);
  }

  // 2. Fallback to parsing the individual subjects sheet tab if not resolved above
  if (!teachingPlanId || !outputSheetId) {
    try {
      var ss = _getSpreadsheet(sheetId);
      var ws = ss.getSheetByName('subjects');
      if (ws) {
        var data = ws.getDataRange().getValues();
        var tpColIdx = -1;
        var outColIdx = -1;
        var codeColIdx = 0;

        var headers = data[0];
        for (var c = 0; c < headers.length; c++) {
          var val = String(headers[c]).toLowerCase().trim();
          if (val.indexOf('teaching plan') !== -1 || val.indexOf('syllabus') !== -1) tpColIdx = c;
          if (val.indexOf('output excel') !== -1 || val.indexOf('output sheet') !== -1 || val.indexOf('output link') !== -1) outColIdx = c;
        }

        var inputParsed = _parseSubjectCode(code);
        for (var i = 1; i < data.length; i++) {
          var rowCode = String(data[i][codeColIdx]).trim();
          var rowParsed = _parseSubjectCode(rowCode);
          if (rowParsed.cleanBaseCode === inputParsed.cleanBaseCode || rowCode.toLowerCase() === code.trim().toLowerCase()) {
            if (!teachingPlanId && tpColIdx !== -1 && data[i][tpColIdx]) {
              var m = String(data[i][tpColIdx]).match(/\/d\/(.*?)(\/|$)/);
              teachingPlanId = m ? m[1] : String(data[i][tpColIdx]).trim();
            }
            if (!outputSheetId && outColIdx !== -1 && data[i][outColIdx]) {
              var m = String(data[i][outColIdx]).match(/\/d\/(.*?)(\/|$)/);
              outputSheetId = m ? m[1] : String(data[i][outColIdx]).trim();
            }
            break;
          }
        }
      }
    } catch(err) {
      Logger.log("Error looking up from subjects sheet tab: " + err.message);
    }
  }

  // Final default fallbacks
  if (!teachingPlanId) teachingPlanId = sheetId;
  if (!outputSheetId) outputSheetId = getOutputSheetId(sheetId);

  return { teachingPlanId: teachingPlanId, outputSheetId: outputSheetId };
}

/**
 * Lookup the global teaching plan link from master config sheet.
 * Used as fallback when a college's subjects sheet has no teaching plan column.
 */
function getGlobalTeachingPlanLink(sheetId) {
  try {
    var MASTER_CONFIG_SHEET_ID = "1p3WoC2s-YYqn9ekqkQ72banxAAd-ujlDoFYpv4fkXmk";
    var masterSs = SpreadsheetApp.openById(MASTER_CONFIG_SHEET_ID);
    var masterWs = masterSs.getSheetByName("smart attendance client sheet") || masterSs.getSheets()[0];
    if (!masterWs) return '';

    var data = masterWs.getDataRange().getValues();
    var headers = data[2] || data[0];
    var inputCol = -1, tpCol = -1;

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).toLowerCase().trim();
      if (h.indexOf('input sheet id') !== -1 || h.indexOf('input link') !== -1) inputCol = c;
      if (h.indexOf('teaching plan link') !== -1 || h.indexOf('teaching plan') !== -1) tpCol = c;
    }

    if (inputCol === -1) inputCol = 4;
    if (tpCol === -1) tpCol = 6;

    for (var r = 3; r < data.length; r++) {
      var rowInputId = String(data[r][inputCol] || '').trim();
      if (rowInputId === sheetId || (sheetId && rowInputId.indexOf(sheetId) !== -1) || (rowInputId && sheetId.indexOf(rowInputId) !== -1)) {
        var tpVal = (tpCol !== -1 && tpCol < data[r].length) ? String(data[r][tpCol] || '').trim() : '';
        return tpVal || '';
      }
    }
  } catch(e) {
    Logger.log("Error getting global teaching plan link: " + e.message);
  }
  return '';
}

/* ═══════════════════════════════════════════════════════════════
   SMART ATTENDANCE LOGIC (UNTOUCHED)
   ═══════════════════════════════════════════════════════════════ */

function saveAttendance(records, outputSheetId, collegeName, managementName, sheetId) {
  if (!records || !records.length) return { error: 'No data' };
  if (!outputSheetId) outputSheetId = getOutputSheetId(sheetId);
  var res = updateOutputMatrix(records, outputSheetId, collegeName, managementName, sheetId);
  if (res === true) {
    // Automatically trigger teaching plan sync for the subject
    try {
      var code = records[0].code;
      var faculty = records[0].faculty || 'Assigned';
      syncTeachingPlan(code, faculty, sheetId);
    } catch(e) {
      Logger.log("Auto-sync teaching plan failed: " + e.message);
    }
    return { success: true, saved: records.length };
  }
  return { success: false, error: String(res) };
}

function updateOutputMatrix(records, outputSheetId, _collegeName, _managementName, sheetId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return "Lock Timeout"; }
  try {
    var outSs = SpreadsheetApp.openById(outputSheetId);
    var grouped = {};
    for (var i = 0; i < records.length; i++) {
      var r = records[i], tab = r.code + " - " + getSubjectName(r.code, sheetId);
      if (r.batch) tab += " - Batch " + r.batch;
      if (!grouped[tab]) grouped[tab] = {};
      if (!grouped[tab][r.date]) grouped[tab][r.date] = [];
      grouped[tab][r.date].push(r);
    }
    var limit = (getAttendanceLimit(sheetId).limit || 75) / 100;
    var config = { collegeName: _collegeName || '', managementName: _managementName || '' };
    if (!config.collegeName || !config.managementName) {
      try {
        var cs = _getSpreadsheet(sheetId).getSheetByName('client sheet') || _getSpreadsheet(sheetId).getSheetByName('subjects');
        if (cs) {
          var cd = cs.getDataRange().getValues();
          for (var cr = 0; cr < cd.length; cr++) {
            for (var cc = 0; cc < cd[cr].length; cc++) {
              var cv = String(cd[cr][cc]).trim().toLowerCase();
              if (!config.collegeName && cv.indexOf('college name') !== -1) {
                if (cr+1 < cd.length) { var bv = String(cd[cr+1][cc]).trim(); if (bv) config.collegeName = bv; }
              }
              if (!config.managementName && cv.indexOf('management name') !== -1) {
                if (cr+1 < cd.length) { var bv = String(cd[cr+1][cc]).trim(); if (bv) config.managementName = bv; }
              }
            }
          }
        }
      } catch(ce) {}
    }
    for (var tab in grouped) {
        var dates = grouped[tab], dKeys = Object.keys(dates), sheet = outSs.getSheetByName(tab);
        if (!sheet) {
            sheet = outSs.insertSheet(tab);
            sheet.getRange("A1:K1").mergeAcross(); sheet.getRange("A2:K2").mergeAcross(); sheet.getRange("A4:K4").mergeAcross();
            sheet.getRange(6, 1, 1, 6).setValues([["Roll No.", "Name", "Total P", "Total A", "Total", "% Att."]]).setFontWeight("bold").setBackground("#F1F5F9").setHorizontalAlignment("center");
            sheet.getRange("B7").setValue("Topic").setFontWeight("bold").setHorizontalAlignment("left");
            var f = dates[dKeys[0]][0], sts = getStudents(f.year, f.batch, sheetId).students || [];
            sts.sort(function(a,b){return parseInt(a.rollNo)-parseInt(b.rollNo)});
            var sd = sts.map(function(s){return [s.rollNo, s.name, 0, 0, 0, 0]});
            if (sd.length > 0) sheet.getRange(8, 1, sd.length, 6).setValues(sd);
            sheet.setColumnWidth(1, 80); sheet.setColumnWidth(2, 280);
        }
        
        // Find header row dynamically if it exists
        var sheetData = sheet.getDataRange().getValues();
        var hdrRowIdx = -1;
        for (var r = 0; r < Math.min(sheetData.length, 30); r++) {
            var rowStr = sheetData[r].map(function(cell) { return String(cell).toLowerCase().trim(); }).join('|');
            if (rowStr.indexOf('roll no') !== -1 && rowStr.indexOf('name') !== -1 && (rowStr.indexOf('total p') !== -1 || rowStr.indexOf('% att') !== -1)) {
                hdrRowIdx = r;
                break;
            }
        }
        if (hdrRowIdx === -1) {
            hdrRowIdx = 5; // Default Row 6 (index 5)
        }
        var hdrRowNumber = hdrRowIdx + 1;
        
        for (var k = 0; k < dKeys.length; k++) {
            var dateKey = dKeys[k], recs = dates[dateKey], dispDate = dbToDisplay(dateKey);
            var hRows = sheet.getRange(hdrRowNumber, 1, 1, Math.max(sheet.getLastColumn(), 10)).getDisplayValues()[0];
            var dCol = -1, tpCol = -1;
            for (var c = 0; c < hRows.length; c++) {
                var val = hRows[c].trim().toLowerCase();
                if (val === dispDate.toLowerCase()) dCol = c + 1;
                if (val.indexOf("total p") !== -1) tpCol = c + 1;
            }
            
            // Find name column dynamically
            var nameColIdx = -1;
            for (var c = 0; c < hRows.length; c++) {
                if (hRows[c].trim().toLowerCase().indexOf('name') !== -1) {
                    nameColIdx = c;
                    break;
                }
            }
            if (nameColIdx === -1) nameColIdx = 1; // Fallback B (index 1)
            
            if (dCol === -1 && tpCol !== -1) {
                sheet.insertColumnBefore(tpCol); dCol = tpCol;
                sheet.getRange(hdrRowNumber, dCol).setValue(dispDate).setFontWeight("bold").setBackground("#F1F5F9").setHorizontalAlignment("center");
                sheet.setColumnWidth(dCol, 100);
                var rs = sheet.getLastRow() - (hdrRowNumber + 1); // Student data starts at hdrRowNumber + 2
                if (rs > 0) {
                    var tpL = columnToLetter(dCol+1), taL = columnToLetter(dCol+2), tL = columnToLetter(dCol+3), deL = columnToLetter(dCol);
                    var firstDateLetter = columnToLetter(nameColIdx + 2); // since first date starts at nameColIdx + 2 (1-indexed)
                    var fms = [];
                    for (var r=0; r<rs; r++) {
                        var rn = r + (hdrRowNumber + 2);
                        fms.push([
                          '=COUNTIF(' + firstDateLetter + rn + ':' + deL + rn + ', "P")', 
                          '=COUNTIF(' + firstDateLetter + rn + ':' + deL + rn + ', "A")', 
                          '=' + tpL + rn + '+' + taL + rn, 
                          '=IF(' + tL + rn + '>0,' + tpL + rn + '/' + tL + rn + ',0)'
                        ]);
                    }
                    sheet.getRange(hdrRowNumber + 2, dCol+1, rs, 4).setFormulas(fms);
                    setupFormulasAndConditions(sheet, rs, dCol+4, hdrRowNumber + 2, nameColIdx + 2, limit*100);
                }
            }
            if (dCol !== -1) {
                var topic = recs[0] && recs[0].topic ? recs[0].topic : "";
                sheet.getRange(hdrRowNumber + 1, dCol).setValue(topic).setFontStyle("italic").setHorizontalAlignment("center");
                
                var rs = sheet.getLastRow() - (hdrRowNumber + 1);
                if (rs > 0) {
                    var ex = sheet.getRange(hdrRowNumber + 2, dCol, rs, 1).getValues(), rolls = sheet.getRange(hdrRowNumber + 2, 1, rs, 1).getValues();
                    var ups = rolls.map(function(r, idx) {
                        var roll = String(r[0]), st = ex[idx][0] || "-";
                        for (var x=0; x<recs.length; x++) { if (String(recs[x].rollNo) === roll) st = recs[x].status; }
                        return [st];
                    });
                    sheet.getRange(hdrRowNumber + 2, dCol, rs, 1).setValues(ups).setHorizontalAlignment("center");
                }
            }
        }
        try {
          var f = dates[dKeys[0]][0], info = getSubjectInfo(f.code, sheetId);
          var row4 = f.code + " - " + info.name + (f.batch ? " | Batch " + f.batch : "") + " | " + info.program + " | " + info.year + " | 01 Jan 2020 to " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy");
          sheet.getRange("A1:K1").unmerge().mergeAcross().setValue(config.managementName || "Management Name").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
          sheet.getRange("A2:K2").unmerge().mergeAcross().setValue(config.collegeName || "College Name").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
          sheet.getRange("A4:K4").unmerge().mergeAcross().setValue(row4).setFontWeight("bold").setBackground("#E2E8F0").setHorizontalAlignment("center");
        } catch(e) {}
    }
    return true;
  } catch(e) { return e.message; } finally { lock.releaseLock(); }
}

function getSubjectName(code, sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects');
  if (!ws) return "Unknown";
  var data = ws.getDataRange().getValues();
  for (var i=1; i<data.length; i++) { if (String(data[i][0]).trim() === String(code).trim()) return String(data[i][1]).trim(); }
  return "Unknown";
}

function getSubjectInfo(code, sheetId) {
  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects');
  if (!ws) return { name: "Unknown", program: "", year: "" };
  var data = ws.getDataRange().getValues();
  for (var i=1; i<data.length; i++) { if (String(data[i][0]).trim() === String(code).trim()) return { name: String(data[i][1]).trim(), year: String(data[i][2]).trim(), program: String(data[i][3]).trim() }; }
  return { name: "Unknown", program: "", year: "" };
}

function dbToDisplay(db) {
  if (!db) return '';
  var m = String(db), s = ""; if (m.indexOf(' (') !== -1) { s = m.substring(m.indexOf(' (')); m = m.substring(0, m.indexOf(' (')); }
  var p = m.split('_')[0].split('-'); if (p.length < 3) return m + s;
  var mos = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var dd = parseInt(p[2]); var dStr = dd < 10 ? '0' + dd : String(dd);
  return dStr + '-' + mos[parseInt(p[1])-1] + s;
}

function displayToDb(disp) {
  var m = String(disp), s = ""; if (m.indexOf(' (') !== -1) { s = m.substring(m.indexOf(' (')); m = m.substring(0, m.indexOf(' (')); }
  var p = m.split('-'); if (p.length !== 2) return disp;
  var mos = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var mi = mos.indexOf(p[1]) + 1;
  var mm = mi < 10 ? '0' + mi : String(mi);
  var dd = parseInt(p[0]); var dStr = dd < 10 ? '0' + dd : String(dd);
  return new Date().getFullYear() + '-' + mm + '-' + dStr + s;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) { temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26; }
  return letter;
}

function setupFormulasAndConditions(sheet, rows, pctCol, startRow, startCol, limit) {
  sheet.getRange(startRow, pctCol, rows, 1).setNumberFormat('0.0%');
  var dataR = sheet.getRange(startRow, startCol, 1000, Math.max(pctCol - startCol, 1));
  var pctR = sheet.getRange(startRow, pctCol, 1000, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("P").setFontColor("#15803D").setBackground("#DCFCE7").setRanges([dataR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("A").setFontColor("#B91C1C").setBackground("#FEE2E2").setRanges([dataR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(limit / 100).setFontColor("#14532D").setBackground("#BBF7D0").setBold(true).setRanges([pctR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(limit / 100).setFontColor("#7F1D1D").setBackground("#FECACA").setBold(true).setRanges([pctR]).build()
  ]);
}

function getAttendance(code, year, date, outputSheetId, sheetId) {
  if (!code) return { error: 'No code' };
  if (!outputSheetId) outputSheetId = getOutputSheetId(sheetId);
  var outSs; try { outSs = SpreadsheetApp.openById(outputSheetId); } catch(e) { return { error: 'Scan Fail' }; }
  var res = [], sheets = outSs.getSheets();
  var parsedInput = _parseSubjectCode(code);
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i], name = s.getName();
    var parsedSheetCode = _parseSubjectCode(name);
    var cleanSheetName = name.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (parsedSheetCode.cleanBaseCode !== parsedInput.cleanBaseCode && cleanSheetName.indexOf(parsedInput.cleanBaseCode) !== 0) continue;
    var batch = name.indexOf(" - Batch ") !== -1 ? name.substring(name.indexOf(" - Batch ") + 9).trim() : "";
    var lc = s.getLastColumn(), lr = s.getLastRow();
    if (lc < 6 || lr < 8) continue;
    
    // Find header row dynamically
    var attData = s.getDataRange().getValues();
    var hdrRowIdx = -1;
    for (var r = 0; r < Math.min(attData.length, 30); r++) {
      var rowStr = attData[r].map(function(cell) { return String(cell).toLowerCase().trim(); }).join('|');
      if (rowStr.indexOf('roll no') !== -1 && rowStr.indexOf('name') !== -1 && (rowStr.indexOf('total p') !== -1 || rowStr.indexOf('% att') !== -1)) {
        hdrRowIdx = r;
        break;
      }
    }
    if (hdrRowIdx === -1) {
      hdrRowIdx = 5; // Default Row 6 (index 5)
    }
    var hdrRowNumber = hdrRowIdx + 1;

    var hdrs = s.getRange(hdrRowNumber, 1, 1, lc).getDisplayValues()[0];
    var raw = s.getRange(hdrRowNumber, 1, 1, lc).getValues()[0];
    
    // Find name and Total P columns dynamically
    var nameColIdx = -1;
    var totalPColIdx = -1;
    for (var c = 0; c < hdrs.length; c++) {
      var val = hdrs[c].toLowerCase().trim();
      if (val.indexOf('name') !== -1) {
        nameColIdx = c;
      }
      if (val.indexOf('total p') !== -1) {
        totalPColIdx = c;
        break;
      }
    }
    if (nameColIdx === -1) nameColIdx = 1;
    if (totalPColIdx === -1) {
      for (var c = 0; c < hdrs.length; c++) {
        var val = hdrs[c].toLowerCase().trim();
        if (val.indexOf('total') !== -1 || val.indexOf('% att') !== -1) {
          totalPColIdx = c;
          break;
        }
      }
    }
    if (totalPColIdx === -1) totalPColIdx = lc;

    var dates = [];
    var firstDateColIdx = nameColIdx + 1;
    for (var c = firstDateColIdx; c < totalPColIdx; c++) {
       if (hdrs[c].trim()) dates.push({ index: c, disp: hdrs[c].trim() });
    }
    if (dates.length === 0) continue;
    var topics = s.getRange(hdrRowNumber + 1, 1, 1, lc).getValues()[0];
    var mtx = s.getRange(hdrRowNumber + 2, 1, lr - (hdrRowNumber + 1), lc).getValues();
    for (var r = 0; r < mtx.length; r++) {
       for (var d = 0; d < dates.length; d++) {
          var st = String(mtx[r][dates[d].index]).trim();
          if (st === 'P' || st === 'A') {
             var dbD = displayToDb(dates[d].disp);
             if (date && dbD.indexOf(date) === -1) continue;
             res.push({ date: dbD, code: code, year: year, batch: batch, faculty: "Assigned", rollNo: mtx[r][0], name: mtx[r][1], status: st, topic: String(topics[dates[d].index] || '') });
          }
       }
    }
  }
  return { success: true, records: res };
}

/* ═══════════════════════════════════════════════════════════════
   ACADEMIC FILE SYLLABUS LOGIC
   ═══════════════════════════════════════════════════════════════ */

function getTeachingPlan(code, teacher, sheetId) {
  if (!code) return { success: false, error: 'Missing subject code' };
  
  function parseAndFormatDate(val, timeZone) {
    if (!val) return '';
    if (val instanceof Date || Object.prototype.toString.call(val) === '[object Date]') {
      try {
        return Utilities.formatDate(val, timeZone, 'yyyy-MM-dd');
      } catch(e) {}
    }
    var str = String(val).trim();
    
    // 1. If it's already yyyy-MM-dd
    var ymdRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (ymdRegex.test(str)) {
      return str;
    }
    
    // 2. Parse DD/MM/YY or DD/MM/YYYY or DD-MM-YYYY or DD.MM.YY format (standard Indian/British formats used by faculty)
    var slashRegex = /^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/;
    var match = str.match(slashRegex);
    if (match) {
      var d = parseInt(match[1], 10);
      var m = parseInt(match[2], 10);
      var y = parseInt(match[3], 10);
      if (y < 100) {
        y += 2000;
      }
      if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
        try {
          var dateObj = new Date(y, m - 1, d);
          return Utilities.formatDate(dateObj, timeZone, 'yyyy-MM-dd');
        } catch(e) {}
      }
    }
    
    // 3. Fallback: DD-MMM-YY (e.g. 13-Jul-26) or other standard patterns
    var dmyRegex = /^\d{1,2}-[A-Za-z]{3}-\d{2,4}$/;
    if (dmyRegex.test(str)) {
      return str;
    }
    
    try {
      var parsed = new Date(str);
      if (!isNaN(parsed.getTime())) {
        return Utilities.formatDate(parsed, timeZone, 'yyyy-MM-dd');
      }
    } catch(e) {}
    return str;
  }
  
  var targetIds = getTargetSheetIds(code, sheetId);
  var tpSs = _getSpreadsheet(targetIds.teachingPlanId);
  var ws = _findSheetByCode(tpSs, code);
  
  if (!ws) {
    return { success: false, error: 'Teaching plan sheet for subject code "' + code + '" not found in spreadsheet.' };
  }
  
  var data = ws.getDataRange().getValues();
  if (data.length < 5) {
    return { success: false, error: 'Teaching plan sheet does not match required format (too short)' };
  }

  // Find the header row dynamically
  var headerRowIdx = -1;
  for (var i = 0; i < Math.min(data.length, 25); i++) {
    for (var j = 0; j < data[i].length; j++) {
      var val = String(data[i][j]).toLowerCase().trim();
      if (val === 'syllabus' || val === 'lecture/practical no' || val === 'lecture/practical no.' || val === 'practical no' || val === 'experiment no' || val === 'expt no' || val === 'lab no' || val === 'topic') {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx !== -1) break;
  }
  
  if (headerRowIdx === -1) {
    headerRowIdx = 10; // Fallback to Row 11 (index 10) if not found
  }

  // Dynamic Metadata Parser
  function findMetadataValue(keywords, defaultRow, defaultCol) {
    for (var r = 0; r < Math.min(data.length, headerRowIdx); r++) {
      for (var c = 0; c < data[r].length; c++) {
        var val = String(data[r][c]).toLowerCase().trim();
        for (var k = 0; k < keywords.length; k++) {
          var kw = keywords[k].toLowerCase();
          if (val.indexOf(kw) !== -1) {
            // Exclude false positives for course/class
            if (kw === 'course' && (val.indexOf('outcome') !== -1 || val.indexOf('co') !== -1)) continue;
            
            if (val.indexOf(':') !== -1) {
              var parts = String(data[r][c]).split(':');
              if (parts[1] && parts[1].trim()) return parts[1].trim();
            }
            if (c + 1 < data[r].length) {
              var rv = String(data[r][c + 1]).trim();
              if (rv && rv.toLowerCase().indexOf(kw) === -1) return rv;
            }
            if (r + 1 < data.length) {
              var bv = String(data[r + 1][c]).trim();
              if (bv && bv.toLowerCase().indexOf(kw) === -1) return bv;
            }
            var cellVal = String(data[r][c]).trim();
            if (cellVal.length > kw.length + 2) {
              return cellVal;
            }
          }
        }
      }
    }
    try {
      if (data[defaultRow] && data[defaultRow][defaultCol] !== undefined) {
        return String(data[defaultRow][defaultCol]).trim();
      }
    } catch(e) {}
    return '';
  }

  var managementName = findMetadataValue(["management name", "management", "society", "sinhgad"], 5, 3);
  var collegeName = findMetadataValue(["college name", "college", "institute", "rmd"], 6, 3);
  var academicYear = findMetadataValue(["academic year", "year", "ay"], 7, 3);
  var course = findMetadataValue(["course"], 8, 2);
  var classCourse = findMetadataValue(["class"], 8, 3);
  var faculty = findMetadataValue(["faculty", "teacher", "instructor"], 8, 4);
  var subject = findMetadataValue(["subject"], 8, 5);
  
  var totalLectures = 0;
  var totalTutorials = 0;
  
  try {
    var foundLectures = false;
    for (var r = 0; r < data.length; r++) {
      for (var c = 0; c < data[r].length; c++) {
        var cellVal = String(data[r][c]).toLowerCase().trim();
        if (cellVal.indexOf('total lectures/practical') !== -1 || cellVal.indexOf('total lectures') !== -1 || cellVal.indexOf('total practicals') !== -1) {
          for (var c2 = c + 1; c2 < data[r].length; c2++) {
            var val = parseInt(data[r][c2]);
            if (!isNaN(val) && val > 0) {
              totalLectures = val;
              foundLectures = true;
              break;
            }
          }
        }
        if (foundLectures) break;
      }
      if (foundLectures) break;
    }
    
    if (!foundLectures) {
      if (data[12] && data[12].length > 8) totalLectures = parseInt(data[12][8]) || 0;
      if (data[13] && data[13].length > 8) totalTutorials = parseInt(data[13][8]) || 0;
    }
  } catch(e) {}

  var topics = [];
  var startRow = headerRowIdx + 1;
  var colIdxSyllabus = 2; // Column C is index 2
  var colIdxLectureNo = 1; // Column B is index 1
  var colIdxPlanned = 3; // Column D is index 3
  var colIdxExecuted = 4; // Column E is index 4
  var colIdxRemark = 5; // Column F is index 5
  
  var headerRow = data[headerRowIdx];
  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c]).toLowerCase().trim();
    if (h === 'syllabus' || h.indexOf('syllabus') !== -1 || h.indexOf('topic') !== -1 || h.indexOf('experiment') !== -1 || h.indexOf('particulars') !== -1) colIdxSyllabus = c;
    if (h.indexOf('lecture/practical no') !== -1 || h.indexOf('lecture no') !== -1 || h.indexOf('practical no') !== -1 || h.indexOf('expt no') !== -1 || h.indexOf('experiment no') !== -1 || h.indexOf('lab no') !== -1 || h.indexOf('sr.no') !== -1 || h.indexOf('sr no') !== -1) colIdxLectureNo = c;
    if (h.indexOf('planned') !== -1) colIdxPlanned = c;
    if (h.indexOf('execution') !== -1 || h.indexOf('executed') !== -1) colIdxExecuted = c;
    if (h.indexOf('remark') !== -1) colIdxRemark = c;
  }

  for (var i = startRow; i < data.length; i++) {
    var row = data[i];
    if (row.length < 3) continue;
    var syllabusText = String(row[colIdxSyllabus] || '').trim();
    var lectNo = String(row[colIdxLectureNo] || '').trim();
    if (!syllabusText || lectNo.toLowerCase() === 'syllabus') continue;
    
    var plannedDate = parseAndFormatDate(row[colIdxPlanned], Session.getScriptTimeZone());
    var executedDate = parseAndFormatDate(row[colIdxExecuted], Session.getScriptTimeZone());

    topics.push({
      rowIndex: i + 1,
      lectureNo: lectNo,
      syllabus: syllabusText,
      plannedDate: plannedDate,
      executedDate: executedDate,
      remark: String(row[colIdxRemark] || '').trim()
    });
  }

  // Deduplicate topics (merge duplicate lecture rows if sheet has repeated tables)
  var uniqueTopics = [];
  var seenKeys = {};
  for (var k = 0; k < topics.length; k++) {
    var top = topics[k];
    var key = String(top.lectureNo).trim().toLowerCase() + '_' + String(top.syllabus).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (!seenKeys[key]) {
      seenKeys[key] = top;
      uniqueTopics.push(top);
    } else {
      if (!seenKeys[key].executedDate && top.executedDate) {
        seenKeys[key].executedDate = top.executedDate;
      } else if (seenKeys[key].executedDate && top.executedDate && String(seenKeys[key].executedDate).indexOf(String(top.executedDate)) === -1) {
        seenKeys[key].executedDate = String(seenKeys[key].executedDate).trim() + ', ' + String(top.executedDate).trim();
      }
      if (!seenKeys[key].remark && top.remark) {
        seenKeys[key].remark = top.remark;
      }
    }
  }
  topics = uniqueTopics;

  var conductedCount = topics.filter(function(t) { return t.executedDate !== ''; }).length;
  var percent = topics.length > 0 ? Math.round((conductedCount / topics.length) * 100) : 0;
  var parsedSubjectCodeInfo = _parseSubjectCode(code, '', subject);

  return {
    success: true,
    metadata: {
      managementName: managementName || 'Sinhgad Technical Education Society',
      collegeName: collegeName || 'RMDIPER',
      academicYear: academicYear || '2024-25',
      course: course,
      classCourse: classCourse,
      faculty: faculty,
      subject: subject,
      isPractical: parsedSubjectCodeInfo.isPractical,
      totalLectures: totalLectures,
      totalTutorials: totalTutorials,
      percent: percent,
      conductedCount: conductedCount,
      totalTopics: topics.length,
      colIdxSyllabus: colIdxSyllabus,
      colIdxLectureNo: colIdxLectureNo,
      colIdxPlanned: colIdxPlanned,
      colIdxExecuted: colIdxExecuted,
      colIdxRemark: colIdxRemark
    },
    topics: topics
  };
}

function syncTeachingPlan(code, teacher, sheetId) {
  if (!code) return { success: false, error: 'Missing subject code' };

  function _normDate(d) {
    if (d === null || d === undefined || d === '') return '';
    var dt = (d instanceof Date) ? d : new Date(d);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return String(d).trim();
  }

  var targetIds = getTargetSheetIds(code, sheetId);
  var outSsId = targetIds.outputSheetId;
  
  if (!outSsId) return { success: false, error: 'Attendance Output Sheet Link not found.' };
  
  var planResult = getTeachingPlan(code, teacher, sheetId);
  if (!planResult.success) return planResult;
  
  var outSs = _getSpreadsheet(outSsId);
  var sheets = outSs.getSheets();
  var attSheet = null;
  var parsedInput = _parseSubjectCode(code);
  
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().toUpperCase();
    var parsedSheet = _parseSubjectCode(name);
    var cleanSheetName = name.replace(/[^A-Z0-9]/g, '');
    if (parsedSheet.cleanBaseCode === parsedInput.cleanBaseCode || cleanSheetName.indexOf(parsedInput.cleanBaseCode) === 0 || name.indexOf(parsedInput.baseCode.toUpperCase()) === 0 || name === code.toUpperCase()) {
      attSheet = sheets[i];
      break;
    }
  }
  
  if (!attSheet) {
    return { success: true, topics: planResult.topics, metadata: planResult.metadata, warning: 'Attendance logs sheet not found' };
  }
  
  var attData = attSheet.getDataRange().getValues();
  if (attData.length < 7) {
    return { success: true, topics: planResult.topics, metadata: planResult.metadata, warning: 'Empty attendance sheet' };
  }

  // Find the header row index dynamically
  var hdrRowIdx = -1;
  for (var r = 0; r < Math.min(attData.length, 30); r++) {
    var rowStr = attData[r].map(function(cell) { return String(cell).toLowerCase().trim(); }).join('|');
    if (rowStr.indexOf('roll no') !== -1 && rowStr.indexOf('name') !== -1 && (rowStr.indexOf('total p') !== -1 || rowStr.indexOf('% att') !== -1)) {
      hdrRowIdx = r;
      break;
    }
  }
  if (hdrRowIdx === -1) {
    hdrRowIdx = 5; // Default fallback to Row 6 (index 5)
  }

  var dateRow = attData[hdrRowIdx];
  var topicRow = attData[hdrRowIdx + 1];
  var lastCol = attSheet.getLastColumn();
  
  // Find name and Total P columns dynamically to locate date columns
  var nameColIdx = -1;
  var totalPColIdx = -1;
  for (var c = 0; c < dateRow.length; c++) {
    var val = String(dateRow[c]).toLowerCase().trim();
    if (val.indexOf('name') !== -1) {
      nameColIdx = c;
    }
    if (val.indexOf('total p') !== -1) {
      totalPColIdx = c;
      break;
    }
  }
  if (nameColIdx === -1) nameColIdx = 1; // Fallback to B (index 1)
  if (totalPColIdx === -1) {
    for (var c = 0; c < dateRow.length; c++) {
      var val = String(dateRow[c]).toLowerCase().trim();
      if (val.indexOf('total') !== -1 || val.indexOf('% att') !== -1) {
        totalPColIdx = c;
        break;
      }
    }
  }
  if (totalPColIdx === -1) totalPColIdx = lastCol;

  var firstDateColIdx = nameColIdx + 1;
  
  var attendanceLogs = [];
  var seenLogKeys = {};
  for (var c = firstDateColIdx; c < totalPColIdx; c++) {
    var val = String(dateRow[c]).trim().toLowerCase();
    if (val.indexOf('total p') !== -1 || val.indexOf('% att') !== -1 || val === '') break;
    
    var dateHeader = dateRow[c];
    var topicConducted = String(topicRow[c] || '').trim();
    var dateIso = _normDate(dateHeader);

    if (topicConducted && dateIso) {
      var key = dateIso + '||' + topicConducted.toLowerCase();
      if (!seenLogKeys[key]) {
        seenLogKeys[key] = true;
        attendanceLogs.push({ date: dateIso, topic: topicConducted });
      }
    }
  }

  var tpSs = _getSpreadsheet(targetIds.teachingPlanId);
  var ws = _findSheetByCode(tpSs, code);
  
  var updated = 0;
  var topics = planResult.topics;
  
  function cleanStr(s) {
    return s.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { return { success: false, error: 'Lock timeout' }; }

  var executedCol = (planResult.metadata && planResult.metadata.colIdxExecuted !== undefined)
                    ? planResult.metadata.colIdxExecuted + 1
                    : 5;

  try {
    for (var j = 0; j < attendanceLogs.length; j++) {
      var log = attendanceLogs[j];
      var dateStr = log.date;
      var cleanLogTopic = cleanStr(log.topic);
      if (!cleanLogTopic) continue;
      
      var target = null;
      
      // PASS 1: Exact Match (prioritized)
      for (var i = 0; i < topics.length; i++) {
        var t = topics[i];
        var cleanSyllabus = cleanStr(t.syllabus);
        if (cleanSyllabus === cleanLogTopic) {
          target = t;
          break;
        }
      }
      
      // PASS 2: High-Confidence Similarity Match (Only if Pass 1 fails)
      if (!target) {
        var bestRatio = 0;
        for (var i = 0; i < topics.length; i++) {
          var t = topics[i];
          var cleanSyllabus = cleanStr(t.syllabus);
          var minLen = Math.min(cleanSyllabus.length, cleanLogTopic.length);
          var maxLen = Math.max(cleanSyllabus.length, cleanLogTopic.length);
          if (minLen >= 4 && (cleanSyllabus.indexOf(cleanLogTopic) !== -1 || cleanLogTopic.indexOf(cleanSyllabus) !== -1)) {
            var ratio = minLen / maxLen;
            // Strict threshold (at least 55% length similarity) to avoid single generic words matching long detailed titles
            if (ratio >= 0.55 && ratio > bestRatio) {
              bestRatio = ratio;
              target = t;
            }
          }
        }
      }
      
      if (target) {
        var currentCellVal = ws.getRange(target.rowIndex, executedCol).getValue();
        var currentDates = String(currentCellVal || '')
          .split(',')
          .map(function(s) { return _normDate(s.trim()); })
          .filter(function(s) { return s !== ''; });

        if (currentDates.indexOf(dateStr) !== -1) {
          target.executedDate = String(currentCellVal || '');
          continue;
        }

        if (!currentCellVal || String(currentCellVal).trim() === '') {
          target.executedDate = dateStr;
          ws.getRange(target.rowIndex, executedCol).setValue(dateStr);
        } else {
          var newVal = String(currentCellVal).trim() + ", " + dateStr;
          target.executedDate = newVal;
          ws.getRange(target.rowIndex, executedCol).setValue(newVal);
        }
        updated++;
      }
    }
    
    if (updated > 0) {
      SpreadsheetApp.flush();
      planResult = getTeachingPlan(code, teacher, sheetId);
    }
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }

  return planResult;
}

function saveRemark(code, rowIndex, remark, sheetId) {
  if (!code || !rowIndex) return { success: false, error: 'Missing parameters' };
  
  var targetIds = getTargetSheetIds(code, sheetId);
  var tpSs = _getSpreadsheet(targetIds.teachingPlanId);
  var ws = _findSheetByCode(tpSs, code);
  
  if (!ws) return { success: false, error: 'Syllabus sheet not found' };

  var remarkCol = 6;
  try {
    var planResult = getTeachingPlan(code, null, sheetId);
    if (planResult.success && planResult.metadata && planResult.metadata.colIdxRemark !== undefined) {
      remarkCol = planResult.metadata.colIdxRemark + 1;
    }
  } catch(e) {}

  ws.getRange(rowIndex, remarkCol).setValue(remark);
  return { success: true };
}

function addCustomSyllabusTopic(data, sheetId) {
  var code = data.code;
  if (!code) return { success: false, error: 'Missing parameters' };

  var targetIds = getTargetSheetIds(code, sheetId);
  var tpSs = _getSpreadsheet(targetIds.teachingPlanId);
  var ws = _findSheetByCode(tpSs, code);

  if (!ws) return { success: false, error: 'Syllabus sheet not found' };

  // Reuse the same smart column detection as getTeachingPlan so the new
  // row lands in the right columns regardless of the sheet's layout.
  var plan = getTeachingPlan(code, null, sheetId);
  var m = (plan.success && plan.metadata) ? plan.metadata : {};
  var colLectNo = (m.colIdxLectureNo !== undefined) ? m.colIdxLectureNo : 1;
  var colSyllabus = (m.colIdxSyllabus !== undefined) ? m.colIdxSyllabus : 2;
  var colPlanned = (m.colIdxPlanned !== undefined) ? m.colIdxPlanned : 3;
  var colExecuted = (m.colIdxExecuted !== undefined) ? m.colIdxExecuted : 4;
  var colRemark = (m.colIdxRemark !== undefined) ? m.colIdxRemark : 5;

  // Next lecture number = highest numeric lecture no in the plan + 1
  var nextLectNo = 1;
  var topics = (plan.success && plan.topics) ? plan.topics : [];
  for (var i = 0; i < topics.length; i++) {
    var n = parseInt(topics[i].lectureNo, 10);
    if (!isNaN(n) && n >= nextLectNo) nextLectNo = n + 1;
  }
  if (topics.length === 0) nextLectNo = ws.getLastRow() - 10; // legacy fallback

  var dateStr = data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var rowLen = Math.max(colLectNo, colSyllabus, colPlanned, colExecuted, colRemark) + 1;
  var row = [];
  for (var c = 0; c < rowLen; c++) row.push('');
  row[colLectNo] = nextLectNo;
  row[colSyllabus] = data.topic;
  row[colPlanned] = dateStr;
  row[colExecuted] = dateStr;
  row[colRemark] = data.remark || 'Extra lecture conducted';

  ws.appendRow(row);

  return { success: true };
}

/**
 * Option A implementation: Automatically scans the Google Drive folder
 * containing the Spreadsheet for a subfolder named "Academic Calendars & Timetable".
 * If found, scans for files with 'timetable' or 'calendar' in their names,
 * returning their preview URLs.
 */
function getAcademicSchedule(sheetId) {
  try {
    var files = DriveApp.getFileById(sheetId).getParents();
    var parentFolder = files.hasNext() ? files.next() : null;
    var scheduleFolder = null;
    
    if (parentFolder) {
      var subfolders = parentFolder.getFoldersByName("Academic Calendars & Timetable");
      if (subfolders.hasNext()) {
        scheduleFolder = subfolders.next();
      } else {
        // Automatically create it for them if not present!
        scheduleFolder = parentFolder.createFolder("Academic Calendars & Timetable");
      }
    }
    
    if (!scheduleFolder) {
      return { success: false, error: "Parent Google Drive folder not accessible." };
    }
    
    var fileIterator = scheduleFolder.getFiles();
    var allFiles = [];
    
    while (fileIterator.hasNext()) {
      var file = fileIterator.next();
      var thumbLink = '';
      try { thumbLink = file.getThumbnail() ? 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400' : ''; } catch(e) { thumbLink = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400'; }
      allFiles.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        webViewLink: file.getUrl(),
        thumbnailLink: thumbLink,
        lastUpdated: file.getLastUpdated().toISOString()
      });
    }
    
    // Sort by last updated descending (newest first)
    allFiles.sort(function(a, b) { return b.lastUpdated > a.lastUpdated ? 1 : -1; });
    
    return {
      success: true,
      files: allFiles,
      // Backward compat: extract timetable/calendar by keyword
      timetable: allFiles.find(function(f) { var n = f.name.toLowerCase(); return n.indexOf('timetable') > -1 || n.indexOf('time table') > -1 || n.indexOf('schedule') > -1; }) || null,
      calendar: allFiles.find(function(f) { var n = f.name.toLowerCase(); return n.indexOf('calendar') > -1 || n.indexOf('calender') > -1 || n.indexOf('event') > -1; }) || null
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function extractSpreadsheetId(url) {
  if (!url) return null;
  if (url.indexOf('docs.google.com') === -1) {
    return url.trim(); // Assume it's already an ID
  }
  var match = url.match(/\/d\/([a-zA-Z0-9\-_]+)/);
  return match ? match[1] : null;
}

function getSyllabus(link, code, sheetId) {
  try {
    if (!link) {
      return { success: false, error: 'No link provided' };
    }
    var points = getSyllabusPointsFromLink(link, code);
    return { success: true, points: points };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function looksLikeSubjectCode(name) {
  if (!name) return false;
  var val = name.trim().toLowerCase().replace(/\s+/g, "");
  if (val.indexOf("sheet") === 0) return false;
  if (val.indexOf("lecture") === 0) return false;
  if (val.indexOf("unit") === 0) return false;
  if (val.indexOf("chap") === 0) return false;
  var hasLetters = /[a-z]/.test(val);
  var hasNumbers = /[0-9]/.test(val);
  return hasLetters && hasNumbers && val.length >= 3;
}

function getSyllabusPointsFromLink(url, code) {
  var id = extractSpreadsheetId(url);
  if (!id) {
    throw new Error("Invalid Google Sheets link. Please check teaching plan link.");
  }
  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error("Cannot access spreadsheet. Please make sure the link is correct and accessible.");
  }
  
  var sheet = _findSheetByCode(ss, code);
  
  if (!sheet) {
    return [];
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  
  // Find the header row and column dynamically
  var colIdx = -1;
  var headerRowIdx = -1;
  var keywords = ["syllabus points", "syllabus point", "syllabus", "topic name", "topics", "topic", "session topic", "particulars", "description", "content", "practical topic", "experiment name", "experiments", "experiment", "lab topic", "practical"];
  
  for (var r = 0; r < Math.min(data.length, 30); r++) {
    var row = data[r].map(function(h) { return String(h).trim().toLowerCase(); });
    
    // First try exact match in the row
    for (var k = 0; k < keywords.length; k++) {
      var idx = row.indexOf(keywords[k]);
      if (idx !== -1) {
        colIdx = idx;
        headerRowIdx = r;
        break;
      }
    }
    if (colIdx !== -1) break;
    
    // Then try partial match in the row
    for (var j = 0; j < row.length; j++) {
      for (var k = 0; k < keywords.length; k++) {
        if (row[j].indexOf(keywords[k]) !== -1) {
          colIdx = j;
          headerRowIdx = r;
          break;
        }
      }
      if (colIdx !== -1) break;
    }
    if (colIdx !== -1) break;
  }
  
  // Fallbacks if header wasn't found dynamically
  if (colIdx === -1) {
    colIdx = 0;
  }
  if (headerRowIdx === -1) {
    headerRowIdx = 0;
  }
  
  function extractFromCol(targetCol) {
    var pts = [];
    var seen = {};
    var hVal = String(data[headerRowIdx][targetCol] || '').trim().toLowerCase();
    for (var r = headerRowIdx + 1; r < data.length; r++) {
      if (!data[r] || targetCol >= data[r].length) continue;
      var val = String(data[r][targetCol]).trim();
      var lowerVal = val.toLowerCase();
      if (val && lowerVal !== hVal && !seen[lowerVal]) {
        seen[lowerVal] = true;
        pts.push(val);
      }
    }
    return pts;
  }

  var points = extractFromCol(colIdx);

  // If extracted points are mostly pure numbers (e.g. 1 to 45), colIdx picked the Lecture No column by mistake!
  var numberCount = points.filter(function(p) {
    return !isNaN(parseInt(p, 10)) && String(parseInt(p, 10)) === p.trim();
  }).length;

  if (points.length > 0 && numberCount > points.length * 0.5) {
    for (var nextC = colIdx + 1; nextC < Math.min(colIdx + 4, data[headerRowIdx].length); nextC++) {
      var altPoints = extractFromCol(nextC);
      var altNumCount = altPoints.filter(function(p) {
        return !isNaN(parseInt(p, 10)) && String(parseInt(p, 10)) === p.trim();
      }).length;
      if (altPoints.length > 0 && altNumCount <= altPoints.length * 0.5) {
        points = altPoints;
        break;
      }
    }
  }

  return points;
}

// ══════════════════════════════════════
// FIREBASE PUSH NOTIFICATION DISPATCHER
// ══════════════════════════════════════
var FCM_SERVER_KEY = "AIzaSyBuw7HMI__3oNgMbjQz-q2L1aoIcfn5H9k"; // Firebase API Key

/**
 * Send push notification to target topic via Firebase FCM
 * @param {string} title - Notification title
 * @param {string} body - Notification text
 * @param {string} topic - e.g. "teachers", "students", or "all"
 * @param {object} customData - Extra JSON metadata
 */
function sendFCMPushNotification(title, body, topic, customData) {
  topic = topic || "teachers";
  var url = "https://fcm.googleapis.com/fcm/send";
  
  var payload = {
    to: "/topics/" + topic,
    notification: {
      title: title || "VibeMantra Alert",
      body: body || "New update available.",
      icon: "icons/icon-192.png",
      click_action: "FLUTTER_NOTIFICATION_CLICK"
    },
    data: customData || { url: "./index.html" }
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "key=" + FCM_SERVER_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("FCM Response: " + response.getContentText());
    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("FCM Error: " + e.message);
    return { success: false, error: e.message };
  }
}


