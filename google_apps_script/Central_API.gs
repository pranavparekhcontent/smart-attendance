/**
 * ═══════════════════════════════════════════════════════════════
 *  SMART ATTENDANCE CENTRAL API — Google Apps Script Web App (v4.0)
 *  High performance, dedicated backend for Smart Attendance PWA.
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
      // ── Smart Attendance Routes ──
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
      case 'getTaughtTopics':
        result = getTaughtTopics(e.parameter.code, e.parameter.outputSheetId, sheetId);
        break;
      case 'getSyllabus':
        result = getSyllabus(e.parameter.link, e.parameter.code, sheetId);
        break;
      case 'getConfig':
      case 'getAllData': 
        result = getAllData(sheetId); 
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
 * Main POST entry point for Smart Attendance
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || (e.parameter && e.parameter.action);
    var sheetId = data.sheetId || (e.parameter && e.parameter.sheetId);
    var result;

    switch (action) {
      case 'saveAttendance': 
        result = saveAttendance(data.records, data.outputSheetId, data.collegeName, data.managementName, sheetId); 
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

function _parseSubjectCode(code, typeHint, nameHint) {
  var raw = String(code || '').trim();
  if (!raw) {
    return { raw: '', baseCode: '', cleanBaseCode: '', cleanFullCode: '', batch: '', isPractical: false };
  }

  var batch = '';
  var bracketMatch = raw.match(/\((?:batch\s*)?([a-zA-Z0-9]+)\)/i);
  if (bracketMatch && bracketMatch[1]) {
    batch = bracketMatch[1].trim();
  } else {
    var trailingMatch = raw.match(/(?:[\s\-_]+(?:batch[\s\-_]*)?|[[\s\-_]+)([a-zA-Z0-9]{1,3})$/i);
    if (trailingMatch && trailingMatch[1]) {
      var candidate = trailingMatch[1].trim();
      if (!/^\d+[PT]$/i.test(candidate)) {
        batch = candidate;
      }
    }
  }

  var baseCode = raw;
  if (batch) {
    baseCode = raw.replace(/\s*\([^)]*\)/gi, '')
                  .replace(new RegExp('(?:[\\s\\-_]+(?:batch[\\s\\-_]*)?|[\\s\\-_]+)' + batch + '$', 'i'), '')
                  .trim();
  } else {
    baseCode = raw.replace(/\s*\([^)]*\)/gi, '').trim();
  }

  var cleanBaseCode = baseCode.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  var cleanBatch = batch.toUpperCase();
  var cleanFullCode = cleanBaseCode + (cleanBatch ? cleanBatch : '');

  var typeStr = String(typeHint || '').toLowerCase();
  var nameStr = String(nameHint || '').toLowerCase();
  var codeUpper = cleanBaseCode;

  var isPractical = false;
  if (typeStr.indexOf('practical') !== -1 || typeStr.indexOf('lab') !== -1 || typeStr === 'pr' || typeStr === 'p') {
    isPractical = true;
  } else if (nameStr.indexOf('practical') !== -1 || nameStr.indexOf('lab') !== -1) {
    isPractical = true;
  } else if (raw.toLowerCase().indexOf('practical') !== -1 || raw.toLowerCase().indexOf('lab') !== -1 || cleanBatch !== '') {
    isPractical = true;
  } else {
    if (/.*?\d+P$/i.test(codeUpper) || codeUpper.endsWith('P')) {
      isPractical = true;
    }
  }

  return {
    raw: raw,
    baseCode: baseCode,
    cleanBaseCode: cleanBaseCode,
    cleanFullCode: cleanFullCode,
    batch: cleanBatch,
    isPractical: isPractical
  };
}

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

    if (sheetName.toLowerCase() === parsedInput.raw.toLowerCase()) {
      score = 100;
    } else if (parsedSheet.cleanFullCode === parsedInput.cleanFullCode && parsedSheet.batch === parsedInput.batch) {
      score = 95;
    } else if (parsedSheet.cleanFullCode === parsedInput.cleanFullCode) {
      score = 90;
    } else if (parsedSheet.cleanBaseCode === parsedInput.cleanBaseCode && parsedInput.batch && parsedSheet.batch && parsedInput.batch === parsedSheet.batch) {
      score = 85;
    } else if (parsedSheet.cleanBaseCode === parsedInput.cleanBaseCode && !parsedInput.batch && !parsedSheet.batch) {
      score = 80;
    } else if (parsedSheet.cleanBaseCode === parsedInput.cleanBaseCode && parsedInput.batch && !parsedSheet.batch) {
      score = 70;
    } else if (parsedSheet.cleanBaseCode === parsedInput.cleanBaseCode) {
      score = 65;
    } else if (parsedInput.cleanBaseCode && (parsedSheet.cleanBaseCode.indexOf(parsedInput.cleanBaseCode) !== -1 || parsedInput.cleanBaseCode.indexOf(parsedSheet.cleanBaseCode) !== -1)) {
      if (parsedInput.batch && parsedSheet.batch === parsedInput.batch) {
        score = 60;
      } else {
        score = 50;
      }
    }

    if (score > maxScore) {
      maxScore = score;
      bestSheet = sheet;
    }
  }

  if (bestSheet && maxScore >= 50) {
    return bestSheet;
  }

  for (var i = 0; i < sheets.length; i++) {
    var nameLower = sheets[i].getName().trim().toLowerCase();
    if (looksLikeSubjectCode(nameLower) && _parseSubjectCode(nameLower).cleanBaseCode !== parsedInput.cleanBaseCode) {
      continue;
    }
    if (nameLower.indexOf("syllabus") !== -1 || nameLower.indexOf("teaching plan") !== -1 || nameLower.indexOf("plan") !== -1) {
      return sheets[i];
    }
  }

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
  var cache = CacheService.getScriptCache();
  var cacheKey = 'allData_' + (sheetId || '');
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  var ss = _getSpreadsheet(sheetId), ws = ss.getSheetByName('subjects'), subs = [], config = { collegeName: '', managementName: '' };
  var teachers = [], limit = 75;
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
    // --- Inline teachers extraction (reuse already-read data, avoid re-reading subjects) ---
    var tMap = {};
    for (var i = 1; i < data.length; i++) {
      var fStr = String(data[i][cols.faculty]).trim(), pStr = String(data[i][cols.pin]).trim();
      if (fStr && fStr !== 'undefined') {
        var fs = fStr.split(','), ps = pStr.split(',');
        for (var f = 0; f < fs.length; f++) {
          var n = fs[f].trim(), p = (ps[f] && ps[f].trim()) || ps[0].trim();
          if (n) {
            if (!tMap[n]) tMap[n] = p;
            else if (tMap[n].split(',').indexOf(p) === -1) tMap[n] += ',' + p;
          }
        }
      }
    }
    for (var k in tMap) teachers.push({ name: k, pin: tMap[k] });
    // --- Inline attendance limit extraction (reuse already-read data) ---
    for (var i = 0; i < data.length; i++) {
      for (var j = 0; j < data[i].length; j++) {
        if (String(data[i][j]).toLowerCase().indexOf('attendance limit') !== -1 && j + 1 < data[i].length) {
          var v = Number(data[i][j + 1]); if (!isNaN(v) && v > 0) limit = v; break;
        }
      }
    }
  }
  var result = { success: !!ws, teachers: teachers, subjects: subs, attendanceLimit: limit, config: config };
  if (ws && (teachers.length > 0 || subs.length > 0)) {
    try { cache.put(cacheKey, JSON.stringify(result), 3600); } catch(ce) {}
  }
  return result;
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

function _getCollegeSheetIds(sheetId) {
  var teachingPlanId = '';
  var outputSheetId = '';
  
  if (!sheetId) return { outputSheetId: '', teachingPlanId: '' };
  
  var cache = CacheService.getScriptCache();
  var cacheKeyOut = 'outLink_' + sheetId;
  var cacheKeyTp = 'tpLink_' + sheetId;
  var cachedOut = cache.get(cacheKeyOut);
  var cachedTp = cache.get(cacheKeyTp);
  
  if (cachedOut !== null && cachedTp !== null) {
    return {
      outputSheetId: cachedOut === 'NONE' ? '' : cachedOut,
      teachingPlanId: cachedTp === 'NONE' ? '' : cachedTp
    };
  }

  try {
    var MASTER_CONFIG_SHEET_ID = "1p3WoC2s-YYqn9ekqkQ72banxAAd-ujlDoFYpv4fkXmk";
    var masterSs = _getSpreadsheet(MASTER_CONFIG_SHEET_ID);
    var masterWs = masterSs.getSheetByName("smart attendance client sheet") || masterSs.getSheets()[0];
    if (masterWs) {
      var data = masterWs.getDataRange().getValues();
      var headers = data[2] || data[0];
      var inputCol = -1, outputCol = -1, tpCol = -1;
      
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]).toLowerCase().trim();
        if (h.indexOf('input sheet id') !== -1 || h.indexOf('input link') !== -1 || h.indexOf('sheet id') !== -1 || h.indexOf('master sheet') !== -1 || h.indexOf('input sheet') !== -1) inputCol = c;
        if (h.indexOf('output link') !== -1 || h.indexOf('output sheet') !== -1 || h.indexOf('output excel') !== -1) outputCol = c;
        if (h.indexOf('teaching plan link') !== -1 || h.indexOf('teaching plan') !== -1 || h.indexOf('syllabus') !== -1 || h.indexOf('tp link') !== -1) tpCol = c;
      }
      
      if (inputCol === -1) inputCol = 4;
      if (outputCol === -1) outputCol = 5;
      if (tpCol === -1) tpCol = 6;
      
      for (var r = 3; r < data.length; r++) {
        var row = data[r];
        var rowInputId = String(row[inputCol] || '').trim();
        if (rowInputId === sheetId || (sheetId && rowInputId.indexOf(sheetId) !== -1) || (rowInputId && sheetId.indexOf(rowInputId) !== -1)) {
          var outVal = (outputCol !== -1 && outputCol < row.length) ? String(row[outputCol] || '').trim() : '';
          if (outVal) {
            var m = outVal.match(/\/d\/(.*?)(\/|$)/);
            outputSheetId = m ? m[1] : outVal;
          }
          
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
    Logger.log("_getCollegeSheetIds: Error looking up from master config sheet: " + err.message);
  }
  
  cache.put(cacheKeyOut, outputSheetId || 'NONE', 21600);
  cache.put(cacheKeyTp, teachingPlanId || 'NONE', 21600);

  return { outputSheetId: outputSheetId, teachingPlanId: teachingPlanId };
}

function getTargetSheetIds(code, sheetId) {
  var collegeIds = _getCollegeSheetIds(sheetId);
  var teachingPlanId = collegeIds.teachingPlanId;
  var outputSheetId = collegeIds.outputSheetId;

  if (!teachingPlanId || !outputSheetId) {
    try {
      var ss = _getSpreadsheet(sheetId);
      var ws = ss.getSheetByName('subjects');
      if (ws) {
        var data = ws.getDataRange().getValues();
        var tpColIdx = -1;
        var outColIdx = -1;
        var codeColIdx = 0;

        var headers = data[0] || [];
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

  if (!teachingPlanId) teachingPlanId = sheetId;
  if (!outputSheetId) outputSheetId = getOutputSheetId(sheetId);

  return { teachingPlanId: teachingPlanId, outputSheetId: outputSheetId };
}

function getGlobalTeachingPlanLink(sheetId) {
  if (!sheetId) return '';
  var cache = CacheService.getScriptCache();
  var cacheKey = 'tpLink_' + sheetId;
  var cachedLink = cache.get(cacheKey);
  if (cachedLink !== null) return cachedLink === 'NONE' ? '' : cachedLink;

  try {
    var MASTER_CONFIG_SHEET_ID = "1p3WoC2s-YYqn9ekqkQ72banxAAd-ujlDoFYpv4fkXmk";
    var masterSs = _getSpreadsheet(MASTER_CONFIG_SHEET_ID);
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
        if (tpVal) {
          var m = tpVal.match(/\/d\/(.*?)(\/|$)/);
          var finalLink = m ? m[1] : tpVal;
          cache.put(cacheKey, finalLink, 21600);
          return finalLink;
        }
        break;
      }
    }
  } catch(e) {
    Logger.log("Error getting global teaching plan link: " + e.message);
  }
  cache.put(cacheKey, 'NONE', 21600);
  return '';
}

/* ═══════════════════════════════════════════════════════════════
   SMART ATTENDANCE LOGIC (FULLY RESTORED & UNTOUCHED)
   ═══════════════════════════════════════════════════════════════ */

function saveAttendance(records, outputSheetId, collegeName, managementName, sheetId) {
  if (!records || !records.length) return { error: 'No data' };
  if (!outputSheetId) outputSheetId = getOutputSheetId(sheetId);
  var res = updateOutputMatrix(records, outputSheetId, collegeName, managementName, sheetId);
  if (res === true) {
    try {
      if (sheetId) CacheService.getScriptCache().remove('dash_' + sheetId);
      var code = records[0] && records[0].code ? records[0].code : '';
      var cleanOutId = extractSpreadsheetId(outputSheetId);
      if (code) {
        CacheService.getScriptCache().remove('attrep_v1_' + code + '__' + cleanOutId);
      }
    } catch(cErr) {}
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
            hdrRowIdx = 5;
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
            
            var nameColIdx = -1;
            for (var c = 0; c < hRows.length; c++) {
                if (hRows[c].trim().toLowerCase().indexOf('name') !== -1) {
                    nameColIdx = c;
                    break;
                }
            }
            if (nameColIdx === -1) nameColIdx = 1;
            
            if (dCol === -1 && tpCol !== -1) {
                sheet.insertColumnBefore(tpCol); dCol = tpCol;
                sheet.getRange(hdrRowNumber, dCol).setValue(dispDate).setFontWeight("bold").setBackground("#F1F5F9").setHorizontalAlignment("center");
                sheet.setColumnWidth(dCol, 100);
                var rs = sheet.getLastRow() - (hdrRowNumber + 1);
                if (rs > 0) {
                    var tpL = columnToLetter(dCol+1), taL = columnToLetter(dCol+2), tL = columnToLetter(dCol+3), deL = columnToLetter(dCol);
                    var firstDateLetter = columnToLetter(nameColIdx + 2);
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
  var cleanOutId = extractSpreadsheetId(outputSheetId || getOutputSheetId(sheetId));
  var cacheKey = 'attrep_v1_' + (code || '') + '_' + (year || '') + '_' + (date || '') + '_' + (cleanOutId || '');
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }
  var result = _getAttendanceUncached(code, year, date, outputSheetId, sheetId);
  if (result && result.success) {
    try {
      var jsonStr = JSON.stringify(result);
      if (jsonStr.length < 95000) {
        cache.put(cacheKey, jsonStr, 300);
      }
    } catch(ce) {}
  }
  return result;
}

function _getAttendanceUncached(code, year, date, outputSheetId, sheetId) {
  if (!code) return { error: 'No code' };
  if (!outputSheetId) outputSheetId = getOutputSheetId(sheetId);
  var cleanOutId = extractSpreadsheetId(outputSheetId);
  if (!cleanOutId) return { error: 'Invalid Output Sheet Link' };
  var outSs; try { outSs = SpreadsheetApp.openById(cleanOutId); } catch(e) { return { error: 'Scan Fail' }; }
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
    
    var attData = s.getDataRange().getValues();
    if (!attData || attData.length < 8) continue;

    var hdrRowIdx = -1;
    for (var r = 0; r < Math.min(attData.length, 30); r++) {
      var rowStr = attData[r].map(function(cell) { return String(cell || '').toLowerCase().trim(); }).join('|');
      if (rowStr.indexOf('roll no') !== -1 && rowStr.indexOf('name') !== -1 && (rowStr.indexOf('total p') !== -1 || rowStr.indexOf('% att') !== -1)) {
        hdrRowIdx = r;
        break;
      }
    }
    if (hdrRowIdx === -1) {
      hdrRowIdx = 5;
    }
    if (attData.length <= hdrRowIdx + 2) continue;

    var rawHeaders = attData[hdrRowIdx] || [];
    var hdrs = rawHeaders.map(function(cell) {
      if (cell instanceof Date) {
        try { return Utilities.formatDate(cell, outSs.getSpreadsheetTimeZone(), 'yyyy-MM-dd'); } catch(e) {}
      }
      return String(cell || '').trim();
    });
    
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
    if (totalPColIdx === -1) totalPColIdx = hdrs.length;

    var dates = [];
    var firstDateColIdx = nameColIdx + 1;
    for (var c = firstDateColIdx; c < totalPColIdx; c++) {
       if (hdrs[c]) dates.push({ index: c, disp: hdrs[c] });
    }
    if (dates.length === 0) continue;

    var topicRow = attData[hdrRowIdx + 1] || [];
    for (var r = hdrRowIdx + 2; r < attData.length; r++) {
       var rowData = attData[r];
       if (!rowData || rowData.length === 0) continue;
       for (var d = 0; d < dates.length; d++) {
          var colIdx = dates[d].index;
          if (colIdx >= rowData.length) continue;
          var st = String(rowData[colIdx] || '').trim();
          if (st === 'P' || st === 'A') {
             var dbD = displayToDb(dates[d].disp);
             if (date && dbD.indexOf(date) === -1) continue;
             var topicVal = colIdx < topicRow.length ? String(topicRow[colIdx] || '') : '';
             res.push({
               date: dbD,
               code: code,
               year: year,
               batch: batch,
               faculty: "Assigned",
               rollNo: rowData[0],
               name: rowData[1],
               status: st,
               topic: topicVal
             });
          }
       }
    }
  }
  return { success: true, records: res };
}

function extractSpreadsheetId(url) {
  if (!url) return '';
  var m = String(url).match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m && m[1]) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(url)) return url;
  return '';
}

function getSyllabus(link, code, sheetId) {
  var cacheKey = 'syl_' + (code || '') + '_' + (link ? extractSpreadsheetId(link) : '');
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }
  try {
    var points = [];

    // Tier 1: Try direct link if provided
    if (link) {
      try {
        var p = getSyllabusPointsFromLink(link, code);
        if (p && p.length > 0) points = p;
      } catch(e1) {
        Logger.log("getSyllabus Tier 1 error: " + e1.message);
      }
    }

    // Tier 2: If Tier 1 failed or no link, resolve Teaching Plan spreadsheet ID via getTargetSheetIds
    if ((!points || points.length === 0) && code) {
      try {
        var targetIds = getTargetSheetIds(code, sheetId);
        if (targetIds && targetIds.teachingPlanId && targetIds.teachingPlanId !== link) {
          var p2 = getSyllabusPointsFromLink(targetIds.teachingPlanId, code);
          if (p2 && p2.length > 0) points = p2;
        }
      } catch(e2) {
        Logger.log("getSyllabus Tier 2 error: " + e2.message);
      }
    }

    if (points && points.length > 0) {
      var res = { success: true, points: points };
      try { cache.put(cacheKey, JSON.stringify(res), 7200); } catch(ce) {}
      return res;
    }
    return { success: false, points: [], error: 'No syllabus points found for ' + (code || 'subject') };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function looksLikeSubjectCode(name) {
  if (!name) return false;
  var val = String(name).trim().toLowerCase().replace(/\s+/g, "");
  if (!val) return false;
  if (val.indexOf("sheet") === 0 || val.indexOf("lecture") === 0 || val.indexOf("unit") === 0 || val.indexOf("chap") === 0) return false;
  if (val.indexOf("syllabus") !== -1 || val.indexOf("plan") !== -1 || val.indexOf("attendance") !== -1 || val.indexOf("index") !== -1) return false;
  // Numeric subject codes: 4 to 8 digits (e.g. 22401, 314001)
  if (/^\d{4,8}$/.test(val)) return true;
  // Alphanumeric subject codes (e.g. CS101, 22401P, DME22401)
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
  
  var colIdx = -1;
  var headerRowIdx = -1;
  var keywords = ["syllabus points", "syllabus point", "syllabus", "topic name", "topics", "topic", "session topic", "particulars", "description", "content", "practical topic", "experiment name", "experiments", "experiment", "lab topic", "practical"];
  
  for (var r = 0; r < Math.min(data.length, 30); r++) {
    var row = data[r].map(function(h) { return String(h).trim().toLowerCase(); });
    
    for (var k = 0; k < keywords.length; k++) {
      var idx = row.indexOf(keywords[k]);
      if (idx !== -1) {
        colIdx = idx;
        headerRowIdx = r;
        break;
      }
    }
    if (colIdx !== -1) break;
    
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
var FCM_SERVER_KEY = "AIzaSyBuw7HMI__3oNgMbjQz-q2L1aoIcfn5H9k";

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


function getTaughtTopics(code, outputSheetId, sheetId) {
  try {
    if (outputSheetId === 'undefined' || outputSheetId === 'null') outputSheetId = '';
    if (!outputSheetId && sheetId) outputSheetId = getOutputSheetId(sheetId);
    var cleanOutId = extractSpreadsheetId(outputSheetId);
    if (!cleanOutId && sheetId) cleanOutId = extractSpreadsheetId(getOutputSheetId(sheetId));
    if (!cleanOutId) return { success: false, error: 'Invalid Output Sheet Link' };
    var outSs = SpreadsheetApp.openById(cleanOutId);
    var sheets = outSs.getSheets();
    var parsedInput = _parseSubjectCode(code);
    var found = {};
    for (var i = 0; i < sheets.length; i++) {
      var s = sheets[i];
      var parsedSheet = _parseSubjectCode(s.getName());
      var cleanName = s.getName().toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (parsedSheet.cleanBaseCode !== parsedInput.cleanBaseCode && cleanName.indexOf(parsedInput.cleanBaseCode) !== 0) continue;
      var lc = s.getLastColumn(), lr = s.getLastRow();
      if (lc < 6 || lr < 6) continue;
      var rows = s.getRange(1, 1, Math.min(15, lr), lc).getValues(); // only top rows, never full matrix
      var hdr = -1;
      for (var r = 0; r < rows.length; r++) {
        var rowStr = rows[r].map(function (c) { return String(c || '').toLowerCase().trim(); }).join('|');
        if (rowStr.indexOf('roll no') !== -1 && rowStr.indexOf('name') !== -1 && (rowStr.indexOf('total p') !== -1 || rowStr.indexOf('% att') !== -1)) { hdr = r; break; }
      }
      if (hdr === -1) hdr = 5;
      var dateHdr = rows[hdr] || [], topicRow = rows[hdr + 1] || [];
      var nameCol = -1;
      for (var c = 0; c < dateHdr.length; c++) if (String(dateHdr[c] || '').toLowerCase().indexOf('name') !== -1) { nameCol = c; break; }
      if (nameCol === -1) nameCol = 1;
      for (var c2 = nameCol + 1; c2 < dateHdr.length; c2++) {
        var dateVal = String(dateHdr[c2] || '').trim().toLowerCase();
        if (!dateVal || dateVal.indexOf('total') !== -1 || dateVal.indexOf('%') !== -1) continue; // topic exists only under real date columns
        var tv = String(topicRow[c2] || '').trim();
        if (tv && tv.toLowerCase() !== 'topic') found[tv] = true;
      }
    }
    return { success: true, topics: Object.keys(found) };
  } catch (e) { return { success: false, error: e.message }; }
}

