/**
 * Smart Attendance PWA — Google Apps Script API Bridge (v1.7)
 * FIXED: Leading-zero date matching + Fuzzy Column Scanner
 */

function doGet(e) {
  try {
    var action = e.parameter.action;
    var result;
    switch (action) {
      case 'getTeachers': result = getTeachers(); break;
      case 'getSubjects': result = getSubjects(e.parameter.teacher); break;
      case 'getStudents': result = getStudents(e.parameter.sheet, e.parameter.batch); break;
      case 'getAttendanceLimit': result = getAttendanceLimit(); break;
      case 'getAttendance': result = getAttendance(e.parameter.code, e.parameter.year, e.parameter.date, e.parameter.outputSheetId); break;
      case 'getConfig':
      case 'getAllData': result = getAllData(); break;
      default: result = { error: 'Unknown action: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || (e.parameter && e.parameter.action);
    var result;
    switch (action) {
      case 'saveAttendance': result = saveAttendance(data.records, data.outputSheetId, data.collegeName, data.managementName); break;
      default: result = { error: 'Unknown POST action: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getTeachers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName('subjects');
  if (!ws) return { success: false, error: 'Sheet "subjects" not found' };
  var data = ws.getDataRange().getValues(), map = {};
  for (var i = 1; i < data.length; i++) {
    var fStr = String(data[i][6]).trim(), pStr = String(data[i][7]).trim();
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

function getSubjects(teacher) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName('subjects');
  if (!ws) return { success: false };
  var data = ws.getDataRange().getValues(), res = [];
  for (var i = 1; i < data.length; i++) {
    var fs = String(data[i][6]).toLowerCase().split(',').map(function(x){return x.trim()});
    if (fs.indexOf(teacher.toLowerCase()) !== -1) {
      res.push({ code: String(data[i][0]).trim(), name: String(data[i][1]).trim(), year: String(data[i][2]).trim(), program: String(data[i][3]).trim(), semester: String(data[i][4]).trim(), type: String(data[i][5]).trim() });
    }
  }
  return { success: true, subjects: res };
}

function getStudents(sheet, batch) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName(sheet);
  if (!ws) return { success: false };
  var data = ws.getDataRange().getValues(), res = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i][0], n = String(data[i][1]).trim(), b = String(data[i][2]).trim();
    if (!r && !n) continue;
    if (batch && b !== batch) continue;
    res.push({ rollNo: r, name: n, batch: b });
  }
  return { success: true, students: res, sheet: sheet };
}

function getAttendanceLimit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName('subjects');
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

function saveAttendance(records, outputSheetId, collegeName, managementName) {
  if (!records || !records.length) return { error: 'No data' };
  if (!outputSheetId) outputSheetId = getOutputSheetId();
  var res = updateOutputMatrix(records, outputSheetId, collegeName, managementName);
  return res === true ? { success: true, saved: records.length } : { success: false, error: String(res) };
}

function updateOutputMatrix(records, outputSheetId, _collegeName, _managementName) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return "Lock Timeout"; }
  try {
    var outSs = SpreadsheetApp.openById(outputSheetId);
    var grouped = {};
    for (var i = 0; i < records.length; i++) {
      var r = records[i], tab = r.code + " - " + getSubjectName(r.code);
      if (r.batch) tab += " - Batch " + r.batch;
      if (!grouped[tab]) grouped[tab] = {};
      if (!grouped[tab][r.date]) grouped[tab][r.date] = [];
      grouped[tab][r.date].push(r);
    }
    var limit = (getAttendanceLimit().limit || 75) / 100;
    var config = { collegeName: _collegeName || '', managementName: _managementName || '' };
    if (!config.collegeName || !config.managementName) {
      try {
        var cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('client sheet') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('subjects');
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
           var f = dates[dKeys[0]][0], sts = getStudents(f.year, f.batch).students || [];
           sts.sort(function(a,b){return parseInt(a.rollNo)-parseInt(b.rollNo)});
           var sd = sts.map(function(s){return [s.rollNo, s.name, 0, 0, 0, 0]});
           if (sd.length > 0) sheet.getRange(8, 1, sd.length, 6).setValues(sd);
           sheet.setColumnWidth(1, 80); sheet.setColumnWidth(2, 280);
       }
       for (var k = 0; k < dKeys.length; k++) {
           var dateKey = dKeys[k], recs = dates[dateKey], dispDate = dbToDisplay(dateKey);
           var hRows = sheet.getRange(6, 1, 1, Math.max(sheet.getLastColumn(), 10)).getDisplayValues()[0];
           var dCol = -1, tpCol = -1;
           for (var c = 0; c < hRows.length; c++) {
               var val = hRows[c].trim().toLowerCase();
               if (val === dispDate.toLowerCase()) dCol = c + 1;
               if (val.indexOf("total p") !== -1) tpCol = c + 1;
           }
           if (dCol === -1 && tpCol !== -1) {
               sheet.insertColumnBefore(tpCol); dCol = tpCol;
               sheet.getRange(6, dCol).setValue(dispDate).setFontWeight("bold").setBackground("#F1F5F9").setHorizontalAlignment("center");
               sheet.setColumnWidth(dCol, 100);
               var rs = sheet.getLastRow() - 7;
               if (rs > 0) {
                   var tpL = columnToLetter(dCol+1), taL = columnToLetter(dCol+2), tL = columnToLetter(dCol+3), deL = columnToLetter(dCol);
                   var fms = [];
                   for (var r=0; r<rs; r++) {
                       var rn = r+8;
                       fms.push(['=COUNTIF(C'+rn+':'+deL+rn+', "P")', '=COUNTIF(C'+rn+':'+deL+rn+', "A")', '='+tpL+rn+'+'+taL+rn, '=IF('+tL+rn+'>0,'+tpL+rn+'/'+tL+rn+',0)']);
                   }
                   sheet.getRange(8, dCol+1, rs, 4).setFormulas(fms);
                   setupFormulasAndConditions(sheet, rs, dCol+4, 8, limit*100);
               }
           }
           if (dCol !== -1) {
               var topic = recs[0] && recs[0].topic ? recs[0].topic : "";
               sheet.getRange(7, dCol).setValue(topic).setFontStyle("italic").setHorizontalAlignment("center");
               
               var rs = sheet.getLastRow() - 7;
               if (rs > 0) {
                   var ex = sheet.getRange(8, dCol, rs, 1).getValues(), rolls = sheet.getRange(8, 1, rs, 1).getValues();
                   var ups = rolls.map(function(r, idx) {
                       var roll = String(r[0]), st = ex[idx][0] || "-";
                       for (var x=0; x<recs.length; x++) { if (String(recs[x].rollNo) === roll) st = recs[x].status; }
                       return [st];
                   });
                   sheet.getRange(8, dCol, rs, 1).setValues(ups).setHorizontalAlignment("center");
               }
           }
       }
       try {
         var f = dates[dKeys[0]][0], info = getSubjectInfo(f.code);
         var row4 = f.code + " - " + info.name + (f.batch ? " | Batch " + f.batch : "") + " | " + info.program + " | " + info.year + " | 01 Jan 2020 to " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy");
         sheet.getRange("A1:K1").unmerge().mergeAcross().setValue(config.managementName || "Management Name").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
         sheet.getRange("A2:K2").unmerge().mergeAcross().setValue(config.collegeName || "College Name").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
         sheet.getRange("A4:K4").unmerge().mergeAcross().setValue(row4).setFontWeight("bold").setBackground("#E2E8F0").setHorizontalAlignment("center");
       } catch(e) {}
    }
    return true;
  } catch(e) { return e.message; } finally { lock.releaseLock(); }
}

function getSubjectName(code) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName('subjects');
  if (!ws) return "Unknown";
  var data = ws.getDataRange().getValues();
  for (var i=1; i<data.length; i++) { if (String(data[i][0]).trim() === String(code).trim()) return String(data[i][1]).trim(); }
  return "Unknown";
}

function getSubjectInfo(code) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName('subjects');
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

function setupFormulasAndConditions(sheet, rows, pctCol, startRow, limit) {
  sheet.getRange(startRow, pctCol, rows, 1).setNumberFormat('0.0%');
  var dataR = sheet.getRange(startRow, 3, 1000, Math.max(pctCol - 3, 1));
  var pctR = sheet.getRange(startRow, pctCol, 1000, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("P").setFontColor("#15803D").setBackground("#DCFCE7").setRanges([dataR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("A").setFontColor("#B91C1C").setBackground("#FEE2E2").setRanges([dataR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(limit / 100).setFontColor("#14532D").setBackground("#BBF7D0").setBold(true).setRanges([pctR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(limit / 100).setFontColor("#7F1D1D").setBackground("#FECACA").setBold(true).setRanges([pctR]).build()
  ]);
}

function getAttendance(code, year, date, outputSheetId) {
  if (!code) return { error: 'No code' };
  if (!outputSheetId) outputSheetId = getOutputSheetId();
  var outSs; try { outSs = SpreadsheetApp.openById(outputSheetId); } catch(e) { return { error: 'Scan Fail' }; }
  var res = [], sheets = outSs.getSheets(), pre = code + " -";
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i], name = s.getName();
    if (name.indexOf(pre) !== 0) continue;
    var batch = name.indexOf(" - Batch ") !== -1 ? name.substring(name.indexOf(" - Batch ") + 9).trim() : "";
    var lc = s.getLastColumn(), lr = s.getLastRow();
    if (lc < 6 || lr < 8) continue;
    var hdrs = s.getRange(6, 1, 1, lc).getDisplayValues()[0];
    var raw = s.getRange(6, 1, 1, lc).getValues()[0];
    var dates = [];
    for (var c = 2; c < hdrs.length; c++) {
       if (String(raw[c]).toLowerCase().indexOf("total p") !== -1) break;
       if (hdrs[c].trim()) dates.push({ index: c, disp: hdrs[c].trim() });
    }
    if (dates.length === 0) continue;
    var topics = s.getRange(7, 1, 1, lc).getValues()[0];
    var mtx = s.getRange(8, 1, lr - 7, lc).getValues();
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

function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName('subjects'), subs = [], config = { collegeName: '', managementName: '' };
  if (ws) {
    var data = ws.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim()) subs.push({ code: String(data[i][0]).trim(), name: String(data[i][1]).trim(), year: String(data[i][2]).trim(), program: String(data[i][3]).trim(), semester: String(data[i][4]).trim(), type: String(data[i][5]).trim(), faculty: String(data[i][6]).trim() }); }
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
  return { success: !!ws, teachers: getTeachers().teachers || [], subjects: subs, attendanceLimit: getAttendanceLimit().limit || 75, config: config };
}

function getOutputSheetId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ws = ss.getSheetByName('subjects');
  var data = ws ? ws.getDataRange().getValues() : [];
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (String(data[i][j]).trim().toLowerCase() === 'output excel link') {
         var f = '';
         for (var n = j + 1; n < data[i].length; n++) { var nv = String(data[i][n]).trim(); if (nv !== '' && ['link','name','text'].indexOf(nv.toLowerCase()) === -1) { f = nv; break; } }
         if (f === '' && i + 1 < data.length) f = String(data[i+1][j]).trim();
         if (f) { var m = f.match(/\/d\/(.*?)(\/|$)/); if (m && m[1]) return m[1]; }
      }
    }
  }
  return '';
}
