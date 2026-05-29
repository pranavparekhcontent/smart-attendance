/**
 * ═══════════════════════════════════════════════════════════════
 *  SCHEMA LEARNER (VibeMantra Smart Engine)
 *  Analyzes all spreadsheet sheets and learns the app structure.
 * ═══════════════════════════════════════════════════════════════
 */

import { Translator } from './translator.js';

export const SchemaLearner = (() => {
  
  // Standard Concepts for common PWA types
  const CONCEPTS = {
    STUDENT_RECORD: {
      id:   ['roll no', 'id', 'roll', 'no', 'student id'],
      name: ['student name', 'name', 'full name', 'fullname'],
      batch:['batch', 'section', 'group'],
      pin:  ['pin', 'password', 'student pin']
    },
    SUBJECT_RECORD: {
      name:     ['subject name', 'subject', 'name', 'title'],
      code:     ['subject code', 'code', 'id'],
      type:     ['type', 'category', 'mode'],
      program:  ['program', 'class', 'course'],
      year:     ['year', 'academic year'],
      semester: ['semester', 'sem']
    }
  };

  /**
   * Main learning engine. Returns a structured DB.
   */
  function learn(rawSheets) {
    const db = {
      subjects: [],
      students: {}, // sheetName -> [students]
      config:   {},
      meta:     {
        discoveredAt: new Date().toISOString(),
        sheetCount: Object.keys(rawSheets).length
      }
    };

    Object.entries(rawSheets).forEach(([sheetName, rawRows]) => {
      if (!Array.isArray(rawRows) || rawRows.length === 0) return;

      // Ensure rows are objects. If it's an array of arrays, convert it.
      let rows = rawRows;
      if (Array.isArray(rawRows[0])) {
        rows = Translator.arrayToObjects(rawRows);
      }

      if (rows.length === 0) return;

      // 1. Is this the Subject sheet?
      if (Translator.identifySheet(rows, CONCEPTS.SUBJECT_RECORD)) {
        db.subjects = rows.map(r => Translator.normalizeRow(r, CONCEPTS.SUBJECT_RECORD));
      } 
      // 2. Is this a Student sheet?
      else if (Translator.identifySheet(rows, CONCEPTS.STUDENT_RECORD)) {
        db.students[sheetName] = rows.map(r => Translator.normalizeRow(r, CONCEPTS.STUDENT_RECORD));
      }
      // 3. Otherwise, store as generic raw data
      else {
        db.config[sheetName] = rows;
      }
    });

    return db;
  }

  return { learn };
})();
