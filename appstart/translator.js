/**
 * ═══════════════════════════════════════════════════════════════
 *  UNIVERSAL DATA TRANSLATOR (VibeMantra Smart Engine)
 *  Fuzzy header matching and row normalization logic.
 * ═══════════════════════════════════════════════════════════════
 */

export const Translator = (() => {
  
  /**
   * Matches a target concept (e.g. "name") against a list of headers
   * using predefined keyword variations.
   */
  function findHeader(headers, keywords) {
    if (!headers || !Array.isArray(headers)) return null;
    
    const normalizedHeaders = headers.map(h => String(h).toLowerCase().replace(/[^a-z0-9]/g, ''));
    const normalizedKeywords = keywords.map(k => k.toLowerCase().replace(/[^a-z0-9]/g, ''));
    
    // 1. Try exact matches first
    for (let i = 0; i < normalizedHeaders.length; i++) {
      if (normalizedKeywords.includes(normalizedHeaders[i])) return headers[i];
    }
    
    // 2. Try partial matches (header contains keyword)
    for (let i = 0; i < normalizedHeaders.length; i++) {
      for (const kw of normalizedKeywords) {
        if (normalizedHeaders[i].includes(kw)) return headers[i];
      }
    }
    
    return null;
  }

  /**
   * Normalizes a raw row object based on a schema mapping.
   * schemaMap: { targetKey: [keywords...] }
   */
  function normalizeRow(row, schemaMap) {
    const headers = Object.keys(row);
    const result = {};
    
    for (const [key, keywords] of Object.entries(schemaMap)) {
      const match = findHeader(headers, keywords);
      const val = match ? row[match] : null;
      
      // Smart cleanup
      if (typeof val === 'string') {
        result[key] = val.trim();
      } else {
        result[key] = val;
      }
    }
    
    return result;
  }

  /**
   * Heuristic to guess if a sheet matches a specific type.
   * Required keywords list.
   */
  function identifySheet(rows, requiredConcepts) {
    if (!rows || rows.length === 0) return false;
    const headers = Object.keys(rows[0]);
    
    let matches = 0;
    for (const keywords of Object.values(requiredConcepts)) {
      if (findHeader(headers, keywords)) matches++;
    }
    
    // If at least 80% of required concepts are found, it's a match
    return (matches / Object.keys(requiredConcepts).length) >= 0.8;
  }

  /**
   * Converts an array of arrays (e.g. from CSV) to array of objects
   * using the first row as headers.
   */
  function arrayToObjects(rows) {
    if (!rows || rows.length < 2) return [];
    const headers = rows[0];
    return rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      return obj;
    });
  }

  return { findHeader, normalizeRow, identifySheet, arrayToObjects };
})();
