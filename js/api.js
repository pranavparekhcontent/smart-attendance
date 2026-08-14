/**
 * Smart Attendance PWA — Google Sheets API Module
 * Handles all communication with Google Apps Script + offline queue
 */

const API = (() => {
  const PENDING_KEY = 'attendance_pending_sync';
  const CACHE_PREFIX = 'attendance_cache_';

  // Retry config (previously from CONFIG object, now inlined)
  const MAX_RETRIES = 3;
  const RETRY_DELAY_MS = 2000;

  // Get server URL from AppStart context (set by engine after boot)
  function _getBaseUrl() {
    return (window.appStartContext && window.appStartContext.serverUrl) || '';
  }

  // ─── Internal helpers ───

  async function _get(action, params = {}) {
    // Inject sheetId if available in context
    if (window.appStartContext && window.appStartContext.sheetId) {
      params.sheetId = window.appStartContext.sheetId;
    }
    let url = _getBaseUrl() + '?action=' + encodeURIComponent(action);
    for (const k in params) {
      url += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    }

    let lastErr;
    for (let i = 0; i < MAX_RETRIES; i++) {
      try {
        const res = await fetch(url, {
          method: 'GET',
          redirect: 'follow'
        });
        
        if (!res.ok) throw new Error('HTTP ' + res.status);
        
        let data = await res.json();
        
        // --- TRANSLATOR / NORMALIZER ---
        // If the server returned { success, data: { ... } }, flatten it.
        if (data && data.success && data.data && typeof data.data === 'object') {
          const innerData = data.data;
          data = { ...data, ...innerData };
        }
        // If success flag is missing but we have records/students/teachers, assume success
        if (data && !data.success && (data.records || data.students || data.teachers)) {
          data.success = true;
        }
        
        return data;
      } catch (err) {
        lastErr = err;
        if (i < MAX_RETRIES - 1) {
          await _sleep(RETRY_DELAY_MS * (i + 1));
        }
      }
    }
    throw lastErr;
  }

  async function _post(action, body, timeoutMs = 10000) {
    return _withTimeout((async () => {
      let url = _getBaseUrl() + '?action=' + encodeURIComponent(action);
      // Inject sheetId into POST body if available
      if (window.appStartContext && window.appStartContext.sheetId) {
        if (typeof body === 'object' && body !== null) {
          body.sheetId = window.appStartContext.sheetId;
        }
      }

      let lastErr;
      for (let i = 0; i < MAX_RETRIES; i++) {
        try {
          const res = await fetch(url, {
            method: 'POST',
            body: JSON.stringify(body),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            redirect: 'follow'
          });
          
          if (!res.ok) throw new Error('HTTP ' + res.status);
          
          let data = await res.json();
          
          // --- TRANSLATOR / NORMALIZER ---
          if (data && data.success && data.data && typeof data.data === 'object') {
            const innerData = data.data;
            data = { ...data, ...innerData };
          }
          
          return data;
        } catch (err) {
          lastErr = err;
          if (i < MAX_RETRIES - 1) {
            await _sleep(RETRY_DELAY_MS * (i + 1));
          }
        }
      }
      throw lastErr;
    })(), timeoutMs);
  }

  function _sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  // ─── Cache ───

  function _setCache(key, data) {
    try {
      localStorage.setItem(CACHE_PREFIX + key, JSON.stringify({ ts: Date.now(), data }));
    } catch (e) { /* quota exceeded — silent */ }
  }

  function _getCache(key) {
    try {
      const raw = localStorage.getItem(CACHE_PREFIX + key);
      if (!raw) return null;
      return JSON.parse(raw).data;
    } catch { return null; }
  }

  // ─── Pending queue ───

  function _getPending() {
    try {
      const raw = localStorage.getItem(PENDING_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch { return []; }
  }

  function _setPending(arr) {
    localStorage.setItem(PENDING_KEY, JSON.stringify(arr));
  }

  function _clearPending() {
    localStorage.removeItem(PENDING_KEY);
  }

  function _addPending(records) {
    _setPending(_getPending().concat(records));
  }

  // ─── Public API ───

  /**
   * Fetch all boot data: teachers, subjects, attendance limit, config
   */
  /**
   * Fetch all boot data: teachers, subjects, attendance limit, config
   * Fetches fresh data ONCE per session on launch, then reuses session cache
   */
  async function getAllData() {
    try {
      const isSessionFetched = sessionStorage.getItem('session_data_fetched');
      if (isSessionFetched === 'true') {
        const sessionCached = _getCache('allData');
        if (sessionCached) return sessionCached;
      }
    } catch(e) {}

    if (navigator.onLine) {
      try {
        const data = await _withTimeout(_get('getAllData'), 25000);
        if (data && (data.success || data.teachers)) {
          try { sessionStorage.setItem('session_data_fetched', 'true'); } catch(e) {}
          _setCache('allData', data);
          return data;
        }
      } catch (e) {
        console.warn('API.getAllData network fail/timeout:', e.message);
      }
    }
    const cached = _getCache('allData');
    if (cached) return cached;
    return { success: false, error: 'No data available. Check internet.' };
  }

  /**
   * Fallback: Fetch all data from a specific URL (used when engine provides serverUrl)
   */
  async function getAllDataFromUrl(serverUrl) {
    if (!serverUrl) return { success: false, error: 'No server URL provided' };
    
    try {
      const isSessionFetched = sessionStorage.getItem('session_data_fetched');
      if (isSessionFetched === 'true') {
        const sessionCached = _getCache('allData');
        if (sessionCached) return sessionCached;
      }
    } catch(e) {}

    // Ensure URL doesn't have trailing slash for consistency
    const baseUrl = serverUrl.replace(/\/$/, "");
    let targetUrl = `${baseUrl}?action=getAllData`;
    
    // Inject sheetId if available in context
    if (window.appStartContext && window.appStartContext.sheetId) {
      targetUrl += '&sheetId=' + encodeURIComponent(window.appStartContext.sheetId);
    }
    
    console.log("🌐 API: Fetching from", targetUrl);

    try {
      const res = await fetch(targetUrl, {
        method: 'GET',
        redirect: 'follow'
      });
      
      if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);
      
      const data = await res.json();
      if (data && (data.success || data.teachers)) {
        try { sessionStorage.setItem('session_data_fetched', 'true'); } catch(e) {}
        _setCache('allData', data);
        return data;
      } else {
        throw new Error(data.error || 'Invalid JSON structure from GAS');
      }
    } catch (e) {
      console.warn('⚠️ API.getAllDataFromUrl fail:', e.message);
      const cached = _getCache('allData');
      if (cached) {
        console.log("📂 Using cached data as emergency fallback");
        return cached;
      }
      return { success: false, error: e.message || 'Network Error' };
    }
  }

  /**
   * Get config values from the subjects sheet (attendance limit, college name, links, etc.)
   */
  async function getConfig() {
    if (navigator.onLine) {
      try {
        const data = await _get('getConfig');
        if (data.success) {
          _setCache('config', data);
          return data;
        }
      } catch (e) {
        console.warn('API.getConfig network fail:', e.message);
      }
    }
    const cached = _getCache('config');
    if (cached) return cached;
    return { success: false };
  }

  function _withTimeout(promise, ms = 25000) {
    return Promise.race([
      promise,
      new Promise((_, reject) => setTimeout(() => reject(new Error('Request timeout (' + ms + 'ms)')), ms))
    ]);
  }

  /**
   * Fetch syllabus points from an external sheet link
   */
  async function getSyllabusPoints(link, code) {
    const cacheKey = 'syl_' + (code || '') + '_' + (link || '');
    const cached = _getCache(cacheKey);
    if (cached && cached.points && cached.points.length > 0) {
      return cached;
    }

    if (navigator.onLine) {
      try {
        const params = {};
        if (link) params.link = link;
        if (code) params.code = code;
        const data = await _withTimeout(_get('getSyllabus', params), 25000);
        if (data && data.success && data.points && data.points.length > 0) {
          _setCache(cacheKey, data);
        }
        return data;
      } catch (e) {
        console.warn('API.getSyllabusPoints network fail:', e.message);
      }
    }
    if (cached) return cached;
    return { success: false, error: 'Offline or timeout' };
  }

  /**
   * Get students for a year sheet
   */
  /**
   * Get students for a year sheet
   */
  async function getStudents(sheetName, batch) {
    const cacheKey = 'students_' + sheetName + (batch ? '_' + batch : '');
    if (navigator.onLine) {
      try {
        const params = { sheet: sheetName };
        if (batch) params.batch = batch;
        const data = await _withTimeout(_get('getStudents', params), 30000);
        if (data.success) {
          _setCache(cacheKey, data);
          return data;
        }
      } catch (e) {
        console.warn('API.getStudents fail:', e.message);
      }
    }
    const cached = _getCache(cacheKey);
    if (cached) return cached;
    return { success: false, error: 'Cannot fetch students offline.' };
  }

  /**
   * Fast scan: fetch taught topic names only (without downloading full student matrix)
   */
  async function getTaughtTopics(code, outputSheetId) {
    const cleanCode = String(code || '').trim();
    const cacheKey = 'taught_' + cleanCode + '_' + String(outputSheetId || '').trim();
    const cached = _getCache(cacheKey);
    if (cached && cached.topics) {
      return cached;
    }
    if (navigator.onLine) {
      try {
        const params = { code: cleanCode };
        if (outputSheetId) params.outputSheetId = outputSheetId;
        const res = await _withTimeout(_get('getTaughtTopics', params), 15000);
        if (res && res.success && res.topics) {
          _setCache(cacheKey, res);
        }
        return res;
      } catch (e) {
        console.warn('API.getTaughtTopics network fail:', e.message);
      }
    }
    if (cached) return cached;
    return { success: false, topics: [] };
  }

  /**
   * Get existing attendance records for session-check
   */
  async function getAttendance(code, year, date, outputSheetId) {
    const cacheKey = 'att_' + (code || '') + '_' + (year || '') + '_' + (date || '') + '_' + (outputSheetId || '');
    if (!navigator.onLine) {
      const cached = _getCache(cacheKey);
      if (cached) return cached;
      return { success: false, error: 'Offline' };
    }
    try {
      const params = { code, year };
      if (date) params.date = date;
      if (outputSheetId) params.outputSheetId = outputSheetId;
      const res = await _withTimeout(_get('getAttendance', params), 45000);
      if (res && res.success) {
        _setCache(cacheKey, res);
      }
      return res;
    } catch (e) {
      const cached = _getCache(cacheKey);
      if (cached) return cached;
      return { success: false, error: e.message };
    }
  }

  /**
   * Save attendance — queue locally first then attempt online sync
   */
  async function saveAttendance(records, outputSheetId) {
    if (!records || !records.length) return { success: false, error: 'No records' };

    // Always queue locally first for safety
    _addPending(records);

    if (navigator.onLine) {
      const syncRes = await syncPending(outputSheetId);
      if (syncRes.success) {
        return { success: true, synced: true, saved: syncRes.synced, isLocallySaved: false };
      }
    }
    // Saved locally safely (offline or network timeout)
    return {
      success: true,
      synced: false,
      saved: records.length,
      isLocallySaved: true,
      message: '(Saved Locally)'
    };
  }

  /**
   * Sync all pending records
   */
  async function syncPending(outputSheetId) {
    const pending = _getPending();
    if (!pending.length) return { success: true, synced: 0 };
    if (!navigator.onLine) return { success: false, pending: pending.length };

    try {
      const res = await _post('saveAttendance', {
        records: pending,
        outputSheetId: outputSheetId || '',
        collegeName: (window.appStartContext && window.appStartContext.collegeName) || '',
        managementName: (window.appStartContext && window.appStartContext.managementName) || ''
      }, 10000);

      if (res && res.success) {
        _clearPending();
        return { success: true, synced: res.saved || pending.length };
      }
      return { success: false, error: (res && res.error) || 'Sync Error', pending: pending.length };
    } catch (e) {
      return { success: false, error: e.message, pending: pending.length };
    }
  }

  function getPendingCount() {
    return _getPending().length;
  }

  /**
   * Helper: extract spreadsheet ID from a Google Sheets URL
   */
  function extractSheetId(url) {
    if (!url) return '';
    if (url.length < 30 && !url.includes('/')) return url; // Already an ID
    const match = url.match(/\/d\/(.*?)(\/|$)/);
    return match ? match[1] : url;
  }

  // Automatic bootup sync & background retry loop
  if (typeof window !== 'undefined') {
    // 1. Bootup auto-sync (runs 3 seconds after page load)
    setTimeout(() => {
      syncPending().then(r => {
        if (r && r.synced > 0 && window.Toast) {
          Toast.show('✅ Auto-synced ' + r.synced + ' offline records', 'success');
        }
      }).catch(() => {});
    }, 3000);

    // 2. Periodic background sync loop every 30 seconds
    setInterval(() => {
      if (navigator.onLine && _getPending().length > 0) {
        syncPending().then(r => {
          if (r && r.synced > 0 && window.Toast) {
            Toast.show('✅ Background synced ' + r.synced + ' records to Sheets', 'success');
          }
        }).catch(() => {});
      }
    }, 30000);

    // 3. Online event auto-sync
    window.addEventListener('online', () => {
      syncPending().then(r => {
        if (r && r.synced > 0 && window.Toast) {
          Toast.show('✅ Synced ' + r.synced + ' offline records', 'success');
        }
      }).catch(() => {});
    });
  }

  return {
    getAllData,
    getAllDataFromUrl,
    getConfig,
    getSyllabusPoints,
    getTaughtTopics,
    getStudents,
    getAttendance,
    saveAttendance,
    syncPending,
    getPendingCount,
    extractSheetId,
    _getCache,
    _setCache
  };
})();
