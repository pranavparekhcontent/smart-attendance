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

  async function _post(action, body) {
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
  async function getAllData() {
    if (navigator.onLine) {
      try {
        const data = await _get('getAllData');
        if (data.success) {
          _setCache('allData', data);
          return data;
        }
      } catch (e) {
        console.warn('API.getAllData network fail:', e.message);
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

  /**
   * Get students for a year sheet
   */
  async function getStudents(sheetName, batch) {
    const cacheKey = 'students_' + sheetName + (batch ? '_' + batch : '');
    if (navigator.onLine) {
      try {
        const params = { sheet: sheetName };
        if (batch) params.batch = batch;
        const data = await _get('getStudents', params);
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
   * Get existing attendance records for session-check
   */
  async function getAttendance(code, year, date, outputSheetId) {
    if (!navigator.onLine) return { success: false, error: 'Offline' };
    try {
      const params = { code, year };
      if (date) params.date = date;
      if (outputSheetId) params.outputSheetId = outputSheetId;
      return await _get('getAttendance', params);
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  /**
   * Save attendance — queue locally then sync
   */
  async function saveAttendance(records, outputSheetId) {
    if (!records || !records.length) return { success: false, error: 'No records' };

    // Always queue first
    _addPending(records);

    if (navigator.onLine) {
      return await syncPending(outputSheetId);
    }
    return { success: true, synced: false, message: 'Saved offline. Will sync when online.' };
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
      });
      if (res.success) {
        _clearPending();
        return { success: true, synced: res.saved };
      }
      return { success: false, error: res.error, pending: pending.length };
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

  // Auto-sync on connectivity restore
  if (typeof window !== 'undefined') {
    window.addEventListener('online', () => {
      syncPending().then(r => {
        if (r.synced > 0 && window.Toast) {
          Toast.show('✅ Synced ' + r.synced + ' offline records', 'success');
        }
      });
    });
  }

  return {
    getAllData,
    getAllDataFromUrl,
    getConfig,
    getStudents,
    getAttendance,
    saveAttendance,
    syncPending,
    getPendingCount,
    extractSheetId
  };
})();
