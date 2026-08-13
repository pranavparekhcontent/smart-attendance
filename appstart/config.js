// ============================================================
//  APPSTART CONFIG — Edit this file for each new app.
//  All other appstart/ files remain untouched between projects.
// ============================================================

const APP_CONFIG = {

  // ── App Identity ──────────────────────────────────────────
  APP_NAME: "Smart Attendance",
  APP_VERSION: "1.0.51",   // Fallback only. Auto-synced from version.json at runtime.

  // ── Layout ───────────────────────────────────────────────
  LAYOUT: "mobile-first",

  // ── Theme (Cyberpunk Dark — matches existing RMDIPER aesthetic) ──
  THEME: {
    primary: "#3B82F6",   // Neon blue accent
    secondary: "#1E40AF",   // Deep blue
    danger: "#EF4444",   // Red for absent / errors
    bg: "#0D0F14",   // Ultra-dark background
    surface: "#1A1D27",   // Card surfaces
    border: "#2A2D3A",   // Subtle borders
    text: "#E8EAF6",   // Light text
    muted: "#6B7280",   // Muted labels
  },

  // ── License ───────────────────────────────────────────────
  LICENSE_STORAGE_KEY: "attendance_license",

  // ── Central API Configuration ──────────────────────────────
  // Change this to your deployed Google Apps Script Web App URL
  CENTRAL_API_URL: "https://script.google.com/macros/s/AKfycbwdEMH_36ryLox45JmzdI6v8z7J0AEgk5gtFHwmy87V5aJhlpxAovaz6UNHdrOp8pH-/exec",

  // ── Config Sheet ──────────────────────────────────────────
  // MASTER CONFIG SHEET (Common for all apps)
  // SMART DETECTION — column order doesn't matter!
  CONFIG_SHEET_URL:
    "https://docs.google.com/spreadsheets/d/1p3WoC2s-YYqn9ekqkQ72banxAAd-ujlDoFYpv4fkXmk/gviz/tq?tqx=out:json",

  dataFetcher: async (serverUrl, sheetId = "") => {
    // Sanitize: Remove trailing slashes and any accidental query params from the sheet string
    const cleanUrl = serverUrl.replace(/\/+$/, "").replace(/\?.*$/, "");

    let targetUrl = cleanUrl + '?action=getAllData';
    if (sheetId) {
      targetUrl += '&sheetId=' + encodeURIComponent(sheetId);
    }

    return {
      allData: (async () => {
        // 1. If offline, return localStorage cache immediately
        if (!navigator.onLine) {
          try {
            const raw = localStorage.getItem('attendance_cache_allData');
            if (raw) {
              const parsed = JSON.parse(raw);
              if (parsed && parsed.data && (parsed.data.teachers || parsed.data.subjects)) {
                return parsed.data;
              }
            }
          } catch(e) {}
        }

        // 2. Fetch fresh data over network with 15s timeout
        try {
          const res = await Promise.race([
            fetch(targetUrl),
            new Promise((_, reject) => setTimeout(() => reject(new Error('Data fetch timeout (15s)')), 15000))
          ]);
          if (!res.ok) throw new Error('HTTP ' + res.status);
          const json = await res.json();
          if (json && (json.success || json.teachers || json.subjects)) {
            try {
              localStorage.setItem('attendance_cache_allData', JSON.stringify({ ts: Date.now(), data: json }));
            } catch(e) {}
            return json;
          }
          throw new Error((json && json.error) || 'Empty data received');
        } catch (err) {
          console.warn("AppStart Data Fetcher network issue, checking local cache...", err.message);
          // 3. Network failed: use local cache if valid
          try {
            const raw = localStorage.getItem('attendance_cache_allData');
            if (raw) {
              const parsed = JSON.parse(raw);
              if (parsed && parsed.data && ((parsed.data.teachers && parsed.data.teachers.length) || (parsed.data.subjects && parsed.data.subjects.length))) {
                return parsed.data;
              }
            }
          } catch(e) {}
          // 4. No cache and network failed: THROW so AppStart shows Error screen with Retry button (never a blank card!)
          throw new Error(err.message || 'Unable to load college roster. Please check internet and tap Retry.');
        }
      })()
    };
  },

  /** CALLBACKS */
  onComplete: (context) => {
    console.log("AppStart complete for:", context.collegeName);
  }
};
