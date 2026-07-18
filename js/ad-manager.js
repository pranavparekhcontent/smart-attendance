/**
 * Smart Attendance PWA - Ad Manager Module
 * Centralized control for Monetag MultiTag script injection and direct-link popups.
 * Includes dynamic event listener interception and DOM cleanup to prevent app freezing.
 */
const AdManager = (() => {
  // ─── CONSTANTS ────────────────────────────────────────────────────────────
  const MONETAG_ZONE = '260367';
  const MONETAG_SRC = 'https://quge5.com/88/tag.min.js';
  const DIRECT_LINK_URL = 'https://omg10.com/4/11324927';

  // ─── PLACEMENT CONFIGURATION ──────────────────────────────────────────────
  // Maps screen IDs to script loading behaviors.
  // Slide 1: Push ads/MultiTag are loaded on the login screen.
  // Slide 2: No ads while taking attendance (attendance-mode screen).
  let _config = {
    'login': { loadMonetag: true, label: 'Login Screen' },
    'faculty-dash': { loadMonetag: false, label: 'Faculty Dashboard' },
    'attendance-mode': { loadMonetag: false, label: 'Attendance Entry' },
    'reports': { loadMonetag: false, label: 'Reports Screen' },
    'install-guide': { loadMonetag: false, label: 'Install Guide' },
  };

  // ─── STATE ────────────────────────────────────────────────────────────────
  const _state = {
    monetagLoaded: false,
    currentScreen: '',
    isInjecting: false, // Flag active during script injection window
  };

  // Store for intercepted event listeners to allow clean teardown
  const _interceptedListeners = [];

  // ─── EVENT LISTENER INTERCEPTION (MONKEY-PATCHING) ────────────────────────
  // Captures and registers event listeners added by Monetag/quge5 to allow removal.
  const originalAddEventListener = EventTarget.prototype.addEventListener;
  const originalRemoveEventListener = EventTarget.prototype.removeEventListener;

  EventTarget.prototype.addEventListener = function(type, listener, options) {
    const stack = new Error().stack || '';
    // Identify if the listener is registered by Monetag/quge5 or during the active injection window
    const isAdListener = stack.includes('tag.min.js') || 
                         stack.includes('quge5') || 
                         stack.includes('monetag') ||
                         (_state.isInjecting && !stack.includes('app.js') && !stack.includes('appstart.js'));

    if (isAdListener) {
      _interceptedListeners.push({ target: this, type, listener, options });
    }
    return originalAddEventListener.call(this, type, listener, options);
  };

  // ─── PRIVATE METHODS ──────────────────────────────────────────────────────

  /**
   * Cleans up all event listeners registered by the ad script.
   */
  function _cleanupListeners() {
    let count = 0;
    while (_interceptedListeners.length > 0) {
      const { target, type, listener, options } = _interceptedListeners.shift();
      try {
        originalRemoveEventListener.call(target, type, listener, options);
        count++;
      } catch (e) {
        console.warn("[AdManager] Failed to remove event listener:", e);
      }
    }
    if (count > 0) {
      console.log(`[AdManager] Cleaned up ${count} event listeners registered by ads.`);
    }
  }

  /**
   * Session load limit check for Monetag script (Slide 1: Max 2 times per launch).
   */
  function _isUnderSessionLimit() {
    try {
      const sessionCount = parseInt(sessionStorage.getItem('sa_session_monetag_count') || '0', 10);
      return sessionCount < 2;
    } catch (e) {
      return true;
    }
  }

  /**
   * Increments the Monetag session load count.
   */
  function _incrementSessionCount() {
    try {
      const sessionCount = parseInt(sessionStorage.getItem('sa_session_monetag_count') || '0', 10);
      sessionStorage.setItem('sa_session_monetag_count', (sessionCount + 1).toString());
      console.log(`[AdManager] Monetag session load count: ${sessionCount + 1}/2`);
    } catch (e) {}
  }

  /**
   * Injects the Monetag script tag into the head.
   */
  function _injectMonetagScript() {
    if (document.getElementById('monetag-ad-script')) return;

    if (!_isUnderSessionLimit()) {
      console.log("[AdManager] Monetag injection skipped (reached session limit of 2 loads).");
      return;
    }

    _state.isInjecting = true;
    const script = document.createElement('script');
    script.id = 'monetag-ad-script';
    script.src = MONETAG_SRC;
    script.setAttribute('data-zone', MONETAG_ZONE);
    script.async = true;
    script.setAttribute('data-cfasync', 'false');

    script.onload = () => {
      // Keep isInjecting true briefly to catch post-load event handlers
      setTimeout(() => { _state.isInjecting = false; }, 2000);
      _state.monetagLoaded = true;
      _incrementSessionCount();
      console.log("[AdManager] Monetag script loaded successfully.");
    };

    script.onerror = () => {
      _state.isInjecting = false;
      console.warn("[AdManager] Monetag script failed to load (offline or adblocker).");
    };

    document.head.appendChild(script);
  }

  /**
   * Removes Monetag script and removes its generated iframes/divs/overlays.
   */
  function _removeMonetagElements() {
    const script = document.getElementById('monetag-ad-script');
    if (script) {
      script.remove();
      _state.monetagLoaded = false;
      console.log("[AdManager] Monetag script tag removed.");
    }

    // Clean up iframes, overlays, and other elements generated by Monetag
    const adSelectors = [
      'iframe[src*="quge5"]',
      'iframe[src*="monetag"]',
      'div[id*="monetag"]',
      'iframe[id*="monetag"]',
      'div[class*="monetag"]',
      '.p-notification-container', // Common push notification containers
      '#monetag-ad-script'
    ];
    
    let removedCount = 0;
    adSelectors.forEach(selector => {
      try {
        const elements = document.querySelectorAll(selector);
        elements.forEach(el => {
          el.remove();
          removedCount++;
        });
      } catch (e) {}
    });

    if (removedCount > 0) {
      console.log(`[AdManager] Removed ${removedCount} ad DOM elements.`);
    }
  }

  // ─── PUBLIC API ───────────────────────────────────────────────────────────
  return {
    /**
     * Loads ads for the given screen if allowed by config and session limits.
     * @param {string} screenId
     */
    load(screenId) {
      _state.currentScreen = screenId;
      const screenConfig = _config[screenId];

      if (!screenConfig || !screenConfig.loadMonetag) {
        // No ads configured for this screen, perform cleanup
        this.remove();
        return;
      }

      console.log(`[AdManager] Evaluating ad load for screen: ${screenId} (${screenConfig.label})`);
      _injectMonetagScript();
    },

    /**
     * Removes ads and cleans up event listeners and DOM elements.
     */
    remove() {
      _state.isInjecting = false;
      _removeMonetagElements();
      _cleanupListeners();
    },

    /**
     * Fires a direct link popunder pop-up programmatically.
     * Triggers the primary fallback action first, then handles the ad pop-up.
     * This avoids event hijacking and resolves double-click issues.
     * @param {string} url - Direct Link Ad URL
     * @param {Function} [primaryAction] - Primary action to execute first
     */
    fireDirectLink(url, primaryAction) {
      const adUrl = url || DIRECT_LINK_URL;

      console.log("[AdManager] Executing primary action first...");
      try {
        if (typeof primaryAction === 'function') {
          primaryAction();
        }
      } catch (e) {
        console.error("[AdManager] Primary action failed:", e);
      }

      console.log("[AdManager] Opening direct-link popup ad:", adUrl);
      try {
        window.open(adUrl, '_blank', 'noopener,noreferrer');
      } catch (e) {
        console.warn("[AdManager] Direct link open blocked by browser:", e);
      }
    },

    /**
     * Fires an ad and navigates the window to a target URL.
     * To prevent popup blocking, window.open is called first, then navigation.
     * @param {string} url - Direct Link Ad URL
     * @param {string} targetUrl - Navigation destination (e.g. WhatsApp link)
     */
    fireDirectLinkAndNavigate(url, targetUrl) {
      const adUrl = url || DIRECT_LINK_URL;
      console.log("[AdManager] Opening direct link and navigating to:", targetUrl);
      try {
        window.open(adUrl, '_blank', 'noopener,noreferrer');
      } catch (e) {
        console.warn("[AdManager] Direct link open blocked:", e);
      }
      window.location.href = targetUrl;
    },

    /**
     * Gets current ad manager state for debugging.
     */
    getState() {
      return {
        ..._state,
        sessionCount: parseInt(sessionStorage.getItem('sa_session_monetag_count') || '0', 10),
        interceptedListenersCount: _interceptedListeners.length
      };
    },

    /**
     * Updates placement config dynamically.
     */
    setConfig(config) {
      _config = { ..._config, ...config };
    }
  };
})();
