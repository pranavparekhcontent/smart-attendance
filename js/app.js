/**
 * Smart Attendance PWA — Core App Logic
 * Single Page App Architecture
 */

const App = (() => {
  // ─── STATE ─────────────────────────────────────────
  const state = {
    currentScreen: '',
    role: null, // 'faculty'
    facultyName: '',

    allData: null,

    // Dashboard Selection
    sessionDate: new Date().toISOString().split('T')[0], // YYYY-MM-DD
    selectedSubject: null,
    sessionTopic: '',

    // Attendance Session
    attBatch: '',
    attStudents: [],
    rollcallIndex: 0,

    // Reports State
    reportsSubject: null,
    reportsBatch: '',
    reportStartDate: '',
    reportEndDate: '',
    reportsActiveTab: 'class',
    reportsExpandedDate: null,
    reportData: [],
  };

  let _pendingRole = 'faculty';

  // ─── DEVICE & VIEWPORT ───────────────────────────────
  const Device = {
    isIOS: /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream,
    isAndroid: /Android/.test(navigator.userAgent),
    isStandalone: window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone,
    getHasNotch: () => {
      if (Device.isIOS) {
        const w = window.screen.width, h = window.screen.height;
        const notchRatios = [
          [375, 812], [812, 375], [414, 896], [896, 414],
          [390, 844], [844, 390], [428, 926], [926, 428],
          [393, 852], [852, 393], [430, 932], [932, 430]
        ];
        return notchRatios.some(r => r[0] === w && r[1] === h);
      }
      return false;
    }
  };

  function fixViewport() {
    let vh = window.innerHeight * 0.01;
    document.documentElement.style.setProperty('--vh', `${vh}px`);

    // Notch handling
    if (Device.getHasNotch()) {
      document.body.classList.add('has-notch');
    }
  }
  window.addEventListener('resize', fixViewport);
  fixViewport();
  // ─── ENGINE ENTRY POINT ────────────────────────────
  // Called by the AppStart engine via the appstart:complete event.
  // All license checking, animations, version sync, and background
  // data fetching have already been handled by the engine.

  async function initFromEngine(context) {
    console.log("🚀 App initializing from AppStart engine...", context);

    // 1. Receive data from the engine's background fetch
    let rawData = null;

    // Aggressive Data Extraction (The "Translator")
    if (context.fetchedData) {
      // Check for 'allData' key (new engine default) or directly nested data
      rawData = context.fetchedData.allData || context.fetchedData.data || context.fetchedData;

      // If it's still a string, try parsing it
      if (typeof rawData === 'string' && rawData.trim().startsWith('{')) {
        try { rawData = JSON.parse(rawData); } catch (e) { }
      }
    }

    // 2. Validate and fallback to direct fetch if engine data is missing
    if (!rawData || (!rawData.success && !rawData.teachers)) {
      console.log("🔄 Engine data missing/invalid, attempting direct translator fetch...");
      try {
        rawData = await API.getAllDataFromUrl(context.serverUrl);
      } catch (e) {
        console.warn('Direct fetch failed:', e.message);
      }
    }

    // 3. Normalize Structure (Bridge from GAS format to App format)
    if (rawData) {
      // If data is wrapped in a .data or .records property, unwrapped it
      const actualData = rawData.data || rawData.records || rawData;

      if (actualData.teachers || actualData.subjects || rawData.success) {
        state.allData = actualData;

        // Safety: Ensure required arrays exist
        if (!state.allData.teachers) state.allData.teachers = [];
        if (!state.allData.subjects) state.allData.subjects = [];

        // Persist to cache
        localStorage.setItem('attendance_cache_allData', JSON.stringify({ ts: Date.now(), data: state.allData }));

        // 🚀 BRIDGE: Master Config → Subjects (Centralized Output ID)
        if (context.config && context.config.output_sheet_id) {
          const masterId = API.extractSheetId(context.config.output_sheet_id);
          if (masterId) {
            state.allData.subjects.forEach(s => {
              if (!s.outputSheetId) s.outputSheetId = masterId;
            });
            if (!state.allData.config) state.allData.config = {};
            state.allData.config.outputSheetId = masterId;
          }
        }

        console.log("✅ Translator: Data successfully mapped to app state.", state.allData);
      }
    }

    // 4. Handle errors if translation failed
    if (!state.allData) {
      const errMsg = (rawData && rawData.error) ? rawData.error : 'Format Mismatch';
      console.error("❌ Translator Error:", errMsg, "Raw Payload:", rawData);
      Toast.show('Sync Error: ' + errMsg, 'error');

      // Emergency: Last-last resort cache
      const fallbackRaw = localStorage.getItem('attendance_cache_allData');
      if (fallbackRaw) {
        try {
          state.allData = JSON.parse(fallbackRaw).data;
          Toast.show('Using offline cache', 'warning');
        } catch (e) { }
      }
    }

    // 5. Apply branding from engine context
    const loginTitle = document.querySelector('#screen-login .topbar-title');
    if (loginTitle && context.collegeName) {
      loginTitle.innerText = 'Welcome to ' + context.collegeName;
    }

    // 6. Navigate based on saved login state
    const savedRole = localStorage.getItem('rmd_role');
    const savedName = localStorage.getItem('rmd_faculty');
    if (savedRole === 'faculty' && savedName) {
      state.role = 'faculty';
      state.facultyName = savedName;
      document.getElementById('dash-faculty-name').innerText = savedName;
      document.getElementById('dash-avatar').innerText = savedName.charAt(0).toUpperCase();
      document.getElementById('dash-date').innerText = formatDate(state.sessionDate);
      preloadFacultyData(savedName);
      navigate('faculty-dash');
    } else {
      navigate('login');
    }
  }

  function preloadFacultyData(facultyName) {
    if (!state.allData || !state.allData.subjects || !facultyName) return;
    const mySubjects = state.allData.subjects.filter(s => {
      if (!s.faculty) return false;
      const facs = s.faculty.split(',').map(f => f.trim().toLowerCase());
      return facs.includes(facultyName.toLowerCase());
    });
    mySubjects.forEach(s => {
      if (s.teachingPlanLink) {
        API.getSyllabusPoints(s.teachingPlanLink, s.code).catch(() => {});
      }
      const outId = _getOutputSheetId(s);
      if (outId) {
        API.getTaughtTopics(s.code, outId).catch(() => {});
      }
    });
  }

  // ─── NAVIGATION ────────────────────────────────────

  function navigate(screenId) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    const target = document.getElementById('screen-' + screenId);
    if (target) target.classList.add('active');
    state.currentScreen = screenId;
    window.scrollTo(0, 0);
  }

  // ─── LICENSE ───────────────────────────────────────
  // License is now handled entirely by the AppStart engine.
  // No handleLicenseActivate() needed here.

  // ─── LOGIN ─────────────────────────────────────────

  function showNameTray(role) {
    let html = `<div class="modal-header">
                  <div class="modal-title">Select Name</div>
                  <div class="modal-subtitle">Tap your name to select</div>
                </div>
                <div class="modal-body" style="padding:10px;">`;

    const teachers = (state.allData && state.allData.teachers) ? state.allData.teachers : [];

    if (teachers.length === 0) {
      const isMissing = !state.allData;
      html += `<div style="padding:40px 20px; text-align:center;">
                 <div style="font-size:40px; margin-bottom:15px; opacity:0.3;"><i class="ph-duotone ph-warning-circle"></i></div>
                 <p style="opacity:0.8; margin-bottom:20px;">${isMissing ? 'Data not loaded yet.' : 'No faculty names found in sheet.'}</p>
                 <div style="display:flex; flex-direction:column; gap:10px;">
                   <button class="btn btn-primary" onclick="location.reload()">Reload App</button>
                   <button class="btn btn-outline" onclick="App.closeModal()">Back</button>
                 </div>
               </div>`;
    } else {
      teachers.forEach(t => {
        html += `<div class="subject-list-item" onclick="App.pickLoginName('${t.name}', '${t.pin}')">${t.name}</div>`;
      });
    }
    html += `</div>`;
    showModal(html);
  }

  function pickLoginName(name, pin) {
    // We reopen the login modal with the picked name
    closeModal();
    setTimeout(() => {
      showLoginModal(_pendingRole || 'faculty', name, pin);
    }, 300);
  }

  function showLoginModal(role, preName = '', prePin = '') {
    _pendingRole = role;
    let html = `<div class="modal-header">
                  <div class="modal-title">Faculty Login</div>
                </div>
                <div class="modal-body">
                  <div class="input-group" onclick="App.showNameTray('${role}')">
                    <label class="input-label">Select Name</label>
                    <div class="input" style="display:flex; align-items:center; justify-content:space-between;">
                      <span id="picked-name" style="${preName ? 'color:var(--text-1)' : 'color:var(--text-4)'}">${preName || '-- Select Name --'}</span>
                      <i class="ph-bold ph-caret-down" style="color:var(--text-4)"></i>
                      <input type="hidden" id="login-name-val" value="${preName}" data-pin="${prePin}" />
                    </div>
                  </div>
                  <div class="input-group" style="margin-top: 20px;">
                    <label class="input-label">Enter PIN</label>
                    <input type="password" id="login-pin" class="input" placeholder="****" inputmode="numeric" pattern="[0-9]*" maxlength="6" />
                  </div>
                </div>
                <div class="modal-footer" style="flex-direction:column; gap:10px;">
                  <button class="btn btn-primary btn-full" onclick="App.processLogin('${role}')">Login</button>
                  <button class="btn btn-outline btn-full" onclick="App.closeModal()">Cancel</button>
                </div>`;
    showModal(html);
  }

  function processLogin(role) {
    const name = document.getElementById('login-name-val').value;
    const expectedPin = document.getElementById('login-name-val').dataset.pin;
    const pin = document.getElementById('login-pin').value;

    if (!name) return Toast.show('Select a name', 'error');
    if (!pin) return Toast.show('Enter PIN', 'error');

    // Split valid pins if multiple exist (comma separated in excel)
    const validPins = expectedPin ? expectedPin.split(',').map(p => p.trim()) : [];

    if (validPins.includes(pin)) {
      state.role = role;
      state.facultyName = name;
      localStorage.setItem('rmd_role', role);
      localStorage.setItem('rmd_faculty', name);

      closeModal();
      document.getElementById('dash-faculty-name').innerText = name;
      document.getElementById('dash-avatar').innerText = name.charAt(0).toUpperCase();
      document.getElementById('dash-date').innerText = formatDate(state.sessionDate);
      preloadFacultyData(name);
      navigate('faculty-dash');
      Toast.show(`Welcome ${name}!`, 'success');
    } else {
      Toast.show('Invalid PIN', 'error');
    }
  }

  function logout() {
    state.role = null;
    state.facultyName = '';
    state.selectedSubject = null;
    localStorage.removeItem('rmd_role');
    localStorage.removeItem('rmd_faculty');
    document.getElementById('dash-subject-name').innerText = 'Tap to select subject';
    document.getElementById('dash-subject-meta').innerText = '';
    navigate('login');
  }

  // ─── DASHBOARD: DATE & SUBJECT ─────────────────────

  function showDatePicker() {
    let html = `<div class="modal-header">
                  <div class="modal-title">Select Date</div>
                </div>
                <div class="modal-body">
                  <input type="date" id="picker-date" class="input" value="${state.sessionDate}" max="${new Date().toISOString().split('T')[0]}" style="color:var(--text-1);" />
                </div>
                <div class="modal-footer">
                  <button class="btn btn-glass" style="flex:1" onclick="App.closeModal()">Cancel</button>
                  <button class="btn btn-primary" style="flex:1" onclick="App.setDate()">Confirm</button>
                </div>`;
    showModal(html);
  }

  function setDate() {
    const d = document.getElementById('picker-date').value;
    if (d) {
      state.sessionDate = d;
      document.getElementById('dash-date').innerText = formatDate(d);
    }
    closeModal();
  }

  function showSubjectPicker(mode = 'dash') {
    if (!state.allData || !state.allData.subjects) return Toast.show('Subjects not loaded', 'error');

    // Filter subjects for logged in faculty
    const mySubjects = state.allData.subjects.filter(s => {
      if (!s.faculty) return false;
      return s.faculty.toLowerCase().includes(state.facultyName.toLowerCase());
    });

    if (mySubjects.length === 0) return Toast.show('No subjects assigned to you', 'warning');

    const colors = [
      { primary: '#3b82f6', light: 'rgba(59, 130, 246, 0.08)' }, // Blue
      { primary: '#10b981', light: 'rgba(16, 185, 129, 0.08)' }, // Green
      { primary: '#f59e0b', light: 'rgba(245, 158, 11, 0.08)' },  // Amber
      { primary: '#8b5cf6', light: 'rgba(139, 92, 246, 0.08)' }, // Violet
      { primary: '#ef4444', light: 'rgba(239, 68, 68, 0.08)' },  // Red
      { primary: '#ec4899', light: 'rgba(236, 72, 153, 0.08)' }  // Pink
    ];

    let html = `<div class="modal-header"><div class="modal-title">Select Subject</div></div><div class="modal-body" style="padding: 16px; display:flex; flex-direction:column; gap:12px;">`;
    mySubjects.forEach((s, idx) => {
      const theme = colors[idx % colors.length];
      html += `<div class="subject-list-item" onclick="App.selectSubject('${s.code}', '${mode}')" 
                    style="border-left: 4px solid ${theme.primary}; background: linear-gradient(90deg, ${theme.light} 0%, transparent 100%);">
                 <div style="flex:1">
                   <div class="subject-selected-name" style="margin-bottom:4px; color:${theme.primary}; font-weight:700;">${s.code} - ${s.name}</div>
                   <div class="subject-selected-meta" style="opacity:0.8;">${s.year} | ${s.program} | ${s.semester} | ${s.type}</div>
                 </div>
                 <i class="ph-bold ph-caret-right" style="color:${theme.primary}; opacity:0.5;"></i>
               </div>`;
    });
    html += `</div>`;
    showModal(html);
  }

  function selectSubject(code, mode) {
    const sub = state.allData.subjects.find(s => s.code === code);
    if (sub) {
      state.selectedSubject = sub;
      if (mode === 'dash') {
        document.getElementById('dash-subject-name').innerText = sub.code + ' - ' + sub.name;
        document.getElementById('dash-subject-meta').innerText = sub.year + ' | ' + sub.type;
        document.getElementById('dash-subject-name').classList.remove('subject-placeholder');
        const card = document.getElementById('subject-card');
        if (card) card.classList.remove('heart-beat');
      } else if (mode === 'reports') {
        state.reportsSubject = sub;
        state.reportsBatch = ''; // Reset batch
        document.getElementById('rep-subject-name').innerText = sub.code + ' - ' + sub.name;
        document.getElementById('rep-subject-name').classList.remove('subject-placeholder');

        const isPractical = sub.type.toUpperCase() === 'PRACTICAL';
        document.getElementById('reports-batch-selector').style.display = isPractical ? 'block' : 'none';

        fetchReportData();
      }
    }
    closeModal();
  }

  function renderReportBatchSelector() {
    const container = document.getElementById('rep-batch-list');
    const batches = state.availableReportBatches || [];

    if (batches.length === 0) {
      container.innerHTML = `<div style="font-size:12px; color:var(--text-4); padding:4px 8px;">Loading batches...</div>`;
      return;
    }

    container.innerHTML = batches.map(b => `
      <div class="batch-chip ${state.reportsBatch === b ? 'active' : ''}" onclick="App.selectReportBatch('${b}')">${b}</div>
    `).join('');
  }

  function selectReportBatch(batch) {
    state.reportsBatch = batch;
    renderReportBatchSelector();
    renderReport();
  }

  // ─── ATTENDANCE LOGIC ──────────────────────────────

  function promptTopic() {
    return new Promise((resolve) => {
      let html = `<div class="modal-header">
                    <div class="modal-title">Topic to be Taught</div>
                    <div class="modal-subtitle">What will you teach in this session?</div>
                  </div>
                  <div class="modal-body">
                    <div class="input-group">
                      <label class="input-label">Topic Name</label>
                      <input type="text" id="picker-topic" class="input" placeholder="e.g., Introduction to Arrays" style="color:var(--text-1);" autocomplete="off" />
                    </div>
                  </div>
                  <div class="modal-footer" style="gap:10px;">
                    <button class="btn btn-primary" style="flex:1" onclick="window._submitTopic()">Start</button>
                  </div>`;

      window._submitTopic = () => {
        const val = document.getElementById('picker-topic').value.trim();
        if (!val) {
          Toast.show('Please enter topic', 'warning');
          return;
        }
        closeModal();
        delete window._submitTopic;
        resolve(val);
      };

      showModal(html);

      setTimeout(() => {
        const input = document.getElementById('picker-topic');
        if (input) input.focus();
      }, 100);
    });
  }

  function escapeHtml(str) {
    return str ? String(str).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[c])) : "";
  }

  function isTopicTaught(syllabusPoint, taughtTopics) {
    if (!taughtTopics || !taughtTopics.size || !syllabusPoint) return false;
    const p = String(syllabusPoint).trim().toLowerCase();
    if (taughtTopics.has(p)) return true;

    // 1. Direct exact or substring match
    for (const t of taughtTopics) {
      const lowerT = String(t).trim().toLowerCase();
      if (!lowerT) continue;
      if (lowerT === p) return true;
      if (p.length >= 3 && (lowerT.indexOf(p) !== -1 || p.indexOf(lowerT) !== -1)) return true;
    }

    // 2. Normalized match (strips unit/chapter/numbers and special chars)
    const normalize = (str) => {
      return String(str || '')
        .toLowerCase()
        .replace(/^(unit\s*\d+|chap\s*\d+|chapter\s*\d+|practical\s*\d+|experiment\s*\d+|exp\s*\d+|\d+[\.\)\-\:\s\d]+)[\.\)\-\:\s]*/i, '')
        .replace(/[^a-z0-9]/g, '')
        .trim();
    };

    const normP = normalize(p);
    if (!normP) return false;
    if (taughtTopics.has(normP)) return true;

    for (const t of taughtTopics) {
      const normT = normalize(t);
      if (!normT) continue;
      if (normT === normP) return true;
      if (normP.length >= 3 && (normT.indexOf(normP) !== -1 || normP.indexOf(normT) !== -1)) return true;
    }
    return false;
  }

  function showSyllabusPicker(points, taughtTopics = new Set()) {
    return new Promise((resolve) => {
      const pickerId = 'syl-picker-' + Date.now();
      let html = `<div class="modal-header">
                    <div class="modal-title">Select Topics</div>
                    <div class="modal-subtitle">Tap to select one or more syllabus points</div>
                  </div>
                  <div class="modal-body" style="padding: 16px 20px; max-height:50vh; overflow-y:auto; display:flex; flex-direction:column; gap:12px;">`;

      // Search input if there are more than 8 points
      if (points.length > 8) {
        html += `<div class="input-group" style="margin-bottom: 4px;">
                   <input type="text" id="syl-search-input" class="input" placeholder="🔍 Search syllabus points..." style="font-size: 13px; padding: 8px 12px; border-radius:8px;" autocomplete="off" oninput="window._filterSylChips(this.value)" />
                 </div>`;
      }

      html += `<div id="${pickerId}" style="display:flex; flex-direction:column; gap:10px; width:100%;">`;

      // Render syllabus chips
      points.forEach((pt, idx) => {
        const safeVal = escapeHtml(pt);
        const isTaught = isTopicTaught(pt, taughtTopics);
        html += `<button type="button" class="syl-chip ${isTaught ? 'taught' : ''}" data-idx="${idx}" data-raw-topic="${encodeURIComponent(pt)}" data-value="${encodeURIComponent(pt)}" onclick="window._toggleSylChip(this)">
                   <span class="syl-chip-dot"></span>
                   <span class="syl-chip-text" style="flex:1;">${safeVal}</span>
                   ${isTaught ? `<span class="syl-chip-badge"><i class="ph-bold ph-check" style="margin-right:2px;"></i> Taught</span>` : ''}
                 </button>`;
      });

      html += `</div>
               </div>
               <div style="padding: 0 20px 12px;">
                 <div class="syl-custom-row" id="syl-custom-row">
                   <button type="button" class="syl-chip syl-chip-other" onclick="window._toggleSylCustom(this)" id="syl-other-btn">
                     <span class="syl-chip-dot" style="background:var(--warning);box-shadow:0 0 6px var(--warning);"></span>
                     <span class="syl-chip-text">✨ Custom Topic</span>
                   </button>
                   <input type="text" id="syl-custom-input" class="input syl-custom-input" placeholder="Enter custom topic..." style="display:none;" autocomplete="off" oninput="window._updateSylConfirm()" />
                 </div>
               </div>
               <div class="modal-footer" style="padding: 10px 20px 20px; gap:10px;">
                 <button class="btn btn-glass" style="flex:1" onclick="window._cancelSyllabus()">Cancel</button>
                 <button class="btn btn-primary" style="flex:1.5; opacity:0.4; pointer-events:none;" id="syl-confirm-btn" onclick="window._confirmSyllabus()">
                   <i class="ph-bold ph-check-circle" style="margin-right:6px;"></i> Confirm
                 </button>
               </div>`;

      // Live update taught chips if background fetch finishes while picker is open
      window._updateTaughtChips = (updatedTaught) => {
        if (!updatedTaught) return;
        const chips = document.querySelectorAll(`#${pickerId} .syl-chip[data-raw-topic]`);
        chips.forEach(chip => {
          const raw = decodeURIComponent(chip.dataset.rawTopic || '');
          if (isTopicTaught(raw, updatedTaught)) {
            chip.classList.add('taught');
            if (!chip.querySelector('.syl-chip-badge')) {
              const badge = document.createElement('span');
              badge.className = 'syl-chip-badge';
              badge.innerHTML = '<i class="ph-bold ph-check" style="margin-right:2px;"></i> Taught';
              chip.appendChild(badge);
            }
          }
        });
      };

      // Filter chips
      window._filterSylChips = (query) => {
        const q = query.toLowerCase().trim();
        const chips = document.querySelectorAll(`#${pickerId} .syl-chip`);
        chips.forEach(chip => {
          const text = chip.querySelector('.syl-chip-text').textContent.toLowerCase();
          if (text.includes(q)) {
            chip.style.display = 'inline-flex';
          } else {
            chip.style.display = 'none';
          }
        });
      };

      // Toggle chip selection
      window._toggleSylChip = (el) => {
        el.classList.toggle('selected');
        updateConfirmBtn();
      };

      // Toggle custom input
      window._toggleSylCustom = (el) => {
        el.classList.toggle('selected');
        const input = document.getElementById('syl-custom-input');
        if (el.classList.contains('selected')) {
          input.style.display = 'block';
          setTimeout(() => input.focus(), 100);
        } else {
          input.style.display = 'none';
          input.value = '';
        }
        updateConfirmBtn();
      };

      // Public handle to update confirm button from inline event listener
      window._updateSylConfirm = () => {
        updateConfirmBtn();
      };

      function updateConfirmBtn() {
        const selected = document.querySelectorAll(`#${pickerId} .syl-chip.selected`);
        const otherBtn = document.getElementById('syl-other-btn');
        const customVal = document.getElementById('syl-custom-input')?.value?.trim();
        const customSelected = otherBtn?.classList.contains('selected') && customVal;
        const totalSelected = selected.length + (customSelected ? 1 : 0);

        const btn = document.getElementById('syl-confirm-btn');
        if (totalSelected > 0) {
          btn.style.opacity = '1';
          btn.style.pointerEvents = 'auto';
          btn.innerHTML = `<i class="ph-bold ph-check-circle" style="margin-right:6px;"></i> Confirm (${totalSelected})`;
        } else {
          // If custom button is selected but no text entered yet, let them confirm ONLY if they have regular chips selected. Otherwise disable confirm.
          if (selected.length > 0) {
            btn.style.opacity = '1';
            btn.style.pointerEvents = 'auto';
            btn.innerHTML = `<i class="ph-bold ph-check-circle" style="margin-right:6px;"></i> Confirm (${selected.length})`;
          } else {
            btn.style.opacity = '0.4';
            btn.style.pointerEvents = 'none';
            btn.innerHTML = `<i class="ph-bold ph-check-circle" style="margin-right:6px;"></i> Confirm`;
          }
        }
      }

      // Confirm selection
      window._confirmSyllabus = () => {
        const selected = document.querySelectorAll(`#${pickerId} .syl-chip.selected`);
        const parts = [];
        selected.forEach(el => parts.push(decodeURIComponent(el.dataset.value)));

        const otherBtn = document.getElementById('syl-other-btn');
        if (otherBtn?.classList.contains('selected')) {
          const customVal = document.getElementById('syl-custom-input')?.value?.trim();
          if (customVal) parts.push(customVal);
        }

        cleanup();
        closeModal();
        resolve(parts.length > 0 ? parts.join(', ') : null);
      };

      window._cancelSyllabus = () => {
        cleanup();
        closeModal();
        resolve(null);
      };

      function cleanup() {
        delete window._toggleSylChip;
        delete window._toggleSylCustom;
        delete window._confirmSyllabus;
        delete window._cancelSyllabus;
      }

      showModal(html);
    });
  }

  function _getOutputSheetId(subject) {
    if (subject && subject.outputSheetId) return subject.outputSheetId;
    if (state.allData && state.allData.config && state.allData.config.outputSheetId) return state.allData.config.outputSheetId;
    if (window.appStartContext) {
      const cfg = window.appStartContext.config || {};
      const possibleKeys = ['output_sheet_id', 'output sheet id', 'output_sheet', 'output sheet', 'outputsheetid', 'output_sheet_url', 'output sheet link', 'output sheet url', 'sheet_id', 'sheet id'];
      for (const k of possibleKeys) {
        if (cfg[k]) return API.extractSheetId(cfg[k]);
      }
      if (window.appStartContext.sheetId) return window.appStartContext.sheetId;
    }
    return '';
  }

  async function startAttendanceFlow() {
    if (!state.selectedSubject) return Toast.show('Please select a subject first', 'warning');

    let topic = '';
    let hasSyllabusPoints = false;
    let syllabusPoints = [];
    const taughtTopics = new Set();

    if (state.selectedSubject.teachingPlanLink) {
      const sylCacheKey = 'syl_' + (state.selectedSubject.code || '') + '_' + (state.selectedSubject.teachingPlanLink || '');
      let cachedSyl = (window.API && API._getCache) ? API._getCache(sylCacheKey) : null;

      const outId = _getOutputSheetId(state.selectedSubject);
      const cleanOutId = (window.API && API.extractSheetId) ? API.extractSheetId(outId) : outId;
      const taughtCacheKey = 'taught_' + (state.selectedSubject.code || '') + '_' + (cleanOutId || '');
      let cachedTaught = (window.API && API._getCache) ? API._getCache(taughtCacheKey) : null;

      if (cachedTaught && cachedTaught.topics && cachedTaught.topics.length > 0) {
        cachedTaught.topics.forEach(t => taughtTopics.add(String(t).trim().toLowerCase()));
      }

      // If either syllabus or taught topics is missing in local cache, fetch them together!
      if (!cachedSyl || !cachedSyl.points || cachedSyl.points.length === 0 || !cachedTaught || !cachedTaught.topics || cachedTaught.topics.length === 0) {
        showSpinner('Fetching Syllabus & Topics...', 'ph-book-open');
        try {
          const [sylRes, taughtRes] = await Promise.allSettled([
            (!cachedSyl || !cachedSyl.points || cachedSyl.points.length === 0) ? API.getSyllabusPoints(state.selectedSubject.teachingPlanLink, state.selectedSubject.code) : Promise.resolve(cachedSyl),
            (!cachedTaught || !cachedTaught.topics || cachedTaught.topics.length === 0) ? API.getTaughtTopics(state.selectedSubject.code, cleanOutId) : Promise.resolve(cachedTaught)
          ]);

          if (sylRes.status === 'fulfilled' && sylRes.value && sylRes.value.points) {
            syllabusPoints = sylRes.value.points;
            hasSyllabusPoints = true;
          }
          if (taughtRes.status === 'fulfilled' && taughtRes.value && taughtRes.value.topics) {
            taughtRes.value.topics.forEach(t => taughtTopics.add(String(t).trim().toLowerCase()));
          }
        } catch (e) {
          console.warn('Error fetching syllabus points / taught topics:', e);
        } finally {
          hideSpinner();
        }
      } else {
        syllabusPoints = cachedSyl.points;
        hasSyllabusPoints = true;
      }
    }

    if (hasSyllabusPoints) {
      const choice = await showSyllabusPicker(syllabusPoints, taughtTopics);
      if (choice === null) {
        // User cancelled syllabus selection
        return;
      }
      topic = choice;
    } else {
      Toast.show('Please add syllabus in teaching plan excel', 'warning');
      topic = await promptTopic();
      if (!topic) return;
    }

    state.sessionTopic = topic;

    showSpinner('Fetching Students...', 'ph-users-three');
    const res = await API.getStudents(state.selectedSubject.year);
    hideSpinner();

    if (!res.success) return Toast.show(res.error, 'error');

    state.attStudents = res.students;
    state.sessionDateSuffix = '';

    if (state.selectedSubject.type.toUpperCase() === 'PRACTICAL') {
      setupAttendanceUI();
      return;
    }

    // Pre-check for THEORY
    if (navigator.onLine) {
      showSpinner('Checking Session...', 'ph-magnifying-glass');
      const sessionRes = await API.getAttendance(state.selectedSubject.code, state.selectedSubject.year, dbFormatDate(state.sessionDate), state.selectedSubject.outputSheetId);
      hideSpinner();

      if (sessionRes.success && sessionRes.records) {
        const existingDates = [...new Set(sessionRes.records.map(r => r.date))];
        if (existingDates.length > 0) {
          const decision = await promptConflict(existingDates);
          if (decision.choice === 'cancel') return;
          if (decision.choice === 'another') {
            await resolveAnotherSession();
          } else if (decision.choice === 'overwrite') {
            if (decision.date.includes(' (')) {
              state.sessionDateSuffix = decision.date.substring(decision.date.indexOf(' ('));
            } else {
              state.sessionDateSuffix = '';
            }
          }
        }
      }
    }

    setupAttendanceUI();
  }

  async function resolveAnotherSession() {
    showSpinner('Allocating Session...', 'ph-browser');
    await sleep(800);
    hideSpinner();
    const type = state.selectedSubject.type.toUpperCase();
    let baseSuffix = type === 'PRACTICAL' ? ' (P' : ' (L';
    let nextNum = 2;

    while (nextNum <= 12) {
      let testDateStr = dbFormatDate(state.sessionDate) + baseSuffix + nextNum + ')';
      let check = await API.getAttendance(state.selectedSubject.code, state.selectedSubject.year, testDateStr);
      let isConflict = false;

      if (check.success && check.records && check.records.length > 0) {
        if (type === 'PRACTICAL') {
          isConflict = check.records.some(r => r.batch === state.attBatch);
        } else {
          isConflict = true;
        }
      }

      if (isConflict) {
        nextNum++;
      } else {
        state.sessionDateSuffix = baseSuffix + nextNum + ')';
        break;
      }
    }
    if (!state.sessionDateSuffix) {
      state.sessionDateSuffix = baseSuffix + nextNum + ')';
    }
    closeModal();
  }

  function setupAttendanceUI() {
    closeModal();
    document.getElementById('att-subject-name').innerText = state.selectedSubject.name;
    document.getElementById('att-subject-meta').innerText = `${state.selectedSubject.year} | ${state.selectedSubject.type} | ${formatDate(state.sessionDate)}${state.sessionDateSuffix}`;

    // Add default status
    state.attStudents.forEach(s => s.status = null);
    updateCounters();

    if (state.selectedSubject.type.toUpperCase() === 'PRACTICAL') {
      document.getElementById('att-batch-selector').style.display = 'block';
      document.getElementById('att-list-view').style.display = 'none';
      document.getElementById('att-rollcall-view').style.display = 'none';
      renderBatches();
    } else {
      document.getElementById('att-batch-selector').style.display = 'none';
      askAttendanceMode();
    }
    navigate('attendance-mode');
  }

  function renderBatches() {
    const batches = [...new Set(state.attStudents.map(s => s.batch).filter(b => b))].sort();
    let html = '';
    batches.forEach(b => {
      const isActive = state.attBatch === b;
      html += `<div class="batch-chip ${isActive ? 'active' : ''}" onclick="App.selectBatch('${b}', this)">${b}</div>`;
    });
    document.getElementById('att-batch-list').innerHTML = html;
  }

  async function selectBatch(batch, elem) {
    document.querySelectorAll('.batch-chip').forEach(c => c.classList.remove('active'));
    elem.classList.add('active');
    state.attBatch = batch;
    updateCounters();

    // Pre-check for PRACTICAL BATCH
    if (navigator.onLine) {
      showSpinner(`Checking session for Batch ${batch}...`, 'ph-magnifying-glass');
      const sessionRes = await API.getAttendance(state.selectedSubject.code, state.selectedSubject.year, dbFormatDate(state.sessionDate), state.selectedSubject.outputSheetId);
      hideSpinner();

      if (sessionRes.success && sessionRes.records) {
        const existingDates = [...new Set(sessionRes.records.filter(r => r.batch === batch).map(r => r.date))];
        if (existingDates.length > 0) {
          const decision = await promptConflict(existingDates);
          if (decision.choice === 'cancel') {
            state.attBatch = '';
            elem.classList.remove('active');
            return;
          } else if (decision.choice === 'another') {
            await resolveAnotherSession();
            document.getElementById('att-subject-meta').innerText = `${state.selectedSubject.year} | ${state.selectedSubject.type} | ${formatDate(state.sessionDate)}${state.sessionDateSuffix}`;
          } else if (decision.choice === 'overwrite') {
            if (decision.date.includes(' (')) {
              state.sessionDateSuffix = decision.date.substring(decision.date.indexOf(' ('));
            } else {
              state.sessionDateSuffix = '';
            }
            document.getElementById('att-subject-meta').innerText = `${state.selectedSubject.year} | ${state.selectedSubject.type} | ${formatDate(state.sessionDate)}${state.sessionDateSuffix}`;
          }
        }
      }
    }

    askAttendanceMode();
  }

  function askAttendanceMode() {
    let html = `<div class="modal-header"><div class="modal-title">Select Entry Mode</div></div>
                <div class="modal-body" style="display:flex; flex-direction:column; gap:12px;">
                  <button class="mode-btn present" onclick="App.setEntryMode('all-present')">
                    <div class="mode-btn-icon" style="color:var(--success)"><i class="ph-fill ph-check-circle"></i></div>
                    <div><div class="mode-btn-name">All Present</div><div class="mode-btn-sub">Mark absentees manually</div></div>
                  </button>
                  <button class="mode-btn absent" onclick="App.setEntryMode('all-absent')">
                    <div class="mode-btn-icon" style="color:var(--danger)"><i class="ph-fill ph-x-circle"></i></div>
                    <div><div class="mode-btn-name">All Absent</div><div class="mode-btn-sub">Mark presenters manually</div></div>
                  </button>
                  <button class="mode-btn rollcall" onclick="App.setEntryMode('rollcall')">
                    <div class="mode-btn-icon" style="color:var(--accent)"><i class="ph-fill ph-microphone-stage"></i></div>
                    <div><div class="mode-btn-name">Roll Call</div><div class="mode-btn-sub">Call out names one by one</div></div>
                  </button>
                </div>`;
    showModal(html, true);
  }

  function setEntryMode(mode) {
    closeModal();
    const studentsToShow = state.selectedSubject.type.toUpperCase() === 'PRACTICAL'
      ? state.attStudents.filter(s => s.batch === state.attBatch)
      : state.attStudents;

    if (mode === 'all-present') {
      studentsToShow.forEach(s => s.status = 'P');
      renderListView(studentsToShow);
    } else if (mode === 'all-absent') {
      studentsToShow.forEach(s => s.status = 'A');
      renderListView(studentsToShow);
    } else if (mode === 'rollcall') {
      state.rollcallIndex = 0;
      studentsToShow.forEach(s => s.status = null); // clear to force selection
      startRollcall(studentsToShow);
    }
  }

  function renderListView(students) {
    document.getElementById('att-list-view').style.display = 'block';
    document.getElementById('att-rollcall-view').style.display = 'none';

    let html = '';
    students.forEach(s => {
      const isP = s.status === 'P';
      html += `<div class="student-row" onclick="App.toggleStudentStatus(event, '${s.rollNo}')">
                 <div class="student-roll">${s.rollNo}</div>
                 <div class="student-name">${s.name}</div>
                 <div class="pa-toggle ${isP ? 'present' : 'absent'}" id="toggle-${s.rollNo}">
                   ${isP ? 'PRESENT' : 'ABSENT'}
                 </div>
               </div>`;
    });
    document.getElementById('att-students-container').innerHTML = html;
    updateCounters();
  }

  function toggleStudentStatus(event, rollNo) {
    if (event) {
      event.stopPropagation();
      event.stopImmediatePropagation();
    }
    const st = state.attStudents.find(s => s.rollNo == rollNo);
    if (st) {
      st.status = st.status === 'P' ? 'A' : 'P';
      const btn = document.getElementById(`toggle-${rollNo}`);
      if (btn) {
        btn.className = `pa-toggle ${st.status === 'P' ? 'present' : 'absent'}`;
        btn.innerText = st.status === 'P' ? 'PRESENT' : 'ABSENT';
      }
      updateCounters();
    }
  }

  function startRollcall(students) {
    document.getElementById('att-list-view').style.display = 'none';
    document.getElementById('att-rollcall-view').style.display = 'block';
    updateRollcallUI(students);
  }

  function updateRollcallUI(students) {
    if (state.rollcallIndex >= students.length) {
      // Done with rollcall
      Toast.show('Roll call complete', 'success');
      renderListView(students); // switch back to list view to verify/save
      return;
    }

    const st = students[state.rollcallIndex];
    document.getElementById('rollcall-counter').innerText = `Student ${state.rollcallIndex + 1} of ${students.length}`;
    document.getElementById('rollcall-roll').innerText = st.rollNo;
    document.getElementById('rollcall-name').innerText = st.name;

    // Draw prev indicators
    let prevHtml = '';
    const startIdx = Math.max(0, state.rollcallIndex - 5);
    for (let i = startIdx; i < state.rollcallIndex; i++) {
      const pSt = students[i];
      const color = pSt.status === 'P' ? 'var(--success)' : 'var(--danger)';
      prevHtml += `<div style="font-size:10px; padding:2px 6px; border-radius:4px; border:1px solid ${color}; color:${color}">${pSt.rollNo}</div>`;
    }
    document.getElementById('rollcall-prev').innerHTML = prevHtml;
    updateCounters();
  }

  function markRollcall(event, status) {
    if (event) {
      event.stopPropagation();
      event.stopImmediatePropagation();
    }
    const students = state.selectedSubject.type.toUpperCase() === 'PRACTICAL'
      ? state.attStudents.filter(s => s.batch === state.attBatch)
      : state.attStudents;

    students[state.rollcallIndex].status = status;
    state.rollcallIndex++;
    updateRollcallUI(students);
  }

  function updateCounters() {
    const isPrac = state.selectedSubject.type.toUpperCase() === 'PRACTICAL';
    if (isPrac && !state.attBatch) {
      document.getElementById('att-count-p').innerText = 0;
      document.getElementById('att-count-a').innerText = 0;
      document.getElementById('att-count-t').innerText = 0;
      return;
    }

    const students = isPrac
      ? state.attStudents.filter(s => s.batch === state.attBatch)
      : state.attStudents;

    const p = students.filter(s => s.status === 'P').length;
    const a = students.filter(s => s.status === 'A').length;
    const t = students.length;

    document.getElementById('att-count-p').innerText = p;
    document.getElementById('att-count-a').innerText = a;
    document.getElementById('att-count-t').innerText = t;
  }

  // --- MODALS & DIALOGS ---
  function promptConflict(existingDates = []) {
    return new Promise((resolve) => {
      let overwriteButtons = '';
      if (existingDates.length === 0) {
        overwriteButtons = `
          <button class="btn btn-full" style="background: rgba(239,68,68,0.1); color: #f87171; margin-bottom: 12px;" onclick="window._resolveConflict({choice: 'overwrite', date: ''})">
            <i class="ph-bold ph-warning"></i> Overwrite Existing
          </button>
        `;
      } else {
        existingDates.forEach(d => {
          let suffix = '';
          if (d.includes(' (')) {
            suffix = d.substring(d.indexOf(' ('));
          }
          const display = formatDate(d.split(' ')[0]) + suffix;
          overwriteButtons += `
            <button class="btn btn-full" style="background: rgba(239,68,68,0.1); color: #f87171; margin-bottom: 12px; height: auto; padding: 10px;" onclick="window._resolveConflict({choice: 'overwrite', date: '${d}'})">
              <i class="ph-bold ph-warning"></i> Overwrite ${display}
            </button>
          `;
        });
      }

      const html = `
        <div style="text-align:center; padding: 16px;">
           <i class="ph-fill ph-warning-circle" style="font-size: 48px; color: var(--danger); margin-bottom: 16px;"></i>
           <h3 style="color: white; margin-bottom: 8px;">Attendance Already Exists</h3>
           <p style="color: var(--text-3); font-size: 14px; margin-bottom: 24px;">Records for this date already exist in the database. What would you like to do?</p>
           
           <button class="btn btn-primary btn-full" style="margin-bottom: 12px; height: auto; padding: 12px; flex-direction: column;" onclick="window._resolveConflict({choice: 'another'})">
             <div style="font-weight: 600; font-size: 15px;">Add Another Session</div>
             <div style="font-size: 12px; opacity: 0.8; font-weight: 400; margin-top: 4px;">Mark as (L2) or (P2)</div>
           </button>
           
           ${overwriteButtons}
           
           <button class="btn btn-outline btn-full" onclick="window._resolveConflict({choice: 'cancel'})">Cancel</button>
        </div>
      `;

      window._resolveConflict = (choice) => {
        closeModal();
        delete window._resolveConflict;
        resolve(choice);
      };

      showModal(html);
    });
  }



  async function saveAttendance(event) {
    if (event) {
      event.stopPropagation();
      event.stopImmediatePropagation();
    }
    const type = state.selectedSubject.type.toUpperCase();
    const students = type === 'PRACTICAL'
      ? state.attStudents.filter(s => s.batch === state.attBatch)
      : state.attStudents;

    if (students.some(s => !s.status)) {
      return Toast.show('Please mark all students before saving', 'warning');
    }

    let targetDateStr = dbFormatDate(state.sessionDate) + (state.sessionDateSuffix || '');

    const records = students.map(s => ({
      date: targetDateStr,
      code: state.selectedSubject.code,
      year: state.selectedSubject.year,
      batch: type === 'PRACTICAL' ? state.attBatch : '',
      faculty: state.facultyName,
      rollNo: s.rollNo,
      name: s.name,
      status: s.status,
      topic: state.sessionTopic || ''
    }));

    showSpinner('Saving Attendance...', 'ph-cloud-arrow-up');
    const res = await API.saveAttendance(records, state.selectedSubject.outputSheetId);
    hideSpinner();

    if (res.success) {
      state.lastSavedRecords = records;

      // Immediately cache newly taught topic locally so green badge appears instantly next time
      if (state.sessionTopic && state.selectedSubject) {
        const taughtKey = 'taught_' + (state.selectedSubject.code || '') + '_' + (state.selectedSubject.outputSheetId || '');
        const cached = (window.API && API._getCache) ? (API._getCache(taughtKey) || { success: true, topics: [] }) : { success: true, topics: [] };
        if (!cached.topics) cached.topics = [];
        if (!cached.topics.includes(state.sessionTopic)) {
          cached.topics.push(state.sessionTopic);
          if (window.API && API._setCache) API._setCache(taughtKey, cached);
        }
      }

      showSessionCompleteDialog(
        students.filter(s => s.status === 'P').length,
        students.filter(s => s.status === 'A').length,
        res.isLocallySaved
      );
      try {
        window.open('https://omg10.com/4/11324927', '_blank', 'noopener,noreferrer');
      } catch (e) {
        console.warn("Direct Link blocked on save:", e);
      }
    } else {
      Toast.show(res.error || 'Failed to save', 'error');
    }
  }

  function showSessionCompleteDialog(p, a, isLocallySaved = false) {
    const titleText = isLocallySaved ? 'Session Saved! (Saved Locally)' : 'Session Saved!';
    const subText = isLocallySaved ? 'Saved on device. Will auto-sync in background.' : 'Attendance securely recorded.';
    let html = `<div class="session-complete-hero">
                  <div class="session-complete-icon"><i class="ph-fill ph-check-circle"></i></div>
                  <div class="session-complete-title">${titleText}</div>
                  <div class="session-complete-sub">${subText}</div>
                </div>
                <div class="stat-pills">
                  <div class="stat-pill green"><div class="stat-pill-val">${p}</div><div class="stat-pill-key">Present</div></div>
                  <div class="stat-pill red"><div class="stat-pill-val">${a}</div><div class="stat-pill-key">Absent</div></div>
                </div>
                <div style="padding: 0 24px 24px;">
                  <button class="btn btn-primary btn-full" style="margin-bottom:12px;" onclick="App.shareLastAttendance()">
                    <i class="ph-bold ph-whatsapp-logo"></i> Share to WhatsApp
                  </button>
                  <button class="btn btn-glass btn-full" onclick="App.handleReturnToDashboard()">Return to Dashboard</button>
                </div>`;
    showModal(html, true);
  }

  function shareLastAttendance() {
    if (state.lastSavedRecords && state.lastSavedRecords.length > 0) {
      generateWhatsAppMessage(state.lastSavedRecords);
    }
  }

  function handleReturnToDashboard() {
    try {
      window.open('https://omg10.com/4/11324927', '_blank', 'noopener,noreferrer');
    } catch (e) {
      console.warn("Direct Link open blocked or failed:", e);
    }
    navigate('faculty-dash');
    closeModal();
  }

  function openAdOnce() {
    if (openAdOnce._shown) return;
    openAdOnce._shown = true;
    try {
      window.open('https://omg10.com/4/11324927', '_blank', 'noopener,noreferrer');
    } catch (e) {
      console.warn("Direct Link blocked on WhatsApp share:", e);
    }
  }

  function generateWhatsAppMessage(records) {
    const sub = state.selectedSubject;
    const p = records.filter(r => r.status === 'P').length;
    const t = records.length;
    const pct = ((p / t) * 100).toFixed(1);
    const absentees = records.filter(r => r.status === 'A').map(r => r.rollNo).join(', ') || 'None';

    let msg = `📅 *Date*      : ${formatDate(state.sessionDate)}${state.sessionDateSuffix ? ' ' + state.sessionDateSuffix : ''}\n`;
    msg += `🎓 *Class*     : ${sub.program} · ${sub.year}\n`;
    msg += `📚 *Subject*   : ${sub.name} (${sub.code}) · ${sub.type}\n`;
    if (records[0].topic && records[0].topic.trim() !== '') msg += `📋 *Topic*     : ${records[0].topic.trim()}\n`;
    if (records[0].batch) msg += `🧪 *Batch*     : ${records[0].batch}\n`;
    msg += `🧑‍🏫 *Faculty*   : ${state.facultyName}\n`;
    msg += `👥 *Attendance*: ${p} / ${t} students\n`;
    msg += `📊 *Percentage*: ${pct}%\n`;
    msg += `🚫 *Absent*    : ${absentees}\n`;

    const url = `whatsapp://send?text=${encodeURIComponent(msg)}`;
    window.location.href = url;

    // Direct Link ad (popunder replacement) — after WhatsApp, once per session
    setTimeout(openAdOnce, 1200);
  }

  // ─── REPORTS ───────────────────────────────────────

  function openReports() {
    navigate('reports');
    if (state.reportsSubject) {
      fetchReportData();
    }
  }

  async function fetchReportData() {
    if (!state.reportsSubject) return;

    const cacheKey = 'report_' + [state.reportsSubject.code, state.reportsSubject.year, state.reportsSubject.outputSheetId || ''].join('_');
    const cached = (window.API && API._getCache) ? API._getCache(cacheKey) : null;
    if (cached && cached.records && cached.records.length > 0) {
      state.reportData = cached.records;
      document.getElementById('reports-filters').style.display = 'block';
      document.getElementById('rep-footer').style.display = 'block';
      renderReport(); // Instant paint from cache
    } else {
      showSpinner('Fetching Report Data...', 'ph-chart-line-up');
    }

    try {
      const [att, stu] = await Promise.allSettled([
        API.getAttendance(state.reportsSubject.code, state.reportsSubject.year, null, state.reportsSubject.outputSheetId),
        API.getStudents(state.reportsSubject.year)
      ]);

      const res = att.status === 'fulfilled' ? att.value : { success: false, error: (att.reason && att.reason.message) || 'Attendance request failed' };
      const studentRes = stu.status === 'fulfilled' ? stu.value : { success: false, students: [] };

      if (res && res.success) {
        state.reportData = res.records || [];
        if (window.API && API._setCache) API._setCache(cacheKey, { records: state.reportData });

        // Dynamic Batch Detection — from student list as intended
        if (state.reportsSubject.type.toUpperCase() === 'PRACTICAL' && studentRes.success) {
          const studentBatches = [...new Set((studentRes.students || []).map(s => s.batch))].filter(b => b).sort();
          state.availableReportBatches = studentBatches;
          renderReportBatchSelector();
        }

        document.getElementById('reports-filters').style.display = 'block';
        document.getElementById('rep-footer').style.display = 'block';
        renderReport();
      } else if (!state.reportData || state.reportData.length === 0) {
        Toast.show((res && res.error) || 'Failed to fetch reports', 'error');
      } else {
        Toast.show('Showing cached report (server slow)', 'warning');
      }
    } catch (e) {
      console.error('fetchReportData error:', e);
      if (!state.reportData || state.reportData.length === 0) {
        Toast.show(e.message || 'Request timeout', 'error');
      }
    } finally {
      hideSpinner(); // GUARANTEED to hide spinner under all circumstances!
    }
  }

  function handleReportFilterChange() {
    state.reportStartDate = document.getElementById('rep-start-date').value;
    state.reportEndDate = document.getElementById('rep-end-date').value;
    renderReport();
  }

  function setReportRange(type) {
    const start = document.getElementById('rep-start-date');
    const end = document.getElementById('rep-end-date');

    if (type === 'all') {
      start.value = '';
      end.value = '';
    }

    handleReportFilterChange();
  }

  function switchReportTab(tab) {
    state.reportsActiveTab = tab;
    document.querySelectorAll('.rep-tab').forEach(t => {
      t.className = 'badge rep-tab';
      t.style.background = 'var(--bg-surface)';
      t.style.color = 'var(--text-3)';
      t.style.border = '1px solid var(--border-color)';
    });
    const active = document.querySelector(`.rep-tab[data-tab="${tab}"]`);
    if (active) {
      active.className = 'badge badge-primary rep-tab';
      active.style = '';
    }
    renderReport();
  }

  function renderReport() {
    const container = document.getElementById('rep-content');
    if (!state.reportsSubject || !state.reportData) return;

    // Filter by batch if practical
    let filtered = state.reportData;
    if (state.reportsSubject.type.toUpperCase() === 'PRACTICAL') {
      if (!state.reportsBatch) {
        container.innerHTML = `<div class="empty-state"><i class="ph-bold ph-hand-pointing empty-icon"></i><div class="empty-title">Select a batch</div><p style="color:var(--text-4);font-size:12px;">Please tap a batch above to view report.</p></div>`;
        return;
      }
      filtered = filtered.filter(r => r.batch === state.reportsBatch);
    }

    // Filter by date
    if (state.reportStartDate) {
      filtered = filtered.filter(r => r.date >= state.reportStartDate);
    }
    if (state.reportEndDate) {
      filtered = filtered.filter(r => r.date <= state.reportEndDate);
    }

    if (filtered.length === 0) {
      const batchInfo = state.reportsBatch ? ` for Batch ${state.reportsBatch}` : '';
      container.innerHTML = `<div class="empty-state"><i class="ph-bold ph-ghost empty-icon"></i><div class="empty-title">No records found</div><p style="color:var(--text-4);font-size:12px;">Try adjusting your filters${batchInfo}.</p></div>`;
      return;
    }

    if (state.reportsActiveTab === 'class') {
      renderClassReport(filtered);
    } else if (state.reportsActiveTab === 'date') {
      renderDateReport(filtered);
    } else if (state.reportsActiveTab === 'defaulter') {
      renderDefaulterReport(filtered);
    }
  }

  function renderClassReport(data) {
    const students = {};
    const totalSessions = new Set(data.map(r => r.date)).size;

    data.forEach(r => {
      if (!students[r.rollNo]) {
        students[r.rollNo] = { name: r.name, present: 0, total: 0 };
      }
      if (r.status === 'P') students[r.rollNo].present++;
      students[r.rollNo].total++;
    });

    let html = `<div style="display:flex; justify-content:space-between; margin-bottom:12px; font-size:12px; color:var(--text-3);">
                  <span>Total Students: ${Object.keys(students).length}</span>
                  <span>Total Sessions: ${totalSessions}</span>
                </div>`;

    Object.keys(students).sort((a, b) => a - b).forEach(roll => {
      const s = students[roll];
      const pct = Math.round((s.present / s.total) * 100);
      const isDefaulter = pct < ((state.allData && state.allData.attendanceLimit) || 75);

      html += `<div class="glass-card" style="padding:12px; margin-bottom:10px; display:flex; align-items:center; gap:12px; justify-content:center; text-align:center;">
                 <div style="width:36px; height:36px; background:var(--bg-surface); border-radius:50%; display:flex; align-items:center; justify-content:center; font-weight:bold; font-size:14px; color:var(--accent); border:1px solid var(--border-color)">${roll}</div>
                 <div style="flex:1">
                   <div style="font-size:14px; font-weight:600; color:white;">${s.name}</div>
                   <div style="font-size:11px; color:var(--text-3); margin-top:2px;">Attended: ${s.present} / ${s.total}</div>
                 </div>
                 <div style="text-align:right">
                   <div style="font-size:16px; font-weight:bold; color:${isDefaulter ? 'var(--danger)' : 'var(--success)'}">${pct}%</div>
                 </div>
               </div>`;
    });

    document.getElementById('rep-content').innerHTML = html;
  }

  function renderDateReport(data) {
    const dates = {};
    data.forEach(r => {
      if (!dates[r.date]) dates[r.date] = { present: [], total: 0 };
      if (r.status === 'P') dates[r.date].present.push(r.rollNo);
      dates[r.date].total++;
    });

    let html = '';
    Object.keys(dates).sort().reverse().forEach(d => {
      const stats = dates[d];
      const pCount = stats.present.length;
      const topicRec = data.find(r => r.date === d && r.topic && r.topic.trim() !== '');
      const pct = Math.round((pCount / stats.total) * 100);
      const isExpanded = state.reportsExpandedDate === d;

      html += `<div class="glass-card" style="padding:12px; margin-bottom:10px; cursor:pointer;" onclick="App.toggleDateDetails('${d}')">
                 <div style="display:flex; align-items:center; gap:12px;">
                   <div style="width:44px; height:44px; background:var(--accent-soft); border-radius:12px; display:flex; align-items:center; justify-content:center; color:var(--accent); font-size:20px;"><i class="ph-bold ph-calendar"></i></div>
                   <div style="flex:1">
                     <div style="font-size:14px; font-weight:600; color:white;">${formatDate(d.split(' ')[0])}${d.includes(' (') ? d.substring(d.indexOf(' (')) : ''}${topicRec ? ' <span style="font-size:10px; color:var(--accent); background:var(--accent-soft); padding:2px 6px; border-radius:4px;">📝 Topic</span>' : ''}</div>
                     <div style="font-size:11px; color:var(--text-3); margin-top:2px;">Present: ${pCount} | Absent: ${stats.total - pCount}</div>
                   </div>
                   <div style="text-align:right">
                     <div style="font-size:16px; font-weight:bold; color:var(--accent)">${pct}%</div>
                   </div>
                 </div>`;

      if (isExpanded) {
        const absentees = data.filter(r => r.date === d && r.status === 'A').map(r => r.rollNo).join(', ') || 'None';
        const sub = state.reportsSubject;
        const dateLabel = formatDate(d.split(' ')[0]) + (d.includes(' (') ? d.substring(d.indexOf(' (')) : '');
        const total = stats.total;
        const pct = ((pCount / total) * 100).toFixed(1);

        let summaryText = `📅 *Date*      : ${dateLabel}\n`;
        summaryText += `🎓 *Class*     : ${sub.program} · ${sub.year}\n`;
        summaryText += `📚 *Subject*   : ${sub.name} (${sub.code}) · ${sub.type}\n`;
        if (sub.type.toUpperCase() === 'PRACTICAL') {
          const batch = state.reportsBatch || data.find(r => r.date === d)?.batch || '';
          summaryText += `🧪 *Batch*     : ${batch}\n`;
        }
        summaryText += `🧑‍🏫 *Faculty*   : ${state.facultyName}\n`;
        if (topicRec) {
          summaryText += `📝 *Topic*     : ${topicRec.topic.trim()}\n`;
        }
        summaryText += `👥 *Attendance*: ${pCount} / ${total} students\n`;
        summaryText += `📊 *Percentage*: ${pct}%\n`;
        summaryText += `🚫 *Absent*    : ${absentees}`;

        html += `<div style="margin-top:16px; padding-top:16px; border-top:1px solid var(--border-color); animation: screenSlideIn 0.3s ease;">`;
        if (topicRec) {
          html += `<div style="font-size:12px; color:var(--accent); margin-bottom:8px;"><i class="ph-bold ph-note-blank" style="margin-right:4px;"></i>Topic: ${topicRec.topic.trim()}</div>`;
        }
        html += `<div style="font-size:12px; color:var(--text-3); margin-bottom:8px;">WhatsApp Message Preview:</div>
                   <div style="background:var(--bg-surface); padding:10px; border-radius:8px; font-size:12px; line-height:1.4; color:var(--text-2); font-family:monospace; white-space:pre-wrap;">${summaryText}</div>
                   <button class="btn btn-primary btn-sm btn-full" style="margin-top:12px; height:36px;" onclick="event.stopPropagation(); App.shareDateReport('${d}')">
                     <i class="ph-bold ph-whatsapp-logo"></i> Share to WhatsApp
                   </button>
                 </div>`;
      }

      html += `</div>`;
    });

    document.getElementById('rep-content').innerHTML = html;
  }

  function toggleDateDetails(date) {
    state.reportsExpandedDate = state.reportsExpandedDate === date ? null : date;
    renderReport();
  }

  function shareDateReport(date) {
    let sessionData = state.reportData.filter(r => r.date === date);

    // Filter by batch if practical
    if (state.reportsSubject.type.toUpperCase() === 'PRACTICAL' && state.reportsBatch) {
      sessionData = sessionData.filter(r => r.batch === state.reportsBatch);
    }

    const p = sessionData.filter(r => r.status === 'P').length;
    const t = sessionData.length;
    const pct = ((p / t) * 100).toFixed(1);
    const absentees = sessionData.filter(r => r.status === 'A').map(r => r.rollNo).join(', ') || 'None';

    const sub = state.reportsSubject;
    const dateLabel = formatDate(date.split(' ')[0]) + (date.includes(' (') ? date.substring(date.indexOf(' (')) : '');

    let msg = `📅 *Date*      : ${dateLabel}\n`;
    msg += `🎓 *Class*     : ${sub.program} · ${sub.year}\n`;
    msg += `📚 *Subject*   : ${sub.name} (${sub.code}) · ${sub.type}\n`;
    // Find any record for this date that has a topic
    const topicRec = sessionData.find(r => r.topic && r.topic.trim() !== '');
    if (topicRec) {
      msg += `📝 *Topic* : ${topicRec.topic.trim()}\n`;
    }
    if (sub.type.toUpperCase() === 'PRACTICAL') {
      msg += `🧪 *Batch*     : ${state.reportsBatch || (sessionData[0] && sessionData[0].batch)}\n`;
    }
    msg += `🧑‍🏫 *Faculty*   : ${state.facultyName}\n`;
    msg += `👥 *Attendance*: ${p} / ${t} students\n`;
    msg += `📊 *Percentage*: ${pct}%\n`;
    msg += `🚫 *Absent*    : ${absentees}\n`;

    const url = `whatsapp://send?text=${encodeURIComponent(msg)}`;
    window.location.href = url;

    // Direct Link ad (popunder replacement) — after WhatsApp, once per session
    setTimeout(openAdOnce, 1200);
  }

  function renderDefaulterReport(data) {
    const students = {};
    const limit = ((state.allData && state.allData.attendanceLimit) || 75);

    data.forEach(r => {
      if (!students[r.rollNo]) {
        students[r.rollNo] = { name: r.name, present: 0, total: 0 };
      }
      if (r.status === 'P') students[r.rollNo].present++;
      students[r.rollNo].total++;
    });

    let html = `<div style="margin-bottom:12px; font-size:12px; color:var(--danger); font-weight:600;">
                  ⚠️ Below ${limit}% Attendance
                </div>`;

    const defaulters = Object.keys(students).filter(roll => {
      const pct = Math.round((students[roll].present / students[roll].total) * 100);
      return pct < limit;
    });

    if (defaulters.length === 0) {
      document.getElementById('rep-content').innerHTML = `<div class="empty-state"><i class="ph-bold ph-smiley-wink empty-icon" style="color:var(--success)"></i><div class="empty-title">No Defaulters!</div><p style="color:var(--text-4);font-size:12px;">Everyone is above ${limit}%.</p></div>`;
      return;
    }

    defaulters.sort((a, b) => a - b).forEach(roll => {
      const s = students[roll];
      const pct = Math.round((s.present / s.total) * 100);

      html += `<div class="glass-card" style="padding:12px; margin-bottom:10px; border-left: 4px solid var(--danger); display:flex; align-items:center; gap:12px;">
                 <div style="width:36px; height:36px; background:var(--danger-bg); border-radius:50%; display:flex; align-items:center; justify-content:center; font-weight:bold; font-size:14px; color:var(--danger); border:1px solid var(--danger-border)">${roll}</div>
                 <div style="flex:1">
                   <div style="font-size:14px; font-weight:600; color:white;">${s.name}</div>
                   <div style="font-size:11px; color:var(--text-3); margin-top:2px;">Attended: ${s.present} / ${s.total}</div>
                 </div>
                 <div style="text-align:right">
                   <div style="font-size:16px; font-weight:bold; color:var(--danger)">${pct}%</div>
                 </div>
               </div>`;
    });

    document.getElementById('rep-content').innerHTML = html;
  }

  function downloadReport() {
    if (!state.reportsSubject || !state.reportData) return;

    // 1. Apply Filters
    let filtered = state.reportData;
    if (state.reportsSubject.type.toUpperCase() === 'PRACTICAL' && state.reportsBatch) {
      filtered = filtered.filter(r => r.batch === state.reportsBatch);
    }
    if (state.reportStartDate) filtered = filtered.filter(r => r.date >= state.reportStartDate);
    if (state.reportEndDate) filtered = filtered.filter(r => r.date <= state.reportEndDate);

    if (filtered.length === 0) return Toast.show('No data for selected filters', 'warning');

    // 2. Prepare Metadata
    const cfg = state.allData?.config || {};
    const sub = state.reportsSubject;
    const rawDates = [...new Set(filtered.map(r => r.date))].sort();
    const isPrac = sub.type.toUpperCase() === 'PRACTICAL';
    const limit = (state.allData && state.allData.attendanceLimit) || 75;

    // Dynamic naming logic
    const batchSuffix = (isPrac && state.reportsBatch) ? ` - ${state.reportsBatch}` : '';

    // SHEET NAME: Must be <= 31 chars and NO forbidden chars (: \ / ? * [ ])
    let sheetName = `${sub.code}${batchSuffix}`.replace(/[:\\/?*\[\]]/g, '');
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

    const dateRange = (state.reportStartDate && state.reportEndDate)
      ? `${formatDate(state.reportStartDate)} to ${formatDate(state.reportEndDate)}`
      : `01 Jan 2020 to ${formatDate(new Date().toISOString().split('T')[0])}`;

    // FILE NAME format: SubjectCode_Batch_DateRange.xlsx
    const batchFilePart = (isPrac && state.reportsBatch) ? `_${state.reportsBatch}` : '';
    const dateFilePart = (state.reportStartDate && state.reportEndDate)
      ? `${formatDate(state.reportStartDate).replace(/-/g, '')}_to_${formatDate(state.reportEndDate).replace(/-/g, '')}`
      : `All_Time`;
    const filename = `${sub.code}${batchFilePart}_${dateFilePart}.xlsx`.replace(/[\\/:*?"<>|]/g, '_');

    // 3. Build XLSX Spreadsheet via xlsx-js-style
    // Prefer full names from master config (context.config) — college sheet / license short form only as fallback
    const mName = (window.appStartContext && window.appStartContext.config && window.appStartContext.config.management_name)
      || cfg.managementName || (window.appStartContext && window.appStartContext.managementName) || "Management";
    const cName = (window.appStartContext && window.appStartContext.config && window.appStartContext.config.college_name)
      || cfg.collegeName || (window.appStartContext && window.appStartContext.collegeName) || "College";
    const metaStr = `${sub.code} - ${sub.name}${batchSuffix} | ${sub.program} | ${sub.year} | ${dateRange}`;

    if (typeof XLSX !== 'undefined') {
      const thinBorder = {
        top: { style: 'thin', color: { rgb: 'CBD5E1' } },
        bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
        left: { style: 'thin', color: { rgb: 'CBD5E1' } },
        right: { style: 'thin', color: { rgb: 'CBD5E1' } }
      };

      const styleTitle = {
        font: { sz: 14, bold: true, color: { rgb: '333333' } },
        alignment: { horizontal: 'center', vertical: 'center' }
      };

      const styleSubTitle = {
        font: { sz: 11, bold: true, color: { rgb: '333333' } },
        alignment: { horizontal: 'center', vertical: 'center' }
      };

      const styleMeta = {
        font: { sz: 10, bold: true, color: { rgb: '0F172A' } },
        fill: { fgColor: { rgb: 'E2E8F0' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder
      };

      const styleHeader = {
        font: { sz: 10, bold: true, color: { rgb: '0F172A' } },
        fill: { fgColor: { rgb: 'F1F5F9' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder
      };

      const styleNormal = {
        font: { sz: 10, color: { rgb: '1E293B' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder
      };

      const styleName = {
        font: { sz: 10, color: { rgb: '1E293B' } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: thinBorder
      };

      const stylePresent = {
        font: { sz: 10, color: { rgb: '15803D' } },
        fill: { fgColor: { rgb: 'DCFCE7' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder
      };

      const styleAbsent = {
        font: { sz: 10, color: { rgb: 'B91C1C' } },
        fill: { fgColor: { rgb: 'FEE2E2' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder
      };

      const stylePctGood = {
        font: { sz: 10, bold: true, color: { rgb: '14532D' } },
        fill: { fgColor: { rgb: 'BBF7D0' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder
      };

      const stylePctBad = {
        font: { sz: 10, bold: true, color: { rgb: '7F1D1D' } },
        fill: { fgColor: { rgb: 'FECACA' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder
      };

      const aoa = [];
      aoa.push([{ v: mName, s: styleTitle }]); // Row 0
      aoa.push([{ v: cName, s: styleSubTitle }]); // Row 1
      aoa.push([]); // Row 2
      aoa.push([{ v: metaStr, s: styleMeta }]); // Row 3
      aoa.push([]); // Row 4

      // Header row (Row 5)
      const headerRow = [
        { v: "Roll No.", s: styleHeader },
        { v: "Name", s: styleHeader }
      ];
      rawDates.forEach(d => {
        const baseDate = d.split(' ')[0];
        headerRow.push({ v: formatDate(baseDate), s: styleHeader });
      });
      headerRow.push(
        { v: "Total P", s: styleHeader },
        { v: "Total A", s: styleHeader },
        { v: "Total", s: styleHeader },
        { v: "% Att.", s: styleHeader }
      );
      aoa.push(headerRow);

      // Topic row (Row 6)
      const topicRow = [
        { v: "", s: styleHeader },
        { v: "Topic", s: styleHeader }
      ];
      rawDates.forEach(d => {
        const recForDate = filtered.find(r => r.date === d);
        topicRow.push({ v: recForDate ? recForDate.topic || '' : '', s: styleNormal });
      });
      topicRow.push(
        { v: "", s: styleNormal },
        { v: "", s: styleNormal },
        { v: "", s: styleNormal },
        { v: "", s: styleNormal }
      );
      aoa.push(topicRow);

      // Student rows (Row 7+)
      const students = {};
      filtered.forEach(r => {
        if (!students[r.rollNo]) students[r.rollNo] = { name: r.name, records: {} };
        students[r.rollNo].records[r.date] = r.status;
      });

      Object.keys(students).sort((a, b) => {
        const na = parseFloat(a), nb = parseFloat(b);
        return (isNaN(na) || isNaN(nb)) ? a.localeCompare(b) : na - nb;
      }).forEach(roll => {
        const s = students[roll];
        const studentRow = [
          { v: roll, s: styleNormal },
          { v: s.name, s: styleName }
        ];
        let pCount = 0, aCount = 0;
        rawDates.forEach(d => {
          const stat = s.records[d] || '-';
          let style = styleNormal;
          if (stat === 'P') { pCount++; style = stylePresent; }
          else if (stat === 'A') { aCount++; style = styleAbsent; }
          studentRow.push({ v: stat, s: style });
        });
        const total = pCount + aCount;
        const pct = total === 0 ? 0 : Math.round((pCount / total) * 100);
        studentRow.push({ v: pCount, s: styleNormal });
        studentRow.push({ v: aCount, s: styleNormal });
        studentRow.push({ v: rawDates.length, s: styleNormal });
        studentRow.push({ v: `${pct}.0%`, s: pct < limit ? stylePctBad : stylePctGood });
        aoa.push(studentRow);
      });

      const ws = XLSX.utils.aoa_to_sheet(aoa);

      const totalCols = rawDates.length + 6;
      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: totalCols - 1 } },
        { s: { r: 3, c: 0 }, e: { r: 3, c: totalCols - 1 } }
      ];

      const cols = [{ wch: 12 }, { wch: 28 }];
      rawDates.forEach(() => cols.push({ wch: 14 }));
      cols.push({ wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 });
      ws['!cols'] = cols;

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);

      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);

      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    // Direct Link ad
    try {
      window.open('https://omg10.com/4/11324927', '_blank', 'noopener,noreferrer');
    } catch (e) {
      console.warn("Direct Link blocked on download:", e);
    }

    Toast.show('Professional Report Generated');
  }



  // ─── UTILS ─────────────────────────────────────────

  function showModal(html, centered = true) {
    const backdrop = document.getElementById('modal-backdrop');
    const content = document.getElementById('modal-content');
    backdrop.className = 'modal-backdrop' + (centered ? ' modal-centered' : '');
    content.className = 'modal' + (centered ? ' modal-centered' : '');
    content.innerHTML = html;
    backdrop.style.display = 'flex';
  }

  function closeModal(e) {
    if (e && e.target !== document.getElementById('modal-backdrop')) return;
    document.getElementById('modal-backdrop').style.display = 'none';
  }

  let _spinnerWatchdog = null;

  function showSpinner(msg, iconClass = 'ph-cloud-arrow-down') {
    const overlay = document.getElementById('loader-overlay');
    const icon = document.getElementById('loader-icon');
    const text = document.getElementById('loader-text');

    if (overlay && icon && text) {
      icon.className = `ph-fill ${iconClass}`;
      text.innerText = msg;
      overlay.style.display = 'flex';
    }

    clearTimeout(_spinnerWatchdog);
    _spinnerWatchdog = setTimeout(() => {
      hideSpinner();
      if (window.Toast) Toast.show('Request timed out. Please tap and try again.', 'warning');
    }, 45000);
  }

  function hideSpinner() {
    clearTimeout(_spinnerWatchdog);
    const overlay = document.getElementById('loader-overlay');
    if (overlay) overlay.style.display = 'none';
  }

  const Toast = {
    show(msg, type = 'success') {
      const c = document.getElementById('toast-container');
      const b = document.getElementById('toast-backdrop');

      const t = document.createElement('div');
      t.className = `toast ${type}`;
      const iconMap = {
        success: 'ph-check-circle',
        error: 'ph-warning-circle',
        warning: 'ph-warning',
        info: 'ph-info'
      };
      const iconClass = iconMap[type] || iconMap.info;
      t.innerHTML = `<span class="toast-icon"><i class="ph-fill ${iconClass}"></i></span><span class="toast-msg">${msg}</span>`;

      c.appendChild(t);
      if (b) {
        b.style.display = 'block';
        setTimeout(() => b.classList.add('active'), 10);
      }

      setTimeout(() => {
        t.classList.add('hiding');
        setTimeout(() => {
          t.remove();
          if (c.children.length === 0 && b) {
            b.classList.remove('active');
            setTimeout(() => { b.style.display = 'none'; }, 300);
          }
        }, 350);
      }, 1500);
    }
  };

  window.Toast = Toast; // global exposure

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  function formatDate(isoDate) {
    const d = new Date(isoDate);
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-');
  }

  function dbFormatDate(isoDate) {
    return isoDate; // YYYY-MM-DD
  }



  return {
    initFromEngine, navigate, showLoginModal, showNameTray, pickLoginName, processLogin, logout,
    showDatePicker, setDate, showSubjectPicker, selectSubject,
    showSyllabusPicker,
    startAttendanceFlow, setupAttendanceUI, selectBatch,
    setEntryMode, toggleStudentStatus, markRollcall, saveAttendance, handleReturnToDashboard,
    openReports, switchReportTab, handleReportFilterChange, setReportRange, downloadReport,
    selectReportBatch, toggleDateDetails, shareDateReport, shareLastAttendance,
    showModal, closeModal
  };
})();

// Boot is now handled via appstart:complete event in index.html.
// No window.onload here — the engine calls AppStart.init() which
// fires the event when ready.

// Particle animation removed — AppStart engine handles boot visuals.
