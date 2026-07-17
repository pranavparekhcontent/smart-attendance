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
        try { rawData = JSON.parse(rawData); } catch(e) {}
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
        } catch(e){}
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
      navigate('faculty-dash');
    } else {
      navigate('login');
    }
  }

  // ─── NAVIGATION ────────────────────────────────────

  function loadDashboardAds() {
    if (document.getElementById('monetag-ad-script')) return;
    const script = document.createElement('script');
    script.id = 'monetag-ad-script';
    script.src = 'https://quge5.com/88/tag.min.js';
    script.setAttribute('data-zone', '260367');
    script.async = true;
    script.setAttribute('data-cfasync', 'false');
    document.head.appendChild(script);
  }

  function removeDashboardAds() {
    const script = document.getElementById('monetag-ad-script');
    if (script) script.remove();
    const adFrames = document.querySelectorAll('iframe[src*="quge5"], div[id*="monetag"], iframe[id*="monetag"]');
    adFrames.forEach(el => el.remove());
  }

  function navigate(screenId) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    const target = document.getElementById('screen-' + screenId);
    if (target) target.classList.add('active');
    state.currentScreen = screenId;
    window.scrollTo(0, 0);

    // Dynamic ad loading based on screen
    if (screenId === 'faculty-dash' || screenId === 'login') {
      loadDashboardAds();
    } else if (screenId === 'attendance-mode') {
      removeDashboardAds();
    }
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
    const validPins = expectedPin ? expectedPin.split(',').map(p=>p.trim()) : [];

    if (validPins.includes(pin)) {
      state.role = role;
      state.facultyName = name;
      localStorage.setItem('rmd_role', role);
      localStorage.setItem('rmd_faculty', name);
      
      closeModal();
      document.getElementById('dash-faculty-name').innerText = name;
      document.getElementById('dash-avatar').innerText = name.charAt(0).toUpperCase();
      document.getElementById('dash-date').innerText = formatDate(state.sessionDate);
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

  function showSyllabusPicker(points) {
    return new Promise((resolve) => {
      const pickerId = 'syl-picker-' + Date.now();
      let html = `<div class="modal-header">
                    <div class="modal-title">Select Topics</div>
                    <div class="modal-subtitle">Tap to select one or more syllabus points</div>
                  </div>
                  <div class="modal-body" style="padding: 16px 20px; max-height:50vh; overflow-y:auto;">
                    <div id="${pickerId}" style="display:flex; flex-wrap:wrap; gap:10px; justify-content:center;">`;

      // Render syllabus chips
      points.forEach((pt, idx) => {
        const safeVal = escapeHtml(pt);
        html += `<button type="button" class="syl-chip" data-idx="${idx}" data-value="${encodeURIComponent(pt)}" onclick="window._toggleSylChip(this)">
                   <span class="syl-chip-dot"></span>
                   <span class="syl-chip-text">${safeVal}</span>
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
                   <input type="text" id="syl-custom-input" class="input syl-custom-input" placeholder="Enter custom topic..." style="display:none;" autocomplete="off" />
                 </div>
               </div>
               <div class="modal-footer" style="padding: 10px 20px 20px; gap:10px;">
                 <button class="btn btn-glass" style="flex:1" onclick="window._cancelSyllabus()">Cancel</button>
                 <button class="btn btn-primary" style="flex:1.5; opacity:0.4; pointer-events:none;" id="syl-confirm-btn" onclick="window._confirmSyllabus()">
                   <i class="ph-bold ph-check-circle" style="margin-right:6px;"></i> Confirm
                 </button>
               </div>`;

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

      function updateConfirmBtn() {
        const selected = document.querySelectorAll(`#${pickerId} .syl-chip.selected`);
        const otherBtn = document.getElementById('syl-other-btn');
        const customVal = document.getElementById('syl-custom-input')?.value?.trim();
        const count = selected.length + ((otherBtn?.classList.contains('selected') && customVal) ? 1 : (otherBtn?.classList.contains('selected') ? 1 : 0));
        const btn = document.getElementById('syl-confirm-btn');
        if (selected.length > 0 || (otherBtn?.classList.contains('selected'))) {
          btn.style.opacity = '1';
          btn.style.pointerEvents = 'auto';
          btn.innerHTML = `<i class="ph-bold ph-check-circle" style="margin-right:6px;"></i> Confirm (${selected.length + (otherBtn?.classList.contains('selected') ? 1 : 0)})`;
        } else {
          btn.style.opacity = '0.4';
          btn.style.pointerEvents = 'none';
          btn.innerHTML = `<i class="ph-bold ph-check-circle" style="margin-right:6px;"></i> Confirm`;
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

  async function startAttendanceFlow() {
    if (!state.selectedSubject) return Toast.show('Please select a subject first', 'warning');
    
    let topic = '';
    let hasSyllabusPoints = false;
    let syllabusPoints = [];

    if (state.selectedSubject.teachingPlanLink) {
      showSpinner('Fetching Syllabus...', 'ph-book-open');
      try {
        const res = await API.getSyllabusPoints(state.selectedSubject.teachingPlanLink, state.selectedSubject.code);
        if (res && res.success && res.points && res.points.length > 0) {
          syllabusPoints = res.points;
          hasSyllabusPoints = true;
        }
      } catch (e) {
        console.warn('Error fetching syllabus points:', e);
      }
      hideSpinner();
    }

    if (hasSyllabusPoints) {
      const choice = await showSyllabusPicker(syllabusPoints);
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
    
    while(nextNum <= 12) {
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
      showSessionCompleteDialog(students.filter(s=>s.status==='P').length, students.filter(s=>s.status==='A').length);
    } else {
      Toast.show(res.error || 'Failed to save', 'error');
    }
  }

  function showSessionCompleteDialog(p, a) {
    let html = `<div class="session-complete-hero">
                  <div class="session-complete-icon"><i class="ph-fill ph-check-circle"></i></div>
                  <div class="session-complete-title">Session Saved!</div>
                  <div class="session-complete-sub">Attendance securely recorded.</div>
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

    // Direct Link ad (popunder replacement)
    try {
      window.open('https://omg10.com/4/11324927', '_blank', 'noopener,noreferrer');
    } catch (e) {
      console.warn("Direct Link blocked on WhatsApp share:", e);
    }

    const url = `https://wa.me/?text=${encodeURIComponent(msg)}`;
    window.location.href = url;
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
    
    showSpinner('Fetching Report Data...', 'ph-chart-line-up');
    const [res, studentRes] = await Promise.all([
      API.getAttendance(state.reportsSubject.code, state.reportsSubject.year, null, state.reportsSubject.outputSheetId),
      API.getStudents(state.reportsSubject.year)
    ]);
    hideSpinner();

    if (res.success) {
      state.reportData = res.records || [];
      
      // Dynamic Batch Detection — from student list as intended
      if (state.reportsSubject.type.toUpperCase() === 'PRACTICAL') {
        const studentBatches = [...new Set((studentRes.students || []).map(s => s.batch))].filter(b => b).sort();
        state.availableReportBatches = studentBatches;
        
        renderReportBatchSelector();
      }

      document.getElementById('reports-filters').style.display = 'block';
      document.getElementById('rep-footer').style.display = 'block';
      renderReport();
    } else {
      Toast.show(res.error || 'Failed to fetch reports', 'error');
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
    if(active) {
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

    Object.keys(students).sort((a,b) => a-b).forEach(roll => {
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

    const url = `https://wa.me/?text=${encodeURIComponent(msg)}`;
    
    // Direct Link ad (popunder replacement)
    try {
      window.open('https://omg10.com/4/11324927', '_blank', 'noopener,noreferrer');
    } catch (e) {
      console.warn("Direct Link blocked on WhatsApp share:", e);
    }

    window.location.href = url;
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

    defaulters.sort((a,b) => a-b).forEach(roll => {
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

    // FILE NAME format: SubjectCode_Batch_DateRange.xls
    const batchFilePart = (isPrac && state.reportsBatch) ? `_${state.reportsBatch}` : '';
    const dateFilePart = (state.reportStartDate && state.reportEndDate)
      ? `${formatDate(state.reportStartDate).replace(/-/g, '')}_to_${formatDate(state.reportEndDate).replace(/-/g, '')}`
      : `All_Time`;
    const filename = `${sub.code}${batchFilePart}_${dateFilePart}.xls`.replace(/[\\/:*?"<>|]/g, '_');

    // 3. Build XML Spreadsheet
    const escapeXml = (str) => str ? String(str).replace(/[<>&"']/g, (c) => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":"&apos;"}[c])) : "";
    const mName = escapeXml(cfg.managementName || (window.appStartContext && window.appStartContext.managementName) || "Management");
    const cName = escapeXml(cfg.collegeName || (window.appStartContext && window.appStartContext.collegeName) || "College");
    const metaStr = escapeXml(`${sub.code} - ${sub.name}${batchSuffix} | ${sub.program} | ${sub.year} | ${dateRange}`);
    
    // Total Columns = Roll(1) + Name(1) + Dates(N) + TotP(1) + TotA(1) + Total(1) + %Att(1) = N + 6
    const mergeVal = rawDates.length + 5; 

    // Helper for Borders (ss:Position="All" is NOT valid SpreadsheetML)
    const borderXml = `
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>`;

    let xml = `<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
    <Styles>
      <Style ss:ID="Title"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:Bold="1" ss:Size="14" ss:Color="#333333"/></Style>
      <Style ss:ID="SubTitle"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:Bold="1" ss:Size="11" ss:Color="#333333"/></Style>
      <Style ss:ID="MetaRow"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:Bold="1" ss:Size="10" ss:Color="#0F172A"/><Interior ss:Color="#E2E8F0" ss:Pattern="Solid"/>${borderXml}</Style>
      <Style ss:ID="Header"><Font ss:Bold="1" ss:Color="#0F172A"/><Interior ss:Color="#F1F5F9" ss:Pattern="Solid"/>${borderXml}<Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
      <Style ss:ID="Normal">${borderXml}<Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
      <Style ss:ID="NameCell">${borderXml}<Alignment ss:Horizontal="Left" ss:Vertical="Center"/></Style>
      <Style ss:ID="PresentCell">${borderXml}<Font ss:Color="#15803D"/><Interior ss:Color="#DCFCE7" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
      <Style ss:ID="AbsentCell">${borderXml}<Font ss:Color="#B91C1C"/><Interior ss:Color="#FEE2E2" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
      <Style ss:ID="PctGood">${borderXml}<Font ss:Color="#14532D" ss:Bold="1"/><Interior ss:Color="#BBF7D0" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
      <Style ss:ID="PctBad">${borderXml}<Font ss:Color="#7F1D1D" ss:Bold="1"/><Interior ss:Color="#FECACA" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
    </Styles>
    <Worksheet ss:Name="${escapeXml(sheetName)}">
    <Table>
      <Column ss:Width="60"/><Column ss:Width="250"/>`;
    
    rawDates.forEach(() => xml += `<Column ss:Width="80"/>`);
    xml += `<Column ss:Width="60"/><Column ss:Width="60"/><Column ss:Width="60"/><Column ss:Width="60"/>`;

    xml += `
      <Row ss:Height="25"><Cell ss:MergeAcross="${mergeVal}" ss:StyleID="Title"><Data ss:Type="String">${mName}</Data></Cell></Row>
      <Row ss:Height="20"><Cell ss:MergeAcross="${mergeVal}" ss:StyleID="SubTitle"><Data ss:Type="String">${cName}</Data></Cell></Row>
      <Row ss:Height="10"><Cell ss:MergeAcross="${mergeVal}"><Data ss:Type="String"></Data></Cell></Row>
      <Row ss:Height="20"><Cell ss:MergeAcross="${mergeVal}" ss:StyleID="MetaRow"><Data ss:Type="String">${metaStr}</Data></Cell></Row>
      <Row ss:Height="10"><Cell ss:MergeAcross="${mergeVal}"><Data ss:Type="String"></Data></Cell></Row>
      <Row ss:Height="22">
        <Cell ss:StyleID="Header"><Data ss:Type="String">Roll No.</Data></Cell>
        <Cell ss:StyleID="Header"><Data ss:Type="String">Name</Data></Cell>`;
    
    rawDates.forEach(d => {
      const baseDate = d.split(' ')[0];
      xml += `<Cell ss:StyleID="Header"><Data ss:Type="String">${formatDate(baseDate)}</Data></Cell>`;
    });
    
    xml += `<Cell ss:StyleID="Header"><Data ss:Type="String">Total P</Data></Cell>
            <Cell ss:StyleID="Header"><Data ss:Type="String">Total A</Data></Cell>
            <Cell ss:StyleID="Header"><Data ss:Type="String">Total</Data></Cell>
            <Cell ss:StyleID="Header"><Data ss:Type="String">% Att.</Data></Cell></Row>`;

    xml += `
      <Row ss:Height="20">
        <Cell ss:StyleID="Header"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="Header"><Data ss:Type="String">Topic</Data></Cell>`;
    
    rawDates.forEach(d => {
      const recForDate = filtered.find(r => r.date === d);
      const topic = recForDate ? recForDate.topic || '' : '';
      xml += `<Cell ss:StyleID="Normal"><Data ss:Type="String">${escapeXml(topic)}</Data></Cell>`;
    });
    
    xml += `
        <Cell ss:StyleID="Normal"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="Normal"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="Normal"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="Normal"><Data ss:Type="String"></Data></Cell>
      </Row>`;

    const students = {};
    filtered.forEach(r => {
      if (!students[r.rollNo]) students[r.rollNo] = { name: r.name, records: {} };
      students[r.rollNo].records[r.date] = r.status;
    });

    Object.keys(students).sort((a,b) => {
      const na = parseFloat(a), nb = parseFloat(b);
      return (isNaN(na) || isNaN(nb)) ? a.localeCompare(b) : na - nb;
    }).forEach(roll => {
      const s = students[roll];
      xml += `<Row ss:Height="18">`;
      xml += `<Cell ss:StyleID="Normal"><Data ss:Type="String">${escapeXml(roll)}</Data></Cell>`;
      xml += `<Cell ss:StyleID="NameCell"><Data ss:Type="String">${escapeXml(s.name)}</Data></Cell>`;
      
      let pCount = 0, aCount = 0;
      rawDates.forEach(d => {
        const stat = s.records[d] || '-';
        if (stat === 'P') pCount++;
        if (stat === 'A') aCount++;
        const style = stat === 'P' ? 'PresentCell' : (stat === 'A' ? 'AbsentCell' : 'Normal');
        xml += `<Cell ss:StyleID="${style}"><Data ss:Type="String">${stat}</Data></Cell>`;
      });
      
      const total = pCount + aCount;
      const pct = total === 0 ? 0 : Math.round((pCount / total) * 100);
      xml += `<Cell ss:StyleID="Normal"><Data ss:Type="Number">${pCount}</Data></Cell>`;
      xml += `<Cell ss:StyleID="Normal"><Data ss:Type="Number">${aCount}</Data></Cell>`;
      xml += `<Cell ss:StyleID="Normal"><Data ss:Type="Number">${rawDates.length}</Data></Cell>`;
      xml += `<Cell ss:StyleID="${pct < limit ? 'PctBad' : 'PctGood'}"><Data ss:Type="String">${pct}.0%</Data></Cell>`;
      xml += `</Row>`;
    });

    xml += `</Table></Worksheet></Workbook>`;

    const blob = new Blob([xml], { type: 'application/vnd.ms-excel' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

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

  function showSpinner(msg, iconClass = 'ph-cloud-arrow-down') {
    const overlay = document.getElementById('loader-overlay');
    const icon = document.getElementById('loader-icon');
    const text = document.getElementById('loader-text');
    
    if (overlay && icon && text) {
      icon.className = `ph-fill ${iconClass}`;
      text.innerText = msg;
      overlay.style.display = 'flex';
    }
  }

  function hideSpinner() {
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
