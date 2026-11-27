(() => {
  'use strict';

  // ==========================================================================
  // 1. State & Constants
  // ==========================================================================
  const TOKEN_KEY = 'ckp_auth_token_v3';
  const AUTO_SYNC_KEY = 'ckp_auto_sync_v3';
  const BOUND_BILLING_SITE_KEY = 'ckp_bound_billing_site_v3';
  const IDB_NAME = 'ckp_bound_excel_v3';
  const IDB_STORE = 'handles';

  const state = {
    token: localStorage.getItem(TOKEN_KEY) || '',
    user: null,
    storageMode: '',
    quota: null,
    activeView: 'front', // 'front' | 'back'
    activeBoTab: 'vehicles', // 'vehicles' | 'workers' | 'sites' | 'routes'

    // Master data (starts empty — No Seed Data)
    vehicles: [],
    workers: [],
    sites: [],
    routes: [],

    // Travel logs (filtered)
    logs: [],
    todayLogsCount: 0,

    // Form state
    editingLogId: null,
    isSwapped: false,
    routeMetaExpanded: false,
    searchTimer: null,
    newSiteOrigins: [],

    // Bound local Excel file handles (File System Access API + IndexedDB)
    boundZoneHandle: null,
    boundBillingHandle: null,
    boundBillingSite: localStorage.getItem(BOUND_BILLING_SITE_KEY) || '',
    boundZoneHandlesBySite: {}, // { [site_name]: FileSystemFileHandle }
    boundBillingHandlesBySite: {}, // { [site_name]: FileSystemFileHandle }
    excelSitesShowAll: false,
    gsheetSitesShowAll: false,
    autoSync: localStorage.getItem(AUTO_SYNC_KEY) !== 'false',

    // Bound Google Sheets config (stored on server / Supabase for PC & mobile)
    gsheet: {
      zone_url: '',
      zone_name: '',
      zone_sheet_url: '',
      zone_by_site: {},
      billing_url: '',
      billing_name: '',
      billing_sheet_url: '',
      billing_site: '',
      billing_by_site: {},
      auto_sync: true
    },
    gsheetHistory: [],
    gsheetHistoryCount: 0,
    gsheetLastHistory: null
  };

  const $ = (id) => document.getElementById(id);

  // ==========================================================================
  // 2. Formatting & Utility Helpers
  // ==========================================================================
  function escapeHtml(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function todayIso() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function currentMonthIso() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }

  function toBE(iso) {
    if (window.CkpExcel && typeof window.CkpExcel.toBE === 'function') {
      return window.CkpExcel.toBE(iso);
    }
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? `${Number(m[3])}/${Number(m[2])}/${Number(m[1]) + 543}` : '';
  }

  function formatMoney(n) {
    const num = Number(n) || 0;
    return num.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatInt(n) {
    return (Number(n) || 0).toLocaleString('th-TH');
  }

  function calcDurationText(dep, arr) {
    const parse = (t) => {
      const s = String(t || '').trim();
      let m = s.match(/^(\d{1,2})[:.](\d{2})$/);
      if (!m) m = s.match(/^(\d{2})(\d{2})$/);
      return m ? Number(m[1]) * 60 + Number(m[2]) : null;
    };
    const d1 = parse(dep);
    const d2 = parse(arr);
    if (d1 === null || d2 === null) return '';
    let diff = d2 - d1;
    if (diff < 0) diff += 24 * 60;
    if (diff === 0) return '0 นาที';
    const hrs = Math.floor(diff / 60);
    const mins = diff % 60;
    if (hrs === 0) return `${mins} นาที`;
    if (mins === 0) return `${hrs} ชม.`;
    return `${hrs} ชม. ${mins} นาที`;
  }

  function normalizeThaiDigits(str) {
    return String(str || '').replace(/[๐-๙]/g, (ch) => String('๐๑๒๓๔๕๖๗๘๙'.indexOf(ch)));
  }

  function formatAutoColonTimeInput(raw, isDeleting) {
    const cleaned = normalizeThaiDigits(raw)
      .replace(/[.,;：]/g, ':')
      .replace(/[^0-9:]/g, '');

    if (!cleaned) return '';
    if (isDeleting) return cleaned.slice(0, 5);

    if (cleaned.includes(':')) {
      const idx = cleaned.indexOf(':');
      const h = cleaned.slice(0, idx).replace(/\D/g, '').slice(0, 2);
      const m = cleaned.slice(idx + 1).replace(/\D/g, '').slice(0, 2);
      if (!h && !m) return '';
      return `${h}:${m}`;
    }

    const digits = cleaned.slice(0, 4);
    if (digits.length === 1) return digits;
    if (digits.length === 2) {
      // ถ้าสองหลักเกิน 23 เช่น 63 (จาก 6:30) ให้จัดเป็น 06:3 ทันที, ถ้า 00-23 ให้เติม : ต่อท้ายทันที
      if (Number(digits) > 23) return `0${digits[0]}:${digits[1]}`;
      return `${digits}:`;
    }
    if (digits.length === 3) {
      if (Number(digits.slice(0, 2)) > 23) return `0${digits[0]}:${digits.slice(1, 3)}`;
      return `${digits.slice(0, 2)}:${digits[2]}`;
    }
    return `${digits.slice(0, 2)}:${digits.slice(2, 4)}`;
  }

  function finalizeAutoColonTime(raw) {
    const cleaned = normalizeThaiDigits(raw)
      .trim()
      .replace(/[.,;：]/g, ':')
      .replace(/[^0-9:]/g, '');
    if (!cleaned) return '';

    let h = '';
    let m = '';
    if (cleaned.includes(':')) {
      const idx = cleaned.indexOf(':');
      h = cleaned.slice(0, idx).replace(/\D/g, '').slice(0, 2);
      m = cleaned.slice(idx + 1).replace(/\D/g, '').slice(0, 2);
      if (!h) return '';
      if (!m) m = '00';
      else if (m.length === 1) return `${h.padStart(2, '0')}:${m}`;
    } else {
      const digits = cleaned.slice(0, 4);
      if (digits.length === 1 || digits.length === 2) {
        if (digits.length === 2 && Number(digits) > 23) {
          return `0${digits[0]}:${digits[1]}`;
        }
        h = digits;
        m = '00';
      } else if (digits.length === 3) {
        h = digits[0];
        m = digits.slice(1, 3);
      } else {
        h = digits.slice(0, 2);
        m = digits.slice(2, 4);
      }
    }

    const hn = Number(h);
    const mn = Number(m);
    if (hn <= 23 && mn <= 59) {
      return `${String(hn).padStart(2, '0')}:${String(mn).padStart(2, '0')}`;
    }
    return cleaned.slice(0, 5);
  }

  function attachAutoColonTimeInput(el, onUpdate) {
    if (!el || el.dataset.autoColonBound === '1') return;
    el.dataset.autoColonBound = '1';

    el.addEventListener('input', (e) => {
      const isDeleting = Boolean(e && e.inputType && e.inputType.startsWith('delete'));
      const formatted = formatAutoColonTimeInput(el.value, isDeleting);
      if (el.value !== formatted) {
        el.value = formatted;
      }
      if (typeof onUpdate === 'function') onUpdate();
    });

    el.addEventListener('blur', () => {
      const finalized = finalizeAutoColonTime(el.value);
      if (el.value !== finalized) {
        el.value = finalized;
      }
      if (typeof onUpdate === 'function') onUpdate();
    });
  }

  function pulseElement(el) {
    if (!el) return;
    el.classList.remove('input-highlight-pulse');
    void el.offsetWidth;
    el.classList.add('input-highlight-pulse');
  }

  // ==========================================================================
  // 3. Toast Notifications & Custom Modals (Zero browser alert/confirm)
  // ==========================================================================
  function showToast(message, type = 'info', duration = 3600) {
    const container = $('toast-container');
    if (!container) return;

    const palette = {
      success: 'bg-emerald-900 text-white border-emerald-700',
      error: 'bg-rose-900 text-white border-rose-700',
      warning: 'bg-amber-900 text-amber-50 border-amber-700',
      info: 'bg-zinc-900 text-white border-zinc-700'
    };
    const icons = {
      success: 'fa-circle-check text-emerald-400',
      error: 'fa-circle-exclamation text-rose-400',
      warning: 'fa-triangle-exclamation text-amber-400',
      info: 'fa-circle-info text-sky-400'
    };

    const toast = document.createElement('div');
    toast.className = `pointer-events-auto toast-animate flex items-start gap-2.5 px-4 py-3 rounded-xl border shadow-xl text-xs sm:text-sm ${palette[type] || palette.info}`;
    toast.innerHTML = `
      <i class="fa-solid ${icons[type] || icons.info} mt-0.5 shrink-0"></i>
      <span class="flex-1 leading-snug">${escapeHtml(message)}</span>
      <button type="button" class="text-white/60 hover:text-white ml-1 cursor-pointer" aria-label="ปิด">
        <i class="fa-solid fa-xmark"></i>
      </button>
    `;
    toast.querySelector('button').addEventListener('click', () => toast.remove());
    container.appendChild(toast);

    setTimeout(() => {
      if (toast.isConnected) toast.remove();
    }, duration);
  }

  function closeModal() {
    const backdrop = $('custom-modal-backdrop');
    const box = $('custom-modal-box');
    if (backdrop) {
      backdrop.classList.add('hidden');
      backdrop.classList.remove('flex');
    }
    if (box) box.innerHTML = '';
  }

  function openModal(html, onReady) {
    const backdrop = $('custom-modal-backdrop');
    const box = $('custom-modal-box');
    if (!backdrop || !box) return;
    box.innerHTML = html;
    backdrop.classList.remove('hidden');
    backdrop.classList.add('flex');
    if (typeof onReady === 'function') onReady(box);
  }

  function confirmModal({ title, message, confirmText = 'ยืนยัน', cancelText = 'ยกเลิก', danger = false }) {
    return new Promise((resolve) => {
      const btnCls = danger
        ? 'bg-rose-600 hover:bg-rose-700 text-white'
        : 'bg-zinc-900 hover:bg-zinc-800 text-white';
      const iconCls = danger
        ? 'bg-rose-100 text-rose-600 fa-triangle-exclamation'
        : 'bg-indigo-100 text-indigo-600 fa-circle-question';

      openModal(
        `
        <div class="p-5 sm:p-6 space-y-4">
          <div class="flex items-start gap-3">
            <div class="w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${iconCls.split(' ').slice(0, 2).join(' ')}">
              <i class="fa-solid ${iconCls.split(' ')[2]}"></i>
            </div>
            <div class="flex-1">
              <h3 class="text-base font-bold text-slate-900">${escapeHtml(title)}</h3>
              <p class="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">${escapeHtml(message)}</p>
            </div>
          </div>
          <div class="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="modal-btn-cancel" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs sm:text-sm font-medium cursor-pointer">${escapeHtml(cancelText)}</button>
            <button type="button" id="modal-btn-confirm" class="min-h-[40px] px-4 py-2 rounded-xl ${btnCls} text-xs sm:text-sm font-semibold cursor-pointer">${escapeHtml(confirmText)}</button>
          </div>
        </div>
        `,
        (box) => {
          box.querySelector('#modal-btn-cancel').addEventListener('click', () => {
            closeModal();
            resolve(false);
          });
          box.querySelector('#modal-btn-confirm').addEventListener('click', () => {
            closeModal();
            resolve(true);
          });
        }
      );
    });
  }

  // ==========================================================================
  // 4. API Helper
  // ==========================================================================
  async function apiFetch(url, options = {}) {
    const headers = { ...(options.headers || {}) };
    if (options.body && !headers['Content-Type']) {
      headers['Content-Type'] = 'application/json';
    }
    if (state.token) {
      headers.Authorization = `Bearer ${state.token}`;
    }

    const res = await fetch(url, { ...options, headers });
    let data = null;
    try {
      data = await res.json();
    } catch {
      data = null;
    }

    if (res.status === 401 && url !== '/api/login') {
      handleForceLogout('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่อีกครั้ง');
      throw new Error((data && data.error) || 'Unauthorized');
    }
    if (!res.ok) {
      const err = new Error((data && data.error) || `เกิดข้อผิดพลาด (${res.status})`);
      err.status = res.status;
      err.data = data;
      throw err;
    }
    return data;
  }

  // ==========================================================================
  // 5. Auth & Navigation
  // ==========================================================================
  function showLoginScreen() {
    $('app-screen').classList.add('hidden');
    $('app-screen').classList.remove('flex');
    $('login-screen').classList.remove('hidden');
    $('login-screen').classList.add('flex');
  }

  function showAppScreen() {
    $('login-screen').classList.add('hidden');
    $('login-screen').classList.remove('flex');
    $('app-screen').classList.remove('hidden');
    $('app-screen').classList.add('flex');

    const isAdmin = state.user && state.user.role === 'admin';
    $('nav-user-fullname').textContent = state.user ? state.user.full_name : '';

    const badge = $('nav-user-role-badge');
    if (isAdmin) {
      badge.className =
        'inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-amber-100 text-amber-900 border border-amber-300';
      badge.innerHTML = '<i class="fa-solid fa-crown text-[10px] text-amber-600"></i><span>แอดมิน</span>';
    } else {
      badge.className =
        'inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-sky-100 text-sky-900 border border-sky-200';
      badge.innerHTML = '<i class="fa-solid fa-user-pen text-[10px]"></i><span>ผู้บันทึกข้อมูล</span>';
    }

    // Toggle Admin controls
    const toggleAdminEl = (id, displayCls = 'flex') => {
      const el = $(id);
      if (!el) return;
      if (isAdmin) {
        el.classList.remove('hidden');
        el.classList.add(displayCls);
      } else {
        el.classList.add('hidden');
        el.classList.remove(displayCls);
      }
    };
    toggleAdminEl('nav-tab-back', 'flex');
    toggleAdminEl('mobile-nav-back', 'flex');
    toggleAdminEl('btn-clear-all-logs', 'flex');
    toggleAdminEl('btn-goto-backoffice-hint', 'flex');

    switchView('front');
  }

  function handleForceLogout(msg) {
    state.token = '';
    state.user = null;
    localStorage.removeItem(TOKEN_KEY);
    showLoginScreen();
    if (msg) showToast(msg, 'warning');
  }

  function switchView(view) {
    if (view === 'back' && (!state.user || state.user.role !== 'admin')) {
      showToast('เฉพาะแอดมินเท่านั้นที่เข้าเมนูตั้งค่าได้ — กรุณาแจ้งแอดมิน', 'warning');
      return;
    }
    state.activeView = view;
    const isFront = view === 'front';

    $('view-front-office').classList.toggle('hidden', !isFront);
    $('view-back-office').classList.toggle('hidden', isFront);

    // Desktop nav button styles
    const activeDesk = 'bg-zinc-900 text-white shadow-sm';
    const inactiveDesk = 'text-zinc-600 hover:text-zinc-900';
    $('nav-tab-front').className = `min-h-[40px] px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-semibold transition flex items-center gap-1.5 cursor-pointer ${isFront ? activeDesk : inactiveDesk}`;
    if (state.user && state.user.role === 'admin') {
      $('nav-tab-back').className = `min-h-[40px] px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-semibold transition flex items-center gap-1.5 cursor-pointer ${!isFront ? activeDesk : inactiveDesk}`;
    }

    // Mobile nav button styles
    $('mobile-nav-front').className = `flex-1 min-h-[44px] rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 cursor-pointer ${isFront ? 'bg-zinc-900 text-white' : 'bg-slate-100 text-slate-700'}`;
    if (state.user && state.user.role === 'admin') {
      $('mobile-nav-back').className = `flex-1 min-h-[44px] rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 cursor-pointer ${!isFront ? 'bg-zinc-900 text-white' : 'bg-slate-100 text-slate-700'}`;
    }
  }

  function switchBoTab(tab) {
    state.activeBoTab = tab;
    const tabs = ['vehicles', 'workers', 'sites', 'routes'];
    for (const t of tabs) {
      const btn = $(`bo-tab-${t}`);
      const panel = $(`bo-panel-${t}`);
      const countBadge = $(`count-${t}`);
      const active = t === tab;
      if (panel) panel.classList.toggle('hidden', !active);
      if (btn) {
        btn.className = `min-h-[40px] px-4 py-2 rounded-xl text-xs sm:text-sm font-semibold transition flex items-center gap-2 cursor-pointer ${
          active ? 'bg-zinc-900 text-white shadow-sm' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
        }`;
      }
      if (countBadge) {
        countBadge.className = `px-2 py-0.5 rounded-full text-[11px] ${
          active ? 'bg-white/15 text-amber-300' : 'bg-slate-200 text-slate-700'
        }`;
      }
    }
  }

  // ==========================================================================
  // 6. Data Loading & Status Rendering
  // ==========================================================================
  function renderDbStatusBanner(quota) {
    const mode = (quota && quota.storageMode) || state.storageMode;
    const dot = $('db-mode-dot');
    const text = $('db-mode-text');
    if (!text) return;

    if (mode === 'supabase_cloud') {
      if (dot) dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
      text.innerHTML =
        '<strong>สถานะฐานข้อมูล: เชื่อมต่อ Supabase PostgreSQL Cloud เรียบร้อยแล้ว</strong> — บันทึกและซิงค์ข้อมูลออนไลน์แบบเรียลไทม์';
    } else {
      if (dot) dot.className = 'w-2 h-2 rounded-full bg-amber-400 animate-pulse';
      text.innerHTML =
        '<strong>สถานะฐานข้อมูล: ฐานข้อมูล SQLite ในเครื่อง (database.sqlite)</strong> — บันทึกข้อมูลถาวรในเครื่อง';
    }
  }

  function renderQuotaUI(quota) {
    if (!quota) return;
    state.quota = quota;
    renderDbStatusBanner(quota);

    $('stat-total-logs').textContent = formatInt(quota.currentCount);
    $('stat-max-quota').textContent = formatInt(quota.maxLimit);

    const pct = Math.min(100, Number(quota.usagePercent) || 0);
    const label = $('quota-progress-label');
    const bar = $('quota-progress-bar');
    const badge = $('quota-status-badge');
    const limitInput = $('input-max-storage-limit');
    if (label) {
      label.textContent = `ใช้งานแล้ว ${formatInt(quota.currentCount)} / ${formatInt(quota.maxLimit)} รายการ (${pct}%)`;
    }
    if (bar) {
      bar.style.width = `${pct}%`;
      bar.className = `h-full rounded-full transition-all duration-300 ${
        quota.isFull ? 'bg-rose-600' : quota.isWarning ? 'bg-amber-500' : 'bg-indigo-600'
      }`;
    }
    if (badge) {
      if (quota.isFull) {
        badge.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-rose-100 text-rose-800';
        badge.textContent = 'เต็มโควต้า (100%)';
      } else if (quota.isWarning) {
        badge.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800';
        badge.textContent = 'ใกล้เต็ม (>= 80%)';
      } else {
        badge.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800';
        badge.textContent = 'ปกติ';
      }
    }
    if (limitInput && document.activeElement !== limitInput) {
      limitInput.value = quota.maxLimit;
    }

    const banner = $('storage-warning-banner');
    if (banner) {
      if (quota.isFull) {
        banner.className = 'rounded-2xl p-4 border shadow-sm bg-rose-50 border-rose-300 text-rose-900 text-xs sm:text-sm flex items-center gap-3';
        banner.innerHTML = `
          <i class="fa-solid fa-circle-exclamation text-rose-600 text-lg shrink-0"></i>
          <div><strong>พื้นที่จัดเก็บข้อมูลเต็มแล้ว (${formatInt(quota.currentCount)}/${formatInt(quota.maxLimit)} รายการ)</strong> — กรุณาส่งออกไฟล์ Excel สำรองข้อมูลแล้วล้างข้อมูลเก่า หรือเพิ่มขีดจำกัดในเมนูตั้งค่า</div>
        `;
      } else if (quota.isWarning) {
        banner.className = 'rounded-2xl p-4 border shadow-sm bg-amber-50 border-amber-300 text-amber-900 text-xs sm:text-sm flex items-center gap-3';
        banner.innerHTML = `
          <i class="fa-solid fa-triangle-exclamation text-amber-600 text-lg shrink-0"></i>
          <div><strong>แจ้งเตือนพื้นที่จัดเก็บข้อมูลใกล้เต็ม (${pct}%)</strong> — ใช้งานแล้ว ${formatInt(quota.currentCount)} จาก ${formatInt(quota.maxLimit)} รายการ แนะนำให้ส่งออกไฟล์ Excel สำรองข้อมูลไว้</div>
        `;
      } else {
        banner.className = 'hidden rounded-2xl p-4 border shadow-sm transition-all';
        banner.innerHTML = '';
      }
    }
  }

  async function loadMasterData() {
    const [vehicles, workers, sites, routes, quota] = await Promise.all([
      apiFetch('/api/master/vehicles'),
      apiFetch('/api/master/workers'),
      apiFetch('/api/master/sites'),
      apiFetch('/api/master/routes'),
      apiFetch('/api/storage-status')
    ]);
    state.vehicles = vehicles || [];
    state.workers = workers || [];
    state.sites = sites || [];
    state.routes = routes || [];

    renderQuotaUI(quota);
    renderMasterSummaryAndHint();
    populateFrontDropdowns();
    renderBackOfficeLists();
    updateBindingUI();
    renderGoogleSheetsUI();
  }

  function getFilterParams() {
    return {
      site_name: $('filter-site').value.trim(),
      trip_type: $('filter-trip-type').value.trim(),
      date_from: $('filter-date-from').value.trim(),
      date_to: $('filter-date-to').value.trim(),
      search: $('filter-search').value.trim()
    };
  }

  async function loadLogs() {
    const params = new URLSearchParams();
    const f = getFilterParams();
    for (const [k, v] of Object.entries(f)) {
      if (v) params.set(k, v);
    }
    const qs = params.toString();
    const today = todayIso();

    const [logs, todayLogs] = await Promise.all([
      apiFetch(`/api/logs${qs ? `?${qs}` : ''}`),
      apiFetch(`/api/logs?date_from=${encodeURIComponent(today)}&date_to=${encodeURIComponent(today)}`)
    ]);

    state.logs = logs || [];
    state.todayLogsCount = (todayLogs || []).length;
    renderStepGuide();
    renderLogs();
    renderFilterSummaryLabel();
    updateBindingUI();
    renderGoogleSheetsUI();
  }

  async function refreshAll() {
    await Promise.all([loadMasterData(), loadGoogleSheetsConfig()]);
    await loadLogs();
  }

  function renderMasterSummaryAndHint() {
    const activeW = state.workers.filter((x) => x.active).length;
    const activeV = state.vehicles.filter((x) => x.active).length;
    const activeS = state.sites.filter((x) => x.active).length;
    const activeR = state.routes.filter((x) => x.active).length;

    $('stat-master-summary').textContent = `คนขับ ${activeW} • รถ ${activeV} • โครงการ ${activeS} • เส้นทาง ${activeR}`;
    $('count-vehicles').textContent = String(state.vehicles.length);
    $('count-workers').textContent = String(state.workers.length);
    $('count-sites').textContent = String(state.sites.length);
    $('count-routes').textContent = String(state.routes.length);

    const missing = [];
    if (activeV === 0) missing.push('ทะเบียนรถ');
    if (activeW === 0) missing.push('รายชื่อคนขับ');
    if (activeS === 0) missing.push('ไซต์งาน/โครงการ');
    if (activeR === 0) missing.push('เส้นทาง');

    const hintBox = $('empty-master-hint');
    const hintText = $('empty-master-hint-text');
    const isAdmin = Boolean(state.user && state.user.role === 'admin');
    if (missing.length > 0) {
      hintBox.classList.remove('hidden');
      hintText.textContent = isAdmin
        ? `ยังขาด: ${missing.join(', ')} — กดปุ่ม "เริ่มตั้งค่า" ทางขวาเพื่อเพิ่ม`
        : `ยังขาด: ${missing.join(', ')} — กรุณาแจ้งแอดมินให้เพิ่มข้อมูลก่อน`;
    } else {
      hintBox.classList.add('hidden');
    }
    renderStepGuide(missing);
  }

  // แถบ 3 ขั้นตอนการใช้งาน (ตั้งค่า → บันทึกเที่ยววิ่ง → ออกเอกสาร)
  function renderStepGuide(missing) {
    const s1 = $('step-1-status');
    const n1 = $('step-1-num');
    if (s1 && Array.isArray(missing)) {
      if (missing.length === 0) {
        s1.className = 'block text-xs font-semibold mt-1 text-emerald-700';
        s1.innerHTML = '<i class="fa-solid fa-circle-check mr-1"></i>ตั้งค่าครบแล้ว';
        if (n1) {
          n1.className = 'w-9 h-9 rounded-full bg-emerald-500 text-white flex items-center justify-center text-base font-bold shrink-0';
          n1.innerHTML = '<i class="fa-solid fa-check"></i>';
        }
      } else {
        s1.className = 'block text-xs font-semibold mt-1 text-amber-700';
        s1.textContent = `ยังขาด: ${missing.join(', ')}`;
        if (n1) {
          n1.className = 'w-9 h-9 rounded-full bg-amber-100 text-amber-800 flex items-center justify-center text-base font-bold shrink-0';
          n1.textContent = '1';
        }
      }
    }
    const s2 = $('step-2-status');
    if (s2) {
      const n = Number(state.todayLogsCount) || 0;
      s2.textContent = n > 0 ? `วันนี้บันทึกแล้ว ${formatInt(n)} เที่ยว` : 'วันนี้ยังไม่ได้บันทึก';
    }
  }

  function scrollToEl(id) {
    const el = $(id);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function populateSelect(selectEl, options, placeholder, preserveValue = true) {
    if (!selectEl) return;
    const prev = preserveValue ? selectEl.value : '';
    const html = [`<option value="">${escapeHtml(placeholder)}</option>`];
    let matched = false;
    for (const opt of options) {
      const val = String(opt.value);
      if (val === prev && prev !== '') matched = true;
      html.push(`<option value="${escapeHtml(val)}">${escapeHtml(opt.label)}</option>`);
    }
    if (preserveValue && prev && !matched) {
      html.push(`<option value="${escapeHtml(prev)}">${escapeHtml(prev)} (รายการเดิม)</option>`);
    }
    selectEl.innerHTML = html.join('');
    if (preserveValue && prev) selectEl.value = prev;
  }

  // ==========================================================================
  // 6B. Searchable Combobox (พิมพ์ค้นหา/พิมพ์เองได้ + ปุ่มดรอปดาวน์ ▾)
  // ==========================================================================
  const comboRegistry = {};

  function getFieldRecommendations() {
    const workerName = $('form-worker') ? $('form-worker').value.trim() : '';
    const plate = $('form-vehicle') ? $('form-vehicle').value.trim() : '';
    const siteName = $('form-site') ? $('form-site').value.trim() : '';
    const originVal = $('form-origin') ? $('form-origin').value.trim() : '';
    const destVal = $('form-destination') ? $('form-destination').value.trim() : '';

    const siteOriginsOf = (s) => {
      if (!s) return [];
      const arr = Array.isArray(s.default_origins) ? s.default_origins : [];
      const out = [];
      const seen = new Set();
      for (const x of [...arr, s.default_origin]) {
        const v = String(x || '').trim();
        if (!v || seen.has(v)) continue;
        seen.add(v);
        out.push(v);
      }
      return out;
    };

    const worker = workerName ? state.workers.find((w) => w.name === workerName) : null;
    const owner = plate ? state.workers.find((w) => w.default_vehicle === plate) : null;
    const siteObj = siteName ? state.sites.find((s) => s.site_name === siteName) : null;
    const matchedRoute =
      (originVal && state.routes.find((r) => r.active && r.origin === originVal)) ||
      (destVal && state.routes.find((r) => r.active && r.origin === destVal)) ||
      null;
    const routesBySite = siteName
      ? state.routes.filter((r) => r.active && (r.default_site === siteName || r.destination === siteName))
      : [];

    // 1. ทะเบียนรถแนะนำ (จากคนขับ)
    const recPlate = worker && worker.default_vehicle ? worker.default_vehicle : '';
    const recPlateReason = recPlate ? `รถประจำของ ${worker.name}` : '';

    // 2. โครงการแนะนำ (จากคนขับ -> จากเส้นทาง/จุดรับที่เลือก — รองรับหลายโครงการ)
    const recSites = [];
    const seenSites = new Set();
    const pushRecSite = (value, reason) => {
      const v = String(value || '').trim();
      if (!v || seenSites.has(v)) return;
      seenSites.add(v);
      recSites.push({ value: v, reason });
    };
    if (worker && worker.default_site) {
      pushRecSite(worker.default_site, `ผูกกับ ${worker.name}`);
    }
    if (matchedRoute && matchedRoute.default_site) {
      pushRecSite(matchedRoute.default_site, `ผูกกับเส้นทาง ${matchedRoute.origin}`);
    }
    if (matchedRoute && matchedRoute.destination && state.sites.some((s) => s.site_name === matchedRoute.destination)) {
      pushRecSite(matchedRoute.destination, `ผูกกับเส้นทาง ${matchedRoute.origin}`);
    }
    // หาโครงการทุกโครงการที่ผูกจุดรับตรงกับ originVal หรือ destVal
    const probeOrigins = [originVal, destVal].filter(Boolean);
    if (probeOrigins.length > 0) {
      for (const s of state.sites) {
        if (!s || !s.active) continue;
        const sOrigs = siteOriginsOf(s);
        const hit = probeOrigins.find((p) => sOrigs.includes(p));
        if (hit) pushRecSite(s.site_name, `ผูกกับเส้นทาง ${hit}`);
      }
      for (const r of state.routes) {
        if (!r || !r.active) continue;
        if (probeOrigins.includes(r.origin) && r.default_site) {
          pushRecSite(r.default_site, `ผูกกับเส้นทาง ${r.origin}`);
        }
      }
    }
    const recSite = recSites[0] ? recSites[0].value : '';
    const recSiteReason = recSites[0] ? recSites[0].reason : '';

    // 3. รวบรวมเส้นทางทั้งหมดที่ผูกไว้ (รองรับหลายเส้นทางต่อ 1 โครงการ!)
    const rawRoutes = [];
    const seenRouteKey = new Set();
    const pushRawRoute = (orig, dest, reason) => {
      const o = String(orig || '').trim();
      const d = String(dest || '').trim();
      if (!o && !d) return;
      const key = `${o}||${d}`;
      if (seenRouteKey.has(key)) return;
      seenRouteKey.add(key);
      rawRoutes.push({ orig: o, dest: d, reason });
    };

    const effectiveDestForSite = siteName || (worker && worker.default_site) || '';
    if (worker && (worker.default_origin || worker.default_destination)) {
      const wDest = worker.default_destination || siteName || worker.default_site || '';
      pushRawRoute(worker.default_origin, wDest, `ผูกกับ ${worker.name}`);
    }
    if (siteObj) {
      for (const orig of siteOriginsOf(siteObj)) {
        const mr = state.routes.find((r) => r.active && r.origin === orig && (r.default_site === siteObj.site_name || r.destination === siteObj.site_name));
        const dest = (mr && mr.destination) || siteObj.site_name;
        pushRawRoute(orig, dest, `ผูกกับโครงการ ${siteObj.site_name}`);
      }
    }
    for (const r of routesBySite) {
      pushRawRoute(r.origin, r.destination || siteName, `ผูกกับโครงการ ${siteName}`);
    }
    if (rawRoutes.length === 0) {
      let fallbackDest = '';
      let fallbackDestReason = '';
      if (siteName) {
        fallbackDest = siteName;
        fallbackDestReason = 'โครงการที่เลือก';
      } else if (matchedRoute && matchedRoute.destination) {
        fallbackDest = matchedRoute.destination;
        fallbackDestReason = `ผูกกับจุดรับ ${matchedRoute.origin}`;
      } else if (effectiveDestForSite) {
        fallbackDest = effectiveDestForSite;
        fallbackDestReason = worker ? `โครงการของ ${worker.name}` : 'โครงการที่เลือก';
      }
      if (fallbackDest) pushRawRoute('', fallbackDest, fallbackDestReason);
    }

    const recRoutes = rawRoutes.map((r) => ({
      orig: state.isSwapped ? r.dest : r.orig,
      dest: state.isSwapped ? r.orig : r.dest,
      reason: r.reason
    }));

    const primary = recRoutes[0] || { orig: '', dest: '', reason: '' };
    const recOrigin = primary.orig;
    const recOriginReason = primary.orig ? primary.reason : '';
    const recDest = primary.dest;
    const recDestReason = primary.dest ? primary.reason : '';

    return {
      worker,
      owner,
      workerName,
      plate,
      siteName,
      originVal,
      destVal,
      recPlate,
      recPlateReason,
      recSite,
      recSiteReason,
      recSites,
      recOrigin,
      recOriginReason,
      recDest,
      recDestReason,
      recRoutes
    };
  }

  function getComboOptions(id) {
    const activeWorkers = state.workers.filter((w) => w.active);
    const activeVehicles = state.vehicles.filter((v) => v.active);
    const activeSites = state.sites.filter((s) => s.active);
    const activeRoutes = state.routes.filter((r) => r.active);
    const rec = getFieldRecommendations();

    const dedup = (items) => {
      const seen = new Set();
      const out = [];
      for (const it of items) {
        const value = String(it.value || '').trim();
        if (!value || seen.has(value)) continue;
        seen.add(value);
        out.push({ value, sub: String(it.sub || '').trim(), recommended: Boolean(it.recommended) });
      }
      return out;
    };

    const siteOriginsOf = (s) => {
      if (!s) return [];
      const arr = Array.isArray(s.default_origins) ? s.default_origins : [];
      const seen = new Set();
      const out = [];
      for (const x of [...arr, s.default_origin]) {
        const v = String(x || '').trim();
        if (!v || seen.has(v)) continue;
        seen.add(v);
        out.push(v);
      }
      return out;
    };

    if (id === 'form-worker') {
      return dedup(
        activeWorkers.map((w) => {
          const parts = [];
          if (w.default_vehicle) parts.push(`รถ ${w.default_vehicle}`);
          if (w.default_site) parts.push(w.default_site);
          else if (w.default_origin) parts.push(w.default_origin);
          return {
            value: w.name,
            sub: parts.join(' • ')
          };
        })
      );
    }
    if (id === 'form-vehicle') {
      const list = [];
      if (rec.recPlate) {
        list.push({ value: rec.recPlate, sub: `★ แนะนำ (${rec.recPlateReason})`, recommended: true });
      }
      for (const v of activeVehicles) {
        const owner = activeWorkers.find((w) => w.default_vehicle === v.license_plate);
        list.push({
          value: v.license_plate,
          sub: owner ? `ประจำ: ${owner.name}` : ''
        });
      }
      return dedup(list);
    }
    if (id === 'form-site') {
      const list = [];
      for (const rs of rec.recSites || []) {
        list.push({ value: rs.value, sub: `★ แนะนำ (${rs.reason})`, recommended: true });
      }
      for (const s of activeSites) {
        const parts = [];
        if (s.zone_name) parts.push(`โซน: ${s.zone_name}`);
        const origs = siteOriginsOf(s);
        if (origs.length === 1) parts.push(`เส้นทาง: ${origs[0]}`);
        else if (origs.length > 1) parts.push(`เส้นทาง: ${origs.slice(0, 2).join(', ')}${origs.length > 2 ? ` +${origs.length - 2}` : ''}`);
        list.push({ value: s.site_name, sub: parts.join(' • ') });
      }
      return dedup(list);
    }
    if (id === 'form-origin') {
      const list = [];
      for (const rr of rec.recRoutes || []) {
        if (rr.orig) {
          list.push({ value: rr.orig, sub: `★ แนะนำ (${rr.reason})`, recommended: true });
        }
      }
      for (const r of activeRoutes) {
        list.push({
          value: r.origin,
          sub: r.destination ? `→ ${r.destination}` : r.default_site ? `โครงการ: ${r.default_site}` : 'จุดรับ'
        });
      }
      for (const s of activeSites) {
        for (const orig of siteOriginsOf(s)) {
          list.push({ value: orig, sub: `จุดรับโครงการ: ${s.site_name}` });
        }
      }
      for (const s of activeSites) {
        list.push({ value: s.site_name, sub: 'โครงการ' });
      }
      return dedup(list);
    }
    if (id === 'form-destination') {
      const list = [];
      for (const rr of rec.recRoutes || []) {
        if (rr.dest) {
          list.push({ value: rr.dest, sub: `★ แนะนำ (${rr.reason})`, recommended: true });
        }
      }
      for (const s of activeSites) {
        list.push({ value: s.site_name, sub: s.zone_name ? `โครงการ • โซน ${s.zone_name}` : 'โครงการ' });
      }
      for (const r of activeRoutes) {
        if (r.destination) list.push({ value: r.destination, sub: `จาก ${r.origin}` });
      }
      for (const r of activeRoutes) {
        list.push({ value: r.origin, sub: 'จุดรับ' });
      }
      return dedup(list);
    }
    return [];
  }

  function closeAllComboboxes(exceptId = '') {
    for (const [id, api] of Object.entries(comboRegistry)) {
      if (id !== exceptId) api.close();
    }
  }

  function setupCombobox(id) {
    const input = $(id);
    const wrap = document.querySelector(`[data-combo="${id}"]`);
    const toggleBtn = document.querySelector(`[data-combo-toggle="${id}"]`);
    const menu = document.querySelector(`[data-combo-menu="${id}"]`);
    if (!input || !wrap || !toggleBtn || !menu) return;

    let isOpen = false;
    let hiIdx = -1;
    let visibleItems = [];

    const close = () => {
      isOpen = false;
      hiIdx = -1;
      menu.classList.add('hidden');
      const icon = toggleBtn.querySelector('i');
      if (icon) icon.style.transform = '';
    };

    const renderMenu = (showAll = false) => {
      const all = getComboOptions(id);
      const q = showAll ? '' : input.value.trim().toLowerCase();
      const curVal = input.value.trim();
      // ถ้าค่าในช่องตรงกับรายการที่มีอยู่แล้ว ให้แสดงทั้งหมดเพื่อให้กดเปลี่ยนเป็นรายการอื่นได้ทันที
      const exactMatch = all.some((o) => o.value === curVal);
      const list = !q || exactMatch
        ? all
        : all.filter((o) => o.value.toLowerCase().includes(q) || o.sub.toLowerCase().includes(q));

      visibleItems = list;
      hiIdx = -1;

      if (list.length === 0) {
        menu.innerHTML = all.length === 0
          ? '<div class="px-3.5 py-2.5 text-xs text-slate-400 text-center">ยังไม่มีรายการตั้งค่า — พิมพ์ข้อความเองได้เลย</div>'
          : `<div class="px-3.5 py-2.5 text-xs text-slate-500 text-center">ไม่พบในรายการ — ใช้ข้อความที่พิมพ์เอง <strong>"${escapeHtml(curVal)}"</strong> ได้ทันที</div>`;
      } else {
        menu.innerHTML = list
          .map((o, idx) => {
            const selected = o.value === curVal;
            const rowCls = selected
              ? 'bg-indigo-50 text-indigo-900 font-semibold'
              : o.recommended
                ? 'bg-amber-50/90 hover:bg-amber-100/80 text-amber-950 font-semibold border-b border-amber-200/80'
                : 'text-slate-800 hover:bg-slate-100';
            const subCls = o.recommended ? 'text-[11px] text-amber-700 font-semibold' : 'text-[11px] text-slate-400 font-normal';
            return `
              <button type="button" data-combo-opt="${idx}"
                class="w-full text-left px-3.5 py-2 text-xs sm:text-sm flex items-center justify-between gap-2 transition cursor-pointer ${rowCls}">
                <span class="truncate">${escapeHtml(o.value)}</span>
                <span class="flex items-center gap-1.5 shrink-0">
                  ${o.sub ? `<span class="${subCls}">${escapeHtml(o.sub)}</span>` : ''}
                  ${selected ? '<i class="fa-solid fa-check text-indigo-600 text-[11px]"></i>' : ''}
                </span>
              </button>
            `;
          })
          .join('');
      }
    };

    const open = (showAll = false) => {
      closeAllComboboxes(id);
      renderMenu(showAll);
      isOpen = true;
      menu.classList.remove('hidden');
      const icon = toggleBtn.querySelector('i');
      if (icon) icon.style.transform = 'rotate(180deg)';
    };

    const selectItem = (val) => {
      input.value = val;
      close();
      input.dispatchEvent(new Event('change', { bubbles: true }));
    };

    const updateHighlight = () => {
      menu.querySelectorAll('[data-combo-opt]').forEach((btn, i) => {
        btn.classList.toggle('bg-indigo-100', i === hiIdx);
        if (i === hiIdx) btn.scrollIntoView({ block: 'nearest' });
      });
    };

    toggleBtn.addEventListener('mousedown', (e) => {
      e.preventDefault();
      if (isOpen) {
        close();
      } else {
        input.focus();
        open(true);
      }
    });

    input.addEventListener('focus', () => open(true));
    input.addEventListener('click', () => {
      if (!isOpen) open(true);
    });
    input.addEventListener('input', () => open(false));

    input.addEventListener('keydown', (e) => {
      if (!isOpen && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
        e.preventDefault();
        open(true);
        return;
      }
      if (!isOpen) return;
      if (e.key === 'Escape') {
        close();
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        if (visibleItems.length > 0) {
          hiIdx = (hiIdx + 1) % visibleItems.length;
          updateHighlight();
        }
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        if (visibleItems.length > 0) {
          hiIdx = (hiIdx - 1 + visibleItems.length) % visibleItems.length;
          updateHighlight();
        }
      } else if (e.key === 'Enter' && hiIdx >= 0 && visibleItems[hiIdx]) {
        e.preventDefault();
        selectItem(visibleItems[hiIdx].value);
      }
    });

    menu.addEventListener('mousedown', (e) => {
      const btn = e.target.closest('[data-combo-opt]');
      if (!btn) return;
      e.preventDefault();
      const idx = Number(btn.dataset.comboOpt);
      if (visibleItems[idx]) selectItem(visibleItems[idx].value);
    });

    input.addEventListener('blur', () => {
      setTimeout(close, 150);
    });

    comboRegistry[id] = { open, close, refresh: () => isOpen && renderMenu(false) };
  }

  function populateFrontDropdowns() {
    // Refresh any open searchable comboboxes
    for (const api of Object.values(comboRegistry)) api.refresh();

    // Filter site dropdown (include all sites + any historical site_name)
    populateSelect(
      $('filter-site'),
      state.sites.map((s) => ({
        value: s.site_name,
        label: s.zone_name ? `${s.site_name} (โซน ${s.zone_name})` : s.site_name
      })),
      '-- ทุกไซต์งาน / โครงการ --'
    );

    // Back-office worker form default_vehicle dropdown
    populateSelect(
      $('input-worker-default-vehicle'),
      state.vehicles.map((v) => ({
        value: v.license_plate,
        label: v.active ? v.license_plate : `${v.license_plate} (ปิดใช้งาน)`
      })),
      '-- ยังไม่ผูกทะเบียนรถ --'
    );

    // Back-office shared datalists (โครงการ, ต้นทาง, ปลายทาง)
    const fillDatalist = (dlId, values) => {
      const dl = $(dlId);
      if (!dl) return;
      const uniq = [...new Set(values.map((x) => String(x || '').trim()).filter(Boolean))];
      dl.innerHTML = uniq.map((v) => `<option value="${escapeHtml(v)}"></option>`).join('');
    };
    fillDatalist('dl-bo-sites', state.sites.map((s) => s.site_name));
    fillDatalist('dl-bo-origins', [
      ...state.routes.map((r) => r.origin),
      ...state.sites.flatMap((s) => (Array.isArray(s.default_origins) ? s.default_origins : [s.default_origin])),
      ...state.sites.map((s) => s.site_name)
    ]);
    fillDatalist('dl-bo-destinations', [
      ...state.sites.map((s) => s.site_name),
      ...state.routes.map((r) => r.destination),
      ...state.routes.map((r) => r.origin)
    ]);
  }

  // ==========================================================================
  // 7. Front-Office Form Logic
  // ==========================================================================
  function setTripType(type) {
    const isCharter = type === 'charter';
    $('form-trip-type').value = isCharter ? 'charter' : 'normal';

    const btnNormal = $('trip-type-normal');
    const btnCharter = $('trip-type-charter');
    if (isCharter) {
      btnNormal.className = 'min-h-[38px] px-2.5 py-1.5 rounded-lg text-xs font-semibold transition flex items-center justify-center gap-1 text-slate-600 hover:text-slate-900 cursor-pointer';
      btnCharter.className = 'min-h-[38px] px-2.5 py-1.5 rounded-lg text-xs font-semibold transition flex items-center justify-center gap-1 bg-amber-500 text-zinc-950 shadow-sm cursor-pointer';
      $('form-trip-type-hint').textContent = 'งานเหมา / ขนย้าย — แยกเข้าชีต "งานเหมา" และ "เหมา" อัตโนมัติ';
      $('form-passengers-hint').textContent = 'เว้นว่างหรือใส่ 0 ได้สำหรับงานเหมาขนย้าย';
      $('form-price-hint').textContent = 'ระบุราคาเหมาสำหรับคำนวณในใบเบิกและสำเนาโซน';
    } else {
      btnNormal.className = 'min-h-[38px] px-2.5 py-1.5 rounded-lg text-xs font-semibold transition flex items-center justify-center gap-1 bg-zinc-900 text-white shadow-sm cursor-pointer';
      btnCharter.className = 'min-h-[38px] px-2.5 py-1.5 rounded-lg text-xs font-semibold transition flex items-center justify-center gap-1 text-slate-600 hover:text-slate-900 cursor-pointer';
      $('form-trip-type-hint').textContent = 'รับ-ส่งพนักงานประจำสายวิ่ง';
      $('form-passengers-hint').textContent = 'ไม่บังคับสำหรับงานเหมา';
      $('form-price-hint').textContent = 'ดึงจากเส้นทาง หรือกรอกราคาเองได้';
    }
    const autoRet = $('form-auto-return');
    if (autoRet && !state.editingLogId) {
      autoRet.checked = !isCharter;
    }
    updateAutoReturnUI();
    updateRouteMetaPreview();
  }

  function updateAutoReturnUI() {
    const box = $('auto-return-box');
    const cb = $('form-auto-return');
    const preview = $('auto-return-preview');
    const submitText = $('btn-submit-text');
    if (!box || !cb) return;

    const isEditing = Boolean(state.editingLogId);
    box.classList.toggle('hidden', isEditing);
    const enabled = !isEditing && cb.checked;

    if (preview) {
      const orig = $('form-origin') ? $('form-origin').value.trim() : '';
      const dest = $('form-destination') ? $('form-destination').value.trim() : '';
      preview.textContent = enabled && orig && dest ? `ขากลับ: ${dest} → ${orig}` : '';
    }

    if (submitText && !isEditing) {
      submitText.textContent = enabled ? 'บันทึกขาไป + ขากลับ (2 เที่ยว)' : 'บันทึกเที่ยวนี้';
    }
  }

  function updateFormDatePreview() {
    const val = $('form-date').value;
    const be = toBE(val);
    $('form-date-thai-preview').textContent = be ? `วันที่ พ.ศ.: ${be}` : '';
  }

  function updateDurationPreview() {
    const dep = $('form-dep-time').value;
    const arr = $('form-arr-time').value;
    const dur = calcDurationText(dep, arr);
    const el = $('form-duration-preview');
    if (!el) return;
    if (dur) {
      el.innerHTML = `<i class="fa-solid fa-stopwatch text-emerald-600"></i><span class="text-emerald-700 font-semibold">ใช้เวลา ${escapeHtml(dur)}</span>`;
    } else {
      el.innerHTML = '<i class="fa-solid fa-stopwatch text-slate-400"></i><span>เว้นว่างได้ หรือระบุตามจริง</span>';
    }
  }

  function updateRouteMetaPreview() {
    const origin = $('form-origin').value.trim();
    const dest = $('form-destination').value.trim();
    const tripType = $('form-trip-type').value;
    const sheet = $('form-route-sheet-name').value.trim() || (origin && dest ? `${origin} - ${dest}` : '-');
    const billing = $('form-billing-title').value.trim() || sheet;
    const mode = $('form-display-mode-zone').value === 'license_plate' ? 'ทะเบียนรถ' : 'ชื่อคนขับ';
    const el = $('form-route-meta-preview');
    if (!el) return;

    if (tripType === 'charter') {
      el.innerHTML = `แผ่นงาน Excel: <strong class="text-amber-700">งานเหมา / เหมา</strong> • รายการในใบเบิก: <strong class="text-slate-700">${escapeHtml(billing)}</strong>`;
    } else {
      el.innerHTML = `แผ่นงาน Excel: <strong class="text-indigo-700">${escapeHtml(sheet)}</strong> • ชื่อในใบเบิก: <strong class="text-slate-700">${escapeHtml(billing)}</strong> • สำเนาโซนแสดง: <strong>${escapeHtml(mode)}</strong>`;
    }
    updateAutoReturnUI();
  }

  function toggleRouteMetaBox(forceOpen) {
    state.routeMetaExpanded = typeof forceOpen === 'boolean' ? forceOpen : !state.routeMetaExpanded;
    const box = $('route-meta-advanced-box');
    const btn = $('btn-toggle-route-meta');
    box.classList.toggle('hidden', !state.routeMetaExpanded);
    btn.textContent = state.routeMetaExpanded ? 'ซ่อนการตั้งค่าเพิ่มเติม' : 'ตั้งค่าเพิ่มเติม (ไม่บังคับ)';
  }

  function refreshAllComboMenus() {
    for (const api of Object.values(comboRegistry)) api.refresh();
  }

  function updateWorkerVehicleHints() {
    const rec = getFieldRecommendations();
    const {
      worker,
      owner,
      workerName,
      plate,
      siteName,
      originVal,
      destVal,
      recPlate,
      recSite,
      recOrigin,
      recDest
    } = rec;

    const wHint = $('form-worker-vehicle-hint');
    const vHint = $('form-vehicle-worker-hint');
    const sHint = $('form-site-rec-hint');
    const rHint = $('form-route-rec-hint');

    const vehEl = $('form-vehicle');
    const siteEl = $('form-site');
    const origEl = $('form-origin');
    const destEl = $('form-destination');

    if (vehEl) {
      vehEl.placeholder = recPlate
        ? `แนะนำ: ${recPlate} (เลือกหรือพิมพ์คันอื่นได้)`
        : 'เลือก หรือ พิมพ์ทะเบียนรถ...';
    }
    if (siteEl) {
      siteEl.placeholder = recSite
        ? `แนะนำ: ${recSite} (เลือกหรือพิมพ์โครงการอื่นได้)`
        : 'เลือก หรือ พิมพ์ชื่อโครงการ...';
    }
    if (origEl) {
      origEl.placeholder = recOrigin
        ? `แนะนำต้นทาง: ${recOrigin} (หรือพิมพ์เอง)...`
        : 'จาก (เลือกจุดรับ หรือ พิมพ์เอง)...';
    }
    if (destEl) {
      destEl.placeholder = recDest
        ? `แนะนำปลายทาง: ${recDest} (หรือพิมพ์เอง)...`
        : 'ถึง (เลือกปลายทาง หรือ พิมพ์เอง)...';
    }

    // 3. ชื่อคนขับ — สรุปค่าที่ผูกไว้ทั้งหมด + ปุ่มลัด "ใช้ที่ผูกทั้งหมด"
    if (wHint) {
      if (worker && (worker.default_vehicle || worker.default_site || worker.default_origin || worker.default_destination)) {
        const boundParts = [];
        if (worker.default_vehicle) boundParts.push(`รถ ${escapeHtml(worker.default_vehicle)}`);
        if (worker.default_site) boundParts.push(`โครงการ ${escapeHtml(worker.default_site)}`);
        if (worker.default_origin || worker.default_destination) {
          const rTxt =
            worker.default_origin && (worker.default_destination || worker.default_site)
              ? `${worker.default_origin} → ${worker.default_destination || worker.default_site}`
              : worker.default_origin || worker.default_destination;
          boundParts.push(`เส้นทาง ${escapeHtml(rTxt)}`);
        }
        const needsApplyAll =
          (recPlate && plate !== recPlate) ||
          (recSite && siteName !== recSite) ||
          (recOrigin && originVal !== recOrigin) ||
          (recDest && destVal !== recDest);
        wHint.innerHTML = `
          <span class="inline-flex flex-wrap items-center gap-1.5">
            <span class="text-slate-600">ผูกไว้: <strong class="text-amber-800">${boundParts.join(' • ')}</strong></span>
            ${
              needsApplyAll
                ? `<button type="button" data-apply-all-worker="1"
                    class="px-2 py-0.5 rounded-md bg-amber-100 hover:bg-amber-200 text-amber-950 border border-amber-300 font-bold text-[11px] cursor-pointer transition">
                    <i class="fa-solid fa-wand-magic-sparkles text-[10px] mr-0.5"></i>ใช้ที่ผูกทั้งหมด
                  </button>`
                : ''
            }
          </span>
        `;
      } else {
        wHint.textContent = 'เลือกจากดรอปดาวน์หรือพิมพ์เองได้ — ระบบแนะนำรถ โครงการ และเส้นทางที่ผูกไว้ให้';
      }
    }

    // 4. ทะเบียนรถ — ป้ายแนะนำคลิกได้แบบไม่บังคับ
    if (vHint) {
      if (recPlate && plate !== recPlate) {
        vHint.innerHTML = `
          <span class="inline-flex flex-wrap items-center gap-1.5">
            <span class="text-amber-700 font-semibold"><i class="fa-solid fa-star text-[10px] mr-0.5"></i>แนะนำ:</span>
            <button type="button" data-use-rec-plate="${escapeHtml(recPlate)}"
              class="px-2 py-0.5 rounded-md bg-amber-100 hover:bg-amber-200 text-amber-950 border border-amber-300 font-bold text-[11px] cursor-pointer transition">
              ${escapeHtml(recPlate)} (กดใช้)
            </button>
            <span class="text-slate-500">${plate ? 'หรือใช้คันที่เลือกนี้ต่อได้เลย' : 'หรือเลือก/พิมพ์คันอื่นได้'}</span>
          </span>
        `;
      } else if (recPlate && plate === recPlate) {
        vHint.innerHTML = `<span class="text-emerald-700 font-semibold"><i class="fa-solid fa-circle-check mr-1"></i>ตรงกับรถประจำ (${escapeHtml(recPlate)}) — แก้หรือเปลี่ยนคันได้อิสระ</span>`;
      } else if (owner && workerName && owner.name !== workerName) {
        vHint.innerHTML = `<span class="text-amber-700 font-semibold"><i class="fa-solid fa-user-pen mr-1"></i>รถคันนี้ผูกกับ ${escapeHtml(owner.name)} (ใช้บันทึกแทนได้ปกติ)</span>`;
      } else {
        vHint.textContent = 'เลือกจากดรอปดาวน์หรือพิมพ์เปลี่ยนคันได้อิสระ';
      }
    }

    // 5. โครงการ — ป้ายแนะนำคลิกได้แบบเดียวกับทะเบียนรถ (รองรับหลายโครงการที่ผูกกับเส้นทางเดียวกัน)
    const recSites = rec.recSites || [];
    if (sHint) {
      const matchedSiteRec = recSites.some((rs) => rs.value === siteName);
      if (recSites.length > 1) {
        const siteBtns = recSites
          .map((rs) => {
            const active = rs.value === siteName;
            return `<button type="button" data-use-rec-site="${escapeHtml(rs.value)}"
              class="px-2 py-0.5 rounded-md ${active ? 'bg-emerald-100 text-emerald-900 border-emerald-300' : 'bg-amber-100 hover:bg-amber-200 text-amber-950 border-amber-300'} border font-bold text-[11px] cursor-pointer transition">
              ${active ? '<i class="fa-solid fa-check text-[10px] mr-0.5"></i>' : ''}${escapeHtml(rs.value)}${active ? '' : ' (กดใช้)'}
            </button>`;
          })
          .join('');
        sHint.innerHTML = `
          <span class="inline-flex flex-wrap items-center gap-1.5">
            <span class="text-amber-700 font-semibold"><i class="fa-solid fa-star text-[10px] mr-0.5"></i>แนะนำ (${recSites.length} โครงการ):</span>
            ${siteBtns}
          </span>
        `;
      } else if (recSite && !matchedSiteRec) {
        sHint.innerHTML = `
          <span class="inline-flex flex-wrap items-center gap-1.5">
            <span class="text-amber-700 font-semibold"><i class="fa-solid fa-star text-[10px] mr-0.5"></i>แนะนำ:</span>
            <button type="button" data-use-rec-site="${escapeHtml(recSite)}"
              class="px-2 py-0.5 rounded-md bg-amber-100 hover:bg-amber-200 text-amber-950 border border-amber-300 font-bold text-[11px] cursor-pointer transition">
              ${escapeHtml(recSite)} (กดใช้)
            </button>
            <span class="text-slate-500">${siteName ? 'หรือใช้โครงการที่พิมพ์นี้ต่อได้เลย' : 'หรือเลือก/พิมพ์โครงการอื่นได้'}</span>
          </span>
        `;
      } else if (recSite && matchedSiteRec) {
        sHint.innerHTML = `<span class="text-emerald-700 font-semibold"><i class="fa-solid fa-circle-check mr-1"></i>ตรงกับโครงการที่ผูกไว้ (${escapeHtml(recSite)}) — พิมพ์หรือเปลี่ยนได้อิสระ</span>`;
      } else {
        sHint.textContent = 'เลือกจากดรอปดาวน์หรือพิมพ์เองได้ — ระบบขึ้นแนะนำโครงการที่ผูกไว้ให้';
      }
    }

    // 6. เส้นทาง (จาก ↔ ถึง) — ป้ายแนะนำคลิกได้ครบทุกเส้นทางที่ผูกกับโครงการ/คนขับ
    const recRoutes = (rec.recRoutes || []).filter((r) => r.orig || r.dest);
    if (rHint) {
      if (recRoutes.length > 1) {
        const routeBtns = recRoutes
          .map((rr) => {
            const isActive =
              (!rr.orig || originVal === rr.orig) && (!rr.dest || destVal === rr.dest);
            const label = rr.orig && rr.dest ? `${rr.orig} → ${rr.dest}` : rr.orig || rr.dest;
            return `
              <button type="button" data-use-rec-route="1"
                data-rec-orig="${escapeHtml(rr.orig)}" data-rec-dest="${escapeHtml(rr.dest)}"
                class="px-2 py-0.5 rounded-md ${isActive ? 'bg-emerald-100 text-emerald-900 border-emerald-300' : 'bg-amber-100 hover:bg-amber-200 text-amber-950 border-amber-300'} border font-bold text-[11px] cursor-pointer transition">
                ${isActive ? '<i class="fa-solid fa-check text-[10px] mr-0.5"></i>' : ''}${escapeHtml(label)}
              </button>
            `;
          })
          .join('');
        rHint.innerHTML = `
          <span class="inline-flex flex-wrap items-center gap-1.5">
            <span class="text-amber-700 font-semibold"><i class="fa-solid fa-star text-[10px] mr-0.5"></i>เส้นทางที่ผูก (${recRoutes.length}):</span>
            ${routeBtns}
            <span class="text-slate-500">คลิกเลือก หรือพิมพ์เองได้อิสระ</span>
          </span>
        `;
      } else {
        const needOrig = Boolean(recOrigin && originVal !== recOrigin);
        const needDest = Boolean(recDest && destVal !== recDest);
        const hasAnyRec = Boolean(recOrigin || recDest);

        if (needOrig || needDest) {
          const btns = [];
          if (recOrigin && recDest && (needOrig || needDest)) {
            btns.push(`
              <button type="button" data-use-rec-route="1"
                data-rec-orig="${escapeHtml(recOrigin)}" data-rec-dest="${escapeHtml(recDest)}"
                class="px-2 py-0.5 rounded-md bg-amber-100 hover:bg-amber-200 text-amber-950 border border-amber-300 font-bold text-[11px] cursor-pointer transition">
                ${escapeHtml(recOrigin)} → ${escapeHtml(recDest)} (กดใช้ทั้งคู่)
              </button>
            `);
          }
          if (needOrig) {
            btns.push(`
              <button type="button" data-use-rec-origin="${escapeHtml(recOrigin)}"
                class="px-2 py-0.5 rounded-md bg-white hover:bg-amber-50 text-amber-900 border border-amber-300 font-semibold text-[11px] cursor-pointer transition">
                จาก: ${escapeHtml(recOrigin)}
              </button>
            `);
          }
          if (needDest) {
            btns.push(`
              <button type="button" data-use-rec-dest="${escapeHtml(recDest)}"
                class="px-2 py-0.5 rounded-md bg-white hover:bg-amber-50 text-amber-900 border border-amber-300 font-semibold text-[11px] cursor-pointer transition">
                ถึง: ${escapeHtml(recDest)}
              </button>
            `);
          }
          rHint.innerHTML = `
            <span class="inline-flex flex-wrap items-center gap-1.5">
              <span class="text-amber-700 font-semibold"><i class="fa-solid fa-star text-[10px] mr-0.5"></i>แนะนำเส้นทาง:</span>
              ${btns.join('')}
              <span class="text-slate-500">หรือพิมพ์/เลือกเองได้อิสระ</span>
            </span>
          `;
        } else if (hasAnyRec) {
          rHint.innerHTML = `<span class="text-emerald-700 font-semibold"><i class="fa-solid fa-circle-check mr-1"></i>ตรงกับเส้นทางที่แนะนำ — พิมพ์หรือเปลี่ยนต้นทาง-ปลายทางได้อิสระ</span>`;
        } else {
          rHint.textContent = 'เลือกจากดรอปดาวน์หรือพิมพ์เองได้อิสระ — ระบบขึ้นแนะนำเส้นทางที่ผูกไว้ให้';
        }
      }
    }
  }

  function onWorkerSelectChange() {
    state.lastAutoWorker = '';
    updateWorkerVehicleHints();
    refreshAllComboMenus();
  }

  function onVehicleSelectChange() {
    state.lastAutoVehicle = '';
    updateWorkerVehicleHints();
  }

  function onSiteInputChange() {
    updateWorkerVehicleHints();
    refreshAllComboMenus();
    updateRouteMetaPreview();
  }

  function applyRoutePresetFromInput() {
    const origVal = $('form-origin').value.trim();
    const destVal = $('form-destination').value.trim();
    const route =
      state.routes.find((r) => r.active && r.origin === origVal && (!r.destination || r.destination === destVal)) ||
      state.routes.find((r) => r.active && r.origin === origVal) ||
      state.routes.find((r) => r.active && r.origin === destVal);

    if (route) {
      // เติมเฉพาะเวลา/ราคา/ชื่อแผ่นงานที่ยังว่างอยู่ (ไม่บังคับเขียนทับต้นทาง-ปลายทางหรือโครงการ — ให้ขึ้นแนะนำแทน)
      if (route.default_dep_time && !$('form-dep-time').value.trim()) {
        $('form-dep-time').value = route.default_dep_time;
      }
      if (route.default_arr_time && !$('form-arr-time').value.trim()) {
        $('form-arr-time').value = route.default_arr_time;
      }
      if (route.unit_price && !$('form-unit-price').value.trim()) {
        $('form-unit-price').value = String(route.unit_price);
        pulseElement($('form-unit-price'));
      }
      if (route.display_mode_zone) {
        $('form-display-mode-zone').value = route.display_mode_zone;
      }
      if (route.route_sheet_name && route.route_sheet_name !== route.origin) {
        $('form-route-sheet-name').value = route.route_sheet_name;
      }
      if (route.billing_title && route.billing_title !== route.origin) {
        $('form-billing-title').value = route.billing_title;
      }
      updateDurationPreview();
    }

    updateWorkerVehicleHints();
    refreshAllComboMenus();
    updateRouteMetaPreview();
  }

  function onSwapRouteClick() {
    const origEl = $('form-origin');
    const destEl = $('form-destination');
    const orig = origEl.value.trim();
    const dest = destEl.value.trim();

    if (!orig && !dest) {
      showToast('กรุณาเลือกเส้นทาง หรือกรอกต้นทางและปลายทางก่อนกดสลับ', 'warning');
      return;
    }

    // คง route_sheet_name และ billing_title ไว้ตามขาไปเดิมเสมอ (แม้ไม่ได้เลือกจาก dropdown)
    const sheetEl = $('form-route-sheet-name');
    const billingEl = $('form-billing-title');
    if (!sheetEl.value.trim() && orig && dest) {
      sheetEl.value = `${orig} - ${dest}`;
    }
    if (!billingEl.value.trim() && sheetEl.value.trim()) {
      billingEl.value = sheetEl.value.trim();
    }

    origEl.value = dest;
    destEl.value = orig;
    state.isSwapped = !state.isSwapped;

    const badge = $('swap-status-badge');
    if (state.isSwapped) {
      badge.classList.remove('hidden');
      badge.classList.add('inline-flex');
    } else {
      badge.classList.add('hidden');
      badge.classList.remove('inline-flex');
    }

    if (state.isSwapped && $('form-auto-return')) {
      $('form-auto-return').checked = false;
    }

    pulseElement(origEl);
    pulseElement(destEl);
    updateWorkerVehicleHints();
    refreshAllComboMenus();
    updateAutoReturnUI();
    updateRouteMetaPreview();
  }

  function resetForm({ keepDateAndSite = false } = {}) {
    const prevDate = $('form-date').value || todayIso();
    const prevSite = $('form-site').value;

    state.editingLogId = null;
    state.isSwapped = false;
    state.lastAutoWorker = '';
    state.lastAutoVehicle = '';

    $('travel-log-form').reset();
    setTripType('normal');
    $('form-date').value = keepDateAndSite ? prevDate : todayIso();
    if (keepDateAndSite && prevSite) {
      $('form-site').value = prevSite;
      state.lastLinkedSite = prevSite;
    } else {
      state.lastLinkedSite = '';
    }
    $('form-route-sheet-name').value = '';
    $('form-billing-title').value = '';
    $('form-display-mode-zone').value = 'worker_name';
    if ($('form-auto-return')) $('form-auto-return').checked = true;
    updateWorkerVehicleHints();
    refreshAllComboMenus();

    $('swap-status-badge').classList.add('hidden');
    $('swap-status-badge').classList.remove('inline-flex');

    // Exit edit visuals
    $('edit-mode-badge-group').classList.add('hidden');
    $('edit-mode-badge-group').classList.remove('flex');
    $('btn-cancel-edit').classList.add('hidden');
    $('btn-cancel-edit').classList.remove('flex');
    $('form-card-container').classList.remove('ring-2', 'ring-amber-400');
    $('form-mode-title').textContent = 'ขั้นที่ 2: บันทึกเที่ยววิ่ง';
    $('form-mode-icon').innerHTML = '<i class="fa-solid fa-plus"></i>';
    $('btn-submit-text').textContent = 'บันทึกข้อมูลการเดินทาง';
    $('btn-submit-icon').className = 'fa-solid fa-floppy-disk text-amber-400';

    updateFormDatePreview();
    updateDurationPreview();
    updateAutoReturnUI();
    updateRouteMetaPreview();
  }

  function startEditLog(log) {
    state.editingLogId = log.id;
    state.isSwapped = false;

    setTripType(log.trip_type || 'normal');
    $('form-date').value = log.trip_date || todayIso();

    state.lastAutoWorker = '';
    state.lastAutoVehicle = '';
    $('form-worker').value = log.worker_name || '';
    $('form-vehicle').value = log.license_plate || '';
    $('form-site').value = log.site_name || '';
    state.lastLinkedSite = log.site_name || '';
    updateWorkerVehicleHints();
    $('form-origin').value = log.origin || '';
    $('form-destination').value = log.destination || '';
    $('form-route-sheet-name').value = log.route_sheet_name || '';
    $('form-billing-title').value = log.billing_title || '';
    $('form-display-mode-zone').value = log.display_mode_zone || 'worker_name';
    $('form-dep-time').value = log.dep_time || '';
    $('form-arr-time').value = log.arr_time || '';
    $('form-passengers').value = log.passengers === null || log.passengers === undefined ? '' : String(log.passengers);
    $('form-unit-price').value = log.unit_price ? String(log.unit_price) : '';
    $('form-remark').value = log.remark || '';

    $('edit-mode-label').textContent = `${log.worker_name} (${toBE(log.trip_date)})`;
    $('edit-mode-badge-group').classList.remove('hidden');
    $('edit-mode-badge-group').classList.add('flex');
    $('btn-cancel-edit').classList.remove('hidden');
    $('btn-cancel-edit').classList.add('flex');
    $('form-card-container').classList.add('ring-2', 'ring-amber-400');
    $('form-mode-title').textContent = 'แก้ไขรายการบันทึกการเดินทาง';
    $('form-mode-icon').innerHTML = '<i class="fa-solid fa-pen-to-square"></i>';
    $('btn-submit-text').textContent = 'อัปเดตรายการเดินทาง';
    $('btn-submit-icon').className = 'fa-solid fa-check text-amber-400';

    updateAutoReturnUI();
    updateFormDatePreview();
    updateDurationPreview();
    updateAutoReturnUI();
    updateRouteMetaPreview();

    $('form-card-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function handleLogFormSubmit(e) {
    e.preventDefault();

    const trip_date = $('form-date').value.trim();
    const trip_type = $('form-trip-type').value;
    const worker_name = $('form-worker').value.trim();
    const license_plate = $('form-vehicle').value.trim();
    const site_name = $('form-site').value.trim();
    const origin = $('form-origin').value.trim();
    const destination = $('form-destination').value.trim();
    const route_sheet_name = $('form-route-sheet-name').value.trim() || (origin && destination ? `${origin} - ${destination}` : '');
    const billing_title = $('form-billing-title').value.trim() || route_sheet_name;
    const dep_time = $('form-dep-time').value.trim();
    const arr_time = $('form-arr-time').value.trim();
    const passRaw = $('form-passengers').value.trim();
    const passengers = passRaw === '' ? null : Number(passRaw);
    const priceRaw = $('form-unit-price').value.trim();
    const unit_price = priceRaw === '' ? 0 : Number(priceRaw);
    const display_mode_zone = $('form-display-mode-zone').value || 'worker_name';
    const remark = $('form-remark').value.trim();

    if (!trip_date) {
      showToast('กรุณาระบุวันที่เดินทาง', 'error');
      $('form-date').focus();
      return;
    }
    if (!worker_name) {
      showToast('กรุณาเลือกชื่อคนขับ', 'error');
      $('form-worker').focus();
      return;
    }
    if (!origin || !destination) {
      showToast('กรุณาระบุต้นทาง (จาก) และปลายทาง (ถึง)', 'error');
      (!origin ? $('form-origin') : $('form-destination')).focus();
      return;
    }

    const payload = {
      trip_date,
      trip_type,
      worker_name,
      license_plate,
      site_name,
      route_sheet_name,
      billing_title,
      origin,
      dep_time,
      destination,
      arr_time,
      passengers,
      unit_price,
      display_mode_zone,
      remark
    };

    const submitBtn = $('btn-submit-log');
    submitBtn.disabled = true;
    try {
      if (state.editingLogId) {
        await apiFetch(`/api/logs/${encodeURIComponent(state.editingLogId)}`, {
          method: 'PUT',
          body: JSON.stringify(payload)
        });
        showToast('อัปเดตรายการเดินทางเรียบร้อยแล้ว', 'success');
        resetForm({ keepDateAndSite: true });
      } else {
        const autoReturn = Boolean($('form-auto-return') && $('form-auto-return').checked);

        const res1 = await apiFetch('/api/logs', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        let latestQuota = res1 && res1.storageQuota;

        if (autoReturn) {
          // ขากลับอัตโนมัติ: สลับต้นทาง–ปลายทาง คงกลุ่มแผ่นงานเดิม และปล่อยเวลา/จำนวนคนเว้นว่างไว้ให้แก้เอง
          const returnPayload = {
            ...payload,
            origin: destination,
            destination: origin,
            dep_time: '',
            arr_time: '',
            passengers: null
          };
          const res2 = await apiFetch('/api/logs', {
            method: 'POST',
            body: JSON.stringify(returnPayload)
          });
          if (res2 && res2.storageQuota) latestQuota = res2.storageQuota;
          if (latestQuota) renderQuotaUI(latestQuota);
          showToast(`บันทึกขาไป + ขากลับสำเร็จ (2 เที่ยว: ${worker_name} • ${origin} ↔ ${destination})`, 'success', 4000);
        } else {
          if (latestQuota) renderQuotaUI(latestQuota);
          showToast(`บันทึกเที่ยววิ่ง ${worker_name} (${origin} → ${destination}) เรียบร้อยแล้ว`, 'success');
        }
        resetForm({ keepDateAndSite: true });
      }
      await loadLogs();
      await autoSyncBoundFiles();
    } catch (err) {
      if (err.data && err.data.storageQuota) renderQuotaUI(err.data.storageQuota);
      showToast(err.message || 'ไม่สามารถบันทึกข้อมูลได้', 'error', 5000);
    } finally {
      submitBtn.disabled = false;
    }
  }

  // คัดลอกคิววิ่งจากวันก่อนหน้า
  function openCopyPreviousDayModal() {
    const defaultTo = $('form-date').value || todayIso();
    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <div class="flex items-start gap-3">
          <div class="w-10 h-10 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center shrink-0">
            <i class="fa-solid fa-copy"></i>
          </div>
          <div>
            <h3 class="text-base font-bold text-slate-900">คัดลอกคิววิ่งจากวันก่อนหน้า</h3>
            <p class="text-xs text-slate-500 mt-0.5">
              ดึงรายการเที่ยววิ่งที่เคยบันทึกไว้มาสร้างเป็นรายการของวันที่เลือกโดยอัตโนมัติ (สามารถแก้ไขจำนวนคนหรือเวลาภายหลังได้)
            </p>
          </div>
        </div>

        <form id="modal-copy-form" class="space-y-3.5">
          <div>
            <label for="modal-copy-to-date" class="block text-xs font-semibold text-slate-700 mb-1">
              สร้างรายการสำหรับวันที่ (วันที่ปลายทาง) *
            </label>
            <input type="date" id="modal-copy-to-date" value="${escapeHtml(defaultTo)}" required
              class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
          </div>
          <div>
            <label for="modal-copy-from-date" class="block text-xs font-semibold text-slate-700 mb-1">
              คัดลอกจากวันที่ (ไม่บังคับ — เว้นว่างเพื่อใช้ข้อมูลวันล่าสุดก่อนหน้าอัตโนมัติ)
            </label>
            <input type="date" id="modal-copy-from-date"
              class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
          </div>
          <div class="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="modal-copy-cancel" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs sm:text-sm font-medium cursor-pointer">ยกเลิก</button>
            <button type="submit" id="modal-copy-submit" class="min-h-[40px] px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold cursor-pointer">
              <i class="fa-solid fa-copy mr-1"></i>คัดลอกคิววิ่ง
            </button>
          </div>
        </form>
      </div>
      `,
      (box) => {
        box.querySelector('#modal-copy-cancel').addEventListener('click', closeModal);
        box.querySelector('#modal-copy-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          const to_date = box.querySelector('#modal-copy-to-date').value.trim();
          const from_date = box.querySelector('#modal-copy-from-date').value.trim();
          const btn = box.querySelector('#modal-copy-submit');
          btn.disabled = true;
          try {
            const res = await apiFetch('/api/logs/copy-previous', {
              method: 'POST',
              body: JSON.stringify({ to_date, from_date })
            });
            closeModal();
            if (res.storageQuota) renderQuotaUI(res.storageQuota);
            showToast(
              `คัดลอกคิววิ่งจากวันที่ ${toBE(res.from_date)} ไปยัง ${toBE(res.to_date)} สำเร็จ (${res.created} รายการ)`,
              'success',
              4500
            );
            await loadLogs();
            await autoSyncBoundFiles();
          } catch (err) {
            if (err.data && err.data.storageQuota) renderQuotaUI(err.data.storageQuota);
            showToast(err.message || 'ไม่สามารถคัดลอกคิววิ่งได้', 'error', 4500);
            btn.disabled = false;
          }
        });
      }
    );
  }

  // ==========================================================================
  // 8. Filters, Quick Periods & Travel Logs Table Rendering
  // ==========================================================================
  function applyQuickPeriod(mode) {
    const monthVal = $('filter-month').value || currentMonthIso();
    const [yStr, mStr] = monthVal.split('-');
    const year = Number(yStr);
    const month = Number(mStr);

    if (mode === '1-15') {
      $('filter-date-from').value = `${year}-${String(month).padStart(2, '0')}-01`;
      $('filter-date-to').value = `${year}-${String(month).padStart(2, '0')}-15`;
    } else if (mode === '16-end') {
      const lastDay = new Date(year, month, 0).getDate();
      $('filter-date-from').value = `${year}-${String(month).padStart(2, '0')}-16`;
      $('filter-date-to').value = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
    } else if (mode === 'today') {
      const t = todayIso();
      $('filter-date-from').value = t;
      $('filter-date-to').value = t;
    } else if (mode === 'all') {
      $('filter-site').value = '';
      $('filter-trip-type').value = '';
      $('filter-date-from').value = '';
      $('filter-date-to').value = '';
      $('filter-search').value = '';
    }
    loadLogs();
  }

  function renderFilterSummaryLabel() {
    const f = getFilterParams();
    const parts = [];
    if (f.site_name) parts.push(`โครงการ: ${f.site_name}`);
    if (f.date_from || f.date_to) {
      parts.push(`ช่วงวันที่: ${f.date_from ? toBE(f.date_from) : '...'} - ${f.date_to ? toBE(f.date_to) : '...'}`);
    }
    if (f.trip_type === 'normal') parts.push('รับ-ส่งปกติ');
    if (f.trip_type === 'charter') parts.push('งานเหมา/ขนย้าย');
    $('filter-period-label').textContent = parts.length ? parts.join(' • ') : 'แสดงทุกงวดวันที่';
  }

  function renderLogs() {
    const logs = state.logs || [];
    const totalAmount = logs.reduce((acc, l) => acc + (Number(l.unit_price) || 0), 0);

    $('stat-filtered-logs').textContent = formatInt(logs.length);
    $('stat-filtered-amount').textContent = formatMoney(totalAmount);
    $('stat-today-logs').textContent = formatInt(state.todayLogsCount);
    $('logs-count-badge').textContent = `${formatInt(logs.length)} รายการ`;

    const tbody = $('logs-table-body');
    const mobile = $('logs-mobile-cards');
    const empty = $('logs-empty-state');

    if (logs.length === 0) {
      tbody.innerHTML = '';
      mobile.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }
    empty.classList.add('hidden');

    tbody.innerHTML = logs
      .map((l, idx) => {
        const isCharter = l.trip_type === 'charter';
        const typeBadge = isCharter
          ? '<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-900 border border-amber-300">งานเหมา</span>'
          : '<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200">รับ-ส่งปกติ</span>';
        const dur = calcDurationText(l.dep_time, l.arr_time);
        const passText = l.passengers === null || l.passengers === undefined ? '-' : `${l.passengers} คน`;

        return `
          <tr class="hover:bg-slate-50/90 transition align-top">
            <td class="py-3 px-3 text-center text-xs font-semibold text-slate-500">${idx + 1}</td>
            <td class="py-3 px-3 whitespace-nowrap">
              <div class="font-semibold text-slate-900 text-xs">${escapeHtml(toBE(l.trip_date))}</div>
              <div class="mt-1">${typeBadge}</div>
            </td>
            <td class="py-3 px-3">
              <div class="font-semibold text-slate-900 text-xs sm:text-sm">${escapeHtml(l.worker_name)}</div>
              ${
                l.license_plate
                  ? `<div class="mt-1"><span class="license-plate-badge !text-[11px] !py-0.5 !px-2">${escapeHtml(l.license_plate)}</span></div>`
                  : '<div class="text-[11px] text-slate-400 mt-0.5">ไม่ระบุทะเบียน</div>'
              }
            </td>
            <td class="py-3 px-3">
              <div class="font-semibold text-slate-800 text-xs">${escapeHtml(l.site_name || '-')}</div>
              <div class="text-[11px] text-slate-500 mt-0.5">ชีต: ${escapeHtml(isCharter ? 'งานเหมา / เหมา' : l.route_sheet_name)}</div>
            </td>
            <td class="py-3 px-3">
              <div class="text-xs font-medium text-slate-900">
                <span>${escapeHtml(l.origin)}</span>
                <i class="fa-solid fa-arrow-right text-[10px] text-amber-500 mx-1"></i>
                <span>${escapeHtml(l.destination)}</span>
              </div>
              <div class="text-[11px] text-slate-500 mt-0.5">ใบเบิก: ${escapeHtml(l.billing_title || l.route_sheet_name)}</div>
            </td>
            <td class="py-3 px-3 text-center whitespace-nowrap">
              <div class="text-xs font-semibold text-slate-800">${escapeHtml(l.dep_time || '-')} – ${escapeHtml(l.arr_time || '-')}</div>
              ${dur ? `<div class="text-[11px] text-emerald-700 font-medium mt-0.5"><i class="fa-solid fa-stopwatch mr-0.5"></i>${escapeHtml(dur)}</div>` : ''}
            </td>
            <td class="py-3 px-3 text-center text-xs font-semibold text-slate-700 whitespace-nowrap">${escapeHtml(passText)}</td>
            <td class="py-3 px-3 text-right text-xs font-bold text-emerald-700 whitespace-nowrap">${formatMoney(l.unit_price)}</td>
            <td class="py-3 px-3 text-xs text-slate-600 max-w-[160px] break-words">${escapeHtml(l.remark || '-')}</td>
            <td class="py-3 px-3 text-right whitespace-nowrap">
              <div class="inline-flex items-center gap-1">
                <button type="button" data-action="edit-log" data-id="${escapeHtml(l.id)}" title="แก้ไข"
                  class="w-8 h-8 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 inline-flex items-center justify-center cursor-pointer transition">
                  <i class="fa-solid fa-pen text-xs"></i>
                </button>
                <button type="button" data-action="delete-log" data-id="${escapeHtml(l.id)}" title="ลบ"
                  class="w-8 h-8 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 inline-flex items-center justify-center cursor-pointer transition">
                  <i class="fa-solid fa-trash text-xs"></i>
                </button>
              </div>
            </td>
          </tr>
        `;
      })
      .join('');

    mobile.innerHTML = logs
      .map((l, idx) => {
        const isCharter = l.trip_type === 'charter';
        const dur = calcDurationText(l.dep_time, l.arr_time);
        const passText = l.passengers === null || l.passengers === undefined ? '-' : `${l.passengers} คน`;
        return `
          <div class="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
            <div class="flex items-start justify-between gap-2">
              <div>
                <div class="flex items-center gap-1.5">
                  <span class="text-xs font-bold text-slate-400">#${idx + 1}</span>
                  <span class="text-xs font-bold text-slate-900">${escapeHtml(toBE(l.trip_date))}</span>
                  <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${isCharter ? 'bg-amber-100 text-amber-900' : 'bg-indigo-100 text-indigo-800'}">
                    ${isCharter ? 'งานเหมา' : 'รับ-ส่งปกติ'}
                  </span>
                </div>
                <div class="font-bold text-sm text-slate-900 mt-1">${escapeHtml(l.worker_name)}</div>
              </div>
              ${l.license_plate ? `<span class="license-plate-badge !text-[11px] !py-0.5 !px-2">${escapeHtml(l.license_plate)}</span>` : ''}
            </div>

            <div class="text-xs text-slate-700 bg-white rounded-lg p-2.5 border border-slate-200/80 space-y-1">
              <div class="font-semibold text-slate-900">
                ${escapeHtml(l.origin)} <i class="fa-solid fa-arrow-right text-amber-500 text-[10px] mx-1"></i> ${escapeHtml(l.destination)}
              </div>
              <div class="flex flex-wrap gap-x-3 gap-y-1 text-[11px] text-slate-600">
                <span><i class="fa-solid fa-building text-amber-500 mr-1"></i>${escapeHtml(l.site_name || '-')}</span>
                <span><i class="fa-regular fa-clock text-indigo-500 mr-1"></i>${escapeHtml(l.dep_time || '-')}–${escapeHtml(l.arr_time || '-')}${dur ? ` (${escapeHtml(dur)})` : ''}</span>
                <span><i class="fa-solid fa-users text-slate-400 mr-1"></i>${escapeHtml(passText)}</span>
                <span class="font-bold text-emerald-700">${formatMoney(l.unit_price)} บ.</span>
              </div>
              ${l.remark ? `<div class="text-[11px] text-slate-500"><i class="fa-regular fa-note-sticky mr-1"></i>${escapeHtml(l.remark)}</div>` : ''}
            </div>

            <div class="flex items-center justify-end gap-2">
              <button type="button" data-action="edit-log" data-id="${escapeHtml(l.id)}"
                class="min-h-[36px] px-3 py-1 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-200 text-xs font-semibold cursor-pointer">
                <i class="fa-solid fa-pen mr-1"></i>แก้ไข
              </button>
              <button type="button" data-action="delete-log" data-id="${escapeHtml(l.id)}"
                class="min-h-[36px] px-3 py-1 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-semibold cursor-pointer">
                <i class="fa-solid fa-trash mr-1"></i>ลบ
              </button>
            </div>
          </div>
        `;
      })
      .join('');
  }

  async function handleLogsActionClick(e) {
    const btn = e.target.closest('button[data-action]');
    if (!btn) return;
    const action = btn.dataset.action;
    const id = btn.dataset.id;
    const log = state.logs.find((l) => l.id === id);
    if (!log) return;

    if (action === 'edit-log') {
      startEditLog(log);
    } else if (action === 'delete-log') {
      const ok = await confirmModal({
        title: 'ยืนยันการลบรายการเดินทาง',
        message: `ต้องการลบเที่ยววิ่งของ "${log.worker_name}" วันที่ ${toBE(log.trip_date)} (${log.origin} → ${log.destination}) ใช่หรือไม่?`,
        confirmText: 'ลบรายการ',
        danger: true
      });
      if (!ok) return;
      try {
        const res = await apiFetch(`/api/logs/${encodeURIComponent(id)}`, { method: 'DELETE' });
        if (res && res.storageQuota) renderQuotaUI(res.storageQuota);
        if (state.editingLogId === id) resetForm({ keepDateAndSite: true });
        showToast('ลบรายการเดินทางเรียบร้อยแล้ว', 'success');
        await loadLogs();
        await autoSyncBoundFiles();
      } catch (err) {
        showToast(err.message || 'ไม่สามารถลบรายการได้', 'error');
      }
    }
  }

  async function handleClearAllLogs() {
    const ok = await confirmModal({
      title: 'ยืนยันการล้างประวัติการเดินทางทั้งหมด',
      message: 'การดำเนินการนี้จะลบรายการบันทึกการเดินทาง (travel_logs) ทั้งหมดออกจากฐานข้อมูลอย่างถาวร (ข้อมูลหลักคนขับ รถ โครงการ เส้นทาง จะยังคงอยู่) แนะนำให้ส่งออกไฟล์ Excel สำรองก่อนดำเนินการ',
      confirmText: 'ล้างรายการทั้งหมด',
      danger: true
    });
    if (!ok) return;
    try {
      const res = await apiFetch('/api/logs/clear', { method: 'POST' });
      if (res && res.storageQuota) renderQuotaUI(res.storageQuota);
      showToast(res.message || 'ล้างรายการเดินทางทั้งหมดแล้ว', 'success');
      await loadLogs();
      await autoSyncBoundFiles();
    } catch (err) {
      showToast(err.message || 'ไม่สามารถล้างข้อมูลได้', 'error');
    }
  }

  // ==========================================================================
  // 9. Dual ExcelJS Export & Local File Binding (ผูกไฟล์สำเนา & ผูกไฟล์ใบเบิก)
  // ==========================================================================
  const supportsFileBinding = () =>
    typeof window.showOpenFilePicker === 'function' && typeof window.showSaveFilePicker === 'function';

  function openHandleDb() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) return resolve(null);
      const req = window.indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(IDB_STORE)) {
          db.createObjectStore(IDB_STORE);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function saveHandleToDb(key, handle) {
    try {
      const db = await openHandleDb();
      if (!db) return;
      await new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, 'readwrite');
        const store = tx.objectStore(IDB_STORE);
        if (handle) store.put(handle, key);
        else store.delete(key);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      db.close();
    } catch {
      // ignore IndexedDB errors
    }
  }

  async function loadHandleFromDb(key) {
    try {
      const db = await openHandleDb();
      if (!db) return null;
      const val = await new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, 'readonly');
        const req = tx.objectStore(IDB_STORE).get(key);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      });
      db.close();
      return val;
    } catch {
      return null;
    }
  }

  const IDB_ZONE_SITE_PREFIX = 'zone_site::';
  const IDB_BILLING_SITE_PREFIX = 'billing_site::';
  const siteHandleKey = (kind, siteName) =>
    `${kind === 'zone' ? IDB_ZONE_SITE_PREFIX : IDB_BILLING_SITE_PREFIX}${String(siteName || '').trim()}`;

  async function loadAllSiteHandlesFromDb() {
    const out = { zone: {}, billing: {} };
    try {
      const db = await openHandleDb();
      if (!db) return out;
      const keys = await new Promise((resolve, reject) => {
        const tx = db.transaction(IDB_STORE, 'readonly');
        const req = tx.objectStore(IDB_STORE).getAllKeys();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => reject(req.error);
      });
      for (const k of keys) {
        if (typeof k !== 'string') continue;
        let bucket = null;
        let siteName = '';
        if (k.startsWith(IDB_ZONE_SITE_PREFIX)) {
          bucket = out.zone;
          siteName = k.slice(IDB_ZONE_SITE_PREFIX.length);
        } else if (k.startsWith(IDB_BILLING_SITE_PREFIX)) {
          bucket = out.billing;
          siteName = k.slice(IDB_BILLING_SITE_PREFIX.length);
        } else {
          continue;
        }
        const h = await new Promise((resolve, reject) => {
          const tx = db.transaction(IDB_STORE, 'readonly');
          const req = tx.objectStore(IDB_STORE).get(k);
          req.onsuccess = () => resolve(req.result || null);
          req.onerror = () => reject(req.error);
        });
        if (siteName && h && bucket) bucket[siteName] = h;
      }
      db.close();
    } catch {
      // ignore IndexedDB errors
    }
    return out;
  }

  function getKnownProjectNames() {
    const set = new Set();
    for (const s of state.sites || []) {
      if (s && s.active && s.site_name) set.add(s.site_name.trim());
    }
    for (const s of state.sites || []) {
      if (s && s.site_name) set.add(s.site_name.trim());
    }
    for (const l of state.logs || []) {
      if (l && l.site_name) set.add(String(l.site_name).trim());
    }
    for (const k of Object.keys(state.boundZoneHandlesBySite || {})) {
      if (k) set.add(k.trim());
    }
    for (const k of Object.keys(state.boundBillingHandlesBySite || {})) {
      if (k) set.add(k.trim());
    }
    const gsZoneBySite = (state.gsheet && state.gsheet.zone_by_site) || {};
    for (const k of Object.keys(gsZoneBySite)) {
      if (k) set.add(k.trim());
    }
    const gsBySite = (state.gsheet && state.gsheet.billing_by_site) || {};
    for (const k of Object.keys(gsBySite)) {
      if (k) set.add(k.trim());
    }
    return [...set].filter(Boolean);
  }

  function getLogCountBySite() {
    const counts = {};
    for (const l of state.logs || []) {
      const s = String((l && l.site_name) || '').trim();
      if (s) counts[s] = (counts[s] || 0) + 1;
    }
    return counts;
  }

  async function restoreBoundHandles() {
    if (!supportsFileBinding()) {
      const note = $('fs-api-unsupported-note');
      if (note) note.classList.remove('hidden');
      await updateBindingUI();
      return;
    }
    const [zoneHandle, billingHandle, siteHandles] = await Promise.all([
      loadHandleFromDb('zone'),
      loadHandleFromDb('billing'),
      loadAllSiteHandlesFromDb()
    ]);
    state.boundZoneHandle = zoneHandle;
    state.boundBillingHandle = billingHandle;
    state.boundZoneHandlesBySite = (siteHandles && siteHandles.zone) || {};
    state.boundBillingHandlesBySite = (siteHandles && siteHandles.billing) || {};
    await updateBindingUI();
  }

  async function renderPerSiteExcelBindingList() {
    const listEl = $('excel-per-site-list');
    const zoneCountBadge = $('excel-zone-sites-bound-count');
    const billingCountBadge = $('excel-sites-bound-count');
    const toggleBtn = $('btn-excel-sites-toggle');
    const searchEl = $('excel-sites-search');
    const syncAllZoneBtn = $('btn-sync-all-excel-zone-sites');
    const syncAllBillingBtn = $('btn-sync-all-excel-sites');
    if (!listEl) return;

    const zoneHandlesMap = state.boundZoneHandlesBySite || {};
    const billingHandlesMap = state.boundBillingHandlesBySite || {};
    const zoneBoundCount = Object.keys(zoneHandlesMap).length;
    const billingBoundCount = Object.keys(billingHandlesMap).length;
    if (zoneCountBadge) zoneCountBadge.textContent = `ผูกแล้ว ${formatInt(zoneBoundCount)} โครงการ`;
    if (billingCountBadge) billingCountBadge.textContent = `ผูกแล้ว ${formatInt(billingBoundCount)} โครงการ`;
    if (syncAllZoneBtn) {
      syncAllZoneBtn.innerHTML = `<i class="fa-solid fa-file-arrow-down"></i><span>เขียนสำเนาทุกโครงการที่ผูก (${formatInt(zoneBoundCount)})</span>`;
    }
    if (syncAllBillingBtn) {
      syncAllBillingBtn.innerHTML = `<i class="fa-solid fa-file-invoice"></i><span>เขียนใบเบิกทุกโครงการที่ผูก (${formatInt(billingBoundCount)})</span>`;
    }

    const logCounts = getLogCountBySite();
    const allProjects = getKnownProjectNames();
    const q = searchEl ? searchEl.value.trim().toLowerCase() : '';

    // เรียง: โครงการที่ผูกแล้วขึ้นก่อน -> โครงการที่มีเที่ยววิ่งในงวดนี้ -> ชื่อ ก-ฮ
    const sorted = [...allProjects].sort((a, b) => {
      const ab = zoneHandlesMap[a] || billingHandlesMap[a] ? 1 : 0;
      const bb = zoneHandlesMap[b] || billingHandlesMap[b] ? 1 : 0;
      if (ab !== bb) return bb - ab;
      const ac = logCounts[a] || 0;
      const bc = logCounts[b] || 0;
      if ((ac > 0) !== (bc > 0)) return bc > 0 ? 1 : -1;
      return a.localeCompare(b, 'th');
    });

    const important = sorted.filter(
      (s) => Boolean(zoneHandlesMap[s] || billingHandlesMap[s]) || (logCounts[s] || 0) > 0
    );
    const defaultList = state.excelSitesShowAll || q || important.length === 0 ? sorted : important;
    const filtered = q ? defaultList.filter((s) => s.toLowerCase().includes(q)) : defaultList;

    if (toggleBtn) {
      if (important.length > 0 && important.length < sorted.length && !q) {
        toggleBtn.classList.remove('hidden');
        toggleBtn.textContent = state.excelSitesShowAll
          ? `ย่อเฉพาะที่ผูก/มีเที่ยว (${important.length})`
          : `แสดงทุกโครงการ (${sorted.length})`;
      } else {
        toggleBtn.classList.add('hidden');
      }
    }

    if (filtered.length === 0) {
      listEl.innerHTML = allProjects.length === 0
        ? '<div class="col-span-full py-3 text-center text-xs text-slate-400">ยังไม่มีโครงการ — เพิ่มโครงการในหน้า "ตั้งค่า" ก่อน เพื่อผูกไฟล์สำเนาและใบเบิกแยกแต่ละโครงการ</div>'
        : '<div class="col-span-full py-3 text-center text-xs text-slate-400">ไม่พบโครงการที่ค้นหา</div>';
      return;
    }

    const zonePermMap = {};
    const billingPermMap = {};
    await Promise.all(
      filtered.map(async (s) => {
        const zh = zoneHandlesMap[s];
        if (zh) {
          try {
            zonePermMap[s] = await zh.queryPermission({ mode: 'readwrite' });
          } catch {
            zonePermMap[s] = 'prompt';
          }
        }
        const bh = billingHandlesMap[s];
        if (bh) {
          try {
            billingPermMap[s] = await bh.queryPermission({ mode: 'readwrite' });
          } catch {
            billingPermMap[s] = 'prompt';
          }
        }
      })
    );

    listEl.innerHTML = filtered
      .map((siteName) => {
        const zHandle = zoneHandlesMap[siteName];
        const bHandle = billingHandlesMap[siteName];
        const trips = logCounts[siteName] || 0;
        const zDotCls = (zonePermMap[siteName] || 'prompt') === 'granted' ? 'bg-emerald-500' : 'bg-amber-400';
        const bDotCls = (billingPermMap[siteName] || 'prompt') === 'granted' ? 'bg-emerald-500' : 'bg-amber-400';
        const hasAnyBound = Boolean(zHandle || bHandle);
        return `
          <div class="p-2.5 rounded-xl bg-white border ${hasAnyBound ? 'border-sky-300 shadow-xs' : 'border-slate-200'} flex flex-col gap-2">
            <div class="flex items-center justify-between gap-2 min-w-0">
              <div class="flex items-center gap-1.5 min-w-0">
                <i class="fa-solid fa-building text-amber-500 text-[11px] shrink-0"></i>
                <span class="font-bold text-xs text-slate-900 truncate" title="${escapeHtml(siteName)}">${escapeHtml(siteName)}</span>
              </div>
              ${trips > 0 ? `<span class="px-1.5 py-0.5 rounded-full bg-indigo-50 text-indigo-700 text-[10px] font-bold shrink-0">${formatInt(trips)} เที่ยว</span>` : ''}
            </div>

            <!-- แถวสำเนาโซน (เฉพาะโครงการนี้) -->
            <div class="p-2 rounded-lg bg-amber-50/70 border border-amber-200/80 space-y-1.5">
              <div class="flex items-center justify-between gap-1.5 min-w-0">
                <div class="flex items-center gap-1.5 min-w-0">
                  <span class="px-1.5 py-0.5 rounded bg-amber-200/80 text-amber-950 text-[10px] font-bold shrink-0">สำเนา</span>
                  ${
                    zHandle
                      ? `<span class="w-2 h-2 rounded-full ${zDotCls} shrink-0"></span><span class="text-[11px] text-amber-950 font-semibold truncate" title="${escapeHtml(zHandle.name)}">${escapeHtml(zHandle.name)}</span>`
                      : '<span class="text-[11px] text-slate-400 truncate">ยังไม่ผูกไฟล์สำเนา</span>'
                  }
                </div>
                ${
                  zHandle
                    ? `<button type="button" data-excel-zone-site-unbind="${escapeHtml(siteName)}" title="ยกเลิกการผูกไฟล์สำเนาของโครงการนี้"
                        class="w-6 h-6 rounded-md bg-white hover:bg-rose-50 text-slate-400 hover:text-rose-600 border border-amber-200 inline-flex items-center justify-center cursor-pointer shrink-0">
                        <i class="fa-solid fa-link-slash text-[10px]"></i>
                      </button>`
                    : ''
                }
              </div>
              <div class="flex flex-wrap items-center gap-1.5">
                <button type="button" data-excel-zone-site-bind="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg bg-white hover:bg-amber-100 text-amber-950 border border-amber-300 text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition">
                  <i class="fa-solid fa-link text-[10px]"></i><span>${zHandle ? 'เปลี่ยนไฟล์' : 'ผูกสำเนา'}</span>
                </button>
                <button type="button" data-excel-zone-site-sync="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg ${zHandle ? 'bg-amber-600 hover:bg-amber-700 text-white' : 'bg-white/90 hover:bg-amber-600 hover:text-white text-amber-900 border border-amber-200'} text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition"
                  title="เขียนเฉพาะข้อมูลสำเนาของโครงการ ${escapeHtml(siteName)} ลงไฟล์ Excel">
                  <i class="fa-solid fa-file-arrow-down text-[10px]"></i><span>เขียนสำเนา</span>
                </button>
                <button type="button" data-excel-zone-site-dl="${escapeHtml(siteName)}"
                  class="min-h-[28px] px-2 py-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition shrink-0"
                  title="ดาวน์โหลดไฟล์สำเนา (.xlsx) เฉพาะโครงการ ${escapeHtml(siteName)}">
                  <i class="fa-solid fa-download text-[10px]"></i><span>โหลด</span>
                </button>
              </div>
            </div>

            <!-- แถวใบเบิก (เฉพาะโครงการนี้) -->
            <div class="p-2 rounded-lg bg-sky-50/70 border border-sky-200/80 space-y-1.5">
              <div class="flex items-center justify-between gap-1.5 min-w-0">
                <div class="flex items-center gap-1.5 min-w-0">
                  <span class="px-1.5 py-0.5 rounded bg-sky-200/80 text-sky-950 text-[10px] font-bold shrink-0">ใบเบิก</span>
                  ${
                    bHandle
                      ? `<span class="w-2 h-2 rounded-full ${bDotCls} shrink-0"></span><span class="text-[11px] text-sky-900 font-semibold truncate" title="${escapeHtml(bHandle.name)}">${escapeHtml(bHandle.name)}</span>`
                      : '<span class="text-[11px] text-slate-400 truncate">ยังไม่ผูกไฟล์ใบเบิก</span>'
                  }
                </div>
                ${
                  bHandle
                    ? `<button type="button" data-excel-site-unbind="${escapeHtml(siteName)}" title="ยกเลิกการผูกไฟล์ใบเบิกของโครงการนี้"
                        class="w-6 h-6 rounded-md bg-white hover:bg-rose-50 text-slate-400 hover:text-rose-600 border border-sky-200 inline-flex items-center justify-center cursor-pointer shrink-0">
                        <i class="fa-solid fa-link-slash text-[10px]"></i>
                      </button>`
                    : ''
                }
              </div>
              <div class="flex flex-wrap items-center gap-1.5">
                <button type="button" data-excel-site-bind="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg bg-white hover:bg-sky-100 text-sky-900 border border-sky-300 text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition">
                  <i class="fa-solid fa-link text-[10px]"></i><span>${bHandle ? 'เปลี่ยนไฟล์' : 'ผูกใบเบิก'}</span>
                </button>
                <button type="button" data-excel-site-sync="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg ${bHandle ? 'bg-sky-600 hover:bg-sky-700 text-white' : 'bg-white/90 hover:bg-sky-600 hover:text-white text-sky-900 border border-sky-200'} text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition"
                  title="เขียนเฉพาะข้อมูลใบเบิกของโครงการ ${escapeHtml(siteName)} ลงไฟล์ Excel">
                  <i class="fa-solid fa-file-arrow-down text-[10px]"></i><span>เขียนใบเบิก</span>
                </button>
                <button type="button" data-excel-site-dl="${escapeHtml(siteName)}"
                  class="min-h-[28px] px-2 py-1 rounded-lg bg-zinc-900 hover:bg-zinc-800 text-white text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition shrink-0"
                  title="ดาวน์โหลดไฟล์ใบเบิก (.xlsx) เฉพาะโครงการ ${escapeHtml(siteName)}">
                  <i class="fa-solid fa-download text-amber-400 text-[10px]"></i><span>โหลด</span>
                </button>
              </div>
            </div>
          </div>
        `;
      })
      .join('');
  }

  async function updateBindingUI() {
    const autoCheck = $('toggle-auto-sync');
    if (autoCheck) autoCheck.checked = Boolean(state.autoSync);

    const legacyZoneWrap = $('legacy-zone-bind-wrap');
    if (legacyZoneWrap) legacyZoneWrap.classList.toggle('hidden', !state.boundZoneHandle);

    const legacyWrap = $('legacy-billing-bind-wrap');
    if (legacyWrap) legacyWrap.classList.toggle('hidden', !state.boundBillingHandle);

    const updateOne = async (kind, handle, labelId, dotId, unbindId) => {
      const label = $(labelId);
      const dot = $(dotId);
      const unbindBtn = $(unbindId);
      if (!label) return;

      if (!handle) {
        label.textContent = 'ยังไม่ได้เลือก (คลิกเพื่อผูกไฟล์)';
        if (dot) dot.classList.add('hidden');
        if (unbindBtn) {
          unbindBtn.classList.add('hidden');
          unbindBtn.classList.remove('inline-flex');
        }
        return;
      }

      const siteSuffix =
        kind === 'billing' && state.boundBillingSite ? ` [${state.boundBillingSite}]` : '';
      label.textContent = `${handle.name}${siteSuffix}`;
      if (unbindBtn) {
        unbindBtn.classList.remove('hidden');
        unbindBtn.classList.add('inline-flex');
      }
      if (dot) {
        dot.classList.remove('hidden');
        let perm = 'prompt';
        try {
          perm = await handle.queryPermission({ mode: 'readwrite' });
        } catch {
          perm = 'prompt';
        }
        dot.className = `w-2.5 h-2.5 rounded-full ml-auto shrink-0 ${
          perm === 'granted' ? 'bg-emerald-500' : 'bg-amber-400'
        }`;
        dot.title =
          perm === 'granted'
            ? 'พร้อมเขียนไฟล์อัตโนมัติ'
            : 'คลิกปุ่ม "บันทึกไฟล์..." 1 ครั้งเพื่ออนุญาตสิทธิ์เขียนไฟล์';
      }
    };

    await Promise.all([
      updateOne('zone', state.boundZoneHandle, 'bind-zone-label', 'bind-zone-dot', 'btn-unbind-zone'),
      updateOne('billing', state.boundBillingHandle, 'bind-billing-label', 'bind-billing-dot', 'btn-unbind-billing'),
      renderPerSiteExcelBindingList()
    ]);
  }

  async function downloadWorkbook(wb, filename) {
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function findSiteMeta(siteName) {
    return state.sites.find((s) => s.site_name === siteName) || null;
  }

  async function buildZonePayload({ siteName, dateFrom, dateTo }) {
    const params = new URLSearchParams();
    if (siteName) params.set('site_name', siteName);
    if (dateFrom) params.set('date_from', dateFrom);
    if (dateTo) params.set('date_to', dateTo);
    const logs = await apiFetch(`/api/logs${params.toString() ? `?${params.toString()}` : ''}`);
    const siteMeta = siteName ? findSiteMeta(siteName) : null;
    const zoneName = siteMeta ? siteMeta.zone_name : '';

    const wb = window.CkpExcel.buildZoneCopyWorkbook(window.ExcelJS, logs, { siteName, zoneName });
    const { start, end } = window.CkpExcel.dateRangeOf(logs, dateFrom, dateTo);
    const filename = window.CkpExcel.zoneCopyFilename({ zoneName, siteName, start, end });
    return { wb, filename, logs };
  }

  async function buildZonePayloadFromFilters() {
    const f = getFilterParams();
    const params = new URLSearchParams();
    if (f.date_from) params.set('date_from', f.date_from);
    if (f.date_to) params.set('date_to', f.date_to);
    if (f.site_name) params.set('site_name', f.site_name);
    if (f.trip_type) params.set('trip_type', f.trip_type);
    const logs = await apiFetch(`/api/logs${params.toString() ? `?${params.toString()}` : ''}`);

    const distinctSites = [...new Set(logs.map((l) => (l.site_name || '').trim()).filter(Boolean))];
    const siteName = f.site_name || (distinctSites.length === 1 ? distinctSites[0] : '');
    const siteMeta = siteName ? findSiteMeta(siteName) : null;
    const zoneName = siteMeta ? siteMeta.zone_name : '';

    const wb = window.CkpExcel.buildZoneCopyWorkbook(window.ExcelJS, logs, { siteName, zoneName });
    const { start, end } = window.CkpExcel.dateRangeOf(logs, f.date_from, f.date_to);
    const filename = window.CkpExcel.zoneCopyFilename({ zoneName, siteName, start, end });
    return { wb, filename, logs };
  }

  async function buildBillingPayload({ siteName, dateFrom, dateTo, issueDate }) {
    const params = new URLSearchParams();
    if (siteName) params.set('site_name', siteName);
    if (dateFrom) params.set('date_from', dateFrom);
    if (dateTo) params.set('date_to', dateTo);

    const logs = await apiFetch(`/api/logs${params.toString() ? `?${params.toString()}` : ''}`);
    const wb = window.CkpExcel.buildBillingWorkbook(window.ExcelJS, logs, {
      siteName,
      dateFrom,
      dateTo,
      issueDate: issueDate || todayIso()
    });
    const { start, end } = window.CkpExcel.dateRangeOf(logs, dateFrom, dateTo);
    const filename = window.CkpExcel.billingFilename({ siteName, start, end });
    return { wb, filename, logs };
  }

  function isFileLockedError(err) {
    if (!err) return false;
    const name = String(err.name || '');
    const msg = String(err.message || '').toLowerCase();
    return (
      name === 'NoModificationAllowedError' ||
      name === 'InvalidStateError' ||
      name === 'NotReadableError' ||
      msg.includes('lock') ||
      msg.includes('in use') ||
      msg.includes('busy') ||
      msg.includes('modified') ||
      msg.includes('permission')
    );
  }

  function showFileLockedModal({ kind, siteName = '', handle, wb, suggestedName, onRetry }) {
    const cleanSite = String(siteName || '').trim();
    const baseLabel = kind === 'billing' ? 'ใบเบิกโครงการ' : 'สำเนาโซน';
    const typeLabel = cleanSite ? `${baseLabel} (${cleanSite})` : baseLabel;
    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <div class="flex items-start gap-3">
          <div class="w-11 h-11 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center shrink-0">
            <i class="fa-solid fa-file-circle-exclamation text-lg"></i>
          </div>
          <div>
            <h3 class="text-base font-bold text-slate-900">ไฟล์ Excel (${escapeHtml(typeLabel)}) กำลังเปิดค้างอยู่ในโปรแกรม Excel</h3>
            <p class="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
              กรุณาปิดโปรแกรม Microsoft Excel ที่เปิดไฟล์ <strong class="text-slate-900">"${escapeHtml(handle.name)}"</strong> ค้างไว้ก่อน แล้วกดปุ่ม <strong>"ลองบันทึกอีกครั้ง"</strong> หรือเลือกบันทึกเป็นไฟล์ใหม่ครับ
            </p>
          </div>
        </div>
        <div class="flex flex-wrap items-center justify-end gap-2 pt-2 border-t border-slate-100">
          <button type="button" id="modal-lock-close" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs sm:text-sm font-medium cursor-pointer">ปิดหน้าต่าง</button>
          <button type="button" id="modal-lock-saveas" class="min-h-[40px] px-4 py-2 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs sm:text-sm font-semibold cursor-pointer">
            <i class="fa-solid fa-floppy-disk mr-1"></i>บันทึกเป็นไฟล์ใหม่
          </button>
          <button type="button" id="modal-lock-retry" class="min-h-[40px] px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-semibold cursor-pointer">
            <i class="fa-solid fa-rotate-right mr-1"></i>ลองบันทึกอีกครั้ง
          </button>
        </div>
      </div>
      `,
      (box) => {
        box.querySelector('#modal-lock-close').addEventListener('click', closeModal);
        box.querySelector('#modal-lock-retry').addEventListener('click', async () => {
          closeModal();
          if (typeof onRetry === 'function') await onRetry();
        });
        box.querySelector('#modal-lock-saveas').addEventListener('click', async () => {
          try {
            const newHandle = await window.showSaveFilePicker({
              suggestedName: suggestedName || handle.name,
              types: [
                {
                  description: 'Excel Workbook (.xlsx)',
                  accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
                }
              ]
            });
            const buffer = await wb.xlsx.writeBuffer();
            const writable = await newHandle.createWritable();
            await writable.write(buffer);
            await writable.close();
            if (cleanSite) {
              if (kind === 'zone') state.boundZoneHandlesBySite[cleanSite] = newHandle;
              else state.boundBillingHandlesBySite[cleanSite] = newHandle;
              await saveHandleToDb(siteHandleKey(kind, cleanSite), newHandle);
            } else if (kind === 'zone') {
              state.boundZoneHandle = newHandle;
              await saveHandleToDb('zone', newHandle);
            } else {
              state.boundBillingHandle = newHandle;
              await saveHandleToDb('billing', newHandle);
            }
            await updateBindingUI();
            closeModal();
            showToast(`บันทึกและผูกไฟล์ใหม่ "${newHandle.name}" เรียบร้อยแล้ว`, 'success', 4500);
          } catch (e) {
            if (e && e.name === 'AbortError') return;
            showToast(`ไม่สามารถบันทึกไฟล์ใหม่ได้: ${e.message}`, 'error');
          }
        });
      }
    );
  }

  async function writeWorkbookToHandle(kind, handle, wb, suggestedName, { interactive = true, siteName = '' } = {}) {
    const cleanSite = String(siteName || '').trim();
    let perm = 'prompt';
    try {
      perm = await handle.queryPermission({ mode: 'readwrite' });
    } catch {
      perm = 'prompt';
    }

    if (perm !== 'granted') {
      if (!interactive) {
        return { status: 'need_permission' };
      }
      perm = await handle.requestPermission({ mode: 'readwrite' });
      if (perm !== 'granted') {
        showToast('ไม่ได้รับสิทธิ์เขียนไฟล์ กรุณากดอนุญาตเมื่อเบราว์เซอร์สอบถาม', 'warning', 4500);
        return { status: 'denied' };
      }
    }

    try {
      const buffer = await wb.xlsx.writeBuffer();
      const writable = await handle.createWritable();
      await writable.write(buffer);
      await writable.close();
      await updateBindingUI();
      return { status: 'ok' };
    } catch (err) {
      if (isFileLockedError(err)) {
        showFileLockedModal({
          kind,
          siteName: cleanSite,
          handle,
          wb,
          suggestedName,
          onRetry: () =>
            cleanSite
              ? syncSiteExcelFile(kind, cleanSite, { interactive: true })
              : syncToBoundFile(kind, { interactive: true })
        });
        return { status: 'locked' };
      }
      throw err;
    }
  }

  function openBindFileModal(kind, fixedSiteName = '') {
    const isBilling = kind === 'billing';
    const fixedSite = String(fixedSiteName || '').trim();
    const kindThai = isBilling ? 'ใบเบิก' : 'สำเนา';
    const title = fixedSite
      ? `ผูกไฟล์ Excel ${kindThai}เฉพาะโครงการ: ${fixedSite}`
      : isBilling
        ? 'ผูกไฟล์ที่ 2: ใบเบิกโครงการ (.xlsx)'
        : 'ผูกไฟล์ที่ 1: สำเนาโซน (.xlsx)';
    const desc = fixedSite
      ? `เลือกไฟล์ .xlsx เดิม หรือสร้างไฟล์ใหม่เพื่อผูก${kindThai}ไว้เฉพาะโครงการ "${fixedSite}" (จะเขียนเฉพาะเที่ยววิ่งของโครงการนี้เท่านั้น ไม่รวมกับโครงการอื่น)`
      : isBilling
        ? 'เลือกไฟล์ใบเบิกเดิมบนเครื่อง หรือสร้างไฟล์ใหม่ไว้ผูกกับระบบ เพื่อให้กดบันทึกข้อมูล (พร้อมโลโก้และสูตรคำนวณ) ลงไฟล์ได้ทันที'
        : 'เลือกไฟล์สำเนาโซนเดิมบนเครื่อง หรือสร้างไฟล์ใหม่ไว้ผูกกับระบบ เพื่อให้กดบันทึกข้อมูล (พร้อมสรุปเที่ยว COUNTIF/SUM) ลงไฟล์ได้ทันที';

    if (!supportsFileBinding()) {
      openModal(
        `
        <div class="p-5 sm:p-6 space-y-4">
          <h3 class="text-base font-bold text-slate-900">${escapeHtml(title)}</h3>
          <p class="text-xs sm:text-sm text-slate-600 leading-relaxed">
            เบราว์เซอร์นี้ยังไม่รองรับการผูกและเขียนทับไฟล์บนเครื่องโดยตรง (File System Access API) กรุณาใช้งานผ่าน <strong>Google Chrome</strong> หรือ <strong>Microsoft Edge</strong> บนคอมพิวเตอร์ หรือกดปุ่ม <strong>โหลด (.xlsx)</strong> / <strong>ไหลลง Google Sheets</strong> แทนครับ
          </p>
          <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="m-bind-close" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm cursor-pointer">ปิด</button>
            <button type="button" id="m-bind-dl" class="min-h-[40px] px-4 py-2 rounded-xl bg-zinc-900 text-white text-xs sm:text-sm font-semibold cursor-pointer">ดาวน์โหลดไฟล์ (.xlsx) แทน</button>
          </div>
        </div>
        `,
        (box) => {
          box.querySelector('#m-bind-close').addEventListener('click', closeModal);
          box.querySelector('#m-bind-dl').addEventListener('click', async () => {
            closeModal();
            const f = getFilterParams();
            if (fixedSite && isBilling) {
              await runBillingExport({ siteName: fixedSite, dateFrom: f.date_from, dateTo: f.date_to, issueDate: todayIso() });
            } else if (fixedSite && !isBilling) {
              await runZoneExport({ siteName: fixedSite, dateFrom: f.date_from, dateTo: f.date_to });
            } else if (isBilling) {
              handleExportBilling();
            } else {
              handleExportZoneCopy();
            }
          });
        }
      );
      return;
    }

    const siteOptions = getKnownProjectNames();
    const currentSite = fixedSite || state.boundBillingSite || $('filter-site').value.trim();

    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <div class="flex items-start gap-3">
          <div class="w-10 h-10 rounded-xl ${isBilling ? 'bg-sky-100 text-sky-700' : 'bg-amber-100 text-amber-700'} flex items-center justify-center shrink-0">
            <i class="fa-solid fa-link"></i>
          </div>
          <div>
            <h3 class="text-base font-bold text-slate-900">${escapeHtml(title)}</h3>
            <p class="text-xs text-slate-500 mt-0.5 leading-relaxed">${escapeHtml(desc)}</p>
          </div>
        </div>

        ${
          fixedSite
            ? `
          <div class="p-3 rounded-xl ${isBilling ? 'bg-sky-50 border-sky-200 text-sky-950' : 'bg-amber-50 border-amber-200 text-amber-950'} border text-xs font-semibold flex items-center gap-2">
            <i class="fa-solid fa-building text-amber-500"></i>
            <span>โครงการที่ผูก (${escapeHtml(kindThai)}): <strong>${escapeHtml(fixedSite)}</strong> (แยกเฉพาะโครงการนี้ ไม่รวมโครงการอื่น)</span>
          </div>
        `
            : isBilling
              ? `
          <div>
            <label for="m-bind-billing-site" class="block text-xs font-semibold text-slate-700 mb-1">
              เลือกโครงการที่ต้องการผูกไฟล์ (แนะนำให้เลือกแยกแต่ละโครงการ)
            </label>
            <select id="m-bind-billing-site" class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-sky-600">
              <option value="">-- ใช้ตามตัวกรองโครงการบนหน้าเว็บ --</option>
              ${siteOptions
                .map(
                  (s) =>
                    `<option value="${escapeHtml(s)}" ${s === currentSite ? 'selected' : ''}>${escapeHtml(s)}</option>`
                )
                .join('')}
            </select>
          </div>
        `
              : ''
        }

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
          <button type="button" id="m-pick-existing"
            class="p-4 rounded-2xl border-2 border-slate-200 hover:border-indigo-500 hover:bg-indigo-50/40 text-left transition cursor-pointer space-y-1.5">
            <div class="w-9 h-9 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center">
              <i class="fa-solid fa-folder-open"></i>
            </div>
            <p class="text-sm font-bold text-slate-900">เลือกไฟล์ที่มีอยู่แล้ว</p>
            <p class="text-[11px] text-slate-500 leading-snug">เลือกไฟล์ .xlsx เดิมบนเครื่องเพื่อผูกไว้สำหรับกดบันทึกอัปเดต</p>
          </button>

          <button type="button" id="m-create-new"
            class="p-4 rounded-2xl border-2 border-slate-200 hover:border-emerald-500 hover:bg-emerald-50/40 text-left transition cursor-pointer space-y-1.5">
            <div class="w-9 h-9 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
              <i class="fa-solid fa-file-circle-plus"></i>
            </div>
            <p class="text-sm font-bold text-slate-900">สร้างไฟล์ใหม่แล้วผูกเลย</p>
            <p class="text-[11px] text-slate-500 leading-snug">สร้างไฟล์ .xlsx ฟอร์มบริษัทใหม่ลงเครื่องพร้อมเขียนข้อมูลทันที</p>
          </button>
        </div>

        <div class="flex justify-end pt-2 border-t border-slate-100">
          <button type="button" id="m-bind-cancel" class="min-h-[38px] px-4 py-1.5 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-medium cursor-pointer">ปิดหน้าต่าง</button>
        </div>
      </div>
      `,
      (box) => {
        const readChosenSite = () => {
          if (fixedSite) return fixedSite;
          if (!isBilling) return '';
          const sel = box.querySelector('#m-bind-billing-site');
          const val = sel ? sel.value.trim() : '';
          state.boundBillingSite = val;
          localStorage.setItem(BOUND_BILLING_SITE_KEY, val);
          return val;
        };

        const storeBoundHandle = async (chosenSite, handle) => {
          if (chosenSite) {
            if (isBilling) state.boundBillingHandlesBySite[chosenSite] = handle;
            else state.boundZoneHandlesBySite[chosenSite] = handle;
            await saveHandleToDb(siteHandleKey(kind, chosenSite), handle);
          } else if (isBilling) {
            state.boundBillingHandle = handle;
            await saveHandleToDb('billing', handle);
          } else {
            state.boundZoneHandle = handle;
            await saveHandleToDb('zone', handle);
          }
        };

        box.querySelector('#m-bind-cancel').addEventListener('click', closeModal);

        box.querySelector('#m-pick-existing').addEventListener('click', async () => {
          const chosenSite = readChosenSite();
          try {
            const [handle] = await window.showOpenFilePicker({
              multiple: false,
              types: [
                {
                  description: 'Excel Workbook (.xlsx)',
                  accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
                }
              ]
            });
            if (!handle) return;
            const perm = await handle.requestPermission({ mode: 'readwrite' });
            await storeBoundHandle(chosenSite, handle);
            await updateBindingUI();
            closeModal();
            const targetLabel = chosenSite ? `${kindThai}โครงการ "${chosenSite}"` : kindThai;
            if (perm !== 'granted') {
              showToast(
                `ผูกไฟล์ "${handle.name}" ให้${targetLabel} แล้ว — กดปุ่มเขียนไฟล์ 1 ครั้งเพื่ออนุญาตสิทธิ์`,
                'warning',
                4800
              );
            } else {
              showToast(
                `ผูกไฟล์ "${handle.name}" ให้${targetLabel} เรียบร้อยแล้ว`,
                'success',
                4500
              );
            }
          } catch (err) {
            if (err && err.name === 'AbortError') return;
            showToast(`ไม่สามารถผูกไฟล์ได้: ${err.message}`, 'error');
          }
        });

        box.querySelector('#m-create-new').addEventListener('click', async () => {
          const chosenSite = readChosenSite();
          try {
            // คำนวณชื่อไฟล์แนะนำแบบ synchronous และเปิด showSaveFilePicker ทันทีเพื่อรักษา User Gesture ใน Chrome/Edge
            const f = getFilterParams();
            const targetSite = chosenSite || f.site_name;
            const siteMeta = targetSite ? findSiteMeta(targetSite) : null;
            const zoneName = siteMeta ? siteMeta.zone_name : '';
            const { start, end } = window.CkpExcel.dateRangeOf(state.logs, f.date_from, f.date_to);
            const suggestedName = isBilling
              ? window.CkpExcel.billingFilename({ siteName: targetSite, start, end })
              : window.CkpExcel.zoneCopyFilename({ zoneName, siteName: targetSite, start, end });

            const handle = await window.showSaveFilePicker({
              suggestedName,
              types: [
                {
                  description: 'Excel Workbook (.xlsx)',
                  accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
                }
              ]
            });

            const payload = isBilling
              ? await buildBillingPayload({
                  siteName: chosenSite || f.site_name,
                  dateFrom: f.date_from,
                  dateTo: f.date_to,
                  issueDate: todayIso()
                })
              : chosenSite
                ? await buildZonePayload({
                    siteName: chosenSite,
                    dateFrom: f.date_from,
                    dateTo: f.date_to
                  })
                : await buildZonePayloadFromFilters();

            const buffer = await payload.wb.xlsx.writeBuffer();
            const writable = await handle.createWritable();
            await writable.write(buffer);
            await writable.close();

            await storeBoundHandle(chosenSite, handle);
            await updateBindingUI();
            closeModal();
            showToast(
              `สร้างและผูกไฟล์${kindThai} "${handle.name}"${chosenSite ? ` สำหรับโครงการ "${chosenSite}"` : ''} สำเร็จ (${payload.wb.worksheets.length} ชีต, ${payload.logs.length} เที่ยว)`,
              'success',
              4500
            );
          } catch (err) {
            if (err && err.name === 'AbortError') return;
            showToast(`ไม่สามารถสร้างไฟล์ได้: ${err.message}`, 'error', 5000);
          }
        });
      }
    );
  }

  async function unbindSiteExcelFile(kind, siteName) {
    const cleanSite = String(siteName || '').trim();
    const isZone = kind === 'zone';
    const map = isZone ? state.boundZoneHandlesBySite : state.boundBillingHandlesBySite;
    const handle = map[cleanSite];
    if (!handle) return;
    const kindThai = isZone ? 'สำเนา' : 'ใบเบิก';
    const ok = await confirmModal({
      title: `ยกเลิกการผูกไฟล์${kindThai}โครงการ "${cleanSite}"`,
      message: `ต้องการยกเลิกการผูกไฟล์ "${handle.name}" ออกจาก${kindThai}โครงการ "${cleanSite}" ใช่หรือไม่? (ไฟล์บนเครื่องของคุณจะไม่ถูกลบ)`,
      confirmText: 'ยกเลิกการผูก'
    });
    if (!ok) return;
    delete map[cleanSite];
    await saveHandleToDb(siteHandleKey(kind, cleanSite), null);
    await updateBindingUI();
    showToast(`ยกเลิกการผูกไฟล์${kindThai}ของโครงการ "${cleanSite}" เรียบร้อยแล้ว`, 'info');
  }

  async function syncSiteExcelFile(kind, siteName, { interactive = true, triggerBtn = null } = {}) {
    if (typeof window.ExcelJS === 'undefined' || !window.CkpExcel) {
      if (interactive) showToast('ไม่พบไลบรารี ExcelJS กรุณารีเฟรชหน้าเว็บ', 'error');
      return { status: 'error' };
    }
    const isZone = kind === 'zone';
    const kindThai = isZone ? 'สำเนา' : 'ใบเบิก';
    const cleanSite = String(siteName || '').trim();
    const map = isZone ? state.boundZoneHandlesBySite : state.boundBillingHandlesBySite;
    const handle = map[cleanSite];
    if (!handle) {
      if (interactive) openBindFileModal(kind, cleanSite);
      return { status: 'unbound' };
    }

    const origHtml = triggerBtn ? triggerBtn.innerHTML : '';
    if (triggerBtn && interactive) {
      triggerBtn.disabled = true;
      triggerBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin text-[10px]"></i><span>กำลังเขียน...</span>';
    }

    try {
      const f = getFilterParams();
      // กรองเฉพาะเที่ยววิ่งของโครงการนี้เท่านั้น ไม่รวมกับโครงการอื่น
      const payload = isZone
        ? await buildZonePayload({
            siteName: cleanSite,
            dateFrom: f.date_from,
            dateTo: f.date_to
          })
        : await buildBillingPayload({
            siteName: cleanSite,
            dateFrom: f.date_from,
            dateTo: f.date_to,
            issueDate: todayIso()
          });
      const res = await writeWorkbookToHandle(kind, handle, payload.wb, payload.filename, {
        interactive,
        siteName: cleanSite
      });
      if (res.status === 'ok' && interactive) {
        showToast(
          `เขียนไฟล์${kindThai}โครงการ "${cleanSite}" (${handle.name}) สำเร็จ (${payload.wb.worksheets.length} ชีต, ${payload.logs.length} เที่ยว)`,
          'success',
          4200
        );
      }
      return { ...res, siteName: cleanSite, name: handle.name, sheets: payload.wb.worksheets.length, count: payload.logs.length };
    } catch (err) {
      if (interactive) showToast(`เขียนไฟล์${kindThai}โครงการ "${cleanSite}" ไม่สำเร็จ: ${err.message}`, 'error', 5000);
      return { status: 'error', siteName: cleanSite, error: err };
    } finally {
      if (triggerBtn && interactive) {
        triggerBtn.disabled = false;
        triggerBtn.innerHTML = origHtml;
      }
    }
  }

  async function syncAllBoundExcelSites(kind = 'billing') {
    const isZone = kind === 'zone';
    const kindThai = isZone ? 'สำเนา' : 'ใบเบิก';
    const map = isZone ? state.boundZoneHandlesBySite : state.boundBillingHandlesBySite;
    const sites = Object.keys(map || {});
    if (sites.length === 0) {
      showToast(`ยังไม่ได้ผูกไฟล์${kindThai} Excel ให้โครงการใด — กดปุ่ม "ผูก${kindThai}" ที่การ์ดโครงการด้านล่างก่อนครับ`, 'warning', 4500);
      return;
    }
    const btn = $(isZone ? 'btn-sync-all-excel-zone-sites' : 'btn-sync-all-excel-sites');
    const origHtml = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i><span>กำลังเขียน${kindThai}ทุกโครงการ...</span>`;
    }
    try {
      const okNames = [];
      let failCount = 0;
      for (const sName of sites) {
        const r = await syncSiteExcelFile(kind, sName, { interactive: true });
        if (r.status === 'ok') okNames.push(sName);
        else if (r.status === 'error') failCount += 1;
      }
      if (okNames.length > 0) {
        showToast(`เขียนไฟล์${kindThai}แยกสำเร็จครบ ${okNames.length} โครงการ: ${okNames.join(', ')}`, 'success', 5000);
      } else if (failCount > 0) {
        showToast(`ไม่สามารถเขียนไฟล์${kindThai}บางโครงการได้ กรุณาตรวจสอบว่าไม่ได้เปิดไฟล์ค้างไว้ใน Excel`, 'error', 5000);
      }
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = origHtml;
      }
      await updateBindingUI();
    }
  }

  async function unbindFile(kind) {
    const handle = kind === 'billing' ? state.boundBillingHandle : state.boundZoneHandle;
    if (!handle) return;
    const label = kind === 'billing' ? 'ใบเบิกโครงการ' : 'สำเนาโซน';
    const ok = await confirmModal({
      title: `ยกเลิกการผูกไฟล์${label}`,
      message: `ต้องการยกเลิกการผูกไฟล์ "${handle.name}" ออกจากหน้าเว็บใช่หรือไม่? (ไฟล์บนเครื่องของคุณจะไม่ถูกลบ)`,
      confirmText: 'ยกเลิกการผูก'
    });
    if (!ok) return;
    if (kind === 'billing') {
      state.boundBillingHandle = null;
      state.boundBillingSite = '';
      localStorage.removeItem(BOUND_BILLING_SITE_KEY);
    } else {
      state.boundZoneHandle = null;
    }
    await saveHandleToDb(kind, null);
    await updateBindingUI();
    showToast(`ยกเลิกการผูกไฟล์${label}เรียบร้อยแล้ว`, 'info');
  }

  async function syncToBoundFile(kind, { interactive = true, billingOpts = null } = {}) {
    if (typeof window.ExcelJS === 'undefined' || !window.CkpExcel) {
      if (interactive) showToast('ไม่พบไลบรารี ExcelJS กรุณารีเฟรชหน้าเว็บ', 'error');
      return { status: 'error' };
    }

    const isBilling = kind === 'billing';
    const handle = isBilling ? state.boundBillingHandle : state.boundZoneHandle;
    if (!handle) {
      if (interactive) openBindFileModal(kind);
      return { status: 'unbound' };
    }

    const btn = $(isBilling ? 'btn-sync-billing' : 'btn-sync-zone');
    if (btn && interactive) btn.disabled = true;

    try {
      const f = getFilterParams();
      let payload;
      if (isBilling) {
        const opts = billingOpts || {
          siteName: state.boundBillingSite || f.site_name || '',
          dateFrom: f.date_from,
          dateTo: f.date_to,
          issueDate: todayIso()
        };
        payload = await buildBillingPayload(opts);
      } else {
        payload = await buildZonePayloadFromFilters();
      }

      const res = await writeWorkbookToHandle(kind, handle, payload.wb, payload.filename, { interactive });
      if (res.status === 'ok' && interactive) {
        showToast(
          `บันทึกลงไฟล์ "${handle.name}" สำเร็จ (${payload.wb.worksheets.length} ชีต, ${payload.logs.length} เที่ยว)`,
          'success',
          4200
        );
      }
      return { ...res, name: handle.name, sheets: payload.wb.worksheets.length, count: payload.logs.length };
    } catch (err) {
      if (interactive) showToast(`บันทึกไฟล์ไม่สำเร็จ: ${err.message}`, 'error', 5000);
      return { status: 'error', error: err };
    } finally {
      if (btn && interactive) btn.disabled = false;
    }
  }

  async function autoSyncBoundFiles() {
    const perSiteZoneKeys = Object.keys(state.boundZoneHandlesBySite || {});
    const perSiteBillingKeys = Object.keys(state.boundBillingHandlesBySite || {});
    if (
      state.autoSync &&
      supportsFileBinding() &&
      (state.boundZoneHandle || state.boundBillingHandle || perSiteZoneKeys.length > 0 || perSiteBillingKeys.length > 0)
    ) {
      const syncedNames = [];
      let needsPerm = false;

      if (state.boundZoneHandle) {
        const r = await syncToBoundFile('zone', { interactive: false });
        if (r.status === 'ok') syncedNames.push(r.name);
        else if (r.status === 'need_permission') needsPerm = true;
      }
      if (state.boundBillingHandle) {
        const r = await syncToBoundFile('billing', { interactive: false });
        if (r.status === 'ok') syncedNames.push(r.name);
        else if (r.status === 'need_permission') needsPerm = true;
      }
      for (const sName of perSiteZoneKeys) {
        const r = await syncSiteExcelFile('zone', sName, { interactive: false });
        if (r.status === 'ok') syncedNames.push(`สำเนา:${sName}`);
        else if (r.status === 'need_permission') needsPerm = true;
      }
      for (const sName of perSiteBillingKeys) {
        const r = await syncSiteExcelFile('billing', sName, { interactive: false });
        if (r.status === 'ok') syncedNames.push(`ใบเบิก:${sName}`);
        else if (r.status === 'need_permission') needsPerm = true;
      }

      if (syncedNames.length > 0) {
        showToast(`ซิงค์ลงไฟล์ที่ผูกไว้อัตโนมัติแล้ว: ${syncedNames.join(', ')}`, 'info', 3500);
      } else if (needsPerm) {
        showToast('กดปุ่ม "เขียนสำเนา" หรือ "เขียนใบเบิก" 1 ครั้งเพื่ออนุญาตสิทธิ์หลังรีเฟรชหน้า', 'warning', 4500);
      }
    }

    await autoSyncGoogleSheets();
  }

  // ==========================================================================
  // 9B. Google Sheets Online Sync (ไหลลง Google Sheets — ทั้งสำเนาโซน & ใบเบิก)
  // ==========================================================================
  const LATEST_APPS_SCRIPT_VERSION = '2.1-logo';

  async function getAppsScriptCode() {
    const res = await fetch('/google-apps-script.js?v=3', { cache: 'no-store' });
    if (!res.ok) throw new Error('ไม่สามารถโหลดไฟล์โค้ด Google Apps Script ได้');
    return res.text();
  }

  async function copyAppsScriptCodeToClipboard() {
    try {
      const code = await getAppsScriptCode();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(code);
      } else {
        const ta = document.createElement('textarea');
        ta.value = code;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        ta.remove();
      }
      showToast('คัดลอกโค้ด Google Apps Script เรียบร้อยแล้ว — นำไปวางใน Extensions → Apps Script ได้เลย', 'success', 4500);
    } catch (err) {
      showToast(err.message || 'คัดลอกโค้ดไม่สำเร็จ', 'error');
    }
  }

  function formatDateTimeBE(isoStr) {
    if (!isoStr) return '-';
    const d = new Date(isoStr);
    if (Number.isNaN(d.getTime())) return String(isoStr);
    const day = d.getDate();
    const month = d.getMonth() + 1;
    const yearBE = d.getFullYear() + 543;
    const hh = String(d.getHours()).padStart(2, '0');
    const mm = String(d.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${yearBE} ${hh}:${mm} น.`;
  }

  function applyGoogleSheetHistoryMeta(historyArr) {
    if (!Array.isArray(historyArr)) return;
    state.gsheetHistory = historyArr;
    state.gsheetHistoryCount = historyArr.length;
    state.gsheetLastHistory = historyArr[0] || null;
  }

  async function loadGoogleSheetsConfig() {
    try {
      const cfg = await apiFetch('/api/gsheet/config');
      if (cfg && typeof cfg === 'object') {
        state.gsheet = { ...state.gsheet, ...cfg };
        if (typeof cfg.history_count === 'number') state.gsheetHistoryCount = cfg.history_count;
        if (cfg.last_history !== undefined) state.gsheetLastHistory = cfg.last_history || null;
      }
    } catch {
      // ignore if offline
    }
    renderGoogleSheetsUI();
  }

  async function loadGoogleSheetsHistory() {
    const res = await apiFetch('/api/gsheet/history');
    const rows = Array.isArray(res && res.history) ? res.history : [];
    applyGoogleSheetHistoryMeta(rows);
    renderGoogleSheetsUI();
    return rows;
  }

  function renderPerSiteGoogleSheetsList() {
    const listEl = $('gsheet-per-site-list');
    const zoneCountBadge = $('gsheet-zone-sites-bound-count');
    const billingCountBadge = $('gsheet-sites-bound-count');
    const toggleBtn = $('btn-gsheet-sites-toggle');
    const searchEl = $('gsheet-sites-search');
    const syncAllZoneBtn = $('btn-sync-all-gsheet-zone-sites');
    const syncAllBillingBtn = $('btn-sync-all-gsheet-sites');
    if (!listEl) return;

    const cfg = state.gsheet || {};
    const zoneBySite = cfg.zone_by_site || {};
    const billingBySite = cfg.billing_by_site || {};
    const zoneBoundCount = Object.keys(zoneBySite).filter((k) => zoneBySite[k] && zoneBySite[k].url).length;
    const billingBoundCount = Object.keys(billingBySite).filter((k) => billingBySite[k] && billingBySite[k].url).length;
    if (zoneCountBadge) zoneCountBadge.textContent = `ผูกแล้ว ${formatInt(zoneBoundCount)} โครงการ`;
    if (billingCountBadge) billingCountBadge.textContent = `ผูกแล้ว ${formatInt(billingBoundCount)} โครงการ`;
    if (syncAllZoneBtn) {
      syncAllZoneBtn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i><span>ไหลสำเนาทุกโครงการที่ผูก (${formatInt(zoneBoundCount)})</span>`;
    }
    if (syncAllBillingBtn) {
      syncAllBillingBtn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i><span>ไหลใบเบิกทุกโครงการที่ผูก (${formatInt(billingBoundCount)})</span>`;
    }

    const logCounts = getLogCountBySite();
    const allProjects = getKnownProjectNames();
    const q = searchEl ? searchEl.value.trim().toLowerCase() : '';

    const sorted = [...allProjects].sort((a, b) => {
      const ab = (zoneBySite[a] && zoneBySite[a].url) || (billingBySite[a] && billingBySite[a].url) ? 1 : 0;
      const bb = (zoneBySite[b] && zoneBySite[b].url) || (billingBySite[b] && billingBySite[b].url) ? 1 : 0;
      if (ab !== bb) return bb - ab;
      const ac = logCounts[a] || 0;
      const bc = logCounts[b] || 0;
      if ((ac > 0) !== (bc > 0)) return bc > 0 ? 1 : -1;
      return a.localeCompare(b, 'th');
    });

    const important = sorted.filter(
      (s) =>
        Boolean((zoneBySite[s] && zoneBySite[s].url) || (billingBySite[s] && billingBySite[s].url)) ||
        (logCounts[s] || 0) > 0
    );
    const defaultList = state.gsheetSitesShowAll || q || important.length === 0 ? sorted : important;
    const filtered = q ? defaultList.filter((s) => s.toLowerCase().includes(q)) : defaultList;

    if (toggleBtn) {
      if (important.length > 0 && important.length < sorted.length && !q) {
        toggleBtn.classList.remove('hidden');
        toggleBtn.textContent = state.gsheetSitesShowAll
          ? `ย่อเฉพาะที่ผูก/มีเที่ยว (${important.length})`
          : `แสดงทุกโครงการ (${sorted.length})`;
      } else {
        toggleBtn.classList.add('hidden');
      }
    }

    if (filtered.length === 0) {
      listEl.innerHTML = allProjects.length === 0
        ? '<div class="col-span-full py-3 text-center text-xs text-slate-400">ยังไม่มีโครงการ — เพิ่มโครงการในหน้า "ตั้งค่า" ก่อน เพื่อผูก Google Sheet สำเนาและใบเบิกแยกแต่ละโครงการ</div>'
        : '<div class="col-span-full py-3 text-center text-xs text-slate-400">ไม่พบโครงการที่ค้นหา</div>';
      return;
    }

    listEl.innerHTML = filtered
      .map((siteName) => {
        const zEntry = zoneBySite[siteName] && zoneBySite[siteName].url ? zoneBySite[siteName] : null;
        const bEntry = billingBySite[siteName] && billingBySite[siteName].url ? billingBySite[siteName] : null;
        const trips = logCounts[siteName] || 0;
        const hasAnyBound = Boolean(zEntry || bEntry);
        return `
          <div class="p-2.5 rounded-xl bg-white border ${hasAnyBound ? 'border-emerald-300 shadow-xs' : 'border-slate-200'} flex flex-col gap-2">
            <div class="flex items-center justify-between gap-2 min-w-0">
              <div class="flex items-center gap-1.5 min-w-0">
                <i class="fa-solid fa-building text-amber-500 text-[11px] shrink-0"></i>
                <span class="font-bold text-xs text-slate-900 truncate" title="${escapeHtml(siteName)}">${escapeHtml(siteName)}</span>
              </div>
              ${trips > 0 ? `<span class="px-1.5 py-0.5 rounded-full bg-indigo-50 text-indigo-700 text-[10px] font-bold shrink-0">${formatInt(trips)} เที่ยว</span>` : ''}
            </div>

            <!-- ชีตสำเนา (เฉพาะโครงการนี้) -->
            <div class="p-2 rounded-lg bg-emerald-50/70 border border-emerald-200/80 space-y-1.5">
              <div class="flex items-center justify-between gap-1.5 min-w-0">
                <div class="flex items-center gap-1.5 min-w-0">
                  <span class="px-1.5 py-0.5 rounded bg-emerald-200/80 text-emerald-950 text-[10px] font-bold shrink-0">สำเนา</span>
                  ${
                    zEntry
                      ? `<span class="w-2 h-2 rounded-full bg-emerald-500 shrink-0"></span><span class="text-[11px] text-emerald-900 font-semibold truncate" title="${escapeHtml(zEntry.name || zEntry.url)}">${escapeHtml(zEntry.name || 'เชื่อมต่อชีตสำเนาแล้ว')}</span>`
                      : '<span class="text-[11px] text-slate-400 truncate">ยังไม่ผูกชีตสำเนา</span>'
                  }
                </div>
                ${
                  zEntry && zEntry.sheet_url
                    ? `<a href="${escapeHtml(zEntry.sheet_url)}" target="_blank" rel="noopener noreferrer" title="เปิด Google Sheet สำเนาโครงการ ${escapeHtml(siteName)}"
                        class="w-6 h-6 rounded-md bg-white hover:bg-emerald-100 text-emerald-700 border border-emerald-200 inline-flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-arrow-up-right-from-square text-[10px]"></i>
                      </a>`
                    : ''
                }
              </div>
              <div class="flex flex-wrap items-center gap-1.5">
                <button type="button" data-gsheet-zone-site-bind="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg bg-white hover:bg-emerald-100 text-emerald-900 border border-emerald-300 text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition">
                  <i class="fa-brands fa-google text-[10px]"></i><span>${zEntry ? 'ตั้งค่าชีต' : 'ผูกชีตสำเนา'}</span>
                </button>
                <button type="button" data-gsheet-zone-site-sync="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg ${zEntry ? 'bg-emerald-600 hover:bg-emerald-700 text-white' : 'bg-white/90 hover:bg-emerald-600 hover:text-white text-emerald-900 border border-emerald-200'} text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition"
                  title="ไหลเฉพาะข้อมูลสำเนาของโครงการ ${escapeHtml(siteName)} ลง Google Sheet">
                  <i class="fa-solid fa-cloud-arrow-up text-[10px]"></i><span>ไหลลงสำเนา</span>
                </button>
              </div>
            </div>

            <!-- ชีตใบเบิก (เฉพาะโครงการนี้) -->
            <div class="p-2 rounded-lg bg-teal-50/70 border border-teal-200/80 space-y-1.5">
              <div class="flex items-center justify-between gap-1.5 min-w-0">
                <div class="flex items-center gap-1.5 min-w-0">
                  <span class="px-1.5 py-0.5 rounded bg-teal-200/80 text-teal-950 text-[10px] font-bold shrink-0">ใบเบิก</span>
                  ${
                    bEntry
                      ? `<span class="w-2 h-2 rounded-full bg-emerald-500 shrink-0"></span><span class="text-[11px] text-teal-900 font-semibold truncate" title="${escapeHtml(bEntry.name || bEntry.url)}">${escapeHtml(bEntry.name || 'เชื่อมต่อชีตใบเบิกแล้ว')}</span>`
                      : '<span class="text-[11px] text-slate-400 truncate">ยังไม่ผูกชีตใบเบิก</span>'
                  }
                </div>
                ${
                  bEntry && bEntry.sheet_url
                    ? `<a href="${escapeHtml(bEntry.sheet_url)}" target="_blank" rel="noopener noreferrer" title="เปิด Google Sheet ใบเบิกโครงการ ${escapeHtml(siteName)}"
                        class="w-6 h-6 rounded-md bg-white hover:bg-teal-100 text-teal-700 border border-teal-200 inline-flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-arrow-up-right-from-square text-[10px]"></i>
                      </a>`
                    : ''
                }
              </div>
              <div class="flex flex-wrap items-center gap-1.5">
                <button type="button" data-gsheet-site-bind="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg bg-white hover:bg-teal-100 text-teal-900 border border-teal-300 text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition">
                  <i class="fa-brands fa-google text-[10px]"></i><span>${bEntry ? 'ตั้งค่าชีต' : 'ผูกชีตใบเบิก'}</span>
                </button>
                <button type="button" data-gsheet-site-sync="${escapeHtml(siteName)}"
                  class="flex-1 min-h-[28px] px-2 py-1 rounded-lg ${bEntry ? 'bg-teal-700 hover:bg-teal-800 text-white' : 'bg-white/90 hover:bg-teal-700 hover:text-white text-teal-900 border border-teal-200'} text-[11px] font-semibold flex items-center justify-center gap-1 cursor-pointer transition"
                  title="ไหลเฉพาะข้อมูลใบเบิกของโครงการ ${escapeHtml(siteName)} ลง Google Sheet">
                  <i class="fa-solid fa-cloud-arrow-up text-[10px]"></i><span>ไหลลงใบเบิก</span>
                </button>
              </div>
            </div>
          </div>
        `;
      })
      .join('');
  }

  function renderGoogleSheetsUI() {
    const cfg = state.gsheet || {};
    const zLabel = $('gsheet-zone-label');
    const zDot = $('gsheet-zone-dot');
    const zOpen = $('btn-gsheet-open-zone');
    const legacyGsZoneWrap = $('legacy-gsheet-zone-wrap');
    const bLabel = $('gsheet-billing-label');
    const bDot = $('gsheet-billing-dot');
    const bOpen = $('btn-gsheet-open-billing');
    const legacyGsWrap = $('legacy-gsheet-billing-wrap');
    const autoToggle = $('toggle-gsheet-auto-sync');
    const histBadge = $('gsheet-history-badge');
    const lastNote = $('gsheet-last-sync-note');

    if (autoToggle) autoToggle.checked = cfg.auto_sync !== false;
    if (histBadge) histBadge.textContent = formatInt(state.gsheetHistoryCount || 0);
    if (legacyGsZoneWrap) legacyGsZoneWrap.classList.toggle('hidden', !cfg.zone_url);
    if (legacyGsWrap) legacyGsWrap.classList.toggle('hidden', !cfg.billing_url);

    if (lastNote) {
      const h = state.gsheetLastHistory;
      if (h) {
        const dt = formatDateTimeBE(h.synced_at);
        const kindLabel =
          h.kind === 'billing'
            ? h.site_name
              ? `ใบเบิก [${h.site_name}]`
              : 'ใบเบิก'
            : h.kind === 'zone'
              ? h.site_name
                ? `สำเนา [${h.site_name}]`
                : 'สำเนาโซน'
              : 'ตั้งค่า/บันทึก';
        const targetText = h.spreadsheet_name ? ` (${escapeHtml(h.spreadsheet_name)})` : '';
        if (h.status === 'error') {
          lastNote.innerHTML = `<i class="fa-solid fa-circle-exclamation text-rose-600 mr-1"></i>บันทึกล่าสุด ${escapeHtml(dt)} • ${escapeHtml(kindLabel)} ไม่สำเร็จ`;
        } else if (h.mode === 'config_saved') {
          lastNote.innerHTML = `<i class="fa-solid fa-floppy-disk text-emerald-600 mr-1"></i>บันทึกการตั้งค่าล่าสุด ${escapeHtml(dt)}${targetText}`;
        } else if (h.mode === 'note') {
          lastNote.innerHTML = `<i class="fa-regular fa-note-sticky text-emerald-600 mr-1"></i>บันทึกล่าสุด ${escapeHtml(dt)} • ${escapeHtml(h.note || kindLabel)}`;
        } else {
          lastNote.innerHTML = `<i class="fa-solid fa-circle-check text-emerald-600 mr-1"></i>ไหลล่าสุด ${escapeHtml(dt)} • ${escapeHtml(kindLabel)}${targetText} (${formatInt(h.log_count)} เที่ยว)`;
        }
      } else {
        lastNote.textContent = 'ยังไม่มีบันทึกประวัติการไหลข้อมูล';
      }
    }

    if (zLabel) {
      if (cfg.zone_url) {
        zLabel.textContent = cfg.zone_name || 'เชื่อมต่อ Google Sheet สำเนาโซนแล้ว';
        zLabel.className = 'truncate text-emerald-900 font-bold';
        if (zDot) zDot.classList.remove('hidden');
        if (zOpen) {
          if (cfg.zone_sheet_url) {
            zOpen.href = cfg.zone_sheet_url;
            zOpen.classList.remove('hidden');
            zOpen.classList.add('inline-flex');
          } else {
            zOpen.classList.add('hidden');
            zOpen.classList.remove('inline-flex');
          }
        }
      } else {
        zLabel.textContent = 'ยังไม่ได้ผูก (กดเพื่อตั้งค่า)';
        zLabel.className = 'truncate text-emerald-700';
        if (zDot) zDot.classList.add('hidden');
        if (zOpen) {
          zOpen.classList.add('hidden');
          zOpen.classList.remove('inline-flex');
        }
      }
    }

    if (bLabel) {
      if (cfg.billing_url) {
        const siteTag = cfg.billing_site ? ` [${cfg.billing_site}]` : '';
        bLabel.textContent = `${cfg.billing_name || 'เชื่อมต่อ Google Sheet ใบเบิกแล้ว'}${siteTag}`;
        bLabel.className = 'truncate text-teal-950 font-bold';
        if (bDot) bDot.classList.remove('hidden');
        if (bOpen) {
          if (cfg.billing_sheet_url) {
            bOpen.href = cfg.billing_sheet_url;
            bOpen.classList.remove('hidden');
            bOpen.classList.add('inline-flex');
          } else {
            bOpen.classList.add('hidden');
            bOpen.classList.remove('inline-flex');
          }
        }
      } else {
        bLabel.textContent = 'ยังไม่ได้ผูก (กดเพื่อตั้งค่า)';
        bLabel.className = 'truncate text-teal-700';
        if (bDot) bDot.classList.add('hidden');
        if (bOpen) {
          bOpen.classList.add('hidden');
          bOpen.classList.remove('inline-flex');
        }
      }
    }

    renderPerSiteGoogleSheetsList();
  }

  async function saveGoogleSheetsBarSettings() {
    const btn = $('btn-gsheet-save-settings');
    const autoToggle = $('toggle-gsheet-auto-sync');
    const autoVal = autoToggle ? Boolean(autoToggle.checked) : true;
    const cfg = state.gsheet || {};
    if (btn) btn.disabled = true;
    try {
      const perSiteZoneCount = Object.keys(cfg.zone_by_site || {}).length;
      const perSiteCount = Object.keys(cfg.billing_by_site || {}).length;
      const zoneSummary =
        perSiteZoneCount > 0
          ? `สำเนาแยกโครงการ: ${perSiteZoneCount} โครงการ` + (cfg.zone_url ? ` + ชีตรวมเดิม (${cfg.zone_name || 'ผูกแล้ว'})` : '')
          : cfg.zone_url
            ? `สำเนา: ${cfg.zone_name || 'ผูกแล้ว'}`
            : 'สำเนาแยกโครงการ: ยังไม่ผูก';
      const billingSummary =
        perSiteCount > 0
          ? `ใบเบิกแยกโครงการ: ${perSiteCount} โครงการ` + (cfg.billing_url ? ` + ชีตรวมเดิม (${cfg.billing_name || 'ผูกแล้ว'})` : '')
          : cfg.billing_url
            ? `ใบเบิก: ${cfg.billing_name || 'ผูกแล้ว'}`
            : 'ใบเบิกแยกโครงการ: ยังไม่ผูก';
      const boundSummary = [
        zoneSummary,
        billingSummary,
        `อัปเดตอัตโนมัติ: ${autoVal ? 'เปิด' : 'ปิด'}`
      ].join(' • ');
      const res = await apiFetch('/api/gsheet/config', {
        method: 'PUT',
        body: JSON.stringify({
          auto_sync: autoVal,
          record_history: true,
          history_kind: 'general',
          history_note: `บันทึกการตั้งค่า (${boundSummary})`
        })
      });
      if (res.config) state.gsheet = { ...state.gsheet, ...res.config };
      if (Array.isArray(res.history)) applyGoogleSheetHistoryMeta(res.history);
      renderGoogleSheetsUI();
      showToast('บันทึกการตั้งค่า Google Sheets และลงบันทึกประวัติเรียบร้อยแล้ว', 'success', 3800);
    } catch (err) {
      showToast(err.message || 'ไม่สามารถบันทึกการตั้งค่าได้', 'error');
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  function openGoogleSheetBindModal(kind = 'zone', fixedSiteName = '') {
    const isBilling = kind === 'billing';
    const fixedSite = String(fixedSiteName || '').trim();
    const kindThai = isBilling ? 'ใบเบิก' : 'สำเนา';
    const cfg = state.gsheet || {};
    const siteMap = isBilling ? cfg.billing_by_site : cfg.zone_by_site;
    const perSiteEntry = fixedSite && siteMap ? siteMap[fixedSite] : null;
    const title = fixedSite
      ? `ผูก Google Sheet ${kindThai}เฉพาะโครงการ: ${fixedSite}`
      : isBilling
        ? 'ผูก Google Sheets ใบเบิกโครงการ'
        : 'ผูก Google Sheets สำเนาโซน';
    const curUrl = fixedSite
      ? ((perSiteEntry && perSiteEntry.url) || '')
      : isBilling
        ? (cfg.billing_url || '')
        : (cfg.zone_url || '');
    const curName = fixedSite
      ? ((perSiteEntry && perSiteEntry.name) || '')
      : isBilling
        ? (cfg.billing_name || '')
        : (cfg.zone_name || '');
    const curSheetUrl = fixedSite
      ? ((perSiteEntry && perSiteEntry.sheet_url) || '')
      : isBilling
        ? (cfg.billing_sheet_url || '')
        : (cfg.zone_sheet_url || '');
    const curSite = fixedSite || cfg.billing_site || $('filter-site').value.trim() || '';
    const siteOptions = getKnownProjectNames();

    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 rounded-xl ${isBilling ? 'bg-teal-100 text-teal-700' : 'bg-emerald-100 text-emerald-700'} flex items-center justify-center shrink-0">
            <i class="fa-brands fa-google text-lg"></i>
          </div>
          <div>
            <h3 class="text-base font-bold text-slate-900">${escapeHtml(title)}</h3>
            <p class="text-xs text-slate-500">${fixedSite ? `ไหลเฉพาะข้อมูล${kindThai}ของโครงการ "${escapeHtml(fixedSite)}" เท่านั้น ไม่รวมกับโครงการอื่น` : 'ตั้งค่าครั้งเดียว ใช้ร่วมกันได้ทุกเครื่องทั้งคอมพิวเตอร์และมือถือ'}</p>
          </div>
        </div>

        <div class="p-3.5 rounded-xl bg-emerald-50/80 border border-emerald-200 text-xs text-emerald-950 space-y-2">
          <div class="flex flex-wrap items-center justify-between gap-2">
            <span class="font-bold"><i class="fa-solid fa-list-ol mr-1"></i>วิธีเชื่อม Google Sheets (ทำครั้งเดียว 2 นาที)</span>
            <button type="button" id="m-gs-copy-code"
              class="min-h-[32px] px-3 py-1 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-semibold flex items-center gap-1.5 cursor-pointer">
              <i class="fa-solid fa-copy"></i><span>คัดลอกโค้ด Apps Script</span>
            </button>
          </div>
          <ol class="list-decimal pl-5 space-y-1 text-[11px] text-emerald-900">
            <li>เปิดไฟล์ Google Sheets ที่ต้องการ → เมนู <strong>ส่วนขยาย (Extensions)</strong> → <strong>Apps Script</strong></li>
            <li>ลบโค้ดเดิมออกทั้งหมด แล้วกด <strong>วาง (Ctrl+V)</strong> โค้ดที่คัดลอกไว้ แล้วกดบันทึก (Ctrl+S)</li>
            <li>กดปุ่ม <strong>ทำให้ใช้งานได้ (Deploy)</strong> → <strong>การทำให้ใช้งานได้รายการใหม่</strong> → เลือกประเภท <strong>เว็บแอป (Web app)</strong></li>
            <li>ตั้งค่า <em>ดำเนินการในฐานะ: ฉัน</em> และ <em>ผู้ที่มีสิทธิ์เข้าถึง: <strong>ทุกคน (Anyone)</strong></em> → กดทำให้ใช้งานได้ แล้วคัดลอก <strong>Web app URL</strong> มาวางด้านล่าง</li>
          </ol>
        </div>

        <form id="m-gs-form" class="space-y-3.5">
          <div>
            <label for="m-gs-url" class="block text-xs font-semibold text-slate-700 mb-1">
              ลิงก์ Web app URL จาก Google Apps Script *
            </label>
            <input type="url" id="m-gs-url" value="${escapeHtml(curUrl)}"
              placeholder="https://script.google.com/macros/s/....../exec"
              class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm focus:outline-none focus:ring-2 focus:ring-emerald-600" />
            <p id="m-gs-test-result" class="mt-1.5 text-xs ${curName ? 'text-emerald-700 font-semibold' : 'text-slate-500'}">
              ${curName ? `เชื่อมต่ออยู่กับ: ${escapeHtml(curName)}` : 'วางลิงก์แล้วกด "ทดสอบเชื่อมต่อ" เพื่อตรวจชื่อชีต หรือกด "บันทึก" ไว้ก่อนได้'}
            </p>
          </div>

          ${
            fixedSite
              ? `
            <div class="p-3 rounded-xl ${isBilling ? 'bg-teal-50 border-teal-200 text-teal-950' : 'bg-emerald-50 border-emerald-200 text-emerald-950'} border text-xs font-semibold flex items-center gap-2">
              <i class="fa-solid fa-building text-amber-500"></i>
              <span>โครงการที่ผูก (${escapeHtml(kindThai)}): <strong>${escapeHtml(fixedSite)}</strong> (แยกชีตเฉพาะโครงการนี้ ไม่รวมโครงการอื่น)</span>
            </div>
          `
              : isBilling
                ? `
            <div>
              <label for="m-gs-billing-site" class="block text-xs font-semibold text-slate-700 mb-1">
                ผูกกับโครงการประจำชีต (แนะนำ — ป้องกันไหลทับข้ามโครงการ)
              </label>
              <select id="m-gs-billing-site" class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white">
                <option value="">-- ตามตัวกรองโครงการหน้าเว็บ --</option>
                ${siteOptions
                  .map((s) => `<option value="${escapeHtml(s)}" ${s === curSite ? 'selected' : ''}>${escapeHtml(s)}</option>`)
                  .join('')}
              </select>
            </div>
          `
                : ''
          }

          <div class="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <div class="flex items-center gap-2">
              <button type="button" id="m-gs-test"
                class="min-h-[40px] px-3 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 text-xs font-semibold cursor-pointer">
                <i class="fa-solid fa-plug-circle-check mr-1 text-emerald-600"></i>ทดสอบเชื่อมต่อ
              </button>
              ${
                curUrl
                  ? `<button type="button" id="m-gs-unbind"
                      class="min-h-[40px] px-3 py-2 rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-semibold cursor-pointer">
                      ยกเลิกผูก
                    </button>`
                  : ''
              }
            </div>
            <div class="flex flex-wrap items-center gap-2">
              <button type="button" id="m-gs-cancel"
                class="min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm cursor-pointer">ยกเลิก</button>
              <button type="button" id="m-gs-save-only"
                class="min-h-[40px] px-3.5 py-2 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-white text-xs sm:text-sm font-semibold cursor-pointer"
                title="บันทึกการตั้งค่าไว้ก่อน โดยยังไม่ไหลข้อมูลทันที">
                <i class="fa-solid fa-floppy-disk mr-1 text-amber-400"></i>บันทึก
              </button>
              <button type="submit" id="m-gs-save"
                class="min-h-[40px] px-3.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-semibold cursor-pointer">
                <i class="fa-solid fa-cloud-arrow-up mr-1"></i>บันทึกและไหลเลย
              </button>
            </div>
          </div>
        </form>
      </div>
      `,
      (box) => {
        let testedMeta = { name: curName, sheetUrl: curSheetUrl };
        const urlEl = box.querySelector('#m-gs-url');
        const resEl = box.querySelector('#m-gs-test-result');
        const saveOnlyBtn = box.querySelector('#m-gs-save-only');
        const saveAndSyncBtn = box.querySelector('#m-gs-save');

        box.querySelector('#m-gs-copy-code').addEventListener('click', copyAppsScriptCodeToClipboard);
        box.querySelector('#m-gs-cancel').addEventListener('click', closeModal);

        const runPing = async () => {
          const url = urlEl.value.trim();
          if (!url) {
            resEl.className = 'mt-1.5 text-xs text-rose-600 font-semibold';
            resEl.textContent = 'กรุณาวาง Web app URL ก่อน';
            return null;
          }
          resEl.className = 'mt-1.5 text-xs text-amber-700 font-semibold';
          resEl.textContent = 'กำลังทดสอบเชื่อมต่อ Google Sheets...';
          try {
            const r = await apiFetch('/api/gsheet/test', {
              method: 'POST',
              body: JSON.stringify({ url })
            });
            testedMeta = {
              name: r.spreadsheetName || 'Google Sheets',
              sheetUrl: r.spreadsheetUrl || '',
              scriptVersion: r.scriptVersion || '1.0'
            };
            if (isBilling && testedMeta.scriptVersion !== LATEST_APPS_SCRIPT_VERSION) {
              resEl.className = 'mt-1.5 text-xs text-amber-700 font-bold';
              resEl.textContent = `เชื่อมต่อพบไฟล์ "${testedMeta.name}" แต่โค้ด Apps Script ยังเป็นเวอร์ชันเก่า (ยังไม่มีโลโก้) — กดคัดลอกโค้ดใหม่ไปวางแล้วเลือก "จัดการการทำให้ใช้งานได้ → แก้ไข → เวอร์ชันใหม่"`;
            } else {
              resEl.className = 'mt-1.5 text-xs text-emerald-700 font-bold';
              resEl.textContent = `เชื่อมต่อสำเร็จ! พบไฟล์: "${testedMeta.name}"${testedMeta.scriptVersion === LATEST_APPS_SCRIPT_VERSION ? ' (โค้ดเวอร์ชันล่าสุด พร้อมโลโก้ CKP)' : ''}`;
            }
            return testedMeta;
          } catch (err) {
            resEl.className = 'mt-1.5 text-xs text-rose-600 font-semibold';
            resEl.textContent = err.message || 'เชื่อมต่อไม่สำเร็จ';
            return null;
          }
        };

        const buildPatch = (url, meta, historyNote) =>
          fixedSite
            ? {
                [isBilling ? 'billing_site_entry' : 'zone_site_entry']: {
                  site_name: fixedSite,
                  url,
                  name: (meta && meta.name) || '',
                  sheet_url: (meta && meta.sheetUrl) || ''
                },
                record_history: true,
                history_kind: kind,
                history_note: historyNote
              }
            : isBilling
              ? {
                  billing_url: url,
                  billing_name: (meta && meta.name) || '',
                  billing_sheet_url: (meta && meta.sheetUrl) || '',
                  billing_site: box.querySelector('#m-gs-billing-site') ? box.querySelector('#m-gs-billing-site').value.trim() : '',
                  record_history: true,
                  history_kind: 'billing',
                  history_note: historyNote
                }
              : {
                  zone_url: url,
                  zone_name: (meta && meta.name) || '',
                  zone_sheet_url: (meta && meta.sheetUrl) || '',
                  record_history: true,
                  history_kind: 'zone',
                  history_note: historyNote
                };

        box.querySelector('#m-gs-test').addEventListener('click', runPing);

        const unbindBtn = box.querySelector('#m-gs-unbind');
        if (unbindBtn) {
          unbindBtn.addEventListener('click', async () => {
            try {
              const patch = fixedSite
                ? {
                    [isBilling ? 'billing_site_entry' : 'zone_site_entry']: {
                      site_name: fixedSite,
                      url: '',
                      name: '',
                      sheet_url: ''
                    },
                    record_history: true,
                    history_kind: kind,
                    history_note: `ยกเลิกการผูก Google Sheet ${kindThai}โครงการ ${fixedSite}`
                  }
                : isBilling
                  ? {
                      billing_url: '',
                      billing_name: '',
                      billing_sheet_url: '',
                      billing_site: '',
                      record_history: true,
                      history_kind: 'billing',
                      history_note: 'ยกเลิกการผูก Google Sheet ใบเบิกโครงการ'
                    }
                  : {
                      zone_url: '',
                      zone_name: '',
                      zone_sheet_url: '',
                      record_history: true,
                      history_kind: 'zone',
                      history_note: 'ยกเลิกการผูก Google Sheet สำเนาโซน'
                    };
              const saved = await apiFetch('/api/gsheet/config', { method: 'PUT', body: JSON.stringify(patch) });
              state.gsheet = saved.config;
              if (Array.isArray(saved.history)) applyGoogleSheetHistoryMeta(saved.history);
              renderGoogleSheetsUI();
              closeModal();
              showToast(`ยกเลิกการผูก ${title} เรียบร้อยแล้ว`, 'info');
            } catch (err) {
              showToast(err.message, 'error');
            }
          });
        }

        // ปุ่ม "บันทึก" อย่างเดียว (เก็บการตั้งค่าไว้ก่อน ยังไม่ไหลข้อมูลทันที)
        saveOnlyBtn.addEventListener('click', async () => {
          const url = urlEl.value.trim();
          if (!url) {
            showToast('กรุณาวางลิงก์ Web app URL ของ Google Apps Script', 'error');
            return;
          }
          saveOnlyBtn.disabled = true;
          saveAndSyncBtn.disabled = true;
          try {
            let meta = url === curUrl ? testedMeta : { name: '', sheetUrl: '' };
            if (!meta.name) {
              const ping = await runPing();
              if (ping) meta = ping;
            }
            const patch = buildPatch(url, meta, `บันทึกการตั้งค่า ${title} (ยังไม่ไหลข้อมูล)`);
            const saved = await apiFetch('/api/gsheet/config', { method: 'PUT', body: JSON.stringify(patch) });
            state.gsheet = saved.config;
            if (Array.isArray(saved.history)) applyGoogleSheetHistoryMeta(saved.history);
            renderGoogleSheetsUI();
            closeModal();
            showToast(
              `บันทึกการตั้งค่า ${title} เรียบร้อยแล้ว (กดปุ่ม "ไหลลงชีตนี้" เมื่อพร้อม)`,
              'success',
              4200
            );
          } catch (err) {
            showToast(err.message || 'ไม่สามารถบันทึกการตั้งค่าได้', 'error');
            saveOnlyBtn.disabled = false;
            saveAndSyncBtn.disabled = false;
          }
        });

        // ปุ่ม "บันทึกและไหลเลย"
        box.querySelector('#m-gs-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          const url = urlEl.value.trim();
          if (!url) {
            showToast('กรุณาวางลิงก์ Web app URL ของ Google Apps Script', 'error');
            return;
          }
          saveOnlyBtn.disabled = true;
          saveAndSyncBtn.disabled = true;
          try {
            const ping = await runPing();
            if (!ping) {
              saveOnlyBtn.disabled = false;
              saveAndSyncBtn.disabled = false;
              return;
            }
            const patch = buildPatch(url, ping, `บันทึกการตั้งค่าและเชื่อมต่อ ${title}`);
            const saved = await apiFetch('/api/gsheet/config', { method: 'PUT', body: JSON.stringify(patch) });
            state.gsheet = saved.config;
            if (Array.isArray(saved.history)) applyGoogleSheetHistoryMeta(saved.history);
            renderGoogleSheetsUI();
            closeModal();
            await syncToGoogleSheet(kind, { interactive: true, mode: 'manual', siteName: fixedSite });
          } catch (err) {
            showToast(err.message || 'ไม่สามารถบันทึกการตั้งค่าได้', 'error');
            saveOnlyBtn.disabled = false;
            saveAndSyncBtn.disabled = false;
          }
        });
      }
    );
  }

  async function openGoogleSheetsHistoryModal() {
    let rows = [];
    try {
      rows = await loadGoogleSheetsHistory();
    } catch (err) {
      showToast(err.message || 'ไม่สามารถโหลดประวัติการไหลข้อมูลได้', 'error');
      return;
    }

    const isAdmin = Boolean(state.user && state.user.role === 'admin');
    const f = getFilterParams();

    const renderListHtml = (list) => {
      if (!list || list.length === 0) {
        return `
          <div class="py-10 text-center text-slate-400">
            <i class="fa-regular fa-clipboard text-2xl mb-2"></i>
            <p class="text-sm font-semibold text-slate-600">ยังไม่มีบันทึกประวัติการไหลข้อมูล</p>
            <p class="text-xs mt-1">เมื่อกดไหลลง Google Sheets หรือบันทึกการตั้งค่า/บันทึกหมายเหตุด้านบน ระบบจะเก็บประวัติไว้ที่นี่อัตโนมัติ</p>
          </div>
        `;
      }
      return list
        .map((item) => {
          const kindBadge =
            item.kind === 'billing'
              ? '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-teal-100 text-teal-800 border border-teal-200">ใบเบิก</span>'
              : item.kind === 'zone'
                ? '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">สำเนาโซน</span>'
                : '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-700 border border-slate-200">ทั่วไป</span>';

          const modeMap = {
            manual: { label: 'กดไหลเอง', cls: 'bg-indigo-50 text-indigo-700 border-indigo-200' },
            auto: { label: 'ไหลอัตโนมัติ', cls: 'bg-sky-50 text-sky-700 border-sky-200' },
            config_saved: { label: 'บันทึกตั้งค่า', cls: 'bg-amber-50 text-amber-800 border-amber-200' },
            note: { label: 'บันทึกเอง', cls: 'bg-purple-50 text-purple-700 border-purple-200' }
          };
          const m = modeMap[item.mode] || modeMap.manual;
          const modeBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-semibold border ${m.cls}">${m.label}</span>`;

          const statusBadge =
            item.status === 'error'
              ? '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-100 text-rose-700 border border-rose-200"><i class="fa-solid fa-xmark mr-0.5"></i>ไม่สำเร็จ</span>'
              : item.status === 'info'
                ? '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200"><i class="fa-solid fa-floppy-disk mr-0.5"></i>บันทึกแล้ว</span>'
                : '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-700 border border-emerald-200"><i class="fa-solid fa-check mr-0.5"></i>สำเร็จ</span>';

          const metaParts = [];
          if (item.site_name) metaParts.push(`โครงการ: ${escapeHtml(item.site_name)}`);
          if (item.date_from || item.date_to) {
            metaParts.push(`ช่วงวันที่: ${escapeHtml(item.date_from ? toBE(item.date_from) : '...')} - ${escapeHtml(item.date_to ? toBE(item.date_to) : '...')}`);
          }
          if (item.mode === 'manual' || item.mode === 'auto') {
            metaParts.push(`${formatInt(item.log_count)} เที่ยว` + (item.sheet_count ? ` • ${formatInt(item.sheet_count)} แท็บ` : ''));
          }
          if (item.triggered_by) metaParts.push(`โดย: ${escapeHtml(item.triggered_by)}`);

          const sheetLink = item.spreadsheet_name
            ? item.spreadsheet_url
              ? `<a href="${escapeHtml(item.spreadsheet_url)}" target="_blank" rel="noopener noreferrer" class="text-xs font-bold text-emerald-800 hover:underline inline-flex items-center gap-1">${escapeHtml(item.spreadsheet_name)} <i class="fa-solid fa-arrow-up-right-from-square text-[10px]"></i></a>`
              : `<span class="text-xs font-bold text-slate-800">${escapeHtml(item.spreadsheet_name)}</span>`
            : '';

          return `
            <div class="p-3 flex items-start justify-between gap-2.5 hover:bg-slate-50/80 transition">
              <div class="min-w-0 flex-1 space-y-1">
                <div class="flex flex-wrap items-center gap-1.5">
                  <span class="text-[11px] font-bold text-slate-700"><i class="fa-regular fa-clock mr-1 text-slate-400"></i>${escapeHtml(formatDateTimeBE(item.synced_at))}</span>
                  ${kindBadge}
                  ${modeBadge}
                  ${statusBadge}
                </div>
                ${sheetLink ? `<div>${sheetLink}</div>` : ''}
                ${item.note ? `<p class="text-xs text-slate-700 break-words">${escapeHtml(item.note)}</p>` : ''}
                ${item.error_message ? `<p class="text-xs text-rose-600 font-medium break-words"><i class="fa-solid fa-triangle-exclamation mr-1"></i>${escapeHtml(item.error_message)}</p>` : ''}
                ${metaParts.length ? `<p class="text-[11px] text-slate-500">${metaParts.join(' • ')}</p>` : ''}
              </div>
              <button
                type="button"
                data-del-gsh-id="${escapeHtml(item.id)}"
                title="ลบบันทึกนี้"
                class="w-7 h-7 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 flex items-center justify-center transition cursor-pointer shrink-0"
              >
                <i class="fa-solid fa-trash-can text-[11px]"></i>
              </button>
            </div>
          `;
        })
        .join('');
    };

    const countStats = (list) => ({
      total: list.length,
      ok: list.filter((x) => x.status === 'success').length,
      err: list.filter((x) => x.status === 'error').length
    });
    const st0 = countStats(rows);

    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center shrink-0">
              <i class="fa-solid fa-clock-rotate-left text-lg"></i>
            </div>
            <div>
              <h3 class="text-base font-bold text-slate-900">บันทึกประวัติการไหลลง Google Sheets</h3>
              <p class="text-xs text-slate-500">เก็บประวัติอัตโนมัติสูงสุด 100 รายการล่าสุด (ใช้ร่วมกันทุกเครื่อง)</p>
            </div>
          </div>
          <div class="flex items-center gap-1.5 text-[11px] font-bold">
            <span id="m-gsh-stat-total" class="px-2 py-0.5 rounded-full bg-slate-100 text-slate-700">ทั้งหมด ${formatInt(st0.total)}</span>
            <span id="m-gsh-stat-ok" class="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800">ไหลสำเร็จ ${formatInt(st0.ok)}</span>
            <span id="m-gsh-stat-err" class="px-2 py-0.5 rounded-full bg-rose-100 text-rose-700">ผิดพลาด ${formatInt(st0.err)}</span>
          </div>
        </div>

        <!-- เพิ่มบันทึกเอง -->
        <form id="m-gsh-form" class="p-3 rounded-xl bg-emerald-50/70 border border-emerald-200 space-y-2">
          <div class="text-xs font-bold text-emerald-950">
            <i class="fa-solid fa-pen-to-square text-emerald-600 mr-1"></i>เพิ่มบันทึกประวัติ / หมายเหตุด้วยตนเอง
          </div>
          <div class="flex flex-col sm:flex-row gap-2">
            <select id="m-gsh-kind" class="sm:w-36 min-h-[38px] px-2.5 py-1.5 rounded-lg border border-emerald-300 bg-white text-xs font-semibold text-slate-800">
              <option value="zone">ชีตสำเนาโซน</option>
              <option value="billing">ชีตใบเบิก</option>
              <option value="general">ทั่วไป</option>
            </select>
            <input
              type="text"
              id="m-gsh-note"
              placeholder="เช่น ตรวจยอดงวด 1-15 บนชีตเรียบร้อยแล้ว / ส่งลิงก์ให้บัญชีแล้ว..."
              class="flex-1 min-h-[38px] px-3 py-1.5 rounded-lg border border-emerald-300 bg-white text-xs text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-600"
            />
            <button
              type="submit"
              id="m-gsh-submit"
              class="min-h-[38px] px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold flex items-center justify-center gap-1.5 cursor-pointer shrink-0"
            >
              <i class="fa-solid fa-floppy-disk"></i><span>บันทึก</span>
            </button>
          </div>
        </form>

        <!-- รายการประวัติ -->
        <div id="m-gsh-list" class="max-h-[46vh] overflow-y-auto divide-y divide-slate-100 border border-slate-200 rounded-xl bg-white">
          ${renderListHtml(rows)}
        </div>

        <div class="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
          <div>
            ${
              isAdmin
                ? `<button type="button" id="m-gsh-clear" class="min-h-[36px] px-3 py-1.5 rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-semibold cursor-pointer">
                    <i class="fa-solid fa-trash-can mr-1"></i>ล้างประวัติทั้งหมด
                  </button>`
                : ''
            }
          </div>
          <button type="button" id="m-gsh-close" class="min-h-[38px] px-4 py-1.5 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-white text-xs font-semibold cursor-pointer">
            ปิดหน้าต่าง
          </button>
        </div>
      </div>
      `,
      (box) => {
        const listEl = box.querySelector('#m-gsh-list');
        const refreshModalList = (nextList) => {
          applyGoogleSheetHistoryMeta(nextList);
          renderGoogleSheetsUI();
          const st = countStats(nextList);
          const elT = box.querySelector('#m-gsh-stat-total');
          const elOk = box.querySelector('#m-gsh-stat-ok');
          const elErr = box.querySelector('#m-gsh-stat-err');
          if (elT) elT.textContent = `ทั้งหมด ${formatInt(st.total)}`;
          if (elOk) elOk.textContent = `ไหลสำเร็จ ${formatInt(st.ok)}`;
          if (elErr) elErr.textContent = `ผิดพลาด ${formatInt(st.err)}`;
          if (listEl) listEl.innerHTML = renderListHtml(nextList);
        };

        box.querySelector('#m-gsh-close').addEventListener('click', closeModal);

        box.querySelector('#m-gsh-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          const noteEl = box.querySelector('#m-gsh-note');
          const kindEl = box.querySelector('#m-gsh-kind');
          const submitBtn = box.querySelector('#m-gsh-submit');
          const note = noteEl.value.trim();
          if (!note) {
            showToast('กรุณากรอกรายละเอียดที่ต้องการบันทึก', 'warning');
            noteEl.focus();
            return;
          }
          submitBtn.disabled = true;
          try {
            const res = await apiFetch('/api/gsheet/history', {
              method: 'POST',
              body: JSON.stringify({
                kind: kindEl.value,
                note,
                site_name: f.site_name,
                date_from: f.date_from,
                date_to: f.date_to,
                log_count: (state.logs || []).length
              })
            });
            noteEl.value = '';
            refreshModalList(res.history || []);
            showToast('เพิ่มบันทึกประวัติเรียบร้อยแล้ว', 'success');
          } catch (err) {
            showToast(err.message || 'ไม่สามารถบันทึกประวัติได้', 'error');
          } finally {
            submitBtn.disabled = false;
          }
        });

        listEl.addEventListener('click', async (e) => {
          const delBtn = e.target.closest('button[data-del-gsh-id]');
          if (!delBtn) return;
          const id = delBtn.dataset.delGshId;
          delBtn.disabled = true;
          try {
            const res = await apiFetch(`/api/gsheet/history/${encodeURIComponent(id)}`, { method: 'DELETE' });
            refreshModalList(res.history || []);
            showToast('ลบบันทึกประวัติแล้ว', 'info');
          } catch (err) {
            showToast(err.message || 'ลบไม่สำเร็จ', 'error');
            delBtn.disabled = false;
          }
        });

        const clearBtn = box.querySelector('#m-gsh-clear');
        if (clearBtn) {
          clearBtn.addEventListener('click', async () => {
            if (!window.confirm('ต้องการล้างประวัติการไหลข้อมูล Google Sheets ทั้งหมดใช่หรือไม่?')) return;
            clearBtn.disabled = true;
            try {
              await apiFetch('/api/gsheet/history', { method: 'DELETE' });
              refreshModalList([]);
              showToast('ล้างประวัติการไหลข้อมูลทั้งหมดแล้ว', 'info');
            } catch (err) {
              showToast(err.message || 'ล้างประวัติไม่สำเร็จ', 'error');
            } finally {
              clearBtn.disabled = false;
            }
          });
        }
      }
    );
  }

  async function syncToGoogleSheet(kind, { interactive = true, mode = '', siteName = '', triggerBtn = null } = {}) {
    const isBilling = kind === 'billing';
    const kindThai = isBilling ? 'ใบเบิก' : 'สำเนา';
    const fixedSite = String(siteName || '').trim();
    const cfg = state.gsheet || {};
    const siteMap = isBilling ? (cfg.billing_by_site || {}) : (cfg.zone_by_site || {});
    const targetUrl = fixedSite
      ? ((siteMap[fixedSite] && siteMap[fixedSite].url) || '')
      : isBilling
        ? cfg.billing_url
        : cfg.zone_url;
    if (!targetUrl) {
      if (interactive) openGoogleSheetBindModal(kind, fixedSite);
      return { status: 'unbound' };
    }

    const btn = triggerBtn || $(isBilling ? 'btn-gsheet-sync-billing' : 'btn-gsheet-sync-zone');
    const origHtml = btn ? btn.innerHTML : '';
    if (btn && interactive) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i><span>กำลังไหลลงชีต...</span>';
    }

    try {
      const f = getFilterParams();
      const res = await apiFetch('/api/gsheet/sync', {
        method: 'POST',
        body: JSON.stringify({
          kind,
          mode: mode || (interactive ? 'manual' : 'auto'),
          site_name: fixedSite ? fixedSite : isBilling ? (cfg.billing_site || f.site_name) : f.site_name,
          per_site_only: Boolean(fixedSite),
          date_from: f.date_from,
          date_to: f.date_to,
          trip_type: fixedSite || isBilling ? '' : f.trip_type,
          search: fixedSite || isBilling ? '' : f.search,
          issue_date: todayIso()
        })
      });
      if (res.config) state.gsheet = { ...state.gsheet, ...res.config };
      if (res.historyEntry) state.gsheetLastHistory = res.historyEntry;
      if (typeof res.historyCount === 'number') state.gsheetHistoryCount = res.historyCount;
      renderGoogleSheetsUI();

      const labelTarget = fixedSite
        ? `${kindThai}โครงการ "${fixedSite}" (${res.spreadsheetName || kindThai})`
        : `"${res.spreadsheetName || (isBilling ? 'ใบเบิก' : 'สำเนาโซน')}"`;
      if (interactive) {
        if (isBilling && res.scriptVersion !== LATEST_APPS_SCRIPT_VERSION) {
          showToast(
            `ไหลลงชีต${labelTarget} สำเร็จ (${res.logCount} เที่ยว) แต่โค้ด Apps Script ในชีตยังเป็นเวอร์ชันเก่าเลยยังไม่มีโลโก้ — กดคัดลอกโค้ดใหม่ไปวางแล้วเลือก "จัดการการทำให้ใช้งานได้ → แก้ไข ✏️ → เวอร์ชันใหม่"`,
            'warning',
            8000
          );
        } else {
          showToast(
            `ไหลข้อมูลลง Google Sheet ${labelTarget} สำเร็จ (${res.sheetCount} แท็บ, ${res.logCount} เที่ยว)`,
            'success',
            4500
          );
        }
      }
      return { status: 'ok', name: fixedSite ? `${kindThai}:${fixedSite}` : res.spreadsheetName || kindThai };
    } catch (err) {
      loadGoogleSheetsConfig();
      if (interactive) showToast(`ไหลลง Google Sheets${fixedSite ? ` (${kindThai}: ${fixedSite})` : ''} ไม่สำเร็จ: ${err.message}`, 'error', 5500);
      return { status: 'error', error: err };
    } finally {
      if (btn && interactive) {
        btn.disabled = false;
        btn.innerHTML = origHtml;
      }
    }
  }

  async function syncAllBoundGoogleSheetSites(kind = 'billing') {
    const isZone = kind === 'zone';
    const kindThai = isZone ? 'สำเนา' : 'ใบเบิก';
    const bySite = (state.gsheet && (isZone ? state.gsheet.zone_by_site : state.gsheet.billing_by_site)) || {};
    const sites = Object.keys(bySite).filter((s) => bySite[s] && bySite[s].url);
    if (sites.length === 0) {
      showToast(`ยังไม่ได้ผูก Google Sheet ${kindThai}ให้โครงการใด — กดปุ่ม "ผูกชีต${kindThai}" ที่การ์ดโครงการด้านล่างก่อนครับ`, 'warning', 4500);
      return;
    }
    const btn = $(isZone ? 'btn-sync-all-gsheet-zone-sites' : 'btn-sync-all-gsheet-sites');
    const origHtml = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i><span>กำลังไหล${kindThai}ทุกโครงการ...</span>`;
    }
    try {
      const okNames = [];
      for (const sName of sites) {
        const r = await syncToGoogleSheet(kind, { interactive: false, mode: 'manual', siteName: sName });
        if (r.status === 'ok') okNames.push(sName);
      }
      if (okNames.length > 0) {
        showToast(`ไหลลง Google Sheet ${kindThai}แยกสำเร็จครบ ${okNames.length} โครงการ: ${okNames.join(', ')}`, 'success', 5000);
      } else {
        showToast(`ไม่สามารถไหลลง Google Sheet ${kindThai}บางโครงการได้ กรุณาตรวจสอบในบันทึกประวัติการไหล`, 'error', 5000);
      }
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = origHtml;
      }
    }
  }

  async function autoSyncGoogleSheets() {
    const cfg = state.gsheet || {};
    const zoneBySite = cfg.zone_by_site || {};
    const billingBySite = cfg.billing_by_site || {};
    const zoneSiteKeys = Object.keys(zoneBySite).filter((s) => zoneBySite[s] && zoneBySite[s].url);
    const billingSiteKeys = Object.keys(billingBySite).filter((s) => billingBySite[s] && billingBySite[s].url);
    if (!cfg.auto_sync) return;
    if (!cfg.zone_url && !cfg.billing_url && zoneSiteKeys.length === 0 && billingSiteKeys.length === 0) return;

    const syncedNames = [];
    if (cfg.zone_url) {
      const r = await syncToGoogleSheet('zone', { interactive: false, mode: 'auto' });
      if (r.status === 'ok') syncedNames.push(r.name);
    }
    if (cfg.billing_url) {
      const r = await syncToGoogleSheet('billing', { interactive: false, mode: 'auto' });
      if (r.status === 'ok') syncedNames.push(r.name);
    }
    for (const sName of zoneSiteKeys) {
      const r = await syncToGoogleSheet('zone', { interactive: false, mode: 'auto', siteName: sName });
      if (r.status === 'ok') syncedNames.push(r.name);
    }
    for (const sName of billingSiteKeys) {
      const r = await syncToGoogleSheet('billing', { interactive: false, mode: 'auto', siteName: sName });
      if (r.status === 'ok') syncedNames.push(r.name);
    }
    if (syncedNames.length > 0) {
      showToast(`ไหลลง Google Sheets อัตโนมัติแล้ว: ${syncedNames.join(', ')}`, 'info', 3500);
    }
  }

  async function runZoneExport({ siteName, dateFrom, dateTo }) {
    const { wb, filename, logs } = await buildZonePayload({ siteName, dateFrom, dateTo });
    await downloadWorkbook(wb, filename);
    showToast(
      `ส่งออกไฟล์สำเนาโครงการ "${siteName}" สำเร็จ: ${filename} (${wb.worksheets.length} ชีต, ${logs.length} เที่ยว)`,
      'success',
      4500
    );
  }

  async function handleExportZoneCopy() {
    if (typeof window.ExcelJS === 'undefined' || !window.CkpExcel) {
      showToast('ไม่พบไลบรารี ExcelJS กรุณารีเฟรชหน้าเว็บ', 'error');
      return;
    }

    const btn = $('btn-export-zone');
    btn.disabled = true;
    try {
      const { wb, filename, logs } = await buildZonePayloadFromFilters();
      await downloadWorkbook(wb, filename);
      showToast(
        `ส่งออกไฟล์สำเนาโซนสำเร็จ: ${filename} (${wb.worksheets.length} ชีต, ${logs.length} เที่ยว)`,
        'success',
        4500
      );
    } catch (err) {
      showToast(`ส่งออกไฟล์สำเนาโซนไม่สำเร็จ: ${err.message}`, 'error', 5000);
    } finally {
      btn.disabled = false;
    }
  }

  async function runBillingExport({ siteName, dateFrom, dateTo, issueDate }) {
    const { wb, filename, logs } = await buildBillingPayload({ siteName, dateFrom, dateTo, issueDate });
    await downloadWorkbook(wb, filename);
    showToast(
      `ส่งออกไฟล์ใบเบิกโครงการสำเร็จ: ${filename} (${wb.worksheets.length} ชีต, ${logs.length} เที่ยว)`,
      'success',
      4500
    );
  }

  async function handleExportBilling() {
    if (typeof window.ExcelJS === 'undefined' || !window.CkpExcel) {
      showToast('ไม่พบไลบรารี ExcelJS กรุณารีเฟรชหน้าเว็บ', 'error');
      return;
    }

    const f = getFilterParams();
    const siteSet = new Set(state.sites.map((s) => s.site_name).filter(Boolean));
    for (const l of state.logs) {
      if (l.site_name) siteSet.add(l.site_name);
    }
    const siteOptions = [...siteSet];
    const preselected = f.site_name || state.boundBillingSite || (siteOptions.length === 1 ? siteOptions[0] : '');

    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <div class="flex items-start gap-3">
          <div class="w-10 h-10 rounded-xl bg-zinc-900 text-amber-400 flex items-center justify-center shrink-0">
            <i class="fa-solid fa-file-invoice-dollar"></i>
          </div>
          <div>
            <h3 class="text-base font-bold text-slate-900">ส่งออกไฟล์ใบเบิกโครงการ (.xlsx)</h3>
            <p class="text-xs text-slate-500 mt-0.5">
              สร้างชีต "ใบเบิก" พร้อมโลโก้ CKP สูตรคำนวณ และชีตแนบท้ายตัดหน้าอัตโนมัติแผ่นละ 30 แถว (A4 แนวตั้ง)
            </p>
          </div>
        </div>

        <form id="modal-billing-form" class="space-y-3.5">
          <div>
            <label for="modal-billing-site" class="block text-xs font-semibold text-slate-700 mb-1">
              โครงการที่วางบิล (site_name)
            </label>
            <select id="modal-billing-site" class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-zinc-900">
              <option value="">-- ทุกโครงการในช่วงวันที่เลือก --</option>
              ${siteOptions
                .map(
                  (s) =>
                    `<option value="${escapeHtml(s)}" ${s === preselected ? 'selected' : ''}>${escapeHtml(s)}</option>`
                )
                .join('')}
            </select>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label for="modal-billing-from" class="block text-xs font-semibold text-slate-700 mb-1">ตั้งแต่วันที่</label>
              <input type="date" id="modal-billing-from" value="${escapeHtml(f.date_from)}"
                class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-zinc-900" />
            </div>
            <div>
              <label for="modal-billing-to" class="block text-xs font-semibold text-slate-700 mb-1">ถึงวันที่</label>
              <input type="date" id="modal-billing-to" value="${escapeHtml(f.date_to)}"
                class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-zinc-900" />
            </div>
          </div>

          <div>
            <label for="modal-billing-issue" class="block text-xs font-semibold text-slate-700 mb-1">วันที่ออกเอกสาร (เซลล์ J2 ในใบเบิก)</label>
            <input type="date" id="modal-billing-issue" value="${escapeHtml(todayIso())}" required
              class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-zinc-900" />
          </div>

          <div class="flex flex-wrap items-center justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="modal-billing-cancel" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs sm:text-sm font-medium cursor-pointer">ยกเลิก</button>
            ${
              state.boundBillingHandle
                ? `<button type="button" id="modal-billing-sync-bound" class="min-h-[40px] px-4 py-2 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs sm:text-sm font-semibold cursor-pointer"><i class="fa-solid fa-link mr-1.5"></i>บันทึกเข้าไฟล์ที่ผูกไว้</button>`
                : ''
            }
            <button type="submit" id="modal-billing-submit" class="min-h-[40px] px-5 py-2 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-white text-xs sm:text-sm font-semibold cursor-pointer">
              <i class="fa-solid fa-download text-amber-400 mr-1.5"></i>สร้างและดาวน์โหลด (.xlsx)
            </button>
          </div>
        </form>
      </div>
      `,
      (box) => {
        const readOpts = () => ({
          siteName: box.querySelector('#modal-billing-site').value.trim(),
          dateFrom: box.querySelector('#modal-billing-from').value.trim(),
          dateTo: box.querySelector('#modal-billing-to').value.trim(),
          issueDate: box.querySelector('#modal-billing-issue').value.trim()
        });
        box.querySelector('#modal-billing-cancel').addEventListener('click', closeModal);
        const syncBoundBtn = box.querySelector('#modal-billing-sync-bound');
        if (syncBoundBtn) {
          syncBoundBtn.addEventListener('click', async () => {
            syncBoundBtn.disabled = true;
            try {
              const opts = readOpts();
              closeModal();
              await syncToBoundFile('billing', { interactive: true, billingOpts: opts });
            } catch (err) {
              showToast(`บันทึกไฟล์ใบเบิกไม่สำเร็จ: ${err.message}`, 'error', 5000);
            }
          });
        }
        box.querySelector('#modal-billing-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          const btn = box.querySelector('#modal-billing-submit');
          btn.disabled = true;
          try {
            await runBillingExport(readOpts());
            closeModal();
          } catch (err) {
            showToast(`ส่งออกไฟล์ใบเบิกไม่สำเร็จ: ${err.message}`, 'error', 5000);
            btn.disabled = false;
          }
        });
      }
    );
  }

  // ==========================================================================
  // 10. Back-Office Master Data CRUD (4 Tables + Quota Settings)
  // ==========================================================================
  function renderBackOfficeLists() {
    renderVehiclesList();
    renderWorkersList();
    renderSitesList();
    renderRoutesList();
  }

  function renderVehiclesList() {
    const el = $('list-vehicles');
    if (state.vehicles.length === 0) {
      el.innerHTML =
        '<div class="col-span-full py-10 text-center text-xs text-slate-400 border border-dashed border-slate-300 rounded-xl">ยังไม่มีทะเบียนรถ — พิมพ์เลขทะเบียนในช่องด้านบน แล้วกด "เพิ่มทะเบียนรถ"</div>';
      return;
    }
    el.innerHTML = state.vehicles
      .map(
        (v) => `
        <div class="p-3.5 rounded-xl border ${v.active ? 'bg-white border-slate-200' : 'bg-slate-50 border-slate-200 opacity-65'} flex items-center justify-between gap-2">
          <div class="flex items-center gap-2 min-w-0">
            <span class="license-plate-badge">${escapeHtml(v.license_plate)}</span>
            ${v.active ? '' : '<span class="text-[10px] px-1.5 py-0.5 rounded bg-slate-200 text-slate-600 font-semibold">ปิดใช้งาน</span>'}
          </div>
          <div class="flex items-center gap-1 shrink-0">
            <button type="button" data-bo="toggle-vehicle" data-id="${v.id}" title="${v.active ? 'ปิดใช้งาน' : 'เปิดใช้งาน'}"
              class="w-8 h-8 rounded-lg ${v.active ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-200 text-slate-600 border-slate-300'} border inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid ${v.active ? 'fa-toggle-on' : 'fa-toggle-off'}"></i>
            </button>
            <button type="button" data-bo="edit-vehicle" data-id="${v.id}" title="แก้ไข"
              class="w-8 h-8 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-pen text-xs"></i>
            </button>
            <button type="button" data-bo="delete-vehicle" data-id="${v.id}" title="ลบ"
              class="w-8 h-8 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-trash text-xs"></i>
            </button>
          </div>
        </div>
      `
      )
      .join('');
  }

  function renderWorkersList() {
    const el = $('list-workers');
    if (state.workers.length === 0) {
      el.innerHTML = '<div class="py-10 text-center text-xs text-slate-400">ยังไม่มีคนขับ — พิมพ์ชื่อในช่องด้านบน เลือกผูกทะเบียนรถ/โครงการ/เส้นทาง แล้วกด "เพิ่มคนขับ"</div>';
      return;
    }
    el.innerHTML = state.workers
      .map((w, i) => {
        const routeBadge =
          w.default_origin || w.default_destination
            ? `<span class="px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 text-[11px] font-semibold"><i class="fa-solid fa-route text-[10px] mr-1"></i>${escapeHtml(w.default_origin || '-')} → ${escapeHtml(w.default_destination || w.default_site || '-')}</span>`
            : '';
        return `
        <div class="px-4 py-3 flex flex-wrap items-center justify-between gap-3 ${w.active ? 'bg-white' : 'bg-slate-50 opacity-65'}">
          <div class="flex flex-wrap items-center gap-2">
            <span class="text-xs font-bold text-slate-400 w-6">${i + 1}.</span>
            <span class="font-semibold text-sm text-slate-900">${escapeHtml(w.name)}</span>
            ${
              w.default_vehicle
                ? `<span class="license-plate-badge !text-[11px] !py-0.5 !px-2"><i class="fa-solid fa-truck-front text-amber-600 text-[10px]"></i>${escapeHtml(w.default_vehicle)}</span>`
                : '<span class="text-[11px] text-slate-400">ยังไม่ผูกรถ</span>'
            }
            ${
              w.default_site
                ? `<span class="px-2 py-0.5 rounded-full bg-amber-50 text-amber-900 border border-amber-200 text-[11px] font-semibold"><i class="fa-solid fa-building text-[10px] mr-1"></i>${escapeHtml(w.default_site)}</span>`
                : ''
            }
            ${routeBadge}
            ${w.active ? '' : '<span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-200 text-slate-600 font-semibold">ปิดใช้งาน</span>'}
          </div>
          <div class="flex items-center gap-1.5">
            <button type="button" data-bo="toggle-worker" data-id="${w.id}" title="${w.active ? 'ปิดใช้งาน' : 'เปิดใช้งาน'}"
              class="w-8 h-8 rounded-lg ${w.active ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-200 text-slate-600 border-slate-300'} border inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid ${w.active ? 'fa-toggle-on' : 'fa-toggle-off'}"></i>
            </button>
            <button type="button" data-bo="edit-worker" data-id="${w.id}" title="แก้ไข"
              class="w-8 h-8 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-pen text-xs"></i>
            </button>
            <button type="button" data-bo="delete-worker" data-id="${w.id}" title="ลบ"
              class="w-8 h-8 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-trash text-xs"></i>
            </button>
          </div>
        </div>
      `;
      })
      .join('');
  }

  function getSiteBoundOrigins(s) {
    if (!s) return [];
    const arr = Array.isArray(s.default_origins) ? s.default_origins : [];
    const seen = new Set();
    const out = [];
    for (const x of [...arr, s.default_origin]) {
      const v = String(x || '').trim();
      if (!v || seen.has(v)) continue;
      seen.add(v);
      out.push(v);
    }
    return out;
  }

  function getQuickRouteOriginChoices() {
    const seen = new Set();
    const out = [];
    for (const r of state.routes || []) {
      if (!r || !r.active) continue;
      const o = String(r.origin || '').trim();
      if (!o || seen.has(o)) continue;
      seen.add(o);
      out.push(o);
    }
    for (const s of state.sites || []) {
      for (const o of getSiteBoundOrigins(s)) {
        if (!o || seen.has(o)) continue;
        seen.add(o);
        out.push(o);
      }
    }
    return out;
  }

  function renderOriginTagPicker(tagsEl, quickEl, selectedList, { removeAttr, toggleAttr }) {
    const clean = [...new Set((selectedList || []).map((x) => String(x || '').trim()).filter(Boolean))];
    if (tagsEl) {
      if (clean.length === 0) {
        tagsEl.innerHTML = '<span class="text-[11px] text-slate-400">ยังไม่ได้ผูกเส้นทาง — กดเลือกจากรายการด้านล่าง หรือพิมพ์แล้วกด "+ เพิ่มเส้นทาง"</span>';
      } else {
        tagsEl.innerHTML = clean
          .map(
            (orig) => `
            <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-900 border border-emerald-300 text-xs font-semibold">
              <i class="fa-solid fa-route text-[10px] text-emerald-700"></i>
              <span>${escapeHtml(orig)}</span>
              <button type="button" ${removeAttr}="${escapeHtml(orig)}" title="เอาเส้นทางนี้ออก"
                class="w-4 h-4 rounded-full hover:bg-emerald-200 text-emerald-800 inline-flex items-center justify-center cursor-pointer">
                <i class="fa-solid fa-xmark text-[10px]"></i>
              </button>
            </span>
          `
          )
          .join('');
      }
    }
    if (quickEl) {
      const choices = getQuickRouteOriginChoices();
      if (choices.length === 0) {
        quickEl.innerHTML = '';
      } else {
        const selectedSet = new Set(clean);
        quickEl.innerHTML =
          '<span class="text-[11px] text-slate-500 font-medium mr-1">คลิกเลือกเส้นทางด่วน:</span>' +
          choices
            .map((orig) => {
              const active = selectedSet.has(orig);
              return `
                <button type="button" ${toggleAttr}="${escapeHtml(orig)}"
                  class="px-2 py-0.5 rounded-full text-[11px] font-semibold border transition cursor-pointer inline-flex items-center gap-1 ${
                    active
                      ? 'bg-emerald-600 text-white border-emerald-700 shadow-xs'
                      : 'bg-white hover:bg-emerald-50 text-slate-700 border-slate-300'
                  }">
                  <i class="fa-solid ${active ? 'fa-check' : 'fa-plus'} text-[9px]"></i>
                  <span>${escapeHtml(orig)}</span>
                </button>
              `;
            })
            .join('');
      }
    }
  }

  function renderAddSiteOriginPicker() {
    renderOriginTagPicker($('site-add-origins-tags'), $('site-add-quick-routes'), state.newSiteOrigins, {
      removeAttr: 'data-remove-new-site-origin',
      toggleAttr: 'data-toggle-new-site-origin'
    });
  }

  function renderSitesList() {
    renderAddSiteOriginPicker();
    const el = $('list-sites');
    if (state.sites.length === 0) {
      el.innerHTML = '<div class="py-10 text-center text-xs text-slate-400">ยังไม่มีโครงการ — พิมพ์ชื่อโครงการในช่องด้านบน แล้วกด "เพิ่มโครงการ"</div>';
      return;
    }
    el.innerHTML = state.sites
      .map((s, i) => {
        const directOrigins = getSiteBoundOrigins(s);
        const routeLinked = (state.routes || [])
          .filter((r) => r && r.active && (r.default_site === s.site_name || r.destination === s.site_name))
          .map((r) => r.origin);
        const allOrigins = [...new Set([...directOrigins, ...routeLinked])].filter(Boolean);
        return `
        <div class="px-4 py-3 flex flex-wrap items-center justify-between gap-3 ${s.active ? 'bg-white' : 'bg-slate-50 opacity-65'}">
          <div class="flex flex-wrap items-center gap-2 min-w-0 flex-1">
            <span class="text-xs font-bold text-slate-400 w-6 shrink-0">${i + 1}.</span>
            <span class="font-semibold text-sm text-slate-900">${escapeHtml(s.site_name)}</span>
            ${
              s.zone_name
                ? `<span class="px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-900 border border-amber-200 text-[11px] font-semibold">โซน: ${escapeHtml(s.zone_name)}</span>`
                : '<span class="text-[11px] text-slate-400">ไม่ระบุชื่อโซน</span>'
            }
            ${allOrigins
              .map(
                (orig) =>
                  `<span class="px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 text-[11px] font-semibold"><i class="fa-solid fa-route text-[10px] mr-1"></i>${escapeHtml(orig)}</span>`
              )
              .join('')}
            ${s.active ? '' : '<span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-200 text-slate-600 font-semibold">ปิดใช้งาน</span>'}
          </div>
          <div class="flex items-center gap-1.5 shrink-0">
            <button type="button" data-bo="edit-site" data-id="${s.id}" title="ผูกหรือแก้ไขเส้นทางของโครงการนี้"
              class="min-h-[32px] px-2.5 py-1 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 text-[11px] font-semibold inline-flex items-center gap-1 cursor-pointer transition">
              <i class="fa-solid fa-route text-[10px]"></i>
              <span>${allOrigins.length > 0 ? `เส้นทาง (${allOrigins.length})` : '+ ผูกเส้นทาง'}</span>
            </button>
            <button type="button" data-bo="toggle-site" data-id="${s.id}" title="${s.active ? 'ปิดใช้งาน' : 'เปิดใช้งาน'}"
              class="w-8 h-8 rounded-lg ${s.active ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-200 text-slate-600 border-slate-300'} border inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid ${s.active ? 'fa-toggle-on' : 'fa-toggle-off'}"></i>
            </button>
            <button type="button" data-bo="edit-site" data-id="${s.id}" title="แก้ไข"
              class="w-8 h-8 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-pen text-xs"></i>
            </button>
            <button type="button" data-bo="delete-site" data-id="${s.id}" title="ลบ"
              class="w-8 h-8 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-trash text-xs"></i>
            </button>
          </div>
        </div>
      `;
      })
      .join('');
  }

  function renderRoutesList() {
    const el = $('list-routes');
    if (state.routes.length === 0) {
      el.innerHTML = '<div class="py-10 text-center text-xs text-slate-400">ยังไม่มีเส้นทาง — พิมพ์ชื่อต้นทาง (จุดรับ) ด้านบน แล้วกด "เพิ่มเส้นทาง"</div>';
      return;
    }
    el.innerHTML = state.routes
      .map(
        (r, i) => `
        <div class="px-4 py-3.5 flex flex-col lg:flex-row lg:items-center justify-between gap-3 ${r.active ? 'bg-white' : 'bg-slate-50 opacity-65'}">
          <div class="space-y-1">
            <div class="flex flex-wrap items-center gap-2">
              <span class="text-xs font-bold text-slate-400">${i + 1}.</span>
              <span class="font-bold text-sm text-slate-900">
                ${r.destination ? `${escapeHtml(r.origin)} <i class="fa-solid fa-arrow-right text-amber-500 text-xs mx-1"></i> ${escapeHtml(r.destination)}` : escapeHtml(r.origin)}
              </span>
              ${
                r.default_site
                  ? `<span class="px-2 py-0.5 rounded-full bg-amber-50 text-amber-900 border border-amber-200 text-[11px] font-semibold"><i class="fa-solid fa-building text-[10px] mr-1"></i>โครงการ: ${escapeHtml(r.default_site)}</span>`
                  : ''
              }
              ${
                r.unit_price > 0
                  ? `<span class="px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-bold">${formatMoney(r.unit_price)} บาท/เที่ยว</span>`
                  : ''
              }
              ${r.active ? '' : '<span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-200 text-slate-600 font-semibold">ปิดใช้งาน</span>'}
            </div>
            <div class="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-600">
              <span><strong class="text-slate-700">แผ่นงาน Excel:</strong> ${escapeHtml(r.route_sheet_name)}</span>
              <span><strong class="text-slate-700">ชื่อในใบเบิก:</strong> ${escapeHtml(r.billing_title)}</span>
              ${
                r.default_dep_time || r.default_arr_time
                  ? `<span><strong class="text-slate-700">เวลา:</strong> ${escapeHtml(r.default_dep_time || '-')} – ${escapeHtml(r.default_arr_time || '-')}</span>`
                  : ''
              }
              <span>
                <strong class="text-slate-700">โซนคอลัมน์ C:</strong>
                <span class="px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-800 font-semibold text-[11px]">
                  ${r.display_mode_zone === 'license_plate' ? 'สำเนาโซนแสดง: ทะเบียนรถ' : 'สำเนาโซนแสดง: ชื่อคนขับ'}
                </span>
              </span>
            </div>
          </div>

          <div class="flex items-center gap-1.5 shrink-0">
            <button type="button" data-bo="toggle-route" data-id="${r.id}" title="${r.active ? 'ปิดใช้งาน' : 'เปิดใช้งาน'}"
              class="w-8 h-8 rounded-lg ${r.active ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-200 text-slate-600 border-slate-300'} border inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid ${r.active ? 'fa-toggle-on' : 'fa-toggle-off'}"></i>
            </button>
            <button type="button" data-bo="edit-route" data-id="${r.id}" title="แก้ไข"
              class="w-8 h-8 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-pen text-xs"></i>
            </button>
            <button type="button" data-bo="delete-route" data-id="${r.id}" title="ลบ"
              class="w-8 h-8 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 inline-flex items-center justify-center cursor-pointer">
              <i class="fa-solid fa-trash text-xs"></i>
            </button>
          </div>
        </div>
      `
      )
      .join('');
  }

  // Add Handlers
  async function handleAddVehicle(e) {
    e.preventDefault();
    const plate = $('input-vehicle-plate').value.trim();
    if (!plate) return;
    try {
      await apiFetch('/api/master/vehicles', { method: 'POST', body: JSON.stringify({ license_plate: plate }) });
      $('input-vehicle-plate').value = '';
      showToast(`เพิ่มทะเบียนรถ "${plate}" เรียบร้อยแล้ว`, 'success');
      await loadMasterData();
    } catch (err) {
      showToast(err.message, 'error');
    }
  }

  async function handleAddWorker(e) {
    e.preventDefault();
    const name = $('input-worker-name').value.trim();
    const default_vehicle = $('input-worker-default-vehicle').value.trim();
    const default_site = $('input-worker-default-site') ? $('input-worker-default-site').value.trim() : '';
    const default_origin = $('input-worker-default-origin') ? $('input-worker-default-origin').value.trim() : '';
    const default_destination = $('input-worker-default-destination') ? $('input-worker-default-destination').value.trim() : '';
    if (!name) return;
    try {
      await apiFetch('/api/master/workers', {
        method: 'POST',
        body: JSON.stringify({ name, default_vehicle, default_site, default_origin, default_destination })
      });
      $('input-worker-name').value = '';
      $('input-worker-default-vehicle').value = '';
      if ($('input-worker-default-site')) $('input-worker-default-site').value = '';
      if ($('input-worker-default-origin')) $('input-worker-default-origin').value = '';
      if ($('input-worker-default-destination')) $('input-worker-default-destination').value = '';
      showToast(`เพิ่มรายชื่อคนขับ "${name}" เรียบร้อยแล้ว`, 'success');
      await loadMasterData();
    } catch (err) {
      showToast(err.message, 'error');
    }
  }

  async function handleAddSite(e) {
    e.preventDefault();
    const site_name = $('input-site-name').value.trim();
    const zone_name = $('input-site-zone').value.trim();
    const typedOrigin = $('input-site-default-origin') ? $('input-site-default-origin').value.trim() : '';
    const default_origins = [...new Set([...(state.newSiteOrigins || []), ...(typedOrigin ? [typedOrigin] : [])])];
    if (!site_name) return;
    try {
      await apiFetch('/api/master/sites', {
        method: 'POST',
        body: JSON.stringify({
          site_name,
          zone_name,
          default_origin: default_origins[0] || '',
          default_origins
        })
      });
      $('input-site-name').value = '';
      $('input-site-zone').value = '';
      if ($('input-site-default-origin')) $('input-site-default-origin').value = '';
      state.newSiteOrigins = [];
      renderAddSiteOriginPicker();
      showToast(`เพิ่มไซต์งาน/โครงการ "${site_name}" เรียบร้อยแล้ว`, 'success');
      await loadMasterData();
    } catch (err) {
      showToast(err.message, 'error');
    }
  }

  async function handleAddRoute(e) {
    e.preventDefault();
    const origin = $('input-route-origin').value.trim();
    if (!origin) {
      showToast('กรุณาระบุต้นทาง (จุดรับ)', 'error');
      return;
    }
    const payload = {
      origin,
      destination: $('input-route-destination') ? $('input-route-destination').value.trim() : '',
      default_site: $('input-route-default-site') ? $('input-route-default-site').value.trim() : '',
      route_sheet_name: $('input-route-sheet-name').value.trim(),
      billing_title: $('input-route-billing-title').value.trim(),
      default_dep_time: $('input-route-dep-time').value.trim(),
      default_arr_time: $('input-route-arr-time').value.trim(),
      unit_price: $('input-route-unit-price').value.trim() === '' ? 0 : Number($('input-route-unit-price').value),
      display_mode_zone: $('input-route-display-mode').value
    };
    try {
      await apiFetch('/api/master/routes', { method: 'POST', body: JSON.stringify(payload) });
      $('form-add-route').reset();
      showToast(`เพิ่มเส้นทาง "${origin}" เรียบร้อยแล้ว`, 'success');
      await loadMasterData();
    } catch (err) {
      showToast(err.message, 'error');
    }
  }

  // Edit Modals for Back-Office Items
  function openEditVehicleModal(v) {
    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <h3 class="text-base font-bold text-slate-900">แก้ไขทะเบียนรถ</h3>
        <form id="modal-edit-form" class="space-y-3.5">
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">เลขทะเบียนรถ *</label>
            <input type="text" id="m-plate" value="${escapeHtml(v.license_plate)}" required
              class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
          </div>
          <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="m-cancel" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm cursor-pointer">ยกเลิก</button>
            <button type="submit" class="min-h-[40px] px-4 py-2 rounded-xl bg-zinc-900 text-white text-xs sm:text-sm font-semibold cursor-pointer">บันทึก</button>
          </div>
        </form>
      </div>
      `,
      (box) => {
        box.querySelector('#m-cancel').addEventListener('click', closeModal);
        box.querySelector('#modal-edit-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          try {
            await apiFetch(`/api/master/vehicles/${v.id}`, {
              method: 'PUT',
              body: JSON.stringify({ license_plate: box.querySelector('#m-plate').value.trim() })
            });
            closeModal();
            showToast('อัปเดตทะเบียนรถเรียบร้อยแล้ว', 'success');
            await refreshAll();
          } catch (err) {
            showToast(err.message, 'error');
          }
        });
      }
    );
  }

  function openEditWorkerModal(w) {
    const vehOpts = state.vehicles
      .map(
        (v) =>
          `<option value="${escapeHtml(v.license_plate)}" ${v.license_plate === w.default_vehicle ? 'selected' : ''}>${escapeHtml(v.license_plate)}</option>`
      )
      .join('');
    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <h3 class="text-base font-bold text-slate-900">แก้ไขข้อมูลคนขับ &amp; ค่าที่ผูกประจำ</h3>
        <form id="modal-edit-form" class="space-y-3.5">
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อคนขับ *</label>
              <input type="text" id="m-name" value="${escapeHtml(w.name)}" required
                class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ทะเบียนรถประจำตัว</label>
              <select id="m-veh" class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-600">
                <option value="">-- ยังไม่ผูกทะเบียนรถ --</option>
                ${vehOpts}
              </select>
            </div>
            <div class="sm:col-span-2">
              <label class="block text-xs font-semibold text-slate-700 mb-1">โครงการที่ผูกประจำ (ไม่บังคับ)</label>
              <input type="text" id="m-worker-site" list="dl-bo-sites" autocomplete="off" value="${escapeHtml(w.default_site || '')}" placeholder="เลือกหรือพิมพ์ชื่อโครงการ..."
                class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ต้นทาง / จาก ที่ผูก (ไม่บังคับ)</label>
              <input type="text" id="m-worker-origin" list="dl-bo-origins" autocomplete="off" value="${escapeHtml(w.default_origin || '')}" placeholder="เลือกหรือพิมพ์จุดรับ..."
                class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ปลายทาง / ถึง ที่ผูก (ไม่บังคับ)</label>
              <input type="text" id="m-worker-dest" list="dl-bo-destinations" autocomplete="off" value="${escapeHtml(w.default_destination || '')}" placeholder="เว้นว่าง = ใช้โครงการที่ผูก"
                class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
            </div>
          </div>
          <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="m-cancel" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm cursor-pointer">ยกเลิก</button>
            <button type="submit" class="min-h-[40px] px-4 py-2 rounded-xl bg-zinc-900 text-white text-xs sm:text-sm font-semibold cursor-pointer">บันทึก</button>
          </div>
        </form>
      </div>
      `,
      (box) => {
        box.querySelector('#m-cancel').addEventListener('click', closeModal);
        box.querySelector('#modal-edit-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          try {
            await apiFetch(`/api/master/workers/${w.id}`, {
              method: 'PUT',
              body: JSON.stringify({
                name: box.querySelector('#m-name').value.trim(),
                default_vehicle: box.querySelector('#m-veh').value.trim(),
                default_site: box.querySelector('#m-worker-site').value.trim(),
                default_origin: box.querySelector('#m-worker-origin').value.trim(),
                default_destination: box.querySelector('#m-worker-dest').value.trim()
              })
            });
            closeModal();
            showToast('อัปเดตข้อมูลคนขับเรียบร้อยแล้ว', 'success');
            await refreshAll();
          } catch (err) {
            showToast(err.message, 'error');
          }
        });
      }
    );
  }

  function openEditSiteModal(s) {
    let selectedOrigins = [...getSiteBoundOrigins(s)];
    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <h3 class="text-base font-bold text-slate-900">แก้ไขโครงการ &amp; ผูกเส้นทาง (ผูกได้หลายเส้นทาง)</h3>
        <form id="modal-edit-form" class="space-y-3.5">
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อโครงการ / ไซต์งาน (site_name) *</label>
            <input type="text" id="m-site-name" value="${escapeHtml(s.site_name)}" required
              class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อโซน (zone_name)</label>
            <input type="text" id="m-zone-name" value="${escapeHtml(s.zone_name)}"
              class="w-full min-h-[42px] px-3 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600" />
          </div>
          <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2.5">
            <label for="m-site-origin" class="block text-xs font-semibold text-slate-700">
              <i class="fa-solid fa-route text-emerald-600 mr-1"></i>เส้นทาง / จุดรับที่ผูกกับโครงการนี้ (เลือกได้หลายเส้นทาง)
            </label>
            <div class="flex items-center gap-2">
              <input type="text" id="m-site-origin" list="dl-bo-origins" autocomplete="off" placeholder="พิมพ์หรือเลือกจุดรับ/เส้นทาง แล้วกด + เพิ่ม..."
                class="flex-1 min-h-[38px] px-3 py-1.5 rounded-xl border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-emerald-600" />
              <button type="button" id="m-site-origin-add"
                class="min-h-[38px] px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold flex items-center gap-1 cursor-pointer shrink-0">
                <i class="fa-solid fa-plus"></i><span>เพิ่ม</span>
              </button>
            </div>
            <div id="m-site-origins-tags" class="flex flex-wrap items-center gap-1.5"></div>
            <div id="m-site-quick-routes" class="flex flex-wrap items-center gap-1.5 pt-1 border-t border-slate-200/70"></div>
          </div>
          <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="m-cancel" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm cursor-pointer">ยกเลิก</button>
            <button type="submit" class="min-h-[40px] px-4 py-2 rounded-xl bg-zinc-900 text-white text-xs sm:text-sm font-semibold cursor-pointer">บันทึก</button>
          </div>
        </form>
      </div>
      `,
      (box) => {
        const originInput = box.querySelector('#m-site-origin');
        const tagsEl = box.querySelector('#m-site-origins-tags');
        const quickEl = box.querySelector('#m-site-quick-routes');

        const redrawPicker = () => {
          renderOriginTagPicker(tagsEl, quickEl, selectedOrigins, {
            removeAttr: 'data-remove-m-origin',
            toggleAttr: 'data-toggle-m-origin'
          });
        };

        const addTypedOrigin = () => {
          const val = originInput.value.trim();
          if (!val) return;
          if (!selectedOrigins.includes(val)) selectedOrigins.push(val);
          originInput.value = '';
          redrawPicker();
        };

        redrawPicker();

        box.querySelector('#m-site-origin-add').addEventListener('click', addTypedOrigin);
        originInput.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' && originInput.value.trim()) {
            e.preventDefault();
            addTypedOrigin();
          }
        });
        tagsEl.addEventListener('click', (e) => {
          const btn = e.target.closest('button[data-remove-m-origin]');
          if (!btn) return;
          const val = btn.dataset.removeMOrigin || '';
          selectedOrigins = selectedOrigins.filter((x) => x !== val);
          redrawPicker();
        });
        quickEl.addEventListener('click', (e) => {
          const btn = e.target.closest('button[data-toggle-m-origin]');
          if (!btn) return;
          const val = btn.dataset.toggleMOrigin || '';
          if (!val) return;
          if (selectedOrigins.includes(val)) selectedOrigins = selectedOrigins.filter((x) => x !== val);
          else selectedOrigins.push(val);
          redrawPicker();
        });

        box.querySelector('#m-cancel').addEventListener('click', closeModal);
        box.querySelector('#modal-edit-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          try {
            const leftover = originInput.value.trim();
            const finalOrigins = [...new Set([...selectedOrigins, ...(leftover ? [leftover] : [])])];
            const newSiteName = box.querySelector('#m-site-name').value.trim();
            await apiFetch(`/api/master/sites/${s.id}`, {
              method: 'PUT',
              body: JSON.stringify({
                site_name: newSiteName,
                zone_name: box.querySelector('#m-zone-name').value.trim(),
                default_origin: finalOrigins[0] || '',
                default_origins: finalOrigins
              })
            });
            if (state.boundBillingSite && state.boundBillingSite === s.site_name && newSiteName) {
              state.boundBillingSite = newSiteName;
              localStorage.setItem(BOUND_BILLING_SITE_KEY, newSiteName);
              await updateBindingUI();
            }
            if (newSiteName && newSiteName !== s.site_name) {
              let migrated = false;
              if (state.boundZoneHandlesBySite[s.site_name]) {
                const movedZone = state.boundZoneHandlesBySite[s.site_name];
                state.boundZoneHandlesBySite[newSiteName] = movedZone;
                delete state.boundZoneHandlesBySite[s.site_name];
                await saveHandleToDb(siteHandleKey('zone', newSiteName), movedZone);
                await saveHandleToDb(siteHandleKey('zone', s.site_name), null);
                migrated = true;
              }
              if (state.boundBillingHandlesBySite[s.site_name]) {
                const movedHandle = state.boundBillingHandlesBySite[s.site_name];
                state.boundBillingHandlesBySite[newSiteName] = movedHandle;
                delete state.boundBillingHandlesBySite[s.site_name];
                await saveHandleToDb(siteHandleKey('billing', newSiteName), movedHandle);
                await saveHandleToDb(siteHandleKey('billing', s.site_name), null);
                migrated = true;
              }
              if (migrated) await updateBindingUI();
            }
            closeModal();
            showToast('อัปเดตไซต์งาน/โครงการเรียบร้อยแล้ว', 'success');
            await refreshAll();
          } catch (err) {
            showToast(err.message, 'error');
          }
        });
      }
    );
  }

  function openEditRouteModal(r) {
    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <h3 class="text-base font-bold text-slate-900">แก้ไขเส้นทาง / ที่หมาย</h3>
        <form id="modal-edit-form" class="space-y-3">
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ต้นทาง (จุดรับ) *</label>
              <input type="text" id="m-origin" value="${escapeHtml(r.origin)}" required
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ปลายทางที่ผูก (ไม่บังคับ)</label>
              <input type="text" id="m-dest" list="dl-bo-destinations" autocomplete="off" value="${escapeHtml(r.destination || '')}" placeholder="เว้นว่าง หรือเลือก/พิมพ์..."
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div class="sm:col-span-2">
              <label class="block text-xs font-semibold text-slate-700 mb-1">โครงการที่ผูกกับเส้นทางนี้ (ไม่บังคับ)</label>
              <input type="text" id="m-route-site" list="dl-bo-sites" autocomplete="off" value="${escapeHtml(r.default_site || '')}" placeholder="เว้นว่าง หรือเลือกโครงการ..."
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อแผ่นงาน (แท็บ) ในไฟล์ Excel</label>
              <input type="text" id="m-sheet" value="${escapeHtml(r.route_sheet_name)}"
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อรายการที่แสดงในใบเบิก</label>
              <input type="text" id="m-billing" value="${escapeHtml(r.billing_title)}"
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">เวลาออกปกติ</label>
              <input type="text" id="m-dep" value="${escapeHtml(r.default_dep_time)}" placeholder="เช่น 06:30"
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">เวลาถึงปกติ</label>
              <input type="text" id="m-arr" value="${escapeHtml(r.default_arr_time)}" placeholder="เช่น 07:15"
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ราคาต่อเที่ยว (บาท)</label>
              <input type="number" id="m-price" min="0" step="0.01" value="${escapeHtml(r.unit_price)}"
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ไฟล์สำเนาโซน: ช่องคอลัมน์ C ให้แสดง</label>
              <select id="m-mode" class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm bg-white">
                <option value="worker_name" ${r.display_mode_zone !== 'license_plate' ? 'selected' : ''}>ชื่อคนขับ</option>
                <option value="license_plate" ${r.display_mode_zone === 'license_plate' ? 'selected' : ''}>เลขทะเบียนรถ</option>
              </select>
            </div>
          </div>
          <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="m-cancel" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm cursor-pointer">ยกเลิก</button>
            <button type="submit" class="min-h-[40px] px-4 py-2 rounded-xl bg-zinc-900 text-white text-xs sm:text-sm font-semibold cursor-pointer">บันทึก</button>
          </div>
        </form>
      </div>
      `,
      (box) => {
        attachAutoColonTimeInput(box.querySelector('#m-dep'));
        attachAutoColonTimeInput(box.querySelector('#m-arr'));
        box.querySelector('#m-cancel').addEventListener('click', closeModal);
        box.querySelector('#modal-edit-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          try {
            await apiFetch(`/api/master/routes/${r.id}`, {
              method: 'PUT',
              body: JSON.stringify({
                origin: box.querySelector('#m-origin').value.trim(),
                destination: box.querySelector('#m-dest').value.trim(),
                default_site: box.querySelector('#m-route-site').value.trim(),
                route_sheet_name: box.querySelector('#m-sheet').value.trim(),
                billing_title: box.querySelector('#m-billing').value.trim(),
                default_dep_time: box.querySelector('#m-dep').value.trim(),
                default_arr_time: box.querySelector('#m-arr').value.trim(),
                unit_price: Number(box.querySelector('#m-price').value) || 0,
                display_mode_zone: box.querySelector('#m-mode').value
              })
            });
            closeModal();
            showToast('อัปเดตเส้นทางเรียบร้อยแล้ว', 'success');
            await loadMasterData();
          } catch (err) {
            showToast(err.message, 'error');
          }
        });
      }
    );
  }

  async function handleBackOfficeActionClick(e) {
    const btn = e.target.closest('button[data-bo]');
    if (!btn) return;
    const action = btn.dataset.bo;
    const id = Number(btn.dataset.id);

    try {
      if (action === 'toggle-vehicle') {
        const v = state.vehicles.find((x) => x.id === id);
        if (!v) return;
        await apiFetch(`/api/master/vehicles/${id}`, { method: 'PUT', body: JSON.stringify({ active: !v.active }) });
        await loadMasterData();
      } else if (action === 'edit-vehicle') {
        const v = state.vehicles.find((x) => x.id === id);
        if (v) openEditVehicleModal(v);
      } else if (action === 'delete-vehicle') {
        const v = state.vehicles.find((x) => x.id === id);
        if (!v) return;
        const ok = await confirmModal({
          title: 'ยืนยันการลบทะเบียนรถ',
          message: `ต้องการลบทะเบียนรถ "${v.license_plate}" ใช่หรือไม่? (คนขับที่ผูกทะเบียนนี้ไว้จะถูกปลดการผูกอัตโนมัติ)`,
          confirmText: 'ลบทะเบียนรถ',
          danger: true
        });
        if (!ok) return;
        await apiFetch(`/api/master/vehicles/${id}`, { method: 'DELETE' });
        showToast(`ลบทะเบียนรถ "${v.license_plate}" แล้ว`, 'success');
        await loadMasterData();
      } else if (action === 'toggle-worker') {
        const w = state.workers.find((x) => x.id === id);
        if (!w) return;
        await apiFetch(`/api/master/workers/${id}`, { method: 'PUT', body: JSON.stringify({ active: !w.active }) });
        await loadMasterData();
      } else if (action === 'edit-worker') {
        const w = state.workers.find((x) => x.id === id);
        if (w) openEditWorkerModal(w);
      } else if (action === 'delete-worker') {
        const w = state.workers.find((x) => x.id === id);
        if (!w) return;
        const ok = await confirmModal({
          title: 'ยืนยันการลบรายชื่อคนขับ',
          message: `ต้องการลบคนขับ "${w.name}" ออกจากรายการข้อมูลหลักใช่หรือไม่?`,
          confirmText: 'ลบรายชื่อ',
          danger: true
        });
        if (!ok) return;
        await apiFetch(`/api/master/workers/${id}`, { method: 'DELETE' });
        showToast(`ลบรายชื่อ "${w.name}" แล้ว`, 'success');
        await loadMasterData();
      } else if (action === 'toggle-site') {
        const s = state.sites.find((x) => x.id === id);
        if (!s) return;
        await apiFetch(`/api/master/sites/${id}`, { method: 'PUT', body: JSON.stringify({ active: !s.active }) });
        await loadMasterData();
      } else if (action === 'edit-site') {
        const s = state.sites.find((x) => x.id === id);
        if (s) openEditSiteModal(s);
      } else if (action === 'delete-site') {
        const s = state.sites.find((x) => x.id === id);
        if (!s) return;
        const ok = await confirmModal({
          title: 'ยืนยันการลบไซต์งาน / โครงการ',
          message: `ต้องการลบโครงการ "${s.site_name}" ออกจากรายการข้อมูลหลักใช่หรือไม่?`,
          confirmText: 'ลบโครงการ',
          danger: true
        });
        if (!ok) return;
        await apiFetch(`/api/master/sites/${id}`, { method: 'DELETE' });
        showToast(`ลบโครงการ "${s.site_name}" แล้ว`, 'success');
        await loadMasterData();
      } else if (action === 'toggle-route') {
        const r = state.routes.find((x) => x.id === id);
        if (!r) return;
        await apiFetch(`/api/master/routes/${id}`, { method: 'PUT', body: JSON.stringify({ active: !r.active }) });
        await loadMasterData();
      } else if (action === 'edit-route') {
        const r = state.routes.find((x) => x.id === id);
        if (r) openEditRouteModal(r);
      } else if (action === 'delete-route') {
        const r = state.routes.find((x) => x.id === id);
        if (!r) return;
        const label = r.destination ? `${r.origin} → ${r.destination}` : r.origin;
        const ok = await confirmModal({
          title: 'ยืนยันการลบเส้นทาง',
          message: `ต้องการลบเส้นทาง "${label}" ออกจากรายการข้อมูลหลักใช่หรือไม่?`,
          confirmText: 'ลบเส้นทาง',
          danger: true
        });
        if (!ok) return;
        await apiFetch(`/api/master/routes/${id}`, { method: 'DELETE' });
        showToast(`ลบเส้นทาง "${label}" แล้ว`, 'success');
        await loadMasterData();
      }
    } catch (err) {
      showToast(err.message || 'ไม่สามารถดำเนินการได้', 'error');
    }
  }

  async function handleQuotaSubmit(e) {
    e.preventDefault();
    const max_storage_limit = Number($('input-max-storage-limit').value);
    if (!Number.isFinite(max_storage_limit) || max_storage_limit < 1) {
      showToast('ขีดจำกัดรายการต้องเป็นตัวเลขตั้งแต่ 1 ขึ้นไป', 'error');
      return;
    }
    try {
      const res = await apiFetch('/api/settings', {
        method: 'PUT',
        body: JSON.stringify({ max_storage_limit })
      });
      if (res && res.storageQuota) renderQuotaUI(res.storageQuota);
      showToast(`บันทึกขีดจำกัดสูงสุดเป็น ${formatInt(max_storage_limit)} รายการแล้ว`, 'success');
    } catch (err) {
      showToast(err.message || 'ไม่สามารถบันทึกโควต้าได้', 'error');
    }
  }

  // ==========================================================================
  // 11. Event Binding & Initialization
  // ==========================================================================
  function bindEvents() {
    // Modal backdrop click & Escape
    $('custom-modal-backdrop').addEventListener('click', (e) => {
      if (e.target === $('custom-modal-backdrop')) closeModal();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeModal();
    });

    // Login events
    $('toggle-password-btn').addEventListener('click', () => {
      const input = $('login-password');
      const icon = $('toggle-password-icon');
      const show = input.type === 'password';
      input.type = show ? 'text' : 'password';
      icon.className = show ? 'fa-solid fa-eye-slash' : 'fa-solid fa-eye';
    });

    $('login-form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const errBox = $('login-error-box');
      errBox.classList.add('hidden');
      errBox.classList.remove('flex');

      const username = $('login-username').value.trim();
      const password = $('login-password').value;
      const btn = $('btn-login-submit');
      btn.disabled = true;

      try {
        const res = await apiFetch('/api/login', {
          method: 'POST',
          body: JSON.stringify({ username, password })
        });
        state.token = res.token;
        state.user = res.user;
        state.storageMode = res.storageMode || '';
        localStorage.setItem(TOKEN_KEY, state.token);
        $('login-password').value = '';
        showAppScreen();
        if (res.storageQuota) renderQuotaUI(res.storageQuota);
        await refreshAll();
        showToast(`ยินดีต้อนรับ ${state.user.full_name}`, 'success');
      } catch (err) {
        $('login-error-text').textContent = err.message || 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง';
        errBox.classList.remove('hidden');
        errBox.classList.add('flex');
      } finally {
        btn.disabled = false;
      }
    });

    $('btn-logout').addEventListener('click', async () => {
      const ok = await confirmModal({
        title: 'ออกจากระบบ',
        message: 'ต้องการออกจากระบบปฏิบัติงานใช่หรือไม่?',
        confirmText: 'ออกจากระบบ'
      });
      if (!ok) return;
      try {
        await apiFetch('/api/logout', { method: 'POST' });
      } catch {
        // ignore network errors on logout
      }
      handleForceLogout('ออกจากระบบเรียบร้อยแล้ว');
    });

    // Navigation
    $('nav-tab-front').addEventListener('click', () => switchView('front'));
    $('nav-tab-back').addEventListener('click', () => switchView('back'));
    $('mobile-nav-front').addEventListener('click', () => switchView('front'));
    $('mobile-nav-back').addEventListener('click', () => switchView('back'));
    $('btn-goto-backoffice-hint').addEventListener('click', () => switchView('back'));

    // แถบ 3 ขั้นตอน + ปุ่ม "ถัดไป" ในหน้าตั้งค่า
    $('step-1').addEventListener('click', () => {
      if (state.user && state.user.role === 'admin') {
        switchView('back');
        window.scrollTo({ top: 0, behavior: 'smooth' });
      } else {
        showToast('การตั้งค่ารายชื่อ รถ และเส้นทาง ทำได้เฉพาะแอดมิน — กรุณาแจ้งแอดมิน', 'warning');
      }
    });
    $('step-2').addEventListener('click', () => {
      scrollToEl('form-card-container');
      setTimeout(() => $('form-date') && $('form-date').focus({ preventScroll: true }), 400);
    });
    $('step-3').addEventListener('click', () => scrollToEl('logs-section'));
    document.querySelectorAll('[data-goto-bo]').forEach((btn) => {
      btn.addEventListener('click', () => {
        switchBoTab(btn.getAttribute('data-goto-bo'));
        window.scrollTo({ top: 0, behavior: 'smooth' });
      });
    });
    document.querySelectorAll('[data-goto-view]').forEach((btn) => {
      btn.addEventListener('click', () => {
        switchView(btn.getAttribute('data-goto-view'));
        window.scrollTo({ top: 0, behavior: 'smooth' });
      });
    });

    // Back-Office Sub-Tabs
    $('bo-tab-vehicles').addEventListener('click', () => switchBoTab('vehicles'));
    $('bo-tab-workers').addEventListener('click', () => switchBoTab('workers'));
    $('bo-tab-sites').addEventListener('click', () => switchBoTab('sites'));
    $('bo-tab-routes').addEventListener('click', () => switchBoTab('routes'));

    // Front-Office Form
    $('trip-type-normal').addEventListener('click', () => setTripType('normal'));
    $('trip-type-charter').addEventListener('click', () => setTripType('charter'));
    $('form-date').addEventListener('change', updateFormDatePreview);
    // Searchable Comboboxes (พิมพ์ได้ + ดรอปดาวน์ ▾)
    ['form-worker', 'form-vehicle', 'form-site', 'form-origin', 'form-destination'].forEach(setupCombobox);
    document.addEventListener('mousedown', (e) => {
      if (!e.target.closest('.combo-wrap')) closeAllComboboxes();
    });

    $('form-worker').addEventListener('input', onWorkerSelectChange);
    $('form-worker').addEventListener('change', onWorkerSelectChange);
    $('form-worker-vehicle-hint').addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-apply-all-worker]');
      if (!btn) return;
      const rec = getFieldRecommendations();
      if (rec.recPlate) {
        $('form-vehicle').value = rec.recPlate;
        pulseElement($('form-vehicle'));
      }
      if (rec.recSite) {
        $('form-site').value = rec.recSite;
        pulseElement($('form-site'));
      }
      if (rec.recOrigin) {
        $('form-origin').value = rec.recOrigin;
        pulseElement($('form-origin'));
      }
      if (rec.recDest) {
        $('form-destination').value = rec.recDest;
        pulseElement($('form-destination'));
      }
      applyRoutePresetFromInput();
    });
    $('form-vehicle').addEventListener('input', onVehicleSelectChange);
    $('form-vehicle').addEventListener('change', onVehicleSelectChange);
    $('form-vehicle-worker-hint').addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-use-rec-plate]');
      if (!btn) return;
      const plate = btn.dataset.useRecPlate || '';
      const vehEl = $('form-vehicle');
      vehEl.value = plate;
      pulseElement(vehEl);
      updateWorkerVehicleHints();
    });
    $('form-site').addEventListener('input', onSiteInputChange);
    $('form-site').addEventListener('change', onSiteInputChange);
    $('form-site-rec-hint').addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-use-rec-site]');
      if (!btn) return;
      const siteVal = btn.dataset.useRecSite || '';
      const siteEl = $('form-site');
      siteEl.value = siteVal;
      pulseElement(siteEl);
      onSiteInputChange();
    });
    $('form-route-rec-hint').addEventListener('click', (e) => {
      const bothBtn = e.target.closest('button[data-use-rec-route]');
      if (bothBtn) {
        const o = bothBtn.dataset.recOrig || '';
        const d = bothBtn.dataset.recDest || '';
        if (o) {
          $('form-origin').value = o;
          pulseElement($('form-origin'));
        }
        if (d) {
          $('form-destination').value = d;
          pulseElement($('form-destination'));
        }
        applyRoutePresetFromInput();
        return;
      }
      const origBtn = e.target.closest('button[data-use-rec-origin]');
      if (origBtn) {
        $('form-origin').value = origBtn.dataset.useRecOrigin || '';
        pulseElement($('form-origin'));
        applyRoutePresetFromInput();
        return;
      }
      const destBtn = e.target.closest('button[data-use-rec-dest]');
      if (destBtn) {
        $('form-destination').value = destBtn.dataset.useRecDest || '';
        pulseElement($('form-destination'));
        applyRoutePresetFromInput();
      }
    });
    $('btn-swap-route').addEventListener('click', onSwapRouteClick);
    $('form-auto-return').addEventListener('change', updateAutoReturnUI);
    $('btn-toggle-route-meta').addEventListener('click', () => toggleRouteMetaBox());
    $('form-origin').addEventListener('input', applyRoutePresetFromInput);
    $('form-origin').addEventListener('change', applyRoutePresetFromInput);
    $('form-destination').addEventListener('input', applyRoutePresetFromInput);
    $('form-destination').addEventListener('change', applyRoutePresetFromInput);
    $('form-route-sheet-name').addEventListener('input', updateRouteMetaPreview);
    $('form-billing-title').addEventListener('input', updateRouteMetaPreview);
    $('form-display-mode-zone').addEventListener('change', updateRouteMetaPreview);
    attachAutoColonTimeInput($('form-dep-time'), updateDurationPreview);
    attachAutoColonTimeInput($('form-arr-time'), updateDurationPreview);
    attachAutoColonTimeInput($('input-route-dep-time'));
    attachAutoColonTimeInput($('input-route-arr-time'));

    $('travel-log-form').addEventListener('submit', handleLogFormSubmit);
    $('btn-clear-form').addEventListener('click', () => resetForm({ keepDateAndSite: false }));
    $('btn-cancel-edit').addEventListener('click', () => resetForm({ keepDateAndSite: true }));
    $('btn-cancel-edit-top').addEventListener('click', () => resetForm({ keepDateAndSite: true }));
    $('btn-copy-previous-day').addEventListener('click', openCopyPreviousDayModal);

    // Filter & Quick Periods
    $('filter-site').addEventListener('change', loadLogs);
    $('filter-trip-type').addEventListener('change', loadLogs);
    $('filter-date-from').addEventListener('change', loadLogs);
    $('filter-date-to').addEventListener('change', loadLogs);
    $('filter-search').addEventListener('input', () => {
      clearTimeout(state.searchTimer);
      state.searchTimer = setTimeout(loadLogs, 220);
    });
    $('btn-period-1-15').addEventListener('click', () => applyQuickPeriod('1-15'));
    $('btn-period-16-end').addEventListener('click', () => applyQuickPeriod('16-end'));
    $('btn-period-today').addEventListener('click', () => applyQuickPeriod('today'));
    $('btn-period-all').addEventListener('click', () => applyQuickPeriod('all'));

    // Exports, File Binding & Logs Actions
    $('btn-export-zone').addEventListener('click', handleExportZoneCopy);
    $('btn-export-billing').addEventListener('click', handleExportBilling);
    $('btn-bind-zone').addEventListener('click', () => openBindFileModal('zone'));
    $('btn-unbind-zone').addEventListener('click', () => unbindFile('zone'));
    $('btn-sync-zone').addEventListener('click', () => syncToBoundFile('zone', { interactive: true }));
    $('btn-bind-billing').addEventListener('click', () => openBindFileModal('billing'));
    $('btn-unbind-billing').addEventListener('click', () => unbindFile('billing'));
    $('btn-sync-billing').addEventListener('click', () => syncToBoundFile('billing', { interactive: true }));
    $('btn-sync-all-excel-zone-sites').addEventListener('click', () => syncAllBoundExcelSites('zone'));
    $('btn-sync-all-excel-sites').addEventListener('click', () => syncAllBoundExcelSites('billing'));
    $('excel-sites-search').addEventListener('input', () => renderPerSiteExcelBindingList());
    $('btn-excel-sites-toggle').addEventListener('click', () => {
      state.excelSitesShowAll = !state.excelSitesShowAll;
      renderPerSiteExcelBindingList();
    });
    $('excel-per-site-list').addEventListener('click', async (e) => {
      const zBindBtn = e.target.closest('button[data-excel-zone-site-bind]');
      if (zBindBtn) {
        openBindFileModal('zone', zBindBtn.dataset.excelZoneSiteBind || '');
        return;
      }
      const zSyncBtn = e.target.closest('button[data-excel-zone-site-sync]');
      if (zSyncBtn) {
        await syncSiteExcelFile('zone', zSyncBtn.dataset.excelZoneSiteSync || '', { interactive: true, triggerBtn: zSyncBtn });
        return;
      }
      const zDlBtn = e.target.closest('button[data-excel-zone-site-dl]');
      if (zDlBtn) {
        const sName = zDlBtn.dataset.excelZoneSiteDl || '';
        zDlBtn.disabled = true;
        try {
          const f = getFilterParams();
          await runZoneExport({ siteName: sName, dateFrom: f.date_from, dateTo: f.date_to });
        } catch (err) {
          showToast(`ส่งออกสำเนา "${sName}" ไม่สำเร็จ: ${err.message}`, 'error', 5000);
        } finally {
          zDlBtn.disabled = false;
        }
        return;
      }
      const zUnbindBtn = e.target.closest('button[data-excel-zone-site-unbind]');
      if (zUnbindBtn) {
        await unbindSiteExcelFile('zone', zUnbindBtn.dataset.excelZoneSiteUnbind || '');
        return;
      }

      const bindBtn = e.target.closest('button[data-excel-site-bind]');
      if (bindBtn) {
        openBindFileModal('billing', bindBtn.dataset.excelSiteBind || '');
        return;
      }
      const syncBtn = e.target.closest('button[data-excel-site-sync]');
      if (syncBtn) {
        await syncSiteExcelFile('billing', syncBtn.dataset.excelSiteSync || '', { interactive: true, triggerBtn: syncBtn });
        return;
      }
      const dlBtn = e.target.closest('button[data-excel-site-dl]');
      if (dlBtn) {
        const sName = dlBtn.dataset.excelSiteDl || '';
        dlBtn.disabled = true;
        try {
          const f = getFilterParams();
          await runBillingExport({ siteName: sName, dateFrom: f.date_from, dateTo: f.date_to, issueDate: todayIso() });
        } catch (err) {
          showToast(`ส่งออกใบเบิก "${sName}" ไม่สำเร็จ: ${err.message}`, 'error', 5000);
        } finally {
          dlBtn.disabled = false;
        }
        return;
      }
      const unbindBtn = e.target.closest('button[data-excel-site-unbind]');
      if (unbindBtn) {
        await unbindSiteExcelFile('billing', unbindBtn.dataset.excelSiteUnbind || '');
      }
    });
    $('toggle-auto-sync').addEventListener('change', (e) => {
      state.autoSync = Boolean(e.target.checked);
      localStorage.setItem(AUTO_SYNC_KEY, state.autoSync ? 'true' : 'false');
      showToast(
        state.autoSync ? 'เปิดซิงค์ลงไฟล์ที่ผูกไว้อัตโนมัติแล้ว' : 'ปิดซิงค์อัตโนมัติแล้ว (กดปุ่มบันทึกไฟล์เองได้ตลอด)',
        'info'
      );
    });

    // Google Sheets Online Sync
    $('btn-gsheet-guide').addEventListener('click', () => openGoogleSheetBindModal('zone'));
    $('btn-gsheet-history').addEventListener('click', openGoogleSheetsHistoryModal);
    $('gsheet-last-sync-note').addEventListener('click', openGoogleSheetsHistoryModal);
    $('btn-gsheet-save-settings').addEventListener('click', saveGoogleSheetsBarSettings);
    $('btn-gsheet-bind-zone').addEventListener('click', () => openGoogleSheetBindModal('zone'));
    $('btn-gsheet-sync-zone').addEventListener('click', () => syncToGoogleSheet('zone', { interactive: true, mode: 'manual' }));
    $('btn-gsheet-bind-billing').addEventListener('click', () => openGoogleSheetBindModal('billing'));
    $('btn-gsheet-sync-billing').addEventListener('click', () => syncToGoogleSheet('billing', { interactive: true, mode: 'manual' }));
    $('btn-sync-all-gsheet-zone-sites').addEventListener('click', () => syncAllBoundGoogleSheetSites('zone'));
    $('btn-sync-all-gsheet-sites').addEventListener('click', () => syncAllBoundGoogleSheetSites('billing'));
    $('gsheet-sites-search').addEventListener('input', () => renderPerSiteGoogleSheetsList());
    $('btn-gsheet-sites-toggle').addEventListener('click', () => {
      state.gsheetSitesShowAll = !state.gsheetSitesShowAll;
      renderPerSiteGoogleSheetsList();
    });
    $('gsheet-per-site-list').addEventListener('click', async (e) => {
      const zBindBtn = e.target.closest('button[data-gsheet-zone-site-bind]');
      if (zBindBtn) {
        openGoogleSheetBindModal('zone', zBindBtn.dataset.gsheetZoneSiteBind || '');
        return;
      }
      const zSyncBtn = e.target.closest('button[data-gsheet-zone-site-sync]');
      if (zSyncBtn) {
        await syncToGoogleSheet('zone', {
          interactive: true,
          mode: 'manual',
          siteName: zSyncBtn.dataset.gsheetZoneSiteSync || '',
          triggerBtn: zSyncBtn
        });
        return;
      }
      const bindBtn = e.target.closest('button[data-gsheet-site-bind]');
      if (bindBtn) {
        openGoogleSheetBindModal('billing', bindBtn.dataset.gsheetSiteBind || '');
        return;
      }
      const syncBtn = e.target.closest('button[data-gsheet-site-sync]');
      if (syncBtn) {
        await syncToGoogleSheet('billing', {
          interactive: true,
          mode: 'manual',
          siteName: syncBtn.dataset.gsheetSiteSync || '',
          triggerBtn: syncBtn
        });
      }
    });
    $('toggle-gsheet-auto-sync').addEventListener('change', async (e) => {
      const val = Boolean(e.target.checked);
      try {
        const res = await apiFetch('/api/gsheet/config', {
          method: 'PUT',
          body: JSON.stringify({ auto_sync: val })
        });
        state.gsheet = res.config;
        renderGoogleSheetsUI();
        showToast(
          val ? 'เปิดไหลข้อมูลลง Google Sheets อัตโนมัติแล้ว' : 'ปิดไหลข้อมูลลง Google Sheets อัตโนมัติแล้ว (กดปุ่มไหลลงชีตเองได้ตลอด)',
          'info'
        );
      } catch (err) {
        e.target.checked = !val;
        showToast(err.message || 'ไม่สามารถบันทึกการตั้งค่าได้', 'error');
      }
    });
    $('btn-clear-all-logs').addEventListener('click', handleClearAllLogs);
    $('logs-table-body').addEventListener('click', handleLogsActionClick);
    $('logs-mobile-cards').addEventListener('click', handleLogsActionClick);

    // Back-Office Forms & Action Delegation
    $('form-add-vehicle').addEventListener('submit', handleAddVehicle);
    $('form-add-worker').addEventListener('submit', handleAddWorker);
    $('form-add-site').addEventListener('submit', handleAddSite);
    const addNewSiteOriginTag = () => {
      const inp = $('input-site-default-origin');
      if (!inp) return;
      const val = inp.value.trim();
      if (!val) return;
      if (!state.newSiteOrigins.includes(val)) state.newSiteOrigins.push(val);
      inp.value = '';
      renderAddSiteOriginPicker();
    };
    $('btn-add-site-origin-tag').addEventListener('click', addNewSiteOriginTag);
    $('input-site-default-origin').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && $('input-site-default-origin').value.trim()) {
        e.preventDefault();
        addNewSiteOriginTag();
      }
    });
    $('site-add-origins-tags').addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-remove-new-site-origin]');
      if (!btn) return;
      const val = btn.dataset.removeNewSiteOrigin || '';
      state.newSiteOrigins = state.newSiteOrigins.filter((x) => x !== val);
      renderAddSiteOriginPicker();
    });
    $('site-add-quick-routes').addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-toggle-new-site-origin]');
      if (!btn) return;
      const val = btn.dataset.toggleNewSiteOrigin || '';
      if (!val) return;
      if (state.newSiteOrigins.includes(val)) {
        state.newSiteOrigins = state.newSiteOrigins.filter((x) => x !== val);
      } else {
        state.newSiteOrigins.push(val);
      }
      renderAddSiteOriginPicker();
    });
    $('form-add-route').addEventListener('submit', handleAddRoute);
    $('form-settings-quota').addEventListener('submit', handleQuotaSubmit);
    $('view-back-office').addEventListener('click', handleBackOfficeActionClick);
  }

  async function init() {
    bindEvents();
    $('form-date').value = todayIso();
    $('filter-month').value = currentMonthIso();
    updateFormDatePreview();
    updateDurationPreview();
    updateAutoReturnUI();
    updateRouteMetaPreview();
    await restoreBoundHandles();

    if (!state.token) {
      showLoginScreen();
      return;
    }

    try {
      const me = await apiFetch('/api/me');
      state.user = me.user;
      showAppScreen();
      if (me.storageQuota) renderQuotaUI(me.storageQuota);
      await refreshAll();
    } catch {
      handleForceLogout();
    }
  }

  document.addEventListener('DOMContentLoaded', init);
})();
