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

    // Bound local Excel file handles (File System Access API + IndexedDB)
    boundZoneHandle: null,
    boundBillingHandle: null,
    boundBillingSite: localStorage.getItem(BOUND_BILLING_SITE_KEY) || '',
    autoSync: localStorage.getItem(AUTO_SYNC_KEY) !== 'false'
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
      const m = String(t || '').match(/^(\d{1,2}):(\d{2})/);
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
  }

  async function refreshAll() {
    await loadMasterData();
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

  function populateFrontDropdowns() {
    const activeWorkers = state.workers.filter((w) => w.active);
    const activeVehicles = state.vehicles.filter((v) => v.active);
    const activeSites = state.sites.filter((s) => s.active);
    const activeRoutes = state.routes.filter((r) => r.active);

    populateSelect(
      $('form-worker'),
      activeWorkers.map((w) => ({
        value: w.name,
        label: w.default_vehicle ? `${w.name} (${w.default_vehicle})` : w.name
      })),
      '-- เลือกชื่อคนขับ --'
    );

    populateSelect(
      $('form-vehicle'),
      activeVehicles.map((v) => ({ value: v.license_plate, label: v.license_plate })),
      '-- เลือกทะเบียนรถ --'
    );

    populateSelect(
      $('form-site'),
      activeSites.map((s) => ({
        value: s.site_name,
        label: s.zone_name ? `${s.site_name} [โซน: ${s.zone_name}]` : s.site_name
      })),
      '-- เลือกไซต์งาน / โครงการ --'
    );

    populateSelect(
      $('form-route-select'),
      activeRoutes.map((r) => ({
        value: String(r.id),
        label: `${r.origin} → ${r.destination}${r.unit_price ? ` (${formatMoney(r.unit_price)} บ.)` : ''}`
      })),
      '-- เลือกเส้นทาง (ระบบเติมข้อมูลให้) --'
    );

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
    updateRouteMetaPreview();
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
      el.innerHTML = '<i class="fa-solid fa-stopwatch text-slate-400"></i><span>คำนวณระยะเวลาอัตโนมัติ</span>';
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
  }

  function toggleRouteMetaBox(forceOpen) {
    state.routeMetaExpanded = typeof forceOpen === 'boolean' ? forceOpen : !state.routeMetaExpanded;
    const box = $('route-meta-advanced-box');
    const btn = $('btn-toggle-route-meta');
    box.classList.toggle('hidden', !state.routeMetaExpanded);
    btn.textContent = state.routeMetaExpanded ? 'ซ่อนการตั้งค่าเพิ่มเติม' : 'ตั้งค่าเพิ่มเติม (ไม่บังคับ)';
  }

  function onWorkerSelectChange() {
    const workerName = $('form-worker').value;
    const worker = state.workers.find((w) => w.name === workerName);
    const hint = $('form-worker-vehicle-hint');
    if (worker && worker.default_vehicle) {
      const vehSelect = $('form-vehicle');
      if (![...vehSelect.options].some((o) => o.value === worker.default_vehicle)) {
        const opt = document.createElement('option');
        opt.value = worker.default_vehicle;
        opt.textContent = worker.default_vehicle;
        vehSelect.appendChild(opt);
      }
      vehSelect.value = worker.default_vehicle;
      pulseElement(vehSelect);
      hint.innerHTML = `<span class="text-emerald-700 font-semibold"><i class="fa-solid fa-wand-magic-sparkles mr-1"></i>เติมรถประจำตัวให้แล้ว: ${escapeHtml(worker.default_vehicle)}</span>`;
    } else {
      hint.textContent = 'เลือกคนขับแล้วระบบเติมทะเบียนรถประจำตัวให้';
    }
  }

  function onRouteSelectChange() {
    const id = Number($('form-route-select').value);
    const route = state.routes.find((r) => r.id === id);
    if (!route) {
      updateRouteMetaPreview();
      return;
    }

    $('form-origin').value = route.origin || '';
    $('form-destination').value = route.destination || '';
    $('form-dep-time').value = route.default_dep_time || '';
    $('form-arr-time').value = route.default_arr_time || '';
    $('form-unit-price').value = route.unit_price ? String(route.unit_price) : '';
    $('form-route-sheet-name').value = route.route_sheet_name || `${route.origin} - ${route.destination}`;
    $('form-billing-title').value = route.billing_title || route.route_sheet_name || `${route.origin} - ${route.destination}`;
    $('form-display-mode-zone').value = route.display_mode_zone || 'worker_name';

    state.isSwapped = false;
    $('swap-status-badge').classList.add('hidden');
    $('swap-status-badge').classList.remove('inline-flex');

    pulseElement($('form-origin'));
    pulseElement($('form-destination'));
    pulseElement($('form-unit-price'));
    updateDurationPreview();
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

    pulseElement(origEl);
    pulseElement(destEl);
    updateRouteMetaPreview();
  }

  function resetForm({ keepDateAndSite = false } = {}) {
    const prevDate = $('form-date').value || todayIso();
    const prevSite = $('form-site').value;

    state.editingLogId = null;
    state.isSwapped = false;

    $('travel-log-form').reset();
    setTripType('normal');
    $('form-date').value = keepDateAndSite ? prevDate : todayIso();
    if (keepDateAndSite && prevSite) $('form-site').value = prevSite;
    $('form-route-sheet-name').value = '';
    $('form-billing-title').value = '';
    $('form-display-mode-zone').value = 'worker_name';
    $('form-worker-vehicle-hint').textContent = 'เลือกคนขับแล้วระบบเติมทะเบียนรถประจำตัวให้';

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
    updateRouteMetaPreview();
  }

  function startEditLog(log) {
    state.editingLogId = log.id;
    state.isSwapped = false;

    setTripType(log.trip_type || 'normal');
    $('form-date').value = log.trip_date || todayIso();

    // Ensure select contains historical values if archived/deleted
    const ensureOpt = (selectId, val) => {
      const sel = $(selectId);
      if (!sel || !val) return;
      if (![...sel.options].some((o) => o.value === val)) {
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        sel.appendChild(opt);
      }
      sel.value = val;
    };

    ensureOpt('form-worker', log.worker_name);
    ensureOpt('form-vehicle', log.license_plate);
    ensureOpt('form-site', log.site_name);
    $('form-route-select').value = '';
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

    updateFormDatePreview();
    updateDurationPreview();
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
        const res = await apiFetch('/api/logs', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (res && res.storageQuota) renderQuotaUI(res.storageQuota);
        showToast(`บันทึกเที่ยววิ่ง ${worker_name} (${origin} → ${destination}) เรียบร้อยแล้ว`, 'success');
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

  async function restoreBoundHandles() {
    if (!supportsFileBinding()) {
      const note = $('fs-api-unsupported-note');
      if (note) note.classList.remove('hidden');
      updateBindingUI();
      return;
    }
    const [zoneHandle, billingHandle] = await Promise.all([
      loadHandleFromDb('zone'),
      loadHandleFromDb('billing')
    ]);
    state.boundZoneHandle = zoneHandle;
    state.boundBillingHandle = billingHandle;
    await updateBindingUI();
  }

  async function updateBindingUI() {
    const autoCheck = $('toggle-auto-sync');
    if (autoCheck) autoCheck.checked = Boolean(state.autoSync);

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
      updateOne('billing', state.boundBillingHandle, 'bind-billing-label', 'bind-billing-dot', 'btn-unbind-billing')
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

  function showFileLockedModal({ kind, handle, wb, suggestedName, onRetry }) {
    const typeLabel = kind === 'billing' ? 'ใบเบิกโครงการ' : 'สำเนาโซน';
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
            if (kind === 'zone') state.boundZoneHandle = newHandle;
            else state.boundBillingHandle = newHandle;
            await saveHandleToDb(kind, newHandle);
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

  async function writeWorkbookToHandle(kind, handle, wb, suggestedName, { interactive = true } = {}) {
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
          handle,
          wb,
          suggestedName,
          onRetry: () => syncToBoundFile(kind, { interactive: true })
        });
        return { status: 'locked' };
      }
      throw err;
    }
  }

  function openBindFileModal(kind) {
    const isBilling = kind === 'billing';
    const title = isBilling
      ? 'ผูกไฟล์ที่ 2: ใบเบิกโครงการ (.xlsx)'
      : 'ผูกไฟล์ที่ 1: สำเนาโซน (.xlsx)';
    const desc = isBilling
      ? 'เลือกไฟล์ใบเบิกเดิมบนเครื่อง หรือสร้างไฟล์ใหม่ไว้ผูกกับระบบ เพื่อให้กดบันทึกข้อมูล (พร้อมโลโก้และสูตรคำนวณ) ลงไฟล์ได้ทันที'
      : 'เลือกไฟล์สำเนาโซนเดิมบนเครื่อง หรือสร้างไฟล์ใหม่ไว้ผูกกับระบบ เพื่อให้กดบันทึกข้อมูล (พร้อมสรุปเที่ยว COUNTIF/SUM) ลงไฟล์ได้ทันที';

    if (!supportsFileBinding()) {
      openModal(
        `
        <div class="p-5 sm:p-6 space-y-4">
          <h3 class="text-base font-bold text-slate-900">${escapeHtml(title)}</h3>
          <p class="text-xs sm:text-sm text-slate-600 leading-relaxed">
            เบราว์เซอร์นี้ยังไม่รองรับการผูกและเขียนทับไฟล์บนเครื่องโดยตรง (File System Access API) กรุณาใช้งานผ่าน <strong>Google Chrome</strong> หรือ <strong>Microsoft Edge</strong> บนคอมพิวเตอร์ หรือกดปุ่ม <strong>ส่งออกไฟล์ (.xlsx)</strong> เพื่อดาวน์โหลดไฟล์ตามปกติครับ
          </p>
          <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button type="button" id="m-bind-close" class="min-h-[40px] px-4 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm cursor-pointer">ปิด</button>
            <button type="button" id="m-bind-dl" class="min-h-[40px] px-4 py-2 rounded-xl bg-zinc-900 text-white text-xs sm:text-sm font-semibold cursor-pointer">ดาวน์โหลดไฟล์ (.xlsx) แทน</button>
          </div>
        </div>
        `,
        (box) => {
          box.querySelector('#m-bind-close').addEventListener('click', closeModal);
          box.querySelector('#m-bind-dl').addEventListener('click', () => {
            closeModal();
            if (isBilling) handleExportBilling();
            else handleExportZoneCopy();
          });
        }
      );
      return;
    }

    const siteSet = new Set(state.sites.map((s) => s.site_name).filter(Boolean));
    for (const l of state.logs) if (l.site_name) siteSet.add(l.site_name);
    const siteOptions = [...siteSet];
    const currentSite = state.boundBillingSite || $('filter-site').value.trim();

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
          isBilling
            ? `
          <div>
            <label for="m-bind-billing-site" class="block text-xs font-semibold text-slate-700 mb-1">
              ผูกกับโครงการ (ไม่บังคับ — หากเว้นว่างจะใช้ตามตัวกรองโครงการบนหน้าเว็บ)
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
          if (!isBilling) return '';
          const sel = box.querySelector('#m-bind-billing-site');
          const val = sel ? sel.value.trim() : '';
          state.boundBillingSite = val;
          localStorage.setItem(BOUND_BILLING_SITE_KEY, val);
          return val;
        };

        box.querySelector('#m-bind-cancel').addEventListener('click', closeModal);

        box.querySelector('#m-pick-existing').addEventListener('click', async () => {
          readChosenSite();
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
            if (isBilling) state.boundBillingHandle = handle;
            else state.boundZoneHandle = handle;
            await saveHandleToDb(kind, handle);
            await updateBindingUI();
            closeModal();
            if (perm !== 'granted') {
              showToast(
                `ผูกไฟล์ "${handle.name}" แล้ว แต่ยังไม่ได้อนุญาตสิทธิ์เขียนไฟล์ — กดปุ่ม "${isBilling ? 'บันทึกไฟล์ใบเบิก' : 'บันทึกไฟล์สำเนา'}" เพื่อขอสิทธิ์อีกครั้ง`,
                'warning',
                4800
              );
            } else {
              showToast(
                `ผูกไฟล์ "${handle.name}" เรียบร้อยแล้ว — กดปุ่ม "${isBilling ? 'บันทึกไฟล์ใบเบิก' : 'บันทึกไฟล์สำเนา'}" เพื่อเขียนข้อมูลลงไฟล์ได้ทันที`,
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
            const targetSite = isBilling ? chosenSite || f.site_name : f.site_name;
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
              : await buildZonePayloadFromFilters();

            const buffer = await payload.wb.xlsx.writeBuffer();
            const writable = await handle.createWritable();
            await writable.write(buffer);
            await writable.close();

            if (isBilling) state.boundBillingHandle = handle;
            else state.boundZoneHandle = handle;
            await saveHandleToDb(kind, handle);
            await updateBindingUI();
            closeModal();
            showToast(
              `สร้างและผูกไฟล์ "${handle.name}" สำเร็จ (${payload.wb.worksheets.length} ชีต, ${payload.logs.length} เที่ยว)`,
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
    if (!state.autoSync || !supportsFileBinding()) return;
    if (!state.boundZoneHandle && !state.boundBillingHandle) return;

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

    if (syncedNames.length > 0) {
      showToast(`ซิงค์ลงไฟล์ที่ผูกไว้อัตโนมัติแล้ว: ${syncedNames.join(', ')}`, 'info', 3500);
    } else if (needsPerm) {
      showToast('กดปุ่ม "บันทึกไฟล์สำเนา / ใบเบิก" 1 ครั้งเพื่ออนุญาตสิทธิ์เขียนไฟล์อัตโนมัติหลังรีเฟรชหน้า', 'warning', 4500);
    }
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
      el.innerHTML = '<div class="py-10 text-center text-xs text-slate-400">ยังไม่มีคนขับ — พิมพ์ชื่อในช่องด้านบน เลือกรถประจำตัว แล้วกด "เพิ่มคนขับ"</div>';
      return;
    }
    el.innerHTML = state.workers
      .map(
        (w, i) => `
        <div class="px-4 py-3 flex flex-wrap items-center justify-between gap-3 ${w.active ? 'bg-white' : 'bg-slate-50 opacity-65'}">
          <div class="flex flex-wrap items-center gap-2.5">
            <span class="text-xs font-bold text-slate-400 w-6">${i + 1}.</span>
            <span class="font-semibold text-sm text-slate-900">${escapeHtml(w.name)}</span>
            ${
              w.default_vehicle
                ? `<span class="license-plate-badge !text-[11px] !py-0.5 !px-2"><i class="fa-solid fa-truck-front text-amber-600 text-[10px]"></i>${escapeHtml(w.default_vehicle)}</span>`
                : '<span class="text-[11px] text-slate-400">ยังไม่ผูกทะเบียนรถประจำตัว</span>'
            }
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
      `
      )
      .join('');
  }

  function renderSitesList() {
    const el = $('list-sites');
    if (state.sites.length === 0) {
      el.innerHTML = '<div class="py-10 text-center text-xs text-slate-400">ยังไม่มีโครงการ — พิมพ์ชื่อโครงการในช่องด้านบน แล้วกด "เพิ่มโครงการ"</div>';
      return;
    }
    el.innerHTML = state.sites
      .map(
        (s, i) => `
        <div class="px-4 py-3 flex flex-wrap items-center justify-between gap-3 ${s.active ? 'bg-white' : 'bg-slate-50 opacity-65'}">
          <div class="flex flex-wrap items-center gap-2.5">
            <span class="text-xs font-bold text-slate-400 w-6">${i + 1}.</span>
            <span class="font-semibold text-sm text-slate-900">${escapeHtml(s.site_name)}</span>
            ${
              s.zone_name
                ? `<span class="px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-900 border border-amber-200 text-[11px] font-semibold">โซน: ${escapeHtml(s.zone_name)}</span>`
                : '<span class="text-[11px] text-slate-400">ไม่ระบุชื่อโซน</span>'
            }
            ${s.active ? '' : '<span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-200 text-slate-600 font-semibold">ปิดใช้งาน</span>'}
          </div>
          <div class="flex items-center gap-1.5">
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
      `
      )
      .join('');
  }

  function renderRoutesList() {
    const el = $('list-routes');
    if (state.routes.length === 0) {
      el.innerHTML = '<div class="py-10 text-center text-xs text-slate-400">ยังไม่มีเส้นทาง — กรอกต้นทาง ปลายทาง เวลา และราคาด้านบน แล้วกด "เพิ่มเส้นทาง"</div>';
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
                ${escapeHtml(r.origin)} <i class="fa-solid fa-arrow-right text-amber-500 text-xs mx-1"></i> ${escapeHtml(r.destination)}
              </span>
              <span class="px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-bold">
                ${formatMoney(r.unit_price)} บาท/เที่ยว
              </span>
              ${r.active ? '' : '<span class="text-[10px] px-2 py-0.5 rounded-full bg-slate-200 text-slate-600 font-semibold">ปิดใช้งาน</span>'}
            </div>
            <div class="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-600">
              <span><strong class="text-slate-700">แผ่นงาน Excel:</strong> ${escapeHtml(r.route_sheet_name)}</span>
              <span><strong class="text-slate-700">ชื่อในใบเบิก:</strong> ${escapeHtml(r.billing_title)}</span>
              <span><strong class="text-slate-700">เวลา:</strong> ${escapeHtml(r.default_dep_time || '-')} – ${escapeHtml(r.default_arr_time || '-')}</span>
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
    if (!name) return;
    try {
      await apiFetch('/api/master/workers', {
        method: 'POST',
        body: JSON.stringify({ name, default_vehicle })
      });
      $('input-worker-name').value = '';
      $('input-worker-default-vehicle').value = '';
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
    if (!site_name) return;
    try {
      await apiFetch('/api/master/sites', {
        method: 'POST',
        body: JSON.stringify({ site_name, zone_name })
      });
      $('input-site-name').value = '';
      $('input-site-zone').value = '';
      showToast(`เพิ่มไซต์งาน/โครงการ "${site_name}" เรียบร้อยแล้ว`, 'success');
      await loadMasterData();
    } catch (err) {
      showToast(err.message, 'error');
    }
  }

  async function handleAddRoute(e) {
    e.preventDefault();
    const origin = $('input-route-origin').value.trim();
    const destination = $('input-route-destination').value.trim();
    if (!origin || !destination) {
      showToast('กรุณาระบุต้นทางและปลายทาง', 'error');
      return;
    }
    const payload = {
      origin,
      destination,
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
      showToast(`เพิ่มเส้นทาง "${origin} → ${destination}" เรียบร้อยแล้ว`, 'success');
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
        <h3 class="text-base font-bold text-slate-900">แก้ไขข้อมูลคนขับ</h3>
        <form id="modal-edit-form" class="space-y-3.5">
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
                default_vehicle: box.querySelector('#m-veh').value.trim()
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
    openModal(
      `
      <div class="p-5 sm:p-6 space-y-4">
        <h3 class="text-base font-bold text-slate-900">แก้ไขไซต์งาน / โครงการ</h3>
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
            const newSiteName = box.querySelector('#m-site-name').value.trim();
            await apiFetch(`/api/master/sites/${s.id}`, {
              method: 'PUT',
              body: JSON.stringify({
                site_name: newSiteName,
                zone_name: box.querySelector('#m-zone-name').value.trim()
              })
            });
            if (state.boundBillingSite && state.boundBillingSite === s.site_name && newSiteName) {
              state.boundBillingSite = newSiteName;
              localStorage.setItem(BOUND_BILLING_SITE_KEY, newSiteName);
              await updateBindingUI();
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
              <label class="block text-xs font-semibold text-slate-700 mb-1">ต้นทาง *</label>
              <input type="text" id="m-origin" value="${escapeHtml(r.origin)}" required
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">ปลายทาง *</label>
              <input type="text" id="m-dest" value="${escapeHtml(r.destination)}" required
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
              <input type="time" id="m-dep" value="${escapeHtml(r.default_dep_time)}"
                class="w-full min-h-[40px] px-3 py-2 rounded-xl border border-slate-300 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">เวลาถึงปกติ</label>
              <input type="time" id="m-arr" value="${escapeHtml(r.default_arr_time)}"
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
        box.querySelector('#m-cancel').addEventListener('click', closeModal);
        box.querySelector('#modal-edit-form').addEventListener('submit', async (e) => {
          e.preventDefault();
          try {
            await apiFetch(`/api/master/routes/${r.id}`, {
              method: 'PUT',
              body: JSON.stringify({
                origin: box.querySelector('#m-origin').value.trim(),
                destination: box.querySelector('#m-dest').value.trim(),
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
        const ok = await confirmModal({
          title: 'ยืนยันการลบเส้นทาง',
          message: `ต้องการลบเส้นทาง "${r.origin} → ${r.destination}" ออกจากรายการข้อมูลหลักใช่หรือไม่?`,
          confirmText: 'ลบเส้นทาง',
          danger: true
        });
        if (!ok) return;
        await apiFetch(`/api/master/routes/${id}`, { method: 'DELETE' });
        showToast(`ลบเส้นทาง "${r.origin} → ${r.destination}" แล้ว`, 'success');
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
    $('form-worker').addEventListener('change', onWorkerSelectChange);
    $('form-route-select').addEventListener('change', onRouteSelectChange);
    $('btn-swap-route').addEventListener('click', onSwapRouteClick);
    $('btn-toggle-route-meta').addEventListener('click', () => toggleRouteMetaBox());
    $('form-origin').addEventListener('input', updateRouteMetaPreview);
    $('form-destination').addEventListener('input', updateRouteMetaPreview);
    $('form-route-sheet-name').addEventListener('input', updateRouteMetaPreview);
    $('form-billing-title').addEventListener('input', updateRouteMetaPreview);
    $('form-display-mode-zone').addEventListener('change', updateRouteMetaPreview);
    $('form-dep-time').addEventListener('input', updateDurationPreview);
    $('form-arr-time').addEventListener('input', updateDurationPreview);

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
    $('toggle-auto-sync').addEventListener('change', (e) => {
      state.autoSync = Boolean(e.target.checked);
      localStorage.setItem(AUTO_SYNC_KEY, state.autoSync ? 'true' : 'false');
      showToast(
        state.autoSync ? 'เปิดซิงค์ลงไฟล์ที่ผูกไว้อัตโนมัติแล้ว' : 'ปิดซิงค์อัตโนมัติแล้ว (กดปุ่มบันทึกไฟล์เองได้ตลอด)',
        'info'
      );
    });
    $('btn-clear-all-logs').addEventListener('click', handleClearAllLogs);
    $('logs-table-body').addEventListener('click', handleLogsActionClick);
    $('logs-mobile-cards').addEventListener('click', handleLogsActionClick);

    // Back-Office Forms & Action Delegation
    $('form-add-vehicle').addEventListener('submit', handleAddVehicle);
    $('form-add-worker').addEventListener('submit', handleAddWorker);
    $('form-add-site').addEventListener('submit', handleAddSite);
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
