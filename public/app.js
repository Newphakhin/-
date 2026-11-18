/**
 * ระบบจัดการบันทึกข้อมูลการเดินทางของคนงาน (Worker Dispatch & Travel Log System v2.0)
 * Cloud Supabase + Cartrack GPS API + Storage Quota Alert + Mobile-First Responsive UI
 */

const state = {
  token: localStorage.getItem('dispatch_auth_token') || null,
  user: JSON.parse(localStorage.getItem('dispatch_auth_user') || 'null'),
  currentView: 'front', // 'front' | 'back'
  editingWorkerName: null, // null = Create mode, string = Edit mode (PK)
  lastCheckedGps: null, // Stores latest checked GPS object from form button

  // Data from Backend
  activeWorkers: [],
  activeRoutes: [],
  activeVehicles: [],
  allWorkers: [],
  allRoutes: [],
  allVehicles: [],
  logs: [],

  // Storage Quota & System Settings
  storageQuota: {
    currentCount: 0,
    maxLimit: 100,
    usagePercent: 0,
    statusLevel: 'normal',
    isFull: false,
    isWarning: false,
    dbSizeBytes: 0,
    supabaseFreeTierBytes: 524288000,
    storageMode: 'local_fallback'
  },
  gpsSettings: {
    cartrack_api_url: 'https://fleetapi-th.cartrack.com/rest/vehicles/status',
    cartrack_username: 'DEMO_FLEET_TH',
    cartrack_api_key: '',
    cartrack_mock_mode: 'true',
    max_storage_limit: '100'
  },

  // Filters
  filterDate: '',
  filterSearch: ''
};

// ============================================================================
// UTILITY HELPERS (Date, Duration, GPS Parser, Escaping, API Fetch)
// ============================================================================

const THAI_MONTHS_SHORT = [
  'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
  'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
];

const THAI_MONTHS_FULL = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

function getTodayISO() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function formatThaiDate(isoDate, fullMonth = false) {
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) {
    return isoDate || '-';
  }
  const [yStr, mStr, dStr] = isoDate.split('-');
  const yearBE = Number(yStr) + 543;
  const monthIdx = Number(mStr) - 1;
  const day = Number(dStr);
  const monthName = fullMonth ? THAI_MONTHS_FULL[monthIdx] : THAI_MONTHS_SHORT[monthIdx];
  return `${day} ${monthName} ${yearBE}`;
}

function calculateTravelDuration(departureTime, arrivalTime) {
  if (!departureTime || !arrivalTime) {
    return { text: '-', isOvernight: false, totalMinutes: null };
  }
  const depMatch = /^(\d{1,2}):(\d{2})$/.exec(departureTime.trim());
  const arrMatch = /^(\d{1,2}):(\d{2})$/.exec(arrivalTime.trim());
  if (!depMatch || !arrMatch) {
    return { text: '-', isOvernight: false, totalMinutes: null };
  }

  const depMins = Number(depMatch[1]) * 60 + Number(depMatch[2]);
  const arrMins = Number(arrMatch[1]) * 60 + Number(arrMatch[2]);

  let diffMins = arrMins - depMins;
  let isOvernight = false;

  if (diffMins < 0) {
    diffMins = 1440 - depMins + arrMins;
    isOvernight = true;
  }

  const hours = Math.floor(diffMins / 60);
  const mins = diffMins % 60;

  const parts = [];
  if (hours > 0) parts.push(`${hours} ชม.`);
  if (mins > 0) parts.push(`${mins} นาที`);
  if (parts.length === 0) parts.push('0 นาที');

  const label = parts.join(' ') + (isOvernight ? ' (ข้ามวัน)' : '');
  return {
    text: label,
    isOvernight,
    totalMinutes: diffMins
  };
}

function parseGpsStatus(rawGps) {
  if (!rawGps) return null;
  if (typeof rawGps === 'object') return rawGps;
  try {
    const parsed = JSON.parse(rawGps);
    if (parsed && typeof parsed === 'object') return parsed;
  } catch {
    // Plain text fallback
    return {
      status: String(rawGps),
      engine: String(rawGps).includes('ดับ') ? 'OFF' : 'ON',
      speed: 0,
      location: '',
      maps_url: 'https://www.google.com/maps?q=13.7563,100.5018',
      updated_at: ''
    };
  }
  return null;
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

async function apiFetch(url, options = {}) {
  const headers = {
    'Content-Type': 'application/json',
    ...(options.headers || {})
  };
  if (state.token) {
    headers.Authorization = `Bearer ${state.token}`;
  }

  const response = await fetch(url, {
    ...options,
    headers
  });

  let data = null;
  try {
    data = await response.json();
  } catch {
    data = null;
  }

  if (response.status === 401 && url !== '/api/login') {
    handleSessionExpired();
    throw new Error(data?.error || 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  }

  return { ok: response.ok, status: response.status, data };
}

// ============================================================================
// TOAST NOTIFICATION & CUSTOM MODAL SYSTEM (No alert() / confirm())
// ============================================================================

function showToast(message, type = 'success', subtitle = '') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const styleMap = {
    success: {
      bg: 'bg-white border-emerald-500',
      iconBg: 'bg-emerald-100 text-emerald-600',
      icon: 'fa-circle-check',
      titleColor: 'text-slate-900'
    },
    warning: {
      bg: 'bg-white border-amber-500',
      iconBg: 'bg-amber-100 text-amber-600',
      icon: 'fa-triangle-exclamation',
      titleColor: 'text-slate-900'
    },
    error: {
      bg: 'bg-white border-rose-500',
      iconBg: 'bg-rose-100 text-rose-600',
      icon: 'fa-circle-exclamation',
      titleColor: 'text-slate-900'
    },
    info: {
      bg: 'bg-white border-indigo-500',
      iconBg: 'bg-indigo-100 text-indigo-600',
      icon: 'fa-circle-info',
      titleColor: 'text-slate-900'
    }
  };

  const cfg = styleMap[type] || styleMap.info;
  const toastEl = document.createElement('div');
  toastEl.className = `pointer-events-auto flex items-start gap-3 p-3.5 rounded-2xl shadow-xl border-l-4 border border-slate-200/90 ${cfg.bg} toast-animate`;
  toastEl.innerHTML = `
    <div class="w-8 h-8 rounded-xl ${cfg.iconBg} flex items-center justify-center shrink-0 mt-0.5">
      <i class="fa-solid ${cfg.icon}"></i>
    </div>
    <div class="flex-1 min-w-0">
      <p class="text-xs sm:text-sm font-semibold ${cfg.titleColor} break-words">${escapeHtml(message)}</p>
      ${subtitle ? `<p class="text-xs text-slate-500 mt-0.5 break-words">${escapeHtml(subtitle)}</p>` : ''}
    </div>
    <button type="button" class="text-slate-400 hover:text-slate-600 p-1 cursor-pointer">
      <i class="fa-solid fa-xmark text-xs"></i>
    </button>
  `;

  toastEl.querySelector('button').addEventListener('click', () => {
    toastEl.remove();
  });

  container.appendChild(toastEl);
  setTimeout(() => {
    if (toastEl.parentNode) {
      toastEl.style.opacity = '0';
      toastEl.style.transition = 'opacity 0.25s ease';
      setTimeout(() => toastEl.remove(), 250);
    }
  }, 4200);
}

function closeCustomModal() {
  const backdrop = document.getElementById('custom-modal-backdrop');
  const box = document.getElementById('custom-modal-box');
  if (backdrop) {
    backdrop.classList.add('hidden');
    backdrop.classList.remove('flex');
  }
  if (box) {
    box.innerHTML = '';
  }
}

function openCustomModal(htmlContent) {
  const backdrop = document.getElementById('custom-modal-backdrop');
  const box = document.getElementById('custom-modal-box');
  if (!backdrop || !box) return;

  box.innerHTML = htmlContent;
  backdrop.classList.remove('hidden');
  backdrop.classList.add('flex');
}

document.getElementById('custom-modal-backdrop')?.addEventListener('click', (e) => {
  if (e.target.id === 'custom-modal-backdrop') {
    closeCustomModal();
  }
});

// ============================================================================
// STORAGE QUOTA METER & 2-LEVEL ALERT SYSTEM (80% Warning / 100% Red Modal)
// ============================================================================

function renderStorageQuotaUI(triggerFullModal = false) {
  const q = state.storageQuota;
  if (!q) return;

  // Update Top Cloud / Local Database Banner
  const modeDot = document.getElementById('db-mode-dot');
  const modeText = document.getElementById('db-mode-text');
  if (modeDot && modeText) {
    if (q.storageMode === 'supabase_cloud') {
      modeDot.className = 'w-2 h-2 rounded-full bg-emerald-400';
      modeText.innerHTML =
        '<strong>สถานะฐานข้อมูล: เชื่อมต่อ Supabase PostgreSQL Cloud + SQLite เรียบร้อยแล้ว</strong> — ซิงค์ข้อมูลออนไลน์แบบเรียลไทม์';
    } else {
      modeDot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
      modeText.innerHTML =
        '<strong>สถานะฐานข้อมูล: เชื่อมต่อฐานข้อมูล SQLite จริงเรียบร้อยแล้ว (database.sqlite)</strong> — บันทึกข้อมูลถาวรครบทั้ง 6 ตาราง (พร้อมซิงค์ขึ้น Supabase Cloud ได้ทันทีเมื่อใส่ค่าใน <code>.env</code>)';
    }
  }

  // Update KPI Max Quota
  const statMaxQuota = document.getElementById('stat-max-quota');
  if (statMaxQuota) statMaxQuota.textContent = q.maxLimit;

  // Update Storage Usage Meter Card
  const progressBar = document.getElementById('quota-progress-bar');
  const progressLabel = document.getElementById('quota-progress-label');
  const dbSizeLabel = document.getElementById('quota-db-size-label');
  const statusBadge = document.getElementById('quota-status-badge');
  const meterIcon = document.getElementById('quota-meter-icon');
  const warningBanner = document.getElementById('storage-warning-banner');
  const adminSimControls = document.getElementById('admin-quota-test-controls');

  if (adminSimControls) {
    if (state.user?.role === 'admin') {
      adminSimControls.classList.remove('hidden');
      adminSimControls.classList.add('flex');
    } else {
      adminSimControls.classList.add('hidden');
      adminSimControls.classList.remove('flex');
    }
  }

  const pct = Math.min(100, Math.max(0, q.usagePercent || 0));
  if (progressBar) {
    progressBar.style.width = `${pct}%`;
  }
  if (progressLabel) {
    progressLabel.textContent = `ความจุข้อมูลปัจจุบัน: ${q.currentCount} / ${q.maxLimit} รายการ (${pct}%)`;
  }
  if (dbSizeLabel) {
    const kb = ((q.dbSizeBytes || 0) / 1024).toFixed(1);
    dbSizeLabel.textContent = `ขนาดฐานข้อมูล: ${kb} KB / Supabase Free Tier 500 MB`;
  }

  // Level styling
  if (q.isFull || q.statusLevel === 'full') {
    if (progressBar) progressBar.className = 'h-full bg-rose-600 rounded-full transition-all duration-300';
    if (statusBadge) {
      statusBadge.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-rose-100 text-rose-800 border border-rose-300';
      statusBadge.innerHTML = '<i class="fa-solid fa-circle-exclamation mr-1"></i>ข้อมูลเต็ม 100% (บล็อกการบันทึกใหม่)';
    }
    if (meterIcon) {
      meterIcon.className = 'w-10 h-10 rounded-xl bg-rose-100 text-rose-600 flex items-center justify-center shrink-0';
    }

    if (warningBanner) {
      warningBanner.className =
        'rounded-2xl p-4 bg-rose-50 border-2 border-rose-300 text-rose-900 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3';
      warningBanner.innerHTML = `
        <div class="flex items-start gap-3">
          <div class="w-10 h-10 rounded-xl bg-rose-600 text-white flex items-center justify-center shrink-0 mt-0.5">
            <i class="fa-solid fa-triangle-exclamation text-lg"></i>
          </div>
          <div>
            <h4 class="text-sm font-bold text-rose-950">
              แจ้งเตือนฉุกเฉิน: พื้นที่เก็บข้อมูลเต็ม 100% (${q.currentCount}/${q.maxLimit} รายการ) — HTTP 507 Insufficient Storage
            </h4>
            <p class="text-xs text-rose-800 mt-0.5">
              ระบบบล็อกการเพิ่มข้อมูลใหม่ชั่วคราว กรุณาส่งออกไฟล์สำรองข้อมูล หรือล้างข้อมูลเก่า/ขยายโควต้าในหน้าตั้งค่าหลังบ้าน
            </p>
          </div>
        </div>
        <div class="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            id="banner-btn-export-csv"
            class="min-h-[40px] px-3.5 py-1.5 rounded-xl bg-white hover:bg-rose-100 text-rose-900 border border-rose-300 text-xs font-bold flex items-center gap-1.5 cursor-pointer"
          >
            <i class="fa-solid fa-file-csv"></i>
            <span>ส่งออก CSV สำรองข้อมูล</span>
          </button>
          <button
            type="button"
            id="banner-btn-open-full-modal"
            class="min-h-[40px] px-3.5 py-1.5 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold flex items-center gap-1.5 cursor-pointer"
          >
            <i class="fa-solid fa-broom"></i>
            <span>จัดการพื้นที่เก็บข้อมูล</span>
          </button>
        </div>
      `;
      warningBanner.querySelector('#banner-btn-export-csv')?.addEventListener('click', exportLogsToCSV);
      warningBanner
        .querySelector('#banner-btn-open-full-modal')
        ?.addEventListener('click', () => showStorageFullEmergencyModal());
    }

    if (triggerFullModal) {
      showStorageFullEmergencyModal();
    }
  } else if (q.isWarning || q.statusLevel === 'warning') {
    if (progressBar) progressBar.className = 'h-full bg-amber-500 rounded-full transition-all duration-300';
    if (statusBadge) {
      statusBadge.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300';
      statusBadge.innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-1"></i>ใกล้เต็ม (${pct}%)`;
    }
    if (meterIcon) {
      meterIcon.className = 'w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0';
    }

    // Level 1 Amber Warning Banner (>= 80%)
    if (warningBanner) {
      warningBanner.className =
        'rounded-2xl p-4 bg-amber-50 border-2 border-amber-300 text-amber-900 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3';
      warningBanner.innerHTML = `
        <div class="flex items-start gap-3">
          <div class="w-10 h-10 rounded-xl bg-amber-500 text-white flex items-center justify-center shrink-0 mt-0.5">
            <i class="fa-solid fa-bell text-lg"></i>
          </div>
          <div>
            <h4 class="text-sm font-bold text-amber-950">
              แจ้งเตือนระดับที่ 1 (Warning): พื้นที่เก็บข้อมูลใกล้เต็ม (${q.currentCount}/${q.maxLimit} รายการ — ${pct}%)
            </h4>
            <p class="text-xs text-amber-800 mt-0.5">
              ความจุข้อมูลเกิน 80% แล้ว แนะนำให้กดส่งออกไฟล์ CSV / Excel เพื่อสำรองข้อมูลไว้ล่วงหน้าก่อนที่พื้นที่เก็บข้อมูลจะเต็ม
            </p>
          </div>
        </div>
        <div class="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            id="banner-warn-export-excel"
            class="min-h-[40px] px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 cursor-pointer"
          >
            <i class="fa-solid fa-file-excel"></i>
            <span>สำรองไฟล์ Excel</span>
          </button>
          <button
            type="button"
            id="banner-warn-export-csv"
            class="min-h-[40px] px-3.5 py-1.5 rounded-xl bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold flex items-center gap-1.5 cursor-pointer"
          >
            <i class="fa-solid fa-file-csv"></i>
            <span>ส่งออกไฟล์ CSV</span>
          </button>
        </div>
      `;
      warningBanner
        .querySelector('#banner-warn-export-excel')
        ?.addEventListener('click', exportLogsToExcel);
      warningBanner
        .querySelector('#banner-warn-export-csv')
        ?.addEventListener('click', exportLogsToCSV);
    }
  } else {
    if (progressBar) progressBar.className = 'h-full bg-indigo-600 rounded-full transition-all duration-300';
    if (statusBadge) {
      statusBadge.className = 'px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800';
      statusBadge.textContent = `ปกติ (${pct}%)`;
    }
    if (meterIcon) {
      meterIcon.className = 'w-10 h-10 rounded-xl bg-indigo-50 text-indigo-600 flex items-center justify-center shrink-0';
    }
    if (warningBanner) {
      warningBanner.className = 'hidden';
      warningBanner.innerHTML = '';
    }
  }
}

function showStorageFullEmergencyModal(customMessage = '') {
  const q = state.storageQuota;
  const isAdmin = state.user?.role === 'admin';

  openCustomModal(`
    <div class="border-t-8 border-rose-600 p-6 bg-white">
      <div class="flex items-start gap-4">
        <div class="w-14 h-14 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center text-2xl shrink-0 animate-bounce">
          <i class="fa-solid fa-database"></i>
        </div>
        <div class="flex-1">
          <span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-rose-100 text-rose-800 border border-rose-300 text-[11px] font-bold mb-1.5">
            HTTP 507 Insufficient Storage — Storage Full Alert
          </span>
          <h3 class="text-lg font-bold text-rose-950">
            แจ้งเตือนฉุกเฉิน: พื้นที่เก็บข้อมูลเต็ม 100%!
          </h3>
          <p class="text-sm text-slate-600 mt-1 leading-relaxed">
            ${escapeHtml(
              customMessage ||
                `ขณะนี้มีข้อมูลบันทึกการเดินทางครบขีดจำกัดแล้ว (${q.currentCount} / ${q.maxLimit} รายการ) ระบบจึงบล็อกการบันทึกข้อมูลใหม่ชั่วคราว`
            )}
          </p>
        </div>
      </div>

      <div class="mt-5 p-4 rounded-xl bg-rose-50 border border-rose-200 space-y-2.5">
        <p class="text-xs font-bold text-rose-900 uppercase tracking-wider">
          <i class="fa-solid fa-bolt mr-1"></i>ทางเลือกดำเนินการฉุกเฉิน:
        </p>
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
          <button
            type="button"
            id="modal-full-export-csv"
            class="min-h-[44px] px-3.5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition cursor-pointer"
          >
            <i class="fa-solid fa-file-csv text-base"></i>
            <span>ส่งออก CSV สำรองข้อมูลทันที</span>
          </button>

          <button
            type="button"
            id="modal-full-export-excel"
            class="min-h-[44px] px-3.5 py-2.5 rounded-xl bg-teal-600 hover:bg-teal-700 text-white text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition cursor-pointer"
          >
            <i class="fa-solid fa-file-excel text-base"></i>
            <span>ส่งออก Excel (.xlsx) ทันที</span>
          </button>

          ${
            isAdmin
              ? `
            <button
              type="button"
              id="modal-full-clear-old"
              class="min-h-[44px] px-3.5 py-2.5 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition cursor-pointer"
            >
              <i class="fa-solid fa-broom text-base"></i>
              <span>ล้างข้อมูลเก่าเพื่อเพิ่มพื้นที่</span>
            </button>

            <button
              type="button"
              id="modal-full-expand-quota"
              class="min-h-[44px] px-3.5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition cursor-pointer"
            >
              <i class="fa-solid fa-sliders text-base"></i>
              <span>ขยายโควต้าในหน้าตั้งค่า (100)</span>
            </button>
          `
              : ''
          }
        </div>
      </div>

      <div class="mt-5 pt-3 border-t border-slate-100 flex justify-end">
        <button
          type="button"
          id="modal-full-close"
          class="min-h-[42px] px-5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-sm font-semibold cursor-pointer"
        >
          ปิดหน้าต่าง
        </button>
      </div>
    </div>
  `);

  document.getElementById('modal-full-close')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-full-export-csv')?.addEventListener('click', () => {
    exportLogsToCSV();
  });
  document.getElementById('modal-full-export-excel')?.addEventListener('click', () => {
    exportLogsToExcel();
  });
  document.getElementById('modal-full-clear-old')?.addEventListener('click', async () => {
    closeCustomModal();
    const res = await apiFetch('/api/logs/clear-old', {
      method: 'POST',
      body: JSON.stringify({ mode: 'all' })
    });
    if (res.ok) {
      showToast(res.data?.message || 'ล้างข้อมูลเก่าเพื่อเพิ่มพื้นที่สำเร็จ', 'success');
      await loadAllData();
    } else {
      showToast(res.data?.error || 'ไม่สามารถล้างข้อมูลเก่าได้', 'error');
    }
  });
  document.getElementById('modal-full-expand-quota')?.addEventListener('click', async () => {
    closeCustomModal();
    await quickSetQuotaLimit(100);
  });
}

async function quickSetQuotaLimit(newLimit) {
  try {
    const res = await apiFetch('/api/settings/gps', {
      method: 'PUT',
      body: JSON.stringify({
        ...state.gpsSettings,
        max_storage_limit: newLimit
      })
    });
    if (res.ok) {
      state.gpsSettings = res.data.settings;
      state.storageQuota = res.data.storageQuota;
      const triggerFull = state.storageQuota.isFull;
      renderStorageQuotaUI(triggerFull);
      showToast(
        `ปรับขีดจำกัดโควต้าพื้นที่เป็น ${newLimit} รายการแล้ว`,
        state.storageQuota.isFull
          ? 'error'
          : state.storageQuota.isWarning
          ? 'warning'
          : 'success'
      );
    }
  } catch (err) {
    showToast(err.message || 'ไม่สามารถปรับโควต้าได้', 'error');
  }
}

// ============================================================================
// AUTHENTICATION & NAVIGATION
// ============================================================================

function saveSession(token, user) {
  state.token = token;
  state.user = user;
  localStorage.setItem('dispatch_auth_token', token);
  localStorage.setItem('dispatch_auth_user', JSON.stringify(user));
}

function clearSession() {
  state.token = null;
  state.user = null;
  localStorage.removeItem('dispatch_auth_token');
  localStorage.removeItem('dispatch_auth_user');
}

function handleSessionExpired() {
  clearSession();
  renderAuthView();
  showToast('กรุณาเข้าสู่ระบบใหม่อีกครั้ง', 'warning');
}

function renderAuthView() {
  const loginScreen = document.getElementById('login-screen');
  const appScreen = document.getElementById('app-screen');

  if (!state.token || !state.user) {
    loginScreen.classList.remove('hidden');
    appScreen.classList.add('hidden');
    appScreen.classList.remove('flex');
    return;
  }

  loginScreen.classList.add('hidden');
  appScreen.classList.remove('hidden');
  appScreen.classList.add('flex');

  const fullNameEl = document.getElementById('nav-user-fullname');
  const roleBadgeEl = document.getElementById('nav-user-role-badge');
  const navTabBack = document.getElementById('nav-tab-back');
  const mobNavBack = document.getElementById('mob-nav-backoffice');

  fullNameEl.textContent = state.user.full_name || state.user.username;

  if (state.user.role === 'admin') {
    fullNameEl.textContent = '';
    fullNameEl.classList.add('hidden');
    roleBadgeEl.className =
      'inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-semibold bg-zinc-900 text-amber-400 border border-zinc-800 shadow-sm';
    roleBadgeEl.innerHTML = '<i class="fa-solid fa-user-shield text-[11px]"></i> แอดมิน';

    navTabBack.classList.remove('hidden');
    navTabBack.classList.add('flex');
    mobNavBack.classList.remove('hidden');
    mobNavBack.classList.add('flex');
  } else {
    fullNameEl.classList.remove('hidden');
    roleBadgeEl.className =
      'inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200';
    roleBadgeEl.innerHTML = '<i class="fa-solid fa-user-check text-[10px]"></i> ผู้บันทึกข้อมูล (Operator)';

    navTabBack.classList.add('hidden');
    navTabBack.classList.remove('flex');
    mobNavBack.classList.add('hidden');
    mobNavBack.classList.remove('flex');

    if (state.currentView === 'back') {
      state.currentView = 'front';
    }
  }

  switchView(state.currentView);
}

function switchView(viewName) {
  if (viewName === 'back' && state.user?.role !== 'admin') {
    showToast('เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่เข้าถึงระบบหลังบ้านได้', 'error');
    return;
  }

  state.currentView = viewName;
  const viewFront = document.getElementById('view-front-office');
  const viewBack = document.getElementById('view-back-office');
  const tabFront = document.getElementById('nav-tab-front');
  const tabBack = document.getElementById('nav-tab-back');
  const mobForm = document.getElementById('mob-nav-form');
  const mobTable = document.getElementById('mob-nav-table');
  const mobBack = document.getElementById('mob-nav-backoffice');

  if (viewName === 'front') {
    viewFront.classList.remove('hidden');
    viewBack.classList.add('hidden');

    tabFront.className =
      'min-h-[40px] px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-semibold transition flex items-center gap-1.5 bg-zinc-900 text-white shadow-sm cursor-pointer';
    tabBack.className =
      state.user?.role === 'admin'
        ? 'flex min-h-[40px] px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-semibold transition items-center gap-1.5 text-zinc-600 hover:text-zinc-900 cursor-pointer'
        : 'hidden';

    if (mobForm) mobForm.className = 'flex-1 min-h-[46px] flex flex-col items-center justify-center rounded-xl text-zinc-900 font-semibold text-[11px] gap-0.5 cursor-pointer';
    if (mobTable) mobTable.className = 'flex-1 min-h-[46px] flex flex-col items-center justify-center rounded-xl text-slate-600 font-medium text-[11px] gap-0.5 cursor-pointer';
    if (mobBack && state.user?.role === 'admin') {
      mobBack.className = 'flex flex-1 min-h-[46px] flex-col items-center justify-center rounded-xl text-slate-600 font-medium text-[11px] gap-0.5 cursor-pointer';
    }
  } else {
    viewFront.classList.add('hidden');
    viewBack.classList.remove('hidden');

    tabFront.className =
      'min-h-[40px] px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-semibold transition flex items-center gap-1.5 text-zinc-600 hover:text-zinc-900 cursor-pointer';
    tabBack.className =
      'flex min-h-[40px] px-3.5 py-1.5 rounded-lg text-xs sm:text-sm font-semibold transition items-center gap-1.5 bg-zinc-900 text-white shadow-sm cursor-pointer';

    if (mobForm) mobForm.className = 'flex-1 min-h-[46px] flex flex-col items-center justify-center rounded-xl text-slate-600 font-medium text-[11px] gap-0.5 cursor-pointer';
    if (mobTable) mobTable.className = 'flex-1 min-h-[46px] flex flex-col items-center justify-center rounded-xl text-slate-600 font-medium text-[11px] gap-0.5 cursor-pointer';
    if (mobBack) mobBack.className = 'flex flex-1 min-h-[46px] flex-col items-center justify-center rounded-xl text-zinc-900 font-semibold text-[11px] gap-0.5 cursor-pointer';

    loadBackOfficeData();
  }
}

// ============================================================================
// DATA LOADING & SYNCHRONIZATION
// ============================================================================

async function loadAllData() {
  try {
    const requests = [
      apiFetch('/api/master/workers?active=1'),
      apiFetch('/api/master/routes?active=1'),
      apiFetch('/api/master/vehicles?active=1'),
      apiFetch('/api/logs'),
      apiFetch('/api/master/workers'),
      apiFetch('/api/master/routes'),
      apiFetch('/api/master/vehicles'),
      apiFetch('/api/storage-status')
    ];

    if (state.user?.role === 'admin') {
      requests.push(apiFetch('/api/settings/gps'));
    }

    const [
      workersRes,
      routesRes,
      vehiclesRes,
      logsRes,
      allWorkersRes,
      allRoutesRes,
      allVehiclesRes,
      quotaRes,
      gpsSettingsRes
    ] = await Promise.all(requests);

    if (workersRes.ok) state.activeWorkers = workersRes.data || [];
    if (routesRes.ok) state.activeRoutes = routesRes.data || [];
    if (vehiclesRes.ok) state.activeVehicles = vehiclesRes.data || [];
    if (logsRes.ok) state.logs = logsRes.data || [];
    if (allWorkersRes.ok) state.allWorkers = allWorkersRes.data || [];
    if (allRoutesRes.ok) state.allRoutes = allRoutesRes.data || [];
    if (allVehiclesRes.ok) state.allVehicles = allVehiclesRes.data || [];
    if (quotaRes.ok && quotaRes.data) state.storageQuota = quotaRes.data;
    if (gpsSettingsRes?.ok && gpsSettingsRes.data?.settings) {
      state.gpsSettings = gpsSettingsRes.data.settings;
      populateGpsSettingsForm();
    }

    renderWorkerDropdown();
    renderVehicleDropdown();
    renderRouteComboboxMenu();
    renderSharedRoutesDatalist();
    renderLogsTable();
    updateSummaryKPIs();
    updateRouteSmartControls();
    renderStorageQuotaUI(false);

    if (state.currentView === 'back' && state.user?.role === 'admin') {
      renderBackOfficeLists();
    }
  } catch (err) {
    console.error('Error loading data:', err);
  }
}

async function loadBackOfficeData() {
  await loadAllData();
  renderBackOfficeLists();
}

function updateSummaryKPIs() {
  const today = getTodayISO();
  const totalLogs = state.logs.length;
  const todayLogs = state.logs.filter((l) => l.travel_date === today).length;
  const loggedWorkerNames = new Set(state.logs.map((l) => l.worker_name));
  const availableActiveWorkers = state.activeWorkers.filter(
    (w) => !loggedWorkerNames.has(w.name)
  ).length;

  document.getElementById('stat-total-logs').textContent = totalLogs;
  document.getElementById('stat-today-logs').textContent = todayLogs;
  document.getElementById('stat-available-workers').textContent = availableActiveWorkers;
  document.getElementById('stat-active-workers').textContent = state.activeWorkers.length;
  document.getElementById('stat-active-routes').textContent = state.activeRoutes.length;
  document.getElementById('stat-active-vehicles').textContent = state.activeVehicles.length;
}

// ============================================================================
// FRONT-OFFICE DROPDOWNS, SMART COMBO BOX & CARTRACK GPS CHECKER
// ============================================================================

function renderWorkerDropdown() {
  const select = document.getElementById('form-worker');
  if (!select) return;

  const currentVal = state.editingWorkerName || select.value;
  const loggedSet = new Set(state.logs.map((l) => l.worker_name));

  let html = '<option value="">-- เลือกรายชื่อคนงาน (บังคับเลือก) --</option>';

  const listToRender = [...state.activeWorkers];
  if (
    state.editingWorkerName &&
    !listToRender.some((w) => w.name === state.editingWorkerName)
  ) {
    const foundInAll = state.allWorkers.find((w) => w.name === state.editingWorkerName);
    listToRender.unshift(
      foundInAll || { id: 0, name: state.editingWorkerName, default_route: '', active: false }
    );
  }

  for (const w of listToRender) {
    const routeSuffix = w.default_route
      ? ` (ไซต์ประจำ: ${w.default_route})`
      : ' (ไม่มีไซต์ประจำ)';
    const alreadyLoggedSuffix = loggedSet.has(w.name) ? ' [มีบันทึกแล้ว]' : '';
    const isSelected = w.name === currentVal ? 'selected' : '';
    html += `<option value="${escapeHtml(w.name)}" ${isSelected}>${escapeHtml(
      w.name + routeSuffix + alreadyLoggedSuffix
    )}</option>`;
  }

  select.innerHTML = html;
  select.disabled = Boolean(state.editingWorkerName);
}

function renderVehicleDropdown() {
  const select = document.getElementById('form-vehicle');
  if (!select) return;

  const currentVal = select.value;
  let html = '<option value="">-- เลือกทะเบียนรถ --</option>';

  const listToRender = [...state.activeVehicles];
  if (currentVal && !listToRender.some((v) => v.plate === currentVal)) {
    listToRender.unshift({ id: 0, plate: currentVal, active: false });
  }

  for (const v of listToRender) {
    const isSelected = v.plate === currentVal ? 'selected' : '';
    html += `<option value="${escapeHtml(v.plate)}" ${isSelected}>${escapeHtml(v.plate)}</option>`;
  }

  select.innerHTML = html;
}

function renderSharedRoutesDatalist() {
  const datalist = document.getElementById('datalist-master-routes');
  if (!datalist) return;
  datalist.innerHTML = state.activeRoutes
    .map((r) => `<option value="${escapeHtml(r.name)}"></option>`)
    .join('');
}

function renderRouteComboboxMenu(filterText = '') {
  const menu = document.getElementById('route-combobox-menu');
  if (!menu) return;

  const q = filterText.trim().toLowerCase();
  const filtered = q
    ? state.activeRoutes.filter((r) => r.name.toLowerCase().includes(q))
    : state.activeRoutes;

  if (filtered.length === 0) {
    menu.innerHTML = `
      <div class="px-3.5 py-3 text-xs text-slate-500">
        ไม่พบในรายการที่หมายเดิม (สามารถใช้ข้อความที่พิมพ์เองได้ หรือกด "+ เพิ่มในรายการที่หมาย")
      </div>
    `;
    return;
  }

  const selectedWorker = getSelectedWorkerObject();
  const workerDefault = selectedWorker?.default_route || '';

  menu.innerHTML = filtered
    .map((r) => {
      const isDefault = workerDefault && r.name === workerDefault;
      return `
        <button
          type="button"
          data-route-name="${escapeHtml(r.name)}"
          class="w-full min-h-[42px] text-left px-3.5 py-2.5 text-xs sm:text-sm hover:bg-indigo-50 text-slate-800 flex items-center justify-between gap-2 transition cursor-pointer"
        >
          <span class="flex items-center gap-2 truncate">
            <i class="fa-solid fa-location-dot text-indigo-500 text-xs"></i>
            <span class="truncate font-medium">${escapeHtml(r.name)}</span>
          </span>
          ${
            isDefault
              ? '<span class="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-semibold shrink-0">ไซต์งานประจำ</span>'
              : ''
          }
        </button>
      `;
    })
    .join('');

  menu.querySelectorAll('button[data-route-name]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const chosen = btn.getAttribute('data-route-name');
      const routeInput = document.getElementById('form-route');
      routeInput.value = chosen;
      menu.classList.add('hidden');
      updateRouteSmartControls();
    });
  });
}

function getSelectedWorkerObject() {
  const workerName = document.getElementById('form-worker')?.value || '';
  if (!workerName) return null;
  return (
    state.allWorkers.find((w) => w.name === workerName) ||
    state.activeWorkers.find((w) => w.name === workerName) ||
    null
  );
}

function updateRouteSmartControls() {
  const routeInput = document.getElementById('form-route');
  const badgeDefault = document.getElementById('badge-default-route');
  const btnResetDefault = document.getElementById('btn-reset-default-route');
  const btnAddToRoutes = document.getElementById('btn-add-to-routes');
  const btnSetWorkerDefault = document.getElementById('btn-set-worker-default-route');

  if (!routeInput) return;

  const currentRoute = routeInput.value.trim();
  const worker = getSelectedWorkerObject();
  const workerDefault = (worker?.default_route || '').trim();

  if (worker && workerDefault && currentRoute === workerDefault) {
    badgeDefault.classList.remove('hidden');
    badgeDefault.classList.add('inline-flex');
  } else {
    badgeDefault.classList.add('hidden');
    badgeDefault.classList.remove('inline-flex');
  }

  if (worker && workerDefault && currentRoute !== workerDefault) {
    btnResetDefault.classList.remove('hidden');
    btnResetDefault.classList.add('inline-flex');
    btnResetDefault.title = `รีเซ็ตกลับเป็น: ${workerDefault}`;
  } else {
    btnResetDefault.classList.add('hidden');
    btnResetDefault.classList.remove('inline-flex');
  }

  const existsInActiveRoutes = state.activeRoutes.some(
    (r) => r.name.toLowerCase() === currentRoute.toLowerCase()
  );
  if (!currentRoute || existsInActiveRoutes) {
    btnAddToRoutes.disabled = true;
    btnAddToRoutes.innerHTML = existsInActiveRoutes
      ? '<i class="fa-solid fa-check text-[11px]"></i><span>มีในรายการที่หมายแล้ว</span>'
      : '<i class="fa-solid fa-plus text-[11px]"></i><span>+ เพิ่มในรายการที่หมาย</span>';
  } else {
    btnAddToRoutes.disabled = false;
    btnAddToRoutes.innerHTML =
      '<i class="fa-solid fa-plus text-[11px]"></i><span>+ เพิ่มในรายการที่หมาย</span>';
  }

  if (!worker || !worker.id || !currentRoute || currentRoute === workerDefault) {
    btnSetWorkerDefault.disabled = true;
    btnSetWorkerDefault.innerHTML =
      worker && workerDefault && currentRoute === workerDefault
        ? '<i class="fa-solid fa-check text-[11px]"></i><span>เป็นไซต์งานประจำแล้ว</span>'
        : '<i class="fa-solid fa-thumbtack text-[11px]"></i><span>ตั้งเป็นไซต์งานประจำคนนี้</span>';
  } else {
    btnSetWorkerDefault.disabled = false;
    btnSetWorkerDefault.innerHTML =
      '<i class="fa-solid fa-thumbtack text-[11px]"></i><span>ตั้งเป็นไซต์งานประจำคนนี้</span>';
  }
}

function updateDateAndDurationPreview() {
  const dateVal = document.getElementById('form-date')?.value || '';
  const datePreview = document.getElementById('form-date-thai-preview');
  if (datePreview) {
    datePreview.textContent = dateVal ? `วันที่ไทย: ${formatThaiDate(dateVal, true)}` : 'ยังไม่ได้ระบุวันที่';
  }

  const depVal = document.getElementById('form-departure')?.value || '';
  const arrVal = document.getElementById('form-arrival')?.value || '';
  const durPreview = document.getElementById('form-duration-preview');
  if (durPreview) {
    const dur = calculateTravelDuration(depVal, arrVal);
    if (dur.totalMinutes !== null) {
      durPreview.className = dur.isOvernight
        ? 'mt-1.5 text-[11px] text-amber-700 font-semibold flex items-center gap-1'
        : 'mt-1.5 text-[11px] text-emerald-700 font-semibold flex items-center gap-1';
      durPreview.innerHTML = `<i class="fa-solid fa-stopwatch"></i><span>ระยะเวลาเดินทาง: ${escapeHtml(
        dur.text
      )}</span>`;
    } else {
      durPreview.className = 'mt-1.5 text-[11px] text-slate-500 flex items-center gap-1';
      durPreview.innerHTML =
        '<i class="fa-solid fa-stopwatch text-slate-400"></i><span>คำนวณระยะเวลาอัตโนมัติ (รองรับข้ามวัน)</span>';
    }
  }
}

async function checkVehicleGpsFromForm() {
  const plate = document.getElementById('form-vehicle')?.value || '';
  const previewCard = document.getElementById('form-gps-preview-card');
  if (!plate) {
    showToast('กรุณาเลือกหมายเลขทะเบียนรถก่อนกดเช็กพิกัด GPS', 'warning');
    document.getElementById('form-vehicle')?.focus();
    return;
  }

  try {
    const res = await apiFetch(`/api/gps/vehicle/${encodeURIComponent(plate)}`);
    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถดึงพิกัด GPS ได้', 'error');
      return;
    }

    const gps = res.data;
    state.lastCheckedGps = gps;

    if (previewCard) {
      previewCard.classList.remove('hidden');
      const isMoving = Number(gps.speed) > 0;
      previewCard.innerHTML = `
        <div class="flex items-start justify-between gap-2">
          <div class="space-y-1">
            <div class="flex flex-wrap items-center gap-1.5">
              <span class="font-bold text-slate-900">
                <i class="fa-solid fa-satellite-dish text-emerald-600 mr-1"></i>${escapeHtml(gps.plate)}
              </span>
              <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${
                isMoving
                  ? 'bg-emerald-600 text-white'
                  : gps.engine === 'ON'
                  ? 'bg-amber-100 text-amber-800 border border-amber-300'
                  : 'bg-slate-200 text-slate-700'
              }">
                ${escapeHtml(gps.status)}
              </span>
            </div>
            <p class="text-[11px] text-slate-700">
              <i class="fa-solid fa-location-dot text-rose-500 mr-1"></i>${escapeHtml(gps.location || '-')}
            </p>
          </div>
          <a
            href="${escapeHtml(gps.maps_url)}"
            target="_blank"
            rel="noopener noreferrer"
            class="px-2.5 py-1.5 rounded-lg bg-white hover:bg-emerald-100 text-emerald-800 border border-emerald-300 text-[11px] font-bold flex items-center gap-1 shrink-0 shadow-sm"
          >
            <i class="fa-solid fa-map-location-dot text-emerald-600"></i>
            <span>เปิดแผนที่</span>
          </a>
        </div>
      `;
    }

    showToast(`ดึงพิกัด GPS รถทะเบียน "${plate}" สำเร็จ`, 'success', `${gps.status} — ${gps.location}`);
    await loadAllData();
  } catch (err) {
    showToast(err.message || 'เกิดข้อผิดพลาดในการเช็กพิกัด GPS', 'error');
  }
}

async function refreshSingleVehicleGps(plate) {
  if (!plate) return;
  try {
    const res = await apiFetch(`/api/gps/vehicle/${encodeURIComponent(plate)}`);
    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถรีเฟรชพิกัด GPS ได้', 'error');
      return;
    }
    const gps = res.data;
    showToast(`อัปเดตพิกัดรถ "${plate}" ล่าสุดแล้ว`, 'success', `${gps.status} • ${gps.location}`);
    await loadAllData();
  } catch (err) {
    showToast(err.message || 'เกิดข้อผิดพลาดในการรีเฟรชพิกัด', 'error');
  }
}

async function syncAllVehiclesGps() {
  const icon = document.getElementById('sync-all-gps-icon');
  icon?.classList.add('fa-spin');
  try {
    const res = await apiFetch('/api/gps/sync-all', { method: 'POST' });
    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถอัปเดตพิกัด GPS ทั้งหมดได้', 'error');
      return;
    }
    showToast(res.data?.message || 'อัปเดตพิกัด GPS รถทุกคันสำเร็จ', 'success');
    await loadAllData();
  } catch (err) {
    showToast(err.message || 'เกิดข้อผิดพลาดในการซิงค์พิกัด GPS', 'error');
  } finally {
    icon?.classList.remove('fa-spin');
  }
}

// ============================================================================
// CREATE / EDIT MODE & FORM SUBMISSION WITH CUSTOM MODALS
// ============================================================================

function enterEditMode(workerName) {
  const log = state.logs.find((l) => l.worker_name === workerName);
  if (!log) {
    showToast(`ไม่พบข้อมูลการเดินทางของ "${workerName}"`, 'error');
    return;
  }

  state.editingWorkerName = log.worker_name;
  state.lastCheckedGps = parseGpsStatus(log.gps_status);
  renderWorkerDropdown();

  const workerSelect = document.getElementById('form-worker');
  workerSelect.value = log.worker_name;
  workerSelect.disabled = true;

  document.getElementById('form-date').value = log.travel_date || '';
  document.getElementById('form-route').value = log.route || '';
  document.getElementById('form-departure').value = log.departure_time || '';
  document.getElementById('form-arrival').value = log.arrival_time || '';

  renderVehicleDropdown();
  document.getElementById('form-vehicle').value = log.vehicle_plate || '';
  const passengersEl = document.getElementById('form-passengers');
  if (passengersEl) {
    passengersEl.value =
      log.passenger_count !== null && log.passenger_count !== undefined ? log.passenger_count : '';
  }

  const formCard = document.getElementById('form-card-container');
  const formHeaderBar = document.getElementById('form-header-bar');
  const modeIcon = document.getElementById('form-mode-icon');
  const modeTitle = document.getElementById('form-mode-title');
  const editBadgeGroup = document.getElementById('edit-mode-badge-group');
  const editWorkerLabel = document.getElementById('edit-mode-worker-label');
  const btnCancelEdit = document.getElementById('btn-cancel-edit');
  const btnSubmit = document.getElementById('btn-submit-log');
  const btnSubmitIcon = document.getElementById('btn-submit-icon');
  const btnSubmitText = document.getElementById('btn-submit-text');

  formCard.classList.add('ring-2', 'ring-amber-400');
  formHeaderBar.className =
    'px-5 sm:px-6 py-4 bg-amber-50/90 border-b border-amber-200 flex flex-wrap items-center justify-between gap-3';
  modeIcon.className =
    'w-8 h-8 rounded-lg bg-amber-600 text-white flex items-center justify-center text-sm shrink-0';
  modeIcon.innerHTML = '<i class="fa-solid fa-pen-to-square"></i>';
  modeTitle.textContent = `แก้ไขข้อมูลการเดินทางของ: ${log.worker_name}`;

  editWorkerLabel.textContent = log.worker_name;
  editBadgeGroup.classList.remove('hidden');
  editBadgeGroup.classList.add('flex');

  btnCancelEdit.classList.remove('hidden');
  btnCancelEdit.classList.add('flex');

  btnSubmit.className =
    'flex-1 sm:flex-initial min-h-[44px] px-6 py-2.5 rounded-xl bg-amber-600 hover:bg-amber-700 active:bg-amber-800 text-white text-sm font-semibold shadow-md shadow-amber-600/20 flex items-center justify-center gap-2 transition cursor-pointer';
  btnSubmitIcon.className = 'fa-solid fa-check-double';
  btnSubmitText.textContent = 'บันทึกการแก้ไขข้อมูล';

  updateDateAndDurationPreview();
  updateRouteSmartControls();
  renderLogsTable();

  formCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function exitEditModeAndResetForm() {
  state.editingWorkerName = null;
  state.lastCheckedGps = null;
  renderWorkerDropdown();

  const workerSelect = document.getElementById('form-worker');
  workerSelect.value = '';
  workerSelect.disabled = false;

  document.getElementById('form-date').value = getTodayISO();
  document.getElementById('form-route').value = '';
  document.getElementById('form-departure').value = '';
  document.getElementById('form-arrival').value = '';
  document.getElementById('form-vehicle').value = '';
  const passengersEl = document.getElementById('form-passengers');
  if (passengersEl) passengersEl.value = '';
  document.getElementById('form-gps-preview-card')?.classList.add('hidden');

  const formCard = document.getElementById('form-card-container');
  const formHeaderBar = document.getElementById('form-header-bar');
  const modeIcon = document.getElementById('form-mode-icon');
  const modeTitle = document.getElementById('form-mode-title');
  const editBadgeGroup = document.getElementById('edit-mode-badge-group');
  const btnCancelEdit = document.getElementById('btn-cancel-edit');
  const btnSubmit = document.getElementById('btn-submit-log');
  const btnSubmitIcon = document.getElementById('btn-submit-icon');
  const btnSubmitText = document.getElementById('btn-submit-text');

  formCard.classList.remove('ring-2', 'ring-amber-400');
  formHeaderBar.className =
    'px-5 sm:px-6 py-4 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3';
  modeIcon.className =
    'w-8 h-8 rounded-lg bg-zinc-900 text-amber-400 flex items-center justify-center text-sm shrink-0';
  modeIcon.innerHTML = '<i class="fa-solid fa-plus"></i>';
  modeTitle.textContent = 'บันทึกข้อมูลการเดินทางพนักงาน';

  editBadgeGroup.classList.add('hidden');
  editBadgeGroup.classList.remove('flex');

  btnCancelEdit.classList.add('hidden');
  btnCancelEdit.classList.remove('flex');

  btnSubmit.className =
    'flex-1 sm:flex-initial min-h-[44px] px-6 py-2.5 rounded-xl bg-zinc-900 hover:bg-zinc-800 active:bg-black text-white text-sm font-semibold shadow-md shadow-zinc-900/20 flex items-center justify-center gap-2 transition cursor-pointer';
  btnSubmitIcon.className = 'fa-solid fa-floppy-disk text-amber-400';
  btnSubmitText.textContent = 'บันทึกข้อมูลการเดินทาง';

  updateDateAndDurationPreview();
  updateRouteSmartControls();
  renderLogsTable();
}

function showDuplicateWorkerModal(workerName, serverMessage) {
  openCustomModal(`
    <div class="p-6">
      <div class="flex items-start gap-4">
        <div class="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center text-xl shrink-0">
          <i class="fa-solid fa-triangle-exclamation"></i>
        </div>
        <div class="flex-1">
          <div class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-rose-50 text-rose-700 border border-rose-200 text-[11px] font-bold mb-1.5">
            HTTP 409 Conflict — Primary Key ซ้ำ
          </div>
          <h3 class="text-lg font-bold text-slate-900">
            ไม่สามารถบันทึกซ้ำได้: "${escapeHtml(workerName)}"
          </h3>
          <p class="text-sm text-slate-600 mt-1 leading-relaxed">
            ${escapeHtml(
              serverMessage ||
                `รายชื่อคนงาน "${workerName}" มีข้อมูลการเดินทางอยู่ในฐานข้อมูลแล้ว กรุณากดปุ่ม "แก้ไข" แทนการบันทึกซ้ำ`
            )}
          </p>
        </div>
      </div>

      <div class="mt-6 pt-4 border-t border-slate-100 flex flex-wrap items-center justify-end gap-2.5">
        <button
          type="button"
          id="modal-btn-close-dup"
          class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-sm font-medium transition cursor-pointer"
        >
          ปิดหน้าต่าง
        </button>
        <button
          type="button"
          id="modal-btn-switch-edit"
          class="min-h-[42px] px-4 py-2 rounded-xl bg-amber-600 hover:bg-amber-700 text-white text-sm font-semibold shadow-sm flex items-center gap-1.5 transition cursor-pointer"
        >
          <i class="fa-solid fa-pen-to-square"></i>
          <span>สลับไปโหมดแก้ไขของ "${escapeHtml(workerName)}" ทันที</span>
        </button>
      </div>
    </div>
  `);

  document.getElementById('modal-btn-close-dup')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-btn-switch-edit')?.addEventListener('click', () => {
    closeCustomModal();
    enterEditMode(workerName);
  });
}

function showIncompleteFieldsModal(missingFields, payload) {
  const listItemsHtml = missingFields
    .map(
      (f) => `
      <li class="flex items-center gap-2 text-sm text-amber-900 py-1">
        <i class="fa-solid fa-circle-dot text-amber-500 text-[10px]"></i>
        <span class="font-medium">${escapeHtml(f)}</span>
      </li>
    `
    )
    .join('');

  openCustomModal(`
    <div class="p-6">
      <div class="flex items-start gap-4">
        <div class="w-12 h-12 rounded-2xl bg-amber-100 text-amber-600 flex items-center justify-center text-xl shrink-0">
          <i class="fa-solid fa-clipboard-question"></i>
        </div>
        <div class="flex-1">
          <h3 class="text-lg font-bold text-slate-900">
            แจ้งเตือน: กรอกข้อมูลยังไม่ครบทุกช่อง
          </h3>
          <p class="text-sm text-slate-600 mt-1">
            คุณกำลังจะบันทึกข้อมูลของ <strong>"${escapeHtml(payload.worker_name)}"</strong> โดยยังไม่ได้ระบุข้อมูลในช่องต่อไปนี้:
          </p>
          <ul class="mt-3 p-3.5 rounded-xl bg-amber-50/90 border border-amber-200 divide-y divide-amber-200/50">
            ${listItemsHtml}
          </ul>
          <p class="text-xs text-slate-500 mt-3">
            คุณต้องการยืนยันการบันทึกข้อมูลลงฐานข้อมูลโดยเว้นว่างช่องเหล่านี้หรือไม่?
          </p>
        </div>
      </div>

      <div class="mt-6 pt-4 border-t border-slate-100 flex flex-wrap items-center justify-end gap-2.5">
        <button
          type="button"
          id="modal-btn-back-to-fill"
          class="min-h-[42px] px-4 py-2.5 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-sm font-medium transition cursor-pointer"
        >
          กลับไปกรอกข้อมูล
        </button>
        <button
          type="button"
          id="modal-btn-confirm-save"
          class="min-h-[42px] px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold shadow-md shadow-indigo-600/20 flex items-center gap-2 transition cursor-pointer"
        >
          <i class="fa-solid fa-check"></i>
          <span>ยืนยันการบันทึก</span>
        </button>
      </div>
    </div>
  `);

  document.getElementById('modal-btn-back-to-fill')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-btn-confirm-save')?.addEventListener('click', async () => {
    closeCustomModal();
    await executeSaveLog(payload);
  });
}

async function executeSaveLog(payload) {
  try {
    const isEdit = Boolean(state.editingWorkerName);
    const url = isEdit
      ? `/api/logs/${encodeURIComponent(state.editingWorkerName)}`
      : '/api/logs';
    const method = isEdit ? 'PUT' : 'POST';

    const res = await apiFetch(url, {
      method,
      body: JSON.stringify(payload)
    });

    // HTTP 409 Conflict (Duplicate Primary Key)
    if (res.status === 409) {
      showDuplicateWorkerModal(payload.worker_name, res.data?.error);
      return;
    }

    // HTTP 507 Insufficient Storage (Storage Quota Full)
    if (res.status === 507) {
      if (res.data?.storageQuota) {
        state.storageQuota = res.data.storageQuota;
      }
      renderStorageQuotaUI(false);
      showStorageFullEmergencyModal(res.data?.error);
      return;
    }

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถบันทึกข้อมูลได้', 'error');
      return;
    }

    showToast(
      isEdit
        ? `อัปเดตข้อมูลการเดินทางของ "${payload.worker_name}" สำเร็จ`
        : `บันทึกข้อมูลการเดินทางของ "${payload.worker_name}" สำเร็จ`,
      'success'
    );

    exitEditModeAndResetForm();
    await loadAllData();
  } catch (err) {
    showToast(err.message || 'เกิดข้อผิดพลาดในการบันทึกข้อมูล', 'error');
  }
}

async function handleTravelLogSubmit(e) {
  e.preventDefault();

  const worker_name = (
    state.editingWorkerName ||
    document.getElementById('form-worker')?.value ||
    ''
  ).trim();
  const travel_date = (document.getElementById('form-date')?.value || '').trim();
  const route = (document.getElementById('form-route')?.value || '').trim();
  const departure_time = (document.getElementById('form-departure')?.value || '').trim();
  const arrival_time = (document.getElementById('form-arrival')?.value || '').trim();
  const vehicle_plate = (document.getElementById('form-vehicle')?.value || '').trim();
  const passenger_count_raw = (document.getElementById('form-passengers')?.value || '').trim();
  const passenger_count =
    passenger_count_raw === '' ? null : Math.max(0, Number.parseInt(passenger_count_raw, 10));

  if (!worker_name) {
    openCustomModal(`
      <div class="p-6">
        <div class="flex items-start gap-4">
          <div class="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center text-xl shrink-0">
            <i class="fa-solid fa-user-xmark"></i>
          </div>
          <div>
            <h3 class="text-lg font-bold text-slate-900">กรุณาเลือกชื่อคนงาน (Primary Key)</h3>
            <p class="text-sm text-slate-600 mt-1">
              ช่อง <strong>"ชื่อคนงาน (PK)"</strong> เป็นข้อมูลบังคับเลือกก่อนทำการบันทึก
            </p>
          </div>
        </div>
        <div class="mt-6 pt-4 border-t border-slate-100 flex justify-end">
          <button
            type="button"
            id="modal-btn-ok-worker"
            class="min-h-[42px] px-5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold cursor-pointer"
          >
            รับทราบ
          </button>
        </div>
      </div>
    `);
    document.getElementById('modal-btn-ok-worker')?.addEventListener('click', () => {
      closeCustomModal();
      document.getElementById('form-worker')?.focus();
    });
    return;
  }

  const payload = {
    worker_name,
    travel_date,
    route,
    departure_time,
    arrival_time,
    vehicle_plate,
    passenger_count: Number.isNaN(passenger_count) ? null : passenger_count
  };

  // If in Create mode and duplicate PK or Storage Full, send directly so backend returns 409 or 507 immediately
  if (
    !state.editingWorkerName &&
    (state.logs.some((l) => l.worker_name === worker_name) || state.storageQuota?.isFull)
  ) {
    await executeSaveLog(payload);
    return;
  }

  // Check incomplete optional fields
  const missingFields = [];
  if (!travel_date) missingFields.push('วันที่หมายที่ต้องไป');
  if (!route) missingFields.push('เส้นทาง / ที่หมาย');
  if (!departure_time) missingFields.push('เวลาขับรถ (ออกเดินทาง)');
  if (!arrival_time) missingFields.push('เวลาที่ต้องไปถึงไซต์งาน');
  if (!vehicle_plate) missingFields.push('ทะเบียนรถ');
  if (passenger_count_raw === '') missingFields.push('จำนวนคนโดยสาร');

  if (missingFields.length > 0) {
    showIncompleteFieldsModal(missingFields, payload);
    return;
  }

  await executeSaveLog(payload);
}

// ============================================================================
// TRAVEL LOGS RENDERING (Desktop Table + Mobile Responsive Cards) & EXPORTS
// ============================================================================

function getFilteredLogs() {
  const dateFilter = state.filterDate.trim();
  const searchFilter = state.filterSearch.trim().toLowerCase();

  return state.logs.filter((row) => {
    if (dateFilter && (row.travel_date || '') !== dateFilter) {
      return false;
    }
    if (searchFilter) {
      const thaiDateShort = formatThaiDate(row.travel_date, false).toLowerCase();
      const thaiDateFull = formatThaiDate(row.travel_date, true).toLowerCase();
      const gpsObj = parseGpsStatus(row.gps_status);
      const haystack = [
        row.worker_name || '',
        row.travel_date || '',
        thaiDateShort,
        thaiDateFull,
        row.route || '',
        row.vehicle_plate || '',
        row.passenger_count !== null && row.passenger_count !== undefined
          ? `${row.passenger_count} คน`
          : '',
        row.departure_time || '',
        row.arrival_time || ''
      ]
        .join(' ')
        .toLowerCase();
      if (!haystack.includes(searchFilter)) {
        return false;
      }
    }
    return true;
  });
}

function buildGpsCellHtml(row) {
  if (!row.vehicle_plate) {
    return '<span class="text-xs text-slate-400 italic">ไม่ได้ระบุรถ</span>';
  }

  const gps = parseGpsStatus(row.gps_status);
  if (!gps) {
    return `
      <div class="flex items-center gap-1.5">
        <span class="text-xs text-slate-400">ยังไม่มีพิกัด</span>
        <button
          type="button"
          data-refresh-gps="${escapeHtml(row.vehicle_plate)}"
          class="px-2 py-1 rounded-lg bg-slate-100 hover:bg-indigo-50 text-indigo-600 text-xs font-medium cursor-pointer"
        >
          <i class="fa-solid fa-rotate"></i> ดึงพิกัด
        </button>
      </div>
    `;
  }

  const isMoving = Number(gps.speed) > 0 || String(gps.status).includes('กำลังวิ่ง');
  const badgeClass = isMoving
    ? 'bg-emerald-100 text-emerald-800 border-emerald-300'
    : gps.engine === 'ON'
    ? 'bg-amber-100 text-amber-800 border-amber-300'
    : 'bg-slate-100 text-slate-700 border-slate-300';

  return `
    <div class="space-y-1 max-w-[250px]">
      <div class="flex flex-wrap items-center gap-1.5">
        <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full border text-[11px] font-bold ${badgeClass}">
          <i class="fa-solid ${isMoving ? 'fa-gauge-high text-emerald-600' : 'fa-circle-pause'} text-[10px]"></i>
          <span>${escapeHtml(gps.status)}</span>
        </span>

        ${
          gps.maps_url
            ? `<a
                href="${escapeHtml(gps.maps_url)}"
                target="_blank"
                rel="noopener noreferrer"
                title="เปิด Google Maps นำทาง/ดูพิกัด"
                class="inline-flex items-center gap-1 px-2 py-0.5 rounded-lg bg-sky-50 hover:bg-sky-100 text-sky-700 border border-sky-200 text-[11px] font-semibold transition"
              >
                <i class="fa-solid fa-map-location-dot"></i>
                <span>แผนที่</span>
              </a>`
            : ''
        }

        <button
          type="button"
          data-refresh-gps="${escapeHtml(row.vehicle_plate)}"
          title="รีเฟรชพิกัด GPS รถคันนี้"
          class="p-1 rounded-lg bg-slate-100 hover:bg-indigo-100 text-slate-600 hover:text-indigo-700 text-[11px] transition cursor-pointer"
        >
          <i class="fa-solid fa-rotate"></i>
        </button>
      </div>
      ${
        gps.location
          ? `<p class="text-[11px] text-slate-500 truncate" title="${escapeHtml(gps.location)}">
              <i class="fa-solid fa-location-dot text-rose-500 mr-1 text-[10px]"></i>${escapeHtml(gps.location)}
            </p>`
          : ''
      }
    </div>
  `;
}

function renderLogsTable() {
  const tbody = document.getElementById('logs-table-body');
  const mobileCards = document.getElementById('logs-mobile-cards');
  const countBadge = document.getElementById('table-count-badge');
  if (!tbody) return;

  const filtered = getFilteredLogs();
  if (countBadge) {
    countBadge.textContent = `${filtered.length} รายการ`;
  }

  if (filtered.length === 0) {
    const emptyHtml = `
      <div class="py-10 px-4 text-center">
        <div class="max-w-sm mx-auto space-y-2">
          <div class="w-12 h-12 rounded-2xl bg-slate-100 text-slate-400 flex items-center justify-center text-xl mx-auto">
            <i class="fa-solid fa-folder-open"></i>
          </div>
          <p class="text-sm font-semibold text-slate-700">ไม่พบรายการบันทึกข้อมูลการเดินทาง</p>
          <p class="text-xs text-slate-500">
            ${
              state.filterDate || state.filterSearch
                ? 'ลองเปลี่ยนตัวกรองวันที่หรือคำค้นหาเพื่อดูรายการทั้งหมด'
                : 'เริ่มต้นบันทึกข้อมูลการเดินทางของคนงานจากฟอร์มด้านบน'
            }
          </p>
        </div>
      </div>
    `;
    tbody.innerHTML = `<tr><td colspan="8">${emptyHtml}</td></tr>`;
    if (mobileCards) mobileCards.innerHTML = emptyHtml;
    return;
  }

  const workerDefaultMap = new Map(
    state.allWorkers.map((w) => [w.name, (w.default_route || '').trim()])
  );

  // 1. Render Desktop Table Rows
  tbody.innerHTML = filtered
    .map((row) => {
      const isEditingThis = state.editingWorkerName === row.worker_name;
      const defaultRoute = workerDefaultMap.get(row.worker_name) || '';
      const isDefaultRoute = defaultRoute && row.route && row.route.trim() === defaultRoute;
      const duration = calculateTravelDuration(row.departure_time, row.arrival_time);

      return `
        <tr class="${
          isEditingThis
            ? 'bg-amber-50/90 ring-1 ring-inset ring-amber-300'
            : 'hover:bg-slate-50/90'
        } transition">
          <td class="py-3.5 px-4 font-semibold text-slate-900 whitespace-nowrap">
            <div class="flex items-center gap-2.5">
              <div class="w-8 h-8 rounded-xl ${
                isEditingThis ? 'bg-amber-600 text-white' : 'bg-indigo-100 text-indigo-700'
              } flex items-center justify-center text-xs font-bold shrink-0">
                ${escapeHtml((row.worker_name || '?').slice(0, 1))}
              </div>
              <div>
                <div class="flex items-center gap-1.5">
                  <span>${escapeHtml(row.worker_name)}</span>
                  <span class="px-1.5 py-0.2 rounded bg-slate-100 text-slate-500 font-mono text-[10px] border border-slate-200">PK</span>
                </div>
                ${
                  defaultRoute
                    ? `<p class="text-[11px] text-slate-400 font-normal">ประจำ: ${escapeHtml(defaultRoute)}</p>`
                    : ''
                }
              </div>
            </div>
          </td>

          <td class="py-3.5 px-4 whitespace-nowrap">
            ${
              row.travel_date
                ? `
                  <div class="font-medium text-slate-800">${escapeHtml(
                    formatThaiDate(row.travel_date, false)
                  )}</div>
                  <div class="text-[11px] text-slate-400 font-mono">${escapeHtml(row.travel_date)}</div>
                `
                : '<span class="text-xs text-slate-400 italic">ไม่ได้ระบุ</span>'
            }
          </td>

          <td class="py-3.5 px-4">
            ${
              row.route
                ? `
                  <div class="flex flex-wrap items-center gap-1.5">
                    <span class="font-medium text-slate-800">${escapeHtml(row.route)}</span>
                    ${
                      isDefaultRoute
                        ? '<span class="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-semibold">ไซต์งานประจำ</span>'
                        : ''
                    }
                  </div>
                `
                : '<span class="text-xs text-slate-400 italic">ไม่ได้ระบุ</span>'
            }
          </td>

          <td class="py-3.5 px-3 text-center whitespace-nowrap">
            <div class="inline-flex items-center gap-1 text-xs font-mono">
              <span class="px-2 py-0.5 rounded bg-slate-100 text-slate-800 font-semibold">${escapeHtml(
                row.departure_time || '-'
              )}</span>
              <i class="fa-solid fa-arrow-right text-[10px] text-slate-400"></i>
              <span class="px-2 py-0.5 rounded bg-indigo-50 text-indigo-800 font-semibold">${escapeHtml(
                row.arrival_time || '-'
              )}</span>
            </div>
          </td>

          <td class="py-3.5 px-3 text-center whitespace-nowrap">
            ${
              duration.totalMinutes !== null
                ? `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold ${
                    duration.isOvernight
                      ? 'bg-amber-100 text-amber-900 border border-amber-300'
                      : 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                  }">
                    <i class="fa-solid ${duration.isOvernight ? 'fa-moon' : 'fa-stopwatch'} text-[10px]"></i>
                    <span>${escapeHtml(duration.text)}</span>
                  </span>`
                : '<span class="text-xs text-slate-400">-</span>'
            }
          </td>

          <td class="py-3.5 px-4 whitespace-nowrap">
            ${
              row.vehicle_plate
                ? `<span class="license-plate-badge"><i class="fa-solid fa-truck text-indigo-600 text-[11px]"></i>${escapeHtml(
                    row.vehicle_plate
                  )}</span>`
                : '<span class="text-xs text-slate-400 italic">ไม่ได้ระบุ</span>'
            }
          </td>

          <td class="py-3.5 px-3 text-center whitespace-nowrap">
            ${
              row.passenger_count !== null && row.passenger_count !== undefined
                ? `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-sky-50 text-sky-800 border border-sky-200 text-xs font-semibold">
                    <i class="fa-solid fa-users text-[10px]"></i>
                    <span>${escapeHtml(row.passenger_count)} คน</span>
                  </span>`
                : '<span class="text-xs text-slate-400 italic">ไม่ได้ระบุ</span>'
            }
          </td>

          <td class="py-3.5 px-4 text-right whitespace-nowrap">
            <div class="inline-flex items-center gap-1.5">
              <button
                type="button"
                data-action="edit-log"
                data-worker="${escapeHtml(row.worker_name)}"
                class="px-2.5 py-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-semibold inline-flex items-center gap-1 transition cursor-pointer"
              >
                <i class="fa-solid fa-pen-to-square"></i>
                <span>แก้ไข</span>
              </button>
              <button
                type="button"
                data-action="delete-log"
                data-worker="${escapeHtml(row.worker_name)}"
                class="px-2.5 py-1.5 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-semibold inline-flex items-center gap-1 transition cursor-pointer"
              >
                <i class="fa-solid fa-trash-can"></i>
                <span>ลบ</span>
              </button>
            </div>
          </td>
        </tr>
      `;
    })
    .join('');

  // 2. Render Mobile Responsive Cards (Mobile-First Touch UI)
  if (mobileCards) {
    mobileCards.innerHTML = filtered
      .map((row) => {
        const defaultRoute = workerDefaultMap.get(row.worker_name) || '';
        const isDefaultRoute = defaultRoute && row.route && row.route.trim() === defaultRoute;
        const duration = calculateTravelDuration(row.departure_time, row.arrival_time);
        const gps = parseGpsStatus(row.gps_status);
        const isMoving = gps && (Number(gps.speed) > 0 || String(gps.status).includes('กำลังวิ่ง'));

        return `
          <div class="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm space-y-3">
            <div class="flex items-start justify-between gap-2">
              <div class="flex items-center gap-2.5">
                <div class="w-9 h-9 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold text-sm shrink-0">
                  ${escapeHtml((row.worker_name || '?').slice(0, 1))}
                </div>
                <div>
                  <h4 class="font-bold text-slate-900 text-sm flex items-center gap-1.5">
                    <span>${escapeHtml(row.worker_name)}</span>
                    <span class="px-1.5 py-0.2 rounded bg-slate-100 text-slate-500 font-mono text-[10px]">PK</span>
                  </h4>
                  <p class="text-xs text-indigo-600 font-medium">
                    <i class="fa-regular fa-calendar mr-1"></i>${escapeHtml(
                      formatThaiDate(row.travel_date, false)
                    )}
                  </p>
                </div>
              </div>

              ${
                row.vehicle_plate
                  ? `<span class="license-plate-badge"><i class="fa-solid fa-truck text-indigo-600 text-[11px]"></i>${escapeHtml(
                      row.vehicle_plate
                    )}</span>`
                  : '<span class="text-xs text-slate-400">ไม่ระบุรถ</span>'
              }
            </div>

            <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-100 space-y-1.5 text-xs">
              <div class="flex items-start justify-between gap-2">
                <span class="text-slate-500 shrink-0">เส้นทาง/ที่หมาย:</span>
                <span class="font-semibold text-slate-800 text-right">
                  ${escapeHtml(row.route || '-')}
                  ${
                    isDefaultRoute
                      ? '<span class="ml-1 px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 text-[10px]">ไซต์ประจำ</span>'
                      : ''
                  }
                </span>
              </div>
              <div class="flex items-center justify-between gap-2">
                <span class="text-slate-500">เวลาเดินทาง:</span>
                <span class="font-mono font-semibold text-slate-800">
                  ${escapeHtml(row.departure_time || '-')} <i class="fa-solid fa-arrow-right text-[10px] text-slate-400 mx-1"></i> ${escapeHtml(
          row.arrival_time || '-'
        )}
                  <span class="ml-1 text-indigo-600">(${escapeHtml(duration.text)})</span>
                </span>
              </div>
              <div class="flex items-center justify-between gap-2">
                <span class="text-slate-500">จำนวนคนโดยสาร:</span>
                <span class="font-semibold text-sky-800">
                  ${
                    row.passenger_count !== null && row.passenger_count !== undefined
                      ? `${escapeHtml(row.passenger_count)} คน`
                      : '<span class="text-slate-400 font-normal italic">ไม่ได้ระบุ</span>'
                  }
                </span>
              </div>
            </div>

            <div class="flex items-center gap-2 pt-1">
              <button
                type="button"
                data-action="edit-log"
                data-worker="${escapeHtml(row.worker_name)}"
                class="flex-1 min-h-[44px] rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-200 font-semibold text-xs flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <i class="fa-solid fa-pen-to-square"></i>
                <span>แก้ไขข้อมูล</span>
              </button>
              <button
                type="button"
                data-action="delete-log"
                data-worker="${escapeHtml(row.worker_name)}"
                class="flex-1 min-h-[44px] rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-semibold text-xs flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <i class="fa-solid fa-trash-can"></i>
                <span>ลบรายการ</span>
              </button>
            </div>
          </div>
        `;
      })
      .join('');
  }

  // Bind Edit, Delete & Refresh GPS buttons for both Desktop & Mobile
  document.querySelectorAll('button[data-action="edit-log"]').forEach((btn) => {
    btn.addEventListener('click', () => {
      enterEditMode(btn.getAttribute('data-worker'));
    });
  });

  document.querySelectorAll('button[data-action="delete-log"]').forEach((btn) => {
    btn.addEventListener('click', () => {
      confirmDeleteTravelLog(btn.getAttribute('data-worker'));
    });
  });

  document.querySelectorAll('button[data-refresh-gps]').forEach((btn) => {
    btn.addEventListener('click', () => {
      refreshSingleVehicleGps(btn.getAttribute('data-refresh-gps'));
    });
  });
}

function confirmDeleteTravelLog(workerName) {
  openCustomModal(`
    <div class="p-6">
      <div class="flex items-start gap-4">
        <div class="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center text-xl shrink-0">
          <i class="fa-solid fa-trash-can"></i>
        </div>
        <div class="flex-1">
          <h3 class="text-lg font-bold text-slate-900">ยืนยันการลบข้อมูลการเดินทาง</h3>
          <p class="text-sm text-slate-600 mt-1">
            คุณต้องการลบข้อมูลการเดินทางของคนงาน <strong>"${escapeHtml(workerName)}"</strong> ออกจากฐานข้อมูลใช่หรือไม่?
          </p>
        </div>
      </div>

      <div class="mt-6 pt-4 border-t border-slate-100 flex justify-end gap-2.5">
        <button
          type="button"
          id="modal-btn-cancel-del-log"
          class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-sm font-medium cursor-pointer"
        >
          ยกเลิก
        </button>
        <button
          type="button"
          id="modal-btn-confirm-del-log"
          class="min-h-[42px] px-5 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-sm font-semibold shadow-sm flex items-center gap-1.5 cursor-pointer"
        >
          <i class="fa-solid fa-trash-can"></i>
          <span>ยืนยันการลบ</span>
        </button>
      </div>
    </div>
  `);

  document.getElementById('modal-btn-cancel-del-log')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-btn-confirm-del-log')?.addEventListener('click', async () => {
    closeCustomModal();
    try {
      const res = await apiFetch(`/api/logs/${encodeURIComponent(workerName)}`, {
        method: 'DELETE'
      });
      if (!res.ok) {
        showToast(res.data?.error || 'ไม่สามารถลบข้อมูลได้', 'error');
        return;
      }
      if (state.editingWorkerName === workerName) {
        exitEditModeAndResetForm();
      }
      showToast(`ลบข้อมูลการเดินทางของ "${workerName}" เรียบร้อยแล้ว`, 'success');
      await loadAllData();
    } catch (err) {
      showToast(err.message || 'เกิดข้อผิดพลาดในการลบข้อมูล', 'error');
    }
  });
}

// Export 1: Native Microsoft Excel (.xlsx)
function exportLogsToExcel() {
  const rows = getFilteredLogs();
  if (rows.length === 0) {
    showToast('ไม่มีรายการข้อมูลให้ส่งออกไฟล์ Excel', 'warning');
    return;
  }

  const dateSuffix = state.filterDate || getTodayISO();
  const filename = `worker_travel_logs_${dateSuffix}.xlsx`;

  const excelData = rows.map((r, idx) => {
    const dur = calculateTravelDuration(r.departure_time, r.arrival_time);
    return {
      'ลำดับ': idx + 1,
      'ชื่อคนงาน (PK)': r.worker_name || '',
      'วันที่หมายที่ต้องไป (YYYY-MM-DD)': r.travel_date || '',
      'วันที่หมายที่ต้องไป (ไทย)': formatThaiDate(r.travel_date, false),
      'เส้นทาง / ที่หมาย': r.route || '',
      'เวลาขับรถ': r.departure_time || '-',
      'เวลาที่ต้องไปถึงไซต์งาน': r.arrival_time || '-',
      'ระยะเวลาเดินทาง': dur.text,
      'ทะเบียนรถ': r.vehicle_plate || '-',
      'จำนวนคนโดยสาร (คน)':
        r.passenger_count !== null && r.passenger_count !== undefined ? r.passenger_count : '-',
      'อัปเดตล่าสุด': r.updated_at || ''
    };
  });

  if (typeof window.XLSX !== 'undefined') {
    const worksheet = window.XLSX.utils.json_to_sheet(excelData);
    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 24 },
      { wch: 24 },
      { wch: 20 },
      { wch: 34 },
      { wch: 14 },
      { wch: 20 },
      { wch: 22 },
      { wch: 18 },
      { wch: 20 },
      { wch: 20 }
    ];
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, 'บันทึกการเดินทาง');
    window.XLSX.writeFile(workbook, filename);

    showToast(
      `ส่งออกไฟล์ Excel (.xlsx) สำเร็จ (${rows.length} รายการ)`,
      'success',
      `ดาวน์โหลดไฟล์ ${filename} เรียบร้อยแล้ว`
    );
    return;
  }

  exportLogsToCSV();
}

// Export 2: CSV with UTF-8 BOM (\uFEFF)
function exportLogsToCSV() {
  const rows = getFilteredLogs();
  if (rows.length === 0) {
    showToast('ไม่มีรายการข้อมูลให้ส่งออกไฟล์ CSV', 'warning');
    return;
  }

  const headers = [
    'ลำดับ',
    'ชื่อคนงาน (PK)',
    'วันที่หมายที่ต้องไป (YYYY-MM-DD)',
    'วันที่หมายที่ต้องไป (ไทย)',
    'เส้นทาง / ที่หมาย',
    'เวลาขับรถ',
    'เวลาที่ต้องไปถึงไซต์งาน',
    'ระยะเวลาเดินทาง',
    'ทะเบียนรถ',
    'จำนวนคนโดยสาร (คน)',
    'อัปเดตล่าสุด'
  ];

  const escapeCsvCell = (val) => {
    const str = String(val ?? '');
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };

  const csvLines = [
    headers.map(escapeCsvCell).join(','),
    ...rows.map((r, idx) => {
      const dur = calculateTravelDuration(r.departure_time, r.arrival_time);
      return [
        idx + 1,
        r.worker_name || '',
        r.travel_date || '',
        formatThaiDate(r.travel_date, false),
        r.route || '',
        r.departure_time || '',
        r.arrival_time || '',
        dur.text,
        r.vehicle_plate || '',
        r.passenger_count !== null && r.passenger_count !== undefined ? r.passenger_count : '',
        r.updated_at || ''
      ]
        .map(escapeCsvCell)
        .join(',');
    })
  ];

  const csvContent = '\uFEFF' + csvLines.join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  const dateSuffix = state.filterDate || getTodayISO();
  link.href = url;
  link.setAttribute('download', `worker_travel_logs_${dateSuffix}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  showToast(
    `ส่งออกไฟล์ CSV สำเร็จ (${rows.length} รายการ)`,
    'success',
    'เข้ารหัส UTF-8 with BOM รองรับ Microsoft Excel และแอปสเปรดชีตบนมือถือ'
  );
}

// ============================================================================
// BACK-OFFICE MANAGEMENT (4 SECTIONS - ADMIN ONLY)
// ============================================================================

function populateGpsSettingsForm() {
  const s = state.gpsSettings;
  if (!s) return;

  const mockToggle = document.getElementById('setting-mock-mode');
  const maxLimitInput = document.getElementById('setting-max-limit');
  const apiUrlInput = document.getElementById('setting-api-url');
  const apiUserInput = document.getElementById('setting-api-username');
  const apiKeyInput = document.getElementById('setting-api-key');
  const modePill = document.getElementById('gps-mode-status-pill');

  const isMock = String(s.cartrack_mock_mode).toLowerCase() !== 'false';
  if (mockToggle) mockToggle.checked = isMock;
  if (maxLimitInput) maxLimitInput.value = s.max_storage_limit || '100';
  if (apiUrlInput) apiUrlInput.value = s.cartrack_api_url || '';
  if (apiUserInput) apiUserInput.value = s.cartrack_username || '';
  if (apiKeyInput) apiKeyInput.value = s.cartrack_api_key || '';

  if (modePill) {
    modePill.className = isMock
      ? 'px-3 py-1 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-300'
      : 'px-3 py-1 rounded-full text-xs font-bold bg-sky-100 text-sky-800 border border-sky-300';
    modePill.textContent = isMock
      ? 'โหมดจำลองข้อมูล (Mock GPS Mode)'
      : 'โหมดเชื่อมต่อจริง (Cartrack Live API)';
  }
}

function renderBackOfficeLists() {
  if (state.user?.role !== 'admin') return;
  populateGpsSettingsForm();

  // 1. Workers
  const workersListEl = document.getElementById('master-workers-list');
  const badgeWorkers = document.getElementById('badge-count-workers');
  if (badgeWorkers) badgeWorkers.textContent = `${state.allWorkers.length} คน`;

  if (workersListEl) {
    workersListEl.innerHTML = state.allWorkers
      .map(
        (w) => `
        <div class="p-3.5 hover:bg-slate-50 flex items-center justify-between gap-2 transition">
          <div class="min-w-0 flex-1">
            <div class="flex items-center gap-2">
              <span class="font-semibold text-sm ${
                w.active ? 'text-slate-900' : 'text-slate-400 line-through'
              } truncate">${escapeHtml(w.name)}</span>
              <span class="px-2 py-0.5 rounded-full text-[10px] font-semibold ${
                w.active ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'
              }">
                ${w.active ? 'เปิดใช้งาน' : 'ปิดใช้งาน'}
              </span>
            </div>
            <p class="text-xs text-slate-500 truncate mt-0.5">
              <i class="fa-solid fa-map-pin text-indigo-500 mr-1 text-[10px]"></i>
              ไซต์ประจำ: ${
                w.default_route
                  ? `<strong class="text-slate-700">${escapeHtml(w.default_route)}</strong>`
                  : '<span class="italic text-slate-400">ไม่ได้กำหนด</span>'
              }
            </p>
          </div>

          <div class="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              data-worker-toggle="${w.id}"
              class="px-2 py-1 rounded-lg border text-xs font-semibold transition cursor-pointer ${
                w.active
                  ? 'bg-emerald-50 border-emerald-300 text-emerald-700 hover:bg-emerald-100'
                  : 'bg-slate-100 border-slate-300 text-slate-500 hover:bg-slate-200'
              }"
            >
              <i class="fa-solid ${w.active ? 'fa-toggle-on text-sm' : 'fa-toggle-off text-sm'}"></i>
            </button>

            <button
              type="button"
              data-worker-edit="${w.id}"
              class="p-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-pen"></i>
            </button>

            <button
              type="button"
              data-worker-delete="${w.id}"
              class="p-1.5 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </div>
      `
      )
      .join('');

    workersListEl.querySelectorAll('[data-worker-toggle]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-worker-toggle'));
        const item = state.allWorkers.find((x) => x.id === id);
        if (item) toggleMasterItemActive('workers', item);
      });
    });

    workersListEl.querySelectorAll('[data-worker-edit]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-worker-edit'));
        const item = state.allWorkers.find((x) => x.id === id);
        if (item) openEditWorkerModal(item);
      });
    });

    workersListEl.querySelectorAll('[data-worker-delete]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-worker-delete'));
        const item = state.allWorkers.find((x) => x.id === id);
        if (item) confirmDeleteMasterItem('workers', item.id, item.name);
      });
    });
  }

  // 2. Routes
  const routesListEl = document.getElementById('master-routes-list');
  const badgeRoutes = document.getElementById('badge-count-routes');
  if (badgeRoutes) badgeRoutes.textContent = `${state.allRoutes.length} เส้นทาง`;

  if (routesListEl) {
    routesListEl.innerHTML = state.allRoutes
      .map((r) => {
        const workersUsingCount = state.allWorkers.filter(
          (w) => w.default_route === r.name
        ).length;
        return `
        <div class="p-3.5 hover:bg-slate-50 flex items-center justify-between gap-2 transition">
          <div class="min-w-0 flex-1">
            <div class="flex items-center gap-2">
              <span class="font-semibold text-sm ${
                r.active ? 'text-slate-900' : 'text-slate-400 line-through'
              } truncate">${escapeHtml(r.name)}</span>
              <span class="px-2 py-0.5 rounded-full text-[10px] font-semibold ${
                r.active ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'
              }">
                ${r.active ? 'เปิดใช้งาน' : 'ปิดใช้งาน'}
              </span>
            </div>
            <p class="text-[11px] text-slate-500 mt-0.5">
              เป็นไซต์ประจำของคนงาน: <strong>${workersUsingCount} คน</strong>
            </p>
          </div>

          <div class="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              data-route-toggle="${r.id}"
              class="px-2 py-1 rounded-lg border text-xs font-semibold transition cursor-pointer ${
                r.active
                  ? 'bg-emerald-50 border-emerald-300 text-emerald-700 hover:bg-emerald-100'
                  : 'bg-slate-100 border-slate-300 text-slate-500 hover:bg-slate-200'
              }"
            >
              <i class="fa-solid ${r.active ? 'fa-toggle-on text-sm' : 'fa-toggle-off text-sm'}"></i>
            </button>

            <button
              type="button"
              data-route-edit="${r.id}"
              class="p-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-pen"></i>
            </button>

            <button
              type="button"
              data-route-delete="${r.id}"
              class="p-1.5 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </div>
      `;
      })
      .join('');

    routesListEl.querySelectorAll('[data-route-toggle]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-route-toggle'));
        const item = state.allRoutes.find((x) => x.id === id);
        if (item) toggleMasterItemActive('routes', item);
      });
    });

    routesListEl.querySelectorAll('[data-route-edit]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-route-edit'));
        const item = state.allRoutes.find((x) => x.id === id);
        if (item) openEditRouteModal(item);
      });
    });

    routesListEl.querySelectorAll('[data-route-delete]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-route-delete'));
        const item = state.allRoutes.find((x) => x.id === id);
        if (item) confirmDeleteMasterItem('routes', item.id, item.name);
      });
    });
  }

  // 3. Vehicles
  const vehiclesListEl = document.getElementById('master-vehicles-list');
  const badgeVehicles = document.getElementById('badge-count-vehicles');
  if (badgeVehicles) badgeVehicles.textContent = `${state.allVehicles.length} คัน`;

  if (vehiclesListEl) {
    vehiclesListEl.innerHTML = state.allVehicles
      .map(
        (v) => `
        <div class="p-3.5 hover:bg-slate-50 flex items-center justify-between gap-2 transition">
          <div class="min-w-0 flex-1 flex items-center gap-2.5">
            <span class="license-plate-badge ${!v.active ? 'opacity-50 line-through' : ''}">
              <i class="fa-solid fa-truck text-amber-600 text-[11px]"></i>
              ${escapeHtml(v.plate)}
            </span>
            <span class="px-2 py-0.5 rounded-full text-[10px] font-semibold ${
              v.active ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'
            }">
              ${v.active ? 'เปิดใช้งาน' : 'ปิดใช้งาน'}
            </span>
          </div>

          <div class="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              data-vehicle-toggle="${v.id}"
              class="px-2 py-1 rounded-lg border text-xs font-semibold transition cursor-pointer ${
                v.active
                  ? 'bg-emerald-50 border-emerald-300 text-emerald-700 hover:bg-emerald-100'
                  : 'bg-slate-100 border-slate-300 text-slate-500 hover:bg-slate-200'
              }"
            >
              <i class="fa-solid ${v.active ? 'fa-toggle-on text-sm' : 'fa-toggle-off text-sm'}"></i>
            </button>

            <button
              type="button"
              data-vehicle-edit="${v.id}"
              class="p-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-pen"></i>
            </button>

            <button
              type="button"
              data-vehicle-delete="${v.id}"
              class="p-1.5 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </div>
      `
      )
      .join('');

    vehiclesListEl.querySelectorAll('[data-vehicle-toggle]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-vehicle-toggle'));
        const item = state.allVehicles.find((x) => x.id === id);
        if (item) toggleMasterItemActive('vehicles', item);
      });
    });

    vehiclesListEl.querySelectorAll('[data-vehicle-edit]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-vehicle-edit'));
        const item = state.allVehicles.find((x) => x.id === id);
        if (item) openEditVehicleModal(item);
      });
    });

    vehiclesListEl.querySelectorAll('[data-vehicle-delete]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-vehicle-delete'));
        const item = state.allVehicles.find((x) => x.id === id);
        if (item) confirmDeleteMasterItem('vehicles', item.id, item.plate);
      });
    });
  }
}

async function toggleMasterItemActive(type, item) {
  try {
    const newActive = !item.active;
    const res = await apiFetch(`/api/master/${type}/${item.id}`, {
      method: 'PUT',
      body: JSON.stringify({ active: newActive })
    });
    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถเปลี่ยนสถานะได้', 'error');
      return;
    }
    const label = item.name || item.plate;
    showToast(
      `${newActive ? 'เปิดใช้งาน' : 'ปิดใช้งาน'} "${label}" ในหน้า Dropdown แล้ว`,
      'info'
    );
    await loadAllData();
  } catch (err) {
    showToast(err.message || 'เกิดข้อผิดพลาด', 'error');
  }
}

function openEditWorkerModal(worker) {
  openCustomModal(`
    <form id="modal-form-edit-worker" class="p-6 space-y-4">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center">
          <i class="fa-solid fa-user-pen"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-slate-900">แก้ไขข้อมูลคนงาน</h3>
          <p class="text-xs text-slate-500">อัปเดตเชื่อมโยงกับตาราง <code>travel_logs</code> อัตโนมัติ</p>
        </div>
      </div>

      <div>
        <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อ-นามสกุลคนงาน *</label>
        <input
          type="text"
          id="modal-edit-worker-name"
          value="${escapeHtml(worker.name)}"
          required
          class="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600"
        />
      </div>

      <div>
        <label class="block text-xs font-semibold text-slate-700 mb-1">ไซต์งานประจำ (default_route)</label>
        <input
          type="text"
          id="modal-edit-worker-route"
          list="datalist-master-routes"
          value="${escapeHtml(worker.default_route || '')}"
          placeholder="เลือกจากรายการหรือพิมพ์เอง (เว้นว่างได้)"
          class="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-600"
        />
      </div>

      <div class="pt-3 border-t border-slate-100 flex justify-end gap-2">
        <button
          type="button"
          id="modal-btn-cancel"
          class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white text-slate-700 text-xs sm:text-sm font-medium cursor-pointer"
        >
          ยกเลิก
        </button>
        <button
          type="submit"
          class="min-h-[42px] px-5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold cursor-pointer"
        >
          บันทึกการแก้ไข
        </button>
      </div>
    </form>
  `);

  document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-form-edit-worker')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = document.getElementById('modal-edit-worker-name').value.trim();
    const default_route = document.getElementById('modal-edit-worker-route').value.trim();
    if (!name) return;

    const res = await apiFetch(`/api/master/workers/${worker.id}`, {
      method: 'PUT',
      body: JSON.stringify({ name, default_route })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถแก้ไขข้อมูลคนงานได้', 'error');
      return;
    }

    closeCustomModal();
    if (state.editingWorkerName === worker.name && name !== worker.name) {
      state.editingWorkerName = name;
    }
    showToast(
      `อัปเดตข้อมูลคนงาน "${name}" สำเร็จ`,
      'success',
      res.data?.cascadedLogsCount
        ? `อัปเดตเชื่อมโยงในตารางบันทึกการเดินทาง ${res.data.cascadedLogsCount} รายการ`
        : ''
    );
    await loadAllData();
  });
}

function openEditRouteModal(routeItem) {
  openCustomModal(`
    <form id="modal-form-edit-route" class="p-6 space-y-4">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
          <i class="fa-solid fa-map-location-dot"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-slate-900">แก้ไขชื่อเส้นทาง / ที่หมาย</h3>
          <p class="text-xs text-slate-500">อัปเดตเชื่อมโยงไปยัง <code>travel_logs</code> และ <code>default_route</code> อัตโนมัติ</p>
        </div>
      </div>

      <div>
        <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อเส้นทาง / ไซต์งาน *</label>
        <input
          type="text"
          id="modal-edit-route-name"
          value="${escapeHtml(routeItem.name)}"
          required
          class="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-600"
        />
      </div>

      <div class="pt-3 border-t border-slate-100 flex justify-end gap-2">
        <button
          type="button"
          id="modal-btn-cancel"
          class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white text-slate-700 text-xs sm:text-sm font-medium cursor-pointer"
        >
          ยกเลิก
        </button>
        <button
          type="submit"
          class="min-h-[42px] px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-semibold cursor-pointer"
        >
          บันทึกการแก้ไข
        </button>
      </div>
    </form>
  `);

  document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-form-edit-route')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = document.getElementById('modal-edit-route-name').value.trim();
    if (!name) return;

    const res = await apiFetch(`/api/master/routes/${routeItem.id}`, {
      method: 'PUT',
      body: JSON.stringify({ name })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถแก้ไขเส้นทางได้', 'error');
      return;
    }

    closeCustomModal();
    showToast(
      `อัปเดตเส้นทาง "${name}" สำเร็จ`,
      'success',
      `อัปเดตเชื่อมโยงบันทึกเดินทาง ${res.data?.cascadedLogsCount || 0} รายการ และไซต์ประจำคนงาน ${
        res.data?.cascadedWorkersCount || 0
      } คน`
    );
    await loadAllData();
  });
}

function openEditVehicleModal(vehicleItem) {
  openCustomModal(`
    <form id="modal-form-edit-vehicle" class="p-6 space-y-4">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center">
          <i class="fa-solid fa-truck-pickup"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-slate-900">แก้ไขหมายเลขทะเบียนรถ</h3>
          <p class="text-xs text-slate-500">อัปเดตเชื่อมโยงไปยังตาราง <code>travel_logs</code> อัตโนมัติ</p>
        </div>
      </div>

      <div>
        <label class="block text-xs font-semibold text-slate-700 mb-1">หมายเลขทะเบียนรถ *</label>
        <input
          type="text"
          id="modal-edit-vehicle-plate"
          value="${escapeHtml(vehicleItem.plate)}"
          required
          class="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-amber-600"
        />
      </div>

      <div class="pt-3 border-t border-slate-100 flex justify-end gap-2">
        <button
          type="button"
          id="modal-btn-cancel"
          class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white text-slate-700 text-xs sm:text-sm font-medium cursor-pointer"
        >
          ยกเลิก
        </button>
        <button
          type="submit"
          class="min-h-[42px] px-5 py-2 rounded-xl bg-amber-600 hover:bg-amber-700 text-white text-xs sm:text-sm font-semibold cursor-pointer"
        >
          บันทึกการแก้ไข
        </button>
      </div>
    </form>
  `);

  document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-form-edit-vehicle')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const plate = document.getElementById('modal-edit-vehicle-plate').value.trim();
    if (!plate) return;

    const res = await apiFetch(`/api/master/vehicles/${vehicleItem.id}`, {
      method: 'PUT',
      body: JSON.stringify({ plate })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถแก้ไขทะเบียนรถได้', 'error');
      return;
    }

    closeCustomModal();
    showToast(
      `อัปเดตทะเบียนรถ "${plate}" สำเร็จ`,
      'success',
      res.data?.cascadedLogsCount
        ? `อัปเดตเชื่อมโยงในตารางบันทึกการเดินทาง ${res.data.cascadedLogsCount} รายการ`
        : ''
    );
    await loadAllData();
  });
}

function confirmDeleteMasterItem(type, id, label) {
  const typeLabelMap = {
    workers: 'รายชื่อคนงาน',
    routes: 'เส้นทาง / ที่หมาย',
    vehicles: 'ทะเบียนรถ'
  };

  openCustomModal(`
    <div class="p-6">
      <div class="flex items-start gap-4">
        <div class="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center text-xl shrink-0">
          <i class="fa-solid fa-trash-can"></i>
        </div>
        <div>
          <h3 class="text-lg font-bold text-slate-900">ยืนยันการลบ${typeLabelMap[type] || 'รายการ'}</h3>
          <p class="text-sm text-slate-600 mt-1">
            คุณต้องการลบ <strong>"${escapeHtml(label)}"</strong> ออกจากรายการตัวเลือกหลักใช่หรือไม่?
          </p>
        </div>
      </div>

      <div class="mt-6 pt-4 border-t border-slate-100 flex justify-end gap-2">
        <button
          type="button"
          id="modal-btn-cancel"
          class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white text-slate-700 text-sm font-medium cursor-pointer"
        >
          ยกเลิก
        </button>
        <button
          type="button"
          id="modal-btn-confirm-del"
          class="min-h-[42px] px-5 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white text-sm font-semibold cursor-pointer"
        >
          ยืนยันการลบ
        </button>
      </div>
    </div>
  `);

  document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-btn-confirm-del')?.addEventListener('click', async () => {
    closeCustomModal();
    const res = await apiFetch(`/api/master/${type}/${id}`, { method: 'DELETE' });
    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถลบรายการได้', 'error');
      return;
    }
    showToast(`ลบ "${label}" สำเร็จ`, 'success');
    await loadAllData();
  });
}

// ============================================================================
// EVENT LISTENERS INITIALIZATION
// ============================================================================

document.addEventListener('DOMContentLoaded', async () => {
  const formDate = document.getElementById('form-date');
  if (formDate) {
    formDate.value = getTodayISO();
  }
  updateDateAndDurationPreview();

  // 1. Login Password Visibility Toggle
  document.getElementById('toggle-password-btn')?.addEventListener('click', () => {
    const pwdInput = document.getElementById('login-password');
    const icon = document.getElementById('toggle-password-icon');
    if (pwdInput.type === 'password') {
      pwdInput.type = 'text';
      icon.className = 'fa-solid fa-eye-slash text-indigo-600';
    } else {
      pwdInput.type = 'password';
      icon.className = 'fa-solid fa-eye';
    }
  });

  // 2. Login Form Submit (No Quick-Fill Buttons)
  document.getElementById('login-form')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = document.getElementById('login-username').value.trim();
    const password = document.getElementById('login-password').value;
    const errBox = document.getElementById('login-error-box');
    const errText = document.getElementById('login-error-text');

    errBox.classList.add('hidden');
    errBox.classList.remove('flex');

    const res = await apiFetch('/api/login', {
      method: 'POST',
      body: JSON.stringify({ username, password })
    });

    if (!res.ok) {
      errText.textContent = res.data?.error || 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง';
      errBox.classList.remove('hidden');
      errBox.classList.add('flex');
      return;
    }

    saveSession(res.data.token, res.data.user);
    if (res.data.storageQuota) {
      state.storageQuota = res.data.storageQuota;
    }
    state.currentView = 'front';
    renderAuthView();
    await loadAllData();
    showToast(`ยินดีต้อนรับ ${res.data.user.full_name}`, 'success');
  });

  // 4. Logout, Desktop Tabs, Mobile Bottom Nav & Database Reset
  document.getElementById('btn-logout')?.addEventListener('click', async () => {
    try {
      await apiFetch('/api/logout', { method: 'POST' });
    } catch {
      // ignore
    }
    exitEditModeAndResetForm();
    clearSession();
    renderAuthView();
    showToast('ออกจากระบบเรียบร้อยแล้ว', 'info');
  });

  document.getElementById('nav-tab-front')?.addEventListener('click', () => switchView('front'));
  document.getElementById('nav-tab-back')?.addEventListener('click', () => switchView('back'));
  document.getElementById('btn-back-to-front')?.addEventListener('click', () => switchView('front'));

  // Mobile Bottom Navigation Bar
  document.getElementById('mob-nav-form')?.addEventListener('click', () => {
    switchView('front');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });
  document.getElementById('mob-nav-table')?.addEventListener('click', () => {
    switchView('front');
    document.getElementById('logs-section-anchor')?.scrollIntoView({ behavior: 'smooth' });
  });
  document.getElementById('mob-nav-backoffice')?.addEventListener('click', () => {
    switchView('back');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  document.getElementById('btn-reset-db')?.addEventListener('click', () => {
    openCustomModal(`
      <div class="p-6">
        <div class="flex items-start gap-4">
          <div class="w-12 h-12 rounded-2xl bg-amber-100 text-amber-600 flex items-center justify-center text-xl shrink-0">
            <i class="fa-solid fa-rotate-left"></i>
          </div>
          <div>
            <h3 class="text-lg font-bold text-slate-900">ยืนยันการรีเซ็ตข้อมูลฐานข้อมูล</h3>
            <p class="text-sm text-slate-600 mt-1">
              ระบบจะรีเซ็ตข้อมูลทั้ง 6 ตารางกลับเป็นข้อมูลเริ่มต้น (Seed Data พร้อมทะเบียนรถ 20 คัน) คุณต้องการดำเนินการต่อหรือไม่?
            </p>
          </div>
        </div>
        <div class="mt-6 pt-4 border-t border-slate-100 flex justify-end gap-2.5">
          <button
            type="button"
            id="modal-btn-cancel-reset"
            class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white text-slate-700 text-sm font-medium cursor-pointer"
          >
            ยกเลิก
          </button>
          <button
            type="button"
            id="modal-btn-confirm-reset"
            class="min-h-[42px] px-5 py-2 rounded-xl bg-amber-600 hover:bg-amber-700 text-white text-sm font-semibold cursor-pointer"
          >
            ยืนยันการรีเซ็ตฐานข้อมูล
          </button>
        </div>
      </div>
    `);

    document.getElementById('modal-btn-cancel-reset')?.addEventListener('click', closeCustomModal);
    document.getElementById('modal-btn-confirm-reset')?.addEventListener('click', async () => {
      closeCustomModal();
      const res = await apiFetch('/api/reset', { method: 'POST' });
      if (!res.ok) {
        showToast(res.data?.error || 'ไม่สามารถรีเซ็ตฐานข้อมูลได้', 'error');
        return;
      }
      exitEditModeAndResetForm();
      await loadAllData();
      showToast(res.data?.message || 'รีเซ็ตฐานข้อมูลสำเร็จ', 'success');
    });
  });

  // 5. Admin Quick Quota Simulator Buttons (Warning 80% / Full 100% / Normal 100)
  document.querySelectorAll('button[data-sim-quota]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const mode = btn.getAttribute('data-sim-quota');
      const count = Math.max(1, state.logs.length || 4);
      if (mode === 'warning') {
        // Set limit so current count is >= 80% (e.g. 4 / 5 = 80%)
        const targetLimit = Math.max(count + 1, Math.ceil(count / 0.8));
        await quickSetQuotaLimit(targetLimit);
      } else if (mode === 'full') {
        // Set limit equal to current count (100% Full!)
        await quickSetQuotaLimit(count);
      } else {
        await quickSetQuotaLimit(100);
      }
    });
  });

  // 6. Front-Office Worker Selection -> Auto-fill Default Route
  document.getElementById('form-worker')?.addEventListener('change', () => {
    const worker = getSelectedWorkerObject();
    const routeInput = document.getElementById('form-route');
    if (worker) {
      if (worker.default_route) {
        routeInput.value = worker.default_route;
        routeInput.classList.remove('input-highlight-pulse');
        void routeInput.offsetWidth;
        routeInput.classList.add('input-highlight-pulse');
      } else if (!state.editingWorkerName) {
        routeInput.value = '';
      }
    }
    renderRouteComboboxMenu();
    updateRouteSmartControls();
  });

  // 7. Cartrack GPS Check Button next to Vehicle Select & Sync All Button
  document
    .getElementById('btn-check-vehicle-gps')
    ?.addEventListener('click', checkVehicleGpsFromForm);
  document.getElementById('btn-sync-all-gps')?.addEventListener('click', syncAllVehiclesGps);

  // 8. Combo Box Route Input & Smart Route Buttons
  const routeInput = document.getElementById('form-route');
  const routeMenu = document.getElementById('route-combobox-menu');

  routeInput?.addEventListener('input', () => {
    renderRouteComboboxMenu(routeInput.value);
    routeMenu?.classList.remove('hidden');
    updateRouteSmartControls();
  });

  routeInput?.addEventListener('focus', () => {
    renderRouteComboboxMenu('');
    routeMenu?.classList.remove('hidden');
  });

  document.getElementById('btn-toggle-route-dropdown')?.addEventListener('click', () => {
    if (routeMenu.classList.contains('hidden')) {
      renderRouteComboboxMenu('');
      routeMenu.classList.remove('hidden');
      routeInput.focus();
    } else {
      routeMenu.classList.add('hidden');
    }
  });

  document.addEventListener('click', (e) => {
    const wrapper = document.getElementById('route-combobox-wrapper');
    if (wrapper && !wrapper.contains(e.target)) {
      routeMenu?.classList.add('hidden');
    }
  });

  document.getElementById('btn-reset-default-route')?.addEventListener('click', () => {
    const worker = getSelectedWorkerObject();
    if (worker && worker.default_route) {
      routeInput.value = worker.default_route;
      updateRouteSmartControls();
      showToast(`รีเซ็ตเส้นทางกลับเป็นไซต์งานประจำของ "${worker.name}" แล้ว`, 'info');
    }
  });

  document.getElementById('btn-add-to-routes')?.addEventListener('click', async () => {
    const val = routeInput.value.trim();
    if (!val) return;

    const res = await apiFetch('/api/master/routes', {
      method: 'POST',
      body: JSON.stringify({ name: val, active: 1, reactivateIfInactive: true })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถเพิ่มในรายการที่หมายได้', 'error');
      return;
    }

    showToast(`เพิ่ม "${val}" เข้าสู่รายการเส้นทาง/ที่หมายหลักแล้ว`, 'success');
    await loadAllData();
  });

  document.getElementById('btn-set-worker-default-route')?.addEventListener('click', async () => {
    const worker = getSelectedWorkerObject();
    const val = routeInput.value.trim();
    if (!worker || !worker.id || !val) return;

    const res = await apiFetch(`/api/master/workers/${worker.id}`, {
      method: 'PUT',
      body: JSON.stringify({ default_route: val })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถตั้งไซต์งานประจำได้', 'error');
      return;
    }

    showToast(`ตั้ง "${val}" เป็นไซต์งานประจำของ "${worker.name}" เรียบร้อยแล้ว`, 'success');
    await loadAllData();
  });

  // 9. Date & Time Inputs Live Preview
  document.getElementById('form-date')?.addEventListener('input', updateDateAndDurationPreview);
  document.getElementById('form-departure')?.addEventListener('input', updateDateAndDurationPreview);
  document.getElementById('form-arrival')?.addEventListener('input', updateDateAndDurationPreview);

  // 10. Form Submit, Clear, Cancel Edit
  document.getElementById('travel-log-form')?.addEventListener('submit', handleTravelLogSubmit);
  document.getElementById('btn-clear-form')?.addEventListener('click', exitEditModeAndResetForm);
  document.getElementById('btn-cancel-edit')?.addEventListener('click', exitEditModeAndResetForm);
  document.getElementById('btn-cancel-edit-top')?.addEventListener('click', exitEditModeAndResetForm);

  // 11. Table Filters & Excel/CSV Export
  const filterDateInput = document.getElementById('filter-date');
  const filterSearchInput = document.getElementById('filter-search');
  const btnClearSearch = document.getElementById('btn-clear-search');

  filterDateInput?.addEventListener('input', () => {
    state.filterDate = filterDateInput.value;
    renderLogsTable();
  });

  document.getElementById('btn-filter-today')?.addEventListener('click', () => {
    const today = getTodayISO();
    filterDateInput.value = today;
    state.filterDate = today;
    renderLogsTable();
  });

  document.getElementById('btn-filter-all-dates')?.addEventListener('click', () => {
    filterDateInput.value = '';
    state.filterDate = '';
    renderLogsTable();
  });

  filterSearchInput?.addEventListener('input', () => {
    state.filterSearch = filterSearchInput.value;
    if (state.filterSearch) {
      btnClearSearch.classList.remove('hidden');
      btnClearSearch.classList.add('flex');
    } else {
      btnClearSearch.classList.add('hidden');
      btnClearSearch.classList.remove('flex');
    }
    renderLogsTable();
  });

  btnClearSearch?.addEventListener('click', () => {
    filterSearchInput.value = '';
    state.filterSearch = '';
    btnClearSearch.classList.add('hidden');
    btnClearSearch.classList.remove('flex');
    renderLogsTable();
  });

  document.getElementById('btn-export-excel')?.addEventListener('click', exportLogsToExcel);
  document.getElementById('btn-export-csv')?.addEventListener('click', exportLogsToCSV);

  // 12. Back-Office: Storage Quota Settings Form
  document.getElementById('form-gps-settings')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const payload = {
      ...state.gpsSettings,
      max_storage_limit: Number(document.getElementById('setting-max-limit').value) || 100
    };

    const res = await apiFetch('/api/settings/gps', {
      method: 'PUT',
      body: JSON.stringify(payload)
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถบันทึกการตั้งค่าได้', 'error');
      return;
    }

    state.gpsSettings = res.data.settings;
    state.storageQuota = res.data.storageQuota;
    populateGpsSettingsForm();
    renderStorageQuotaUI(false);
    showToast('บันทึกการตั้งค่าขีดจำกัดพื้นที่เก็บข้อมูลเรียบร้อยแล้ว', 'success');
  });

  // 13. Back-Office Add Forms (Workers, Routes, Vehicles)
  document.getElementById('form-add-worker')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const nameInput = document.getElementById('new-worker-name');
    const routeInputEl = document.getElementById('new-worker-default-route');
    const name = nameInput.value.trim();
    const default_route = routeInputEl.value.trim();
    if (!name) return;

    const res = await apiFetch('/api/master/workers', {
      method: 'POST',
      body: JSON.stringify({ name, default_route, active: 1 })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถเพิ่มรายชื่อคนงานได้', 'error');
      return;
    }

    nameInput.value = '';
    routeInputEl.value = '';
    showToast(`เพิ่มคนงาน "${name}" สำเร็จ`, 'success');
    await loadAllData();
  });

  document.getElementById('form-add-route')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const nameInput = document.getElementById('new-route-name');
    const name = nameInput.value.trim();
    if (!name) return;

    const res = await apiFetch('/api/master/routes', {
      method: 'POST',
      body: JSON.stringify({ name, active: 1 })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถเพิ่มเส้นทางได้', 'error');
      return;
    }

    nameInput.value = '';
    showToast(`เพิ่มเส้นทาง "${name}" สำเร็จ`, 'success');
    await loadAllData();
  });

  document.getElementById('form-add-vehicle')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const plateInput = document.getElementById('new-vehicle-plate');
    const plate = plateInput.value.trim();
    if (!plate) return;

    const res = await apiFetch('/api/master/vehicles', {
      method: 'POST',
      body: JSON.stringify({ plate, active: 1 })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถเพิ่มทะเบียนรถได้', 'error');
      return;
    }

    plateInput.value = '';
    showToast(`เพิ่มทะเบียนรถ "${plate}" สำเร็จ`, 'success');
    await loadAllData();
  });

  // 14. Verify existing session on startup
  if (state.token) {
    try {
      const meRes = await apiFetch('/api/me');
      if (meRes.ok && meRes.data?.user) {
        state.user = meRes.data.user;
        if (meRes.data.storageQuota) {
          state.storageQuota = meRes.data.storageQuota;
        }
        localStorage.setItem('dispatch_auth_user', JSON.stringify(state.user));
        renderAuthView();
        await loadAllData();
        return;
      }
    } catch {
      // session expired
    }
    clearSession();
  }

  renderAuthView();
});
