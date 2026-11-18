/**
 * ระบบจัดการบันทึกข้อมูลการเดินทางของคนงาน (Worker Dispatch & Travel Log System v2.0)
 * Cloud Supabase + Cartrack GPS API + Storage Quota Alert + Mobile-First Responsive UI
 */

const state = {
  token: localStorage.getItem('dispatch_auth_token') || null,
  user: JSON.parse(localStorage.getItem('dispatch_auth_user') || 'null'),
  currentView: 'front', // 'front' | 'back'
  editingWorkerName: null, // Worker name of the trip currently being edited
  editingLogId: null, // Unique log_id of the trip currently being edited
  lastCheckedGps: null, // Stores latest checked GPS object from form button
  linkedCopyHandle: null, // Persistent FileSystemFileHandle for แบบสำเนา
  linkedBillingHandle: null, // Persistent FileSystemFileHandle for แบบใบเบิก

  // Data from Backend
  activeWorkers: [],
  activeSites: [],
  activeRoutes: [],
  activeVehicles: [],
  allWorkers: [],
  allSites: [],
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
  filterSite: '',
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
              ความจุข้อมูลเกิน 80% แล้ว แนะนำให้กดส่งออกแบบฟอร์มไฟล์ Excel เพื่อสำรองข้อมูลไว้ล่วงหน้าก่อนที่พื้นที่เก็บข้อมูลจะเต็ม
            </p>
          </div>
        </div>
        <div class="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            id="banner-warn-export-excel"
            class="min-h-[40px] px-4 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 cursor-pointer"
          >
            <i class="fa-solid fa-file-excel"></i>
            <span>ส่งออกแบบฟอร์ม Excel (ครบทั้ง 2 แบบฟอร์ม)</span>
          </button>
        </div>
      `;
      warningBanner
        .querySelector('#banner-warn-export-excel')
        ?.addEventListener('click', exportLogsToExcel);
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
            id="modal-full-export-excel"
            class="sm:col-span-2 min-h-[44px] px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center justify-center gap-2 shadow-sm transition cursor-pointer"
          >
            <i class="fa-solid fa-file-excel text-base"></i>
            <span>ส่งออกแบบฟอร์ม Excel (ครบทั้ง 2 แบบฟอร์ม) ทันที</span>
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
      apiFetch('/api/master/sites?active=1'),
      apiFetch('/api/master/routes?active=1'),
      apiFetch('/api/master/vehicles?active=1'),
      apiFetch('/api/logs'),
      apiFetch('/api/master/workers'),
      apiFetch('/api/master/sites'),
      apiFetch('/api/master/routes'),
      apiFetch('/api/master/vehicles'),
      apiFetch('/api/storage-status')
    ];

    if (state.user?.role === 'admin') {
      requests.push(apiFetch('/api/settings/gps'));
    }

    const [
      workersRes,
      sitesRes,
      routesRes,
      vehiclesRes,
      logsRes,
      allWorkersRes,
      allSitesRes,
      allRoutesRes,
      allVehiclesRes,
      quotaRes,
      gpsSettingsRes
    ] = await Promise.all(requests);

    if (workersRes.ok) state.activeWorkers = workersRes.data || [];
    if (sitesRes.ok) state.activeSites = sitesRes.data || [];
    if (routesRes.ok) state.activeRoutes = routesRes.data || [];
    if (vehiclesRes.ok) state.activeVehicles = vehiclesRes.data || [];
    if (logsRes.ok) state.logs = logsRes.data || [];
    if (allWorkersRes.ok) state.allWorkers = allWorkersRes.data || [];
    if (allSitesRes.ok) state.allSites = allSitesRes.data || [];
    if (allRoutesRes.ok) state.allRoutes = allRoutesRes.data || [];
    if (allVehiclesRes.ok) state.allVehicles = allVehiclesRes.data || [];
    if (quotaRes.ok && quotaRes.data) state.storageQuota = quotaRes.data;
    if (gpsSettingsRes?.ok && gpsSettingsRes.data?.settings) {
      state.gpsSettings = gpsSettingsRes.data.settings;
      populateGpsSettingsForm();
    }

    renderWorkerDropdown();
    renderSiteDropdown();
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

  const currentVal = select.value || state.editingWorkerName || '';
  const tripCountMap = new Map();
  for (const l of state.logs) {
    if (l.worker_name) {
      tripCountMap.set(l.worker_name, (tripCountMap.get(l.worker_name) || 0) + 1);
    }
  }

  let html = '<option value="">-- เลือกชื่อคนขับ / พนักงาน (บังคับเลือก) --</option>';

  const listToRender = [...state.activeWorkers];
  if (
    currentVal &&
    !listToRender.some((w) => w.name === currentVal)
  ) {
    const foundInAll = state.allWorkers.find((w) => w.name === currentVal);
    listToRender.unshift(
      foundInAll || { id: 0, name: currentVal, default_route: '', active: false }
    );
  }

  for (const w of listToRender) {
    const routeSuffix = w.default_route
      ? ` (ไซต์ประจำ: ${w.default_route})`
      : '';
    const count = tripCountMap.get(w.name) || 0;
    const countSuffix = count > 0 ? ` [บันทึกสะสม ${count} เที่ยว]` : '';
    const isSelected = w.name === currentVal ? 'selected' : '';
    html += `<option value="${escapeHtml(w.name)}" ${isSelected}>${escapeHtml(
      w.name + routeSuffix + countSuffix
    )}</option>`;
  }

  select.innerHTML = html;
  select.disabled = false;
}

function renderSiteDropdown() {
  const select = document.getElementById('form-site');
  if (select) {
    const currentVal = select.value;
    let html = '<option value="">-- เลือกไซต์งาน / โครงการ --</option>';

    const listToRender = [...state.activeSites];
    if (currentVal && !listToRender.some((s) => s.name === currentVal)) {
      listToRender.unshift({ id: 0, name: currentVal, active: false });
    }

    for (const s of listToRender) {
      const isSelected = s.name === currentVal ? 'selected' : '';
      html += `<option value="${escapeHtml(s.name)}" ${isSelected}>${escapeHtml(s.name)}</option>`;
    }

    select.innerHTML = html;
  }

  const filterSelect = document.getElementById('filter-site');
  if (filterSelect) {
    const currentFilter = state.filterSite || filterSelect.value || '';
    const allSiteNames = new Set([
      ...state.activeSites.map((s) => s.name),
      ...state.logs.map((l) => (l.site_name || '').trim()).filter(Boolean)
    ]);
    let filterHtml = '<option value="">ทุกไซต์งาน / โครงการ</option>';
    for (const sName of allSiteNames) {
      const isSel = sName === currentFilter ? 'selected' : '';
      filterHtml += `<option value="${escapeHtml(sName)}" ${isSel}>${escapeHtml(sName)}</option>`;
    }
    filterSelect.innerHTML = filterHtml;
  }
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

function enterEditMode(logId) {
  const log = state.logs.find((l) => (l.log_id || l.worker_name) === logId);
  if (!log) {
    showToast(`ไม่พบข้อมูลรายการที่ต้องการแก้ไข`, 'error');
    return;
  }

  state.editingLogId = log.log_id || log.worker_name;
  state.editingWorkerName = log.worker_name;
  state.lastCheckedGps = parseGpsStatus(log.gps_status);
  renderWorkerDropdown();

  const workerSelect = document.getElementById('form-worker');
  workerSelect.value = log.worker_name;
  workerSelect.disabled = false;

  document.getElementById('form-date').value = log.travel_date || '';
  document.getElementById('form-route').value = log.route || '';
  document.getElementById('form-departure').value = log.departure_time || '';
  document.getElementById('form-arrival').value = log.arrival_time || '';

  renderSiteDropdown();
  const siteEl = document.getElementById('form-site');
  if (siteEl) {
    siteEl.value = log.site_name || '';
    renderSiteDropdown();
    siteEl.value = log.site_name || '';
  }

  renderVehicleDropdown();
  document.getElementById('form-vehicle').value = log.vehicle_plate || '';
  const passengersEl = document.getElementById('form-passengers');
  if (passengersEl) {
    passengersEl.value =
      log.passenger_count !== null && log.passenger_count !== undefined ? log.passenger_count : '';
  }
  const noteEl = document.getElementById('form-note');
  if (noteEl) {
    noteEl.value = log.note || '';
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
  modeTitle.textContent = `แก้ไขรายการเดินทางของ: ${log.worker_name}`;

  editWorkerLabel.textContent = `${log.worker_name} (${formatShortThaiDateBE(log.travel_date) || 'ไม่ระบุวัน'})`;
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
  state.editingLogId = null;
  state.editingWorkerName = null;
  state.lastCheckedGps = null;
  renderWorkerDropdown();

  const workerSelect = document.getElementById('form-worker');
  workerSelect.value = '';
  workerSelect.disabled = false;

  document.getElementById('form-date').value = getTodayISO();
  const siteEl = document.getElementById('form-site');
  if (siteEl) siteEl.value = '';
  renderSiteDropdown();
  document.getElementById('form-route').value = '';
  document.getElementById('form-departure').value = '';
  document.getElementById('form-arrival').value = '';
  document.getElementById('form-vehicle').value = '';
  const passengersEl = document.getElementById('form-passengers');
  if (passengersEl) passengersEl.value = '';
  const noteEl = document.getElementById('form-note');
  if (noteEl) noteEl.value = '';
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
  modeTitle.textContent = 'บันทึกข้อมูลการเดินทางพนักงาน (สะสมรายเที่ยว)';

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
    const isEdit = Boolean(state.editingLogId);
    const url = isEdit
      ? `/api/logs/${encodeURIComponent(state.editingLogId)}`
      : '/api/logs';
    const method = isEdit ? 'PUT' : 'POST';

    const res = await apiFetch(url, {
      method,
      body: JSON.stringify(payload)
    });

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

  const worker_name = (document.getElementById('form-worker')?.value || '').trim();
  const site_name = (document.getElementById('form-site')?.value || '').trim();
  const travel_date = (document.getElementById('form-date')?.value || '').trim();
  const route = (document.getElementById('form-route')?.value || '').trim();
  const departure_time = (document.getElementById('form-departure')?.value || '').trim();
  const arrival_time = (document.getElementById('form-arrival')?.value || '').trim();
  const vehicle_plate = (document.getElementById('form-vehicle')?.value || '').trim();
  const passenger_count_raw = (document.getElementById('form-passengers')?.value || '').trim();
  const passenger_count =
    passenger_count_raw === '' ? null : Math.max(0, Number.parseInt(passenger_count_raw, 10));
  const note = (document.getElementById('form-note')?.value || '').trim();

  if (!worker_name) {
    openCustomModal(`
      <div class="p-6">
        <div class="flex items-start gap-4">
          <div class="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center text-xl shrink-0">
            <i class="fa-solid fa-user-xmark"></i>
          </div>
          <div>
            <h3 class="text-lg font-bold text-slate-900">กรุณาเลือกชื่อคนขับ / พนักงาน</h3>
            <p class="text-sm text-slate-600 mt-1">
              ช่อง <strong>"ชื่อคนขับ / พนักงาน"</strong> เป็นข้อมูลบังคับเลือกก่อนทำการบันทึก
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
    site_name,
    travel_date,
    route,
    departure_time,
    arrival_time,
    vehicle_plate,
    passenger_count: Number.isNaN(passenger_count) ? null : passenger_count,
    note
  };

  if (!state.editingLogId && state.storageQuota?.isFull) {
    await executeSaveLog(payload);
    return;
  }

  // Check incomplete optional fields
  const missingFields = [];
  if (!site_name) missingFields.push('ไซต์งาน / โครงการ (หัวตาราง)');
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
  const siteFilter = (state.filterSite || '').trim();
  const dateFilter = state.filterDate.trim();
  const searchFilter = state.filterSearch.trim().toLowerCase();

  return state.logs.filter((row) => {
    if (siteFilter && (row.site_name || '').trim() !== siteFilter) {
      return false;
    }
    if (dateFilter && (row.travel_date || '') !== dateFilter) {
      return false;
    }
    if (searchFilter) {
      const thaiDateShort = formatThaiDate(row.travel_date, false).toLowerCase();
      const thaiDateFull = formatThaiDate(row.travel_date, true).toLowerCase();
      const haystack = [
        row.worker_name || '',
        row.site_name || '',
        row.travel_date || '',
        thaiDateShort,
        thaiDateFull,
        row.route || '',
        row.vehicle_plate || '',
        row.passenger_count !== null && row.passenger_count !== undefined
          ? `${row.passenger_count} คน`
          : '',
        row.note || '',
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
    tbody.innerHTML = `<tr><td colspan="9">${emptyHtml}</td></tr>`;
    if (mobileCards) mobileCards.innerHTML = emptyHtml;
    return;
  }

  const workerDefaultMap = new Map(
    state.allWorkers.map((w) => [w.name, (w.default_route || '').trim()])
  );

  // 1. Render Desktop Table Rows
  tbody.innerHTML = filtered
    .map((row) => {
      const rowKey = row.log_id || row.worker_name;
      const isEditingThis = state.editingLogId === rowKey;
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
              row.site_name
                ? `<span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-amber-50 text-amber-900 border border-amber-200 text-xs font-semibold">
                    <i class="fa-solid fa-building-user text-amber-600 text-[11px]"></i>
                    <span>${escapeHtml(row.site_name)}</span>
                  </span>`
                : '<span class="text-xs text-slate-400 italic">ไม่ได้ระบุ</span>'
            }
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
                  ${
                    row.note
                      ? `<p class="text-[11px] text-amber-700 mt-0.5"><i class="fa-regular fa-note-sticky mr-1"></i>${escapeHtml(
                          row.note
                        )}</p>`
                      : ''
                  }
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
                data-log-id="${escapeHtml(rowKey)}"
                class="px-2.5 py-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-semibold inline-flex items-center gap-1 transition cursor-pointer"
              >
                <i class="fa-solid fa-pen-to-square"></i>
                <span>แก้ไข</span>
              </button>
              <button
                type="button"
                data-action="delete-log"
                data-log-id="${escapeHtml(rowKey)}"
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
        const rowKey = row.log_id || row.worker_name;
        const defaultRoute = workerDefaultMap.get(row.worker_name) || '';
        const isDefaultRoute = defaultRoute && row.route && row.route.trim() === defaultRoute;
        const duration = calculateTravelDuration(row.departure_time, row.arrival_time);

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
                <span class="text-slate-500 shrink-0">ไซต์งาน/โครงการ:</span>
                <span class="font-semibold text-amber-800 text-right">
                  ${escapeHtml(row.site_name || '-')}
                </span>
              </div>
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
              ${
                row.note
                  ? `<div class="flex items-center justify-between gap-2">
                      <span class="text-slate-500">หมายเหตุ:</span>
                      <span class="font-medium text-amber-700">${escapeHtml(row.note)}</span>
                    </div>`
                  : ''
              }
            </div>

            <div class="flex items-center gap-2 pt-1">
              <button
                type="button"
                data-action="edit-log"
                data-log-id="${escapeHtml(rowKey)}"
                class="flex-1 min-h-[44px] rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-200 font-semibold text-xs flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <i class="fa-solid fa-pen-to-square"></i>
                <span>แก้ไขข้อมูล</span>
              </button>
              <button
                type="button"
                data-action="delete-log"
                data-log-id="${escapeHtml(rowKey)}"
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
      enterEditMode(btn.getAttribute('data-log-id'));
    });
  });

  document.querySelectorAll('button[data-action="delete-log"]').forEach((btn) => {
    btn.addEventListener('click', () => {
      confirmDeleteTravelLog(
        btn.getAttribute('data-log-id'),
        btn.getAttribute('data-worker')
      );
    });
  });

  document.querySelectorAll('button[data-refresh-gps]').forEach((btn) => {
    btn.addEventListener('click', () => {
      refreshSingleVehicleGps(btn.getAttribute('data-refresh-gps'));
    });
  });
}

function confirmDeleteTravelLog(logId, workerLabel = '') {
  const displayLabel = workerLabel || logId;
  openCustomModal(`
    <div class="p-6">
      <div class="flex items-start gap-4">
        <div class="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center text-xl shrink-0">
          <i class="fa-solid fa-trash-can"></i>
        </div>
        <div class="flex-1">
          <h3 class="text-lg font-bold text-slate-900">ยืนยันการลบรายการเดินทาง</h3>
          <p class="text-sm text-slate-600 mt-1">
            คุณต้องการลบรายการเดินทางของ <strong>"${escapeHtml(displayLabel)}"</strong> ออกจากฐานข้อมูลใช่หรือไม่?
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
      const res = await apiFetch(`/api/logs/${encodeURIComponent(logId)}`, {
        method: 'DELETE'
      });
      if (!res.ok) {
        showToast(res.data?.error || 'ไม่สามารถลบข้อมูลได้', 'error');
        return;
      }
      if (state.editingLogId === logId) {
        exitEditModeAndResetForm();
      }
      showToast(`ลบรายการของ "${displayLabel}" เรียบร้อยแล้ว`, 'success');
      await loadAllData();
    } catch (err) {
      showToast(err.message || 'เกิดข้อผิดพลาดในการลบข้อมูล', 'error');
    }
  });
}

// Format YYYY-MM-DD into short Thai Buddhist date (e.g. 18/11/2569)
function formatShortThaiDateBE(isoDate) {
  if (!isoDate) return '';
  const m = String(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return String(isoDate);
  const yearBE = Number(m[1]) + 543;
  const month = Number(m[2]);
  const day = Number(m[3]);
  return `${day}/${month}/${yearBE}`;
}

// Split route string into { from, to } (e.g. "นาคนิวาส48 - ไลฟ์รัชดา-พระราม9")
function splitRouteParts(routeStr, fallbackSite = '') {
  const raw = String(routeStr || '').trim();
  const site = String(fallbackSite || '').trim();
  if (!raw) return { from: site, to: site };
  // Prefer splitting on spaced separator (" - ", " -> ", " ถึง ") first so names like "ไลฟ์รัชดา-พระราม9" stay intact
  const spacedParts = raw.split(/\s+(?:-|->|→|ถึง)\s*|\s*(?:-|->|→|ถึง)\s+/).filter(Boolean);
  if (spacedParts.length >= 2) {
    return {
      from: spacedParts[0].trim(),
      to: spacedParts.slice(1).join(' - ').trim()
    };
  }
  const parts = raw.split(/\s*(?:-|->|→|ถึง)\s*/).filter(Boolean);
  if (parts.length >= 2) {
    return {
      from: parts[0].trim(),
      to: parts.slice(1).join('-').trim()
    };
  }
  return {
    from: raw,
    to: site || raw
  };
}

// Get directional { from, to } based on departure time (เช้า < 12:00 = แคมป์ -> โครงการ, บ่าย/เย็น >= 12:00 = โครงการ -> แคมป์)
function getDirectionalRouteParts(routeStr, departureTimeStr, fallbackSite = '') {
  const { from, to } = splitRouteParts(routeStr, fallbackSite);
  if (!from || !to || from === to) return { from, to };
  const m = String(departureTimeStr || '').trim().match(/^(\d{1,2}):(\d{2})/);
  if (m) {
    const hour = Number(m[1]);
    if (hour >= 12) {
      return { from: to, to: from };
    }
  }
  return { from, to };
}

function formatExcelTimeDisplay(timeStr) {
  const s = String(timeStr || '').trim();
  if (!s) return '';
  const m = s.match(/^0?(\d{1,2}):(\d{2})/);
  if (m) return `${m[1]}:${m[2]}`;
  return s;
}

// ============================================================================
// INDEXED-DB FILE HANDLE PERSISTENCE (ไหลข้อมูลเข้าไฟล์ Excel ไฟล์เดิมอัตโนมัติ)
// ============================================================================

const EXCEL_SYNC_DB_NAME = 'ckp_excel_sync_db';
const EXCEL_SYNC_STORE = 'file_handles';
const EXCEL_SYNC_KEY_COPY = 'copy_excel_file';
const EXCEL_SYNC_KEY_BILLING = 'billing_excel_file';

function openExcelSyncIdb() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      resolve(null);
      return;
    }
    const req = window.indexedDB.open(EXCEL_SYNC_DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(EXCEL_SYNC_STORE)) {
        db.createObjectStore(EXCEL_SYNC_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveLinkedExcelHandleByType(type, handle) {
  if (type === 'copy') {
    state.linkedCopyHandle = handle;
  } else {
    state.linkedBillingHandle = handle;
  }
  updateLinkedExcelUI();
  try {
    const db = await openExcelSyncIdb();
    if (!db) return;
    const key = type === 'copy' ? EXCEL_SYNC_KEY_COPY : EXCEL_SYNC_KEY_BILLING;
    await new Promise((resolve, reject) => {
      const tx = db.transaction(EXCEL_SYNC_STORE, 'readwrite');
      tx.objectStore(EXCEL_SYNC_STORE).put(handle, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // ignore IndexedDB storage errors
  }
}

async function loadLinkedExcelHandles() {
  try {
    const db = await openExcelSyncIdb();
    if (!db) return;
    const [copyHandle, billingHandle] = await Promise.all([
      new Promise((resolve) => {
        const tx = db.transaction(EXCEL_SYNC_STORE, 'readonly');
        const req = tx.objectStore(EXCEL_SYNC_STORE).get(EXCEL_SYNC_KEY_COPY);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      }),
      new Promise((resolve) => {
        const tx = db.transaction(EXCEL_SYNC_STORE, 'readonly');
        const req = tx.objectStore(EXCEL_SYNC_STORE).get(EXCEL_SYNC_KEY_BILLING);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      })
    ]);
    if (copyHandle) state.linkedCopyHandle = copyHandle;
    if (billingHandle) state.linkedBillingHandle = billingHandle;
    updateLinkedExcelUI();
  } catch {
    // ignore
  }
}

function updateLinkedExcelUI() {
  const copyNameEl = document.getElementById('linked-copy-filename');
  const billingNameEl = document.getElementById('linked-billing-filename');
  const btnCopy = document.getElementById('btn-link-copy-excel');
  const btnBilling = document.getElementById('btn-link-billing-excel');

  const hasCopy = Boolean(state.linkedCopyHandle && state.linkedCopyHandle.name);
  const hasBilling = Boolean(state.linkedBillingHandle && state.linkedBillingHandle.name);

  if (copyNameEl) {
    copyNameEl.textContent = hasCopy ? state.linkedCopyHandle.name : 'เลือกไฟล์...';
  }
  if (btnCopy) {
    btnCopy.className = hasCopy
      ? 'min-h-[40px] px-3 py-1.5 rounded-xl border border-emerald-400 bg-emerald-50 hover:bg-emerald-100 text-emerald-950 text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer truncate max-w-[260px]'
      : 'min-h-[40px] px-3 py-1.5 rounded-xl bg-white hover:bg-amber-100 text-amber-950 border border-amber-300 text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer truncate max-w-[260px]';
  }

  if (billingNameEl) {
    billingNameEl.textContent = hasBilling ? state.linkedBillingHandle.name : 'เลือกไฟล์...';
  }
  if (btnBilling) {
    btnBilling.className = hasBilling
      ? 'min-h-[40px] px-3 py-1.5 rounded-xl border border-emerald-400 bg-emerald-50 hover:bg-emerald-100 text-emerald-950 text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer truncate max-w-[260px]'
      : 'min-h-[40px] px-3 py-1.5 rounded-xl bg-white hover:bg-sky-100 text-sky-950 border border-sky-300 text-xs font-semibold flex items-center gap-1.5 transition cursor-pointer truncate max-w-[260px]';
  }
}

// เลือกผูกไฟล์แยกกัน (ผูกไฟล์เฉยๆ ยังไม่เขียนทับจนกว่าจะกดปุ่ม "บันทึกไฟล์สำเนา" หรือ "บันทึกไฟล์ใบเบิก")
async function pickAndLinkExcelFileOnly(type) {
  const isCopy = type === 'copy';
  const typeLabel = isCopy ? 'สำเนา' : 'ใบเบิก';

  if (typeof window.showOpenFilePicker !== 'function') {
    openExcelLinkModalByType(type);
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: `ไฟล์ Excel แบบ${typeLabel} (.xlsx)`,
          accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
          }
        }
      ]
    });
    if (handle) {
      await saveLinkedExcelHandleByType(type, handle);
      showToast(
        `ผูกไฟล์${typeLabel} "${handle.name}" สำเร็จ!`,
        'success',
        `เมื่อต้องการให้ข้อมูลไหลเข้าไฟล์นี้ ให้กดปุ่ม "บันทึกไฟล์${typeLabel}" ได้เลยครับ`
      );
    }
  } catch (err) {
    if (err?.name !== 'AbortError') {
      showToast(`ยกเลิกการเลือกผูกไฟล์${typeLabel}`, 'info');
    }
  }
}

// ============================================================================
// EXACT TEMPLATE BUILDERS (ตามแบบฟอร์มต้นฉบับที่ผู้ใช้ระบุ 100%)
// ============================================================================

// Section 1: ใบเบิกผลงานผู้รับเหมา (มีเว้นบรรทัดคั่นระหว่างรายการ 1, 2, 3... และคำนวณยอดรวมอัตโนมัติ)
function appendBillingSectionToAoa(
  aoa,
  merges,
  formulas,
  rows,
  projectTitle,
  periodLabel,
  issueDateBE,
  existingPricesMap = new Map()
) {
  const startR = aoa.length;
  const pad12 = (arr) => {
    const copy = [...arr];
    while (copy.length < 12) copy.push('');
    return copy;
  };

  // Group trips by route
  const routeGroupsMap = new Map();
  for (const r of rows) {
    const routeName =
      (r.route || '').trim() ||
      (r.site_name ? `${r.site_name}` : 'ไม่ระบุเส้นทาง');
    if (!routeGroupsMap.has(routeName)) {
      routeGroupsMap.set(routeName, {
        route: routeName,
        trips: 0,
        note: ''
      });
    }
    const g = routeGroupsMap.get(routeName);
    g.trips += 1;
  }

  const routeGroups = Array.from(routeGroupsMap.values());

  // Header rows (0..6 relative to startR)
  aoa.push(pad12(['ใบเบิกผลงานผู้รับเหมา']));
  aoa.push(
    pad12([
      'ผู้รับเหมาชุด    ',
      '',
      'บริษัท จักรพงษ์ ขนส่ง ( 2003 ) จำกัด',
      '',
      '',
      '',
      '',
      'วันที่',
      '',
      issueDateBE
    ])
  );
  aoa.push(
    pad12([
      'โครงการ',
      '',
      projectTitle,
      '',
      '',
      '',
      '',
      'ประจำงวด',
      '',
      periodLabel
    ])
  );
  aoa.push(pad12([]));
  aoa.push(
    pad12([
      'ลำดับ',
      'รายละเอียดงาน',
      '',
      '',
      'ปริมาณงาน',
      '',
      '',
      'จำนวนเงิน',
      '',
      'หมายเหตุ'
    ])
  );
  aoa.push(pad12(['', '', '', '', 'จำนวน', 'หน่วย', 'ราคา/หน่วย', '', '', '']));
  aoa.push(pad12(['', 'ค่ารถรับ -  ส่ง   ', '', '', '', '', '', '', '', '']));

  merges.push(
    { s: { r: startR + 0, c: 0 }, e: { r: startR + 0, c: 9 } }, // A1:J1
    { s: { r: startR + 1, c: 2 }, e: { r: startR + 1, c: 5 } }, // C2:F2
    { s: { r: startR + 1, c: 7 }, e: { r: startR + 1, c: 8 } }, // H2:I2
    { s: { r: startR + 2, c: 2 }, e: { r: startR + 2, c: 5 } }, // C3:F3
    { s: { r: startR + 2, c: 7 }, e: { r: startR + 2, c: 8 } }, // H3:I3
    { s: { r: startR + 4, c: 0 }, e: { r: startR + 5, c: 0 } }, // A5:A6
    { s: { r: startR + 4, c: 1 }, e: { r: startR + 5, c: 3 } }, // B5:D6
    { s: { r: startR + 4, c: 4 }, e: { r: startR + 4, c: 6 } }, // E5:G5
    { s: { r: startR + 4, c: 7 }, e: { r: startR + 5, c: 8 } }, // H5:I6
    { s: { r: startR + 4, c: 9 }, e: { r: startR + 5, c: 9 } }, // J5:J6
    { s: { r: startR + 6, c: 1 }, e: { r: startR + 6, c: 3 } }, // B7:D7
    { s: { r: startR + 6, c: 7 }, e: { r: startR + 6, c: 8 } }  // H7:I7
  );

  const firstDetailExcelRow = startR + 8; // 1-based Excel row 8
  const minDetailRows = Math.max(21, routeGroups.length * 2);
  let itemIdx = 0;
  let totalBillingAmount = 0;

  for (let slot = 0; slot < minDetailRows; slot++) {
    const r0 = aoa.length;
    const excelRow = r0 + 1;
    merges.push(
      { s: { r: r0, c: 1 }, e: { r: r0, c: 3 } }, // B..D
      { s: { r: r0, c: 7 }, e: { r: r0, c: 8 } }  // H..I
    );

    if (slot % 2 === 0 && itemIdx < routeGroups.length) {
      const g = routeGroups[itemIdx];
      itemIdx += 1;
      const normRoute = String(g.route || '').replace(/[\s\-]+/g, '').toLowerCase();
      let savedPrice = existingPricesMap.get(g.route);
      if (savedPrice === undefined) {
        for (const [k, v] of existingPricesMap.entries()) {
          const normK = String(k || '').replace(/[\s\-]+/g, '').toLowerCase();
          if (normK && (normK === normRoute || normK.includes(normRoute) || normRoute.includes(normK))) {
            savedPrice = v;
            break;
          }
        }
      }
      const unitPrice =
        savedPrice !== undefined && savedPrice !== '' ? Number(savedPrice) : '';
      const calcAmount =
        unitPrice !== '' && !Number.isNaN(Number(unitPrice))
          ? g.trips * Number(unitPrice)
          : 0;
      totalBillingAmount += calcAmount;

      aoa.push(
        pad12([
          itemIdx,
          g.route,
          '',
          '',
          g.trips,
          'เที่ยว',
          unitPrice,
          calcAmount,
          '',
          ''
        ])
      );
      formulas.push({
        cell: `H${excelRow}`,
        f: `SUM(E${excelRow}*G${excelRow})`,
        v: calcAmount
      });
    } else {
      aoa.push(pad12([]));
    }
  }

  const lastDetailExcelRow = aoa.length;
  const totalR0 = aoa.length;
  const totalExcelRow = totalR0 + 1;
  aoa.push(pad12(['', 'รวม', '', '', '', '', '', totalBillingAmount, '', '']));
  merges.push(
    { s: { r: totalR0, c: 1 }, e: { r: totalR0, c: 3 } }, // B29:D29
    { s: { r: totalR0, c: 7 }, e: { r: totalR0, c: 8 } }  // H29:I29
  );
  formulas.push({
    cell: `H${totalExcelRow}`,
    f: `SUM(H${firstDetailExcelRow}:I${lastDetailExcelRow})`,
    v: totalBillingAmount
  });

  aoa.push(pad12([]));
  aoa.push(pad12([]));
  aoa.push(pad12(['', '', '', '', 'ผู้ตรวจรับ / Checked by         ']));
  aoa.push(pad12(['', '', '', '', 'วันที่ / Date                            ']));
}

// Section 2: ใบบันทึกการปฏิบัติงานรถ 6 ล้อ รับ - ส่งพนักงาน (แบ่งแผ่นละ 30 แถว แผ่นที่ 1, 2, 3... ตามเส้นทาง)
function appendOperationPagesToAoa(
  aoa,
  merges,
  rows,
  defaultProjectTitle,
  startDateBE,
  endDateBE
) {
  // Group rows by route so each route gets its own 30-row page(s), matching แผ่นที่ 1..8 in sample
  const routeBuckets = new Map();
  if (rows.length === 0) {
    routeBuckets.set('__empty__', []);
  } else {
    for (const r of rows) {
      const key = (r.route || '').trim() || '__default__';
      if (!routeBuckets.has(key)) routeBuckets.set(key, []);
      routeBuckets.get(key).push(r);
    }
  }

  let sheetPageNo = 1;
  for (const [, bucketRows] of routeBuckets.entries()) {
    const pageCount = Math.max(1, Math.ceil(bucketRows.length / 30));
    for (let p = 0; p < pageCount; p++) {
      const pageRows = bucketRows.slice(p * 30, (p + 1) * 30);
      const firstRowSite =
        (pageRows.find((x) => x.site_name)?.site_name || '').trim() || defaultProjectTitle;
      const startR = aoa.length;

      aoa.push([
        'ใบบันทึกการปฏิบัติงานรถ 6 ล้อ รับ - ส่งพนักงาน',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        ''
      ]);
      aoa.push(['', '', '', '', '', '', '', '', '', 'แผ่นที่', sheetPageNo, '']);
      aoa.push([
        'สถานที่ปฏิบัติงานโครงการ :-',
        '',
        '',
        firstRowSite,
        '',
        'ระหว่างวันที่',
        '',
        startDateBE,
        '',
        'ถึงวันที่ ',
        endDateBE,
        ''
      ]);
      aoa.push(['', '', '', '', '', '', '', '', '', '', '', '']);
      aoa.push([
        'ลำดับ',
        'ว/ด/ป',
        'ทะเบียนรถ',
        'จาก',
        'เวลา',
        'ถึง',
        '',
        '',
        'เวลา',
        'จำนวน',
        'ลงชื่อผู้ตรวจ',
        'หมายเหตุ'
      ]);
      aoa.push(['', '', '', '', '( น. )', '', '', '', '( น. )', '(คน)', 'แคมป์ Boss /รปภ.', '']);

      merges.push(
        { s: { r: startR + 0, c: 0 }, e: { r: startR + 0, c: 10 } },
        { s: { r: startR + 1, c: 10 }, e: { r: startR + 1, c: 11 } },
        { s: { r: startR + 2, c: 3 }, e: { r: startR + 2, c: 4 } },
        { s: { r: startR + 2, c: 5 }, e: { r: startR + 2, c: 6 } },
        { s: { r: startR + 2, c: 7 }, e: { r: startR + 2, c: 8 } },
        { s: { r: startR + 2, c: 10 }, e: { r: startR + 2, c: 11 } },
        { s: { r: startR + 4, c: 0 }, e: { r: startR + 5, c: 0 } },
        { s: { r: startR + 4, c: 1 }, e: { r: startR + 5, c: 1 } },
        { s: { r: startR + 4, c: 2 }, e: { r: startR + 5, c: 2 } },
        { s: { r: startR + 4, c: 3 }, e: { r: startR + 5, c: 3 } },
        { s: { r: startR + 4, c: 5 }, e: { r: startR + 5, c: 7 } },
        { s: { r: startR + 4, c: 11 }, e: { r: startR + 5, c: 11 } }
      );

      let prevDateStr = '';
      for (let i = 0; i < 30; i++) {
        const r = pageRows[i];
        const seqNo = p * 30 + i + 1;
        const row0 = aoa.length;
        merges.push({ s: { r: row0, c: 5 }, e: { r: row0, c: 7 } }); // F..H

        if (r) {
          const dateBE = formatShortThaiDateBE(r.travel_date);
          const displayDate = dateBE !== prevDateStr ? dateBE : '';
          if (dateBE) prevDateStr = dateBE;

          const { from, to } = getDirectionalRouteParts(
            r.route,
            r.departure_time || r.arrival_time,
            r.site_name || firstRowSite
          );
          const pax =
            r.passenger_count !== null &&
            r.passenger_count !== undefined &&
            r.passenger_count !== ''
              ? Number(r.passenger_count)
              : '';
          // ช่อง L (หมายเหตุ): ใส่เฉพาะหมายเหตุที่กรอกจริงๆ เท่านั้น ตามต้นฉบับ
          const remarkText = String(r.note || '').trim();
          aoa.push([
            seqNo,
            displayDate,
            r.vehicle_plate || r.worker_name || '',
            from,
            formatExcelTimeDisplay(r.departure_time),
            to,
            '',
            '',
            formatExcelTimeDisplay(r.arrival_time),
            pax,
            '',
            remarkText
          ]);
        } else {
          aoa.push([seqNo, '', '', '', '', '', '', '', '', '', '', '']);
        }
      }

      const f1 = aoa.length;
      const f2 = aoa.length + 1;
      aoa.push([
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'ผู้ตรวจ / Checked by',
        '',
        '',
        '…........................................................................',
        ''
      ]);
      aoa.push([
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'วันที่ / Date',
        '',
        '',
        '…............../…..................../…............',
        ''
      ]);
      merges.push(
        { s: { r: f1, c: 10 }, e: { r: f1, c: 11 } },
        { s: { r: f2, c: 10 }, e: { r: f2, c: 11 } }
      );

      sheetPageNo += 1;
    }
  }
}

// Section 3: บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด - บันทึกรายการรถ 6 ล้อ รับ-ส่งคนงาน + สรุปจำนวนเที่ยวคนขับ/ทะเบียนรถท้ายตาราง
function appendZone6WheelBlocksToAoa(aoa, merges, rows, defaultProjectTitle, formulas = []) {
  const pad12 = (arr) => {
    const copy = [...arr];
    while (copy.length < 12) copy.push('');
    return copy;
  };

  const routeBuckets = new Map();
  if (rows.length === 0) {
    routeBuckets.set('__empty__', []);
  } else {
    for (const r of rows) {
      const key = (r.route || r.site_name || '').trim() || '__default__';
      if (!routeBuckets.has(key)) routeBuckets.set(key, []);
      routeBuckets.get(key).push(r);
    }
  }

  for (const [, bucketRows] of routeBuckets.entries()) {
    const blockProject =
      (bucketRows.find((x) => x.site_name)?.site_name || '').trim() || defaultProjectTitle;
    const startR = aoa.length;

    aoa.push(pad12(['บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด']));
    aoa.push(pad12(['บันทึกรายการรถ 6 ล้อ รับ-ส่งคนงาน']));
    aoa.push(pad12([`โครงการ  ${blockProject}`]));
    aoa.push(pad12([]));
    aoa.push(
      pad12(['ลำดับ', 'วัน เดือน ปี', 'ทะเบียนรถ', 'จาก', 'เวลา', 'จำนวนคน', 'ถึง', 'หมายเหตุ'])
    );

    merges.push(
      { s: { r: startR + 0, c: 0 }, e: { r: startR + 0, c: 7 } },
      { s: { r: startR + 1, c: 0 }, e: { r: startR + 1, c: 7 } },
      { s: { r: startR + 2, c: 0 }, e: { r: startR + 2, c: 7 } }
    );

    let prevDateStr = '';
    const tallyMap = new Map();
    const firstRowRefMap = new Map();

    for (let i = 0; i < bucketRows.length; i++) {
      const r = bucketRows[i];
      const dateBE = formatShortThaiDateBE(r.travel_date);
      const displayDate = dateBE !== prevDateStr ? dateBE : '';
      if (dateBE) prevDateStr = dateBE;

      const tripTime = r.departure_time || r.arrival_time;
      const { from } = getDirectionalRouteParts(
        r.route,
        tripTime,
        r.site_name || blockProject
      );
      const pax =
        r.passenger_count !== null &&
        r.passenger_count !== undefined &&
        r.passenger_count !== ''
          ? Number(r.passenger_count)
          : '';
      // ช่อง C ของไฟล์สำเนา: ใช้ชื่อคนขับ (หรือทะเบียนรถถ้าไม่มีชื่อ) ให้ตรงกับสูตร COUNTIF(C:C,...) ด้านล่าง 100%
      const driverOrPlate = (r.worker_name || r.vehicle_plate || '').trim();
      const currentExcelRow = aoa.length + 1;
      if (driverOrPlate) {
        tallyMap.set(driverOrPlate, (tallyMap.get(driverOrPlate) || 0) + 1);
        if (!firstRowRefMap.has(driverOrPlate)) {
          firstRowRefMap.set(driverOrPlate, currentExcelRow);
        }
      }

      // ช่อง H (หมายเหตุ): ใส่เฉพาะหมายเหตุจริงๆ เท่านั้น และช่อง G (ถึง) ในแถวข้อมูลเว้นว่างตามต้นฉบับสำเนา
      const remarkText = String(r.note || '').trim();

      aoa.push(
        pad12([
          i + 1,
          displayDate,
          driverOrPlate,
          from,
          formatExcelTimeDisplay(tripTime),
          pax,
          '',
          remarkText
        ])
      );
    }

    // 2 blank rows before summary (exactly like lines 45-46 in user's template)
    aoa.push(pad12([bucketRows.length + 1]));
    aoa.push(pad12([bucketRows.length + 2]));

    let seqCursor = bucketRows.length + 3;
    let totalTally = 0;
    const firstSummaryExcelRow = aoa.length + 1;
    for (const [personOrPlate, count] of tallyMap.entries()) {
      totalTally += count;
      const sumExcelRow = aoa.length + 1;
      const refRow = firstRowRefMap.get(personOrPlate) || sumExcelRow;
      aoa.push(pad12([seqCursor, '', '', '', '', personOrPlate, count, '']));
      formulas.push({
        cell: `G${sumExcelRow}`,
        f: `COUNTIF(C:C,C${refRow})`,
        v: count
      });
      seqCursor += 1;
    }
    const lastSummaryExcelRow = aoa.length;
    if (tallyMap.size > 0) {
      const totalExcelRow = aoa.length + 1;
      aoa.push(pad12([seqCursor, '', '', '', '', '', totalTally, '']));
      formulas.push({
        cell: `G${totalExcelRow}`,
        f: `SUM(G${firstSummaryExcelRow}:G${lastSummaryExcelRow})`,
        v: totalTally
      });
      seqCursor += 1;
    }

    // Pad up to at least 70 numbered rows as in template
    while (seqCursor <= Math.max(70, bucketRows.length + tallyMap.size + 15)) {
      aoa.push(pad12([seqCursor]));
      seqCursor += 1;
    }
  }
}

function applyFormulasToWorksheet(ws, formulas) {
  for (const item of formulas) {
    ws[item.cell] = { t: 'n', f: item.f, v: item.v ?? 0 };
  }
}

// Parse existing history & unit prices from an existing Excel workbook buffer
function normalizeExcelCellDateToISO(val) {
  if (val === '' || val === null || val === undefined) return '';
  const numVal = Number(val);
  // Handle Excel serial date numbers (including Thai Buddhist year serials like 244357 = 1/8/2569)
  if (!Number.isNaN(numVal) && numVal > 30000 && numVal < 300000 && /^\d+(\.\d+)?$/.test(String(val).trim())) {
    const d = new Date(Math.round((numVal - 25569) * 86400 * 1000));
    let y = d.getUTCFullYear();
    let month = d.getUTCMonth() + 1;
    let day = d.getUTCDate();
    if (y >= 2400) {
      y -= 543;
      // In Thai BE Excel sheets entered on M/D/Y locale, serial 244357 has month=day(1..12) and day=month(1..12)
      const swappedDay = month;
      const swappedMonth = day;
      day = swappedDay;
      month = swappedMonth;
    }
    return `${y}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  const s = String(val || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let p1 = Number(m[1]);
    let p2 = Number(m[2]);
    let y = Number(m[3]);
    const isThaiShortYear = y >= 60 && y <= 99;
    if (y >= 2400) {
      y -= 543;
    } else if (y < 100) {
      y = isThaiShortYear ? 2500 + y - 543 : 2000 + y;
    }
    let day = p1;
    let month = p2;
    if (month > 12 && day <= 12) {
      day = p2;
      month = p1;
    }
    return `${y}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  return s;
}

function normalizeExcelCellTime(val) {
  if (val === '' || val === null || val === undefined) return '';
  const numVal = Number(val);
  if (!Number.isNaN(numVal) && numVal >= 0 && numVal < 1 && String(val).trim() !== '') {
    const totalMins = Math.round(numVal * 24 * 60);
    const hh = String(Math.floor(totalMins / 60) % 24).padStart(2, '0');
    const mm = String(totalMins % 60).padStart(2, '0');
    return `${hh}:${mm}`;
  }
  const s = String(val).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (m) {
    return `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`;
  }
  return s;
}

function extractExistingDataFromWorkbook(existingWorkbook) {
  const existingRows = [];
  const existingPricesMap = new Map();
  if (!existingWorkbook || !existingWorkbook.SheetNames) {
    return { existingRows, existingPricesMap };
  }

  // 1. Extract unit prices from ใบเบิก sheet if present
  for (const sheetName of existingWorkbook.SheetNames) {
    const ws = existingWorkbook.Sheets[sheetName];
    if (!ws) continue;
    const rows2d = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    for (const r of rows2d) {
      const colB = String(r[1] || '').trim();
      const colF = String(r[5] || '').trim();
      const colG = r[6];
      if (colB && colF === 'เที่ยว' && colG !== '' && colG !== undefined && !Number.isNaN(Number(colG))) {
        existingPricesMap.set(colB, Number(colG));
      }
    }
  }

  // 2. Extract saved history from _CKP_HISTORY sheet if present
  if (existingWorkbook.Sheets['_CKP_HISTORY']) {
    const rawHistory = window.XLSX.utils.sheet_to_json(existingWorkbook.Sheets['_CKP_HISTORY'], {
      defval: ''
    });
    for (const item of rawHistory) {
      if (item.worker_name || item.vehicle_plate || item.route) {
        existingRows.push({
          worker_name: String(item.worker_name || ''),
          site_name: String(item.site_name || ''),
          travel_date: String(item.travel_date || ''),
          route: String(item.route || ''),
          departure_time: String(item.departure_time || ''),
          arrival_time: String(item.arrival_time || ''),
          vehicle_plate: String(item.vehicle_plate || ''),
          passenger_count:
            item.passenger_count !== '' && item.passenger_count !== null && item.passenger_count !== undefined
              ? Number(item.passenger_count)
              : null,
          note: String(item.note || '')
        });
      }
    }
    return { existingRows, existingPricesMap };
  }

  // 3. Fallback: Parse native สำเนา or ใบบันทึก sheets when linking a pre-existing Excel file without _CKP_HISTORY
  for (const sheetName of existingWorkbook.SheetNames) {
    if (
      sheetName === '_CKP_HISTORY' ||
      sheetName === 'รวมสำเนาทุกสาย' ||
      sheetName === 'รวมใบเบิกและใบบันทึก' ||
      sheetName === 'ใบเบิก'
    ) {
      continue;
    }
    const ws = existingWorkbook.Sheets[sheetName];
    if (!ws) continue;
    const rows2d = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rows2d || rows2d.length < 5) continue;

    const cleanSheetRoute = sheetName
      .replace(/^\d+_/, '')
      .replace(/\s*\(\d+\)\s*$/, '')
      .trim();

    let siteName = '';
    let mode = ''; // 'copy' | 'operation'
    let dataStartIndex = -1;

    for (let idx = 0; idx < Math.min(rows2d.length, 12); idx++) {
      const r = rows2d[idx] || [];
      const c0 = String(r[0] || '').trim();
      const c1 = String(r[1] || '').trim();
      const c3 = String(r[3] || '').trim();
      const c5 = String(r[5] || '').trim();

      if (c0.startsWith('โครงการ')) {
        siteName = c0.replace(/^โครงการ\s*/i, '').trim();
      } else if (c0.includes('สถานที่ปฏิบัติงานโครงการ') && c3) {
        siteName = c3;
      }

      if (c0 === 'ลำดับ' && c1.includes('วัน') && c5.includes('จำนวนคน')) {
        mode = 'copy';
        dataStartIndex = idx + 1;
        break;
      }
      if (c0 === 'ลำดับ' && c1.includes('ว/ด/ป')) {
        mode = 'operation';
        dataStartIndex = idx + 2; // skip 2-row header
        break;
      }
    }

    if (!mode || dataStartIndex < 0) continue;

    let lastDateISO = '';
    for (let idx = dataStartIndex; idx < rows2d.length; idx++) {
      const r = rows2d[idx] || [];
      const rawDate = r[1];
      if (rawDate !== '' && rawDate !== null && rawDate !== undefined) {
        lastDateISO = normalizeExcelCellDateToISO(rawDate);
      }

      if (mode === 'copy') {
        const plateOrDriver = String(r[2] || '').trim();
        const fromStr = String(r[3] || '').trim();
        const depTime = normalizeExcelCellTime(r[4]);
        const rawPax = String(r[5] || '').trim();
        const col6 = String(r[6] || '').trim();
        const noteStr = String(r[7] || '').trim();

        // In สำเนา, summary rows at the bottom have empty plate (col 2) and empty from (col 3)
        if (!plateOrDriver && !fromStr) continue;

        const isCol6Time = /^\d{1,2}:\d{2}/.test(col6);
        const arrTime = isCol6Time ? normalizeExcelCellTime(col6) : '';
        const routeStr = cleanSheetRoute || fromStr;
        const paxNum =
          rawPax !== '' && !Number.isNaN(Number(rawPax)) ? Number(rawPax) : null;

        existingRows.push({
          worker_name: plateOrDriver,
          site_name: siteName,
          travel_date: lastDateISO,
          route: routeStr,
          departure_time: depTime,
          arrival_time: arrTime,
          vehicle_plate: plateOrDriver,
          passenger_count: paxNum,
          note: noteStr
        });
      } else if (mode === 'operation') {
        const plate = String(r[2] || '').trim();
        const fromStr = String(r[3] || '').trim();
        const depTime = normalizeExcelCellTime(r[4]);
        const toStr = String(r[5] || '').trim();
        const arrTime = normalizeExcelCellTime(r[8]);
        const rawPax = String(r[9] || '').trim();
        const noteStr = String(r[11] || '').trim();

        if (!plate && !fromStr && !toStr) continue;

        const routeStr =
          cleanSheetRoute && !cleanSheetRoute.startsWith('แผ่นที่')
            ? cleanSheetRoute
            : toStr && toStr !== fromStr
            ? `${fromStr} - ${toStr}`
            : fromStr || toStr;
        const paxNum =
          rawPax !== '' && !Number.isNaN(Number(rawPax)) ? Number(rawPax) : null;

        existingRows.push({
          worker_name: plate,
          site_name: siteName,
          travel_date: lastDateISO,
          route: routeStr,
          departure_time: depTime,
          arrival_time: arrTime,
          vehicle_plate: plate,
          passenger_count: paxNum,
          note: noteStr
        });
      }
    }
  }

  return { existingRows, existingPricesMap };
}

// Merge existing Excel rows with current web app rows (ไหลข้อมูลใหม่ต่อท้ายไฟล์เดิมโดยไม่ตกหล่นและไม่ซ้ำ)
function mergeLogsForExcel(existingRows, currentWebRows) {
  const mergedMap = new Map();
  const makeSignature = (r) =>
    [
      String(r.worker_name || '').trim(),
      String(r.travel_date || '').trim(),
      String(r.departure_time || '').trim(),
      String(r.arrival_time || '').trim(),
      String(r.route || '').trim(),
      String(r.vehicle_plate || '').trim(),
      String(r.passenger_count ?? '').trim(),
      String(r.note || '').trim()
    ].join('||');

  const existingCounts = new Map();
  for (const r of existingRows) {
    const sig = makeSignature(r);
    const occ = (existingCounts.get(sig) || 0) + 1;
    existingCounts.set(sig, occ);
    mergedMap.set(`${sig}##${occ}`, { ...r });
  }

  const webCounts = new Map();
  for (const r of currentWebRows) {
    const sig = makeSignature(r);
    const occ = (webCounts.get(sig) || 0) + 1;
    webCounts.set(sig, occ);
    mergedMap.set(`${sig}##${occ}`, { ...r });
  }

  return Array.from(mergedMap.values()).sort((a, b) => {
    const d1 = a.travel_date || '';
    const d2 = b.travel_date || '';
    if (d1 !== d2) return d1.localeCompare(d2);
    return String(a.departure_time || a.arrival_time || '').localeCompare(
      String(b.departure_time || b.arrival_time || '')
    );
  });
}

function getWorkbookMetaFromRows(rows) {
  const todayISO = getTodayISO();
  const issueDateBE = formatShortThaiDateBE(todayISO);

  const validDates = rows
    .map((r) => r.travel_date)
    .filter(Boolean)
    .sort();
  const startDateBE = validDates.length > 0 ? formatShortThaiDateBE(validDates[0]) : issueDateBE;
  const endDateBE =
    validDates.length > 0
      ? formatShortThaiDateBE(validDates[validDates.length - 1])
      : issueDateBE;
  const periodLabel = startDateBE === endDateBE ? startDateBE : `${startDateBE} - ${endDateBE}`;

  const uniqueSites = Array.from(
    new Set(rows.map((r) => (r.site_name || '').trim()).filter(Boolean))
  );
  const selectedFormSite = (document.getElementById('form-site')?.value || '').trim();
  const uniqueRoutes = Array.from(
    new Set(rows.map((r) => (r.route || '').trim()).filter(Boolean))
  );
  const projectTitle =
    uniqueSites.length === 1
      ? uniqueSites[0]
      : uniqueSites.length > 1
      ? uniqueSites.join(', ')
      : selectedFormSite
      ? selectedFormSite
      : uniqueRoutes.length === 1
      ? uniqueRoutes[0]
      : 'ไลฟ์รัชดา -พระราม9';

  const defaultCols = [
    { wch: 8 },  // A: ลำดับ
    { wch: 15 }, // B: ว/ด/ป / รายละเอียด
    { wch: 16 }, // C: ทะเบียนรถ / โครงการ
    { wch: 22 }, // D: จาก
    { wch: 12 }, // E: เวลา / จำนวน
    { wch: 14 }, // F: ถึง / หน่วย / จำนวนคน
    { wch: 12 }, // G: ราคา/หน่วย / สรุปเที่ยว
    { wch: 14 }, // H: จำนวนเงิน
    { wch: 10 }, // I: เวลาถึง
    { wch: 12 }, // J: จำนวน(คน) / วันที่
    { wch: 20 }, // K: ลงชื่อผู้ตรวจ
    { wch: 20 }  // L: หมายเหตุ
  ];

  return { issueDateBE, startDateBE, endDateBE, periodLabel, projectTitle, defaultCols };
}

function appendHistorySheet(workbook, rows) {
  const wsHistory = window.XLSX.utils.json_to_sheet(
    rows.map((r) => ({
      worker_name: r.worker_name || '',
      site_name: r.site_name || '',
      travel_date: r.travel_date || '',
      route: r.route || '',
      departure_time: r.departure_time || '',
      arrival_time: r.arrival_time || '',
      vehicle_plate: r.vehicle_plate || '',
      passenger_count:
        r.passenger_count !== null && r.passenger_count !== undefined ? r.passenger_count : '',
      note: r.note || ''
    }))
  );
  window.XLSX.utils.book_append_sheet(workbook, wsHistory, '_CKP_HISTORY');
}

function sanitizeExcelSheetName(rawName, fallbackName) {
  const cleaned = String(rawName || '')
    .replace(/[\\/?*[\]:]/g, ' ')
    .trim();
  return (cleaned || fallbackName).slice(0, 28);
}

function getUniqueSheetName(workbook, desiredName) {
  const base = sanitizeExcelSheetName(desiredName, 'Sheet');
  if (!workbook.SheetNames.includes(base)) return base;
  let counter = 2;
  while (counter < 200) {
    const suffix = ` (${counter})`;
    const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    if (!workbook.SheetNames.includes(candidate)) return candidate;
    counter += 1;
  }
  return `${base.slice(0, 24)}_${Date.now() % 1000}`;
}

// 1. สร้างไฟล์แบบ "สำเนา" (แยก Sheet ตามสายรถเหมือนไฟล์ สำเนา27 โซนวิภา + สูตร COUNTIF สรุปจำนวนเที่ยวท้ายตาราง)
function buildCopyOnlyWorkbook(rows) {
  const { projectTitle, defaultCols } = getWorkbookMetaFromRows(rows);
  const workbook = window.XLSX.utils.book_new();

  // แยก Sheet ตามเส้นทาง/โครงการเหมือนไฟล์ สำเนา27 โซนวิภา (ตั้งชื่อแท็บตามสายรถตรงๆ)
  const routeBuckets = new Map();
  if (rows.length === 0) {
    routeBuckets.set('สำเนา', []);
  } else {
    for (const r of rows) {
      const key = (r.route || r.site_name || '').trim() || 'ทั่วไป';
      if (!routeBuckets.has(key)) routeBuckets.set(key, []);
      routeBuckets.get(key).push(r);
    }
  }

  let tabIndex = 1;
  for (const [bucketKey, bucketRows] of routeBuckets.entries()) {
    const singleAoa = [];
    const singleMerges = [];
    const singleFormulas = [];
    const tabProject =
      (bucketRows.find((x) => x.site_name)?.site_name || '').trim() || projectTitle;
    appendZone6WheelBlocksToAoa(
      singleAoa,
      singleMerges,
      bucketRows,
      tabProject,
      singleFormulas
    );
    const wsTab = window.XLSX.utils.aoa_to_sheet(singleAoa);
    applyFormulasToWorksheet(wsTab, singleFormulas);
    wsTab['!merges'] = singleMerges;
    wsTab['!cols'] = defaultCols;
    const sheetName = getUniqueSheetName(
      workbook,
      sanitizeExcelSheetName(bucketKey, `สำเนา_${tabIndex}`)
    );
    window.XLSX.utils.book_append_sheet(workbook, wsTab, sheetName);
    tabIndex += 1;
  }

  // รวมสำเนาทุกสายในหน้าเดียว
  const zoneAoa = [];
  const zoneMerges = [];
  const zoneFormulas = [];
  appendZone6WheelBlocksToAoa(zoneAoa, zoneMerges, rows, projectTitle, zoneFormulas);
  const wsZone = window.XLSX.utils.aoa_to_sheet(zoneAoa);
  applyFormulasToWorksheet(wsZone, zoneFormulas);
  wsZone['!merges'] = zoneMerges;
  wsZone['!cols'] = defaultCols;
  window.XLSX.utils.book_append_sheet(workbook, wsZone, 'รวมสำเนาทุกสาย');

  appendHistorySheet(workbook, rows);
  return workbook;
}

// 2. สร้างไฟล์แบบ "ใบเบิก" (Sheet 1: ใบเบิก + Sheet 2..N: ตั้งชื่อแท็บตามสายรถ แผ่นละ 30 แถว เหมือนต้นฉบับ 100%)
function buildBillingOnlyWorkbook(rows, existingPricesMap = new Map()) {
  const { issueDateBE, startDateBE, endDateBE, periodLabel, projectTitle, defaultCols } =
    getWorkbookMetaFromRows(rows);
  const workbook = window.XLSX.utils.book_new();

  // Sheet 1: ใบเบิก (ใบเบิกผลงานผู้รับเหมา)
  const billAoa = [];
  const billMerges = [];
  const billFormulas = [];
  appendBillingSectionToAoa(
    billAoa,
    billMerges,
    billFormulas,
    rows,
    projectTitle,
    periodLabel,
    issueDateBE,
    existingPricesMap
  );
  const wsBill = window.XLSX.utils.aoa_to_sheet(billAoa);
  applyFormulasToWorksheet(wsBill, billFormulas);
  wsBill['!merges'] = billMerges;
  wsBill['!cols'] = defaultCols;
  window.XLSX.utils.book_append_sheet(workbook, wsBill, 'ใบเบิก');

  // Sheet 2..N: แยกแผ่นละ 30 แถวตามเส้นทาง (ตั้งชื่อแท็บตามสายรถ เช่น "ไลฟ์รัชดา - พระราม9", "ไลฟ์รัชดา - พระราม9 (2)")
  const routeBuckets = new Map();
  if (rows.length === 0) {
    routeBuckets.set(projectTitle || 'แผ่นที่ 1', []);
  } else {
    for (const r of rows) {
      const key = (r.route || r.site_name || '').trim() || projectTitle || 'ทั่วไป';
      if (!routeBuckets.has(key)) routeBuckets.set(key, []);
      routeBuckets.get(key).push(r);
    }
  }

  let pageNo = 1;
  for (const [bucketKey, bucketRows] of routeBuckets.entries()) {
    const pageCount = Math.max(1, Math.ceil(bucketRows.length / 30));
    const baseTabName = sanitizeExcelSheetName(bucketKey, `แผ่นที่ ${pageNo}`);
    for (let p = 0; p < pageCount; p++) {
      const slice30 = bucketRows.slice(p * 30, (p + 1) * 30);
      const pageAoa = [];
      const pageMerges = [];
      appendOperationPagesToAoa(
        pageAoa,
        pageMerges,
        slice30,
        projectTitle,
        startDateBE,
        endDateBE
      );
      if (pageAoa[1]) {
        pageAoa[1][10] = pageNo; // Set แผ่นที่ X in cell K2
      }
      // Adjust sequence numbers for multi-page routes (31..60, 61..90)
      if (p > 0) {
        for (let i = 0; i < 30; i++) {
          if (pageAoa[6 + i]) {
            pageAoa[6 + i][0] = p * 30 + i + 1;
          }
        }
      }
      const wsPage = window.XLSX.utils.aoa_to_sheet(pageAoa);
      wsPage['!merges'] = pageMerges;
      wsPage['!cols'] = defaultCols;
      const desiredTabName = p === 0 ? baseTabName : `${baseTabName.slice(0, 24)} (${p + 1})`;
      const finalTabName = getUniqueSheetName(workbook, desiredTabName);
      window.XLSX.utils.book_append_sheet(workbook, wsPage, finalTabName);
      pageNo += 1;
    }
  }

  // Sheet รวมใบเบิก + ใบบันทึกทุกแผ่นในหน้าเดียว
  const combinedBillingAoa = [];
  const combinedBillingMerges = [];
  const combinedBillingFormulas = [];
  appendBillingSectionToAoa(
    combinedBillingAoa,
    combinedBillingMerges,
    combinedBillingFormulas,
    rows,
    projectTitle,
    periodLabel,
    issueDateBE,
    existingPricesMap
  );
  appendOperationPagesToAoa(
    combinedBillingAoa,
    combinedBillingMerges,
    rows,
    projectTitle,
    startDateBE,
    endDateBE
  );
  const wsCombinedBilling = window.XLSX.utils.aoa_to_sheet(combinedBillingAoa);
  applyFormulasToWorksheet(wsCombinedBilling, combinedBillingFormulas);
  wsCombinedBilling['!merges'] = combinedBillingMerges;
  wsCombinedBilling['!cols'] = defaultCols;
  window.XLSX.utils.book_append_sheet(workbook, wsCombinedBilling, 'รวมใบเบิกและใบบันทึก');

  appendHistorySheet(workbook, rows);
  return workbook;
}

// บันทึกแยกเฉพาะไฟล์สำเนา
async function saveCopyExcelFile() {
  if (typeof window.XLSX === 'undefined') {
    showToast('ไม่พบไลบรารี SheetJS (XLSX)', 'error');
    return;
  }
  if (state.linkedCopyHandle) {
    await writeWorkbookIntoHandleByType('copy', state.linkedCopyHandle, false);
  } else {
    openExcelLinkModalByType('copy');
  }
}

// บันทึกแยกเฉพาะไฟล์ใบเบิก
async function saveBillingExcelFile() {
  if (typeof window.XLSX === 'undefined') {
    showToast('ไม่พบไลบรารี SheetJS (XLSX)', 'error');
    return;
  }
  if (state.linkedBillingHandle) {
    await writeWorkbookIntoHandleByType('billing', state.linkedBillingHandle, false);
  } else {
    openExcelLinkModalByType('billing');
  }
}

async function writeWorkbookIntoHandleByType(type, handle, silentToast = false) {
  const typeLabel = type === 'copy' ? 'แบบสำเนา' : 'แบบใบเบิก';
  try {
    if (handle.queryPermission) {
      let perm = await handle.queryPermission({ mode: 'readwrite' });
      if (perm !== 'granted' && handle.requestPermission) {
        perm = await handle.requestPermission({ mode: 'readwrite' });
      }
      if (perm !== 'granted') {
        showToast(`ไม่ได้รับสิทธิ์เขียนทับไฟล์${typeLabel} กรุณากดอนุญาตสิทธิ์`, 'warning');
        return false;
      }
    }

    let existingRows = [];
    let existingPricesMap = new Map();
    try {
      const existingFile = await handle.getFile();
      if (existingFile && existingFile.size > 0) {
        const arrayBuf = await existingFile.arrayBuffer();
        const existingWb = window.XLSX.read(arrayBuf, { type: 'array' });
        const extracted = extractExistingDataFromWorkbook(existingWb);
        existingRows = extracted.existingRows;
        existingPricesMap = extracted.existingPricesMap;
      }
    } catch {
      // Newly created file or unreadable prior content — start clean
    }

    const currentWebRows = getFilteredLogs();
    const mergedRows = mergeLogsForExcel(existingRows, currentWebRows);
    const workbook =
      type === 'copy'
        ? buildCopyOnlyWorkbook(mergedRows)
        : buildBillingOnlyWorkbook(mergedRows, existingPricesMap);
    const outArrayBuffer = window.XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const writable = await handle.createWritable();
    await writable.write(outArrayBuffer);
    await writable.close();

    await saveLinkedExcelHandleByType(type, handle);
    if (!silentToast) {
      showToast(
        `ไหลข้อมูลเข้าไฟล์${typeLabel} "${handle.name}" สำเร็จ!`,
        'success',
        `อัปเดตข้อมูลสะสมรวมทั้งหมด ${mergedRows.length} รายการ ลงในไฟล์เดิมเรียบร้อยแล้ว`
      );
    }
    return true;
  } catch (err) {
    const msg = String(err?.message || err || '');
    if (
      err?.name === 'NoModificationAllowedError' ||
      err?.name === 'InvalidStateError' ||
      msg.toLowerCase().includes('lock') ||
      msg.toLowerCase().includes('used by another')
    ) {
      openCustomModal(`
        <div class="p-6 space-y-4">
          <div class="flex items-start gap-3.5">
            <div class="w-12 h-12 rounded-2xl bg-amber-100 text-amber-700 flex items-center justify-center text-xl shrink-0">
              <i class="fa-solid fa-file-circle-exclamation"></i>
            </div>
            <div>
              <h3 class="text-base font-bold text-slate-900">ไฟล์ Excel (${escapeHtml(
                typeLabel
              )}) กำลังเปิดค้างอยู่ในโปรแกรม Excel</h3>
              <p class="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
                กรุณา<strong>ปิดโปรแกรม Microsoft Excel</strong> ที่เปิดไฟล์ <strong>"${escapeHtml(
                  handle.name
                )}"</strong> ค้างไว้ก่อน แล้วกดปุ่ม <strong>"ลองไหลข้อมูลอีกครั้ง"</strong> เพื่อให้ระบบบันทึกทับลงไฟล์เดิมได้ครับ
              </p>
            </div>
          </div>
          <div class="pt-3 border-t border-slate-100 flex justify-end gap-2">
            <button
              type="button"
              id="modal-btn-cancel"
              class="min-h-[42px] px-4 py-2 rounded-xl border border-slate-300 bg-white text-slate-700 text-xs sm:text-sm font-medium cursor-pointer"
            >
              ปิดหน้าต่าง
            </button>
            <button
              type="button"
              id="modal-btn-retry-write"
              class="min-h-[42px] px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-semibold cursor-pointer"
            >
              ลองไหลข้อมูลเข้าไฟล์เดิมอีกครั้ง
            </button>
          </div>
        </div>
      `);
      document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);
      document.getElementById('modal-btn-retry-write')?.addEventListener('click', async () => {
        closeCustomModal();
        await writeWorkbookIntoHandleByType(type, handle, silentToast);
      });
      return false;
    }

    showToast(err.message || `ไม่สามารถเขียนข้อมูลลงไฟล์${typeLabel}ได้`, 'error');
    return false;
  }
}

function downloadWorkbookByType(type) {
  const rows = mergeLogsForExcel([], getFilteredLogs());
  const dateSuffix = state.filterDate || getTodayISO();
  if (type === 'copy') {
    const wb = buildCopyOnlyWorkbook(rows);
    const filename = `สำเนา_บันทึกรายการรถ6ล้อ_${dateSuffix}.xlsx`;
    window.XLSX.writeFile(wb, filename);
    showToast(`ดาวน์โหลดไฟล์สำเนาสำเร็จ (${rows.length} รายการ)`, 'success', filename);
  } else {
    const wb = buildBillingOnlyWorkbook(rows, new Map());
    const filename = `ใบเบิก_และใบบันทึกปฏิบัติงานรถ6ล้อ_${dateSuffix}.xlsx`;
    window.XLSX.writeFile(wb, filename);
    showToast(`ดาวน์โหลดไฟล์ใบเบิกสำเร็จ (${rows.length} รายการ)`, 'success', filename);
  }
}

// Modal สำหรับเลือกผูกไฟล์ 1. แบบสำเนา หรือ 2. แบบใบเบิก
function openExcelLinkModalByType(type) {
  const isCopy = type === 'copy';
  const titleText = isCopy
    ? 'ผูกไฟล์ที่ 1: แบบ "สำเนา" (บันทึกรายการรถ 6 ล้อ + สรุปเที่ยว)'
    : 'ผูกไฟล์ที่ 2: แบบ "ใบเบิก" (ใบเบิกผู้รับเหมา + ใบบันทึกปฏิบัติงาน)';
  const subtitleText = isCopy
    ? 'เช่น ไฟล์ "สำเนา27 โซนวิภา ( 1 - 15 ส.ค. 69 ).xlsx" — เมื่อผูกแล้ว กดปุ่มส่งออกข้อมูลจะไหลต่อท้ายเข้าไฟล์สำเนานี้ทันที'
    : 'เช่น ไฟล์ "ใบเบิกไลฟ์รัชดา-พระราม9 ( 1-15-6).xlsx" — เมื่อผูกแล้ว กดปุ่มส่งออกข้อมูลจะไหลเข้าหน้าใบเบิกและใบบันทึกของไฟล์นี้ทันที';
  const currentHandle = isCopy ? state.linkedCopyHandle : state.linkedBillingHandle;

  const supportsFileSystemAccess =
    typeof window.showOpenFilePicker === 'function' &&
    typeof window.showSaveFilePicker === 'function';

  if (!supportsFileSystemAccess) {
    downloadWorkbookByType(type);
    return;
  }

  openCustomModal(`
    <div class="p-6 space-y-4">
      <div class="flex items-start gap-3.5">
        <div class="w-12 h-12 rounded-2xl ${
          isCopy ? 'bg-amber-100 text-amber-700' : 'bg-sky-100 text-sky-700'
        } flex items-center justify-center text-xl shrink-0">
          <i class="fa-solid fa-file-excel"></i>
        </div>
        <div>
          <h3 class="text-base sm:text-lg font-bold text-slate-900">
            ${escapeHtml(titleText)}
          </h3>
          <p class="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
            ${escapeHtml(subtitleText)}
          </p>
          ${
            currentHandle && currentHandle.name
              ? `<p class="mt-2 inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-semibold">
                  <i class="fa-solid fa-check-circle text-emerald-600"></i>
                  <span>ไฟล์ที่ผูกอยู่ปัจจุบัน: <strong>${escapeHtml(currentHandle.name)}</strong></span>
                </p>`
              : ''
          }
        </div>
      </div>

      <div class="space-y-2.5 pt-2">
        ${
          currentHandle && currentHandle.name
            ? `<button
                type="button"
                id="modal-btn-flow-current"
                class="w-full min-h-[48px] px-4 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold shadow-sm flex items-center justify-between gap-3 transition cursor-pointer"
              >
                <span class="flex items-center gap-2.5 text-left">
                  <i class="fa-solid fa-bolt text-amber-300 text-base"></i>
                  <span>ไหลข้อมูลเข้าไฟล์เดิม "${escapeHtml(currentHandle.name)}" ทันที</span>
                </span>
                <i class="fa-solid fa-chevron-right text-xs"></i>
              </button>`
            : ''
        }

        <button
          type="button"
          id="modal-btn-pick-existing"
          class="w-full min-h-[48px] px-4 py-3 rounded-xl ${
            isCopy
              ? 'bg-amber-600 hover:bg-amber-700 text-white'
              : 'bg-sky-600 hover:bg-sky-700 text-white'
          } text-xs sm:text-sm font-bold shadow-sm flex items-center justify-between gap-3 transition cursor-pointer"
        >
          <span class="flex items-center gap-2.5 text-left">
            <i class="fa-solid fa-folder-open text-base"></i>
            <span>1. เลือกไฟล์ Excel "${isCopy ? 'สำเนา' : 'ใบเบิก'}" เดิมที่มีอยู่ในเครื่อง</span>
          </span>
          <i class="fa-solid fa-chevron-right text-xs"></i>
        </button>

        <button
          type="button"
          id="modal-btn-create-new"
          class="w-full min-h-[48px] px-4 py-3 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-white text-xs sm:text-sm font-bold shadow-sm flex items-center justify-between gap-3 transition cursor-pointer"
        >
          <span class="flex items-center gap-2.5 text-left">
            <i class="fa-solid fa-file-circle-plus text-amber-400 text-base"></i>
            <span>2. สร้างไฟล์ Excel "${isCopy ? 'สำเนา' : 'ใบเบิก'}" หลักใหม่ไว้ในเครื่อง</span>
          </span>
          <i class="fa-solid fa-chevron-right text-xs"></i>
        </button>

        <button
          type="button"
          id="modal-btn-download-once"
          class="w-full min-h-[44px] px-4 py-2.5 rounded-xl border border-slate-300 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs sm:text-sm font-semibold flex items-center justify-center gap-2 transition cursor-pointer"
        >
          <i class="fa-solid fa-download"></i>
          <span>ดาวน์โหลดไฟล์ "${isCopy ? 'สำเนา' : 'ใบเบิก'}" แยกเป็นไฟล์ .xlsx ทันที</span>
        </button>
      </div>

      <div class="pt-3 border-t border-slate-100 flex justify-end">
        <button
          type="button"
          id="modal-btn-cancel"
          class="min-h-[40px] px-4 py-1.5 rounded-xl border border-slate-300 bg-white text-slate-600 text-xs font-medium cursor-pointer"
        >
          ปิดหน้าต่าง
        </button>
      </div>
    </div>
  `);

  document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);

  document.getElementById('modal-btn-flow-current')?.addEventListener('click', async () => {
    closeCustomModal();
    if (currentHandle) {
      await writeWorkbookIntoHandleByType(type, currentHandle, false);
    }
  });

  document.getElementById('modal-btn-pick-existing')?.addEventListener('click', async () => {
    closeCustomModal();
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: 'Excel Workbook (.xlsx)',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
            }
          }
        ]
      });
      if (handle) {
        await writeWorkbookIntoHandleByType(type, handle, false);
      }
    } catch (err) {
      if (err?.name !== 'AbortError') {
        showToast('ยกเลิกการเลือกไฟล์ Excel', 'info');
      }
    }
  });

  document.getElementById('modal-btn-create-new')?.addEventListener('click', async () => {
    closeCustomModal();
    try {
      const dateSuffix = state.filterDate || getTodayISO();
      const suggestedName = isCopy
        ? `สำเนา_บันทึกรายการรถ6ล้อ_${dateSuffix}.xlsx`
        : `ใบเบิก_และใบบันทึกปฏิบัติงานรถ6ล้อ_${dateSuffix}.xlsx`;
      const handle = await window.showSaveFilePicker({
        suggestedName,
        types: [
          {
            description: 'Excel Workbook (.xlsx)',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
            }
          }
        ]
      });
      if (handle) {
        await writeWorkbookIntoHandleByType(type, handle, false);
      }
    } catch (err) {
      if (err?.name !== 'AbortError') {
        showToast('ยกเลิกการสร้างไฟล์ Excel', 'info');
      }
    }
  });

  document.getElementById('modal-btn-download-once')?.addEventListener('click', () => {
    closeCustomModal();
    downloadWorkbookByType(type);
  });
}

// ปุ่มหลัก: ไหลข้อมูลเข้าไฟล์ที่ผูกไว้ (ทั้งไฟล์สำเนา และ/หรือ ไฟล์ใบเบิก ในปุ่มเดียว!)
async function exportLogsToExcel() {
  if (typeof window.XLSX === 'undefined') {
    showToast('ไม่พบไลบรารี SheetJS (XLSX) สำหรับสร้างไฟล์ Excel', 'error');
    return;
  }

  const hasCopy = Boolean(state.linkedCopyHandle);
  const hasBilling = Boolean(state.linkedBillingHandle);

  if (hasCopy && hasBilling) {
    const okCopy = await writeWorkbookIntoHandleByType('copy', state.linkedCopyHandle, true);
    const okBilling = await writeWorkbookIntoHandleByType('billing', state.linkedBillingHandle, true);
    if (okCopy && okBilling) {
      showToast(
        'ไหลข้อมูลเข้าทั้ง 2 ไฟล์ (สำเนา + ใบเบิก) สำเร็จ!',
        'success',
        `อัปเดตไฟล์ "${state.linkedCopyHandle.name}" และ "${state.linkedBillingHandle.name}" เรียบร้อยแล้ว`
      );
    }
    return;
  }

  if (hasCopy) {
    await writeWorkbookIntoHandleByType('copy', state.linkedCopyHandle, false);
    return;
  }

  if (hasBilling) {
    await writeWorkbookIntoHandleByType('billing', state.linkedBillingHandle, false);
    return;
  }

  // หากยังไม่ได้ผูกไฟล์ใดเลย ให้เปิดหน้าต่างเลือกผูกไฟล์ทั้ง 2 แบบ
  openCustomModal(`
    <div class="p-6 space-y-4">
      <div class="flex items-start gap-3.5">
        <div class="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center text-xl shrink-0">
          <i class="fa-solid fa-file-excel"></i>
        </div>
        <div>
          <h3 class="text-base sm:text-lg font-bold text-slate-900">
            เลือกผูกไฟล์ Excel 2 แบบฟอร์ม (1. แบบสำเนา / 2. แบบใบเบิก)
          </h3>
          <p class="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
            กรุณาเลือกผูกไฟล์ที่ต้องการอย่างน้อย 1 ไฟล์ (หรือผูกทั้ง 2 ไฟล์) เพื่อให้เวลากดปุ่มนี้ ข้อมูลจะ<strong>ไหลเข้าไฟล์เดิมที่ผูกไว้โดยอัตโนมัติ</strong>ครับ
          </p>
        </div>
      </div>

      <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
        <button
          type="button"
          id="modal-choose-link-copy"
          class="p-4 rounded-2xl border-2 border-amber-300 bg-amber-50/70 hover:bg-amber-100 text-left space-y-1.5 transition cursor-pointer"
        >
          <div class="flex items-center justify-between text-amber-800 font-bold text-sm">
            <span><i class="fa-solid fa-link mr-1.5"></i>1. ผูกไฟล์แบบ "สำเนา"</span>
            <i class="fa-solid fa-chevron-right text-xs"></i>
          </div>
          <p class="text-xs text-amber-900/80">
            บันทึกรายการรถ 6 ล้อ รับ-ส่งคนงาน + สรุปจำนวนเที่ยวคนขับ/ทะเบียนรถ
          </p>
        </button>

        <button
          type="button"
          id="modal-choose-link-billing"
          class="p-4 rounded-2xl border-2 border-sky-300 bg-sky-50/70 hover:bg-sky-100 text-left space-y-1.5 transition cursor-pointer"
        >
          <div class="flex items-center justify-between text-sky-800 font-bold text-sm">
            <span><i class="fa-solid fa-link mr-1.5"></i>2. ผูกไฟล์แบบ "ใบเบิก"</span>
            <i class="fa-solid fa-chevron-right text-xs"></i>
          </div>
          <p class="text-xs text-sky-900/80">
            ใบเบิกผลงานผู้รับเหมา + ใบบันทึกการปฏิบัติงานรถ 6 ล้อ (แผ่นที่ 1..N)
          </p>
        </button>
      </div>

      <div class="pt-2">
        <button
          type="button"
          id="modal-download-both-now"
          class="w-full min-h-[44px] px-4 py-2.5 rounded-xl border border-slate-300 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs sm:text-sm font-semibold flex items-center justify-center gap-2 transition cursor-pointer"
        >
          <i class="fa-solid fa-download"></i>
          <span>ดาวน์โหลดทั้ง 2 ไฟล์ (สำเนา + ใบเบิก) แยกเป็น .xlsx ทันที</span>
        </button>
      </div>

      <div class="pt-3 border-t border-slate-100 flex justify-end">
        <button
          type="button"
          id="modal-btn-cancel"
          class="min-h-[40px] px-4 py-1.5 rounded-xl border border-slate-300 bg-white text-slate-600 text-xs font-medium cursor-pointer"
        >
          ปิดหน้าต่าง
        </button>
      </div>
    </div>
  `);

  document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-choose-link-copy')?.addEventListener('click', () => {
    closeCustomModal();
    openExcelLinkModalByType('copy');
  });
  document.getElementById('modal-choose-link-billing')?.addEventListener('click', () => {
    closeCustomModal();
    openExcelLinkModalByType('billing');
  });
  document.getElementById('modal-download-both-now')?.addEventListener('click', () => {
    closeCustomModal();
    downloadWorkbookByType('copy');
    setTimeout(() => downloadWorkbookByType('billing'), 350);
  });
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

  // 1B. Sites (ไซต์งาน / โครงการ หัวตาราง)
  const sitesListEl = document.getElementById('master-sites-list');
  const badgeSites = document.getElementById('badge-count-sites');
  if (badgeSites) badgeSites.textContent = `${state.allSites.length} ไซต์งาน`;

  if (sitesListEl) {
    sitesListEl.innerHTML = state.allSites
      .map((s) => {
        const logsUsingCount = state.logs.filter((l) => (l.site_name || '') === s.name).length;
        return `
        <div class="p-3.5 hover:bg-slate-50 flex items-center justify-between gap-2 transition">
          <div class="min-w-0 flex-1">
            <div class="flex items-center gap-2">
              <span class="font-semibold text-sm ${
                s.active ? 'text-slate-900' : 'text-slate-400 line-through'
              } truncate">${escapeHtml(s.name)}</span>
              <span class="px-2 py-0.5 rounded-full text-[10px] font-semibold ${
                s.active ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'
              }">
                ${s.active ? 'เปิดใช้งาน' : 'ปิดใช้งาน'}
              </span>
            </div>
            <p class="text-[11px] text-slate-500 mt-0.5">
              บันทึกในไซต์งานนี้: <strong>${logsUsingCount} รายการ</strong>
            </p>
          </div>

          <div class="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              data-site-toggle="${s.id}"
              class="px-2 py-1 rounded-lg border text-xs font-semibold transition cursor-pointer ${
                s.active
                  ? 'bg-emerald-50 border-emerald-300 text-emerald-700 hover:bg-emerald-100'
                  : 'bg-slate-100 border-slate-300 text-slate-500 hover:bg-slate-200'
              }"
            >
              <i class="fa-solid ${s.active ? 'fa-toggle-on text-sm' : 'fa-toggle-off text-sm'}"></i>
            </button>

            <button
              type="button"
              data-site-edit="${s.id}"
              class="p-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-pen"></i>
            </button>

            <button
              type="button"
              data-site-delete="${s.id}"
              class="p-1.5 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs transition cursor-pointer"
            >
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </div>
      `;
      })
      .join('');

    sitesListEl.querySelectorAll('[data-site-toggle]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-site-toggle'));
        const item = state.allSites.find((x) => x.id === id);
        if (item) toggleMasterItemActive('sites', item);
      });
    });

    sitesListEl.querySelectorAll('[data-site-edit]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-site-edit'));
        const item = state.allSites.find((x) => x.id === id);
        if (item) openEditSiteModal(item);
      });
    });

    sitesListEl.querySelectorAll('[data-site-delete]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = Number(btn.getAttribute('data-site-delete'));
        const item = state.allSites.find((x) => x.id === id);
        if (item) confirmDeleteMasterItem('sites', item.id, item.name);
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

function openEditSiteModal(siteItem) {
  openCustomModal(`
    <form id="modal-form-edit-site" class="p-6 space-y-4">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-sky-100 text-sky-700 flex items-center justify-center">
          <i class="fa-solid fa-building-user"></i>
        </div>
        <div>
          <h3 class="text-base font-bold text-slate-900">แก้ไขชื่อไซต์งาน / โครงการ (หัวตาราง)</h3>
          <p class="text-xs text-slate-500">อัปเดตเชื่อมโยงไปยังตาราง <code>travel_logs</code> อัตโนมัติ</p>
        </div>
      </div>

      <div>
        <label class="block text-xs font-semibold text-slate-700 mb-1">ชื่อไซต์งาน / โครงการ *</label>
        <input
          type="text"
          id="modal-edit-site-name"
          value="${escapeHtml(siteItem.name)}"
          required
          class="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-sky-600"
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
          class="min-h-[42px] px-5 py-2 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs sm:text-sm font-semibold cursor-pointer"
        >
          บันทึกการแก้ไข
        </button>
      </div>
    </form>
  `);

  document.getElementById('modal-btn-cancel')?.addEventListener('click', closeCustomModal);
  document.getElementById('modal-form-edit-site')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = document.getElementById('modal-edit-site-name').value.trim();
    if (!name) return;

    const res = await apiFetch(`/api/master/sites/${siteItem.id}`, {
      method: 'PUT',
      body: JSON.stringify({ name })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถแก้ไขไซต์งานได้', 'error');
      return;
    }

    closeCustomModal();
    showToast(
      `อัปเดตไซต์งาน "${name}" สำเร็จ`,
      'success',
      `อัปเดตเชื่อมโยงบันทึกเดินทาง ${res.data?.cascadedLogsCount || 0} รายการ`
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
    sites: 'ไซต์งาน / โครงการ',
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

  // 11. Table Filters & Excel Export
  const filterSiteSelect = document.getElementById('filter-site');
  const filterDateInput = document.getElementById('filter-date');
  const filterSearchInput = document.getElementById('filter-search');
  const btnClearSearch = document.getElementById('btn-clear-search');

  filterSiteSelect?.addEventListener('change', () => {
    state.filterSite = filterSiteSelect.value;
    renderLogsTable();
  });

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
  document.getElementById('btn-save-copy-excel')?.addEventListener('click', saveCopyExcelFile);
  document.getElementById('btn-save-billing-excel')?.addEventListener('click', saveBillingExcelFile);
  document.getElementById('btn-link-copy-excel')?.addEventListener('click', () => {
    pickAndLinkExcelFileOnly('copy');
  });
  document.getElementById('btn-link-billing-excel')?.addEventListener('click', () => {
    pickAndLinkExcelFileOnly('billing');
  });
  loadLinkedExcelHandles();

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

  document.getElementById('form-add-site')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const nameInput = document.getElementById('new-site-name');
    const name = nameInput.value.trim();
    if (!name) return;

    const res = await apiFetch('/api/master/sites', {
      method: 'POST',
      body: JSON.stringify({ name, active: 1 })
    });

    if (!res.ok) {
      showToast(res.data?.error || 'ไม่สามารถเพิ่มไซต์งานได้', 'error');
      return;
    }

    nameInput.value = '';
    showToast(`เพิ่มไซต์งาน/โครงการ "${name}" สำเร็จ`, 'success');
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
