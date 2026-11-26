const { db, supabase, isSupabaseConfigured, verifyPassword, seedDatabase } = require('./db');

seedDatabase();

const useSupabase = Boolean(isSupabaseConfigured && supabase);

// ตารางจริงที่ใช้เก็บข้อมูล (ใช้ตารางเดิมที่มีอยู่แล้วใน Supabase / SQLite เพื่อให้ใช้งานได้ทันทีโดยไม่ต้องรัน SQL เพิ่ม)
//  - คนขับ      → master_workers   (name, default_route ← ใช้เก็บ default_vehicle, active)
//  - ทะเบียนรถ  → master_vehicles  (plate ← license_plate, active)
//  - โครงการ    → system_settings  (key = v3_sites,  JSON)
//  - เส้นทาง    → system_settings  (key = v3_routes, JSON)
//  - เที่ยววิ่ง → travel_logs      (worker_name ← id ของรายการ, gps_status ← JSON เต็มรูปแบบ, คอลัมน์อื่นเก็บค่าสำคัญไว้ดูใน Dashboard)
const T = {
  workers: 'master_workers',
  vehicles: 'master_vehicles',
  logs: 'travel_logs',
  settings: 'system_settings'
};
const KV_SITES = 'v3_sites';
const KV_ROUTES = 'v3_routes';
const KV_WORKER_BINDINGS = 'v3_worker_bindings';
const KV_GSHEET = 'v3_gsheet_config';
const KV_GSHEET_HISTORY = 'v3_gsheet_history';
const MAX_GSHEET_HISTORY = 100;
const LOG_VERSION = 3;
const DEFAULT_MAX_LIMIT = 10000;

// ============================================================================
// Helpers
// ============================================================================

function httpError(status, message, extra = {}) {
  const err = new Error(message);
  err.status = status;
  Object.assign(err, extra);
  return err;
}

const str = (v) => String(v ?? '').trim();

function normalizeTime(v, label) {
  const s = str(v);
  if (!s) return '';
  // รองรับทั้งแบบโคลอน 15:00 และแบบจุด 15.00 หรือติดกัน 1500
  let m = s.match(/^(\d{1,2})[:.](\d{2})$/);
  if (!m) {
    m = s.match(/^(\d{2})(\d{2})$/);
  }
  if (!m || Number(m[1]) > 23 || Number(m[2]) > 59) {
    throw httpError(400, `รูปแบบ${label}ไม่ถูกต้อง (เช่น 15:00 หรือ 15.00)`);
  }
  return `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`;
}

function normalizeMoney(v, label) {
  if (v === undefined || v === null || str(v) === '') return 0;
  const n = Number(v);
  if (!Number.isFinite(n) || n < 0) throw httpError(400, `${label}ต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป`);
  return Math.round(n * 100) / 100;
}

function normalizeIntOrNull(v, label) {
  if (v === undefined || v === null || str(v) === '') return null;
  const n = Number.parseInt(String(v), 10);
  if (!Number.isFinite(n) || n < 0) throw httpError(400, `${label}ต้องเป็นจำนวนเต็มตั้งแต่ 0 ขึ้นไป`);
  return n;
}

function normalizeDisplayMode(v) {
  return str(v) === 'license_plate' ? 'license_plate' : 'worker_name';
}

function makeId(prefix) {
  return `${prefix}${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
}

// ============================================================================
// Low-level table access (Supabase หรือ SQLite) — ใช้ร่วมกันทุกเอนทิตี
// ============================================================================

async function selectAll(table, orderCol) {
  if (useSupabase) {
    const out = [];
    const page = 1000;
    for (let from = 0; ; from += page) {
      const { data, error } = await supabase
        .from(table)
        .select('*')
        .order(orderCol, { ascending: true })
        .range(from, from + page - 1);
      if (error) throw new Error(error.message);
      out.push(...(data || []));
      if (!data || data.length < page) break;
    }
    return out;
  }
  return db.prepare(`SELECT * FROM ${table} ORDER BY ${orderCol} ASC`).all();
}

async function selectOne(table, keyCol, key) {
  if (useSupabase) {
    const { data, error } = await supabase.from(table).select('*').eq(keyCol, key).maybeSingle();
    if (error) throw new Error(error.message);
    return data || null;
  }
  return db.prepare(`SELECT * FROM ${table} WHERE ${keyCol} = ?`).get(key) || null;
}

function toSqlValue(v) {
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === undefined) return null;
  return v;
}

async function insertRow(table, row, keyCol) {
  if (useSupabase) {
    let { data, error } = await supabase.from(table).insert(row).select('*').single();
    if (error && error.code === '23505' && String(error.message).includes('_pkey') && keyCol === 'id') {
      const { data: maxRow } = await supabase
        .from(table)
        .select('id')
        .order('id', { ascending: false })
        .limit(1)
        .maybeSingle();
      const retry = await supabase
        .from(table)
        .insert({ ...row, id: (maxRow?.id || 0) + 1 })
        .select('*')
        .single();
      data = retry.data;
      error = retry.error;
    }
    if (error) throw new Error(error.message);
    return data;
  }
  const cols = Object.keys(row);
  const result = db
    .prepare(`INSERT INTO ${table} (${cols.join(', ')}) VALUES (${cols.map(() => '?').join(', ')})`)
    .run(...cols.map((c) => toSqlValue(row[c])));
  const key = keyCol === 'id' ? result.lastInsertRowid : row[keyCol];
  return db.prepare(`SELECT * FROM ${table} WHERE ${keyCol} = ?`).get(key);
}

async function insertRows(table, rows) {
  if (rows.length === 0) return;
  if (useSupabase) {
    const chunk = 500;
    for (let i = 0; i < rows.length; i += chunk) {
      const { error } = await supabase.from(table).insert(rows.slice(i, i + chunk));
      if (error) throw new Error(error.message);
    }
    return;
  }
  for (const r of rows) {
    await insertRow(table, r, 'worker_name');
  }
}

async function updateRow(table, keyCol, key, patch) {
  if (useSupabase) {
    const { data, error } = await supabase
      .from(table)
      .update(patch)
      .eq(keyCol, key)
      .select('*')
      .single();
    if (error) throw new Error(error.message);
    return data;
  }
  const cols = Object.keys(patch);
  db.prepare(`UPDATE ${table} SET ${cols.map((c) => `${c} = ?`).join(', ')} WHERE ${keyCol} = ?`).run(
    ...cols.map((c) => toSqlValue(patch[c])),
    key
  );
  const newKey = patch[keyCol] !== undefined ? patch[keyCol] : key;
  return db.prepare(`SELECT * FROM ${table} WHERE ${keyCol} = ?`).get(newKey);
}

async function deleteRow(table, keyCol, key) {
  if (useSupabase) {
    const { error } = await supabase.from(table).delete().eq(keyCol, key);
    if (error) throw new Error(error.message);
    return;
  }
  db.prepare(`DELETE FROM ${table} WHERE ${keyCol} = ?`).run(key);
}

async function deleteAllRows(table, keyCol) {
  if (useSupabase) {
    const { error } = await supabase.from(table).delete().neq(keyCol, '__never_match__');
    if (error) throw new Error(error.message);
    return;
  }
  db.prepare(`DELETE FROM ${table}`).run();
}

async function countRows(table) {
  if (useSupabase) {
    const { count, error } = await supabase.from(table).select('*', { count: 'exact', head: true });
    if (error) throw new Error(error.message);
    return Number(count || 0);
  }
  return Number(db.prepare(`SELECT COUNT(*) AS c FROM ${table}`).get().c);
}

async function getKV(key, fallback) {
  let raw = null;
  if (useSupabase) {
    const { data, error } = await supabase
      .from(T.settings)
      .select('value')
      .eq('key', key)
      .maybeSingle();
    if (error) throw new Error(error.message);
    raw = data ? data.value : null;
  } else {
    const row = db.prepare(`SELECT value FROM ${T.settings} WHERE key = ?`).get(key);
    raw = row ? row.value : null;
  }
  if (raw === null || raw === undefined || raw === '') return fallback;
  try {
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

async function getRawKV(key) {
  if (useSupabase) {
    const { data } = await supabase.from(T.settings).select('value').eq('key', key).maybeSingle();
    return data ? data.value : null;
  }
  const row = db.prepare(`SELECT value FROM ${T.settings} WHERE key = ?`).get(key);
  return row ? row.value : null;
}

async function setRawKV(key, value) {
  if (useSupabase) {
    const { error } = await supabase
      .from(T.settings)
      .upsert([{ key, value: String(value) }], { onConflict: 'key' });
    if (error) throw new Error(error.message);
    return;
  }
  db.prepare(`INSERT OR REPLACE INTO ${T.settings} (key, value) VALUES (?, ?)`).run(key, String(value));
}

async function setKV(key, value) {
  await setRawKV(key, JSON.stringify(value));
}

// ============================================================================
// Users
// ============================================================================

async function findUserByUsername(username) {
  if (useSupabase) {
    const { data, error } = await supabase
      .from('users')
      .select('id, username, password, full_name, role')
      .eq('username', username)
      .maybeSingle();
    if (error) throw new Error(error.message);
    return data || null;
  }
  return (
    db
      .prepare('SELECT id, username, password, full_name, role FROM users WHERE username = ?')
      .get(username) || null
  );
}

async function findUserById(id) {
  if (useSupabase) {
    const { data, error } = await supabase
      .from('users')
      .select('id, username, full_name, role')
      .eq('id', id)
      .maybeSingle();
    if (error) throw new Error(error.message);
    return data || null;
  }
  return db.prepare('SELECT id, username, full_name, role FROM users WHERE id = ?').get(id) || null;
}

// ============================================================================
// Settings & Storage quota
// ============================================================================

async function getSettings() {
  const raw = await getRawKV('max_storage_limit');
  const n = Number(raw);
  return { max_storage_limit: Number.isFinite(n) && n >= 1 ? n : DEFAULT_MAX_LIMIT };
}

async function saveSettings({ max_storage_limit }) {
  const n = Math.max(1, Number.parseInt(max_storage_limit, 10) || DEFAULT_MAX_LIMIT);
  await setRawKV('max_storage_limit', n);
  return getSettings();
}

async function getQuotaStatus() {
  const { max_storage_limit: maxLimit } = await getSettings();
  const currentCount = await countRows(T.logs);
  const usagePercent = Math.min(100, Math.round((currentCount / maxLimit) * 100));
  let statusLevel = 'normal';
  if (currentCount >= maxLimit) statusLevel = 'full';
  else if (usagePercent >= 80) statusLevel = 'warning';
  return {
    currentCount,
    maxLimit,
    usagePercent,
    statusLevel,
    isFull: currentCount >= maxLimit,
    isWarning: statusLevel === 'warning',
    storageMode: useSupabase ? 'supabase_cloud' : 'local_sqlite'
  };
}

// ============================================================================
// Master: Workers  (id, name, default_vehicle, default_site, default_origin, default_destination, active)
// ============================================================================

async function getWorkerBindingsMap() {
  const raw = await getKV(KV_WORKER_BINDINGS, {});
  return raw && typeof raw === 'object' && !Array.isArray(raw) ? raw : {};
}

async function saveWorkerBinding(workerId, binding) {
  const map = await getWorkerBindingsMap();
  const key = String(workerId);
  const clean = {
    default_site: str(binding && binding.default_site),
    default_origin: str(binding && binding.default_origin),
    default_destination: str(binding && binding.default_destination)
  };
  if (clean.default_site || clean.default_origin || clean.default_destination) {
    map[key] = clean;
  } else {
    delete map[key];
  }
  await setKV(KV_WORKER_BINDINGS, map);
  return clean;
}

const workerFromRow = (r, bindingsMap = {}) => {
  const b = (bindingsMap && bindingsMap[String(r.id)]) || {};
  return {
    id: Number(r.id),
    name: r.name,
    default_vehicle: r.default_route || '',
    default_site: str(b.default_site),
    default_origin: str(b.default_origin),
    default_destination: str(b.default_destination),
    active: Boolean(r.active)
  };
};

async function listWorkers(activeOnly = false) {
  const [rawRows, bindingsMap] = await Promise.all([selectAll(T.workers, 'id'), getWorkerBindingsMap()]);
  const rows = rawRows.map((r) => workerFromRow(r, bindingsMap));
  return activeOnly ? rows.filter((r) => r.active) : rows;
}

async function assertVehicleExistsIfSet(plate) {
  if (!plate) return;
  const vehicles = await listVehicles();
  if (!vehicles.some((v) => v.license_plate === plate)) {
    throw httpError(400, `ไม่พบทะเบียนรถ "${plate}" ในตารางทะเบียนรถ กรุณาเพิ่มทะเบียนรถก่อน`);
  }
}

async function createWorker({ name, default_vehicle, default_site, default_origin, default_destination, active }) {
  const cleanName = str(name);
  if (!cleanName) throw httpError(400, 'กรุณาระบุชื่อคนขับ');
  const plate = str(default_vehicle);
  await assertVehicleExistsIfSet(plate);
  const existing = await listWorkers();
  if (existing.some((w) => w.name === cleanName)) {
    throw httpError(409, `ชื่อคนขับ "${cleanName}" มีอยู่ในระบบแล้ว`);
  }
  const row = await insertRow(
    T.workers,
    { name: cleanName, default_route: plate, active: active === undefined ? true : Boolean(active) },
    'id'
  );
  const binding = await saveWorkerBinding(row.id, {
    default_site,
    default_origin,
    default_destination
  });
  return workerFromRow(row, { [String(row.id)]: binding });
}

async function updateWorker(id, { name, default_vehicle, default_site, default_origin, default_destination, active }) {
  const all = await listWorkers();
  const existing = all.find((w) => w.id === Number(id));
  if (!existing) throw httpError(404, 'ไม่พบคนขับที่ต้องการแก้ไข');

  const newName = name !== undefined ? str(name) : existing.name;
  if (!newName) throw httpError(400, 'ชื่อคนขับห้ามเว้นว่าง');
  const newPlate = default_vehicle !== undefined ? str(default_vehicle) : existing.default_vehicle;
  if (default_vehicle !== undefined) await assertVehicleExistsIfSet(newPlate);
  if (newName !== existing.name && all.some((w) => w.name === newName)) {
    throw httpError(409, `ชื่อคนขับ "${newName}" มีอยู่ในระบบแล้ว`);
  }

  const row = await updateRow(T.workers, 'id', existing.id, {
    name: newName,
    default_route: newPlate,
    active: active !== undefined ? Boolean(active) : existing.active
  });

  const binding = await saveWorkerBinding(existing.id, {
    default_site: default_site !== undefined ? default_site : existing.default_site,
    default_origin: default_origin !== undefined ? default_origin : existing.default_origin,
    default_destination: default_destination !== undefined ? default_destination : existing.default_destination
  });

  let cascadedLogsCount = 0;
  if (newName !== existing.name) {
    cascadedLogsCount = await cascadeLogField('worker_name', existing.name, newName);
  }
  return { ...workerFromRow(row, { [String(existing.id)]: binding }), cascadedLogsCount };
}

async function deleteWorker(id) {
  const existing = (await listWorkers()).find((w) => w.id === Number(id));
  if (!existing) throw httpError(404, 'ไม่พบคนขับที่ต้องการลบ');
  await deleteRow(T.workers, 'id', existing.id);
  const map = await getWorkerBindingsMap();
  if (map[String(existing.id)]) {
    delete map[String(existing.id)];
    await setKV(KV_WORKER_BINDINGS, map);
  }
  return existing;
}

// ============================================================================
// Master: Vehicles  (id, license_plate, active)
// ============================================================================

const vehicleFromRow = (r) => ({
  id: Number(r.id),
  license_plate: r.plate,
  active: Boolean(r.active)
});

async function listVehicles(activeOnly = false) {
  const rows = (await selectAll(T.vehicles, 'id')).map(vehicleFromRow);
  return activeOnly ? rows.filter((r) => r.active) : rows;
}

async function createVehicle({ license_plate, active }) {
  const plate = str(license_plate);
  if (!plate) throw httpError(400, 'กรุณาระบุเลขทะเบียนรถ');
  const existing = await listVehicles();
  if (existing.some((v) => v.license_plate === plate)) {
    throw httpError(409, `ทะเบียนรถ "${plate}" มีอยู่ในระบบแล้ว`);
  }
  const row = await insertRow(
    T.vehicles,
    { plate, active: active === undefined ? true : Boolean(active) },
    'id'
  );
  return vehicleFromRow(row);
}

async function updateVehicle(id, { license_plate, active }) {
  const all = await listVehicles();
  const existing = all.find((v) => v.id === Number(id));
  if (!existing) throw httpError(404, 'ไม่พบทะเบียนรถที่ต้องการแก้ไข');
  const newPlate = license_plate !== undefined ? str(license_plate) : existing.license_plate;
  if (!newPlate) throw httpError(400, 'เลขทะเบียนรถห้ามเว้นว่าง');
  if (newPlate !== existing.license_plate && all.some((v) => v.license_plate === newPlate)) {
    throw httpError(409, `ทะเบียนรถ "${newPlate}" มีอยู่ในระบบแล้ว`);
  }
  const row = await updateRow(T.vehicles, 'id', existing.id, {
    plate: newPlate,
    active: active !== undefined ? Boolean(active) : existing.active
  });

  let cascadedLogsCount = 0;
  if (newPlate !== existing.license_plate) {
    cascadedLogsCount = await cascadeLogField('license_plate', existing.license_plate, newPlate);
    for (const w of await listWorkers()) {
      if (w.default_vehicle === existing.license_plate) {
        await updateRow(T.workers, 'id', w.id, { default_route: newPlate });
      }
    }
  }
  return { ...vehicleFromRow(row), cascadedLogsCount };
}

async function deleteVehicle(id) {
  const existing = (await listVehicles()).find((v) => v.id === Number(id));
  if (!existing) throw httpError(404, 'ไม่พบทะเบียนรถที่ต้องการลบ');
  await deleteRow(T.vehicles, 'id', existing.id);
  for (const w of await listWorkers()) {
    if (w.default_vehicle === existing.license_plate) {
      await updateRow(T.workers, 'id', w.id, { default_route: '' });
    }
  }
  return existing;
}

// ============================================================================
// Master: Sites  (id, site_name, zone_name, default_origin, active)  — เก็บใน system_settings
// ============================================================================

const siteFromItem = (s) => ({
  id: Number(s.id),
  site_name: String(s.site_name || ''),
  zone_name: String(s.zone_name || ''),
  default_origin: String(s.default_origin || ''),
  active: s.active === undefined ? true : Boolean(s.active)
});

async function listSites(activeOnly = false) {
  const items = (await getKV(KV_SITES, [])).map(siteFromItem);
  return activeOnly ? items.filter((s) => s.active) : items;
}

async function createSite({ site_name, zone_name, default_origin, active }) {
  const name = str(site_name);
  if (!name) throw httpError(400, 'กรุณาระบุชื่อไซต์งาน / โครงการ');
  const items = await listSites();
  if (items.some((s) => s.site_name === name)) {
    throw httpError(409, `ไซต์งาน/โครงการ "${name}" มีอยู่ในระบบแล้ว`);
  }
  const created = {
    id: items.reduce((m, s) => Math.max(m, s.id), 0) + 1,
    site_name: name,
    zone_name: str(zone_name),
    default_origin: str(default_origin),
    active: active === undefined ? true : Boolean(active)
  };
  items.push(created);
  await setKV(KV_SITES, items);
  return created;
}

async function updateSite(id, { site_name, zone_name, default_origin, active }) {
  const items = await listSites();
  const idx = items.findIndex((s) => s.id === Number(id));
  if (idx === -1) throw httpError(404, 'ไม่พบไซต์งาน/โครงการที่ต้องการแก้ไข');
  const existing = items[idx];
  const newName = site_name !== undefined ? str(site_name) : existing.site_name;
  if (!newName) throw httpError(400, 'ชื่อไซต์งาน/โครงการห้ามเว้นว่าง');
  if (newName !== existing.site_name && items.some((s) => s.site_name === newName)) {
    throw httpError(409, `ไซต์งาน/โครงการ "${newName}" มีอยู่ในระบบแล้ว`);
  }
  items[idx] = {
    id: existing.id,
    site_name: newName,
    zone_name: zone_name !== undefined ? str(zone_name) : existing.zone_name,
    default_origin: default_origin !== undefined ? str(default_origin) : existing.default_origin,
    active: active !== undefined ? Boolean(active) : existing.active
  };
  await setKV(KV_SITES, items);
  let cascadedLogsCount = 0;
  if (newName !== existing.site_name) {
    cascadedLogsCount = await cascadeLogField('site_name', existing.site_name, newName);
    // อัปเดตโครงการที่ผูกไว้ในคนขับและเส้นทางด้วย
    const bMap = await getWorkerBindingsMap();
    let bChanged = false;
    for (const k of Object.keys(bMap)) {
      if (bMap[k] && bMap[k].default_site === existing.site_name) {
        bMap[k].default_site = newName;
        bChanged = true;
      }
      if (bMap[k] && bMap[k].default_destination === existing.site_name) {
        bMap[k].default_destination = newName;
        bChanged = true;
      }
    }
    if (bChanged) await setKV(KV_WORKER_BINDINGS, bMap);

    const routes = await listRoutes();
    let rChanged = false;
    for (const r of routes) {
      if (r.default_site === existing.site_name) {
        r.default_site = newName;
        rChanged = true;
      }
      if (r.destination === existing.site_name) {
        r.destination = newName;
        rChanged = true;
      }
    }
    if (rChanged) await setKV(KV_ROUTES, routes);
  }
  return { ...items[idx], cascadedLogsCount };
}

async function deleteSite(id) {
  const items = await listSites();
  const idx = items.findIndex((s) => s.id === Number(id));
  if (idx === -1) throw httpError(404, 'ไม่พบไซต์งาน/โครงการที่ต้องการลบ');
  const [deleted] = items.splice(idx, 1);
  await setKV(KV_SITES, items);
  return deleted;
}

// ============================================================================
// Master: Routes — เส้นทาง / ที่หมาย  — เก็บใน system_settings
// ============================================================================

const routeFromItem = (r) => ({
  id: Number(r.id),
  origin: String(r.origin || ''),
  destination: String(r.destination || ''),
  default_site: String(r.default_site || ''),
  route_sheet_name: String(r.route_sheet_name || ''),
  billing_title: String(r.billing_title || ''),
  default_dep_time: String(r.default_dep_time || ''),
  default_arr_time: String(r.default_arr_time || ''),
  unit_price: Number(r.unit_price) || 0,
  display_mode_zone: normalizeDisplayMode(r.display_mode_zone),
  active: r.active === undefined ? true : Boolean(r.active)
});

async function listRoutes(activeOnly = false) {
  const items = (await getKV(KV_ROUTES, [])).map(routeFromItem);
  return activeOnly ? items.filter((r) => r.active) : items;
}

function buildRouteRecord(input, base = {}) {
  const origin = input.origin !== undefined ? str(input.origin) : base.origin || '';
  const destination =
    input.destination !== undefined ? str(input.destination) : base.destination || '';
  const default_site =
    input.default_site !== undefined ? str(input.default_site) : base.default_site || '';
  if (!origin) throw httpError(400, 'กรุณาระบุต้นทาง (จุดรับ) ของเส้นทาง');
  const defaultSheet = destination ? `${origin} - ${destination}` : origin;
  const sheet =
    (input.route_sheet_name !== undefined ? str(input.route_sheet_name) : base.route_sheet_name) ||
    defaultSheet;
  const billing =
    (input.billing_title !== undefined ? str(input.billing_title) : base.billing_title) || sheet;
  return {
    origin,
    destination,
    default_site,
    route_sheet_name: sheet,
    billing_title: billing,
    default_dep_time:
      input.default_dep_time !== undefined
        ? normalizeTime(input.default_dep_time, 'เวลาออกมาตรฐาน')
        : base.default_dep_time || '',
    default_arr_time:
      input.default_arr_time !== undefined
        ? normalizeTime(input.default_arr_time, 'เวลาถึงมาตรฐาน')
        : base.default_arr_time || '',
    unit_price:
      input.unit_price !== undefined ? normalizeMoney(input.unit_price, 'ราคาต่อเที่ยว') : base.unit_price || 0,
    display_mode_zone:
      input.display_mode_zone !== undefined
        ? normalizeDisplayMode(input.display_mode_zone)
        : normalizeDisplayMode(base.display_mode_zone),
    active: input.active !== undefined ? Boolean(input.active) : base.active === undefined ? true : base.active
  };
}

async function createRoute(input) {
  const items = await listRoutes();
  const record = buildRouteRecord(input);
  if (items.some((r) => r.origin === record.origin && r.destination === record.destination)) {
    const label = record.destination ? `${record.origin} → ${record.destination}` : record.origin;
    throw httpError(409, `เส้นทาง "${label}" มีอยู่ในระบบแล้ว`);
  }
  const created = { id: items.reduce((m, r) => Math.max(m, r.id), 0) + 1, ...record };
  items.push(created);
  await setKV(KV_ROUTES, items);
  return created;
}

async function updateRoute(id, input) {
  const items = await listRoutes();
  const idx = items.findIndex((r) => r.id === Number(id));
  if (idx === -1) throw httpError(404, 'ไม่พบเส้นทางที่ต้องการแก้ไข');
  const record = buildRouteRecord(input, items[idx]);
  if (
    items.some(
      (r, i) => i !== idx && r.origin === record.origin && r.destination === record.destination
    )
  ) {
    const label = record.destination ? `${record.origin} → ${record.destination}` : record.origin;
    throw httpError(409, `เส้นทาง "${label}" มีอยู่ในระบบแล้ว`);
  }
  items[idx] = { id: items[idx].id, ...record };
  await setKV(KV_ROUTES, items);
  return items[idx];
}

async function deleteRoute(id) {
  const items = await listRoutes();
  const idx = items.findIndex((r) => r.id === Number(id));
  if (idx === -1) throw httpError(404, 'ไม่พบเส้นทางที่ต้องการลบ');
  const [deleted] = items.splice(idx, 1);
  await setKV(KV_ROUTES, items);
  return deleted;
}

// ============================================================================
// Travel logs
// ============================================================================

const LOG_FIELDS = [
  'trip_date',
  'trip_type',
  'worker_name',
  'license_plate',
  'site_name',
  'route_sheet_name',
  'billing_title',
  'origin',
  'dep_time',
  'destination',
  'arr_time',
  'passengers',
  'unit_price',
  'display_mode_zone',
  'remark'
];

function logToRow(id, log) {
  const body = { v: LOG_VERSION };
  for (const f of LOG_FIELDS) body[f] = log[f];
  return {
    worker_name: id,
    travel_date: log.trip_date,
    route: log.route_sheet_name,
    departure_time: log.dep_time,
    arrival_time: log.arr_time,
    vehicle_plate: log.license_plate,
    gps_status: JSON.stringify(body),
    updated_at: new Date().toISOString()
  };
}

function logFromRow(r) {
  let body = null;
  try {
    body = JSON.parse(r.gps_status);
  } catch {
    return null;
  }
  if (!body || body.v !== LOG_VERSION) return null;
  const out = { id: r.worker_name };
  for (const f of LOG_FIELDS) out[f] = body[f];
  out.passengers = body.passengers === null || body.passengers === undefined ? null : Number(body.passengers);
  out.unit_price = Number(body.unit_price) || 0;
  return out;
}

async function loadAllLogs() {
  return (await selectAll(T.logs, 'worker_name')).map(logFromRow).filter(Boolean);
}

function timeSortKey(row) {
  return String((row && (row.dep_time || row.arr_time)) || '99:99');
}

function compareLogsChrono(a, b) {
  if (a.trip_date !== b.trip_date) return String(a.trip_date).localeCompare(String(b.trip_date));
  const tc = timeSortKey(a).localeCompare(timeSortKey(b));
  if (tc !== 0) return tc;
  return String(a.id || '').localeCompare(String(b.id || ''));
}

function normalizeLogInput(input, base = {}) {
  const pick = (field) => (input[field] !== undefined ? input[field] : base[field]);

  const trip_date = str(pick('trip_date'));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(trip_date)) throw httpError(400, 'กรุณาระบุวันที่เดินทางให้ถูกต้อง');

  const trip_type = str(pick('trip_type')) || 'normal';
  if (!['normal', 'charter'].includes(trip_type)) {
    throw httpError(400, 'ประเภทงานต้องเป็น normal (รับ-ส่งปกติ) หรือ charter (งานเหมา/ขนย้าย)');
  }

  const worker_name = str(pick('worker_name'));
  if (!worker_name) throw httpError(400, 'กรุณาเลือกชื่อคนขับ');

  const origin = str(pick('origin'));
  const destination = str(pick('destination'));
  if (!origin || !destination) throw httpError(400, 'กรุณาระบุต้นทาง (จาก) และปลายทาง (ถึง)');

  const route_sheet_name = str(pick('route_sheet_name')) || `${origin} - ${destination}`;
  const billing_title = str(pick('billing_title')) || route_sheet_name;

  return {
    trip_date,
    trip_type,
    worker_name,
    license_plate: str(pick('license_plate')),
    site_name: str(pick('site_name')),
    route_sheet_name,
    billing_title,
    origin,
    dep_time: normalizeTime(pick('dep_time'), 'เวลาออก'),
    destination,
    arr_time: normalizeTime(pick('arr_time'), 'เวลาถึง'),
    passengers: normalizeIntOrNull(pick('passengers'), 'จำนวนคนโดยสาร'),
    unit_price: normalizeMoney(pick('unit_price'), 'ราคาต่อเที่ยว'),
    display_mode_zone: normalizeDisplayMode(pick('display_mode_zone')),
    remark: str(pick('remark'))
  };
}

async function assertQuotaAllows(extra = 1) {
  const quota = await getQuotaStatus();
  if (quota.currentCount + extra > quota.maxLimit) {
    throw httpError(
      507,
      `พื้นที่เก็บข้อมูลไม่พอ (${quota.currentCount}/${quota.maxLimit} รายการ) กรุณาส่งออกไฟล์สำรองแล้วล้างข้อมูลเก่า หรือเพิ่มขีดจำกัดในระบบหลังบ้าน`,
      { code: 'STORAGE_FULL', storageQuota: quota }
    );
  }
  return quota;
}

async function listLogs(filters = {}) {
  const { date_from, date_to, site_name, trip_type, search } = filters;
  let rows = await loadAllLogs();
  if (date_from) rows = rows.filter((r) => r.trip_date >= date_from);
  if (date_to) rows = rows.filter((r) => r.trip_date <= date_to);
  if (site_name) rows = rows.filter((r) => r.site_name === site_name);
  if (trip_type) rows = rows.filter((r) => r.trip_type === trip_type);
  if (search) {
    const q = String(search).toLowerCase();
    rows = rows.filter((r) =>
      [r.worker_name, r.license_plate, r.site_name, r.route_sheet_name, r.origin, r.destination, r.remark]
        .join(' ')
        .toLowerCase()
        .includes(q)
    );
  }
  // ใหม่สุดขึ้นก่อน (วันที่มาก → น้อย) และเวลาออกเรียงจากเช้า → เย็น (รายการที่ยังไม่ใส่เวลาอยู่ด้านล่างของวันนั้น)
  rows.sort((a, b) => {
    if (a.trip_date !== b.trip_date) return String(b.trip_date).localeCompare(String(a.trip_date));
    const tc = timeSortKey(a).localeCompare(timeSortKey(b));
    if (tc !== 0) return tc;
    return String(a.id || '').localeCompare(String(b.id || ''));
  });
  return rows;
}

async function createLog(input) {
  await assertQuotaAllows(1);
  const log = normalizeLogInput(input);
  const id = makeId('T');
  const row = await insertRow(T.logs, logToRow(id, log), 'worker_name');
  return { ...logFromRow(row), storageQuota: await getQuotaStatus() };
}

async function updateLog(id, updates) {
  const existingRow = await selectOne(T.logs, 'worker_name', id);
  const existing = existingRow ? logFromRow(existingRow) : null;
  if (!existing) throw httpError(404, 'ไม่พบข้อมูลการเดินทางที่ต้องการแก้ไข');
  const merged = normalizeLogInput(updates, existing);
  const { worker_name: _key, ...patch } = logToRow(id, merged);
  const row = await updateRow(T.logs, 'worker_name', id, patch);
  return logFromRow(row);
}

async function deleteLog(id) {
  const row = await selectOne(T.logs, 'worker_name', id);
  const existing = row ? logFromRow(row) : null;
  if (!existing) throw httpError(404, 'ไม่พบข้อมูลการเดินทางที่ต้องการลบ');
  await deleteRow(T.logs, 'worker_name', id);
  return { deleted: existing, storageQuota: await getQuotaStatus() };
}

async function cascadeLogField(field, oldValue, newValue) {
  const rows = await selectAll(T.logs, 'worker_name');
  let count = 0;
  for (const r of rows) {
    const log = logFromRow(r);
    if (!log || log[field] !== oldValue) continue;
    const updated = { ...log, [field]: newValue };
    const { worker_name: _k, ...patch } = logToRow(r.worker_name, updated);
    await updateRow(T.logs, 'worker_name', r.worker_name, patch);
    count += 1;
  }
  return count;
}

// คัดลอกคิววิ่งจากวันก่อนหน้า → วันที่เลือก
async function copyPreviousDay({ to_date, from_date }) {
  const target = str(to_date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(target)) throw httpError(400, 'กรุณาระบุวันที่ปลายทางให้ถูกต้อง');
  const all = await loadAllLogs();

  let source = str(from_date);
  if (!source) {
    const earlier = all.map((l) => l.trip_date).filter((d) => d < target);
    if (earlier.length === 0) throw httpError(404, 'ไม่พบรายการวันก่อนหน้าให้คัดลอก');
    source = earlier.sort().pop();
  }
  const sourceLogs = all.filter((l) => l.trip_date === source).sort(compareLogsChrono);
  if (sourceLogs.length === 0) throw httpError(404, `ไม่พบรายการของวันที่ ${source} ให้คัดลอก`);

  await assertQuotaAllows(sourceLogs.length);

  const rows = sourceLogs.map((l, i) => {
    const id = makeId('T') + String(i).padStart(3, '0');
    return logToRow(id, { ...l, trip_date: target });
  });
  await insertRows(T.logs, rows);
  return {
    created: rows.length,
    from_date: source,
    to_date: target,
    storageQuota: await getQuotaStatus()
  };
}

async function clearAllLogs() {
  const before = await countRows(T.logs);
  await deleteAllRows(T.logs, 'worker_name');
  return { deletedCount: before, storageQuota: await getQuotaStatus() };
}

// ============================================================================
// Google Sheets Sync Config — เก็บใน system_settings (ใช้ร่วมกันทุกเครื่อง)
// ============================================================================
const GSHEET_DEFAULT = {
  zone_url: '',
  zone_name: '',
  zone_sheet_url: '',
  billing_url: '',
  billing_name: '',
  billing_sheet_url: '',
  billing_site: '',
  auto_sync: true
};

function validateAppsScriptUrl(url, label) {
  const u = str(url);
  if (!u) return '';
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec(\?.*)?$/.test(u)) {
    throw httpError(
      400,
      `ลิงก์ ${label} ไม่ถูกต้อง — ต้องเป็น Web app URL ของ Google Apps Script (ขึ้นต้นด้วย https://script.google.com/macros/s/.../exec)`
    );
  }
  return u;
}

async function getGoogleSheetsConfig() {
  const saved = await getKV(KV_GSHEET, {});
  return {
    ...GSHEET_DEFAULT,
    ...(saved && typeof saved === 'object' ? saved : {}),
    auto_sync: saved && typeof saved.auto_sync === 'boolean' ? saved.auto_sync : true
  };
}

async function saveGoogleSheetsConfig(patch = {}) {
  const cur = await getGoogleSheetsConfig();
  const next = { ...cur };
  if (patch.zone_url !== undefined) next.zone_url = validateAppsScriptUrl(patch.zone_url, 'Google Sheet สำเนาโซน');
  if (patch.zone_name !== undefined) next.zone_name = str(patch.zone_name);
  if (patch.zone_sheet_url !== undefined) next.zone_sheet_url = str(patch.zone_sheet_url);
  if (patch.billing_url !== undefined) next.billing_url = validateAppsScriptUrl(patch.billing_url, 'Google Sheet ใบเบิกโครงการ');
  if (patch.billing_name !== undefined) next.billing_name = str(patch.billing_name);
  if (patch.billing_sheet_url !== undefined) next.billing_sheet_url = str(patch.billing_sheet_url);
  if (patch.billing_site !== undefined) next.billing_site = str(patch.billing_site);
  if (patch.auto_sync !== undefined) next.auto_sync = Boolean(patch.auto_sync);
  if (!next.zone_url) {
    next.zone_name = '';
    next.zone_sheet_url = '';
  }
  if (!next.billing_url) {
    next.billing_name = '';
    next.billing_sheet_url = '';
  }
  await setKV(KV_GSHEET, next);
  return next;
}

// ============================================================================
// Google Sheets Sync History — บันทึกประวัติการไหลข้อมูลลง Google Sheets
// ============================================================================
async function listGoogleSheetsHistory() {
  const rows = await getKV(KV_GSHEET_HISTORY, []);
  return Array.isArray(rows) ? rows : [];
}

async function addGoogleSheetsHistory(input = {}) {
  const allowedKinds = ['zone', 'billing', 'general'];
  const allowedModes = ['manual', 'auto', 'config_saved', 'note'];
  const allowedStatuses = ['success', 'error', 'info'];

  const kind = allowedKinds.includes(input.kind) ? input.kind : 'general';
  const mode = allowedModes.includes(input.mode) ? input.mode : 'manual';
  const status = allowedStatuses.includes(input.status) ? input.status : 'success';

  const entry = {
    id: makeId('GH'),
    synced_at: str(input.synced_at) || new Date().toISOString(),
    kind,
    mode,
    status,
    spreadsheet_name: str(input.spreadsheet_name),
    spreadsheet_url: str(input.spreadsheet_url),
    site_name: str(input.site_name),
    date_from: str(input.date_from),
    date_to: str(input.date_to),
    sheet_count: Number.isFinite(Number(input.sheet_count)) ? Math.max(0, Number(input.sheet_count)) : 0,
    log_count: Number.isFinite(Number(input.log_count)) ? Math.max(0, Number(input.log_count)) : 0,
    triggered_by: str(input.triggered_by) || 'ผู้ใช้ระบบ',
    note: str(input.note).slice(0, 500),
    error_message: str(input.error_message).slice(0, 500)
  };

  const cur = await listGoogleSheetsHistory();
  const next = [entry, ...cur].slice(0, MAX_GSHEET_HISTORY);
  await setKV(KV_GSHEET_HISTORY, next);
  return { entry, history: next };
}

async function deleteGoogleSheetsHistoryItem(id) {
  const key = str(id);
  const cur = await listGoogleSheetsHistory();
  const next = cur.filter((item) => item && item.id !== key);
  if (next.length === cur.length) throw httpError(404, 'ไม่พบบันทึกประวัติที่ต้องการลบ');
  await setKV(KV_GSHEET_HISTORY, next);
  return { history: next };
}

async function clearGoogleSheetsHistory() {
  const cur = await listGoogleSheetsHistory();
  await setKV(KV_GSHEET_HISTORY, []);
  return { deletedCount: cur.length };
}

module.exports = {
  verifyPassword,
  findUserByUsername,
  findUserById,
  getSettings,
  saveSettings,
  getQuotaStatus,
  listWorkers,
  createWorker,
  updateWorker,
  deleteWorker,
  listVehicles,
  createVehicle,
  updateVehicle,
  deleteVehicle,
  listSites,
  createSite,
  updateSite,
  deleteSite,
  listRoutes,
  createRoute,
  updateRoute,
  deleteRoute,
  listLogs,
  createLog,
  updateLog,
  deleteLog,
  copyPreviousDay,
  clearAllLogs,
  getGoogleSheetsConfig,
  saveGoogleSheetsConfig,
  validateAppsScriptUrl,
  listGoogleSheetsHistory,
  addGoogleSheetsHistory,
  deleteGoogleSheetsHistoryItem,
  clearGoogleSheetsHistory
};
