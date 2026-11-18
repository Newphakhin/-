require('dotenv').config();
const path = require('node:path');
const fs = require('node:fs');
const crypto = require('node:crypto');
const { DatabaseSync } = require('node:sqlite');
const { createClient } = require('@supabase/supabase-js');

const os = require('node:os');

const isServerless = Boolean(process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME);
const WRITABLE_DIR = isServerless ? os.tmpdir() : __dirname;

const DB_PATH = path.join(WRITABLE_DIR, 'database.sqlite');
const JSON_FALLBACK_PATH = path.join(WRITABLE_DIR, 'fallback_storage.json');

// Check if Supabase credentials are provided in .env
const SUPABASE_URL = (process.env.SUPABASE_URL || '').trim();
const SUPABASE_SERVICE_ROLE_KEY = (process.env.SUPABASE_SERVICE_ROLE_KEY || '').trim();
const isSupabaseConfigured =
  Boolean(SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY) &&
  SUPABASE_URL.startsWith('http');

const supabase = isSupabaseConfigured
  ? createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
      auth: { persistSession: false }
    })
  : null;

// Local SQLite database (acts as fast local engine & automatic fallback when SUPABASE_URL is not set)
const db = new DatabaseSync(DB_PATH);
db.exec('PRAGMA journal_mode = WAL;');
db.exec('PRAGMA foreign_keys = ON;');

function hashPassword(plain) {
  const hash = crypto.createHash('sha256').update(String(plain)).digest('hex');
  return `sha256:${hash}`;
}

function verifyPassword(plain, stored) {
  if (!stored) return false;
  if (stored.startsWith('sha256:')) {
    return hashPassword(plain) === stored;
  }
  return String(plain) === String(stored);
}

function getTodayDateString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function createSchema() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password TEXT NOT NULL,
      full_name TEXT NOT NULL,
      role TEXT NOT NULL CHECK (role IN ('admin', 'operator'))
    );

    CREATE TABLE IF NOT EXISTS master_workers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      default_route TEXT DEFAULT '',
      active BOOLEAN DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS master_routes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      active BOOLEAN DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS master_vehicles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plate TEXT UNIQUE NOT NULL,
      active BOOLEAN DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS master_sites (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE NOT NULL,
      active BOOLEAN DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS travel_logs (
      worker_name TEXT PRIMARY KEY,
      site_name TEXT DEFAULT '',
      travel_date TEXT,
      route TEXT,
      departure_time TEXT,
      arrival_time TEXT,
      vehicle_plate TEXT,
      passenger_count INTEGER DEFAULT NULL,
      gps_status TEXT DEFAULT '',
      updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS system_settings (
      key TEXT PRIMARY KEY,
      value TEXT
    );
  `);

  // Ensure columns exist if upgrading from an older travel_logs table
  const columns = db.prepare('PRAGMA table_info(travel_logs)').all();
  if (!columns.some((c) => c.name === 'gps_status')) {
    db.exec("ALTER TABLE travel_logs ADD COLUMN gps_status TEXT DEFAULT '';");
  }
  if (!columns.some((c) => c.name === 'passenger_count')) {
    db.exec('ALTER TABLE travel_logs ADD COLUMN passenger_count INTEGER DEFAULT NULL;');
  }
  if (!columns.some((c) => c.name === 'site_name')) {
    db.exec("ALTER TABLE travel_logs ADD COLUMN site_name TEXT DEFAULT '';");
  }
}

const INITIAL_VEHICLES = [
  '34-5798',
  '34-4546',
  '34-3771',
  '34-5793',
  '34-5579',
  '34-5795',
  '34-4596',
  '34-6756',
  '86-9334',
  '34-5273',
  '34-6703',
  '34-5797',
  '34-4373',
  '34-3062',
  '34-4597',
  '34-7124',
  '34-4370',
  '10-2784',
  '34-6716',
  '1นข 1609'
];

function buildSampleGpsStatus(statusText, engine, speed, lat, lng, location) {
  return JSON.stringify({
    status: statusText,
    engine,
    speed,
    lat,
    lng,
    location,
    maps_url: `https://www.google.com/maps?q=${lat},${lng}`,
    updated_at: new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }) + ' น.'
  });
}

function seedDatabase(forceReset = false) {
  createSchema();

  if (forceReset) {
    db.exec(`
      DELETE FROM travel_logs;
      DELETE FROM master_vehicles;
      DELETE FROM master_routes;
      DELETE FROM master_workers;
      DELETE FROM users;
      DELETE FROM system_settings;
      DELETE FROM sqlite_sequence WHERE name IN ('users', 'master_workers', 'master_routes', 'master_vehicles');
    `);
  }

  // 1. Seed users (Both admin and operator roles, password: 1234)
  const upsertUser = db.prepare(
    'INSERT OR REPLACE INTO users (id, username, password, full_name, role) VALUES (?, ?, ?, ?, ?)'
  );
  const existingAdmin = db.prepare("SELECT id FROM users WHERE username = 'admin'").get();
  if (!existingAdmin) {
    upsertUser.run(1, 'admin', hashPassword('1234'), 'แอดมิน', 'admin');
  } else {
    db.prepare("UPDATE users SET full_name = 'แอดมิน' WHERE username = 'admin'").run();
  }
  const existingOperator = db.prepare("SELECT id FROM users WHERE username = 'user'").get();
  if (!existingOperator) {
    upsertUser.run(2, 'user', hashPassword('1234'), 'ผู้บันทึกข้อมูล', 'operator');
  }

  // 2. Seed master_vehicles (All 20 license plates)
  const vehicleCount = db.prepare('SELECT COUNT(*) AS count FROM master_vehicles').get().count;
  if (vehicleCount === 0) {
    const insertVehicle = db.prepare('INSERT INTO master_vehicles (plate, active) VALUES (?, ?)');
    for (const plate of INITIAL_VEHICLES) {
      insertVehicle.run(plate, 1);
    }
  }

  // 6. Seed system_settings
  const defaultSettings = {
    cartrack_api_url:
      process.env.CARTRACK_API_URL || 'https://fleetapi-th.cartrack.com/rest/vehicles/status',
    cartrack_username: process.env.CARTRACK_USERNAME || 'DEMO_FLEET_TH',
    cartrack_api_key: process.env.CARTRACK_API_KEY || '',
    cartrack_mock_mode: process.env.CARTRACK_MOCK_MODE || 'true',
    max_storage_limit: process.env.MAX_STORAGE_LIMIT || '100'
  };

  const upsertSetting = db.prepare(
    'INSERT OR IGNORE INTO system_settings (key, value) VALUES (?, ?)'
  );
  for (const [k, v] of Object.entries(defaultSettings)) {
    upsertSetting.run(k, String(v));
  }

  syncJsonFallbackFile();
}

function syncJsonFallbackFile() {
  try {
    const workers = db.prepare('SELECT * FROM master_workers').all();
    const routes = db.prepare('SELECT * FROM master_routes').all();
    const vehicles = db.prepare('SELECT * FROM master_vehicles').all();
    const logs = db.prepare('SELECT * FROM travel_logs').all();
    const settings = db.prepare('SELECT * FROM system_settings').all();

    const snapshot = {
      updated_at: new Date().toISOString(),
      mode: isSupabaseConfigured ? 'supabase_cloud' : 'local_sqlite',
      master_workers: workers,
      master_routes: routes,
      master_vehicles: vehicles,
      travel_logs: logs,
      system_settings: settings
    };
    fs.writeFileSync(JSON_FALLBACK_PATH, JSON.stringify(snapshot, null, 2), 'utf8');

    // If Supabase Cloud is configured, sync tables to Supabase PostgreSQL asynchronously
    if (isSupabaseConfigured && supabase) {
      Promise.all([
        supabase.from('master_workers').upsert(
          workers.map((w) => ({
            id: w.id,
            name: w.name,
            default_route: w.default_route || '',
            active: Boolean(w.active)
          })),
          { onConflict: 'name' }
        ),
        supabase.from('master_routes').upsert(
          routes.map((r) => ({
            id: r.id,
            name: r.name,
            active: Boolean(r.active)
          })),
          { onConflict: 'name' }
        ),
        supabase.from('master_vehicles').upsert(
          vehicles.map((v) => ({
            id: v.id,
            plate: v.plate,
            active: Boolean(v.active)
          })),
          { onConflict: 'plate' }
        ),
        supabase.from('travel_logs').upsert(
          logs.map((l) => ({
            worker_name: l.worker_name,
            travel_date: l.travel_date || '',
            route: l.route || '',
            departure_time: l.departure_time || '',
            arrival_time: l.arrival_time || '',
            vehicle_plate: l.vehicle_plate || '',
            gps_status: l.gps_status || ''
          })),
          { onConflict: 'worker_name' }
        ),
        supabase.from('system_settings').upsert(settings, { onConflict: 'key' })
      ]).catch((err) => {
        console.error('Supabase background sync warning:', err.message);
      });
    }
  } catch {
    // ignore snapshot write errors
  }
}

function getSystemSettingsMap() {
  const rows = db.prepare('SELECT key, value FROM system_settings').all();
  const map = {
    cartrack_api_url: 'https://fleetapi-th.cartrack.com/rest/vehicles/status',
    cartrack_username: 'DEMO_FLEET_TH',
    cartrack_api_key: '',
    cartrack_mock_mode: 'true',
    max_storage_limit: '100'
  };
  for (const r of rows) {
    map[r.key] = r.value;
  }
  return map;
}

function updateSystemSettings(newSettings) {
  const upsert = db.prepare('INSERT OR REPLACE INTO system_settings (key, value) VALUES (?, ?)');
  db.exec('BEGIN');
  try {
    for (const [k, v] of Object.entries(newSettings)) {
      if (v !== undefined && v !== null) {
        upsert.run(k, String(v));
      }
    }
    db.exec('COMMIT');
  } catch (err) {
    db.exec('ROLLBACK');
    throw err;
  }
  syncJsonFallbackFile();
  return getSystemSettingsMap();
}

function getStorageQuotaStatus() {
  const settings = getSystemSettingsMap();
  const maxLimit = Math.max(1, Number(settings.max_storage_limit) || 100);
  const currentCount = db.prepare('SELECT COUNT(*) AS count FROM travel_logs').get().count;
  const usagePercent = Math.min(100, Math.round((currentCount / maxLimit) * 100));

  let dbSizeBytes = 0;
  try {
    const stat = fs.statSync(DB_PATH);
    dbSizeBytes = stat.size;
  } catch {
    dbSizeBytes = currentCount * 512;
  }

  const supabaseFreeTierBytes = 500 * 1024 * 1024; // 500 MB
  let statusLevel = 'normal';
  if (currentCount >= maxLimit || usagePercent >= 100) {
    statusLevel = 'full';
  } else if (usagePercent >= 80) {
    statusLevel = 'warning';
  }

  return {
    currentCount,
    maxLimit,
    usagePercent,
    statusLevel, // 'normal' | 'warning' | 'full'
    isFull: currentCount >= maxLimit,
    isWarning: usagePercent >= 80 && currentCount < maxLimit,
    dbSizeBytes,
    supabaseFreeTierBytes,
    storageMode: isSupabaseConfigured ? 'supabase_cloud' : 'local_fallback'
  };
}

// Realistic Thai GPS Locations for Cartrack Mock Mode
const THAI_GPS_LOCATIONS = [
  {
    lat: 13.4258,
    lng: 101.0021,
    location: 'ถ.เทพรัตน (บางนา-ตราด) กม.57 อ.เมือง ชลบุรี'
  },
  {
    lat: 13.6195,
    lng: 100.7428,
    location: 'ถ.บางนา-ตราด กม.18 อ.บางพลี สมุทรปราการ'
  },
  {
    lat: 13.6892,
    lng: 100.4912,
    location: 'ถ.สมเด็จพระเจ้าตากสิน แขวงดาวคะนอง กทม.'
  },
  {
    lat: 13.0145,
    lng: 101.1422,
    location: 'นิคมฯ อีสเทิร์นซีบอร์ด อ.ปลวกแดง ระยอง'
  },
  {
    lat: 13.0819,
    lng: 100.8961,
    location: 'ท่าเรือแหลมฉบัง ถ.สุขุมวิท อ.ศรีราชา ชลบุรี'
  },
  {
    lat: 14.3312,
    lng: 100.6918,
    location: 'นิคมอุตสาหกรรมโรจนะ อ.อุทัย พระนครศรีอยุธยา'
  },
  {
    lat: 13.8854,
    lng: 100.4092,
    location: 'ถ.กาญจนาภิเษก (วงแหวนตะวันตก) อ.บางบัวทอง นนทบุรี'
  },
  {
    lat: 14.1652,
    lng: 100.6184,
    location: 'ถ.พหลโยธิน กม.52 อ.บางปะอิน พระนครศรีอยุธยา'
  },
  {
    lat: 13.5621,
    lng: 100.2894,
    location: 'ถ.พระราม 2 กม.34 อ.เมือง สมุทรสาคร'
  },
  {
    lat: 13.3611,
    lng: 100.9847,
    location: 'ทางหลวงพิเศษมอเตอร์เวย์ สาย 7 อ.เมือง ชลบุรี'
  }
];

function generateMockCartrackGps(plate) {
  const cleanPlate = String(plate || '34-5798').trim();
  const randIdx = Math.floor(Math.random() * THAI_GPS_LOCATIONS.length);
  const baseSpot = THAI_GPS_LOCATIONS[randIdx];

  // Add slight realistic coordinate jitter
  const lat = Number((baseSpot.lat + (Math.random() - 0.5) * 0.012).toFixed(5));
  const lng = Number((baseSpot.lng + (Math.random() - 0.5) * 0.012).toFixed(5));

  const stateRoll = Math.random();
  let speed = 0;
  let engine = 'ON';
  let status = '';

  if (stateRoll < 0.68) {
    speed = Math.floor(42 + Math.random() * 46); // 42 - 87 km/h
    engine = 'ON';
    status = `กำลังวิ่ง ${speed} กม./ชม.`;
  } else if (stateRoll < 0.85) {
    speed = 0;
    engine = 'ON';
    status = 'จอดอยู่ (ติดเครื่อง)';
  } else {
    speed = 0;
    engine = 'OFF';
    status = 'จอดอยู่ (ดับเครื่อง)';
  }

  const nowTimeStr =
    new Date().toLocaleTimeString('th-TH', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    }) + ' น.';

  return {
    plate: cleanPlate,
    status,
    engine,
    speed,
    lat,
    lng,
    location: baseSpot.location,
    maps_url: `https://www.google.com/maps?q=${lat},${lng}`,
    updated_at: nowTimeStr,
    source: 'mock'
  };
}

module.exports = {
  db,
  supabase,
  isSupabaseConfigured,
  DB_PATH,
  JSON_FALLBACK_PATH,
  hashPassword,
  verifyPassword,
  seedDatabase,
  syncJsonFallbackFile,
  getSystemSettingsMap,
  updateSystemSettings,
  getStorageQuotaStatus,
  generateMockCartrackGps
};
