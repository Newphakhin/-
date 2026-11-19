require('dotenv').config();
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');
const { DatabaseSync } = require('node:sqlite');
const { createClient } = require('@supabase/supabase-js');

/**
 * ชั้นฐานข้อมูล (Database layer)
 * - ใช้ Supabase PostgreSQL เมื่อมีค่า SUPABASE_URL + SUPABASE_SERVICE_ROLE_KEY ใน .env
 * - ถ้าไม่มี จะใช้ SQLite ในเครื่อง (database.sqlite) เป็นตัวสำรอง
 * - ไม่มีข้อมูลตัวอย่าง (No Seed Data): ข้อมูลหลักและบันทึกการเดินทางเริ่มต้นว่างเปล่าทั้งหมด
 *   สิ่งเดียวที่สร้างให้อัตโนมัติคือบัญชีผู้ใช้เข้าระบบ (admin / user) เพื่อให้เข้าใช้งานได้
 */

const isServerless = Boolean(process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME);
const WRITABLE_DIR = isServerless ? os.tmpdir() : __dirname;
const DB_PATH = path.join(WRITABLE_DIR, 'database.sqlite');

const SUPABASE_URL = (process.env.SUPABASE_URL || '').trim();
const SUPABASE_SERVICE_ROLE_KEY = (process.env.SUPABASE_SERVICE_ROLE_KEY || '').trim();
const isSupabaseConfigured =
  Boolean(SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY) && SUPABASE_URL.startsWith('http');

const supabase = isSupabaseConfigured
  ? createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, { auth: { persistSession: false } })
  : null;

const db = new DatabaseSync(DB_PATH);
db.exec('PRAGMA journal_mode = WAL;');

function hashPassword(plain) {
  const hash = crypto.createHash('sha256').update(String(plain)).digest('hex');
  return `sha256:${hash}`;
}

function verifyPassword(plain, stored) {
  if (!stored) return false;
  if (stored.startsWith('sha256:')) return hashPassword(plain) === stored;
  return String(plain) === String(stored);
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

    CREATE TABLE IF NOT EXISTS master_vehicles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plate TEXT UNIQUE NOT NULL,
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
      updated_at TEXT
    );

    CREATE TABLE IF NOT EXISTS system_settings (
      key TEXT PRIMARY KEY,
      value TEXT
    );
  `);
}

function seedDatabase() {
  createSchema();
  // เฉพาะบัญชีผู้ใช้งานระบบ (รหัสผ่านเริ่มต้น 1234) — ไม่มีข้อมูลคนขับ/รถ/โครงการ/เส้นทาง/เที่ยววิ่ง
  const insertUser = db.prepare(
    'INSERT INTO users (username, password, full_name, role) VALUES (?, ?, ?, ?)'
  );
  if (!db.prepare("SELECT id FROM users WHERE username = 'admin'").get()) {
    insertUser.run('admin', hashPassword('1234'), 'แอดมิน', 'admin');
  }
  if (!db.prepare("SELECT id FROM users WHERE username = 'user'").get()) {
    insertUser.run('user', hashPassword('1234'), 'ผู้บันทึกข้อมูล', 'operator');
  }
}

module.exports = {
  db,
  supabase,
  isSupabaseConfigured,
  DB_PATH,
  hashPassword,
  verifyPassword,
  seedDatabase
};
