-- ============================================================================
-- ระบบจัดการบันทึกข้อมูลการเดินทางพนักงาน — บริษัท จักรพงษ์ ขนส่ง (2003) จำกัด
-- Supabase / PostgreSQL Database Schema (V3 — No Seed Data)
-- วิธีใช้งาน: คัดลอกโค้ดทั้งหมดนี้ไปรันในเมนู SQL Editor บน Supabase Dashboard
-- หมายเหตุ: ตารางข้อมูลหลักและบันทึกการเดินทางเริ่มต้นว่างเปล่าทั้งหมด (No Seed Data)
--          มีเฉพาะบัญชีผู้ใช้งานระบบ (admin / user) เพื่อให้เข้าสู่ระบบได้
-- ============================================================================

-- 1. ตารางผู้ใช้งานระบบ (users)
CREATE TABLE IF NOT EXISTS users (
  id SERIAL PRIMARY KEY,
  username TEXT UNIQUE NOT NULL,
  password TEXT NOT NULL,
  full_name TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('admin', 'operator'))
);

-- 2. ตารางรายชื่อคนงาน/คนขับ (master_workers)
--    default_route ใช้เก็บทะเบียนรถประจำตัวคนขับ (default_vehicle) เพื่อ Auto-fill
CREATE TABLE IF NOT EXISTS master_workers (
  id SERIAL PRIMARY KEY,
  name TEXT UNIQUE NOT NULL,
  default_route TEXT DEFAULT '',
  active BOOLEAN DEFAULT TRUE
);

-- 3. ตารางทะเบียนรถ (master_vehicles)
CREATE TABLE IF NOT EXISTS master_vehicles (
  id SERIAL PRIMARY KEY,
  plate TEXT UNIQUE NOT NULL,
  active BOOLEAN DEFAULT TRUE
);

-- 4. ตารางบันทึกการเดินทาง (travel_logs)
--    worker_name ใช้เก็บ id เฉพาะของแต่ละเที่ยววิ่ง (เช่น T...)
--    gps_status ใช้เก็บ JSON โครงสร้างเต็มรูปแบบของเที่ยววิ่ง V3
CREATE TABLE IF NOT EXISTS travel_logs (
  worker_name TEXT PRIMARY KEY,
  travel_date TEXT,
  route TEXT,
  departure_time TEXT,
  arrival_time TEXT,
  vehicle_plate TEXT,
  gps_status TEXT DEFAULT '',
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. ตารางตั้งค่าระบบและจัดเก็บข้อมูลหลัก V3 (system_settings)
--    key = 'v3_sites'          -> JSON รายการไซต์งาน/โครงการ (id, site_name, zone_name, active)
--    key = 'v3_routes'         -> JSON รายการเส้นทาง/ที่หมาย (id, origin, destination, route_sheet_name, billing_title, default_dep_time, default_arr_time, unit_price, display_mode_zone, active)
--    key = 'max_storage_limit' -> ขีดจำกัดจำนวนรายการสูงสุด
CREATE TABLE IF NOT EXISTS system_settings (
  key TEXT PRIMARY KEY,
  value TEXT
);

-- ============================================================================
-- บัญชีผู้ใช้งานระบบเริ่มต้น (รหัสผ่านเริ่มต้น: 1234)
-- ห้ามใส่ข้อมูลตัวอย่างคนขับ/ทะเบียนรถ/โครงการ/เส้นทาง/บันทึกการเดินทาง
-- ============================================================================
INSERT INTO users (username, password, full_name, role)
VALUES
  ('admin', '1234', 'แอดมิน', 'admin'),
  ('user', '1234', 'ผู้บันทึกข้อมูล', 'operator')
ON CONFLICT (username) DO NOTHING;

INSERT INTO system_settings (key, value)
VALUES
  ('max_storage_limit', '10000')
ON CONFLICT (key) DO NOTHING;
