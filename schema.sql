-- ============================================================================
-- ระบบจัดการบันทึกข้อมูลการเดินทางของคนงาน (Worker Dispatch & Travel Log System)
-- Supabase / PostgreSQL Database Schema & Initial Seed Data (6 Tables)
-- วิธีใช้งาน: คัดลอกโค้ดทั้งหมดนี้ไปรันในเมนู SQL Editor บน Supabase Dashboard
-- ============================================================================

-- 1. ตารางผู้ใช้งานระบบ (users)
CREATE TABLE IF NOT EXISTS users (
  id SERIAL PRIMARY KEY,
  username TEXT UNIQUE NOT NULL,
  password TEXT NOT NULL,
  full_name TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('admin', 'operator'))
);

-- 2. ตารางรายชื่อคนงานสำหรับ Dropdown (master_workers)
CREATE TABLE IF NOT EXISTS master_workers (
  id SERIAL PRIMARY KEY,
  name TEXT UNIQUE NOT NULL,
  default_route TEXT DEFAULT '',
  active BOOLEAN DEFAULT TRUE
);

-- 3. ตารางเส้นทาง / ที่หมายสำหรับ Dropdown (master_routes)
CREATE TABLE IF NOT EXISTS master_routes (
  id SERIAL PRIMARY KEY,
  name TEXT UNIQUE NOT NULL,
  active BOOLEAN DEFAULT TRUE
);

-- 4. ตารางทะเบียนรถสำหรับ Dropdown (master_vehicles)
CREATE TABLE IF NOT EXISTS master_vehicles (
  id SERIAL PRIMARY KEY,
  plate TEXT UNIQUE NOT NULL,
  active BOOLEAN DEFAULT TRUE
);

-- 5. ตารางบันทึกข้อมูลการเดินทาง (travel_logs)
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

-- 6. ตารางตั้งค่าการเชื่อมต่อ API และโควต้าข้อมูล (system_settings)
CREATE TABLE IF NOT EXISTS system_settings (
  key TEXT PRIMARY KEY,
  value TEXT
);

-- ============================================================================
-- ข้อมูลเริ่มต้น (SEED DATA)
-- ============================================================================

-- Seed 1: users (รหัสผ่านเริ่มต้น: 1234)
INSERT INTO users (username, password, full_name, role)
VALUES
  ('admin', '1234', 'แอดมิน', 'admin'),
  ('user', '1234', 'ผู้บันทึกข้อมูล', 'operator')
ON CONFLICT (username) DO NOTHING;

-- Seed 2: master_vehicles (ทะเบียนรถครบทั้ง 20 คัน)
INSERT INTO master_vehicles (plate, active)
VALUES
  ('34-5798', TRUE),
  ('34-4546', TRUE),
  ('34-3771', TRUE),
  ('34-5793', TRUE),
  ('34-5579', TRUE),
  ('34-5795', TRUE),
  ('34-4596', TRUE),
  ('34-6756', TRUE),
  ('86-9334', TRUE),
  ('34-5273', TRUE),
  ('34-6703', TRUE),
  ('34-5797', TRUE),
  ('34-4373', TRUE),
  ('34-3062', TRUE),
  ('34-4597', TRUE),
  ('34-7124', TRUE),
  ('34-4370', TRUE),
  ('10-2784', TRUE),
  ('34-6716', TRUE),
  ('1นข 1609', TRUE)
ON CONFLICT (plate) DO NOTHING;

-- Seed 3: system_settings
INSERT INTO system_settings (key, value)
VALUES
  ('cartrack_api_url', 'https://fleetapi-th.cartrack.com/rest/vehicles/status'),
  ('cartrack_username', 'DEMO_FLEET_TH'),
  ('cartrack_api_key', ''),
  ('cartrack_mock_mode', 'true'),
  ('max_storage_limit', '100')
ON CONFLICT (key) DO NOTHING;
