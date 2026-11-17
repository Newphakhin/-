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
  ('admin', '1234', 'สมชาย บริหารงาน (ผู้ดูแลระบบ Admin)', 'admin'),
  ('user', '1234', 'วิภาวดี จัดคิวรถ (เจ้าหน้าที่บันทึกข้อมูล Operator)', 'operator')
ON CONFLICT (username) DO NOTHING;

-- Seed 2: master_routes
INSERT INTO master_routes (name, active)
VALUES
  ('นิคมอุตสาหกรรมอมตะซิตี้ ชลบุรี', TRUE),
  ('ไซต์งานก่อสร้างรถไฟฟ้า สายสีม่วง', TRUE),
  ('คลังสินค้าบางนา กม.18', TRUE),
  ('โรงงานประกอบรถยนต์ ระยอง', TRUE),
  ('ท่าเรือแหลมฉบัง ชลบุรี', TRUE),
  ('นิคมอุตสาหกรรมโรจนะ อยุธยา', TRUE),
  ('ศูนย์กระจายสินค้า วังน้อย (ปิดปรับปรุง)', FALSE)
ON CONFLICT (name) DO NOTHING;

-- Seed 3: master_workers
INSERT INTO master_workers (name, default_route, active)
VALUES
  ('สมศักดิ์ ใจดี', 'นิคมอุตสาหกรรมอมตะซิตี้ ชลบุรี', TRUE),
  ('ประยุทธ์ ขยันงาน', 'ไซต์งานก่อสร้างรถไฟฟ้า สายสีม่วง', TRUE),
  ('มานะ อดทน', 'คลังสินค้าบางนา กม.18', TRUE),
  ('กิตติพงศ์ รักงาน', 'โรงงานประกอบรถยนต์ ระยอง', TRUE),
  ('วีระชัย สายลุย', 'ท่าเรือแหลมฉบัง ชลบุรี', TRUE),
  ('ธนากร พร้อมลุย', '', TRUE),
  ('อนุชา มั่นคง', 'นิคมอุตสาหกรรมโรจนะ อยุธยา', TRUE),
  ('สุรพล พักงาน (ลาพักร้อน)', 'คลังสินค้าบางนา กม.18', FALSE)
ON CONFLICT (name) DO NOTHING;

-- Seed 4: master_vehicles (ทะเบียนรถครบทั้ง 20 คัน)
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

-- Seed 5: travel_logs
INSERT INTO travel_logs (worker_name, travel_date, route, departure_time, arrival_time, vehicle_plate, gps_status)
VALUES
  ('สมศักดิ์ ใจดี', TO_CHAR(CURRENT_DATE, 'YYYY-MM-DD'), 'นิคมอุตสาหกรรมอมตะซิตี้ ชลบุรี', '06:30', '08:00', '34-5798', '{"status":"กำลังวิ่ง 68 กม./ชม.","engine":"ON","speed":68,"lat":13.4258,"lng":101.0021,"location":"ถ.เทพรัตน (บางนา-ตราด) กม.57 ชลบุรี","maps_url":"https://www.google.com/maps?q=13.4258,101.0021","updated_at":"วันนี้"}'),
  ('ประยุทธ์ ขยันงาน', TO_CHAR(CURRENT_DATE, 'YYYY-MM-DD'), 'ไซต์งานก่อสร้างรถไฟฟ้า สายสีม่วง', '07:00', '08:15', '34-4546', '{"status":"จอดอยู่ (ติดเครื่อง)","engine":"ON","speed":0,"lat":13.6892,"lng":100.4912,"location":"ถ.สมเด็จพระเจ้าตากสิน แขวงดาวคะนอง กทม.","maps_url":"https://www.google.com/maps?q=13.6892,100.4912","updated_at":"วันนี้"}'),
  ('มานะ อดทน', TO_CHAR(CURRENT_DATE, 'YYYY-MM-DD'), 'คลังสินค้าบางนา กม.18', '22:30', '01:00', '34-3771', '{"status":"กำลังวิ่ง 54 กม./ชม.","engine":"ON","speed":54,"lat":13.6195,"lng":100.7428,"location":"ถ.บางนา-ตราด กม.18 บางพลี สมุทรปราการ","maps_url":"https://www.google.com/maps?q=13.6195,100.7428","updated_at":"วันนี้"}'),
  ('กิตติพงศ์ รักงาน', TO_CHAR(CURRENT_DATE, 'YYYY-MM-DD'), 'โรงงานประกอบรถยนต์ ระยอง', '05:45', '08:30', '34-5793', '{"status":"จอดอยู่ (ดับเครื่อง)","engine":"OFF","speed":0,"lat":13.0145,"lng":101.1422,"location":"นิคมฯ อีสเทิร์นซีบอร์ด ปลวกแดง ระยอง","maps_url":"https://www.google.com/maps?q=13.0145,101.1422","updated_at":"วันนี้"}')
ON CONFLICT (worker_name) DO NOTHING;

-- Seed 6: system_settings
INSERT INTO system_settings (key, value)
VALUES
  ('cartrack_api_url', 'https://fleetapi-th.cartrack.com/rest/vehicles/status'),
  ('cartrack_username', 'DEMO_FLEET_TH'),
  ('cartrack_api_key', ''),
  ('cartrack_mock_mode', 'true'),
  ('max_storage_limit', '100')
ON CONFLICT (key) DO NOTHING;
