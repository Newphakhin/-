# 📘 คู่มือการนำระบบขึ้นคลาวด์ออนไลน์ (Supabase + Render.com ฟรี)

เอกสารนี้อธิบายขั้นตอนการนำ **"ระบบจัดการบันทึกข้อมูลการเดินทางของคนงาน (Worker Dispatch & Travel Log System)"** ขึ้นใช้งานบนคลาวด์ออนไลน์ เพื่อให้เปิดใช้งานได้จากทั้ง **โทรศัพท์มือถือ (iOS / Android)** และ **คอมพิวเตอร์** ได้ทุกที่

---

## ขั้นตอนที่ 1: สร้างฐานข้อมูล PostgreSQL ฟรีบน Supabase

1. เข้าไปที่เว็บไซต์ [https://supabase.com](https://supabase.com) แล้วสมัครสมาชิกหรือเข้าสู่ระบบ
2. กดปุ่ม **"New Project"**
   - ตั้งชื่อโปรเจกต์ เช่น `worker-dispatch-db`
   - กำหนดรหัสผ่าน Database Password
   - เลือก Region เป็น **`Southeast Asia (Singapore)`** เพื่อความรวดเร็วในประเทศไทย
   - กด **"Create new project"** และรอประมาณ 1-2 นาที
3. เมื่อสร้างเสร็จแล้ว ให้ไปที่เมนูด้านซ้ายเลือก **"SQL Editor"** -> กด **"New query"**
4. เปิดไฟล์ **`schema.sql`** ในโปรเจกต์นี้ คัดลอกโค้ด SQL ทั้งหมดไปวางในช่อง SQL Editor แล้วกดปุ่ม **"Run"**
   - ระบบจะสร้างตารางทั้ง **6 ตารางหลัก** (`users`, `master_workers`, `master_routes`, `master_vehicles`, `travel_logs`, `system_settings`) พร้อมข้อมูลเริ่มต้นและทะเบียนรถทั้ง 20 คันทันที
5. ไปที่เมนู **"Project Settings"** (ไอคอนรูปเฟืองด้านล่างซ้าย) -> เลือกเมนู **"API"**
   - คัดลอกค่า **`Project URL`** (ใช้สำหรับ `SUPABASE_URL`)
   - คัดลอกค่า **`service_role` (`secret`) key** (ใช้สำหรับ `SUPABASE_SERVICE_ROLE_KEY`)

---

## ขั้นตอนที่ 2: การรันทดสอบในเครื่อง (Local Development)

1. คัดลอกไฟล์ `.env.example` เป็นไฟล์ชื่อ `.env`:
   ```bash
   cp .env.example .env
   ```
2. หากต้องการเชื่อมต่อ Supabase ทันที ให้ใส่ค่า `SUPABASE_URL` และ `SUPABASE_SERVICE_ROLE_KEY` ลงในไฟล์ `.env`
   *(หมายเหตุ: หากยังไม่ได้ใส่ค่า `SUPABASE_URL` ระบบจะมี **Local Fallback (SQLite + JSON)** ทำงานอัตโนมัติพร้อมแจ้งเตือนบนหน้าเว็บ ทำให้รันทดสอบทุกฟีเจอร์ในเครื่องได้ทันทีโดยไม่ Error)*
3. ติดตั้งแพ็กเกจและรันเซิร์ฟเวอร์:
   ```bash
   npm install
   npm start
   ```
4. เปิดเว็บเบราว์เซอร์ไปที่ `http://localhost:3000`

---

## ขั้นตอนที่ 3: นำเว็บแอปพลิเคชันขึ้นออนไลน์ฟรีบน Render.com

1. อัปโหลดโค้ดโปรเจกต์นี้ขึ้น **GitHub Repository** ของคุณ (ระวังอย่าอัปโหลดไฟล์ `.env` ที่เก็บความลับ)
2. เข้าไปที่เว็บไซต์ [https://render.com](https://render.com) และเข้าสู่ระบบด้วยบัญชี GitHub
3. กดปุ่ม **"New +"** -> เลือก **"Web Service"**
4. เชื่อมต่อกับ GitHub Repository ของโปรเจกต์นี้แล้วตั้งค่าดังนี้:
   - **Name**: `worker-dispatch-system`
   - **Region**: `Singapore (Southeast Asia)`
   - **Runtime**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
   - **Instance Type**: เลือก **`Free`**
5. เลื่อนลงมาที่หัวข้อ **"Environment Variables"** แล้วกด **"Add Environment Variable"** เพื่อใส่ค่าดังนี้:
   - `SUPABASE_URL` = *(ใส่ Project URL จาก Supabase)*
   - `SUPABASE_SERVICE_ROLE_KEY` = *(ใส่ service_role secret key จาก Supabase)*
   - `TOKEN_SECRET` = *(ตั้งรหัสลับใดๆ เช่น `my-dispatch-secret-2026`)*
6. กดปุ่ม **"Create Web Service"** รอประมาณ 2 นาที เมื่อขึ้นสถานะ **`Live`** คุณจะได้รับลิงก์ URL เช่น `https://worker-dispatch-system.onrender.com` ซึ่งสามารถส่งให้ทีมงานเปิดผ่านมือถือหรือคอมพิวเตอร์ได้ทันที!

---

## ขั้นตอนที่ 4: การตั้งค่าเชื่อมต่อ Cartrack GPS API

1. เข้าสู่ระบบด้วยสิทธิ์ **ผู้ดูแลระบบ (Admin)**: `admin` / รหัสผ่าน `1234`
2. กดที่เมนู **"ระบบหลังบ้าน & ตั้งค่า GPS"** -> ไปที่ส่วนที่ 4 **"ตั้งค่าการเชื่อมต่อ Cartrack GPS & โควต้าพื้นที่"**
3. โดยค่าเริ่มต้น ระบบจะเปิด **"โหมดจำลองข้อมูล (Mock GPS Mode)"** ไว้เพื่อให้คุณทดสอบกดปุ่มเช็กพิกัดรถ, ความเร็ว, สถานะเครื่องยนต์ และเปิดลิงก์ Google Maps ได้ทันที
4. เมื่อต้องการใช้งานกับระบบ Cartrack จริง:
   - กรอก `Cartrack API Endpoint URL` (เช่น `https://fleetapi-th.cartrack.com/rest/vehicles/status`)
   - กรอก `Cartrack Username` และ `Cartrack API Key`
   - ปิดสวิตช์ **Mock GPS Mode** เพื่อสลับเป็น **เชื่อมต่อจริง (Live API)**
   - กดปุ่ม **"ทดสอบการเชื่อมต่อ API"** และกด **"บันทึกการตั้งค่า"**
