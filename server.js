require('dotenv').config();
const express = require('express');
const path = require('node:path');
const crypto = require('node:crypto');
const {
  isSupabaseConfigured,
  verifyPassword,
  seedDatabase,
  generateMockCartrackGps
} = require('./db');
const store = require('./store');

const app = express();
const PORT = process.env.PORT || 3000;
const TOKEN_SECRET = process.env.TOKEN_SECRET || 'worker-dispatch-secret-key-2026';

// Initialize local SQLite schema/seed on startup (used as local fallback when Supabase is not configured)
seedDatabase(false);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Token signing & verification (HMAC-SHA256 stateless token + revocation set)
const revokedTokens = new Set();

function createToken(user) {
  const payload = {
    id: user.id,
    username: user.username,
    full_name: user.full_name,
    role: user.role,
    iat: Date.now()
  };
  const base64Payload = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const signature = crypto
    .createHmac('sha256', TOKEN_SECRET)
    .update(base64Payload)
    .digest('base64url');
  return `${base64Payload}.${signature}`;
}

function verifyToken(token) {
  if (!token || revokedTokens.has(token)) return null;
  const parts = token.split('.');
  if (parts.length !== 2) return null;
  const [base64Payload, signature] = parts;
  const expectedSig = crypto
    .createHmac('sha256', TOKEN_SECRET)
    .update(base64Payload)
    .digest('base64url');
  if (signature !== expectedSig) return null;
  try {
    return JSON.parse(Buffer.from(base64Payload, 'base64url').toString('utf8'));
  } catch {
    return null;
  }
}

function extractToken(req) {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    return authHeader.slice(7).trim();
  }
  if (req.headers['x-auth-token']) {
    return String(req.headers['x-auth-token']).trim();
  }
  return null;
}

function requireAuth(req, res, next) {
  const token = extractToken(req);
  const payload = verifyToken(token);
  if (!payload) {
    return res.status(401).json({ error: 'กรุณาเข้าสู่ระบบก่อนใช้งาน (Unauthorized)' });
  }
  req.user = payload;
  req.token = token;
  next();
}

function requireAdmin(req, res, next) {
  if (!req.user || req.user.role !== 'admin') {
    return res
      .status(403)
      .json({ error: 'เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่ได้รับสิทธิ์ใช้งานส่วนนี้' });
  }
  next();
}

// ============================================================================
// 1. Authentication, Session & Storage Quota APIs
// ============================================================================

app.post('/api/login', async (req, res) => {
  try {
    const username = String(req.body.username || '').trim();
    const password = String(req.body.password || '');

    if (!username || !password) {
      return res.status(400).json({ error: 'กรุณากรอกชื่อผู้ใช้งานและรหัสผ่าน' });
    }

    const user = await store.findUserByUsername(username);
    if (!user || !verifyPassword(password, user.password)) {
      return res.status(401).json({ error: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' });
    }

    const safeUser = {
      id: user.id,
      username: user.username,
      full_name: user.full_name,
      role: user.role
    };

    const token = createToken(safeUser);
    const storageQuota = await store.getQuotaStatus();
    return res.json({
      message: 'เข้าสู่ระบบสำเร็จ',
      token,
      user: safeUser,
      storageQuota
    });
  } catch (err) {
    console.error('Login error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเข้าสู่ระบบ' });
  }
});

app.get('/api/me', requireAuth, async (req, res) => {
  try {
    const user = await store.findUserById(req.user.id);
    if (!user) {
      return res.status(401).json({ error: 'ไม่พบข้อมูลผู้ใช้งานในระบบ' });
    }
    return res.json({
      user,
      storageQuota: await store.getQuotaStatus()
    });
  } catch (err) {
    console.error('GET /api/me error:', err);
    return res.status(500).json({ error: 'ไม่สามารถตรวจสอบข้อมูลผู้ใช้งานได้' });
  }
});

app.post('/api/logout', requireAuth, (req, res) => {
  if (req.token) {
    revokedTokens.add(req.token);
  }
  return res.json({ message: 'ออกจากระบบสำเร็จ' });
});

app.get('/api/storage-status', async (req, res) => {
  try {
    return res.json(await store.getQuotaStatus());
  } catch (err) {
    console.error('GET /api/storage-status error:', err);
    return res.status(500).json({ error: 'ไม่สามารถตรวจสอบสถานะพื้นที่เก็บข้อมูลได้' });
  }
});

app.post('/api/logs/clear-old', requireAuth, requireAdmin, async (req, res) => {
  try {
    const mode = req.body.mode || 'all';
    const result = await store.clearOldLogs(mode);
    return res.json({
      message: `ล้างข้อมูลบันทึกการเดินทางเรียบร้อยแล้ว (${result.deletedCount} รายการ)`,
      ...result
    });
  } catch (err) {
    console.error('POST /api/logs/clear-old error:', err);
    return res.status(500).json({ error: 'ไม่สามารถล้างข้อมูลเก่าได้' });
  }
});

app.post('/api/reset', requireAuth, requireAdmin, async (req, res) => {
  try {
    const storageQuota = await store.resetAllToSeed();
    return res.json({
      message: 'รีเซ็ตฐานข้อมูลกลับเป็นค่าเริ่มต้น (Seed Data) สำเร็จเรียบร้อย',
      storageQuota
    });
  } catch (err) {
    console.error('Database reset error:', err);
    return res.status(500).json({ error: 'ไม่สามารถรีเซ็ตฐานข้อมูลได้' });
  }
});

// ============================================================================
// 2. Settings APIs (/api/settings/gps)
// ============================================================================

app.get('/api/settings/gps', requireAuth, requireAdmin, async (req, res) => {
  try {
    return res.json({
      settings: await store.getSettingsMap(),
      storageQuota: await store.getQuotaStatus(),
      isSupabaseConfigured
    });
  } catch (err) {
    console.error('GET /api/settings/gps error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงค่าการตั้งค่าระบบได้' });
  }
});

app.put('/api/settings/gps', requireAuth, requireAdmin, async (req, res) => {
  try {
    const {
      cartrack_api_url,
      cartrack_username,
      cartrack_api_key,
      cartrack_mock_mode,
      max_storage_limit
    } = req.body;

    const limitNum = Math.max(1, Number(max_storage_limit) || 100);
    const updated = await store.saveSettingsMap({
      cartrack_api_url: String(
        cartrack_api_url || 'https://fleetapi-th.cartrack.com/rest/vehicles/status'
      ).trim(),
      cartrack_username: String(cartrack_username || '').trim(),
      cartrack_api_key: String(cartrack_api_key || '').trim(),
      cartrack_mock_mode:
        cartrack_mock_mode === false || cartrack_mock_mode === 'false' ? 'false' : 'true',
      max_storage_limit: String(limitNum)
    });

    return res.json({
      message: 'บันทึกการตั้งค่าเรียบร้อยแล้ว',
      settings: updated,
      storageQuota: await store.getQuotaStatus()
    });
  } catch (err) {
    console.error('PUT /api/settings/gps error:', err);
    return res.status(500).json({ error: 'ไม่สามารถบันทึกการตั้งค่าได้' });
  }
});

// ============================================================================
// 3. Master Workers APIs (/api/master/workers)
// ============================================================================

app.get('/api/master/workers', async (req, res) => {
  try {
    const activeOnly = req.query.active === '1' || req.query.active === 'true';
    return res.json(await store.listWorkers(activeOnly));
  } catch (err) {
    console.error('GET /api/master/workers error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงรายชื่อคนงานได้' });
  }
});

app.post('/api/master/workers', requireAuth, requireAdmin, async (req, res) => {
  try {
    const name = String(req.body.name || '').trim();
    const default_route = String(req.body.default_route || '').trim();
    const active = req.body.active === undefined ? true : Boolean(req.body.active);
    if (!name) {
      return res.status(400).json({ error: 'กรุณาระบุชื่อคนงาน' });
    }
    const created = await store.createWorker({ name, default_route, active });
    return res.status(201).json(created);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการเพิ่มรายชื่อคนงาน' });
  }
});

app.put('/api/master/workers/:id', requireAuth, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const name = req.body.name !== undefined ? String(req.body.name).trim() : undefined;
    const default_route =
      req.body.default_route !== undefined ? String(req.body.default_route).trim() : undefined;
    const active = req.body.active !== undefined ? Boolean(req.body.active) : undefined;

    if (name !== undefined && !name) {
      return res.status(400).json({ error: 'ชื่อคนงานห้ามเว้นว่าง' });
    }

    const updated = await store.updateWorker(id, { name, default_route, active }, req.user.role);
    return res.json(updated);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการแก้ไขข้อมูลคนงาน' });
  }
});

app.delete('/api/master/workers/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const deleted = await store.deleteWorker(id);
    return res.json({ message: 'ลบรายชื่อคนงานสำเร็จ', deleted });
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'ไม่สามารถลบรายชื่อคนงานได้' });
  }
});

// ============================================================================
// 4. Master Routes APIs (/api/master/routes)
// ============================================================================

app.get('/api/master/routes', async (req, res) => {
  try {
    const activeOnly = req.query.active === '1' || req.query.active === 'true';
    return res.json(await store.listRoutes(activeOnly));
  } catch (err) {
    console.error('GET /api/master/routes error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงรายการเส้นทาง/ที่หมายได้' });
  }
});

app.post('/api/master/routes', requireAuth, async (req, res) => {
  try {
    const name = String(req.body.name || '').trim();
    const active = req.body.active === undefined ? true : Boolean(req.body.active);
    if (!name) {
      return res.status(400).json({ error: 'กรุณาระบุชื่อเส้นทาง/ที่หมาย' });
    }
    const created = await store.createRoute({
      name,
      active,
      reactivateIfInactive: Boolean(req.body.reactivateIfInactive)
    });
    return res.status(created.reactivated ? 200 : 201).json(created);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการเพิ่มเส้นทาง/ที่หมาย' });
  }
});

app.put('/api/master/routes/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const name = req.body.name !== undefined ? String(req.body.name).trim() : undefined;
    const active = req.body.active !== undefined ? Boolean(req.body.active) : undefined;
    if (name !== undefined && !name) {
      return res.status(400).json({ error: 'ชื่อเส้นทาง/ที่หมายห้ามเว้นว่าง' });
    }
    const updated = await store.updateRoute(id, { name, active });
    return res.json(updated);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการแก้ไขเส้นทาง/ที่หมาย' });
  }
});

app.delete('/api/master/routes/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const deleted = await store.deleteRoute(id);
    return res.json({ message: 'ลบเส้นทาง/ที่หมายสำเร็จ', deleted });
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'ไม่สามารถลบเส้นทาง/ที่หมายได้' });
  }
});

// ============================================================================
// 5. Master Vehicles APIs (/api/master/vehicles)
// ============================================================================

app.get('/api/master/vehicles', async (req, res) => {
  try {
    const activeOnly = req.query.active === '1' || req.query.active === 'true';
    return res.json(await store.listVehicles(activeOnly));
  } catch (err) {
    console.error('GET /api/master/vehicles error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงรายการทะเบียนรถได้' });
  }
});

app.post('/api/master/vehicles', requireAuth, requireAdmin, async (req, res) => {
  try {
    const plate = String(req.body.plate || '').trim();
    const active = req.body.active === undefined ? true : Boolean(req.body.active);
    if (!plate) {
      return res.status(400).json({ error: 'กรุณาระบุหมายเลขทะเบียนรถ' });
    }
    const created = await store.createVehicle({ plate, active });
    return res.status(201).json(created);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการเพิ่มทะเบียนรถ' });
  }
});

app.put('/api/master/vehicles/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const plate = req.body.plate !== undefined ? String(req.body.plate).trim() : undefined;
    const active = req.body.active !== undefined ? Boolean(req.body.active) : undefined;
    if (plate !== undefined && !plate) {
      return res.status(400).json({ error: 'หมายเลขทะเบียนรถห้ามเว้นว่าง' });
    }
    const updated = await store.updateVehicle(id, { plate, active });
    return res.json(updated);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการแก้ไขทะเบียนรถ' });
  }
});

app.delete('/api/master/vehicles/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const deleted = await store.deleteVehicle(id);
    return res.json({ message: 'ลบทะเบียนรถสำเร็จ', deleted });
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'ไม่สามารถลบทะเบียนรถได้' });
  }
});

// ============================================================================
// 5B. Master Sites APIs (/api/master/sites - ไซต์งาน / โครงการ หัวตาราง)
// ============================================================================

app.get('/api/master/sites', async (req, res) => {
  try {
    const activeOnly = req.query.active === '1' || req.query.active === 'true';
    return res.json(await store.listSites(activeOnly));
  } catch (err) {
    console.error('GET /api/master/sites error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงรายการไซต์งาน/โครงการได้' });
  }
});

app.post('/api/master/sites', requireAuth, async (req, res) => {
  try {
    const name = String(req.body.name || '').trim();
    const active = req.body.active === undefined ? true : Boolean(req.body.active);
    if (!name) {
      return res.status(400).json({ error: 'กรุณาระบุชื่อไซต์งาน/โครงการ' });
    }
    const created = await store.createSite({
      name,
      active,
      reactivateIfInactive: Boolean(req.body.reactivateIfInactive)
    });
    return res.status(created.reactivated ? 200 : 201).json(created);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการเพิ่มไซต์งาน/โครงการ' });
  }
});

app.put('/api/master/sites/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const name = req.body.name !== undefined ? String(req.body.name).trim() : undefined;
    const active = req.body.active !== undefined ? Boolean(req.body.active) : undefined;
    if (name !== undefined && !name) {
      return res.status(400).json({ error: 'ชื่อไซต์งาน/โครงการห้ามเว้นว่าง' });
    }
    const updated = await store.updateSite(id, { name, active });
    return res.json(updated);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการแก้ไขไซต์งาน/โครงการ' });
  }
});

app.delete('/api/master/sites/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const deleted = await store.deleteSite(id);
    return res.json({ message: 'ลบไซต์งาน/โครงการสำเร็จ', deleted });
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'ไม่สามารถลบไซต์งาน/โครงการได้' });
  }
});

// ============================================================================
// 6. Travel Logs APIs (/api/logs)
// ============================================================================

app.get('/api/logs', async (req, res) => {
  try {
    const dateFilter = String(req.query.date || '').trim();
    const searchFilter = String(req.query.search || '').trim();
    return res.json(await store.listLogs(dateFilter, searchFilter));
  } catch (err) {
    console.error('GET /api/logs error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงข้อมูลบันทึกการเดินทางได้' });
  }
});

app.post('/api/logs', requireAuth, async (req, res) => {
  try {
    const worker_name = String(req.body.worker_name || '').trim();
    const site_name = String(req.body.site_name || '').trim();
    const travel_date = String(req.body.travel_date || '').trim();
    const route = String(req.body.route || '').trim();
    const departure_time = String(req.body.departure_time || '').trim();
    const arrival_time = String(req.body.arrival_time || '').trim();
    const vehicle_plate = String(req.body.vehicle_plate || '').trim();
    const passenger_count = req.body.passenger_count;

    if (!worker_name) {
      return res.status(400).json({ error: 'กรุณาเลือกชื่อคนงาน (Primary Key)' });
    }

    const created = await store.createLog({
      worker_name,
      site_name,
      travel_date,
      route,
      departure_time,
      arrival_time,
      vehicle_plate,
      passenger_count,
      gps_status: String(req.body.gps_status || '').trim()
    });
    return res.status(201).json(created);
  } catch (err) {
    return res.status(err.status || 500).json({
      error: err.message || 'เกิดข้อผิดพลาดในการบันทึกข้อมูลการเดินทาง',
      code: err.code,
      existing: err.existing,
      storageQuota: err.storageQuota
    });
  }
});

app.put('/api/logs/:worker_name', requireAuth, async (req, res) => {
  try {
    const worker_name = decodeURIComponent(req.params.worker_name || '').trim();
    const updates = {};
    if (req.body.site_name !== undefined)
      updates.site_name = String(req.body.site_name).trim();
    if (req.body.travel_date !== undefined)
      updates.travel_date = String(req.body.travel_date).trim();
    if (req.body.route !== undefined) updates.route = String(req.body.route).trim();
    if (req.body.departure_time !== undefined)
      updates.departure_time = String(req.body.departure_time).trim();
    if (req.body.arrival_time !== undefined)
      updates.arrival_time = String(req.body.arrival_time).trim();
    if (req.body.vehicle_plate !== undefined)
      updates.vehicle_plate = String(req.body.vehicle_plate).trim();
    if (req.body.passenger_count !== undefined)
      updates.passenger_count = req.body.passenger_count;

    const updated = await store.updateLog(worker_name, updates);
    return res.json(updated);
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'เกิดข้อผิดพลาดในการอัปเดตข้อมูลการเดินทาง' });
  }
});

app.delete('/api/logs/:worker_name', requireAuth, async (req, res) => {
  try {
    const worker_name = decodeURIComponent(req.params.worker_name || '').trim();
    const result = await store.deleteLog(worker_name);
    return res.json({
      message: `ลบข้อมูลการเดินทางของ "${worker_name}" สำเร็จ`,
      ...result
    });
  } catch (err) {
    return res
      .status(err.status || 500)
      .json({ error: err.message || 'ไม่สามารถลบข้อมูลการเดินทางได้' });
  }
});

// Fallback to index.html for SPA
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log('==================================================================');
    console.log('🚛 Worker Dispatch & Travel Log System v2.0');
    console.log(`🌐 URL:      http://localhost:${PORT}`);
    console.log(
      `☁️  Storage:  ${
        isSupabaseConfigured
          ? 'Supabase PostgreSQL Cloud (Direct Mode)'
          : 'Local Fallback Mode (SQLite + JSON)'
      }`
    );
    console.log('==================================================================');
  });
}

module.exports = app;
