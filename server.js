require('dotenv').config();
const express = require('express');
const path = require('node:path');
const crypto = require('node:crypto');
const {
  db,
  supabase,
  isSupabaseConfigured,
  DB_PATH,
  verifyPassword,
  seedDatabase,
  syncJsonFallbackFile,
  getSystemSettingsMap,
  updateSystemSettings,
  getStorageQuotaStatus,
  generateMockCartrackGps
} = require('./db');

const app = express();
const PORT = process.env.PORT || 3000;
const TOKEN_SECRET = process.env.TOKEN_SECRET || 'worker-dispatch-secret-key-2026';

// Initialize and seed database automatically on startup
seedDatabase(false);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static frontend files from /public
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
    return res.status(403).json({ error: 'เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่ได้รับสิทธิ์ใช้งานส่วนนี้' });
  }
  next();
}

// Helper: Fetch live Cartrack Fleet API for a vehicle or all vehicles
async function fetchCartrackLiveVehicles(settings) {
  const apiUrl = (settings.cartrack_api_url || 'https://fleetapi-th.cartrack.com/rest/vehicles/status').trim();
  const username = (settings.cartrack_username || '').trim();
  const apiKey = (settings.cartrack_api_key || '').trim();

  if (!username || !apiKey) {
    throw new Error('ยังไม่ได้ตั้งค่า Cartrack Username หรือ API Key ในระบบหลังบ้าน');
  }

  const basicAuth = Buffer.from(`${username}:${apiKey}`).toString('base64');
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 8000);

  try {
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        Authorization: `Basic ${basicAuth}`,
        Accept: 'application/json'
      },
      signal: controller.signal
    });
    clearTimeout(timeout);

    if (!response.ok) {
      throw new Error(`Cartrack API ตอบกลับด้วยสถานะ HTTP ${response.status}`);
    }

    const json = await response.json();
    const list = Array.isArray(json?.data) ? json.data : Array.isArray(json) ? json : [];
    return list;
  } catch (err) {
    clearTimeout(timeout);
    throw err;
  }
}

function normalizePlateForMatch(p) {
  return String(p || '')
    .replace(/[\s\-]/g, '')
    .toLowerCase();
}

function formatCartrackLiveItem(plate, item) {
  const lat = Number(item?.latitude ?? item?.lat ?? 13.7563);
  const lng = Number(item?.longitude ?? item?.lng ?? 100.5018);
  const speed = Number(item?.speed ?? 0);
  const ignition = Boolean(item?.ignition ?? item?.engine_on ?? speed > 0);
  const engine = ignition ? 'ON' : 'OFF';
  const status =
    speed > 0
      ? `กำลังวิ่ง ${speed} กม./ชม.`
      : ignition
      ? 'จอดอยู่ (ติดเครื่อง)'
      : 'จอดอยู่ (ดับเครื่อง)';
  const location =
    item?.location_description ||
    item?.location ||
    item?.road ||
    `พิกัด ${lat.toFixed(4)}, ${lng.toFixed(4)}`;

  return {
    plate,
    status,
    engine,
    speed,
    lat,
    lng,
    location,
    maps_url: `https://www.google.com/maps?q=${lat},${lng}`,
    updated_at:
      new Date().toLocaleTimeString('th-TH', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      }) + ' น.',
    source: 'live'
  };
}

// ============================================================================
// 1. Authentication, Session & System / Storage Quota APIs
// ============================================================================

app.post('/api/login', (req, res) => {
  try {
    // Support both explicit username ('admin' / 'user') or default to 'admin' if left blank
    const username = String(req.body.username || 'admin').trim() || 'admin';
    const password = String(req.body.password || '');

    if (!password) {
      return res.status(400).json({ error: 'กรุณากรอกชื่อผู้ใช้งานและรหัสผ่าน' });
    }

    const user = db
      .prepare('SELECT id, username, password, full_name, role FROM users WHERE username = ?')
      .get(username);

    if (!user || !verifyPassword(password, user.password)) {
      return res.status(401).json({ error: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง (ทดสอบใช้: admin/1234 หรือ user/1234)' });
    }

    const safeUser = {
      id: user.id,
      username: user.username,
      full_name: user.full_name,
      role: user.role
    };

    const token = createToken(safeUser);
    return res.json({
      message: 'เข้าสู่ระบบสำเร็จ',
      token,
      user: safeUser,
      storageQuota: getStorageQuotaStatus()
    });
  } catch (err) {
    console.error('Login error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเข้าสู่ระบบ' });
  }
});

app.get('/api/me', requireAuth, (req, res) => {
  const user = db
    .prepare('SELECT id, username, full_name, role FROM users WHERE id = ?')
    .get(req.user.id);
  if (!user) {
    return res.status(401).json({ error: 'ไม่พบข้อมูลผู้ใช้งานในระบบ' });
  }
  return res.json({
    user,
    storageQuota: getStorageQuotaStatus()
  });
});

app.post('/api/logout', requireAuth, (req, res) => {
  if (req.token) {
    revokedTokens.add(req.token);
  }
  return res.json({ message: 'ออกจากระบบสำเร็จ' });
});

app.get('/api/storage-status', (req, res) => {
  try {
    return res.json(getStorageQuotaStatus());
  } catch (err) {
    console.error('GET /api/storage-status error:', err);
    return res.status(500).json({ error: 'ไม่สามารถตรวจสอบสถานะพื้นที่เก็บข้อมูลได้' });
  }
});

app.post('/api/logs/clear-old', requireAuth, requireAdmin, (req, res) => {
  try {
    const mode = req.body.mode || 'all'; // 'all' | 'before_today'
    let deletedCount = 0;

    if (mode === 'before_today') {
      const now = new Date();
      const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
        now.getDate()
      ).padStart(2, '0')}`;
      const result = db.prepare('DELETE FROM travel_logs WHERE travel_date < ? OR travel_date = ""').run(today);
      deletedCount = Number(result.changes || 0);
      if (deletedCount === 0) {
        const allRes = db.prepare('DELETE FROM travel_logs').run();
        deletedCount = Number(allRes.changes || 0);
      }
    } else {
      const result = db.prepare('DELETE FROM travel_logs').run();
      deletedCount = Number(result.changes || 0);
    }

    syncJsonFallbackFile();
    return res.json({
      message: `ล้างข้อมูลบันทึกการเดินทางเรียบร้อยแล้ว (${deletedCount} รายการ)`,
      deletedCount,
      storageQuota: getStorageQuotaStatus()
    });
  } catch (err) {
    console.error('POST /api/logs/clear-old error:', err);
    return res.status(500).json({ error: 'ไม่สามารถล้างข้อมูลเก่าได้' });
  }
});

app.post('/api/reset', requireAuth, requireAdmin, (req, res) => {
  try {
    seedDatabase(true);
    return res.json({
      message: 'รีเซ็ตฐานข้อมูลกลับเป็นค่าเริ่มต้น (Seed Data) สำเร็จเรียบร้อย',
      storageQuota: getStorageQuotaStatus()
    });
  } catch (err) {
    console.error('Database reset error:', err);
    return res.status(500).json({ error: 'ไม่สามารถรีเซ็ตฐานข้อมูลได้' });
  }
});

// ============================================================================
// 2. Cartrack GPS Proxy & Settings APIs
// ============================================================================

app.get('/api/gps/vehicle/:plate', requireAuth, async (req, res) => {
  try {
    const plate = decodeURIComponent(req.params.plate || '').trim();
    if (!plate) {
      return res.status(400).json({ error: 'กรุณาระบุทะเบียนรถที่ต้องการตรวจสอบพิกัด GPS' });
    }

    const settings = getSystemSettingsMap();
    const isMock = String(settings.cartrack_mock_mode).toLowerCase() !== 'false';

    let gpsData = null;

    if (isMock) {
      gpsData = generateMockCartrackGps(plate);
    } else {
      const liveList = await fetchCartrackLiveVehicles(settings);
      const targetNorm = normalizePlateForMatch(plate);
      const matched = liveList.find(
        (v) => normalizePlateForMatch(v.registration || v.plate || '') === targetNorm
      );
      if (!matched) {
        return res.status(404).json({
          error: `ไม่พบข้อมูลพิกัดของทะเบียน "${plate}" จากระบบ Cartrack Live API`
        });
      }
      gpsData = formatCartrackLiveItem(plate, matched);
    }

    // Update any travel_logs rows currently assigned to this vehicle_plate
    db.prepare(
      "UPDATE travel_logs SET gps_status = ?, updated_at = datetime('now', 'localtime') WHERE vehicle_plate = ?"
    ).run(JSON.stringify(gpsData), plate);
    syncJsonFallbackFile();

    return res.json(gpsData);
  } catch (err) {
    console.error('GET /api/gps/vehicle/:plate error:', err);
    return res.status(502).json({
      error: err.message || 'ไม่สามารถเชื่อมต่อกับ Cartrack GPS API ได้'
    });
  }
});

app.post('/api/gps/sync-all', requireAuth, async (req, res) => {
  try {
    const settings = getSystemSettingsMap();
    const isMock = String(settings.cartrack_mock_mode).toLowerCase() !== 'false';
    const logs = db.prepare('SELECT worker_name, vehicle_plate FROM travel_logs').all();

    let liveList = [];
    if (!isMock) {
      liveList = await fetchCartrackLiveVehicles(settings);
    }

    const updateStmt = db.prepare(
      "UPDATE travel_logs SET gps_status = ?, updated_at = datetime('now', 'localtime') WHERE worker_name = ?"
    );

    let updatedCount = 0;
    db.exec('BEGIN');
    try {
      for (const row of logs) {
        if (!row.vehicle_plate) continue;
        let gpsData = null;
        if (isMock) {
          gpsData = generateMockCartrackGps(row.vehicle_plate);
        } else {
          const targetNorm = normalizePlateForMatch(row.vehicle_plate);
          const matched = liveList.find(
            (v) => normalizePlateForMatch(v.registration || v.plate || '') === targetNorm
          );
          if (matched) {
            gpsData = formatCartrackLiveItem(row.vehicle_plate, matched);
          }
        }

        if (gpsData) {
          updateStmt.run(JSON.stringify(gpsData), row.worker_name);
          updatedCount++;
        }
      }
      db.exec('COMMIT');
    } catch (txErr) {
      db.exec('ROLLBACK');
      throw txErr;
    }

    syncJsonFallbackFile();
    const updatedLogs = db
      .prepare('SELECT * FROM travel_logs ORDER BY travel_date DESC, departure_time ASC, worker_name ASC')
      .all();

    return res.json({
      message: `อัปเดตพิกัด GPS ของรถในรายการสำเร็จ (${updatedCount} คัน)`,
      updatedCount,
      source: isMock ? 'mock' : 'live',
      logs: updatedLogs
    });
  } catch (err) {
    console.error('POST /api/gps/sync-all error:', err);
    return res.status(502).json({
      error: err.message || 'ไม่สามารถซิงค์พิกัด GPS จาก Cartrack ได้'
    });
  }
});

app.get('/api/settings/gps', requireAuth, requireAdmin, (req, res) => {
  try {
    const settings = getSystemSettingsMap();
    return res.json({
      settings,
      storageQuota: getStorageQuotaStatus(),
      isSupabaseConfigured
    });
  } catch (err) {
    console.error('GET /api/settings/gps error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงค่าการตั้งค่าระบบได้' });
  }
});

app.put('/api/settings/gps', requireAuth, requireAdmin, (req, res) => {
  try {
    const {
      cartrack_api_url,
      cartrack_username,
      cartrack_api_key,
      cartrack_mock_mode,
      max_storage_limit
    } = req.body;

    const limitNum = Math.max(1, Number(max_storage_limit) || 100);

    const updated = updateSystemSettings({
      cartrack_api_url: String(cartrack_api_url || 'https://fleetapi-th.cartrack.com/rest/vehicles/status').trim(),
      cartrack_username: String(cartrack_username || '').trim(),
      cartrack_api_key: String(cartrack_api_key || '').trim(),
      cartrack_mock_mode: cartrack_mock_mode === false || cartrack_mock_mode === 'false' ? 'false' : 'true',
      max_storage_limit: String(limitNum)
    });

    return res.json({
      message: 'บันทึกการตั้งค่า Cartrack GPS และโควต้าพื้นที่เรียบร้อยแล้ว',
      settings: updated,
      storageQuota: getStorageQuotaStatus()
    });
  } catch (err) {
    console.error('PUT /api/settings/gps error:', err);
    return res.status(500).json({ error: 'ไม่สามารถบันทึกการตั้งค่าได้' });
  }
});

app.post('/api/settings/gps/test', requireAuth, requireAdmin, async (req, res) => {
  try {
    const current = getSystemSettingsMap();
    const testSettings = {
      cartrack_api_url: req.body.cartrack_api_url ?? current.cartrack_api_url,
      cartrack_username: req.body.cartrack_username ?? current.cartrack_username,
      cartrack_api_key: req.body.cartrack_api_key ?? current.cartrack_api_key,
      cartrack_mock_mode:
        req.body.cartrack_mock_mode !== undefined
          ? String(req.body.cartrack_mock_mode)
          : current.cartrack_mock_mode
    };

    const isMock = String(testSettings.cartrack_mock_mode).toLowerCase() !== 'false';
    if (isMock) {
      const sample = generateMockCartrackGps('34-5798');
      return res.json({
        success: true,
        mode: 'mock',
        message: 'เชื่อมต่อโหมดจำลองข้อมูล (Mock GPS Mode) สำเร็จ พร้อมใช้งานทันที',
        sample
      });
    }

    const liveList = await fetchCartrackLiveVehicles(testSettings);
    return res.json({
      success: true,
      mode: 'live',
      message: `เชื่อมต่อ Cartrack Live API สำเร็จ! พบข้อมูลรถในระบบ ${liveList.length} คัน`
    });
  } catch (err) {
    return res.status(400).json({
      success: false,
      error: `ทดสอบการเชื่อมต่อไม่สำเร็จ: ${err.message}`
    });
  }
});

// ============================================================================
// 3. Master Workers APIs (/api/master/workers)
// ============================================================================

app.get('/api/master/workers', (req, res) => {
  try {
    const activeOnly = req.query.active === '1' || req.query.active === 'true';
    const sql = activeOnly
      ? 'SELECT id, name, default_route, active FROM master_workers WHERE active = 1 ORDER BY id ASC'
      : 'SELECT id, name, default_route, active FROM master_workers ORDER BY id ASC';
    const rows = db.prepare(sql).all().map((r) => ({
      id: r.id,
      name: r.name,
      default_route: r.default_route || '',
      active: Boolean(r.active)
    }));
    return res.json(rows);
  } catch (err) {
    console.error('GET /api/master/workers error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงรายชื่อคนงานได้' });
  }
});

app.post('/api/master/workers', requireAuth, requireAdmin, (req, res) => {
  try {
    const name = String(req.body.name || '').trim();
    const default_route = String(req.body.default_route || '').trim();
    const active = req.body.active === undefined ? 1 : req.body.active ? 1 : 0;

    if (!name) {
      return res.status(400).json({ error: 'กรุณาระบุชื่อคนงาน' });
    }

    const existing = db.prepare('SELECT id FROM master_workers WHERE name = ?').get(name);
    if (existing) {
      return res.status(409).json({ error: `ชื่อคนงาน "${name}" มีอยู่ในระบบแล้ว` });
    }

    const result = db
      .prepare('INSERT INTO master_workers (name, default_route, active) VALUES (?, ?, ?)')
      .run(name, default_route, active);

    syncJsonFallbackFile();
    const created = db
      .prepare('SELECT id, name, default_route, active FROM master_workers WHERE id = ?')
      .get(result.lastInsertRowid);

    return res.status(201).json({
      id: created.id,
      name: created.name,
      default_route: created.default_route || '',
      active: Boolean(created.active)
    });
  } catch (err) {
    console.error('POST /api/master/workers error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเพิ่มรายชื่อคนงาน' });
  }
});

app.put('/api/master/workers/:id', requireAuth, (req, res) => {
  try {
    const id = Number(req.params.id);
    const existing = db
      .prepare('SELECT id, name, default_route, active FROM master_workers WHERE id = ?')
      .get(id);

    if (!existing) {
      return res.status(404).json({ error: 'ไม่พบข้อมูลคนงานที่ต้องการแก้ไข' });
    }

    const isRenaming = req.body.name !== undefined && String(req.body.name).trim() !== existing.name;
    const isTogglingActive =
      req.body.active !== undefined && Boolean(req.body.active) !== Boolean(existing.active);
    if ((isRenaming || isTogglingActive) && req.user.role !== 'admin') {
      return res.status(403).json({
        error: 'เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่สามารถเปลี่ยนชื่อหรือสถานะคนงานได้'
      });
    }

    const newName = req.body.name !== undefined ? String(req.body.name).trim() : existing.name;
    const newDefaultRoute =
      req.body.default_route !== undefined
        ? String(req.body.default_route).trim()
        : existing.default_route || '';
    const newActive =
      req.body.active !== undefined ? (req.body.active ? 1 : 0) : existing.active ? 1 : 0;

    if (!newName) {
      return res.status(400).json({ error: 'ชื่อคนงานห้ามเว้นว่าง' });
    }

    if (newName !== existing.name) {
      const duplicate = db
        .prepare('SELECT id FROM master_workers WHERE name = ? AND id != ?')
        .get(newName, id);
      if (duplicate) {
        return res.status(409).json({ error: `ชื่อคนงาน "${newName}" มีอยู่ในระบบแล้ว` });
      }
    }

    let cascadedLogsCount = 0;
    db.exec('BEGIN');
    try {
      db.prepare(
        'UPDATE master_workers SET name = ?, default_route = ?, active = ? WHERE id = ?'
      ).run(newName, newDefaultRoute, newActive, id);

      if (newName !== existing.name) {
        const cascadeRes = db
          .prepare(
            "UPDATE travel_logs SET worker_name = ?, updated_at = datetime('now', 'localtime') WHERE worker_name = ?"
          )
          .run(newName, existing.name);
        cascadedLogsCount = Number(cascadeRes.changes || 0);
      }
      db.exec('COMMIT');
    } catch (txErr) {
      db.exec('ROLLBACK');
      throw txErr;
    }

    syncJsonFallbackFile();
    const updated = db
      .prepare('SELECT id, name, default_route, active FROM master_workers WHERE id = ?')
      .get(id);

    return res.json({
      id: updated.id,
      name: updated.name,
      default_route: updated.default_route || '',
      active: Boolean(updated.active),
      cascadedLogsCount
    });
  } catch (err) {
    console.error('PUT /api/master/workers/:id error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการแก้ไขข้อมูลคนงาน' });
  }
});

app.delete('/api/master/workers/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const existing = db.prepare('SELECT id, name FROM master_workers WHERE id = ?').get(id);
    if (!existing) {
      return res.status(404).json({ error: 'ไม่พบรายชื่อคนงานที่ต้องการลบ' });
    }
    db.prepare('DELETE FROM master_workers WHERE id = ?').run(id);
    if (isSupabaseConfigured && supabase) {
      await supabase.from('master_workers').delete().eq('name', existing.name);
    }
    syncJsonFallbackFile();
    return res.json({ message: 'ลบรายชื่อคนงานสำเร็จ', deleted: existing });
  } catch (err) {
    console.error('DELETE /api/master/workers/:id error:', err);
    return res.status(500).json({ error: 'ไม่สามารถลบรายชื่อคนงานได้' });
  }
});

// ============================================================================
// 4. Master Routes APIs (/api/master/routes)
// ============================================================================

app.get('/api/master/routes', (req, res) => {
  try {
    const activeOnly = req.query.active === '1' || req.query.active === 'true';
    const sql = activeOnly
      ? 'SELECT id, name, active FROM master_routes WHERE active = 1 ORDER BY id ASC'
      : 'SELECT id, name, active FROM master_routes ORDER BY id ASC';
    const rows = db.prepare(sql).all().map((r) => ({
      id: r.id,
      name: r.name,
      active: Boolean(r.active)
    }));
    return res.json(rows);
  } catch (err) {
    console.error('GET /api/master/routes error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงรายการเส้นทาง/ที่หมายได้' });
  }
});

app.post('/api/master/routes', requireAuth, (req, res) => {
  try {
    const name = String(req.body.name || '').trim();
    const active = req.body.active === undefined ? 1 : req.body.active ? 1 : 0;

    if (!name) {
      return res.status(400).json({ error: 'กรุณาระบุชื่อเส้นทาง/ที่หมาย' });
    }

    const existing = db.prepare('SELECT id, name, active FROM master_routes WHERE name = ?').get(name);
    if (existing) {
      if (!existing.active && req.body.reactivateIfInactive) {
        db.prepare('UPDATE master_routes SET active = 1 WHERE id = ?').run(existing.id);
        syncJsonFallbackFile();
        return res
          .status(200)
          .json({ id: existing.id, name: existing.name, active: true, reactivated: true });
      }
      return res.status(409).json({ error: `เส้นทาง/ที่หมาย "${name}" มีอยู่ในระบบแล้ว` });
    }

    const result = db
      .prepare('INSERT INTO master_routes (name, active) VALUES (?, ?)')
      .run(name, active);

    syncJsonFallbackFile();
    const created = db
      .prepare('SELECT id, name, active FROM master_routes WHERE id = ?')
      .get(result.lastInsertRowid);

    return res.status(201).json({
      id: created.id,
      name: created.name,
      active: Boolean(created.active)
    });
  } catch (err) {
    console.error('POST /api/master/routes error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเพิ่มเส้นทาง/ที่หมาย' });
  }
});

app.put('/api/master/routes/:id', requireAuth, requireAdmin, (req, res) => {
  try {
    const id = Number(req.params.id);
    const existing = db
      .prepare('SELECT id, name, active FROM master_routes WHERE id = ?')
      .get(id);

    if (!existing) {
      return res.status(404).json({ error: 'ไม่พบเส้นทาง/ที่หมายที่ต้องการแก้ไข' });
    }

    const newName = req.body.name !== undefined ? String(req.body.name).trim() : existing.name;
    const newActive =
      req.body.active !== undefined ? (req.body.active ? 1 : 0) : existing.active ? 1 : 0;

    if (!newName) {
      return res.status(400).json({ error: 'ชื่อเส้นทาง/ที่หมายห้ามเว้นว่าง' });
    }

    if (newName !== existing.name) {
      const duplicate = db
        .prepare('SELECT id FROM master_routes WHERE name = ? AND id != ?')
        .get(newName, id);
      if (duplicate) {
        return res.status(409).json({ error: `เส้นทาง/ที่หมาย "${newName}" มีอยู่ในระบบแล้ว` });
      }
    }

    let cascadedLogsCount = 0;
    let cascadedWorkersCount = 0;
    db.exec('BEGIN');
    try {
      db.prepare('UPDATE master_routes SET name = ?, active = ? WHERE id = ?').run(
        newName,
        newActive,
        id
      );

      if (newName !== existing.name) {
        const logRes = db
          .prepare(
            "UPDATE travel_logs SET route = ?, updated_at = datetime('now', 'localtime') WHERE route = ?"
          )
          .run(newName, existing.name);
        cascadedLogsCount = Number(logRes.changes || 0);

        const workerRes = db
          .prepare('UPDATE master_workers SET default_route = ? WHERE default_route = ?')
          .run(newName, existing.name);
        cascadedWorkersCount = Number(workerRes.changes || 0);
      }

      db.exec('COMMIT');
    } catch (txErr) {
      db.exec('ROLLBACK');
      throw txErr;
    }

    syncJsonFallbackFile();
    const updated = db
      .prepare('SELECT id, name, active FROM master_routes WHERE id = ?')
      .get(id);

    return res.json({
      id: updated.id,
      name: updated.name,
      active: Boolean(updated.active),
      cascadedLogsCount,
      cascadedWorkersCount
    });
  } catch (err) {
    console.error('PUT /api/master/routes/:id error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการแก้ไขเส้นทาง/ที่หมาย' });
  }
});

app.delete('/api/master/routes/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const existing = db.prepare('SELECT id, name FROM master_routes WHERE id = ?').get(id);
    if (!existing) {
      return res.status(404).json({ error: 'ไม่พบเส้นทาง/ที่หมายที่ต้องการลบ' });
    }

    db.exec('BEGIN');
    try {
      db.prepare('DELETE FROM master_routes WHERE id = ?').run(id);
      db.prepare("UPDATE master_workers SET default_route = '' WHERE default_route = ?").run(
        existing.name
      );
      db.exec('COMMIT');
    } catch (txErr) {
      db.exec('ROLLBACK');
      throw txErr;
    }

    if (isSupabaseConfigured && supabase) {
      await supabase.from('master_routes').delete().eq('name', existing.name);
    }
    syncJsonFallbackFile();
    return res.json({ message: 'ลบเส้นทาง/ที่หมายสำเร็จ', deleted: existing });
  } catch (err) {
    console.error('DELETE /api/master/routes/:id error:', err);
    return res.status(500).json({ error: 'ไม่สามารถลบเส้นทาง/ที่หมายได้' });
  }
});

// ============================================================================
// 5. Master Vehicles APIs (/api/master/vehicles)
// ============================================================================

app.get('/api/master/vehicles', (req, res) => {
  try {
    const activeOnly = req.query.active === '1' || req.query.active === 'true';
    const sql = activeOnly
      ? 'SELECT id, plate, active FROM master_vehicles WHERE active = 1 ORDER BY id ASC'
      : 'SELECT id, plate, active FROM master_vehicles ORDER BY id ASC';
    const rows = db.prepare(sql).all().map((r) => ({
      id: r.id,
      plate: r.plate,
      active: Boolean(r.active)
    }));
    return res.json(rows);
  } catch (err) {
    console.error('GET /api/master/vehicles error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงรายการทะเบียนรถได้' });
  }
});

app.post('/api/master/vehicles', requireAuth, requireAdmin, (req, res) => {
  try {
    const plate = String(req.body.plate || '').trim();
    const active = req.body.active === undefined ? 1 : req.body.active ? 1 : 0;

    if (!plate) {
      return res.status(400).json({ error: 'กรุณาระบุหมายเลขทะเบียนรถ' });
    }

    const existing = db.prepare('SELECT id FROM master_vehicles WHERE plate = ?').get(plate);
    if (existing) {
      return res.status(409).json({ error: `ทะเบียนรถ "${plate}" มีอยู่ในระบบแล้ว` });
    }

    const result = db
      .prepare('INSERT INTO master_vehicles (plate, active) VALUES (?, ?)')
      .run(plate, active);

    syncJsonFallbackFile();
    const created = db
      .prepare('SELECT id, plate, active FROM master_vehicles WHERE id = ?')
      .get(result.lastInsertRowid);

    return res.status(201).json({
      id: created.id,
      plate: created.plate,
      active: Boolean(created.active)
    });
  } catch (err) {
    console.error('POST /api/master/vehicles error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเพิ่มทะเบียนรถ' });
  }
});

app.put('/api/master/vehicles/:id', requireAuth, requireAdmin, (req, res) => {
  try {
    const id = Number(req.params.id);
    const existing = db
      .prepare('SELECT id, plate, active FROM master_vehicles WHERE id = ?')
      .get(id);

    if (!existing) {
      return res.status(404).json({ error: 'ไม่พบทะเบียนรถที่ต้องการแก้ไข' });
    }

    const newPlate = req.body.plate !== undefined ? String(req.body.plate).trim() : existing.plate;
    const newActive =
      req.body.active !== undefined ? (req.body.active ? 1 : 0) : existing.active ? 1 : 0;

    if (!newPlate) {
      return res.status(400).json({ error: 'หมายเลขทะเบียนรถห้ามเว้นว่าง' });
    }

    if (newPlate !== existing.plate) {
      const duplicate = db
        .prepare('SELECT id FROM master_vehicles WHERE plate = ? AND id != ?')
        .get(newPlate, id);
      if (duplicate) {
        return res.status(409).json({ error: `ทะเบียนรถ "${newPlate}" มีอยู่ในระบบแล้ว` });
      }
    }

    let cascadedLogsCount = 0;
    db.exec('BEGIN');
    try {
      db.prepare('UPDATE master_vehicles SET plate = ?, active = ? WHERE id = ?').run(
        newPlate,
        newActive,
        id
      );

      if (newPlate !== existing.plate) {
        const logRes = db
          .prepare(
            "UPDATE travel_logs SET vehicle_plate = ?, updated_at = datetime('now', 'localtime') WHERE vehicle_plate = ?"
          )
          .run(newPlate, existing.plate);
        cascadedLogsCount = Number(logRes.changes || 0);
      }

      db.exec('COMMIT');
    } catch (txErr) {
      db.exec('ROLLBACK');
      throw txErr;
    }

    syncJsonFallbackFile();
    const updated = db
      .prepare('SELECT id, plate, active FROM master_vehicles WHERE id = ?')
      .get(id);

    return res.json({
      id: updated.id,
      plate: updated.plate,
      active: Boolean(updated.active),
      cascadedLogsCount
    });
  } catch (err) {
    console.error('PUT /api/master/vehicles/:id error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการแก้ไขทะเบียนรถ' });
  }
});

app.delete('/api/master/vehicles/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const existing = db.prepare('SELECT id, plate FROM master_vehicles WHERE id = ?').get(id);
    if (!existing) {
      return res.status(404).json({ error: 'ไม่พบทะเบียนรถที่ต้องการลบ' });
    }
    db.prepare('DELETE FROM master_vehicles WHERE id = ?').run(id);
    if (isSupabaseConfigured && supabase) {
      await supabase.from('master_vehicles').delete().eq('plate', existing.plate);
    }
    syncJsonFallbackFile();
    return res.json({ message: 'ลบทะเบียนรถสำเร็จ', deleted: existing });
  } catch (err) {
    console.error('DELETE /api/master/vehicles/:id error:', err);
    return res.status(500).json({ error: 'ไม่สามารถลบทะเบียนรถได้' });
  }
});

// ============================================================================
// 6. Travel Logs APIs (/api/logs) + Storage Quota Check (HTTP 507)
// ============================================================================

app.get('/api/logs', (req, res) => {
  try {
    const dateFilter = String(req.query.date || '').trim();
    const searchFilter = String(req.query.search || '').trim().toLowerCase();

    let rows = db
      .prepare(
        `SELECT worker_name, travel_date, route, departure_time, arrival_time, vehicle_plate, gps_status, updated_at
         FROM travel_logs
         ORDER BY travel_date DESC, departure_time ASC, worker_name ASC`
      )
      .all();

    if (dateFilter) {
      rows = rows.filter((r) => (r.travel_date || '') === dateFilter);
    }

    if (searchFilter) {
      rows = rows.filter((r) => {
        const combined = [
          r.worker_name || '',
          r.travel_date || '',
          r.route || '',
          r.vehicle_plate || '',
          r.departure_time || '',
          r.arrival_time || '',
          r.gps_status || ''
        ]
          .join(' ')
          .toLowerCase();
        return combined.includes(searchFilter);
      });
    }

    return res.json(rows);
  } catch (err) {
    console.error('GET /api/logs error:', err);
    return res.status(500).json({ error: 'ไม่สามารถดึงข้อมูลบันทึกการเดินทางได้' });
  }
});

app.post('/api/logs', requireAuth, (req, res) => {
  try {
    const worker_name = String(req.body.worker_name || '').trim();
    const travel_date = String(req.body.travel_date || '').trim();
    const route = String(req.body.route || '').trim();
    const departure_time = String(req.body.departure_time || '').trim();
    const arrival_time = String(req.body.arrival_time || '').trim();
    const vehicle_plate = String(req.body.vehicle_plate || '').trim();

    if (!worker_name) {
      return res.status(400).json({ error: 'กรุณาเลือกชื่อคนงาน (Primary Key)' });
    }

    // 1. Check Duplicate Primary Key (HTTP 409 Conflict)
    const existing = db
      .prepare('SELECT * FROM travel_logs WHERE worker_name = ?')
      .get(worker_name);

    if (existing) {
      return res.status(409).json({
        error: `คนงาน "${worker_name}" มีข้อมูลการเดินทางในระบบแล้ว ห้ามบันทึกซ้ำ กรุณากดปุ่ม "แก้ไข" แทน`,
        code: 'DUPLICATE_WORKER_PK',
        existing
      });
    }

    // 2. Check Storage Quota Limit (HTTP 507 Insufficient Storage)
    const quota = getStorageQuotaStatus();
    if (quota.isFull) {
      return res.status(507).json({
        error: `พื้นที่เก็บข้อมูลเต็ม (${quota.currentCount}/${quota.maxLimit} รายการ - 100%) ไม่สามารถบันทึกข้อมูลใหม่ได้ชั่วคราว กรุณาส่งออกไฟล์สำรองข้อมูลหรือล้างข้อมูลเก่า`,
        code: 'STORAGE_FULL',
        storageQuota: quota
      });
    }

    // 3. Generate initial GPS status if vehicle_plate is provided
    let gps_status = String(req.body.gps_status || '').trim();
    if (!gps_status && vehicle_plate) {
      const settings = getSystemSettingsMap();
      if (String(settings.cartrack_mock_mode).toLowerCase() !== 'false') {
        gps_status = JSON.stringify(generateMockCartrackGps(vehicle_plate));
      }
    }

    db.prepare(
      `INSERT INTO travel_logs (worker_name, travel_date, route, departure_time, arrival_time, vehicle_plate, gps_status, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now', 'localtime'))`
    ).run(worker_name, travel_date, route, departure_time, arrival_time, vehicle_plate, gps_status);

    syncJsonFallbackFile();
    const created = db
      .prepare('SELECT * FROM travel_logs WHERE worker_name = ?')
      .get(worker_name);

    return res.status(201).json({
      ...created,
      storageQuota: getStorageQuotaStatus()
    });
  } catch (err) {
    console.error('POST /api/logs error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการบันทึกข้อมูลการเดินทาง' });
  }
});

app.put('/api/logs/:worker_name', requireAuth, (req, res) => {
  try {
    const worker_name = decodeURIComponent(req.params.worker_name || '').trim();
    const existing = db
      .prepare('SELECT * FROM travel_logs WHERE worker_name = ?')
      .get(worker_name);

    if (!existing) {
      return res.status(404).json({ error: `ไม่พบข้อมูลการเดินทางของ "${worker_name}"` });
    }

    const travel_date =
      req.body.travel_date !== undefined ? String(req.body.travel_date).trim() : existing.travel_date;
    const route =
      req.body.route !== undefined ? String(req.body.route).trim() : existing.route;
    const departure_time =
      req.body.departure_time !== undefined
        ? String(req.body.departure_time).trim()
        : existing.departure_time;
    const arrival_time =
      req.body.arrival_time !== undefined
        ? String(req.body.arrival_time).trim()
        : existing.arrival_time;
    const vehicle_plate =
      req.body.vehicle_plate !== undefined
        ? String(req.body.vehicle_plate).trim()
        : existing.vehicle_plate;

    let gps_status =
      req.body.gps_status !== undefined ? String(req.body.gps_status).trim() : existing.gps_status || '';
    if (vehicle_plate && vehicle_plate !== existing.vehicle_plate && !req.body.gps_status) {
      gps_status = JSON.stringify(generateMockCartrackGps(vehicle_plate));
    } else if (!vehicle_plate) {
      gps_status = '';
    }

    db.prepare(
      `UPDATE travel_logs
       SET travel_date = ?, route = ?, departure_time = ?, arrival_time = ?, vehicle_plate = ?, gps_status = ?, updated_at = datetime('now', 'localtime')
       WHERE worker_name = ?`
    ).run(
      travel_date,
      route,
      departure_time,
      arrival_time,
      vehicle_plate,
      gps_status,
      worker_name
    );

    syncJsonFallbackFile();
    const updated = db
      .prepare('SELECT * FROM travel_logs WHERE worker_name = ?')
      .get(worker_name);

    return res.json({
      ...updated,
      storageQuota: getStorageQuotaStatus()
    });
  } catch (err) {
    console.error('PUT /api/logs/:worker_name error:', err);
    return res.status(500).json({ error: 'เกิดข้อผิดพลาดในการอัปเดตข้อมูลการเดินทาง' });
  }
});

app.delete('/api/logs/:worker_name', requireAuth, async (req, res) => {
  try {
    const worker_name = decodeURIComponent(req.params.worker_name || '').trim();
    const existing = db
      .prepare('SELECT * FROM travel_logs WHERE worker_name = ?')
      .get(worker_name);

    if (!existing) {
      return res.status(404).json({ error: `ไม่พบข้อมูลการเดินทางของ "${worker_name}"` });
    }

    db.prepare('DELETE FROM travel_logs WHERE worker_name = ?').run(worker_name);
    if (isSupabaseConfigured && supabase) {
      await supabase.from('travel_logs').delete().eq('worker_name', worker_name);
    }
    syncJsonFallbackFile();
    return res.json({
      message: `ลบข้อมูลการเดินทางของ "${worker_name}" สำเร็จ`,
      deleted: existing,
      storageQuota: getStorageQuotaStatus()
    });
  } catch (err) {
    console.error('DELETE /api/logs/:worker_name error:', err);
    return res.status(500).json({ error: 'ไม่สามารถลบข้อมูลการเดินทางได้' });
  }
});

// Fallback to index.html for SPA
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log('==================================================================');
    console.log('🚛 Worker Dispatch & Travel Log System v2.0 (Cloud + Cartrack GPS)');
    console.log(`🌐 URL:      http://localhost:${PORT}`);
    console.log(
      `☁️  Storage:  ${
        isSupabaseConfigured
          ? 'Supabase PostgreSQL Cloud Connected'
          : 'Local Fallback Mode (SQLite + JSON)'
      }`
    );
    console.log('🔑 Demo Accounts:');
    console.log('   - Admin:    admin / 1234');
    console.log('   - Operator: user  / 1234');
    console.log('==================================================================');
  });
}

module.exports = app;
