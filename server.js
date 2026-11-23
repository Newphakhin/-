require('dotenv').config();
const express = require('express');
const path = require('node:path');
const crypto = require('node:crypto');
const { isSupabaseConfigured } = require('./db');
const store = require('./store');

const app = express();
const PORT = process.env.PORT || 3000;
const TOKEN_SECRET = process.env.TOKEN_SECRET || 'worker-dispatch-secret-key-2026';

app.use(express.json({ limit: '1mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// ----------------------------------------------------------------------------
// Token (HMAC-SHA256 แบบ stateless + รายการ revoke)
// ----------------------------------------------------------------------------
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
  const signature = crypto.createHmac('sha256', TOKEN_SECRET).update(base64Payload).digest('base64url');
  return `${base64Payload}.${signature}`;
}

function verifyToken(token) {
  if (!token || revokedTokens.has(token)) return null;
  const parts = token.split('.');
  if (parts.length !== 2) return null;
  const [base64Payload, signature] = parts;
  const expected = crypto.createHmac('sha256', TOKEN_SECRET).update(base64Payload).digest('base64url');
  if (signature !== expected) return null;
  try {
    return JSON.parse(Buffer.from(base64Payload, 'base64url').toString('utf8'));
  } catch {
    return null;
  }
}

function extractToken(req) {
  const auth = req.headers.authorization;
  if (auth && auth.startsWith('Bearer ')) return auth.slice(7).trim();
  if (req.headers['x-auth-token']) return String(req.headers['x-auth-token']).trim();
  return null;
}

function requireAuth(req, res, next) {
  const token = extractToken(req);
  const payload = verifyToken(token);
  if (!payload) return res.status(401).json({ error: 'กรุณาเข้าสู่ระบบก่อนใช้งาน (Unauthorized)' });
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

// ห่อ handler ให้ส่ง error เป็น JSON พร้อม status ที่เหมาะสมเสมอ
const handle = (fn, fallbackMessage) => async (req, res) => {
  try {
    await fn(req, res);
  } catch (err) {
    if (!err.status) console.error(`${req.method} ${req.originalUrl}:`, err);
    res.status(err.status || 500).json({
      error: err.status ? err.message : fallbackMessage || err.message || 'เกิดข้อผิดพลาดของระบบ',
      code: err.code,
      storageQuota: err.storageQuota
    });
  }
};

// ----------------------------------------------------------------------------
// 1. Auth & ระบบ
// ----------------------------------------------------------------------------
app.post(
  '/api/login',
  handle(async (req, res) => {
    const username = String(req.body.username || '').trim();
    const password = String(req.body.password || '');
    if (!username || !password) {
      return res.status(400).json({ error: 'กรุณากรอกชื่อผู้ใช้งานและรหัสผ่าน' });
    }
    const user = await store.findUserByUsername(username);
    if (!user || !store.verifyPassword(password, user.password)) {
      return res.status(401).json({ error: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' });
    }
    const safeUser = { id: user.id, username: user.username, full_name: user.full_name, role: user.role };
    res.json({
      message: 'เข้าสู่ระบบสำเร็จ',
      token: createToken(safeUser),
      user: safeUser,
      storageQuota: await store.getQuotaStatus(),
      storageMode: isSupabaseConfigured ? 'supabase_cloud' : 'local_sqlite'
    });
  }, 'เกิดข้อผิดพลาดในการเข้าสู่ระบบ')
);

app.get(
  '/api/me',
  requireAuth,
  handle(async (req, res) => {
    const user = await store.findUserById(req.user.id);
    if (!user) return res.status(401).json({ error: 'ไม่พบข้อมูลผู้ใช้งานในระบบ' });
    res.json({ user, storageQuota: await store.getQuotaStatus() });
  })
);

app.post('/api/logout', requireAuth, (req, res) => {
  revokedTokens.add(req.token);
  res.json({ message: 'ออกจากระบบสำเร็จ' });
});

app.get(
  '/api/storage-status',
  requireAuth,
  handle(async (req, res) => res.json(await store.getQuotaStatus()))
);

app.get(
  '/api/settings',
  requireAuth,
  requireAdmin,
  handle(async (req, res) =>
    res.json({ settings: await store.getSettings(), storageQuota: await store.getQuotaStatus() })
  )
);

app.put(
  '/api/settings',
  requireAuth,
  requireAdmin,
  handle(async (req, res) => {
    const settings = await store.saveSettings({ max_storage_limit: req.body.max_storage_limit });
    res.json({ message: 'บันทึกการตั้งค่าเรียบร้อยแล้ว', settings, storageQuota: await store.getQuotaStatus() });
  })
);

// ----------------------------------------------------------------------------
// 2. ข้อมูลหลัก /api/master/*  (ดูได้ทุกคนที่ล็อกอิน / แก้ไขเฉพาะแอดมิน)
// ----------------------------------------------------------------------------
function registerMaster(name, api, labels) {
  app.get(
    `/api/master/${name}`,
    requireAuth,
    handle(async (req, res) => {
      const activeOnly = req.query.active === '1' || req.query.active === 'true';
      res.json(await api.list(activeOnly));
    }, `ไม่สามารถดึงข้อมูล${labels}ได้`)
  );
  app.post(
    `/api/master/${name}`,
    requireAuth,
    requireAdmin,
    handle(async (req, res) => res.status(201).json(await api.create(req.body || {})))
  );
  app.put(
    `/api/master/${name}/:id`,
    requireAuth,
    requireAdmin,
    handle(async (req, res) => res.json(await api.update(Number(req.params.id), req.body || {})))
  );
  app.delete(
    `/api/master/${name}/:id`,
    requireAuth,
    requireAdmin,
    handle(async (req, res) => {
      const deleted = await api.remove(Number(req.params.id));
      res.json({ message: `ลบ${labels}สำเร็จ`, deleted });
    })
  );
}

registerMaster(
  'workers',
  { list: store.listWorkers, create: store.createWorker, update: store.updateWorker, remove: store.deleteWorker },
  'รายชื่อคนขับ'
);
registerMaster(
  'vehicles',
  { list: store.listVehicles, create: store.createVehicle, update: store.updateVehicle, remove: store.deleteVehicle },
  'ทะเบียนรถ'
);
registerMaster(
  'sites',
  { list: store.listSites, create: store.createSite, update: store.updateSite, remove: store.deleteSite },
  'ไซต์งาน/โครงการ'
);
registerMaster(
  'routes',
  { list: store.listRoutes, create: store.createRoute, update: store.updateRoute, remove: store.deleteRoute },
  'เส้นทาง'
);

// ----------------------------------------------------------------------------
// 3. บันทึกการเดินทาง /api/logs
// ----------------------------------------------------------------------------
app.get(
  '/api/logs',
  requireAuth,
  handle(async (req, res) => {
    res.json(
      await store.listLogs({
        date_from: String(req.query.date_from || '').trim(),
        date_to: String(req.query.date_to || '').trim(),
        site_name: String(req.query.site_name || '').trim(),
        trip_type: String(req.query.trip_type || '').trim(),
        search: String(req.query.search || '').trim()
      })
    );
  }, 'ไม่สามารถดึงข้อมูลบันทึกการเดินทางได้')
);

app.post(
  '/api/logs',
  requireAuth,
  handle(async (req, res) => res.status(201).json(await store.createLog(req.body || {})))
);

app.post(
  '/api/logs/copy-previous',
  requireAuth,
  handle(async (req, res) => {
    const result = await store.copyPreviousDay({
      to_date: req.body.to_date,
      from_date: req.body.from_date
    });
    res.status(201).json({ message: `คัดลอกคิววิ่งจากวันที่ ${result.from_date} จำนวน ${result.created} รายการ`, ...result });
  })
);

app.post(
  '/api/logs/clear',
  requireAuth,
  requireAdmin,
  handle(async (req, res) => {
    const result = await store.clearAllLogs();
    res.json({ message: `ล้างบันทึกการเดินทางทั้งหมดแล้ว (${result.deletedCount} รายการ)`, ...result });
  })
);

app.put(
  '/api/logs/:id',
  requireAuth,
  handle(async (req, res) => res.json(await store.updateLog(String(req.params.id), req.body || {})))
);

app.delete(
  '/api/logs/:id',
  requireAuth,
  handle(async (req, res) => {
    const result = await store.deleteLog(String(req.params.id));
    res.json({ message: 'ลบรายการเดินทางสำเร็จ', ...result });
  })
);

// ----------------------------------------------------------------------------
// 4. เชื่อมต่อและไหลข้อมูลลง Google Sheets (ผ่าน Apps Script Web App Proxy)
// ----------------------------------------------------------------------------
async function postToAppsScript(webAppUrl, payload) {
  const safeUrl = store.validateAppsScriptUrl(webAppUrl, 'Google Apps Script');
  if (!safeUrl) {
    const err = new Error('ยังไม่ได้ตั้งค่าลิงก์ Google Apps Script');
    err.status = 400;
    throw err;
  }

  let response;
  try {
    response = await fetch(safeUrl, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(28000)
    });
  } catch (netErr) {
    const err = new Error(`ไม่สามารถเชื่อมต่อ Google Sheets ได้ (${netErr.message})`);
    err.status = 502;
    throw err;
  }

  const rawText = await response.text();
  let parsed = null;
  try {
    parsed = JSON.parse(rawText);
  } catch {
    if (rawText.includes('<html') || rawText.includes('<!DOCTYPE') || rawText.includes('accounts.google.com')) {
      const err = new Error(
        'Google Apps Script ยังไม่ได้เปิดสิทธิ์เข้าถึง กรุณาตั้งค่า Deploy → Who has access เป็น "ทุกคน (Anyone)" แล้วนำ Web app URL ใหม่มาวาง'
      );
      err.status = 400;
      throw err;
    }
    const err = new Error(`Google Sheets ตอบกลับรูปแบบไม่ถูกต้อง (HTTP ${response.status})`);
    err.status = 502;
    throw err;
  }

  if (!response.ok || !parsed || parsed.ok === false) {
    const err = new Error((parsed && parsed.error) || `ส่งข้อมูลไป Google Sheets ไม่สำเร็จ (HTTP ${response.status})`);
    err.status = 502;
    throw err;
  }
  return parsed;
}

app.get(
  '/api/gsheet/config',
  requireAuth,
  handle(async (req, res) => res.json(await store.getGoogleSheetsConfig()))
);

app.put(
  '/api/gsheet/config',
  requireAuth,
  handle(async (req, res) => {
    const config = await store.saveGoogleSheetsConfig(req.body || {});
    res.json({ message: 'บันทึกการตั้งค่า Google Sheets เรียบร้อยแล้ว', config });
  })
);

app.post(
  '/api/gsheet/test',
  requireAuth,
  handle(async (req, res) => {
    const url = String((req.body && req.body.url) || '').trim();
    const result = await postToAppsScript(url, { action: 'ping' });
    res.json({
      ok: true,
      spreadsheetName: result.spreadsheetName || 'Google Sheets',
      spreadsheetUrl: result.spreadsheetUrl || ''
    });
  })
);

app.post(
  '/api/gsheet/sync',
  requireAuth,
  handle(async (req, res) => {
    const body = req.body || {};
    const kind = body.kind === 'billing' ? 'billing' : 'zone';
    const cfg = await store.getGoogleSheetsConfig();
    const targetUrl = kind === 'billing' ? cfg.billing_url : cfg.zone_url;
    if (!targetUrl) {
      return res.status(400).json({
        error: kind === 'billing' ? 'ยังไม่ได้ผูก Google Sheet ใบเบิกโครงการ' : 'ยังไม่ได้ผูก Google Sheet สำเนาโซน'
      });
    }

    const effectiveSite = kind === 'billing'
      ? String(cfg.billing_site || body.site_name || '').trim()
      : String(body.site_name || '').trim();

    const logs = await store.listLogs({
      date_from: String(body.date_from || '').trim(),
      date_to: String(body.date_to || '').trim(),
      site_name: effectiveSite,
      trip_type: kind === 'billing' ? '' : String(body.trip_type || '').trim(),
      search: kind === 'billing' ? '' : String(body.search || '').trim()
    });

    const result = await postToAppsScript(targetUrl, {
      action: 'sync',
      kind,
      siteName: effectiveSite,
      dateFrom: String(body.date_from || '').trim(),
      dateTo: String(body.date_to || '').trim(),
      issueDate: String(body.issue_date || '').trim(),
      logs
    });

    // อัปเดตชื่อไฟล์และลิงก์เปิดชีตอัตโนมัติหากมีการเปลี่ยนชื่อบน Google Drive
    const metaPatch = kind === 'billing'
      ? { billing_name: result.spreadsheetName || cfg.billing_name, billing_sheet_url: result.spreadsheetUrl || cfg.billing_sheet_url }
      : { zone_name: result.spreadsheetName || cfg.zone_name, zone_sheet_url: result.spreadsheetUrl || cfg.zone_sheet_url };
    const updatedConfig = await store.saveGoogleSheetsConfig(metaPatch);

    res.json({
      ok: true,
      kind,
      sheetCount: result.sheetCount || (Array.isArray(result.sheets) ? result.sheets.length : 0),
      sheets: result.sheets || [],
      logCount: logs.length,
      spreadsheetName: result.spreadsheetName || '',
      spreadsheetUrl: result.spreadsheetUrl || '',
      config: updatedConfig
    });
  })
);

app.use('/api', (req, res) => res.status(404).json({ error: 'ไม่พบ API ที่เรียกใช้' }));

// SPA fallback
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

if (require.main === module) {
  app.listen(PORT, () => {
    console.log('==================================================================');
    console.log('🚛 CKP Transport — ระบบจัดการบันทึกข้อมูลการเดินทางพนักงาน');
    console.log(`🌐 URL:      http://localhost:${PORT}`);
    console.log(`☁️  Storage:  ${isSupabaseConfigured ? 'Supabase PostgreSQL Cloud' : 'SQLite (local)'}`);
    console.log('==================================================================');
  });
}

module.exports = app;
