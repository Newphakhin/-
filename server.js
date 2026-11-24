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
  handle(async (req, res) => {
    const [config, history] = await Promise.all([
      store.getGoogleSheetsConfig(),
      store.listGoogleSheetsHistory()
    ]);
    res.json({
      ...config,
      history_count: history.length,
      last_history: history[0] || null
    });
  })
);

app.put(
  '/api/gsheet/config',
  requireAuth,
  handle(async (req, res) => {
    const body = req.body || {};
    const config = await store.saveGoogleSheetsConfig(body);
    let history = null;
    if (body.record_history) {
      const kind = body.history_kind === 'billing' ? 'billing' : body.history_kind === 'zone' ? 'zone' : 'general';
      const sheetName =
        kind === 'billing'
          ? config.billing_name || (config.billing_url ? 'Google Sheet ใบเบิก' : '')
          : kind === 'zone'
            ? config.zone_name || (config.zone_url ? 'Google Sheet สำเนาโซน' : '')
            : '';
      const sheetUrl = kind === 'billing' ? config.billing_sheet_url : kind === 'zone' ? config.zone_sheet_url : '';
      const added = await store.addGoogleSheetsHistory({
        kind,
        mode: 'config_saved',
        status: 'info',
        spreadsheet_name: sheetName,
        spreadsheet_url: sheetUrl,
        site_name: kind === 'billing' ? config.billing_site : '',
        triggered_by: (req.user && (req.user.full_name || req.user.username)) || 'ผู้ใช้ระบบ',
        note: String(body.history_note || 'บันทึกการตั้งค่า Google Sheets').trim()
      });
      history = added.history;
    }
    res.json({ message: 'บันทึกการตั้งค่า Google Sheets เรียบร้อยแล้ว', config, ...(history ? { history } : {}) });
  })
);

app.get(
  '/api/gsheet/history',
  requireAuth,
  handle(async (req, res) => {
    const history = await store.listGoogleSheetsHistory();
    res.json({ history });
  })
);

app.post(
  '/api/gsheet/history',
  requireAuth,
  handle(async (req, res) => {
    const body = req.body || {};
    const note = String(body.note || '').trim();
    if (!note) {
      return res.status(400).json({ error: 'กรุณากรอกรายละเอียดที่ต้องการบันทึก' });
    }
    const cfg = await store.getGoogleSheetsConfig();
    const kind = body.kind === 'billing' ? 'billing' : body.kind === 'zone' ? 'zone' : 'general';
    const spreadsheetName =
      String(body.spreadsheet_name || '').trim() ||
      (kind === 'billing' ? cfg.billing_name : kind === 'zone' ? cfg.zone_name : '');
    const spreadsheetUrl =
      String(body.spreadsheet_url || '').trim() ||
      (kind === 'billing' ? cfg.billing_sheet_url : kind === 'zone' ? cfg.zone_sheet_url : '');

    const { entry, history } = await store.addGoogleSheetsHistory({
      kind,
      mode: 'note',
      status: 'info',
      spreadsheet_name: spreadsheetName,
      spreadsheet_url: spreadsheetUrl,
      site_name: String(body.site_name || '').trim(),
      date_from: String(body.date_from || '').trim(),
      date_to: String(body.date_to || '').trim(),
      log_count: Number(body.log_count) || 0,
      triggered_by: (req.user && (req.user.full_name || req.user.username)) || 'ผู้ใช้ระบบ',
      note
    });
    res.status(201).json({ message: 'บันทึกประวัติการไหลข้อมูลเรียบร้อยแล้ว', entry, history });
  })
);

app.delete(
  '/api/gsheet/history/:id',
  requireAuth,
  handle(async (req, res) => {
    const { history } = await store.deleteGoogleSheetsHistoryItem(req.params.id);
    res.json({ message: 'ลบบันทึกประวัติเรียบร้อยแล้ว', history });
  })
);

app.delete(
  '/api/gsheet/history',
  requireAuth,
  requireAdmin,
  handle(async (req, res) => {
    const r = await store.clearGoogleSheetsHistory();
    res.json({ message: `ล้างประวัติการไหลข้อมูลทั้งหมด (${r.deletedCount} รายการ) เรียบร้อยแล้ว`, history: [] });
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
    const mode = body.mode === 'auto' ? 'auto' : 'manual';
    const triggeredBy = (req.user && (req.user.full_name || req.user.username)) || 'ผู้ใช้ระบบ';
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
    const dateFrom = String(body.date_from || '').trim();
    const dateTo = String(body.date_to || '').trim();

    const logs = await store.listLogs({
      date_from: dateFrom,
      date_to: dateTo,
      site_name: effectiveSite,
      trip_type: kind === 'billing' ? '' : String(body.trip_type || '').trim(),
      search: kind === 'billing' ? '' : String(body.search || '').trim()
    });

    const hostHeader = String(req.headers['x-forwarded-host'] || req.headers.host || '').trim();
    const protoHeader = String(req.headers['x-forwarded-proto'] || 'https').split(',')[0].trim();
    const logoUrl =
      hostHeader && !/^(localhost|127\.0\.0\.1)(:\d+)?$/i.test(hostHeader)
        ? `${protoHeader}://${hostHeader}/logo.jpg`
        : 'https://worker-dispatch-xi.vercel.app/logo.jpg';

    let result;
    try {
      result = await postToAppsScript(targetUrl, {
        action: 'sync',
        kind,
        siteName: effectiveSite,
        dateFrom,
        dateTo,
        issueDate: String(body.issue_date || '').trim(),
        logoUrl,
        logs
      });
    } catch (syncErr) {
      await store.addGoogleSheetsHistory({
        kind,
        mode,
        status: 'error',
        spreadsheet_name: kind === 'billing' ? cfg.billing_name : cfg.zone_name,
        spreadsheet_url: kind === 'billing' ? cfg.billing_sheet_url : cfg.zone_sheet_url,
        site_name: effectiveSite,
        date_from: dateFrom,
        date_to: dateTo,
        sheet_count: 0,
        log_count: logs.length,
        triggered_by: triggeredBy,
        error_message: syncErr.message || 'ส่งข้อมูลไม่สำเร็จ'
      });
      throw syncErr;
    }

    // อัปเดตชื่อไฟล์และลิงก์เปิดชีตอัตโนมัติหากมีการเปลี่ยนชื่อบน Google Drive
    const spreadsheetName = result.spreadsheetName || (kind === 'billing' ? cfg.billing_name : cfg.zone_name) || '';
    const spreadsheetUrl = result.spreadsheetUrl || (kind === 'billing' ? cfg.billing_sheet_url : cfg.zone_sheet_url) || '';
    const sheetCount = result.sheetCount || (Array.isArray(result.sheets) ? result.sheets.length : 0);

    const metaPatch = kind === 'billing'
      ? { billing_name: spreadsheetName, billing_sheet_url: spreadsheetUrl }
      : { zone_name: spreadsheetName, zone_sheet_url: spreadsheetUrl };
    const updatedConfig = await store.saveGoogleSheetsConfig(metaPatch);

    const { entry: historyEntry, history } = await store.addGoogleSheetsHistory({
      kind,
      mode,
      status: 'success',
      spreadsheet_name: spreadsheetName,
      spreadsheet_url: spreadsheetUrl,
      site_name: effectiveSite,
      date_from: dateFrom,
      date_to: dateTo,
      sheet_count: sheetCount,
      log_count: logs.length,
      triggered_by: triggeredBy,
      note: String(body.note || '').trim()
    });

    res.json({
      ok: true,
      kind,
      sheetCount,
      sheets: result.sheets || [],
      logCount: logs.length,
      spreadsheetName,
      spreadsheetUrl,
      config: updatedConfig,
      historyEntry,
      historyCount: history.length
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
