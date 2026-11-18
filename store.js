const {
  db,
  supabase,
  isSupabaseConfigured,
  DB_PATH,
  hashPassword,
  verifyPassword,
  seedDatabase,
  syncJsonFallbackFile,
  getSystemSettingsMap: getLocalSettingsMap,
  updateSystemSettings: updateLocalSettings,
  getStorageQuotaStatus: getLocalQuotaStatus,
  INITIAL_VEHICLES
} = require('./db');
const fs = require('node:fs');

function getTodayDateString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// ============================================================================
// USERS & AUTH
// ============================================================================

async function findUserByUsername(username) {
  if (isSupabaseConfigured && supabase) {
    const { data, error } = await supabase
      .from('users')
      .select('id, username, password, full_name, role')
      .eq('username', username)
      .maybeSingle();
    if (error) throw new Error(error.message);
    return data || null;
  }
  return (
    db
      .prepare('SELECT id, username, password, full_name, role FROM users WHERE username = ?')
      .get(username) || null
  );
}

async function findUserById(id) {
  if (isSupabaseConfigured && supabase) {
    const { data, error } = await supabase
      .from('users')
      .select('id, username, full_name, role')
      .eq('id', id)
      .maybeSingle();
    if (error) throw new Error(error.message);
    return data || null;
  }
  return (
    db.prepare('SELECT id, username, full_name, role FROM users WHERE id = ?').get(id) || null
  );
}

// ============================================================================
// SYSTEM SETTINGS & STORAGE QUOTA
// ============================================================================

async function getSettingsMap() {
  const defaults = {
    cartrack_api_url: 'https://fleetapi-th.cartrack.com/rest/vehicles/status',
    cartrack_username: 'DEMO_FLEET_TH',
    cartrack_api_key: '',
    cartrack_mock_mode: 'true',
    max_storage_limit: '100'
  };

  if (isSupabaseConfigured && supabase) {
    const { data, error } = await supabase.from('system_settings').select('key, value');
    if (error) throw new Error(error.message);
    for (const r of data || []) {
      defaults[r.key] = r.value;
    }
    return defaults;
  }

  return getLocalSettingsMap();
}

async function saveSettingsMap(newSettings) {
  if (isSupabaseConfigured && supabase) {
    const rows = Object.entries(newSettings)
      .filter(([, v]) => v !== undefined && v !== null)
      .map(([key, value]) => ({ key, value: String(value) }));
    if (rows.length > 0) {
      const { error } = await supabase.from('system_settings').upsert(rows, { onConflict: 'key' });
      if (error) throw new Error(error.message);
    }
    return getSettingsMap();
  }
  return updateLocalSettings(newSettings);
}

async function getQuotaStatus() {
  if (isSupabaseConfigured && supabase) {
    const settings = await getSettingsMap();
    const maxLimit = Math.max(1, Number(settings.max_storage_limit) || 100);
    const { count, error } = await supabase
      .from('travel_logs')
      .select('*', { count: 'exact', head: true });
    if (error) throw new Error(error.message);
    const currentCount = Number(count || 0);
    const usagePercent = Math.min(100, Math.round((currentCount / maxLimit) * 100));

    let statusLevel = 'normal';
    if (currentCount >= maxLimit || usagePercent >= 100) statusLevel = 'full';
    else if (usagePercent >= 80) statusLevel = 'warning';

    return {
      currentCount,
      maxLimit,
      usagePercent,
      statusLevel,
      isFull: currentCount >= maxLimit,
      isWarning: usagePercent >= 80 && currentCount < maxLimit,
      dbSizeBytes: currentCount * 512,
      supabaseFreeTierBytes: 500 * 1024 * 1024,
      storageMode: 'supabase_cloud'
    };
  }

  return getLocalQuotaStatus();
}

// ============================================================================
// MASTER WORKERS
// ============================================================================

async function listWorkers(activeOnly = false) {
  if (isSupabaseConfigured && supabase) {
    let query = supabase
      .from('master_workers')
      .select('id, name, default_route, active')
      .order('id', { ascending: true });
    if (activeOnly) query = query.eq('active', true);
    const { data, error } = await query;
    if (error) throw new Error(error.message);
    return (data || []).map((r) => ({
      id: r.id,
      name: r.name,
      default_route: r.default_route || '',
      active: Boolean(r.active)
    }));
  }

  const sql = activeOnly
    ? 'SELECT id, name, default_route, active FROM master_workers WHERE active = 1 ORDER BY id ASC'
    : 'SELECT id, name, default_route, active FROM master_workers ORDER BY id ASC';
  return db
    .prepare(sql)
    .all()
    .map((r) => ({
      id: r.id,
      name: r.name,
      default_route: r.default_route || '',
      active: Boolean(r.active)
    }));
}

async function insertWithSequenceFix(table, row, selectCols) {
  let { data, error } = await supabase.from(table).insert(row).select(selectCols).single();
  if (error && error.code === '23505' && String(error.message).includes('_pkey')) {
    const { data: maxRow } = await supabase
      .from(table)
      .select('id')
      .order('id', { ascending: false })
      .limit(1)
      .maybeSingle();
    const nextId = (maxRow?.id || 0) + 1;
    const retry = await supabase
      .from(table)
      .insert({ ...row, id: nextId })
      .select(selectCols)
      .single();
    data = retry.data;
    error = retry.error;
  }
  if (error) throw new Error(error.message);
  return data;
}

async function createWorker({ name, default_route, active }) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('master_workers')
      .select('id')
      .eq('name', name)
      .maybeSingle();
    if (existing) {
      const err = new Error(`ชื่อคนงาน "${name}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
    }

    const data = await insertWithSequenceFix(
      'master_workers',
      { name, default_route: default_route || '', active: Boolean(active) },
      'id, name, default_route, active'
    );
    return {
      id: data.id,
      name: data.name,
      default_route: data.default_route || '',
      active: Boolean(data.active)
    };
  }

  const existing = db.prepare('SELECT id FROM master_workers WHERE name = ?').get(name);
  if (existing) {
    const err = new Error(`ชื่อคนงาน "${name}" มีอยู่ในระบบแล้ว`);
    err.status = 409;
    throw err;
  }
  const result = db
    .prepare('INSERT INTO master_workers (name, default_route, active) VALUES (?, ?, ?)')
    .run(name, default_route || '', active ? 1 : 0);
  syncJsonFallbackFile();
  const created = db
    .prepare('SELECT id, name, default_route, active FROM master_workers WHERE id = ?')
    .get(result.lastInsertRowid);
  return {
    id: created.id,
    name: created.name,
    default_route: created.default_route || '',
    active: Boolean(created.active)
  };
}

async function updateWorker(id, { name, default_route, active }, userRole) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing, error: fetchErr } = await supabase
      .from('master_workers')
      .select('id, name, default_route, active')
      .eq('id', id)
      .maybeSingle();
    if (fetchErr) throw new Error(fetchErr.message);
    if (!existing) {
      const err = new Error('ไม่พบข้อมูลคนงานที่ต้องการแก้ไข');
      err.status = 404;
      throw err;
    }

    const isRenaming = name !== undefined && name !== existing.name;
    const isTogglingActive = active !== undefined && Boolean(active) !== Boolean(existing.active);
    if ((isRenaming || isTogglingActive) && userRole !== 'admin') {
      const err = new Error('เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่สามารถเปลี่ยนชื่อหรือสถานะคนงานได้');
      err.status = 403;
      throw err;
    }

    const newName = name !== undefined ? name : existing.name;
    const newDefaultRoute =
      default_route !== undefined ? default_route : existing.default_route || '';
    const newActive = active !== undefined ? Boolean(active) : Boolean(existing.active);

    if (newName !== existing.name) {
      const { data: dup } = await supabase
        .from('master_workers')
        .select('id')
        .eq('name', newName)
        .neq('id', id)
        .maybeSingle();
      if (dup) {
        const err = new Error(`ชื่อคนงาน "${newName}" มีอยู่ในระบบแล้ว`);
        err.status = 409;
        throw err;
      }
    }

    const { data: updated, error: updErr } = await supabase
      .from('master_workers')
      .update({ name: newName, default_route: newDefaultRoute, active: newActive })
      .eq('id', id)
      .select('id, name, default_route, active')
      .single();
    if (updErr) throw new Error(updErr.message);

    let cascadedLogsCount = 0;
    if (newName !== existing.name) {
      const allLogs = await listLogs();
      for (const log of allLogs) {
        if ((log.worker_name || '') === existing.name) {
          await updateLog(log.log_id || log.worker_name, { worker_name: newName });
          cascadedLogsCount += 1;
        }
      }
    }

    return {
      id: updated.id,
      name: updated.name,
      default_route: updated.default_route || '',
      active: Boolean(updated.active),
      cascadedLogsCount
    };
  }

  const existing = db
    .prepare('SELECT id, name, default_route, active FROM master_workers WHERE id = ?')
    .get(id);
  if (!existing) {
    const err = new Error('ไม่พบข้อมูลคนงานที่ต้องการแก้ไข');
    err.status = 404;
    throw err;
  }

  const isRenaming = name !== undefined && name !== existing.name;
  const isTogglingActive = active !== undefined && Boolean(active) !== Boolean(existing.active);
  if ((isRenaming || isTogglingActive) && userRole !== 'admin') {
    const err = new Error('เฉพาะผู้ดูแลระบบ (Admin) เท่านั้นที่สามารถเปลี่ยนชื่อหรือสถานะคนงานได้');
    err.status = 403;
    throw err;
  }

  const newName = name !== undefined ? name : existing.name;
  const newDefaultRoute =
    default_route !== undefined ? default_route : existing.default_route || '';
  const newActive = active !== undefined ? (active ? 1 : 0) : existing.active ? 1 : 0;

  if (newName !== existing.name) {
    const dup = db
      .prepare('SELECT id FROM master_workers WHERE name = ? AND id != ?')
      .get(newName, id);
    if (dup) {
      const err = new Error(`ชื่อคนงาน "${newName}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
    }
  }

  let cascadedLogsCount = 0;
  db.exec('BEGIN');
  try {
    db.prepare('UPDATE master_workers SET name = ?, default_route = ?, active = ? WHERE id = ?').run(
      newName,
      newDefaultRoute,
      newActive,
      id
    );
    if (newName !== existing.name) {
      const matchingLogs = db
        .prepare("SELECT worker_name FROM travel_logs WHERE worker_name = ? OR worker_name LIKE ?")
        .all(existing.name, `${existing.name}__#%`);
      const updStmt = db.prepare(
        "UPDATE travel_logs SET worker_name = ?, updated_at = datetime('now', 'localtime') WHERE worker_name = ?"
      );
      for (const m of matchingLogs) {
        const { suffix } = splitWorkerLogKey(m.worker_name);
        const nextKey = suffix ? `${newName}${suffix}` : newName;
        updStmt.run(nextKey, m.worker_name);
        cascadedLogsCount += 1;
      }
    }
    db.exec('COMMIT');
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }
  syncJsonFallbackFile();

  const updated = db
    .prepare('SELECT id, name, default_route, active FROM master_workers WHERE id = ?')
    .get(id);
  return {
    id: updated.id,
    name: updated.name,
    default_route: updated.default_route || '',
    active: Boolean(updated.active),
    cascadedLogsCount
  };
}

async function deleteWorker(id) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('master_workers')
      .select('id, name')
      .eq('id', id)
      .maybeSingle();
    if (!existing) {
      const err = new Error('ไม่พบรายชื่อคนงานที่ต้องการลบ');
      err.status = 404;
      throw err;
    }
    const { error } = await supabase.from('master_workers').delete().eq('id', id);
    if (error) throw new Error(error.message);
    return existing;
  }

  const existing = db.prepare('SELECT id, name FROM master_workers WHERE id = ?').get(id);
  if (!existing) {
    const err = new Error('ไม่พบรายชื่อคนงานที่ต้องการลบ');
    err.status = 404;
    throw err;
  }
  db.prepare('DELETE FROM master_workers WHERE id = ?').run(id);
  syncJsonFallbackFile();
  return existing;
}

// ============================================================================
// MASTER ROUTES
// ============================================================================

async function listRoutes(activeOnly = false) {
  if (isSupabaseConfigured && supabase) {
    let query = supabase
      .from('master_routes')
      .select('id, name, active')
      .order('id', { ascending: true });
    if (activeOnly) query = query.eq('active', true);
    const { data, error } = await query;
    if (error) throw new Error(error.message);
    return (data || []).map((r) => ({ id: r.id, name: r.name, active: Boolean(r.active) }));
  }

  const sql = activeOnly
    ? 'SELECT id, name, active FROM master_routes WHERE active = 1 ORDER BY id ASC'
    : 'SELECT id, name, active FROM master_routes ORDER BY id ASC';
  return db
    .prepare(sql)
    .all()
    .map((r) => ({ id: r.id, name: r.name, active: Boolean(r.active) }));
}

async function createRoute({ name, active, reactivateIfInactive }) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('master_routes')
      .select('id, name, active')
      .eq('name', name)
      .maybeSingle();
    if (existing) {
      if (!existing.active && reactivateIfInactive) {
        await supabase.from('master_routes').update({ active: true }).eq('id', existing.id);
        return { id: existing.id, name: existing.name, active: true, reactivated: true };
      }
      const err = new Error(`เส้นทาง/ที่หมาย "${name}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
    }

    const data = await insertWithSequenceFix(
      'master_routes',
      { name, active: Boolean(active) },
      'id, name, active'
    );
    return { id: data.id, name: data.name, active: Boolean(data.active) };
  }

  const existing = db.prepare('SELECT id, name, active FROM master_routes WHERE name = ?').get(name);
  if (existing) {
    if (!existing.active && reactivateIfInactive) {
      db.prepare('UPDATE master_routes SET active = 1 WHERE id = ?').run(existing.id);
      syncJsonFallbackFile();
      return { id: existing.id, name: existing.name, active: true, reactivated: true };
    }
    const err = new Error(`เส้นทาง/ที่หมาย "${name}" มีอยู่ในระบบแล้ว`);
    err.status = 409;
    throw err;
  }
  const result = db
    .prepare('INSERT INTO master_routes (name, active) VALUES (?, ?)')
    .run(name, active ? 1 : 0);
  syncJsonFallbackFile();
  const created = db
    .prepare('SELECT id, name, active FROM master_routes WHERE id = ?')
    .get(result.lastInsertRowid);
  return { id: created.id, name: created.name, active: Boolean(created.active) };
}

async function updateRoute(id, { name, active }) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing, error: fetchErr } = await supabase
      .from('master_routes')
      .select('id, name, active')
      .eq('id', id)
      .maybeSingle();
    if (fetchErr) throw new Error(fetchErr.message);
    if (!existing) {
      const err = new Error('ไม่พบเส้นทาง/ที่หมายที่ต้องการแก้ไข');
      err.status = 404;
      throw err;
    }

    const newName = name !== undefined ? name : existing.name;
    const newActive = active !== undefined ? Boolean(active) : Boolean(existing.active);

    if (newName !== existing.name) {
      const { data: dup } = await supabase
        .from('master_routes')
        .select('id')
        .eq('name', newName)
        .neq('id', id)
        .maybeSingle();
      if (dup) {
        const err = new Error(`เส้นทาง/ที่หมาย "${newName}" มีอยู่ในระบบแล้ว`);
        err.status = 409;
        throw err;
      }
    }

    const { data: updated, error: updErr } = await supabase
      .from('master_routes')
      .update({ name: newName, active: newActive })
      .eq('id', id)
      .select('id, name, active')
      .single();
    if (updErr) throw new Error(updErr.message);

    let cascadedLogsCount = 0;
    let cascadedWorkersCount = 0;
    if (newName !== existing.name) {
      const { data: lUpd } = await supabase
        .from('travel_logs')
        .update({ route: newName, updated_at: new Date().toISOString() })
        .eq('route', existing.name)
        .select('worker_name');
      cascadedLogsCount = (lUpd || []).length;

      const { data: wUpd } = await supabase
        .from('master_workers')
        .update({ default_route: newName })
        .eq('default_route', existing.name)
        .select('id');
      cascadedWorkersCount = (wUpd || []).length;
    }

    return {
      id: updated.id,
      name: updated.name,
      active: Boolean(updated.active),
      cascadedLogsCount,
      cascadedWorkersCount
    };
  }

  const existing = db.prepare('SELECT id, name, active FROM master_routes WHERE id = ?').get(id);
  if (!existing) {
    const err = new Error('ไม่พบเส้นทาง/ที่หมายที่ต้องการแก้ไข');
    err.status = 404;
    throw err;
  }
  const newName = name !== undefined ? name : existing.name;
  const newActive = active !== undefined ? (active ? 1 : 0) : existing.active ? 1 : 0;

  if (newName !== existing.name) {
    const dup = db
      .prepare('SELECT id FROM master_routes WHERE name = ? AND id != ?')
      .get(newName, id);
    if (dup) {
      const err = new Error(`เส้นทาง/ที่หมาย "${newName}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
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
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }
  syncJsonFallbackFile();
  const updated = db.prepare('SELECT id, name, active FROM master_routes WHERE id = ?').get(id);
  return {
    id: updated.id,
    name: updated.name,
    active: Boolean(updated.active),
    cascadedLogsCount,
    cascadedWorkersCount
  };
}

async function deleteRoute(id) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('master_routes')
      .select('id, name')
      .eq('id', id)
      .maybeSingle();
    if (!existing) {
      const err = new Error('ไม่พบเส้นทาง/ที่หมายที่ต้องการลบ');
      err.status = 404;
      throw err;
    }
    const { error } = await supabase.from('master_routes').delete().eq('id', id);
    if (error) throw new Error(error.message);
    await supabase
      .from('master_workers')
      .update({ default_route: '' })
      .eq('default_route', existing.name);
    return existing;
  }

  const existing = db.prepare('SELECT id, name FROM master_routes WHERE id = ?').get(id);
  if (!existing) {
    const err = new Error('ไม่พบเส้นทาง/ที่หมายที่ต้องการลบ');
    err.status = 404;
    throw err;
  }
  db.exec('BEGIN');
  try {
    db.prepare('DELETE FROM master_routes WHERE id = ?').run(id);
    db.prepare("UPDATE master_workers SET default_route = '' WHERE default_route = ?").run(
      existing.name
    );
    db.exec('COMMIT');
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }
  syncJsonFallbackFile();
  return existing;
}

// ============================================================================
// MASTER VEHICLES
// ============================================================================

async function listVehicles(activeOnly = false) {
  if (isSupabaseConfigured && supabase) {
    let query = supabase
      .from('master_vehicles')
      .select('id, plate, active')
      .order('id', { ascending: true });
    if (activeOnly) query = query.eq('active', true);
    const { data, error } = await query;
    if (error) throw new Error(error.message);
    return (data || []).map((r) => ({ id: r.id, plate: r.plate, active: Boolean(r.active) }));
  }

  const sql = activeOnly
    ? 'SELECT id, plate, active FROM master_vehicles WHERE active = 1 ORDER BY id ASC'
    : 'SELECT id, plate, active FROM master_vehicles ORDER BY id ASC';
  return db
    .prepare(sql)
    .all()
    .map((r) => ({ id: r.id, plate: r.plate, active: Boolean(r.active) }));
}

async function createVehicle({ plate, active }) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('master_vehicles')
      .select('id')
      .eq('plate', plate)
      .maybeSingle();
    if (existing) {
      const err = new Error(`ทะเบียนรถ "${plate}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
    }

    const data = await insertWithSequenceFix(
      'master_vehicles',
      { plate, active: Boolean(active) },
      'id, plate, active'
    );
    return { id: data.id, plate: data.plate, active: Boolean(data.active) };
  }

  const existing = db.prepare('SELECT id FROM master_vehicles WHERE plate = ?').get(plate);
  if (existing) {
    const err = new Error(`ทะเบียนรถ "${plate}" มีอยู่ในระบบแล้ว`);
    err.status = 409;
    throw err;
  }
  const result = db
    .prepare('INSERT INTO master_vehicles (plate, active) VALUES (?, ?)')
    .run(plate, active ? 1 : 0);
  syncJsonFallbackFile();
  const created = db
    .prepare('SELECT id, plate, active FROM master_vehicles WHERE id = ?')
    .get(result.lastInsertRowid);
  return { id: created.id, plate: created.plate, active: Boolean(created.active) };
}

async function updateVehicle(id, { plate, active }) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing, error: fetchErr } = await supabase
      .from('master_vehicles')
      .select('id, plate, active')
      .eq('id', id)
      .maybeSingle();
    if (fetchErr) throw new Error(fetchErr.message);
    if (!existing) {
      const err = new Error('ไม่พบทะเบียนรถที่ต้องการแก้ไข');
      err.status = 404;
      throw err;
    }

    const newPlate = plate !== undefined ? plate : existing.plate;
    const newActive = active !== undefined ? Boolean(active) : Boolean(existing.active);

    if (newPlate !== existing.plate) {
      const { data: dup } = await supabase
        .from('master_vehicles')
        .select('id')
        .eq('plate', newPlate)
        .neq('id', id)
        .maybeSingle();
      if (dup) {
        const err = new Error(`ทะเบียนรถ "${newPlate}" มีอยู่ในระบบแล้ว`);
        err.status = 409;
        throw err;
      }
    }

    const { data: updated, error: updErr } = await supabase
      .from('master_vehicles')
      .update({ plate: newPlate, active: newActive })
      .eq('id', id)
      .select('id, plate, active')
      .single();
    if (updErr) throw new Error(updErr.message);

    let cascadedLogsCount = 0;
    if (newPlate !== existing.plate) {
      const { data: lUpd } = await supabase
        .from('travel_logs')
        .update({ vehicle_plate: newPlate, updated_at: new Date().toISOString() })
        .eq('vehicle_plate', existing.plate)
        .select('worker_name');
      cascadedLogsCount = (lUpd || []).length;
    }

    return {
      id: updated.id,
      plate: updated.plate,
      active: Boolean(updated.active),
      cascadedLogsCount
    };
  }

  const existing = db.prepare('SELECT id, plate, active FROM master_vehicles WHERE id = ?').get(id);
  if (!existing) {
    const err = new Error('ไม่พบทะเบียนรถที่ต้องการแก้ไข');
    err.status = 404;
    throw err;
  }
  const newPlate = plate !== undefined ? plate : existing.plate;
  const newActive = active !== undefined ? (active ? 1 : 0) : existing.active ? 1 : 0;

  if (newPlate !== existing.plate) {
    const dup = db
      .prepare('SELECT id FROM master_vehicles WHERE plate = ? AND id != ?')
      .get(newPlate, id);
    if (dup) {
      const err = new Error(`ทะเบียนรถ "${newPlate}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
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
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }
  syncJsonFallbackFile();
  const updated = db.prepare('SELECT id, plate, active FROM master_vehicles WHERE id = ?').get(id);
  return {
    id: updated.id,
    plate: updated.plate,
    active: Boolean(updated.active),
    cascadedLogsCount
  };
}

async function deleteVehicle(id) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('master_vehicles')
      .select('id, plate')
      .eq('id', id)
      .maybeSingle();
    if (!existing) {
      const err = new Error('ไม่พบทะเบียนรถที่ต้องการลบ');
      err.status = 404;
      throw err;
    }
    const { error } = await supabase.from('master_vehicles').delete().eq('id', id);
    if (error) throw new Error(error.message);
    return existing;
  }

  const existing = db.prepare('SELECT id, plate FROM master_vehicles WHERE id = ?').get(id);
  if (!existing) {
    const err = new Error('ไม่พบทะเบียนรถที่ต้องการลบ');
    err.status = 404;
    throw err;
  }
  db.prepare('DELETE FROM master_vehicles WHERE id = ?').run(id);
  syncJsonFallbackFile();
  return existing;
}

// ============================================================================
// MASTER SITES (ไซต์งาน / โครงการ - หัวตาราง เช่น ไลฟ์รัชดา-พระราม9)
// ============================================================================

async function getSupabaseSitesArray() {
  const { data, error } = await supabase
    .from('system_settings')
    .select('value')
    .eq('key', 'master_sites_json')
    .maybeSingle();
  if (error) throw new Error(error.message);
  if (!data || !data.value) return [];
  try {
    const parsed = JSON.parse(data.value);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

async function saveSupabaseSitesArray(sitesArray) {
  const { error } = await supabase
    .from('system_settings')
    .upsert([{ key: 'master_sites_json', value: JSON.stringify(sitesArray) }], {
      onConflict: 'key'
    });
  if (error) throw new Error(error.message);
}

async function listSites(activeOnly = false) {
  if (isSupabaseConfigured && supabase) {
    const sites = await getSupabaseSitesArray();
    const mapped = sites.map((s) => ({
      id: Number(s.id),
      name: String(s.name || ''),
      active: Boolean(s.active)
    }));
    return activeOnly ? mapped.filter((s) => s.active) : mapped;
  }

  const sql = activeOnly
    ? 'SELECT id, name, active FROM master_sites WHERE active = 1 ORDER BY id ASC'
    : 'SELECT id, name, active FROM master_sites ORDER BY id ASC';
  return db
    .prepare(sql)
    .all()
    .map((r) => ({ id: r.id, name: r.name, active: Boolean(r.active) }));
}

async function createSite({ name, active, reactivateIfInactive }) {
  if (isSupabaseConfigured && supabase) {
    const sites = await getSupabaseSitesArray();
    const existing = sites.find((s) => String(s.name).toLowerCase() === name.toLowerCase());
    if (existing) {
      if (!existing.active && reactivateIfInactive) {
        existing.active = true;
        await saveSupabaseSitesArray(sites);
        return { id: existing.id, name: existing.name, active: true, reactivated: true };
      }
      const err = new Error(`ไซต์งาน/โครงการ "${name}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
    }
    const nextId = sites.reduce((max, s) => Math.max(max, Number(s.id) || 0), 0) + 1;
    const newItem = { id: nextId, name, active: Boolean(active) };
    sites.push(newItem);
    await saveSupabaseSitesArray(sites);
    return newItem;
  }

  const existing = db.prepare('SELECT id, name, active FROM master_sites WHERE name = ?').get(name);
  if (existing) {
    if (!existing.active && reactivateIfInactive) {
      db.prepare('UPDATE master_sites SET active = 1 WHERE id = ?').run(existing.id);
      syncJsonFallbackFile();
      return { id: existing.id, name: existing.name, active: true, reactivated: true };
    }
    const err = new Error(`ไซต์งาน/โครงการ "${name}" มีอยู่ในระบบแล้ว`);
    err.status = 409;
    throw err;
  }
  const result = db
    .prepare('INSERT INTO master_sites (name, active) VALUES (?, ?)')
    .run(name, active ? 1 : 0);
  syncJsonFallbackFile();
  const created = db
    .prepare('SELECT id, name, active FROM master_sites WHERE id = ?')
    .get(result.lastInsertRowid);
  return { id: created.id, name: created.name, active: Boolean(created.active) };
}

async function updateSite(id, { name, active }) {
  if (isSupabaseConfigured && supabase) {
    const sites = await getSupabaseSitesArray();
    const idx = sites.findIndex((s) => Number(s.id) === Number(id));
    if (idx === -1) {
      const err = new Error('ไม่พบไซต์งาน/โครงการที่ต้องการแก้ไข');
      err.status = 404;
      throw err;
    }
    const existing = sites[idx];
    const newName = name !== undefined ? name : existing.name;
    const newActive = active !== undefined ? Boolean(active) : Boolean(existing.active);

    if (newName !== existing.name) {
      const dup = sites.find(
        (s) => Number(s.id) !== Number(id) && String(s.name).toLowerCase() === newName.toLowerCase()
      );
      if (dup) {
        const err = new Error(`ไซต์งาน/โครงการ "${newName}" มีอยู่ในระบบแล้ว`);
        err.status = 409;
        throw err;
      }
    }

    sites[idx] = { id: Number(existing.id), name: newName, active: newActive };
    await saveSupabaseSitesArray(sites);

    let cascadedLogsCount = 0;
    if (newName !== existing.name) {
      const allLogs = await listLogs();
      for (const log of allLogs) {
        if ((log.site_name || '') === existing.name) {
          await updateLog(log.log_id || log.worker_name, { site_name: newName });
          cascadedLogsCount += 1;
        }
      }
    }

    return {
      id: Number(existing.id),
      name: newName,
      active: newActive,
      cascadedLogsCount
    };
  }

  const existing = db.prepare('SELECT id, name, active FROM master_sites WHERE id = ?').get(id);
  if (!existing) {
    const err = new Error('ไม่พบไซต์งาน/โครงการที่ต้องการแก้ไข');
    err.status = 404;
    throw err;
  }
  const newName = name !== undefined ? name : existing.name;
  const newActive = active !== undefined ? (active ? 1 : 0) : existing.active ? 1 : 0;

  if (newName !== existing.name) {
    const dup = db
      .prepare('SELECT id FROM master_sites WHERE name = ? AND id != ?')
      .get(newName, id);
    if (dup) {
      const err = new Error(`ไซต์งาน/โครงการ "${newName}" มีอยู่ในระบบแล้ว`);
      err.status = 409;
      throw err;
    }
  }

  let cascadedLogsCount = 0;
  db.exec('BEGIN');
  try {
    db.prepare('UPDATE master_sites SET name = ?, active = ? WHERE id = ?').run(
      newName,
      newActive,
      id
    );
    if (newName !== existing.name) {
      const logRes = db
        .prepare(
          "UPDATE travel_logs SET site_name = ?, updated_at = datetime('now', 'localtime') WHERE site_name = ?"
        )
        .run(newName, existing.name);
      cascadedLogsCount = Number(logRes.changes || 0);
    }
    db.exec('COMMIT');
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }
  syncJsonFallbackFile();
  const updated = db.prepare('SELECT id, name, active FROM master_sites WHERE id = ?').get(id);
  return {
    id: updated.id,
    name: updated.name,
    active: Boolean(updated.active),
    cascadedLogsCount
  };
}

async function deleteSite(id) {
  if (isSupabaseConfigured && supabase) {
    const sites = await getSupabaseSitesArray();
    const idx = sites.findIndex((s) => Number(s.id) === Number(id));
    if (idx === -1) {
      const err = new Error('ไม่พบไซต์งาน/โครงการที่ต้องการลบ');
      err.status = 404;
      throw err;
    }
    const [deleted] = sites.splice(idx, 1);
    await saveSupabaseSitesArray(sites);
    return deleted;
  }

  const existing = db.prepare('SELECT id, name FROM master_sites WHERE id = ?').get(id);
  if (!existing) {
    const err = new Error('ไม่พบไซต์งาน/โครงการที่ต้องการลบ');
    err.status = 404;
    throw err;
  }
  db.prepare('DELETE FROM master_sites WHERE id = ?').run(id);
  syncJsonFallbackFile();
  return existing;
}

// ============================================================================
// TRAVEL LOGS
// ============================================================================

function normalizePassengerCount(val) {
  if (val === null || val === undefined || String(val).trim() === '') {
    return null;
  }
  const num = Number.parseInt(String(val).trim(), 10);
  if (Number.isNaN(num) || num < 0) return null;
  return num;
}

function encodeExtraFieldsToGpsStatus(passenger_count, site_name, note) {
  const cleanSite = String(site_name || '').trim();
  const cleanNote = String(note || '').trim();
  if (passenger_count === null && !cleanSite && !cleanNote) {
    return '';
  }
  return JSON.stringify({
    passenger_count,
    site_name: cleanSite,
    note: cleanNote
  });
}

function splitWorkerLogKey(rawKey) {
  const str = String(rawKey || '');
  const sepIdx = str.indexOf('__#');
  if (sepIdx === -1) {
    return {
      log_id: str,
      worker_name: str,
      suffix: ''
    };
  }
  return {
    log_id: str,
    worker_name: str.slice(0, sepIdx),
    suffix: str.slice(sepIdx)
  };
}

function withResolvedPassengerCount(row) {
  if (!row) return row;
  const { log_id, worker_name } = splitWorkerLogKey(row.worker_name);
  let count =
    row.passenger_count !== undefined && row.passenger_count !== null
      ? normalizePassengerCount(row.passenger_count)
      : null;
  let siteName =
    row.site_name !== undefined && row.site_name !== null ? String(row.site_name).trim() : '';
  let noteText =
    row.note !== undefined && row.note !== null ? String(row.note).trim() : '';

  if (row.gps_status) {
    try {
      const parsed = JSON.parse(row.gps_status);
      if (parsed && typeof parsed === 'object') {
        if (count === null && parsed.passenger_count !== undefined && parsed.passenger_count !== null) {
          count = normalizePassengerCount(parsed.passenger_count);
        }
        if (!siteName && parsed.site_name) {
          siteName = String(parsed.site_name).trim();
        }
        if (!noteText && parsed.note) {
          noteText = String(parsed.note).trim();
        }
      }
    } catch {
      // ignore non-JSON gps_status
    }
  }

  return {
    ...row,
    log_id,
    worker_name,
    site_name: siteName,
    passenger_count: count,
    note: noteText
  };
}

function isMissingExtraColumnError(error) {
  if (!error) return false;
  const msg = String(error.message || '').toLowerCase();
  return (
    error.code === 'PGRST204' ||
    error.code === '42703' ||
    msg.includes('passenger_count') ||
    msg.includes('site_name')
  );
}

async function listLogs(dateFilter = '', searchFilter = '') {
  if (dateFilter && typeof dateFilter === 'object') {
    searchFilter = dateFilter.search || '';
    dateFilter = dateFilter.date || '';
  }
  let rows = [];
  if (isSupabaseConfigured && supabase) {
    const { data, error } = await supabase
      .from('travel_logs')
      .select('*')
      .order('travel_date', { ascending: false })
      .order('departure_time', { ascending: true })
      .order('worker_name', { ascending: true });
    if (error) throw new Error(error.message);
    rows = (data || []).map(withResolvedPassengerCount);
  } else {
    rows = db
      .prepare(
        `SELECT worker_name, site_name, travel_date, route, departure_time, arrival_time, vehicle_plate, passenger_count, gps_status, updated_at
         FROM travel_logs
         ORDER BY travel_date DESC, departure_time ASC, worker_name ASC`
      )
      .all()
      .map(withResolvedPassengerCount);
  }

  if (dateFilter) {
    rows = rows.filter((r) => (r.travel_date || '') === dateFilter);
  }

  if (searchFilter) {
    const q = searchFilter.toLowerCase();
    rows = rows.filter((r) => {
      const combined = [
        r.worker_name || '',
        r.site_name || '',
        r.travel_date || '',
        r.route || '',
        r.vehicle_plate || '',
        r.passenger_count !== null && r.passenger_count !== undefined ? `${r.passenger_count} คน` : '',
        r.note || '',
        r.departure_time || '',
        r.arrival_time || ''
      ]
        .join(' ')
        .toLowerCase();
      return combined.includes(q);
    });
  }

  return rows;
}

async function createLog(payload) {
  const {
    worker_name,
    site_name,
    travel_date,
    route,
    departure_time,
    arrival_time,
    vehicle_plate,
    note
  } = payload;
  const cleanWorkerName = String(worker_name || '').trim();
  const cleanSiteName = String(site_name || '').trim();
  const cleanNote = String(note || '').trim();
  const passenger_count = normalizePassengerCount(payload.passenger_count);
  const encodedExtra = encodeExtraFieldsToGpsStatus(passenger_count, cleanSiteName, cleanNote);
  const uniqueLogKey = `${cleanWorkerName}__#${Date.now()}_${Math.floor(Math.random() * 10000)}`;

  if (isSupabaseConfigured && supabase) {
    const quota = await getQuotaStatus();
    if (quota.isFull) {
      const err = new Error(
        `พื้นที่เก็บข้อมูลเต็ม (${quota.currentCount}/${quota.maxLimit} รายการ - 100%) ไม่สามารถบันทึกข้อมูลใหม่ได้ชั่วคราว กรุณาส่งออกไฟล์สำรองข้อมูลหรือล้างข้อมูลเก่า`
      );
      err.status = 507;
      err.code = 'STORAGE_FULL';
      err.storageQuota = quota;
      throw err;
    }

    const insertRow = {
      worker_name: uniqueLogKey,
      travel_date: travel_date || '',
      route: route || '',
      departure_time: departure_time || '',
      arrival_time: arrival_time || '',
      vehicle_plate: vehicle_plate || '',
      passenger_count,
      gps_status: encodedExtra,
      updated_at: new Date().toISOString()
    };

    let { data: created, error } = await supabase
      .from('travel_logs')
      .insert(insertRow)
      .select('*')
      .single();

    if (error && isMissingExtraColumnError(error)) {
      const { passenger_count: _omit, ...fallbackRow } = insertRow;
      const retry = await supabase
        .from('travel_logs')
        .insert(fallbackRow)
        .select('*')
        .single();
      created = retry.data;
      error = retry.error;
    }

    if (error) throw new Error(error.message);

    return {
      ...withResolvedPassengerCount(created),
      storageQuota: await getQuotaStatus()
    };
  }

  const quota = await getQuotaStatus();
  if (quota.isFull) {
    const err = new Error(
      `พื้นที่เก็บข้อมูลเต็ม (${quota.currentCount}/${quota.maxLimit} รายการ - 100%) ไม่สามารถบันทึกข้อมูลใหม่ได้ชั่วคราว กรุณาส่งออกไฟล์สำรองข้อมูลหรือล้างข้อมูลเก่า`
    );
    err.status = 507;
    err.code = 'STORAGE_FULL';
    err.storageQuota = quota;
    throw err;
  }

  db.prepare(
    `INSERT INTO travel_logs (worker_name, site_name, travel_date, route, departure_time, arrival_time, vehicle_plate, passenger_count, gps_status, updated_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now', 'localtime'))`
  ).run(
    uniqueLogKey,
    cleanSiteName,
    travel_date || '',
    route || '',
    departure_time || '',
    arrival_time || '',
    vehicle_plate || '',
    passenger_count,
    encodedExtra
  );

  syncJsonFallbackFile();
  const created = db.prepare('SELECT * FROM travel_logs WHERE worker_name = ?').get(uniqueLogKey);
  return { ...withResolvedPassengerCount(created), storageQuota: await getQuotaStatus() };
}

async function updateLog(log_id, updates) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('travel_logs')
      .select('*')
      .eq('worker_name', log_id)
      .maybeSingle();
    if (!existing) {
      const err = new Error(`ไม่พบข้อมูลการเดินทางที่ต้องการแก้ไข`);
      err.status = 404;
      throw err;
    }

    const resolvedExisting = withResolvedPassengerCount(existing);
    const { suffix } = splitWorkerLogKey(existing.worker_name);
    const nextWorkerName =
      updates.worker_name !== undefined && String(updates.worker_name).trim()
        ? String(updates.worker_name).trim()
        : resolvedExisting.worker_name;
    const nextLogKey = suffix ? `${nextWorkerName}${suffix}` : existing.worker_name;

    const nextPassengerCount =
      updates.passenger_count !== undefined
        ? normalizePassengerCount(updates.passenger_count)
        : resolvedExisting.passenger_count;
    const nextSiteName =
      updates.site_name !== undefined
        ? String(updates.site_name || '').trim()
        : resolvedExisting.site_name || '';
    const nextNote =
      updates.note !== undefined
        ? String(updates.note || '').trim()
        : resolvedExisting.note || '';
    const encodedExtra = encodeExtraFieldsToGpsStatus(nextPassengerCount, nextSiteName, nextNote);

    const nextRow = {
      worker_name: nextLogKey,
      travel_date:
        updates.travel_date !== undefined ? updates.travel_date : existing.travel_date || '',
      route: updates.route !== undefined ? updates.route : existing.route || '',
      departure_time:
        updates.departure_time !== undefined
          ? updates.departure_time
          : existing.departure_time || '',
      arrival_time:
        updates.arrival_time !== undefined ? updates.arrival_time : existing.arrival_time || '',
      vehicle_plate:
        updates.vehicle_plate !== undefined ? updates.vehicle_plate : existing.vehicle_plate || '',
      passenger_count: nextPassengerCount,
      gps_status: encodedExtra,
      updated_at: new Date().toISOString()
    };

    let { data: updated, error } = await supabase
      .from('travel_logs')
      .update(nextRow)
      .eq('worker_name', log_id)
      .select('*')
      .single();

    if (error && isMissingExtraColumnError(error)) {
      const { passenger_count: _omit, ...fallbackRow } = nextRow;
      const retry = await supabase
        .from('travel_logs')
        .update(fallbackRow)
        .eq('worker_name', log_id)
        .select('*')
        .single();
      updated = retry.data;
      error = retry.error;
    }

    if (error) throw new Error(error.message);

    return {
      ...withResolvedPassengerCount(updated),
      storageQuota: await getQuotaStatus()
    };
  }

  const existing = db
    .prepare('SELECT * FROM travel_logs WHERE worker_name = ?')
    .get(log_id);
  if (!existing) {
    const err = new Error(`ไม่พบข้อมูลการเดินทางที่ต้องการแก้ไข`);
    err.status = 404;
    throw err;
  }

  const resolvedExisting = withResolvedPassengerCount(existing);
  const { suffix } = splitWorkerLogKey(existing.worker_name);
  const nextWorkerName =
    updates.worker_name !== undefined && String(updates.worker_name).trim()
      ? String(updates.worker_name).trim()
      : resolvedExisting.worker_name;
  const nextLogKey = suffix ? `${nextWorkerName}${suffix}` : existing.worker_name;

  const site_name =
    updates.site_name !== undefined
      ? String(updates.site_name || '').trim()
      : resolvedExisting.site_name || '';
  const note =
    updates.note !== undefined
      ? String(updates.note || '').trim()
      : resolvedExisting.note || '';
  const travel_date =
    updates.travel_date !== undefined ? updates.travel_date : existing.travel_date;
  const route = updates.route !== undefined ? updates.route : existing.route;
  const departure_time =
    updates.departure_time !== undefined ? updates.departure_time : existing.departure_time;
  const arrival_time =
    updates.arrival_time !== undefined ? updates.arrival_time : existing.arrival_time;
  const vehicle_plate =
    updates.vehicle_plate !== undefined ? updates.vehicle_plate : existing.vehicle_plate;
  const passenger_count =
    updates.passenger_count !== undefined
      ? normalizePassengerCount(updates.passenger_count)
      : resolvedExisting.passenger_count;
  const gps_status = encodeExtraFieldsToGpsStatus(passenger_count, site_name, note);

  db.prepare(
    `UPDATE travel_logs
     SET worker_name = ?, site_name = ?, travel_date = ?, route = ?, departure_time = ?, arrival_time = ?, vehicle_plate = ?, passenger_count = ?, gps_status = ?, updated_at = datetime('now', 'localtime')
     WHERE worker_name = ?`
  ).run(
    nextLogKey,
    site_name,
    travel_date,
    route,
    departure_time,
    arrival_time,
    vehicle_plate,
    passenger_count,
    gps_status,
    log_id
  );

  syncJsonFallbackFile();
  const updated = db.prepare('SELECT * FROM travel_logs WHERE worker_name = ?').get(nextLogKey);
  return { ...withResolvedPassengerCount(updated), storageQuota: await getQuotaStatus() };
}

async function deleteLog(log_id) {
  if (isSupabaseConfigured && supabase) {
    const { data: existing } = await supabase
      .from('travel_logs')
      .select('*')
      .eq('worker_name', log_id)
      .maybeSingle();
    if (!existing) {
      const err = new Error(`ไม่พบข้อมูลการเดินทางที่ต้องการลบ`);
      err.status = 404;
      throw err;
    }

    const { error } = await supabase.from('travel_logs').delete().eq('worker_name', log_id);
    if (error) throw new Error(error.message);
    return { deleted: withResolvedPassengerCount(existing), storageQuota: await getQuotaStatus() };
  }

  const existing = db
    .prepare('SELECT * FROM travel_logs WHERE worker_name = ?')
    .get(log_id);
  if (!existing) {
    const err = new Error(`ไม่พบข้อมูลการเดินทางที่ต้องการลบ`);
    err.status = 404;
    throw err;
  }

  db.prepare('DELETE FROM travel_logs WHERE worker_name = ?').run(log_id);
  syncJsonFallbackFile();
  return { deleted: withResolvedPassengerCount(existing), storageQuota: await getQuotaStatus() };
}

async function clearOldLogs(mode = 'all') {
  if (isSupabaseConfigured && supabase) {
    const { data: allRows } = await supabase.from('travel_logs').select('worker_name, travel_date');
    const rows = allRows || [];
    const today = getTodayDateString();
    const toDelete =
      mode === 'before_today'
        ? rows.filter((r) => !r.travel_date || r.travel_date < today)
        : rows;
    const targets = toDelete.length > 0 ? toDelete : rows;
    const names = targets.map((r) => r.worker_name);

    if (names.length > 0) {
      const { error } = await supabase.from('travel_logs').delete().in('worker_name', names);
      if (error) throw new Error(error.message);
    }

    return { deletedCount: names.length, storageQuota: await getQuotaStatus() };
  }

  let deletedCount = 0;
  if (mode === 'before_today') {
    const today = getTodayDateString();
    const result = db
      .prepare('DELETE FROM travel_logs WHERE travel_date < ? OR travel_date = ""')
      .run(today);
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
  return { deletedCount, storageQuota: await getQuotaStatus() };
}

async function resetAllToSeed() {
  seedDatabase(true);

  if (isSupabaseConfigured && supabase) {
    // Clear existing rows in Supabase and re-insert clean seed data
    await supabase.from('travel_logs').delete().neq('worker_name', '__never__');
    await supabase.from('master_workers').delete().gte('id', 0);
    await supabase.from('master_routes').delete().gte('id', 0);
    await supabase.from('master_vehicles').delete().gte('id', 0);

    const workers = db
      .prepare('SELECT id, name, default_route, active FROM master_workers')
      .all()
      .map((w) => ({ ...w, active: Boolean(w.active) }));
    const routes = db
      .prepare('SELECT id, name, active FROM master_routes')
      .all()
      .map((r) => ({ ...r, active: Boolean(r.active) }));
    const vehicles = db
      .prepare('SELECT id, plate, active FROM master_vehicles')
      .all()
      .map((v) => ({ ...v, active: Boolean(v.active) }));
    const logs = db
      .prepare(
        'SELECT worker_name, travel_date, route, departure_time, arrival_time, vehicle_plate, gps_status FROM travel_logs'
      )
      .all();
    const settings = db.prepare('SELECT key, value FROM system_settings').all();

    await supabase.from('master_workers').upsert(workers, { onConflict: 'name' });
    await supabase.from('master_routes').upsert(routes, { onConflict: 'name' });
    await supabase.from('master_vehicles').upsert(vehicles, { onConflict: 'plate' });
    await supabase.from('travel_logs').upsert(logs, { onConflict: 'worker_name' });
    await supabase.from('system_settings').upsert(settings, { onConflict: 'key' });
  }

  return getQuotaStatus();
}

module.exports = {
  findUserByUsername,
  findUserById,
  getSettingsMap,
  saveSettingsMap,
  getQuotaStatus,
  listWorkers,
  createWorker,
  updateWorker,
  deleteWorker,
  listRoutes,
  createRoute,
  updateRoute,
  deleteRoute,
  listVehicles,
  createVehicle,
  updateVehicle,
  deleteVehicle,
  listSites,
  createSite,
  updateSite,
  deleteSite,
  listLogs,
  createLog,
  updateLog,
  deleteLog,
  clearOldLogs,
  resetAllToSeed
};
