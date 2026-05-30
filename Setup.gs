// ═══════════════════════════════════════════════════════════════════════════
// Monumento Public Witnessing — COMBINED SETUP + BACKEND SCRIPT
// ═══════════════════════════════════════════════════════════════════════════
//
// HOW TO USE (takes ~3 minutes):
//
//   1. Go to https://script.google.com → click "+ New project"
//   2. Delete all existing code in the editor
//   3. Paste THIS ENTIRE FILE into the editor
//   4. Click Save (💾), name the project "Monumento Witnessing Scheduler"
//   5. In the function dropdown (top bar), select "runSetup" → click Run ▶
//   6. When prompted, click "Review Permissions" → choose your Google account
//      → click "Advanced" → "Go to Monumento..." → Allow
//   7. After it finishes, open the Execution Log (View → Logs)
//      You will see:  ✅ Spreadsheet URL: https://...
//                     📋 Copy this URL into index.html → CONFIG.API_URL
//   8. Click Deploy → New deployment → ⚙ gear → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//      → click Deploy → copy the Web App URL
//   9. Open index.html in VS Code, find CONFIG near the top, and:
//         API_URL: 'paste-the-web-app-url-here',
//         MOCK_MODE: false,
//
//  That's it — the app is live on real data!
// ═══════════════════════════════════════════════════════════════════════════

// ─── Sheet names ──────────────────────────────────────────────────────────
const SHEETS = {
  LOCATIONS:    'Locations',
  TIMESLOTS:    'TimeSlots',
  SCHEDULES:    'Schedules',
  PARTICIPANTS: 'Participants',
  AUDIT_LOG:    'AuditLog',
  CONFIG:       'Config',
};

// ─── Get Spreadsheet ID from script properties (set by runSetup) ──────────
function getSpreadsheetId() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run runSetup() first to create and link the spreadsheet.');
  return id;
}

// ═══════════════════════════════════════════════════════════════════════════
// ONE-TIME SETUP — run this function ONCE from the Apps Script editor
// ═══════════════════════════════════════════════════════════════════════════
function runSetup() {
  // Create a brand-new Google Spreadsheet in the user's Drive
  const ss = SpreadsheetApp.create('Monumento Public Witnessing — Scheduler');
  const id = ss.getId();

  // Save the ID so the API functions can find it after deployment
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);

  // Remove the default blank sheet
  const defaultSheet = ss.getSheets()[0];

  // ── Create all sheets with headers ──────────────────────────────────────
  function makeSheet(name, headers) {
    const sh = ss.insertSheet(name);
    sh.appendRow(headers);
    const headerRange = sh.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
               .setBackground('#1B4332')
               .setFontColor('#ffffff')
               .setFontSize(11);
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, headers.length, 160);
    return sh;
  }

  const locSheet  = makeSheet(SHEETS.LOCATIONS,    ['locationId','name','address','capacity','mapsUrl','isActive','createdAt']);
  const slotSheet = makeSheet(SHEETS.TIMESLOTS,    ['slotId','label','startTime','endTime','isActive','sortOrder']);
  const schSheet  = makeSheet(SHEETS.SCHEDULES,    ['scheduleId','locationId','slotId','date','capacity','seatsUsed','status','createdBy','createdAt','cancelledAt','cancelledBy']);
  const partSheet = makeSheet(SHEETS.PARTICIPANTS, ['participantId','scheduleId','name','companion','seatsHeld','status','waitlistPosition','joinedAt','cancelledAt','promotedAt','contactNote']);
  const logSheet  = makeSheet(SHEETS.AUDIT_LOG,    ['logId','timestamp','action','actorName','targetType','targetId','details']);
  const cfgSheet  = makeSheet(SHEETS.CONFIG,       ['key','value','description']);

  // Delete the original blank sheet
  ss.deleteSheet(defaultSheet);

  // ── Seed Config ──────────────────────────────────────────────────────────
  cfgSheet.appendRow(['nearlyFullThreshold', '0.75',                        'Fraction of capacity that triggers Nearly Full status']);
  cfgSheet.appendRow(['defaultCapacity',     '6',                           'Default maximum participants per schedule']);
  cfgSheet.appendRow(['appTitle',            'Monumento Public Witnessing', 'Application display name']);
  cfgSheet.appendRow(['adminCode',           '2026Admin',                   'Admin passcode — change this to something only admins know!']);

  // ── Seed Locations ───────────────────────────────────────────────────────
  const now = new Date().toISOString();
  locSheet.appendRow(['LOC-001', 'Riza Residence',  'Monumento Area', 6, '', true,  now]);
  locSheet.appendRow(['LOC-002', 'Cielo Residence', 'Monumento Area', 6, '', true,  now]);
  locSheet.appendRow(['LOC-003', 'DXN Office',      'Monumento Area', 6, '', true,  now]);

  // ── Seed Time Slots ──────────────────────────────────────────────────────
  slotSheet.appendRow(['SLOT-001', 'Early Morning (6:00–7:00 AM)',  '06:00', '07:00', true, 1]);
  slotSheet.appendRow(['SLOT-002', 'Morning (7:00–8:00 AM)',         '07:00', '08:00', true, 2]);
  slotSheet.appendRow(['SLOT-003', 'Early Afternoon (4:00–5:00 PM)', '16:00', '17:00', true, 3]);
  slotSheet.appendRow(['SLOT-004', 'Afternoon (5:00–6:00 PM)',       '17:00', '18:00', true, 4]);
  slotSheet.appendRow(['SLOT-005', 'Early Evening (6:00–7:00 PM)',   '18:00', '19:00', true, 5]);
  slotSheet.appendRow(['SLOT-006', 'Evening (7:00–8:00 PM)',         '19:00', '20:00', true, 6]);

  // ── Format column widths for readability ─────────────────────────────────
  locSheet.setColumnWidth(2, 200);  // name
  locSheet.setColumnWidth(3, 250);  // address
  slotSheet.setColumnWidth(2, 260); // label
  schSheet.setColumnWidth(7, 120);  // status
  partSheet.setColumnWidth(3, 200); // name
  partSheet.setColumnWidth(4, 200); // companion

  // ── Done ─────────────────────────────────────────────────────────────────
  Logger.log('');
  Logger.log('✅ Spreadsheet created and seeded successfully!');
  Logger.log('');
  Logger.log('📊 Open your spreadsheet here:');
  Logger.log('   ' + ss.getUrl());
  Logger.log('');
  Logger.log('📋 Spreadsheet ID (for reference): ' + id);
  Logger.log('');
  Logger.log('──────────────────────────────────────────────');
  Logger.log('NEXT STEP: Deploy this script as a Web App:');
  Logger.log('  Click Deploy → New deployment → ⚙ gear → Web app');
  Logger.log('  Execute as: Me | Who has access: Anyone');
  Logger.log('  Then paste the Web App URL into index.html CONFIG.API_URL');
  Logger.log('  and set MOCK_MODE: false');
  Logger.log('──────────────────────────────────────────────');
}


// ═══════════════════════════════════════════════════════════════════════════
// WEB APP API — these functions serve live data after deployment
// ═══════════════════════════════════════════════════════════════════════════

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    let params;
    if (e.postData && e.postData.contents) {
      try { params = JSON.parse(e.postData.contents); } catch(_) { params = e.parameter; }
    } else {
      params = e.parameter || {};
    }
    const action = params.action;
    if (!action) throw new Error('Missing action parameter');
    output.setContent(JSON.stringify({ success: true, data: dispatch(action, params) }));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
  }
  return output;
}

function dispatch(action, p) {
  switch (action) {
    case 'getSchedules':        return getSchedules(p);
    case 'getLocations':        return getLocations(p);
    case 'getTimeSlots':        return getTimeSlots(p);
    case 'getParticipants':     return getParticipants(p);
    case 'getConfig':           return getConfig(p);
    case 'getAuditLog':         return getAuditLog(p);
    case 'verifyAdmin':         return verifyAdmin(p);
    case 'joinSchedule':        return joinSchedule(p);
    case 'cancelParticipation': return cancelParticipation(p);
    case 'createSchedule':      return createSchedule(p);
    case 'cancelSchedule':      return cancelSchedule(p);
    case 'createLocation':      return createLocation(p);
    case 'updateLocation':      return updateLocation(p);
    case 'createTimeSlot':      return createTimeSlot(p);
    case 'updateTimeSlot':      return updateTimeSlot(p);
    case 'exportReport':        return exportReport(p);
    default: throw new Error('Unknown action: ' + action);
  }
}

// ─── Sheet helpers ────────────────────────────────────────────────────────

// GAS reads date/time cells back as JS Date objects, not strings.
// These helpers normalise them to the formats the frontend expects.
function fmtDate(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    const y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
    return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }
  return String(v);
}

function fmtTime(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    return String(v.getHours()).padStart(2, '0') + ':' + String(v.getMinutes()).padStart(2, '0');
  }
  return String(v);
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

function getSheet(name) {
  const sh = getSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] === '' ? '' : row[i]; });
    return obj;
  });
}

function findRowIndex(sheet, field, value) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return -1;
  const headers = data[0].map(h => String(h).trim());
  const col = headers.indexOf(field);
  if (col === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]).toLowerCase() === String(value).toLowerCase()) return i + 1;
  }
  return -1;
}

function getRowObject(sheet, rowIndex) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values  = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => { obj[String(h).trim()] = values[i]; });
  return obj;
}

function updateRowField(sheet, rowIndex, field, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = headers.indexOf(field);
  if (col === -1) throw new Error('Column not found: ' + field);
  sheet.getRange(rowIndex, col + 1).setValue(value);
}

function generateId(prefix) {
  return prefix + '-' + new Date().getTime().toString(36).toUpperCase();
}

function nowIso() { return new Date().toISOString(); }

// ─── Config & Admin ───────────────────────────────────────────────────────

function getConfigMap() {
  const rows = sheetToObjects(getSheet(SHEETS.CONFIG));
  const map = {};
  rows.forEach(r => { map[r.key] = r.value; });
  return map;
}

function validateAdminCode(code) {
  const cfg = getConfigMap();
  if (String(code).trim() !== String(cfg.adminCode || '').trim()) throw new Error('Invalid admin code');
}

function verifyAdmin(p) {
  try { validateAdminCode(p.adminCode); return { isAdmin: true }; }
  catch(_) { return { isAdmin: false }; }
}

function getConfig(p) {
  const cfg = getConfigMap();
  delete cfg.adminCode;
  return cfg;
}

// ─── Audit log ────────────────────────────────────────────────────────────

function writeAuditLog(action, actorName, targetType, targetId, details) {
  try {
    getSheet(SHEETS.AUDIT_LOG).appendRow([
      generateId('LOG'), nowIso(), action, actorName || '',
      targetType || '', targetId || '',
      typeof details === 'object' ? JSON.stringify(details) : (details || ''),
    ]);
  } catch(_) {}
}

function getAuditLog(p) {
  validateAdminCode(p.adminCode);
  const rows = sheetToObjects(getSheet(SHEETS.AUDIT_LOG)).reverse();
  const limit = parseInt(p.limit) || 100;
  let result = rows;
  if (p.targetType) result = result.filter(r => r.targetType === p.targetType);
  if (p.targetId)   result = result.filter(r => r.targetId === p.targetId);
  return result.slice(0, limit);
}

// ─── Locations ────────────────────────────────────────────────────────────

function getLocations(p) {
  const rows = sheetToObjects(getSheet(SHEETS.LOCATIONS));
  const inc  = p.includeInactive === 'true' || p.includeInactive === true;
  return inc ? rows : rows.filter(r => String(r.isActive) === 'TRUE' || r.isActive === true);
}

function createLocation(p) {
  validateAdminCode(p.adminCode);
  if (!p.name) throw new Error('Location name is required');
  const id = generateId('LOC');
  getSheet(SHEETS.LOCATIONS).appendRow([id, p.name, p.address||'', parseInt(p.capacity)||6, p.mapsUrl||'', true, nowIso()]);
  writeAuditLog('CREATE_LOCATION', 'Admin', 'LOCATION', id, { name: p.name });
  return { locationId: id };
}

function updateLocation(p) {
  validateAdminCode(p.adminCode);
  const sh = getSheet(SHEETS.LOCATIONS);
  const row = findRowIndex(sh, 'locationId', p.locationId);
  if (row === -1) throw new Error('Location not found');
  if (p.name     !== undefined) updateRowField(sh, row, 'name',     p.name);
  if (p.address  !== undefined) updateRowField(sh, row, 'address',  p.address);
  if (p.capacity !== undefined) updateRowField(sh, row, 'capacity', parseInt(p.capacity));
  if (p.mapsUrl  !== undefined) updateRowField(sh, row, 'mapsUrl',  p.mapsUrl);
  if (p.isActive !== undefined) updateRowField(sh, row, 'isActive', p.isActive === 'true' || p.isActive === true);
  writeAuditLog('UPDATE_LOCATION', 'Admin', 'LOCATION', p.locationId, {});
  return { updated: true };
}

// ─── Time Slots ───────────────────────────────────────────────────────────

function getTimeSlots(p) {
  return sheetToObjects(getSheet(SHEETS.TIMESLOTS))
    .filter(r => String(r.isActive) === 'TRUE' || r.isActive === true)
    .sort((a, b) => (parseInt(a.sortOrder)||0) - (parseInt(b.sortOrder)||0))
    .map(r => ({ ...r, startTime: fmtTime(r.startTime), endTime: fmtTime(r.endTime) }));
}

function createTimeSlot(p) {
  validateAdminCode(p.adminCode);
  if (!p.label) throw new Error('Slot label is required');
  const id = generateId('SLOT');
  getSheet(SHEETS.TIMESLOTS).appendRow([id, p.label, p.startTime||'', p.endTime||'', true, parseInt(p.sortOrder)||99]);
  return { slotId: id };
}

function updateTimeSlot(p) {
  validateAdminCode(p.adminCode);
  const sh = getSheet(SHEETS.TIMESLOTS);
  const row = findRowIndex(sh, 'slotId', p.slotId);
  if (row === -1) throw new Error('Time slot not found');
  if (p.label     !== undefined) updateRowField(sh, row, 'label',     p.label);
  if (p.startTime !== undefined) updateRowField(sh, row, 'startTime', p.startTime);
  if (p.endTime   !== undefined) updateRowField(sh, row, 'endTime',   p.endTime);
  if (p.sortOrder !== undefined) updateRowField(sh, row, 'sortOrder', parseInt(p.sortOrder));
  if (p.isActive  !== undefined) updateRowField(sh, row, 'isActive',  p.isActive === 'true' || p.isActive === true);
  return { updated: true };
}

// ─── Schedules ────────────────────────────────────────────────────────────

function computeStatus(seatsUsed, capacity, threshold) {
  const u = parseInt(seatsUsed)||0, c = parseInt(capacity)||1, t = parseFloat(threshold)||0.75;
  if (u >= c)      return 'FULL';
  if (u/c >= t)    return 'NEARLY_FULL';
  return 'AVAILABLE';
}

function getSchedules(p) {
  const locs  = {}; sheetToObjects(getSheet(SHEETS.LOCATIONS)).forEach(l => { locs[l.locationId] = l; });
  const slots = {}; sheetToObjects(getSheet(SHEETS.TIMESLOTS)).forEach(s => { slots[s.slotId]    = s; });
  const parts = {}; sheetToObjects(getSheet(SHEETS.PARTICIPANTS)).forEach(pp => {
    if (!parts[pp.scheduleId]) parts[pp.scheduleId] = [];
    parts[pp.scheduleId].push(pp);
  });
  const cfg = getConfigMap();
  const threshold = parseFloat(cfg.nearlyFullThreshold) || 0.75;

  let scheds = sheetToObjects(getSheet(SHEETS.SCHEDULES));
  if (p.dateFrom)   scheds = scheds.filter(s => String(s.date) >= p.dateFrom);
  if (p.dateTo)     scheds = scheds.filter(s => String(s.date) <= p.dateTo);
  if (p.locationId) scheds = scheds.filter(s => s.locationId === p.locationId);

  return scheds.map(s => {
    const loc  = locs[s.locationId]  || {};
    const slot = slots[s.slotId]     || {};
    const pList = (parts[s.scheduleId] || []).map(pp => ({
      participantId: pp.participantId, scheduleId: pp.scheduleId,
      name: pp.name, companion: pp.companion, seatsHeld: pp.seatsHeld,
      status: pp.status, waitlistPosition: pp.waitlistPosition,
      joinedAt: pp.joinedAt, cancelledAt: pp.cancelledAt, promotedAt: pp.promotedAt,
    }));
    return {
      scheduleId: s.scheduleId, locationId: s.locationId, slotId: s.slotId,
      date: fmtDate(s.date), capacity: parseInt(s.capacity)||6, seatsUsed: parseInt(s.seatsUsed)||0,
      status: s.status,
      availabilityStatus: s.status === 'CANCELLED' ? 'CANCELLED' : computeStatus(s.seatsUsed, s.capacity, threshold),
      locationName: loc.name||'', locationAddress: loc.address||'', mapsUrl: loc.mapsUrl||'',
      slotLabel: slot.label||'', slotStart: fmtTime(slot.startTime)||'', slotEnd: fmtTime(slot.endTime)||'',
      createdBy: s.createdBy||'', createdAt: s.createdAt||'',
      cancelledAt: s.cancelledAt||'', participants: pList,
    };
  }).sort((a, b) => a.date.localeCompare(b.date) || a.slotStart.localeCompare(b.slotStart));
}

function checkMeetingConflict(date, slotId) {
  const slots = sheetToObjects(getSheet(SHEETS.TIMESLOTS));
  const slot = slots.filter(function(s) { return s.slotId === slotId; })[0];
  if (!slot) return;
  const parts = date.split('-');
  const day = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2])).getDay();
  const toMins = function(t) { const hm = fmtTime(t).split(':'); return parseInt(hm[0]) * 60 + (parseInt(hm[1]) || 0); };
  const s = toMins(slot.startTime), e = toMins(slot.endTime);
  if (day === 4 && s < 20 * 60 && e > 18 * 60)
    throw new Error('Cannot schedule during the midweek meeting (Thursday 6–8 PM)');
  if (day === 0 && s < 17 * 60 && e > 15 * 60)
    throw new Error('Cannot schedule during the weekend meeting (Sunday 3–5 PM)');
}

function createSchedule(p) {
  if (!p.locationId || !p.slotId || !p.date) throw new Error('locationId, slotId, and date are required');
  checkMeetingConflict(p.date, p.slotId);
  const locSh = getSheet(SHEETS.LOCATIONS);
  const locRow = findRowIndex(locSh, 'locationId', p.locationId);
  if (locRow === -1) throw new Error('Location not found');
  const loc = getRowObject(locSh, locRow);
  const existing = sheetToObjects(getSheet(SHEETS.SCHEDULES));
  if (existing.find(s => s.locationId===p.locationId && s.slotId===p.slotId && fmtDate(s.date)===p.date && s.status==='OPEN'))
    throw new Error('A schedule already exists for this location, date, and time slot');
  const id = generateId('SCH');
  getSheet(SHEETS.SCHEDULES).appendRow([id, p.locationId, p.slotId, p.date, parseInt(loc.capacity)||6, 0, 'OPEN', p.createdBy||'Admin', nowIso(), '', '']);
  writeAuditLog('CREATE_SCHEDULE', p.createdBy||'Admin', 'SCHEDULE', id, { locationId: p.locationId, slotId: p.slotId, date: p.date });
  return { scheduleId: id };
}

function cancelSchedule(p) {
  validateAdminCode(p.adminCode);
  const sh = getSheet(SHEETS.SCHEDULES);
  const row = findRowIndex(sh, 'scheduleId', p.scheduleId);
  if (row === -1) throw new Error('Schedule not found');
  updateRowField(sh, row, 'status', 'CANCELLED');
  updateRowField(sh, row, 'cancelledAt', nowIso());
  updateRowField(sh, row, 'cancelledBy', p.cancelledBy||'Admin');
  writeAuditLog('CANCEL_SCHEDULE', p.cancelledBy||'Admin', 'SCHEDULE', p.scheduleId, {});
  return { cancelled: true };
}

// ─── Participants ─────────────────────────────────────────────────────────

function getParticipants(p) {
  if (!p.scheduleId) throw new Error('scheduleId is required');
  return sheetToObjects(getSheet(SHEETS.PARTICIPANTS))
    .filter(r => r.scheduleId === p.scheduleId)
    .map(r => ({ participantId: r.participantId, scheduleId: r.scheduleId, name: r.name, companion: r.companion, seatsHeld: r.seatsHeld, status: r.status, waitlistPosition: r.waitlistPosition, joinedAt: r.joinedAt, cancelledAt: r.cancelledAt, promotedAt: r.promotedAt }));
}

function joinSchedule(p) {
  const name = String(p.name||'').trim();
  if (!name) throw new Error('Name is required');
  const schSh = getSheet(SHEETS.SCHEDULES);
  const schRow = findRowIndex(schSh, 'scheduleId', p.scheduleId);
  if (schRow === -1) throw new Error('Schedule not found');
  const sch = getRowObject(schSh, schRow);
  if (sch.status === 'CANCELLED') throw new Error('This schedule has been cancelled');
  const companion = String(p.companion||'').trim();
  const seatsHeld = companion ? 2 : 1;
  const allParts  = sheetToObjects(getSheet(SHEETS.PARTICIPANTS));
  // Duplicate check
  if (allParts.find(r => r.scheduleId===p.scheduleId && (r.status==='ACTIVE'||r.status==='WAITLISTED'||r.status==='PROMOTED') && String(r.name).toLowerCase()===name.toLowerCase()))
    throw new Error('You are already registered for this schedule');
  // Conflict check
  const allScheds = sheetToObjects(schSh);
  const conflict = allParts.find(r => {
    if (String(r.name).toLowerCase()!==name.toLowerCase() || (r.status!=='ACTIVE'&&r.status!=='PROMOTED')) return false;
    const rs = allScheds.find(s => s.scheduleId===r.scheduleId);
    return rs && String(rs.date)===String(sch.date) && String(rs.slotId)===String(sch.slotId) && rs.scheduleId!==p.scheduleId;
  });
  if (conflict) throw new Error('You are already registered at another location during this time slot');
  const seatsUsed = parseInt(sch.seatsUsed)||0;
  const capacity  = parseInt(sch.capacity)||6;
  const isWL = seatsUsed + seatsHeld > capacity;
  const wlPos = isWL ? allParts.filter(r => r.scheduleId===p.scheduleId && r.status==='WAITLISTED').length + 1 : null;
  const pid = generateId('PRT');
  getSheet(SHEETS.PARTICIPANTS).appendRow([pid, p.scheduleId, name, companion, seatsHeld, isWL?'WAITLISTED':'ACTIVE', isWL?wlPos:'', nowIso(), '', '', p.contactNote||'']);
  if (!isWL) updateRowField(schSh, schRow, 'seatsUsed', seatsUsed + seatsHeld);
  writeAuditLog('JOIN', name, 'PARTICIPANT', pid, { scheduleId: p.scheduleId, seatsHeld, status: isWL?'WAITLISTED':'ACTIVE' });
  return { participantId: pid, status: isWL?'WAITLISTED':'ACTIVE', waitlistPosition: wlPos };
}

function cancelParticipation(p) {
  const name = String(p.name||'').trim();
  if (!p.participantId || !name) throw new Error('participantId and name are required');
  const partSh = getSheet(SHEETS.PARTICIPANTS);
  const partRow = findRowIndex(partSh, 'participantId', p.participantId);
  if (partRow === -1) throw new Error('Registration not found');
  const part = getRowObject(partSh, partRow);
  if (String(part.name).toLowerCase() !== name.toLowerCase()) throw new Error('Name does not match registration');
  if (part.status === 'CANCELLED') throw new Error('Already cancelled');
  updateRowField(partSh, partRow, 'status', 'CANCELLED');
  updateRowField(partSh, partRow, 'cancelledAt', nowIso());
  const schSh = getSheet(SHEETS.SCHEDULES);
  const schRow = findRowIndex(schSh, 'scheduleId', part.scheduleId);
  const promoted = [];
  if (schRow !== -1 && (part.status==='ACTIVE'||part.status==='PROMOTED')) {
    const sch = getRowObject(schSh, schRow);
    const newSeats = Math.max(0, (parseInt(sch.seatsUsed)||0) - (parseInt(part.seatsHeld)||1));
    updateRowField(schSh, schRow, 'seatsUsed', newSeats);
    promoted.push(...promoteWaitlist(part.scheduleId, newSeats, parseInt(sch.capacity)||6));
  }
  writeAuditLog('CANCEL', name, 'PARTICIPANT', p.participantId, { scheduleId: part.scheduleId, promoted });
  return { cancelled: true, promoted };
}

function promoteWaitlist(scheduleId, currentSeats, capacity) {
  const partSh = getSheet(SHEETS.PARTICIPANTS);
  const schSh  = getSheet(SHEETS.SCHEDULES);
  const waiting = sheetToObjects(partSh)
    .filter(r => r.scheduleId===scheduleId && r.status==='WAITLISTED')
    .sort((a, b) => (parseInt(a.waitlistPosition)||999) - (parseInt(b.waitlistPosition)||999));
  const promoted = [];
  let seats = currentSeats;
  for (const wp of waiting) {
    const need = parseInt(wp.seatsHeld)||1;
    if (seats + need <= capacity) {
      const row = findRowIndex(partSh, 'participantId', wp.participantId);
      if (row !== -1) {
        updateRowField(partSh, row, 'status', 'PROMOTED');
        updateRowField(partSh, row, 'promotedAt', nowIso());
        updateRowField(partSh, row, 'waitlistPosition', '');
        seats += need;
        promoted.push(wp.name);
        writeAuditLog('PROMOTE_WAITLIST', wp.name, 'PARTICIPANT', wp.participantId, { scheduleId });
      }
    }
  }
  // Renumber remaining waitlist
  sheetToObjects(partSh)
    .filter(r => r.scheduleId===scheduleId && r.status==='WAITLISTED' && !promoted.includes(r.name))
    .sort((a,b) => (parseInt(a.waitlistPosition)||999)-(parseInt(b.waitlistPosition)||999))
    .forEach((wp, i) => {
      const row = findRowIndex(partSh, 'participantId', wp.participantId);
      if (row !== -1) updateRowField(partSh, row, 'waitlistPosition', i+1);
    });
  const schRow = findRowIndex(schSh, 'scheduleId', scheduleId);
  if (schRow !== -1) updateRowField(schSh, schRow, 'seatsUsed', seats);
  return promoted;
}

// ─── Export ───────────────────────────────────────────────────────────────

function exportReport(p) {
  validateAdminCode(p.adminCode);
  const type = p.reportType || 'schedules';
  let rows = [];
  if (type === 'schedules') {
    rows = [['Schedule ID','Location','Date','Time Slot','Capacity','Seats Used','Status','Created By']];
    getSchedules(p).forEach(s => rows.push([s.scheduleId, s.locationName, s.date, s.slotLabel, s.capacity, s.seatsUsed, s.status, s.createdBy]));
  } else if (type === 'participants') {
    rows = [['Participant ID','Name','Companion','Seats','Location','Date','Slot','Status','Joined At']];
    const schMap = {}; sheetToObjects(getSheet(SHEETS.SCHEDULES)).forEach(s => { schMap[s.scheduleId]=s; });
    const locMap = {}; sheetToObjects(getSheet(SHEETS.LOCATIONS)).forEach(l => { locMap[l.locationId]=l; });
    const slotMap= {}; sheetToObjects(getSheet(SHEETS.TIMESLOTS)).forEach(s => { slotMap[s.slotId]=s; });
    sheetToObjects(getSheet(SHEETS.PARTICIPANTS)).forEach(pp => {
      const s = schMap[pp.scheduleId]||{}, l = locMap[s.locationId]||{}, sl = slotMap[s.slotId]||{};
      rows.push([pp.participantId, pp.name, pp.companion, pp.seatsHeld, l.name||'', s.date||'', sl.label||'', pp.status, pp.joinedAt]);
    });
  } else if (type === 'waitlist') {
    rows = [['Name','Companion','Location','Date','Slot','Position','Joined At']];
    const schMap = {}; sheetToObjects(getSheet(SHEETS.SCHEDULES)).forEach(s => { schMap[s.scheduleId]=s; });
    const locMap = {}; sheetToObjects(getSheet(SHEETS.LOCATIONS)).forEach(l => { locMap[l.locationId]=l; });
    const slotMap= {}; sheetToObjects(getSheet(SHEETS.TIMESLOTS)).forEach(s => { slotMap[s.slotId]=s; });
    sheetToObjects(getSheet(SHEETS.PARTICIPANTS)).filter(pp => pp.status==='WAITLISTED').forEach(pp => {
      const s = schMap[pp.scheduleId]||{}, l = locMap[s.locationId]||{}, sl = slotMap[s.slotId]||{};
      rows.push([pp.name, pp.companion, l.name||'', s.date||'', sl.label||'', pp.waitlistPosition, pp.joinedAt]);
    });
  }
  return { csvContent: rows.map(r => r.map(c => '"' + String(c??'').replace(/"/g,'""') + '"').join(',')).join('\n') };
}
