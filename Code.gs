// ─────────────────────────────────────────────────────────────────────────────
// Monumento Public Witnessing — Google Apps Script Backend
// ─────────────────────────────────────────────────────────────────────────────
// SETUP INSTRUCTIONS:
//   1. Create a new Google Spreadsheet.
//   2. Copy this file into the Apps Script editor (Extensions → Apps Script).
//   3. Set SPREADSHEET_ID below to your spreadsheet's ID (from the URL).
//   4. Run seedData() once from the editor to create all sheets and sample data.
//   5. Deploy: Deploy → New Deployment → Web App
//      - Execute as: Me
//      - Who has access: Anyone
//   6. Copy the deployed URL into CONFIG.API_URL in index.html.
//   7. Set CONFIG.MOCK_MODE = false in index.html.
// ─────────────────────────────────────────────────────────────────────────────

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

const SHEETS = {
  LOCATIONS:    'Locations',
  TIMESLOTS:    'TimeSlots',
  SCHEDULES:    'Schedules',
  PARTICIPANTS: 'Participants',
  AUDIT_LOG:    'AuditLog',
  CONFIG:       'Config',
};

// ─── HTTP Handlers ────────────────────────────────────────────────────────────

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

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
    const result = dispatch(action, params);
    output.setContent(JSON.stringify({ success: true, data: result }));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
  }

  return output;
}

function dispatch(action, p) {
  switch (action) {
    // READ
    case 'getSchedules':       return getSchedules(p);
    case 'getLocations':       return getLocations(p);
    case 'getTimeSlots':       return getTimeSlots(p);
    case 'getParticipants':    return getParticipants(p);
    case 'getConfig':          return getConfig(p);
    case 'getAuditLog':        return getAuditLog(p);
    case 'verifyAdmin':        return verifyAdmin(p);
    // WRITE
    case 'joinSchedule':       return joinSchedule(p);
    case 'cancelParticipation':return cancelParticipation(p);
    case 'createSchedule':     return createSchedule(p);
    case 'cancelSchedule':     return cancelSchedule(p);
    case 'createLocation':     return createLocation(p);
    case 'updateLocation':     return updateLocation(p);
    case 'createTimeSlot':     return createTimeSlot(p);
    case 'updateTimeSlot':     return updateTimeSlot(p);
    case 'exportReport':       return exportReport(p);
    default: throw new Error('Unknown action: ' + action);
  }
}

// ─── Sheet Helpers ────────────────────────────────────────────────────────────

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(name) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
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
  const colIdx = headers.indexOf(field);
  if (colIdx === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIdx]).toLowerCase() === String(value).toLowerCase()) {
      return i + 1; // 1-indexed sheet row
    }
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
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIdx  = headers.map(h => String(h).trim()).indexOf(field);
  if (colIdx === -1) throw new Error('Column not found: ' + field);
  sheet.getRange(rowIndex, colIdx + 1).setValue(value);
}

function generateId(prefix) {
  return prefix + '-' + new Date().getTime().toString(36).toUpperCase();
}

function nowIso() {
  return new Date().toISOString();
}

// ─── Config & Admin ───────────────────────────────────────────────────────────

function getConfigMap() {
  const sheet = getSheet(SHEETS.CONFIG);
  const rows = sheetToObjects(sheet);
  const map = {};
  rows.forEach(r => { map[r.key] = r.value; });
  return map;
}

function validateAdminCode(code) {
  const cfg = getConfigMap();
  if (!cfg.adminCode) throw new Error('Admin code not configured');
  if (String(code).trim() !== String(cfg.adminCode).trim()) throw new Error('Invalid admin code');
}

function verifyAdmin(p) {
  try {
    validateAdminCode(p.adminCode);
    return { isAdmin: true };
  } catch(_) {
    return { isAdmin: false };
  }
}

function getConfig(p) {
  const cfg = getConfigMap();
  delete cfg.adminCode; // never expose
  return cfg;
}

// ─── Audit Logging ────────────────────────────────────────────────────────────

function writeAuditLog(action, actorName, targetType, targetId, details) {
  try {
    const sheet = getSheet(SHEETS.AUDIT_LOG);
    sheet.appendRow([
      generateId('LOG'),
      nowIso(),
      action,
      actorName || '',
      targetType || '',
      targetId || '',
      typeof details === 'object' ? JSON.stringify(details) : (details || ''),
    ]);
  } catch(e) {
    // Audit failures must not break the main operation
  }
}

function getAuditLog(p) {
  validateAdminCode(p.adminCode);
  const sheet = getSheet(SHEETS.AUDIT_LOG);
  const rows = sheetToObjects(sheet);
  rows.reverse(); // newest first
  const limit = parseInt(p.limit) || 100;
  let result = rows;
  if (p.targetType) result = result.filter(r => r.targetType === p.targetType);
  if (p.targetId)   result = result.filter(r => r.targetId === p.targetId);
  return result.slice(0, limit);
}

// ─── Locations ────────────────────────────────────────────────────────────────

function getLocations(p) {
  const sheet = getSheet(SHEETS.LOCATIONS);
  const rows = sheetToObjects(sheet);
  const includeInactive = p.includeInactive === 'true' || p.includeInactive === true;
  return includeInactive ? rows : rows.filter(r => String(r.isActive) === 'TRUE' || r.isActive === true);
}

function createLocation(p) {
  validateAdminCode(p.adminCode);
  if (!p.name) throw new Error('Location name is required');
  const sheet = getSheet(SHEETS.LOCATIONS);
  const locId = generateId('LOC');
  sheet.appendRow([
    locId,
    p.name,
    p.address || '',
    parseInt(p.capacity) || 6,
    p.mapsUrl || '',
    true,
    nowIso(),
  ]);
  writeAuditLog('CREATE_LOCATION', p.name, 'LOCATION', locId, { name: p.name });
  return { locationId: locId };
}

function updateLocation(p) {
  validateAdminCode(p.adminCode);
  const sheet = getSheet(SHEETS.LOCATIONS);
  const rowIdx = findRowIndex(sheet, 'locationId', p.locationId);
  if (rowIdx === -1) throw new Error('Location not found');
  const before = getRowObject(sheet, rowIdx);
  if (p.name     !== undefined) updateRowField(sheet, rowIdx, 'name',     p.name);
  if (p.address  !== undefined) updateRowField(sheet, rowIdx, 'address',  p.address);
  if (p.capacity !== undefined) updateRowField(sheet, rowIdx, 'capacity', parseInt(p.capacity));
  if (p.mapsUrl  !== undefined) updateRowField(sheet, rowIdx, 'mapsUrl',  p.mapsUrl);
  if (p.isActive !== undefined) updateRowField(sheet, rowIdx, 'isActive', p.isActive === 'true' || p.isActive === true);
  writeAuditLog('UPDATE_LOCATION', 'Admin', 'LOCATION', p.locationId, { before });
  return { updated: true };
}

// ─── Time Slots ───────────────────────────────────────────────────────────────

function getTimeSlots(p) {
  const sheet = getSheet(SHEETS.TIMESLOTS);
  const rows = sheetToObjects(sheet);
  return rows
    .filter(r => String(r.isActive) === 'TRUE' || r.isActive === true)
    .sort((a, b) => (parseInt(a.sortOrder) || 0) - (parseInt(b.sortOrder) || 0));
}

function createTimeSlot(p) {
  validateAdminCode(p.adminCode);
  if (!p.label) throw new Error('Slot label is required');
  const sheet = getSheet(SHEETS.TIMESLOTS);
  const slotId = generateId('SLOT');
  sheet.appendRow([slotId, p.label, p.startTime || '', p.endTime || '', true, parseInt(p.sortOrder) || 99]);
  writeAuditLog('CREATE_TIMESLOT', 'Admin', 'TIMESLOT', slotId, { label: p.label });
  return { slotId };
}

function updateTimeSlot(p) {
  validateAdminCode(p.adminCode);
  const sheet = getSheet(SHEETS.TIMESLOTS);
  const rowIdx = findRowIndex(sheet, 'slotId', p.slotId);
  if (rowIdx === -1) throw new Error('Time slot not found');
  if (p.label     !== undefined) updateRowField(sheet, rowIdx, 'label',     p.label);
  if (p.startTime !== undefined) updateRowField(sheet, rowIdx, 'startTime', p.startTime);
  if (p.endTime   !== undefined) updateRowField(sheet, rowIdx, 'endTime',   p.endTime);
  if (p.sortOrder !== undefined) updateRowField(sheet, rowIdx, 'sortOrder', parseInt(p.sortOrder));
  if (p.isActive  !== undefined) updateRowField(sheet, rowIdx, 'isActive',  p.isActive === 'true' || p.isActive === true);
  return { updated: true };
}

// ─── Schedules ────────────────────────────────────────────────────────────────

function computeAvailabilityStatus(seatsUsed, capacity, threshold) {
  const used = parseInt(seatsUsed) || 0;
  const cap  = parseInt(capacity)  || 1;
  const t    = parseFloat(threshold) || 0.75;
  if (used >= cap)         return 'FULL';
  if (used / cap >= t)     return 'NEARLY_FULL';
  return 'AVAILABLE';
}

function getSchedules(p) {
  const locSheet  = getSheet(SHEETS.LOCATIONS);
  const slotSheet = getSheet(SHEETS.TIMESLOTS);
  const schSheet  = getSheet(SHEETS.SCHEDULES);
  const partSheet = getSheet(SHEETS.PARTICIPANTS);

  const locations = {};
  sheetToObjects(locSheet).forEach(l => { locations[l.locationId] = l; });

  const slots = {};
  sheetToObjects(slotSheet).forEach(s => { slots[s.slotId] = s; });

  const allParticipants = sheetToObjects(partSheet);
  const partsBySchedule = {};
  allParticipants.forEach(p => {
    if (!partsBySchedule[p.scheduleId]) partsBySchedule[p.scheduleId] = [];
    partsBySchedule[p.scheduleId].push(p);
  });

  const cfg = getConfigMap();
  const threshold = parseFloat(cfg.nearlyFullThreshold) || 0.75;

  let schedules = sheetToObjects(schSheet);

  if (p.dateFrom)    schedules = schedules.filter(s => String(s.date) >= p.dateFrom);
  if (p.dateTo)      schedules = schedules.filter(s => String(s.date) <= p.dateTo);
  if (p.locationId)  schedules = schedules.filter(s => s.locationId === p.locationId);

  return schedules.map(s => {
    const loc  = locations[s.locationId] || {};
    const slot = slots[s.slotId]         || {};
    const parts = (partsBySchedule[s.scheduleId] || []).map(pp => ({
      participantId:   pp.participantId,
      scheduleId:      pp.scheduleId,
      name:            pp.name,
      companion:       pp.companion,
      seatsHeld:       pp.seatsHeld,
      status:          pp.status,
      waitlistPosition:pp.waitlistPosition,
      joinedAt:        pp.joinedAt,
      cancelledAt:     pp.cancelledAt,
      promotedAt:      pp.promotedAt,
      // contactNote intentionally excluded from public list
    }));

    const availabilityStatus = s.status === 'CANCELLED'
      ? 'CANCELLED'
      : computeAvailabilityStatus(s.seatsUsed, s.capacity, threshold);

    return {
      scheduleId:         s.scheduleId,
      locationId:         s.locationId,
      slotId:             s.slotId,
      date:               String(s.date),
      capacity:           parseInt(s.capacity) || 6,
      seatsUsed:          parseInt(s.seatsUsed) || 0,
      status:             s.status,
      availabilityStatus,
      locationName:       loc.name        || '',
      locationAddress:    loc.address     || '',
      mapsUrl:            loc.mapsUrl     || '',
      slotLabel:          slot.label      || '',
      slotStart:          slot.startTime  || '',
      slotEnd:            slot.endTime    || '',
      createdBy:          s.createdBy     || '',
      createdAt:          s.createdAt     || '',
      cancelledAt:        s.cancelledAt   || '',
      participants:       parts,
    };
  }).sort((a, b) => a.date.localeCompare(b.date) || a.slotStart.localeCompare(b.slotStart));
}

function checkMeetingConflict(date, slotId) {
  var slots = sheetToObjects(getSheet(SHEETS.TIMESLOTS));
  var slot = slots.filter(function(s) { return s.slotId === slotId; })[0];
  if (!slot) return;
  var parts = date.split('-');
  var day = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2])).getDay();
  var toMins = function(t) { var hm = String(t).split(':'); return parseInt(hm[0]) * 60 + (parseInt(hm[1]) || 0); };
  var s = toMins(slot.startTime), e = toMins(slot.endTime);
  if (day === 4 && s < 20 * 60 && e > 18 * 60)
    throw new Error('Cannot schedule during the midweek meeting (Thursday 6–8 PM)');
  if (day === 0 && s < 17 * 60 && e > 15 * 60)
    throw new Error('Cannot schedule during the weekend meeting (Sunday 3–5 PM)');
}

function createSchedule(p) {
  if (!p.locationId || !p.slotId || !p.date) throw new Error('locationId, slotId, and date are required');
  checkMeetingConflict(p.date, p.slotId);

  const locSheet  = getSheet(SHEETS.LOCATIONS);
  const slotSheet = getSheet(SHEETS.TIMESLOTS);
  const schSheet  = getSheet(SHEETS.SCHEDULES);

  const locRowIdx = findRowIndex(locSheet, 'locationId', p.locationId);
  if (locRowIdx === -1) throw new Error('Location not found');
  const loc = getRowObject(locSheet, locRowIdx);

  const slotRowIdx = findRowIndex(slotSheet, 'slotId', p.slotId);
  if (slotRowIdx === -1) throw new Error('Time slot not found');

  // Duplicate check
  const existing = sheetToObjects(schSheet);
  const dup = existing.find(s =>
    s.locationId === p.locationId &&
    s.slotId     === p.slotId     &&
    String(s.date) === p.date     &&
    s.status === 'OPEN'
  );
  if (dup) throw new Error('A schedule already exists for this location, date, and time slot');

  const schId = generateId('SCH');
  schSheet.appendRow([
    schId,
    p.locationId,
    p.slotId,
    p.date,
    parseInt(loc.capacity) || 6,
    0, // seatsUsed
    'OPEN',
    p.createdBy || 'Admin',
    nowIso(),
    '', // cancelledAt
    '', // cancelledBy
  ]);
  writeAuditLog('CREATE_SCHEDULE', p.createdBy || 'Admin', 'SCHEDULE', schId, { locationId: p.locationId, slotId: p.slotId, date: p.date });
  return { scheduleId: schId };
}

function cancelSchedule(p) {
  validateAdminCode(p.adminCode);
  const sheet = getSheet(SHEETS.SCHEDULES);
  const rowIdx = findRowIndex(sheet, 'scheduleId', p.scheduleId);
  if (rowIdx === -1) throw new Error('Schedule not found');
  updateRowField(sheet, rowIdx, 'status',      'CANCELLED');
  updateRowField(sheet, rowIdx, 'cancelledAt', nowIso());
  updateRowField(sheet, rowIdx, 'cancelledBy', p.cancelledBy || 'Admin');
  writeAuditLog('CANCEL_SCHEDULE', p.cancelledBy || 'Admin', 'SCHEDULE', p.scheduleId, {});
  return { cancelled: true };
}

// ─── Participants ─────────────────────────────────────────────────────────────

function getParticipants(p) {
  if (!p.scheduleId) throw new Error('scheduleId is required');
  const sheet = getSheet(SHEETS.PARTICIPANTS);
  return sheetToObjects(sheet)
    .filter(r => r.scheduleId === p.scheduleId)
    .map(r => ({
      participantId:    r.participantId,
      scheduleId:       r.scheduleId,
      name:             r.name,
      companion:        r.companion,
      seatsHeld:        r.seatsHeld,
      status:           r.status,
      waitlistPosition: r.waitlistPosition,
      joinedAt:         r.joinedAt,
      cancelledAt:      r.cancelledAt,
      promotedAt:       r.promotedAt,
      // contactNote excluded
    }));
}

function joinSchedule(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('Name is required');
  if (!p.scheduleId) throw new Error('scheduleId is required');

  const schSheet  = getSheet(SHEETS.SCHEDULES);
  const partSheet = getSheet(SHEETS.PARTICIPANTS);

  const schRowIdx = findRowIndex(schSheet, 'scheduleId', p.scheduleId);
  if (schRowIdx === -1) throw new Error('Schedule not found');
  const sch = getRowObject(schSheet, schRowIdx);

  if (sch.status === 'CANCELLED') throw new Error('This schedule has been cancelled');

  const companion  = String(p.companion || '').trim();
  const seatsHeld  = companion ? 2 : 1;
  const allParts   = sheetToObjects(partSheet);

  // Duplicate check
  const dup = allParts.find(r =>
    r.scheduleId === p.scheduleId &&
    (r.status === 'ACTIVE' || r.status === 'WAITLISTED' || r.status === 'PROMOTED') &&
    String(r.name).toLowerCase() === name.toLowerCase()
  );
  if (dup) throw new Error('You are already registered for this schedule');

  // Conflict check — same name, same date, same slotId, different scheduleId
  const schDate   = String(sch.date);
  const schSlotId = String(sch.slotId);
  const schSheet2 = getSheet(SHEETS.SCHEDULES);
  const allScheds = sheetToObjects(schSheet2);
  const conflicting = allParts.find(r => {
    if (String(r.name).toLowerCase() !== name.toLowerCase()) return false;
    if (r.status !== 'ACTIVE' && r.status !== 'PROMOTED') return false;
    const rSch = allScheds.find(s => s.scheduleId === r.scheduleId);
    if (!rSch) return false;
    return String(rSch.date) === schDate && String(rSch.slotId) === schSlotId && rSch.scheduleId !== p.scheduleId;
  });
  if (conflicting) {
    const cSch = allScheds.find(s => s.scheduleId === conflicting.scheduleId);
    throw new Error('Conflict: you are already registered at another location during this time slot' + (cSch ? ` (${cSch.locationId})` : ''));
  }

  const seatsUsed = parseInt(sch.seatsUsed) || 0;
  const capacity  = parseInt(sch.capacity)  || 6;
  const isWaitlisted = seatsUsed + seatsHeld > capacity;

  let waitlistPosition = null;
  if (isWaitlisted) {
    const wlCount = allParts.filter(r => r.scheduleId === p.scheduleId && r.status === 'WAITLISTED').length;
    waitlistPosition = wlCount + 1;
  }

  const pid = generateId('PRT');
  partSheet.appendRow([
    pid,
    p.scheduleId,
    name,
    companion,
    seatsHeld,
    isWaitlisted ? 'WAITLISTED' : 'ACTIVE',
    isWaitlisted ? waitlistPosition : '',
    nowIso(),
    '', // cancelledAt
    '', // promotedAt
    p.contactNote || '',
  ]);

  if (!isWaitlisted) {
    updateRowField(schSheet, schRowIdx, 'seatsUsed', seatsUsed + seatsHeld);
  }

  writeAuditLog('JOIN', name, 'PARTICIPANT', pid, { scheduleId: p.scheduleId, seatsHeld, status: isWaitlisted ? 'WAITLISTED' : 'ACTIVE' });
  return { participantId: pid, status: isWaitlisted ? 'WAITLISTED' : 'ACTIVE', waitlistPosition };
}

function cancelParticipation(p) {
  const name = String(p.name || '').trim();
  if (!p.participantId) throw new Error('participantId is required');
  if (!name) throw new Error('name is required for verification');

  const partSheet = getSheet(SHEETS.PARTICIPANTS);
  const schSheet  = getSheet(SHEETS.SCHEDULES);

  const partRowIdx = findRowIndex(partSheet, 'participantId', p.participantId);
  if (partRowIdx === -1) throw new Error('Registration not found');
  const part = getRowObject(partSheet, partRowIdx);

  if (String(part.name).toLowerCase() !== name.toLowerCase()) throw new Error('Name does not match registration');
  if (part.status === 'CANCELLED') throw new Error('Registration is already cancelled');

  updateRowField(partSheet, partRowIdx, 'status',      'CANCELLED');
  updateRowField(partSheet, partRowIdx, 'cancelledAt', nowIso());

  const schRowIdx = findRowIndex(schSheet, 'scheduleId', part.scheduleId);
  const promoted = [];

  if (schRowIdx !== -1 && (part.status === 'ACTIVE' || part.status === 'PROMOTED')) {
    const sch = getRowObject(schSheet, schRowIdx);
    const newSeatsUsed = Math.max(0, (parseInt(sch.seatsUsed) || 0) - (parseInt(part.seatsHeld) || 1));
    updateRowField(schSheet, schRowIdx, 'seatsUsed', newSeatsUsed);

    // Promote waitlisted participants
    promoted.push(...promoteWaitlist(part.scheduleId, newSeatsUsed, parseInt(sch.capacity) || 6));
  }

  writeAuditLog('CANCEL', name, 'PARTICIPANT', p.participantId, { scheduleId: part.scheduleId, promoted });
  return { cancelled: true, promoted };
}

function promoteWaitlist(scheduleId, currentSeatsUsed, capacity) {
  const partSheet = getSheet(SHEETS.PARTICIPANTS);
  const schSheet  = getSheet(SHEETS.SCHEDULES);
  const allParts  = sheetToObjects(partSheet);

  const waitlisted = allParts
    .filter(r => r.scheduleId === scheduleId && r.status === 'WAITLISTED')
    .sort((a, b) => {
      const posA = parseInt(a.waitlistPosition) || 999;
      const posB = parseInt(b.waitlistPosition) || 999;
      return posA - posB;
    });

  const promoted = [];
  let seats = currentSeatsUsed;

  for (const wp of waitlisted) {
    const needed = parseInt(wp.seatsHeld) || 1;
    if (seats + needed <= capacity) {
      const rowIdx = findRowIndex(partSheet, 'participantId', wp.participantId);
      if (rowIdx !== -1) {
        updateRowField(partSheet, rowIdx, 'status',          'PROMOTED');
        updateRowField(partSheet, rowIdx, 'promotedAt',      nowIso());
        updateRowField(partSheet, rowIdx, 'waitlistPosition','');
        seats += needed;
        promoted.push(wp.name);
        writeAuditLog('PROMOTE_WAITLIST', wp.name, 'PARTICIPANT', wp.participantId, { scheduleId });
      }
    }
  }

  // Re-number remaining waitlist
  const stillWaiting = allParts
    .filter(r => r.scheduleId === scheduleId && r.status === 'WAITLISTED' && !promoted.includes(r.name))
    .sort((a, b) => (parseInt(a.waitlistPosition)||999) - (parseInt(b.waitlistPosition)||999));

  stillWaiting.forEach((wp, i) => {
    const rowIdx = findRowIndex(partSheet, 'participantId', wp.participantId);
    if (rowIdx !== -1) updateRowField(partSheet, rowIdx, 'waitlistPosition', i + 1);
  });

  // Update seatsUsed on schedule
  const schRowIdx = findRowIndex(schSheet, 'scheduleId', scheduleId);
  if (schRowIdx !== -1) updateRowField(schSheet, schRowIdx, 'seatsUsed', seats);

  return promoted;
}

// ─── Export ───────────────────────────────────────────────────────────────────

function exportReport(p) {
  validateAdminCode(p.adminCode);
  const type = p.reportType || 'schedules';

  let rows = [];

  if (type === 'schedules') {
    rows = [['Schedule ID','Location','Date','Time Slot','Capacity','Seats Used','Status','Created By','Created At','Cancelled At']];
    const scheds = getSchedules(p);
    scheds.forEach(s => rows.push([
      s.scheduleId, s.locationName, s.date, s.slotLabel,
      s.capacity, s.seatsUsed, s.status, s.createdBy, s.createdAt, s.cancelledAt,
    ]));
  } else if (type === 'participants') {
    rows = [['Participant ID','Schedule ID','Location','Date','Slot','Name','Companion','Seats Held','Status','Joined At','Cancelled At','Promoted At']];
    const partSheet = getSheet(SHEETS.PARTICIPANTS);
    const schSheet  = getSheet(SHEETS.SCHEDULES);
    const locSheet  = getSheet(SHEETS.LOCATIONS);
    const slotSheet = getSheet(SHEETS.TIMESLOTS);
    const schMap = {}; sheetToObjects(schSheet).forEach(s => { schMap[s.scheduleId] = s; });
    const locMap = {}; sheetToObjects(locSheet).forEach(l => { locMap[l.locationId] = l; });
    const slotMap= {}; sheetToObjects(slotSheet).forEach(s => { slotMap[s.slotId]   = s; });
    sheetToObjects(partSheet).forEach(pp => {
      const sch = schMap[pp.scheduleId] || {};
      const loc = locMap[sch.locationId] || {};
      const slot= slotMap[sch.slotId]   || {};
      if (p.dateFrom && String(sch.date) < p.dateFrom) return;
      if (p.dateTo   && String(sch.date) > p.dateTo)   return;
      rows.push([pp.participantId, pp.scheduleId, loc.name||'', sch.date||'', slot.label||'', pp.name, pp.companion, pp.seatsHeld, pp.status, pp.joinedAt, pp.cancelledAt, pp.promotedAt]);
    });
  } else if (type === 'waitlist') {
    rows = [['Participant ID','Name','Companion','Schedule ID','Location','Date','Slot','Waitlist Position','Joined At']];
    const partSheet = getSheet(SHEETS.PARTICIPANTS);
    const schSheet  = getSheet(SHEETS.SCHEDULES);
    const locSheet  = getSheet(SHEETS.LOCATIONS);
    const slotSheet = getSheet(SHEETS.TIMESLOTS);
    const schMap = {}; sheetToObjects(schSheet).forEach(s => { schMap[s.scheduleId] = s; });
    const locMap = {}; sheetToObjects(locSheet).forEach(l => { locMap[l.locationId] = l; });
    const slotMap= {}; sheetToObjects(slotSheet).forEach(s => { slotMap[s.slotId]   = s; });
    sheetToObjects(partSheet).filter(pp => pp.status === 'WAITLISTED').forEach(pp => {
      const sch  = schMap[pp.scheduleId] || {};
      const loc  = locMap[sch.locationId] || {};
      const slot = slotMap[sch.slotId]   || {};
      rows.push([pp.participantId, pp.name, pp.companion, pp.scheduleId, loc.name||'', sch.date||'', slot.label||'', pp.waitlistPosition, pp.joinedAt]);
    });
  }

  const csv = rows.map(r =>
    r.map(c => '"' + String(c == null ? '' : c).replace(/"/g, '""') + '"').join(',')
  ).join('\n');

  return { csvContent: csv };
}

// ─────────────────────────────────────────────────────────────────────────────
// SEEDING — Run seedData() once from the Apps Script editor after setup
// ─────────────────────────────────────────────────────────────────────────────

function seedData() {
  const ss = getSpreadsheet();

  function ensureSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1B4332').setFontColor('#ffffff');
    }
    return sheet;
  }

  ensureSheet(SHEETS.LOCATIONS,    ['locationId','name','address','capacity','mapsUrl','isActive','createdAt']);
  ensureSheet(SHEETS.TIMESLOTS,    ['slotId','label','startTime','endTime','isActive','sortOrder']);
  ensureSheet(SHEETS.SCHEDULES,    ['scheduleId','locationId','slotId','date','capacity','seatsUsed','status','createdBy','createdAt','cancelledAt','cancelledBy']);
  ensureSheet(SHEETS.PARTICIPANTS, ['participantId','scheduleId','name','companion','seatsHeld','status','waitlistPosition','joinedAt','cancelledAt','promotedAt','contactNote']);
  ensureSheet(SHEETS.AUDIT_LOG,    ['logId','timestamp','action','actorName','targetType','targetId','details']);

  const configSheet = ensureSheet(SHEETS.CONFIG, ['key','value','description']);
  if (configSheet.getLastRow() < 2) {
    configSheet.appendRow(['nearlyFullThreshold', '0.75', 'Fraction of capacity for "Nearly Full" status']);
    configSheet.appendRow(['defaultCapacity',     '6',    'Default seats per location']);
    configSheet.appendRow(['appTitle',             'Monumento Public Witnessing', 'Application display title']);
    configSheet.appendRow(['adminCode',            '2026Admin', 'Admin passcode — change this!']);
  }

  // Sample locations
  const locSheet = getSheet(SHEETS.LOCATIONS);
  if (locSheet.getLastRow() < 2) {
    locSheet.appendRow(['LOC-001', 'Riza Residence',  'Monumento Area', 6, '', true, nowIso()]);
    locSheet.appendRow(['LOC-002', 'Cielo Residence', 'Monumento Area', 6, '', true, nowIso()]);
    locSheet.appendRow(['LOC-003', 'DXN Office',      'Monumento Area', 6, '', true, nowIso()]);
  }

  // Sample time slots
  const slotSheet = getSheet(SHEETS.TIMESLOTS);
  if (slotSheet.getLastRow() < 2) {
    slotSheet.appendRow(['SLOT-001', 'Early Morning (6:00–7:00 AM)',  '06:00', '07:00', true, 1]);
    slotSheet.appendRow(['SLOT-002', 'Morning (7:00–8:00 AM)',         '07:00', '08:00', true, 2]);
    slotSheet.appendRow(['SLOT-003', 'Early Afternoon (4:00–5:00 PM)', '16:00', '17:00', true, 3]);
    slotSheet.appendRow(['SLOT-004', 'Afternoon (5:00–6:00 PM)',       '17:00', '18:00', true, 4]);
    slotSheet.appendRow(['SLOT-005', 'Early Evening (6:00–7:00 PM)',   '18:00', '19:00', true, 5]);
    slotSheet.appendRow(['SLOT-006', 'Evening (7:00–8:00 PM)',         '19:00', '20:00', true, 6]);
  }

  Logger.log('✅ Seed complete for Monumento Public Witnessing. Deploy as Web App and paste the URL into index.html CONFIG.API_URL.');
}
