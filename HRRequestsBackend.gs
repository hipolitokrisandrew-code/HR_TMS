/*******************************************************
 * HRRequest.gs — Backend for HR Requests tab
 *
 * Key behavior:
 *  - "Details" is written only by the Request Form (TMS).
 *    HR Request treats it as read-only.
 *  - HR Request can update: Service, Process Step, Remarks,
 *    Start / Pause / Resume / End, Status, TAT fields.
 *  - Request Date, Due Date, Details, Attachments, Remarks
 *    are all stored in the TMS sheet and read from there.
 *
 * This version fixes:
 *  - Start timestamp persists in the UI (via canonicalized
 *    log data that always includes Start).
 *  - Due Date is guaranteed to be present after Start
 *    (either from Request Form or computed at Start using
 *    existing SLA logic helper computeDueDateForService_).
 *  - Request Date & Due Date are always formatted as
 *    "MM-dd-yyyy HH:mm:ss" in the HR Request table,
 *    regardless of sheet formatting.
 *
 * Relies on shared utilities in your project:
 *  - CONFIG, getCompany(), companyFromRequestId_(),
 *    getCompanySpreadsheet_()
 *  - getOnwardSs(), getItamSs(), getIrealSs(), getLatteSs()
 *  - getAdminSs(), getSheetByAnyName_(), readSheetAsObjects()
 *  * - requireAuth()
 *******************************************************/

/* =====================================================
 * Canonical header aliases (for inconsistent column names)
 * ===================================================== */
const __TMS_ALIASES__ = Object.freeze({
  "Start"            : ["Start", "Start Time", "Started", "Time Start", "Date Start"],
  "Pause"            : ["Pause", "Paused", "Pause Time", "Time Pause"],
  "Resume"           : ["Resume", "Resumed", "Resume Time", "Time Resume"],
  "End"              : ["End", "End Time", "Closed", "Completed", "Time End", "Date End"],
  "Process Step"     : ["Process Step", "Specific", "Step"],
  "Details"          : ["Details", "Specific", "Description", "Notes", "Form Details"],
  "Remarks"          : ["Remarks", "HR Remarks", "Handler Remarks", "Follow-up Notes"],
  "Request Date"     : ["Request Date", "Date Requested", "Form Date"],
  "Due Date"         : ["Due Date", "Deadline", "Target Date"],
  "TAT (mins)"       : ["TAT (mins)", "TAT"],
  "Total TAT (mins)" : ["Total TAT (mins)", "Total TAT", "TAT Total"]
});

/**
 * Adds canonical keys (Start, Due Date, etc.) to a row object
 * based on header aliases, without removing original keys.
 */
function __canonicalizeTmsRow__(row) {
  if (!row) return row;
  const out = Object.assign({}, row);

  const pick = (canonKey) => {
    if (out[canonKey]) return;
    const aliases = __TMS_ALIASES__[canonKey] || [];
    for (let i = 0; i < aliases.length; i++) {
      const k = aliases[i];
      if (Object.prototype.hasOwnProperty.call(row, k) &&
          String(row[k] || "").trim() !== "") {
        out[canonKey] = row[k];
        break;
      }
    }
  };

  Object.keys(__TMS_ALIASES__).forEach(pick);
  return out;
}

/* ==============================
 * Small helpers
 * ============================== */

function _normalizeId_(v) {
  return String(v == null ? "" : v).trim();
}

function _cleanUnique_(arr) {
  const seen = Object.create(null);
  const out  = [];
  (arr || []).forEach(v => {
    const s = String(v || "").trim();
    if (!s || seen[s]) return;
    seen[s] = true;
    out.push(s);
  });
  return out;
}

function _listMatrixValues_() {
  return getListSheet_().getDataRange().getDisplayValues();
}

/* ==============================
 * List / Services / Process Steps
 * ============================== */

function getListSheet_() {
  const sh = getSheetByAnyName_(getAdminSs(), [CONFIG.SHEETS.LIST, "LIST", "list"]);
  if (!sh) {
    throw new Error('List sheet not found. Create "List" with headers: Service | Process Step');
  }
  return sh;
}

function getServicesBackend() {
  requireAuth();
  const rows = readSheetAsObjects(getListSheet_());
  const fromRows = _cleanUnique_(
    rows
      .map(r => r["Service"] || r["service"])
      .filter(Boolean)
  );
  if (fromRows.length) return fromRows;

  const vals   = _listMatrixValues_();
  const header = (vals[0] || []);
  return _cleanUnique_(header);
}

function getProcessStepsBackend(service) {
  requireAuth();
  if (!service) return [];
  const listRows = readSheetAsObjects(getListSheet_());
  const stepsFromRows = listRows
    .filter(r => String(r["Service"] ?? r["service"] ?? "").trim() === String(service).trim())
    .map(r => r["Process Step"] ?? r["process step"])
    .filter(Boolean);
  if (stepsFromRows.length) return _cleanUnique_(stepsFromRows);

  const vals = _listMatrixValues_();
  if (!vals.length) return [];
  const header = vals[0] || [];
  const want   = String(service).trim().toLowerCase();
  let col = -1;
  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] || "").trim().toLowerCase();
    if (h && h === want) {
      col = c;
      break;
    }
  }
  if (col === -1) return [];
  const steps = [];
  for (let r = 1; r < vals.length; r++) {
    const cell = String((vals[r] && vals[r][col]) || "").trim();
    if (cell) steps.push(cell);
  }
  return _cleanUnique_(steps);
}

/* ==========================================
 * Unified multi-company log data (source for
 * HR Request table and Request ID dropdowns)
 * ========================================== */

function fromCompany_(ss, companyLabel) {
  if (!ss) return [];
  const sh = ss.getSheetByName(CONFIG.SHEETS.TMS);
  if (!sh) return [];

  // Read both raw values and display values
  const range         = sh.getDataRange();
  const displayValues = range.getDisplayValues();
  const rawValues     = range.getValues();
  if (!displayValues.length) return [];
  const header      = displayValues[0] || [];
  const headerLower = header.map(h => String(h || "").trim().toLowerCase());

  // Helper: find index of a canonical key using alias mapping
  function findAliasIdx(canonKey) {
    const aliases = (__TMS_ALIASES__ && __TMS_ALIASES__[canonKey]) || [];
    for (let i = 0; i < aliases.length; i++) {
      const t = String(aliases[i] || "").trim().toLowerCase();
      const idx = headerLower.indexOf(t);
      if (idx !== -1) return idx;
    }
    return -1;
  }

  const reqDateIdx = findAliasIdx("Request Date");
  const dueDateIdx = findAliasIdx("Due Date");

  const out = [];

  for (let r = 1; r < displayValues.length; r++) {
    const rowVals = displayValues[r] || [];
    const rawRow  = rawValues[r]     || [];
    const obj = {};

    for (let c = 0; c < header.length; c++) {
      const keyRaw = header[c];
      const key    = String(keyRaw || "").trim();
      if (!key) continue;

      let val = rowVals[c];

      // Force Request Date & Due Date to "MM-dd-yyyy HH:mm:ss"
      if (c === reqDateIdx || c === dueDateIdx) {
        const raw = rawRow[c];
        if (raw instanceof Date && !isNaN(raw)) {
          val = Utilities.formatDate(raw, CONFIG.TIMEZONE, "MM-dd-yyyy HH:mm:ss");
        }
      }

      obj[key] = val;
    }

    // Tag company if not already present
    if (!obj["Company"]) obj["Company"] = companyLabel;
    out.push(__canonicalizeTmsRow__(obj));
  }
  return out;
}

function __buildLogDataBackendCore__(session) {
  // session is expected from requireAuth(); tolerate missing to stay fail-open for legacy callers
  const safeSession = session || {};

  let rows = []
    .concat(fromCompany_(getOnwardSs(), "Onward"))
    .concat(fromCompany_(getItamSs(),   "ITAM"))
    .concat(fromCompany_(getIrealSs(),  "IREAL"))
    .concat(fromCompany_(getLatteSs(),  "LATTE"));

  // De-duplicate rows by Company + Request ID + Start/End/Status
  const seen = Object.create(null);
  rows = rows.filter(r => {
    const key = [
      String(r["Company"] || "").toUpperCase(),
      _normalizeId_(r["Request ID"] || r["request id"]),
      String(r["Start"] || r["start"] || "").trim(),
      String(r["End"] || r["end"] || "").trim(),
      String(r["Status"] || r["status"] || "").trim().toLowerCase()
    ].join("||");
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });

  try {
    const role = (safeSession && safeSession.role || "").toLowerCase();
    const email = (safeSession && safeSession.email || "").toLowerCase();
    const dept = (safeSession && safeSession.department || "").toLowerCase();

    if (role === "employee") {
      const emailKey = rows.length ? Object.keys(rows[0]).find(k => k.toLowerCase().includes("email")) : null;
      if (emailKey) {
        rows = rows.filter(r => String(r[emailKey] || "").toLowerCase() === email);
      }
    } else if (role === "department head") {
      rows = rows.filter(r => (r["Department"] || "").toLowerCase() === dept);
    }
  } catch (e) {
    // fail-open on any session error
  }

  return rows;
}

function getLogDataBackend(enforceAccess) {
  const session = requireAuth();
  const shouldEnforce = enforceAccess !== false;
  if (shouldEnforce) {
    requireRoleAccess_('tab-requests');
  }

  return __buildLogDataBackendCore__(session);
}

function getFilteredLogDataBackend(filters) {
  requireRoleAccess_('tab-requests');
  const rows = getLogDataBackend();
  const f = filters || {};

  return rows.filter(r => {
    const company    = (r["Company"]      || "").toString();
    const service    = (r["Service"]      || "").toString();
    const step       = (r["Process Step"] || "").toString();
    const details    = (r["Details"]      || "").toString();
    const department = (r["Department"]   || "").toString();
    const status     = (r["Status"]       || "").toString();
    const requestId  = (r["Request ID"]   || "").toString();

    return (!f.company     || company   === f.company)
        && (!f.service     || service.toLowerCase().indexOf(String(f.service).toLowerCase()) !== -1)
        && (!f.processStep || step.toLowerCase().indexOf(String(f.processStep).toLowerCase()) !== -1)
        && (!f.details     || details.toLowerCase().indexOf(String(f.details).toLowerCase()) !== -1)
        && (!f.department  || department.toLowerCase().indexOf(String(f.department).toLowerCase()) !== -1)
        && (!f.status      || status.toLowerCase().indexOf(String(f.status).toLowerCase()) !== -1)
        && (!f.requestId   || requestId.toLowerCase().indexOf(String(f.requestId).toLowerCase()) !== -1);
  });
}

/* ==========================================
 * Row lookup by Request ID (single-row ops)
 * ========================================== */

function findRowInfoByRequestId_(sheet, requestId) {
  const id = _normalizeId_(requestId);
  if (!id) return { rowIndex: -1, status: "", startHasValue: false };

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { rowIndex: -1, status: "", startHasValue: false };
  }

  const header = values[0].map(h => String(h).trim());
  const idIdx     = header.findIndex(h => h.toLowerCase() === "request id");
  const statusIdx = header.findIndex(h => h.toLowerCase() === "status");
  const startIdx  = header.findIndex(h => h.toLowerCase() === "start");

  if (idIdx < 0)     throw new Error("Request ID column not found.");
  if (statusIdx < 0) throw new Error("Status column not found.");
  if (startIdx < 0)  throw new Error("Start column not found.");

  for (let r = 1; r < values.length; r++) {
    const cellId = _normalizeId_(values[r][idIdx]);
    if (cellId === id) {
      const status        = String(values[r][statusIdx] || "").trim();
      const start         = values[r][startIdx];
      const startHasValue = (start instanceof Date) || (String(start || "").trim() !== "");
      return { rowIndex: r + 1, status: status, startHasValue: startHasValue };
    }
  }

  return { rowIndex: -1, status: "", startHasValue: false };
}

/* ==========================================
 * Core action logic: Start / Pause / Resume / End
 * ========================================== */

function logActionBackend(action, requestId, service, processStep, remarks) {
  requireRoleAccess_('tab-requests');

  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    hasLock = lock.tryLock(5000);
    if (!hasLock) {
      lock.waitLock(5000);
      hasLock = true;
    }

    const id = _normalizeId_(requestId);
    if (!id) throw new Error("Request ID is required.");

    // Validate company against Request ID prefix
    const idCompany = companyFromRequestId_(id); // Onward / ITAM / IREAL / LATTE
    const selectedCompany = getCompany();

    if (!idCompany) {
      const msgNotFound = {
        "Start" : "Cannot start: Request ID not found.",
        "Pause" : "Cannot pause: Request ID not found.",
        "Resume": "Cannot resume: Request ID not found.",
        "End"   : "Cannot end: Request ID not found."
      }[action] || "Request ID not found.";
      throw new Error(msgNotFound);
    }

    if (selectedCompany && idCompany !== selectedCompany) {
      const msgMismatch = {
        "Start" : "Cannot start: Selected company does not match the Request ID.",
        "Pause" : "Cannot pause: Selected company does not match the Request ID.",
        "Resume": "Cannot resume: Selected company does not match the Request ID.",
        "End"   : "Cannot end: Selected company does not match the Request ID."
      }[action] || "Company mismatch.";
      throw new Error(msgMismatch);
    }

    const ss    = getCompanySpreadsheet_(idCompany);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.TMS);
    if (!sheet) throw new Error("TMS sheet not found for this Request ID.");

    const hdrs  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const lower = hdrs.map(h => h.toLowerCase());

    function idx(name) {
      const i = lower.indexOf(String(name).toLowerCase());
      if (i < 0) throw new Error("Column not found: " + name);
      return i + 1;
    }

    function idxOrNull(names) {
      for (let i = 0; i < names.length; i++) {
        const n = String(names[i]).toLowerCase();
        const j = lower.indexOf(n);
        if (j >= 0) return j + 1;
      }
      return null;
    }

    const col = {
      requestId   : idx("request id"),
      service     : idx("service"),
      step        : idxOrNull(["process step", "specific"]),
      details     : idxOrNull(["details", "specific", "description", "notes", "form details"]), // READ-ONLY
      remarks     : idxOrNull(["remarks", "hr remarks", "handler remarks", "follow-up notes"]),
      start       : idx("start"),
      pause       : idx("pause"),
      resume      : idx("resume"),
      end         : idx("end"),
      tat         : idxOrNull(["tat (mins)", "tat"]),
      totalTat    : idxOrNull(["total tat (mins)", "total tat", "tat total"]),
      status      : idx("status"),
      requestDate : idxOrNull(["request date", "date requested", "form date"]),
      dueDate     : idxOrNull(["due date", "deadline", "target date"])
    };

    const TS_FORMAT = "MM-dd-yyyy HH:mm:ss";
    const now       = new Date();
    const nowStr    = Utilities.formatDate(now, CONFIG.TIMEZONE, TS_FORMAT);

    // Locate the row for this Request ID
    const info = findRowInfoByRequestId_(sheet, id);
    if (info.rowIndex === -1) {
      const msg2 = {
        "Start" : "Cannot start: Request ID not found.",
        "Pause" : "Cannot pause: Request ID not found.",
        "Resume": "Cannot resume: Request ID not found.",
        "End"   : "Cannot end: Request ID not found."
      }[action] || "Request ID not found.";
      throw new Error(msg2);
    }
    const r = info.rowIndex;

    function getDate(c) {
      if (!c) return null;
      const v = sheet.getRange(r, c).getValue();
      return (v instanceof Date && !isNaN(v)) ? v : null;
    }

    const startDt  = getDate(col.start);
    const pauseDt  = getDate(col.pause);
    const resumeDt = getDate(col.resume);
    const endDt    = getDate(col.end);

    const statusLower    = String(info.status || "").trim().toLowerCase();
    const isPausedState  = (statusLower === "paused");
    const isInProgress   = statusLower === "in progress" ||
                           statusLower === "resumed" ||
                           (!!startDt && !isPausedState && !endDt);
    const isStartedAtAll = !!startDt;

    let tatSoFar = 0;
    if (col.tat) {
      const rawTat = sheet.getRange(r, col.tat).getDisplayValue();
      tatSoFar = parseFloat(String(rawTat || "0").replace(/,/g, "")) || 0;
    }

    // Action-specific guards
    if (action === "Start" && isStartedAtAll) {
      throw new Error("Cannot start: Request already started.");
    }
    if (action === "Pause" && !isInProgress) {
      throw new Error("Cannot pause: Request is not currently started.");
    }
    if (action === "Resume" && !isPausedState) {
      throw new Error("Cannot resume: Request is not currently paused.");
    }
    if (action === "End" && !(isInProgress || isPausedState)) {
      throw new Error("Cannot end: Request is not in progress.");
    }

    // Update descriptive fields (Service, Process Step, Remarks)
    sheet.getRange(r, col.service).setValue(service);
    if (col.step)    sheet.getRange(r, col.step).setValue(processStep);
    if (col.remarks) sheet.getRange(r, col.remarks).setValue(remarks);

    // Helper: accumulate active window into TAT
    function addActiveWindowToTat(anchorDate, stopDate) {
      if (anchorDate && stopDate && stopDate >= anchorDate) {
        const delta = (stopDate.getTime() - anchorDate.getTime()) / 60000;
        tatSoFar = Math.max(0, tatSoFar + delta);
      }
    }

    // START
    if (action === "Start") {
      // Set Start timestamp
      sheet.getRange(r, col.start).setValue(now).setNumberFormat(TS_FORMAT);

      // Ensure Request Date exists if column is present but empty
      if (col.requestDate && !getDate(col.requestDate)) {
        sheet.getRange(r, col.requestDate).setValue(now).setNumberFormat(TS_FORMAT);
      }

      // If Due Date is still blank, compute using SLA helper (if available)
      if (col.dueDate && !getDate(col.dueDate) && typeof computeDueDateForService_ === "function") {
        const reqDate = getDate(col.requestDate) || now;
        const due = computeDueDateForService_(service, processStep, reqDate);
        if (due instanceof Date && !isNaN(due)) {
          sheet.getRange(r, col.dueDate).setValue(due).setNumberFormat(TS_FORMAT);
        }
      }

      // Reset Pause/Resume/End and TAT
      if (col.pause)  sheet.getRange(r, col.pause).clearContent();
      if (col.resume) sheet.getRange(r, col.resume).clearContent();
      if (col.end)    sheet.getRange(r, col.end).clearContent();
      tatSoFar = 0;
    }

    // PAUSE
    if (action === "Pause") {
      const pauseAnchor = resumeDt || startDt;
      addActiveWindowToTat(pauseAnchor, now);
      sheet.getRange(r, col.pause).setValue(now).setNumberFormat(TS_FORMAT);
    }

    // RESUME
    if (action === "Resume") {
      sheet.getRange(r, col.resume).setValue(now).setNumberFormat(TS_FORMAT);
      // Next active window starts from this resume time
    }

    // END
    if (action === "End") {
      const endAnchor = resumeDt || startDt;
      addActiveWindowToTat(endAnchor, now);
      sheet.getRange(r, col.end).setValue(now).setNumberFormat(TS_FORMAT);
    }

    const nextStatus = ({
      "Start" : "In Progress",
      "Pause" : "Paused",
      "Resume": "Resumed",
      "End"   : "Completed"
    })[action] || "In Progress";

    const tatFixed = tatSoFar > 0 ? tatSoFar.toFixed(2) : "0.00";
    if (col.tat)      sheet.getRange(r, col.tat).setValue(tatFixed);
    if (col.totalTat) sheet.getRange(r, col.totalTat).setValue(tatFixed);
    sheet.getRange(r, col.status).setValue(nextStatus);

    SpreadsheetApp.flush();

    // Re-read the entire row with display values so the UI sees
    // Start / Due Date / Details / Remarks exactly as stored in TMS
    const rowValues = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getDisplayValues()[0] || [];
    const rowObj = {};
    hdrs.forEach((h, i) => {
      rowObj[h] = rowValues[i];
    });

    // Ensure we return canonical keys and updated fields
    rowObj["Service"]        = service;
    rowObj["Process Step"]   = processStep;
    if (col.remarks) rowObj["Remarks"] = remarks;
    // Details stays exactly as stored in TMS
    rowObj["TAT (mins)"]       = tatFixed;
    rowObj["Total TAT (mins)"] = tatFixed;
    rowObj["Status"]           = nextStatus;
    rowObj["Company"]          = idCompany;

    // In case formats haven't applied yet, guarantee a Start string
    if (action === "Start")  rowObj["Start"]  = rowObj["Start"]  || nowStr;
    if (action === "Pause")  rowObj["Pause"]  = rowObj["Pause"]  || nowStr;
    if (action === "Resume") rowObj["Resume"] = rowObj["Resume"] || nowStr;
    if (action === "End")    rowObj["End"]    = rowObj["End"]    || nowStr;

    return {
      ok     : true,
      message: 'Action "' + action + '" completed.',
      row    : __canonicalizeTmsRow__(rowObj)
    };
  } finally {
    try {
      if (hasLock) lock.releaseLock();
    } catch (e) {}
  }
}

/* ==========================================
 * Active Requests + single-record info
 * ========================================== */

function getActiveRequestsBackend() {
  requireRoleAccess_('tab-requests');
  const rows = getLogDataBackend();
  const ACTIVE = {
    "open": true,
    "paused": true,
    "resumed": true,
    "in-progress": true,
    "in progress": true
  };

  const seen = Object.create(null);
  const out  = [];

  rows.forEach(orig => {
    const r   = __canonicalizeTmsRow__(orig);
    const id  = _normalizeId_(r["Request ID"] || r["request id"]);
    if (!id) return;

    const status = String(r["Status"] || "").trim();
    const sl     = status.toLowerCase();
    if (!ACTIVE[sl]) return;
    if (seen[id]) return;
    seen[id] = true;

    out.push({
      requestId   : id,
      company     : r["Company"]      || "",
      service     : r["Service"]      || "",
      processStep : r["Process Step"] || "",
      details     : r["Details"]      || "",
      remarks     : r["Remarks"]      || "",
      status      : status,
      requestDate : r["Request Date"] || "",
      dueDate     : r["Due Date"]     || ""
    });
  });

  return out;
}

function getRequestInfoBackend(requestId) {
  requireRoleAccess_('tab-requests');
  const id = _normalizeId_(requestId);
  if (!id) return null;

  const rows = getLogDataBackend();
  const hit = rows.find(r =>
    _normalizeId_(r["Request ID"] || r["request id"]) === id
  );
  if (!hit) return null;

  const c = __canonicalizeTmsRow__(hit);
  return {
    requestId   : id,
    company     : c["Company"]      || "",
    service     : c["Service"]      || "",
    processStep : c["Process Step"] || "",
    details     : c["Details"]      || "",
    remarks     : c["Remarks"]      || "",
    status      : c["Status"]       || "",
    requestDate : c["Request Date"] || "",
    dueDate     : c["Due Date"]     || ""
  };
}

/* ==========================================
 * Public functions used by HTML frontend
 * ========================================== */

function getLogData() {
  return getLogDataBackend();
}
function getFilteredLogData(filters) {
  return getFilteredLogDataBackend(filters);
}
function getServices() {
  return getServicesBackend();
}
function getProcessSteps(service) {
  return getProcessStepsBackend(service);
}
function logAction(action, requestId, service, processStep, remarks) {
  return logActionBackend(action, requestId, service, processStep, remarks);
}
function getActiveRequests() {
  return getActiveRequestsBackend();
}
function getRequestInfo(requestId) {
  return getRequestInfoBackend(requestId);
}
