/*******************************************************
 * EmployeeConcernsBackend.gs — Employee Concerns (TMS-based)
 *
 * Changes vs old version:
 *  - Completely stops using CONFIG.SHEETS.CONCERNS.
 *  - Reads only from the TMS sheet (same data as HR Request).
 *  - Uses the same row structure as getLogDataBackend() / HRRequest.
 *  - Derives Company from the first 3 chars of Request ID:
 *      ITM → ITAM
 *      ONW → Onward
 *      IRL → IREAL
 *      LTE → LATTE
 *  - Filters rows by active company from getCompany()/setCompany()
 *    using the derived Company (not any old Concerns sheet).
 *  - Returns headers + rows that support:
 *      • 8-column limited view (handled in frontend)
 *      • Full TMS “Show all records” view
 *
 * Exposed functions:
 *  - getConcernsDataActiveCompanyBackend(opts)  // core backend
 *  - getConcernsDataActiveCompany(opts)         // wrapper for google.script.run
 *  - getConcernsDataBackend()                   // compatibility (rows only)
 *
 * Relies on shared utilities already present in your project:
 *  - CONFIG, CONFIG.SHEETS.TMS, CONFIG.TIMEZONE
 *  - requireAuth()
 *  - getCompany()
 *  - getLogDataBackend()   // from HRRequest.gs
 *  - getCompanySpreadsheet_()
 *  - _normalizeId_()       // from HRRequest.gs
 *******************************************************/

/**
 * Derive Company from Request ID prefix.
 * Mapping (spec):
 *   ITM → ITAM
 *   ONW → ONWARD
 *   IRL → IREAL
 *   LTE → LATTE
 *
 * NOTE:
 *  - For filtering we must match the existing company selector values:
 *      Onward, ITAM, IREAL, LATTE
 *  - So we map ONW → "Onward" (capital O, rest lower case),
 *    and keep the others as ITAM / IREAL / LATTE.
 */
function EC_deriveCompanyFromRequestId_(requestId) {
  const id = (typeof _normalizeId_ === 'function')
    ? _normalizeId_(requestId)
    : String(requestId == null ? "" : requestId).trim();

  if (!id) return '';

  const prefix = id.substring(0, 3).toUpperCase();
  switch (prefix) {
    case 'ITM': return 'ITAM';    // ITM → ITAM
    case 'ONW': return 'Onward';  // ONW → ONWARD (UI uses "Onward")
    case 'IRL': return 'IREAL';   // IRL → IREAL
    case 'LTE': return 'LATTE';   // LTE → LATTE
    default:    return '';
  }
}

/**
 * Read the raw header row from the TMS sheet for a given company.
 * Returns the *sheet headers only* (no synthetic "Company" column).
 */
function EC_getTmsHeadersForCompany_(company) {
  try {
    if (!company) return [];
    if (typeof getCompanySpreadsheet_ !== 'function') return [];

    const ss = getCompanySpreadsheet_(company);
    if (!ss) return [];

    const sh = ss.getSheetByName(CONFIG.SHEETS.TMS);
    if (!sh) return [];

    const vals = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues();
    if (!vals || !vals.length) return [];

    const header = vals[0] || [];
    return header
      .map(function (h) { return String(h || '').trim(); })
      .filter(function (h) { return h; });
  } catch (e) {
    // Fail-open: if anything goes wrong, let caller fall back to row keys.
    return [];
  }
}

/**
 * Core data provider for Employee Concerns.
 *
 * Returns:
 *   {
 *     company : <active company from getCompany()>,
 *     headers : ['Company', <raw TMS headers...>],
 *     rows    : [ { Company:..., <TMS header>:value, ... }, ... ]
 *   }
 *
 * Notes:
 *  - Uses getLogDataBackend() so structure matches HRRequest.
 *  - Company is *derived from Request ID* and then used for:
 *      • filter by getCompany()
 *      • Company column in the UI
 *  - showAll + limit are kept for compatibility, but the
 *    "Show all records" toggle is handled purely in the frontend
 *    (columns, not rows).
 */
function getConcernsDataActiveCompanyBackend(opts) {
  requireAuth();

  const showAllFlag = !!(opts && opts.showAll);
  const limit = (opts && Number.isFinite(opts.limit) && opts.limit > 0)
    ? Math.floor(opts.limit)
    : null;

  const activeCompany = (typeof getCompany === 'function')
    ? getCompany()
    : '';

  // Reuse unified TMS log data (same as HR Request table).
  var rows = (typeof getLogDataBackend === 'function')
    ? getLogDataBackend()
    : [];

  if (!Array.isArray(rows) || !rows.length) {
    return { company: activeCompany, headers: [], rows: [] };
  }

  // 1) Derive Company from Request ID prefix.
  rows = rows.map(function (original) {
    var r = Object.assign({}, original);
    var rawId =
      r['Request ID']   ||
      r['request id']   ||
      r['REQUEST ID']   ||
      '';

    var derived = EC_deriveCompanyFromRequestId_(rawId);
    if (derived) {
      r['Company'] = derived;
    } else if (!r['Company']) {
      // Ensure the key exists (even if empty) for consistent UI behavior.
      r['Company'] = '';
    }
    return r;
  });

  // 2) Filter by active company from getCompany(), using the *derived* Company.
  var companyFilter = activeCompany && activeCompany !== 'All'
    ? String(activeCompany)
    : '';

  if (companyFilter) {
    rows = rows.filter(function (r) {
      return String(r['Company'] || '').trim() === companyFilter;
    });
  }

  if (!rows.length) {
    return { company: activeCompany, headers: [], rows: [] };
  }

  // 3) Build headers: synthetic "Company" + raw TMS headers.
  //    Prefer reading headers directly from the TMS sheet for the active company.
  var tmsHeaders = EC_getTmsHeadersForCompany_(
    companyFilter || rows[0]['Company'] || ''
  );

  // Fallback: use keys from the first row if we can't read from the sheet.
  if (!tmsHeaders.length) {
    tmsHeaders = Object.keys(rows[0] || {}).filter(function (k) {
      return k && k !== 'Company';
    });
  }

  // Final header list for "show all" view:
  //   [ 'Company', <TMS headers except any existing Company column> ]
  var headers = ['Company'].concat(
    tmsHeaders.filter(function (h) {
      return String(h || '').trim().toLowerCase() !== 'company';
    })
  );

  // 4) Normalize rows so they expose exactly the headers above
  //    (no extra keys; safe for KPI / downstream code).
  var normalized = rows.map(function (r) {
    var o = {};
    headers.forEach(function (h) {
      if (h === 'Company') {
        o[h] = r['Company'] || '';
      } else if (Object.prototype.hasOwnProperty.call(r, h)) {
        o[h] = r[h] == null ? '' : r[h];
      } else {
        o[h] = '';
      }
    });
    return o;
  });

  // 5) Optional row limit (kept for compatibility).
  if (!showAllFlag && limit != null) {
    normalized = normalized.slice(0, limit);
  }

  return {
    company: activeCompany,
    headers: headers,
    rows: normalized
  };
}

/**
 * Wrapper so the frontend can call:
 *   google.script.run.getConcernsDataActiveCompany(...)
 *
 * This simply delegates to the backend function above.
 */
function getConcernsDataActiveCompany(opts) {
  return getConcernsDataActiveCompanyBackend(opts);
}

/**
 * Compatibility helper (used by bootstrap/KPI code).
 * Returns *rows only* for active company, with full TMS columns.
 */
function getConcernsDataBackend() {
  var pack = getConcernsDataActiveCompanyBackend({ showAll: true });
  return pack.rows || [];
}
