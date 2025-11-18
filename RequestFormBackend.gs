/*******************************************************
 * RequestForm.gs — Request form submission backend
 *
 * New behavior:
 *  - "Request Date" is set at creation (exact submission
 *    timestamp) in the TMS sheet.
 *  - "Due Date" is computed based on Service/Process Step
 *    using SLA rules (where mapping exists).
 *  - For configured Services, Due Date is always attempted
 *    via SLA logic and logged if missing.
 *  - "Attachments" are saved as Drive URLs in TMS.
 *
 * Existing behavior/output is preserved for all other fields.
 *******************************************************/

function submitFormBackend(dto, files) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);

  try {
    const s = getSessionInfoForClient();

    const accountCode = (typeof _detectAccountCodeForEmail_ === 'function')
      ? (_detectAccountCodeForEmail_(s && s.email) || s.accountCode || '')
      : (s.accountCode || '');

    // Unique ID against correct company code
    const requestId = generateRequestId_(s.companyCode, accountCode);

    // Company inferred from Request ID prefix
    const inferredCompany = companyFromRequestId_(requestId)
      || (typeof companyNameFromCode_ === 'function' ? companyNameFromCode_(s.companyCode) : null)
      || s.company || s.companyName || 'Onward';

    // Accept either dto.processStep (new) or dto.specific (legacy)
    const processStepVal = (dto && (dto.processStep || dto.specific)) || '';
    const when = new Date(); // server timestamp; Request Date uses this

    const row = {
      'Timestamp'   : when,
      'Request ID'  : requestId,
      'Company'     : inferredCompany,
      'Company Code': s.companyCode || '',
      'Account Code': accountCode || '',
      'Display Name': (dto && dto.name)  || s.displayName || '',
      'Email'       : (dto && dto.email) || s.email || '',
      'Employee ID' : (dto && dto.empId) || '',
      'Service'     : (dto && dto.service) || '',
      'Process Step': processStepVal,
      'Details'     : (dto && dto.details) || '',
      'Status'      : 'Open'
    };

    // Request Date: original submission date/time (single source of truth)
    row['Request Date'] = Utilities.formatDate(
      when,
      Session.getScriptTimeZone(),
      "MM-dd-yyyy HH:mm:ss"
    );

    // ------------------ DUE DATE HANDLING ------------------
    // Always attempt SLA-based Due Date computation for the configured Services.
    const serviceNameForSla = row['Service'];
    const stepForSla        = row['Process Step'];

    const dueRaw = computeDueDateForService_(serviceNameForSla, stepForSla, when);

    if (serviceRequiresDueDate_(serviceNameForSla)) {
      // These Services are required to have a Due Date via SLA logic
      if (dueRaw instanceof Date && !isNaN(dueRaw)) {
        row['Due Date'] = Utilities.formatDate(
          dueRaw,
          Session.getScriptTimeZone(),
          "MM-dd-yyyy HH:mm:ss"
        );
      } else {
        // Log when we *expected* a Due Date but SLA logic could not produce one
        logSlaDebug_('request_form_due_date_missing_required', {
          source     : 'submitFormBackend',
          service    : serviceNameForSla,
          processStep: stepForSla,
          requestDate: when.toString(),
          note       : 'computeDueDateForService_ returned null or invalid Date for a required Service'
        });
      }
    } else {
      // For other Services, keep original (best-effort) behavior
      if (dueRaw instanceof Date && !isNaN(dueRaw)) {
        row['Due Date'] = Utilities.formatDate(
          dueRaw,
          Session.getScriptTimeZone(),
          "MM-dd-yyyy HH:mm:ss"
        );
      }
    }
    // ---------------- END DUE DATE HANDLING ----------------

    // Optional attachments → stored in TMS as newline-separated links
    if (files && files.length) {
      const folder = getOrCreateUploadFolder_(inferredCompany, requestId);
      const links  = files.map(f => {
        const dataPart = (f.dataUrl || '').split(',')[1] || '';
        const blob = Utilities.newBlob(
          Utilities.base64Decode(dataPart),
          f.mimeType || 'application/octet-stream',
          f.name
        );
        const file = folder.createFile(blob)
          .setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
      });
      row['Attachments'] = links.join('\n');
    }

    // Append with header-safe helper to the correct company's TMS sheet
    appendToCompanyTms_(inferredCompany, row);

    return { ok: true, requestId, company: inferredCompany };
  } catch (err) {
    return { ok:false, message: String(err && err.message || err) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* Compatibility alias: frontend calls submitForm(dto, files) */
function submitForm(dto, files) {
  return submitFormBackend(dto, files);
}

// Optional helper to standardize SLA debug logs.
// - Only logs to Logger (no side effects on return values).
// - Safe against circular structures in payload.
function logSlaDebug_(event, data) {
  try {
    Logger.log(
      '[SLA DEBUG] %s :: %s',
      event,
      JSON.stringify(data || {})
    );
  } catch (e) {
    Logger.log('[SLA DEBUG] %s :: [payload not serializable]', event);
  }
}

/* -----------------------------------------------------------
 * Services that MUST always attempt a Due Date on submit
 * (exact strings must match the Service column values)
 * --------------------------------------------------------- */
const SLA_SERVICE_NAMES = [
  "Employee Relations",
  "Employee Concerns",
  "Timekeeping - Remind & Endorse",
  "Timekeeping - Missing Logs",
  "Timekeeping - Sick Leave",
  "Timekeeping - Vacation Leave",
  "Timekeeping - Emergency Leave",
  "Timekeeping - Overtime",
  "Timekeeping - Undertime",
  "Timekeeping - Official Business",
  "Timekeeping - Change Shift",
  "Separation - Voluntary",
  "Separation - Termination",
  "Separation - ENDO",
  "Separation - Retraction",
  "Movement - Regularization",
  "Movement - Promotion",
  "Movement - Level Up",
  "Movement - Salary Adjustment",
  "Movement - Force Transfer",
  "Training - Employee Orientation",
  "Training - Promotional Training",
  "Training - Level Up Training",
  "Training - Transfer Training",
  "Training - PIP Training",
  "Training - Training Materials",
  "Benefits - HMO",
  "Benefits - Leave Allocation",
  "Benefits - Leave Reports",
  "Clinic - Validation of Sickness Availments",
  "Clinic - Illness Report",
  "Clinic - Health & Wellness Program",
  "Performance Management",
  "Org - Org Development",
  "Org - Organizational Report",
  "Org - Compensation Report",
  "Payroll - Payroll Computation & Crediting (LOCAL)",
  "Payroll - Payroll Computation & Crediting (EXPAT)",
  "Payroll - Last Pay",
  "Payroll - Separation Pay",
  "Payroll - Tax Annualization",
  "SSS - New Hire Reporting",
  "SSS - Contributions",
  "SSS - Loan Remittances",
  "SSS - Unemployment Certification",
  "SSS - Certificate of Employee's Loan",
  "SSS - Sickness Notification",
  "SSS - Maternity Notification",
  "SSS - Disability Notification",
  "SSS - Annual Processing of L-501",
  "PHIC - New Hire Reporting",
  "PHIC - Contributions",
  "PHIC - CSF, CSI, MDR, Certificate of Contributions",
  "DOLE - Annual",
  "DOLE - Monthly (WAIR)",
  "DOLE - Termination Report",
  "DOLE - Quarterly (AEP Report)",
  "PAGIBIG - Certification of Employee's Loan",
  "PAGIBIG - Contributions",
  "PAGIBIG - Loan Remittances",
  "Records Management",
  "Recruitment",
  "Ad Hoc",
  "Approval"
];

function serviceRequiresDueDate_(serviceName) {
  if (!serviceName) return false;
  const s = String(serviceName).trim();
  return SLA_SERVICE_NAMES.indexOf(s) !== -1;
}

/* -----------------------------------------------------------
 * SLA Due Date Logic
 * -----------------------------------------------------------
 * Due Date is based on Service / Process Step using your
 * SLA descriptions. The mapping relies on the text labels
 * used in "Service" (and sometimes Process Step).
 *
 * NOTE:
 * - Returns a Date when an SLA rule matches.
 * - Returns null when no SLA rule applies.
 * - Additional logging added for debugging only.
 */

function computeDueDateForService_(service, processStep, requestDate) {
  // 🔎 Log the raw inputs on every call
  logSlaDebug_('computeDueDateForService_call', {
    service,
    processStep,
    requestDateType: Object.prototype.toString.call(requestDate),
    requestDateRaw: requestDate
  });

  if (!(requestDate instanceof Date) || isNaN(requestDate)) {
    // 🚨 Invalid or missing requestDate (not a valid Date)
    logSlaDebug_('invalid_request_date', {
      reason: 'requestDate is not a valid Date instance',
      service,
      processStep,
      requestDateType: Object.prototype.toString.call(requestDate),
      requestDateRaw: requestDate
    });
    return null;
  }

  const svcRaw = String(service || '');
  const svc    = svcRaw.toLowerCase();
  const svcKey = svc.replace(/\s+/g, ' ').trim(); // normalized, lower-cased
  const step   = String(processStep || '').toLowerCase();

  if (!svcKey) {
    // 🚨 Missing or empty service text → no SLA rule can match
    logSlaDebug_('missing_service_value', {
      reason: 'service is empty or undefined',
      service,
      processStep,
      requestDate: requestDate.toString()
    });
    return null;
  }

  const base = new Date(requestDate.getTime()); // copy

  // Helpers
  const addCalDays = (d, days) => {
    const x = new Date(d.getTime());
    x.setDate(x.getDate() + days);
    return x;
  };

  const isWeekend = d => d.getDay() === 0 || d.getDay() === 6;

  const addWorkingDays = (d, days) => {
    if (!Number.isFinite(days) || days === 0) return new Date(d.getTime());
    const dir = days > 0 ? 1 : -1;
    let remaining = Math.abs(days);
    let cur = new Date(d.getTime());
    while (remaining > 0) {
      cur.setDate(cur.getDate() + dir);
      if (!isWeekend(cur)) remaining--;
    }
    return cur;
  };

  const nextFixedMonthDay = (d, monthIndex, day) => {
    // monthIndex = 0..11
    const year = d.getFullYear();
    let target = new Date(year, monthIndex, day);
    if (target < d) {
      target = new Date(year + 1, monthIndex, day);
    }
    return target;
  };

  const lastDayOfMonth = (d) => {
    return new Date(d.getFullYear(), d.getMonth() + 1, 0);
  };

  const secondWorkingDayOfMonth = (year, month) => {
    let d = new Date(year, month, 1);
    let count = 0;
    while (true) {
      if (!isWeekend(d)) {
        count++;
        if (count === 2) return d;
      }
      d.setDate(d.getDate() + 1);
    }
  };

  const secondWorkingDayOfNextMonth = (d) => {
    const y = d.getFullYear();
    const m = d.getMonth();
    const nextMonth = (m + 1) % 12;
    const year = m === 11 ? y + 1 : y;
    return secondWorkingDayOfMonth(year, nextMonth);
  };

  const secondWorkingDayOfNextQuarter = (d) => {
    const y = d.getFullYear();
    const m = d.getMonth();
    const thisQuarter = Math.floor(m / 3); // 0..3
    const nextQuarter = (thisQuarter + 1) % 4;
    const year = nextQuarter === 0 ? y + 1 : y;
    const firstMonthOfQuarter = nextQuarter * 3;
    return secondWorkingDayOfMonth(year, firstMonthOfQuarter);
  };

  const nthWorkingDayOfMonth = (year, month, n) => {
    let d = new Date(year, month, 1);
    let count = 0;
    while (true) {
      if (!isWeekend(d)) {
        count++;
        if (count === n) return d;
      }
      d.setDate(d.getDate() + 1);
    }
  };

  const nthWorkingDayOfNextMonth = (d, n) => {
    const m = d.getMonth();
    const y = m === 11 ? d.getFullYear() + 1 : d.getFullYear();
    const month = (m + 1) % 12;
    return nthWorkingDayOfMonth(y, month, n);
  };

  const nextDayOfMonth = (d, day) => {
    const year = d.getFullYear();
    const month = d.getMonth();
    let target = new Date(year, month, day);
    if (target < d) {
      target = new Date(year, month + 1, day);
    }
    return target;
  };

  const nextDateFromList = (d, monthDayPairs) => {
    const baseYear = d.getFullYear();
    let best = null;

    function consider(y, m, day) {
      const candidate = new Date(y, m, day);
      if (candidate >= d && (!best || candidate < best)) {
        best = candidate;
      }
    }

    monthDayPairs.forEach(function (pair) {
      consider(baseYear,     pair[0], pair[1]);
      consider(baseYear + 1, pair[0], pair[1]);
    });

    return best;
  };

  const nextLevelUpDate = (d) => {
    // Level Up: every Dec 1 and June 1
    const year = d.getFullYear();
    const candidates = [
      new Date(year, 5,  1), // June 1
      new Date(year, 11, 1)  // Dec 1
    ];
    let best = null;
    for (let i = 0; i < candidates.length; i++) {
      const c = candidates[i];
      if (c >= d && (!best || c < best)) best = c;
    }
    if (!best) {
      // if both in current year already passed, pick next year June 1
      return new Date(year + 1, 5, 1);
    }
    return best;
  };

  const nextHealthWellnessDate = (d) => {
    // Health & Wellness Program: Every 2nd week of Feb, Jun, Sep, Dec.
    // Approximate as 8th of those months (start of 2nd week).
    const months = [1, 5, 8, 11]; // Feb, Jun, Sep, Dec
    const baseYear = d.getFullYear();
    let best = null;

    [baseYear, baseYear + 1].forEach(function (y) {
      months.forEach(function (m) {
        const candidate = new Date(y, m, 8);
        if (candidate >= d && (!best || candidate < best)) {
          best = candidate;
        }
      });
    });

    return best;
  };

  const nextAepQuarterDate = (d) => {
    // AEP Quarterly Report: Jan 31, Apr 30, Jul 31, Oct 31
    return nextDateFromList(d, [
      [0, 31],  // Jan 31
      [3, 30],  // Apr 30
      [6, 31],  // Jul 31
      [9, 31]   // Oct 31
    ]);
  };

  const nextDoleAnnualDate = (d) => {
    // DOLE Annual (AMR, AEDR, Annual Wages, 13th Month Pay Report)
    // Approx: choose the nearest of Jan 15, Jan 31, Mar 31, Jun 15
    return nextDateFromList(d, [
      [0, 15],  // Jan 15
      [0, 31],  // Jan 31
      [2, 31],  // Mar 31
      [5, 15]   // Jun 15
    ]);
  };

  const nextLocalPayrollDate = (d) => {
    // Local payroll: Every 2nd and 17th of the month
    const startYear = d.getFullYear();
    const startMonth = d.getMonth();
    let best = null;

    for (let offset = 0; offset <= 2; offset++) {
      const tmp = new Date(startYear, startMonth + offset, 1);
      const y = tmp.getFullYear();
      const m = tmp.getMonth();
      [2, 17].forEach(function (day) {
        const candidate = new Date(y, m, day);
        if (candidate >= d && (!best || candidate < best)) {
          best = candidate;
        }
      });
    }
    return best || new Date(startYear, startMonth, 2);
  };

  // 3. Employee Relations – 12 working days from incident
  if (svcKey.indexOf('employee relations') !== -1) {
    const due = addWorkingDays(base, 12);
    logSlaDebug_('sla_match_employee_relations', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'employee_relations_12_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  // 4. Employee Concerns – 1 or 2 working days
  if (svcKey.indexOf('employee concerns') !== -1) {
    // If process step mentions escalation / major / high, treat as major
    let days = 1;
    if (/major|escalat|high/.test(step)) days = 2;
    const due = addWorkingDays(base, days);
    logSlaDebug_('sla_match_employee_concerns', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: days === 1
        ? 'employee_concerns_minor_1_working_day'
        : 'employee_concerns_major_2_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  // Timekeeping-related services
  if (svcKey.indexOf('timekeeping - remind & endorse') !== -1) {
    // Reminder / endorse window → 1 working day from request
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_timekeeping_remind_endorse', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'timekeeping_remind_endorse_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - missing logs') !== -1) {
    // Missing Logs: within 1 working day from request
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_timekeeping_missing_logs', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'timekeeping_missing_logs_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - sick leave') !== -1) {
    // Sick Leave: within 1 working day upon return (requestDate assumed as return date)
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_timekeeping_sick_leave', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'timekeeping_sick_leave_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - vacation leave') !== -1) {
    // Vacation Leave: 5 days prior to intended date.
    // Here we treat requestDate as intended date and compute due 5 working days before.
    const due = addWorkingDays(base, -5);
    logSlaDebug_('sla_match_timekeeping_vacation_leave', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'timekeeping_vacation_leave_5_working_days_prior',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - emergency leave') !== -1) {
    // Emergency Leave: within 1 working day upon return
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_timekeeping_emergency_leave', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'timekeeping_emergency_leave_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - overtime') !== -1) {
    // Overtime:
    // - Planned: within 1 working day from request
    // - Unplanned: same day (approximation)
    let days = 1;
    if (/unplanned/i.test(step)) {
      days = 0;
    }
    const due = addWorkingDays(base, days);
    logSlaDebug_('sla_match_timekeeping_overtime', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: /unplanned/i.test(step)
        ? 'timekeeping_overtime_unplanned_same_day'
        : 'timekeeping_overtime_planned_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - undertime') !== -1) {
    // Undertime:
    // - Planned: 1 day prior (treated as 1 working day before base)
    // - Unplanned: within 1 working day after
    let days = 1;
    if (/planned/i.test(step)) {
      days = -1;
    } else if (/unplanned/i.test(step)) {
      days = 1;
    }
    const due = addWorkingDays(base, days);
    logSlaDebug_('sla_match_timekeeping_undertime', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: /planned/i.test(step)
        ? 'timekeeping_undertime_planned_1_working_day_prior'
        : (/unplanned/i.test(step)
          ? 'timekeeping_undertime_unplanned_1_working_day_after'
          : 'timekeeping_undertime_default_1_working_day_after'),
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - official business') !== -1) {
    // Official Business:
    // - Planned: 3 days prior
    // - Unplanned: within 1st hour upon return (approx same day)
    let days = 0;
    if (/planned/i.test(step)) {
      days = -3;
    }
    const due = addWorkingDays(base, days);
    logSlaDebug_('sla_match_timekeeping_official_business', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: /planned/i.test(step)
        ? 'timekeeping_official_business_planned_3_working_days_prior'
        : 'timekeeping_official_business_unplanned_same_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('timekeeping - change shift') !== -1) {
    // Change Shift: 3 days prior
    const due = addWorkingDays(base, -3);
    logSlaDebug_('sla_match_timekeeping_change_shift', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'timekeeping_change_shift_3_working_days_prior',
      dueDate: due.toString()
    });
    return due;
  }

  // Separation-related services
  if (svcKey.indexOf('separation - voluntary') !== -1) {
    // Voluntary: 30 days (Rank & File) or 60 days (Managerial) from resignation letter
    let days = 30;
    if (/manager/i.test(step)) days = 60;
    const due = addCalDays(base, days);
    logSlaDebug_('sla_match_separation_voluntary', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: days === 60
        ? 'separation_voluntary_managerial_60_calendar_days'
        : 'separation_voluntary_rank_and_file_30_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('separation - termination') !== -1) {
    // Termination:
    // - Authorized: 30 calendar days
    // - Just causes: 7 calendar days
    let days = 30;
    if (/just/i.test(step)) {
      days = 7;
    }
    const due = addCalDays(base, days);
    logSlaDebug_('sla_match_separation_termination', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: /just/i.test(step)
        ? 'separation_termination_just_7_calendar_days'
        : 'separation_termination_authorized_30_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

if (svcKey.indexOf('separation - endo') !== -1) {
  // ENDO (Independent Contractor / Probationary)
  // Approximation:
  //  - Probationary: 7 calendar days from request
  //  - Independent Contractor: 30 calendar days from request
  let days = 30;
  if (/probationary|probation/i.test(step)) {
    days = 7;
  }
  const due = addCalDays(base, days);
  logSlaDebug_('sla_match_separation_endo', {
    service,
    processStep,
    requestDate: requestDate.toString(),
    rule: days === 7
      ? 'separation_endo_probationary_7_calendar_days'
      : 'separation_endo_independent_30_calendar_days',
    dueDate: due.toString()
  });
  return due;
}

  if (svcKey.indexOf('separation - retraction') !== -1) {
    // Retraction: 5 working days from effective date / 1 day prior
    const due = addWorkingDays(base, 5);
    logSlaDebug_('sla_match_separation_retraction', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'separation_retraction_5_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  // Movement services
  if (svcKey.indexOf('movement - regularization') !== -1) {
    // Regularization: 14 working days (treated as window from request)
    const due = addWorkingDays(base, 14);
    logSlaDebug_('sla_match_movement_regularization', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'movement_regularization_14_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('movement - promotion') !== -1) {
    // Promotion: 2 working days to complete
    const due = addWorkingDays(base, 2);
    logSlaDebug_('sla_match_movement_promotion', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'movement_promotion_2_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('movement - salary adjustment') !== -1) {
    // Salary Adjustment: 2 working days to complete
    const due = addWorkingDays(base, 2);
    logSlaDebug_('sla_match_movement_salary_adjustment', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'movement_salary_adjustment_2_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('movement - force transfer') !== -1) {
    // Force Transfer: 2 working days to complete
    const due = addWorkingDays(base, 2);
    logSlaDebug_('sla_match_movement_force_transfer', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'movement_force_transfer_2_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('movement - level up') !== -1) {
    // Level Up: Every Dec 1 and June 1
    const due = nextLevelUpDate(base);
    logSlaDebug_('sla_match_movement_level_up', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'movement_level_up_next_jun_1_or_dec_1',
      dueDate: due.toString()
    });
    return due;
  }

  // Training services
  if (svcKey.indexOf('training - employee orientation') !== -1) {
    // Employee Orientation: first day of employee (treated as base)
    const due = new Date(base.getTime());
    logSlaDebug_('sla_match_training_employee_orientation', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'training_employee_orientation_same_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('training - training materials') !== -1) {
    // Training Materials: 1 month from date of request
    const due = addCalDays(base, 30);
    logSlaDebug_('sla_match_training_materials', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'training_materials_30_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (
    svcKey.indexOf('training - promotional training') !== -1 ||
    svcKey.indexOf('training - level up training') !== -1 ||
    svcKey.indexOf('training - transfer training') !== -1 ||
    svcKey.indexOf('training - pip training') !== -1
  ) {
    // Training completion: 5 working days window
    const due = addWorkingDays(base, 5);
    logSlaDebug_('sla_match_training_generic_5days', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'training_generic_5_working_days',
      dueDate: due.toString()
    });
    return due;
  }

  // Benefits / Leaves services
  if (svcKey.indexOf('benefits - leave allocation') !== -1) {
    // Leave Allocation:
    // Annual: Every Jan 1 (next Jan 1)
    const year = base.getFullYear();
    let target = new Date(year, 0, 1);
    if (target < base) {
      target = new Date(year + 1, 0, 1);
    }
    const due = target;
    logSlaDebug_('sla_match_benefits_leave_allocation', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'benefits_leave_allocation_next_jan_1',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('benefits - leave reports') !== -1) {
    // Leave Reports: Monthly – 2nd working day of following month
    const due = secondWorkingDayOfNextMonth(base);
    logSlaDebug_('sla_match_benefits_leave_reports', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'benefits_leave_reports_2nd_working_day_next_month',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('benefits - hmo') !== -1) {
    // HMO: treat as 30 calendar days window (approximation)
    const due = addCalDays(base, 30);
    logSlaDebug_('sla_match_benefits_hmo', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'benefits_hmo_30_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  // Clinic services
  if (svcKey.indexOf('clinic - validation of sickness availments') !== -1) {
    // Validation: within the day of submission
    const due = new Date(base.getTime());
    logSlaDebug_('sla_match_clinic_validation', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'clinic_validation_same_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('clinic - illness report') !== -1) {
    // Illness Report: every end of the month
    const due = lastDayOfMonth(base);
    logSlaDebug_('sla_match_clinic_illness_report', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'clinic_illness_report_end_of_month',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('clinic - health & wellness program') !== -1) {
    // Health & Wellness: every 2nd week of Feb, Jun, Sep, Dec
    const due = nextHealthWellnessDate(base);
    logSlaDebug_('sla_match_clinic_health_wellness', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'clinic_health_wellness_next_cycle',
      dueDate: due ? due.toString() : null
    });
    return due;
  }

  // Performance Management
  if (svcKey.indexOf('performance management') !== -1) {
    // 9-box report: 17th working day of the following month
    const due = nthWorkingDayOfNextMonth(base, 17);
    logSlaDebug_('sla_match_performance_management', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'performance_9box_17th_working_day_next_month',
      dueDate: due.toString()
    });
    return due;
  }

  // Organizational / OD / Reports
  if (svcKey.indexOf('org - org development') !== -1) {
    // Org Development: 2nd week of November
    const y = base.getFullYear();
    let target = new Date(y, 10, 8); // Nov 8 ~ start of 2nd week
    if (target < base) {
      target = new Date(y + 1, 10, 8);
    }
    const due = target;
    logSlaDebug_('sla_match_org_development', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'org_development_2nd_week_november',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('org - organizational report') !== -1) {
    // Organizational Report – Monthly: 2nd working day of following month
    const due = secondWorkingDayOfNextMonth(base);
    logSlaDebug_('sla_match_org_organizational_report', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'org_organizational_report_2nd_working_day_next_month',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('org - compensation report') !== -1) {
    // Compensation Report – Quarterly: 2nd working day of succeeding quarter
    const due = secondWorkingDayOfNextQuarter(base);
    logSlaDebug_('sla_match_org_compensation_report', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'org_compensation_report_2nd_working_day_next_quarter',
      dueDate: due.toString()
    });
    return due;
  }

  // Payroll services
  if (svcKey.indexOf('payroll - payroll computation & crediting (local)') !== -1) {
    // Local payroll: every 2nd and 17th of the month
    const due = nextLocalPayrollDate(base);
    logSlaDebug_('sla_match_payroll_local', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'payroll_local_next_2nd_or_17th',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('payroll - payroll computation & crediting (expat)') !== -1) {
    // Expat payroll: every 10th of the month
    const due = nextDayOfMonth(base, 10);
    logSlaDebug_('sla_match_payroll_expat', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'payroll_expat_next_10th',
      dueDate: due.toString()
    });
    return due;
  }

  if (
    svcKey.indexOf('payroll - last pay') !== -1 ||
    svcKey.indexOf('payroll - separation pay') !== -1
  ) {
    // Last Pay & Separation Pay: 30 days from accomplished clearance
    const due = addCalDays(base, 30);
    logSlaDebug_('sla_match_payroll_last_or_separation_pay', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'payroll_last_or_separation_pay_30_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('payroll - tax annualization') !== -1) {
    // Annualization: on or before January 31
    const due = nextFixedMonthDay(base, 0, 31); // Jan 31
    logSlaDebug_('sla_match_payroll_tax_annualization', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'payroll_tax_annualization_next_jan_31',
      dueDate: due.toString()
    });
    return due;
  }

  // 13. Records Management — within 1 working day from receipt
  if (svcKey.indexOf('records management') !== -1) {
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_records_management', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'records_management_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  // Generic "within 1 working day" from date of request/information
  // Now covers both "Certification of Employee's Loan"
  // and "Certificate of Employee's Loan" (SSS / PAGIBIG variants).
  if (
    /certification of employee'?s loan|certificate of employee'?s loan|unemployment certification|certificate of contributions|csf|cfi|mdr/i
      .test(svcRaw)
  ) {
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_generic_1_working_day', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'generic_certification_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  // SSS services (contributions, remittances, notifications)
  if (svcKey.indexOf('sss - new hire reporting') !== -1) {
    // Reporting of newly hired: onboarded date (treated as base)
    const due = new Date(base.getTime());
    logSlaDebug_('sla_match_sss_new_hire_reporting', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'sss_new_hire_reporting_same_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (
    svcKey.indexOf('sss - contributions') !== -1 ||
    svcKey.indexOf('sss - loan remittances') !== -1
  ) {
    // Contributions & Loan Remittances: on or before end of the month
    const due = lastDayOfMonth(base);
    logSlaDebug_('sla_match_sss_contributions_or_loan_remittances', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'sss_contributions_or_loan_remittances_end_of_month',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('sss - sickness notification') !== -1) {
    // Sickness Notification: 5 calendar days after receipt
    const due = addCalDays(base, 5);
    logSlaDebug_('sla_match_sss_sickness_notification', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'sss_sickness_notification_5_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('sss - maternity notification') !== -1) {
    // Maternity Notification: 60 days from conception (approx 60 days from request)
    const due = addCalDays(base, 60);
    logSlaDebug_('sla_match_sss_maternity_notification', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'sss_maternity_notification_60_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('sss - disability notification') !== -1) {
    // Disability Notification: treat similar to sickness (5 calendar days)
    const due = addCalDays(base, 5);
    logSlaDebug_('sla_match_sss_disability_notification', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'sss_disability_notification_5_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('sss - annual processing of l-501') !== -1) {
    // Annual Processing of L-501: Every January 15
    const due = nextFixedMonthDay(base, 0, 15); // Jan 15
    logSlaDebug_('sla_match_sss_l501', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'sss_l501_next_jan_15',
      dueDate: due.toString()
    });
    return due;
  }

  // PHIC services
  if (svcKey.indexOf('phic - new hire reporting') !== -1) {
    // Reporting of newly hired: onboarded date (treated as base)
    const due = new Date(base.getTime());
    logSlaDebug_('sla_match_phic_new_hire_reporting', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'phic_new_hire_reporting_same_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('phic - contributions') !== -1) {
    // PHIC Contributions & Remittances: on or before the 15th day of the month
    const due = nextDayOfMonth(base, 15);
    logSlaDebug_('sla_match_phic_contributions', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'phic_contributions_next_15th',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('phic - csf, csi, mdr, certificate of contributions') !== -1) {
    // CSF, CSI, MDR; Certificate of Contributions: within 1 working day
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_phic_csf_csi_mdr', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'phic_csf_csi_mdr_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  // DOLE services
  if (svcKey.indexOf('dole - annual') !== -1) {
    // DOLE Annual: nearest of Jan 15, Jan 31, Mar 31, Jun 15
    const due = nextDoleAnnualDate(base);
    logSlaDebug_('sla_match_dole_annual', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'dole_annual_next_key_date',
      dueDate: due ? due.toString() : null
    });
    return due;
  }

  if (svcKey.indexOf('dole - monthly (wair)') !== -1) {
    // WAIR Monthly: on or before end of the month
    const due = lastDayOfMonth(base);
    logSlaDebug_('sla_match_dole_monthly_wair', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'dole_monthly_wair_end_of_month',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('dole - termination report') !== -1) {
    // DOLE Termination Report: 30 days from effectivity date (treated as base)
    const due = addCalDays(base, 30);
    logSlaDebug_('sla_match_dole_termination_report', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'dole_termination_report_30_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (svcKey.indexOf('dole - quarterly (aep report)') !== -1) {
    // AEP Quarterly Report: Jan 31, Apr 30, Jul 31, Oct 31
    const due = nextAepQuarterDate(base);
    logSlaDebug_('sla_match_dole_quarterly_aep', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'dole_quarterly_aep_next_cycle',
      dueDate: due ? due.toString() : null
    });
    return due;
  }

  // PAGIBIG services
  if (svcKey.indexOf('pagibig - certification of employee\'s loan') !== -1) {
    // Certification of Employee's Loan: within 1 working day
    const due = addWorkingDays(base, 1);
    logSlaDebug_('sla_match_pagibig_certification_loan', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'pagibig_certification_loan_1_working_day',
      dueDate: due.toString()
    });
    return due;
  }

  if (
    svcKey.indexOf('pagibig - contributions') !== -1 ||
    svcKey.indexOf('pagibig - loan remittances') !== -1
  ) {
    // Contributions & Loan Remittances: on or before the 15th day of the month
    const due = nextDayOfMonth(base, 15);
    logSlaDebug_('sla_match_pagibig_contributions_or_loan_remittances', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'pagibig_contributions_or_loan_remittances_next_15th',
      dueDate: due.toString()
    });
    return due;
  }

  // 5. Compliance with Government Regulations – some date-based examples (generic)
  if (svc.indexOf('termination report') !== -1) {
    // "30 days from the effectivity date" – assume effectivity = request date
    const due = addCalDays(base, 30);
    logSlaDebug_('sla_match_generic_termination_report', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'generic_termination_report_30_calendar_days',
      dueDate: due.toString()
    });
    return due;
  }

  if (/13th month pay report/i.test(svcRaw)) {
    // "every January 15"
    const due = nextFixedMonthDay(base, 0, 15); // January 15
    logSlaDebug_('sla_match_13th_month_pay_report', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: '13th_month_pay_report_jan_15',
      dueDate: due.toString()
    });
    return due;
  }

  if (/annual wages report/i.test(svcRaw)) {
    // "every June 15"
    const due = nextFixedMonthDay(base, 5, 15); // June 15
    logSlaDebug_('sla_match_annual_wages_report', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'annual_wages_report_jun_15',
      dueDate: due.toString()
    });
    return due;
  }

  // Monthly report: "Every 2nd working day of the following month"
  if (/monthly/i.test(svcRaw) && /report|organizational|compensation/i.test(svcRaw)) {
    const due = secondWorkingDayOfNextMonth(base);
    logSlaDebug_('sla_match_monthly_report_generic', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'monthly_report_2nd_working_day_next_month',
      dueDate: due.toString()
    });
    return due;
  }

  // Quarterly report: "Every 2nd working day of the succeeding quarter"
  if (/quarterly/i.test(svcRaw) && /report|organizational|compensation/i.test(svcRaw)) {
    const due = secondWorkingDayOfNextQuarter(base);
    logSlaDebug_('sla_match_quarterly_report_generic', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'quarterly_report_2nd_working_day_next_quarter',
      dueDate: due.toString()
    });
    return due;
  }

  // Annual Jan 15 style (organizational / compensation reports, etc.)
  if (/annual/i.test(svcRaw) && /report|organizational|compensation/i.test(svcRaw)) {
    const due = nextFixedMonthDay(base, 0, 15); // January 15
    logSlaDebug_('sla_match_annual_report_jan_15_generic', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'annual_report_jan_15_generic',
      dueDate: due.toString()
    });
    return due;
  }

  // Recruitment (if you want a conservative SLA; here 12 working days similar to Employee Relations)
  if (svcKey.indexOf('recruitment') !== -1) {
    const due = addWorkingDays(base, 12);
    logSlaDebug_('sla_match_recruitment', {
      service,
      processStep,
      requestDate: requestDate.toString(),
      rule: 'recruitment_12_working_days',
      dueDate: due.toString()
    });
    return due;
  }

// Ad Hoc – still no fixed SLA in table, keep as "no due date"
if (svcKey.indexOf('ad hoc') !== -1) {
  logSlaDebug_('sla_unmapped_ad_hoc', {
    service,
    processStep,
    requestDate: requestDate.toString(),
    reason: 'no explicit SLA rule defined for Ad Hoc'
  });
  return null;
}

// Approval – treat as "within 1 working day" from request
if (svcKey.indexOf('approval') !== -1) {
  const due = addWorkingDays(base, 1);
  logSlaDebug_('sla_match_approval_1_working_day', {
    service,
    processStep,
    requestDate: requestDate.toString(),
    rule: 'approval_1_working_day_from_request',
    dueDate: due.toString()
  });
  return due;
}

  // ❌ If no rule matches, log that we are returning null (no SLA match)
  logSlaDebug_('no_sla_rule_match', {
    reason: 'no SLA rule matched service/processStep',
    service,
    processStep,
    requestDate: requestDate.toString()
  });

  // If no rule matches, return null so we don't set a Due Date
  return null;
}

/* -------- Category options & Process Steps pass-throughs -------- */

function getCategoryOptions() {
  // If a specific backend implementation exists, use it; else fallback
  if (typeof getCategoryOptionsBackend === 'function') {
    return getCategoryOptionsBackend();
  }
  return getCategoryOptionsLegacy_();
}

function getProcessSteps(service) {
  if (typeof getProcessStepsBackend === 'function') {
    return getProcessStepsBackend(service);
  }
  return getProcessStepsLegacy_(service);
}

/* Legacy implementations kept as fallback (used only if modular versions are missing) */

function getCategoryOptionsLegacy_() {
  const sh = getSheetByAnyName_(getAdminSs(), [CONFIG.SHEETS.LIST, 'LIST', 'list']);
  if (!sh) return {};
  const vals = sh.getDataRange().getDisplayValues();
  if (!vals.length) return {};
  const header = vals[0] || [];
  const map = {};
  for (let c = 0; c < header.length; c++) {
    const cat = String(header[c] || '').trim();
    if (!cat) continue;
    const options = [];
    for (let r = 1; r < vals.length; r++) {
      const cell = String((vals[r] && vals[r][c]) || '').trim();
      if (cell) options.push(cell);
    }
    if (options.length) map[cat] = Array.from(new Set(options));
  }
  return map;
}

function getProcessStepsLegacy_(service) {
  if (!service) return [];
  const sh = getSheetByAnyName_(getAdminSs(), [CONFIG.SHEETS.LIST, 'LIST', 'list']);
  if (!sh) return [];
  const vals = sh.getDataRange().getDisplayValues();
  if (!vals.length) return [];
  const header = vals[0] || [];
  const want   = String(service).trim().toLowerCase();
  let col = -1;
  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] || '').trim().toLowerCase();
    if (h && h === want) {
      col = c;
      break;
    }
  }
  if (col === -1) return [];
  const steps = [];
  for (let r = 1; r < vals.length; r++) {
    const cell = String((vals[r] && vals[r][col]) || '').trim();
    if (cell) steps.push(cell);
  }
  return Array.from(new Set(steps));
}

/**
 * Pass-throughs so Request Form can also see active requests if needed.
 * (Uses the HRRequest backend implementations)
 */
function getActiveRequests() {
  return (typeof getActiveRequestsBackend === 'function')
    ? getActiveRequestsBackend()
    : [];
}

function getRequestInfo(requestId) {
  return (typeof getRequestInfoBackend === 'function')
    ? getRequestInfoBackend(requestId)
    : null;
}
