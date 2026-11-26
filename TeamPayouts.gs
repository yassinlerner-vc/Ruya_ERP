// ============================================================
//  TEAM PAYOUTS — deferred / partial payment of revenue shares
//
//  Every Revenue_Distribution row is an ALLOCATION: it says how
//  much of a project's collected EGP belongs to a person. For
//  non-partner team members the admin chooses, per row, whether
//  that allocation is PAID now or DEFERRED (pay later).
//
//  Source of truth for "how much has actually been paid" is the
//  Team_Payouts sheet (one row per payment, so partial payments,
//  dates and methods are all captured). Revenue_Distribution's
//  Status column (DIST.STATUS) is the stored, human-readable flag:
//
//    PAID      — fully paid (or blank = legacy row, read as paid)
//    DEFERRED  — allocated, not yet fully paid
//
//  A DEFERRED row flips to PAID the moment its payouts add up to
//  the full allocation, and flips back if a payout is undone.
//
//  Team balance = sum of (amount − paid) over a member's rows.
//  Computed live every time — never stored — so it can't drift.
//
//  PARTNERS are NOT part of this: their share flows straight into
//  the Ledger (getPartnerBalances) and they are paid through
//  Drawings. Partner rows are always stored as PAID with no
//  payout record, and are excluded from team balances.
//
//  All amounts are EGP (distribution amounts always are), and all
//  maths is done in whole piastres (cents) so nothing drifts.
// ============================================================

// ------------------------------------------------------------
//  SMALL HELPERS
// ------------------------------------------------------------
function normDistStatus(v) {
  return String(v || '').toUpperCase() === DIST_STATUS_DEFERRED
    ? DIST_STATUS_DEFERRED : DIST_STATUS_PAID;
}

function toCents_(n) {
  return Math.round(Number(n || 0) * 100);
}

// Sheet date cells come back as Date objects (or, occasionally, as
// 'yyyy-MM-dd' text) — normalise both to 'yyyy-MM-dd'.
function fmtSheetDate_(val, tz) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  const s = String(val);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch (e) { return ''; }
}

function getPayoutsSheet_() {
  const sh = getSheet('Team_Payouts');
  if (!sh) throw new Error('The Team_Payouts sheet is missing. Run setupTeamPayouts() once from the Apps Script editor.');
  return sh;
}

// Set of Team IDs flagged IsPartner.
function getPartnerIdSet_() {
  const set = {};
  const sh = getSheet('Team');
  if (!sh) return set;
  sh.getDataRange().getValues().slice(1).forEach(row => {
    if (row[TM.TEAM_ID] && row[TM.IS_PARTNER]) set[String(row[TM.TEAM_ID])] = true;
  });
  return set;
}

// { distributionId: [ {payoutId, amount, date, method, notes} … ] }, oldest first.
function getPayoutsByDist_() {
  const map = {};
  const sh = getSheet('Team_Payouts');
  if (!sh) return map;
  const tz = Session.getScriptTimeZone();
  sh.getDataRange().getValues().slice(1).forEach(row => {
    if (!row[TPAY.PAYOUT_ID]) return;
    const key = String(row[TPAY.DISTRIBUTION_ID]);
    (map[key] = map[key] || []).push({
      payoutId: String(row[TPAY.PAYOUT_ID]),
      amount:   Number(row[TPAY.AMOUNT] || 0),
      date:     fmtSheetDate_(row[TPAY.DATE], tz),
      method:   String(row[TPAY.METHOD] || ''),
      notes:    String(row[TPAY.NOTES]  || '')
    });
  });
  Object.keys(map).forEach(k => map[k].sort((a, b) => (a.date || '').localeCompare(b.date || '')));
  return map;
}

// Payment state of ONE distribution row.
//   returns { status, paid, remaining, partial }
// A row marked PAID counts as fully paid whatever the payouts say;
// a DEFERRED row's paid amount is the sum of its payouts (capped).
function distPayState_(amount, status, payouts) {
  const amountC = toCents_(amount);
  let paidC;
  if (normDistStatus(status) === DIST_STATUS_PAID) {
    paidC = amountC;
  } else {
    paidC = (payouts || []).reduce((s, p) => s + toCents_(p.amount), 0);
    if (paidC > amountC) paidC = amountC;
  }
  const remC = amountC - paidC;
  return {
    status:    remC <= 0 ? DIST_STATUS_PAID : DIST_STATUS_DEFERRED,
    paid:      paidC / 100,
    remaining: remC / 100,
    partial:   paidC > 0 && remC > 0
  };
}

// Used by Projects.gs getProjectDetail — payment info for each of one
// project's distribution rows, keyed by distributionId.
function getDistPayInfoForRows_(distRows) {
  const payouts  = getPayoutsByDist_();
  const partners = getPartnerIdSet_();
  const info = {};
  distRows.forEach(row => {
    const id  = String(row[DIST.DISTRIBUTION_ID]);
    const pid = String(row[DIST.PERSON_ID] || '');
    const isPartner = !!partners[pid];
    const st = distPayState_(row[DIST.AMOUNT], row[DIST.STATUS], payouts[id]);
    st.isPartner = isPartner;
    // payState: what the UI shows
    st.payState = isPartner ? 'partner'
                : st.status === DIST_STATUS_PAID ? 'paid'
                : st.partial ? 'partial' : 'deferred';
    info[id] = st;
  });
  return info;
}

function hasPayoutsForProject_(projectId) {
  const sh = getSheet('Team_Payouts');
  if (!sh) return false;
  return sh.getDataRange().getValues().slice(1)
    .some(r => r[TPAY.PAYOUT_ID] && String(r[TPAY.PROJECT_ID]) === String(projectId));
}

// Returns a function that hands out PAYOUT IDs in sequence (TP000001…)
// without re-reading the sheet each time — used when saveDistribution
// writes several payouts in one go.
function makePayoutIdGen_() {
  let maxNum = 0;
  const sh = getSheet('Team_Payouts');
  if (sh) {
    sh.getDataRange().getValues().slice(1).forEach(row => {
      const m = String(row[TPAY.PAYOUT_ID] || '').match(/^TP(\d+)$/);
      if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
    });
  }
  return function () { maxNum++; return 'TP' + String(maxNum).padStart(6, '0'); };
}

// ------------------------------------------------------------
//  BALANCES  — every non-partner member's allocated / paid / owed
//  returns { personId: { allocated, paid, outstanding } } (EGP)
// ------------------------------------------------------------
function computeTeamBalances_() {
  const out = {};
  const distSheet = getSheet('Revenue_Distribution');
  if (!distSheet) return out;
  const partners = getPartnerIdSet_();
  const payouts  = getPayoutsByDist_();

  distSheet.getDataRange().getValues().slice(1).forEach(row => {
    if (!row[DIST.DISTRIBUTION_ID]) return;
    const pid = String(row[DIST.PERSON_ID] || '');
    if (!pid || partners[pid]) return;
    const st = distPayState_(row[DIST.AMOUNT], row[DIST.STATUS], payouts[String(row[DIST.DISTRIBUTION_ID])]);
    const b = out[pid] = out[pid] || { allocatedC: 0, paidC: 0, outstandingC: 0 };
    b.allocatedC   += toCents_(row[DIST.AMOUNT]);
    b.paidC        += toCents_(st.paid);
    b.outstandingC += toCents_(st.remaining);
  });

  Object.keys(out).forEach(pid => {
    const b = out[pid];
    out[pid] = { allocated: b.allocatedC / 100, paid: b.paidC / 100, outstanding: b.outstandingC / 100 };
  });
  return out;
}

// Team list for the Team module — getTeam() plus a live Balance.
// Partners get balance = null (they're settled via the Ledger).
function getTeamWithBalances() {
  const team = getTeam();
  const bal  = computeTeamBalances_();
  return team.map(t => Object.assign({}, t, {
    balance: t.isPartner ? null : (bal[t.id] ? bal[t.id].outstanding : 0)
  }));
}

// ------------------------------------------------------------
//  TEAM MEMBER PROFILE — summary + every project they're in
// ------------------------------------------------------------
function getTeamMemberPayouts(personId) {
  const member = getTeamById(personId);
  if (!member) return null;
  if (member.isPartner) {
    return { isPartner: true, summary: null, rows: [] };
  }

  const distSheet = getSheet('Revenue_Distribution');
  const tz       = Session.getScriptTimeZone();
  const payouts  = getPayoutsByDist_();
  const rows = [];
  let allocC = 0, paidC = 0, outC = 0;

  if (distSheet) {
    distSheet.getDataRange().getValues().slice(1).forEach(row => {
      if (!row[DIST.DISTRIBUTION_ID]) return;
      if (String(row[DIST.PERSON_ID]) !== String(personId)) return;

      const id   = String(row[DIST.DISTRIBUTION_ID]);
      const list = payouts[id] || [];
      const st   = distPayState_(row[DIST.AMOUNT], row[DIST.STATUS], list);
      const dateAllocated = fmtSheetDate_(row[DIST.CREATED_AT], tz);

      // Date paid: latest payout date; legacy PAID rows with no payout
      // record fall back to the allocation date.
      let datePaid = '';
      if (list.length) datePaid = list[list.length - 1].date;
      else if (st.status === DIST_STATUS_PAID) datePaid = dateAllocated;

      allocC += toCents_(row[DIST.AMOUNT]);
      paidC  += toCents_(st.paid);
      outC   += toCents_(st.remaining);

      rows.push({
        distributionId: id,
        projectId:      String(row[DIST.PROJECT_ID]   || ''),
        projectName:    String(row[DIST.PROJECT_NAME] || ''),
        accountName:    String(row[DIST.ACCOUNT_NAME] || ''),
        dateAllocated:  dateAllocated,
        mode:           normDistMode(row[DIST.MODE]),
        percent:        Number(row[DIST.PERCENT] || 0),
        amount:         Number(row[DIST.AMOUNT]  || 0),
        paid:           st.paid,
        remaining:      st.remaining,
        status:         st.status,
        partial:        st.partial,
        datePaid:       datePaid,
        payouts:        list
      });
    });
  }

  // Newest allocation first; unpaid rows float to the top within a date.
  rows.sort((a, b) =>
    (b.dateAllocated || '').localeCompare(a.dateAllocated || '') ||
    (b.remaining > 0 ? 1 : 0) - (a.remaining > 0 ? 1 : 0));

  return {
    isPartner: false,
    summary: { totalAllocated: allocC / 100, totalPaid: paidC / 100, outstanding: outC / 100 },
    rows: rows
  };
}

// ------------------------------------------------------------
//  PAY — record a (full or partial) payment against one row
//  data: { amount, method, date ('yyyy-MM-dd', optional), notes }
// ------------------------------------------------------------
function recordTeamPayout(distributionId, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const distSheet = getSheet('Revenue_Distribution');
    const paySheet  = getPayoutsSheet_();
    if (!distSheet) return { success: false, error: 'Revenue_Distribution sheet missing.' };

    const distData = distSheet.getDataRange().getValues();
    let rowIndex = -1, row = null;
    for (let i = 1; i < distData.length; i++) {
      if (String(distData[i][DIST.DISTRIBUTION_ID]) === String(distributionId)) {
        rowIndex = i + 1; row = distData[i]; break;
      }
    }
    if (rowIndex === -1) return { success: false, error: 'Allocation not found.' };

    const personId = String(row[DIST.PERSON_ID] || '');
    if (getPartnerIdSet_()[personId]) {
      return { success: false, error: 'Partners are settled through the Ledger / Drawings, not team payouts.' };
    }

    const method = String((data && data.method) || '');
    if (PAYMENT_METHODS.indexOf(method) === -1) {
      return { success: false, error: 'Choose a payment method.' };
    }

    const tz    = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const dateStr = (data && data.date) ? String(data.date) : today;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { success: false, error: 'Invalid payment date.' };
    if (dateStr > today) return { success: false, error: 'The payment date cannot be in the future.' };

    const amount  = Number(String((data && data.amount) || '').replace(/,/g, ''));
    const amountC = toCents_(amount);
    if (!isFinite(amount) || amountC <= 0) return { success: false, error: 'Enter a valid amount to pay.' };

    const distId  = String(row[DIST.DISTRIBUTION_ID]);
    const existing = (getPayoutsByDist_()[distId]) || [];
    const before  = distPayState_(row[DIST.AMOUNT], row[DIST.STATUS], existing);
    const remC    = toCents_(before.remaining);
    if (remC <= 0) return { success: false, error: 'This allocation is already fully paid.' };
    if (amountC > remC) {
      return { success: false, error: 'The amount exceeds what is still owed (EGP ' + (remC / 100).toFixed(2) + ').' };
    }

    const user     = Session.getActiveUser().getEmail();
    const now      = new Date();
    const payoutId = makePayoutIdGen_()();
    const projectName = String(row[DIST.PROJECT_NAME] || '');
    const personName  = String(row[DIST.PERSON_NAME]  || '');

    paySheet.appendRow([
      payoutId, distId, String(row[DIST.PROJECT_ID] || ''), projectName,
      personId, personName, amountC / 100, dateStr, method,
      (data && data.notes) ? String(data.notes) : '', user, now
    ]);

    const newRemC = remC - amountC;
    const newStatus = newRemC <= 0 ? DIST_STATUS_PAID : DIST_STATUS_DEFERRED;
    distSheet.getRange(rowIndex, DIST.STATUS + 1).setValue(newStatus);

    writeLog('Team_Log', 'Team Payout', payoutId,
      personId + ' — ' + personName, '', String(row[DIST.ACCOUNT_NAME] || ''),
      'Payout', '', amountC / 100,
      'Paid EGP ' + (amountC / 100).toFixed(2) + ' via ' + method + ' for ' + projectName +
      ' (' + String(row[DIST.PROJECT_ID] || '') + '); EGP ' + (newRemC / 100).toFixed(2) + ' remaining. By ' + user);

    return { success: true, payoutId: payoutId, remaining: newRemC / 100, status: newStatus };
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
//  UNDO — delete a payout entered by mistake, re-opening the row
// ------------------------------------------------------------
function voidTeamPayout(payoutId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const paySheet  = getPayoutsSheet_();
    const distSheet = getSheet('Revenue_Distribution');

    const payData = paySheet.getDataRange().getValues();
    let payRow = -1, p = null;
    for (let i = 1; i < payData.length; i++) {
      if (String(payData[i][TPAY.PAYOUT_ID]) === String(payoutId)) { payRow = i + 1; p = payData[i]; break; }
    }
    if (payRow === -1) return { success: false, error: 'Payment not found.' };

    const distId = String(p[TPAY.DISTRIBUTION_ID]);
    paySheet.deleteRow(payRow);

    // Re-evaluate the allocation this payout belonged to.
    if (distSheet) {
      const distData = distSheet.getDataRange().getValues();
      for (let i = 1; i < distData.length; i++) {
        if (String(distData[i][DIST.DISTRIBUTION_ID]) !== distId) continue;
        const left = (getPayoutsByDist_()[distId]) || [];
        const leftC = left.reduce((s, x) => s + toCents_(x.amount), 0);
        const status = leftC >= toCents_(distData[i][DIST.AMOUNT]) ? DIST_STATUS_PAID : DIST_STATUS_DEFERRED;
        distSheet.getRange(i + 1, DIST.STATUS + 1).setValue(status);
        break;
      }
    }

    const user = Session.getActiveUser().getEmail();
    writeLog('Team_Log', 'Team Payout', payoutId,
      String(p[TPAY.PERSON_ID]) + ' — ' + String(p[TPAY.PERSON_NAME]), '', '',
      'Payout', Number(p[TPAY.AMOUNT] || 0), 'Undone',
      'Payment of EGP ' + Number(p[TPAY.AMOUNT] || 0).toFixed(2) + ' (' + String(p[TPAY.METHOD] || '') +
      ') for ' + String(p[TPAY.PROJECT_NAME] || '') + ' undone by ' + user);

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  ONE-TIME SETUP — run once from the Apps Script editor
//  (select setupTeamPayouts in the function dropdown ▸ Run).
//  Safe to run again: it never overwrites existing data.
//    1. Creates the Team_Payouts sheet with its headers.
//    2. Adds the "Status" header to Revenue_Distribution (col O).
//    3. Backfills Status = PAID on every existing distribution row
//       that has no status yet (they were all paid-by-definition
//       before this feature).
// ============================================================
function setupTeamPayouts() {
  const ss = SpreadsheetApp.getActive();

  let pay = ss.getSheetByName('Team_Payouts');
  if (!pay) {
    pay = ss.insertSheet('Team_Payouts');
    const headers = ['Payout ID', 'Distribution ID', 'Project ID', 'Project Name',
      'Person ID', 'Person Name', 'Amount (EGP)', 'Date', 'Method', 'Notes',
      'Created By', 'Created At'];
    pay.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#c9daf8');
    pay.setFrozenRows(1);
    pay.setTabColor('#c9daf8');
  }

  const dist = ss.getSheetByName('Revenue_Distribution');
  if (!dist) throw new Error('Revenue_Distribution sheet not found.');

  const hdrCell = dist.getRange(1, DIST.STATUS + 1);
  if (!String(hdrCell.getValue() || '').trim()) {
    hdrCell.setValue('Status');
    dist.getRange(1, DIST.MODE + 1).copyTo(hdrCell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const lastRow = dist.getLastRow();
  let filled = 0;
  if (lastRow >= 2) {
    const ids  = dist.getRange(2, DIST.DISTRIBUTION_ID + 1, lastRow - 1, 1).getValues();
    const stat = dist.getRange(2, DIST.STATUS + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < stat.length; i++) {
      if (ids[i][0] && !String(stat[i][0] || '').trim()) { stat[i][0] = DIST_STATUS_PAID; filled++; }
    }
    dist.getRange(2, DIST.STATUS + 1, stat.length, 1).setValues(stat);
  }

  Logger.log('Team_Payouts ready. Status header set; backfilled ' + filled + ' existing row(s) as PAID.');
  try { refreshHome(); } catch (e) { /* home index is cosmetic */ }
}
