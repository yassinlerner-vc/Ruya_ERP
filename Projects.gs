// ============================================================
//  PROJECTS — GENERATE PAYMENT ID
// ============================================================
function generatePaymentId() {
  const sheet = getSheet('Payments');
  const data  = sheet.getDataRange().getValues();
  let maxNum  = 0;
  data.slice(1).forEach(row => {
    const match = String(row[PAY.PAYMENT_ID] || '').match(/^PAY(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  });
  return 'PAY' + String(maxNum + 1).padStart(6, '0');
}

// ============================================================
//  REVENUE DISTRIBUTION — MODE COLUMN
//  Each distribution row is either a PERCENT share of the EGP
//  collected, or a fixed AMOUNT in EGP. The mode is stored in the
//  "Mode" column of Revenue_Distribution (DIST.MODE, 0-based
//  index 13 = column N). Rows written before this column existed
//  have it blank, which is read as 'percent', which is exactly how
//  they were created.
// ============================================================

function normDistMode(v) {
  return String(v || '').toLowerCase() === 'amount' ? 'amount' : 'percent';
}

// ============================================================
//  FREE-PROJECT HELPER
//  A project doesn't snapshot IsFree onto its own row (today it
//  only snapshots Total/Currency at confirmation) — this looks it
//  up through the existing Quotation link every time. Single
//  source of truth for "is this project free?", used everywhere
//  that matters: listing, detail, payments, distribution, and
//  delivery/fulfillment.
// ============================================================
function isProjectFree(quotationId) {
  if (!quotationId) return false;
  const qSheet = getSheet('Quotations');
  if (!qSheet) return false;
  const qData = qSheet.getDataRange().getValues();
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      return !!qData[i][Q.IS_FREE];
    }
  }
  return false;
}

// ============================================================
//  PROJECTS — LIST
//  Payment status is computed live from the Payments sheet,
//  not stored, so it can never drift out of sync. Overall
//  project STATUS ('Active' / 'Fulfilled'), by contrast, IS
//  stored — it's a real lifecycle milestone set once by
//  fulfillProjectAndQuotation(), not something to recompute.
// ============================================================
function getProjects() {
  try {
    const pSheet    = getSheet('Projects');
    const paySheet  = getSheet('Payments');
    const distSheet = getSheet('Revenue_Distribution');
    if (!pSheet) return [];
    const data = pSheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const tz = Session.getScriptTimeZone();

    function fmtDate(val) {
      if (!val) return '';
      try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch(e) { return ''; }
    }

    // Build paid-amount map: projectId -> sum of payments
    const paidMap = {};
    if (paySheet) {
      paySheet.getDataRange().getValues().slice(1).forEach(row => {
        const pid = String(row[PAY.PROJECT_ID]);
        if (!pid) return;
        paidMap[pid] = (paidMap[pid] || 0) + (Number(row[PAY.AMOUNT]) || 0);
      });
    }

    // Which projects already have a distribution recorded
    const distributedIds = new Set();
    if (distSheet) {
      distSheet.getDataRange().getValues().slice(1).forEach(row => {
        const pid = String(row[DIST.PROJECT_ID]);
        if (pid) distributedIds.add(pid);
      });
    }

    const projects = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[P.PROJECT_ID]) continue;
      const pid         = String(row[P.PROJECT_ID]);
      const quotationId = String(row[P.QUOTATION_ID] || '');
      const totalAmount = Number(row[P.TOTAL_AMOUNT] || 0);
      const amountPaid  = paidMap[pid] || 0;
      const remaining   = Math.max(0, totalAmount - amountPaid);
      const isFree      = isProjectFree(quotationId);
      const status      = String(row[P.STATUS] || 'Active');

      // Free projects never collect payments, so they're excluded
      // from the Unpaid/Partially Paid/Fully Paid buckets entirely —
      // a distinct 'Free' paymentStatus keeps them out of those
      // filters (see the "Needs Review" filter tab in ProjectsForm).
      let paymentStatus = isFree ? 'Free' : 'Unpaid';
      if (!isFree) {
        if (totalAmount > 0 && amountPaid >= totalAmount) paymentStatus = 'Fully Paid';
        else if (amountPaid > 0) paymentStatus = 'Partially Paid';
      }

      try {
        projects.push({
          projectId:        pid,
          quotationId:      quotationId,
          accountId:        String(row[P.ACCOUNT_ID]   || ''),
          accountName:      String(row[P.ACCOUNT_NAME] || ''),
          projectName:      String(row[P.PROJECT_NAME] || ''),
          projectDesc:      String(row[P.PROJECT_DESC] || ''),
          deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
          dueDate:          fmtDate(row[P.DUE_DATE]),
          totalAmount:      totalAmount,
          currency:         String(row[P.CURRENCY] || 'EGP'),
          amountPaid:       amountPaid,
          remainingToPay:   remaining,
          paymentStatus:    paymentStatus,
          isDistributed:    distributedIds.has(pid),
          isFree:           isFree,
          deliveredDate:    fmtDate(row[P.DELIVERED_DATE]),
          status:           status,
          createdAt:        fmtDate(row[P.CREATED_AT])
        });
      } catch(e) {
        Logger.log('getProjects: skipping row ' + i + ': ' + e);
      }
    }

    const order = { 'Unpaid': 0, 'Partially Paid': 1, 'Fully Paid': 2 };
    projects.sort((a, b) => {
      const ao = order[a.paymentStatus] !== undefined ? order[a.paymentStatus] : 1;
      const bo = order[b.paymentStatus] !== undefined ? order[b.paymentStatus] : 1;
      if (ao !== bo) return ao - bo;
      return b.createdAt.localeCompare(a.createdAt);
    });

    return projects;
  } catch(e) {
    Logger.log('getProjects fatal error: ' + e);
    return [];
  }
}

// ============================================================
//  PROJECTS — GET DETAIL
//  Single call powering the whole detail view: project meta,
//  items (reference only), payments, distribution, team list.
// ============================================================
function getProjectDetail(projectId) {
  const pSheet    = getSheet('Projects');
  const piSheet   = getSheet('Project_Items');
  const paySheet  = getSheet('Payments');
  const distSheet = getSheet('Revenue_Distribution');
  const tSheet    = getSheet('Team');
  if (!pSheet) return null;

  const tz = Session.getScriptTimeZone();
  function fmtDate(val) {
    if (!val) return '';
    try { return Utilities.formatDate(new Date(val), tz, 'yyyy-MM-dd'); } catch(e) { return ''; }
  }

  // ── Project row ───────────────────────────────────────────
  const pData = pSheet.getDataRange().getValues();
  let project = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      const row = pData[i];
      project = {
        projectId:        String(row[P.PROJECT_ID]),
        quotationId:      String(row[P.QUOTATION_ID] || ''),
        accountId:        String(row[P.ACCOUNT_ID]   || ''),
        accountName:      String(row[P.ACCOUNT_NAME] || ''),
        projectName:      String(row[P.PROJECT_NAME] || ''),
        projectDesc:      String(row[P.PROJECT_DESC] || ''),
        deliveryDeadline: fmtDate(row[P.DELIVERY_DEADLINE]),
        dueDate:          fmtDate(row[P.DUE_DATE]),
        totalAmount:      Number(row[P.TOTAL_AMOUNT] || 0),
        currency:         String(row[P.CURRENCY] || 'EGP'),
        status:           String(row[P.STATUS] || 'Active'),
        deliveredDate:    fmtDate(row[P.DELIVERED_DATE])
      };
      break;
    }
  }
  if (!project) return null;

  project.isFree = isProjectFree(project.quotationId);

  // Quotation's Date Issued — the earliest allowed delivery date,
  // enforced both client- and server-side in confirmDelivery().
  let quotationDateIssued = '';
  const qSheetForDate = getSheet('Quotations');
  if (qSheetForDate) {
    const qDataForDate = qSheetForDate.getDataRange().getValues();
    for (let i = 1; i < qDataForDate.length; i++) {
      if (String(qDataForDate[i][Q.QUOTATION_ID]) === String(project.quotationId)) {
        quotationDateIssued = qDataForDate[i][Q.DATE_ISSUED]
          ? Utilities.formatDate(new Date(qDataForDate[i][Q.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '';
        break;
      }
    }
  }
  project.quotationDateIssued = quotationDateIssued;

  // ── Items (reference only — no assignment) ─────────────────
  const items = piSheet ? piSheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PI.PROJECT_ID]) === String(projectId))
    .map(row => ({
      itemId:       String(row[PI.ITEM_ID]),
      itemName:     String(row[PI.ITEM_NAME]),
      displayValue: String(row[PI.DISPLAY_VALUE] || ''),
      quantity:     Number(row[PI.QUANTITY] || 0),
      description:  String(row[PI.DESCRIPTION] || ''),
      notes:        String(row[PI.NOTES] || '')
    })) : [];

  // ── Payments ─────────────────────────────────────────────
  const payments = paySheet ? paySheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PAY.PROJECT_ID]) === String(projectId))
    .map(row => ({
      paymentId:     String(row[PAY.PAYMENT_ID]),
      amount:        Number(row[PAY.AMOUNT] || 0),
      date:          row[PAY.DATE] ? fmtDate(row[PAY.DATE]) : '',
      method:        String(row[PAY.METHOD] || ''),
      notes:         String(row[PAY.NOTES] || ''),
      recordedBy:    String(row[PAY.RECORDED_BY] || ''),
      recordedAt:    row[PAY.RECORDED_AT] ? fmtDate(row[PAY.RECORDED_AT]) : '',
      egpEquivalent: Number(row[PAY.EGP_EQUIVALENT] || 0),
      exchangeRate:  Number(row[PAY.EXCHANGE_RATE]  || 0)
    }))
    .sort((a, b) => (a.date || '').localeCompare(b.date || '')) : [];

  const amountPaid    = payments.reduce((sum, p) => sum + p.amount, 0);
  const remainingToPay = Math.max(0, project.totalAmount - amountPaid);
  const isFullyPaid    = project.totalAmount > 0 && amountPaid >= project.totalAmount;

  // Total EGP actually collected — same figure saveDistribution() uses to
  // compute amounts. Sent to the frontend so the distribution editor/table
  // can preview and label against the real EGP pool instead of the
  // project's native-currency total (which is what was being shown before,
  // even though the stored distribution amounts are always EGP).
  const totalEgpCollected = payments.reduce((sum, p) => sum + (p.egpEquivalent || 0), 0);

  // ── Distribution ─────────────────────────────────────────
  // Each row carries its MODE ('percent' | 'amount'). Rows saved
  // before the Mode column existed come back as 'percent'.
  const distribution = distSheet ? distSheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[DIST.PROJECT_ID]) === String(projectId))
    .map(row => ({
      distributionId: String(row[DIST.DISTRIBUTION_ID]),
      personId:       String(row[DIST.PERSON_ID]   || ''),
      personName:     String(row[DIST.PERSON_NAME] || ''),
      mode:           normDistMode(row[DIST.MODE]),
      percent:        Number(row[DIST.PERCENT] || 0),
      amount:         Number(row[DIST.AMOUNT]  || 0),
      notes:          String(row[DIST.NOTES]   || '')
    })) : [];

  // ── Team members (for distribution dropdown) ───────────────
  const teamMembers = tSheet ? tSheet.getDataRange().getValues().slice(1)
    .filter(row => row[TM.TEAM_ID])
    .map(row => ({
      id:   String(row[TM.TEAM_ID]),
      name: String(row[TM.NAME]),
      role: String(row[TM.ROLE])
    })) : [];

  return { project, items, payments, distribution, teamMembers, amountPaid, remainingToPay, isFullyPaid, totalEgpCollected };
}

// ============================================================
//  PAYMENTS — RECORD
//  Caps total payments at the project's quotation total.
//
//  EGP equivalent: required whenever the project's currency
//  isn't EGP (mirrors the same rule used on Expenses). For
//  EGP-currency projects the EGP equivalent is just the amount
//  itself and the rate is 1. This is what makes it possible to
//  later distribute commission off the ACTUAL EGP collected
//  rather than re-converting the quotation total — each
//  installment can land at a different rate and that difference
//  is captured here, payment by payment.
// ============================================================
function recordPayment(projectId, data) {
  const pSheet   = getSheet('Projects');
  const paySheet = getSheet('Payments');
  if (!pSheet || !paySheet) return { success: false, error: 'Required sheets missing.' };

  const pData = pSheet.getDataRange().getValues();
  let project = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) { project = pData[i]; break; }
  }
  if (!project) return { success: false, error: 'Project not found.' };

  const totalAmount     = Number(project[P.TOTAL_AMOUNT] || 0);
  const quotationId     = String(project[P.QUOTATION_ID] || '');
  const accountName     = String(project[P.ACCOUNT_NAME] || '');
  const projectName     = String(project[P.PROJECT_NAME] || '');
  const projectCurrency = String(project[P.CURRENCY] || 'EGP');

  if (isProjectFree(quotationId)) {
    throw new Error('This project is free — payments cannot be recorded against it.');
  }

  const existingPaid = paySheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PAY.PROJECT_ID]) === String(projectId))
    .reduce((sum, row) => sum + (Number(row[PAY.AMOUNT]) || 0), 0);

  const amount = Number(data.amount) || 0;
  if (amount <= 0) return { success: false, error: 'Enter a valid payment amount.' };
  if (existingPaid + amount > totalAmount + 0.01) {
    return { success: false, error: 'Payment exceeds the remaining balance of ' +
      (totalAmount - existingPaid).toFixed(2) + '.' };
  }

  const isForeign = projectCurrency !== 'EGP';
  let egpEquivalent = isForeign ? Number(data.egpEquivalent) || 0 : amount;
  if (isForeign && egpEquivalent <= 0) {
    return { success: false, error: 'Enter the EGP equivalent received for this payment.' };
  }
  const exchangeRate = isForeign ? (egpEquivalent / amount) : 1;

  const paymentId = generatePaymentId();
  const user = Session.getActiveUser().getEmail();
  const now  = new Date();
  const dateVal = data.date || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  paySheet.appendRow([
    paymentId, projectId, quotationId, accountName, projectName,
    amount, dateVal, data.method || '', data.notes || '', user, now,
    egpEquivalent, exchangeRate
  ]);

  writeLog('Payments_Log', 'Payments', paymentId,
    paymentId + ' — ' + projectName, '', accountName,
    'Payment', '', amount, 'Payment of ' + amount + ' recorded by ' + user +
    (isForeign ? ' (EGP equiv. ' + egpEquivalent.toFixed(2) + ')' : ''));

  return { success: true, paymentId };
}

// ============================================================
//  PAYMENTS — DELETE
//  Blocked once revenue has been distributed for the project,
//  to keep the books from going inconsistent.
// ============================================================
function deletePayment(paymentId) {
  const paySheet  = getSheet('Payments');
  const distSheet = getSheet('Revenue_Distribution');
  if (!paySheet) return { success: false, error: 'Payments sheet missing.' };

  const data = paySheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][PAY.PAYMENT_ID]) === String(paymentId)) {
      const projectId   = String(data[i][PAY.PROJECT_ID]);
      const projectName = String(data[i][PAY.PROJECT_NAME]);
      const quotationId = String(data[i][PAY.QUOTATION_ID] || '');

      if (isProjectFree(quotationId)) {
        throw new Error('This project is free — payments cannot be modified.');
      }

      if (distSheet) {
        const hasDist = distSheet.getDataRange().getValues().slice(1)
          .some(r => String(r[DIST.PROJECT_ID]) === projectId);
        if (hasDist) {
          return { success: false, error: 'Cannot delete a payment after revenue has been distributed for this project.' };
        }
      }

      paySheet.deleteRow(i + 1);
      writeLog('Payments_Log', 'Payments', paymentId,
        paymentId + ' — ' + projectName, '', '',
        'Status', 'Recorded', 'Deleted', 'Payment deleted');
      return { success: true };
    }
  }
  return { success: false, error: 'Payment not found.' };
}

// ============================================================
//  REVENUE DISTRIBUTION — SAVE
//  Only allowed once the project is fully paid. Replaces any
//  existing distribution rows for this project.
//
//  MIXED MODE: each incoming row has a `mode`:
//    • 'percent' — row.percent is a share (%) of the EGP pool
//    • 'amount'  — row.amount is a fixed EGP amount
//  Both can be mixed freely in one distribution. The pool is
//  always the ACTUAL EGP collected (see below), and the rows
//  together must allocate all of it: percent-rows' EGP value
//  plus fixed amounts must equal the pool (within a few piastres
//  of rounding, which is absorbed into the largest percent row —
//  or the largest row if there are no percent rows — so the
//  stored amounts always sum to the pool exactly).
//
//  Every stored row gets BOTH a Percent and an Amount: percent
//  rows keep the percent that was entered; fixed-amount rows get
//  a derived percent (amount ÷ pool). The Mode column records
//  which one was the input, so editing later reopens the row the
//  way it was entered. Anything that reads the stored Amount or
//  Percent columns keeps working unchanged.
//
//  Amount is computed server-side from the ACTUAL EGP collected
//  (summed from each payment's EGP_EQUIVALENT), NOT from the
//  quotation's native-currency total. This is what correctly
//  absorbs exchange-rate differences between installments —
//  e.g. a $500 quotation paid in two USD installments at
//  different rates might net EGP 10,000 total; that EGP 10,000
//  (not a fresh $500→EGP conversion) is what gets distributed.
//  Distribution is always recorded in EGP, regardless of the
//  project's native currency.
//
//  Once a distribution save succeeds, the project is by
//  definition both fully paid (checked below) and now
//  distributed — the two conditions "Fulfilled" requires beyond
//  Confirmed (which is implicit; only confirmed quotations have
//  a Project row at all). So this is the single trigger point
//  for fulfillProjectAndQuotation().
// ============================================================
function saveDistribution(projectId, rows) {
  const pSheet    = getSheet('Projects');
  const paySheet  = getSheet('Payments');
  const distSheet = getSheet('Revenue_Distribution');
  if (!pSheet || !distSheet) return { success: false, error: 'Required sheets missing.' };

  const pData = pSheet.getDataRange().getValues();
  let project = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) { project = pData[i]; break; }
  }
  if (!project) return { success: false, error: 'Project not found.' };

  const totalAmount = Number(project[P.TOTAL_AMOUNT] || 0);
  const quotationId = String(project[P.QUOTATION_ID] || '');
  const accountName = String(project[P.ACCOUNT_NAME] || '');
  const projectName = String(project[P.PROJECT_NAME] || '');

  if (isProjectFree(quotationId)) {
    throw new Error('This project is free — revenue distribution does not apply.');
  }

  const payments = paySheet ? paySheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[PAY.PROJECT_ID]) === String(projectId)) : [];

  const amountPaid = payments.reduce((sum, row) => sum + (Number(row[PAY.AMOUNT]) || 0), 0);

  if (totalAmount <= 0 || amountPaid < totalAmount) {
    return { success: false, error: 'Project must be fully paid before distributing revenue.' };
  }

  // Actual EGP collected — summed from each payment's EGP equivalent,
  // so exchange-rate differences between installments are absorbed
  // correctly rather than re-converting the quotation total once.
  const totalEgpCollected = payments.reduce((sum, row) => sum + (Number(row[PAY.EGP_EQUIVALENT]) || 0), 0);

  if (totalEgpCollected <= 0) {
    return { success: false, error: 'No EGP-equivalent payment amounts found for this project — cannot distribute.' };
  }

  if (!rows || !rows.length) {
    return { success: false, error: 'Add at least one person to distribute to.' };
  }

  // ── Work in whole piastres (cents) so nothing drifts ───────
  const poolCents = Math.round(totalEgpCollected * 100);
  const computed  = [];

  for (let i = 0; i < rows.length; i++) {
    const r    = rows[i] || {};
    const mode = normDistMode(r.mode);
    const who  = r.personName || 'a row';
    let cents, percentInput = null;

    if (mode === 'amount') {
      const amt = Number(r.amount);
      if (!isFinite(amt) || amt <= 0) {
        return { success: false, error: 'Enter a valid fixed amount for ' + who + '.' };
      }
      cents = Math.round(amt * 100);
    } else {
      const pct = Number(r.percent);
      if (!isFinite(pct) || pct <= 0) {
        return { success: false, error: 'Enter a valid percentage for ' + who + '.' };
      }
      percentInput = pct;
      cents = Math.round(poolCents * pct / 100);
    }
    computed.push({ r: r, mode: mode, cents: cents, percentInput: percentInput });
  }

  const sumCents = computed.reduce((s, c) => s + c.cents, 0);
  const diff     = poolCents - sumCents;                 // + = unallocated, − = over-allocated
  const tol      = Math.max(5, rows.length);             // rounding tolerance in piastres

  if (Math.abs(diff) > tol) {
    return { success: false, error: diff > 0
      ? 'EGP ' + (diff / 100).toFixed(2) + ' is still unallocated. The distribution must add up to the full EGP ' +
        totalEgpCollected.toFixed(2) + ' collected.'
      : 'Distribution is over by EGP ' + (-diff / 100).toFixed(2) + ' — it cannot exceed the EGP ' +
        totalEgpCollected.toFixed(2) + ' collected.' };
  }

  // Absorb sub-piastre rounding into the largest percent row (else largest row)
  // so stored amounts reconcile to the pool exactly.
  if (diff !== 0) {
    let target = -1;
    computed.forEach((c, idx) => {
      if (c.mode === 'percent' && (target === -1 || c.cents > computed[target].cents)) target = idx;
    });
    if (target === -1) {
      computed.forEach((c, idx) => { if (target === -1 || c.cents > computed[target].cents) target = idx; });
    }
    computed[target].cents += diff;
    if (computed[target].cents <= 0) {
      return { success: false, error: 'Distribution could not be reconciled — check the amounts.' };
    }
  }

  // Clear any existing distribution rows for this project
  const distData = distSheet.getDataRange().getValues();
  for (let i = distData.length - 1; i >= 1; i--) {
    if (String(distData[i][DIST.PROJECT_ID]) === String(projectId)) distSheet.deleteRow(i + 1);
  }

  const user = Session.getActiveUser().getEmail();
  const now  = new Date();
  const logRows = [];

  computed.forEach(c => {
    const amount  = c.cents / 100;
    const percent = c.mode === 'percent'
      ? c.percentInput
      : Math.round(c.cents / poolCents * 100 * 10000) / 10000;   // derived, 4 dp

    // Columns 0–12 are the original layout; Mode is column 13 (DIST_MODE_COL).
    distSheet.appendRow([
      Utilities.getUuid(), projectId, quotationId, accountName, projectName,
      c.r.personId || '', c.r.personName || '', percent, amount,
      'EGP', c.r.notes || '', user, now,
      c.mode
    ]);
    logRows.push({ name: c.r.personName, mode: c.mode, percent: percent, amount: amount });
  });

  writeLog('Revenue_Distribution_Log', 'Revenue_Distribution', projectId,
    projectId + ' — ' + projectName, '', accountName,
    'Distribution', '',
    JSON.stringify(logRows),
    'Revenue distributed by ' + user + ' — based on EGP ' + totalEgpCollected.toFixed(2) + ' actually collected');

  // Fully paid (checked above) and, as of this line, distributed —
  // two of the three gates. Delivery may or may not be confirmed
  // yet, so defer to the shared helper rather than fulfilling
  // unconditionally here.
  maybeFulfillProject(projectId, quotationId);

  return { success: true };
}

// ============================================================
//  FULFILLMENT GATE-CHECK — a normal (non-free) project becomes
//  Fulfilled once ALL THREE are true: fully paid, distributed, AND
//  delivery confirmed. These can complete in any order, so both
//  saveDistribution() and confirmDelivery() call this after they
//  complete their own gate — whichever call finds all three true
//  is the one that actually fulfills. Idempotent (fulfillProjectAndQuotation
//  itself no-ops if already Fulfilled), so it's safe to call this
//  speculatively every time.
// ============================================================
function maybeFulfillProject(projectId, quotationId) {
  const pSheet    = getSheet('Projects');
  const paySheet  = getSheet('Payments');
  const distSheet = getSheet('Revenue_Distribution');
  if (!pSheet) return;

  const pData = pSheet.getDataRange().getValues();
  let pRow = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) { pRow = pData[i]; break; }
  }
  if (!pRow) return;
  if (pRow[P.STATUS] === 'Fulfilled') return; // already done — nothing to check

  const totalAmount = Number(pRow[P.TOTAL_AMOUNT] || 0);
  const amountPaid = paySheet ? paySheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[PAY.PROJECT_ID]) === String(projectId))
    .reduce((s, r) => s + (Number(r[PAY.AMOUNT]) || 0), 0) : 0;
  const isFullyPaid = totalAmount > 0 && amountPaid >= totalAmount;

  const isDistributed = distSheet ? distSheet.getDataRange().getValues().slice(1)
    .some(r => String(r[DIST.PROJECT_ID]) === String(projectId)) : false;

  const isDelivered = !!pRow[P.DELIVERED_DATE];

  if (isFullyPaid && isDistributed && isDelivered) {
    fulfillProjectAndQuotation(projectId, quotationId);
  }
}

// ============================================================
//  DELIVERY CONFIRMATION — applies to ALL projects (free and
//  normal). Sets Projects.DeliveredDate. For a free project this
//  IS the fulfillment trigger (no payment/distribution gates
//  apply); for a normal project it's one of three independent
//  gates checked by maybeFulfillProject().
//
//  Validation (re-checked here even though the client already
//  enforces it, per "fail loudly" / never trust the client):
//    • date is required and must parse
//    • cannot be after today (no future dates)
//    • cannot be before the quotation's Date Issued
//
//  A free project that has already reached Fulfilled cannot have
//  its delivery date edited (that action is what fulfilled it) —
//  a normal Fulfilled project CAN still have its date edited,
//  since delivery is otherwise independent of payment timing.
// ============================================================
function confirmDelivery(projectId, dateStr) {
  const pSheet = getSheet('Projects');
  const qSheet = getSheet('Quotations');
  if (!pSheet || !qSheet) return { success: false, error: 'Required sheets missing.' };

  const pData = pSheet.getDataRange().getValues();
  let pRowIndex = -1, pRow = null;
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      pRowIndex = i + 1; pRow = pData[i]; break;
    }
  }
  if (!pRow) return { success: false, error: 'Project not found.' };

  const quotationId = String(pRow[P.QUOTATION_ID] || '');
  const isFree       = isProjectFree(quotationId);

  if (isFree && pRow[P.STATUS] === 'Fulfilled') {
    return { success: false, error: 'This free project is already Fulfilled — the delivery date cannot be changed.' };
  }

  // ── Validate date ──────────────────────────────────────────
  if (!dateStr) return { success: false, error: 'A delivery date is required.' };
  const parsed = new Date(dateStr);
  if (isNaN(parsed.getTime())) return { success: false, error: 'Invalid delivery date.' };

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (dateStr > todayStr) {
    return { success: false, error: 'Delivery date cannot be in the future.' };
  }

  let qDateIssuedStr = '';
  const qData = qSheet.getDataRange().getValues();
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      qDateIssuedStr = qData[i][Q.DATE_ISSUED]
        ? Utilities.formatDate(new Date(qData[i][Q.DATE_ISSUED]), tz, 'yyyy-MM-dd') : '';
      break;
    }
  }
  if (qDateIssuedStr && dateStr < qDateIssuedStr) {
    return { success: false, error: 'Delivery date cannot be before the quotation date (' + qDateIssuedStr + ').' };
  }

  const user   = Session.getActiveUser().getEmail();
  const now    = new Date();
  const oldVal = pRow[P.DELIVERED_DATE]
    ? Utilities.formatDate(new Date(pRow[P.DELIVERED_DATE]), tz, 'yyyy-MM-dd') : '';

  pSheet.getRange(pRowIndex, P.DELIVERED_DATE + 1).setValue(parsed);

  writeLog('Projects_Log', 'Projects', projectId,
    quotationId + ' — ' + pRow[P.PROJECT_NAME], pRow[P.ACCOUNT_ID], pRow[P.ACCOUNT_NAME],
    'DeliveredDate', oldVal, dateStr,
    'Delivery date ' + (oldVal ? 'edited' : 'confirmed') + ' by ' + user);

  if (isFree) {
    // Free projects: delivery confirmation IS the fulfillment
    // trigger — reuses the same routine a normal project uses, no
    // duplicated logic, and it creates no ledger/team/revenue
    // entries because none exist to create (no Payments/
    // Revenue_Distribution rows ever get written for a free project).
    fulfillProjectAndQuotation(projectId, quotationId);
  } else {
    maybeFulfillProject(projectId, quotationId);
  }

  return { success: true, deliveredDate: dateStr };
}

// ============================================================
//  FULFILLMENT — the quotation/project becomes "Fulfilled" once
//  it's Confirmed (implicit — a Project row only exists post-
//  confirm) and all applicable gates are satisfied: for a normal
//  project that's fully paid + distributed + delivered (checked by
//  maybeFulfillProject); for a free project it's just delivered
//  (checked directly in confirmDelivery). This function itself
//  doesn't re-check anything — callers are responsible for only
//  calling it once their gates are satisfied.
//
//  On trigger:
//    • Quotation PDF moved Confirmed → Fulfilled (Drive)
//    • Quotation.STATUS  → 'Fulfilled'
//    • Project.STATUS    → 'Fulfilled', COMPLETED_AT stamped
//    • Both changes logged
//
//  Idempotent — safe to call again (e.g. on a distribution edit
//  after the fact): if the quotation is already 'Fulfilled' the
//  PDF move/status/log are skipped so nothing double-fires or
//  tries to move an already-moved file.
// ============================================================
function fulfillProjectAndQuotation(projectId, quotationId) {
  const pSheet = getSheet('Projects');
  const qSheet = getSheet('Quotations');
  if (!pSheet || !qSheet) return;

  const user = Session.getActiveUser().getEmail();
  const now  = new Date();

  // ── Quotation side ───────────────────────────────────────
  const qData = qSheet.getDataRange().getValues();
  let qRowIndex = -1, qRow = null;
  for (let i = 1; i < qData.length; i++) {
    if (String(qData[i][Q.QUOTATION_ID]) === String(quotationId)) {
      qRowIndex = i + 1; qRow = qData[i]; break;
    }
  }

  if (qRow && qRow[Q.STATUS] !== 'Fulfilled') {
    const pdfUrl = extractUrl(qRow[Q.FOLDER_URL]) || '';
    try {
      const pdfFile = getQuotationPdfFile(pdfUrl);
      if (pdfFile) {
        const fulfilledFolder = DriveApp.getFolderById(RUYA_QUOTATIONS_FULFILLED_ID);
        const parents = pdfFile.getParents();
        fulfilledFolder.addFile(pdfFile);
        while (parents.hasNext()) parents.next().removeFile(pdfFile);
      }
    } catch(e) { Logger.log('Fulfillment PDF move error: ' + e); }

    const oldStatus = qRow[Q.STATUS];
    qSheet.getRange(qRowIndex, Q.STATUS          + 1).setValue('Fulfilled');
    qSheet.getRange(qRowIndex, Q.LAST_UPDATED_BY + 1).setValue(user);
    qSheet.getRange(qRowIndex, Q.LAST_UPDATED_AT + 1).setValue(now);

    writeLog('Quotations_Log', 'Quotations', quotationId,
      quotationId + ' — ' + qRow[Q.PROJECT_NAME], qRow[Q.ACCOUNT_ID], qRow[Q.ACCOUNT_NAME],
      'Status', oldStatus, 'Fulfilled',
      'Quotation fulfilled — all applicable conditions met; PDF moved to Fulfilled');
  }

  // ── Project side ─────────────────────────────────────────
  const pData = pSheet.getDataRange().getValues();
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][P.PROJECT_ID]) === String(projectId)) {
      if (pData[i][P.STATUS] !== 'Fulfilled') {
        const oldStatus = pData[i][P.STATUS];
        pSheet.getRange(i + 1, P.STATUS       + 1).setValue('Fulfilled');
        pSheet.getRange(i + 1, P.COMPLETED_AT + 1).setValue(now);

        writeLog('Projects_Log', 'Projects', projectId,
          quotationId + ' — ' + pData[i][P.PROJECT_NAME], pData[i][P.ACCOUNT_ID], pData[i][P.ACCOUNT_NAME],
          'Status', oldStatus, 'Fulfilled', 'Project fulfilled — all applicable conditions met');
      }
      break;
    }
  }
}
