// ============================================================
// BIB LOANS — Google Apps Script Backend (code.gs)
// Build, Invest, Borrow | Loans Made Simple
// Interest Rate: 10% per month (flat)
// ============================================================

// ── CONFIGURATION ────────────────────────────────────────────────────────────
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE', // <-- Replace with your Sheet ID
  SHEETS: {
    BORROWERS: 'Borrowers',
    LOANS:     'Loans',
    PAYMENTS:  'Payments',
    REQUESTS:  'LoanRequests',
    SETTINGS:  'Settings'
  },
  INTEREST_RATE: 0.10,          // 10% per month
  APP_TITLE:     'BIB Loans',
  APP_VERSION:   '1.0.0'
};

// ── WEB APP ENTRY POINTS ──────────────────────────────────────────────────────

function doGet(e) {
  // Allow GET-based API calls (avoids CORS preflight issues from standalone HTML)
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = e.parameter.action;
      const data   = e.parameter.data ? JSON.parse(e.parameter.data) : {};
      switch (action) {
        case 'getLoanByBorrower': return respond(getLoanByBorrower(data));
        case 'getRequests':       return respond(getRequests());
        case 'updateRequestStatus': return respond(updateRequestStatus(data));
        case 'getDashboard':      return respond(getDashboard());
        case 'getBorrowers':      return respond(getBorrowers());
        case 'getLoans':          return respond(getLoans(data.borrowerId));
        default:                  return respond({ error: 'Unknown action: ' + action }, false);
      }
    } catch (err) {
      return respond({ error: err.message }, false);
    }
  }
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('BIB Loans — Loans Made Simple')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    const data    = payload.data || {};

    switch (action) {
      case 'getBorrowers':  return respond(getBorrowers());
      case 'addBorrower':    return respond(addBorrower(data));
      case 'updateBorrower': return respond(updateBorrower(data));
      case 'getLoans':      return respond(getLoans(data.borrowerId));
      case 'addLoan':       return respond(addLoan(data));
      case 'addPayment':    return respond(addPayment(data));
      case 'deletePayment': return respond(deletePayment(data));
      case 'getDashboard':    return respond(getDashboard());
      case 'getLoanDetail':   return respond(getLoanDetail(data.loanId));
      case 'addLoanRequest':        return respond(addLoanRequest(data));
      case 'getRequests':           return respond(getRequests());
      case 'updateRequestStatus':   return respond(updateRequestStatus(data));
      case 'getLoanByBorrower':     return respond(getLoanByBorrower(data));
      default:                  return respond({ error: 'Unknown action: ' + action }, false);
    }
  } catch (err) {
    return respond({ error: err.message }, false);
  }
}

function respond(data, success = true) {
  const output = ContentService.createTextOutput(
    JSON.stringify({ success, data })
  );
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── SHEET HELPERS ─────────────────────────────────────────────────────────────

function getSheet(name) {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) sheet = createSheet(ss, name);
  return sheet;
}

function createSheet(ss, name) {
  const sheet = ss.insertSheet(name);
  const headers = {
    Borrowers: ['ID', 'Name', 'Department', 'Employee No', 'Phone', 'Email', 'Date Added', 'Notes', 'ATM PIN'],
    Loans:     ['Loan ID', 'Borrower ID', 'Borrower Name', 'Principal', 'Interest Rate', 'Term (Months)',
                'Monthly Payment', 'Total Payable', 'Date Released', 'Due Date', 'Status', 'Purpose', 'Notes'],
    Payments:  ['Payment ID', 'Loan ID', 'Borrower ID', 'Amount', 'Date', 'Running Balance', 'Notes'],
    LoanRequests: ['Ref ID', 'Name', 'Phone', 'Dept', 'Employee No', 'Amount', 'Term', 'Purpose', 'Notes', 'ATM PIN', 'Date Submitted', 'Status'],
    Settings:  ['Key', 'Value']
  };
  if (headers[name]) {
    sheet.appendRow(headers[name]);
    sheet.getRange(1, 1, 1, headers[name].length)
      .setFontWeight('bold')
      .setBackground('#1a4fba')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateId(prefix) {
  return prefix + '-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
}

// ── BORROWER FUNCTIONS ────────────────────────────────────────────────────────

function getBorrowers() {
  const sheet = getSheet(CONFIG.SHEETS.BORROWERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(row => ({
    id:         row[0],
    name:       row[1],
    department: row[2],
    employeeNo: String(row[3] ?? ''),
    phone:      String(row[4] ?? ''),
    email:      row[5],
    dateAdded:  row[6] ? Utilities.formatDate(new Date(row[6]), 'Asia/Manila', 'yyyy-MM-dd') : '',
    notes:      row[7],
    atmPin:     row[8] ? String(row[8]) : ''
  })).filter(b => b.id);
}

function addBorrower(data) {
  const sheet = getSheet(CONFIG.SHEETS.BORROWERS);
  const id    = generateId('BOR');
  const now   = new Date();

  sheet.appendRow([
    id,
    data.name       || '',
    data.department || '',
    data.employeeNo || '',
    (data.phone || '').replace(/\D/g, '').replace(/^0+/, ''), // store without leading 0
    data.email      || '',
    now,
    data.notes      || '',
    data.atmPin     || ''
  ]);

  return { id, message: 'Borrower added successfully.' };
}

function updateBorrower(data) {
  const sheet = getSheet(CONFIG.SHEETS.BORROWERS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      if (data.name       !== undefined) sheet.getRange(i + 1, 2).setValue(data.name);
      if (data.department !== undefined) sheet.getRange(i + 1, 3).setValue(data.department);
      if (data.employeeNo !== undefined) sheet.getRange(i + 1, 4).setValue(data.employeeNo);
      if (data.phone      !== undefined) sheet.getRange(i + 1, 5).setValue((data.phone || '').replace(/\D/g, '').replace(/^0+/, ''));
      if (data.notes      !== undefined) sheet.getRange(i + 1, 8).setValue(data.notes);
      if (data.atmPin     !== undefined) sheet.getRange(i + 1, 9).setValue(data.atmPin); // column I
      return { success: true, message: 'Borrower updated.' };
    }
  }
  return { success: false, message: 'Borrower not found.' };
}

// ── LOAN FUNCTIONS ────────────────────────────────────────────────────────────

function getLoans(borrowerId) {
  const sheet = getSheet(CONFIG.SHEETS.LOANS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let loans = data.slice(1).map(row => ({
    loanId:         row[0],
    borrowerId:     row[1],
    borrowerName:   row[2],
    principal:      parseFloat(row[3])  || 0,
    interestRate:   parseFloat(row[4])  || 0,
    term:           parseInt(row[5])    || 0,
    monthlyPayment: parseFloat(row[6])  || 0,
    totalPayable:   parseFloat(row[7])  || 0,
    dateReleased:   row[8] ? Utilities.formatDate(new Date(row[8]), 'Asia/Manila', 'yyyy-MM-dd') : '',
    dueDate:        row[9] ? Utilities.formatDate(new Date(row[9]), 'Asia/Manila', 'yyyy-MM-dd') : '',
    status:         row[10] || 'Active',
    purpose:        row[11] || '',
    notes:          row[12] || ''
  })).filter(l => l.loanId);

  if (borrowerId) loans = loans.filter(l => l.borrowerId === borrowerId);

  // Enrich each loan with payment totals so progress bars work on all devices
  return loans.map(loan => {
    const payments  = getPaymentsForLoan(loan.loanId);
    const totalPaid = payments.reduce((s, p) => s + p.amount, 0);
    const balance   = Math.max(0, loan.totalPayable - totalPaid);
    return { ...loan, totalPaid, balance };
  });
}

function addLoan(data) {
  const sheet   = getSheet(CONFIG.SHEETS.LOANS);
  const loanId  = generateId('LN');
  const principal     = parseFloat(data.principal) || 0;
  const term          = parseInt(data.term)        || 1;
  const interestRate  = CONFIG.INTEREST_RATE;         // 10% per month

  // Flat interest calculation
  const totalInterest  = principal * interestRate * term;
  const totalPayable   = principal + totalInterest;
  const monthlyPayment = totalPayable / term;

  const dateReleased  = new Date(data.dateReleased || new Date());
  const dueDate       = new Date(dateReleased);
  dueDate.setMonth(dueDate.getMonth() + term);

  sheet.appendRow([
    loanId,
    data.borrowerId    || '',
    data.borrowerName  || '',
    principal,
    interestRate,
    term,
    monthlyPayment,
    totalPayable,
    dateReleased,
    dueDate,
    'Active',
    data.purpose       || '',
    data.notes         || ''
  ]);

  return {
    loanId,
    principal,
    totalInterest,
    totalPayable,
    monthlyPayment,
    dueDate: Utilities.formatDate(dueDate, 'Asia/Manila', 'yyyy-MM-dd'),
    message: 'Loan created successfully.'
  };
}

function getLoanDetail(loanId) {
  const loans    = getLoans();
  const loan     = loans.find(l => l.loanId === loanId);
  if (!loan) return { error: 'Loan not found.' };

  const payments = getPaymentsForLoan(loanId);
  const totalPaid    = payments.reduce((s, p) => s + p.amount, 0);
  const balance      = loan.totalPayable - totalPaid;

  return { ...loan, payments, totalPaid, balance };
}

// ── PAYMENT FUNCTIONS ─────────────────────────────────────────────────────────

function getPaymentsForLoan(loanId) {
  const sheet = getSheet(CONFIG.SHEETS.PAYMENTS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .filter(row => row[1] === loanId)
    .map(row => ({
      paymentId:      row[0],
      loanId:         row[1],
      borrowerId:     row[2],
      amount:         parseFloat(row[3]) || 0,
      date:           row[4] ? Utilities.formatDate(new Date(row[4]), 'Asia/Manila', 'yyyy-MM-dd') : '',
      runningBalance: parseFloat(row[5]) || 0,
      notes:          row[6] || ''
    }));
}

function addPayment(data) {
  const sheet     = getSheet(CONFIG.SHEETS.PAYMENTS);
  const paymentId = generateId('PAY');
  const amount    = parseFloat(data.amount) || 0;

  // Calculate running balance
  const loanDetail     = getLoanDetail(data.loanId);
  const runningBalance = (loanDetail.balance || 0) - amount;

  sheet.appendRow([
    paymentId,
    data.loanId     || '',
    data.borrowerId || '',
    amount,
    new Date(data.date || new Date()),
    Math.max(0, runningBalance),
    data.notes      || ''
  ]);

  // Auto-mark loan as Paid if balance <= 0
  if (runningBalance <= 0) {
    updateLoanStatus(data.loanId, 'Paid');
  }

  return { paymentId, runningBalance: Math.max(0, runningBalance), message: 'Payment recorded.' };
}

function deletePayment(data) {
  const sheet = getSheet(CONFIG.SHEETS.PAYMENTS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.paymentId) {
      sheet.deleteRow(i + 1);
      // Recalculate running balances for remaining payments on this loan
      recalcRunningBalances(data.loanId);
      // Recalculate loan status (may need to flip back to Active from Paid)
      recalcLoanStatus(data.loanId);
      return { success: true, message: 'Payment deleted.' };
    }
  }
  return { success: false, message: 'Payment not found.' };
}

function recalcRunningBalances(loanId) {
  const sheet    = getSheet(CONFIG.SHEETS.PAYMENTS);
  const rows     = sheet.getDataRange().getValues();
  const loanData = getLoanDetail(loanId);
  let balance    = loanData.totalPayable || 0;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === loanId) {
      balance -= (parseFloat(rows[i][3]) || 0);
      sheet.getRange(i + 1, 6).setValue(Math.max(0, balance));
    }
  }
}

function recalcLoanStatus(loanId) {
  const loanDetail = getLoanDetail(loanId);
  const status     = (loanDetail.balance <= 0) ? 'Paid' : 'Active';
  updateLoanStatus(loanId, status);
}

function updateLoanStatus(loanId, status) {
  const sheet = getSheet(CONFIG.SHEETS.LOANS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === loanId) {
      sheet.getRange(i + 1, 11).setValue(status);
      break;
    }
  }
}

// ── DASHBOARD SUMMARY ─────────────────────────────────────────────────────────

function getDashboard() {
  const loans     = getLoans();
  const active    = loans.filter(l => l.status === 'Active');
  const paid      = loans.filter(l => l.status === 'Paid');
  const overdue   = active.filter(l => new Date(l.dueDate) < new Date());

  const totalPrincipal = active.reduce((s, l) => s + l.principal, 0);
  const totalReceivable = active.reduce((s, l) => s + l.totalPayable, 0);

  // Collect all payments
  const sheet    = getSheet(CONFIG.SHEETS.PAYMENTS);
  const pData    = sheet.getDataRange().getValues();
  const payments = pData.length > 1 ? pData.slice(1) : [];
  const totalCollected = payments.reduce((s, r) => s + (parseFloat(r[3]) || 0), 0);

  const borrowers = getBorrowers();

  return {
    totalBorrowers:  borrowers.length,
    activeLoans:     active.length,
    paidLoans:       paid.length,
    overdueLoans:    overdue.length,
    totalPrincipal,
    totalReceivable,
    totalCollected,
    recentLoans:     loans.slice(-10).reverse()
  };
}

// ── UTILITIES ─────────────────────────────────────────────────────────────────

function initializeSheets() {
  Object.values(CONFIG.SHEETS).forEach(name => getSheet(name));
  return { message: 'Sheets initialized.' };
}

function fixLoanRequestsHeader() {
  const sheet   = getSheet(CONFIG.SHEETS.REQUESTS);
  const headers = ['Ref ID', 'Name', 'Phone', 'Dept', 'Employee No', 'Amount', 'Term', 'Purpose', 'Notes', 'ATM PIN', 'Date Submitted', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight('bold')
       .setBackground('#1a4fba')
       .setFontColor('#ffffff');
  return { message: 'LoanRequests header updated.' };
}

// ── LOAN REQUEST FUNCTIONS ────────────────────────────────────────

function getRequests() {
  const sheet = getSheet(CONFIG.SHEETS.REQUESTS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({
    refId:   row[0],
    name:    row[1],
    phone:   String(row[2] ?? ''),
    dept:    row[3],
    empNo:   String(row[4] ?? ''),
    amount:  parseFloat(row[5]) || 0,
    term:    parseInt(row[6])   || 0,
    purpose: row[7],
    notes:   row[8],
    atmPin:  row[9] ? String(row[9]) : '',
    date:    row[10] ? Utilities.formatDate(new Date(row[10]), 'Asia/Manila', 'yyyy-MM-dd') : '',
    status:  row[11] || 'Pending'
  })).filter(r => r.refId);
}

function addLoanRequest(data) {
  const sheet = getSheet(CONFIG.SHEETS.REQUESTS);
  sheet.appendRow([
    data.refId       || generateId('REQ'),
    data.name        || '',
    data.phone       || '',
    data.dept        || '',
    data.empNo       || '',
    parseFloat(data.amount) || 0,
    parseInt(data.term)     || 0,
    data.purpose     || '',
    data.notes       || '',
    data.atmPin      || '',
    new Date(),
    'Pending'
  ]);
  return { message: 'Request saved.' };
}

function updateRequestStatus(data) {
  const sheet = getSheet(CONFIG.SHEETS.REQUESTS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.refId) {
      sheet.getRange(i + 1, 12).setValue(data.status); // column L = Status
      if (data.reason) sheet.getRange(i + 1, 13).setValue(data.reason);
      return { message: 'Status updated.' };
    }
  }
  return { error: 'Request not found.' };
}

// ── BORROWER SELF-LOOKUP ──────────────────────────────────────────

function getLoanByBorrower(data) {
  // Normalize phone: strip non-digits then strip leading 0 (handles both 09XXXXXXXXX and 9XXXXXXXXX)
  const phone      = (data.phone      || '').trim().replace(/\D/g, '').replace(/^0+/, '');
  const employeeNo = (data.employeeNo || '').trim().toLowerCase();

  if (!phone && !employeeNo) {
    return { error: 'Please provide a phone number or employee number.' };
  }

  // Find matching borrower
  const borrowers = getBorrowers();
  const borrower  = borrowers.find(b => {
    const bPhone = String(b.phone      ?? '').replace(/\D/g, '').replace(/^0+/, '');
    const bEmpNo = String(b.employeeNo ?? '').trim().toLowerCase();
    // Must match BOTH fields when both are provided, OR at least one if only one given
    const phoneMatch  = phone      && bPhone === phone;
    const empNoMatch  = employeeNo && bEmpNo === employeeNo;
    if (phone && employeeNo) return phoneMatch && empNoMatch;
    if (phone)      return phoneMatch;
    if (employeeNo) return empNoMatch;
    return false;
  });

  if (!borrower) {
    return { error: 'No records found. Please double-check your phone number and employee number.' };
  }

  // Get all loans for this borrower
  const allLoans = getLoans(borrower.id);
  if (!allLoans.length) {
    return { error: 'No active loans found for your account.' };
  }

  // Enrich each loan with payment data
  const loans = allLoans.map(loan => {
    const payments     = getPaymentsForLoan(loan.loanId);
    const totalPaid    = payments.reduce((s, p) => s + p.amount, 0);
    const balance      = Math.max(0, loan.totalPayable - totalPaid);
    const progressPct  = loan.totalPayable > 0
      ? Math.min(100, Math.round((totalPaid / loan.totalPayable) * 100))
      : 0;
    return { ...loan, payments, totalPaid, balance, progressPct };
  });

  return {
    borrower: {
      id:         borrower.id,
      name:       borrower.name,
      department: borrower.department,
      employeeNo: borrower.employeeNo,
      phone:      borrower.phone
    },
    loans
  };
}
