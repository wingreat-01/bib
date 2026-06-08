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
      case 'addBorrower':   return respond(addBorrower(data));
      case 'getLoans':      return respond(getLoans(data.borrowerId));
      case 'addLoan':       return respond(addLoan(data));
      case 'addPayment':    return respond(addPayment(data));
      case 'getDashboard':    return respond(getDashboard());
      case 'getLoanDetail':   return respond(getLoanDetail(data.loanId));
      case 'addLoanRequest':  return respond(addLoanRequest(data));
      default:                return respond({ error: 'Unknown action: ' + action }, false);
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
    Borrowers: ['ID', 'Name', 'Department', 'Employee No', 'Phone', 'Email', 'Date Added', 'Notes'],
    Loans:     ['Loan ID', 'Borrower ID', 'Borrower Name', 'Principal', 'Interest Rate', 'Term (Months)',
                'Monthly Payment', 'Total Payable', 'Date Released', 'Due Date', 'Status', 'Purpose', 'Notes'],
    Payments:  ['Payment ID', 'Loan ID', 'Borrower ID', 'Amount', 'Date', 'Running Balance', 'Notes'],
    LoanRequests: ['Ref ID', 'Name', 'Phone', 'Dept', 'Employee No', 'Amount', 'Term', 'Purpose', 'Notes', 'Date Submitted', 'Status'],
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
    employeeNo: row[3],
    phone:      row[4],
    email:      row[5],
    dateAdded:  row[6] ? Utilities.formatDate(new Date(row[6]), 'Asia/Manila', 'yyyy-MM-dd') : '',
    notes:      row[7]
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
    data.phone      || '',
    data.email      || '',
    now,
    data.notes      || ''
  ]);

  return { id, message: 'Borrower added successfully.' };
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
  return loans;
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
    recentLoans:     loans.slice(-5).reverse()
  };
}

// ── UTILITIES ─────────────────────────────────────────────────────────────────

function initializeSheets() {
  Object.values(CONFIG.SHEETS).forEach(name => getSheet(name));
  return { message: 'Sheets initialized.' };
}

// ── LOAN REQUEST FUNCTIONS ────────────────────────────────────────

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
    new Date(),
    'Pending'
  ]);
  return { message: 'Request saved.' };
}
