/**
 * ====================================================================
 * UNIVERSAL OLEORESINS - EXHIBITION MEETING SCHEDULER BACKEND
 * Target Google Sheet ID: 1YLKPpuAhTqvvfgwUi29l-U7vws8bHwdT5y41QysqGn0
 * ====================================================================
 */

const SHEET_NAME = "Bookings";

function doGet(e) {
  const params = e ? e.parameter : {};
  const action = params.action || 'fetch';

  if (action === 'book') {
    return handleBook(params);
  } else if (action === 'update') {
    return handleUpdate(params);
  } else if (action === 'cancel') {
    return handleCancel(params);
  } else if (action === 'report') {
    return handleReport(params);
  } else {
    return handleFetch(params);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'book') {
      return handleBook(data);
    } else if (data.action === 'cancel') {
      return handleCancel(data);
    } else {
      return handleFetch(data);
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp", "Reference", "Exhibition", "Salesman ID", "Salesman Name",
      "Date", "Time", "First Name", "Last Name", "Company", "Email", "Phone", "Status"
    ]);
    sheet.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#d54e1f").setFontColor("#ffffff");
  }
  return sheet;
}

function handleFetch(params) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const bookings = [];

  if (data.length > 1) {
    const salesmanFilter = params.salesmanId || params.s;
    const exhibitionFilter = params.exhibitionId || params.e;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[12] || 'confirmed';
      if (status === 'cancelled') continue;

      const sid = row[3];
      const exId = row[2];

      if (salesmanFilter && sid !== salesmanFilter) continue;
      if (exhibitionFilter && exId !== exhibitionFilter) continue;

      bookings.push({
        ref: row[1],
        exhibitionId: row[2],
        salesmanId: row[3],
        salesmanName: row[4],
        date: formatDateStr(row[5]),
        time: row[6],
        firstName: row[7],
        lastName: row[8],
        company: row[9],
        email: row[10],
        phone: row[11],
        status: status
      });
    }
  }

  return jsonResponse(bookings);
}

function handleBook(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return jsonResponse({ status: "error", message: "Server busy. Please try again." });
  }

  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();

    const salesmanId = params.salesmanId || params.s;
    const dateStr = params.date;
    const timeStr = params.time;
    const exhibitionId = params.exhibitionId || "fi-india-2026";

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowStatus = row[12] || 'confirmed';
      const rowExId = row[2];
      const rowSid = row[3];
      const rowDate = formatDateStr(row[5]);
      const rowTime = row[6];

      if (rowStatus !== 'cancelled' && rowExId === exhibitionId && rowSid === salesmanId && rowDate === dateStr && rowTime === timeStr) {
        lock.releaseLock();
        return jsonResponse({ status: "error", message: "Slot already booked by another user." });
      }
    }

    const timestamp = new Date().toISOString();
    const ref = params.ref || ('UO-' + Math.random().toString(36).substring(2, 8).toUpperCase());
    const salesmanName = params.salesmanName || getSalesmanName(salesmanId);

    sheet.appendRow([
      timestamp,
      ref,
      exhibitionId,
      salesmanId,
      salesmanName,
      "'" + dateStr, // Force text to prevent Google Sheet auto-formatting dates
      "'" + timeStr, // Force text to prevent Google Sheet auto-formatting times
      params.firstName,
      params.lastName,
      params.company,
      params.email,
      params.phone,
      "confirmed"
    ]);

    lock.releaseLock();
    return jsonResponse({ status: "success", ref: ref, date: dateStr, time: timeStr });
  } catch (err) {
    lock.releaseLock();
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ==========================================
// UPDATE ACTION (Edit Meeting)
// ==========================================
function handleUpdate(params) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return jsonResponse({ status: "error", message: "Server busy" }); }

  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const ref = params.ref;

    let found = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === ref) {
        // Update Date, Time, First, Last, Company, Email, Phone
        sheet.getRange(i + 1, 6).setValue("'" + params.date); 
        sheet.getRange(i + 1, 7).setValue("'" + params.time); 
        sheet.getRange(i + 1, 8).setValue(params.firstName);
        sheet.getRange(i + 1, 9).setValue(params.lastName);
        sheet.getRange(i + 1, 10).setValue(params.company);
        sheet.getRange(i + 1, 11).setValue(params.email);
        sheet.getRange(i + 1, 12).setValue(params.phone);
        found = true;
        break;
      }
    }

    lock.releaseLock();
    if (found) {
      return jsonResponse({ status: "success", message: "Booking updated" });
    } else {
      return jsonResponse({ status: "error", message: "Booking reference not found" });
    }
  } catch (err) {
    lock.releaseLock();
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ==========================================
// DELETE ACTION (Cancel/Delete Booking Permanently)
// ==========================================
function handleCancel(params) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return jsonResponse({ status: "error", message: "Server busy" }); }

  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const ref = params.ref;

    let found = false;
    // Loop backwards when deleting rows to not mess up the row indexes
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (ref && row[1] === ref) {
        sheet.deleteRow(i + 1); // Permanently deletes the row from Google Sheets
        found = true;
      }
    }

    lock.releaseLock();
    if (found) {
      return jsonResponse({ status: "success", message: "Booking deleted permanently" });
    } else {
      return jsonResponse({ status: "error", message: "Booking reference not found" });
    }
  } catch (err) {
    lock.releaseLock();
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function handleReport(params) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const stats = {
    totalBookings: 0,
    cancelledBookings: 0,
    bySalesman: {},
    byDate: {},
    byCompany: {}
  };

  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[12] || 'confirmed';
      const sid = row[3];
      const sname = row[4];
      const date = formatDateStr(row[5]);
      const company = row[9];

      if (status === 'cancelled') {
        stats.cancelledBookings++;
        continue;
      }

      stats.totalBookings++;

      if (!stats.bySalesman[sid]) {
        stats.bySalesman[sid] = { name: sname, count: 0 };
      }
      stats.bySalesman[sid].count++;

      stats.byDate[date] = (stats.byDate[date] || 0) + 1;

      if (company) {
        stats.byCompany[company] = (stats.byCompany[company] || 0) + 1;
      }
    }
  }

  return jsonResponse({ status: "success", report: stats });
}

function formatDateStr(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(d);
}

function getSalesmanName(id) {
  const map = {
    'S002': 'Jai Shah',
    'S003': 'Nishesh Shah',
    'S005': 'Paul Thampy',
    'S006': 'Bikash Kar',
    'S007': 'Harshita Shah',
    'S009': 'Payal',
    'S010': 'Machindranath',
    'S011': 'Shishir Shah',
    'S013': 'Saurabh',
    'S014': 'Kiruthi Kumar',
    'S015': 'Shubham',
    'S020': 'Jigesh Shah'
  };
  return map[id] || id;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
