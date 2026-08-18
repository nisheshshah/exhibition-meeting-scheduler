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

    // Send automated email confirmation to client & sales representative
    sendConfirmationEmail(params, ref, salesmanName);

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
    const newDate = params.date;
    const newTime = params.time;

    let targetIndex = -1;
    let targetExhibitionId = null;
    let targetSalesmanId = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === ref) {
        targetIndex = i;
        targetExhibitionId = data[i][2];
        targetSalesmanId = data[i][3];
        break;
      }
    }

    if (targetIndex === -1) {
      lock.releaseLock();
      return jsonResponse({ status: "error", message: "Booking reference not found" });
    }

    // Make sure the new date/time isn't already taken by a DIFFERENT
    // booking for this same rep before we overwrite anything.
    for (let i = 1; i < data.length; i++) {
      if (i === targetIndex) continue;
      const row = data[i];
      const rowStatus = row[12] || 'confirmed';
      const rowDate = formatDateStr(row[5]);
      const rowTime = row[6];
      if (rowStatus !== 'cancelled' && row[2] === targetExhibitionId && row[3] === targetSalesmanId && rowDate === newDate && rowTime === newTime) {
        lock.releaseLock();
        return jsonResponse({ status: "error", message: "That slot is already booked. Please choose another time." });
      }
    }

    const i = targetIndex;
    sheet.getRange(i + 1, 6).setValue("'" + newDate);
    sheet.getRange(i + 1, 7).setValue("'" + newTime);
    sheet.getRange(i + 1, 8).setValue(params.firstName);
    sheet.getRange(i + 1, 9).setValue(params.lastName);
    sheet.getRange(i + 1, 10).setValue(params.company);
    sheet.getRange(i + 1, 11).setValue(params.email);
    sheet.getRange(i + 1, 12).setValue(params.phone);

    lock.releaseLock();
    return jsonResponse({ status: "success", message: "Booking updated" });
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

function getSalesmanEmail(id) {
  const map = {
    'S002': 'jaishah@universaloleoresins.com',
    'S003': 'nisheshshah@universaloleoresins.com',
    'S005': 'intsales@universaloleoresins.com',
    'S006': 'domsaleseast@universaloleoresins.com',
    'S007': 'harshitashah@universaloleoresins.com',
    'S009': 'domsales@universaloleoresins.com',
    'S010': 'domsalesmh@universaloleoresins.com',
    'S011': 'shishirshah@xtractiva.com',
    'S013': 'domsalesgj@universaloleoresins.com',
    'S014': 'domsalessouth@universaloleoresins.com',
    'S015': 'horecawest@universaloleoresins.com',
    'S020': 'jigeshshah@universaloleoresins.com',
    'S021': 'padmadeshraj@universaloleoresins.com'
  };
  return map[id] || '';
}

function sendConfirmationEmail(params, ref, salesmanName) {
  try {
    const clientEmail = params.email;
    const salesmanEmail = getSalesmanEmail(params.salesmanId || params.s);
    if (!clientEmail) return;

    const subject = `Meeting Confirmation: Universal Oleoresins [Ref: ${ref}]`;
    const bodyText = `Dear ${params.firstName} ${params.lastName},\n\n` +
      `Your meeting with ${salesmanName} at Universal Oleoresins is confirmed!\n\n` +
      `Meeting Details:\n` +
      `- Reference Code: ${ref}\n` +
      `- Date: ${params.date}\n` +
      `- Time: ${params.time} IST\n` +
      `- Company: ${params.company}\n` +
      `- Sales Representative: ${salesmanName}\n` +
      `- Location: Stall 3D38, Hall 3 (BEC, Goregaon, Mumbai)\n\n` +
      `Thank you,\nUniversal Oleoresins Team`;

    const bodyHtml = `
      <div style="font-family: Arial, sans-serif; color: #1a1410; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden;">
        <div style="background-color: #d54e1f; padding: 20px; text-align: center; color: white;">
          <h2 style="margin: 0; font-size: 22px;">UNIVERSAL OLEORESINS</h2>
          <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.9;">Meeting Confirmation · Fi India 2026</p>
        </div>
        <div style="padding: 24px; background-color: #ffffff;">
          <p style="font-size: 16px; margin-top: 0;">Dear <strong>${params.firstName} ${params.lastName}</strong>,</p>
          <p style="font-size: 14px; color: #4a3f33; line-height: 1.5;">Your meeting has been successfully confirmed. Here are your reservation details:</p>
          
          <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fbf7f0; border-radius: 8px; overflow: hidden;">
            <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold; width: 40%;">Reference Code:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${ref}</td></tr>
            <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Date & Time:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${params.date} at ${params.time} IST</td></tr>
            <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Representative:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${salesmanName}</td></tr>
            <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Company:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${params.company}</td></tr>
            <tr><td style="padding: 10px 14px; font-weight: bold;">Location:</td><td style="padding: 10px 14px;">Stall 3D38, Hall 3 (BEC, Goregaon, Mumbai)</td></tr>
          </table>

          <p style="font-size: 13px; color: #666;">We look forward to meeting you at the exhibition!</p>
        </div>
      </div>
    `;

    // Send email to client
    MailApp.sendEmail({
      to: clientEmail,
      subject: subject,
      body: bodyText,
      htmlBody: bodyHtml
    });

    // Also send copy to Sales Representative if valid email exists
    if (salesmanEmail && salesmanEmail !== clientEmail) {
      MailApp.sendEmail({
        to: salesmanEmail,
        subject: `New Meeting Booked: ${params.firstName} ${params.lastName} (${params.company}) [Ref: ${ref}]`,
        body: bodyText,
        htmlBody: bodyHtml
      });
    }
  } catch(e) {
    console.warn("Could not send confirmation email:", e);
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}