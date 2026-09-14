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
    } else if (data.action === 'update') {
      return handleUpdate(data);
    } else if (data.action === 'cancel') {
      return handleCancel(data);
    } else if (data.action === 'report') {
      return handleReport(data);
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
        phone: formatPhoneStr(row[11]),
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
    const exhibitionId = params.exhibitionId || "default-event";
    const duration = parseInt(params.duration) || 15;
    const slotsToBook = params.slots ? params.slots.split(',') : [timeStr];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowStatus = row[12] || 'confirmed';
      const rowExId = row[2];
      const rowSid = row[3];
      const rowDate = formatDateStr(row[5]);
      const rowTime = row[6];

      if (rowStatus !== 'cancelled' && rowExId === exhibitionId && rowSid === salesmanId && rowDate === dateStr && slotsToBook.indexOf(rowTime) !== -1) {
        lock.releaseLock();
        return jsonResponse({ status: "error", message: "One or more 15-minute lots for this meeting were just booked by another user." });
      }
    }

    const timestamp = new Date().toISOString();
    const ref = params.ref || ('UO-' + Math.random().toString(36).substring(2, 8).toUpperCase());
    const baseRef = ref.split('-P')[0];
    const salesmanName = params.salesmanName || getSalesmanName(salesmanId);

    slotsToBook.forEach((slotTime, idx) => {
      const partRef = idx === 0 ? ref : (baseRef + '-P' + (idx + 1));
      sheet.appendRow([
        timestamp,
        partRef,
        exhibitionId,
        salesmanId,
        salesmanName,
        "'" + dateStr, // Force text to prevent Google Sheet auto-formatting dates
        "'" + slotTime, // Force text to prevent Google Sheet auto-formatting times
        params.firstName,
        params.lastName,
        params.company,
        params.email,
        params.phone ? ("'" + formatPhoneStr(params.phone)) : "",
        "confirmed"
      ]);
    });

    // Send automated email confirmation to client & sales representative (skip if notify=false or silent=true)
    if (!params || (params.notify !== 'false' && params.silent !== 'true' && params.noEmail !== 'true')) {
      sendConfirmationEmail(params, baseRef, salesmanName);
    }

    lock.releaseLock();
    return jsonResponse({ status: "success", ref: baseRef, date: dateStr, time: timeStr, duration: duration, slots: slotsToBook });
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
    const newDuration = parseInt(params.duration) || 15;
    const newSlots = params.slots ? params.slots.split(',') : [newTime];

    let targetExhibitionId = null;
    let targetSalesmanId = null;
    let salesmanName = null;

    // Find all existing rows for this booking (ref and ref-P*)
    const existingIndices = [];
    for (let i = 1; i < data.length; i++) {
      const rowRef = String(data[i][1]);
      if (rowRef === ref || rowRef.indexOf(ref + '-P') === 0) {
        existingIndices.push(i);
        if (!targetExhibitionId) {
          targetExhibitionId = data[i][2];
          targetSalesmanId = data[i][3];
          salesmanName = data[i][4] || getSalesmanName(targetSalesmanId);
        }
      }
    }

    if (existingIndices.length === 0) {
      lock.releaseLock();
      return jsonResponse({ status: "error", message: "Booking reference not found" });
    }

    // Make sure the new slots aren't already taken by another booking
    for (let i = 1; i < data.length; i++) {
      if (existingIndices.indexOf(i) !== -1) continue;
      const row = data[i];
      const rowStatus = row[12] || 'confirmed';
      const rowDate = formatDateStr(row[5]);
      const rowTime = row[6];
      if (rowStatus !== 'cancelled' && row[2] === targetExhibitionId && row[3] === targetSalesmanId && rowDate === newDate && newSlots.indexOf(rowTime) !== -1) {
        lock.releaseLock();
        return jsonResponse({ status: "error", message: "One or more requested slots are already booked. Please choose another time." });
      }
    }

    // Delete existing rows backwards
    for (let k = existingIndices.length - 1; k >= 0; k--) {
      sheet.deleteRow(existingIndices[k] + 1);
    }

    // Re-append updated slots
    const timestamp = new Date().toISOString();
    newSlots.forEach((slotTime, idx) => {
      const partRef = idx === 0 ? ref : (ref + '-P' + (idx + 1));
      sheet.appendRow([
        timestamp,
        partRef,
        targetExhibitionId,
        targetSalesmanId,
        salesmanName,
        "'" + newDate,
        "'" + slotTime,
        params.firstName,
        params.lastName,
        params.company,
        params.email,
        params.phone ? ("'" + formatPhoneStr(params.phone)) : "",
        "confirmed"
      ]);
    });

    // Send updated meeting email notification to client & sales representative (skip if notify=false or silent=true)
    if (!params || (params.notify !== 'false' && params.silent !== 'true' && params.noEmail !== 'true')) {
      const updateParams = Object.assign({}, params || {}, {
        exhibitionId: (params && params.exhibitionId) || targetExhibitionId,
        date: (params && params.date) || newDate,
        time: (params && params.time) || newTime
      });
      sendUpdateEmail(updateParams, ref, salesmanName, targetSalesmanId);
    }

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
    let cancelInfo = null;
    // Loop backwards when deleting rows to not mess up the row indexes
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowRef = String(row[1]);
      if (ref && (rowRef === ref || rowRef.indexOf(ref + '-P') === 0)) {
        if (!cancelInfo) {
          cancelInfo = {
            clientEmail: row[10],
            exhibitionId: row[2],
            salesmanId: row[3],
            salesmanName: row[4],
            salesmanEmail: getSalesmanEmail(row[3]),
            firstName: row[7],
            lastName: row[8],
            company: row[9],
            date: formatDateStr(row[5]),
            time: row[6]
          };
        }
        sheet.deleteRow(i + 1); // Permanently deletes row from Google Sheets
        found = true;
      }
    }

    lock.releaseLock();
    if (found && cancelInfo) {
      // Send single cancellation email notification to client & sales representative (skip if notify=false or silent=true)
      if (!params || (params.notify !== 'false' && params.silent !== 'true' && params.noEmail !== 'true')) {
        const cancelParams = Object.assign({}, params || {}, {
          exhibitionId: (params && params.exhibitionId) || cancelInfo.exhibitionId,
          date: (params && params.date) || cancelInfo.date,
          time: (params && params.time) || cancelInfo.time
        });
        sendCancellationEmail(
          cancelInfo.clientEmail, cancelInfo.salesmanEmail, ref,
          cancelInfo.firstName, cancelInfo.lastName, cancelInfo.company,
          cancelInfo.date, cancelInfo.time, cancelInfo.salesmanName, cancelParams
        );
      }
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

function formatPhoneStr(p) {
  if (!p) return '';
  const str = String(p).trim();
  if (str.startsWith("'")) return str.substring(1).trim();
  if (str.includes('#ERROR') || str.includes('#REF') || str.includes('#VALUE') || str.toLowerCase().includes('error')) return '';
  const digitsOnly = str.replace(/[\s\-\(\)\+]/g, '');
  if (digitsOnly === '66812345678' || digitsOnly === '812345678') return '';
  return str;
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

function formatTime12h(timeStr) {
  if (!timeStr) return '';
  const str = String(timeStr).trim();
  if (str.toUpperCase().includes('AM') || str.toUpperCase().includes('PM')) return str;
  const parts = str.split(':');
  if (parts.length < 2) return str;
  let h = parseInt(parts[0], 10);
  const m = parts[1].padStart(2, '0');
  if (isNaN(h)) return str;
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12;
  h = h ? h : 12;
  return `${h}:${m} ${ampm}`;
}

function calculateMeetingEndTime(timeStr, durMins) {
  if (!timeStr) return '';
  const parts = String(timeStr).split(':');
  if (parts.length < 2) return timeStr;
  let h = parseInt(parts[0], 10);
  let m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) return timeStr;
  const totalMins = h * 60 + m + durMins;
  const endH = Math.floor(totalMins / 60) % 24;
  const endM = totalMins % 60;
  return `${String(endH).padStart(2, '0')}:${String(endM).padStart(2, '0')}`;
}

function resolveExhibitionMetadata(params) {
  params = params || {};
  const rawExId = params.exhibitionId || params.e || '';
  const exId = String(rawExId).trim().toLowerCase();
  const dateStr = String(params.date || '').trim();

  // Known exhibitions catalog for fallback inference
  const KNOWN_EXHIBITIONS = {
    'ifeat-2026': {
      title: 'IFEAT 2026 Bangkok',
      venue: "Marriott Marquis Queen's Park, Bangkok, Thailand",
      location: 'Universal Oleoresins Suite / Meeting Room',
      tzAbbr: 'ICT'
    },
    'fi-india-2026': {
      title: 'Fi India 2026',
      venue: '(BEC), Goregaon, Mumbai',
      location: 'Stall 3D38, Hall 3',
      tzAbbr: 'IST'
    },
    'gulfood-2027': {
      title: 'Gulfood Dubai 2027',
      venue: 'Dubai World Trade Centre, UAE',
      location: 'Hall 4, Stand S4-C12',
      tzAbbr: 'GST'
    }
  };

  // 1. Identify preset match by ID or by Date range
  let preset = KNOWN_EXHIBITIONS[exId] || null;
  if (!preset && dateStr) {
    if (dateStr >= '2026-10-01' && dateStr <= '2026-10-31') {
      preset = KNOWN_EXHIBITIONS['ifeat-2026'];
    } else if (dateStr >= '2026-08-01' && dateStr <= '2026-08-31') {
      preset = KNOWN_EXHIBITIONS['fi-india-2026'];
    } else if (dateStr >= '2027-02-01' && dateStr <= '2027-02-28') {
      preset = KNOWN_EXHIBITIONS['gulfood-2027'];
    }
  }

  // 2. Resolve Title
  let title = (params.exhibitionTitle && String(params.exhibitionTitle).trim()) || '';
  if (!title && preset) title = preset.title;
  if (!title && rawExId && rawExId !== 'default-event') {
    title = String(rawExId).replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
  }
  if (!title) title = 'Universal Oleoresins Exhibition';

  // 3. Resolve Venue & Location
  let venue = (params.exhibitionVenue && String(params.exhibitionVenue).trim()) || '';
  if (!venue && preset) venue = preset.venue;

  let location = (params.exhibitionLocation && String(params.exhibitionLocation).trim()) || '';
  if (!location && preset) location = preset.location;
  if (!location) location = 'Universal Oleoresins Suite / Meeting Room';

  // Format combined location: "Room / Stall (Venue)"
  let fullLocation = location;
  if (venue && !location.toLowerCase().includes(venue.toLowerCase())) {
    fullLocation = `${location} (${venue})`;
  }

  // 4. Resolve Timezone Abbreviation
  let tzAbbr = (params.venueTzAbbr && String(params.venueTzAbbr).trim()) || '';
  if (!tzAbbr && preset) tzAbbr = preset.tzAbbr;
  if (!tzAbbr) {
    if (title.toLowerCase().includes('ifeat') || title.toLowerCase().includes('bangkok') || (dateStr >= '2026-10-01' && dateStr <= '2026-10-31')) {
      tzAbbr = 'ICT';
    } else if (title.toLowerCase().includes('fi india') || title.toLowerCase().includes('mumbai') || (dateStr >= '2026-08-01' && dateStr <= '2026-08-31')) {
      tzAbbr = 'IST';
    } else if (title.toLowerCase().includes('gulfood') || title.toLowerCase().includes('dubai') || (dateStr >= '2027-02-01' && dateStr <= '2027-02-28')) {
      tzAbbr = 'GST';
    } else {
      tzAbbr = 'Venue Time';
    }
  }

  return { title, venue, location: fullLocation, tzAbbr };
}

function sendConfirmationEmail(params, ref, salesmanName) {
  if (params && (params.notify === 'false' || params.silent === 'true' || params.noEmail === 'true')) return;
  const clientEmail = (params.email || '').trim();
  const salesmanEmail = params.salesmanEmail || getSalesmanEmail(params.salesmanId || params.s);
  
  if (!clientEmail || !clientEmail.includes('@') || 
      clientEmail.endsWith('@example.com') || 
      clientEmail.endsWith('@test.com') || 
      clientEmail.endsWith('@sample.com') ||
      clientEmail.endsWith('@localhost')) return;

  const duration = parseInt(params.duration) || 15;
  const lotCount = Math.max(1, Math.round(duration / 15));
  const meta = resolveExhibitionMetadata(params);
  const exTitle = meta.title;
  const exLocation = meta.location;
  const tzAbbr = meta.tzAbbr;

  // Format venue time span
  const startTime12 = params.time12 || formatTime12h(params.time);
  const endTime24 = params.endTime || calculateMeetingEndTime(params.time, duration);
  const endTime12 = params.endTime12 || formatTime12h(endTime24);
  const venueSpan = params.timeSpan || (duration > 15 ? `${startTime12} – ${endTime12}` : `${startTime12} – ${endTime12}`);
  const venueTimeDesc = `${venueSpan} ${tzAbbr}`;

  // Client local time if provided
  const hasUserTime = !!(params.userTimeSpan && params.userTzAbbr);
  const userTimeDesc = hasUserTime ? `${params.userDate ? params.userDate + ' at ' : ''}${params.userTimeSpan} ${params.userTzAbbr}` : '';

  const subject = `Meeting Confirmation: Universal Oleoresins [Ref: ${ref}]`;
  
  let bodyText = `Dear ${params.firstName} ${params.lastName},\n\n` +
    `Your meeting with ${salesmanName} at Universal Oleoresins is confirmed!\n\n` +
    `Meeting Details:\n` +
    `- Reference Code: ${ref}\n` +
    `- Event: ${exTitle}\n` +
    `- Date: ${params.date}\n` +
    `- Venue Time: ${venueTimeDesc}\n`;
  if (hasUserTime) {
    bodyText += `- Your Local Time: ${userTimeDesc}\n`;
  }
  bodyText += `- Duration: ${duration} Minutes (${lotCount} x 15-minute lots)\n` +
    `- Company: ${params.company}\n` +
    `- Sales Representative: ${salesmanName}\n` +
    `- Location: ${exLocation}\n\n` +
    `We look forward to meeting you at ${exTitle}!\n\n` +
    `Thank you,\nUniversal Oleoresins Team`;

  let userTimeHtml = '';
  if (hasUserTime) {
    userTimeHtml = `<tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold; color: #777;">Your Local Time:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #555;">${userTimeDesc}</td></tr>`;
  }

  const bodyHtml = `
    <div style="font-family: Arial, sans-serif; color: #1a1410; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden;">
      <div style="background-color: #d54e1f; padding: 22px 20px; text-align: center; color: white;">
        <h2 style="margin: 0; font-size: 22px; letter-spacing: 0.5px;">UNIVERSAL OLEORESINS</h2>
        <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.92; font-weight: 500;">Meeting Confirmation · ${exTitle}</p>
      </div>
      <div style="padding: 24px; background-color: #ffffff;">
        <p style="font-size: 16px; margin-top: 0;">Dear <strong>${params.firstName} ${params.lastName}</strong>,</p>
        <p style="font-size: 14px; color: #4a3f33; line-height: 1.5;">Your meeting has been successfully confirmed. Here are your reservation details:</p>
        
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fbf7f0; border-radius: 8px; overflow: hidden;">
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold; width: 40%;">Reference Code:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${ref}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Event:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">${exTitle}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Date & Time (Venue):</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${params.date} at ${venueTimeDesc}</td></tr>
          ${userTimeHtml}
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Duration:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${duration} Minutes (${lotCount} x 15-min lots)</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Representative:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${salesmanName}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Company:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${params.company}</td></tr>
          <tr><td style="padding: 10px 14px; font-weight: bold;">Location:</td><td style="padding: 10px 14px;">${exLocation}</td></tr>
        </table>

        <p style="font-size: 13px; color: #666; margin-top: 20px;">We look forward to meeting you at ${exTitle}!</p>
      </div>
    </div>
  `;

  // 1. Send confirmation to client
  try {
    MailApp.sendEmail({ to: clientEmail, subject: subject, body: bodyText, htmlBody: bodyHtml });
  } catch(e1) {
    try {
      GmailApp.sendEmail(clientEmail, subject, bodyText, { htmlBody: bodyHtml });
    } catch(e2) {
      console.warn("Client email send error:", e2);
    }
  }

  // 2. Send notification to salesman
  if (salesmanEmail && salesmanEmail !== clientEmail) {
    try {
      MailApp.sendEmail({
        to: salesmanEmail,
        subject: `New Meeting Booked (${duration} Mins): ${params.firstName} ${params.lastName} (${params.company}) [Ref: ${ref}]`,
        body: bodyText,
        htmlBody: bodyHtml
      });
    } catch(e3) {
      try {
        GmailApp.sendEmail(salesmanEmail, `New Meeting Booked: ${params.firstName} ${params.lastName}`, bodyText, { htmlBody: bodyHtml });
      } catch(e4) {
        console.warn("Salesman email send error:", e4);
      }
    }
  }
}

function sendUpdateEmail(params, ref, salesmanName, salesmanId) {
  if (params && (params.notify === 'false' || params.silent === 'true' || params.noEmail === 'true')) return;
  const clientEmail = (params.email || '').trim();
  const salesmanEmail = params.salesmanEmail || getSalesmanEmail(salesmanId || params.salesmanId || params.s);
  
  if (!clientEmail || !clientEmail.includes('@') || 
      clientEmail.endsWith('@example.com') || 
      clientEmail.endsWith('@test.com') || 
      clientEmail.endsWith('@sample.com') ||
      clientEmail.endsWith('@localhost')) return;

  const duration = parseInt(params.duration) || 15;
  const lotCount = Math.max(1, Math.round(duration / 15));
  const meta = resolveExhibitionMetadata(params);
  const exTitle = meta.title;
  const exLocation = meta.location;
  const tzAbbr = meta.tzAbbr;

  // Format venue time span
  const startTime12 = params.time12 || formatTime12h(params.time);
  const endTime24 = params.endTime || calculateMeetingEndTime(params.time, duration);
  const endTime12 = params.endTime12 || formatTime12h(endTime24);
  const venueSpan = params.timeSpan || (duration > 15 ? `${startTime12} – ${endTime12}` : `${startTime12} – ${endTime12}`);
  const venueTimeDesc = `${venueSpan} ${tzAbbr}`;

  // Client local time if provided
  const hasUserTime = !!(params.userTimeSpan && params.userTzAbbr);
  const userTimeDesc = hasUserTime ? `${params.userDate ? params.userDate + ' at ' : ''}${params.userTimeSpan} ${params.userTzAbbr}` : '';

  const subject = `Updated Meeting Schedule: Universal Oleoresins [Ref: ${ref}]`;
  
  let bodyText = `Dear ${params.firstName} ${params.lastName},\n\n` +
    `Your meeting with ${salesmanName} at Universal Oleoresins has been updated/rescheduled.\n\n` +
    `Updated Meeting Details:\n` +
    `- Reference Code: ${ref}\n` +
    `- Event: ${exTitle}\n` +
    `- New Date: ${params.date}\n` +
    `- New Venue Time: ${venueTimeDesc}\n`;
  if (hasUserTime) {
    bodyText += `- Your Local Time: ${userTimeDesc}\n`;
  }
  bodyText += `- Duration: ${duration} Minutes (${lotCount} x 15-minute lots)\n` +
    `- Company: ${params.company}\n` +
    `- Representative: ${salesmanName}\n` +
    `- Location: ${exLocation}\n\n` +
    `We look forward to meeting you at ${exTitle}!\n\n` +
    `Thank you,\nUniversal Oleoresins Team`;

  let userTimeHtml = '';
  if (hasUserTime) {
    userTimeHtml = `<tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold; color: #777;">Your Local Time:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #555;">${userTimeDesc}</td></tr>`;
  }

  const bodyHtml = `
    <div style="font-family: Arial, sans-serif; color: #1a1410; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden;">
      <div style="background-color: #e8852b; padding: 22px 20px; text-align: center; color: white;">
        <h2 style="margin: 0; font-size: 22px; letter-spacing: 0.5px;">UNIVERSAL OLEORESINS</h2>
        <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.92; font-weight: 500;">Meeting Updated / Rescheduled · ${exTitle}</p>
      </div>
      <div style="padding: 24px; background-color: #ffffff;">
        <p style="font-size: 16px; margin-top: 0;">Dear <strong>${params.firstName} ${params.lastName}</strong>,</p>
        <p style="font-size: 14px; color: #4a3f33; line-height: 1.5;">Your meeting details have been updated. Here is your revised schedule:</p>
        
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fbf7f0; border-radius: 8px; overflow: hidden;">
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold; width: 40%;">Reference Code:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${ref}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Event:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">${exTitle}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">New Date & Time (Venue):</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${params.date} at ${venueTimeDesc}</td></tr>
          ${userTimeHtml}
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Duration:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${duration} Minutes (${lotCount} x 15-min lots)</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Representative:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${salesmanName}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Company:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${params.company}</td></tr>
          <tr><td style="padding: 10px 14px; font-weight: bold;">Location:</td><td style="padding: 10px 14px;">${exLocation}</td></tr>
        </table>

        <p style="font-size: 13px; color: #666; margin-top: 20px;">We look forward to meeting you at ${exTitle}!</p>
      </div>
    </div>
  `;

  try { MailApp.sendEmail({ to: clientEmail, subject: subject, body: bodyText, htmlBody: bodyHtml }); } catch(e) { try { GmailApp.sendEmail(clientEmail, subject, bodyText, { htmlBody: bodyHtml }); } catch(err) {} }
  if (salesmanEmail && salesmanEmail !== clientEmail) {
    try { MailApp.sendEmail({ to: salesmanEmail, subject: `Meeting Updated (${duration} Mins): ${params.firstName} ${params.lastName} [Ref: ${ref}]`, body: bodyText, htmlBody: bodyHtml }); } catch(e) { try { GmailApp.sendEmail(salesmanEmail, `Meeting Updated: ${params.firstName} ${params.lastName}`, bodyText, { htmlBody: bodyHtml }); } catch(err) {} }
  }
}

function sendCancellationEmail(clientEmail, salesmanEmail, ref, firstName, lastName, company, date, time, salesmanName, params) {
  if (params && (params.notify === 'false' || params.silent === 'true' || params.noEmail === 'true')) return;
  clientEmail = (clientEmail || '').trim();
  if (!clientEmail || !clientEmail.includes('@') || 
      clientEmail.endsWith('@example.com') || 
      clientEmail.endsWith('@test.com') || 
      clientEmail.endsWith('@sample.com') ||
      clientEmail.endsWith('@localhost')) return;

  const meta = resolveExhibitionMetadata(params || { date: date });
  const exTitle = meta.title;
  const exLocation = meta.location;
  const tzAbbr = meta.tzAbbr;
  const timeFormatted = formatTime12h(time) + (tzAbbr ? ` ${tzAbbr}` : '');

  const subject = `Meeting Cancelled: Universal Oleoresins [Ref: ${ref}]`;
  const bodyText = `Dear ${firstName} ${lastName},\n\n` +
    `Your meeting scheduled for ${date} at ${timeFormatted} with ${salesmanName} at Universal Oleoresins (${exTitle}) has been cancelled.\n\n` +
    `Reference Code: ${ref}\n` +
    `Event: ${exTitle}\n` +
    `Location: ${exLocation}\n` +
    `Company: ${company}\n\n` +
    `All reserved slots for this meeting have been released. If you wish to reschedule, please visit our scheduler website.\n\n` +
    `Thank you,\nUniversal Oleoresins Team`;

  const bodyHtml = `
    <div style="font-family: Arial, sans-serif; color: #1a1410; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden;">
      <div style="background-color: #555555; padding: 22px 20px; text-align: center; color: white;">
        <h2 style="margin: 0; font-size: 22px; letter-spacing: 0.5px;">UNIVERSAL OLEORESINS</h2>
        <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.92; font-weight: 500;">Meeting Cancellation Notice · ${exTitle}</p>
      </div>
      <div style="padding: 24px; background-color: #ffffff;">
        <p style="font-size: 16px; margin-top: 0;">Dear <strong>${firstName} ${lastName}</strong>,</p>
        <p style="font-size: 14px; color: #4a3f33; line-height: 1.5;">This is to confirm that your meeting scheduled for <strong>${date} at ${timeFormatted}</strong> with ${salesmanName} (${company}) at <strong>${exTitle}</strong> has been cancelled and all reserved slots have been released.</p>
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fbf7f0; border-radius: 8px; overflow: hidden;">
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold; width: 40%;">Reference Code:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${ref}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Event:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">${exTitle}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Cancelled Schedule:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee; color: #d54e1f; font-weight: bold;">${date} at ${timeFormatted}</td></tr>
          <tr><td style="padding: 10px 14px; border-bottom: 1px solid #eee; font-weight: bold;">Location:</td><td style="padding: 10px 14px; border-bottom: 1px solid #eee;">${exLocation}</td></tr>
          <tr><td style="padding: 10px 14px; font-weight: bold;">Representative:</td><td style="padding: 10px 14px;">${salesmanName}</td></tr>
        </table>
        <p style="font-size: 13px; color: #666; margin-top: 20px;">If you would like to pick a different date or time, please feel free to book a new slot on our online scheduler.</p>
      </div>
    </div>
  `;

  try { MailApp.sendEmail({ to: clientEmail, subject: subject, body: bodyText, htmlBody: bodyHtml }); } catch(e) { try { GmailApp.sendEmail(clientEmail, subject, bodyText, { htmlBody: bodyHtml }); } catch(err) {} }
  if (salesmanEmail && salesmanEmail !== clientEmail) {
    try { MailApp.sendEmail({ to: salesmanEmail, subject: `Meeting Cancelled: ${firstName} ${lastName} (${company}) [Ref: ${ref}]`, body: bodyText, htmlBody: bodyHtml }); } catch(e) { try { GmailApp.sendEmail(salesmanEmail, `Meeting Cancelled: ${firstName} ${lastName}`, bodyText, { htmlBody: bodyHtml }); } catch(err) {} }
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    resolveExhibitionMetadata,
    formatTime12h,
    calculateMeetingEndTime,
    sendConfirmationEmail,
    sendUpdateEmail,
    sendCancellationEmail,
    formatPhoneStr,
    formatDateStr,
    getSalesmanName,
    getSalesmanEmail
  };
}

/**
 * Helper function: Run this ONCE inside Apps Script editor to authorize email permissions!
 */
function testEmailPermissions() {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  if (email) {
    MailApp.sendEmail(email, "Universal Oleoresins Email Test", "Email authorization test successful!");
    Logger.log("Test email sent to " + email);
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}