// Copy this file into the Apps Script project attached to the directory sheet.
// Updating this file in GitHub alone does not update the deployed web app.
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
    if (!sheet) throw new Error('Submissions sheet was not found');

    // The script runs as the sheet owner, but Session.getEffectiveUser() needs
    // an additional OAuth scope. Use the verified notification address here.
    const ownerEmail = 'divinedivas333@gmail.com';

    if (data.action === 'delete') {
      const email = String(data.email || '').trim();
      if (!email) return reply_('missing_email');

      MailApp.sendEmail({
        to: ownerEmail,
        subject: '⚠️ Divine Purple Pages — Listing Removal Request',
        body: 'A member has requested their listing be removed from the Divine Purple Pages Community Directory.\n\n' +
          'Name: ' + (data.name || 'Not provided') + '\nEmail: ' + email + '\n\n' +
          'Find the matching email in the Submissions sheet and remove that row. ' +
          'The published directory will update after the sheet refreshes.'
      });
      return reply_('removal_requested');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    let notice = null;
    let status;
    try {
      const rows = sheet.getDataRange().getValues();
      const email = String(data.email || '').trim().toLowerCase();

      if (data.action === 'update') {
        if (!email) return reply_('missing_email');
        const index = rows.findIndex((row, i) => i > 0 &&
          String(row[3] || '').trim().toLowerCase() === email);
        if (index < 0) return reply_('not_found');

        const newEmail = String(data.newEmail || '').trim();
        if (newEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(newEmail)) {
          return reply_('invalid_email');
        }
        if (newEmail && rows.some((row, i) => i > 0 && i !== index &&
            String(row[3] || '').trim().toLowerCase() === newEmail.toLowerCase())) {
          return reply_('duplicate_email');
        }
        const websiteIssue = websiteIssue_(data.newWebsites, rows);
        if (websiteIssue) return reply_(websiteIssue);

        const rowNumber = index + 1;
        if (data.newName) sheet.getRange(rowNumber, 2).setValue(String(data.newName).trim());
        if (newEmail) sheet.getRange(rowNumber, 4).setValue(newEmail);
        if (data.newOffer) sheet.getRange(rowNumber, 5).setValue(String(data.newOffer).trim());
        if (Array.isArray(data.newWebsites) && data.newWebsites.length) {
          sheet.getRange(rowNumber, 6).setValue(mergeLinks_(rows[index][5], data.newWebsites, 'url'));
        }
        if (Array.isArray(data.newSocials) && data.newSocials.length) {
          sheet.getRange(rowNumber, 7).setValue(mergeLinks_(rows[index][6], data.newSocials, 'handle'));
        }

        notice = {
          subject: '✦ Divine Purple Pages — Listing Updated',
          body: 'A member submitted a listing update.\n\nOriginal email: ' + email +
            (newEmail ? '\nNew email: ' + newEmail : '') +
            (data.newName ? '\nNew name: ' + data.newName : '') +
            (data.newOffer ? '\nNew offer: ' + data.newOffer : '') +
            (data.newWebsites && data.newWebsites.length ? '\nWebsites added: ' + JSON.stringify(data.newWebsites) : '') +
            (data.newSocials && data.newSocials.length ? '\nSocials added: ' + JSON.stringify(data.newSocials) : '') +
            (data.removeItems ? '\nRequested removals (manual action needed): ' + data.removeItems : '')
        };
        status = 'updated';
      } else {
        if (!email) return reply_('missing_email');
        if (rows.some((row, i) => i > 0 && String(row[3] || '').trim().toLowerCase() === email)) {
          return reply_('duplicate_email');
        }
        const websiteIssue = websiteIssue_(data.websites, rows);
        if (websiteIssue) return reply_(websiteIssue);
        sheet.appendRow([
          new Date(),
          data.name || '',
          data.nickname || '',
          data.email || '',
          data.whatYouOffer || '',
          JSON.stringify(data.websites || []),
          JSON.stringify(data.socials || []),
          data.submittedAt || ''
        ]);
        status = 'success';
      }
    } finally {
      lock.releaseLock();
    }

    // A mail failure must not cause a member to retry an update that already saved.
    if (notice) {
      try {
        MailApp.sendEmail({ to: ownerEmail, subject: notice.subject, body: notice.body });
      } catch (err) {
        console.error('Listing saved; owner notification failed: ' + err);
        status = 'updated_notice_failed';
      }
    }
    return reply_(status);
  } catch (err) {
    console.error('Directory submission failed: ' + err);
    return reply_('error');
  }
}

// Read-only website check used by the public forms. Return only a status, never
// sheet rows or member information. The callback is restricted to a safe name.
function doGet(e) {
  const callback = String((e.parameter || {}).callback || '');
  if (!/^[$A-Za-z_][$\w]{0,100}$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback');
  }
  let status = 'unavailable';
  try {
    const urls = JSON.parse(e.parameter.websites || '[]');
    if (Array.isArray(urls) && urls.length <= 20) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
      const rows = sheet.getDataRange().getValues();
      status = websiteIssue_(urls.map(url => ({ url })), rows) || 'available';
    }
  } catch (err) {
    console.error('Website lookup failed: ' + err);
  }
  return ContentService.createTextOutput(callback + '(' + JSON.stringify({ status }) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Check the live sheet inside the submission lock. The published CSV can lag
// behind new rows, so browser validation alone cannot prevent a duplicate.
function websiteIssue_(additions, rows) {
  if (!Array.isArray(additions) || !additions.length) return null;
  const keys = new Set();
  for (const website of additions) {
    const key = websiteKey_(website && website.url);
    if (!key) return 'invalid_website';
    if (keys.has(key)) return 'duplicate_website';
    keys.add(key);
  }
  for (let i = 1; i < rows.length; i++) {
    let existing;
    try { existing = JSON.parse(rows[i][5] || '[]'); }
    catch (err) { continue; }
    if (!Array.isArray(existing)) continue;
    for (const website of existing) {
      const key = websiteKey_(website && website.url);
      if (key && keys.has(key)) return 'duplicate_website';
    }
  }
  return null;
}

function websiteKey_(value) {
  const input = String(value || '').trim()
    .replace(/^[a-z][a-z\d+.-]*:\/\//i, '').replace(/^\/\//, '');
  const match = input.match(/^([^/?#]+)([^?#]*)/);
  if (!match || /\s|@/.test(match[1])) return '';
  const host = match[1].toLowerCase().replace(/:\d+$/, '')
    .replace(/^www\./, '').replace(/\.$/, '');
  if (!/^[a-z\d.-]+\.[a-z\d-]+$/.test(host)) return '';
  const path = match[2].replace(/\/+$/, '').toLowerCase();
  const sharedHosts = [
    'etsy.com', 'amazon.com', 'stan.store', 'beacons.ai', 'gumroad.com',
    'payhip.com', 'skool.com', 'shopify.com', 'facebook.com',
    'instagram.com', 'tiktok.com', 'youtube.com', 'linktr.ee'
  ];
  return sharedHosts.indexOf(host) >= 0 ? host + path : host;
}

function mergeLinks_(existingValue, additions, addressKey) {
  let existing;
  try { existing = JSON.parse(existingValue || '[]'); } catch (err) { existing = []; }
  if (!Array.isArray(existing)) existing = [];

  const seen = new Set(existing.map(item =>
    String(item.type || '').trim().toLowerCase() + '|' +
    String(item[addressKey] || '').trim().toLowerCase()));
  additions.forEach(item => {
    const type = String(item.type || '').trim();
    const address = String(item[addressKey] || '').trim();
    if (!type || !address) return;
    const key = type.toLowerCase() + '|' + address.toLowerCase();
    if (!seen.has(key)) {
      existing.push({ type: type, [addressKey]: address });
      seen.add(key);
    }
  });
  return JSON.stringify(existing);
}

function reply_(status) {
  return ContentService.createTextOutput(JSON.stringify({ status: status }))
    .setMimeType(ContentService.MimeType.JSON);
}
