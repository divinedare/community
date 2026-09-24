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
