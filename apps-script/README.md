# Directory Apps Script deployment

The GitHub Pages HTML and the Google Apps Script web app are deployed separately. A GitHub update does not change the code running at the `/exec` URL.

1. In the directory Google Sheet, open **Extensions → Apps Script**.
2. Review the current `Code.gs` against this version, then replace the existing `doPost` implementation with the contents of [Code.gs](Code.gs). This version keeps columns A–H and has no photo fields.
3. Save the project. Open **Deploy → Manage deployments**, select the existing web app, and click the pencil icon. Select **New version**, keep **Execute as: Me** and **Who has access: Anyone**, and click **Deploy**. Reauthorize Sheets and Mail when prompted.
4. Confirm the web app URL still matches `FORM_ENDPOINT_URL` in `manage.html` and `submit.html`. If a new deployment URL was created instead, update both HTML files before publishing them.
5. Check **Executions** in Apps Script after a controlled listing update, and confirm that the sheet changed and the owner received the notification email. Test a removal request separately; it must send an email and should not delete the row automatically.

The site uses `mode: 'no-cors'` for the cross-origin POST, so it cannot read the Apps Script result. The site therefore says the request was *sent*, not that the update was saved. Check the Sheet or wait for the published CSV to refresh before retrying.
