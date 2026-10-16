# GitHub Pages + Google Sheets setup (no separate paid host)

The public site is in /docs and can be published using GitHub Pages. GitHub Pages does not run the older Node.js backend; that version remains in /public and server.js only as a separate alternative.

## 1. Create the private registration spreadsheet

1. Open https://sheets.new in **your Google account**, name it e.g. DIIT 21st CSE Registrations. **Do not publicly share this sheet**.
2. From the sheet, open Extensions → Apps Script. Replace Code.gs with the contents of `apps-script/Code.gs` from this repository.
3. Create a new HTML file called `Registration` and paste the contents of `apps-script/Registration.html` into it.
4. From the Apps Script editor's function picker, select `setup_` and click Run. Authorize access to your sheet and email. This creates the Registrations tab, header row and installable payment approval trigger.
5. Deploy → New deployment → Web app. Select **Execute as: Me** and **Who has access: Anyone**. Deploy/authorize and copy the URL ending in `/exec` (not /dev). Public access is needed for students to submit; **your spreadsheet itself must stay private**.
6. Edit `docs/config.js` on GitHub: paste that exact Apps Script URL between the quotes. Commit it. The landing page will embed the registration form.

## 2. Enable GitHub Pages

In GitHub → repository Settings → Pages, select **Deploy from a branch**, Branch **main**, folder **/docs**, then Save. Your page URL will normally be:
https://ar210038.github.io/diit-21-cse-closing-ceremony/

GitHub Free generally requires the repository to be **public** for Pages. If Pages is unavailable while private, change the **repository visibility to Public** under Settings → General → Danger Zone; keep private student information only in the Google Sheet. Do **not** make the sheet public. On eligible paid GitHub plans, Pages can be published from a private repository, but the site is still publicly viewable.

## 3. Take payments and confirm them

- The public form calculates ৳1,000 + ৳500 per guest. Students choose Cash or bKash and may enter the bKash transaction ID.
- Incoming registrations appear in the private **Registrations** Google Sheet as PENDING.
- **After actually receiving/verifying payment**, tick the checkbox in column K (Payment received). The installable on-edit trigger marks that row CONFIRMED and sends a confirmation email with the QR ticket image and verification link.
- If the email fails, column P (Email error) shows why. Uncheck and recheck column K to retry the failed email. Don't repeatedly resend successful confirmations.
- The QR code opens a ticket verification page listing the student name, ID, guest count and reference; it **does not** include phone or email.

## Notes

- The Apps Script email sender is the Google account that authorized the trigger; Google imposes daily mail quotas. Check quotas before a large batch: https://developers.google.com/apps-script/guides/services/quotas
- Google may request authorization because the script edits Sheets and sends emails. Review permissions carefully; deploy only code you trust.
- This is **manual** bKash verification, not a bKash gateway. Never treat a transaction ID as proof of payment without checking your bKash account.
- Anyone with a valid QR URL can view the ticket's name, ID and guest count. Treat the QR code as a bearer ticket and do not share it publicly.
- Register with a test student ID, verify a received payment, scan the resulting QR and verify email sending **before** opening registration. The deployment and flow have not been end-to-end tested in this environment.
