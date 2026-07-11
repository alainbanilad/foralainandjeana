# Backend

Reference copies of the backend code for foralainandjeana.com. Nothing in this
folder runs from the repo — these are deployed elsewhere. Keep them in sync:
if you edit the live version, update the copy here (and vice versa).

## Files

### `apps-script.gs`
The complete Google Apps Script bound to the wedding spreadsheet.
Two independent parts:

- **Part 1 — RSVP email parser.** Runs hourly on a time-based trigger.
  Reads "New RSVP from" emails in Gmail, writes rows to the **RSVPs** tab,
  and sends a personalised reply to the guest (CC to us both).
- **Part 2 — Menu selection endpoint.** A `doPost` web app that receives
  submissions from `/menu.html` and writes rows to the **Menu Selections**
  tab (created automatically on first submission).

Deployed at: Google Sheets → Extensions → Apps Script.
Reminder: after editing, create a **new deployment version**
(Deploy → Manage deployments), otherwise the live URL serves stale code.

### `emailjs-template-menu.html`
The HTML content of the EmailJS template that sends menu confirmation
emails to guests (Bcc copies to us both). The setup comment at the top of
the file lists the Subject / To / Bcc settings. Pasted into:
emailjs.com → Email Templates.

The main RSVP form also uses EmailJS, but that template is a plain
notification email whose real logic lives in Part 1 of the Apps Script.

## No secrets here

Everything in this folder is safe for a public repo. The EmailJS public key
and Apps Script web app URL are exposed client-side in the site's HTML by
design — the Apps Script only accepts writes (it can't read data back out),
and EmailJS sends are locked to our templates.
