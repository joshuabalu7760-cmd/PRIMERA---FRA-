# Koritsu FRA Works

A local app for importing your property workbook, assigning embedded FRA PDFs to properties, reviewing likely works findings, selecting an Outlook template, and sending messages through Microsoft 365.

## Run locally

Requires Python 3.10 or newer. From this directory:

```sh
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000. Download the FRA PDFs from Odoo, then insert each PDF as an embedded file object in your `.xlsx` or `.xlsm` workbook. The app does not connect to Odoo; it reads the report copies embedded in the workbook. The first worksheet should use these headings: `Managing Agent`, `MA Email`, `Property Address`, `Postcode`, `FRA Report`, `inspection date`, `due date`, and `notes`. Inspection date, due date, and notes are optional to the app; the first five headings are required. The app lists embedded PDFs for you to assign to property rows; it preselects a report only when its extracted filename matches the `FRA Report` cell.

## Outlook templates

In the email step, import one or more Outlook `.oft` files. The imported subject and message body are stored in the local `instance/email_templates` folder and remain available across app restarts. Choose a different entry in the template selector to switch templates for a batch. The app imports the message body and subject; Outlook inline images, signatures that depend on embedded images, and other template attachments are not carried over.

Template fields can be written as `{{Managing Agent}}`, `{{MA Email}}`, `{{Property Address}}`, `{{Postcode}}`, `{{Inspection Date}}`, `{{Due Date}}`, or `{{Notes}}`. The app also recognizes `«Property Address»` and `[[Property Address]]` forms. Unknown fields are left unchanged.

## Microsoft 365 setup

Sending uses Microsoft Graph delegated `Mail.Send` permission and device-code sign-in. Your Microsoft 365 administrator may need to register an Entra application, enable public-client/device-code authentication, add the delegated `Mail.Send` permission, and grant consent. Enter the app's client ID and your tenant ID in the app, select Connect Microsoft 365, and complete the sign-in prompt. The connection control becomes Disconnect Microsoft 365 when connected. Disconnect removes the token from Koritsu's memory so the app cannot use it again; Microsoft-issued access tokens may remain valid until they expire. The app does not use Odoo's API or SMTP.

## Review and limits

PDF text is checked for common works/action phrases, but this is only a triage aid. Scanned PDFs are not OCR'd, and every report status should be reviewed before sending. Larger PDFs use Microsoft Graph attachment upload sessions. The app keeps workbook/report data and Microsoft access tokens in memory while it runs; imported templates are stored on disk. Run it only on a trusted machine and do not expose it to a public network. Test with a small batch before operational use.