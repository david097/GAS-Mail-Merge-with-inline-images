# GAS Mail Merge with inline images

**English** | [한국어](README.ko.md)

Send personalised emails to everyone in a Google Sheet, using a Gmail draft as the template — with inline images, attachments, emoji and "Send mail as" aliases. Built with Google Apps Script, no add-on install required.

## Features

- **Gmail draft as template** — write the email in Gmail and use `{{Column Name}}` merge fields in the subject and body.
- **Inline images & attachments** are kept exactly as in the draft.
- **Emoji-safe** — subject, body and sender name are sent as proper UTF-8 MIME via the Gmail API, so 🎄🎁 don't turn into `��`.
- **Send from an alias** configured in Gmail ("Send mail as"), with Reply-To set to the same address.
- **Test email** to yourself before sending to everyone.
- **Resumable** — sent rows are marked `EMAIL_SENT`; failures are recorded per row and don't stop the run. Long lists stop safely before the 6-minute limit and continue on the next run.
- **Personalised slide image** (optional) — render the first slide of a Google Slides file per recipient and embed it with `{{SlideImage}}`.
- **Multilingual UI** — English and Korean, auto-detected from the Google account language and switchable in the sidebar.

## Installation

### Option A — copy the sample spreadsheet

1. Open the [sample spreadsheet](https://bit.ly/2ZJmLwL) and choose **File → Make a copy**.
2. Open your copy and reload it. A **▼Mail Merge▼** menu appears.
3. The first time you run it, Google asks for authorisation. If you see "Google hasn't verified this app", click **Advanced → Go to (project name)**. The script runs only in your account.

### Option B — install manually

1. Create a Google Sheet and open **Extensions → Apps Script**.
2. Add the files from this repository (keep the same names):

   | File | Type |
   |---|---|
   | `Code.gs`, `i18n.gs`, `Lang_en.gs`, `Lang_ko.gs` | Script |
   | `MailMerge.html`, `Manual_en.html`, `Manual_ko.html` | HTML |

3. **Project Settings → Show "appsscript.json" manifest file in editor**, then replace its content with [`appsscript.json`](appsscript.json). This enables the **Gmail API** and **Google Slides API** advanced services.
   (Alternatively add them under **Services (+)**.)
4. Reload the spreadsheet.

## Usage

1. **Create a draft in Gmail.** Use the sheet's column headers as merge fields:

   ```
   Subject: 🎄 {{Name}}, you're invited to our year-end party
   Body:    Hi {{Name}} from {{Company}}, ...
   ```

   Matching ignores case, spaces and symbols: `{{Email Address}}` = `{{emailaddress}}`.

2. **List recipients in a sheet.** Row 1 is the header. The email column must be named `Email` or `Email Address` (`이메일` also works).

   | Company | Name | Email | Mail Merge Status | Mail Merge Sent At |
   |---|---|---|---|---|
   | Acme | Jane | jane@example.com | | |

3. **▼Mail Merge▼ → Open Mail Merge**, choose the draft, sheet and From address.
4. Choose **Send a test email** first and check it in your inbox.
5. Choose **Run mail merge** → **RUN**. The button shows how many people will receive it; click again to send.

### Results

| Status | Meaning |
|---|---|
| `EMAIL_SENT` | Sent. This row is skipped next time — clear the cell to resend. (Case-insensitive: `Email_Sent` also counts.) |
| `ERROR: …` | Failed, with the reason. Fix it and click RUN again. |

The **`Mail Merge Sent At`** column (created next to the status column) records the time of the last successful send. Test emails don't change the sheet.

### Slide image (optional)

1. In a Google Slides file, design the **first slide** with fields such as `{{Name}}`.
2. Put `{{SlideImage}}` in the email body.
3. Tick **Insert slide image** in the sidebar and pick the presentation.

## Troubleshooting

| Problem | Solution |
|---|---|
| `Gmail API service is not added` | Add **Gmail API** under Services (+), or use the provided `appsscript.json`. |
| `… API has not been used in project … or it is disabled` | Your script is linked to a standard Google Cloud project. Open the link in the message and click **Enable** (for Gmail API and, if used, Google Slides API). |
| `Email column not found` | Name the email column's header `Email`. |
| A field is blank in the sent email | The result dialog lists merge fields not found in the sheet — check the spelling against the header row. |
| Daily limit | Gmail's own sending limits apply (about 2,000/day for Google Workspace accounts, lower for personal Gmail accounts). |
| Large attachments fail | Gmail API requests have a size limit; share large files as Drive links instead. |

## Adding a language

1. Copy `Lang_en.gs` to `Lang_<code>.gs` (e.g. `Lang_ja.gs`) and change `LANGS.en` to `LANGS.ja`.
2. Translate the values. Keep `{placeholders}` and `{{merge fields}}` unchanged. Missing keys fall back to English.
3. Copy `Manual_en.html` to `Manual_<code>.html` and translate it.

The language is chosen automatically from the user's Google account language when it is supported, otherwise English.

## Project structure

```
Code.gs           Main logic: sheet reading, merge fields, MIME building, sending
i18n.gs           Translation helper t(), language detection and selection
Lang_en.gs        English strings (default / fallback)
Lang_ko.gs        Korean strings
MailMerge.html    Sidebar UI (HtmlService template)
Manual_en.html    In-app manual (English)
Manual_ko.html    In-app manual (Korean)
appsscript.json   Manifest: V8 runtime, Gmail & Slides advanced services
```

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

## Author

© David Sung (2019–2026)
