# Changelog

**English** | [한국어](CHANGELOG.ko.md)

## 2026 — Rewrite

### Upgrading from 2023

1. In the Apps Script editor, replace `Code.gs` and `MailMerge.html`, and add `i18n.gs`, `Lang_en.gs`, `Lang_ko.gs`, `Manual_en.html`, `Manual_ko.html`.
2. Delete `UserManual.html`, `createDraft.html` and `createSlide.html` (no longer used).
3. Replace `appsscript.json` (or add **Gmail API** and **Google Slides API** under Services).
4. Reload the spreadsheet and send a test email.

Existing sheets keep working: `EMAIL_SENT` statuses and `{{Name}}`-style fields are compatible.

### Fixed

- **Emoji in the subject were garbled** (`🎄` → `��`). `GmailApp.sendEmail` mis-encodes 4-byte characters, while `MailApp` ignores the `from` alias. Messages are now built as UTF-8 MIME (RFC 2047 headers) and sent via the Gmail API, so both emoji and aliases work.
- **Status written to the wrong row** when the sheet had blank rows (EMAIL_SENT could be marked on someone who wasn't emailed).
- **Fields shifted by one column** when a header was blank or non-ASCII (e.g. Korean), putting wrong values into emails.
- **Wrong draft could be sent**: drafts were listed with `in:draft` but picked with `in:drafts` by index; reply drafts used the first message of the thread. Drafts are now selected by ID.
- One failing row stopped the whole run, and the error dialog itself failed (`showModalDialog(e)`). Failures are now recorded per row and the run continues.
- Long lists hit the 6-minute limit without any record of progress. The run now stops safely at ~5 minutes and resumes on the next RUN.
- The slide-image feature duplicated a slide per recipient and never deleted it.
- Dates and numbers were inserted as raw JS values; `0` became blank. Values now appear as displayed in the sheet.
- Values containing `&`, `<` broke the HTML; values containing `$` were mangled. Values are now escaped and inserted literally.
- Merge fields split by Gmail's editor (`{{<span>Name</span>}}`) were not replaced.
- A missing `Email` header prompted for a column and overwrote the existing header.
- A `Mail Merge Status` column that wasn't the last column was duplicated and statuses were written to the last column.
- Double-clicking RUN could send twice.
- "Clear sheet" failed on a sheet with only a header row.
- Draft subjects were truncated in the middle of an emoji.

### Added

- English / Korean UI (`i18n.gs`, `Lang_*.gs`), auto-detected from the Google account language, switchable in the sidebar.
- Pre-send confirmation showing the number of recipients.
- `Mail Merge Sent At` column with the time of the last successful send (created next to the status column; never overwrites existing data). Test emails no longer modify the sheet.
- The `EMAIL_SENT` check ignores case, spaces, `_` and `-` (`Email_Sent`, `email sent` are skipped too).
- CC field; CC/BCC address validation.
- Sender name auto-filled from the Gmail "Send mail as" settings and remembered per user.
- Warning for merge fields that don't exist in the sheet.
- Plain-text alternative body.
- Automatic retry for transient Gmail API errors.
- Clear message with an enable link when the Gmail or Slides API is disabled in the Cloud project.

### Removed

- `getMyName()`, which created a `Mail Merge.txt` file in Drive just to read the owner's name.
- `getResult()`, `rangeToObjects()`, `onInstall()` and the old row-reading helpers (replaced by `readTable_()`).
- Four separate sidebar calls (replaced by `getSidebarData()`), redundant JSON round trips, jQuery, commented-out code.
- `createDraft.html` and `createSlide.html` (now handled inside the sidebar).
- The "emails remaining today" toast (it showed the MailApp quota, which doesn't match Gmail sending).

## 2023-11-15

- Switched from `MailApp.sendEmail` to `GmailApp.sendEmail` so the selected From alias is used.
- Added Reply-To set to the sender address.

## 2019

- Initial release: Gmail draft templates, inline images, attachments, test email, BCC, slide image.
