/**
 * English (default / fallback language).
 * To add a language: copy this file to Lang_<code>.gs (e.g. Lang_ja.gs), change "en" below
 * to the language code, translate the values, and add Manual_<code>.html.
 * Keep {placeholders} and {{merge fields}} unchanged.
 */
var LANGS = LANGS || {};

LANGS.en = {
  _name: 'English',

  // Menu & dialogs
  menu_open: 'Open Mail Merge',
  menu_clear: 'Clear current sheet',
  menu_manual: 'How to use',
  sidebar_title: 'Mail Merge',
  manual_title: 'How to use Mail Merge',
  clear_title: 'Caution',
  clear_confirm: "Delete all contents of '{sheet}' except the header row?",

  // Sidebar
  language_hint: 'Language (the menu changes after you reload the spreadsheet)',
  loading: 'Loading...',
  draft_label: 'Email template (Gmail draft)',
  reload: 'Refresh',
  draft_hint: 'Use column headers as merge fields in the subject and body, e.g. {{Name}}.',
  no_drafts: '(No drafts found)',
  no_drafts_hint: 'Write the email in Gmail, save it as a draft, then click Refresh.',
  gmail_compose: 'Compose in Gmail',
  sheet_label: 'Recipient sheet',
  from_label: 'From address',
  name_label: "Sender's name",
  cc_label: 'CC (optional)',
  bcc_label: 'BCC (optional)',
  comma_separated: 'Comma-separated',
  slide_check: 'Insert slide image',
  slide_open: 'Open',
  slide_new: 'New',
  slide_hint: 'The first slide is the template. Its {{Name}} etc. are filled in and inserted as an image.',
  slide_none: '(No slides — click New)',
  slide_creating: 'Creating a slide...',
  slide_created: "New slide created. Click 'Open' and add fields such as {{Name}} to the first slide.",
  slide_new_name: 'Mail Merge slide - {date}',
  mode_legend: 'What would you like to do?',
  mode_test: 'Send a test email (1 email to me)',
  mode_run: 'Run mail merge',
  btn_test: 'TEST',
  btn_run: 'RUN',
  btn_reset: 'Reload',
  btn_confirm: 'Send to {n}',
  sending_test: 'Sending a test email...',
  sending_run: 'Sending... Please keep this sidebar open.',
  confirm_text: '{n} recipient(s) to send (already sent: {sent}, invalid address: {invalid}).\nClick the button again to start sending.',
  nothing_to_send: 'Nothing to send (already sent: {sent}, invalid address: {invalid}).\nTo resend, clear EMAIL_SENT in the Mail Merge Status column.',

  // Progress & results
  toast_test: 'Sending a test email...',
  toast_run: 'Running mail merge...',
  toast_done: 'Done.',
  result_title_test: 'Test email result',
  result_title_run: 'Mail merge result',
  result_subject: 'Subject',
  result_sheet: 'Sheet',
  result_from: 'From',
  result_finished: 'Finished',
  result_summary: 'Sent {sent} · Skipped {skipped} · Failed {failed}',
  result_stopped: 'Stopped because of the execution time limit. Click RUN again to send the remaining {n}.',
  result_missing: 'Not found in the sheet (left blank): {fields}',
  label_ok: 'Sent',
  label_skip: 'Skipped',
  label_error: 'Failed',
  log_already_sent: '{email} (row {row}) — already sent',
  log_invalid_email: 'Row {row} — missing or invalid email address ({value})',
  log_test_sent: 'Test email → {to} (using data from row {row})',
  status_invalid_email: 'missing or invalid email address',
  value_empty: 'empty',
  draft_no_subject: '(no subject)',

  // Errors
  err_busy: 'Mail merge is already running. Please try again when it finishes.',
  err_gmail_service: 'The Gmail API service is not added. In the Apps Script editor, go to Services (+) and add "Gmail API".',
  err_slides_service: 'The Google Slides API service is not added. In the Apps Script editor, go to Services (+) and add "Google Slides API".',
  err_api_disabled: 'The {api} API is disabled for this script\'s Google Cloud project. Open the link below, click ENABLE, wait a minute or two and try again.\n{url}',
  err_no_draft_selected: 'Please select an email template (Gmail draft).',
  err_draft_not_found: 'The selected draft was not found. Click Refresh in the sidebar.',
  err_no_rows: 'There are no data rows to send.',
  err_sheet_not_found: "Sheet '{name}' was not found.",
  err_sheet_empty: "Sheet '{name}' is empty.",
  err_no_email_column: "Email column not found. Name the header of the email column 'Email'.",
  err_invalid_address: 'Invalid {label} address: {list}',
};
