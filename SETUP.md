# Voice Memo Journal Setup

This guide walks through setting up the Voice Memo Journal system from a copied workbook/template to an automated voice memo processing pipeline.

## Requirements

You need:

- A Google account with access to Google Sheets, Google Drive, and Apps Script.
- An OpenAI API key.
- A Google Drive folder that will act as the voice memo inbox.
- A way to upload phone voice memos into that Drive folder, such as an iPhone Shortcut.
- This repository available locally so you can copy every `src/*.js` backend file and `src/Index.html`.

## 1. Copy The Workbook And Add The Script

1. Make a copy of the Voice Journal workbook/template.
2. Give the copy a clear name, such as `Voice Journal`.
3. Open `Extensions > Apps Script`.
4. Delete any starter code in the Apps Script editor.
5. For each backend file in `src` whose name ends in `.js`, create a matching script file in Apps Script and paste that file's contents.
6. Keep the numeric filename prefixes, such as `00_Constants`, `10_Main`, and `20_SetupDesign`, so the project remains easy to compare with the repository.
7. Add a new HTML file named `Index`.
8. Copy the full contents of `src/Index.html` into that HTML file.
9. Save the Apps Script project.

## 2. Create The Drive Inbox Folder

1. Create a folder in Google Drive for incoming voice memos.
2. Open the folder in your browser.
3. Copy the folder ID from the URL.

The URL usually looks like this:

```text
https://drive.google.com/drive/folders/FOLDER_ID_HERE
```

Copy only the `FOLDER_ID_HERE` part.

## 3. Fill In Required Config Values

Open the `Config` tab in the copied workbook.

Set:

- `DRIVE_INBOX_FOLDER_ID` to the Drive folder ID.
- `OPENAI_API_KEY_SETUP` to your OpenAI API key. This is temporary; `installVoiceJournal()` stores it in Script Properties and clears the cell.
- `DIGEST_RECIPIENT_EMAIL` to the email address that should receive journal digests, or leave it blank to skip digest trigger installation.

Optional values can stay at their defaults at first. The most useful ones to revisit later are:

- `DIGEST_FREQUENCY_PER_WEEK`
- `DIGEST_SEND_HOUR`, which defaults to `18` for 6 PM.
- `DIGEST_LOOKBACK_DAYS`
- `DEFAULT_GOAL_ACTIVE_DAYS`
- `DEFAULT_GOAL_REMINDER_FREQUENCY_DAYS`
- `DEFAULT_FOLLOW_UP_REMINDER_FREQUENCY_DAYS`
- `MOOD_TAGS`

## 4. Run The Installer

In Apps Script, run:

```javascript
installVoiceJournal()
```

Google will ask you to authorize the script. Approve the requested permissions so the script can manage the bound sheet, read files from Drive, call external APIs, install triggers, and send email digests.

After setup has rendered the Config controls, the `Install / Repair` quick-action button on the `Config` tab runs the same installer.

After the function finishes, the spreadsheet should contain these tabs:

- `Dashboard`
- `Entries`
- `Goals`
- `Follow Ups`
- `Reflections`
- `Design`
- `Search Index` hidden
- `Config`

`installVoiceJournal()` repairs/migrates the workbook, moves `Dashboard` to the leftmost tab, formats the source tables, hides `Search Index`, seeds `Design`, stores the OpenAI API key securely, clears `OPENAI_API_KEY_SETUP`, installs the inbox poller, installs the digest trigger if `DIGEST_RECIPIENT_EMAIL` is filled, and refreshes the dashboard.

If it fails, read the error message, fill the missing Config values, and run `installVoiceJournal()` again.

## 5. Test Manual Processing

Upload one small audio file to the Drive inbox folder.

Then run:

```javascript
processVoiceMemoInbox()
```

Check the spreadsheet after it finishes:

- `Entries` should contain the processed memo and transcript.
- `Goals`, `Follow Ups`, and `Reflections` may contain extracted records depending on the memo content.
- `Dashboard` should refresh with active goals, active follow ups, and the last 7 days of mood tags.

If processing fails, check the `status`, `error`, and `retry_count` columns in `Entries`.

## 6. Deploy And Test The Search Web App

After setup has created the sheet tabs and the OpenAI API key is stored, deploy the search interface:

1. In Apps Script, click `Deploy > New deployment`.
2. Choose `Web app`.
3. Set `Execute as` to `Me`.
4. Set access to only yourself or authorized users.
5. Deploy and authorize any requested permissions.
6. Copy the web app URL.

Open the web app URL in your browser. The page should show `Voice Journal Search`, index status, a query box, and buttons for `Search`, `Refresh Index`, and `Rebuild Index`.

Before the first search, click `Rebuild Index` in the web app, or run `Voice Journal > Rebuild Search Index` from the spreadsheet menu. After that, new processed memos are indexed automatically, and `Refresh Index` updates only changed entries.

## 7. Test The Dashboard

After at least one goal, follow up, or reflection has been added, run:

```javascript
refreshDashboardNow()
```

Confirm that:

- Active follow ups appear on the `Dashboard` tab.
- Checking a dashboard follow up updates the matching source row in `Follow Ups` to `Done`.
- Unchecking a dashboard follow up updates the matching source row in `Follow Ups` to `Open`.
- Active goals appear on the `Dashboard` tab.
- Selecting `Complete` or `Archive` from a dashboard goal dropdown updates the matching source row in `Goals` and removes the goal from the dashboard.
- Mood tags from `Reflections` rows created in the last 7 days appear in the mood summary table and pie chart.

You can also refresh from the custom `Voice Journal > Refresh Dashboard` menu after reopening the spreadsheet.
The `Config` tab also has an `Update Dashboard` quick-action button.

## 8. Test The Digest

After at least one memo has been processed, run:

```javascript
sendDailyDigestNow()
```

Confirm that the configured recipient receives an email. If no email arrives, check:

- The `DIGEST_RECIPIENT_EMAIL` value.
- Apps Script execution logs.
- Gmail spam or filtering rules.
- Whether Apps Script authorization was completed.

## 9. Try A Custom Design Row

Open the `Design` tab. It starts with built-in rows for `follow_ups`, `goals`, and `reflections`.

To add a custom extraction:

1. Add a row with a unique `extraction_key`, a display name, a destination `sheet_name`, a description, and comma-separated keywords.
2. Optionally add `fact_table_dictionary` JSON to define custom fields. Leave it blank to use `summary`, `evidence`, and `status`.
3. Leave `ai_decides_without_keyword` as `FALSE` if the AI should only run when at least one keyword appears in the transcript.
4. Set `status` to `Active` if future inbox processing should apply this extraction.
5. Choose `Create Sheet` in the `action` cell to create the destination sheet, or choose `Run Backfill` to process existing entries in batches.

Example `fact_table_dictionary` for a `books_to_read` extraction:

```json
{
  "fields": [
    { "key": "summary", "type": "string", "description": "Book title or recommendation summary." },
    { "key": "author", "type": "string", "description": "Author name, blank if unknown." },
    { "key": "priority", "type": "enum", "description": "Reading priority.", "values": ["low", "medium", "high"] }
  ]
}
```

Custom field keys must be `snake_case`. Supported types are `string`, `number`, `boolean`, `date`, and `enum`. Reserved metadata columns cannot be overridden: the custom id column, `entry_id`, `matched_keywords`, `created_at`, and `extracted_at`. Adding fields adds columns; removing fields does not delete old columns; renaming a field creates a new column.

If the inline action does not run because of Apps Script trigger authorization, select the row and use `Voice Journal > Run Selected Design Action`.

## 10. Manual Repair Functions

Most users should only need `installVoiceJournal()`. The individual functions remain available for repair or advanced testing:

- `setupVoiceJournalSheet()` repairs tabs, headers, formatting, Config rows, and Design rows.
- `installVoiceMemoPoller()` reinstalls only the Drive inbox poller.
- `installDailyDigestTrigger()` reinstalls only the digest trigger.
- `setOpenAiApiKey("your-api-key")` stores the API key directly in Script Properties if you do not want to use `OPENAI_API_KEY_SETUP`.

The digest trigger runs daily at `DIGEST_SEND_HOUR`, then uses `DIGEST_FREQUENCY_PER_WEEK` to decide whether that day should send.

## 11. Create An iPhone Shortcut

Create a Shortcut that:

1. Records audio or receives a shared Voice Memo.
2. Saves the audio file to the configured Google Drive inbox folder.
3. Uses a filename that includes the recording date and time.

Example filename:

```text
Voice Journal 2026-05-16 09-30.m4a
```

The current script uses the Google Drive file ID for deduplication and the Drive file last-updated time for `uploaded_at` and extracted fact `created_at` values.

## 12. Operating The Journal

Use the sheet as the review surface:

- Use `Dashboard` for quick active goal review, active follow up review, and weekly mood trends.
- Mark completed follow ups as `Done`, `Complete`, or `Archived`.
- Mark completed goals as `Done`, `Complete`, or `Archived`.
- Keep active goals as `Active`.
- Keep pending follow ups as `Open`.
- Review mood tags and reflection summaries in `Reflections`.

Use the web app as the search surface:

- Type a query into the web app and click `Search`.
- Review ranked exact transcript quotes and source metadata.
- Use `Refresh Index` if transcripts or reflections were edited manually.
- Use `Rebuild Index` if embeddings need to be regenerated from scratch.

Goals and follow ups update their own reminder state after digests are sent. Completed or expired records stop appearing as active digest reminders.

## 13. Local Tests

From the repository root, run:

```bash
node test/voice_journal.test.js
```

These tests use local mocks and do not call Google or OpenAI. They are useful for checking helper behavior before copying code into Apps Script.

## Troubleshooting

If no files are processed, confirm that `DRIVE_INBOX_FOLDER_ID` is set correctly and that the Apps Script project has permission to read the folder.

If transcription fails, confirm that `installVoiceJournal()` cleared `OPENAI_API_KEY_SETUP`, that the OpenAI API key exists in Script Properties, and that the account has API access.

If extraction fails, check the `EXTRACTION_MODEL` and `EXTRACTION_PROMPT` config values.

If the digest does not send on a scheduled day, confirm `DIGEST_FREQUENCY_PER_WEEK`, `DIGEST_SEND_HOUR`, and the installed triggers in Apps Script.

If the dashboard does not update, run `refreshDashboardNow()` manually and check that `Goals` has `goal_id`, `summary`, and `status` values, that `Follow Ups` has `follow_up_id`, `summary`, and `status` values, and that `Reflections` has `moods` and `created_at` values.

If duplicate files appear, check whether the same audio was uploaded as separate Drive files. Deduplication is based on Google Drive file ID, not audio content.
