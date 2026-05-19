# Voice Memo Journal Setup

This guide walks through setting up the Voice Memo Journal system from a blank Google Sheet to an automated voice memo processing pipeline.

## Requirements

You need:

- A Google account with access to Google Sheets, Google Drive, and Apps Script.
- An OpenAI API key.
- A Google Drive folder that will act as the voice memo inbox.
- A way to upload phone voice memos into that Drive folder, such as an iPhone Shortcut.
- This repository available locally so you can copy `src/Code.js`.

## 1. Create The Google Sheet

1. Create a new Google Sheet.
2. Give it a clear name, such as `Voice Journal`.
3. Open `Extensions > Apps Script`.
4. Delete any starter code in the Apps Script editor.
5. Copy the full contents of `src/Code.js` from this repository.
6. Paste that code into the Apps Script editor.
7. Save the Apps Script project.

## 2. Create The Sheet Tabs

In Apps Script, run:

```javascript
setupVoiceJournalSheet()
```

Google will ask you to authorize the script. Approve the requested permissions so the script can manage the bound sheet, read files from Drive, call external APIs, install triggers, and send email digests.

After the function finishes, the spreadsheet should contain these tabs:

- `Dashboard`
- `Entries`
- `Goals`
- `To-Dos`
- `Thoughts`
- `Reminders`
- `Config`

## 3. Create The Drive Inbox Folder

1. Create a folder in Google Drive for incoming voice memos.
2. Open the folder in your browser.
3. Copy the folder ID from the URL.

The URL usually looks like this:

```text
https://drive.google.com/drive/folders/FOLDER_ID_HERE
```

Copy only the `FOLDER_ID_HERE` part.

## 4. Fill In Required Config Values

Open the `Config` tab in the Google Sheet.

Set:

- `DRIVE_INBOX_FOLDER_ID` to the Drive folder ID.
- `DIGEST_RECIPIENT_EMAIL` to the email address that should receive journal digests.

Optional values can stay at their defaults at first. The most useful ones to revisit later are:

- `DIGEST_FREQUENCY_PER_WEEK`
- `DIGEST_SEND_HOUR`, which defaults to `18` for 6 PM.
- `DIGEST_LOOKBACK_DAYS`
- `DEFAULT_GOAL_ACTIVE_DAYS`
- `DEFAULT_GOAL_REMINDER_FREQUENCY_DAYS`
- `DEFAULT_TODO_REMINDER_FREQUENCY_DAYS`
- `MOOD_TAGS`

## 5. Store The OpenAI API Key

In Apps Script, run:

```javascript
setOpenAiApiKey("your-api-key")
```

Replace `your-api-key` with your actual OpenAI API key.

The key is stored in Apps Script Script Properties. It is not written to the Google Sheet. After running the function once, remove the literal key from the Apps Script editor if you typed it directly into the code.

## 6. Test Manual Processing

Upload one small audio file to the Drive inbox folder.

Then run:

```javascript
processVoiceMemoInbox()
```

Check the spreadsheet after it finishes:

- `Entries` should contain the processed memo and transcript.
- `Goals`, `To-Dos`, and `Thoughts` may contain extracted records depending on the memo content.
- `Reminders` should contain linked reminder rows for extracted goals and to-dos.
- `Dashboard` should refresh with active goals, active to-dos, and the last 7 days of mood tags.

If processing fails, check the `status`, `error`, and `retry_count` columns in `Entries`.

## 7. Test The Dashboard

After at least one goal, to-do, or thought has been added, run:

```javascript
refreshDashboardNow()
```

Confirm that:

- Active to-dos appear on the `Dashboard` tab.
- Checking a dashboard to-do updates the matching source row in `To-Dos` to `Done`.
- Unchecking a dashboard to-do updates the matching source row in `To-Dos` to `Open`.
- Active goals appear on the `Dashboard` tab.
- Selecting `Complete` or `Archive` from a dashboard goal dropdown updates the matching source row in `Goals` and removes the goal from the dashboard.
- Mood tags from `Thoughts` rows created in the last 7 days appear in the mood summary table and pie chart.

You can also refresh from the custom `Voice Journal > Refresh Dashboard` menu after reopening the spreadsheet.

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

## 9. Install Automation Triggers

To check the Drive inbox every 15 minutes, run:

```javascript
installVoiceMemoPoller()
```

To send scheduled digest emails, run:

```javascript
installDailyDigestTrigger()
```

The digest trigger runs daily at `DIGEST_SEND_HOUR`, then uses `DIGEST_FREQUENCY_PER_WEEK` to decide whether that day should send.

## 10. Create An iPhone Shortcut

Create a Shortcut that:

1. Records audio or receives a shared Voice Memo.
2. Saves the audio file to the configured Google Drive inbox folder.
3. Uses a filename that includes the recording date and time.

Example filename:

```text
Voice Journal 2026-05-16 09-30.m4a
```

The current script uses the Google Drive file ID for deduplication and the Drive file creation time for `uploaded_at`.

## 11. Operating The Journal

Use the sheet as the review surface:

- Use `Dashboard` for quick active goal review, active to-do review, and weekly mood trends.
- Mark completed to-dos as `Done`, `Complete`, or `Archived`.
- Mark completed goals as `Done`, `Complete`, or `Archived`.
- Keep active goals as `Active`.
- Keep pending to-dos as `Open`.
- Review mood tags and reflection summaries in `Thoughts`.

Reminder rows are updated after digests are sent. Completed or expired parent records become inactive in reminder logic.

## 12. Local Tests

From the repository root, run:

```bash
node test/voice_journal.test.js
```

These tests use local mocks and do not call Google or OpenAI. They are useful for checking helper behavior before copying code into Apps Script.

## Troubleshooting

If no files are processed, confirm that `DRIVE_INBOX_FOLDER_ID` is set correctly and that the Apps Script project has permission to read the folder.

If transcription fails, confirm that the OpenAI API key was stored with `setOpenAiApiKey()` and that the account has API access.

If extraction fails, check the `EXTRACTION_MODEL` and `EXTRACTION_PROMPT` config values.

If the digest does not send on a scheduled day, confirm `DIGEST_FREQUENCY_PER_WEEK`, `DIGEST_SEND_HOUR`, and the installed triggers in Apps Script.

If the dashboard does not update, run `refreshDashboardNow()` manually and check that `Goals` has `goal_id`, `summary`, and `status` values, that `To-Dos` has `todo_id`, `task`, and `status` values, and that `Thoughts` has `moods` and `created_at` values.

If duplicate files appear, check whether the same audio was uploaded as separate Drive files. Deduplication is based on Google Drive file ID, not audio content.
