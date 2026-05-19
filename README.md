# Voice Memo Journal System

Lightweight Google Sheets + Google Apps Script system for turning voice memos into a searchable journal with derived Goals, To-Dos, and Thoughts tabs.

## What this includes

- `src/Code.js`: Google Apps Script implementation.
- `test/voice_journal.test.js`: local mock tests for dedupe, extraction routing, and retry behavior.
- Google Sheet tabs created by `setupVoiceJournalSheet()`:
  - `Entries`
  - `Goals`
  - `To-Dos`
  - `Thoughts`
  - `Reminders`
  - `Config`

## Setup

1. Create a Google Sheet.
2. Open `Extensions > Apps Script`.
3. Paste the contents of `src/Code.js` into the Apps Script editor.
4. Run `setupVoiceJournalSheet()` once and authorize the script.
5. In the `Config` tab, set `DRIVE_INBOX_FOLDER_ID` to the Google Drive folder where phone voice memos will be uploaded.
6. In the `Config` tab, set `DIGEST_RECIPIENT_EMAIL` to the email address that should receive daily digests.
7. In Apps Script, run:

```javascript
setOpenAiApiKey("your-api-key")
```

The key is stored in Apps Script Script Properties, not in the sheet.

8. Run `processVoiceMemoInbox()` manually once to test voice memo processing.
9. Run `sendDailyDigestNow()` manually once to test email delivery.
10. Run `installVoiceMemoPoller()` to install a 15-minute voice memo poller.
11. Run `installDailyDigestTrigger()` to install the daily digest trigger.

## iPhone Shortcut shape

Create a Shortcut that:

1. Records audio or receives a Voice Memo share sheet item.
2. Saves the audio file to the configured Google Drive inbox folder.
3. Uses a filename that includes the date/time, such as `Voice Journal 2026-05-16 09-30.m4a`.

The Apps Script poller will pick up new Drive files by file ID.

## Review workflow

- Every processed memo creates one row in `Entries`.
- `uploaded_at` is the date the audio file was created in Google Drive, not necessarily the original device recording time.
- Goals and to-dos create linked rows in `Reminders`.
- Thoughts are extracted as big ideas from each entry and tagged with moods.
- Use `status` values like `Open`, `Active`, `Done`, `Complete`, or `Archived` to control what stays active.
- The email digest summarizes active goals, open/recent to-dos, and recent thought reflections.

## Reminder and digest config

Useful `Config` keys:

- `DIGEST_RECIPIENT_EMAIL`: required before email can send.
- `DIGEST_FREQUENCY_PER_WEEK`: default `7`; `5` means weekdays, `3` means Monday/Wednesday/Friday.
- `DIGEST_SEND_HOUR`: default `8`.
- `DIGEST_LOOKBACK_DAYS`: default `7`.
- `DEFAULT_GOAL_ACTIVE_DAYS`: default `30`.
- `DEFAULT_GOAL_REMINDER_FREQUENCY_DAYS`: default `7`.
- `DEFAULT_TODO_REMINDER_FREQUENCY_DAYS`: default `3`.
- `MOOD_TAGS`: allowed mood tags for extracted thoughts.

## Local tests

Run:

```bash
node test/voice_journal.test.js
```

These tests use mocks and do not call Google or OpenAI.
