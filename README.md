# Voice Memo Journal System

Voice Memo Journal is a lightweight Google Sheets and Google Apps Script project for turning spoken voice memos into a searchable personal journal. It uses a Google Drive inbox folder, OpenAI transcription and extraction, and a Google Sheet as the review surface for entries, goals, to-dos, thoughts, reminders, and digest configuration.

For installation instructions, see [SETUP.md](SETUP.md).

## Purpose

The project is designed for fast voice capture. Instead of manually rewriting notes, tasks, and reflections after recording a memo, you upload audio files to a Drive folder and let the Apps Script pipeline process them into structured rows.

The system is meant to help with:

- Capturing journal entries from voice memos.
- Extracting goals and to-dos from natural speech.
- Saving larger reflections as tagged thoughts.
- Creating reminder rows for active goals and open tasks.
- Sending an email digest of active goals, relevant to-dos, and recent reflections.

## Pipeline

1. A voice memo is recorded on a phone.
2. The audio file is uploaded into a configured Google Drive inbox folder.
3. A Google Apps Script poller checks the folder every 15 minutes.
4. New audio files are transcribed with the OpenAI transcription API.
5. The transcript is sent to an OpenAI text model for structured extraction.
6. The system writes the original entry and derived records into Google Sheet tabs.
7. A scheduled digest email summarizes the active items and recent reflections.

## Project Structure

- `src/Code.js`: Google Apps Script implementation.
- `test/voice_journal.test.js`: Local mock tests for pure helper behavior.
- `README.md`: High-level project overview.
- `Setup.MD`: Detailed setup and operating instructions.

## Google Sheet Tabs

Running `setupVoiceJournalSheet()` creates and maintains these tabs:

- `Entries`: One row per processed voice memo, including transcript, status, retry count, and error details.
- `Goals`: Goal records extracted from transcripts.
- `To-Dos`: Task records extracted from transcripts.
- `Thoughts`: Reflection records with mood tags.
- `Reminders`: Reminder rows linked to goals and to-dos.
- `Config`: User-editable configuration values.

## Review Workflow

Every processed memo creates one `Entries` row. Extracted goals, to-dos, and thoughts are linked back to that entry through `entry_id`.

The `status` fields are intentionally simple. Values like `Open`, `Active`, `Done`, `Complete`, and `Archived` control what remains active in reminders and digests. Completed or archived goals and to-dos are ignored by future reminder logic.

The `uploaded_at` value comes from the Google Drive file creation time. If the original device recording time matters, include it in the filename or transcript context.

## Configuration

The most important `Config` values are:

- `DRIVE_INBOX_FOLDER_ID`: Google Drive folder where phone voice memos are uploaded.
- `DIGEST_RECIPIENT_EMAIL`: Email address that receives journal digests.
- `TRANSCRIPTION_MODEL`: OpenAI audio transcription model.
- `EXTRACTION_MODEL`: OpenAI text model for structured extraction.
- `DIGEST_FREQUENCY_PER_WEEK`: Number of digest days per week.
- `DIGEST_SEND_HOUR`: Hour of day for scheduled digest delivery.
- `DIGEST_LOOKBACK_DAYS`: Number of days included in recent thought and to-do review.
- `MOOD_TAGS`: Allowed mood tags for extracted thoughts.

See [Setup.MD](Setup.MD) for the full setup flow.

## Local Tests

Run the local mock tests with:

```bash
node test/voice_journal.test.js
```

These tests do not call Google or OpenAI services. They are intended to validate pure logic such as normalization, retry behavior, and digest scheduling rules.

## Current Limitations

- The Apps Script source is currently deployed by copying `src/Code.js` into a bound Apps Script project.
- The pipeline does not archive processed audio files by default.
- Date inference depends on transcript context and may be imperfect for phrases like "tomorrow" or "next week."
- The digest is plain-text email, not an interactive dashboard.

## Credits

- Project concept and repository: Marko Suchy.
- Implementation assistance: OpenAI ChatGPT/Codex.
- Platform services: Google Sheets, Google Drive, Google Apps Script, and OpenAI APIs.
