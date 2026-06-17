/**
 * Inbox
 */

function installPollerFromConfigButton() {
  installVoiceMemoPoller();
  SpreadsheetApp.getActiveSpreadsheet().toast('Voice memo poller installed.', 'Voice Journal', 5);
}

function refreshDashboardFromConfigButton() {
  refreshDashboardNow();
  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard refreshed.', 'Voice Journal', 5);
}

function setOpenAiApiKey(apiKey) {
  apiKey = apiKey || 'Paste your API key here, Run, then delete it!';
  if (!apiKey || !String(apiKey).trim()) {
    throw new Error('API key is required.');
  }
  if (apiKey === 'Paste your API key here, Run, then delete it!') {
    throw new Error('Replace the placeholder text in setOpenAiApiKey() with your real API key, run the function once, then delete the key from the script.');
  }
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', String(apiKey).trim());
}

function installVoiceMemoPoller() {
  removeVoiceMemoPollers_();
  ScriptApp.newTrigger('processVoiceMemoInbox')
    .timeBased()
    .everyMinutes(15)
    .create();
}

function removeVoiceMemoPollers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processVoiceMemoInbox') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function processVoiceMemoInbox() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(spreadsheet);
  if (!config.DRIVE_INBOX_FOLDER_ID) {
    throw new Error('Set DRIVE_INBOX_FOLDER_ID in the Config tab before processing.');
  }

  const folder = DriveApp.getFolderById(config.DRIVE_INBOX_FOLDER_ID);
  const entriesSheet = spreadsheet.getSheetByName(SHEET_NAMES.ENTRIES);
  const existing = buildEntriesIndex_(entriesSheet);
  const files = folder.getFiles();
  const maxRetries = parseInteger_(config.MAX_RETRIES, 3);

  while (files.hasNext()) {
    const file = files.next();
    const driveFileId = file.getId();
    const existingEntry = existing[driveFileId];

    if (existingEntry && !shouldRetryEntry_(existingEntry.values, maxRetries)) {
      continue;
    }

    const entryId = existingEntry
      ? existingEntry.values.entry_id
      : makeId_('entry');

    const retryCount = existingEntry
      ? parseInteger_(existingEntry.values.retry_count, 0)
      : 0;
    const sourceTimestamp = getDriveSourceTimestamp_(file);

    try {
      const result = processVoiceMemoFile_(file, entryId, config);
      upsertEntry_(entriesSheet, existingEntry, {
        entry_id: entryId,
        drive_file_id: driveFileId,
        audio_url: file.getUrl(),
        uploaded_at: sourceTimestamp,
        processed_at: new Date(),
        transcript: result.transcript,
        has_goal: result.extraction.goals.length > 0,
        has_follow_up: result.extraction.followUps.length > 0,
        has_reflection: result.extraction.reflections.length > 0,
        status: 'Processed',
        error: '',
        retry_count: retryCount,
      });

      removeDerivedRowsForEntry_(spreadsheet, entryId);
      appendExtractedRows_(spreadsheet, entryId, result.extraction, config, sourceTimestamp);
      appendCustomExtractionsForEntry_(spreadsheet, entryId, result.transcript, config, sourceTimestamp);
      try {
        indexEntryForSearch_(spreadsheet, entryId, config);
      } catch (indexError) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Search indexing failed for ' + entryId + '.', 'Voice Journal', 5);
      }
      existing[driveFileId] = {
        rowNumber: existingEntry ? existingEntry.rowNumber : entriesSheet.getLastRow(),
        values: { drive_file_id: driveFileId, entry_id: entryId, status: 'Processed', retry_count: retryCount },
      };
    } catch (error) {
      upsertEntry_(entriesSheet, existingEntry, {
        entry_id: entryId,
        drive_file_id: driveFileId,
        audio_url: file.getUrl(),
        uploaded_at: sourceTimestamp,
        processed_at: new Date(),
        transcript: existingEntry ? existingEntry.values.transcript : '',
        has_goal: false,
        has_follow_up: false,
        has_reflection: false,
        status: 'Failed',
        error: errorToString_(error),
        retry_count: retryCount + 1,
      });
    }
  }
  refreshDashboard_(spreadsheet, new Date());
}

function getDriveSourceTimestamp_(file) {
  return file.getLastUpdated ? file.getLastUpdated() : file.getDateCreated();
}

function processVoiceMemoFile_(file, entryId, config) {
  const transcript = transcribeAudio_(file.getBlob(), config);
  const extraction = extractJournalItems_(transcript, entryId, config);
  return { transcript: transcript, extraction: extraction };
}
