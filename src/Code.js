/**
 * Voice Memo Journal System
 *
 * Paste this file into a Google Apps Script project bound to your Google Sheet.
 * Configure the Drive inbox folder in the Config tab and set your OpenAI API key
 * with setOpenAiApiKey("sk-...").
 */

const SHEET_NAMES = {
  DASHBOARD: 'Dashboard',
  ENTRIES: 'Entries',
  GOALS: 'Goals',
  TODOS: 'To-Dos',
  THOUGHTS: 'Thoughts',
  CONFIG: 'Config',
};

const MOOD_TAGS = [
  'energized',
  'calm',
  'hopeful',
  'curious',
  'focused',
  'uncertain',
  'frustrated',
  'anxious',
  'tired',
  'grateful',
  'reflective',
  'excited',
];

const DONE_STATUSES = ['done', 'complete', 'completed', 'archived'];

const DASHBOARD = {
  TODO_START_ROW: 5,
  TODO_START_COLUMN: 1,
  TODO_COLUMN_COUNT: 6,
  TODO_ID_COLUMN: 6,
  GOAL_TITLE_ROW: 3,
  GOAL_HEADER_ROW: 4,
  GOAL_START_ROW: 5,
  GOAL_START_COLUMN: 8,
  GOAL_COLUMN_COUNT: 6,
  GOAL_ID_COLUMN: 13,
  MOOD_START_ROW: 5,
  MOOD_START_COLUMN: 15,
  MOOD_COLUMN_COUNT: 2,
  MOOD_LOOKBACK_DAYS: 7,
};

const HEADERS = {
  Entries: [
    'entry_id',
    'drive_file_id',
    'audio_url',
    'uploaded_at',
    'processed_at',
    'transcript',
    'has_goal',
    'has_todo',
    'has_thought',
    'status',
    'error',
    'retry_count',
  ],
  Goals: [
    'goal_id',
    'entry_id',
    'summary',
    'status',
    'created_at',
    'active_until',
    'reminder_frequency_days',
    'next_reminder_at',
    'last_reminded_at',
    'reminder_count',
  ],
  'To-Dos': [
    'todo_id',
    'entry_id',
    'task',
    'status',
    'created_at',
    'due_date_optional',
    'reminder_frequency_days',
    'next_reminder_at',
    'last_reminded_at',
    'reminder_count',
  ],
  Thoughts: [
    'thought_id',
    'entry_id',
    'summary',
    'moods',
    'created_at',
  ],
  Config: ['key', 'value', 'notes'],
};

const DEFAULT_CONFIG = [
  ['DRIVE_INBOX_FOLDER_ID', '', 'Google Drive folder ID where phone voice memos are uploaded.'],
  ['OPENAI_API_KEY_PROPERTY', 'OPENAI_API_KEY', 'Script Properties key that stores the OpenAI API key. Do not put the API key in this sheet.'],
  ['TRANSCRIPTION_MODEL', 'gpt-4o-mini-transcribe', 'OpenAI audio transcription model.'],
  ['EXTRACTION_MODEL', 'gpt-5.2', 'OpenAI text model used for structured extraction.'],
  ['MAX_RETRIES', '3', 'Maximum processing attempts per memo.'],
  ['DEFAULT_TODO_STATUS', 'Open', 'Initial status for new to-do rows.'],
  ['DEFAULT_GOAL_STATUS', 'Active', 'Initial status for new goal rows.'],
  ['DEFAULT_GOAL_ACTIVE_DAYS', '30', 'Default number of days a goal remains active when the memo is unclear.'],
  ['DEFAULT_GOAL_REMINDER_FREQUENCY_DAYS', '7', 'Default reminder frequency for goals when the memo is unclear.'],
  ['DEFAULT_TODO_REMINDER_FREQUENCY_DAYS', '3', 'Default reminder frequency for to-dos when the memo is unclear.'],
  ['DIGEST_RECIPIENT_EMAIL', '', 'Required. Daily digest recipient email address.'],
  ['DIGEST_FREQUENCY_PER_WEEK', '7', 'How many days per week to send the digest. 7 = daily, 5 = weekdays, 3 = Mon/Wed/Fri.'],
  ['DIGEST_SEND_HOUR', '18', 'Hour of day, 0-23, for the daily digest trigger. Default 18 = 6 PM.'],
  ['DIGEST_LOOKBACK_DAYS', '7', 'Number of days of recent thoughts and to-dos to include in the digest.'],
  ['MOOD_TAGS', MOOD_TAGS.join(','), 'Allowed mood tags for extracted thoughts.'],
  ['EXTRACTION_PROMPT', defaultExtractionPrompt_(), 'Prompt used to extract goals, to-dos, and big-idea thoughts from transcripts.'],
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Voice Journal')
    .addItem('Setup Sheet', 'setupVoiceJournalSheet')
    .addItem('Process Inbox Now', 'processVoiceMemoInbox')
    .addItem('Refresh Dashboard', 'refreshDashboardNow')
    .addItem('Install 15-Minute Poller', 'installVoiceMemoPoller')
    .addSeparator()
    .addItem('Send Digest Now', 'sendDailyDigestNow')
    .addItem('Install Daily Digest Trigger', 'installDailyDigestTrigger')
    .addItem('Remove Daily Digest Trigger', 'removeDailyDigestTrigger')
    .addToUi();
}

function setupVoiceJournalSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  migrateKnownSchemaChanges_(spreadsheet);
  Object.keys(HEADERS).forEach(function(sheetName) {
    const sheet = getOrCreateSheet_(spreadsheet, sheetName);
    ensureHeaders_(sheet, HEADERS[sheetName]);
    sheet.setFrozenRows(1);
  });
  getOrCreateSheet_(spreadsheet, SHEET_NAMES.DASHBOARD);
  seedConfigDefaults_(spreadsheet.getSheetByName(SHEET_NAMES.CONFIG));
}

function refreshDashboardNow() {
  setupVoiceJournalSheet();
  refreshDashboard_(SpreadsheetApp.getActiveSpreadsheet(), new Date());
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.DASHBOARD) {
    return;
  }

  const spreadsheet = e.source || SpreadsheetApp.getActiveSpreadsheet();
  if (handleDashboardTodoEdit_(spreadsheet, sheet, e)) {
    return;
  }

  handleDashboardGoalEdit_(spreadsheet, sheet, e);
}

function handleDashboardTodoEdit_(spreadsheet, sheet, e) {
  if (e.range.getColumn() !== DASHBOARD.TODO_START_COLUMN || e.range.getRow() < DASHBOARD.TODO_START_ROW) {
    return false;
  }

  const todoId = sheet.getRange(e.range.getRow(), DASHBOARD.TODO_ID_COLUMN).getValue();
  if (!todoId) {
    return false;
  }

  const nextStatus = dashboardCheckboxStatus_(e.value);
  updateTodoStatusById_(spreadsheet.getSheetByName(SHEET_NAMES.TODOS), todoId, nextStatus);
  sheet.getRange(e.range.getRow(), 3).setValue(nextStatus);
  return true;
}

function handleDashboardGoalEdit_(spreadsheet, sheet, e) {
  if (e.range.getColumn() !== DASHBOARD.GOAL_START_COLUMN || e.range.getRow() < DASHBOARD.GOAL_START_ROW) {
    return false;
  }

  const nextStatus = dashboardGoalActionStatus_(e.value);
  if (!nextStatus) {
    return false;
  }

  const goalId = sheet.getRange(e.range.getRow(), DASHBOARD.GOAL_ID_COLUMN).getValue();
  if (!goalId) {
    return false;
  }

  updateGoalStatusById_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS), goalId, nextStatus);
  refreshDashboard_(spreadsheet, new Date());
  return true;
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

function installDailyDigestTrigger() {
  setupVoiceJournalSheet();
  removeDailyDigestTrigger();
  const config = readConfig_(SpreadsheetApp.getActiveSpreadsheet());
  const sendHour = clampInteger_(parseInteger_(config.DIGEST_SEND_HOUR, 18), 0, 23);
  ScriptApp.newTrigger('sendScheduledDailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(sendHour)
    .create();
}

function removeDailyDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendScheduledDailyDigest') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function sendDailyDigestNow() {
  return sendDigest_(true);
}

function sendScheduledDailyDigest() {
  return sendDigest_(false);
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

    try {
      const result = processVoiceMemoFile_(file, entryId, config);
      upsertEntry_(entriesSheet, existingEntry, {
        entry_id: entryId,
        drive_file_id: driveFileId,
        audio_url: file.getUrl(),
        uploaded_at: file.getDateCreated(),
        processed_at: new Date(),
        transcript: result.transcript,
        has_goal: result.extraction.goals.length > 0,
        has_todo: result.extraction.todos.length > 0,
        has_thought: result.extraction.thoughts.length > 0,
        status: 'Processed',
        error: '',
        retry_count: retryCount,
      });

      removeDerivedRowsForEntry_(spreadsheet, entryId);
      appendExtractedRows_(spreadsheet, entryId, result.extraction, config);
      existing[driveFileId] = {
        rowNumber: existingEntry ? existingEntry.rowNumber : entriesSheet.getLastRow(),
        values: { drive_file_id: driveFileId, entry_id: entryId, status: 'Processed', retry_count: retryCount },
      };
    } catch (error) {
      upsertEntry_(entriesSheet, existingEntry, {
        entry_id: entryId,
        drive_file_id: driveFileId,
        audio_url: file.getUrl(),
        uploaded_at: file.getDateCreated(),
        processed_at: new Date(),
        transcript: existingEntry ? existingEntry.values.transcript : '',
        has_goal: false,
        has_todo: false,
        has_thought: false,
        status: 'Failed',
        error: errorToString_(error),
        retry_count: retryCount + 1,
      });
    }
  }
  refreshDashboard_(spreadsheet, new Date());
}

function processVoiceMemoFile_(file, entryId, config) {
  const transcript = transcribeAudio_(file.getBlob(), config);
  const extraction = extractJournalItems_(transcript, entryId, config);
  return { transcript: transcript, extraction: extraction };
}

function transcribeAudio_(audioBlob, config) {
  const apiKey = getOpenAiApiKey_(config);
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions', {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
    payload: {
      model: config.TRANSCRIPTION_MODEL || 'gpt-4o-mini-transcribe',
      file: audioBlob,
      response_format: 'json',
    },
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Transcription failed with HTTP ' + code + ': ' + body);
  }

  const parsed = JSON.parse(body);
  if (!parsed.text) {
    throw new Error('Transcription response did not include text.');
  }
  return parsed.text;
}

function extractJournalItems_(transcript, entryId, config) {
  const apiKey = getOpenAiApiKey_(config);
  const prompt = config.EXTRACTION_PROMPT || defaultExtractionPrompt_();
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
    payload: JSON.stringify({
      model: config.EXTRACTION_MODEL || 'gpt-5.2',
      input: [
        {
          role: 'system',
          content: prompt,
        },
        {
          role: 'user',
          content: 'Entry ID: ' + entryId + '\n\nTranscript:\n' + transcript,
        },
      ],
      text: {
        format: {
          type: 'json_schema',
          name: 'voice_journal_extraction',
          strict: true,
          schema: extractionSchema_(),
        },
      },
    }),
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Extraction failed with HTTP ' + code + ': ' + body);
  }

  const parsed = JSON.parse(body);
  const outputText = parsed.output_text || findOutputText_(parsed);
  if (!outputText) {
    throw new Error('Extraction response did not include output text.');
  }

  return normalizeExtraction_(JSON.parse(outputText));
}

function appendExtractedRows_(spreadsheet, entryId, extraction, config) {
  const now = new Date();
  const todoStatus = config.DEFAULT_TODO_STATUS || 'Open';
  const goalStatus = config.DEFAULT_GOAL_STATUS || 'Active';
  const defaultGoalDays = parseInteger_(config.DEFAULT_GOAL_ACTIVE_DAYS, 30);
  const defaultGoalFrequency = parseInteger_(config.DEFAULT_GOAL_REMINDER_FREQUENCY_DAYS, 7);
  const defaultTodoFrequency = parseInteger_(config.DEFAULT_TODO_REMINDER_FREQUENCY_DAYS, 3);

  const goalsSheet = spreadsheet.getSheetByName(SHEET_NAMES.GOALS);
  extraction.goals.forEach(function(goal) {
    const goalId = makeId_('goal');
    const activeUntil = parseDate_(goal.active_until) || addDays_(now, defaultGoalDays);
    const frequencyDays = positiveInteger_(goal.reminder_frequency_days, defaultGoalFrequency);
    appendObjectRow_(goalsSheet, {
      goal_id: goalId,
      entry_id: entryId,
      summary: goal.summary,
      status: goal.status || goalStatus,
      created_at: now,
      active_until: activeUntil,
      reminder_frequency_days: frequencyDays,
      next_reminder_at: now,
      last_reminded_at: '',
      reminder_count: 0,
    });
  });

  const todosSheet = spreadsheet.getSheetByName(SHEET_NAMES.TODOS);
  extraction.todos.forEach(function(todo) {
    const todoId = makeId_('todo');
    const dueDate = parseDate_(todo.due_date_optional);
    const frequencyDays = positiveInteger_(todo.reminder_frequency_days, defaultTodoFrequency);
    appendObjectRow_(todosSheet, {
      todo_id: todoId,
      entry_id: entryId,
      task: todo.task,
      status: todo.status || todoStatus,
      created_at: now,
      due_date_optional: dueDate || todo.due_date_optional || '',
      reminder_frequency_days: frequencyDays,
      next_reminder_at: now,
      last_reminded_at: '',
      reminder_count: 0,
    });
  });

  const thoughtsSheet = spreadsheet.getSheetByName(SHEET_NAMES.THOUGHTS);
  extraction.thoughts.forEach(function(thought) {
    appendObjectRow_(thoughtsSheet, {
      thought_id: makeId_('thought'),
      entry_id: entryId,
      summary: thought.summary,
      moods: thought.moods.join(', '),
      created_at: now,
    });
  });
}

function removeDerivedRowsForEntry_(spreadsheet, entryId) {
  [SHEET_NAMES.GOALS, SHEET_NAMES.TODOS, SHEET_NAMES.THOUGHTS].forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    deleteRowsWhereColumnEquals_(sheet, 'entry_id', entryId);
  });
}

function refreshDashboard_(spreadsheet, now) {
  const dashboardSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.DASHBOARD);
  const todoRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.TODOS));
  const goalRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS));
  const thoughtRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.THOUGHTS));
  const dashboardTodos = buildDashboardTodoRows_(todoRows);
  const dashboardGoals = buildDashboardGoalRows_(goalRows);
  const moodCounts = buildRecentMoodCounts_(thoughtRows, now, DASHBOARD.MOOD_LOOKBACK_DAYS);
  const moodRows = moodCountsToRows_(moodCounts);

  removeCharts_(dashboardSheet);
  dashboardSheet.clear();
  dashboardSheet.setHiddenGridlines(true);
  dashboardSheet.setFrozenRows(0);
  dashboardSheet.getRange('A1').setValue('Voice Journal Dashboard').setFontSize(18).setFontWeight('bold');
  dashboardSheet.getRange('A2').setValue('Last refreshed: ' + formatDateTimeForDashboard_(now));

  renderDashboardTodos_(dashboardSheet, dashboardTodos);
  renderDashboardGoals_(dashboardSheet, dashboardGoals);
  renderDashboardMoods_(dashboardSheet, moodRows);
}

function renderDashboardTodos_(sheet, dashboardTodos) {
  const headerRange = sheet.getRange(4, DASHBOARD.TODO_START_COLUMN, 1, DASHBOARD.TODO_COLUMN_COUNT);
  headerRange
    .setValues([['Done', 'Task', 'Status', 'Due Date', 'Created', 'todo_id']])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.getRange('A3').setValue('Active To-Dos').setFontWeight('bold');

  if (dashboardTodos.length === 0) {
    sheet.getRange(DASHBOARD.TODO_START_ROW, DASHBOARD.TODO_START_COLUMN).setValue('No active to-dos.');
  } else {
    const range = sheet.getRange(
      DASHBOARD.TODO_START_ROW,
      DASHBOARD.TODO_START_COLUMN,
      dashboardTodos.length,
      DASHBOARD.TODO_COLUMN_COUNT
    );
    range.setValues(dashboardTodos);
    sheet.getRange(DASHBOARD.TODO_START_ROW, DASHBOARD.TODO_START_COLUMN, dashboardTodos.length, 1).insertCheckboxes();
    sheet.getRange(DASHBOARD.TODO_START_ROW, 4, dashboardTodos.length, 2).setNumberFormat('yyyy-mm-dd');
  }

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 360);
  sheet.getRange(1, 2, sheet.getMaxRows(), 1).setWrap(true);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.hideColumns(DASHBOARD.TODO_ID_COLUMN);
}

function renderDashboardGoals_(sheet, dashboardGoals) {
  const headerRange = sheet.getRange(DASHBOARD.GOAL_HEADER_ROW, DASHBOARD.GOAL_START_COLUMN, 1, DASHBOARD.GOAL_COLUMN_COUNT);
  headerRange
    .setValues([['Action', 'Goal', 'Status', 'Active Until', 'Created', 'goal_id']])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.getRange(DASHBOARD.GOAL_TITLE_ROW, DASHBOARD.GOAL_START_COLUMN).setValue('Active Goals').setFontWeight('bold');

  if (dashboardGoals.length === 0) {
    sheet.getRange(DASHBOARD.GOAL_START_ROW, DASHBOARD.GOAL_START_COLUMN).setValue('No active goals.');
  } else {
    const range = sheet.getRange(
      DASHBOARD.GOAL_START_ROW,
      DASHBOARD.GOAL_START_COLUMN,
      dashboardGoals.length,
      DASHBOARD.GOAL_COLUMN_COUNT
    );
    range.setValues(dashboardGoals);
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Complete', 'Archive'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(DASHBOARD.GOAL_START_ROW, DASHBOARD.GOAL_START_COLUMN, dashboardGoals.length, 1).setDataValidation(validation);
    sheet.getRange(DASHBOARD.GOAL_START_ROW, DASHBOARD.GOAL_START_COLUMN + 3, dashboardGoals.length, 2).setNumberFormat('yyyy-mm-dd');
  }

  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN, 100);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 1, 340);
  sheet.getRange(1, DASHBOARD.GOAL_START_COLUMN + 1, sheet.getMaxRows(), 1).setWrap(true);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 2, 90);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 3, 120);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 4, 110);
  sheet.hideColumns(DASHBOARD.GOAL_ID_COLUMN);
}

function renderDashboardMoods_(sheet, moodRows) {
  sheet.getRange(3, DASHBOARD.MOOD_START_COLUMN).setValue('Moods - Last 7 Days').setFontWeight('bold');
  sheet.getRange(
    4,
    DASHBOARD.MOOD_START_COLUMN,
    1,
    DASHBOARD.MOOD_COLUMN_COUNT
  ).setValues([['Mood', 'Count']]).setFontWeight('bold').setBackground('#f1f3f4');

  if (moodRows.length === 0) {
    sheet.getRange(DASHBOARD.MOOD_START_ROW, DASHBOARD.MOOD_START_COLUMN).setValue('No moods in the last 7 days.');
    return;
  }

  const moodRange = sheet.getRange(
    DASHBOARD.MOOD_START_ROW,
    DASHBOARD.MOOD_START_COLUMN,
    moodRows.length,
    DASHBOARD.MOOD_COLUMN_COUNT
  );
  moodRange.setValues(moodRows);
  sheet.setColumnWidth(DASHBOARD.MOOD_START_COLUMN, 140);
  sheet.setColumnWidth(DASHBOARD.MOOD_START_COLUMN + 1, 80);

  const chart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange(4, DASHBOARD.MOOD_START_COLUMN, moodRows.length + 1, DASHBOARD.MOOD_COLUMN_COUNT))
    .setPosition(4, DASHBOARD.MOOD_START_COLUMN + 3, 0, 0)
    .setOption('title', 'Moods - Last 7 Days')
    .setOption('pieHole', 0.35)
    .build();
  sheet.insertChart(chart);
}

function removeCharts_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
}

function buildDashboardTodoRows_(rows) {
  return normalizeRowsForValues_(rows)
    .filter(function(todo) {
      return todo.todo_id && todo.task && !isDoneStatus_(todo.status);
    })
    .map(function(todo) {
      return [
        false,
        todo.task,
        todo.status || 'Open',
        parseDate_(todo.due_date_optional) || '',
        parseDate_(todo.created_at) || '',
        todo.todo_id,
      ];
    });
}

function buildDashboardGoalRows_(rows) {
  return normalizeRowsForValues_(rows)
    .filter(function(goal) {
      return goal.goal_id && goal.summary && !isDoneStatus_(goal.status);
    })
    .map(function(goal) {
      return [
        '',
        goal.summary,
        goal.status || 'Active',
        parseDate_(goal.active_until) || '',
        parseDate_(goal.created_at) || '',
        goal.goal_id,
      ];
    });
}

function buildRecentMoodCounts_(rows, now, lookbackDays) {
  const since = addDays_(now, -lookbackDays);
  const counts = {};
  normalizeRowsForValues_(rows).forEach(function(thought) {
    const createdAt = parseDate_(thought.created_at);
    if (!isRecent_(createdAt, since)) {
      return;
    }

    String(thought.moods || '').split(',').forEach(function(rawMood) {
      const mood = rawMood.trim().toLowerCase();
      if (mood) {
        counts[mood] = (counts[mood] || 0) + 1;
      }
    });
  });
  return counts;
}

function moodCountsToRows_(counts) {
  return Object.keys(counts || {})
    .sort(function(a, b) {
      if (counts[b] !== counts[a]) {
        return counts[b] - counts[a];
      }
      return a < b ? -1 : a > b ? 1 : 0;
    })
    .map(function(mood) {
      return [mood, counts[mood]];
    });
}

function normalizeRowsForValues_(rows) {
  return normalizeArray_(rows).map(function(row) {
    return row && row.values ? row.values : row;
  });
}

function dashboardCheckboxStatus_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE' ? 'Done' : 'Open';
}

function dashboardGoalActionStatus_(value) {
  const action = String(value || '').trim().toLowerCase();
  if (action === 'complete') {
    return 'Complete';
  }
  if (action === 'archive') {
    return 'Archived';
  }
  return '';
}

function updateTodoStatusById_(todosSheet, todoId, status) {
  const headers = getHeaders_(todosSheet);
  const idColumnIndex = headers.indexOf('todo_id');
  const statusColumnIndex = headers.indexOf('status');
  if (idColumnIndex === -1 || statusColumnIndex === -1 || todosSheet.getLastRow() < 2) {
    return false;
  }

  const ids = todosSheet.getRange(2, idColumnIndex + 1, todosSheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === todoId) {
      todosSheet.getRange(i + 2, statusColumnIndex + 1).setValue(status);
      return true;
    }
  }
  return false;
}

function updateGoalStatusById_(goalsSheet, goalId, status) {
  const headers = getHeaders_(goalsSheet);
  const idColumnIndex = headers.indexOf('goal_id');
  const statusColumnIndex = headers.indexOf('status');
  if (idColumnIndex === -1 || statusColumnIndex === -1 || goalsSheet.getLastRow() < 2) {
    return false;
  }

  const ids = goalsSheet.getRange(2, idColumnIndex + 1, goalsSheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === goalId) {
      goalsSheet.getRange(i + 2, statusColumnIndex + 1).setValue(status);
      return true;
    }
  }
  return false;
}

function sendDigest_(forceSend) {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(spreadsheet);
  const recipient = String(config.DIGEST_RECIPIENT_EMAIL || '').trim();
  if (!recipient) {
    throw new Error('Set DIGEST_RECIPIENT_EMAIL in the Config tab before sending the digest.');
  }

  const now = new Date();
  if (!forceSend && !shouldSendDigestOnDate_(now, config.DIGEST_FREQUENCY_PER_WEEK)) {
    return 'Skipped: today is not an enabled digest day.';
  }

  const digest = buildDigest_(spreadsheet, config, now);
  MailApp.sendEmail({
    to: recipient,
    subject: 'Voice Journal Digest - ' + formatDateForDigest_(now),
    body: digest.body,
    htmlBody: digest.htmlBody,
  });
  applyDigestStateUpdates_(spreadsheet, digest);
  return 'Sent digest to ' + recipient + '.';
}

function buildDigest_(spreadsheet, config, now) {
  return buildDigestFromRows_(
    getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS)),
    getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.TODOS)),
    getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.THOUGHTS)),
    config,
    now,
    function(summary) {
      return generateDigestNarrative_(config, summary);
    }
  );
}

function buildDigestFromRows_(goalRows, todoRows, thoughtRows, config, now, narrativeGenerator) {
  const lookbackDays = parseInteger_(config.DIGEST_LOOKBACK_DAYS, 7);
  const since = addDays_(now, -lookbackDays);
  const goalUpdates = [];
  const todoUpdates = [];

  const activeGoals = normalizeRowsForValues_(goalRows)
    .map(function(goal, index) {
      return { values: goal, rowNumber: goalRows[index] && goalRows[index].rowNumber };
    })
    .filter(function(row) {
      const activeUntil = parseDate_(row.values.active_until);
      return !isDoneStatus_(row.values.status)
        && (!activeUntil || startOfDay_(activeUntil).getTime() >= startOfDay_(now).getTime());
    })
    .map(function(row) {
      const due = isReminderDue_(row.values, now);
      if (due) {
        goalUpdates.push(buildDigestStateUpdate_(row, now));
      }
      return copyObject_(row.values, { reminder_due: due });
    });

  const openTodos = normalizeRowsForValues_(todoRows)
    .map(function(todo, index) {
      return { values: todo, rowNumber: todoRows[index] && todoRows[index].rowNumber };
    })
    .filter(function(row) {
      const todo = row.values;
      const createdAt = parseDate_(todo.created_at);
      const dueDate = parseDate_(todo.due_date_optional);
      return !isDoneStatus_(todo.status)
        && (isRecent_(createdAt, since) || isDueSoon_(dueDate, now) || isReminderDue_(todo, now));
    })
    .map(function(row) {
      const due = isReminderDue_(row.values, now);
      if (due) {
        todoUpdates.push(buildDigestStateUpdate_(row, now));
      }
      return copyObject_(row.values, { reminder_due: due });
    });

  const recentThoughts = normalizeRowsForValues_(thoughtRows)
    .filter(function(thought) {
      return isRecent_(parseDate_(thought.created_at), since);
    });

  const summary = {
    goals: activeGoals,
    todos: openTodos,
    thoughts: recentThoughts,
    moodCounts: countMoods_(recentThoughts),
  };
  const narrative = buildDigestNarrative_(summary, narrativeGenerator);

  return {
    body: renderDigestTextBody_(activeGoals, openTodos, recentThoughts, narrative, now, lookbackDays),
    htmlBody: renderDigestHtmlBody_(activeGoals, openTodos, recentThoughts, narrative, now, lookbackDays),
    goalUpdates: goalUpdates,
    todoUpdates: todoUpdates,
  };
}

function renderDigestTextBody_(goals, todos, thoughts, narrative, now, lookbackDays) {
  const lines = [];
  lines.push('Voice Journal Digest');
  lines.push(formatDateForDigest_(now));
  lines.push('');
  if (narrative.intro) {
    lines.push(narrative.intro);
    lines.push('');
  }

  lines.push('__Goals__');
  if (narrative.goalsSummary) {
    lines.push(narrative.goalsSummary);
  }
  if (goals.length === 0) {
    lines.push('  1. No active goals.');
  } else {
    goals.slice(0, 12).forEach(function(goal, index) {
      const activeUntil = goal.active_until ? ' through ' + formatDateForDigest_(goal.active_until) : '';
      const dueNow = goal.reminder_due ? ' [reminder due]' : '';
      lines.push('  ' + (index + 1) + '. ' + goal.summary + activeUntil + dueNow);
    });
  }
  lines.push('');

  lines.push('__To-Dos__');
  if (narrative.todosSummary) {
    lines.push(narrative.todosSummary);
  }
  if (todos.length === 0) {
    lines.push('  1. No open to-dos needing attention.');
  } else {
    todos.slice(0, 15).forEach(function(todo, index) {
      const due = todo.due_date_optional ? ' due ' + formatDateForDigest_(todo.due_date_optional) : '';
      const dueNow = todo.reminder_due ? ' [reminder due]' : '';
      lines.push('  ' + (index + 1) + '. ' + todo.task + due + dueNow);
    });
  }
  lines.push('');

  lines.push('__Thoughts__');
  if (narrative.thoughtsSummary) {
    lines.push(narrative.thoughtsSummary);
  }
  if (thoughts.length === 0) {
    lines.push('  1. No recent thoughts from the last ' + lookbackDays + ' days.');
  } else {
    const moodCounts = countMoods_(thoughts);
    const moodSummary = Object.keys(moodCounts)
      .sort(function(a, b) { return moodCounts[b] - moodCounts[a]; })
      .map(function(mood) { return mood + ' (' + moodCounts[mood] + ')'; })
      .join(', ');
    if (moodSummary) {
      lines.push('Mood tags: ' + moodSummary);
    }
    thoughts.slice(0, 10).forEach(function(thought, index) {
      const moods = thought.moods ? ' [' + thought.moods + ']' : '';
      lines.push('  ' + (index + 1) + '. ' + thought.summary + moods);
    });
  }

  return lines.join('\n');
}

function renderDigestHtmlBody_(goals, todos, thoughts, narrative, now, lookbackDays) {
  return [
    '<div style="font-family:Arial,sans-serif;line-height:1.45;color:#202124;">',
    '<h2 style="margin:0 0 4px 0;">Voice Journal Digest</h2>',
    '<p style="margin:0 0 16px 0;color:#5f6368;">' + escapeHtml_(formatDateForDigest_(now)) + '</p>',
    narrative.intro ? '<p>' + escapeHtml_(narrative.intro) + '</p>' : '',
    renderDigestHtmlSection_('Goals', narrative.goalsSummary, goals.slice(0, 12), function(goal) {
      const activeUntil = goal.active_until ? ' through ' + formatDateForDigest_(goal.active_until) : '';
      const dueNow = goal.reminder_due ? ' [reminder due]' : '';
      return goal.summary + activeUntil + dueNow;
    }, 'No active goals.'),
    renderDigestHtmlSection_('To-Dos', narrative.todosSummary, todos.slice(0, 15), function(todo) {
      const due = todo.due_date_optional ? ' due ' + formatDateForDigest_(todo.due_date_optional) : '';
      const dueNow = todo.reminder_due ? ' [reminder due]' : '';
      return todo.task + due + dueNow;
    }, 'No open to-dos needing attention.'),
    renderDigestHtmlThoughtsSection_(thoughts.slice(0, 10), narrative.thoughtsSummary, lookbackDays),
    '</div>',
  ].join('');
}

function renderDigestHtmlSection_(title, summary, items, itemText, emptyText) {
  const itemHtml = items.length === 0
    ? '<li>' + escapeHtml_(emptyText) + '</li>'
    : items.map(function(item) { return '<li>' + escapeHtml_(itemText(item)) + '</li>'; }).join('');
  return [
    '<p style="margin:20px 0 4px 0;"><strong><u>' + escapeHtml_(title) + '</u></strong></p>',
    summary ? '<p style="margin:0 0 8px 0;">' + escapeHtml_(summary) + '</p>' : '',
    '<ol style="margin:0 0 0 24px;padding-left:18px;">',
    itemHtml,
    '</ol>',
  ].join('');
}

function renderDigestHtmlThoughtsSection_(thoughts, summary, lookbackDays) {
  const moodCounts = countMoods_(thoughts);
  const moodSummary = Object.keys(moodCounts)
    .sort(function(a, b) { return moodCounts[b] - moodCounts[a]; })
    .map(function(mood) { return mood + ' (' + moodCounts[mood] + ')'; })
    .join(', ');
  const items = thoughts.length === 0
    ? ['No recent thoughts from the last ' + lookbackDays + ' days.']
    : thoughts.map(function(thought) {
      const moods = thought.moods ? ' [' + thought.moods + ']' : '';
      return thought.summary + moods;
    });
  return [
    '<p style="margin:20px 0 4px 0;"><strong><u>Thoughts</u></strong></p>',
    summary ? '<p style="margin:0 0 8px 0;">' + escapeHtml_(summary) + '</p>' : '',
    moodSummary ? '<p style="margin:0 0 8px 0;">Mood tags: ' + escapeHtml_(moodSummary) + '</p>' : '',
    '<ol style="margin:0 0 0 24px;padding-left:18px;">',
    items.map(function(item) { return '<li>' + escapeHtml_(item) + '</li>'; }).join(''),
    '</ol>',
  ].join('');
}

function applyDigestStateUpdates_(spreadsheet, digest) {
  applyObjectRowUpdates_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS), digest.goalUpdates);
  applyObjectRowUpdates_(spreadsheet.getSheetByName(SHEET_NAMES.TODOS), digest.todoUpdates);
}

function applyObjectRowUpdates_(sheet, updates) {
  normalizeArray_(updates).forEach(function(update) {
    if (update.rowNumber) {
      writeObjectRow_(sheet, update.rowNumber, update.values);
    }
  });
}

function buildDigestStateUpdate_(row, now) {
  const values = row.values;
  const frequencyDays = positiveInteger_(values.reminder_frequency_days, 1);
  return {
    rowNumber: row.rowNumber,
    values: copyObject_(values, {
      reminder_count: parseInteger_(values.reminder_count, 0) + 1,
      last_reminded_at: now,
      next_reminder_at: addDays_(now, frequencyDays),
    }),
  };
}

function isReminderDue_(values, now) {
  const nextReminderAt = parseDate_(values.next_reminder_at);
  return !!nextReminderAt && nextReminderAt.getTime() <= now.getTime();
}

function countMoods_(thoughts) {
  const counts = {};
  thoughts.forEach(function(thought) {
    String(thought.moods || '').split(',').forEach(function(rawMood) {
      const mood = rawMood.trim().toLowerCase();
      if (mood) {
        counts[mood] = (counts[mood] || 0) + 1;
      }
    });
  });
  return counts;
}

function buildDigestNarrative_(summary, narrativeGenerator) {
  try {
    if (narrativeGenerator) {
      const narrative = sanitizeDigestNarrative_(narrativeGenerator(summary));
      if (narrative.intro || narrative.goalsSummary || narrative.todosSummary || narrative.thoughtsSummary) {
        return narrative;
      }
    }
  } catch (error) {
    // Digest delivery should not depend on optional AI narrative.
  }
  return fallbackDigestNarrative_(summary);
}

function generateDigestNarrative_(config, summary) {
  const apiKey = getOpenAiApiKey_(config);
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
    payload: JSON.stringify({
      model: config.EXTRACTION_MODEL || 'gpt-5.2',
      input: [
        {
          role: 'system',
          content: [
            'Write concise narrative copy for a personal voice journal email digest.',
            'Be accurate, realistic, and grounded in the provided journal content.',
            'Do not be overly warm, upbeat, optimistic, or motivational.',
            'It is okay to be compassionate when the content warrants it.',
            'Do not add bullets, markdown, advice, or claims not supported by the content.',
            'The intro must be one or two sentences.',
            'The goals and to-dos summaries must be one sentence each, two only if needed.',
            'The thoughts summary must be two sentences when possible and no more than three.',
          ].join(' '),
        },
        {
          role: 'user',
          content: digestNarrativePrompt_(summary),
        },
      ],
      text: {
        format: {
          type: 'json_schema',
          name: 'voice_journal_digest_narrative',
          strict: true,
          schema: digestNarrativeSchema_(),
        },
      },
    }),
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Digest intro failed with HTTP ' + code + ': ' + body);
  }

  const parsed = JSON.parse(body);
  const outputText = parsed.output_text || findOutputText_(parsed);
  if (!outputText) {
    throw new Error('Digest narrative response did not include output text.');
  }
  return JSON.parse(outputText);
}

function digestNarrativePrompt_(summary) {
  const goals = normalizeArray_(summary.goals).slice(0, 8).map(function(goal) { return goal.summary; }).filter(Boolean);
  const todos = normalizeArray_(summary.todos).slice(0, 10).map(function(todo) { return todo.task; }).filter(Boolean);
  const thoughts = normalizeArray_(summary.thoughts).slice(0, 8).map(function(thought) { return thought.summary; }).filter(Boolean);
  const moods = moodCountsToRows_(summary.moodCounts || {}).slice(0, 5).map(function(row) {
    return row[0] + ' (' + row[1] + ')';
  });

  return [
    'Active goals: ' + (goals.join('; ') || 'none'),
    'To-dos: ' + (todos.join('; ') || 'none'),
    'Recent reflections: ' + (thoughts.join('; ') || 'none'),
    'Mood counts: ' + (moods.join(', ') || 'none'),
  ].join('\n');
}

function digestNarrativeSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      intro: { type: 'string' },
      goalsSummary: { type: 'string' },
      todosSummary: { type: 'string' },
      thoughtsSummary: { type: 'string' },
    },
    required: ['intro', 'goalsSummary', 'todosSummary', 'thoughtsSummary'],
  };
}

function sanitizeDigestNarrative_(narrative) {
  if (typeof narrative === 'string') {
    return copyObject_(fallbackDigestNarrative_({}), { intro: sanitizeDigestText_(narrative, 500) });
  }
  return {
    intro: sanitizeDigestText_(narrative && narrative.intro, 500),
    goalsSummary: sanitizeDigestText_(narrative && narrative.goalsSummary, 350),
    todosSummary: sanitizeDigestText_(narrative && narrative.todosSummary, 350),
    thoughtsSummary: sanitizeDigestText_(narrative && narrative.thoughtsSummary, 600),
  };
}

function sanitizeDigestText_(value, maxLength) {
  const clean = String(value || '').replace(/\s+/g, ' ').trim();
  if (!clean) {
    return '';
  }
  return clean.length > maxLength ? clean.slice(0, maxLength - 3) + '...' : clean;
}

function fallbackDigestNarrative_(summary) {
  return {
    intro: fallbackDigestIntro_(summary),
    goalsSummary: fallbackSectionSummary_('goals', normalizeArray_(summary && summary.goals), summary),
    todosSummary: fallbackSectionSummary_('to-dos', normalizeArray_(summary && summary.todos), summary),
    thoughtsSummary: fallbackSectionSummary_('thoughts', normalizeArray_(summary && summary.thoughts), summary),
  };
}

function fallbackDigestIntro_(summary) {
  const goals = normalizeArray_(summary && summary.goals);
  const todos = normalizeArray_(summary && summary.todos);
  const thoughts = normalizeArray_(summary && summary.thoughts);
  const moodRows = moodCountsToRows_((summary && summary.moodCounts) || {});
  const leadingMoods = moodRows.slice(0, 2).map(function(row) { return row[0]; });
  const moodSentence = leadingMoods.length > 0
    ? ' The most common moods were ' + leadingMoods.join(' and ') + '.'
    : '';
  return 'Today\'s digest includes ' + goals.length + ' active goals, '
    + todos.length + ' to-dos, and ' + thoughts.length + ' recent reflections.'
    + moodSentence;
}

function fallbackSectionSummary_(section, items, summary) {
  if (section === 'goals') {
    const dueGoals = items.filter(function(goal) { return goal.reminder_due; }).length;
    return items.length === 0
      ? 'There are no active goals in this digest.'
      : items.length + ' active goals are included' + (dueGoals ? ', with ' + dueGoals + ' currently reminder-due.' : '.');
  }
  if (section === 'to-dos') {
    const dueTodos = items.filter(function(todo) { return todo.reminder_due; }).length;
    return items.length === 0
      ? 'There are no open to-dos needing attention in this digest.'
      : items.length + ' to-dos are included' + (dueTodos ? ', with ' + dueTodos + ' currently reminder-due.' : '.');
  }
  const moodRows = moodCountsToRows_((summary && summary.moodCounts) || {});
  if (items.length === 0) {
    return 'There are no recent thoughts in the current lookback window.';
  }
  if (moodRows.length === 0) {
    return items.length + ' recent thoughts are included.';
  }
  return items.length + ' recent thoughts are included. The most common mood tags are '
    + moodRows.slice(0, 2).map(function(row) { return row[0]; }).join(' and ') + '.';
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function deleteRowsWhereColumnEquals_(sheet, headerName, expectedValue) {
  const headers = getHeaders_(sheet);
  const columnIndex = headers.indexOf(headerName);
  if (columnIndex === -1 || sheet.getLastRow() < 2) {
    return;
  }

  const values = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === expectedValue) {
      sheet.deleteRow(i + 2);
    }
  }
}

function upsertEntry_(entriesSheet, existingEntry, values) {
  if (existingEntry) {
    writeObjectRow_(entriesSheet, existingEntry.rowNumber, values);
    return;
  }
  appendObjectRow_(entriesSheet, values);
}

function buildEntriesIndex_(entriesSheet) {
  const rows = getRowsAsObjects_(entriesSheet);
  const index = {};
  rows.forEach(function(row) {
    if (row.values.drive_file_id) {
      index[row.values.drive_file_id] = row;
    }
  });
  return index;
}

function shouldRetryEntry_(entry, maxRetries) {
  return entry.status === 'Failed' && parseInteger_(entry.retry_count, 0) < maxRetries;
}

function appendObjectRow_(sheet, values) {
  const headers = getHeaders_(sheet);
  const row = headers.map(function(header) {
    return values[header] === undefined ? '' : values[header];
  });
  sheet.appendRow(row);
}

function writeObjectRow_(sheet, rowNumber, values) {
  const headers = getHeaders_(sheet);
  const row = headers.map(function(header) {
    return values[header] === undefined ? '' : values[header];
  });
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
}

function getRowsAsObjects_(sheet) {
  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(function(row, index) {
    const object = {};
    headers.forEach(function(header, columnIndex) {
      object[header] = row[columnIndex];
    });
    return {
      rowNumber: index + 2,
      values: object,
    };
  });
}

function getHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return [];
  }
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function migrateKnownSchemaChanges_(spreadsheet) {
  const entriesSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ENTRIES);
  renameHeaderIfNeeded_(entriesSheet, 'recorded_at', 'uploaded_at');
  removeColumnsByHeader_(entriesSheet, ['recorded_at']);
  [SHEET_NAMES.GOALS, SHEET_NAMES.TODOS, SHEET_NAMES.THOUGHTS].forEach(function(sheetName) {
    const sheet = getOrCreateSheet_(spreadsheet, sheetName);
    removeColumnsByHeader_(sheet, ['review_status']);
  });
}

function renameHeaderIfNeeded_(sheet, oldHeader, newHeader) {
  const headers = getHeaders_(sheet);
  const oldIndex = headers.indexOf(oldHeader);
  const newIndex = headers.indexOf(newHeader);
  if (oldIndex !== -1 && newIndex === -1) {
    sheet.getRange(1, oldIndex + 1).setValue(newHeader);
  }
}

function removeColumnsByHeader_(sheet, headersToRemove) {
  const headers = getHeaders_(sheet);
  for (let i = headers.length - 1; i >= 0; i--) {
    if (headersToRemove.indexOf(headers[i]) !== -1) {
      sheet.deleteColumn(i + 1);
    }
  }
}

function ensureHeaders_(sheet, headers) {
  const existingHeaders = getHeaders_(sheet);
  if (existingHeaders.join('|') === headers.join('|')) {
    return;
  }

  if (existingHeaders.length === 0 || existingHeaders.every(function(header) { return !header; })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const missingHeaders = headers.filter(function(header) {
    return existingHeaders.indexOf(header) === -1;
  });

  if (missingHeaders.length > 0) {
    sheet.getRange(1, existingHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
}

function seedConfigDefaults_(configSheet) {
  const existing = readConfigRows_(configSheet);
  const existingKeys = {};
  existing.forEach(function(row) {
    existingKeys[row.key] = true;
  });

  DEFAULT_CONFIG.forEach(function(row) {
    if (!existingKeys[row[0]]) {
      configSheet.appendRow(row);
    }
  });
}

function readConfig_(spreadsheet) {
  const config = {};
  readConfigRows_(spreadsheet.getSheetByName(SHEET_NAMES.CONFIG)).forEach(function(row) {
    if (row.key) {
      config[row.key] = row.value;
    }
  });
  return config;
}

function readConfigRows_(configSheet) {
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = configSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return values.map(function(row) {
    return {
      key: row[0],
      value: row[1],
      notes: row[2],
    };
  });
}

function getOpenAiApiKey_(config) {
  const propertyName = config.OPENAI_API_KEY_PROPERTY || 'OPENAI_API_KEY';
  const apiKey = PropertiesService.getScriptProperties().getProperty(propertyName);
  if (!apiKey) {
    throw new Error('Missing OpenAI API key. Run setOpenAiApiKey("your-api-key") first.');
  }
  return apiKey;
}

function normalizeExtraction_(raw) {
  return {
    goals: normalizeArray_(raw.goals).map(function(goal) {
      return {
        summary: String(goal.summary || '').trim(),
        status: String(goal.status || '').trim(),
        active_until: String(goal.active_until || '').trim(),
        reminder_frequency_days: positiveInteger_(goal.reminder_frequency_days, 0),
      };
    }).filter(function(goal) { return goal.summary; }),
    todos: normalizeArray_(raw.todos).map(function(todo) {
      return {
        task: String(todo.task || '').trim(),
        status: String(todo.status || '').trim(),
        due_date_optional: String(todo.due_date_optional || '').trim(),
        reminder_frequency_days: positiveInteger_(todo.reminder_frequency_days, 0),
      };
    }).filter(function(todo) { return todo.task; }),
    thoughts: normalizeArray_(raw.thoughts).map(function(thought) {
      return {
        summary: String(thought.summary || '').trim(),
        moods: normalizeMoodTags_(thought.moods),
      };
    }).filter(function(thought) { return thought.summary; }),
  };
}

function emptyExtraction_() {
  return { goals: [], todos: [], thoughts: [] };
}

function normalizeArray_(value) {
  return Array.isArray(value) ? value : [];
}

function normalizeMoodTags_(value) {
  return normalizeArray_(value)
    .map(function(mood) { return String(mood || '').trim().toLowerCase(); })
    .filter(function(mood, index, moods) {
      return MOOD_TAGS.indexOf(mood) !== -1 && moods.indexOf(mood) === index;
    });
}

function findOutputText_(response) {
  const output = response.output || [];
  for (let i = 0; i < output.length; i++) {
    const item = output[i];
    const content = item.content || [];
    for (let j = 0; j < content.length; j++) {
      if (content[j].type === 'output_text' && content[j].text) {
        return content[j].text;
      }
    }
  }
  return '';
}

function extractionSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      goals: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            summary: { type: 'string' },
            status: { type: 'string' },
            active_until: { type: 'string' },
            reminder_frequency_days: { type: 'integer' },
          },
          required: ['summary', 'status', 'active_until', 'reminder_frequency_days'],
        },
      },
      todos: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            task: { type: 'string' },
            status: { type: 'string' },
            due_date_optional: { type: 'string' },
            reminder_frequency_days: { type: 'integer' },
          },
          required: ['task', 'status', 'due_date_optional', 'reminder_frequency_days'],
        },
      },
      thoughts: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            summary: { type: 'string' },
            moods: {
              type: 'array',
              items: { type: 'string', enum: MOOD_TAGS },
            },
          },
          required: ['summary', 'moods'],
        },
      },
    },
    required: ['goals', 'todos', 'thoughts'],
  };
}

function defaultExtractionPrompt_() {
  return [
    'You extract structured journal records from personal voice memo transcripts.',
    'Return only JSON matching the requested schema.',
    'Extract goals and to-dos from explicit markers and reasonable contextual clues.',
    'Always consider whether the entry contains one or more big ideas worth saving as thoughts, even without the word thought.',
    'For goals, infer active_until as an ISO date and reminder_frequency_days from context; if unclear, use 30 days from the memo date and 7 days.',
    'For to-dos, infer due_date_optional as an ISO date when possible and reminder_frequency_days from urgency; if unclear, use an empty due date and 3 days.',
    'Thought moods must come only from this list: ' + MOOD_TAGS.join(', ') + '.',
    'Do not invent unrelated tasks, goals, or dates; make useful conservative guesses only when the transcript implies intent.',
  ].join(' ');
}

function shouldSendDigestOnDate_(date, frequencyPerWeek) {
  const frequency = clampInteger_(parseInteger_(frequencyPerWeek, 7), 1, 7);
  if (frequency >= 7) {
    return true;
  }

  const day = date.getDay();
  if (frequency === 5) {
    return day >= 1 && day <= 5;
  }
  if (frequency === 3) {
    return day === 1 || day === 3 || day === 5;
  }
  if (frequency === 2) {
    return day === 1 || day === 4;
  }
  if (frequency === 1) {
    return day === 1;
  }

  const enabledDays = [1, 2, 3, 4, 5, 6, 0].slice(0, frequency);
  return enabledDays.indexOf(day) !== -1;
}

function isDoneStatus_(status) {
  return DONE_STATUSES.indexOf(String(status || '').trim().toLowerCase()) !== -1;
}

function isRecent_(date, since) {
  return !!date && date.getTime() >= since.getTime();
}

function isDueSoon_(date, now) {
  if (!date) {
    return false;
  }
  return startOfDay_(date).getTime() <= addDays_(startOfDay_(now), 7).getTime();
}

function parseDate_(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return value;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateForDigest_(value) {
  const date = parseDate_(value) || new Date();
  if (typeof Utilities === 'undefined' || typeof Session === 'undefined') {
    return date.toISOString().slice(0, 10);
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateTimeForDashboard_(value) {
  const date = parseDate_(value) || new Date();
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function startOfDay_(date) {
  const copy = new Date(date.getTime());
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function addDays_(date, days) {
  const copy = new Date(date.getTime());
  copy.setDate(copy.getDate() + days);
  return copy;
}

function copyObject_(source, overrides) {
  const copy = {};
  Object.keys(source || {}).forEach(function(key) {
    copy[key] = source[key];
  });
  Object.keys(overrides || {}).forEach(function(key) {
    copy[key] = overrides[key];
  });
  return copy;
}

function makeId_(prefix) {
  return prefix + '_' + Utilities.getUuid();
}

function positiveInteger_(value, fallback) {
  const parsed = parseInteger_(value, fallback);
  return parsed > 0 ? parsed : fallback;
}

function parseInteger_(value, fallback) {
  const parsed = parseInt(value, 10);
  return Number.isNaN(parsed) ? fallback : parsed;
}

function clampInteger_(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function errorToString_(error) {
  if (!error) {
    return 'Unknown error';
  }
  return error.stack || error.message || String(error);
}

if (typeof module !== 'undefined') {
  module.exports = {
    SHEET_NAMES,
    MOOD_TAGS,
    DONE_STATUSES,
    HEADERS,
    DEFAULT_CONFIG,
    normalizeExtraction_,
    normalizeMoodTags_,
    shouldRetryEntry_,
    shouldSendDigestOnDate_,
    isDoneStatus_,
    buildDigestFromRows_,
    fallbackDigestIntro_,
    buildDashboardTodoRows_,
    buildDashboardGoalRows_,
    buildRecentMoodCounts_,
    moodCountsToRows_,
    dashboardCheckboxStatus_,
    dashboardGoalActionStatus_,
    extractionSchema_,
    defaultExtractionPrompt_,
  };
}
