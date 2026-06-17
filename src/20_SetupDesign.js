/**
 * Setup Design
 */

function setupVoiceJournalSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  migrateKnownSchemaChanges_(spreadsheet);
  Object.keys(HEADERS).forEach(function(sheetName) {
    if (sheetName === SHEET_NAMES.SEARCH) {
      return;
    }
    const sheet = getOrCreateSheet_(spreadsheet, sheetName);
    ensureHeaders_(sheet, HEADERS[sheetName]);
    sheet.setFrozenRows(1);
  });
  const dashboardSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.DASHBOARD);
  moveSheetToFront_(spreadsheet, dashboardSheet);
  const configSheet = spreadsheet.getSheetByName(SHEET_NAMES.CONFIG);
  seedConfigDefaults_(configSheet);
  const designSheet = spreadsheet.getSheetByName(SHEET_NAMES.DESIGN);
  seedDesignDefaults_(designSheet);
  formatWorkbookSheets_(spreadsheet);
  renderSearchSheet_(spreadsheet.getSheetByName(SHEET_NAMES.SEARCH));
  hideSearchIndexSheet_(spreadsheet);
  renderConfigControls_(configSheet);
}

function moveSheetToFront_(spreadsheet, sheet) {
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.moveActiveSheet(1);
}

function formatWorkbookSheets_(spreadsheet) {
  formatSheetColumns_(spreadsheet.getSheetByName(SHEET_NAMES.ENTRIES), {
    drive_file_id: 180,
    audio_url: 260,
    uploaded_at: 150,
    processed_at: 150,
    transcript: 900,
    error: 420,
  }, ['transcript', 'error']);
  formatSheetColumns_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS), {
    summary: 460,
    created_at: 140,
    active_until: 130,
    next_reminder_at: 160,
    last_reminded_at: 160,
  }, ['summary']);
  formatSheetColumns_(spreadsheet.getSheetByName(SHEET_NAMES.FOLLOW_UPS), {
    summary: 460,
    created_at: 140,
    due_date_optional: 130,
    next_reminder_at: 160,
    last_reminded_at: 160,
  }, ['summary']);
  formatSheetColumns_(spreadsheet.getSheetByName(SHEET_NAMES.REFLECTIONS), {
    summary: 560,
    moods: 220,
    created_at: 140,
  }, ['summary', 'moods']);
  formatSheetColumns_(spreadsheet.getSheetByName(SHEET_NAMES.CONFIG), {
    key: 280,
    value: 420,
    notes: 560,
  }, ['value', 'notes']);
  formatSheetColumns_(spreadsheet.getSheetByName(SHEET_NAMES.DESIGN), {
    extraction_key: 160,
    display_name: 160,
    sheet_name: 160,
    description: 460,
    fact_table_dictionary: 520,
    keywords: 420,
    ai_decides_without_keyword: 190,
    status: 100,
    action: 130,
    backfill_cursor: 130,
    last_backfilled_at: 160,
    last_error: 360,
  }, ['description', 'fact_table_dictionary', 'keywords', 'last_error']);
  renderDesignControls_(spreadsheet.getSheetByName(SHEET_NAMES.DESIGN));
  formatSheetColumns_(spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX), {
    chunk_id: 220,
    entry_id: 220,
    quote_text: 420,
    context_before: 320,
    context_after: 320,
    related_reflections: 360,
    embedding_json: 220,
    index_error: 360,
  }, ['quote_text', 'context_before', 'context_after', 'related_reflections', 'embedding_json', 'index_error']);
}

function renderSearchSheet_(sheet) {
  if (!sheet) {
    return;
  }
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(SEARCH.RESULT_HEADER_ROW);
  sheet.getRange(1, 1, 1, SEARCH.RESULT_COLUMN_COUNT).clearContent();
  sheet.getRange('A1').setValue('Journal Search').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A2').setValue('Query').setFontWeight('bold');
  sheet.getRange('A3').setValue('Max results').setFontWeight('bold');
  sheet.getRange('A4').setValue('Last search').setFontWeight('bold');
  SEARCH_BUTTONS.forEach(function(button) {
    renderActionButtonFallback_(sheet, button);
  });
  removeActionButtonImages_(sheet);
  SEARCH_BUTTONS.forEach(function(button) {
    try {
      insertActionButtonImage_(sheet, button);
    } catch (error) {
      sheet.getRange(button.row, button.column).setNote(
        'Could not create the image button automatically. Use the Voice Journal menu for this action.'
      );
    }
  });
  sheet.getRange(SEARCH.RESULT_HEADER_ROW, 1, 1, SEARCH.RESULT_COLUMN_COUNT)
    .setValues([HEADERS.Search])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 360);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 260);
  sheet.setColumnWidth(8, 360);
  sheet.setColumnWidth(9, 180);
  sheet.getRange(1, 1, sheet.getMaxRows(), SEARCH.RESULT_COLUMN_COUNT).setVerticalAlignment('top');
  sheet.getRange(1, 3, sheet.getMaxRows(), 2).setWrap(true);
  sheet.getRange(1, 8, sheet.getMaxRows(), 2).setWrap(true);
}

function hideSearchIndexSheet_(spreadsheet) {
  const indexSheet = spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX);
  if (indexSheet && !indexSheet.isSheetHidden()) {
    indexSheet.hideSheet();
  }
}

function formatSheetColumns_(sheet, widthsByHeader, wrapHeaders) {
  if (!sheet) {
    return;
  }
  const headers = getHeaders_(sheet);
  headers.forEach(function(header, index) {
    const column = index + 1;
    if (widthsByHeader[header]) {
      sheet.setColumnWidth(column, widthsByHeader[header]);
    }
    if (wrapHeaders.indexOf(header) !== -1) {
      sheet.getRange(1, column, sheet.getMaxRows(), 1).setWrap(true).setVerticalAlignment('top');
    }
  });
  sheet.getRange(1, 1, sheet.getMaxRows(), Math.max(headers.length, 1)).setVerticalAlignment('top');
}

function renderConfigControls_(configSheet) {
  const titleRange = configSheet.getRange('E1:F1');
  if (!titleRange.isPartOfMerge()) {
    titleRange.merge();
  }
  titleRange.setValue('Quick Actions').setFontWeight('bold').setBackground('#f1f3f4');
  CONFIG_BUTTONS.forEach(function(button) {
    renderActionButtonFallback_(configSheet, button);
  });
  removeActionButtonImages_(configSheet);
  CONFIG_BUTTONS.forEach(function(button) {
    try {
      insertActionButtonImage_(configSheet, button);
    } catch (error) {
      configSheet.getRange(button.row, button.column).setNote(
        'Could not create the image button automatically. Use the Voice Journal menu for this action.'
      );
    }
  });
}

function renderDesignControls_(designSheet) {
  if (!designSheet) {
    return;
  }
  designSheet.setFrozenRows(1);
  const actionValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(DESIGN.ACTIONS, true)
    .setAllowInvalid(false)
    .build();
  designSheet.getRange(DESIGN.FIRST_DATA_ROW, DESIGN.ACTION_COLUMN, designSheet.getMaxRows() - 1, 1)
    .setDataValidation(actionValidation);
  const booleanValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  designSheet.getRange(DESIGN.FIRST_DATA_ROW, 7, designSheet.getMaxRows() - 1, 1)
    .setDataValidation(booleanValidation);
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Paused'], true)
    .setAllowInvalid(false)
    .build();
  designSheet.getRange(DESIGN.FIRST_DATA_ROW, 8, designSheet.getMaxRows() - 1, 1)
    .setDataValidation(statusValidation);
}

function renderActionButtonFallback_(sheet, button) {
  const range = sheet.getRange(button.row, button.column, 2, 2);
  if (!range.isPartOfMerge()) {
    range.merge();
  }
  range
    .setValue(button.title)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground(button.color)
    .setNote('If the image button does not appear, use the Voice Journal menu for this action.');
  sheet.setColumnWidth(button.column, 150);
  sheet.setColumnWidth(button.column + 1, 150);
  sheet.setRowHeight(button.row, 34);
  sheet.setRowHeight(button.row + 1, 10);
}

function removeActionButtonImages_(sheet) {
  sheet.getImages().forEach(function(image) {
    const title = image.getAltTextTitle && image.getAltTextTitle();
    if (String(title || '').indexOf('Voice Journal:') === 0) {
      image.remove();
    }
  });
}

function insertActionButtonImage_(sheet, button) {
  const transparentPixel = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=';
  const blob = Utilities.newBlob(Utilities.base64Decode(transparentPixel), 'image/png', button.title + '.png');
  const image = sheet.insertImage(blob, button.column, button.row);
  image
    .setAltTextTitle('Voice Journal: ' + button.title)
    .setAltTextDescription('Runs ' + button.script)
    .assignScript(button.script)
    .setWidth(220)
    .setHeight(44);
}

function runSelectedDesignAction() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  if (!sheet || sheet.getName() !== SHEET_NAMES.DESIGN) {
    throw new Error('Select a row in the Design tab before running a design action.');
  }
  const row = sheet.getActiveRange().getRow();
  if (row < DESIGN.FIRST_DATA_ROW) {
    throw new Error('Select a design row, not the header row.');
  }
  return processDesignActionRow_(spreadsheet, row, 'Run Backfill');
}

function handleDesignActionEdit_(spreadsheet, sheet, e) {
  if (e.range.getColumn() !== DESIGN.ACTION_COLUMN || e.range.getRow() < DESIGN.FIRST_DATA_ROW) {
    return false;
  }
  const action = String(e.value || '').trim();
  if (!action) {
    return false;
  }
  processDesignActionRow_(spreadsheet, e.range.getRow(), action);
  sheet.getRange(e.range.getRow(), DESIGN.ACTION_COLUMN).clearContent();
  return true;
}

function processDesignActionRow_(spreadsheet, rowNumber, action) {
  const designSheet = spreadsheet.getSheetByName(SHEET_NAMES.DESIGN);
  const designRow = getRowsAsObjects_(designSheet).filter(function(row) {
    return row.rowNumber === rowNumber;
  })[0];
  if (!designRow) {
    throw new Error('No design row found at row ' + rowNumber + '.');
  }

  const design = normalizeDesignRow_(designRow.values);
  if (!design.extraction_key) {
    throw new Error('Design row ' + rowNumber + ' is missing extraction_key.');
  }

  try {
    if (!isBuiltInExtraction_(design)) {
      ensureCustomExtractionSheet_(spreadsheet, design);
      if (String(action || '').trim() === 'Run Backfill') {
        return backfillCustomExtraction_(spreadsheet, design, rowNumber);
      }
    }
    updateDesignRowState_(designSheet, rowNumber, {
      action: '',
      last_error: '',
      last_backfilled_at: new Date(),
    });
    spreadsheet.toast('Design action completed for ' + design.display_name + '.', 'Voice Journal', 5);
    return 'Completed design action for ' + design.display_name + '.';
  } catch (error) {
    updateDesignRowState_(designSheet, rowNumber, {
      action: '',
      last_error: errorToString_(error),
    });
    throw error;
  }
}

function ensureCustomExtractionSheet_(spreadsheet, design) {
  const normalized = normalizeDesignRow_(design);
  if (isBuiltInExtraction_(normalized)) {
    return spreadsheet.getSheetByName(normalized.sheet_name);
  }
  const sheet = getOrCreateSheet_(spreadsheet, normalized.sheet_name);
  const factFields = customFactFields_(normalized.fact_table_dictionary, normalized.extraction_key);
  ensureHeaders_(sheet, customExtractionHeaders_(normalized.extraction_key, normalized.fact_table_dictionary));
  sheet.setFrozenRows(1);
  const widths = {
    entry_id: 220,
    matched_keywords: 260,
    created_at: 150,
    extracted_at: 150,
  };
  const wrapHeaders = ['matched_keywords'];
  factFields.forEach(function(field) {
    widths[field.key] = field.key === 'evidence' ? 520 : 360;
    if (field.type === 'string' || field.type === 'enum') {
      wrapHeaders.push(field.key);
    }
  });
  formatSheetColumns_(sheet, widths, wrapHeaders);
  return sheet;
}

function backfillCustomExtraction_(spreadsheet, design, rowNumber) {
  const config = readConfig_(spreadsheet);
  const entries = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.ENTRIES))
    .filter(function(row) {
      return row.values.entry_id
        && row.values.transcript
        && String(row.values.status || '').toLowerCase() === 'processed';
    });
  const startIndex = clampInteger_(parseInteger_(design.backfill_cursor, 0), 0, entries.length);
  const endIndex = Math.min(entries.length, startIndex + DESIGN.BACKFILL_BATCH_SIZE);
  const now = new Date();
  let processedCount = 0;

  ensureCustomExtractionSheet_(spreadsheet, design);
  for (let i = startIndex; i < endIndex; i++) {
    const entry = entries[i].values;
    const sheet = spreadsheet.getSheetByName(design.sheet_name);
    deleteRowsWhereColumnEquals_(sheet, 'entry_id', entry.entry_id);
    const records = extractCustomJournalItems_(entry.transcript, entry.entry_id, design, config);
    appendCustomExtractionRecords_(
      spreadsheet,
      design,
      entry.entry_id,
      records,
      parseDate_(entry.uploaded_at) || now,
      now
    );
    processedCount += 1;
  }

  const nextCursor = endIndex >= entries.length ? '' : String(endIndex);
  updateDesignRowState_(spreadsheet.getSheetByName(SHEET_NAMES.DESIGN), rowNumber, {
    action: '',
    backfill_cursor: nextCursor,
    last_backfilled_at: now,
    last_error: '',
  });
  const message = nextCursor
    ? 'Backfilled ' + processedCount + ' entries for ' + design.display_name + '. Run again to continue.'
    : 'Backfilled ' + processedCount + ' entries for ' + design.display_name + '.';
  spreadsheet.toast(message, 'Voice Journal', 5);
  return message;
}

function updateDesignRowState_(designSheet, rowNumber, values) {
  const headers = getHeaders_(designSheet);
  Object.keys(values).forEach(function(key) {
    const index = headers.indexOf(key);
    if (index !== -1) {
      designSheet.getRange(rowNumber, index + 1).setValue(values[key]);
    }
  });
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function migrateKnownSchemaChanges_(spreadsheet) {
  renameSheetIfNeeded_(spreadsheet, 'To-Dos', SHEET_NAMES.FOLLOW_UPS);
  renameSheetIfNeeded_(spreadsheet, 'Thoughts', SHEET_NAMES.REFLECTIONS);
  const entriesSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ENTRIES);
  renameHeaderIfNeeded_(entriesSheet, 'recorded_at', 'uploaded_at');
  renameHeaderIfNeeded_(entriesSheet, 'has_todo', 'has_follow_up');
  renameHeaderIfNeeded_(entriesSheet, 'has_thought', 'has_reflection');
  removeColumnsByHeader_(entriesSheet, ['recorded_at']);
  const followUpsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.FOLLOW_UPS);
  renameHeaderIfNeeded_(followUpsSheet, 'todo_id', 'follow_up_id');
  renameHeaderIfNeeded_(followUpsSheet, 'task', 'summary');
  const reflectionsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.REFLECTIONS);
  renameHeaderIfNeeded_(reflectionsSheet, 'thought_id', 'reflection_id');
  const searchIndexSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.SEARCH_INDEX);
  renameHeaderIfNeeded_(searchIndexSheet, 'related_thoughts', 'related_reflections');
  const configSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.CONFIG);
  renameConfigKeyIfNeeded_(configSheet, 'DEFAULT_TODO_STATUS', 'DEFAULT_FOLLOW_UP_STATUS');
  renameConfigKeyIfNeeded_(configSheet, 'DEFAULT_TODO_REMINDER_FREQUENCY_DAYS', 'DEFAULT_FOLLOW_UP_REMINDER_FREQUENCY_DAYS');
  [SHEET_NAMES.GOALS, SHEET_NAMES.FOLLOW_UPS, SHEET_NAMES.REFLECTIONS].forEach(function(sheetName) {
    removeColumnsByHeader_(getOrCreateSheet_(spreadsheet, sheetName), ['review_status']);
  });
}

function renameSheetIfNeeded_(spreadsheet, oldName, newName) {
  const oldSheet = spreadsheet.getSheetByName(oldName);
  const newSheet = spreadsheet.getSheetByName(newName);
  if (oldSheet && !newSheet) {
    oldSheet.setName(newName);
  }
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

function renameConfigKeyIfNeeded_(configSheet, oldKey, newKey) {
  const rows = readConfigRows_(configSheet);
  const hasNewKey = rows.some(function(row) {
    return row.key === newKey;
  });
  if (hasNewKey) {
    return;
  }
  const oldRow = rows.find(function(row) {
    return row.key === oldKey;
  });
  if (!oldRow) {
    return;
  }
  const keyColumn = 1;
  configSheet.getRange(rows.indexOf(oldRow) + 2, keyColumn).setValue(newKey);
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

function seedDesignDefaults_(designSheet) {
  const existing = readDesignRows_(designSheet);
  const existingKeys = {};
  existing.forEach(function(row) {
    if (row.values.extraction_key) {
      existingKeys[row.values.extraction_key] = true;
    }
  });

  DEFAULT_DESIGN_ROWS.forEach(function(row) {
    if (!existingKeys[row.extraction_key]) {
      appendObjectRow_(designSheet, copyObject_(row, {
        action: '',
        backfill_cursor: '',
        last_backfilled_at: '',
        last_error: '',
      }));
    }
  });
}

function readDesignRows_(designSheet) {
  return getRowsAsObjects_(designSheet).map(function(row) {
    return {
      rowNumber: row.rowNumber,
      values: normalizeDesignRow_(row.values),
    };
  });
}

function normalizeDesignRow_(row) {
  const key = normalizeExtractionKey_(row && row.extraction_key);
  const builtInSheetName = builtInSheetNameForExtractionKey_(key);
  return {
    extraction_key: key,
    display_name: String((row && row.display_name) || '').trim() || titleFromExtractionKey_(key),
    sheet_name: builtInSheetName || sanitizeSheetName_((row && row.sheet_name) || titleFromExtractionKey_(key)),
    description: String((row && row.description) || '').trim(),
    fact_table_dictionary: String((row && row.fact_table_dictionary) || '').trim(),
    keywords: parseKeywords_(row && row.keywords).join(', '),
    ai_decides_without_keyword: parseBoolean_(row && row.ai_decides_without_keyword),
    status: String((row && row.status) || DESIGN.ACTIVE_STATUS).trim() || DESIGN.ACTIVE_STATUS,
    action: String((row && row.action) || '').trim(),
    backfill_cursor: String((row && row.backfill_cursor) || '').trim(),
    last_backfilled_at: row && row.last_backfilled_at ? row.last_backfilled_at : '',
    last_error: String((row && row.last_error) || '').trim(),
  };
}

function normalizeExtractionKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function titleFromExtractionKey_(key) {
  return String(key || 'Custom Extraction')
    .split('_')
    .filter(Boolean)
    .map(function(part) { return part.charAt(0).toUpperCase() + part.slice(1); })
    .join(' ');
}

function sanitizeSheetName_(value) {
  const clean = String(value || '').replace(/[\[\]\*\/\\\?:]/g, ' ').replace(/\s+/g, ' ').trim();
  return (clean || 'Custom Extraction').slice(0, 99);
}

function parseKeywords_(value) {
  return String(value || '')
    .split(/[,;\n]/)
    .map(function(keyword) { return keyword.trim().toLowerCase(); })
    .filter(function(keyword, index, keywords) {
      return keyword && keywords.indexOf(keyword) === index;
    });
}

function parseBoolean_(value) {
  if (value === true) {
    return true;
  }
  const normalized = String(value || '').trim().toLowerCase();
  return normalized === 'true' || normalized === 'yes' || normalized === 'y' || normalized === '1';
}

function isBuiltInExtraction_(design) {
  return DESIGN.BUILT_IN_KEYS.indexOf(design.extraction_key) !== -1;
}

function builtInSheetNameForExtractionKey_(key) {
  if (key === 'follow_ups') {
    return SHEET_NAMES.FOLLOW_UPS;
  }
  if (key === 'goals') {
    return SHEET_NAMES.GOALS;
  }
  if (key === 'reflections') {
    return SHEET_NAMES.REFLECTIONS;
  }
  return '';
}

function customExtractionHeaders_(extractionKey, dictionary) {
  const key = normalizeExtractionKey_(extractionKey);
  const customHeaders = customFactFields_(dictionary, key).map(function(field) {
    return field.key;
  });
  return [
    key + '_id',
    'entry_id',
    'matched_keywords',
    'created_at',
    'extracted_at',
  ].concat(customHeaders);
}

function customFactFields_(dictionary, extractionKey) {
  return parseFactTableDictionary_(dictionary, extractionKey).fields;
}

function parseFactTableDictionary_(dictionary, extractionKey) {
  const raw = String(dictionary || '').trim();
  if (!raw) {
    return { fields: defaultCustomFactFields_() };
  }
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error('fact_table_dictionary must be valid JSON: ' + error.message);
  }
  const fields = normalizeArray_(parsed && parsed.fields);
  if (fields.length === 0) {
    throw new Error('fact_table_dictionary.fields must include at least one field.');
  }
  const normalizedFields = fields.map(function(field) {
    return normalizeFactField_(field, extractionKey);
  });
  const seen = {};
  normalizedFields.forEach(function(field) {
    if (seen[field.key]) {
      throw new Error('fact_table_dictionary field key is duplicated: ' + field.key);
    }
    seen[field.key] = true;
  });
  return { fields: normalizedFields };
}

function defaultCustomFactFields_() {
  return [
    { key: 'summary', type: 'string', description: 'Short extracted fact summary.' },
    { key: 'evidence', type: 'string', description: 'Short exact or near-exact transcript excerpt supporting the summary.' },
    { key: 'status', type: 'string', description: 'Concise review status. Use Open unless the transcript clearly implies another status.' },
  ];
}

function normalizeFactField_(field, extractionKey) {
  const key = String(field && field.key || '').trim();
  const type = String(field && field.type || '').trim().toLowerCase();
  const description = String(field && field.description || '').trim();
  const values = normalizeArray_(field && field.values)
    .map(function(value) { return String(value || '').trim(); })
    .filter(Boolean);

  if (!/^[a-z][a-z0-9]*(?:_[a-z0-9]+)*$/.test(key)) {
    throw new Error('fact_table_dictionary field key must be snake_case: ' + key);
  }
  if (reservedCustomFactColumns_(extractionKey).indexOf(key) !== -1) {
    throw new Error('fact_table_dictionary field key is reserved: ' + key);
  }
  if (['string', 'number', 'boolean', 'date', 'enum'].indexOf(type) === -1) {
    throw new Error('fact_table_dictionary field type is unsupported for ' + key + ': ' + type);
  }
  if (type === 'enum' && values.length === 0) {
    throw new Error('fact_table_dictionary enum field must define values: ' + key);
  }
  return {
    key: key,
    type: type,
    description: description,
    values: values,
  };
}

function reservedCustomFactColumns_(extractionKey) {
  return [
    normalizeExtractionKey_(extractionKey) + '_id',
    'entry_id',
    'matched_keywords',
    'created_at',
    'extracted_at',
  ];
}

function findMatchedKeywords_(transcript, keywords) {
  const text = String(transcript || '').toLowerCase();
  return parseKeywords_(keywords).filter(function(keyword) {
    return text.indexOf(keyword) !== -1;
  });
}

function shouldRunCustomExtraction_(design, transcript) {
  const normalized = normalizeDesignRow_(design);
  const matchedKeywords = findMatchedKeywords_(transcript, normalized.keywords);
  return normalized.ai_decides_without_keyword || matchedKeywords.length > 0;
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
