/**
 * Utilities
 */

function normalizeRowsForValues_(rows) {
  return normalizeArray_(rows).map(function(row) {
    return row && row.values ? row.values : row;
  });
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function clearSheetRows_(sheet, startRow) {
  const row = startRow || 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < row) {
    return;
  }
  sheet.getRange(row, 1, lastRow - row + 1, Math.max(sheet.getLastColumn(), 1)).clearContent();
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

function getOpenAiApiKey_(config) {
  const propertyName = getOpenAiApiKeyPropertyName_(config);
  const apiKey = PropertiesService.getScriptProperties().getProperty(propertyName);
  if (!apiKey) {
    throw new Error('Missing OpenAI API key. Fill OPENAI_API_KEY_SETUP in Config and run installVoiceJournal(), or run setOpenAiApiKey("your-api-key") first.');
  }
  return apiKey;
}

function normalizeArray_(value) {
  return Array.isArray(value) ? value : [];
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
  if (typeof Utilities === 'undefined' || typeof Session === 'undefined') {
    return date.toISOString().slice(0, 16).replace('T', ' ');
  }
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
