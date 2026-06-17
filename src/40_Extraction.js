/**
 * Extraction
 */

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

function appendExtractedRows_(spreadsheet, entryId, extraction, config, sourceTimestamp) {
  const now = sourceTimestamp || new Date();
  const followUpStatus = config.DEFAULT_FOLLOW_UP_STATUS || 'Open';
  const goalStatus = config.DEFAULT_GOAL_STATUS || 'Active';
  const defaultGoalDays = parseInteger_(config.DEFAULT_GOAL_ACTIVE_DAYS, 30);
  const defaultGoalFrequency = parseInteger_(config.DEFAULT_GOAL_REMINDER_FREQUENCY_DAYS, 7);
  const defaultFollowUpFrequency = parseInteger_(config.DEFAULT_FOLLOW_UP_REMINDER_FREQUENCY_DAYS, 3);

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

  const followUpsSheet = spreadsheet.getSheetByName(SHEET_NAMES.FOLLOW_UPS);
  extraction.followUps.forEach(function(followUp) {
    const followUpId = makeId_('followUp');
    const dueDate = parseDate_(followUp.due_date_optional);
    const frequencyDays = positiveInteger_(followUp.reminder_frequency_days, defaultFollowUpFrequency);
    appendObjectRow_(followUpsSheet, {
      follow_up_id: followUpId,
      entry_id: entryId,
      summary: followUp.summary,
      status: followUp.status || followUpStatus,
      created_at: now,
      due_date_optional: dueDate || followUp.due_date_optional || '',
      reminder_frequency_days: frequencyDays,
      next_reminder_at: now,
      last_reminded_at: '',
      reminder_count: 0,
    });
  });

  const reflectionsSheet = spreadsheet.getSheetByName(SHEET_NAMES.REFLECTIONS);
  extraction.reflections.forEach(function(reflection) {
    appendObjectRow_(reflectionsSheet, {
      reflection_id: makeId_('reflection'),
      entry_id: entryId,
      summary: reflection.summary,
      moods: reflection.moods.join(', '),
      created_at: now,
    });
  });
}

function appendCustomExtractionsForEntry_(spreadsheet, entryId, transcript, config, sourceTimestamp) {
  const designRows = readDesignRows_(spreadsheet.getSheetByName(SHEET_NAMES.DESIGN));
  designRows.forEach(function(row) {
    const design = row.values;
    if (!design.extraction_key || isBuiltInExtraction_(design) || String(design.status).toLowerCase() !== 'active') {
      return;
    }
    ensureCustomExtractionSheet_(spreadsheet, design);
    deleteRowsWhereColumnEquals_(spreadsheet.getSheetByName(design.sheet_name), 'entry_id', entryId);
    const records = extractCustomJournalItems_(transcript, entryId, design, config);
    appendCustomExtractionRecords_(spreadsheet, design, entryId, records, sourceTimestamp || new Date(), new Date());
  });
}

function extractCustomJournalItems_(transcript, entryId, design, config) {
  const matchedKeywords = findMatchedKeywords_(transcript, design.keywords);
  if (!design.ai_decides_without_keyword && matchedKeywords.length === 0) {
    return [];
  }
  const factFields = customFactFields_(design.fact_table_dictionary, design.extraction_key);

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
          content: customExtractionPrompt_(design, matchedKeywords, factFields),
        },
        {
          role: 'user',
          content: 'Entry ID: ' + entryId + '\n\nTranscript:\n' + transcript,
        },
      ],
      text: {
        format: {
          type: 'json_schema',
          name: 'voice_journal_custom_extraction',
          strict: true,
          schema: customExtractionSchema_(factFields),
        },
      },
    }),
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Custom extraction failed with HTTP ' + code + ': ' + body);
  }
  const parsed = JSON.parse(body);
  const outputText = parsed.output_text || findOutputText_(parsed);
  if (!outputText) {
    throw new Error('Custom extraction response did not include output text.');
  }
  return normalizeCustomExtraction_(JSON.parse(outputText), matchedKeywords, factFields);
}

function customExtractionPrompt_(design, matchedKeywords, factFields) {
  const fields = normalizeArray_(factFields).length > 0
    ? factFields
    : customFactFields_(design.fact_table_dictionary, design.extraction_key);
  const fieldLines = fields.map(function(field) {
    const enumValues = field.type === 'enum' ? ' Allowed values: ' + field.values.join(', ') + '.' : '';
    return '- ' + field.key + ' (' + field.type + '): ' + (field.description || 'No description provided.') + enumValues;
  });
  return [
    'You extract one configured kind of structured journal record from personal voice memo transcripts.',
    'Return only JSON matching the requested schema.',
    'Extraction name: ' + design.display_name + '.',
    'Extraction description: ' + (design.description || 'No extra description provided.') + '.',
    'Configured keywords: ' + (design.keywords || 'none') + '.',
    'Matched keywords in this transcript: ' + (matchedKeywords.join(', ') || 'none') + '.',
    'Return every configured field for every record. If a value is unknown, use a blank string.',
    'Configured fields:\n' + fieldLines.join('\n'),
    'Create records only when the transcript contains a clear instance of this extraction type.',
  ].join(' ');
}

function customExtractionSchema_(factFields) {
  const fields = normalizeArray_(factFields).length > 0 ? factFields : defaultCustomFactFields_();
  const properties = {};
  const required = [];
  fields.forEach(function(field) {
    properties[field.key] = customFactFieldSchema_(field);
    required.push(field.key);
  });
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      records: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: properties,
          required: required,
        },
      },
    },
    required: ['records'],
  };
}

function customFactFieldSchema_(field) {
  if (field.type === 'number') {
    return { type: ['number', 'string'] };
  }
  if (field.type === 'boolean') {
    return { type: ['boolean', 'string'] };
  }
  if (field.type === 'enum') {
    return { type: 'string', enum: [''].concat(field.values) };
  }
  return { type: 'string' };
}

function normalizeCustomExtraction_(raw, matchedKeywords, factFields) {
  const fields = normalizeArray_(factFields).length > 0 ? factFields : defaultCustomFactFields_();
  return normalizeArray_(raw && raw.records).map(function(record) {
    const normalized = {
      matched_keywords: parseKeywords_(matchedKeywords).join(', '),
    };
    fields.forEach(function(field) {
      normalized[field.key] = normalizeCustomFactValue_(record && record[field.key], field);
    });
    return normalized;
  }).filter(function(record) {
    return fields.some(function(field) {
      return String(record[field.key] || '').trim();
    });
  });
}

function normalizeCustomFactValue_(value, field) {
  if (value === null || value === undefined) {
    return '';
  }
  if (field.type === 'number') {
    return value === '' ? '' : Number(value);
  }
  if (field.type === 'boolean') {
    if (value === true || value === false) {
      return value;
    }
    const normalized = String(value || '').trim().toLowerCase();
    if (normalized === 'true' || normalized === 'yes' || normalized === '1') {
      return true;
    }
    if (normalized === 'false' || normalized === 'no' || normalized === '0') {
      return false;
    }
    return '';
  }
  if (field.type === 'enum') {
    const enumValue = String(value || '').trim();
    return field.values.indexOf(enumValue) === -1 ? '' : enumValue;
  }
  return String(value || '').trim();
}

function appendCustomExtractionRecords_(spreadsheet, design, entryId, records, createdAt, extractedAt) {
  const sheet = spreadsheet.getSheetByName(design.sheet_name);
  const idHeader = normalizeExtractionKey_(design.extraction_key) + '_id';
  const factFields = customFactFields_(design.fact_table_dictionary, design.extraction_key);
  normalizeArray_(records).forEach(function(record) {
    const values = {};
    values[idHeader] = makeId_(normalizeExtractionKey_(design.extraction_key));
    values.entry_id = entryId;
    values.matched_keywords = record.matched_keywords || '';
    values.created_at = createdAt;
    values.extracted_at = extractedAt;
    factFields.forEach(function(field) {
      values[field.key] = record[field.key] === undefined ? '' : record[field.key];
    });
    appendObjectRow_(sheet, values);
  });
}

function removeDerivedRowsForEntry_(spreadsheet, entryId) {
  [SHEET_NAMES.GOALS, SHEET_NAMES.FOLLOW_UPS, SHEET_NAMES.REFLECTIONS].forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    deleteRowsWhereColumnEquals_(sheet, 'entry_id', entryId);
  });
  readDesignRows_(spreadsheet.getSheetByName(SHEET_NAMES.DESIGN)).forEach(function(row) {
    const design = row.values;
    if (!design.extraction_key || isBuiltInExtraction_(design)) {
      return;
    }
    const sheet = spreadsheet.getSheetByName(design.sheet_name);
    if (sheet) {
      deleteRowsWhereColumnEquals_(sheet, 'entry_id', entryId);
    }
  });
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
    followUps: normalizeArray_(raw.followUps).map(function(followUp) {
      return {
        summary: String(followUp.summary || '').trim(),
        status: String(followUp.status || '').trim(),
        due_date_optional: String(followUp.due_date_optional || '').trim(),
        reminder_frequency_days: positiveInteger_(followUp.reminder_frequency_days, 0),
      };
    }).filter(function(followUp) { return followUp.summary; }),
    reflections: normalizeArray_(raw.reflections).map(function(reflection) {
      return {
        summary: String(reflection.summary || '').trim(),
        moods: normalizeMoodTags_(reflection.moods),
      };
    }).filter(function(reflection) { return reflection.summary; }),
  };
}

function emptyExtraction_() {
  return { goals: [], followUps: [], reflections: [] };
}

function normalizeMoodTags_(value) {
  return normalizeArray_(value)
    .map(function(mood) { return String(mood || '').trim().toLowerCase(); })
    .filter(function(mood, index, moods) {
      return MOOD_TAGS.indexOf(mood) !== -1 && moods.indexOf(mood) === index;
    });
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
      followUps: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            summary: { type: 'string' },
            status: { type: 'string' },
            due_date_optional: { type: 'string' },
            reminder_frequency_days: { type: 'integer' },
          },
          required: ['summary', 'status', 'due_date_optional', 'reminder_frequency_days'],
        },
      },
      reflections: {
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
    required: ['goals', 'followUps', 'reflections'],
  };
}

function defaultExtractionPrompt_() {
  return DEFAULT_EXTRACTION_PROMPT_TEXT;
}
