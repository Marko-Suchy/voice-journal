/**
 * Search
 */

function searchJournalFromMenu() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const message = searchJournal_(spreadsheet, readConfig_(spreadsheet), new Date());
    spreadsheet.toast(message, 'Voice Journal', 5);
    return message;
  } catch (error) {
    spreadsheet.toast('Search failed: ' + errorToString_(error), 'Voice Journal', 8);
    throw error;
  }
}

function refreshSearchIndexFromMenu() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const count = refreshSearchIndex_(spreadsheet, readConfig_(spreadsheet), false);
    spreadsheet.toast('Refreshed ' + count + ' search index entries.', 'Voice Journal', 5);
    return count;
  } catch (error) {
    spreadsheet.toast('Search index refresh failed: ' + errorToString_(error), 'Voice Journal', 8);
    throw error;
  }
}

function rebuildSearchIndexFromMenu() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const count = refreshSearchIndex_(spreadsheet, readConfig_(spreadsheet), true);
    spreadsheet.toast('Rebuilt search index with ' + count + ' entries.', 'Voice Journal', 5);
    return count;
  } catch (error) {
    spreadsheet.toast('Search index rebuild failed: ' + errorToString_(error), 'Voice Journal', 8);
    throw error;
  }
}

function getSearchAppState() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return buildSearchAppStateFromRows_(
    getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX)),
    readConfig_(spreadsheet)
  );
}

function searchJournalWeb(query, options) {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(spreadsheet);
  const cleanQuery = validateSearchQuery_(query);
  const searchOptions = normalizeSearchOptions_(options, config);
  const queryEmbedding = embedText_(cleanQuery, config);
  const indexRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX));
  const ranked = rankSearchRows_(indexRows, queryEmbedding, searchOptions.minScore, searchOptions.maxResults);
  return {
    query: cleanQuery,
    options: searchOptions,
    results: buildSearchResultObjects_(ranked),
    state: buildSearchAppStateFromRows_(indexRows, config),
  };
}

function refreshSearchIndexWeb() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(spreadsheet);
  const changedEntries = refreshSearchIndex_(spreadsheet, config, false);
  return {
    changedEntries: changedEntries,
    state: buildSearchAppStateFromRows_(
      getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX)),
      config
    ),
  };
}

function rebuildSearchIndexWeb() {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(spreadsheet);
  const changedEntries = refreshSearchIndex_(spreadsheet, config, true);
  return {
    changedEntries: changedEntries,
    state: buildSearchAppStateFromRows_(
      getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX)),
      config
    ),
  };
}

function searchJournal_(spreadsheet, config, now) {
  const searchSheet = spreadsheet.getSheetByName(SHEET_NAMES.SEARCH);
  const query = String(searchSheet.getRange(SEARCH.QUERY_CELL).getValue() || '').trim();
  if (!query) {
    throw new Error('Enter a query in Search!' + SEARCH.QUERY_CELL + ' before searching.');
  }

  const queryEmbedding = embedText_(query, config);
  const indexRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX));
  const maxResultsOverride = parseInteger_(searchSheet.getRange(SEARCH.MAX_RESULTS_CELL).getValue(), 0);
  const maxResults = maxResultsOverride > 0
    ? maxResultsOverride
    : positiveInteger_(config.SEARCH_MAX_RESULTS, 12);
  const minScore = parseFloat(config.SEARCH_MIN_SCORE || '0.20');
  const ranked = rankSearchRows_(indexRows, queryEmbedding, Number.isNaN(minScore) ? 0.20 : minScore, maxResults);
  renderSearchResults_(searchSheet, ranked, now);
  return 'Found ' + ranked.length + ' search results.';
}

function refreshSearchIndex_(spreadsheet, config, forceRebuild) {
  const entriesSheet = spreadsheet.getSheetByName(SHEET_NAMES.ENTRIES);
  const indexSheet = spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX);
  if (forceRebuild) {
    clearSheetRows_(indexSheet, 2);
  }

  const existingIndex = forceRebuild ? {} : buildSearchIndexState_(getRowsAsObjects_(indexSheet));
  const entries = getRowsAsObjects_(entriesSheet);
  let indexedCount = 0;
  entries.forEach(function(row) {
    const entry = row.values;
    const entryId = entry.entry_id;
    const transcript = String(entry.transcript || '');
    if (!entryId || !transcript.trim() || String(entry.status || '').toLowerCase() !== 'processed') {
      return;
    }
    const transcriptHash = transcriptHash_(transcript);
    if (!forceRebuild && existingIndex[entryId] && existingIndex[entryId].transcript_hash === transcriptHash) {
      return;
    }
    indexEntryForSearch_(spreadsheet, entryId, config);
    indexedCount += 1;
  });
  hideSearchIndexSheet_(spreadsheet);
  return indexedCount;
}

function indexEntryForSearch_(spreadsheet, entryId, config) {
  const entries = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.ENTRIES));
  const entryRow = entries.find(function(row) {
    return row.values.entry_id === entryId;
  });
  if (!entryRow) {
    return 0;
  }

  const indexSheet = spreadsheet.getSheetByName(SHEET_NAMES.SEARCH_INDEX);
  deleteRowsWhereColumnEquals_(indexSheet, 'entry_id', entryId);
  const reflectionContext = buildReflectionContextByEntryId_(getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.REFLECTIONS)));
  const rows = buildSearchIndexRowsForEntry_(entryRow, reflectionContext[entryId], config, new Date(), function(text) {
    return embedText_(text, config);
  });
  rows.forEach(function(row) {
    appendObjectRow_(indexSheet, row);
  });
  const failedCount = rows.filter(function(row) {
    return row.index_status === 'Failed';
  }).length;
  if (failedCount > 0) {
    spreadsheet.toast('Search indexing failed for ' + failedCount + ' chunk(s) in ' + entryId + '.', 'Voice Journal', 5);
  }
  return rows.length;
}

function buildSearchIndexRowsForEntry_(entryRow, reflectionContext, config, now, embeddingFn) {
  const entry = entryRow.values || entryRow;
  const transcript = String(entry.transcript || '');
  const targetChars = positiveInteger_(config.SEARCH_CHUNK_TARGET_CHARS, 1200);
  const overlapChars = Math.min(positiveInteger_(config.SEARCH_CHUNK_OVERLAP_CHARS, 200), Math.floor(targetChars / 2));
  const chunks = chunkTranscript_(transcript, targetChars, overlapChars);
  const model = config.EMBEDDING_MODEL || 'text-embedding-3-small';
  const dimensions = positiveInteger_(config.EMBEDDING_DIMENSIONS, 512);
  const transcriptHash = transcriptHash_(transcript);
  const context = reflectionContext || { related_reflections: '', moods: '' };

  return chunks.map(function(chunk, index) {
    try {
      const embedding = embeddingFn(chunk.quote_text);
      return {
        chunk_id: entry.entry_id + '_chunk_' + (index + 1),
        entry_id: entry.entry_id,
        source_row: entryRow.rowNumber || '',
        uploaded_at: entry.uploaded_at || '',
        processed_at: entry.processed_at || '',
        audio_url: entry.audio_url || '',
        chunk_start_char: chunk.chunk_start_char,
        chunk_end_char: chunk.chunk_end_char,
        quote_text: chunk.quote_text,
        context_before: chunk.context_before,
        context_after: chunk.context_after,
        related_reflections: context.related_reflections,
        moods: context.moods,
        embedding_json: JSON.stringify(embedding),
        embedding_model: model,
        embedding_dimensions: dimensions,
        transcript_hash: transcriptHash,
        indexed_at: now,
        index_status: 'Indexed',
        index_error: '',
      };
    } catch (error) {
      return {
        chunk_id: entry.entry_id + '_chunk_' + (index + 1),
        entry_id: entry.entry_id,
        source_row: entryRow.rowNumber || '',
        uploaded_at: entry.uploaded_at || '',
        processed_at: entry.processed_at || '',
        audio_url: entry.audio_url || '',
        chunk_start_char: chunk.chunk_start_char,
        chunk_end_char: chunk.chunk_end_char,
        quote_text: chunk.quote_text,
        context_before: chunk.context_before,
        context_after: chunk.context_after,
        related_reflections: context.related_reflections,
        moods: context.moods,
        embedding_json: '',
        embedding_model: model,
        embedding_dimensions: dimensions,
        transcript_hash: transcriptHash,
        indexed_at: now,
        index_status: 'Failed',
        index_error: errorToString_(error),
      };
    }
  });
}

function embedText_(text, config) {
  const apiKey = getOpenAiApiKey_(config);
  const dimensions = positiveInteger_(config.EMBEDDING_DIMENSIONS, 512);
  const payload = {
    model: config.EMBEDDING_MODEL || 'text-embedding-3-small',
    input: text,
    encoding_format: 'float',
    dimensions: dimensions,
  };
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/embeddings', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
    payload: JSON.stringify(payload),
  });
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Embedding failed with HTTP ' + code + ': ' + body);
  }
  const parsed = JSON.parse(body);
  const embedding = parsed.data && parsed.data[0] && parsed.data[0].embedding;
  if (!Array.isArray(embedding)) {
    throw new Error('Embedding response did not include an embedding vector.');
  }
  return embedding;
}

function rankSearchRows_(indexRows, queryEmbedding, minScore, maxResults) {
  return normalizeRowsForValues_(indexRows)
    .filter(function(row) {
      return row.index_status === 'Indexed' && row.embedding_json && row.quote_text;
    })
    .map(function(row) {
      const embedding = parseEmbeddingJson_(row.embedding_json);
      return copyObject_(row, { score: cosineSimilarity_(queryEmbedding, embedding) });
    })
    .filter(function(row) {
      return row.score >= minScore;
    })
    .sort(function(a, b) {
      return b.score - a.score;
    })
    .slice(0, maxResults);
}

function renderSearchResults_(sheet, rankedRows, now) {
  clearSheetRows_(sheet, SEARCH.RESULT_START_ROW);
  sheet.getRange(SEARCH.RESULT_HEADER_ROW, 1, 1, SEARCH.RESULT_COLUMN_COUNT)
    .setValues([HEADERS.Search])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.getRange(SEARCH.LAST_SEARCH_CELL).setValue(formatDateTimeForDashboard_(now));
  if (rankedRows.length === 0) {
    sheet.getRange(SEARCH.RESULT_START_ROW, 1).setValue('No matching quotes found.');
    return;
  }
  const values = buildSearchResultRows_(rankedRows);
  sheet.getRange(SEARCH.RESULT_START_ROW, 1, values.length, SEARCH.RESULT_COLUMN_COUNT).setValues(values);
  sheet.getRange(SEARCH.RESULT_START_ROW, 2, values.length, 1).setNumberFormat('0.000');
  sheet.getRange(SEARCH.RESULT_START_ROW, 6, values.length, 1).setNumberFormat('yyyy-mm-dd');
}

function buildSearchResultRows_(rankedRows) {
  return rankedRows.map(function(row, index) {
    return [
      index + 1,
      row.score,
      row.quote_text,
      combineContext_(row.context_before, row.context_after),
      row.entry_id,
      parseDate_(row.uploaded_at) || row.uploaded_at || '',
      row.audio_url || '',
      row.related_reflections || '',
      row.moods || '',
    ];
  });
}

function buildSearchResultObjects_(rankedRows) {
  return normalizeRowsForValues_(rankedRows).map(function(row, index) {
    return {
      rank: index + 1,
      score: roundScore_(row.score),
      quote: row.quote_text || '',
      contextBefore: row.context_before || '',
      contextAfter: row.context_after || '',
      entryId: row.entry_id || '',
      uploadedAt: formatDateForWeb_(row.uploaded_at),
      audioUrl: row.audio_url || '',
      relatedReflections: row.related_reflections || '',
      moods: row.moods || '',
    };
  });
}

function buildSearchAppStateFromRows_(indexRows, config) {
  let indexedChunks = 0;
  let failedChunks = 0;
  let lastIndexedAt = null;
  normalizeRowsForValues_(indexRows).forEach(function(row) {
    if (row.index_status === 'Indexed') {
      indexedChunks += 1;
    } else if (row.index_status === 'Failed') {
      failedChunks += 1;
    }
    const indexedAt = parseDate_(row.indexed_at);
    if (indexedAt && (!lastIndexedAt || indexedAt.getTime() > lastIndexedAt.getTime())) {
      lastIndexedAt = indexedAt;
    }
  });

  return {
    indexedChunks: indexedChunks,
    failedChunks: failedChunks,
    totalChunks: normalizeRowsForValues_(indexRows).length,
    lastIndexedAt: lastIndexedAt ? formatDateTimeForWeb_(lastIndexedAt) : '',
    defaults: {
      maxResults: positiveInteger_(config.SEARCH_MAX_RESULTS, 12),
      minScore: parseSearchScore_(config.SEARCH_MIN_SCORE, 0.20),
      embeddingModel: config.EMBEDDING_MODEL || 'text-embedding-3-small',
      embeddingDimensions: positiveInteger_(config.EMBEDDING_DIMENSIONS, 512),
    },
  };
}

function validateSearchQuery_(query) {
  const cleanQuery = String(query || '').trim();
  if (!cleanQuery) {
    throw new Error('Enter a search query before searching.');
  }
  return cleanQuery;
}

function normalizeSearchOptions_(options, config) {
  const provided = options || {};
  const configuredMax = positiveInteger_(config.SEARCH_MAX_RESULTS, 12);
  const configuredMinScore = parseSearchScore_(config.SEARCH_MIN_SCORE, 0.20);
  const maxResults = positiveInteger_(provided.maxResults, configuredMax);
  return {
    maxResults: clampInteger_(maxResults, 1, 50),
    minScore: parseSearchScore_(provided.minScore, configuredMinScore),
  };
}

function parseSearchScore_(value, fallback) {
  const parsed = parseFloat(value);
  if (Number.isNaN(parsed)) {
    return fallback;
  }
  return Math.max(0, Math.min(1, parsed));
}

function roundScore_(score) {
  return Math.round((Number(score) || 0) * 1000) / 1000;
}

function formatDateForWeb_(value) {
  const date = parseDate_(value);
  if (!date) {
    return value ? String(value) : '';
  }
  return formatDateForDigest_(date);
}

function formatDateTimeForWeb_(value) {
  const date = parseDate_(value);
  if (!date) {
    return value ? String(value) : '';
  }
  return formatDateTimeForDashboard_(date);
}

function chunkTranscript_(transcript, targetChars, overlapChars) {
  const text = String(transcript || '').trim();
  if (!text) {
    return [];
  }
  const target = positiveInteger_(targetChars, 1200);
  const overlap = Math.min(Math.max(parseInteger_(overlapChars, 0), 0), Math.floor(target / 2));
  if (text.length <= target) {
    return [buildTranscriptChunk_(text, 0, text.length, text)];
  }

  const chunks = [];
  let start = 0;
  while (start < text.length) {
    const end = chooseTranscriptChunkEnd_(text, start, target);
    const rawQuote = text.slice(start, end);
    const leadingWhitespace = rawQuote.search(/\S/);
    const quoteStart = start + (leadingWhitespace === -1 ? 0 : leadingWhitespace);
    const quote = text.slice(quoteStart, end).trimEnd();
    if (quote) {
      chunks.push(buildTranscriptChunk_(text, quoteStart, quoteStart + quote.length, quote));
    }
    if (end >= text.length) {
      break;
    }
    start = Math.max(end - overlap, start + 1);
    while (start < text.length && /\s/.test(text.charAt(start))) {
      start += 1;
    }
  }
  return chunks;
}

function chooseTranscriptChunkEnd_(text, start, targetChars) {
  const maxEnd = Math.min(text.length, start + targetChars);
  if (maxEnd >= text.length) {
    return text.length;
  }
  const minEnd = start + Math.floor(targetChars * 0.5);
  const candidate = text.slice(start, maxEnd);
  const boundaryMatches = candidate.match(/[.!?]\s+|\n+/g) || [];
  if (boundaryMatches.length > 0) {
    let searchFrom = 0;
    let boundaryEnd = -1;
    boundaryMatches.forEach(function(match) {
      const index = candidate.indexOf(match, searchFrom);
      searchFrom = index + match.length;
      const absoluteEnd = start + index + match.length;
      if (absoluteEnd >= minEnd) {
        boundaryEnd = absoluteEnd;
      }
    });
    if (boundaryEnd > start) {
      return boundaryEnd;
    }
  }
  return maxEnd;
}

function buildTranscriptChunk_(fullText, start, end, quote) {
  const sourceText = String(fullText || '');
  const safeStart = Math.max(0, start);
  const safeEnd = Math.min(sourceText.length, end);
  return {
    chunk_start_char: safeStart,
    chunk_end_char: safeEnd,
    quote_text: String(quote || ''),
    context_before: sourceText.slice(Math.max(0, safeStart - 240), safeStart).trim(),
    context_after: sourceText.slice(safeEnd, Math.min(sourceText.length, safeEnd + 240)).trim(),
  };
}

function transcriptHash_(transcript) {
  const text = String(transcript || '');
  if (typeof Utilities !== 'undefined' && Utilities.computeDigest) {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
    return digest.map(function(byte) {
      const value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    }).join('');
  }
  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    hash = ((hash << 5) - hash + text.charCodeAt(i)) | 0;
  }
  return String(hash);
}

function buildReflectionContextByEntryId_(reflectionRows) {
  const contextByEntryId = {};
  normalizeRowsForValues_(reflectionRows).forEach(function(reflection) {
    const entryId = reflection.entry_id;
    if (!entryId) {
      return;
    }
    if (!contextByEntryId[entryId]) {
      contextByEntryId[entryId] = { reflections: [], moods: [] };
    }
    if (reflection.summary) {
      contextByEntryId[entryId].reflections.push(reflection.summary);
    }
    String(reflection.moods || '').split(',').forEach(function(rawMood) {
      const mood = rawMood.trim();
      if (mood && contextByEntryId[entryId].moods.indexOf(mood) === -1) {
        contextByEntryId[entryId].moods.push(mood);
      }
    });
  });
  Object.keys(contextByEntryId).forEach(function(entryId) {
    contextByEntryId[entryId] = {
      related_reflections: contextByEntryId[entryId].reflections.join('; '),
      moods: contextByEntryId[entryId].moods.join(', '),
    };
  });
  return contextByEntryId;
}

function buildSearchIndexState_(indexRows) {
  const state = {};
  normalizeRowsForValues_(indexRows).forEach(function(row) {
    if (row.entry_id && row.transcript_hash && row.index_status === 'Indexed') {
      state[row.entry_id] = { transcript_hash: row.transcript_hash };
    }
  });
  return state;
}

function parseEmbeddingJson_(embeddingJson) {
  try {
    const embedding = JSON.parse(embeddingJson);
    return Array.isArray(embedding) ? embedding : [];
  } catch (error) {
    return [];
  }
}

function cosineSimilarity_(a, b) {
  const length = Math.min(normalizeArray_(a).length, normalizeArray_(b).length);
  if (length === 0) {
    return 0;
  }
  let dot = 0;
  let magnitudeA = 0;
  let magnitudeB = 0;
  for (let i = 0; i < length; i++) {
    const valueA = Number(a[i]) || 0;
    const valueB = Number(b[i]) || 0;
    dot += valueA * valueB;
    magnitudeA += valueA * valueA;
    magnitudeB += valueB * valueB;
  }
  if (!magnitudeA || !magnitudeB) {
    return 0;
  }
  return dot / (Math.sqrt(magnitudeA) * Math.sqrt(magnitudeB));
}

function combineContext_(before, after) {
  const parts = [];
  if (before) {
    parts.push('Before: ' + before);
  }
  if (after) {
    parts.push('After: ' + after);
  }
  return parts.join('\n\n');
}
