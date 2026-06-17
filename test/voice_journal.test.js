const assert = require('assert');

const {
  HEADERS,
  SHEET_NAMES,
  DEFAULT_DESIGN_ROWS,
  DESIGN,
  assertVoiceJournalInstallReady_,
  buildVoiceJournalInstallPlan_,
  buildVoiceJournalInstallSummary_,
  buildDigestFromRows_,
  buildDashboardGoalRows_,
  buildDashboardFollowUpRows_,
  buildRecentMoodCounts_,
  buildSearchIndexRowsForEntry_,
  buildSearchAppStateFromRows_,
  buildSearchResultObjects_,
  buildSearchResultRows_,
  buildReflectionContextByEntryId_,
  chunkTranscript_,
  cosineSimilarity_,
  dashboardGoalActionStatus_,
  dashboardCheckboxStatus_,
  fallbackDigestIntro_,
  findMatchedKeywords_,
  moodCountsToRows_,
  customFactFields_,
  defaultCustomFactFields_,
  normalizeCustomExtraction_,
  normalizeDesignRow_,
  normalizeSearchOptions_,
  parseFactTableDictionary_,
  parseKeywords_,
  rankSearchRows_,
  reservedCustomFactColumns_,
  shouldRunCustomExtraction_,
  customExtractionHeaders_,
  customExtractionSchema_,
  customFactFieldSchema_,
  clearConfigValue_,
  extractionSchema_,
  getOpenAiApiKeyPropertyName_,
  transcriptHash_,
  validateSearchQuery_,
} = require('./load_app_script');

function test(name, fn) {
  try {
    fn();
    console.log('ok - ' + name);
  } catch (error) {
    console.error('not ok - ' + name);
    throw error;
  }
}

test('filters active dashboard follow ups', function() {
  const rows = [
    { values: { follow_up_id: 'followUp_1', summary: 'Call Sam', status: 'Open', due_date_optional: '2026-05-20', created_at: '2026-05-18' } },
    { values: { follow_up_id: 'followUp_2', summary: 'File taxes', status: 'Done', due_date_optional: '', created_at: '2026-05-17' } },
    { values: { follow_up_id: 'followUp_3', summary: 'Draft outline', status: 'Active', due_date_optional: '', created_at: '2026-05-16' } },
    { values: { follow_up_id: 'followUp_4', summary: 'Old thing', status: 'Archived', due_date_optional: '', created_at: '2026-05-15' } },
  ];

  const dashboardRows = buildDashboardFollowUpRows_(rows);

  assert.strictEqual(dashboardRows.length, 2);
  assert.deepStrictEqual(dashboardRows.map(function(row) { return row[5]; }), ['followUp_1', 'followUp_3']);
  assert.strictEqual(dashboardRows[0][0], false);
});

test('counts recent comma-separated moods', function() {
  const now = new Date('2026-05-19T12:00:00Z');
  const rows = [
    { values: { moods: 'focused, calm', created_at: '2026-05-19T09:00:00Z' } },
    { values: { moods: 'calm, grateful', created_at: '2026-05-18T09:00:00Z' } },
  ];

  assert.deepStrictEqual(buildRecentMoodCounts_(rows, now, 7), {
    calm: 2,
    focused: 1,
    grateful: 1,
  });
});

test('ignores reflections outside the mood lookback window', function() {
  const now = new Date('2026-05-19T12:00:00Z');
  const rows = [
    { values: { moods: 'focused', created_at: '2026-05-19T09:00:00Z' } },
    { values: { moods: 'tired', created_at: '2026-05-01T09:00:00Z' } },
  ];

  assert.deepStrictEqual(buildRecentMoodCounts_(rows, now, 7), { focused: 1 });
});

test('sorts mood chart rows by count then name', function() {
  const rows = moodCountsToRows_({ calm: 2, focused: 3, anxious: 2 });

  assert.deepStrictEqual(rows, [
    ['focused', 3],
    ['anxious', 2],
    ['calm', 2],
  ]);
});

test('maps dashboard checkbox edits to follow up statuses', function() {
  assert.strictEqual(dashboardCheckboxStatus_(true), 'Done');
  assert.strictEqual(dashboardCheckboxStatus_('TRUE'), 'Done');
  assert.strictEqual(dashboardCheckboxStatus_(false), 'Open');
  assert.strictEqual(dashboardCheckboxStatus_('FALSE'), 'Open');
});

test('filters active dashboard goals', function() {
  const rows = [
    { values: { goal_id: 'goal_1', summary: 'Run a half marathon', status: 'Active', active_until: '2026-06-01', created_at: '2026-05-18' } },
    { values: { goal_id: 'goal_2', summary: 'Read more', status: 'Open', active_until: '', created_at: '2026-05-17' } },
    { values: { goal_id: 'goal_3', summary: 'Done goal', status: 'Complete', active_until: '', created_at: '2026-05-16' } },
  ];

  const dashboardRows = buildDashboardGoalRows_(rows);

  assert.strictEqual(dashboardRows.length, 2);
  assert.deepStrictEqual(dashboardRows.map(function(row) { return row[5]; }), ['goal_1', 'goal_2']);
  assert.strictEqual(dashboardRows[0][0], '');
});

test('excludes complete completed and archived dashboard goals', function() {
  const rows = [
    { values: { goal_id: 'goal_1', summary: 'Done goal', status: 'Complete' } },
    { values: { goal_id: 'goal_2', summary: 'Also done', status: 'Completed' } },
    { values: { goal_id: 'goal_3', summary: 'Archived goal', status: 'Archived' } },
    { values: { goal_id: 'goal_4', summary: 'Active goal', status: 'Active' } },
  ];

  const dashboardRows = buildDashboardGoalRows_(rows);

  assert.strictEqual(dashboardRows.length, 1);
  assert.strictEqual(dashboardRows[0][5], 'goal_4');
});

test('maps dashboard goal actions to source statuses', function() {
  assert.strictEqual(dashboardGoalActionStatus_('Complete'), 'Complete');
  assert.strictEqual(dashboardGoalActionStatus_('complete'), 'Complete');
  assert.strictEqual(dashboardGoalActionStatus_('Archive'), 'Archived');
  assert.strictEqual(dashboardGoalActionStatus_('archive'), 'Archived');
  assert.strictEqual(dashboardGoalActionStatus_(''), '');
  assert.strictEqual(dashboardGoalActionStatus_('Active'), '');
});

test('search chunking creates one exact chunk for short transcripts', function() {
  const transcript = 'I felt clear about the project today. The next step is a small prototype.';
  const chunks = chunkTranscript_(transcript, 1200, 200);

  assert.strictEqual(chunks.length, 1);
  assert.strictEqual(chunks[0].quote_text, transcript);
  assert.strictEqual(chunks[0].chunk_start_char, 0);
  assert.strictEqual(chunks[0].chunk_end_char, transcript.length);
});

test('search chunking overlaps long transcripts', function() {
  const transcript = [
    'First sentence about planning.',
    'Second sentence about building.',
    'Third sentence about reviewing.',
    'Fourth sentence about shipping.',
  ].join(' ');
  const chunks = chunkTranscript_(transcript, 65, 15);

  assert.ok(chunks.length > 1);
  assert.ok(chunks[1].chunk_start_char < chunks[0].chunk_end_char);
  assert.strictEqual(transcript.slice(chunks[0].chunk_start_char, chunks[0].chunk_end_char), chunks[0].quote_text);
});

test('search chunking ignores empty transcripts', function() {
  assert.deepStrictEqual(chunkTranscript_('', 1200, 200), []);
  assert.deepStrictEqual(chunkTranscript_('   ', 1200, 200), []);
});

test('transcript hash changes when transcript changes', function() {
  const first = transcriptHash_('A stable note.');
  const second = transcriptHash_('A changed note.');

  assert.strictEqual(transcriptHash_('A stable note.'), first);
  assert.notStrictEqual(first, second);
});

test('cosine similarity ranks matching search rows first', function() {
  const ranked = rankSearchRows_(
    [
      { values: { chunk_id: 'chunk_1', quote_text: 'close', embedding_json: JSON.stringify([1, 0]), index_status: 'Indexed' } },
      { values: { chunk_id: 'chunk_2', quote_text: 'far', embedding_json: JSON.stringify([0, 1]), index_status: 'Indexed' } },
    ],
    [1, 0],
    0,
    2
  );

  assert.strictEqual(cosineSimilarity_([1, 0], [1, 0]), 1);
  assert.strictEqual(ranked[0].chunk_id, 'chunk_1');
});

test('search result rows preserve quotes and source metadata', function() {
  const rows = buildSearchResultRows_([
    {
      score: 0.91,
      quote_text: 'Exact transcript quote.',
      context_before: 'Before text.',
      context_after: 'After text.',
      entry_id: 'entry_1',
      uploaded_at: '2026-05-20T12:00:00Z',
      audio_url: 'https://drive.example/audio',
      related_reflections: 'A reflection',
      moods: 'focused',
    },
  ]);

  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0][0], 1);
  assert.strictEqual(rows[0][2], 'Exact transcript quote.');
  assert.strictEqual(rows[0][4], 'entry_1');
  assert.strictEqual(rows[0][7], 'A reflection');
});

test('search result objects preserve exact quotes and metadata for web app', function() {
  const results = buildSearchResultObjects_([
    {
      score: 0.91234,
      quote_text: 'Exact transcript quote.',
      context_before: 'Before text.',
      context_after: 'After text.',
      entry_id: 'entry_1',
      uploaded_at: '2026-05-20T12:00:00Z',
      audio_url: 'https://drive.example/audio',
      related_reflections: 'A reflection',
      moods: 'focused',
    },
  ]);

  assert.strictEqual(results.length, 1);
  assert.strictEqual(results[0].rank, 1);
  assert.strictEqual(results[0].score, 0.912);
  assert.strictEqual(results[0].quote, 'Exact transcript quote.');
  assert.strictEqual(results[0].contextBefore, 'Before text.');
  assert.strictEqual(results[0].contextAfter, 'After text.');
  assert.strictEqual(results[0].entryId, 'entry_1');
  assert.strictEqual(results[0].uploadedAt, '2026-05-20');
  assert.strictEqual(results[0].audioUrl, 'https://drive.example/audio');
});

test('web search rejects empty queries', function() {
  assert.throws(function() {
    validateSearchQuery_('   ');
  }, /Enter a search query/);
  assert.strictEqual(validateSearchQuery_(' focus '), 'focus');
});

test('search app state summarizes indexed and failed chunks', function() {
  const state = buildSearchAppStateFromRows_(
    [
      { values: { index_status: 'Indexed', indexed_at: '2026-05-20T12:00:00Z' } },
      { values: { index_status: 'Failed', indexed_at: '2026-05-21T12:00:00Z' } },
      { values: { index_status: 'Indexed', indexed_at: '2026-05-19T12:00:00Z' } },
    ],
    { SEARCH_MAX_RESULTS: '20', SEARCH_MIN_SCORE: '0.25', EMBEDDING_MODEL: 'embed-test', EMBEDDING_DIMENSIONS: '256' }
  );

  assert.strictEqual(state.indexedChunks, 2);
  assert.strictEqual(state.failedChunks, 1);
  assert.strictEqual(state.totalChunks, 3);
  assert.strictEqual(state.lastIndexedAt, '2026-05-21 12:00');
  assert.strictEqual(state.defaults.maxResults, 20);
  assert.strictEqual(state.defaults.minScore, 0.25);
  assert.strictEqual(state.defaults.embeddingModel, 'embed-test');
  assert.strictEqual(state.defaults.embeddingDimensions, 256);
});

test('web search options use config defaults and clamp max results', function() {
  assert.deepStrictEqual(
    normalizeSearchOptions_({}, { SEARCH_MAX_RESULTS: '12', SEARCH_MIN_SCORE: '0.2' }),
    { maxResults: 12, minScore: 0.2 }
  );
  assert.deepStrictEqual(
    normalizeSearchOptions_({ maxResults: '100', minScore: '1.5' }, { SEARCH_MAX_RESULTS: '12', SEARCH_MIN_SCORE: '0.2' }),
    { maxResults: 50, minScore: 1 }
  );
});

test('search index enrichment joins reflections by entry id', function() {
  const context = buildReflectionContextByEntryId_([
    { values: { entry_id: 'entry_1', summary: 'Prototype felt promising', moods: 'focused, curious' } },
    { values: { entry_id: 'entry_1', summary: 'Need a calmer review loop', moods: 'calm, focused' } },
    { values: { entry_id: 'entry_2', summary: 'Different entry', moods: 'tired' } },
  ]);

  assert.strictEqual(context.entry_1.related_reflections, 'Prototype felt promising; Need a calmer review loop');
  assert.strictEqual(context.entry_1.moods, 'focused, curious, calm');
});

test('search index rows include enriched context and embeddings', function() {
  const rows = buildSearchIndexRowsForEntry_(
    {
      rowNumber: 2,
      values: {
        entry_id: 'entry_1',
        transcript: 'I want to build a focused search tab.',
        uploaded_at: '2026-05-20',
        audio_url: 'https://drive.example/audio',
      },
    },
    { related_reflections: 'Search matters', moods: 'focused' },
    { EMBEDDING_MODEL: 'test-embedding', EMBEDDING_DIMENSIONS: 2, SEARCH_CHUNK_TARGET_CHARS: 1200, SEARCH_CHUNK_OVERLAP_CHARS: 200 },
    new Date('2026-05-20T12:00:00Z'),
    function() { return [1, 0]; }
  );

  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0].entry_id, 'entry_1');
  assert.strictEqual(rows[0].related_reflections, 'Search matters');
  assert.strictEqual(rows[0].moods, 'focused');
  assert.strictEqual(rows[0].embedding_json, JSON.stringify([1, 0]));
  assert.strictEqual(rows[0].index_status, 'Indexed');
});

test('digest reads reminder state from goals and follow ups', function() {
  const now = new Date('2026-05-20T12:00:00Z');
  const digest = buildDigestFromRows_(
    [
      {
        rowNumber: 2,
        values: {
          goal_id: 'goal_1',
          summary: 'Run a half marathon',
          status: 'Active',
          active_until: '2026-06-01',
          reminder_frequency_days: 7,
          next_reminder_at: '2026-05-20T10:00:00Z',
          reminder_count: 2,
        },
      },
      {
        rowNumber: 3,
        values: {
          goal_id: 'goal_2',
          summary: 'Read more',
          status: 'Active',
          active_until: '2026-06-01',
          reminder_frequency_days: 7,
          next_reminder_at: '2026-05-27T10:00:00Z',
          reminder_count: 0,
        },
      },
    ],
    [
      {
        rowNumber: 2,
        values: {
          follow_up_id: 'followUp_1',
          summary: 'Call Sam',
          status: 'Open',
          created_at: '2026-05-10T09:00:00Z',
          due_date_optional: '',
          reminder_frequency_days: 3,
          next_reminder_at: '2026-05-20T09:00:00Z',
          reminder_count: 1,
        },
      },
      {
        rowNumber: 3,
        values: {
          follow_up_id: 'followUp_2',
          summary: 'Prep report',
          status: 'Open',
          created_at: '2026-05-01T09:00:00Z',
          due_date_optional: '2026-05-23',
          reminder_frequency_days: 3,
          next_reminder_at: '2026-06-01T09:00:00Z',
          reminder_count: 0,
        },
      },
    ],
    [
      { values: { reflection_id: 'reflection_1', summary: 'Felt focused during planning', moods: 'focused', created_at: '2026-05-19T09:00:00Z' } },
      { values: { reflection_id: 'reflection_2', summary: 'Old note', moods: 'tired', created_at: '2026-04-01T09:00:00Z' } },
    ],
    { DIGEST_LOOKBACK_DAYS: 7 },
    now,
    function() { throw new Error('intro failed'); }
  );

  assert.match(digest.body, /Today's digest includes 2 active goals, 2 follow ups, and 1 recent reflections\./);
  assert.match(digest.body, /__Goals__\n2 active goals are included, with 1 currently reminder-due\./);
  assert.match(digest.body, /__Follow Ups__\n2 follow ups are included, with 1 currently reminder-due\./);
  assert.match(digest.body, /__Reflections__\n1 recent reflections are included\. The most common mood tags are focused\./);
  assert.match(digest.body, /Run a half marathon through 2026-06-01 \[reminder due\]/);
  assert.match(digest.body, /Read more through 2026-06-01\n/);
  assert.match(digest.body, /Call Sam \[reminder due\]/);
  assert.match(digest.body, /Prep report due 2026-05-23/);
  assert.match(digest.body, /Felt focused during planning \[focused\]/);
  assert.doesNotMatch(digest.body, /Old note/);

  assert.strictEqual(digest.goalUpdates.length, 1);
  assert.strictEqual(digest.goalUpdates[0].rowNumber, 2);
  assert.strictEqual(digest.goalUpdates[0].values.reminder_count, 3);
  assert.strictEqual(digest.goalUpdates[0].values.last_reminded_at, now);
  assert.strictEqual(digest.goalUpdates[0].values.next_reminder_at.toISOString(), '2026-05-27T12:00:00.000Z');

  assert.strictEqual(digest.followUpUpdates.length, 1);
  assert.strictEqual(digest.followUpUpdates[0].rowNumber, 2);
  assert.strictEqual(digest.followUpUpdates[0].values.reminder_count, 2);
  assert.strictEqual(digest.followUpUpdates[0].values.next_reminder_at.toISOString(), '2026-05-23T12:00:00.000Z');
});

test('digest html bolds underlines and numbers section items', function() {
  const digest = buildDigestFromRows_(
    [
      {
        values: {
          goal_id: 'goal_1',
          summary: 'Keep training',
          status: 'Active',
          active_until: '',
          reminder_frequency_days: 7,
          next_reminder_at: '2026-05-27T10:00:00Z',
        },
      },
    ],
    [],
    [],
    { DIGEST_LOOKBACK_DAYS: 7 },
    new Date('2026-05-20T12:00:00Z'),
    function() {
      return {
        intro: 'A realistic opening.',
        goalsSummary: 'One active goal is present.',
        followUpsSummary: 'No follow ups are currently included.',
        reflectionsSummary: 'No recent reflections are currently included.',
      };
    }
  );

  assert.match(digest.htmlBody, /<strong><u>Goals<\/u><\/strong>/);
  assert.match(digest.htmlBody, /<strong><u>Follow Ups<\/u><\/strong>/);
  assert.match(digest.htmlBody, /<strong><u>Reflections<\/u><\/strong>/);
  assert.match(digest.htmlBody, /<ol style=/);
  assert.match(digest.htmlBody, /<li>Keep training<\/li>/);
  assert.match(digest.htmlBody, /One active goal is present\./);
});

test('digest uses AI introduction when available', function() {
  const digest = buildDigestFromRows_(
    [],
    [],
    [],
    { DIGEST_LOOKBACK_DAYS: 7 },
    new Date('2026-05-20T12:00:00Z'),
    function() { return 'A compact AI-written opening.'; }
  );

  assert.match(digest.body, /A compact AI-written opening\./);
});

test('fallback digest intro summarizes counts and moods', function() {
  const intro = fallbackDigestIntro_({
    goals: [{ summary: 'Goal' }],
    followUps: [{ summary: 'Task' }, { summary: 'Second summary' }],
    reflections: [{ summary: 'reflection', moods: 'calm, focused' }],
    moodCounts: { calm: 1, focused: 1 },
  });

  assert.match(intro, /1 active goals, 2 follow ups, and 1 recent reflections/);
  assert.match(intro, /calm and focused|focused and calm/);
});

test('reminders table is no longer part of the configured schema', function() {
  assert.strictEqual(SHEET_NAMES.REMINDERS, undefined);
  assert.strictEqual(HEADERS.Reminders, undefined);
  assert.ok(HEADERS.Goals.indexOf('next_reminder_at') !== -1);
  assert.ok(HEADERS['Follow Ups'].indexOf('last_reminded_at') !== -1);
});

test('built-in extraction schema uses follow ups and reflections', function() {
  const schema = extractionSchema_();
  assert.deepStrictEqual(schema.required, ['goals', 'followUps', 'reflections']);
  assert.ok(schema.properties.followUps);
  assert.ok(schema.properties.reflections);
  assert.deepStrictEqual(customExtractionSchema_().required, ['records']);
  assert.ok(HEADERS.Entries.indexOf('has_follow_up') !== -1);
  assert.ok(HEADERS.Entries.indexOf('has_reflection') !== -1);
  assert.ok(HEADERS['Follow Ups'].indexOf('follow_up_id') !== -1);
  assert.ok(HEADERS.Reflections.indexOf('reflection_id') !== -1);
});

test('design defaults seed the built-in extraction rows', function() {
  assert.strictEqual(SHEET_NAMES.DESIGN, 'Design');
  assert.deepStrictEqual(DEFAULT_DESIGN_ROWS.map(function(row) { return row.extraction_key; }), [
    'follow_ups',
    'goals',
    'reflections',
  ]);
  assert.strictEqual(DESIGN.ACTION_COLUMN, 9);
});

test('normalizes design rows and fallback custom headers', function() {
  const row = normalizeDesignRow_({
    extraction_key: 'Recurring Themes!',
    display_name: '',
    sheet_name: 'Recurring/Themes',
    keywords: 'theme, pattern; theme',
    ai_decides_without_keyword: 'FALSE',
    status: '',
  });

  assert.strictEqual(row.extraction_key, 'recurring_themes');
  assert.strictEqual(row.display_name, 'Recurring Themes');
  assert.strictEqual(row.sheet_name, 'Recurring Themes');
  assert.strictEqual(row.fact_table_dictionary, '');
  assert.strictEqual(row.keywords, 'theme, pattern');
  assert.strictEqual(row.ai_decides_without_keyword, false);
  assert.strictEqual(row.status, 'Active');
  assert.deepStrictEqual(customExtractionHeaders_(row.extraction_key), [
    'recurring_themes_id',
    'entry_id',
    'matched_keywords',
    'created_at',
    'extracted_at',
    'summary',
    'evidence',
    'status',
  ]);
});

test('parses custom fact table dictionary fields', function() {
  const dictionary = JSON.stringify({
    fields: [
      { key: 'summary', type: 'string', description: 'Short summary.' },
      { key: 'priority', type: 'enum', description: 'Priority.', values: ['low', 'medium', 'high'] },
      { key: 'is_recurring', type: 'boolean', description: 'Whether this repeats.' },
      { key: 'score', type: 'number', description: 'Confidence score.' },
      { key: 'target_date', type: 'date', description: 'Relevant date.' },
    ],
  });

  const parsed = parseFactTableDictionary_(dictionary, 'themes');

  assert.deepStrictEqual(parsed.fields.map(function(field) { return field.key; }), [
    'summary',
    'priority',
    'is_recurring',
    'score',
    'target_date',
  ]);
  assert.deepStrictEqual(parsed.fields[1].values, ['low', 'medium', 'high']);
});

test('rejects invalid fact table dictionary definitions', function() {
  assert.throws(function() {
    parseFactTableDictionary_('{', 'themes');
  }, /valid JSON/);
  assert.throws(function() {
    parseFactTableDictionary_(JSON.stringify({ fields: [{ key: 'BadKey', type: 'string' }] }), 'themes');
  }, /snake_case/);
  assert.throws(function() {
    parseFactTableDictionary_(JSON.stringify({ fields: [{ key: 'rating', type: 'object' }] }), 'themes');
  }, /unsupported/);
  assert.throws(function() {
    parseFactTableDictionary_(JSON.stringify({ fields: [{ key: 'priority', type: 'enum' }] }), 'themes');
  }, /enum field must define values/);
  assert.throws(function() {
    parseFactTableDictionary_(JSON.stringify({ fields: [{ key: 'entry_id', type: 'string' }] }), 'themes');
  }, /reserved/);
  assert.throws(function() {
    parseFactTableDictionary_(JSON.stringify({ fields: [{ key: 'themes_id', type: 'string' }] }), 'themes');
  }, /reserved/);
  assert.throws(function() {
    parseFactTableDictionary_(JSON.stringify({
      fields: [
        { key: 'summary', type: 'string' },
        { key: 'summary', type: 'string' },
      ],
    }), 'themes');
  }, /duplicated/);
});

test('dynamic custom headers add dictionary fields after metadata', function() {
  const dictionary = JSON.stringify({
    fields: [
      { key: 'summary', type: 'string', description: 'Summary.' },
      { key: 'priority', type: 'enum', description: 'Priority.', values: ['low', 'high'] },
    ],
  });

  assert.deepStrictEqual(customExtractionHeaders_('themes', dictionary), [
    'themes_id',
    'entry_id',
    'matched_keywords',
    'created_at',
    'extracted_at',
    'summary',
    'priority',
  ]);
  assert.deepStrictEqual(reservedCustomFactColumns_('themes'), [
    'themes_id',
    'entry_id',
    'matched_keywords',
    'created_at',
    'extracted_at',
  ]);
});

test('custom extraction schema requires all dictionary fields', function() {
  const fields = customFactFields_(JSON.stringify({
    fields: [
      { key: 'summary', type: 'string', description: 'Summary.' },
      { key: 'priority', type: 'enum', description: 'Priority.', values: ['low', 'high'] },
      { key: 'is_recurring', type: 'boolean', description: 'Repeats.' },
    ],
  }), 'themes');
  const schema = customExtractionSchema_(fields);
  const itemSchema = schema.properties.records.items;

  assert.deepStrictEqual(itemSchema.required, ['summary', 'priority', 'is_recurring']);
  assert.deepStrictEqual(itemSchema.properties.priority.enum, ['', 'low', 'high']);
  assert.deepStrictEqual(customFactFieldSchema_(fields[2]).type, ['boolean', 'string']);
});

test('blank fact table dictionary falls back to summary evidence status', function() {
  assert.deepStrictEqual(defaultCustomFactFields_().map(function(field) { return field.key; }), [
    'summary',
    'evidence',
    'status',
  ]);
  assert.deepStrictEqual(customFactFields_('', 'themes').map(function(field) { return field.key; }), [
    'summary',
    'evidence',
    'status',
  ]);
});

test('keywords gate custom extraction unless AI may decide without a hit', function() {
  const design = normalizeDesignRow_({
    extraction_key: 'themes',
    keywords: 'calm, focus',
    ai_decides_without_keyword: false,
  });

  assert.deepStrictEqual(parseKeywords_('calm, focus; calm'), ['calm', 'focus']);
  assert.deepStrictEqual(findMatchedKeywords_('I felt calm today.', design.keywords), ['calm']);
  assert.strictEqual(shouldRunCustomExtraction_(design, 'I felt calm today.'), true);
  assert.strictEqual(shouldRunCustomExtraction_(design, 'A different note.'), false);
  assert.strictEqual(shouldRunCustomExtraction_(
    { extraction_key: 'themes', keywords: '', ai_decides_without_keyword: true },
    'A different note.'
  ), true);
});

test('normalizes custom extraction records with matched keywords', function() {
  const fields = customFactFields_(JSON.stringify({
    fields: [
      { key: 'summary', type: 'string', description: 'Summary.' },
      { key: 'priority', type: 'enum', description: 'Priority.', values: ['low', 'high'] },
      { key: 'is_recurring', type: 'boolean', description: 'Repeats.' },
      { key: 'score', type: 'number', description: 'Score.' },
    ],
  }), 'themes');
  const records = normalizeCustomExtraction_(
    {
      records: [
        { summary: 'Keep using the calmer review loop', priority: 'high', is_recurring: 'true', score: '4' },
        { summary: 'Unknown priority', priority: 'urgent', is_recurring: 'maybe', score: '' },
        { summary: '   ', priority: '', is_recurring: '', score: '' },
      ],
    },
    ['calm', 'review'],
    fields
  );

  assert.deepStrictEqual(records, [
    {
      summary: 'Keep using the calmer review loop',
      priority: 'high',
      is_recurring: true,
      score: 4,
      matched_keywords: 'calm, review',
    },
    {
      summary: 'Unknown priority',
      priority: '',
      is_recurring: '',
      score: '',
      matched_keywords: 'calm, review',
    },
  ]);
});

test('built-in design rows ignore fact table dictionary JSON', function() {
  const row = normalizeDesignRow_({
    extraction_key: 'goals',
    sheet_name: 'Custom Goals',
    fact_table_dictionary: JSON.stringify({ fields: [{ key: 'priority', type: 'string' }] }),
  });

  assert.strictEqual(row.sheet_name, SHEET_NAMES.GOALS);
  assert.strictEqual(row.fact_table_dictionary.indexOf('priority') !== -1, true);
  assert.strictEqual(row.extraction_key, 'goals');
});

test('install plan requires Drive folder and API key setup when no stored key exists', function() {
  const plan = buildVoiceJournalInstallPlan_({}, false);

  assert.strictEqual(plan.isReady, false);
  assert.deepStrictEqual(plan.missingConfigKeys, ['DRIVE_INBOX_FOLDER_ID', 'OPENAI_API_KEY_SETUP']);
  assert.throws(function() {
    assertVoiceJournalInstallReady_(plan);
  }, /DRIVE_INBOX_FOLDER_ID, OPENAI_API_KEY_SETUP/);
});

test('install plan accepts already stored API key and skips blank digest', function() {
  const plan = buildVoiceJournalInstallPlan_(
    {
      DRIVE_INBOX_FOLDER_ID: 'folder_123',
      OPENAI_API_KEY_PROPERTY: 'VOICE_JOURNAL_OPENAI_KEY',
      DIGEST_RECIPIENT_EMAIL: '',
    },
    true
  );

  assert.strictEqual(plan.isReady, true);
  assert.strictEqual(plan.apiKeyPropertyName, 'VOICE_JOURNAL_OPENAI_KEY');
  assert.strictEqual(plan.shouldStoreApiKey, false);
  assert.strictEqual(plan.shouldInstallDigest, false);
});

test('install plan moves setup API key and installs digest when recipient is set', function() {
  const plan = buildVoiceJournalInstallPlan_(
    {
      DRIVE_INBOX_FOLDER_ID: 'folder_123',
      OPENAI_API_KEY_SETUP: 'sk-test',
      DIGEST_RECIPIENT_EMAIL: 'me@example.com',
    },
    false
  );

  assert.strictEqual(getOpenAiApiKeyPropertyName_({}), 'OPENAI_API_KEY');
  assert.strictEqual(plan.isReady, true);
  assert.strictEqual(plan.shouldStoreApiKey, true);
  assert.strictEqual(plan.apiKeyToStore, 'sk-test');
  assert.strictEqual(plan.shouldInstallDigest, true);
});

test('clears temporary API key config cell after storage', function() {
  const calls = [];
  const sheet = {
    getLastRow: function() { return 3; },
    getRange: function(row, column, rowCount, columnCount) {
      if (row === 2 && column === 1 && rowCount === 2 && columnCount === 3) {
        return {
          getValues: function() {
            return [
              ['DRIVE_INBOX_FOLDER_ID', 'folder_123', ''],
              ['OPENAI_API_KEY_SETUP', 'sk-test', ''],
            ];
          },
        };
      }
      return {
        clearContent: function() {
          calls.push(['clearContent', row, column]);
          return this;
        },
        setNote: function(note) {
          calls.push(['setNote', row, column, note]);
          return this;
        },
      };
    },
  };

  assert.strictEqual(clearConfigValue_(sheet, 'OPENAI_API_KEY_SETUP', 'stored'), true);
  assert.deepStrictEqual(calls, [
    ['clearContent', 3, 2],
    ['setNote', 3, 2, 'stored'],
  ]);
});

test('install summary reports digest installed or skipped', function() {
  assert.match(
    buildVoiceJournalInstallSummary_({ pollerInstalled: true, apiKeyStored: true, digestInstalled: true }),
    /API key stored.*Digest trigger installed/
  );
  assert.match(
    buildVoiceJournalInstallSummary_({ pollerInstalled: true, apiKeyStored: false, digestInstalled: false }),
    /already stored.*Digest trigger skipped/
  );
});
