const assert = require('assert');

const {
  HEADERS,
  SHEET_NAMES,
  buildDigestFromRows_,
  buildDashboardGoalRows_,
  buildDashboardTodoRows_,
  buildRecentMoodCounts_,
  buildSearchIndexRowsForEntry_,
  buildSearchAppStateFromRows_,
  buildSearchResultObjects_,
  buildSearchResultRows_,
  buildThoughtContextByEntryId_,
  chunkTranscript_,
  cosineSimilarity_,
  dashboardGoalActionStatus_,
  dashboardCheckboxStatus_,
  fallbackDigestIntro_,
  moodCountsToRows_,
  normalizeSearchOptions_,
  rankSearchRows_,
  transcriptHash_,
  validateSearchQuery_,
} = require('../src/Code.js');

function test(name, fn) {
  try {
    fn();
    console.log('ok - ' + name);
  } catch (error) {
    console.error('not ok - ' + name);
    throw error;
  }
}

test('filters active dashboard to-dos', function() {
  const rows = [
    { values: { todo_id: 'todo_1', task: 'Call Sam', status: 'Open', due_date_optional: '2026-05-20', created_at: '2026-05-18' } },
    { values: { todo_id: 'todo_2', task: 'File taxes', status: 'Done', due_date_optional: '', created_at: '2026-05-17' } },
    { values: { todo_id: 'todo_3', task: 'Draft outline', status: 'Active', due_date_optional: '', created_at: '2026-05-16' } },
    { values: { todo_id: 'todo_4', task: 'Old thing', status: 'Archived', due_date_optional: '', created_at: '2026-05-15' } },
  ];

  const dashboardRows = buildDashboardTodoRows_(rows);

  assert.strictEqual(dashboardRows.length, 2);
  assert.deepStrictEqual(dashboardRows.map(function(row) { return row[5]; }), ['todo_1', 'todo_3']);
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

test('ignores thoughts outside the mood lookback window', function() {
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

test('maps dashboard checkbox edits to todo statuses', function() {
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
      related_thoughts: 'A thought',
      moods: 'focused',
    },
  ]);

  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0][0], 1);
  assert.strictEqual(rows[0][2], 'Exact transcript quote.');
  assert.strictEqual(rows[0][4], 'entry_1');
  assert.strictEqual(rows[0][7], 'A thought');
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
      related_thoughts: 'A thought',
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

test('search index enrichment joins thoughts by entry id', function() {
  const context = buildThoughtContextByEntryId_([
    { values: { entry_id: 'entry_1', summary: 'Prototype felt promising', moods: 'focused, curious' } },
    { values: { entry_id: 'entry_1', summary: 'Need a calmer review loop', moods: 'calm, focused' } },
    { values: { entry_id: 'entry_2', summary: 'Different entry', moods: 'tired' } },
  ]);

  assert.strictEqual(context.entry_1.related_thoughts, 'Prototype felt promising; Need a calmer review loop');
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
    { related_thoughts: 'Search matters', moods: 'focused' },
    { EMBEDDING_MODEL: 'test-embedding', EMBEDDING_DIMENSIONS: 2, SEARCH_CHUNK_TARGET_CHARS: 1200, SEARCH_CHUNK_OVERLAP_CHARS: 200 },
    new Date('2026-05-20T12:00:00Z'),
    function() { return [1, 0]; }
  );

  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0].entry_id, 'entry_1');
  assert.strictEqual(rows[0].related_thoughts, 'Search matters');
  assert.strictEqual(rows[0].moods, 'focused');
  assert.strictEqual(rows[0].embedding_json, JSON.stringify([1, 0]));
  assert.strictEqual(rows[0].index_status, 'Indexed');
});

test('digest reads reminder state from goals and to-dos', function() {
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
          todo_id: 'todo_1',
          task: 'Call Sam',
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
          todo_id: 'todo_2',
          task: 'Prep report',
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
      { values: { thought_id: 'thought_1', summary: 'Felt focused during planning', moods: 'focused', created_at: '2026-05-19T09:00:00Z' } },
      { values: { thought_id: 'thought_2', summary: 'Old note', moods: 'tired', created_at: '2026-04-01T09:00:00Z' } },
    ],
    { DIGEST_LOOKBACK_DAYS: 7 },
    now,
    function() { throw new Error('intro failed'); }
  );

  assert.match(digest.body, /Today's digest includes 2 active goals, 2 to-dos, and 1 recent reflections\./);
  assert.match(digest.body, /__Goals__\n2 active goals are included, with 1 currently reminder-due\./);
  assert.match(digest.body, /__To-Dos__\n2 to-dos are included, with 1 currently reminder-due\./);
  assert.match(digest.body, /__Thoughts__\n1 recent thoughts are included\. The most common mood tags are focused\./);
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

  assert.strictEqual(digest.todoUpdates.length, 1);
  assert.strictEqual(digest.todoUpdates[0].rowNumber, 2);
  assert.strictEqual(digest.todoUpdates[0].values.reminder_count, 2);
  assert.strictEqual(digest.todoUpdates[0].values.next_reminder_at.toISOString(), '2026-05-23T12:00:00.000Z');
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
        todosSummary: 'No to-dos are currently included.',
        thoughtsSummary: 'No recent thoughts are currently included.',
      };
    }
  );

  assert.match(digest.htmlBody, /<strong><u>Goals<\/u><\/strong>/);
  assert.match(digest.htmlBody, /<strong><u>To-Dos<\/u><\/strong>/);
  assert.match(digest.htmlBody, /<strong><u>Thoughts<\/u><\/strong>/);
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
    todos: [{ task: 'Task' }, { task: 'Second task' }],
    thoughts: [{ summary: 'Thought', moods: 'calm, focused' }],
    moodCounts: { calm: 1, focused: 1 },
  });

  assert.match(intro, /1 active goals, 2 to-dos, and 1 recent reflections/);
  assert.match(intro, /calm and focused|focused and calm/);
});

test('reminders table is no longer part of the configured schema', function() {
  assert.strictEqual(SHEET_NAMES.REMINDERS, undefined);
  assert.strictEqual(HEADERS.Reminders, undefined);
  assert.ok(HEADERS.Goals.indexOf('next_reminder_at') !== -1);
  assert.ok(HEADERS['To-Dos'].indexOf('last_reminded_at') !== -1);
});
