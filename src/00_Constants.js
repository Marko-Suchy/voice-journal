/**
 * Constants
 */

/**
 * Voice Memo Journal System
 *
 * Copy all src/*.js files into a Google Apps Script project bound to your Google Sheet.
 * Fill the required Config values in the copied workbook, then run installVoiceJournal().
 */

const SHEET_NAMES = {
  DASHBOARD: 'Dashboard',
  ENTRIES: 'Entries',
  GOALS: 'Goals',
  FOLLOW_UPS: 'Follow Ups',
  REFLECTIONS: 'Reflections',
  SEARCH: 'Search',
  SEARCH_INDEX: 'Search Index',
  CONFIG: 'Config',
  DESIGN: 'Design',
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

const DEFAULT_EXTRACTION_PROMPT_TEXT = [
  'You extract structured journal records from personal voice memo transcripts.',
  'Return only JSON matching the requested schema.',
  'Extract goals and follow ups from explicit markers and reasonable contextual clues.',
  'Always consider whether the entry contains one or more big ideas worth saving as reflections, even without the word reflection.',
  'For goals, infer active_until as an ISO date and reminder_frequency_days from context; if unclear, use 30 days from the memo date and 7 days.',
  'For follow ups, infer due_date_optional as an ISO date when possible and reminder_frequency_days from urgency; if unclear, use an empty due date and 3 days.',
  'Reflection moods must come only from this list: ' + MOOD_TAGS.join(', ') + '.',
  'Do not invent unrelated follow ups, goals, or dates; make useful conservative guesses only when the transcript implies intent.',
].join(' ');

const DONE_STATUSES = ['done', 'complete', 'completed', 'archived'];

const DASHBOARD = {
  FOLLOW_UP_START_ROW: 5,
  FOLLOW_UP_START_COLUMN: 1,
  FOLLOW_UP_COLUMN_COUNT: 6,
  FOLLOW_UP_ID_COLUMN: 6,
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

const SEARCH = {
  QUERY_CELL: 'B2',
  MAX_RESULTS_CELL: 'B3',
  LAST_SEARCH_CELL: 'B4',
  RESULT_HEADER_ROW: 6,
  RESULT_START_ROW: 7,
  RESULT_COLUMN_COUNT: 9,
};

const DESIGN = {
  ACTION_COLUMN: 9,
  FIRST_DATA_ROW: 2,
  ACTIONS: ['', 'Create Sheet', 'Run Backfill'],
  ACTIVE_STATUS: 'Active',
  BACKFILL_BATCH_SIZE: 10,
  BUILT_IN_KEYS: ['follow_ups', 'goals', 'reflections'],
};

const CONFIG_BUTTONS = [
  {
    title: 'Install / Repair',
    script: 'installVoiceJournal',
    row: 2,
    column: 5,
    color: '#1a73e8',
  },
  {
    title: 'Update Dashboard',
    script: 'refreshDashboardFromConfigButton',
    row: 5,
    column: 5,
    color: '#188038',
  },
];

const SEARCH_BUTTONS = [
  {
    title: 'Send Query',
    script: 'searchJournalFromMenu',
    row: 2,
    column: 4,
    color: '#1a73e8',
  },
];

const HEADERS = {
  Entries: [
    'entry_id',
    'drive_file_id',
    'audio_url',
    'uploaded_at',
    'processed_at',
    'transcript',
    'has_goal',
    'has_follow_up',
    'has_reflection',
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
  'Follow Ups': [
    'follow_up_id',
    'entry_id',
    'summary',
    'status',
    'created_at',
    'due_date_optional',
    'reminder_frequency_days',
    'next_reminder_at',
    'last_reminded_at',
    'reminder_count',
  ],
  Reflections: [
    'reflection_id',
    'entry_id',
    'summary',
    'moods',
    'created_at',
  ],
  Search: [
    'Rank',
    'Score',
    'Quote',
    'Context',
    'entry_id',
    'uploaded_at',
    'Audio',
    'Related Reflections',
    'Moods',
  ],
  'Search Index': [
    'chunk_id',
    'entry_id',
    'source_row',
    'uploaded_at',
    'processed_at',
    'audio_url',
    'chunk_start_char',
    'chunk_end_char',
    'quote_text',
    'context_before',
    'context_after',
    'related_reflections',
    'moods',
    'embedding_json',
    'embedding_model',
    'embedding_dimensions',
    'transcript_hash',
    'indexed_at',
    'index_status',
    'index_error',
  ],
  Config: ['key', 'value', 'notes'],
  Design: [
    'extraction_key',
    'display_name',
    'sheet_name',
    'description',
    'fact_table_dictionary',
    'keywords',
    'ai_decides_without_keyword',
    'status',
    'action',
    'backfill_cursor',
    'last_backfilled_at',
    'last_error',
  ],
};

const DEFAULT_DESIGN_ROWS = [
  {
    extraction_key: 'follow_ups',
    display_name: 'Follow Ups',
    sheet_name: SHEET_NAMES.FOLLOW_UPS,
    description: 'Actionable commitments, errands, and next steps that should remain visible until completed or archived.',
    fact_table_dictionary: '',
    keywords: 'follow up, follow-up, remind me, remember to, need to, should, task, action item',
    ai_decides_without_keyword: true,
    status: 'Active',
  },
  {
    extraction_key: 'goals',
    display_name: 'Goals',
    sheet_name: SHEET_NAMES.GOALS,
    description: 'Longer-running aims, habits, outcomes, or intentions that should stay active across entries.',
    fact_table_dictionary: '',
    keywords: 'goal, want to, working toward, aiming to, by the end, this month, this year',
    ai_decides_without_keyword: true,
    status: 'Active',
  },
  {
    extraction_key: 'reflections',
    display_name: 'Reflections',
    sheet_name: SHEET_NAMES.REFLECTIONS,
    description: 'Meaningful observations, emotional notes, patterns, or larger ideas worth preserving with mood tags.',
    fact_table_dictionary: '',
    keywords: 'reflect, reflection, realized, noticed, felt, learned, thinking about, idea',
    ai_decides_without_keyword: true,
    status: 'Active',
  },
];

const DEFAULT_CONFIG = [
  ['DRIVE_INBOX_FOLDER_ID', '', 'Google Drive folder ID where phone voice memos are uploaded.'],
  ['OPENAI_API_KEY_PROPERTY', 'OPENAI_API_KEY', 'Script Properties key that stores the OpenAI API key. Do not put the API key in this sheet.'],
  ['OPENAI_API_KEY_SETUP', '', 'Temporary setup field. Paste your OpenAI API key here before running installVoiceJournal(); the installer stores it in Script Properties and clears this cell.'],
  ['TRANSCRIPTION_MODEL', 'gpt-4o-mini-transcribe', 'OpenAI audio transcription model.'],
  ['EXTRACTION_MODEL', 'gpt-5.2', 'OpenAI text model used for structured extraction.'],
  ['MAX_RETRIES', '3', 'Maximum processing attempts per memo.'],
  ['DEFAULT_FOLLOW_UP_STATUS', 'Open', 'Initial status for new follow up rows.'],
  ['DEFAULT_GOAL_STATUS', 'Active', 'Initial status for new goal rows.'],
  ['DEFAULT_GOAL_ACTIVE_DAYS', '30', 'Default number of days a goal remains active when the memo is unclear.'],
  ['DEFAULT_GOAL_REMINDER_FREQUENCY_DAYS', '7', 'Default reminder frequency for goals when the memo is unclear.'],
  ['DEFAULT_FOLLOW_UP_REMINDER_FREQUENCY_DAYS', '3', 'Default reminder frequency for follow ups when the memo is unclear.'],
  ['DIGEST_RECIPIENT_EMAIL', '', 'Required. Daily digest recipient email address.'],
  ['DIGEST_FREQUENCY_PER_WEEK', '7', 'How many days per week to send the digest. 7 = daily, 5 = weekdays, 3 = Mon/Wed/Fri.'],
  ['DIGEST_SEND_HOUR', '18', 'Hour of day, 0-23, for the daily digest trigger. Default 18 = 6 PM.'],
  ['DIGEST_LOOKBACK_DAYS', '7', 'Number of days of recent reflections and follow ups to include in the digest.'],
  ['MOOD_TAGS', MOOD_TAGS.join(','), 'Allowed mood tags for extracted reflections.'],
  ['EMBEDDING_MODEL', 'text-embedding-3-small', 'OpenAI embedding model used for journal semantic search.'],
  ['EMBEDDING_DIMENSIONS', '512', 'Embedding dimensions for Search Index rows.'],
  ['SEARCH_MAX_RESULTS', '12', 'Default number of Search results to display.'],
  ['SEARCH_MIN_SCORE', '0.20', 'Minimum cosine similarity score shown in Search results.'],
  ['SEARCH_CHUNK_TARGET_CHARS', '1200', 'Approximate target transcript characters per search chunk.'],
  ['SEARCH_CHUNK_OVERLAP_CHARS', '200', 'Approximate overlapping characters between long search chunks.'],
  ['EXTRACTION_PROMPT', DEFAULT_EXTRACTION_PROMPT_TEXT, 'Prompt used to extract goals, follow ups, and big-idea reflections from transcripts.'],
];
