const assert = require('assert');
const journal = require('../src/Code.js');

function testHeaders() {
  assert.strictEqual(journal.HEADERS.Entries.includes('uploaded_at'), true);
  assert.strictEqual(journal.HEADERS.Entries.includes('recorded_at'), false);
  assert.strictEqual(journal.HEADERS.Goals.includes('review_status'), false);
  assert.strictEqual(journal.HEADERS['To-Dos'].includes('review_status'), false);
  assert.strictEqual(journal.HEADERS.Thoughts.includes('review_status'), false);
  assert.deepStrictEqual(journal.HEADERS.Reminders, [
    'reminder_id',
    'source_type',
    'source_id',
    'entry_id',
    'reminder_text',
    'next_reminder_at',
    'active_until',
    'frequency_days',
    'reminder_count',
    'last_sent_at',
    'status',
  ]);
}

function testExtractionNormalization() {
  const normalized = journal.normalizeExtraction_({
    goals: [
      {
        summary: '  Publish the essay ',
        status: '',
        active_until: '2026-06-15',
        reminder_frequency_days: 7,
      },
      { summary: '', status: 'Active' },
    ],
    todos: [
      {
        task: 'Call Sam',
        status: '',
        due_date_optional: '',
        reminder_frequency_days: 3,
      },
    ],
    thoughts: [
      {
        summary: ' This idea may become a project. ',
        moods: ['Curious', 'reflective', 'not-a-real-mood', 'curious'],
      },
    ],
  });

  assert.deepStrictEqual(normalized, {
    goals: [
      {
        summary: 'Publish the essay',
        status: '',
        active_until: '2026-06-15',
        reminder_frequency_days: 7,
      },
    ],
    todos: [
      {
        task: 'Call Sam',
        status: '',
        due_date_optional: '',
        reminder_frequency_days: 3,
      },
    ],
    thoughts: [
      {
        summary: 'This idea may become a project.',
        moods: ['curious', 'reflective'],
      },
    ],
  });
}

function testRetryRule() {
  assert.strictEqual(journal.shouldRetryEntry_({ status: 'Failed', retry_count: 0 }, 3), true);
  assert.strictEqual(journal.shouldRetryEntry_({ status: 'Failed', retry_count: 3 }, 3), false);
  assert.strictEqual(journal.shouldRetryEntry_({ status: 'Processed', retry_count: 0 }, 3), false);
}

function testDigestFrequency() {
  const monday = new Date('2026-05-18T12:00:00Z');
  const tuesday = new Date('2026-05-19T12:00:00Z');
  const saturday = new Date('2026-05-23T12:00:00Z');

  assert.strictEqual(journal.shouldSendDigestOnDate_(saturday, 7), true);
  assert.strictEqual(journal.shouldSendDigestOnDate_(monday, 5), true);
  assert.strictEqual(journal.shouldSendDigestOnDate_(saturday, 5), false);
  assert.strictEqual(journal.shouldSendDigestOnDate_(monday, 3), true);
  assert.strictEqual(journal.shouldSendDigestOnDate_(tuesday, 3), false);
}

function testDoneStatuses() {
  assert.strictEqual(journal.isDoneStatus_('Done'), true);
  assert.strictEqual(journal.isDoneStatus_('Complete'), true);
  assert.strictEqual(journal.isDoneStatus_('Archived'), true);
  assert.strictEqual(journal.isDoneStatus_('Open'), false);
}

function testSchemaShape() {
  const schema = journal.extractionSchema_();
  assert.deepStrictEqual(schema.required, ['goals', 'todos', 'thoughts']);
  assert.strictEqual(schema.additionalProperties, false);
  assert.strictEqual(schema.properties.goals.items.required.includes('active_until'), true);
  assert.strictEqual(schema.properties.goals.items.required.includes('reminder_frequency_days'), true);
  assert.strictEqual(schema.properties.todos.items.required.includes('due_date_optional'), true);
  assert.strictEqual(schema.properties.todos.items.required.includes('reminder_frequency_days'), true);
  assert.strictEqual(schema.properties.thoughts.items.required.includes('moods'), true);
  assert.deepStrictEqual(schema.properties.thoughts.items.properties.moods.items.enum, journal.MOOD_TAGS);
}

testHeaders();
testExtractionNormalization();
testRetryRule();
testDigestFrequency();
testDoneStatuses();
testSchemaShape();

console.log('All voice journal tests passed.');
