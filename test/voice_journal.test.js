const assert = require('assert');

const {
  buildDashboardTodoRows_,
  buildRecentMoodCounts_,
  dashboardCheckboxStatus_,
  moodCountsToRows_,
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
