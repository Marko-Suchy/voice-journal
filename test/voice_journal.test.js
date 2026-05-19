const assert = require('assert');

const {
  buildDashboardGoalRows_,
  buildDashboardTodoRows_,
  buildRecentMoodCounts_,
  dashboardGoalActionStatus_,
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
