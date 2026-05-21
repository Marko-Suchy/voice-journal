const assert = require('assert');

const {
  HEADERS,
  SHEET_NAMES,
  buildDigestFromRows_,
  buildDashboardGoalRows_,
  buildDashboardTodoRows_,
  buildRecentMoodCounts_,
  dashboardGoalActionStatus_,
  dashboardCheckboxStatus_,
  fallbackDigestIntro_,
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
