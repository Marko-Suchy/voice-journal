/**
 * Dashboard
 */

function refreshDashboardNow() {
  setupVoiceJournalSheet();
  refreshDashboard_(SpreadsheetApp.getActiveSpreadsheet(), new Date());
}

function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  const spreadsheet = e.source || SpreadsheetApp.getActiveSpreadsheet();
  if (sheet.getName() === SHEET_NAMES.DESIGN) {
    handleDesignActionEdit_(spreadsheet, sheet, e);
    return;
  }

  if (sheet.getName() !== SHEET_NAMES.DASHBOARD) {
    return;
  }

  if (handleDashboardFollowUpEdit_(spreadsheet, sheet, e)) {
    return;
  }

  handleDashboardGoalEdit_(spreadsheet, sheet, e);
}

function handleDashboardFollowUpEdit_(spreadsheet, sheet, e) {
  if (e.range.getColumn() !== DASHBOARD.FOLLOW_UP_START_COLUMN || e.range.getRow() < DASHBOARD.FOLLOW_UP_START_ROW) {
    return false;
  }

  const followUpId = sheet.getRange(e.range.getRow(), DASHBOARD.FOLLOW_UP_ID_COLUMN).getValue();
  if (!followUpId) {
    return false;
  }

  const nextStatus = dashboardCheckboxStatus_(e.value);
  updateFollowUpStatusById_(spreadsheet.getSheetByName(SHEET_NAMES.FOLLOW_UPS), followUpId, nextStatus);
  sheet.getRange(e.range.getRow(), 3).setValue(nextStatus);
  return true;
}

function handleDashboardGoalEdit_(spreadsheet, sheet, e) {
  if (e.range.getColumn() !== DASHBOARD.GOAL_START_COLUMN || e.range.getRow() < DASHBOARD.GOAL_START_ROW) {
    return false;
  }

  const nextStatus = dashboardGoalActionStatus_(e.value);
  if (!nextStatus) {
    return false;
  }

  const goalId = sheet.getRange(e.range.getRow(), DASHBOARD.GOAL_ID_COLUMN).getValue();
  if (!goalId) {
    return false;
  }

  updateGoalStatusById_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS), goalId, nextStatus);
  sheet.getRange(e.range.getRow(), DASHBOARD.GOAL_START_COLUMN + 2).setValue(nextStatus);
  return true;
}

function refreshDashboard_(spreadsheet, now) {
  const dashboardSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.DASHBOARD);
  const followUpRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.FOLLOW_UPS));
  const goalRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS));
  const reflectionRows = getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.REFLECTIONS));
  const dashboardFollowUps = buildDashboardFollowUpRows_(followUpRows);
  const dashboardGoals = buildDashboardGoalRows_(goalRows);
  const moodCounts = buildRecentMoodCounts_(reflectionRows, now, DASHBOARD.MOOD_LOOKBACK_DAYS);
  const moodRows = moodCountsToRows_(moodCounts);

  removeCharts_(dashboardSheet);
  dashboardSheet.clear();
  dashboardSheet.setHiddenGridlines(true);
  dashboardSheet.setFrozenRows(0);
  dashboardSheet.getRange('A1').setValue('Voice Journal Dashboard').setFontSize(18).setFontWeight('bold');
  dashboardSheet.getRange('A2').setValue('Last refreshed: ' + formatDateTimeForDashboard_(now));

  renderDashboardFollowUps_(dashboardSheet, dashboardFollowUps);
  renderDashboardGoals_(dashboardSheet, dashboardGoals);
  renderDashboardMoods_(dashboardSheet, moodRows);
}

function renderDashboardFollowUps_(sheet, dashboardFollowUps) {
  const headerRange = sheet.getRange(4, DASHBOARD.FOLLOW_UP_START_COLUMN, 1, DASHBOARD.FOLLOW_UP_COLUMN_COUNT);
  headerRange
    .setValues([['Done', 'Follow Up', 'Status', 'Due Date', 'Created', 'follow_up_id']])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.getRange('A3').setValue('Active Follow Ups').setFontWeight('bold');

  if (dashboardFollowUps.length === 0) {
    sheet.getRange(DASHBOARD.FOLLOW_UP_START_ROW, DASHBOARD.FOLLOW_UP_START_COLUMN).setValue('No active follow ups.');
  } else {
    const range = sheet.getRange(
      DASHBOARD.FOLLOW_UP_START_ROW,
      DASHBOARD.FOLLOW_UP_START_COLUMN,
      dashboardFollowUps.length,
      DASHBOARD.FOLLOW_UP_COLUMN_COUNT
    );
    range.setValues(dashboardFollowUps);
    sheet.getRange(DASHBOARD.FOLLOW_UP_START_ROW, DASHBOARD.FOLLOW_UP_START_COLUMN, dashboardFollowUps.length, 1).insertCheckboxes();
    sheet.getRange(DASHBOARD.FOLLOW_UP_START_ROW, 4, dashboardFollowUps.length, 2).setNumberFormat('yyyy-mm-dd');
  }

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 360);
  sheet.getRange(1, 2, sheet.getMaxRows(), 1).setWrap(true);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.hideColumns(DASHBOARD.FOLLOW_UP_ID_COLUMN);
}

function renderDashboardGoals_(sheet, dashboardGoals) {
  const headerRange = sheet.getRange(DASHBOARD.GOAL_HEADER_ROW, DASHBOARD.GOAL_START_COLUMN, 1, DASHBOARD.GOAL_COLUMN_COUNT);
  headerRange
    .setValues([['Action', 'Goal', 'Status', 'Active Until', 'Created', 'goal_id']])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.getRange(DASHBOARD.GOAL_TITLE_ROW, DASHBOARD.GOAL_START_COLUMN).setValue('Active Goals').setFontWeight('bold');

  if (dashboardGoals.length === 0) {
    sheet.getRange(DASHBOARD.GOAL_START_ROW, DASHBOARD.GOAL_START_COLUMN).setValue('No active goals.');
  } else {
    const range = sheet.getRange(
      DASHBOARD.GOAL_START_ROW,
      DASHBOARD.GOAL_START_COLUMN,
      dashboardGoals.length,
      DASHBOARD.GOAL_COLUMN_COUNT
    );
    range.setValues(dashboardGoals);
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Complete', 'Archive'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(DASHBOARD.GOAL_START_ROW, DASHBOARD.GOAL_START_COLUMN, dashboardGoals.length, 1).setDataValidation(validation);
    sheet.getRange(DASHBOARD.GOAL_START_ROW, DASHBOARD.GOAL_START_COLUMN + 3, dashboardGoals.length, 2).setNumberFormat('yyyy-mm-dd');
  }

  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN, 100);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 1, 340);
  sheet.getRange(1, DASHBOARD.GOAL_START_COLUMN + 1, sheet.getMaxRows(), 1).setWrap(true);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 2, 90);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 3, 120);
  sheet.setColumnWidth(DASHBOARD.GOAL_START_COLUMN + 4, 110);
  sheet.hideColumns(DASHBOARD.GOAL_ID_COLUMN);
}

function renderDashboardMoods_(sheet, moodRows) {
  sheet.getRange(3, DASHBOARD.MOOD_START_COLUMN).setValue('Moods - Last 7 Days').setFontWeight('bold');
  sheet.getRange(
    4,
    DASHBOARD.MOOD_START_COLUMN,
    1,
    DASHBOARD.MOOD_COLUMN_COUNT
  ).setValues([['Mood', 'Count']]).setFontWeight('bold').setBackground('#f1f3f4');

  if (moodRows.length === 0) {
    sheet.getRange(DASHBOARD.MOOD_START_ROW, DASHBOARD.MOOD_START_COLUMN).setValue('No moods in the last 7 days.');
    return;
  }

  const moodRange = sheet.getRange(
    DASHBOARD.MOOD_START_ROW,
    DASHBOARD.MOOD_START_COLUMN,
    moodRows.length,
    DASHBOARD.MOOD_COLUMN_COUNT
  );
  moodRange.setValues(moodRows);
  sheet.setColumnWidth(DASHBOARD.MOOD_START_COLUMN, 140);
  sheet.setColumnWidth(DASHBOARD.MOOD_START_COLUMN + 1, 80);

  const chart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange(4, DASHBOARD.MOOD_START_COLUMN, moodRows.length + 1, DASHBOARD.MOOD_COLUMN_COUNT))
    .setPosition(4, DASHBOARD.MOOD_START_COLUMN + 3, 0, 0)
    .setOption('title', 'Moods - Last 7 Days')
    .setOption('pieHole', 0.35)
    .build();
  sheet.insertChart(chart);
}

function removeCharts_(sheet) {
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
}

function buildDashboardFollowUpRows_(rows) {
  return normalizeRowsForValues_(rows)
    .filter(function(followUp) {
      return followUp.follow_up_id && followUp.summary && !isDoneStatus_(followUp.status);
    })
    .map(function(followUp) {
      return [
        false,
        followUp.summary,
        followUp.status || 'Open',
        parseDate_(followUp.due_date_optional) || '',
        parseDate_(followUp.created_at) || '',
        followUp.follow_up_id,
      ];
    });
}

function buildDashboardGoalRows_(rows) {
  return normalizeRowsForValues_(rows)
    .filter(function(goal) {
      return goal.goal_id && goal.summary && !isDoneStatus_(goal.status);
    })
    .map(function(goal) {
      return [
        '',
        goal.summary,
        goal.status || 'Active',
        parseDate_(goal.active_until) || '',
        parseDate_(goal.created_at) || '',
        goal.goal_id,
      ];
    });
}

function buildRecentMoodCounts_(rows, now, lookbackDays) {
  const since = addDays_(now, -lookbackDays);
  const counts = {};
  normalizeRowsForValues_(rows).forEach(function(reflection) {
    const createdAt = parseDate_(reflection.created_at);
    if (!isRecent_(createdAt, since)) {
      return;
    }

    String(reflection.moods || '').split(',').forEach(function(rawMood) {
      const mood = rawMood.trim().toLowerCase();
      if (mood) {
        counts[mood] = (counts[mood] || 0) + 1;
      }
    });
  });
  return counts;
}

function moodCountsToRows_(counts) {
  return Object.keys(counts || {})
    .sort(function(a, b) {
      if (counts[b] !== counts[a]) {
        return counts[b] - counts[a];
      }
      return a < b ? -1 : a > b ? 1 : 0;
    })
    .map(function(mood) {
      return [mood, counts[mood]];
    });
}

function dashboardCheckboxStatus_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE' ? 'Done' : 'Open';
}

function dashboardGoalActionStatus_(value) {
  const action = String(value || '').trim().toLowerCase();
  if (action === 'complete') {
    return 'Complete';
  }
  if (action === 'archive') {
    return 'Archived';
  }
  return '';
}

function updateFollowUpStatusById_(followUpsSheet, followUpId, status) {
  const headers = getHeaders_(followUpsSheet);
  const idColumnIndex = headers.indexOf('follow_up_id');
  const statusColumnIndex = headers.indexOf('status');
  if (idColumnIndex === -1 || statusColumnIndex === -1 || followUpsSheet.getLastRow() < 2) {
    return false;
  }

  const ids = followUpsSheet.getRange(2, idColumnIndex + 1, followUpsSheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === followUpId) {
      followUpsSheet.getRange(i + 2, statusColumnIndex + 1).setValue(status);
      return true;
    }
  }
  return false;
}

function updateGoalStatusById_(goalsSheet, goalId, status) {
  const headers = getHeaders_(goalsSheet);
  const idColumnIndex = headers.indexOf('goal_id');
  const statusColumnIndex = headers.indexOf('status');
  if (idColumnIndex === -1 || statusColumnIndex === -1 || goalsSheet.getLastRow() < 2) {
    return false;
  }

  const ids = goalsSheet.getRange(2, idColumnIndex + 1, goalsSheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === goalId) {
      goalsSheet.getRange(i + 2, statusColumnIndex + 1).setValue(status);
      return true;
    }
  }
  return false;
}
