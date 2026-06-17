/**
 * Digest
 */

function installDailyDigestTrigger() {
  setupVoiceJournalSheet();
  removeDailyDigestTrigger();
  const config = readConfig_(SpreadsheetApp.getActiveSpreadsheet());
  const sendHour = clampInteger_(parseInteger_(config.DIGEST_SEND_HOUR, 18), 0, 23);
  ScriptApp.newTrigger('sendScheduledDailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(sendHour)
    .create();
}

function removeDailyDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendScheduledDailyDigest') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function sendDailyDigestNow() {
  return sendDigest_(true);
}

function sendScheduledDailyDigest() {
  return sendDigest_(false);
}

function sendDigest_(forceSend) {
  setupVoiceJournalSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(spreadsheet);
  const recipient = String(config.DIGEST_RECIPIENT_EMAIL || '').trim();
  if (!recipient) {
    throw new Error('Set DIGEST_RECIPIENT_EMAIL in the Config tab before sending the digest.');
  }

  const now = new Date();
  if (!forceSend && !shouldSendDigestOnDate_(now, config.DIGEST_FREQUENCY_PER_WEEK)) {
    return 'Skipped: today is not an enabled digest day.';
  }

  const digest = buildDigest_(spreadsheet, config, now);
  MailApp.sendEmail({
    to: recipient,
    subject: 'Voice Journal Digest - ' + formatDateForDigest_(now),
    body: digest.body,
    htmlBody: digest.htmlBody,
  });
  applyDigestStateUpdates_(spreadsheet, digest);
  return 'Sent digest to ' + recipient + '.';
}

function buildDigest_(spreadsheet, config, now) {
  return buildDigestFromRows_(
    getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS)),
    getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.FOLLOW_UPS)),
    getRowsAsObjects_(spreadsheet.getSheetByName(SHEET_NAMES.REFLECTIONS)),
    config,
    now,
    function(summary) {
      return generateDigestNarrative_(config, summary);
    }
  );
}

function buildDigestFromRows_(goalRows, followUpRows, reflectionRows, config, now, narrativeGenerator) {
  const lookbackDays = parseInteger_(config.DIGEST_LOOKBACK_DAYS, 7);
  const since = addDays_(now, -lookbackDays);
  const goalUpdates = [];
  const followUpUpdates = [];

  const activeGoals = normalizeRowsForValues_(goalRows)
    .map(function(goal, index) {
      return { values: goal, rowNumber: goalRows[index] && goalRows[index].rowNumber };
    })
    .filter(function(row) {
      const activeUntil = parseDate_(row.values.active_until);
      return !isDoneStatus_(row.values.status)
        && (!activeUntil || startOfDay_(activeUntil).getTime() >= startOfDay_(now).getTime());
    })
    .map(function(row) {
      const due = isReminderDue_(row.values, now);
      if (due) {
        goalUpdates.push(buildDigestStateUpdate_(row, now));
      }
      return copyObject_(row.values, { reminder_due: due });
    });

  const openFollowUps = normalizeRowsForValues_(followUpRows)
    .map(function(followUp, index) {
      return { values: followUp, rowNumber: followUpRows[index] && followUpRows[index].rowNumber };
    })
    .filter(function(row) {
      const followUp = row.values;
      const createdAt = parseDate_(followUp.created_at);
      const dueDate = parseDate_(followUp.due_date_optional);
      return !isDoneStatus_(followUp.status)
        && (isRecent_(createdAt, since) || isDueSoon_(dueDate, now) || isReminderDue_(followUp, now));
    })
    .map(function(row) {
      const due = isReminderDue_(row.values, now);
      if (due) {
        followUpUpdates.push(buildDigestStateUpdate_(row, now));
      }
      return copyObject_(row.values, { reminder_due: due });
    });

  const recentReflections = normalizeRowsForValues_(reflectionRows)
    .filter(function(reflection) {
      return isRecent_(parseDate_(reflection.created_at), since);
    });

  const summary = {
    goals: activeGoals,
    followUps: openFollowUps,
    reflections: recentReflections,
    moodCounts: countMoods_(recentReflections),
  };
  const narrative = buildDigestNarrative_(summary, narrativeGenerator);

  return {
    body: renderDigestTextBody_(activeGoals, openFollowUps, recentReflections, narrative, now, lookbackDays),
    htmlBody: renderDigestHtmlBody_(activeGoals, openFollowUps, recentReflections, narrative, now, lookbackDays),
    goalUpdates: goalUpdates,
    followUpUpdates: followUpUpdates,
  };
}

function renderDigestTextBody_(goals, followUps, reflections, narrative, now, lookbackDays) {
  const lines = [];
  lines.push('Voice Journal Digest');
  lines.push(formatDateForDigest_(now));
  lines.push('');
  if (narrative.intro) {
    lines.push(narrative.intro);
    lines.push('');
  }

  lines.push('__Goals__');
  if (narrative.goalsSummary) {
    lines.push(narrative.goalsSummary);
  }
  if (goals.length === 0) {
    lines.push('  1. No active goals.');
  } else {
    goals.slice(0, 12).forEach(function(goal, index) {
      const activeUntil = goal.active_until ? ' through ' + formatDateForDigest_(goal.active_until) : '';
      const dueNow = goal.reminder_due ? ' [reminder due]' : '';
      lines.push('  ' + (index + 1) + '. ' + goal.summary + activeUntil + dueNow);
    });
  }
  lines.push('');

  lines.push('__Follow Ups__');
  if (narrative.followUpsSummary) {
    lines.push(narrative.followUpsSummary);
  }
  if (followUps.length === 0) {
    lines.push('  1. No open follow ups needing attention.');
  } else {
    followUps.slice(0, 15).forEach(function(followUp, index) {
      const due = followUp.due_date_optional ? ' due ' + formatDateForDigest_(followUp.due_date_optional) : '';
      const dueNow = followUp.reminder_due ? ' [reminder due]' : '';
      lines.push('  ' + (index + 1) + '. ' + followUp.summary + due + dueNow);
    });
  }
  lines.push('');

  lines.push('__Reflections__');
  if (narrative.reflectionsSummary) {
    lines.push(narrative.reflectionsSummary);
  }
  if (reflections.length === 0) {
    lines.push('  1. No recent reflections from the last ' + lookbackDays + ' days.');
  } else {
    const moodCounts = countMoods_(reflections);
    const moodSummary = Object.keys(moodCounts)
      .sort(function(a, b) { return moodCounts[b] - moodCounts[a]; })
      .map(function(mood) { return mood + ' (' + moodCounts[mood] + ')'; })
      .join(', ');
    if (moodSummary) {
      lines.push('Mood tags: ' + moodSummary);
    }
    reflections.slice(0, 10).forEach(function(reflection, index) {
      const moods = reflection.moods ? ' [' + reflection.moods + ']' : '';
      lines.push('  ' + (index + 1) + '. ' + reflection.summary + moods);
    });
  }

  return lines.join('\n');
}

function renderDigestHtmlBody_(goals, followUps, reflections, narrative, now, lookbackDays) {
  return [
    '<div style="font-family:Arial,sans-serif;line-height:1.45;color:#202124;">',
    '<h2 style="margin:0 0 4px 0;">Voice Journal Digest</h2>',
    '<p style="margin:0 0 16px 0;color:#5f6368;">' + escapeHtml_(formatDateForDigest_(now)) + '</p>',
    narrative.intro ? '<p>' + escapeHtml_(narrative.intro) + '</p>' : '',
    renderDigestHtmlSection_('Goals', narrative.goalsSummary, goals.slice(0, 12), function(goal) {
      const activeUntil = goal.active_until ? ' through ' + formatDateForDigest_(goal.active_until) : '';
      const dueNow = goal.reminder_due ? ' [reminder due]' : '';
      return goal.summary + activeUntil + dueNow;
    }, 'No active goals.'),
    renderDigestHtmlSection_('Follow Ups', narrative.followUpsSummary, followUps.slice(0, 15), function(followUp) {
      const due = followUp.due_date_optional ? ' due ' + formatDateForDigest_(followUp.due_date_optional) : '';
      const dueNow = followUp.reminder_due ? ' [reminder due]' : '';
      return followUp.summary + due + dueNow;
    }, 'No open follow ups needing attention.'),
    renderDigestHtmlReflectionsSection_(reflections.slice(0, 10), narrative.reflectionsSummary, lookbackDays),
    '</div>',
  ].join('');
}

function renderDigestHtmlSection_(title, summary, items, itemText, emptyText) {
  const itemHtml = items.length === 0
    ? '<li>' + escapeHtml_(emptyText) + '</li>'
    : items.map(function(item) { return '<li>' + escapeHtml_(itemText(item)) + '</li>'; }).join('');
  return [
    '<p style="margin:20px 0 4px 0;"><strong><u>' + escapeHtml_(title) + '</u></strong></p>',
    summary ? '<p style="margin:0 0 8px 0;">' + escapeHtml_(summary) + '</p>' : '',
    '<ol style="margin:0 0 0 24px;padding-left:18px;">',
    itemHtml,
    '</ol>',
  ].join('');
}

function renderDigestHtmlReflectionsSection_(reflections, summary, lookbackDays) {
  const moodCounts = countMoods_(reflections);
  const moodSummary = Object.keys(moodCounts)
    .sort(function(a, b) { return moodCounts[b] - moodCounts[a]; })
    .map(function(mood) { return mood + ' (' + moodCounts[mood] + ')'; })
    .join(', ');
  const items = reflections.length === 0
    ? ['No recent reflections from the last ' + lookbackDays + ' days.']
    : reflections.map(function(reflection) {
      const moods = reflection.moods ? ' [' + reflection.moods + ']' : '';
      return reflection.summary + moods;
    });
  return [
    '<p style="margin:20px 0 4px 0;"><strong><u>Reflections</u></strong></p>',
    summary ? '<p style="margin:0 0 8px 0;">' + escapeHtml_(summary) + '</p>' : '',
    moodSummary ? '<p style="margin:0 0 8px 0;">Mood tags: ' + escapeHtml_(moodSummary) + '</p>' : '',
    '<ol style="margin:0 0 0 24px;padding-left:18px;">',
    items.map(function(item) { return '<li>' + escapeHtml_(item) + '</li>'; }).join(''),
    '</ol>',
  ].join('');
}

function applyDigestStateUpdates_(spreadsheet, digest) {
  applyObjectRowUpdates_(spreadsheet.getSheetByName(SHEET_NAMES.GOALS), digest.goalUpdates);
  applyObjectRowUpdates_(spreadsheet.getSheetByName(SHEET_NAMES.FOLLOW_UPS), digest.followUpUpdates);
}

function applyObjectRowUpdates_(sheet, updates) {
  normalizeArray_(updates).forEach(function(update) {
    if (update.rowNumber) {
      writeObjectRow_(sheet, update.rowNumber, update.values);
    }
  });
}

function buildDigestStateUpdate_(row, now) {
  const values = row.values;
  const frequencyDays = positiveInteger_(values.reminder_frequency_days, 1);
  return {
    rowNumber: row.rowNumber,
    values: copyObject_(values, {
      reminder_count: parseInteger_(values.reminder_count, 0) + 1,
      last_reminded_at: now,
      next_reminder_at: addDays_(now, frequencyDays),
    }),
  };
}

function isReminderDue_(values, now) {
  const nextReminderAt = parseDate_(values.next_reminder_at);
  return !!nextReminderAt && nextReminderAt.getTime() <= now.getTime();
}

function countMoods_(reflections) {
  const counts = {};
  reflections.forEach(function(reflection) {
    String(reflection.moods || '').split(',').forEach(function(rawMood) {
      const mood = rawMood.trim().toLowerCase();
      if (mood) {
        counts[mood] = (counts[mood] || 0) + 1;
      }
    });
  });
  return counts;
}

function buildDigestNarrative_(summary, narrativeGenerator) {
  try {
    if (narrativeGenerator) {
      const narrative = sanitizeDigestNarrative_(narrativeGenerator(summary));
      if (narrative.intro || narrative.goalsSummary || narrative.followUpsSummary || narrative.reflectionsSummary) {
        return narrative;
      }
    }
  } catch (error) {
    // Digest delivery should not depend on optional AI narrative.
  }
  return fallbackDigestNarrative_(summary);
}

function generateDigestNarrative_(config, summary) {
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
          content: [
            'Write concise narrative copy for a personal voice journal email digest.',
            'Be accurate, realistic, and grounded in the provided journal content.',
            'Do not be overly warm, upbeat, optimistic, or motivational.',
            'It is okay to be compassionate when the content warrants it.',
            'Do not add bullets, markdown, advice, or claims not supported by the content.',
            'The intro must be one or two sentences.',
            'The goals and follow ups summaries must be one sentence each, two only if needed.',
            'The reflections summary must be two sentences when possible and no more than three.',
          ].join(' '),
        },
        {
          role: 'user',
          content: digestNarrativePrompt_(summary),
        },
      ],
      text: {
        format: {
          type: 'json_schema',
          name: 'voice_journal_digest_narrative',
          strict: true,
          schema: digestNarrativeSchema_(),
        },
      },
    }),
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Digest intro failed with HTTP ' + code + ': ' + body);
  }

  const parsed = JSON.parse(body);
  const outputText = parsed.output_text || findOutputText_(parsed);
  if (!outputText) {
    throw new Error('Digest narrative response did not include output text.');
  }
  return JSON.parse(outputText);
}

function digestNarrativePrompt_(summary) {
  const goals = normalizeArray_(summary.goals).slice(0, 8).map(function(goal) { return goal.summary; }).filter(Boolean);
  const followUps = normalizeArray_(summary.followUps).slice(0, 10).map(function(followUp) { return followUp.summary; }).filter(Boolean);
  const reflections = normalizeArray_(summary.reflections).slice(0, 8).map(function(reflection) { return reflection.summary; }).filter(Boolean);
  const moods = moodCountsToRows_(summary.moodCounts || {}).slice(0, 5).map(function(row) {
    return row[0] + ' (' + row[1] + ')';
  });

  return [
    'Active goals: ' + (goals.join('; ') || 'none'),
    'Follow ups: ' + (followUps.join('; ') || 'none'),
    'Recent reflections: ' + (reflections.join('; ') || 'none'),
    'Mood counts: ' + (moods.join(', ') || 'none'),
  ].join('\n');
}

function digestNarrativeSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      intro: { type: 'string' },
      goalsSummary: { type: 'string' },
      followUpsSummary: { type: 'string' },
      reflectionsSummary: { type: 'string' },
    },
    required: ['intro', 'goalsSummary', 'followUpsSummary', 'reflectionsSummary'],
  };
}

function sanitizeDigestNarrative_(narrative) {
  if (typeof narrative === 'string') {
    return copyObject_(fallbackDigestNarrative_({}), { intro: sanitizeDigestText_(narrative, 500) });
  }
  return {
    intro: sanitizeDigestText_(narrative && narrative.intro, 500),
    goalsSummary: sanitizeDigestText_(narrative && narrative.goalsSummary, 350),
    followUpsSummary: sanitizeDigestText_(narrative && narrative.followUpsSummary, 350),
    reflectionsSummary: sanitizeDigestText_(narrative && narrative.reflectionsSummary, 600),
  };
}

function sanitizeDigestText_(value, maxLength) {
  const clean = String(value || '').replace(/\s+/g, ' ').trim();
  if (!clean) {
    return '';
  }
  return clean.length > maxLength ? clean.slice(0, maxLength - 3) + '...' : clean;
}

function fallbackDigestNarrative_(summary) {
  return {
    intro: fallbackDigestIntro_(summary),
    goalsSummary: fallbackSectionSummary_('goals', normalizeArray_(summary && summary.goals), summary),
    followUpsSummary: fallbackSectionSummary_('follow ups', normalizeArray_(summary && summary.followUps), summary),
    reflectionsSummary: fallbackSectionSummary_('reflections', normalizeArray_(summary && summary.reflections), summary),
  };
}

function fallbackDigestIntro_(summary) {
  const goals = normalizeArray_(summary && summary.goals);
  const followUps = normalizeArray_(summary && summary.followUps);
  const reflections = normalizeArray_(summary && summary.reflections);
  const moodRows = moodCountsToRows_((summary && summary.moodCounts) || {});
  const leadingMoods = moodRows.slice(0, 2).map(function(row) { return row[0]; });
  const moodSentence = leadingMoods.length > 0
    ? ' The most common moods were ' + leadingMoods.join(' and ') + '.'
    : '';
  return 'Today\'s digest includes ' + goals.length + ' active goals, '
    + followUps.length + ' follow ups, and ' + reflections.length + ' recent reflections.'
    + moodSentence;
}

function fallbackSectionSummary_(section, items, summary) {
  if (section === 'goals') {
    const dueGoals = items.filter(function(goal) { return goal.reminder_due; }).length;
    return items.length === 0
      ? 'There are no active goals in this digest.'
      : items.length + ' active goals are included' + (dueGoals ? ', with ' + dueGoals + ' currently reminder-due.' : '.');
  }
  if (section === 'follow ups') {
    const dueFollowUps = items.filter(function(followUp) { return followUp.reminder_due; }).length;
    return items.length === 0
      ? 'There are no open follow ups needing attention in this digest.'
      : items.length + ' follow ups are included' + (dueFollowUps ? ', with ' + dueFollowUps + ' currently reminder-due.' : '.');
  }
  const moodRows = moodCountsToRows_((summary && summary.moodCounts) || {});
  if (items.length === 0) {
    return 'There are no recent reflections in the current lookback window.';
  }
  if (moodRows.length === 0) {
    return items.length + ' recent reflections are included.';
  }
  return items.length + ' recent reflections are included. The most common mood tags are '
    + moodRows.slice(0, 2).map(function(row) { return row[0]; }).join(' and ') + '.';
}
