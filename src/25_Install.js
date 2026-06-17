/**
 * Install
 */

function installVoiceJournal() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  setupVoiceJournalSheet();

  const config = readConfig_(spreadsheet);
  const propertyName = getOpenAiApiKeyPropertyName_(config);
  const hasStoredApiKey = !!PropertiesService.getScriptProperties().getProperty(propertyName);
  const plan = buildVoiceJournalInstallPlan_(config, hasStoredApiKey);
  assertVoiceJournalInstallReady_(plan);

  DriveApp.getFolderById(plan.folderId);

  if (plan.shouldStoreApiKey) {
    PropertiesService.getScriptProperties().setProperty(plan.apiKeyPropertyName, plan.apiKeyToStore);
    clearConfigValue_(spreadsheet.getSheetByName(SHEET_NAMES.CONFIG), 'OPENAI_API_KEY_SETUP', 'Stored securely in Script Properties by installVoiceJournal(); this sheet cell was cleared.');
  }

  installVoiceMemoPoller();
  let digestInstalled = false;
  if (plan.shouldInstallDigest) {
    installDailyDigestTrigger();
    digestInstalled = true;
  }

  refreshDashboard_(spreadsheet, new Date());

  const message = buildVoiceJournalInstallSummary_({
    pollerInstalled: true,
    digestInstalled: digestInstalled,
    apiKeyStored: plan.shouldStoreApiKey,
  });
  spreadsheet.toast(message, 'Voice Journal', 8);
  return message;
}

function buildVoiceJournalInstallPlan_(config, hasStoredApiKey) {
  const folderId = String(config && config.DRIVE_INBOX_FOLDER_ID || '').trim();
  const apiKeyToStore = String(config && config.OPENAI_API_KEY_SETUP || '').trim();
  const digestRecipient = String(config && config.DIGEST_RECIPIENT_EMAIL || '').trim();
  const apiKeyPropertyName = getOpenAiApiKeyPropertyName_(config);
  const missingConfigKeys = [];

  if (!folderId) {
    missingConfigKeys.push('DRIVE_INBOX_FOLDER_ID');
  }
  if (!apiKeyToStore && !hasStoredApiKey) {
    missingConfigKeys.push('OPENAI_API_KEY_SETUP');
  }

  return {
    folderId: folderId,
    apiKeyPropertyName: apiKeyPropertyName,
    apiKeyToStore: apiKeyToStore,
    shouldStoreApiKey: !!apiKeyToStore,
    shouldInstallDigest: !!digestRecipient,
    digestRecipient: digestRecipient,
    missingConfigKeys: missingConfigKeys,
    isReady: missingConfigKeys.length === 0,
  };
}

function assertVoiceJournalInstallReady_(plan) {
  if (plan && plan.isReady) {
    return;
  }
  const keys = (plan && plan.missingConfigKeys || []).join(', ');
  throw new Error('Fill these Config values before running installVoiceJournal(): ' + keys + '. Then rerun installVoiceJournal().');
}

function buildVoiceJournalInstallSummary_(result) {
  const parts = ['Voice Journal installed. Inbox poller installed.'];
  if (result && result.apiKeyStored) {
    parts.push('OpenAI API key stored and cleared from Config.');
  } else {
    parts.push('OpenAI API key already stored.');
  }
  if (result && result.digestInstalled) {
    parts.push('Digest trigger installed.');
  } else {
    parts.push('Digest trigger skipped because DIGEST_RECIPIENT_EMAIL is blank.');
  }
  return parts.join(' ');
}

function getOpenAiApiKeyPropertyName_(config) {
  return String(config && config.OPENAI_API_KEY_PROPERTY || 'OPENAI_API_KEY').trim() || 'OPENAI_API_KEY';
}

function clearConfigValue_(configSheet, key, note) {
  const rowNumber = findConfigRowNumber_(configSheet, key);
  if (!rowNumber) {
    return false;
  }
  configSheet.getRange(rowNumber, 2).clearContent().setNote(note || '');
  return true;
}

function findConfigRowNumber_(configSheet, key) {
  const rows = readConfigRows_(configSheet);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].key === key) {
      return i + 2;
    }
  }
  return 0;
}
