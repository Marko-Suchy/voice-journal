/**
 * Main
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Voice Journal')
    .addItem('Install / Repair Voice Journal', 'installVoiceJournal')
    .addSeparator()
    .addItem('Setup Sheet', 'setupVoiceJournalSheet')
    .addItem('Process Inbox Now', 'processVoiceMemoInbox')
    .addItem('Refresh Search Index', 'refreshSearchIndexFromMenu')
    .addItem('Rebuild Search Index', 'rebuildSearchIndexFromMenu')
    .addItem('Refresh Dashboard', 'refreshDashboardNow')
    .addItem('Install 15-Minute Poller', 'installVoiceMemoPoller')
    .addSeparator()
    .addItem('Run Selected Design Action', 'runSelectedDesignAction')
    .addSeparator()
    .addItem('Send Digest Now', 'sendDailyDigestNow')
    .addItem('Install Daily Digest Trigger', 'installDailyDigestTrigger')
    .addItem('Remove Daily Digest Trigger', 'removeDailyDigestTrigger')
    .addToUi();
}

function doGet() {
  setupVoiceJournalSheet();
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Voice Journal Search');
}
