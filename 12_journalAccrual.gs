function runHourlyAccrualExport() {
  const folderIn  = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_FOLDER_ID);
  const folderOut = DriveApp.getFolderById(CONFIG.HOURLY_JOURNAL_OUTPUT_FOLDER_ID);

  const file = getCsvFileByKeyword(folderIn, 'accrual');
  const blob = file.getBlob();

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  const filename = `hourlyAccrual_journal_${timestamp}.csv`;

  const copied = folderOut.createFile(blob.setName(filename));
  return copied;
}
