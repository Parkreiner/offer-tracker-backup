/**
 * @file Defines the top-level entry point functions for performing daily back-
 * ups of the offer tracker spreadsheet.
 *
 * NOTE: This file cannot have any export statements at all, or the build
 * process will not work. GAS still doesn't support ES module syntax.
 */
import { compileBackupReport_, logBackupInfo_ } from "./backup.js";

import {
  BACKUP_FOLDER_ID,
  TRACKER_SPREADSHEET_ID,
  PERSONAL_EMAIL,
  DISPLAY_MESSAGES,
} from "./constants.js";

import {
  sendEmail_,
  getFormattedDateStamp_,
  copySpreadsheet_,
  getSpreadsheetById_,
  getFolderById_,
} from "./gasHelpers.js";

/**
 * For forcing a backup to happen, even if some part of the logic thinks
 * that it isn't needed.
 *
 * Just here in case some of my logic is broken, and I need to make backups
 * while getting things fixed. Should only ever be run directly from the Apps
 * Scripts console.
 */
// @ts-expect-error
function forceManualBackup(): void {
  backupDailyContents(true);
}

/**
 * Entrypoint for the script logic. Should be set up to run daily via an Apps
 * Script trigger.
 */
function backupDailyContents(forceBackup = false): void {
  try {
    const sourceSpreadsheet = getSpreadsheetById_(TRACKER_SPREADSHEET_ID);
    const backupsFolder = getFolderById_(BACKUP_FOLDER_ID);

    const baseSpreadsheetName = `tracker_${getFormattedDateStamp_()}`;
    const backupReport = compileBackupReport_(
      sourceSpreadsheet,
      backupsFolder,
      baseSpreadsheetName
    );

    logBackupInfo_(backupReport, forceBackup);

    if (!forceBackup && backupReport.backupAlreadyExists) {
      console.log(DISPLAY_MESSAGES.status.stop);
      return;
    }

    console.log(DISPLAY_MESSAGES.status.proceed);
    copySpreadsheet_(sourceSpreadsheet, baseSpreadsheetName, backupsFolder);
    console.log(DISPLAY_MESSAGES.status.success);

    const emailBody = `Changes detected:\n${backupReport.changes.join("\n")}`;
    sendEmail_(PERSONAL_EMAIL, DISPLAY_MESSAGES.email.success, emailBody);
  } catch (err: unknown) {
    const logValue = err instanceof Error ? err.stack : err;
    const emailBody =
      err instanceof Error ? err.stack : `Non-error value ${err} thrown`;

    console.error(logValue);
    sendEmail_(PERSONAL_EMAIL, DISPLAY_MESSAGES.email.error, emailBody);
  }
}
