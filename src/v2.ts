/**
 * @file Makes a back-up copy of the NY/ECRI offer tracker.
 *
 * This tool is here in the off chance that a malicious actor tries to delete
 * the entire contents of the spreadsheet, or that someone accidentally breaks
 * things.
 *
 * @todo
 *   - Test the functionality for the newer functions.
 *   - Split up the functionality across more files.
 *   - Update the email logic to have better formatting
 */

import { DriveFolder, getValues, Sheet, Spreadsheet } from "./gasTypeHelpers";
import {
  BACKUP_FOLDER_ID,
  TRACKER_SPREADSHEET_ID,
  PERSONAL_EMAIL,
} from "./constants";

/**
 * Compiles basic information about what has changed since the last backup.
 *
 * There is no relation between backupNeeded and backupAlreadyExists. A backup
 * can exist for the day, but already be out of date if later changes are made
 * in the same day.
 */
type BackupReport = {
  /**
   * Indicates whether there are differences between the source sheet and the
   * most recent backup.
   */
  backupNeeded: boolean;

  /** Indicates whether a backup has already been created for the day. */
  backupAlreadyExists: boolean;

  /** Compiles all changes between the source sheet and the last backup */
  changes: string[];
};

/**
 * For forcing a backup to happen, even if some part of the logic thinks
 * that it isn't needed.
 *
 * Just here in case some of my logic is broken, and I need to make backups
 * while getting things fixed. Should only ever be run directly from the Apps
 * Scripts console.
 */
function forceManualBackup(): void {
  backupDailyContents(true);
}

/**
 * Entrypoint for the script logic. Should be set up to run daily via
 * an Apps Script trigger.
 */
function backupDailyContents(forceBackup = false): void {
  try {
    const [backupsFolder, sourceSpreadsheet] = getBackupResources_(
      BACKUP_FOLDER_ID,
      TRACKER_SPREADSHEET_ID
    );

    const baseName = `tracker_${getFormattedDateStamp_()}`;
    const backupReport = compileBackupReport_(
      sourceSpreadsheet,
      backupsFolder,
      baseName
    );

    logBackupInfo_(backupReport, forceBackup);

    const shouldProceed = forceBackup || !backupReport.backupAlreadyExists;
    if (!shouldProceed) {
      console.log("Not proceeding with back up. Exiting script.");
      return;
    }

    copySpreadsheet_(
      sourceSpreadsheet,
      backupsFolder,
      SpreadsheetApp.create(getNextPossibleName_(backupsFolder, baseName))
    );

    console.log("Backup complete.");
    sendEmail_(
      PERSONAL_EMAIL,
      "[Offer letter tracker] Backup complete",
      backupReport.changes.join("\n")
    );
  } catch (err: unknown) {
    const logValue = err instanceof Error ? err.stack : err;
    const emailBody =
      err instanceof Error ? err.stack : `Non-error value ${err} thrown`;

    console.error(logValue);
    sendEmail_(
      PERSONAL_EMAIL,
      "[Offer letter tracker] Error when backing up offer tracker",
      emailBody
    );
  }
}

function logBackupInfo_(report: BackupReport, forceBackup: boolean): void {
  const toWord = (b: boolean) => (b ? "Yes" : "No");
  const changeContent =
    report.changes.length > 0
      ? `Changes detected:\n${report.changes.join("\n")}`
      : "Changes detected: None.";

  console.log("Backup info:");
  console.log(`Backup needed? ${toWord(report.backupNeeded)}.`);
  console.log(`Backup already exists? ${toWord(report.backupAlreadyExists)}.`);
  console.log(`Backup forced? ${toWord(forceBackup)}.`);
  console.log(changeContent);
}

function getNextPossibleName_(folder: DriveFolder, baseName: string): string {
  let index = 0;
  let fileIterator = folder.getFilesByName(baseName);
  let currentName: string;

  do {
    currentName = index === 0 ? baseName : `${baseName} (${index})`;
    fileIterator = folder.getFilesByName(currentName);
    index++;
  } while (fileIterator.hasNext());

  return currentName;
}

/**
 * Tries retrieving the folder and spreadsheet needed to perform backups.
 *
 * Ideally, this logic wouldn't need to be in a separate function, but Google
 * Apps Script's error messages when you fail to retrieve a resource are
 * terrible and don't tell you anything. This is here to intercept the errors
 * and modify their messages before re-throwing them.
 * @private
 *
 * @throws {Error} If either resource cannot be retrieved.
 */
function getBackupResources_(
  folderId: string,
  spreadsheetId: string
): [backupFolder: DriveFolder, sourceSpreadsheet: Spreadsheet] {
  let backupsFolder: DriveFolder;
  try {
    backupsFolder = DriveApp.getFolderById(folderId);
  } catch (err: unknown) {
    if (err instanceof Error) err.message = "Backups folder unavailable";
    throw err;
  }

  let sourceSpreadsheet: Spreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (err: unknown) {
    if (err instanceof Error)
      err.message = "Offer tracker spreadsheet unavailable";
    throw err;
  }

  return [backupsFolder, sourceSpreadsheet];
}

/**
 * Gets the ID of the most recent file in a folder.
 * @private
 */
function getIdNewestFile_(folder: DriveFolder): string {
  const fileIterator = folder.getFiles();
  if (!fileIterator.hasNext()) {
    throw new Error("Folder is empty.");
  }

  let newestFileRef = fileIterator.next();
  while (fileIterator.hasNext()) {
    const nextFileRef = fileIterator.next();
    if (nextFileRef.getDateCreated() > newestFileRef.getDateCreated()) {
      newestFileRef = nextFileRef;
    }
  }

  return newestFileRef.getId();
}

/**
 * Goes through the source spreadsheet and the last backed-up spreadsheet, and
 * returns an object reporting all their changes.
 * @private
 */
function compileBackupReport_(
  sourceSpreadsheet: Spreadsheet,
  backupsFolder: DriveFolder,
  spreadsheetNameToFind: string
): BackupReport {
  const detectedChanges: string[] = [];
  const sheetPairs = pairUpSheets_(
    sourceSpreadsheet.getSheets(),
    SpreadsheetApp.openById(getIdNewestFile_(backupsFolder)).getSheets()
  );

  for (const [sourceSheet, lastBackupSheet] of sheetPairs) {
    // Rule out that a sheet is missing from a given pair
    if (!sourceSheet) {
      detectedChanges.push(
        `Sheet ${lastBackupSheet.getName()} deleted from source spreadsheet`
      );
      continue;
    }

    if (!lastBackupSheet) {
      detectedChanges.push(
        `Sheet ${sourceSheet.getName()} added since last backup`
      );
      continue;
    }

    // Gather values now that both sheets are definitely defined
    const sourceName = sourceSheet.getName();
    const sourceValues = getValues(sourceSheet.getDataRange());
    const backupValues = getValues(lastBackupSheet.getDataRange());

    // Handle changes in row count
    const rowDiff = sourceValues.length - backupValues.length;
    if (rowDiff !== 0) {
      detectedChanges.push(
        rowDiff > 0
          ? `${rowDiff} rows added to sheet ${sourceName}`
          : `${rowDiff * -1} rows deleted from sheet ${sourceName}`
      );
    }

    // Start iterating through individual cell values
    for (const [i, sourceRow] of sourceValues.entries()) {
      const backupRow = backupValues[i];
      if (!backupRow) break;

      // Handle differences in column count
      const colDiff = sourceRow.length - backupRow.length;
      if (colDiff !== 0) {
        detectedChanges.push(
          colDiff > 0
            ? `${colDiff} columns added to sheet ${sourceName}`
            : `${colDiff * -1} columns deleted from sheet ${sourceName}`
        );
      }

      for (const [j, sourceValue] of sourceRow.entries()) {
        const backupValue = backupRow[j];
        if (backupValue === undefined) break;

        const valuesDifferent =
          sourceValue instanceof Date && backupValue instanceof Date
            ? sourceValue.getTime() === backupValue.getTime()
            : sourceValue === backupValue;

        if (valuesDifferent) {
          detectedChanges.push(
            `Values changes for row ${j} in sheet ${sourceName}`
          );
        }
      }
    }
  }

  return {
    backupNeeded: detectedChanges.length > 0,
    changes: detectedChanges,
    backupAlreadyExists: backupsFolder
      .getFilesByName(spreadsheetNameToFind)
      .hasNext(),
  };
}

/**
 * Pairs up all sheets in the source spreadsheet with ones in the last backed-
 * up spreadsheet, by turning each pair into a two-element tuple.
 *
 * At least one element in each tuple is guaranteed to be defined.
 * @private
 */
function pairUpSheets_(
  sourceSheets: Sheet[],
  lastBackupSheets: Sheet[]
): ([Sheet, Sheet] | [Sheet, null] | [null, Sheet])[] {
  // The whole function isn't the most efficient, but should be easy to maintain
  const toMapEntry = (s: Sheet) => [s.getName(), s] as const;
  const sourceMap = new Map(sourceSheets.map(toMapEntry));
  const backupMap = new Map(lastBackupSheets.map(toMapEntry));

  const uniqueSheetNames = [
    ...new Set([
      ...sourceSheets.map((s) => s.getName()),
      ...lastBackupSheets.map((s) => s.getName()),
    ]),
  ].sort();

  return uniqueSheetNames.map((name) => {
    const inSource = sourceMap.has(name);
    const inBackup = backupMap.has(name);

    if (inSource) {
      if (inBackup) {
        return [sourceMap.get(name) as Sheet, backupMap.get(name) as Sheet];
      }

      return [sourceMap.get(name) as Sheet, null];
    }

    return [null, backupMap.get(name) as Sheet];
  });
}

/**
 * Encapsulates the steps needed to make a formatted date stamp. Timestamp
 * is formatted to work with > comparisons right out of the box.
 * @private
 */
function getFormattedDateStamp_(): string {
  const date = new Date();
  const month = String(1 + date.getMonth()).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${date.getFullYear()}-${month}-${day}`;
}

function copySpreadsheet_(
  sourceSpreadsheet: Spreadsheet,
  destinationFolder: DriveFolder,
  targetSpreadsheet: Spreadsheet
): void {
  const oldSheets = targetSpreadsheet.getSheets();
  const copyPrefixMatcher = /^Copy.*?of */i;

  for (const sourceSheet of sourceSpreadsheet.getSheets()) {
    const newSheet = sourceSheet.copyTo(targetSpreadsheet);
    newSheet.setName(newSheet.getName().replace(copyPrefixMatcher, ""));
  }

  for (const sheet of oldSheets) {
    targetSpreadsheet.deleteSheet(sheet);
  }

  const newSpreadsheetRef = DriveApp.getFileById(targetSpreadsheet.getId());
  newSpreadsheetRef.moveTo(destinationFolder);
}

/**
 * Sends an email.
 *
 * Preemptively splitting this off into a separate function, in case the
 * functionality needs to be beefed up down the line.
 * @private
 */
function sendEmail_(
  emailAddress: string,
  subject: string,
  messageText = ""
): void {
  GmailApp.sendEmail(emailAddress, subject, messageText);
}
