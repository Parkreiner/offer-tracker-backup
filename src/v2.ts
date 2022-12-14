/**
 * @file Makes a back-up copy of the NY/ECRI offer tracker.
 *
 * This tool is here in the off chance that a malicious actor tries to delete
 * the entire contents of the spreadsheet, or that someone accidentally breaks
 * things.
 *
 *
 * @todo
 *   - Test out the new functionality for tracking sheet changes, and get it
 *     hooked up
 *   - Get the project set up in CLASP and TypeScript; actually put it in a repo
 *   - Set the code up so that it tracks and lists all changes between the
 *     current sheet and the last back-up. List more than the very first thing
 *     that was detected as a change. (requires changing detectSheetChanges_ to
 *     track all the new info – maybe call it compileSheetChanges_, and also the
 *     email function to output the array items as a proper HTML list)
 */

import {
  BACKUP_FOLDER_ID,
  TRACKER_SPREADSHEET_ID,
  PERSONAL_EMAIL,
} from "./constants";

import {
  CellValue,
  DriveFolder,
  getValues,
  Sheet,
  Spreadsheet,
} from "./gasTypeHelpers";

/** Compiles basic information about what has changed since the last backup */
type BackupReport = {
  backupNeeded: boolean;
  backupAlreadyExists: boolean;
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

    const newName = `tracker_${getFormattedDateStamp_()}`;
    const [backupNeeded, message] = detectSheetChanges_(
      sourceSpreadsheet,
      backupsFolder,
      newName
    );

    if (!forceBackup && !backupNeeded) {
      console.log(`${message} Exiting.`);
      return;
    }

    console.log("Backup in progress.");
    console.log(message);
    console.log(`Force update value: ${forceBackup}`);

    const newSpreadsheet = copySpreadsheetContents_(
      sourceSpreadsheet,
      SpreadsheetApp.create(newName)
    );

    const newSpreadsheetRef = DriveApp.getFileById(newSpreadsheet.getId());
    newSpreadsheetRef.moveTo(backupsFolder);

    console.log("Backup complete");
    sendEmail_(
      PERSONAL_EMAIL,
      "[Offer letter tracker] Backup complete",
      message
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
 * Determines whether a backup is needed, by comparing the main offer tracker
 * spreadsheet against the most recent backup in the backups folder.
 * @private
 */
function INCOMPLETE_compileSheetChanges_(
  sourceSpreadsheet: Spreadsheet,
  backupsFolder: DriveFolder,
  newSpreadsheetName: string
): BackupReport {
  let backupNeeded = false;
  let backupAlreadyExists = false;
  const changes = [];

  const fileIteratorByName = backupsFolder.getFilesByName(newSpreadsheetName);
  if (fileIteratorByName.hasNext()) {
    backupAlreadyExists = true;
  }

  let sourceSheets: (Sheet | null)[];
  let newestSheets: (Sheet | null)[];
  {
    const tempSource = sourceSpreadsheet.getSheets().sort(orderBySheetName_);
    const tempNewest = SpreadsheetApp.openById(getIdNewestFile_(backupsFolder))
      .getSheets()
      .sort(orderBySheetName_);

    [sourceSheets, newestSheets] = pairUpSheets_(tempSource, tempNewest);
  }

  if (sourceSheets.length !== newestSheets.length) {
    backupNeeded = true;
    changes.push("Source spreadsheet sheet count changed.");
  }

  // Iterate over each pair of sheets. The entire loop is a little messy, but it
  // should technically be O(n) still, since each cell is only iterated over
  // once
  for (let i = 0; i < sourceSheets.length; i++) {
    // If one sheet is null, the other should be guaranteed to be defined
    const sourceSheet = sourceSheets[i];
    const newestSheet = newestSheets[i];

    if (!sourceSheet) {
      backupNeeded = true;
      const newName = newestSheet?.getName() ?? "N/A";
      changes.push(`Sheet ${newName} deleted from source spreadsheet`);
      continue;
    }

    if (!newestSheet) {
      backupNeeded = true;
      changes.push(
        `Sheet ${sourceSheet.getName()} deleted from source spreadsheet`
      );
      continue;
    }

    const sourceSheetName = sourceSheet.getName();
    const sourceValues = sourceSheet.getDataRange().getValues();
    const newestValues = newestSheet.getDataRange().getValues();

    // Detect change in row count
    if (sourceValues.length !== newestValues.length) {
      const diff = sourceValues.length - newestValues.length;
      const print = diff > 0 ? `+${diff}` : String(diff);

      backupNeeded = true;
      changes.push(
        `Row count for Source spreadsheet ${sourceSheetName}: ${print}.`
      );
    }

    // Iterate over each row
    for (let j = 0; j < sourceValues.length; j++) {
      const sourceRow = sourceValues[j];
      const newestRow = newestValues[j];

      // Detect change in column count
      if (sourceRow.length !== newestRow.length) {
        const diff = sourceRow.length - newestRow.length;
        const print = diff > 0 ? `+${diff}` : String(diff);

        backupNeeded = true;
        changes.push(
          `Column count for Source spreadsheet ${sourceSheetName}: ${print}.`
        );
      }

      // Iterate over each cell in the row
      for (let k = 0; k < sourceRow.length; k++) {
        // Detect value change
        if (sourceRow[k] !== newestRow[k]) {
          backupNeeded = true;
          changes.push(
            `Cell values changed for Source spreadsheet ${sourceSheetName}: row ${
              j + 1
            }.`
          );

          // Moving to next iteration of loop to avoid a bunch of changes being
          // registered on the same row (common for new offers being added)
          continue;
        }
      }
    }
  }

  return { backupNeeded, backupAlreadyExists, changes };
}

/**
 * Pairs up all sheets in the source spreadsheet with ones in the last backed-
 * up spreadsheet.
 *
 * If a sheet exists in one, but not the other, a value of null is inserted for
 * the other. All sheet arrays are assumed to be sorted alphabetically
 * beforehand.
 *
 * @todo Rewrite this logic. It seems messy and not that type-safe.
 */
function pairUpSheets_<S1 extends Sheet, S2 extends Sheet>(
  sourceSheets: S1[],
  lastBackupSheets: S2[]
): [sourceSheet: (S1 | null)[], backupSheets: (S2 | null)[]] {
  const organizedSources = [];
  const organizedBackups = [];

  let sourceIndex = 0;
  let backupIndex = 0;

  while (
    sourceIndex < sourceSheets.length &&
    backupIndex < lastBackupSheets.length
  ) {
    const sourceSheet = sourceSheets[sourceIndex];
    if (!sourceSheet) {
      throw new Error(`Unable to pull sourceSheet for index ${sourceIndex}`);
    }

    const backupSheet = lastBackupSheets[backupIndex];
    if (!backupSheet) {
      throw new Error(`Unable to pull backupSheet for index ${backupIndex}`);
    }

    const sourceName = sourceSheet.getName();
    const backupName = backupSheet.getName();

    // Source spreadsheet has a sheet the backups doesn't (added)
    if (sourceName < backupName) {
      organizedSources.push(sourceSheet);
      organizedBackups.push(null);
      sourceIndex++;
      continue;
    }

    // Backup spreadsheet has a sheet the backups doesn't (deleted)
    if (backupName < sourceName) {
      organizedSources.push(null);
      organizedBackups.push(backupSheet);
      backupIndex++;
      continue;
    }

    // Base case - names are the same
    organizedSources.push(sourceSheet);
    organizedBackups.push(backupSheet);
    sourceIndex++;
    backupIndex++;
  }

  return [organizedSources, organizedBackups];
}

/**
 * Determines whether a backup is needed, by comparing the main offer tracker
 * spreadsheet against the most recent backup in the backups folder.
 * @private
 */
function detectSheetChanges_(
  sourceSpreadsheet: Spreadsheet,
  backupsFolder: DriveFolder,
  newSsName: string
): [backupNeeded: boolean, message: string] {
  const fileIteratorByName = backupsFolder.getFilesByName(newSsName);
  if (fileIteratorByName.hasNext()) {
    return [false, "Backup already exists."];
  }

  const sourceSheets = sourceSpreadsheet.getSheets();
  const newestSheets = SpreadsheetApp.openById(
    getIdNewestFile_(backupsFolder)
  ).getSheets();
  if (sourceSheets.length !== newestSheets.length) {
    return [true, "Source spreadsheet sheet count changed."];
  }

  sourceSheets.sort(orderBySheetName_);
  newestSheets.sort(orderBySheetName_);

  // Three loops nested within each other. It's messy, but still technically
  // O(n), since each cell will only be iterated over once
  for (const [i, sourceSheet] of sourceSheets.entries()) {
    // Previous check ensures that sheet arrays have the same length
    const newSheet = newestSheets[i] as Sheet;

    const sourceSheetName = sourceSheet.getName();
    const sourceValues = getValues(sourceSheet.getDataRange());
    const newestValues = getValues(newSheet.getDataRange());

    if (sourceValues.length !== newestValues.length) {
      return [
        true,
        `Source spreadsheet row count changed for sheet ${sourceSheetName}.`,
      ];
    }

    for (const [j, sourceRow] of sourceValues.entries()) {
      const newRow = newestValues[j] as CellValue[];

      if (sourceRow.length !== newRow.length) {
        return [
          true,
          `Source spreadsheet column count changed for sheet ${sourceSheetName}.`,
        ];
      }

      for (const [k, cellValue] of sourceRow.entries()) {
        if (cellValue !== newRow[k]) {
          const indexOffset = j + 1;
          return [
            true,
            `Source spreadsheet cell values changed for sheet ${sourceSheetName} on row ${indexOffset}.`,
          ];
        }
      }
    }
  }

  return [false, "No changes detected."];
}

/**
 * Determines an orderinal value for two sheets, based on their names.
 * @private
 */
function orderBySheetName_(s1: Sheet, s2: Sheet): -1 | 0 | 1 {
  const sheet1Name = s1.getName();
  const sheet2Name = s2.getName();

  if (sheet1Name < sheet2Name) {
    return -1;
  } else if (sheet1Name > sheet2Name) {
    return 1;
  }

  return 0;
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

/**
 * Copies all the contents of source spreadsheet into target spreadsheet.
 *
 * Copying manually instead of using the built-in spreadsheet.copy to avoid
 * copying Google Apps Script script files over and over.
 * @private
 */
function copySpreadsheetContents_<SS extends Spreadsheet>(
  sourceSpreadsheet: Spreadsheet,
  targetSpreadsheet: SS
): SS {
  const oldSheets = targetSpreadsheet.getSheets();
  const copyPrefixMatcher = /^Copy.*?of */i;

  for (const sourceSheet of sourceSpreadsheet.getSheets()) {
    const newSheet = sourceSheet.copyTo(targetSpreadsheet);
    const cleanedName = newSheet.getName().replace(copyPrefixMatcher, "");
    newSheet.setName(cleanedName);
  }

  for (const sheet of oldSheets) {
    targetSpreadsheet.deleteSheet(sheet);
  }

  return targetSpreadsheet;
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
