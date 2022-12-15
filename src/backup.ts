/**
 * @file Makes a back-up copy of the NY/ECRI offer tracker.
 *
 * This tool is here in the off chance that a malicious actor tries to delete
 * the entire contents of the spreadsheet, or that someone accidentally breaks
 * things.
 */

import { DriveFolder, getValues, Sheet, Spreadsheet } from "./gasHelpers";

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
 * Logs all relevant info about a backup operation.
 */
export function logBackupInfo_(
  report: BackupReport,
  forceBackup: boolean
): void {
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

/**
 * Gets the ID of the most recent file in a folder.
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
 */
export function compileBackupReport_(
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

    const sourceName = sourceSheet.getName();
    const sourceValues = getValues(sourceSheet.getDataRange());
    const backupValues = getValues(lastBackupSheet.getDataRange());

    const rowDiff = sourceValues.length - backupValues.length;
    if (rowDiff !== 0) {
      detectedChanges.push(
        rowDiff > 0
          ? `${rowDiff} rows added to sheet ${sourceName}`
          : `${rowDiff * -1} rows deleted from sheet ${sourceName}`
      );
    }

    for (const [i, sourceRow] of sourceValues.entries()) {
      const backupRow = backupValues[i];
      if (backupRow === undefined) break;

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

        const valuesAreDifferent =
          sourceValue instanceof Date && backupValue instanceof Date
            ? sourceValue.getTime() === backupValue.getTime()
            : sourceValue === backupValue;

        if (valuesAreDifferent) {
          detectedChanges.push(
            `Values changes for row ${i + 1} in sheet ${sourceName}`
          );

          break;
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

    if (inSource && inBackup) {
      return [sourceMap.get(name) as Sheet, backupMap.get(name) as Sheet];
    }

    if (inSource) {
      return [sourceMap.get(name) as Sheet, null];
    }

    return [null, backupMap.get(name) as Sheet];
  });
}
