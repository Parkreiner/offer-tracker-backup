/**
 * @file Defines domain-specific logic for backing up data from the Codesmith
 * offer tracker spreadsheet.
 */

import {
  DriveFolder,
  getValues_,
  Sheet,
  Spreadsheet,
  getIdNewestFile_,
  convertToColumnLetters_,
} from "./gasHelpers.js";

type DriveResource = { name: string; id: string };

/**
 * Compiles basic information about what has changed since the last backup.
 *
 * There is no relation between backupNeeded and backupAlreadyExists. A backup
 * can exist for the day, but already be out of date if later changes are made
 * in the same day.
 */
type BackupReport = {
  folder: DriveResource;
  sourceSpreadsheet: DriveResource;
  comparisonSpreadsheet: DriveResource;

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
  console.log("Backup info:");
  console.log(`Backups folder: ${report.folder.name} (${report.folder.id})`);
  console.log(
    `Source spreadsheet: ${report.sourceSpreadsheet.name} (${report.sourceSpreadsheet.id})`
  );
  console.log(
    `Comparison spreadsheet: ${report.comparisonSpreadsheet.name} (${report.comparisonSpreadsheet.id})`
  );

  const toWord = (b: boolean) => (b ? "Yes" : "No");

  console.log(`Backup already exists? ${toWord(report.backupAlreadyExists)}.`);
  console.log(`Changes since last backup? ${toWord(report.backupNeeded)}.`);
  console.log(`Backup forced? ${toWord(forceBackup)}.`);

  const changesBody =
    report.changes.length > 0
      ? `Changes detected:\n${report.changes.join("\n")}`
      : "Changes detected: None.";
  console.log(changesBody);
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
  const comparisonSpreadsheet = SpreadsheetApp.openById(
    getIdNewestFile_(backupsFolder)
  );

  const detectedChanges: string[] = [];
  const sheetPairs = pairUpSheets_(
    sourceSpreadsheet.getSheets(),
    comparisonSpreadsheet.getSheets()
  );

  for (const [sourceSheet, lastBackupSheet] of sheetPairs) {
    if (!sourceSheet) {
      detectedChanges.push(
        `Sheet ${lastBackupSheet.getName()} deleted from source spreadsheet`
      );
      continue;
    }

    const sourceName = sourceSheet.getName();
    if (!lastBackupSheet) {
      detectedChanges.push(`Sheet ${sourceName} added since last backup`);
      continue;
    }

    const sourceValues = getValues_(sourceSheet.getDataRange());
    const backupValues = getValues_(lastBackupSheet.getDataRange());

    const rowDiff = sourceValues.length - backupValues.length;
    if (rowDiff !== 0) {
      detectedChanges.push(
        rowDiff > 0
          ? `${rowDiff} row(s) added to sheet ${sourceName}`
          : `${rowDiff * -1} row(s) deleted from sheet ${sourceName}`
      );
    }

    for (const [i, sourceRow] of sourceValues.entries()) {
      const backupRow = backupValues[i];
      if (backupRow === undefined) break;

      const colDiff = sourceRow.length - backupRow.length;
      if (colDiff !== 0) {
        detectedChanges.push(
          colDiff > 0
            ? `${colDiff} column(s) added to sheet ${sourceName}`
            : `${colDiff * -1} column(s) deleted from sheet ${sourceName}`
        );
      }

      for (const [j, sourceValue] of sourceRow.entries()) {
        const backupValue = backupRow[j];
        if (backupValue === undefined) break;

        const valuesAreDifferent =
          sourceValue instanceof Date && backupValue instanceof Date
            ? sourceValue.getTime() !== backupValue.getTime()
            : sourceValue !== backupValue;

        if (valuesAreDifferent) {
          const row = i + 1;
          const col = convertToColumnLetters_(j + 1);

          detectedChanges.push(
            `Values changes for Row ${row}, Column ${col} in sheet ${sourceName}`
          );
        }
      }
    }
  }

  return {
    folder: {
      name: backupsFolder.getName(),
      id: backupsFolder.getId(),
    },

    sourceSpreadsheet: {
      name: sourceSpreadsheet.getName(),
      id: sourceSpreadsheet.getId(),
    },

    comparisonSpreadsheet: {
      name: comparisonSpreadsheet.getName(),
      id: comparisonSpreadsheet.getId(),
    },

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
