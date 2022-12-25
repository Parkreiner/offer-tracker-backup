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
  CellValue,
} from "./gasHelpers.js";

/** Indicates info about a folder/document in a BackupReport */
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

export function formatBackupReport_(
  report: BackupReport,
  backupForced: boolean
): string {
  const toWord = (b: boolean) => (b === true ? "Yes" : "No");
  const changeList = report.changes.map((l) => `- ${l}`).join("\n") || "None.";

  return [
    "Backup info:",
    `Backups folder: "${report.folder.name}" (ID ${report.folder.id})`,
    `Source spreadsheet: "${report.sourceSpreadsheet.name}" (ID ${report.sourceSpreadsheet.id})`,
    `Comparison spreadsheet: "${report.comparisonSpreadsheet.name}" (ID ${report.comparisonSpreadsheet.id})`,
    "",
    `Backup already exists? ${toWord(report.backupAlreadyExists)}.`,
    `Changes since last backup? ${toWord(report.backupNeeded)}.`,
    `Backup forced? ${toWord(backupForced)}.`,
    "",
    "Changes detected:",
    changeList,
  ].join("\n");
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

        if (areValuesDifferent_(sourceValue, backupValue)) {
          const row = i + 1;
          const col = convertToColumnLetters_(j + 1);

          detectedChanges.push(
            `Values changed for cell ${col}${row} in sheet ${sourceName}`
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
 * Determines if two Google Sheet cell values are different.
 */
function areValuesDifferent_(v1: CellValue, v2: CellValue): boolean {
  if (v1 instanceof Date && v2 instanceof Date) {
    return v1.getTime() !== v2.getTime();
  }

  // Have to do duck-typing to determine whether a value is a CellImageBuilder;
  // GAS doesn't give you direct access to the classes, so no instanceof
  const v1IsImage = typeof v1 === "object" && "getUrl" in v1;
  const v2IsImage = typeof v2 === "object" && "getUrl" in v2;
  if (v1IsImage && v2IsImage) {
    return v1.getUrl() !== v2.getUrl();
  }

  return v1 !== v2;
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
