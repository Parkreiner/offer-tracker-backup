/**
 * @file Very general helper functions and types for working with Apps Script.
 *
 * A number of these utilities wrap built-in methods to provide things like
 * better type information or better error details.
 *
 * @todo Update the email logic to have better formatting
 */

/** A Google Drive folder. */
export type DriveFolder = GoogleAppsScript.Drive.Folder;

/** A Google Drive spreadsheet. */
export type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;

/** A sheet within a Google Drive spreadsheet. */
export type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

/** A 2-D range of cells within the sheet of a spreadsheet. */
export type Range = GoogleAppsScript.Spreadsheet.Range;

/** All possible values for a cell in Google Sheets. */
export type CellValue = string | number | boolean | Date;

/**
 * Tries getting the folder specified by a given ID.
 * @throws {Error} If the folder cannot be retrieved.
 */
export function getFolderById_(folderId: string): DriveFolder {
  let folder: DriveFolder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (err: unknown) {
    if (err instanceof Error) err.message = `Folder ${folderId} unavailable`;
    throw err;
  }

  return folder;
}

/**
 * Tries getting the spreadsheet specified by a given ID.
 * @throws {Error} If the spreadsheet cannot be retrieved.
 */
export function getSpreadsheetById_(spreadsheetId: string): Spreadsheet {
  let spreadsheet: Spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (err: unknown) {
    if (err instanceof Error)
      err.message = `Spreadsheet ${spreadsheetId} unavailable`;
    throw err;
  }

  return spreadsheet;
}

/**
 * Takes a sheet range and gets the 2-D array of its cell values.
 *
 * The default getValues method returns a 2-D array of any values. This
 * overrides that type information to use the CellValue type.
 */
export function getValues_(r: Range): CellValue[][] {
  return r.getValues();
}

/**
 * Takes a base file name and a folder, and keeps appending numbers to the end
 * of the file name until there are no conflicts with files already in the
 * folder.
 */
export function getNextPossibleName_(
  folder: DriveFolder,
  fileName: string
): string {
  let index = 0;
  let iterator: GoogleAppsScript.Drive.FileIterator;
  let currentName: string;

  do {
    currentName = index === 0 ? fileName : `${fileName} (${index})`;
    iterator = folder.getFilesByName(currentName);
    index++;
  } while (iterator.hasNext());

  return currentName;
}

/**
 * Sends an email.
 *
 * Preemptively splitting this off into a separate function, in case the
 * functionality needs to be beefed up down the line.
 */
export function sendEmail_(
  emailAddress: string,
  subject: string,
  messageText = ""
): void {
  GmailApp.sendEmail(emailAddress, subject, messageText);
}

/**
 * Copies the contents of a source spreadsheet over to a specific folder, under
 * a specified name.
 */
export function copySpreadsheet_(
  sourceSpreadsheet: Spreadsheet,
  targetName: string,
  destinationFolder: DriveFolder
): void {
  const targetSpreadsheet = SpreadsheetApp.create(
    getNextPossibleName_(destinationFolder, targetName)
  );

  const oldSheets = targetSpreadsheet.getSheets();
  const copyOfMatcher = /^Copy of */i;

  for (const sourceSheet of sourceSpreadsheet.getSheets()) {
    const newSheet = sourceSheet.copyTo(targetSpreadsheet);

    // Google adds "Copy of" even if the sheet is going to a different
    // spreadsheet and there wouldn't be any name conflicts.
    newSheet.setName(newSheet.getName().replace(copyOfMatcher, ""));
  }

  for (const sheet of oldSheets) {
    targetSpreadsheet.deleteSheet(sheet);
  }

  const newSpreadsheetRef = DriveApp.getFileById(targetSpreadsheet.getId());
  newSpreadsheetRef.moveTo(destinationFolder);
}

/**
 * Creates a datestamp for the current day in the format YYYY-MM-DD.
 */
export function getFormattedDateStamp_(): string {
  const date = new Date();
  const month = String(1 + date.getMonth()).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${date.getFullYear()}-${month}-${day}`;
}

/**
 * Gets the ID of the most recent file in a folder.
 * @throws {Error} If folder is completely empty.
 */
export function getIdNewestFile_(folder: DriveFolder): string {
  const fileIterator = folder.getFiles();
  if (!fileIterator.hasNext()) {
    throw new Error(`Folder ${folder.getId()} is empty.`);
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
 * Converts a column number into the letter(s) used in a sheet.
 */
export function convertToColumnLetters_(column: number): string {
  if (!Number.isInteger(column) || column < 1 || column >= 18278) {
    throw new RangeError(`Column ${column} is not a valid integer.`);
  }

  const ASCII_A = 65;
  const letterBuffer: string[] = [];

  let remainder = column;
  while (remainder > 0) {
    const columnGroupValue = (remainder - 1) % 26;
    remainder = (remainder - columnGroupValue - 1) / 26;

    const newLetter = String.fromCharCode(columnGroupValue + ASCII_A);
    letterBuffer.unshift(newLetter);
  }

  return letterBuffer.join("");
}
