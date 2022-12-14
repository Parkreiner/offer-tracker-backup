/**
 * @file Defines a number of helper types and functions for working with Google
 * Apps Script on the type level.
 */

/** A Google Drive folder. */
export type DriveFolder = GoogleAppsScript.Drive.Folder;

/** A Google Drive spreadsheet. */
export type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;

/** A sheet within a Google Drive spreadsheet. */
export type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

/** All possible values for a cell in Google Sheets. */
export type CellValue = string | number | boolean | Date;

/**
 * Takes a sheet range and gets the 2-D array of its cell values.
 *
 * The default getValues method returns a 2-D array of any values. This
 * overrides that type information with the CellValue type.
 */
export function getValues(
  range: GoogleAppsScript.Spreadsheet.Range
): CellValue[][] {
  return range.getValues();
}
