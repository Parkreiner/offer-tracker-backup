export type DriveFolder = GoogleAppsScript.Drive.Folder;
export type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
export type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

export type CellValue = string | number | boolean | Date;

/**
 * Takes a sheet range and gets the 2-D array of its cell values.
 *
 * The default getValues method returns an array of any values; this overrides
 * that type information with the true possible cell types.
 */
export function getValues(
  range: GoogleAppsScript.Spreadsheet.Range
): CellValue[][] {
  return range.getValues();
}
