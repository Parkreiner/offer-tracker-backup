/**
 * @file Defines the global constants used throughout the script.
 */

/** Provides basic validation for a Gmail email address */
type GmailAddress = `${string}@gmail.com`;

/** Reference to offer tracker spreadsheet */
export const TRACKER_SPREADSHEET_ID =
  "1L3gsmEUAslu91eJ6B55DZaSdXNFzw6kuVL_L2jjUGDI";

/** Reference to personal folder where files will be backed up */
export const BACKUP_FOLDER_ID = "1gny3Ry9unE1MuD587YAzcg8T53IrUZRY";

/** Email address to send error messages to if things break */
export const PERSONAL_EMAIL =
  "rhinocerocketman@gmail.com" satisfies GmailAddress;

export const DISPLAY_MESSAGES = {
  email: {
    success: "[Offer letter tracker] Backup complete",
    error: "[Offer letter tracker] Error when backing up offer tracker",
  },

  status: {
    proceed: "Proceeding with back up.",
    stop: "Not proceeding with back up. Exiting script.",
    success: "Backup complete.",
  },
} as const satisfies Record<string, Record<string, string>>;
