/**
 * @fileoverview A comprehensive Google Apps Script to automate a driver referral program.
 * It handles form submissions, validates data, manages records, and sends conditional emails.
 */

// --- Configuration ---
// Define sheet names here for easy management and to avoid hard-coding strings.
const SHEET_NAMES = {
  referrals: 'Referrals',
  scouts: 'Scouts',
  churnedDrivers: 'Churned Drivers',
  validReferrals: 'Valid Referrals',
  notEligible: 'Not Eligible Referrals',
  driverActivity: 'Driver Activity',
  compensationDue: 'Compensation Due',
  blockedDrivers: 'Blocked Drivers'
};

const EMAIL_SENDER_INFO = {
  from: "no-reply@yourcompany.com",
  name: "Your Company Alias",
  replyTo: "no-reply@yourcompany.com"
};

/**
 * The main function triggered by a new Google Form submission.
 * It orchestrates the entire workflow: validation, duplicate checking,
 * data processing, and email sending.
 * @param {object} e The event object from the form submission trigger.
 */
function onFormSubmit(e) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const referralsSheet = spreadsheet.getSheetByName(SHEET_NAMES.referrals);

    if (!referralsSheet) {
      throw new Error(`Sheet not found: ${SHEET_NAMES.referrals}`);
    }

    validateNewRecords(referralsSheet);
    checkDuplicate(referralsSheet);

    const lastRow = referralsSheet.getLastRow();
    const [referrerMail, referredUserMail, referrerEligibility, referredUserEligibility, , , , , , duplicateCheck] = referralsSheet.getRange(lastRow, 6, 1, 14).getValues()[0];

    // Send the appropriate email based on validation and duplicate status.
    if (referrerEligibility === "Eligible" && referredUserEligibility === "Eligible" && duplicateCheck >= 2) {
      sendDuplicateEmail(referrerMail, referredUserMail);
    } else if (referrerEligibility === "Eligible" && referredUserEligibility === "Not Eligible") {
      sendIneligibleReferralEmails(referrerMail, referredUserMail);
    } else if (referrerEligibility === "Eligible" && referredUserEligibility === "Eligible" && duplicateCheck === 1) {
      sendConfirmationEmails(referrerMail, referredUserMail);
    }
    
    // Route the referral data to the correct sheet.
    routeReferralRecord(referralsSheet);
  } catch (error) {
    Logger.log(`An error occurred in onFormSubmit: ${error.message}`);
  }
}

/**
 * Validates a new referral against 'Scouts' and 'Churned Drivers' lists.
 * It populates eligibility status and other lookup details in the 'Referrals' sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} referralsSheet The active 'Referrals' sheet.
 */
function validateNewRecords(referralsSheet) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scoutSheet = spreadsheet.getSheetByName(SHEET_NAMES.scouts);
  const churnDriversSheet = spreadsheet.getSheetByName(SHEET_NAMES.churnedDrivers);

  if (!scoutSheet || !churnDriversSheet) {
    throw new Error("Validation sheets not found. Check sheet names.");
  }

  const lastRow = referralsSheet.getLastRow();
  const [referrerCode, churnDriverPhone, , , , churnDriverEmail] = referralsSheet.getRange(lastRow, 2, 1, 6).getValues()[0];

  const scoutData = scoutSheet.getDataRange().getValues();
  const churnData = churnDriversSheet.getDataRange().getValues();

  // Validate Scout's eligibility.
  const scoutMatch = scoutData.find(row => row[0] === referrerCode);
  if (scoutMatch) {
    referralsSheet.getRange(lastRow, 7).setValue("Eligible");
    referralsSheet.getRange(lastRow, 6).setValue(scoutMatch[4]); // Scout email
    referralsSheet.getRange(lastRow, 11).setValue(scoutMatch[1]); // Scout ID
    referralsSheet.getRange(lastRow, 5).setValue(scoutMatch[3]); // Scout name
  } else {
    referralsSheet.getRange(lastRow, 7).setValue("Not Eligible");
  }

  // Validate Referred Driver's eligibility.
  const churnedDriverMatch = churnData.find(row => row[2] === churnDriverPhone || row[3] === churnDriverEmail);
  if (churnedDriverMatch) {
    referralsSheet.getRange(lastRow, 8).setValue("Eligible");
    referralsSheet.getRange(lastRow, 12).setValue(churnedDriverMatch[0]); // Referred User ID
    referralsSheet.getRange(lastRow, 13).setValue(churnedDriverMatch[3]); // Referred User Email
  } else {
    referralsSheet.getRange(lastRow, 8).setValue("Not Eligible");
    referralsSheet.getRange(lastRow, 13).setValue("noreply-invalid@invalid-referrals.com");
  }
}

/**
 * Checks for previous eligible referrals for the same driver to detect duplicates.
 * A duplicate count is added to the last row of the 'Referrals' sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active 'Referrals' sheet.
 */
function checkDuplicate(sheet) {
  const lastRow = sheet.getLastRow();
  const [lastPhone, scoutEligible, driverEligible] = sheet.getRange(lastRow, 3, 1, 6).getValues()[0];
  let duplicateCount = 0;

  if (scoutEligible === "Eligible" && driverEligible === "Eligible") {
    const previousData = sheet.getRange(2, 3, lastRow - 1, 6).getValues();
    duplicateCount = previousData.filter(row => row[0] == lastPhone && row[4] === "Eligible" && row[5] === "Eligible").length;
    duplicateCount++; // Add the current record to the count.
  }

  sheet.getRange(lastRow, 14).setValue(duplicateCount);
}

/**
 * Routes a newly submitted referral record to the correct sheet ('Valid' or 'Not Eligible').
 * @param {GoogleAppsScript.Spreadsheet.Sheet} referralsSheet The active 'Referrals' sheet.
 */
function routeReferralRecord(referralsSheet) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lastRow = referralsSheet.getLastRow();
  const record = referralsSheet.getRange(lastRow, 1, 1, 14).getValues()[0];
  const referrerEligible = record[6]; // Column G
  const refereeEligible = record[7]; // Column H
  const duplicateCount = record[13]; // Column N

  if (referrerEligible === "Eligible" && refereeEligible === "Eligible" && duplicateCount === 1) {
    const validSheet = spreadsheet.getSheetByName(SHEET_NAMES.validReferrals);
    validSheet.appendRow(record);
  } else {
    const notEligibleSheet = spreadsheet.getSheetByName(SHEET_NAMES.notEligible);
    notEligibleSheet.appendRow(record);
  }
}

// --- Email Sending Functions ---
/**
 * Sends a general email with predefined sender options.
 * @param {string} toEmail The recipient's email address.
 * @param {string} subject The email subject line.
 * @param {string} body The email body content.
 */
function sendCustomEmail(toEmail, subject, body) {
  MailApp.sendEmail({
    to: toEmail,
    subject: subject,
    body: body,
    from: EMAIL_SENDER_INFO.from,
    name: EMAIL_SENDER_INFO.name,
    replyTo: EMAIL_SENDER_INFO.replyTo
  });
}

/**
 * Sends emails for a confirmed, unique referral.
 * @param {string} referrerEmail The referrer's email.
 * @param {string} referredUserEmail The referred user's email.
 */
function sendConfirmationEmails(referrerEmail, referredUserEmail) {
  const scoutSubject = "Reactivation Program - Confirmation 🎉";
  const scoutBody = "Hello, your referral has been successfully processed!";
  const driverSubject = "Reactivation Program - Confirmation 🎉";
  const driverBody = "Hello, you have been successfully referred to drive under our program!";

  sendCustomEmail(referrerEmail, scoutSubject, scoutBody);
  sendCustomEmail(referredUserEmail, driverSubject, driverBody);
}

/**
 * Sends emails when a referral is not eligible.
 * @param {string} referrerEmail The referrer's email.
 * @param {string} referredUserEmail The referred user's email.
 */
function sendIneligibleReferralEmails(referrerEmail, referredUserEmail) {
  const scoutSubject = "Reactivation Program - Non-Confirmation 😔";
  const scoutBody = "Hello, the referral you submitted is not eligible.";
  const driverSubject = "Reactivation Program - Not Eligible 😔";
  const driverBody = "Hello, your account was referred, but we couldn't find a match.";

  sendCustomEmail(referrerEmail, scoutSubject, scoutBody);
  sendCustomEmail(referredUserEmail, driverSubject, driverBody);
}

/**
 * Sends emails for a duplicate referral.
 * @param {string} referrerEmail The referrer's email.
 * @param {string} referredUserEmail The referred user's email.
 */
function sendDuplicateEmail(referrerEmail, referredUserEmail) {
  const subject = "Reactivation Program - Not Eligible 😔";
  const scoutBody = "Hello, the driver you referred has already been submitted by another agent.";
  const driverBody = "Hello, you have been referred by another agent. Kindly complete your trips to win big!";

  sendCustomEmail(referrerEmail, subject, scoutBody);
  sendCustomEmail(referredUserEmail, subject, driverBody);
}

/**
 * Sends conditional emails to referrers based on a referred user's progress.
 * This is intended to be run on a time-based trigger (e.g., daily or weekly).
 */
function sendConditionalEmails() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.validReferrals);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const [ , , , , , , , , scoutEmail, scoutName, , , hashedEmail, hashedPhone, scenario] = data[i];

    let subject, body;
    if (scenario === "No Trips") {
      subject = `Update on your Referral with phone number: ${hashedPhone}.`;
      body = `Hello ${scoutName}, The partner you referred has not completed a ride.`;
    } else if (scenario === "In progress") {
      subject = `Update on your Referral with phone number: ${hashedPhone}.`;
      body = `Hello ${scoutName}, The partner you referred is progressing well!`;
    } else if (scenario === "Missed") {
      subject = `Update on your Referral with phone number: ${hashedPhone}.`;
      body = `Hello ${scoutName}, The referral will not be compensated as they didn't meet the trip requirements.`;
    } else if (scenario === "Completed") {
      subject = `Update on your Referral with phone number: ${hashedPhone}.`;
      body = `Hello ${scoutName}, The partner you referred has completed the requirements! The bonus will be credited.`;
    } else {
      continue;
    }
    sendCustomEmail(scoutEmail, subject, body);
  }
}

/**
 * Updates reactivation and last activity dates for drivers in the 'Valid Referrals' sheet.
 * This function should be run on a trigger after new driver activity data is fetched.
 */
function updateReactivationAndLastActivityDates() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const referralsSheet = spreadsheet.getSheetByName(SHEET_NAMES.validReferrals);
  const activitySheet = spreadsheet.getSheetByName(SHEET_NAMES.driverActivity);

  if (!referralsSheet || !activitySheet) {
    throw new Error("Required sheets ('Valid Referrals' or 'Driver Activity') not found.");
  }

  const referralsData = referralsSheet.getDataRange().getValues();
  const activityData = activitySheet.getDataRange().getValues();

  const activityMap = new Map();
  for (let i = 1; i < activityData.length; i++) {
    const userId = activityData[i][0];
    const reactivationDate = activityData[i][1];
    if (userId) {
      activityMap.set(userId, reactivationDate);
    }
  }

  const updates = [];
  for (let j = 1; j < referralsData.length; j++) {
    const userId = referralsData[j][0];
    const reactivationDate = referralsData[j][4];
    if (activityMap.has(userId)) {
      if (!reactivationDate) {
        updates.push([j + 1, 5, activityMap.get(userId)]);
      } else {
        updates.push([j + 1, 6, activityMap.get(userId)]);
      }
    }
  }

  if (updates.length > 0) {
    updates.forEach(update => referralsSheet.getRange(update[0], update[1]).setValue(update[2]));
  }
}

/**
 * Checks for completed referrals and moves their records to the 'Compensation Due' sheet.
 * This function is typically run on a scheduled trigger.
 */
function compenationDueCheck() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(SHEET_NAMES.validReferrals);
  const targetSheet = spreadsheet.getSheetByName(SHEET_NAMES.compensationDue);

  if (!sourceSheet || !targetSheet) {
    throw new Error("Required sheets not found for compensation check.");
  }

  const lastRow = sourceSheet.getLastRow();
  for (let i = lastRow; i >= 2; i--) {
    const cellValue = sourceSheet.getRange(i, 15).getValue(); // Column O
    if (cellValue === "Completed") {
      const rowValues = sourceSheet.getRange(i, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
      targetSheet.appendRow(rowValues);
      sourceSheet.deleteRow(i);
    }
  }
}

/**
 * Fetches and updates churned driver data from an external source spreadsheet.
 * This function is designed to be run on a scheduled trigger.
 */
function fetchChurnData() {
  // Replace these with the actual spreadsheet IDs and sheet names you want to sync.
  const SOURCE_SPREADSHEET_ID = "";
  const SOURCE_SHEET_NAME = "";
  const TARGET_SHEET_NAME = SHEET_NAMES.churnedDrivers;

  if (!SOURCE_SPREADSHEET_ID || !SOURCE_SHEET_NAME) {
    Logger.log("Source spreadsheet ID or sheet name is not configured.");
    return;
  }

  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);

  if (!sourceSheet || !targetSheet) {
    throw new Error("Source or target sheet not found for fetchChurnData.");
  }

  const sourceRange = sourceSheet.getDataRange();
  const sourceData = sourceRange.getValues();

  targetSheet.clear();
  targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
}

/**
 * Fetches and updates driver activity data from an external source spreadsheet.
 * This function is designed to be run on a scheduled trigger.
 */
function fetchActivityData() {
  const SOURCE_SPREADSHEET_ID = "";
  const SOURCE_SHEET_NAME = "";
  const TARGET_SHEET_NAME = SHEET_NAMES.driverActivity;

  if (!SOURCE_SPREADSHEET_ID || !SOURCE_SHEET_NAME) {
    Logger.log("Source spreadsheet ID or sheet name is not configured.");
    return;
  }

  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);

  if (!sourceSheet || !targetSheet) {
    throw new Error("Source or target sheet not found for fetchActivityData.");
  }

  const sourceRange = sourceSheet.getDataRange();
  const sourceData = sourceRange.getValues();

  targetSheet.clear();
  targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
}

/**
 * A function to clear data in a specific range of the 'Blocked Drivers' sheet.
 */
function deleteEscalations() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.blockedDrivers);
  if (!sheet) {
    throw new Error(`Sheet not found: ${SHEET_NAMES.blockedDrivers}`);
  }
  // Clear the contents of a specific range (e.g., column D).
  sheet.getRange('D2:D').clearContent();
}

/**
 * This function processes rows in the 'Not Eligible Referrals' sheet that have been manually escalated
 * and meet specific criteria. It moves them to the 'Valid Referrals' sheet and marks them as resolved.
 */
function processEscalatedReferrals() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(SHEET_NAMES.notEligible);
  const targetSheet = spreadsheet.getSheetByName(SHEET_NAMES.validReferrals);
  const scoutSheet = spreadsheet.getSheetByName(SHEET_NAMES.scouts);

  if (!sourceSheet || !targetSheet || !scoutSheet) {
    throw new Error("Required sheets not found for processing escalated referrals.");
  }

  const data = sourceSheet.getDataRange().getValues();
  const scoutData = scoutSheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    // Check for eligibility conditions in the 'Not Eligible Referrals' sheet
    if (row[3] === "Eligible" && row[4] === "Escalated" && row[5] !== "" && row[8] === "" && row[9] === 0) {
      
      const scoutCode = row[2];
      const scoutMatch = scoutData.find(sRow => sRow[0] === scoutCode);

      if (scoutMatch) {
        // Prepare new record for 'Valid Referrals' sheet
        const newRecord = [
          row[0], // Phone
          row[1], // Email
          scoutMatch[4], // Scout Email
          scoutMatch[3], // Scout Name
          scoutMatch[1], // Scout ID
          row[2] // Scout Code
        ];
        
        targetSheet.appendRow(newRecord);
        sourceSheet.getRange(i + 1, 9).setValue("Resolved"); // Mark as resolved in Column I
      }
    }
  }
}

/**
 * Anonymizes a phone number for privacy.
 * @param {string} phone The phone number to hash.
 * @return {string} The hashed phone number.
 */
function hashPhone(phone) {
  const phoneStr = String(phone);
  if (phoneStr.length < 9) return phoneStr;
  return phoneStr.slice(0, 5) + "****" + phoneStr.slice(-4);
}

/**
 * Anonymizes an email address for privacy.
 * @param {string} email The email address to hash.
 * @return {string} The hashed email address.
 */
function hashEmail(email) {
  const atIndex = email.indexOf("@");
  if (atIndex < 3) return email;
  return email.slice(0, 2) + "****" + email.slice(atIndex);
}