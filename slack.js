/**
 * This script syncs data from Notion Database (sync) - Mandates (Google Sheet) to Slack by updating user profiles with custom fields.
 * Notion Database (sync) - Mandates (Google Sheet) : https://docs.google.com/spreadsheets/d/1uqCK0JDHKkuzDEfOlBcSqJCsVxHjCRouOW5F8hJZRUA/edit?gid=2131663677#gid=2131663677
 * It reads user-related data from a specified sheet, retrieves Slack user IDs based on email addresses,
 * and updates Slack user profiles with details such as position, team, Notion page URL, mandate status,
 * availability, profile creation date, and additional time tracking fields.
 *
 * The script includes:
 * - `syncSheetToSlack`: Main function that processes the data from the sheet and updates Slack profiles.
 * - `getUserIdByEmail`: Helper function that fetches the Slack user ID using an email address.
 * - `updateUserProfile`: Function that updates the Slack user profile with custom fields.
 * - `constructProfileFields`: Function that constructs the JSON payload for the custom Slack profile fields.
 *
 * New Custom Slack Profile Fields Added:
 * - `Time Tracker (Last Update)` (Linked to 'Last Update' column in Google Sheet) - Field ID: Xf07HUS9GSSC
 * - `Time Tracker (Total)` (Linked to 'Hours (decimal)' column in Google Sheet) - Field ID: Xf07GZDPHHV4
 *
 * The script uses a Slack token, securely stored in the script properties, to authenticate API requests.
 * It also handles errors gracefully, ensuring that any issues during the sync process are logged for review.
 * Notion link: https://www.notion.so/grey-box/Sync-Mandate-Google-Sheet-to-Slack-syncSheetsMandatesToSlackGreyBox-gs-ff5228987fbb409bbd0177f44deb9bf1?pvs=4
 */


const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
const SHEET_NAME = 'Mandates'; // Change to your actual sheet name

function syncSheetToSlack() {
    logToSlack(
        "📢 Starting execution of \`syncSheetsMandatesToSlackGreyBox\` script"
    );
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const header = data.shift(); // Remove the header row

    const emailIndex = header.indexOf('Email (Org)');
    const positionIndex = header.indexOf('Position');
    const teamIndex = header.indexOf('Team (Current)');
    const notionUrlIndex = header.indexOf('Notion Page URL');
    const mandateStatusIndex = header.indexOf('Mandate (Status)');
    const availabilityIndex = header.indexOf('Availability (avg h/w)');
    const createdProfileIndex = header.indexOf('Created (Profile)');

    // New fields indices
    const lastUpdateIndex = header.indexOf('Last Update');
    const hoursDecimalIndex = header.indexOf('Hours (decimal)');

    if (emailIndex === -1 || positionIndex === -1 || teamIndex === -1 || notionUrlIndex === -1 || mandateStatusIndex === -1 || availabilityIndex === -1 || createdProfileIndex === -1 || lastUpdateIndex === -1 || hoursDecimalIndex === -1) {
        throw new Error('Required columns not found.');
    }

    data.forEach(row => {
        const email = row[emailIndex];
        const position = row[positionIndex];
        const team = row[teamIndex];
        const notionUrl = row[notionUrlIndex];
        const mandateStatus = row[mandateStatusIndex];
        const availability = row[availabilityIndex];
        const createdProfile = row[createdProfileIndex];
        const lastUpdate = row[lastUpdateIndex]; // Time Tracker (Last Update)
        const hoursDecimal = row[hoursDecimalIndex]; // Time Tracker (Total)

        if (email) {
            const userId = getUserIdByEmail(email);
            if (userId) {
                const customFields = {
                    "Position": position,
                    "Team": team,
                    "NotionPageURL": notionUrl,
                    "MandateStatus": mandateStatus,
                    "Availability": availability,
                    "CreatedProfile": createdProfile.toString(), // Ensure it's treated as a plain string
                    "TimeTrackerLastUpdate": lastUpdate, // Linked to "Last Update" column
                    "TimeTrackerTotal": hoursDecimal // Linked to "Hours (decimal)" column
                };
                updateUserProfile(userId, customFields);
            }
        }
    });
    logToSlack(
        "📢 Execution of \`syncSheetsMandatesToSlackGreyBox\` script finished"
    );
}

function getUserIdByEmail(email) {
    const response = UrlFetchApp.fetch(`https://slack.com/api/users.lookupByEmail?email=${email}`, {
        headers: { Authorization: `Bearer ${SLACK_TOKEN}` },
        muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());

    return result.ok ? result.user.id : null;
}

function updateUserProfile(userId, fields) {
    const profileFields = constructProfileFields(fields);
    const payload = {
        user: userId,
        profile: JSON.stringify(profileFields)
    };

    const response = UrlFetchApp.fetch('https://slack.com/api/users.profile.set', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        headers: { Authorization: `Bearer ${SLACK_TOKEN}` },
        muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());

    if (!result.ok) {
        console.error(`Failed to update profile for user ${userId}: ${result.error}`);
        logToSlack(`Failed to update profile for user ${userId}: ${result.error}`);
    }
}

function constructProfileFields(fields) {
    return {
        "fields": {
            "Xf06JZK27DRA": { "value": fields.Position }, // Position field identifier
            "Xf03V366R202": { "value": fields.Team }, // Team field identifier
            "Xf06JGJMBZPZ": { "value": fields.NotionPageURL }, // Notion Page URL field identifier
            "Xf0759PXS7BP": { "value": fields.MandateStatus }, // Mandate (Status) field identifier
            "Xf074Y4V1KHV": { "value": fields.Availability }, // Availability (avg h/w) field identifier
            "Xf075CJ4SXEF": { "value": fields.CreatedProfile }, // Created (Profile) field identifier
            "Xf07HUS9GSSC": { "value": fields.TimeTrackerLastUpdate }, // Time Tracker (Last Update) linked to "Last Update" column
            "Xf07GZDPHHV4": { "value": fields.TimeTrackerTotal } // Time Tracker (Total) linked to "Hours (decimal)" column
        }
    };
}
