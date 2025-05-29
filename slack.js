/**
 * This script synchronizes teams from a Google Sheet to Slack user groups, managing both group creation and user membership.
 *
 * It reads user data from a specified Google Sheet, identifies team assignments, and ensures Slack user groups exist for each team.
 * The script then manages user membership in these groups based on sheet data, adding users to appropriate groups
 * and maintaining group membership. It also includes functionality to deactivate empty groups.
 *
 * The script includes:
 * - `myFunction`: Main function that orchestrates the sync process and demonstrates group/user management.
 * - `readUsers`: Reads user data from the sheet and organizes it into a Map with team names as keys and arrays of emails as values.
 * - `createUserGroup`: Creates a new Slack user group if it doesn't already exist.
 * - `checkForExistingGroups`: Verifies whether a specified Slack user group already exists.
 * - `getUserIdByEmail`: Retrieves a Slack user ID based on their email address.
 * - `addUsersToUsergroup`: Adds one or more users to a Slack user group, avoiding duplicate additions.
 * - `getUserGroupMembers`: Retrieves a list of all user IDs that are members of a specified Slack user group.
 * - `createMessageQueue`: Implements a message queuing system with rate limiting for sending messages to Slack.
 * - `logToSlack`: Logs messages to both Google Apps Script Logger and a specified Slack channel.
 * - `fetchWithRetry`: Handles HTTP requests with automatic retry logic and exponential backoff for rate limits.
 *
 * Key Features:
 * - Filters users based on status fields (ignoring those marked as 'To Verify', 'Completed', or 'Archived')
 * - Handles comma-separated team assignments, placing users in multiple groups when needed
 * - Checks for existing groups before creating new ones to prevent duplicates
 * - Verifies user membership before attempting to add users to groups
 * - Includes rate limiting handling for Slack API requests
 * - Provides comprehensive logging for debugging and tracking
 * - Supports batch processing of users across multiple teams
 * - Implements asynchronous message queuing to respect Slack API rate limits
 * - Uses exponential backoff for failed API requests
 * - Supports dual logging to both Script Logger and a dedicated Slack channel
 *
 * Authentication:
 * - Uses a user token with permissions for group management
 * - Uses a bot token for logging to Slack channels
 * - Securely stores tokens in Script Properties
 *
 * Data Structure:
 * The script expects a sheet with columns:
 * - 'Email (Org)': User email addresses
 * - 'Mandate (Status)': Status field for filtering
 * - 'Team (Current)': Comma-separated list of team names
 */



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
/**
 * Creates a Slack user group with the specified name and a generated handle.
 * Checks if the group exists before attempting to create it.
 * Logs events, including handle generation defaults via generateSlackHandle.
 *
 * @param {string} slackGroupName - The desired name for the user group.
 * @param {string} slackToken - Slack API token with group creation permissions.
 * @return {Object | undefined} Response object from Slack API or undefined on error.
 * @requires generateSlackHandle - Function to generate Slack-compliant handles.
 * @requires checkForExistingGroups - Function to check for existing groups.
 * @requires fetchWithRetry - Function for making HTTP requests with retries.
 * @requires logToSlack - Global function for logging messages.
 * @requires Logger - Assumed Google Apps Script Logger or similar.
 */
function createUserGroup(slackGroupName, slackToken = SLACK_USER_TOKEN) {
    // Input validation for the name itself
    if (!slackGroupName || typeof slackGroupName !== 'string' || slackGroupName.trim() === '') {
        const errorMessage = `Invalid group name provided: "${slackGroupName}". Cannot create group.`;
        Logger.log(errorMessage);
        // Assuming logToSlack is defined and accessible
        logToSlack(errorMessage);
        return { ok: false, error: 'invalid_group_name_provided' };
    }

    const trimmedName = slackGroupName.trim(); // Use trimmed name for consistency

    try {
        Logger.log(`Attempting to find or create user group: "${trimmedName}"`);
        // Check using trimmed name
        const groupId = checkForExistingGroups(trimmedName, slackToken);

        if (groupId) {
            Logger.log(
                `Group "${trimmedName}" already exists with ID: ${groupId}`
            );
            // Return structure consistent with Slack API success response
            return {
                ok: true,
                usergroup: {
                    id: groupId,
                    name: trimmedName, // Return the name we checked for
                },
            };
        } else {
            Logger.log(
                `Group "${trimmedName}" not found. Proceeding with creation.`
            );

            // *** Generate the handle using the dedicated function ***
            // Pass the original (but trimmed) name to the handle generator
            const groupHandle = generateSlackHandle(trimmedName);
            // Note: generateSlackHandle internally calls logToSlack if it creates a default handle

            Logger.log(
                `Generated handle for group "${trimmedName}": "${groupHandle}". Creating new group via API.`
            );

            const url = 'https://slack.com/api/usergroups.create';
            const options = {
                method: 'post',
                headers: {
                    Authorization: 'Bearer ' + slackToken,
                    'Content-Type': 'application/json',
                },
                // *** Use the generated handle and the trimmed name in the payload ***
                payload: JSON.stringify({ name: trimmedName, handle: groupHandle }),
                muteHttpExceptions: true, // Handle errors based on Slack's JSON response
            };

            // Assuming fetchWithRetry is defined and accessible
            const response = fetchWithRetry(url, options);
            const responseText = response.getContentText();
            const data = JSON.parse(responseText);

            // Log the outcome based on Slack's response
            if (data.ok) {
                Logger.log(
                    `Successfully created group "${
                        data.usergroup?.name || trimmedName // Use name from response if available
                    }" with handle "${
                        data.usergroup?.handle || groupHandle
                    }". ID: ${data.usergroup?.id}.`
                );
                // Optional: Log full success response if needed for debugging
                // Logger.log("Full Slack API response (Success): " + responseText);
            } else {
                // Log the failure and the specific error from Slack
                const errorMessage = `Failed to create Slack group "${trimmedName}" (handle: "${groupHandle}"). Slack API Error: ${
                    data.error || 'unknown_error'
                }. Response: ${responseText}`;
                Logger.log(errorMessage);
                // Assuming logToSlack is defined and accessible
                logToSlack(errorMessage); // Send Slack API errors to Slack for visibility
            }
            return data; // Return the full response object from Slack API
        }
    } catch (error) {
        // Catch unexpected errors (network issues, parsing errors, etc.)
        const errorMessage = `Unexpected error in createUserGroup for "${trimmedName}": ${error.message}`;
        Logger.log(`${errorMessage}\n${error.stack || ''}`);
        // Assuming logToSlack is defined and accessible
        logToSlack(errorMessage);
        // Return a generic error structure
        return { ok: false, error: `internal_script_error: ${error.message}` };
    }
}

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
/**
 * Gets the slack userID for a user by their email
 * @param {string} email - Email of the user
 * @param {string} token - Slack API token with group management permissions
 * @return {Object} UserID
 */
 function getUserIdByEmail(email, token) {
    try {
        const url = `https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(email)}`;
        const options = {
            method: 'get',
            headers: {Authorization: 'Bearer ' + token},
            muteHttpExceptions: true,
        };

        const response = fetchWithRetry(url, options);
        const result = JSON.parse(response.getContentText());

        if (result.ok) {
            return result.user.id;
        } else {
            Logger.log(`Error finding user by email (${email}): ${result.error}`);
            return null;
        }

    } catch (error) {
        logToSlack(`An error occurred while retrieving user ID by email (${email}): ${error}`);
        return null;
    }
}

/**
 * Adds the users to a specific group if they're not in it already
 * @param {string} userGroupId - The id of the group to add users
 * @param {Array} userIds - all the user id to be added into a group
 * @param {string} token - Slack API token with group management permissions
 * @return {Object} Statistics about adding users to groups
 */
 function addUsersToUsergroup(userGroupId, userIds, token) {
    try {
        const existingUserIds = getUserGroupMembers(userGroupId, token);
        const newUsers = userIds.filter((id) => !existingUserIds.includes(id));
        if (newUsers.length === 0) {
            Logger.log('All users are already in the group. No update needed.');
            return true;
        }
        Logger.log('Adding new users to group: ' + newUsers.join(', '));
        const url = 'https://slack.com/api/usergroups.users.update';
        const payload = {
            usergroup: userGroupId,
            users: existingUserIds.concat(newUsers).join(','),
        };
        const options = {
            method: 'post',
            headers: {
                Authorization: 'Bearer ' + token,
                'Content-Type': 'application/json',
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true,
        };
        const response = fetchWithRetry(url, options);
        const result = JSON.parse(response.getContentText());
        if (!result.ok) {
            Logger.log(`Error updating user group: ${result.error}`);
            return false;
        } else {
            Logger.log('Successfully added users to the group.');
            return true;
        }
    } catch (error) {
        logToSlack(`Error in addUsersToUsergroup: ${error}`);
        return false;
    }
}

/**
 * Gets all the users in the group by group id
 * @param {string} userGroupId The id of the group to be checked
 * @param {string} token - Slack API token with group management permissions
 * @return {Object} Returns all userID  of a group
 */
 function getUserGroupMembers(userGroupId, token) {
    try {
        const url = `https://slack.com/api/usergroups.users.list?usergroup=${userGroupId}`;
        const options = {
            method: 'get',
            headers: {Authorization: 'Bearer ' + token},
            muteHttpExceptions: true,
        };
        const response = fetchWithRetry(url, options);
        const result = JSON.parse(response.getContentText());
        return result.ok ? result.users : [];
    } catch (error) {
        logToSlack(`Error in getUserGroupMembers: ${error}`);
        return [];
    }
}

/**
 * Generates a Slack-compatible handle from a given name string.
 * Cleans the name by converting to lowercase, replacing spaces and invalid
 * characters, handling consecutive/leading/trailing separators, and truncating.
 * If the cleaning process results in an empty string, a default handle is
 * generated and the event is logged using logToSlack.
 *
 * @param {string} name - The input name string.
 * @returns {string} A cleaned, Slack-compatible handle.
 * @requires logToSlack - Global function for logging messages.
 */
 function generateSlackHandle(name) {
    // Basic input validation
    if (!name || typeof name !== 'string') {
        const defaultHandle = 'default-user-' + Date.now();
        // Log the reason for using the default (invalid input)
        // Ensure logToSlack is accessible here
        logToSlack(
            `Invalid input provided for handle generation (received: ${name}). Assigned default handle: "${defaultHandle}".`
        );
        return defaultHandle;
    }

    const originalName = name; // Keep original name for logging if needed
    let handle = name.toLowerCase();

    // Replace whitespace with hyphens
    handle = handle.replace(/\s+/g, '-');
    // Remove disallowed characters (keep a-z, 0-9, ., -, _)
    handle = handle.replace(/[^a-z0-9._-]/g, '');
    // Replace multiple consecutive separators (-, _, .) with a single hyphen
    handle = handle.replace(/[-_.]{2,}/g, '-');
    // Remove leading/trailing separators
    handle = handle.replace(/^[-_.]+|[-_.]+$/g, '');
    // Truncate to 21 characters (adjust if Slack's limit differs)
    handle = handle.substring(0, 21);

    // Final check: If handle is empty after cleaning, generate default and log
    if (!handle) {
        const defaultHandle = 'default-user-' + Date.now(); // Generate default
        // Log the event using the global logToSlack function
        logToSlack(
            `Original name "${originalName}" resulted in an empty handle after cleaning. Assigned default handle: "${defaultHandle}".`
        );
        return defaultHandle; // Return the generated default
    }

    // Return the cleaned handle if it's valid
    return handle;
}


/**
 * Sends a direct message to a specific Slack user ID using fetchWithRetry.
 *
 * @param {string} userId - The Slack User ID to send the message to.
 * @param {string} message - The text message to send.
 * @param {string} token - Slack Bot token with chat:write permission.
 * @return {boolean} True if the message was sent successfully (API returned ok: true), false otherwise.
 */
function sendDirectMessageToUser(userId, message, token) {
    if (!userId || !message || !token) {
        Logger.log("sendDirectMessageToUser: Missing userId, message, or token.");
        return false;
    }
    try {
        const url = `https://slack.com/api/chat.postMessage`;
        const payload = {
            channel: userId, // For DMs, the channel is the User ID
            text: message,
            link_names: true, // Ensures @mentions work if needed, though not used here
        };

        const options = {
            method: "post",
            headers: {
                Authorization: "Bearer " + token,
                "Content-Type": "application/json; charset=utf-8",
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true,
        };

        const response = fetchWithRetry(url, options); // Use shared fetchWithRetry
        const result = JSON.parse(response.getContentText());

        if (result.ok) {
            Logger.log(`Successfully sent DM to user ID: ${userId}`);
            return true;
        } else {
            Logger.log(
                `Error sending DM to user ID ${userId}: ${
                    result.error || "Unknown error"
                }. Response: ${response.getContentText()}`
            );
            return false;
        }
    } catch (error) {
        logToSlack(
            `Exception occurred in sendDirectMessageToUser (User ID: ${userId}): ${error}`
        );
        return false;
    }
}