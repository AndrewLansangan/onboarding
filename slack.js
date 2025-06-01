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
 * - `checkForExistingSlackGroups`: Verifies whether a specified Slack user group already exists.
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
 * - `syncGoogleSheetToSlack`: Main function that processes the data from the sheet and updates Slack profiles.
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
 * @requires checkForExistingSlackGroups - Function to check for existing groups.
 * @requires fetchWithRetries - Function for making HTTP requests with retries.
 * @requires logToSlack - Global function for logging messages.
 * @requires Logger - Assumed Google Apps Script Logger or similar.
 */

/**
 * ========================================
 * 🧩 Slack Integration Master Script
 * ========================================
 *
 * This file contains all logic for integrating Slack with Google Sheets and Notion,
 * including user group syncing, profile updating, direct messaging, and utility helpers.
 *
 * ----------------------------------------
 * 🔁 SECTION 1: Sync Slack User Groups from Sheet
 * ----------------------------------------
 * Synchronizes Google Sheet teams to Slack user groups:
 * - Reads user data from sheet (`readUsers`)
 * - Filters users based on status ('To Verify', 'Completed', 'Archived')
 * - Creates user groups if missing (`createUserGroup`)
 * - Adds users to corresponding groups (`addUsersToUsergroup`)
 * - Logs all major events to Slack and Logger
 *
 * 🧾 Required Sheet Columns:
 * - 'Email (Org)', 'Mandate (Status)', 'Team (Current)'
 *
 * 🔐 Auth:
 * - Requires user token with `usergroups:read` and `usergroups:write`
 *
 * ----------------------------------------
 * 🔄 SECTION 2: Sync Slack Profiles from Sheet
 * ----------------------------------------
 * Updates Slack user profiles with metadata from Notion-based sheet:
 * - Fields: Position, Team, Notion Page URL, Mandate Status, Availability
 * - Time tracker fields: Last Update, Total Hours
 *
 * 📄 Data Source:
 * - "Notion Database (sync) - Mandates" Google Sheet
 *
 * 🔧 Key Functions:
 * - `syncGoogleSheetToSlack`
 * - `extractUserDataFromSheet`
 * - `updateUserProfile`
 *
 * 🔐 Auth:
 * - Requires user token with `users.profile:write`
 *
 * ----------------------------------------
 * ⚙️ SECTION 3: Slack Utility Functions
 * ----------------------------------------
 * Shared helper functions used across all Slack operations:
 *
 * - `getSlackUserIdByEmail(email)`
 *   → Looks up Slack user ID by email using `users.lookupByEmail`
 *
 * - `generateSlackHandle(name)`
 *   → Normalizes team/user name into valid Slack handle
 *
 * - `constructProfileFields(fields)`
 *   → Maps data to Slack profile field IDs
 *
 * - `createUserGroup(name, token)`
 *   → Creates a Slack user group after checking for duplicates
 *
 * - `checkForExistingSlackGroups(name, token)`
 *   → Verifies if a group already exists and returns its ID
 *
 * - `getUserGroupMembers(groupId, token)`
 *   → Returns list of user IDs currently in a user group
 *
 * - `addUsersToUsergroup(groupId, userIds, token)`
 *   → Adds missing users to the group
 *
 * - `sendDirectMessageToUser(userId, message, token)`
 *   → Sends DM to user via `chat.postMessage`
 *
 * - `fetchWithRetry(url, options)`
 *   → Makes HTTP requests with retry logic (assumed defined globally)
 *
 * - `logToSlack(message)`
 *   → Sends logs to configured Slack channel
 *
 * ----------------------------------------
 * 🔐 Authentication Overview
 * ----------------------------------------
 * - `SLACK_USER_TOKEN`: Used for most write operations (group/user updates)
 * - `SLACK_BOT_TOKEN`: Used for logging and messaging
 * - Tokens are stored securely in Script Properties
 */

/**
* Gets the Slack user ID for a user by their email using fetchWithRetry.
* @param {string} email - Email of the user.
* @returns {string|null} Slack user ID or null if not found or an error occurs.
*/
function getSlackUserIdByEmail(email) {
    if (!email || !SLACK_USER_TOKEN) {
        Logger.log("getUserIdByEmail: Missing email or token.");
        return null;
    }
    try {
        const url = `https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(
            email
        )}`;
        const options = {
            method: "get",
            headers: {Authorization: "Bearer " + SLACK_USER_TOKEN},
            muteHttpExceptions: true,
        };

        const response = fetchWithRetries(url, options); // Use shared fetchWithRetry
        const result = JSON.parse(response.getContentText());

        if (result.ok && result.user && result.user.id) {
            return result.user.id;
        } else {
            Logger.log(
                `Error finding Slack user ID by email (${email}): ${
                    result.error || "Unknown error"
                }. Response: ${response.getContentText()}`
            );
            return null;
        }
    } catch (error) {
        // Log error using the provided logToSlack
        logToSlack(
            `Exception occurred while retrieving user ID by email (${email}): ${error}`
        );
        return null;
    }
}

/**
 * Generates a sanitized Slack handle from a display name.
 *
 * - Converts name to lowercase
 * - Replaces spaces with hyphens
 * - Removes disallowed characters
 * - Trims and truncates to 21 characters
 * - Falls back to a default if resulting handle is empty
 *
 * @param {string} name - Raw display name or team name.
 * @returns {string} A Slack-compatible handle.
 *
 * @example
 * generateSlackHandle("Grey Box Team")  // "grey-box-team"
 */
function generateSlackHandle(name) {
    if (!name || typeof name !== 'string') {
        const defaultHandle = 'default-user-' + Date.now();
        logToSlack?.(`⚠️ Invalid input for Slack handle. Assigned default: ${defaultHandle}`);
        return defaultHandle;
    }

    const originalName = name;
    let handle = name.toLowerCase()
        .replace(/\s+/g, '-')                  // spaces to dashes
        .replace(/[^a-z0-9._-]/g, '')          // remove illegal characters
        .replace(/[-_.]{2,}/g, '-')            // squash multiple separators
        .replace(/^[-_.]+|[-_.]+$/g, '')       // trim leading/trailing
        .substring(0, 21);                     // truncate to Slack limit

    if (!handle) {
        const defaultHandle = 'default-user-' + Date.now();
        logToSlack?.(`⚠️ Name "${originalName}" resulted in empty handle. Fallback: ${defaultHandle}`);
        return defaultHandle;
    }

    return handle;
}

// ====================================
// ======= Slack User Profiles ========
// ====================================

/**
 * Constructs Slack custom profile field JSON from provided user field data.
 *
 * @param {Object} fields - Object containing Slack profile field values.
 * @return {Object} Slack API compatible profile.fields payload.
 */
function constructProfileFields(fields) {
    return {
        fields: {
            "Xf06JZK27DRA": {value: fields.Position},
            "Xf03V366R202": {value: fields.Team},
            "Xf06JGJMBZPZ": {value: fields.NotionPageURL},
            "Xf0759PXS7BP": {value: fields.MandateStatus},
            "Xf074Y4V1KHV": {value: fields.Availability},
            "Xf075CJ4SXEF": {value: fields.CreatedProfile},
            "Xf07HUS9GSSC": {value: fields.TimeTrackerLastUpdate},
            "Xf07GZDPHHV4": {value: fields.TimeTrackerTotal}
        }
    };
}

/**
 * Updates a Slack user's profile using custom fields.
 *
 * @param {string} userId - Slack user ID to update.
 * @param {Object} fields - Custom profile field values, used to build Slack payload.
 * @returns {void}
 *
 * @example
 * updateUserProfile("U123456", {
 *   Position: "Developer",
 *   Team: "Engineering",
 *   NotionPageURL: "https://notion.so/...etc"
 * });
 */
function updateUserProfile(userId, fields) {
    const profileFields = constructProfileFields(fields);

    const payload = {
        user: userId,
        profile: JSON.stringify(profileFields)
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        headers: {
            Authorization: `Bearer ${SLACK_USER_TOKEN}`
        },
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch('https://slack.com/api/users.profile.set', options);
        const result = JSON.parse(response.getContentText());

        if (!result.ok) {
            const msg = `❌ Failed to update profile for ${userId}: ${result.error}`;
            Logger.log(msg);
            logToSlack?.(msg);
        }
    } catch (error) {
        logToSlack?.(`🚨 Exception in updateUserProfile(${userId}): ${error}`);
    }
}

function extractUserDataFromSheet(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const [header, ...rows] = sheet.getDataRange().getValues();

    const indexes = {
        email: header.indexOf('Email (Org)'),
        position: header.indexOf('Position'),
        team: header.indexOf('Team (Current)'),
        notionUrl: header.indexOf('Notion Page URL'),
        mandateStatus: header.indexOf('Mandate (Status)'),
        availability: header.indexOf('Availability (avg h/w)'),
        createdProfile: header.indexOf('Created (Profile)'),
        lastUpdate: header.indexOf('Last Update'),
        hoursDecimal: header.indexOf('Hours (decimal)'),
    };

    // Validate column presence
    const missingFields = Object.entries(indexes).filter(([_, i]) => i === -1).map(([k]) => k);
    if (missingFields.length > 0) {
        throw new Error('Missing required columns: ' + missingFields.join(', '));
    }

    return rows
        .filter(row => row[indexes.email])
        .map(row => ({
            email: row[indexes.email],
            Position: row[indexes.position],
            Team: row[indexes.team],
            NotionPageURL: row[indexes.notionUrl],
            MandateStatus: row[indexes.mandateStatus],
            Availability: row[indexes.availability],
            CreatedProfile: row[indexes.createdProfile]?.toString(),
            TimeTrackerLastUpdate: row[indexes.lastUpdate],
            TimeTrackerTotal: row[indexes.hoursDecimal],
        }));
}

/**
 * Creates a Slack user group with the specified name and a generated handle.
 * Checks if the group exists before attempting to create it.
 * Logs events, including handle generation defaults via generateSlackHandle.
 *
 * @param {string} slackGroupName - The desired name for the user group.
 * @param {string} [slackToken=SLACK_USER_TOKEN] - Slack API token with group creation permissions.
 * @return {Object} Response object from Slack API.
 * @requires generateSlackHandle - Function to generate Slack-compliant handles.
 * @requires checkForExistingSlackGroups - Function to check for existing groups.
 * @requires fetchWithRetries - Function for making HTTP requests with retries.
 * @requires logToSlack - Global function for logging messages.
 */
function createUserGroup(slackGroupName, slackToken = SLACK_USER_TOKEN) {
    if (!slackGroupName || typeof slackGroupName !== 'string' || slackGroupName.trim() === '') {
        const errorMessage = `Invalid group name provided: "${slackGroupName}". Cannot create group.`;
        Logger.log(errorMessage);
        logToSlack(errorMessage);
        return {ok: false, error: 'invalid_group_name_provided'};
    }

    const trimmedName = slackGroupName.trim();

    try {
        Logger.log(`Attempting to find or create user group: "${trimmedName}"`);
        const groupId = checkForExistingSlackGroups(trimmedName, slackToken);

        if (groupId) {
            Logger.log(`Group "${trimmedName}" already exists with ID: ${groupId}`);
            return {
                ok: true,
                usergroup: {
                    id: groupId,
                    name: trimmedName,
                },
            };
        } else {
            Logger.log(`Group "${trimmedName}" not found. Proceeding with creation.`);
            const groupHandle = generateSlackHandle(trimmedName);

            Logger.log(`Generated handle for group "${trimmedName}": "${groupHandle}". Creating new group via API.`);

            const url = 'https://slack.com/api/usergroups.create';
            const options = {
                method: 'post',
                headers: {
                    Authorization: 'Bearer ' + slackToken,
                    'Content-Type': 'application/json',
                },
                payload: JSON.stringify({name: trimmedName, handle: groupHandle}),
                muteHttpExceptions: true,
            };

            const response = fetchWithRetries(url, options, 3);
            const data = JSON.parse(response.getContentText());

            if (data.ok) {
                Logger.log(`Successfully created group "${data.usergroup?.name || trimmedName}" with handle "${data.usergroup?.handle || groupHandle}". ID: ${data.usergroup?.id}.`);
            } else {
                const errorMessage = `Failed to create Slack group "${trimmedName}" (handle: "${groupHandle}"). Slack API Error: ${data.error || 'unknown_error'}.`;
                Logger.log(errorMessage);
                logToSlack(errorMessage);
            }

            return data;
        }
    } catch (error) {
        const errorMessage = `Unexpected error in createUserGroup for "${trimmedName}": ${error.message}`;
        Logger.log(`${errorMessage}\n${error.stack || ''}`);
        logToSlack(errorMessage);
        return {ok: false, error: `internal_script_error: ${error.message}`};
    }
}

// ====================================
// ==== Slack Group Membership Ops ====
// ====================================

/**
 * Retrieves all Slack user IDs from a given user group.
 *
 * @param {string} userGroupId - Slack user group ID to query.
 * @param {string} token - Slack API token with `usergroups:read` scope.
 * @returns {string[]} Array of user IDs, or an empty array on failure.
 */
function getUserGroupMembers(userGroupId, token) {
    try {
        const url = `https://slack.com/api/usergroups.users.list?usergroup=${userGroupId}`;
        const options = {
            method: 'get',
            headers: {
                Authorization: `Bearer ${token}`
            },
            muteHttpExceptions: true
        };

        const response = fetchWithRetries(url, options);
        const result = JSON.parse(response.getContentText());

        if (result.ok) {
            return result.users;
        } else {
            Logger.log(`❌ Slack error in getUserGroupMembers: ${result.error}`);
            return [];
        }
    } catch (error) {
        logToSlack?.(`🚨 Exception in getUserGroupMembers(${userGroupId}): ${error}`);
        return [];
    }
}

/**
 * Adds users to a Slack user group if they're not already members.
 *
 * @param {string} userGroupId - Slack group ID to modify.
 * @param {string[]} userIds - Array of Slack user IDs to add.
 * @param {string} token - Slack API token with `usergroups:write` scope.
 * @returns {boolean} True if operation succeeded or no update was needed, false otherwise.
 */
function addUsersToUsergroup(userGroupId, userIds, token) {
    try {
        const existingUserIds = getUserGroupMembers(userGroupId, token);
        const newUsers = userIds.filter(id => !existingUserIds.includes(id));

        if (newUsers.length === 0) {
            Logger.log(`ℹ️ All users already in group ${userGroupId}. No changes made.`);
            return true;
        }

        Logger.log(`🔧 Adding users to group ${userGroupId}: ${newUsers.join(', ')}`);
        const url = 'https://slack.com/api/usergroups.users.update';
        const payload = {
            usergroup: userGroupId,
            users: existingUserIds.concat(newUsers).join(',')
        };
        const options = {
            method: 'post',
            headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const response = fetchWithRetries(url, options);
        const result = JSON.parse(response.getContentText());

        if (!result.ok) {
            Logger.log(`❌ Failed to update group ${userGroupId}: ${result.error}`);
            return false;
        }

        Logger.log(`✅ Successfully updated group ${userGroupId}.`);
        return true;
    } catch (error) {
        logToSlack?.(`🚨 Exception in addUsersToUsergroup(${userGroupId}): ${error}`);
        return false;
    }
}

// ====================================
// ======= Slack Direct Messages ======
// ====================================

/**
 * Sends a direct message to a Slack user via `chat.postMessage`.
 *
 * @param {string} userId - Slack user ID (e.g., "U123456").
 * @param {string} message - Message content to send.
 * @param {string} token - Bot token with `chat:write` permission.
 * @returns {boolean} True if successful, false if error or missing params.
 *
 * @example
 * sendDirectMessageToUser("UABC123", "Welcome to the workspace!", SLACK_BOT_TOKEN);
 */
function sendDirectMessageToUser(userId, message, token = SLACK_BOT_TOKEN) {
    if (!userId || !message || !token) {
        Logger.log("⚠️ sendDirectMessageToUser: Missing parameter(s).");
        return false;
    }

    const url = 'https://slack.com/api/chat.postMessage';
    const payload = {
        channel: userId,
        text: message,
        link_names: true
    };

    const options = {
        method: 'post',
        headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json; charset=utf-8'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const response = fetchWithRetries(url, options);
        const result = JSON.parse(response.getContentText());

        if (result.ok) {
            Logger.log(`✅ DM sent to user ${userId}`);
            return true;
        } else {
            Logger.log(`❌ DM to ${userId} failed: ${result.error}`);
            return false;
        }
    } catch (error) {
        logToSlack?.(`🚨 Exception in sendDirectMessageToUser(${userId}): ${error}`);
        return false;
    }
}

/**
 * Checks if a user group with the specified name already exists in Slack.
 * @param slackGroupName
 * @param {string} slackToken - Slack API token with group read permissions.
 * @return {string|boolean} Group ID if found, false otherwise.
 */
function checkForExistingSlackGroups(slackGroupName, slackToken = SLACK_USER_TOKEN) {
    if (!name || !slackToken) {
        logError("checkForExistingSlackGroups: Name or token is missing.");
        return false;
    }

    try {
        const options = {
            method: 'get',
            headers: {Authorization: `Bearer ${slackToken}`},
            muteHttpExceptions: true,
        };

        const response = fetchWithRetries(SLACK_USERGROUP_LIST_URL, options);
        const data = JSON.parse(response.getContentText());

        if (!data.ok) {
            logError(`Slack API error ${response.getResponseCode()}: ${data.error}`);
            return false;
        }

        const match = data.usergroups.find(g => g.name === name);
        return match?.id || false;
    } catch (error) {
        logError(`checkForExistingSlackGroups error: ${error.message || error}`);
        return false;
    }
}

