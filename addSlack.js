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

const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
const SLACK_USER_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_USER_TOKEN');
const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const sheetName = PropertiesService.getScriptProperties().getProperty('SHEET_NAME');
const lastRunTime = PropertiesService.getScriptProperties().getProperty('LAST_RUN_TIME');

function runScript() {
    logToSlack(
        "📢 Starting execution of \`addGroupsToSlack\` script"
    );
    let externalSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = externalSpreadsheet.getSheetByName(sheetName);

    const groupsMap = readUsers(sheet);

    for (let [groupName, emailList] of groupsMap.entries()) {
        try {
            let userGroupId = checkForExistingGroups(groupName, SLACK_USER_TOKEN);

            if (!userGroupId) {
                Logger.log(`Group '${groupName}' does not exist. Creating it...`);
                let groupData = createUserGroup(groupName, SLACK_USER_TOKEN);

                if (groupData && groupData.usergroup && groupData.usergroup.id) {
                    userGroupId = groupData.usergroup.id;
                    Logger.log(`Group '${groupName}' created with ID: ${userGroupId}`);
                    Utilities.sleep(2000);
                } else {
                    Logger.log(`Failed to create group '${groupName}', skipping...`);
                    continue;
                }
            } else {
                Logger.log(`Group '${groupName}' already exists with ID: ${userGroupId}`);
            }

            const userIds = [];
            for (let email of emailList) {
                let userId = getUserIdByEmail(email, SLACK_USER_TOKEN);
                if (userId) {
                    userIds.push(userId);
                } else {
                    Logger.log(`User not found or error for email: ${email}`);
                }
            }

            let existingUserIds = getUserGroupMembers(userGroupId, SLACK_USER_TOKEN);
            const newUsers = userIds.filter(id => !existingUserIds.includes(id));

            if (newUsers.length === 0) {
                Logger.log(`All users already in group '${groupName}'. Skipping update.`);
                continue;
            }

            Logger.log(`Adding users to group '${groupName}': ${newUsers.join(', ')}`);
            let success = addUsersToUsergroup(userGroupId, newUsers, SLACK_USER_TOKEN);

            if (success) {
                Logger.log(`Successfully added users to group '${groupName}'`);
            } else {
                Logger.log(`Failed to add users to group '${groupName}'`);
            }

        } catch (err) {
            Logger.log(`Error processing group '${groupName}': ${err}`);
        }
    }
    PropertiesService.getScriptProperties().setProperty('LAST_RUN_TIME', Date.now().toString());
    Logger.log("Finished processing all groups and users.");
    logToSlack(
        "📢 Execution of \`addGroupsToSlack\` script finished"
    );
}

/**
 * Reads the sheet and creates a Map of team names to arrays of user emails
 * Filters out users with invalid statuses (To Verify, Completed, Archived)
 * @param {Sheet} sheet - The Google Sheet to read from
 * @return {Map} Map with team names as keys and arrays of emails as values
 */
function readUsers(sheet) {
    Logger.log(`Starting to read users from sheet: ${sheet.getName()}`);
    let data = sheet.getDataRange();
    let values = data.getValues();
    const header = values.shift();
    const invalidStatuses = new Set(['To Verify', 'Completed', 'Archived']);
    const emailIndex = header.indexOf('Email (Org)');
    const statusIndex = header.indexOf('Mandate (Status)');
    const teamIndex = header.indexOf('Team (Current)');
    const lastUpdateIndex = header.indexOf("Last Update");

    Logger.log(
        `Found columns - Email: ${emailIndex}, Status: ${statusIndex}, Team: ${teamIndex}`,
    );

    let groupsMap = new Map();
    let processedRows = 0;
    let skippedRows = 0;

    for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const email = row[emailIndex];
        const status = row[statusIndex];
        const usergroupCell = row[teamIndex];
        const lastUpdate = row[lastUpdateIndex];
        const parsedDate = parseCustomDate(lastUpdate);

        if (parsedDate < lastRunTime) {
            Logger.log("User already processed skipping");
            continue;
        }

        if (!invalidStatuses.has(status)) {
            const usergroups = usergroupCell.split(',').map((group) => group.trim());
            Logger.log(
                `Processing user ${email} with ${usergroups.length} team assignments`,
            );
            processedRows++;

            usergroups.forEach((userGroup) => {
                if (userGroup === '' || userGroup.toLowerCase().includes('admin')) {
                    Logger.log(`skipping ${userGroup} because it's an admin group.`)
                    return;
                }

                if (groupsMap.has(userGroup)) {
                    const emails = groupsMap.get(userGroup);
                    emails.push(email);
                    Logger.log(`Added ${email} to existing group ${userGroup}`);
                } else {
                    groupsMap.set(userGroup, [email]);
                    Logger.log(
                        `Created new group ${userGroup} with first member ${email}`,
                    );
                }
            });
        } else {
            Logger.log(`Skipping user ${email} with status: ${status}`);
            skippedRows++;
        }
    }

    Logger.log(
        `Finished processing ${processedRows} users, skipped ${skippedRows} users`,
    );
    Logger.log(`Created ${groupsMap.size} total groups`);
    return groupsMap;
}
function parseCustomDate(dateStr) {
    let [day, month, year] = String(dateStr).split('/').map(Number);
    return new Date(year, month - 1, day);
}


/**
 * Creates a Slack user group with the specified name and a generated handle.
 * Checks if the group exists before attempting to create it.
 * Logs events, including handle generation defaults via generateSlackHandle.
 *
 * @param {string} name - The desired name for the user group.
 * @param {string} token - Slack API token with group creation permissions.
 * @return {Object | undefined} Response object from Slack API or undefined on error.
 * @requires generateSlackHandle - Function to generate Slack-compliant handles.
 * @requires checkForExistingGroups - Function to check for existing groups.
 * @requires fetchWithRetry - Function for making HTTP requests with retries.
 * @requires logToSlack - Global function for logging messages.
 * @requires Logger - Assumed Google Apps Script Logger or similar.
 */
function createUserGroup(name, token) {
    // Input validation for the name itself
    if (!name || typeof name !== 'string' || name.trim() === '') {
        const errorMessage = `Invalid group name provided: "${name}". Cannot create group.`;
        Logger.log(errorMessage);
        // Assuming logToSlack is defined and accessible
        logToSlack(errorMessage);
        return { ok: false, error: 'invalid_group_name_provided' };
    }

    const trimmedName = name.trim(); // Use trimmed name for consistency

    try {
        Logger.log(`Attempting to find or create user group: "${trimmedName}"`);
        // Check using trimmed name
        const groupId = checkForExistingGroups(trimmedName, token);

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
                    Authorization: 'Bearer ' + token,
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

/**
 * Checks if a user group with the specified name already exists
 * @param {string} name - The name of the user group to check
 * @param {string} token - Slack API token with group read permissions
 * @return {string|boolean} Group ID if found, false otherwise
 */
function checkForExistingGroups(name, token) {
    try {
        Logger.log(`Checking if group exists: ${name}`);
        const url = 'https://slack.com/api/usergroups.list';
        const options = {
            method: 'get',
            contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + token },
            muteHttpExceptions: true,
        };
        const response = fetchWithRetry(url, options);
        const data = JSON.parse(response.getContentText());

        if (!data.ok) {
            Logger.log(
                `Slack API error: ${response.getResponseCode()} - ${data.error}`,
            );
            return false;
        }

        const group = data.usergroups.find((group) => group.name === name);
        return group ? group.id : false;
    } catch (error) {
        logToSlack(`Error checking for existing groups: ${error}`);
        return false;
    }
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
            headers: { Authorization: 'Bearer ' + token },
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
            headers: { Authorization: 'Bearer ' + token },
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
