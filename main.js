/**
 * This script syncs data from the People Directory (Notion database) to the Notion Database (sync) - Mandates (Google Sheet).
 * People Directory (Notion database): https://www.notion.so/grey-box/People-da052a0ffb3a428d8e7013c540c42665
 * Notion Database (sync) - Mandates (Google Sheet): https://docs.google.com/spreadsheets/d/1uqCK0JDHKkuzDEfOlBcSqJCsVxHjCRouOW5F8hJZRUA/edit?gid=2131663677#gid=2131663677
 *
 * It retrieves entries from the Notion database using the Notion API, processes the data, and updates a Google Sheet named "Mandates".
 * The script handles pagination to retrieve all entries, supports multiple Notion property types, and fetches additional information
 * for relation properties to display related page titles instead of IDs.
 *
 * The script includes:
 * - `syncDataToSheet`: Main function that retrieves data from Notion and updates the Google Sheet.
 * - `constructPayload`: Helper function to build the request payload for pagination.
 * - `retrieveNotionData`: Function to fetch data from Notion using the Notion API.
 * - `getPropertyData`: Function that extracts data from Notion properties, including titles and relations.
 * - `fetchPageTitle`: Function that retrieves the title of related Notion pages for display in the sheet.
 * - `convertDate`: Utility function to convert dates from ISO format to a readable format.
 * - `handleNotionResponse`: Function to process the Notion API response and handle errors.
 * - `parseRowFromPage`: Function to extract and format a row of data from a Notion page.
 *
 * Key Features:
 * - Supports various Notion property types, including text, number, select, multi-select, email, date, and relation.
 * - Dynamically fetches and displays the title of related pages, ensuring user-friendly information in the "Team (Current)" column.
 * - Logs details of missing or undefined properties to assist with debugging.
 * - Utilizes a Notion API key stored securely in the script properties for authenticated API requests.
 * - Gracefully handles errors and logs any issues encountered during the sync process.
 * Notion link: https://www.notion.so/grey-box/Sync-Notion-People-Directory-to-Google-Sheet-syncNotionPeopleDirectoryToSheetsMandates-bca68b43894946e7847267aff967180e
 */

/**
 * Entrypoint: Sync mandate data from Google Sheets to Slack user profiles.
 * - Extracts data from the Google Sheet (via `extractUserDataFromSheet`)
 * - Gets Slack user ID (via `getUserIdByEmail`)
 * - Updates Slack profile fields (via `updateUserProfile`)
 */
/**
 * Synchronizes the "People Directory" Notion database with the "Mandates" sheet in Google Sheets.
 *
 * 🔁 This replaces the previous `syncDataToSheet()` function and consolidates:
 * - `retrieveNotionData()` → now uses `fetchNotionData()`
 * - `parseRowFromPage()` → now `parseNotionPageRow()` for clarity
 *
 * The script:
 * - Fetches all pages from the specified Notion database with pagination support
 * - Parses each page into a row of relevant values
 * - Clears and repopulates the "Mandates" sheet with up-to-date data
 *
 * ⚠️ Make sure the headers in the sheet match the parsed values
 * ✅ Sheet will always be cleared and reloaded from scratch
 *
 * Dependencies:
 * - `fetchNotionData(apiUrl, headers, payload)`
 * - `handleNotionApiResponse(response)`
 * - `parseNotionPageRow(page)`
 * - `logToSlack(message)`
 */

function syncNotionPeopleDirectoryToGoogleSheet() {
    logToSlack(
        "📢 Starting execution of \`syncNotionPeopleDirectoryToSheetsMandates\` script"
    );
    const databaseId = Config.notionDbPeople;
    const apiUrl = `https://api.notion.com/v1/databases/${databaseId}/query`;
    const headers = Config.notionHeaders;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.getSheetConfig("MANDATES"));
    sheet.clearContents();

    sheet.appendRow(SHEET_HEADERS.MANDATES);

    let allRows = [];
    let hasMore = true;
    let startCursor = null;

    while (hasMore) {
        const payload = constructPayload(startCursor);

        try {
            const response = fetchNotionData(apiUrl, headers, payload);
            const {results, has_more, next_cursor} = handleNotionApiResponse(response);

            const rows = results.map(page => parseNotionPageRow(page));
            allRows = allRows.concat(rows);

            hasMore = has_more;
            startCursor = next_cursor;
        } catch (error) {
            logToSlack(`Error fetching data: ${error}`);

            break;
        }
    }

    if (allRows.length > 0) {
        const startRow = 2;
        const startColumn = 1;
        const numRows = allRows.length;
        const numColumns = allRows[0].length;
        sheet.getRange(startRow, startColumn, numRows, numColumns).setValues(allRows);
    }

    logToSlack("Sync completed!");
    logToSlack(
        "📢 Execution of \`syncNotionPeopleDirectoryToSheetsMandates\` script finished"
    );
}

/**
 * Synchronizes data from the "Mandates" Google Sheet to Slack user profiles.
 *
 * Extracts each user's data, gets their Slack ID, and updates their profile fields.
 *
 * @function syncGoogleSheetToSlack
 * @returns {void}
 */
function syncGoogleSheetToSlack() {
    const { header, rows } = loadSheetData("Mandates");


    if (!headersMatch(header, SHEET_HEADERS.MANDATES)) {
        logToSlack("⚠️ Header mismatch in Mandates sheet. Aborting sync.");
        return;
    }

    logToSlack("📢 Starting syncSheetsMandatesToSlackGreyBox...");
//reads user data sheet process them
    const users = extractUserDataFromSheet(SHEET_NAMES.MANDATES);
    users.forEach(user => {
        const userId = getSlackUserIdByEmail(user.email);
        if (userId) {
            updateUserProfile(userId, constructProfileFields(user));
        }
    });

    logToSlack("✅ Finished syncSheetsMandatesToSlackGreyBox.");
}

/**
 * Synchronizes data from the "Mandates" Google Sheet to Notion's "People Directory" database.
 *
 * For each row in the sheet, updates "Hours (Current)" and "Hours (Last Update)" in Notion,
 * if changes are detected.
 *
 * @function syncGoogleSheetToNotion
 * @returns {void}
 */
function syncGoogleSheetToNotion() {

    const { header, rows } = loadSheetData("Mandates");


    if (!headersMatch(header, SHEET_HEADERS.MANDATES)) {
        logToSlack("⚠️ Header mismatch in Mandates sheet. Aborting sync.");
        return;
    }

    logToSlack(
        "📢 Starting execution of \`syncSheetsMandatesToNotionPeopleDirectory\` script"
    );
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MANDATES);
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const header = data.shift(); // Remove the header row

    const notionUrlIndex = header.indexOf('Notion Page URL');
    const hoursCurrentIndex = header.indexOf('Hours (decimal)');
    const lastUpdateIndex = header.indexOf('Last Update');

    data.forEach((row, rowIndex) => {
        const notionUrl = row[notionUrlIndex];
        const hoursCurrent = row[hoursCurrentIndex];
        const lastUpdate = row[lastUpdateIndex];

        // Skip row if both relevant Google Sheet values are null/empty
        if (!hoursCurrent && !lastUpdate) {
            logInfo(`Skipping row ${rowIndex + 2} as both Hours and Last Update are empty.`);
            return;
        }

        const notionPageId = extractNotionPageId(notionUrl);
        if (!notionPageId) {
            logInfo(`Skipping row ${rowIndex + 2} as Notion Page ID is missing or invalid.`);
            return;
        }

        // Fetch current Notion page properties
        const notionProperties = getNotionPageProperties(notionPageId);

        // Prepare properties to update in Notion
        const propertiesToUpdate = {};

        if (hoursCurrent) {
            const roundedHoursCurrent = parseFloat(hoursCurrent.toFixed(1));
            if (notionProperties['Hours (Current)']?.number !== roundedHoursCurrent) {
                propertiesToUpdate['Hours (Current)'] = {number: roundedHoursCurrent};
            }
        }

        if (lastUpdate) {
            const formattedLastUpdate = new Date(lastUpdate).toISOString();
            if (notionProperties['Hours (Last Update)']?.date?.start !== formattedLastUpdate) {
                propertiesToUpdate['Hours (Last Update)'] = {date: {start: formattedLastUpdate}};
            }
        }

        // Update Notion only if there are changes
        if (Object.keys(propertiesToUpdate).length > 0) {
            const updateSuccess = updateNotionPageProperties(notionPageId, propertiesToUpdate);
            if (updateSuccess) {
                logInfo(`Successfully updated Notion page ${notionPageId}`);
            } else {
                logToSlack(`Failed to update Notion page ${notionPageId}`);
            }
        } else {
            logInfo(`No changes detected for Notion Page ID: ${notionPageId}. Skipping update.`);
        }
    });

    logToSlack(
        "📢 Execution of \`syncSheetsMandatesToNotionPeopleDirectory\` script finished"
    );
}

/**
 * Synchronizes the Notion Team Directory database to a Google Sheet.
 *
 * Fetches paginated data from the Notion Team database and writes it to the sheet.
 *
 * @function syncTeamDirectoryToSheet
 * @returns {void}
 */
function syncTeamDirectoryToSheet() {
    logInfo("📢 Starting execution of `syncTeamDirectoryToSheet` script");

    const apiUrl = NOTION_QUERY_URL(NOTION_TEAM_DB_ID);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TEAM_DIRECTORY);

    sheet.clearContents();
    sheet.appendRow(TEAM_DIRECTORY_COLUMNS);

    let allRows = [];
    let hasMore = true;
    let startCursor = null;

    while (hasMore) {
        const payload = buildNotionPayload(startCursor);

        try {
            const response = fetchNotionData(apiUrl, NOTION_HEADERS, payload);
            const { results, has_more, next_cursor } = processNotionResponse(response);

            const rows = results.map(page => transformNotionTeamPageToRow(page));
            allRows = allRows.concat(rows);

            hasMore = has_more;
            startCursor = next_cursor;
        } catch (error) {
            logError(`Error fetching data: ${error}`);
            break;
        }
    }

    if (allRows.length > 0) {
        sheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);
    } else {
        logInfo("No rows found in Notion Team Directory Database to sync.");
    }

    logInfo("📢 `syncTeamDirectoryToSheet` script finished");
}

/**
 * Checks for completed team mandates in Notion and notifies the Scrum Master on Slack.
 */
function notifyScrumMasterOfCompletedMandates() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const notionApiKey = scriptProperties.getProperty("NOTION_API_KEY");
    const teamDatabaseId = scriptProperties.getProperty(TEAM_DB_ID_PROP_NAME);
    const slackbottoken = scriptProperties.getProperty('SLACK_BOT_TOKEN');

    logToSlack(
        "📢 Starting execution of \`notifyScrumIfTeamMandateIsComplete\` script"
    );

    // Check essential configuration
    if (!notionApiKey || !slackbottoken || !teamDatabaseId) {
        const errorMsg =
            "🚨 *Critical Error*: Missing required script properties (`NOTION_API_KEY`, `SLACK_BOT_TOKEN`, `NOTION_TEAM_DB_ID`). Please configure them.";
        Logger.log(errorMsg);
        // Try logging to Slack if possible, otherwise just log locally
        if (loggingChannelId && slackbottoken) {
            logToSlack(errorMsg); // Use the provided logToSlack function
        }
        return;
    }

    logToSlack(
        "🚀 Starting check for completed team mandates in order to notify scrum masters..."
    ); // Log start

    if(SILENCETEAMSWITHOUTSCRUM){
        logToSlack("⚠️ Notifications about teams without scrum master is currently *disabled* change the flag \`SILENCETEAMSWITHOUTSCRUM\` on the script to enable them.");
    }

    // Load list of already notified team IDs
    let notifiedTeamIds = [];
    try {
        const notifiedJson = scriptProperties.getProperty(
            NOTIFIED_TEAMS_PROPERTY_KEY
        );
        if (notifiedJson) {
            notifiedTeamIds = JSON.parse(notifiedJson);
        }
    } catch (e) {
        logToSlack(
            `⚠️ Error parsing notified teams list from Script Properties: \`${e}\`. Starting with empty list.`
        );
        notifiedTeamIds = [];
    }

    const apiUrl = `https://api.notion.com/v1/databases/${teamDatabaseId}/query`;
    const headers = {
        Authorization: `Bearer ${notionApiKey}`,
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    };

    let hasMore = true;
    let startCursor = null;
    let teamsProcessedCount = 0;
    let notificationsSentCount = 0;
    let newNotifications = false; // Flag to track if we need to save updated list

    while (hasMore) {
        const payload = buildNotionFilterPayload(
            startCursor,
            TEAM_STATUS_PROPERTY,
            COMPLETED_STATUS_VALUE
        );

        try {
            const responseJson = fetchNotionDataWithRetry(apiUrl, headers, payload);
            const { results, has_more, next_cursor } =
                processNotionResponse(responseJson);

            Logger.log(`Fetched ${results.length} results from Notion Team DB.`);

            for (const teamPage of results) {
                const teamId = teamPage.id;
                const teamName = extractPropertyValue(
                    teamPage.properties[TEAM_NAME_PROPERTY],
                    TEAM_NAME_PROPERTY
                );

                if (!teamName) {
                    Logger.log(
                        `Skipping page ID ${teamId}: Could not extract team name.`
                    );
                    continue;
                }

                // Check if already notified
                if (notifiedTeamIds.includes(teamId)) {
                    Logger.log(
                        `Skipping team "${teamName}" (ID: ${teamId}): Already notified.`
                    );
                    continue; // Skip this team
                }

                teamsProcessedCount++;
                Logger.log(`Processing completed team: "${teamName}" (ID: ${teamId})`);

                // 1. Extract Scrum Master Relation ID
                const smRelationIds = extractRelationIds(
                    teamPage.properties[SCRUM_MASTER_RELATION_PROPERTY],
                    SCRUM_MASTER_RELATION_PROPERTY
                );

                if (!smRelationIds || smRelationIds.length === 0) {
                    if(!SILENCETEAMSWITHOUTSCRUM){
                        logToSlack(
                            `⚠️ No Scrum Master relation found for completed team *${teamName}* (ID: \`${teamId}\`). Cannot notify.`
                        );
                    }
                    continue;
                }

                // Assuming the first relation is the correct SM
                const scrumMasterPageId = smRelationIds[0];
                if (smRelationIds.length > 1) {
                    logToSlack(
                        `⚠️ Multiple SM relations for team *${teamName}* (ID: \`${teamId}\`). Using first: \`${scrumMasterPageId}\`.`
                    );
                }

                // 2. Fetch SM Email from related Person Page
                const scrumMasterEmail = fetchPageProperty(
                    scrumMasterPageId,
                    SM_EMAIL_PROPERTY_IN_PEOPLE_DB,
                    notionApiKey
                );

                if (!scrumMasterEmail || !validateEmail(scrumMasterEmail)) {
                    logToSlack(
                        `⚠️ Could not get valid email for SM (Page ID: \`${scrumMasterPageId}\`) for team *${teamName}* (ID: \`${teamId}\`). Email found: \`${
                            scrumMasterEmail || "None"
                        }\`. Cannot notify.`
                    );
                    continue;
                }

                Logger.log(
                    `Found SM Email for team "${teamName}": ${scrumMasterEmail}`
                );

                // 3. Get Slack User ID
                const slackUserId = getUserIdByEmail(
                    scrumMasterEmail,
                    slackbottoken
                );

                if (slackUserId) {
                    Logger.log(
                        `Found Slack User ID for ${scrumMasterEmail}: ${slackUserId}`
                    );
                    // 4. Send Slack DM
                    const teamNotionLink = `https://www.notion.so/${teamId.replace(
                        /-/g,
                        ""
                    )}`;

                    const message = `Hi <@${slackUserId}>! The mandate for the team *${teamName}* has been marked as '${COMPLETED_STATUS_VALUE}' in Notion. Please review the team's status and consider disabling the corresponding Slack user group if appropriate.\nTeam Page: <${teamNotionLink}|${teamName}>`;

                    const success = sendDirectMessageToUser(
                        slackUserId,
                        message,
                        slackbottoken
                    );

                    const actionDescription = "Mandate Completion Notification";

                    if (success) {
                        notificationsSentCount++;
                        // Success Log
                        logToSlack(
                            `✅ Successfully notified \`${scrumMasterEmail}\` for team *${teamName}* that ${actionDescription} was *sent*.`
                        );
                        // 5. Mark as notified
                        notifiedTeamIds.push(teamId);
                        newNotifications = true; // Mark that we need to save the updated list
                    } else {
                        // Failure Log
                        logToSlack(
                            `❌ Failed to notify \`${scrumMasterEmail}\` (Slack ID: \`${slackUserId}\`) for team *${teamName}* that ${actionDescription} was *sent*. Check GAS logs for DM error details.`
                        );
                    }
                } else {
                    // Use logToSlack for warnings
                    logToSlack(
                        `⚠️ Could not find Slack User ID for SM email: \`${scrumMasterEmail}\` (Notion Page ID: \`${scrumMasterPageId}\`) for team *${teamName}*. Cannot send DM.`
                    );
                }
            } // End loop through results

            hasMore = has_more;
            startCursor = next_cursor;
        } catch (error) {
            logToSlack(
                `🚨 *Critical Error* fetching/processing Notion data: \`${error}\``
            );
            hasMore = false; // Stop processing on critical error
        }
    } // End while(hasMore)

    // Save the updated list of notified team IDs if new notifications were sent
    if (newNotifications) {
        try {
            scriptProperties.setProperty(
                NOTIFIED_TEAMS_PROPERTY_KEY,
                JSON.stringify(notifiedTeamIds)
            );
            Logger.log("Successfully updated the list of notified team IDs.");
        } catch (e) {
            logToSlack(
                `🚨 *Critical Error*: Failed to save updated notified teams list to Script Properties: \`${e}\``
            );
        }
    }

    logToSlack(
        `🏁 Script finished. Processed ${teamsProcessedCount} new completed teams. Sent ${notificationsSentCount} notifications.`
    );
}


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
 * @param {string} [purpose] - Optional token purpose: 'bot', 'user', or 'default'. Uses Config.getSlackToken(purpose).- Slack API token with group creation permissions.
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
 /**
 * Gets the Slack user ID for a user by their email using fetchWithRetry.
 * @param {string} email - Email of the user.
 * @param purpose
 * @returns {string|null} Slack user ID or null if not found or an error occurs.
 */
function getSlackUserIdByEmail(email, purpose = 'user') {
    const token = Config.getSlackToken(purpose);

    if (!email || !token) {
        Logger.log("getUserIdByEmail: Missing email or token.");
        return null;
    }

    try {
        const url = `https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(email)}`;
        const options = {
            method: "get",
            headers: { Authorization: "Bearer " + token },
            muteHttpExceptions: true,
        };

        const response = fetchWithRetries(url, options);
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
        logToSlack(`Exception occurred while retrieving user ID by email (${email}): ${error}`);
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
            Authorization: `Bearer ${Config.slackUserToken}`
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
 * @param {{BOT_TOKEN: *, USER_TOKEN: *, LOGGING_CHANNEL_ID: *}} [slackToken=Config.slackUserToken] - Slack API token with group creation permissions.
 * @return {Object} Response object from Slack API.
 * @requires generateSlackHandle - Function to generate Slack-compliant handles.
 * @requires checkForExistingSlackGroups - Function to check for existing groups.
 * @requires fetchWithRetries - Function for making HTTP requests with retries.
 * @requires logToSlack - Global function for logging messages.
 */
function createUserGroup(slackGroupName, slackToken = Config.slackUserToken) {
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
 * @param {{BOT_TOKEN: *, USER_TOKEN: *, LOGGING_CHANNEL_ID: *}} slackUserToken - Slack API token with group read permissions.
 * @return {string|boolean} Group ID if found, false otherwise.
 */
function checkForExistingSlackGroups(slackGroupName, slackUserToken = Config.slackUserToken) {
    if (!name || !slackUserToken) {
        logError("checkForExistingSlackGroups: Name or token is missing.");
        return false;
    }

    try {
        const options = {
            method: 'get',
            headers: {Authorization: `Bearer ${slackUserToken}`},
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


// Main function to link pages across databases
function linkDatabases() {
    logToSlack(
        "📢 Starting execution of \`syncNotionPeopleRelations\` script"
    );
    const config = initializeConfig();

    if (!config) {
        logToSlack("Script aborted due to missing configuration.");
        return;  // Exit if configuration is not valid
    }

    const { headers, databaseId1, databaseId2 } = config;

    logToSlack("Starting the process to link databases.");

    const pagesDatabase1 = fetchAllNotionPages(databaseId1, headers);
    const pagesDatabase2 = fetchAllNotionPages(databaseId2, headers);

    if (!pagesDatabase1.length || !pagesDatabase2.length) {
        logToSlack("Failed to fetch pages from one or both databases. Please check the fetchAllNotionPages function and the database IDs.");
        return;
    }

    Logger.log("Mapping 'Email (Org)' to page IDs for database 2");
    const emailToPageIdMap = pagesDatabase2.reduce((map, page) => {
        const email = (page.properties["Email (Org)"]?.email || "").toLowerCase();
        if (email) {
            if (!map[email]) {
                map[email] = []; // Initialize array if not already created
            }
            map[email].push(page.id); // Add the page ID to the array
        }
        return map;
    }, {});

    logToSlack("Iterating through pages in database 1 to update relations.");
    pagesDatabase1.forEach(page => {
        const email = (page.properties["Email (Org)"]?.email || "").toLowerCase();
        if (email && emailToPageIdMap[email]) {
            Logger.log(`Found matching email: ${email}`);
            if (!page.properties["People Directory (Sync)"]?.relation?.length) {
                const relatedPageIds = emailToPageIdMap[email]; // Get all matching page IDs

                // Prepare the relation payload with multiple IDs
                const relationPayload = relatedPageIds.map(pageId => ({ id: pageId }));

                Logger.log(`Updating page ${page.id} with relations: ${JSON.stringify(relationPayload)}`);

                const success = updatePageRelationWithMultiple(page.id, relationPayload, headers); // Updated function to handle multiple relations
                if (!success) {
                    logToSlack(`Failed to link page with email: ${email}`);
                }
            } else {
                Logger.log(`Page with email ${email} is already linked.`);
            }
        }
    });
    logToSlack("Linking process completed.");
    logToSlack(
        "📢 Execution of \`syncNotionPeopleRelations\` script finished"
    );
}
