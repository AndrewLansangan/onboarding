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

    const pagesDatabase1 = fetchAllPages(databaseId1, headers);
    const pagesDatabase2 = fetchAllPages(databaseId2, headers);

    if (!pagesDatabase1.length || !pagesDatabase2.length) {
        logToSlack("Failed to fetch pages from one or both databases. Please check the fetchAllPages function and the database IDs.");
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
// --- Slack Helper Functions ---

/**
 * Gets the Slack userID for a user by their email using fetchWithRetry.
 * @param {string} email - Email of the user.
 * @param {string} token - Slack API token with users:read.email permission.
 * @return {string|null} Slack UserID or null if not found or error occurs.
 */
function getUserIdByEmail(email, token) {
    if (!email || !token) {
        Logger.log("getUserIdByEmail: Missing email or token.");
        return null;
    }
    try {
        const url = `https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(
            email
        )}`;
        const options = {
            method: "get",
            headers: { Authorization: "Bearer " + token },
            muteHttpExceptions: true,
        };

        const response = fetchWithRetry(url, options); // Use shared fetchWithRetry
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
function extractTeamDirectoryRow(page) {
    const properties = page.properties;

    const name = extractPropertyValue(properties["Name"], "Name");
    const status = extractPropertyValue(properties["Status"], "Status");
    const people = extractPropertyValue(properties["People (Current)"], "People (Current)");
    const scrumMaster = extractPropertyValue(properties["Scrum Master"], "Scrum Master");
    const activityEpic = extractPropertyValue(properties["Activity (Epic)"], "Activity (Epic)");
    const dateEpic = extractPropertyValue(properties["Date (Epic)"], "Date (Epic)");

    return [name, status, people, scrumMaster, activityEpic, dateEpic];
}

function extractPropertyValue(property, propertyName) {
    if (!property) {
        logToSlack(`Property '${propertyName}' is undefined.`);
        return "";
    }

    const type = property.type;
    let value = "";

    switch (type) {
        case 'title':
            value = property.title.length > 0 ? property.title[0].plain_text : "";
            break;
        case 'number':
            value = property.number !== null ? property.number : "";
            break;
        case 'select':
            value = property.select ? property.select.name : "";
            break;
        case 'multi_select':
            value = property.multi_select.length > 0 ? property.multi_select.map(item => item.name).join(", ") : "";
            break;
        case 'email':
            value = property.email || "";
            break;
        case 'date':
            if (property.date && property.date.start) {
                value = formatDate(property.date.start);
            }
            break;
        case 'created_time':
            if (property.created_time) {
                value = new Date(property.created_time).toLocaleDateString('en-CA');
            }
            break;
        case 'rich_text':
            value = property.rich_text.length > 0 ? property.rich_text.map(item => item.plain_text).join("\n") : "";
            break;
        case 'status':
            value = property.status ? property.status.name : "";
            break;
        case 'relation':
            value = property.relation.length > 0 ? property.relation.map(item => item.name || item.plain_text).join(", ") : "";
            break;
        case 'formula':
            if (property.formula.string) {
                value = property.formula.string;
            } else if (property.formula.number !== null) {
                value = property.formula.number;
            } else if (property.formula.boolean !== null) {
                value = property.formula.boolean.toString();
            }
            break;
        default:
            logToSlack(`Unhandled property type: ${type}`);
            value = "";
    }

    return value;
}

/**
 * The `timeTrackerRecap` function updates the "Notion Database (sync) - Mandates" Google Sheet with data from the
 * "All Recap - Current and Archived" (Google Sheet).
 *  Notion Database (sync) - Mandates (Google Sheet) : https://docs.google.com/spreadsheets/d/1uqCK0JDHKkuzDEfOlBcSqJCsVxHjCRouOW5F8hJZRUA/edit?gid=2131663677#gid=2131663677
 *  All Recap - Current and Archived" (Google Sheet) : https://docs.google.com/spreadsheets/d/1jnXb0JTCy6C0ORsGkepJBzQbPYIZHpBWO0gqvQVOvX4/edit?gid=0#gid=0
 *  It matches the "Greybox ID" from the target sheet with the "Recap"
 * name in the external sheet and retrieves the following columns: 'Error Detection', 'Last Update', 'Start Date',
 * and 'Hours (decimal)'. The data is then written into columns N to Q of the target sheet, ensuring accurate
 * and up-to-date information is reflected for each corresponding entry.
 */


function timeTrackerRecap() {
    logToSlack(
        "📢 Starting execution of \`TimeTrackerRecap\` script"
    );
    const externalSpreadsheetId = "1jnXb0JTCy6C0ORsGkepJBzQbPYIZHpBWO0gqvQVOvX4";
    const externalSheetName = "All Recap - Current and Archived";
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Ensure that the relevant columns exist in the target sheet
    const headers = ['Error Detection', 'Last Update', 'Start Date', 'Hours (decimal)'];
    for (let i = 0; i < headers.length; i++) {
        targetSheet.getRange(1, 14 + i).setValue(headers[i]); // Starting at column N (14th column)
    }

    const externalSs = SpreadsheetApp.openById(externalSpreadsheetId);
    const externalSheet = externalSs.getSheetByName(externalSheetName);
    if (!externalSheet) {
        logToSlack("External sheet not found.");
        return;
    }

    const externalData = externalSheet.getRange("A2:G" + externalSheet.getLastRow()).getValues();

    // Create a map of names to their corresponding data
    const nameToDataMap = {};
    externalData.forEach(row => {
        const name = row[0] ? row[0].trim().toLowerCase() : ""; // Recap names are in column A, normalized to lowercase and trimmed
        if (name) {
            nameToDataMap[name] = {
                errorDetection: row[2], // Error Detection in column C
                lastUpdate: formatDateTime(row[3]),     // Last Update in column D, format to "DD/MM/YYYY HH:MM:SS"
                startDate: formatDateTime(row[4]),      // Start Date in column E, format to "DD/MM/YYYY HH:MM:SS"
                hoursDecimal: row[5]    // Hours (decimal) in column F
            };
        }
    });

    // Iterate over 'Greybox ID' in the target sheet to prepare data for updating
    const targetDataRange = targetSheet.getRange("A2:A" + targetSheet.getLastRow());
    const targetNames = targetDataRange.getValues();

    const dataToWrite = targetNames.map(row => {
        const name = row[0] ? row[0].trim().toLowerCase() : ""; // Normalize and trim the name
        const data = nameToDataMap[name] || {
            errorDetection: "",
            lastUpdate: "",
            startDate: "",
            hoursDecimal: ""
        };
        return [data.errorDetection, data.lastUpdate, data.startDate, data.hoursDecimal];
    });

    // Update columns N to Q with the extracted data
    targetSheet.getRange("N2:Q" + (1 + dataToWrite.length)).setValues(dataToWrite);
    logToSlack(
        "📢 Execution of \`TimeTrackerRecap\` script finished."
    );
}


function formatDateTime(dateValue) {
    if (!dateValue) return "";

    // Check if the dateValue is a Date object, convert it to the desired string format
    if (Object.prototype.toString.call(dateValue) === '[object Date]') {
        return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }

    // If dateValue is already a string, assume it's in the correct format and return it as is
    if (typeof dateValue === 'string') {
        return dateValue;
    }

    return ""; // Return an empty string if dateValue is not valid
}
