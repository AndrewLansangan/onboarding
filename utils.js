/**
 * Reads the sheet and creates a Map of team names to arrays of user emails
 * Filters out users with invalid statuses (To Verify, Completed, Archived)
 * @param {Sheet} sheet - The Google Sheet to read from
 * @return {Map} Map with team names as keys and arrays of emails as values
 */
/**
 * Reads user data from a Google Sheet and groups users by their active team names.
 * Filters out users with invalid mandate statuses and already processed entries based on `lastRunTime`.
 *
 * @return {Map<string, string[]>} - A map of team names to arrays of user emails.
 * @param dateStr
 */

function parseCustomDate(dateStr) {
    let [day, month, year] = String(dateStr).split('/').map(Number);
    return new Date(year, month - 1, day);
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
 * External dependency (shared utility).
 * Defined in: `global-utils.gs` or similar.
 *
 * @function fetchWithRetries
 * @param {string} url
 * @param {Object} options
 * @param maxRetries
 * @returns {HTTPResponse}
 */
// fetchWithRetry() is assumed to be globally available

function fetchWithRetries(url, options, maxRetries = 3) {
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const response = UrlFetchApp.fetch(url, options);
        const data = JSON.parse(response.getContentText());

        if (data.ok) return response; // ✅ Success
        else if (shouldRetry(data.error)) {
            Utilities.sleep(2 ** attempt * 1000); // wait 1s, 2s, 4s...
        } else {
            break; // non-retriable error
        }
    }
    return lastFailedResponse;
}

function formatDateTimeForSheet(dateValue) {
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

function convertDate(isoDateString) {
    const date = new Date(isoDateString);
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return date.toLocaleDateString('en-US', options);
}