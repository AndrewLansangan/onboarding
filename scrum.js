/**
 * This script checks the Notion Team Directory database for teams whose mandate
 * is marked as complete. For each completed team, it identifies the Scrum Master
 * via a relation property, fetches their email from the related People Directory page,
 * finds their Slack ID, and notifies them via direct message on Slack.
 * It prompts them to review the team and potentially disable the Slack user group.
 * It uses Script Properties to track teams already notified to prevent duplicates.
 *
 * Team Directory DB: https://www.notion.so/grey-box/70779d3ee3cf467b9b86171acabc3321
 * People Directory DB: https://www.notion.so/grey-box/People-da052a0ffb3a428d8e7013c540c42665
 * Notion Page for Script: https://www.notion.so/grey-box/Notify-SM-on-Mandate-Completion-YOUR_PAGE_ID_HERE?pvs=4
 */

// --- Configuration ---
// Store these values in Script Properties (File > Project properties > Script properties)
// NOTION_API_KEY: Your Notion integration token
// SLACK_BOT_TOKEN: Your Slack Bot token with chat:write and users:read.email permissions
// NOTION_TEAM_DB_ID: The ID of the Team Directory Notion database (e.g., 70779d3ee3cf467b9b86171acabc3321)
// SLACK_LOGGING_CHANNEL_ID: The Slack channel ID for logging script errors/info.
// SILENCETEAMSWITHOUTSCRUM: Will stop the script from sending the warning if no scrum is assigned to a team


const SILENCETEAMSWITHOUTSCRUM = false; //Stops the warning about teams with no scrum.


// --- Constants ---
// Property names in the *Team Directory* Database
const TEAM_DB_ID_PROP_NAME = "NOTION_TEAM_DB_ID";
const TEAM_NAME_PROPERTY = "Name"; // Title property
const TEAM_STATUS_PROPERTY = "Status"; // Status or Select property
const SCRUM_MASTER_RELATION_PROPERTY = "Scrum Master"; // Relation property linking to People DB

// Property names in the *People Directory* Database (used when fetching related SM page)
const SM_EMAIL_PROPERTY_IN_PEOPLE_DB = "Email (Org)"; // Email property on the Person page

// Values
const COMPLETED_STATUS_VALUE = "Completed"; // The value indicating completion in TEAM_STATUS_PROPERTY
const NOTIFIED_TEAMS_PROPERTY_KEY = "notifiedCompletedTeamIds"; // Key for storing notified IDs in Script Properties

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

// --- Notion Helper Functions ---

/**
 * Builds the payload for the Notion API query to filter by a specific property and value.
 * Assumes the property type is 'status' based on common usage for this script's purpose.
 * If your 'Status' property in Notion is actually a 'Select' type, change "status" below to "select".
 *
 * @param {string|null} startCursor - The cursor for pagination, or null to start.
 * @param {string} filterProperty - The name of the property (e.g., "Status").
 * @param {string} filterValue - The value to filter by (e.g., "Complete").
 * @return {Object} The payload object for the Notion API request.
 */
function buildNotionFilterPayload(startCursor, filterProperty, filterValue) {
    // *** Correction: Use the property TYPE ('status') as the key, not the property NAME ('Status') ***
    // If your Notion property is actually a 'Select' type, change 'status' to 'select' below.
    const payload = {
        page_size: 100, // Max allowed by Notion API
        filter: {
            property: filterProperty, // The name of the property to filter (e.g., "Status")
            status: {
                // The TYPE of the property being filtered
                equals: filterValue, // The condition (e.g., "Complete")
            },
        },
    };

    if (startCursor) {
        payload.start_cursor = startCursor;
    }

    // Logger.log(`Notion Payload: ${JSON.stringify(payload)}`); // Keep for debugging if needed
    return payload;
}

/**
 * Fetches data from the Notion API using a POST request with retry logic.
 * (Uses fetchWithRetry internally)
 *
 * @param {string} apiUrl - The Notion API endpoint URL.
 * @param {Object} headers - The request headers including Authorization.
 * @param {Object} payload - The request payload (body).
 * @return {Object} The parsed JSON response from Notion.
 * @throws {Error} If the API call fails after retries or returns a Notion error object.
 */
function fetchNotionDataWithRetry(apiUrl, headers, payload) {
    const options = {
        method: "post",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
    };

    const response = fetchWithRetry(apiUrl, options); // Uses shared fetchWithRetry
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
        try {
            return JSON.parse(responseBody);
        } catch (e) {
            throw new Error(`Failed to parse Notion API response: ${e.message}`);
        }
    } else {
        let errorMessage = `Notion API request failed (${apiUrl}) with status code ${responseCode}.`;
        try {
            const errorResponse = JSON.parse(responseBody);
            errorMessage += ` Notion Error: ${errorResponse.message || responseBody}`;
        } catch (e) {
            errorMessage += ` Response Body: ${responseBody}`;
        }
        Logger.log(errorMessage); // Log detailed error
        throw new Error(errorMessage); // Throw to be caught by main loop
    }
}

/**
 * Processes the raw JSON response from the Notion API query.
 * Checks for Notion-specific error objects.
 *
 * @param {Object} responseJson - The parsed JSON response from Notion query.
 * @return {Object} The processed response, typically { results, has_more, next_cursor }.
 * @throws {Error} If the response object indicates an error or unexpected structure.
 */
function processNotionResponse(responseJson) {
    if (responseJson.object && responseJson.object === "error") {
        Logger.log(`Notion API Error: ${JSON.stringify(responseJson)}`);
        throw new Error(
            responseJson.message ||
            `Notion API returned an error: ${responseJson.code}`
        );
    }
    if (!responseJson.results) {
        Logger.log(
            `Unexpected Notion API response structure: ${JSON.stringify(
                responseJson
            )}`
        );
        throw new Error(
            "Unexpected Notion API response structure (missing results)."
        );
    }
    return responseJson;
}

/**
 * Extracts relation IDs from a Notion relation property object.
 *
 * @param {Object} property - The Notion relation property object.
 * @param {string} propertyName - The name of the property (for logging).
 * @return {string[]} An array of relation page IDs, or an empty array if none found.
 */
function extractRelationIds(property, propertyName) {
    if (
        !property ||
        property.type !== "relation" ||
        !property.relation ||
        property.relation.length === 0
    ) {
        // Logger.log(`Property '${propertyName}' is not a valid relation or is empty.`);
        return [];
    }
    return property.relation.map((rel) => rel.id);
}

/**
 * Fetches a specific property value from a specific Notion page ID.
 * Used here to get the SM's email from their Person page.
 *
 * @param {string} pageId - The ID of the Notion page to fetch.
 * @param {string} propertyToExtract - The exact name of the property to extract from the page.
 * @param {string} notionApiKey - The Notion API key.
 * @return {string|null} The value of the specified property, or null if not found/error.
 */
function fetchPageProperty(pageId, propertyToExtract, notionApiKey) {
    const apiUrl = `https://api.notion.com/v1/pages/${pageId}`;
    const headers = {
        Authorization: `Bearer ${notionApiKey}`,
        "Notion-Version": "2022-06-28",
    };
    const options = {
        method: "get",
        headers: headers,
        muteHttpExceptions: true,
    };

    try {
        const response = fetchWithRetry(apiUrl, options); // Use retry mechanism
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode >= 200 && responseCode < 300) {
            const pageData = JSON.parse(responseBody);
            if (pageData.properties && pageData.properties[propertyToExtract]) {
                // Use extractPropertyValue to handle different types correctly
                const value = extractPropertyValue(
                    pageData.properties[propertyToExtract],
                    propertyToExtract
                );
                return value; // Returns string, number, boolean, or empty string
            } else {
                Logger.log(
                    `Property '${propertyToExtract}' not found on page ID ${pageId}.`
                );
                return null;
            }
        } else {
            Logger.log(
                `Error fetching page ${pageId}: Status ${responseCode}. Response: ${responseBody}`
            );
            return null;
        }
    } catch (error) {
        logToSlack(
            `Exception fetching page property '${propertyToExtract}' for page ID ${pageId}: ${error}`
        );
        return null;
    }
}

/**
 * Extracts a property value from a Notion page property object.
 * (Adapted from your provided functions to be reusable here)
 *
 * @param {Object} property - The Notion property object.
 * @param {string} propertyName - The name of the property (for logging).
 * @return {string|number|boolean} The extracted value, or an empty string if not found or type is unhandled.
 */
function extractPropertyValue(property, propertyName) {
    if (!property) {
        return "";
    }
    const type = property.type;
    let value = "";

    try {
        switch (type) {
            case "title":
                value =
                    property.title && property.title.length > 0
                        ? property.title[0].plain_text
                        : "";
                break;
            case "rich_text":
                value =
                    property.rich_text && property.rich_text.length > 0
                        ? property.rich_text.map((item) => item.plain_text).join("")
                        : "";
                break;
            case "number":
                value = property.number !== null ? property.number : "";
                break;
            case "select":
                value = property.select ? property.select.name : "";
                break;
            case "status":
                value = property.status ? property.status.name : "";
                break;
            case "email": // *** Crucial for getting SM email ***
                value = property.email || "";
                break;
            case "date":
                value = property.date && property.date.start ? property.date.start : "";
                break;
            case "checkbox":
                value = property.checkbox; // Returns true/false
                break;
            // Add other cases if needed (multi_select, people, files, url, phone_number, formula, relation (just IDs), created_time, etc.)
            // For this script, we primarily need title, status, email, and relation (handled by extractRelationIds)
            default:
                Logger.log(
                    `Unhandled property type: '${type}' for property '${propertyName}' in extractPropertyValue.`
                );
                value = "";
        }
    } catch (e) {
        Logger.log(
            `Error extracting value for property '${propertyName}' (Type: ${type}): ${e}`
        );
        value = "";
    }
    // Return primitive types directly, convert others to string if needed for consistency elsewhere
    return typeof value === "boolean" || typeof value === "number"
        ? value
        : String(value);
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