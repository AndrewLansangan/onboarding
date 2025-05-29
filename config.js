/**
 * This script facilitates the linking of pages between the People Directory Notion databases and the internal HR-only version by establishing relations based on email addresses.
 * People Directory (Notion database) : https://www.notion.so/grey-box/People-da052a0ffb3a428d8e7013c540c42665
 * People Directory (Internal) (Notion database) : https://www.notion.so/grey-box/47fbed712f3e4558b032edb9ec081f00?v=2d969f6b09084313823bca813f39db69
 *
 * Key Components:
 * - Configuration: Retrieves Notion API credentials and initializes database IDs for the two databases to be linked.
 * - `fetchAllPages`: Fetches all pages from a specified Notion database, handling pagination to ensure complete data retrieval.
 * - `updatePageRelationWithMultiple`: Updates a specific relation property of a page in Notion, linking it to one or more pages based on shared email addresses.
 * - `linkDatabases`: Main function that orchestrates the linking process by:
 *   - Fetching pages from both databases.
 *   - Mapping emails from one database to corresponding page IDs.
 *   - Iterating over the pages in the first database and updating their relation properties to link them with the corresponding pages in the second database.
 *   - Handles cases where multiple matches are found by linking all corresponding relations to the page.
 *
 * The script is designed to run in sequence, ensuring that database pages are fetched and processed in a manner that maintains data integrity
 * and avoids redundant updates. Logging is extensively used to track the progress and identify any issues during execution.
 *
 * Notion link: https://www.notion.so/grey-box/Sync-Relation-Notion-Team-Directory-with-People-Directory-syncNotionPeopleRelations-gs-a906389d2dd440b6a65c6ffe0130787e
 */

function getScriptConfig() {
    const props = PropertiesService.getScriptProperties();
    return {
        NOTION_API_KEY: props.getProperty('NOTION_API_KEY'),
        SLACK_BOT_TOKEN: props.getProperty('SLACK_BOT_TOKEN'),
        SLACK_USER_TOKEN: props.getProperty('SLACK_USER_TOKEN'),
        LOGGING_CHANNEL_ID: props.getProperty('LOGGING_CHANNEL_ID'),
        NOTION_PEOPLE_DB_ID: props.getProperty('NOTION_PEOPLE_DB_ID'),
        NOTION_INTERNAL_PEOPLE_DB_ID: props.getProperty('NOTION_INTERNAL_PEOPLE_DB_ID'),
        NOTION_TEAM_DB_ID: props.getProperty('NOTION_TEAM_DB_ID'),
        LAST_RUN_TIME: props.getProperty('LAST_RUN_TIME'),
        SPREADSHEET_ID: props.getProperty('SPREADSHEET_ID'),
        SHEET_NAME: props.getProperty('SHEET_NAME'),
        NOTIFIED_TEAMS_PROPERTY_KEY: 'notifiedCompletedTeamIds',
    };
}
const [SLACK_USER_TOKEN] = getScriptConfig()
// RUN linkDatabases first
// Configuration
function initializeConfig() {
    const {
        NOTION_API_KEY,
        NOTION_PEOPLE_DB_ID,
        NOTION_INTERNAL_PEOPLE_DB_ID
    } = getScriptConfig();

    if (!NOTION_API_KEY || !NOTION_PEOPLE_DB_ID || !NOTION_INTERNAL_PEOPLE_DB_ID) {
        logError("❌ Missing configuration values in Script Properties.");
        return null;
    }

    logInfo("✅ Notion configuration loaded.");
    logDebug(`People DB: ${NOTION_PEOPLE_DB_ID}, Internal DB: ${NOTION_INTERNAL_PEOPLE_DB_ID}`);

    return {
        headers: {
            Authorization: `Bearer ${NOTION_API_KEY}`,
            'Notion-Version': '2022-06-28',
            'Content-Type': 'application/json',
        },
        databaseId1: NOTION_PEOPLE_DB_ID,
        databaseId2: NOTION_INTERNAL_PEOPLE_DB_ID,
    };
}
