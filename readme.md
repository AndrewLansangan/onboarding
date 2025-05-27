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

// --- Notion Helper Functions ---

