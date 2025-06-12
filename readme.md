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

# 🔄 Sync Summary

This document provides an overview of the main synchronization functions between **Notion**, **Google Sheets**, and **Slack** in the script.

---

## 📥 `syncNotionPeopleDirectoryToGoogleSheet`

**From:**  
→ Notion Database: `People Directory`

**To:**  
→ Google Sheet: `Mandates`

**Purpose:**  
Fetches paginated data from Notion, parses it, and writes it into the `Mandates` sheet.

**Dependencies:**
- `Config.notionDbPeople`
- `fetchNotionData()`
- `handleNotionApiResponse()`
- `parseNotionPageRow()`

---

## 📤 `syncGoogleSheetToSlack`

**From:**  
→ Google Sheet: `Mandates`

**To:**  
→ Slack Profiles

**Purpose:**  
Reads users from the sheet, finds their Slack ID, and updates their Slack profile fields.

**Dependencies:**
- `extractUserDataFromSheet()`
- `getSlackUserIdByEmail()`
- `updateUserProfile()`

---

## 🔁 `syncGoogleSheetToNotion`

**From:**  
→ Google Sheet: `Mandates`

**To:**  
→ Notion Database: `People Directory`

**Purpose:**  
Updates `Hours (Current)` and `Hours (Last Update)` properties in Notion if values changed in the sheet.

**Dependencies:**
- `extractNotionPageId()`
- `getNotionPageProperties()`
- `updateNotionPageProperties()`

---

## 📥 `syncTeamDirectoryToSheet`

**From:**  
→ Notion Database: `Team Directory`

**To:**  
→ Google Sheet: `Teams`

**Purpose:**  
Fetches team data from Notion and syncs it to the corresponding sheet.

**Dependencies:**
- `NOTION_QUERY_URL()`
- `fetchNotionData()`
- `processNotionResponse()`
- `transformNotionTeamPageToRow()`

---
