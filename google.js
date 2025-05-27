function syncTeamDirectoryToSheet() {
    logInfo("📢 Starting execution of `syncTeamDirectoryToSheet` script");

    const apiUrl = NOTION_QUERY_URL(NOTION_TEAM_DB_ID);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEAM_SHEET_NAME);
    sheet.clearContents();
    sheet.appendRow(TEAM_DIRECTORY_COLUMNS);

    let allRows = [];
    let hasMore = true;
    let startCursor = null;

    while (hasMore) {
        const payload = buildPayload(startCursor);

        try {
            const response = fetchNotionData(apiUrl, NOTION_HEADERS, payload);
            const { results, has_more, next_cursor } = processNotionResponse(response);

            const rows = results.map(page => extractTeamDirectoryRow(page));
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
    }

    logInfo("📢 `syncTeamDirectoryToSheet` script finished");
}
