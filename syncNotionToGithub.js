function syncNotionToGitHub() {
    const notionApiKey = PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
    const githubToken = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    const databaseId = 'YOUR_NOTION_DATABASE_ID';
    const repoOwner = 'your-org';
    const repoName = 'your-repo';

    const notionPages = fetchNotionPages(notionApiKey, databaseId);

    notionPages.forEach(page => {
        const title = extractTitle(page);
        const body = extractBody(page);
        const githubIssueId = getSyncedGithubId(page.id);

        if (githubIssueId) {
            updateGithubIssue(githubIssueId, title, body, githubToken, repoOwner, repoName);
        } else {
            const issueId = createGithubIssue(title, body, githubToken, repoOwner, repoName);
            storeMapping(page.id, issueId);
        }
    });
}
