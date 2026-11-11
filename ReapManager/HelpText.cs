namespace ReapManager;

/// Text of the in-app Help tab. Keep in sync with the "User guide" in README.md.
static class HelpText
{
    public static readonly (string Title, string Body)[] Sections =
    [
        ("Check these 3 things before changing anything",
         "The blue header at the top always shows:\n" +
         "  - Language badge (orange), e.g. 'Language: Tamil (ta)'\n" +
         "  - Auth folder, e.g. 'reap_Tamil'\n" +
         "  - YouTube channel: the real channel name from Google after you sign in\n" +
         "All three must belong to the same language. The window title and the bottom-right of the status bar show it too."),

        ("1. One-time setup: one auth folder per language",
         "Every language is a separate YouTube channel, so each language needs its own auth folder with that channel's files:\n\n" +
         "  D:\\REAP Auth\\reap_English\\     oauth_client.json  +  api.txt\n" +
         "  D:\\REAP Auth\\reap_Hindi\\       oauth_client.json  +  api.txt\n" +
         "  D:\\REAP Auth\\reap_Tamil\\       oauth_client.json  +  api.txt\n" +
         "  ... one folder for each of the 12 languages\n\n" +
         "Tip: put the language name in the folder name (reap_Tamil, REAP Tamil, ...). The app then picks the right language automatically.\n" +
         "Keep these folders private. Never upload oauth_client.json, api.txt or token files anywhere."),

        ("2. First time for each language",
         "a. Choose the language in 'Channel Language' (tab 1).\n" +
         "b. Click 'Browse...' and select that language's auth folder. The app remembers it for this language.\n" +
         "c. Click 'Authorize (Browser)'. Your web browser opens: log in with THAT channel's Google account and allow access.\n" +
         "d. The light in the bottom-left turns green and the header shows the channel name.\n" +
         "A saved login (token) is stored inside the auth folder, so next time no browser is needed.\n" +
         "If the folder already has token.json from the old Python app, just click 'Authorize with Token'.\n" +
         "Repeat a-d once for every language you work on. At first the folder list is empty; it fills up as you add folders."),

        ("3. Everyday use",
         "Pick the language from 'Channel Language', or pick a saved folder from the 'Auth Folder' list (it shows every folder with its language).\n" +
         "The app signs out of the previous channel, loads that language's folder and signs in with the saved token automatically\n" +
         "(untick 'Sign in automatically' if you prefer to click 'Authorize with Token' yourself).\n" +
         "Check the header (language, folder, channel) and then use tab 2 or tab 3."),

        ("4. Rename & Describe (tab 2)",
         "Select a playlist and click 'Load Videos'.\n" +
         "Titles in bold are not in the format 'Chapter N - Topic' and will be renamed.\n" +
         "Check the 'Corrected' and 'Description' columns; double-click a cell to edit it.\n" +
         "Click 'Apply Changes' and confirm. Only the bold rows are updated on YouTube. Results are in the Log tab."),

        ("5. Excel Export (tab 3)",
         "Select a playlist and click 'Load Video Info' to preview the rows.\n" +
         "Click 'Generate Excel Sheet'. The file is saved in '<auth folder>\\Excel sheet\\' and the full path is shown.\n" +
         "The language code comes from the end of the playlist title (e.g. NARQ30056_ta -> ta). If it does not match the selected\n" +
         "language you get a warning, so a Tamil sheet is not made from a Hindi channel by mistake."),

        ("6. Problems and fixes",
         "'oauth_client.json is missing' / 'api.txt is missing': wrong folder selected, or the file is missing in it.\n" +
         "'The saved login expired or was revoked': click 'Authorize (Browser)' and log in again.\n" +
         "Login expires about every 7 days: the Google Cloud OAuth app is in 'Testing' mode. Fix once in Google Cloud Console:\n" +
         "  APIs & Services > OAuth consent screen > Publish app > 'In production'.\n" +
         "Wrong channel name in the header: you logged in with the wrong Google account. Click 'Sign out / Reset', then delete the file\n" +
         "  'Google.Apis.Auth.OAuth2.Responses.TokenResponse-user' (and token.json) in that auth folder and use 'Authorize (Browser)' again.\n" +
         "Windows says 'Windows protected your PC': click 'More info' > 'Run anyway' (the app is not code-signed)."),

        ("7. Reporting a problem",
         "Open the Log tab, tick 'Show debug details', repeat what went wrong, then click 'Copy all' and send it.\n" +
         "Or send the newest file from the logs folder (full path in the Log and About tabs; 'Open logs folder' opens it).\n" +
         "Logs never contain passwords or keys."),
    ];
}
