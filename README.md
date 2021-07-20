This is a mail thread attendance counter, which takes in a subject string `SEARCH_SUBJECT` and searches for specific `SEARCH_MSG` and `SEARCH_MSG_NEG` and exports all the names using the `snippet` of the mail content to a `.xls` sheet.

It also reads the `snippet` for the reason in case of a `SEARCH_MSG_NEG` hit and exports those to the excel sheet as well.

It finally makes a table in the excel summarising the attendance on that day. 

The file `autoGenAttendance.xls` is not been committed to protect the privacy of the members. `credentials.json` and `token.json` have not been committed as they contain the API token for the Gmail API used in the script.

You can implement the idea for your organization (or anywhere you think you would need it). Check out the [Gmail API Documentation](https://developers.google.com/gmail/api).
