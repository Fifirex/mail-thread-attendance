# Mail attendance

This is a mail thread attendance counter, which takes in a subject string `SEARCH_SUBJECT` and searches for specific `SEARCH_MSG` and `SEARCH_MSG_NEG` and exports all the names using the `snippet` of the mail content to a `.xls` sheet.

It also reads the `snippet` for the reason in case of a `SEARCH_MSG_NEG` hit and exports those to the excel sheet as well.

It finally makes a table in the excel summarising the attendance on that day. For example:

DATE | 15 July 2021
------------ | -------------
**RES_CTR** | 37
**PLUS_CTR** | 31
**MINUS_CTR** | 6

---

## Data required

There are 2 main json files that act as look up dictionaries for the script:

### 1. database/index-map.json

```js
{<name> : [<disctionary_index>, <used_flag>, <sub_team>(optional)]}
```

Used for indexing the final excel and determining the "No-Responses" with the flag. The `sub_team` is a string type that tells if the person belongs to a specific team is not called during the meet.

### 2. database/name-map.json

```js
{<email_snippet> : <name>}
```

Used for looking up the name of an email alias. It is kept to hold several alias for the same name while avoiding duplicates in the excel.

---

> The directory `database` (which contains the output excel `autoGenAttendance.xls` and the email databse) is not been committed to protect the privacy of the membersof the organization. `credentials.json` and `token.json` have not been committed as they contain the API token for the Gmail API used in the script.

You can implement the idea for your organization (or anywhere you think you would need it). Check out the [Gmail API Documentation](https://developers.google.com/gmail/api).
