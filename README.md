# lc-statements
Lending Club statement fetcher that writes to Google Sheets. 

Useful for fetching down your latest statement from lending club, then writing via api to your google spreadsheet. Intended to be setup with a cronjob or Heroku's scheduler.

### Usage:

```
git clone https://github.com/robgraeber/lc-statements.git && cd lc-statements 
(setup env variables)
npm start
```

Env variables required:  
`CLIENT_ID`: Google Sheets Client Id  
`CLIENT_SECRET`: Google Sheets Client Secret  
`REFRESH_TOKEN`: Google Sheets Client Refresh Token  
`LENDING_CLUB_EMAIL`: Lending Club Email  
`LENDING_CLUB_PASSWORD`: Lending Club Password  
`GOOGLE_SPREADSHEET_NAME`: Spreadsheet Name  
`GOOGLE_WORKSHEET_NAME`: Worksheet Name  
