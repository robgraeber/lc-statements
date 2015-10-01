'use strict';
const Constants = {
    ClientId: process.env.CLIENT_ID,
    ClientSecret: process.env.CLIENT_SECRET,
    RefreshToken: process.env.REFRESH_TOKEN,
    LendingClubEmail: process.env.LENDING_CLUB_EMAIL,
    LendingClubPassword: process.env.LENDING_CLUB_PASSWORD,
    GoogleSpreadsheetName: process.env.GOOGLE_SPREADSHEET_NAME,
    GoogleWorksheetName: process.env.GOOGLE_WORKSHEET_NAME
};

module.exports = Constants;