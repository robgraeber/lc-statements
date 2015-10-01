'use strict';
const Util2 = require('./Util2');
const Logger = require('winston2');
const Promise = require('bluebird');
const Constants = require('./Constants');
const SheetHelper = require('./SheetHelper');
const Spreadsheet = require('edit-google-spreadsheet');
const RequestLib = require('request-promise');
const Request = RequestLib.defaults({
    headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.1 Safari/537.36'
    },
    followAllRedirects: true,
    jar: true
});
Promise.promisifyAll(Spreadsheet);

Logger.info('Starting script');
/**
 * Fetches LendingClub balance + net deposits from latest statement
 * @example fetchLCStatementInfo(6, 2015) // Promise => {balance: 6500.00, deposits: 500.00}
 * @returns {Promise.<Object>} Rsp object with deposits and balance
 */
const fetchLCStatementInfo = (month, year) => {
    Logger.info('Fetching LC statement:', month+'-'+year);
    return Request.post({
        url: 'https://www.lendingclub.com/account/login.action',
        form: {
            login_url: '',
            login_email: Constants.LendingClubEmail,
            login_password: Constants.LendingClubPassword
        }
    }).then(rsp => {
        Logger.info('LC login rsp:', rsp.trim().replace(/(\r\n|\n|\r)/gm,""));
        const urlTemplate = 'https://www.lendingclub.com/account/monthlyStatementDownload.action?file_extension=pdf&start_date_monthly_statements=YEAR-MONTH-01&attachmentName=Monthly_Statement_YEAR_MONTH.pdf';
        const yearStr = year + '';
        const monthStr = month < 10 ? '0' + month : month + '';
        return Request.get(urlTemplate.replace(/YEAR/g, yearStr).replace(/MONTH/g, monthStr));
    }).then(rsp => {
        Logger.info('LC pre-pdf rsp:', rsp);
        const url = 'https://www.lendingclub.com/account/downloadMonthlyStatement.action?'+JSON.parse(rsp).queryString;
        Logger.info('LC pdf url:', url);
        return Request.get(url, {
            encoding: null
        });
    }).then(pdfBuffer => {
        Logger.info('Received pdf');
        return Util2.pdf2text(pdfBuffer);
    }).then(pages => {
        const accountTotalIndex = pages[0].indexOf('ACCOUNT TOTAL') + 5;
        const accountTotal = pages[0][accountTotalIndex].replace('$', '');
        const depositIndex = pages[1].indexOf('Deposits') + 6;
        const deposits = pages[1][depositIndex].replace('$', '').replace('-', '0.00');

        Logger.info('Account total:', accountTotal);
        Logger.info('Deposits:', deposits);
        return {
            deposits: deposits,
            balance: accountTotal
        };
    });
};
/**
 * Updates your google spreadsheet w/ a new row
 * @example updateGoogleSheet('500', '5642.24') // success
 * @returns {Promise.<null>} Successfully completed on resolve
 */
const updateGoogleSheet = (deposits, balance) => {
    Logger.info('Updating google sheet with:', deposits, balance);

    return Spreadsheet.loadAsync({
        spreadsheetName: Constants.GoogleSpreadsheetName,
        worksheetName: Constants.GoogleWorksheetName,
        oauth2: {
            client_id: Constants.ClientId,
            client_secret: Constants.ClientSecret,
            refresh_token: Constants.RefreshToken
        }
    }).then(spreadsheet => {
        Promise.promisifyAll(spreadsheet);
        Logger.info('Received spreadsheet');
        return [spreadsheet.receiveAsync(), spreadsheet];
    }).spread((rspArray, spreadsheet) => {
        Logger.info('Rows:', rspArray[0]);
        Logger.info('Spreadsheet info:', rspArray[1]);
        const rows = rspArray[0];
        const info = rspArray[1];
        const newRow = SheetHelper.extrapolateRow(rows[info.lastRow-1], rows[info.lastRow], info.lastRow-1, info.lastRow);
        newRow[3] = deposits;
        newRow[6] = balance;

        const spreadsheetRow = {};

        spreadsheetRow[info.nextRow] = newRow;
        Logger.info('New row:', spreadsheetRow);
        spreadsheet.add(spreadsheetRow);
        return spreadsheet.sendAsync();
    });
};

const date = new Date();
date.setMonth(date.getMonth() - 1);
const targetMonth = date.getMonth() + 1; //add 1 to make it 1-12
const targetYear = date.getFullYear();

fetchLCStatementInfo(targetMonth, targetYear).then(rsp => {
    return updateGoogleSheet(rsp.deposits, rsp.balance);
}).then(() => {
    Logger.info('Spreadsheet updated!');
    Logger.info('Great success!');
}).error(err => {
    Logger.error('Err:', err);
});
