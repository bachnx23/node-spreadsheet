'use trict'
const {google} = require('googleapis');
const dotenv = require('dotenv');
const sheets = google.sheets('v4');
const fs = require('fs');

dotenv.load();
var config = dotenv.config();

module.exports = class NodeSpreadsheet {
    constructor() {
        this.maxRows = 5;
    }

    setMaxRows(maxRows) {
        this.maxRows = maxRows;
    }

    getMaxRows() {
        return this.maxRows;
    }

    getSpreadSheetsValue() {
        var auth2Client = this.authorize();
        var requestOptions = {
            // The ID of the spreadsheet to retrieve data from.
            spreadsheetId: config.parsed.GOOGLE_SPREADSHEET_ID,  // TODO: Update placeholder value.
    
            range: 'Sheet1',  // TODO: Update placeholder value.
        
            valueRenderOption: 'FORMATTED_VALUE',  // TODO: Update placeholder value.
        
            dateTimeRenderOption: 'FORMATTED_STRING',  // TODO: Update placeholder value.
        
            auth: auth2Client,
          };
        return new Promise((resolve, reject) => {
            sheets.spreadsheets.values.get(requestOptions, (err, response) => {
                if ( err ) reject(err);
                resolve(response.data);
            });
        }); 
    }

    countSpreadSheedsRows(tabName) {
        var auth2Client = this.authorize();
        var requestOptions = {
            // The ID of the spreadsheet to retrieve data from.
            spreadsheetId: config.parsed.GOOGLE_SPREADSHEET_ID,  // TODO: Update placeholder value.
    
            range: tabName,  // TODO: Update placeholder value.
        
            valueRenderOption: 'FORMATTED_VALUE',  // TODO: Update placeholder value.
        
            dateTimeRenderOption: 'FORMATTED_STRING',  // TODO: Update placeholder value.
        
            auth: auth2Client,
          };
        return new Promise((resolve, reject) => {
            sheets.spreadsheets.values.get(requestOptions, (err, response) => {
                if ( err ) {
                    resolve(err);
                } else {
                    if ( response.data.values ) {
                        resolve(response.data.values.length);
                    } else {
                        resolve(0);
                    }
                    
                }
            });
        });    
    }

    updateSpreadSheets(sheetName, sheet_values) {
        var oauth2Client = this.authorize();
        sheets.spreadsheets.values.append({
            spreadsheetId: config.parsed.GOOGLE_SPREADSHEET_ID,
            range: sheetName,
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            resource: {
              values: [
                sheet_values
              ],
            },
            auth: oauth2Client
          }, (err, response) => {
            if (err) return console.error(err)
          })
    }
    
    createSpreadSheetsTab(tabName) {
        var oauth2Client = this.authorize();
        var options = {
            spreadsheetId: config.parsed.GOOGLE_SPREADSHEET_ID,
            resource: {
                requests: [
                    {
                        'addSheet': {
                            'properties': {
                                'title': tabName
                            }
                        }
                    }
                ]
            },
            auth: oauth2Client
        };
        return new Promise((resolve, reject) => {
            sheets.spreadsheets.batchUpdate(options, ( err, response) => {
                if ( err ) {
                    resolve(err);
                } else {
                    resolve(response.data.replies[0].addSheet.properties);
                }
            });
        });
    }

    authorize() {
        var oAuth2Client = new google.auth.OAuth2(
            config.parsed.GOOGLE_CLIENT_ID,
            config.parsed.GOOGLE_CLIENT_SECRET
        );
        oAuth2Client.setCredentials({
            refresh_token: config.parsed.GOOGLE_REFRESH_TOKEN
        });
        oAuth2Client.refreshAccessToken((err, tokens) => {
            if ( err ) return console.log('Error for refresh access token: %s', err);
            oAuth2Client.setCredentials({
                access_token: tokens.access_token
            });
        });
        return oAuth2Client;
    }

    getAllSheets() {
        var oauth2Client = this.authorize();
        var requestOptions = {
            spreadsheetId: config.parsed.GOOGLE_SPREADSHEET_ID,
            ranges: [],
            includeGridData: false,
            auth: oauth2Client
        };
        return new Promise((resovle) => {
            sheets.spreadsheets.get(requestOptions, (err, response) => {
                if ( err ) {
                    resolve(err);
                } else {
                    resovle(response.data.sheets);
                }
            });
        });
    }

    run(sheet_values) {
        var coreSheetName = 'Sheet';
        this.getAllSheets().then((allSheet) => {
            if ( !allSheet.code ) {
                var latestSheet = allSheet.pop();
                var tabName = latestSheet.properties.title;
                this.countSpreadSheedsRows(tabName).then((countRows) => {
                    if ( !countRows.code ) {
                        if ( countRows > this.maxRows ) {
                            var splitTabName = tabName.split(' ');
                            var incrementTitleIndex = parseInt(splitTabName[1]) + 1;
                            this.createSpreadSheetsTab(coreSheetName + ' ' + incrementTitleIndex).then((createTab) => {
                                if ( !createTab.code ) {
                                    var newTabName = createTab.title; 
                                    this.updateSpreadSheets(newTabName, sheet_values);
                                }
                            });
                        } else {
                            this.updateSpreadSheets(tabName, sheet_values);
                        }
                    } 
                })
            }
        });
    }

    saveFile(dest, file ,value, clearDest) {
        if (clearDest) {
            this.removeDir(clearDest);
        }
        if ( !fs.existsSync(dest) ) {
            try {
                fs.mkdirSync(dest);
                fs.chmodSync(dest, '0775');
            } catch (e) {
                throw e;
            }
        }
        var writeStream = fs.createWriteStream(dest + '/' + file);
        writeStream.write(value);
        writeStream.close();
     }

     removeDir(path) {
        fs.readdir(path, function(err, files) {
            if (err) {
                // console.log(err.toString());
            }
            else {
                if (files.length == 0) {
                    fs.rmdir(path, function(err) {
                        if (err) {
                            // console.log(err.toString());
                        }
                    });
                }
                else {
                    files.forEach(function(file) {
                        var filePath = path + file + "/";
                        fs.stat(filePath, function(err, stats) {
                            if (stats.isFile()) {
                                fs.unlink(filePath, function(err) {
                                    if (err) {
                                        // console.log(err.toString());
                                    }
                                });
                            }
                            if (stats.isDirectory()) {
                                removeDirForce(filePath);
                            }
                        });
                    });
                }
            }
        });
    }

}