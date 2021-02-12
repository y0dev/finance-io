const { promisify } = require('util');
const fs = require('fs').promises;
const colors = require('colors');

class Settings {
    constructor(document,log = true) {
        this.doc = document;
        this.log = log;
    }

    getSheets() {
        for (var i = 0; i < this.doc.sheetCount; i++) {
          let sheet = this.doc.sheetsByIndex[i];
          if (sheet.title.includes('Setting')) {
            this.getRows(sheet);
          }
        }
      }

    async getRows(sheet) {
    // console.log('Title: '+ sheet.title);
    // console.log('Rows: ' + sheet.rowCount + ', Column: ' + sheet.columnCount);
    await sheet.loadCells();
    var budgetSet = [];
    for (var i = 0; i < sheet.columnCount; i++) {
        var settings = [];
        for (var j = 0; j < sheet.rowCount; j++) {
            var cell = sheet.getCell(j,i);
            if (cell.value != null) {
                if (j == 0) { var title = cell.value; }
                else {
                    settings.push(cell.value);
                }
                // console.log(cell.formattedValue);
                // console.log(cell.value);
            }
        }
        let budgetSetting = {
            title: title,
            settings: settings
        };
        budgetSet.push(budgetSetting);
    }
        this.setting = {
            settings: budgetSet
        }
        // console.log(this.setting);
        if (this.log) { this.logSettings(this.setting); }
    }

    logSettings(setting) {
        try {
            let data = JSON.stringify(setting, null, 2);
            fs.writeFile('./json/other_json/settings.json', data);
            console.log(`Data written to settings.json`.underline);
        } catch (error) {
            console.log(error)
        } 
    }
}

module.exports = Settings;