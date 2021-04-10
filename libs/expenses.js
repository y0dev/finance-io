const fs = require('fs').promises;
const colors = require('colors');


class Expenses {

    constructor(document,title,log = true) {
        this.doc = document;
        this.title = title;
        this.log = log;
        this.yearlyExpenses = [];
    }

    async getSheets(multipleSheets = false,print=false) {
        const sheets = [];
        for (var i = 0; i < this.doc.sheetCount; i++) {
            let sheet = this.doc.sheetsByIndex[i];
            if (multipleSheets) {
                if (sheet.title != "Settings" && sheet.title != "Loan") {
                    await this.getRows(sheet,multipleSheets,print);
                    sheets.push(sheet);
                }
            }
            else {
                if (sheet.title == this.title  || sheet.title.includes(this.title)) { 
                    await this.getRows(sheet,multipleSheets,print);
                    sheets.push(sheet);
                }
            }
        }
        return sheets;
      }

    async getRows(sheet,multipleSheets = false,print=false) {
        // console.log('Title: '+ sheet.title);
        // console.log('Rows: ' + sheet.rowCount + ', Column: ' + sheet.columnCount);
        await sheet.loadCells();
        var expenses = [];
        var monthlyTotal = [];
        var sheet_date = ""
        for (var row = 0; row < sheet.rowCount; row++) {
            var cell = sheet.getCell(row,0);
            if (cell.value != null && cell.value != 'Total') {
                // if (j == 1) { var title = cell.value; }
                if (row > 1) {
                    var expense = {
                        date: sheet.getCell(row,0).formattedValue,
                        description: sheet.getCell(row,1).value,
                        store: sheet.getCell(row,2).value,
                        category: sheet.getCell(row,3).value,
                        amount: sheet.getCell(row,4).value,
                        row: `$E$${cell._row+1}`
                    };
                    
                    expenses.push(expense);
                } 
            }
            if (cell.value == 'Total') {
                var total = sheet.getCell(row,4).value;
            }
        }

        sheet_date = new Date(sheet.getCell(2,0).formattedValue);
        this.budgetExpenses = {
            title: sheet.title + sheet_date.getFullYear() +'_Expenses',
            spent: total,
            monthlyExpenses: expenses
        };
        
        if (multipleSheets) { 
            const year = this.doc.title.substr(this.doc.title.length - 4);
            this.yearlyExpenses.push(this.budgetExpenses);
            this.logExpenses(this.yearlyExpenses,year,print);
        }
        else {
            if (this.log) { this.logExpensesForMonth(this.budgetExpenses,print); }
        }
        await sheet.addNamedRange(`${sheet.title}!Sample`,`${sheet.title}!A1:A4`);
    }

    async logExpensesForMonth(expenses,print=false) {
        let data = JSON.stringify(expenses, null, 2);
        try {
            await fs.writeFile(`./json/expense_json/${expenses.title}.json`, data);
            if (print) { console.log(`Data written to ${expenses.title}.json`); }
            
            
        } catch (error) {
            console.log(error)
        }
    }

    async logExpenses(expenses,year,print=false) {
        let data = JSON.stringify(expenses, null, 2);
        // console.log(`Data: ${data}`);
        try {
            await fs.writeFile(`./json/expense_json/expenses${year}.json`, data);
            if (print) { console.log(`Data written to expenses${year}.json`); }
        } catch (error) {
            console.log(error)
        }   
    }
}

module.exports = Expenses;

    