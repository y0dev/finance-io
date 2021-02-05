const fs = require('fs').promises;

class Loan {

    constructor(document) {
        this.doc = document;
        this.loans = [];
    }

    async getSheets(print=false) {
        for (var i = 0; i < this.doc.sheetCount; i++) {
            let sheet = this.doc.sheetsByIndex[i];
            await this.getRows(sheet,print);
            return sheet;
        }
      }

    async getRows(sheet,print=false) {
        // console.log('Title: '+ sheet.title);
        // console.log('Rows: ' + sheet.rowCount + ', Column: ' + sheet.columnCount);
        await sheet.loadCells();
        var loanPayments = [];
        var loans = {};
        for (var j = 0; j < sheet.rowCount; j++) {
            var cell = sheet.getCell(j,0);
            if (cell.value != null) {
                if (j == 1) { 
                    loans.loan_title = sheet.getCell(j,0).value
                    loans.loan_remaining = sheet.getCell(j,1).value.toFixed(2)
                    loans.loan_payment = sheet.getCell(j,2).value.toFixed(2)
                    loans.total_payments_left = sheet.getCell(j,3).value.toFixed(1)
                 }
                if (j > 4) {
                    var payment = {
                        date: sheet.getCell(j,0).formattedValue,
                        amount: sheet.getCell(j,1).value,
                        row: cell._row+1
                    };
                    loanPayments.push(payment);
                }
            }else {
                cell.value = 9;
            }
        }

        let loan = {
            title: sheet.title,
            original_payment: 12614.16,
            payments: loanPayments,
            loans: loans
        };

        if (sheet.title === 'Student Loan') { loan.original_payment = 12614.16; }
        if (sheet.title === 'Car Loan') 
        { 
            loan.original_payment = 28578.09;
            loan.apr = 3.34;
            loan.finance_charge = 3041.43;
        }
        this.loans.push(loan);
        await this.logSettings(loan);
        // console.log(this.budgetExpenses);
    }

    async logSettings(expenses,print=false) {
        let data = JSON.stringify(expenses, null, 2);
        // console.log(`Data: ${data}`);
        try {
            await fs.writeFile('./json/loan.json', data);
            if (print) { console.log(`Data written to loan.json`.underline); }
        } catch (error) {
            console.log(error)
        } 
    }
}

module.exports = Loan;

    