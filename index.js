const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./creds.json');
const fs = require('fs');
const Settings = require('./libs/settings.js');
const Expenses = require('./libs/expenses.js');
const Loan = require('./libs/loan.js');
const sum = require('./libs/sums');
const sheet_info = require('./json/sheet_info.json');

const update_sheet = require('./libs/helpers/update_sheet.js');
let _sheet = sheet_info[2021]
// Month - 1 to get index of month needed (ex: January = 1; index = 0)
month = _sheet.months[2-1];


const sheetTitle = month;
const print = false;

console.log(sheetTitle,_sheet.key);
async function accessSpreadsheet(sheetTitle,update_loan=false,print=false) {
  const doc = new GoogleSpreadsheet(_sheet.key);
  // console.log(doc);
  await doc.useServiceAccountAuth({
    client_email: creds.client_email,
    private_key:  creds.private_key,
  });

  await doc.loadInfo();
  // console.log('Spreadsheet Title: ' + doc.title);
  let setting = new Settings(doc);
  // await setting.getSheets();

  if (update_loan)
  {
    console.log('Update Loan');
    const loanSpreadsheet = new GoogleSpreadsheet(creds.sheet_names[2021]);
    await loanSpreadsheet.useServiceAccountAuth({
      client_email: creds.client_email,
      private_key:  creds.private_key,
    });
    await loanSpreadsheet.loadInfo();

    let loan = new Loan(loanSpreadsheet);
    const loanSheet = await loan.getSheets(true);
    const [update,payment] = await sum.updateLoanSum(print);
    
    // Update Loan Sheet if need be
    if (update) {
      await loanSheet.loadCells();
      const loanD5 = loanSheet.getCellByA1('D5'); 
      loanD5.value = `=${payment}`;
      await loanSheet.saveUpdatedCells(); // save all updates in one call
    }
  }

  

  
  console.log(doc.title);
  let exp = new Expenses(doc,sheetTitle);
  const sheets = await exp.getSheets(print);
  const multipleSheets = sheets.length > 1;

  if (multipleSheets) {
    sheets.forEach(async sheet => {
    await sheet.loadCells();
    update_sheet(sum,doc,sheet,multipleSheets,print);
    });
  } else {
    const sheet = sheets[0];
    await sheet.loadCells();

    update_sheet(sum,doc,sheet,multipleSheets,print);

  }
}

//------------MAIN------------//
accessSpreadsheet(sheetTitle,true,print);
