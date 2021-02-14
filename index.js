const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./json/hidden_json/creds.json');
const fs = require('fs');
const Settings = require('./libs/settings.js');
const Expenses = require('./libs/expenses.js');
const Loan = require('./libs/loan.js');
const sum = require('./libs/sums');
const sheet_info = require('./json/hidden_json/sheet_info.json');

const update_sheet = require('./libs/helpers/update_sheet.js');


async function accessSpreadsheet(sheetTitle,update_loan=false,multipleSheets = false,print=false) { 
  const doc = new GoogleSpreadsheet(creds.sheet_names[2021]);
  // console.log(doc);
  await doc.useServiceAccountAuth({
    client_email: creds.client_email,
    private_key:  creds.private_key,
  });

  await doc.loadInfo();
  console.log(doc.title);
  if (!multipleSheets) { console.log(sheetTitle,doc.spreadsheetId); }
  else { console.log(`Updating budget for multiple sheets in ${doc.title.substr(doc.title.length - 4)} using this key: ${doc.spreadsheetId}`);}
  // console.log('Spreadsheet Title: ' + doc.title);
  let setting = new Settings(doc);
  // await setting.getSheets();
  if (update_loan)
  {
    console.log('Update Loan');
    const loanSpreadsheet = new GoogleSpreadsheet(doc.spreadsheetId);
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

  let exp = new Expenses(doc,sheetTitle);
  const sheets = await exp.getSheets(multipleSheets,print);
  
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

// Month - 1 to get index of month needed (ex: January = 1; index = 0)
let chosen_month;
const print = false;
let multipleSheets = true;
let foundMonth = false;
const updateLoan = false;


const args = process.argv.slice(2);
if (args.length > 0) {
  creds.months.forEach(month => {
    /* 
     *  How to run program
     * 
     *  Run for every sheet in excel
     *  node index.js 
     * 
     *  Run for a specific sheet
     *  node index.js January
     *  or
     *  node index.js Jan
     * 
     *  Ex: args[0] = January
     *  or
     *  Ex: args[0].length > 2 = Jan
     * 
     */ 
    if ((args[0] == month) || (args[0].length > 2 && month.includes(args[0]))) { 
      multipleSheets = false;
      foundMonth = true;
      chosen_month = month;
    }
  });
  if(!foundMonth && multipleSheets) {
    console.log(`ERROR: ${args[0]} is not a month or sheet does not exist`);
    console.log("GOOD BYE!");
    process.exit();
  }
}


//------------MAIN------------//
accessSpreadsheet(chosen_month,updateLoan,multipleSheets,print);
