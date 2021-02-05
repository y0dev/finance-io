const { settings } = require('cluster');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { promisify } = require('util');
const creds = require('./creds.json');
const fs = require('fs');
const Settings = require('./libs/settings.js');
const Expenses = require('./libs/expenses.js');
const Loan = require('./libs/loan.js');
const sum = require('./libs/sums');
const colors = require('colors');

const sheetTitle = 'November2020'
const print = false;


async function accessSpreadsheet() {
  const doc = new GoogleSpreadsheet('1-Fp-bAL-CtxxRjG6Zge9ASXlP2orTaRCESg3fuolqvI');
  // console.log(doc);
  await doc.useServiceAccountAuth({
    client_email: creds.client_email,
    private_key:  creds.private_key,
  });

  await doc.loadInfo();
  console.log('Spreadsheet Title: ' + doc.title);
  let setting = new Settings(doc);
  // await setting.getSheets();

  let loan = new Loan(doc);
  const loanSheet = await loan.getSheets();
  const [update,payment] = await sum.updateLoanSum(print);
  
  // Update Loan Sheet if need be
  if (update) {
    await loanSheet.loadCells();
    const loanD5 = loanSheet.getCellByA1('D5'); 
    loanD5.value = `=${payment}`;
    await loanSheet.saveUpdatedCells(); // save all updates in one call
  }

  let exp = new Expenses(doc,sheetTitle);
  const sheets = await exp.getSheets();
  const multipleSheets = sheets.length > 1;

  if (multipleSheets) {
    sheets.forEach(async sheet => {
    await sheet.loadCells();
    // const storeSums = await sum.getSums(sheet,multipleSheets,print);
    // console.log(storeSums);

      // const h39 = sheet.getCellByA1('H39'); // or A1 style notation
      // const e50 = sheet.getCellByA1('E50');
      // console.log(`${h39.value} ${e50.value}`);
      // console.log(value);
      // h39.value = `=${value}`;
      await sheet.saveUpdatedCells(); // save all updates in one call
    });
  } else {
    const sheet = sheets[0];
    await sheet.loadCells();
    const [storeSums,cateSums] = await Promise.all([sum.getSumOfStores(sheet,multipleSheets,print),sum.getSumOfCategories(sheet,multipleSheets,print)]);
    console.log(storeSums,cateSums);
    
    // sum.sum_of_category('Groceries',sheet,false);
    // sum.sum_of_category('Car',sheet,false);
    // const costcoFuelValue = await sum.sum_of_category_at('Costco','Car',sheet,false);
    // const value = await sum.sum_with_store_title('Costco',sheet,false);
    // const h39 = sheet.getCellByA1('H39'); // or A1 style notation
    // const e50 = sheet.getCellByA1('E50');
    // console.log(`${h39.value} ${e50.value}`);
    // console.log(sheet.getCellByA1('I9').numberFormat);
    // sheet.getCellByA1('I14').textFormat = sheet.getCellByA1('I9').textFormat;
    // sheet.getCellByA1('I14').numberFormat = sheet.getCellByA1('I9').numberFormat;
    // sheet.getCellByA1('I14').value = `=${costcoFuelValue}`;
    // await sheet.saveUpdatedCells(); // save all updates in one call
  }
  // let sum = await exp.getStoresWith('Costco');
  // console.log(`Spent ${sum[0]} at ${sum[1]}`.green);
  // let set = new Settings(doc);
  // set.logSettings();
  // getSheets(doc);

  // const rows
  // const info = doc.getInfo();
  // const sheet = info.worksheets[0];
  // console.log(info);
  // console.log('Title: ${sheet.title}, Rows: ${sheet.rowCount}');

  /*
  const newSheet = doc.addSheet({
    title: 'Testing',
    headerValues: ['Pet Name','Animal Type']
  });

  const row = (await newSheet).addRow({
    'Pet Name': 'Skip',
    'Animal Type': 'Dog',
  });
  */
}

function getSheets(document) {
  for (var i = 0; i < document.sheetCount; i++) {
    let sheet = document.sheetsByIndex[i];
    if (sheet.title.includes('Setting')) {
      // getRows(sheet);
    }
  }
}

async function getRows(sheet) {
  console.log('Title: '+ sheet.title);
  console.log('Rows: ' + sheet.rowCount + ', Column: ' + sheet.columnCount);
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
  let setting = {
      settings: budgetSet
  }
  let data = JSON.stringify(setting, null, 2);

  fs.appendFile('./json/settings.json', data, (err) => {
      if (err) throw err;
      console.log('Data written to file');
  });
}

//------------MAIN------------//
accessSpreadsheet();
