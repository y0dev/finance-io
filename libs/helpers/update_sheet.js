module.exports = (sum,doc,sheet,multipleSheets,print) => {

    var sheet_title = sheet.title + doc.title.substr(doc.title.length - 4);
    sum.getSumForMonthA1(sheet_title,multipleSheets,print)
    .then(result => {
      const column = 'G';
      let row = 2;
      for (const key in result) {
        if (result.hasOwnProperty(key)) {
          const cell1 = `${column}${row}`;
          const cell2 = `${column}${row+1}`;
          sheet.getCellByA1(`${cell1}`).textFormat = sheet.getCellByA1('D3').textFormat;
          sheet.getCellByA1(`${cell2}`).numberFormat = sheet.getCellByA1('E3').numberFormat;
          const element = result[key];
          switch (key) {
            case "Total Amount Spent":
              sheet.getCellByA1(`${column}2`).value = `${key}`;
              sheet.getCellByA1(`${column}3`).value = `=SUM(${sheet.title}!TotalAmount)`;
              break;
            case "Left To Spend":
              sheet.getCellByA1(`${column}4`).value = `${key}`;
              sheet.getCellByA1(`${column}5`).value = `=(3500 -$G$3)`;//`${element}`;
              break;
            default:
              break;
          }
        }
      }
      sheet.saveUpdatedCells();
    });

    // sum.getSumForMonth(sheet_title,multipleSheets,print)
    // .then(result => {
      
    // });

    sum.getSumOfCategories(sheet_title,multipleSheets,print)
    .then(result => {
      const column = 'I';
      let row = 1;
      for (const key in result) {
        if (result.hasOwnProperty(key)) {
          const cell1 = `${column}${row}`;
          const cell2 = `${column}${row+1}`;
          sheet.getCellByA1(`${cell1}`).textFormat = sheet.getCellByA1('D3').textFormat;
          sheet.getCellByA1(`${cell2}`).numberFormat = sheet.getCellByA1('E3').numberFormat;
          const element = result[key];
          switch (key) {
            case "Books":
              sheet.getCellByA1(`${column}1`).value = `${key}`;
              sheet.getCellByA1(`${column}2`).value = `=${element}`;
              break;
            case "Car":
              sheet.getCellByA1(`${column}3`).value = `${key}`;
              sheet.getCellByA1(`${column}4`).value = `=${element}`;
              break;
            case "Cleaning Supplies":
              sheet.getCellByA1(`${column}5`).value = `${key}`;
              sheet.getCellByA1(`${column}6`).value = `=${element}`;
              break;
            case "Clothing":
              sheet.getCellByA1(`${column}7`).value = `${key}`;
              sheet.getCellByA1(`${column}8`).value = `=${element}`;
              break;
            case "Entertainment":
              sheet.getCellByA1(`${column}9`).value = `${key}`;
              sheet.getCellByA1(`${column}10`).value = `=${element}`;
              break;
            case "Fast Food/Restaurant":
              sheet.getCellByA1(`${column}11`).value = `${key}`;
              sheet.getCellByA1(`${column}12`).value = `=${element}`;
              break;
            case "Gifts":
              sheet.getCellByA1(`${column}13`).value = `${key}`;
              sheet.getCellByA1(`${column}14`).value = `=${element}`;
              break;
            case "Groceries":
              sheet.getCellByA1(`${column}15`).value = `${key}`;
              sheet.getCellByA1(`${column}16`).value = `=${element}`;
              break;
            case "Health":
              sheet.getCellByA1(`${column}17`).value = `${key}`;
              sheet.getCellByA1(`${column}18`).value = `=${element}`;
              break;
            case "Home":
              sheet.getCellByA1(`${column}19`).value = `${key}`;
              sheet.getCellByA1(`${column}20`).value = `=${element}`;
              break;
            case "Internet":
              sheet.getCellByA1(`${column}21`).value = `${key}`;
              sheet.getCellByA1(`${column}22`).value = `=${element}`;
              break;
            case "Loan":
              sheet.getCellByA1(`${column}23`).value = `${key}`;
              sheet.getCellByA1(`${column}24`).value = `=${element}`;
              break;
            case "Phone Bill":
              sheet.getCellByA1(`${column}25`).value = `${key}`;
              sheet.getCellByA1(`${column}26`).value = `=${element}`;
              break;
            case "Rent":
              sheet.getCellByA1(`${column}27`).value = `${key}`;
              sheet.getCellByA1(`${column}28`).value = `=${element}`;
              break;
            case "School":
              sheet.getCellByA1(`${column}29`).value = `${key}`;
              sheet.getCellByA1(`${column}30`).value = `=${element}`;
              break;
            case "Tithes":
              sheet.getCellByA1(`${column}31`).value = `${key}`;
              sheet.getCellByA1(`${column}32`).value = `=${element}`;
              break;
            case "Tools":
              sheet.getCellByA1(`${column}33`).value = `${key}`;
              sheet.getCellByA1(`${column}34`).value = `=${element}`;
              break;
            case "Utilities":
              sheet.getCellByA1(`${column}35`).value = `${key}`;
              sheet.getCellByA1(`${column}36`).value = `=${element}`;
              break;
            case "Vacation/Trip":
              sheet.getCellByA1(`${column}37`).value = `${key}`;
              sheet.getCellByA1(`${column}38`).value = `=${element}`;
              break;
            case "Zelle Payment":
              sheet.getCellByA1(`${column}39`).value = `${key}`;
              sheet.getCellByA1(`${column}40`).value = `=${element}`;
              break;
            default:
              break;
          }
          row += 2;
        }
      }
      sheet.saveUpdatedCells();
    });

    sum.getSumOfSubCategories(sheet_title,multipleSheets,print)
    .then(results => {
      const column = 'J';
      let row = 1;
      for (const key in results) {
        if (results.hasOwnProperty(key)) {
          const cell1 = `${column}${row}`;
          const cell2 = `${column}${row+1}`;
          sheet.getCellByA1(`${cell2}`).numberFormat = sheet.getCellByA1('E3').numberFormat;
          const element = results[key];
          // console.log(key);
          switch (key) {
            case 'Car Insurance':
              sheet.getCellByA1(`${column}1`).value = `${key}`;
              sheet.getCellByA1(`${column}2`).value = `=${element}`;
              break;
            case 'Car Maintenance':
              sheet.getCellByA1(`${column}3`).value = `${key}`;
              sheet.getCellByA1(`${column}4`).value = `=${element}`;
              break;
            case 'Fuel':
              sheet.getCellByA1(`${column}5`).value = `${key}`;
              sheet.getCellByA1(`${column}6`).value = `=${element}`;
              break;
            case 'Car Other':
              sheet.getCellByA1(`${column}7`).value = `${key}`;
              sheet.getCellByA1(`${column}8`).value = `=${element}`;
              break;
            default:
              break;
          }
          row += 2;
        }
      }
      sheet.saveUpdatedCells();
    });

    sum.getSumOfStores(sheet_title,multipleSheets,print)
    .then(result => {
      const column = 'K';
      let row = 1;
      // console.log(result);
      for (const key in result) {
        if (result.hasOwnProperty(key)) {
          const cell1 = `${column}${row}`;
          const cell2 = `${column}${row+1}`;
          sheet.getCellByA1(`${cell1}`).textFormat = sheet.getCellByA1('D3').textFormat;
          sheet.getCellByA1(`${cell2}`).numberFormat = sheet.getCellByA1('E3').numberFormat;
          const element = result[key];
          switch (key) {
            case "Apple":
              sheet.getCellByA1(`${column}1`).value = `${key}`;
              sheet.getCellByA1(`${column}2`).value = `=${element}`;
              break;
          
            default:
              break;
          }
          row += 2;
        }
      }
      sheet.saveUpdatedCells();
    });

}