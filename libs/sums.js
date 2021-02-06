const fs = require('fs').promises;

async function getSumForMonth(sheet_title,multipleSheets=true,print=false) {
  const total = {};
  const sheetLength = sheet_title.length;
  const year = sheet_title.substr(sheet_title.length - 4);
  try {
    if (multipleSheets){
      const data = await fs.readFile(`./json/expenses${year}.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.forEach(async (month) => {
        if (month.title == `${sheet_title}_Expenses` || month.title.includes(sheet_title)) {
          total["Total Amount Spent"] = Number (month["Total Amount Spent"].toFixed(2));
        }
      })
      
    } else {
      const data = await fs.readFile(`./json/${sheet_title}_Expenses.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      // console.log(expensesJson["Total Amount Spent"]);
      total["Total Amount Spent"] = Number (expensesJson["Total Amount Spent"].toFixed(2));
    }
  } catch (error) {
    console.log(error.red);
  }
  if (print) { console.log(`Spent ${total.amount} in the month of ${sheet_title.slice(0,sheetLength-4)}`.green); }
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve(total);
    }, 1500);
  });
}

async function getStoresWith(title,sheet_title,multipleSheets=true,print=false) {
  
  var sum = 0;
  const sheetLength = sheet_title.length;
  const month = sheet_title.slice(0,sheetLength-4);
  const year = sheet_title.substr(sheet_title.length - 4);
  try {
    if (multipleSheets){
      const data = await fs.readFile(`./json/expenses${year}.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.forEach(async (month) => {
        if (month.title == `${sheet_title}_Expenses` || month.title.includes(sheet_title)) {
          // console.log(month.title);
          month.expenses_for_the_month.forEach( (expense) => {
            if (expense.store.includes(title) || expense.store == title) {
              sum += expense.amount;
            }
          });
        }
      })
      
    } else {
      const data = await fs.readFile(`./json/${sheet_title}_Expenses.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.expenses_for_the_month.forEach( (expense) => {
        if (expense.store.includes(title) || expense.store == title) {
            sum += expense.amount;
        }
      });
    }
  } catch (error) {
    console.log(error.red);
  }
  if (print) { console.log(`Spent ${sum.toFixed(2)} at ${title} in the month of ${month}`.green); }
  return Promise.resolve(sum.toFixed(2)); 
}

async function getCategoryWith(title,sheet_title,multipleSheets=true,print=false) {
  
  var sum = 0;
  const sheetLength = sheet_title.length;
  const month = sheet_title.slice(0,sheetLength-4);
  const year = sheet_title.substr(sheet_title.length - 4);
  try {
    if (multipleSheets){
      const data = await fs.readFile(`./json/expenses${year}.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.forEach(async (month) => {
        if (month.title == `${sheet_title}_Expenses` || month.title.includes(sheet_title)) {
          month.expenses_for_the_month.forEach( (expense) => {
            if (expense.category.includes(title) || expense.category == title) {
                sum += expense.amount;
            }
        });
        }
      })
      
    } else {
      const data = await fs.readFile(`./json/${sheet_title}_Expenses.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.expenses_for_the_month.forEach( (expense) => {
        if (expense.category.includes(title) || expense.category == title) {
          sum += expense.amount;
        }
      });
    }
  } catch (error) {
    console.log(error.red);
  }

  if (print) { console.log(`Spent ${sum.toFixed(2)} in buying ${title} in the month of ${month}`.cyan); }
  return Promise.resolve(sum.toFixed(2)); 
}

async function getSumOfCategoryWithStoreName(category,store,sheet_title,multipleSheets=true,print=false) {
  
  var sum = 0;
  const sheetLength = sheet_title.length;
  const month = sheet_title.slice(0,sheetLength-4);
  const year = sheet_title.substr(sheet_title.length - 4);
  try {
    if (multipleSheets){
      const data = await fs.readFile(`./json/expenses${year}.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.forEach(async (month) => {
        if (month.title == `${sheet_title}_Expenses` || month.title.includes(sheet_title)) {
          month.expenses_for_the_month.forEach( (expense) => {
            if ((expense.category.includes(category) || expense.category == category) &&
               (expense.store.includes(store) || expense.store == store)) {
                sum += expense.amount;
            }
        });
        }
      })
      
    } else {
      const data = await fs.readFile(`./json/${sheet_title}_Expenses.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.expenses_for_the_month.forEach( (expense) => {
        if ((expense.category.includes(category) || expense.category == category) &&
            (expense.store.includes(store) || expense.store == store)) {
              sum += expense.amount;
        }
      });
    }
  } catch (error) {
    console.log(error.red);
  }

  if (print) { console.log(`Spent ${sum.toFixed(2)} in buying ${category} at ${store} in the month of ${month}`.cyan); }
  return Promise.resolve(sum.toFixed(2)); 
}

async function getStoreWithCategoryWith(storeTitle,categoryTitle,sheet,multipleSheets=true,print=false) {

  var sum = 0;
  const sheetLength = sheet.title.length;
  const month = sheet.title.slice(0,sheetLength-4)
  try {
    if (multipleSheets){
      const data = await fs.readFile(`./json/expenses${year}.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.forEach(async (month) => {
        if (month.title == `${sheet.title}_Expenses` || month.title.includes(sheet.title)) {
          month.expenses_for_the_month.forEach(async (expense) => {
            if ((expense.category.includes(categoryTitle) || expense.category == categoryTitle) 
                    && (expense.store.includes(storeTitle) || expense.store == storeTitle)) {
                sum += expense.amount;
            }
        });
        }
      })
      
    } else {
      const data = await fs.readFile(`./json/${sheet.title}_Expenses.json`, 'utf8');
      var expensesJson = JSON.parse(data);
      expensesJson.expenses_for_the_month.forEach( (expense) => {
          if ((expense.category.includes(categoryTitle) || expense.category == categoryTitle) 
                  && (expense.store.includes(storeTitle) || expense.store == storeTitle)) {
              sum += expense.amount;
          }
      });
    }
  } catch (error) {
    console.log(error.red);
  }
  if (print) { console.log(`Spent ${sum.toFixed(2)} in buying ${categoryTitle} at ${storeTitle} in the month of ${month}`.yellow); }
  return Promise.resolve(sum.toFixed(2));
}


async function getSumOfStores(sheet_title,multipleSheets=true,print=false) {
  const sums = {};
  try {
    const data = await fs.readFile(`./json/settings.json`, 'utf8');
    const settingsJSON = JSON.parse(data);
    settingsJSON.settings.forEach(async (setting) => {
      if (setting.title == "Store Names") {
        setting.settings.forEach(async (storeName) => {
          sums[storeName] = await getStoresWith(storeName,sheet_title,multipleSheets,print);
        });
      }// end if
    });
  } catch (error) {
    console.log(error);
  }
  finally {
    if (print) { console.log(`Sums for the month: ${sums}`.bgBlue); }
    return new Promise((resolve) => {
      setTimeout(() => {
        resolve(sums);
      }, 500);
    });
  }
}

async function getSumOfCategories(sheet_title,multipleSheets=true,print=false) {
  const sums = {};
  try {
    const data = await fs.readFile(`./json/settings.json`, 'utf8');
    const settingsJSON = JSON.parse(data);
    settingsJSON.settings.forEach(async (setting) => {
      if (setting.title == "Expense Categories") {
        setting.settings.forEach(async (category) => {
          sums[category] = await getCategoryWith(category,sheet_title,multipleSheets,print);
        });
      }
    });
  } catch (error) {
    console.log(error)
  }
  if (print) { console.log(`Sums for the month: ${sums}`.bgMagenta); }
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve(sums);
    }, 500);
  });
}

async function getSumOfSubCategories(sheet_title,multipleSheets=true,print=false) {
  const car_cat = {
    'Car Insurance': ['Allstate','Statefarm'],
    'Car Maintenance': ['Amazon','Autozone','Toyota','Honda','Orielly\'s','PepBoys'],
    'Fuel': ['Costco','Chevron'],
    'Car Other': ['N/A','Circle Marina Speedwash']
  };
  const sums = {};
  try {
    const data = await fs.readFile(`./json/settings.json`, 'utf8');
    const settingsJSON = JSON.parse(data);
    settingsJSON.settings.forEach(async (setting) => {
      if (setting.title == "Expense Categories") {
        setting.settings.forEach(async (category) => {
          if (category === "Car") {
            for (const key in car_cat) {
              if (car_cat.hasOwnProperty(key)) {
                const element = car_cat[key];
                sums[key] = 0;
                element.forEach(async (store) => {
                  getSumOfCategoryWithStoreName(category,store,sheet_title,multipleSheets,print)
                  .then(result => sums[key] += Number(result));
                })
              }
            }// end for
          }
        });
      }
    });
  } catch (error) {
    console.log(error)
  }
  if (print) { console.log(`Sums for the month: ${sums}`.bgMagenta); }
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve(sums);
    }, 500);
  });
}

async function updateLoanSum(print=false) {
  try {
    var sum = 0;
    const data = await fs.readFile(`./json/loan.json`, 'utf8');
    const loanJson = JSON.parse(data);
    
    loanJson.payments.forEach( (loan_payment) => {
      sum += loan_payment.amount;
    });
    if (print) {
      console.log(`Sum: ${sum}`);
      console.log(`Original: ${loanJson.original_payment} - ${sum} = ${loanJson.original_payment - sum}`);
      console.log(`Loan Payment: ${loanJson.loans.loan_remaining}`);
    }
    
    if ((loanJson.original_payment - sum).toFixed(2) == loanJson.loans.loan_remaining) {
      if(print) { console.log('No loan update needed'); }
      
      return [false,sum.toFixed(2)];
    } 
    return [true,sum.toFixed(2)];
  } catch (error) {
    console.log(error);
  }
}


module.exports = {
  updateLoanSum,
  getSumOfStores,
  getSumOfCategories,
  getSumOfSubCategories,
  getSumForMonth,
};