const {By,Key,Builder, until} = require("selenium-webdriver");
require("chromedriver");
const reader = require('xlsx');


// code to read the xlsx data

const filePath = './XXXXX.xlsx'
const workbook = reader.readFile(filePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
let data = reader.utils.sheet_to_json(worksheet);


async function emailscrape(){

      let driver = await new Builder().forBrowser("chrome").build();
      await driver.get("https://portal.austinisd.org");



      for (let i=0; i < 130; i++) {
        // load the school to search
        var searchString = data[i].school;

        // goes to the webpage to check

        await driver.wait(until.elementLocated(By.id("UniqueId"))).then(el => el.sendKeys(searchString,Key.RETURN));

        await driver.wait(until.elementLocated(By.linkText(data[i].name))).then(el => el.click());

        await driver.wait(until.elementLocated(By.xpath("//a[contains(text(),'Notice of Referral Decision')]"))).then(el => el.click());

        await driver.get("https://austin.acceliplan.com/plan/Students/Landing");

      };

      let writeToWorkBook = reader.utils.book_new();
      const writeToWorkSheet = reader.utils.json_to_sheet(data);
      let exportFileName = `results.xlsx`;
      reader.utils.book_append_sheet(writeToWorkBook, writeToWorkSheet, `response`);
      reader.writeFile(writeToWorkBook, exportFileName);

      await driver.quit();
}

emailscrape();
