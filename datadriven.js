const { Builder, By, Key } = require('selenium-webdriver');
const excel = require('exceljs');

async function writeData(file_path, sheet_name, row_num, col_num, data) {
  const workbook = new excel.Workbook();
  await workbook.xlsx.readFile(file_path);
  const sheet = workbook.getWorksheet(sheet_name);
  sheet.getCell(row_num, col_num).value = data;
  await workbook.xlsx.writeFile(file_path);
}

(async function example() {
  console.log('Sample test case started');
  const driver = await new Builder().forBrowser('chrome').build();
  await driver.get('https://www.google.com/');
  await driver.sleep(3000);
  await driver.manage().window().maximize();

  const driver_path = 'chromedriver.exe';
  const path = 'options.xlsx';

  const workbook = new excel.Workbook();
  await workbook.xlsx.readFile(path);

  const sheetNames = [];

  workbook.eachSheet((sheet) => {
    sheetNames.push(sheet.name);
  });



  for (const day of sheetNames) {
    console.log(`Processing ${day}...`);

    const sheet = workbook.getWorksheet(day);
    const rows = sheet.rowCount;

    for (let r = 3; r <= rows; r++) {
      const search_query = sheet.getCell(r, 3).value;
      const search_box = await driver.findElement(By.name('q'));
      await search_box.clear();
      await search_box.sendKeys(search_query);



      // Clear the previous data
      await search_box.sendKeys("");


      // Introduce a delay of 3 seconds
      await driver.sleep(3000);

      const options = await driver.findElements(By.xpath("//ul[@role='listbox']//li[@class='sbct']"));
      let min_length = Infinity;
      let min_length_option = null;
      let max_length = 0;
      let max_length_option = null;





      for (const option of options) {
        const option_text = await option.getText();

        console.log(option_text)
        const option_length = option_text.length;
        console.log(option_length)
        if (option_length < min_length) {
          min_length = option_length;
          min_length_option = option_text;
        }
        if (option_length > max_length) {
          max_length = option_length;
          max_length_option = option_text;
        }
      }

      console.log('Minimum length:', min_length);
      console.log('Option with minimum length:', min_length_option);

      console.log('Maximum length:', max_length);
      console.log('Option with maximum length:', max_length_option);

      await writeData(path, day, r, 4, max_length_option);
      await writeData(path, day, r, 5, min_length_option);




      await search_box.clear();
    }
  }

  await driver.quit();
})();
