const { Builder, By } = require('selenium-webdriver');
const XLSX = require('xlsx');

async function readExcelForToday(fileName) {
    const workbook = XLSX.readFile(fileName);
    const today = new Date();
    const sheetName = today.toLocaleString('default', { weekday: 'long' });
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    const dataDict = {};
    data.forEach(row => {
        dataDict[row['Value']] = row['Value_Content'];
    });
    return dataDict;
}

async function getGoogleSuggestions(searchQuery) {
    let driver = await new Builder().forBrowser('chrome').build();
    await driver.get('https://www.google.com');
    await driver.findElement(By.name('q')).sendKeys(searchQuery);
    await driver.sleep(500);

    const suggestions = await driver.findElement(By.id('Alh6id')).getText();
    const suggestionsList = suggestions.split('\n');

    await driver.quit();
    return suggestionsList;
}

async function appendToExcel(fileName, data) {
    const workbook = XLSX.readFile(fileName);
    const today = new Date();
    const sheetName = today.toLocaleString('default', { weekday: 'long' });

    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    console.log("Appending..");
    jsonData.forEach(row => {
        if (data[row['Value']]) {
            const [suggestions, shortest, longest] = data[row['Value']];
            row['Shortest Option'] = shortest;
            row['Longest Option'] = longest;
        }
    });

    const updatedSheet = XLSX.utils.json_to_sheet(jsonData);
    workbook.Sheets[sheetName] = updatedSheet;
    XLSX.writeFile(workbook, fileName);
}

(async () => {
    const FILE_NAME = 'Excel.xlsx';
    const searchData = await readExcelForToday(FILE_NAME);

    const resultData = {};
    for (const key in searchData) {
        const suggestions = await getGoogleSuggestions(searchData[key]);
        if (suggestions.length) {
            const shortest = suggestions.reduce((a, b) => a.length <= b.length ? a : b);
            const longest = suggestions.reduce((a, b) => a.length >= b.length ? a : b);
            resultData[key] = [suggestions, shortest, longest];
        }
    }

    await appendToExcel(FILE_NAME, resultData);
    console.log("Success");
})();
