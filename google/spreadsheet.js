function getCellValue(sheet, address) {
  if (!address) return null;
  Logger.log(JSON.stringify(address));
  const cell = sheet.getRange(address);
  const value = cell.getValue();
  Logger.log(value);
  return value;
}

function setCellValue(sheet, address, value) {
  if (!address) return null;
  const cell = sheet.getRange(address);
  cell.setValue(value);
  Logger.log(value);
  return value;
}

function numberToColumnLetter(n) {
  let letter = '';
  while (n > 0) {
    let remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// tokens are expected on the last sheet
function getOAuthToken()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const lastSheet = sheets[sheets.length - 1]; // last sheet in the tab order
  const cell = lastSheet.getRange('B1'); // in B1 we expect to see the valid token
  const value = cell.getValue();
  return value;
}

// tokens are expected on the last sheet
function getRefreshUrl()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oauthSheet = ss.getSheetByName('OAuth');
  if (oauthSheet) {
    const cell = oauthSheet.getRange('B4'); // in B4 we expect to see the valid refresh URL
    const value = cell.getValue();
    return value;
  } else {
    return null;
  }
}

function createSheetWithName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if sheet already exists
  const existing = ss.getSheetByName(sheetName);
  if (existing) {
    return ss;
  }

  // Create new sheet
  const newSheet = ss.insertSheet(sheetName);
  return newSheet;
}

function startOAuth()
{
  createSheetWithName('OAuth');

  const url = "https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/oauth-start"

  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${url}", "_blank");google.script.host.close();</script>`
  ).setWidth(100).setHeight(100);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

function refreshOAuth()
{
  const url = getRefreshUrl();

  if (url)
  {
    const html = HtmlService.createHtmlOutput(
      `<script>window.open("${url}", "_blank");google.script.host.close();</script>`
    ).setWidth(100).setHeight(100);
  
    SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
  }
}

function formatLastRow(sheet, row_style, currency_symbol)
{
    const lastRow = sheet.getLastRow(); // index of last row
    for (let i = 0; i < row_style.length; i++)
    {
      if (row_style[i] == 'text') continue;

      const columnI = numberToColumnLetter(i+1);
      const range = `${columnI}${lastRow}`;
      Logger.log(JSON.stringify(range));
      const cell = sheet.getRange(range); // Change to your target cell
      if (row_style[i] == 'currency_symbol')
      {
        cell.setNumberFormat(`${currency_symbol}#,##0.00`);
      }
      else if (row_style[i] == '0.00')
      {
        cell.setNumberFormat('0.00');
      }
    }
}

function insertJSONDataFromURL(lastName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const quoteInfo = getCellValue(sheet, 'A1');

  let idStr = '';
  let quoteId = 0;
  if (quoteInfo.toString().startsWith('https://'))
  {
    // we've got raw url and can extract quote id
    const parts = quoteInfo.toString().split('/');
    idStr = parts[parts.length - 1];
  }
  else if (quoteInfo.toString().match(/^-?\d+$/)) //valid integer (positive or negative)
  {
    idStr = quoteInfo.toString();
  }

  if (idStr.match(/^-?\d+$/)) //valid integer (positive or negative)
  {
    quoteId = idStr;
  }

  if (isNaN(quoteId) || quoteId == 0 || !quoteId)
  {
    Logger.log("A1 has no valid quote");
    return;
  }

  const token = getOAuthToken();

  Logger.log("fetching URL: ");

  const url = 'https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/quote'
  +`?id=${quoteId}&token=${token}`;
  Logger.log(url);

  const response = UrlFetchApp.fetch(url);
  const jsonData = JSON.parse(response.getContentText());
  Logger.log(JSON.stringify(jsonData));

  if (!Array.isArray(jsonData)) {
    throw new Error("Expected JSON array of objects");
  }

  // Clear old data
  sheet.clearContents();
  setCellValue(sheet, 'A1', quoteInfo);

  // Write headers
  const headers = ['id','name', 'quantity', 'list_price',	'total']; // Object.keys(jsonData[0]);
  const priceColumns = ['list_price', 'total'];
  const sumColumns = ['total'];

  // Write rows that come straight
  for (let i = 0; i < jsonData.length; i++) {
    const obj = jsonData[i];
    if (Array.isArray(obj)) {
      sheet.appendRow(obj);
    }
  }
  sheet.appendRow([' ',' ',' ',' ']); // add empty row

  // write headers
  sheet.appendRow(headers);

  let row_sums = new Array(headers.length).fill(0);
  let row_style = [];

  let currency_symbol = null;

  // Write rows
  for (let i = 0; i < jsonData.length; i++) {
    const obj = jsonData[i];
    if (Array.isArray(obj)) {
      continue;
    } // skip array rows, as they are already added before
    
    Logger.log(JSON.stringify(obj));
    let row = [];
    row_style = [];

    for (let i = 0; i < headers.length; i++)
    {
      const propName = headers[i];

      const propValue = obj[propName];
      row.push(propValue);

      // collect sum for some rows
      if (sumColumns.indexOf(propName) != -1)
      {
        row_sums[i] += propValue;
      } else {
        row_sums[i] = ''; // no sum for this row
      }

      // remember necessary formatting for some rows
      if (priceColumns.indexOf(propName) != -1)
      {
        if (!currency_symbol) {
          currency_symbol = obj['currency_symbol'];
        }
        if (currency_symbol) {
          row_style.push('currency_symbol');
        } else {
          row_style.push('0.00');
        }
      } else {
        row_style.push('text');
      }

    }
    sheet.appendRow(row);
    formatLastRow(sheet, row_style, currency_symbol);
  }

  Logger.log(JSON.stringify(row_sums));
  sheet.appendRow(row_sums);
  formatLastRow(sheet, row_style, currency_symbol);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Nordimpianti')
    .addItem('Authenticate', 'startOAuth')
    .addItem('Keep Alive', 'refreshOAuth')
    .addSeparator()
    .addItem('Load Quote #', 'insertJSONDataFromURL')
    .addToUi();
}

function onEditTrigger(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (sheet.getName() === "Sheet1" && range.getA1Notation() === "D7") {
    // const lastName = getCellValue("D7");
    // insertJSONDataFromURL(lastName);
  }
}


