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

function refreshToken()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oauthSheet = ss.getSheetByName('OAuth');
  if (!oauthSheet) {
    Logger.log("Can't get OAuth sheet");
    return;
  }

  let refreshUrl = getRefreshUrl();
  refreshUrl = refreshUrl + '&format=json';

  const response = UrlFetchApp.fetch(refreshUrl);
  const data = JSON.parse(response.getContentText());
  Logger.log(data);

  const tokenCell = oauthSheet.getRange('B1'); // in B1 we expect to see the valid token
  tokenCell.setValue(data.access_token);
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
      const columnI = numberToColumnLetter(i+1);
      const range = `${columnI}${lastRow}`;
      Logger.log(JSON.stringify(range));
      const cell = sheet.getRange(range); // Change to your target cell
      if (row_style[i] == 'text')
      {
        cell.clearFormat();
      }
      else if (row_style[i] == 'currency_symbol')
      {
        cell.setNumberFormat(`${currency_symbol}#,##0.00`);
      }
      else if (row_style[i] == '0.00')
      {
        cell.setNumberFormat('0.00');
      } 
      else
      {
        cell.clearFormat();
      }
    }
}

function loadData(quoteInfo, isRU)
{
  const token = getOAuthToken();
  if (!token)
  {
    Logger.log("Missing OAuth token");
    return;
  }

  let idStr = '';
  let numberStr = '';
  let quoteId = 0;
  let quoteNumber = 0;
  let url = '';

  if (quoteInfo.toString().startsWith('https://'))
  {
    // we've got raw url and can extract quote id
    const parts = quoteInfo.toString().split('/');
    idStr = parts[parts.length - 1];
    if (idStr.match(/^-?\d+$/)) //valid integer (positive or negative)
    {
      quoteId = idStr;
    }
    if (isNaN(quoteId) || quoteId == 0 || !quoteId)
    {
      Logger.log("A1 has no valid quote id");
      return;
    }
    url = 'https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/quote'
          +`?id=${quoteId}&token=${token}&ru=${isRU}`;
  }
  else if (quoteInfo.toString().match(/^-?\d+$/)) //valid integer (positive or negative)
  {
    numberStr = quoteInfo.toString();
    if (numberStr.match(/^-?\d+$/)) //valid integer (positive or negative)
    {
      quoteNumber = numberStr;
    }
    if (isNaN(quoteNumber) || quoteNumber == 0 || !quoteNumber)
    {
      Logger.log("A1 has no valid quote number");
      return;
    }
    url = 'https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/quote'
          +`?number=${quoteNumber}&token=${token}&ru=${isRU}`;
  } else {
      Logger.log("A1 must have valid quote number or URL on quote in CRM");
      return;
  }

  Logger.log("fetching URL: ");
  Logger.log(url);

  const response = UrlFetchApp.fetch(url);
  const jsonData = JSON.parse(response.getContentText());
  Logger.log(JSON.stringify(jsonData));

  if (!Array.isArray(jsonData)) {
    throw new Error("Expected JSON array of objects");
  }

  return jsonData;
}

function insertJSONDataFromURL(isRU) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const quoteInfo = getCellValue(sheet, 'A1');
  const jsonData = loadData(quoteInfo, isRU);

  // Clear old data
  sheet.clearContents();
  setCellValue(sheet, 'A1', quoteInfo);

  // Write headers
  const headers = ['id', 'product_id', 'type', 'name', 'quantity', 'list_price', 'amount', 'k', 'unit_price', 'net_total'];
  const headersLocalized = (!isRU)
    ? ['#', 'id', 'Type',    'Name',        'Quantity',  'List Price', 'Amount', 'K', 'Unit Price',      'Total']
    : ['№', 'id', 'Тип', 'Наименование', 'Количество',    'Цена',     'Сумма', 'K', 'Стоимость (ед.)', 'Стоимость'];

  const priceColumns = ['list_price', 'amount', 'unit_price', 'net_total'];
  const intColumns = ['k'];
  const quantityColumns = ['quantity'];
  const sumColumns = ['net_total'];

  // Write rows that come straight
  for (let i = 0; i < jsonData.length; i++) {
    const obj = jsonData[i];
    if (Array.isArray(obj)) {
      sheet.appendRow(obj);
    }
  }
  sheet.appendRow([' ',' ',' ',' ']); // add empty row

  // write headers
  sheet.appendRow(headersLocalized);

  let row_sums = new Array(headers.length).fill(0);
  let row_style = [];

  let currency_symbol = null;

  // Write rows
  for (let i = 0; i < jsonData.length; i++) {
    const obj = jsonData[i];
    if (Array.isArray(obj)) {
      continue;
    } // skip array rows, as they are already added before
    
    // Logger.log(JSON.stringify(obj));
    let row = [];
    row_style = [];

    for (let i = 0; i < headers.length; i++)
    {
      const propName = headers[i];
      let propValue = obj[propName];

      const isGroup = obj.type == "Group";
      {
        // collect sum for some rows
        if (sumColumns.indexOf(propName) != -1)
        {
          if (!isGroup) {
            row_sums[i] += propValue;
          }
        } else {
          row_sums[i] = ''; // no sum for this row
        }

        // modify for units of measurement
        if (quantityColumns.indexOf(propName) != -1)
        {
          propValue = obj.usage_unit ? `${propValue} ${obj.usage_unit}` : propValue;
        }

        // remember necessary formatting for some rows
        if (intColumns.indexOf(propName) != -1)
        {
          row_style.push('0');
        }
        else if (quantityColumns.indexOf(propName) != -1)
        {
          row_style.push('text');
        }
        else if (priceColumns.indexOf(propName) != -1)
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
        row.push(propValue);
      }
    }
    sheet.appendRow(row);
    formatLastRow(sheet, row_style, currency_symbol);
  }

  Logger.log(JSON.stringify(row_sums));
  sheet.appendRow(row_sums);
  formatLastRow(sheet, row_style, currency_symbol);
}

function insertJSONDataFromURL_EN() {
  insertJSONDataFromURL(false);
}

function insertJSONDataFromURL_RU() {
  insertJSONDataFromURL(true);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Nordimpianti')
    .addItem('Load Quote # (EN)', 'insertJSONDataFromURL_EN')
    .addItem('Load Quote # (RU)', 'insertJSONDataFromURL_RU')
    .addSeparator()
    .addItem('Authenticate', 'startOAuth')
    .addItem('Keep Alive', 'refreshToken')
    .addToUi();

  refreshToken();
}

function onHalfHourTrigger()
{
  refreshToken();
}

function onEditTrigger(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (sheet.getName() === "Sheet1" && range.getA1Notation() === "D7") {
    // const lastName = getCellValue("D7");
    // insertJSONDataFromURL(lastName);
  }
}


