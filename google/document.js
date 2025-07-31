function extractDateFromIso(isoString) {
  const datePart = isoString.split('T')[0];  // "2025-07-31"
  Logger.log(datePart);
  return datePart;
}

function replaceInAllHeadersAndFooters(placeholder, replacement) {
  const doc = DocumentApp.getActiveDocument();
  const numSections = doc.getNumChildren();  // Each section is a Body

  for (let i = 0; i < numSections; i++) {
    const section = doc.getChild(i);
    if (section.getType() === DocumentApp.ElementType.BODY_SECTION) {
      const sectionHeader = doc.getHeader(i);
      const sectionFooter = doc.getFooter(i);

      if (sectionHeader) {
        replaceInElement(sectionHeader, placeholder, replacement);
      }
      if (sectionFooter) {
        replaceInElement(sectionFooter, placeholder, replacement);
      }
    }
  }
}

function replaceInElement(container, searchText, replaceText) {
  const numChildren = container.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const element = container.getChild(i);

    if (element.getType() === DocumentApp.ElementType.PARAGRAPH ||
        element.getType() === DocumentApp.ElementType.TABLE_CELL) {

      const text = element.asText();
      const t = text.getText();
      const index = t.indexOf(searchText);
      if (index !== -1) {
        Logger.log(`Replacing in ${t}: ${searchText} by ${replaceText}`);
        text.deleteText(index, index + searchText.length - 1);
        text.insertText(index, replaceText);
      }
    } else if (element.getType() === DocumentApp.ElementType.TABLE) {
      const table = element.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          replaceInElement(row.getCell(c), searchText, replaceText);
        }
      }
    }
  }
}

function replaceAll(searchText, replaceText) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  for (const paragraph of paragraphs) {
    const t = paragraph.getText();
    if (t.includes(searchText)) {
      Logger.log(`Replacing in ${t}: ${searchText} by ${replaceText}`);
      paragraph.replaceText(searchText, replaceText);
    }
  }

  replaceInAllHeadersAndFooters(searchText, replaceText);
}

function replaceParagraph(targetText, replaceText, canRemove) {
  const body = DocumentApp.getActiveDocument().getBody();
  const numChildren = body.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);

    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      const text = paragraph.getText();

      if (text === targetText) {
        let newParagraph = null;
        if (replaceText)
        {
          newParagraph = body.insertParagraph(i, replaceText);
        }

        // Remove the original paragraph
        if (canRemove) {
          body.removeChild(paragraph);
        }

        Logger.log('Paragraph replaced');
        return newParagraph;
      }
    }
  }

  Logger.log('No matching paragraph found: ' + targetText);
  return null;
}

function replaceParagraphWithTable(targetText, tableData, canRemove) {
  const body = DocumentApp.getActiveDocument().getBody();
  const numChildren = body.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);

    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      const text = paragraph.getText();

      if (text === targetText) {
        // Insert an empty table at the position of the paragraph
        const newTable = body.insertTable(i);

        // Fill in table rows and cells
        for (let row = 0; row < tableData.length; row++) {
          const rowData = tableData[row];
          const tableRow = newTable.appendTableRow();
          for (let col = 0; col < rowData.length; col++) {
            tableRow.appendTableCell(rowData[col]);
          }
        }

        // Remove the original paragraph
        if (canRemove) {
          body.removeChild(paragraph);
        }

        Logger.log('Paragraph replaced with table.');
        return newTable;
      }
    }
  }

  Logger.log('No matching paragraph found: ' + targetText);
  return null;
}

function insertImageFromUrl(fileUrl) {
   const response = UrlFetchApp.fetch(fileUrl);
  const blob = response.getBlob();

  Logger.log('MIME type: ' + blob.getContentType());
  Logger.log('Blob size: ' + blob.getBytes().length);

  // Example: Insert into Google Doc (if it's an image)
  const mimeType = blob.getContentType();
  if (mimeType.startsWith('image/')) {
    return blob;
  }
}

function setMetadata(obj)
{
  const body = DocumentApp.getActiveDocument().getBody();
  const lastIndex = body.getNumChildren() - 1;
  const lastElement = body.getChild(lastIndex);

  if (lastElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
    if (lastElement.getText().trim().startsWith('meta:')) {
      if (obj)
      {
        lastElement.setText("meta:" + JSON.stringify(obj));
        lastElement.setForegroundColor("#C0C0C0"); // white font, effectively hidden
        lastElement.setFontSize(10);
      }
      else
      {
        lastElement.clear(); // Clears text without removing the element
      }
    }
    else
    {
      if (obj)
      {
        const text = body.appendParagraph("meta:" + JSON.stringify(obj));
        text.setForegroundColor("#C0C0C0"); // white font, effectively hidden
        text.setFontSize(10);
      }
    }
  }
}

function getMetadata() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();

  let result = null;

  for (let p of paragraphs) {
    const text = p.getText().trim();
    if (text.startsWith('meta:')) {
      const jsonString = text.substring(5); // remove 'meta:'
      try {
        const metaText = jsonString.replace(/\\"/g, '"');
        const metadata = JSON.parse(metaText);
        // Logger.log('Found metadata: ' + JSON.stringify(metadata));
        // Now you can use metadata.key, metadata.status, etc.
        result = metadata;
      } catch (e) {
        Logger.log('Invalid JSON in metadata: ' + jsonString);
      }
      break;
    }
  }

  return result;
}

// tokens are expected on the last paragraph
function getOAuthToken()
{
  const metadata = getMetadata();
  if (metadata && metadata.oauth)
  {
    return metadata.oauth["access_token"];
  }
  else
  {
    return null;
  }
}

// tokens are expected on the last sheet
function getRefreshUrl()
{
  const metadata = getMetadata();
  if (metadata && metadata.oauth)
  {
    return metadata.oauth["refresh-url"];
  }
  else
  {
    return null;
  }
}

function tryQuickRefresh()
{
  const quickRefreshUrl = `https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/oauth-token`;
  Logger.log('Trying quick refresh:');
  const response = UrlFetchApp.fetch(quickRefreshUrl);
  const data = JSON.parse(response.getContentText());
  Logger.log(data);
  if (data && data.access_token)
  {
    let metadata = getMetadata()|| { oauth: {} };
    metadata.oauth = data;
    metadata.oauth["refresh-url"] = data.refresh_url;
    setMetadata(metadata);
    return true;
  }
  else
  {
    return false;
  }
}

function refreshToken()
{
  // try to do quick refresh with hardcoded refresh-url on DigitalOcean function
  if (tryQuickRefresh())
  {
    return;
  }

  let refreshUrl = getRefreshUrl();
  if (!refreshUrl.endsWith('format=json'))
  {
    refreshUrl = refreshUrl + '&format=json';
  }
  Logger.log(refreshUrl);

  const response = UrlFetchApp.fetch(refreshUrl);
  const data = JSON.parse(response.getContentText());
  Logger.log(data);

  let metadata = getMetadata()|| { oauth: {} };
  metadata.oauth = data;
  metadata.oauth["refresh-url"] = refreshUrl;
  setMetadata(metadata);
}

function startOAuth()
{
  // try to do quick refresh with hardcoded refresh-url on DigitalOcean function
  if (tryQuickRefresh())
  {
    return;
  }

  const metadata = getMetadata()|| { oauth: {} };
  delete metadata.oauth;
  setMetadata(metadata);

  const url = "https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/oauth-start?format=json"

  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${url}", "_blank");google.script.host.close();</script>`
  ).setWidth(100).setHeight(100);
  
  DocumentApp.getUi().showModalDialog(html, 'Opening...');
}

function refreshOAuth()
{
  // try to do quick refresh with hardcoded refresh-url on DigitalOcean function
  if (tryQuickRefresh())
  {
    return;
  }

  const url = getRefreshUrl();
  if (url)
  {
    const html = HtmlService.createHtmlOutput(
      `<script>window.open("${url}", "_blank");google.script.host.close();</script>`
    ).setWidth(100).setHeight(100);
  
    DocumentApp.getUi().showModalDialog(html, 'Opening...');
  }
}

function loadQuoteData(quoteInfo, isRU, isRaw)
{
  const token = getOAuthToken();
  if (!token)
  {
    Logger.log("Missing OAuth token");
    DocumentApp.getUi().alert('Error: Something went wrong, check logs');
    return;
  }

  let numberStr = '';
  let quoteNumber = 0;
  let url = '';

  if (quoteInfo.toString().match(/^-?\d+$/)) //valid integer (positive or negative)
  {
    numberStr = quoteInfo.toString();
    if (numberStr.match(/^-?\d+$/)) //valid integer (positive or negative)
    {
      quoteNumber = numberStr;
    }
    if (isNaN(quoteNumber) || quoteNumber == 0 || !quoteNumber)
    {
      Logger.log("Must specify valid quote number");
      return;
    }
    url = 'https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/quote'
          +`?number=${quoteNumber}&token=${token}&ru=${isRU}`;
    if (isRaw) {
      url += `&raw=1`;
    }
  } else {
      Logger.log("Must specify valid quote number");
      return;
  }

  Logger.log("Loading quote");
  Logger.log(url);

  const response = UrlFetchApp.fetch(url);
  const statusCode = response.getResponseCode();
  if (statusCode >= 400)
  {
    Logger.log('Failed to load quote data');
    return null;
  }

  if (isRaw)
  {
    const jsonData = JSON.parse(response.getContentText());

    if (!Array.isArray(jsonData.data)) {
      throw new Error("Expected JSON array of objects");
    }

    return jsonData.data[0];
  }
  else
  {
    const jsonData = JSON.parse(response.getContentText());

    if (!Array.isArray(jsonData)) {
      throw new Error("Expected JSON array of objects");
    }

    return jsonData;
  }
}

function loadProductData(productId, isRU, tokenOverride)
{
  const token = tokenOverride || getOAuthToken();
  if (!token)
  {
    Logger.log("Missing OAuth token");
    DocumentApp.getUi().alert('Error: Something went wrong, check logs');
    return;
  }

  const productUrl = `https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/product?ru=${isRU}&id=${productId}&token=${token}`;
  Logger.log(productUrl);

  const response = UrlFetchApp.fetch(productUrl);
  const statusCode = response.getResponseCode();
  if (statusCode >= 400)
  {
    Logger.log('Failed to load product data');
    DocumentApp.getUi().alert('Error: Something went wrong, check logs');
    return null;
  }

  const data = JSON.parse(response.getContentText());
  return data;
}

function askQuoteNumber()
{
  const ui = DocumentApp.getUi();
  const response = ui.prompt('Enter Quote Number', 'Please enter the quote number:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() != ui.Button.OK) {
    return 0;
  }

  const quoteNumber = response.getResponseText();
  return quoteNumber;
}

function addAttachment(table, productData, attachmentData, token)
{
  const imageUrl = `https://faas-fra1-afec6ce7.doserverless.co/api/v1/web/fn-cc862375-8740-4e17-bd3e-7bfe3aaf577b/functions/attachment?module_name=Products&parent_id=${productData.id}&attachment_id=${attachmentData.attachment_Id}&file_name=${attachmentData.file_Name}&token=${token}`;
  Logger.log("Loading attachment image");
  Logger.log(imageUrl);

  const imageBlob = insertImageFromUrl(imageUrl);
  if (imageBlob) {
    const row = table.appendTableRow();
    const cell = row.appendTableCell();
    const insertedImage = cell.appendImage(imageBlob);
    if (insertedImage)
    {
      // Adjust the image width so it fits within the page (assuming 8.5 inches page width)
      // Convert inches to pixels (Google Docs usually uses 72 DPI for page size)
      const pageWidthInPixels = 8.5 * 72;  // 8.5 inches * 72 pixels/inch (standard page width in pixels)
  
      // set the image's width (keep the aspect ratio intact)
      const originalWidth =  insertedImage.getWidth();
      const originalHeight = insertedImage.getHeight();

      const desiredWidth = pageWidthInPixels * 1.0;
      const scaleFactor = desiredWidth / originalWidth;
      if (scaleFactor < 1.0) // let only downscale images
      {
        insertedImage.setWidth(originalWidth * scaleFactor);
        insertedImage.setHeight(originalHeight * scaleFactor);
      }
    }
  }
}

function startLoadQuote(quoteNumber, isRU)
{
  Logger.log("Loading raw Quote: " + quoteNumber);
  const jsonData = loadQuoteData(quoteNumber, isRU, false);
  if (!jsonData)
  {
    DocumentApp.getUi().alert('Error: Something went wrong, check logs');
    return;
  }

  Logger.log("Loading raw Quote: " + quoteNumber);
  const rawQuote = loadQuoteData(quoteNumber, isRU, true);
  if (!rawQuote)
  {
    DocumentApp.getUi().alert('Error: Something went wrong, check logs');
    return;
  }

  replaceAll('{{Quote.Quote_Number}}', rawQuote.Quote_Number);
  replaceAll('{{Quote.Account_Name.name}}', rawQuote.Account_Name.name);
  replaceAll('{{Quote.Billing_City}}', rawQuote.Billing_City);
  replaceAll('{{Quote.Billing_Country}}', rawQuote.Billing_Country);
  replaceAll('{{Quote.Created_Time}}', extractDateFromIso(rawQuote.Created_Time));
  replaceAll('{{Quote.Production_Time}}', rawQuote.Production_Time);
  replaceAll('{{Quote.Carrier}}', rawQuote.Carrier);
  replaceAll('{{Quote.Delivery_place}}', rawQuote.Delivery_place);
  replaceAll('{{Quote.Valid_Till}}', rawQuote.Valid_Till);

  return;

  const token = getOAuthToken();

  const metadata = getMetadata();
  setMetadata(null); // remove it for now

  const tableData = [];
  const tableStyle = [];

  const tableOfContents = [
    // here we will collect '1.1' => {product info}, so that chapters will be generated same as they are listed in the table
  ];

  // this is needed to swap descriptions and table at the end
  // we generate table, collect chapter map like 1.1, 1.2, and groups too
  // then descriptions are generated to correspond chapter numeration.
  // But we still need table of contents to be at the end of the document.
  let elementTableOfContents = null;

  try
  {
    // Write headers
    const headers = ['id','name', 'quantity', 'net_total'];
    const headersLocalized = (!isRU)
      ? ['#','Name', 'Quantity', 'Total']
      : ['№', 'Наименование', 'Кол-во', 'Стоимость'];

    const columnWidths = [40, 320, 60, 80];

    const priceColumns = ['net_total'];
    const intColumns = [];
    const quantityColumns = ['quantity'];
    const sumColumns = ['net_total'];
    const alignRight = ['net_total'];
    const alignCenter = ['id'];

    // write headers
    tableData.push(headersLocalized);
    tableStyle.push(new Array(headers.length).fill('text,bold,bg=#e6e6e6'));

    let row_sums = new Array(headers.length).fill(0);
    let row_style = [];

    let currency_symbol = null;

    let numberGroup = 0;
    let numberProduct = 0;
    let lastSeenGroup = null;

    // Write rows
    for (let i = 0; i < jsonData.length; i++)
    {
      const obj = jsonData[i];
      Logger.log(obj);
      if (Array.isArray(obj)) {
        continue;
      } // skip array rows, as they are already added before
      
      let row = new Array(headers.length).fill('');
      row_style = new Array(headers.length).fill('');

      const isGroup = obj.type == "Group";
      if (isGroup) {
        if (lastSeenGroup)
        {
          row       = new Array(headers.length).fill('');
          row_style = new Array(headers.length).fill('text,bold');

          row[headers.indexOf('id')] = `${numberGroup}.0`;
          row[headers.indexOf('name')] = isRU ? "ИТОГО" : "TOTAL";
          row[headers.indexOf('net_total')] = 
            (new Intl.NumberFormat('en-US', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2
            }).format(lastSeenGroup.net_total)) + ' ' + currency_symbol;

          tableData.push(row);
          tableStyle.push(row_style);
        }
      
        numberGroup += 1;
        numberProduct = 0;

        row       = new Array(headers.length).fill('');
        row_style = new Array(headers.length).fill('text,bold');

        const idStr = `${numberGroup}.0`;
        row[headers.indexOf('id')] = idStr;
        row[headers.indexOf('name')] = obj.name;

        tableData.push(row);
        tableStyle.push(row_style);

        lastSeenGroup = obj;

        tableOfContents.push({
          index: i, // get object by jsonData[i]
          chapter: idStr,
          product_id: null,
          is_group: true,
          name: obj.name
        });

        continue;
      } else {
        numberProduct += 1;

        const idStr = `${numberGroup}.${numberProduct}`;
        row[headers.indexOf('id')] = idStr;
        row[headers.indexOf('name')] = obj.name;

        tableOfContents.push({
          index: i, // get object by jsonData[i]
          chapter: idStr,
          product_id: obj.product_id,
          is_group: false,
          name: obj.name
        });
      }

      for (let i = 0; i < headers.length; i++)
      {
        const propName = headers[i];
        if (propName == 'id' || propName == 'name') {
          continue;
        }

        let propValue = obj[propName];

        // modify for units of measurement
        if (quantityColumns.indexOf(propName) != -1)
        {
          propValue = obj.usage_unit ? `${propValue} ${obj.usage_unit}` : propValue;
        }

        // collect sum for some rows
        if (sumColumns.indexOf(propName) != -1)
        {
          row_sums[i] += propValue;
        } else {
          row_sums[i] = ''; // no sum for this row
        }

        // remember necessary formatting for some rows
        if (intColumns.indexOf(propName) != -1)
        {
          // format as integer
          propValue = new Intl.NumberFormat('en-US', {
            minimumFractionDigits: 0,
            maximumFractionDigits: 0
          }).format(propValue);

          row_style[i] = '0';
        }
        else if (priceColumns.indexOf(propName) != -1)
        {
          if (!currency_symbol) {
            currency_symbol = obj['currency_symbol'];
          }
          if (currency_symbol) {
            // format as 382,386.00 €
            propValue = (new Intl.NumberFormat('en-US', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2
            }).format(propValue)) + ' ' + currency_symbol;

            row_style[i] = 'currency_symbol';
          } else {
            // format as 382,386.00
            propValue = new Intl.NumberFormat('en-US', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2
            }).format(propValue);

            row_style[i] = '0.00';
          }
        } else {
          row_style = 'text';
        }
        row[i] = propValue;
      }

      tableData.push(row);
      tableStyle.push(row_style);
    }

    for (let i = 0; i < row_sums.length; i++)
    {
      if (row_sums[i])
      {
        // format sum into currency
        row_sums[i] = (new Intl.NumberFormat('en-US', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2
            }).format(row_sums[i])) + ' ' + currency_symbol;
      }
    }

    row_sums[0] = '';
    row_sums[1] = isRU ? 'Общая стоимость коммерческого предложения' : 'GRAND TOTAL';
    tableData.push(row_sums);
    tableStyle.push(new Array(headers.length).fill('text,bold,bg=#e2efd9'));

    //const doc = DocumentApp.getActiveDocument();
    //const body = doc.getBody();
    // const table = body.appendTable(tableData);

    const table = replaceParagraphWithTable("<PLACEHOLDER_PRICE>", tableData, true);
    if (!table) {
      DocumentApp.getUi().alert('Error: <PLACEHOLDER_PRICE> not found');
    }

    // Loop through all rows and cells
    for (let i = 0; i < table.getNumRows(); i++){
      const row = table.getRow(i);
      for (let j = 0; j < row.getNumCells(); j++) {
        const cell = row.getCell(j);
        for (let c = 0; c < cell.getNumChildren(); c++) {
          const text = cell.getChild(c).asText();

          cell.setWidth(columnWidths[j]);

          if (alignRight.indexOf(headers[j]) != -1)
          {
            var cellStyle = {};
            cellStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
            cell.getChild(0).asParagraph().setAttributes(cellStyle);
          }
          else if (alignCenter.indexOf(headers[j]) != -1)
          {
            var cellStyle = {};
            cellStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
            cell.getChild(0).asParagraph().setAttributes(cellStyle);
          }

          if (tableStyle[i][j].includes(',bold'))
          {
            text.setBold(true);
          }
          if (tableStyle[i][j].includes(',bg=#'))
          {
            const parts = tableStyle[i][j].split(',');
            const bgfound = parts.filter(function(item) {
              return item.startsWith("bg=#") && item.length == ("bg=#F0F1F2").length; // find color style
            });
            if (bgfound.length == 1)
            {
              const bgColor = bgfound[0].replace("bg=","");
              cell.setBackgroundColor(bgColor);
            }
          }

          // Set font family, size, and color
          text.setFontFamily('Arial');
          text.setFontSize(11);
          text.setForegroundColor('#000000');
        }
      }
    }

    elementTableOfContents = table;
  }
  catch (e) {
    // Catch and log the error
    Logger.log('❌ An error occurred: ' + e.message);
    Logger.log('Stack trace:\n' + e.stack);
  }

  try
  {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    for (const chapterInfo of tableOfContents)
    {
      Logger.log('Creating chapter: ');
      Logger.log(chapterInfo);

      const obj = jsonData[chapterInfo.index];
      if (Array.isArray(obj)) {
        // must never reach here!
        continue;
      }

      const tableRows = [
        [ `${chapterInfo.chapter} ${chapterInfo.name}` ],
      ];

      // const table = body.appendTable(tableRows);
      const table = replaceParagraphWithTable("<PLACEHOLDER_DESCRIPTION>", tableRows, false);
      if (!table) {
        DocumentApp.getUi().alert('Error: <PLACEHOLDER_DESCRIPTION> not found');
        break;
      }

      let productData = null;
      if (!chapterInfo.is_group)
      {
        Logger.log(`Loading product ${chapterInfo.product_id}`);
        productData = loadProductData(chapterInfo.product_id, isRU, token);
        Logger.log('Got product data');
      }
      if (productData)
      {
        Logger.log('Checking for image');

        if (productData.Product_Image_1)
        {
          const attachment = productData.Product_Image_1[0];
          if (attachment)
          {
            addAttachment(table, productData, attachment, token);
          }
        }
        if (productData.Product_Image_2)
        {
          const attachment = productData.Product_Image_2[0];
          if (attachment)
          {
            addAttachment(table, productData, attachment, token);
          }
        }
        if (productData.Product_Image_3)
        {
          const attachment = productData.Product_Image_3[0];
          if (attachment)
          {
            addAttachment(table, productData, attachment, token);
          }
        }

        const descriptionStr = isRU ? (productData['Description_RU'] || productData['Description']) : productData['Description'];
        const row = table.appendTableRow();

        const hasSpecs = productData.Attributes && Array.isArray(productData.Attributes) && productData.Attributes.length > 0;
        row.appendTableCell(`${descriptionStr}`);

        if (hasSpecs)
        {
          const row = table.appendTableRow();

          const specsExtraHeader = isRU ? "Технические характеристики:" : "Technical specifications:";
          const cell = row.appendTableCell(`${specsExtraHeader}`);
          const specsTable = cell.appendTable();

          for (let i = 0; i < productData.Attributes.length; i++)
          {
            const attr = productData.Attributes[i];

            const row = specsTable.appendTableRow();
            row.appendTableCell( isRU ? attr["Name_RU"] || attr["Name_EN"] : attr["Name_EN"]);
            row.appendTableCell( attr["Value"] );
            row.appendTableCell( attr["Unit"] );
          }
        }
      }

      table.setBorderColor('#ffffff');
      table.setBorderWidth(0);
      // Loop through all rows and cells
      for (let i = 0; i < table.getNumRows(); i++){
        const row = table.getRow(i);
        for (let j = 0; j < row.getNumCells(); j++) {
          const cell = row.getCell(j);
          if (chapterInfo.is_group && i == 0)
          {
            cell.setBackgroundColor('#b5b4b4');
          }
          for (let c = 0; c < cell.getNumChildren(); c++) {
            const text = cell.getChild(c).asText();
            text.setFontFamily('Arial');
            text.setFontSize(11);
            text.setForegroundColor('#000000');
            if (i == 0)
            {
              text.setBold(true);
            }
          }
        }
      }
    }

    replaceParagraph("<PLACEHOLDER_DESCRIPTION>", null, true);

    // swap description and table of contents
    //elementTableOfContents.removeFromParent();
    //body.appendPageBreak();
    //body.appendTable(elementTableOfContents);

  } catch (e)
  {
    // Catch and log the error
    Logger.log('❌ An error occurred: ' + e.message);
    Logger.log('Stack trace:\n' + e.stack);
  }

  setMetadata(metadata); // restore it to be the last paragraph in the doc
}

function reloadQuote()
{
  const metadata = getMetadata();
  const quoteNumber = metadata["quoteNumber"];
  const isRU = metadata["isRU"];
  if (quoteNumber && isRU !== undefined)
  {
    startLoadQuote(quoteNumber, isRU);
  }
}

function startLoadQuote_EN()
{
  const quoteNumber = askQuoteNumber();
  if (quoteNumber)
  {
    const metadata = getMetadata();
    metadata["quoteNumber"] = quoteNumber;
    metadata["isRU"] = false;
    setMetadata(metadata);
    startLoadQuote(quoteNumber, false);
  }
}

function startLoadQuote_RU()
{
  const quoteNumber = askQuoteNumber();
  if (quoteNumber)
  {
    const metadata = getMetadata();
    metadata["quoteNumber"] = quoteNumber;
    metadata["isRU"] = true;
    setMetadata(metadata);
    startLoadQuote(quoteNumber, true);
  }
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Nordimpianti')
    .addItem('Load Quote # (EN)', 'startLoadQuote_EN')
    .addItem('Load Quote # (RU)', 'startLoadQuote_RU')
    .addItem('Reload Quote', 'reloadQuote')
    .addSeparator()
    .addItem('Authenticate', 'startOAuth')
    .addToUi();

  refreshToken();
}

function onHalfHourTrigger()
{
  refreshToken();
}
