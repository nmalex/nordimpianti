const fetch = require('node-fetch');

// shared config hardcoded for visibility
const config =
{
  "crm_api_url": "https://crm.zoho.eu/crm/v4"
};

const _local = {};
_local['ru'] = {
  'Quote Id': "Идентификатор предложения",
  'Quote Number': "Номер предложения",
  'Account Name': "Аккаунт",
  'Subject': "Тема",
  'Agent': "Агент",
  'Created By': "Создал",
  'Contact Name': "Контакт",
  'Deal Name': "Сделка",
  'Created Time': "Дата создания",
  'Production Time': "Время производства",
};

exports.main = async function (args) {
  try {
    const accessToken = args['token'] || args.query && args.query['token'];
    if (!accessToken) {
      return { statusCode: 400, body: 'Missing token' };
    }

    let isRaw = args['raw'] || args.query && args.query['raw'];
    isRaw = !!isRaw;

    let isRU = args['ru'] || args.query && args.query['ru'];
    isRU = isRU == '1' || isRU == 'true';
    const lang = isRU ? 'ru' : 'en';
    const $loc = (str) => {
      return isRU ? (_local[lang][str] || str) : str;
    };

    let quoteId = args['id'] || args.query && args.query['id'];
    let quoteNumber = args['number'] || args.query && args.query['number'];
    if (!quoteId && !quoteNumber) {
      const url = `${config.crm_api_url}/Quotes?fields=id`;
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          Authorization: `Zoho-oauthtoken ${accessToken}`
        }
      });

      const data = await response.json();

      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(data),
      };
    } else {
      let url = null;
      if (quoteId) {
        url = `${config.crm_api_url}/Quotes/${quoteId}`;
      } else if (quoteNumber) {
        url = `${config.crm_api_url}/Quotes/search?criteria=(Quote_Number:equals:"${quoteNumber}")`;
        console.log("searching by Quote_Number: ", url);

        const response = await fetch(url, {
          method: 'GET',
          headers: {
            Authorization: `Zoho-oauthtoken ${accessToken}`
          }
        });

        const data = await response.json();
        //console.log(JSON.stringify(data));

        quoteId = data.data[0].id;
        url = `${config.crm_api_url}/Quotes/${quoteId}`;
      } else {
        return { statusCode: 400, body: 'quote id or number must be provided' };
      }

      console.log(url);
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          Authorization: `Zoho-oauthtoken ${accessToken}`
        }
      });
      if (response.status != 200)
      {
        return { statusCode: response.status, body: `Status code is ${response.status}` };
      }
      const data = await response.json();

      if (isRaw) {
        return {
          statusCode: 200,
          headers: {
            'Content-Type': 'application/json'
          },
          body: data,
        };
      }

      const res = [];
      if (!data.data) {
        return { statusCode: 500, body: 'data.data empty' };
      }
      const item = data.data[0];
      if (!data.data[0]) {
        return { statusCode: 500, body: 'data.data[0] empty' };
      }

      res.push([$loc('Quote Id'), item.id]);
      res.push([$loc('Quote Number'), item.Quote_Number]);
      res.push([$loc('Account Name'), item.Account_Name.name]);
      res.push([$loc('Subject'), item.Subject]);
      res.push([$loc('Agent'), item.Agent.name]);
      res.push([$loc('Created By'), item.Created_By.name]);
      res.push([$loc('Contact Name'), item.Contact_Name.name]);
      res.push([$loc('Deal Name'), item.Deal_Name.name]);
      res.push([$loc('Created Time'), item.Created_Time]);
      res.push([$loc('Production Time'), item.Production_Time]);

      const currency_symbol = item['$currency_symbol'];

      let productPromises = [];
      for (const i of item.Quoted_Items) {
        const productId = i.Product_Name.id;
        {
          const url = `${config.crm_api_url}/Products/${productId}?fields=id,Product_Name,Product_Name_RU,Usage_Unit,Type,Group`;
          console.log(url);
          const response = fetch(url, {
            method: 'GET',
            headers: {
              Authorization: `Zoho-oauthtoken ${accessToken}`
            }
          });
          productPromises.push(response);
        }
      }

      const productDict = {};
      const productResponses = await Promise.all(productPromises);
      for (const response of productResponses) {
        const data = await response.json();
        console.log(JSON.stringify(data));
        if (data && data.data) {
          productDict[data.data[0].id] = data.data[0];
        }
      }

      let id = 1;
      for (const i of item.Quoted_Items) {
        if (!isRaw) {
          const productId = i.Product_Name.id;
          const productDesc = productDict[productId];
          const name = isRU
            ? (productDesc && productDesc.Product_Name_RU || i.Product_Name.name)
            : (productDesc && productDesc.Product_Name || i.Product_Name.name);

          res.push({
            id,
            product_id: productId,
            name,
            quantity: i.Quantity,
            usage_unit: productDesc.Usage_Unit || null,
            type: productDesc.Type || null,
            list_price: i.List_Price,
            unit_price: i.Unit_price,
            net_total: i.Net_Total,
            amount: i.Total,
            k: i.K,
            currency_symbol,
          });
        } else {
          // return 1:1 the whole object
          res.push(i);
        }
        id += 1;
      }

      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'application/json'
        },
        body: res,
      };
    }
  }
  catch (ex) {
    return {
      statusCode: 500,
      body: `OAuth error: ${JSON.stringify(ex.message)}`
    };
  }
};
