const fetch = require('node-fetch');

// shared config hardcoded for visibility
const config = 
{
	"crm_api_url": "https://crm.zoho.eu/crm/v2"
};

exports.main = async function (args) {
  try {
    const accessToken = args['token'] || args.query && args.query['token'];
    if (!accessToken) {
      return { statusCode: 400, body: 'Missing token' };
    }

    let isRaw = args['raw'] || args.query && args.query['raw'];
    isRaw = !!isRaw;

    const quoteId = args['id'] || args.query && args.query['id'];
    if (!quoteId) {
      const url = `${config.crm_api_url}/quotes?fields=id`;
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
      const url = `${config.crm_api_url}/quotes/${quoteId}`;
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          Authorization: `Zoho-oauthtoken ${accessToken}`
        }
      });

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
      const datadata = data.data;
      if (!datadata) {
        return { statusCode: 400, body: 'datadata empty' };
      }
      const item = datadata[0];

      res.push([ 'Quote Id', item.id ]);
      res.push([ 'Quote Number', item.Quote_Number ]);
      res.push([ 'Account Name', item.Account_Name.name ]);
      res.push([ 'Subject', item.Subject ]);
      res.push([ 'Agent', item.Agent.name ]);
      res.push([ 'Created By', item.Created_By.name ]);
      res.push([ 'Contact Name', item.Contact_Name.name ]);
      res.push([ 'Deal Name', item.Deal_Name.name ]);
      res.push([ 'Created Time', item.Created_Time ]);
      res.push([ 'Production Time', item.Production_Time ]);

      const currency_symbol = item['$currency_symbol'];

      let id = 1;
      for (const i of item.Product_Details) {
        if (!isRaw) {
          res.push({
            id,
            name: i.product.name,
            quantity: i.quantity,
            list_price: i.list_price,
            unit_price: i.unit_price,
            total: i.total,
            total_after_discount: i.total_after_discount,
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
