const fetch = require('node-fetch');

// shared config hardcoded for visibility
const config =
{
  "crm_api_url": "https://crm.zoho.eu/crm/v4"
};

exports.main = async function (args) {
  try {
    const accessToken = args['token'] || args.query && args.query['token'];
    if (!accessToken) {
      return { statusCode: 400, body: 'Missing token' };
    }

    let isRaw = args['raw'] || args.query && args.query['raw'];
    isRaw = isRaw == 1 || isRaw == 'true';

    let isRU = args['ru'] || args.query && args.query['ru'];
    isRU = isRU == 1 || isRU == 'true';

    const productId = args['id'] || args.query && args.query['id'];
    if (!productId) {
      const url = `${config.crm_api_url}/Products?fields=id,name`;
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

      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(data.data[0]),
      };
    } else {
      const url = `${config.crm_api_url}/Products/${productId}?organization_id=20081329037`;
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

      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'application/json'
        },
        body: data.data[0],
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
