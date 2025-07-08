const fetch = require('node-fetch');

exports.main = async function (args) {
  try {
    const accessToken = args['token'] || args.query && args.query['token'];
    if (!accessToken) {
      return { statusCode: 400, body: 'Missing token' };
    }

    const refreshUrl = `https://crm.zoho.eu/crm/v2/organizations?include=info&from=crm_org_profile`;
    const response = await fetch(refreshUrl, {
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
      body: data,
    };
  }
  catch (ex) {
    return {
      statusCode: 500,
      body: `OAuth error: ${JSON.stringify(ex.message)}`
    };
  }
};
