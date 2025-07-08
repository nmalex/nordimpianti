// shared config hardcoded for visibility
const config = 
{
	"crm_api_url":             "https://crm.zoho.eu/crm/v2",
  "DigitalOceanFunctionUrl": "",
  "ZohoClientID": "",
  "ZohoSecret": ""
};

exports.main = async function (args) {
  try {
    const redirectUri = `${config.DigitalOceanFunctionUrl}/functions/oauth-callback`;

    const authUrl = new URL('https://accounts.zoho.com/oauth/v2/auth');
    authUrl.searchParams.set('client_id', config.ZohoClientID);
    authUrl.searchParams.set('redirect_uri', redirectUri);
    authUrl.searchParams.set('response_type', 'code');
    authUrl.searchParams.set('scope', 'ZohoCRM.org.READ ZohoSubscriptions.quotes.READ ZohoSubscriptions.subscriptions.READ ZohoSubscriptions.customers.READ ZohoCRM.modules.ALL');
    authUrl.searchParams.set('access_type', 'offline');
    authUrl.searchParams.set('prompt', 'consent');

    return {
      statusCode: 302,
      headers: {
        Location: authUrl.toString()
      },
      body: ''
    };
  }
  catch (ex) {
    return {
      statusCode: 500,
      body: `OAuth error: ${JSON.stringify(ex.message)}`
    };
  }
};
