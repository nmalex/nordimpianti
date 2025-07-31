const fetch = require('node-fetch');

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
    const refreshToken = args['refresh-token'] || args.query && args.query['refresh-token'];
    const accountsServer = args['accounts-server'] || args.query && args.query['accounts-server'];
    const format = args['format'] || args.query && args.query['format'];
    const isJsonOutput = format == 'json';

    if (!refreshToken) {
      return { statusCode: 400, body: 'Missing refresh token' };
    }
    if (!accountsServer) {
      return { statusCode: 400, body: 'Missing accounts server' };
    }

    const params = new URLSearchParams({
      client_id: config.ZohoClientID,
      client_secret: config.ZohoSecret,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    });

    const refreshUrl = `${accountsServer}/oauth/v2/token`;
    const response = await fetch(refreshUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString()
    });

    const data = await response.json();

    const DO = `${config.DigitalOceanFunctionUrl}/functions/oauth-refresh`;
    if (data.access_token) {

      const refreshUrl = `${DO}?refresh-token=${refreshToken}&accounts-server=${accountsServer}`;
      const refreshUrlEl = `<a href=\"${refreshUrl}\">link</a>`;

      if (isJsonOutput) {

        return {
          statusCode: 200,
          headers: {
            'Content-Type': 'application/json'
          },
          body: {
            access_token: data.access_token,
            refresh_token: refreshToken,
            accounts_server: accountsServer,
            refresh_url: refreshUrl,
          },
        };
      }

      let html = '';
      // render html page
      {
        html += `<!DOCTYPE html>
<html>
<head>
<title>OAuth Refresh Success</title>
<script>
    function copyText() {
      const textToCopy = 'access-token\\t${data.access_token}\\r\\nrefresh-token\\t${refreshToken}\\r\\naccounts-server\\t${accountsServer}\\r\\nrefresh URL\\t${refreshUrl}';
      navigator.clipboard.writeText(textToCopy)
        .then(() => {
          alert('Copied to clipboard!');
        })
        .catch(err => {
          console.error('Failed to copy:', err);
        });
    }
</script>
</head>
<body>

<h1>Tokens refreshed:</h1>
<button onclick="copyText()">Copy</button>
<p>Hint: copy it to the last sheet of the google spreadsheets.</p>
<table>
<tr><td>access-token</td><td>${data.access_token}</td></tr>
<tr><td>refresh-token</td><td>${refreshToken}</td></tr>
<tr><td>accounts-server</td><td>${accountsServer}</td></tr>
<tr><td>refresh URL</td><td>${refreshUrlEl}</td></tr>
</table>

</body>
</html>`;
      }

      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'text/html'
        },
        body: html,
      };
    } else {
      return {
        statusCode: 500,
        body: `OAuth error: ${JSON.stringify(data)}`
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
