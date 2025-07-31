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
    const code = args['code'] || args.query && args.query['code'];;
    if (!code) {
      return { statusCode: 400, body: 'Missing authorization code' };
    }
    const accountsServer = args['accounts-server'] || args.query['accounts-server'];
    if (!accountsServer) {
      return { statusCode: 400, body: 'Missing accounts-server' };
    }
    const format = args['format'] || args.query && args.query['format'];
    const isJsonOutput = format == 'json';
    let redirectUri = `${config.DigitalOceanFunctionUrl}/functions/oauth-callback`;
    if (isJsonOutput) {
      redirectUri += `?format=json`;
    }

    const params = new URLSearchParams({
      code,
      client_id: config.ZohoClientID,
      client_secret: config.ZohoSecret,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code'
    });

    const response = await fetch(`${accountsServer}/oauth/v2/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString()
    });

    const data = await response.json();

    const DO = `${config.DigitalOceanFunctionUrl}/functions/oauth-refresh`;
    if (data.access_token) {

      const refreshUrl = `${DO}?refresh-token=${data.refresh_token}&accounts-server=${accountsServer}`;
      const refreshUrlEl = `<a href=\"${refreshUrl}\">link</a>`;

      if (isJsonOutput) {
        const meta = {
          oauth: {
            "access-token": data.access_token,
            "refresh-token": data.refresh_token,
            "accounts-server": accountsServer,
            "refresh-url": refreshUrl,
          }
        };
        let metaText = JSON.stringify(meta);
        const obj = {
          text: "meta: " + metaText,
          meta,
        };
        return {
          statusCode: 200,
          headers: {
            'Content-Type': 'application/json'
          },
          body: obj,
        };
      }

      let html = '';
      // render html page
      {
        html += `<!DOCTYPE html>
<html>
<head>
<title>OAuth Success</title>
<script>
    function copyText() {
      const textToCopy = 'access-token\\t${data.access_token}\\r\\nrefresh-token\\t${data.refresh_token}\\r\\naccounts-server\\t${accountsServer}\\r\\nrefresh URL\\t${refreshUrl}';
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

<h1>Tokens received:</h1>
<button onclick="copyText()">Copy</button>
<p>Hint: copy it to the last sheet of the google spreadsheets.</p>
<table>
<tr><td>access-token</td><td>${data.access_token}</td></tr>
<tr><td>refresh-token</td><td>${data.refresh_token}</td></tr>
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
