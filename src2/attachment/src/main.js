const fetch = require('node-fetch');

// shared config hardcoded for visibility
const config =
{
  "crm_api_url": "https://crm.zoho.eu/crm/v4"
};

const https = require('https'); // Use 'http' if the URL starts with http://
const fs = require('fs');
const path = require('path');
const pngToJpeg = require('png-to-jpeg');

// Function to download and save image
async function downloadAttachment(moduleName, parentId, attachmentId, accessToken, filePath) {
  const p = new Promise(
    (resolve, reject) => {
      const file = fs.createWriteStream(filePath);

      // Use the appropriate protocol ('https' or 'http')
      const protocol = https;

      // Define the request options, including the Authorization header
      const options = {
        headers: {
          Authorization: `Zoho-oauthtoken ${accessToken}`
        }
      };

      const apiUrl = `${config.crm_api_url}/${moduleName}/${parentId}/Attachments/${attachmentId}`;
      console.log(apiUrl);

      protocol.get(apiUrl, options, (response) => {
        // Check if the response status is 200 (OK)
        console.log(`response ${response.statusCode}`);

        if (response.statusCode === 200) {
          response.pipe(file); // Pipe the image data to the file

          file.on('finish', () => {
            file.close(() => {
              console.log(`Image downloaded and saved as ${filePath}`);
              resolve();
            });
          });
        } else {
          console.error(`Failed to download image. Status code: ${response.statusCode}`);
          reject();
        }
      }).on('error', (err) => {
        console.error(`Error downloading image: ${err.message}`);
        reject();
      });
    });

  return p;
}

exports.main = async function (args) {
  try {
    const accessToken = args['token'] || args.query && args.query['token'];
    if (!accessToken) {
      return { statusCode: 400, body: 'Missing token' };
    }
    let moduleName = args['module_name'] || args.query && args.query['module_name'];
    if (!moduleName) {
      return { statusCode: 400, body: 'Missing module_name' };
    }
    let attachmentId = args['attachment_id'] || args.query && args.query['attachment_id'];
    if (!attachmentId) {
      return { statusCode: 400, body: 'Missing attachment_id' };
    }
    let parentId = args['parent_id'] || args.query && args.query['parent_id'];
    if (!parentId) {
      return { statusCode: 400, body: 'Missing parent_id' };
    }
    const fileName = args['file_name'] || args.query && args.query['file_name'];
    if (!fileName) {
      return { statusCode: 400, body: 'Missing file_Name' };
    }

    // Temporary file path inside /tmp
    let tempFilePath = path.join('/tmp', `${fileName}`);

    // Download and save the image
    console.log(`downloadAttachment(${moduleName}, ${parentId}, ${attachmentId}, ${accessToken}, ${tempFilePath})`);
    await downloadAttachment(moduleName, parentId, attachmentId, accessToken, tempFilePath);

    // Check if the image file exists
    if (fs.existsSync(tempFilePath)) {

      let imageBuffer = fs.readFileSync(tempFilePath);  // Read the image as a buffer
      let imageSize = imageBuffer.length;
      console.log(`The size of the image is ${imageSize} bytes.`);

      // Set the appropriate header based on the image type
      let contentType = '';
      if (fileName.endsWith('.png')) {
        console.log("Converting to JPEG...");

        tempFilePath = tempFilePath.replace(".png", ".jpg");

        const output = await pngToJpeg({quality: 90})(imageBuffer);
        fs.writeFileSync(tempFilePath, output);

        imageBuffer = fs.readFileSync(tempFilePath);  // Read the image as a buffer
        imageSize = imageBuffer.length;
        console.log(`The size of the image is ${imageSize} bytes.`);

        contentType = 'image/jpeg';
      } else if (fileName.endsWith('.jpg')) {
        contentType = 'image/jpeg';
      }
      if (!contentType) {
        return {
          statusCode: 500,
          body: `Image type not supported: ${fileName}`
        };
      }

      return {
        statusCode: 200,
        headers: {
          'Content-Type': contentType,
        },
        body: imageBuffer.toString('base64'),
        isBase64Encoded: true,  // Since we're sending the image as base64 encoded
      };
    } else {
      return {
        statusCode: 500,
        body: `Image not found`
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
