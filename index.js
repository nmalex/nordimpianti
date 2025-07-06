/**
 * A simple DigitalOcean function (Node.js)
 * Responds with "Hello from DigitalOcean Function!"
 */
module.exports = async function (context) {
  return {
    statusCode: 200,
    body: "Hello from DigitalOcean Function!",
  };
};
