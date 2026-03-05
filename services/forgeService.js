const dotenv = require('dotenv');
dotenv.config();

// Patch fetch for CommonJS
const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));

async function getAccessToken() {
  const clientId = process.env.FORGE_CLIENT_ID;
  const clientSecret = process.env.FORGE_CLIENT_SECRET;

  const response = await fetch('https://developer.api.autodesk.com/authentication/v1/authenticate', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    body: `client_id=${clientId}&client_secret=${clientSecret}&grant_type=client_credentials&scope=data:read data:write data:create bucket:read`
  });

  if (!response.ok) {
    throw new Error(`Token request failed: ${response.statusText}`);
  }

  return response.json();
}

module.exports = { getAccessToken };
