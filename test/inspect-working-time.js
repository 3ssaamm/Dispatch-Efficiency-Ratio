const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

async function getAccessToken() {
  const globalClaspRc = path.join(os.homedir(), '.clasprc.json');
  const claspData = JSON.parse(fs.readFileSync(globalClaspRc, 'utf8'));
  const tokens = claspData.tokens?.default || claspData.token;
  return tokens.access_token;
}

async function run() {
  const spreadsheetId = '1bDHi8a2TuopuWyM1vOH3cA-Vx3umsQqosipCWAhr9Fc';
  const token = await getAccessToken();
  
  const req = https.get({
    hostname: 'sheets.googleapis.com',
    path: `/v4/spreadsheets/${spreadsheetId}`,
    headers: {
      'Authorization': `Bearer ${token}`
    }
  }, (res) => {
    let data = '';
    res.on('data', chunk => data += chunk);
    res.on('end', () => {
      console.log('Status code:', res.statusCode);
      try {
        const json = JSON.parse(data);
        console.log('Response JSON:', JSON.stringify(json, null, 2).slice(0, 1000));
      } catch (e) {
        console.log('Raw data:', data);
      }
    });
  });
  req.on('error', console.error);
}

run().catch(console.error);
