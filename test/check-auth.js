const fs = require('fs');
const path = require('path');
const os = require('os');

const globalClaspRc = path.join(os.homedir(), '.clasprc.json');
if (fs.existsSync(globalClaspRc)) {
  const content = fs.readFileSync(globalClaspRc, 'utf8');
  console.log('File size:', content.length);
  try {
    const creds = JSON.parse(content);
    console.log('Keys in .clasprc.json:', Object.keys(creds));
    if (creds.token) console.log('Keys in creds.token:', Object.keys(creds.token));
    if (creds.tokens) console.log('Keys in creds.tokens:', Object.keys(creds.tokens));
    if (creds.oauth2ClientSettings) console.log('Keys in creds.oauth2ClientSettings:', Object.keys(creds.oauth2ClientSettings));
  } catch (e) {
    console.log('JSON parse error:', e.message);
  }
}
