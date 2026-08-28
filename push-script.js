const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');

async function getAccessToken() {
  const globalClaspRc = path.join(os.homedir(), '.clasprc.json');
  if (!fs.existsSync(globalClaspRc)) {
    throw new Error('No .clasprc.json found. Please run clasp login.');
  }

  const claspData = JSON.parse(fs.readFileSync(globalClaspRc, 'utf8'));
  const tokens = claspData.tokens?.default || claspData.token;
  if (!tokens) {
    throw new Error('No valid tokens found in .clasprc.json.');
  }

  const expiry = tokens.expiry_date || (tokens.expires_at ? tokens.expires_at * 1000 : 0);
  const isExpired = Date.now() > (expiry - 60000);

  if (!isExpired && tokens.access_token) {
    return tokens.access_token;
  }

  if (!tokens.refresh_token) {
    if (tokens.access_token) return tokens.access_token;
    throw new Error('Token is expired and no refresh_token available.');
  }

  console.log('Refreshing Google OAuth2 Access Token...');
  const clientId = tokens.client_id || '1072944905499-vm2v2i533lu0ckvthh1ehjk4opi20q49.apps.googleusercontent.com';
  const clientSecret = tokens.client_secret || 'v62Bp6IHluSODxPyp8tGzjhE';

  const postData = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: tokens.refresh_token,
    grant_type: 'refresh_token'
  }).toString();

  const refreshed = await new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'oauth2.googleapis.com',
      path: '/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.access_token) {
            resolve(parsed);
          } else {
            reject(new Error(`Failed to refresh token: ${data}`));
          }
        } catch (e) {
          reject(e);
        }
      });
    });
    req.on('error', reject);
    req.write(postData);
    req.end();
  });

  tokens.access_token = refreshed.access_token;
  if (refreshed.expires_in) {
    tokens.expiry_date = Date.now() + refreshed.expires_in * 1000;
  }
  if (claspData.tokens?.default) {
    claspData.tokens.default = tokens;
  } else {
    claspData.token = tokens;
  }
  fs.writeFileSync(globalClaspRc, JSON.stringify(claspData, null, 2), 'utf8');

  return refreshed.access_token;
}

async function pushToAppsScript() {
  const claspConfigPath = path.join(__dirname, '.clasp.json');
  if (!fs.existsSync(claspConfigPath)) {
    throw new Error('.clasp.json not found');
  }
  const claspConfig = JSON.parse(fs.readFileSync(claspConfigPath, 'utf8'));
  const scriptId = claspConfig.scriptId;

  console.log(`🚀 Preparing to push files to Google Apps Script (Script ID: ${scriptId})...`);

  const accessToken = await getAccessToken();

  const filesToPush = [
    {
      name: 'appsscript',
      type: 'JSON',
      filePath: path.join(__dirname, 'appsscript.json')
    },
    {
      name: 'Config',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'Config.js')
    },
    {
      name: 'DispatcherSchedule',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'DispatcherSchedule.js')
    },
    {
      name: 'DriveSync',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'DriveSync.js')
    },
    {
      name: 'AnalysisBuilder',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'AnalysisBuilder.js')
    },
    {
      name: 'ReadMeBuilder',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'ReadMeBuilder.js')
    },
    {
      name: 'MasterSummaryBuilder',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'MasterSummaryBuilder.js')
    },
    {
      name: 'DispatcherAuditBuilder',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'DispatcherAuditBuilder.js')
    },
    {
      name: 'SettingsManager',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'SettingsManager.js')
    },
    {
      name: 'Menu',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'Menu.js')
    },
    {
      name: 'Utils',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'Utils.js')
    },
    {
      name: 'MockDataGenerator',
      type: 'SERVER_JS',
      filePath: path.join(__dirname, 'MockDataGenerator.js')
    }
  ];

  const payloadFiles = filesToPush.map(f => {
    let content = fs.readFileSync(f.filePath, 'utf8');
    return {
      name: f.name,
      type: f.type,
      source: content
    };
  });

  const requestBody = JSON.stringify({ files: payloadFiles });

  console.log(`📦 Uploading ${payloadFiles.length} files:`);
  payloadFiles.forEach(f => console.log(`   • ${f.name} (${f.type})`));

  const result = await new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'script.googleapis.com',
      path: `/v1/projects/${scriptId}/content`,
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(requestBody)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (res.statusCode === 200) {
            resolve(parsed);
          } else {
            reject(new Error(`API Error (HTTP ${res.statusCode}): ${JSON.stringify(parsed, null, 2)}`));
          }
        } catch (e) {
          reject(new Error(`Non-JSON response (HTTP ${res.statusCode}): ${data}`));
        }
      });
    });
    req.on('error', reject);
    req.write(requestBody);
    req.end();
  });

  console.log('\n✅ PUSH SUCCESSFUL!');
  console.log(`🎉 Project files updated in Google Apps Script!`);
  return result;
}

pushToAppsScript().catch(err => {
  console.error('\n❌ Push failed:', err.message);
  process.exit(1);
});
