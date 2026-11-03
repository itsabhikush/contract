const https = require('https');

const SANDBOX_API_KEY = process.env.SANDBOX_API_KEY || 'key_live_3a0f223b4ad04f9994e8b41296205e20';
const SANDBOX_API_SECRET = process.env.SANDBOX_API_SECRET || 'secret_live_254b88a313454d73bfa156d449b1bb66';
const SANDBOX_BASE_URL = 'https://api.sandbox.co.in';

let cachedToken = null;
let tokenExpiry = 0;

/**
 * Authenticate with Sandbox API and cache the JWT token for 23 hours.
 */
function getSandboxToken() {
  if (cachedToken && Date.now() < tokenExpiry - 60000) {
    return Promise.resolve(cachedToken);
  }

  return new Promise((resolve, reject) => {
    const req = https.request(SANDBOX_BASE_URL + '/authenticate', {
      method: 'POST',
      timeout: 10000,
      headers: {
        'accept': 'application/json',
        'x-api-key': SANDBOX_API_KEY,
        'x-api-secret': SANDBOX_API_SECRET,
        'x-api-version': '1.0.0',
        'Content-Length': 0
      }
    }, res => {
      let body = '';
      res.on('data', d => body += d);
      res.on('end', () => {
        try {
          const json = JSON.parse(body);
          const token = json.access_token || (json.data && json.data.access_token);
          if (token) {
            cachedToken = token;
            tokenExpiry = Date.now() + 23 * 3600 * 1000;
            return resolve(token);
          }
          reject(new Error(json.message || 'Failed to authenticate with Sandbox API'));
        } catch (e) {
          reject(new Error('Invalid response from Sandbox auth endpoint'));
        }
      });
    });

    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Sandbox authentication timed out'));
    });
    req.on('error', reject);
    req.end();
  });
}

const GST_STATES = {
  "01": "Jammu and Kashmir",
  "02": "Himachal Pradesh",
  "03": "Punjab",
  "04": "Chandigarh",
  "05": "Uttarakhand",
  "06": "Haryana",
  "07": "Delhi",
  "08": "Rajasthan",
  "09": "Uttar Pradesh",
  "10": "Bihar",
  "11": "Sikkim",
  "12": "Arunachal Pradesh",
  "13": "Nagaland",
  "14": "Manipur",
  "15": "Mizoram",
  "16": "Tripura",
  "17": "Meghalaya",
  "18": "Assam",
  "19": "West Bengal",
  "20": "Jharkhand",
  "21": "Odisha",
  "22": "Chhattisgarh",
  "23": "Madhya Pradesh",
  "24": "Gujarat",
  "25": "Daman and Diu",
  "26": "Dadra and Nagar Haveli and Daman and Diu",
  "27": "Maharashtra",
  "29": "Karnataka",
  "30": "Goa",
  "31": "Lakshadweep",
  "32": "Kerala",
  "33": "Tamil Nadu",
  "34": "Puducherry",
  "35": "Andaman and Nicobar Islands",
  "36": "Telangana",
  "37": "Andhra Pradesh",
  "38": "Ladakh",
  "97": "Other Territory",
  "99": "CBIC"
};

/**
 * Formats GST address into the exact official order as per Form GST REG-06 and GST Portal:
 * 1. Floor Number (flno)
 * 2. Building / Flat / Door Number (bno)
 * 3. Name of Premises / Building (bnm)
 * 4. Road / Street (st)
 * 5. Locality / Sub-locality (locality)
 * 6. Landmark (landMark)
 * 7. City / Town / Village / Location (loc)
 * 8. District (dst)
 * 9. State Name (stcd)
 * 10. PIN Code (pncd)
 */
function formatGstAddress(pradr) {
  if (!pradr) return '';
  if (typeof pradr === 'string') return pradr.trim();
  if (pradr.adr && typeof pradr.adr === 'string' && pradr.adr.trim()) {
    return pradr.adr.trim();
  }

  const a = pradr.addr || pradr;
  if (!a || typeof a !== 'object') return '';

  let stateName = a.stcd || '';
  if (stateName && GST_STATES[stateName]) {
    stateName = GST_STATES[stateName];
  }

  const parts = [
    a.flno,
    a.bno,
    a.bnm,
    a.st,
    a.locality,
    a.landMark,
    a.loc,
    a.dst,
    stateName,
    a.pncd
  ];

  const cleaned = [];
  for (let part of parts) {
    if (part === undefined || part === null) continue;
    const s = String(part).trim();
    if (!s) continue;
    // Skip consecutive identical tokens
    if (cleaned.length > 0 && cleaned[cleaned.length - 1].toLowerCase() === s.toLowerCase()) {
      continue;
    }
    cleaned.push(s);
  }

  return cleaned.join(', ');
}

/**
 * Verify GST taxpayer details using Sandbox.co.in API (no captcha required).
 */
async function searchGstTaxpayer(gstin) {
  const cleanGstin = (gstin || '').trim().toUpperCase();
  if (!cleanGstin || cleanGstin.length !== 15) {
    return {
      success: false,
      error: 'GST Number must be exactly 15 characters.'
    };
  }

  let token = await getSandboxToken();
  const payload = JSON.stringify({ gstin: cleanGstin });

  const executeRequest = (authToken) => {
    return new Promise((resolve, reject) => {
      const req = https.request(SANDBOX_BASE_URL + '/gst/compliance/public/gstin/search', {
        method: 'POST',
        timeout: 15000,
        headers: {
          'Content-Type': 'application/json',
          'authorization': authToken,
          'x-api-key': SANDBOX_API_KEY,
          'x-api-version': '1.0.0',
          'Content-Length': Buffer.byteLength(payload)
        }
      }, res => {
        let body = '';
        res.on('data', d => body += d);
        res.on('end', () => {
          resolve({ statusCode: res.statusCode, body });
        });
      });

      req.on('timeout', () => {
        req.destroy();
        reject(new Error('GST verification timed out. Please try again.'));
      });
      req.on('error', reject);
      req.write(payload);
      req.end();
    });
  };

  let response = await executeRequest(token);

  // If token expired, refresh once and retry
  if (response.statusCode === 401 || response.statusCode === 403) {
    cachedToken = null;
    token = await getSandboxToken();
    response = await executeRequest(token);
  }

  try {
    const json = JSON.parse(response.body);
    if (response.statusCode !== 200) {
      return {
        success: false,
        error: json.message || json.error || 'GSTIN verification failed.'
      };
    }

    const tpData = (json.data && json.data.data) || json.data || json;
    const address = formatGstAddress(tpData.pradr || (Array.isArray(tpData.adadr) && tpData.adadr[0]));

    return {
      success: true,
      gstin: tpData.gstin || cleanGstin,
      sts: tpData.sts || 'Active',
      dty: tpData.dty || 'Regular',
      lgnm: tpData.lgnm || '',
      tradeNam: tpData.tradeNam || tpData.lgnm || '',
      address: address,
      ctb: tpData.ctb || '',
      rgdt: tpData.rgdt || '',
      raw: tpData
    };
  } catch (e) {
    throw new Error('Failed to parse response from GST API: ' + response.body.slice(0, 100));
  }
}

/**
 * Legacy Captcha fetcher for PAN modal search (services.gst.gov.in)
 */
function fetchGstCaptcha() {
  return new Promise((resolve, reject) => {
    const req = https.request('https://services.gst.gov.in/services/captcha', {
      method: 'GET',
      timeout: 10000,
      headers: {
        'User-Agent': 'curl/8.7.1',
        'Accept': '*/*'
      }
    }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        if (res.statusCode !== 200) {
          return reject(new Error('GST Portal returned HTTP ' + res.statusCode));
        }
        const setCookies = res.headers['set-cookie'] || [];
        const cookieHeader = setCookies.map(c => c.split(';')[0]).join('; ');
        const imgBuffer = Buffer.concat(chunks);
        const captchaImage = 'data:image/png;base64,' + imgBuffer.toString('base64');
        const sessionId = Buffer.from(cookieHeader).toString('base64');
        resolve({ captchaImage, sessionId });
      });
    });
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('GST Portal connection timed out.'));
    });
    req.on('error', reject);
    req.end();
  });
}

/**
 * Search GST by PAN via GST Portal
 */
function searchGstByPan(pan, captcha, sessionId) {
  return new Promise((resolve, reject) => {
    let cookieHeader = '';
    try {
      cookieHeader = Buffer.from(sessionId, 'base64').toString('utf8');
    } catch {
      return reject(new Error('Invalid or expired captcha session. Please refresh captcha.'));
    }

    const cleanPan = (pan || '').trim().toUpperCase();
    const payload = JSON.stringify({
      panNO: cleanPan,
      captcha: (captcha || '').trim()
    });

    const req = https.request('https://services.gst.gov.in/services/api/get/gstndtls', {
      method: 'POST',
      timeout: 12000,
      headers: {
        'User-Agent': 'curl/8.7.1',
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json;charset=UTF-8',
        'Content-Length': Buffer.byteLength(payload),
        'Referer': 'https://services.gst.gov.in/services/searchtpbypan',
        'Origin': 'https://services.gst.gov.in',
        'Cookie': cookieHeader
      }
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.errorCode) {
            if (json.errorCode === 'SWEB_9000') {
              return resolve({
                success: false,
                error: 'Incorrect captcha code. Please try again.',
                isCaptchaError: true
              });
            }
            if (json.errorCode === 'SWEB_10001') {
              return resolve({
                success: true,
                pan: cleanPan,
                notFound: true,
                count: 0,
                gstinList: [],
                message: `No GST records found registered under PAN: ${cleanPan}`
              });
            }
            if (json.errorCode === 'SWEB_9032' || json.errorCode === 'SWEB_9034' || json.errorCode === 'SWEB_9035') {
              return resolve({
                success: false,
                error: 'Invalid PAN Number. Please verify your 10-character PAN number.',
                isInvalidPan: true
              });
            }
            return resolve({
              success: false,
              error: `GST Portal Error (${json.errorCode}): ${json.message || 'Search failed.'}`
            });
          }

          const rawList = Array.isArray(json.gstinResList) ? json.gstinResList : [];
          if (rawList.length === 0) {
            return resolve({
              success: true,
              pan: cleanPan,
              notFound: true,
              count: 0,
              gstinList: [],
              message: `No GST records found registered under PAN: ${cleanPan}`
            });
          }

          const gstinList = rawList.map((item, idx) => {
            const gstin = (item.gstin || '').trim();
            const stateCd = item.stateCd || (gstin.length >= 2 ? gstin.slice(0, 2) : '');
            const stateName = item.stateName || GST_STATES[stateCd] || (stateCd ? `State (${stateCd})` : 'N/A');
            return {
              sNo: idx + 1,
              gstin: gstin,
              authStatus: item.authStatus || item.status || item.sts || 'Active',
              stateCd: stateCd,
              stateName: stateName
            };
          });

          resolve({
            success: true,
            pan: cleanPan,
            notFound: false,
            count: gstinList.length,
            gstinList: gstinList,
            raw: json
          });
        } catch {
          reject(new Error('Invalid response from GST Portal. Response: ' + data.slice(0, 100)));
        }
      });
    });

    req.on('timeout', () => {
      req.destroy();
      reject(new Error('GST Portal search timed out.'));
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

module.exports = {
  fetchGstCaptcha,
  searchGstTaxpayer,
  searchGstByPan,
  formatGstAddress,
  GST_STATES
};
