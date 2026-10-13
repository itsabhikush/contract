const https = require('https');

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

function searchGstTaxpayer(gstin, captcha, sessionId) {
  return new Promise((resolve, reject) => {
    let cookieHeader = '';
    try {
      cookieHeader = Buffer.from(sessionId, 'base64').toString('utf8');
    } catch {
      return reject(new Error('Invalid or expired captcha session. Please refresh captcha.'));
    }

    const payload = JSON.stringify({
      gstin: (gstin || '').trim().toUpperCase(),
      captcha: (captcha || '').trim()
    });

    const req = https.request('https://services.gst.gov.in/services/api/search/taxpayerDetails', {
      method: 'POST',
      timeout: 12000,
      headers: {
        'User-Agent': 'curl/8.7.1',
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json;charset=UTF-8',
        'Content-Length': Buffer.byteLength(payload),
        'Referer': 'https://services.gst.gov.in/services/searchtp',
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
            if (json.errorCode === 'SWEB_9035') {
              return resolve({
                success: false,
                error: 'Invalid GSTIN / UIN. Please verify your 15-character GST number.',
                isInvalidGstin: true
              });
            }
            return resolve({
              success: false,
              error: `GST Portal Error (${json.errorCode}): ${json.message || 'Verification failed.'}`
            });
          }

          // Build principal place of business address
          let address = '';
          if (json.pradr && json.pradr.adr) {
            address = json.pradr.adr;
          } else if (json.pradr && json.pradr.addr) {
            const a = json.pradr.addr;
            address = [a.bno, a.bnm, a.st, a.loc, a.dst, a.stcd, a.pncd].filter(Boolean).join(', ');
          }

          resolve({
            success: true,
            gstin: json.gstin || gstin,
            sts: json.sts || '', // e.g. "Active"
            dty: json.dty || '', // e.g. "Regular"
            lgnm: json.lgnm || '', // Legal Name of Business
            tradeNam: json.tradeNam || json.lgnm || '', // Trade Name
            address: address, // Principle Place of Business
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

module.exports = { fetchGstCaptcha, searchGstTaxpayer, searchGstByPan, GST_STATES };
