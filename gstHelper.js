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

module.exports = { fetchGstCaptcha, searchGstTaxpayer };
