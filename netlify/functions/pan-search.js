let gstHelper;
try {
  gstHelper = require('./gstHelper');
} catch (e) {
  gstHelper = require('../../gstHelper');
}
const { searchGstByPan } = gstHelper;

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ success: false, error: 'Method Not Allowed' })
    };
  }
  try {
    const body = JSON.parse(event.body || '{}');
    const { pan, captcha, sessionId } = body;
    if (!pan || !captcha || !sessionId) {
      return {
        statusCode: 400,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ success: false, error: 'Missing PAN, Captcha, or Session ID' })
      };
    }
    const result = await searchGstByPan(pan, captcha, sessionId);
    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify(result)
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ success: false, error: err.message || 'Failed to search GST by PAN' })
    };
  }
};
