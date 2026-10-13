let gstHelper;
try {
  gstHelper = require('./gstHelper');
} catch (e) {
  gstHelper = require('../../gstHelper');
}
const { fetchGstCaptcha } = gstHelper;

exports.handler = async (event) => {
  if (event.httpMethod !== 'GET') {
    return {
      statusCode: 405,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ success: false, error: 'Method Not Allowed' })
    };
  }
  try {
    const data = await fetchGstCaptcha();
    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({ success: true, ...data })
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ success: false, error: err.message || 'Failed to fetch captcha' })
    };
  }
};
