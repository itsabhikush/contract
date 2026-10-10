const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Busboy = require('busboy');

function getTemplatePath() {
  const candidates = [
    path.join(__dirname, 'Form_GST.docx'),
    path.join(__dirname, '..', '..', 'Form_GST.docx'),
    path.join(process.cwd(), 'netlify', 'functions', 'Form_GST.docx'),
    path.join(process.cwd(), 'Form_GST.docx')
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return path.join(__dirname, 'Form_GST.docx');
}

const TEMPLATE_PATH = getTemplatePath();

// ── Date formatter ──────────────────────────────────────────────────────────
function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d)) return dateStr;
  const day = d.getDate();
  let suffix = 'th';
  if (day % 100 < 11 || day % 100 > 13) {
    const ld = day % 10;
    if (ld === 1) suffix = 'st';
    else if (ld === 2) suffix = 'nd';
    else if (ld === 3) suffix = 'rd';
  }
  const months = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
  return `${day}${suffix} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

// ── XML token replacer ──────────────────────────────────────────────────────
function replaceInXml(xml, token, value) {
  const escaped = (value || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  return xml.split(token).join(escaped);
}

// ── ANNEXURE E cell updater ─────────────────────────────────────────────────
function setCellValue(cellXml, value) {
  if (!value) return cellXml.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, '<w:t></w:t>');
  const escaped = value.replace(/&/g,'&amp;');
  if (/<w:t/.test(cellXml))
    return cellXml.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, `<w:t>${escaped}</w:t>`);
  return cellXml.replace('</w:p>', `<w:r><w:t>${escaped}</w:t></w:r></w:p>`);
}

function updateAnnexureETable(docXml, guestsLocalId, unmarriedCouple, foreigner) {
  const anchorIdx = docXml.indexOf('Guests with Local ID');
  if (anchorIdx === -1) return docXml;
  const tableStart = docXml.lastIndexOf('<w:tbl>', anchorIdx);
  const foreignerIdx = docXml.indexOf('Foreigner', anchorIdx);
  const tableEnd = docXml.indexOf('</w:tbl>', foreignerIdx) + '</w:tbl>'.length;
  const beforeTable = docXml.substring(0, tableStart);
  const afterTable  = docXml.substring(tableEnd);
  let tableXml = docXml.substring(tableStart, tableEnd);

  const rowRegex = /<w:tr[ >][\s\S]*?<\/w:tr>/g;
  const rowMatches = [...tableXml.matchAll(rowRegex)];
  if (rowMatches.length < 4) return docXml;

  const dataRows = [
    { rowIndex: 1, selection: guestsLocalId },
    { rowIndex: 2, selection: unmarriedCouple },
    { rowIndex: 3, selection: foreigner },
  ];

  for (const { rowIndex, selection } of dataRows) {
    const rowMatch = rowMatches[rowIndex];
    if (!rowMatch) continue;
    const rowXml  = rowMatch[0];
    const cells   = [...rowXml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)];
    if (cells.length < 3) continue;
    const yesVal = selection === 'Yes' ? 'Yes' : '';
    const noVal  = selection === 'No'  ? 'No'  : '';
    let newRowXml = rowXml;
    newRowXml = newRowXml.replace(cells[1][0], setCellValue(cells[1][0], yesVal));
    newRowXml = newRowXml.replace(cells[2][0], setCellValue(cells[2][0], noVal));
    tableXml = tableXml.replace(rowXml, newRowXml);
  }
  return beforeTable + tableXml + afterTable;
}

// ── Multipart parser ────────────────────────────────────────────────────────
function parseMultipart(event) {
  return new Promise((resolve, reject) => {
    const contentType = event.headers['content-type'] || event.headers['Content-Type'] || '';
    const busboy = Busboy({ headers: { 'content-type': contentType } });
    const fields = {};
    const files  = {};

    busboy.on('file', (fieldname, file, info) => {
      const chunks = [];
      file.on('data', chunk => chunks.push(chunk));
      file.on('end', () => {
        files[fieldname] = { buffer: Buffer.concat(chunks), mimetype: info.mimeType, filename: info.filename };
      });
    });
    busboy.on('field', (fieldname, value) => { fields[fieldname] = value; });
    busboy.on('finish', () => resolve({ fields, files }));
    busboy.on('error', reject);

    const body = event.isBase64Encoded
      ? Buffer.from(event.body, 'base64')
      : Buffer.from(event.body || '');
    busboy.write(body);
    busboy.end();
  });
}

// ── Netlify handler ─────────────────────────────────────────────────────────
exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    const { fields, files } = await parseMultipart(event);

    const {
      DATE, GST_NO, GST_ADDRESS, OWNER_NAME, TRADE_NAME,
      HOTEL_NAME, HOTEL_ADDRESS, INVENTRY,
      LAT_D, LONG_D, MOBILE_NO, LANDLINE_NO, EMAIL_ADD,
      GUESTS_LOCAL_ID, UNMARRIED_COUPLE, FOREIGNER
    } = fields;

    const formattedDate = formatDate(DATE);
    const content = fs.readFileSync(TEMPLATE_PATH, 'binary');
    const zip = new PizZip(content);

    let docXml = zip.file('word/document.xml').asText();

    const replacements = {
      DATE: formattedDate, GST_NO, GST_ADDRESS, OWNER_NAME, TRADE_NAME,
      HOTEL_NAME, HOTEL_ADDRESS, INVENTRY,
      LAT_D, LONG_D, MOBILE_NO, LANDLINE_NO, EMAIL_ADD,
    };

    for (const [token, value] of Object.entries(replacements)) {
      docXml = replaceInXml(docXml, token, value || '');
    }

    docXml = updateAnnexureETable(docXml, GUESTS_LOCAL_ID || '', UNMARRIED_COUPLE || '', FOREIGNER || '');
    zip.file('word/document.xml', docXml);

    // Footer replacements
    Object.keys(zip.files).filter(f => f.startsWith('word/footer')).forEach(fFile => {
      let fXml = zip.file(fFile).asText();
      for (const [token, value] of Object.entries(replacements)) {
        fXml = replaceInXml(fXml, token, value || '');
      }
      zip.file(fFile, fXml);
    });

    // Replace hotel photo if uploaded
    if (files.hotelPhoto && files.hotelPhoto.buffer.length > 0) {
      zip.file('word/media/image3.png', files.hotelPhoto.buffer, { binary: true });
    }

    const output = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const safeHotel = (HOTEL_NAME || 'Contract').replace(/[^a-zA-Z0-9_\- ]/g, '').trim();
    const filename  = `Contract_${safeHotel}_${formattedDate}`.replace(/\s+/g,'_') + '.docx';

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
      body: output.toString('base64'),
      isBase64Encoded: true,
    };

  } catch (err) {
    console.error('Function error:', err);
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};
