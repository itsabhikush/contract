const express = require('express');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const multer = require('multer');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

app.use(express.static(path.join(__dirname, 'public')));

function getTemplatePath(contractType) {
  const filename = contractType === 'NGST' ? 'Form_NGST.docx' : 'Form_GST.docx';
  const candidates = [
    path.join(__dirname, filename),
    path.join(__dirname, 'netlify', 'functions', filename),
    path.join(process.cwd(), filename),
    path.join(process.cwd(), 'netlify', 'functions', filename),
    path.join(process.cwd(), contractType === 'NGST' ? 'formet NGST.docx' : 'Form_GST.docx')
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return path.join(__dirname, filename);
}

// Format date: "2026-10-11" -> "11th October 2026"
function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d)) return dateStr;
  const day = d.getDate();
  let suffix = 'th';
  if (day % 100 < 11 || day % 100 > 13) {
    const lastDigit = day % 10;
    if (lastDigit === 1) suffix = 'st';
    else if (lastDigit === 2) suffix = 'nd';
    else if (lastDigit === 3) suffix = 'rd';
  }
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return `${day}${suffix} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

// Replace plain text tokens in OOXML
function replaceInXml(xml, token, value) {
  const escaped = (value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return xml.split(token).join(escaped);
}

// Build a Yes/No/blank cell run XML for the ANNEXURE E table
function setCellValue(cellXml, value) {
  if (!value) {
    return cellXml.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, '<w:t></w:t>');
  }
  const escaped = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  if (/<w:t[ >]/.test(cellXml)) {
    return cellXml.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, `<w:t>${escaped}</w:t>`);
  }
  return cellXml.replace('</w:p>', `<w:r><w:t>${escaped}</w:t></w:r></w:p>`);
}

// Update ANNEXURE E table rows 1-3
function updateAnnexureETable(docXml, guestsLocalId, unmarriedCouple, foreigner) {
  const anchorIdx = docXml.indexOf('Guests with Local ID');
  if (anchorIdx === -1) return docXml;

  const tableStart = docXml.lastIndexOf('<w:tbl>', anchorIdx);
  const foreignerIdx = docXml.indexOf('Foreigner', anchorIdx);
  const tableEnd = docXml.indexOf('</w:tbl>', foreignerIdx) + '</w:tbl>'.length;

  const beforeTable = docXml.substring(0, tableStart);
  const afterTable = docXml.substring(tableEnd);
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

    const rowXml = rowMatch[0];
    const cellRegex = /<w:tc>[\s\S]*?<\/w:tc>/g;
    const cells = [...rowXml.matchAll(cellRegex)];
    if (cells.length < 3) continue;

    const yesVal = selection === 'Yes' ? 'Yes' : '';
    const noVal = selection === 'No' ? 'No' : '';

    let newRowXml = rowXml;
    const originalYesCell = cells[1][0];
    const originalNoCell = cells[2][0];
    const newYesCell = setCellValue(originalYesCell, yesVal);
    const newNoCell = setCellValue(originalNoCell, noVal);

    newRowXml = newRowXml.replace(originalYesCell, newYesCell);
    newRowXml = newRowXml.replace(originalNoCell, newNoCell);

    tableXml = tableXml.replace(rowXml, newRowXml);
  }

  return beforeTable + tableXml + afterTable;
}

app.post('/generate', upload.single('hotelPhoto'), (req, res) => {
  try {
    const {
      contractType,
      DATE, GST_NO, GST_ADDRESS,
      PAN_NO, PAN_ADDRESS, ADHAR_ADDRESS,
      OWNER_NAME, TRADE_NAME,
      HOTEL_NAME, HOTEL_ADDRESS, INVENTRY,
      LAT_D, LONG_D, MOBILE_NO, LANDLINE_NO, EMAIL_ADD,
      GUESTS_LOCAL_ID, UNMARRIED_COUPLE, FOREIGNER
    } = req.body;

    const isNGST = contractType === 'NGST';
    const templatePath = getTemplatePath(isNGST ? 'NGST' : 'GST');

    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template not found at ${templatePath}`);
    }

    const formattedDate = formatDate(DATE);
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);

    let docXml = zip.file('word/document.xml').asText();

    // Replacements dictionary covering both GST and NGST placeholders
    const replacements = {
      'DATE': formattedDate,
      'GST_NO': GST_NO || '',
      'GST_ADDRESS': GST_ADDRESS || '',
      'PAN_NO': PAN_NO || '',
      'PAN_ADDRESS': PAN_ADDRESS || '',
      'ADHAR_ADDRESS': ADHAR_ADDRESS || '',
      'OWNER_NAME': OWNER_NAME || '',
      'TRADE_NAME': TRADE_NAME || '',
      'HOTEL_NAME': HOTEL_NAME || '',
      'HOTEL_ADDRESS': HOTEL_ADDRESS || '',
      'INVENTRY': INVENTRY || '',
      'LAT_D': LAT_D || '',
      'LONG_D': LONG_D || '',
      'MOBILE_NO': MOBILE_NO || '',
      'LANDLINE_NO': LANDLINE_NO || '',
      'EMAIL_ADD': EMAIL_ADD || '',
    };

    for (const [token, value] of Object.entries(replacements)) {
      docXml = replaceInXml(docXml, token, value);
    }

    // Update ANNEXURE E table
    docXml = updateAnnexureETable(docXml, GUESTS_LOCAL_ID || '', UNMARRIED_COUPLE || '', FOREIGNER || '');

    zip.file('word/document.xml', docXml);

    // Replace Brand Display photo if uploaded (Annexure C-1: media/image3.png)
    if (req.file && req.file.buffer && req.file.buffer.length > 0) {
      zip.file('word/media/image3.png', req.file.buffer, { binary: true });
    }

    // Process footers
    const footerFiles = Object.keys(zip.files).filter(f => f.startsWith('word/footer'));
    for (const fFile of footerFiles) {
      let fXml = zip.file(fFile).asText();
      for (const [token, value] of Object.entries(replacements)) {
        fXml = replaceInXml(fXml, token, value);
      }
      zip.file(fFile, fXml);
    }

    const output = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const safeHotelName = (HOTEL_NAME || 'Contract').replace(/[^a-zA-Z0-9_\- ]/g, '').trim();
    const typeLabel = isNGST ? 'NGST' : 'GST';
    const filename = (`Contract_${typeLabel}_${safeHotelName}_${formattedDate || 'draft'}`).replace(/\s+/g, '_') + '.docx';

    res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(output);

  } catch (err) {
    console.error('Error generating docx:', err);
    res.status(500).json({ error: 'Failed to generate document: ' + err.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => console.log('Contract Form Server running at http://localhost:' + PORT));
