const express = require('express');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const multer = require('multer');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

app.use(express.static(path.join(__dirname, 'public')));

const TEMPLATE_PATH = path.join(__dirname, 'Form_GST.docx');

// Format date: "2026-10-11" -> "11th October 2026"
function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d)) return dateStr;
  const day = d.getDate();
  // 11, 12, 13 are always "th"; otherwise last digit decides
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
// Given the existing cell XML (which has the paragraph/run structure), replace the w:t text content
function setCellValue(cellXml, value) {
  // value: 'Yes', 'No', or '' (blank)
  if (!value) {
    // Remove any w:t content but keep structure
    return cellXml.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, '<w:t></w:t>');
  }
  const escaped = value.replace(/&/g, '&amp;');
  // If there's already a w:t, replace it; otherwise inject before </w:r>
  if (/<w:t/.test(cellXml)) {
    return cellXml.replace(/<w:t[^>]*>[^<]*<\/w:t>/g, `<w:t>${escaped}</w:t>`);
  }
  // No run exists — inject a run with w:t
  return cellXml.replace('</w:p>', `<w:r><w:t>${escaped}</w:t></w:r></w:p>`);
}

// Update ANNEXURE E table rows 1-3 (0-indexed, row 0 is header)
// For each data row, set Yes/No/blank in the Yes column (col index 1) and No column (col index 2)
function updateAnnexureETable(docXml, guestsLocalId, unmarriedCouple, foreigner) {
  const selections = { guestsLocalId, unmarriedCouple, foreigner };

  // Find the ANNEXURE E table by locating 'Guests with Local ID'
  const anchorIdx = docXml.indexOf('Guests with Local ID');
  if (anchorIdx === -1) return docXml;

  const tableStart = docXml.lastIndexOf('<w:tbl>', anchorIdx);
  const foreignerIdx = docXml.indexOf('Foreigner', anchorIdx);
  const tableEnd = docXml.indexOf('</w:tbl>', foreignerIdx) + '</w:tbl>'.length;

  const beforeTable = docXml.substring(0, tableStart);
  const afterTable = docXml.substring(tableEnd);
  let tableXml = docXml.substring(tableStart, tableEnd);

  // Split into rows
  const rowRegex = /<w:tr[ >][\s\S]*?<\/w:tr>/g;
  const rowMatches = [...tableXml.matchAll(rowRegex)];

  if (rowMatches.length < 4) return docXml; // Safety check

  // Process rows 1, 2, 3 (data rows)
  const dataRows = [
    { rowIndex: 1, selection: guestsLocalId },
    { rowIndex: 2, selection: unmarriedCouple },
    { rowIndex: 3, selection: foreigner },
  ];

  for (const { rowIndex, selection } of dataRows) {
    const rowMatch = rowMatches[rowIndex];
    if (!rowMatch) continue;

    const rowXml = rowMatch[0];
    // Split row into cells
    const cellRegex = /<w:tc>[\s\S]*?<\/w:tc>/g;
    const cells = [...rowXml.matchAll(cellRegex)];
    if (cells.length < 3) continue;

    const yesVal = selection === 'Yes' ? 'Yes' : '';
    const noVal = selection === 'No' ? 'No' : '';

    // Update Yes cell (index 1) and No cell (index 2)
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
      DATE, GST_NO, GST_ADDRESS, OWNER_NAME, TRADE_NAME,
      HOTEL_NAME, HOTEL_ADDRESS, INVENTRY,
      LAT_D, LONG_D, MOBILE_NO, LANDLINE_NO, EMAIL_ADD,
      GUESTS_LOCAL_ID, UNMARRIED_COUPLE, FOREIGNER
    } = req.body;

    const formattedDate = formatDate(DATE);

    const content = fs.readFileSync(TEMPLATE_PATH, 'binary');
    const zip = new PizZip(content);

    let docXml = zip.file('word/document.xml').asText();

    // Replace all plain text tokens
    const replacements = {
      'DATE': formattedDate,
      'GST_NO': GST_NO || '',
      'GST_ADDRESS': GST_ADDRESS || '',
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

    // Replace Brand Display photo if uploaded (rId10 = media/image3.png)
    if (req.file) {
      const uploadedImageBuffer = req.file.buffer;
      const uploadedMime = req.file.mimetype; // e.g. image/jpeg, image/png

      // Determine extension
      const extMap = { 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/png': 'png', 'image/webp': 'webp' };
      const ext = extMap[uploadedMime] || 'png';

      // rId10 points to media/image3.png — replace the file content
      // Always overwrite image3.png regardless of format (Word will still render it)
      zip.file('word/media/image3.png', uploadedImageBuffer, { binary: true });

      // If it's not PNG, update the relationship to reflect the new content type
      // For simplicity, Word is lenient about file extension vs content, so overwriting works fine
    }

    // Process footers too
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
    const filename = ('Contract_' + safeHotelName + '_' + (formattedDate || 'draft')).replace(/\s+/g, '_') + '.docx';

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
