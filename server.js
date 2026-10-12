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

// Format percentage value: ensure trailing % if non-empty
function formatPercentage(val) {
  if (val === undefined || val === null) return '';
  let s = String(val).trim();
  if (!s) return '';
  if (!s.endsWith('%')) s += '%';
  return s;
}

// Render cell content paragraphs handling multiple lines
function renderCellContent(text) {
  if (!text) {
    return `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>`;
  }
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length === 0) lines.push('');
  return lines.map(line => {
    const escaped = line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>${escaped}</w:t></w:r></w:p>`;
  }).join('');
}

// Build Platform fees table XML
function buildPlatformFeesTableXml(rows, isNGST) {
  const col1W = isNGST ? 3386 : 2977;
  const col2W = isNGST ? 3452 : 3402;
  const tblIndW = isNGST ? 720 : 1413;

  let rowsXml = '';
  for (const r of rows) {
    rowsXml += `<w:tr>
      <w:trPr><w:trHeight w:val="697"/></w:trPr>
      <w:tc>
        <w:tcPr><w:tcW w:w="${col1W}" w:type="dxa"/></w:tcPr>
        ${renderCellContent(r.occupancy)}
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="${col2W}" w:type="dxa"/></w:tcPr>
        ${renderCellContent(r.fee)}
      </w:tc>
    </w:tr>`;
  }

  return `<w:tbl>
    <w:tblPr>
      <w:tblStyle w:val="TableGrid"/>
      <w:tblW w:w="0" w:type="auto"/>
      <w:tblInd w:w="${tblIndW}" w:type="dxa"/>
      <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
    </w:tblPr>
    <w:tblGrid>
      <w:gridCol w:w="${col1W}"/>
      <w:gridCol w:w="${col2W}"/>
    </w:tblGrid>
    <w:tr>
      <w:trPr><w:trHeight w:val="841"/></w:trPr>
      <w:tc>
        <w:tcPr><w:tcW w:w="${col1W}" w:type="dxa"/></w:tcPr>
        <w:p><w:pPr><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>Franchisor's Booking</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>Occupancy (%)</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="${col2W}" w:type="dxa"/></w:tcPr>
        <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>Platform fees</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
    ${rowsXml}
  </w:tbl>`;
}

// Update Clause 7.1.b margins and optional platform fees table
function updateClause71b(docXml, franchisorMargin, franchiseeMargin, includePlatformFees, rows, isNGST, includeOtaRider, otaRiderVal) {
  const formulaAnchor = 'Reseller margin % * Net Booking Value';
  const formulaIdx = docXml.indexOf(formulaAnchor);
  if (formulaIdx === -1) return docXml;

  const pFormEnd = docXml.indexOf('</w:p>', formulaIdx) + '</w:p>'.length;
  const tblIdx = docXml.indexOf('<w:tbl>', pFormEnd);
  if (tblIdx === -1) return docXml;

  const tblEnd = docXml.indexOf('</w:tbl>', tblIdx) + '</w:tbl>'.length;

  const escFranchisor = (formatPercentage(franchisorMargin) || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escFranchisee = (formatPercentage(franchiseeMargin) || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  let newXml = '';
  // Spacing paragraph
  newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>`;

  // Franchisor margin paragraph
  newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">Reseller margin  on Franchisor Bookings = ${escFranchisor}</w:t></w:r></w:p>`;

  // Optional OTA Rider line (added after Franchisor Bookings)
  if (includeOtaRider && otaRiderVal) {
    const escOtaRider = (formatPercentage(otaRiderVal) || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>`;
    newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">OTA Rider = ${escOtaRider}</w:t></w:r></w:p>`;
  }

  // Spacing paragraph
  newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>`;

  // Franchisee margin paragraph (matching uploaded format: "Reseller margin on Franchisee Bookings= Y%")
  newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">Reseller margin on Franchisee Bookings= ${escFranchisee}</w:t></w:r></w:p>`;

  if (includePlatformFees && rows && rows.length > 0) {
    // Spacing paragraph
    newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>`;

    // Heading "Platform fees:"
    newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>Platform fees:</w:t></w:r></w:p>`;

    // Spacing paragraph
    newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>`;

    // Table
    newXml += buildPlatformFeesTableXml(rows, isNGST);
  }

  return docXml.substring(0, pFormEnd) + newXml + docXml.substring(tblEnd);
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
      GUESTS_LOCAL_ID, UNMARRIED_COUPLE, FOREIGNER,
      RESELLER_MARGIN_FRANCHISOR, RESELLER_MARGIN_FRANCHISEE,
      INCLUDE_OTA_RIDER, OTA_RIDER,
      INCLUDE_PLATFORM_FEES, PLATFORM_FEES_DATA
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

    // Update Clause 7.1.b (Margins and Platform fees table)
    let parsedTableRows = [];
    if (PLATFORM_FEES_DATA) {
      try {
        parsedTableRows = typeof PLATFORM_FEES_DATA === 'string' ? JSON.parse(PLATFORM_FEES_DATA) : PLATFORM_FEES_DATA;
      } catch (e) {
        parsedTableRows = [];
      }
    }
    const shouldIncludeTable = INCLUDE_PLATFORM_FEES === 'true' || INCLUDE_PLATFORM_FEES === 'Yes' || INCLUDE_PLATFORM_FEES === true;
    const shouldIncludeOtaRider = INCLUDE_OTA_RIDER === 'true' || INCLUDE_OTA_RIDER === 'Yes' || INCLUDE_OTA_RIDER === true;

    docXml = updateClause71b(
      docXml,
      RESELLER_MARGIN_FRANCHISOR !== undefined && RESELLER_MARGIN_FRANCHISOR !== '' ? RESELLER_MARGIN_FRANCHISOR : '25%',
      RESELLER_MARGIN_FRANCHISEE !== undefined && RESELLER_MARGIN_FRANCHISEE !== '' ? RESELLER_MARGIN_FRANCHISEE : '0%',
      shouldIncludeTable,
      parsedTableRows,
      isNGST,
      shouldIncludeOtaRider,
      OTA_RIDER
    );

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
