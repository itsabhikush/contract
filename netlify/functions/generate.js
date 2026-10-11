const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Busboy = require('busboy');

function getTemplatePath(contractType) {
  const filename = contractType === 'NGST' ? 'Form_NGST.docx' : 'Form_GST.docx';
  const candidates = [
    path.join(__dirname, filename),
    path.join(__dirname, '..', '..', filename),
    path.join(process.cwd(), 'netlify', 'functions', filename),
    path.join(process.cwd(), filename),
    path.join(process.cwd(), contractType === 'NGST' ? 'formet NGST.docx' : 'Form_GST.docx')
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return path.join(__dirname, filename);
}

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
  const escaped = value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  if (/<w:t[ >]/.test(cellXml))
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

// ── Format percentage value ─────────────────────────────────────────────────
function formatPercentage(val) {
  if (val === undefined || val === null) return '';
  let s = String(val).trim();
  if (!s) return '';
  if (!s.endsWith('%')) s += '%';
  return s;
}

// ── Render cell content paragraphs ─────────────────────────────────────────
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

// ── Build Platform fees table XML ───────────────────────────────────────────
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
        <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>Franchisor's Booking Occupancy (%)</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="${col2W}" w:type="dxa"/></w:tcPr>
        <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>Platform fees</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
    ${rowsXml}
  </w:tbl>`;
}

// ── Update Clause 7.1.b margins and optional platform fees table ────────────
function updateClause71b(docXml, franchisorMargin, franchiseeMargin, includePlatformFees, rows, isNGST) {
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

  // Spacing paragraph
  newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>`;

  // Franchisee margin paragraph
  newXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:spacing w:line="200" w:lineRule="atLeast"/><w:jc w:val="both"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">Reseller margin on Franchisee Bookings = ${escFranchisee}</w:t></w:r></w:p>`;

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
      contractType,
      DATE, GST_NO, GST_ADDRESS,
      PAN_NO, PAN_ADDRESS, ADHAR_ADDRESS,
      OWNER_NAME, TRADE_NAME,
      HOTEL_NAME, HOTEL_ADDRESS, INVENTRY,
      LAT_D, LONG_D, MOBILE_NO, LANDLINE_NO, EMAIL_ADD,
      GUESTS_LOCAL_ID, UNMARRIED_COUPLE, FOREIGNER,
      RESELLER_MARGIN_FRANCHISOR, RESELLER_MARGIN_FRANCHISEE,
      INCLUDE_PLATFORM_FEES, PLATFORM_FEES_DATA
    } = fields;

    const isNGST = contractType === 'NGST';
    const templatePath = getTemplatePath(isNGST ? 'NGST' : 'GST');

    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template not found at ${templatePath}`);
    }

    const formattedDate = formatDate(DATE);
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);

    let docXml = zip.file('word/document.xml').asText();

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

    docXml = updateClause71b(
      docXml,
      RESELLER_MARGIN_FRANCHISOR !== undefined && RESELLER_MARGIN_FRANCHISOR !== '' ? RESELLER_MARGIN_FRANCHISOR : '25%',
      RESELLER_MARGIN_FRANCHISEE !== undefined && RESELLER_MARGIN_FRANCHISEE !== '' ? RESELLER_MARGIN_FRANCHISEE : '0%',
      shouldIncludeTable,
      parsedTableRows,
      isNGST
    );
    zip.file('word/document.xml', docXml);

    // Footer replacements
    Object.keys(zip.files).filter(f => f.startsWith('word/footer')).forEach(fFile => {
      let fXml = zip.file(fFile).asText();
      for (const [token, value] of Object.entries(replacements)) {
        fXml = replaceInXml(fXml, token, value);
      }
      zip.file(fFile, fXml);
    });

    // Replace hotel photo if uploaded (Annexure C-1: media/image3.png)
    if (files.hotelPhoto && files.hotelPhoto.buffer && files.hotelPhoto.buffer.length > 0) {
      zip.file('word/media/image3.png', files.hotelPhoto.buffer, { binary: true });
    }

    const output = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const safeHotel = (HOTEL_NAME || 'Contract').replace(/[^a-zA-Z0-9_\- ]/g, '').trim();
    const typeLabel = isNGST ? 'NGST' : 'GST';
    const filename  = `Contract_${typeLabel}_${safeHotel}_${formattedDate || 'draft'}`.replace(/\s+/g,'_') + '.docx';

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
