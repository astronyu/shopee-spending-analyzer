const te = new TextEncoder();

function xmlEscape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function colName(index) {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function normalizeCell(value) {
  if (value && typeof value === 'object' && !Array.isArray(value) && value.__xlsxType) return value;
  if (typeof value === 'number' && Number.isFinite(value)) return { __xlsxType: 'number', value };
  if (typeof value === 'boolean') return { __xlsxType: 'boolean', value };
  return { __xlsxType: 'string', value: value ?? '' };
}

export function xlsxCurrency(value, style = 2) {
  return { __xlsxType: 'number', value: Number(value) || 0, style };
}

export function xlsxInteger(value, style = 3) {
  return { __xlsxType: 'number', value: Number(value) || 0, style };
}

export function xlsxLink(url, label = 'Open') {
  return { __xlsxType: 'link', value: url || '', label };
}

export function xlsxTitle(text) {
  return { __xlsxType: 'string', value: text, style: 4 };
}

export function xlsxSubtle(text) {
  return { __xlsxType: 'string', value: text, style: 5 };
}

function cellXml(value, rowIndex, colIndex, defaultStyle = 0) {
  const c = normalizeCell(value);
  const ref = `${colName(colIndex)}${rowIndex + 1}`;
  const style = c.style ?? defaultStyle;
  const styleAttr = style ? ` s="${style}"` : '';

  if (c.__xlsxType === 'number') {
    return `<c r="${ref}"${styleAttr}><v>${Number(c.value) || 0}</v></c>`;
  }
  if (c.__xlsxType === 'boolean') {
    return `<c r="${ref}" t="b"${styleAttr}><v>${c.value ? 1 : 0}</v></c>`;
  }
  if (c.__xlsxType === 'link' && c.value) {
    return `<c r="${ref}" s="6" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(c.label || c.value)}</t></is></c>`;
  }
  const text = xmlEscape(c.value);
  return `<c r="${ref}" t="inlineStr"${styleAttr}><is><t xml:space="preserve">${text}</t></is></c>`;
}

function sheetLinks(sheet) {
  const links = [];
  (sheet.rows || []).forEach((row, r) => row.forEach((cell, c) => {
    if (cell?.__xlsxType === 'link' && cell.value) {
      links.push({ ref: `${colName(c)}${r + 1}`, id: `rId${links.length + 1}`, url: cell.value });
    }
  }));
  return links;
}

function sheetXml(sheet, links) {
  const rows = sheet.rows || [];
  const widths = sheet.widths || [];
  const cols = widths.length
    ? `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${Math.max(6, Number(w) || 12)}" customWidth="1"/>`).join('')}</cols>`
    : '';
  const maxCols = rows.reduce((m, row) => Math.max(m, row.length), 0);
  const dim = rows.length && maxCols ? `A1:${colName(maxCols - 1)}${rows.length}` : 'A1';
  const headerRows = Number(sheet.headerRows ?? 1);
  const rowXml = rows.map((row, r) => {
    const defaultStyle = r < headerRows ? 1 : 0;
    const cells = row.map((v, c) => cellXml(v, r, c, defaultStyle)).join('');
    return `<row r="${r + 1}">${cells}</row>`;
  }).join('');
  const pane = sheet.freezeHeader && rows.length > 1
    ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${headerRows}" topLeftCell="A${headerRows + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
    : `<sheetViews><sheetView workbookViewId="0"/></sheetViews>`;
  const autoFilter = sheet.autoFilter && rows.length > headerRows && maxCols
    ? `<autoFilter ref="A${headerRows}:${colName(maxCols - 1)}${rows.length}"/>`
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="${dim}"/>
${pane}
<sheetFormatPr defaultRowHeight="15"/>
${cols}
<sheetData>${rowXml}</sheetData>
${autoFilter}
${links.length ? `<hyperlinks>${links.map(link => `<hyperlink ref="${link.ref}" r:id="${link.id}"/>`).join('')}</hyperlinks>` : ''}
</worksheet>`;
}

function sheetLinksXml(links) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${links.map(link => `<Relationship Id="${link.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(link.url)}" TargetMode="External"/>`).join('')}
</Relationships>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><name val="Calibri"/></font>
    <font><color rgb="FF777777"/><sz val="10"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEE4D2D"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="left"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
}

function workbookXml(sheets) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>
  <sheets>${sheets.map((s, i) => `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>
  <calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/>
</workbook>`;
}

function workbookRelsXml(sheets) {
  const sheetRels = sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetRels}
<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function contentTypesXml(sheets) {
  const overrides = sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${overrides}
</Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function u16(n) { return new Uint8Array([n & 255, (n >>> 8) & 255]); }
function u32(n) { return new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]); }
function concat(parts) {
  const size = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(size);
  let off = 0;
  for (const p of parts) { out.set(p, off); off += p.length; }
  return out;
}

function dosDateTime(date = new Date()) {
  const year = Math.max(1980, date.getFullYear());
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const day = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, day };
}

function zipStore(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  const dt = dosDateTime();

  for (const file of files) {
    const name = te.encode(file.name);
    const data = typeof file.data === 'string' ? te.encode(file.data) : file.data;
    const crc = crc32(data);
    const local = concat([
      u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(dt.time), u16(dt.day),
      u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), name, data
    ]);
    locals.push(local);
    const central = concat([
      u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(dt.time), u16(dt.day),
      u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0),
      u16(0), u16(0), u32(0), u32(offset), name
    ]);
    centrals.push(central);
    offset += local.length;
  }

  const centralBlob = concat(centrals);
  const localBlob = concat(locals);
  const eocd = concat([
    u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length),
    u32(centralBlob.length), u32(localBlob.length), u16(0)
  ]);
  return concat([localBlob, centralBlob, eocd]);
}

export function buildXlsx(sheets) {
  const files = [
    { name: '[Content_Types].xml', data: contentTypesXml(sheets) },
    { name: '_rels/.rels', data: rootRelsXml() },
    { name: 'xl/workbook.xml', data: workbookXml(sheets) },
    { name: 'xl/_rels/workbook.xml.rels', data: workbookRelsXml(sheets) },
    { name: 'xl/styles.xml', data: stylesXml() },
    ...sheets.flatMap((s, i) => {
      const links = sheetLinks(s);
      const parts = [{ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s, links) }];
      if (links.length) parts.push({ name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, data: sheetLinksXml(links) });
      return parts;
    })
  ];
  return new Blob([zipStore(files)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
