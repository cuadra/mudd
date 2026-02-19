const XML_ESCAPES = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&apos;"
};

function xmlEscape(input) {
  return String(input).replace(/[&<>"']/g, (m) => XML_ESCAPES[m]);
}

function createCrc32Table() {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let j = 0; j < 8; j += 1) c = (c & 1 ? 0xedb88320 : 0) ^ (c >>> 1);
    table[i] = c >>> 0;
  }
  return table;
}

const CRC32_TABLE = createCrc32Table();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    c = CRC32_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function dosTimeParts(date = new Date()) {
  const year = Math.max(1980, date.getFullYear());
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = Math.floor(date.getSeconds() / 2);
  const time = (hours << 11) | (minutes << 5) | seconds;
  const dosDate = ((year - 1980) << 9) | (month << 5) | day;
  return { time, date: dosDate };
}

function writeUint16LE(target, value, offset) {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
}

function writeUint32LE(target, value, offset) {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
  target[offset + 2] = (value >>> 16) & 0xff;
  target[offset + 3] = (value >>> 24) & 0xff;
}

function zipStore(files) {
  const encoder = new TextEncoder();
  const dos = dosTimeParts();
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;

  for (const file of files) {
    const nameBytes = encoder.encode(file.name);
    const dataBytes = encoder.encode(file.content);
    const crc = crc32(dataBytes);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    writeUint32LE(localHeader, 0x04034b50, 0);
    writeUint16LE(localHeader, 20, 4);
    writeUint16LE(localHeader, 0, 6);
    writeUint16LE(localHeader, 0, 8);
    writeUint16LE(localHeader, dos.time, 10);
    writeUint16LE(localHeader, dos.date, 12);
    writeUint32LE(localHeader, crc, 14);
    writeUint32LE(localHeader, dataBytes.length, 18);
    writeUint32LE(localHeader, dataBytes.length, 22);
    writeUint16LE(localHeader, nameBytes.length, 26);
    writeUint16LE(localHeader, 0, 28);
    localHeader.set(nameBytes, 30);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    writeUint32LE(centralHeader, 0x02014b50, 0);
    writeUint16LE(centralHeader, 20, 4);
    writeUint16LE(centralHeader, 20, 6);
    writeUint16LE(centralHeader, 0, 8);
    writeUint16LE(centralHeader, 0, 10);
    writeUint16LE(centralHeader, dos.time, 12);
    writeUint16LE(centralHeader, dos.date, 14);
    writeUint32LE(centralHeader, crc, 16);
    writeUint32LE(centralHeader, dataBytes.length, 20);
    writeUint32LE(centralHeader, dataBytes.length, 24);
    writeUint16LE(centralHeader, nameBytes.length, 28);
    writeUint16LE(centralHeader, 0, 30);
    writeUint16LE(centralHeader, 0, 32);
    writeUint16LE(centralHeader, 0, 34);
    writeUint16LE(centralHeader, 0, 36);
    writeUint32LE(centralHeader, 0, 38);
    writeUint32LE(centralHeader, localOffset, 42);
    centralHeader.set(nameBytes, 46);

    localParts.push(localHeader, dataBytes);
    centralParts.push(centralHeader);
    localOffset += localHeader.length + dataBytes.length;
  }

  const centralLength = centralParts.reduce((sum, p) => sum + p.length, 0);
  const end = new Uint8Array(22);
  writeUint32LE(end, 0x06054b50, 0);
  writeUint16LE(end, 0, 4);
  writeUint16LE(end, 0, 6);
  writeUint16LE(end, files.length, 8);
  writeUint16LE(end, files.length, 10);
  writeUint32LE(end, centralLength, 12);
  writeUint32LE(end, localOffset, 16);
  writeUint16LE(end, 0, 20);

  return new Blob([...localParts, ...centralParts, end], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

export class TextRun {
  constructor({ text = "", bold = false, italic = false, underline = false } = {}) {
    this.text = text;
    this.bold = bold;
    this.italic = italic;
    this.underline = underline;
  }
}

export class ExternalHyperlink {
  constructor({ link, children = [] } = {}) {
    this.link = link || "";
    this.children = children;
  }
}

export class Paragraph {
  constructor({ headingLevel = null, numbering = null, children = [] } = {}) {
    this.headingLevel = headingLevel;
    this.numbering = numbering;
    this.children = children;
  }
}

export class TableCell {
  constructor({ children = [] } = {}) {
    this.children = children;
  }
}

export class TableRow {
  constructor({ children = [] } = {}) {
    this.children = children;
  }
}

export class Table {
  constructor({ rows = [] } = {}) {
    this.rows = rows;
  }
}

export class Document {
  constructor({ title = "Document", children = [] } = {}) {
    this.title = title;
    this.children = children;
  }
}

function paragraphPropsXml(paragraph) {
  const props = [];
  if (paragraph.headingLevel) {
    props.push(`<w:pStyle w:val="Heading${paragraph.headingLevel}"/>`);
  }
  if (paragraph.numbering) {
    props.push(
      `<w:numPr><w:ilvl w:val="${paragraph.numbering.level || 0}"/><w:numId w:val="${paragraph.numbering.numId}"/></w:numPr>`
    );
  }
  return props.length ? `<w:pPr>${props.join("")}</w:pPr>` : "";
}

function textRunXml(run) {
  const runProps = [];
  if (run.bold) runProps.push("<w:b/>");
  if (run.italic) runProps.push("<w:i/>");
  if (run.underline) runProps.push('<w:u w:val="single"/>');
  const props = runProps.length ? `<w:rPr>${runProps.join("")}</w:rPr>` : "";
  const preserve = /^[\s]|[\s]$/.test(run.text) ? ' xml:space="preserve"' : "";
  return `<w:r>${props}<w:t${preserve}>${xmlEscape(run.text)}</w:t></w:r>`;
}

function paragraphXml(paragraph, linkState) {
  const chunks = [paragraphPropsXml(paragraph)];
  for (const child of paragraph.children) {
    if (child instanceof ExternalHyperlink) {
      if (!child.link) continue;
      const relId = `rId${linkState.nextId++}`;
      linkState.rels.push({
        id: relId,
        target: child.link
      });
      const linkRuns = child.children.map((r) => textRunXml(r)).join("");
      chunks.push(`<w:hyperlink r:id="${relId}">${linkRuns}</w:hyperlink>`);
    } else {
      chunks.push(textRunXml(child));
    }
  }
  return `<w:p>${chunks.join("")}</w:p>`;
}

function tableXml(table, linkState) {
  const rows = table.rows
    .map((row) => {
      const cells = row.children
        .map((cell) => {
          const children = cell.children.length
            ? cell.children
            : [new Paragraph({ children: [new TextRun({ text: "" })] })];
          const inner = children.map((p) => paragraphXml(p, linkState)).join("");
          return `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
        })
        .join("");
      return `<w:tr>${cells}</w:tr>`;
    })
    .join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>${rows}</w:tbl>`;
}

function buildDocumentXml(documentModel, linkState) {
  const blocks = documentModel.children
    .map((child) => (child instanceof Table ? tableXml(child, linkState) : paragraphXml(child, linkState)))
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:w10="urn:schemas-microsoft-com:office:word"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 mc:Ignorable="w14 wp14">
 <w:body>
  ${blocks}
  <w:sectPr>
   <w:pgSz w:w="12240" w:h="15840"/>
   <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
   <w:cols w:space="720"/>
   <w:docGrid w:linePitch="360"/>
  </w:sectPr>
 </w:body>
</w:document>`;
}

function buildDocumentRelsXml(linkState) {
  const base = [
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
  ];
  const links = linkState.rels.map(
    (rel) =>
      `<Relationship Id="${xmlEscape(rel.id)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(rel.target)}" TargetMode="External"/>`
  );
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${base.concat(links).join("")}
</Relationships>`;
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/>
 </w:style>
 <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="20"/></w:rPr></w:style>
</w:styles>`;
}

function buildNumberingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:abstractNum w:abstractNumId="1">
  <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
 </w:abstractNum>
 <w:abstractNum w:abstractNumId="2">
  <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
 </w:abstractNum>
 <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
 <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;
}

function buildSettingsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:zoom w:percent="100"/>
</w:settings>`;
}

function buildRootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function buildContentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
 <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
 <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
 <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
 <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
 <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function buildCoreXml(title) {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <dc:title>${xmlEscape(title)}</dc:title>
 <dc:creator>ToDOCX Exporter</dc:creator>
 <cp:lastModifiedBy>ToDOCX Exporter</cp:lastModifiedBy>
 <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
 <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function buildAppXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
 <Application>ToDOCX Exporter</Application>
</Properties>`;
}

export class Packer {
  static toBlob(documentModel) {
    const linkState = { rels: [], nextId: 100 };
    const files = [
      { name: "[Content_Types].xml", content: buildContentTypesXml() },
      { name: "_rels/.rels", content: buildRootRelsXml() },
      { name: "docProps/core.xml", content: buildCoreXml(documentModel.title || "Document") },
      { name: "docProps/app.xml", content: buildAppXml() },
      { name: "word/document.xml", content: buildDocumentXml(documentModel, linkState) },
      { name: "word/_rels/document.xml.rels", content: buildDocumentRelsXml(linkState) },
      { name: "word/styles.xml", content: buildStylesXml() },
      { name: "word/numbering.xml", content: buildNumberingXml() },
      { name: "word/settings.xml", content: buildSettingsXml() }
    ];
    return zipStore(files);
  }
}
