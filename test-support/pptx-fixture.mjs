import { crc32, deflateRawSync } from 'node:zlib';
import { posix } from 'node:path';

// Minimal ZIP writer used only by tests to construct PPTX-like packages with
// controlled parts, relationships and display switches.
export function buildZip(files) {
  const locals = [];
  const central = [];
  let offset = 0;
  for (const [name, value] of Object.entries(files)) {
    const data = Buffer.isBuffer(value) ? value : Buffer.from(value, 'utf8');
    const nameBytes = Buffer.from(name, 'utf8');
    const compressed = deflateRawSync(data);
    const useDeflate = compressed.length < data.length;
    const body = useDeflate ? compressed : data;
    const method = useDeflate ? 8 : 0;
    const checksum = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(method, 8); local.writeUInt16LE(0, 10); local.writeUInt16LE(0, 12);
    local.writeUInt32LE(checksum, 14); local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(data.length, 22); local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBytes, body);
    const record = Buffer.alloc(46);
    record.writeUInt32LE(0x02014b50, 0); record.writeUInt16LE(20, 4); record.writeUInt16LE(20, 6);
    record.writeUInt16LE(0x0800, 8); record.writeUInt16LE(method, 10); record.writeUInt16LE(0, 12);
    record.writeUInt16LE(0, 14); record.writeUInt32LE(checksum, 16); record.writeUInt32LE(body.length, 20);
    record.writeUInt32LE(data.length, 24); record.writeUInt16LE(nameBytes.length, 28);
    record.writeUInt16LE(0, 30); record.writeUInt16LE(0, 32); record.writeUInt16LE(0, 34);
    record.writeUInt16LE(0, 36); record.writeUInt32LE(0, 38); record.writeUInt32LE(offset, 42);
    central.push(record, nameBytes);
    offset += local.length + nameBytes.length + body.length;
  }
  const directory = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(0, 4); end.writeUInt16LE(0, 6);
  end.writeUInt16LE(Object.keys(files).length, 8); end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(offset, 16); end.writeUInt16LE(0, 20);
  return Buffer.concat([...locals, directory, end]);
}

export const namespaces = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
export const rel = (id, role, target, extra = '') =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${role}" Target="${target}"${extra}/>`;
export const relationships = items => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.join('')}</Relationships>`;
export const paragraph = text => `<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`;
export const emptyParagraph = () => '<a:p/>';
export const shape = ({ id = '2', name = 'Shape', placeholder = '', paragraphs = [] } = {}) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${placeholder}</p:nvPr></p:nvSpPr>`
  + `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs.join('')}</p:txBody></p:sp>`;
export const slidePart = (shapes, attrs = '') =>
  `<p:sld ${namespaces}${attrs}><p:cSld><p:spTree>${shapes.join('')}</p:spTree></p:cSld></p:sld>`;
export const layoutPart = (shapes, attrs = '') =>
  `<p:sldLayout ${namespaces}${attrs}><p:cSld><p:spTree>${shapes.join('')}</p:spTree></p:cSld></p:sldLayout>`;
export const masterPart = shapes =>
  `<p:sldMaster ${namespaces}><p:cSld><p:spTree>${shapes.join('')}</p:spTree></p:cSld></p:sldMaster>`;
export const notesPart = shapes =>
  `<p:notes ${namespaces}><p:cSld><p:spTree>${shapes.join('')}</p:spTree></p:cSld></p:notes>`;
export const presentationPart = slides =>
  `<p:presentation ${namespaces}><p:sldIdLst>${slides.join('')}</p:sldIdLst></p:presentation>`;
export const slideId = (id, relationship) => `<p:sldId id="${id}" r:id="${relationship}"/>`;

export const relsPath = part => {
  const directory = posix.dirname(part);
  return posix.join(directory === '.' ? '' : directory, '_rels', `${posix.basename(part)}.rels`);
};

// Builds a complete synthetic package. Every part body is supplied explicitly
// so a test can omit a relationship target without another same-named part
// standing in for it.
export function buildPackage({ slides, layouts = {}, masters = {}, notes = {},
  presentation = 'ppt/presentation.xml', contentTypes = true }) {
  const directory = posix.dirname(presentation);
  const target = part => (part.startsWith('/') ? part.slice(1) : posix.join(directory, part));
  const parts = {
    '_rels/.rels': relationships([rel('rId1', 'officeDocument', presentation)])
  };
  if (contentTypes) {
    parts['[Content_Types].xml'] =
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';
  }
  parts[presentation] = presentationPart(slides.map((slide, index) =>
    slideId(slide.id ?? String(256 + index), slide.relationship)));
  parts[relsPath(presentation)] = relationships(slides.map(slide =>
    rel(slide.relationship, 'slide', slide.target)));
  for (const [part, { body, rels = [] }] of Object.entries({ ...layouts, ...masters, ...notes })) {
    parts[part] = body;
    if (rels.length > 0) parts[relsPath(part)] = relationships(rels);
  }
  for (const slide of slides) {
    if (slide.body) parts[target(slide.target)] = slide.body;
    if (slide.rels?.length) parts[relsPath(target(slide.target))] = relationships(slide.rels);
  }
  return buildZip(parts);
}
