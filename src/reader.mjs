import { inflateRawSync } from 'node:zlib';
export function zipEntries(bytes) {
  let end = -1;
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 65557); at--) {
    if (bytes.readUInt32LE(at) === 0x06054b50) { end = at; break; }
  }
  if (end < 0) throw new Error('zip_end_missing');
  const count = bytes.readUInt16LE(end + 10);
  const directorySize = bytes.readUInt32LE(end + 12);
  let offset = bytes.readUInt32LE(end + 16);
  if (offset + directorySize > end || count > 1000) throw new Error('zip_directory_bounds');
  const entries = new Map(); let expanded = 0;
  for (let index = 0; index < count; index++) {
    if (offset + 46 > bytes.length || bytes.readUInt32LE(offset) !== 0x02014b50) throw new Error('zip_directory_record');
    const method = bytes.readUInt16LE(offset + 10); const compressed = bytes.readUInt32LE(offset + 20);
    const uncompressed = bytes.readUInt32LE(offset + 24); const nameLength = bytes.readUInt16LE(offset + 28);
    const extraLength = bytes.readUInt16LE(offset + 30); const commentLength = bytes.readUInt16LE(offset + 32);
    const localOffset = bytes.readUInt32LE(offset + 42);
    const name = bytes.toString('utf8', offset + 46, offset + 46 + nameLength);
    if (name.startsWith('/') || name.split('/').includes('..') || entries.has(name)) throw new Error('zip_member_name');
    if (localOffset + 30 > bytes.length || bytes.readUInt32LE(localOffset) !== 0x04034b50) throw new Error('zip_local_header');
    const body = localOffset + 30 + bytes.readUInt16LE(localOffset + 26) + bytes.readUInt16LE(localOffset + 28);
    if (body + compressed > bytes.length || uncompressed > 4194304) throw new Error('zip_member_bounds');
    const raw = bytes.subarray(body, body + compressed);
    const data = method === 0 ? raw : method === 8 ? inflateRawSync(raw, { maxOutputLength: 4194304 }) : null;
    if (!data || data.length !== uncompressed) throw new Error('zip_member_size');
    expanded += data.length;
    if (expanded > 8388608) throw new Error('zip_expanded_limit');
    entries.set(name, data);
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}
import { parseXml, local, descendants, content } from './xml.mjs';
export function inspect(bytes) {
  const entries = zipEntries(bytes);
  if (!entries.has('[Content_Types].xml')) throw new Error('pptx_content_types');
  const slides = [];
  for (const [name, data] of entries) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
    const root = parseXml(data);
    if (local(root) !== 'sld') throw new Error('pptx_slide_root');
    slides.push({ part: name, visible: root.attrs.show !== '0',
      paragraphs: descendants(root, 'p').map(paragraph => descendants(paragraph, 't').map(content).join('')) });
  }
  return { kind: 'pptx-slide-text', parts: [...entries.keys()], slides };
}
