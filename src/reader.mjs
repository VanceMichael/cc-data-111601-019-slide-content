import { inflateRawSync } from 'node:zlib';
import { parseXml, local, descendants, content } from './xml.mjs';
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
// 部件对应的关系部件路径；包根的关系部件为 _rels/.rels
function relsPartName(part) {
  if (part === '') return '_rels/.rels';
  const at = part.lastIndexOf('/');
  return `${part.slice(0, at + 1)}_rels/${part.slice(at + 1)}.rels`;
}
// 关系 Target 按声明部件的相对 URI 解析为包内部件路径，不允许逃逸出包根
function resolveTarget(base, target) {
  const raw = target.startsWith('/') ? target.slice(1) : base + target;
  const segments = [];
  for (const segment of raw.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (segments.length === 0) throw new Error('pptx_target_escape');
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join('/');
}
// 读取部件的 package relationships：沿 relationship Type 末段判角色，Target 解析为包内路径
function relationships(entries, part) {
  const name = relsPartName(part);
  if (!entries.has(name)) return [];
  const root = parseXml(entries.get(name));
  if (local(root) !== 'Relationships') throw new Error('pptx_rels_root');
  const base = part === '' ? '' : part.slice(0, part.lastIndexOf('/') + 1);
  return root.children.filter(node => local(node) === 'Relationship')
    .filter(node => node.attrs.TargetMode !== 'External')
    .map(node => ({ id: node.attrs.Id ?? '', role: (node.attrs.Type ?? '').split('/').at(-1),
      target: resolveTarget(base, node.attrs.Target ?? '') }));
}
const paragraphsOf = root => descendants(root, 'p')
  .map(paragraph => descendants(paragraph, 't').map(content).join(''));
// 拆分形状文字：占位形状的文字是版式/母版的提示定义，非占位形状是固定显示文字
function shapesOf(root) {
  const fixed = [];
  const prompts = [];
  for (const sp of descendants(root, 'sp')) {
    const paragraphs = paragraphsOf(sp);
    const ph = descendants(sp, 'ph')[0];
    if (!ph) { fixed.push(...paragraphs); continue; }
    const placeholder = { type: ph.attrs.type ?? 'body' };
    if (ph.attrs.idx !== undefined) placeholder.idx = ph.attrs.idx;
    for (const text of paragraphs) prompts.push({ text, placeholder });
  }
  return { fixed, prompts };
}
// 采集单页：页面自有文字、经实际关系链解析的版式/母版固定形状、讲者注释分层记录
function slideEntry(entries, parsePart, part, position) {
  const relations = {};
  const missing = [];
  const entry = { part, position, visible: null, paragraphs: [], texts: [],
    placeholderDefinitions: [], relations, missing };
  if (!entries.has(part)) { missing.push(part); return entry; }
  const root = parsePart(part);
  if (local(root) !== 'sld') throw new Error('pptx_slide_root');
  entry.visible = root.attrs.show !== '0';
  entry.paragraphs = paragraphsOf(root);
  for (const text of entry.paragraphs)
    entry.texts.push({ text, audience: 'learner', origin: 'slide', part });
  const rels = relationships(entries, part);
  const layoutRel = rels.find(rel => rel.role === 'slideLayout');
  const notesRel = rels.find(rel => rel.role === 'notesSlide');
  let layoutRoot = null;
  if (layoutRel) {
    relations.layout = layoutRel.target;
    if (entries.has(layoutRel.target)) layoutRoot = parsePart(layoutRel.target);
    else missing.push(layoutRel.target);
  }
  let masterRoot = null;
  if (layoutRoot) {
    if (local(layoutRoot) !== 'sldLayout') throw new Error('pptx_layout_root');
    const masterRel = relationships(entries, relations.layout).find(rel => rel.role === 'slideMaster');
    if (masterRel) {
      relations.master = masterRel.target;
      if (entries.has(masterRel.target)) masterRoot = parsePart(masterRel.target);
      else missing.push(masterRel.target);
    }
  }
  if (layoutRoot) {
    const { fixed, prompts } = shapesOf(layoutRoot);
    for (const text of fixed)
      entry.texts.push({ text, audience: 'learner', origin: 'layout', part: relations.layout });
    for (const prompt of prompts)
      entry.placeholderDefinitions.push({ ...prompt, origin: 'layout', part: relations.layout });
  }
  if (masterRoot) {
    if (local(masterRoot) !== 'sldMaster') throw new Error('pptx_master_root');
    // 页面或版式关闭母版显示（showMasterSp="0"）时，母版固定形状不出现在该页
    const shown = root.attrs.showMasterSp !== '0' && layoutRoot.attrs.showMasterSp !== '0';
    const { fixed, prompts } = shapesOf(masterRoot);
    if (shown) {
      for (const text of fixed)
        entry.texts.push({ text, audience: 'learner', origin: 'master', part: relations.master });
    }
    for (const prompt of prompts)
      entry.placeholderDefinitions.push({ ...prompt, origin: 'master', part: relations.master });
  }
  if (notesRel) {
    relations.notes = notesRel.target;
    if (entries.has(notesRel.target)) {
      const notesRoot = parsePart(notesRel.target);
      if (local(notesRoot) !== 'notes') throw new Error('pptx_notes_root');
      for (const text of paragraphsOf(notesRoot))
        entry.texts.push({ text, audience: 'speaker', origin: 'notes', part: notesRel.target });
    } else missing.push(notesRel.target);
  }
  return entry;
}
export function inspect(bytes) {
  const entries = zipEntries(bytes);
  if (!entries.has('[Content_Types].xml')) throw new Error('pptx_content_types');
  const documentRel = relationships(entries, '').find(rel => rel.role === 'officeDocument');
  const presentationPart = documentRel?.target ?? 'ppt/presentation.xml';
  if (!entries.has(presentationPart)) throw new Error('pptx_presentation_missing');
  const presentation = parseXml(entries.get(presentationPart));
  if (local(presentation) !== 'presentation') throw new Error('pptx_presentation_root');
  const presRels = relationships(entries, presentationPart);
  const ordered = descendants(presentation, 'sldId')
    .map(node => presRels.find(rel => rel.id === (node.attrs['r:id'] ?? '') && rel.role === 'slide')?.target)
    .filter(target => target !== undefined);
  const cache = new Map([[presentationPart, presentation]]);
  const parsePart = part => {
    if (!cache.has(part)) cache.set(part, parseXml(entries.get(part)));
    return cache.get(part);
  };
  const slides = ordered.map((part, index) => slideEntry(entries, parsePart, part, index + 1));
  return { kind: 'pptx-slide-text', parts: [...entries.keys()], slides };
}
