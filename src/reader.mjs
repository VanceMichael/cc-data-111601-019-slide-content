import { inflateRawSync } from 'node:zlib';
import { posix } from 'node:path';
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

// Part roles and page order come only from package relationships. A target is
// resolved against the part that owns the relationship. If the target part is
// absent we report it and stop: another part with a similar name is never used.
const children = (node, name) => node.children.filter(child => local(child) === name);
const child = (node, name) => children(node, name)[0] ?? null;
const firstDescendant = (node, name) => descendants(node, name)[0] ?? null;
const isOff = value => value === '0' || value === 'false';

function relationshipsPart(part) {
  if (part === null) return '_rels/.rels';
  const directory = posix.dirname(part);
  return posix.join(directory === '.' ? '' : directory, '_rels', `${posix.basename(part)}.rels`);
}

function resolvePart(sourcePart, target) {
  if (typeof target !== 'string' || target.trim() === '') return null;
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith('//')) return null;
  let path = target.split(/[?#]/)[0];
  try { path = decodeURIComponent(path); } catch { return null; }
  if (path.includes('\0')) return null;
  const base = sourcePart === null ? '' : posix.dirname(sourcePart);
  const resolved = path.startsWith('/') ? path.slice(1) : posix.join(base, path);
  const normalized = posix.normalize(resolved);
  if (!normalized || normalized === '.' || normalized.startsWith('../') || normalized.startsWith('/')) return null;
  return normalized;
}

function relationshipModels(root) {
  if (local(root) !== 'Relationships') throw new Error('pptx_relationships_root');
  return children(root, 'Relationship').map(node => ({
    id: node.attrs.Id ?? null,
    type: node.attrs.Type ?? null,
    role: typeof node.attrs.Type === 'string' ? node.attrs.Type.split('/').at(-1) : null,
    mode: node.attrs.TargetMode === 'External' ? 'External' : 'Internal',
    target: node.attrs.Target ?? null
  }));
}

function relationshipAttribute(node) {
  const entries = Object.entries(node.attrs)
    .filter(([name]) => name !== 'id' && (name === 'r:id' || name.endsWith(':id')));
  return entries[0]?.[1] ?? null;
}

// Text of a paragraph comes from its runs. Field elements (<a:fld>, such as
// slide numbers and dates) are computed presentation data, not authored text,
// and the project scope excludes footer field computation, so they are skipped.
function paragraphText(paragraph) {
  let text = '';
  for (const node of paragraph.children) {
    const name = local(node);
    if (name === 'fld') continue;
    if (name === 't') text += content(node);
    else text += paragraphText(node);
  }
  return text;
}

function shapeModels(root) {
  const tree = firstDescendant(root, 'spTree');
  if (!tree) return [];
  return children(tree, 'sp').map(shape => {
    const nonVisual = firstDescendant(shape, 'nvSpPr');
    const properties = nonVisual ? child(nonVisual, 'nvPr') : null;
    const placeholder = properties ? child(properties, 'ph') : null;
    const identity = nonVisual ? child(nonVisual, 'cNvPr') : null;
    const textBody = firstDescendant(shape, 'txBody');
    return {
      id: identity?.attrs.id ?? null,
      name: identity?.attrs.name ?? null,
      placeholder: placeholder !== null,
      placeholderType: placeholder?.attrs.type ?? null,
      placeholderIndex: placeholder?.attrs.idx ?? null,
      // Paragraphs are kept even when empty: an explicit empty txBody is still
      // explicit slide content and must not be filled from a placeholder prompt.
      paragraphs: textBody ? children(textBody, 'p').map(paragraphText) : []
    };
  });
}

function renderSegments({ shapes, include, role, audience, speaker, part, slide, slideId, slidePart,
  hidden, reason, reasonCode }) {
  const segments = [];
  for (const shape of shapes) {
    if (!include(shape)) continue;
    shape.paragraphs.forEach((text, paragraph) => {
      if (text.trim() === '') return;
      segments.push({
        slide,
        slideIndex: slide - 1,
        slideId,
        slidePart,
        hidden,
        text,
        speaker,
        audience,
        role,
        // `origin`/`source` are stable aliases for the part role that produced
        // this text, so consumers can group by slide/layout/master/notes.
        origin: role,
        source: role,
        part,
        reasonCode,
        reason,
        shapeId: shape.id,
        shapeName: shape.name,
        placeholder: shape.placeholder,
        paragraph
      });
    });
  }
  return segments;
}

function placeholderDefinition(part, role, shape, text, paragraph) {
  return {
    part,
    role,
    text,
    paragraph,
    shapeId: shape.id,
    shapeName: shape.name,
    placeholder: { type: shape.placeholderType, idx: shape.placeholderIndex },
    audience: null,
    includedAs: 'definition',
    meaning: 'placeholder prompt definition, not learner text'
  };
}

function describeSlide({ entries, roots, relationships, part, slideId, slide, definitions, issues }) {
  const result = {
    index: slide,
    number: slide,
    slideIndex: slide - 1,
    slideId,
    part,
    visible: true,
    hidden: false,
    // Kept for existing callers: every paragraph in the slide part, including
    // an explicit empty one, never filled from a layout or master prompt.
    paragraphs: [],
    slideText: [],
    masterText: [],
    notesText: [],
    learnerText: [],
    speakerText: [],
    layoutPart: null,
    masterPart: null,
    notesPart: null,
    masterShapesShown: false,
    missing: [],
    missingParts: [],
    unresolvedRelationships: [],
    segments: []
  };
  if (part === null) return result;

  const root = roots(part, 'sld', 'pptx_slide_root');
  result.visible = !isOff(root.attrs.show);
  result.hidden = !result.visible;

  const slideRelationships = relationships(part);
  const layoutRelationship = slideRelationships.find(item => item.role === 'slideLayout');
  const notesRelationship = slideRelationships.find(item => item.role === 'notesSlide');
  const layout = resolveRelationship(entries, part, 'slideLayout', layoutRelationship);
  // A slide without notes is ordinary; only a broken layout link is an issue.
  const notes = resolveRelationship(entries, part, 'notesSlide', notesRelationship);
  if (layout.status !== 'ok') {
    const issue = { kind: layout.status, slide, relationship: 'slideLayout',
      target: layout.target };
    result.missing.push(issue);
    issues.push(issue);
  }
  if (notes.status !== 'ok' && notes.status !== 'missing_relationship') {
    const issue = { kind: notes.status, slide, relationship: 'notesSlide',
      target: notes.target };
    result.missing.push(issue);
    issues.push(issue);
  }

  const ownShapes = shapeModels(root);
  result.paragraphs = ownShapes.flatMap(shape => shape.paragraphs);
  result.slideText = ownShapes.flatMap(shape => shape.paragraphs.filter(text => text.trim() !== ''));
  result.segments.push(...renderSegments({
    shapes: ownShapes,
    include: () => true,
    role: 'slide',
    audience: 'learner',
    speaker: false,
    part,
    slide,
    slideId,
    slidePart: part,
    hidden: result.hidden,
    reasonCode: 'slide_text',
    reason: 'explicit text stored in the slide part'
  }));

  if (layout.status === 'ok') {
    result.layoutPart = layout.part;
    const layoutRoot = roots(layout.part, 'sldLayout', 'pptx_layout_root');
    const layoutShapes = shapeModels(layoutRoot);
    for (const shape of layoutShapes) {
      if (!shape.placeholder) continue;
      shape.paragraphs.forEach((text, paragraph) => {
        if (text.trim() !== '') definitions.push(placeholderDefinition(
          layout.part, 'layout', shape, text, paragraph));
      });
    }

    const layoutRelationships = relationships(layout.part);
    const masterRelationship = layoutRelationships.find(item => item.role === 'slideMaster');
    const master = resolveRelationship(entries, layout.part, 'slideMaster', masterRelationship);
    if (master.status !== 'ok') {
      const issue = { kind: master.status, slide, relationship: 'slideMaster',
        target: master.target };
      result.missing.push(issue);
      issues.push(issue);
    }
    if (master.status === 'ok') {
      result.masterPart = master.part;
      const masterRoot = roots(master.part, 'sldMaster', 'pptx_master_root');
      const masterShapes = shapeModels(masterRoot);
      for (const shape of masterShapes) {
        if (!shape.placeholder) continue;
        shape.paragraphs.forEach((text, paragraph) => {
          if (text.trim() !== '') definitions.push(placeholderDefinition(
            master.part, 'master', shape, text, paragraph));
        });
      }
      // showMasterSp defaults to true. The slide and its layout can each turn
      // the master's ordinary shapes off for this page.
      result.masterShapesShown = !isOff(root.attrs.showMasterSp)
        && !isOff(layoutRoot.attrs.showMasterSp);
      if (result.masterShapesShown) {
        const masterSegments = renderSegments({
          shapes: masterShapes,
          // Placeholder text in a master is a prompt definition. Only the
          // master's ordinary shapes are displayed fixed text.
          include: shape => !shape.placeholder,
          role: 'master',
          audience: 'learner',
          speaker: false,
          part: master.part,
          slide,
          slideId,
          slidePart: part,
          hidden: result.hidden,
          reasonCode: 'master_shape_shown',
          reason: 'non-placeholder master shape displayed because showMasterSp is enabled on slide and layout'
        });
        result.segments.push(...masterSegments);
        result.masterText = masterSegments.map(segment => segment.text);
      }
    }
  }

  if (notes.status === 'ok') {
    result.notesPart = notes.part;
    const notesRoot = roots(notes.part, 'notes', 'pptx_notes_root');
    const notesSegments = renderSegments({
      shapes: shapeModels(notesRoot),
      include: () => true,
      role: 'notes',
      audience: 'speaker',
      speaker: true,
      part: notes.part,
      slide,
      slideId,
      slidePart: part,
      hidden: result.hidden,
      reasonCode: 'notes_text',
      reason: 'text from the notesSlide related to this slide'
    });
    result.segments.push(...notesSegments);
    result.notesText = notesSegments.map(segment => segment.text);
    result.speakerText = result.notesText;
  }

  result.learnerText = [...result.slideText, ...result.masterText];
  result.missingParts = result.missing.map(issue => issue.target).filter(Boolean);
  result.unresolvedRelationships = [...new Set(result.missing.map(issue => issue.relationship))];
  return result;
}

function resolveRelationship(entries, sourcePart, role, relationship) {
  if (!relationship) return { status: 'missing_relationship', role, target: null };
  if (relationship.mode === 'External') {
    return { status: 'external', role, target: relationship.target, relationship };
  }
  const part = resolvePart(sourcePart, relationship.target);
  if (!part || !entries.has(part)) {
    return { status: 'missing_part', role, target: part ?? relationship.target, relationship };
  }
  return { status: 'ok', role, part, relationship };
}

function uniquePlaceholders(definitions) {
  const seen = new Set();
  return definitions.filter(definition => {
    const key = JSON.stringify([definition.part, definition.shapeId, definition.paragraph,
      definition.placeholder.type, definition.placeholder.idx, definition.text]);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function inspect(bytes) {
  const entries = zipEntries(bytes);
  if (!entries.has('[Content_Types].xml')) throw new Error('pptx_content_types');

  const parsed = new Map();
  const roots = (part, expected, error) => {
    if (!parsed.has(part)) parsed.set(part, parseXml(entries.get(part)));
    const root = parsed.get(part);
    if (local(root) !== expected) throw new Error(error);
    return root;
  };
  const relCache = new Map();
  const relationships = part => {
    if (!relCache.has(part)) {
      const data = entries.get(relationshipsPart(part));
      relCache.set(part, data ? relationshipModels(parseXml(data)) : []);
    }
    return relCache.get(part);
  };

  const office = relationships(null).find(item => item.role === 'officeDocument'
    && item.mode !== 'External');
  if (!office) throw new Error('pptx_presentation_relationship');
  const presentationPart = resolvePart(null, office.target);
  if (!presentationPart || !entries.has(presentationPart)) throw new Error('pptx_presentation_part');
  const presentationRoot = roots(presentationPart, 'presentation', 'pptx_presentation_root');

  const slideList = firstDescendant(presentationRoot, 'sldIdLst');
  const slideReferences = slideList ? children(slideList, 'sldId') : [];
  const presentationRelationships = relationships(presentationPart);
  const definitions = [];
  const issues = [];
  const slides = [];

  slideReferences.forEach((reference, offset) => {
    const slide = offset + 1;
    const slideId = reference.attrs.id ?? null;
    const relationship = presentationRelationships.find(item =>
      item.id === relationshipAttribute(reference) && item.role === 'slide');
    const resolved = resolveRelationship(entries, presentationPart, 'slide', relationship);
    if (resolved.status !== 'ok') {
      const issue = { kind: resolved.status, slide, relationship: 'slide',
        target: resolved.target, slideId };
      issues.push(issue);
      const empty = describeSlide({ entries, roots, relationships, part: null, slideId, slide,
        definitions, issues });
      empty.missing.push(issue);
      empty.missingParts = issue.target ? [issue.target] : [];
      empty.unresolvedRelationships = ['slide'];
      slides.push(empty);
      return;
    }
    slides.push(describeSlide({ entries, roots, relationships,
      part: resolved.part, slideId, slide, definitions, issues }));
  });

  const segments = slides.flatMap(slide => slide.segments);
  // A layout or master can be reused by several pages. Record which pages
  // reference a prompt definition rather than copying it per page.
  const referencedBy = new Map();
  for (const slide of slides) {
    for (const part of [slide.layoutPart, slide.masterPart]) {
      if (!part) continue;
      if (!referencedBy.has(part)) referencedBy.set(part, []);
      referencedBy.get(part).push(slide.index);
    }
  }
  for (const definition of definitions) {
    definition.referencedBy = [...new Set(referencedBy.get(definition.part) ?? [])];
  }
  return {
    kind: 'pptx-slide-text',
    presentation: presentationPart,
    parts: [...entries.keys()],
    slides,
    segments,
    learner: segments.filter(segment => !segment.speaker),
    notes: segments.filter(segment => segment.speaker),
    hiddenSlides: slides.filter(slide => slide.hidden).map(slide => slide.index),
    placeholders: uniquePlaceholders(definitions),
    issues,
    unresolved: issues
  };
}
