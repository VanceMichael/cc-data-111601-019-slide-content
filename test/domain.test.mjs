import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { inspect, zipEntries } from '../src/reader.mjs';
import { buildPackage, rel, shape, paragraph, emptyParagraph, slidePart, layoutPart,
  masterPart, notesPart } from '../test-support/pptx-fixture.mjs';

const primary = await readFile(new URL('../fixtures/primary.bin', import.meta.url));
const manifest = JSON.parse(await readFile(new URL('../fixtures/upstream.json', import.meta.url), 'utf8'));
async function resource(path) {
  const item = manifest.resources[path];
  return item.file ? readFile(new URL(`../fixtures/${item.file}`, import.meta.url)) : Buffer.from(item.body);
}

test('页面自有文字与显式空段落可区分', () => {
  const result = inspect(primary);
  assert.equal(result.slides.length, 2);
  assert.deepEqual(result.slides[0].slideText, ['现场课程']);
  assert.deepEqual(result.slides[1].paragraphs, ['']);
  assert.deepEqual(result.slides[1].slideText, []);
  assert.equal(result.slides[1].visible, false);
  assert.equal(result.slides[1].hidden, true);
  const entries = zipEntries(primary);
  assert.match(entries.get('ppt/slides/slide2.xml').toString(), /showMasterSp="0"/);
  assert.match(entries.get('ppt/slideMasters/slideMaster1.xml').toString(), /培训资料仅供内部使用/);
});

test('讲者部件与复用布局真实存在于包内', () => {
  const entries = zipEntries(primary);
  assert.ok(entries.has('ppt/notesSlides/notesSlide1.xml'));
  assert.match(entries.get('ppt/slides/_rels/slide1.xml.rels').toString(), /slideLayout1/);
  assert.match(entries.get('ppt/slides/_rels/slide2.xml.rels').toString(), /slideLayout1/);
});

test('母版普通形状按显示开关进入学员正文，讲者注释保持独立', () => {
  const result = inspect(primary);
  const first = result.slides[0];
  assert.deepEqual(first.learnerText, ['现场课程', '培训资料仅供内部使用']);
  assert.deepEqual(first.masterText, ['培训资料仅供内部使用']);
  assert.deepEqual(first.speakerText, ['仅讲者可见的补充说明']);
  assert.ok(!first.learnerText.includes('仅讲者可见的补充说明'));
  assert.ok(!first.learnerText.includes('培训标题提示'));
  const segments = result.segments;
  assert.equal(segments.length, 3);
  for (const segment of segments) {
    assert.equal(segment.slide, 1);
    assert.equal(segment.slidePart, 'ppt/slides/slide1.xml');
    assert.equal(typeof segment.reason, 'string');
    assert.ok(segment.reason.length > 0);
    assert.equal(segment.speaker, segment.role === 'notes');
  }
  const notes = segments.filter(segment => segment.speaker);
  assert.deepEqual(notes.map(segment => segment.text), ['仅讲者可见的补充说明']);
  assert.equal(notes[0].audience, 'speaker');
  const master = segments.filter(segment => segment.role === 'master');
  assert.equal(master[0].reasonCode, 'master_shape_shown');
  assert.ok(result.learner.every(segment => segment.speaker === false));
  assert.deepEqual(result.hiddenSlides, [2]);
});

test('showMasterSp 关闭时母版形状不加入该页，隐藏页保留并独立标识', () => {
  const result = inspect(primary);
  const hidden = result.slides[1];
  assert.equal(hidden.hidden, true);
  assert.equal(hidden.visible, false);
  assert.equal(hidden.masterPart, 'ppt/slideMasters/slideMaster1.xml');
  assert.equal(hidden.masterShapesShown, false);
  assert.deepEqual(hidden.masterText, []);
  assert.deepEqual(hidden.learnerText, []);
  assert.deepEqual(hidden.segments, []);
  // The hidden page is still listed and still records its layout provenance.
  assert.equal(hidden.layoutPart, 'ppt/slideLayouts/slideLayout1.xml');
});

test('版式与母版占位提示保留为定义而不是正文', () => {
  const result = inspect(primary);
  const layoutPrompts = result.placeholders.filter(item => item.role === 'layout');
  assert.deepEqual(layoutPrompts.map(item => item.text), ['培训标题提示']);
  assert.equal(layoutPrompts[0].includedAs, 'definition');
  assert.deepEqual(layoutPrompts[0].referencedBy, [1, 2]);
  assert.ok(!result.segments.some(segment => segment.text === '培训标题提示'));
});

test('复用布局只按关系引用，与同名部件无关', () => {
  const result = inspect(primary);
  assert.equal(result.slides[0].layoutPart, 'ppt/slideLayouts/slideLayout1.xml');
  assert.equal(result.slides[1].layoutPart, 'ppt/slideLayouts/slideLayout1.xml');
});

test('缺关系目标时报告缺失且不借同名布局补出内容', async () => {
  const bytes = await resource('/missing-rel');
  const entries = zipEntries(bytes);
  assert.match(entries.get('ppt/slides/_rels/slide1.xml.rels').toString(), /slideLayout9/);
  assert.ok(!entries.has('ppt/slideLayouts/slideLayout9.xml'));
  const result = inspect(bytes);
  const first = result.slides[0];
  assert.equal(first.layoutPart, null);
  assert.equal(first.masterPart, null);
  assert.deepEqual(first.masterText, []);
  assert.deepEqual(first.learnerText, ['现场课程']);
  assert.deepEqual(first.missing, [{ kind: 'missing_part', slide: 1,
    relationship: 'slideLayout', target: 'ppt/slideLayouts/slideLayout9.xml' }]);
  assert.deepEqual(result.issues, first.missing);
  // The unrelated layout1 that exists in the package is not used for slide 1.
  assert.ok(!result.segments.some(segment => segment.slide === 1 && segment.role === 'layout'));
  assert.ok(!result.segments.some(segment => segment.text === '培训资料仅供内部使用' && segment.slide === 1));
  assert.throws(() => zipEntries(primary.subarray(0, 100)));
});

test('缺关系目标变体与损坏 ZIP 有不同表现', async () => {
  const entries = zipEntries(await resource('/missing-rel'));
  assert.match(entries.get('ppt/slides/_rels/slide1.xml.rels').toString(), /slideLayout9/);
  assert.ok(!entries.has('ppt/slideLayouts/slideLayout9.xml'));
  assert.throws(() => zipEntries(primary.subarray(0, 100)));
});

const master = part => ({ body: masterPart([shape({ id: '10', name: 'Fixed',
  paragraphs: [paragraph(`${part} 固定说明`)] })]) });
const layout = (target = '../slideMasters/slideMaster1.xml', attrs = '') => ({
  body: layoutPart([shape({ name: 'Title', placeholder: '<p:ph type="title" idx="1"/>',
    paragraphs: [paragraph('版式提示')] })], attrs),
  rels: [rel('rMaster', 'slideMaster', target)]
});

test('版式 showMasterSp=false 关闭母版形状，slide 默认继承打开', () => {
  const bytes = buildPackage({
    slides: [{ relationship: 'rId1', target: 'slides/slide1.xml',
      rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
      body: slidePart([shape({ name: 'Body', paragraphs: [paragraph('页面文字')] })]) }],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout('../slideMasters/slideMaster1.xml', ' showMasterSp="0"') },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') }
  });
  const result = inspect(bytes);
  assert.equal(result.slides[0].masterShapesShown, false);
  assert.deepEqual(result.slides[0].learnerText, ['页面文字']);
  assert.equal(result.slides[0].masterPart, 'ppt/slideMasters/slideMaster1.xml');
});

test('slide 的 showMasterSp=false 只影响该页，复用的另一页仍显示母版文字', () => {
  const bytes = buildPackage({
    slides: [
      { relationship: 'rId1', target: 'slides/slide1.xml',
        rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
        body: slidePart([shape({ paragraphs: [paragraph('第一页')] })], ' showMasterSp="0"') },
      { relationship: 'rId2', target: 'slides/slide2.xml',
        rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
        body: slidePart([shape({ paragraphs: [paragraph('第二页')] })]) }
    ],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout() },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') }
  });
  const result = inspect(bytes);
  assert.deepEqual(result.slides[0].learnerText, ['第一页']);
  assert.deepEqual(result.slides[1].learnerText, ['第二页', '母版 固定说明']);
  const segments = result.segments.filter(segment => segment.role === 'master');
  assert.deepEqual(segments.map(segment => [segment.slide, segment.text]), [[2, '母版 固定说明']]);
});

test('母版占位提示与普通形状分开：提示不计正文，普通形状按页计入', () => {
  const bytes = buildPackage({
    slides: [{ relationship: 'rId1', target: 'slides/slide1.xml',
      rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
      body: slidePart([shape({ name: 'Title', placeholder: '<p:ph type="title" idx="1"/>',
        paragraphs: [paragraph('第一页')] })]) }],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': { body: layoutPart([],
      ''), rels: [rel('rMaster', 'slideMaster', '../slideMasters/slideMaster1.xml')] } },
    masters: { 'ppt/slideMasters/slideMaster1.xml': { body: masterPart([
      shape({ id: '10', name: 'Fixed', paragraphs: [paragraph('固定说明')] }),
      shape({ id: '11', name: 'Prompt', placeholder: '<p:ph type="body" idx="2"/>',
        paragraphs: [paragraph('母版提示词')] })
    ]) } }
  });
  const result = inspect(bytes);
  assert.deepEqual(result.slides[0].learnerText, ['第一页', '固定说明']);
  assert.deepEqual(result.placeholders.map(item => [item.role, item.text]), [['master', '母版提示词']]);
  assert.ok(!result.segments.some(segment => segment.text === '母版提示词'));
});

test('notes 缺失是普通情况，布局关系缺失才报告', () => {
  const bytes = buildPackage({
    slides: [{ relationship: 'rId1', target: 'slides/slide1.xml',
      rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
      body: slidePart([shape({ paragraphs: [paragraph('页面')] })]) }],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout() },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') }
  });
  const result = inspect(bytes);
  assert.deepEqual(result.issues, []);
  assert.equal(result.slides[0].notesPart, null);
  assert.deepEqual(result.notes, []);
});

test('讲者注释只属于关联的单页，不串到复用同一版式的其他页', () => {
  const notes = {
    'ppt/notesSlides/notesSlide1.xml': { body: notesPart([shape({ paragraphs: [paragraph('第一页讲稿')] })]) }
  };
  const bytes = buildPackage({
    slides: [
      { relationship: 'rId1', target: 'slides/slide1.xml',
        rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml'),
          rel('rNotes', 'notesSlide', '../notesSlides/notesSlide1.xml')],
        body: slidePart([shape({ paragraphs: [paragraph('第一页')] })]) },
      { relationship: 'rId2', target: 'slides/slide2.xml',
        rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
        body: slidePart([shape({ paragraphs: [paragraph('第二页')] })]) }
    ],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout() },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') },
    notes
  });
  const result = inspect(bytes);
  assert.deepEqual(result.slides[0].speakerText, ['第一页讲稿']);
  assert.deepEqual(result.slides[1].speakerText, []);
  assert.deepEqual(result.notes.map(segment => [segment.slide, segment.text]), [[1, '第一页讲稿']]);
  assert.deepEqual(result.slides[0].learnerText, ['第一页', '母版 固定说明']);
  assert.ok(!result.slides[1].learnerText.includes('第一页讲稿'));
});

test('页序取自 presentation 的 sldId 关系，不按部件文件名排序', () => {
  const bytes = buildPackage({
    slides: [
      { relationship: 'rId2', target: 'slides/slide2.xml',
        rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
        body: slidePart([shape({ paragraphs: [paragraph('先出现')] })]) },
      { relationship: 'rId1', target: 'slides/slide1.xml',
        rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
        body: slidePart([shape({ paragraphs: [paragraph('后出现')] })]) }
    ],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout() },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') }
  });
  const result = inspect(bytes);
  assert.deepEqual(result.slides.map(slide => slide.slideText), [['先出现'], ['后出现']]);
  assert.deepEqual(result.slides.map(slide => slide.part), ['ppt/slides/slide2.xml', 'ppt/slides/slide1.xml']);
  assert.deepEqual(result.slides.map(slide => slide.index), [1, 2]);
});

test('显式空段落不会从版式提示补文字', () => {
  const bytes = buildPackage({
    slides: [{ relationship: 'rId1', target: 'slides/slide1.xml',
      rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
      body: slidePart([shape({ name: 'Title', placeholder: '<p:ph type="title" idx="1"/>',
        paragraphs: [emptyParagraph()] })]) }],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout() },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') }
  });
  const result = inspect(bytes);
  assert.deepEqual(result.slides[0].paragraphs, ['']);
  assert.deepEqual(result.slides[0].slideText, []);
  assert.ok(!result.segments.some(segment => segment.text === '版式提示'));
  assert.deepEqual(result.placeholders.map(item => item.text), ['版式提示']);
});

test('外部或越界关系目标不被读取，按缺失处理', () => {
  const bytes = buildPackage({
    slides: [{ relationship: 'rId1', target: 'slides/slide1.xml',
      rels: [rel('rLayout', 'slideLayout', 'https://example.invalid/layout.xml', ' TargetMode="External"')],
      body: slidePart([shape({ paragraphs: [paragraph('页面')] })]) }]
  });
  const result = inspect(bytes);
  assert.equal(result.slides[0].layoutPart, null);
  assert.deepEqual(result.slides[0].learnerText, ['页面']);
  assert.equal(result.issues[0].kind, 'external');
});

test('关系目标越出包根时按缺失处理，不读取包外部件', () => {
  const bytes = buildPackage({
    slides: [{ relationship: 'rId1', target: 'slides/slide1.xml',
      rels: [rel('rLayout', 'slideLayout', '../../../../etc/layout.xml')],
      body: slidePart([shape({ paragraphs: [paragraph('页面')] })]) }]
  });
  const result = inspect(bytes);
  assert.equal(result.slides[0].layoutPart, null);
  assert.equal(result.issues[0].kind, 'missing_part');
  assert.deepEqual(result.slides[0].learnerText, ['页面']);
});

test('计算字段（页码/日期）不当作创作者文字收入', () => {
  const field = '<a:p><a:fld id="{1}" type="slidenum"><a:t>99</a:t></a:fld>'
    + '<a:r><a:t>真实文字</a:t></a:r></a:p>';
  const bytes = buildPackage({
    slides: [{ relationship: 'rId1', target: 'slides/slide1.xml',
      rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
      body: slidePart([shape({ paragraphs: [field] })]) }],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout() },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') }
  });
  const result = inspect(bytes);
  assert.deepEqual(result.slides[0].slideText, ['真实文字']);
});

test('presentation 指向的页部件缺失时保留页位并报告，不静默跳过', () => {
  const bytes = buildPackage({
    slides: [
      { relationship: 'rId1', target: 'slides/slide1.xml',
        rels: [rel('rLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml')],
        body: slidePart([shape({ paragraphs: [paragraph('第一页')] })]) },
      { relationship: 'rId2', target: 'slides/slide2.xml' }
    ],
    layouts: { 'ppt/slideLayouts/slideLayout1.xml': layout() },
    masters: { 'ppt/slideMasters/slideMaster1.xml': master('母版') }
  });
  const result = inspect(bytes);
  assert.deepEqual(result.slides.map(slide => slide.index), [1, 2]);
  assert.equal(result.slides[1].part, null);
  assert.deepEqual(result.slides[1].learnerText, []);
  assert.equal(result.issues.at(-1).kind, 'missing_part');
  assert.equal(result.issues.at(-1).target, 'ppt/slides/slide2.xml');
});
