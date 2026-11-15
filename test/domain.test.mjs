import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { inspect, zipEntries } from '../src/reader.mjs';
const primary=await readFile(new URL('../fixtures/primary.bin',import.meta.url));
const manifest=JSON.parse(await readFile(new URL('../fixtures/upstream.json',import.meta.url),'utf8'));
async function resource(path) {
  const item=manifest.resources[path];
  return item.file ? readFile(new URL(`../fixtures/${item.file}`,import.meta.url)) : Buffer.from(item.body);
}

test('页面自有文字与显式空段落可区分',()=>{
  const result=inspect(primary);assert.equal(result.slides.length,2);assert.equal(result.slides[0].paragraphs[0],'现场课程');
  assert.equal(result.slides[1].paragraphs[0],'');assert.equal(result.slides[1].visible,false);
  const entries=zipEntries(primary);assert.match(entries.get('ppt/slides/slide2.xml').toString(),/showMasterSp="0"/);
  assert.match(entries.get('ppt/slideMasters/slideMaster1.xml').toString(),/培训资料仅供内部使用/);
});
test('讲者部件与复用布局真实存在于包内',()=>{
  const entries=zipEntries(primary);assert.ok(entries.has('ppt/notesSlides/notesSlide1.xml'));
  assert.match(entries.get('ppt/slides/_rels/slide1.xml.rels').toString(),/slideLayout1/);
  assert.match(entries.get('ppt/slides/_rels/slide2.xml.rels').toString(),/slideLayout1/);
});
test('缺关系目标变体与损坏 ZIP 有不同表现',async()=>{
  const entries=zipEntries(await resource('/missing-rel'));
  assert.match(entries.get('ppt/slides/_rels/slide1.xml.rels').toString(),/slideLayout9/);
  assert.ok(!entries.has('ppt/slideLayouts/slideLayout9.xml'));assert.throws(()=>zipEntries(primary.subarray(0,100)));
});
