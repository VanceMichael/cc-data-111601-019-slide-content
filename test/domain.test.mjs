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
test('页面文字、母版固定说明与讲者注释分层采集',()=>{
  const result=inspect(primary);const [first,second]=result.slides;
  assert.equal(first.position,1);assert.equal(first.visible,true);
  assert.deepEqual(first.texts,[
    {text:'现场课程',audience:'learner',origin:'slide',part:'ppt/slides/slide1.xml'},
    {text:'培训资料仅供内部使用',audience:'learner',origin:'master',part:'ppt/slideMasters/slideMaster1.xml'},
    {text:'仅讲者可见的补充说明',audience:'speaker',origin:'notes',part:'ppt/notesSlides/notesSlide1.xml'}]);
  assert.deepEqual(first.relations,{layout:'ppt/slideLayouts/slideLayout1.xml',
    master:'ppt/slideMasters/slideMaster1.xml',notes:'ppt/notesSlides/notesSlide1.xml'});
  assert.deepEqual(first.missing,[]);
  // 版式占位提示仅保留为定义，不作为该页正文
  assert.deepEqual(first.placeholderDefinitions,[{text:'培训标题提示',
    placeholder:{type:'title',idx:'1'},origin:'layout',part:'ppt/slideLayouts/slideLayout1.xml'}]);
  assert.ok(!first.texts.some(t=>t.text==='培训标题提示'));
  // 隐藏页保留并标识；showMasterSp="0" 关闭母版固定说明；空占位不被提示词填充
  assert.equal(second.visible,false);
  assert.deepEqual(second.texts,[{text:'',audience:'learner',origin:'slide',part:'ppt/slides/slide2.xml'}]);
  assert.deepEqual(second.placeholderDefinitions.map(d=>d.text),['培训标题提示']);
  assert.ok(!second.texts.some(t=>t.origin==='master'));
});
test('缺关系目标不借同名布局补全，其余页不受影响',async()=>{
  const result=inspect(await resource('/missing-rel'));const [first,second]=result.slides;
  assert.equal(first.relations.layout,'ppt/slideLayouts/slideLayout9.xml');
  assert.deepEqual(first.missing,['ppt/slideLayouts/slideLayout9.xml']);
  assert.deepEqual(first.placeholderDefinitions,[]);
  assert.ok(!first.texts.some(t=>t.text==='培训资料仅供内部使用'));
  assert.ok(!first.texts.some(t=>t.text==='培训标题提示'));
  assert.ok(first.texts.some(t=>t.audience==='speaker'&&t.text==='仅讲者可见的补充说明'));
  assert.deepEqual(second.missing,[]);
  assert.deepEqual(second.placeholderDefinitions.map(d=>d.text),['培训标题提示']);
});
