import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startProvider } from '../src/provider.mjs';
import { startService } from '../src/service.mjs';
import { inspect } from '../src/reader.mjs';
const directory = await mkdtemp(join(tmpdir(), 'intake-demo-'));
const provider = await startProvider();
const service = await startService({ source: provider.url, directory });
try {
  console.log(JSON.stringify({ source: provider.url, product: service.url }));
  console.log('来源目录', await (await fetch(`${provider.url}/catalog`)).json());
  const response = await fetch(`${service.url}/intakes`, { method: 'POST',
    headers: { 'content-type': 'application/json' }, body: JSON.stringify({ path: '/primary' }) });
  if (response.status !== 201) throw new Error(await response.text());
  const record = await response.json();
  console.log('接收结果', record);
  console.log('已保存资料', await (await fetch(`${service.url}/records`)).json());
  const original = Buffer.from(await (await fetch(`${service.url}/objects/${record.digest}`)).arrayBuffer());
  const parsed = inspect(original);
  for (const slide of parsed.slides) {
    console.log(JSON.stringify({ slide: slide.index, hidden: slide.hidden,
      layout: slide.layoutPart, master: slide.masterPart, notes: slide.notesPart,
      learnerText: slide.learnerText, speakerText: slide.speakerText }));
  }
  console.log('逐段来源', parsed.segments.map(segment => ({ slide: segment.slide,
    role: segment.role, audience: segment.audience, text: segment.text, reason: segment.reasonCode })));
  console.log('占位提示（非正文）', parsed.placeholders.map(item => ({ part: item.part, text: item.text })));
} finally { await service.close(); await provider.close(); await rm(directory, { recursive: true, force: true }); }
