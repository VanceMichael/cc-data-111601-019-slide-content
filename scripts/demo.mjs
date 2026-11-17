import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startProvider } from '../src/provider.mjs';
import { startService } from '../src/service.mjs';
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
  for (const slide of record.parsed.slides) {
    console.log(`第${slide.position}页 ${slide.part} visible=${slide.visible}`);
    for (const item of slide.texts)
      console.log(`  [${item.audience}/${item.origin}] ${JSON.stringify(item.text)} (${item.part})`);
    for (const definition of slide.placeholderDefinitions)
      console.log(`  [定义/不进正文] ${JSON.stringify(definition.text)} (${definition.part})`);
    for (const part of slide.missing) console.log(`  [缺失部件] ${part}`);
  }
  console.log('已保存资料', await (await fetch(`${service.url}/records`)).json());
} finally { await service.close(); await provider.close(); await rm(directory, { recursive: true, force: true }); }
