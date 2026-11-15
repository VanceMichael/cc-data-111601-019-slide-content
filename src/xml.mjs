export function parseXml(bytes) {
  const text = Buffer.isBuffer(bytes) ? bytes.toString('utf8') : bytes;
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('unsupported_xml_declaration');
  const document = { name: '#document', attrs: {}, children: [], text: '' };
  const stack = [document];
  const decode = value => value.replace(/&(?:amp|lt|gt|quot|apos);/g,
    entity => ({ '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" })[entity]);
  const tokens = text.match(/<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<[^>]+>|[^<]+/g) ?? [];
  for (const token of tokens) {
    if (token.startsWith('<!--') || token.startsWith('<?')) continue;
    if (token.startsWith('<![CDATA[')) { stack.at(-1).text += token.slice(9, -3); continue; }
    if (token.startsWith('</')) {
      if (stack.length === 1 || stack.at(-1).name !== token.slice(2, -1).trim()) throw new Error('xml_nesting');
      stack.pop();
    } else if (token.startsWith('<')) {
      const name = /^<([^\s/>]+)/.exec(token)?.[1];
      if (!name) throw new Error('xml_tag');
      const attrs = {};
      for (const match of token.matchAll(/([^\s=<>]+)\s*=\s*(["'])(.*?)\2/g)) attrs[match[1]] = decode(match[3]);
      const node = { name, attrs, children: [], text: '' };
      stack.at(-1).children.push(node);
      if (!token.endsWith('/>')) stack.push(node);
    } else stack.at(-1).text += decode(token);
  }
  if (stack.length !== 1 || document.children.length !== 1) throw new Error('xml_document');
  return document.children[0];
}
export const local = node => node.name.split(':').at(-1);
export function descendants(node, name) {
  return node.children.flatMap(child => [ ...(local(child) === name ? [child] : []), ...descendants(child, name) ]);
}
export const content = node => node.text + node.children.map(content).join('');
