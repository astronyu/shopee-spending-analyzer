import assert from 'node:assert/strict';
import fs from 'node:fs';

const source = fs.readFileSync(new URL('../xlsx-writer.js', import.meta.url), 'utf8');
const { buildXlsx, xlsxLink } = await import(`data:text/javascript;base64,${Buffer.from(source).toString('base64')}`);
const blob = buildXlsx([
  { name: 'Orders', rows: [['Shop'], [xlsxLink('https://shopee.com.my/shop/123', 'Open Shop')]], autoFilter: true },
  { name: 'Items', rows: [['Product', 'Shop', 'Missing'], [xlsxLink('https://shopee.com.my/product/123/456?a=1&b=2', 'Open Product'), xlsxLink('https://shopee.com.my/shop/123', 'Shop & Seller'), xlsxLink('')]] },
  { name: 'Summary', rows: [['Total', 25]] }
]);
// Read the uncompressed ZIP entries without relying on the writer's helpers.
const bytes = Buffer.from(await blob.arrayBuffer());
const entries = new Map();
let offset = 0;
while (bytes.readUInt32LE(offset) === 0x04034b50) {
  assert.equal(bytes.readUInt16LE(offset + 8), 0);
  const size = bytes.readUInt32LE(offset + 18);
  const nameLength = bytes.readUInt16LE(offset + 26);
  const extraLength = bytes.readUInt16LE(offset + 28);
  const start = offset + 30 + nameLength + extraLength;
  const name = bytes.subarray(offset + 30, offset + 30 + nameLength).toString();
  entries.set(name, bytes.subarray(start, start + size).toString());
  offset = start + size;
}
const orders = entries.get('xl/worksheets/sheet1.xml');
const items = entries.get('xl/worksheets/sheet2.xml');
const rels = entries.get('xl/worksheets/_rels/sheet2.xml.rels');
assert.match(orders, /<c r="A2" s="6" t="inlineStr"><is><t xml:space="preserve">Open Shop<\/t>/);
assert.match(items, /Open Product/);
assert.match(items, /Shop &amp; Seller/);
assert.match(items, /<hyperlink ref="A2" r:id="rId1"\/>/);
assert.match(items, /<hyperlink ref="B2" r:id="rId2"\/>/);
assert.doesNotMatch(items, /<hyperlink ref="C2"|<f>/);
assert.match(rels, /Target="https:\/\/shopee.com.my\/product\/123\/456\?a=1&amp;b=2" TargetMode="External"/);
assert.match(rels, /Id="rId2"[^>]+Target="https:\/\/shopee.com.my\/shop\/123"/);
assert.ok(orders.indexOf('<hyperlinks>') > orders.indexOf('<autoFilter'));
assert.ok(!entries.has('xl/worksheets/_rels/sheet3.xml.rels'));
assert.match(entries.get('xl/worksheets/sheet3.xml'), /<v>25<\/v>/);
console.log('PASS: visible labels, external targets, XML escaping, worksheet-local relationship IDs, missing links and numeric cells.');
