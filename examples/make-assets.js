// 見本用の PNG を作る（外部ツール不要）。node examples/make-assets.js
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
};

function png(w, h, pixel) {
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w * 3 + 1)] = 0;
    for (let x = 0; x < w; x++) {
      const [r, g, b] = pixel(x, y);
      const o = y * (w * 3 + 1) + 1 + x * 3;
      raw[o] = r; raw[o + 1] = g; raw[o + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

// 色の背景に、中心の円と 10% 刻みの格子（トリミングの位置が分かるように）
function sample(w, h, base) {
  return png(w, h, (x, y) => {
    const fx = x / w, fy = y / h;
    const grid = Math.abs(fx * 10 - Math.round(fx * 10)) < 0.02 || Math.abs(fy * 10 - Math.round(fy * 10)) < 0.02;
    if (grid) return [255, 255, 255];
    const d = Math.hypot(fx - 0.5, (fy - 0.5) * (h / w));
    if (d < 0.18) return [30, 30, 30];
    return base.map((c) => Math.round(c * (0.6 + 0.4 * fy)));
  });
}

const dir = path.join(path.dirname(fileURLToPath(import.meta.url)), 'demo');
fs.mkdirSync(dir, { recursive: true });
const specs = [
  ['a1.png', 1200, 900, [220, 90, 80]],
  ['a2.png', 1200, 900, [80, 160, 220]],
  ['a3.png', 900, 1200, [90, 190, 110]],
  ['a4.png', 1600, 900, [230, 170, 60]],
  ['a5.png', 1000, 1000, [150, 100, 200]],
  ['a6.png', 400, 300, [120, 120, 120]], // 低解像度（警告の見本）
];
for (const [name, w, h, c] of specs) fs.writeFileSync(path.join(dir, name), sample(w, h, c));
console.log(`作成: ${dir}`);
