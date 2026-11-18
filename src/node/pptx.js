// PPTX の書き出し（Node 用）。組み立ては src/core/pptx.js。
import fs from 'node:fs';
import path from 'node:path';
import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';
import { buildPPTX, finishPPTX } from '../core/pptx.js';
import { hasImageUnderline, hasTextBg } from '../core/html.js';

export async function exportPPTX(project, outFile) {
  // 画像の線があれば、Chrome で描いて位置を測る
  // 文字の下地があれば、同じく文字の塊の大きさを測る
  const ul = hasImageUnderline(project.layout);
  const tb = hasTextBg(project.layout);
  const { underlines, textBgs } = ul || tb ? await (await import('./chrome.js')).measureProjectForPPTX(project, { underlines: ul, textBgs: tb }) : { underlines: [], textBgs: [] };
  const pres = await buildPPTX(pptxgen, project.layout, (src) => ({ path: path.join(project.dir, src) }), { underlines, textBgs });
  fs.writeFileSync(outFile, await finishPPTX(pres, JSZip, 'nodebuffer'));
}
