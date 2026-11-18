// PPTX の書き出し（Node 用）。組み立ては src/core/pptx.js。
import fs from 'node:fs';
import path from 'node:path';
import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';
import { buildPPTX, finishPPTX } from '../core/pptx.js';
import { hasImageUnderline } from '../core/html.js';

export async function exportPPTX(project, outFile) {
  // 画像の線があれば、Chrome で描いて位置を測る
  const underlines = hasImageUnderline(project.layout) ? await (await import('./chrome.js')).measureProjectUnderlines(project) : [];
  const pres = await buildPPTX(pptxgen, project.layout, (src) => ({ path: path.join(project.dir, src) }), { underlines });
  fs.writeFileSync(outFile, await finishPPTX(pres, JSZip, 'nodebuffer'));
}
