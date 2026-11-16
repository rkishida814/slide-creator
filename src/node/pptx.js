// PPTX の書き出し（Node 用）。組み立ては src/core/pptx.js。
import path from 'node:path';
import pptxgen from 'pptxgenjs';
import { buildPPTX } from '../core/pptx.js';

export async function exportPPTX(project, outFile) {
  const pres = await buildPPTX(pptxgen, project.layout, (src) => ({ path: path.join(project.dir, src) }));
  await pres.writeFile({ fileName: outFile });
}
