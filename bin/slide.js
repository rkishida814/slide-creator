#!/usr/bin/env node
// slide-creator の CLI（AI もこれを使う）。使い方は `slide help`。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadProject, findTarget, LAYOUT_FILE } from '../src/node/project.js';

const HELP = `slide-creator

  slide init <dir>                     新しいプロジェクト（layout.yaml）を作る
  slide tree <dir> [--slide ID] [--json]
                                       ノードの木と、各枠の位置・大きさ（キャンバスの単位）を表示
  slide check <dir>                    設定の誤り・画像の解像度・文字のはみ出しを検査
  slide render <dir> [--slide ID] [--node ID] [--outline] [--width PX] [-o FILE]
                                       スライド全体または特定ノードの枠を PNG にする
                                       （--outline で枠と id を重ねて描く）
  slide export <dir> [--pptx] [--pdf] [-o OUTDIR]
                                       PPTX / PDF を書き出す（指定なしなら両方。既定は <dir>/out/）

ノードの指定: id（"photos"）、スライド付き（"s1/photos"）、自動の経路（"s1.1.0"）`;

function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '-o') args.o = argv[++i];
    else if (a.startsWith('--')) {
      const key = a.slice(2);
      const next = argv[i + 1];
      if (next != null && !next.startsWith('-') && ['slide', 'node', 'width'].includes(key)) args[key] = argv[++i];
      else args[key] = true;
    } else args._.push(a);
  }
  return args;
}

const fmt = (v) => v.toFixed(2);

function printTree(layout, slideId) {
  const u = layout.unit;
  for (const s of layout.slides) {
    if (slideId && s.id !== slideId) continue;
    console.log(`${s.id}  [${fmt(layout.width)} x ${fmt(layout.height)} ${u}]`);
    const byPath = new Map(s.elements.filter((e) => e.type !== 'rect').map((e) => [e.path, e]));
    for (const b of s.boxes) {
      if (b.depth === 0) continue;
      const e = byPath.get(b.path);
      let info = '';
      if (e?.type === 'image') {
        const c = e.crop.map((x) => +x.toFixed(3));
        info = `${e.src}  表示 ${fmt(e.rect.w)}x${fmt(e.rect.h)}  crop [${c}]  ${e.dpi}dpi`;
      } else if (e?.type === 'missing') info = `画像が無い: ${e.src}`;
      else if (e?.type === 'text') info = JSON.stringify(e.text.length > 30 ? e.text.slice(0, 30) + '…' : e.text) + `  ${e.font.size}pt ${e.align}/${e.valign}`;
      const name = b.id ? `${b.path}` : b.path;
      const r = b.rect;
      console.log(`${'  '.repeat(b.depth)}${name}  ${b.kind}  (${fmt(r.x)}, ${fmt(r.y)}) ${fmt(r.w)}x${fmt(r.h)}  ${info}`);
    }
    for (const i of s.issues) console.log(`  [${i.level}] ${i.path}: ${i.msg}`);
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const [cmd, dir] = args._;
  if (!cmd || cmd === 'help' || args.help) return console.log(HELP);
  if (!dir) throw new Error('プロジェクトのフォルダを指定してください');

  if (cmd === 'init') {
    if (fs.existsSync(path.join(dir, LAYOUT_FILE))) throw new Error(`${dir}/${LAYOUT_FILE} はすでにあります`);
    const here = path.dirname(fileURLToPath(import.meta.url));
    fs.mkdirSync(dir, { recursive: true });
    fs.copyFileSync(path.join(here, '..', 'templates', 'init.yaml'), path.join(dir, LAYOUT_FILE));
    return console.log(`作成: ${path.join(dir, LAYOUT_FILE)}`);
  }

  const project = loadProject(dir);
  const outDir = path.join(project.dir, 'out');

  if (cmd === 'tree') {
    if (args.json) return console.log(JSON.stringify(project.layout, null, 1));
    return printTree(project.layout, args.slide);
  }

  if (cmd === 'check') {
    const { measureText } = await import('../src/node/chrome.js');
    const issues = [...project.layout.issues, ...project.layout.slides.flatMap((s) => s.issues)];
    for (const o of await measureText(project)) {
      issues.push({
        level: 'error',
        path: o.path,
        msg: o.overflowWidth ? '文字が横にはみ出している（長い語）' : `文字が枠の高さを ${o.overflowHeightPct}% 超えている`,
      });
    }
    if (!issues.length) return console.log('問題なし');
    for (const i of issues) console.log(`[${i.level}] ${i.path}: ${i.msg}`);
    if (issues.some((i) => i.level === 'error')) process.exitCode = 1;
    return;
  }

  if (cmd === 'render') {
    const { renderPNG } = await import('../src/node/chrome.js');
    const target = findTarget(project.layout, args.slide, args.node);
    const name = `preview_${target.slide.id}${target.box ? '_' + target.box.path.replace(/[/.]/g, '_') : ''}.png`;
    const out = args.o || path.join(outDir, name);
    fs.mkdirSync(path.dirname(path.resolve(out)), { recursive: true });
    await renderPNG(project, target, out, { width: +(args.width || 1600), outline: !!args.outline });
    return console.log(`書き出し: ${path.resolve(out)}`);
  }

  if (cmd === 'export') {
    const both = !args.pptx && !args.pdf;
    const od = args.o || outDir;
    fs.mkdirSync(od, { recursive: true });
    const base = path.basename(project.dir);
    if (both || args.pptx) {
      const { exportPPTX } = await import('../src/node/pptx.js');
      const f = path.join(od, `${base}.pptx`);
      await exportPPTX(project, f);
      console.log(`書き出し: ${path.resolve(f)}`);
    }
    if (both || args.pdf) {
      const { exportPDF } = await import('../src/node/chrome.js');
      const f = path.join(od, `${base}.pdf`);
      await exportPDF(project, f);
      console.log(`書き出し: ${path.resolve(f)}`);
    }
    return;
  }
  throw new Error(`知らないコマンド: ${cmd}\n\n${HELP}`);
}

main().catch((e) => {
  console.error(`エラー: ${e.message}`);
  process.exit(1);
});
