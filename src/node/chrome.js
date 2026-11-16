// Chrome（puppeteer-core）で HTML を開き、PDF・PNG の書き出しと文字のはみ出し検査をする。
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import puppeteer from 'puppeteer-core';
import { renderHTML } from '../core/html.js';
import { PER_INCH } from '../core/units.js';

const CANDIDATES = [
  process.env.SLIDE_CHROME,
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  '/Applications/Chromium.app/Contents/MacOS/Chromium',
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  '/usr/bin/google-chrome',
  '/usr/bin/chromium',
].filter(Boolean);

function chromePath() {
  const p = CANDIDATES.find((c) => fs.existsSync(c));
  if (!p) throw new Error('Chrome が見つかりません。環境変数 SLIDE_CHROME に実行ファイルの場所を設定してください');
  return p;
}

// HTML はプロジェクトの out/ に書き、file:// で開く（画像を相対参照で読むため）
async function openProject(project, htmlOpts) {
  const outDir = path.join(project.dir, 'out');
  fs.mkdirSync(outDir, { recursive: true });
  const html = renderHTML(project.layout, {
    ...htmlOpts,
    srcURL: (src) => pathToFileURL(path.join(project.dir, src)).href,
  });
  const htmlFile = path.join(outDir, '.render.html');
  fs.writeFileSync(htmlFile, html);
  const browser = await puppeteer.launch({ executablePath: chromePath(), headless: true, args: ['--allow-file-access-from-files'] });
  const page = await browser.newPage();
  await page.goto(pathToFileURL(htmlFile).href, { waitUntil: 'load' });
  await page.evaluate(() => document.fonts.ready);
  return { browser, page };
}

export async function exportPDF(project, outFile) {
  const { browser, page } = await openProject(project, {});
  try {
    const { unit, width, height } = project.layout;
    // Chrome の PDF は in / px 指定が確実なので in に直す
    await page.pdf({
      path: outFile,
      width: `${width / PER_INCH[unit]}in`,
      height: `${height / PER_INCH[unit]}in`,
      printBackground: true,
      preferCSSPageSize: true,
    });
  } finally {
    await browser.close();
  }
}

/**
 * スライド全体か、特定ノードの枠を PNG にする。
 * @param target { slide, box }  box が null ならスライド全体
 */
export async function renderPNG(project, target, outFile, { width = 1600, outline = false } = {}) {
  const { browser, page } = await openProject(project, { slides: [target.slide.id], outline });
  try {
    const px = (v) => (v / PER_INCH[project.layout.unit]) * 96;
    const r = target.box ? target.box.rect : target.slide.rect;
    const clipW = px(r.w);
    const scale = Math.min(width / clipW, 8);
    await page.setViewport({ width: Math.ceil(px(project.layout.width)) + 1, height: Math.ceil(px(project.layout.height)) + 1, deviceScaleFactor: scale });
    const sec = await page.$eval(`section[id="slide-${target.slide.id}"]`, (el) => {
      const b = el.getBoundingClientRect();
      return { x: b.x, y: b.y };
    });
    await page.screenshot({ path: outFile, clip: { x: sec.x + px(r.x), y: sec.y + px(r.y), width: clipW, height: px(r.h) } });
  } finally {
    await browser.close();
  }
}

// 文字が枠からはみ出していないかを、実際に描いて測る
export async function measureText(project) {
  const { browser, page } = await openProject(project, {});
  try {
    return await page.evaluate(() => {
      const out = [];
      for (const box of document.querySelectorAll('.txt')) {
        const inner = box.firstElementChild;
        const b = box.getBoundingClientRect();
        const i = inner.getBoundingClientRect();
        const overH = i.height - b.height;
        const overW = inner.scrollWidth - box.clientWidth;
        if (overH > 0.5 || overW > 0.5) {
          out.push({ path: box.dataset.path, overflowHeightPct: Math.round((overH / b.height) * 100), overflowWidth: overW > 0.5 });
        }
      }
      return out;
    });
  } finally {
    await browser.close();
  }
}
