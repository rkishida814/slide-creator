// 配置結果 → HTML。PDF/PNG 出力（Chrome）とブラウザ版の表示で共用する。
import { richToHTML, paragraphs, isBullet } from './richtext.js';
// 長さはキャンバスの単位をそのまま CSS の単位として使う（cm, mm, in, pt, px）。

// フォントは、英数字（英語）用と日本語用を別々に選ぶ（font.family と font.family_ja）。
// 日本語の字（かな・漢字・全角の記号）は日本語用、それ以外は英語用のフォントで書く。PPTX でも同じに分ける
export const DEFAULT_FONT = 'Helvetica';
export const DEFAULT_FONT_JA = 'BIZ UDゴシック';
export const JA_FONTS = ['BIZ UDゴシック', 'ヒラギノ角ゴシック', 'Noto Sans JP', 'メイリオ'];
// 英語用には、日本語のフォントも選べる（Helvetica・Arial・Times New Roman は英語用だけ）
export const FONTS = [...JA_FONTS, 'Helvetica', 'Arial', 'Times New Roman'];
// css：ブラウザで使う名前、pptx：PowerPoint に書く名前、web：手元に無いとき Google Fonts から読み込む名前
const FONT_DEF = {
  // PowerPoint に書く名前は英語の名前にする（Mac で Google Fonts から入れたものも、Windows に入っているものも、この名前で見つかる）
  'BIZ UDゴシック': { css: `'BIZ UDGothic', 'BIZ UDゴシック'`, pptx: 'BIZ UDGothic', web: 'BIZ+UDGothic' },
  // 前の版の既定（選択肢には出さない。すでに書いてあるファイルのため）
  'BIZ UDPゴシック': { css: `'BIZ UDPGothic', 'BIZ UDPゴシック'`, pptx: 'BIZ UDPGothic', web: 'BIZ+UDPGothic' },
  'ヒラギノ角ゴシック': { css: `'Hiragino Sans', 'Hiragino Kaku Gothic ProN'`, pptx: 'Hiragino Sans' },
  'Noto Sans JP': { css: `'Noto Sans JP'`, pptx: 'Noto Sans JP', web: 'Noto+Sans+JP' },
  'メイリオ': { css: `'Meiryo', 'メイリオ'`, pptx: 'Meiryo' },
};
const cssName = (name) => FONT_DEF[name]?.css || `'${String(name).replace(/['"<>;]/g, '')}'`;
export const pptxFontName = (name) => FONT_DEF[name]?.pptx || String(name);
// 日本語の字（かな・漢字・全角の記号など）
export const JA_CHARS = /[\u2E80-\u9FFF\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFFEF]+/g;
const jaSpans = (html) => html.split(/(<[^>]*>)/).map((part) => (part.startsWith('<') ? part : part.replace(JA_CHARS, (m) => `<span class="j">${m}</span>`))).join('');

export function color(c, fallback = '#000000') {
  if (!c) return fallback;
  const s = String(c);
  return /^[0-9a-fA-F]{6}$/.test(s) ? `#${s}` : s;
}

// 文字の光彩（PowerPoint の「光彩」）。font.glow: true か {color, size（pt）, opacity（0〜1）}。無ければ null
export function glowOf(f) {
  const g = f?.glow;
  if (!g) return null;
  const o = g === true ? {} : g;
  return { color: o.color ?? '#ffffff', size: Number(o.size ?? 6), opacity: Number(o.opacity ?? 1) };
}

// 光彩を CSS の text-shadow で近似する（ぼかしの小さい影を重ねて濃くし、外側を半径 size までぼかす）
function glowCSS(g) {
  const c = color(g.color, '#ffffff');
  const m = /^#?([0-9a-fA-F]{6})$/.exec(c);
  const col = m
    ? `rgba(${parseInt(m[1].slice(0, 2), 16)},${parseInt(m[1].slice(2, 4), 16)},${parseInt(m[1].slice(4, 6), 16)},${g.opacity})`
    : c;
  const r = g.size;
  return [0.12, 0.12, 0.2, 0.2, 0.3, 0.3, 0.45, 0.65, 1].map((b) => `0 0 ${(r * b).toFixed(2)}pt ${col}`).join(',');
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

/**
 * @param layout  layoutDoc の結果
 * @param opts    { srcURL(src) => url, slides: [id]（省略で全部）, outline: bool }
 */
// 文書で使っているフォントを、実際の文字について読み込み終えるまで待つ（PDF に書き出す前に使う）。
// フォントは、使う字の部分ごとに分けて配られるので、書いてある文字を渡して読み込ませる
export async function loadUsedFonts(doc) {
  const jobs = [];
  for (const el of doc.querySelectorAll('.txt')) {
    const cs = doc.defaultView.getComputedStyle(el);
    const text = el.textContent || '';
    if (!text.trim()) continue;
    const ja = cs.getPropertyValue('--ja').trim();
    for (const fam of [cs.fontFamily, ja].filter(Boolean)) {
      for (const w of ['400', '700']) jobs.push(doc.fonts.load(`${w} 16px ${fam}`, text).catch(() => null));
    }
  }
  await Promise.all(jobs);
  await doc.fonts.ready;
}

export function renderHTML(layout, opts = {}) {
  const u = layout.unit;
  const L = (v) => `${+v.toFixed(4)}${u}`;
  const slides = opts.slides ? layout.slides.filter((s) => opts.slides.includes(s.id)) : layout.slides;
  const labelSize = layout.width / 90;
  // Google Fonts のフォントを使う文字があるときだけ、描画を止めない形で読み込む（つながらなくても他の書体で描く）
  const web = new Set();
  for (const sl of slides) {
    for (const e of sl.elements) {
      if (e.type !== 'text') continue;
      for (const n of [e.font?.family || DEFAULT_FONT, e.font?.family_ja || DEFAULT_FONT_JA]) if (FONT_DEF[n]?.web) web.add(FONT_DEF[n].web);
    }
  }
  const fontLink = [...web]
    // PDF に書き出すとき（opts.waitFonts）は、読み込みを待つ形で読み込む
    .map((w) => (opts.waitFonts
      ? `<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=${w}:wght@400;700&display=block">\n`
      : `<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=${w}:wght@400;700&display=swap" media="print" onload="this.media='all'">\n`))
    .join('');

  const body = slides
    .map((s) => {
      const els = s.elements.map((e) => renderElement(e, L, opts)).join('\n');
      const outline = opts.outline
        ? (opts.outlineBoxes || s.boxes.filter((b) => b.depth > 0))
            .map((b) => {
              return `<div class="ol" style="${pos(b.rect, L)}"></div>`;
            })
            .join('\n')
        : '';
      return `<section class="slide" id="slide-${esc(s.id)}" style="width:${L(layout.width)};height:${L(layout.height)}">\n${els}\n${outline}\n</section>`;
    })
    .join('\n');

  return `<!doctype html>
<html><head><meta charset="utf-8">
${fontLink}<style>
@page { size: ${L(layout.width)} ${L(layout.height)}; margin: 0; }
html, body { margin: 0; padding: 0; background: #fff; }
.slide { position: relative; overflow: hidden; background: #fff; break-after: page; }
.slide > div { position: absolute; box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.img { overflow: hidden; }
.img img { position: absolute; display: block; max-width: none; }
.txt { display: flex; flex-direction: column; overflow: visible; }
.txt > div { white-space: pre-wrap; overflow-wrap: anywhere; }
.txt .j { font-family: var(--ja); }
/* 箇条書き：行頭に「•」。折り返した行は文字の頭に揃える */
.pa.li { position: relative; padding-left: 1.1em; }
.pa.li::before { content: '\\2022'; position: absolute; left: 0.1em; }
/* 下線：二重線と、画像の線（画像を文字の下に、縦横比を保たずに伸ばして敷く） */
u.ud { text-decoration-style: double; }
u.ui { text-decoration: none; background-repeat: no-repeat; background-position: 0 100%; -webkit-box-decoration-break: clone; box-decoration-break: clone; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.missing { background: repeating-linear-gradient(45deg, #fdd, #fdd 10px, #fff 10px, #fff 20px); }
.ol { border: 1px dashed rgba(220, 0, 80, .8); pointer-events: none; }
.ol span { position: absolute; left: 0; top: 0; background: rgba(220, 0, 80, .85); color: #fff; padding: 0 .3em; font-family: sans-serif; line-height: 1.3; }
</style></head>
<body>
${body}
</body></html>`;
}

function pos(r, L) {
  return `left:${L(r.x)};top:${L(r.y)};width:${L(r.w)};height:${L(r.h)}`;
}

function renderElement(e, L, opts) {
  const data = `data-path="${esc(e.path)}"`;
  if (e.type === 'rect') {
    const bw = e.border ? `${e.border.width ?? 1}pt` : '0';
    const bc = e.border ? color(e.border.color) : 'transparent';
    return `<div ${data} style="${pos(e.rect, L)};background:${e.fill ? color(e.fill) : 'transparent'};border:${bw} solid ${bc}"></div>`;
  }
  if (e.type === 'missing') return `<div ${data} class="missing" style="${pos(e.rect, L)}"></div>`;
  if (e.type === 'image') {
    const [l, t, r, b] = e.crop;
    const fw = e.rect.w / (r - l); // 切り抜く前の画像全体を表示した場合の大きさ
    const fh = e.rect.h / (b - t);
    const url = opts.srcURL ? opts.srcURL(e.src) : e.src;
    return `<div ${data} class="img${e.background ? ' bg' : ''}" style="${pos(e.rect, L)}"><img src="${esc(url)}" style="left:${L(-l * fw)};top:${L(-t * fh)};width:${L(fw)};height:${L(fh)}"></div>`;
  }
  if (e.type === 'text') {
    const f = e.font;
    const jc = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[e.valign];
    const ja = `${cssName(f.family_ja || DEFAULT_FONT_JA)}, 'Hiragino Sans', sans-serif`;
    const fam = `${cssName(f.family || DEFAULT_FONT)}, ${ja}`;
    const st = [
      pos(e.rect, L),
      `justify-content:${jc}`,
      `text-align:${e.align}`,
      `font-family:${fam}`,
      `--ja:${ja}`,
      `font-size:${f.size}pt`,
      `font-weight:${f.bold ? 700 : 400}`,
      `font-style:${f.italic ? 'italic' : 'normal'}`,
      `color:${color(f.color, '#000000')}`,
      `line-height:${e.line_height}`,
      ...(glowOf(f) ? [`text-shadow:${glowCSS(glowOf(f))}`] : []),
    ].join(';');
    const paras = paragraphs(e.text)
      .map((p, i) => `<div class="pa${isBullet(p) ? ' li' : ''}"${i && e.para_space ? ` style="margin-top:${e.para_space}pt"` : ''}>${jaSpans(richToHTML(p, '<br>', opts.srcURL || ((x) => x))) || '<br>'}</div>`)
      .join('');
    return `<div ${data} class="txt" style="${st}"><div>${paras}</div></div>`;
  }
  return '';
}

// 画像の線（<u img=…>）の位置を、描いた HTML から測る（PPTX では、測った位置に画像を置く）。
// doc は renderHTML の HTML を拡大・縮小せずに開いた文書。width はキャンバスの幅（単位はキャンバスの単位）。
// Chrome（puppeteer）の中でも動かすので、この関数の外のものを使わない
export function measureUnderlines(doc, width) {
  const out = [];
  for (const sec of doc.querySelectorAll('section.slide')) {
    const sr = sec.getBoundingClientRect();
    const k = width / sr.width; // px → キャンバスの単位
    const slide = sec.id.replace(/^slide-/, '');
    for (const u of sec.querySelectorAll('u.ui')) {
      const box = u.closest('[data-path]');
      const th = ((parseFloat(u.dataset.w) || 2) * 96) / 72; // 太さ（pt → px）
      // PowerPoint では文字が画面より少し下に描かれ、画面と同じ位置だと線が文字に重なる。
      // そのため PPTX では、画面の線の位置から太さの分だけ下げて置く
      for (const r of u.getClientRects()) {
        if (r.width < 0.5) continue;
        out.push({ slide, path: box ? box.dataset.path : '', src: u.dataset.img, rect: { x: (r.left - sr.left) * k, y: (r.bottom - sr.top) * k, w: r.width * k, h: th * k } });
      }
    }
  }
  return out;
}

// 文書の中に画像の線があるか（無ければ測らない）
export function hasImageUnderline(layout) {
  return layout.slides.some((s) => s.elements.some((e) => e.type === 'text' && /<u\s[^>]*img\s*=/i.test(String(e.text))));
}
