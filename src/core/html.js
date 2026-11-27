// 配置結果 → HTML。PDF/PNG 出力（Chrome）とブラウザ版の表示で共用する。
import { richToHTML, paragraphs, bulletType, paraSpaceOf, glowShadowCSS } from './richtext.js';
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

// 色に不透明度を付けた CSS の色（#RRGGBB 以外はそのまま）
function rgba(c0, opacity, fallback) {
  const c = color(c0, fallback);
  const m = /^#?([0-9a-fA-F]{6})$/.exec(c);
  return m ? `rgba(${parseInt(m[1].slice(0, 2), 16)},${parseInt(m[1].slice(2, 4), 16)},${parseInt(m[1].slice(4, 6), 16)},${opacity})` : c;
}

// 光彩を CSS の text-shadow で近似する（ぼかしの小さい影を重ねて濃くし、外側を半径 size までぼかす）
function glowCSS(g) {
  return glowShadowCSS({ ...g, color: color(g.color, '#ffffff') });
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
/* 文字の下地：文字の塊の幅に縮め（折り返す文字は枠の幅いっぱい）、外側に margin の分だけ広げる（広げた分は影で描き、文字の位置を変えない） */
.txt > div.tb { max-width: 100%; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
/* 日本語の字の部分は、行の高さを広げない（英語用と日本語用のフォントで、行の高さの帯の上下の位置が違うため、
   広げると日本語の入った行だけが「文字の大きさ × 行間」より少し高くなり、左右に並ぶ列の行がずれていく） */
.txt .j { font-family: var(--ja); line-height: 0; }
/* 箇条書き：行頭に「•」。折り返した行は文字の頭に揃える */
.pa.li { position: relative; padding-left: 1.1em; }
.pa.li::before { content: '\\2022'; position: absolute; left: 0.1em; }
.pa.li.lih::before { content: none; }
/* 下線：二重線と、画像の線（画像を文字の下に、縦横比を保たずに伸ばして敷く） */
u.ud { text-decoration-style: double; }
u.ui { text-decoration: none; background-repeat: no-repeat; background-position: 0 100%; -webkit-box-decoration-break: clone; box-decoration-break: clone; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.missing { background: repeating-linear-gradient(45deg, #fdd, #fdd 10px, #fff 10px, #fff 20px); }
.ol { border: 1px dashed rgba(128, 128, 128, .8); pointer-events: none; }
.ol span { position: absolute; left: 0; top: 0; background: rgba(220, 0, 80, .85); color: #fff; padding: 0 .3em; font-family: sans-serif; line-height: 1.3; }
</style></head>
<body>
${body}
</body></html>`;
}

function pos(r, L) {
  return `left:${L(r.x)};top:${L(r.y)};width:${L(r.w)};height:${L(r.h)}`;
}

// 矢印：線と矢じり（三角）を SVG で描く。矢じりの大きさは PowerPoint の「三角」（中）に近づけ、長さ・幅とも線の太さの 3 倍。
// 光彩は、ぼかした影を重ねて近似する（PPTX では PowerPoint の「光彩」）
const r4 = (v) => +v.toFixed(4);
export function arrowGeom(e) {
  const { x1, y1, x2, y2, wu } = e;
  const len = Math.hypot(x2 - x1, y2 - y1) || 1e-9;
  const ux = (x2 - x1) / len;
  const uy = (y2 - y1) / len;
  const hl = Math.min(wu * 3, len / (e.head === 'both' ? 2 : 1));
  const hw = wu * 1.5;
  const tri = (tx, ty, dx, dy) => [[tx, ty], [tx - dx * hl - dy * hw, ty - dy * hl + dx * hw], [tx - dx * hl + dy * hw, ty - dy * hl - dx * hw]];
  const heads = [];
  let a = [x1, y1];
  let b = [x2, y2];
  if (e.head === 'end' || e.head === 'both') {
    heads.push(tri(x2, y2, ux, uy));
    b = [x2 - ux * hl * 0.9, y2 - uy * hl * 0.9];
  }
  if (e.head === 'start' || e.head === 'both') {
    heads.push(tri(x1, y1, -ux, -uy));
    a = [x1 + ux * hl * 0.9, y1 + uy * hl * 0.9];
  }
  return { a, b, heads };
}
function arrowSVG(e, L) {
  const R = e.rect;
  const { a, b, heads } = arrowGeom(e);
  const col = color(e.color);
  const g = e.glow;
  const filt = g ? `filter:${[0.25, 0.5, 1].map((k) => `drop-shadow(0 0 ${L(e.glowU * k)} ${rgba(g.color, g.opacity, '#ffff00')})`).join(' ')};` : '';
  const pts = (t) => t.map((p) => p.map(r4).join(',')).join(' ');
  return `<svg data-path="${esc(e.path)}" class="arw" xmlns="http://www.w3.org/2000/svg" viewBox="${r4(R.x)} ${r4(R.y)} ${r4(R.w)} ${r4(R.h)}" style="position:absolute;${pos(R, L)};overflow:visible;pointer-events:none;${filt}">`
    + `<line x1="${r4(a[0])}" y1="${r4(a[1])}" x2="${r4(b[0])}" y2="${r4(b[1])}" stroke="${col}" stroke-width="${r4(e.wu)}" stroke-linecap="butt"/>`
    + heads.map((t) => `<polygon points="${pts(t)}" fill="${col}"/>`).join('')
    + '</svg>';
}

function renderElement(e, L, opts) {
  const data = `data-path="${esc(e.path)}"`;
  if (e.type === 'rect') {
    const bw = e.border ? `${e.border.width ?? 1}pt` : '0';
    const bc = e.border ? color(e.border.color) : 'transparent';
    return `<div ${data} style="${pos(e.rect, L)};background:${e.fill ? color(e.fill) : 'transparent'};border:${bw} solid ${bc}"></div>`;
  }
  if (e.type === 'arrow') return arrowSVG(e, L);
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
      .map((p, i) => {
        // 段落の前の間隔：段落に書いた間隔（<space N>）か、ボックスの段落の間隔。最初の段落には付けない
        const own = paraSpaceOf(p);
        const sp = i ? own ?? e.para_space : 0;
        return `<div class="pa${{ normal: ' li', hidden: ' li lih' }[bulletType(p)] || ''}"${own != null ? ` data-space="${own}"` : ''}${sp ? ` style="margin-top:${sp}pt"` : ''}>${jaSpans(richToHTML(p, '<br class="soft">', opts.srcURL || ((x) => x))) || '<br>'}</div>`;
      })
      .join('');
    const tb = e.text_bg;
    const tbAttr = tb
      ? ` class="tb" style="align-self:${{ left: 'flex-start', center: 'center', right: 'flex-end' }[e.align] || 'stretch'};background:${rgba(tb.color, tb.opacity, '#ffff00')};box-shadow:0 0 0 ${L(tb.margin)} ${rgba(tb.color, tb.opacity, '#ffff00')}"`
      : '';
    return `<div ${data} class="txt" style="${st}"><div${tbAttr}>${paras}</div></div>`;
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
      // PPTX の行間は画面と同じポイントで書くので（pptx.js）、画面と同じ位置に置けば PowerPoint でも文字の下に来る
      // （PowerPoint で開いて測った差は、24pt で 0.2pt、48pt で 1pt 程度）
      for (const r of u.getClientRects()) {
        if (r.width < 0.5) continue;
        out.push({ slide, path: box ? box.dataset.path : '', src: u.dataset.img, rect: { x: (r.left - sr.left) * k, y: (r.bottom - sr.top - th) * k, w: r.width * k, h: th * k } });
      }
    }
  }
  return out;
}

// 文字の下地（text_bg）の四角を、描いた HTML から測る（PPTX では、テキストボックスをこの大きさにして塗りつぶす）。
// rect は文字の塊（下地の margin を含まない）。wrap は、文字が枠の幅いっぱいで折り返しているか（していなければ PPTX で折り返さない）。
// Chrome（puppeteer）の中でも動かすので、この関数の外のものを使わない
export function measureTextBgs(doc, width) {
  const out = [];
  for (const sec of doc.querySelectorAll('section.slide')) {
    const sr = sec.getBoundingClientRect();
    const k = width / sr.width;
    const slide = sec.id.replace(/^slide-/, '');
    for (const tb of sec.querySelectorAll('.txt > div.tb')) {
      const box = tb.parentElement;
      const r = tb.getBoundingClientRect();
      const full = r.width >= box.clientWidth - 0.5;
      out.push({ slide, path: box.dataset.path, rect: { x: (r.left - sr.left) * k, y: (r.top - sr.top) * k, w: r.width * k, h: r.height * k }, wrap: full });
    }
  }
  return out;
}

export function hasTextBg(layout) {
  return layout.slides.some((s) => s.elements.some((e) => e.type === 'text' && e.text_bg));
}

// 文書の中に画像の線があるか（無ければ測らない）
export function hasImageUnderline(layout) {
  return layout.slides.some((s) => s.elements.some((e) => e.type === 'text' && /<u\s[^>]*img\s*=/i.test(String(e.text))));
}
