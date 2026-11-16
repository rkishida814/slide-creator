// 配置結果 → HTML。PDF/PNG 出力（Chrome）とブラウザ版の表示で共用する。
import { richToHTML } from './richtext.js';
// 長さはキャンバスの単位をそのまま CSS の単位として使う（cm, mm, in, pt, px）。

// PPTX と同じ既定フォント（Arial）。和文は PowerPoint 側では Office テーマの東アジアフォントになる
export const DEFAULT_FONT = 'Arial';
// 画面で選べるフォント。Noto Sans JP は Google Fonts から読み込む（PowerPoint では、その PC に入っているときだけ同じ字形になる）
export const FONTS = ['Helvetica', 'Arial', 'Times New Roman', 'Noto Sans JP'];
const FALLBACK_FONTS = `'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Yu Gothic', 'Meiryo', sans-serif`;

export function color(c, fallback = '#000000') {
  if (!c) return fallback;
  const s = String(c);
  return /^[0-9a-fA-F]{6}$/.test(s) ? `#${s}` : s;
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

/**
 * @param layout  layoutDoc の結果
 * @param opts    { srcURL(src) => url, slides: [id]（省略で全部）, outline: bool }
 */
export function renderHTML(layout, opts = {}) {
  const u = layout.unit;
  const L = (v) => `${+v.toFixed(4)}${u}`;
  const slides = opts.slides ? layout.slides.filter((s) => opts.slides.includes(s.id)) : layout.slides;
  const labelSize = layout.width / 90;
  // Noto Sans JP を使う文字があるときだけ、描画を止めない形で読み込む（つながらなくても他の書体で描く）
  const usesNoto = slides.some((sl) => sl.elements.some((e) => e.type === 'text' && e.font?.family === 'Noto Sans JP'));
  const fontLink = usesNoto
    ? '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap" media="print" onload="this.media=\'all\'">\n'
    : '';

  const body = slides
    .map((s) => {
      const els = s.elements.map((e) => renderElement(e, L, opts)).join('\n');
      const outline = opts.outline
        ? s.boxes
            .filter((b) => b.depth > 0)
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
.tex { background-repeat: repeat; background-position: 0 0; }
.img { overflow: hidden; }
.img img { position: absolute; display: block; max-width: none; }
.txt { display: flex; flex-direction: column; overflow: visible; }
.txt > div { white-space: pre-wrap; overflow-wrap: anywhere; }
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
  if (e.type === 'texture') {
    const url = opts.srcURL ? opts.srcURL(e.src) : e.src;
    return `<div ${data} class="tex" style="${pos(e.rect, L)};background-image:url('${esc(url)}');background-size:${L(e.tileW)} auto"></div>`;
  }
  if (e.type === 'missing') return `<div ${data} class="missing" style="${pos(e.rect, L)}"></div>`;
  if (e.type === 'image') {
    const [l, t, r, b] = e.crop;
    const fw = e.rect.w / (r - l); // 切り抜く前の画像全体を表示した場合の大きさ
    const fh = e.rect.h / (b - t);
    const url = opts.srcURL ? opts.srcURL(e.src) : e.src;
    return `<div ${data} class="img" style="${pos(e.rect, L)}"><img src="${esc(url)}" style="left:${L(-l * fw)};top:${L(-t * fh)};width:${L(fw)};height:${L(fh)}"></div>`;
  }
  if (e.type === 'text') {
    const f = e.font;
    const jc = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[e.valign];
    const fam = `'${String(f.family || DEFAULT_FONT).replace(/['"<>]/g, '')}', ${FALLBACK_FONTS}`;
    const st = [
      pos(e.rect, L),
      `justify-content:${jc}`,
      `text-align:${e.align}`,
      `font-family:${fam}`,
      `font-size:${f.size}pt`,
      `font-weight:${f.bold ? 700 : 400}`,
      `font-style:${f.italic ? 'italic' : 'normal'}`,
      `color:${color(f.color, '#000000')}`,
      `line-height:${e.line_height}`,
    ].join(';');
    return `<div ${data} class="txt" style="${st}"><div>${richToHTML(e.text)}</div></div>`;
  }
  return '';
}
