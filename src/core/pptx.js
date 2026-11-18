// 配置結果 → 編集できる PPTX。
// 画像は元画像をそのまま埋め込み、トリミングは PowerPoint の「トリミング」（srcRect）で表す。
// 文字はテキストボックス（余白0・自動調整なし）。
// ファイルに依存しない。Node・ブラウザの両方から使う。
import { PER_INCH } from './units.js';
import { DEFAULT_FONT, DEFAULT_FONT_JA, pptxFontName, glowOf } from './html.js';
import { parseRich, hasRich, paragraphs, bulletType } from './richtext.js';

const hex = (c, fb) => {
  if (!c) return fb;
  const s = String(c).replace(/^#/, '');
  if (/^[0-9a-fA-F]{3}$/.test(s)) return s.split('').map((ch) => ch + ch).join('');
  if (/^[0-9a-fA-F]{6}$/.test(s)) return s;
  throw new Error(`PPTX では色を 16 進（#RRGGBB）で書いてください: ${c}`);
};

/**
 * @param PptxGen   pptxgenjs のクラス
 * @param layout    layoutDoc の結果
 * @param imageRef  (src) => { path } | { data: 'image/png;base64,...' }
 * @param opts      { underlines: measureUnderlines の結果（画像の線を、測った位置に画像として置く）,
 *                    textBgs: measureTextBgs の結果（文字の下地のあるテキストボックスを、測った文字の塊の大きさにする） }
 * @returns pptxgenjs のインスタンス（writeFile / write で書き出す）
 */
export async function buildPPTX(PptxGen, layout, imageRef, opts = {}) {
  const k = 1 / PER_INCH[layout.unit]; // キャンバスの単位 → インチ
  const pres = new PptxGen();
  pres.defineLayout({ name: 'SLIDE_CREATOR', width: layout.width * k, height: layout.height * k });
  pres.layout = 'SLIDE_CREATOR';

  for (const s of layout.slides) {
    const slide = pres.addSlide();
    for (const e of s.elements) {
      const box = { x: e.rect.x * k, y: e.rect.y * k, w: e.rect.w * k, h: e.rect.h * k };
      if (e.type === 'rect') {
        slide.addShape(pres.ShapeType.rect, {
          ...box,
          fill: e.fill ? { color: hex(e.fill) } : { type: 'none' },
          line: e.border ? { color: hex(e.border.color, '000000'), width: e.border.width ?? 1 } : { type: 'none' },
        });
      } else if (e.type === 'missing') {
        slide.addShape(pres.ShapeType.rect, { ...box, fill: { color: 'FFDDDD' }, line: { color: 'FF0000', width: 1 } });
        slide.addText(`画像が無い: ${e.src}`, { ...box, fontSize: 10, color: 'CC0000', align: 'center', valign: 'middle' });
      } else if (e.type === 'image') {
        const [l, t, r, b] = e.crop;
        const full = { w: box.w / (r - l), h: box.h / (b - t) }; // 切り抜く前の全体の大きさ
        const cropped = l > 1e-6 || t > 1e-6 || r < 1 - 1e-6 || b < 1 - 1e-6;
        slide.addImage({
          ...(await imageRef(e.src)),
          x: box.x,
          y: box.y,
          ...(cropped
            ? { w: full.w, h: full.h, sizing: { type: 'crop', x: l * full.w, y: t * full.h, w: box.w, h: box.h } }
            : { w: box.w, h: box.h }),
          altText: e.path,
        });
      } else if (e.type === 'text') {
        const f = e.font;
        const g = glowOf(f);
        const glow = g ? { color: hex(g.color, 'FFFFFF'), size: g.size, opacity: g.opacity } : undefined;
        // 文字の下地：テキストボックスを、画面で測った文字の塊に下地の margin を足した大きさにし、塗りつぶす。
        // margin はテキストボックスの内側の余白にするので、文字の位置は画面と同じ。測っていなければ、ボックス全体を塗る
        const tb = e.text_bg;
        const m = tb ? (opts.textBgs || []).find((x) => x.slide === s.id && x.path === e.path) : null;
        const mg = tb && m ? tb.margin * k : 0;
        const tbox = m ? { x: m.rect.x * k - mg, y: m.rect.y * k - mg, w: m.rect.w * k + 2 * mg, h: m.rect.h * k + 2 * mg } : box;
        slide.addText(textRuns(e.text, f, e.para_space, glow, f.size, e.line_height), {
          ...(glow ? { glow } : {}),
          ...tbox,
          ...(tb ? { fill: { color: hex(tb.color, 'FFFF00'), transparency: Math.round((1 - tb.opacity) * 100) } } : {}),
          // 英語用と日本語用を「|」でつないで渡し、書き出した後で分ける（finishPPTX）
          fontFace: `${pptxFontName(f.family || DEFAULT_FONT)}|${pptxFontName(f.family_ja || DEFAULT_FONT_JA)}`,
          fontSize: f.size,
          bold: !!f.bold,
          italic: !!f.italic,
          color: hex(f.color, '000000'),
          align: e.align,
          valign: m ? 'top' : e.valign,
          margin: Math.round(mg * 72 * 100) / 100, // pt
          // 行の間隔は、倍数ではなくポイントで決める（文字の大きさ × 行間）。PowerPoint の倍数は、フォント本来の行の高さに掛けるため、
          // 画面（文字の大きさ × 行間）より行の間隔が広くなり、行ごとに文字が下にずれる（画像の線などの位置が合わなくなる）
          lineSpacing: Math.round(f.size * e.line_height * 100) / 100,
          wrap: m ? m.wrap : true,
          fit: 'none',
          objectName: e.path,
        });
        // 画像の線：画面で測った位置に、画像を縦横比を保たずに置く
        for (const u of (opts.underlines || []).filter((x) => x.slide === s.id && x.path === e.path)) {
          slide.addImage({ ...(await imageRef(u.src)), x: u.rect.x * k, y: u.rect.y * k, w: u.rect.w * k, h: u.rect.h * k, altText: `${e.path} の下線` });
        }
      }
    }
  }
  return pres;
}

// 飾りも改行も無い文字はそのまま。あれば、区切り（run）の配列にする
// 箇条書きの段落は、行頭に「•」を付け、折り返した行を文字の頭に揃える（字下げは文字の大きさの 1.1 倍。画面と同じ）
function textRuns(text, f, paraSpace, glow, size = 24, lineHeight = 1.2) {
  if (!hasRich(text)) return glow ? [{ text, options: { glow } }] : text;
  const runs = [];
  const paras = paragraphs(text);
  paras.forEach((p, pi) => {
    const start = runs.length;
    let soft = false;
    for (const g of parseRich(p)) {
      if (g.br) {
        soft = true;
        continue;
      }
      const options = {};
      if (glow) options.glow = glow;
      if (g.b || f.bold) options.bold = true;
      if (g.i || f.italic) options.italic = true;
      if (g.color) options.color = g.color;
      if (g.size) options.fontSize = g.size;
      if (g.u && g.u.type !== 'img') options.underline = { style: g.u.type === 'double' ? 'dbl' : 'sng' };
      if (g.sup) options.superscript = true;
      else if (g.sub) options.subscript = true;
      if (soft) options.softBreakBefore = true;
      soft = false;
      runs.push({ text: g.text, options });
    }
    if (soft) runs.push({ text: '', options: { softBreakBefore: true } });
    if (runs.length === start) runs.push({ text: '', options: {} }); // 空の段落
    if (pi > 0 && paraSpace > 0) runs[start].options.paraSpaceBefore = paraSpace;
    // 一部の文字の大きさを変えた段落は、いちばん大きい文字に合わせて行の間隔を決める（画面と同じ）
    const big = Math.max(0, ...parseRich(p).map((g) => g.size || 0));
    if (big > size) for (let i = start; i < runs.length; i++) runs[i].options.lineSpacing = Math.round(big * lineHeight * 100) / 100;
    // 不可視の箇条書きは、行頭の記号を空白（見えない文字）にして、字下げだけを残す
    const bt = bulletType(p);
    if (bt) runs[start].options.bullet = { indent: Math.round(size * 1.1 * 10) / 10, ...(bt === 'hidden' ? { characterCode: '00A0' } : {}) };
    if (pi < paras.length - 1) runs[runs.length - 1].options.breakLine = true;
  });
  return runs;
}

/**
 * pptxgenjs で作ったものを書き出し、文字のフォントを英語用と日本語用に分けて書き直す。
 * @param JSZip  Node では jszip、ブラウザでは window.JSZip
 * @param type   'nodebuffer' / 'blob' / 'arraybuffer'
 */
export async function finishPPTX(pres, JSZip, type) {
  const buf = await pres.write({ outputType: 'arraybuffer' });
  const zip = await JSZip.loadAsync(buf);
  for (const name of Object.keys(zip.files)) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
    let xml = await zip.file(name).async('string');
    // 文字のフォント：英語用（latin・cs）と日本語用（ea）に分ける。ea の文字コードの指定は日本語（Shift_JIS）にする
    xml = xml.replace(/<a:(latin|ea|cs) typeface="([^"|]*)\|([^"]*)"([^>]*?)\/>/g, (m, kind, latin, ja, rest) =>
      kind === 'ea' ? `<a:ea typeface="${ja}"${rest.replace(/charset="[^"]*"/, 'charset="-128"')}/>` : `<a:${kind} typeface="${latin}"${rest}/>`);
    xml = xml.replace(/typeface="([^"|]*)\|[^"]*"/g, 'typeface="$1"');
    zip.file(name, xml);
  }
  return zip.generateAsync({ type, compression: 'DEFLATE' });
}
