// 配置結果 → 編集できる PPTX。
// 画像は元画像をそのまま埋め込み、トリミングは PowerPoint の「トリミング」（srcRect）で表す。
// 文字はテキストボックス（余白0・自動調整なし）。
// ファイルに依存しない。Node・ブラウザの両方から使う。
import { PER_INCH } from './units.js';
import { DEFAULT_FONT, DEFAULT_FONT_JA, pptxFontName, glowOf } from './html.js';
import { parseRich, hasRich, paragraphs } from './richtext.js';

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
 * @returns pptxgenjs のインスタンス（writeFile / write で書き出す）
 */
export async function buildPPTX(PptxGen, layout, imageRef) {
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
        slide.addText(textRuns(e.text, f, e.para_space, glow), {
          ...(glow ? { glow } : {}),
          ...box,
          // 英語用と日本語用を「|」でつないで渡し、書き出した後で分ける（finishPPTX）
          fontFace: `${pptxFontName(f.family || DEFAULT_FONT)}|${pptxFontName(f.family_ja || DEFAULT_FONT_JA)}`,
          fontSize: f.size,
          bold: !!f.bold,
          italic: !!f.italic,
          color: hex(f.color, '000000'),
          align: e.align,
          valign: e.valign,
          margin: 0,
          lineSpacingMultiple: e.line_height,
          wrap: true,
          fit: 'none',
          objectName: e.path,
        });
      }
    }
  }
  return pres;
}

// 飾りも改行も無い文字はそのまま。あれば、区切り（run）の配列にする
function textRuns(text, f, paraSpace, glow) {
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
      if (g.sup) options.superscript = true;
      else if (g.sub) options.subscript = true;
      if (soft) options.softBreakBefore = true;
      soft = false;
      runs.push({ text: g.text, options });
    }
    if (soft) runs.push({ text: '', options: { softBreakBefore: true } });
    if (runs.length === start) runs.push({ text: '', options: {} }); // 空の段落
    if (pi > 0 && paraSpace > 0) runs[start].options.paraSpaceBefore = paraSpace;
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
