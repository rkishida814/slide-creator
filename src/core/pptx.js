// 配置結果 → 編集できる PPTX。
// 画像は元画像をそのまま埋め込み、トリミングは PowerPoint の「トリミング」（srcRect）で表す。
// 文字はテキストボックス（余白0・自動調整なし）。
// ファイルに依存しない。Node・ブラウザの両方から使う。
import { PER_INCH } from './units.js';
import { DEFAULT_FONT } from './html.js';

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
        slide.addText(e.text, {
          ...box,
          fontFace: f.family || DEFAULT_FONT,
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
