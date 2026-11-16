// 配置の計算。layout.yaml（テンプレート展開済み）を受け取り、各ノードの枠と、
// 描く要素（rect / image / text）の絶対座標をキャンバスの単位で返す。
// ファイルやブラウザに依存しない（ブラウザ版でもそのまま使う）。
import { assertUnit, toUnit, parseBox } from './units.js';

// 親から子へ引き継ぐ設定
const INHERIT = ['align', 'valign', 'fit', 'font', 'line_height'];
const H_ALIGN = ['left', 'center', 'right'];
const V_ALIGN = ['top', 'middle', 'bottom'];
const FITS = ['contain', 'cover', 'stretch'];

export function nodeKind(node) {
  if (node == null) return 'empty';
  if (typeof node === 'string') return 'text';
  if (node.split) return 'split';
  if (node.grid) return 'grid';
  if (node.image != null) return 'image';
  if (node.text != null) return 'text';
  return 'empty';
}

/**
 * @param {object} doc  展開済みの文書
 * @param {object} env  { imageSize(src) => {w,h} | null }
 */
export function layoutDoc(doc, env) {
  if (!doc || !doc.canvas) throw new Error('canvas がありません');
  const unit = assertUnit(doc.canvas.unit || 'cm');
  const W = toUnit(doc.canvas.width, unit);
  const H = toUnit(doc.canvas.height, unit);
  if (!(W > 0 && H > 0)) throw new Error('canvas の width / height が不正です');
  const minDpi = doc.check?.min_dpi ?? 150;
  const issues = [];
  // PowerPoint のスライドは 1 辺 56 インチ（142.24 cm）まで
  const maxIn = 56;
  for (const [name, v] of [['幅', W], ['高さ', H]]) {
    const inch = v / { in: 1, cm: 2.54, mm: 25.4, pt: 72, px: 96 }[unit];
    if (inch > maxIn) issues.push({ level: 'warn', path: 'canvas', msg: `${name} ${(inch * 2.54).toFixed(1)} cm は、PowerPoint のスライドの上限 142.24 cm を超える。PPTX は PowerPoint で正しく開けないおそれがある（PDF には書き出せる）` });
  }
  const slides = (doc.slides || []).map((slide, i) => {
    const id = slide.id || `s${i + 1}`;
    const ctx = { unit, minDpi, env, boxes: [], elements: [], issues: [], ids: new Set(), slideId: id };
    const inh = pickInherit(doc.defaults || {}, {});
    if (slide.background) ctx.elements.push({ type: 'rect', path: id, rect: { x: 0, y: 0, w: W, h: H }, fill: slide.background });
    walk(slide.root, { x: 0, y: 0, w: W, h: H }, inh, `${id}`, ctx, 0, null);
    return { id, rect: { x: 0, y: 0, w: W, h: H }, boxes: ctx.boxes, elements: ctx.elements, issues: ctx.issues };
  });
  const ids = slides.map((s) => s.id);
  const dup = ids.filter((x, i) => ids.indexOf(x) !== i);
  if (dup.length) throw new Error(`スライドの id が重複: ${dup.join(', ')}`);
  return { unit, width: W, height: H, font: doc.defaults?.font, slides, issues };
}

function pickInherit(node, base) {
  const out = { ...base };
  for (const k of INHERIT) {
    if (node[k] == null) continue;
    out[k] = k === 'font' ? { ...(base.font || {}), ...node.font } : node[k];
  }
  return out;
}

function walk(node, rect, inherited, autoPath, ctx, depth, parent) {
  const kind = nodeKind(node);
  if (typeof node === 'string') node = { text: node };
  node = node || {};
  const path = node.id ? `${ctx.slideId}/${node.id}` : autoPath;
  if (node.id) {
    if (ctx.ids.has(node.id)) ctx.issues.push({ level: 'error', path, msg: `id "${node.id}" が同じスライド内で重複` });
    ctx.ids.add(node.id);
  }
  const box = { path, id: node.id || null, kind, rect, depth, parent, src: node.__src || null, template: node.__template || null };
  ctx.boxes.push(box);

  const style = pickInherit(node, inherited);
  // 画面の設定欄で使う（JSON には出さない）
  Object.defineProperties(box, { node: { value: node }, inherited: { value: inherited }, style: { value: style } });
  if (node.background || node.border) {
    ctx.elements.push({ type: 'rect', path, rect, fill: node.background || null, border: node.border || null });
  }
  const pad = parseBox(node.padding, ctx.unit);
  const inner = { x: rect.x + pad.l, y: rect.y + pad.t, w: rect.w - pad.l - pad.r, h: rect.h - pad.t - pad.b };
  if (inner.w <= 0 || inner.h <= 0) {
    ctx.issues.push({ level: 'error', path, msg: 'padding が大きすぎて中身の領域が無い' });
    return;
  }

  if (kind === 'split') return layoutSplit(node, inner, style, path, ctx, depth);
  if (kind === 'grid') return layoutGrid(node, inner, style, path, ctx, depth);
  if (kind === 'image') return layoutImage(node, inner, style, path, ctx);
  if (kind === 'text') return layoutText(node, inner, style, path, ctx);
}

// split: rows（上下に並べる）/ cols（左右に並べる）。
// ratio の数値は比率、"3cm" のような長さは固定長。
function layoutSplit(node, rect, style, path, ctx, depth) {
  const dir = node.split;
  if (dir !== 'rows' && dir !== 'cols') {
    ctx.issues.push({ level: 'error', path, msg: `split は rows か cols（"${dir}"）` });
    return;
  }
  const children = node.children || [];
  const n = children.length;
  if (!n) return;
  const ratio = node.ratio || children.map(() => 1);
  if (ratio.length !== n) {
    ctx.issues.push({ level: 'error', path, msg: `ratio の数（${ratio.length}）と children の数（${n}）が違う` });
    return;
  }
  const total = dir === 'rows' ? rect.h : rect.w;
  const gap = toUnit(node.gap ?? 0, ctx.unit);
  let free = total - gap * (n - 1);
  let sumRatio = 0;
  const fixed = ratio.map((r) => (typeof r === 'number' ? null : toUnit(r, ctx.unit)));
  fixed.forEach((f, i) => (f == null ? (sumRatio += ratio[i]) : (free -= f)));
  if (free < -1e-9) ctx.issues.push({ level: 'error', path, msg: '固定長と gap の合計が領域を超える' });
  let pos = dir === 'rows' ? rect.y : rect.x;
  children.forEach((child, i) => {
    const len = fixed[i] != null ? fixed[i] : sumRatio > 0 ? (Math.max(free, 0) * ratio[i]) / sumRatio : 0;
    const r = dir === 'rows' ? { x: rect.x, y: pos, w: rect.w, h: len } : { x: pos, y: rect.y, w: len, h: rect.h };
    walk(child, r, style, `${path}.${i}`, ctx, depth + 1, path);
    pos += len + gap;
  });
}

// grid: rows × cols の同じ大きさのセル。items は行ごと（order: col なら列ごと）に詰める。
function layoutGrid(node, rect, style, path, ctx, depth) {
  const g = node.grid;
  const rows = g.rows || 1;
  const cols = g.cols || 1;
  const gapv = g.gap ?? node.gap ?? 0;
  const [rg, cg] = Array.isArray(gapv) ? gapv.map((v) => toUnit(v, ctx.unit)) : [toUnit(gapv, ctx.unit), toUnit(gapv, ctx.unit)];
  const items = node.items || [];
  if (items.length > rows * cols) {
    ctx.issues.push({ level: 'warn', path, msg: `items が ${items.length} 個あり、セル ${rows * cols} 個を超えた分は置かれない` });
  }
  const cw = (rect.w - cg * (cols - 1)) / cols;
  const ch = (rect.h - rg * (rows - 1)) / rows;
  const byCol = g.order === 'col';
  for (let k = 0; k < Math.min(items.length, rows * cols); k++) {
    const r = byCol ? k % rows : Math.floor(k / cols);
    const c = byCol ? Math.floor(k / rows) : k % cols;
    const cell = { x: rect.x + c * (cw + cg), y: rect.y + r * (ch + rg), w: cw, h: ch };
    walk(items[k], cell, style, `${path}.${k}`, ctx, depth + 1, path);
  }
}

function layoutImage(node, rect, style, path, ctx) {
  const src = node.image;
  const fit = node.fit || style.fit || 'contain';
  const align = style.align || 'center';
  const valign = style.valign || 'middle';
  if (!FITS.includes(fit)) ctx.issues.push({ level: 'error', path, msg: `fit は ${FITS.join('/')}（"${fit}"）` });
  checkAlign(align, valign, path, ctx);
  const nat = ctx.env.imageSize(src);
  if (!nat) {
    ctx.issues.push({ level: 'error', path, msg: `画像が読めない: ${src}` });
    ctx.elements.push({ type: 'missing', path, rect, src });
    return;
  }
  let [l, t, r, b] = node.crop || [0, 0, 1, 1];
  if (!(0 <= l && l < r && r <= 1 && 0 <= t && t < b && b <= 1)) {
    ctx.issues.push({ level: 'error', path, msg: `crop は [左, 上, 右, 下] の割合（0〜1、左<右、上<下）: ${JSON.stringify(node.crop)}` });
    [l, t, r, b] = [0, 0, 1, 1];
  }
  const srcW = (r - l) * nat.w; // 切り抜き後の画素数
  const srcH = (b - t) * nat.h;
  const imgAspect = srcW / srcH;
  const boxAspect = rect.w / rect.h;
  let disp = { ...rect };
  if (fit === 'contain') {
    if (imgAspect > boxAspect) disp.h = rect.w / imgAspect;
    else disp.w = rect.h * imgAspect;
    disp.x = rect.x + offset(rect.w - disp.w, align, H_ALIGN);
    disp.y = rect.y + offset(rect.h - disp.h, valign, V_ALIGN);
  } else if (fit === 'cover') {
    // 枠いっぱいに表示し、はみ出す分をさらに切り抜く（どちら側を残すかは align / valign）
    if (imgAspect > boxAspect) {
      const keep = (r - l) * (boxAspect / imgAspect);
      const cut = (r - l) - keep;
      l = l + offset(cut, align, H_ALIGN);
      r = l + keep;
    } else {
      const keep = (b - t) * (imgAspect / boxAspect);
      const cut = (b - t) - keep;
      t = t + offset(cut, valign, V_ALIGN);
      b = t + keep;
    }
  }
  const dispIn = disp.w / perInch(ctx.unit);
  const dpi = ((r - l) * nat.w) / dispIn;
  if (dpi < ctx.minDpi) {
    ctx.issues.push({ level: 'warn', path, msg: `画像の解像度が低い: ${Math.round(dpi)} dpi（基準 ${ctx.minDpi}）${src}` });
  }
  ctx.elements.push({ type: 'image', path, src, rect: disp, cell: rect, crop: [l, t, r, b], natural: nat, dpi: Math.round(dpi) });
}

function layoutText(node, rect, style, path, ctx) {
  const align = style.align || 'left';
  const valign = style.valign || 'top';
  checkAlign(align, valign, path, ctx);
  const font = { size: 18, ...(style.font || {}) };
  ctx.elements.push({
    type: 'text',
    path,
    rect,
    text: String(node.text),
    font,
    align,
    valign,
    line_height: style.line_height ?? 1.2,
  });
}

function checkAlign(align, valign, path, ctx) {
  if (!H_ALIGN.includes(align)) ctx.issues.push({ level: 'error', path, msg: `align は ${H_ALIGN.join('/')}（"${align}"）` });
  if (!V_ALIGN.includes(valign)) ctx.issues.push({ level: 'error', path, msg: `valign は ${V_ALIGN.join('/')}（"${valign}"）` });
}

function offset(space, a, list) {
  const i = list.indexOf(a);
  return i <= 0 ? 0 : i === 1 ? space / 2 : space;
}

function perInch(unit) {
  return { in: 1, cm: 2.54, mm: 25.4, pt: 72, px: 96 }[unit];
}
