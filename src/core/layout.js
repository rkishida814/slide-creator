// 配置の計算。layout.yaml（テンプレート展開済み）を受け取り、各ノードの枠と、
// 描く要素（rect / image / text）の絶対座標をキャンバスの単位で返す。
// ファイルやブラウザに依存しない（ブラウザ版でもそのまま使う）。
import { assertUnit, toUnit, parseBox, defaultPadding, isBgTree } from './units.js';

// 親から子へ引き継ぐ設定
const INHERIT = ['align', 'valign', 'fit', 'font', 'line_height', 'para_space'];
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
    const ctx = { unit, minDpi, env, boxes: [], bgBoxes: {}, elements: [], issues: [], ids: new Set(), slideId: id };
    const inh = pickInherit(doc.defaults || {}, {});
    pushBackground(slide.background ?? doc.background, { x: 0, y: 0, w: W, h: H }, id, ctx);
    walk(slide.root, { x: 0, y: 0, w: W, h: H }, inh, `${id}`, ctx, 0, null);
    return { id, rect: { x: 0, y: 0, w: W, h: H }, boxes: ctx.boxes, bgBoxes: ctx.bgBoxes, elements: ctx.elements, issues: ctx.issues };
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
  // 背景を分けた木の中は、決まった余白を付けない（隙間なく敷く）
  const pad = parseBox(node.padding ?? (ctx.bg ? null : defaultPadding(kind)), ctx.unit);
  const inner = { x: rect.x + pad.l, y: rect.y + pad.t, w: rect.w - pad.l - pad.r, h: rect.h - pad.t - pad.b };
  // background_inset: true なら、背景と枠線も余白の分だけ内側に縮める。
  // 背景を分けた木の中の枠は、余白を書けば、いつもその分だけ背景を内側に縮める
  const bgRect = (node.background_inset || ctx.bg) && inner.w > 0 && inner.h > 0 ? inner : rect;
  if (isBgTree(node.background)) layoutBgTree(node.background, bgRect, path, ctx);
  else pushBackground(node.background, bgRect, path, ctx);
  if (node.border) ctx.elements.push({ type: 'rect', path, rect: bgRect, fill: null, border: node.border });
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

// 背景を分けた木。ボックスと同じ配置の計算で、背景だけを描く。
// 分けた枠は slide.bgBoxes[ボックスの path] に入れる（画面の「背景を編集」で選ぶ枠）。中身のボックスの一覧には入れない
function layoutBgTree(tree, rect, host, ctx) {
  const sub = { ...ctx, boxes: [], bg: true };
  walk(tree, rect, {}, `${host}/bg`, sub, 0, null);
  ctx.bgBoxes[host] = sub.boxes;
}

// 背景。色だけなら "#RRGGBB"。画像を使うときは
//   { image: 画像のファイル, mode: stretch | cover, color: 下地の色 }
//   mode: stretch（既定）… 縦横比を変えて、ボックスいっぱいに広げる
//         cover … 縦横比を保ってボックスいっぱいに広げ、はみ出しを中央から切る
//   cover 以外の mode は stretch として描く
function pushBackground(bg, rect, path, ctx) {
  if (!bg) return;
  if (typeof bg !== 'object') {
    ctx.elements.push({ type: 'rect', path, rect, fill: bg });
    return;
  }
  if (bg.color) ctx.elements.push({ type: 'rect', path, rect, fill: bg.color });
  if (!bg.image) return;
  const nat = ctx.env.imageSize(bg.image);
  if (!nat) {
    ctx.issues.push({ level: 'error', path, msg: `背景の画像が読めない: ${bg.image}` });
    return;
  }
  // 背景の画像のトリミング：crop: [左, 上, 右, 下]（元画像に対する割合）。残した範囲を、mode に従って広げる
  const c0 = Array.isArray(bg.crop) && bg.crop.length === 4 && bg.crop.every((v) => typeof v === 'number') ? bg.crop : [0, 0, 1, 1];
  const cw = c0[2] - c0[0];
  const ch = c0[3] - c0[1];
  if (!(cw > 0 && ch > 0)) {
    ctx.issues.push({ level: 'error', path, msg: `背景の crop の範囲が無い: ${JSON.stringify(bg.crop)}` });
    return;
  }
  let crop = [...c0];
  if (bg.mode === 'cover') {
    // 枠いっぱいに広げ、はみ出しを中央から切る（トリミングで残した範囲の中で）
    const ia = (cw * nat.w) / (ch * nat.h);
    const ba = rect.w / rect.h;
    if (ia > ba) {
      const keep = ba / ia;
      crop = [c0[0] + (cw * (1 - keep)) / 2, c0[1], c0[0] + (cw * (1 + keep)) / 2, c0[3]];
    } else {
      const keep = ia / ba;
      crop = [c0[0], c0[1] + (ch * (1 - keep)) / 2, c0[2], c0[1] + (ch * (1 + keep)) / 2];
    }
  }
  const dpi = Math.min(((crop[2] - crop[0]) * nat.w) / (rect.w / perInch(ctx.unit)), ((crop[3] - crop[1]) * nat.h) / (rect.h / perInch(ctx.unit)));
  ctx.elements.push({ type: 'image', path, src: bg.image, rect, cell: rect, crop, natural: nat, dpi: Math.round(dpi), background: true });
}

function layoutText(node, rect, style, path, ctx) {
  const align = style.align || 'left';
  const valign = style.valign || 'top';
  checkAlign(align, valign, path, ctx);
  const font = { size: 24, ...(style.font || {}) };
  const lineHeight = style.line_height ?? 1.2;
  let textRect = rect;
  // 見出しアイコン：icon（画像のファイル）、icon_size（高さ。pt。省くと文字の大きさ）
  if (node.icon) {
    const nat = ctx.env.imageSize(node.icon);
    if (!nat) {
      ctx.issues.push({ level: 'error', path, msg: `見出しアイコンの画像が読めない: ${node.icon}` });
    } else {
      const pt = (v) => (v / 72) * perInch(ctx.unit); // pt → キャンバスの単位
      const ih = pt(node.icon_size ?? font.size);
      const iw = (ih * nat.w) / nat.h;
      const gap = pt(font.size * 0.3);
      const lineH = pt(font.size * lineHeight);
      let iy;
      if (valign === 'top') iy = rect.y + (lineH - ih) / 2;
      else if (valign === 'bottom') iy = rect.y + rect.h - lineH + (lineH - ih) / 2;
      else iy = rect.y + (rect.h - ih) / 2;
      const ir = { x: rect.x, y: iy, w: iw, h: ih };
      const dpi = nat.w / (iw / perInch(ctx.unit));
      ctx.elements.push({ type: 'image', path, src: node.icon, rect: ir, cell: ir, crop: [0, 0, 1, 1], natural: nat, dpi: Math.round(dpi), icon: true });
      textRect = { x: rect.x + iw + gap, y: rect.y, w: Math.max(rect.w - iw - gap, 0), h: rect.h };
    }
  }
  ctx.elements.push({
    type: 'text',
    path,
    rect: textRect,
    text: String(node.text),
    font,
    align,
    valign,
    line_height: lineHeight,
    para_space: style.para_space ?? 0,
    text_bg: textBg(node, font, path, ctx),
  });
}

// 文字の下地（text_bg）。文字の塊にぴったりの四角に、色を敷く（背景とは別。PPTX ではテキストボックスの塗りつぶし）。
// "#RRGGBB" か {color, opacity（0〜1）, margin（文字の塊の外に広げる長さ。省くと文字の大きさの 0.15 倍）}。無ければ null
function textBg(node, font, path, ctx) {
  const tb = node && typeof node === 'object' ? node.text_bg : null;
  if (!tb) return null;
  const o = typeof tb === 'string' ? { color: tb } : tb;
  if (!o.color) return null;
  const opacity = Math.min(1, Math.max(0, Number(o.opacity ?? 1)));
  let margin = (font.size * 0.15 * perInch(ctx.unit)) / 72;
  if (o.margin != null) {
    try {
      margin = Math.max(0, toUnit(o.margin, ctx.unit));
    } catch (e) {
      ctx.issues.push({ level: 'error', path, msg: `text_bg の margin が読めない: ${o.margin}` });
    }
  }
  return { color: String(o.color), opacity: Number.isFinite(opacity) ? opacity : 1, margin };
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
