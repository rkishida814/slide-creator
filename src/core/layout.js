// 配置の計算。layout.yaml（テンプレート展開済み）を受け取り、各ノードの枠と、
// 描く要素（rect / image / text）の絶対座標をキャンバスの単位で返す。
// ファイルやブラウザに依存しない（ブラウザ版でもそのまま使う）。
import { assertUnit, toUnit, parseBox, defaultPadding, convert } from './units.js';

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
  if (node.layout != null) return 'layout';
  if (node.image != null) return 'image';
  if (node.text != null) return 'text';
  return 'empty';
}

/**
 * @param {object} doc  展開済みの文書
 * @param {object} env  { imageSize(src) => {w,h} | null,
 *                        layout(file) => 展開済みの文書 | null（ボックスに入れる別のレイアウト。任意） }
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
    pushBackground(slide.background ?? doc.background, { x: 0, y: 0, w: W, h: H }, id, ctx);
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
  const pad = parseBox(node.padding ?? defaultPadding(kind), ctx.unit);
  const inner = { x: rect.x + pad.l, y: rect.y + pad.t, w: rect.w - pad.l - pad.r, h: rect.h - pad.t - pad.b };
  // background_inset: true なら、背景と枠線も余白の分だけ内側に縮める
  const bgRect = node.background_inset && inner.w > 0 && inner.h > 0 ? inner : rect;
  pushBackground(node.background, bgRect, path, ctx);
  if (node.border) ctx.elements.push({ type: 'rect', path, rect: bgRect, fill: null, border: node.border });
  if (inner.w <= 0 || inner.h <= 0) {
    ctx.issues.push({ level: 'error', path, msg: 'padding が大きすぎて中身の領域が無い' });
    return;
  }

  if (kind === 'split') return layoutSplit(node, inner, style, path, ctx, depth);
  if (kind === 'grid') return layoutGrid(node, inner, style, path, ctx, depth);
  if (kind === 'layout') return layoutNested(node, inner, style, path, ctx);
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

// 背景。色だけなら "#RRGGBB"。画像を使うときは
//   { image: 画像のファイル, mode: tile（並べる・既定）| cover（広げる）, tile: 模様 1 つの幅, color: 下地の色 }
// tile を省くと、画像を 96 dpi として測った幅にする
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
  if (bg.mode === 'cover') {
    // 枠いっぱいに広げ、はみ出しを中央から切る
    const ia = nat.w / nat.h;
    const ba = rect.w / rect.h;
    let crop = [0, 0, 1, 1];
    if (ia > ba) {
      const keep = ba / ia;
      crop = [(1 - keep) / 2, 0, (1 + keep) / 2, 1];
    } else {
      const keep = ia / ba;
      crop = [0, (1 - keep) / 2, 1, (1 + keep) / 2];
    }
    const dpi = ((crop[2] - crop[0]) * nat.w) / (rect.w / perInch(ctx.unit));
    ctx.elements.push({ type: 'image', path, src: bg.image, rect, cell: rect, crop, natural: nat, dpi: Math.round(dpi), background: true });
    return;
  }
  const tileW = bg.tile != null ? toUnit(bg.tile, ctx.unit) : (nat.w / 96) * perInch(ctx.unit);
  if (!(tileW > 0)) {
    ctx.issues.push({ level: 'error', path, msg: `背景の模様の幅が不正: ${bg.tile}` });
    return;
  }
  ctx.elements.push({ type: 'texture', path, rect, src: bg.image, tileW, natural: nat });
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
  });
}

// ---------------------------------------------------------------- 別のレイアウトを入れる
// layout: fig_intro（フォルダ。中の layout.yaml を使う）か fig_intro/layout.yaml。slide: s1（省くと 1 枚目）。
// 入れるレイアウトの 1 枚を、縦横比を保ったままボックスに収める（揃え方は align / valign。既定は中央）。
// 長さも文字の大きさも同じ倍率で縮める（図として貼るのと同じ見た目）。中の画像のファイルは、入れるレイアウトの
// フォルダからの場所に直す。中のボックスはこの文書からは選べない（直すときは、入れるレイアウトのほうを開く）

// "a/b" + "../c.png" → "a/c.png"（区切りは /）
export function joinPath(dir, file) {
  const f = String(file);
  if (!dir || /^([a-z]+:|\/)/i.test(f)) return f;
  const out = [];
  for (const part of `${dir}/${f}`.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..' && out.length && out[out.length - 1] !== '..') out.pop();
    else out.push(part);
  }
  return out.join('/');
}

// layout の値 → 読むファイル（フォルダなら layout.yaml を足す）
export function layoutFileOf(ref) {
  const r = String(ref).replace(/\\/g, '/').replace(/\/+$/, '');
  return /\.ya?ml$/i.test(r) ? r : `${r}/layout.yaml`;
}

export function dirOf(file) {
  const i = file.lastIndexOf('/');
  return i < 0 ? '' : file.slice(0, i);
}

function layoutNested(node, rect, style, path, ctx) {
  const file = layoutFileOf(node.layout);
  const stack = ctx.env.stack || [];
  const fail = (msg) => {
    ctx.issues.push({ level: 'error', path, msg });
    ctx.elements.push({ type: 'missing', path, rect, src: file });
  };
  if (stack.includes(file)) return fail(`レイアウトが自分自身を入れている: ${[...stack, file].join(' → ')}`);
  const sub = ctx.env.layout ? ctx.env.layout(file) : null;
  if (!sub) return fail(`レイアウトが読めない: ${file}`);
  if (sub.error) return fail(`レイアウト ${file} の誤り: ${sub.error}`);
  const dir = dirOf(file);
  const subEnv = {
    imageSize: (src) => ctx.env.imageSize(joinPath(dir, src)),
    layout: (f) => (ctx.env.layout ? ctx.env.layout(joinPath(dir, f)) : null),
    stack: [...stack, file],
  };
  let L;
  try {
    L = layoutDoc(sub, subEnv);
  } catch (e) {
    return fail(`レイアウト ${file} の誤り: ${e.message}`);
  }
  const s = node.slide != null ? L.slides.find((x) => x.id === String(node.slide)) : L.slides[0];
  if (!s) return fail(`レイアウト ${file} にスライド ${node.slide ?? ''} が無い`);
  const align = style.align || 'center';
  const valign = style.valign || 'middle';
  checkAlign(align, valign, path, ctx);
  // 入れるレイアウトの大きさ（この文書の単位）と、縮める倍率 k
  const cu = convert(1, L.unit, ctx.unit);
  const k = Math.min(rect.w / (L.width * cu), rect.h / (L.height * cu));
  const w = L.width * cu * k;
  const h = L.height * cu * k;
  const ox = rect.x + offset(rect.w - w, align, H_ALIGN);
  const oy = rect.y + offset(rect.h - h, valign, V_ALIGN);
  const c = cu * k;
  const tr = (r) => r && { x: ox + r.x * c, y: oy + r.y * c, w: r.w * c, h: r.h * c };
  const pre = (p) => `${path}>${p}`;
  for (const e of s.elements) {
    const o = { ...e, path: pre(e.path), rect: tr(e.rect) };
    if (e.cell) o.cell = tr(e.cell);
    if (e.src && e.type !== 'missing') o.src = joinPath(dir, e.src);
    if (e.tileW != null) o.tileW = e.tileW * c;
    if (e.border) o.border = { ...e.border, width: (e.border.width ?? 1) * k };
    if (e.type === 'text') {
      const f = { ...e.font, size: e.font.size * k };
      if (f.glow) {
        const g = f.glow === true ? {} : f.glow;
        f.glow = { ...g, size: (g.size ?? 6) * k };
      }
      o.font = f;
      o.para_space = (e.para_space || 0) * k;
    }
    if (e.type === 'image') {
      o.dpi = Math.round(e.dpi / k);
      if (o.dpi < ctx.minDpi && !e.background) {
        ctx.issues.push({ level: 'warn', path: o.path, msg: `画像の解像度が低い: ${o.dpi} dpi（基準 ${ctx.minDpi}）${o.src}` });
      }
    }
    ctx.elements.push(o);
  }
  // 解像度の警告は、縮めた後の大きさで出し直した（上）。それ以外の指摘は引き継ぐ
  for (const i of s.issues) if (!/解像度/.test(i.msg)) ctx.issues.push({ ...i, path: pre(i.path) });
  const box = ctx.boxes[ctx.boxes.length - 1];
  if (box && box.path === path) Object.defineProperty(box, 'nested', { value: { file, slide: s.id, scale: k, rect: { x: ox, y: oy, w, h }, slides: L.slides.map((x) => x.id) } });
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
