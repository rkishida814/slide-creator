// 別に作ったレイアウトの 1 枚を、ボックスの中身として取り込む（ボックスの木を写す）。
// 取り込んだ後は、この文書の普通のボックスになる（分割・比率・文字などをこの文書の中で直す。元のファイルとはつながらない）。
// ファイルに依存しない。Node・ブラウザの両方から使う。
import { toUnit, convert } from './units.js';

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

export function dirOf(file) {
  const i = file.lastIndexOf('/');
  return i < 0 ? '' : file.slice(0, i);
}

// "fig_intro" → "fig_intro/layout.yaml"
export function layoutFileOf(ref) {
  const r = String(ref).replace(/\\/g, '/').replace(/\/+$/, '');
  return /\.ya?ml$/i.test(r) ? r : `${r}/layout.yaml`;
}

const INHERIT = ['align', 'valign', 'fit', 'font', 'line_height', 'para_space'];

/**
 * @param sub      取り込むレイアウト（テンプレート展開済み。resolveDoc の結果）
 * @param opts     { slide: id（省くと 1 枚目）, dir: 取り込むレイアウトのフォルダ（この文書のフォルダからの場所）,
 *                   unit: この文書の単位, rect: {w, h} 取り込む先のボックスの中身の大きさ（この文書の単位）,
 *                   usedIds: この文書のスライドで使っている id（重なる id は外す） }
 * @returns { node, scale, dropped: [id] }  node は layout.yaml にそのまま書ける形
 */
export function importLayout(sub, opts) {
  const slides = sub.slides || [];
  const s = opts.slide != null ? slides.find((x) => x.id === String(opts.slide)) : slides[0];
  if (!s) throw new Error(`取り込むレイアウトにスライド ${opts.slide ?? ''} が無い`);
  const su = sub.canvas?.unit || 'cm';
  const sw = toUnit(sub.canvas.width, su);
  const sh = toUnit(sub.canvas.height, su);
  const u = opts.unit;
  const cu = convert(1, su, u);
  // 縦横比を保って収める倍率。ほぼ 1 なら 1 にする（数値を元のままにする）
  let k = opts.rect ? Math.min(opts.rect.w / (sw * cu), opts.rect.h / (sh * cu)) : 1;
  if (Math.abs(k - 1) < 0.01) k = 1;
  const r4 = (v) => +v.toFixed(4);
  const r2 = (v) => +v.toFixed(2);
  // 長さ：数値は取り込むレイアウトの単位、文字列は単位付き → この文書の単位に直して k 倍
  const len = (v) => r4(toUnit(v, su) * cu * k);
  const lenStr = (v) => `${r4(toUnit(v, su) * cu * k)}${u}`;
  const used = new Set(opts.usedIds || []);
  const dropped = [];

  const fontOf = (f) => {
    if (!f || typeof f !== 'object') return f;
    const o = { ...f };
    if (o.size != null) o.size = r2(o.size * k);
    if (o.glow && k !== 1) {
      const g = o.glow === true ? {} : o.glow;
      o.glow = { ...g, size: r2((g.size ?? 6) * k) };
    }
    return o;
  };
  const bgOf = (bg) => {
    if (!bg || typeof bg !== 'object') return bg;
    const o = { ...bg };
    if (o.image) o.image = joinPath(opts.dir, o.image);
    if (o.tile != null) o.tile = len(o.tile);
    return o;
  };
  const boxOf = (p) => (Array.isArray(p) ? p.map(len) : len(p));

  const conv = (n) => {
    if (n == null) return {};
    if (typeof n !== 'object') return { text: String(n) };
    const o = {};
    for (const [key, v] of Object.entries(n)) {
      if (key.startsWith('__') || v === undefined) continue;
      o[key] = v;
    }
    if (o.id != null) {
      if (used.has(String(o.id))) {
        dropped.push(String(o.id));
        delete o.id;
      } else used.add(String(o.id));
    }
    if (o.image != null) o.image = joinPath(opts.dir, o.image);
    if (o.icon != null) o.icon = joinPath(opts.dir, o.icon);
    if (o.icon_size != null) o.icon_size = r2(o.icon_size * k);
    if (o.background != null) o.background = bgOf(o.background);
    if (o.border) o.border = { ...o.border, width: r2((o.border.width ?? 1) * k) };
    if (o.font) o.font = fontOf(o.font);
    if (o.para_space != null) o.para_space = r2(o.para_space * k);
    if (o.padding != null) o.padding = boxOf(o.padding);
    if (o.gap != null) o.gap = boxOf(o.gap);
    if (Array.isArray(o.ratio)) o.ratio = o.ratio.map((x) => (typeof x === 'number' ? x : lenStr(x)));
    if (o.grid && typeof o.grid === 'object') {
      o.grid = { ...o.grid };
      if (o.grid.gap != null) o.grid.gap = boxOf(o.grid.gap);
    }
    if (Array.isArray(o.children)) o.children = o.children.map(conv);
    if (Array.isArray(o.items)) o.items = o.items.map(conv);
    return o;
  };

  const root = conv(s.root);
  // 取り込むレイアウト全体の既定（defaults）と背景は、根に書く（この文書の既定より優先させる）
  const d = sub.defaults || {};
  for (const key of INHERIT) {
    if (d[key] == null) continue;
    if (key === 'font') root.font = { ...fontOf(d.font), ...(root.font || {}) };
    else if (key === 'para_space') root.para_space ??= r2(d.para_space * k);
    else root[key] ??= d[key];
  }
  const bg = s.background ?? sub.background;
  if (bg != null && root.background == null) root.background = bgOf(bg);
  return { node: root, scale: k, dropped };
}

// 取り込む先のノード：枠の設定（id・余白・背景・枠線）があれば残し、その中に取り込んだ木を入れる
export function placeImported(target, imported) {
  const keep = {};
  if (target && typeof target === 'object') {
    for (const key of ['id', 'padding', 'background', 'background_inset', 'border']) if (target[key] != null) keep[key] = target[key];
  }
  if (!Object.keys(keep).length) return imported;
  return { ...keep, split: 'rows', children: [imported] };
}
