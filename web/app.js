// slide-creator ブラウザ版
import { resolveDoc } from '../src/core/resolve.js';
import { layoutDoc } from '../src/core/layout.js';
import { renderHTML } from '../src/core/html.js';
import { buildPPTX } from '../src/core/pptx.js';
import { PER_INCH } from '../src/core/units.js';
import { FolderSource, UrlSource, download } from './source.js';
import { parse, stringify, setProp, replaceNode, getJS, YAML } from './edit.js';
import { openCrop } from './crop.js';

const $ = (s) => document.querySelector(s);
const IMAGE_RE = /\.(png|jpe?g|gif|svg|webp)$/i;

const S = {
  source: null,
  files: {}, // 'layout.yaml' / 'templates/x.yaml' → { text, doc, dirty }
  assets: [],
  layout: null,
  error: null,
  slideId: null,
  selected: null, // 選んだ枠の path
  focus: null, // 拡大している枠の path
  outline: true,
  imageSizes: new Map(),
  imageURLs: new Map(),
  overflow: [],
  view: null,
};

// ---------------------------------------------------------------- 読み込み

async function openSource(source) {
  if (anyDirty() && !confirm('保存していない変更があります。捨てて開きますか？')) return;
  try {
    status('読み込み中…');
    const files = {};
    files['layout.yaml'] = mkFile(await source.readText('layout.yaml'));
    for (const f of await source.list('templates', 0)) {
      if (/\.ya?ml$/.test(f)) files[f] = mkFile(await source.readText(f));
    }
    S.source = source;
    S.files = files;
    S.assets = (await source.list('assets')).filter((f) => IMAGE_RE.test(f));
    S.imageSizes.clear();
    S.imageURLs.clear();
    S.slideId = null;
    S.selected = null;
    S.focus = null;
    $('#project-name').textContent = source.name + (source.writable ? '' : '（見本）');
    document.title = `${source.name} — slide-creator`;
    $('#empty').hidden = true;
    for (const id of ['#btn-reload', '#btn-pptx', '#btn-pdf']) $(id).disabled = false;
    await rebuild();
    status('');
  } catch (e) {
    status(e.message, true);
  }
}

function mkFile(text) {
  return { text, doc: parse(text), dirty: false };
}

function anyDirty() {
  return Object.values(S.files).some((f) => f.dirty);
}

function collectImages(node, out) {
  if (!node || typeof node !== 'object') return;
  if (node.image != null) out.add(String(node.image));
  (node.children || []).forEach((c) => collectImages(c, out));
  (node.items || []).forEach((c) => collectImages(c, out));
}

async function loadImage(src) {
  if (S.imageSizes.has(src)) return;
  try {
    const url = await S.source.imageURL(src);
    const img = new Image();
    img.src = url;
    await img.decode();
    S.imageURLs.set(src, url);
    S.imageSizes.set(src, { w: img.naturalWidth, h: img.naturalHeight });
  } catch {
    S.imageSizes.set(src, null);
  }
}

// YAML → 展開 → 配置計算 → 表示
async function rebuild() {
  try {
    const templates = {};
    for (const [name, f] of Object.entries(S.files)) {
      if (name.startsWith('templates/')) templates[name.slice(10).replace(/\.ya?ml$/, '')] = f.doc.toJS();
    }
    const resolved = resolveDoc(S.files['layout.yaml'].doc.toJS() || {}, templates);
    const srcs = new Set();
    (resolved.slides || []).forEach((s) => collectImages(s.root, srcs));
    await Promise.all([...srcs].map(loadImage));
    S.layout = layoutDoc(resolved, { imageSize: (src) => S.imageSizes.get(src) || null });
    S.error = null;
  } catch (e) {
    S.error = e.message;
  }
  if (S.layout && !S.layout.slides.find((s) => s.id === S.slideId)) S.slideId = S.layout.slides[0]?.id || null;
  if (S.selected && !findBox(S.selected)) S.selected = null;
  if (S.focus && !findBox(S.focus)) S.focus = null;
  status(S.error || '', !!S.error);
  renderSlideTabs();
  renderPreview();
  renderGlobal();
  renderProps();
  $('#btn-save').disabled = !anyDirty();
}

// ---------------------------------------------------------------- 表示

function curSlide() {
  return S.layout?.slides.find((s) => s.id === S.slideId) || null;
}

function findBox(path) {
  const s = curSlide();
  return s?.boxes.find((b) => b.path === path) || null;
}

function renderSlideTabs() {
  const el = $('#slide-tabs');
  el.innerHTML = '';
  for (const s of S.layout?.slides || []) {
    const b = document.createElement('button');
    b.textContent = s.id;
    b.className = s.id === S.slideId ? 'on' : '';
    b.onclick = () => {
      S.slideId = s.id;
      S.selected = null;
      S.focus = null;
      renderSlideTabs();
      renderPreview();
      renderProps();
    };
    el.append(b);
  }
}

const FRAME_CSS = `
html { background: #dfe3e8; }
body { background: transparent !important; overflow: hidden; }
section.slide { position: absolute; left: 0; top: 0; box-shadow: 0 1px 6px rgba(0,0,0,.25); }
.sel { position: absolute; pointer-events: none; z-index: 10; }
.focusdim { position: absolute; pointer-events: none; z-index: 9; }
`;

function renderPreview() {
  const frame = $('#frame');
  const slide = curSlide();
  if (!slide) {
    frame.srcdoc = '';
    return;
  }
  const html = renderHTML(S.layout, {
    slides: [slide.id],
    outline: S.outline,
    srcURL: (src) => S.imageURLs.get(src) || '',
  }).replace('</style>', `${FRAME_CSS}</style>`);
  frame.onload = setupFrame;
  frame.srcdoc = html;
}

function setupFrame() {
  const frame = $('#frame');
  const doc = frame.contentDocument;
  const sec = doc?.querySelector('section.slide');
  if (!sec) return;
  sec.style.transform = 'none';
  S.ppu = sec.getBoundingClientRect().width / S.layout.width; // キャンバスの単位あたりの px
  for (const cls of ['focusdim', 'sel']) {
    const d = doc.createElement('div');
    d.className = cls;
    sec.append(d);
  }
  doc.addEventListener('click', (e) => selectAt(e, false));
  doc.addEventListener('dblclick', (e) => selectAt(e, true));
  doc.addEventListener('keydown', onKey);
  applyView();
  measureOverflow(doc);
}

// 拡大・位置合わせ
function applyView() {
  const frame = $('#frame');
  const doc = frame.contentDocument;
  const sec = doc?.querySelector('section.slide');
  const slide = curSlide();
  if (!sec || !slide) return;
  const fb = S.focus ? findBox(S.focus) : null;
  const r = fb ? fb.rect : slide.rect;
  const ppu = S.ppu;
  const m = 24;
  const fw = frame.clientWidth;
  const fh = frame.clientHeight;
  const s = Math.min((fw - 2 * m) / (r.w * ppu), (fh - 2 * m) / (r.h * ppu));
  const offX = (fw - r.w * ppu * s) / 2;
  const offY = (fh - r.h * ppu * s) / 2;
  sec.style.transformOrigin = '0 0';
  sec.style.transform = `translate(${offX}px, ${offY}px) scale(${s}) translate(${-r.x * ppu}px, ${-r.y * ppu}px)`;
  S.view = { s, offX, offY, r };

  const place = (el, rect) => {
    el.style.display = rect ? 'block' : 'none';
    if (!rect) return;
    Object.assign(el.style, { left: `${rect.x * ppu}px`, top: `${rect.y * ppu}px`, width: `${rect.w * ppu}px`, height: `${rect.h * ppu}px` });
  };
  const dim = sec.querySelector('.focusdim');
  place(dim, fb ? fb.rect : null);
  dim.style.boxShadow = `0 0 0 ${Math.max(S.layout.width, S.layout.height) * ppu}px rgba(40, 44, 52, .45)`;
  const sel = S.selected ? findBox(S.selected) : null;
  const selEl = sec.querySelector('.sel');
  place(selEl, sel ? sel.rect : null);
  selEl.style.outline = `${2.5 / s}px solid #2563eb`;
  selEl.style.background = 'rgba(37, 99, 235, .08)';
  renderCrumbs();
  $('#btn-focus').disabled = !S.selected || S.selected === S.focus;
  $('#btn-whole').disabled = !S.focus;
}

function selectAt(e, zoom) {
  const slide = curSlide();
  if (!slide || !S.view) return;
  const { s, offX, offY, r } = S.view;
  const x = (e.clientX - offX) / s / S.ppu + r.x;
  const y = (e.clientY - offY) / s / S.ppu + r.y;
  let best = null;
  for (const b of slide.boxes) {
    const q = b.rect;
    if (x >= q.x && x <= q.x + q.w && y >= q.y && y <= q.y + q.h && (!best || b.depth >= best.depth)) best = b;
  }
  if (!best) return;
  select(best.path);
  if (zoom) focusOn(best.path);
}

function select(path) {
  S.selected = path;
  applyView();
  renderProps();
}

function focusOn(path) {
  const b = findBox(path);
  S.focus = b && b.depth > 0 ? path : null;
  applyView();
}

function renderCrumbs() {
  const el = $('#crumbs');
  el.innerHTML = '';
  const sel = S.selected ? findBox(S.selected) : null;
  if (!sel) return;
  const chain = [];
  for (let b = sel; b; b = b.parent ? findBox(b.parent) : null) chain.unshift(b);
  chain.forEach((b, i) => {
    if (i) el.append(' › ');
    const a = document.createElement('a');
    a.textContent = b.id || b.path.split(/[./]/).pop() || b.path;
    if (b.depth === 0) a.textContent = b.path;
    a.className = b === sel ? 'cur' : '';
    a.title = b.path;
    a.onclick = () => select(b.path);
    el.append(a);
  });
}

function measureOverflow(doc) {
  const out = [];
  for (const box of doc.querySelectorAll('.txt')) {
    const inner = box.firstElementChild;
    const b = box.getBoundingClientRect();
    const i = inner.getBoundingClientRect();
    const overH = (i.height - b.height) / b.height;
    const overW = inner.scrollWidth - box.clientWidth > 0.5;
    if (overH > 0.005 || overW) {
      out.push({ level: 'error', path: box.dataset.path, msg: overW ? '文字が横にはみ出している（長い語）' : `文字が枠の高さを ${Math.round(overH * 100)}% 超えている` });
    }
  }
  S.overflow = out;
}

// ---------------------------------------------------------------- 編集

function fileChanged(file) {
  const f = S.files[file];
  f.text = stringify(f.doc);
  f.dirty = true;
  return rebuild();
}

// 選んだ枠の設定を 1 つ変える
function edit(box, key, value) {
  const { file, path } = box.src;
  try {
    setProp(S.files[file].doc, path, key, value);
  } catch (e) {
    return status(`書き換えられません: ${e.message}`, true);
  }
  return fileChanged(file);
}

function editMany(box, pairs) {
  const { file, path } = box.src;
  for (const [k, v] of pairs) setProp(S.files[file].doc, path, k, v);
  return fileChanged(file);
}

// 型から来た値を YAML に書き戻すときに、内部の印（__src など）を外す
function plain(node) {
  if (node == null || typeof node !== 'object') return node;
  if (Array.isArray(node)) return node.map(plain);
  if (node.__scalar) return node.text;
  const out = {};
  for (const [k, v] of Object.entries(node)) if (!k.startsWith('__')) out[k] = plain(v);
  return out;
}

const CONTENT_KEYS = ['split', 'children', 'ratio', 'grid', 'items', 'image', 'crop', 'text'];

function changeKind(box, kind) {
  const { file, path, mode } = box.src;
  const doc = S.files[file].doc;
  const fresh = {
    empty: {},
    split: { split: 'rows', children: [{}, {}] },
    grid: { grid: { rows: 2, cols: 2 }, items: [{}, {}, {}, {}] },
    image: { image: S.assets[0] || 'assets/' },
    text: { text: '（ここに文字を入力）' },
  }[kind];
  if (mode === 'direct') {
    const cur = getJS(doc, path);
    const keep = cur && typeof cur === 'object' ? Object.fromEntries(Object.entries(cur).filter(([k]) => !CONTENT_KEYS.includes(k))) : {};
    replaceNode(doc, path, { ...keep, ...fresh });
  } else {
    // 型の値を打ち消すため、他の種類の中身は null にする
    for (const k of CONTENT_KEYS) setProp(doc, path, k, k in fresh ? fresh[k] : null);
  }
  return fileChanged(file);
}

function wrapSplit(box, dir) {
  const { file, path } = box.src;
  const doc = S.files[file].doc;
  const cur = getJS(doc, path);
  replaceNode(doc, path, { split: dir, children: [cur ?? {}, {}] });
  S.selected = box.path; // 自動の経路は変わる。親の位置を選び直す
  return fileChanged(file);
}

function clearNode(box) {
  const { file, path, mode } = box.src;
  const doc = S.files[file].doc;
  if (mode === 'direct') {
    const cur = getJS(doc, path);
    const keep = cur && typeof cur === 'object' ? Object.fromEntries(Object.entries(cur).filter(([k]) => ['id', 'padding'].includes(k))) : {};
    replaceNode(doc, path, keep);
  } else {
    for (const k of CONTENT_KEYS) setProp(doc, path, k, null);
  }
  return fileChanged(file);
}

// 兄弟と入れ替える（direct のときだけ）
function moveNode(box, delta) {
  const { file, path } = box.src;
  const doc = S.files[file].doc;
  const seq = doc.getIn(path.slice(0, -1), true);
  const k = path[path.length - 1];
  if (!YAML.isSeq(seq) || typeof k !== 'number') return;
  const j = k + delta;
  if (j < 0 || j >= seq.items.length) return;
  [seq.items[k], seq.items[j]] = [seq.items[j], seq.items[k]];
  const parent = findBox(box.parent);
  const sib = curSlide().boxes.find((b) => b.parent === box.parent && b.src?.path.at(-1) === j);
  S.selected = sib ? sib.path : parent?.path || null;
  return fileChanged(file);
}

// split の子の数・grid のセルの数を変える
function setChildCount(box, listKey, n) {
  const { file, path, mode } = box.src;
  const doc = S.files[file].doc;
  const node = box.node;
  n = Math.max(1, Math.min(60, n | 0));
  const seq = doc.getIn([...path, listKey], true);
  if (mode === 'direct' && YAML.isSeq(seq)) {
    while (seq.items.length < n) seq.add(doc.createNode({}));
    while (seq.items.length > n) seq.items.pop();
  } else {
    const arr = plain(node[listKey] || []).slice(0, n);
    while (arr.length < n) arr.push({});
    setProp(doc, path, listKey, arr);
  }
  if (listKey === 'children' && Array.isArray(node.ratio)) {
    const ratio = node.ratio.slice(0, n);
    while (ratio.length < n) ratio.push(1);
    setProp(doc, path, 'ratio', ratio);
  }
  return fileChanged(file);
}

// ---------------------------------------------------------------- 設定欄

function renderProps() {
  const el = $('#node');
  el.innerHTML = '';
  if (!S.layout) return;
  const box = S.selected ? findBox(S.selected) : null;
  el.append(h('h2', null, '選んだ枠'));
  if (!box) {
    el.append(h('p', 'note', 'プレビューで枠をクリックすると、その枠の設定がここに出ます。'));
    return;
  }
  const node = box.node || {};
  const inh = box.inherited || {};
  const u = S.layout.unit;
  el.append(h('p', 'meta', `${KIND_LABEL[box.kind]}　${fmt(box.rect.w)} × ${fmt(box.rect.h)} ${u}`));
  if (!box.src) return;
  const mode = box.src.mode;
  if (mode === 'shared') el.append(h('p', 'warn', 'この枠は、複数のスライドで共通に使う枠の組み合わせ（テンプレート）の一部です。変えると、このテンプレートを使う全てのスライドが変わります。'));

  // 種類
  const fsKind = h('div');
  fsKind.append(
    row('種類', selectEl(KIND_OPTS, box.kind, (v) => changeKind(box, v))),
  );
  el.append(fsKind);

  // 中身
  if (box.kind === 'split') el.append(splitFields(box, node));
  if (box.kind === 'grid') el.append(gridFields(box, node));
  if (box.kind === 'image') el.append(imageFields(box, node, inh));
  if (box.kind === 'text') el.append(textFields(box, node, inh));

  // 引き継ぐ設定（分割・格子のとき）
  if (box.kind === 'split' || box.kind === 'grid') el.append(inheritFields(box, node, inh));

  // 枠
  const fsBox = fieldset('余白・背景・枠線');
  fsBox.append(
    row('内側の余白', textEl(fmtList(node.padding), '0', (v) => edit(box, 'padding', parseList(v))), '1 つ書くと四方共通。2 つは上下, 左右、4 つは上, 右, 下, 左'),
    row('背景色', colorEl(node.background, (v) => edit(box, 'background', v || undefined))),
    row('枠線', h2(
      numEl(node.border?.width ?? '', 'pt', (v) => edit(box, 'border', v === '' ? undefined : { width: +v, color: node.border?.color || '#333333' })),
      'pt',
      colorEl(node.border?.color, (v) => node.border && edit(box, ['border', 'color'], v || '#333333')),
    )),
  );
  el.append(fsBox);

  // 操作
  const acts = h('div', 'actions');
  if (box.parent) acts.append(btn('外側の枠を選ぶ', () => select(box.parent)));
  acts.append(btn('この枠を拡大', () => focusOn(box.path)));
  if (mode === 'direct') {
    acts.append(btn('上下に分ける', () => wrapSplit(box, 'rows')), btn('左右に分ける', () => wrapSplit(box, 'cols')));
    if (typeof box.src.path.at(-1) === 'number') acts.append(btn('前の枠と入れ替える', () => moveNode(box, -1)), btn('次の枠と入れ替える', () => moveNode(box, 1)));
  }
  if (box.kind !== 'empty' && box.depth > 0) {
    const b = btn('中身を消す', () => confirm('この枠の中身を消しますか？') && clearNode(box));
    b.className = 'danger';
    acts.append(b);
  }
  el.append(acts);
}

// スライド全体の設定（全スライド共通）。用紙の寸法は cm で持ち、書き込むときに YAML の単位に直す
const PAPERS = [
  ['A0 縦', 84.1, 118.9],
  ['A0 横', 118.9, 84.1],
  ['A1 縦', 59.4, 84.1],
  ['A1 横', 84.1, 59.4],
  ['B0 縦', 103, 145.6],
  ['B0 横', 145.6, 103],
  ['B1 縦', 72.8, 103],
  ['B1 横', 103, 72.8],
  ['PowerPoint の標準 16:9', 33.867, 19.05],
  ['PowerPoint の標準 4:3', 25.4, 19.05],
];
const RATIOS = [
  ['16:9', 16 / 9],
  ['4:3', 4 / 3],
  ['1:1', 1],
  ['3:4', 3 / 4],
  ['9:16', 9 / 16],
  ['A判・B判の縦（1:√2）', 1 / Math.SQRT2],
  ['A判・B判の横（√2:1）', Math.SQRT2],
];
let lockRatio = true;
try {
  lockRatio = localStorage.getItem('slide-creator.lockRatio') !== '0';
} catch {
  /* 使えなければ固定のまま */
}

function renderGlobal() {
  const el = $('#global');
  el.innerHTML = '';
  if (!S.layout) return;
  const doc = S.files['layout.yaml'].doc;
  const c = getJS(doc, ['canvas']) || {};
  const unit = c.unit || 'cm';
  const W = +c.width;
  const H = +c.height;
  const r3 = (v) => Math.round(v * 1000) / 1000;
  const fromCm = (v) => r3((v / 2.54) * PER_INCH[unit]);
  const toCm = (v) => (v / PER_INCH[unit]) * 2.54;
  const set = (pairs) => {
    for (const [k, v] of pairs) setProp(doc, ['canvas'], k, v);
    return fileChanged('layout.yaml');
  };

  el.append(h('h2', null, 'スライド全体'));

  const wIn = numEl(c.width ?? '', '', (v) => {
    if (!(+v > 0)) return;
    set(lockRatio ? [['width', +v], ['height', r3((+v * H) / W)]] : [['width', +v]]);
  });
  const hIn = numEl(c.height ?? '', '', (v) => {
    if (!(+v > 0)) return;
    set(lockRatio ? [['height', +v], ['width', r3((+v * W) / H)]] : [['height', +v]]);
  });
  const unitIn = selectEl([['cm', 'cm'], ['mm', 'mm'], ['in', 'in'], ['pt', 'pt'], ['px', 'px']], unit, (nu) => {
    const k = PER_INCH[nu] / PER_INCH[unit];
    set([['width', r3(W * k)], ['height', r3(H * k)], ['unit', nu]]);
  });

  const ratioHit = RATIOS.findIndex(([, v]) => Math.abs(v / (W / H) - 1) < 0.003);
  const ratioOpts = [...(ratioHit < 0 ? [['', `${(W / H).toFixed(3)} : 1（一覧に無い比）`]] : []), ...RATIOS.map(([l], i) => [String(i), l])];
  const ratioIn = selectEl(ratioOpts, ratioHit >= 0 ? String(ratioHit) : '', (v) => {
    if (v === '') return;
    // 幅を保ち、高さを比に合わせる
    set([['height', r3(W / RATIOS[+v][1])]]);
  });

  const lockIn = checkEl('縦横比を固定する', lockRatio, (v) => {
    lockRatio = v;
    try {
      localStorage.setItem('slide-creator.lockRatio', v ? '1' : '0');
    } catch {
      /* 保存できなくても動く */
    }
  });

  const paperHit = PAPERS.findIndex(([, w, hh]) => Math.abs(w - toCm(W)) < 0.05 && Math.abs(hh - toCm(H)) < 0.05);
  const paperOpts = [...(paperHit < 0 ? [['', '一覧に無い大きさ']] : []), ...PAPERS.map(([l, w, hh], i) => [String(i), `${l}（${w} × ${hh} cm）`])];
  const paperIn = selectEl(paperOpts, paperHit >= 0 ? String(paperHit) : '', (v) => {
    if (v === '') return;
    const [, w, hh] = PAPERS[+v];
    set([['width', fromCm(w)], ['height', fromCm(hh)]]);
  });

  el.append(
    row('幅', h2(wIn, unit)),
    row('高さ', h2(hIn, unit)),
    row('単位', unitIn),
    row('縦横比', ratioIn),
    row('', lockIn),
    row('用紙', paperIn),
  );
  const acts = h('div', 'actions');
  acts.append(btn('縦と横を入れ替える', () => set([['width', H], ['height', W]])));
  el.append(acts);
  for (const i of S.layout.issues || []) el.append(h('p', 'warn', i.msg));
}

function findTemplateName(box) {
  for (let b = box; b; b = b.parent ? findBox(b.parent) : null) if (b.template) return b.template;
  return null;
}

function splitFields(box, node) {
  const fs = fieldset('分割');
  const n = (node.children || []).length;
  fs.append(
    row('向き', selectEl([['rows', '上下に並べる'], ['cols', '左右に並べる']], node.split, (v) => edit(box, 'split', v))),
    row('分ける数', numEl(n, '', (v) => setChildCount(box, 'children', +v))),
    row('長さの比', textEl(fmtList(node.ratio), Array(n).fill(1).join(', '), (v) => edit(box, 'ratio', parseList(v))), '部分ごとの長さの比を「,」で区切って書く（例：1, 2）。「3cm」のように単位を付けた部分は、その長さに固定される'),
    row('部分の間隔', textEl(node.gap ?? '', '0', (v) => edit(box, 'gap', parseLen(v)))),
  );
  return fs;
}

function gridFields(box, node) {
  const fs = fieldset('格子');
  const g = node.grid || {};
  const cells = (g.rows || 1) * (g.cols || 1);
  const items = (node.items || []).length;
  fs.append(
    row('行 × 列', h2(numEl(g.rows || 1, '', (v) => edit(box, ['grid', 'rows'], Math.max(1, +v))), '×', numEl(g.cols || 1, '', (v) => edit(box, ['grid', 'cols'], Math.max(1, +v))))),
    row('セルの間隔', textEl(fmtList(g.gap ?? node.gap), '0', (v) => edit(box, ['grid', 'gap'], parseList(v))), '1 つ書くと縦横共通。2 つ書くと、行の間, 列の間'),
    row('並べる順', selectEl([['', '横に（行ごと）'], ['col', '縦に（列ごと）']], g.order || '', (v) => edit(box, ['grid', 'order'], v || undefined))),
  );
  const info = h('div', 'note', `置いたもの ${items} 個／セル ${cells} 個`);
  fs.append(info);
  if (items !== cells) fs.append(h('div', 'actions', btn(`中身の数をセルの数（${cells}）に合わせる`, () => setChildCount(box, 'items', cells))));
  return fs;
}

function imageFields(box, node, inh) {
  const fs = fieldset('画像');
  const opts = [...new Set([...S.assets, String(node.image)])].map((a) => [a, a]);
  const srcIn = S.assets.length
    ? selectEl(opts, String(node.image), (v) => edit(box, 'image', v))
    : textEl(String(node.image), 'assets/xxx.png', (v) => edit(box, 'image', v));
  const el = S.layout && curSlide().elements.find((e) => e.path === box.path && e.type === 'image');
  const crop = node.crop || [0, 0, 1, 1];
  const cropIns = crop.map((c, i) =>
    numEl(c, '', (v) => {
      const next = [...crop];
      next[i] = +v;
      edit(box, 'crop', next);
    }, 0.01),
  );
  const cropRow = h2(...cropIns);
  cropRow.classList.add('crop4');
  fs.append(
    row('ファイル', srcIn),
    row('収め方', selectEl(withInherit(FIT_OPTS, inh.fit, 'contain'), node.fit ?? '', (v) => edit(box, 'fit', v || undefined))),
    row('横の位置', selectEl(withInherit(H_OPTS, inh.align, 'center'), node.align ?? '', (v) => edit(box, 'align', v || undefined))),
    row('縦の位置', selectEl(withInherit(V_OPTS, inh.valign, 'middle'), node.valign ?? '', (v) => edit(box, 'valign', v || undefined))),
    row('トリミング', cropRow, '残す範囲の左端・上端・右端・下端を、元画像の幅・高さに対する割合（0〜1）で書く'),
  );
  const acts = h('div', 'actions');
  acts.append(btn('画像を見ながらトリミング…', () => cropDialog(box, node, el)));
  if (node.crop) acts.append(btn('トリミングを外す', () => edit(box, 'crop', undefined)));
  fs.append(acts);
  if (el) fs.append(h('div', 'note', `元画像 ${el.natural.w} × ${el.natural.h} px、表示する大きさ ${fmt(el.rect.w)} × ${fmt(el.rect.h)} ${S.layout.unit}、解像度 ${el.dpi} dpi`));
  return fs;
}

async function cropDialog(box, node, el) {
  const nat = S.imageSizes.get(String(node.image));
  if (!nat) return status('画像が読めません', true);
  const cell = el?.cell || box.rect;
  const result = await openCrop({
    src: String(node.image),
    url: S.imageURLs.get(String(node.image)),
    natural: nat,
    crop: node.crop || [0, 0, 1, 1],
    cellAspect: cell.w / cell.h,
    lockDefault: (node.fit ?? box.inherited?.fit) === 'cover',
  });
  if (result === 'clear') return edit(box, 'crop', undefined);
  if (Array.isArray(result)) return edit(box, 'crop', result);
}

function textFields(box, node, inh) {
  const fs = fieldset('文字');
  const f = node.font || {};
  const fi = inh.font || {};
  const ta = document.createElement('textarea');
  ta.value = node.text ?? '';
  ta.onchange = () => edit(box, 'text', ta.value);
  fs.append(
    row('文字', ta),
    row('大きさ', h2(numEl(f.size ?? '', String(fi.size ?? 18), (v) => edit(box, ['font', 'size'], v === '' ? undefined : +v)), 'pt')),
    row('フォント', textEl(f.family ?? '', fi.family || 'Arial', (v) => edit(box, ['font', 'family'], v || undefined))),
    row('太字・斜体・色', h2(
      checkEl('太字', f.bold ?? false, (v) => edit(box, ['font', 'bold'], v || undefined)),
      checkEl('斜体', f.italic ?? false, (v) => edit(box, ['font', 'italic'], v || undefined)),
      colorEl(f.color ?? fi.color, (v) => edit(box, ['font', 'color'], v || undefined)),
    )),
    row('横の揃え方', selectEl(withInherit(H_OPTS, inh.align, 'left'), node.align ?? '', (v) => edit(box, 'align', v || undefined))),
    row('縦の揃え方', selectEl(withInherit(V_OPTS, inh.valign, 'top'), node.valign ?? '', (v) => edit(box, 'valign', v || undefined))),
    row('行の高さ', numEl(node.line_height ?? '', String(inh.line_height ?? 1.2), (v) => edit(box, 'line_height', v === '' ? undefined : +v), 0.05)),
  );
  return fs;
}

function inheritFields(box, node, inh) {
  const fs = fieldset('中の枠すべてに効く設定');
  const f = node.font || {};
  fs.append(
    row('横の揃え方', selectEl(withInherit(H_OPTS, inh.align, '画像は中央、文字は左'), node.align ?? '', (v) => edit(box, 'align', v || undefined))),
    row('縦の揃え方', selectEl(withInherit(V_OPTS, inh.valign, '画像は中央、文字は上'), node.valign ?? '', (v) => edit(box, 'valign', v || undefined))),
    row('画像の収め方', selectEl(withInherit(FIT_OPTS, inh.fit, 'contain'), node.fit ?? '', (v) => edit(box, 'fit', v || undefined))),
    row('文字の大きさ', h2(numEl(f.size ?? '', String(inh.font?.size ?? 18), (v) => edit(box, ['font', 'size'], v === '' ? undefined : +v)), 'pt')),
  );
  return fs;
}

// ---------------------------------------------------------------- 部品

const KIND_LABEL = { empty: '空き', split: '分割', grid: '格子', image: '画像', text: '文字' };
const KIND_OPTS = [['empty', '空き（何も置かない）'], ['split', '分割（中を上下か左右に分ける）'], ['grid', '格子（同じ大きさで縦横に並べる）'], ['image', '画像'], ['text', '文字']];
const FIT_OPTS = [['contain', '切らずに全体を収める'], ['cover', '枠いっぱいに広げ、はみ出た部分を切る'], ['stretch', '縦横比を変えて枠に合わせる']];
const H_OPTS = [['left', '左'], ['center', '中央'], ['right', '右']];
const V_OPTS = [['top', '上'], ['middle', '中央'], ['bottom', '下']];

function withInherit(opts, inherited, dflt) {
  const label = inherited != null ? `（外側の枠の設定：${(opts.find((o) => o[0] === inherited) || [, inherited])[1]}）` : `（指定なし：${(opts.find((o) => o[0] === dflt) || [, dflt])[1]}）`;
  return [['', label], ...opts];
}

function h(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}
function h2(...parts) {
  const d = h('span', 'in');
  d.append(...parts);
  return d;
}
function fieldset(title) {
  const f = h('fieldset');
  f.append(h('legend', null, title));
  return f;
}
function row(label, input, hint) {
  const r = h('div', 'row');
  const wrap = input.classList?.contains('in') ? input : h2(input);
  r.append(h('label', null, label), wrap);
  if (hint) r.append(h('div', 'hint', hint));
  return r;
}
function textEl(value, placeholder, onCommit) {
  const i = document.createElement('input');
  i.type = 'text';
  i.value = value;
  i.placeholder = placeholder;
  i.onchange = () => onCommit(i.value.trim());
  return i;
}
function numEl(value, placeholder, onCommit, step = 'any') {
  const i = document.createElement('input');
  i.type = 'number';
  i.step = step;
  i.value = value;
  i.placeholder = placeholder;
  i.onchange = () => onCommit(i.value);
  return i;
}
function selectEl(opts, value, onChange) {
  const s = document.createElement('select');
  for (const [v, l] of opts) {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = l;
    s.append(o);
  }
  s.value = value ?? '';
  s.onchange = () => onChange(s.value);
  return s;
}
function checkEl(label, value, onChange) {
  const l = h('label', 'chk');
  const i = document.createElement('input');
  i.type = 'checkbox';
  i.checked = !!value;
  i.onchange = () => onChange(i.checked);
  l.append(i, label);
  return l;
}
function colorEl(value, onCommit) {
  const wrap = h('span', 'in');
  const c = document.createElement('input');
  c.type = 'color';
  const norm = normColor(value);
  c.value = norm || '#000000';
  const t = textEl(value ?? '', '（なし）', (v) => onCommit(v));
  c.oninput = () => (t.value = c.value);
  c.onchange = () => onCommit(c.value);
  wrap.append(c, t);
  return wrap;
}
function btn(label, fn) {
  const b = h('button', null, label);
  b.type = 'button';
  b.onclick = fn;
  return b;
}
function normColor(v) {
  if (!v) return null;
  const s = String(v).replace(/^#/, '');
  if (/^[0-9a-f]{6}$/i.test(s)) return `#${s}`;
  if (/^[0-9a-f]{3}$/i.test(s)) return `#${s.split('').map((c) => c + c).join('')}`;
  return null;
}
const fmt = (v) => (+v).toFixed(2);
function fmtList(v) {
  if (v == null) return '';
  return Array.isArray(v) ? v.join(', ') : String(v);
}
function parseLen(s) {
  if (s === '') return undefined;
  return /^-?[\d.]+$/.test(s) ? +s : s;
}
function parseList(s) {
  if (s === '') return undefined;
  const parts = s.split(/[,\s]+/).filter(Boolean).map(parseLen);
  return parts.length === 1 ? parts[0] : parts;
}

// ---------------------------------------------------------------- 保存・書き出し

async function save() {
  if (!S.source) return;
  const dirty = Object.entries(S.files).filter(([, f]) => f.dirty);
  if (!dirty.length) return;
  try {
    for (const [name, f] of dirty) {
      await S.source.writeFile(name, f.text);
      f.dirty = false;
    }
    $('#btn-save').disabled = true;
    status(S.source.writable ? '保存しました' : '見本はフォルダに保存できないため、ファイルをダウンロードしました');
  } catch (e) {
    status(`保存できません: ${e.message}`, true);
  }
}

async function reload() {
  if (anyDirty() && !confirm('保存していない変更を捨てて、ファイルを読み直しますか？')) return;
  const keep = { slide: S.slideId, sel: S.selected, focus: S.focus };
  S.source.forgetImages();
  for (const f of Object.values(S.files)) f.dirty = false;
  await openSource(S.source);
  Object.assign(S, { slideId: keep.slide, selected: keep.sel, focus: keep.focus });
  await rebuild();
}

const blobToDataURL = (blob) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsDataURL(blob);
  });

async function exportPPTX() {
  if (!S.layout) return;
  try {
    status('PPTX を作成中…');
    const pres = await buildPPTX(window.PptxGenJS, S.layout, async (src) => ({ data: await blobToDataURL(await S.source.readBlob(src)) }));
    const blob = await pres.write({ outputType: 'blob' });
    const name = `${S.source.name}.pptx`;
    if (S.source.writable) {
      await S.source.writeFile(`out/${name}`, blob);
      status(`out/${name} に書き出しました`);
    } else {
      download(blob, name);
      status('ダウンロードしました');
    }
  } catch (e) {
    status(`PPTX を作れません: ${e.message}`, true);
  }
}

function exportPDF() {
  if (!S.layout) return;
  const html = renderHTML(S.layout, { srcURL: (src) => S.imageURLs.get(src) || '' });
  const f = document.createElement('iframe');
  f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  document.body.append(f);
  f.onload = async () => {
    await f.contentDocument.fonts.ready;
    await Promise.all([...f.contentDocument.images].map((i) => i.decode().catch(() => {})));
    f.contentWindow.print();
    setTimeout(() => f.remove(), 1000);
  };
  f.srcdoc = html;
  status('印刷画面で「PDFに保存」を選んでください');
}

// ---------------------------------------------------------------- 前回のフォルダ

const IDB = {
  open: () =>
    new Promise((res, rej) => {
      const r = indexedDB.open('slide-creator', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }),
  async get(k) {
    const db = await this.open();
    return new Promise((res) => {
      const q = db.transaction('kv').objectStore('kv').get(k);
      q.onsuccess = () => res(q.result);
      q.onerror = () => res(null);
    });
  },
  async set(k, v) {
    const db = await this.open();
    db.transaction('kv', 'readwrite').objectStore('kv').put(v, k);
  },
};

async function initReopen() {
  try {
    const handle = await IDB.get('lastFolder');
    if (!handle) return;
    const b = $('#btn-reopen');
    b.textContent = `前回のフォルダを開く（${handle.name}）`;
    b.hidden = false;
    b.onclick = async () => {
      if ((await handle.requestPermission({ mode: 'readwrite' })) === 'granted') openSource(new FolderSource(handle));
    };
  } catch {
    /* 使えない環境では出さない */
  }
}

// ---------------------------------------------------------------- 全体

function status(msg, isErr = false) {
  const s = $('#status');
  s.textContent = msg;
  s.className = isErr ? 'err' : '';
}


function onKey(e) {
  if ((e.metaKey || e.ctrlKey) && e.key === 's') {
    e.preventDefault();
    save();
    return;
  }
  const tag = (e.target.tagName || '').toLowerCase();
  if (['input', 'textarea', 'select'].includes(tag)) return;
  const box = S.selected ? findBox(S.selected) : null;
  if (e.key === 'Escape') {
    if (box?.parent) select(box.parent);
    else if (S.focus) focusOn(null);
  } else if (e.key === 'Enter' && box) {
    focusOn(box.path);
  }
}

$('#btn-open').onclick = async () => {
  try {
    const src = await FolderSource.pick();
    if (!(await src.exists('layout.yaml'))) {
      if (!confirm(`${src.name} に layout.yaml がありません。新しく作りますか？`)) return;
      const tpl = await (await fetch('../templates/init.yaml')).text();
      await src.writeFile('layout.yaml', tpl);
    }
    IDB.set('lastFolder', src.handle).catch(() => {});
    openSource(src);
  } catch (e) {
    if (e.name !== 'AbortError') status(e.message, true);
  }
};
$('#btn-demo').onclick = () => openSource(new UrlSource('../examples/demo/', 'demo'));
$('#btn-save').onclick = save;
$('#btn-reload').onclick = reload;
$('#btn-pptx').onclick = exportPPTX;
$('#btn-pdf').onclick = exportPDF;
$('#btn-focus').onclick = () => S.selected && focusOn(S.selected);
$('#btn-whole').onclick = () => focusOn(null);
$('#chk-outline').onchange = (e) => {
  S.outline = e.target.checked;
  renderPreview();
};
document.addEventListener('keydown', onKey);
window.addEventListener('resize', () => applyView());
window.addEventListener('beforeunload', (e) => {
  if (anyDirty()) e.preventDefault();
});
initReopen();
window.slideCreator = { state: S, buildPPTX }; // 開発者ツールからの確認用
if (new URLSearchParams(location.search).has('demo')) $('#btn-demo').click();
