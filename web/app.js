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
  if (anyDirty() && !confirm('保存していない変更があります。破棄して開きますか？')) return;
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
    $('#project-name').textContent = source.name + (source.writable ? '' : '（読むだけ）');
    document.title = `${source.name} — slide-creator`;
    $('#empty').hidden = true;
    for (const id of ['#btn-reload', '#btn-pptx', '#btn-pdf']) $(id).disabled = false;
    fillYamlSelect();
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
  renderProps();
  renderIssues();
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
  renderIssues();
}

// ---------------------------------------------------------------- 問題

function renderIssues() {
  const el = $('#tab-issues');
  el.innerHTML = '';
  const list = [];
  if (S.error) list.push({ level: 'error', path: '', msg: S.error });
  list.push(...(S.layout?.issues || []));
  for (const s of S.layout?.slides || []) list.push(...s.issues.map((i) => ({ ...i, slide: s.id })));
  list.push(...S.overflow.map((i) => ({ ...i, slide: S.slideId })));
  $('#issue-count').textContent = list.length ? `(${list.length})` : '';
  if (!list.length) {
    el.innerHTML = '<p class="note">問題はありません。文字のはみ出しは、表示中のスライドだけを測っています。</p>';
    return;
  }
  for (const i of list) {
    const d = document.createElement('div');
    d.className = `issue ${i.level}`;
    d.innerHTML = `<b>${i.level === 'error' ? '誤り' : '注意'}</b> <code></code><div></div>`;
    d.querySelector('code').textContent = i.path || '—';
    d.querySelector('div').textContent = i.msg;
    d.onclick = () => {
      if (!i.slide) return showTab('yaml');
      if (i.slide !== S.slideId) {
        S.slideId = i.slide;
        renderSlideTabs();
        renderPreview();
      }
      S.selected = i.path;
      showTab('props');
      setTimeout(applyView, 50);
      renderProps();
    };
    el.append(d);
  }
}

// ---------------------------------------------------------------- 編集

function fileChanged(file) {
  const f = S.files[file];
  f.text = stringify(f.doc);
  f.dirty = true;
  if ($('#yaml-file').value === file) $('#yaml-text').value = f.text;
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
    text: { text: 'テキスト' },
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

const MODE_NOTE = {
  direct: '',
  fill: '型「{t}」の中のノード。このスライドの fill に書き込みます。',
  override: '型「{t}」の根。このスライドの use の行に上書きします。',
  shared: '型の定義を書き換えます。この型を使う全ての箇所に効きます。',
};

function renderProps() {
  const el = $('#tab-props');
  el.innerHTML = '';
  if (!S.layout) return;
  const box = S.selected ? findBox(S.selected) : null;
  if (!box) {
    el.append(canvasFields());
    el.append(h('p', 'note', `スライド ${S.layout.slides.length} 枚。プレビューで枠をクリックすると、その枠の設定が出ます。`));
    return;
  }
  if (box.depth === 0) el.append(canvasFields());
  const node = box.node || {};
  const inh = box.inherited || {};
  const u = S.layout.unit;

  const head = h('div', 'p-head');
  head.append(h('div', 'path', box.path), h('div', 'meta', `${KIND_LABEL[box.kind]}　位置 (${fmt(box.rect.x)}, ${fmt(box.rect.y)})　大きさ ${fmt(box.rect.w)} × ${fmt(box.rect.h)} ${u}`));
  el.append(head);

  if (!box.src) {
    el.append(h('p', 'note', 'この枠は編集先が分かりません。'));
    return;
  }
  const mode = box.src.mode;
  const tgt = h('div', `p-target${mode === 'shared' ? ' shared' : ''}`);
  const note = (MODE_NOTE[mode] || '').replace('{t}', findTemplateName(box) || '');
  tgt.append(note ? `${note} ` : '', '書き込み先：');
  const a = h('a', null, `${box.src.file} › ${box.src.path.join('.')}`);
  a.onclick = () => revealInYaml(box.src);
  tgt.append(a);
  el.append(tgt);

  // 種類
  const fsKind = fieldset('種類');
  fsKind.append(
    row('種類', selectEl(Object.entries(KIND_LABEL), box.kind, (v) => changeKind(box, v))),
  );
  if (mode === 'direct') {
    fsKind.append(row('id', textEl(node.id ?? '', '（なし）', (v) => edit(box, 'id', v || undefined)), 'CLI やテンプレートの fill で使う名前'));
  }
  el.append(fsKind);

  // 中身
  if (box.kind === 'split') el.append(splitFields(box, node));
  if (box.kind === 'grid') el.append(gridFields(box, node));
  if (box.kind === 'image') el.append(imageFields(box, node, inh));
  if (box.kind === 'text') el.append(textFields(box, node, inh));

  // 引き継ぐ設定（分割・格子のとき）
  if (box.kind === 'split' || box.kind === 'grid') el.append(inheritFields(box, node, inh));

  // 枠
  const fsBox = fieldset('枠');
  fsBox.append(
    row('内側の余白', textEl(fmtList(node.padding), '0', (v) => edit(box, 'padding', parseList(v))), '1 つ＝四方、2 つ＝上下, 左右、4 つ＝上, 右, 下, 左'),
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
  if (box.parent) acts.append(btn('親を選ぶ', () => select(box.parent)));
  acts.append(btn('拡大', () => focusOn(box.path)));
  if (mode === 'direct') {
    acts.append(btn('上下に分ける', () => wrapSplit(box, 'rows')), btn('左右に分ける', () => wrapSplit(box, 'cols')));
    if (typeof box.src.path.at(-1) === 'number') acts.append(btn('◀ 前と入替', () => moveNode(box, -1)), btn('後と入替 ▶', () => moveNode(box, 1)));
  }
  if (box.kind !== 'empty' && box.depth > 0) {
    const b = btn('空にする', () => confirm('この枠の中身を消しますか？') && clearNode(box));
    b.className = 'danger';
    acts.append(b);
  }
  el.append(acts);
}

// 1 枚の大きさ（全スライド共通）。値は cm で持ち、書き込むときに YAML の単位に直す
const CANVAS_PRESETS = [
  ['16:9（PowerPoint 標準）', 33.867, 19.05],
  ['4:3', 25.4, 19.05],
  ['A0 縦', 84.1, 118.9],
  ['A0 横', 118.9, 84.1],
  ['A1 縦', 59.4, 84.1],
  ['A1 横', 84.1, 59.4],
  ['B0 縦（PowerPoint の上限超え）', 103, 145.6],
  ['B0 横（PowerPoint の上限超え）', 145.6, 103],
  ['B1 縦', 72.8, 103],
  ['B1 横', 103, 72.8],
];

function canvasFields() {
  const doc = S.files['layout.yaml'].doc;
  const c = getJS(doc, ['canvas']) || {};
  const unit = c.unit || 'cm';
  const toCm = (v) => (+v / PER_INCH[unit]) * 2.54;
  const fromCm = (v) => Math.round(((v / 2.54) * PER_INCH[unit]) * 1000) / 1000;
  const wCm = toCm(c.width);
  const hCm = toCm(c.height);
  const hit = CANVAS_PRESETS.findIndex(([, w, hh]) => Math.abs(w - wCm) < 0.05 && Math.abs(hh - hCm) < 0.05);

  const editCanvas = (pairs) => {
    for (const [k, v] of pairs) setProp(doc, ['canvas'], k, v);
    return fileChanged('layout.yaml');
  };
  const fs = fieldset('キャンバス（1 枚の大きさ・全スライド共通）');
  const presetOpts = [['', '自由に指定'], ...CANVAS_PRESETS.map(([l, w, hh], i) => [String(i), `${l}　${w} × ${hh} cm`])];
  fs.append(
    row('よく使う大きさ', selectEl(presetOpts, hit >= 0 ? String(hit) : '', (v) => {
      if (v === '') return;
      const [, w, hh] = CANVAS_PRESETS[+v];
      editCanvas([['width', fromCm(w)], ['height', fromCm(hh)]]);
    })),
    row('幅 × 高さ', h2(
      numEl(c.width ?? '', '', (v) => v !== '' && +v > 0 && editCanvas([['width', +v]])),
      '×',
      numEl(c.height ?? '', '', (v) => v !== '' && +v > 0 && editCanvas([['height', +v]])),
      selectEl([['cm', 'cm'], ['mm', 'mm'], ['in', 'in'], ['pt', 'pt'], ['px', 'px']], unit, (nu) => {
        const k = PER_INCH[nu] / PER_INCH[unit];
        const r = (v) => Math.round(+v * k * 1000) / 1000;
        editCanvas([['width', r(c.width)], ['height', r(c.height)], ['unit', nu]]);
      }),
    )),
  );
  const acts = h('div', 'actions');
  acts.append(btn('縦と横を入れ替える', () => editCanvas([['width', c.height], ['height', c.width]])));
  fs.append(acts);
  // 20 以下の整数比で 0.3% 以内に収まれば、それで表す（33.867 × 19.05 → 16:9）
  let ratio = `${(wCm / hCm).toFixed(3)} : 1`;
  outer: for (let q = 1; q <= 20; q++) {
    for (let p = 1; p <= 20; p++) {
      if (Math.abs(p / q / (wCm / hCm) - 1) < 0.003) {
        ratio = `${p}:${q}`;
        break outer;
      }
    }
  }
  fs.append(h('div', 'note', `縦横比 ${ratio}。中の枠は比率で決まるので、大きさを変えても配置の関係は保たれます（文字の pt は変わらないので、大きくしたら文字も大きくしてください）。`));
  return fs;
}

function findTemplateName(box) {
  for (let b = box; b; b = b.parent ? findBox(b.parent) : null) if (b.template) return b.template;
  return null;
}

function splitFields(box, node) {
  const fs = fieldset('分割');
  const n = (node.children || []).length;
  fs.append(
    row('向き', selectEl([['rows', '上下に並べる（rows）'], ['cols', '左右に並べる（cols）']], node.split, (v) => edit(box, 'split', v))),
    row('数', numEl(n, '', (v) => setChildCount(box, 'children', +v))),
    row('比率', textEl(fmtList(node.ratio), Array(n).fill(1).join(', '), (v) => edit(box, 'ratio', parseList(v))), '数値は比率、"3cm" のような長さは固定長'),
    row('間隔', textEl(node.gap ?? '', '0', (v) => edit(box, 'gap', parseLen(v)))),
  );
  return fs;
}

function gridFields(box, node) {
  const fs = fieldset('格子（同じ大きさのセル）');
  const g = node.grid || {};
  const cells = (g.rows || 1) * (g.cols || 1);
  const items = (node.items || []).length;
  fs.append(
    row('行 × 列', h2(numEl(g.rows || 1, '', (v) => edit(box, ['grid', 'rows'], Math.max(1, +v))), '×', numEl(g.cols || 1, '', (v) => edit(box, ['grid', 'cols'], Math.max(1, +v))))),
    row('間隔', textEl(fmtList(g.gap ?? node.gap), '0', (v) => edit(box, ['grid', 'gap'], parseList(v))), '1 つ＝共通、2 つ＝行の間, 列の間'),
    row('詰める順', selectEl([['', '行ごと'], ['col', '列ごと']], g.order || '', (v) => edit(box, ['grid', 'order'], v || undefined))),
  );
  const info = h('div', 'note', `中身 ${items} 個 ／ セル ${cells} 個`);
  fs.append(info);
  if (items !== cells) fs.append(h('div', 'actions', btn(`中身の数をセルに合わせる（${cells} 個）`, () => setChildCount(box, 'items', cells))));
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
    row('収め方', selectEl(withInherit([['contain', 'contain：切らずに収める'], ['cover', 'cover：枠いっぱい、はみ出しを切る'], ['stretch', 'stretch：伸ばす']], inh.fit, 'contain'), node.fit ?? '', (v) => edit(box, 'fit', v || undefined))),
    row('横の位置', selectEl(withInherit(H_OPTS, inh.align, 'center'), node.align ?? '', (v) => edit(box, 'align', v || undefined))),
    row('縦の位置', selectEl(withInherit(V_OPTS, inh.valign, 'middle'), node.valign ?? '', (v) => edit(box, 'valign', v || undefined))),
    row('トリミング', cropRow, '元画像に対する 左, 上, 右, 下 の位置（0〜1）'),
  );
  const acts = h('div', 'actions');
  acts.append(btn('トリミング画面…', () => cropDialog(box, node, el)));
  if (node.crop) acts.append(btn('トリミングを外す', () => edit(box, 'crop', undefined)));
  fs.append(acts);
  if (el) fs.append(h('div', 'note', `元画像 ${el.natural.w} × ${el.natural.h} px ／ 表示 ${fmt(el.rect.w)} × ${fmt(el.rect.h)} ${S.layout.unit} ／ ${el.dpi} dpi`));
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
    row('書体', textEl(f.family ?? '', fi.family || 'Arial', (v) => edit(box, ['font', 'family'], v || undefined))),
    row('飾り', h2(
      checkEl('太字', f.bold ?? false, (v) => edit(box, ['font', 'bold'], v || undefined)),
      checkEl('斜体', f.italic ?? false, (v) => edit(box, ['font', 'italic'], v || undefined)),
      colorEl(f.color ?? fi.color, (v) => edit(box, ['font', 'color'], v || undefined)),
    )),
    row('横の揃え', selectEl(withInherit(H_OPTS, inh.align, 'left'), node.align ?? '', (v) => edit(box, 'align', v || undefined))),
    row('縦の揃え', selectEl(withInherit(V_OPTS, inh.valign, 'top'), node.valign ?? '', (v) => edit(box, 'valign', v || undefined))),
    row('行の高さ', numEl(node.line_height ?? '', String(inh.line_height ?? 1.2), (v) => edit(box, 'line_height', v === '' ? undefined : +v), 0.05)),
  );
  return fs;
}

function inheritFields(box, node, inh) {
  const fs = fieldset('中に引き継ぐ設定');
  const f = node.font || {};
  fs.append(
    row('横の揃え', selectEl(withInherit(H_OPTS, inh.align, '画像は中央・文字は左'), node.align ?? '', (v) => edit(box, 'align', v || undefined))),
    row('縦の揃え', selectEl(withInherit(V_OPTS, inh.valign, '画像は中段・文字は上'), node.valign ?? '', (v) => edit(box, 'valign', v || undefined))),
    row('画像の収め方', selectEl(withInherit([['contain', 'contain'], ['cover', 'cover'], ['stretch', 'stretch']], inh.fit, 'contain'), node.fit ?? '', (v) => edit(box, 'fit', v || undefined))),
    row('文字の大きさ', h2(numEl(f.size ?? '', String(inh.font?.size ?? 18), (v) => edit(box, ['font', 'size'], v === '' ? undefined : +v)), 'pt')),
  );
  return fs;
}

// ---------------------------------------------------------------- 部品

const KIND_LABEL = { empty: '空き', split: '分割', grid: '格子', image: '画像', text: '文字' };
const H_OPTS = [['left', '左'], ['center', '中央'], ['right', '右']];
const V_OPTS = [['top', '上'], ['middle', '中段'], ['bottom', '下']];

function withInherit(opts, inherited, dflt) {
  const label = inherited != null ? `（引き継ぎ：${(opts.find((o) => o[0] === inherited) || [, inherited])[1]}）` : `（既定：${(opts.find((o) => o[0] === dflt) || [, dflt])[1]}）`;
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

// ---------------------------------------------------------------- YAML タブ

function fillYamlSelect() {
  const sel = $('#yaml-file');
  sel.innerHTML = '';
  for (const f of Object.keys(S.files)) {
    const o = document.createElement('option');
    o.value = o.textContent = f;
    sel.append(o);
  }
  sel.value = 'layout.yaml';
  $('#yaml-text').value = S.files['layout.yaml'].text;
}

let yamlTimer = null;
function onYamlInput() {
  clearTimeout(yamlTimer);
  yamlTimer = setTimeout(() => {
    const file = $('#yaml-file').value;
    const text = $('#yaml-text').value;
    try {
      S.files[file] = { text, doc: parse(text), dirty: true };
      $('#yaml-error').textContent = '';
      rebuild();
    } catch (e) {
      S.files[file].text = text;
      S.files[file].dirty = true;
      $('#yaml-error').textContent = e.message;
      $('#btn-save').disabled = false;
    }
  }, 350);
}

function revealInYaml(src) {
  showTab('yaml');
  const sel = $('#yaml-file');
  sel.value = src.file;
  const ta = $('#yaml-text');
  const f = S.files[src.file];
  ta.value = f.text;
  // 編集先そのものが無ければ（fill がまだ無いなど）、在る所まで遡る
  let node = null;
  for (let n = src.path.length; n >= 0 && !node; n--) node = src.path.length ? f.doc.getIn(src.path.slice(0, n), true) : f.doc.contents;
  const start = node?.range?.[0] ?? 0;
  ta.focus();
  ta.setSelectionRange(start, start);
  const line = f.text.slice(0, start).split('\n').length;
  ta.scrollTop = Math.max(0, (line - 4) * parseFloat(getComputedStyle(ta).lineHeight));
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
    status(S.source.writable ? '保存しました' : '読むだけのため、ダウンロードしました');
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
    b.textContent = `前回：${handle.name}`;
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

function showTab(name) {
  document.querySelectorAll('.tabs button').forEach((b) => b.classList.toggle('on', b.dataset.tab === name));
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('on', t.id === `tab-${name}`));
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
$('#yaml-text').oninput = onYamlInput;
$('#yaml-text').onkeydown = (e) => {
  if (e.key === 'Tab') {
    e.preventDefault();
    document.execCommand('insertText', false, '  ');
  }
};
$('#yaml-file').onchange = (e) => ($('#yaml-text').value = S.files[e.target.value].text);
document.querySelectorAll('.tabs button').forEach((b) => (b.onclick = () => showTab(b.dataset.tab)));
document.addEventListener('keydown', onKey);
window.addEventListener('resize', () => applyView());
window.addEventListener('beforeunload', (e) => {
  if (anyDirty()) e.preventDefault();
});
initReopen();
window.slideCreator = { state: S, buildPPTX }; // 開発者ツールからの確認用
if (new URLSearchParams(location.search).has('demo')) $('#btn-demo').click();
