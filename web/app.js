// slide-creator ブラウザ版
import { resolveDoc } from '../src/core/resolve.js';
import { layoutDoc } from '../src/core/layout.js';
import { renderHTML, FONTS } from '../src/core/html.js';
import { richToHTML } from '../src/core/richtext.js';
import { buildPPTX, finishPPTX } from '../src/core/pptx.js';
import { PER_INCH, toUnit } from '../src/core/units.js';
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
    S.assets = (await source.list('')).filter((f) => IMAGE_RE.test(f) && !f.startsWith('out/'));
    S.imageSizes.clear();
    S.imageURLs.clear();
    S.slideId = null;
    S.selected = null;
    S.focus = null;
    $('#project-name').textContent = source.name + (source.writable ? '' : '（見本）');
    document.title = `${source.name} — slide-creator`;
    $('#empty').hidden = true;
    for (const id of ['#btn-pptx', '#btn-pdf']) $(id).disabled = false;
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
  if (node.background?.image) out.add(String(node.background.image));
  (node.children || []).forEach((c) => collectImages(c, out));
  (node.items || []).forEach((c) => collectImages(c, out));
}

async function loadImage(src) {
  if (S.imageSizes.has(src)) return;
  try {
    const url = await S.source.imageURL(src);
    // decode() は画面が隠れていると終わらないことがあるので、onload で待つ
    const img = await new Promise((res, rej) => {
      const im = new Image();
      im.onload = () => res(im);
      im.onerror = rej;
      im.src = url;
    });
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
    (resolved.slides || []).forEach((s) => {
      collectImages(s.root, srcs);
      if (s.background?.image) srcs.add(String(s.background.image));
    });
    if (resolved.background?.image) srcs.add(String(resolved.background.image));
    await Promise.all([...srcs].map(loadImage));
    // 選択肢には、フォルダの画像に加えて layout.yaml で使っている画像も入れる（Web 上のフォルダは一覧できないため）
    for (const x of srcs) if (!S.assets.includes(x)) S.assets.push(x);
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
.drop { position: absolute; pointer-events: none; z-index: 11; display: none; }
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
  for (const cls of ['focusdim', 'sel', 'drop']) {
    const d = doc.createElement('div');
    d.className = cls;
    sec.append(d);
  }
  // ドラッグで別のボックスと入れ替える。6 px 以上動かしたらドラッグとみなし、クリックのメニューは出さない
  let drag = null;
  let dragged = false;
  doc.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    const b = boxAt(e);
    if (b && b.depth > 0) {
      e.preventDefault();
      drag = { box: b, x: e.clientX, y: e.clientY, moving: false };
    }
  });
  doc.addEventListener('mousemove', (e) => {
    if (!drag) return;
    if (!drag.moving && Math.hypot(e.clientX - drag.x, e.clientY - drag.y) > 6) {
      drag.moving = true;
      closeMenu();
      doc.body.style.cursor = 'grabbing';
    }
    if (drag.moving) {
      const t = boxAt(e);
      showDrop(t && canSwap(drag.box, t) ? t : null);
    }
  });
  doc.addEventListener('mouseup', (e) => {
    if (!drag) return;
    const d = drag;
    drag = null;
    doc.body.style.cursor = '';
    showDrop(null);
    if (!d.moving) return;
    dragged = true;
    const t = boxAt(e);
    if (t && t.path !== d.box.path) {
      if (canSwap(d.box, t)) swapBoxes(d.box, t);
      else status('このボックスとは入れ替えられません。テンプレートの一部のボックスと、互いに外側と内側の関係にあるボックスは入れ替えられません。', true);
    }
  });
  doc.addEventListener('click', (e) => {
    if (dragged) {
      dragged = false;
      return;
    }
    const box = boxAt(e);
    if (!box) return closeMenu();
    const fr = $('#frame').getBoundingClientRect();
    openMenu(box, fr.left + e.clientX, fr.top + e.clientY);
  });
  doc.addEventListener('keydown', onKey);
  applyView();
  measureOverflow(doc);
}

function showDrop(box) {
  const sec = $('#frame').contentDocument?.querySelector('section.slide');
  const el = sec?.querySelector('.drop');
  if (!el) return;
  if (!box) {
    el.style.display = 'none';
    return;
  }
  const ppu = S.ppu;
  const s = S.view?.s || 1;
  Object.assign(el.style, {
    display: 'block',
    left: `${box.rect.x * ppu}px`,
    top: `${box.rect.y * ppu}px`,
    width: `${box.rect.w * ppu}px`,
    height: `${box.rect.h * ppu}px`,
    outline: `${3 / s}px dashed #ea580c`,
    background: 'rgba(234, 88, 12, .12)',
  });
}

// 入れ替えられるのは、どちらも YAML にそのまま書かれたボックスで、外側と内側の関係にないもの
function canSwap(a, b) {
  if (!a || !b || a.path === b.path || !a.src || !b.src) return false;
  if (a.depth === 0 || b.depth === 0) return false;
  if (a.src.file !== b.src.file || a.src.mode !== 'direct' || b.src.mode !== 'direct') return false;
  const pre = (x, y) => x.length <= y.length && x.every((v, i) => v === y[i]);
  return !pre(a.src.path, b.src.path) && !pre(b.src.path, a.src.path);
}

// 中身を入れ替える。ボックスの位置と大きさ（外側の比率）はそのまま
function swapBoxes(a, b) {
  const file = a.src.file;
  const doc = S.files[file].doc;
  const na = doc.getIn(a.src.path, true) ?? doc.createNode({});
  const nb = doc.getIn(b.src.path, true) ?? doc.createNode({});
  doc.setIn(a.src.path, nb);
  doc.setIn(b.src.path, na);
  S.selected = null;
  return fileChanged(file);
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
  const selPath = S.pending || S.selected;
  const sel = selPath ? findBox(selPath) : null;
  const selEl = sec.querySelector('.sel');
  place(selEl, sel ? sel.rect : null);
  selEl.style.outline = `${2.5 / s}px solid #2563eb`;
  selEl.style.background = 'rgba(37, 99, 235, .08)';
  renderCrumbs();
  $('#btn-focus').disabled = !S.selected || S.selected === S.focus;
  $('#btn-whole').disabled = !S.focus;
}

// クリックした点を含む、いちばん内側の枠
function boxAt(e) {
  const slide = curSlide();
  if (!slide || !S.view) return null;
  const { s, offX, offY, r } = S.view;
  const x = (e.clientX - offX) / s / S.ppu + r.x;
  const y = (e.clientY - offY) / s / S.ppu + r.y;
  let best = null;
  for (const b of slide.boxes) {
    const q = b.rect;
    if (x >= q.x && x <= q.x + q.w && y >= q.y && y <= q.y + q.h && (!best || b.depth >= best.depth)) best = b;
  }
  return best;
}

// 枠をクリックしたときのメニュー：分けるか、編集するか
function openMenu(box, x, y) {
  const m = $('#menu');
  S.pending = box.path;
  applyView();
  m.innerHTML = '';
  m.append(h('div', 'menu-title', `${KIND_LABEL[box.kind]}　${fmt(box.rect.w)} × ${fmt(box.rect.h)} ${S.layout.unit}`));
  if (box.src) {
    const count = selectEl([2, 3, 4, 5, 6].map((n) => [String(n), `${n} つ`]), String(S.splitCount || 2), (v) => (S.splitCount = +v));
    const r = h('div', 'menu-row');
    r.append(h('span', null, '分ける数'), count);
    m.append(
      r,
      btn('上下に分ける', () => {
        closeMenu();
        splitBox(box, 'rows', S.splitCount || 2);
      }),
      btn('左右に分ける', () => {
        closeMenu();
        splitBox(box, 'cols', S.splitCount || 2);
      }),
    );
  }
  const ed = btn('このボックスを編集', () => {
    closeMenu();
    select(box.path);
  });
  ed.className = 'primary';
  m.append(ed);
  if (canRemove(box)) {
    const rm = btn('このボックスを消す', () => {
      closeMenu();
      removeBox(box);
    });
    rm.className = 'danger';
    m.append(rm);
  }
  m.hidden = false;
  const mw = m.offsetWidth;
  const mh = m.offsetHeight;
  m.style.left = `${Math.min(x + 8, window.innerWidth - mw - 8)}px`;
  m.style.top = `${Math.min(y + 8, window.innerHeight - mh - 8)}px`;
}

function closeMenu() {
  const m = $('#menu');
  if (m.hidden) return;
  m.hidden = true;
  S.pending = null;
  applyView();
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
      out.push({ level: 'error', path: box.dataset.path, msg: overW ? '文字が横にはみ出している（長い語）' : `文字がボックスの高さを ${Math.round(overH * 100)}% 超えている` });
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

// 枠を n 個に分ける。中身（と揃え方などの設定）は 1 つ目に移し、
// 枠そのものの設定（id・余白・背景・枠線）は外側に残す。表示の拡大はしない
const FRAME_KEYS = ['id', 'padding', 'background', 'border'];
function splitBox(box, dir, n) {
  const { file, path, mode } = box.src;
  const doc = S.files[file].doc;
  let pn = plain(box.node || {});
  if (typeof pn === 'string') pn = { text: pn };
  const content = Object.fromEntries(Object.entries(pn || {}).filter(([k]) => !FRAME_KEYS.includes(k)));
  const hasContent = box.kind !== 'empty';
  const children = Array.from({ length: n }, (_, i) => (i === 0 && hasContent ? content : {}));
  if (mode === 'direct') {
    const cur = getJS(doc, path);
    const keep = cur && typeof cur === 'object' ? Object.fromEntries(Object.entries(cur).filter(([k]) => FRAME_KEYS.includes(k))) : {};
    replaceNode(doc, path, { ...keep, split: dir, children });
  } else {
    for (const k of CONTENT_KEYS) setProp(doc, path, k, null);
    setProp(doc, path, 'split', dir);
    setProp(doc, path, 'children', children);
  }
  return fileChanged(file);
}

// 枠そのものを消す。分割・格子の中の枠だけが対象。
// 分割の中の枠が 1 つだけ残ったら、分割をやめて、残った枠を外側の枠の位置に置く
function canRemove(box) {
  const parent = box.parent ? findBox(box.parent) : null;
  return !!(parent && parent.src && box.src && (parent.kind === 'split' || parent.kind === 'grid'));
}

function hasContent(box) {
  const slide = curSlide();
  const inside = (b) => b.path === box.path || (b.parent && inside(findBox(b.parent)));
  return slide.boxes.some((b) => (b.kind === 'image' || b.kind === 'text') && inside(b));
}

function removeBox(box) {
  if (!canRemove(box)) return;
  if (hasContent(box) && !confirm('このボックスと中身を消しますか？')) return;
  const parent = findBox(box.parent);
  const listKey = parent.kind === 'split' ? 'children' : 'items';
  const { file, path, mode } = parent.src;
  const doc = S.files[file].doc;
  const pn = parent.node || {};
  const sibs = curSlide().boxes.filter((b) => b.parent === parent.path);
  const i = sibs.findIndex((b) => b.path === box.path);
  // 書いてあるとおりの中身を使う（use でテンプレートを呼んでいる枠を、展開した形で書き直さない）
  const written = mode === 'direct' ? getJS(doc, [...path, listKey]) : null;
  const list = Array.isArray(written) && written.length === sibs.length ? written : plain(pn[listKey] || []);
  list.splice(i, 1);
  const ratio = Array.isArray(pn.ratio) && pn.ratio.length === sibs.length ? [...pn.ratio] : null;
  if (ratio) ratio.splice(i, 1);

  if (parent.kind === 'split' && list.length === 1 && mode === 'direct') {
    // 残った 1 つを外側の枠の位置に戻す（外側の id・余白・背景・枠線は残す）
    const cur = getJS(doc, path) || {};
    const keep = Object.fromEntries(Object.entries(cur).filter(([k]) => FRAME_KEYS.includes(k)));
    let child = list[0];
    if (typeof child === 'string') child = { text: child };
    replaceNode(doc, path, { ...(child || {}), ...keep });
  } else {
    const seq = doc.getIn([...path, listKey], true);
    if (mode === 'direct' && YAML.isSeq(seq) && seq.items.length === sibs.length) {
      seq.items.splice(i, 1);
      if (ratio) setProp(doc, path, 'ratio', ratio);
    } else {
      setProp(doc, path, listKey, list);
      if (ratio) setProp(doc, path, 'ratio', ratio);
    }
  }
  S.selected = null;
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
  renderGlobal();
  if (S.keepPanel && S.keepPanel === S.selected) {
    S.keepPanel = null;
    return;
  }
  S.keepPanel = null;
  const el = $('#node');
  el.innerHTML = '';
  if (!S.layout) return;
  const box = S.selected ? findBox(S.selected) : null;
  el.append(h('h2', null, '選択中のボックス'));
  if (!box) {
    el.append(h('p', 'note', 'プレビューでボックスをクリックし、「このボックスを編集」を選ぶと、ここに設定が出ます。'));
    return;
  }
  const node = box.node || {};
  const inh = box.inherited || {};
  const u = S.layout.unit;
  el.append(h('p', 'meta', `${KIND_LABEL[box.kind]}　${fmt(box.rect.w)} × ${fmt(box.rect.h)} ${u}`));
  if (!box.src) return;
  const mode = box.src.mode;
  if (mode === 'shared') el.append(h('p', 'warn', 'このボックスは、複数のスライドで共通に使うボックスの組み合わせ（テンプレート）の一部です。変えると、このテンプレートを使う全てのスライドが変わります。'));

  // 種類
  const fsSize = sizeFields(box);
  if (fsSize) el.append(fsSize);
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
    ...bgRows(node.background, (v) => edit(box, 'background', v)),
    row('枠線', h2(
      numEl(node.border?.width ?? '', 'pt', (v) => edit(box, 'border', v === '' ? undefined : { width: +v, color: node.border?.color || '#333333' })),
      'pt',
      colorEl(node.border?.color, (v) => node.border && edit(box, ['border', 'color'], v || '#333333')),
    )),
  );
  el.append(fsBox);

  // 操作
  const acts = h('div', 'actions');
  acts.append(btn('このボックスを拡大', () => focusOn(box.path)));
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
  ['16:9 のスライド', 32, 18],
  ['4:3 のスライド', 32, 24],
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

  // 枠を編集しているあいだは折りたたむ。選ぶ枠が変わったときだけ開閉を決め直し、それ以外は利用者の開閉を保つ
  if (S.globalFor !== S.selected) {
    S.globalOpen = !S.selected;
    S.globalFor = S.selected;
  }
  const fold = document.createElement('details');
  fold.open = S.globalOpen;
  fold.ontoggle = () => (S.globalOpen = fold.open);
  const sum = document.createElement('summary');
  sum.append(h('h2', null, 'スライド全体'), h('span', 'sum-size', `${fmt(W)} × ${fmt(H)} ${unit}`));
  const body = h('div');
  fold.append(sum, body);
  el.append(fold);

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
  const ratioOpts = [...(ratioHit < 0 ? [['', `ユーザー設定（${(W / H).toFixed(3)} : 1）`]] : []), ...RATIOS.map(([l], i) => [String(i), l])];
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
  const paperOpts = [...(paperHit < 0 ? [['', 'ユーザー設定']] : []), ...PAPERS.map(([l, w, hh], i) => [String(i), `${l}（${w} × ${hh} cm）`])];
  const paperIn = selectEl(paperOpts, paperHit >= 0 ? String(paperHit) : '', (v) => {
    if (v === '') return;
    const [, w, hh] = PAPERS[+v];
    set([['width', fromCm(w)], ['height', fromCm(hh)]]);
  });

  body.append(
    row('幅', h2(wIn, unit)),
    row('高さ', h2(hIn, unit)),
    row('単位', unitIn),
    row('縦横比', ratioIn),
    row('', lockIn),
    row('用紙', paperIn),
  );
  body.append(...bgRows(getJS(doc, ['background']), (v) => {
    setProp(doc, [], 'background', v);
    fileChanged('layout.yaml');
  }));
  const acts = h('div', 'actions');
  acts.append(btn('縦と横を入れ替える', () => set([['width', H], ['height', W]])));
  body.append(acts);
  for (const i of S.layout.issues || []) el.append(h('p', 'warn', i.msg));
}

// 分割の中の枠の大きさ。外側の分割の ratio のうち、この枠の分を書き換える。
// 数値は比率、"3cm" のような単位付きは固定の長さ（ほかの部分は残りを比率で分ける）
function sizeFields(box) {
  const parent = box.parent ? findBox(box.parent) : null;
  if (!parent || parent.kind !== 'split' || !parent.src) return null;
  const pn = parent.node || {};
  const n = (pn.children || []).length;
  const sibs = curSlide().boxes.filter((b) => b.parent === parent.path);
  const i = sibs.findIndex((b) => b.path === box.path);
  if (i < 0 || sibs.length !== n) return null;
  const u = S.layout.unit;
  const ratio = Array.isArray(pn.ratio) && pn.ratio.length === n ? [...pn.ratio] : Array(n).fill(1);
  const isLen = (v) => typeof v === 'string';
  const rows = pn.split === 'rows';
  const dim = rows ? '高さ' : '幅';
  const lenOf = (b) => (rows ? b.rect.h : b.rect.w);
  const r3 = (v) => Math.round(v * 1000) / 1000;
  const write = (v) => {
    ratio[i] = v;
    edit(parent, 'ratio', ratio);
  };

  const fs = fieldset('このボックスの大きさ');
  const mode = selectEl([['ratio', '比率で指定'], ['len', '長さで指定']], isLen(ratio[i]) ? 'len' : 'ratio', (v) => {
    if (v === 'len') return write(`${r3(lenOf(box))}${u}`);
    // 比率に戻すときは、今の大きさを保つ比にする（比率で決めているほかの部分を基準にする）
    const j = sibs.findIndex((b, k) => k !== i && !isLen(ratio[k]) && ratio[k] > 0);
    write(j >= 0 ? r3((lenOf(box) / lenOf(sibs[j])) * ratio[j]) : 1);
  });
  let valueIn;
  let hint;
  if (isLen(ratio[i])) {
    valueIn = h2(numEl(r3(toUnit(ratio[i], u)), '', (v) => +v > 0 && write(`${+v}${u}`)), u);
    hint = `${u} で固定する。ほかの部分は、残りの長さを比率で分ける`;
  } else {
    valueIn = numEl(ratio[i], '1', (v) => +v > 0 && write(+v));
    const shown = ratio.map((v) => (isLen(v) ? v : String(v))).join(' : ');
    hint = `ほかの部分との比。今は ${shown}`;
  }
  if (parent.src.mode === 'shared') fs.append(h('p', 'warn', 'この大きさは、同じテンプレートを使う全てのスライドで共通です。'));
  fs.append(
    row(`${dim}の指定`, mode),
    row(dim, valueIn, hint),
    h('div', 'note', `今の${dim} ${fmt(lenOf(box))} ${u}`),
  );
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
  fs.append(
    row('文字', richEditor(String(node.text ?? ''), (v) => {
      S.keepPanel = box.path;
      edit(box, 'text', v);
    })),
    row('大きさ', h2(numEl(f.size ?? '', String(fi.size ?? 18), (v) => edit(box, ['font', 'size'], v === '' ? undefined : +v)), 'pt')),
    row('フォント', selectEl(withInherit(FONT_OPTS, fi.family, 'Arial'), f.family ?? '', (v) => edit(box, ['font', 'family'], v || undefined))),
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
  const fs = fieldset('中のボックスすべてに効く設定');
  const f = node.font || {};
  fs.append(
    row('横の揃え方', selectEl(withInherit(H_OPTS, inh.align, '画像は中央、文字は左'), node.align ?? '', (v) => edit(box, 'align', v || undefined))),
    row('縦の揃え方', selectEl(withInherit(V_OPTS, inh.valign, '画像は中央、文字は上'), node.valign ?? '', (v) => edit(box, 'valign', v || undefined))),
    row('画像の収め方', selectEl(withInherit(FIT_OPTS, inh.fit, 'contain'), node.fit ?? '', (v) => edit(box, 'fit', v || undefined))),
    row('文字の大きさ', h2(numEl(f.size ?? '', String(inh.font?.size ?? 18), (v) => edit(box, ['font', 'size'], v === '' ? undefined : +v)), 'pt')),
    row('フォント', selectEl(withInherit(FONT_OPTS, inh.font?.family, 'Arial'), f.family ?? '', (v) => edit(box, ['font', 'family'], v || undefined))),
  );
  return fs;
}

// ---------------------------------------------------------------- 部品

// 文字の欄。ドラッグで選んだ範囲を、ボタンで太字・斜体・上付き・下付きにする。
// 中身は <b> <i> <sup> <sub> の付いた文字として layout.yaml に書く
function richEditor(value, onCommit) {
  const wrap = h('div', 'rich-wrap');
  const bar = h('div', 'rich-bar');
  const ed = document.createElement('div');
  ed.className = 'rich';
  ed.contentEditable = 'true';
  ed.innerHTML = richToHTML(value);
  let last = value;
  let timer = null;
  const commit = () => {
    clearTimeout(timer);
    const v = serializeRich(ed);
    if (v === last) return;
    last = v;
    onCommit(v);
  };
  for (const [cmd, label, title] of [['bold', 'B', '太字'], ['italic', 'I', '斜体'], ['superscript', 'x²', '上付き'], ['subscript', 'x₂', '下付き']]) {
    const b = h('button', `rich-btn rich-${cmd}`, label);
    b.type = 'button';
    b.title = title;
    b.onmousedown = (e) => {
      e.preventDefault(); // 文字の欄の選択を保つ
      document.execCommand('styleWithCSS', false, false);
      document.execCommand(cmd);
      commit();
    };
    bar.append(b);
  }
  ed.oninput = () => {
    clearTimeout(timer);
    timer = setTimeout(commit, 600);
  };
  ed.onblur = commit;
  ed.onpaste = (e) => {
    e.preventDefault();
    document.execCommand('insertText', false, e.clipboardData.getData('text/plain'));
  };
  wrap.append(bar, ed);
  return wrap;
}

function serializeRich(root) {
  let out = '';
  const TAGS = { B: 'b', STRONG: 'b', I: 'i', EM: 'i', SUP: 'sup', SUB: 'sub' };
  const walk = (n) => {
    for (const c of n.childNodes) {
      if (c.nodeType === 3) out += c.nodeValue.replace(/ /g, ' ');
      else if (c.nodeName === 'BR') out += '\n';
      else if (c.nodeType === 1) {
        const block = c.nodeName === 'DIV' || c.nodeName === 'P';
        if (block && out && !out.endsWith('\n')) out += '\n';
        const t = TAGS[c.nodeName];
        if (t) out += `<${t}>`;
        walk(c);
        if (t) out += `</${t}>`;
      }
    }
  };
  walk(root);
  return out.replace(/<(b|i|sup|sub)><\/\1>/g, '').replace(/\n$/, '');
}

// 背景の欄。値は "#RRGGBB"（色）か { image, mode: tile | cover, tile, color }
const BG_OPTS = [['', 'なし'], ['color', '色'], ['tile', '画像を並べる（模様）'], ['cover', '画像を広げる']];
function bgRows(cur, write) {
  const kind = !cur ? '' : typeof cur !== 'object' ? 'color' : cur.image ? (cur.mode === 'cover' ? 'cover' : 'tile') : cur.color ? 'color' : '';
  const rows = [
    row('背景', selectEl(BG_OPTS, kind, (v) => {
      if (v === '') return write(undefined);
      if (v === 'color') return write(typeof cur === 'string' ? cur : cur?.color || '#ffffff');
      const image = (typeof cur === 'object' && cur?.image) || S.assets[0];
      if (!image) {
        status('フォルダに画像がありません', true);
        return renderProps();
      }
      const keep = typeof cur === 'object' && cur ? cur : {};
      write(v === 'tile' ? { ...keep, image, mode: 'tile' } : { image, mode: 'cover', ...(keep.color ? { color: keep.color } : {}) });
    })),
  ];
  if (kind === 'color') rows.push(row('色', colorEl(typeof cur === 'string' ? cur : cur.color, (v) => write(v || undefined))));
  if (kind === 'tile' || kind === 'cover') {
    const opts = [...new Set([...S.assets, cur.image])].map((a) => [a, a]);
    rows.push(row('画像', selectEl(opts, cur.image, (v) => write({ ...cur, image: v }))));
    if (kind === 'tile') {
      const u = S.layout.unit;
      const nat = S.imageSizes.get(cur.image);
      const dflt = nat ? fmt((nat.w / 96) * PER_INCH[u]) : '';
      const { tile, ...rest } = cur;
      rows.push(row('模様の幅', h2(numEl(tile ?? '', dflt, (v) => write(v === '' ? rest : { ...cur, tile: +v })), u), '模様 1 つの幅。空けておくと、画像を 96 dpi とみなした幅になる'));
    }
  }
  return rows;
}

const KIND_LABEL = { empty: '空き', split: '分割', grid: '格子', image: '画像', text: '文字' };
const KIND_OPTS = [['empty', '空き（何も置かない）'], ['split', '分割（中を上下か左右に分ける）'], ['grid', '格子（同じ大きさで縦横に並べる）'], ['image', '画像'], ['text', '文字']];
const FIT_OPTS = [['contain', '切らずに全体を収める'], ['cover', 'ボックスいっぱいに広げ、はみ出た部分を切る'], ['stretch', '縦横比を変えてボックスに合わせる']];
const H_OPTS = [['left', '左'], ['center', '中央'], ['right', '右']];
const FONT_OPTS = FONTS.map((f) => [f, f]);
const V_OPTS = [['top', '上'], ['middle', '中央'], ['bottom', '下']];

function withInherit(opts, inherited, dflt) {
  const label = inherited != null ? `（外側のボックスの設定：${(opts.find((o) => o[0] === inherited) || [, inherited])[1]}）` : `（指定なし：${(opts.find((o) => o[0] === dflt) || [, dflt])[1]}）`;
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
const fmt = (v) => (+v).toFixed(2).replace(/\.?0+$/, '');
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
    const blob = await finishPPTX(pres, window.JSZip, 'blob');
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
    await Promise.all([...f.contentDocument.images].map((i) => (i.complete ? null : new Promise((r) => { i.onload = i.onerror = r; }))));
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
  if (e.key === 'Escape' && !$('#menu').hidden) return closeMenu();
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
$('#btn-save').onclick = save;
$('#btn-pptx').onclick = exportPPTX;
$('#btn-pdf').onclick = exportPDF;
$('#btn-focus').onclick = () => S.selected && focusOn(S.selected);
$('#btn-whole').onclick = () => focusOn(null);
$('#chk-outline').onchange = (e) => {
  S.outline = e.target.checked;
  renderPreview();
};
document.addEventListener('keydown', onKey);
document.addEventListener('mousedown', (e) => {
  if (!$('#menu').contains(e.target)) closeMenu();
});
window.addEventListener('resize', () => applyView());
window.addEventListener('beforeunload', (e) => {
  if (anyDirty()) e.preventDefault();
});
initReopen();
window.slideCreator = { state: S, buildPPTX, finishPPTX }; // 開発者ツールからの確認用
// 開発用：URL に ?demo を付けると examples/demo を開く（画面にボタンは出さない）
if (new URLSearchParams(location.search).has('demo')) openSource(new UrlSource('../examples/demo/', 'demo'));
