// slide-creator ブラウザ版
import { resolveDoc } from '../src/core/resolve.js';
import { layoutDoc } from '../src/core/layout.js';
import { renderHTML, loadUsedFonts, FONTS, JA_FONTS, DEFAULT_FONT, DEFAULT_FONT_JA } from '../src/core/html.js';
import { richToHTML, paragraphs } from '../src/core/richtext.js';
import { buildPPTX, finishPPTX } from '../src/core/pptx.js';
import { PER_INCH, toUnit, defaultPadding, parseBox } from '../src/core/units.js';
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
  selected: null, // 1 つだけ選んでいるときの path
  sel: [], // 選択中のボックスの path（複数のこともある）
  zoom: 1,
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
    S.saved = Object.fromEntries(Object.entries(files).map(([k, f]) => [k, f.text]));
    S.hist = [];
    S.histI = -1;
    pushHistory();
    S.assets = (await source.list('')).filter((f) => IMAGE_RE.test(f) && !f.startsWith('out/'));
    S.imageSizes.clear();
    S.imageURLs.clear();
    S.slideId = null;
    setSel([]);
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
  if (node.icon) out.add(String(node.icon));
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
  if ((S.sel || []).some((x) => !findBox(x))) setSel([]);
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
  const slides = S.layout?.slides || [];
  if (!slides.length) return;
  const sel = selectEl(slides.map((s, i) => [s.id, `${i + 1}ページ`]), S.slideId, (v) => {
    S.slideId = v;
    setSel([]);
    setZoom(1);
    renderPreview();
    renderProps();
  });
  sel.className = 'page-select';
  el.append(sel);
}

const FRAME_CSS = `
html { background: #dfe3e8; }
body { background: transparent !important; overflow: auto; margin: 0; }
.zoom-space { position: absolute; left: 0; top: 0; pointer-events: none; }
section.slide { position: absolute; left: 0; top: 0; box-shadow: 0 1px 6px rgba(0,0,0,.25); }
.sel, .drop, .hov { position: absolute; pointer-events: none; display: none; }
.hov { z-index: 9; }
.sel { z-index: 10; }
.drop { z-index: 11; }
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

// 選択。1 つならその枠の設定、2 つ以上なら比率と背景を出す
function setSel(paths) {
  S.sel = paths;
  S.selected = paths.length === 1 ? paths[0] : null;
}

function select(path) {
  setSel(path ? [path] : []);
  applyView();
  renderProps();
}

function setZoom(z) {
  S.zoom = z;
  $('#zoom').value = String(Math.round(z * 100));
  $('#zoom-val').textContent = `${Math.round(z * 100)}%`;
}

function setupFrame() {
  const frame = $('#frame');
  const doc = frame.contentDocument;
  const sec = doc?.querySelector('section.slide');
  if (!sec) return;
  sec.style.transform = 'none';
  S.ppu = sec.getBoundingClientRect().width / S.layout.width; // キャンバスの単位あたりの px
  for (const cls of ['hov', 'sel', 'drop']) {
    const d = doc.createElement('div');
    d.className = cls;
    sec.append(d);
  }
  // 拡大したときにスクロールできる広さを作る
  const space = doc.createElement('div');
  space.className = 'zoom-space';
  doc.body.append(space);
  // 書き換えて描き直しても、見ていた位置を保つ
  doc.defaultView.addEventListener('scroll', () => {
    S.scroll = { x: doc.defaultView.scrollX, y: doc.defaultView.scrollY };
  });
  // クリックで 1 つ選ぶ。ドラッグで、長方形になるように複数選ぶ。
  // ⌘（cmd。Windows では Ctrl）を押しながらドラッグすると、離した先のボックスと中身を入れ替える
  let drag = null;
  let dragged = false;
  doc.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    const b = leafAt(e);
    if (!b) return;
    e.preventDefault();
    drag = { box: b, x: e.clientX, y: e.clientY, moving: false, swap: e.metaKey || e.ctrlKey };
  });
  doc.addEventListener('mousemove', (e) => {
    if (!drag) return hover(leafAt(e));
    if (!drag.moving && Math.hypot(e.clientX - drag.x, e.clientY - drag.y) > 6) {
      drag.moving = true;
      hover(null);
      if (drag.swap) doc.body.style.cursor = 'grabbing';
    }
    if (!drag.moving) return;
    const t = leafAt(e);
    if (drag.swap) {
      showDrop(t && canSwap(drag.box, t) ? t : null);
    } else if (t) {
      S.dragSel = rectSelect(drag.box, t);
      applyView();
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
    if (d.swap) {
      const t = leafAt(e);
      if (t && t.path !== d.box.path) {
        if (canSwap(d.box, t)) swapBoxes(d.box, t);
        else status('このボックスとは入れ替えられません。テンプレートの一部のボックスは入れ替えられません。', true);
      }
      return;
    }
    const r = S.dragSel;
    S.dragSel = null;
    if (r) setSel(r.leaves.map((l) => l.path));
    applyView();
    renderProps();
  });
  doc.addEventListener('mouseleave', () => hover(null));
  doc.addEventListener('click', (e) => {
    if (dragged) {
      dragged = false;
      return;
    }
    const b = leafAt(e);
    select(b ? b.path : null);
  });
  doc.addEventListener('keydown', onKey);
  applyView();
  measureOverflow(doc);
  // 文字の高さに合わせ直すボックスを、1 つずつ処理する（書き換えると描き直され、ここに戻って次を処理する）
  while (S.pendingFit && S.pendingFit.length) {
    const p = S.pendingFit.shift();
    if (fitTextBelow(p, doc)) break;
  }
}

// テキストのボックスを文字の高さに合わせ、下に余った部分を別のボックス（空き）にする。
// すでにそうしてあれば（上下の分割で、このボックスの高さが長さで決まっていて、すぐ下が空きのボックス）、高さだけ合わせ直す
function fitTextBelow(path, doc) {
  const box = findBox(path);
  if (!box || box.kind !== 'text' || !box.src) return;
  const node = box.node || {};
  if (node.split_below === false || !String(node.text ?? '').trim()) return;
  const v = box.style?.valign || 'top';
  if (v !== 'top') return; // 中央・下に揃えた文字は、下に余白を作らない
  const el = [...doc.querySelectorAll('.txt')].find((x) => x.dataset.path === path);
  if (!el || !el.firstElementChild) return;
  const u = S.layout.unit;
  const textH = el.firstElementChild.getBoundingClientRect().height / (S.ppu * S.view.s);
  const pad = parseBox(node.padding ?? defaultPadding('text'), u);
  const need = Math.ceil((textH + pad.t + pad.b) * 100) / 100;
  const len = `${need}${u}`;
  const parent = box.parent ? findBox(box.parent) : null;
  const sibs = parent ? curSlide().boxes.filter((b) => b.parent === parent.path) : [];
  const i = sibs.findIndex((b) => b.path === box.path);
  const ratio0 = parent && Array.isArray(parent.node?.ratio) && parent.node.ratio.length === sibs.length ? parent.node.ratio : null;
  const ours = parent && parent.kind === 'split' && parent.node.split === 'rows' && parent.src && ratio0
    && typeof ratio0[i] === 'string' && sibs[i + 1] && sibs[i + 1].kind === 'empty';
  if (ours) {
    // 下の空きのボックスが無くならない範囲で、高さを合わせる
    const avail = box.rect.h + sibs[i + 1].rect.h;
    const h = Math.min(need, Math.floor(avail * 0.95 * 100) / 100);
    if (Math.abs(toUnit(ratio0[i], u) - h) < 0.01) return;
    const ratio = [...ratio0];
    ratio[i] = `${h}${u}`;
    return edit(parent, 'ratio', ratio);
  }
  // 余りが 0.5 cm に満たないときは分けない
  if (box.rect.h - need < toUnit('0.5cm', u)) return;
  const { file, path: ypath, mode } = box.src;
  const ydoc = S.files[file].doc;
  let pn = writtenNode(box);
  if (typeof pn === 'string') pn = { text: pn };
  const content = Object.fromEntries(Object.entries(pn).filter(([k]) => !FRAME_KEYS.includes(k)));
  const children = [content, {}];
  if (mode === 'direct') {
    const keep = Object.fromEntries(Object.entries(pn).filter(([k]) => FRAME_KEYS.includes(k)));
    replaceNode(ydoc, ypath, { ...keep, split: 'rows', ratio: [len, 1], children });
  } else {
    for (const k of CONTENT_KEYS) setProp(ydoc, ypath, k, null);
    for (const k of Object.keys(content)) if (!CONTENT_KEYS.includes(k)) setProp(ydoc, ypath, k, null);
    setProp(ydoc, ypath, 'split', 'rows');
    setProp(ydoc, ypath, 'ratio', [len, 1]);
    setProp(ydoc, ypath, 'children', children);
  }
  // 選択は、文字の入ったボックス（分けた上側）に移す（複数選んでいるときは、そのボックスだけを置き換える）
  setSel((S.sel || []).map((x) => (x === path ? `${path}.0` : x)));
  return fileChanged(file);
}

// 複数のテキストのボックスに「テキストの下に生じた余白を別のボックスにする」を設定する
async function setSplitBelowMany(paths, on) {
  if (on) {
    const files = new Set();
    for (const p of paths) {
      const b = findBox(p);
      if (!b || !b.src) continue;
      setProp(S.files[b.src.file].doc, b.src.path, 'split_below', undefined);
      files.add(b.src.file);
    }
    // 描き直したあと、1 つずつ文字の高さに合わせる
    S.pendingFit = [...paths];
    for (const f of files) await fileChanged(f);
    return;
  }
  for (const p of paths) {
    const b = findBox(p);
    if (b && b.src) await unfitTextBelow(b);
  }
}

// 「テキストの下に生じた余白を別のボックスにする」を外したら、下の空きのボックスを消して元の 1 つに戻す
function unfitTextBelow(box) {
  const parent = box.parent ? findBox(box.parent) : null;
  const sibs = parent ? curSlide().boxes.filter((b) => b.parent === parent.path) : [];
  const i = sibs.findIndex((b) => b.path === box.path);
  const { file, path } = box.src;
  setProp(S.files[file].doc, path, 'split_below', false);
  const ratio0 = parent && Array.isArray(parent.node?.ratio) ? parent.node.ratio : null;
  if (parent && parent.kind === 'split' && parent.node.split === 'rows' && ratio0 && typeof ratio0[i] === 'string' && sibs[i + 1] && sibs[i + 1].kind === 'empty') {
    const keep = parent.path;
    const before = S.sel || [];
    const p = removeFromParent(parent, [i + 1]);
    // 1 つに戻ったときは外側の位置に、そうでなければ同じ場所を選んだままにする（ほかに選んでいるボックスはそのまま）
    setSel(before.map((x) => (x === box.path ? (sibs.length === 2 ? keep : box.path) : x)));
    return p;
  }
  return fileChanged(file);
}

function placeIn(el, rect) {
  if (!el) return;
  el.style.display = rect ? 'block' : 'none';
  if (!rect) return;
  const ppu = S.ppu;
  Object.assign(el.style, { left: `${rect.x * ppu}px`, top: `${rect.y * ppu}px`, width: `${rect.w * ppu}px`, height: `${rect.h * ppu}px` });
}

function frameEl(cls) {
  return $('#frame').contentDocument?.querySelector(`section.slide .${cls}`);
}

// 選んでいないボックスにマウスを乗せたら、青く示す
function hover(box) {
  const el = frameEl('hov');
  if (!el) return;
  const show = box && !(S.sel || []).includes(box.path);
  placeIn(el, show ? box.rect : null);
  const s = S.view?.s || 1;
  el.style.outline = `${1.5 / s}px solid rgba(37, 99, 235, .7)`;
  el.style.background = 'rgba(37, 99, 235, .14)';
}

function showDrop(box) {
  const el = frameEl('drop');
  if (!el) return;
  placeIn(el, box ? box.rect : null);
  const s = S.view?.s || 1;
  el.style.outline = `${3 / s}px dashed #ea580c`;
  el.style.background = 'rgba(234, 88, 12, .12)';
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
  setSel([]);
  return fileChanged(file);
}

const isLeaf = (b) => b.kind !== 'split' && b.kind !== 'grid';

function unionRect(rs) {
  const x = Math.min(...rs.map((r) => r.x));
  const y = Math.min(...rs.map((r) => r.y));
  const x2 = Math.max(...rs.map((r) => r.x + r.w));
  const y2 = Math.max(...rs.map((r) => r.y + r.h));
  return { x, y, w: x2 - x, h: y2 - y };
}

function overlaps(q, r) {
  const e = 1e-6;
  return q.x < r.x + r.w - e && q.x + q.w > r.x + e && q.y < r.y + r.h - e && q.y + q.h > r.y + e;
}

// 2 つのボックスを含む長方形を作り、その長方形にかかるボックスが入るまで広げる
function rectSelect(a, b) {
  const leaves = curSlide().boxes.filter(isLeaf);
  let r = unionRect([a.rect, b.rect]);
  for (let k = 0; k < 50; k++) {
    const inc = leaves.filter((l) => overlaps(l.rect, r));
    const r2 = unionRect(inc.map((l) => l.rect));
    const same = Math.abs(r2.x - r.x) + Math.abs(r2.y - r.y) + Math.abs(r2.w - r.w) + Math.abs(r2.h - r.h) < 1e-6;
    r = r2;
    if (same) return { rect: r, leaves: inc };
  }
  return { rect: r, leaves: leaves.filter((l) => overlaps(l.rect, r)) };
}

function selectionRect() {
  const bs = (S.sel || []).map(findBox).filter(Boolean);
  return bs.length ? unionRect(bs.map((b) => b.rect)) : null;
}

// 拡大・位置合わせ。拡大は、いま見えている左上の点を動かさずに右下へ広げる。拡大したらスクロールで全体を見られる
function applyView() {
  const frame = $('#frame');
  const doc = frame.contentDocument;
  const sec = doc?.querySelector('section.slide');
  const slide = curSlide();
  if (!sec || !slide) return;
  const win = doc.defaultView;
  const R = slide.rect;
  const ppu = S.ppu;
  const m = 24;
  const fw = frame.clientWidth;
  const fh = frame.clientHeight;
  const base = Math.min((fw - 2 * m) / (R.w * ppu), (fh - 2 * m) / (R.h * ppu));
  const s = base * (S.zoom || 1);
  const W = R.w * ppu * s;
  const H = R.h * ppu * s;
  // 収まる向きは中央に、はみ出す向きは余白 m から並べる
  const offX = W + 2 * m <= fw ? (fw - W) / 2 : m;
  const offY = H + 2 * m <= fh ? (fh - H) / 2 : m;
  sec.style.transformOrigin = '0 0';
  sec.style.transform = `translate(${offX}px, ${offY}px) scale(${s})`;
  const prev = S.view;
  let sx = S.scroll?.x || 0;
  let sy = S.scroll?.y || 0;
  const t = S.scrollTarget;
  if (t) {
    // 虫眼鏡：選んだボックスの左上を、プレビューの左上（余白 m）に置く。右下の端のボックスでも寄せられるよう、スクロールできる広さを足す
    sx = offX + t.x * ppu * s - m;
    sy = offY + t.y * ppu * s - m;
    S.extraSpace = { w: sx + fw, h: sy + fh };
    S.scrollTarget = null;
  }
  const space = doc.querySelector('.zoom-space');
  const ex = S.extraSpace || { w: 0, h: 0 };
  if (space) Object.assign(space.style, { width: `${Math.max(fw, W + 2 * m, ex.w)}px`, height: `${Math.max(fh, H + 2 * m, ex.h)}px` });
  if (!t && prev && Math.abs(prev.s - s) > 1e-9) {
    // 見えている左上の点（スライドの上の位置）を、拡大の前後で同じ場所に置く
    // 左上の角がまだ見えているとき（100% から広げるときなど）は、左上の角から右下へ広げる
    const px = (sx - prev.offX) / prev.s;
    const py = (sy - prev.offY) / prev.s;
    sx = px > 0 ? px * s + offX : 0;
    sy = py > 0 ? py * s + offY : 0;
  }
  win.scrollTo(sx, sy);
  S.scroll = { x: win.scrollX, y: win.scrollY };
  S.view = { s, offX, offY, r: { x: 0, y: 0 } };
  const selRect = S.dragSel?.rect || selectionRect();

  const selEl = sec.querySelector('.sel');
  placeIn(selEl, selRect);
  selEl.style.outline = `${2.5 / s}px solid #2563eb`;
  selEl.style.background = 'rgba(37, 99, 235, .08)';
  const one = S.selected ? findBox(S.selected) : null;
  $('#btn-split').disabled = !(one && one.src && isLeaf(one));
  $('#btn-zoom-fit').disabled = !(S.sel || []).length;
  $('#btn-dup').disabled = !dupPlan();
  $('#btn-merge').disabled = !((S.sel || []).length > 1 && selectionUnit());
  $('#btn-delete').disabled = !((S.sel || []).length > 1 ? selectionUnit() : one && canRemove(one));
}

// クリックした点を含む、いちばん内側のボックス（分割・格子の間の隙間なら null）
function leafAt(e) {
  const slide = curSlide();
  if (!slide || !S.view) return null;
  const { s, offX, offY } = S.view;
  const x = (e.pageX - offX) / s / S.ppu;
  const y = (e.pageY - offY) / s / S.ppu;
  let best = null;
  for (const b of slide.boxes) {
    const q = b.rect;
    if (x >= q.x && x <= q.x + q.w && y >= q.y && y <= q.y + q.h && (!best || b.depth >= best.depth)) best = b;
  }
  return best && isLeaf(best) ? best : null;
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
  f.dirty = f.text !== S.saved?.[file];
  pushHistory();
  return rebuild();
}

// 書き換えるたびに、全てのファイルの中身を記録する（⌘Z で戻り、⌘⇧Z でやり直す）
function snapshot() {
  return Object.fromEntries(Object.entries(S.files).map(([k, f]) => [k, f.text]));
}

function pushHistory() {
  S.hist = (S.hist || []).slice(0, (S.histI ?? -1) + 1);
  S.hist.push(snapshot());
  if (S.hist.length > 300) S.hist.shift();
  S.histI = S.hist.length - 1;
}

function restoreHistory(i) {
  const snap = S.hist[i];
  S.histI = i;
  for (const [k, text] of Object.entries(snap)) S.files[k] = { text, doc: parse(text), dirty: text !== S.saved?.[k] };
  S.keepPanel = null;
  return rebuild();
}

function undo() {
  if (!S.hist || S.histI <= 0) return status('これ以上は戻せません');
  restoreHistory(S.histI - 1).then(() => status('元に戻しました'));
}

function redo() {
  if (!S.hist || S.histI >= S.hist.length - 1) return status('これ以上はやり直せません');
  restoreHistory(S.histI + 1).then(() => status('やり直しました'));
}

// 選んだ枠の設定を 1 つ変える
const editOne = (box, key, value) => edit(box, key, value);
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

function changeKind(box, kind, imagePath) {
  const { file, path, mode } = box.src;
  const doc = S.files[file].doc;
  const fresh = {
    empty: {},
    split: { split: 'rows', children: [{}, {}] },
    grid: { grid: { rows: 2, cols: 2 }, items: [{}, {}, {}, {}] },
    image: { image: imagePath || S.assets[0] || '' },
    text: { text: '' },
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
// 外側の枠に残す設定（余白は残さない。外側の枠はクリックで選べず、見えない余白になるため）
const FRAME_KEYS = ['id', 'background', 'border'];
const MOVE_TO_CHILD = ['padding', 'background_inset'];
function splitBox(box, dir, n) {
  const { file, path, mode } = box.src;
  const doc = S.files[file].doc;
  let pn = plain(box.node || {});
  if (typeof pn === 'string') pn = { text: pn };
  const content = Object.fromEntries(Object.entries(pn || {}).filter(([k]) => !FRAME_KEYS.includes(k) && !MOVE_TO_CHILD.includes(k)));
  const pad = pn && typeof pn === 'object' && pn.padding != null ? { padding: pn.padding } : {};
  const hasContent = box.kind !== 'empty';
  // 余白は、分けた内側のボックスの 1 つずつに写す（選べば欄に見え、そこで直せる）
  const children = Array.from({ length: n }, (_, i) => ({ ...(i === 0 && hasContent ? content : {}), ...structuredClone(pad) }));
  if (mode === 'direct') {
    const cur = getJS(doc, path);
    const keep = cur && typeof cur === 'object' ? Object.fromEntries(Object.entries(cur).filter(([k]) => FRAME_KEYS.includes(k))) : {};
    replaceNode(doc, path, { ...keep, split: dir, children });
  } else {
    for (const k of CONTENT_KEYS) setProp(doc, path, k, null);
    if (pad.padding != null) setProp(doc, path, 'padding', 0);
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

// ---- 分割の中の、選んだ範囲を正確に扱う
// 分割 lca の部分を、選んだボックス selSet で「全部選んだ／選んでいない」に分けて並べる。
// 一部だけ選んだ部分が同じ向きの分割なら、その中身を lca の並びに展開する（例：タイトル＋下の空き の組の、空きだけを選んだとき）。
// 返す並び：[{ node（YAML に書く中身）, ratio（比率の数か長さの文字列）, len（いまの長さ）, sel }]
const round4 = (x) => Math.round(x * 10000) / 10000;
const round2 = (x) => Math.round(x * 100) / 100;
function within(b, box) {
  for (let x = b; x; x = x.parent ? findBox(x.parent) : null) if (x.path === box.path) return true;
  return false;
}
function splitEntries(lca, selSet) {
  const u = S.layout.unit;
  const dir = lca.node.split;
  const L = (b) => (dir === 'rows' ? b.rect.h : b.rect.w);
  const slide = curSlide();
  const kids = slide.boxes.filter((b) => b.parent === lca.path);
  const n = kids.length;
  const doc = S.files[lca.src.file].doc;
  const written = lca.src.mode === 'direct' ? getJS(doc, [...lca.src.path, 'children']) : null;
  const nodes = Array.isArray(written) && written.length === n ? written : plain(lca.node.children || []);
  const ratio = Array.isArray(lca.node.ratio) && lca.node.ratio.length === n ? lca.node.ratio : Array(n).fill(1);
  const gap = toUnit(lca.node.gap ?? 0, u);
  // 比率 1 あたりの長さ
  let k = null;
  kids.forEach((b, i) => {
    if (k == null && typeof ratio[i] === 'number' && ratio[i] > 0) k = L(b) / ratio[i];
  });
  const state = (box) => {
    const ls = slide.boxes.filter((b) => isLeaf(b) && within(b, box));
    const c = ls.filter((l) => selSet.has(l.path)).length;
    return c === 0 ? 'none' : c === ls.length ? 'all' : 'part';
  };
  const out = [];
  const add = (box, node, r) => {
    const st = state(box);
    if (st !== 'part') {
      out.push({ node, ratio: r, len: L(box), sel: st === 'all' });
      return true;
    }
    // 一部だけ選んだ部分：同じ向きの分割で、名前・余白・背景などが付いていなければ、中身を展開する
    const nd = node && typeof node === 'object' ? node : null;
    const ok = ['split', 'ratio', 'children', 'gap', 'id', 'split_below'];
    if (!nd || nd.split !== dir || !Array.isArray(nd.children) || Object.entries(nd).some(([x, v]) => v != null && !ok.includes(x))) return false;
    if (Math.abs(toUnit(nd.gap ?? 0, u) - gap) > 1e-6) return false;
    const subs = slide.boxes.filter((b) => b.parent === box.path);
    if (subs.length !== nd.children.length) return false;
    const sr = Array.isArray(nd.ratio) && nd.ratio.length === subs.length ? nd.ratio : Array(subs.length).fill(1);
    for (let j = 0; j < subs.length; j++) {
      const r2 = typeof sr[j] === 'string' ? sr[j] : k ? round4(L(subs[j]) / k) : `${round2(L(subs[j]))}${u}`;
      if (!add(subs[j], nd.children[j], r2)) return false;
    }
    return true;
  };
  for (let i = 0; i < n; i++) if (!add(kids[i], nodes[i], ratio[i])) return null;
  const idx = out.map((e, i) => (e.sel ? i : -1)).filter((i) => i >= 0);
  if (!idx.length) return null;
  const i0 = idx[0];
  const i1 = idx[idx.length - 1];
  if (i1 - i0 + 1 !== idx.length) return null; // 選んだ部分が続いていない
  return { entries: out, i0, i1, gap, k, u };
}

// 左右（上下）に並ぶ部分が、どれも同じ位置で上下（左右）に分かれているとき、並びを組み替える。
// 例：[CA の上下の組, PCA の上下の組, ZPCA の上下の組] を左右に並べたもの
//   → [上段（CA・PCA・ZPCA を左右に）, 下段（下の空き 3 つを左右に）] を上下に並べたもの。見た目は変わらない
// 組み替えられないときは null を返す
function transposeSplit(lca) {
  if (!lca || lca.kind !== 'split' || !lca.src) return null;
  const D = lca.node.split;
  const P = D === 'rows' ? 'cols' : 'rows';
  const L = (b) => (P === 'rows' ? b.rect.h : b.rect.w);
  const slide = curSlide();
  const kids = slide.boxes.filter((b) => b.parent === lca.path);
  const n = kids.length;
  const { file, path, mode } = lca.src;
  const doc = S.files[file].doc;
  const written = mode === 'direct' ? getJS(doc, [...path, 'children']) : null;
  const nodes = Array.isArray(written) && written.length === n ? written : plain(lca.node.children || []);
  const ok = ['split', 'ratio', 'children', 'gap', 'id', 'split_below'];
  let m = null;
  let gap0 = null;
  const subs = [];
  for (let i = 0; i < n; i++) {
    const nd = nodes[i];
    if (!nd || typeof nd !== 'object' || nd.split !== P || !Array.isArray(nd.children)) return null;
    if (Object.entries(nd).some(([x, v]) => v != null && !ok.includes(x))) return null;
    const g = toUnit(nd.gap ?? 0, S.layout.unit);
    if (m == null) {
      m = nd.children.length;
      gap0 = g;
    }
    if (nd.children.length !== m || Math.abs(g - gap0) > 1e-6) return null;
    const sb = slide.boxes.filter((b) => b.parent === kids[i].path);
    if (sb.length !== m) return null;
    subs.push(sb);
  }
  if (!m || m < 2) return null;
  // 分かれ目の位置がそろっていること
  for (let i = 1; i < n; i++) for (let j = 0; j < m; j++) if (Math.abs(L(subs[i][j]) - L(subs[0][j])) > 0.02) return null;
  const r0 = Array.isArray(nodes[0].ratio) && nodes[0].ratio.length === m ? nodes[0].ratio : null;
  const bands = Array.from({ length: m }, (_, j) => ({
    split: D,
    ...(Array.isArray(lca.node.ratio) ? { ratio: plain(lca.node.ratio) } : {}),
    ...(lca.node.gap != null ? { gap: lca.node.gap } : {}),
    children: nodes.map((nd) => nd.children[j]),
  }));
  const content = { split: P, ...(r0 ? { ratio: r0 } : {}), ...(nodes[0].gap != null ? { gap: nodes[0].gap } : {}), children: bands };
  if (mode === 'direct') {
    const cur = getJS(doc, path) || {};
    const keep = Object.fromEntries(Object.entries(cur).filter(([x]) => !CONTENT_KEYS.includes(x) && x !== 'gap'));
    replaceNode(doc, path, { ...keep, ...content });
  } else {
    setProp(doc, path, 'ratio', null);
    setProp(doc, path, 'gap', null);
    for (const [x, v] of Object.entries(content)) setProp(doc, path, x, v);
  }
  return fileChanged(file);
}

// 選んだ範囲がそのままでは扱えないとき、並びを組み替えてから、同じ範囲を選び直して fn をやり直す
function retryAfterTranspose(lca, fn) {
  const r = selectionRect();
  const done = transposeSplit(lca);
  if (!done) return null;
  return done.then(() => {
    const e = 1e-6;
    const inside = curSlide().boxes.filter((b) => isLeaf(b) && b.rect.x >= r.x - e && b.rect.y >= r.y - e && b.rect.x + b.rect.w <= r.x + r.w + e && b.rect.y + b.rect.h <= r.y + r.h + e);
    setSel(inside.map((b) => b.path));
    applyView();
    renderProps();
    return fn();
  });
}

// ---- 複製して追加
// 中身の無いボックス（空き）か
function isEmptyNode(n) {
  if (n == null) return true;
  if (typeof n !== 'object' || Array.isArray(n)) return false;
  return !['text', 'image', 'split', 'grid', 'children', 'items'].some((k) => n[k] != null);
}

// 選んだ範囲が、上下の並びの中の「横いっぱいの行」になっているかを調べ、複製の段取りを返す。
// 複製した行の分だけ下のボックスを下にずらし、その分は、下のほうの空きのボックスを縮めて埋める。
// 空きが足りない（スライドからはみ出す）ときや、行全体を選んでいないときは null
function dupPlan() {
  const sel = S.sel || [];
  const slide = curSlide();
  if (!sel.length || !slide) return null;
  const boxes = sel.map(findBox).filter(Boolean);
  if (!boxes.length) return null;
  const selSet = new Set(sel);
  const full = (b) => slide.boxes.filter((x) => isLeaf(x) && within(x, b)).every((l) => selSet.has(l.path));
  const isRows = (b) => b && b.kind === 'split' && b.node?.split === 'rows' && b.src;
  // 行を含む上下の並び R を探す
  const a = boxes.length === 1 ? boxes[0] : commonAncestor(boxes);
  if (!a) return null;
  let R = isRows(a) && !full(a) ? a : null;
  let x = a;
  while (!R) {
    if (!full(x) || !x.parent) return null;
    const p = findBox(x.parent);
    if (isRows(p)) R = p;
    else x = p;
  }
  // 空きが足りなければ、外側の上下の並びでも探す
  for (let anc = R; isRows(anc); anc = anc.parent ? findBox(anc.parent) : null) {
    const r = splitEntries(anc, selSet);
    if (!r) return null;
    const part = r.entries.slice(r.i0, r.i1 + 1);
    const D = part.reduce((s2, e) => s2 + e.len, 0) + r.gap * part.length;
    const room = r.entries.slice(r.i1 + 1).filter((e) => isEmptyNode(e.node)).reduce((s2, e) => s2 + e.len, 0);
    if (D <= room + 1e-6) return { anc, r, D };
  }
  return null;
}

function duplicateRow() {
  const plan = dupPlan();
  if (!plan) return status('複製できません（行全体を選んでいないか、下の空きが足りません）。', true);
  const { anc, r, D } = plan;
  const { entries, i0, i1, gap, k, u } = r;
  // 複製は、元と同じ長さに固定する（中の名前は外す）
  const copies = entries.slice(i0, i1 + 1).map((e) => ({ node: stripIds(structuredClone(e.node ?? {})), ratio: `${round2(e.len)}${u}`, len: e.len }));
  const after = entries.slice(i1 + 1).map((e) => ({ ...e }));
  // 下のほうの空きから縮める
  let need = D;
  for (let j = after.length - 1; j >= 0 && need > 1e-6; j--) {
    const e = after[j];
    if (!isEmptyNode(e.node)) continue;
    if (e.len <= need + 1e-6) {
      need -= e.len + gap;
      e.drop = true;
      continue;
    }
    if (typeof e.ratio === 'string' || !k) e.ratio = `${round2(e.len - need)}${u}`;
    else e.ratio = round4(e.ratio - need / k);
    need = 0;
  }
  const list = [...entries.slice(0, i1 + 1), ...copies, ...after.filter((e) => !e.drop)];
  return writeEntries(anc, list).then(() => status('複製して追加しました'));
}

// 並びを分割 lca に書き戻す。1 つだけになったら、分割をやめてその 1 つを lca の位置に置く
function writeEntries(lca, list) {
  const { file, path, mode } = lca.src;
  const doc = S.files[file].doc;
  if (list.length === 1) {
    let child = list[0].node;
    if (typeof child === 'string') child = { text: child };
    child = structuredClone(child || {});
    const cur = getJS(doc, path) || {};
    // 外側の枠に余白が残っていて、残す中身に余白が無ければ、外側の余白を引き継ぐ（見た目を変えないため）
    if (cur && cur.padding != null && child.padding == null) child.padding = cur.padding;
    if (mode === 'direct') {
      const keep = Object.fromEntries(Object.entries(cur).filter(([x]) => FRAME_KEYS.includes(x)));
      replaceNode(doc, path, { ...child, ...keep });
    } else {
      // テンプレートの値を打ち消すため、使わない種類の中身は null にする
      for (const x of CONTENT_KEYS) setProp(doc, path, x, x in child ? child[x] : null);
      for (const [x, v] of Object.entries(child)) if (!CONTENT_KEYS.includes(x) && !FRAME_KEYS.includes(x)) setProp(doc, path, x, v);
    }
  } else {
    setProp(doc, path, 'children', list.map((e) => e.node));
    setProp(doc, path, 'ratio', list.map((e) => e.ratio));
  }
  setSel([]);
  return fileChanged(file);
}

// 分割 lca の中の選んだ範囲を消し、消した場所を、すぐ下（右）のボックスが伸びて埋める（いちばん下なら、すぐ上が伸びる）。
// ほかのボックスの大きさは変えない
function removeInSplit(lca, selSet) {
  const r = splitEntries(lca, selSet);
  if (!r) {
    const again = retryAfterTranspose(lca, () => {
      S.confirmed = true;
      try {
        return removeSelection();
      } finally {
        S.confirmed = false;
      }
    });
    return again || status('この範囲は削除できません。', true);
  }
  // 消す範囲が組（同じ向きの分割）のいちばん下・右で、その組の外側に、見た目ですぐ下・右のボックスがあるときは、
  // 外側の並びで扱い、そのボックスを伸ばす（例：タイトルの下に分けた空きを消すと、さらに下のボックスが上に伸びる）
  const { entries, i0, i1, gap, k, u } = r;
  const rest = entries.filter((_, i) => i < i0 || i > i1);
  if (!rest.length) return status('全部は削除できません。1 つ以上残してください', true);
  if (i1 + 1 >= entries.length) {
    // 外側へたどり、同じ向きの並びで、すぐ下（右）にボックスがあるところを探す。
    // 見つかったら、消した分だけ途中の枠を縮め、そのボックスを伸ばす（途中の枠に余白などが付いていてもよい）
    const dir = lca.node.split;
    const chain = [];
    let a = lca;
    let found = null;
    while (a.parent) {
      const q = findBox(a.parent);
      if (!q || q.kind !== 'split' || q.node.split !== dir || !q.src) break;
      const sibs = curSlide().boxes.filter((b) => b.parent === q.path);
      const ia = sibs.findIndex((b) => b.path === a.path);
      chain.push({ q, sibs, ia });
      if (ia + 1 < sibs.length) {
        found = true;
        break;
      }
      a = q;
    }
    if (found) {
      const removed = entries.slice(i0, i1 + 1);
      const freed = removed.reduce((x, e) => x + e.len, 0) + gap * removed.length;
      const files = new Set();
      chain.forEach(({ q, sibs, ia }, level) => {
        const L = (b) => (dir === 'rows' ? b.rect.h : b.rect.w);
        const n = sibs.length;
        const ratio = Array.isArray(q.node.ratio) && q.node.ratio.length === n ? plain(q.node.ratio) : Array(n).fill(1);
        let kq = null;
        sibs.forEach((b, i) => {
          if (kq == null && typeof ratio[i] === 'number' && ratio[i] > 0) kq = L(b) / ratio[i];
        });
        const change = (i, delta) => {
          if (typeof ratio[i] === 'string' || !kq) ratio[i] = `${round2(L(sibs[i]) + delta)}${u}`;
          else ratio[i] = round4(ratio[i] + delta / kq);
        };
        change(ia, -freed);
        if (level === chain.length - 1) change(ia + 1, freed);
        setProp(S.files[q.src.file].doc, q.src.path, 'ratio', ratio);
        files.add(q.src.file);
      });
      // 消した範囲の中の部分は、比率をそのまま残す（枠が縮んだ分だけ、ちょうど収まる）
      const list = rest.map((e) => ({ ...e }));
      for (const f of files) if (f !== lca.src.file) fileChanged(f);
      return writeEntries(lca, list);
    }
  }
  const j = i1 + 1 < entries.length ? i1 + 1 : i0 - 1;
  const nb = { ...entries[j] };
  const removed = entries.slice(i0, i1 + 1);
  const freed = removed.reduce((a, e) => a + e.len, 0) + gap * removed.length;
  if (typeof nb.ratio === 'number' && k) {
    // 比率 1 あたりの長さ（k）を変えずに、空いた長さの分だけ比率を足す
    const fixedLen = removed.filter((e) => typeof e.ratio === 'string').reduce((a, e) => a + e.len, 0);
    const num = removed.filter((e) => typeof e.ratio === 'number').reduce((a, e) => a + e.ratio, 0);
    nb.ratio = round4(nb.ratio + num + (fixedLen + gap * removed.length) / k);
  } else {
    nb.ratio = `${round2(nb.len + freed)}${u}`;
  }
  const list = entries.map((e, i) => (i === j ? nb : e)).filter((_, i) => i < i0 || i > i1);
  return writeEntries(lca, list);
}

function removeBox(box) {
  if (!canRemove(box)) return;
  if (hasContent(box) && !confirm('このボックスと中身を消しますか？')) return;
  const parent = findBox(box.parent);
  if (parent.kind === 'split') return removeInSplit(parent, new Set(curSlide().boxes.filter((b) => isLeaf(b) && within(b, box)).map((b) => b.path)));
  const sibs = curSlide().boxes.filter((b) => b.parent === parent.path);
  const i = sibs.findIndex((b) => b.path === box.path);
  return removeFromParent(parent, [i]);
}

// 選択中のボックスを消す（ドラッグで複数選んだときも）
function removeSelection() {
  const u = selectionUnit();
  if (!u) return;
  if ((S.sel || []).length < 2) {
    if (!canRemove(u.box)) return status('このボックスは削除できません（スライド全体のボックスは消せません）', true);
    return removeBox(u.box);
  }
  const lca = u.kind === 'range' ? u.lca : u.box;
  const boxes = (S.sel || []).map(findBox).filter(Boolean);
  if (boxes.some(hasContent) && !S.confirmed && !confirm('選んだボックスと中身を消しますか？')) return;
  const sel = new Set(S.sel);
  // 分割の中身を全部選んだときは、その分割ごと消す
  const all = curSlide().boxes.filter((b) => isLeaf(b) && within(b, lca)).every((l) => sel.has(l.path));
  if (all && canRemove(lca)) {
    const parent = findBox(lca.parent);
    if (parent.kind === 'split') return removeInSplit(parent, sel);
    const sibs = curSlide().boxes.filter((b) => b.parent === parent.path);
    return removeFromParent(parent, [sibs.findIndex((b) => b.path === lca.path)]);
  }
  if (lca.kind === 'split' && lca.src) return removeInSplit(lca, sel);
  return status('この範囲は削除できません。', true);
}

// ドラッグで選んだ範囲を 1 つのボックスにまとめる。選んでいないボックスは、そのまま残す。
// 内容は、選んだ中でいちばん上か左にある、内容のあるボックスのものを使う（ほかに内容があれば、確かめてから消す）
function mergeSelection() {
  const u = selectionUnit();
  if (!u || (S.sel || []).length < 2) return;
  const slide = curSlide();
  const order = (b) => slide.boxes.indexOf(b);
  const leaves = (S.sel || []).map(findBox).filter(Boolean).sort((a, b) => order(a) - order(b));
  const filled = leaves.filter((b) => b.kind === 'image' || b.kind === 'text');
  if (filled.length > 1 && !S.confirmed && !confirm('内容は、選んだ中でいちばん上か左にあるボックスのものだけを残し、ほかは消します。まとめますか？')) return;
  let node = filled.length ? writtenNode(filled[0]) : {};
  if (typeof node === 'string') node = { text: node };
  delete node.id;
  const lca = u.kind === 'range' ? u.lca : u.box;
  if (!lca.src) return status('このボックスは、まとめられません。', true);
  if (lca.kind === 'split') {
    const r = splitEntries(lca, new Set(S.sel));
    if (!r) {
      const again = retryAfterTranspose(lca, () => {
        S.confirmed = true;
        try {
          return mergeSelection();
        } finally {
          S.confirmed = false;
        }
      });
      return again || status('この範囲は、まとめられません。', true);
    }
    const { entries, i0, i1, gap, u: unit } = r;
    const part = entries.slice(i0, i1 + 1);
    const ratio = part.every((e) => typeof e.ratio === 'number')
      ? round4(part.reduce((a, e) => a + e.ratio, 0))
      : `${round2(part.reduce((a, e) => a + e.len, 0) + gap * (part.length - 1))}${unit}`;
    const list = [...entries.slice(0, i0), { node, ratio }, ...entries.slice(i1 + 1)];
    return writeEntries(lca, list).then(() => status('まとめました'));
  }
  // 格子：全部のセルを選んだときだけ、1 つにまとめる
  const kids = slide.boxes.filter((b) => b.parent === lca.path);
  if (!kids.every((kb) => within(kb, lca) && slide.boxes.filter((b) => isLeaf(b) && within(b, kb)).every((l) => S.sel.includes(l.path)))) {
    return status('格子の一部のセルは、まとめられません。', true);
  }
  return writeEntries(lca, [{ node }]).then(() => status('まとめました'));
}

// 分割・格子 parent の idx 番目の部分を消す。分割の中が 1 つだけ残ったら、分割をやめて外側の位置に置く
function removeFromParent(parent, idx) {
  const listKey = parent.kind === 'split' ? 'children' : 'items';
  const { file, path, mode } = parent.src;
  const doc = S.files[file].doc;
  const pn = parent.node || {};
  const sibs = curSlide().boxes.filter((b) => b.parent === parent.path);
  // 書いてあるとおりの中身を使う（use でテンプレートを呼んでいる枠を、展開した形で書き直さない）
  const written = mode === 'direct' ? getJS(doc, [...path, listKey]) : null;
  const all = Array.isArray(written) && written.length === sibs.length ? written : plain(pn[listKey] || []);
  const drop = new Set(idx);
  const list = all.filter((_, k) => !drop.has(k));
  const ratio0 = Array.isArray(pn.ratio) && pn.ratio.length === sibs.length ? pn.ratio : null;
  const ratio = ratio0 ? ratio0.filter((_, k) => !drop.has(k)) : null;
  if (!list.length) return status('全部は削除できません。1 つ以上残してください', true);

  if (parent.kind === 'split' && list.length === 1) {
    // 残った 1 つを外側の位置に戻す（外側の名前・背景・枠線は残す）
    let child = list[0];
    if (typeof child === 'string') child = { text: child };
    child = child || {};
    if (mode === 'direct') {
      const cur = getJS(doc, path) || {};
      const keep = Object.fromEntries(Object.entries(cur).filter(([k]) => FRAME_KEYS.includes(k)));
      replaceNode(doc, path, { ...child, ...keep });
    } else {
      // テンプレートの値を打ち消すため、使わない種類の中身は null にする
      for (const k of CONTENT_KEYS) setProp(doc, path, k, k in child ? child[k] : null);
      for (const [k, v] of Object.entries(child)) if (!CONTENT_KEYS.includes(k) && !FRAME_KEYS.includes(k)) setProp(doc, path, k, v);
    }
  } else {
    setProp(doc, path, listKey, list);
    if (ratio) {
      keepOneFlexible(ratio, 0);
      setProp(doc, path, 'ratio', ratio);
    }
  }
  setSel([]);
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

// 名前の列の幅を、画面に出ている名前のうち最も長いものに合わせる（入力欄の左端を全部揃える）
function fitLabels() {
  // 項目名の列の幅は、「スライド全体」と「選択中のボックス」の欄で別々に、いちばん長い項目名に合わせる
  for (const el of [$('#global'), $('#node')]) {
    if (!el) continue;
    el.style.setProperty('--label-w', 'max-content');
    const ws = [...el.querySelectorAll('.row > label')].map((l) => l.getBoundingClientRect().width);
    el.style.setProperty('--label-w', `${Math.ceil(Math.max(40, ...ws))}px`);
  }
}

function renderProps() {
  renderGlobal();
  queueMicrotask(fitLabels);
  if (S.keepPanel && S.keepPanel === S.selected) {
    S.keepPanel = null;
    return;
  }
  S.keepPanel = null;
  const el = $('#node');
  el.innerHTML = '';
  if (!S.layout) return;
  if ((S.sel || []).length > 1) return renderMulti(el);
  const box = S.selected ? findBox(S.selected) : null;
  el.append(h('h2', null, '選択中のボックス'));
  if (!box) {
    el.append(h('p', 'note', 'プレビューでボックスをクリックすると選択し、ここに設定が出ます。ドラッグすると、長方形になる範囲で複数のボックスを選択します。'));
    return;
  }
  const node = box.node || {};
  const inh = box.inherited || {};
  const u = S.layout.unit;
  // 中身の無いボックスは、初めから文字のボックスとして扱う
  const kind = box.kind === 'empty' ? 'text' : box.kind;
  el.append(h('p', 'meta', `${KIND_LABEL[kind]}　${fmt(box.rect.w)} × ${fmt(box.rect.h)} ${u}`));
  if (!box.src) return;
  const mode = box.src.mode;
  if (mode === 'shared') el.append(h('p', 'warn', 'このボックスは、複数のスライドで共通に使うボックスの組み合わせ（テンプレート）の一部です。変えると、このテンプレートを使う全てのスライドが変わります。'));

  // 種類
  const fsSize = sizeFields(box);
  if (fsSize) el.append(fsSize);
  if (isLeaf(box)) {
    el.append(row('内容', selectEl(KIND_OPTS, kind, (v) => {
      if (v === kind) return;
      // 画像にするとき、フォルダに画像が無ければファイルを選ぶ画面を出す
      if (v === 'image' && !S.assets.length) return pickImageFile().then((path) => (path ? changeKind(box, v, path) : renderProps()));
      changeKind(box, v);
    })));
  }

  // 中身
  if (box.kind === 'image') el.append(imageFields(box, node, inh));
  if (kind === 'text') el.append(textFields(box, node, inh));

  // 枠
  const fsBox = fieldset('余白・背景・枠線');
  fsBox.append(
    paddingRow(box, node),
    ...bgRows(node.background, (v) => edit(box, 'background', v)),
    row('枠線', h2(
      numEl(node.border?.width ?? 0, '', (v) => edit(box, 'border', +v > 0 ? { width: +v, color: node.border?.color || '#333333' } : undefined)),
      'pt',
      colorEl(node.border?.color, (v) => node.border && edit(box, ['border', 'color'], v || '#333333')),
    )),
  );
  el.append(fsBox);
}

// 内側の余白。上・右・下・左の 4 つを 1 行で
// 余白の 4 つの値 a（上・右・下・左）を、ボックス box の padding に書く値にする。
// 決まった余白と同じなら書かない。テキストのボックスで 0 にしたときは、0 と書いて残す
function padValue(box, a) {
  const def = defaultPadding(box.kind);
  const u = S.layout.unit;
  const defV = def ? def.map((x) => Math.round(toUnit(x, u) * 1000) / 1000) : null;
  const out = a.every((y) => y === a[0]) ? a[0] : a;
  if (defV ? a.every((y, k) => y === defV[k]) : out === 0) return undefined;
  return out;
}

// 余白の欄。targets（省略すると box だけ）の 1 つずつに、同じ余白を書く。値は box のものを出す
function paddingRow(box, node, targets) {
  const list = targets || [box];
  // テキストのボックスで余白を書いていないときは、決まった余白（上下左右に 0.5 cm）を出す
  const def = defaultPadding(box.kind);
  const u = S.layout.unit;
  const defV = def ? def.map((x) => Math.round(toUnit(x, u) * 1000) / 1000) : null;
  const p = node.padding ?? defV;
  let v;
  if (Array.isArray(p)) v = p.length === 2 ? [p[0], p[1], p[0], p[1]] : [p[0], p[1], p[2], p[3]];
  else v = [p ?? 0, p ?? 0, p ?? 0, p ?? 0];
  const writeAll = (key, valueOf) => {
    const files = new Set();
    try {
      for (const t of list) {
        if (!t.src) continue;
        setProp(S.files[t.src.file].doc, t.src.path, key, valueOf(t));
        files.add(t.src.file);
      }
    } catch (e) {
      return status(`書き換えられません: ${e.message}`, true);
    }
    return Promise.all([...files].map((f) => fileChanged(f)));
  };
  const write = () => {
    const a = v.map((y) => +y || 0);
    writeAll('padding', (t) => padValue(t, a));
  };
  // ふだんは、上下と左右を 1 つずつの欄で同時に変える
  const pair = h('span', 'in pad4');
  [['↕', '上下', [0, 2]], ['↔', '左右', [1, 3]]].forEach(([icon, title, ks]) => {
    const ic = h('span', 'pad-ic', icon);
    ic.title = title;
    const same = v[ks[0]] === v[ks[1]];
    const inp = numEl(same ? v[ks[0]] ?? 0 : '', same ? '0' : '別々', (x) => {
      for (const k of ks) v[k] = x === '' ? 0 : +x;
      write();
    });
    pair.append(ic, h('span', 'pad-name', title), inp);
  });
  // 上・右・下・左を別々に変える欄は、折りたたんでおく（上下・左右がそろっていないときは開く）
  const four = h('span', 'in pad4');
  [['↑', '上'], ['→', '右'], ['↓', '下'], ['←', '左']].forEach(([icon, title], k) => {
    const ic = h('span', 'pad-ic', icon);
    ic.title = title;
    const inp = numEl(v[k] ?? 0, '0', (x) => {
      v[k] = x === '' ? 0 : +x;
      write();
    });
    four.append(ic, inp);
  });
  const det = document.createElement('details');
  det.className = 'pad-detail';
  det.open = S.padOpen || v[0] !== v[2] || v[1] !== v[3];
  det.ontoggle = () => (S.padOpen = det.open);
  const sum = document.createElement('summary');
  sum.textContent = '上・右・下・左を別々に設定する';
  det.append(sum, four);
  const r = row('余白', pair);
  const rd = row('', det);
  const inset = checkEl('背景ごと余白を設ける', !!node.background_inset, (x) => writeAll('background_inset', () => x || undefined));
  const r2 = row('', inset);
  const both = h('div');
  both.append(r, rd, r2);
  return both;
}

// 複数のボックスを選んだとき。共通の外側の分割について、選んだ部分どうしの比率と、まとめた背景を設定する
function renderMulti(el) {
  const boxes = S.sel.map(findBox).filter(Boolean);
  const r = selectionRect();
  const u = S.layout.unit;
  el.append(h('h2', null, '選択中のボックス'), h('p', 'meta', `${boxes.length} 個　${fmt(r.w)} × ${fmt(r.h)} ${u}`));
  const lca = commonAncestor(boxes);
  if (!lca || !lca.src) return;
  const kids = curSlide().boxes.filter((b) => b.parent === lca.path);
  const covered = kids.map((k, i) => [k, i]).filter(([k]) => overlaps(k.rect, r));
  const all = covered.length === kids.length;
  if (lca.src.mode === 'shared') el.append(h('p', 'warn', 'このボックスは、複数のスライドで共通に使うボックスの組み合わせ（テンプレート）の一部です。変えると、このテンプレートを使う全てのスライドが変わります。'));
  el.append(multiSizeFields(lca, r, all));

  if (lca.kind === 'split' && covered.length >= 2) {
    const n = kids.length;
    const ratio = Array.isArray(lca.node.ratio) && lca.node.ratio.length === n ? [...lca.node.ratio] : Array(n).fill(1);
    const parts = [];
    covered.forEach(([, i], j) => {
      if (j) parts.push(h('span', 'ratio-sep', ':'));
      const t = textEl(String(ratio[i]), '1', (x) => {
        const pv = parseLen(x);
        if (pv === undefined) return;
        ratio[i] = pv;
        keepOneFlexible(ratio, i);
        edit(lca, 'ratio', ratio);
      });
      t.className = 'ratio-in';
      parts.push(t);
    });
    const fs = fieldset(lca.node.split === 'rows' ? '上下の比率' : '左右の比率');
    fs.append(h2(...parts), h('div', 'note', `${lca.node.split === 'rows' ? '上から' : '左から'}順に並ぶ。「3cm」のように単位を付けると、その長さに固定される`));
    el.append(fs);
  }

  // テキストのボックスを選んでいれば、「テキストの下に生じた余白を別のボックスにする」を 1 つずつに設定する
  const texts = boxes.filter((b) => b.kind === 'text' && b.src);
  if (texts.length) {
    const fsT = fieldset('テキスト（選んだボックスの 1 つずつ）');
    const on = texts.every((b) => (b.node || {}).split_below !== false);
    fsT.append(row('', checkEl('テキストの下に生じた余白を別のボックスにする', on, (x) => setSplitBelowMany(texts.map((b) => b.path), x))));
    el.append(fsT);
  }
  // 余白は、選んだボックスの 1 つずつに同じ値を設定する（値は、いちばん上か左のボックスのものを出す）
  const slideBoxes = curSlide().boxes;
  const leaves = boxes.filter((b) => b.src).sort((a, b) => slideBoxes.indexOf(a) - slideBoxes.indexOf(b));
  if (leaves.length) {
    const fsPad = fieldset('余白（選んだボックスの 1 つずつ）');
    fsPad.append(paddingRow(leaves[0], leaves[0].node || {}, leaves));
    el.append(fsPad);
  }
  const outerPad = all && lca.node?.padding != null && JSON.stringify(lca.node.padding) !== '0';
  const fsBg = fieldset(outerPad ? '外側の枠の余白・背景（選んだ範囲全体）' : '背景（選んだ範囲全体）');
  if (all) {
    // 外側の枠に余白が残っているときだけ、その余白も出す（クリックでは選べない枠のため）
    if (outerPad) fsBg.append(paddingRow(lca, lca.node || {}));
    fsBg.append(...bgRows(lca.node.background, (v) => edit(lca, 'background', v)));
  } else if (lca.kind === 'split') {
    fsBg.append(...bgRows(undefined, (v) => v && groupWithBackground(lca, kids, covered[0][1], covered[covered.length - 1][1], v)));
  } else {
    fsBg.append(h('p', 'note', '格子の一部のセルには、まとめた背景を付けられません。'));
  }
  el.append(fsBg);
}

// 複数選んだときの、選んだ範囲全体の幅と高さ。
// 外側の分割の並ぶ向きで一部を選んだときは、選んだ部分を、合計がその長さになるように長さで固定する（中の比は保つ）。
// それ以外の向き（と、分割の全部を選んだとき）は、外側の枠そのものの大きさを変える
function multiSizeFields(lca, r, all) {
  const u = S.layout.unit;
  const r3 = (v) => Math.round(v * 1000) / 1000;
  const fs = fieldset('ボックスの大きさ（選んだ範囲全体）');
  const dimInput = (dir) => {
    const cur = dir === 'rows' ? r.h : r.w;
    if (lca.kind === 'split' && lca.node.split === dir && !all) {
      return numEl(r3(cur), '', (v) => +v > 0 && setRangeLength(lca, +v));
    }
    // 格子の一部のセルを選んだときは、値を出すだけ
    const c = lca.kind === 'grid' && !all ? null : sizeCtrl(lca, dir);
    if (!c) {
      // 外側で分けられていない向きは、値を出すだけ
      const i = numEl(r3(cur), '', () => {});
      i.disabled = true;
      return i;
    }
    const n = c.kids.length;
    const ratio = Array.isArray(c.p.node.ratio) && c.p.node.ratio.length === n ? [...c.p.node.ratio] : Array(n).fill(1);
    return numEl(r3(cur), '', (v) => {
      if (!(+v > 0)) return;
      ratio[c.i] = `${+v}${u}`;
      keepOneFlexible(ratio, c.i);
      edit(c.p, 'ratio', ratio);
    });
  };
  const line = h('span', 'in row-line');
  line.append(dimInput('cols'), u, h('label', null, '高さ'), dimInput('rows'), u);
  fs.append(row('幅', line));
  return fs;
}

// 分割 lca の中の選んだ部分を、合計の長さが total になるように長さで固定する
function setRangeLength(lca, total) {
  const r = splitEntries(lca, new Set(S.sel));
  if (!r) return retryAfterTranspose(lca, () => status('並びを組み替えました。もう一度、長さを入れてください。')) || status('この範囲の大きさは変えられません。', true);
  const { entries, i0, i1, gap, u } = r;
  const part = entries.slice(i0, i1 + 1);
  const sum = part.reduce((a, e) => a + e.len, 0);
  const inner = total - gap * (part.length - 1);
  if (!(inner > 0) || !(sum > 0)) return status('この長さにはできません。', true);
  const list = entries.map((e, i) => (i >= i0 && i <= i1 ? { ...e, ratio: `${round2((e.len * inner) / sum)}${u}` } : { ...e }));
  // 全部が長さで固定になると、空いた長さを受け取る部分が無くなるので、選んでいない隣を比率に戻す
  if (list.every((e) => typeof e.ratio === 'string')) {
    const j = i1 + 1 < list.length ? i1 + 1 : i0 - 1;
    if (j >= 0) list[j].ratio = 1;
  }
  if (entries.length === curSlide().boxes.filter((b) => b.parent === lca.path).length) {
    // 並びの組み替えが無いときは比率だけを書き換え、選択はそのまま
    return edit(lca, 'ratio', list.map((e) => e.ratio));
  }
  return writeEntries(lca, list);
}

function commonAncestor(boxes) {
  const chain = (b) => {
    const c = [];
    for (let x = b; x; x = x.parent ? findBox(x.parent) : null) c.unshift(x);
    return c;
  };
  const chains = boxes.map(chain);
  let lca = null;
  for (let d = 0; ; d++) {
    const c = chains[0][d];
    if (!c || !chains.every((ch) => ch[d] && ch[d].path === c.path)) break;
    lca = c;
  }
  return lca;
}

// 分割の一部の部分に背景を付ける。その部分をまとめて 1 つの分割にし、まとめた分割に背景を付ける
// 選んだ範囲を、置き換えの単位（1 つのボックスか、ある分割の連続した部分）として捉える
//   { kind: 'node', box }           … その 1 つのボックス（または分割全体）
//   { kind: 'range', lca, i0, i1 }  … 分割 lca の i0〜i1 番目の部分
function selectionUnit() {
  const boxes = (S.sel || []).map(findBox).filter(Boolean);
  if (!boxes.length) return null;
  if (boxes.length === 1) return { kind: 'node', box: boxes[0] };
  const lca = commonAncestor(boxes);
  if (!lca) return null;
  const r = selectionRect();
  const kids = curSlide().boxes.filter((b) => b.parent === lca.path);
  const covered = kids.map((k, i) => [k, i]).filter(([k]) => overlaps(k.rect, r));
  if (covered.length === kids.length || lca.kind !== 'split') return { kind: 'node', box: lca };
  return { kind: 'range', lca, kids, i0: covered[0][1], i1: covered[covered.length - 1][1] };
}

// YAML に書いてあるとおりの中身（use でテンプレートを呼ぶ枠は展開しない）
function writtenNode(box) {
  if (box.src?.mode === 'direct') {
    const v = getJS(S.files[box.src.file].doc, box.src.path);
    if (v !== undefined) return structuredClone(v ?? {});
  }
  return plain(box.node || {});
}

// 貼り付けた先で id が重ならないように、中の id を外す
function stripIds(n) {
  if (Array.isArray(n)) return n.map(stripIds);
  if (!n || typeof n !== 'object') return n;
  const out = {};
  for (const [k, v] of Object.entries(n)) if (k !== 'id') out[k] = stripIds(v);
  return out;
}

function copySelection() {
  const u = selectionUnit();
  if (!u) return;
  const r = selectionRect();
  let node;
  if (u.kind === 'node') {
    node = writtenNode(u.box);
    if (typeof node === 'string') node = { text: node };
  } else {
    const { lca, kids, i0, i1 } = u;
    const n = kids.length;
    const ratio = Array.isArray(lca.node.ratio) && lca.node.ratio.length === n ? lca.node.ratio : Array(n).fill(1);
    node = {
      split: lca.node.split,
      ...(lca.node.gap != null ? { gap: lca.node.gap } : {}),
      ratio: ratio.slice(i0, i1 + 1),
      children: kids.slice(i0, i1 + 1).map((k) => writtenNode(k)),
    };
  }
  S.clip = { node: stripIds(node), w: r.w, h: r.h, count: (S.sel || []).length };
  status(`コピーしました（${fmt(r.w)} × ${fmt(r.h)} ${S.layout.unit}）`);
}

function pasteSelection() {
  const u = selectionUnit();
  if (!u || !S.clip) return;
  const r = selectionRect();
  const same = (a, b) => Math.abs(a - b) <= Math.max(a, b) * 0.005;
  const unit = S.layout.unit;
  if (!same(r.w, S.clip.w) || !same(r.h, S.clip.h)) {
    return status(`大きさが違うので貼り付けられません（コピーしたもの ${fmt(S.clip.w)} × ${fmt(S.clip.h)} ${unit}、選択中 ${fmt(r.w)} × ${fmt(r.h)} ${unit}）`, true);
  }
  const node = structuredClone(S.clip.node);
  if (u.kind === 'node') {
    const box = u.box;
    if (!box.src || box.src.mode !== 'direct') return status('テンプレートの一部には貼り付けられません。', true);
    const doc = S.files[box.src.file].doc;
    const cur = getJS(doc, box.src.path);
    // 貼り付け先の id は残す（CLI やテンプレートの fill が使う名前のため）
    if (cur && typeof cur === 'object' && cur.id) node.id = cur.id;
    replaceNode(doc, box.src.path, node);
    setSel([]);
    return fileChanged(box.src.file).then(() => status('貼り付けました'));
  }
  const done = replaceRange(u.lca, u.kids, u.i0, u.i1, node);
  if (done && done.then) done.then(() => status('貼り付けました'));
}

// 分割 lca の i0〜i1 番目を、1 つのノードに置き換える（外側での大きさは、置き換える部分の合計）
function replaceRange(lca, kids, i0, i1, node) {
  if (!lca.src) return status('この範囲には貼り付けられません。', true);
  const doc = S.files[lca.src.file].doc;
  const path = lca.src.path;
  const n = kids.length;
  // テンプレートから来た枠では、展開した中身を書き戻す
  const children = lca.src.mode === 'direct' ? getJS(doc, [...path, 'children']) : plain(lca.node.children || []);
  if (!Array.isArray(children) || children.length !== n) return status('この範囲には貼り付けられません。', true);
  const ratio = Array.isArray(lca.node.ratio) && lca.node.ratio.length === n ? [...lca.node.ratio] : Array(n).fill(1);
  const subR = ratio.slice(i0, i1 + 1);
  const u = S.layout.unit;
  const gap = toUnit(lca.node.gap ?? 0, u);
  let outer;
  if (subR.every((x) => typeof x === 'number')) outer = subR.reduce((a, x) => a + x, 0);
  else if (subR.every((x) => typeof x === 'string')) outer = `${Math.round((subR.reduce((a, x) => a + toUnit(x, u), 0) + gap * (subR.length - 1)) * 1000) / 1000}${u}`;
  else {
    // 比率と長さが混ざっているときは、いま画面に出ている長さで置き換える
    const a = kids[i0].rect, b = kids[i1].rect;
    const len = lca.node.split === 'cols' ? b.x + b.w - a.x : b.y + b.h - a.y;
    outer = `${Math.round(len * 1000) / 1000}${u}`;
  }
  children.splice(i0, subR.length, node);
  ratio.splice(i0, subR.length, outer);
  setProp(doc, path, 'children', children);
  setProp(doc, path, 'ratio', ratio);
  setSel([]);
  return fileChanged(lca.src.file);
}

function groupWithBackground(lca, kids, i0, i1, bg) {
  if (lca.src.mode !== 'direct') return status('テンプレートの一部なので、この範囲にまとめた背景を付けられません。', true);
  const doc = S.files[lca.src.file].doc;
  const path = lca.src.path;
  const n = kids.length;
  const children = getJS(doc, [...path, 'children']);
  if (!Array.isArray(children) || children.length !== n) return status('この範囲にはまとめた背景を付けられません。', true);
  const ratio = Array.isArray(lca.node.ratio) && lca.node.ratio.length === n ? [...lca.node.ratio] : Array(n).fill(1);
  const sub = children.slice(i0, i1 + 1);
  const subR = ratio.slice(i0, i1 + 1);
  const u = S.layout.unit;
  const gap = toUnit(lca.node.gap ?? 0, u);
  let outer;
  if (subR.every((x) => typeof x === 'number')) outer = subR.reduce((a, x) => a + x, 0);
  else if (subR.every((x) => typeof x === 'string')) outer = `${Math.round((subR.reduce((a, x) => a + toUnit(x, u), 0) + gap * (sub.length - 1)) * 1000) / 1000}${u}`;
  else return status('比率と長さが混ざった範囲には、まとめた背景を付けられません。', true);
  const group = { split: lca.node.split, ...(lca.node.gap != null ? { gap: lca.node.gap } : {}), ratio: subR, children: sub, background: bg };
  children.splice(i0, sub.length, group);
  ratio.splice(i0, sub.length, outer);
  setProp(doc, path, 'children', children);
  setProp(doc, path, 'ratio', ratio);
  setSel([]);
  // まとめた後も同じ範囲を選んだままにする（続けて色や画像を変えられるように）
  const groupPath = `${lca.path}.${i0}`;
  return fileChanged(lca.src.file).then(() => {
    const inside = (b) => b && (b.path === groupPath || (b.parent && inside(findBox(b.parent))));
    const leaves = curSlide().boxes.filter((b) => isLeaf(b) && inside(b));
    if (leaves.length > 1) {
      setSel(leaves.map((b) => b.path));
      applyView();
      renderProps();
    }
  });
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
  const selKey = (S.sel || []).join('|');
  if (S.globalFor !== selKey) {
    S.globalOpen = !selKey; // 1 つでも複数でも、ボックスを選んでいるあいだは折りたたむ
    S.globalFor = selKey;
  }
  const fold = document.createElement('details');
  fold.open = S.globalOpen;
  fold.ontoggle = () => {
    S.globalOpen = fold.open;
    fitLabels();
  };
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
    row('フォント（英語）', selectEl(FONT_OPTS, getJS(doc, ['defaults', 'font', 'family']) ?? DEFAULT_FONT, (v) => {
      setProp(doc, ['defaults'], ['font', 'family'], v);
      fileChanged('layout.yaml');
    })),
    row('フォント（日本語）', selectEl(JA_FONT_OPTS, getJS(doc, ['defaults', 'font', 'family_ja']) ?? DEFAULT_FONT_JA, (v) => {
      setProp(doc, ['defaults'], ['font', 'family_ja'], v);
      fileChanged('layout.yaml');
    })),
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

// そのボックスの幅（左右に分ける分割）・高さ（上下に分ける分割）を決めている分割と、その中の何番目か
// 同じ並びの部分がすべて長さで固定になると、空いた長さを受け取る部分が無くなり、隙間が残る。
// そうなるときは、書き換えた部分の隣（最後の部分なら 1 つ前）を比率に戻し、残りの長さを受け取らせる
function keepOneFlexible(ratio, i) {
  if (!ratio.every((x) => typeof x === 'string')) return;
  const j = i + 1 < ratio.length ? i + 1 : i - 1;
  if (j >= 0) ratio[j] = 1;
}

function sizeCtrl(box, dir) {
  for (let b = box; b && b.parent; b = findBox(b.parent)) {
    const p = findBox(b.parent);
    if (p && p.kind === 'split' && p.node?.split === dir && p.src) {
      const kids = curSlide().boxes.filter((x) => x.parent === p.path);
      const i = kids.findIndex((x) => x.path === b.path);
      if (i >= 0 && kids.length === (p.node.children || []).length) return { p, kids, i, child: b };
    }
  }
  return null;
}

// ボックスの大きさ。長さで指定すると、その長さに固定する（同じ並びのほかの部分は、残りを比率で分ける）
function sizeFields(box) {
  const u = S.layout.unit;
  const mode = S.sizeMode || 'len';
  const r3 = (v) => Math.round(v * 1000) / 1000;
  const fs = fieldset('ボックスの大きさ');
  fs.append(row('サイズの指定', selectEl([['len', '長さで指定'], ['ratio', '比率で指定']], mode, (v) => {
    S.sizeMode = v;
    renderProps();
  })));
  let shared = false;
  const dimInput = (dir) => {
    const lenOf = (b) => (dir === 'rows' ? b.rect.h : b.rect.w);
    const c = sizeCtrl(box, dir);
    if (!c) {
      // 外側で分けられていない向きは、値を出すだけ
      const i = numEl(r3(lenOf(box)), '', () => {});
      i.disabled = true;
      return i;
    }
    if (c.p.src.mode === 'shared') shared = true;
    const n = c.kids.length;
    const ratio = Array.isArray(c.p.node.ratio) && c.p.node.ratio.length === n ? [...c.p.node.ratio] : Array(n).fill(1);
    const write = (v) => {
      ratio[c.i] = v;
      keepOneFlexible(ratio, c.i);
      edit(c.p, 'ratio', ratio);
    };
    if (mode === 'len') return numEl(r3(lenOf(c.child)), '', (v) => +v > 0 && write(`${+v}${u}`));
    let val = ratio[c.i];
    if (typeof val === 'string') {
      // 長さで決めている部分を比率で表すときは、比率で決めているほかの部分を基準にする
      const j = c.kids.findIndex((k, ix) => ix !== c.i && typeof ratio[ix] === 'number' && ratio[ix] > 0);
      val = j >= 0 ? r3((lenOf(c.child) / lenOf(c.kids[j])) * ratio[j]) : 1;
    }
    return numEl(val, '', (v) => +v > 0 && write(+v));
  };
  const unitTxt = mode === 'len' ? u : '';
  const line = h('span', 'in row-line');
  line.append(dimInput('cols'), unitTxt, h('label', null, '高さ'), dimInput('rows'), unitTxt);
  fs.append(row('幅', line));
  if (shared) fs.append(h('p', 'warn', 'この大きさは、同じテンプレートを使う全てのスライドで共通です。'));
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
  const fs = fieldset(null);
  const opts = [...new Set([...S.assets, String(node.image)])].map((a) => [a, a.split('/').pop()]);
  const srcIn = h2(
    selectEl(opts, String(node.image), (v) => edit(box, 'image', v)),
    pickBtn((path) => edit(box, 'image', path)),
  );
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
    row('画像', srcIn),
    row('収め方', selectEl(FIT_OPTS, node.fit ?? inh.fit ?? 'contain', (v) => edit(box, 'fit', v))),
    row('横の位置', selectEl(H_OPTS, node.align ?? inh.align ?? 'center', (v) => edit(box, 'align', v))),
    row('縦の位置', selectEl(V_OPTS, node.valign ?? inh.valign ?? 'middle', (v) => edit(box, 'valign', v))),
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

function wrapWhole(text, f) {
  let t = text;
  if (f.italic) t = `<i>${t}</i>`;
  if (f.bold) t = `<b>${t}</b>`;
  return t;
}

// 揃え方（横）と（縦）を 1 行に
function alignRow(box, node, inh, hDflt, vDflt) {
  // ほかの行と同じく、名前の列と入力の列に分ける（入力の左端を揃える）
  const vals = h('span', 'in row-line');
  vals.append(
    selectEl(H_OPTS, node.align ?? inh.align ?? hDflt, (v) => edit(box, 'align', v)),
    h('label', null, '揃え方（縦）'),
    selectEl(V_OPTS, node.valign ?? inh.valign ?? vDflt, (v) => edit(box, 'valign', v)),
  );
  return row('揃え方（横）', vals);
}

function textFields(box, node, inh) {
  const fs = fieldset(null);
  // 文字の高さが変わりうる書き換えのあとで、下の余白を合わせ直す
  const edit = (b, k, v) => {
    S.pendingFit = [box.path];
    return editOne(b, k, v);
  };
  const f = node.font || {};
  const fi = inh.font || {};
  fs.append(
    row('テキスト', richEditor(wrapWhole(String(node.text ?? ''), f), (v) => {
      S.keepPanel = box.path;
      const off = box.src.mode === 'direct' ? undefined : false;
      const pairs = [['text', v]];
      if (f.bold) pairs.push([['font', 'bold'], off]);
      if (f.italic) pairs.push([['font', 'italic'], off]);
      S.pendingFit = [box.path];
      editMany(box, pairs);
    })),
    row('', checkEl('テキストの下に生じた余白を別のボックスにする', node.split_below !== false, (x) => {
      if (x) {
        S.pendingFit = [box.path];
        return editOne(box, 'split_below', undefined);
      }
      return unfitTextBelow(box);
    })),
    row('大きさ', h2(numEl(f.size ?? fi.size ?? DEFAULT_SIZE, '', (v) => +v > 0 && edit(box, ['font', 'size'], +v)), 'pt')),
    row('フォント（英語）', selectEl(FONT_OPTS, f.family ?? fi.family ?? DEFAULT_FONT, (v) => edit(box, ['font', 'family'], v))),
    row('フォント（日本語）', selectEl(JA_FONT_OPTS, f.family_ja ?? fi.family_ja ?? DEFAULT_FONT_JA, (v) => edit(box, ['font', 'family_ja'], v))),
    row('色', colorEl(f.color ?? fi.color, (v) => edit(box, ['font', 'color'], v || undefined))),
    row('見出しアイコン', h2(
      selectEl([['', 'なし'], ...[...new Set([...S.assets, ...(node.icon ? [String(node.icon)] : [])])].map((a) => [a, a.split('/').pop()])], node.icon ?? '', (v) => edit(box, 'icon', v || undefined)),
      pickBtn((path) => edit(box, 'icon', path)),
    )),
    ...(node.icon ? [row('アイコンの大きさ', h2(numEl(node.icon_size ?? f.size ?? fi.size ?? DEFAULT_SIZE, '', (v) => +v > 0 && edit(box, 'icon_size', +v)), 'pt'))] : []),
    alignRow(box, node, inh, 'left', 'top'),
    row('行間', h2(numEl(node.line_height ?? inh.line_height ?? 1.2, '', (v) => +v > 0 && edit(box, 'line_height', +v), 0.05), '倍')),
    row('段落の間隔', h2(numEl(node.para_space ?? inh.para_space ?? 0, '', (v) => edit(box, 'para_space', +v > 0 ? +v : undefined)), 'pt')),
  );
  return fs;
}


// ---------------------------------------------------------------- 部品

// 「ファイルから選ぶ…」：Finder で画像を選び、layout.yaml と同じフォルダに写してから使う
function pickBtn(onPicked) {
  const b = btn('ファイルから選ぶ…', async () => {
    const path = await pickImageFile();
    if (!path) return;
    await onPicked(path);
    if (S.pickMsg) status(S.pickMsg); // 描き直しで消えないよう、書き換えの後に出す
  });
  b.className = 'pick-btn';
  return b;
}

async function pickImageFile() {
  if (!S.source?.writable) {
    status('画像は、フォルダを開いているときだけ足せます', true);
    return null;
  }
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/png,image/jpeg,image/gif,image/svg+xml,image/webp';
  const file = await new Promise((res) => {
    input.onchange = () => res(input.files[0] || null);
    input.addEventListener('cancel', () => res(null));
    input.click();
  });
  if (!file) return null;
  const name = file.name.replace(/[\\/:*?"<>|]/g, '_');
  const m = name.match(/^(.*?)(\.[^.]+)?$/);
  let path = name;
  for (let k = 1; await S.source.exists(path); k++) {
    // 同じ名前で中身も同じなら、写さずにそのまま使う
    const old = await S.source.readBlob(path);
    if (old.size === file.size && (await old.arrayBuffer().then((a) => new Uint8Array(a).join())) === (await file.arrayBuffer().then((a) => new Uint8Array(a).join()))) {
      S.pickMsg = `${path} を使います`;
      return path;
    }
    path = `${m[1]}_${k}${m[2] || ''}`;
  }
  try {
    await S.source.writeFile(path, file);
  } catch (e) {
    status(`画像をフォルダに入れられません: ${e.message}`, true);
    return null;
  }
  if (!S.assets.includes(path)) S.assets.push(path);
  S.pickMsg = `画像 ${path} をフォルダに入れました`;
  return path;
}

// テキストの欄。ドラッグで選んだ範囲を、ボタンで太字・斜体・上付き・下付きにする。
// Enter は段落の区切り（行末に ¶）、Shift+Enter は段落の中の改行（↵）。
// layout.yaml には、段落の区切りを改行、段落の中の改行を <br> として書く
const SOFT_BREAK = '<span class="sb" contenteditable="false">↵</span><br>';
function editorHTML(value) {
  return paragraphs(value)
    .map((p) => `<div>${richToHTML(p, SOFT_BREAK) || '<br>'}</div>`)
    .join('');
}

function richEditor(value, onCommit) {
  const wrap = h('div', 'rich-wrap');
  const bar = h('div', 'rich-bar');
  const ed = document.createElement('div');
  ed.className = 'rich';
  ed.contentEditable = 'true';
  ed.innerHTML = editorHTML(value);
  let last = value;
  let timer = null;
  const commit = () => {
    clearTimeout(timer);
    const v = serializeRich(ed);
    if (v === last) return;
    last = v;
    onCommit(v);
  };
  for (const [cmd, label, title] of [['bold', 'B', '太字'], ['italic', 'I', '斜体'], ['superscript', 'x<sup>2</sup>', '上付き'], ['subscript', 'x<sub>2</sub>', '下付き']]) {
    const b = h('button', `rich-btn rich-${cmd}`);
    b.innerHTML = label;
    b.type = 'button';
    b.title = title;
    b.onmousedown = (e) => {
      e.preventDefault(); // テキストの欄の選択を保つ
      document.execCommand('styleWithCSS', false, false);
      document.execCommand(cmd);
      commit();
    };
    bar.append(b);
  }
  ed.onfocus = () => document.execCommand('defaultParagraphSeparator', false, 'div');
  // Shift+Enter：カーソルのある段落の中に ↵ と改行を差し込み、カーソルをその直後に置く
  const insertSoftBreak = () => {
    const sel = getSelection();
    if (!sel.rangeCount) return;
    let r = sel.getRangeAt(0);
    r.deleteContents();
    if (r.startContainer === ed) {
      // カーソルが段落の外にあるときは、直前の段落の末尾に移す
      const div = ed.children[Math.max(0, r.startOffset - 1)] || ed.firstElementChild;
      if (!div) return;
      r = document.createRange();
      r.selectNodeContents(div);
      r.collapse(false);
      const last = div.lastChild;
      if (last && last.nodeName === 'BR' && !afterMark(last)) r.setStartBefore(last);
    }
    const sb = document.createElement('span');
    sb.className = 'sb';
    sb.contentEditable = 'false';
    sb.textContent = '↵';
    const br = document.createElement('br');
    const z = document.createTextNode('\u200b'); // 行末の改行を見えるようにするゼロ幅の文字（書き出すときに外す）
    r.insertNode(z);
    r.insertNode(br);
    r.insertNode(sb);
    const nr = document.createRange();
    nr.setStart(z, 1);
    nr.collapse(true);
    sel.removeAllRanges();
    sel.addRange(nr);
    ed.dispatchEvent(new Event('input'));
  };
  ed.onkeydown = (e) => {
    if (e.key === 'Enter' && e.shiftKey) {
      e.preventDefault();
      insertSoftBreak();
    }
  };
  ed.oninput = () => {
    clearTimeout(timer);
    timer = setTimeout(commit, 600);
  };
  ed.onblur = commit;
  ed.onpaste = (e) => {
    e.preventDefault();
    document.execCommand('insertText', false, e.clipboardData.getData('text/plain'));
  };
  wrap.append(bar, ed, h('div', 'rich-legend', '¶ 段落の区切り（Enter）　↵ 段落の中の改行（Shift+Enter）'));
  return wrap;
}

// その <br> が ↵ の直後にあるか（間の空の文字の節は読み飛ばす）
function afterMark(br) {
  let p = br.previousSibling;
  while (p && p.nodeType === 3 && !p.nodeValue.replace(/\u200b/g, '')) p = p.previousSibling;
  return !!p?.classList?.contains('sb');
}

function serializeRich(root) {
  const TAGS = { B: 'b', STRONG: 'b', I: 'i', EM: 'i', SUP: 'sup', SUB: 'sub' };
  const inline = (n) => {
    let out = '';
    for (const c of n.childNodes) {
      if (c.nodeType === 3) out += c.nodeValue.replace(/ /g, ' ').replace(/​/g, '');
      else if (c.nodeType === 1) {
        if (c.classList?.contains('sb')) continue;
        if (c.nodeName === 'BR') {
          // ↵ の直後の <br> だけが段落の中の改行。ほかの <br> は空の行を保つための置き物
          if (afterMark(c)) out += '<br>';
          continue;
        }
        const t = TAGS[c.nodeName];
        if (t) out += `<${t}>`;
        out += inline(c);
        if (t) out += `</${t}>`;
      }
    }
    return out;
  };
  const paras = [];
  let pending = null;
  for (const c of root.childNodes) {
    if (c.nodeType === 1 && (c.nodeName === 'DIV' || c.nodeName === 'P')) {
      if (pending !== null) paras.push(pending);
      pending = null;
      paras.push(inline(c));
    } else {
      pending = (pending ?? '') + inline({ childNodes: [c] });
    }
  }
  if (pending !== null) paras.push(pending);
  return paras.join('\n').replace(/<(b|i|sup|sub)><\/\1>/g, '');
}

// 背景の欄。値は "#RRGGBB"（色）か { image, mode: tile | cover, tile, color }
// 画像の背景は、画像を並べて敷く（テクスチャ）。layout.yaml の mode: cover（広げる）は書けば効くが、ここでは選ばせない
const BG_OPTS = [['', 'なし'], ['color', '色'], ['tile', '画像']];
function bgRows(cur, write) {
  const kind = !cur ? '' : typeof cur !== 'object' ? 'color' : cur.image ? 'tile' : cur.color ? 'color' : '';
  const rows = [
    row('背景', selectEl(BG_OPTS, kind, (v) => {
      if (v === '') return write(undefined);
      if (v === 'color') return write(typeof cur === 'string' ? cur : cur?.color || '#ffffff');
      const keep = typeof cur === 'object' && cur ? cur : {};
      const image = (typeof cur === 'object' && cur?.image) || S.assets[0];
      if (image) return write({ ...keep, image, mode: 'tile' });
      pickImageFile().then((path) => (path ? write({ ...keep, image: path, mode: 'tile' }) : renderProps()));
    })),
  ];
  if (kind === 'color') rows.push(row('色', colorEl(typeof cur === 'string' ? cur : cur.color, (v) => write(v || undefined))));
  if (kind === 'tile') {
    const opts = [...new Set([...S.assets, cur.image])].map((a) => [a, a.split('/').pop()]);
    rows.push(row('画像', h2(selectEl(opts, cur.image, (v) => write({ ...cur, image: v })), pickBtn((path) => write({ ...cur, image: path })))));
    if (cur.mode !== 'cover') {
      const u = S.layout.unit;
      const nat = S.imageSizes.get(cur.image);
      const dflt = nat ? fmt((nat.w / 96) * PER_INCH[u]) : '';
      rows.push(row('画像1枚の幅', h2(numEl(cur.tile ?? dflt, '', (v) => +v > 0 && write({ ...cur, tile: +v })), u)));
    }
  }
  return rows;
}

const KIND_LABEL = { empty: '空き', split: '分割', grid: '格子', image: '画像', text: 'テキスト' };
const KIND_OPTS = [['text', 'テキスト'], ['image', '画像']];
const FIT_OPTS = [['contain', '切らずに全体を収める'], ['cover', 'ボックスいっぱいに広げ、はみ出た部分を切る'], ['stretch', '縦横比を変えてボックスに合わせる']];
const H_OPTS = [['left', '左'], ['center', '中央'], ['right', '右']];
const FONT_OPTS = FONTS.map((f) => [f, f]);
const JA_FONT_OPTS = JA_FONTS.map((f) => [f, f]);
const DEFAULT_SIZE = 24; // 文字の大きさの既定（pt）
const V_OPTS = [['top', '上'], ['middle', '中央'], ['bottom', '下']];


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
  if (title) f.append(h('legend', null, title));
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
      S.saved = { ...(S.saved || {}), [name]: f.text };
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
  const html = renderHTML(S.layout, { srcURL: (src) => S.imageURLs.get(src) || '', waitFonts: true });
  const f = document.createElement('iframe');
  f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  document.body.append(f);
  f.onload = async () => {
    // フォントを読み込み終えてから印刷する（待たないと、日本語が別のフォントで PDF に入る）
    await loadUsedFonts(f.contentDocument);
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
  const st = $('#status');
  st.textContent = msg;
  st.className = isErr ? 'err' : '';
  // プレビューの上にも、少しのあいだ目立つ形で出す
  const t = $('#toast');
  if (!t) return;
  clearTimeout(status.timer);
  if (!msg || msg === '読み込み中…') {
    t.hidden = true;
    return;
  }
  t.textContent = msg;
  t.className = isErr ? 'toast err' : 'toast';
  t.hidden = false;
  status.timer = setTimeout(() => (t.hidden = true), isErr ? 5000 : 2200);
}


function onKey(e) {
  if ((e.metaKey || e.ctrlKey) && e.code === 'KeyS') {
    e.preventDefault();
    save();
    return;
  }
  const tag = (e.target.tagName || '').toLowerCase();
  if (['input', 'textarea', 'select'].includes(tag) || e.target.isContentEditable) return;
  const mod = e.metaKey || e.ctrlKey;
  if (mod && (e.code === 'KeyZ' || e.code === 'KeyY')) {
    e.preventDefault();
    if (e.shiftKey || e.code === 'KeyY') redo();
    else undo();
    return;
  }
  if (mod && e.code === 'KeyC' && (S.sel || []).length) {
    e.preventDefault();
    copySelection();
    return;
  }
  if (mod && e.code === 'KeyV' && (S.sel || []).length) {
    if (!S.clip) {
      e.preventDefault();
      return status('先にコピーしてください（ボックスを選んで ⌘C）', true);
    }
    e.preventDefault();
    pasteSelection();
    return;
  }
  if (e.key === 'Escape') select(null);
}

// 選択中のボックスを分割する（分割数と、上下か左右かを選ぶ）
function openSplitDialog() {
  const box = S.selected ? findBox(S.selected) : null;
  if (!box || !box.src) return;
  const dlg = $('#split-dialog');
  // ボタンを押したその場で分割する（close の知らせは、画面が隠れていると遅れることがある）
  const run = (dir) => {
    const n = Math.max(2, Math.min(12, parseInt($('#split-n').value, 10) || 2));
    dlg.close();
    setSel([]);
    splitBox(box, dir, n);
  };
  $('#split-rows').onclick = () => run('rows');
  $('#split-cols').onclick = () => run('cols');
  $('#split-cancel').onclick = () => dlg.close();
  dlg.showModal();
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
$('#btn-split').onclick = openSplitDialog;
$('#btn-merge').onclick = () => mergeSelection();
$('#btn-dup').onclick = () => duplicateRow();
$('#btn-delete').onclick = () => removeSelection();
// 虫眼鏡：選んだボックスを左上に寄せ、いまの縦横比のまま、縦か横がプレビューいっぱいになるまで拡大する（つまみもそこまで動く）
$('#btn-zoom-fit').onclick = () => {
  const r = S.dragSel?.rect || selectionRect();
  const slide = curSlide();
  if (!r || !slide || !S.ppu) return;
  const frame = $('#frame');
  const m = 24;
  const fw = frame.clientWidth;
  const fh = frame.clientHeight;
  const R = slide.rect;
  const base = Math.min((fw - 2 * m) / (R.w * S.ppu), (fh - 2 * m) / (R.h * S.ppu));
  const z = Math.min((fw - 2 * m) / (r.w * S.ppu * base), (fh - 2 * m) / (r.h * S.ppu * base));
  setZoom(Math.max(1, Math.min(8, z)));
  S.scrollTarget = { x: r.x, y: r.y };
  applyView();
};
$('#zoom').oninput = (e) => {
  S.extraSpace = null;
  setZoom(+e.target.value / 100);
  applyView();
};
$('#btn-whole').onclick = () => {
  setSel([]);
  S.extraSpace = null;
  setZoom(1);
  applyView();
  renderProps();
};
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
window.slideCreator = { state: S, buildPPTX, finishPPTX, parse, rebuild, setSel, removeSelection, mergeSelection, curSlide }; // 開発者ツールからの確認用
// 開発用：URL に ?demo を付けると examples/demo を開く（画面にボタンは出さない）
if (new URLSearchParams(location.search).has('demo')) openSource(new UrlSource('../examples/demo/', 'demo'));
