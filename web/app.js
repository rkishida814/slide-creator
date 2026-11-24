// slide-creator ブラウザ版
import { resolveDoc } from '../src/core/resolve.js';
import { layoutDoc } from '../src/core/layout.js';
import { renderHTML, loadUsedFonts, measureUnderlines, hasImageUnderline, measureTextBgs, hasTextBg, FONTS, JA_FONTS, DEFAULT_FONT, DEFAULT_FONT_JA } from '../src/core/html.js';
import { richToHTML, paragraphs, bulletType, underlineHTML } from '../src/core/richtext.js';
import { buildPPTX, finishPPTX } from '../src/core/pptx.js';
import { PER_INCH, toUnit, defaultPadding, parseBox, isBgTree } from '../src/core/units.js';
import { FolderSource, UrlSource, download } from './source.js';
import { parse, stringify, setProp, replaceNode, getJS, YAML } from './edit.js';
import { openCrop } from './crop.js';
import { importLayout, placeImported, joinPath, dirOf } from '../src/core/import.js';

const $ = (s) => document.querySelector(s);
const IMAGE_RE = /\.(png|jpe?g|gif|svg|webp)$/i;
// HEIC（iPhone の写真）。ブラウザでは表示できないので、使うときに JPEG に変えてフォルダに入れる
const HEIC_RE = /\.(heic|heif)$/i;

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
  layoutFiles: [], // 取り込めるレイアウト（サブフォルダの layout.yaml）
  importFor: null, // 「レイアウトを取り込む」の欄を開いているボックスの path
  overflow: [],
  view: null,
  bgEdit: null, // 背景を編集しているボックスの path（このあいだ、プレビューで選ぶのは背景の枠）
  layer: 'main', // 選択対象：main（主レイヤー）/ bg（背景）/ block（既存レイアウト）
  importSigs: null, // 取り込み元の layout.yaml の分割の形（既存レイアウトを見分ける）
  autoBlocks: new Map(), // 取り込み元と分割の形が同じボックスの path → 取り込み元
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
    const all = await source.list('');
    S.assets = all.filter((f) => IMAGE_RE.test(f) && !f.startsWith('out/'));
    S.heics = all.filter((f) => HEIC_RE.test(f) && !f.startsWith('out/'));
    S.importSigs = null;
    S.layoutFiles = all.filter((f) => /\/layout\.ya?ml$/.test(f) && !/(^|\/)(out|templates)\//.test(f));
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
  if (isBgTree(node.background)) collectImages(node.background, out);
  if (node.icon) out.add(String(node.icon));
  // 文字の中の画像の線（<u img="…">）
  if (typeof node.text === 'string') for (const m of node.text.matchAll(/<u\s[^>]*img\s*=\s*(?:"([^"]*)"|([^\s>]+))/gi)) out.add(m[1] ?? m[2]);
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
// ボックスの配置と画像の並び。これが変わった書き込みでは、書いている途中でもプレビューを描き直す
function layoutSig() {
  const s = realSlide();
  if (!s) return '';
  const r = (x) => [x.x, x.y, x.w, x.h].map((v) => v.toFixed(2)).join(',');
  return [...s.boxes.map((b) => `${b.path}:${b.kind}:${r(b.rect)}`), ...s.elements.filter((e) => e.type !== 'text').map((e) => `${e.type}:${e.path}:${r(e.rect)}`)].join('|');
}

async function rebuild() {
  const sig0 = layoutSig();
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
  if (S.bgEdit && !realSlide()?.bgBoxes?.[S.bgEdit]) S.bgEdit = null;
  if (S.layer === 'block') refreshAutoBlocks();
  if ((S.sel || []).some((x) => !findBox(x))) setSel([]);
  status(S.error || '', !!S.error);
  renderSlideTabs();
  // プレビューに直接書いている途中の書き込み。描き直すと書く欄が作り直されるので、描き直さない（書き終えたときに描き直す）。
  // ただし、ほかの書き込み（余白のボックスを作るなど）と重なってボックスの配置が変わったときは描き直す
  const skip = S.skipPreview && layoutSig() === sig0;
  S.skipPreview = false;
  if (skip) syncSupport();
  else renderPreview();
  renderProps();
  $('#btn-save').disabled = !anyDirty();
}

// ---------------------------------------------------------------- 表示

function realSlide() {
  return S.layout?.slides.find((s) => s.id === S.slideId) || null;
}

// いま選べる枠の一覧を持つスライド。背景の編集中は、そのボックスの背景を分けた枠だけを選べる
function curSlide() {
  const s = realSlide();
  if (!s || S.layer !== 'bg') return s;
  return { ...s, boxes: (S.bgEdit && s.bgBoxes?.[S.bgEdit]) || [] };
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
    S.bgEdit = null;
    if (S.layer === 'block') refreshAutoBlocks();
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
.ied { outline: none; cursor: text; }
.txt { cursor: default; }
body.bgedit .txt, body.bgedit .img:not(.bg) { opacity: .25; }
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
    outline: S.outline || S.layer !== 'main',
    outlineBoxes: S.layer === 'bg' ? curSlide().boxes : S.layer === 'block' ? blockRoots() : undefined,
    srcURL: (src) => S.imageURLs.get(src) || '',
  }).replace('</style>', `${FRAME_CSS}</style>`);
  frame.onload = setupFrame;
  // プレビューに直接書いていたら、カーソルの位置を覚えておき、描き直したあとに戻す
  if (S.inline) {
    const ed = S.inline.ed;
    S.inlineRestore = { path: S.inline.path, pos: offsetsIn(ed), focus: ed.ownerDocument.activeElement === ed };
    S.support?.ctl.detach();
    S.inline = null;
  }
  // プレビューを描き直すあいだ、キー操作（⌘Z など）がどこにも届かなくならないように、プレビューにあったカーソルをアプリに戻す
  // （プレビューに直接書いていたときは外さない。描き直したあと、書く欄にカーソルを戻す）
  if (document.activeElement === frame && !S.inlineRestore?.focus) {
    frame.blur();
    window.focus();
  }
  frame.srcdoc = html;
}

// 選択。1 つならその枠の設定、2 つ以上なら比率と背景を出す
function setSel(paths) {
  if (S.inline && !(paths.length === 1 && paths[0] === S.inline.path)) finishInline();
  S.sel = paths;
  S.selected = paths.length === 1 ? paths[0] : null;
  if (S.importFor && S.importFor !== S.selected) S.importFor = null;
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
  doc.body.classList.toggle('bgedit', S.layer === 'bg');
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
    positionSupport();
  });
  // クリックで 1 つ選ぶ。ドラッグで、長方形になるように複数選ぶ。
  // ⌘（cmd。Windows では Ctrl）を押しながらドラッグすると、離した先のボックスと中身を入れ替える
  let drag = null;
  let dragged = false;
  let downInEd = false; // 書いている文字の中で押した（文字をドラッグで選んでいる）
  doc.addEventListener('mousedown', (e) => {
    downInEd = false;
    // 設定欄の入力欄にカーソルが残っていると、⌘Z がその欄の「元に戻す」になり、ボックスの操作が戻らない。プレビューを押したら外す
    const ae = document.activeElement;
    if (ae && ae !== document.body && ae !== frame) ae.blur();
    if (e.button !== 0) return;
    // 書いている文字の中を押したら、ふつうの文字の選択（ドラッグで範囲を選ぶ）にする
    if (S.inline && S.inline.ed.contains(e.target)) {
      downInEd = true;
      return;
    }
    // 背景では、いま開いている背景の枠だけをドラッグで選ぶ（別のボックスはクリックで開く）
    const b = S.layer === 'bg' ? leafAt(e) : pickAt(e);
    if (!b) return;
    e.preventDefault();
    drag = { box: b, x: e.clientX, y: e.clientY, moving: false, swap: e.metaKey || e.ctrlKey };
    // 複数選んだ行（横いっぱい）の中から ⌘ドラッグを始めたら、行ごと入れ替える
    if (S.layer !== 'block' && drag.swap && (S.sel || []).length > 1 && S.sel.includes(b.path)) drag.row = rowSwapPlan();
  });
  doc.addEventListener('mousemove', (e) => {
    if (!drag) return hover(S.layer === 'bg' ? bgHoverAt(e) : pickAt(e));
    if (!drag.moving && Math.hypot(e.clientX - drag.x, e.clientY - drag.y) > 6) {
      drag.moving = true;
      hover(null);
      if (drag.swap) doc.body.style.cursor = 'grabbing';
    }
    if (!drag.moving) return;
    const t = pickAt(e);
    if (drag.swap && drag.row) {
      showDrop(rowDropTarget(drag.row, t));
    } else if (drag.swap) {
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
    if (d.swap && d.row) {
      const t = leafAt(e);
      if (t) swapRows(d.row, t);
      return;
    }
    if (d.swap) {
      const t = pickAt(e);
      if (t && t.path !== d.box.path) {
        if (canSwap(d.box, t)) swapBoxes(d.box, t);
        else status('このボックスとは入れ替えられません。テンプレートの一部のボックスは入れ替えられません。', true);
      }
      return;
    }
    const r = S.dragSel;
    S.dragSel = null;
    if (r) setSel(S.layer === 'block' ? blocksLeaves(r.leaves) : r.leaves.map((l) => l.path));
    applyView();
    renderProps();
  });
  doc.addEventListener('mouseleave', () => hover(null));
  doc.addEventListener('click', (e) => {
    if (dragged) {
      dragged = false;
      return;
    }
    // 書いている文字の中（カーソルはそのまま）。文字の中で押してボックスの外まで選んで離したときも、ほかのボックスを選ばない
    if (downInEd || (S.inline && S.inline.ed.contains(e.target))) {
      downInEd = false;
      return;
    }
    const b = pickAt(e);
    if (S.layer === 'block') {
      setSel(b ? blockLeafPaths(b) : []);
      applyView();
      return renderProps();
    }
    if (S.layer === 'bg' && b && !curSlide().boxes.includes(b)) return switchBgHost(b, canvasPoint(e));
    if (S.layer === 'bg') {
      // 開いている背景の中でも、クリックした位置に、別の（もっと内側の）背景が付いたボックスがあれば、そちらを開く
      const r = realLeafAt(e);
      const owner = r ? bgOwnerOf(bgHostOf(r)) : null;
      if (owner && owner.path !== S.bgEdit && hasOwnBg(owner)) return switchBgHost(owner, canvasPoint(e));
    }
    select(b ? b.path : null);
    // ボックスを選んだら、そのまま書けるようにする（クリックした位置にカーソルを置く）
    if (b && S.layer === 'main') startInline({ point: { x: e.clientX, y: e.clientY } });
  });
  doc.addEventListener('keydown', onKey);
  applyView();
  measureOverflow(doc);
  // 選んでいるボックスは、描き直したあとも書ける状態にする（書いていたなら、同じ位置にカーソルを戻す）
  let rs = S.inlineRestore;
  S.inlineRestore = null;
  // 描き直しが始まったあとに書き始めたとき（書く欄が古いプレビューにある）は、新しいプレビューで書く欄にカーソルを入れ直す
  if (!rs && S.inline && S.inline.ed.ownerDocument !== doc) {
    rs = { path: S.inline.path, pos: null, focus: true };
    S.support?.ctl.detach();
    S.inline = null;
  }
  if (S.support && S.layer === 'main') startInline(rs && rs.path === S.support.path ? { pos: rs.pos, focus: rs.focus } : { focus: false });
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
  return fitBoxHeight(box, need, toUnit('0.5cm', u));
}

// 画像のボックスの高さを、表示している画像の高さ（余白を含む）まで縮め、下に余った部分を別のボックス（空き）にする
// 画像の上下に隙間があるか（ボックスの余白の分は除く）
function imageHasGap(box) {
  const el = curSlide().elements.find((e) => e.path === box.path && e.type === 'image');
  if (!el) return false;
  const pad = parseBox((box.node || {}).padding, S.layout.unit);
  return box.rect.h - pad.t - pad.b - el.rect.h > 0.01;
}

const L_rows = (b) => b.rect.h;

// 複数の画像のボックスを詰める。左右に並ぶ段の画像なら、段全体を、いちばん高い画像に合わせて縮める。
// それ以外は 1 つずつ詰める（詰めるたびに描き直されるので、順に待つ）
async function fitImagesBelow(paths) {
  const boxes = paths.map(findBox).filter(Boolean);
  const lca = boxes.length > 1 ? commonAncestor(boxes) : null;
  if (lca && lca.kind === 'split' && lca.node.split === 'cols') {
    const needOf = (b) => {
      const el = curSlide().elements.find((e) => e.path === b.path && e.type === 'image');
      const pad = parseBox((b.node || {}).padding, S.layout.unit);
      return el ? el.rect.h + pad.t + pad.b : b.rect.h;
    };
    const need = Math.ceil(Math.max(...boxes.map(needOf)) * 100) / 100;
    const cur = Math.max(...boxes.map((b) => b.rect.h));
    if (cur - need < 0.01) return status('画像の上下に余白はありません。');
    const r = resizeSelectionCross(lca, 'rows', need - cur);
    if (r && r.then) await r;
    return status('画像の上下の余白を詰めました');
  }
  for (const p of paths) {
    const b = findBox(p);
    if (!b || b.kind !== 'image' || !imageHasGap(b)) continue;
    const r = fitImageBelow(b);
    if (r && r.then) await r;
  }
  status('画像の上下の余白を詰めました');
}

function fitImageBelow(box) {
  const el = curSlide().elements.find((e) => e.path === box.path && e.type === 'image');
  if (!el) return;
  const u = S.layout.unit;
  const pad = parseBox((box.node || {}).padding, u);
  const need = Math.ceil((el.rect.h + pad.t + pad.b) * 100) / 100;
  if (box.rect.h - need < 0.01) return status('画像の上下に余白はありません。');
  // 空きのボックスは作らない。ボックスを縮め、空いた分はすぐ下のボックスが受け取る（下のボックスが上に伸びる）
  const c = sizeCtrl(box, 'rows');
  if (!c) return status('このボックスの高さは変えられません。', true);
  return resizeAlong(c, 'rows', c.child.path === box.path ? need : L_rows(c.child) - (box.rect.h - need));
}

// ボックスの高さを need にする。すでに上下の分割で高さが長さで決まり、すぐ下が空きなら、高さだけ合わせ直す。
// そうでなければ、上下に分けて下を空きにする（余りが minRest に満たないときは分けない）
function fitBoxHeight(box, need, minRest) {
  const u = S.layout.unit;
  const path = box.path;
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
  if (box.rect.h - need < minRest) return;
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
  // 選択は、文字の入ったボックス（分けた上側）に移す（複数選んでいるときは、そのボックスだけを置き換える）。
  // 書いている途中なら、書く欄も上側に移す（描き直したあと、同じ位置にカーソルを戻す）
  if (S.inline?.path === path) S.inline.path = `${path}.0`;
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
  const show = box && !(S.sel || []).includes(box.path) && !(S.layer === 'block' && blockLeafPaths(box).every((p) => S.sel.includes(p)));
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

// 2 つの長方形が重なるか。固定長の合計が枠を少し超えて、隣のボックスとわずかに重なることがあるので、
// スライドの幅の 0.05% 未満の重なりは、重なりとみなさない（端のボックスを含めて選ぶと、下の空きまで巻き込んでいた）
function overlaps(q, r) {
  const e = Math.max(1e-6, (S.layout?.width || 0) * 0.0005);
  return q.x < r.x + r.w - e && q.x + q.w > r.x + e && q.y < r.y + r.h - e && q.y + q.h > r.y + e;
}

// 2 つのボックスを含む長方形を作り、その長方形にかかるボックスが入るまで広げる。
// 広げる先は、2 つのボックスを両方含む、いちばん小さい組の中のボックスだけにする
// （ほかの組のボックスが、固定長の合計が領域を超えるなどで少しはみ出して重なっていても、巻き込まない）
function rectSelect(a, b) {
  const lca = a.path === b.path ? a : commonAncestor([a, b]);
  const leaves = curSlide().boxes.filter((l) => isLeaf(l) && (!lca || within(l, lca)));
  const hits = overlaps;
  let r = unionRect([a.rect, b.rect]);
  for (let k = 0; k < 50; k++) {
    const inc = leaves.filter((l) => hits(l.rect, r));
    const r2 = unionRect(inc.map((l) => l.rect));
    const same = Math.abs(r2.x - r.x) + Math.abs(r2.y - r.y) + Math.abs(r2.w - r.w) + Math.abs(r2.h - r.h) < 1e-6;
    r = r2;
    if (same) return { rect: r, leaves: inc };
  }
  return { rect: r, leaves: leaves.filter((l) => hits(l.rect, r)) };
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
  $('#layer').value = S.layer;
  $('#btn-dup').disabled = !dupPlan();
  $('#btn-merge').disabled = !((S.sel || []).length > 1 && selectionUnit());
  syncSupport();
  $('#btn-delete').disabled = !((S.sel || []).length > 1 ? selectionUnit() : one && canRemove(one));
}

// クリックした点を含む、いちばん内側のボックス（分割・格子の間の隙間なら null）
function canvasPoint(e) {
  const { s, offX, offY } = S.view;
  return { x: (e.pageX - offX) / s / S.ppu, y: (e.pageY - offY) / s / S.ppu };
}
function leafAt(e) {
  return leafIn(curSlide(), e);
}
// 背景を選んでいるときにも、中身のボックスを探す
function realLeafAt(e) {
  return leafIn(realSlide(), e);
}
function leafIn(slide, e) {
  if (!slide || !S.view) return null;
  const { x, y } = canvasPoint(e);
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
  if (S.support) S.support.stale = true;
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
function splitBox(box, dir, n, cols = 1) {
  // cols > 1：分けた各段を、さらに左右に cols 個に分ける（マス目で「行 × 列」を選んだとき）
  const grid = (kids, rest) => (cols > 1 ? kids.map((k, i) => ({ split: 'cols', children: [k, ...Array.from({ length: cols - 1 }, () => rest(i))] })) : kids);
  const { file, path, mode } = box.src;
  const doc = S.files[file].doc;
  if (S.layer === 'bg') {
    // 背景の枠を分けるときは、分けた枠のそれぞれに同じ背景を写す（見た目は変わらず、あとで 1 つずつ変えられる）
    const bg = plain((box.node || {}).background);
    const leaf = () => (bg != null ? { background: structuredClone(bg) } : {});
    const children = grid(Array.from({ length: n }, leaf), leaf);
    // 枠の余白は、分けた外側に残す（分けた枠は、余白の内側に並ぶ）
    const pad = plain((box.node || {}).padding);
    replaceNode(doc, path, { split: dir, ...(pad != null ? { padding: pad } : {}), children });
    return fileChanged(file);
  }
  let pn = plain(box.node || {});
  if (typeof pn === 'string') pn = { text: pn };
  const content = Object.fromEntries(Object.entries(pn || {}).filter(([k]) => !FRAME_KEYS.includes(k) && !MOVE_TO_CHILD.includes(k)));
  const pad = pn && typeof pn === 'object' && pn.padding != null ? { padding: pn.padding } : {};
  const hasContent = box.kind !== 'empty';
  // 余白は、分けた内側のボックスの 1 つずつに写す（選べば欄に見え、そこで直せる）
  const children = grid(Array.from({ length: n }, (_, i) => ({ ...(i === 0 && hasContent ? content : {}), ...structuredClone(pad) })), () => structuredClone(pad));
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
function splitEntries(lca, selSet, opts = {}) {
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
    // opts.expandAll：全部選んだ組でも、同じ向きの組なら展開する（比率を部分ごとに出すため）
    const sameDir = node && typeof node === 'object' && node.split === dir && Array.isArray(node.children);
    if (st !== 'part' && !(opts.expandAll && st === 'all' && sameDir)) {
      out.push({ node, ratio: r, len: L(box), sel: st === 'all', box });
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
// range（{ i0, i1 }）を渡すと、i0〜i1 番目の部分だけを 1 つの組にまとめて組み替え、ほかの部分はそのまま残す
// 例：[見出しの左右の組, 画像の左右の組, 文章] の上 2 段だけを、[左の列（見出し・画像）, 右の列] にする
function transposeSplit(lca, range = null) {
  if (!lca || lca.kind !== 'split' || !lca.src) return null;
  const D = lca.node.split;
  const P = D === 'rows' ? 'cols' : 'rows';
  const L = (b) => (P === 'rows' ? b.rect.h : b.rect.w);
  const slide = curSlide();
  const allKids = slide.boxes.filter((b) => b.parent === lca.path);
  const N = allKids.length;
  const { file, path, mode } = lca.src;
  const doc = S.files[file].doc;
  const written = mode === 'direct' ? getJS(doc, [...path, 'children']) : null;
  const allNodes = Array.isArray(written) && written.length === N ? written : plain(lca.node.children || []);
  if (allNodes.length !== N) return null;
  const allRatio = Array.isArray(lca.node.ratio) && lca.node.ratio.length === N ? plain(lca.node.ratio) : null;
  const sub = range && (range.i0 > 0 || range.i1 < N - 1);
  if (sub && (range.i0 < 0 || range.i1 >= N || range.i1 - range.i0 < 1)) return null;
  const kids = sub ? allKids.slice(range.i0, range.i1 + 1) : allKids;
  const nodes = sub ? allNodes.slice(range.i0, range.i1 + 1) : allNodes;
  const n = kids.length;
  const ok = ['split', 'ratio', 'children', 'gap', 'id', 'split_below'];
  const unit = S.layout.unit;
  const plainSplit = (nd) => nd && typeof nd === 'object' && nd.split === P && Array.isArray(nd.children) && !Object.entries(nd).some(([x, v]) => v != null && !ok.includes(x));
  // 部分 nd（ボックス b）の中身を、同じ向きの入れ子の組を開いて 1 列に並べる。[{ node, box, ratio }]。開けなければ null
  // 例：[見出し, [見出し, 空き]] → [見出し, 見出し, 空き]（入れ子の比率は、この列の比率 1 あたりの長さで直す）
  const flat = (nd, b, g0) => {
    const sb = slide.boxes.filter((x) => x.parent === b.path);
    if (sb.length !== nd.children.length) return null;
    const sr = Array.isArray(nd.ratio) && nd.ratio.length === sb.length ? nd.ratio : Array(sb.length).fill(1);
    let k = null;
    sb.forEach((x, j) => {
      if (k == null && typeof sr[j] === 'number' && sr[j] > 0) k = L(x) / sr[j];
    });
    const out = [];
    for (let j = 0; j < sb.length; j++) {
      const c = nd.children[j];
      if (plainSplit(c) && Math.abs(toUnit(c.gap ?? 0, unit) - g0) <= 1e-6) {
        const inner = flat(c, sb[j], g0);
        if (!inner) return null;
        for (const e of inner) out.push({ ...e, ratio: typeof e.ratio === 'string' ? e.ratio : k ? round4(L(e.box) / k) : `${round4(L(e.box))}${unit}` });
      } else out.push({ node: c, box: sb[j], ratio: sr[j] });
    }
    return out;
  };
  let gap0 = null;
  const flats = [];
  for (let i = 0; i < n; i++) {
    const nd = nodes[i];
    if (!plainSplit(nd)) return null;
    const g = toUnit(nd.gap ?? 0, unit);
    if (gap0 == null) gap0 = g;
    if (Math.abs(g - gap0) > 1e-6) return null;
    const f = flat(nd, kids[i], g);
    if (!f) return null;
    flats.push(f);
  }
  const m = flats[0]?.length;
  if (!m || m < 2 || flats.some((f) => f.length !== m)) return null;
  // 分かれ目の位置がそろっていること
  for (let i = 1; i < n; i++) for (let j = 0; j < m; j++) if (Math.abs(L(flats[i][j].box) - L(flats[0][j].box)) > 0.02) return null;
  const opened = flats.some((f, i) => f.length !== nodes[i].children.length);
  const r0 = opened ? flats[0].map((e) => e.ratio) : Array.isArray(nodes[0].ratio) && nodes[0].ratio.length === m ? nodes[0].ratio : null;
  const bandRatio = sub ? (allRatio || Array(N).fill(1)).slice(range.i0, range.i1 + 1) : allRatio;
  const bands = Array.from({ length: m }, (_, j) => ({
    split: D,
    ...(bandRatio ? { ratio: bandRatio } : {}),
    ...(lca.node.gap != null ? { gap: lca.node.gap } : {}),
    children: flats.map((f) => f[j].node),
  }));
  const content = { split: P, ...(r0 ? { ratio: r0 } : {}), ...(nodes[0].gap != null ? { gap: nodes[0].gap } : {}), children: bands };
  if (sub) {
    // 組み替えた組を、まとめた部分の位置に置く。組の長さは、まとめた部分の比率の和（長さで書いた部分があれば、いまの長さ）
    const u = S.layout.unit;
    const LD = (b) => (D === 'rows' ? b.rect.h : b.rect.w);
    const gap = toUnit(lca.node.gap ?? 0, u);
    const rr = allRatio || Array(N).fill(1);
    const part = rr.slice(range.i0, range.i1 + 1);
    const merged = part.every((x) => typeof x === 'number')
      ? round4(part.reduce((a, x) => a + x, 0))
      : `${round4(kids.reduce((a, b) => a + LD(b), 0) + gap * (n - 1))}${u}`;
    const children = [...allNodes.slice(0, range.i0), content, ...allNodes.slice(range.i1 + 1)];
    const ratio = [...rr.slice(0, range.i0), merged, ...rr.slice(range.i1 + 1)];
    keepOneFlexible(ratio, range.i0);
    setProp(doc, path, 'children', children);
    setProp(doc, path, 'ratio', ratio);
    return fileChanged(file);
  }
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

// 左右（上下）に並ぶ列の、上下（左右）の分け方の数がそろっていなくても、選んだ範囲の上端と下端（左端と右端）で
// どの列も分かれていれば、各列を「上・選んだ段・下」の 3 つに分けて、段ごとの左右の並びに組み替える。見た目は変わらない。
// 例：[[図, 空き, 空き], [図と文, 空き, 空き], [[ビーカー, 空き], 空き, 空き]] の 2 段目の空きを 3 つ選んだとき
// range（{ i0, i1 }）なら、その列だけを組み替え、ほかの列はそのまま残す。組み替えられないときは null
function bandSplit(lca, r, range = null) {
  if (!lca || lca.kind !== 'split' || !lca.src || lca.src.mode !== 'direct' || !r) return null;
  const D = lca.node.split;
  const P = D === 'rows' ? 'cols' : 'rows';
  const unit = S.layout.unit;
  const pos = (b) => (P === 'rows' ? b.rect.y : b.rect.x);
  const L = (b) => (P === 'rows' ? b.rect.h : b.rect.w);
  const top = P === 'rows' ? r.y : r.x;
  const bot = top + (P === 'rows' ? r.h : r.w);
  const tol = Math.max(0.02, (S.layout.width || 0) * 0.0005);
  const slide = curSlide();
  const allKids = slide.boxes.filter((b) => b.parent === lca.path);
  const N = allKids.length;
  const { file, path } = lca.src;
  const doc = S.files[file].doc;
  const allNodes = getJS(doc, [...path, 'children']);
  if (!Array.isArray(allNodes) || allNodes.length !== N) return null;
  const allRatio = Array.isArray(lca.node.ratio) && lca.node.ratio.length === N ? plain(lca.node.ratio) : Array(N).fill(1);
  const i0 = range ? range.i0 : 0;
  const i1 = range ? range.i1 : N - 1;
  if (i1 - i0 < 1) return null;
  const ok = ['split', 'ratio', 'children', 'gap', 'id', 'split_below'];
  const plainSplit = (nd) => nd && typeof nd === 'object' && nd.split === P && Array.isArray(nd.children) && !Object.entries(nd).some(([x, v]) => v != null && !ok.includes(x));
  // 列の中身を、同じ向きの入れ子を開いて 1 列に並べる（比率は、この列の比率 1 あたりの長さで直す）
  const flat = (nd, b, g0) => {
    const sb = slide.boxes.filter((x) => x.parent === b.path);
    if (sb.length !== nd.children.length) return null;
    const sr = Array.isArray(nd.ratio) && nd.ratio.length === sb.length ? nd.ratio : Array(sb.length).fill(1);
    let k = null;
    sb.forEach((x, j) => {
      if (k == null && typeof sr[j] === 'number' && sr[j] > 0) k = L(x) / sr[j];
    });
    const out = [];
    for (let j = 0; j < sb.length; j++) {
      const c = nd.children[j];
      if (plainSplit(c) && Math.abs(toUnit(c.gap ?? 0, unit) - g0) <= 1e-6) {
        const inner = flat(c, sb[j], g0);
        if (!inner) return null;
        for (const e of inner) out.push({ ...e, ratio: typeof e.ratio === 'string' ? e.ratio : k ? round4(L(e.box) / k) : `${round4(L(e.box))}${unit}` });
      } else out.push({ node: c, box: sb[j], ratio: sr[j] });
    }
    return out;
  };
  let gap0 = null;
  const cols = [];
  for (let i = i0; i <= i1; i++) {
    const nd = allNodes[i];
    if (!plainSplit(nd)) return null;
    const g = toUnit(nd.gap ?? 0, unit);
    if (gap0 == null) gap0 = g;
    if (Math.abs(g - gap0) > 1e-6) return null;
    const f = flat(nd, allKids[i], g);
    if (!f) return null;
    // 選んだ範囲の上端・下端で分かれているところ
    const a = f.findIndex((e) => Math.abs(pos(e.box) - top) <= tol);
    const z = f.findIndex((e) => Math.abs(pos(e.box) + L(e.box) - bot) <= tol);
    if (a < 0 || z < a) return null;
    cols.push([f.slice(0, a), f.slice(a, z + 1), f.slice(z + 1)]);
  }
  // どの列も、上・下の組があるか無いかがそろっていること
  for (const g of [0, 2]) if (new Set(cols.map((c) => c[g].length > 0)).size > 1) return null;
  const len = (es) => es.reduce((s2, e) => s2 + L(e.box), 0) + gap0 * Math.max(0, es.length - 1);
  const group = (es) => (es.length === 1 ? es[0].node : { split: P, ratio: es.map((e) => e.ratio), ...(gap0 ? { gap: `${round4(gap0)}${unit}` } : {}), children: es.map((e) => e.node) });
  const bands = [0, 1, 2].filter((g) => cols[0][g].length > 0);
  if (bands.length < 2) return null;
  const colRatio = allRatio.slice(i0, i1 + 1);
  const content = {
    split: P,
    ratio: bands.map((g) => round4(len(cols[0][g]))),
    ...(gap0 ? { gap: `${round4(gap0)}${unit}` } : {}),
    children: bands.map((g) => ({
      split: D,
      ratio: colRatio,
      ...(lca.node.gap != null ? { gap: lca.node.gap } : {}),
      children: cols.map((c) => group(c[g])),
    })),
  };
  if (range && (i0 > 0 || i1 < N - 1)) {
    const LD = (b) => (D === 'rows' ? b.rect.h : b.rect.w);
    const part = allRatio.slice(i0, i1 + 1);
    const gapD = toUnit(lca.node.gap ?? 0, unit);
    const merged = part.every((x) => typeof x === 'number')
      ? round4(part.reduce((s2, x) => s2 + x, 0))
      : `${round4(allKids.slice(i0, i1 + 1).reduce((s2, b) => s2 + LD(b), 0) + gapD * (i1 - i0))}${unit}`;
    const children = [...allNodes.slice(0, i0), content, ...allNodes.slice(i1 + 1)];
    const ratio = [...allRatio.slice(0, i0), merged, ...allRatio.slice(i1 + 1)];
    keepOneFlexible(ratio, i0);
    setProp(doc, path, 'children', children);
    setProp(doc, path, 'ratio', ratio);
  } else {
    const cur = getJS(doc, path) || {};
    const keep = Object.fromEntries(Object.entries(cur).filter(([x]) => !CONTENT_KEYS.includes(x) && x !== 'gap'));
    replaceNode(doc, path, { ...keep, ...content });
  }
  return fileChanged(file);
}

// 選んだ範囲がそのままでは扱えないとき、並びを組み替えてから、同じ範囲を選び直して fn をやり直す
function retryAfterTranspose(lca, fn) {
  const r = selectionRect();
  // 全部の部分を組み替えられなければ、選んだ範囲にかかる部分だけを組み替える。
  // それもできなければ、選んだ範囲の上端と下端で、各列を「上・選んだ段・下」に分けて組み替える
  const hit = r && lca ? curSlide().boxes.filter((b) => b.parent === lca.path).map((b, i) => (overlaps(b.rect, r) ? i : -1)).filter((i) => i >= 0) : [];
  const rng = hit.length >= 2 ? { i0: hit[0], i1: hit[hit.length - 1] } : null;
  const done = transposeSplit(lca) || (rng ? transposeSplit(lca, rng) : null) || (rng ? bandSplit(lca, r, rng) : null);
  if (!done) return null;
  return done.then(() => {
    const e = 0.02; // 組み替えで長さを丸めた分のずれを許す（transposeSplit の判定と同じ）
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
// 選んだボックスが、上下の並びの中の「横いっぱいの行」になっているとき、その上下の並びを返す
function rowSplitOf(boxes, selSet) {
  const slide = curSlide();
  const full = (b) => slide.boxes.filter((x) => isLeaf(x) && within(x, b)).every((l) => selSet.has(l.path));
  const isRows = (b) => b && b.kind === 'split' && b.node?.split === 'rows' && b.src;
  const a = boxes.length === 1 ? boxes[0] : commonAncestor(boxes);
  if (!a) return null;
  if (isRows(a) && !full(a)) return a;
  let x = a;
  for (;;) {
    if (!full(x) || !x.parent) return null;
    const p = findBox(x.parent);
    if (isRows(p)) return p;
    x = p;
  }
}

// ⌘ドラッグで行を入れ替える段取り：選んだ行（横いっぱい）と、その上下の並びの部分の一覧
function rowSwapPlan() {
  const sel = S.sel || [];
  if (!sel.length) return null;
  const boxes = sel.map(findBox).filter(Boolean);
  const selSet = new Set(sel);
  const R = rowSplitOf(boxes, selSet);
  if (!R) return null;
  const r = splitEntries(R, selSet);
  return r ? { R, r } : null;
}

// 選んだ行と、t を含む行を入れ替える（行の高さも一緒に動く）
function swapRows(plan, t) {
  const { R, r } = plan;
  const { entries, i0, i1 } = r;
  const j = entries.findIndex((e) => e.box && within(t, e.box));
  if (j < 0) return status('この行とは入れ替えられません（同じ上下の並びの中の行に落としてください）。', true);
  if (j >= i0 && j <= i1) return;
  const block = entries.slice(i0, i1 + 1);
  const list = j > i1
    ? [...entries.slice(0, i0), entries[j], ...entries.slice(i1 + 1, j), ...block, ...entries.slice(j + 1)]
    : [...entries.slice(0, j), ...block, ...entries.slice(j + 1, i0), entries[j], ...entries.slice(i1 + 1)];
  return writeEntries(R, list).then(() => status('行を入れ替えました'));
}

// ドラッグ中、落とす先の行を示す
function rowDropTarget(plan, t) {
  if (!t) return null;
  const e = plan.r.entries.find((x) => x.box && within(t, x.box));
  if (!e) return null;
  const k = plan.r.entries.indexOf(e);
  return k >= plan.r.i0 && k <= plan.r.i1 ? null : e.box;
}

function dupPlan() {
  const sel = S.sel || [];
  const slide = curSlide();
  if (!sel.length || !slide) return null;
  const boxes = sel.map(findBox).filter(Boolean);
  if (!boxes.length) return null;
  const selSet = new Set(sel);
  const full = (b) => slide.boxes.filter((x) => isLeaf(x) && within(x, b)).every((l) => selSet.has(l.path));
  const isRows = (b) => b && b.kind === 'split' && b.node?.split === 'rows' && b.src;
  const R = rowSplitOf(boxes, selSet);
  // 空きが足りなければ、外側の上下の並びでも探す
  for (let anc = R; isRows(anc); anc = anc.parent ? findBox(anc.parent) : null) {
    const r = splitEntries(anc, selSet);
    if (!r) break;
    const part = r.entries.slice(r.i0, r.i1 + 1);
    const D = part.reduce((s2, e) => s2 + e.len, 0) + r.gap * part.length;
    const room = r.entries.slice(r.i1 + 1).filter((e) => isEmptyNode(e.node)).reduce((s2, e) => s2 + e.len, 0);
    if (D <= room + 1e-6) return { anc, r, D };
  }
  return partDupPlan(boxes, selSet);
}

// 行の一部（左右に並ぶうちのいくつか）を選んだとき：選んだボックスを、同じ横の位置・幅・高さで、すぐ下に新しい行として足す。
// 新しい行の、選んでいない場所は空きにする。足した高さの分は、下のほうの空きを縮めて埋める（外側の上下の並びでも探す）
function partDupPlan(boxes, selSet) {
  const slide = curSlide();
  const u = selectionUnit();
  if (!u) return null;
  const target = u.kind === 'range' ? u.lca : u.box;
  if (!target?.src) return null;
  const unit = S.layout.unit;
  const sr = selectionRect();
  const D = sr.h;
  const kidsOf = (b) => slide.boxes.filter((x) => x.parent === b.path);
  const leavesOf = (b) => slide.boxes.filter((x) => isLeaf(x) && within(x, b)).map((x) => x.path);
  const copyOf = (b) => {
    let n = writtenNode(b);
    if (typeof n === 'string') n = { text: n };
    return stripIds(n);
  };
  // 左右の余白だけを残す（上下の余白は、新しい行の高さに入れない）
  const sidePad = (b) => {
    const p = parseBox(b.node?.padding, unit);
    return p.l || p.r ? { padding: [0, round4(p.r), 0, round4(p.l)] } : {};
  };
  // b の中の、選んだ所だけを写した形（左右の並びは、ほかを空きにして位置と幅を保つ。上下の並びは、選んだ所を含む部分だけを残す）
  const skel = (b) => {
    if (b.path === target.path) {
      if (u.kind === 'node') return copyOf(b);
      const ks = kidsOf(b);
      const n = b.node;
      return {
        split: n.split,
        ...(Array.isArray(n.ratio) && n.ratio.length === ks.length ? { ratio: plain(n.ratio) } : {}),
        ...(n.gap != null ? { gap: n.gap } : {}),
        ...sidePad(b),
        children: ks.map((k, i) => (i >= u.i0 && i <= u.i1 ? copyOf(k) : {})),
      };
    }
    if (b.kind !== 'split') return null;
    const ks = kidsOf(b);
    const c = ks.find((k) => within(target, k));
    if (!c) return null;
    const inner = skel(c);
    if (!inner) return null;
    const n = b.node;
    if (n.split === 'rows') {
      const pad = sidePad(b);
      return pad.padding ? { split: 'rows', ...pad, children: [inner] } : inner;
    }
    return {
      split: 'cols',
      ...(Array.isArray(n.ratio) && n.ratio.length === ks.length ? { ratio: plain(n.ratio) } : {}),
      ...(n.gap != null ? { gap: n.gap } : {}),
      ...sidePad(b),
      children: ks.map((k) => (k === c ? inner : {})),
    };
  };
  // 選んだ所を含む上下の並びを、内側から外側へ。下に十分な空きがあるところに足す
  for (let anc = target.parent ? findBox(target.parent) : null; anc; anc = anc.parent ? findBox(anc.parent) : null) {
    if (anc.kind !== 'split' || anc.node?.split !== 'rows' || !anc.src) continue;
    const E = kidsOf(anc).find((k) => within(target, k));
    if (!E) continue;
    const r = splitEntries(anc, new Set(leavesOf(E)));
    if (!r) continue;
    const room = r.entries.slice(r.i1 + 1).filter((e) => isEmptyNode(e.node)).reduce((s2, e) => s2 + e.len, 0);
    if (D + r.gap > room + 1e-6) continue;
    const node = skel(E);
    if (!node) return null;
    return { anc, r, D: D + r.gap, copies: [{ node, ratio: `${round4(D)}${unit}`, len: D }] };
  }
  return null;
}

function duplicateRow() {
  const plan = dupPlan();
  if (!plan) return status('複製できません（下の空きが足りません）。', true);
  const { anc, r, D } = plan;
  const { entries, i0, i1, gap, k, u } = r;
  // 複製は、元と同じ長さに固定する（中の名前は外す）
  const copies = plan.copies || entries.slice(i0, i1 + 1).map((e) => ({ node: stripIds(structuredClone(e.node ?? {})), ratio: `${round2(e.len)}${u}`, len: e.len }));
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
  let node = filled.length ? writtenNode(filled[0]) : S.bgEdit && leaves.length ? writtenNode(leaves[0]) : {};
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

// 設定欄を作り直す。テキストの欄（とその大きさの欄）を使っている最中なら、作り直した欄に、選んでいた範囲（先頭から何文字目か）と
// カーソルを写す（打った文字や大きさを書き込むと、ボックスを文字の高さに合わせるために描き直すことがあり、そのたびに欄が作り直されるため。
//   ボックスを直すとボックスの場所が変わることがあるので、古い欄は戻さず、新しい欄に移る）
function renderProps() {
  const act = document.activeElement;
  const oldWrap = act?.closest?.('.rich-wrap');
  const inUse = oldWrap && $('#node').contains(oldWrap) ? oldWrap : null;
  let pos = null;
  let toSize = false;
  if (inUse) {
    const oldEd = inUse.querySelector('.rich');
    toSize = act.classList.contains('rich-size');
    const sel = getSelection();
    const r = sel.rangeCount && oldEd.contains(sel.getRangeAt(0).startContainer) ? sel.getRangeAt(0) : inUse._getSaved?.();
    if (r && oldEd.contains(r.startContainer) && oldEd.contains(r.endContainer)) {
      const off = (node, o) => {
        const x = document.createRange();
        x.setStart(oldEd, 0);
        x.setEnd(node, o);
        return x.toString().length;
      };
      pos = { start: off(r.startContainer, r.startOffset), end: off(r.endContainer, r.endOffset) };
    }
  }
  renderPropsInner();
  if (!inUse || inUse.isConnected) return;
  const nw = $('#node').querySelector('.rich-wrap');
  if (!nw) return;
  const ed = nw.querySelector('.rich');
  // 先頭から n 文字目の位置（文字の節と、その中の位置）。節の境目では、範囲の始まりは次の節の頭に、終わりは前の節の末尾に置く
  const at = (n, isStart) => {
    const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
    let last = null;
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      if (isStart ? n < t.nodeValue.length : n <= t.nodeValue.length) return [t, n];
      n -= t.nodeValue.length;
      last = t;
    }
    return last ? [last, last.nodeValue.length] : [ed, ed.childNodes.length];
  };
  let r = null;
  if (pos) {
    r = document.createRange();
    r.setStart(...at(pos.start, pos.start < pos.end));
    r.setEnd(...at(pos.end, false));
  }
  if (toSize) {
    // 大きさの欄を使っていたら、その欄にカーソルを戻し、選んでいた範囲を覚えさせる
    nw._setSaved?.(r);
    nw.querySelector('.rich-size')?.focus();
    return;
  }
  ed.focus();
  if (r) {
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }
}

function renderPropsInner() {
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
  if (S.layer === 'bg') return renderBgProps(el);
  if (S.layer === 'block') return renderBlockProps(el);
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
  // 文字・画像は、プレビューのボックスに直接書き、サポートパネルで設定する（この欄には出さない）
  if (isLeaf(box) && S.importFor !== box.path) {
    el.append(row('', btn('レイアウトを取り込む…', () => {
      S.importFor = box.path;
      renderProps();
    })));
  }

  // 中身
  if (S.importFor === box.path && isLeaf(box)) el.append(importFields(box));
  if (box.kind === 'image') el.append(imageFields(box, node, inh));

  // 枠
  const fsBox = fieldset('余白・背景・枠線');
  fsBox.append(
    paddingRow(box, node),
    ...(isBgTree(node.background)
      ? [row('背景', h('span', 'note', '分けて設定している（「背景を編集」で変える）'))]
      : bgRows(node.background, (v) => edit(box, 'background', v), { rect: box.rect })),
    row('', bgEditBtn(box)),
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
  // 背景の枠には、決まった余白が無い
  const def = S.layer === 'bg' ? null : defaultPadding(box.kind);
  const u = S.layout.unit;
  const defV = def ? def.map((x) => Math.round(toUnit(x, u) * 1000) / 1000) : null;
  const out = a.every((y) => y === a[0]) ? a[0] : a;
  if (defV ? a.every((y, k) => y === defV[k]) : out === 0) return undefined;
  return out;
}

// 余白の欄。targets（省略すると box だけ）の 1 つずつに、同じ余白を書く。値は box のものを出す
function paddingRow(box, node, targets) {
  const list = targets || [box];
  // テキストのボックスで余白を書いていないときは、決まった余白（上下左右に 0.5 cm）を出す（背景の枠には無い）
  const bgMode = S.layer === 'bg';
  const def = bgMode ? null : defaultPadding(box.kind);
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
  // ふだんは、上下と左右を 1 つずつの欄で同時に変える（ほかの 2 つの組の行と同じ並べ方）
  const pairIn = (ks) => {
    const same = v[ks[0]] === v[ks[1]];
    return numEl(same ? v[ks[0]] ?? 0 : '', same ? '0' : '別々', (x) => {
      for (const k of ks) v[k] = x === '' ? 0 : +x;
      write();
    }, 0.5);
  };
  const pair = pairLine(['余白（上下）', pairIn([0, 2])], ['余白（左右）', pairIn([1, 3])]);
  // 上・右・下・左を別々に変える欄は、折りたたんでおく（上下・左右がそろっていないときは開く）
  const four = h('span', 'in pad4');
  [['↑', '上'], ['→', '右'], ['↓', '下'], ['←', '左']].forEach(([icon, title], k) => {
    const ic = h('span', 'pad-ic', icon);
    ic.title = title;
    const inp = numEl(v[k] ?? 0, '0', (x) => {
      v[k] = x === '' ? 0 : +x;
      write();
    }, 0.5);
    four.append(ic, inp);
  });
  const det = document.createElement('details');
  det.className = 'pad-detail';
  det.open = S.padOpen || v[0] !== v[2] || v[1] !== v[3];
  det.ontoggle = () => (S.padOpen = det.open);
  const sum = document.createElement('summary');
  sum.textContent = '上・右・下・左を別々に設定する';
  det.append(sum, four);
  const r = fullRow(pair);
  const rd = row('', det);
  const inset = checkEl('背景ごと余白を設ける', !!node.background_inset, (x) => writeAll('background_inset', () => x || undefined));
  const r2 = row('', inset);
  const both = h('div');
  // 背景の枠では、余白の分だけ背景が縮むので、「背景ごと余白を設ける」は出さない
  both.append(r, rd, ...(bgMode ? [] : [r2]));
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

  // 選んだ部分どうしの比率。同じ向きに分けた組が入れ子になっていても、外側の並びに展開して、選んだ部分の数だけ出す
  const ent = lca.kind === 'split' ? splitEntries(lca, new Set(S.sel), { expandAll: true }) || splitEntries(lca, new Set(S.sel)) : null;
  if (ent && ent.i1 - ent.i0 + 1 >= 2) {
    const { entries, i0, i1 } = ent;
    const ratio = entries.map((e) => (typeof e.ratio === 'number' ? Math.round(e.ratio * 10000) / 10000 : e.ratio));
    const parts = [];
    for (let i = i0; i <= i1; i++) {
      if (i > i0) parts.push(h('span', 'ratio-sep', ':'));
      const t = textEl(String(ratio[i]), '1', (x) => {
        const pv = parseLen(x);
        if (pv === undefined) return;
        ratio[i] = pv;
        keepOneFlexible(ratio, i);
        setRatioEntries(lca, entries, ratio);
      });
      t.className = 'ratio-in';
      parts.push(t);
    }
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
  // 画像のボックスを選んでいて、上下に隙間のあるものがあれば、まとめて詰めるボタンを出す
  const gapImgs = boxes.filter((b) => b.kind === 'image' && b.src && imageHasGap(b));
  if (gapImgs.length) {
    const fsI = fieldset('画像（選んだボックスの 1 つずつ）');
    const acts = h('div', 'actions');
    acts.append(btn('画像の上下の余白を詰める', () => fitImagesBelow(gapImgs.map((b) => b.path))));
    fsI.append(acts);
    el.append(fsI);
  }
  // 横の位置・縦の位置を、選んだボックス（テキスト・画像）の 1 つずつに設定する（値は、いちばん上か左のボックスのものを出す）
  const posBoxes = boxes.filter((b) => (b.kind === 'text' || b.kind === 'image' || b.kind === 'empty') && b.src)
    .sort((a, b) => curSlide().boxes.indexOf(a) - curSlide().boxes.indexOf(b));
  if (posBoxes.length) {
    const f0 = posBoxes[0];
    const dflt = (b, key) => (b.kind === 'image' ? (key === 'align' ? 'center' : 'middle') : key === 'align' ? 'left' : 'top');
    const cur = (key) => (f0.node || {})[key] ?? (f0.inherited || {})[key] ?? dflt(f0, key);
    const setAll = (key, v) => {
      const files = new Set();
      for (const b of posBoxes) {
        setProp(S.files[b.src.file].doc, b.src.path, key, v);
        files.add(b.src.file);
      }
      for (const f of files) fileChanged(f);
    };
    const fsP = fieldset('位置（選んだボックスの 1 つずつ）');
    fsP.append(fullRow(pairLine(
      ['横の位置', selectEl(H_OPTS, cur('align'), (v) => setAll('align', v))],
      ['縦の位置', selectEl(V_OPTS, cur('valign'), (v) => setAll('valign', v))],
    )));
    el.append(fsP);
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
    // 選んだボックスだけを伸び縮みさせる（上・左のボックスは動かさない）
    if (lca.kind === 'split' && lca.node.split !== dir) {
      return numEl(r3(cur), '', (v) => +v > 0 && resizeSelectionCross(lca, dir, +v - cur));
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
  fs.append(fullRow(pairLine(['幅', dimInput('cols'), u], ['高さ', dimInput('rows'), u])));
  return fs;
}

// 左右に並ぶ部分（lca）にまたがって選んだボックスの高さ（上下に並ぶときは幅）を delta だけ変える。
// 1 つ目の部分では resizeAlong で外側まで伸ばし（下のボックスとやり取りする）、ほかの部分では、選んだボックスを同じだけ伸ばす
function resizeSelectionCross(lca, dir, delta) {
  if (Math.abs(delta) < 1e-6) return;
  const u = S.layout.unit;
  const L = (b) => (dir === 'rows' ? b.rect.h : b.rect.w);
  const leaves = (S.sel || []).map(findBox).filter(Boolean);
  const cs = [];
  const seen = new Set();
  for (const l of leaves) {
    const c = sizeCtrl(l, dir);
    if (!c || seen.has(c.child.path)) continue;
    seen.add(c.child.path);
    cs.push(c);
  }
  if (!cs.length) return status('この範囲の大きさは変えられません。', true);
  // 選んだボックスが、それぞれの部分の中で上下（左右）に分けた一部なら、部分ごとに伸ばす
  const inside = cs.every((c) => c.p.path !== lca.path && within(c.p, lca));
  if (!inside) {
    const c = sizeCtrl(lca, dir);
    if (!c) return status('この範囲の大きさは変えられません。', true);
    return resizeAlong(c, dir, L(c.child) + delta);
  }
  const extra = [];
  for (const c of cs.slice(1)) {
    const sibs = curSlide().boxes.filter((b) => b.parent === c.p.path);
    const ratio = Array.isArray(c.p.node.ratio) && c.p.node.ratio.length === sibs.length ? plain(c.p.node.ratio) : Array(sibs.length).fill(1);
    ratio[c.i] = `${round2(L(c.child) + delta)}${u}`;
    extra.push([c.p, ratio]);
  }
  return resizeAlong(cs[0], dir, L(cs[0].child) + delta, extra);
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

// 並び lca に、展開した部分の一覧 entries と比率 ratio を書き戻す。
// 入れ子の組を展開したときは並びが組み替わるので、同じ範囲のボックスを選び直す
function setRatioEntries(lca, entries, ratio) {
  const kids = curSlide().boxes.filter((b) => b.parent === lca.path);
  if (entries.length === kids.length) return edit(lca, 'ratio', ratio);
  const { file, path } = lca.src;
  const doc = S.files[file].doc;
  setProp(doc, path, 'children', entries.map((e) => e.node));
  setProp(doc, path, 'ratio', ratio);
  return fileChanged(file).then(() => {
    // 比率を変えると位置が変わるので、選んでいた部分（並びの中の同じ番号）を選び直す
    const R = findBox(lca.path);
    if (!R) return;
    const nk = curSlide().boxes.filter((b) => b.parent === R.path);
    const picked = nk.filter((b, i) => entries[i] && entries[i].sel);
    const leaves = curSlide().boxes.filter((b) => isLeaf(b) && picked.some((p) => within(b, p)));
    setSel(leaves.map((b) => b.path));
    applyView();
    renderProps();
  });
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
    // 描き直して画面から外れた欄の知らせ（遅れて届く）は無視する
    if (!fold.isConnected) return;
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
    row('文字の大きさ', h2(numEl(getJS(doc, ['defaults', 'font', 'size']) ?? DEFAULT_SIZE, '', (v) => {
      if (!(+v > 0)) return;
      setProp(doc, ['defaults'], ['font', 'size'], +v);
      fileChanged('layout.yaml');
    }), 'pt')),
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
    if (!c && S.layer === 'bg' && box.src) {
      // 背景の枠で、分けていない向き：背景の長さ（枠から余白を引いた長さ）を出し、変えると、左（上）はそのままに、右（下）の余白で合わせる
      const p = parseBox(box.node?.padding, u);
      const near = dir === 'rows' ? p.t : p.l;
      const far = dir === 'rows' ? p.b : p.r;
      const frame = lenOf(box);
      return numEl(r3(frame - near - far), '', (v) => {
        const nf = frame - near - +v;
        if (!(+v > 0) || nf < -1e-6) return status(`背景の長さは、0 より大きく、枠の長さ（${fmt(frame - near)} ${u}）以下にしてください。`, true);
        const a = [p.t, p.r, p.b, p.l].map((x) => Math.round(x * 1000) / 1000);
        a[dir === 'rows' ? 2 : 1] = Math.round(Math.max(0, nf) * 1000) / 1000;
        return edit(box, 'padding', padValue(box, a));
      });
    }
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
    // 長さで指定：上（左）のボックスはそのままにし、増えた分・減った分は、すぐ下（右）のボックスとやり取りする
    if (mode === 'len') return numEl(r3(lenOf(c.child)), '', (v) => +v > 0 && (resizeAlong(c, dir, +v) ?? write(`${+v}${u}`)));
    let val = ratio[c.i];
    if (typeof val === 'string') {
      // 長さで決めている部分を比率で表すときは、比率で決めているほかの部分を基準にする
      const j = c.kids.findIndex((k, ix) => ix !== c.i && typeof ratio[ix] === 'number' && ratio[ix] > 0);
      val = j >= 0 ? r3((lenOf(c.child) / lenOf(c.kids[j])) * ratio[j]) : 1;
    }
    return numEl(val, '', (v) => +v > 0 && write(+v));
  };
  const unitTxt = mode === 'len' ? u : '';
  fs.append(fullRow(pairLine(['幅', dimInput('cols'), unitTxt], ['高さ', dimInput('rows'), unitTxt])));
  if (shared) fs.append(h('p', 'warn', 'この大きさは、同じテンプレートを使う全てのスライドで共通です。'));
  return fs;
}

// ボックス（c.child）の長さを v にする。上（左）のボックスは動かさず、増減の分は、すぐ下（右）のボックスとやり取りする。
// いちばん下（右）なら外側へたどり、外側の並びで、すぐ下（右）のボックスとやり取りする（途中の枠は一緒に伸び縮みする）。
// 扱えない形（格子の中など）なら null を返す
function resizeAlong(c, dir, v, extra = []) {
  const u = S.layout.unit;
  const L = (b) => (dir === 'rows' ? b.rect.h : b.rect.w);
  const delta = v - L(c.child);
  if (Math.abs(delta) < 1e-6) return true;
  const writes = [];
  let x = c.child;
  let q = c.p;
  let first = true;
  while (q) {
    if (q.kind === 'split' && q.node?.split === dir && q.src) {
      const sibs = curSlide().boxes.filter((b) => b.parent === q.path);
      const i = sibs.findIndex((b) => b.path === x.path);
      if (i < 0) return null;
      const ratio = Array.isArray(q.node.ratio) && q.node.ratio.length === sibs.length ? plain(q.node.ratio) : Array(sibs.length).fill(1);
      let k = null;
      sibs.forEach((b, j) => {
        if (k == null && typeof ratio[j] === 'number' && ratio[j] > 0) k = L(b) / ratio[j];
      });
      const change = (j, d) => {
        if (typeof ratio[j] === 'string' || !k) ratio[j] = `${round2(L(sibs[j]) + d)}${u}`;
        else ratio[j] = round4(ratio[j] + d / k);
      };
      if (first) ratio[i] = `${round2(v)}${u}`;
      else change(i, delta);
      first = false;
      writes.push([q, ratio]);
      const commit = () => {
        const files = new Set();
        for (const [qq, r] of [...extra, ...writes]) {
          setProp(S.files[qq.src.file].doc, qq.src.path, 'ratio', r);
          files.add(qq.src.file);
        }
        for (const f of files) fileChanged(f);
        return true;
      };
      // 縮めるとき：空いた分は、すぐ下（右）のボックスが受け取る
      if (delta < 0 && i + 1 < sibs.length) {
        change(i + 1, -delta);
        return commit();
      }
      // 伸ばすとき：下（右）のボックスは押し下げ、その分は、下のほうの空きのボックスから取る。
      // この並びの空きで足りなければ、この並びごと伸ばし、外側の並びの空きから取る
      if (delta > 0) {
        const empties = sibs.map((b, j) => [b, j]).filter(([b, j]) => j > i && b.kind === 'empty');
        const room = empties.reduce((a, [b]) => a + L(b), 0);
        if (room >= delta - 1e-6) {
          let need = delta;
          for (let t = empties.length - 1; t >= 0 && need > 1e-6; t--) {
            const [b, j] = empties[t];
            const take = Math.min(L(b), need);
            change(j, -take);
            need -= take;
          }
          return commit();
        }
      }
    } else if (!(q.kind === 'split' && q.src)) {
      return null;
    }
    // いちばん下（右）か、向きの違う並び：外側へたどる
    x = q;
    q = q.parent ? findBox(q.parent) : null;
  }
  status(dir === 'rows' ? '下に空きのボックスが足りないので、この高さにはできません。' : '右に空きのボックスが足りないので、この幅にはできません。', true);
  return true;
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
  // 画像の上下に隙間があるときだけ、詰めるボタンを出す
  if (el && imageHasGap(box)) acts.append(btn('画像の上下の余白を詰める', () => fitImageBelow(box)));
  if (node.crop) acts.append(btn('トリミングを外す', () => edit(box, 'crop', undefined)));
  fs.append(acts);
  if (el) fs.append(h('div', 'note', `元画像 ${el.natural.w} × ${el.natural.h} px、表示する大きさ ${fmt(el.rect.w)} × ${fmt(el.rect.h)} ${S.layout.unit}、解像度 ${el.dpi} dpi`));
  return fs;
}

// ---------------------------------------------------------------- レイアウトを取り込む
// サブフォルダの layout.yaml の 1 枚の、ボックスの木を写す。写した後はこの文書の普通のボックスになる
async function readSubLayout(file) {
  const sub = parse(await S.source.readText(file)).toJS() || {};
  const templates = {};
  for (const f of await S.source.list(joinPath(dirOf(file), 'templates'), 0)) {
    if (/\.ya?ml$/.test(f)) templates[f.split('/').pop().replace(/\.ya?ml$/, '')] = parse(await S.source.readText(f)).toJS();
  }
  return resolveDoc(sub, templates);
}

function importFields(box) {
  const fs = fieldset('レイアウトを取り込む');
  if (!S.layoutFiles.length) {
    fs.append(h('p', 'note', 'このフォルダの中のサブフォルダに、取り込める layout.yaml がありません（例：fig_intro/layout.yaml）。'));
    const acts0 = h('div', 'actions');
    acts0.append(btn('やめる', () => { S.importFor = null; renderProps(); }));
    fs.append(acts0);
    return fs;
  }
  const opts = S.layoutFiles.map((f) => [f, f.replace(/\/layout\.ya?ml$/, '')]);
  let file = opts[0][0];
  let page = null;
  const pageRow = h('div');
  const loadPages = async () => {
    pageRow.innerHTML = '';
    page = null;
    try {
      const sub = await readSubLayout(file);
      const ids = (sub.slides || []).map((x, i) => x.id || `s${i + 1}`);
      if (ids.length > 1) {
        page = ids[0];
        pageRow.append(row('ページ', selectEl(ids.map((id, i) => [id, `${i + 1}ページ（${id}）`]), page, (v) => { page = v; })));
      }
    } catch (e) {
      pageRow.append(h('p', 'warn', `読めません: ${e.message}`));
    }
  };
  fs.append(
    row('レイアウト', selectEl(opts, file, (v) => { file = v; loadPages(); })),
    pageRow,
    h('p', 'note', 'そのレイアウトのボックスの組み合わせ（分割・比率・文字・画像）を、このボックスの中身として写します。写した後は、このスライドの普通のボックスとして直せます（元のファイルとはつながりません）。'),
  );
  const acts = h('div', 'actions');
  acts.append(btn('取り込む', () => doImport(box, file, page)), btn('やめる', () => { S.importFor = null; renderProps(); }));
  fs.append(acts);
  loadPages();
  return fs;
}

async function doImport(box, file, page) {
  if (!box.src || box.src.mode !== 'direct') return status('取り込めるのは、layout.yaml に直接書いたボックスだけです（テンプレートの中は不可）', true);
  try {
    const sub = await readSubLayout(file);
    const node = box.node && typeof box.node === 'object' ? box.node : {};
    const pad = parseBox(node.padding, S.layout.unit);
    const rect = { w: box.rect.w - pad.l - pad.r, h: box.rect.h - pad.t - pad.b };
    const usedIds = curSlide().boxes.filter((b) => b.id && b.path !== box.path).map((b) => b.id);
    const { node: imported, scale, dropped } = importLayout(sub, { slide: page, dir: dirOf(file), unit: S.layout.unit, rect, usedIds });
    const target = Object.fromEntries(Object.entries(plain(node) || {}).filter(([k]) => typeof k === 'string'));
    // 取り込み元の印を残す（選択対象「既存レイアウト」で、塊ごと選ぶため）
    imported.from_layout = file;
    replaceNode(S.files[box.src.file].doc, box.src.path, placeImported(target, imported));
    S.importFor = null;
    await fileChanged(box.src.file);
    const msg = [`${file} を取り込みました（倍率 ${+(scale * 100).toFixed(1)}%）`];
    if (dropped.length) msg.push(`このスライドですでに使っている id を外しました: ${dropped.join(', ')}`);
    status(msg.join('。'));
  } catch (e) {
    status(`取り込めません: ${e.message}`, true);
  }
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
  return fullRow(pairLine(
    ['揃え方（横）', selectEl(H_OPTS, node.align ?? inh.align ?? hDflt, (v) => edit(box, 'align', v))],
    ['揃え方（縦）', selectEl(V_OPTS, node.valign ?? inh.valign ?? vDflt, (v) => edit(box, 'valign', v))],
  ));
}

// 1 行に 2 つの組（名前・入力欄・単位）を並べる。1 つ目の名前は、ほかの行と同じ名前の列に置き、
// 1 つ目の入力欄は、ほかの行の入力欄と左端をそろえる。2 つ目の名前と入力欄は、同じ行の右半分に置く
function pairLine(a, b) {
  const line = h('span', 'in pair-line');
  const group = (ctl, unit) => {
    const g = h('span', 'pair');
    g.append(ctl);
    if (unit) g.append(h('span', 'pair-unit', unit));
    return g;
  };
  line.append(group(a[1], a[2]), h('label', 'pair-label', b[0]), group(b[1], b[2]));
  line._rowLabel = a[0];
  return line;
}
function fullRow(line) {
  return row(line._rowLabel || '', line);
}

// 文字の下地（text_bg）：付けるか・色・透明度（%）の 1 行
function textBgRows(box, node, edit) {
  const raw = node.text_bg;
  const cur = raw && typeof raw === 'object' ? plain(raw) : raw ? { color: String(raw) } : null;
  const on = !!cur?.color;
  const write = (patch) => {
    const next = { ...(cur || {}), ...patch };
    for (const k of Object.keys(next)) if (next[k] === undefined) delete next[k];
    return edit(box, 'text_bg', next);
  };
  const chk = document.createElement('input');
  chk.type = 'checkbox';
  chk.checked = on;
  chk.onchange = () => (chk.checked ? write({ color: S.lastTextBg || '#ffff00' }) : edit(box, 'text_bg', undefined));
  const c = document.createElement('input');
  c.type = 'color';
  c.setAttribute('list', refreshColorList());
  c.value = normColor(cur?.color) || S.lastTextBg || '#ffff00';
  c.onchange = () => {
    S.lastTextBg = c.value;
    return write({ color: c.value });
  };
  const t = numEl(on ? Math.round((1 - Number(cur.opacity ?? 1)) * 100) : 0, '', (v) => {
    const x = Math.min(100, Math.max(0, +v || 0));
    return write({ color: cur?.color || c.value, opacity: x ? Math.round((1 - x / 100) * 100) / 100 : undefined });
  }, 5);
  t.className = 'tb-alpha';
  t.disabled = !on;
  return [row('下地', h2(chk, c, h('span', null, '透明度'), t, '%'))];
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
  input.accept = 'image/png,image/jpeg,image/gif,image/svg+xml,image/webp,image/heic,image/heif,.heic,.heif';
  let file = await new Promise((res) => {
    input.onchange = () => res(input.files[0] || null);
    input.addEventListener('cancel', () => res(null));
    input.click();
  });
  if (!file) return null;
  let name = file.name.replace(/[\\/:*?"<>|]/g, '_');
  if (HEIC_RE.test(name)) {
    // HEIC は JPEG に変えてから入れる（名前の拡張子も .jpg にする）
    try {
      status('HEIC の写真を JPEG に変えています…');
      file = await heicToJpeg(file);
      status('');
    } catch (e) {
      status(`HEIC の写真を読めません: ${e.message || e}`, true);
      return null;
    }
    name = name.replace(HEIC_RE, '.jpg');
  }
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

// HEIC の写真を JPEG に変える。変換のライブラリ（heic2any）は、初めて使うときに読み込む
let heicLib = null;
function heicToJpeg(blob) {
  heicLib ||= new Promise((res, rej) => {
    const sc = document.createElement('script');
    sc.src = 'vendor/heic2any.min.js';
    sc.onload = () => res(window.heic2any);
    sc.onerror = () => rej(new Error('変換のライブラリを読み込めません'));
    document.head.append(sc);
  });
  return heicLib.then((f) => f({ blob, toType: 'image/jpeg', quality: 0.92 })).then((r) => (Array.isArray(r) ? r[0] : r));
}

// フォルダの HEIC を JPEG に変えて、隣に「名前_heic.jpg」として入れ、その場所を返す（すでにあれば、それを使う）
function heicJpegPath(path) {
  return path.replace(HEIC_RE, '_heic.jpg');
}
async function useHeic(path) {
  const out = heicJpegPath(path);
  if (!(await S.source.exists(out))) {
    status(`${path.split('/').pop()} を JPEG に変えています…`);
    const jpg = await heicToJpeg(await S.source.readBlob(path));
    await S.source.writeFile(out, jpg);
  }
  if (!S.assets.includes(out)) S.assets.push(out);
  S.pickMsg = `${path.split('/').pop()} を JPEG に変えて ${out} として使います`;
  return out;
}

// ▾ のメニューが画面の下にはみ出すときは、上に開く。上にも入りきらなければ、広いほうに開いてスクロールできるようにする
function fitMenu(menu) {
  menu.style.top = menu.style.bottom = menu.style.maxHeight = menu.style.overflowY = '';
  const r = menu.getBoundingClientRect();
  const below = innerHeight - r.top - 8;
  if (r.height <= below) return;
  const host = menu.offsetParent?.getBoundingClientRect();
  const above = host ? host.top - 8 : 0;
  if (above > below) {
    menu.style.top = 'auto';
    menu.style.bottom = '100%';
  }
  const room = Math.max(above, below);
  if (r.height > room) {
    menu.style.maxHeight = `${Math.max(120, room)}px`;
    menu.style.overflowY = 'auto';
  }
}

// 文字の書式の道具（サポートパネルの書式のボタンと、文字を書く欄の操作）。
// 書く欄（ed）は、プレビューの中の文字の要素で、プレビューを描き直すたびに attach で差し替える。
// D は ed のある文書（プレビューの中の文書）。選んだ範囲・execCommand は D で扱う
//   opts.base()：ボックスの文字の大きさ（pt）
//   opts.onBaseSize(n) / onBoxColor(c) / onBoxToggle('bold' | 'italic')：文字を選んでいないときに、ボックス全体を変える
function textCtl(opts, onCommit) {
  const baseNow = () => opts.base() || DEFAULT_SIZE;
  const bar = h('div', 'rich-bar');
  let ed = null;
  let D = document;
  let last = null;
  let timer = null;
  const commit = () => {
    clearTimeout(timer);
    if (!ed) return;
    const v = serializeRich(ed);
    if (v === last) return;
    last = v;
    onCommit(v);
  };
  // 書式のボタン（PowerPoint のホームの「フォント」「段落」に似せる）
  // ボタンを押しても、テキストの欄の選択を保つ（mousedown で preventDefault）。▾ のメニューを開くときは選択を覚えておき、付けるときに戻す
  let saved = null;
  const saveSel = () => {
    if (!ed) return;
    const sel = D.getSelection();
    if (sel.rangeCount && ed.contains(sel.getRangeAt(0).commonAncestorContainer)) saved = sel.getRangeAt(0).cloneRange();
  };
  const restoreSel = () => {
    if (!saved || !ed) return false;
    const sel = D.getSelection();
    sel.removeAllRanges();
    sel.addRange(saved);
    return true;
  };
  const iconBtn = (cls, html, title, fn) => {
    const b = h('button', `rich-btn ${cls}`);
    b.type = 'button';
    b.innerHTML = html;
    b.title = title;
    b.onmousedown = (e) => {
      e.preventDefault();
      fn();
    };
    return b;
  };
  const exec = (cmd) => () => {
    if (!hasSel() && (cmd === 'bold' || cmd === 'italic') && opts.onBoxToggle) return opts.onBoxToggle(cmd);
    if (!ed) return;
    restoreSel();
    D.execCommand('styleWithCSS', false, false);
    D.execCommand(cmd);
    commit();
  };
  // ▾ のメニュー。開くと、テキストの欄の選択を覚えておく
  const menus = [];
  const dropMenu = (title, build) => {
    const wrapM = h('span', 'rich-drop');
    const menu = h('div', 'rich-menu');
    menu.hidden = true;
    const close = () => (menu.hidden = true);
    const arrow = iconBtn('rich-arrow', '▾', title, () => {
      const open = menu.hidden;
      menus.forEach((m) => (m.hidden = true));
      if (open) {
        saveSel();
        menu.innerHTML = '';
        build(menu, close);
        menu.hidden = false;
        fitMenu(menu);
      }
    });
    menus.push(menu);
    wrapM.append(arrow, menu);
    return wrapM;
  };
  const sep = () => h('span', 'rich-sep');

  // 太字・斜体・下線・上付き・下付き
  // 文字の大きさの欄（PowerPoint のフォントサイズの欄）。文字を選んでいれば、その部分の大きさを出して、その部分だけを変える。
  // 選んでいなければ（カーソルだけ・欄の外）、ボックス全体の文字の大きさを変える
  const sizeIn = document.createElement('input');
  sizeIn.type = 'number';
  sizeIn.min = '1';
  sizeIn.step = '1';
  sizeIn.className = 'rich-size';
  sizeIn.title = '文字の大きさ（pt）。文字を選んでいれば、その部分だけを変える';
  sizeIn.value = String(baseNow());
  sizeIn.onchange = () => {
    const n = +sizeIn.value;
    if (!(n > 0)) return;
    if (hasSel()) {
      restoreSel();
      applySize(n);
      saveSel();
      sizeIn.focus(); // 続けて数字を変えられるように、大きさの欄にカーソルを戻す
    } else if (opts.onBaseSize) opts.onBaseSize(n);
  };
  // 範囲 r の始まりの文字の大きさ
  const sizeAt = (r) => {
    const n = r?.startContainer;
    if (!n || !ed || !ed.contains(n)) return baseNow();
    const sp = (n.nodeType === 1 ? n : n.parentElement)?.closest('span.fs');
    return sp && ed.contains(sp) ? +sp.dataset.size : baseNow();
  };
  // テキストの欄で選んだ範囲を覚えておき（大きさの欄を押すと欄の選択が外れるため）、大きさの欄にその大きさを出す
  const onSel = () => {
    if (!bar.isConnected || !ed) return;
    const sel = D.getSelection();
    if (!sel.rangeCount || !ed.contains(sel.getRangeAt(0).startContainer)) return;
    saved = sel.getRangeAt(0).cloneRange();
    if (document.activeElement !== sizeIn) sizeIn.value = String(sizeAt(saved));
    tellSel();
    // パネルをカーソルの右に置いているときは、カーソルが動いたら追いかける
    if (S.support?.atCaret) requestAnimationFrame(positionSupport);
  };
  // 範囲 r の最初の文字に付いている一部の色（"#RRGGBB"）。付いていなければ null
  const colorAt = (r) => {
    if (!r) return null;
    if (!ed) return null;
    const tw = D.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
    let t = null;
    for (let n = tw.nextNode(); n; n = tw.nextNode()) {
      if (!n.nodeValue.replace(/​/g, '') || !r.intersectsNode(n)) continue;
      if (n === r.startContainer && r.startOffset >= n.nodeValue.length && !r.collapsed) continue; // 選んだ範囲が、この文字の後ろから始まる
      if (n === r.endContainer && r.endOffset === 0 && !r.collapsed) continue;
      t = n;
      break;
    }
    for (let e = t?.parentElement; e && e !== ed; e = e.parentElement) {
      const raw = e.nodeName === 'FONT' ? e.getAttribute('color') : e.style?.color;
      const x = raw ? toHex6(raw) : null;
      if (x) return x === NO_COLOR.slice(1).toUpperCase() ? null : `#${x}`;
    }
    return null;
  };
  // 選んだ範囲が変わったことを、設定欄の「色」に知らせる（選んだ部分の色を出すため）
  // A の帯に、いまの色（選んだ部分の色。選んでいなければボックス全体の文字の色）を出す
  const shownColor = () => (hasSel() && colorAt(saved)) || normColor(opts.boxColor?.()) || '#000000';
  const tellSel = () => {
    const b = bar.querySelector('.rich-colorA .bar');
    if (b) b.style.background = shownColor();
  };
  // 文字を選んでいるか（カーソルだけなら false）
  const hasSel = () => !!(ed && saved && !saved.collapsed && ed.contains(saved.commonAncestorContainer));
  const grpSize = h('span', 'rich-group');
  grpSize.append(sizeIn, h('span', 'rich-unit', 'pt'));

  const grpFont = h('span', 'rich-group');
  grpFont.append(
    iconBtn('rich-bold', 'B', '太字', exec('bold')),
    iconBtn('rich-italic', 'I', '斜体', exec('italic')),
    // U：選んだ部分に一重線を付ける。選んだ部分がすでに全部下線なら外す
    iconBtn('rich-under', 'U', '下線（もう一度押すと外す）', () => applyUnderline(allUnderlined() ? null : { type: 'single' })),
    dropMenu('下線の種類', (menu, close) => {
      const pickU = (u) => {
        restoreSel();
        applyUnderline(u);
        close();
      };
      const item = (html, u) => menu.append(iconBtn('rich-item', html, '', () => pickU(u)));
      item('<u>一重線</u>', { type: 'single' });
      item('<u class="ud">二重線</u>', { type: 'double' });
      // 画像の線：太さを決めてから、Finder で画像を選ぶ。一度選んだ画像（とこの文書で使っている画像の線）は、次から候補に出す
      menu.append(h('div', 'rich-menu-label', '画像の線'));
      const wRow = h('div', 'rich-menu-row');
      const w = numEl(S.lastUnderW ?? 2, '', (v) => (S.lastUnderW = +v > 0 ? +v : 2), 0.5);
      w.min = '0.5';
      w.className = 'rich-w';
      w.oninput = () => (S.lastUnderW = +w.value > 0 ? +w.value : 2);
      wRow.append(h('span', null, '太さ'), w, h('span', null, 'pt'));
      menu.append(wRow);
      menu.append(iconBtn('rich-item', 'Finder から選ぶ…', '', () => {
        const keep = saved;
        close();
        pickImageFile().then((path) => {
          if (!path) return;
          addUnderImg(path);
          saved = keep;
          pickU({ type: 'img', img: path, w: S.lastUnderW ?? 2 });
        });
      }));
      const cands = underImgs();
      if (cands.length) {
        const list = h('div', 'rich-imgs');
        menu.append(list);
        Promise.all(cands.map(loadImage)).then(() => {
          for (const a of cands) {
            const b = iconBtn('rich-item rich-img', '', a, () => pickU({ type: 'img', img: a, w: S.lastUnderW ?? 2 }));
            const line = h('span', 'line');
            line.style.backgroundImage = `url('${S.imageURLs.get(a) || ''}')`;
            b.append(line, h('span', 'name', a.split('/').pop()));
            list.append(b);
          }
        });
      }
    }),
    iconBtn('rich-sup', '<span>x<sup>2</sup></span>', '上付き', exec('superscript')),
    iconBtn('rich-sub', '<span>x<sub>2</sub></span>', '下付き', exec('subscript')),
  );

  // 文字の色：A の帯は、いまの色（選んだ部分の色。選んでいなければボックス全体の文字の色）。
  // A か ▾ で色のメニューを開き、選んだ色を、文字を選んでいればその部分、選んでいなければボックス全体に付ける
  const colorA = iconBtn('rich-colorA', `<span class="a">A</span><span class="bar" style="background:${normColor(opts.boxColor?.()) || '#000000'}"></span>`, '文字の色', () => colorDrop.querySelector('.rich-arrow').onmousedown(new MouseEvent('mousedown')));
  const grpColor = h('span', 'rich-group');
  const colorDrop = dropMenu('文字の色', (menu, close) => {
    const sel = hasSel();
    const use = (c) => {
      close();
      if (!sel) return opts.onBoxColor?.(c);
      restoreSel();
      applyColor(c);
      tellSel();
    };
    // このスライドで使っている色
    const used = usedColors();
    if (used.length) {
      menu.append(h('div', 'rich-menu-label', 'このスライドで使っている色'));
      const sw = h('div', 'rich-swatches');
      for (const c of used) {
        const b = iconBtn('rich-swatch', '', c, () => use(c));
        b.style.background = c;
        sw.append(b);
      }
      menu.append(sw);
    }
    const pick = document.createElement('input');
    pick.type = 'color';
    pick.className = 'rich-color';
    pick.value = shownColor();
    pick.onchange = () => use(pick.value);
    const row = h('div', 'rich-menu-row');
    row.append(h('span', null, 'ほかの色'), pick);
    menu.append(row, iconBtn('rich-item', '色を外す', '', () => {
      close();
      if (!sel) return opts.onBoxColor?.(undefined);
      restoreSel();
      applyColor(NO_COLOR);
      tellSel();
    }));
  });
  grpColor.append(colorA, colorDrop);

  // 箇条書き：カーソルのある段落（選んだ段落）の行頭に「•」を付ける。もう一度押すと外す
  const bulletIcon = '<svg viewBox="0 0 20 16" width="20" height="16"><g fill="currentColor"><circle cx="2.5" cy="3" r="1.6"/><circle cx="2.5" cy="8" r="1.6"/><circle cx="2.5" cy="13" r="1.6"/></g><g stroke="currentColor" stroke-width="1.6"><path d="M7 3h12M7 8h12M7 13h12"/></g></svg>';
  const grpPara = h('span', 'rich-group');
  grpPara.append(
    iconBtn('rich-bullet', bulletIcon, '箇条書き（もう一度押すと外す）', () => toggleBullet()),
    dropMenu('箇条書きの種類', (menu, close) => {
      const item = (html, type) => menu.append(iconBtn('rich-item', html, '', () => {
        restoreSel();
        setBullet(type);
        close();
      }));
      item('• 通常', 'normal');
      item('不可視', 'hidden');
    }),
  );

  bar.append(grpFont, sep(), grpColor, sep(), grpSize, grpPara);
  // メニューの外を押したら閉じる（欄が画面から消えたら、この見張りも外す）
  const outside = (e) => {
    if (!bar.isConnected) return document.removeEventListener('mousedown', outside);
    if (!e.target.closest?.('.rich-drop, .rich-colorA')) menus.forEach((m) => (m.hidden = true));
  };
  document.addEventListener('mousedown', outside);

  const applyColor = (value) => {
    if (!ed) return;
    D.execCommand('styleWithCSS', false, true);
    D.execCommand('foreColor', false, value);
    D.execCommand('styleWithCSS', false, false);
    commit();
  };

  // 選んだ範囲の段落（ed の直下の div）
  const blocksIn = (r) => {
    const out = [];
    for (const d of ed.children) if (r.intersectsNode(d)) out.push(d);
    if (!out.length) {
      // カーソルが段落の外（ed の直下の文字）にあるときは、段落に包む
      let n = r.startContainer;
      while (n && n.parentNode !== ed) n = n.parentNode;
      if (n && n !== ed) {
        const d = D.createElement('div');
        n.replaceWith(d);
        d.append(n);
        out.push(d);
      }
    }
    return out;
  };

  // 選んだ部分の文字が、全部下線の中にあるか
  const allUnderlined = () => {
    if (!ed) return false;
    const sel = D.getSelection();
    if (!sel.rangeCount) return false;
    const r = sel.getRangeAt(0);
    if (r.collapsed || !ed.contains(r.commonAncestorContainer)) return false;
    const w = D.createTreeWalker(r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentNode, NodeFilter.SHOW_TEXT);
    let any = false;
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!r.intersectsNode(n) || !n.nodeValue.replace(/\u200b/g, '')) continue;
      any = true;
      if (!n.parentElement.closest('u')) return false;
    }
    return any;
  };

  // 選んだ範囲を、make() で作る要素で包む（make が null なら外す）。sel は同じ種類の要素（下線・大きさ）を見分ける。
  // 選んだ範囲の中にある同じ種類の要素は外し、外側にある同じ種類の要素からは出す（一部を選んだときは、その前後に残す）。
  // keep なら、包んだあとも同じ文字を選んだままにする
  const wrapSelection = (selector, make, keep) => {
    if (!ed) return false;
    const sel = D.getSelection();
    if (!sel.rangeCount) return false;
    const range = sel.getRangeAt(0);
    if (range.collapsed || !ed.contains(range.commonAncestorContainer)) return false;
    const els = [];
    for (const d of blocksIn(range)) {
      const r = D.createRange();
      r.selectNodeContents(d);
      if (d.contains(range.startContainer)) r.setStart(range.startContainer, range.startOffset);
      if (d.contains(range.endContainer)) r.setEnd(range.endContainer, range.endOffset);
      if (r.collapsed) continue;
      const frag = r.extractContents();
      for (const x of frag.querySelectorAll(selector)) x.replaceWith(...x.childNodes);
      const el = (make && make()) || D.createElement('span');
      el.append(frag);
      r.insertNode(el);
      for (let o = el.parentElement?.closest(selector); o && ed.contains(o); o = el.parentElement?.closest(selector)) {
        const rb = D.createRange();
        rb.setStart(o, 0);
        rb.setEndBefore(el);
        const before = o.cloneNode(false);
        before.append(rb.extractContents());
        const ra = D.createRange();
        ra.setStartAfter(el);
        ra.setEnd(o, o.childNodes.length);
        const after = o.cloneNode(false);
        after.append(ra.extractContents());
        o.replaceWith(before, ...o.childNodes, after);
      }
      els.push(el);
    }
    // 空になった要素を消す
    for (const x of ed.querySelectorAll(selector)) if (!x.textContent) x.remove();
    const nr = D.createRange();
    if (keep && els.length) {
      nr.setStartBefore(els[0]);
      nr.setEndAfter(els[els.length - 1]);
    }
    if (!make) {
      // 外したときは、包んでいた仮の要素をほどく（選択はその前後に置く）
      for (const el of els) el.replaceWith(...el.childNodes);
    }
    sel.removeAllRanges();
    if (keep && els.length && make) sel.addRange(nr);
    commit();
    return true;
  };

  // 下線を付ける（u が null なら外す）。選んだ範囲の、すでにある下線は付け替える
  const applyUnderline = (u) => {
    const ok = wrapSelection('u', u ? () => {
      const tmp = D.createElement('div');
      tmp.innerHTML = underlineHTML(u, '', (src) => S.imageURLs.get(src) || '');
      return tmp.firstChild;
    } : null);
    if (!ok) status('下線を付ける文字を選んでください', true);
  };

  // 一部の文字の大きさ（pt）。ボックスの文字の大きさと同じなら、大きさの指定を外す
  const applySize = (n) => {
    const ok = wrapSelection('span.fs', n === baseNow() ? null : () => {
      const sp = D.createElement('span');
      sp.className = 'fs';
      sp.dataset.size = String(n);
      sp.style.fontSize = `${n}pt`;
      return sp;
    }, true);
    if (!ok) status('大きさを変える文字を選んでください', true);
  };

  // 箇条書きにする段落：カーソルのある段落（選んだ段落）。欄にカーソルが無ければ全部の段落
  const bulletTargets = () => {
    if (!ed) return [];
    const sel = D.getSelection();
    if (!sel.rangeCount || !ed.contains(sel.getRangeAt(0).commonAncestorContainer)) return [...ed.children];
    return blocksIn(sel.getRangeAt(0));
  };
  // ボタン：通常の箇条書きを付ける。全部が箇条書きなら外す
  const toggleBullet = () => {
    const ds = bulletTargets();
    const on = !ds.every((d) => d.classList.contains('li'));
    ds.forEach((d) => {
      d.classList.toggle('li', on);
      d.classList.remove('lih');
    });
    commit();
  };
  // ▾ のメニュー：通常か不可視にする
  const setBullet = (type) => {
    for (const d of bulletTargets()) {
      d.classList.add('li');
      d.classList.toggle('lih', type === 'hidden');
    }
    commit();
  };
  // Shift+Enter：段落の中の改行（<br class="soft">）を差し込み、カーソルをその直後に置く
  const insertSoftBreak = () => {
    const sel = D.getSelection();
    if (!sel.rangeCount) return;
    const r = sel.getRangeAt(0);
    r.deleteContents();
    const br = D.createElement('br');
    br.className = 'soft';
    const z = D.createTextNode('​'); // 行末の改行を見えるようにするゼロ幅の文字（書き出すときに外す）
    r.insertNode(z);
    r.insertNode(br);
    const nr = D.createRange();
    nr.setStart(z, 1);
    nr.collapse(true);
    sel.removeAllRanges();
    sel.addRange(nr);
    ed.dispatchEvent(new Event('input'));
  };
  // 書く欄を差し替える（プレビューを描き直したとき）。差し替えた時点の中身を、書き込み済みの中身とする
  const attach = (el) => {
    if (D && D !== el.ownerDocument) D.removeEventListener('selectionchange', onSel);
    ed = el;
    D = el.ownerDocument;
    D.addEventListener('selectionchange', onSel);
    last = serializeRich(ed);
    saved = null;
    sizeIn.value = String(baseNow());
    ed.onfocus = () => D.execCommand('defaultParagraphSeparator', false, 'div');
    ed.onkeydown = (e) => {
      if (e.key === 'Enter' && e.shiftKey) {
        e.preventDefault();
        insertSoftBreak();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        commit();
        select(null); // 選択を外す（書くのをやめる）
      }
    };
    ed.oninput = () => {
      clearTimeout(timer);
      timer = setTimeout(commit, 600);
    };
    ed.onblur = commit;
    ed.onpaste = (e) => {
      e.preventDefault();
      D.execCommand('insertText', false, e.clipboardData.getData('text/plain'));
    };
  };
  const detach = () => {
    clearTimeout(timer);
    if (D) D.removeEventListener('selectionchange', onSel);
    if (ed) {
      ed.onkeydown = ed.oninput = ed.onblur = ed.onpaste = ed.onfocus = null;
    }
    ed = null;
  };
  return {
    bar,
    attach,
    detach,
    commit,
    ed: () => ed,
    getSaved: () => saved,
    setSaved: (r) => {
      saved = r;
      sizeIn.value = String(sizeAt(r));
      tellSel();
    },
    // ボックス全体の色・大きさが変わったとき、ボタンの表示を合わせる
    refresh: () => {
      if (document.activeElement !== sizeIn) sizeIn.value = String(sizeAt(saved));
      tellSel();
    },
  };
}

// ---------------------------------------------------------------- サポートパネル・プレビューに直接書く
// ボックスを 1 つ選ぶと、プレビューのボックスにそのまま文字を書ける（クリックした位置にカーソルを置く）。
// 選んでいるあいだ、ボックスに重ならない近くに、文字の書式のパネル（サポートパネル）を出す。ドラッグで動かせ、選択を外すと消える。
// 文字は、書いている途中は layout.yaml にだけ書き、プレビューは描き直さない（描き直すと書く欄が作り直されるため）。
// 書き終えたとき（選択を外したとき・別のボックスを選んだとき）に描き直し、文字の高さにボックスを合わせる

// 一度 Finder で選んだ画像の線の画像（次から候補に出す）。ブラウザに覚えておき、この文書で使っている画像の線も候補にする
function underImgs() {
  let saved = [];
  try {
    saved = JSON.parse(localStorage.getItem('slideCreator.underImgs') || '[]');
  } catch {
    saved = [];
  }
  const used = new Set();
  for (const f of Object.values(S.files)) {
    for (const m of String(f.text || '').matchAll(/<u\s[^>]*img\s*=\s*\\?"([^"\\]+)\\?"/gi)) used.add(m[1]);
  }
  return [...new Set([...saved.filter((p) => S.assets.includes(p) || used.has(p)), ...used])];
}
function addUnderImg(path) {
  try {
    const cur = JSON.parse(localStorage.getItem('slideCreator.underImgs') || '[]');
    localStorage.setItem('slideCreator.underImgs', JSON.stringify([path, ...cur.filter((p) => p !== path)].slice(0, 20)));
  } catch {
    // 覚えられなくても、この場では使える
  }
}

// 書く欄の中の範囲 ↔ 先頭から何文字目か（プレビューを描き直したあと、同じ位置にカーソルを戻すため）
function offsetsIn(ed) {
  const sel = ed.ownerDocument.getSelection();
  if (!sel.rangeCount || !ed.contains(sel.getRangeAt(0).startContainer)) return null;
  const r = sel.getRangeAt(0);
  const off = (node, o) => {
    const x = ed.ownerDocument.createRange();
    x.setStart(ed, 0);
    x.setEnd(node, o);
    return x.toString().length;
  };
  return { start: off(r.startContainer, r.startOffset), end: off(r.endContainer, r.endOffset) };
}
function rangeAt(ed, pos) {
  const D = ed.ownerDocument;
  const at = (n, isStart) => {
    const w = D.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
    let last = null;
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      if (isStart ? n < t.nodeValue.length : n <= t.nodeValue.length) return [t, n];
      n -= t.nodeValue.length;
      last = t;
    }
    return last ? [last, last.nodeValue.length] : [ed.lastElementChild || ed, 0];
  };
  const r = D.createRange();
  r.setStart(...at(pos.start, pos.start < pos.end));
  r.setEnd(...at(pos.end, false));
  return r;
}

// 選んだボックスの文字の要素を、書ける状態にする。how.point（プレビューの中の位置）か how.pos（何文字目か）にカーソルを置く
function startInline(how = {}) {
  const sp = S.support;
  if (!sp || S.layer !== 'main') return;
  const doc = $('#frame').contentDocument;
  const el = doc && [...doc.querySelectorAll('.txt')].find((x) => x.dataset.path === sp.path);
  const ed = el?.firstElementChild;
  if (!ed) return;
  if (S.inline?.ed === ed) return placeCaret(ed, how);
  ed.contentEditable = 'true';
  ed.spellcheck = false;
  ed.classList.add('ied');
  sp.ctl.attach(ed);
  S.inline = { path: sp.path, ed };
  placeCaret(ed, how);
}
function placeCaret(ed, how) {
  const D = ed.ownerDocument;
  let r = null;
  if (how.point) {
    r = D.caretRangeFromPoint(how.point.x, how.point.y);
    if (r && !ed.contains(r.startContainer)) r = null;
  }
  if (!r && how.pos) r = rangeAt(ed, how.pos);
  if (!r && (how.point || how.focus)) {
    r = D.createRange();
    r.selectNodeContents(ed.lastElementChild || ed);
    r.collapse(false);
  }
  if (how.focus === false && !how.point) {
    if (r) {
      D.getSelection().removeAllRanges();
      D.getSelection().addRange(r);
    }
    return;
  }
  const frame = $('#frame');
  if (document.activeElement !== frame) frame.focus({ preventScroll: true });
  ed.focus({ preventScroll: true });
  if (r) {
    D.getSelection().removeAllRanges();
    D.getSelection().addRange(r);
  }
  positionSupport(); // カーソルの位置に合わせてパネルを置く（ボックスが大きいとき）
}

// 書くのを終える。書いた文字を書き込み、プレビューを描き直して、文字の高さにボックスを合わせる
function finishInline() {
  const inl = S.inline;
  if (!inl) return;
  S.inline = null;
  const box = findBox(inl.path);
  if (box && (box.kind === 'text' || box.kind === 'empty')) S.pendingFit = [inl.path];
  const before = S.histI;
  S.support?.ctl.commit(); // S.inline を外してから書き込むので、プレビューも描き直す
  S.support?.ctl.detach();
  inl.ed.contentEditable = 'false';
  inl.ed.classList.remove('ied');
  if (S.histI === before && S.inlineDirty) renderPreview(); // 書き込みは済んでいて、描き直しだけが残っている
  S.inlineDirty = false;
}

// 選択に合わせて、サポートパネルを出す・消す・置き直す（applyView のたびに呼ぶ）
function syncSupport() {
  const one = S.layer === 'main' && S.selected ? findBox(S.selected) : null;
  if (!one || !one.src || !isLeaf(one)) return closeSupport();
  if (!S.support || S.support.path !== one.path || S.support.stale) buildSupport(one);
  positionSupport();
}
function closeSupport() {
  if (!S.support) return;
  S.support.ctl.detach();
  S.support.el.remove();
  S.support = null;
}

// パネルを、ボックスに重ならない近くに置く（下・上・右・左の順に、入る所）。ドラッグで動かしたあとは、その位置のまま
function positionSupport() {
  const sp = S.support;
  if (!sp || sp.moved) return;
  const box = findBox(sp.path);
  const frame = $('#frame');
  const win = frame.contentWindow;
  if (!box || !S.view || !win) return;
  const fr = frame.getBoundingClientRect();
  const { s, offX, offY } = S.view;
  const k = S.ppu * s;
  const bx = fr.left + offX + box.rect.x * k - win.scrollX;
  const by = fr.top + offY + box.rect.y * k - win.scrollY;
  const bw = box.rect.w * k;
  const bh = box.rect.h * k;
  const pw = sp.el.offsetWidth;
  const ph = sp.el.offsetHeight;
  const m = 10;
  const fits = (x, y) => x >= fr.left && y >= fr.top && x + pw <= fr.right && y + ph <= fr.bottom;
  const clampX = (x) => Math.max(fr.left + 4, Math.min(x, fr.right - pw - 4));
  const clampY = (y) => Math.max(fr.top + 4, Math.min(y, fr.bottom - ph - 4));
  // まず選んだボックスの上に置く。入らなければ下、右、左の順
  const cands = [
    [clampX(bx), by - ph - m],
    [clampX(bx), by + bh + m],
    [bx + bw + m, clampY(by)],
    [bx - pw - m, clampY(by)],
  ];
  let pos = cands.find(([cx, cy]) => fits(cx, cy));
  sp.atCaret = !pos;
  if (!pos) {
    // ボックスの外に入る場所が無い（ボックスが大きい）ときは、カーソルの右に置く。右に入らなければ、カーソルの行の下に置く
    const c = caretRect(fr);
    const cx = c ? c.right : bx;
    const cy = c ? c.top : by;
    const ch = c ? c.height : 0;
    if (cx + 24 + pw <= fr.right) pos = [cx + 24, clampY(cy + ch / 2 - 20)];
    else pos = [clampX(cx - 20), cy + ch + m + ph <= fr.bottom ? cy + ch + m : clampY(cy - ph - m)];
  }
  const [x, y] = pos;
  sp.el.style.left = `${Math.round(x)}px`;
  sp.el.style.top = `${Math.round(y)}px`;
}

// 書いている文字のカーソルの位置（画面の座標）。書いていなければ null
function caretRect(fr) {
  const ed = S.inline?.ed;
  if (!ed) return null;
  const sel = ed.ownerDocument.getSelection();
  let r = null;
  if (sel.rangeCount && ed.contains(sel.getRangeAt(0).endContainer)) {
    const rg = sel.getRangeAt(0).cloneRange();
    rg.collapse(false);
    r = rg.getClientRects()[0] || (rg.endContainer.nodeType === 1 ? rg.endContainer.getBoundingClientRect() : rg.endContainer.parentElement?.getBoundingClientRect());
  }
  r ||= ed.getBoundingClientRect();
  return { right: fr.left + r.right, top: fr.top + r.top, height: r.height };
}

function buildSupport(box) {
  closeSupport();
  const path = box.path;
  const cur = () => findBox(path);
  const node = () => cur()?.node || {};
  const inh = () => cur()?.inherited || {};
  const font = () => ({ ...(inh().font || {}), ...(node().font || {}) });
  // 書式を変えると文字の高さが変わりうるので、描き直したあとにボックスを合わせる（文字のボックスだけ）
  const set = (k, v) => {
    const b = cur();
    if (!b) return;
    if (b.kind === 'text') S.pendingFit = [b.path];
    return editOne(b, k, v);
  };
  const el = h('div', 'support');
  const head = h('div', 'support-head');
  head.append(h('span', 'grip', '⋮⋮'), h('span', 'support-title', '書式'));
  // 文字の書式の道具。文字を選んでいなければ、ボックス全体の大きさ・色・太字・斜体を変える
  const ctl = textCtl({
    base: () => font().size ?? DEFAULT_SIZE,
    boxColor: () => font().color,
    onBaseSize: (n) => set(['font', 'size'], n),
    onBoxColor: (c) => set(['font', 'color'], c || undefined),
    onBoxToggle: (k) => set(['font', k], font()[k] ? undefined : true),
  }, (v) => {
    const b = cur();
    if (!b) return;
    // 書いている途中は、プレビューを描き直さない（書く欄が作り直されるため）。描き直しは書き終えたときにする
    if (S.inline) {
      S.skipPreview = true;
      S.inlineDirty = true;
    }
    const empty = !v.replace(/<[^>]*>/g, '').trim() && !/<li/.test(v);
    const done = editOne(b, 'text', empty && b.kind !== 'text' ? undefined : v);
    // 書いている途中でも（手を止めて書き込んだとき）、ボックスを文字の高さに合わせ、下に余った部分を別のボックスにする
    Promise.resolve(done).then(() => {
      const inl = S.inline;
      if (inl && inl.path === b.path && findBox(b.path)?.kind === 'text') fitTextBelow(b.path, inl.ed.ownerDocument);
    });
  });
  // 光彩（文字の周りをぼかした色で縁取る）。押すと付ける・外す
  const glowBtn = h('button', 'rich-btn rich-glow');
  glowBtn.type = 'button';
  glowBtn.innerHTML = '<span>A</span>';
  glowBtn.title = '光彩（もう一度押すと外す）';
  glowBtn.classList.toggle('on', !!font().glow);
  glowBtn.onmousedown = (e) => {
    e.preventDefault();
    set(['font', 'glow'], font().glow ? undefined : true);
  };
  // 画像：▾ で、Finder から選ぶか、フォルダの画像を選ぶ。画像は文字の裏に置く
  const imgWrap = h('span', 'rich-drop');
  const imgBtn = h('button', 'rich-btn rich-image');
  imgBtn.type = 'button';
  imgBtn.title = '画像（文字の裏に置く）';
  imgBtn.innerHTML = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="5" width="18" height="14" rx="1.5"/><circle cx="9" cy="10" r="1.8"/><path d="M4 18l5-5 3 3 3-4 5 6"/></svg><span class="arr">▾</span>';
  const imgMenu = h('div', 'rich-menu');
  imgMenu.hidden = true;
  const setImage = (p) => {
    imgMenu.hidden = true;
    const b = cur();
    if (!b) return;
    const pairs = [['image', p]];
    if (node().crop != null) pairs.push(['crop', undefined]);
    editMany(b, pairs).then(() => p && status(S.pickMsg || ''));
  };
  imgBtn.onmousedown = (e) => {
    e.preventDefault();
    if (!imgMenu.hidden) return (imgMenu.hidden = true);
    imgMenu.innerHTML = '';
    const item = (html, fn) => {
      const b = h('button', 'rich-btn rich-item');
      b.type = 'button';
      b.innerHTML = html;
      b.onmousedown = (ev) => {
        ev.preventDefault();
        fn();
      };
      imgMenu.append(b);
      return b;
    };
    item('Finder から選ぶ…', () => {
      imgMenu.hidden = true;
      pickImageFile().then((p) => p && setImage(p));
    });
    if (node().image != null) item('画像を外す', () => setImage(undefined));
    const list = h('div', 'rich-imgs');
    imgMenu.append(list);
    const imgs = [...new Set(S.assets)];
    // HEIC は、JPEG に変えたものがまだ無いものだけ出す（選ぶと JPEG に変えて使う）
    const heics = (S.heics || []).filter((x) => !S.assets.includes(heicJpegPath(x)));
    if (!imgs.length && !heics.length) list.append(h('div', 'note', 'フォルダに画像がありません'));
    for (const a of heics) {
      const b = h('button', 'rich-btn rich-item rich-thumb');
      b.type = 'button';
      b.title = `${a}（選ぶと JPEG に変えて使います）`;
      b.append(h('span', 'heic-mark', 'HEIC'), h('span', 'name', a.split('/').pop()));
      b.onmousedown = (ev) => {
        ev.preventDefault();
        imgMenu.hidden = true;
        useHeic(a).then(setImage, (e) => status(`HEIC の写真を読めません: ${e.message || e}`, true));
      };
      list.append(b);
    }
    Promise.all(imgs.map(loadImage)).then(() => {
      for (const a of imgs) {
        if (!S.imageURLs.get(a)) continue;
        const b = h('button', 'rich-btn rich-item rich-thumb');
        b.type = 'button';
        b.title = a;
        const im = document.createElement('img');
        im.src = S.imageURLs.get(a);
        b.append(im, h('span', 'name', a.split('/').pop()));
        b.onmousedown = (ev) => {
          ev.preventDefault();
          setImage(a);
        };
        list.append(b);
      }
      fitMenu(imgMenu);
    });
    imgMenu.hidden = false;
    fitMenu(imgMenu);
  };
  imgWrap.append(imgBtn, imgMenu);
  ctl.bar.append(h('span', 'rich-sep'), glowBtn, imgWrap);
  document.addEventListener('mousedown', function outside(e) {
    if (!el.isConnected) return document.removeEventListener('mousedown', outside);
    if (!e.target.closest?.('.rich-drop')) imgMenu.hidden = true;
  });

  // そのほかの文字の設定（左の欄から移した）。開くかどうかは覚えておく
  const more = document.createElement('details');
  more.className = 'support-more';
  try {
    more.open = localStorage.getItem('slideCreator.supportMore') === '1';
  } catch {
    more.open = false;
  }
  more.ontoggle = () => {
    try {
      localStorage.setItem('slideCreator.supportMore', more.open ? '1' : '0');
    } catch {
      // 覚えられなくてもよい
    }
    positionSupport();
  };
  more.append(h('summary', null, 'そのほかの文字の設定'));
  const b0 = cur();
  const n0 = node();
  const f0 = font();
  const i0 = inh();
  more.append(
    row('フォント（英語）', selectEl(FONT_OPTS, f0.family ?? DEFAULT_FONT, (v) => set(['font', 'family'], v))),
    row('フォント（日本語）', selectEl(JA_FONT_OPTS, f0.family_ja ?? DEFAULT_FONT_JA, (v) => set(['font', 'family_ja'], v))),
    fullRow(pairLine(
      ['揃え方（横）', selectEl(H_OPTS, n0.align ?? i0.align ?? 'left', (v) => set('align', v))],
      ['揃え方（縦）', selectEl(V_OPTS, n0.valign ?? i0.valign ?? 'top', (v) => set('valign', v))],
    )),
    fullRow(pairLine(
      ['行間', numEl(n0.line_height ?? i0.line_height ?? 1.2, '', (v) => +v > 0 && set('line_height', +v), 0.05), '倍'],
      ['段落の間隔', numEl(n0.para_space ?? i0.para_space ?? 0, '', (v) => set('para_space', +v > 0 ? +v : undefined)), 'pt'],
    )),
    ...textBgRows(b0, n0, (b, k, v) => set(k, v)),
    row('光彩の色', h2(colorEl(typeof f0.glow === 'object' ? f0.glow.color : f0.glow ? '#ffffff' : '', (v) => set(['font', 'glow'], v ? { ...(typeof f0.glow === 'object' ? f0.glow : {}), color: v } : undefined)))),
    row('見出しアイコン', h2(
      selectEl([['', 'なし'], ...[...new Set([...S.assets, ...(n0.icon ? [String(n0.icon)] : [])])].map((a) => [a, a.split('/').pop()])], n0.icon ?? '', (v) => set('icon', v || undefined)),
      pickBtn((p) => set('icon', p)),
    )),
    ...(n0.icon ? [row('アイコンの大きさ', h2(numEl(n0.icon_size ?? f0.size ?? DEFAULT_SIZE, '', (v) => +v > 0 && set('icon_size', +v)), 'pt'))] : []),
    row('', checkEl('テキストの下に生じた余白を別のボックスにする', n0.split_below !== false, (x) => {
      const b = cur();
      if (!b) return;
      if (x) {
        S.pendingFit = [b.path];
        return editOne(b, 'split_below', undefined);
      }
      return unfitTextBelow(b);
    })),
  );
  el.append(head, ctl.bar, more);
  document.body.append(el);
  S.support = { path, el, ctl, moved: false };

  // ドラッグで動かす（つまみを持って動かす）
  head.onpointerdown = (e) => {
    e.preventDefault();
    head.setPointerCapture(e.pointerId);
    const r = el.getBoundingClientRect();
    const dx = e.clientX - r.left;
    const dy = e.clientY - r.top;
    head.onpointermove = (ev) => {
      S.support.moved = true;
      el.style.left = `${Math.round(ev.clientX - dx)}px`;
      el.style.top = `${Math.round(ev.clientY - dy)}px`;
    };
    head.onpointerup = () => {
      head.onpointermove = head.onpointerup = null;
    };
  };
  queueMicrotask(fitSupportLabels);
}

// パネルの中の名前の列の幅を、いちばん長い名前に合わせる
function fitSupportLabels() {
  const more = S.support?.el.querySelector('.support-more');
  if (!more) return;
  let w = 0;
  for (const l of more.querySelectorAll('.row > label')) w = Math.max(w, l.scrollWidth);
  more.style.setProperty('--label-w', `${Math.ceil(w) + 8}px`);
}

// ---------------------------------------------------------------- 分割（マス目）
// 分割ボタンを押すと、下にマス目を出す。マウスを乗せた範囲の「行 × 列」に分ける（PowerPoint の表の挿入と同じ）
const SPLIT_MAX = { rows: 8, cols: 10 };
function openSplitGrid() {
  const btnEl = $('#btn-split');
  let pop = $('#split-grid');
  if (pop && !pop.hidden) return (pop.hidden = true);
  const box = S.selected ? findBox(S.selected) : null;
  if (!box || !box.src) return;
  if (!pop) {
    pop = h('div', 'split-grid');
    pop.id = 'split-grid';
    document.body.append(pop);
    document.addEventListener('mousedown', (e) => {
      if (!pop.hidden && !e.target.closest?.('#split-grid, #btn-split')) pop.hidden = true;
    });
  }
  pop.innerHTML = '';
  const label = h('div', 'split-label', '分割');
  const grid = h('div', 'split-cells');
  grid.style.gridTemplateColumns = `repeat(${SPLIT_MAX.cols}, 18px)`;
  const cells = [];
  const show = (r, c) => {
    label.textContent = r && c ? `${r} 行 × ${c} 列` : '分割';
    for (const x of cells) x.el.classList.toggle('on', x.r <= r && x.c <= c);
  };
  for (let r = 1; r <= SPLIT_MAX.rows; r++) {
    for (let c = 1; c <= SPLIT_MAX.cols; c++) {
      const cell = h('span', 'split-cell');
      cell.onmouseenter = () => show(r, c);
      cell.onmousedown = (e) => {
        e.preventDefault();
        pop.hidden = true;
        if (r * c > 1) splitGrid(findBox(S.selected) || box, r, c);
      };
      cells.push({ el: cell, r, c });
      grid.append(cell);
    }
  }
  grid.onmouseleave = () => show(0, 0);
  pop.append(label, grid);
  const b = btnEl.getBoundingClientRect();
  pop.style.left = `${Math.round(b.left)}px`;
  pop.style.top = `${Math.round(b.bottom + 4)}px`;
  pop.hidden = false;
}

// r 行 × c 列に分ける。1 行か 1 列なら、その向きだけに分ける。両方なら、上下に r 段に分け、各段を左右に c 個に分ける
function splitGrid(box, r, c) {
  if (!box?.src) return;
  setSel([]);
  if (r === 1) return splitBox(box, 'cols', c);
  if (c === 1) return splitBox(box, 'rows', r);
  return splitBox(box, 'rows', r, c);
}

// その <br> が ↵ の直後にあるか（間の空の文字の節は読み飛ばす）
function afterMark(br) {
  let p = br.previousSibling;
  while (p && p.nodeType === 3 && !p.nodeValue.replace(/\u200b/g, '')) p = p.previousSibling;
  return !!p?.classList?.contains('sb');
}

// 「色を外す」で一時的に付ける色（書き出すときに色の印ごと外す）
const NO_COLOR = '#010203';
// "rgb(r, g, b)" や "#rgb" を "RRGGBB" にする
function toHex6(c) {
  const m = String(c).match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (m) return [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
  const s = String(c).replace(/^#/, '');
  if (/^[0-9a-f]{3}$/i.test(s)) return s.split('').map((x) => x + x).join('').toUpperCase();
  return /^[0-9a-f]{6}$/i.test(s) ? s.toUpperCase() : null;
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
          // ↵ の直後の <br> と、<br class="soft">（プレビューに直接書くときの段落の中の改行）だけが段落の中の改行。
          // ほかの <br> は空の行を保つための置き物
          if (afterMark(c) || c.classList?.contains('soft')) out += '<br>';
          continue;
        }
        const t = TAGS[c.nodeName];
        // 一部の文字の大きさ：<span class="fs" data-size>（色も付いていれば、その色も書く）
        if (c.nodeName === 'SPAN' && c.classList.contains('fs')) {
          const cc = c.style?.color ? toHex6(c.style.color) : null;
          const ct = cc && cc !== NO_COLOR.slice(1).toUpperCase() ? cc : null;
          out += `${ct ? `<color #${ct}>` : ''}<size ${+c.dataset.size}>${inline(c)}</size>${ct ? '</color>' : ''}`;
          continue;
        }
        // 下線：<u>（一重線）、<u class="ud">（二重線）、<u class="ui" data-img data-w>（画像の線）
        if (c.nodeName === 'U') {
          const open = c.classList.contains('ui') ? `<u img="${c.dataset.img}" w="${c.dataset.w || 2}">` : c.classList.contains('ud') ? '<u double>' : '<u>';
          out += `${open}${inline(c)}</u>`;
          continue;
        }
        // 色：<span style="color:…">（入力欄で付けたもの）か <font color="…">
        const raw = c.nodeName === 'FONT' ? c.getAttribute('color') : c.style?.color;
        const col = raw ? toHex6(raw) : null;
        const colTag = col && col !== NO_COLOR.slice(1).toUpperCase() ? col : null;
        if (colTag) out += `<color #${colTag}>`;
        if (t) out += `<${t}>`;
        out += inline(c);
        if (t) out += `</${t}>`;
        if (colTag) out += '</color>';
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
      paras.push((c.classList.contains('li') ? (c.classList.contains('lih') ? '<li hidden>' : '<li>') : '') + inline(c));
    } else {
      pending = (pending ?? '') + inline({ childNodes: [c] });
    }
  }
  if (pending !== null) paras.push(pending);
  return paras.join('\n')
    .replace(/<(b|i|sup|sub)><\/\1>/g, '')
    .replace(/<u[^>]*><\/u>/g, '')
    .replace(/<size [\d.]+><\/size>/g, '')
    .replace(/<color #[0-9A-F]{6}><\/color>/g, '')
    // 隣り合う同じ色はまとめる
    .replace(/<color (#[0-9A-F]{6})>((?:(?!<\/?color)[\s\S])*)<\/color><color \1>/g, '<color $1>$2');
}

// 背景の欄。値は "#RRGGBB"（色）か { image, mode: stretch | cover, color }
// 画像の背景は、既定では縦横比を変えてボックスいっぱいに広げる
const BG_OPTS = [['', 'なし'], ['color', '色'], ['image', '画像']];
const BG_MODE_OPTS = [['stretch', '縦横比を変えて、ボックスいっぱいに広げる'], ['cover', '縦横比を保って広げ、はみ出た部分を切る']];
// opts.rect：背景を敷く枠（トリミングの画面で、枠の縦横比に合わせるため）
function bgRows(cur, write, opts = {}) {
  const kind = !cur ? '' : typeof cur !== 'object' ? 'color' : cur.image ? 'image' : cur.color ? 'color' : '';
  const rows = [
    row('背景', selectEl(BG_OPTS, kind, (v) => {
      if (v === '') return write(undefined);
      if (v === 'color') return write(typeof cur === 'string' ? cur : cur?.color || '#ffffff');
      const keep = typeof cur === 'object' && cur ? cur : {};
      const image = (typeof cur === 'object' && cur?.image) || S.assets[0];
      if (image) return write({ ...keep, image });
      pickImageFile().then((path) => (path ? write({ ...keep, image: path }) : renderProps()));
    })),
  ];
  if (kind === 'color') rows.push(row('色', colorEl(typeof cur === 'string' ? cur : cur.color, (v) => write(v || undefined))));
  if (kind === 'image') {
    const opts = [...new Set([...S.assets, cur.image])].map((a) => [a, a.split('/').pop()]);
    rows.push(row('画像', h2(selectEl(opts, cur.image, (v) => write({ ...cur, image: v })), pickBtn((path) => write({ ...cur, image: path })))));
    rows.push(row('広げ方', selectEl(BG_MODE_OPTS, cur.mode === 'cover' ? 'cover' : 'stretch', (v) => {
      const { mode: _m, tile: _t, ...rest } = cur;
      write(v === 'cover' ? { ...rest, mode: v } : rest);
    })));
    // トリミング：画像のボックスと同じく、残す範囲の左端・上端・右端・下端（元画像に対する割合）
    const base = plain(cur);
    const crop = Array.isArray(base.crop) && base.crop.length === 4 ? base.crop : [0, 0, 1, 1];
    const setCrop = (c) => {
      const { crop: _c, ...rest } = base;
      return write(c ? { ...rest, crop: c } : rest);
    };
    const cropRow = h2(...crop.map((c, i) => numEl(c, '', (v) => {
      const next = [...crop];
      next[i] = +v;
      setCrop(next);
    }, 0.01)));
    cropRow.classList.add('crop4');
    rows.push(row('トリミング', cropRow));
    const acts = h('div', 'actions');
    acts.append(btn('画像を見ながらトリミング…', async () => {
      const nat = S.imageSizes.get(String(base.image));
      if (!nat) return status('画像が読めません', true);
      const r = opts.rect;
      const result = await openCrop({
        src: String(base.image),
        url: S.imageURLs.get(String(base.image)),
        natural: nat,
        crop,
        cellAspect: r && r.h > 0 ? r.w / r.h : nat.w / nat.h,
        lockDefault: base.mode === 'cover',
      });
      if (result === 'clear') return setCrop(null);
      if (Array.isArray(result)) return setCrop(result);
    }));
    if (base.crop) acts.append(btn('トリミングを外す', () => setCrop(null)));
    rows.push(acts);
  }
  return rows;
}

// ---------------------------------------------------------------- 選択対象（主レイヤー・背景・既存レイアウト）
// S.layer: 'main'（ふつうのボックス）/ 'bg'（クリックしたボックスの背景）/ 'block'（取り込んだレイアウトの塊）
// 既存レイアウト：取り込み元の印（from_layout）があるボックスの組か、
// フォルダの中の layout.yaml（取り込み元）と分割の形が同じボックスの組を、塊ごと選ぶ

// 分割の形。上下・左右と、子の並び（中身の種類は見ない。取り込んだ後に文字や画像を直しても同じ形になる）
function shapeSig(n) {
  if (!n || typeof n !== 'object') return 'L';
  if (n.split) return `${n.split}(${(n.children || []).map(shapeSig).join(',')})`;
  if (n.grid) return `grid(${(n.items || []).map(shapeSig).join(',')})`;
  return 'L';
}

async function loadImportSigs() {
  if (S.importSigs) return;
  const sigs = [];
  for (const f of S.layoutFiles) {
    try {
      const sub = await readSubLayout(f);
      for (const sl of sub.slides || []) {
        const g = shapeSig(sl.root);
        if (g.includes('(')) sigs.push([g, f]);
      }
    } catch {
      // 読めない layout.yaml は飛ばす
    }
  }
  S.importSigs = sigs;
}

function refreshAutoBlocks() {
  S.autoBlocks = new Map();
  const s = realSlide();
  if (!s || !S.importSigs) return;
  for (const b of s.boxes) {
    if (b.node?.from_layout) continue;
    const g = shapeSig(b.node);
    const hit = S.importSigs.find(([x]) => x === g);
    if (hit) S.autoBlocks.set(b.path, hit[1]);
  }
}

function blockFile(b) {
  return b?.node?.from_layout || S.autoBlocks.get(b?.path) || null;
}

function findReal(path) {
  return realSlide()?.boxes.find((b) => b.path === path) || null;
}

// b を含む、いちばん外側の取り込んだ塊（無ければ null）
function blockOf(b) {
  let top = null;
  for (let x = b; x; x = x.parent ? findReal(x.parent) : null) if (blockFile(x)) top = x;
  return top;
}

function blockRoots() {
  return (realSlide()?.boxes || []).filter((b) => blockFile(b) && blockOf(b) === b);
}

function blockLeafPaths(block) {
  return realSlide().boxes.filter((b) => isLeaf(b) && within(b, block)).map((b) => b.path);
}

function blocksLeaves(leaves) {
  const out = new Set();
  for (const l of leaves) {
    const blk = blockOf(l);
    if (blk) for (const p of blockLeafPaths(blk)) out.add(p);
  }
  return [...out];
}

// 背景で、クリックしたボックスのどの背景を選ぶか：自分か外側の枠で、背景を分けているもの。無ければ自分
function bgHostOf(b) {
  for (let x = b; x; x = x.parent ? findReal(x.parent) : null) if (isBgTree(x.node?.background)) return x;
  return b;
}

// b か、その外側のうち、背景が付いているいちばん内側のボックス（どこにも無ければ b）。
// 背景モードで開くのはこのボックス。例：背景を付けたボックスに文字を入れて、文字の下に余白のボックスができたとき、
// 文字のボックスではなく、外側の背景を開く
function bgOwnerOf(b) {
  for (let x = b; x; x = x.parent ? findReal(x.parent) : null) if (hasOwnBg(x)) return x;
  return b;
}

// 背景モードでマウスを乗せたときに示すもの（クリックしたときに開く・選ぶものと同じ）。
// もっと内側の背景が付いたボックスがあればそのボックス、無ければ、開いている背景の枠。背景と関係のない中身のボックスは示さない
function bgHoverAt(e) {
  const r = realLeafAt(e);
  const owner = r ? bgOwnerOf(bgHostOf(r)) : null;
  if (owner && owner.path !== S.bgEdit && hasOwnBg(owner)) return owner;
  return leafAt(e);
}

// ボックス x に背景が付いているか。中身の無い背景（{}。内部の印 __src などだけのもの）や、
// 背景を編集するために書き直した、分けていない空の木（{ background: {} }）は、背景なしとみなす
function hasOwnBg(x) {
  const filled = (o) => o != null && !(typeof o === 'object' && !Array.isArray(o) && !Object.keys(o).some((k) => !k.startsWith('__')));
  const bg = x?.node?.background;
  if (!filled(bg)) return false;
  if (isBgTree(bg) && !bg.split && !bg.grid) return filled(bg.background);
  return true;
}

// クリックした点で選ぶもの
//   主レイヤー：いちばん内側のボックス
//   背景：いま開いている背景の枠。その外なら、クリックしたボックス（背景を開き直す）
//   既存レイアウト：取り込んだ塊
function pickAt(e) {
  const b = leafAt(e);
  if (S.layer === 'block') return b ? blockOf(b) : null;
  if (S.layer === 'bg') {
    if (b) return b;
    const r = realLeafAt(e);
    return r ? bgHostOf(r) : null;
  }
  return b;
}

async function setLayer(v) {
  if (v === S.layer) return;
  const prev = S.layer;
  const prevSel = (S.sel || []).map(findBox).filter(Boolean);
  if (S.bgEdit) await leaveBgHost(v, []);
  else {
    S.layer = v;
    setSel([]);
  }
  if (v === 'bg') {
    // 選んでいたボックスがあれば、その背景から始める
    let host = null;
    if (prevSel.length) host = prev === 'block' ? blockOf(prevSel[0]) : prevSel.length === 1 ? prevSel[0] : commonAncestor(prevSel);
    host = host && findReal(host.path);
    if (host) host = bgOwnerOf(host);
    if (host && host.src?.mode === 'direct') return enterBgHost(host);
  }
  if (v === 'block') {
    await loadImportSigs();
    refreshAutoBlocks();
  }
  renderPreview();
  renderProps();
  if (v === 'bg') status('ボックスをクリックすると、そのボックスの背景を選びます。');
  else if (v === 'block') status(blockRoots().length ? '取り込んだレイアウトをクリックすると、塊ごと選びます。' : 'このスライドには、取り込んだレイアウトが見つかりません。');
  else status('');
}

function renderBlockProps(el) {
  const boxes = (S.sel || []).map(findBox).filter(Boolean);
  const blk = boxes.length ? blockOf(boxes[0]) : null;
  if (!blk) {
    el.append(h('h2', null, '既存レイアウト'));
    el.append(h('p', 'note', '取り込んだレイアウトをクリックすると、その塊ごと選びます。⌘ドラッグで、別の塊と入れ替えます。'));
    return;
  }
  // 塊の中のボックスを全部選んだのと同じ欄を出し、取り込み元を書き添える
  if (boxes.length > 1) renderMulti(el);
  else {
    S.layer = 'main';
    S.keepPanel = null;
    try {
      renderProps();
    } finally {
      S.layer = 'block';
    }
  }
  const meta = el.querySelector('p.meta');
  const src = h('p', 'meta', `取り込み元：${blockFile(blk)}`);
  if (meta) meta.after(src);
  else el.prepend(src);
}

// ---------------------------------------------------------------- 背景の編集
// ボックスの背景だけを、中身のボックスの分割とは別に分ける。
// 背景は layout.yaml の background に、ボックスと同じ形の木（split / ratio / children、葉は { background: 背景 }）で書く

function hostBox() {
  return S.bgEdit ? findReal(S.bgEdit) : null;
}

function bgEditBtn(box) {
  const b = btn('背景を編集', () => enterBgHost(box));
  // テンプレートの中のボックスは、背景を分けられない（書き先が 1 つに決まらないため）
  if (box.src?.mode !== 'direct') {
    b.disabled = true;
    b.title = 'テンプレートの一部のボックスでは使えません';
  }
  return b;
}

// box の背景を開く（選択対象を「背景」にする）
async function enterBgHost(box) {
  if (box.src?.mode !== 'direct') {
    status('テンプレートの一部のボックスは、背景を分けられません。', true);
    return false;
  }
  const bg = (box.node || {}).background;
  S.layer = 'bg';
  S.bgEdit = box.path;
  setSel([`${box.path}/bg`]);
  // 1 つの背景（や背景なし）は、分けていない木（葉 1 つ）に書き直してから編集する。見た目は変わらない
  if (!isBgTree(bg)) await edit(box, 'background', bg == null ? {} : { background: plain(bg) });
  else await rebuild();
  status('背景を選んでいます。分割・統合・削除をすると、背景の枠を分けたり、まとめたり、消したりします。中身のボックスは変わりません。');
  return true;
}

// 開いている背景を閉じる。分けていない木（葉 1 つ）は、もとの 1 つの背景の書き方に戻す
async function leaveBgHost(nextLayer, nextSel) {
  const host = hostBox();
  S.bgEdit = null;
  S.layer = nextLayer;
  setSel(nextSel);
  if (host?.src) {
    const t = getJS(S.files[host.src.file].doc, [...host.src.path, 'background']);
    // 余白を付けた背景は、1 つの背景の書き方では書けないので、木のまま残す（戻すと余白が消えていた）
    if (isBgTree(t) && !t.split && !t.grid && t.padding == null) {
      const v = t.background == null || (typeof t.background === 'object' && !Object.keys(t.background).length) ? undefined : t.background;
      return edit(host, 'background', v);
    }
  }
  return rebuild();
}

// 背景の編集を終えて、主レイヤーに戻る
async function exitBgEdit() {
  const hp = S.bgEdit;
  await leaveBgHost('main', hp ? [hp] : []);
  status('');
}

// 背景で、別のボックスをクリックしたとき：いまの背景を閉じ、そのボックスの背景を開いて、クリックした位置の枠を選ぶ
// 開くのは、クリックしたボックスか、その外側のうち、背景が付いているいちばん内側のボックス（bgOwnerOf）
async function switchBgHost(box, pt) {
  box = bgOwnerOf(box);
  const path = box.path;
  if (box.src?.mode !== 'direct') return status('テンプレートの一部のボックスは、背景を分けられません。', true);
  if (S.bgEdit) await leaveBgHost('bg', []);
  const host = findReal(path);
  if (!host || !(await enterBgHost(host))) return;
  const hit = curSlide().boxes.filter((b) => isLeaf(b) && pt.x >= b.rect.x && pt.x <= b.rect.x + b.rect.w && pt.y >= b.rect.y && pt.y <= b.rect.y + b.rect.h);
  hit.sort((a, b) => b.depth - a.depth);
  if (hit[0]) select(hit[0].path);
}

function renderBgProps(el) {
  const host = hostBox();
  el.append(h('h2', null, '背景を編集中'));
  el.append(h('p', 'note', 'このボックスの背景だけを分けています。中身のボックスの分割とは別です。上の「分割」「統合」「削除」を押すと、背景の枠を分けたり、まとめたり、消したりします。'));
  const done = h('div', 'actions');
  done.append(btn('背景の編集を終える', () => exitBgEdit()));
  el.append(done);
  if (!host) {
    el.append(h('p', 'note', 'プレビューでボックスをクリックすると、そのボックスの背景を選びます。'));
    return;
  }
  const boxes = (S.sel || []).map(findBox).filter(Boolean);
  const u = S.layout.unit;
  if (!boxes.length) {
    el.append(h('p', 'note', 'プレビューで背景の枠をクリックすると選択し、ここに設定が出ます。'));
    return;
  }
  if (boxes.length === 1) {
    const box = boxes[0];
    el.append(h('p', 'meta', `背景の枠　${fmt(box.rect.w)} × ${fmt(box.rect.h)} ${u}`));
    const fsSize = sizeFields(box);
    if (fsSize) el.append(fsSize);
  } else {
    const r = selectionRect();
    el.append(h('p', 'meta', `背景の枠 ${boxes.length} 個　${fmt(r.w)} × ${fmt(r.h)} ${u}`));
    const lca = commonAncestor(boxes);
    if (lca?.src) {
      const kids = curSlide().boxes.filter((b) => b.parent === lca.path);
      const all = kids.filter((k) => overlaps(k.rect, r)).length === kids.length;
      el.append(multiSizeFields(lca, r, all));
    }
  }
  // 背景は、選んだ枠の 1 つずつに同じ値を設定する（値は、いちばん上か左の枠のものを出す）
  const first = boxes[0];
  // 余白：その分だけ、背景を枠の内側に縮める
  const fsPad = fieldset(boxes.length > 1 ? '余白（選んだ枠の 1 つずつ）' : '余白');
  fsPad.append(paddingRow(first, first.node || {}, boxes));
  el.append(fsPad);
  const fsBg = fieldset(boxes.length > 1 ? '背景（選んだ枠の 1 つずつ）' : '背景');
  fsBg.append(...bgRows((first.node || {}).background, (v) => {
    const files = new Set();
    for (const b of boxes) {
      if (!b.src) continue;
      setProp(S.files[b.src.file].doc, b.src.path, 'background', v === undefined ? undefined : structuredClone(v));
      files.add(b.src.file);
    }
    for (const f of files) fileChanged(f);
  }, { rect: bgInner(first) }));
  el.append(fsBg);
}

// 背景の枠 b の、余白を除いた範囲（背景を敷く範囲）
function bgInner(b) {
  const p = parseBox(b.node?.padding, S.layout.unit);
  return { x: b.rect.x + p.l, y: b.rect.y + p.t, w: b.rect.w - p.l - p.r, h: b.rect.h - p.t - p.b };
}

const KIND_LABEL = { empty: '空き', split: '分割', grid: '格子', image: '画像', text: 'テキスト' };
const KIND_OPTS = [['text', 'テキスト'], ['image', '画像'], ['layout', 'レイアウトを取り込む…']];
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
  // 色を選ぶ画面に、このスライドで使っている色を候補として出す
  c.setAttribute('list', refreshColorList());
  const norm = normColor(value);
  c.value = norm || '#000000';
  const t = textEl(value ?? '', '（なし）', (v) => onCommit(v));
  c.oninput = () => (t.value = c.value);
  c.onchange = () => onCommit(c.value);
  wrap.append(c, t);
  return wrap;
}

// いまのスライドで使っている色（多く使っている順。"#RRGGBB"）。文字の色・一部の文字の色・文字の下地・光彩・背景・枠線
function usedColors() {
  const s = realSlide();
  if (!s) return [];
  const count = new Map();
  const add = (c) => {
    const x = c && typeof c === 'string' ? toHex6(c) : null;
    if (x && x !== NO_COLOR.slice(1).toUpperCase()) count.set(x, (count.get(x) || 0) + 1);
  };
  for (const e of s.elements || []) {
    if (e.type === 'text') {
      add(e.font?.color);
      add(e.font?.glow?.color);
      add(e.text_bg?.color);
      for (const m of String(e.text).matchAll(/<color\s+#?([0-9a-fA-F]{6})\s*>/gi)) add(m[1]);
    } else if (e.type === 'rect') {
      add(e.fill);
      add(e.border?.color);
    }
  }
  return [...count.entries()].sort((a, b) => b[1] - a[1]).map(([x]) => `#${x}`);
}

// 色の入力欄の候補（<datalist>）を、いまのスライドの色に作り直し、その id を返す
function refreshColorList() {
  let dl = document.getElementById('used-colors');
  if (!dl) {
    dl = document.createElement('datalist');
    dl.id = 'used-colors';
    document.body.append(dl);
  }
  dl.replaceChildren(...usedColors().map((c) => Object.assign(document.createElement('option'), { value: c.toLowerCase() })));
  return dl.id;
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
    const { underlines, textBgs } = await measureForPPTX();
    const pres = await buildPPTX(window.PptxGenJS, S.layout, async (src) => ({ data: await blobToDataURL(await S.source.readBlob(src)) }), { underlines, textBgs });
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

// 画像の線（<u img=…>）の位置と、文字の下地の大きさを、隠した画面に全部のスライドを描いて測る
// （画像の線は PPTX に画像として置き、下地はテキストボックスの大きさにする）
async function measureForPPTX() {
  const ul = hasImageUnderline(S.layout);
  const tb = hasTextBg(S.layout);
  if (!ul && !tb) return { underlines: [], textBgs: [] };
  const f = document.createElement('iframe');
  f.style.cssText = 'position:fixed;left:-30000px;top:0;width:3000px;height:2000px;border:0;visibility:hidden';
  document.body.append(f);
  try {
    await new Promise((res) => {
      f.onload = res;
      f.srcdoc = renderHTML(S.layout, { srcURL: (src) => S.imageURLs.get(src) || '', waitFonts: true });
    });
    await loadUsedFonts(f.contentDocument);
    return {
      underlines: ul ? measureUnderlines(f.contentDocument, S.layout.width) : [],
      textBgs: tb ? measureTextBgs(f.contentDocument, S.layout.width) : [],
    };
  } finally {
    f.remove();
  }
}

// PDF は、このページそのものを印刷して作る。印刷のあいだだけ、アプリの画面を隠してポスターだけを出し、
// 用紙をポスターの大きさ・余白 0 にする（余白が 0 なので、日付・題名・URL・ページ番号のヘッダーとフッターは入らない）。
// 画面の中の見えない枠（iframe）を印刷すると、Chrome が枠の用紙の指定を使わず、普通の用紙と余白で印刷し、
// ヘッダーとフッターを入れていた。新しいウィンドウを開く方法は、ポップアップとして止められることがある
// 印刷のための、ポスターだけの部分（#print-root）と、印刷のときだけ効く書き方を、このページに足す。片付ける関数を返す
function mountPrint() {
  const html = renderHTML(S.layout, { srcURL: (src) => S.imageURLs.get(src) || '', waitFonts: true });
  const src = new DOMParser().parseFromString(html, 'text/html');
  const posterCSS = [...src.querySelectorAll('style')].map((s) => s.textContent).join('\n');
  // 印刷のときだけ効かせる（画面のアプリの見た目は変えない）。アプリの「* { box-sizing: border-box }」などを打ち消す
  const style = document.createElement('style');
  style.id = 'print-css';
  style.textContent = `#print-root { display: none; }
@media print {
${posterCSS}
html, body { height: auto !important; display: block !important; background: #fff !important; margin: 0 !important; padding: 0 !important; }
body > *:not(#print-root) { display: none !important; }
#print-root { display: block !important; }
#print-root *:not(section.slide > div) { box-sizing: content-box; }
}`;
  // Google Fonts を使う文字があれば、その読み込みもこのページに足す
  const links = [...src.querySelectorAll('link[rel=stylesheet]')].map((l) => {
    const n = document.createElement('link');
    n.rel = 'stylesheet';
    n.href = l.href;
    n.dataset.print = '1';
    return n;
  });
  const root = document.createElement('div');
  root.id = 'print-root';
  root.innerHTML = src.body.innerHTML;
  document.head.append(style, ...links);
  document.body.append(root);
  const title = document.title;
  // PDF の名前の既定は、ページの題名になる
  document.title = String(S.source?.name || 'slide');
  S.printing = true;
  const cleanup = () => {
    root.remove();
    style.remove();
    for (const l of links) l.remove();
    document.title = title;
    S.printing = false;
  };
  return { root, links, cleanup };
}

// ブラウザの印刷（Ctrl+P・⌘P）で印刷するときも、アプリの画面ではなく、ポスターだけを余白 0 で印刷する
window.addEventListener('beforeprint', () => {
  if (S.printing || !S.layout) return;
  const { cleanup } = mountPrint();
  window.addEventListener('afterprint', cleanup, { once: true });
});

async function exportPDF() {
  if (!S.layout) return;
  const { root, links, cleanup } = mountPrint();
  status('印刷画面で「PDFに保存」を選んでください');
  try {
    // フォントと画像を読み込み終えてから印刷する（待たないと、日本語が別のフォントで PDF に入る）
    await Promise.all(links.map((l) => new Promise((r) => { l.onload = l.onerror = r; })));
    await loadUsedFonts(document);
    await Promise.all([...root.querySelectorAll('img')].map((i) => (i.complete ? null : new Promise((r) => { i.onload = i.onerror = r; }))));
    window.addEventListener('afterprint', cleanup, { once: true });
    window.print();
  } catch (e) {
    cleanup();
    status(`PDF を作れません: ${e.message}`, true);
  }
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
  if (S.layer === 'bg' && mod && (e.code === 'KeyC' || e.code === 'KeyV')) return;
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
  if ((e.key === 'Backspace' || e.key === 'Delete') && !mod && (S.sel || []).length && S.layer !== 'bg') {
    e.preventDefault();
    clearSelection();
    return;
  }
  if (e.key === 'Escape') {
    if (S.layer === 'bg' && !(S.sel || []).length) return exitBgEdit();
    select(null);
  }
}

// 選んだボックス（複数も）の中身を消して、空のテキストボックスにする。ボックスの大きさ・余白・背景・枠線・文字の設定は残す
const CLEAR_KEYS = ['image', 'crop', 'fit', 'split', 'children', 'ratio', 'gap', 'grid', 'items', 'icon', 'icon_size', 'text_bg', 'from_layout'];
function clearSelection() {
  const slide = curSlide();
  const leaves = [...new Set((S.sel || []).map(findBox).filter(Boolean).flatMap((b) => slide.boxes.filter((x) => isLeaf(x) && within(x, b))))];
  const targets = leaves.filter((b) => b.src && !(b.kind === 'text' && String(b.node?.text ?? '') === ''));
  if (!targets.length) return;
  const files = new Set();
  for (const b of targets) {
    const { file, path, mode } = b.src;
    const doc = S.files[file].doc;
    if (mode === 'direct') {
      let cur = getJS(doc, path);
      if (typeof cur === 'string') cur = {};
      const keep = cur && typeof cur === 'object' ? Object.fromEntries(Object.entries(cur).filter(([k]) => !CLEAR_KEYS.includes(k) && k !== 'text')) : {};
      replaceNode(doc, path, { ...keep, text: '' });
    } else {
      // 型の値を打ち消すため、ほかの種類の中身は null にする
      for (const k of CLEAR_KEYS) setProp(doc, path, k, null);
      setProp(doc, path, 'text', '');
    }
    files.add(file);
  }
  const sel = S.sel;
  return Promise.all([...files].map((f) => fileChanged(f))).then(() => {
    setSel(sel.filter((p) => findBox(p)));
    applyView();
    renderProps();
    status(targets.length > 1 ? `${targets.length} 個のボックスを空にしました` : 'ボックスを空にしました');
  });
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
$('#btn-split').onclick = openSplitGrid;
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
$('#layer').onchange = (e) => setLayer(e.target.value);
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
