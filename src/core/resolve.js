import { isBgTree } from './units.js';

// テンプレートの展開。ファイルに依存しない。
// ノードに use: 名前 があれば、テンプレートの木を写し、
//   fill: { 子のid: {上書きする設定} } で中身を入れ、
//   use / fill 以外のキーはテンプレートの根に上書きする。
//
// 展開の前に各ノードへ __src（編集先）を付ける。画面で選んだノードの設定を変えるとき、
// どのファイルのどの位置を書き換えればよいかを示す。
//   { file: 'layout.yaml' | 'templates/x.yaml', path: [...], mode }
//   mode: 'direct'   … そのノード自身の位置（構造も変えてよい）
//         'fill'     … テンプレートの中の id 付きノード。use 側の fill[id] に書く（このスライドだけに効く）
//         'override' … テンプレートの根。use を書いたノードに書く（このスライドだけに効く）
//         'shared'   … テンプレートの中の id の無いノード。型の定義に書く（型を使う全箇所に効く）

export function resolveDoc(doc, templates = {}) {
  doc = structuredClone(doc);
  const all = {};
  for (const [name, t] of Object.entries(templates)) {
    all[name] = annotate(structuredClone(t), `templates/${name}.yaml`, [], 'shared');
  }
  for (const [name, t] of Object.entries(doc.templates || {})) {
    all[name] = annotate(t, 'layout.yaml', ['templates', name], 'shared');
  }
  const out = { ...doc };
  out.slides = (doc.slides || []).map((s, i) => {
    if (s.use) {
      // スライドそのものに use を書いた場合は root に展開する
      const { id, background, use, fill, ...rest } = s;
      const node = annotate({ use, fill, ...rest }, 'layout.yaml', ['slides', i], 'direct');
      return { id, background, root: resolveNode(node, all, []) };
    }
    const root = annotate(s.root, 'layout.yaml', ['slides', i, 'root'], 'direct');
    return { ...s, root: resolveNode(root, all, []) };
  });
  return out;
}

// 文字だけ・空のノードも、編集先を付けられるようにオブジェクトにする
function annotate(node, file, path, mode) {
  if (node == null) node = {};
  else if (typeof node !== 'object') node = { text: String(node), __scalar: true };
  else if (Array.isArray(node)) return node;
  node.__src = { file, path, mode };
  // 背景を分けた木も、編集先を付ける
  if (isBgTree(node.background)) node.background = annotate(node.background, file, [...path, 'background'], mode);
  if (Array.isArray(node.children)) node.children = node.children.map((c, k) => annotate(c, file, [...path, 'children', k], mode));
  if (Array.isArray(node.items)) node.items = node.items.map((c, k) => annotate(c, file, [...path, 'items', k], mode));
  if (node.fill && typeof node.fill === 'object') {
    for (const [id, f] of Object.entries(node.fill)) {
      if (!f || typeof f !== 'object') continue;
      if (Array.isArray(f.children)) f.children = f.children.map((c, k) => annotate(c, file, [...path, 'fill', id, 'children', k], mode));
      if (Array.isArray(f.items)) f.items = f.items.map((c, k) => annotate(c, file, [...path, 'fill', id, 'items', k], mode));
    }
  }
  return node;
}

function resolveNode(node, templates, stack) {
  if (node == null || typeof node !== 'object') return node;
  if (node.use) {
    const name = node.use;
    if (stack.includes(name)) throw new Error(`テンプレートが循環している: ${[...stack, name].join(' → ')}`);
    const t = templates[name];
    if (!t) throw new Error(`テンプレートが無い: ${name}`);
    const useSrc = node.__src;
    let base = resolveNode(structuredClone(t), templates, [...stack, name]);
    const { use, fill, __src, ...override } = node;
    base = { ...base, ...override };
    if (useSrc) {
      // 型の中の id 付きノードは、この use の fill[id] を編集先にする
      retarget(base, useSrc);
      base.__src = { ...useSrc, mode: useSrc.mode === 'direct' ? 'override' : useSrc.mode };
      base.__template = name;
    }
    if (fill) {
      const missing = applyFill(base, fill);
      if (missing.length) throw new Error(`テンプレート ${name} に id が無い: ${missing.join(', ')}`);
    }
    return resolveNode(base, templates, stack);
  }
  const out = { ...node };
  if (Array.isArray(node.children)) out.children = node.children.map((c) => resolveNode(c, templates, stack));
  if (Array.isArray(node.items)) out.items = node.items.map((c) => resolveNode(c, templates, stack));
  return out;
}

function retarget(root, useSrc) {
  const visit = (n, isRoot) => {
    if (!n || typeof n !== 'object') return;
    if (!isRoot && n.id) {
      n.__src = { file: useSrc.file, path: [...useSrc.path, 'fill', n.id], mode: useSrc.mode === 'shared' ? 'shared' : 'fill' };
    }
    (n.children || []).forEach((c) => visit(c, false));
    (n.items || []).forEach((c) => visit(c, false));
  };
  visit(root, true);
}

function applyFill(root, fill) {
  const found = new Set();
  const visit = (n) => {
    if (!n || typeof n !== 'object') return;
    if (n.id && fill[n.id]) {
      const { __src, ...f } = fill[n.id];
      Object.assign(n, f);
      found.add(n.id);
    }
    (n.children || []).forEach(visit);
    (n.items || []).forEach(visit);
  };
  visit(root);
  return Object.keys(fill).filter((k) => !found.has(k));
}
