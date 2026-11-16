// テンプレートの展開。ファイルに依存しない。
// ノードに use: 名前 があれば、テンプレートの木を写し、
//   fill: { 子のid: {上書きする設定} } で中身を入れ、
//   use / fill 以外のキーはテンプレートの根に上書きする。

export function resolveDoc(doc, templates = {}) {
  const all = { ...templates, ...(doc.templates || {}) };
  const out = { ...doc };
  out.slides = (doc.slides || []).map((s) => {
    if (s.use) {
      // スライドそのものに use を書いた場合は root に展開する
      const { id, background, use, fill, ...rest } = s;
      return { id, background, root: resolveNode({ use, fill, ...rest }, all, []) };
    }
    return { ...s, root: resolveNode(s.root, all, []) };
  });
  return out;
}

function resolveNode(node, templates, stack) {
  if (node == null || typeof node !== 'object') return node;
  if (node.use) {
    const name = node.use;
    if (stack.includes(name)) throw new Error(`テンプレートが循環している: ${[...stack, name].join(' → ')}`);
    const t = templates[name];
    if (!t) throw new Error(`テンプレートが無い: ${name}`);
    let base = resolveNode(structuredClone(t), templates, [...stack, name]);
    const { use, fill, ...override } = node;
    base = { ...base, ...override };
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

function applyFill(root, fill) {
  const found = new Set();
  const visit = (n) => {
    if (!n || typeof n !== 'object') return;
    if (n.id && fill[n.id]) {
      Object.assign(n, fill[n.id]);
      found.add(n.id);
    }
    (n.children || []).forEach(visit);
    (n.items || []).forEach(visit);
  };
  visit(root);
  return Object.keys(fill).filter((k) => !found.has(k));
}
