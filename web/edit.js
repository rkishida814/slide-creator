// YAML の書き換え。コメントや書き方をできるだけ残すため、yaml の Document を直接いじる。
import * as YAML from './vendor/yaml/index.js';

export function parse(text) {
  const doc = YAML.parseDocument(text);
  if (doc.errors.length) {
    const e = doc.errors[0];
    const line = e.linePos?.[0]?.line;
    throw new Error(`layout.yaml の書き方に誤りがあります${line ? `（${line}行目）` : ''}: ${e.message.split('\n')[0]}`);
  }
  return doc;
}

export function stringify(doc) {
  quoteRiskyFlowStrings(doc);
  const text = doc.toString({ lineWidth: 0 });
  if (!YAML.parseDocument(text).errors.length) return text;
  // それでも読み直せない書き方になったら、{ } [ ] の中の改行を含む文字を、すべて引用符で囲んで書き直す
  quoteRiskyFlowStrings(doc, true);
  return doc.toString({ lineWidth: 0 });
}

// yaml（2.9.1）は、{ } [ ] の中の文字で、行が「:」で終わり改行が続くもの（例："Objectives:\n本文"）を、
// 引用符なしで書き出してしまい、読み直すと誤りになる（「text:」の中に別の対応表があるように読める）。
// そうなる文字は、書き出す前に二重引用符で囲む。all なら、{ } [ ] の中の改行を含む文字をすべて囲む
function quoteRiskyFlowStrings(doc, all = false) {
  YAML.visit(doc, {
    Scalar(_, node, path) {
      if (typeof node.value !== 'string' || !node.value.includes('\n')) return;
      if (node.type && node.type !== YAML.Scalar.PLAIN) return;
      if (!path.some((p) => YAML.isCollection(p) && p.flow)) return;
      if (all || /:[ \t]*$/m.test(node.value)) node.type = YAML.Scalar.QUOTE_DOUBLE;
    },
  });
}

// 値を YAML のノードにする。数値・文字の配列（ratio, crop, padding）は [a, b] の形で書く
function toNode(doc, value) {
  const n = toNode1(doc, value);
  // 中の小さい配列・対応表も、1 行の形にする（取り込んだレイアウトなど、大きな木を書くとき）
  if (YAML.isCollection(n)) {
    YAML.visit(n, {
      Seq(_, s) {
        if (s.items.every((v) => YAML.isScalar(v))) s.flow = true;
      },
      Map(_, m) {
        if (m.items.length <= 4 && m.items.every((p) => YAML.isScalar(p.value) || (YAML.isCollection(p.value) && p.value.flow))) m.flow = true;
      },
    });
  }
  return n;
}

function toNode1(doc, value) {
  if (Array.isArray(value) && value.every((v) => v == null || typeof v !== 'object')) {
    const n = doc.createNode(value, { aliasDuplicateObjects: false });
    n.flow = true;
    return n;
  }
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    const n = doc.createNode(value, { aliasDuplicateObjects: false });
    // 小さい対応表（font など）は { } の形で 1 行に書く
    if (Object.values(value).every((v) => v == null || typeof v !== 'object') && Object.keys(value).length <= 4) n.flow = true;
    return n;
  }
  return doc.createNode(value, { aliasDuplicateObjects: false });
}

// path の位置がノード（対応表）でなければ、対応表に直す。文字だけのノードは { text: … } にする
function ensureMap(doc, path) {
  if (!path.length) return;
  const cur = doc.getIn(path, true);
  if (YAML.isMap(cur)) return;
  if (cur == null || (YAML.isScalar(cur) && cur.value == null)) {
    doc.setIn(path, doc.createNode({}, { aliasDuplicateObjects: false }));
  } else if (YAML.isScalar(cur)) {
    doc.setIn(path, doc.createNode({ text: cur.value }, { aliasDuplicateObjects: false }));
  }
}

/**
 * ノードの設定を 1 つ変える。
 * @param key    'gap' や ['font', 'size']
 * @param value  undefined なら消す
 */
export function setProp(doc, path, key, value) {
  ensureMap(doc, path);
  const keys = Array.isArray(key) ? key : [key];
  const full = [...path, ...keys];
  if (value === undefined) {
    if (doc.hasIn(full)) doc.deleteIn(full);
    // 空になった font: {} などを消す
    if (keys.length > 1) {
      const parent = doc.getIn([...path, ...keys.slice(0, -1)], true);
      if (YAML.isMap(parent) && parent.items.length === 0) doc.deleteIn([...path, ...keys.slice(0, -1)]);
    }
    return;
  }
  if (keys.length > 1) {
    const parentPath = [...path, ...keys.slice(0, -1)];
    const parent = doc.getIn(parentPath, true);
    if (!YAML.isMap(parent)) doc.setIn(parentPath, doc.createNode({}, { aliasDuplicateObjects: false }));
  }
  doc.setIn(full, toNode(doc, value));
}

// ノードを丸ごと置き換える（分割・空にする など）
export function replaceNode(doc, path, value) {
  doc.setIn(path, toNode(doc, value));
}

export function getJS(doc, path) {
  const n = doc.getIn(path, true);
  return n && typeof n.toJSON === 'function' ? n.toJSON() : n;
}

export { YAML };
