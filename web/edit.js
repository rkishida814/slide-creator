// YAML の書き換え。コメントや書き方をできるだけ残すため、yaml の Document を直接いじる。
import * as YAML from './vendor/yaml/index.js';

export function parse(text) {
  const doc = YAML.parseDocument(text);
  if (doc.errors.length) {
    const e = doc.errors[0];
    const line = e.linePos?.[0]?.line;
    throw new Error(`YAML の誤り${line ? `（${line}行目）` : ''}: ${e.message.split('\n')[0]}`);
  }
  return doc;
}

export function stringify(doc) {
  return doc.toString({ lineWidth: 0 });
}

// 値を YAML のノードにする。数値・文字の配列（ratio, crop, padding）は [a, b] の形で書く
function toNode(doc, value) {
  if (Array.isArray(value) && value.every((v) => v == null || typeof v !== 'object')) {
    const n = doc.createNode(value);
    n.flow = true;
    return n;
  }
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    const n = doc.createNode(value);
    // 小さい対応表（font など）は { } の形で 1 行に書く
    if (Object.values(value).every((v) => v == null || typeof v !== 'object') && Object.keys(value).length <= 4) n.flow = true;
    return n;
  }
  return doc.createNode(value);
}

// path の位置がノード（対応表）でなければ、対応表に直す。文字だけのノードは { text: … } にする
function ensureMap(doc, path) {
  if (!path.length) return;
  const cur = doc.getIn(path, true);
  if (YAML.isMap(cur)) return;
  if (cur == null || (YAML.isScalar(cur) && cur.value == null)) {
    doc.setIn(path, doc.createNode({}));
  } else if (YAML.isScalar(cur)) {
    doc.setIn(path, doc.createNode({ text: cur.value }));
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
    if (!YAML.isMap(parent)) doc.setIn(parentPath, doc.createNode({}));
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
