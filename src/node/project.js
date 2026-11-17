// プロジェクトフォルダの読み込み（Node 用）。
import fs from 'node:fs';
import path from 'node:path';
import YAML from 'yaml';
import { imageSize } from 'image-size';
import { resolveDoc } from '../core/resolve.js';
import { layoutDoc } from '../core/layout.js';

export const LAYOUT_FILE = 'layout.yaml';

export function loadProject(dir) {
  dir = path.resolve(dir);
  const file = path.join(dir, LAYOUT_FILE);
  if (!fs.existsSync(file)) throw new Error(`${file} がありません（slide init で作れます）`);
  const doc = YAML.parse(fs.readFileSync(file, 'utf8'));
  const templates = loadTemplates(path.join(dir, 'templates'));
  const resolved = resolveDoc(doc, templates);
  const cache = new Map();
  const nested = new Map();
  const env = {
    // ボックスに入れる別のレイアウト（file はこのフォルダからの場所。中の templates/ も読む）
    layout(file) {
      if (nested.has(file)) return nested.get(file);
      let v = null;
      const abs = path.join(dir, file);
      if (fs.existsSync(abs)) {
        try {
          const sub = YAML.parse(fs.readFileSync(abs, 'utf8'));
          v = resolveDoc(sub, loadTemplates(path.join(path.dirname(abs), 'templates')));
        } catch (e) {
          v = { error: e.message };
        }
      }
      nested.set(file, v);
      return v;
    },
    imageSize(src) {
      if (cache.has(src)) return cache.get(src);
      let v = null;
      try {
        const info = imageSize(fs.readFileSync(path.join(dir, src)));
        const rot = info.orientation >= 5 && info.orientation <= 8; // EXIF で90度回転
        v = rot ? { w: info.height, h: info.width } : { w: info.width, h: info.height };
      } catch {
        v = null;
      }
      cache.set(src, v);
      return v;
    },
  };
  const layout = layoutDoc(resolved, env);
  return { dir, doc, resolved, layout };
}

function loadTemplates(tdir) {
  const out = {};
  if (!fs.existsSync(tdir)) return out;
  for (const f of fs.readdirSync(tdir)) {
    if (!/\.ya?ml$/.test(f)) continue;
    out[f.replace(/\.ya?ml$/, '')] = YAML.parse(fs.readFileSync(path.join(tdir, f), 'utf8'));
  }
  return out;
}

// "s1" / "s1/photos" / "photos"（全スライドから探す）/ 自動の経路 "s1.1.0"
export function findTarget(layout, slideArg, nodeArg) {
  let slide = slideArg ? layout.slides.find((s) => s.id === slideArg) : null;
  if (slideArg && !slide) throw new Error(`スライドが無い: ${slideArg}`);
  if (!nodeArg) return { slide: slide || layout.slides[0], box: null };
  if (nodeArg.includes('/') && !slide) {
    const [sid] = nodeArg.split('/');
    slide = layout.slides.find((s) => s.id === sid);
  }
  const pool = slide ? [slide] : layout.slides;
  for (const s of pool) {
    const box = s.boxes.find((b) => b.path === nodeArg || b.id === nodeArg || b.path === `${s.id}/${nodeArg}`);
    if (box) return { slide: s, box };
  }
  throw new Error(`ノードが無い: ${nodeArg}`);
}
