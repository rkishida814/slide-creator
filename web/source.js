// プロジェクトの読み書き先。
//   FolderSource … 手元のフォルダ（Chrome / Edge のフォルダアクセス機能）。保存はフォルダに直接書く
//   UrlSource    … Web 上のフォルダ（見本用。読むだけで、保存はダウンロードになる）

export class FolderSource {
  constructor(handle) {
    this.handle = handle;
    this.name = handle.name;
    this.writable = true;
    this.urls = new Map();
  }

  static async pick() {
    if (!window.showDirectoryPicker) throw new Error('このブラウザはフォルダを開けません（Chrome か Edge を使ってください）');
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    return new FolderSource(handle);
  }

  async dirOf(parts, create = false) {
    let d = this.handle;
    for (const p of parts) d = await d.getDirectoryHandle(p, { create });
    return d;
  }

  async fileHandle(path, create = false) {
    const parts = path.split('/').filter(Boolean);
    const name = parts.pop();
    const dir = await this.dirOf(parts, create);
    return dir.getFileHandle(name, { create });
  }

  async exists(path) {
    try {
      await this.fileHandle(path);
      return true;
    } catch {
      return false;
    }
  }

  async readText(path) {
    return (await (await this.fileHandle(path)).getFile()).text();
  }

  async readBlob(path) {
    return (await this.fileHandle(path)).getFile();
  }

  async writeFile(path, data) {
    const fh = await this.fileHandle(path, true);
    const w = await fh.createWritable();
    await w.write(data);
    await w.close();
  }

  async imageURL(path) {
    if (!this.urls.has(path)) this.urls.set(path, URL.createObjectURL(await this.readBlob(path)));
    return this.urls.get(path);
  }

  // 画像の再読み込み（外で画像を差し替えたとき）
  forgetImages() {
    for (const u of this.urls.values()) URL.revokeObjectURL(u);
    this.urls.clear();
  }

  // dir の下のファイルを再帰で列挙（"assets/a.png" の形）
  async list(dir, depth = 3) {
    const out = [];
    let d;
    try {
      d = await this.dirOf(dir.split('/').filter(Boolean));
    } catch {
      return out;
    }
    const walk = async (h, prefix, n) => {
      for await (const [name, child] of h.entries()) {
        if (name.startsWith('.')) continue;
        if (child.kind === 'file') out.push(prefix + name);
        else if (n > 0) await walk(child, `${prefix}${name}/`, n - 1);
      }
    };
    await walk(d, dir ? `${dir.replace(/\/$/, '')}/` : '', depth);
    return out.sort();
  }
}

export class UrlSource {
  constructor(base, name) {
    this.base = base.endsWith('/') ? base : `${base}/`;
    this.name = name;
    this.writable = false;
  }

  url(path) {
    return new URL(path, new URL(this.base, location.href)).href;
  }

  async exists(path) {
    const r = await fetch(this.url(path), { method: 'HEAD' });
    return r.ok;
  }

  async readText(path) {
    const r = await fetch(this.url(path));
    if (!r.ok) throw new Error(`${path} を読めません（${r.status}）`);
    return r.text();
  }

  async readBlob(path) {
    const r = await fetch(this.url(path));
    if (!r.ok) throw new Error(`${path} を読めません（${r.status}）`);
    return r.blob();
  }

  async writeFile(path, data) {
    download(data instanceof Blob ? data : new Blob([data]), path.split('/').pop());
  }

  async imageURL(path) {
    return this.url(path);
  }

  forgetImages() {}

  async list() {
    return []; // Web 上のフォルダは一覧できない
  }
}

export function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}
