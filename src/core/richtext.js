// 文字の飾りと改行。layout.yaml の text の中では、次だけを特別に読む（それ以外の < > はそのままの文字）。
//   <b>太字</b>  <i>斜体</i>  <sup>上付き</sup>  <sub>下付き</sub>
//   <u>下線（一重線）</u>  <u double>二重線</u>  <u img="line.png" w="2">画像の線</u>
//     画像の線は、画像を縦横比を保たずに細長く伸ばして、文字の下に引く。w は線の太さ（pt。省くと 2）
//   改行（\n）… 段落の区切り（Enter）。段落の間隔（para_space）が足される
//   <br>      … 段落の中の改行（Shift+Enter）。行間だけで、段落の間隔は足さない
//   <li>      … 段落の頭に書くと、その段落を箇条書きにする（行頭に「•」。折り返した行は文字の頭に揃う）
//   <li hidden> … 「•」を描かない箇条書き（字下げだけ。前の箇条書きの文字の頭に揃える）
//   <color #RRGGBB>色を付ける部分</color> … 一部の文字の色
// 例: "New bone area (mm<sup>2</sup>)"、"<color #C00000>有意差あり</color>"、"<li>1 つ目\n<li>2 つ目"
const TAG = /<(\/?)(b|i|sup|sub)>|<color\s+#?([0-9a-fA-F]{6})\s*>|<\/(color)>|<br\s*\/?>|<u(\s[^>]*)?>|<\/(u)>|<li(?:\s+hidden)?>/gi;

// 段落に分ける
export function paragraphs(s) {
  return String(s ?? '').split('\n');
}

// 段落が箇条書きか（段落の頭の <li>）
export function isBullet(p) {
  return /^<li(\s+hidden)?>/i.test(String(p ?? ''));
}

// 箇条書きの種類：'normal'（•）/ 'hidden'（•を描かない。字下げだけ）/ null（箇条書きでない）
export function bulletType(p) {
  const m = String(p ?? '').match(/^<li(\s+hidden)?>/i);
  return m ? (m[1] ? 'hidden' : 'normal') : null;
}

// <u …> の中身 → { type: single | double | img, img, w }
function underlineOf(attr) {
  const a = String(attr || '');
  const img = a.match(/img\s*=\s*"([^"]*)"/i) || a.match(/img\s*=\s*([^\s"]+)/i);
  if (img) {
    const w = a.match(/\bw\s*=\s*"?([\d.]+)"?/i);
    return { type: 'img', img: img[1], w: w ? +w[1] : 2 };
  }
  return { type: /\bdouble\b/i.test(a) ? 'double' : 'single' };
}

// 1 つの段落 → [{ text, b, i, sup, sub, color, u, br }]（br は段落の中の改行。段落の頭の <li> は読み飛ばす）
export function parseRich(s) {
  const out = [];
  const st = { b: 0, i: 0, sup: 0, sub: 0 };
  const colors = []; // 色の入れ子（いちばん内側が効く）
  const unders = []; // 下線の入れ子（いちばん内側が効く）
  let last = 0;
  const src = String(s ?? '');
  const push = (text) => {
    if (!text) return;
    const g = { text, b: st.b > 0, i: st.i > 0, sup: st.sup > 0, sub: st.sub > 0 };
    if (colors.length) g.color = colors[colors.length - 1];
    if (unders.length) g.u = unders[unders.length - 1];
    out.push(g);
  };
  for (const m of src.matchAll(TAG)) {
    const [all, close, tag, color, colorEnd, uAttr, uEnd] = m;
    push(src.slice(last, m.index));
    last = m.index + all.length;
    if (color) colors.push(color.toUpperCase());
    else if (colorEnd) colors.pop();
    else if (uEnd) unders.pop();
    else if (/^<u[\s>]/i.test(all)) unders.push(underlineOf(uAttr));
    else if (/^<li[\s>]/i.test(all)) continue;
    else if (!tag) out.push({ text: '', br: true });
    else {
      const k = tag.toLowerCase();
      st[k] = Math.max(0, st[k] + (close ? -1 : 1));
    }
  }
  push(src.slice(last));
  return out;
}

export function hasRich(s) {
  return /<\/?(b|i|sup|sub|u)>|<u\s|<\/?color[\s>]|<br\s*\/?>|<li[\s>]|\n/i.test(String(s ?? ''));
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

// 1 つの段落 → HTML（段落の中の改行は <br>。brHTML で差し替えられる。srcURL は画像の線の画像の場所）
export function richToHTML(s, brHTML = '<br>', srcURL = (x) => x) {
  return parseRich(s)
    .map((g) => {
      if (g.br) return brHTML;
      let h = esc(g.text);
      if (g.sub) h = `<sub>${h}</sub>`;
      if (g.sup) h = `<sup>${h}</sup>`;
      if (g.i) h = `<i>${h}</i>`;
      if (g.b) h = `<b>${h}</b>`;
      if (g.u) h = underlineHTML(g.u, h, srcURL);
      if (g.color) h = `<span class="fc" style="color:#${g.color}">${h}</span>`;
      return h;
    })
    .join('');
}

// 下線。画像の線は、画像を文字の下に、縦横比を保たずに伸ばして敷く（行を折り返しても行ごとに引く）
export function underlineHTML(u, inner, srcURL = (x) => x) {
  if (u.type === 'double') return `<u class="ud">${inner}</u>`;
  if (u.type === 'img') {
    const url = String(srcURL(u.img) || '').replace(/'/g, '%27');
    return `<u class="ui" data-img="${esc(u.img)}" data-w="${u.w}" style="text-decoration:none;background-image:url('${esc(url)}');background-repeat:no-repeat;background-position:0 100%;background-size:100% ${u.w}pt;padding-bottom:${u.w}pt">${inner}</u>`;
  }
  return `<u>${inner}</u>`;
}

// 飾りを外した文字だけ（段落の中の改行も改行にする）
export function plainText(s) {
  return paragraphs(s)
    .map((p) => parseRich(p).map((g) => (g.br ? '\n' : g.text)).join(''))
    .join('\n');
}
