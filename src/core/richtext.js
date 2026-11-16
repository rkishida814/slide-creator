// 文字の飾りと改行。layout.yaml の text の中では、次だけを特別に読む（それ以外の < > はそのままの文字）。
//   <b>太字</b>  <i>斜体</i>  <sup>上付き</sup>  <sub>下付き</sub>
//   改行（\n）… 段落の区切り（Enter）。段落の間隔（para_space）が足される
//   <br>      … 段落の中の改行（Shift+Enter）。行間だけで、段落の間隔は足さない
// 例: "New bone area (mm<sup>2</sup>)"
const TAG = /<(\/?)(b|i|sup|sub)>|<br\s*\/?>/gi;

// 段落に分ける
export function paragraphs(s) {
  return String(s ?? '').split('\n');
}

// 1 つの段落 → [{ text, b, i, sup, sub, br }]（br は段落の中の改行）
export function parseRich(s) {
  const out = [];
  const st = { b: 0, i: 0, sup: 0, sub: 0 };
  let last = 0;
  const src = String(s ?? '');
  const push = (text) => {
    if (text) out.push({ text, b: st.b > 0, i: st.i > 0, sup: st.sup > 0, sub: st.sub > 0 });
  };
  src.replace(TAG, (m, close, tag, at) => {
    push(src.slice(last, at));
    if (!tag) out.push({ text: '', br: true });
    else {
      const k = tag.toLowerCase();
      st[k] = Math.max(0, st[k] + (close ? -1 : 1));
    }
    last = at + m.length;
    return m;
  });
  push(src.slice(last));
  return out;
}

export function hasRich(s) {
  return /<\/?(b|i|sup|sub)>|<br\s*\/?>|\n/i.test(String(s ?? ''));
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

// 1 つの段落 → HTML（段落の中の改行は <br>。brHTML で差し替えられる）
export function richToHTML(s, brHTML = '<br>') {
  return parseRich(s)
    .map((g) => {
      if (g.br) return brHTML;
      let h = esc(g.text);
      if (g.sub) h = `<sub>${h}</sub>`;
      if (g.sup) h = `<sup>${h}</sup>`;
      if (g.i) h = `<i>${h}</i>`;
      if (g.b) h = `<b>${h}</b>`;
      return h;
    })
    .join('');
}

// 飾りを外した文字だけ（段落の中の改行も改行にする）
export function plainText(s) {
  return paragraphs(s)
    .map((p) => parseRich(p).map((g) => (g.br ? '\n' : g.text)).join(''))
    .join('\n');
}
