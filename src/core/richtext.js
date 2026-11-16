// 文字の飾り。layout.yaml の text の中では、次のタグだけを飾りとして読む（それ以外の < > はそのままの文字）。
//   <b>太字</b>  <i>斜体</i>  <sup>上付き</sup>  <sub>下付き</sub>
// 例: "New bone area (mm<sup>2</sup>)"
const TAG = /<(\/?)(b|i|sup|sub)>/gi;

// → [{ text, b, i, sup, sub }]（改行は text の中に \n のまま残る）
export function parseRich(s) {
  const out = [];
  const st = { b: 0, i: 0, sup: 0, sub: 0 };
  let last = 0;
  const push = (text) => {
    if (text) out.push({ text, b: st.b > 0, i: st.i > 0, sup: st.sup > 0, sub: st.sub > 0 });
  };
  String(s ?? '').replace(TAG, (m, close, tag, at) => {
    push(s.slice(last, at));
    const k = tag.toLowerCase();
    st[k] = Math.max(0, st[k] + (close ? -1 : 1));
    last = at + m.length;
    return m;
  });
  push(String(s ?? '').slice(last));
  return out;
}

export function hasRich(s) {
  return /<\/?(b|i|sup|sub)>/i.test(String(s ?? ''));
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);

export function richToHTML(s) {
  return parseRich(s)
    .map((g) => {
      let h = esc(g.text);
      if (g.sub) h = `<sub>${h}</sub>`;
      if (g.sup) h = `<sup>${h}</sup>`;
      if (g.i) h = `<i>${h}</i>`;
      if (g.b) h = `<b>${h}</b>`;
      return h;
    })
    .join('');
}

// 飾りを外した文字だけ
export function plainText(s) {
  return parseRich(s).map((g) => g.text).join('');
}
