// トリミング画面。元画像の上で範囲を動かし、[左, 上, 右, 下]（0〜1）を返す。
// 戻り値：配列（適用）/ 'clear'（外す）/ null（取り消し）
const $ = (s) => document.querySelector(s);
const r3 = (v) => Math.round(v * 1000) / 1000;

export function openCrop({ src, url, natural, crop, cellAspect, lockDefault }) {
  const dlg = $('#crop-dialog');
  const img = $('#crop-img');
  const rectEl = $('#crop-rect');
  const lock = $('#crop-lock');
  const values = $('#crop-values');
  $('#crop-src').textContent = `${src}（${natural.w} × ${natural.h} px）`;
  lock.checked = !!lockDefault;
  let [l, t, r, b] = crop;

  // 枠の縦横比（px の比）を、割合の空間での比に直す：(幅の割合 × W) / (高さの割合 × H) = cellAspect
  const fracAspect = () => (cellAspect * natural.h) / natural.w; // 幅の割合 / 高さの割合

  const draw = () => {
    const W = img.clientWidth;
    const H = img.clientHeight;
    Object.assign(rectEl.style, {
      left: `${l * W}px`,
      top: `${t * H}px`,
      width: `${(r - l) * W}px`,
      height: `${(b - t) * H}px`,
      backgroundImage: `url("${url}")`,
      backgroundSize: `${W}px ${H}px`,
      backgroundPosition: `${-l * W}px ${-t * H}px`,
    });
    const pw = Math.round((r - l) * natural.w);
    const ph = Math.round((b - t) * natural.h);
    values.textContent = `[${r3(l)}, ${r3(t)}, ${r3(r)}, ${r3(b)}]　${pw} × ${ph} px`;
  };

  // 縦横比を合わせる（中心を保ち、はみ出したら縮める）
  const applyLock = (anchor) => {
    if (!lock.checked) return;
    const A = fracAspect();
    let w = r - l;
    let h = b - t;
    if (w / h > A) w = h * A;
    else h = w / A;
    if (w > 1) { w = 1; h = w / A; }
    if (h > 1) { h = 1; w = h * A; }
    if (anchor) {
      // 角を動かしたときは、反対側の角を固定する
      const [ax, ay, sx, sy] = anchor;
      l = sx > 0 ? ax : ax - w;
      r = l + w;
      t = sy > 0 ? ay : ay - h;
      b = t + h;
    } else {
      const cx = (l + r) / 2;
      const cy = (t + b) / 2;
      l = cx - w / 2; r = cx + w / 2; t = cy - h / 2; b = cy + h / 2;
    }
    clampMove();
  };

  const clampMove = () => {
    const w = r - l;
    const h = b - t;
    if (l < 0) { l = 0; r = w; }
    if (r > 1) { r = 1; l = 1 - w; }
    if (t < 0) { t = 0; b = h; }
    if (b > 1) { b = 1; t = 1 - h; }
  };

  let drag = null;
  const onDown = (e) => {
    e.preventDefault();
    drag = { kind: e.target.dataset.h || 'move', x: e.clientX, y: e.clientY, start: [l, t, r, b] };
  };
  const onMove = (e) => {
    if (!drag) return;
    const W = img.clientWidth;
    const H = img.clientHeight;
    const dx = (e.clientX - drag.x) / W;
    const dy = (e.clientY - drag.y) / H;
    const [l0, t0, r0, b0] = drag.start;
    const min = 0.02;
    if (drag.kind === 'move') {
      [l, t, r, b] = [l0 + dx, t0 + dy, r0 + dx, b0 + dy];
      clampMove();
    } else {
      const k = drag.kind;
      l = k.includes('w') ? Math.min(Math.max(0, l0 + dx), r0 - min) : l0;
      r = k.includes('e') ? Math.max(Math.min(1, r0 + dx), l0 + min) : r0;
      t = k.includes('n') ? Math.min(Math.max(0, t0 + dy), b0 - min) : t0;
      b = k.includes('s') ? Math.max(Math.min(1, b0 + dy), t0 + min) : b0;
      if (lock.checked) {
        // 固定する角：動かしている角の反対側
        const ax = k.includes('w') ? r0 : l0;
        const ay = k.includes('n') ? b0 : t0;
        applyLock([ax, ay, k.includes('w') ? -1 : 1, k.includes('n') ? -1 : 1]);
      }
    }
    draw();
  };
  const onUp = () => (drag = null);

  rectEl.onpointerdown = onDown;
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
  lock.onchange = () => {
    applyLock(null);
    draw();
  };

  return new Promise((resolve) => {
    img.onload = () => {
      applyLock(null);
      draw();
    };
    img.src = url;
    if (img.complete && img.naturalWidth) draw();
    dlg.returnValue = '';
    dlg.showModal();
    // Enter は適用（form の既定では、最初のボタン「トリミングを外す」になってしまう）
    dlg.onkeydown = (e) => {
      if (e.key !== 'Enter' || e.isComposing) return;
      e.preventDefault();
      dlg.close('ok');
    };
    dlg.onclose = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      const v = dlg.returnValue;
      if (v === 'ok') {
        const full = l <= 0.0005 && t <= 0.0005 && r >= 0.9995 && b >= 0.9995;
        resolve(full ? 'clear' : [r3(l), r3(t), r3(r), r3(b)]);
      } else if (v === 'clear') resolve('clear');
      else resolve(null);
    };
  });
}
