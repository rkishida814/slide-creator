// 長さの単位。キャンバスの単位（cm/mm/in/pt/px）で内部計算し、出力時に変換する。
export const PER_INCH = { in: 1, cm: 2.54, mm: 25.4, pt: 72, px: 96 };

export function assertUnit(unit) {
  if (!PER_INCH[unit]) throw new Error(`未対応の単位: ${unit}（cm, mm, in, pt, px のいずれか）`);
  return unit;
}

// 数値はキャンバスの単位とみなす。"5mm" "0.3cm" "12pt" のような文字列は変換する。
export function toUnit(value, unit) {
  if (value == null) return 0;
  if (typeof value === 'number') return value;
  const m = String(value).trim().match(/^(-?[\d.]+)\s*(cm|mm|in|pt|px)?$/);
  if (!m) throw new Error(`長さとして読めない値: ${value}`);
  const v = parseFloat(m[1]);
  const u = m[2] || unit;
  return (v / PER_INCH[u]) * PER_INCH[unit];
}

export function convert(value, from, to) {
  return (value / PER_INCH[from]) * PER_INCH[to];
}

// テキストのボックス（と、テキストを入れる前の空のボックス）で padding を書かないときの余白：上下左右に 0.5 cm
export const TEXT_PADDING = ['0.5cm', '0.5cm', '0.5cm', '0.5cm'];
export function defaultPadding(kind) {
  return kind === 'text' || kind === 'empty' ? TEXT_PADDING : null;
}

// padding: 数値 / [上下, 左右] / [上, 右, 下, 左]
export function parseBox(value, unit) {
  if (value == null) return { t: 0, r: 0, b: 0, l: 0 };
  if (!Array.isArray(value)) {
    const v = toUnit(value, unit);
    return { t: v, r: v, b: v, l: v };
  }
  const a = value.map((x) => toUnit(x, unit));
  if (a.length === 2) return { t: a[0], r: a[1], b: a[0], l: a[1] };
  if (a.length === 4) return { t: a[0], r: a[1], b: a[2], l: a[3] };
  throw new Error(`padding の形が不正: ${JSON.stringify(value)}`);
}
