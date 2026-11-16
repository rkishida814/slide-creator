# slide-creator

決まった型で図と文字を置いていく、スライド・ポスター作成ツール。PowerPoint の自由度をわざと下げ、
**「領域を上下か左右に分ける → 比率を決める → 各領域に何を入れるか、どう揃えるか」** を入れ子で書いた
設定ファイル（`layout.yaml`）を正本にする。人もAIも同じファイルを編集し、同じ CLI で確かめる。

- 出力：編集できる PPTX（画像は元画像のまま埋め込み、トリミングは PowerPoint の「トリミング」として残る）、PDF、PNG
- 状態：試作 v0.1（AI 用の CLI のみ。ブラウザ版は未着手）

## 準備

```bash
npm install
npm run demo          # 見本の画像を作り、検査し、examples/demo/out/ に PPTX と PDF を出す
```

PDF と PNG は Chrome で描く。見つからないときは環境変数 `SLIDE_CHROME` に実行ファイルの場所を入れる。

## プロジェクトの形

```
my-poster/
  layout.yaml     ← どこに何を置くか（正本）
  assets/         ← 画像（png / jpg / gif / svg。TIFF はブラウザで描けないので PNG にする）
  templates/      ← 使い回す型（任意。ファイル名が型の名前）
  out/            ← 書き出し先（自動）
```

## コマンド

```bash
node bin/slide.js init my-poster                       # 雛形を作る
node bin/slide.js tree my-poster                       # 木と各枠の座標・大きさ・crop・dpi
node bin/slide.js check my-poster                      # 誤り・解像度・文字のはみ出し（error があれば終了コード1）
node bin/slide.js render my-poster --slide s1 --outline  # PNG（--outline で枠とidを重ねる）
node bin/slide.js render my-poster --node photos --width 1200   # 特定の領域だけ PNG
node bin/slide.js export my-poster                     # PPTX と PDF（--pptx / --pdf で片方）
```

## layout.yaml の書き方

```yaml
canvas: {width: 33.867, height: 19.05, unit: cm}   # 1枚の大きさ。unit は cm / mm / in / pt / px
defaults:                                          # 全体の既定（下の「引き継ぐ設定」）
  font: {size: 18, family: Arial}
check: {min_dpi: 150}                              # これ未満の画像を警告
templates: {...}                                   # この文書の中だけで使う型（templates/ フォルダと同じ形）

slides:
  - id: s1
    background: "F4F6F8"
    root: <ノード>
```

### ノード

長さは数値ならキャンバスの単位。`"5mm"` のように単位付きでも書ける。文字の大きさだけは pt。

| 種類 | 書き方 | 意味 |
|---|---|---|
| 分割 | `split: rows` / `split: cols`、`children: [...]` | rows は上下に、cols は左右に並べる |
| | `ratio: [1, 2]` | 比率。`"3cm"` のような長さを混ぜると、その子は固定長になり、残りを比率で分ける。省略すると等分 |
| | `gap: 0.3` | 子と子の間隔 |
| 格子 | `grid: {rows: 2, cols: 3, gap: 0.2}`、`items: [...]` | 同じ大きさのセルに、行ごとに詰める（`order: col` で列ごと）。`gap: [行間, 列間]` も可 |
| 画像 | `image: assets/a.png` | |
| | `crop: [左, 上, 右, 下]` | 元画像に対する割合（0〜1）。`[0.1, 0, 0.9, 1]` は左右10%ずつ切る |
| | `fit: contain / cover / stretch` | contain は切らずに収める、cover は枠いっぱいにしてはみ出しを切る、stretch は伸ばす |
| 文字 | `text: "1行目\n2行目"` | `font: {size, bold, italic, color, family}`、`line_height: 1.2` |
| 空き | 何も書かない（`{}` や `~`） | |

どのノードにも書けるもの：

- `id`：名前。CLI で `--node photos` や `s1/photos` と指定できる。テンプレートの差し込み口にもなる
- `padding`：内側の余白。`1` / `[上下, 左右]` / `[上, 右, 下, 左]`
- `background: "#RRGGBB"`、`border: {width: 1 (pt), color: "#RRGGBB"}`

### 引き継ぐ設定

`align`（left / center / right）、`valign`（top / middle / bottom）、`fit`、`font`、`line_height` は、
親に書くと子孫すべてに効く（子で上書きできる）。たとえば格子に `fit: cover` と書けば全セルに効く。
どこにも書いていなければ、画像は中央・中段、文字は左・上に揃う。

### テンプレート（使い回し）

```yaml
templates:
  photos_and_note:          # 中身を空けた型
    split: rows
    ratio: ["2.4cm", 1]
    children:
      - {id: title, text: ""}
      - id: photos
        grid: {rows: 2, cols: 3}
        items: []

slides:
  - id: s1
    use: photos_and_note    # 型を使う
    fill:                   # id ごとに中身を入れる（その id のノードの設定を上書き）
      title: {text: "結果"}
      photos: {items: [{image: assets/a1.png}, {image: assets/a2.png}]}
```

`use` はノードの途中にも書ける（型の中に型を入れてもよい）。`use` と `fill` 以外のキー（`padding` など）は型の根を上書きする。

## AI が使うとき

1. `layout.yaml` を直接編集する（専用の編集コマンドは無い。テキストなので差分も読める）
2. `slide tree` で数値を、`slide render --outline` で見た目を確かめる。PNG は画像として読む
3. `slide check` で error が無いことを確かめてから `slide export`

`tree --json` は全要素の絶対座標を JSON で出す（整列の検査ツールなどに渡せる）。

## 既知の制約

- PPTX を PowerPoint で手直ししても、`layout.yaml` には戻らない。正本は常に設定ファイル
- 文字の折り返しとはみ出しの判定は Chrome で測る。PowerPoint とはフォントの寸法が少し違うため、ぎりぎりの枠では PPTX 側だけはみ出すことがある
- 既定フォントは Arial。和文は、PDF では Hiragino 等、PPTX では Office テーマの東アジアフォントで表示されるので、字形が一致しない。合わせたいときは `font.family` に両方にあるフォントを書く
- EXIF の回転情報付きの写真は、縦横を入れ替えて計算するが、PPTX での表示は未確認
- 図形・矢印・表・箇条書きの書式（字下げ）・文字の一部だけの太字は未対応

## 構成

```
src/core/   layout.js（配置計算）, resolve.js（テンプレート展開）, html.js（HTML描画）, units.js
            ← ファイルに依存しない。ブラウザ版でも使う
src/node/   project.js（読み込み）, chrome.js（PDF/PNG/はみ出し測定）, pptx.js（PPTX）
bin/slide.js  CLI
```
