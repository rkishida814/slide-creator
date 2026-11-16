# slide-creator

決まった型で図と文字を置いていく、スライド・ポスター作成ツール。PowerPoint の自由度をわざと下げ、
**「領域を上下か左右に分ける → 比率を決める → 各領域に何を入れるか、どう揃えるか」** を入れ子で書いた
設定ファイル（`layout.yaml`）を正本にする。人もAIも同じファイルを編集し、同じ CLI で確かめる。

- 出力：編集できる PPTX（画像は元画像のまま埋め込み、トリミングは PowerPoint の「トリミング」として残る）、PDF、PNG
- 状態：試作 v0.2（AI 用の CLI と、ブラウザ版の表示・編集画面）

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

## ブラウザ版

```bash
npm run serve        # http://127.0.0.1:8769/web/ を Chrome / Edge で開く
```

GitHub Pages に置けば、そのままブラウザで使える（ビルド不要の静的ファイル）。

- 「フォルダを開く」でプロジェクトフォルダを選ぶ。`layout.yaml` を直接読み書きする（Chrome / Edge のフォルダアクセス機能。Safari・Firefox は不可）。`layout.yaml` が無いフォルダなら雛形を作る
- 開発用：URL に `?demo` を付けると `examples/demo` を開く（保存はダウンロードになる。画面にボタンは無い）
- 左の欄の上の「スライド全体」で、1 枚の大きさ（幅・高さ・単位、縦横比とその固定、用紙）と全体の背景を決める。ボックスを選んでいるあいだは折りたたむ。PowerPoint の上限（1 辺 142.24 cm）を超えると注意を出す
- プレビューの上の帯：ページの選択（「1ページ」…）、「分割」（分割数と、上下か左右かを選ぶ）、ゴミ箱（選択中のボックスを削除）、虫眼鏡とつまみ（選択中のボックスを中心に 100〜800% で拡大）、「全体表示」（選択を外して全体を映す）
- ボックスにマウスを乗せると青く示す。クリックで選択する
- ドラッグすると、長方形になる範囲で複数のボックスを選ぶ（はみ出すボックスがあれば、そのボックスが入るまで範囲を広げる）。複数選んだときは、共通の外側の分割について、選んだ部分どうしの上下か左右の比率と、選んだ範囲全体の背景を設定する。分割の一部の部分に背景を付けると、その部分を 1 つの分割にまとめて背景を付ける
- ⌘（cmd。Windows では Ctrl）を押しながらドラッグして別のボックスの上で離すと、2 つの中身が入れ替わる（位置と大きさはそのまま）
- 1 つ選んだときは、大きさ（比率か長さ）、中身（画像か文字）、画像の収め方・位置・トリミング、文字（B・I・x²・x₂ で、選んだ範囲を太字・斜体・上付き・下付き）・大きさ・フォント（Helvetica・Arial・Times New Roman・Noto Sans JP）・色・見出しアイコン（フォルダの中の画像）・揃え方（横・縦）・行の高さ、内側の余白（上・右・下・左）、背景、枠線を変える。変えた内容は layout.yaml に書き込まれる（コメントは残る）
- 「画像を見ながらトリミング…」で、元画像の上で範囲をドラッグして決める
- 保存（⌘S）でフォルダに書き込む。「PPTX に書き出す」はフォルダの `out/` に書き出す。「PDF に書き出す」は印刷画面から「PDFに保存」を選ぶ

テンプレート（型）の中の枠を変えるときは、id があれば、そのスライドの `fill` に書き込む（そのスライドだけに効く）。
id が無ければ型の定義を書き換える（型を使う全ての箇所に効く）。後者のときは、設定欄に注意を出す。設定の誤り・文字のはみ出しの検査は、CLI の `slide check` で行う。

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
canvas: {width: 32, height: 18, unit: cm}   # 1枚の大きさ。unit は cm / mm / in / pt / px
defaults:                                          # 全体の既定（下の「引き継ぐ設定」）
  font: {size: 24, family: Helvetica}
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
| 文字 | `text: "1行目\n2行目"` | `font: {size, bold, italic, color, family}`、`line_height: 1.2`。文字の左に見出しアイコンを置くときは `icon: assets/x.png`（高さは `icon_size` で pt。省くと文字の大きさ）。文字の一部の飾りは `<b>太字</b>` `<i>斜体</i>` `<sup>上付き</sup>` `<sub>下付き</sub>`（例：`mm<sup>2</sup>`） |
| 空き | 何も書かない（`{}` や `~`） | |

どのノードにも書けるもの：

- `id`：名前。CLI で `--node photos` や `s1/photos` と指定できる。テンプレートの差し込み口にもなる
- `padding`：内側の余白。`1` / `[上下, 左右]` / `[上, 右, 下, 左]`
- `background`：色なら `"#RRGGBB"`。画像なら `{image: assets/x.png, mode: tile（並べる）| cover（広げる）, tile: 模様 1 つの幅, color: 下地の色}`。文書の一番上に書くと全スライドの背景になる（スライドごとの `background` が優先）
- `border: {width: 1 (pt), color: "#RRGGBB"}`

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
- 既定フォントは Helvetica（Windows の PowerPoint では Arial などに置き換わることがある）。和文は、PDF では Hiragino 等、PPTX では Office テーマの東アジアフォントで表示されるので、字形が一致しない。合わせたいときは `font.family` に両方にあるフォントを書く
- EXIF の回転情報付きの写真は、縦横を入れ替えて計算するが、PPTX での表示は未確認
- 図形・矢印・表・箇条書きの書式（字下げ）は未対応
- PPTX の「画像を並べる」背景は、画像を 96 dpi とみなして並べる倍率を決めている。PowerPoint で開いた見え方は未確認
- Noto Sans JP は、プレビューと PDF では Google Fonts から読み込む。PowerPoint では、その PC に入っているときだけ同じ字形になる
- ブラウザ版で設定欄から書き換えると、YAML の `{a: 1}` が `{ a: 1 }` のように書き直される（意味は同じ）

## 構成

```
src/core/   layout.js（配置計算）, resolve.js（テンプレート展開）, html.js（HTML描画）, units.js,
            ← ファイルに依存しない。ブラウザ版でも使う
            pptx.js（PPTX の組み立て）
src/node/   project.js（読み込み）, chrome.js（PDF/PNG/はみ出し測定）, pptx.js（PPTX の書き出し）
bin/slide.js  CLI
web/        ブラウザ版（app.js 画面, edit.js YAML の書き換え, source.js フォルダの読み書き, crop.js トリミング）
web/vendor/ ブラウザ用に同梱したライブラリ（npm run vendor で更新）
```
