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
- 「見本」は `examples/demo` を読むだけで開く（保存はダウンロードになる）。URL に `?demo` を付けても開く
- 左の欄の上の「スライド全体」で、1 枚の大きさを決める。幅・高さ・単位を直接入れるか、縦横比（16:9、4:3、A判・B判の縦横など）や用紙（A0・A1・B0・B1 の縦横、PowerPoint の標準）から選ぶ。「縦横比を固定する」をオンにすると、幅か高さの片方を変えたとき、もう片方も比を保って変わる。PowerPoint の上限（1 辺 142.24 cm）を超えると注意を出す
- プレビューの枠をクリックで選ぶ（いちばん内側の枠）。ダブルクリックか Enter で、その枠を拡大して表示。Esc で親の枠へ
- 左の欄の下の「選んだ枠」で、種類（空き・分割・格子・画像・文字）、分割の向き・数・長さの比・間隔、格子の行と列、画像の収め方・位置・トリミング、文字の大きさ・揃え方、余白・背景・枠線を変える。変えた内容は layout.yaml に書き込まれる（コメントは残る）
- 「上下に分ける」「左右に分ける」で、選んだ枠を、元の中身と空きの 2 つに分ける。「前と入替」「後と入替」で兄弟の順を入れ替える
- 「トリミング画面…」で、元画像の上で範囲をドラッグして決める（「枠の縦横比に合わせる」で、表示される範囲と一致させる）
- 保存（⌘S）でフォルダに書き込む。「PPTX」はフォルダの `out/` に書き出す。「PDF」は印刷画面から「PDFに保存」を選ぶ

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
