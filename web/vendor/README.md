# 同梱ライブラリ

ブラウザ版がビルドなしで動くように、次のライブラリを npm のパッケージからそのまま写している（`npm run vendor` で更新）。

| ファイル | ライブラリ | 版 | ライセンス |
|---|---|---|---|
| pptxgen.bundle.js | PptxGenJS（JSZip を含む） | 4.0.1（JSZip 3.10.2） | MIT（LICENSE.pptxgenjs）、JSZip は MIT または GPLv3 から MIT を選択（LICENSE.jszip.md） |
| yaml/ | yaml | 2.9.1 | ISC（LICENSE.yaml） |
| heic-to.js | heic-to（HEIC の写真を JPEG に変える。中の libheif は 1.23.7。HEIC を使うときだけ読み込む） | 1.6.7 | LGPL-3.0（LICENSE.heic-to）。ビルドせずにそのまま同梱しているので、差し替えられる |
