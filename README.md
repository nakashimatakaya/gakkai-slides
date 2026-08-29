# gakkai-slides

学会発表スライド（pptx）を作るための Claude Code / Codex / claude.ai 用スキルです。抄録や論文を渡して「学会スライドにして」と頼むと、発表時間から枚数を設計し、アウトラインの承認を取ってから pptx を生成します。

デザインは決め打ちです。白背景、黒の箇条書き、強調だけ赤、フォントは Noto Sans JP。塗りつぶしのbox、角丸カード、アイコン、グラデーション、タイトル下のアクセント線は作りません（AI生成スライドの典型なので）。本文は1枚につき1つのテキスト枠にまとめ、タイトル以外の全スライド右下にページ番号が入ります。COIスライドは2枚目に自動で入ります。

## 入れ方

Claude Code / Codex（ターミナルで1行）

```bash
npx skills add nakashimatakaya/gakkai-slides -g -a claude-code -a codex
```

claude.ai は、このページの Code → Download ZIP を「カスタマイズ」→「スキル」からアップロードしてください。

生成スクリプトを使うには python-pptx が要ります。

```bash
pip3 install python-pptx
```

フォントについて。macOS に Noto Sans JP は入っていません。Google Fonts（fonts.google.com/noto/specimen/Noto+Sans+JP）からダウンロードして Font Book でインストールしてください。入れない場合は生成時に `--font "ヒラギノ角ゴシック"` で代替できます。

## 使い方

Claude Code なら `/gakkai-slides`、Codex なら `$gakkai-slides` を付けて、抄録・論文・メモを貼り、発表時間と学会名を伝えます。

```
/gakkai-slides 添付の抄録から7分の口演スライドを作って。第XX回九州麻酔科学会、16:9、COIなし。
```

Claude はまずアウトライン（Markdown）を出します。数値の間違いはここで直すのがいちばん楽です。承認するとpptxが生成され、検査結果と一緒に返ってきます。

スクリプトを直接使うこともできます。

```bash
python3 ~/.agents/skills/gakkai-slides/scripts/build_deck.py outline.md -o deck.pptx
python3 ~/.agents/skills/gakkai-slides/scripts/check_deck.py deck.pptx
```

check_deck.py は自分で作った既存のスライドにも使えます。box、ページ番号の欠落、フォント、パレット外の色、16pt未満の文字、1文1ボックスの分割を報告します。

## アウトラインの書式

```markdown
# 演題名
演者: 中島誉也、共同演者A
所属: 長崎大学病院 麻酔集中治療部
学会: 第XX回 〇〇学会学術集会
日付: 2026年9月X日
発表時間: 7分
COI: なし
比率: 16:9

## 背景と目的
- 箇条書き（**で囲むと赤太字**）
  - スペース2つで下位項目
![](figures/fig1.png)
主張: このスライドで言いたいこと1行
出典: 1) Russell JA, et al. N Engl J Med. 2008.
```

サンプルが `evals/samples/outline_sample.md` にあります（数値はダミー）。

## うまくいかないとき

- `ModuleNotFoundError: pptx` → `pip3 install python-pptx` を実行
- 生成したpptxのフォントが明朝や中華フォントに置き換わる → Noto Sans JP が未インストール。Google Fonts から入れるか `--font` で代替
- 会場が4:3指定 → アウトラインの `比率: 4:3` で作り直す
- COIの文面が学会指定と違う → 学会の様式（PPTダウンロード）で2枚目を差し替える。日本麻酔科学会系は口演スライド2枚目への開示が義務
- 提出先で崩れる → PowerPoint の「オプション → 保存 → ファイルにフォントを埋め込む」か、PDFで提出

## 出典

ガレメン（Garage members）「研究発表スライドの作り方」、ドキュメントプラス「分かりやすい学会発表のスライドの作り方」、にゅ〜ろろぐ、Antaa Slide（学会のフォント指定調査）、日本麻酔科学会・日本医学会のCOI開示規定、Anthropic skills の pptx スキル。詳細は references/ を参照。

## ライセンス

MIT
