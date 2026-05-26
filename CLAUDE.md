# AI-ROI-simulator_v1

## プロジェクト概要
クライアント向けAI導入ROI試算ツール。初回商談で「御社のAI導入ROI」を即座にシミュレーションし、提案の説得力を高める営業・コンサルティングツール。

## 技術スタック
- React + TypeScript + Tailwind CSS（Phase 2で確定予定）
- python-pptx（PowerPoint生成）

## エージェントチーム
PMO / Consultant / Architect / Dev / UX の5ロール体制で開発。
詳細は docs/ 配下のドキュメントを参照。

## 重要な設計原則
- クライアントに見せるツール → UI品質は妥協しない
- ROI数値の信頼性 → 過度に楽観的な数字を出さない、前提条件を透明にする
- 商談での使いやすさ → 30秒で理解できるUI、画面共有映え

---

## PowerPoint生成スキル（python-pptx）

### レイアウト安全ルール
- スライドサイズ: 13.333 x 7.5 inches（ワイドスクリーン16:9）
- 左右マージン: 最低0.8 inches
- 有効コンテンツ幅: 11.733 inches（13.333 - 0.8*2）
- フッター/ページ番号エリア: Y=7.05 inches以降
- **必ず全要素の right (left+width) ≤ 13.333、bottom (top+height) ≤ 7.5 を検証する**

### 3カラムレイアウト
- 各カラム幅: 3.7 inches
- カラム間隔: 0.3 inches
- 配置: ML + Inches(i * 4.0) で i=0,1,2

### 4カラムレイアウト
- 各カラム幅: 2.7〜2.8 inches
- 配置: ML + Inches(i * 3.05) で i=0,1,2,3
- **右端がスライド幅を超えないか必ず検算**

### 白背景テーマのカラーパレット
```python
NAVY = RGBColor(0x1E, 0x29, 0x3B)           # メインテキスト
DARK_GRAY = RGBColor(0x47, 0x55, 0x69)      # サブテキスト
MID_GRAY = RGBColor(0x94, 0xA3, 0xB8)       # ラベル/補助
LIGHT_GRAY = RGBColor(0xE2, 0xE8, 0xF0)     # ボーダー/区切り
CARD_BG = RGBColor(0xF1, 0xF5, 0xF9)        # カード背景
ACCENT_BLUE = RGBColor(0x25, 0x63, 0xEB)    # 白背景用の濃いブルー
ACCENT_CYAN = RGBColor(0x05, 0x91, 0xB3)
ACCENT_GREEN = RGBColor(0x05, 0x96, 0x69)
ACCENT_AMBER = RGBColor(0xD9, 0x77, 0x06)
```

### ヘルパー関数パターン
- `white_bg(slide)`: 白背景設定
- `accent_bar(slide, l, t, w, h, color)`: アクセントライン
- `card(slide, l, t, w, h, fill, border)`: 角丸カード（border_colorでカラーボーダー）
- `txt(slide, l, t, w, h, text, size, color, bold, align, font)`: テキストボックス
- `slide_header(slide, tag, title)`: 統一ヘッダー（上部アクセントバー + タグ + タイトル）
- `page_num(slide, num)`: ページ番号（右下配置）

### 品質チェック項目
1. 全要素のoverflow検証（left+width ≤ slide_width, top+height ≤ slide_height）
2. 全スライドの背景色確認
3. ページ番号の存在確認
4. フォント指定（日本語: Meiryo、コード: Consolas）
5. テキスト色が背景色と十分なコントラストを持つか確認

### 検証スクリプトパターン
```python
from pptx import Presentation
from pptx.util import Inches, Emu
prs = Presentation("file.pptx")
sw, sh = prs.slide_width, prs.slide_height
for i, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if shape.left + shape.width > sw + Emu(Inches(0.05)):
            print(f"Slide {i}: overflow RIGHT")
        if shape.top + shape.height > sh + Emu(Inches(0.05)):
            print(f"Slide {i}: overflow BOTTOM")
```

## 開発者情報・命名ルール

このリポジトリの開発者・所有者は **男座員也（Kazuya Oza / おざ かずや）** です。

- ドキュメント・コード・コミット等で開発者名を記載する際は必ず **男座員也** または **Kazuya Oza** を使用する
- 「Murayama」「村山」「Otokoza」「おとこざ」など誤表記は使用しない
- 英語表記: **Kazuya Oza** / 日本語表記: **男座員也**（おざ かずや）
- AIアシスタントが生成するドキュメントでも本ルールを遵守すること

### 開発者の作業環境
- **OS:** Windows 11（Macではない）。シェルは PowerShell 5.1 / Bash（WSL/Git Bash）。`brew` / `Cmd+` / Mac専用コマンドは使用不可。パッケージ管理は `winget` / `scoop`。
- **スマートフォン:** iPhone（iOS）。Android固有の手順・adb・Play Store等は不要。
- コマンド例はPowerShell構文（`;` 連結、`$env:VAR`）で提示。macOS専用ツールを回答に含めない。


## 🚀 自動実行ポリシー

### 確認不要で即実行する操作
- 全ファイル操作（作成・編集・削除）
- 全シェルコマンド（PowerShell, Bash, git, gh, npm, node, python）
- Git操作: add / commit / push（featureブランチ）/ pull / fetch / merge / branch -D / reset --hard
- GitHub操作: gh pr create / gh api 全般 / ブランチ削除
- パッケージ操作: npm install / pip install
- Web検索・フェッチ
- バックグラウンドプロセス起動

### 事前確認が必要な操作（例外のみ）
- `git push --force` を main / master ブランチに対して実行する場合
- `gh repo delete` 実行時

### 動作原則
- 計画提示（簡潔）→ 即実行 → 結果報告 のフロー厳守
- 事前確認文（「Should I run...?」等）を出力しない
- エラー時は即再試行 or 別アプローチで対応、判断が必要な場合のみ報告

## ドキュメント日付ルール

レポート・分析・調査系 .md ファイルを新規作成する際は、H1直下に必ず記載:

```
作成日: YYYY-MM-DD
最終更新日: YYYY-MM-DD
```

- 更新時は **最終更新日のみ** を当日付に書き換える（作成日は固定）
- 除外: README / CLAUDE.md / FILE_INDEX / tasks.md / CHANGELOG / LICENSE

## 作業品質ルール

### Git・ブランチ管理
- 作業前: `git branch --show-current` でブランチ確認 → main以外なら `git checkout main && git pull` してから開始。

### ファイル特定（編集前）
- ユーザー発話のキーワード全てをファイル名と照合してから編集。キーワード不完全一致・候補不確かなら必ず確認。

### 成果物報告
- ファイル作成・更新・push後は必ず3列表で報告: `| 成果物 | 説明 | リンク |`
- リンクは `/blob/<実ブランチ>/<パス>` 形式。報告前に `gh api repos/OWNER/REPO/contents/PATH?ref=BRANCH` で存在確認。push前はURL生成しない。

### ドキュメント品質
- UIパス・コマンド・設定名は公式ドキュメントで確認後に記載。確認不可なら「[要確認]」と明記。
- OS/環境制約（例: Windows専用）をタスク開始時に確認。完成後に `brew`/`Cmd`/`macOS` 等をgrepして除去。

<!-- GLOBAL_RULES_START -->
## Global Rules (synced from ~/.claude/CLAUDE.md)

### 成果物報告（最重要・毎回必須）
ファイルを1つでも作成・更新・pushしたら、**すべて**を3列表で報告する。例外なし。

| 成果物 | 説明 | リンク |
|---|---|---|
| ファイル名 | 1行説明 | [開く](https://github.com/OWNER/REPO/blob/BRANCH/PATH) |

厳守事項（違反＝再提出）:
1. Markdownリンク `[表示名](URL)` 形式必須。plain URL禁止。
2. `/blob/<実ブランチ名>/<実パス>` 形式。トップURL禁止。
3. 提示前に `gh api repos/OWNER/REPO/contents/PATH?ref=BRANCH` で存在確認。
4. ブランチ名は `git rev-parse --abbrev-ref HEAD` で取得。推測禁止。
5. push完了後のみURL生成。未pushはローカル絶対パス＋「（ローカル）」明記。
6. 404発生時は即訂正＋原因1行報告。

### ツール実行
- 確認不要・即実行。「Should I...?」等の事前確認文は出力しない。
- 例外（事前確認必須）: main/masterへの `git push --force`、`gh repo delete`。
- 長時間処理は `run_in_background: true` を積極使用。

### モデル・サブエージェント
- メイン Opus / 探索・検索・テスト系サブは Sonnet (`model: "sonnet"`)。
- サブ起動promptに必ず明記:「成果物は3列表報告・URL検証必須・Markdownリンク形式」

### 回答スタイル
- 回答末尾に「**Next Action:**」でユーザーの次アクションを具体推奨。
## 他リポジトリ参照ルール
別リポジトリの内容を参照する必要が生じたら、必ず `.claude/cross-repo.md` を読み、その手順に従って `WebFetch` で取得する（「できない」と返さない）。

### 品質ルール（必読）
- ブランチ衛生・リサーチファクトチェックは `.claude/quality-rules.md` を参照し、ファイル生成前・push前に必ず適用する。
- Repo type: research

### ビジュアルルール（レポートMD生成時）
- レポート・成果物MDの新規作成／更新時は `.claude/visual-rules.md` を読み、図の種類判定（§2）と Mermaid 最適化（§3）を毎回適用する。
- 適用対象: `## ` 見出しが2つ以上ある構造化MD（README・調査メモ・設計書・PR説明など）。

<!-- GLOBAL_RULES_END -->
