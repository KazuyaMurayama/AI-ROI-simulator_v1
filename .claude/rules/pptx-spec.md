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
