# FILE_INDEX.md — AI-ROI-simulator_v1

> **新セッション開始時に必ずこのファイルを読む。**
> ファイル追加・削除・移動時は必ずこのファイルを更新すること。
> 最終更新: 2026-04-30

## 概要
AI導入ROIシミュレーターの設計・要件定義・PPTX生成スクリプト。

**スタック:** Python, Markdown, PPTX

---

## 📋 最初に読むべきファイル

| 優先度 | ファイル | 内容 |
|---|---|---|
| ★★★ | `CLAUDE.md` | 運用ルール |
| ★★★ | `docs/requirements.md` | 要件定義 |
| ★★★ | `generate_pptx_v2.py` | PPTXプレゼン生成スクリプト |
| ★★ | `docs/roi-model-design.md` | ROIモデル設計 |
| ★★ | `AI-ROI-Simulator_Guide.pptx` | 完成プレゼン資料 |

---

## 🗂️ ディレクトリ構造

```
AI-ROI-simulator_v1/
├── CLAUDE.md
├── generate_pptx_v2.py          ← PPTX生成スクリプト
├── AI-ROI-Simulator_Guide.pptx  ← 完成プレゼン資料
└── docs/
    ├── discussion-log.md
    ├── requirements.md          ← 要件定義
    └── roi-model-design.md      ← ROIモデル設計
```

---

## 📑 全ファイル一覧

| パス | 種別 | 説明 |
|---|---|---|
| `CLAUDE.md` | ドキュメント | 運用ルール |
| `generate_pptx_v2.py` | Python | PPTXプレゼン生成スクリプト（v2） |
| `AI-ROI-Simulator_Guide.pptx` | 資料 | 完成プレゼン資料 |
| `docs/requirements.md` | ドキュメント | 要件定義 |
| `docs/roi-model-design.md` | ドキュメント | ROIモデル設計 |
| `docs/discussion-log.md` | ドキュメント | 設計討論ログ |

---

## 🔖 ファイル更新ルール

1. 新ファイル追加時: 該当セクションに1行追加
2. ファイル削除・移動時: 該当行を削除または更新
3. 更新後: `git add FILE_INDEX.md && git commit -m "docs: FILE_INDEX.md更新"`
