# AI ROI Simulator v1 — AI導入ROIシミュレーター

> 企業のAI導入による投資収益率（ROI）をシミュレーション・可視化するツールです。

## 📋 概要

企業のAI導入による投資収益率（ROI）をシミュレーション・可視化するツールです。業務工数削減・品質向上・売上増加など複数の効果を定量化し、経営判断をデータドリブンにサポートします。

## ✨ 主な機能

- AI導入コスト（初期/運用）の自動試算
- 業務工数削減効果のシミュレーション
- ROI・回収期間・NPV の自動計算
- 感度分析によるシナリオ比較
- エグゼクティブサマリーの自動生成

## 🛠️ 技術スタック

| カテゴリ | 技術・ライブラリ |
|----------|----------------|
| 言語 | Python 3.10+ |
| UI | Streamlit |
| 計算エンジン | pandas, numpy |
| 可視化 | plotly |
| AI | Claude API（レポート生成） |

## 🚀 セットアップ

### 前提条件

- Python 3.9 以上
- APIキー（Claude / OpenAI 等）を `.env` ファイルに設定

### インストール

```bash
git clone https://github.com/KazuyaMurayama/AI-ROI-simulator_v1.git
cd AI-ROI-simulator_v1
pip install -r requirements.txt
```

### 環境設定

```bash
cp .env.example .env
# .env ファイルに必要なAPIキーを設定
```

## 💻 使い方

```bash
streamlit run app.py
```

## 👨‍💻 開発者情報

**男座員也（Kazuya Oza / おざ かずや）**

| | |
|---|---|
| GitHub | [@KazuyaMurayama](https://github.com/KazuyaMurayama) |
| 専門領域 | データサイエンス・生成AIコンサルタント |
| 主要スキル | Python, LightGBM, LangChain, RAG, Streamlit, React, TypeScript |
| 事業 | AIコンサルティング（月単価目標300万円）/ SaaS開発 / 定量投資 |

## 📄 ライセンス

© 2025 男座員也（Kazuya Oza）. All rights reserved.

---

> このリポジトリは **男座員也（Kazuya Oza）** が開発・管理しています。
> 命名・ドキュメント等での表記は必ず **男座員也** または **Kazuya Oza** を使用してください。
