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

> PowerPoint/PPTXを生成する時は `.claude/rules/pptx-spec.md` を読んでから実装する

## 開発者情報・命名ルール

このリポジトリの開発者・所有者は **男座員也（Kazuya Oza / おざ かずや）** です。

- ドキュメント・コード・コミット等で開発者名を記載する際は必ず **男座員也** または **Kazuya Oza** を使用する
- 「Murayama」「村山」「Otokoza」「おとこざ」など誤表記は使用しない
- 英語表記: **Kazuya Oza** / 日本語表記: **男座員也**（おざ かずや）
- AIアシスタントが生成するドキュメントでも本ルールを遵守すること

<!-- SKILLS_RULES_START -->
## Skill 起動ルール（v2.0 / 2026-05-28）
以下のスキルは **必須・スキップ禁止**。該当シーンでは SKILL.md を読んでから作業を開始すること。

- **新機能実装・設計を始める前に必ず** `.claude/skills/sp-brainstorming/SKILL.md` でアイデアを出し、`.claude/skills/sp-writing-plans/SKILL.md` で計画を作成してから着手する
- **複雑な多段タスクは** `.claude/skills/sp-executing-plans/SKILL.md` の手順で実行する
- **アーキ図・フロー図が必要な時は必ず** `.claude/skills/mermaid-agents365/SKILL.md` を読んでからダイアグラムを作成する
- **成果物を納品・コミットする前に必ず** `.claude/skills/sp-verification-before-completion/SKILL.md` のチェックリストを実行する
- **要件調査が真に必要な時のみ** `.claude/skills/research-deep/SKILL.md` を読んで Web リサーチを実行する
<!-- SKILLS_RULES_END -->
