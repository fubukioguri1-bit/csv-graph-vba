# CSV → グラフ作成ツール（Excel VBA版）

## 概要
- 入力: 任意の CSV（UTF‑8）  
- 処理: 「参加者氏名」列を自動削除 → 残りの数値でグラフ作成 → PNG 出力 & シート貼り付け  

## ディレクトリ構成
csv-graph-vba/
├─ .github/
│ └─ workflows/
│ └─ ci.yml # CI (markdownlint + mermaid + VBA test)
├─ docs/ # 画像や外部資料
├─ src/
│ └─ modules/
│ ├─ ModMain.bas
│ ├─ ModCSV.bas
│ ├─ ModChart.bas
│ └─ ModUtil.bas
├─ sample/
│ └─ participants.csv
├─ 0_overview/
│ ├─ 0.1_background.md
│ ├─ 0.2_goals.md
│ └─ 0.3_stakeholders.md
├─ 1_requirements/
│ ├─ 1.1_functional.md
│ └─ 1.2_nonfunctional.md
├─ 2_architecture/
│ ├─ 2.1_system-context.mermaid.md
│ └─ 2.2_component-diagram.mermaid.md
├─ 3_data/
│ ├─ 3.1_entity-relationship.md
│ └─ 3.2_schema.md
├─ 4_operations/
│ └─ 4.1_ci-cd.md
├─ 5_testing/
│ └─ 5.1_test-strategy.md
├─ 6_risk-management/
│ └─ 6.1_risk-register.md
└─ CHANGELOG.md

## 使い方
1. `src/modules/` の `.bas` ファイルを Excel の VBE に **Import**  
2. 「Data」シートを作成し、リボンかシートボタンからマクロを実行  
3. `graph.png` が生成され、同シートに貼り付けられます  
## ライセンス
MIT
