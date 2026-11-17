#!/usr/bin/env bash
# ==============================================================================
# init‑csv‑graph‑vba.sh
# 目的: 「CSV → 氏名列削除 → グラフ作成」ツール用の Git リポジトリ構成を自動生成
# ==============================================================================
set -euo pipefail

# ------------------------------ 1️⃣ 基本設定 ------------------------------
# ★ 必要に応じて以下変数を書き換えてください
REPO_NAME="csv-graph-vba"                         # 作成したいリポジトリ名
REMOTE_URL="fubukioguri1@1bit:https://github.com/fubukioguri1-bit/csv-graph-vba.git"                                      # 例: git@github.com:yourorg/csv-graph-vba.git
#   空のままにするとローカルだけで完了。GitHub に自動作成したい場合は下の「GitHub CLI」セクション参照

# ------------------------------ 2️⃣ ディレクトリ作成 ------------------------------
mkdir -p "${REPO_NAME}"
cd "${REPO_NAME}"

# ---------- 2‑1. .gitignore ----------
cat > .gitignore <<'EOF'
# OS が自動生成するファイル
.DS_Store
Thumbs.db

# Node / npm 関連
node_modules/
npm-debug.log*

# VBA / Excel 生成物
*.xlsm
*.xlam
*.xlsx
*.xlsb
*.xls

# 画像・PDF 出力（自動生成物は Git 管理しない方が楽です）
*.png
*.jpg
*.pdf

# ログ
error.log
EOF

# ---------- 2‑2. README ----------
cat > README.md <<'EOF'
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
EOF
# ---------- 2‑3. 空の CHANGELOG ----------
cat > CHANGELOG.md <<'EOF'
# CHANGELOG
All notable changes to this project will be documented in this file.
## [Unreleased]
- 初期リポジトリ作成
EOF
# ---------- 2‑4. 0_overview ----------
mkdir -p 0_overview
cat > 0_overview/0.1_background.md <<'EOF'
# 0.1 背景
社内の勉強会や研修では参加者リストが CSV で配布されますが、個人情報保護のため氏名列はレポートに含めたくありません。  
本ツールは CSV から氏名列を自動除去し、残りの数値データでグラフ画像（PNG）を生成し、Excel シートに貼り付けます。
EOF
cat > 0_overview/0.2_goals.md <<'EOF'
# 0.2 Goals（目標）
| ビジネスゴール | 定量指標 |
|----------------|-----------|
| 個人情報漏洩リスクを 0 に | 氏名列が出力・画像に残らない |
| グラフ作成工数を 80% 削減 | 手動 5 分 → 1 分 |
| エラーハンドリング品質向上 | エラーログが 0 件で完了 |
EOF
cat > 0_overview/0.3_stakeholders.md <<'EOF'
# 0.3 ステークホルダー
- **プロダクトオーナー**: 研修運営チーム
- **開発者**: データ分析部 VBA エンジニア
- **利用者**: 研修講師・運営スタッフ
- **情報セキュリティ部**: 個人情報保護ポリシー確認
EOF
# ---------- 2‑5. 1_requirements ----------
mkdir -p 1_requirements
cat > 1_requirements/1.1_functional.md <<'EOF'
# 1.1 Functional Requirements
| 番号 | 機能 | 入力 | 出力 | 優先度 |
|------|------|------|------|--------|
| FR‑001 | CSV 読み込み | 任意の UTF‑8 CSV | Data シートにテーブル | 高 |
| FR‑002 | 氏名列削除 | 読み込んだテーブル | 氏名列が除去されたテーブル | 高 |
| FR‑003 | グラフ作成 | 数値列（1列以上） | PNG 画像 `graph.png` | 高 |
| FR‑004 | 画像貼り付け | PNG | 同シートに画像オブジェクト | 中 |
| FR‑005 | エラー処理 | 例外時 | メッセージボックス + `error.log` | 高 |
EOF
cat > 1_requirements/1.2_nonfunctional.md <<'EOF'
# 1.2 Non‑Functional Requirements
| 項目 | 要件 |
|------|------|
| NFR‑001 可用性 | ローカルで 99.9% 起動成功率 |
| NFR‑002 性能 | 100 行程度の CSV を ≤ 2 秒で処理 |
| NFR‑003 対応 Excel バージョン | Excel 2019 以降、または Microsoft 365 |
| NFR‑004 セキュリティ | 外部通信は行わない |
| NFR‑005 保守性 | VBA はモジュール単位に分割、コメント必須 |
EOF
# ---------- 2‑6. 2_architecture ----------
mkdir -p 2_architecture
cat > 2_architecture/2.1_system-context.mermaid.md <<'EOF'
# System Context Diagram
```mermaid
flowchart LR
    User[ユーザー] -->|CSV選択| VBA_Tool[VBA マクロ]
    VBA_Tool -->|内部処理| ModCSV[CSV 読み込み・列削除]
    ModCSV --> ModChart[グラフ作成 & PNG出力]
    ModChart -->|画像| Sheet[Excel シート]
    Sheet -->|結果| User
EOF

cat > 2_architecture/2.2_component-diagram.mermaid.md <<'EOF'

Component Diagram
graph TD
    subgraph UI
        Btn[リボンボタン] --> Run[RunCsvGraphTool]
    end
    subgraph Core
        Run --> LoadCSV[ModCSV.LoadAndCleanCsv]
        LoadCSV --> CreateChart[ModChart.CreateChartAndExport]
        CreateChart --> InsertImg[ModChart.InsertImage]
    end
    style UI fill:#f9f,stroke:#333,stroke-width:2px
    style Core fill:#bbf,stroke:#333,stroke-width:2px
EOF

cat > 2_architecture/2.5_implementation.md <<'EOF'

2.5 Implementation – VBA
4.1 ディレクトリ構成（Git 管理）
src/
├─ modules/
│ ├─ ModMain.bas ← エントリーポイント
│ ├─ ModCSV.bas ← CSV 読み込み・氏名列削除
│ ├─ ModChart.bas ← グラフ作成・PNG出力・画像貼り付け
│ └─ ModUtil.bas ← ロギング・エラーハンドリング
└─ README.md ← 各モジュール概要

4.2 主なコード（抜粋）
ModMain.bas
vba

Option Explicit
Sub RunCsvGraphTool()
    On Error GoTo ErrHandler
    Dim csvPath As String
    csvPath = Application.GetOpenFilename("CSV Files (*.csv),*.csv")
    If csvPath = "False" Then Exit Sub
    Call LoadAndCleanCsv(csvPath)   ' ModCSV
    Call CreateChartAndExport()     ' ModChart
    MsgBox "グラフ画像が作成されました。", vbInformation
    Exit Sub
ErrHandler:
    Call LogError(Err.Number, Err.Description)
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
ModCSV.bas
vba

Option Explicit
Public Sub LoadAndCleanCsv(ByVal fPath As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data")
    ws.Cells.Clear
    With ws.QueryTables.Add(Connection:="TEXT;" & fPath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    Dim colIdx As Long
    colIdx = Application.Match("Name", ws.Rows(1), 0)
    If IsError(colIdx) Then colIdx = Application.Match("氏名", ws.Rows(1), 0)
    If IsError(colIdx) Then Err.Raise vbObjectError + 1000, "ModCSV", "氏名列が見つかりません"
    ws.Columns(colIdx).Delete
End Sub
ModChart.bas
vba

Option Explicit
Public Sub CreateChartAndExport()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data")
    Dim chartObj As ChartObject
    Dim rng As Range: Set rng = ws.UsedRange
    Set chartObj = ws.ChartObjects.Add(300, 10, 500, 300)
    With chartObj.Chart
        .SetSourceData Source:=rng
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Data Overview"
        .Export Filename:=ThisWorkbook.Path & "\graph.png", FilterName:="PNG"
    End With
    ws.Pictures.Insert ThisWorkbook.Path & "\graph.png"
End Sub
ModUtil.bas
vba

Option Explicit
Public Sub LogError(ByVal errNo As Long, ByVal errMsg As String)
    Dim logPath As String: logPath = ThisWorkbook.Path & "\error.log"
    Dim f As Integer: f = FreeFile
    Open logPath For Append As #f
    Print #f, Now & " - Error " & errNo & ": " & errMsg
    Close #f
End Sub
4.3 UI 連携
リボンボタン もしくは シート上のフォームボタン に RunCsvGraphTool を割り当てます。
実行すると Data シートにクリーンデータが、同シートに graph.png が貼り付けられます。
EOF
---------- 2‑7. 3_data ----------
mkdir -p 3_data
cat > 3_data/3.1_entity-relationship.md <<'EOF'

3.1 Entity Relationship

erDiagram
    CSV {
        string ParticipantID PK
        string Name
        int Score
        int Hours
    }
    CleanTable {
        string ParticipantID PK
        int Score
        int Hours
    }
    CSV ||--|| CleanTable : transforms
EOF

cat > 3_data/3.2_schema.md <<'EOF'

3.2 CSV スキーマ（サンプル）

列名	データ型	説明
ParticipantID	文字列	ユニーク ID
Name	文字列	削除対象
Score	整数	評価点数
Hours	整数	受講時間
EOF		
---------- 2‑8. 4_operations ----------
mkdir -p 4_operations
cat > 4_operations/4.1_ci-cd.md <<'EOF'

4.1 CI/CD
GitHub Actions ワークフロー (.github/workflows/ci.yml)
yaml

name: CI

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - name: Install markdownlint
        run: npm install -g markdownlint-cli
      - name: Lint markdown
        run: markdownlint "**/*.md"

  mermaid:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - name: Install mermaid-cli
        run: npm install -g @mermaid-js/mermaid-cli
      - name: Validate mermaid diagrams
        run: |
          find . -name "*.md" -exec grep -l "mermaid" {} \; | \
          xargs -n1 mmdc -i {} -o /dev/null

  vba-test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      # ここでは Rubberduck の CLI がインストール済みである前提です
      - name: Run VBA unit tests (Rubberduck)
        run: |
          rubby duck run-tests -ProjectPath src/modules
デプロイ戦略
完成した .xlsm（または .xlam）は GitHub Releases にバイナリ添付。
README.md に Import 手順と「リボンにボタンを追加する方法」を記載。
EOF
# ---------- 2‑9. 5_testing ----------
mkdir -p 5_testing
cat > 5_testing/5.1_test-strategy.md <<'EOF'
# 5.1 Test Strategy
## テストレベル
| レベル | 手法 | 主なチェック項目 |
|--------|------|------------------|
| Unit   | Rubberduck VBA Unit Test | `ModCSV.LoadAndCleanCsv` が氏名列を正しく削除できるか |
| Integration | 手動シナリオ (ボタン実行) | CSV → データシート → PNG が生成・貼り付けられるか |
| UI     | ユーザーテストシナリオ | エラーメッセージが適切に表示されるか |
## Rubberduck テスト例（`src/modules/TestModCSV.bas`）
```vba
Option Explicit
Public Sub Test_LoadAndCleanCsv_RemovesName()
    Dim tmpCsv As String
    tmpCsv = ThisWorkbook.Path & "\sample\participants.csv"
    Call LoadAndCleanCsv(tmpCsv)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data")
    Dim nameCol As Variant
    nameCol = Application.Match("Name", ws.Rows(1), 0)
    Assert.IsError nameCol, "Name 列が削除されていること"
End Sub
Rubberduck → Run All Tests で実行。全て Passed すれば OK。
EOF

---------- 2‑10. 6_risk-management ----------
mkdir -p 6_risk-management
cat > 6_risk-management/6.1_risk-register.md <<'EOF'

6.1 Risk Register

ID	リスク	発生確率	影響度	緩和策
R‑001	CSV の列名が「Name」以外（例: FirstName、姓）	中	高	正規表現で *name* 系列をすべて除去
R‑002	大容量 CSV（数千行）で処理が遅くなる	低	中	Application.ScreenUpdating = False、計算モード xlCalculationManual
R‑003	ユーザーがマクロ無効化している	中	高	起動時に Application.AutomationSecurity をチェックし警告
R‑004	PNG 出力先フォルダへの書き込み権限が無い	低	中	エラーハンドリングで LogError に出力し、メッセージ表示
EOF				
---------- 2‑11. sample ディレクトリ ----------
mkdir -p sample
cat > sample/participants.csv <<'EOF'
ParticipantID,Name,Score,Hours
P001,山田太郎,85,12
P002,鈴木花子,90,15
P003,佐藤次郎,78,9
EOF

---------- 2‑12. src/modules ----------
mkdir -p src/modules

ModMain.bas
cat > src/modules/ModMain.bas <<'EOF'
Option Explicit

Sub RunCsvGraphTool()
On Error GoTo ErrHandler

Dim csvPath As String
csvPath = Application.GetOpenFilename("CSV Files (*.csv),*.csv")
If csvPath = "False" Then Exit Sub   ' キャンセル
Call LoadAndCleanCsv(csvPath)      ' ModCSV
Call CreateChartAndExport()        ' ModChart
MsgBox "グラフ画像が作成され、シートに貼り付けました。", vbInformation
Exit Sub
ErrHandler:
Call LogError(Err.Number, Err.Description)
MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub
EOF

ModCSV.bas
cat > src/modules/ModCSV.bas <<'EOF'
Option Explicit

Public Sub LoadAndCleanCsv(ByVal fPath As String)
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Data") ' 予め「Data」シートを作成しておく

ws.Cells.Clear
With ws.QueryTables.Add(Connection:="TEXT;" & fPath, Destination:=ws.Range("A1"))
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    .TextFilePlatform = xlWindows
    .Refresh BackgroundQuery:=False
End With
' 「Name」または「氏名」列を削除
Dim colIdx As Long
colIdx = Application.Match("Name", ws.Rows(1), 0)
If IsError(colIdx) Then
    colIdx = Application.Match("氏名", ws.Rows(1), 0)
End If
If IsError(colIdx) Then
    Err.Raise vbObjectError + 1000, "ModCSV", "氏名列が見つかりません"
End If
ws.Columns(colIdx).Delete
End Sub
EOF

ModChart.bas
cat > src/modules/ModChart.bas <<'EOF'
Option Explicit

Public Sub CreateChartAndExport()
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data")
Dim chartObj As ChartObject
Dim rng As Range

Set rng = ws.UsedRange
Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=10, Width:=500, Height:=300)
With chartObj.Chart
    .SetSourceData Source:=rng
    .ChartType = xlColumnClustered      ' 棒グラフ例
    .HasTitle = True
    .ChartTitle.Text = "Data Overview"
    .Export Filename:=ThisWorkbook.Path & "\graph.png", FilterName:="PNG"
End With
' PNG を同シートに貼り付け
ws.Pictures.Insert ThisWorkbook.Path & "\graph.png"
End Sub
EOF

ModUtil.bas
cat > src/modules/ModUtil.bas <<'EOF'
Option Explicit

Public Sub LogError(ByVal errNo As Long, ByVal errMsg As String)
Dim logPath As String
logPath = ThisWorkbook.Path & "\error.log"
Dim f As Integer: f = FreeFile
Open logPath For Append As #f
Print #f, Now & " - Error " & errNo & ": " & errMsg
Close #f
End Sub
EOF

---------- 3️⃣ Git 初期化 ----------
git init
git add .
git commit -m "Initial commit: repository skeleton for CSV→Graph VBA tool"
git branch -M main

---------- 4️⃣ （任意）GitHub リモート作成 ----------
if [[ -n "REMOTEURL"]];thenecho"SettingremotetoREMOTE U​RL"]];thenecho"Settingremoteto{REMOTE_URL}"
git remote add origin "${REMOTE_URL}"
git push -u origin main
else
echo "リモート URL が未設定です。ローカルリポジトリだけが作成されました。"
fi

---------- 5️⃣ Node パッケージインストール（CI 用） ----------
markdownlint と mermaid-cli は CI 時にインストールされますが、ローカルでも試したい場合は以下を実行
npm install -g markdownlint-cli @mermaid-js/mermaid-cli
echo ""
echo "✅ 初期化が完了しました！"
echo " - ディレクトリ構成: $(pwd)"
echo " - 必要に応じて GitHub リモートを手動で設定してください（例: gh repo create …）"
echo " - VBA ファイルは src/modules/ 以下にあります。Excel の VBE へ Import してご利用ください。"
echo ""

---
## 3️⃣ 使い方
```bash
# 1️⃣ スクリプトに実行権限を付与
chmod +x init‑csv‑graph‑vba.sh
# 2️⃣ 実行（カレントディレクトリにリポジトリが作られます）
./init‑csv‑graph‑vba.sh	
