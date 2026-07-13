# 中小企業 財務分析・コンサルティングレポート自動生成スキル

中小企業の財務データを入力するだけで、**中小企業診断士レベル × 大手コンサルティングファーム流**の
財務診断レポート（Word 10ページ）と財務モデル（Excel 6シート）を自動生成します。

---

## 特徴

| 項目 | 内容 |
|------|------|
| 分析フレームワーク | 中小企業診断士4軸（収益性・安全性・効率性・成長性） |
| 業界比較 | 8業種 × 17 KPI のベンチマーク比較（中小企業庁・TKC BAST 参考値） |
| RAGシグナル | 各KPIを Green/Amber/Red の信号機で評価 |
| 課題特定 | MECEフレームワークで4カテゴリに分類 |
| 戦略提言 | 即時対応 / 短期 / 中長期の3フェーズでロードマップ化 |
| Word出力 | A4 10ページ、游ゴシック、ネイビー×RAGカラースキーム |
| Excel出力 | 6シート（財務3表・KPI・趨勢・業界比較・シナリオ・資金繰り） |

---

## ファイル構成

```
中小企業の財務分析/
├── SKILL.md                    # スキル定義（Claudeへの指示書）
├── generate_sme_report.py      # メイン生成スクリプト ← ここから実行
├── industry_benchmarks.py      # 業種別ベンチマーク・RAG評価ロジック
├── word_template.py            # Word（.docx）レポート生成モジュール
├── excel_model.py              # Excel（.xlsx）財務モデル生成モジュール
├── sample_input.json           # サンプル入力データ（製造業・株式会社サンプル製造）
├── requirements.txt            # 依存パッケージ
└── README.md                   # このファイル
```

---

## セットアップ

### 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

> Python 3.9 以上を推奨

### 2. サンプルデータで動作確認

```bash
python3 generate_sme_report.py --data sample_input.json
```

成功すると以下のファイルが生成されます：

- `株式会社サンプル製造_2024年3月期_財務診断レポート.docx`
- `株式会社サンプル製造_2024年3月期_財務モデル.xlsx`

---

## 使い方

### コマンドライン

```bash
# カレントディレクトリに出力
python3 generate_sme_report.py --data 自社データ.json

# 出力先を指定
python3 generate_sme_report.py --data 自社データ.json --out-dir ./reports
```

### Python からの呼び出し

```python
import json
from generate_sme_report import SMEFinancialAnalyzer

with open("自社データ.json", encoding="utf-8") as f:
    data = json.load(f)

analyzer = SMEFinancialAnalyzer(data)
word_path, excel_path = analyzer.run(out_dir="./reports")
```

---

## 入力データ形式（JSON）

```json
{
  "meta": {
    "company_name": "株式会社○○",
    "fiscal_year": "2024年3月期",
    "industry": "製造業",
    "employees": 45,
    "report_date": "2024-06-30"
  },
  "pl": {
    "_unit": "千円",
    "revenue": 500000,
    "revenue_prev": 470000,
    "cogs": 360000,
    "gross_profit": 140000,
    "sga": 105000,
    "operating_profit": 35000,
    "interest_expense": 3500,
    "ordinary_profit": 32000,
    "net_profit": 21000,
    "depreciation": 12000
  },
  "bs": {
    "_unit": "千円",
    "cash": 65000,
    "accounts_receivable": 95000,
    "inventory": 55000,
    "current_assets": 230000,
    "fixed_assets": 120000,
    "total_assets": 350000,
    "accounts_payable": 60000,
    "short_term_debt": 50000,
    "current_liabilities": 140000,
    "long_term_debt": 70000,
    "equity": 135000
  },
  "cf": {
    "_unit": "千円",
    "operating_cf": 48000,
    "investing_cf": -18000,
    "financing_cf": -22000
  },
  "trend": []
}
```

### 対応業種（`industry` フィールド）

| 設定値 | 業種 |
|--------|------|
| `製造業` | 製造業全般 |
| `卸売業` | 卸売業 |
| `小売業` | 小売業 |
| `建設業` | 建設業・土木 |
| `サービス業` | サービス業全般 |
| `IT・情報通信業` | IT・SaaS・情報通信 |
| `飲食業` | 飲食・フード |
| `不動産業` | 不動産・管理 |

業種が一致しない場合は「全産業平均」のベンチマークを使用します。

---

## 生成成果物

### Word レポート（10ページ）

| ページ | 内容 |
|--------|------|
| P1 | 表紙（会社名・診断日・診断機関） |
| P2 | エグゼクティブサマリー（KPIシグナル・Top3課題・Top3提言） |
| P3-4 | 財務ハイライト（PL/BS サマリー・KPI 業界比較テーブル） |
| P5 | 収益性分析（売上構造・マージン分析） |
| P6 | 安全性・効率性分析（流動性・CCCほか） |
| P7 | 成長性・キャッシュフロー分析 |
| P8 | 課題の特定（MECEフレームワーク・根本原因） |
| P9-10 | 戦略提言ロードマップ（フェーズ別アクションプラン） |

### Excel モデル（6シート）

| シート | 内容 |
|--------|------|
| ① 財務3表 | PL・BS・CF の一覧 |
| ② 経営指標 | 17+ KPI・RAGシグナル・業界比較 |
| ③ 趨勢分析 | 複数期の売上・利益トレンド（折れ線グラフ付） |
| ④ 業界比較 | 主要KPIの業界平均対比（棒グラフ） |
| ⑤ シナリオ | 楽観・中立・悲観の3シナリオP/L |
| ⑥ 資金繰り | 12ヶ月キャッシュフロー予測 |

---

## 分析フレームワーク

```
収益性 ─── 営業利益率 / ROA / ROE / EBITDAマージン
安全性 ─── 流動比率 / 当座比率 / 自己資本比率 / ICR / 純有利子負債/EBITDA
効率性 ─── CCC / 売掛金回転日数 / 在庫回転日数 / 総資産回転率
成長性 ─── 売上成長率 / 営業利益成長率 / 自己資本成長率
```

各KPIは業種別ベンチマーク（中小企業庁「中小企業実態基本調査」・TKC経営指標BAST参考）との
比較で自動評価されます。

---

## 注意事項

- 本ツールの出力は **経営判断の参考情報** であり、公認会計士・税理士による正式な財務監査の代替ではありません
- 業界ベンチマーク値は中小企業庁・TKC の公開統計を参考に設定しており、個社の状況と異なる場合があります
- 入力データの正確性が分析精度に直結します。入力値は必ずご確認ください

---

## ライセンス

MIT License — 商用・非商用問わず自由に使用・改変可能です。

---

*Powered by Claude Agent SDK × 中小企業診断士フレームワーク*
