"""
中小企業財務分析・コンサルティングレポート自動生成スクリプト
中小企業診断士基準 × 大手コンサルティングフレームワーク

Usage:
    python3 generate_sme_report.py --data sample_input.json
    python3 generate_sme_report.py --data sample_input.json --out-dir ./output
"""

import json
import argparse
import sys
import os
from datetime import date
from pathlib import Path

# ─── ローカルモジュール ───────────────────────────────
try:
    from industry_benchmarks import get_benchmark, rag_signal, RAG_THRESHOLDS
    from word_template import generate_word
    from excel_model import generate_excel
except ImportError as e:
    print(f"[ERROR] モジュールが見つかりません: {e}")
    print("このスクリプトと同じディレクトリに以下のファイルが必要です:")
    print("  - industry_benchmarks.py")
    print("  - word_template.py")
    print("  - excel_model.py")
    sys.exit(1)


# ─────────────────────────────────────────────────────
# SMEFinancialAnalyzer — メインクラス
# ─────────────────────────────────────────────────────
class SMEFinancialAnalyzer:
    """
    中小企業財務分析エンジン

    Parameters
    ----------
    data : dict
        財務データ（sample_input.json 参照）
    """

    def __init__(self, data: dict):
        self.data     = data
        self.meta     = data.get("meta", {})
        self.pl       = data.get("pl", {})
        self.bs       = data.get("bs", {})
        self.cf       = data.get("cf", {})
        self.trend    = data.get("trend", [])

        self.industry = self.meta.get("industry", "")
        self.benchmark = get_benchmark(self.industry)

        self.ratios: dict   = {}
        self.issues: list   = []
        self.recommendations: list = []
        self.overall_rag: str = "AMBER"

    # ─────────────────────────────────────────────────
    # Step 2: 財務指標の計算
    # ─────────────────────────────────────────────────
    def calculate_ratios(self) -> dict:
        pl = self.pl
        bs = self.bs
        cf = self.cf

        rev  = pl.get("revenue") or 1
        cogs = pl.get("cogs", rev * 0.6)
        gp   = pl.get("gross_profit") or (rev - cogs)
        sga  = pl.get("sga", 0)
        op   = pl.get("operating_profit") or pl.get("operating_profit", gp - sga)
        oc   = pl.get("ordinary_profit", op)
        net  = pl.get("net_profit", oc * 0.65)
        int_exp = pl.get("interest_expense", 0)
        dep  = pl.get("depreciation", 0)

        ta   = bs.get("total_assets") or 1
        ca   = bs.get("current_assets", ta * 0.5)
        inv  = bs.get("inventory", 0)
        ar   = bs.get("accounts_receivable", 0)
        cash = bs.get("cash", 0)
        fa   = bs.get("fixed_assets", ta - ca)

        cl   = bs.get("current_liabilities", ta * 0.3)
        ap   = bs.get("accounts_payable", 0)
        st_debt = bs.get("short_term_debt", 0)
        lt_debt = bs.get("long_term_debt", 0)
        eq   = bs.get("equity") or (ta - cl - lt_debt) or 1

        ebitda = op + dep
        total_debt = st_debt + lt_debt
        net_debt   = total_debt - cash

        # ─── 収益性 ───
        r = {}
        r["gross_margin"]     = gp / rev
        r["operating_margin"] = op / rev
        r["ordinary_margin"]  = oc / rev
        r["roa"]              = oc / ta
        r["roe"]              = net / (eq or 1)
        r["ebitda_margin"]    = ebitda / rev

        # ─── 安全性 ───
        r["current_ratio"]    = ca / (cl or 1)
        r["quick_ratio"]      = (ca - inv) / (cl or 1)
        r["equity_ratio"]     = eq / ta
        r["debt_ratio"]       = total_debt / (eq or 1)
        r["interest_coverage"] = op / (int_exp or 0.001)
        r["net_debt_ebitda"]  = net_debt / (ebitda or 0.001)

        # ─── 効率性 ───
        r["asset_turnover"]   = rev / ta
        r["receivable_days"]  = (ar / (rev / 365)) if ar > 0 else 0
        r["inventory_days"]   = (inv / ((cogs or rev * 0.6) / 365)) if inv > 0 else 0
        r["payable_days"]     = (ap / ((cogs or rev * 0.6) / 365)) if ap > 0 else 0
        r["ccc"]              = r["receivable_days"] + r["inventory_days"] - r["payable_days"]

        # ─── 成長性 ───
        rev_prev = pl.get("revenue_prev")
        op_prev  = pl.get("operating_profit_prev")
        eq_prev  = bs.get("equity_prev")

        r["revenue_growth"]    = (rev - rev_prev) / rev_prev if rev_prev else None
        r["op_profit_growth"]  = (op - op_prev) / abs(op_prev) if op_prev else None
        r["equity_growth"]     = (eq - eq_prev) / abs(eq_prev) if eq_prev else None

        # ─── RAG評価 ───
        for key, thr in RAG_THRESHOLDS.items():
            val = r.get(key)
            if val is not None:
                r[f"_rag_{key}"] = rag_signal(
                    val, thr["low"], thr["high"], thr["inverse"])

        # ─── 総合RAG ───
        rag_vals = [v for k, v in r.items() if k.startswith("_rag_")]
        red_cnt  = rag_vals.count("RED")
        green_cnt = rag_vals.count("GREEN")
        if red_cnt >= 3:
            self.overall_rag = "RED"
        elif green_cnt >= 5 and red_cnt == 0:
            self.overall_rag = "GREEN"
        else:
            self.overall_rag = "AMBER"

        self.ratios = r
        return r

    # ─────────────────────────────────────────────────
    # Step 3: 課題の特定（MECE）
    # ─────────────────────────────────────────────────
    def identify_issues(self) -> list:
        r  = self.ratios
        bm = self.benchmark
        issues = []

        # ── ① 収益構造課題 ────────────────────────────────
        op_m = r.get("operating_margin", 0)
        bm_op = bm.get("operating_margin", 0.03)
        gm    = r.get("gross_margin", 0)
        bm_gm = bm.get("gross_margin", 0.25)

        if op_m < bm_op * 0.5:
            issues.append({
                "category": "収益構造",
                "title":    "営業利益率が業界平均の半分以下",
                "description": f"営業利益率 {op_m:.1%}（業界平均 {bm_op:.1%}）",
                "current_state": f"営業利益率は {op_m:.1%}と低水準で、業界平均（{bm_op:.1%}）を大きく下回っています。",
                "root_cause":   "価格競争による売価下落、または固定費・変動費の膨張が主因と考えられます。",
                "impact":       "持続的な収益力の低下、借入返済余力の縮小、投資原資の枯渇リスク。",
            })
        elif gm < bm_gm - 0.05:
            issues.append({
                "category": "収益構造",
                "title":    "売上総利益率が業界平均を下回る",
                "description": f"粗利率 {gm:.1%}（業界平均 {bm_gm:.1%}）",
                "current_state": f"売上総利益率は {gm:.1%}で業界平均（{bm_gm:.1%}）を下回っています。",
                "root_cause":   "原価高騰、廉価販売による価格競争、製品ミックス悪化が考えられます。",
                "impact":       "粗利率1%改善で年間約 {:.0f}千円の利益改善余地があります。".format(
                    (self.pl.get("revenue", 0)) * 0.01 / 1000),
            })

        # ── ② 財務健全性課題 ──────────────────────────────
        cr  = r.get("current_ratio", 1.5)
        er  = r.get("equity_ratio", 0.3)
        ic  = r.get("interest_coverage", 5)
        nde = r.get("net_debt_ebitda", 2)

        if cr < 1.0:
            issues.append({
                "category": "財務健全性",
                "title":    "流動比率1倍割れ — 短期流動性リスク",
                "description": f"流動比率 {cr:.2f}倍",
                "current_state": f"流動比率は {cr:.2f}倍で危険水準。短期債務返済能力に懸念があります。",
                "root_cause":   "短期借入依存度の高さ、または売掛金回収サイトの長期化が原因。",
                "impact":       "資金繰り悪化、取引先・金融機関の信用低下、最悪の場合支払不能リスク。",
            })
        elif er < 0.20:
            issues.append({
                "category": "財務健全性",
                "title":    "自己資本比率の低さ — レバレッジリスク",
                "description": f"自己資本比率 {er:.1%}",
                "current_state": f"自己資本比率は {er:.1%}で業界平均（{bm.get('equity_ratio', 0.3):.1%}）を下回っています。",
                "root_cause":   "過去の損失累積や積極的な借入による財務レバレッジの高止まり。",
                "impact":       "景気悪化・業績悪化時の財務耐久力が低く、借入コスト上昇リスク。",
            })

        if ic < 3 and ic > 0:
            issues.append({
                "category": "財務健全性",
                "title":    "インタレストカバレッジ比率の低さ",
                "description": f"ICR {ic:.1f}倍（目安: 3倍以上）",
                "current_state": f"インタレストカバレッジは {ic:.1f}倍と低水準。金利上昇への耐性が不足。",
                "root_cause":   "有利子負債の過大または営業利益の低水準。",
                "impact":       "金利上昇局面で利息負担が急増し、利益が圧迫されるリスク。",
            })

        # ── ③ 事業効率課題 ────────────────────────────────
        ccc = r.get("ccc", 30)
        bm_ccc = (bm.get("receivable_days", 45) +
                  bm.get("inventory_days", 30) -
                  bm.get("payable_days", 35))
        rec_d = r.get("receivable_days", 0)
        inv_d = r.get("inventory_days", 0)

        if ccc > bm_ccc + 15:
            issues.append({
                "category": "事業効率",
                "title":    "CCCの長期化 — 運転資金の固定化",
                "description": f"CCC {ccc:.0f}日（業界目安 {bm_ccc:.0f}日）",
                "current_state": f"現金転換サイクルは {ccc:.0f}日と業界平均比 {ccc-bm_ccc:.0f}日超過。",
                "root_cause":   (f"売掛金回収サイト {rec_d:.0f}日・"
                                 f"在庫回転日数 {inv_d:.0f}日の長期化が主因。"),
                "impact":       "追加運転資本の調達コスト増加、キャッシュフローの悪化、成長投資余力の喪失。",
            })

        at = r.get("asset_turnover", 1.0)
        bm_at = bm.get("asset_turnover", 1.0)
        if at < bm_at * 0.7:
            issues.append({
                "category": "事業効率",
                "title":    "資産効率の低下 — 総資産回転率",
                "description": f"総資産回転率 {at:.2f}回（業界平均 {bm_at:.2f}回）",
                "current_state": f"総資産回転率は {at:.2f}回で業界平均（{bm_at:.2f}回）を大幅に下回っています。",
                "root_cause":   "遊休資産の放置、売掛金の滞留、または過大な設備投資の可能性。",
                "impact":       "ROAの低下、資産生産性の悪化、資本コストを上回るリターン獲得が困難。",
            })

        # ── ④ 成長・競争力課題 ───────────────────────────
        rg = r.get("revenue_growth")
        bm_rg = bm.get("revenue_growth", 0.02)

        if rg is not None and rg < 0:
            issues.append({
                "category": "成長・競争力",
                "title":    "売上高の減少傾向",
                "description": f"売上高成長率 {rg:+.1%}",
                "current_state": f"売上高は前期比 {rg:.1%}と減少。業界成長率（{bm_rg:.1%}）と対照的。",
                "root_cause":   "市場縮小、競合台頭、主要顧客の離反、または製品陳腐化が考えられます。",
                "impact":       "固定費カバー余力の低下、従業員士気への悪影響、銀行格付けの低下リスク。",
            })
        elif rg is not None and rg < bm_rg - 0.02:
            issues.append({
                "category": "成長・競争力",
                "title":    "業界平均以下の成長率",
                "description": f"売上高成長率 {rg:+.1%}（業界平均 {bm_rg:+.1%}）",
                "current_state": f"売上高成長率は業界平均（{bm_rg:.1%}）を下回っており、相対的なシェア低下が懸念されます。",
                "root_cause":   "営業力不足、製品競争力の相対的低下、または新規顧客開拓の停滞。",
                "impact":       "中期的な規模の経済喪失、固定費負担増大リスク。",
            })

        # 課題が0件の場合は健全メッセージ
        if not issues:
            issues.append({
                "category": "総合",
                "title":    "総合的に健全な財務状態",
                "description": "主要KPIが業界平均水準を充足",
                "current_state": "収益性・安全性・効率性・成長性の全軸で業界平均水準以上を維持しています。",
                "root_cause":   "継続的なコスト管理と収益重視の経営が奏功しています。",
                "impact":       "更なる成長投資と付加価値向上に集中できる環境にあります。",
            })

        self.issues = issues
        return issues

    # ─────────────────────────────────────────────────
    # Step 4: 戦略提言の生成
    # ─────────────────────────────────────────────────
    def generate_recommendations(self) -> list:
        r    = self.ratios
        bm   = self.benchmark
        issues = self.issues
        recs = []

        # 課題カテゴリ別に施策を設定
        for issue in issues:
            cat = issue.get("category", "")

            # ── 収益構造 ──────────────────────────────────
            if "収益構造" in cat:
                op_m = r.get("operating_margin", 0)
                bm_op = bm.get("operating_margin", 0.03)

                if op_m < 0:
                    recs.append({
                        "timeline": "即時対応（0〜3ヶ月）",
                        "priority": 1,
                        "action":   "緊急コスト削減プログラムの立案・実行",
                        "detail":   "変動費の緊急圧縮（外注費・広告費・消耗品）と、"
                                    "固定費の聖域なき見直し（家賃交渉・人件費構造改革）",
                        "kpi":      "営業利益率 → 0%超（損益分岐点の突破）",
                        "expected_effect": "3ヶ月以内に月次単位での黒字化",
                        "owner":    "経営企画部門・財務部門",
                    })
                else:
                    recs.append({
                        "timeline": "短期施策（3〜12ヶ月）",
                        "priority": 2,
                        "action":   "粗利率改善プログラム（価格政策・原価管理）",
                        "detail":   "①製品・顧客別の利益貢献度分析による低採算案件の見直し "
                                    "②主要仕入先との価格再交渉 ③値上げ戦略の立案",
                        "kpi":      f"粗利率 → 業界平均 {bm.get('gross_margin', 0.25):.1%} 水準",
                        "expected_effect": "粗利率1%改善で売上高の1%相当の利益増",
                        "owner":    "営業部門・購買部門・経営企画部門",
                    })

            # ── 財務健全性 ────────────────────────────────
            if "財務健全性" in cat:
                cr = r.get("current_ratio", 1.5)
                if cr < 1.0:
                    recs.append({
                        "timeline": "即時対応（0〜3ヶ月）",
                        "priority": 1,
                        "action":   "緊急資金繰り対策 — 売掛金回収加速",
                        "detail":   "①滞留売掛金の一覧化と督促強化 "
                                    "②ファクタリング活用による早期現金化 "
                                    "③金融機関とのリスケジュール協議",
                        "kpi":      "流動比率 → 1.0倍超（短期安全水準）",
                        "expected_effect": "3ヶ月以内の流動比率改善と資金ショートリスク解消",
                        "owner":    "財務部門・経営トップ",
                    })
                recs.append({
                    "timeline": "中長期施策（1〜3年）",
                    "priority": 3,
                    "action":   "財務体質強化 — 自己資本比率の向上",
                    "detail":   "①内部留保の積み上げ（増益分の積極的留保方針） "
                                "②不要資産の売却による財務スリム化 "
                                "③DES（デット・エクイティ・スワップ）等の資本政策検討",
                    "kpi":      f"自己資本比率 → {bm.get('equity_ratio', 0.30):.0%} 以上",
                    "expected_effect": "銀行格付け改善、調達金利低下、財務耐久力向上",
                    "owner":    "財務部門・経営企画部門",
                })

            # ── 事業効率 ──────────────────────────────────
            if "事業効率" in cat:
                recs.append({
                    "timeline": "短期施策（3〜12ヶ月）",
                    "priority": 2,
                    "action":   "運転資本マネジメント（CCC短縮）",
                    "detail":   "①売掛金回収サイトの短縮（回収期間の契約見直し・督促プロセス強化） "
                                "②在庫最適化（SKU削減・発注ロット見直し・ABC分析） "
                                "③買掛金支払サイトの適正化（可能な範囲での支払延長交渉）",
                    "kpi":      f"CCC → 業界平均 {bm.get('receivable_days', 45)+bm.get('inventory_days', 30)-bm.get('payable_days', 35):.0f}日 以内",
                    "expected_effect": "CCC10日短縮で売上高×10/365分の運転資金解放",
                    "owner":    "営業部門・物流・購買部門",
                })

            # ── 成長・競争力 ──────────────────────────────
            if "成長" in cat or "競争" in cat:
                recs.append({
                    "timeline": "短期施策（3〜12ヶ月）",
                    "priority": 2,
                    "action":   "既存顧客深耕・新規顧客開拓の強化",
                    "detail":   "①上位20%顧客への重点営業（LTV向上） "
                                "②休眠顧客の再活性化キャンペーン "
                                "③デジタルマーケティング活用による新規リード創出",
                    "kpi":      "売上高成長率 → +5%以上（業界平均超え）",
                    "expected_effect": "12ヶ月以内に売上高5%増（固定費カバー余力の拡大）",
                    "owner":    "営業部門・マーケティング部門",
                })
                recs.append({
                    "timeline": "中長期施策（1〜3年）",
                    "priority": 3,
                    "action":   "新規事業・サービス拡充による収益源多様化",
                    "detail":   "①周辺市場への展開可能性調査 "
                                "②既存ケイパビリティを活用した付帯サービス提供 "
                                "③M&A・業務提携による市場拡大",
                    "kpi":      "新規事業売上比率 → 全体の20%以上",
                    "expected_effect": "3年後に売上高10〜20%増、収益安定化",
                    "owner":    "経営トップ・事業開発部門",
                })

        # 共通施策（常に追加）
        recs.append({
            "timeline": "中長期施策（1〜3年）",
            "priority": 3,
            "action":   "経営管理体制の高度化（KPI管理と月次PDCAサイクル）",
            "detail":   "①月次財務レポート（P/L・B/S・CF）の整備と経営会議への定期報告 "
                        "②KPIダッシュボードの導入 "
                        "③財務診断の四半期実施",
            "kpi":      "財務KPI達成率 → 設定KPIの80%以上",
            "expected_effect": "早期異常検知による損失最小化、銀行・投資家への信頼向上",
            "owner":    "経営企画部門・経営トップ",
        })

        # 重複除去・優先度順ソート
        seen = set()
        unique_recs = []
        for rec in recs:
            key = rec["action"]
            if key not in seen:
                seen.add(key)
                unique_recs.append(rec)

        unique_recs.sort(key=lambda x: x.get("priority", 9))
        self.recommendations = unique_recs
        return unique_recs

    # ─────────────────────────────────────────────────
    # 趨勢データの整形
    # ─────────────────────────────────────────────────
    def _build_trend_data(self) -> dict:
        """Excelトレンドシート用データを構築"""
        trend = self.trend or []
        current = {
            "year": self.meta.get("fiscal_year", "当期"),
            "revenue":          self.pl.get("revenue"),
            "gross_profit":     self.pl.get("gross_profit"),
            "operating_profit": self.pl.get("operating_profit"),
            "net_profit":       self.pl.get("net_profit"),
            "total_assets":     self.bs.get("total_assets"),
            "equity":           self.bs.get("equity"),
            "operating_margin": self.ratios.get("operating_margin"),
            "roa":              self.ratios.get("roa"),
            "equity_ratio":     self.ratios.get("equity_ratio"),
        }
        all_years = list(trend) + [current]
        return {"years": all_years}

    # ─────────────────────────────────────────────────
    # Step 5a: Wordレポート生成
    # ─────────────────────────────────────────────────
    def generate_word_report(self, output_path: str):
        meta = {
            **self.meta,
            "report_date": self.meta.get("report_date", str(date.today())),
        }
        generate_word(
            meta=meta,
            data=self.data,
            ratios=self.ratios,
            benchmark=self.benchmark,
            issues=self.issues,
            recommendations=self.recommendations,
            overall_rag=self.overall_rag,
            output_path=output_path,
        )

    # ─────────────────────────────────────────────────
    # Step 5b: Excelモデル生成
    # ─────────────────────────────────────────────────
    def generate_excel_model(self, output_path: str):
        trend_data_dict = self._build_trend_data()
        trend_list = trend_data_dict.get("years", [])
        generate_excel(
            data=self.data,
            ratios=self.ratios,
            benchmark=self.benchmark,
            trend_data=trend_list,
            output_path=output_path,
        )

    # ─────────────────────────────────────────────────
    # 全ステップを実行して両ファイルを生成
    # ─────────────────────────────────────────────────
    def run(self, out_dir: str = "."):
        company = self.meta.get("company_name", "会社").replace(" ", "_").replace("　", "_")
        fiscal  = self.meta.get("fiscal_year", "当期").replace(" ", "").replace("　", "")

        print("=" * 60)
        print(f"  中小企業財務分析スタート")
        print(f"  会社名: {self.meta.get('company_name', '─')}")
        print(f"  業種  : {self.industry}")
        print(f"  決算期: {fiscal}")
        print("=" * 60)

        print("\n[Step 2] 財務指標の計算中...")
        self.calculate_ratios()
        self._print_ratios_summary()

        print("\n[Step 3] 課題の特定中（MECEフレームワーク）...")
        self.identify_issues()
        for i, issue in enumerate(self.issues, 1):
            print(f"  課題{i}: {issue['title']}")

        print("\n[Step 4] 戦略提言の生成中...")
        self.generate_recommendations()
        for i, rec in enumerate(self.recommendations, 1):
            print(f"  提言{i}: [{rec['timeline']}] {rec['action']}")

        print("\n[Step 5] 成果物の生成中...")
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        word_path  = out_dir / f"{company}_{fiscal}_財務診断レポート.docx"
        excel_path = out_dir / f"{company}_{fiscal}_財務モデル.xlsx"

        print(f"  → Wordレポート: {word_path}")
        self.generate_word_report(str(word_path))

        print(f"  → Excelモデル: {excel_path}")
        self.generate_excel_model(str(excel_path))

        print("\n" + "=" * 60)
        print(f"✅ 生成完了！")
        print(f"   📄 Word: {word_path}")
        print(f"   📊 Excel: {excel_path}")
        print(f"   総合評価: {self.overall_rag} "
              f"({'◎ 良好' if self.overall_rag=='GREEN' else '△ 要注意' if self.overall_rag=='AMBER' else '✕ 要改善'})")
        print("=" * 60)

        return str(word_path), str(excel_path)

    def _print_ratios_summary(self):
        print("\n  ── 主要KPI ──")
        kpis = [
            ("営業利益率",  "operating_margin",  "{:.1%}"),
            ("ROA",         "roa",               "{:.1%}"),
            ("自己資本比率","equity_ratio",       "{:.1%}"),
            ("流動比率",    "current_ratio",     "{:.2f}倍"),
            ("CCC",         "ccc",               "{:.0f}日"),
        ]
        for label, key, fmt in kpis:
            val = self.ratios.get(key)
            rag = self.ratios.get(f"_rag_{key}", "─")
            rag_icon = {"GREEN": "◎", "AMBER": "△", "RED": "✕"}.get(rag, "─")
            val_str = fmt.format(val) if val is not None else "─"
            print(f"  {label:<18} {val_str:>10}  {rag_icon}")


# ─────────────────────────────────────────────────────
# CLI エントリポイント
# ─────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="中小企業財務分析 — Word & Excel 自動生成",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
例:
  python3 generate_sme_report.py --data sample_input.json
  python3 generate_sme_report.py --data sample_input.json --out-dir ./reports
        """
    )
    parser.add_argument("--data",    required=True, help="財務データJSON ファイルパス")
    parser.add_argument("--out-dir", default=".",   help="出力先ディレクトリ（デフォルト: カレント）")
    args = parser.parse_args()

    # データ読み込み
    data_path = Path(args.data)
    if not data_path.exists():
        print(f"[ERROR] データファイルが見つかりません: {data_path}")
        sys.exit(1)

    with open(data_path, encoding="utf-8") as f:
        data = json.load(f)

    # 分析・生成
    analyzer = SMEFinancialAnalyzer(data)
    analyzer.run(out_dir=args.out_dir)


if __name__ == "__main__":
    main()
