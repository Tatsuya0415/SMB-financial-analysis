"""
業種別ベンチマークデータ
出典: 中小企業庁「中小企業実態基本調査」/ TKC経営指標(BAST) / 中小企業診断士試験参考値

■ 指標体系（PDFテキスト「中小企業の財務分析」準拠）

【収益性】
  1.売上高総利益率  2.売上高営業利益率  3.売上高経常利益率
  4.売上高純利益率  5.ROA  6.ROE  7.EBITDA率

【安全性】
  8.流動比率  9.当座比率  10.自己資本比率  11.負債比率
  12.固定比率  13.固定長期適合率  14.ICR  15.債務償還年数

【効率性（CCC）】
  16.総資産回転率  17.売上債権回転日数  18.棚卸資産回転日数
  19.買入債務回転日数  20.CCC  21.固定資産回転率

【生産性】
  22.1人当たり売上高  23.1人当たり付加価値額  24.労働生産性
  25.労働分配率  26.売上高人件費率  27.付加価値率

【成長性】
  28.売上高増加率  29.営業利益増加率  30.総資産増加率
"""

INDUSTRY_BENCHMARKS = {
    "製造業": {
        "label": "製造業",
        # 収益性
        "gross_margin": 0.25,          # 売上総利益率（優良30%↑ 平均20-30% 要注意20%↓）
        "operating_margin": 0.04,      # 営業利益率（優良5%↑ 平均2-5% 要注意2%↓）
        "ordinary_margin": 0.04,       # 経常利益率
        "roa": 0.03,                   # ROA
        "roe": 0.06,                   # ROE
        "ebitda_margin": 0.07,         # EBITDAマージン
        # 安全性
        "current_ratio": 1.50,         # 流動比率（優良150%↑）
        "quick_ratio": 1.10,           # 当座比率
        "equity_ratio": 0.38,          # 自己資本比率（優良40%↑ 要注意20%↓）
        # 効率性
        "asset_turnover": 0.90,        # 総資産回転率（優良1.5回↑）
        "receivable_days": 55,         # 売掛金回転日数（優良30日↓）
        "inventory_days": 45,          # 棚卸資産回転日数（優良30日↓ 要注意60日↑）
        "payable_days": 40,            # 買掛金回転日数
        # 生産性
        "revenue_per_employee": 1200,  # 1人当たり売上高（万円）
        "value_added_per_employee": 500, # 1人当たり付加価値額（万円）
        "labor_distribution_ratio": 0.55, # 労働分配率
        # 成長性
        "revenue_growth": 0.02,        # 売上高成長率（優良5%↑）
    },
    "卸売業": {
        "label": "卸売業",
        # 収益性
        "gross_margin": 0.14,
        "operating_margin": 0.02,
        "ordinary_margin": 0.02,
        "roa": 0.02,
        "roe": 0.05,
        "ebitda_margin": 0.03,
        # 安全性
        "current_ratio": 1.30,
        "quick_ratio": 0.90,
        "equity_ratio": 0.25,
        # 効率性
        "asset_turnover": 1.60,
        "receivable_days": 50,
        "inventory_days": 25,
        "payable_days": 45,
        # 生産性
        "revenue_per_employee": 3000,  # 1人当たり売上高（万円）卸売は売上規模大
        "value_added_per_employee": 400,
        "labor_distribution_ratio": 0.65,
        # 成長性
        "revenue_growth": 0.01,
    },
    "小売業": {
        "label": "小売業",
        # 収益性
        "gross_margin": 0.28,
        "operating_margin": 0.02,
        "ordinary_margin": 0.02,
        "roa": 0.02,
        "roe": 0.05,
        "ebitda_margin": 0.04,
        # 安全性
        "current_ratio": 1.10,
        "quick_ratio": 0.70,
        "equity_ratio": 0.22,
        # 効率性
        "asset_turnover": 1.40,
        "receivable_days": 10,
        "inventory_days": 40,
        "payable_days": 35,
        # 生産性
        "revenue_per_employee": 2000,
        "value_added_per_employee": 380,
        "labor_distribution_ratio": 0.70,
        # 成長性
        "revenue_growth": 0.01,
    },
    "建設業": {
        "label": "建設業",
        # 収益性
        "gross_margin": 0.20,
        "operating_margin": 0.03,
        "ordinary_margin": 0.03,
        "roa": 0.03,
        "roe": 0.07,
        "ebitda_margin": 0.05,
        # 安全性
        "current_ratio": 1.40,
        "quick_ratio": 1.10,
        "equity_ratio": 0.30,
        # 効率性
        "asset_turnover": 1.10,
        "receivable_days": 60,
        "inventory_days": 30,
        "payable_days": 50,
        # 生産性
        "revenue_per_employee": 1500,
        "value_added_per_employee": 550,
        "labor_distribution_ratio": 0.60,
        # 成長性
        "revenue_growth": 0.03,
    },
    "サービス業": {
        "label": "サービス業",
        # 収益性
        "gross_margin": 0.55,
        "operating_margin": 0.05,
        "ordinary_margin": 0.05,
        "roa": 0.04,
        "roe": 0.08,
        "ebitda_margin": 0.10,
        # 安全性
        "current_ratio": 1.20,
        "quick_ratio": 1.10,
        "equity_ratio": 0.35,
        # 効率性
        "asset_turnover": 0.80,
        "receivable_days": 35,
        "inventory_days": 5,
        "payable_days": 25,
        # 生産性
        "revenue_per_employee": 800,
        "value_added_per_employee": 450,
        "labor_distribution_ratio": 0.62,
        # 成長性
        "revenue_growth": 0.03,
    },
    "IT・情報通信業": {
        "label": "IT・情報通信業",
        # 収益性
        "gross_margin": 0.50,
        "operating_margin": 0.07,
        "ordinary_margin": 0.07,
        "roa": 0.05,
        "roe": 0.10,
        "ebitda_margin": 0.12,
        # 安全性
        "current_ratio": 1.50,
        "quick_ratio": 1.40,
        "equity_ratio": 0.45,
        # 効率性
        "asset_turnover": 0.70,
        "receivable_days": 40,
        "inventory_days": 5,
        "payable_days": 30,
        # 生産性
        "revenue_per_employee": 1000,
        "value_added_per_employee": 700,
        "labor_distribution_ratio": 0.55,
        # 成長性
        "revenue_growth": 0.05,
    },
    "飲食業": {
        "label": "飲食業",
        # 収益性
        "gross_margin": 0.65,
        "operating_margin": 0.03,
        "ordinary_margin": 0.03,
        "roa": 0.03,
        "roe": 0.06,
        "ebitda_margin": 0.08,
        # 安全性
        "current_ratio": 0.90,
        "quick_ratio": 0.70,
        "equity_ratio": 0.20,
        # 効率性
        "asset_turnover": 1.20,
        "receivable_days": 5,
        "inventory_days": 10,
        "payable_days": 20,
        # 生産性
        "revenue_per_employee": 500,
        "value_added_per_employee": 300,
        "labor_distribution_ratio": 0.75,
        # 成長性
        "revenue_growth": 0.02,
    },
    "不動産業": {
        "label": "不動産業",
        # 収益性
        "gross_margin": 0.30,
        "operating_margin": 0.10,
        "ordinary_margin": 0.10,
        "roa": 0.03,
        "roe": 0.05,
        "ebitda_margin": 0.15,
        # 安全性
        "current_ratio": 1.20,
        "quick_ratio": 0.90,
        "equity_ratio": 0.25,
        # 効率性
        "asset_turnover": 0.30,
        "receivable_days": 20,
        "inventory_days": 200,
        "payable_days": 30,
        # 生産性
        "revenue_per_employee": 2500,
        "value_added_per_employee": 800,
        "labor_distribution_ratio": 0.45,
        # 成長性
        "revenue_growth": 0.02,
    },
}

# デフォルト（業種不明時）
DEFAULT_BENCHMARK = {
    "label": "全産業平均",
    "gross_margin": 0.28,
    "operating_margin": 0.03,
    "ordinary_margin": 0.03,
    "roa": 0.03,
    "roe": 0.06,
    "ebitda_margin": 0.07,
    "current_ratio": 1.30,
    "quick_ratio": 0.95,
    "equity_ratio": 0.30,
    "asset_turnover": 1.00,
    "receivable_days": 45,
    "inventory_days": 35,
    "payable_days": 35,
    "revenue_per_employee": 1000,
    "value_added_per_employee": 450,
    "labor_distribution_ratio": 0.60,
    "revenue_growth": 0.02,
}

def get_benchmark(industry: str) -> dict:
    """業種名から最も近いベンチマークを返す"""
    for key in INDUSTRY_BENCHMARKS:
        if key in (industry or ""):
            return INDUSTRY_BENCHMARKS[key]
    # 部分一致検索
    industry_lower = (industry or "").lower()
    for key, val in INDUSTRY_BENCHMARKS.items():
        if any(k in industry_lower for k in [key.lower(), key[:2].lower()]):
            return val
    return {**DEFAULT_BENCHMARK, "label": f"{industry or '全産業'}（参考値）"}


def rag_signal(value: float, low: float, high: float, inverse: bool = False) -> str:
    """
    RAG（Red/Amber/Green）信号を返す
    inverse=True: 値が低いほど良い指標（負債比率等）
    """
    if inverse:
        if value <= low:
            return "GREEN"
        elif value <= high:
            return "AMBER"
        else:
            return "RED"
    else:
        if value >= high:
            return "GREEN"
        elif value >= low:
            return "AMBER"
        else:
            return "RED"


# 指標別RAG閾値定義
RAG_THRESHOLDS = {
    "gross_margin":       {"low": 0.10, "high": 0.25, "inverse": False, "label": "売上総利益率"},
    "operating_margin":   {"low": 0.02, "high": 0.05, "inverse": False, "label": "営業利益率"},
    "ordinary_margin":    {"low": 0.02, "high": 0.05, "inverse": False, "label": "経常利益率"},
    "roa":                {"low": 0.02, "high": 0.05, "inverse": False, "label": "ROA"},
    "roe":                {"low": 0.05, "high": 0.10, "inverse": False, "label": "ROE"},
    "current_ratio":      {"low": 1.00, "high": 1.30, "inverse": False, "label": "流動比率"},
    "quick_ratio":        {"low": 0.80, "high": 1.00, "inverse": False, "label": "当座比率"},
    "equity_ratio":       {"low": 0.20, "high": 0.40, "inverse": False, "label": "自己資本比率"},
    "debt_ratio":         {"low": 1.00, "high": 2.00, "inverse": True,  "label": "負債比率"},
    "interest_coverage":  {"low": 3.00, "high": 8.00, "inverse": False, "label": "インタレストカバレッジ"},
    "net_debt_ebitda":    {"low": 2.00, "high": 4.00, "inverse": True,  "label": "純有利子負債/EBITDA"},
    "asset_turnover":     {"low": 0.60, "high": 1.20, "inverse": False, "label": "総資産回転率"},
    "receivable_days":    {"low": 30,   "high": 60,   "inverse": True,  "label": "売掛金回転日数"},
    "inventory_days":     {"low": 30,   "high": 60,   "inverse": True,  "label": "棚卸資産回転日数"},
    "ccc":                {"low": 20,   "high": 60,   "inverse": True,  "label": "CCC"},
    "revenue_growth":     {"low": 0.00, "high": 0.03, "inverse": False, "label": "売上成長率"},
    "ebitda_margin":      {"low": 0.03, "high": 0.08, "inverse": False, "label": "EBITDAマージン"},
}
