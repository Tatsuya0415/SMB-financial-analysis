"""
業種別ベンチマークデータ
出典: 中小企業庁「中小企業実態基本調査」/ TKC経営指標(BAST) / 中小企業診断士試験参考値
"""

INDUSTRY_BENCHMARKS = {
    "製造業": {
        "label": "製造業",
        "gross_margin": 0.25,          # 売上総利益率
        "operating_margin": 0.04,      # 営業利益率
        "ordinary_margin": 0.04,       # 経常利益率
        "roa": 0.03,                   # 総資産利益率
        "roe": 0.06,                   # 自己資本利益率
        "current_ratio": 1.50,         # 流動比率
        "quick_ratio": 1.10,           # 当座比率
        "equity_ratio": 0.38,          # 自己資本比率
        "asset_turnover": 0.90,        # 総資産回転率
        "receivable_days": 55,         # 売掛金回転日数
        "inventory_days": 45,          # 棚卸資産回転日数
        "payable_days": 40,            # 買掛金回転日数
        "revenue_growth": 0.02,        # 売上高成長率
        "ebitda_margin": 0.07,
    },
    "卸売業": {
        "label": "卸売業",
        "gross_margin": 0.14,
        "operating_margin": 0.02,
        "ordinary_margin": 0.02,
        "roa": 0.02,
        "roe": 0.05,
        "current_ratio": 1.30,
        "quick_ratio": 0.90,
        "equity_ratio": 0.25,
        "asset_turnover": 1.60,
        "receivable_days": 50,
        "inventory_days": 25,
        "payable_days": 45,
        "revenue_growth": 0.01,
        "ebitda_margin": 0.03,
    },
    "小売業": {
        "label": "小売業",
        "gross_margin": 0.28,
        "operating_margin": 0.02,
        "ordinary_margin": 0.02,
        "roa": 0.02,
        "roe": 0.05,
        "current_ratio": 1.10,
        "quick_ratio": 0.70,
        "equity_ratio": 0.22,
        "asset_turnover": 1.40,
        "receivable_days": 10,
        "inventory_days": 40,
        "payable_days": 35,
        "revenue_growth": 0.01,
        "ebitda_margin": 0.04,
    },
    "建設業": {
        "label": "建設業",
        "gross_margin": 0.20,
        "operating_margin": 0.03,
        "ordinary_margin": 0.03,
        "roa": 0.03,
        "roe": 0.07,
        "current_ratio": 1.40,
        "quick_ratio": 1.10,
        "equity_ratio": 0.30,
        "asset_turnover": 1.10,
        "receivable_days": 60,
        "inventory_days": 30,
        "payable_days": 50,
        "revenue_growth": 0.03,
        "ebitda_margin": 0.05,
    },
    "サービス業": {
        "label": "サービス業",
        "gross_margin": 0.55,
        "operating_margin": 0.05,
        "ordinary_margin": 0.05,
        "roa": 0.04,
        "roe": 0.08,
        "current_ratio": 1.20,
        "quick_ratio": 1.10,
        "equity_ratio": 0.35,
        "asset_turnover": 0.80,
        "receivable_days": 35,
        "inventory_days": 5,
        "payable_days": 25,
        "revenue_growth": 0.03,
        "ebitda_margin": 0.10,
    },
    "IT・情報通信業": {
        "label": "IT・情報通信業",
        "gross_margin": 0.50,
        "operating_margin": 0.07,
        "ordinary_margin": 0.07,
        "roa": 0.05,
        "roe": 0.10,
        "current_ratio": 1.50,
        "quick_ratio": 1.40,
        "equity_ratio": 0.45,
        "asset_turnover": 0.70,
        "receivable_days": 40,
        "inventory_days": 5,
        "payable_days": 30,
        "revenue_growth": 0.05,
        "ebitda_margin": 0.12,
    },
    "飲食業": {
        "label": "飲食業",
        "gross_margin": 0.65,
        "operating_margin": 0.03,
        "ordinary_margin": 0.03,
        "roa": 0.03,
        "roe": 0.06,
        "current_ratio": 0.90,
        "quick_ratio": 0.70,
        "equity_ratio": 0.20,
        "asset_turnover": 1.20,
        "receivable_days": 5,
        "inventory_days": 10,
        "payable_days": 20,
        "revenue_growth": 0.02,
        "ebitda_margin": 0.08,
    },
    "不動産業": {
        "label": "不動産業",
        "gross_margin": 0.30,
        "operating_margin": 0.10,
        "ordinary_margin": 0.10,
        "roa": 0.03,
        "roe": 0.05,
        "current_ratio": 1.20,
        "quick_ratio": 0.90,
        "equity_ratio": 0.25,
        "asset_turnover": 0.30,
        "receivable_days": 20,
        "inventory_days": 200,
        "payable_days": 30,
        "revenue_growth": 0.02,
        "ebitda_margin": 0.15,
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
    "current_ratio": 1.30,
    "quick_ratio": 0.95,
    "equity_ratio": 0.30,
    "asset_turnover": 1.00,
    "receivable_days": 45,
    "inventory_days": 35,
    "payable_days": 35,
    "revenue_growth": 0.02,
    "ebitda_margin": 0.07,
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
