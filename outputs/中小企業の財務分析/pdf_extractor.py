#!/usr/bin/env python3
"""
pdf_extractor.py — PDFから財務データを自動抽出してJSONに変換

使い方:
  python pdf_extractor.py 決算書.pdf
  python pdf_extractor.py 決算書.pdf --out 出力.json
  python pdf_extractor.py 決算書.pdf --vision   # Claude Vision API使用（画像PDF向け）

対応フォーマット:
  - テキスト型PDF（e.g. 会計ソフト出力）: pdfplumber で直接抽出
  - 画像型PDF（スキャン・写真）: pdftoppm → PNG → Claude Vision API

出力形式: generate_sme_report.py の sample_input.json と同じ構造
"""

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path


# =============================================================================
# 財務用語マッピング（日本語 → JSONキー）
# =============================================================================

# P/L（損益計算書）マッピング
PL_MAPPING = {
    # 売上高
    "売上高": "revenue",
    "売上収益": "revenue",
    "純売上高": "revenue",
    "営業収益": "revenue",
    # 売上原価
    "売上原価": "cost_of_sales",
    "商品売上原価": "cost_of_sales",
    "製品売上原価": "cost_of_sales",
    # 売上総利益
    "売上総利益": "gross_profit",
    "粗利益": "gross_profit",
    # 販管費
    "販売費及び一般管理費": "sg_and_a",
    "販管費": "sg_and_a",
    "販売費・一般管理費": "sg_and_a",
    # 営業利益
    "営業利益": "operating_profit",
    "営業損失": "operating_profit",  # マイナス処理は後で
    # 営業外収益
    "営業外収益": "non_operating_income",
    "受取利息": "interest_income",
    "受取配当金": "dividend_income",
    # 営業外費用
    "営業外費用": "non_operating_expenses",
    "支払利息": "interest_expense",
    "支払利息・割引料": "interest_expense",
    # 経常利益
    "経常利益": "ordinary_profit",
    "経常損失": "ordinary_profit",
    # 特別損益
    "特別利益": "extraordinary_income",
    "特別損失": "extraordinary_loss",
    # 税引前利益
    "税引前当期純利益": "profit_before_tax",
    "税金等調整前当期純利益": "profit_before_tax",
    # 法人税等
    "法人税等": "income_taxes",
    "法人税、住民税及び事業税": "income_taxes",
    # 当期純利益
    "当期純利益": "net_profit",
    "当期純損失": "net_profit",
    "当期利益": "net_profit",
    # 減価償却費
    "減価償却費": "depreciation",
    "のれん償却額": "amortization",
    # 人件費
    "人件費": "personnel_expenses",
    "給料及び手当": "personnel_expenses",
    # 地代家賃
    "地代家賃": "rent_expenses",
    "賃借料": "rent_expenses",
}

# B/S（貸借対照表）マッピング
BS_MAPPING = {
    # 流動資産
    "流動資産": "current_assets",
    "現金及び預金": "cash_and_deposits",
    "現金預金": "cash_and_deposits",
    "受取手形": "notes_receivable",
    "売掛金": "accounts_receivable",
    "受取手形及び売掛金": "accounts_receivable",
    "棚卸資産": "inventories",
    "商品": "inventories",
    "製品": "inventories",
    "前払費用": "prepaid_expenses",
    "その他流動資産": "other_current_assets",
    # 固定資産
    "固定資産": "non_current_assets",
    "有形固定資産": "tangible_assets",
    "建物及び構築物": "buildings",
    "機械装置": "machinery",
    "土地": "land",
    "無形固定資産": "intangible_assets",
    "のれん": "goodwill",
    "投資その他の資産": "investments_and_others",
    "投資有価証券": "investment_securities",
    # 資産合計
    "資産の部合計": "total_assets",
    "資産合計": "total_assets",
    "総資産": "total_assets",
    # 流動負債
    "流動負債": "current_liabilities",
    "支払手形": "notes_payable",
    "買掛金": "accounts_payable",
    "支払手形及び買掛金": "accounts_payable",
    "短期借入金": "short_term_borrowings",
    "1年以内返済長期借入金": "current_portion_of_long_term_debt",
    "未払費用": "accrued_expenses",
    "前受金": "advances_received",
    # 固定負債
    "固定負債": "non_current_liabilities",
    "長期借入金": "long_term_borrowings",
    "社債": "bonds_payable",
    "退職給付引当金": "retirement_benefit_liability",
    # 負債合計
    "負債の部合計": "total_liabilities",
    "負債合計": "total_liabilities",
    # 純資産
    "純資産の部合計": "total_equity",
    "純資産合計": "total_equity",
    "自己資本": "total_equity",
    "資本金": "capital_stock",
    "利益剰余金": "retained_earnings",
    # 負債・純資産合計
    "負債及び純資産の部合計": "total_liabilities_and_equity",
    "負債純資産合計": "total_liabilities_and_equity",
}

# CF（キャッシュフロー計算書）マッピング
CF_MAPPING = {
    "営業活動によるキャッシュ・フロー": "operating_cf",
    "営業活動によるキャッシュフロー": "operating_cf",
    "投資活動によるキャッシュ・フロー": "investing_cf",
    "投資活動によるキャッシュフロー": "investing_cf",
    "財務活動によるキャッシュ・フロー": "financing_cf",
    "財務活動によるキャッシュフロー": "financing_cf",
    "現金及び現金同等物の増減額": "net_change_in_cash",
    "現金及び現金同等物の期末残高": "cash_end_of_period",
}


# =============================================================================
# テキスト抽出（pdfplumber使用）
# =============================================================================

def extract_text_from_pdf(pdf_path: str) -> list[str]:
    """pdfplumber でPDFからテキストを抽出して行リストを返す"""
    try:
        import pdfplumber
    except ImportError:
        print("pdfplumber が未インストールです。インストール中...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pdfplumber", "-q"], check=True)
        import pdfplumber

    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split("\n"))
    return lines


# =============================================================================
# テキスト行から財務数値をパース
# =============================================================================

NUM_PATTERN = re.compile(
    r"[△▲\-\(]?\s*[\d,，０-９]+(?:\.\d+)?\s*[）\)]?"
)

def parse_number(raw: str) -> float | None:
    """日本語財務数値文字列を float に変換（△/▲/括弧はマイナス）"""
    raw = raw.strip()
    negative = False
    if raw.startswith(("△", "▲", "-", "▽")):
        negative = True
        raw = raw[1:].strip()
    if raw.startswith("(") or raw.startswith("（"):
        negative = True
        raw = raw.lstrip("(（").rstrip(")）").strip()

    # 全角数字 → 半角
    raw = raw.translate(str.maketrans("０１２３４５６７８９，．", "0123456789,."))
    raw = raw.replace(",", "").replace("，", "").strip()

    try:
        val = float(raw)
        return -val if negative else val
    except ValueError:
        return None


def extract_values_from_lines(lines: list[str], mapping: dict) -> dict:
    """テキスト行リストから指定マッピングの数値を抽出"""
    result = {}

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # マッピングの各キーを確認
        for jp_name, json_key in mapping.items():
            if jp_name in line:
                # 行内の数値を検索
                nums = NUM_PATTERN.findall(line)
                if nums:
                    # 最後の数値を採用（小計・合計行の場合、最後が合計値）
                    val = parse_number(nums[-1])
                    if val is not None and json_key not in result:
                        result[json_key] = val

    return result


# =============================================================================
# Claude Vision API 使用（画像PDF向け）
# =============================================================================

def extract_via_claude_vision(pdf_path: str) -> dict:
    """
    画像型PDFの場合: pdftoppm で PNG に変換 → Claude API で財務数値を抽出
    ANTHROPIC_API_KEY 環境変数が必要
    """
    import anthropic
    import base64

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError(
            "ANTHROPIC_API_KEY 環境変数が設定されていません。\n"
            "export ANTHROPIC_API_KEY='sk-ant-...' を実行してください。"
        )

    print("画像型PDFを処理中: pdftoppm でPNGに変換...")
    with tempfile.TemporaryDirectory() as tmpdir:
        # PDF → PNG
        out_prefix = os.path.join(tmpdir, "page")
        result = subprocess.run(
            ["pdftoppm", "-r", "150", "-png", pdf_path, out_prefix],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise RuntimeError(f"pdftoppm 失敗: {result.stderr}")

        png_files = sorted(Path(tmpdir).glob("*.png"))
        if not png_files:
            raise RuntimeError("PDFからPNGへの変換に失敗しました。")

        print(f"  {len(png_files)} ページをClaude APIで分析中...")

        # 全ページのbase64エンコード
        images_content = []
        for png_path in png_files[:20]:  # 最大20ページ
            with open(png_path, "rb") as f:
                img_data = base64.standard_b64encode(f.read()).decode("utf-8")
            images_content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/png", "data": img_data}
            })
            images_content.append({
                "type": "text",
                "text": f"（上記: {png_path.name}）"
            })

        prompt = """以下のPDF（決算書・財務諸表）の画像から財務データを抽出してください。

【抽出対象（損益計算書・貸借対照表・CF計算書から）】
単位は万円または千円として抽出し、以下のJSON形式で出力してください。
単位が百万円の場合は×100、千円の場合は÷10して万円に統一してください。

出力形式（コードブロックなしのJSONのみ）:
{
  "pl": {
    "revenue": <売上高>,
    "cost_of_sales": <売上原価>,
    "gross_profit": <売上総利益>,
    "sg_and_a": <販売費及び一般管理費>,
    "operating_profit": <営業利益>,
    "interest_expense": <支払利息>,
    "ordinary_profit": <経常利益>,
    "net_profit": <当期純利益>,
    "depreciation": <減価償却費>,
    "personnel_expenses": <人件費>,
    "rent_expenses": <地代家賃>
  },
  "bs": {
    "cash_and_deposits": <現金及び預金>,
    "accounts_receivable": <売掛金>,
    "inventories": <棚卸資産>,
    "current_assets": <流動資産合計>,
    "tangible_assets": <有形固定資産>,
    "total_assets": <資産合計>,
    "accounts_payable": <買掛金>,
    "short_term_borrowings": <短期借入金>,
    "current_liabilities": <流動負債合計>,
    "long_term_borrowings": <長期借入金>,
    "total_liabilities": <負債合計>,
    "total_equity": <純資産合計>
  },
  "cf": {
    "operating_cf": <営業CF>,
    "investing_cf": <投資CF>,
    "financing_cf": <財務CF>
  },
  "unit": "万円",
  "periods": <決算期数（例: 3）>,
  "fiscal_year_end": <最新決算年月 例: "2024-03">
}

数値が見つからない項目はnullにしてください。複数期ある場合は最新期の数値を使用してください。"""

        images_content.append({"type": "text", "text": prompt})

        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-opus-4-8",
            max_tokens=2000,
            messages=[{"role": "user", "content": images_content}]
        )

        response_text = response.content[0].text.strip()

        # JSONを抽出
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        else:
            raise ValueError(f"Claude APIからのレスポンスにJSONが含まれていません:\n{response_text}")


# =============================================================================
# テキスト抽出結果を sample_input.json 形式に変換
# =============================================================================

def build_output_json(pl: dict, bs: dict, cf: dict, meta: dict) -> dict:
    """抽出した財務データをgenerate_sme_report.py用JSONに整形"""

    def v(d, key, default=0):
        val = d.get(key)
        return val if val is not None else default

    revenue = v(pl, "revenue")
    operating_profit = v(pl, "operating_profit")
    ordinary_profit = v(pl, "ordinary_profit")
    net_profit = v(pl, "net_profit")
    gross_profit = v(pl, "gross_profit")
    depreciation = v(pl, "depreciation")
    interest_expense = v(pl, "interest_expense")
    personnel_expenses = v(pl, "personnel_expenses")
    rent_expenses = v(pl, "rent_expenses")

    total_assets = v(bs, "total_assets")
    current_assets = v(bs, "current_assets")
    current_liabilities = v(bs, "current_liabilities")
    total_equity = v(bs, "total_equity")
    accounts_receivable = v(bs, "accounts_receivable")
    inventories = v(bs, "inventories")
    accounts_payable = v(bs, "accounts_payable")
    short_term_borrowings = v(bs, "short_term_borrowings")
    long_term_borrowings = v(bs, "long_term_borrowings")
    cash = v(bs, "cash_and_deposits")

    operating_cf = v(cf, "operating_cf")
    investing_cf = v(cf, "investing_cf")
    financing_cf = v(cf, "financing_cf")

    # 付加価値額（加算法）
    value_added = ordinary_profit + personnel_expenses + depreciation + interest_expense + rent_expenses

    # 有利子負債
    interest_bearing_debt = short_term_borrowings + long_term_borrowings

    output = {
        "company": {
            "name": meta.get("company_name", "（会社名未設定）"),
            "industry": meta.get("industry", "製造業"),
            "fiscal_year": meta.get("fiscal_year_end", "2024-03"),
            "employees": meta.get("employees", 0),
            "founded_year": meta.get("founded_year", 2000),
            "capital": meta.get("capital", 0)
        },
        "financials": {
            "pl": {
                "revenue": revenue,
                "cost_of_sales": v(pl, "cost_of_sales"),
                "gross_profit": gross_profit if gross_profit else revenue - v(pl, "cost_of_sales"),
                "sg_and_a": v(pl, "sg_and_a"),
                "operating_profit": operating_profit,
                "non_operating_income": v(pl, "non_operating_income"),
                "interest_income": v(pl, "interest_income"),
                "non_operating_expenses": v(pl, "non_operating_expenses"),
                "interest_expense": interest_expense,
                "ordinary_profit": ordinary_profit,
                "extraordinary_income": v(pl, "extraordinary_income"),
                "extraordinary_loss": v(pl, "extraordinary_loss"),
                "profit_before_tax": v(pl, "profit_before_tax"),
                "income_taxes": v(pl, "income_taxes"),
                "net_profit": net_profit,
                "depreciation": depreciation,
                "personnel_expenses": personnel_expenses,
                "rent_expenses": rent_expenses
            },
            "bs": {
                "cash_and_deposits": cash,
                "accounts_receivable": accounts_receivable,
                "inventories": inventories,
                "current_assets": current_assets,
                "tangible_assets": v(bs, "tangible_assets"),
                "total_assets": total_assets,
                "accounts_payable": accounts_payable,
                "short_term_borrowings": short_term_borrowings,
                "current_liabilities": current_liabilities,
                "long_term_borrowings": long_term_borrowings,
                "total_liabilities": v(bs, "total_liabilities"),
                "capital_stock": v(bs, "capital_stock"),
                "retained_earnings": v(bs, "retained_earnings"),
                "total_equity": total_equity
            },
            "cf": {
                "operating_cf": operating_cf,
                "investing_cf": investing_cf,
                "financing_cf": financing_cf,
                "free_cf": operating_cf + investing_cf,
                "net_change_in_cash": operating_cf + investing_cf + financing_cf
            },
            "derived": {
                "ebitda": ordinary_profit + interest_expense + depreciation,
                "value_added": value_added,
                "interest_bearing_debt": interest_bearing_debt,
                "net_debt": interest_bearing_debt - cash
            }
        }
    }
    return output


# =============================================================================
# メイン処理
# =============================================================================

def detect_pdf_type(pdf_path: str) -> str:
    """PDFがテキスト型か画像型かを判定"""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            total_chars = 0
            for page in pdf.pages[:3]:  # 最初の3ページで判定
                text = page.extract_text() or ""
                total_chars += len(text.strip())
        return "text" if total_chars > 100 else "image"
    except Exception:
        return "image"


def interactive_meta() -> dict:
    """会社基本情報を対話的に入力"""
    print("\n--- 会社基本情報の入力 ---")
    print("（Enterでスキップ可）")
    meta = {}
    meta["company_name"] = input("会社名: ").strip() or "株式会社サンプル"
    meta["industry"] = input("業種（製造業/卸売業/小売業/建設業/サービス業/IT・情報通信業/飲食業/不動産業）: ").strip() or "製造業"
    meta["fiscal_year_end"] = input("決算年月（例: 2024-03）: ").strip() or "2024-03"
    emp = input("従業員数（人）: ").strip()
    meta["employees"] = int(emp) if emp.isdigit() else 0
    return meta


def main():
    parser = argparse.ArgumentParser(
        description="PDFから財務データを抽出してJSONに変換"
    )
    parser.add_argument("pdf", help="入力PDFファイルパス")
    parser.add_argument("--out", "-o", help="出力JSONファイルパス（省略時: <pdf名>.json）")
    parser.add_argument("--vision", action="store_true", help="Claude Vision APIを強制使用（画像PDF向け）")
    parser.add_argument("--no-meta", action="store_true", help="会社情報の入力をスキップ")
    parser.add_argument("--unit", choices=["万円", "千円", "百万円"], default="万円", help="財務数値の単位")
    args = parser.parse_args()

    pdf_path = args.pdf
    if not os.path.exists(pdf_path):
        print(f"エラー: ファイルが見つかりません: {pdf_path}")
        sys.exit(1)

    out_path = args.out or Path(pdf_path).with_suffix(".json")

    print(f"PDFを処理中: {pdf_path}")

    # 会社情報の入力
    meta = {} if args.no_meta else interactive_meta()

    # PDF種別の判定
    if args.vision:
        pdf_type = "image"
    else:
        pdf_type = detect_pdf_type(pdf_path)
        print(f"PDF種別: {'テキスト型' if pdf_type == 'text' else '画像型'}")

    pl, bs, cf = {}, {}, {}

    if pdf_type == "text":
        # テキスト抽出
        print("テキスト抽出中...")
        lines = extract_text_from_pdf(pdf_path)
        print(f"  {len(lines)} 行を抽出しました")

        pl = extract_values_from_lines(lines, PL_MAPPING)
        bs = extract_values_from_lines(lines, BS_MAPPING)
        cf = extract_values_from_lines(lines, CF_MAPPING)

        # 単位変換
        if args.unit == "千円":
            factor = 0.1  # 千円 → 万円
            pl = {k: v * factor for k, v in pl.items()}
            bs = {k: v * factor for k, v in bs.items()}
            cf = {k: v * factor for k, v in cf.items()}
        elif args.unit == "百万円":
            factor = 100  # 百万円 → 万円
            pl = {k: v * factor for k, v in pl.items()}
            bs = {k: v * factor for k, v in bs.items()}
            cf = {k: v * factor for k, v in cf.items()}

        extracted_count = len(pl) + len(bs) + len(cf)
        print(f"  抽出成功: {extracted_count} 項目 (PL:{len(pl)}, BS:{len(bs)}, CF:{len(cf)})")

        if extracted_count < 5:
            print("\n  ⚠ テキスト抽出の項目数が少ないです。画像型PDFの可能性があります。")
            print("  --vision オプションを付けて再実行することをお勧めします:")
            print(f"  python pdf_extractor.py {pdf_path} --vision")

    else:
        # Vision API使用
        vision_result = extract_via_claude_vision(pdf_path)
        pl = vision_result.get("pl", {})
        bs = vision_result.get("bs", {})
        cf = vision_result.get("cf", {})
        if "fiscal_year_end" in vision_result and not meta.get("fiscal_year_end"):
            meta["fiscal_year_end"] = vision_result["fiscal_year_end"]
        print(f"  Vision API 抽出成功: PL:{len(pl)}, BS:{len(bs)}, CF:{len(cf)}")

    # JSON生成
    output = build_output_json(pl, bs, cf, meta)

    # 保存
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✅ JSON出力完了: {out_path}")
    print("\n次のステップ:")
    print(f"  python generate_sme_report.py {out_path}")
    print("")
    print("主要抽出値（万円）:")
    pl_data = output["financials"]["pl"]
    bs_data = output["financials"]["bs"]
    if pl_data.get("revenue"):
        print(f"  売上高        : {pl_data['revenue']:>12,.0f}")
    if pl_data.get("ordinary_profit"):
        print(f"  経常利益      : {pl_data['ordinary_profit']:>12,.0f}")
    if bs_data.get("total_assets"):
        print(f"  総資産        : {bs_data['total_assets']:>12,.0f}")
    if bs_data.get("total_equity"):
        print(f"  純資産        : {bs_data['total_equity']:>12,.0f}")
    cf_data = output["financials"]["cf"]
    if cf_data.get("operating_cf"):
        print(f"  営業CF        : {cf_data['operating_cf']:>12,.0f}")


if __name__ == "__main__":
    main()
