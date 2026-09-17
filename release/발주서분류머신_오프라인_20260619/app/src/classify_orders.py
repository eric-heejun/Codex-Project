from __future__ import annotations

import argparse
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = PROJECT_ROOT / "config" / "supplier_rules.json"
OUTPUT_DIR = PROJECT_ROOT / "data" / "output"


def normalize_key(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def compact_text(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_key(value)).lower()


def load_rules(path: Path = CONFIG_PATH) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"발주처 규칙 파일이 없습니다: {path}")

    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def latest_excel_file(input_dir: Path) -> Path:
    files = [
        path
        for path in input_dir.glob("*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"입력 폴더에 엑셀 파일이 없습니다: {input_dir}")

    return max(files, key=lambda path: path.stat().st_mtime)


def combine_text(row: pd.Series, columns: Iterable[str]) -> str:
    parts = []
    for column in columns:
        if column in row and pd.notna(row[column]):
            parts.append(str(row[column]))
    return " ".join(parts)


def find_keyword(text: str, keywords: Iterable[str]) -> str:
    compacted_text = compact_text(text)
    for keyword in keywords:
        compacted_keyword = compact_text(keyword)
        if compacted_keyword and compacted_keyword in compacted_text:
            return str(keyword)
    return ""


def classify_row(row: pd.Series, rules: dict[str, Any]) -> tuple[str, str, str, bool]:
    default_supplier = str(rules.get("default_supplier", "로지스존"))
    text = combine_text(row, rules.get("text_columns", ["상품명", "옵션", "배송메시지"]))

    for rule in rules.get("forced_keyword_rules", []):
        matched = find_keyword(text, rule.get("keywords", []))
        if matched:
            return str(rule["supplier"]), "상품명 예외", matched, False

    field7_column = str(rules.get("field7_column", "자사몰필드7"))
    field7_value = normalize_key(row[field7_column]) if field7_column in row.index else ""
    field7_prefix = field7_value[:1].upper()
    field7_rules = rules.get("field7_prefix_rules", {})
    if field7_prefix in field7_rules:
        return str(field7_rules[field7_prefix]), "자사몰필드7", f"{field7_column}={field7_value}", False

    for rule in rules.get("brand_keyword_rules", []):
        matched = find_keyword(text, rule.get("keywords", []))
        if matched:
            return str(rule["supplier"]), "브랜드명", matched, False

    return default_supplier, "기본값", "A/B 외 또는 빈 값", False


def classify_orders(input_path: Path) -> Path:
    rules = load_rules()
    target_column = str(rules.get("target_column", "발주처"))
    orders = pd.read_excel(input_path)

    if orders.empty:
        raise ValueError(f"입력 파일에 데이터가 없습니다: {input_path}")

    classified_rows = [classify_row(row, rules) for _, row in orders.iterrows()]
    predicted_suppliers = [item[0] for item in classified_rows]
    methods = [item[1] for item in classified_rows]
    reasons = [item[2] for item in classified_rows]
    review_flags = [item[3] for item in classified_rows]

    if target_column in orders.columns:
        output_target_column = f"예측{target_column}"
        orders[output_target_column] = predicted_suppliers
        orders["분류일치"] = [
            normalize_key(actual) == normalize_key(predicted)
            for actual, predicted in zip(orders[target_column], predicted_suppliers)
        ]
        review_flags = [
            needs_review or not is_match
            for needs_review, is_match in zip(review_flags, orders["분류일치"])
        ]
    else:
        output_target_column = target_column
        orders[output_target_column] = predicted_suppliers

    orders["분류방식"] = methods
    orders["분류근거"] = reasons
    orders["검수필요"] = ["예" if flag else "아니오" for flag in review_flags]

    quantity_column = "내품수량" if "내품수량" in orders.columns else None
    summary = (
        orders.groupby(output_target_column, dropna=False)
        .agg(건수=(output_target_column, "size"), 총수량=(quantity_column, "sum"))
        .reset_index()
        if quantity_column
        else orders.groupby(output_target_column, dropna=False).agg(건수=(output_target_column, "size")).reset_index()
    )

    method_summary = (
        orders.groupby(["분류방식", output_target_column], dropna=False)
        .agg(건수=(output_target_column, "size"))
        .reset_index()
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"발주처분류결과_{timestamp}.xlsx"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        orders.to_excel(writer, index=False, sheet_name="발주처분류")
        summary.to_excel(writer, index=False, sheet_name="요약")
        method_summary.to_excel(writer, index=False, sheet_name="분류방식요약")
        review_rows = orders[orders["검수필요"] == "예"]
        review_rows.to_excel(writer, index=False, sheet_name="검수필요")

        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 10), 40)

    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="사방넷 주문서 엑셀에 발주처를 자동 분류합니다.")
    parser.add_argument("--input", type=Path, help="분류할 엑셀 파일 경로")
    parser.add_argument("--input-dir", type=Path, default=PROJECT_ROOT / "data" / "input", help="최신 엑셀 파일을 찾을 입력 폴더")
    args = parser.parse_args()

    input_path = args.input if args.input else latest_excel_file(args.input_dir)
    output_path = classify_orders(input_path)
    print(f"분류 완료: {output_path}")


if __name__ == "__main__":
    main()
