from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from classify_orders import classify_row, load_rules, normalize_key


def validate_rules(labeled_file: Path, target_column: str) -> None:
    rules = load_rules()
    df = pd.read_excel(labeled_file)

    if target_column not in df.columns:
        raise ValueError(f"검증 파일에 '{target_column}' 열이 없습니다: {labeled_file}")

    predicted = [classify_row(row, rules)[0] for _, row in df.iterrows()]
    actual = [normalize_key(value) for value in df[target_column]]
    matches = [normalize_key(left) == normalize_key(right) for left, right in zip(predicted, actual)]

    print(f"검증 파일: {labeled_file}")
    print(f"전체 행: {len(df)}")
    print(f"일치: {sum(matches)}")
    print(f"불일치: {len(matches) - sum(matches)}")

    if not all(matches):
        mismatch = df.loc[[not item for item in matches]].copy()
        mismatch["예측발주처"] = [value for value, ok in zip(predicted, matches) if not ok]
        columns = [column for column in ["자사몰필드7", "상품명", "옵션", target_column, "예측발주처"] if column in mismatch.columns]
        print(mismatch[columns].head(20).to_string(index=False))


def main() -> None:
    parser = argparse.ArgumentParser(description="현재 발주처 규칙을 발주처가 채워진 엑셀 파일로 검증합니다.")
    parser.add_argument("--labeled-file", type=Path, required=True, help="발주처 열이 채워진 엑셀 파일")
    parser.add_argument("--output", type=Path, help="이전 버전 호환용 인자입니다. 현재는 사용하지 않습니다.")
    parser.add_argument("--target-column", default="발주처", help="정답 발주처 열 이름")
    args = parser.parse_args()

    validate_rules(args.labeled_file, args.target_column)


if __name__ == "__main__":
    main()
