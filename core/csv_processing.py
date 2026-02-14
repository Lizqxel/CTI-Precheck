import csv
import re
from pathlib import Path
from typing import Dict, List, Tuple

from utils.address_utils import normalize_address

ZIP_PATTERN = re.compile(r"^\d{3}-?\d{4}$")


def decode_csv_bytes(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("unknown", b"", 0, 1, "サポートされていないエンコーディングです")


def normalize_zipcode(value: str) -> str:
    cleaned = (value or "").strip().replace("－", "-").replace("ー", "-")
    digits_only = re.sub(r"\D", "", cleaned)
    if len(digits_only) == 7:
        return f"{digits_only[:3]}-{digits_only[3:]}"
    return cleaned


def read_csv(file_path: Path) -> List[List[str]]:
    file_bytes = file_path.read_bytes()
    text = decode_csv_bytes(file_bytes)
    reader = csv.reader(text.splitlines())
    return [row for row in reader]


def validate_rows(rows: List[List[str]]) -> Tuple[List[Dict[str, str]], List[int]]:
    parsed: List[Dict[str, str]] = []
    invalid_line_numbers: List[int] = []

    for index, row in enumerate(rows, start=1):
        zipcode = row[0].strip() if len(row) >= 1 and row[0] is not None else ""
        address = row[1].strip() if len(row) >= 2 and row[1] is not None else ""

        normalized_zipcode = normalize_zipcode(zipcode)
        normalized_address = normalize_address(address) if address else ""

        status = "OK"
        if not zipcode and not address:
            status = "空行"
        elif not zipcode or not address:
            status = "入力不足"
            invalid_line_numbers.append(index)
        elif not ZIP_PATTERN.match(normalized_zipcode):
            status = "郵便番号形式エラー"
            invalid_line_numbers.append(index)

        parsed.append(
            {
                "行": str(index),
                "郵便番号": normalized_zipcode,
                "住所": normalized_address,
                "状態": status,
                "判定結果": "未実行",
                "備考": "",
            }
        )

    return parsed, invalid_line_numbers
