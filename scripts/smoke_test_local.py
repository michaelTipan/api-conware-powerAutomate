"""
Genera un Excel de prueba y llama a POST /parse-excel.

Uso (con el servidor ya levantado):
  python scripts/smoke_test_local.py

Variables:
  BASE_URL: opcional (también en .env)
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

import pandas as pd


def load_env_file(path: Path) -> None:
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def build_sample_excel_b64() -> str:
    df = pd.DataFrame(
        {
            "A": [1, 2, 3, 4],
            "B": ["x", "y", "z", "w"],
            "C": [" Ana ", "BRUNO", "carla", "diego"],
        }
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return base64.b64encode(buf.getvalue()).decode("ascii")


def post_parse(base_url: str, excel_b64: str) -> tuple[int, str]:
    url = base_url.rstrip("/") + "/parse-excel"
    body = {
        "file_name": "smoke_test.xlsx",
        "excel_base64": excel_b64,
        "run_id": "smoke-local",
        "column_letter": "C",
    }
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        method="POST",
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.getcode(), resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", errors="replace")


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    load_env_file(root / ".env")

    parser = argparse.ArgumentParser(description="Prueba local de /parse-excel")
    parser.add_argument(
        "--base-url",
        default=os.environ.get("BASE_URL", "http://127.0.0.1:8000"),
        help="URL base del API (sin barra final)",
    )
    parser.add_argument(
        "--print-b64-only",
        action="store_true",
        help="Solo imprime excel_base64 y sale (no llama al servidor)",
    )
    args = parser.parse_args()

    excel_b64 = build_sample_excel_b64()

    if args.print_b64_only:
        print(excel_b64)
        return 0

    code, text = post_parse(args.base_url, excel_b64)
    print(f"HTTP {code}")
    try:
        parsed = json.loads(text)
        print(json.dumps(parsed, indent=2, ensure_ascii=False))
    except json.JSONDecodeError:
        print(text)

    if code != 200:
        return 1
    data = json.loads(text)
    if data.get("status") != "ok":
        return 1
    if data.get("unique_count") != 4:
        print("Esperaba unique_count=4.", file=sys.stderr)
        return 1
    if len(data.get("values_in_row_order", [])) != 4:
        print("Esperaba 4 valores en values_in_row_order.", file=sys.stderr)
        return 1
    print("\nOK: smoke test pasó.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
