"""
Genera un Excel de prueba y llama a POST /parse-excel (multipart/form-data).

Uso (con el servidor ya levantado):
  python scripts/smoke_test_local.py

Variables:
  BASE_URL: opcional (también en .env)
"""

from __future__ import annotations

import argparse
import io
import json
import os
import sys
import urllib.error
import urllib.request
import uuid
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


def build_sample_excel_bytes() -> bytes:
    df = pd.DataFrame(
        {
            "A": [1, 2, 3, 4],
            "B": ["x", "y", "z", "w"],
            "C": [" Ana ", "BRUNO", "carla", "diego"],
        }
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


def _multipart_body(
    fields: dict[str, str],
    file_field: str,
    filename: str,
    file_bytes: bytes,
    content_type: str,
) -> tuple[str, bytes]:
    boundary = uuid.uuid4().hex
    crlf = b"\r\n"
    parts: list[bytes] = []

    for name, value in fields.items():
        parts.append(f"--{boundary}".encode())
        parts.append(f'Content-Disposition: form-data; name="{name}"'.encode())
        parts.append(b"")
        parts.append(value.encode("utf-8"))

    parts.append(f"--{boundary}".encode())
    parts.append(
        f'Content-Disposition: form-data; name="{file_field}"; filename="{filename}"'.encode()
    )
    parts.append(f"Content-Type: {content_type}".encode())
    parts.append(b"")
    parts.append(file_bytes)

    parts.append(f"--{boundary}--".encode())
    parts.append(b"")
    body = crlf.join(parts)
    return boundary, body


def post_parse_multipart(base_url: str, excel_bytes: bytes) -> tuple[int, str]:
    url = base_url.rstrip("/") + "/parse-excel"
    boundary, body = _multipart_body(
        {"run_id": "smoke-local"},
        "file",
        "smoke_test.xlsx",
        excel_bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    req = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.getcode(), resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", errors="replace")


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    load_env_file(root / ".env")

    parser = argparse.ArgumentParser(description="Prueba local de /parse-excel (multipart)")
    parser.add_argument(
        "--base-url",
        default=os.environ.get("BASE_URL", "http://127.0.0.1:8000"),
        help="URL base del API (sin barra final)",
    )
    args = parser.parse_args()

    excel_bytes = build_sample_excel_bytes()

    code, text = post_parse_multipart(args.base_url, excel_bytes)
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
    if data.get("row_count") != 4:
        print("Esperaba row_count=4.", file=sys.stderr)
        return 1
    if len(data.get("rows", [])) != 4:
        print("Esperaba 4 filas en rows.", file=sys.stderr)
        return 1
    print("\nOK: smoke test pasó.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
