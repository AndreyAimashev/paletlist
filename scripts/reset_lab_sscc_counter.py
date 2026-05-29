#!/usr/bin/env python3
"""Сброс счётчика SSCC паллет ЛАБ (по умолчанию следующий паллет = 12)."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from api_server import init_db, reset_lab_sscc_pallet_counter  # noqa: E402


def main() -> None:
    p = argparse.ArgumentParser(description="Сброс lab_sscc_last_shipped")
    p.add_argument(
        "--next-pallet",
        type=int,
        default=12,
        help="Номер следующего паллета в SSCC (по умолчанию 12)",
    )
    p.add_argument(
        "--last-shipped",
        type=int,
        default=None,
        help="Последний отгруженный номер (если задан, --next-pallet игнорируется)",
    )
    args = p.parse_args()
    init_db()
    out = reset_lab_sscc_pallet_counter(
        last_shipped=args.last_shipped,
        next_pallet=None if args.last_shipped is not None else args.next_pallet,
    )
    print(out)


if __name__ == "__main__":
    main()
