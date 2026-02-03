from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from .logic import build_cash_recon
from .models import Period
from .parse import load_hotcake_bills_xlsx, load_hotcake_orders_xlsx, load_pos_history_orders_xlsx
from .report import save_cash_recon_report


def _parse_dt(s: str) -> datetime:
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError("日期時間格式請用 YYYY-MM-DD HH:MM[:SS]")


def main() -> int:
    p = argparse.ArgumentParser(description="Generate cash reconciliation report.")
    p.add_argument("--store", required=True, help="分店名稱（需與報表一致）")
    p.add_argument("--start", required=True, help="區間開始 YYYY-MM-DD HH:MM[:SS]")
    p.add_argument("--end", required=True, help="區間結束 YYYY-MM-DD HH:MM[:SS]")
    p.add_argument("--hotcake-bills", required=True, type=Path)
    p.add_argument("--hotcake-orders", required=True, type=Path)
    p.add_argument("--pos-orders", type=Path, default=None)
    p.add_argument("--out", type=Path, default=Path("output/spreadsheet/cash_recon.xlsx"))
    args = p.parse_args()

    period = Period(start=_parse_dt(args.start), end=_parse_dt(args.end))
    orders = load_hotcake_orders_xlsx(args.hotcake_orders)
    bills = load_hotcake_bills_xlsx(args.hotcake_bills)
    pos_orders = load_pos_history_orders_xlsx(args.pos_orders) if args.pos_orders else None

    result = build_cash_recon(period=period, store=args.store, orders=orders, bills=bills, pos_orders=pos_orders)
    save_cash_recon_report(result, args.out)
    print(args.out)
    if result.missing_bills:
        print(f"WARNING: MissingBills={len(result.missing_bills)} (現金對帳不保證正確)")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
