from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import re

from .models import (
    CashTotals,
    HotcakeTimeMismatchRow,
    MissingBillRow,
    Period,
    PosTimeMismatchRow,
    ServiceBillRow,
    TopUpBillRow,
)
from .parse import HotcakeBills, OrdersRow, PosRow


def _in_period(dt: datetime, period: Period) -> bool:
    return period.start <= dt <= period.end


def _normalize_store(value: str) -> str:
    return (value or "").replace(" ", "").replace("\u3000", "").lower()


def _normalize_name(value: str) -> str:
    return (value or "").replace(" ", "").replace("\u3000", "").lower()


def _extract_minutes(text: str) -> Optional[int]:
    if not text:
        return None
    m = re.search(r"(\\d+)\\s*分鐘", text)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return None


def _parse_pos_designer(product_name: str) -> str:
    if not product_name:
        return ""
    part = product_name.split(",")[0]
    return part.strip()


@dataclass(frozen=True)
class CashReconResult:
    period: Period
    store: str
    missing_bills: list[MissingBillRow]
    service_bill_rows: list[ServiceBillRow]
    topup_bill_rows: list[TopUpBillRow]
    hotcake_time_mismatches: list[HotcakeTimeMismatchRow]
    pos_time_mismatches: list[PosTimeMismatchRow]
    totals: CashTotals


def build_cash_recon(
    *,
    period: Period,
    store: str,
    orders: list[OrdersRow],
    bills: HotcakeBills,
    pos_orders: Optional[list[PosRow]] = None,
    topup_mode: str = "settlement_time",
    time_tolerance_minutes: int = 30,
) -> CashReconResult:
    store_norm = _normalize_store(store)
    scoped_orders = [
        o for o in orders if _normalize_store(o.store) == store_norm and _in_period(o.service_start, period)
    ]

    missing_bills: list[MissingBillRow] = []
    bill_ids_in_scope: set[str] = set()
    for o in scoped_orders:
        if o.order_status == "已報到" and (not o.bill_id) and (o.bill_amount == 0.0):
            missing_bills.append(
                MissingBillRow(
                    store=o.store,
                    order_id=o.order_id,
                    order_code=o.order_code,
                    service_start=o.service_start,
                    designer=o.designer,
                    service=o.service,
                    order_status=o.order_status,
                    checkin_time=o.checkin_time,
                    member_name=o.member_name,
                    phone=o.phone,
                )
            )
        if o.bill_id:
            bill_ids_in_scope.add(o.bill_id)

    service_bill_rows: list[ServiceBillRow] = []
    if bill_ids_in_scope:
        for r in bills.service:
            if r.bill_id not in bill_ids_in_scope:
                continue
            service_bill_rows.append(
                ServiceBillRow(
                    store=r.store,
                    bill_id=r.bill_id,
                    settlement_time=r.settlement_time,
                    attributed_date=r.attributed_date,
                    designer=r.designer,
                    item=r.item,
                    cash=r.cash,
                    bill_amount=r.bill_amount,
                )
            )

    # Top-up cash is not tied to service start in orders, so filter by settlement time in the same period.
    topup_bill_rows: list[TopUpBillRow] = []
    if topup_mode != "exclude":
        for r in bills.topup:
            if _normalize_store(r.store) != store_norm:
                continue
            if not _in_period(r.settlement_time, period):
                continue
            topup_bill_rows.append(
                TopUpBillRow(
                    store=r.store,
                    bill_id=r.bill_id,
                    settlement_time=r.settlement_time,
                    attributed_date=r.attributed_date,
                    designer=r.designer,
                    item=r.item,
                    cash=r.cash,
                    bill_amount=r.bill_amount,
                )
            )

    hotcake_service_cash = sum(r.cash for r in service_bill_rows)
    hotcake_topup_cash = sum(r.cash for r in topup_bill_rows)
    hotcake_cash_total = hotcake_service_cash + hotcake_topup_cash

    # Build bill cash map for mismatch reporting
    bill_cash_map: dict[str, float] = {}
    for r in service_bill_rows:
        bill_cash_map[r.bill_id] = bill_cash_map.get(r.bill_id, 0.0) + r.cash

    pos_cash_total: Optional[float] = None
    pos_cash_diff: Optional[float] = None
    if pos_orders is not None:
        scoped_pos = [p for p in pos_orders if _in_period(p.created_time, period)]
        pos_cash_total = sum(p.cash_paid for p in scoped_pos)
        pos_cash_diff = pos_cash_total - hotcake_cash_total
    else:
        scoped_pos = []

    # Time mismatch detection (Hotcake vs POS)
    hotcake_time_mismatches: list[HotcakeTimeMismatchRow] = []
    pos_time_mismatches: list[PosTimeMismatchRow] = []

    if scoped_pos:
        pos_candidates = []
        for p in scoped_pos:
            designer = _parse_pos_designer(p.product_name)
            pos_candidates.append(
                {
                    "row": p,
                    "designer_norm": _normalize_name(designer),
                    "minutes": _extract_minutes(p.product_name),
                }
            )
        used_pos = set()

        for o in scoped_orders:
            designer_norm = _normalize_name(o.designer)
            minutes = _extract_minutes(o.service)
            if not designer_norm:
                hotcake_time_mismatches.append(
                    HotcakeTimeMismatchRow(
                        store=o.store,
                        service_start=o.service_start,
                        designer=o.designer,
                        service=o.service,
                        minutes=minutes,
                        bill_id=o.bill_id,
                        bill_amount=o.bill_amount,
                        cash=bill_cash_map.get(o.bill_id, 0.0),
                        nearest_pos_time=None,
                        nearest_diff_minutes=None,
                    )
                )
                continue
            # candidate POS rows: same designer, minutes match if both present
            best_idx = None
            best_diff = None
            nearest_diff = None
            nearest_time = None
            for idx, p in enumerate(pos_candidates):
                if idx in used_pos:
                    continue
                if designer_norm and designer_norm != p["designer_norm"]:
                    continue
                if minutes is not None and p["minutes"] is not None and minutes != p["minutes"]:
                    continue
                diff = abs((p["row"].created_time - o.service_start).total_seconds() / 60)
                if nearest_diff is None or diff < nearest_diff:
                    nearest_diff = diff
                    nearest_time = p["row"].created_time
                if best_diff is None or diff < best_diff:
                    best_diff = diff
                    best_idx = idx

            if best_idx is not None and best_diff is not None and best_diff <= time_tolerance_minutes:
                used_pos.add(best_idx)
                continue

            hotcake_time_mismatches.append(
                HotcakeTimeMismatchRow(
                    store=o.store,
                    service_start=o.service_start,
                    designer=o.designer,
                    service=o.service,
                    minutes=minutes,
                    bill_id=o.bill_id,
                    bill_amount=o.bill_amount,
                    cash=bill_cash_map.get(o.bill_id, 0.0),
                    nearest_pos_time=nearest_time,
                    nearest_diff_minutes=int(nearest_diff) if nearest_diff is not None else None,
                )
            )

        for idx, p in enumerate(pos_candidates):
            if idx in used_pos:
                continue
            row = p["row"]
            # find nearest hotcake order for context
            nearest_diff = None
            nearest_time = None
            for o in scoped_orders:
                if _normalize_name(o.designer) != p["designer_norm"]:
                    continue
                diff = abs((row.created_time - o.service_start).total_seconds() / 60)
                if nearest_diff is None or diff < nearest_diff:
                    nearest_diff = diff
                    nearest_time = o.service_start
            pos_time_mismatches.append(
                PosTimeMismatchRow(
                    terminal_name=row.terminal_name,
                    created_time=row.created_time,
                    designer=_parse_pos_designer(row.product_name),
                    product_name=row.product_name,
                    minutes=p["minutes"],
                    cash_paid=row.cash_paid,
                    nearest_hotcake_time=nearest_time,
                    nearest_diff_minutes=int(nearest_diff) if nearest_diff is not None else None,
                )
            )

    return CashReconResult(
        period=period,
        store=store,
        missing_bills=missing_bills,
        service_bill_rows=service_bill_rows,
        topup_bill_rows=topup_bill_rows,
        hotcake_time_mismatches=hotcake_time_mismatches,
        pos_time_mismatches=pos_time_mismatches,
        totals=CashTotals(
            hotcake_service_cash=hotcake_service_cash,
            hotcake_topup_cash=hotcake_topup_cash,
            hotcake_cash_total=hotcake_cash_total,
            pos_cash_total=pos_cash_total,
            pos_cash_diff=pos_cash_diff,
        ),
    )
