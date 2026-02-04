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
    CardMatchRow,
    CardMismatchRow,
)
from .parse import HotcakeBills, OrdersRow, PosRow, CardMachineRow


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


def _normalize_pay_method(text: str) -> str:
    t = (text or "").replace(" ", "").replace("_", "").lower()
    if "linepay" in t or "linepay" in t or "linepay" in t or "linepay" in t or "linepay" in t:
        return "linepay"
    if "line" in t and "pay" in t:
        return "linepay"
    if "信用卡" in (text or "") or "credit" in t or "card" in t:
        return "credit_card"
    return ""


@dataclass(frozen=True)
class CashReconResult:
    period: Period
    store: str
    missing_bills: list[MissingBillRow]
    service_bill_rows: list[ServiceBillRow]
    topup_bill_rows: list[TopUpBillRow]
    hotcake_time_mismatches: list[HotcakeTimeMismatchRow]
    pos_time_mismatches: list[PosTimeMismatchRow]
    card_machine_rows: list[CardMachineRow]
    card_matches: list[CardMatchRow]
    card_mismatches: list[CardMismatchRow]
    totals: CashTotals


def build_cash_recon(
    *,
    period: Period,
    store: str,
    orders: list[OrdersRow],
    bills: HotcakeBills,
    pos_orders: Optional[list[PosRow]] = None,
    card_machine_rows: Optional[list[CardMachineRow]] = None,
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
                    credit_card=r.credit_card,
                    linepay=r.linepay,
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
                    credit_card=r.credit_card,
                    linepay=r.linepay,
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

    # Card machine
    card_machine_total: Optional[float] = None
    scoped_card = []
    if card_machine_rows is not None:
        scoped_card = [
            r
            for r in card_machine_rows
            if _normalize_store(r.store) == store_norm and _in_period(r.transaction_time, period)
        ]
        card_machine_total = sum(r.paid_amount for r in scoped_card)

    # Card machine matching to Hotcake (credit card / linepay)
    card_matches: list[CardMatchRow] = []
    card_mismatches: list[CardMismatchRow] = []
    if scoped_card:
        hotcake_card_candidates = []
        for r in service_bill_rows:
            if r.credit_card and r.credit_card > 0:
                hotcake_card_candidates.append(
                    {
                        "bill_id": r.bill_id,
                        "pay_type": "credit_card",
                        "amount": r.credit_card,
                        "time": r.settlement_time,
                    }
                )
            if r.linepay and r.linepay > 0:
                hotcake_card_candidates.append(
                    {
                        "bill_id": r.bill_id,
                        "pay_type": "linepay",
                        "amount": r.linepay,
                        "time": r.settlement_time,
                    }
                )

        used_hotcake = set()

        for card in scoped_card:
            pay_type = _normalize_pay_method(card.pay_method)
            if pay_type not in ("credit_card", "linepay"):
                card_mismatches.append(
                    CardMismatchRow(
                        source="card",
                        store=card.store,
                        pay_type=card.pay_method,
                        amount=card.paid_amount,
                        time=card.transaction_time,
                        bill_id="",
                        nearest_time=None,
                        nearest_diff_minutes=None,
                    )
                )
                continue

            best_idx = None
            best_diff = None
            nearest_time = None
            nearest_diff = None
            for idx, h in enumerate(hotcake_card_candidates):
                if idx in used_hotcake:
                    continue
                if h["pay_type"] != pay_type:
                    continue
                if abs(h["amount"] - card.paid_amount) > 0.0001:
                    continue
                diff = abs((card.transaction_time - h["time"]).total_seconds() / 60)
                if nearest_diff is None or diff < nearest_diff:
                    nearest_diff = diff
                    nearest_time = h["time"]
                if best_diff is None or diff < best_diff:
                    best_diff = diff
                    best_idx = idx

            if best_idx is not None and best_diff is not None and best_diff <= time_tolerance_minutes:
                h = hotcake_card_candidates[best_idx]
                used_hotcake.add(best_idx)
                card_matches.append(
                    CardMatchRow(
                        store=store,
                        bill_id=h["bill_id"],
                        pay_type=pay_type,
                        hotcake_amount=h["amount"],
                        hotcake_time=h["time"],
                        card_amount=card.paid_amount,
                        card_time=card.transaction_time,
                        time_diff_minutes=int(best_diff),
                    )
                )
            else:
                card_mismatches.append(
                    CardMismatchRow(
                        source="card",
                        store=card.store,
                        pay_type=pay_type,
                        amount=card.paid_amount,
                        time=card.transaction_time,
                        bill_id="",
                        nearest_time=nearest_time,
                        nearest_diff_minutes=int(nearest_diff) if nearest_diff is not None else None,
                    )
                )

        for idx, h in enumerate(hotcake_card_candidates):
            if idx in used_hotcake:
                continue
            card_mismatches.append(
                CardMismatchRow(
                    source="hotcake",
                    store=store,
                    pay_type=h["pay_type"],
                    amount=h["amount"],
                    time=h["time"],
                    bill_id=h["bill_id"],
                    nearest_time=None,
                    nearest_diff_minutes=None,
                )
            )

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
                        cash_diff=None,
                        nearest_pos_time=None,
                        nearest_diff_minutes=None,
                        reason="設計師缺失",
                    )
                )
                continue
            # candidate POS rows: same designer, minutes match if both present
            best_idx = None
            best_diff = None
            best_cash_match = None
            nearest_diff = None
            nearest_time = None
            nearest_cash = None
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
                    nearest_cash = p["row"].cash_paid
                if best_diff is None or diff < best_diff:
                    best_diff = diff
                    best_idx = idx
                    best_cash_match = p["row"].cash_paid

            hotcake_cash = bill_cash_map.get(o.bill_id, 0.0)
            if (
                best_idx is not None
                and best_diff is not None
                and best_diff <= time_tolerance_minutes
                and best_cash_match is not None
                and abs(best_cash_match - hotcake_cash) < 0.0001
            ):
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
                    cash=hotcake_cash,
                    cash_diff=(nearest_cash - hotcake_cash) if nearest_cash is not None else None,
                    nearest_pos_time=nearest_time,
                    nearest_diff_minutes=int(nearest_diff) if nearest_diff is not None else None,
                    reason="時間超過容忍或金額不符",
                )
            )

        for idx, p in enumerate(pos_candidates):
            if idx in used_pos:
                continue
            row = p["row"]
            # find nearest hotcake order for context
            nearest_diff = None
            nearest_time = None
            nearest_cash = None
            for o in scoped_orders:
                if _normalize_name(o.designer) != p["designer_norm"]:
                    continue
                diff = abs((row.created_time - o.service_start).total_seconds() / 60)
                if nearest_diff is None or diff < nearest_diff:
                    nearest_diff = diff
                    nearest_time = o.service_start
                    nearest_cash = bill_cash_map.get(o.bill_id, 0.0)
            pos_time_mismatches.append(
                PosTimeMismatchRow(
                    terminal_name=row.terminal_name,
                    created_time=row.created_time,
                    designer=_parse_pos_designer(row.product_name),
                    product_name=row.product_name,
                    minutes=p["minutes"],
                    cash_paid=row.cash_paid,
                    cash_diff=(row.cash_paid - nearest_cash) if nearest_cash is not None else None,
                    nearest_hotcake_time=nearest_time,
                    nearest_diff_minutes=int(nearest_diff) if nearest_diff is not None else None,
                    reason="時間超過容忍或金額不符",
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
        card_machine_rows=scoped_card,
        card_matches=card_matches,
        card_mismatches=card_mismatches,
        totals=CashTotals(
            hotcake_service_cash=hotcake_service_cash,
            hotcake_topup_cash=hotcake_topup_cash,
            hotcake_cash_total=hotcake_cash_total,
            pos_cash_total=pos_cash_total,
            pos_cash_diff=pos_cash_diff,
            card_machine_total=card_machine_total,
        ),
    )
