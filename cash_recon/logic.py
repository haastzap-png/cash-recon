from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional

from .models import CashTotals, MissingBillRow, Period, ServiceBillRow, TopUpBillRow
from .parse import HotcakeBills, OrdersRow, PosRow


def _in_period(dt: datetime, period: Period) -> bool:
    return period.start <= dt <= period.end


@dataclass(frozen=True)
class CashReconResult:
    period: Period
    store: str
    missing_bills: list[MissingBillRow]
    service_bill_rows: list[ServiceBillRow]
    topup_bill_rows: list[TopUpBillRow]
    totals: CashTotals


def build_cash_recon(
    *,
    period: Period,
    store: str,
    orders: list[OrdersRow],
    bills: HotcakeBills,
    pos_orders: Optional[list[PosRow]] = None,
) -> CashReconResult:
    scoped_orders = [o for o in orders if o.store == store and _in_period(o.service_start, period)]

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
    for bill_id in sorted(bill_ids_in_scope):
        r = bills.service.get(bill_id)
        if r is None:
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
    for r in bills.topup.values():
        if r.store != store:
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

    pos_cash_total: Optional[float] = None
    pos_cash_diff: Optional[float] = None
    if pos_orders is not None:
        scoped_pos = [p for p in pos_orders if _in_period(p.created_time, period)]
        pos_cash_total = sum(p.cash_paid for p in scoped_pos)
        pos_cash_diff = pos_cash_total - hotcake_cash_total

    return CashReconResult(
        period=period,
        store=store,
        missing_bills=missing_bills,
        service_bill_rows=service_bill_rows,
        topup_bill_rows=topup_bill_rows,
        totals=CashTotals(
            hotcake_service_cash=hotcake_service_cash,
            hotcake_topup_cash=hotcake_topup_cash,
            hotcake_cash_total=hotcake_cash_total,
            pos_cash_total=pos_cash_total,
            pos_cash_diff=pos_cash_diff,
        ),
    )
