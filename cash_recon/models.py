from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass(frozen=True)
class Period:
    start: datetime
    end: datetime


@dataclass(frozen=True)
class CashTotals:
    hotcake_service_cash: float
    hotcake_topup_cash: float
    hotcake_cash_total: float
    pos_cash_total: Optional[float]
    pos_cash_diff: Optional[float]


@dataclass(frozen=True)
class MissingBillRow:
    store: str
    order_id: str
    order_code: str
    service_start: datetime
    designer: str
    service: str
    order_status: str
    checkin_time: Optional[datetime]
    member_name: str
    phone: str


@dataclass(frozen=True)
class ServiceBillRow:
    store: str
    bill_id: str
    settlement_time: datetime
    attributed_date: datetime
    designer: str
    item: str
    cash: float
    bill_amount: float


@dataclass(frozen=True)
class TopUpBillRow:
    store: str
    bill_id: str
    settlement_time: datetime
    attributed_date: datetime
    designer: str
    item: str
    cash: float
    bill_amount: float

