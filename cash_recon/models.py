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


@dataclass(frozen=True)
class HotcakeTimeMismatchRow:
    store: str
    service_start: datetime
    designer: str
    service: str
    minutes: Optional[int]
    bill_id: str
    bill_amount: float
    cash: float
    nearest_pos_time: Optional[datetime]
    nearest_diff_minutes: Optional[int]


@dataclass(frozen=True)
class PosTimeMismatchRow:
    terminal_name: str
    created_time: datetime
    designer: str
    product_name: str
    minutes: Optional[int]
    cash_paid: float
    nearest_hotcake_time: Optional[datetime]
    nearest_diff_minutes: Optional[int]
