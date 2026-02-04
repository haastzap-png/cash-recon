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
    card_machine_total: Optional[float]


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
    credit_card: float
    linepay: float
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
    credit_card: float
    linepay: float
    bill_amount: float


@dataclass(frozen=True)
class CardMachineRow:
    order_id: str
    store: str
    device_name: str
    amount: float
    paid_amount: float
    transaction_time: datetime
    pay_method: str


@dataclass(frozen=True)
class CardMatchRow:
    store: str
    bill_id: str
    pay_type: str
    hotcake_amount: float
    hotcake_time: datetime
    card_amount: float
    card_time: datetime
    time_diff_minutes: int


@dataclass(frozen=True)
class CardMismatchRow:
    source: str
    store: str
    pay_type: str
    amount: float
    time: datetime
    bill_id: str
    nearest_time: Optional[datetime]
    nearest_diff_minutes: Optional[int]


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
    cash_diff: Optional[float]
    nearest_pos_time: Optional[datetime]
    nearest_diff_minutes: Optional[int]
    reason: str


@dataclass(frozen=True)
class PosTimeMismatchRow:
    terminal_name: str
    created_time: datetime
    designer: str
    product_name: str
    minutes: Optional[int]
    cash_paid: float
    cash_diff: Optional[float]
    nearest_hotcake_time: Optional[datetime]
    nearest_diff_minutes: Optional[int]
    reason: str
