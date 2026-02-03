from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Optional

import streamlit as st

from cash_recon.detect import detect_xlsx_kind
from cash_recon.logic import build_cash_recon
from cash_recon.models import Period
from cash_recon.parse import load_hotcake_bills_xlsx, load_hotcake_orders_xlsx, load_pos_history_orders_xlsx
from cash_recon.report import build_cash_recon_workbook


def _parse_dt(text: str) -> Optional[datetime]:
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


st.set_page_config(page_title="現金對帳表", layout="wide")
st.title("現金對帳表 (MVP)")

st.markdown(
    """
上傳 Hotcake 的 **帳單紀錄** + **訂單(預約)報表**，可選上傳 **收銀機歷史訂單**。
輸出重點：該區間每間店 **要收回來多少現金**，以及 **漏結帳清單** 是否為空。
"""
)

with st.sidebar:
    st.header("輸入")
    store = st.text_input("分店名稱（需與報表一致）", value="中壢三光店")
    start_text = st.text_input("區間開始 (YYYY-MM-DD HH:MM)", value="2026-01-01 00:00")
    end_text = st.text_input("區間結束 (YYYY-MM-DD HH:MM)", value="2026-01-31 23:59")

st.header("上傳檔案")
uploads = st.file_uploader(
    "把 xlsx 拖進來（可一次上傳多個，系統會自動判斷：帳單紀錄/訂單(預約)/收銀機歷史訂單）",
    type=["xlsx"],
    accept_multiple_files=True,
)

hotcake_bills_up = None
hotcake_orders_up = None
pos_orders_up = None

if uploads:
    st.subheader("辨識結果")
    for up in uploads:
        content = up.getvalue()
        d = detect_xlsx_kind(content)
        st.write(f"- `{up.name}` → `{d.kind}`（{d.reason}）")
        if d.kind == "hotcake_bills" and hotcake_bills_up is None:
            hotcake_bills_up = up
        elif d.kind == "hotcake_orders" and hotcake_orders_up is None:
            hotcake_orders_up = up
        elif d.kind == "pos_orders" and pos_orders_up is None:
            pos_orders_up = up

if not uploads:
    st.info("請先上傳檔案。至少需要 Hotcake：帳單紀錄 + 訂單/預約報表。")

run = st.button("產出現金對帳表", type="primary", disabled=not (hotcake_bills_up and hotcake_orders_up))

if run:
    start_dt = _parse_dt(start_text)
    end_dt = _parse_dt(end_text)
    if start_dt is None or end_dt is None or start_dt > end_dt:
        st.error("區間日期時間格式不正確，請用 YYYY-MM-DD HH:MM 或 YYYY-MM-DD HH:MM:SS，且開始需早於結束。")
        st.stop()

    if not store.strip():
        st.error("請輸入分店名稱。")
        st.stop()

    with st.spinner("讀取報表中..."):
        bills_bytes = hotcake_bills_up.getvalue()
        orders_bytes = hotcake_orders_up.getvalue()
        pos_bytes = pos_orders_up.getvalue() if pos_orders_up else None

        orders = load_hotcake_orders_xlsx(orders_bytes)
        bills = load_hotcake_bills_xlsx(bills_bytes)
        pos_orders = load_pos_history_orders_xlsx(pos_bytes) if pos_bytes else None

    with st.spinner("計算中..."):
        result = build_cash_recon(
            period=Period(start=start_dt, end=end_dt),
            store=store.strip(),
            orders=orders,
            bills=bills,
            pos_orders=pos_orders,
        )

    st.header("結果")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Hotcake 服務現金", f"{result.totals.hotcake_service_cash:,.0f}")
    c2.metric("Hotcake 儲值金現金", f"{result.totals.hotcake_topup_cash:,.0f}")
    c3.metric("Hotcake 現金合計", f"{result.totals.hotcake_cash_total:,.0f}")
    if result.totals.pos_cash_total is not None:
        c4.metric("收銀機現金合計", f"{result.totals.pos_cash_total:,.0f}")
    else:
        c4.metric("收銀機現金合計", "（未上傳）")

    if result.missing_bills:
        st.error(f"漏結帳清單不為空：{len(result.missing_bills)} 筆。依你的規則，現金對帳表不可視為正確。")
        st.dataframe(
            [
                {
                    "分店": r.store,
                    "日期": r.service_start.strftime("%Y-%m-%d"),
                    "日期時間(服務開始)": r.service_start.strftime("%Y-%m-%d %H:%M:%S"),
                    "師傅": r.designer,
                    "服務": r.service,
                    "訂單編號": r.order_id,
                    "訂單代碼": r.order_code,
                    "會員姓名": r.member_name,
                    "手機號碼": r.phone,
                }
                for r in result.missing_bills
            ],
            use_container_width=True,
        )
    else:
        st.success("漏結帳清單為空：可視為正確現金對帳表。")

    if result.totals.pos_cash_diff is not None:
        st.info(f"收銀機現金 - Hotcake 現金合計：{result.totals.pos_cash_diff:,.0f}")

    wb = build_cash_recon_workbook(result)
    out_name = f"現金對帳表_{store.strip()}_{start_dt.strftime('%Y%m%d')}-{end_dt.strftime('%Y%m%d')}.xlsx"
    bio = BytesIO()
    wb.save(bio)
    report_bytes = bio.getvalue()

    st.download_button(
        "下載現金對帳表 .xlsx",
        data=report_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
