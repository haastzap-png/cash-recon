from __future__ import annotations

from datetime import datetime
from io import BytesIO
import inspect

import streamlit as st

from cash_recon.detect import detect_xlsx_kind
from cash_recon.logic import build_cash_recon
from cash_recon.models import Period
from cash_recon.parse import load_hotcake_bills_xlsx, load_hotcake_orders_xlsx, load_pos_history_orders_xlsx

try:
    from cash_recon.parse import load_card_machine_xlsx
except ImportError:
    load_card_machine_xlsx = None
from cash_recon.report import build_cash_recon_workbook


def _reset_app():
    st.session_state["upload_key"] = st.session_state.get("upload_key", 0) + 1
    st.session_state.pop("result_ready", None)
    st.rerun()


def _build_cash_recon_compat(*, period, store, orders, bills, pos_orders, card_machine_rows, topup_mode, time_tolerance_minutes):
    kwargs = {
        "period": period,
        "store": store,
        "orders": orders,
        "bills": bills,
        "pos_orders": pos_orders,
        "card_machine_rows": card_machine_rows,
        "topup_mode": topup_mode,
        "time_tolerance_minutes": time_tolerance_minutes,
    }

    sig = inspect.signature(build_cash_recon)
    supported = set(sig.parameters.keys())

    # Backward-compat: some older versions use time_tolerance instead of time_tolerance_minutes.
    if "time_tolerance_minutes" not in supported and "time_tolerance" in supported:
        kwargs["time_tolerance"] = kwargs.pop("time_tolerance_minutes")

    filtered = {k: v for k, v in kwargs.items() if k in supported}
    return build_cash_recon(**filtered)


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
    store_options = ["中壢三光店", "桃園中平店", "楊梅四維店", "中壢體育園區店", "其他(手動輸入)"]
    store_choice = st.selectbox("分店", options=store_options, index=0)
    if store_choice == "其他(手動輸入)":
        store = st.text_input("分店名稱（手動輸入）", value="")
    else:
        store = store_choice

    st.caption("區間時間可直接手動輸入，或點選日曆/時間")
    default_start = datetime(2026, 1, 1, 0, 0)
    default_end = datetime(2026, 1, 31, 23, 59)
    start_date = st.date_input("區間開始日期", value=default_start.date())
    start_time = st.time_input("區間開始時間", value=default_start.time(), step=60)
    end_date = st.date_input("區間結束日期", value=default_end.date())
    end_time = st.time_input("區間結束時間", value=default_end.time(), step=60)

    topup_mode = "settlement_time"
    st.caption("儲值金：依結帳操作時間計入")
    time_tolerance = st.number_input("POS 對帳時間容忍(分鐘)", min_value=0, max_value=240, value=120)
    auto_clear = st.checkbox("下載後自動清空", value=True)
    if st.button("清空/重置"):
        _reset_app()

st.header("上傳檔案")
if "upload_key" not in st.session_state:
    st.session_state["upload_key"] = 0

uploads = st.file_uploader(
    "把 xlsx 拖進來（可一次上傳多個，系統會自動判斷：帳單紀錄/訂單(預約)/收銀機歷史訂單）",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploads_{st.session_state['upload_key']}",
)
card_machine_up = st.file_uploader(
    "智慧刷卡機交易紀錄（可選）",
    type=["xlsx"],
    key=f"card_{st.session_state['upload_key']}",
)
if load_card_machine_xlsx is None:
    st.info("目前版本未包含智慧刷卡機解析功能；如需使用此功能請更新程式碼後再試。")

hotcake_bills_up = None
hotcake_orders_up = None
pos_orders_up = None

if uploads:
    st.subheader("辨識結果")
    assignments = []
    for up in uploads:
        content = up.getvalue()
        d = detect_xlsx_kind(content)
        choice = st.selectbox(
            f"{up.name}（自動判斷：{d.kind}）",
            options=[
                ("自動判斷", "auto"),
                ("Hotcake 帳單紀錄", "hotcake_bills"),
                ("Hotcake 訂單/預約報表", "hotcake_orders"),
                ("收銀機 歷史訂單", "pos_orders"),
                ("忽略此檔", "ignore"),
            ],
            format_func=lambda x: x[0],
            key=f"type_{up.name}",
        )
        assignments.append((up, d.kind, choice[1]))

    for up, detected, manual in assignments:
        kind = detected if manual == "auto" else manual
        if kind == "hotcake_bills" and hotcake_bills_up is None:
            hotcake_bills_up = up
        elif kind == "hotcake_orders" and hotcake_orders_up is None:
            hotcake_orders_up = up
        elif kind == "pos_orders" and pos_orders_up is None:
            pos_orders_up = up

if not uploads:
    st.info("請先上傳檔案。至少需要 Hotcake：帳單紀錄 + 訂單/預約報表。")

run = st.button("產出現金對帳表", type="primary", disabled=not (hotcake_bills_up and hotcake_orders_up))

if run:
    start_dt = datetime.combine(start_date, start_time)
    end_dt = datetime.combine(end_date, end_time)
    if start_dt > end_dt:
        st.error("區間日期時間不正確：開始需早於結束。")
        st.stop()

    if not store.strip():
        st.error("請輸入分店名稱。")
        st.stop()

    with st.spinner("讀取報表中..."):
        bills_bytes = hotcake_bills_up.getvalue()
        orders_bytes = hotcake_orders_up.getvalue()
        pos_bytes = pos_orders_up.getvalue() if pos_orders_up else None
        card_bytes = card_machine_up.getvalue() if card_machine_up else None

        try:
            if card_bytes and load_card_machine_xlsx is None:
                raise ValueError("目前版本未包含智慧刷卡機解析功能，請更新程式碼後再試。")
            orders = load_hotcake_orders_xlsx(orders_bytes)
            bills = load_hotcake_bills_xlsx(bills_bytes)
            pos_orders = load_pos_history_orders_xlsx(pos_bytes) if pos_bytes else None
            card_rows = load_card_machine_xlsx(card_bytes) if card_bytes else None
        except Exception as e:
            st.error(f"讀取報表失敗：{e}")
            st.info("若表頭有變動，請更新 GitHub 版本後再試。")
            st.stop()

    with st.spinner("計算中..."):
        result = _build_cash_recon_compat(
            period=Period(start=start_dt, end=end_dt),
            store=store.strip(),
            orders=orders,
            bills=bills,
            pos_orders=pos_orders,
            card_machine_rows=card_rows,
            topup_mode=topup_mode,
            time_tolerance_minutes=int(time_tolerance),
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
    if result.totals.card_machine_total is not None:
        st.info(f"刷卡機實付合計：{result.totals.card_machine_total:,.0f}")
        card_mis_card = len([r for r in result.card_mismatches if r.source == "card"])
        card_mis_hotcake = len([r for r in result.card_mismatches if r.source == "hotcake"])
        st.info(f"刷卡機未匹配(卡機)：{card_mis_card} ；刷卡機未匹配(Hotcake)：{card_mis_hotcake}")

    if pos_orders is not None:
        st.subheader("時間容忍外的資料")
        c5, c6 = st.columns(2)
        c5.metric("Hotcake 未匹配筆數", f"{len(result.hotcake_time_mismatches)}")
        c6.metric("POS 未匹配筆數", f"{len(result.pos_time_mismatches)}")
        if result.hotcake_time_mismatches:
            st.dataframe(
                [
                    {
                        "分店": r.store,
                        "日期": r.service_start.strftime("%Y-%m-%d"),
                        "日期時間": r.service_start.strftime("%Y-%m-%d %H:%M:%S"),
                        "師傅": r.designer,
                        "服務": r.service,
                        "分鐘": r.minutes,
                        "帳單編號": r.bill_id,
                        "帳單金額": r.bill_amount,
                        "現金": r.cash,
                        "現金差額": r.cash_diff,
                        "最近POS時間": r.nearest_pos_time.strftime("%Y-%m-%d %H:%M:%S") if r.nearest_pos_time else "",
                        "時間差(分)": r.nearest_diff_minutes,
                        "原因": r.reason,
                    }
                    for r in result.hotcake_time_mismatches
                ],
                use_container_width=True,
            )
        if result.pos_time_mismatches:
            st.dataframe(
                [
                    {
                        "機台名稱": r.terminal_name,
                        "日期時間": r.created_time.strftime("%Y-%m-%d %H:%M:%S"),
                        "師傅": r.designer,
                        "商品名稱": r.product_name,
                        "分鐘": r.minutes,
                        "現金支付": r.cash_paid,
                        "現金差額": r.cash_diff,
                        "最近Hotcake時間": r.nearest_hotcake_time.strftime("%Y-%m-%d %H:%M:%S") if r.nearest_hotcake_time else "",
                        "時間差(分)": r.nearest_diff_minutes,
                        "原因": r.reason,
                    }
                    for r in result.pos_time_mismatches
                ],
                use_container_width=True,
            )

    wb = build_cash_recon_workbook(result)
    out_name = f"現金對帳表_{store.strip()}_{start_dt.strftime('%Y%m%d')}-{end_dt.strftime('%Y%m%d')}.xlsx"
    bio = BytesIO()
    wb.save(bio)
    report_bytes = bio.getvalue()

    download_kwargs = {}
    if auto_clear:
        download_kwargs["on_click"] = _reset_app
    st.download_button(
        "下載現金對帳表 .xlsx",
        data=report_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        **download_kwargs,
    )
