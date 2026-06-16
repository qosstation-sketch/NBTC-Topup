import calendar
import datetime
import io
import os
import re

import gspread
import pandas as pd
import requests
import streamlit as st
from google.oauth2.service_account import Credentials

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except Exception:
    A4 = None
    pdfmetrics = None
    TTFont = None

THAI_FONT_PATHS = [
    "THSarabunNew.ttf",                   # ⭐️ สำหรับตอนรันบนเว็บ Streamlit Cloud
    "THSarabun.ttf",                      # ⭐️ เผื่อไว้
    r"C:\Windows\Fonts\THSarabunNew.ttf", # สำหรับตอนรันรันในคอมตัวเอง
    r"C:\Windows\Fonts\THSarabun.ttf"
]
THAI_PDF_FONT = None

if pdfmetrics and TTFont:
    for font_path in THAI_FONT_PATHS:
        try:
            if os.path.exists(font_path):
                font_name = os.path.splitext(os.path.basename(font_path))[0]
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                THAI_PDF_FONT = font_name
                break
        except Exception:
            continue


SHEET_URL = "https://docs.google.com/spreadsheets/d/1rvc8juySIft-YPNu5X5VQ4-0ksei9XSEktqJoQBEhCc/edit?usp=sharing"
WORKSHEET_NAME = "StockV3"
OPERATORS = ["DTN", "AWN", "TUC", "CAT"]
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

OPERATOR_COLORS = {
    "AWN": {"bg": "#DCFCE7", "border": "#22C55E", "text": "#166534"},
    "DTN": {"bg": "#DBEAFE", "border": "#3B82F6", "text": "#1D4ED8"},
    "TUC": {"bg": "#FEE2E2", "border": "#EF4444", "text": "#991B1B"},
    "CAT": {"bg": "#FFEDD5", "border": "#F97316", "text": "#9A3412"},
}
DEFAULT_OPERATOR_COLOR = {"bg": "#F1F5F9", "border": "#94A3B8", "text": "#334155"}


def inject_styles():
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }
        .dashboard-header {
            border: none;
            background: transparent;
            padding: 0;
            margin-bottom: 0;
        }
        .dashboard-header h1 {
            margin: 0 0 4px 0;
            font-size: 1.7rem;
            line-height: 1.2;
            color: #111827;
        }
        .dashboard-header p {
            margin: 0;
            color: #64748b;
        }
        .operator-legend {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin: 12px 0 20px 0;
        }
        .operator-chip {
            display: inline-flex;
            align-items: center;
            border: 1.5px solid var(--operator-border);
            background: var(--operator-bg);
            color: var(--operator-text);
            border-radius: 999px;
            padding: 8px 16px;
            font-weight: 900;
            font-size: 1.05rem;
            box-shadow: 0 2px 6px rgba(0,0,0,0.18);
        }
        .operator-card {
            border: 1px solid var(--operator-border);
            border-left: 6px solid var(--operator-border);
            border-radius: 8px;
            background: var(--operator-bg);
            padding: 12px 12px;
            min-height: 74px;
        }
        .operator-card .operator-name {
            color: var(--operator-text);
            font-weight: 900;
            font-size: 0.95rem;
        }
        .operator-card .operator-count {
            color: #111827;
            font-weight: 900;
            font-size: 1.35rem;
            line-height: 1.1;
            margin-top: 5px;
        }
        .operator-card .operator-value {
            color: #475569;
            font-size: 0.82rem;
            margin-top: 4px;
        }
        div[data-testid="stMetric"] {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            background: #ffffff;
            padding: 10px 12px;
            min-height: 80px;
        }
        div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
            font-size: 0.95rem;
        }
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            font-size: 1.25rem;
            font-weight: 700;
        }
        div[data-testid="stMetric"] label,
        div[data-testid="stMetric"] [data-testid="stMetricDelta"],
        div[data-testid="stMetric"] p {
            color: #111827 !important;
        }
        .top-label {
            font-size: 0.86rem;
            color: #94a3b8;
            margin-bottom: 0.75rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def operator_style(operator):
    return OPERATOR_COLORS.get(str(operator).strip().upper(), DEFAULT_OPERATOR_COLOR)


def render_operator_legend():
    chips = []
    for operator in OPERATORS:
        colors = operator_style(operator)
        chips.append(
            f"<span class=\"operator-chip\" style=\"--operator-bg:{colors['bg']};--operator-border:{colors['border']};--operator-text:{colors['text']};\">"
            f"{operator}"
            "</span>"
        )
    st.markdown(f"<div class='operator-legend'>{''.join(chips)}</div>", unsafe_allow_html=True)


def render_operator_card(operator, count, value):
    colors = operator_style(operator)
    st.markdown(
        f"""
        <div class="operator-card" style="--operator-bg:{colors['bg']};--operator-border:{colors['border']};--operator-text:{colors['text']};">
            <div class="operator-name">{operator}</div>
            <div class="operator-count">{count:,.0f} ใบ</div>
            <div class="operator-value">มูลค่ารวม {value:,.0f} บาท</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def style_operator_rows(dataframe, operator_column):
    def row_style(row):
        colors = operator_style(row.get(operator_column, ""))
        return [f"background-color: {colors['bg']}; color: #111827;" for _ in row]

    return dataframe.style.apply(row_style, axis=1)


def authorize_google_sheets():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
        return gspread.authorize(creds)
    except Exception:
        try:
            creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPE)
            return gspread.authorize(creds)
        except FileNotFoundError:
            st.error("🚨 ไม่พบไฟล์รหัสผ่าน กรุณาตั้งค่า Secrets หรือวางไฟล์ credentials.json")
            st.stop()


def safe_rerun():
    """Try to rerun the Streamlit script using available API, otherwise no-op."""
    try:
        st.experimental_rerun()
        return
    except Exception:
        pass
    try:
        if hasattr(st, "rerun"):
            st.rerun()
            return
    except Exception:
        pass


def open_stock_sheet(client):
    try:
        spreadsheet = client.open_by_url(SHEET_URL)
    except requests.exceptions.RequestException as err:
        st.error(
            "ไม่สามารถเชื่อมต่อ Google Sheets ได้ กรุณาตรวจสอบการเชื่อมต่ออินเทอร์เน็ตหรือการตั้งค่า DNS ของคุณ"
        )
        st.error(str(err))
        st.stop()
    except Exception as err:
        st.error(f"เกิดข้อผิดพลาดในการเปิด Google Sheets: {err}")
        st.stop()

    try:
        return spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        st.error(f"ไม่พบ worksheet ชื่อ {WORKSHEET_NAME}")
        st.stop()


def load_sheet_dataframe(worksheet):
    values = worksheet.get_all_values()
    if not values:
        return pd.DataFrame()

    headers = [header.strip() for header in values[0]]
    row_width = len(headers)
    rows = []
    sheet_rows = []

    for sheet_row, row in enumerate(values[1:], start=2):
        padded_row = row + [""] * (row_width - len(row))
        padded_row = padded_row[:row_width]
        if any(str(cell).strip() for cell in padded_row):
            rows.append(padded_row)
            sheet_rows.append(sheet_row)

    dataframe = pd.DataFrame(rows, columns=headers)
    if not dataframe.empty:
        dataframe["sheet_row"] = sheet_rows
    return dataframe


def find_column(columns, *names):
    normalized_columns = {column.strip().lower(): column for column in columns}
    for name in names:
        column = normalized_columns.get(name.strip().lower())
        if column:
            return column
    return None


def parse_expired_dates(series):
    parsed = pd.to_datetime(series, format="%m/%d/%y", errors="coerce")
    missing = parsed.isna()
    parsed.loc[missing] = pd.to_datetime(series.loc[missing], format="%m/%d/%Y", errors="coerce")
    missing = parsed.isna()
    parsed.loc[missing] = pd.to_datetime(series.loc[missing], format="%d/%m/%Y", errors="coerce")
    return parsed


def require_columns(column_map, required_names, section_name):
    missing = [name for name in required_names if not column_map.get(name)]
    if missing:
        st.error(f"ไม่พบคอลัมน์ที่จำเป็นสำหรับ{section_name}: {', '.join(missing)}")
        return False
    return True


def normalize_card_id(value):
    return str(value).strip()


def get_added_by_column(dataframe):
    if dataframe.empty:
        return None
    column = find_column(
        dataframe.columns,
        "ผู้เพิ่มบัตร",
        "Added By",
        "AddedBy",
        "Adder",
        "Created By",
        "Creator",
    )
    if column:
        return column
    return dataframe.columns[1] if len(dataframe.columns) > 1 else None


def get_added_by_options(dataframe):
    added_by_column = get_added_by_column(dataframe)
    if not added_by_column:
        return []

    names = []
    seen = set()
    for value in dataframe[added_by_column].dropna().astype(str):
        name = value.strip()
        if not name:
            continue
        normalized = name.casefold()
        if normalized in seen:
            continue
        seen.add(normalized)
        names.append(name)
    return sorted(names, key=str.casefold)


def get_requester_options(dataframe, column_map):
    user_column = column_map.get("User")
    if dataframe.empty or not user_column:
        return []

    names = []
    seen = set()
    for value in dataframe[user_column].dropna().astype(str):
        name = value.strip()
        if not name:
            continue
        normalized = name.casefold()
        if normalized in seen:
            continue
        seen.add(normalized)
        names.append(name)
    return sorted(names, key=str.casefold)


def get_price_options(dataframe, column_map):
    price_column = column_map.get("Price")
    if dataframe.empty or not price_column:
        return []

    prices = pd.to_numeric(dataframe[price_column], errors="coerce").dropna()
    return [int(price) if float(price).is_integer() else float(price) for price in sorted(prices.unique())]


def parse_price_input(value):
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    return int(number) if number.is_integer() else number


def get_existing_card_ids(dataframe, column_map):
    id_column = column_map.get("ID")
    if dataframe.empty or not id_column:
        return set()
    return {
        card_id
        for card_id in dataframe[id_column].map(normalize_card_id)
        if card_id
    }


def get_next_card_id(dataframe, column_map, operator):
    if dataframe.empty or not column_map.get("Operator") or not column_map.get("ID"):
        return "1"

    operator_rows = dataframe[
        dataframe[column_map["Operator"]].astype(str).str.strip().str.upper() == str(operator).strip().upper()
    ]
    if operator_rows.empty:
        return "1"

    best_number = None
    best_width = 0
    for raw_card_id in operator_rows[column_map["ID"]].dropna().astype(str):
        match = re.search(r"(\d+)", raw_card_id.strip())
        if not match:
            continue

        number_text = match.group(1)
        number = int(number_text)
        if best_number is None or number > best_number:
            best_number = number
            best_width = len(number_text)

    if best_number is None:
        return "1"

    next_number = best_number + 1
    return str(next_number).zfill(best_width) if best_width > 1 else str(next_number)


def build_withdraw_updates(sheet_rows, withdraw_date, requester, job_no):
    updates = []
    for sheet_row in sheet_rows:
        updates.extend(
            [
                {"range": f"I{sheet_row}", "values": [[True]]},
                {"range": f"J{sheet_row}", "values": [[withdraw_date]]},
                {"range": f"K{sheet_row}", "values": [[requester]]},
                {"range": f"L{sheet_row}", "values": [[job_no]]},
            ]
        )
    return updates


def build_column_map(df):
    return {
        "Status": find_column(df.columns, "Status", "สถานะ"),
        "Operator": find_column(df.columns, "Operator", "เครือข่าย", "ผู้ให้บริการ"),
        "Price": find_column(df.columns, "Price", "ราคา", "Amount", "Value", "ราคาบัตร"),
        "ID": find_column(
            df.columns,
            "ID",
            "Card ID",
            "Card No",
            "Card Number",
            "เลขที่บัตร",
            "เลขบัตร",
        ),
        "Expired date": find_column(df.columns, "Expired date", "Expired Date", "วันหมดอายุ"),
        "Day(s) left": find_column(
            df.columns,
            "Day(s) left",
            "Days left",
            "Day left",
            "จำนวนวัน",
            "วันเหลือ",
            "วันคงเหลือ",
        ),
        "Withdraw Date": find_column(
            df.columns,
            "วันที่ขอเบิก",
            "วันที่เบิก",
            "Withdraw Date",
            "WithdrawDate",
            "Date",
            "วันที่",
        ),
        "User": find_column(
            df.columns,
            "ผู้ใช้บัตร",
            "ผู้ขอเบิก",
            "ผู้เบิก",
            "Requester",
            "Requestor",
            "Requestee",
            "User",
            "ชื่อผู้ใช้",
            "ชื่อผู้ขอเบิก",
            "ชื่อผู้ใช้บัตร",
        ),
        "Job No": find_column(
            df.columns,
            "Job No.",
            "Job No",
            "Job",
            "JobNo",
            "Job number",
            "เลข Job",
            "Job ID",
        ),
        "Used?": find_column(df.columns, "Used?", "Used", "Withdrawn"),
    }


def format_used_display(status):
    status_text = str(status).strip().lower()
    if status_text in ["หมดอายุแล้ว", "expired"]:
        return "🔴 หมดอายุ"
    if status_text in ["ใช้งานแล้ว", "used"]:
        return "✅ ใช้งานแล้ว"
    return "🟩 พร้อมใช้งาน"


def render_stock_summary(df, columns):
    st.subheader("ยอดบัตรคงเหลือในระบบ")
    if df.empty:
        st.info("ยังไม่มีข้อมูลบัตรในระบบ")
        return

    if not require_columns(columns, ["Status", "Operator", "Price"], "สรุปยอด"):
        return

    available_df = df[df[columns["Status"]].isin(["ปกติ", "ใกล้หมดอายุ"])].copy()
    available_df[columns["Price"]] = pd.to_numeric(available_df[columns["Price"]], errors="coerce").fillna(0)

    total_cards = len(available_df)
    total_value = available_df[columns["Price"]].sum()
    near_expiry = (available_df[columns["Status"]] == "ใกล้หมดอายุ").sum()

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("พร้อมเบิก", f"{total_cards:,.0f} ใบ")
    metric2.metric("มูลค่ารวม", f"{total_value:,.0f} บาท")
    metric3.metric("ใกล้หมดอายุ", f"{near_expiry:,.0f} ใบ")

    operator_summary = (
        available_df.groupby(columns["Operator"])
        .agg(
            card_count=(columns["Operator"], "size"),
            total_value=(columns["Price"], "sum"),
        )
        .reset_index()
    )

    card_columns = st.columns(4)
    for index, operator in enumerate(OPERATORS):
        operator_rows = operator_summary[
            operator_summary[columns["Operator"]].astype(str).str.upper() == operator
        ]
        count = int(operator_rows["card_count"].sum()) if not operator_rows.empty else 0
        value = float(operator_rows["total_value"].sum()) if not operator_rows.empty else 0
        with card_columns[index]:
            render_operator_card(operator, count, value)

    available_cards = (
        available_df
        .groupby([columns["Operator"], columns["Price"]])
        .size()
        .reset_index(name="คงเหลือ (ใบ)")
        .sort_values([columns["Operator"], columns["Price"]])
    )
    card_table_height = min(520, max(120, 42 + available_cards.shape[0] * 34))
    st.dataframe(
        style_operator_rows(available_cards, columns["Operator"]),
        width="stretch",
        height=card_table_height,
        hide_index=True,
    )


def render_add_card_form(sheet):
    try:
        id_col = columns.get("ID") if "columns" in globals() else None
        chips = []
        if id_col and not df.empty:
            for operator in OPERATORS:
                next_id = get_next_card_id(df, columns, operator)
                colors = operator_style(operator)
                chips.append(
                    f"<span class=\"operator-chip\" style=\"--operator-bg:{colors['bg']};--operator-border:{colors['border']};--operator-text:{colors['text']};\">{operator} — ถัดไป: {next_id}</span>"
                )
        else:
            for operator in OPERATORS:
                colors = operator_style(operator)
                chips.append(
                    f"<span class=\"operator-chip\" style=\"--operator-bg:{colors['bg']};--operator-border:{colors['border']};--operator-text:{colors['text']};\">{operator}</span>"
                )
        st.markdown(f"<div class='operator-legend'>{''.join(chips)}</div>", unsafe_allow_html=True)
    except Exception:
        render_operator_legend()

    st.subheader("เพิ่มบัตรใหม่เข้าสต็อก")
    operator = st.selectbox("ผู้ให้บริการ", OPERATORS, key="add_card_operator")
    suggested_card_id = get_next_card_id(df, columns, operator)
    added_by_options = get_added_by_options(df)
    price_options = get_price_options(df, columns)

    with st.form("add_card_form", clear_on_submit=True):
        col1, col2, col3 = st.columns([1.1, 1, 1.1])
        with col1:
            if added_by_options:
                added_by = st.selectbox(
                    "ผู้เพิ่มบัตร",
                    added_by_options,
                    index=None,
                    placeholder="ค้นหาหรือเลือกผู้เพิ่มบัตร",
                    accept_new_options=True,
                    filter_mode="fuzzy",
                )
            else:
                added_by = st.text_input("ผู้เพิ่มบัตร", placeholder="ชื่อผู้เพิ่มบัตร")
        with col2:
            if price_options:
                price = st.selectbox(
                    "ราคาบัตร",
                    price_options,
                    index=None,
                    placeholder="ค้นหาหรือเลือกราคา",
                    accept_new_options=True,
                    filter_mode="fuzzy",
                )
            else:
                price = st.number_input("ราคาบัตร", min_value=0, step=10)
        with col3:
            expiry_month = st.selectbox(
                "เดือนหมดอายุ",
                list(range(1, 13)),
                format_func=lambda month: f"{month:02d}",
            )
        expiry_year = st.number_input(
            "ปีหมดอายุ",
            min_value=datetime.date.today().year,
            max_value=2100,
            value=datetime.date.today().year,
            step=1,
        )
        card_id = st.text_input(
            "เลขที่บัตร",
            value=suggested_card_id,
            key=f"add_card_id_{operator}_{suggested_card_id}",
        )

        submitted = st.form_submit_button("บันทึกเข้าสต็อก", type="primary")
        if not submitted:
            return

        added_by = str(added_by or "").strip()
        card_id = normalize_card_id(card_id)
        price = parse_price_input(price)

        if not card_id:
            st.error("กรุณากรอกเลขที่บัตร")
            return
        if price is None or price < 0:
            st.error("กรุณาเลือกราคาบัตรหรือกรอกราคาเป็นตัวเลข")
            return

        latest_df = load_sheet_dataframe(sheet)
        if not latest_df.empty:
            latest_df.columns = latest_df.columns.str.strip()
        latest_columns = build_column_map(latest_df)
        existing_ids = get_existing_card_ids(latest_df, latest_columns)
        if card_id in existing_ids:
            st.error(f"ไม่สามารถบันทึกได้: เลขที่บัตร {card_id} มีอยู่ในชีทแล้ว")
            return

        next_row = len(sheet.col_values(1)) + 1
        today = datetime.date.today().strftime("%m/%d/%y")
        last_day = calendar.monthrange(expiry_year, expiry_month)[1]
        expired_date = f"{expiry_month:02d}/{last_day:02d}/{expiry_year % 100:02d}"

        new_row_data = [
            operator,
            added_by,
            price,
            today,
            expired_date,
            card_id,
            f'=IF(I{next_row}=TRUE, "ใช้งานแล้ว", IF(E{next_row}="", "", IF(E{next_row}<TODAY(), "หมดอายุแล้ว", IF(E{next_row}-TODAY()<=30, "ใกล้หมดอายุ", "ปกติ"))))',
            f'=IF(G{next_row}="ใช้งานแล้ว", "-", E{next_row}-TODAY())',
            "FALSE",
            "",
            "",
            "",
        ]

        sheet.append_row(new_row_data, value_input_option="USER_ENTERED", table_range="A1:A")
        sheet.spreadsheet.batch_update(
            {
                "requests": [
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": sheet.id,
                                "startRowIndex": next_row - 1,
                                "endRowIndex": next_row,
                                "startColumnIndex": 8,
                                "endColumnIndex": 9,
                            },
                            "rule": {
                                "condition": {"type": "BOOLEAN"},
                                "strict": True,
                                "showCustomUi": True,
                            },
                        }
                    }
                ]
            }
        )

        st.success(f"บันทึกบัตรเลขที่ {card_id} เรียบร้อยแล้ว พร้อมสูตรและ Checkbox!")
        st.rerun()


def render_withdraw(df, sheet, columns):
    st.subheader("เบิกใช้บัตรเติมเงิน")

    if df.empty:
        st.info("ยังไม่มีข้อมูลบัตรในระบบ")
        return

    required = ["Status", "Operator", "Price", "ID", "Expired date", "Day(s) left"]
    if not require_columns(columns, required, "หน้าเบิก"):
        return

    active_cards = df[df[columns["Status"]].isin(["ปกติ", "ใกล้หมดอายุ"])].copy()
    active_cards[columns["Price"]] = pd.to_numeric(active_cards[columns["Price"]], errors="coerce").fillna(0)
    active_cards[columns["Day(s) left"]] = pd.to_numeric(active_cards[columns["Day(s) left"]], errors="coerce")
    if active_cards.empty:
        st.success("🎉 บัตรในสต็อกถูกเบิกหมดแล้ว")
        return

    filter1, filter2, filter3 = st.columns([1, 1, 1.4])
    with filter1:
        operator_filter = st.multiselect("เครือข่าย", options=OPERATORS, default=OPERATORS)
    with filter2:
        prices = sorted(active_cards[columns["Price"]].astype(int).unique().tolist())
        price_filter = st.selectbox("ราคา", ["ทั้งหมด"] + prices)
    with filter3:
        status_filter = st.multiselect(
            "สถานะ",
            ["ปกติ", "ใกล้หมดอายุ"],
            default=["ปกติ", "ใกล้หมดอายุ"],
        )

    if operator_filter:
        selected_ops = [op.upper() for op in operator_filter]
        active_cards = active_cards[active_cards[columns["Operator"]].astype(str).str.upper().isin(selected_ops)]
    if price_filter != "ทั้งหมด":
        active_cards = active_cards[active_cards[columns["Price"]] == price_filter]
    if status_filter:
        active_cards = active_cards[active_cards[columns["Status"]].isin(status_filter)]

    if active_cards.empty:
        st.info("ไม่พบบัตรที่ตรงกับเงื่อนไข")
        return

    active_cards = active_cards.sort_values(
        [columns["Day(s) left"], "sheet_row"],
        na_position="last",
    ).reset_index(drop=True)
    active_cards.insert(0, "เลือกเบิก", False)
    display_df = active_cards[
        [
            "เลือกเบิก",
            columns["Operator"],
            columns["Price"],
            columns["ID"],
            columns["Expired date"],
            columns["Day(s) left"],
            columns["Status"],
            "sheet_row",
        ]
    ]

    requester_options = get_requester_options(df, columns)
    col1, col2 = st.columns(2)
    with col1:
        if requester_options:
            requester = st.selectbox(
                "👤 ชื่อผู้ขอเบิก",
                requester_options,
                index=None,
                placeholder="ค้นหาหรือเลือกชื่อผู้ขอเบิก",
                accept_new_options=True,
                filter_mode="fuzzy",
            )
        else:
            requester = st.text_input("👤 ชื่อผู้ขอเบิก", placeholder="กรอกชื่อ-นามสกุล")
    with col2:
        job_no = st.text_input("💼 Job No.", placeholder="เช่น QoS, 42/69")

    edited_df = st.data_editor(
        display_df,
        hide_index=True,
        disabled=[
            columns["Operator"],
            columns["Price"],
            columns["ID"],
            columns["Expired date"],
            columns["Day(s) left"],
            columns["Status"],
            "sheet_row",
        ],
        column_config={
            "เลือกเบิก": st.column_config.CheckboxColumn("เลือกเบิก"),
            columns["Day(s) left"]: st.column_config.NumberColumn(columns["Day(s) left"]),
            "sheet_row": st.column_config.NumberColumn("แถวในชีต"),
        },
        width="stretch",
        key="withdraw_editor",
    )

    selected_rows = edited_df[edited_df["เลือกเบิก"] == True]
    if selected_rows.empty:
        st.warning("⚠️ กรุณาเลือกบัตรที่ต้องการเบิกจากตารางด้านบนอย่างน้อย 1 ใบ")
    else:
        selected_operators = ", ".join(selected_rows[columns["Operator"]].astype(str).unique())
        st.info(f"👉 ท่านกำลังเลือกบัตรค่าย **{selected_operators}** รวมทั้งหมด **{len(selected_rows)}** ใบ")

    if st.button("🚀 ยืนยันการเบิกบัตรที่เลือก", type="primary"):
        requester = str(requester or "").strip()
        job_no = str(job_no or "").strip()
        if not requester or not job_no:
            st.error("❌ ไม่สามารถเบิกได้: กรุณากรอกชื่อผู้ขอเบิกและ Job No. ให้ครบถ้วน")
            return
        if selected_rows.empty:
            st.error("❌ ไม่สามารถเบิกได้: ยังไม่ได้เลือกบัตร")
            return

        today_str = datetime.date.today().strftime("%d/%m/%Y")
        sheet_rows = [int(row["sheet_row"]) for _, row in selected_rows.iterrows()]
        updates = build_withdraw_updates(sheet_rows, today_str, requester, job_no)

        with st.spinner("กำลังทำการเบิกบัตรและอัปเดตข้อมูลลง Google Sheets..."):
            sheet.batch_update(updates, value_input_option="USER_ENTERED")

        st.success(f"ทำการเบิกบัตรเติมเงินจำนวน {len(selected_rows)} ใบให้คุณ {requester} เรียบร้อยแล้ว")
        st.rerun()


def render_nearest_expiry(df, columns):
    st.subheader("บัตรที่ใกล้หมดอายุที่สุดแยกตามเครือข่าย")

    if df.empty:
        st.info("ยังไม่มีข้อมูลบัตรในระบบ")
        return

    if not require_columns(columns, ["Status", "Operator", "Price", "Expired date"], "สรุปวันหมดอายุ"):
        return

    available_df = df[df[columns["Status"]] != "ใช้งานแล้ว"].copy()
    if available_df.empty:
        st.info("ยังไม่มีบัตรที่พร้อมใช้งานในระบบ")
        return

    available_df["Expired Date Parsed"] = parse_expired_dates(available_df[columns["Expired date"]])
    available_df[columns["Price"]] = pd.to_numeric(available_df[columns["Price"]], errors="coerce")
    available_df = available_df.dropna(subset=["Expired Date Parsed", columns["Price"]])
    available_df = available_df[available_df["Expired Date Parsed"].dt.date >= datetime.date.today()]

    if available_df.empty:
        st.warning("ยังไม่มีบัตรที่พร้อมใช้งานและอ่านวันหมดอายุได้")
        return

    nearest_expiry = available_df.groupby(columns["Operator"])["Expired Date Parsed"].transform("min")
    nearest_cards = available_df[available_df["Expired Date Parsed"] == nearest_expiry].copy()

    detail_summary = (
        nearest_cards.groupby([columns["Operator"], "Expired Date Parsed", columns["Price"]])
        .size()
        .reset_index(name="จำนวน (ใบ)")
        .sort_values([columns["Operator"], "Expired Date Parsed", columns["Price"]])
    )
    detail_summary["มูลค่ารวม"] = detail_summary[columns["Price"]] * detail_summary["จำนวน (ใบ)"]
    detail_summary["จำนวนวันที่เหลือ"] = (
        detail_summary["Expired Date Parsed"].dt.date - datetime.date.today()
    ).apply(lambda days: days.days)
    detail_summary["วันหมดอายุใกล้สุด"] = detail_summary["Expired Date Parsed"].dt.strftime("%d/%m/%y")

    operator_summary = (
        detail_summary.groupby([columns["Operator"], "วันหมดอายุใกล้สุด", "จำนวนวันที่เหลือ"]) 
        .agg(
            **{
                "จำนวนรวม (ใบ)": ("จำนวน (ใบ)", "sum"),
                "มูลค่ารวม": ("มูลค่ารวม", "sum"),
            }
        )
        .reset_index()
    )

    price_details = (
        detail_summary.assign(
            รายละเอียด=lambda data: data[columns["Price"]].astype(int).astype(str)
            + " บาท x "
            + data["จำนวน (ใบ)"].astype(str)
            + " ใบ"
        )
        .groupby(columns["Operator"])["รายละเอียด"]
        .apply(", ".join)
        .reset_index()
    )
    operator_summary = operator_summary.merge(price_details, on=columns["Operator"], how="left")

    st.dataframe(
        style_operator_rows(operator_summary, columns["Operator"]),
        width="stretch",
        hide_index=True,
    )
    st.caption("แจกแจงตามมูลค่าบัตรของวันหมดอายุที่ใกล้ที่สุด")
    st.dataframe(
        style_operator_rows(
            detail_summary[
                [
                    columns["Operator"],
                    "วันหมดอายุใกล้สุด",
                    "จำนวนวันที่เหลือ",
                    columns["Price"],
                    "จำนวน (ใบ)",
                    "มูลค่ารวม",
                ]
            ],
            columns["Operator"],
        ),
        width="stretch",
        hide_index=True,
    )


def render_withdrawal_report_tab(df, columns):
    st.subheader("📋 รายการคำขอสำรองเบิกใช้บัตรเติมเงิน ")

    if df.empty:
        st.info("💡 ยังไม่มีข้อมูลในระบบ หรือไม่สามารถโหลดข้อมูลจาก Google Sheets ได้")
        return

    required = ["Status", "Job No", "Withdraw Date"]
    if not require_columns(columns, required, "รายงานการเบิก"):
        st.warning("โปรดตรวจสอบว่าหัวตารางใน Google Sheets มีคอลัมน์: 'สถานะ', 'Job No.', และ 'วันที่ขอเบิก'")
        return

    job_column = columns["Job No"]
    user_column = columns.get("User")
    date_column = columns["Withdraw Date"]

    df_clean = df.copy()
    df_clean[columns["Status"]] = df_clean[columns["Status"]].astype(str).str.strip()
    df_clean[job_column] = df_clean[job_column].astype(str).str.strip()
    df_clean[date_column] = df_clean[date_column].astype(str).str.strip()

    withdrawn_df = df_clean[df_clean[columns["Status"]] == "ใช้งานแล้ว"]
    if withdrawn_df.empty:
        st.info("ℹ️ ปัจจุบันยังไม่มีสถานะ 'ใช้งานแล้ว' ในระบบ (ไม่มีประวัติการเบิก)")
        return

    unique_jobs = sorted([j for j in withdrawn_df[job_column].unique() if j and j != "nan" and j != "-"])
    if not unique_jobs:
        st.warning("⚠️ มีรายการเบิกใช้แล้ว แต่ไม่พบรหัส Job No. ที่สมบูรณ์ในระบบ")
        return

    filter_col1, filter_col2, filter_col3 = st.columns([1.2, 1, 1])

    if user_column and user_column in withdrawn_df.columns:
        all_requesters = sorted([
            r for r in withdrawn_df[user_column].dropna().astype(str).str.strip().unique()
            if r and r != "nan" and r != "-"
        ])
    else:
        all_requesters = []

    with filter_col1:
        if all_requesters:
            selected_requester = st.selectbox(
                "1. เลือกชื่อผู้ขอเบิก",
                [""] + all_requesters,
                format_func=lambda x: x or "เลือกชื่อผู้ขอเบิก",
            )
        else:
            selected_requester = st.text_input("1. ชื่อผู้ขอเบิก (พิมพ์)", placeholder="พิมพ์ชื่อผู้ขอเบิก")

    if selected_requester and user_column and user_column in withdrawn_df.columns:
        date_source = withdrawn_df[withdrawn_df[user_column].astype(str).str.contains(selected_requester, case=False, na=False)]
    else:
        date_source = withdrawn_df

    unique_dates = sorted([d for d in date_source[date_column].unique() if d and d != "nan" and d != "-"])
    if not unique_dates:
        st.warning("⚠️ ไม่พบวันที่ขอเบิกสำหรับเงื่อนไขที่เลือก")
        return

    with filter_col2:
        selected_date = st.selectbox(
            "2. เลือกวันที่ขอเบิก",
            [""] + unique_dates,
            format_func=lambda x: x or "เลือกวันที่ขอเบิก",
        )

    job_source = withdrawn_df
    if selected_requester and user_column and user_column in job_source.columns:
        job_source = job_source[job_source[user_column].astype(str).str.contains(selected_requester, case=False, na=False)]
    if selected_date:
        job_source = job_source[job_source[date_column] == selected_date]

    unique_jobs_for_filters = sorted([
        j for j in job_source[job_column].dropna().astype(str).str.strip().unique()
        if j and j != "nan" and j != "-"
    ])
    if not unique_jobs_for_filters:
        st.warning("⚠️ ไม่พบ Job No. สำหรับเงื่อนไขที่เลือก")
        return

    with filter_col3:
        selected_job = st.selectbox(
            "3. เลือก Job No.",
            [""] + unique_jobs_for_filters,
            format_func=lambda x: x or "เลือก Job No.",
        )

    requester_filled = bool(selected_requester) if isinstance(selected_requester, str) else True
    if (not requester_filled) or (not selected_date) or (not selected_job):
        st.info("กรุณาเลือก: ชื่อผู้ขอเบิก, วันที่ขอเบิก และ Job No. ก่อนดูสรุปยอดและรายละเอียด")
        return

    report_data = withdrawn_df.copy()
    if user_column and user_column in report_data.columns:
        report_data = report_data[report_data[user_column].astype(str).str.contains(selected_requester, case=False, na=False)]

    report_data = report_data[report_data[date_column] == selected_date]
    report_data = report_data[report_data[job_column] == selected_job]

    if report_data.empty:
        st.warning("⚠️ ไม่พบข้อมูลการเบิกที่ตรงกับเงื่อนไขที่เลือก")
        return

    def generate_withdrawal_statement_pdf(report_df, columns_map, requester, selected_date, selected_job, approver=""):
        if A4 is None:
            raise RuntimeError("reportlab library is not available. Install reportlab to enable PDF export.")

        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "TitleThai",
            parent=styles["Title"],
            fontName=THAI_PDF_FONT or styles["Title"].fontName,
            fontSize=14,
            leading=20,
            alignment=1,
        )
        body_style = ParagraphStyle(
            "BodyThai",
            parent=styles["Normal"],
            fontName=THAI_PDF_FONT or styles["Normal"].fontName,
            fontSize=10,
            leading=12,
        )
        label_style = ParagraphStyle(
            "LabelThai",
            parent=body_style,
            fontSize=10,
            leading=12,
        )
        elems = []

        title = Paragraph("<b>รายการคำขอสำรองเบิกใช้บัตรเติมเงินสำหรับงานทดสอบคุณภาพการให้บริการโทรศัพท์เคลื่อนที่</b>", title_style)
        elems.append(title)
        elems.append(Spacer(1, 14))

        header_data = [
            [Paragraph("<b>ชื่อผู้ขอเบิก</b>", label_style), requester, Paragraph("<b>Job No.</b>", label_style), selected_job],
            [Paragraph("<b>วันที่ขอเบิก</b>", label_style), selected_date, Paragraph("<b>ผู้อนุมัติ</b>", label_style), approver or ""],
        ]
        hdr_table = Table(header_data, colWidths=[70, 150, 70, 140], hAlign="CENTER")
        hdr_table.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("FONTNAME", (0, 0), (-1, -1), THAI_PDF_FONT or "Helvetica"),
            ("LINEBELOW", (1, 0), (1, 0), 0.5, colors.black),
            ("LINEBELOW", (3, 0), (3, 0), 0.5, colors.black),
            ("LINEBELOW", (1, 1), (1, 1), 0.5, colors.black),
            ("LINEBELOW", (3, 1), (3, 1), 0.5, colors.black),
        ]))
        elems.append(hdr_table)
        elems.append(Spacer(1, 16))

        cn_price = columns_map.get("Price") or columns_map.get("Value")
        cn_provider = columns_map.get("Operator")
        report_df = report_df.copy()
        if cn_price and cn_price in report_df.columns:
            report_df[cn_price] = pd.to_numeric(report_df[cn_price], errors="coerce").fillna(0)

        if cn_provider and cn_provider in report_df.columns and cn_price and cn_price in report_df.columns:
            grouped = report_df.groupby([cn_provider, cn_price]).size().reset_index(name="count")
            summary_data = []
            for _, row in grouped.iterrows():
                service = str(row[cn_provider])
                price = str(int(row[cn_price])) if pd.notna(row[cn_price]) else ""
                count = str(int(row["count"]))
                summary_data.append([service, "บัตรใบละ", price, "จำนวน", count, "ใบ"])

            if summary_data:
                summary_tbl = Table(summary_data, colWidths=[90, 55, 50, 50, 45, 30], hAlign="CENTER")
                summary_tbl.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, -1), THAI_PDF_FONT or "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 10),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("ALIGN", (2, 0), (2, -1), "CENTER"),
                    ("ALIGN", (4, 0), (4, -1), "CENTER"),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ]))
                elems.append(summary_tbl)
                elems.append(Spacer(1, 14))

        cn_card_no = columns_map.get("ID") or columns_map.get("Card No") or columns_map.get("Card Number")
        cn_provider = columns_map.get("Operator")
        cn_user = columns_map.get("User")

        table_data = [["ลำดับ", "เลขที่บัตร", "ผู้ให้บริการ", "ราคาบัตร", "ผู้ใช้บัตร", "หมายเหตุ"]]
        for i, row in enumerate(report_df.itertuples(index=False), start=1):
            card_no = getattr(row, cn_card_no) if cn_card_no in report_df.columns else ""
            provider = getattr(row, cn_provider) if cn_provider in report_df.columns else ""
            price = getattr(row, cn_price) if cn_price in report_df.columns else ""
            user = getattr(row, cn_user) if (cn_user and cn_user in report_df.columns) else ""
            table_data.append([str(i), str(card_no), str(provider), str(price), str(user), ""])

        tbl = Table(table_data, colWidths=[40, 100, 80, 80, 80, 100])
        tbl.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("FONTNAME", (0, 0), (-1, -1), THAI_PDF_FONT or "Helvetica"),
        ]))
        elems.append(tbl)

        total_count = len(report_df)
        total_value = int(report_df[cn_price].sum()) if cn_price and cn_price in report_df.columns else 0
        elems.append(Spacer(1, 12))
        total_summary = Table(
            [[Paragraph("<b>รวมทั้งหมด</b>", label_style), f"{total_count} ใบ", f"มูลค่ารวม {total_value:,} บาท"]],
            colWidths=[120, 80, 140],
            hAlign="RIGHT",
        )
        total_summary.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), THAI_PDF_FONT or "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
        ]))
        elems.append(total_summary)

        doc.build(elems)
        buf.seek(0)
        return buf.read()

    try:
        pdf_bytes = generate_withdrawal_statement_pdf(
            report_data,
            columns,
            selected_requester if isinstance(selected_requester, str) else "",
            selected_date,
            selected_job,
        )
        safe_job = re.sub(r"[\\/:*?\"<>|]+", "_", str(selected_job))
        safe_date = re.sub(r"[\\/:*?\"<>|]+", "_", str(selected_date))
        st.download_button(
            "ดาวน์โหลด PDF",
            data=pdf_bytes,
            file_name=f"รายการคำขอเบิก_{safe_job}_{safe_date}.pdf",
            mime="application/pdf",
        )
    except RuntimeError as e:
        st.info(str(e))
        try:
            csv_bytes = report_data.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            safe_job = re.sub(r"[\\/:*?\"<>|]+", "_", str(selected_job))
            safe_date = re.sub(r"[\\/:*?\"<>|]+", "_", str(selected_date))
            st.download_button(
                "ดาวน์โหลด CSV (สำรอง)",
                data=csv_bytes,
                file_name=f"statement_{safe_job}_{safe_date}.csv",
                mime="text/csv",
            )
            st.info("หรือ ติดตั้ง reportlab (`pip install reportlab`) เพื่อดาวน์โหลดเป็น PDF")
        except Exception as e2:
            st.error(f"ไม่สามารถสร้างไฟล์สำรองได้: {e2}")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดขณะสร้าง PDF: {e}")

    requester_name = report_data[user_column].iloc[0] if user_column and user_column in report_data.columns else "ไม่ระบุชื่อผู้ใช้บัตร"

    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    c1.markdown(f"👤 **ชื่อผู้ขอเบิก:** `{requester_name}`")
    c2.markdown(f"💼 **Job No:** `{selected_job}`")
    c3.markdown(f"📅 **วันที่ขอเบิก:** `{selected_date}`")
    st.markdown("---")
    st.markdown("##### 📊 **สรุปยอดรวมการเบิก (แยกตามราคา)**")
    report_data[columns["Price"]] = pd.to_numeric(report_data[columns["Price"]], errors="coerce").fillna(0)

    price_summary = report_data.groupby(columns["Price"]).size().reset_index(name="จำนวน (ใบ)")
    price_summary["มูลค่ารวม (บาท)"] = price_summary[columns["Price"]] * price_summary["จำนวน (ใบ)"]

    st.dataframe(price_summary, hide_index=True, width="stretch")

    st.markdown("##### 🎫 **รายละเอียดหมายเลขบัตรที่เบิก**")
    report_data = report_data.reset_index(drop=True)
    report_data.insert(0, "ลำดับ", report_data.index + 1)

    display_cols = ["ลำดับ", columns["ID"], columns["Operator"], columns["Price"], user_column]
    existing_cols = [col for col in display_cols if col in report_data.columns]

    st.dataframe(report_data[existing_cols], hide_index=True, width="stretch")


def render_dashboard_header():
    st.markdown(
        '<div class="top-label">ฐานข้อมูล StockV3</div>',
        unsafe_allow_html=True,
    )


def render_raw_table(df, columns):
    if df.empty:
        st.info("ยังไม่มีข้อมูลบัตรในระบบ")
        return

    display_df = df.drop(columns=["sheet_row"], errors="ignore").copy()
    if columns.get("Used?") and columns.get("Status"):
        display_df[columns["Used?"]] = display_df[columns["Status"]].apply(format_used_display)

    if columns.get("Operator"):
        op_filter = st.multiselect("กรองเครือข่าย", options=OPERATORS, default=OPERATORS)
        if op_filter:
            selected_ops = [op.upper() for op in op_filter]
            display_df = display_df[display_df[columns["Operator"]].astype(str).str.upper().isin(selected_ops)]

        st.dataframe(
            style_operator_rows(display_df, columns["Operator"]),
            width="stretch",
            hide_index=True,
        )
    else:
        st.dataframe(display_df, width="stretch", hide_index=True)


st.set_page_config(page_title="ระบบบริหารจัดการบัตรเติมเงิน", layout="wide")
inject_styles()

client = authorize_google_sheets()
sheet = open_stock_sheet(client)

df = load_sheet_dataframe(sheet)
if not df.empty:
    df.columns = df.columns.str.strip()
columns = build_column_map(df)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.username = ""
    st.session_state.role = ""

render_dashboard_header()

TEST_USERS = {
    "admin": {"password": "Qosnbtc0", "role": "Admin"},
    "tester": {"password": "userpass", "role": "User"},
}


def _do_logout():
    st.session_state.authenticated = False
    st.session_state.username = ""
    st.session_state.role = ""
    safe_rerun()


st.sidebar.header("เข้าสู่ระบบ")
if not st.session_state.authenticated:
    with st.sidebar.form("login_form"):
        u = st.text_input("ชื่อผู้ใช้", key="login_username")
        p = st.text_input("รหัสผ่าน", type="password", key="login_password")
        login_submitted = st.form_submit_button("ล็อกอิน")
    if login_submitted:
        user = TEST_USERS.get(u)
        if user and p == user["password"]:
            st.session_state.authenticated = True
            st.session_state.username = u
            st.session_state.role = user["role"]
            st.sidebar.success(f"ล็อกอินสำเร็จ: {st.session_state.role}")
            safe_rerun()
        else:
            st.sidebar.error("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
else:
    st.sidebar.markdown(f"**ผู้ใช้:** {st.session_state.username}")
    st.sidebar.markdown(f"**บทบาท:** {st.session_state.role}")
    if st.sidebar.button("ออกจากระบบ"):
        _do_logout()


overview_tab, add_tab, withdraw_tab, expiry_tab, data_tab, report_tab = st.tabs(
    ["ภาพรวม", "เพิ่มบัตร", "เบิกบัตร", "ใกล้หมดอายุ", "ข้อมูลทั้งหมด", "รายงานการเบิก"]
)

with overview_tab:
    render_stock_summary(df, columns)

with add_tab:
    if not st.session_state.get("authenticated", False):
        st.error("❌ กรุณาเข้าสู่ระบบเพื่อใช้ฟีเจอร์นี้")
    elif st.session_state.get("role") != "Admin":
        st.error("❌ คุณไม่มีสิทธิ์เพิ่มบัตร — ติดต่อแอดมิน")
    else:
        render_add_card_form(sheet)

with withdraw_tab:
    if not st.session_state.get("authenticated", False):
        st.error("❌ กรุณาเข้าสู่ระบบเพื่อใช้ฟีเจอร์นี้")
    elif st.session_state.get("role") != "Admin":
        st.error("❌ คุณไม่มีสิทธิ์เบิกบัตร — ติดต่อแอดมิน")
    else:
        render_withdraw(df, sheet, columns)

with expiry_tab:
    render_nearest_expiry(df, columns)

with data_tab:
    render_raw_table(df, columns)

with report_tab:
    render_withdrawal_report_tab(df, columns)
