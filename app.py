"""
Cement Distribution – Outstanding Tracker
-----------------------------------------
Reads the CEMENT.xlsx workbook (Loading / Unloading / Payment details),
calculates the outstanding per customer and shows a small dashboard.
Also sends an SMS to each customer about their outstanding.

Run:   streamlit run app.py
"""

import json
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

APP_DIR = Path(__file__).parent
CONFIG_FILE = APP_DIR / "config.json"
SMS_LOG_FILE = APP_DIR / "sms_log.csv"

DEFAULT_TEMPLATE = (
    "Customer Name : {customer}\n"
    "Outstanding: Rs. {outstanding}\n"
    "Unloaded Qty: {qty}"
)

st.set_page_config(page_title="YM Cement Distribution", page_icon="🏗️", layout="wide")


# --------------------------------------------------------------------------
# Config (file path is saved so the user only configures it once)
# --------------------------------------------------------------------------
def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_config(cfg: dict) -> None:
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


# --------------------------------------------------------------------------
# Excel parsing
# --------------------------------------------------------------------------
def _norm(v) -> str:
    return str(v).strip().lower() if pd.notna(v) else ""


def _key(name) -> str:
    """Normalised customer key (so 'CENTRAL ' matches 'Central')."""
    return re.sub(r"\s+", " ", str(name)).strip().upper()


def normalize_lk_phone(raw) -> str | None:
    """Return a number in 94XXXXXXXXX format, or None if it looks invalid."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    digits = re.sub(r"\D", "", str(raw))
    if digits.endswith(".0"):
        digits = digits[:-2]
    if digits.startswith("0094"):
        digits = digits[2:]
    if digits.startswith("94") and len(digits) == 11:
        return digits
    if digits.startswith("0") and len(digits) == 10:
        return "94" + digits[1:]
    if len(digits) == 9:  # 7XXXXXXXX
        return "94" + digits
    return None


@st.cache_data(show_spinner=False)
def read_workbook(path: str, mtime: float, main_sheet: str, phone_sheet: str):
    """
    Parse the workbook. `mtime` is only there so the cache refreshes when the
    file is saved. Returns (unloading_df, loading_df, payments_df, phones_df).
    """
    raw = pd.read_excel(path, sheet_name=main_sheet, header=None, engine="openpyxl")

    # header row = the row that contains the 'Customer' cell
    hdr = None
    for i, row in raw.iterrows():
        if any(_norm(v) == "customer" for v in row):
            hdr = i
            break
    if hdr is None:
        raise ValueError("Could not find a 'Customer' column header in the sheet.")

    headers = [_norm(v) for v in raw.iloc[hdr].tolist()]
    body = raw.iloc[hdr + 1:].reset_index(drop=True)

    c = headers.index("customer")  # unloading block: Customer, Qty, Rate, Amount

    # ---- loading block (columns left of Customer) -------------------------
    def find_left(name):
        for idx in range(0, c):
            if headers[idx] == name:
                return idx
        return None

    L_date, L_po, L_inv = find_left("date"), find_left("p.o"), find_left("invoice")
    L_qty, L_rate = find_left("qty"), find_left("rate")

    # ---- payment columns (every 'payment' column right of Customer) -------
    pay_cols = [i for i in range(c + 1, len(headers)) if "payment" in headers[i]]

    rows_unload, rows_load, rows_pay = [], [], []
    cur_po, cur_inv = None, None

    for _, r in body.iterrows():
        # carry PO / invoice down (they are only typed on the first row)
        if L_po is not None and pd.notna(r.iloc[L_po]):
            cur_po = r.iloc[L_po]
        if L_inv is not None and pd.notna(r.iloc[L_inv]):
            cur_inv = r.iloc[L_inv]

        # loading rows
        if L_qty is not None and pd.notna(r.iloc[L_qty]):
            rows_load.append(
                {
                    "Date": r.iloc[L_date] if L_date is not None else None,
                    "PO": cur_po,
                    "Invoice": cur_inv,
                    "Loaded Qty": float(r.iloc[L_qty]),
                    "Rate": float(r.iloc[L_rate]) if L_rate is not None and pd.notna(r.iloc[L_rate]) else 0.0,
                }
            )

        # unloading rows
        cust = r.iloc[c]
        if pd.isna(cust) or str(cust).strip() == "":
            continue
        qty = pd.to_numeric(r.iloc[c + 1], errors="coerce")
        rate = pd.to_numeric(r.iloc[c + 2], errors="coerce")
        qty = 0.0 if pd.isna(qty) else float(qty)
        rate = 0.0 if pd.isna(rate) else float(rate)

        paid_total = 0.0
        for pc in pay_cols:
            amt = pd.to_numeric(r.iloc[pc], errors="coerce")
            if pd.notna(amt) and amt != 0:
                paid_total += float(amt)
                # date column sits immediately to the left of each payment column
                pdate = r.iloc[pc - 1] if pc - 1 > c and "date" in headers[pc - 1] else None
                rows_pay.append(
                    {
                        "Customer": str(cust).strip(),
                        "Key": _key(cust),
                        "Date": pdate,
                        "Payment": float(amt),
                    }
                )

        rows_unload.append(
            {
                "Customer": str(cust).strip(),
                "Key": _key(cust),
                "PO": cur_po,
                "Invoice": cur_inv,
                "Unloaded Qty": qty,
                "Rate": rate,
                "Amount": qty * rate,  # rate * qty = amount to pay
                "Paid": paid_total,
            }
        )

    unload = pd.DataFrame(rows_unload)
    loading = pd.DataFrame(rows_load)
    pay = pd.DataFrame(rows_pay, columns=["Customer", "Key", "Date", "Payment"])

    # ---- phone sheet -------------------------------------------------------
    phones = pd.DataFrame(columns=["Key", "Phone"])
    try:
        ph = pd.read_excel(path, sheet_name=phone_sheet, engine="openpyxl", dtype=str)
        ph.columns = [str(x).strip().lower() for x in ph.columns]
        ccol = next((x for x in ph.columns if "customer" in x), None)
        pcol = next((x for x in ph.columns if "phone" in x or "mobile" in x or "tel" in x), None)
        if ccol and pcol:
            ph = ph[[ccol, pcol]].dropna(subset=[ccol])
            phones = pd.DataFrame({"Key": ph[ccol].map(_key), "Phone": ph[pcol]})
    except Exception:
        pass

    return unload, loading, pay, phones


def summarise(unload: pd.DataFrame) -> pd.DataFrame:
    g = (
        unload.groupby("Key", sort=False)
        .agg(
            Customer=("Customer", "first"),
            **{
                "Unloaded Qty": ("Unloaded Qty", "sum"),
                "Amount to Pay": ("Amount", "sum"),
                "Paid": ("Paid", "sum"),
            },
        )
        .reset_index()
    )
    g["Outstanding"] = g["Amount to Pay"] - g["Paid"]
    g["Collection %"] = (g["Paid"] / g["Amount to Pay"].where(g["Amount to Pay"] != 0) * 100).fillna(0).round(1)
    g["Avg Rate"] = (g["Amount to Pay"] / g["Unloaded Qty"].where(g["Unloaded Qty"] != 0)).fillna(0)
    return g


def save_phones_to_excel(path: str, sheet: str, df: pd.DataFrame) -> None:
    """Write the Customer / Phone table back into the phone sheet (formulas elsewhere are kept)."""
    from openpyxl import load_workbook

    wb = load_workbook(path)  # not data_only -> formulas preserved
    ws = wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
    ws.delete_rows(1, ws.max_row)
    ws.append(["Customer", "Phone Number"])
    for _, r in df.iterrows():
        ws.append([r["Customer"], r["Phone"] or ""])
    wb.save(path)


# --------------------------------------------------------------------------
# SMS providers
# --------------------------------------------------------------------------
def send_notify_lk(user_id, api_key, sender_id, to, message):
    r = requests.post(
        "https://app.notify.lk/api/v1/send",
        data={"user_id": user_id, "api_key": api_key, "sender_id": sender_id, "to": to, "message": message},
        timeout=30,
    )
    try:
        j = r.json()
    except Exception:
        return False, f"HTTP {r.status_code}: {r.text[:120]}"
    ok = str(j.get("status", "")).lower() == "success"
    return ok, "sent" if ok else str(j.get("message") or j)[:150]


def send_twilio(sid, token, from_no, to, message):
    r = requests.post(
        f"https://api.twilio.com/2010-04-01/Accounts/{sid}/Messages.json",
        auth=(sid, token),
        data={"To": "+" + to, "From": from_no, "Body": message},
        timeout=30,
    )
    ok = r.status_code in (200, 201)
    return ok, "sent" if ok else r.text[:150]


def secret(name: str, default: str = "") -> str:
    """Credential lookup: environment variable first, then .streamlit/secrets.toml."""
    if os.environ.get(name):
        return os.environ[name]
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


# --------------------------------------------------------------------------
# Sidebar – configuration
# --------------------------------------------------------------------------
cfg = load_config()
st.sidebar.title("⚙️ Configuration")

file_path = st.sidebar.text_input(
    "Excel file path",
    value=cfg.get("file_path", ""),
    placeholder=r"C:\Users\you\Documents\CEMENT.xlsx",
    help="Full path to the CEMENT.xlsx file. It is saved, so you only enter it once.",
).strip().strip('"')

main_sheet = cfg.get("main_sheet", "Sheet1")
phone_sheet = cfg.get("phone_sheet", "Sheet2")

sheet_names = []
if file_path and Path(file_path).is_file():
    try:
        sheet_names = pd.ExcelFile(file_path, engine="openpyxl").sheet_names
    except Exception as e:
        st.sidebar.error(f"Cannot open file: {e}")

if sheet_names:
    main_sheet = st.sidebar.selectbox(
        "Main sheet (Loading / Unloading / Payments)",
        sheet_names,
        index=sheet_names.index(main_sheet) if main_sheet in sheet_names else 0,
    )
    phone_sheet = st.sidebar.selectbox(
        "Phone number sheet",
        sheet_names,
        index=sheet_names.index(phone_sheet) if phone_sheet in sheet_names else min(1, len(sheet_names) - 1),
    )

if st.sidebar.button("💾 Save configuration", width="stretch"):
    save_config({"file_path": file_path, "main_sheet": main_sheet, "phone_sheet": phone_sheet})
    st.sidebar.success("Saved")

if st.sidebar.button("🔄 Reload data", width="stretch"):
    st.cache_data.clear()
    st.session_state.pop("phone_table", None)
    st.rerun()

st.title("🏗️ YM CEMENT DISTRIBUTION – Outstanding Tracker")

if not file_path:
    st.info("Enter the **Excel file path** in the sidebar to begin.")
    st.stop()
if not Path(file_path).is_file():
    st.error(f"File not found: `{file_path}`")
    st.stop()

try:
    unload, loading, pay, phones = read_workbook(
        file_path, os.path.getmtime(file_path), main_sheet, phone_sheet
    )
except Exception as e:
    st.error(f"Could not read the workbook: {e}")
    st.stop()

if unload.empty:
    st.warning("No unloading rows found in the sheet.")
    st.stop()

summary = summarise(unload)

# ---- phone table (editable, kept in session) ------------------------------
phone_lookup = dict(zip(phones["Key"], phones["Phone"]))
if "phone_table" not in st.session_state or set(st.session_state.phone_table["Key"]) != set(summary["Key"]):
    st.session_state.phone_table = pd.DataFrame(
        {
            "Key": summary["Key"],
            "Customer": summary["Customer"],
            "Phone": [phone_lookup.get(k, "") or "" for k in summary["Key"]],
        }
    )

phone_map = {
    r["Key"]: normalize_lk_phone(r["Phone"]) for _, r in st.session_state.phone_table.iterrows()
}
summary["Phone"] = summary["Key"].map(phone_map)

rs = lambda v: f"Rs. {v:,.2f}"

tab_dash, tab_data, tab_sms = st.tabs(["📊 Dashboard", "📋 Details", "📱 SMS"])

# ==========================================================================
# DASHBOARD
# ==========================================================================
with tab_dash:
    total_loaded = loading["Loaded Qty"].sum() if not loading.empty else 0
    total_unloaded = summary["Unloaded Qty"].sum()
    total_amount = summary["Amount to Pay"].sum()
    total_paid = summary["Paid"].sum()
    total_out = summary["Outstanding"].sum()
    coll = (total_paid / total_amount * 100) if total_amount else 0

    k = st.columns(5)
    k[0].metric("Loaded (bags)", f"{total_loaded:,.0f}")
    k[1].metric("Unloaded (bags)", f"{total_unloaded:,.0f}", delta=f"{total_unloaded - total_loaded:,.0f} vs loaded" if total_loaded else None, delta_color="off")
    k[2].metric("Total to collect", rs(total_amount))
    k[3].metric("Total paid", rs(total_paid), delta=f"{coll:.1f}% collected", delta_color="off")
    k[4].metric("Outstanding", rs(total_out))

    if total_out <= 0.005:
        st.success("All customers are fully paid – no outstanding balance.")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            summary.sort_values("Unloaded Qty", ascending=False),
            x="Customer", y="Unloaded Qty", text="Unloaded Qty",
            title="Bags unloaded per customer",
        )
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig.update_layout(yaxis_title="Bags", xaxis_title=None, margin=dict(t=50, b=10))
        st.plotly_chart(fig, width="stretch")
    with c2:
        fig = px.pie(summary, names="Customer", values="Unloaded Qty", hole=0.45, title="Share of unloaded bags")
        fig.update_layout(margin=dict(t=50, b=10))
        st.plotly_chart(fig, width="stretch")

    c3, c4 = st.columns(2)
    with c3:
        s = summary.sort_values("Amount to Pay", ascending=False)
        fig = go.Figure()
        fig.add_bar(x=s["Customer"], y=s["Amount to Pay"], name="Amount to pay")
        fig.add_bar(x=s["Customer"], y=s["Paid"], name="Paid")
        fig.add_bar(x=s["Customer"], y=s["Outstanding"], name="Outstanding")
        fig.update_layout(barmode="group", title="Amount to pay vs paid vs outstanding (Rs.)",
                          yaxis_title="LKR", margin=dict(t=50, b=10), legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig, width="stretch")
    with c4:
        s = summary.sort_values("Collection %")
        fig = px.bar(s, x="Collection %", y="Customer", orientation="h", text="Collection %",
                     range_x=[0, 105], title="Payment performance – % of bill paid")
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(margin=dict(t=50, b=10), yaxis_title=None)
        st.plotly_chart(fig, width="stretch")

    if len(pay) and pay["Date"].notna().any():
        st.caption("Payments received over time")
        p = pay.dropna(subset=["Date"]).copy()
        p["Date"] = pd.to_datetime(p["Date"], errors="coerce")
        p = p.dropna(subset=["Date"]).groupby("Date", as_index=False)["Payment"].sum().sort_values("Date")
        fig = px.bar(p, x="Date", y="Payment", title="Payments received by date (Rs.)")
        fig.update_layout(margin=dict(t=50, b=10))
        st.plotly_chart(fig, width="stretch")

# ==========================================================================
# DETAILS
# ==========================================================================
with tab_data:
    st.subheader("Customer summary")
    show = summary[["Customer", "Unloaded Qty", "Avg Rate", "Amount to Pay", "Paid", "Outstanding", "Collection %", "Phone"]]
    st.dataframe(
        show,
        hide_index=True,
        width="stretch",
        column_config={
            "Unloaded Qty": st.column_config.NumberColumn(format="%d"),
            "Avg Rate": st.column_config.NumberColumn(format="Rs. %.2f"),
            "Amount to Pay": st.column_config.NumberColumn(format="Rs. %.2f"),
            "Paid": st.column_config.NumberColumn(format="Rs. %.2f"),
            "Outstanding": st.column_config.NumberColumn(format="Rs. %.2f"),
            "Collection %": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100),
        },
    )

    st.subheader("Loading vs unloading (by PO / invoice)")
    if not loading.empty:
        by_inv = unload.groupby(["PO", "Invoice"], dropna=False)["Unloaded Qty"].sum().reset_index()
        lv = loading.merge(by_inv, on=["PO", "Invoice"], how="left")
        lv["Unloaded Qty"] = lv["Unloaded Qty"].fillna(0)
        lv["Still on vehicle"] = lv["Loaded Qty"] - lv["Unloaded Qty"]
        lv["Load Amount"] = lv["Loaded Qty"] * lv["Rate"]
        st.dataframe(lv, hide_index=True, width="stretch")
        if (lv["Still on vehicle"].abs() > 0.001).any():
            st.warning("Some loads have bags that are not yet unloaded (or unloaded more than loaded).")
    else:
        st.info("No loading rows found.")

    st.subheader("Payment history")
    if len(pay):
        st.dataframe(pay.drop(columns="Key"), hide_index=True, width="stretch")
        st.caption("Note: check payment dates in Excel – some look like day/month were swapped (e.g. 2025-03-11 instead of 2025-11-03).")
    else:
        st.info("No payments recorded.")

# ==========================================================================
# SMS
# ==========================================================================
with tab_sms:
    st.subheader("1. Customer phone numbers")
    st.caption("Numbers are read from the phone sheet. Fill in or correct them here (07X XXX XXXX or 94X format).")
    edited = st.data_editor(
        st.session_state.phone_table[["Customer", "Phone"]],
        hide_index=True,
        width="stretch",
        disabled=["Customer"],
        key="phone_editor",
    )
    st.session_state.phone_table["Phone"] = edited["Phone"].fillna("").values
    phone_map = {r["Key"]: normalize_lk_phone(r["Phone"]) for _, r in st.session_state.phone_table.iterrows()}
    summary["Phone"] = summary["Key"].map(phone_map)

    if st.button("💾 Save these numbers into the Excel phone sheet"):
        try:
            save_phones_to_excel(file_path, phone_sheet, st.session_state.phone_table[["Customer", "Phone"]])
            st.cache_data.clear()
            st.success("Phone numbers saved. (Make sure the file is not open in Excel when saving.)")
        except PermissionError:
            st.error("The file is open in Excel – close it and try again.")
        except Exception as e:
            st.error(f"Could not save: {e}")

    st.subheader("2. Message")
    template = st.text_area(
        "SMS template",
        value=cfg.get("template", DEFAULT_TEMPLATE),
        height=110,
        help="Placeholders: {customer}  {outstanding}  {qty}",
    )
    include_zero = st.checkbox("Also message customers with zero outstanding", value=False)

    targets = summary if include_zero else summary[summary["Outstanding"] > 0.005]
    msgs = []
    for _, r in targets.iterrows():
        try:
            text = template.format(
                customer=r["Customer"],
                outstanding=f"{r['Outstanding']:,.2f}",
                qty=f"{r['Unloaded Qty']:,.0f}",
            )
        except (KeyError, IndexError, ValueError) as e:
            st.error(f"Template error: {e}")
            st.stop()
        msgs.append({"Customer": r["Customer"], "Phone": r["Phone"], "Message": text})
    preview = pd.DataFrame(msgs)

    if preview.empty:
        st.success("Nobody has an outstanding balance, so there is nothing to send.")
    else:
        st.dataframe(preview, hide_index=True, width="stretch")
        missing = preview["Phone"].isna().sum()
        if missing:
            st.warning(f"{missing} customer(s) have a missing / invalid phone number and will be skipped.")

    st.subheader("3. Send")
    provider = st.selectbox(
        "SMS provider",
        ["Dry run (preview only – nothing is sent)", "Notify.lk", "Twilio"],
    )

    creds = {}
    if provider == "Notify.lk":
        creds["user_id"] = st.text_input("Notify.lk User ID", value=secret("NOTIFYLK_USER_ID"))
        creds["api_key"] = st.text_input("Notify.lk API Key", value=secret("NOTIFYLK_API_KEY"), type="password")
        creds["sender_id"] = st.text_input("Sender ID", value=secret("NOTIFYLK_SENDER_ID", "NotifyDEMO"))
    elif provider == "Twilio":
        creds["sid"] = st.text_input("Account SID", value=secret("TWILIO_ACCOUNT_SID"))
        creds["token"] = st.text_input("Auth Token", value=secret("TWILIO_AUTH_TOKEN"), type="password")
        creds["from"] = st.text_input("From number (e.g. +1555...)", value=secret("TWILIO_FROM"))

    live = provider != "Dry run (preview only – nothing is sent)"
    confirm = True
    if live:
        confirm = st.checkbox(f"I confirm – send {int(preview['Phone'].notna().sum()) if not preview.empty else 0} real SMS now")

    if st.button("📤 Send SMS to customers", type="primary", disabled=preview.empty or (live and not confirm)):
        results = []
        bar = st.progress(0.0)
        for i, m in enumerate(msgs):
            if not m["Phone"]:
                status, ok = "skipped – no valid phone", False
            elif not live:
                status, ok = "dry run (not sent)", None
            else:
                try:
                    if provider == "Notify.lk":
                        ok, status = send_notify_lk(creds["user_id"], creds["api_key"], creds["sender_id"], m["Phone"], m["Message"])
                    else:
                        ok, status = send_twilio(creds["sid"], creds["token"], creds["from"], m["Phone"], m["Message"])
                except Exception as e:
                    ok, status = False, f"error: {e}"
            results.append({"Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Customer": m["Customer"],
                            "Phone": m["Phone"], "Provider": provider, "Result": status, "Message": m["Message"]})
            bar.progress((i + 1) / len(msgs))

        res = pd.DataFrame(results)
        st.dataframe(res.drop(columns="Message"), hide_index=True, width="stretch")
        res.to_csv(SMS_LOG_FILE, mode="a", header=not SMS_LOG_FILE.exists(), index=False)
        if live:
            sent = sum(1 for r in results if r["Result"] == "sent")
            st.success(f"{sent} of {len(results)} SMS sent. Log saved to {SMS_LOG_FILE.name}")

    # remember the template
    if template != cfg.get("template", DEFAULT_TEMPLATE):
        if st.button("Save this template as default"):
            cfg["template"] = template
            cfg.update({"file_path": file_path, "main_sheet": main_sheet, "phone_sheet": phone_sheet})
            save_config(cfg)
            st.success("Template saved")
