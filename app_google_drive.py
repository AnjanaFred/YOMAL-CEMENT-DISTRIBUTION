"""
Cement Distribution – Outstanding Tracker
-----------------------------------------
Uploads and reads the CEMENT.xlsx workbook (Loading / Unloading / Payment details),
calculates the outstanding per customer, records new payments, and shows a dashboard.
Also sends an SMS to each customer about their outstanding.

Run:   streamlit run app.py
"""

import json
import os
import re
import tempfile
import hashlib
import io
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

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
                "Date": r.iloc[L_date] if L_date is not None else None,
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

    # ---- separately entered payment ledger (created by the Payment tab) ----
    ledger_pay = pd.DataFrame(columns=["Customer", "Key", "Date", "Payment", "Reference"])
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        if "Payment Entries" in xls.sheet_names:
            ledger = pd.read_excel(path, sheet_name="Payment Entries", engine="openpyxl")
            if not ledger.empty:
                ledger.columns = [str(x).strip() for x in ledger.columns]
                required = {"Date", "Customer", "Payment Amount"}
                if required.issubset(set(ledger.columns)):
                    ledger_pay = pd.DataFrame({
                        "Customer": ledger["Customer"].astype(str).str.strip(),
                        "Key": ledger["Customer"].map(_key),
                        "Date": pd.to_datetime(ledger["Date"], errors="coerce"),
                        "Payment": pd.to_numeric(ledger["Payment Amount"], errors="coerce").fillna(0.0),
                        "Reference": ledger["Reference"].fillna("").astype(str) if "Reference" in ledger.columns else "",
                    })
                    ledger_pay = ledger_pay[ledger_pay["Payment"] > 0].copy()
                    pay = pd.concat(
                        [pay.assign(Reference="Existing workbook payment"), ledger_pay],
                        ignore_index=True,
                    )
    except Exception:
        # A missing or unreadable payment ledger should not prevent the main
        # workbook from loading; the Payment tab will create it when needed.
        pass

    return unload, loading, pay, phones, ledger_pay


def summarise(unload: pd.DataFrame, ledger_pay: pd.DataFrame | None = None) -> pd.DataFrame:
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
    # Payments entered through the Payment tab are stored separately from
    # the legacy payment columns in the original workbook, so add them once.
    if ledger_pay is not None and not ledger_pay.empty:
        extra_paid = ledger_pay.groupby("Key")["Payment"].sum()
        g["Paid"] = g["Paid"] + g["Key"].map(extra_paid).fillna(0.0)

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


def append_payment_to_excel(path: str, payment_date, customer: str, amount: float, reference: str = "") -> None:
    """Append a payment to the dedicated Payment Entries worksheet."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    sheet_name = "Payment Entries"
    headers = ["Date", "Customer", "Payment Amount", "Reference", "Recorded At"]
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row == 1 and all(ws.cell(1, c).value is None for c in range(1, 6)):
            for col, value in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=value)
        elif ws.cell(1, 1).value != "Date":
            ws.insert_rows(1)
            for col, value in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=value)
    else:
        ws = wb.create_sheet(sheet_name)
        for col, value in enumerate(headers, start=1):
            ws.cell(row=1, column=col, value=value)

    ws.append([
        payment_date,
        customer,
        float(amount),
        reference,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ])
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


def google_drive_settings():
    """Read Google Drive settings from Streamlit secrets; never hardcode credentials."""
    try:
        cfg = st.secrets.get("gdrive", {})
        raw_json = cfg.get("service_account_json", "")
        file_id = cfg.get("workbook_file_id", "")
        if not raw_json or not file_id:
            return None
        info = json.loads(raw_json) if isinstance(raw_json, str) else dict(raw_json)
        return info, str(file_id).strip()
    except Exception:
        return None


def google_drive_service():
    settings = google_drive_settings()
    if not settings:
        return None
    info, _ = settings
    credentials = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def download_workbook_from_drive(local_path: str) -> str:
    """Download the configured workbook from Google Drive to a temporary local copy."""
    settings = google_drive_settings()
    if not settings:
        raise RuntimeError("Google Drive is not configured. Add [gdrive] settings to Streamlit secrets.")
    _, file_id = settings
    service = google_drive_service()
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    with open(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return local_path


def upload_workbook_to_drive(local_path: str) -> None:
    """Replace the configured Google Drive workbook with the edited local workbook."""
    settings = google_drive_settings()
    if not settings:
        raise RuntimeError("Google Drive is not configured.")
    _, file_id = settings
    service = google_drive_service()
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    service.files().update(
        fileId=file_id,
        media_body=media,
        supportsAllDrives=True,
    ).execute()


def sync_workbook(path: str) -> None:
    """Persist workbook changes to Google Drive when cloud mode is enabled."""
    if google_drive_settings():
        upload_workbook_to_drive(path)


def read_optional_sheet(path: str, sheet_name: str) -> pd.DataFrame:
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        if sheet_name in xls.sheet_names:
            return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        pass
    return pd.DataFrame()


def load_customers(path: str, unload: pd.DataFrame) -> pd.DataFrame:
    customers = read_optional_sheet(path, "Customers")
    if customers.empty or "Customer" not in customers.columns:
        customers = pd.DataFrame(columns=["Customer", "Status", "Phone"])
    customers.columns = [str(c).strip() for c in customers.columns]
    if "Status" not in customers.columns:
        customers["Status"] = "Active"
    if "Phone" not in customers.columns:
        customers["Phone"] = ""
    customers["Customer"] = customers["Customer"].astype(str).str.strip()
    customers["Key"] = customers["Customer"].map(_key)
    customers["Status"] = customers["Status"].fillna("Active").astype(str).str.title()
    customers["Phone"] = customers["Phone"].fillna("").astype(str)
    customers = customers[customers["Customer"].ne("")].drop_duplicates("Key", keep="last")

    # Preserve visibility of every customer already present in the source workbook.
    if not unload.empty:
        existing = set(customers["Key"])
        historical = unload[["Customer", "Key"]].drop_duplicates("Key")
        missing = historical[~historical["Key"].isin(existing)].copy()
        if not missing.empty:
            missing["Status"] = "Active"
            missing["Phone"] = ""
            customers = pd.concat([customers, missing[["Customer", "Status", "Phone", "Key"]]], ignore_index=True)
    return customers


def load_rate_history(path: str) -> pd.DataFrame:
    rates = read_optional_sheet(path, "Customer Rates")
    required = ["Customer", "Rate", "Effective Date"]
    if rates.empty or not set(required).issubset(rates.columns):
        return pd.DataFrame(columns=required + ["Key"])
    rates = rates[required].copy()
    rates["Customer"] = rates["Customer"].astype(str).str.strip()
    rates["Key"] = rates["Customer"].map(_key)
    rates["Rate"] = pd.to_numeric(rates["Rate"], errors="coerce")
    rates["Effective Date"] = pd.to_datetime(rates["Effective Date"], errors="coerce")
    return rates.dropna(subset=["Rate", "Effective Date"])


def apply_effective_rates(unload: pd.DataFrame, rates: pd.DataFrame) -> pd.DataFrame:
    """Apply the latest effective customer rate only to unloading dated on/after it."""
    if unload.empty or rates.empty or "Date" not in unload.columns:
        return unload
    result = unload.copy()
    result["Date"] = pd.to_datetime(result["Date"], errors="coerce")
    rates = rates.sort_values("Effective Date")
    for key, group in rates.groupby("Key"):
        mask = result["Key"].eq(key) & result["Date"].notna()
        if not mask.any():
            continue
        for idx in result.index[mask]:
            row_date = result.at[idx, "Date"]
            eligible = group[group["Effective Date"] <= row_date]
            if not eligible.empty:
                result.at[idx, "Rate"] = float(eligible.iloc[-1]["Rate"])
    result["Amount"] = result["Unloaded Qty"] * result["Rate"]
    return result


def append_customer(path: str, customer: str, phone: str = "") -> None:
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb["Customers"] if "Customers" in wb.sheetnames else wb.create_sheet("Customers")
    headers = ["Customer", "Status", "Phone", "Created At"]
    if ws.cell(1, 1).value != "Customer":
        if ws.max_row > 1 or ws.cell(1, 1).value is not None:
            ws.insert_rows(1)
        for col, value in enumerate(headers, 1):
            ws.cell(1, col, value)
    ws.append([customer, "Active", phone, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    wb.save(path)


def set_customer_status(path: str, customer: str, status: str) -> None:
    from openpyxl import load_workbook
    wb = load_workbook(path)
    if "Customers" not in wb.sheetnames:
        ws = wb.create_sheet("Customers")
        for col, value in enumerate(["Customer", "Status", "Phone", "Created At"], 1):
            ws.cell(1, col, value)
    ws = wb["Customers"]
    headers = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
    if "Customer" not in headers or "Status" not in headers:
        raise ValueError("Customers sheet must contain Customer and Status columns.")
    found = False
    for r in range(2, ws.max_row + 1):
        if _key(ws.cell(r, headers["Customer"]).value) == _key(customer):
            ws.cell(r, headers["Status"], status)
            found = True
    if not found:
        ws.append([customer, status])
    wb.save(path)


def append_customer_rate(path: str, customer: str, rate: float, effective_date) -> None:
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb["Customer Rates"] if "Customer Rates" in wb.sheetnames else wb.create_sheet("Customer Rates")
    headers = ["Customer", "Rate", "Effective Date", "Recorded At"]
    if ws.cell(1, 1).value != "Customer":
        if ws.max_row > 1 or ws.cell(1, 1).value is not None:
            ws.insert_rows(1)
        for col, value in enumerate(headers, 1):
            ws.cell(1, col, value)
    ws.append([customer, float(rate), effective_date, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    wb.save(path)


# --------------------------------------------------------------------------
# Sidebar – Google Drive workbook connection (uploader is fallback)
# --------------------------------------------------------------------------
cfg = load_config()
st.sidebar.title("⚙️ Configuration")
drive_enabled = bool(google_drive_settings())

if drive_enabled:
    st.sidebar.success("Connected to Google Drive")
    refresh_drive = st.sidebar.button("🔄 Refresh from Google Drive", width="stretch")
    current_path = st.session_state.get("workbook_path", "")
    if refresh_drive or not current_path or not Path(current_path).is_file():
        tmp = tempfile.NamedTemporaryFile(prefix="cement_drive_", suffix=".xlsx", delete=False)
        tmp.close()
        try:
            download_workbook_from_drive(tmp.name)
            st.session_state["workbook_path"] = tmp.name
            st.session_state["workbook_upload_hash"] = "google-drive"
        except Exception as e:
            st.sidebar.error(f"Google Drive download failed: {e}")
    file_path = st.session_state.get("workbook_path")
else:
    st.sidebar.warning("Google Drive is not configured yet")
    uploaded_file = st.sidebar.file_uploader(
        "Upload CEMENT.xlsx (temporary mode)", type=["xlsx"],
        help="Configure Google Drive secrets to make workbook changes persistent across sessions.",
    )
    file_path = st.session_state.get("workbook_path")
    if uploaded_file is not None:
        uploaded_bytes = uploaded_file.getvalue()
        upload_hash = hashlib.sha256(uploaded_bytes).hexdigest()
        if st.session_state.get("workbook_upload_hash") != upload_hash or not file_path or not Path(file_path).is_file():
            tmp = tempfile.NamedTemporaryFile(prefix="cement_tracker_", suffix=".xlsx", delete=False)
            tmp.write(uploaded_bytes)
            tmp.flush()
            tmp.close()
            file_path = tmp.name
            st.session_state["workbook_path"] = file_path
            st.session_state["workbook_upload_hash"] = upload_hash

main_sheet = cfg.get("main_sheet", "Sheet1")
phone_sheet = cfg.get("phone_sheet", "Sheet2")
sheet_names = []
if file_path and Path(file_path).is_file():
    try:
        sheet_names = pd.ExcelFile(file_path, engine="openpyxl").sheet_names
    except Exception as e:
        st.sidebar.error(f"Cannot open workbook: {e}")
if sheet_names:
    main_sheet = st.sidebar.selectbox("Main sheet (Loading / Unloading / Payments)", sheet_names,
        index=sheet_names.index(main_sheet) if main_sheet in sheet_names else 0)
    phone_sheet = st.sidebar.selectbox("Phone number sheet", sheet_names,
        index=sheet_names.index(phone_sheet) if phone_sheet in sheet_names else min(1, len(sheet_names) - 1))
if st.sidebar.button("💾 Save configuration", width="stretch"):
    save_config({"main_sheet": main_sheet, "phone_sheet": phone_sheet, "template": cfg.get("template", DEFAULT_TEMPLATE)})
    st.sidebar.success("Saved")
if st.sidebar.button("🔄 Reload data", width="stretch"):
    st.cache_data.clear()
    st.session_state.pop("phone_table", None)
    st.rerun()

st.title("🏗️ YM CEMENT DISTRIBUTION – Outstanding Tracker")
if not file_path or not Path(file_path).is_file():
    st.info("Connect Google Drive using Streamlit secrets, or upload your CEMENT.xlsx workbook in the Configuration panel.")
    st.stop()
try:
    unload, loading, pay, phones, ledger_pay = read_workbook(
        file_path, os.path.getmtime(file_path), main_sheet, phone_sheet
    )
except Exception as e:
    st.error(f"Could not read the workbook: {e}")
    st.stop()

if unload.empty:
    st.warning("No unloading rows found in the sheet yet. Customer configuration is still available.")
    unload = pd.DataFrame(columns=["Customer", "Key", "Date", "PO", "Invoice", "Unloaded Qty", "Rate", "Amount", "Paid"])

customers = load_customers(file_path, unload)
# Use customer-master phone numbers as a fallback when the legacy phone sheet has no entry.
if "Phone" in customers.columns and not phones.empty:
    phone_existing = set(phones["Key"])
    fallback = customers.loc[~customers["Key"].isin(phone_existing) & customers["Phone"].ne(""), ["Key", "Phone"]]
    if not fallback.empty:
        phones = pd.concat([phones, fallback], ignore_index=True)
elif "Phone" in customers.columns:
    phones = customers.loc[customers["Phone"].ne(""), ["Key", "Phone"]].copy()
rate_history = load_rate_history(file_path)
unload = apply_effective_rates(unload, rate_history)
summary = summarise(unload, ledger_pay)

# Include active and inactive customer master records, even when they have no unloading yet.
for _, customer_row in customers.iterrows():
    key = customer_row["Key"]
    if key not in set(summary["Key"]):
        summary.loc[len(summary)] = {
            "Key": key, "Customer": customer_row["Customer"], "Unloaded Qty": 0.0,
            "Amount to Pay": 0.0, "Paid": 0.0, "Outstanding": 0.0,
            "Collection %": 0.0, "Avg Rate": 0.0,
        }
if st.session_state.get("payment_flash"):
    st.success(st.session_state.pop("payment_flash"))

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

tab_dash, tab_data, tab_payment, tab_customers, tab_rates, tab_sms = st.tabs([
    "📊 Dashboard", "📋 Details", "💳 Payment", "👥 Customers", "🏷️ Rate Config", "📱 SMS"
])

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

    if summary.empty:
        st.info("Add customers and enter unloading records to populate the dashboard charts.")
    else:
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
# PAYMENT ENTRY
# ==========================================================================
with tab_payment:
    st.subheader("Record a customer payment")
    st.caption("Payments entered here are appended to the workbook's 'Payment Entries' worksheet.")

    customer_options = customers.loc[customers["Status"].eq("Active"), "Customer"].dropna().astype(str).tolist()
    submitted = False
    if customer_options:
        with st.form("payment_entry_form", clear_on_submit=True):
            p_date = st.date_input("Payment date", value=datetime.now().date())
            p_customer = st.selectbox("Customer", customer_options)
            p_amount = st.number_input(
                "Payment amount (Rs.)",
                min_value=0.01,
                value=1000.00,
                step=100.00,
                format="%.2f",
            )
            p_reference = st.text_input("Reference / remarks (optional)")
            submitted = st.form_submit_button("💾 Save payment", type="primary")
    else:
        st.info("Add or reactivate a customer before recording a payment.")

    if submitted:
        try:
            append_payment_to_excel(file_path, p_date, p_customer, p_amount, p_reference.strip())
            sync_workbook(file_path)
            st.cache_data.clear()
            st.session_state.pop("phone_table", None)
            st.session_state["payment_flash"] = f"Payment of {rs(float(p_amount))} recorded for {p_customer}."
            st.rerun()
        except PermissionError:
            st.error("The workbook could not be updated. Please try again.")
        except Exception as e:
            st.error(f"Could not save payment: {e}")

    st.subheader("Payment history")
    if not pay.empty:
        history = pay.drop(columns=["Key"], errors="ignore").copy()
        st.dataframe(history, hide_index=True, width="stretch")
    else:
        st.info("No payments recorded yet.")

    st.download_button(
        "⬇️ Download updated Excel workbook",
        data=Path(file_path).read_bytes(),
        file_name="CEMENT_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    if drive_enabled:
        st.caption("The workbook is synchronized to Google Drive. Download a backup whenever needed.")
    else:
        st.caption("Temporary upload mode: download the updated workbook before ending this session. Configure Google Drive for persistent cloud storage.")

# ==========================================================================
# CUSTOMER MANAGEMENT
# ==========================================================================
with tab_customers:
    st.subheader("Customer configuration")
    with st.form("add_customer_form", clear_on_submit=True):
        new_customer = st.text_input("Customer name")
        new_phone = st.text_input("Phone number (optional)")
        add_customer_clicked = st.form_submit_button("➕ Add customer", type="primary")
    if add_customer_clicked:
        if not new_customer.strip():
            st.error("Please enter a customer name.")
        elif _key(new_customer) in set(customers["Key"]):
            st.error("This customer already exists. You can reactivate an inactive customer below.")
        else:
            try:
                append_customer(file_path, new_customer.strip(), new_phone.strip())
                sync_workbook(file_path)
                st.session_state["customer_flash"] = f"Customer {new_customer.strip()} added successfully."
                st.rerun()
            except Exception as e:
                st.error(f"Could not add customer: {e}")
    if st.session_state.get("customer_flash"):
        st.success(st.session_state.pop("customer_flash"))

    st.markdown("#### Customer status")
    customer_names = customers["Customer"].tolist()
    if customer_names:
        selected_customer = st.selectbox("Select customer", customer_names, key="customer_status_select")
        selected_status = customers.loc[customers["Key"].eq(_key(selected_customer)), "Status"].iloc[0]
        st.caption(f"Current status: {selected_status}")
        if selected_status == "Active":
            if st.button("Deactivate customer", type="secondary"):
                try:
                    set_customer_status(file_path, selected_customer, "Inactive")
                    sync_workbook(file_path)
                    st.success(f"{selected_customer} deactivated. Historical transactions are preserved.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not deactivate customer: {e}")
        else:
            if st.button("Reactivate customer", type="primary"):
                try:
                    set_customer_status(file_path, selected_customer, "Active")
                    sync_workbook(file_path)
                    st.success(f"{selected_customer} reactivated.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not reactivate customer: {e}")
    st.dataframe(customers[["Customer", "Status"]], hide_index=True, width="stretch")
    st.caption("Deactivation prevents new payment/customer selection but does not delete prior unloading or payment history.")

# ==========================================================================
# CUSTOMER RATE CONFIGURATION
# ==========================================================================
with tab_rates:
    st.subheader("Customer-wise rate configuration")
    st.caption("A new rate applies only to unloading records dated on or after its effective date. Earlier records retain their previous rate.")
    active_customers = customers.loc[customers["Status"].eq("Active"), "Customer"].tolist()
    if active_customers:
        with st.form("customer_rate_form", clear_on_submit=True):
            rate_customer = st.selectbox("Customer", active_customers)
            new_rate = st.number_input("New rate per bag (Rs.)", min_value=0.01, value=2000.00, step=50.00, format="%.2f")
            effective_date = st.date_input("Effective date", value=datetime.now().date(), min_value=datetime.now().date())
            save_rate_clicked = st.form_submit_button("💾 Save customer rate", type="primary")
        if save_rate_clicked:
            try:
                append_customer_rate(file_path, rate_customer, new_rate, effective_date)
                sync_workbook(file_path)
                st.session_state["rate_flash"] = f"Rate for {rate_customer} updated to Rs. {new_rate:,.2f} from {effective_date:%d %b %Y}."
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"Could not save rate: {e}")
    else:
        st.info("Add or reactivate a customer before configuring a rate.")
    if st.session_state.get("rate_flash"):
        st.success(st.session_state.pop("rate_flash"))
    st.markdown("#### Rate history")
    if not rate_history.empty:
        st.dataframe(rate_history[["Customer", "Rate", "Effective Date"]].sort_values(["Customer", "Effective Date"], ascending=[True, False]), hide_index=True, width="stretch")
    else:
        st.info("No customer-specific rate history has been configured yet. Existing workbook rates remain in use.")

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
            sync_workbook(file_path)
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
