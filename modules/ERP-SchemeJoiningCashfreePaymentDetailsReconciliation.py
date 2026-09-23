# app.py
"""
ERP Scheme Joining  ↔  Cashfree Payment Details  —  Reconciliation Tool
========================================================================
v3.0.0  —  Adds Dashboard & Date-wise Joining Details.
"""

from __future__ import annotations

import io
import logging
import re
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------- #
#  Logging
# --------------------------------------------------------------------------- #
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
)
logger = logging.getLogger("erp_cashfree_recon")

# --------------------------------------------------------------------------- #
#  Page config & styling
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="ERP ↔ Cashfree Reconciliation",
    page_icon="🔗",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
    .main .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
    .metric-card {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px;
        padding: 14px 16px; text-align: center;
    }
    .metric-card .label { font-size: 0.8rem; color: #64748b; text-transform: uppercase; letter-spacing: .5px; }
    .metric-card .value { font-size: 1.6rem; font-weight: 700; color: #0f172a; }
    .metric-card .sub   { font-size: 0.75rem; color: #94a3b8; margin-top: 4px; }
    .status-matched   { color: #15803d; font-weight: 600; }
    .status-unmatched { color: #b91c1c; font-weight: 600; }
    .status-dup       { color: #b45309; font-weight: 600; }
    div[data-testid="stDownloadButton"] button { width: 100%; }
    .kpi-row { margin-bottom: 10px; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
#  Constants
# --------------------------------------------------------------------------- #
APP_TITLE = "ERP Scheme Joining  ↔  Cashfree Payment Reconciliation"
APP_VERSION = "3.0.0"

ERP_PHONE_ALIASES = ["Mobileno", "Mobile No", "Mobile Number", "Mobile", "Phone", "Phone No", "Contact"]
ERP_DATE_ALIASES = ["Date", "Joining Date", "Payment Date", "Txn Date", "Transaction Date", "Paid Date"]
ERP_AMOUNT_ALIASES = ["Installment Amount", "Amount", "Paid Amount", "Payment Amount", "Installment Amt"]
ERP_DOC_ALIASES = ["Doc No", "DocNo", "Document No", "Doc Number", "Voucher No"]
ERP_SCHEME_ALIASES = ["Scheme"]
ERP_CUSTOMER_ALIASES = ["Customer", "Customer Name", "Cust Name"]
ERP_BOARD_RATE_ALIASES = ["Board Rate", "BoardRate", "Rate"]
ERP_BOOKED_WT_ALIASES = ["Booked Wt", "Booked Weight", "BookedWt", "Wt"]
ERP_ONLINE_ALIASES = ["Online", "Online Status"]
ERP_MATURITY_ALIASES = ["Maturity Dt", "Maturity Dt.", "Maturity Date", "MaturityDt"]

CF_REF_ALIASES = [
    "Transaction Reference", "Transaction Ref", "Reference", "Txn Reference",
    "Order ID", "Order Id", "Reference ID", "Ref ID", "Payment ID",
]
CF_NAME_ALIASES = ["Customer Name", "Name", "Payer Name"]
CF_EMAIL_ALIASES = ["Customer Email", "Email", "Email ID", "Payer Email"]
CF_PHONE_ALIASES = ["Customer Phone", "Phone", "Mobile", "Mobileno", "Contact"]
CF_CUST_ID_ALIASES = ["Customer ID", "Cust ID", "CustomerId", "Cust Id"]
CF_AMOUNT_ALIASES = ["Amount", "Paid Amount", "Payment Amount", "Txn Amount"]
CF_DATE_ALIASES = ["Date", "Payment Date", "Txn Date", "Transaction Date"]
CF_SCHEME_ALIASES = [
    "Scheme Name / Passbook Number", "Scheme Name", "Passbook Number",
    "Passbook No", "Scheme", "Passbook",
]

OUTPUT_COLUMNS: List[str] = [
    "S.No", "Transaction Reference", "Customer Name", "Customer Email",
    "Customer Phone", "Customer ID", "Amount", "Date",
    "Scheme Name / Passbook Number", "Status",
    "ERP Date", "Doc No", "Scheme", "Customer", "Mobileno",
    "Board Rate", "Booked Wt", "Online", "Maturity Dt.",
]

STATUS_MATCHED = "Matched"
STATUS_NOT_MATCHED = "Not Matched"
STATUS_DUPLICATE = "Duplicate Match"

# --------------------------------------------------------------------------- #
#  Data classes
# --------------------------------------------------------------------------- #
@dataclass
class MatchOptions:
    strict_three_fields: bool = True
    date_tolerance_days: int = 0
    amount_tolerance: float = 0.0
    flag_duplicates: bool = True
    strip_country_code: bool = True
    country_code: str = "91"


@dataclass
class FileBundle:
    raw: pd.DataFrame
    phone_col: str
    date_col: str
    amount_col: str
    extra: Dict[str, Optional[str]] = field(default_factory=dict)
    label: str = ""


# --------------------------------------------------------------------------- #
#  Session state
# --------------------------------------------------------------------------- #
def _init_state() -> None:
    defaults: Dict[str, Any] = {
        "erp_file_bytes": None,
        "erp_file_name": None,
        "cf_file_bytes": None,
        "cf_file_name": None,
        "erp_sheet": None,
        "cf_sheet": None,
        "output_df": None,
        "summary_df": None,
        "erp_count": 0,
        "cf_count": 0,
        "run_clicked": False,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


# --------------------------------------------------------------------------- #
#  Column detection helpers
# --------------------------------------------------------------------------- #
def _norm_col(name: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def find_col(df: pd.DataFrame, aliases: Sequence[str]) -> Optional[str]:
    if df is None or len(df.columns) == 0:
        return None
    lookup = {_norm_col(c): c for c in df.columns}
    for alias in aliases:
        key = _norm_col(alias)
        if key in lookup:
            return lookup[key]
    for alias in aliases:
        key = _norm_col(alias)
        if not key:
            continue
        for ncol, original in lookup.items():
            if key in ncol:
                return original
    return None


# --------------------------------------------------------------------------- #
#  Normalization utilities
# --------------------------------------------------------------------------- #
_PHONE_CLEAN_RE = re.compile(r"\D")


def normalize_phone(value: Any, strip_cc: bool = True, cc: str = "91") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, (int, np.integer)):
        s = str(int(value))
    elif isinstance(value, (float, np.floating)):
        s = str(int(value)) if float(value).is_integer() else str(value)
    else:
        s = str(value).strip()
        if s.lower() in {"nan", "none", "null", ""}:
            return ""

    digits = _PHONE_CLEAN_RE.sub("", s)
    if not digits:
        return ""
    if strip_cc and cc and digits.startswith(cc) and len(digits) == len(cc) + 10:
        digits = digits[len(cc):]
    if len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    return digits


_DATE_FORMATS = (
    "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%d-%m-%y", "%d/%m/%y",
    "%m/%d/%Y", "%m-%d-%Y",
    "%d-%b-%Y", "%d %b %Y", "%d-%b-%y", "%d %b %y",
    "%Y-%m-%d %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S",
    "%Y/%m/%d %H:%M:%S", "%d-%b-%Y %H:%M:%S",
)


def normalize_date(value: Any) -> pd.Timestamp:
    if value is None:
        return pd.NaT
    if isinstance(value, float) and np.isnan(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        try:
            return pd.Timestamp(value).normalize().tz_localize(None)
        except TypeError:
            return pd.Timestamp(value).normalize()
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        try:
            return (pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")).normalize()
        except Exception:
            return pd.NaT

    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null", "nat"}:
        return pd.NaT
    for fmt in _DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt)).normalize()
        except ValueError:
            continue
    try:
        ts = pd.to_datetime(s, dayfirst=True, errors="raise")
        return pd.Timestamp(ts).normalize().tz_localize(None)
    except Exception:
        return pd.NaT


_AMOUNT_CLEAN_RE = re.compile(r"[^\d.\-]")


def normalize_amount(value: Any) -> float:
    if value is None:
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        if isinstance(value, float) and np.isnan(value):
            return np.nan
        return float(value)
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null"}:
        return np.nan
    s = _AMOUNT_CLEAN_RE.sub("", s)
    if s in {"", "-", ".", "-."}:
        return np.nan
    if s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(s)
    except ValueError:
        return np.nan


def _format_date_string(value: Any) -> str:
    ts = normalize_date(value)
    if pd.isna(ts):
        if isinstance(value, str):
            s = value.strip()
            if s and "00:00:00" not in s and "T00:00" not in s:
                return s
        return ""
    return ts.strftime("%d-%m-%Y")


def _format_date_series(series: pd.Series) -> pd.Series:
    return series.apply(_format_date_string)


# --------------------------------------------------------------------------- #
#  File reading
# --------------------------------------------------------------------------- #
def list_excel_sheets(file) -> List[str]:
    try:
        try:
            file.seek(0)
        except Exception:
            pass
        xls = pd.ExcelFile(file)
        return xls.sheet_names
    except Exception:
        return []
    finally:
        try:
            file.seek(0)
        except Exception:
            pass


def read_table(file, sheet_name: Optional[str] = None) -> pd.DataFrame:
    name = getattr(file, "name", "") or ""
    name_l = name.lower()
    try:
        file.seek(0)
    except Exception:
        pass

    is_csv = name_l.endswith(".csv")
    if not is_csv and name_l == "":
        try:
            return pd.read_excel(file, sheet_name=sheet_name, dtype=object)
        except Exception:
            try:
                file.seek(0)
                return pd.read_csv(file, dtype=object, keep_default_na=False, encoding="utf-8")
            except Exception:
                pass
        raise RuntimeError("Unable to read file (unknown type)")

    if is_csv:
        for enc in ("utf-8", "utf-8-sig", "latin-1"):
            try:
                file.seek(0)
                return pd.read_csv(file, dtype=object, keep_default_na=False, encoding=enc)
            except Exception:
                continue
        raise RuntimeError(f"Unable to read CSV file: {name}")

    try:
        return pd.read_excel(file, sheet_name=sheet_name, dtype=object)
    except Exception as exc:
        raise RuntimeError(f"Unable to read Excel file {name}: {exc}") from exc


def build_bundle(df: pd.DataFrame, label: str, kind: str,
                 overrides: Optional[Dict[str, str]] = None) -> FileBundle:
    overrides = overrides or {}

    if kind == "erp":
        phone_col = overrides.get("phone") or find_col(df, ERP_PHONE_ALIASES)
        date_col = overrides.get("date") or find_col(df, ERP_DATE_ALIASES)
        amount_col = overrides.get("amount") or find_col(df, ERP_AMOUNT_ALIASES)
        extra = {
            "doc": find_col(df, ERP_DOC_ALIASES),
            "scheme": find_col(df, ERP_SCHEME_ALIASES),
            "customer": find_col(df, ERP_CUSTOMER_ALIASES),
            "board_rate": find_col(df, ERP_BOARD_RATE_ALIASES),
            "booked_wt": find_col(df, ERP_BOOKED_WT_ALIASES),
            "online": find_col(df, ERP_ONLINE_ALIASES),
            "maturity": find_col(df, ERP_MATURITY_ALIASES),
        }
    else:
        phone_col = overrides.get("phone") or find_col(df, CF_PHONE_ALIASES)
        date_col = overrides.get("date") or find_col(df, CF_DATE_ALIASES)
        amount_col = overrides.get("amount") or find_col(df, CF_AMOUNT_ALIASES)
        extra = {
            "ref": find_col(df, CF_REF_ALIASES),
            "name": find_col(df, CF_NAME_ALIASES),
            "email": find_col(df, CF_EMAIL_ALIASES),
            "cust_id": find_col(df, CF_CUST_ID_ALIASES),
            "scheme": find_col(df, CF_SCHEME_ALIASES),
        }

    missing = [n for n, c in (("phone", phone_col), ("date", date_col), ("amount", amount_col)) if not c]
    if missing:
        raise ValueError(f"[{label}] Could not detect required column(s): {', '.join(missing)}")

    return FileBundle(
        raw=df, phone_col=phone_col, date_col=date_col, amount_col=amount_col,
        extra=extra, label=label,
    )


# --------------------------------------------------------------------------- #
#  Normalization of bundles
# --------------------------------------------------------------------------- #
def _valid(row: pd.Series) -> bool:
    return (
        str(row["__phone"]).strip() != ""
        and pd.notna(row["__date"])
        and pd.notna(row["__amount"])
    )


def normalize_bundle(bundle: FileBundle, opts: MatchOptions) -> pd.DataFrame:
    df = bundle.raw.copy()
    df["__phone"] = df[bundle.phone_col].apply(
        lambda v: normalize_phone(v, opts.strip_country_code, opts.country_code)
    )
    df["__date"] = df[bundle.date_col].apply(normalize_date)
    df["__amount"] = df[bundle.amount_col].apply(normalize_amount)
    df["__amount_r"] = df["__amount"].round(2)
    df["__key"] = (
        df["__phone"].astype(str) + "|"
        + df["__date"].astype(str) + "|"
        + df["__amount_r"].astype(str)
    )
    df["__valid"] = df.apply(_valid, axis=1)
    return df


# --------------------------------------------------------------------------- #
#  Value getter
# --------------------------------------------------------------------------- #
def _get(row: pd.Series, col: Optional[str]) -> Any:
    if not col:
        return ""
    v = row.get(col, "")
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(v, (pd.Timestamp, datetime, np.datetime64)):
        return _format_date_string(v)
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float, np.integer, np.floating)):
        return v

    s = str(v).strip()
    if s == "" or s.lower() in {"nan", "nat", "none", "null"}:
        return ""
    if re.match(r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$", s):
        return _format_date_string(s)
    return s


# --------------------------------------------------------------------------- #
#  Matching engine
# --------------------------------------------------------------------------- #
def match_transactions(
    erp_df: pd.DataFrame,
    cf_df: pd.DataFrame,
    erp_bundle: FileBundle,
    cf_bundle: FileBundle,
    opts: MatchOptions,
    progress_cb=None,
) -> pd.DataFrame:
    erp_valid = erp_df[erp_df["__valid"]].copy()
    cf_valid = cf_df[cf_df["__valid"]].copy()

    if opts.strict_three_fields:
        erp_counts = erp_valid.groupby("__key").size().to_dict()
        cf_counts = cf_valid.groupby("__key").size().to_dict()
        erp_lookup = {k: g for k, g in erp_valid.groupby("__key")}

        def resolve(cf_row: pd.Series) -> Tuple[Optional[pd.Series], str]:
            if not cf_row["__valid"]:
                return None, STATUS_NOT_MATCHED
            key = cf_row["__key"]
            erp_grp = erp_lookup.get(key)
            if erp_grp is None or erp_grp.empty:
                return None, STATUS_NOT_MATCHED
            if opts.flag_duplicates and (erp_counts.get(key, 0) > 1 or cf_counts.get(key, 0) > 1):
                return erp_grp.iloc[0], STATUS_DUPLICATE
            return erp_grp.iloc[0], STATUS_MATCHED
    else:
        erp_by_phone: Dict[str, pd.DataFrame] = {
            p: g for p, g in erp_valid.groupby("__phone") if p
        }
        erp_key_counts = erp_valid.groupby("__key").size().to_dict()
        cf_key_counts = cf_valid.groupby("__key").size().to_dict()

        def resolve(cf_row: pd.Series) -> Tuple[Optional[pd.Series], str]:
            if not cf_row["__valid"]:
                return None, STATUS_NOT_MATCHED
            cands = erp_by_phone.get(cf_row["__phone"])
            if cands is None or cands.empty:
                return None, STATUS_NOT_MATCHED

            mask = pd.Series(True, index=cands.index)
            if opts.amount_tolerance > 0:
                mask &= (cands["__amount"] - cf_row["__amount"]).abs() <= opts.amount_tolerance
            else:
                mask &= cands["__amount_r"] == cf_row["__amount_r"]

            if opts.date_tolerance_days > 0:
                delta = (cands["__date"] - cf_row["__date"]).abs()
                mask &= delta <= pd.Timedelta(days=opts.date_tolerance_days)
            else:
                mask &= cands["__date"] == cf_row["__date"]

            hits = cands[mask]
            if hits.empty:
                return None, STATUS_NOT_MATCHED

            key = cf_row["__key"]
            if opts.flag_duplicates and (
                len(hits) > 1 or erp_key_counts.get(key, 0) > 1 or cf_key_counts.get(key, 0) > 1
            ):
                return hits.iloc[0], STATUS_DUPLICATE
            return hits.iloc[0], STATUS_MATCHED

    erp_extra = erp_bundle.extra
    cf_extra = cf_bundle.extra
    total = len(cf_df)
    records: List[Dict[str, Any]] = []

    for i, (_, cf_row) in enumerate(cf_df.iterrows(), start=1):
        erp_row, status = resolve(cf_row)

        rec: Dict[str, Any] = {
            "S.No": i,
            "Transaction Reference": _get(cf_row, cf_extra.get("ref")),
            "Customer Name": _get(cf_row, cf_extra.get("name")),
            "Customer Email": _get(cf_row, cf_extra.get("email")),
            "Customer Phone": _get(cf_row, cf_bundle.phone_col),
            "Customer ID": _get(cf_row, cf_extra.get("cust_id")),
            "Amount": _get(cf_row, cf_bundle.amount_col),
            "Date": _get(cf_row, cf_bundle.date_col),
            "Scheme Name / Passbook Number": _get(cf_row, cf_extra.get("scheme")),
            "Status": status,
            "ERP Date": "", "Doc No": "", "Scheme": "", "Customer": "",
            "Mobileno": "", "Board Rate": "", "Booked Wt": "",
            "Online": "", "Maturity Dt.": "",
        }

        if erp_row is not None:
            rec["ERP Date"] = _get(erp_row, erp_bundle.date_col)
            rec["Doc No"] = _get(erp_row, erp_extra.get("doc"))
            rec["Scheme"] = _get(erp_row, erp_extra.get("scheme"))
            rec["Customer"] = _get(erp_row, erp_extra.get("customer"))
            rec["Mobileno"] = _get(erp_row, erp_bundle.phone_col)
            rec["Board Rate"] = _get(erp_row, erp_extra.get("board_rate"))
            rec["Booked Wt"] = _get(erp_row, erp_extra.get("booked_wt"))
            rec["Online"] = _get(erp_row, erp_extra.get("online"))
            rec["Maturity Dt."] = _get(erp_row, erp_extra.get("maturity"))

        records.append(rec)
        if progress_cb and i % 500 == 0:
            progress_cb(i / max(total, 1))

    if progress_cb:
        progress_cb(1.0)

    return pd.DataFrame(records, columns=OUTPUT_COLUMNS)


# --------------------------------------------------------------------------- #
#  Export helpers
# --------------------------------------------------------------------------- #
def to_excel_bytes(df: pd.DataFrame, summary: Optional[pd.DataFrame] = None,
                   daily: Optional[pd.DataFrame] = None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="All Results")
        df[df["Status"] == STATUS_MATCHED].to_excel(writer, index=False, sheet_name="Matched")
        df[df["Status"] == STATUS_NOT_MATCHED].to_excel(writer, index=False, sheet_name="Not Matched")
        df[df["Status"] == STATUS_DUPLICATE].to_excel(writer, index=False, sheet_name="Duplicate Match")
        if summary is not None:
            summary.to_excel(writer, index=False, sheet_name="Summary")
        if daily is not None and not daily.empty:
            daily.to_excel(writer, index=False, sheet_name="Date-wise")

        for sheet in writer.sheets.values():
            for col_cells in sheet.columns:
                max_len = 0
                letter = col_cells[0].column_letter
                for cell in col_cells:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                sheet.column_dimensions[letter].width = min(max(max_len + 2, 10), 45)

    return buf.getvalue()


def build_summary(output_df: pd.DataFrame, erp_count: int, cf_count: int) -> pd.DataFrame:
    total = len(output_df)
    matched = int((output_df["Status"] == STATUS_MATCHED).sum())
    unmatched = int((output_df["Status"] == STATUS_NOT_MATCHED).sum())
    dup = int((output_df["Status"] == STATUS_DUPLICATE).sum())
    rate = (matched / total * 100) if total else 0.0

    amount_total = pd.to_numeric(output_df["Amount"], errors="coerce").fillna(0).sum()
    amount_matched = pd.to_numeric(
        output_df.loc[output_df["Status"] == STATUS_MATCHED, "Amount"], errors="coerce"
    ).fillna(0).sum()

    rows = [
        ("Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("App Version", APP_VERSION),
        ("ERP Rows", erp_count),
        ("Cashfree Rows", cf_count),
        ("Total Cashfree Rows Exported", total),
        ("Matched", matched),
        ("Not Matched", unmatched),
        ("Duplicate Match", dup),
        ("Match Rate (%)", round(rate, 2)),
        ("Total Amount (Cashfree)", round(float(amount_total), 2)),
        ("Total Amount (Matched)", round(float(amount_matched), 2)),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


# --------------------------------------------------------------------------- #
#  Date-wise Joining Details
# --------------------------------------------------------------------------- #
def build_daily_details(output_df: pd.DataFrame) -> pd.DataFrame:
    """
    Produce a date-wise breakdown using the Cashfree Date column.
    Columns: Date, Total, Matched, Not Matched, Duplicate Match, Match Rate %,
             Total Amount, Matched Amount, Unique Customers (by phone).
    """
    if output_df is None or output_df.empty:
        return pd.DataFrame()

    df = output_df.copy()
    df["__dt"] = df["Date"].apply(normalize_date)
    df = df[df["__dt"].notna()]
    if df.empty:
        return pd.DataFrame()

    df["__amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)

    grouped = df.groupby("__dt")
    rows = []
    for dt, grp in grouped:
        total = len(grp)
        m = int((grp["Status"] == STATUS_MATCHED).sum())
        nm = int((grp["Status"] == STATUS_NOT_MATCHED).sum())
        d = int((grp["Status"] == STATUS_DUPLICATE).sum())
        amt_total = float(grp["__amount"].sum())
        amt_matched = float(grp.loc[grp["Status"] == STATUS_MATCHED, "__amount"].sum())
        unique_cust = grp["Customer Phone"].astype(str).nunique()
        rows.append({
            "Date": dt.strftime("%d-%m-%Y"),
            "__sort": dt,
            "Total": total,
            "Matched": m,
            "Not Matched": nm,
            "Duplicate Match": d,
            "Match Rate (%)": round((m / total * 100) if total else 0.0, 2),
            "Total Amount": round(amt_total, 2),
            "Matched Amount": round(amt_matched, 2),
            "Unique Customers": unique_cust,
        })

    out = pd.DataFrame(rows).sort_values("__sort", ascending=False).drop(columns="__sort").reset_index(drop=True)
    return out


# --------------------------------------------------------------------------- #
#  Sidebar
# --------------------------------------------------------------------------- #
def render_sidebar():
    st.sidebar.title("⚙️ Configuration")

    st.sidebar.subheader("1. Upload files")
    erp_file = st.sidebar.file_uploader(
        "ERP Scheme Joining Report", type=["xlsx", "xls", "csv"], key="erp_file_widget",
    )
    cf_file = st.sidebar.file_uploader(
        "Cashfree Payment Details", type=["xlsx", "xls", "csv"], key="cf_file_widget",
    )

    if erp_file is not None:
        st.session_state["erp_file_bytes"] = erp_file.getvalue()
        st.session_state["erp_file_name"] = erp_file.name
    if cf_file is not None:
        st.session_state["cf_file_bytes"] = cf_file.getvalue()
        st.session_state["cf_file_name"] = cf_file.name

    erp_name = st.session_state.get("erp_file_name")
    cf_name = st.session_state.get("cf_file_name")

    if erp_name and erp_name.lower().endswith((".xlsx", ".xls")):
        buf = io.BytesIO(st.session_state["erp_file_bytes"])
        sheets = list_excel_sheets(buf)
        if sheets:
            current = st.session_state.get("erp_sheet")
            idx = sheets.index(current) if current in sheets else 0
            chosen = st.sidebar.selectbox("ERP sheet", sheets, index=idx, key="erp_sheet_widget")
            st.session_state["erp_sheet"] = chosen
    else:
        st.session_state["erp_sheet"] = None

    if cf_name and cf_name.lower().endswith((".xlsx", ".xls")):
        buf = io.BytesIO(st.session_state["cf_file_bytes"])
        sheets = list_excel_sheets(buf)
        if sheets:
            current = st.session_state.get("cf_sheet")
            idx = sheets.index(current) if current in sheets else 0
            chosen = st.sidebar.selectbox("Cashfree sheet", sheets, index=idx, key="cf_sheet_widget")
            st.session_state["cf_sheet"] = chosen
    else:
        st.session_state["cf_sheet"] = None

    st.sidebar.subheader("2. Matching options")
    strict = st.sidebar.checkbox("Strict 3-field match (Phone + Date + Amount)", value=True)
    date_tol = 0
    amt_tol = 0.0
    if not strict:
        date_tol = st.sidebar.number_input("Date tolerance (± days)", 0, 15, 1, 1)
        amt_tol = st.sidebar.number_input("Amount tolerance (±)", 0.0, 100.0, 1.0, 0.5)

    flag_dup = st.sidebar.checkbox("Flag duplicate matches", value=True)
    strip_cc = st.sidebar.checkbox("Strip country code from phone", value=True)
    cc = st.sidebar.text_input("Country code", value="91", disabled=not strip_cc, max_chars=4)

    opts = MatchOptions(
        strict_three_fields=strict,
        date_tolerance_days=int(date_tol),
        amount_tolerance=float(amt_tol),
        flag_duplicates=flag_dup,
        strip_country_code=strip_cc,
        country_code=cc.strip() or "91",
    )

    st.sidebar.subheader("3. Column overrides (optional)")
    st.sidebar.caption("Leave blank for auto-detection.")
    erp_over = {
        "phone":  st.sidebar.text_input("ERP phone column", ""),
        "date":   st.sidebar.text_input("ERP date column", ""),
        "amount": st.sidebar.text_input("ERP amount column", ""),
    }
    cf_over = {
        "phone":  st.sidebar.text_input("Cashfree phone column", ""),
        "date":   st.sidebar.text_input("Cashfree date column", ""),
        "amount": st.sidebar.text_input("Cashfree amount column", ""),
    }

    if st.sidebar.button("🚀 Run Reconciliation", type="primary", width="stretch"):
        st.session_state["run_clicked"] = True

    if st.sidebar.button("♻️ Reset", width="stretch"):
        for k in ("erp_file_bytes", "erp_file_name", "cf_file_bytes", "cf_file_name",
                  "erp_sheet", "cf_sheet", "output_df", "summary_df"):
            st.session_state[k] = None
        st.session_state["erp_count"] = 0
        st.session_state["cf_count"] = 0
        st.session_state["run_clicked"] = False
        st.rerun()

    return erp_over, cf_over, opts


# --------------------------------------------------------------------------- #
#  UI helpers
# --------------------------------------------------------------------------- #
def render_metric(label: str, value: Any, css_class: str = "", sub: str = "") -> None:
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="label">{label}</div>
            <div class="value {css_class}">{value}</div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard(output_df: pd.DataFrame, erp_count: int, cf_count: int) -> None:
    st.markdown("### 📊 Dashboard")

    total = len(output_df)
    matched = int((output_df["Status"] == STATUS_MATCHED).sum())
    unmatched = int((output_df["Status"] == STATUS_NOT_MATCHED).sum())
    dup = int((output_df["Status"] == STATUS_DUPLICATE).sum())
    rate = (matched / total * 100) if total else 0.0

    amt_series = pd.to_numeric(output_df["Amount"], errors="coerce").fillna(0.0)
    amt_total = float(amt_series.sum())
    amt_matched = float(amt_series[output_df["Status"] == STATUS_MATCHED].sum())
    amt_unmatched = float(amt_series[output_df["Status"] == STATUS_NOT_MATCHED].sum())

    # ---------- KPI row 1: counts ----------
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: render_metric("Total Cashfree Txns", f"{total:,}")
    with c2: render_metric("Matched", f"{matched:,}", "status-matched")
    with c3: render_metric("Not Matched", f"{unmatched:,}", "status-unmatched")
    with c4: render_metric("Duplicate Match", f"{dup:,}", "status-dup")
    with c5: render_metric("Match Rate", f"{rate:.2f}%")

    # ---------- KPI row 2: amounts ----------
    c6, c7, c8, c9, c10 = st.columns(5)
    with c6: render_metric("Total Amount", f"₹{amt_total:,.2f}")
    with c7: render_metric("Matched Amount", f"₹{amt_matched:,.2f}", "status-matched")
    with c8: render_metric("Unmatched Amount", f"₹{amt_unmatched:,.2f}", "status-unmatched")
    with c9: render_metric("ERP Rows", f"{erp_count:,}")
    with c10: render_metric("Cashfree Rows", f"{cf_count:,}")

    st.markdown("---")

    # ---------- Status distribution + top dates ----------
    daily = build_daily_details(output_df)

    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.markdown("#### Status Distribution")
        dist = pd.DataFrame({
            "Status": [STATUS_MATCHED, STATUS_NOT_MATCHED, STATUS_DUPLICATE],
            "Count": [matched, unmatched, dup],
        }).set_index("Status")
        st.bar_chart(dist, width="stretch", height=280)

    with col_right:
        st.markdown("#### Transactions by Date (Top 15)")
        if not daily.empty:
            top = daily.head(15).copy()
            top = top.set_index("Date")[["Matched", "Not Matched", "Duplicate Match"]]
            st.bar_chart(top, width="stretch", height=280)
        else:
            st.info("No valid dates found in Cashfree data.")

    # ---------- Match rate over time ----------
    if not daily.empty:
        st.markdown("#### Match Rate by Date")
        trend = daily.copy()
        # Display oldest -> newest for a natural time axis
        trend = trend.sort_values(
            "Date", key=lambda s: pd.to_datetime(s, format="%d-%m-%Y")
        )
        trend = trend.set_index("Date")[["Match Rate (%)"]]
        st.line_chart(trend, width="stretch", height=260)

    # ---------- Amount flow by date ----------
    if not daily.empty:
        st.markdown("#### Amount Flow by Date")
        flow = daily.sort_values(
            "Date", key=lambda s: pd.to_datetime(s, format="%d-%m-%Y")
        ).set_index("Date")[["Total Amount", "Matched Amount"]]
        st.area_chart(flow, width="stretch", height=260)


def render_datewise_tab(output_df: pd.DataFrame) -> None:
    st.markdown("### 📅 Date-wise Joining Details")
    st.caption(
        "Grouped by the Cashfree transaction date. Every Cashfree row is included in "
        "the **Total** column; **Matched** counts only rows where Phone + Date + Amount "
        "matched an ERP record."
    )

    daily = build_daily_details(output_df)
    if daily.empty:
        st.info("No valid dates found to aggregate. Check that the Cashfree Date column is populated.")
        return

    # ---- Filters ----
    f1, f2 = st.columns([2, 2])
    with f1:
        date_min = pd.to_datetime(daily["Date"], format="%d-%m-%Y").min().date()
        date_max = pd.to_datetime(daily["Date"], format="%d-%m-%Y").max().date()
        date_range = st.date_input(
            "Date range",
            value=(date_min, date_max),
            min_value=date_min,
            max_value=date_max,
        )
    with f2:
        min_rate = st.slider("Minimum match rate (%)", 0, 100, 0, 5)

    # Apply filters
    daily_f = daily.copy()
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start, end = date_range
        daily_f = daily_f[
            (pd.to_datetime(daily_f["Date"], format="%d-%m-%Y").dt.date >= start)
            & (pd.to_datetime(daily_f["Date"], format="%d-%m-%Y").dt.date <= end)
        ]
    daily_f = daily_f[daily_f["Match Rate (%)"] >= min_rate]

    # ---- KPI strip ----
    if not daily_f.empty:
        k1, k2, k3, k4 = st.columns(4)
        with k1: render_metric("Days covered", f"{len(daily_f)}")
        with k2: render_metric("Total Txns", f"{int(daily_f['Total'].sum()):,}")
        with k3: render_metric("Matched", f"{int(daily_f['Matched'].sum()):,}", "status-matched")
        with k4: render_metric("Amount (Matched)", f"₹{daily_f['Matched Amount'].sum():,.2f}")

    # ---- Table ----
    st.dataframe(daily_f, width="stretch", height=420, hide_index=True)

    # ---- Download of the date-wise table ----
    csv_bytes = daily_f.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Download Date-wise CSV",
        data=csv_bytes,
        file_name=f"datewise_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        key="dl_datewise_csv",
    )

    # ---- Drill-down: pick a date to see that day's transactions ----
    st.markdown("#### 🔎 Drill-down — pick a date to see its transactions")
    picked = st.selectbox(
        "Select a date",
        options=["—"] + daily_f["Date"].tolist(),
        index=0,
        key="drill_date",
    )
    if picked and picked != "—":
        mask = output_df["Date"].apply(
            lambda v: _format_date_string(v) == picked
        )
        day_df = output_df[mask]
        st.caption(f"{len(day_df):,} transaction(s) on **{picked}**")
        st.dataframe(day_df, width="stretch", height=420, hide_index=True)


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def main() -> None:
    _init_state()

    st.title(APP_TITLE)
    st.caption(f"v{APP_VERSION} · Match by Phone + Date + Amount")

    erp_over, cf_over, opts = render_sidebar()

    erp_bytes = st.session_state.get("erp_file_bytes")
    cf_bytes = st.session_state.get("cf_file_bytes")

    if not erp_bytes or not cf_bytes:
        st.info("👈 Upload **both** files from the sidebar to begin.")
        st.stop()

    # ---------- Load ----------
    try:
        with st.spinner("Reading uploaded files…"):
            erp_raw = read_table(io.BytesIO(erp_bytes), st.session_state.get("erp_sheet"))
            cf_raw = read_table(io.BytesIO(cf_bytes), st.session_state.get("cf_sheet"))
    except Exception as exc:
        st.error(f"Failed to read files: {exc}")
        st.code(traceback.format_exc())
        st.stop()

    st.success(f"Loaded · ERP rows: **{len(erp_raw):,}** · Cashfree rows: **{len(cf_raw):,}**")

    # ---------- Column detection ----------
    try:
        erp_bundle = build_bundle(erp_raw, "ERP", "erp", erp_over)
        cf_bundle = build_bundle(cf_raw, "Cashfree", "cf", cf_over)
    except ValueError as exc:
        st.error(str(exc))
        with st.expander("Show available columns"):
            c1, c2 = st.columns(2)
            c1.write("**ERP columns**");  c1.write(list(erp_raw.columns))
            c2.write("**Cashfree columns**"); c2.write(list(cf_raw.columns))
        st.stop()

    with st.expander("🧭 Detected columns", expanded=False):
        c1, c2 = st.columns(2)
        c1.markdown("**ERP**")
        c1.json({"phone": erp_bundle.phone_col, "date": erp_bundle.date_col,
                 "amount": erp_bundle.amount_col, **erp_bundle.extra})
        c2.markdown("**Cashfree**")
        c2.json({"phone": cf_bundle.phone_col, "date": cf_bundle.date_col,
                 "amount": cf_bundle.amount_col, **cf_bundle.extra})

    # ---------- Compute only on Run (or first time) ----------
    need_compute = (
        st.session_state.get("run_clicked", False)
        or st.session_state.get("output_df") is None
    )

    if need_compute:
        try:
            with st.spinner("Normalizing data…"):
                erp_norm = normalize_bundle(erp_bundle, opts)
                cf_norm = normalize_bundle(cf_bundle, opts)

            progress = st.progress(0.0, text="Matching transactions…")
            try:
                output_df = match_transactions(
                    erp_norm, cf_norm, erp_bundle, cf_bundle, opts,
                    progress_cb=lambda p: progress.progress(
                        p, text=f"Matching… {int(p * 100)}%"
                    ),
                )
            finally:
                progress.empty()

            for col in ("ERP Date", "Maturity Dt."):
                if col in output_df.columns:
                    output_df[col] = output_df[col].apply(_format_date_string)

            st.session_state["output_df"] = output_df
            st.session_state["summary_df"] = build_summary(
                output_df, len(erp_norm), len(cf_norm)
            )
            st.session_state["erp_count"] = len(erp_norm)
            st.session_state["cf_count"] = len(cf_norm)
            st.session_state["run_clicked"] = False
        except Exception as exc:
            st.error(f"Matching failed: {exc}")
            st.code(traceback.format_exc())
            st.stop()

    output_df = st.session_state["output_df"]
    summary_df = st.session_state["summary_df"]
    erp_count = st.session_state["erp_count"]
    cf_count = st.session_state["cf_count"]

    # ---------- Tabs ----------
    tab_dash, tab_daily, tab_results, tab_download = st.tabs(
        ["📊 Dashboard", "📅 Date-wise Joining Details", "📋 All Results", "⬇️ Downloads"]
    )

    with tab_dash:
        render_dashboard(output_df, erp_count, cf_count)

    with tab_daily:
        render_datewise_tab(output_df)

    with tab_results:
        st.markdown("### 🔍 Results")
        status_filter = st.multiselect(
            "Filter by status",
            options=[STATUS_MATCHED, STATUS_NOT_MATCHED, STATUS_DUPLICATE],
            default=[STATUS_MATCHED, STATUS_NOT_MATCHED, STATUS_DUPLICATE],
            key="status_filter_widget",
        )
        view = output_df[output_df["Status"].isin(status_filter)] if status_filter else output_df
        st.dataframe(view, width="stretch", height=520, hide_index=True)

    with tab_download:
        st.markdown("### ⬇️ Downloads")
        st.caption(
            "Excel workbook contains: All Results · Matched · Not Matched · "
            "Duplicate Match · Summary · Date-wise"
        )
        daily_df = build_daily_details(output_df)
        xlsx_bytes = to_excel_bytes(output_df, summary_df, daily_df)
        csv_bytes = output_df.to_csv(index=False).encode("utf-8")
        daily_csv = daily_df.to_csv(index=False).encode("utf-8") if not daily_df.empty else b""

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "📥 Download Excel (multi-sheet)",
                data=xlsx_bytes,
                file_name=f"erp_cashfree_reconciliation_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                key="dl_xlsx",
            )
        with d2:
            st.download_button(
                "📥 Download CSV (all results)",
                data=csv_bytes,
                file_name=f"erp_cashfree_reconciliation_{ts}.csv",
                mime="text/csv",
                width="stretch",
                key="dl_csv",
            )
        if daily_csv:
            d3, d4 = st.columns(2)
            with d3:
                st.download_button(
                    "📥 Download CSV (date-wise)",
                    data=daily_csv,
                    file_name=f"datewise_{ts}.csv",
                    mime="text/csv",
                    width="stretch",
                    key="dl_daily_csv",
                )

    st.caption("Tip: switch tabs to explore the Dashboard, Date-wise breakdown, and full Results.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logger.exception("Fatal error in reconciliation app")
        st.error("An unexpected error occurred. See logs for details.")
        st.code(traceback.format_exc())