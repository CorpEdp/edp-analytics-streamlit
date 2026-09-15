"""
Auto-converted from: App-CampaignTransactionAnalyzer.py
Review this file before using — the converter does a best-effort wrap;
double-check indentation around any unusual control flow (loops, if/else
blocks that span large sections, etc.).
"""

import streamlit as st
import pandas as pd
import numpy as np
import time
import traceback
import re
import zipfile
from io import BytesIO
from pathlib import Path
def _excel_read_engine():
    try:
        import python_calamine  # noqa: F401
        return "calamine"
    except ImportError:
        return "openpyxl"
def clean_column_names(df):
    df = df.copy()
    df.columns = (
        pd.Index(df.columns).astype(str).str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    return df
def apply_column_aliases(df, aliases):
    df = df.copy()
    normalized = {str(col).lower().strip(): col for col in df.columns}
    rename_map = {}
    for alias, target in aliases.items():
        if target in df.columns:
            continue
        alias_key = str(alias).lower().strip()
        if alias_key in normalized:
            actual_col = normalized[alias_key]
            if actual_col not in rename_map and actual_col != target:
                rename_map[actual_col] = target
    if rename_map:
        df = df.rename(columns=rename_map)
    return df
def ensure_columns(df, required_columns, file_name=""):
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        found = sorted([str(c) for c in df.columns.tolist()])
        raise ValueError(
            f"Missing required columns in {file_name or 'file'}:\n"
            + "\n".join(f"  ❌ {x}" for x in missing)
            + "\n\nColumns actually found in the file:\n"
            + "\n".join(f"  ✓ {x}" for x in found)
        )
def normalize_phone_series(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.str.replace(r"\D", "", regex=True)
    s = s.str.replace(r"^(?:91|0)(\d{10})$", r"\1", regex=True)
    s = s.replace("nan", "")
    return s.fillna("")
def _read_excel_with_fallback(uploaded_file, sheet_names=None):
    engines = [_excel_read_engine()]
    if engines[0] != "openpyxl":
        engines.append("openpyxl")
    last_error = None
    for engine in engines:
        try:
            uploaded_file.seek(0)
            xls = pd.ExcelFile(uploaded_file, engine=engine)
            names = (
                xls.sheet_names if sheet_names is None
                else xls.sheet_names[:sheet_names]
            )
            result = {}
            for name in names:
                result[name] = xls.parse(name)
            return result, engine
        except Exception as e:
            last_error = e
            if engine == engines[-1]:
                raise
            st.warning(
                f"⚠️ {engine} failed on '{uploaded_file.name}' — "
                f"retrying with fallback engine..."
            )
    raise last_error
def read_single_campaign_file(uploaded_file):
    file_name = uploaded_file.name
    extension = Path(file_name).suffix.lower()
    frames = []

    if extension in (".xlsx", ".xls"):
        sheets, _ = _read_excel_with_fallback(
            uploaded_file, sheet_names=2
        )
        for sheet_name, df in sheets.items():
            if df is None or df.empty:
                continue
            df = clean_column_names(df)
            df["Campaign File Name"] = file_name
            df["Campaign Sheet Name"] = sheet_name
            frames.append(df)
    elif extension == ".csv":
        df = pd.read_csv(uploaded_file, low_memory=False)
        df = clean_column_names(df)
        df["Campaign File Name"] = file_name
        df["Campaign Sheet Name"] = "CSV"
        frames.append(df)
    else:
        raise ValueError(
            f"Unsupported Campaign file format: {file_name}"
        )

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)
def read_campaign_files(uploaded_files):
    all_frames = []
    total = len(uploaded_files)
    progress = st.progress(0, text="📥 Reading Campaign files...")
    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            df = read_single_campaign_file(uploaded_file)
            if not df.empty:
                all_frames.append(df)
        except Exception as e:
            st.warning(
                f"⚠️ Skipped '{uploaded_file.name}': "
                f"{type(e).__name__}: {e}"
            )
        progress.progress(
            (idx + 1) / total,
            text=f"📥 Read {idx + 1}/{total} campaign files"
        )
    progress.empty()
    if not all_frames:
        return pd.DataFrame()
    return pd.concat(all_frames, ignore_index=True)
def process_campaign_data(df):
    if df.empty:
        return df

    df = clean_column_names(df)
    df = apply_column_aliases(df, CAMPAIGN_COLUMN_ALIASES)
    ensure_columns(df, CAMPAIGN_REQUIRED_COLUMNS, "Campaign data")

    df["number"] = normalize_phone_series(df["number"])
    df = df[df["number"].ne("")].copy()

    raw_statuses = (
        df["status"].astype(str).str.strip().str.lower().unique().tolist()
    )

    df["status"] = df["status"].astype(str).str.strip().str.lower()

    before = len(df)
    df = df[df["status"].isin(ALLOWED_CAMPAIGN_STATUSES)].copy()
    after = len(df)

    if after == 0:
        raise ValueError(
            f"All {before} campaign rows were filtered out by status.\n"
            f"Allowed statuses: {list(ALLOWED_CAMPAIGN_STATUSES)}\n"
            f"Statuses found in file: {raw_statuses}"
        )

    df["status"] = df["status"].astype("string")

    for col in ["sent_at", "delivered_at", "read_at", "failed_at"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["_campaign_date"] = (
        df["read_at"]
        .combine_first(df["delivered_at"])
        .combine_first(df["sent_at"])
    )

    df["_status_priority"] = (
        df["status"].astype(str).map(STATUS_PRIORITY).fillna(0)
    )

    df = df.sort_values(
        ["number", "_status_priority", "_campaign_date"],
        ascending=[True, False, False],
        kind="mergesort"
    )
    df = df.drop_duplicates(subset=["number"], keep="first")

    df = df.drop(
        columns=["_campaign_date", "_status_priority"],
        errors="ignore"
    )

    df["Campaign File Name"] = df["Campaign File Name"].astype("string")

    return df.reset_index(drop=True)
def read_transaction_file(uploaded_file):
    file_name = uploaded_file.name
    extension = Path(file_name).suffix.lower()

    if extension in (".xlsx", ".xls"):
        sheets, _ = _read_excel_with_fallback(
            uploaded_file, sheet_names=1
        )
        df = list(sheets.values())[0]
    elif extension == ".csv":
        df = pd.read_csv(uploaded_file, low_memory=False)
    else:
        raise ValueError(
            f"Unsupported Transaction file format: {file_name}"
        )

    return clean_column_names(df)
def process_transaction_data(df):
    if df.empty:
        return df

    df = clean_column_names(df)
    df = apply_column_aliases(df, TRANSACTION_COLUMN_ALIASES)
    ensure_columns(df, TRANSACTION_REQUIRED_COLUMNS, "Transaction data")

    df["Customer Phone Number"] = normalize_phone_series(
        df["Customer Phone Number"]
    )

    df["Installment number"] = pd.to_numeric(
        df["Installment number"], errors="coerce"
    )

    df = df[
        df["Installment number"].notna()
        & (df["Installment number"] > 0)
    ].copy()

    df["Transaction Type"] = np.where(
        df["Installment number"].eq(1),
        ENROLLMENT,
        COLLECTION
    )
    # ✅ FIX #1: use nullable string dtype instead of category
    df["Transaction Type"] = df["Transaction Type"].astype("string")

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    numeric_columns = [
        "Saved Amount", "Refunded Amount", "Reward Amount",
        "Installment number", "Metal Rate", "Saved Metal Weight",
        "Rewards Metal Weight", "Benefit Metal Amount",
        "Benefit Metal Weight", "Benefit Metal Percentage",
    ]
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df.reset_index(drop=True)
def create_final_data(campaign_df, transaction_df):
    if campaign_df.empty:
        return pd.DataFrame()

    campaign_phone_set = set(campaign_df["number"].dropna().unique())

    transaction_matched = transaction_df[
        transaction_df["Customer Phone Number"].isin(campaign_phone_set)
    ].copy()

    final_df = campaign_df.merge(
        transaction_matched,
        left_on="number",
        right_on="Customer Phone Number",
        how="left",
        suffixes=("", "_Transaction"),
        sort=False,
        copy=False
    )

    matched_mask = (
        final_df["Customer Phone Number"].fillna("").astype(str).ne("")
    )

    final_df["Transaction Match"] = np.where(
        matched_mask, MATCHED, NOT_MATCHED
    )

    return final_df
def get_matched_transactions(final_df):
    if final_df.empty:
        return final_df
    return final_df[final_df["Transaction Match"].eq(MATCHED)]
def _make_zero_rows(
    file_names,
    scheme_label=NO_MATCH,
    transaction_types=(ENROLLMENT, COLLECTION),
):
    rows = []
    for fn in file_names:
        for tt in transaction_types:
            rows.append({
                "Campaign File Name": fn,
                "Scheme Name": scheme_label,
                "Transaction Type": tt,
                "Transaction Count": 0,
                "Total Amount": 0.0,
            })
    return rows
def create_campaign_summary(campaign_df, matched_df):
    if campaign_df.empty:
        return pd.DataFrame()

    campaign_numbers = campaign_df["number"].nunique()
    matched_transactions = len(matched_df)

    if matched_df.empty:
        enrollment = collection = 0
        enrollment_amount = collection_amount = 0
        total_saved = 0
    else:
        is_enroll = matched_df["Transaction Type"].eq(ENROLLMENT)
        is_collect = matched_df["Transaction Type"].eq(COLLECTION)

        enrollment = int(is_enroll.sum())
        collection = int(is_collect.sum())
        enrollment_amount = float(matched_df.loc[is_enroll, "Saved Amount"].sum())
        collection_amount = float(matched_df.loc[is_collect, "Saved Amount"].sum())
        total_saved = float(matched_df["Saved Amount"].sum())

    return pd.DataFrame([{
        "Unique Campaign Numbers": campaign_numbers,
        "Matched Transactions": matched_transactions,
        "Enrollment": enrollment,
        "Enrollment Amount": enrollment_amount,
        "Collection": collection,
        "Collection Amount": collection_amount,
        "Total Saved Amount": total_saved,
    }])
def create_campaign_file_summary(campaign_df, matched_df):
    campaign_counts = (
        campaign_df
        .groupby("Campaign File Name", sort=False, observed=False)["number"]
        .nunique()
        .reset_index(name="Campaign Numbers")
    )

    if matched_df.empty:
        summary = campaign_counts.copy()
        summary["Matched Transactions"] = 0
        summary["Enrollment"] = 0
        summary["Enrollment Amount"] = 0.0
        summary["Collection"] = 0
        summary["Collection Amount"] = 0.0
        summary["Total Saved Amount"] = 0.0
        return summary[[
            "Campaign File Name", "Campaign Numbers",
            "Matched Transactions", "Enrollment", "Enrollment Amount",
            "Collection", "Collection Amount", "Total Saved Amount",
        ]]

    type_summary = (
        matched_df
        .groupby(["Campaign File Name", "Transaction Type"],
                 sort=False, observed=False)
        .agg(
            Transaction_Count=("Id", "size"),
            Amount=("Saved Amount", "sum"),
        )
        .reset_index()
    )

    count_pivot = (
        type_summary.pivot(
            index="Campaign File Name",
            columns="Transaction Type",
            values="Transaction_Count",
        ).fillna(0)
    )
    amount_pivot = (
        type_summary.pivot(
            index="Campaign File Name",
            columns="Transaction Type",
            values="Amount",
        ).fillna(0)
    )

    for col in [ENROLLMENT, COLLECTION]:
        if col not in count_pivot.columns:
            count_pivot[col] = 0
        if col not in amount_pivot.columns:
            amount_pivot[col] = 0.0

    amount_pivot = amount_pivot.rename(columns={
        ENROLLMENT: "Enrollment Amount",
        COLLECTION: "Collection Amount",
    })

    matched_counts = (
        matched_df
        .groupby("Campaign File Name", sort=False, observed=False)
        .size()
        .reset_index(name="Matched Transactions")
    )

    summary = campaign_counts.merge(
        matched_counts, on="Campaign File Name", how="left"
    )
    summary = summary.merge(
        count_pivot.reset_index(), on="Campaign File Name", how="left"
    )
    summary = summary.merge(
        amount_pivot.reset_index(), on="Campaign File Name", how="left"
    )

    zero_columns = [
        "Matched Transactions", ENROLLMENT, "Enrollment Amount",
        COLLECTION, "Collection Amount",
    ]
    for col in zero_columns:
        if col not in summary.columns:
            summary[col] = 0
        summary[col] = summary[col].fillna(0)

    summary["Total Saved Amount"] = (
        summary["Enrollment Amount"] + summary["Collection Amount"]
    )

    for col in ["Campaign Numbers", "Matched Transactions", ENROLLMENT, COLLECTION]:
        summary[col] = summary[col].fillna(0).astype(int)

    return summary[[
        "Campaign File Name", "Campaign Numbers", "Matched Transactions",
        ENROLLMENT, "Enrollment Amount", COLLECTION, "Collection Amount",
        "Total Saved Amount",
    ]].reset_index(drop=True)
def create_campaign_transaction_summary(campaign_df, matched_df):
    campaign_files = (
        campaign_df[["Campaign File Name"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    transaction_types = pd.DataFrame({
        "Transaction Type": [ENROLLMENT, COLLECTION]
    })

    campaign_files["_key"] = 1
    transaction_types["_key"] = 1

    base = (
        campaign_files
        .merge(transaction_types, on="_key", how="inner")
        .drop(columns="_key")
    )

    if matched_df.empty:
        base["Transaction Count"] = 0
        base["Total Amount"] = 0.0
        return base[[
            "Campaign File Name", "Transaction Type",
            "Transaction Count", "Total Amount",
        ]]

    actual = (
        matched_df
        .groupby(["Campaign File Name", "Transaction Type"],
                 sort=False, observed=False)
        .agg(
            Transaction_Count=("Id", "size"),
            Total_Amount=("Saved Amount", "sum"),
        )
        .reset_index()
    )

    actual = actual.rename(columns={
        "Transaction_Count": "Transaction Count",
        "Total_Amount": "Total Amount",
    })

    summary = base.merge(
        actual,
        on=["Campaign File Name", "Transaction Type"],
        how="left"
    )

    summary["Transaction Count"] = (
        summary["Transaction Count"].fillna(0).astype(int)
    )
    summary["Total Amount"] = summary["Total Amount"].fillna(0.0)

    # ✅ FIX #3: guard against Categorical dtype before map+fillna
    summary["_order"] = (
        summary["Transaction Type"]
        .astype(str)
        .map(TYPE_ORDER)
        .fillna(99)
        .astype(int)
    )

    summary = (
        summary
        .sort_values(["Campaign File Name", "_order"], kind="stable")
        .drop(columns="_order")
        .reset_index(drop=True)
    )

    return summary[[
        "Campaign File Name", "Transaction Type",
        "Transaction Count", "Total Amount",
    ]]
def create_campaign_scheme_summary(campaign_df, matched_df):
    campaign_files = (
        campaign_df[["Campaign File Name"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    if matched_df.empty:
        return pd.DataFrame(
            _make_zero_rows(campaign_files["Campaign File Name"].tolist())
        )

    actual = (
        matched_df
        .groupby(["Campaign File Name", "Scheme Name", "Transaction Type"],
                 sort=False, observed=False)
        .agg(
            Transaction_Count=("Id", "size"),
            Total_Amount=("Saved Amount", "sum"),
        )
        .reset_index()
    )

    actual = actual.rename(columns={
        "Transaction_Count": "Transaction Count",
        "Total_Amount": "Total Amount",
    })

    matched_files = set(actual["Campaign File Name"].dropna().unique())
    unmatched_files = campaign_files[
        ~campaign_files["Campaign File Name"].isin(matched_files)
    ]

    zero_rows = _make_zero_rows(
        unmatched_files["Campaign File Name"].tolist()
    )
    if zero_rows:
        actual = pd.concat(
            [actual, pd.DataFrame(zero_rows)],
            ignore_index=True
        )

    # ✅ FIX #2: guard against Categorical dtype before map+fillna
    actual["_type_order"] = (
        actual["Transaction Type"]
        .astype(str)
        .map(TYPE_ORDER)
        .fillna(99)
        .astype(int)
    )

    actual = (
        actual
        .sort_values(
            ["Campaign File Name", "Scheme Name", "_type_order"],
            kind="stable"
        )
        .drop(columns="_type_order")
        .reset_index(drop=True)
    )

    return actual[[
        "Campaign File Name", "Scheme Name", "Transaction Type",
        "Transaction Count", "Total Amount",
    ]]
def create_campaign_phone_summary(campaign_df, matched_df):
    if campaign_df.empty:
        return pd.DataFrame()

    if matched_df.empty:
        phone_summary = (
            campaign_df[["Campaign File Name", "number"]]
            .drop_duplicates()
            .copy()
        )
        phone_summary["Matched Transactions"] = 0
        phone_summary[ENROLLMENT] = 0
        phone_summary["Enrollment Amount"] = 0.0
        phone_summary[COLLECTION] = 0
        phone_summary["Collection Amount"] = 0.0
        phone_summary["Total Saved Amount"] = 0.0

    else:
        summary = (
            matched_df
            .groupby(["Campaign File Name", "number"],
                     sort=False, observed=False)
            .agg(Matched_Transactions=("Id", "size"))
            .reset_index()
        )

        enrollment = (
            matched_df[matched_df["Transaction Type"].eq(ENROLLMENT)]
            .groupby(["Campaign File Name", "number"],
                     sort=False, observed=False)
            .agg(
                Enrollment=("Id", "size"),
                Enrollment_Amount=("Saved Amount", "sum"),
            )
            .reset_index()
        )

        collection = (
            matched_df[matched_df["Transaction Type"].eq(COLLECTION)]
            .groupby(["Campaign File Name", "number"],
                     sort=False, observed=False)
            .agg(
                Collection=("Id", "size"),
                Collection_Amount=("Saved Amount", "sum"),
            )
            .reset_index()
        )

        phone_summary = (
            campaign_df[["Campaign File Name", "number"]]
            .drop_duplicates()
            .merge(summary, on=["Campaign File Name", "number"], how="left")
            .merge(enrollment, on=["Campaign File Name", "number"], how="left")
            .merge(collection, on=["Campaign File Name", "number"], how="left")
        )

        zero_columns = [
            "Matched_Transactions", ENROLLMENT, "Enrollment_Amount",
            COLLECTION, "Collection_Amount",
        ]
        for col in zero_columns:
            if col not in phone_summary.columns:
                phone_summary[col] = 0
            phone_summary[col] = phone_summary[col].fillna(0)

        phone_summary["Total Saved Amount"] = (
            phone_summary["Enrollment_Amount"]
            + phone_summary["Collection_Amount"]
        )

        phone_summary = phone_summary.rename(columns={
            "Matched_Transactions": "Matched Transactions",
            "Enrollment_Amount": "Enrollment Amount",
            "Collection_Amount": "Collection Amount",
        })

    for col in ["Matched Transactions", ENROLLMENT, COLLECTION]:
        phone_summary[col] = phone_summary[col].fillna(0).astype(int)

    for col in ["Enrollment Amount", "Collection Amount", "Total Saved Amount"]:
        phone_summary[col] = phone_summary[col].fillna(0.0)

    return phone_summary[[
        "Campaign File Name", "number", "Matched Transactions",
        ENROLLMENT, "Enrollment Amount", COLLECTION, "Collection Amount",
        "Total Saved Amount",
    ]].rename(columns={"number": "Campaign Number"})
def create_not_matched_summary(campaign_df, matched_df):
    if campaign_df.empty:
        return pd.DataFrame()

    matched_phones = set()
    if not matched_df.empty:
        matched_phones = set(matched_df["number"].dropna().unique())

    unmatched = campaign_df[
        ~campaign_df["number"].isin(matched_phones)
    ].copy()

    if unmatched.empty:
        return pd.DataFrame(columns=[
            "Campaign File Name", "Campaign Number", "status",
            "sent_at", "delivered_at", "read_at",
        ])

    columns = [
        "Campaign File Name", "number", "status",
        "sent_at", "delivered_at", "read_at",
    ]
    available = [c for c in columns if c in unmatched.columns]

    return (
        unmatched[available]
        .rename(columns={"number": "Campaign Number"})
        .reset_index(drop=True)
    )
def _build_formats(workbook):
    return {
        "header": workbook.add_format({
            "bold": True, "font_color": "FFFFFF", "bg_color": "#1F4E78",
            "align": "center", "valign": "vcenter", "border": 1,
        }),
        "text": workbook.add_format({
            "num_format": "@", "border": 1, "valign": "vcenter",
        }),
        "amount": workbook.add_format({
            "num_format": "₹#,##0.00", "border": 1, "valign": "vcenter",
        }),
        "count": workbook.add_format({
            "num_format": "#,##0", "border": 1, "valign": "vcenter",
        }),
        "date": workbook.add_format({
            "num_format": "dd-mm-yyyy hh:mm", "border": 1, "valign": "vcenter",
        }),
        "cell": workbook.add_format({"border": 1, "valign": "vcenter"}),
        "cell_alt": workbook.add_format({
            "border": 1, "valign": "vcenter", "bg_color": "#EAF2F8",
        }),
        "text_alt": workbook.add_format({
            "num_format": "@", "border": 1, "valign": "vcenter",
            "bg_color": "#EAF2F8",
        }),
        "amount_alt": workbook.add_format({
            "num_format": "₹#,##0.00", "border": 1, "valign": "vcenter",
            "bg_color": "#EAF2F8",
        }),
        "count_alt": workbook.add_format({
            "num_format": "#,##0", "border": 1, "valign": "vcenter",
            "bg_color": "#EAF2F8",
        }),
        "date_alt": workbook.add_format({
            "num_format": "dd-mm-yyyy hh:mm", "border": 1,
            "valign": "vcenter", "bg_color": "#EAF2F8",
        }),
    }
def _pick_format_key(header, alt=False):
    header_str = str(header)
    if header_str in PHONE_HEADERS:
        base = "text"
    elif "Amount" in header_str or header_str == "Total Saved Amount":
        base = "amount"
    elif header_str in COUNT_HEADERS:
        base = "count"
    elif ("Date" in header_str or header_str.endswith("_at")
          or header_str in DATE_HEADERS):
        base = "date"
    else:
        base = "cell"
    if alt and base != "header":
        return f"{base}_alt"
    return base
def _compute_col_widths(df):
    widths = {}
    for col in df.columns:
        try:
            header_len = len(str(col))
            max_data = df[col].head(500).astype(str).str.len().max()
            if pd.isna(max_data):
                max_data = 0
            widths[col] = min(max(int(max_data), header_len) + 2, 40)
        except Exception:
            widths[col] = 15
    return widths
def _to_python_scalar(value):
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return float(value)
    if isinstance(value, np.bool_):
        return bool(value)
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value
def _is_missing(value):
    if value is None:
        return True
    try:
        result = pd.isna(value)
        if isinstance(result, (bool, np.bool_)):
            return bool(result)
        return False
    except (TypeError, ValueError):
        return False
def _extract_campaign_date_from_filename(file_name: str) -> str:
    name = Path(str(file_name)).stem
    name_clean = name.replace("_", " ").replace("-", " ")

    month_names = (
        "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|"
        "january|february|march|april|june|july|august|september|"
        "october|november|december"
    )

    pat1 = re.search(
        rf"\b({month_names})\s+(\d{{1,2}})(?:[,\s]+(\d{{4}}))?\b",
        name_clean, re.IGNORECASE
    )
    if pat1:
        mon = pat1.group(1).capitalize()[:3]
        day = pat1.group(2)
        yr = pat1.group(3) or ""
        return f"{mon} {day} {yr}".strip()

    pat2 = re.search(
        rf"\b(\d{{1,2}})\s+({month_names})(?:[,\s]+(\d{{4}}))?\b",
        name_clean, re.IGNORECASE
    )
    if pat2:
        day = pat2.group(1)
        mon = pat2.group(2).capitalize()[:3]
        yr = pat2.group(3) or ""
        return f"{mon} {day} {yr}".strip()

    pat3 = re.search(r"\b(\d{1,2})[\s/-]+(\d{1,2})[\s/-]+(\d{2,4})\b", name_clean)
    if pat3:
        return f"{pat3.group(1)}-{pat3.group(2)}-{pat3.group(3)}"

    return ""
def _sanitize_sheet_name(name: str, used: set) -> str:
    safe = re.sub(r'[\[\]\:\*\?\/\\]', "_", str(name)).strip()
    safe = safe[:31] or "Sheet"
    base = safe
    i = 1
    while safe.lower() in used:
        suffix = f"_{i}"
        safe = (base[: 31 - len(suffix)]) + suffix
        i += 1
    used.add(safe.lower())
    return safe
def build_per_campaign_matrix(campaign_df, matched_df, scheme_order=None):
    """
    Build per-campaign blocks for EVERY campaign file in campaign_df,
    even those with ZERO matched transactions.
    """
    result = {}

    if campaign_df is None or campaign_df.empty:
        return result

    # ── Determine the full scheme list from ALL matched data ──
    all_schemes = []
    if matched_df is not None and not matched_df.empty:
        all_schemes = (
            matched_df["Scheme Name"].dropna().astype(str)
            .replace("", np.nan).dropna().unique().tolist()
        )

    ordered = [s for s in PREFERRED_SCHEME_ORDER if s in all_schemes]
    extras = sorted([s for s in all_schemes if s not in ordered])
    scheme_list = ordered + extras

    if scheme_order:
        scheme_list = [s for s in scheme_order if s in scheme_list] + [
            s for s in scheme_list if s not in scheme_order
        ]

    if not scheme_list:
        scheme_list = [NO_MATCH]

    # ── Iterate over EVERY campaign file in campaign_df ──
    campaign_files = (
        campaign_df["Campaign File Name"].astype(str).dropna().unique().tolist()
    )

    for campaign_name in campaign_files:

        if matched_df is not None and not matched_df.empty:
            sub = matched_df[
                matched_df["Campaign File Name"].astype(str) == campaign_name
            ]
        else:
            sub = pd.DataFrame()

        has_match = not sub.empty

        rows_data = {}

        for txn_type in [ENROLLMENT, COLLECTION]:
            if sub.empty:
                sub_t = pd.DataFrame()
            else:
                sub_t = sub[sub["Transaction Type"].astype(str) == txn_type]

            row = {}
            total_count = 0
            total_amount = 0.0

            for scheme in scheme_list:
                if sub_t.empty:
                    cnt = 0
                    amt = 0.0
                else:
                    sub_ts = sub_t[sub_t["Scheme Name"] == scheme]
                    cnt = int(len(sub_ts))
                    amt = float(sub_ts["Saved Amount"].sum()) if not sub_ts.empty else 0.0

                row[f"{scheme} Count"] = cnt
                row[f"{scheme} Amount"] = amt
                total_count += cnt
                total_amount += amt

            row["Total Count"] = total_count
            row["Total Amount"] = total_amount
            rows_data[txn_type] = row

        total_row = {}
        grand_count = 0
        grand_amount = 0.0
        for scheme in scheme_list:
            c = rows_data[ENROLLMENT].get(f"{scheme} Count", 0) + \
                rows_data[COLLECTION].get(f"{scheme} Count", 0)
            a = rows_data[ENROLLMENT].get(f"{scheme} Amount", 0.0) + \
                rows_data[COLLECTION].get(f"{scheme} Amount", 0.0)
            total_row[f"{scheme} Count"] = c
            total_row[f"{scheme} Amount"] = a
            grand_count += c
            grand_amount += a

        total_row["Total Count"] = grand_count
        total_row["Total Amount"] = grand_amount
        rows_data["Total"] = total_row

        rows_df = pd.DataFrame([
            {"Transaction Type": ENROLLMENT, **rows_data[ENROLLMENT]},
            {"Transaction Type": COLLECTION, **rows_data[COLLECTION]},
            {"Transaction Type": "Total", **rows_data["Total"]},
        ])

        result[campaign_name] = {
            "campaign_name": campaign_name,
            "campaign_date": _extract_campaign_date_from_filename(campaign_name),
            "schemes": scheme_list,
            "rows": rows_df,
            "has_match": has_match,
        }

    return result
def _write_campaign_block(workbook, ws, block):
    """
    Write the stacked layout with enhanced formatting.
    """
    campaign_name = block["campaign_name"]
    campaign_date = block["campaign_date"]
    schemes = block["schemes"]
    rows_df = block["rows"]

    enrollment_row = rows_df.iloc[0].to_dict()
    collection_row = rows_df.iloc[1].to_dict()

    # ────────────────────────────────────────────────────────
    # FORMATS
    # ────────────────────────────────────────────────────────

    fmt_title = workbook.add_format({
        "bold": True, "font_size": 14,
        "font_color": "FFFFFF", "bg_color": "#C55A11",
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#8B3A0A",
    })

    fmt_date = workbook.add_format({
        "italic": True, "font_size": 11,
        "font_color": "FFFFFF", "bg_color": "#E67E22",
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#8B3A0A",
    })

    fmt_section_label = workbook.add_format({
        "bold": True, "font_size": 12,
        "font_color": "0F2027", "bg_color": "#B4C7DC",
        "align": "left", "valign": "vcenter",
        "border": 1, "border_color": "#7A90A8",
        "text_wrap": True,
    })

    scheme_palette = [
        {"bg": "#B4A7D6", "fg": "0F2027"},
        {"bg": "#C9C9C9", "fg": "0F2027"},
        {"bg": "#D9C7E8", "fg": "0F2027"},
        {"bg": "#E8F0D8", "fg": "0F2027"},
        {"bg": "#D6EAF8", "fg": "0F2027"},
        {"bg": "#F5D7B7", "fg": "0F2027"},
    ]

    def make_scheme_fmt(idx):
        c = scheme_palette[idx % len(scheme_palette)]
        return workbook.add_format({
            "bold": True, "font_size": 11,
            "font_color": c["fg"], "bg_color": c["bg"],
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": "#7A90A8",
        })

    scheme_formats = [make_scheme_fmt(i) for i in range(len(schemes))]

    fmt_total_header = workbook.add_format({
        "bold": True, "font_size": 11,
        "font_color": "0F2027", "bg_color": "#9CC3D5",
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#5E7C93",
    })

    fmt_sub = workbook.add_format({
        "bold": True, "font_size": 10,
        "font_color": "0F2027", "bg_color": "#F5DDC8",
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#B99A82",
    })

    fmt_count = workbook.add_format({
        "num_format": "#,##0", "border": 1,
        "align": "right", "valign": "vcenter",
        "border_color": "#B99A82",
    })
    fmt_amount = workbook.add_format({
        "num_format": "₹#,##0.00", "border": 1,
        "align": "right", "valign": "vcenter",
        "border_color": "#B99A82",
    })

    fmt_total_count = workbook.add_format({
        "num_format": "#,##0", "border": 1, "bold": True,
        "bg_color": "#FFF2CC", "align": "right", "valign": "vcenter",
        "border_color": "#B99A82",
    })
    fmt_total_amount = workbook.add_format({
        "num_format": "₹#,##0.00", "border": 1, "bold": True,
        "bg_color": "#FFF2CC", "align": "right", "valign": "vcenter",
        "border_color": "#B99A82",
    })

    fmt_empty_cell = workbook.add_format({
        "bg_color": "#B4C7DC",
        "border": 1, "border_color": "#7A90A8",
    })

    # ────────────────────────────────────────────────────────
    # LAYOUT
    # ────────────────────────────────────────────────────────

    n_schemes = len(schemes)
    n_cols = 1 + 2 * n_schemes + 2

    ws.merge_range(0, 0, 0, n_cols - 1, campaign_name, fmt_title)
    date_text = f"Date: {campaign_date}" if campaign_date else ""
    ws.merge_range(1, 0, 1, n_cols - 1, date_text, fmt_date)

    def _write_section(top_row, section_label, data_row):
        mid = top_row + 1
        data_r = top_row + 2

        ws.merge_range(top_row, 0, mid, 0, section_label, fmt_section_label)

        c = 1
        for idx, scheme in enumerate(schemes):
            fmt_s = scheme_formats[idx]
            ws.merge_range(top_row, c, top_row, c + 1, scheme, fmt_s)
            c += 2

        ws.merge_range(top_row, c, top_row, c + 1, "Total", fmt_total_header)

        c = 1
        for _ in schemes:
            ws.write(mid, c, "Count", fmt_sub)
            ws.write(mid, c + 1, "Amount", fmt_sub)
            c += 2
        ws.write(mid, c, "Count", fmt_sub)
        ws.write(mid, c + 1, "Amount", fmt_sub)

        ws.write(data_r, 0, "", fmt_empty_cell)

        c = 1
        for scheme in schemes:
            cnt_val = data_row.get(f"{scheme} Count", 0)
            amt_val = data_row.get(f"{scheme} Amount", 0.0)

            try:
                ws.write_number(data_r, c, float(cnt_val), fmt_count)
            except Exception:
                ws.write(data_r, c, cnt_val, fmt_count)

            try:
                ws.write_number(data_r, c + 1, float(amt_val), fmt_amount)
            except Exception:
                ws.write(data_r, c + 1, amt_val, fmt_amount)

            c += 2

        tot_cnt = data_row.get("Total Count", 0)
        tot_amt = data_row.get("Total Amount", 0.0)

        try:
            ws.write_number(data_r, c, float(tot_cnt), fmt_total_count)
        except Exception:
            ws.write(data_r, c, tot_cnt, fmt_total_count)

        try:
            ws.write_number(data_r, c + 1, float(tot_amt), fmt_total_amount)
        except Exception:
            ws.write(data_r, c + 1, tot_amt, fmt_total_amount)

    enrollment_top = 3
    _write_section(enrollment_top, "Enrollment", enrollment_row)

    collection_top = enrollment_top + 3 + 2
    _write_section(collection_top, "Subsequent Collection", collection_row)

    ws.set_row(0, 24)
    ws.set_row(1, 20)
    ws.set_row(enrollment_top, 20)
    ws.set_row(enrollment_top + 1, 18)
    ws.set_row(collection_top, 20)
    ws.set_row(collection_top + 1, 18)

    ws.set_column(0, 0, 22)
    for c in range(1, n_cols):
        ws.set_column(c, c, 12)

    ws.freeze_panes(3, 1)
def create_single_campaign_excel(campaign_block):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet_name = _sanitize_sheet_name(
            campaign_block["campaign_name"], set()
        )
        ws = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws

        _write_campaign_block(workbook, ws, campaign_block)

    output.seek(0)
    return output
def create_excel_file(
    final_df,
    campaign_summary,
    campaign_file_summary,
    campaign_transaction_summary,
    campaign_scheme_summary,
    campaign_phone_summary,
    not_matched_summary,
    matched_df=None,
    campaign_df_for_sheets=None,
):
    output = BytesIO()

    sheets = [
        ("Final Data", final_df),
        ("Campaign Summary", campaign_summary),
        ("Campaign File Summary", campaign_file_summary),
        ("Campaign Transaction Summary", campaign_transaction_summary),
        ("Campaign Scheme Summary", campaign_scheme_summary),
        ("Campaign Phone Summary", campaign_phone_summary),
        ("Not Matched Phones", not_matched_summary),
    ]

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs={"options": {"constant_memory": True}},
    ) as writer:

        workbook = writer.book
        formats = _build_formats(workbook)

        for sheet_name, df in sheets:

            ws = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws
            ws.freeze_panes(1, 0)
            ws.set_tab_color(TAB_COLORS.get(sheet_name, "#1F4E78"))

            if df is None or len(df.columns) == 0:
                continue

            widths = _compute_col_widths(df)
            for col_idx, col_name in enumerate(df.columns):
                ws.set_column(col_idx, col_idx, widths.get(col_name, 15))

            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, str(col_name), formats["header"])

            n_cols = len(df.columns)
            fmt_keys_normal = [
                _pick_format_key(df.columns[c], alt=False)
                for c in range(n_cols)
            ]
            fmt_keys_alt = [
                _pick_format_key(df.columns[c], alt=True)
                for c in range(n_cols)
            ]

            df_reset = df.reset_index(drop=True)

            for row_idx in range(len(df_reset)):
                alt = (row_idx % 2 == 1)
                fmt_keys = fmt_keys_alt if alt else fmt_keys_normal
                for col_idx in range(n_cols):
                    value = df_reset.iat[row_idx, col_idx]
                    fmt = formats[fmt_keys[col_idx]]
                    if _is_missing(value):
                        ws.write_blank(row_idx + 1, col_idx, None, fmt)
                        continue
                    value = _to_python_scalar(value)
                    ws.write(row_idx + 1, col_idx, value, fmt)

            if len(df_reset) > 0:
                ws.autofilter(0, 0, len(df_reset), n_cols - 1)

        if campaign_df_for_sheets is not None and not campaign_df_for_sheets.empty:
            per_campaign = build_per_campaign_matrix(
                campaign_df_for_sheets, matched_df
            )
            used_names = set()
            for campaign_name, block in per_campaign.items():
                sheet_name = _sanitize_sheet_name(campaign_name, used_names)
                ws = workbook.add_worksheet(sheet_name)
                ws.set_tab_color("#F4B183")
                _write_campaign_block(workbook, ws, block)

    output.seek(0)
    return output
def create_zip_bundle(single_campaign_files: dict) -> BytesIO:
    """Bundle all per-campaign xlsx files into one ZIP."""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for campaign_name, info in single_campaign_files.items():
            zf.writestr(info["filename"], info["bytes"])
    zip_buffer.seek(0)
    return zip_buffer

LABEL = "Campaign Transaction Analyzer"


def run():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        .main .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }

        /* ---- HERO HEADER ---- */
        .hero-header {
            background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
            padding: 2rem 2.5rem;
            border-radius: 18px;
            margin-bottom: 1.5rem;
            box-shadow: 0 10px 30px rgba(0,0,0,0.25);
            color: #fff;
        }
        .hero-header h1 {
            margin: 0;
            font-size: 2.1rem;
            font-weight: 800;
            letter-spacing: -0.5px;
            color: #fff;
        }
        .hero-header p {
            margin: 0.5rem 0 0 0;
            font-size: 1rem;
            opacity: 0.85;
            color: #e0f2fe;
        }

        /* ---- SECTION HEADERS ---- */
        .section-title {
            font-size: 1.25rem;
            font-weight: 700;
            color: #0f2027;
            margin: 0.5rem 0 0.75rem 0;
            padding-left: 0.75rem;
            border-left: 5px solid #2c5364;
            line-height: 1.3;
        }

        /* ---- METRIC CARDS ---- */
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, #ffffff 0%, #f0f9ff 100%);
            border: 1px solid #dbeafe;
            border-radius: 14px;
            padding: 1rem 1.25rem;
            box-shadow: 0 4px 12px rgba(44, 83, 100, 0.08);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        div[data-testid="stMetric"]:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 20px rgba(44, 83, 100, 0.18);
        }
        div[data-testid="stMetric"] label {
            color: #64748b !important;
            font-weight: 600 !important;
            font-size: 0.8rem !important;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            color: #0f2027 !important;
            font-weight: 800 !important;
            font-size: 1.6rem !important;
        }

        /* ---- FILE UPLOADERS ---- */
        div[data-testid="stFileUploader"] {
            background: #f8fafc;
            border: 2px dashed #94a3b8;
            border-radius: 14px;
            padding: 0.75rem;
            transition: border-color 0.2s ease, background 0.2s ease;
        }
        div[data-testid="stFileUploader"]:hover {
            border-color: #2c5364;
            background: #f0f9ff;
        }

        /* ---- BUTTONS ---- */
        div.stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #0f2027 0%, #2c5364 100%);
            color: #fff;
            border: none;
            border-radius: 12px;
            padding: 0.75rem 1.5rem;
            font-weight: 700;
            font-size: 1rem;
            letter-spacing: 0.3px;
            box-shadow: 0 6px 18px rgba(15, 32, 39, 0.3);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        div.stButton > button[kind="primary"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 25px rgba(15, 32, 39, 0.45);
        }

        div.stDownloadButton > button {
            background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
            color: #fff;
            border: none;
            border-radius: 10px;
            font-weight: 600;
            padding: 0.55rem 1rem;
            transition: transform 0.15s ease, box-shadow 0.15s ease;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
        }
        div.stDownloadButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(37, 99, 235, 0.4);
        }

        /* ---- DATAFRAMES ---- */
        div[data-testid="stDataFrame"] {
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.04);
        }

        /* ---- TABS ---- */
        button[data-baseweb="tab"] {
            font-weight: 600;
            font-size: 0.95rem;
            border-radius: 10px 10px 0 0;
        }

        /* ---- DIVIDER ---- */
        hr {
            margin: 1.25rem 0;
            border: none;
            border-top: 1px solid #e2e8f0;
        }

        /* ---- DOWNLOAD CARD ---- */
        .download-card {
            background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
            border: 1px solid #bae6fd;
            border-radius: 14px;
            padding: 1rem 1.25rem;
            margin-bottom: 0.75rem;
            transition: box-shadow 0.2s ease;
        }
        .download-card:hover {
            box-shadow: 0 6px 18px rgba(37, 99, 235, 0.15);
        }
        .download-card.no-match {
            background: linear-gradient(135deg, #fff7ed 0%, #ffedd5 100%);
            border-color: #fdba74;
        }

        /* ---- SIDEBAR ---- */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0f2027 0%, #203a43 100%);
        }
        section[data-testid="stSidebar"] * {
            color: #e2e8f0 !important;
        }
        section[data-testid="stSidebar"] .stCheckbox label span {
            color: #e2e8f0 !important;
        }
        section[data-testid="stSidebar"] hr {
            border-top: 1px solid rgba(255,255,255,0.12);
        }

        /* ---- RADIO AS TABS (persistent navigation) ---- */
        div[role="radiogroup"] {
            flex-direction: row;
            gap: 0.5rem;
            border-bottom: 2px solid #e2e8f0;
            padding-bottom: 0;
            margin-bottom: 1rem;
        }
        div[role="radiogroup"] > label {
            background: #f1f5f9;
            border-radius: 10px 10px 0 0;
            padding: 0.6rem 1.2rem;
            font-weight: 600;
            font-size: 0.95rem;
            color: #475569;
            cursor: pointer;
            transition: all 0.15s ease;
            border: 1px solid transparent;
            border-bottom: none;
            margin-bottom: -2px;
        }
        div[role="radiogroup"] > label:hover {
            background: #e2e8f0;
            color: #0f2027;
        }
        div[role="radiogroup"] > label[data-checked="true"] {
            background: linear-gradient(135deg, #0f2027 0%, #2c5364 100%);
            color: #fff !important;
            box-shadow: 0 4px 12px rgba(15, 32, 39, 0.25);
        }
        div[role="radiogroup"] > label[data-checked="true"] p,
        div[role="radiogroup"] > label[data-checked="true"] span {
            color: #fff !important;
        }
        /* Hide the radio dot circles */
        div[role="radiogroup"] > label > div:first-child {
            display: none;
        }
    </style>
    """, unsafe_allow_html=True)
    ENROLLMENT = "Enrollment"
    COLLECTION = "Collection"
    MATCHED = "Matched"
    NOT_MATCHED = "Not Matched"
    NO_MATCH = "No Match"
    STATUS_SENT = "sent"
    STATUS_DELIVERED = "delivered"
    STATUS_READ = "read"
    ALLOWED_CAMPAIGN_STATUSES = (STATUS_SENT, STATUS_DELIVERED, STATUS_READ)
    STATUS_PRIORITY = {
        STATUS_SENT: 1,
        STATUS_DELIVERED: 2,
        STATUS_READ: 3,
    }
    TYPE_ORDER = {
        ENROLLMENT: 1,
        COLLECTION: 2,
    }
    PREFERRED_SCHEME_ORDER = [
        "e-Gold", "e-Silver", "Akshaya Tritiya",
        "Christmas", "Diwali", "Pongal",
    ]
    CAMPAIGN_REQUIRED_COLUMNS = [
        "number", "status", "sent_at", "delivered_at", "read_at",
        "failed_at", "error_code", "campName", "TemplateName",
    ]
    TRANSACTION_REQUIRED_COLUMNS = [
        "Id", "Scheme Participation Id", "Date", "Status", "Saved Amount",
        "Refunded Amount", "Reward Amount", "Transaction Reference",
        "Installment number", "Metal Type", "Metal Rate",
        "Saved Metal Weight", "Rewards Metal Weight", "Benefit Metal Amount",
        "Benefit Metal Weight", "Benefit Metal Percentage", "Receipt ID",
        "Customer Name", "Customer Phone Number", "Passbook number",
        "Scheme Name",
    ]
    CAMPAIGN_COLUMN_ALIASES = {
        "phone": "number", "phone number": "number", "mobile": "number",
        "mobile number": "number", "whatsapp number": "number",
        "contact": "number", "contact number": "number",
        "sent": "sent_at", "sent at": "sent_at", "sent time": "sent_at",
        "delivered": "delivered_at", "delivered at": "delivered_at",
        "delivered time": "delivered_at",
        "read": "read_at", "read at": "read_at", "read time": "read_at",
        "failed": "failed_at", "failed at": "failed_at",
        "failed time": "failed_at",
        "error code": "error_code", "errorcode": "error_code",
        "error": "error_code",
        "camp name": "campName", "campaign name": "campName",
        "campname": "campName", "campaign": "campName",
        "template name": "TemplateName", "templatename": "TemplateName",
        "template": "TemplateName",
    }
    TRANSACTION_COLUMN_ALIASES = {
        "customer phone": "Customer Phone Number",
        "customer phone number": "Customer Phone Number",
        "customer mobile": "Customer Phone Number",
        "phone": "Customer Phone Number",
        "phone number": "Customer Phone Number",
        "mobile": "Customer Phone Number",
        "mobile number": "Customer Phone Number",
        "installment no": "Installment number",
        "installment no.": "Installment number",
        "installment": "Installment number",
        "installment_number": "Installment number",
        "id": "Id",
        "scheme participation id": "Scheme Participation Id",
        "date": "Date", "status": "Status",
        "saved amount": "Saved Amount",
        "refunded amount": "Refunded Amount",
        "reward amount": "Reward Amount",
        "transaction reference": "Transaction Reference",
        "metal type": "Metal Type", "metal rate": "Metal Rate",
        "saved metal weight": "Saved Metal Weight",
        "rewards metal weight": "Rewards Metal Weight",
        "benefit metal amount": "Benefit Metal Amount",
        "benefit metal weight": "Benefit Metal Weight",
        "benefit metal percentage": "Benefit Metal Percentage",
        "receipt id": "Receipt ID",
        "customer name": "Customer Name",
        "passbook number": "Passbook number",
        "passbook no": "Passbook number",
        "scheme name": "Scheme Name",
    }
    PHONE_HEADERS = {"number", "Campaign Number", "Customer Phone Number"}
    COUNT_HEADERS = {
        "Campaign Numbers", "Unique Campaign Numbers",
        "Matched Transactions", "Transaction Count",
        "Enrollment", "Collection",
    }
    DATE_HEADERS = {"sent_at", "delivered_at", "read_at", "failed_at"}
    TAB_COLORS = {
        "Final Data": "#1F4E78",
        "Campaign Summary": "#70AD47",
        "Campaign File Summary": "#8064A2",
        "Campaign Transaction Summary": "#ED7D31",
        "Campaign Scheme Summary": "#5B9BD5",
        "Campaign Phone Summary": "#A64D79",
        "Not Matched Phones": "#C00000",
    }
    st.markdown("""
    <div class="hero-header">
        <h1>📊 Campaign Transaction Analyzer</h1>
        <p>Upload your campaign files and a transaction file to generate a comprehensive analysis with matched transactions, scheme breakdowns, and downloadable Excel reports.</p>
    </div>
    """, unsafe_allow_html=True)
    with st.sidebar:
        st.markdown("## ⚙️ Settings")
        st.markdown("---")

        debug_mode = st.checkbox(
            "🔍 Debug mode",
            value=False,
            help="Show extra diagnostics and raw data shapes during processing."
        )

        st.markdown("---")
        st.markdown("### 📖 How it works")
        st.markdown("""
        1. **Upload** campaign files (Excel/CSV) — first 2 sheets are combined.
        2. **Upload** one transaction file.
        3. Click **Process & Generate Excel**.
        4. Filter: only `Sent`, `Delivered`, `Read` statuses.
        5. Deduplicate phone numbers globally.
        6. Match campaign numbers with transaction phone numbers.
        7. Split into **Enrollment** (installment 1) and **Collection** (installment > 1).
        8. Download combined report + individual campaign sheets.
        """)

        st.markdown("---")
        st.caption("Built with ❤️ using Streamlit")
    st.markdown('<div class="section-title">📁 Upload Files</div>', unsafe_allow_html=True)
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        campaign_files = st.file_uploader(
            "📁 Campaign files (Excel/CSV, multiple allowed)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="campaign_uploader"
        )
    with col_up2:
        transaction_file = st.file_uploader(
            "💳 Transaction file (Excel/CSV)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=False,
            key="transaction_uploader"
        )
    if "excel_bytes" not in st.session_state:
        st.session_state.excel_bytes = None
    if "excel_filename" not in st.session_state:
        st.session_state.excel_filename = None
    if "results" not in st.session_state:
        st.session_state.results = None
    if "single_campaign_files" not in st.session_state:
        st.session_state.single_campaign_files = None
    if "zip_bytes" not in st.session_state:
        st.session_state.zip_bytes = None
    if "active_tab" not in st.session_state:
        st.session_state.active_tab = "📈 Overview"
    if campaign_files and transaction_file:

        st.markdown("---")

        process_button = st.button(
            "🚀 Process & Generate Excel",
            type="primary",
            width="stretch"
        )

        if process_button:

            st.session_state.excel_bytes = None
            st.session_state.excel_filename = None
            st.session_state.results = None
            st.session_state.single_campaign_files = None
            st.session_state.zip_bytes = None

            try:

                t_start = time.perf_counter()

                # READ CAMPAIGNS
                t0 = time.perf_counter()
                campaign_raw = read_campaign_files(campaign_files)
                t_read_campaign = time.perf_counter() - t0

                if debug_mode:
                    st.write("**Raw Campaign shape:**", campaign_raw.shape)

                if campaign_raw.empty:
                    st.error("❌ No Campaign data was found.")
                    st.stop()

                # PROCESS CAMPAIGNS
                t0 = time.perf_counter()
                with st.spinner("⚙️ Processing Campaign data..."):
                    campaign_df = process_campaign_data(campaign_raw)
                t_proc_campaign = time.perf_counter() - t0

                if campaign_df.empty:
                    st.error("❌ No valid Campaign records were found after filtering.")
                    st.stop()

                # READ TRANSACTION
                t0 = time.perf_counter()
                with st.spinner("📥 Reading Transaction file..."):
                    transaction_raw = read_transaction_file(transaction_file)
                t_read_txn = time.perf_counter() - t0

                # PROCESS TRANSACTIONS
                t0 = time.perf_counter()
                with st.spinner("⚙️ Processing Transaction data..."):
                    transaction_df = process_transaction_data(transaction_raw)
                t_proc_txn = time.perf_counter() - t0

                # MATCH
                t0 = time.perf_counter()
                with st.spinner("🔗 Matching Campaign numbers..."):
                    final_df = create_final_data(campaign_df, transaction_df)
                t_match = time.perf_counter() - t0

                matched_df = get_matched_transactions(final_df)

                # SUMMARIES
                t0 = time.perf_counter()
                with st.spinner("📊 Creating summaries..."):

                    campaign_summary = create_campaign_summary(campaign_df, matched_df)
                    campaign_file_summary = create_campaign_file_summary(campaign_df, matched_df)
                    campaign_transaction_summary = create_campaign_transaction_summary(campaign_df, matched_df)
                    campaign_scheme_summary = create_campaign_scheme_summary(campaign_df, matched_df)
                    campaign_phone_summary = create_campaign_phone_summary(campaign_df, matched_df)
                    not_matched_summary = create_not_matched_summary(campaign_df, matched_df)
                t_summaries = time.perf_counter() - t0

                # PER-CAMPAIGN SINGLE-SHEET EXCEL FILES
                t0 = time.perf_counter()
                with st.spinner("📗 Building per-campaign Excel files..."):
                    per_campaign_blocks = build_per_campaign_matrix(
                        campaign_df, matched_df
                    )

                    single_campaign_files = {}
                    for campaign_name, block in per_campaign_blocks.items():
                        bytes_io = create_single_campaign_excel(block)
                        safe_name = re.sub(r'[\[\]\:\*\?\/\\]', "_", str(campaign_name))
                        safe_name = Path(safe_name).stem or "campaign"
                        single_campaign_files[campaign_name] = {
                            "filename": f"{safe_name}.xlsx",
                            "bytes": bytes_io.getvalue(),
                            "date": block["campaign_date"],
                            "scheme_count": len(block["schemes"]),
                            "row_count": len(block["rows"]),
                            "has_match": block.get("has_match", True),
                        }
                t_per_campaign = time.perf_counter() - t0

                # ZIP BUNDLE
                t0 = time.perf_counter()
                zip_bytes = None
                if single_campaign_files:
                    zip_buffer = create_zip_bundle(single_campaign_files)
                    zip_bytes = zip_buffer.getvalue()
                t_zip = time.perf_counter() - t0

                # EXCEL (MAIN WORKBOOK)
                t0 = time.perf_counter()
                with st.spinner("📗 Creating combined Excel workbook..."):
                    excel_file = create_excel_file(
                        final_df,
                        campaign_summary,
                        campaign_file_summary,
                        campaign_transaction_summary,
                        campaign_scheme_summary,
                        campaign_phone_summary,
                        not_matched_summary,
                        matched_df=matched_df,
                        campaign_df_for_sheets=campaign_df,
                    )
                t_excel = time.perf_counter() - t0

                t_total = time.perf_counter() - t_start

                st.session_state.excel_bytes = excel_file.getvalue()
                st.session_state.excel_filename = "Campaign_Transaction_Report.xlsx"
                st.session_state.single_campaign_files = single_campaign_files
                st.session_state.zip_bytes = zip_bytes
                st.session_state.results = {
                    "campaign_df": campaign_df,
                    "final_df": final_df,
                    "matched_df": matched_df,
                    "campaign_file_summary": campaign_file_summary,
                    "campaign_transaction_summary": campaign_transaction_summary,
                    "campaign_scheme_summary": campaign_scheme_summary,
                    "not_matched_summary": not_matched_summary,
                    "timings": {
                        "read_campaign": t_read_campaign,
                        "process_campaign": t_proc_campaign,
                        "read_transaction": t_read_txn,
                        "process_transaction": t_proc_txn,
                        "match": t_match,
                        "summaries": t_summaries,
                        "per_campaign": t_per_campaign,
                        "single_campaign": t_per_campaign,
                        "zip": t_zip,
                        "excel": t_excel,
                        "total": t_total,
                    },
                }

                # 👇 Auto-switch to Overview tab after successful processing
                st.session_state.active_tab = "📈 Overview"

                st.success("✅ Processing completed successfully!")

            except Exception as e:

                st.error(f"❌ Error while processing files: {type(e).__name__}: {e}")

                with st.expander("🔍 Full traceback", expanded=True):
                    st.code(traceback.format_exc(), language="python")

                st.info(
                    "💡 Common causes:\n"
                    "- Missing/misnamed column\n"
                    "- Status column has values other than `sent`/`delivered`/`read`\n"
                    "- Empty campaign data after filtering\n"
                    "- Password-protected or merged-header Excel file"
                )
    if st.session_state.results:

        r = st.session_state.results

        st.markdown("---")

        # ---- Persistent tab controller (survives reruns/downloads) ----
        TAB_OPTIONS = ["📈 Overview", "📊 Summaries", "⬇️ Downloads"]

        # Guard: reset if stored value is invalid
        if st.session_state.active_tab not in TAB_OPTIONS:
            st.session_state.active_tab = TAB_OPTIONS[0]

        selected_tab = st.radio(
            "Navigation",
            TAB_OPTIONS,
            index=TAB_OPTIONS.index(st.session_state.active_tab),
            horizontal=True,
            key="active_tab",
            label_visibility="collapsed",
        )

        # ========================================================
        # TAB 1: OVERVIEW
        # ========================================================
        if selected_tab == TAB_OPTIONS[0]:

            campaign_df = r["campaign_df"]
            final_df = r["final_df"]
            matched_df = r["matched_df"]

            st.markdown('<div class="section-title">📈 Key Metrics</div>', unsafe_allow_html=True)

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric(
                    "🎯 Unique Campaign Numbers",
                    f"{campaign_df['number'].nunique():,}"
                )
            with col2:
                matched_count = (
                    final_df[final_df["Transaction Match"].eq(MATCHED)]["number"]
                    .nunique()
                )
                st.metric("✅ Matched Numbers", f"{matched_count:,}")
            with col3:
                st.metric("🔗 Matched Transactions", f"{len(matched_df):,}")
            with col4:
                st.metric(
                    "📁 Campaign Files",
                    f"{campaign_df['Campaign File Name'].nunique():,}"
                )

            st.markdown("---")

            st.markdown('<div class="section-title">⏱️ Processing Timings</div>', unsafe_allow_html=True)
            timings = r["timings"]

            def _t(key):
                return timings.get(key, 0.0)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Read Campaign", f"{_t('read_campaign'):.2f}s")
            c2.metric("Read Transaction", f"{_t('read_transaction'):.2f}s")
            c3.metric("Summaries", f"{_t('summaries'):.2f}s")
            c4.metric("Per-Campaign", f"{_t('per_campaign'):.2f}s")
            c5.metric("Excel Write", f"{_t('excel'):.2f}s")

            st.caption(
                f"Total: **{_t('total'):.2f}s** — "
                f"Match: {_t('match'):.2f}s — "
                f"Process Campaign: {_t('process_campaign'):.2f}s — "
                f"Process Transaction: {_t('process_transaction'):.2f}s — "
                f"Single Campaign Files: {_t('single_campaign'):.2f}s — "
                f"ZIP: {_t('zip'):.2f}s"
            )

            st.markdown("---")

            st.markdown('<div class="section-title">📁 Campaign File Summary — Enrollment vs Collection</div>', unsafe_allow_html=True)
            st.dataframe(
                r["campaign_file_summary"],
                width="stretch",
                hide_index=True
            )

        # ========================================================
        # TAB 2: SUMMARIES
        # ========================================================
        elif selected_tab == TAB_OPTIONS[1]:

            st.markdown('<div class="section-title">💳 Campaign Transaction Summary</div>', unsafe_allow_html=True)
            st.dataframe(
                r["campaign_transaction_summary"],
                width="stretch",
                hide_index=True
            )

            st.markdown("---")

            st.markdown('<div class="section-title">📋 Campaign Scheme Summary — Breakdown by Scheme</div>', unsafe_allow_html=True)
            st.dataframe(
                r["campaign_scheme_summary"],
                width="stretch",
                hide_index=True
            )

            st.markdown("---")

            st.markdown('<div class="section-title">🚫 Not Matched Phones</div>', unsafe_allow_html=True)
            if r["not_matched_summary"].empty:
                st.success("🎉 Every campaign number matched at least one transaction.")
            else:
                st.dataframe(
                    r["not_matched_summary"],
                    width="stretch",
                    hide_index=True
                )

        # ========================================================
        # TAB 3: DOWNLOADS
        # ========================================================
        elif selected_tab == TAB_OPTIONS[2]:

            st.markdown('<div class="section-title">📘 Combined Excel Report</div>', unsafe_allow_html=True)
            st.markdown("Download the full workbook with all summary sheets plus one sheet per campaign file.")

            st.download_button(
                label="📘 Download Combined Excel Report (all sheets)",
                data=st.session_state.excel_bytes,
                file_name=st.session_state.excel_filename,
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                width="stretch"
            )

            st.markdown("---")

            single_files = st.session_state.single_campaign_files or {}

            if single_files:
                st.markdown('<div class="section-title">📄 Individual Campaign Sheets</div>', unsafe_allow_html=True)
                st.markdown(
                    f"Download each campaign as its own single-sheet Excel file "
                    f"({len(single_files)} file(s) — including ones with **no matches**). "
                    f"Or grab them all at once as a ZIP."
                )

                if st.session_state.zip_bytes:
                    st.download_button(
                        label=f"🗂️ Download ALL {len(single_files)} Campaign Sheets (ZIP)",
                        data=st.session_state.zip_bytes,
                        file_name="Campaign_Individual_Sheets.zip",
                        mime="application/zip",
                        width="stretch",
                        key="dl_zip_all",
                    )
                    st.markdown("")

                for campaign_name, info in single_files.items():
                    date_str = f" · 📅 {info['date']}" if info["date"] else ""
                    has_match = info.get("has_match", True)
                    if has_match:
                        badge = "✅ Has Matches"
                        card_class = "download-card"
                    else:
                        badge = "⚠️ No Matches"
                        card_class = "download-card no-match"

                    st.markdown(
                        f"""
                        <div class="{card_class}">
                            <div style="font-weight:700; color:#0f2027; font-size:1.05rem;">
                                📄 {campaign_name}
                            </div>
                            <div style="color:#64748b; font-size:0.85rem; margin-top:4px;">
                                Schemes: <b>{info['scheme_count']}</b>{date_str} · {badge}
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    st.download_button(
                        label=f"⬇️ Download {info['filename']}",
                        data=info["bytes"],
                        file_name=info["filename"],
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        ),
                        key=f"dl_{campaign_name}",
                        width="stretch",
                    )


    # ============================================================
    # EMPTY STATE MESSAGES
    # ============================================================

    elif not campaign_files and not transaction_file:
        st.info("👆 Upload Campaign files and a Transaction file to begin.")
    elif not campaign_files:
        st.warning("⚠️ Please upload at least one Campaign file.")
    elif not transaction_file:
        st.warning("⚠️ Please upload one Transaction file.")
