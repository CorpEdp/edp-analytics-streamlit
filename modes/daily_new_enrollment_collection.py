import streamlit as st
import pandas as pd
import numpy as np
import io
import re

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="E-Gold / E-Silver / Sessional Analysis",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# CONSTANTS
# ============================================================

CATEGORY_GOLD = "e-Gold"
CATEGORY_SILVER = "e-Silver"
CATEGORY_SESSIONAL = "Sessional"

GENERIC_BUCKET_NAMES = {
    "e-gold", "e silver", "e-silver", "egold", "esilver",
    "sessional", "e gold", "e silver",
}

SESSIONAL_LAUNCH_DATE = pd.Timestamp("2026-09-01")
SESSIONAL_AVAILABLE_DATE_FALLBACK = pd.Timestamp("2026-09-14")
REPORT_ANCHOR_MONTH = pd.Timestamp("2026-09-01")

DAY_START = 1
DAY_END_FALLBACK = 14

MAX_SHEET_COLS = 9


# ============================================================
# CORE HELPERS
# ============================================================

def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_column_name(column):
    column = str(column).strip()
    column = re.sub(r"\s+", " ", column)
    return column


def find_column(df, possible_names, required=True):
    columns = list(df.columns)
    normalized_columns = {str(c).strip().lower(): c for c in columns}
    for name in possible_names:
        key = str(name).strip().lower()
        if key in normalized_columns:
            return normalized_columns[key]
    for name in possible_names:
        search_name = str(name).strip().lower()
        for col in columns:
            if search_name in str(col).strip().lower():
                return col
    if required:
        raise ValueError(
            f"Required column not found.\n\n"
            f"Expected one of:\n{possible_names}\n\n"
            f"Columns available:\n{columns}"
        )
    return None


def safe_numeric(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def classify_scheme(scheme_name):
    name = normalize_text(scheme_name).lower()
    compact = re.sub(r"[\s_\-]+", "", name)
    if "egold" in compact:
        return CATEGORY_GOLD
    if "esilver" in compact:
        return CATEGORY_SILVER
    return CATEGORY_SESSIONAL


def is_generic_bucket(scheme_name):
    key = re.sub(r"[\s_\-]+", "", str(scheme_name).strip().lower())
    return key in {re.sub(r"[\s_\-]+", "", g) for g in GENERIC_BUCKET_NAMES}


def classify_transaction(df, installment_column):
    df["_Installment_Number"] = safe_numeric(df[installment_column])
    df["Transaction Type"] = np.select(
        [
            df["_Installment_Number"] == 1,
            df["_Installment_Number"] > 1
        ],
        ["Enrolment", "Collection"],
        default="Exclude"
    )
    return df


def pct_change(old_value, new_value):
    if pd.isna(old_value) or old_value == 0:
        if new_value == 0:
            return 0.0
        return np.nan
    return (new_value - old_value) / old_value * 100


def fmt_pct_value(pct, new_value=0):
    if pd.isna(pct):
        if new_value > 0:
            return "NEW"
        return 0.0
    return float(pct)


# ============================================================
# WINDOW ENGINE
# ============================================================

def _month_window(year, month, day_start, day_end):
    last_day = (
        pd.Timestamp(year=year, month=month, day=1)
        + pd.DateOffset(months=1) - pd.DateOffset(days=1)
    ).day
    start_ts = pd.Timestamp(year=year, month=month, day=day_start)
    end_ts = pd.Timestamp(year=year, month=month, day=min(day_end, last_day))
    return start_ts, end_ts


def clip_df_to_matched_window(df, day_start, day_end):
    if df.empty:
        return df
    months = df["Month Date"].dropna().unique()
    bounds = {}
    for m in months:
        m_ts = pd.Timestamp(m)
        s, e = _month_window(m_ts.year, m_ts.month, day_start, day_end)
        bounds[m] = (s, e)
    mask = df.apply(
        lambda r: (
            r["Month Date"] in bounds
            and bounds[r["Month Date"]][0] <= r["Date"] <= bounds[r["Month Date"]][1]
        ),
        axis=1,
    )
    return df[mask].copy()


# ============================================================
# STYLE HELPERS
# ============================================================

def get_styles():
    return {
        "title_font": Font(name="Calibri", size=18, bold=True, color="FFFFFF"),
        "title_fill": PatternFill(fill_type="solid", fgColor="17365D"),
        "section_font": Font(name="Calibri", size=13, bold=True, color="FFFFFF"),
        "section_fill": PatternFill(fill_type="solid", fgColor="2E75B6"),
        "header_font": Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
        "header_fill": PatternFill(fill_type="solid", fgColor="4472C4"),
        "subheader_font": Font(name="Calibri", size=11, bold=True, color="17365D"),
        "subheader_fill": PatternFill(fill_type="solid", fgColor="D9E1F2"),
        "label_font": Font(name="Calibri", size=11, bold=True),
        "value_font": Font(name="Calibri", size=11),
        "total_font": Font(name="Calibri", size=11, bold=True, color="17365D"),
        "center_align": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "left_align": Alignment(horizontal="left", vertical="center", wrap_text=True),
        "right_align": Alignment(horizontal="right", vertical="center"),
        "thin_border": Border(
            left=Side(style="thin", color="B4C6E7"),
            right=Side(style="thin", color="B4C6E7"),
            top=Side(style="thin", color="B4C6E7"),
            bottom=Side(style="thin", color="B4C6E7")
        ),
        "medium_border": Border(
            left=Side(style="medium", color="2E75B6"),
            right=Side(style="medium", color="2E75B6"),
            top=Side(style="medium", color="2E75B6"),
            bottom=Side(style="medium", color="2E75B6")
        ),
        "currency_fmt": '₹#,##0.00',
        "number_fmt": '#,##0',
        "pct_fmt": '0.00"%"',
        "green_fill": PatternFill(fill_type="solid", fgColor="C6EFCE"),
        "lightgreen_fill": PatternFill(fill_type="solid", fgColor="E2EFDA"),
        "orange_fill": PatternFill(fill_type="solid", fgColor="FFEB9C"),
        "red_fill": PatternFill(fill_type="solid", fgColor="FFC7CE"),
        "gray_fill": PatternFill(fill_type="solid", fgColor="F2F2F2"),
        "highlight_fill": PatternFill(fill_type="solid", fgColor="FFF2CC"),
        "live_fill": PatternFill(fill_type="solid", fgColor="FFF2CC"),
        "total_fill": PatternFill(fill_type="solid", fgColor="BDD7EE"),
        "green_font": Font(name="Calibri", size=11, bold=True, color="006100"),
        "red_font": Font(name="Calibri", size=11, bold=True, color="9C0006"),
        "orange_font": Font(name="Calibri", size=11, bold=True, color="9C6500"),
    }


def apply_performance_fill(cell, pct):
    S = get_styles()
    if isinstance(pct, str):
        cell.fill = S["green_fill"]
        cell.font = S["green_font"]
        return
    if pd.isna(pct):
        cell.fill = S["gray_fill"]
        return
    if pct > 10:
        cell.fill = S["green_fill"]
        cell.font = S["green_font"]
    elif pct > 0:
        cell.fill = S["lightgreen_fill"]
    elif pct > -10:
        cell.fill = S["orange_fill"]
        cell.font = S["orange_font"]
    else:
        cell.fill = S["red_fill"]
        cell.font = S["red_font"]


# ============================================================
# SINGLE SHEET WRITER
# ============================================================

class SingleSheetWriter:
    def __init__(self, ws, max_cols=MAX_SHEET_COLS):
        self.ws = ws
        self.max_cols = max_cols
        self.row = 1
        self.S = get_styles()

    def blank(self, n=1):
        self.row += n

    def title(self, text):
        S = self.S
        self.ws.merge_cells(start_row=self.row, start_column=1,
                            end_row=self.row, end_column=self.max_cols)
        c = self.ws.cell(row=self.row, column=1, value=text)
        c.font = S["title_font"]; c.fill = S["title_fill"]
        c.alignment = S["center_align"]
        self.ws.row_dimensions[self.row].height = 38
        for col in range(1, self.max_cols + 1):
            self.ws.cell(row=self.row, column=col).fill = S["title_fill"]
        self.row += 1

    def subtitle(self, text):
        self.ws.merge_cells(start_row=self.row, start_column=1,
                            end_row=self.row, end_column=self.max_cols)
        c = self.ws.cell(row=self.row, column=1, value=text)
        c.font = Font(italic=True, size=10, color="404040")
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        self.ws.row_dimensions[self.row].height = 24
        self.row += 1

    def note(self, text):
        S = self.S
        self.ws.merge_cells(start_row=self.row, start_column=1,
                            end_row=self.row, end_column=self.max_cols)
        c = self.ws.cell(row=self.row, column=1, value=text)
        c.font = Font(italic=True, size=10, color="404040")
        c.alignment = Alignment(horizontal="left", vertical="center",
                                wrap_text=True)
        c.fill = S["gray_fill"]
        self.ws.row_dimensions[self.row].height = 28
        self.row += 1

    def section(self, text):
        S = self.S
        self.ws.merge_cells(start_row=self.row, start_column=1,
                            end_row=self.row, end_column=self.max_cols)
        c = self.ws.cell(row=self.row, column=1, value=text)
        c.font = S["section_font"]; c.fill = S["section_fill"]
        c.alignment = S["left_align"]
        self.ws.row_dimensions[self.row].height = 26
        for col in range(1, self.max_cols + 1):
            self.ws.cell(row=self.row, column=col).fill = S["section_fill"]
        self.row += 1

    def subsection(self, text):
        S = self.S
        self.ws.merge_cells(start_row=self.row, start_column=1,
                            end_row=self.row, end_column=self.max_cols)
        c = self.ws.cell(row=self.row, column=1, value=text)
        c.font = S["subheader_font"]; c.fill = S["subheader_fill"]
        c.alignment = S["left_align"]
        self.ws.row_dimensions[self.row].height = 22
        for col in range(1, self.max_cols + 1):
            self.ws.cell(row=self.row, column=col).fill = S["subheader_fill"]
        self.row += 1

    def headers(self, headers):
        S = self.S
        headers = list(headers) + [""] * (self.max_cols - len(headers))
        headers = headers[:self.max_cols]
        for i, h in enumerate(headers, start=1):
            c = self.ws.cell(row=self.row, column=i, value=h)
            c.font = S["header_font"]; c.fill = S["header_fill"]
            c.alignment = S["center_align"]; c.border = S["thin_border"]
        self.ws.row_dimensions[self.row].height = 32
        self.row += 1

    def row_values(self, values, formats=None, aligns=None,
                   pct_cols=None, total=False, live=False):
        S = self.S
        pct_cols = pct_cols or []
        values = list(values) + [""] * (self.max_cols - len(values))
        values = values[:self.max_cols]
        if formats is None:
            formats = [None] * self.max_cols
        else:
            formats = list(formats) + [None] * (self.max_cols - len(formats))
            formats = formats[:self.max_cols]
        if aligns is None:
            aligns = [None] * self.max_cols
        else:
            aligns = list(aligns) + [None] * (self.max_cols - len(aligns))
            aligns = aligns[:self.max_cols]

        for i, v in enumerate(values, start=1):
            c = self.ws.cell(row=self.row, column=i, value=v)
            c.font = S["total_font"] if total else S["value_font"]
            c.border = S["thin_border"]
            c.alignment = (
                aligns[i - 1] if aligns[i - 1]
                else (S["left_align"] if i == 1 else S["right_align"])
            )
            if formats[i - 1]:
                if not (isinstance(v, str) and v == "NEW"):
                    c.number_format = formats[i - 1]
            if total:
                c.fill = S["total_fill"]
            if live:
                c.fill = S["live_fill"]
            if i - 1 in pct_cols:
                if isinstance(v, (int, float, np.number)) or v == "NEW":
                    apply_performance_fill(c, v)
        self.row += 1

    def insight_block(self, lines):
        S = self.S
        for line in lines:
            self.ws.merge_cells(start_row=self.row, start_column=1,
                                end_row=self.row, end_column=self.max_cols)
            c = self.ws.cell(row=self.row, column=1, value=line)
            c.font = Font(size=11, bold=True, color="17365D")
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.fill = S["highlight_fill"]
            self.ws.row_dimensions[self.row].height = 32
            self.row += 1


# ============================================================
# ANALYSIS BUILDERS
# ============================================================

def build_monthly_matched_table(df, id_col, reference_month, day_start, day_end):
    rows = []
    for month_date in sorted(df["Month Date"].unique()):
        month_ts = pd.Timestamp(month_date)
        start_ts, end_ts = _month_window(
            month_ts.year, month_ts.month, day_start, day_end)
        month_df = df[(df["Date"] >= start_ts) & (df["Date"] <= end_ts)]
        is_ref = (month_ts == reference_month)
        label = (f"{month_ts.strftime('%Y-%m')} (Day {day_start}–{end_ts.day})"
                 + ("  ⚠️ LIVE" if is_ref else ""))
        row = {"Month": label, "_is_live": is_ref}
        for lbl, cat in [("e-Gold", CATEGORY_GOLD),
                         ("e-Silver", CATEGORY_SILVER),
                         ("Sessional", CATEGORY_SESSIONAL)]:
            enrol = month_df[(month_df["Scheme Category"] == cat) &
                             (month_df["Transaction Type"] == "Enrolment")]
            coll = month_df[(month_df["Scheme Category"] == cat) &
                            (month_df["Transaction Type"] == "Collection")]
            row[f"{lbl} Enrolment Count"] = len(enrol)
            row[f"{lbl} Enrolment Amount"] = enrol["Saved Amount"].sum()
            row[f"{lbl} Enrolment Average Ticket"] = (
                row[f"{lbl} Enrolment Amount"] / len(enrol) if len(enrol) else 0)
            row[f"{lbl} Collection Count"] = len(coll)
            row[f"{lbl} Collection Amount"] = coll["Saved Amount"].sum()
            row[f"{lbl} Collection Average Ticket"] = (
                row[f"{lbl} Collection Amount"] / len(coll) if len(coll) else 0)
        row["Total Enrolment Count"] = (row["e-Gold Enrolment Count"] +
                                        row["e-Silver Enrolment Count"] +
                                        row["Sessional Enrolment Count"])
        row["Total Enrolment Amount"] = (row["e-Gold Enrolment Amount"] +
                                         row["e-Silver Enrolment Amount"] +
                                         row["Sessional Enrolment Amount"])
        row["Total Collection Count"] = (row["e-Gold Collection Count"] +
                                         row["e-Silver Collection Count"] +
                                         row["Sessional Collection Count"])
        row["Total Collection Amount"] = (row["e-Gold Collection Amount"] +
                                          row["e-Silver Collection Amount"] +
                                          row["Sessional Collection Amount"])
        rows.append(row)
    return pd.DataFrame(rows)


def build_window(df, start_ts, end_ts, label, scheme=None):
    mask = (df["Date"] >= start_ts) & (df["Date"] <= end_ts)
    if scheme is not None:
        mask &= (df["Scheme Category"] == scheme)
    win = df[mask]
    enrol = win[win["Transaction Type"] == "Enrolment"]
    coll = win[win["Transaction Type"] == "Collection"]
    e_count, e_amount = len(enrol), enrol["Saved Amount"].sum()
    c_count, c_amount = len(coll), coll["Saved Amount"].sum()
    return {
        "label": label, "start": start_ts, "end": end_ts,
        "days": (end_ts - start_ts).days + 1,
        "Enrol. Count": e_count, "Enrol. Amount": e_amount,
        "Enrol. Avg Ticket": e_amount / e_count if e_count else 0,
        "Coll. Count": c_count, "Coll. Amount": c_amount,
        "Coll. Avg Ticket": c_amount / c_count if c_count else 0,
        "Total Count": e_count + c_count, "Total Amount": e_amount + c_amount,
    }


def build_window_list(df, scheme, base_year, base_month, day_start, day_end):
    windows = []
    for i in range(3, 0, -1):
        m = pd.Timestamp(year=base_year, month=base_month, day=1) - pd.DateOffset(months=i)
        start_ts, end_ts = _month_window(m.year, m.month, day_start, day_end)
        label = f"{m.strftime('%b %Y')} (Day {day_start}–{end_ts.day})"
        windows.append(build_window(df, start_ts, end_ts, label, scheme))
    start_ts, end_ts = _month_window(base_year, base_month, day_start, day_end)
    label = (f"{pd.Timestamp(year=base_year, month=base_month, day=1).strftime('%b %Y')} "
             f"(Day {day_start}–{end_ts.day})")
    windows.append(build_window(df, start_ts, end_ts, label, scheme))
    return windows


def build_combined_window_list(df, schemes, base_year, base_month,
                               day_start, day_end):
    subset = df[df["Scheme Category"].isin(schemes)]
    return build_window_list(subset, None, base_year, base_month,
                             day_start, day_end)


def build_scheme_impact_analysis(filtered_df, launch_date, id_col,
                                 reference_month, day_start, day_end):
    if filtered_df.empty:
        return pd.DataFrame()
    named_df = filtered_df[~filtered_df["Actual Scheme Name"].apply(is_generic_bucket)]
    if named_df.empty:
        return pd.DataFrame()
    pre_clipped = clip_df_to_matched_window(
        named_df[named_df["Date"] < launch_date], day_start, day_end)
    post_clipped = clip_df_to_matched_window(
        named_df[named_df["Date"] >= launch_date], day_start, day_end)
    pre_months = pre_clipped["Month Date"].nunique() if not pre_clipped.empty else 0
    post_months = post_clipped["Month Date"].nunique() if not post_clipped.empty else 0
    rows = []
    for scheme in sorted(named_df["Actual Scheme Name"].unique()):
        cat = classify_scheme(scheme)
        pre_s = pre_clipped[pre_clipped["Actual Scheme Name"] == scheme]
        post_s = post_clipped[post_clipped["Actual Scheme Name"] == scheme]
        pre_enrol = pre_s[pre_s["Transaction Type"] == "Enrolment"]
        post_enrol = post_s[post_s["Transaction Type"] == "Enrolment"]
        pre_count, post_count = len(pre_enrol), len(post_enrol)
        pre_amount = pre_enrol["Saved Amount"].sum()
        post_amount = post_enrol["Saved Amount"].sum()
        pre_avg_m_count = pre_count / pre_months if pre_months else 0
        post_avg_m_count = post_count / post_months if post_months else 0
        pre_avg_m_amount = pre_amount / pre_months if pre_months else 0
        post_avg_m_amount = post_amount / post_months if post_months else 0
        pre_ticket = pre_amount / pre_count if pre_count else 0
        post_ticket = post_amount / post_count if post_count else 0
        total_pre_amount = pre_clipped["Saved Amount"].sum()
        total_post_amount = post_clipped["Saved Amount"].sum()
        pre_share = (pre_amount / total_pre_amount * 100) if total_pre_amount else 0
        post_share = (post_amount / total_post_amount * 100) if total_post_amount else 0
        rows.append({
            "Scheme Name": scheme, "Category": cat,
            "Pre Enrol. Count": pre_count, "Post Enrol. Count": post_count,
            "Pre Enrol. Amount": pre_amount, "Post Enrol. Amount": post_amount,
            "Pre Avg Monthly Count": pre_avg_m_count,
            "Post Avg Monthly Count": post_avg_m_count,
            "Pre Avg Monthly Amount": pre_avg_m_amount,
            "Post Avg Monthly Amount": post_avg_m_amount,
            "Pre Avg Ticket": pre_ticket, "Post Avg Ticket": post_ticket,
            "Count Change %": pct_change(pre_avg_m_count, post_avg_m_count),
            "Amount Change %": pct_change(pre_avg_m_amount, post_avg_m_amount),
            "Ticket Change %": pct_change(pre_ticket, post_ticket),
            "Pre Share %": pre_share, "Post Share %": post_share,
            "Share Change %": post_share - pre_share,
        })
    return pd.DataFrame(rows)


def build_scheme_share_by_month(df, id_col, day_start, day_end):
    if df.empty:
        return pd.DataFrame()
    clipped = clip_df_to_matched_window(df, day_start, day_end)
    if clipped.empty:
        return pd.DataFrame()
    pivot = (clipped.groupby(["Month", "Actual Scheme Name"])
             .agg(Amount=("Saved Amount", "sum")).reset_index())
    total_by_month = pivot.groupby("Month")["Amount"].sum().rename("Total")
    pivot = pivot.merge(total_by_month, on="Month")
    pivot["Share %"] = np.where(
        pivot["Total"] > 0, pivot["Amount"] / pivot["Total"] * 100, 0)
    return pivot.pivot_table(index="Month", columns="Actual Scheme Name",
                             values="Share %", aggfunc="sum",
                             fill_value=0).reset_index()


def compute_grand_totals(df):
    enrol = df[df["Transaction Type"] == "Enrolment"]
    coll = df[df["Transaction Type"] == "Collection"]

    def per_cat(cat):
        e = enrol[enrol["Scheme Category"] == cat]
        c = coll[coll["Scheme Category"] == cat]
        return {"Enrol_Count": len(e),
                "Enrol_Amount": float(e["Saved Amount"].sum()),
                "Coll_Count": len(c),
                "Coll_Amount": float(c["Saved Amount"].sum())}

    g = per_cat(CATEGORY_GOLD)
    s = per_cat(CATEGORY_SILVER)
    ss = per_cat(CATEGORY_SESSIONAL)
    return {"e-Gold": g, "e-Silver": s, "Sessional": ss,
            "Grand": {
                "Enrol_Count": g["Enrol_Count"] + s["Enrol_Count"] + ss["Enrol_Count"],
                "Enrol_Amount": g["Enrol_Amount"] + s["Enrol_Amount"] + ss["Enrol_Amount"],
                "Coll_Count": g["Coll_Count"] + s["Coll_Count"] + ss["Coll_Count"],
                "Coll_Amount": g["Coll_Amount"] + s["Coll_Amount"] + ss["Coll_Amount"],
                "Total_Count": (g["Enrol_Count"] + s["Enrol_Count"] + ss["Enrol_Count"] +
                                g["Coll_Count"] + s["Coll_Count"] + ss["Coll_Count"]),
                "Total_Amount": (g["Enrol_Amount"] + s["Enrol_Amount"] +
                                 ss["Enrol_Amount"] + g["Coll_Amount"] +
                                 s["Coll_Amount"] + ss["Coll_Amount"]),
            }}


# ============================================================
# SECTION 16 — SESSIONAL OVERLAP (e-Gold / e-Silver / Both)
# ============================================================

def build_sessional_overlap_report(filtered_df, launch_date,
                                   customer_col="Customer_ID",
                                   post_cutoff=None):
    if filtered_df.empty or customer_col not in filtered_df.columns:
        return None
    df = filtered_df[filtered_df[customer_col].notna()].copy()
    df = df[df[customer_col].astype(str).str.len() > 0]
    if df.empty:
        return None
    sess = df[(df["Scheme Category"] == CATEGORY_SESSIONAL) &
              (df["Transaction Type"] == "Enrolment") &
              (df["Date"] >= launch_date)].copy()
    if sess.empty:
        return None

    sessional_enrollees = (
        sess.groupby(customer_col)
        .agg(Sessional_Scheme=("Actual Scheme Name",
                               lambda s: " | ".join(sorted(set(s)))),
             Sessional_Enrol_Count=("Saved Amount", "size"),
             Sessional_Enrol_Amount=("Saved Amount", "sum"),
             First_Sessional_Date=("Date", "min"))
        .reset_index()
        .rename(columns={customer_col: "Customer_ID"})
    )

    legacy_pre = df[(df["Scheme Category"].isin([CATEGORY_GOLD, CATEGORY_SILVER])) &
                    (df["Date"] < launch_date)].copy()

    gold_ids = set(legacy_pre.loc[legacy_pre["Scheme Category"] == CATEGORY_GOLD,
                                  customer_col].dropna().astype(str).unique())
    silver_ids = set(legacy_pre.loc[legacy_pre["Scheme Category"] == CATEGORY_SILVER,
                                    customer_col].dropna().astype(str).unique())

    def classify_bucket(cid):
        cid = str(cid)
        in_gold = cid in gold_ids
        in_silver = cid in silver_ids
        if in_gold and in_silver:  return "Existing — Both"
        if in_gold:                return "Existing — e-Gold only"
        if in_silver:              return "Existing — e-Silver only"
        return "New"

    sessional_enrollees["Legacy_Bucket"] = (
        sessional_enrollees["Customer_ID"].apply(classify_bucket)
    )

    overlap_df = sessional_enrollees[
        sessional_enrollees["Legacy_Bucket"] != "New"].copy()
    new_df = sessional_enrollees[
        sessional_enrollees["Legacy_Bucket"] == "New"].copy()

    if not overlap_df.empty:
        overlap_ids = set(overlap_df["Customer_ID"].astype(str))
        pre_legacy_grp = (
            legacy_pre[legacy_pre[customer_col].astype(str).isin(overlap_ids)]
            .groupby(customer_col)
            .agg(Pre_Legacy_Amount=("Saved Amount", "sum"),
                 Pre_Legacy_Count=("Saved Amount", "size"),
                 Pre_Legacy_Schemes=("Scheme Category",
                                     lambda s: " | ".join(sorted(set(s)))),
                 Last_Legacy_Date=("Date", "max"))
            .reset_index()
            .rename(columns={customer_col: "Customer_ID"})
        )
        overlap_df = overlap_df.merge(pre_legacy_grp, on="Customer_ID", how="left")

    buckets = ["Existing — Both",
               "Existing — e-Gold only",
               "Existing — e-Silver only",
               "New"]
    bucket_totals = {}
    for b in buckets:
        sub = sessional_enrollees[sessional_enrollees["Legacy_Bucket"] == b]
        bucket_totals[b] = {
            "count": int(len(sub)),
            "amount": float(sub["Sessional_Enrol_Amount"].sum()),
            "avg_ticket": float(sub["Sessional_Enrol_Amount"].mean())
                          if len(sub) else 0.0,
        }

    total_count = sum(v["count"] for v in bucket_totals.values())
    total_amount = sum(v["amount"] for v in bucket_totals.values())

    existing_count = (bucket_totals["Existing — e-Gold only"]["count"]
                      + bucket_totals["Existing — e-Silver only"]["count"]
                      + bucket_totals["Existing — Both"]["count"])
    existing_amount = (bucket_totals["Existing — e-Gold only"]["amount"]
                       + bucket_totals["Existing — e-Silver only"]["amount"]
                       + bucket_totals["Existing — Both"]["amount"])

    summary_rows = [
        ("Total Sessional Enrollees", total_count, "count"),
        ("Sessional Enrolment Amount — Total", total_amount, "amount"),
        ("Existing — e-Gold only (count)",
         bucket_totals["Existing — e-Gold only"]["count"], "count"),
        ("Existing — e-Gold only (amount)",
         bucket_totals["Existing — e-Gold only"]["amount"], "amount"),
        ("Existing — e-Silver only (count)",
         bucket_totals["Existing — e-Silver only"]["count"], "count"),
        ("Existing — e-Silver only (amount)",
         bucket_totals["Existing — e-Silver only"]["amount"], "amount"),
        ("Existing — Both e-Gold & e-Silver (count)",
         bucket_totals["Existing — Both"]["count"], "count"),
        ("Existing — Both e-Gold & e-Silver (amount)",
         bucket_totals["Existing — Both"]["amount"], "amount"),
        ("New (no prior legacy) — count",
         bucket_totals["New"]["count"], "count"),
        ("New (no prior legacy) — amount",
         bucket_totals["New"]["amount"], "amount"),
        ("Total Existing (any legacy) — count", existing_count, "count"),
        ("Total Existing (any legacy) — amount", existing_amount, "amount"),
        ("% of Sessional Amount from Existing",
         (existing_amount / total_amount * 100) if total_amount else 0.0, "pct"),
    ]

    # ---------- Per-scheme × bucket breakdown (UNIQUE CUSTOMERS) ----------
    per_scheme_rows = []
    for scheme in sorted(sess["Actual Scheme Name"].unique()):
        sch_sess = sess[sess["Actual Scheme Name"] == scheme]
        scheme_customers = set(sch_sess[customer_col].astype(str).unique())

        row = {"Scheme": scheme}
        for b in buckets:
            ids_b = set(sessional_enrollees.loc[
                sessional_enrollees["Legacy_Bucket"] == b,
                "Customer_ID"].astype(str))
            bucket_scheme_ids = scheme_customers & ids_b
            sub = sch_sess[sch_sess[customer_col].astype(str).isin(bucket_scheme_ids)]
            unique_cnt = sub[customer_col].astype(str).nunique()
            row[f"{b} — Enrollees"] = int(unique_cnt)
            row[f"{b} — Amount"] = float(sub["Saved Amount"].sum())
        row["Total Enrollees"] = int(len(scheme_customers))
        row["Total Amount"] = float(sch_sess["Saved Amount"].sum())
        per_scheme_rows.append(row)

    post_cutoff = post_cutoff or pd.Timestamp(df["Date"].max())
    overlap_ids = set(sessional_enrollees.loc[
        sessional_enrollees["Legacy_Bucket"] != "New", "Customer_ID"].astype(str))
    overlap_legacy_post = df[
        (df[customer_col].astype(str).isin(overlap_ids)) &
        (df["Scheme Category"].isin([CATEGORY_GOLD, CATEGORY_SILVER])) &
        (df["Date"] >= launch_date) & (df["Date"] <= post_cutoff)]
    pre_total = float(legacy_pre[
        legacy_pre[customer_col].astype(str).isin(overlap_ids)]
        ["Saved Amount"].sum())
    post_total = float(overlap_legacy_post["Saved Amount"].sum())

    return {
        "sessional_enrollees": sessional_enrollees,
        "overlap_df": overlap_df,
        "new_df": new_df,
        "summary_rows": summary_rows,
        "per_scheme_rows": per_scheme_rows,
        "bucket_totals": bucket_totals,
        "pre_activity_of_overlap": {
            "pre_total_legacy_amount": pre_total,
            "post_total_legacy_amount": post_total,
            "legacy_amount_change": post_total - pre_total,
        },
    }


# ============================================================
# EXCEL SHEET BUILDER
# ============================================================

def build_single_overall_sheet(
    wb, filtered_df, id_col, start_date, end_date,
    totals_block, category_summary, monthly_analysis,
    pre_post_df, gold_silver_comparison, gold_window, silver_window,
    combined_window, sessional_first_days, sessional_scheme_table,
    scheme_impact_df, scheme_share_df, cannibalisation_rows,
    insights, interpretation, sessional_interpretation,
    sessional_launch_date, sessional_live_start, sessional_live_end,
    last_report_live_date, day_start, day_end,
    reference_month_label, day_window_label, day_window_days,
    overlap_report=None,
):
    S = get_styles()
    ws = wb.create_sheet("Overall Report")
    W = SingleSheetWriter(ws, max_cols=MAX_SHEET_COLS)

    W.title("E-GOLD / E-SILVER / SESSIONAL — OVERALL IMPACT REPORT")
    W.subtitle(
        f"Analysis Period: {start_date.strftime('%d-%b-%Y')} → "
        f"{end_date.strftime('%d-%b-%Y')}   |   "
        f"Sessional Launch: {sessional_launch_date.strftime('%d-%b-%Y')}   |   "
        f"Last Report Live: {last_report_live_date.strftime('%d-%b-%Y')}   |   "
        f"Live Window (ALL sections): {day_window_label}"
    )
    W.blank()

    # ---------- 1 ----------
    W.section("1. EXECUTIVE SUMMARY")
    W.headers(["Metric", "Count", "Amount", "Average Ticket"])
    for label, count, amount, avg in totals_block:
        W.row_values([label, int(count), float(amount), float(avg)],
                     formats=[None, S["number_fmt"], S["currency_fmt"],
                              S["currency_fmt"]],
                     aligns=[S["left_align"], S["right_align"],
                             S["right_align"], S["right_align"]])
    W.blank()

    # ---------- 2 ----------
    W.section("2. CATEGORY PERFORMANCE — COUNT & AMOUNT")
    W.headers(["Scheme Category", "Transaction Type", "Count",
               "Amount", "Average Ticket"])
    if not category_summary.empty:
        for _, r in category_summary.iterrows():
            W.row_values([r["Scheme Category"], r["Transaction Type"],
                          int(r["Count"]), float(r["Amount"]),
                          float(r["Average Ticket"] or 0)],
                         formats=[None, None, S["number_fmt"], S["currency_fmt"],
                                  S["currency_fmt"]],
                         aligns=[S["left_align"], S["center_align"],
                                 S["right_align"], S["right_align"],
                                 S["right_align"]])
    W.blank()

    # ---------- 3 ----------
    W.section(f"3. MONTHLY BUSINESS ANALYSIS — MATCHED LIVE WINDOW ({day_window_label})")
    W.note(f"⚠️ Every month is clipped to the SAME live window "
           f"{day_window_label} ({day_window_days} days). All rows are directly "
           f"comparable. Reference month ({reference_month_label}) is highlighted.")
    W.headers(["Month (Day Window)",
               "e-Gold Enrol Count", "e-Gold Enrol Amount",
               "e-Silver Enrol Count", "e-Silver Enrol Amount",
               "Sessional Enrol Count", "Sessional Enrol Amount",
               "Total Enrol Amount"])
    if not monthly_analysis.empty:
        for _, r in monthly_analysis.iterrows():
            is_live = bool(r.get("_is_live", False))
            W.row_values([
                r.get("Month", ""),
                int(r.get("e-Gold Enrolment Count", 0)),
                float(r.get("e-Gold Enrolment Amount", 0)),
                int(r.get("e-Silver Enrolment Count", 0)),
                float(r.get("e-Silver Enrolment Amount", 0)),
                int(r.get("Sessional Enrolment Count", 0)),
                float(r.get("Sessional Enrolment Amount", 0)),
                float(r.get("Total Enrolment Amount", 0)),
            ], formats=[None, S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["currency_fmt"], S["currency_fmt"]],
               aligns=[S["center_align"]] + [S["right_align"]] * 7, live=is_live)
        W.row_values([
            f"TOTAL ({day_window_label} × {len(monthly_analysis)} months)",
            int(monthly_analysis["e-Gold Enrolment Count"].sum()),
            float(monthly_analysis["e-Gold Enrolment Amount"].sum()),
            int(monthly_analysis["e-Silver Enrolment Count"].sum()),
            float(monthly_analysis["e-Silver Enrolment Amount"].sum()),
            int(monthly_analysis["Sessional Enrolment Count"].sum()),
            float(monthly_analysis["Sessional Enrolment Amount"].sum()),
            float(monthly_analysis["Total Enrolment Amount"].sum()),
        ], formats=[None, S["number_fmt"], S["currency_fmt"],
                    S["number_fmt"], S["currency_fmt"],
                    S["number_fmt"], S["currency_fmt"], S["currency_fmt"]],
           aligns=[S["center_align"]] + [S["right_align"]] * 7, total=True)
    W.blank()

    # ---------- 4 ----------
    W.section("4. ENROLMENT vs COLLECTION BREAKDOWN")
    W.headers(["Category", "Enrol Count", "Enrol Amount",
               "Enrol Avg Ticket", "Coll Count", "Coll Amount",
               "Coll Avg Ticket"])
    enrol = filtered_df[filtered_df["Transaction Type"] == "Enrolment"]
    coll = filtered_df[filtered_df["Transaction Type"] == "Collection"]
    for cat in [CATEGORY_GOLD, CATEGORY_SILVER, CATEGORY_SESSIONAL]:
        e = enrol[enrol["Scheme Category"] == cat]
        c = coll[coll["Scheme Category"] == cat]
        W.row_values([cat, len(e), float(e["Saved Amount"].sum()),
                      float(e["Saved Amount"].mean()) if len(e) else 0,
                      len(c), float(c["Saved Amount"].sum()),
                      float(c["Saved Amount"].mean()) if len(c) else 0],
                     formats=[None, S["number_fmt"], S["currency_fmt"],
                              S["currency_fmt"], S["number_fmt"],
                              S["currency_fmt"], S["currency_fmt"]])
    W.row_values(["GRAND TOTAL", len(enrol), float(enrol["Saved Amount"].sum()),
                  float(enrol["Saved Amount"].mean()) if len(enrol) else 0,
                  len(coll), float(coll["Saved Amount"].sum()),
                  float(coll["Saved Amount"].mean()) if len(coll) else 0],
                 formats=[None, S["number_fmt"], S["currency_fmt"],
                          S["currency_fmt"], S["number_fmt"],
                          S["currency_fmt"], S["currency_fmt"]], total=True)
    W.blank()

    # ---------- 5 ----------
    W.section(f"5. e-GOLD vs e-SILVER — PREVIOUS vs CURRENT MONTH (both {day_window_label})")
    W.headers(["Scheme", "Prev Count", "Prev Amount", "Curr Count",
               "Curr Amount", "Count Δ", "Count Δ %", "Amount Δ %"])
    if not gold_silver_comparison.empty:
        for _, r in gold_silver_comparison.iterrows():
            W.row_values([
                r.get("Scheme Category", ""),
                int(r.get("Previous Month Enrol. Count", 0)),
                float(r.get("Previous Month Enrol. Amount", 0)),
                int(r.get("Current Month Enrol. Count", 0)),
                float(r.get("Current Month Enrol. Amount", 0)),
                float(r.get("Count Change", 0)),
                fmt_pct_value(r.get("Count Change %", np.nan),
                              r.get("Current Month Enrol. Count", 0)),
                fmt_pct_value(r.get("Amount Change %", np.nan),
                              r.get("Current Month Enrol. Amount", 0)),
            ], formats=[None, S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["pct_fmt"], S["pct_fmt"]],
               pct_cols=[6, 7])
    W.blank()

    # ---------- 6 ----------
    W.section(f"6. PREVIOUS 3 MONTHS (DAY {day_start}–{day_end}) vs "
              f"CURRENT MONTH (DAY {day_start}–{day_end}) — MATCHED WINDOWS")

    def render_window_block(title, windows):
        W.subsection(title)
        W.headers(["Window", "Enrol Count", "Enrol Amount",
                   "Coll Count", "Coll Amount",
                   "Δ Enrol Count", "Δ Enrol Count %",
                   "Δ Enrol Amount", "Δ Enrol Amount %"])
        for i, w in enumerate(windows):
            prev = windows[i - 1] if i > 0 else None
            d_count = w["Enrol. Count"] - prev["Enrol. Count"] if prev else 0
            d_amt = w["Enrol. Amount"] - prev["Enrol. Amount"] if prev else 0
            pc_count = pct_change(prev["Enrol. Count"], w["Enrol. Count"]) if prev else 0
            pc_amt = pct_change(prev["Enrol. Amount"], w["Enrol. Amount"]) if prev else 0
            W.row_values([
                w["label"],
                int(w["Enrol. Count"]), float(w["Enrol. Amount"]),
                int(w["Coll. Count"]), float(w["Coll. Amount"]),
                float(d_count),
                fmt_pct_value(pc_count, w["Enrol. Count"]) if prev else 0,
                float(d_amt),
                fmt_pct_value(pc_amt, w["Enrol. Amount"]) if prev else 0,
            ], formats=[None, S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["pct_fmt"],
                        S["currency_fmt"], S["pct_fmt"]],
               pct_cols=[5, 6, 8])
        W.blank()

    render_window_block("A. e-GOLD — MATCHED DAY WINDOWS", gold_window)
    render_window_block("B. e-SILVER — MATCHED DAY WINDOWS", silver_window)
    render_window_block("C. e-GOLD + e-SILVER COMBINED — MATCHED DAY WINDOWS",
                        combined_window)

    # Explicit blank row so Section D doesn't visually bleed into C
    W.blank()

    W.subsection(f"D. SESSIONAL LIVE WINDOW "
                 f"({sessional_live_start.strftime('%d-%b')} → "
                 f"{sessional_live_end.strftime('%d-%b')})")
    W.headers(["Category", "Transaction Type", "Count",
               "Amount", "Avg Ticket"])
    if sessional_first_days["days"] > 0 and sessional_first_days["details"]:
        for row in sessional_first_days["details"]:
            W.row_values([row["Scheme Category"], row["Transaction Type"],
                          int(row["Count"]), float(row["Amount"]),
                          float(row["Average Ticket"])],
                         formats=[None, None, S["number_fmt"],
                                  S["currency_fmt"], S["currency_fmt"]])
    else:
        W.note("No Sessional live data available yet.")
    W.blank()

    # ---------- 7 ----------
    W.section("7. SESSIONAL SCHEME DETAIL — ACTUAL SCHEME NAMES")
    W.headers(["Month", "Actual Scheme Name", "Transaction Type",
               "Count", "Amount", "Average Ticket"])
    if not sessional_scheme_table.empty:
        for _, r in sessional_scheme_table.iterrows():
            W.row_values([r.get("Month", ""), r.get("Actual Scheme Name", ""),
                          r.get("Transaction Type", ""),
                          int(r.get("Count", 0)), float(r.get("Amount", 0)),
                          float(r.get("Average Ticket") or 0)],
                         formats=[None, None, None, S["number_fmt"],
                                  S["currency_fmt"], S["currency_fmt"]],
                         aligns=[S["center_align"], S["left_align"],
                                 S["center_align"], S["right_align"],
                                 S["right_align"], S["right_align"]])
    W.blank()

    # ---------- 8 ----------
    W.section(f"8. PRE-SESSIONAL vs POST-SESSIONAL — Avg Monthly "
              f"(both clipped to {day_window_label})")
    W.headers(["Category", "Type", "Pre Avg Count", "Post Avg Count",
               "Δ Count %", "Pre Avg Amt", "Post Avg Amt", "Δ Amt %"])
    if not pre_post_df.empty:
        for _, r in pre_post_df.iterrows():
            W.row_values([
                r["Scheme Category"], r["Transaction Type"],
                float(r["Pre Monthly Avg Count"]),
                float(r["Post Monthly Avg Count"]),
                fmt_pct_value(r["Count Change %"], r["Post Monthly Avg Count"]),
                float(r["Pre Monthly Avg Amount"]),
                float(r["Post Monthly Avg Amount"]),
                fmt_pct_value(r["Amount Change %"], r["Post Monthly Avg Amount"]),
            ], formats=[None, None, S["number_fmt"], S["number_fmt"],
                        S["pct_fmt"], S["currency_fmt"],
                        S["currency_fmt"], S["pct_fmt"]],
               pct_cols=[4, 7])
    W.blank()

    # ---------- 9 ----------
    W.section("9. SCHEME-NAME LEVEL IMPACT — PRE vs POST (Named Schemes Only)")
    W.note(f"Excludes generic bucket names (e-Gold, e-Silver, Sessional). "
           f"All values clipped to {day_window_label}.")
    W.headers(["Scheme Name", "Category", "Pre Enrol", "Post Enrol",
               "Pre Avg Amt", "Post Avg Amt", "Δ Amt %", "Share Δ (pp)"])
    if not scheme_impact_df.empty:
        for _, r in scheme_impact_df.iterrows():
            W.row_values([r["Scheme Name"], r["Category"],
                          int(r["Pre Enrol. Count"]), int(r["Post Enrol. Count"]),
                          float(r["Pre Avg Monthly Amount"]),
                          float(r["Post Avg Monthly Amount"]),
                          fmt_pct_value(r["Amount Change %"],
                                        r["Post Avg Monthly Amount"]),
                          float(r["Share Change %"])],
                         formats=[None, None, S["number_fmt"], S["number_fmt"],
                                  S["currency_fmt"], S["currency_fmt"],
                                  S["pct_fmt"], S["pct_fmt"]],
                         pct_cols=[6, 7])
    else:
        W.note("No named schemes found (all entries are generic buckets).")
    W.blank()

    # ---------- 10 ----------
    W.section(f"10. SCHEME-NAME SHARE BY MONTH — all months clipped to "
              f"{day_window_label}")
    if not scheme_share_df.empty:
        cols = ["Month"] + [c for c in scheme_share_df.columns if c != "Month"]
        headers = cols[:MAX_SHEET_COLS]
        W.headers(headers)
        for _, r in scheme_share_df.iterrows():
            vals, fmts, aligns = [], [], []
            for c in headers:
                if c == "Month":
                    vals.append(r[c]); fmts.append(None)
                    aligns.append(S["center_align"])
                else:
                    vals.append(float(r[c])); fmts.append(S["pct_fmt"])
                    aligns.append(S["right_align"])
            W.row_values(vals, formats=fmts, aligns=aligns)
    W.blank()

    # ---------- 11 ----------
    W.section("11. SESSIONAL IMPACT ON e-GOLD / e-SILVER — MATCHED WINDOW")
    W.headers(["Metric", "Value"])
    for label, value, fmt in cannibalisation_rows:
        W.row_values([label, float(value)],
                     formats=[None, fmt],
                     aligns=[S["left_align"], S["right_align"]])
    W.blank()

    # ---------- 12 ----------
    W.section("12. PERCENTAGE GROWTH / DECLINE HIGHLIGHTS")

    def band(pct):
        if isinstance(pct, str):   return "🟢 New Business"
        if pd.isna(pct):           return "⚪ No change"
        if pct > 10:               return "🟢 Strong Growth"
        if pct > 0:                return "🟡 Moderate Growth"
        if pct > -10:              return "🟠 Moderate Decline"
        return "🔴 Sharp Decline"

    W.headers(["Source", "Metric", "Δ %",
               "Post Amount", "Δ Amount", "Signal"])

    highlight_rows = []

    for _, r in pre_post_df.iterrows():
        pre_amt = float(r["Pre Monthly Avg Amount"])
        post_amt = float(r["Post Monthly Avg Amount"])
        delta_amt = post_amt - pre_amt
        if pre_amt == 0 and post_amt > 0:
            delta_amt = post_amt
        highlight_rows.append({
            "Source": "Pre/Post Launch",
            "Metric": f"{r['Scheme Category']} — {r['Transaction Type']} Amount",
            "Δ %": r["Amount Change %"],
            "Post Amount": post_amt,
            "Δ Amount": delta_amt,
            "Post": post_amt,
        })

    if not gold_silver_comparison.empty:
        for _, r in gold_silver_comparison.iterrows():
            prev_amt = float(r.get("Previous Month Enrol. Amount", 0) or 0)
            curr_amt = float(r.get("Current Month Enrol. Amount", 0) or 0)
            highlight_rows.append({
                "Source": "Prev vs Curr Month",
                "Metric": f"{r['Scheme Category']} Enrolment Amount",
                "Δ %": r.get("Amount Change %", np.nan),
                "Post Amount": curr_amt,
                "Δ Amount": curr_amt - prev_amt,
                "Post": curr_amt,
            })

    if not scheme_impact_df.empty:
        for _, r in scheme_impact_df.iterrows():
            pre_amt = float(r["Pre Avg Monthly Amount"])
            post_amt = float(r["Post Avg Monthly Amount"])
            highlight_rows.append({
                "Source": "Scheme-Level Impact",
                "Metric": f"{r['Scheme Name']} — Avg Monthly Amount",
                "Δ %": r["Amount Change %"],
                "Post Amount": post_amt,
                "Δ Amount": post_amt - pre_amt,
                "Post": post_amt,
            })

    for row in highlight_rows:
        display_pct = fmt_pct_value(row["Δ %"], row["Post"])
        W.row_values(
            [row["Source"], row["Metric"], display_pct,
             row["Post Amount"], row["Δ Amount"],
             band(display_pct)],
            formats=[None, None, S["pct_fmt"],
                     S["currency_fmt"], S["currency_fmt"], None],
            pct_cols=[2]
        )
    W.blank()

    # ---------- 13 ----------
    W.section("13. GRAND TOTALS")
    W.headers(["Category", "Enrol Count", "Enrol Amount",
               "Coll Count", "Coll Amount", "Total Count", "Total Amount"])
    totals = compute_grand_totals(filtered_df)
    for cat in ["e-Gold", "e-Silver", "Sessional"]:
        t = totals[cat]
        W.row_values([cat, t["Enrol_Count"], t["Enrol_Amount"],
                      t["Coll_Count"], t["Coll_Amount"],
                      t["Enrol_Count"] + t["Coll_Count"],
                      t["Enrol_Amount"] + t["Coll_Amount"]],
                     formats=[None, S["number_fmt"], S["currency_fmt"],
                              S["number_fmt"], S["currency_fmt"],
                              S["number_fmt"], S["currency_fmt"]])
    g = totals["Grand"]
    W.row_values(["GRAND TOTAL", g["Enrol_Count"], g["Enrol_Amount"],
                  g["Coll_Count"], g["Coll_Amount"],
                  g["Total_Count"], g["Total_Amount"]],
                 formats=[None, S["number_fmt"], S["currency_fmt"],
                          S["number_fmt"], S["currency_fmt"],
                          S["number_fmt"], S["currency_fmt"]], total=True)
    W.blank()

    # ---------- 14 ----------
    W.section("14. MANAGEMENT INSIGHTS")
    W.insight_block(insights)
    W.blank()

    # ---------- 15 ----------
    W.section("15. INTERPRETATION — SEASONAL SCHEME IMPACT")
    W.insight_block([interpretation, sessional_interpretation])
    W.blank()

    # ---------- 16 ----------
    W.section("16. SESSIONAL ENROLLEES — PRIOR e-GOLD / e-SILVER BREAKDOWN")

    if overlap_report is None:
        W.note("⚠️ No customer identifier column in the uploaded file. "
               "Rename your phone/mobile column to 'Phone Number' / 'Mobile' "
               "to enable this report.")
    else:
        W.note("Splits Sessional enrollees into 4 buckets: e-Gold only, "
               "e-Silver only, Both (e-Gold AND e-Silver), and New "
               "(no prior legacy activity before 01-Sep-2026). "
               "Counts in part C are UNIQUE CUSTOMERS per (scheme × bucket).")
        bt = overlap_report["bucket_totals"]

        W.subsection("A. Sessional Enrollees — Split by Prior Legacy")
        W.headers(["Bucket", "Enrollees", "Amount", "Avg Ticket"])
        for bucket_label in ["Existing — e-Gold only",
                             "Existing — e-Silver only",
                             "Existing — Both",
                             "New"]:
            b = bt[bucket_label]
            W.row_values([bucket_label, int(b["count"]),
                          float(b["amount"]), float(b["avg_ticket"])],
                         formats=[None, S["number_fmt"], S["currency_fmt"],
                                  S["currency_fmt"]],
                         aligns=[S["left_align"], S["right_align"],
                                 S["right_align"], S["right_align"]])
        W.blank()

        W.subsection("B. Summary")
        W.headers(["Metric", "Value"])
        for label, value, fmt in overlap_report["summary_rows"]:
            if fmt == "amount":
                W.row_values([label, float(value)],
                             formats=[None, S["currency_fmt"]],
                             aligns=[S["left_align"], S["right_align"]])
            elif fmt == "count":
                W.row_values([label, int(value)],
                             formats=[None, S["number_fmt"]],
                             aligns=[S["left_align"], S["right_align"]])
            else:
                W.row_values([label, float(value)],
                             formats=[None, S["pct_fmt"]],
                             aligns=[S["left_align"], S["right_align"]])
        W.blank()

        W.subsection("C. Per Sessional Scheme × Prior Legacy Bucket (Unique Customers)")
        W.headers(["Sessional Scheme",
                   "e-Gold only (count)", "e-Gold only (₹)",
                   "e-Silver only (count)", "e-Silver only (₹)",
                   "Both (count)", "Both (₹)", "New (₹)"])
        for row in overlap_report["per_scheme_rows"]:
            W.row_values([
                row["Scheme"],
                int(row["Existing — e-Gold only — Enrollees"]),
                float(row["Existing — e-Gold only — Amount"]),
                int(row["Existing — e-Silver only — Enrollees"]),
                float(row["Existing — e-Silver only — Amount"]),
                int(row["Existing — Both — Enrollees"]),
                float(row["Existing — Both — Amount"]),
                float(row["New — Amount"]),
            ], formats=[None, S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["currency_fmt"],
                        S["number_fmt"], S["currency_fmt"], S["currency_fmt"]],
               aligns=[S["left_align"]] + [S["right_align"]] * 7)
        W.row_values([
            "TOTAL (unique customers across all schemes)",
            int(sum(r["Existing — e-Gold only — Enrollees"]
                    for r in overlap_report["per_scheme_rows"])),
            float(sum(r["Existing — e-Gold only — Amount"]
                      for r in overlap_report["per_scheme_rows"])),
            int(sum(r["Existing — e-Silver only — Enrollees"]
                    for r in overlap_report["per_scheme_rows"])),
            float(sum(r["Existing — e-Silver only — Amount"]
                      for r in overlap_report["per_scheme_rows"])),
            int(sum(r["Existing — Both — Enrollees"]
                    for r in overlap_report["per_scheme_rows"])),
            float(sum(r["Existing — Both — Amount"]
                      for r in overlap_report["per_scheme_rows"])),
            float(sum(r["New — Amount"]
                      for r in overlap_report["per_scheme_rows"])),
        ], formats=[None, S["number_fmt"], S["currency_fmt"],
                    S["number_fmt"], S["currency_fmt"],
                    S["number_fmt"], S["currency_fmt"], S["currency_fmt"]],
           aligns=[S["left_align"]] + [S["right_align"]] * 7, total=True)
        W.blank()

        W.subsection("D. Legacy Activity Change of the Overlap Group")
        W.headers(["Metric", "Value"])
        pa = overlap_report["pre_activity_of_overlap"]
        for label, value in [
            ("Legacy amount pre-launch", pa["pre_total_legacy_amount"]),
            ("Legacy amount post-launch", pa["post_total_legacy_amount"]),
            ("Change", pa["legacy_amount_change"]),
        ]:
            W.row_values([label, float(value)],
                         formats=[None, S["currency_fmt"]],
                         aligns=[S["left_align"], S["right_align"]])
        W.blank()

        def write_customer_list(header, subset):
            W.subsection(header)
            if subset.empty:
                W.note("No customers in this bucket.")
                W.blank()
                return
            W.headers(["Customer ID", "Sessional Scheme", "Sess. Count",
                       "Sess. Amount", "First Sess.",
                       "Pre-Legacy Amount", "Pre-Legacy Schemes",
                       "Last Legacy Date"])
            for _, r in subset.sort_values(
                "Sessional_Enrol_Amount", ascending=False).iterrows():
                fs = r.get("First_Sessional_Date")
                ll = r.get("Last_Legacy_Date")
                W.row_values([
                    str(r.get("Customer_ID", "")),
                    str(r.get("Sessional_Scheme", "")),
                    int(r.get("Sessional_Enrol_Count", 0)),
                    float(r.get("Sessional_Enrol_Amount", 0)),
                    fs.strftime("%d-%b-%Y") if pd.notna(fs) else "",
                    float(r.get("Pre_Legacy_Amount", 0) or 0),
                    str(r.get("Pre_Legacy_Schemes", "")),
                    ll.strftime("%d-%b-%Y") if pd.notna(ll) else "",
                ], formats=[None, None, S["number_fmt"], S["currency_fmt"],
                            None, S["currency_fmt"], None, None],
                   aligns=[S["left_align"], S["left_align"],
                           S["right_align"], S["right_align"],
                           S["center_align"], S["right_align"],
                           S["left_align"], S["center_align"]])
            W.blank()

        ov = overlap_report["overlap_df"]
        write_customer_list("E1. Existing — e-Gold only",
            ov[ov["Legacy_Bucket"] == "Existing — e-Gold only"])
        write_customer_list("E2. Existing — e-Silver only",
            ov[ov["Legacy_Bucket"] == "Existing — e-Silver only"])
        write_customer_list("E3. Existing — Both e-Gold & e-Silver",
            ov[ov["Legacy_Bucket"] == "Existing — Both"])

        W.subsection("F. New — No Prior Legacy Activity")
        nw = overlap_report["new_df"]
        if nw.empty:
            W.note("No new customers found.")
        else:
            W.headers(["Customer ID", "Sessional Scheme", "Sess. Count",
                       "Sess. Amount", "First Sess."])
            for _, r in nw.sort_values(
                "Sessional_Enrol_Amount", ascending=False).iterrows():
                fs = r.get("First_Sessional_Date")
                W.row_values([
                    str(r.get("Customer_ID", "")),
                    str(r.get("Sessional_Scheme", "")),
                    int(r.get("Sessional_Enrol_Count", 0)),
                    float(r.get("Sessional_Enrol_Amount", 0)),
                    fs.strftime("%d-%b-%Y") if pd.notna(fs) else "",
                ], formats=[None, None, S["number_fmt"], S["currency_fmt"],
                            None],
                   aligns=[S["left_align"], S["left_align"],
                           S["right_align"], S["right_align"],
                           S["center_align"]])
        W.blank()

    ws.column_dimensions["A"].width = 38
    for col in ["B", "C", "D", "E", "F", "G", "H", "I"]:
        ws.column_dimensions[col].width = 22
    ws.freeze_panes = "A3"
    return ws


# ============================================================
# STREAMLIT APP
# ============================================================

st.title("📊 E-Gold / E-Silver / Sessional — Overall Impact Report")
st.markdown("### Single-sheet Excel report — every section uses the LIVE window")

uploaded_file = st.file_uploader(
    "Upload your transaction Excel / CSV file",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is None:
    st.info("Upload your transaction file to start.")
    st.stop()

try:
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

df.columns = [clean_column_name(c) for c in df.columns]

try:
    date_col = find_column(df, ["Date", "Transaction Date", "Created Date"])
    installment_col = find_column(df, ["Installment number", "Installment Number",
                                        "Installment No", "Installment"])
    scheme_col = find_column(df, ["Scheme Name", "Scheme"])
    amount_col = find_column(df, ["Saved Amount", "Amount",
                                   "Paid Amount", "Transaction Amount"])
    id_col = find_column(df, ["Id", "ID", "Transaction ID"], required=False)
except Exception as e:
    st.error(str(e))
    st.write("### Columns detected in your Excel")
    st.write(list(df.columns))
    st.stop()

if id_col is None:
    df["Id"] = range(1, len(df) + 1)
    id_col = "Id"


# ============================================================
# CUSTOMER ID LOOKUP
# ============================================================

customer_id_col = find_column(
    df,
    [
        "Customer ID", "Customer Id", "CustomerID", "CIF",
        "Customer Code", "Member ID", "Member Id", "Member Code",
        "Cust ID", "Cust Id", "Cust Code",
        "Client ID", "Client Id", "Client Code",
        "User ID", "User Id", "User Code",
        "Phone Number", "Phone No", "PhoneNo", "Phone",
        "Mobile Number", "Mobile No", "MobileNo", "Mobile",
        "Contact Number", "Contact No", "ContactNo", "Contact",
        "WhatsApp Number", "WhatsApp No", "WhatsApp",
        "Cell Number", "Cell No", "Cell",
        "Primary Mobile", "Primary Phone",
        "Registered Mobile", "Registered Phone",
        "Customer Mobile", "Customer Phone",
        "Contact Mobile", "Contact Phone",
        "Aadhaar", "Aadhar", "Aadhaar Number", "Aadhar Number",
        "PAN", "PAN Number", "PAN No",
        "Email", "Email ID", "Email Id",
    ],
    required=False,
)

if customer_id_col is None:
    for col in df.columns:
        c = str(col).strip().lower().replace("_", " ").replace("-", " ")
        if any(w in c for w in ["phone", "mobile", "contact",
                                 "whatsapp", "cell"]):
            customer_id_col = col
            break

HAS_CUSTOMER_ID = customer_id_col is not None

if HAS_CUSTOMER_ID:
    def _norm_cid(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        digits = re.sub(r"\D", "", s)
        if len(digits) >= 7:
            return digits
        return s.lower()

    df["Customer_ID"] = df[customer_id_col].apply(_norm_cid)
    st.sidebar.success(f"🧬 Customer ID detected: `{customer_id_col}`")
else:
    df["Customer_ID"] = None
    st.sidebar.error(
        "⚠️ No customer identifier column found. "
        "Rename your phone/mobile column to **Phone Number** or **Mobile** "
        "to enable Section 16."
    )


# ============================================================
# PREPARE ANALYSIS DATA
# ============================================================

df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
df = df[df[date_col].notna()].copy()
df[amount_col] = safe_numeric(df[amount_col])

df["Actual Scheme Name"] = df[scheme_col].fillna("Unknown").astype(str).str.strip()
df["Scheme Category"] = df["Actual Scheme Name"].apply(classify_scheme)
df = classify_transaction(df, installment_col)

analysis_df = df[df["Transaction Type"].isin(["Enrolment", "Collection"])].copy()
analysis_df["Saved Amount"] = analysis_df[amount_col]
analysis_df["Date"] = analysis_df[date_col].dt.normalize()
analysis_df["Month"] = analysis_df["Date"].dt.to_period("M").astype(str)
analysis_df["Month Date"] = analysis_df["Date"].dt.to_period("M").dt.to_timestamp()

min_date = analysis_df["Date"].min().date()
max_date = analysis_df["Date"].max().date()
last_report_live_date = pd.Timestamp(analysis_df["Date"].max())

available_months_all = sorted(analysis_df["Month Date"].unique())
if available_months_all:
    latest_month_ts_all = pd.Timestamp(available_months_all[-1])
    REFERENCE_MONTH = (REPORT_ANCHOR_MONTH
                       if latest_month_ts_all >= REPORT_ANCHOR_MONTH
                       else latest_month_ts_all)
else:
    REFERENCE_MONTH = REPORT_ANCHOR_MONTH

ref_month_df = analysis_df[analysis_df["Month Date"] == REFERENCE_MONTH]
if not ref_month_df.empty:
    DAY_START = int(ref_month_df["Date"].dt.day.min())
    DAY_END = int(ref_month_df["Date"].dt.day.max())
else:
    DAY_START, DAY_END = 1, DAY_END_FALLBACK
if DAY_END < DAY_START:
    DAY_END = DAY_END_FALLBACK

DAY_WINDOW_LABEL = f"Day {DAY_START}–{DAY_END}"
DAY_WINDOW_DAYS = DAY_END - DAY_START + 1
REFERENCE_MONTH_LABEL = REFERENCE_MONTH.strftime("%b %Y")

sessional_live_all = analysis_df[
    (analysis_df["Scheme Category"] == CATEGORY_SESSIONAL) &
    (analysis_df["Date"] >= SESSIONAL_LAUNCH_DATE)]
if not sessional_live_all.empty:
    SESSIONAL_LIVE_START = pd.Timestamp(sessional_live_all["Date"].min())
else:
    SESSIONAL_LIVE_START = pd.Timestamp(
        year=REFERENCE_MONTH.year, month=REFERENCE_MONTH.month,
        day=min(DAY_END, SESSIONAL_AVAILABLE_DATE_FALLBACK.day))
SESSIONAL_LIVE_END = pd.Timestamp(analysis_df["Date"].max())


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Filters")

date_range = st.sidebar.date_input(
    "Analysis Date Range",
    value=(min_date, max_date),
    min_value=min_date, max_value=max_date)
if isinstance(date_range, tuple):
    start_date, end_date = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
else:
    start_date = pd.Timestamp(date_range)
    end_date = start_date

filtered_df = analysis_df[
    (analysis_df["Date"] >= start_date) & (analysis_df["Date"] <= end_date)].copy()

st.sidebar.markdown("---")
st.sidebar.header("🚀 Sessional Timeline")
st.sidebar.success(f"Launch: {SESSIONAL_LAUNCH_DATE.strftime('%d-%b-%Y')}")
st.sidebar.info(f"Sessional Live: {SESSIONAL_LIVE_START.strftime('%d-%b-%Y')} → "
                f"{SESSIONAL_LIVE_END.strftime('%d-%b-%Y')}")
st.sidebar.caption(f"Last report live: {last_report_live_date.strftime('%d-%b-%Y')}")
st.sidebar.warning(f"Matched Window: {DAY_WINDOW_LABEL} ({DAY_WINDOW_DAYS} days)")
st.sidebar.success(f"🔒 Live window locked: {DAY_WINDOW_LABEL} — ALL sections")
st.sidebar.markdown("---")

selected_categories = st.sidebar.multiselect(
    "Scheme Category",
    options=[CATEGORY_GOLD, CATEGORY_SILVER, CATEGORY_SESSIONAL],
    default=[CATEGORY_GOLD, CATEGORY_SILVER, CATEGORY_SESSIONAL])
if selected_categories:
    filtered_df = filtered_df[
        filtered_df["Scheme Category"].isin(selected_categories)].copy()

sessional_scheme_names = sorted(
    filtered_df.loc[filtered_df["Scheme Category"] == CATEGORY_SESSIONAL,
                    "Actual Scheme Name"].dropna().unique().tolist())
if sessional_scheme_names:
    st.sidebar.markdown("**Actual Sessional Scheme Names**")
    selected_sessional_schemes = st.sidebar.multiselect(
        "Select Sessional Schemes",
        options=sessional_scheme_names,
        default=sessional_scheme_names)
    if selected_sessional_schemes:
        filtered_df = filtered_df[
            (filtered_df["Scheme Category"] != CATEGORY_SESSIONAL) |
            (filtered_df["Actual Scheme Name"].isin(selected_sessional_schemes))
        ].copy()


# ============================================================
# LIVE BANNER
# ============================================================

st.caption(f"⚠️ Every section uses the live window **{DAY_WINDOW_LABEL}** of "
           f"the reference month ({REFERENCE_MONTH_LABEL}).")


# ============================================================
# BASE AGGREGATIONS
# ============================================================

enrolment_df = filtered_df[filtered_df["Transaction Type"] == "Enrolment"]
collection_df = filtered_df[filtered_df["Transaction Type"] == "Collection"]
total_enrolments = len(enrolment_df)
total_collections = len(collection_df)
enrolment_amount = enrolment_df["Saved Amount"].sum()
collection_amount = collection_df["Saved Amount"].sum()
total_business = enrolment_amount + collection_amount
enrolment_avg_ticket = enrolment_amount / total_enrolments if total_enrolments else 0
collection_avg_ticket = collection_amount / total_collections if total_collections else 0

st.header("📌 Executive Summary")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Enrolments", f"{total_enrolments:,}")
c2.metric("Enrolment Amount", f"₹{enrolment_amount:,.0f}")
c3.metric("Collections", f"{total_collections:,}")
c4.metric("Collection Amount", f"₹{collection_amount:,.0f}")
c5, c6, c7, c8 = st.columns(4)
c5.metric("Enrol. Avg Ticket", f"₹{enrolment_avg_ticket:,.0f}")
c6.metric("Coll. Avg Ticket", f"₹{collection_avg_ticket:,.0f}")
c7.metric("Total Business", f"₹{total_business:,.0f}")
c8.metric("Total Txns", f"{total_enrolments + total_collections:,}")


# ============================================================
# CATEGORY SUMMARY
# ============================================================

category_summary = (
    filtered_df.groupby(["Scheme Category", "Transaction Type"], dropna=False)
    .agg(Count=(id_col, "count"), Amount=("Saved Amount", "sum"))
    .reset_index())
category_summary["Average Ticket"] = (
    category_summary["Amount"] / category_summary["Count"].replace(0, np.nan))

st.header("🏷️ Category Performance")
st.dataframe(category_summary, use_container_width=True, hide_index=True)


# ============================================================
# MONTHLY ANALYSIS
# ============================================================

monthly_analysis = build_monthly_matched_table(
    filtered_df, id_col, REFERENCE_MONTH, DAY_START, DAY_END)

st.header(f"📈 Monthly Analysis — Matched Live Window ({DAY_WINDOW_LABEL})")
display_monthly = monthly_analysis.drop(columns=["_is_live"], errors="ignore")
st.dataframe(display_monthly, use_container_width=True, hide_index=True)

if not monthly_analysis.empty:
    chart_data = monthly_analysis[
        ["Month", "e-Gold Enrolment Amount",
         "e-Silver Enrolment Amount", "Sessional Enrolment Amount"]
    ].set_index("Month")
    st.line_chart(chart_data)


# ============================================================
# PREV vs CURR
# ============================================================

st.header(f"🔄 Previous vs Current Month (both {DAY_WINDOW_LABEL})")
curr_month_ts = REFERENCE_MONTH
prev_month_ts = REFERENCE_MONTH - pd.DateOffset(months=1)
prev_month_label = prev_month_ts.strftime("%b %Y")
curr_month_label = REFERENCE_MONTH.strftime("%b %Y")


def _enrol_in_window(df, category, year, month):
    s, e = _month_window(year, month, DAY_START, DAY_END)
    sub = df[(df["Scheme Category"] == category) &
             (df["Transaction Type"] == "Enrolment") &
             (df["Date"] >= s) & (df["Date"] <= e)]
    cnt, amt = len(sub), sub["Saved Amount"].sum()
    return cnt, amt, (amt / cnt if cnt else 0)


comparison_rows = []
for cat in [CATEGORY_GOLD, CATEGORY_SILVER]:
    pc, pa, pt = _enrol_in_window(filtered_df, cat, prev_month_ts.year, prev_month_ts.month)
    cc, ca, ct = _enrol_in_window(filtered_df, cat, curr_month_ts.year, curr_month_ts.month)
    comparison_rows.append({
        "Scheme Category": cat,
        "Previous Month Enrol. Count": pc, "Previous Month Enrol. Amount": pa,
        "Current Month Enrol. Count": cc, "Current Month Enrol. Amount": ca,
        "Count Change": cc - pc, "Count Change %": pct_change(pc, cc),
        "Amount Change": ca - pa, "Amount Change %": pct_change(pa, ca),
    })
gold_silver_comparison = pd.DataFrame(comparison_rows)
if not gold_silver_comparison.empty:
    st.dataframe(gold_silver_comparison, use_container_width=True, hide_index=True)


# ============================================================
# SESSIONAL SCHEME DETAIL
# ============================================================

st.header("🟣 Sessional Scheme Detail")
sessional_df = filtered_df[filtered_df["Scheme Category"] == CATEGORY_SESSIONAL].copy()
if sessional_df.empty:
    sessional_scheme_table = pd.DataFrame()
else:
    sessional_scheme_table = (
        sessional_df.groupby(["Month", "Actual Scheme Name", "Transaction Type"],
                             dropna=False)
        .agg(Count=(id_col, "count"), Amount=("Saved Amount", "sum"))
        .reset_index())
    sessional_scheme_table["Average Ticket"] = (
        sessional_scheme_table["Amount"] /
        sessional_scheme_table["Count"].replace(0, np.nan))
    sessional_scheme_table = sessional_scheme_table.sort_values(
        ["Month", "Actual Scheme Name", "Transaction Type"])
    st.dataframe(sessional_scheme_table, use_container_width=True, hide_index=True)


# ============================================================
# PRE vs POST
# ============================================================

st.header(f"🔄 Pre vs Post Sessional (both clipped to {DAY_WINDOW_LABEL})")
pre_df = filtered_df[filtered_df["Date"] < SESSIONAL_LAUNCH_DATE].copy()
post_df = filtered_df[filtered_df["Date"] >= SESSIONAL_LAUNCH_DATE].copy()
pre_matched = clip_df_to_matched_window(pre_df, DAY_START, DAY_END)
post_matched = clip_df_to_matched_window(post_df, DAY_START, DAY_END)


def matched_avg(data, category, ttype, metric):
    temp = data[(data["Scheme Category"] == category) &
                (data["Transaction Type"] == ttype)].copy()
    if temp.empty:
        return 0
    monthly = temp.groupby("Month").agg(
        Count=(id_col, "count"), Amount=("Saved Amount", "sum"))
    monthly["Average Ticket"] = monthly["Amount"] / monthly["Count"].replace(0, np.nan)
    if metric == "Count":           return monthly["Count"].mean()
    if metric == "Amount":          return monthly["Amount"].mean()
    if metric == "Average Ticket":  return monthly["Average Ticket"].mean()
    return 0


pre_post_rows = []
for category in [CATEGORY_GOLD, CATEGORY_SILVER, CATEGORY_SESSIONAL]:
    for ttype in ["Enrolment", "Collection"]:
        pre_count = matched_avg(pre_matched, category, ttype, "Count")
        post_count = matched_avg(post_matched, category, ttype, "Count")
        pre_amount = matched_avg(pre_matched, category, ttype, "Amount")
        post_amount = matched_avg(post_matched, category, ttype, "Amount")
        pre_ticket = matched_avg(pre_matched, category, ttype, "Average Ticket")
        post_ticket = matched_avg(post_matched, category, ttype, "Average Ticket")
        pre_post_rows.append({
            "Scheme Category": category, "Transaction Type": ttype,
            "Pre Monthly Avg Count": pre_count, "Post Monthly Avg Count": post_count,
            "Count Change": post_count - pre_count,
            "Count Change %": pct_change(pre_count, post_count),
            "Pre Monthly Avg Amount": pre_amount, "Post Monthly Avg Amount": post_amount,
            "Amount Change": post_amount - pre_amount,
            "Amount Change %": pct_change(pre_amount, post_amount),
            "Pre Avg Ticket": pre_ticket, "Post Avg Ticket": post_ticket,
            "Average Ticket Change %": pct_change(pre_ticket, post_ticket)})
pre_post_analysis = pd.DataFrame(pre_post_rows)
st.dataframe(pre_post_analysis, use_container_width=True, hide_index=True)


# ============================================================
# SCHEME-NAME IMPACT
# ============================================================

st.header("🔬 Scheme-Name Level Impact (Named Schemes Only)")
scheme_impact_df = build_scheme_impact_analysis(
    filtered_df, SESSIONAL_LAUNCH_DATE, id_col,
    REFERENCE_MONTH, DAY_START, DAY_END)
if not scheme_impact_df.empty:
    st.dataframe(scheme_impact_df, use_container_width=True, hide_index=True)
else:
    st.info("No named schemes found.")


# ============================================================
# SHARE BY MONTH
# ============================================================

st.header(f"📊 Scheme Share by Month — clipped to {DAY_WINDOW_LABEL}")
scheme_share_df = build_scheme_share_by_month(filtered_df, id_col, DAY_START, DAY_END)
if not scheme_share_df.empty:
    st.dataframe(scheme_share_df, use_container_width=True, hide_index=True)
    st.area_chart(scheme_share_df.set_index("Month"))


# ============================================================
# MATCHED WINDOW COMPARISONS
# ============================================================

base_year, base_month = REFERENCE_MONTH.year, REFERENCE_MONTH.month
gold_window_comparison = build_window_list(
    filtered_df, CATEGORY_GOLD, base_year, base_month, DAY_START, DAY_END)
silver_window_comparison = build_window_list(
    filtered_df, CATEGORY_SILVER, base_year, base_month, DAY_START, DAY_END)
combined_window_comparison = build_combined_window_list(
    filtered_df, [CATEGORY_GOLD, CATEGORY_SILVER],
    base_year, base_month, DAY_START, DAY_END)

sessional_live_df = filtered_df[
    (filtered_df["Date"] >= SESSIONAL_LIVE_START) &
    (filtered_df["Date"] <= SESSIONAL_LIVE_END) &
    (filtered_df["Scheme Category"] == CATEGORY_SESSIONAL)]
sessional_details = []
for ttype in ["Enrolment", "Collection"]:
    sub = sessional_live_df[sessional_live_df["Transaction Type"] == ttype]
    cnt, amt = len(sub), sub["Saved Amount"].sum()
    sessional_details.append({"Scheme Category": CATEGORY_SESSIONAL,
        "Transaction Type": ttype, "Count": cnt, "Amount": amt,
        "Average Ticket": amt / cnt if cnt else 0})
sessional_first_days = {
    "label": f"Sessional Live ({SESSIONAL_LIVE_START.strftime('%d-%b')} → "
             f"{SESSIONAL_LIVE_END.strftime('%d-%b')})",
    "start": SESSIONAL_LIVE_START, "end": SESSIONAL_LIVE_END,
    "days": ((SESSIONAL_LIVE_END - SESSIONAL_LIVE_START).days + 1
             if SESSIONAL_LIVE_END >= SESSIONAL_LIVE_START else 0),
    "details": sessional_details}


# ============================================================
# CANNIBALISATION
# ============================================================

st.header("🧠 Sessional Impact on e-Gold / e-Silver — Matched Window")


def get_pre_post(category, ttype):
    result = pre_post_analysis[
        (pre_post_analysis["Scheme Category"] == category) &
        (pre_post_analysis["Transaction Type"] == ttype)]
    if result.empty:
        return 0, 0
    return (float(result.iloc[0]["Pre Monthly Avg Amount"]),
            float(result.iloc[0]["Post Monthly Avg Amount"]))


_, sessional_post = get_pre_post(CATEGORY_SESSIONAL, "Enrolment")

if len(combined_window_comparison) >= 4:
    matched_pre_avg = np.mean([w["Enrol. Amount"]
                               for w in combined_window_comparison[:-1]])
    matched_post = combined_window_comparison[-1]["Enrol. Amount"]
else:
    matched_pre_avg, matched_post = 0.0, 0.0

existing_pre = matched_pre_avg
existing_post = matched_post
existing_scheme_loss = max(existing_pre - existing_post, 0)
sessional_gain = sessional_post
net_business_change = existing_post + sessional_gain - existing_pre

c1, c2, c3, c4 = st.columns(4)
c1.metric("e-Gold+e-Silver Pre (matched avg)", f"₹{existing_pre:,.0f}")
c2.metric("e-Gold+e-Silver Post (matched)", f"₹{existing_post:,.0f}")
c3.metric("Sessional Post (matched)", f"₹{sessional_gain:,.0f}")
c4.metric("Net Business Change", f"₹{net_business_change:,.0f}")

if post_matched.empty:
    interpretation = "⚪ No Post-Sessional data in the live window."
elif sessional_gain == 0:
    interpretation = "⚪ No measurable Sessional enrolment amount after launch."
elif existing_post >= existing_pre:
    interpretation = "🟢 Incremental — Sessional grew while e-Gold/e-Silver did not decline."
elif sessional_gain > existing_scheme_loss:
    interpretation = "🟡 Mostly incremental — e-Gold/e-Silver declined but Sessional exceeds it."
elif sessional_gain == existing_scheme_loss:
    interpretation = "🟠 Possible replacement — Sessional ≈ e-Gold/e-Silver decline."
else:
    interpretation = "🔴 Possible cannibalisation — decline exceeds Sessional gain."
st.info(interpretation)


# ============================================================
# MATCHED COMPARISON DISPLAY
# ============================================================

if len(gold_window_comparison) >= 2:
    gold_base_avg = np.mean([w["Enrol. Amount"] for w in gold_window_comparison[:-1]])
    silver_base_avg = np.mean([w["Enrol. Amount"] for w in silver_window_comparison[:-1]])
    gold_curr = gold_window_comparison[-1]["Enrol. Amount"]
    silver_curr = silver_window_comparison[-1]["Enrol. Amount"]
    gold_chg_pct = pct_change(gold_base_avg, gold_curr)
    silver_chg_pct = pct_change(silver_base_avg, silver_curr)
    sessional_amt = sum(d["Amount"] for d in sessional_details)

    if sessional_amt == 0:
        sessional_interpretation = "⚪ No Sessional amount in the live window."
    elif (gold_curr + silver_curr) >= (gold_base_avg + silver_base_avg):
        sessional_interpretation = (
            f"🟢 Incremental. e-Gold {gold_chg_pct:+.1f}%, "
            f"e-Silver {silver_chg_pct:+.1f}%. Sessional added ₹{sessional_amt:,.0f}.")
    else:
        combined_loss = (gold_base_avg + silver_base_avg) - (gold_curr + silver_curr)
        if sessional_amt > combined_loss:
            sessional_interpretation = (
                f"🟡 Mostly incremental. Decline ₹{combined_loss:,.0f} vs "
                f"Sessional gain ₹{sessional_amt:,.0f}.")
        elif sessional_amt == combined_loss:
            sessional_interpretation = (
                f"🟠 Possible replacement (₹{sessional_amt:,.0f}).")
        else:
            sessional_interpretation = (
                f"🔴 Possible cannibalisation. Decline ₹{combined_loss:,.0f} "
                f"exceeds Sessional gain ₹{sessional_amt:,.0f}.")
else:
    sessional_interpretation = "⚪ Insufficient monthly data."
st.info(sessional_interpretation)


# ============================================================
# SECTION 16 — SESSIONAL OVERLAP
# ============================================================

st.header("🧬 Sessional Enrollees — Prior e-Gold / e-Silver Breakdown")
st.caption("Splits Sessional enrollees into 4 buckets: e-Gold only, "
           "e-Silver only, Both, and New.")

overlap_report = None
if not HAS_CUSTOMER_ID:
    st.warning("⚠️ No customer identifier column found. Section 16 skipped.")
else:
    overlap_report = build_sessional_overlap_report(filtered_df, SESSIONAL_LAUNCH_DATE)
    if overlap_report is None:
        st.info("No Sessional enrolments in the selected date range.")
    else:
        bt = overlap_report["bucket_totals"]
        st.subheader("A. Sessional Enrollees — Split by Prior Legacy")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("e-Gold only",
                  f"{bt['Existing — e-Gold only']['count']:,}",
                  f"₹{bt['Existing — e-Gold only']['amount']:,.0f}")
        c2.metric("e-Silver only",
                  f"{bt['Existing — e-Silver only']['count']:,}",
                  f"₹{bt['Existing — e-Silver only']['amount']:,.0f}")
        c3.metric("Both",
                  f"{bt['Existing — Both']['count']:,}",
                  f"₹{bt['Existing — Both']['amount']:,.0f}")
        c4.metric("New",
                  f"{bt['New']['count']:,}",
                  f"₹{bt['New']['amount']:,.0f}")

        st.subheader("B. Per Sessional Scheme × Prior Legacy Bucket")
        st.dataframe(pd.DataFrame(overlap_report["per_scheme_rows"]),
                     use_container_width=True, hide_index=True)

        st.subheader("C. Legacy Activity Change of Overlap Group")
        pa = overlap_report["pre_activity_of_overlap"]
        c1, c2, c3 = st.columns(3)
        c1.metric("Pre-launch", f"₹{pa['pre_total_legacy_amount']:,.0f}")
        c2.metric("Post-launch", f"₹{pa['post_total_legacy_amount']:,.0f}")
        c3.metric("Change", f"₹{pa['legacy_amount_change']:,.0f}")

        st.subheader("D. Customer Lists")
        tab1, tab2, tab3, tab4 = st.tabs(["e-Gold only", "e-Silver only", "Both", "New"])
        ov = overlap_report["overlap_df"]
        with tab1:
            st.dataframe(ov[ov["Legacy_Bucket"] == "Existing — e-Gold only"],
                         use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(ov[ov["Legacy_Bucket"] == "Existing — e-Silver only"],
                         use_container_width=True, hide_index=True)
        with tab3:
            st.dataframe(ov[ov["Legacy_Bucket"] == "Existing — Both"],
                         use_container_width=True, hide_index=True)
        with tab4:
            st.dataframe(overlap_report["new_df"],
                         use_container_width=True, hide_index=True)


# ============================================================
# BUILD WORKBOOK
# ============================================================

totals_block = [
    ("Total Enrolments", total_enrolments, enrolment_amount, enrolment_avg_ticket),
    ("Total Collections", total_collections, collection_amount, collection_avg_ticket),
    ("Total Business", total_enrolments + total_collections, total_business,
     (total_business / (total_enrolments + total_collections))
     if (total_enrolments + total_collections) else 0),
]

cannibalisation_rows = [
    ("e-Gold + e-Silver Pre (matched avg)", existing_pre, '₹#,##0.00'),
    ("e-Gold + e-Silver Post (matched)", existing_post, '₹#,##0.00'),
    ("Sessional Post (matched)", sessional_gain, '₹#,##0.00'),
    ("Observed Existing Scheme Loss", existing_scheme_loss, '₹#,##0.00'),
    ("Observed Net Business Change", net_business_change, '₹#,##0.00'),
]

insights = [interpretation]
if overlap_report is not None:
    bt = overlap_report["bucket_totals"]
    existing_total = (bt["Existing — e-Gold only"]["count"]
                      + bt["Existing — e-Silver only"]["count"]
                      + bt["Existing — Both"]["count"])
    total = existing_total + bt["New"]["count"]
    pct_existing = (existing_total / total * 100) if total else 0
    flag = "🔴 HIGH" if pct_existing >= 70 else ("🟡 MODERATE" if pct_existing >= 40 else "🟢 LOW")
    insights.append(
        f"{flag} overlap — {pct_existing:.0f}% of Sessional enrollees had prior "
        f"legacy activity: e-Gold only {bt['Existing — e-Gold only']['count']}, "
        f"e-Silver only {bt['Existing — e-Silver only']['count']}, "
        f"Both {bt['Existing — Both']['count']}, New {bt['New']['count']}.")
if not scheme_impact_df.empty:
    best = scheme_impact_df.sort_values("Amount Change %", ascending=False).iloc[0]
    worst = scheme_impact_df.sort_values("Amount Change %").iloc[0]
    if best["Scheme Name"] != worst["Scheme Name"]:
        insights.append(f"✅ Best named scheme: {best['Scheme Name']} "
                        f"({best['Amount Change %']:+.1f}%).")
        insights.append(f"⚠️ Weakest named scheme: {worst['Scheme Name']} "
                        f"({worst['Amount Change %']:+.1f}%).")

wb = Workbook()
wb.remove(wb.active)

build_single_overall_sheet(
    wb, filtered_df, id_col, start_date, end_date,
    totals_block, category_summary, monthly_analysis,
    pre_post_analysis, gold_silver_comparison,
    gold_window_comparison, silver_window_comparison,
    combined_window_comparison, sessional_first_days,
    sessional_scheme_table, scheme_impact_df, scheme_share_df,
    cannibalisation_rows, insights, interpretation,
    sessional_interpretation,
    SESSIONAL_LAUNCH_DATE, SESSIONAL_LIVE_START, SESSIONAL_LIVE_END,
    last_report_live_date,
    DAY_START, DAY_END,
    REFERENCE_MONTH_LABEL, DAY_WINDOW_LABEL, DAY_WINDOW_DAYS,
    overlap_report=overlap_report,
)

output = io.BytesIO()
wb.save(output)
output.seek(0)
excel_file = output.getvalue()


# ============================================================
# DOWNLOAD
# ============================================================

st.header("📥 Download Excel Report")
st.download_button(
    label="📥 Download Single-Sheet Formatted Report",
    data=excel_file,
    file_name="E-Gold_Silver_Sessional_Overall_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.sidebar.markdown("---")
st.sidebar.subheader("📌 Data Information")
st.sidebar.write(f"Rows loaded: {len(df):,}")
st.sidebar.write(f"Rows analysed: {len(filtered_df):,}")
st.sidebar.write(f"Data range: {min_date} → {max_date}")
st.sidebar.write(f"Sessional schemes: {len(sessional_scheme_names)}")
st.sidebar.write(f"Matched Window: {DAY_WINDOW_LABEL} ({DAY_WINDOW_DAYS} days)")
st.sidebar.write(f"Reference Month: {REFERENCE_MONTH_LABEL}")
st.sidebar.write(f"Customer ID column: {customer_id_col if HAS_CUSTOMER_ID else 'NOT FOUND'}")