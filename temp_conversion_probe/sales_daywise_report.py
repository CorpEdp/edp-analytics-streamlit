"""
Auto-converted from: ERP-SalesDaywiseReport.py
Review this file before using — the converter does a best-effort wrap;
double-check indentation around any unusual control flow (loops, if/else
blocks that span large sections, etc.).
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
def get_branch_display_name(code):
    """
    Get display name for branch code
    """
    # If code is already a full name, return it
    if code in BRANCH_DISPLAY_NAMES.values():
        return code

    # Try to match code to display name
    if code in BRANCH_DISPLAY_NAMES:
        return BRANCH_DISPLAY_NAMES[code]

    # Try case-insensitive match
    code_upper = code.upper()
    for key, value in BRANCH_DISPLAY_NAMES.items():
        if key.upper() == code_upper:
            return value

    # If no match, return the original code
    return code
def sort_branches(branches):
    branches_list = list(branches)

    return sorted(
        branches_list,
        key=lambda x: (
            BRANCH_ORDER.index(x)
            if x in BRANCH_ORDER
            else len(BRANCH_ORDER)
        )
    )
def clean_numeric_series(series):
    """
    Converts values such as:
    11,770.00
    ₹50,751.71
    3.000
    blank
    into numeric values.
    """

    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace(["", "nan", "None", "NaN", "null"], "0")
    )

    return pd.to_numeric(cleaned, errors="coerce").fillna(0)
def _has_code_token(label_str, codes):
    """
    True if any code in `codes` (e.g. 'V-', 'R-') appears as a
    standalone token: at the start of the label, or immediately
    after a space/slash/comma boundary - never as a mid-word
    substring.
    """
    for code in codes:
        pattern = r'(?:^|[\s/,]){}'.format(re.escape(code))
        if re.search(pattern, label_str):
            return True
    return False
def detect_item_from_label(label):

    if pd.isna(label):
        return "Unknown"

    label_str = str(label).upper().strip()

    # --------------------------------------------------------
    # IMPORTANT:
    # Check Repair first
    # --------------------------------------------------------

    if "REPAIR" in label_str:
        return "Skip"

    # --------------------------------------------------------
    # Silver
    # --------------------------------------------------------

    if (
        "SILVER" in label_str
        or _has_code_token(label_str, ["V-"])
    ):
        return "Silver"

    # --------------------------------------------------------
    # Diamond
    # --------------------------------------------------------

    if (
        "DIAMOND" in label_str
        or _has_code_token(label_str, ["R-"])
    ):
        return "Diamond"

    # --------------------------------------------------------
    # Platinum
    # --------------------------------------------------------

    if (
        "PLATINUM" in label_str
        or _has_code_token(label_str, ["P-"])
    ):
        return "Platinum"

    # --------------------------------------------------------
    # Gold
    # --------------------------------------------------------

    if (
        "MAIN" in label_str
        or "ANTIQUE" in label_str
        or "DESIGNER" in label_str
        or "BOUTIQUE" in label_str
        or "COUNTER" in label_str
        or _has_code_token(label_str, ["M-", "A-", "D-", "B-"])
    ):
        return "Gold"

    # ========================================================
    # FALLBACK: safe keyword-only checks
    # (loose single-letter-plus-space checks removed - see note above)
    # ========================================================

    if "GOLD" in label_str:
        return "Gold"

    if "SILVER" in label_str:
        return "Silver"

    if "PLATINUM" in label_str:
        return "Platinum"

    if "DIAMOND" in label_str:
        return "Diamond"

    # If it has any numeric value or seems like a valid item, classify as Gold
    # This ensures we don't lose data
    if re.search(r'\d', label_str):
        return "Gold"

    return "Unknown"
def extract_branch_code(branch_value):
    """
    Extract branch code from branch name with improved handling
    """
    if pd.isna(branch_value):
        return "N/A"

    branch_str = str(branch_value).strip().upper()

    # Remove "BRANCH" prefix if present
    branch_str = branch_str.replace("BRANCH", "").strip()
    branch_str = " ".join(branch_str.split())

    # Check for exact full name match (single source of truth)
    if branch_str in _FULL_NAME_TO_CODE:
        return _FULL_NAME_TO_CODE[branch_str]

    # Check for partial matches
    for full_name, code in _FULL_NAME_TO_CODE.items():
        if full_name in branch_str or branch_str in full_name:
            return code

    # Check if it's already a valid code
    valid_codes = list(BRANCH_CODE_TO_NAME.keys())

    # Try to extract code from string
    # Look for 2-4 letter codes
    parts = branch_str.split()
    for part in parts:
        # Remove any non-alphabetic characters
        clean_part = re.sub(r"[^A-Z]", "", part)
        if len(clean_part) >= 2:
            # Check if it matches any valid code
            for code in valid_codes:
                if clean_part.startswith(code) or code.startswith(clean_part):
                    return code

    # If we find a 2-4 letter code at the start
    first_part = re.sub(r"[^A-Za-z]", "", branch_str.split()[0] if branch_str.split() else "")
    if len(first_part) >= 2:
        for code in valid_codes:
            if first_part.upper().startswith(code) or code.startswith(first_part.upper()):
                return code

    # Fallback: return first 3 characters if nothing else works
    if len(branch_str) >= 3:
        return branch_str[:3]

    return branch_str[:2] if len(branch_str) >= 2 else branch_str
def fill_blank_branches(df, branch_col):

    if branch_col not in df.columns:
        return df, 0

    df_filled = df.copy()

    original = df_filled[branch_col].copy()

    df_filled[branch_col] = (
        df_filled[branch_col]
        .astype(str)
        .str.strip()
    )

    df_filled[branch_col] = df_filled[branch_col].replace(
        [
            "",
            "nan",
            "None",
            "NaN",
            "null",
        ],
        pd.NA
    )

    # Forward fill
    df_filled[branch_col] = df_filled[branch_col].ffill()

    # Fill remaining blanks from next valid branch
    df_filled[branch_col] = df_filled[branch_col].bfill()

    filled_count = (
        original.astype(str).str.strip()
        != df_filled[branch_col].astype(str).str.strip()
    ).sum()

    return df_filled, int(filled_count)
def auto_detect_columns(df):

    cols = list(df.columns)

    lower = {
        c: str(c).lower().strip()
        for c in cols
    }

    def find(
        include_any=None,
        include_all=None,
        exclude=None,
        exact=None
    ):

        include_any = include_any or []
        include_all = include_all or []
        exclude = exclude or []

        # Exact
        if exact:

            for c in cols:

                if lower[c] == exact:
                    return c

        # ALL required words
        if include_all:

            for c in cols:

                lc = lower[c]

                if (
                    all(tok in lc for tok in include_all)
                    and not any(ex in lc for ex in exclude)
                ):
                    return c

        # ANY keyword
        for c in cols:

            lc = lower[c]

            if (
                any(tok in lc for tok in include_any)
                and not any(ex in lc for ex in exclude)
            ):
                return c

        return None


    # --------------------------------------------------------
    # Date
    # --------------------------------------------------------

    date_col = find(
        exact="date",
        include_any=["date"]
    )


    # --------------------------------------------------------
    # Branch
    # --------------------------------------------------------

    branch_col = find(
        exact="branch location",
        include_all=["branch"],
        exclude=["label"]
    )

    if branch_col is None:

        branch_col = find(
            include_any=["branch"],
            exclude=["label"]
        )


    # --------------------------------------------------------
    # Label
    # --------------------------------------------------------

    label_col = find(
        exact="label location",
        include_any=["label", "counter"]
    )


    # --------------------------------------------------------
    # Item Amount
    # --------------------------------------------------------

    amount_col = find(
        exact="item amount",
        include_all=["item", "amount"]
    )

    if amount_col is None:

        amount_col = find(
            include_all=["taxable", "amount"]
        )

    if amount_col is None:

        amount_col = find(
            include_any=["amount", "taxable"]
        )


    # --------------------------------------------------------
    # Gross Weight
    # --------------------------------------------------------

    gross_wt_col = find(
        exact="gross wt",
        include_all=["gross"],
        exclude=["diamond"]
    )

    if gross_wt_col is None:
        gross_wt_col = find(
            exact="gross weight",
            include_all=["gross"],
            exclude=["diamond"]
        )

    if gross_wt_col is None:

        gross_wt_col = find(
            include_any=["gross"],
            exclude=["diamond"]
        )


    # --------------------------------------------------------
    # Diamond Weight
    # --------------------------------------------------------

    diamond_wt_col = find(
        exact="diamond weight ct",
        include_all=["diamond", "weight"]
    )

    if diamond_wt_col is None:
        diamond_wt_col = find(
            exact="diamond weight (ct)",
            include_all=["diamond", "weight"]
        )

    if diamond_wt_col is None:
        diamond_wt_col = find(
            include_all=["diamond", "weight"]
        )

    if diamond_wt_col is None:

        diamond_wt_col = find(
            include_all=["diamond", "wt"]
        )

    if diamond_wt_col is None:
        diamond_wt_col = find(
            include_all=["stone", "weight"]
        )

    if diamond_wt_col is None:

        diamond_wt_col = find(
            include_any=["diamond"]
        )


    # --------------------------------------------------------
    # Diamond Amount
    # --------------------------------------------------------

    diamond_amount_col = find(
        exact="diamond amount",
        include_all=["diamond", "amount"]
    )

    if diamond_amount_col is None:
        diamond_amount_col = find(
            exact="stone amount",
            include_all=["stone", "amount"]
        )

    if diamond_amount_col is None:

        diamond_amount_col = find(
            include_all=["diamond", "amount"]
        )

    if diamond_amount_col is None:

        diamond_amount_col = find(
            include_all=["diamond", "amt"]
        )

    if diamond_amount_col is None:
        diamond_amount_col = find(
            include_all=["stone", "amount"]
        )

    if diamond_amount_col is None:
        diamond_amount_col = find(
            include_any=["stone", "amount", "diamond", "amt"]
        )


    # --------------------------------------------------------
    # Fallback
    # --------------------------------------------------------

    fallback = cols[0] if cols else None

    return {

        "date": date_col or fallback,

        "branch": branch_col or fallback,

        "label": label_col or fallback,

        "amount": amount_col or fallback,

        "gross_wt": gross_wt_col or fallback,

        "diamond_wt": diamond_wt_col or fallback,

        "diamond_amount": diamond_amount_col or None,
    }
def build_raw_data_sheet(
    df_filtered,
    branch_col,
    label_col,
    amount_col,
    gross_wt_col,
    diamond_wt_col,
    diamond_amount_col,
    item_col=None  # kept for compatibility; not used in the export layout
):
    """
    Build a raw export sheet using only the original columns requested by the user.
    Final layout:
    Date, Branch Location, Item, Gross Wt, Stone Wt, Diamond Weight Ct,
    Net Wt, Stone Am, Item Amount, Label Location
    """

    raw_df = df_filtered.copy()

    # Keep original source values
    if branch_col in raw_df.columns:
        raw_df["Branch Location"] = raw_df[branch_col]
    else:
        raw_df["Branch Location"] = raw_df.get("Branch_Code", "")

    if label_col in raw_df.columns:
        raw_df["Label Location"] = raw_df[label_col]
    else:
        raw_df["Label Location"] = raw_df.get("Item", "")

    if "Item" in raw_df.columns:
        raw_df["Item"] = raw_df["Item"]
    elif item_col and item_col in raw_df.columns:
        raw_df["Item"] = raw_df[item_col]
    else:
        raw_df["Item"] = ""

    raw_df["Gross Wt"] = raw_df[gross_wt_col] if gross_wt_col in raw_df.columns else 0
    raw_df["Diamond Weight Ct"] = raw_df[diamond_wt_col] if diamond_wt_col in raw_df.columns else 0
    raw_df["Item Amount"] = raw_df[amount_col] if amount_col in raw_df.columns else 0

    # Stone Wt uses the same value as diamond weight for the raw export set
    raw_df["Stone Wt"] = raw_df["Diamond Weight Ct"]

    # Net Wt remains equal to Gross Wt unless a specific conversion rule is requested later
    raw_df["Net Wt"] = raw_df["Gross Wt"]

    if diamond_amount_col and diamond_amount_col in raw_df.columns:
        raw_df["Stone Am"] = raw_df[diamond_amount_col]
    else:
        raw_df["Stone Am"] = 0

    columns_order = [
        "Date",
        "Branch Location",
        "Item",
        "Gross Wt",
        "Stone Wt",
        "Diamond Weight Ct",
        "Net Wt",
        "Stone Am",
        "Item Amount",
        "Label Location"
    ]

    for col in columns_order:
        if col not in raw_df.columns:
            raw_df[col] = ""

    raw_df = raw_df[columns_order]
    raw_df = raw_df.sort_values(["Date", "Branch Location"])

    return raw_df
def build_single_sheet_report(
    df_filtered,
    amount_col,
    gross_wt_col,
    diamond_wt_col,
    diamond_amount_col,
    branch_col,
    label_col,
    item_col=None  # New parameter
):

    items = [
        "Gold",
        "Silver",
        "Diamond",
        "Platinum"
    ]

    unique_branches = df_filtered["Branch_Code"].unique()

    branches = sort_branches(unique_branches)

    # ========================================================
    # TOTAL AMOUNTS FOR PERCENTAGES
    # ========================================================

    total_amounts = {}

    for item in items:

        item_data = df_filtered[
            df_filtered["Item"] == item
        ]

        total_amounts[item] = (
            float(item_data[amount_col].sum())
            if len(item_data)
            else 0.0
        )

    # ========================================================
    # WORKBOOK
    # ========================================================

    wb = Workbook()

    ws = wb.active

    ws.title = "Sales Report"

    # ========================================================
    # STYLES
    # ========================================================

    title_font = Font(
        bold=True,
        size=16,
        color="FFFFFF"
    )

    title_fill = PatternFill(
        start_color="1F4E79",
        end_color="1F4E79",
        fill_type="solid"
    )

    section_font = Font(
        bold=True,
        size=14,
        color="FFFFFF"
    )

    section_fill = PatternFill(
        start_color="2E75B6",
        end_color="2E75B6",
        fill_type="solid"
    )

    branch_font = Font(
        bold=True,
        size=11,
        color="FFFFFF"
    )

    branch_fill = PatternFill(
        start_color="4472C4",
        end_color="4472C4",
        fill_type="solid"
    )

    header_font = Font(
        bold=True,
        size=10,
        color="000000"
    )

    header_fill = PatternFill(
        start_color="D9E1F2",
        end_color="D9E1F2",
        fill_type="solid"
    )

    total_font = Font(
        bold=True,
        size=10,
        color="000000"
    )

    total_fill = PatternFill(
        start_color="FFF2CC",
        end_color="FFF2CC",
        fill_type="solid"
    )

    grand_total_font = Font(
        bold=True,
        size=11,
        color="FFFFFF"
    )

    grand_total_fill = PatternFill(
        start_color="C00000",
        end_color="C00000",
        fill_type="solid"
    )

    center = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    right = Alignment(
        horizontal="right",
        vertical="center"
    )

    left_align = Alignment(
        horizontal="left",
        vertical="center"
    )

    thin = Side(
        style="thin",
        color="B7B7B7"
    )

    medium = Side(
        style="medium",
        color="000000"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    thick_border = Border(
        left=medium,
        right=medium,
        top=medium,
        bottom=medium
    )

    # ========================================================
    # COLUMNS PER BRANCH
    #
    # Branch Name  = 1 (NEW)
    # Gold         = 3
    # Silver       = 3
    # Diamond      = 5 (Gross Wt, Wt Ct, Amt, Item Amt, %)
    # Platinum     = 3
    #
    # TOTAL = 15 per branch
    # ========================================================

    cols_per_branch = 15

    total_cols = (
        2  # Date and Day Name columns
        + len(branches) * cols_per_branch
    )

    # ========================================================
    # SUB HEADERS
    # ========================================================

    def get_sub_headers():

        return [
            "Branch Name",  # NEW
            "Gold Wt (g)",
            "Gold Item Amt",
            "Gold %",
            "Silver Wt (g)",
            "Silver Item Amt",
            "Silver %",
            "Diamond Gross Wt (g)",
            "Diamond Wt (Ct)",
            "Diamond Amt",
            "Diamond Item Amt",
            "Diamond %",
            "Platinum Wt (g)",
            "Platinum Item Amt",
            "Platinum %",
        ]

    # ========================================================
    # CALCULATE BRANCH TOTALS
    # ========================================================

    def branch_item_totals(data):

        result = {}

        branch_weights = {}

        # ----------------------------------------------------
        # Calculate each branch/category
        # ----------------------------------------------------

        for b in branches:

            bdf = data[
                data["Branch_Code"] == b
            ]

            for item in items:

                idf = bdf[
                    bdf["Item"] == item
                ]

                item_amt = (
                    float(idf[amount_col].sum())
                    if len(idf)
                    else 0.0
                )

                # ====================================================
                # GOLD
                # ====================================================

                if item == "Gold":

                    gold_df = bdf[
                        bdf["Item"] == "Gold"
                    ]

                    gold_weight = (
                        float(
                            gold_df[gross_wt_col].sum()
                        )
                        if len(gold_df)
                        else 0.0
                    )

                    result[(b, item)] = {
                        "wt": gold_weight,
                        "amt": item_amt,
                        "pct": 0.0,
                    }

                    branch_weights[(b, item)] = gold_weight

                # ====================================================
                # DIAMOND
                # ====================================================

                elif item == "Diamond":

                    # Diamond Gross Wt (g) - ONLY from diamond rows
                    diamond_gross_wt = (
                        float(idf[gross_wt_col].sum())
                        if len(idf)
                        else 0.0
                    )

                    # Diamond Weight (Ct) - from ALL rows with Diamond Wt (Ct) > 0
                    diamond_weight = (
                        float(
                            bdf[
                                bdf[diamond_wt_col] > 0
                            ][diamond_wt_col].sum()
                        )
                        if len(bdf)
                        else 0.0
                    )

                    if (
                        diamond_amount_col
                        and diamond_amount_col in idf.columns
                    ):

                        diamond_amount = (
                            float(
                                idf[
                                    diamond_amount_col
                                ].sum()
                            )
                            if len(idf)
                            else 0.0
                        )

                    else:

                        diamond_amount = 0.0

                    result[(b, item)] = {

                        "gross_wt": diamond_gross_wt,

                        "stone_wt": diamond_weight,

                        "diamond_amt": diamond_amount,

                        "item_amt": item_amt,

                        "amt": item_amt,

                        "pct": 0.0,
                    }

                    branch_weights[(b, item)] = diamond_weight

                # ====================================================
                # SILVER / PLATINUM
                # ====================================================

                else:

                    weight = (
                        float(
                            idf[gross_wt_col].sum()
                        )
                        if len(idf)
                        else 0.0
                    )

                    result[(b, item)] = {

                        "wt": weight,

                        "amt": item_amt,

                        "pct": 0.0,
                    }

                    branch_weights[(b, item)] = weight

        # ========================================================
        # ITEM TOTALS (WEIGHT WISE)
        # ========================================================

        item_totals = {}

        for b in branches:

            for item in items:

                wt = branch_weights.get(
                    (b, item),
                    0.0
                )

                item_totals[item] = (
                    item_totals.get(item, 0.0)
                    + wt
                )

        # ========================================================
        # PERCENTAGES (WEIGHT WISE)
        # ========================================================

        for b in branches:

            for item in items:

                wt = branch_weights.get(
                    (b, item),
                    0.0
                )

                denominator = item_totals.get(
                    item,
                    0.0
                )

                pct = (
                    wt / denominator * 100
                    if denominator > 0
                    else 0.0
                )

                result[
                    (b, item)
                ]["pct"] = pct

        return result

    # ========================================================
    # WRITE SUB HEADER
    # ========================================================

    def write_sub_header_row(
        row,
        first_col_label="",
        second_col_label="",
        first_branch_col=3
    ):

        c0 = ws.cell(
            row=row,
            column=1,
            value=first_col_label
        )

        c0.font = header_font
        c0.fill = header_fill
        c0.alignment = center
        c0.border = border

        if first_branch_col == 3:

            c1 = ws.cell(
                row=row,
                column=2,
                value=second_col_label
            )

            c1.font = header_font
            c1.fill = header_fill
            c1.alignment = center
            c1.border = border

        headers = get_sub_headers()

        col = first_branch_col

        for _ in branches:

            # Branch Name column (NEW)
            cell = ws.cell(
                row=row,
                column=col,
                value="Branch"
            )
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.alignment = center
            cell.border = border
            cell.fill = PatternFill(
                start_color="4472C4",
                end_color="4472C4",
                fill_type="solid"
            )
            col += 1

            for header in headers[1:]:  # Skip "Branch Name" as we already added it

                cell = ws.cell(
                    row=row,
                    column=col,
                    value=header
                )

                cell.font = header_font
                cell.alignment = center
                cell.border = border

                # Weight columns
                if (
                    "Wt" in header
                    or "Weight" in header
                ):

                    cell.fill = PatternFill(
                        start_color="E2EFDA",
                        end_color="E2EFDA",
                        fill_type="solid"
                    )

                # Amount columns
                elif (
                    "Amt" in header
                    or "Amount" in header
                ):

                    cell.fill = PatternFill(
                        start_color="FFEB9C",
                        end_color="FFEB9C",
                        fill_type="solid"
                    )

                # Percentage
                elif "%" in header:

                    cell.fill = PatternFill(
                        start_color="E6E6FA",
                        end_color="E6E6FA",
                        fill_type="solid"
                    )

                else:

                    cell.fill = header_fill

                col += 1

    # ========================================================
    # WRITE TOTAL ROW
    # ========================================================

    def write_totals_row(
        row,
        totals,
        row_label,
        day_name="",
        bold=False,
        is_grand=False,
        first_branch_col=3
    ):

        label_cell = ws.cell(
            row=row,
            column=1,
            value=row_label
        )

        label_cell.alignment = center
        label_cell.border = (
            thick_border
            if is_grand
            else border
        )

        if is_grand:

            label_cell.font = grand_total_font
            label_cell.fill = grand_total_fill

        elif bold:

            label_cell.font = total_font
            label_cell.fill = total_fill

        if first_branch_col == 3:

            day_cell = ws.cell(
                row=row,
                column=2,
                value=day_name
            )

            day_cell.alignment = center
            day_cell.border = thick_border if is_grand else border

            if is_grand:
                day_cell.font = grand_total_font
                day_cell.fill = grand_total_fill
            elif bold:
                day_cell.font = total_font
                day_cell.fill = total_fill

        col = first_branch_col

        for b in branches:

            # Write Branch Name (NEW)
            branch_cell = ws.cell(
                row=row,
                column=col,
                value=get_branch_display_name(b)
            )
            branch_cell.alignment = center
            branch_cell.border = thick_border if is_grand else border
            branch_cell.fill = PatternFill(
                start_color="E2EFDA" if not is_grand else "C00000",
                end_color="E2EFDA" if not is_grand else "C00000",
                fill_type="solid"
            )
            branch_cell.font = Font(bold=True, size=9, color="000000" if not is_grand else "FFFFFF")
            col += 1

            # ====================================================
            # GOLD
            # ====================================================

            gold_data = totals.get(
                (b, "Gold"),
                {}
            )

            gold_wt = gold_data.get(
                "wt",
                0.0
            )

            gold_amt = gold_data.get(
                "amt",
                0.0
            )

            gold_pct = gold_data.get(
                "pct",
                0.0
            )

            values = [
                gold_wt,
                gold_amt,
                100.0 if is_grand else gold_pct,
            ]

            for idx, value in enumerate(values):

                cell = ws.cell(
                    row=row,
                    column=col,
                    value=round(value, 2)
                )

                cell.alignment = (
                    center
                    if idx == 2
                    else right
                )

                cell.border = (
                    thick_border
                    if is_grand
                    else border
                )

                if idx == 2:

                    cell.number_format = (
                        '0.0"%"'
                    )

                elif idx == 1:

                    cell.number_format = (
                        '₹#,##0.00'
                    )

                else:

                    cell.number_format = (
                        '#,##0.00'
                    )

                if is_grand:

                    cell.font = grand_total_font
                    cell.fill = grand_total_fill

                elif bold:

                    cell.font = total_font
                    cell.fill = total_fill

                else:

                    if idx == 1:

                        cell.fill = PatternFill(
                            start_color="FFEB9C",
                            end_color="FFEB9C",
                            fill_type="solid"
                        )

                    elif idx == 2:

                        cell.fill = PatternFill(
                            start_color="E6E6FA",
                            end_color="E6E6FA",
                            fill_type="solid"
                        )

                    else:

                        cell.fill = PatternFill(
                            start_color="E2EFDA",
                            end_color="E2EFDA",
                            fill_type="solid"
                        )

                col += 1

            # ====================================================
            # SILVER
            # ====================================================

            silver_data = totals.get(
                (b, "Silver"),
                {}
            )

            silver_wt = silver_data.get(
                "wt",
                0.0
            )

            silver_amt = silver_data.get(
                "amt",
                0.0
            )

            silver_pct = silver_data.get(
                "pct",
                0.0
            )

            values = [
                silver_wt,
                silver_amt,
                100.0 if is_grand else silver_pct,
            ]

            for idx, value in enumerate(values):

                cell = ws.cell(
                    row=row,
                    column=col,
                    value=round(value, 2)
                )

                cell.alignment = (
                    center
                    if idx == 2
                    else right
                )

                cell.border = (
                    thick_border
                    if is_grand
                    else border
                )

                if idx == 2:

                    cell.number_format = (
                        '0.0"%"'
                    )

                elif idx == 1:

                    cell.number_format = (
                        '₹#,##0.00'
                    )

                else:

                    cell.number_format = (
                        '#,##0.00'
                    )

                if is_grand:

                    cell.font = grand_total_font
                    cell.fill = grand_total_fill

                elif bold:

                    cell.font = total_font
                    cell.fill = total_fill

                else:

                    if idx == 1:

                        cell.fill = PatternFill(
                            start_color="FFEB9C",
                            end_color="FFEB9C",
                            fill_type="solid"
                        )

                    elif idx == 2:

                        cell.fill = PatternFill(
                            start_color="E6E6FA",
                            end_color="E6E6FA",
                            fill_type="solid"
                        )

                    else:

                        cell.fill = PatternFill(
                            start_color="E2EFDA",
                            end_color="E2EFDA",
                            fill_type="solid"
                        )

                col += 1

            # ====================================================
            # DIAMOND
            # ====================================================

            diamond_data = totals.get(
                (b, "Diamond"),
                {}
            )

            diamond_gross_wt = diamond_data.get(
                "gross_wt",
                0.0
            )

            diamond_wt = diamond_data.get(
                "stone_wt",
                0.0
            )

            diamond_amt = diamond_data.get(
                "diamond_amt",
                0.0
            )

            diamond_item_amt = diamond_data.get(
                "item_amt",
                0.0
            )

            diamond_pct = diamond_data.get(
                "pct",
                0.0
            )

            values = [
                diamond_gross_wt,
                diamond_wt,
                diamond_amt,
                diamond_item_amt,
                100.0 if is_grand else diamond_pct,
            ]

            for idx, value in enumerate(values):

                cell = ws.cell(
                    row=row,
                    column=col,
                    value=round(value, 2)
                )

                cell.alignment = (
                    center
                    if idx == 4
                    else right
                )

                cell.border = (
                    thick_border
                    if is_grand
                    else border
                )

                if idx == 4:

                    cell.number_format = (
                        '0.0"%"'
                    )

                elif idx in [2, 3]:

                    cell.number_format = (
                        '₹#,##0.00'
                    )

                else:

                    cell.number_format = (
                        '#,##0.00'
                    )

                if is_grand:

                    cell.font = grand_total_font
                    cell.fill = grand_total_fill

                elif bold:

                    cell.font = total_font
                    cell.fill = total_fill

                else:

                    if idx in [2, 3]:

                        cell.fill = PatternFill(
                            start_color="FFEB9C",
                            end_color="FFEB9C",
                            fill_type="solid"
                        )

                    elif idx == 4:

                        cell.fill = PatternFill(
                            start_color="E6E6FA",
                            end_color="E6E6FA",
                            fill_type="solid"
                        )

                    else:

                        cell.fill = PatternFill(
                            start_color="E2EFDA",
                            end_color="E2EFDA",
                            fill_type="solid"
                        )

                col += 1

            # ====================================================
            # PLATINUM
            # ====================================================

            platinum_data = totals.get(
                (b, "Platinum"),
                {}
            )

            platinum_wt = platinum_data.get(
                "wt",
                0.0
            )

            platinum_amt = platinum_data.get(
                "amt",
                0.0
            )

            platinum_pct = platinum_data.get(
                "pct",
                0.0
            )

            values = [
                platinum_wt,
                platinum_amt,
                100.0 if is_grand else platinum_pct,
            ]

            for idx, value in enumerate(values):

                cell = ws.cell(
                    row=row,
                    column=col,
                    value=round(value, 2)
                )

                cell.alignment = (
                    center
                    if idx == 2
                    else right
                )

                cell.border = (
                    thick_border
                    if is_grand
                    else border
                )

                if idx == 2:

                    cell.number_format = (
                        '0.0"%"'
                    )

                elif idx == 1:

                    cell.number_format = (
                        '₹#,##0.00'
                    )

                else:

                    cell.number_format = (
                        '#,##0.00'
                    )

                if is_grand:

                    cell.font = grand_total_font
                    cell.fill = grand_total_fill

                elif bold:

                    cell.font = total_font
                    cell.fill = total_fill

                else:

                    if idx == 1:

                        cell.fill = PatternFill(
                            start_color="FFEB9C",
                            end_color="FFEB9C",
                            fill_type="solid"
                        )

                    elif idx == 2:

                        cell.fill = PatternFill(
                            start_color="E6E6FA",
                            end_color="E6E6FA",
                            fill_type="solid"
                        )

                    else:

                        cell.fill = PatternFill(
                            start_color="E2EFDA",
                            end_color="E2EFDA",
                            fill_type="solid"
                        )

                col += 1

    # ========================================================
    # TITLE
    # ========================================================

    row = 1

    ws.merge_cells(
        start_row=row,
        start_column=1,
        end_row=row,
        end_column=total_cols
    )

    c = ws.cell(
        row=row,
        column=1,
        value=(
            "📊 BRANCH WISE SALES REPORT - "
            "CUMULATIVE & DAILY"
        )
    )

    c.font = title_font
    c.fill = title_fill
    c.alignment = center

    row += 2

    # ========================================================
    # CUMULATIVE HEADER - REMOVED MERGED BRANCH HEADER ROW
    # ========================================================

    write_sub_header_row(
        row,
        first_col_label="",
        second_col_label="",
        first_branch_col=3
    )

    row += 1

    # ========================================================
    # CUMULATIVE TOTAL
    # ========================================================

    cumulative_totals = branch_item_totals(
        df_filtered
    )

    write_totals_row(
        row,
        cumulative_totals,
        row_label="📌 TOTAL",
        bold=True,
        first_branch_col=3
    )

    row += 2

    # ========================================================
    # DAY WISE SECTION
    # ========================================================

    ws.merge_cells(
        start_row=row,
        start_column=1,
        end_row=row,
        end_column=total_cols
    )

    c = ws.cell(
        row=row,
        column=1,
        value="📅 DAY WISE SALES REPORT"
    )

    c.font = section_font
    c.fill = section_fill
    c.alignment = center

    row += 2

    write_sub_header_row(
        row,
        first_col_label="📅 Date",
        second_col_label="📅 Day Name",
        first_branch_col=3
    )

    date_header_row = row

    row += 1

    # ========================================================
    # DATE LIST
    # ========================================================

    dates = sorted(
        df_filtered["Date"].unique()
    )

    for d in dates:

        day_df = df_filtered[
            df_filtered["Date"] == d
        ]

        date_label = (
            d.strftime("%d-%b-%Y")
            if hasattr(d, "strftime")
            else str(d)
        )

        day_name = (
            d.strftime("%A")
            if hasattr(d, "strftime")
            else ""
        )

        day_totals = branch_item_totals(
            day_df
        )

        write_totals_row(
            row,
            day_totals,
            row_label=date_label,
            day_name=day_name,
            first_branch_col=3
        )

        row += 1

    row += 1

    # ========================================================
    # GRAND TOTAL
    # ========================================================

    write_totals_row(
        row,
        cumulative_totals,
        row_label="🔴 GRAND TOTAL",
        bold=True,
        is_grand=True,
        first_branch_col=3
    )

    # ========================================================
    # COLUMN WIDTH
    # ========================================================

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 15

    for c_idx in range(
        2,
        total_cols + 1
    ):

        ws.column_dimensions[
            get_column_letter(c_idx)
        ].width = 15

    # ========================================================
    # FREEZE PANES
    # ========================================================

    ws.freeze_panes = ws.cell(
        row=date_header_row + 1,
        column=3
    )

    # ========================================================
    # LEGEND
    # ========================================================

    legend_row = row + 2

    ws.merge_cells(
        start_row=legend_row,
        start_column=1,
        end_row=legend_row,
        end_column=total_cols
    )

    legend_cell = ws.cell(
        row=legend_row,
        column=1,
        value=(
            "📌 Legend: "
            "🟢 Weight = grams / carats | "
            "🟡 Amount = ₹ | "
            "🟣 Percentage = % (weight wise) | "
            "🔴 Grand Total | "
            "💎 Diamond Gross Wt = Gross Wt of diamond items only | "
            "Diamond Wt (Ct) = ALL diamond weights from ALL categories | "
            "Diamond Amount shown separately | "
            "📋 Raw Data: Product Item = Original product name from Item column, Category = Detected category"
        )
    )

    legend_cell.font = Font(
        size=9,
        italic=True
    )

    legend_cell.alignment = center

    # ========================================================
    # SHEET 2 - GOLD SALES REPORT
    # ========================================================

    ws2 = wb.create_sheet("Gold Only")

    gold_only_df = df_filtered[
        df_filtered["Item"] == "Gold"
    ].copy()

    gold_branches = sort_branches(
        gold_only_df["Branch_Code"].unique()
    )

    gold_cols_per_branch = 4  # Added Branch Name column
    gold_total_cols = (
        2
        + len(gold_branches) * gold_cols_per_branch
    )

    # Define color fills for gold sheet
    gold_weight_fill = PatternFill(
        start_color="E2EFDA",
        end_color="E2EFDA",
        fill_type="solid"
    )

    gold_amount_fill = PatternFill(
        start_color="FFEB9C",
        end_color="FFEB9C",
        fill_type="solid"
    )

    gold_pct_fill = PatternFill(
        start_color="E6E6FA",
        end_color="E6E6FA",
        fill_type="solid"
    )

    def write_gold_headers(row, date_label="", day_label=""):

        # Date column
        cell = ws2.cell(row=row, column=1, value=date_label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

        # Day Name column
        cell = ws2.cell(row=row, column=2, value=day_label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

        col = 3

        for _ in gold_branches:

            # Branch Name column
            cell = ws2.cell(row=row, column=col, value="Branch")
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.alignment = center
            cell.border = border
            cell.fill = PatternFill(
                start_color="4472C4",
                end_color="4472C4",
                fill_type="solid"
            )
            col += 1

            # Gold Wt (g) - Green fill
            cell = ws2.cell(row=row, column=col, value="Gold Wt (g)")
            cell.font = header_font
            cell.alignment = center
            cell.border = border
            cell.fill = gold_weight_fill
            col += 1

            # Gold Item Amt - Yellow fill
            cell = ws2.cell(row=row, column=col, value="Gold Item Amt")
            cell.font = header_font
            cell.alignment = center
            cell.border = border
            cell.fill = gold_amount_fill
            col += 1

            # Gold % - Purple fill
            cell = ws2.cell(row=row, column=col, value="Gold %")
            cell.font = header_font
            cell.alignment = center
            cell.border = border
            cell.fill = gold_pct_fill
            col += 1

    def gold_totals(data):

        total_weight = data[gross_wt_col].sum()
        totals = {}

        for branch in gold_branches:

            branch_data = data[
                data["Branch_Code"] == branch
            ]
            weight = float(branch_data[gross_wt_col].sum())
            amount = float(branch_data[amount_col].sum())
            totals[branch] = {
                "wt": weight,
                "amt": amount,
                "pct": weight / total_weight * 100
                if total_weight > 0 else 0.0,
            }

        return totals

    def write_gold_row(row, totals, label, day_name="", grand=False):

        # Label column
        cell = ws2.cell(row=row, column=1, value=label)
        cell.alignment = center
        cell.border = thick_border if grand else border
        if grand:
            cell.font = grand_total_font
            cell.fill = grand_total_fill
        elif label == "📌 TOTAL":
            cell.font = total_font
            cell.fill = total_fill

        # Day Name column
        cell = ws2.cell(row=row, column=2, value=day_name)
        cell.alignment = center
        cell.border = thick_border if grand else border
        if grand:
            cell.font = grand_total_font
            cell.fill = grand_total_fill
        elif label == "📌 TOTAL":
            cell.font = total_font
            cell.fill = total_fill

        col = 3

        for branch in gold_branches:

            data = totals.get(branch, {"wt": 0.0, "amt": 0.0, "pct": 0.0})

            # Branch Name
            cell = ws2.cell(row=row, column=col, value=get_branch_display_name(branch))
            cell.alignment = center
            cell.border = thick_border if grand else border
            cell.fill = PatternFill(
                start_color="E2EFDA" if not grand else "C00000",
                end_color="E2EFDA" if not grand else "C00000",
                fill_type="solid"
            )
            cell.font = Font(bold=True, size=9, color="000000" if not grand else "FFFFFF")
            col += 1

            # Gold Wt (g) - Green fill
            cell = ws2.cell(row=row, column=col, value=round(data["wt"], 2))
            cell.alignment = right
            cell.border = thick_border if grand else border
            cell.number_format = '#,##0.00'
            if grand:
                cell.font = grand_total_font
                cell.fill = grand_total_fill
            elif label == "📌 TOTAL":
                cell.font = total_font
                cell.fill = total_fill
            else:
                cell.fill = gold_weight_fill
            col += 1

            # Gold Item Amt - Yellow fill
            cell = ws2.cell(row=row, column=col, value=round(data["amt"], 2))
            cell.alignment = right
            cell.border = thick_border if grand else border
            cell.number_format = '₹#,##0.00'
            if grand:
                cell.font = grand_total_font
                cell.fill = grand_total_fill
            elif label == "📌 TOTAL":
                cell.font = total_font
                cell.fill = total_fill
            else:
                cell.fill = gold_amount_fill
            col += 1

            # Gold % - Purple fill
            pct_value = 100.0 if grand else data["pct"]
            cell = ws2.cell(row=row, column=col, value=round(pct_value, 1))
            cell.alignment = center
            cell.border = thick_border if grand else border
            cell.number_format = '0.0"%"'
            if grand:
                cell.font = grand_total_font
                cell.fill = grand_total_fill
            elif label == "📌 TOTAL":
                cell.font = total_font
                cell.fill = total_fill
            else:
                cell.fill = gold_pct_fill
            col += 1

    row2 = 1
    ws2.merge_cells(
        start_row=row2,
        start_column=1,
        end_row=row2,
        end_column=gold_total_cols
    )
    title = ws2.cell(
        row=row2,
        column=1,
        value="🥇 GOLD SALES REPORT - CUMULATIVE & DAILY"
    )
    title.font = title_font
    title.fill = title_fill
    title.alignment = center

    row2 += 2
    write_gold_headers(row2)
    row2 += 1

    cumulative_gold_totals = gold_totals(gold_only_df)
    write_gold_row(row2, cumulative_gold_totals, "📌 TOTAL")
    row2 += 2

    ws2.merge_cells(
        start_row=row2,
        start_column=1,
        end_row=row2,
        end_column=gold_total_cols
    )
    section = ws2.cell(
        row=row2,
        column=1,
        value="📅 GOLD DAY WISE SALES REPORT"
    )
    section.font = section_font
    section.fill = section_fill
    section.alignment = center

    row2 += 2
    write_gold_headers(row2, "📅 Date", "📅 Day Name")
    gold_date_header_row = row2
    row2 += 1

    for date_value in sorted(gold_only_df["Date"].unique()):

        day_data = gold_only_df[
            gold_only_df["Date"] == date_value
        ]
        write_gold_row(
            row2,
            gold_totals(day_data),
            date_value.strftime("%d-%b-%Y"),
            date_value.strftime("%A")
        )
        row2 += 1

    row2 += 1
    write_gold_row(
        row2,
        cumulative_gold_totals,
        "🔴 GRAND TOTAL",
        grand=True
    )

    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 15
    for column in range(3, gold_total_cols + 1):
        ws2.column_dimensions[get_column_letter(column)].width = 15

    ws2.freeze_panes = ws2.cell(
        row=gold_date_header_row + 1,
        column=3
    )

    # ========================================================
    # SHEET 3 - RAW DATA
    # ========================================================

    ws3 = wb.create_sheet("Raw Data")

    raw_df = build_raw_data_sheet(
        df_filtered,
        branch_col,
        label_col,
        amount_col,
        gross_wt_col,
        diamond_wt_col,
        diamond_amount_col,
        item_col  # Pass the item column
    )

    # Write headers
    header_row = 1
    for col_idx, col_name in enumerate(raw_df.columns, 1):
        cell = ws3.cell(row=header_row, column=col_idx, value=col_name)
        cell.font = Font(bold=True, size=10, color="FFFFFF")
        cell.fill = PatternFill(
            start_color="4472C4",
            end_color="4472C4",
            fill_type="solid"
        )
        cell.alignment = center
        cell.border = border

    # Write data
    for row_idx, row_data in enumerate(raw_df.values, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws3.cell(row=row_idx, column=col_idx, value=value)

            # Apply number formatting for numeric columns
            if col_idx in [4, 5, 6, 7, 8, 9]:
                if isinstance(value, (int, float)):
                    if col_idx in [8, 9]:
                        cell.number_format = '₹#,##0.00'
                    else:
                        cell.number_format = '#,##0.000'
                    cell.alignment = right
            elif col_idx == 1:
                cell.alignment = center
            elif col_idx == 3:
                cell.alignment = left_align
            elif col_idx == 10:
                cell.alignment = left_align

            cell.border = border

    # Set column widths
    column_widths = {
        1: 18,   # Date
        2: 25,   # Branch Location
        3: 25,   # Item
        4: 15,   # Gross Wt
        5: 15,   # Stone Wt
        6: 18,   # Diamond Weight Ct
        7: 15,   # Net Wt
        8: 18,   # Stone Am
        9: 18,   # Item Amount
        10: 30,  # Label Location
    }

    for col_idx, width in column_widths.items():
        ws3.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze header row
    ws3.freeze_panes = ws3.cell(row=2, column=1)

    # ========================================================
    # SAVE
    # ========================================================

    output = io.BytesIO()

    wb.save(output)

    return output.getvalue()
def process_data_cached(
    df,
    date_col,
    branch_col,
    label_col,
    amount_col,
    gross_wt_col,
    diamond_wt_col,
    diamond_amount_col,
    item_col,  # New parameter
    fill_blanks
):

    return process_data(
        df,
        date_col,
        branch_col,
        label_col,
        amount_col,
        gross_wt_col,
        diamond_wt_col,
        diamond_amount_col,
        item_col,
        fill_blanks
    )
def process_data(
    df,
    date_col,
    branch_col,
    label_col,
    amount_col,
    gross_wt_col,
    diamond_wt_col,
    diamond_amount_col,
    item_col,  # New parameter
    fill_blanks
):

    # ========================================================
    # VALIDATION
    # ========================================================

    required_cols = [
        date_col,
        branch_col,
        label_col,
        amount_col,
        gross_wt_col,
        diamond_wt_col,
    ]

    missing_cols = [
        col
        for col in required_cols
        if col not in df.columns
    ]

    if missing_cols:

        raise ValueError(
            "Missing required columns: "
            + ", ".join(
                map(str, missing_cols)
            )
        )

    df_clean = df.copy()

    # Track original row count
    original_count = len(df_clean)
    st.info(f"📊 Original rows: {original_count}")

    # ========================================================
    # FILL BLANK BRANCHES
    # ========================================================

    filled_count = 0

    if fill_blanks:

        df_clean, filled_count = (
            fill_blank_branches(
                df_clean,
                branch_col
            )
        )

    # ========================================================
    # DATE
    # ========================================================

    df_clean["Date"] = pd.to_datetime(
        df_clean[date_col],
        errors="coerce"
    ).dt.date

    # Count invalid dates
    invalid_dates = df_clean["Date"].isna().sum()
    if invalid_dates > 0:
        st.warning(f"⚠️ {invalid_dates} rows had invalid dates and were removed")

    # Remove invalid dates
    df_clean = df_clean[
        df_clean["Date"].notna()
    ].copy()

    after_dates = len(df_clean)
    st.info(f"📊 After date validation: {after_dates} rows")

    # ========================================================
    # BRANCH - IMPROVED HANDLING
    # ========================================================

    # First, try to extract branch code from the column
    df_clean["Branch_Code"] = (
        df_clean[branch_col]
        .astype(str)
        .apply(extract_branch_code)
    )

    # If we have a "Branch" column that contains full names, use it directly
    # Check if the branch column contains full names like "Madurai"
    branch_values = df_clean[branch_col].astype(str).str.strip().unique()

    # Create a mapping from branch value to code
    branch_code_map = {}
    for val in branch_values:
        code = extract_branch_code(val)
        branch_code_map[val] = code

        # If the code is not in BRANCH_DISPLAY_NAMES, try to find a match
        if code not in BRANCH_DISPLAY_NAMES:
            # Try to match the original value directly
            for display_name in BRANCH_DISPLAY_NAMES.values():
                if display_name.upper() in val.upper() or val.upper() in display_name.upper():
                    # Find the code for this display name
                    for code_key, code_val in BRANCH_DISPLAY_NAMES.items():
                        if code_val == display_name:
                            branch_code_map[val] = code_key
                            break
                    break

    # Apply the mapping
    df_clean["Branch_Code"] = df_clean[branch_col].astype(str).str.strip().map(branch_code_map)

    # If any remain unmapped, try the original extraction
    mask_unmapped = df_clean["Branch_Code"].isna()
    if mask_unmapped.any():
        df_clean.loc[mask_unmapped, "Branch_Code"] = (
            df_clean.loc[mask_unmapped, branch_col]
            .astype(str)
            .apply(extract_branch_code)
        )

    # ========================================================
    # ITEM
    # ========================================================

    df_clean["Item"] = (
        df_clean[label_col]
        .astype(str)
        .apply(detect_item_from_label)
    )

    # Show item distribution
    item_dist = df_clean["Item"].value_counts()
    st.write("### Item Distribution:")
    st.dataframe(item_dist)

    # ========================================================
    # NUMERIC COLUMNS
    # ========================================================

    df_clean[gross_wt_col] = (
        clean_numeric_series(
            df_clean[gross_wt_col]
        )
    )

    df_clean[diamond_wt_col] = (
        clean_numeric_series(
            df_clean[diamond_wt_col]
        )
    )

    df_clean[amount_col] = (
        clean_numeric_series(
            df_clean[amount_col]
        )
    )

    # ========================================================
    # DIAMOND AMOUNT
    # ========================================================

    if (
        diamond_amount_col
        and diamond_amount_col in df_clean.columns
    ):

        df_clean[diamond_amount_col] = (
            clean_numeric_series(
                df_clean[diamond_amount_col]
            )
        )

    else:

        diamond_amount_col = None

    # ========================================================
    # FILTER VALID ITEMS
    # ========================================================

    # Count before filtering
    before_filter = len(df_clean)

    df_filtered = df_clean[
        ~df_clean["Item"].isin(
            [
                "Skip",
                "Unknown"
            ]
        )
    ].copy()

    after_filter = len(df_filtered)

    # Show filtering stats
    removed_count = before_filter - after_filter
    st.info(f"📊 After filtering (removed Skip/Unknown): {after_filter} rows (removed {removed_count} rows)")

    # Show what was removed if any
    if removed_count > 0:
        removed = df_clean[df_clean["Item"].isin(["Skip", "Unknown"])]
        with st.expander(f"⚠️ {removed_count} rows were removed (Skip/Unknown) - Click to view"):
            st.write(removed[[branch_col, label_col, "Item"]].head(20))

    if len(df_filtered) == 0:

        raise ValueError(
            "No valid data found after filtering. "
            "Please check your Label Location mapping."
        )

    # ========================================================
    # TOTAL AMOUNTS
    # ========================================================

    total_amounts = {}

    for item in [
        "Gold",
        "Silver",
        "Diamond",
        "Platinum"
    ]:

        item_data = df_filtered[
            df_filtered["Item"] == item
        ]

        total_amounts[item] = (
            float(
                item_data[amount_col].sum()
            )
            if len(item_data)
            else 0.0
        )

    # ========================================================
    # BRANCH LIST - SHOW ORIGINAL NAMES
    # ========================================================

    unique_branches = sort_branches(
        df_filtered[
            "Branch_Code"
        ].unique()
    )

    # Also get the original branch names for display
    branch_name_map = {}
    for code in unique_branches:
        # Find the first occurrence of this branch code
        mask = df_filtered["Branch_Code"] == code
        if mask.any():
            # Get the original branch name
            original_name = df_filtered.loc[mask, branch_col].iloc[0]
            branch_name_map[code] = str(original_name).strip()
        else:
            branch_name_map[code] = code

    # ========================================================
    # REPORT DATAFRAME
    # ========================================================

    report_df = pd.DataFrame(
        index=unique_branches
    )

    # ========================================================
    # GOLD
    # ========================================================

    gold_data = df_filtered[
        df_filtered["Item"] == "Gold"
    ]

    gold_weight = (
        gold_data.groupby(
            "Branch_Code"
        )[gross_wt_col].sum()
        if len(gold_data)
        else pd.Series(dtype=float)
    )

    gold_weight_series = pd.Series(
        0.0,
        index=unique_branches
    )

    gold_weight_series.update(
        gold_weight
    )

    report_df["Gold Wt (g)"] = (
        gold_weight_series
    )

    # Gold Item Amount
    gold_amount = (
        gold_data.groupby(
            "Branch_Code"
        )[amount_col].sum()
        if len(gold_data)
        else pd.Series(dtype=float)
    )

    gold_amount_series = pd.Series(
        0.0,
        index=unique_branches
    )

    gold_amount_series.update(
        gold_amount
    )

    report_df["Gold Item Amt"] = (
        gold_amount_series
    )

    # Gold % — WEIGHT WISE (based on actual Gold weight)
    total_gold_weight = report_df["Gold Wt (g)"].sum()

    if total_gold_weight > 0:

        report_df["Gold %"] = (
            report_df["Gold Wt (g)"]
            / total_gold_weight
            * 100
        ).round(1)

    else:

        report_df["Gold %"] = 0.0

    # ========================================================
    # SILVER
    # ========================================================

    silver_data = df_filtered[
        df_filtered["Item"] == "Silver"
    ]

    if len(silver_data) > 0:

        silver_weight = (
            silver_data.groupby(
                "Branch_Code"
            )[gross_wt_col].sum()
        )

        silver_amount = (
            silver_data.groupby(
                "Branch_Code"
            )[amount_col].sum()
        )

    else:

        silver_weight = pd.Series(
            dtype=float
        )

        silver_amount = pd.Series(
            dtype=float
        )

    silver_weight_series = pd.Series(
        0.0,
        index=unique_branches
    )

    silver_amount_series = pd.Series(
        0.0,
        index=unique_branches
    )

    silver_weight_series.update(
        silver_weight
    )

    silver_amount_series.update(
        silver_amount
    )

    report_df["Silver Wt (g)"] = (
        silver_weight_series
    )

    report_df["Silver Item Amt"] = (
        silver_amount_series
    )

    # Silver % — WEIGHT WISE
    total_silver_weight = report_df["Silver Wt (g)"].sum()

    if total_silver_weight > 0:

        report_df["Silver %"] = (
            report_df["Silver Wt (g)"]
            / total_silver_weight
            * 100
        ).round(1)

    else:

        report_df["Silver %"] = 0.0

    # ========================================================
    # DIAMOND
    # ========================================================

    # Diamond Gross Wt (g) - ONLY from diamond items
    diamond_data = df_filtered[
        df_filtered["Item"] == "Diamond"
    ]

    if len(diamond_data) > 0:

        diamond_gross_wt = (
            diamond_data.groupby(
                "Branch_Code"
            )[gross_wt_col].sum()
        )

        # Diamond Item Amount
        diamond_item_amount = (
            diamond_data.groupby(
                "Branch_Code"
            )[amount_col].sum()
        )

        # Diamond Amount
        if (
            diamond_amount_col
            and diamond_amount_col in diamond_data.columns
        ):

            diamond_amount = (
                diamond_data.groupby(
                    "Branch_Code"
                )[diamond_amount_col].sum()
            )

        else:

            diamond_amount = pd.Series(
                0.0,
                index=unique_branches
            )

    else:

        diamond_gross_wt = pd.Series(
            dtype=float
        )

        diamond_item_amount = pd.Series(
            dtype=float
        )

        diamond_amount = pd.Series(
            dtype=float
        )

    diamond_gross_wt_series = pd.Series(
        0.0,
        index=unique_branches
    )

    diamond_item_amount_series = pd.Series(
        0.0,
        index=unique_branches
    )

    diamond_amount_series = pd.Series(
        0.0,
        index=unique_branches
    )

    diamond_gross_wt_series.update(
        diamond_gross_wt
    )

    diamond_item_amount_series.update(
        diamond_item_amount
    )

    diamond_amount_series.update(
        diamond_amount
    )

    report_df["Diamond Gross Wt (g)"] = (
        diamond_gross_wt_series
    )

    # Diamond Wt (Ct) - from ALL rows with Diamond Wt (Ct) > 0
    diamond_weight_all = (
        df_filtered[
            df_filtered[diamond_wt_col] > 0
        ].groupby("Branch_Code")[diamond_wt_col].sum()
    )

    diamond_weight_series = pd.Series(
        0.0,
        index=unique_branches
    )

    diamond_weight_series.update(
        diamond_weight_all
    )

    report_df["Diamond Wt (Ct)"] = (
        diamond_weight_series
    )

    report_df["Diamond Amt"] = (
        diamond_amount_series
    )

    report_df["Diamond Item Amt"] = (
        diamond_item_amount_series
    )

    # Diamond % — WEIGHT WISE (based on ALL Diamond Wt Ct)
    total_diamond_weight = report_df["Diamond Wt (Ct)"].sum()

    if total_diamond_weight > 0:

        report_df["Diamond %"] = (
            report_df["Diamond Wt (Ct)"]
            / total_diamond_weight
            * 100
        ).round(1)

    else:

        report_df["Diamond %"] = 0.0

    # ========================================================
    # PLATINUM
    # ========================================================

    platinum_data = df_filtered[
        df_filtered["Item"] == "Platinum"
    ]

    if len(platinum_data) > 0:

        platinum_weight = (
            platinum_data.groupby(
                "Branch_Code"
            )[gross_wt_col].sum()
        )

        platinum_amount = (
            platinum_data.groupby(
                "Branch_Code"
            )[amount_col].sum()
        )

    else:

        platinum_weight = pd.Series(
            dtype=float
        )

        platinum_amount = pd.Series(
            dtype=float
        )

    platinum_weight_series = pd.Series(
        0.0,
        index=unique_branches
    )

    platinum_amount_series = pd.Series(
        0.0,
        index=unique_branches
    )

    platinum_weight_series.update(
        platinum_weight
    )

    platinum_amount_series.update(
        platinum_amount
    )

    report_df["Platinum Wt (g)"] = (
        platinum_weight_series
    )

    report_df["Platinum Item Amt"] = (
        platinum_amount_series
    )

    # Platinum % — WEIGHT WISE
    total_platinum_weight = report_df["Platinum Wt (g)"].sum()

    if total_platinum_weight > 0:

        report_df["Platinum %"] = (
            report_df["Platinum Wt (g)"]
            / total_platinum_weight
            * 100
        ).round(1)

    else:

        report_df["Platinum %"] = 0.0

    # ========================================================
    # COLUMN ORDER
    # ========================================================

    column_order = [

        "Gold Wt (g)",
        "Gold Item Amt",
        "Gold %",

        "Silver Wt (g)",
        "Silver Item Amt",
        "Silver %",

        "Diamond Gross Wt (g)",
        "Diamond Wt (Ct)",
        "Diamond Amt",
        "Diamond Item Amt",
        "Diamond %",

        "Platinum Wt (g)",
        "Platinum Item Amt",
        "Platinum %",
    ]

    report_df = report_df[
        column_order
    ]

    # ========================================================
    # DISPLAY DATAFRAME - USE ORIGINAL BRANCH NAMES
    # ========================================================

    report_display = report_df.copy()

    # Map branch codes to display names
    display_names = []
    for code in report_display.index:
        # First try to get the original branch name from the data
        if code in branch_name_map:
            original_name = branch_name_map[code]
            # Check if original name is in the display name mapping
            if original_name in BRANCH_DISPLAY_NAMES.values():
                display_names.append(original_name)
            else:
                # Try to find a match
                display_name = get_branch_display_name(code)
                display_names.append(display_name)
        else:
            display_names.append(get_branch_display_name(code))

    report_display.index = display_names

    report_display.index.name = "Branch"

    report_display = (
        report_display.reset_index()
    )

    # ========================================================
    # GRAND TOTAL
    # ========================================================

    total_row = {
        "Branch": "GRAND TOTAL"
    }

    for col in report_display.columns:

        if col == "Branch":
            continue

        if "%" in col:

            total_row[col] = 100.0

        else:

            total_row[col] = (
                report_display[col].sum()
            )

    report_display = pd.concat(
        [
            report_display,
            pd.DataFrame([total_row])
        ],
        ignore_index=True
    )

    # ========================================================
    # DETECTION STATISTICS
    # ========================================================

    detection_stats = (
        df_clean["Item"].value_counts()
    )

    stats_data = []

    for cat in [
        "Gold",
        "Silver",
        "Diamond",
        "Platinum",
        "Skip",
        "Unknown"
    ]:

        count = detection_stats.get(
            cat,
            0
        )

        pct = (
            count
            / len(df_clean)
            * 100
            if len(df_clean) > 0
            else 0
        )

        stats_data.append(
            {
                "Category": cat,
                "Count": count,
                "Percentage": f"{pct:.1f}%"
            }
        )

    stats_df = pd.DataFrame(
        stats_data
    )

    # ========================================================
    # BRANCH MAPPING
    # ========================================================

    branch_mapping = (
        df_clean.groupby(
            branch_col
        )["Branch_Code"]
        .first()
        .reset_index()
    )

    # ========================================================
    # UNKNOWN DATA
    # ========================================================

    unknown_data = df_clean[
        df_clean["Item"] == "Unknown"
    ][
        [
            branch_col,
            label_col,
            "Item"
        ]
    ].head(20)

    # ========================================================
    # DAY-WISE AMOUNT
    # ========================================================

    day_wise = df_filtered.pivot_table(
        index="Date",
        columns="Item",
        values=amount_col,
        aggfunc="sum",
        fill_value=0
    )

    # ========================================================
    # DAY-WISE WEIGHT
    # ========================================================

    day_wise_weight = pd.DataFrame(
        index=sorted(
            df_filtered["Date"].unique()
        )
    )

    for d in day_wise_weight.index:

        day_data = df_filtered[
            df_filtered["Date"] == d
        ]

        # Gold weight = Actual Gold Gross Wt ONLY
        gold_weight_day = (
            day_data[
                day_data["Item"] == "Gold"
            ][gross_wt_col].sum()
        )

        # Diamond Gross Weight - from diamond items only
        diamond_gross_day = (
            day_data[
                day_data["Item"] == "Diamond"
            ][gross_wt_col].sum()
        )

        # Diamond Wt (Ct) - from ALL rows with Diamond Wt (Ct) > 0
        diamond_weight_day = (
            day_data[
                day_data[diamond_wt_col] > 0
            ][diamond_wt_col].sum()
        )

        silver_weight_day = (
            day_data[
                day_data["Item"] == "Silver"
            ][gross_wt_col].sum()
        )

        platinum_weight_day = (
            day_data[
                day_data["Item"] == "Platinum"
            ][gross_wt_col].sum()
        )

        day_wise_weight.loc[
            d,
            "Gold Wt (g)"
        ] = gold_weight_day

        day_wise_weight.loc[
            d,
            "Silver Wt (g)"
        ] = silver_weight_day

        day_wise_weight.loc[
            d,
            "Diamond Gross Wt (g)"
        ] = diamond_gross_day

        day_wise_weight.loc[
            d,
            "Diamond Wt (Ct)"
        ] = diamond_weight_day

        day_wise_weight.loc[
            d,
            "Platinum Wt (g)"
        ] = platinum_weight_day

    # ========================================================
    # TOTAL BY ITEM
    # ========================================================

    total_by_item = {}

    for item in [
        "Gold",
        "Silver",
        "Diamond",
        "Platinum"
    ]:

        item_data = df_filtered[
            df_filtered["Item"] == item
        ]

        if item == "Gold":

            total_weight = float(item_data[gross_wt_col].sum())

            total_by_item[item] = {
                "Weight": total_weight,
                "Item Amount": float(item_data[amount_col].sum()),
                "Diamond Amount": 0.0,
            }

        elif item == "Silver":

            total_weight = float(item_data[gross_wt_col].sum())

            total_by_item[item] = {
                "Weight": total_weight,
                "Item Amount": float(item_data[amount_col].sum()),
                "Diamond Amount": 0.0,
            }

        elif item == "Diamond":

            diamond_gross_wt = float(item_data[gross_wt_col].sum())

            diamond_weight_ct = float(
                df_filtered[
                    df_filtered[diamond_wt_col] > 0
                ][diamond_wt_col].sum()
            )

            total_diamond_amount = (
                float(item_data[diamond_amount_col].sum())
                if diamond_amount_col and diamond_amount_col in item_data.columns
                else 0.0
            )

            total_by_item[item] = {
                "Gross Weight": diamond_gross_wt,
                "Weight": diamond_weight_ct,
                "Item Amount": float(item_data[amount_col].sum()),
                "Diamond Amount": total_diamond_amount,
            }

        elif item == "Platinum":

            total_weight = float(item_data[gross_wt_col].sum())

            total_by_item[item] = {
                "Weight": total_weight,
                "Item Amount": float(item_data[amount_col].sum()),
                "Diamond Amount": 0.0,
            }

    total_by_item = pd.DataFrame(total_by_item).T

    # ========================================================
    # EXCEL REPORT - Updated to pass item_col
    # ========================================================

    formatted_bytes = build_single_sheet_report(
        df_filtered,
        amount_col,
        gross_wt_col,
        diamond_wt_col,
        diamond_amount_col,
        branch_col,
        label_col,
        item_col  # Pass the item column
    )

    # ========================================================
    # SAMPLE DATA
    # ========================================================

    sample_data = df_clean[
        [
            branch_col,
            label_col,
            "Branch_Code",
            "Item"
        ]
    ].head(10)

    return {

        "filled_count": filled_count,

        "sample_data": sample_data,

        "stats_df": stats_df,

        "branch_mapping": branch_mapping,

        "unknown_data": unknown_data,

        "n_rows_filtered": len(
            df_filtered
        ),

        "n_branches": (
            df_filtered[
                "Branch_Code"
            ].nunique()
        ),

        "report_display": report_display,

        "day_wise": day_wise,

        "day_wise_weight": day_wise_weight,

        "total_by_item": total_by_item,

        "amount_col": amount_col,

        "gross_wt_col": gross_wt_col,

        "diamond_wt_col": diamond_wt_col,

        "diamond_amount_col": diamond_amount_col,

        "formatted_bytes": formatted_bytes,

        "total_amounts": total_amounts,
    }

LABEL = "Sales Daywise Report"


def run():
    st.title("📊 Branch-wise Cumulative Sales Report")
    st.markdown("### Gold | Silver | Diamond | Platinum")
    st.markdown(
        "#### Gold Gross Weight includes Diamond Jewellery Gross Weight | "
        "Diamond Gross Weight, Diamond Weight & Diamond Amount shown separately"
    )
    BRANCH_CODE_TO_NAME = {
        "AN": "Anna Nagar-Madurai",
        "DGL": "Dindigul",
        "MDM": "Marthandam",
        "MDU": "Madurai",
        "NDA": "Noida",
        "RJP": "Rajapalayam",
        "SLM": "Salem",
        "TCY": "Trichy",
        "TVL": "Tirunelveli",
        "VNR": "Virudhunagar",
    }
    _FULL_NAME_TO_CODE = {
        name.upper(): code
        for code, name in BRANCH_CODE_TO_NAME.items()
    }
    _FULL_NAME_TO_CODE.update({
        "MADURAI": "MDU",
        "ANNA NAGAR": "AN",
        "VIRUDHURNGAR": "VNR",  # common typo seen in source data
    })
    BRANCH_DISPLAY_NAMES = dict(BRANCH_CODE_TO_NAME)
    BRANCH_DISPLAY_NAMES.update({
        "MADURAI": "Madurai",
        "ANNA NAGAR": "Anna Nagar-Madurai",
        "ANNA NAGAR-MADURAI": "Anna Nagar-Madurai",
        "DINDIGUL": "Dindigul",
        "MARTHANDAM": "Marthandam",
        "NOIDA": "Noida",
        "RAJAPALAYAM": "Rajapalayam",
        "SALEM": "Salem",
        "TRICHY": "Trichy",
        "TIRUNELVELI": "Tirunelveli",
        "VIRUDHUNAGAR": "Virudhunagar",
    })
    BRANCH_ORDER = [
        "MDU",
        "AN",
        "MDM",
        "SLM",
        "TVL",
        "TCY",
        "RJP",
        "DGL",
        "NDA",
        "VNR",
    ]
    st.sidebar.header("📁 Upload Data")
    uploaded_file = st.sidebar.file_uploader(
        "Choose Excel or CSV file",
        type=[
            "xlsx",
            "xls",
            "csv"
        ],
        help="Upload your sales data file"
    )
    with st.sidebar:

        st.divider()

        st.markdown(
            "### 📋 Expected Columns"
        )

        st.markdown(
            """
            - **Date** - Transaction date
            - **Branch Location** - Branch name
            - **Label Location** - Counter/category (used for detection)
            - **Item** - Product name (e.g., RING, NECKLACE, etc.)
            - **Item Amount** - Item sales amount
            - **Gross Wt** - Gross weight in grams
            - **Diamond Weight Ct** - Diamond weight in carats
            - **Diamond Amount** - Separate diamond amount
            """
        )

        st.divider()

        st.markdown(
            "### 💎 Diamond Weight Rule"
        )

        st.info(
            """
            **Gold Wt (g)** = Actual Gold Gross Weight ONLY

            **Diamond Gross Wt (g)** = Gross Wt of diamond items only

            **Diamond Wt (Ct)** = ALL diamond weights from ALL categories
            (Platinum + Diamond, Gold + Diamond, etc.)

            Diamond Weight (Ct) and Diamond Amount are
            reported separately.
            """
        )

        st.divider()

        st.markdown(
            "### 🔍 Branch Order"
        )

        st.markdown(
            """
            1. Madurai
            2. Anna Nagar-Madurai
            3. Marthandam
            4. Salem
            5. Tirunelveli
            6. Trichy
            7. Rajapalayam
            8. Dindigul
            9. Noida
            10. Virudhunagar
            """
        )

        st.divider()

        st.markdown(
            "### 🔍 Detection Rules"
        )

        st.markdown(
            """
            **Auto-detects from Label Location:**

            🥇 Gold
            - MAIN
            - ANTIQUE
            - DESIGNER
            - BOUTIQUE
            - M-
            - A-
            - D-
            - B-

            🥈 Silver
            - SILVER
            - V-

            💎 Diamond
            - DIAMOND
            - R-

            🔹 Platinum
            - PLATINUM
            - P-

            ⏭️ Skip
            - REPAIR
            """
        )

        st.divider()

        fill_blanks = st.checkbox(
            "✅ Fill blank Branch Locations",
            value=True,
            help=(
                "Fill blank branch cells using "
                "the previous branch name"
            )
        )
    if uploaded_file is not None:

        try:

            # ====================================================
            # READ FILE
            # ====================================================

            if uploaded_file.name.lower().endswith(
                ".csv"
            ):

                df = pd.read_csv(
                    uploaded_file
                )

            else:

                df = pd.read_excel(
                    uploaded_file
                )

            # ====================================================
            # EMPTY FILE
            # ====================================================

            if len(df) == 0:

                st.error(
                    "❌ The uploaded file is empty."
                )

                st.stop()

            # ====================================================
            # RESET SESSION WHEN NEW FILE
            # ====================================================

            if (
                st.session_state.get(
                    "_last_uploaded_name"
                )
                != uploaded_file.name
            ):

                st.session_state[
                    "_last_uploaded_name"
                ] = uploaded_file.name

                st.session_state[
                    "report_generated"
                ] = False

                st.session_state.pop(
                    "results",
                    None
                )

            # ====================================================
            # RAW DATA
            # ====================================================

            st.subheader(
                "📄 Raw Data Preview"
            )

            col1, col2, col3 = st.columns(3)

            with col1:

                st.metric(
                    "Total Rows",
                    len(df)
                )

            with col2:

                st.metric(
                    "Total Columns",
                    len(df.columns)
                )

            with col3:

                if "Branch Location" in df.columns:

                    blank_count = (
                        df[
                            "Branch Location"
                        ].isna().sum()
                        +
                        (
                            df[
                                "Branch Location"
                            ].astype(str)
                            .str.strip()
                            == ""
                        ).sum()
                    )

                    st.metric(
                        "Blank Branches",
                        blank_count
                    )

            st.dataframe(
                df.head(10),
                use_container_width=True
            )

            with st.expander(
                "📋 All Columns in Your File"
            ):

                st.write(
                    df.columns.tolist()
                )

            # ====================================================
            # COLUMN MAPPING
            # ====================================================

            st.subheader(
                "🔧 Map Your Columns"
            )

            st.caption(
                "Columns are auto-detected. "
                "Please verify the selections before generating."
            )

            guess = auto_detect_columns(
                df
            )

            cols_list = list(
                df.columns
            )

            def idx(colname):

                if (
                    colname is not None
                    and colname in cols_list
                ):

                    return cols_list.index(
                        colname
                    )

                return 0

            # ====================================================
            # ROW 1
            # ====================================================

            col1, col2, col3 = st.columns(3)

            with col1:

                date_col = st.selectbox(
                    "📅 Date Column",
                    cols_list,
                    index=idx(
                        guess["date"]
                    )
                )

            with col2:

                branch_col = st.selectbox(
                    "🏢 Branch Location Column",
                    cols_list,
                    index=idx(
                        guess["branch"]
                    )
                )

            with col3:

                label_col = st.selectbox(
                    "🏷️ Label Location Column (for detection)",
                    cols_list,
                    index=idx(
                        guess["label"]
                    ),
                    help="This column is used to detect the category (Gold, Silver, Diamond, Platinum)"
                )

            # ====================================================
            # ROW 2
            # ====================================================

            col4, col5, col6 = st.columns(3)

            with col4:

                amount_col = st.selectbox(
                    "💰 Item Amount Column",
                    cols_list,
                    index=idx(
                        guess["amount"]
                    )
                )

            with col5:

                gross_wt_col = st.selectbox(
                    "⚖️ Gross Weight Column",
                    cols_list,
                    index=idx(
                        guess["gross_wt"]
                    )
                )

            with col6:

                diamond_wt_col = st.selectbox(
                    "💎 Diamond Weight Column",
                    cols_list,
                    index=idx(
                        guess["diamond_wt"]
                    )
                )

            # ====================================================
            # ITEM COLUMN (NEW)
            # ====================================================

            st.markdown(
                "#### 📦 Product Item Column"
            )

            # Auto-detect item column
            item_col_guess = None
            for col in cols_list:
                if col.lower() == "item":
                    item_col_guess = col
                    break
                elif "product" in col.lower() or "item" in col.lower():
                    item_col_guess = col
                    break

            item_col_options = [
                "None"
            ] + cols_list

            item_col_selected = st.selectbox(
                "📦 Product Item Column (e.g., RING, NECKLACE, etc.)",
                item_col_options,
                index=(
                    item_col_options.index(item_col_guess)
                    if item_col_guess in item_col_options
                    else 0
                ),
                help="Select the column that contains the actual product names like RING, NECKLACE, etc."
            )

            if item_col_selected == "None":
                item_col = None
            else:
                item_col = item_col_selected

            # ====================================================
            # DIAMOND AMOUNT
            # ====================================================

            st.markdown(
                "#### 💎 Diamond Amount Column"
            )

            diamond_amount_options = [
                "None"
            ] + cols_list

            detected_diamond_amount = (
                guess["diamond_amount"]
                if guess["diamond_amount"]
                in cols_list
                else "None"
            )

            diamond_amount_col_selected = st.selectbox(
                "💎 Separate Diamond Amount Column",
                diamond_amount_options,
                index=(
                    diamond_amount_options.index(
                        detected_diamond_amount
                    )
                    if detected_diamond_amount
                    in diamond_amount_options
                    else 0
                ),
                help=(
                    "Select the Diamond Amount column. "
                    "For example: Diamond Amount = ₹5,740.00"
                )
            )

            if (
                diamond_amount_col_selected
                == "None"
            ):

                diamond_amount_col = None

            else:

                diamond_amount_col = (
                    diamond_amount_col_selected
                )

            # ====================================================
            # SHOW LOGIC
            # ====================================================

            st.info(
                """
                **Report Logic:**

                🥇 **Gold Wt** = Actual Gold Gross Wt ONLY

                💎 **Diamond Gross Wt (g)** = Gross Wt of diamond items only

                💎 **Diamond Wt (Ct)** = ALL diamond weights from ALL categories
                (Platinum + Diamond, Gold + Diamond, etc.)

                💎 **Diamond Amt** = Diamond Amount column

                💎 **Diamond Item Amt** = Item Amount of Diamond jewellery

                🥈 **Silver Wt** = Silver Gross Wt

                🔹 **Platinum Wt** = Platinum Gross Wt

                📊 **Percentages are WEIGHT WISE** (branch weight ÷ total weight of that category × 100)

                📋 **Raw Data Sheet** - All transaction details including:
                - Date
                - Branch Location
                - **Product Item** - Original product name from the Item column (e.g., RING, NECKLACE)
                - **Category** - Detected category (Gold, Silver, Diamond, Platinum)
                - Gross Wt, Stone Wt, Diamond Weight Ct, Net Wt, Stone Am, Diamond Amount, Item Amount
                - Label Location - Original label used for detection
                """
            )

            # ====================================================
            # GENERATE
            # ====================================================

            if st.button(
                "🚀 Generate Branch-wise Report",
                type="primary",
                use_container_width=True
            ):

                with st.spinner(
                    "Analyzing and generating reports..."
                ):

                    try:

                        results = process_data_cached(
                            df,
                            date_col,
                            branch_col,
                            label_col,
                            amount_col,
                            gross_wt_col,
                            diamond_wt_col,
                            diamond_amount_col,
                            item_col,
                            fill_blanks
                        )

                        st.session_state[
                            "results"
                        ] = results

                        st.session_state[
                            "report_generated"
                        ] = True

                    except Exception as gen_err:

                        st.session_state[
                            "report_generated"
                        ] = False

                        st.error(
                            f"❌ Error while generating report: {gen_err}"
                        )

                        import traceback

                        st.code(
                            traceback.format_exc()
                        )

            # ====================================================
            # DISPLAY RESULTS
            # ====================================================

            if (
                st.session_state.get(
                    "report_generated"
                )
                and
                "results"
                in st.session_state
            ):

                results = (
                    st.session_state[
                        "results"
                    ]
                )

                if results[
                    "filled_count"
                ] > 0:

                    st.success(
                        f"✅ Filled "
                        f"{results['filled_count']} "
                        "blank Branch Location cells"
                    )

                # =================================================
                # SAMPLE DETECTION
                # =================================================

                st.subheader(
                    "🔍 Sample Detection Results"
                )

                st.dataframe(
                    results["sample_data"],
                    use_container_width=True
                )

                # =================================================
                # DETECTION STATS
                # =================================================

                st.subheader(
                    "📊 Detection Statistics"
                )

                st.dataframe(
                    results["stats_df"],
                    use_container_width=True
                )

                # =================================================
                # BRANCH MAPPING
                # =================================================

                st.subheader(
                    "🏢 Branch Code Mapping Results"
                )

                st.dataframe(
                    results["branch_mapping"],
                    use_container_width=True
                )

                # =================================================
                # UNKNOWN
                # =================================================

                if len(
                    results["unknown_data"]
                ) > 0:

                    with st.expander(
                        f"⚠️ "
                        f"{len(results['unknown_data'])} "
                        "Unknown Items - Click to view"
                    ):

                        st.dataframe(
                            results["unknown_data"]
                        )

                st.success(
                    f"✅ Processing "
                    f"{results['n_rows_filtered']} rows "
                    f"from {results['n_branches']} branches"
                )

                # =================================================
                # REPORT
                # =================================================

                st.subheader(
                    "🏢 Branch-wise Cumulative Report"
                )

                report_display = (
                    results["report_display"]
                )

                format_dict = {}

                for col in report_display.columns:

                    if col == "Branch":
                        continue

                    if "%" in col:

                        format_dict[col] = (
                            "{:.1f}%"
                        )

                    elif (
                        "Amt" in col
                        or "Amount" in col
                    ):

                        format_dict[col] = (
                            "₹{:,.2f}"
                        )

                    elif (
                        "Wt" in col
                        or "Weight" in col
                    ):

                        format_dict[col] = (
                            "{:,.3f}"
                        )

                st.dataframe(
                    report_display.style.format(
                        format_dict
                    ),
                    use_container_width=True
                )

                # =================================================
                # DOWNLOAD CSV
                # =================================================

                csv_buffer = io.StringIO()

                report_display.to_csv(
                    csv_buffer,
                    index=False
                )

                st.download_button(
                    label="📥 Download Summary as CSV",
                    data=csv_buffer.getvalue(),
                    file_name=(
                        "Branch_Summary_"
                        +
                        datetime.now().strftime(
                            "%Y%m%d_%H%M"
                        )
                        +
                        ".csv"
                    ),
                    mime="text/csv",
                    key="download_csv_btn",
                    use_container_width=True
                )

                # =================================================
                # VISUALIZATION
                # =================================================

                st.subheader(
                    "📊 Branch-wise Visualization"
                )

                chart_data = report_display[
                    report_display["Branch"]
                    != "GRAND TOTAL"
                ].copy()

                if len(chart_data) > 0:

                    # Percentage
                    pct_cols = [
                        "Gold %",
                        "Silver %",
                        "Diamond %",
                        "Platinum %"
                    ]

                    pct_cols_exist = [
                        c
                        for c in pct_cols
                        if c in chart_data.columns
                    ]

                    if (
                        pct_cols_exist
                        and
                        chart_data[
                            pct_cols_exist
                        ].sum().sum() > 0
                    ):

                        st.markdown(
                            "#### Percentage Distribution by Branch"
                        )

                        st.bar_chart(
                            chart_data.set_index(
                                "Branch"
                            )[
                                pct_cols_exist
                            ]
                        )

                    # Weight
                    weight_cols = [
                        "Gold Wt (g)",
                        "Silver Wt (g)",
                        "Diamond Wt (Ct)",
                        "Platinum Wt (g)"
                    ]

                    weight_cols_exist = [
                        c
                        for c in weight_cols
                        if c in chart_data.columns
                    ]

                    if (
                        weight_cols_exist
                        and
                        chart_data[
                            weight_cols_exist
                        ].sum().sum() > 0
                    ):

                        st.markdown(
                            "#### Weight by Branch"
                        )

                        st.bar_chart(
                            chart_data.set_index(
                                "Branch"
                            )[
                                weight_cols_exist
                            ]
                        )

                    # Item Amount
                    amount_cols = [
                        "Gold Item Amt",
                        "Silver Item Amt",
                        "Diamond Item Amt",
                        "Platinum Item Amt"
                    ]

                    amount_cols_exist = [
                        c
                        for c in amount_cols
                        if c in chart_data.columns
                    ]

                    if (
                        amount_cols_exist
                        and
                        chart_data[
                            amount_cols_exist
                        ].sum().sum() > 0
                    ):

                        st.markdown(
                            "#### Item Amount by Branch"
                        )

                        st.bar_chart(
                            chart_data.set_index(
                                "Branch"
                            )[
                                amount_cols_exist
                            ]
                        )

                # =================================================
                # DAY WISE
                # =================================================

                st.subheader(
                    "📅 Day-wise Summary"
                )

                if len(
                    results["day_wise"]
                ) > 0:

                    day_tab1, day_tab2 = st.tabs(
                        [
                            "💰 Amount",
                            "⚖️ Weight"
                        ]
                    )

                    with day_tab1:

                        st.dataframe(
                            results[
                                "day_wise"
                            ].style.format(
                                "₹{:,.2f}"
                            ),
                            use_container_width=True
                        )

                    with day_tab2:

                        st.dataframe(
                            results[
                                "day_wise_weight"
                            ].style.format(
                                "{:,.3f}"
                            ),
                            use_container_width=True
                        )

                else:

                    st.info(
                        "No day-wise data available"
                    )

                # =================================================
                # DOWNLOAD EXCEL
                # =================================================

                st.divider()

                st.subheader(
                    "📥 Download Formatted Report"
                )

                st.download_button(
                    label=(
                        "📥 Download Branch-wise "
                        "Excel Report"
                    ),
                    data=results[
                        "formatted_bytes"
                    ],
                    file_name=(
                        "Branch_Sales_Report_"
                        +
                        datetime.now().strftime(
                            "%Y%m%d_%H%M"
                        )
                        +
                        ".xlsx"
                    ),
                    mime=(
                        "application/"
                        "vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="primary",
                    use_container_width=True,
                    key="download_report_btn"
                )

                # =================================================
                # SUMMARY METRICS
                # =================================================

                st.divider()

                st.subheader(
                    "📈 Summary"
                )

                total_by_item = (
                    results["total_by_item"]
                )

                metrics_cols = st.columns(4)

                icons = {
                    "Gold": "🥇",
                    "Silver": "🥈",
                    "Diamond": "💎",
                    "Platinum": "🔹"
                }

                for metric_idx, item in enumerate(
                    [
                        "Gold",
                        "Silver",
                        "Diamond",
                        "Platinum"
                    ]
                ):

                    with metrics_cols[
                        metric_idx
                    ]:

                        if item in total_by_item.index:

                            item_amount = (
                                total_by_item.loc[
                                    item,
                                    "Item Amount"
                                ]
                            )

                            weight = (
                                total_by_item.loc[
                                    item,
                                    "Weight"
                                ]
                            )

                            diamond_amount = (
                                total_by_item.loc[
                                    item,
                                    "Diamond Amount"
                                ]
                            )

                            if item == "Diamond":
                                if "Gross Weight" in total_by_item.columns:
                                    gross_weight = total_by_item.loc[item, "Gross Weight"]
                                    st.metric(
                                        f"{icons[item]} Diamond",
                                        f"₹{item_amount:,.0f}",
                                        delta=(
                                            f"{gross_weight:.3f} g Gross | "
                                            f"{weight:.3f} Ct (ALL categories) | "
                                            f"Diamond Amt ₹{diamond_amount:,.0f}"
                                        )
                                    )
                                else:
                                    st.metric(
                                        f"{icons[item]} Diamond",
                                        f"₹{item_amount:,.0f}",
                                        delta=(
                                            f"{weight:.3f} Ct (ALL categories) | "
                                            f"Diamond Amt ₹{diamond_amount:,.0f}"
                                        )
                                    )

                            else:
                                st.metric(
                                    f"{icons[item]} {item}",
                                    f"₹{item_amount:,.0f}",
                                    delta=(
                                        f"{weight:.3f} g"
                                    )
                                )

                st.success(
                    "✅ Report generated successfully!"
                )

        except Exception as e:

            st.error(
                f"❌ Error: {str(e)}"
            )

            import traceback

            st.code(
                traceback.format_exc()
            )


    # ============================================================
    # NO FILE
    # ============================================================

    else:

        st.session_state[
            "report_generated"
        ] = False

        st.session_state.pop(
            "results",
            None
        )

        st.info(
            "👈 Please upload your Excel/CSV file "
            "from the sidebar to begin"
        )

        with st.expander(
            "📋 How This Works"
        ):

            st.markdown(
                """
                ### 🔍 Item Detection

                The app detects the category from
                **Label Location** (not from the Item column).

                - `V- SILVER` → **Silver**
                - `M- MAIN COUNTER` → **Gold**
                - `A- ANTIQUE COUNTER` → **Gold**
                - `D- DESIGNER COUNTER` → **Gold**
                - `B- BOUTIQUE COUNTER` → **Gold**
                - `R-DIAMONDS COUNTER` → **Diamond**
                - `P- PLATINUM` → **Platinum**
                - `REPAIR` → **Skip**

                ---

                ### 📦 Product Item Column

                The **Product Item** column in the Raw Data sheet shows the actual product name
                from your **Item** column (e.g., RING, NECKLACE, BRACELET, etc.).

                This is separate from the **Category** column which shows the detected category
                (Gold, Silver, Diamond, Platinum).

                ---

                ### 💎 Diamond Jewellery Handling

                **Gold Wt (g)** = Actual Gold Gross Wt ONLY

                **Diamond Gross Wt (g)** = Gross Wt of diamond items only

                **Diamond Wt (Ct)** = ALL diamond weights from ALL categories
                (Platinum + Diamond, Gold + Diamond, etc.)

                Example:

                Platinum + Diamond transaction:

                - Gross Wt = 9.39 g
                - Diamond Wt (Ct) = 1.20 Ct

                Report:

                - Platinum Wt (g) → **9.39 g**
                - Diamond Gross Wt (g) → **0 g** (not a diamond item)
                - Diamond Wt (Ct) → **1.20 Ct** (included in total)

                Example with Gold + Diamond:

                - Gross Wt = 15.00 g
                - Diamond Wt (Ct) = 0.50 Ct

                Report:

                - Gold Wt (g) → **15.00 g**
                - Diamond Gross Wt (g) → **0 g** (not a diamond item)
                - Diamond Wt (Ct) → **0.50 Ct** (included in total)

                **Gold Wt (g) and Diamond Gross Wt (g) are now separate!**

                **Diamond Wt (Ct) = SUM of all Diamond Wt (Ct) from ALL categories!**

                ---

                ### 📊 Report Structure

                **Gold**
                - Gold Wt (g)
                - Gold Item Amt
                - Gold %

                **Silver**
                - Silver Wt (g)
                - Silver Item Amt
                - Silver %

                **Diamond**
                - Diamond Gross Wt (g)
                - Diamond Wt (Ct) (ALL categories)
                - Diamond Amt
                - Diamond Item Amt
                - Diamond %

                **Platinum**
                - Platinum Wt (g)
                - Platinum Item Amt
                - Platinum %

                ---

                ### 📌 Important Rule

                **Gold Wt (g)** = Actual Gold Gross Wt ONLY

                **Diamond Gross Wt (g)** = Gross Wt of diamond items only

                **Diamond Wt (Ct)** = SUM of Diamond Wt (Ct) from ALL rows where Diamond Wt (Ct) > 0

                **Gold and Diamond weights are now SEPARATE!**

                ---

                ### 📈 Percentage

                Percentages are calculated **weight wise**
                (not amount wise).

                For example:

                Diamond % =
                Branch Diamond Weight (Ct) /
                Total Diamond Weight (Ct) × 100

                Gold % =
                Branch Gold Wt (g) /
                Total Gold Wt (g) × 100

                Grand Total percentage = **100%**.

                ---

                ### 📍 Branch Order

                1. Madurai
                2. Anna Nagar-Madurai
                3. Marthandam
                4. Salem
                5. Tirunelveli
                6. Trichy
                7. Rajapalayam
                8. Dindigul
                9. Noida
                10. Virudhunagar

                ---

                ### 📋 Raw Data Sheet

                The Excel report includes a third sheet with all transaction details:

                - **Date** - Transaction date
                - **Branch Location** - Branch name
                - **Product Item** - Original product name from the Item column (e.g., RING, NECKLACE)
                - **Category** - Detected item category (Gold, Silver, Diamond, Platinum)
                - **Gross Wt** - Gross weight in grams
                - **Stone Wt** - Stone weight (same as Diamond Weight for diamond items)
                - **Diamond Weight Ct** - Diamond weight in carats
                - **Net Wt** - currently equal to Gross Wt (not yet netted of stone weight - see code comment)
                - **Stone Am** - Stone amount (same as Diamond Amount)
                - **Diamond Amount** - Separate diamond amount
                - **Item Amount** - Item sales amount
                - **Label Location** - Original label location value (used for detection)
                """
            )
    st.divider()
    st.caption(
        "Built with ❤️ using Streamlit | "
        "Branch-wise Cumulative & Day-wise Report | "
        "Gold Wt = Actual Gold Gross Wt ONLY | "
        "Diamond Gross Wt = Gross Wt of diamond items only | "
        "Diamond Wt (Ct) = ALL diamond weights from ALL categories | "
        "Diamond Amount shown separately | "
        "Raw Data sheet with Product Item (from Item column) and Category (detected)"
    )
