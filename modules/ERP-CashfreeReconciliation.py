import streamlit as st
import pandas as pd
from itertools import combinations
from io import BytesIO

st.set_page_config(page_title="Payment Reconciliation", layout="wide")
st.title("💳 External vs Internal Payment Reconciliation Dashboard")

# ============================================================
# HELPERS
# ============================================================

def read_file(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)

def clean_mobile(x):
    try:
        x = str(x)
        x = x.replace(".0", "").replace(" ", "").replace("-", "").replace("+91", "")
        x = ''.join(filter(str.isdigit, x))
        if len(x) > 10:
            x = x[-10:]
        return x
    except:
        return ""

def clean_amount(x):
    try:
        return round(float(x), 2)
    except:
        return 0.0

def clean_date(x):
    try:
        return pd.to_datetime(x, dayfirst=True, errors="coerce").date()
    except:
        return None

# ============================================================
# FILE UPLOAD
# ============================================================

external_file = st.file_uploader(
    "📤 Upload External Payment File",
    type=["xlsx", "csv"]
)

internal_file = st.file_uploader(
    "📤 Upload Internal Payment File",
    type=["xlsx", "csv"]
)

# ============================================================
# MAIN PROCESS
# ============================================================

if external_file and internal_file:

    external_payment = read_file(external_file)
    internal_payment = read_file(internal_file)

    st.success("✅ Files Uploaded Successfully")

    # ========================================================
    # COLUMN MAPPING UI
    # ========================================================

    st.subheader("🧩 Column Mapping")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### External Payment Columns")
        ext_phone_col = st.selectbox("Phone Column", external_payment.columns)
        ext_amount_col = st.selectbox("Amount Column", external_payment.columns)
        ext_date_col = st.selectbox("Date Column", external_payment.columns)

    with col2:
        st.markdown("### Internal Payment Columns")
        int_phone_col = st.selectbox("Phone Column ", internal_payment.columns)
        int_amount_col = st.selectbox("Amount Column ", internal_payment.columns)
        int_date_col = st.selectbox("Date Column ", internal_payment.columns)

    # ========================================================
    # STANDARDIZE DATA
    # ========================================================

    external_payment["Customer Phone"] = external_payment[ext_phone_col].apply(clean_mobile)
    internal_payment["mobile"] = internal_payment[int_phone_col].apply(clean_mobile)

    external_payment["Amount"] = external_payment[ext_amount_col].apply(clean_amount)
    internal_payment["payment_amount"] = internal_payment[int_amount_col].apply(clean_amount)

    external_payment["MATCH_DATE"] = external_payment[ext_date_col].apply(clean_date)
    internal_payment["MATCH_DATE"] = internal_payment[int_date_col].apply(clean_date)

    internal_payment["USED"] = False

    # ========================================================
    # DEBUG DATE CHECK
    # ========================================================

    with st.expander("🔍 Debug Date Parsing"):
        st.write("External Date Sample:")
        st.write(external_payment[[ext_date_col, "MATCH_DATE"]].head())
        st.write("Internal Date Sample:")
        st.write(internal_payment[[int_date_col, "MATCH_DATE"]].head())

    # ========================================================
    # MATCHING LOGIC
    # ========================================================

    results = []
    exact_match_count = 0
    split_match_count = 0
    unmatched_count = 0
    split_internal_count = 0

    progress_bar = st.progress(0)
    status_text = st.empty()

    total_external = len(external_payment)

    for idx, (_, ext_row) in enumerate(external_payment.iterrows()):

        # Update progress
        progress = (idx + 1) / total_external
        progress_bar.progress(progress)
        status_text.text(f"Processing {idx + 1} of {total_external} external payments...")

        phone = ext_row["Customer Phone"]
        amount = ext_row["Amount"]
        date = ext_row["MATCH_DATE"]

        matched = False

        # EXACT MATCH
        exact_matches = internal_payment[
            (internal_payment["mobile"] == phone) &
            (internal_payment["payment_amount"] == amount) &
            (internal_payment["MATCH_DATE"] == date) &
            (internal_payment["USED"] == False)
        ]

        if not exact_matches.empty:

            idx_match = exact_matches.index[0]
            internal_payment.loc[idx_match, "USED"] = True

            merged = {}

            for col in external_payment.columns:
                merged[f"EXT_{col}"] = ext_row[col]

            for col in internal_payment.columns:
                merged[f"INT_{col}"] = exact_matches.iloc[0][col]

            merged["MATCH_TYPE"] = "EXACT MATCH"

            results.append(merged)

            exact_match_count += 1
            matched = True

        # SPLIT MATCH - UP TO 6 INTERNAL PAYMENTS
        if not matched:

            candidates = internal_payment[
                (internal_payment["mobile"] == phone) &
                (internal_payment["MATCH_DATE"] == date) &
                (internal_payment["USED"] == False)
            ]

            indices = list(candidates.index)
            found_combo = None
            combo_size = 0

            # Try combinations from 2 to 6 internal payments
            for r in range(2, min(7, len(indices) + 1)):
                
                # Optional: Show which combination size is being checked
                # status_text.text(f"Checking combinations of {r} payments...")
                
                for combo in combinations(indices, r):

                    if round(
                        internal_payment.loc[list(combo), "payment_amount"].sum(),
                        2
                    ) == amount:

                        found_combo = combo
                        combo_size = r
                        break

                if found_combo:
                    break

            if found_combo:

                internal_payment.loc[list(found_combo), "USED"] = True

                for _, row in internal_payment.loc[list(found_combo)].iterrows():

                    merged = {}

                    for col in external_payment.columns:
                        merged[f"EXT_{col}"] = ext_row[col]

                    for col in internal_payment.columns:
                        merged[f"INT_{col}"] = row[col]

                    merged["MATCH_TYPE"] = f"SPLIT MATCH ({combo_size} payments)"

                    results.append(merged)

                split_match_count += 1
                split_internal_count += len(found_combo)
                matched = True

        # UNMATCHED
        if not matched:

            merged = {}

            for col in external_payment.columns:
                merged[f"EXT_{col}"] = ext_row[col]

            merged["MATCH_TYPE"] = "UNMATCHED"

            results.append(merged)

            unmatched_count += 1

    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()

    final_df = pd.DataFrame(results)

    # ========================================================
    # DASHBOARD
    # ========================================================

    st.subheader("📊 Dashboard")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("✅ Exact Matches", exact_match_count)
    with col2:
        st.metric("🔗 Split Matches", split_match_count)
    with col3:
        st.metric("❌ Unmatched", unmatched_count)
    with col4:
        st.metric("📦 Internal Payments in Splits", split_internal_count)

    # Summary metrics
    total_matched = exact_match_count + split_match_count
    match_percentage = (total_matched / total_external * 100) if total_external > 0 else 0
    
    st.markdown("---")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("📊 Total External Payments", total_external)
    with col_b:
        st.metric("✅ Total Matched", total_matched)
    with col_c:
        st.metric("📈 Match Rate", f"{match_percentage:.1f}%")

    st.divider()

    # ========================================================
    # FILTERS
    # ========================================================

    st.subheader("🔍 Filter Results")

    filter_col1, filter_col2 = st.columns(2)
    
    with filter_col1:
        match_type_filter = st.multiselect(
            "Match Type",
            options=final_df["MATCH_TYPE"].unique(),
            default=final_df["MATCH_TYPE"].unique()
        )
    
    with filter_col2:
        if "EXT_Customer Phone" in final_df.columns:
            phone_filter = st.text_input("Filter by Phone Number (contains)")
        else:
            phone_filter = ""

    # Apply filters
    filtered_df = final_df[final_df["MATCH_TYPE"].isin(match_type_filter)]
    
    if phone_filter and "EXT_Customer Phone" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["EXT_Customer Phone"].astype(str).str.contains(phone_filter, na=False)]

    # ========================================================
    # TABLE
    # ========================================================

    st.subheader("📄 Reconciliation Result")
    
    st.info(f"Showing {len(filtered_df)} of {len(final_df)} rows")

    st.dataframe(
        filtered_df,
        width="stretch",
        height=500
    )

    # ========================================================
    # STATISTICS EXPANDER
    # ========================================================

    with st.expander("📊 Detailed Statistics"):
        
        st.write("### Match Type Distribution")
        match_distribution = final_df["MATCH_TYPE"].value_counts()
        st.bar_chart(match_distribution)
        
        st.write("### Split Match Sizes")
        split_sizes = final_df[final_df["MATCH_TYPE"].str.contains("SPLIT MATCH", na=False)]["MATCH_TYPE"].value_counts()
        if not split_sizes.empty:
            st.write(split_sizes)
        else:
            st.write("No split matches found")

    # ========================================================
    # EXPORT
    # ========================================================

    st.subheader("💾 Export Results")

    col_export1, col_export2 = st.columns(2)

    with col_export1:
        # Excel Export
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            filtered_df.to_excel(writer, sheet_name="Reconciliation Results", index=False)
            
            # Add summary sheet
            summary_data = {
                "Metric": ["Total External Payments", "Exact Matches", "Split Matches", "Unmatched", "Match Rate"],
                "Value": [total_external, exact_match_count, split_match_count, unmatched_count, f"{match_percentage:.1f}%"]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
        
        st.download_button(
            "📊 Download Excel Report",
            data=output.getvalue(),
            file_name="reconciliation_report.xlsx",
            width="stretch"
        )

    with col_export2:
        # CSV Export
        st.download_button(
            "📄 Download CSV",
            data=filtered_df.to_csv(index=False).encode('utf-8'),
            file_name="reconciliation_results.csv",
            mime="text/csv",
            width="stretch"
        )

    # ========================================================
    # SUCCESS MESSAGE
    # ========================================================
    
    st.success(f"✅ Reconciliation complete! Matched {total_matched} out of {total_external} external payments ({match_percentage:.1f}%)")

else:
    st.info("👈 Please upload both External and Internal payment files to begin reconciliation")
