# Merged Streamlit Toolkit — Setup Guide

## What's in this zip

```
my_project/
├── main.py              ← run this: `streamlit run main.py`
├── requirements.txt      ← starter list, replace with pipreqs output later
├── README.md             ← this file
└── modules/
    ├── __init__.py
    ├── audit_scheme_participation_transaction.py
    ├── branch_employee_wise_referral_report.py
    ├── campaign_checkins_enroll_collection.py
    ├── campaign_reconciliation_studio.py
    ├── campaign_transaction_analyzer.py
    ├── customer_register_location_find.py
    ├── daily_scheme_payment_frequency_tracker_slabwise.py
    ├── day_scheme_enrollment_collection_report.py
    ├── due_remaining_report.py
    ├── due_remaining_report_checking.py
    ├── egold_customer_payment_days_followup_report.py
    ├── gateway_cashfree_payment_mode_analytics.py
    ├── month_scheme_enrollment_collection_report.py
    ├── our_team_wise_call_tracking_status.py
    ├── our_team_wise_referral_status.py
    ├── rate_wise_customer_trackers.py
    ├── erp_bhima_jewellery_customer_payment_logimax_cashfree.py
    ├── erp_cashfree_reconciliation.py
    ├── erp_item_category_weight_range_wastage_check.py
    ├── erp_sales_daywise_report.py
    └── numbers_duplicate_removed.py
```

Each module file is currently a **stub** — it just shows an `st.info(...)`
placeholder. You need to paste your real code into each one.

## How to fill in each module (do this for all 21 files)

Open your old `App-XyzName.py` and the matching new stub side by side.

1. **Copy** everything from the old file.
2. **Delete** any line that calls `st.set_page_config(...)` — only
   `main.py` is allowed to call this, once.
3. **Paste** the rest of the old code *inside* the `def run():` function
   in the stub (indent it one level).
4. **Namespace** anything that touches shared state, so two modules can
   never collide:
   - `st.session_state["data"]` → `st.session_state["<modname>_data"]`
   - `st.button("Submit")` → `st.button("Submit", key="<modname>_submit")`
   - Do this for every `st.session_state[...]` and every widget that has
     (or needs) a `key=`.
5. Keep the `LABEL = "..."` line at the top — that's what shows up in the
   sidebar dropdown.

Repeat for all 21 files. You don't need to touch `main.py` at all — it
auto-discovers every file in `modules/`.

## Running it

```bash
cd my_project
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run main.py
```

## Adding a mode to the active app

The active entry point is `app.py`. To add a mode, paste a valid Python module
with either a `render()` or `run()` function into `modes/`, using a Python-safe
filename such as `customer_summary.py`. The sidebar discovers new files on the
next Streamlit rerun; no edit to `app.py` or `modes/registry.py` is required.

For a deployed Streamlit link, the new file must also reach the deployment
environment. With Streamlit Community Cloud, commit and push the file to the
connected repository; the deployment then reruns automatically. Pasting a
file only on your local computer cannot change a remote deployment.

## Once all 21 modules have real code: regenerate requirements.txt

The current `requirements.txt` is a reasonable starting guess. Replace it
with an accurate, auto-scanned version:

```bash
pip install pipreqs
pipreqs . --force
```

This reads every `import` across all 21 real modules and writes exactly
what's needed — no more, no less.

## If something breaks

- **"No module named modules"** → make sure you're running
  `streamlit run main.py` from *inside* `my_project/`, not from one level
  up or down.
- **Duplicate widget key error** → two modules have widgets with the same
  key (or no key at all with the same auto-generated one). Add unique
  `key="<modname>_..."` to the offending widgets.
- **A module doesn't show up in the sidebar** → check the sidebar's
  "⚠️ N module(s) failed to load" expander — it shows the exact import
  error for that file.
- **Data from one mode "leaking" into another** → you forgot to namespace
  a `st.session_state` key — see step 4 above.
