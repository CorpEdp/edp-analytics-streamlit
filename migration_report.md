# Unified Streamlit Migration Audit

## Summary
This project was migrated from multiple standalone Streamlit script files into a single entry point. The original files remain under `modules/` and a backup copy was created under `legacy/`.

## Original file -> New mode mapping

| Original file | New mode | Status | Problems | Dependencies | Manual review required |
| --- | --- | --- | --- | --- | --- |
| App-Audit-SchemeParticipation&Transaction.py | modes.audit_scheme_participation_transaction | Safe adapter | Page config removed at runtime | pandas, streamlit, numpy | No |
| App-BranchEmployeeWiseReferralReport.py | modes.branch_employee_wise_referral_report | Safe adapter | None expected | pandas, streamlit, openpyxl | No |
| App-CampaignCheckingsEnroll&Collection.py | modes.campaign_checkings_enroll_collection | Safe adapter | None expected | pandas, streamlit, numpy | No |
| App-CampaignReconciliationStudio.py | modes.campaign_reconciliation_studio | Safe adapter | None expected | pandas, streamlit, plotly | No |
| App-CampaignTransactionAnalyzer.py | modes.campaign_transaction_analyzer | Safe adapter | None expected | pandas, streamlit, numpy | No |
| App-CustomerRegisterLocationFind.py | modes.customer_register_location_find | Safe adapter | None expected | pandas, streamlit | No |
| App-DailySchemePaymentFrequencyTrackerSlapWise.py | modes.daily_scheme_payment_frequency_tracker_slap_wise | Safe adapter | None expected | pandas, streamlit, plotly | No |
| App-Day-SchemeEnrollmentCollectionReport.py | modes.day_scheme_enrollment_collection_report | Safe adapter | None expected | pandas, streamlit, numpy | No |
| App-DueRemainingReport.py | modes.due_remaining_report | Safe adapter | None expected | pandas, streamlit, plotly | No |
| App-DueRemainingReportChecking.py | modes.due_remaining_report_checking | Safe adapter | None expected | pandas, streamlit | No |
| App-E-GoldCustomerPaymentDaysFollowupReport.py | modes.e_gold_customer_payment_days_followup_report | Safe adapter | None expected | pandas, streamlit, plotly | No |
| App-GatewayCashfreePaymentModeAnalytics.py | modes.gateway_cashfree_payment_mode_analytics | Safe adapter | None expected | pandas, streamlit | No |
| App-Month-SchemeEnrollmentCollectionReport.py | modes.month_scheme_enrollment_collection_report | Safe adapter | None expected | pandas, streamlit, numpy | No |
| App-OurTeamWiseCallTrackingStatus.py | modes.our_team_wise_call_tracking_status | Safe adapter | None expected | pandas, streamlit | No |
| App-OurTeamwiseReferralStatus.py | modes.our_teamwise_referral_status | Safe adapter | None expected | pandas, streamlit, plotly | No |
| AppRateWiseCustomerTrackers.py | modes.rate_wise_customer_trackers | Safe adapter | None expected | pandas, streamlit | No |
| ERP-BhimaJewelleryCustomerPaymentLogimaxCashfree.py | modes.bhima_jewellery_customer_payment_logimax_cashfree | Safe adapter | None expected | pandas, streamlit, numpy | No |
| ERP-CashfreeReconciliation.py | modes.cashfree_reconciliation | Safe adapter | None expected | pandas, streamlit | No |
| ERP-ItemCategoryWeightRangeWastageCheck.py | modes.item_category_weight_range_wastage_check | Safe adapter | None expected | pandas, streamlit, numpy | No |
| ERP-SalesDaywiseReport.py | modes.sales_daywise_report | Safe adapter | None expected | pandas, streamlit | No |
| NumbersDublicateRemoved.py | modes.numbers_dublicate_removed | Safe adapter | None expected | pandas, streamlit, numpy | No |

## Notes
- `st.set_page_config()` calls were treated as runtime-only page definitions and are stripped before executing the legacy script body.
- No legacy module is imported merely to populate the registry.
- Only the selected mode runs.
- Original files remain available for review under `legacy/` and `modules/`.
