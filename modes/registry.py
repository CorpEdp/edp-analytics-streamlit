"""Mode metadata plus runtime discovery for the unified Streamlit app."""

from __future__ import annotations

import ast
import importlib
import pkgutil
import re
from pathlib import Path

MODE_REGISTRY = [
    {
        "id": "home",
        "name": "Dashboard / Home",
        "description": "Unified EDP Analytics home screen",
        "category": "General",
        "module": "modes.home",
        "source": "app.py",
    },
    {
        "id": "audit_scheme_participation_transaction",
        "name": "Audit: Scheme Participation and Transaction",
        "description": "Legacy Streamlit mode converted from App-Audit-SchemeParticipation&Transaction.py",
        "category": "App",
        "module": "modes.audit_scheme_participation_transaction",
        "source": "modules/App-Audit-SchemeParticipation&Transaction.py",
    },
    {
        "id": "branch_employee_wise_referral_report",
        "name": "Branch Employee Wise Referral Report",
        "description": "Legacy Streamlit mode converted from App-BranchEmployeeWiseReferralReport.py",
        "category": "App",
        "module": "modes.branch_employee_wise_referral_report",
        "source": "modules/App-BranchEmployeeWiseReferralReport.py",
    },
    {
        "id": "campaign_checkings_enroll_collection",
        "name": "Campaign Checkings Enrollment and Collection",
        "description": "Legacy Streamlit mode converted from App-CampaignCheckingsEnroll&Collection.py",
        "category": "Campaign",
        "module": "modes.campaign_checkings_enroll_collection",
        "source": "modules/App-CampaignCheckingsEnroll&Collection.py",
    },
    {
        "id": "campaign_reconciliation_studio",
        "name": "Campaign Reconciliation Studio",
        "description": "Legacy Streamlit mode converted from App-CampaignReconciliationStudio.py",
        "category": "Campaign",
        "module": "modes.campaign_reconciliation_studio",
        "source": "modules/App-CampaignReconciliationStudio.py",
    },
    {
        "id": "campaign_transaction_analyzer",
        "name": "Campaign Transaction Analyzer",
        "description": "Legacy Streamlit mode converted from App-CampaignTransactionAnalyzer.py",
        "category": "Campaign",
        "module": "modes.campaign_transaction_analyzer",
        "source": "modules/App-CampaignTransactionAnalyzer.py",
    },
    {
        "id": "customer_register_location_find",
        "name": "Customer Register Location Find",
        "description": "Legacy Streamlit mode converted from App-CustomerRegisterLocationFind.py",
        "category": "Customer",
        "module": "modes.customer_register_location_find",
        "source": "modules/App-CustomerRegisterLocationFind.py",
    },
    {
        "id": "daily_scheme_payment_frequency_tracker_slap_wise",
        "name": "Daily Scheme Payment Frequency Tracker (Slap Wise)",
        "description": "Legacy Streamlit mode converted from App-DailySchemePaymentFrequencyTrackerSlapWise.py",
        "category": "App",
        "module": "modes.daily_scheme_payment_frequency_tracker_slap_wise",
        "source": "modules/App-DailySchemePaymentFrequencyTrackerSlapWise.py",
    },
    {
        "id": "day_scheme_enrollment_collection_report",
        "name": "Day Scheme Enrollment and Collection Report",
        "description": "Legacy Streamlit mode converted from App-Day-SchemeEnrollmentCollectionReport.py",
        "category": "App",
        "module": "modes.day_scheme_enrollment_collection_report",
        "source": "modules/App-Day-SchemeEnrollmentCollectionReport.py",
    },
    {
        "id": "due_remaining_report",
        "name": "Due Remaining Report",
        "description": "Legacy Streamlit mode converted from App-DueRemainingReport.py",
        "category": "ERP",
        "module": "modes.due_remaining_report",
        "source": "modules/App-DueRemainingReport.py",
    },
    {
        "id": "due_remaining_report_checking",
        "name": "Due Remaining Report Checking",
        "description": "Legacy Streamlit mode converted from App-DueRemainingReportChecking.py",
        "category": "ERP",
        "module": "modes.due_remaining_report_checking",
        "source": "modules/App-DueRemainingReportChecking.py",
    },
    {
        "id": "e_gold_customer_payment_days_followup_report",
        "name": "E-Gold Customer Payment Days Followup Report",
        "description": "Legacy Streamlit mode converted from App-E-GoldCustomerPaymentDaysFollowupReport.py",
        "category": "App",
        "module": "modes.e_gold_customer_payment_days_followup_report",
        "source": "modules/App-E-GoldCustomerPaymentDaysFollowupReport.py",
    },
    {
        "id": "gateway_cashfree_payment_mode_analytics",
        "name": "Gateway Cashfree Payment Mode Analytics",
        "description": "Legacy Streamlit mode converted from App-GatewayCashfreePaymentModeAnalytics.py",
        "category": "Cashfree",
        "module": "modes.gateway_cashfree_payment_mode_analytics",
        "source": "modules/App-GatewayCashfreePaymentModeAnalytics.py",
    },
    {
        "id": "month_scheme_enrollment_collection_report",
        "name": "Month Scheme Enrollment and Collection Report",
        "description": "Legacy Streamlit mode converted from App-Month-SchemeEnrollmentCollectionReport.py",
        "category": "App",
        "module": "modes.month_scheme_enrollment_collection_report",
        "source": "modules/App-Month-SchemeEnrollmentCollectionReport.py",
    },
    {
        "id": "our_team_wise_call_tracking_status",
        "name": "Our Team Wise Call Tracking Status",
        "description": "Legacy Streamlit mode converted from App-OurTeamWiseCallTrackingStatus.py",
        "category": "Team",
        "module": "modes.our_team_wise_call_tracking_status",
        "source": "modules/App-OurTeamWiseCallTrackingStatus.py",
    },
    {
        "id": "our_teamwise_referral_status",
        "name": "Our Teamwise Referral Status",
        "description": "Legacy Streamlit mode converted from App-OurTeamwiseReferralStatus.py",
        "category": "Team",
        "module": "modes.our_teamwise_referral_status",
        "source": "modules/App-OurTeamwiseReferralStatus.py",
    },
    {
        "id": "rate_wise_customer_trackers",
        "name": "Rate Wise Customer Trackers",
        "description": "Legacy Streamlit mode converted from App-RateWiseCustomerTrackers.py",
        "category": "App",
        "module": "modes.rate_wise_customer_trackers",
        "source": "modules/AppRateWiseCustomerTrackers.py",
    },
    {
        "id": "bhima_jewellery_customer_payment_logimax_cashfree",
        "name": "Bhima Jewellery Customer Payment Logimax Cashfree",
        "description": "Legacy Streamlit mode converted from ERP-BhimaJewelleryCustomerPaymentLogimaxCashfree.py",
        "category": "ERP",
        "module": "modes.bhima_jewellery_customer_payment_logimax_cashfree",
        "source": "modules/ERP-BhimaJewelleryCustomerPaymentLogimaxCashfree.py",
    },
    {
        "id": "cashfree_reconciliation",
        "name": "Cashfree Reconciliation",
        "description": "Legacy Streamlit mode converted from ERP-CashfreeReconciliation.py",
        "category": "ERP",
        "module": "modes.cashfree_reconciliation",
        "source": "modules/ERP-CashfreeReconciliation.py",
    },
    {
        "id": "item_category_weight_range_wastage_check",
        "name": "Item Category Weight Range Wastage Check",
        "description": "Legacy Streamlit mode converted from ERP-ItemCategoryWeightRangeWastageCheck.py",
        "category": "ERP",
        "module": "modes.item_category_weight_range_wastage_check",
        "source": "modules/ERP-ItemCategoryWeightRangeWastageCheck.py",
    },
    {
        "id": "sales_daywise_report",
        "name": "Sales Daywise Report",
        "description": "Legacy Streamlit mode converted from ERP-SalesDaywiseReport.py",
        "category": "ERP",
        "module": "modes.sales_daywise_report",
        "source": "modules/ERP-SalesDaywiseReport.py",
    },
    {
        "id": "numbers_dublicate_removed",
        "name": "Numbers Dublicate Removed",
        "description": "Legacy Streamlit mode converted from Custom-NumbersDublicateRemoved.py",
        "category": "Custom",
        "module": "modes.numbers_dublicate_removed",
        "source": "modules/Custom-NumbersDublicateRemoved.py",
    },
        {
        "id": "scheme_joining_cashfree_payment_details_reconciliation",
        "name": "Scheme Joining Cashfree Payment Details Reconciliation",
        "description": "Legacy Streamlit mode converted from ERP-SchemeJoiningCashfreePaymentDetailsReconciliation.py",
        "category": "ERP",
        "module": "modes.scheme_joining_cashfree_payment_details_reconciliation",
        "source": "modules/ERP-SchemeJoiningCashfreePaymentDetailsReconciliation.py",
    },
        {
        "id": "employee_referral_performance_report",
        "name": "Employee Referral PerformanceReport",
        "description": "Legacy Streamlit mode converted from App-EmployeeReferralPerformanceReport.py",
        "category": "App",
        "module": "modes.employee_referral_performance_report",
        "source": "modules/App-EmployeeReferralPerformanceReport.py",
        },
]

def _display_name(module_name: str) -> str:
    return re.sub(r"\s+", " ", module_name.replace("_", " ")).strip().title()


def _mode_type(path: Path) -> str:
    try:
        tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
    except SyntaxError:
        return "script"

    has_entrypoint = any(
        isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name in {"render", "run"}
        for node in tree.body
    )
    return "callable" if has_entrypoint else "script"


def get_mode_registry() -> list[dict[str, str]]:
    """Return saved metadata plus every valid mode file currently on disk."""
    importlib.invalidate_caches()

    modes_by_module = {mode["module"]: mode.copy() for mode in MODE_REGISTRY}
    import modes

    for _, module_name, is_package in pkgutil.iter_modules(modes.__path__):
        if (
            is_package
            or module_name in {"registry", "home"}
            or not module_name.isidentifier()
        ):
            continue

        qualified_name = f"modes.{module_name}"
        if qualified_name in modes_by_module:
            continue

        source_path = Path(modes.__file__).parent / f"{module_name}.py"
        modes_by_module[qualified_name] = {
            "id": module_name,
            "name": _display_name(module_name),
            "description": "Automatically discovered Streamlit mode",
            "category": "Custom",
            "module": qualified_name,
            "source": f"modes/{module_name}.py",
            "mode_type": _mode_type(source_path),
        }

    return list(modes_by_module.values())
