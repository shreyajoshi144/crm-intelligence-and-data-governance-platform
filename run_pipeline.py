"""
PIPELINE RUNNER - Enterprise CRM Governance & Analytics Platform

Run this file to execute Modules 1 -> 8 end-to-end.

Usage:
    python run_pipeline.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.module1_ingestion import run_ingestion
from modules.module2_standardization import run_standardization
from modules.module3_entity_resolution import run_entity_resolution
from modules.module4_deduplication import run_deduplication
from modules.module5_data_quality import run_data_quality
from modules.module6_golden_record_engine import run_golden_record_engine
from modules.module7_governance_analytics import run_governance_analytics
from modules.module8_business_analytics import run_business_analytics


def main():
    print("\n" + "=" * 70)
    print("  ENTERPRISE CRM GOVERNANCE & ANALYTICS PLATFORM")
    print("  Pipeline Runner - Modules 1 -> 8")
    print("=" * 70 + "\n")

    staging_data = run_ingestion()
    print()

    std_data = run_standardization(staging_data)
    print()

    entity_match_table = run_entity_resolution(std_data)
    print()

    dedup_data = run_deduplication(std_data, entity_match_table)
    print()

    dq_data = run_data_quality(dedup_data, std_data)
    print()

    golden_data = run_golden_record_engine(dedup_data, dq_data, std_data)
    print()

    governance_data = run_governance_analytics(dq_data, golden_data, dedup_data, std_data)
    print()

    business_data = run_business_analytics(golden_data)
    print()

    print("\n" + "=" * 70)
    print("  PIPELINE COMPLETE  (Modules 1 -> 8)")
    print("=" * 70)

    print("\nStaging tables         (Module 1):")
    for name, df in staging_data.items():
        print(f"  {name:<30} {len(df):>8,} rows")

    print("\nStandardised tables    (Module 2):")
    for name, df in std_data.items():
        print(f"  {name:<30} {len(df):>8,} rows")

    print(f"\nEntity match table     (Module 3):  {len(entity_match_table):,} matches")

    print("\nDeduplicated tables    (Module 4):")
    for name, df in dedup_data.items():
        print(f"  {name:<30} {len(df):>8,} rows")

    print("\nData quality outputs   (Module 5):")
    print(f"  CRM Health Score              {dq_data['crm_health_score']:>8.1f}%")
    print(f"  DQ Issues                     {len(dq_data['issue_df']):>8,} rows")
    print(f"  Account Quality Scores        {len(dq_data['quality_scores_accounts']):>8,} rows")
    print(f"  Contact Quality Scores        {len(dq_data['quality_scores_contacts']):>8,} rows")
    print(f"  Lead Quality Scores           {len(dq_data['quality_scores_leads']):>8,} rows")

    print("\nGolden records         (Module 6):")
    print(f"  golden_accounts               {len(golden_data['golden_accounts']):>8,} rows")
    print(f"  golden_contacts               {len(golden_data['golden_contacts']):>8,} rows")
    print(f"  golden_leads                  {len(golden_data['golden_leads']):>8,} rows")

    print("\nGovernance analytics   (Module 7):")
    print(f"  governance_metrics            {len(governance_data['governance_metrics']):>8,} rows")
    print(f"  governance_report             {governance_data['governance_report_path']}")

    print("\nBusiness analytics     (Module 8):")
    print(f"  business_analytics_dataset    {len(business_data['business_analytics_dataset']):>8,} rows")
    print(f"  lead_funnel_summary           {len(business_data['lead_funnel']):>8,} rows")
    print(f"  lead_source_performance       {len(business_data['lead_source_performance']):>8,} rows")
    print(f"  business_analytics_report     {business_data['business_analytics_report']}")

    print("\nDeliverables written:")
    print("  logs/module1_ingestion.log")
    print("  logs/module2_standardization.log")
    print("  logs/module3_entity_resolution.log")
    print("  logs/module4_deduplication.log")
    print("  logs/module5_data_quality.log")
    print("  logs/module6_golden_record.log")
    print("  logs/module7_governance.log")
    print("  logs/module8_business_analytics.log")
    print("  reports/schema_validation_report.xlsx")
    print("  reports/standardization_report.xlsx")
    print("  reports/entity_resolution_report.xlsx")
    print("  reports/deduplication_report.xlsx")
    print("  reports/data_quality_report.xlsx")
    print("  reports/golden_record_report.xlsx")
    print("  reports/governance_report.xlsx")
    print("  reports/business_analytics_report.xlsx")
    print("  data/processed/governance_metrics.csv")
    print("  data/processed/business_analytics_dataset.csv")
    print("  data/processed/lead_funnel_summary.csv")
    print("  data/processed/lead_source_performance.csv")
    print("  data/processed/industry_distribution.csv")
    print("  data/processed/country_distribution.csv")
    print("  data/processed/golden_accounts.csv")
    print("  data/processed/golden_contacts.csv")
    print("  data/processed/golden_leads.csv")
    print()


if __name__ == "__main__":
    main()
