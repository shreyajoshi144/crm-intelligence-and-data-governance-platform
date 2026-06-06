"""
Enterprise CRM Governance & Analytics Platform
Configuration File
"""

import os

# ─── Base Paths ───────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR    = os.path.join(BASE_DIR, "data")
RAW_DIR     = os.path.join(DATA_DIR, "raw")
STAGING_DIR = os.path.join(DATA_DIR, "staging")
PROCESSED_DIR = os.path.join(DATA_DIR, "processed")
LOGS_DIR    = os.path.join(BASE_DIR, "logs")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")

# ─── Source File ──────────────────────────────────────────────────────────────
SOURCE_FILE = os.path.join(RAW_DIR, "Enterprise_CRM_Governance_Dataset.xlsx")

# ─── Source Sheets → Staging Table Names ─────────────────────────────────────
SOURCE_SHEETS = {
    "Accounts":          "stg_accounts",
    "Contacts":          "stg_contacts",
    "Leads":             "stg_leads",
    "Data_Stewardship":  "stg_stewardship",
    "Golden_Record_Map": "stg_golden_map",
}

# ─── Expected Schemas (column → expected dtype) ───────────────────────────────
EXPECTED_SCHEMAS = {
    "Accounts": {
        "account_id":     "object",
        "account_name":   "object",
        "industry":       "object",
        "country":        "object",
        "annual_revenue": "float64",
        "employee_count": "float64",
    },
    "Contacts": {
        "contact_id": "object",
        "first_name":  "object",
        "last_name":   "object",
        "email":       "object",
        "job_title":   "object",
        "account_id":  "object",
    },
    "Leads": {
        "lead_id":   "object",
        "lead_name": "object",
        "source":    "object",
        "status":    "object",
    },
    "Data_Stewardship": {
        "account_id":       "object",
        "steward":          "object",
        "dq_issue_status":  "object",
    },
    "Golden_Record_Map": {
        "source_record":    "object",
        "account_id":       "object",
        "golden_record_id": "object",
    },
}

# ─── Standardization Rules ────────────────────────────────────────────────────

INDUSTRY_MAP = {
    "tech": "Technology", "it": "Technology", "information technology": "Technology",
    "software": "Technology", "technology": "Technology",
    "healthcare": "Healthcare", "health care": "Healthcare", "pharma": "Healthcare",
    "pharmaceutical": "Healthcare",
    "finance": "Finance", "financial services": "Finance", "banking": "Finance",
    "fintech": "Finance",
    "retail": "Retail", "e-commerce": "Retail", "ecommerce": "Retail",
    "manufacturing": "Manufacturing", "mfg": "Manufacturing",
    "education": "Education", "edtech": "Education",
    "media": "Media & Entertainment", "entertainment": "Media & Entertainment",
    "real estate": "Real Estate", "realty": "Real Estate",
    "consulting": "Consulting", "professional services": "Consulting",
    "logistics": "Logistics & Supply Chain", "supply chain": "Logistics & Supply Chain",
    "energy": "Energy", "oil & gas": "Energy", "utilities": "Energy",
}

COUNTRY_MAP = {
    "us": "USA", "u.s.": "USA", "u.s.a.": "USA", "united states": "USA",
    "united states of america": "USA",
    "uk": "United Kingdom", "u.k.": "United Kingdom", "great britain": "United Kingdom",
    "england": "United Kingdom",
    "in": "India", "ind": "India",
    "de": "Germany", "deu": "Germany",
    "au": "Australia", "aus": "Australia",
    "ca": "Canada", "can": "Canada",
    "sg": "Singapore", "sin": "Singapore",
    "ae": "UAE", "united arab emirates": "UAE",
    "fr": "France", "fra": "France",
    "jp": "Japan", "jpn": "Japan",
    "cn": "China", "chn": "China",
    "br": "Brazil", "bra": "Brazil",
}

VALID_LEAD_SOURCES  = ["Web", "Referral", "Event", "LinkedIn", "Cold Call",
                        "Partner", "Advertisement", "Other"]
VALID_LEAD_STATUSES = ["New", "Working", "Qualified", "Converted", "Unqualified",
                        "Nurturing"]
VALID_JOB_TITLES    = ["VP", "Director", "Analyst", "Manager", "Engineer",
                        "Consultant", "Executive", "Associate", "Intern", "Other"]
VALID_DQ_STATUSES   = ["Open", "In Review", "Resolved", "Escalated"]
