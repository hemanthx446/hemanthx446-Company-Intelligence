# hemanthx446 Company Intelligence

**Owner:** Hemanth  
**System Identity:** hemanthx446 Company Intelligence  
**System Purpose:** A personal, reusable, evidence-grounded corporate intelligence research system designed to investigate target companies across any industry. Operates autonomously to analyze business models, technological footprints, customer/partner ecosystems, key executives, public business signals, and commercial opportunities without speculative claims or ungrounded assumptions.

---

## 1. System Architecture & Directory Structure

```
hemanthx446-Company-Intelligence/
├── AGENTS.md                   # Permanent operating instructions and ethical safeguards
├── README.md                   # System documentation and architecture overview
├── RUN_RESEARCH.md             # Operator quick-start guide, commands, and launch prompt
├── requirements.txt            # Python dependencies (openpyxl, pytest)
│
├── config/
│   ├── research_profile.md     # Hemanth's research interests and analytical focal areas
│   └── research_settings.json  # Configurable investigation settings, scoring weights, allowances
│
├── templates/
│   ├── report_template.md      # Standardized 15-section markdown intelligence report template
│   ├── research_log_template.md # 13-stage chronological audit trail template
│   ├── source_register_template.csv # Source citation tabular schema
│   └── company_intelligence_schema.md # JSON data contract specification
│
├── workflows/
│   └── research_company.md     # Step-by-step 13-stage investigation workflow
│
├── scripts/
│   ├── generate_workbook.py    # Generates management-grade 12-worksheet company_intelligence.xlsx
│   └── validate_research.py    # Enforces data contracts, scoring math, and source integrity
│
├── research/                   # Snapshot storage for investigated companies (e.g., <company>_<date>/)
│   ├── .gitkeep
│   └── fictional_novafacture_20261026/ # Demonstration snapshot with complete sample data
│       ├── data.json
│       ├── report.md
│       ├── research_log.md
│       └── sources.csv
│
├── outputs/                    # Export directory for generated Excel workbooks
│   ├── .gitkeep
│   └── fictional_novafacture_20261026/
│       └── company_intelligence.xlsx
│
└── tests/                      # Automated regression tests
    ├── test_workbook.py        # Validates workbook generation, 12 sheets, links, and schema
    └── test_data/
        └── synthetic_company_data.json # Ground-truth synthetic test fixture
```

---

## 2. Core Operational Pillars

1. **Evidence Grounding & Six Standard Categories:**
   Every factual claim must cite a valid primary or corroborated source. Information is strictly classified into:
   - `Verified fact`
   - `Corroborated finding`
   - `Reasonable inference`
   - `Unverified claim`
   - `Historical information`
   - `Unknown information`
2. **Transparent 100-Point Opportunity Prioritization:**
   Commercial opportunities are scored across 5 dimensions:
   - **Strategic Fit (30 pts):** Alignment with research profile (SAP S/4HANA, BTP, MES, Cloud ALM, Advisory).
   - **Need Evidence (25 pts):** Publicly documented pain, modernization initiatives, or job postings.
   - **Account/Contact Access (20 pts):** Identified department and verifiable decision-maker.
   - **Commercial Relevance (15 pts):** Strategic scale, enterprise scope, and budget magnitude.
   - **Feasibility & Timing (10 pts):** Near-term readiness (0–6 month actionable window).
3. **12-Worksheet Excel Workbook (`company_intelligence.xlsx`):**
   Automated generation of an executive-ready workbook featuring:
   - Frozen header panes & auto-filters
   - Dynamic auto-fit column widths & text wrapping
   - Executive Summary dashboard with KPI cards and top opportunity table
   - Clickable source URLs and verified profile links
   - Zero hallucinated fields; explicit handling of unknown parameters

---

## 3. Quick-Start Execution

### Setup Environment
```bash
# Initialize virtual environment and install dependencies
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

### Generate Excel Workbook
```bash
.venv\Scripts\python scripts/generate_workbook.py --input research/<target_dir>/data.json --output outputs/<target_dir>/company_intelligence.xlsx
```

### Validate Research Dossier
```bash
.venv\Scripts\python scripts/validate_research.py --input research/<target_dir>/data.json --workbook outputs/<target_dir>/company_intelligence.xlsx
```

### Run Test Suite
```bash
.venv\Scripts\pytest tests/ -v
```

---

## 4. Privacy & Governance
- **No Inferred Emails:** Only verifiable public professional roles and profiles are recorded.
- **No Outbound Actions:** The system never transmits emails, fills forms, contacts targets, or initiates commercial purchases.
- **Untrusted Web Content:** Web data is treated as untrusted evidence, avoiding prompt injection or promotional distortions.
- **Neutrality:** The workspace remains completely independent, unbranded by any employer or commercial entity.
