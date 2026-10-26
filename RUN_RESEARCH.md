# Operator Guide: Running Company Intelligence Investigations

**System:** hemanthx446 Company Intelligence  
**Owner:** Hemanth  
**Target Audience:** Operators and automated agents executing corporate investigations.

---

## 1. Quick-Start Execution Overview

An intelligence investigation consists of four operational steps:
```
1. Configure Target  -->  2. Conduct Research & Log  -->  3. Generate Workbook  -->  4. Validate & Review
   (settings.json)        (data.json + report.md)         (generate_workbook.py)    (validate_research.py)
```

---

## 2. Step-by-Step Execution Guide

### Step 1: Configure Target & Interests
1. Open [config/research_profile.md](file:///c:/Users/Lumbini_User/Desktop/hemanthx446-Company-Intelligence/config/research_profile.md) to inspect or refine your analytical focus (e.g. SAP S/4HANA, BTP, MES, Cloud ALM, Advisory).
2. Edit [config/research_settings.json](file:///c:/Users/Lumbini_User/Desktop/hemanthx446-Company-Intelligence/config/research_settings.json) with target company parameters:
   ```json
   {
     "target_company": {
       "name": "Target Company Name",
       "official_website": "https://www.targetcompany.com",
       "country_or_region": "United States",
       "industry": "Industrial Discrete Manufacturing",
       "research_objective": "Comprehensive intelligence analysis, technology footprint, and commercial opportunity scoring."
     },
     "execution_parameters": {
       "research_date": "2026-10-26",
       "research_depth": "standard"
     }
   }
   ```

### Step 2: Launch the Investigation (Agent Prompt)
Provide the prompt below to the Antigravity agent to initiate the live 13-stage investigation workflow:

```text
Please execute a comprehensive company intelligence investigation following the instructions in AGENTS.md and workflows/research_company.md.

Target Company: [INSERT COMPANY NAME]
Website: [INSERT WEBSITE URL OR "Auto-discover"]
Country: [INSERT COUNTRY / REGION]
Industry: [INSERT INDUSTRY OR "Auto-discover"]

Requirements:
1. Create a snapshot directory: research/[company_slug]_[YYYYMMDD]/
2. Adhere strictly to the 13-stage workflow in workflows/research_company.md.
3. Classify all evidence into the 6 standard categories and cite all sources with SRC-001+ IDs.
4. Evaluate opportunities using the 100-point scoring model against config/research_profile.md.
5. Generate report.md, research_log.md, sources.csv, and data.json adhering to templates/company_intelligence_schema.md.
6. Execute scripts/generate_workbook.py to produce outputs/[company_slug]_[YYYYMMDD]/company_intelligence.xlsx.
7. Validate all outputs using scripts/validate_research.py.
```

### Step 3: Run the Excel Workbook Generator
To generate or refresh the Excel workbook manually:
```bash
.venv\Scripts\python scripts/generate_workbook.py --input research/<target_slug>_<date>/data.json --output outputs/<target_slug>_<date>/company_intelligence.xlsx
```

### Step 4: Run the Research Validator
To verify data completeness, source link integrity, and Excel structure:
```bash
.venv\Scripts\python scripts/validate_research.py --input research/<target_slug>_<date>/data.json --workbook outputs/<target_slug>_<date>/company_intelligence.xlsx
```

### Step 5: Run Automated System Tests
To run the full regression test suite:
```bash
.venv\Scripts\pytest tests/ -v
```

---

## 3. Interpreting Deliverables

| Deliverable | Location | Description & Business Use |
| :--- | :--- | :--- |
| **Executive Dossier** | `research/<target>/report.md` | Complete 15-section narrative analysis for executive reading and strategic decision-making. |
| **Excel Workbook** | `outputs/<target>/company_intelligence.xlsx` | 12-sheet structured model with frozen headers, auto-fit columns, clickable source links, and scoring matrix. |
| **Audit Log** | `research/<target>/research_log.md` | Search queries executed, pages visited, hypotheses tested, and unresolved questions. |
| **Source Register** | `research/<target>/sources.csv` | Tabular registry of all primary and secondary evidence citations. |
| **Structured Data** | `research/<target>/data.json` | JSON data contract enabling programmatic downstream integrations or automated updates. |

---

## 4. Tips for Maximizing Research Quality

1. **Target Specific Subdomains:**
   - Use Google search operators: `site:target.com/investors`, `site:target.com/press`, `site:target.com/case-studies`.
2. **Detect Stealth Signals:**
   - Search public job board requisitions (`site:lever.co`, `site:greenhouse.io`, or corporate career portals) for specific software keywords like `SAP S/4HANA`, `BTP`, `Cloud ALM`, `MES`, or `SCADA`. Job postings disclose internal technology architectures months before public press releases.
3. **Verify Partner Tiers:**
   - Query official vendor directories (e.g., SAP Partner Finder, Microsoft Solution Provider Directory, AWS Partner Network) rather than relying on self-declared website footer logos.
4. **Distinguish Facts from Marketing Claims:**
   - Require joint customer quotes, signed SEC filings, or press releases for verified customer relationships. If an organization merely speaks at a vendor conference, classify the relationship as a `Reasonable inference` or `Unverified claim`.
