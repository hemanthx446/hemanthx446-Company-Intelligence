# Research Workflow: End-to-End Company Intelligence Investigation

**System:** hemanthx446 Company Intelligence  
**Owner:** Hemanth  
**Standard Execution Workflow:** 13-Stage Investigation Framework

---

## Overview
This workflow defines the operational protocol for conducting a corporate intelligence investigation on any target enterprise. The workflow guarantees evidence integrity, distinguishes facts from inferences, logs source citations, evaluates commercial opportunities using a transparent 100-point scoring model, and produces both narrative and Excel deliverables.

---

## Pre-Flight Preparation
1. Review [config/research_profile.md](file:///c:/Users/Lumbini_User/Desktop/hemanthx446-Company-Intelligence/config/research_profile.md) for active research focus domains (e.g., SAP S/4HANA, BTP, Cloud ALM, MES/OT integration, advisory).
2. Configure target parameters in [config/research_settings.json](file:///c:/Users/Lumbini_User/Desktop/hemanthx446-Company-Intelligence/config/research_settings.json) (Company Name, Website, Country, Industry, Allowance).
3. Create the target snapshot directory:
   `research/<company_slug>_<YYYYMMDD>/`
4. Initialize `research_log.md` from [templates/research_log_template.md](file:///c:/Users/Lumbini_User/Desktop/hemanthx446-Company-Intelligence/templates/research_log_template.md).

---

## 13-Stage Research Execution

### Stage 1: Resolve Company Identity
- **Goal:** Unambiguously identify the target legal entity, parent corporation, corporate structure, global headquarters, and primary web domains.
- **Actions:**
  - Execute web searches for official investor relations, corporate registrations, and "About Us" pages.
  - Determine legal entity status (publicly traded ticker, private equity-backed, subsidiary, or privately held).
- **Deliverables:** Record in `company_profile` with `PROF-001`+ IDs, citing primary sources.

### Stage 2: Research Company Background & Business Model
- **Goal:** Understand commercial mechanics, revenue generation streams, segment contributions, and geographical footprint.
- **Actions:**
  - Review annual reports (10-K, 20-F, annual integrated report), investor presentations, or corporate press kits.
  - Map revenue models: recurring subscriptions, capital sales, services, software licenses, maintenance fees.
- **Deliverables:** Financial scale summary, operational footprint, and monetization overview.

### Stage 3: Inventory Products, Solutions & Services
- **Goal:** Catalog commercial offerings, technology architecture, and solution capabilities.
- **Actions:**
  - Inspect product catalog pages, solution datasheets, whitepapers, and technical specifications.
  - Document core software stacks, deployment models (cloud, edge, on-premises), and hardware integrations.
- **Deliverables:** Structured `offerings` records with `OFF-001`+ IDs.

### Stage 4: Investigate Customers & Publicly Evidenced Projects
- **Goal:** Identify referenceable client implementations and evidenced customer accounts.
- **Actions:**
  - Search official case studies, joint customer press releases, customer success stories, and conference presentation decks.
  - **Rule:** Do NOT assume customer relationships from logo bars alone. Require narrative evidence of an actual deployment or signed contract.
- **Deliverables:** Populated `customers` records with project scopes and documented outcomes (`CUST-001`+).

### Stage 5: Investigate Official Partners, Vendors & Strategic Alliances
- **Goal:** Map the technology ecosystem, ERP/MES software partnerships, cloud hyperscalers, and systems integrators.
- **Actions:**
  - Search official vendor partner directories (e.g., SAP Partner Finder, Microsoft/AWS Partner Network).
  - Verify official alliance tier (Gold, Strategic, Reseller, Co-development).
- **Deliverables:** Verified `partners` records (`PART-001`+).

### Stage 6: Identify Relevant Professional Contacts
- **Goal:** Discover key decision-makers and influencers within target operational units (Enterprise IT, Manufacturing, Operations, Supply Chain, Architecture).
- **Actions:**
  - Search public LinkedIn company rosters and executive profile summaries.
  - Gather verifiable business roles, department oversight, and public profile links.
  - **Rule:** Never invent or extrapolate email addresses or phone numbers.
- **Deliverables:** `contacts` records (`CONT-001`+).

### Stage 7: Research Competitors & Market Positioning
- **Goal:** Benchmark the target against direct and indirect market competitors.
- **Actions:**
  - Review industry analyst reports (Gartner, IDC, Forrester if publicly summarized), market comparisons, and competitor claims.
  - Detail unique competitive advantages and known product/architectural vulnerabilities.
- **Deliverables:** `competitors` records (`COMP-001`+).

### Stage 8: Detect Public Business Signals
- **Goal:** Capture operational catalysts, hiring surges, leadership transitions, and capital expenditures.
- **Actions:**
  - Search recent press announcements (past 6–12 months), job board requisitions (e.g., searches for SAP BTP, MES, Cloud ALM, or integration engineers), and earnings call transcripts.
  - Document strategic significance of each signal.
- **Deliverables:** `business_signals` records (`SIG-001`+).

### Stage 9: Verify Findings & Resolve Conflicting Evidence
- **Goal:** Cross-reference findings across multiple sources and resolve ambiguities.
- **Actions:**
  - Assign strict evidence status: `Verified fact`, `Corroborated finding`, `Reasonable inference`, `Unverified claim`, `Historical information`, or `Unknown information`.
  - Remove duplicate records and demote claims lacking secondary corroboration.
  - Check whether search allowance has been exhausted. Log remaining unanswered questions in `research_gaps`.

### Stage 10: Identify & Score Commercial Opportunities
- **Goal:** Formulate actionable commercial hypotheses and compute objective 100-point scores.
- **Scoring Dimensions:**
  - **Strategic Fit (0–30):** Alignment with research profile domains.
  - **Need Evidence (0–25):** Public documentation of architectural pain or strategic need.
  - **Account/Contact Access (0–20):** Identified executive or functional champion.
  - **Commercial Relevance (0–15):** Budget scope, strategic value, or enterprise impact.
  - **Feasibility & Timing (0–10):** Actionable within 0–6 months.
- **Deliverables:** Scored `opportunities` records (`OPP-001`+) with component rationale and independent confidence rating.

### Stage 11: Create Report & Structured Research Records
- **Goal:** Generate narrative intelligence report and machine-readable data contracts.
- **Actions:**
  - Populate [templates/report_template.md](file:///c:/Users/Lumbini_User/Desktop/hemanthx446-Company-Intelligence/templates/report_template.md) into `report.md`.
  - Build `data.json` adhering to [templates/company_intelligence_schema.md](file:///c:/Users/Lumbini_User/Desktop/hemanthx446-Company-Intelligence/templates/company_intelligence_schema.md).
  - Export `sources.csv`.

### Stage 12: Generate & Validate Excel Workbook
- **Goal:** Produce management-grade `company_intelligence.xlsx` and run automated validation.
- **Actions:**
  - Run `python scripts/generate_workbook.py --input research/<target_dir>/data.json --output outputs/<target_dir>/company_intelligence.xlsx`
  - Run `python scripts/validate_research.py --input research/<target_dir>/data.json --workbook outputs/<target_dir>/company_intelligence.xlsx`
  - Confirm all 12 worksheets are populated, hyperlinks functional, and zero validation errors.

### Stage 13: Summarize Priority Opportunities, Contacts & 7-Day Action Plan
- **Goal:** Synthesize top findings for executive review.
- **Actions:**
  - Highlight the top-scoring opportunities.
  - Pair each opportunity with specific decision-maker outreach angles.
  - Provide an actionable 7-day tactical execution roadmap.
