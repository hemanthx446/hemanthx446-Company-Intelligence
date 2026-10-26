# Company Intelligence Data Schema Specification

**System:** hemanthx446 Company Intelligence  
**File Format:** JSON (`data.json`) and accompanying CSVs (`sources.csv`)  
**Purpose:** Defines the standardized data contract for structured corporate research records. Used by `scripts/generate_workbook.py` to populate all 12 worksheets of `company_intelligence.xlsx` and by `scripts/validate_research.py` to enforce data integrity.

---

## 1. Top-Level Structure

A valid `data.json` file must contain the following top-level keys:

```json
{
  "metadata": { ... },
  "executive_summary": { ... },
  "company_profile": [ ... ],
  "offerings": [ ... ],
  "customers": [ ... ],
  "partners": [ ... ],
  "contacts": [ ... ],
  "competitors": [ ... ],
  "business_signals": [ ... ],
  "opportunities": [ ... ],
  "outreach_plan": [ ... ],
  "sources": [ ... ],
  "research_gaps": [ ... ]
}
```

---

## 2. Entity Schemas & Enums

### A. Metadata (`metadata`)
- `project_name` (string): Fixed to `"hemanthx446 Company Intelligence"`
- `owner` (string): Fixed to `"Hemanth"`
- `target_company_name` (string, required): Official legal or commercial name.
- `target_website` (string, required, valid URL format).
- `target_country` (string, required): Headquarters country / region.
- `target_industry` (string, required).
- `research_date` (string, YYYY-MM-DD format).
- `overall_confidence` (string, Enum: `High`, `Medium`, `Low`).

### B. Standard Allowed Enums
- **Evidence Status (`evidence_status`):**
  - `Verified fact`
  - `Corroborated finding`
  - `Reasonable inference`
  - `Unverified claim`
  - `Historical information`
  - `Unknown information`
- **Confidence Level (`confidence`):**
  - `High`
  - `Medium`
  - `Low`

### C. Record ID Conventions
All entities must use stable, uppercase, prefixed record identifiers:
- `sources`: `SRC-001`, `SRC-002`, ...
- `company_profile`: `PROF-001`, `PROF-002`, ...
- `offerings`: `OFF-001`, `OFF-002`, ...
- `customers`: `CUST-001`, `CUST-002`, ...
- `partners`: `PART-001`, `PART-002`, ...
- `contacts`: `CONT-001`, `CONT-002`, ...
- `competitors`: `COMP-001`, `COMP-002`, ...
- `business_signals`: `SIG-001`, `SIG-002`, ...
- `opportunities`: `OPP-001`, `OPP-002`, ...
- `outreach_plan`: `OUT-001`, `OUT-002`, ...
- `research_gaps`: `GAP-001`, `GAP-002`, ...

### D. Detailed Key Definitions

#### 1. Sources (`sources`)
```json
{
  "id": "SRC-001",
  "title": "Corporate Annual 10-K Filing",
  "url": "https://example.com/sec/filing",
  "publication_date": "2026-03-31",
  "access_date": "2026-10-26",
  "evidence_category": "Verified fact",
  "confidence": "High",
  "notes": "Primary regulatory disclosure regarding segment revenues."
}
```

#### 2. Company Profile (`company_profile`)
```json
{
  "id": "PROF-001",
  "attribute": "Legal Entity Name",
  "value": "Target Industrial Technologies Corp.",
  "category": "Corporate Identity",
  "evidence_status": "Verified fact",
  "sources": ["SRC-001"],
  "notes": "State incorporation records."
}
```

#### 3. Offerings (`offerings`)
```json
{
  "id": "OFF-001",
  "name": "Cloud MES Suite",
  "category": "Software Solution",
  "tech_stack": "Kubernetes, SAP NetWeaver connector, MQTT broker",
  "value_proposition": "Real-time production scheduling and shopfloor telemetry",
  "target_market": "Tier 1 Discrete Manufacturers",
  "evidence_status": "Verified fact",
  "sources": ["SRC-002"]
}
```

#### 4. Customers (`customers`)
```json
{
  "id": "CUST-001",
  "customer_name": "Apex Automotive Group",
  "industry": "Automotive Tier 1",
  "project_scope": "Full shopfloor MES deployment across 4 assembly plants",
  "technologies_used": "SAP S/4HANA, Custom MES Bridge",
  "documented_outcomes": "14% reduction in cycle downtime",
  "evidence_status": "Verified fact",
  "confidence": "High",
  "sources": ["SRC-003"]
}
```

#### 5. Partners (`partners`)
```json
{
  "id": "PART-001",
  "partner_name": "Cloud Platform Provider",
  "partner_type": "Strategic Hyperscaler",
  "strategic_focus": "SaaS hosting and enterprise data lake co-selling",
  "verification_method": "Official Cloud Partner Directory",
  "evidence_status": "Verified fact",
  "sources": ["SRC-004"]
}
```

#### 6. Key Contacts (`contacts`)
```json
{
  "id": "CONT-001",
  "name": "Alex Mercer",
  "title": "VP of Enterprise Architecture & Applications",
  "functional_area": "IT & Enterprise Systems",
  "decision_making_role": "Executive Sponsor / Budget Approver",
  "profile_url": "https://www.linkedin.com/in/example-alex-mercer",
  "sources": ["SRC-005"]
}
```

#### 7. Competitors (`competitors`)
```json
{
  "id": "COMP-001",
  "competitor_name": "Global Factory Systems Inc.",
  "market_overlap": "Direct competition in MES and shopfloor automation",
  "strengths": "Established legacy presence in North America",
  "vulnerabilities": "Monolithic architecture, slow cloud transition",
  "sources": ["SRC-006"]
}
```

#### 8. Business Signals (`business_signals`)
```json
{
  "id": "SIG-001",
  "signal_category": "Executive Appointment",
  "event_description": "New Chief Digital Officer announced to spearhead Clean Core transformation",
  "date_detected": "2026-08-15",
  "commercial_impact": "Accelerated timeline for enterprise application modernizations",
  "sources": ["SRC-007"]
}
```

#### 9. Opportunities (`opportunities`)
Scoring dimensions:
- `strategic_fit_score` (0–30)
- `need_evidence_score` (0–25)
- `contact_access_score` (0–20)
- `commercial_relevance_score` (0–15)
- `feasibility_timing_score` (0–10)
- `total_score` (sum of the 5, 0–100)
- `confidence` (`High`, `Medium`, `Low`)
- `rationale` (detailed analytical explanation)
- `related_contacts` (list of contact IDs, e.g. `["CONT-001"]`)
- `related_signals` (list of signal IDs, e.g. `["SIG-001"]`)
- `sources` (list of source IDs)

#### 10. Outreach Plan (`outreach_plan`)
```json
{
  "id": "OUT-001",
  "target_contact_id": "CONT-001",
  "opportunity_id": "OPP-001",
  "value_angle": "SAP BTP extension architecture for zero-downtime MES integration",
  "suggested_asset": "Clean Core Shopfloor Architecture Blueprint",
  "channel": "LinkedIn InMail / Corporate Message",
  "sequence_stage": "Stage 1: Thought Leadership Share",
  "status": "Planned"
}
```

#### 11. Research Gaps (`research_gaps`)
```json
{
  "id": "GAP-001",
  "dimension": "Shopfloor Telemetry Protocol",
  "description": "Specific industrial protocol (OPC-UA vs MQTT) standard in European plants",
  "cause": "Proprietary plant telemetry specifications not published externally",
  "next_action": "Monitor upcoming automation symposium presentations and tech job postings"
}
```
