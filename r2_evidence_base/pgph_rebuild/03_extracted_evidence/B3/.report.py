"""Regenerate B3's Markdown report from the batch tables and retrieval records."""

import csv
import json
from collections import Counter
from pathlib import Path

output = Path(__file__).resolve().parent
batch = output.name
logs_dir = output.parents[1] / "01_raw_search_logs" / batch


def table(name):
    with (output / (name + ".csv")).open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


searches = table("SEARCH_LOG")
statuses = table("JURISDICTION_STATUS")
sources = table("SOURCE_REGISTER")
pathways = table("POLICY_PATHWAYS")
evidence = table("EVIDENCE_ROWS")
cases = table("OBSERVED_DENIALS")
post = table("POST_TREATMENT_RISK")
stages = Counter(row["stage"] for row in searches)
reviewed = [row for row in sources if row["status_in_force"] != "not_used_for_adjudication"]
fetches = {p.stem[6:]:json.loads(p.read_text()) for p in sorted(logs_dir.glob("fetch_*.json"))}

critical = {
    "ARE_dha_decree_pdf": "Full UAE Cabinet Resolution 75/2022 annex. DHA returned a portal shell; official summaries were used for the wording they contain.",
    "ARE_wafid_current": "Wafid’s updated regulations and current administrative and technical appendices. The guideline requests returned 403; the live Arabic FAQ and 2021 version remained available.",
    "AZE_migration_health": "Azerbaijan's full Cabinet list referenced by the migration procedure. The attempted legal-database result did not yield a verified instrument.",
    "DZA_work_rendered": "Algeria's full current labour procedure and separate medical forms. The rendered labour procedure was partial; the residence checklist was available.",
    "EGY_permits_current": "Egypt's current work-permit packet under Decision 279/2025. The official gazette and legacy instructions were retrievable, but the contemporary checklist was unverified.",
    "IRQ_recruitment_rules": "Official gazette PDF of Iraq’s 2026 recruitment instructions. An unofficial transcription was available; the residence law was independently located on the ministry site.",
    "KWT_official_gazette": "Kuwait Ministerial Decision 139/2026. Official text was unavailable; a legal report served as a lead.",
    "KWT_worker_exam": "Kuwait's government worker examination page returned 403. A national screening panel remains unverified.",
    "LBN_labour_faq": "Lebanon’s labour ministry site returned 403. The General Security domestic-worker instructions did not specify the testing panel.",
    "LBY_ncdc": "Libya’s disease-control centre material did not yield a verified analyte list and eligibility schedule for the worker certificate.",
    "MAR_employment_guide": "Morocco’s full labour ministry procedure PDF did not download. The rendered copy omitted detailed checklists; contemporary residence and contract procedures remain unresolved.",
    "PSE_work_regulation": "Palestinian Decision 45/2004 and its implementing checklist remain unverified. Legal portal requests returned an index.",
    "QAT_moph_rendered": "Qatar’s Medical Commission service page was opened. The domestic panel and clinical fitness rules were missing from the official description.",
    "SDN_non_sudanese_law": "Official text and current implementing instructions under Sudan’s non-Sudanese employment law remain unverified.",
    "SYR_work_rules": "Syrian implementing work rules could not be opened through the ILO link; continuing use of the one-year consular checklist remains unverified.",
    "TUN_labour_fr": "Complete current employment-endorsement document lists. The ministry index and the model contract were accessible; the underlying checklists were not verified.",
    "TUR_residence_pdf": "Türkiye’s residence-document download timed out. The list associated with Law 6458’s public-health restriction remains unverified.",
    "YEM_aden_ministry": "Yemen’s current permitting instructions. The Aden ministry page returned 503; statutory wording was retrieved separately.",
}


def ids(rows, field, value):
    return ", ".join(row["jurisdiction_id"] for row in rows if row[field] == value) or "None"


def names(rows, field, values):
    return [
        f"- `{value}` ({sum(row[field] == value for row in rows)}) — {ids(rows, field, value)}"
        for value in values
    ]


paths_absent = ", ".join(row["pathway_id"] for row in pathways if row["dss_status"] == "documented_absence")
case_ids = ", ".join(row["case_id"] for row in cases)
regional = [
    row for row in pathways
    if row["pathway_name"] == "Wafid pre-departure examination for covered applicants"
]
confirmed = [row for row in pathways if row["hd_named_status"] == "confirmed"]
lines = [
    "# B3 evidence report",
    "",
    "Northern Africa and Western Asia. The batch covers the assigned jurisdictions in the frozen review frame.",
    "",
    "## Search and source inventory",
    "",
    f"Executed {len(searches)} logged searches across {len(statuses)} jurisdictions. "
    + ", ".join(f"{n} {stage.split('_', 1)[0]}" for stage, n in sorted(stages.items())) + ". "
    + "All jurisdictions have an English pass and at least one local-language pass. "
    + f"The register contains {len(sources)} saved retrievals, of which {len(reviewed)} were reviewed as policy evidence, case evidence or contextual leads. "
    + "The other records preserve supplementary, partial, off-topic or unsuccessful download responses. "
    + f"The extraction contains {len(pathways)} pathways and {len(evidence)} evidence rows.",
    "",
    "The raw query responses and `fetch_*.json` document the search and acquisition trail. "
    "The source register lists the opened documents and records URL, UTC retrieval time, bytes, SHA-256 and redistribution status. "
    "Selected government pages were captured as rendered text. Partial renderings and portal shells are identified in the register. "
    "Local queries and translations were model-assisted and the original wording was preserved. "
    "Exact active time per jurisdiction could not be recovered from the parallel search and extraction workflow. Selective follow-ups are recorded in the jurisdiction notes.",
    "",
    "## Findings by jurisdiction",
    "",
    "### Disease-specific screening",
    "",
    *names(statuses, "dss_status", ["documented_positive","documented_absence","unresolved"]),
    "",
    "### Hansen’s disease explicitly named",
    "",
    *names(statuses, "hd_named_status", ["confirmed","not_established","cannot_determine"]),
    "",
    f"The confirmed findings represent {len(confirmed) - len(regional) + 1} policy mechanisms. "
    f"Wafid is one shared pre-departure mechanism across {len(regional)} destinations. "
    "UAE government descriptions separately name leprosy in the destination residence programme. "
    "They describe unfitness on a positive result at first issue and renewal. The 2021 Wafid regulations were read as version history. "
    "The Gulf Health Council reports a subsequent update and its official Arabic FAQ retains leprosy in the unfitness list.",
    "",
    f"Pathway-level documented absences are {paths_absent}. Each is limited to its named procedure. "
    "Georgia, Lebanon, Tunisia and Türkiye retain separate unresolved medical or public-health pathways. "
    "The Japan B/1 checklist allows a physician’s fitness certificate without specified tests; other Israeli posts specify named tests. "
    "The Cyprus renewal checklist has no medical item and its initial-stage checklists specify named tests.",
    "",
    "## Key adjudications",
    "",
    "- Armenia's official residence law cross-references a disease list. The relevant official list was read separately. Azerbaijan's permit and residence pages explicitly name conditions; the full referenced Cabinet list remains unresolved.",
    "- Algeria's residence form requires general medicine, phtisiology and serology certificates. These labels leave the clinical and serological criteria unspecified. Egypt’s undated Alexandria work-permit procedure specifies an AIDS-free certificate; continuity under Decision 279/2025 remains unresolved.",
    "- Iraq documents an AIDS-related federal entry condition and a separate Kurdistan Region scheme with a country-dependent panel. Jordan’s Ministry of Health certificate workflow specifies named laboratory tests and conditional TB investigation.",
    "- Bahrain and Oman publish destination examination material distinct from Wafid. Oman’s follow-up and conditional fitness provisions concern named conditions other than HD. Qatar and Saudi Arabia provide official evidence of domestic examinations; their full domestic panels remain unresolved.",
    "- Syria’s published consular checklist specifies HIV testing for one-year residence requiring labour authority approval. The positive classification is limited to that documented route. Current domestic implementation and the work-permit requirements remain uncertain.",
    "",
    "## Treatment status, removals and safeguards",
    "",
    f"{len(post)} pathway records describe HD rules without an explicit distinction by treatment status. "
    "No examined policy text established an exclusion after completed MDT, clinical cure or for residual sequelae. "
    "The Wafid FAQ leaves the implications of the unqualified leprosy listing unresolved. "
    "The UAE summaries give no definition of a positive result as applied to a previously treated applicant. "
    "The 2021 Wafid complaint process concerned the examination service. Current applicant appeal or waiver against HD-related unfitness was unverified.",
    "",
    f"The case register records {case_ids}. Lambert and Walker, `Leprosy Review` 86(1), 108–111 (2015), reports the two removals from Kuwait. "
    "Both workers had received treatment or had medication dispensed before removal and completed MDT in Ethiopia later. "
    "One had a bacterial index of zero. Clinical non-infectiousness at removal and a completed-MDT exclusion were unverified.",
    "",
    "## Inaccessible or incomplete critical sources",
    "",
    *[f"- {detail} [Acquisition record]({fetches[key]['url']})" for key, detail in critical.items()],
    "",
    "## Reproduction",
    "",
    "From the `pgph_rebuild` directory run the batch assembly, report generator and validator.",
    "",
    "```sh",
    "python3 03_extracted_evidence/B3/.assemble.py",
    "python3 03_extracted_evidence/B3/.report.py",
    "PYTHONDONTWRITEBYTECODE=1 python3 scripts/validate_batch.py B3",
    "```",
    "",
    "The validator checks structure and snapshot integrity. Independent verification of the substantive classifications is a separate review step.",
    "",
]

(output / "BATCH_REPORT.md").write_text("\n".join(lines), encoding="utf-8")
