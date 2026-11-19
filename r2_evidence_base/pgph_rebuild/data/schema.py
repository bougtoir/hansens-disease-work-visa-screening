"""Single schema definition for all rebuild evidence tables (imported by validators and analysis)."""
SEARCH_LOG = ["search_id", "jurisdiction_id", "batch", "stage", "language", "timestamp_utc", "query", "engine_or_database", "result_url", "hit_title", "relevant_excerpt", "opened_full_text", "access_status", "translation_method", "notes"]
SOURCE_REGISTER = ["source_id", "jurisdiction_id", "title", "issuing_authority", "source_tier", "doc_type", "language", "url", "archive_url", "retrieved_utc", "local_path", "bytes", "sha256", "effective_date", "status_in_force", "redistribution", "notes"]
EVIDENCE_ROWS = ["evidence_id", "jurisdiction_id", "pathway_id", "source_id", "provision_id", "exam_country_scope", "condition_code", "condition_other_text", "criterion_type", "source_wording_original", "source_wording_english", "explicit_named", "operative_effect", "authority_level", "effective_from", "effective_to", "clinical_state_rule", "safeguard_rule", "trigger_rule", "excerpt_locator", "evidence_tier", "confidence", "classification_note", "reviewer", "record_created_at", "modified_at", "supersedes", "reason_for_change", "provenance_class"]
POLICY_PATHWAYS = ["pathway_id", "jurisdiction_id", "pathway_type", "pathway_name", "medical_exam_required", "exam_location", "governing_source_ids", "dss_status", "dss_basis", "hd_named_status", "hd_basis", "consequence_values", "notes"]
CLINICAL_STATUS_RULES = ["jurisdiction_id", "pathway_id", "active_untreated", "on_treatment", "mdt_completed", "cured", "residual_sequelae", "source_ids", "notes"]
SAFEGUARDS = ["jurisdiction_id", "pathway_id", "discretion", "individual_assessment", "retest", "waiver", "appeal_review", "source_ids", "notes"]
NATIONALITY_ORIGIN_TRIGGERS = ["jurisdiction_id", "pathway_id", "trigger_type", "trigger_value", "source_ids", "notes"]
OBSERVED_DENIALS = ["case_id", "jurisdiction_id", "date", "person_status", "clinical_status", "decision", "source_type", "source_id", "notes"]
POST_TREATMENT_RISK = ["jurisdiction_id", "pathway_id", "risk_status", "mechanism", "source_ids", "notes"]
JURISDICTION_STATUS = ["jurisdiction_id", "batch", "english_pass_done", "local_pass_done", "local_languages_used", "n_queries", "n_sources_opened", "best_source_tier", "dss_status", "hd_named_status", "unresolved_reason", "hours_spent", "notes"]
CONDITION_CODES = ["HD", "TB", "HIV", "SYPH", "STI_OTHER", "HBV", "HCV", "MENTAL", "SUBST_DEP", "DRUG_TEST", "ALCOHOL", "MALARIA", "FILARIASIS", "OTHER_NAMED", "GENERIC_FITNESS"]
ENUMS = dict(
    stage=["P1_english", "P2_local", "P3_applicability", "P4_ambiguity"],
    source_tier=["S1", "S2", "S3", "S4", "S5", "S6", "S7"],
    explicit_named=["yes", "no", "ambiguous"],
    dss_status=["documented_positive", "documented_absence", "unresolved"],
    hd_named_status=["confirmed", "not_established", "cannot_determine"],
    pathway_type=["P1_work", "P2_residence_for_workers", "P3_other"],
    clinical=["mentioned_exempt", "mentioned_conditional", "mentioned_excluded", "not_addressed"],
    safeguard=["explicit_available", "explicit_unavailable", "no_statement_located", "unknown"],
    confidence=["high", "medium", "low"],
    criterion_type=["diagnostic", "behavioral", "administrative_test"],
    provenance_class=["historical_unverified", "recovered_historical", "new_collected", "independently_verified", "unresolved"],
    risk_status=["explicit_post_treatment_exclusion", "no_treatment_distinction_named_condition", "treatment_distinction_exempts_or_conditional", "not_applicable_hd_not_named", "cannot_determine"],
)
TABLES = dict(SEARCH_LOG=SEARCH_LOG, SOURCE_REGISTER=SOURCE_REGISTER, EVIDENCE_ROWS=EVIDENCE_ROWS, POLICY_PATHWAYS=POLICY_PATHWAYS,
              CLINICAL_STATUS_RULES=CLINICAL_STATUS_RULES, SAFEGUARDS=SAFEGUARDS, NATIONALITY_ORIGIN_TRIGGERS=NATIONALITY_ORIGIN_TRIGGERS,
              OBSERVED_DENIALS=OBSERVED_DENIALS, POST_TREATMENT_RISK=POST_TREATMENT_RISK, JURISDICTION_STATUS=JURISDICTION_STATUS)
