#!/usr/bin/env bash
# Regenerates all counts, the legal-consequence audit, the WHO burden comparison and Figures 1-6.
set -euo pipefail
cd "$(dirname "$0")"
python3 pgph_rebuild/scripts/gate4_merge_and_count.py
python3 pgph_rebuild/scripts/gate4b_exclusion_consequences.py
python3 pgph_rebuild/scripts/gate4d_legal_audit.py > /dev/null
python3 pgph_rebuild/scripts/gate4c_post_treatment.py
python3 pgph_rebuild/scripts/gate5_numbers.py > /dev/null
python3 pgph_rebuild/scripts/gate6_figures.py
python3 pgph_rebuild/scripts/test_legal_rectification_pipeline.py | tail -1
sha256sum -c EXPECTED_OUTPUTS.sha256
