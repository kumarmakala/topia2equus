# UAT-02 — Configuration-driven business rules, approval status, import readiness

Task: OCNORA T-1387 (job-947f04a88a11). Sources: `clarifications/UAT_MASTER_COMMAND.md`
(§5, §6, §13, architecture requirement), `clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md`
(D4 table; §4 "gate behind explicit approval status"). Status vocabulary as in the assessment.
Nothing in this document is SIGNED OFF; Equus UAT execution is **NOT EXECUTED** (no Equus access
from this environment).

## 1. What changed

| File | Change | Reason |
|---|---|---|
| `lambda/src/business_rules.json` (new) | Declares the six D4 rules with `id`, `Q`, `title`, `status`, `approved_for_import: false`, `owner`, `parameters`; file `version` 1.0.0 | D4: unapproved rules must not reach import-ready output |
| `lambda/src/reference_data/*.json` (new) | `resource_usernames` (Q2), `cost_centres` (Q15), `country_state` (Q12/Q3): version/approval metadata, **empty** `values` | Assessment §4: reference lists from a versioned file; absence = exception |
| `lambda/src/mapping_config.py:262-721` | `ConfigurationError`, strict loaders `load_business_rules` (`:539`) and `load_reference_data` (`:662`), parameter schema per rule (`RULE_PARAMETERS`, `:408`), `rule_enabled()` helper (`:712`); value maps, cancellation constants, host-job-level override, `DEFAULT_DATES_FROM_START` and `RULE_TARGETS` removed from code constants | Architecture requirement: business configuration versioned, rejected when invalid |
| `lambda/src/mapping_config.py:43` | Brown fill `FFB45F06` = `none (Q8: no meaning)` | Q8 CONFIRMED BY JAKE |
| `lambda/src/transform.py:464` `apply_assignment_rules`, `:494` `default_child_dates` | Q34 / Q40 / Q22-Q23 rules applied only when approved; otherwise rows counted | D4 rows 1-3 |
| `lambda/src/transform.py:386` `_convert` | Q35 value maps gated; Q12/Q3 country conversion performed but recorded as provisional while unapproved; resource lookup from the approved resource list only | D4 rows 4-5 |
| `lambda/src/transform.py:251` `build_column_plan` | Cancellation Reason and Host Job Level columns are `held` (flag `HELD`) while their rule is unapproved | D4 rows 2, 6 |
| `lambda/src/transform.py:710` `compute_readiness` | `NOT_READY` / `READY_FOR_UAT_EXECUTION` with reasons (category, item, Q, owner, detail) | Master command §5 (STRICT_PENDING concern), §10 |
| `lambda/src/transform.py:882,904` | New report sheets **Rules** and **Readiness** | AC-4 |
| `lambda/src/handler.py:159-205` | `*_EQUUS_DRAFT_*` name while `NOT_READY`; `readiness`, `readiness-reasons`, `business-rules-version` in S3 metadata and log summary; duplicate check covers both names; code version hashes the JSON configuration | AC-4; idempotency kept |
| `tests/test_migration.py` | `test_conversions_and_rules` assertions now expect the gated outcome (Q35, Q34, Q40, Q32, each cited); `test_governance_flags` brown assertion (Q8) | Approved-change rule: assertions updated with Q citation, nothing removed |
| `tests/test_business_rules.py` (new) | AC-1 … AC-4 tests (listed in §4) | |

## 2. Behaviour per rule (shipped configuration)

| Rule id | Q | Status | Owner | Not approved (shipped) | Approved (test copy) |
|---|---|---|---|---|---|
| `international_relo_derivation` | Q40 | open | Schneider | `INTERNATIONAL_RELO` not derived | Yes/No from FROM_COUNTRY ≠ TO_COUNTRY |
| `cancellation_reason_sets_phase` | Q34 | open | Jake + Schneider | column HELD, `PHASE` only from a mapped column | non-empty reason ⇒ `PHASE = Cancelled` |
| `default_effective_dates_from_start` | Q22/Q23 | open | Jake (Q22) / Alka, John (Q23) | no defaulted salary / family-size effective date, address `BEGIN_DATE`, contact `START_DATE` | defaults to `ASSIGNMENT!START_DATE` |
| `value_maps_employee_type_and_status` | Q35 | open | Jake | values sent as in the source | `employee→Employee`, `user→User`, `employed→Active` |
| `host_job_level_to_salary_grade` | Q32 | open | Jake / Alka | column HELD (not mapped by its row-2 `TO_TITLE` either) | `TO_SALARY_GRADE` |
| `country_name_to_iso3` | Q12/Q3 | partially_clarified | Jake (Equus country/state list) | conversion performed; every converted value listed as `provisional conversion` and counted against readiness | conversion without provisional marking |

Rollback / previous behaviour: set `approved_for_import: true` for the rule (add an
`approval_reference`). This changes the handler's code version, so new output keys are produced.

## 3. Assumptions (recorded instead of asking)

1. **A gated rule counts against readiness only when it affected at least one row** of the run;
   a gated rule with no affected rows is still listed in the Rules sheet ("no (gated)", 0 rows).
2. **Held columns are not migrated anywhere.** Mapping Host Job Level by its row-2 key
   (`TO_TITLE`) would apply a field the source sheet itself marks as wrong, so the column is held
   and reported (Q32 owner) rather than re-pointed.
3. **Reference lists are used only when `approved_for_import` is true and `values` is non-empty.**
   Equus-supplied lists still need the owner to set the flag; until then every value of the
   covered headers is a readiness exception. An absent file is treated like an empty one (not a
   configuration error); an unknown file in `reference_data/` is a configuration error.
   `applies_to` patterns are business configuration in `mapping_config.REFERENCE_LISTS`.
   State-value validation against `country_state` is not implemented (only countries are
   checked); recorded as open for when Q12's list arrives.
4. **Dropdown headers without an Equus list (Q14)** count against readiness ("unknown value =
   exception"). `lambda/src/equus_lists.json` keeps its location and format.
5. **Readiness owners for non-rule items** (`mapping_config.READINESS_ITEMS`) are taken from the
   assessment and OPEN_QUESTIONS.md; they are routing hints, not approvals.
6. **Configuration headers missing from the Equus template** are reported (warning + readiness
   reason) rather than raised, because the template is a read-only input that the rule owner
   may still change; malformed configuration files are always raised.
7. **`approved_for_import` can be true for any status**: the task's status vocabulary has no
   `signed_off`, and the flag is the owner's explicit switch. `signed_off` as a status is rejected.
8. **Identifier fallback (D1) is unchanged** (UAT-03 scope); every fallback row is now a readiness
   reason (Q1).

## 4. Test evidence

Tests (synthetic fixture `tests/fixtures/topia_sample.xlsx` only):

- AC-1: `tests/test_business_rules.py::test_malformed_business_rules_fail_fast_naming_path`
  (16 parametrised malformations incl. unknown status, duplicate id, missing owner),
  `::test_duplicate_json_key_invalid_json_and_missing_file_rejected`,
  `::test_malformed_reference_list_fails_fast`, `::test_unsupported_reference_list_file_rejected`,
  `::test_valid_configuration_loads_and_version_is_reported`.
- AC-2: `::test_default_output_has_no_unapproved_rule_values`,
  `::test_default_readiness_not_ready_and_lists_every_gated_and_pending_item`,
  `::test_shipped_configuration_declares_every_d4_rule_unapproved`,
  `tests/test_migration.py::test_conversions_and_rules`.
- AC-3: `::test_enable_international_relo_restores_derivation`,
  `::test_enable_cancellation_phase_restores_cancelled`,
  `::test_enable_default_dates_restores_start_date_defaults`,
  `::test_enable_value_maps_restores_translation`,
  `::test_enable_host_job_level_restores_salary_grade`,
  `::test_enable_country_iso3_removes_provisional_marking`,
  `::test_rule_functions_gated_and_enabled`.
- AC-4: `::test_rules_and_readiness_sheets`, `::test_handler_names_draft_and_records_readiness`,
  `::test_handler_uses_import_name_only_when_ready`.

**Execution status: NOT RUN by the implementer.** The development session that produced this
change could not execute Python (every interpreter invocation was refused by the sandbox), so no
test, `py_compile` or JSON-parse result is recorded here. The runner's gate executes the full
suite independently. No real-data reconciliation run was performed for the same reason.

## 5. Readiness verdict

With the shipped configuration every run is **NOT READY FOR UAT** (`NOT_READY`): all six D4 rules
are unapproved, reference lists are empty (Q2, Q12, Q15), dropdown lists are missing (Q14), and
yellow columns hold data. Gates G6–G8 (Equus import, validation, duplicate/retry in Equus) are
NOT EXECUTED; G11–G12 not obtained.
