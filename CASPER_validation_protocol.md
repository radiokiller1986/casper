# Validation Study Protocol — CASPER ADR Assessment Tool

**Title:** Agreement between an automated multi-scale ADR assessment tool (CASPER) and expert consensus for causality, severity and preventability assessment of adverse drug reactions: a reliability and agreement study.

**Tool version under test:** CASPER v0.2.0 (scoring ruleset frozen; commit/DOI to be cited).
**Protocol version / date:** v1.0 — to be finalised before data collection.
**Author:** Puneet Paliwal.

> This is a draft protocol. Numbers in the sample-size section are worked examples; finalise them with your statistician and institutional ethics committee before starting.

---

## 1. Background and rationale

No causality assessment scale is a true gold standard; published reviews repeatedly note inter-rater variability and limited reproducibility across WHO-UMC, Naranjo and related tools. An automated tool removes *rater* variability (the same input always yields the same output), but that consistency is worthless if the automated output does not *agree* with careful expert judgement. CASPER therefore cannot be presented as producing "authentic" results until its outputs have been compared, on real cases, against an expert reference standard, with agreement quantified. This study provides that evidence.

CASPER is also partly *derivation-based*: it infers temporal relationship, de-challenge and severity flags from routinely collected fields, and flags items it cannot determine. The study must therefore test not only overall agreement but whether agreement degrades when input data are incomplete (the rows CASPER flags).

## 2. Objectives

**Primary.** To quantify the agreement between CASPER and an expert-consensus reference standard for each of:
- Causality — Naranjo category (Definite / Probable / Possible / Doubtful)
- Causality — WHO-UMC category (Certain / Probable / Possible / Unlikely / Conditional / Unassessable)
- Severity — Modified Hartwig & Siegel grade (Mild / Moderate / Severe)
- Preventability — Modified Schumock & Thornton (Definitely / Probably / Not preventable)

**Secondary.**
1. Inter-rater reliability among the expert assessors themselves (the ceiling against which CASPER should be judged).
2. Agreement of CASPER on the continuous Naranjo *score* (not just category) vs experts.
3. Effect of input completeness on agreement (complete rows vs CASPER-flagged rows; stratified by "Naranjo items known/10").
4. Frequency and correctness of CASPER's data-integrity flags.
5. Time taken per case: CASPER-assisted workflow vs fully manual.

## 3. Study design

Cross-sectional reliability and agreement study using a retrospective, anonymised sample of completed ADR reports. Single-centre to begin; a multi-centre replication is recommended before any claim of generalisability. Reported per the **GRRAS** guidelines for reliability and agreement studies.

## 4. Reference standard (the comparator)

There is no external gold standard, so the reference standard is **structured expert consensus**:

- **Two independent assessors**, each trained in pharmacovigilance and blinded to (a) CASPER's output and (b) each other's scoring, independently assess every case on all four scales using the original source records and the published scale instruments.
- **Disagreements are resolved by a third senior assessor** (or a consensus meeting) to produce the final reference label for each case.
- The two independent assessors' pre-consensus scores are retained to compute *expert inter-rater reliability* (Objective S1). This number is critical: CASPER should not be expected to agree with the consensus better than the experts agree with each other.

Assessors must be blinded to CASPER throughout, and must not have been involved in building or tuning the tool.

## 5. Population, inclusion and sampling

- **Source:** consecutive ADR reports from the [AMC / department / registry], reporting period to be specified.
- **Inclusion:** reports with the mandatory fields needed for assessment (suspected drug, reaction, onset date, drug start date, outcome).
- **Exclusion:** duplicate reports; reports with no identifiable suspected drug.
- **Sampling for category spread:** simple causality categories (Definite/Certain) are rare. To avoid empty cells in the agreement matrices, use either consecutive sampling of a large enough cohort or *stratified* sampling to ensure a minimum number of cases in each major category. Pre-specify a target of **≥ 20–30 cases per major causality category** where the source data allow.
- Include a deliberate subset of **incomplete-data cases** so the completeness analysis (Objective S3) is powered.

## 6. Sample size

Sample size is driven by the precision required on Cohen's kappa.

**Worked example (causality, the primary endpoint).** Assuming an expected agreement of κ ≈ 0.75 and testing against a minimally acceptable κ₀ = 0.40 (Donner & Eliasziw approach), with two-category-collapsed prevalence near 0.5, α = 0.05 and power = 80%, the required number of cases is on the order of **~85–110**. Allowing ~15% unusable/incomplete records, **recruit ≈ 130 cases.**

**Multi-category adjustment.** WHO-UMC has up to six categories and Hartwig three; populating those agreement matrices reliably needs more cases than a two-category calculation implies. A pragmatic target of **150–200 ADR cases** is recommended for the full multi-scale study, with the minimum-per-category rule in §5. Confirm with reference tables (Bujang & Baharum 2017; Flack et al.; Donner & Eliasziw 1992) and your statistician.

State the final n and its basis explicitly in the protocol and manuscript.

## 7. Procedure

1. Extract and anonymise eligible reports into the CASPER input template (one row per suspected drug–reaction pair). Assign a study ID; strip patient identifiers.
2. **CASPER arm:** run the frozen tool version; record every output (Naranjo score + category, WHO-UMC, Hartwig level + grade, Schumock category, items-known/10, all flags). Lock these before manual assessment.
3. **Expert arm:** the two blinded assessors independently score each case from the source records; the third assessor adjudicates disagreements to the reference label.
4. Record per-case time for the CASPER-assisted vs fully manual workflow on a random subset (Objective S5).
5. Merge CASPER and reference labels by study ID for analysis. Analysts handling the merged file should be blinded to which column is which where feasible.

## 8. Statistical analysis

- **Primary:** Cohen's kappa (κ) with 95% CI for each scale, CASPER vs reference. For the ordered scales (Naranjo categories, WHO-UMC, severity) report **weighted kappa** (linear and quadratic weights) in addition to unweighted, because misclassifying "Certain" as "Unlikely" is worse than as "Probable". Report raw percent agreement and per-category sensitivity/specificity, with full confusion matrices.
- **Naranjo score:** intraclass correlation coefficient (ICC) and Bland–Altman limits of agreement for the continuous score.
- **Expert inter-rater reliability:** same statistics between the two independent assessors.
- **Completeness analysis:** repeat the primary agreement statistics stratified by complete vs CASPER-flagged rows, and across "Naranjo items known/10" bands; test whether agreement differs.
- **Integrity flags:** report sensitivity/specificity of CASPER's data-integrity flags against manual review.
- **Interpretation (Landis & Koch):** <0 poor; 0.01–0.20 slight; 0.21–0.40 fair; 0.41–0.60 moderate; 0.61–0.80 substantial; 0.81–1.00 almost perfect.
- **Pre-specified acceptance criterion:** CASPER is considered acceptable for research screening on a given scale if the **lower bound of the 95% CI for weighted κ ≥ 0.60 (substantial)**, and if its agreement with the reference is not meaningfully lower than the experts' agreement with each other. Scales that fail this should be reported as "not recommended for automated use" rather than quietly published.

## 9. Bias control

Blinding of assessors to CASPER and to each other; locking CASPER output before manual scoring; assessors independent of tool development; pre-registration of the analysis plan and the acceptance threshold; reporting all four scales regardless of result (no selective reporting of the scale that performed best).

## 10. Software reproducibility

The exact tool version (frozen scoring ruleset, version string / Zenodo DOI) is recorded. The input dataset (anonymised) and the analysis code are archived so the agreement statistics can be reproduced. Any change to the scoring logic invalidates the validation and requires re-testing.

## 11. Ethics and data governance

Institutional Ethics Committee / IRB approval before data access. Use of retrospective anonymised ADR data; no patient identifiers leave the source system; processing is local (CASPER runs offline in the browser, so case data are not transmitted). A waiver of individual consent is typically appropriate for anonymised retrospective pharmacovigilance data but is for the committee to decide. Comply with the PvPI/national data-handling requirements.

## 12. Limitations (to state up front)

The reference standard is expert consensus, not biological truth — the study measures agreement with expert judgement, not correctness of causality per se. Single-centre results may not generalise. Rare categories may remain under-represented. CASPER's derivation means part of what is being validated is the *imputation* logic, not only the scales; the completeness analysis addresses but does not eliminate this.

## 13. Timeline (indicative)

Ethics approval → assessor training and calibration on pilot cases → data extraction and anonymisation → parallel CASPER and expert scoring → adjudication → analysis → reporting (GRRAS-compliant manuscript). Build in a small pilot (~15–20 cases) to calibrate assessors and finalise the data dictionary before the main run.

## 14. Key references to cite

- Manjhi PK, Singh MP, Kumar M. Causality, Severity, Preventability and Predictability Assessment Scales for ADRs: A Review. *Cureus* 2024;16(5):e59975.
- Naranjo CA et al. A method for estimating the probability of adverse drug reactions. *Clin Pharmacol Ther* 1981;30:239–45.
- WHO-UMC system for standardised case causality assessment (Uppsala Monitoring Centre).
- Hartwig SC, Siegel J, Schneider PJ. Severity assessment scale. *Am J Hosp Pharm* 1992.
- Schumock GT, Thornton JP. Focusing on the preventability of ADRs. *Hosp Pharm* 1992.
- Landis JR, Koch GG. The measurement of observer agreement for categorical data. *Biometrics* 1977;33:159–74.
- Donner A, Eliasziw M. A goodness-of-fit approach to inference procedures for the kappa statistic. *Stat Med* 1992.
- Kottner J et al. Guidelines for Reporting Reliability and Agreement Studies (GRRAS). *J Clin Epidemiol* 2011.
- Bujang MA, Baharum N. Guidelines of the minimum sample size requirements for Cohen's kappa. *Epidemiol Biostat Public Health* 2017.

---

## Addendum (v0.2) — efficiency, usability, and the equivalence framing

Reviewers of v0.2 asked for more than agreement statistics. The validation study should therefore report three co-primary/secondary endpoints:

1. **Agreement / equivalence (primary).** CASPER vs blinded expert consensus, per scale, by weighted Cohen's κ with 95% CI, against the pre-specified acceptance threshold (lower bound ≥ 0.60). Frame it as an *equivalence* question — CASPER should agree with experts about as well as experts agree with each other (report the experts' inter-rater κ as the ceiling). Now includes **ICH E2A seriousness** (Serious/Non-serious) as a fifth endpoint.
2. **Efficiency / throughput (secondary).** Record time-to-completion per case, CASPER-assisted vs fully manual, on a random subset; compare with a paired test (paired t-test or Wilcoxon signed-rank). Hypothesis: CASPER significantly reduces time and cognitive load.
3. **Usability (secondary).** Administer the **System Usability Scale (SUS)** to ≥ 30 prospective users (PV officers, clinical-pharmacology residents, physicians). Target mean SUS > 72.5 (Good/Excellent). Report mean ± SD and the adjective rating.

**Transparency & integrity for reproducibility.** Each CASPER output row now carries a per-rule score breakdown, an assessment timestamp, and a SHA-256 hash over (case + drug + reaction + outputs + timestamp + version). Archive the frozen tool version (DOI), the anonymised dataset, and the analysis code so the κ values can be independently reproduced. Note explicitly in the limitations that CASPER is **not** a 21 CFR Part 11 / GAMP 5 validated system; the hash provides tamper-evidence, not regulated electronic-records compliance.
