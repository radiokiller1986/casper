# PvPI ADR Reporting Form → CASPER mapping

How fields on the **IPC / PvPI Suspected ADR Reporting Form (v1.4)** feed the CASPER scoring engine, and which assessments the form alone can and cannot fully support.

## What the form *can* score on its own

The PvPI form was designed for **WHO-UMC causality**, so it natively carries the four WHO-UMC criteria. These fields map directly:

| PvPI form field | CASPER column | Used by |
|---|---|---|
| Patient initials | `patient_id` | reporting |
| Age / date of birth (2) | `age_years` | reporting |
| Gender (3) | `sex` | reporting |
| Weight (4) | `weight_kg` | Schumock (dose check) |
| Reaction start date (5) | `adr_onset_date` | temporal |
| Reaction stop date (6) | — (with therapy dates → `c_temporal`) | causality |
| Reaction management (7) | → `s_antidote_required` | severity |
| Suspected medication: Name / Dose / Route / Frequency / Therapy start–stop / Indication (8) | `suspected_drug`, `dose`, `route`, `frequency`, `therapy_start_date`, `therapy_stop_date`, `indication` | all |
| Action taken after reaction (9: withdrawn / reduced…) | `c_dechallenge` | Naranjo Q3, WHO-UMC, RUCAM |
| Reaction reappeared after reintroduction (10) | `c_rechallenge` | Naranjo Q4, WHO-UMC, RUCAM |
| Concomitant medications (11) | `concomitant_drugs` → informs `c_alt_causes`, `p_b2_interaction` | causality, preventability |
| Relevant investigations with dates (12) | → `c_toxic_level`, `c_objective_confirm` | Naranjo Q7/Q10 |
| Medical / medication history (13) | `comorbidities` → informs `c_alt_causes`, `p_a1_allergy_history` | causality, preventability |
| Seriousness (14: death / life-threatening / hospitalisation / disability) | `seriousness` → informs `s_caused_admission`, `s_increased_los`, `s_permanent_harm` | severity |
| Outcome (15: recovered / fatal / sequelae…) | `adr_outcome` → informs `s_contributed_death`, `s_permanent_harm` | severity |

**Derived automatically:** `c_temporal` is computed from therapy start date vs. reaction onset date.

## What the form does *not* capture (reviewer must add)

The official form has no field for these, so the engine flags them and a reviewer codes them once. Without them, the tool still scores — it just marks the row in the **Data flags** sheet rather than guessing.

| CASPER column | Needed for | Where the answer usually comes from |
|---|---|---|
| `c_prev_reports` | Naranjo Q1 | Known/literature ADR for that drug |
| `c_placebo` | Naranjo Q6 | Rarely available; leave "Not done" |
| `c_dose_response` | Naranjo Q8 | Dose-change records |
| `c_prior_exposure` | Naranjo Q9 | Sometimes in history (13) |
| `c_objective_confirm` | Naranjo Q10 | Investigations (12) |
| `c_event_definitive` | WHO-UMC "Certain" | Clinical/pharmacological judgement |
| `s_caused_admission`, `s_increased_los`, `s_intensive_care`, `s_permanent_harm`, `s_contributed_death` | Hartwig level 4–7 | Seriousness (14) + outcome (15), coded explicitly |
| `p_a2`, `p_a3`, `p_a5`, `p_b1`, `p_b3`, `p_b4` | Schumock preventability | Appropriateness/monitoring judgement |

## Bottom line by assessment

- **WHO-UMC causality** — fully supported by the form. True one-click.
- **Naranjo causality** — ~6–7 of 10 items come from the form; the rest default to "Don't know / Not done" unless a reviewer adds them. Score is valid but conservative until completed.
- **Hartwig–Siegel severity** — partially derivable from seriousness + outcome; explicit coding of admission/LOS/ICU sharpens it.
- **Schumock–Thornton preventability** — mostly a judgement instrument; the form supplies allergy history and interactions, the reviewer supplies appropriateness.

## Workflow

1. Export form data into the **CASPER upload template** columns (one row per drug–ADR pair). The table above tells you which form box fills which column.
2. Fill the reviewer-only columns where known; leave blank where not.
3. Run the engine (`python casper_engine.py data.xlsx`) or the app (`streamlit run casper_app.py`) and download the results workbook.
4. Check the **Data flags** sheet — those rows had missing inputs and should be reviewed before publication.
