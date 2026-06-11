# CASPER

**Causality, Severity & Preventability Evaluation for Research**

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)
[![DOI](https://img.shields.io/badge/DOI-10.5281%2Fzenodo.XXXXXXX-blue.svg)](https://doi.org/10.5281/zenodo.XXXXXXX)
<!-- Replace the DOI badge above with the real Zenodo DOI after your first release. -->

**▶ Live tool (no install):** `https://radiokiller1986.github.io/casper/`

CASPER is an open-source tool for automated assessment of adverse drug reactions (ADRs). Give it a single record or upload a dataset (CSV/Excel), and it generates standardized **causality**, **severity**, and **preventability** assessments with per-case and aggregate, research-ready outputs.

- **Causality:** Naranjo algorithm, WHO-UMC scale (implementation verified against the official WHO-UMC criteria)
- **Severity:** Modified Hartwig & Siegel scale
- **Preventability:** Modified Schumock & Thornton criteria
- **Batch mode:** assess whole research cohorts at once
- **Browser app:** `CASPER.html` runs fully offline — open it in any browser, no install, no server, no data leaves the device
- **Outputs:** per-case scores, items-known transparency, data-integrity flags, aggregate tables, charts, exportable reports (every export carries the citation)

> ⚠️ CASPER is a research and decision-support tool. It does not replace clinical judgment or formal regulatory assessment. Outputs depend on input completeness; the tool flags what it cannot determine rather than guessing.

---

## What's in this repository

| File | Purpose |
|---|---|
| `CASPER.html` | The offline browser app (upload → pick scales → results). Open directly or host on Pages. |
| `index.html` | Redirect so the Pages root URL opens the tool. |
| `casper_engine.py` | Python scoring engine for batch/scripted use. |
| `casper_app.py` | Optional Streamlit version of the app. |
| `CASPER_thesis_template.xlsx` | Data-collection template with the scoring columns and dropdowns. |
| `CASPER_validation_protocol.md` | Draft protocol for the formal validation study. |
| `LICENSE`, `NOTICE`, `CITATION.cff` | Licence and citation metadata. |

## Publish as a live link (GitHub Pages)

1. Create a public repo named `casper` and upload every file in this folder (including `index.html` and `.nojekyll`).
2. Repo **Settings → Pages → Build and deployment → Deploy from a branch → `main` / root → Save**.
3. After a minute your tool is live at `https://radiokiller1986.github.io/casper/`.
4. Replace `radiokiller1986` in this README, `CITATION.cff`, and the citation strings once the URL is known.

---

## How to cite

If you use CASPER in research, teaching, or any published work, please cite it. This is the only thing the author asks in return for free use.

> Paliwal, P. (2026). *CASPER: Causality, Severity & Preventability Evaluation for Research* (Version 0.1.0) [Computer software]. https://github.com