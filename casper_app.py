"""
CASPER web app — upload an ADR dataset, click once, get causality + severity +
preventability for every case, with a downloadable results workbook.

Run:
    pip install streamlit pandas openpyxl
    streamlit run casper_app.py

Author: Puneet Paliwal. Licensed under Apache-2.0.
"""
import io
import pandas as pd
import streamlit as st
import casper_engine as ce

st.set_page_config(page_title="CASPER — ADR Assessment", layout="wide")
st.title("CASPER")
st.caption("Causality, Severity & Preventability Evaluation for Research")

st.markdown(
    "Upload an Excel/CSV built from the **CASPER upload template** "
    "(or a PvPI ADR-form export mapped to those columns). "
    "CASPER scores **Naranjo**, **WHO-UMC**, **Hartwig–Siegel severity** and "
    "**Schumock–Thornton preventability** for every row in one click."
)
st.info("CASPER standardises and makes scoring reproducible. It does not replace "
        "clinical judgement, and scales may legitimately disagree.", icon="⚠️")

up = st.file_uploader("Upload ADR dataset (.xlsx or .csv)", type=["xlsx", "csv"])

if up is not None:
    suffix = ".csv" if up.name.lower().endswith(".csv") else ".xlsx"
    tmp = "_uploaded" + suffix
    with open(tmp, "wb") as f:
        f.write(up.getbuffer())

    if st.button("Run assessment", type="primary"):
        df = ce.load_rows(tmp)
        if len(df) == 0:
            st.error("No valid case rows found. Each row needs a case_id and a suspected_drug.")
        else:
            res = ce.score(df)
            n = len(res)
            st.success(f"Scored {n} case(s).")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Naranjo: Probable+Definite",
                      int(res["Naranjo causality"].isin(["Probable", "Definite"]).sum()))
            c2.metric("WHO-UMC: Certain+Probable",
                      int(res["WHO-UMC causality"].isin(["Certain", "Probable / Likely"]).sum()))
            c3.metric("Severe ADRs", int((res["Severity grade"] == "Severe").sum()))
            c4.metric("Preventable",
                      int(res["Preventability"].isin(["Definitely preventable", "Probably preventable"]).sum()))

            st.subheader("Per-case results")
            st.dataframe(res, use_container_width=True)

            colA, colB = st.columns(2)
            with colA:
                st.subheader("Causality (WHO-UMC)")
                st.bar_chart(res["WHO-UMC causality"].value_counts())
                st.subheader("Severity")
                st.bar_chart(res["Severity grade"].value_counts())
            with colB:
                st.subheader("Causality (Naranjo)")
                st.bar_chart(res["Naranjo causality"].value_counts())
                st.subheader("Preventability")
                st.bar_chart(res["Preventability"].value_counts())

            flagged = res[res["Data flags"] != "complete"]
            if len(flagged):
                st.warning(f"{len(flagged)} row(s) had missing/assumed inputs — review the "
                           "'Missing/assumed fields' column before publishing.")

            out = "_RESULTS.xlsx"
            ce.write_report(res, out)
            with open(out, "rb") as f:
                st.download_button("Download results workbook (.xlsx)", f,
                                   file_name="CASPER_RESULTS.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
