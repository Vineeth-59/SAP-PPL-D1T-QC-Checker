import streamlit as st
import pandas as pd

from src.qc_engine import run_qc, make_excel


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="SAP PPL QC Checker",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown("""
<style>

    /* Main page */
    .main {
        padding-top: 1rem;
    }

    /* Header */
    .main-header {
        padding: 1.2rem 1.5rem;
        border-radius: 12px;
        background: linear-gradient(90deg, #0f4c81, #1769aa);
        color: white;
        margin-bottom: 1.2rem;
    }

    .main-header h1 {
        color: white;
        margin-bottom: 0.2rem;
        font-size: 2rem;
    }

    .main-header p {
        color: #e8f2fb;
        margin-bottom: 0;
        font-size: 0.95rem;
    }

    /* Section cards */
    .section-card {
        padding: 1.2rem;
        border-radius: 12px;
        border: 1px solid #d9e2ec;
        background-color: #ffffff;
        margin-bottom: 1rem;
    }

    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #17324d;
        margin-bottom: 0.7rem;
    }

    /* Metric cards */
    .metric-card {
        padding: 1rem;
        border-radius: 12px;
        background-color: #ffffff;
        border: 1px solid #d9e2ec;
        text-align: center;
        min-height: 105px;
    }

    .metric-label {
        font-size: 0.85rem;
        color: #64748b;
        margin-bottom: 0.3rem;
    }

    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #17324d;
    }

    /* Info box */
    .logic-box {
        padding: 1rem;
        border-radius: 10px;
        background-color: #f4f8fc;
        border: 1px solid #d7e6f3;
        margin-bottom: 1rem;
    }

    .logic-title {
        font-weight: 700;
        color: #174a73;
        margin-bottom: 0.4rem;
    }

    /* Status */
    .pass-box {
        padding: 0.8rem;
        border-radius: 10px;
        background-color: #eef9f1;
        border: 1px solid #b7dfc0;
        color: #176b2c;
        font-weight: 600;
        text-align: center;
    }

    .fail-box {
        padding: 0.8rem;
        border-radius: 10px;
        background-color: #fff1f1;
        border: 1px solid #efb4b4;
        color: #a12626;
        font-weight: 600;
        text-align: center;
    }

    /* Run button */
    div.stButton > button {
        width: 100%;
        height: 3rem;
        border-radius: 9px;
        font-size: 1.05rem;
        font-weight: 700;
    }

    /* Download buttons */
    div.stDownloadButton > button {
        width: 100%;
        border-radius: 8px;
    }

    /* Small footer */
    .footer {
        text-align: center;
        color: #7a8793;
        font-size: 0.8rem;
        padding-top: 2rem;
    }

</style>
""", unsafe_allow_html=True)


# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div class="main-header">
    <h1>✅ SAP PPL → D1T QC Checker</h1>
    <p>
        Automated configuration validation | Expected PPL vs Actual D1T
    </p>
     <p style="font-size: 0.85rem; margin-top: 0.5rem;">
        POC implemented by <b>Vineeth</b>
    </p>
</div>
""", unsafe_allow_html=True)


# =========================================================
# EXPLANATION
# =========================================================

st.markdown("""
<div class="logic-box">
    <div class="logic-title">How this QC works</div>
    <div>
        <b>PPL</b> represents the expected configuration.
        <b>D1T</b> represents the actual configuration.
        The QC engine compares the two and identifies
        <b>PASS</b>, <b>FAIL</b>, or <b>NOT CHECKABLE</b> results.
    </div>
</div>
""", unsafe_allow_html=True)


# =========================================================
# FILE UPLOAD SECTION
# =========================================================

st.markdown("""
<div class="section-title">📂 Input Files</div>
""", unsafe_allow_html=True)

upload_col1, upload_col2 = st.columns(2)

with upload_col1:

    st.markdown("### 📘 PPL — Expected Configuration")

    ppl_file = st.file_uploader(
        "Upload PPL Excel",
        type=["xlsx", "xls"],
        key="ppl_upload"
    )

    if ppl_file:
        st.success(f"✓ PPL loaded: {ppl_file.name}")
    else:
        st.info("Upload the Post Processing Log (PPL).")


with upload_col2:

    st.markdown("### 🗄️ D1T — Actual Configuration")

    d1t_file = st.file_uploader(
        "Upload D1T Excel",
        type=["xlsx", "xls"],
        key="d1t_upload"
    )

    if d1t_file:
        st.success(f"✓ D1T loaded: {d1t_file.name}")
    else:
        st.info("Upload the D1T configuration export.")


st.divider()


# =========================================================
# PROCESS ONLY WHEN FILES ARE AVAILABLE
# =========================================================

if ppl_file and d1t_file:

    try:

        # -------------------------------------------------
        # READ FILES
        # -------------------------------------------------

        ppl = pd.read_excel(ppl_file)
        d1t = pd.read_excel(d1t_file)


        # -------------------------------------------------
        # VALIDATE PPL
        # -------------------------------------------------

        required = [
            "Step",
            "Value Stream",
            "Agile Team",
            "Owner",
            "T-Code",
            "IMG-Activity",
            "Configuration key field NAMES (identifier)",
            "Template Value",
            "Rule",
            "New Org Value",
            "Post Processing field Value",
            "IMG Path",
            "Local/ Client Level/ Cross Client",
            "Remarks",
            "Configuration Guide",
            "CG Update needed",
            "Subject to change for localisation",
            "Owner.1"
        ]

        missing = [
            x for x in required
            if x not in ppl.columns
        ]

        if missing:

            st.error(
                "❌ PPL structure mismatch. Missing columns: "
                + ", ".join(missing)
            )

            st.stop()


        # -------------------------------------------------
        # VALIDATE D1T
        # -------------------------------------------------

        if "Company Code" not in d1t.columns:

            st.error(
                "❌ D1T must contain a 'Company Code' column."
            )

            st.stop()


        # =================================================
        # QC FILTERS
        # =================================================

        st.markdown("""
        <div class="section-title">🎯 QC Scope</div>
        """, unsafe_allow_html=True)

        filter1, filter2, filter3 = st.columns(3)


        # -------------------------------------------------
        # T-CODE
        # -------------------------------------------------

        with filter1:

            tcodes = sorted(
                ppl["T-Code"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )

            default_index = (
                tcodes.index("V_T001")
                if "V_T001" in tcodes
                else 0
            )

            tcode = st.selectbox(
                "⚙️ T-Code",
                tcodes,
                index=default_index
            )


        # -------------------------------------------------
        # VALUE STREAM
        # -------------------------------------------------

        with filter2:

            streams = (
                ppl.loc[
                    ppl["T-Code"].astype(str).str.strip() == tcode,
                    "Value Stream"
                ]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )

            streams = ["All"] + sorted(streams)

            stream = st.selectbox(
                "🔗 Value Stream",
                streams
            )


        # -------------------------------------------------
        # COMPANY CODE COUNT
        # -------------------------------------------------

        with filter3:

            companies = sorted(
                d1t["Company Code"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )

            st.metric(
                "Available Company Codes",
                len(companies)
            )


        # -------------------------------------------------
        # MULTI SELECT COMPANY CODES
        # -------------------------------------------------

        selected_companies = st.multiselect(
            "🏢 Company Code(s) to verify",
            companies,
            default=companies[:5],
            help="Select one or multiple Company Codes for batch QC."
        )


        if selected_companies:

            st.caption(
                f"Selected {len(selected_companies)} Company Code(s): "
                + ", ".join(selected_companies)
            )

        else:

            st.warning(
                "Please select at least one Company Code."
            )


        st.markdown("<br>", unsafe_allow_html=True)


        # =================================================
        # RUN QC
        # =================================================

        run_button = st.button(
            "🚀 RUN QC",
            type="primary",
            use_container_width=True
        )


        if run_button:

            if not selected_companies:

                st.warning(
                    "Please select at least one Company Code."
                )

                st.stop()


            # -------------------------------------------------
            # PROGRESS
            # -------------------------------------------------

            progress = st.progress(0)

            status_text = st.empty()

            all_results = []

            total = 0
            passed = 0
            failed = 0
            not_checkable = 0


            # -------------------------------------------------
            # RUN QC FOR EACH COMPANY
            # -------------------------------------------------

            total_companies = len(selected_companies)


            for i, company in enumerate(selected_companies):

                status_text.info(
                    f"Running QC for Company Code "
                    f"{company} ({i + 1}/{total_companies})..."
                )


                result, company_summary = run_qc(
                    ppl,
                    d1t,
                    company,
                    tcode,
                    stream
                )


                if result is not None and not result.empty:

                    all_results.append(result)


                total += company_summary["total"]

                passed += company_summary["pass"]

                failed += company_summary["fail"]

                not_checkable += company_summary["not_checkable"]


                progress.progress(
                    int(((i + 1) / total_companies) * 100)
                )


            status_text.success(
                "✓ QC execution completed."
            )


            # -------------------------------------------------
            # COMBINE RESULTS
            # -------------------------------------------------

            if all_results:

                result = pd.concat(
                    all_results,
                    ignore_index=True
                )

            else:

                result = pd.DataFrame()


            # -------------------------------------------------
            # SUMMARY
            # -------------------------------------------------

            pass_rate = (
                (passed / total) * 100
                if total > 0
                else 0
            )


            summary = {
                "total": total,
                "pass": passed,
                "fail": failed,
                "not_checkable": not_checkable,
                "pass_rate": pass_rate
            }


            st.divider()


            # =================================================
            # RESULTS DASHBOARD
            # =================================================

            st.markdown("""
            <div class="section-title">📊 QC Dashboard</div>
            """, unsafe_allow_html=True)


            m1, m2, m3, m4, m5 = st.columns(5)


            with m1:

                st.metric(
                    "Company Codes",
                    len(selected_companies)
                )


            with m2:

                st.metric(
                    "Total Checks",
                    total
                )


            with m3:

                st.metric(
                    "✅ PASS",
                    passed
                )


            with m4:

                st.metric(
                    "❌ FAIL",
                    failed
                )


            with m5:

                st.metric(
                    "⚠️ Not Checkable",
                    not_checkable
                )


            # -------------------------------------------------
            # PASS RATE
            # -------------------------------------------------

            st.markdown("### Overall QC Pass Rate")

            st.progress(
                min(pass_rate / 100, 1.0)
            )

            st.write(
                f"**{pass_rate:.1f}%** of configuration checks passed."
            )


            # -------------------------------------------------
            # OVERALL STATUS
            # -------------------------------------------------

            if failed == 0 and total > 0:

                st.markdown("""
                <div class="pass-box">
                    ✅ QC PASSED — No configuration discrepancies detected.
                </div>
                """, unsafe_allow_html=True)

            elif failed > 0:

                st.markdown("""
                <div class="fail-box">
                    ⚠️ QC COMPLETED — Configuration discrepancies require review.
                </div>
                """, unsafe_allow_html=True)


            st.markdown("<br>", unsafe_allow_html=True)


            # =================================================
            # DETAILED RESULTS
            # =================================================

            st.markdown("""
            <div class="section-title">📋 Detailed QC Evidence</div>
            """, unsafe_allow_html=True)


            if result.empty:

                st.info(
                    "No QC results were generated for the selected criteria."
                )

            else:

                st.dataframe(
                    result,
                    use_container_width=True,
                    hide_index=True
                )


                # =================================================
                # DOWNLOAD SECTION
                # =================================================

                st.markdown("### 📥 Download QC Evidence")

                download1, download2 = st.columns(2)


                # -------------------------------------------------
                # CSV
                # -------------------------------------------------

                with download1:

                    csv_data = result.to_csv(
                        index=False
                    ).encode("utf-8-sig")


                    st.download_button(
                        "⬇️ Download CSV Report",
                        csv_data,
                        "QC_Report_Multi_Company_Codes.csv",
                        "text/csv",
                        use_container_width=True
                    )


                # -------------------------------------------------
                # EXCEL
                # -------------------------------------------------

                with download2:

                    excel_data = make_excel(
                        result,
                        summary
                    )


                    st.download_button(
                        "⬇️ Download Excel Report",
                        excel_data,
                        "QC_Report_Multi_Company_Codes.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )


    # =========================================================
    # ERROR HANDLING
    # =========================================================

    except Exception as e:

        st.error(
            f"❌ An error occurred while running the QC: {e}"
        )


# =========================================================
# INITIAL STATE
# =========================================================

else:

    st.markdown("<br>", unsafe_allow_html=True)

    st.info(
        "👆 Upload both the PPL and D1T Excel files to begin QC."
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown("""
<div class="footer">
    SAP PPL → D1T QC Checker | Local Proof of Concept |
    Deterministic configuration validation
</div>
""", unsafe_allow_html=True)