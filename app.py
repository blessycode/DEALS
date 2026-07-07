from io import BytesIO

import pandas as pd
import streamlit as st

from assets import build_assets_workbook


def clean_df(df):
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df.columns = df.columns.astype(str).str.strip()

    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    return df


def clean_value_list(df, column, label):
    values = df[[column]].copy()
    values.columns = [label]
    values[label] = values[label].astype(str).str.strip()
    values = values[values[label].ne("")]
    values = values[values[label].str.lower().ne("nan")]
    return values.drop_duplicates().sort_values(label).reset_index(drop=True)


def sheet_names(uploaded_file):
    uploaded_file.seek(0)
    excel_file = pd.ExcelFile(uploaded_file)
    return excel_file.sheet_names


def read_sheet(uploaded_file, sheet_name):
    uploaded_file.seek(0)
    return clean_df(pd.read_excel(uploaded_file, sheet_name=sheet_name))


def make_comparison_report(df1, df2, column1, column2):
    label = "VALUE"
    values1 = clean_value_list(df1, column1, label)
    values2 = clean_value_list(df2, column2, label)

    merged = values1.merge(
        values2,
        on=label,
        how="outer",
        indicator=True
    )

    only_file1 = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])
    only_file2 = merged[merged["_merge"] == "right_only"].drop(columns=["_merge"])
    matched = merged[merged["_merge"] == "both"].drop(columns=["_merge"])

    counts = {
        "file1": len(values1),
        "file2": len(values2),
        "matched": len(matched),
        "only_file1": len(only_file1),
        "only_file2": len(only_file2),
    }
    same_values = only_file1.empty and only_file2.empty

    summary = pd.DataFrame({
        "CHECK": [
            "Same selected values in both files?",
            "Selected value count in first file",
            "Selected value count in second file",
            "Matched selected value count",
            "Only in first file count",
            "Only in second file count",
        ],
        "RESULT": [
            "YES" if same_values else "NO",
            f"{counts['file1']:,}",
            f"{counts['file2']:,}",
            f"{counts['matched']:,}",
            f"{counts['only_file1']:,}",
            f"{counts['only_file2']:,}",
        ]
    })

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="SUMMARY", index=False)
        matched.to_excel(writer, sheet_name="MATCHED", index=False)
        only_file1.to_excel(writer, sheet_name="ONLY_IN_FILE_1", index=False)
        only_file2.to_excel(writer, sheet_name="ONLY_IN_FILE_2", index=False)

    output.seek(0)
    return summary, output, same_values, counts


def make_assets_report(uploaded_file):
    uploaded_file.seek(0)
    workbook, mm_count, sec_count = build_assets_workbook(uploaded_file)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    summary = pd.DataFrame({
        "ITEM": ["MM rows", "SECURITIES rows"],
        "COUNT": [f"{mm_count:,}", f"{sec_count:,}"],
    })
    return summary, output, mm_count, sec_count


st.set_page_config(page_title="Deals Workbench", layout="wide")

st.markdown(
    """
    <style>
        :root {
            --ink: #172033;
            --muted: #667085;
            --line: #d8dee8;
            --panel: #ffffff;
            --panel-soft: #f7f9fc;
            --primary: #244b7a;
            --primary-dark: #183657;
            --accent: #0f8f7e;
            --warning: #a15c00;
            --success: #0c6b4f;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(15, 143, 126, 0.12), transparent 28rem),
                linear-gradient(180deg, #f5f7fb 0%, #eef2f7 100%);
            color: var(--ink);
        }

        .block-container {
            max-width: 1180px;
            padding-top: 1.2rem;
            padding-bottom: 3.5rem;
        }

        h1, h2, h3, p, label, span, div, button {
            letter-spacing: 0;
        }

        .app-shell {
            display: grid;
            gap: 1rem;
        }

        .hero {
            position: relative;
            overflow: hidden;
            border-radius: 8px;
            background: #10233f;
            border: 1px solid rgba(255, 255, 255, 0.16);
            box-shadow: 0 18px 50px rgba(16, 35, 63, 0.18);
            color: #ffffff;
            margin-bottom: 1rem;
        }

        .hero::before {
            content: "";
            position: absolute;
            inset: auto 0 0;
            height: 5px;
            background: linear-gradient(90deg, #0f8f7e, #d8a11d, #6c83b5);
            pointer-events: none;
        }

        .hero-inner {
            position: relative;
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            gap: 1.25rem;
            align-items: end;
            padding: 1.4rem 1.55rem 1.55rem;
        }

        .hero h1 {
            margin: 0;
            color: #ffffff;
            font-size: clamp(2rem, 4vw, 3rem);
            line-height: 1;
            font-weight: 850;
        }

        .hero p {
            max-width: 720px;
            margin: 0.55rem 0 0;
            color: #e7eef7 !important;
            font-size: 1.02rem;
            line-height: 1.55;
        }

        .hero-stat {
            min-width: 160px;
            border-left: 1px solid rgba(255, 255, 255, 0.2);
            padding-left: 1rem;
        }

        .hero-stat strong,
        .hero-stat span {
            display: block;
            color: #ffffff;
        }

        .hero-stat strong {
            font-size: 1.6rem;
            line-height: 1;
        }

        .hero-stat span {
            color: #d8e4f2;
            font-size: 0.84rem;
            margin-top: 0.35rem;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-color: var(--line);
            background: var(--panel);
            border-radius: 8px;
            box-shadow: 0 12px 34px rgba(23, 32, 51, 0.08);
        }

        .panel-heading {
            margin: -0.1rem 0 1rem;
            padding-bottom: 0.85rem;
            border-bottom: 1px solid var(--line);
        }

        .panel-heading h2 {
            margin: 0;
            color: var(--ink);
            font-size: 1.28rem;
            line-height: 1.25;
            font-weight: 820;
        }

        .panel-heading p {
            margin: 0.35rem 0 0;
            color: var(--muted) !important;
            font-size: 0.96rem;
            line-height: 1.45;
        }

        .step-card {
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 0.85rem 0.9rem;
            background: var(--panel-soft);
            min-height: 5.25rem;
        }

        .step-card span {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 1.55rem;
            height: 1.55rem;
            border-radius: 999px;
            background: var(--primary);
            color: #ffffff;
            font-size: 0.8rem;
            font-weight: 800;
            margin-bottom: 0.45rem;
        }

        .step-card strong {
            display: block;
            margin: 0;
            color: var(--ink);
            font-size: 0.98rem;
        }

        .step-card p {
            margin: 0.25rem 0 0;
            color: var(--muted);
            font-size: 0.87rem;
            line-height: 1.35;
        }

        .result-strip {
            border: 1px solid var(--line);
            border-left: 5px solid var(--accent);
            border-radius: 8px;
            background: #ffffff;
            padding: 0.9rem 1rem;
            margin: 0.35rem 0 0.85rem;
        }

        .result-strip.warning {
            border-left-color: var(--warning);
        }

        .result-strip.success {
            border-left-color: var(--success);
        }

        .result-strip strong {
            display: block;
            color: var(--ink);
            font-size: 1.02rem;
        }

        .result-strip p {
            margin: 0.25rem 0 0;
            color: var(--muted);
        }

        label,
        div[data-testid="stFileUploader"] label,
        div[data-testid="stSelectbox"] label {
            color: var(--ink) !important;
            font-weight: 760 !important;
        }

        div[data-testid="stMarkdownContainer"] p,
        div[data-testid="stCaptionContainer"] {
            color: var(--muted);
        }

        div[data-testid="stMetric"] {
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 0.85rem 0.95rem;
            background: #ffffff;
            box-shadow: 0 8px 20px rgba(23, 32, 51, 0.05);
        }

        div[data-testid="stMetricLabel"] p {
            color: var(--muted) !important;
            font-weight: 760;
        }

        div[data-testid="stMetricValue"] {
            color: var(--primary-dark) !important;
            font-size: 1.75rem;
        }

        div[data-testid="stFileUploader"] {
            border: 1px dashed #9ba8b8;
            border-radius: 8px;
            padding: 0.8rem 0.9rem;
            background: var(--panel-soft);
        }

        div[data-testid="stFileUploader"] small,
        div[data-testid="stFileUploader"] span,
        div[data-testid="stFileUploader"] p {
            color: var(--ink) !important;
        }

        div[data-testid="stFileUploader"] section {
            border-color: #aab6c5;
            background: #ffffff;
        }

        div[data-testid="stFileUploader"] button {
            background: var(--primary);
            border: 1px solid var(--primary);
            color: #ffffff !important;
            border-radius: 8px;
            font-weight: 760;
        }

        div[data-testid="stFileUploader"] button *,
        div[data-testid="stFileUploader"] button p,
        div[data-testid="stFileUploader"] button span,
        div[data-testid="stFileUploader"] button div {
            color: #ffffff !important;
        }

        div[data-testid="stFileUploader"] button:hover {
            background: var(--primary-dark);
            border-color: var(--primary-dark);
            color: #ffffff !important;
        }

        div[data-baseweb="select"] > div {
            border-color: #8fa0b5 !important;
            background: #ffffff !important;
            border-radius: 8px !important;
        }

        div[data-testid="stSelectbox"] div[data-baseweb="select"],
        div[data-testid="stSelectbox"] div[data-baseweb="select"] *,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] input {
            color: var(--ink) !important;
            -webkit-text-fill-color: var(--ink) !important;
        }

        div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
            background: #ffffff !important;
            color: var(--ink) !important;
        }

        div[data-baseweb="select"] span {
            color: var(--ink) !important;
            -webkit-text-fill-color: var(--ink) !important;
        }

        div[data-baseweb="select"] svg {
            color: #334155 !important;
            fill: #334155 !important;
        }

        .stButton > button,
        .stDownloadButton > button {
            width: 100%;
            border-radius: 8px;
            font-weight: 800;
            min-height: 2.8rem;
            border: 1px solid var(--primary);
        }

        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"] {
            background: var(--primary);
            color: #ffffff !important;
            border-color: var(--primary);
        }

        .stButton > button[kind="primary"] *,
        .stButton > button[kind="primary"] p,
        .stButton > button[kind="primary"] span,
        .stButton > button[kind="primary"] div,
        .stDownloadButton > button[kind="primary"] *,
        .stDownloadButton > button[kind="primary"] p {
            color: #ffffff !important;
        }

        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button[kind="primary"]:hover {
            background: var(--primary-dark);
            border-color: var(--primary-dark);
            color: #ffffff !important;
        }

        .stButton > button:disabled {
            color: #334155 !important;
            border-color: #cbd5e1;
            background: #eef2f7;
        }

        .stButton > button:disabled *,
        .stButton > button:disabled p,
        .stButton > button:disabled span,
        .stButton > button:disabled div {
            color: #334155 !important;
        }

        .stTabs [data-baseweb="tab-list"] {
            display: inline-flex;
            gap: 0.25rem;
            align-items: center;
            width: auto;
            max-width: 100%;
            margin: 0.1rem 0 1.15rem;
            padding: 0.3rem;
            border: 1px solid rgba(36, 75, 122, 0.14);
            border-radius: 8px;
            background: rgba(232, 239, 248, 0.86);
            box-shadow: 0 10px 28px rgba(23, 32, 51, 0.07);
        }

        .stTabs [data-baseweb="tab-border"],
        .stTabs [data-baseweb="tab-highlight"] {
            display: none;
        }

        .stTabs [data-baseweb="tab"] {
            min-height: 2.45rem;
            min-width: 8.75rem;
            padding: 0.58rem 1.05rem;
            border: 1px solid transparent;
            border-radius: 6px;
            background: transparent;
            color: #40546d;
            font-weight: 760;
            transition: background 150ms ease, border-color 150ms ease, box-shadow 150ms ease, color 150ms ease;
        }

        .stTabs [data-baseweb="tab"]:hover {
            background: rgba(36, 75, 122, 0.1);
            border-color: rgba(36, 75, 122, 0.14);
            color: var(--primary-dark);
        }

        .stTabs [data-baseweb="tab"] p {
            color: inherit !important;
            font-size: 0.96rem;
            font-weight: 760;
            line-height: 1.1;
            white-space: nowrap;
        }

        .stTabs [aria-selected="true"] {
            background: linear-gradient(180deg, #285684 0%, var(--primary) 100%);
            border-color: rgba(24, 54, 87, 0.5);
            box-shadow: 0 8px 18px rgba(24, 54, 87, 0.2);
            color: #ffffff;
            position: relative;
        }

        .stTabs [aria-selected="true"]::after {
            content: "";
            position: absolute;
            left: 0.7rem;
            right: 0.7rem;
            bottom: 0.24rem;
            height: 2px;
            border-radius: 999px;
            background: var(--accent);
        }

        .stTabs [aria-selected="true"] p {
            color: #ffffff !important;
        }

        div[data-testid="stAlert"] {
            border-radius: 8px;
            border: 1px solid var(--line);
        }

        div[data-testid="stAlert"] p,
        div[data-testid="stAlert"] span,
        div[data-testid="stAlert"] div {
            color: var(--ink) !important;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--line);
            border-radius: 8px;
            overflow: hidden;
        }

        @media (max-width: 760px) {
            .hero-inner {
                grid-template-columns: 1fr;
            }

            .hero-stat {
                border-left: 0;
                border-top: 1px solid rgba(255, 255, 255, 0.2);
                padding-left: 0;
                padding-top: 0.9rem;
            }

            .stTabs [data-baseweb="tab-list"] {
                display: flex;
                width: 100%;
            }

            .stTabs [data-baseweb="tab"] {
                flex: 1 1 0;
                min-width: 0;
                padding-inline: 0.7rem;
            }

            .stTabs [data-baseweb="tab"] p {
                font-size: 0.9rem;
            }
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="hero">
        <div class="hero-inner">
            <div>
                <h1>Deals Workbench</h1>
                <p>Prepare monthly assets workbooks and reconcile selected Excel values from one focused workspace.</p>
            </div>
            <div class="hero-stat">
                <strong>2</strong>
                <span>Excel workflows</span>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

assets_tab, compare_tab = st.tabs(["Assets Workbook", "Sheet Compare"])

with assets_tab:
    step1, step2, step3 = st.columns(3)
    step1.markdown(
        """<div class="step-card"><span>1</span><strong>Upload</strong><p>Select the dated source workbook.</p></div>""",
        unsafe_allow_html=True,
    )
    step2.markdown(
        """<div class="step-card"><span>2</span><strong>Build</strong><p>Extract monthly MM and securities rows.</p></div>""",
        unsafe_allow_html=True,
    )
    step3.markdown(
        """<div class="step-card"><span>3</span><strong>Download</strong><p>Review counts and save the output workbook.</p></div>""",
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown(
            """
            <div class="panel-heading">
                <h2>Assets Workbook</h2>
                <p>Upload the source workbook and generate monthly MM and securities sheets in one clean file.</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        assets_file = st.file_uploader("Source workbook", type=["xlsx", "xlsm"], key="assets_file")

        build_assets = st.button(
            "Build Assets Workbook",
            disabled=assets_file is None,
            type="primary",
            use_container_width=True
        )

        if build_assets:
            with st.spinner("Building assets workbook..."):
                assets_summary, assets_report, mm_count, sec_count = make_assets_report(assets_file)

            st.markdown(
                """
                <div class="result-strip success">
                    <strong>Assets workbook is ready.</strong>
                    <p>Review the extracted row counts below, then download the generated workbook.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            col1, col2 = st.columns(2)
            col1.metric("MM rows", f"{mm_count:,}")
            col2.metric("Securities rows", f"{sec_count:,}")
            st.dataframe(assets_summary, use_container_width=True, hide_index=True)
            st.download_button(
                label="Download Assets Workbook",
                data=assets_report,
                file_name="Combined_Assets_By_Month_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

with compare_tab:
    step1, step2, step3 = st.columns(3)
    step1.markdown(
        """<div class="step-card"><span>1</span><strong>Load files</strong><p>Upload both workbooks for comparison.</p></div>""",
        unsafe_allow_html=True,
    )
    step2.markdown(
        """<div class="step-card"><span>2</span><strong>Select values</strong><p>Choose the sheet and column in each file.</p></div>""",
        unsafe_allow_html=True,
    )
    step3.markdown(
        """<div class="step-card"><span>3</span><strong>Export</strong><p>Download matches and differences as Excel.</p></div>""",
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown(
            """
            <div class="panel-heading">
                <h2>Compare Sheets</h2>
                <p>Match selected values across two Excel files and export the matched and unmatched rows.</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        file_col1, file_col2 = st.columns(2)
        with file_col1:
            file1 = st.file_uploader("First Excel file", type=["xlsx", "xlsm"], key="compare_file_1")
        with file_col2:
            file2 = st.file_uploader("Second Excel file", type=["xlsx", "xlsm"], key="compare_file_2")

        if file1 and file2:
            sheet_col1, sheet_col2 = st.columns(2)
            with sheet_col1:
                sheet1 = st.selectbox("Sheet in first file", sheet_names(file1))
            with sheet_col2:
                sheet2 = st.selectbox("Sheet in second file", sheet_names(file2))

            df1 = read_sheet(file1, sheet1)
            df2 = read_sheet(file2, sheet2)

            column_col1, column_col2 = st.columns(2)
            with column_col1:
                column1 = st.selectbox("Column in first file", list(df1.columns))
            with column_col2:
                column2 = st.selectbox("Column in second file", list(df2.columns))

            preview_col1, preview_col2 = st.columns(2)
            with preview_col1:
                st.caption(f"{len(df1):,} rows loaded from {sheet1}")
            with preview_col2:
                st.caption(f"{len(df2):,} rows loaded from {sheet2}")

            if st.button("Compare Files", type="primary", use_container_width=True):
                with st.spinner("Comparing selected columns..."):
                    compare_summary, compare_report, same_values, counts = make_comparison_report(
                        df1,
                        df2,
                        column1,
                        column2
                    )

                if same_values:
                    st.markdown(
                        """
                        <div class="result-strip success">
                            <strong>The selected values match.</strong>
                            <p>No values appear exclusively in either file.</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        """
                        <div class="result-strip warning">
                            <strong>Differences found in the selected values.</strong>
                            <p>Use the exported workbook to review values that appear in only one file.</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                metric_col1, metric_col2, metric_col3 = st.columns(3)
                metric_col1.metric("Matched", f"{counts['matched']:,}")
                metric_col2.metric("Only in first", f"{counts['only_file1']:,}")
                metric_col3.metric("Only in second", f"{counts['only_file2']:,}")
                st.dataframe(compare_summary, use_container_width=True, hide_index=True)
                st.download_button(
                    label="Download Comparison Report",
                    data=compare_report,
                    file_name="Staff_Comparison_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        else:
            st.info("Upload both Excel files to choose sheets and columns.")
