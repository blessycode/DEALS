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
            "YES" if only_file1.empty and only_file2.empty else "NO",
            len(values1),
            len(values2),
            len(matched),
            len(only_file1),
            len(only_file2),
        ]
    })

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="SUMMARY", index=False)
        matched.to_excel(writer, sheet_name="MATCHED", index=False)
        only_file1.to_excel(writer, sheet_name="ONLY_IN_FILE_1", index=False)
        only_file2.to_excel(writer, sheet_name="ONLY_IN_FILE_2", index=False)

    output.seek(0)
    return summary, output


def make_assets_report(uploaded_file):
    uploaded_file.seek(0)
    workbook, mm_count, sec_count = build_assets_workbook(uploaded_file)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    summary = pd.DataFrame({
        "ITEM": ["MM rows", "SECURITIES rows"],
        "COUNT": [mm_count, sec_count],
    })
    return summary, output


st.set_page_config(page_title="Deals Workbench", layout="wide")

st.markdown(
    """
    <style>
        :root {
            --ink: #111827;
            --muted: #4b5563;
            --line: #d7dee8;
            --panel: #ffffff;
            --navy: #10233f;
            --navy-2: #17365f;
            --accent: #0f9f8f;
            --gold: #d99b16;
            --soft: #f3f6fa;
        }

        .stApp {
            background: linear-gradient(180deg, #eef2f6 0%, #f8fafc 48%, #eef2f6 100%);
            color: var(--ink);
        }

        .block-container {
            padding-top: 1.35rem;
            padding-bottom: 3rem;
            max-width: 1120px;
        }

        h1, h2, h3, p, label, span, div {
            letter-spacing: 0;
        }

        .main-header {
            position: relative;
            overflow: hidden;
            border-radius: 8px;
            padding: 1.35rem 1.55rem 1.45rem;
            background: linear-gradient(135deg, var(--navy) 0%, var(--navy-2) 66%, #0d5f66 100%);
            border: 1px solid rgba(255, 255, 255, 0.12);
            box-shadow: 0 22px 60px rgba(16, 35, 63, 0.22);
            margin-bottom: 1.05rem;
            color: #ffffff;
        }

        .main-header::before {
            content: "";
            position: absolute;
            inset: 0;
            border-top: 5px solid var(--gold);
            pointer-events: none;
        }

        .main-header h1 {
            position: relative;
            margin: 0;
            color: #ffffff;
            font-size: 2.35rem;
            line-height: 1.02;
            letter-spacing: 0;
        }

        .main-header p {
            position: relative;
            max-width: 720px;
            margin: 0.58rem 0 0;
            color: #e7eef7 !important;
            font-size: 1rem;
        }

        .main-header *,
        div[data-testid="stMarkdownContainer"] .main-header p {
            color: #ffffff;
        }

        div[data-testid="stMarkdownContainer"] .main-header p {
            color: #e7eef7 !important;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-color: var(--line);
            background: var(--panel);
            box-shadow: 0 14px 35px rgba(17, 24, 39, 0.07);
        }

        div[data-testid="stVerticalBlockBorderWrapper"] > div > div > div {
            color: var(--ink);
        }

        .section-title {
            margin: 0.05rem 0 1rem;
            padding-bottom: 0.8rem;
            border-bottom: 1px solid var(--line);
        }

        .section-title h2 {
            margin: 0;
            color: var(--ink);
            font-size: 1.32rem;
            font-weight: 800;
            letter-spacing: 0;
        }

        .section-title p {
            margin: 0.3rem 0 0;
            color: var(--muted);
            font-weight: 500;
        }

        label,
        div[data-testid="stFileUploader"] label,
        div[data-testid="stSelectbox"] label {
            color: var(--ink) !important;
            font-weight: 700 !important;
        }

        div[data-testid="stMarkdownContainer"] p,
        div[data-testid="stCaptionContainer"] {
            color: var(--muted);
        }

        .section-title p,
        div[data-testid="stMarkdownContainer"] .section-title p {
            color: var(--muted) !important;
        }

        div[data-testid="stMetric"] {
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 0.8rem 0.95rem;
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        }

        div[data-testid="stMetricLabel"] p {
            color: var(--muted) !important;
            font-weight: 700;
        }

        div[data-testid="stMetricValue"] {
            color: var(--navy) !important;
        }

        div[data-testid="stFileUploader"] {
            border: 1px dashed #9aa8ba;
            border-radius: 8px;
            padding: 0.75rem 0.85rem;
            background: var(--soft);
        }

        div[data-testid="stFileUploader"] small,
        div[data-testid="stFileUploader"] span,
        div[data-testid="stFileUploader"] p {
            color: var(--ink) !important;
        }

        div[data-testid="stFileUploader"] section {
            border-color: #aab7c7;
            background: #ffffff;
        }

        div[data-testid="stFileUploader"] button {
            background: var(--navy);
            border: 1px solid var(--navy);
            color: #ffffff !important;
            border-radius: 8px;
            font-weight: 800;
        }

        div[data-testid="stFileUploader"] button *,
        div[data-testid="stFileUploader"] button p,
        div[data-testid="stFileUploader"] button span,
        div[data-testid="stFileUploader"] button div {
            color: #ffffff !important;
        }

        div[data-testid="stFileUploader"] button:hover {
            background: #0b1a30;
            border-color: #0b1a30;
            color: #ffffff !important;
        }

        div[data-testid="stFileUploaderDropzoneInstructions"] span {
            color: var(--ink) !important;
        }

        div[data-baseweb="select"] > div {
            border-color: #8fa0b5 !important;
            background: #ffffff !important;
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

        div[role="listbox"],
        ul[role="listbox"] {
            background: #ffffff !important;
            color: var(--ink) !important;
        }

        div[role="option"],
        ul[role="listbox"] li {
            background: #ffffff !important;
            color: var(--ink) !important;
            -webkit-text-fill-color: var(--ink) !important;
        }

        div[role="option"]:hover,
        ul[role="listbox"] li:hover {
            background: #eef2f7 !important;
            color: var(--ink) !important;
        }

        .stButton > button,
        .stDownloadButton > button {
            width: 100%;
            border-radius: 8px;
            font-weight: 800;
            min-height: 2.75rem;
            border: 1px solid var(--navy);
        }

        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"] {
            background: var(--navy);
            color: #ffffff !important;
            border-color: var(--navy);
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
            background: #0b1a30;
            border-color: #0b1a30;
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
            gap: 0.5rem;
            border-bottom: 1px solid var(--line);
            padding-bottom: 0.45rem;
        }

        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 0.55rem 1rem;
            background: #ffffff;
            border: 1px solid var(--line);
            color: var(--ink);
            font-weight: 800;
        }

        .stTabs [data-baseweb="tab"] p {
            color: var(--ink);
            font-weight: 800;
        }

        .stTabs [aria-selected="true"] {
            background: var(--navy);
            border-color: var(--navy);
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
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="main-header">
        <h1>Deals Workbench</h1>
        <p>Fast Excel preparation for assets and sheet comparisons.</p>
    </div>
    """,
    unsafe_allow_html=True
)

assets_tab, compare_tab = st.tabs(["Assets", "Compare"])

with assets_tab:
    with st.container(border=True):
        st.markdown(
            """
            <div class="section-title">
                <h2>Assets Workbook</h2>
                <p>Upload the source workbook and generate the monthly MM and securities output.</p>
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
                assets_summary, assets_report = make_assets_report(assets_file)

            mm_count = int(assets_summary.loc[assets_summary["ITEM"] == "MM rows", "COUNT"].iloc[0])
            sec_count = int(assets_summary.loc[assets_summary["ITEM"] == "SECURITIES rows", "COUNT"].iloc[0])

            st.success("Assets workbook is ready.")
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
    with st.container(border=True):
        st.markdown(
            """
            <div class="section-title">
                <h2>Compare Sheets</h2>
                <p>Match selected values across two Excel files and export the differences.</p>
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

            if st.button("Compare Files", type="primary", use_container_width=True):
                with st.spinner("Comparing selected columns..."):
                    compare_summary, compare_report = make_comparison_report(df1, df2, column1, column2)

                result = compare_summary.loc[
                    compare_summary["CHECK"] == "Same selected values in both files?",
                    "RESULT"
                ].iloc[0]
                matched = int(compare_summary.loc[
                    compare_summary["CHECK"] == "Matched selected value count",
                    "RESULT"
                ].iloc[0])
                only_first = int(compare_summary.loc[
                    compare_summary["CHECK"] == "Only in first file count",
                    "RESULT"
                ].iloc[0])
                only_second = int(compare_summary.loc[
                    compare_summary["CHECK"] == "Only in second file count",
                    "RESULT"
                ].iloc[0])

                if result == "YES":
                    st.success("The selected values match.")
                else:
                    st.warning("Differences found in the selected values.")

                metric_col1, metric_col2, metric_col3 = st.columns(3)
                metric_col1.metric("Matched", f"{matched:,}")
                metric_col2.metric("Only in first", f"{only_first:,}")
                metric_col3.metric("Only in second", f"{only_second:,}")
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
