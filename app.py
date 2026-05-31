from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


APP_TITLE = "Revenue Operations Dashboard"
DEFAULT_DATA_PATH = Path(__file__).parent / "data" / "revenue_operations.xlsx"
RAW_CRM_COLUMN = "_Client CRM Raw"

CLOSED_WON = "Closed Won"
CLOSED_LOST = "Closed Lost"
CLOSED_STAGES = {CLOSED_WON, CLOSED_LOST}

REQUIRED_COLUMNS = {
    "AQL date",
    "Source",
    "Client CRM",
    "Client country",
    "Number of sales reps",
    "PPC budget USD",
    "Subscription period",
    "Stage",
    "Loss reason description",
    "Closing Date",
}

# Centralized visual tokens keep the dashboard presentation consistent without
# touching the calculations, filters, or data-quality business rules.
COLORS = {
    "background": "#0B1020",
    "panel": "#151E31",
    "panel_alt": "#1D2A42",
    "border": "rgba(203, 213, 225, 0.20)",
    "text": "#F8FAFC",
    "muted": "#CBD5E1",
    "accent": "#38BDF8",
    "accent_2": "#34D399",
    "warning": "#FBBF24",
    "danger": "#FB7185",
}

CHART_TEMPLATE = "plotly_dark"
CHART_SEQUENCE = ["#38BDF8", "#34D399", "#A78BFA", "#FBBF24", "#FB7185", "#2DD4BF"]


@dataclass(frozen=True)
class QualityCheck:
    name: str
    mask: pd.Series


@dataclass(frozen=True)
class FilterContext:
    countries: list[str]
    crms: list[str]


st.set_page_config(
    page_title=APP_TITLE,
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_theme() -> None:
    """Apply a readable dark executive-dashboard visual layer."""
    st.markdown(
        f"""
        <style>
            :root {{
                --app-bg: {COLORS["background"]};
                --panel: {COLORS["panel"]};
                --panel-alt: {COLORS["panel_alt"]};
                --border: {COLORS["border"]};
                --text: {COLORS["text"]};
                --muted: {COLORS["muted"]};
                --accent: {COLORS["accent"]};
                --success: {COLORS["accent_2"]};
                --warning: {COLORS["warning"]};
                --danger: {COLORS["danger"]};
            }}

            .stApp {{
                background:
                    radial-gradient(circle at top left, rgba(56, 189, 248, 0.11), transparent 30rem),
                    linear-gradient(180deg, #0B1020 0%, #101827 100%);
                color: var(--text);
            }}

            .block-container {{
                padding-top: 2rem;
                padding-bottom: 3rem;
            }}

            [data-testid="stSidebar"] {{
                background: rgba(14, 21, 35, 0.98);
                border-right: 1px solid var(--border);
            }}

            [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
            [data-testid="stSidebar"] label {{
                color: #DCE6F4;
            }}

            [data-testid="stSidebar"] .stRadio label {{
                color: var(--text);
                font-weight: 680;
            }}

            [data-testid="stSidebar"] [role="radiogroup"] {{
                background: rgba(21, 30, 49, 0.84);
                border: 1px solid var(--border);
                border-radius: 8px;
                padding: 0.45rem 0.65rem;
            }}

            h1, h2, h3 {{
                letter-spacing: 0;
            }}

            h1 {{
                font-size: 2.1rem;
                font-weight: 780;
                margin-bottom: 0.25rem;
            }}

            .dashboard-subtitle {{
                color: #D4DEEC;
                font-size: 1rem;
                margin-bottom: 1.25rem;
            }}

            .section-title {{
                color: var(--text);
                font-size: 1.15rem;
                font-weight: 720;
                margin: 0.9rem 0 0.75rem;
            }}

            .kpi-card {{
                background: linear-gradient(145deg, rgba(21, 30, 49, 0.98), rgba(29, 42, 66, 0.98));
                border: 1px solid var(--border);
                border-radius: 8px;
                padding: 1.1rem 1.15rem;
                min-height: 8.2rem;
                box-shadow: 0 18px 42px rgba(0, 0, 0, 0.26);
            }}

            .kpi-topline {{
                align-items: center;
                color: var(--muted);
                display: flex;
                font-size: 0.8rem;
                font-weight: 650;
                gap: 0.55rem;
                letter-spacing: 0.02em;
                text-transform: uppercase;
            }}

            .kpi-icon {{
                align-items: center;
                background: rgba(56, 189, 248, 0.16);
                border: 1px solid rgba(56, 189, 248, 0.38);
                border-radius: 8px;
                color: var(--accent);
                display: inline-flex;
                font-size: 0.8rem;
                font-weight: 800;
                height: 2rem;
                justify-content: center;
                width: 2rem;
            }}

            .kpi-value {{
                color: var(--text);
                font-size: clamp(2rem, 4vw, 3rem);
                font-weight: 820;
                line-height: 1.05;
                margin-top: 0.9rem;
            }}

            .kpi-caption {{
                color: var(--muted);
                font-size: 0.86rem;
                margin-top: 0.45rem;
            }}

            .context-card,
            .insight-card,
            .quality-card {{
                background: rgba(21, 30, 49, 0.92);
                border: 1px solid var(--border);
                border-left: 4px solid var(--accent);
                border-radius: 8px;
                min-height: 12rem;
                padding: 1.3rem 1.35rem;
                box-shadow: 0 16px 36px rgba(0, 0, 0, 0.18);
            }}

            .insight-card.positive {{
                background: linear-gradient(145deg, rgba(7, 52, 42, 0.96), rgba(21, 30, 49, 0.96));
                border-color: rgba(52, 211, 153, 0.38);
                border-left-color: var(--success);
            }}

            .insight-card.important {{
                background: linear-gradient(145deg, rgba(59, 45, 12, 0.96), rgba(21, 30, 49, 0.96));
                border-color: rgba(251, 191, 36, 0.36);
                border-left-color: var(--warning);
            }}

            .insight-card.warning {{
                background: linear-gradient(145deg, rgba(67, 40, 14, 0.96), rgba(21, 30, 49, 0.96));
                border-color: rgba(251, 146, 60, 0.38);
                border-left-color: #FB923C;
            }}

            .insight-card.risk {{
                background: linear-gradient(145deg, rgba(68, 21, 32, 0.96), rgba(21, 30, 49, 0.96));
                border-color: rgba(251, 113, 133, 0.40);
                border-left-color: var(--danger);
            }}

            .context-card {{
                min-height: 6.8rem;
            }}

            .quality-card {{
                border-left-color: var(--accent);
                min-height: auto;
                padding: 1rem 1.1rem;
            }}

            .context-value {{
                color: var(--text);
                font-size: 1.35rem;
                font-weight: 780;
                line-height: 1.2;
                margin: 0.2rem 0 0.25rem;
            }}

            .context-caption {{
                color: var(--muted);
                font-size: 0.84rem;
                line-height: 1.45;
            }}

            .status-pill {{
                background: rgba(56, 189, 248, 0.16);
                border: 1px solid rgba(56, 189, 248, 0.34);
                border-radius: 999px;
                color: var(--accent);
                display: inline-block;
                font-size: 0.72rem;
                font-weight: 760;
                letter-spacing: 0.04em;
                margin-bottom: 0.55rem;
                padding: 0.22rem 0.55rem;
                text-transform: uppercase;
            }}

            .insight-label {{
                align-items: center;
                color: var(--text);
                display: flex;
                font-size: 1.02rem;
                font-weight: 700;
                gap: 0.5rem;
                letter-spacing: 0;
                line-height: 1.25;
                margin-bottom: 0.75rem;
            }}

            .insight-value {{
                color: #FFFFFF;
                font-size: 1.5rem;
                font-weight: 820;
                line-height: 1.1;
                margin-bottom: 0.65rem;
            }}

            .insight-body {{
                color: #E2E8F0;
                font-size: 0.96rem;
                line-height: 1.55;
            }}

            .finding-card {{
                background: linear-gradient(145deg, rgba(14, 35, 56, 0.98), rgba(21, 30, 49, 0.98));
                border: 1px solid rgba(56, 189, 248, 0.34);
                border-left: 5px solid var(--accent);
                border-radius: 8px;
                box-shadow: 0 16px 36px rgba(0, 0, 0, 0.22);
                margin: 1rem 0 1.3rem;
                padding: 1.35rem 1.45rem;
            }}

            .finding-card.positive {{
                background: linear-gradient(145deg, rgba(7, 52, 42, 0.96), rgba(21, 30, 49, 0.98));
                border-color: rgba(52, 211, 153, 0.38);
                border-left-color: var(--success);
            }}

            .finding-card.warning {{
                background: linear-gradient(145deg, rgba(67, 40, 14, 0.96), rgba(21, 30, 49, 0.98));
                border-color: rgba(251, 146, 60, 0.38);
                border-left-color: #FB923C;
            }}

            .finding-card.risk {{
                background: linear-gradient(145deg, rgba(68, 21, 32, 0.96), rgba(21, 30, 49, 0.98));
                border-color: rgba(251, 113, 133, 0.40);
                border-left-color: var(--danger);
            }}

            .finding-title {{
                color: var(--text);
                font-size: 1.28rem;
                font-weight: 760;
                margin-bottom: 1rem;
            }}

            .finding-row {{
                color: #E2E8F0;
                font-size: 0.98rem;
                line-height: 1.55;
                margin: 0.45rem 0;
            }}

            .finding-row strong {{
                color: var(--text);
            }}

            .investigation-grid {{
                display: grid;
                gap: 0.8rem;
            }}

            .investigation-step {{
                background: rgba(15, 23, 42, 0.48);
                border: 1px solid rgba(203, 213, 225, 0.14);
                border-left: 4px solid var(--accent);
                border-radius: 8px;
                padding: 0.9rem 1rem;
            }}

            .investigation-step.evidence {{
                border-left-color: var(--accent);
            }}

            .investigation-step.hypothesis,
            .insight-note {{
                background: linear-gradient(145deg, rgba(67, 51, 18, 0.68), rgba(21, 30, 49, 0.92));
                border: 1px solid rgba(251, 191, 36, 0.38);
                border-left: 4px solid var(--warning);
            }}

            .investigation-step.recommendation {{
                border-left-color: var(--success);
            }}

            .investigation-step.risk {{
                border-left-color: var(--danger);
            }}

            .investigation-step.positive {{
                border-left-color: var(--success);
            }}

            .investigation-step.warning {{
                border-left-color: #FB923C;
            }}

            .investigation-step.important {{
                border-left-color: var(--warning);
            }}

            .step-label {{
                color: #F8FAFC;
                font-size: 0.78rem;
                font-weight: 780;
                letter-spacing: 0.04em;
                margin-bottom: 0.35rem;
                text-transform: uppercase;
            }}

            .step-body {{
                color: #E2E8F0;
                font-size: 0.98rem;
                line-height: 1.55;
            }}

            .exec-summary-card {{
                background: rgba(21, 30, 49, 0.94);
                border: 1px solid rgba(203, 213, 225, 0.18);
                border-left: 4px solid var(--accent);
                border-radius: 8px;
                box-shadow: 0 14px 30px rgba(0, 0, 0, 0.18);
                min-height: 8rem;
                padding: 1rem 1.05rem;
            }}

            .exec-summary-card.positive {{
                border-left-color: var(--success);
            }}

            .exec-summary-card.warning {{
                border-left-color: #FB923C;
            }}

            .exec-summary-card.risk {{
                border-left-color: var(--danger);
            }}

            .exec-summary-card.important {{
                border-left-color: var(--warning);
            }}

            .exec-title {{
                color: var(--muted);
                font-size: 0.78rem;
                font-weight: 750;
                letter-spacing: 0.04em;
                margin-bottom: 0.45rem;
                text-transform: uppercase;
            }}

            .exec-metric {{
                color: #FFFFFF;
                font-size: 1.22rem;
                font-weight: 820;
                line-height: 1.2;
                margin-bottom: 0.45rem;
            }}

            .exec-body {{
                color: #CBD5E1;
                font-size: 0.86rem;
                line-height: 1.45;
            }}

            .report-card {{
                background: rgba(21, 30, 49, 0.9);
                border: 1px solid rgba(203, 213, 225, 0.16);
                border-radius: 8px;
                box-shadow: 0 14px 30px rgba(0, 0, 0, 0.16);
                margin: 1rem 0 1.35rem;
                padding: 1.45rem 1.55rem;
            }}

            .report-title {{
                color: var(--text);
                font-size: 1.2rem;
                font-weight: 780;
                letter-spacing: 0.02em;
                margin-bottom: 1rem;
                text-transform: uppercase;
            }}

            .report-section {{
                border-top: 1px solid rgba(203, 213, 225, 0.10);
                padding: 0.9rem 0 0;
                margin-top: 0.85rem;
            }}

            .report-section:first-of-type {{
                border-top: 0;
                margin-top: 0;
                padding-top: 0;
            }}

            .report-label {{
                color: #F8FAFC;
                font-size: 0.82rem;
                font-weight: 760;
                letter-spacing: 0.04em;
                margin-bottom: 0.35rem;
                text-transform: uppercase;
            }}

            .report-body {{
                color: #E2E8F0;
                font-size: 0.98rem;
                line-height: 1.62;
            }}

            .evidence-snapshot {{
                background: rgba(15, 23, 42, 0.62);
                border: 1px solid rgba(203, 213, 225, 0.14);
                border-left: 4px solid var(--accent);
                border-radius: 8px;
                margin: -0.6rem 0 1.35rem;
                padding: 1rem 1.1rem;
            }}

            .evidence-snapshot-title {{
                color: #CBD5E1;
                font-size: 0.78rem;
                font-weight: 760;
                letter-spacing: 0.04em;
                margin-bottom: 0.65rem;
                text-transform: uppercase;
            }}

            .evidence-snapshot-row {{
                align-items: baseline;
                border-top: 1px solid rgba(203, 213, 225, 0.08);
                display: flex;
                gap: 1rem;
                justify-content: space-between;
                padding: 0.55rem 0;
            }}

            .evidence-snapshot-row:first-of-type {{
                border-top: 0;
                padding-top: 0;
            }}

            .evidence-snapshot-label {{
                color: #E2E8F0;
                font-size: 0.92rem;
                line-height: 1.35;
            }}

            .evidence-snapshot-value {{
                color: #FFFFFF;
                flex: 0 0 auto;
                font-size: 1rem;
                font-weight: 820;
                text-align: right;
            }}

            div[data-testid="stMetric"] {{
                background: rgba(21, 30, 49, 0.92);
                border: 1px solid var(--border);
                border-radius: 8px;
                padding: 1rem;
            }}

            .recommendation-card {{
                background: linear-gradient(145deg, rgba(6, 42, 34, 0.96), rgba(21, 30, 49, 0.96));
                border: 1px solid rgba(52, 211, 153, 0.42);
                border-left: 4px solid var(--success);
                border-radius: 8px;
                box-shadow: 0 16px 36px rgba(0, 0, 0, 0.22);
                padding: 1rem 1.1rem;
            }}

            .sidebar-note {{
                color: #D4DEEC;
                font-size: 0.8rem;
                line-height: 1.42;
                margin: -0.2rem 0 0.75rem;
            }}

            .sidebar-insight-card {{
                background: rgba(21, 30, 49, 0.88);
                border: 1px solid rgba(203, 213, 225, 0.18);
                border-left: 3px solid var(--accent);
                border-radius: 8px;
                margin: 0.55rem 0;
                padding: 0.72rem 0.78rem;
            }}

            .sidebar-insight-title {{
                color: var(--text);
                font-size: 0.84rem;
                font-weight: 760;
                line-height: 1.3;
                margin-bottom: 0.25rem;
            }}

            .sidebar-insight-body {{
                color: #CBD5E1;
                font-size: 0.78rem;
                line-height: 1.42;
            }}

            .sidebar-insight-card.warning {{
                border-left-color: var(--warning);
            }}

            .sidebar-insight-card.risk {{
                border-left-color: var(--danger);
            }}

            .sidebar-insight-card.success {{
                border-left-color: var(--success);
            }}

            div[data-testid="stAlert"] {{
                background: rgba(251, 191, 36, 0.14);
                border: 1px solid rgba(251, 191, 36, 0.46);
                border-left: 4px solid var(--warning);
                border-radius: 8px;
                color: #FEF3C7;
            }}

            .warning-note {{
                background: rgba(251, 191, 36, 0.14);
                border: 1px solid rgba(251, 191, 36, 0.46);
                border-left: 4px solid var(--warning);
                border-radius: 8px;
                color: #FEF3C7;
                font-size: 0.95rem;
                line-height: 1.5;
                margin: 0.35rem 0 0.75rem;
                padding: 0.75rem 1rem;
            }}

            .stTabs [data-baseweb="tab-list"] {{
                gap: 0.4rem;
                border-bottom: 1px solid var(--border);
            }}

            .stTabs [data-baseweb="tab"] {{
                background: transparent;
                border-radius: 8px 8px 0 0;
                color: var(--muted);
                font-weight: 650;
                padding: 0.85rem 1rem;
            }}

            .stTabs [aria-selected="true"] {{
                background: rgba(56, 189, 248, 0.1);
                color: var(--text);
            }}

            [data-testid="stDataFrame"] {{
                border: 1px solid var(--border);
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 14px 30px rgba(0, 0, 0, 0.16);
            }}

            [data-testid="stPlotlyChart"] {{
                background: rgba(21, 30, 49, 0.72);
                border: 1px solid var(--border);
                border-radius: 8px;
                padding: 0.5rem;
                box-shadow: 0 14px 30px rgba(0, 0, 0, 0.16);
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_crm(value: object) -> str:
    """Standardize CRM labels before filters, grouping, win rates, and charts."""
    crm = normalize_text(value)
    compact = " ".join(crm.casefold().replace("-", " ").replace("_", " ").split())
    compact_no_space = compact.replace(" ", "")
    crm_words_present = "crm" in compact or "crm" in compact_no_space

    no_crm_signals = (
        compact == "no crm",
        compact_no_space == "nocrm",
        "not using" in compact,
        "not use" in compact,
        "does not use" in compact,
        "doesn't use" in compact,
        "do not use" in compact,
        "dont use" in compact,
        "don't use" in compact,
        "does not have" in compact and crm_words_present,
        "doesn't have" in compact and crm_words_present,
        "do not have" in compact and crm_words_present,
        "dont have" in compact and crm_words_present,
        "don't have" in compact and crm_words_present,
        "no use" in compact and crm_words_present,
        "no usage" in compact and crm_words_present,
        "without crm" in compact,
        "without a crm" in compact,
        "without any crm" in compact,
        "no client crm" in compact,
    )

    if any(no_crm_signals):
        return "No CRM"

    return crm


def first_existing_sheet(excel_file: str | Path | bytes) -> pd.DataFrame:
    return pd.read_excel(excel_file, sheet_name=0)


@st.cache_data(show_spinner=False)
def load_data(uploaded_file: bytes | None = None) -> pd.DataFrame:
    source = BytesIO(uploaded_file) if uploaded_file is not None else DEFAULT_DATA_PATH
    df = first_existing_sheet(source)
    missing_columns = REQUIRED_COLUMNS.difference(df.columns)

    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"The workbook is missing required columns: {missing}")

    df = df.copy()
    df[RAW_CRM_COLUMN] = df["Client CRM"].map(normalize_text)
    for column in ["Source", "Client CRM", "Client country", "PPC budget USD", "Stage", "Loss reason description"]:
        df[column] = df[column].map(normalize_text)
    df["Client CRM"] = df["Client CRM"].map(normalize_crm)

    df["AQL date"] = pd.to_datetime(df["AQL date"], errors="coerce")
    df["Closing Date"] = pd.to_datetime(df["Closing Date"], errors="coerce")
    df["Is Closed"] = df["Stage"].isin(CLOSED_STAGES)
    df["Is Won"] = df["Stage"].eq(CLOSED_WON)
    df["Is Lost"] = df["Stage"].eq(CLOSED_LOST)
    df["Is No CRM"] = df["Client CRM"].eq("No CRM")

    return df


def ordered_options(values: Iterable[str]) -> list[str]:
    return sorted({value for value in values if normalize_text(value)}, key=str.casefold)


def ppc_budget_sort_key(value: str) -> tuple[int, float, str]:
    text = normalize_text(value)
    if text.endswith("+"):
        number = pd.to_numeric(text.rstrip("+"), errors="coerce")
        return (1, float(number) if pd.notna(number) else float("inf"), text)
    if "-" in text:
        start = pd.to_numeric(text.split("-", 1)[0], errors="coerce")
        return (0, float(start) if pd.notna(start) else float("inf"), text)
    number = pd.to_numeric(text, errors="coerce")
    return (0, float(number) if pd.notna(number) else float("inf"), text)


def ppc_budget_estimate(value: str) -> float | None:
    text = normalize_text(value)
    if not text:
        return None
    if text.endswith("+"):
        number = pd.to_numeric(text.rstrip("+"), errors="coerce")
        return float(number) if pd.notna(number) else None
    if "-" in text:
        start_text, end_text = text.split("-", 1)
        start = pd.to_numeric(start_text, errors="coerce")
        end = pd.to_numeric(end_text, errors="coerce")
        if pd.notna(start) and pd.notna(end):
            return float((start + end) / 2)
        return None
    number = pd.to_numeric(text, errors="coerce")
    return float(number) if pd.notna(number) else None


def win_rate_summary(df: pd.DataFrame, dimension: str) -> pd.DataFrame:
    grouped = (
        df.groupby(dimension, dropna=False)
        .agg(
            deals=("Stage", "size"),
            closed_deals=("Is Closed", "sum"),
            won_deals=("Is Won", "sum"),
            lost_deals=("Is Lost", "sum"),
        )
        .reset_index()
    )
    grouped = grouped[grouped["closed_deals"] > 0].copy()
    grouped["win_rate"] = grouped["won_deals"] / grouped["closed_deals"]
    grouped["win_rate_label"] = (grouped["win_rate"] * 100).round(1).astype(str) + "%"
    return grouped.sort_values(["win_rate", "closed_deals"], ascending=[False, False])


def count_summary(df: pd.DataFrame, dimension: str, name: str = "deals") -> pd.DataFrame:
    return (
        df.groupby(dimension, dropna=False)
        .size()
        .reset_index(name=name)
        .sort_values(name, ascending=False)
    )


def quality_checks(df: pd.DataFrame) -> list[QualityCheck]:
    is_closed = df["Stage"].isin(CLOSED_STAGES)
    closing_date_missing = df["Closing Date"].isna()
    closing_date_present = df["Closing Date"].notna()
    loss_reason_missing = df["Loss reason description"].eq("")
    required_fields_missing = (
        df["Source"].eq("")
        | df["Client country"].eq("")
        | df["Stage"].eq("")
        | df["AQL date"].isna()
    )

    return [
        QualityCheck("Closed Won / Closed Lost with empty Closing Date", is_closed & closing_date_missing),
        QualityCheck("Closing Date filled but Stage is not closed", closing_date_present & ~is_closed),
        QualityCheck("Closed Lost with empty Loss reason description", df["Is Lost"] & loss_reason_missing),
        QualityCheck("AQL date later than Closing Date", df["AQL date"].notna() & closing_date_present & (df["AQL date"] > df["Closing Date"])),
        QualityCheck("Required fields are empty", required_fields_missing),
    ]


def apply_filters(df: pd.DataFrame) -> tuple[pd.DataFrame, FilterContext]:
    # Filter controls remain in the sidebar with the same options and filtering
    # logic; only labels and surrounding presentation are made cleaner.
    st.sidebar.header("Filters")

    countries = st.sidebar.multiselect(
        "Country",
        options=ordered_options(df["Client country"]),
        default=[],
        placeholder="All countries",
    )

    crms = st.sidebar.multiselect(
        "CRM",
        options=ordered_options(df["Client CRM"]),
        default=[],
        placeholder="All CRMs",
    )

    valid_dates = df["AQL date"].dropna()
    if valid_dates.empty:
        st.sidebar.warning("No valid AQL dates found.")
        return df, FilterContext(countries=countries, crms=crms)

    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()
    date_range = st.sidebar.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

    filtered = df.copy()
    if countries:
        filtered = filtered[filtered["Client country"].isin(countries)]
    if crms:
        filtered = filtered[filtered["Client CRM"].isin(crms)]
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered = filtered[
            filtered["AQL date"].dt.date.between(start_date, end_date, inclusive="both")
        ]

    return filtered, FilterContext(countries=countries, crms=crms)


def style_chart(fig: go.Figure, *, height: int = 420, show_legend: bool = False) -> go.Figure:
    """Give all charts a shared executive-dashboard treatment."""
    fig.update_layout(
        template=CHART_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter, Segoe UI, Arial, sans-serif", "color": COLORS["text"], "size": 13},
        title={"font": {"size": 18, "color": COLORS["text"]}, "x": 0.02, "xanchor": "left"},
        margin={"l": 32, "r": 56, "t": 64, "b": 44},
        height=height,
        showlegend=show_legend,
        coloraxis_showscale=False,
        bargap=0.34,
        hoverlabel={
            "bgcolor": "#151E31",
            "bordercolor": "rgba(203, 213, 225, 0.26)",
            "font": {"color": COLORS["text"]},
        },
    )
    fig.update_xaxes(
        automargin=True,
        gridcolor="rgba(203, 213, 225, 0.12)",
        linecolor="rgba(203, 213, 225, 0.22)",
        zerolinecolor="rgba(203, 213, 225, 0.16)",
        title_font={"color": COLORS["muted"]},
        tickfont={"color": COLORS["muted"]},
    )
    fig.update_yaxes(
        automargin=True,
        gridcolor="rgba(203, 213, 225, 0.08)",
        linecolor="rgba(203, 213, 225, 0.22)",
        title_font={"color": COLORS["muted"]},
        tickfont={"color": COLORS["muted"]},
    )
    return fig


def render_context_strip(df: pd.DataFrame) -> None:
    """Add executive context cards without changing the core KPI calculations."""
    closed_deals = int(df["Is Closed"].sum())
    open_deals = len(df) - closed_deals
    quality_summary = build_quality_summary(df)
    total_issues = int(quality_summary["Issue count"].sum())
    top_source = count_summary(df, "Source").iloc[0] if not df.empty else None

    context_cards = [
        (
            "Pipeline Exposure",
            f"{open_deals:,} open deals",
            "Open stages still visible after the current sidebar filters.",
        ),
        (
            "Primary Source",
            f"{top_source['Source']}" if top_source is not None else "No source data",
            f"{int(top_source['deals']):,} deals from the largest source." if top_source is not None else "No source concentration available.",
        ),
        (
            "Data Readiness",
            f"{total_issues:,} quality issues",
            "Validation flags in the selected view, shown in detail on the Data Quality tab.",
        ),
    ]

    columns = st.columns(3)
    for column, (label, value, caption) in zip(columns, context_cards):
        with column:
            st.markdown(
                f"""
                <div class="context-card">
                    <div class="status-pill">{label}</div>
                    <div class="context-value">{value}</div>
                    <div class="context-caption">{caption}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def draw_metric_row(df: pd.DataFrame) -> None:
    closed_deals = int(df["Is Closed"].sum())
    won_deals = int(df["Is Won"].sum())
    lost_deals = int(df["Is Lost"].sum())
    win_rate = won_deals / closed_deals if closed_deals else 0

    # The calculation lines above are unchanged; this section only upgrades the
    # KPI presentation into executive-style cards.
    cards = [
        ("TD", "Total Deals", f"{len(df):,}", "All opportunities in the selected view"),
        ("CD", "Closed Deals", f"{closed_deals:,}", f"{lost_deals:,} lost included in closed volume"),
        ("WD", "Won Deals", f"{won_deals:,}", "Closed Won opportunities"),
        ("WR", "Win Rate", f"{win_rate:.1%}", "Closed Won / closed deals"),
    ]

    columns = st.columns(4)
    for column, (icon, label, value, caption) in zip(columns, cards):
        with column:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-topline"><span class="kpi-icon">{icon}</span>{label}</div>
                    <div class="kpi-value">{value}</div>
                    <div class="kpi-caption">{caption}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def plot_win_rate(df: pd.DataFrame, dimension: str, title: str, *, top_n: int | None = None):
    summary = win_rate_summary(df, dimension)
    if top_n:
        summary = summary.head(top_n)

    fig = px.bar(
        summary,
        x="win_rate",
        y=dimension,
        orientation="h",
        color="closed_deals",
        color_continuous_scale=[[0, "#1D4ED8"], [0.55, "#38BDF8"], [1, "#22C55E"]],
        text="win_rate_label",
        hover_data={
            "win_rate": ":.1%",
            "closed_deals": True,
            "won_deals": True,
            "lost_deals": True,
            dimension: False,
        },
        labels={"win_rate": "Win rate", "closed_deals": "Closed deals", dimension: ""},
        title=title,
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_tickformat=".0%")
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    return style_chart(fig)


def selected_values(value: str | list[str] | tuple[str, ...] | None) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value] if value else []
    return [item for item in value if item]


def segment_win_rate_data(df: pd.DataFrame, dimension: str) -> pd.DataFrame:
    summary = (
        df.groupby(dimension, dropna=False)
        .agg(
            closed_deals=("Is Closed", "sum"),
            won_deals=("Is Won", "sum"),
        )
        .reset_index()
        .rename(columns={dimension: "segment_name"})
    )
    summary = summary[summary["closed_deals"] > 0].copy()

    if summary.empty:
        return pd.DataFrame(columns=["segment_name", "closed_deals", "won_deals", "win_rate", "chart_label"])

    summary["win_rate"] = summary["won_deals"] / summary["closed_deals"]
    summary["chart_label"] = (
        (summary["win_rate"] * 100).round(1).astype(str)
        + "% | "
        + summary["closed_deals"].astype(int).map("{:,}".format)
        + " closed"
    )
    return summary.sort_values(["closed_deals", "win_rate"], ascending=[False, False])


def get_country_chart_data(
    filtered_df: pd.DataFrame,
    selected_country: str | list[str] | None,
    selected_crm: str | list[str] | None = None,
) -> pd.DataFrame:
    selected = selected_values(selected_country)
    summary = segment_win_rate_data(filtered_df, "Client country")

    if selected:
        return summary[summary["segment_name"].isin(selected)].copy()

    if selected_values(selected_crm):
        return summary

    return summary[summary["closed_deals"] >= 20].copy()


def get_crm_chart_data(
    filtered_df: pd.DataFrame,
    selected_crm: str | list[str] | None,
    selected_country: str | list[str] | None = None,
) -> pd.DataFrame:
    selected = selected_values(selected_crm)
    summary = segment_win_rate_data(filtered_df, "Client CRM")

    if selected:
        return summary[summary["segment_name"].isin(selected)].copy()

    if selected_values(selected_country):
        return summary

    return summary[summary["closed_deals"] >= 10].copy()


def plot_segment_win_rate(summary: pd.DataFrame, title: str, *, global_mode: bool) -> go.Figure:
    if summary.empty:
        empty_message = (
            "No segments meet the minimum closed-deal threshold for the selected filters."
            if global_mode
            else "No closed deals for the selected filters."
        )
        fig = go.Figure()
        fig.update_layout(
            title=title,
            annotations=[
                {
                    "text": empty_message,
                    "xref": "paper",
                    "yref": "paper",
                    "x": 0.5,
                    "y": 0.5,
                    "showarrow": False,
                    "font": {"size": 14, "color": COLORS["muted"]},
                    "align": "center",
                }
            ],
            xaxis={"visible": False},
            yaxis={"visible": False},
        )
        return style_chart(fig)

    fig = px.bar(
        summary,
        x="win_rate",
        y="segment_name",
        orientation="h",
        color="closed_deals",
        color_continuous_scale=[[0, "#1D4ED8"], [0.55, "#38BDF8"], [1, "#22C55E"]],
        text="chart_label",
        hover_data={
            "win_rate": ":.1%",
            "closed_deals": True,
            "won_deals": True,
            "chart_label": False,
            "segment_name": False,
        },
        labels={"win_rate": "Win rate", "closed_deals": "Closed deals", "segment_name": ""},
        title=title,
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, xaxis_tickformat=".0%")
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    return style_chart(fig)


def crm_win_rate_chart_dataset(df: pd.DataFrame, *, min_closed_deals: int | None = 10) -> pd.DataFrame:
    """Build the CRM chart dataset with the Analytics tab threshold."""
    summary = segment_win_rate_data(df, "Client CRM")
    if min_closed_deals is None:
        return summary
    return summary[summary["closed_deals"] >= min_closed_deals].copy()


def country_win_rate_chart_dataset(df: pd.DataFrame, *, min_closed_deals: int | None = 20) -> pd.DataFrame:
    """Build the country chart dataset with the Analytics tab threshold."""
    summary = segment_win_rate_data(df, "Client country")
    if min_closed_deals is None:
        return summary
    return summary[summary["closed_deals"] >= min_closed_deals].copy()


def plot_crm_win_rate(df: pd.DataFrame, *, min_closed_deals: int | None = 10) -> go.Figure:
    return plot_segment_win_rate(
        crm_win_rate_chart_dataset(df, min_closed_deals=min_closed_deals),
        "Win Rate by CRM",
        global_mode=min_closed_deals is not None,
    )


def plot_country_win_rate(df: pd.DataFrame, *, min_closed_deals: int | None = 20) -> go.Figure:
    return plot_segment_win_rate(
        country_win_rate_chart_dataset(df, min_closed_deals=min_closed_deals),
        "Win Rate by Country",
        global_mode=min_closed_deals is not None,
    )


def plot_distribution(df: pd.DataFrame, dimension: str, title: str):
    summary = count_summary(df, dimension)
    fig = px.bar(
        summary,
        x="deals",
        y=dimension,
        orientation="h",
        text="deals",
        color="deals",
        color_continuous_scale=[[0, "#1E3A8A"], [1, "#38BDF8"]],
        labels={"deals": "Deal count", dimension: ""},
        title=title,
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"})
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    return style_chart(fig)


def plot_stage_distribution(df: pd.DataFrame) -> go.Figure:
    # Uses the existing stage count summary and adds display-only percentages
    # so the chart reads as a current-state distribution, not a conversion path.
    summary = count_summary(df, "Stage")
    total_deals = summary["deals"].sum()
    summary["share"] = summary["deals"] / total_deals if total_deals else 0
    summary["label"] = summary.apply(lambda row: f"{row['deals']:,} | {row['share']:.1%}", axis=1)

    fig = px.bar(
        summary,
        x="deals",
        y="Stage",
        orientation="h",
        text="label",
        color="deals",
        color_continuous_scale=[[0, "#1E3A8A"], [0.55, "#38BDF8"], [1, "#22C55E"]],
        labels={"deals": "Deal count", "Stage": ""},
        title="Deal Distribution by Stage",
        hover_data={
            "deals": ":,",
            "share": ":.1%",
            "Stage": False,
            "label": False,
        },
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"})
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    return style_chart(fig, height=440)


def plot_ppc_budget_win_rate(df: pd.DataFrame) -> go.Figure:
    budget_summary = win_rate_summary(df, "PPC budget USD")
    budget_summary = budget_summary.sort_values(
        "PPC budget USD",
        key=lambda series: series.map(ppc_budget_sort_key),
    )
    budget_fig = px.bar(
        budget_summary,
        x="PPC budget USD",
        y="win_rate",
        color="closed_deals",
        color_continuous_scale=[[0, "#1D4ED8"], [0.55, "#38BDF8"], [1, "#22C55E"]],
        text="win_rate_label",
        hover_data={
            "win_rate": ":.1%",
            "closed_deals": True,
            "won_deals": True,
            "lost_deals": True,
            "PPC budget USD": False,
        },
        labels={"win_rate": "Win rate", "PPC budget USD": "PPC budget USD", "closed_deals": "Closed deals"},
        title="Win Rate by PPC Budget",
    )
    budget_fig.update_layout(yaxis_tickformat=".0%")
    budget_fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    return style_chart(budget_fig)


def crm_ppc_heatmap_data(df: pd.DataFrame, *, min_closed_deals: int = 5) -> tuple[pd.DataFrame, list[str], list[str]]:
    summary = (
        df.groupby(["Client CRM", "PPC budget USD"], dropna=False)
        .agg(
            closed_deals=("Is Closed", "sum"),
            won_deals=("Is Won", "sum"),
        )
        .reset_index()
    )
    summary = summary[summary["closed_deals"] > 0].copy()

    if summary.empty:
        return summary, [], []

    summary["win_rate"] = summary["won_deals"] / summary["closed_deals"]
    summary["display_win_rate"] = summary["win_rate"].where(summary["closed_deals"] >= min_closed_deals)

    visible_crms = (
        summary.groupby("Client CRM", dropna=False)["display_win_rate"]
        .apply(lambda values: values.notna().any())
    )
    visible_crms = visible_crms[visible_crms].index.tolist()
    summary = summary[summary["Client CRM"].isin(visible_crms)].copy()

    if summary.empty:
        return summary, [], []

    crm_order = (
        summary.groupby("Client CRM", dropna=False)["closed_deals"]
        .sum()
        .sort_values(ascending=False)
        .index
        .tolist()
    )
    budget_order = sorted(summary["PPC budget USD"].dropna().unique().tolist(), key=ppc_budget_sort_key)

    return summary, crm_order, budget_order


def plot_crm_ppc_heatmap(df: pd.DataFrame, *, min_closed_deals: int = 5) -> go.Figure:
    summary, crm_order, budget_order = crm_ppc_heatmap_data(df, min_closed_deals=min_closed_deals)

    if summary.empty:
        fig = go.Figure()
        fig.update_layout(
            title="Win Rate by CRM and PPC Budget",
            annotations=[
                {
                    "text": "No CRM and PPC budget segments meet the minimum closed-deal threshold for the selected filters.",
                    "xref": "paper",
                    "yref": "paper",
                    "x": 0.5,
                    "y": 0.5,
                    "showarrow": False,
                    "font": {"size": 14, "color": COLORS["muted"]},
                    "align": "center",
                }
            ],
            xaxis={"visible": False},
            yaxis={"visible": False},
        )
        return style_chart(fig, height=520)

    value_matrix = (
        summary.pivot(index="Client CRM", columns="PPC budget USD", values="display_win_rate")
        .reindex(index=crm_order, columns=budget_order)
    )
    closed_matrix = (
        summary.pivot(index="Client CRM", columns="PPC budget USD", values="closed_deals")
        .reindex(index=crm_order, columns=budget_order)
    )
    won_matrix = (
        summary.pivot(index="Client CRM", columns="PPC budget USD", values="won_deals")
        .reindex(index=crm_order, columns=budget_order)
    )
    actual_win_rate_matrix = (
        summary.pivot(index="Client CRM", columns="PPC budget USD", values="win_rate")
        .reindex(index=crm_order, columns=budget_order)
    )

    customdata = []
    text = []
    for crm in crm_order:
        custom_row = []
        text_row = []
        for budget in budget_order:
            closed = closed_matrix.loc[crm, budget]
            won = won_matrix.loc[crm, budget]
            actual_win_rate = actual_win_rate_matrix.loc[crm, budget]
            shown_win_rate = value_matrix.loc[crm, budget]

            if pd.isna(closed):
                custom_row.append([crm, budget, None, None, None, "No closed deals"])
                text_row.append("")
            elif closed < min_closed_deals:
                custom_row.append([crm, budget, int(closed), int(won), actual_win_rate, "Low sample"])
                text_row.append("")
            else:
                custom_row.append([crm, budget, int(closed), int(won), actual_win_rate, "Shown"])
                text_row.append(f"{actual_win_rate:.0%}")
        customdata.append(custom_row)
        text.append(text_row)

    fig = go.Figure(
        data=go.Heatmap(
            z=value_matrix.values,
            x=budget_order,
            y=crm_order,
            customdata=customdata,
            text=text,
            texttemplate="%{text}",
            colorscale=[[0, "#1E3A8A"], [0.5, "#38BDF8"], [1, "#22C55E"]],
            zmin=0,
            zmax=1,
            colorbar={"title": "Win rate", "tickformat": ".0%"},
            hovertemplate=(
                "CRM: %{customdata[0]}<br>"
                "PPC budget: %{customdata[1]}<br>"
                "Closed deals: %{customdata[2]}<br>"
                "Won deals: %{customdata[3]}<br>"
                "Win rate: %{customdata[4]:.1%}<br>"
                "Status: %{customdata[5]}<extra></extra>"
            ),
        )
    )
    fig.update_layout(title="Win Rate by CRM and PPC Budget")
    return style_chart(fig, height=max(520, 34 * len(crm_order) + 140), show_legend=False)


def crm_adoption_maturity_summary(df: pd.DataFrame) -> pd.DataFrame:
    analysis_df = df.copy()
    analysis_df["CRM Adoption Segment"] = analysis_df["Client CRM"].where(
        analysis_df["Client CRM"].eq("No CRM"),
        "CRM Users",
    )
    analysis_df["PPC Budget Estimate"] = analysis_df["PPC budget USD"].map(ppc_budget_estimate)

    rows = []
    for segment in ["No CRM", "CRM Users"]:
        segment_df = analysis_df[analysis_df["CRM Adoption Segment"].eq(segment)]
        closed_deals = int(segment_df["Is Closed"].sum())
        won_deals = int(segment_df["Is Won"].sum())
        win_rate = won_deals / closed_deals if closed_deals else 0
        total_deals = len(segment_df)

        rows.append(
            {
                "Segment": segment,
                "Total Deals": total_deals,
                "Closed Deals": closed_deals,
                "Won Deals": won_deals,
                "Win Rate": win_rate,
                "Average PPC Budget": segment_df["PPC Budget Estimate"].mean(),
                "Median PPC Budget": segment_df["PPC Budget Estimate"].median(),
                "% of deals in 0-500 budget segment": (segment_df["PPC budget USD"].eq("0-500").sum() / total_deals) if total_deals else 0,
                "% of deals in 500-1000 budget segment": (segment_df["PPC budget USD"].eq("500-1000").sum() / total_deals) if total_deals else 0,
                "% of deals above 1000 budget": (segment_df["PPC Budget Estimate"].gt(1000).sum() / total_deals) if total_deals else 0,
            }
        )

    return pd.DataFrame(rows)


def format_crm_adoption_maturity_table(summary: pd.DataFrame) -> pd.DataFrame:
    formatted = summary.copy()
    formatted["Win Rate"] = formatted["Win Rate"].map(lambda value: f"{value:.1%}")
    for column in ["Average PPC Budget", "Median PPC Budget"]:
        formatted[column] = formatted[column].map(lambda value: "N/A" if pd.isna(value) else f"${value:,.0f}")
    for column in [
        "% of deals in 0-500 budget segment",
        "% of deals in 500-1000 budget segment",
        "% of deals above 1000 budget",
    ]:
        formatted[column] = formatted[column].map(lambda value: f"{value:.1%}")
    return formatted


def finding_card(title: str, investigated: str, evidence: str, conclusion: str) -> None:
    st.markdown(
        f"""
        <div class="finding-card">
            <div class="finding-title">{title}</div>
            <div class="finding-row"><strong>What we investigated &rarr;</strong> {investigated}</div>
            <div class="finding-row"><strong>Evidence &rarr;</strong> {evidence}</div>
            <div class="finding-row"><strong>Conclusion / Hypothesis &rarr;</strong> {conclusion}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_crm_adoption_maturity_section(df: pd.DataFrame) -> None:
    summary = crm_adoption_maturity_summary(df)
    formatted_summary = format_crm_adoption_maturity_table(summary)[
        [
            "Segment",
            "Total Deals",
            "Closed Deals",
            "Won Deals",
            "Win Rate",
            "Average PPC Budget",
            "Median PPC Budget",
        ]
    ]

    st.markdown('<div class="section-title">CRM Adoption vs Business Maturity</div>', unsafe_allow_html=True)
    st.dataframe(formatted_summary, hide_index=True, use_container_width=True)
    st.caption("Average and median PPC budget use estimated numeric values from budget bands, using band midpoints and the lower bound for open-ended bands.")

    lookup = summary.set_index("Segment")
    no_crm = lookup.loc["No CRM"]
    crm_users = lookup.loc["CRM Users"]
    _, sales_rep_stats = crm_sales_team_size_matrix(df)
    sales_lookup = sales_rep_stats.set_index("Segment")
    budget_rate = loss_category_rate_by_crm_adoption(
        df,
        "Budget / Price",
        "Budget/Price Losses",
        "Budget/Price Loss Rate",
    ).set_index("Segment")

    evidence = (
        f"No CRM win rate is {no_crm['Win Rate']:.1%} vs {crm_users['Win Rate']:.1%} for CRM Users. "
        f"Median PPC budget is ${no_crm['Median PPC Budget']:,.0f} for No CRM and ${crm_users['Median PPC Budget']:,.0f} for CRM Users. "
        f"Average sales reps are {sales_lookup.loc['No CRM', 'Average sales reps']:.1f} for No CRM vs "
        f"{sales_lookup.loc['CRM Users', 'Average sales reps']:.1f} for CRM Users. "
        f"Budget/Price loss rates are {budget_rate.loc['No CRM', 'Budget/Price Loss Rate']:.1%} for No CRM vs "
        f"{budget_rate.loc['CRM Users', 'Budget/Price Loss Rate']:.1%} for CRM Users."
    )
    finding_card(
        "Finding: No CRM underconverts, but budget alone does not explain the gap.",
        "We compared No CRM against all CRM-adopting companies on conversion, PPC budget, sales-team size, and pricing objections.",
        evidence,
        "Businesses without CRM adoption convert worse than CRM users (29.3% vs 36.2%). This difference cannot be fully explained by budget size, sales-team size, or pricing objections. CRM adoption appears to act as a proxy for broader operational maturity rather than directly causing improved conversion outcomes.",
    )


def sales_rep_bucket(value: object) -> str:
    reps = pd.to_numeric(value, errors="coerce")
    if pd.isna(reps):
        return "Unknown"
    if reps <= 1:
        return "1 rep"
    if reps <= 5:
        return "2-5 reps"
    if reps <= 10:
        return "6-10 reps"
    if reps <= 20:
        return "11-20 reps"
    return "21+ reps"


def win_rate_table(df: pd.DataFrame, dimension: str, order: list[str] | None = None) -> pd.DataFrame:
    summary = (
        df.groupby(dimension, dropna=False)
        .agg(
            **{
                "Total Deals": ("Stage", "size"),
                "Closed Deals": ("Is Closed", "sum"),
                "Won Deals": ("Is Won", "sum"),
            }
        )
        .reset_index()
        .rename(columns={dimension: "Segment"})
    )
    summary["Win Rate"] = summary["Won Deals"] / summary["Closed Deals"].replace(0, pd.NA)
    summary["Win Rate"] = summary["Win Rate"].fillna(0)

    if order:
        summary["Segment"] = pd.Categorical(summary["Segment"], categories=order, ordered=True)
        summary = summary.sort_values("Segment")
        summary["Segment"] = summary["Segment"].astype(str)
    else:
        summary = summary.sort_values(["Closed Deals", "Win Rate"], ascending=[False, False])

    return summary


def format_win_rate_table(summary: pd.DataFrame) -> pd.DataFrame:
    formatted = summary.copy()
    formatted["Win Rate"] = formatted["Win Rate"].map(lambda value: f"{value:.1%}")
    return formatted


def sales_maturity_tables(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    analysis_df = df.copy()
    analysis_df["Sales Rep Bucket"] = analysis_df["Number of sales reps"].map(sales_rep_bucket)
    rep_order = ["1 rep", "2-5 reps", "6-10 reps", "11-20 reps", "21+ reps", "Unknown"]
    reps_summary = win_rate_table(analysis_df, "Sales Rep Bucket", rep_order)

    subscription_df = analysis_df.copy()
    subscription_df["Subscription period"] = subscription_df["Subscription period"].fillna("Unknown")
    subscription_summary = win_rate_table(subscription_df, "Subscription period")

    return reps_summary, subscription_summary


def sales_maturity_interpretation(reps_summary: pd.DataFrame, subscription_summary: pd.DataFrame) -> str:
    known_reps = reps_summary[~reps_summary["Segment"].eq("Unknown")].copy()
    best_rep_segment = known_reps.sort_values(["Win Rate", "Closed Deals"], ascending=[False, False]).head(1)
    largest_team_segments = known_reps[known_reps["Segment"].isin(["11-20 reps", "21+ reps"])]
    best_subscription = subscription_summary[~subscription_summary["Segment"].eq("Unknown")].sort_values(
        ["Win Rate", "Closed Deals"],
        ascending=[False, False],
    ).head(1)

    if best_rep_segment.empty:
        rep_text = "Sales team size data is not sufficient to determine whether larger teams convert better."
    elif not largest_team_segments.empty and best_rep_segment.iloc[0]["Segment"] not in ["11-20 reps", "21+ reps"]:
        rep_text = (
            f"The strongest sales-team bucket is {best_rep_segment.iloc[0]['Segment']} at {best_rep_segment.iloc[0]['Win Rate']:.1%}. "
            "In the current filtered data, larger sales teams do not appear to convert better, so this field may weaken a simple 'larger team equals higher maturity' interpretation."
        )
    else:
        rep_text = (
            f"The strongest sales-team bucket is {best_rep_segment.iloc[0]['Segment']} at {best_rep_segment.iloc[0]['Win Rate']:.1%}, "
            "which may indicate that sales-team scale is associated with conversion quality."
        )

    subscription_text = (
        f"The strongest subscription-period segment is {best_subscription.iloc[0]['Segment']} at {best_subscription.iloc[0]['Win Rate']:.1%}. Because subscription period is often populated on won deals, this suggests commitment may be related to conversion but requires further validation."
        if not best_subscription.empty
        else "Subscription period data requires further validation before linking it to conversion quality."
    )

    return (
        f"{rep_text} {subscription_text} CRM adoption alone may not fully explain conversion differences. "
        "Sales team size and subscription commitment can act as additional maturity signals. If larger sales teams or longer subscription periods show stronger win rates, this would support the hypothesis that operational maturity is a key driver of conversion. "
        "These patterns are directional and require further validation before making causal claims."
    )


def render_sales_maturity_subscription_section(df: pd.DataFrame) -> None:
    reps_summary, subscription_summary = sales_maturity_tables(df)

    st.markdown('<div class="section-title">Sales Maturity & Subscription Insights</div>', unsafe_allow_html=True)
    left, right = st.columns(2)
    with left:
        st.markdown("#### Win Rate by Number of Sales Reps")
        st.dataframe(format_win_rate_table(reps_summary), hide_index=True, use_container_width=True)
    with right:
        st.markdown("#### Win Rate by Subscription Period")
        st.dataframe(format_win_rate_table(subscription_summary), hide_index=True, use_container_width=True)

    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">Interpretation</div>
            <div class="insight-body">{sales_maturity_interpretation(reps_summary, subscription_summary)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def crm_sales_team_size_matrix(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    analysis_df = df.copy()
    analysis_df["CRM Adoption Segment"] = analysis_df["Client CRM"].where(
        analysis_df["Client CRM"].eq("No CRM"),
        "CRM Users",
    )
    analysis_df["Sales Rep Bucket"] = analysis_df["Number of sales reps"].map(sales_rep_bucket)
    analysis_df["Sales Reps Numeric"] = pd.to_numeric(analysis_df["Number of sales reps"], errors="coerce")

    bucket_order = ["1 rep", "2-5 reps", "6-10 reps", "11-20 reps", "21+ reps"]
    segment_order = ["No CRM", "CRM Users"]
    count_matrix = (
        analysis_df[analysis_df["Sales Rep Bucket"].isin(bucket_order)]
        .pivot_table(
            index="CRM Adoption Segment",
            columns="Sales Rep Bucket",
            values="Stage",
            aggfunc="size",
            fill_value=0,
        )
        .reindex(index=segment_order, columns=bucket_order, fill_value=0)
    )
    share_matrix = count_matrix.div(count_matrix.sum(axis=1).replace(0, pd.NA), axis=0).fillna(0)

    display_matrix = count_matrix.copy().astype(str)
    for segment in segment_order:
        for bucket in bucket_order:
            display_matrix.loc[segment, bucket] = f"{int(count_matrix.loc[segment, bucket]):,} ({share_matrix.loc[segment, bucket]:.1%})"

    display_matrix = display_matrix.reset_index().rename(columns={"CRM Adoption Segment": "Segment"})

    sales_rep_stats = (
        analysis_df.groupby("CRM Adoption Segment", dropna=False)
        .agg(
            **{
                "Average sales reps": ("Sales Reps Numeric", "mean"),
                "Median sales reps": ("Sales Reps Numeric", "median"),
            }
        )
        .reindex(segment_order)
        .reset_index()
        .rename(columns={"CRM Adoption Segment": "Segment"})
    )

    return display_matrix, sales_rep_stats


def format_sales_rep_stats(stats: pd.DataFrame) -> pd.DataFrame:
    formatted = stats.copy()
    for column in ["Average sales reps", "Median sales reps"]:
        formatted[column] = formatted[column].map(lambda value: "N/A" if pd.isna(value) else f"{value:.1f}")
    return formatted


def crm_sales_team_size_interpretation(matrix: pd.DataFrame, stats: pd.DataFrame) -> str:
    stat_lookup = stats.set_index("Segment")
    no_crm_avg = stat_lookup.loc["No CRM", "Average sales reps"]
    crm_users_avg = stat_lookup.loc["CRM Users", "Average sales reps"]
    no_crm_median = stat_lookup.loc["No CRM", "Median sales reps"]
    crm_users_median = stat_lookup.loc["CRM Users", "Median sales reps"]

    share_lookup = matrix.set_index("Segment")

    def parse_share(value: str) -> float:
        percent_text = value.split("(", 1)[1].rstrip(")")
        return float(percent_text.rstrip("%")) / 100

    no_crm_small_share = parse_share(share_lookup.loc["No CRM", "1 rep"]) + parse_share(share_lookup.loc["No CRM", "2-5 reps"])
    crm_small_share = parse_share(share_lookup.loc["CRM Users", "1 rep"]) + parse_share(share_lookup.loc["CRM Users", "2-5 reps"])
    no_crm_large_share = parse_share(share_lookup.loc["No CRM", "11-20 reps"]) + parse_share(share_lookup.loc["No CRM", "21+ reps"])
    crm_large_share = parse_share(share_lookup.loc["CRM Users", "11-20 reps"]) + parse_share(share_lookup.loc["CRM Users", "21+ reps"])

    small_team_text = (
        "No CRM businesses are more concentrated in smaller sales teams, which may indicate lower sales-organization maturity."
        if no_crm_small_share > crm_small_share
        else "No CRM businesses are not more concentrated in smaller sales teams in the current filtered view."
    )
    larger_team_text = (
        "CRM Users are more represented in larger sales-team buckets, which suggests CRM adoption may be associated with more developed sales operations."
        if crm_large_share > no_crm_large_share
        else "CRM Users are not clearly more concentrated in larger sales-team buckets in the current filtered view."
    )
    gap_text = (
        "Because No CRM has fewer average sales reps, sales-team size may explain part of the observed CRM win-rate gap, but this requires further investigation."
        if no_crm_avg < crm_users_avg
        else "Average sales-team size does not clearly explain the CRM win-rate gap on its own, and requires further investigation."
    )

    return (
        f"{small_team_text} {larger_team_text} {gap_text} "
        f"Average sales reps are {no_crm_avg:.1f} for No CRM vs {crm_users_avg:.1f} for CRM Users; medians are {no_crm_median:.1f} vs {crm_users_median:.1f}. "
        "This is directional evidence only and should not be treated as causal."
    )


def render_crm_sales_team_size_section(df: pd.DataFrame) -> None:
    matrix, stats = crm_sales_team_size_matrix(df)

    st.markdown('<div class="section-title">CRM Adoption vs Sales Team Size</div>', unsafe_allow_html=True)
    st.dataframe(matrix, hide_index=True, use_container_width=True)
    st.caption("Each cell shows Total Deals and percent of that CRM adoption segment.")
    st.dataframe(format_sales_rep_stats(stats), hide_index=True, use_container_width=True)

    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">Interpretation</div>
            <div class="insight-body">{crm_sales_team_size_interpretation(matrix, stats)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def top_segment_table(df: pd.DataFrame, group_col: str, bucket: str, *, top_n: int = 5) -> pd.DataFrame:
    bucket_df = df[df["Sales Rep Bucket"].eq(bucket)]
    summary = (
        bucket_df.groupby(group_col, dropna=False)
        .agg(
            **{
                "Total deals": ("Stage", "size"),
                "Closed deals": ("Is Closed", "sum"),
                "Won deals": ("Is Won", "sum"),
            }
        )
        .reset_index()
        .rename(columns={group_col: "Segment"})
    )
    if summary.empty:
        return pd.DataFrame(columns=["Segment", "Total deals", "Closed deals", "Won deals", "Win rate"])

    summary["Win rate"] = summary["Won deals"] / summary["Closed deals"].replace(0, pd.NA)
    summary["Win rate"] = summary["Win rate"].fillna(0)
    return summary.sort_values(["Total deals", "Closed deals"], ascending=[False, False]).head(top_n)


def format_diagnostic_table(summary: pd.DataFrame) -> pd.DataFrame:
    formatted = summary.copy()
    if "Win rate" in formatted:
        formatted["Win rate"] = formatted["Win rate"].map(lambda value: f"{value:.1%}")
    return formatted


def diagnostic_concentration_insights(df: pd.DataFrame) -> list[str]:
    bucket_order = ["1 rep", "2-5 reps", "6-10 reps", "11-20 reps", "21+ reps"]
    larger_buckets = ["11-20 reps", "21+ reps"]
    insights = []

    for group_col, label in [
        ("Client country", "countries"),
        ("Client CRM", "CRM systems"),
        ("Source", "acquisition channels"),
    ]:
        larger_df = df[df["Sales Rep Bucket"].isin(larger_buckets)]
        if larger_df.empty:
            continue

        top_segment = larger_df[group_col].value_counts().head(1)
        if top_segment.empty:
            continue

        segment_name = top_segment.index[0]
        larger_share = top_segment.iloc[0] / len(larger_df)
        all_share = df[group_col].eq(segment_name).sum() / len(df) if len(df) else 0
        if larger_share > all_share:
            insights.append(
                f"Larger sales-team buckets over-index in {segment_name} for {label} ({larger_share:.1%} of 11-20 and 21+ rep deals vs {all_share:.1%} overall), which may indicate a mix effect that requires further investigation."
            )
        else:
            insights.append(
                f"The largest {label} segment among larger sales teams is {segment_name}, but it does not over-index versus the overall mix ({larger_share:.1%} of 11-20 and 21+ rep deals vs {all_share:.1%} overall). This weakens a simple concentration explanation and requires further investigation."
            )

    if not insights:
        insights.append("The larger sales-team buckets do not show a clear concentration pattern in the current filtered view; this requires further investigation.")

    insights.append(
        "If larger sales teams are concentrated in specific countries, CRM systems, or acquisition channels with lower conversion, those concentration effects could potentially explain part of the lower win rates observed in larger sales-team segments."
    )
    return insights


def render_largest_sales_team_sources_section(df: pd.DataFrame) -> None:
    analysis_df = df.copy()
    analysis_df["Sales Rep Bucket"] = analysis_df["Number of sales reps"].map(sales_rep_bucket)
    bucket_order = ["1 rep", "2-5 reps", "6-10 reps", "11-20 reps", "21+ reps"]

    st.markdown('<div class="section-title">Where do the largest sales teams come from?</div>', unsafe_allow_html=True)
    for bucket in bucket_order:
        bucket_deals = int(analysis_df["Sales Rep Bucket"].eq(bucket).sum())
        if bucket_deals == 0:
            continue

        with st.expander(f"{bucket} ({bucket_deals:,} deals)", expanded=bucket in ["11-20 reps", "21+ reps"]):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("##### Top Countries")
                st.dataframe(
                    format_diagnostic_table(top_segment_table(analysis_df, "Client country", bucket)),
                    hide_index=True,
                    use_container_width=True,
                )
            with col2:
                st.markdown("##### Top CRM Categories")
                st.dataframe(
                    format_diagnostic_table(top_segment_table(analysis_df, "Client CRM", bucket)),
                    hide_index=True,
                    use_container_width=True,
                )
            with col3:
                st.markdown("##### Top Lead Sources")
                st.dataframe(
                    format_diagnostic_table(top_segment_table(analysis_df, "Source", bucket)),
                    hide_index=True,
                    use_container_width=True,
                )

    insight_text = "<br><br>".join(diagnostic_concentration_insights(analysis_df))
    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">Diagnostic Note</div>
            <div class="insight-body">{insight_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def basic_win_rate_summary(df: pd.DataFrame) -> pd.DataFrame:
    closed_deals = int(df["Is Closed"].sum())
    won_deals = int(df["Is Won"].sum())
    lost_deals = int(df["Is Lost"].sum())
    win_rate = won_deals / closed_deals if closed_deals else 0
    return pd.DataFrame(
        [
            {
                "Total Deals": len(df),
                "Closed Deals": closed_deals,
                "Won Deals": won_deals,
                "Lost Deals": lost_deals,
                "Win Rate": f"{win_rate:.1%}",
            }
        ]
    )


def stage_distribution_table(df: pd.DataFrame) -> pd.DataFrame:
    total = len(df)
    summary = count_summary(df, "Stage", name="Deal count").rename(columns={"Stage": "Stage"})
    summary["% of Demo Request deals"] = summary["Deal count"].map(lambda value: value / total if total else 0)
    summary["% of Demo Request deals"] = summary["% of Demo Request deals"].map(lambda value: f"{value:.1%}")
    return summary


def loss_reason_table(df: pd.DataFrame) -> pd.DataFrame:
    lost_df = df[df["Is Lost"]].copy()
    total_lost = len(lost_df)
    if lost_df.empty:
        return pd.DataFrame(columns=["Loss Reason", "Count", "% of lost Demo Request deals"])

    lost_df["Loss Reason"] = lost_df["Loss reason description"].replace("", "Missing loss reason")
    summary = (
        lost_df.groupby("Loss Reason", dropna=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    summary["% of lost Demo Request deals"] = summary["Count"].map(lambda value: f"{(value / total_lost):.1%}" if total_lost else "0.0%")
    return summary


def dimension_win_rate_table(df: pd.DataFrame, dimension: str, label: str) -> pd.DataFrame:
    summary = (
        df.groupby(dimension, dropna=False)
        .agg(
            **{
                "Total Deals": ("Stage", "size"),
                "Closed Deals": ("Is Closed", "sum"),
                "Won Deals": ("Is Won", "sum"),
            }
        )
        .reset_index()
        .rename(columns={dimension: label})
    )
    summary["Win Rate"] = summary["Won Deals"] / summary["Closed Deals"].replace(0, pd.NA)
    summary["Win Rate"] = summary["Win Rate"].fillna(0)
    summary = summary.sort_values(["Total Deals", "Closed Deals"], ascending=[False, False])
    summary["Win Rate"] = summary["Win Rate"].map(lambda value: f"{value:.1%}")
    return summary


def demo_request_breakdowns(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    demo_df = df[df["Source"].eq("Demo Request")].copy()
    demo_df["Sales Rep bucket"] = demo_df["Number of sales reps"].map(sales_rep_bucket)
    return {
        "summary": basic_win_rate_summary(demo_df),
        "stage": stage_distribution_table(demo_df),
        "loss": loss_reason_table(demo_df),
        "country": dimension_win_rate_table(demo_df, "Client country", "Country"),
        "crm": dimension_win_rate_table(demo_df, "Client CRM", "CRM"),
        "sales_team": dimension_win_rate_table(demo_df, "Sales Rep bucket", "Sales Rep bucket"),
        "budget": dimension_win_rate_table(demo_df, "PPC budget USD", "Budget segment"),
    }


def demo_request_interpretation(tables: dict[str, pd.DataFrame]) -> str:
    summary = tables["summary"].iloc[0]
    country = tables["country"]
    crm = tables["crm"]

    top_country = country.iloc[0] if not country.empty else None
    top_crm = crm.iloc[0] if not crm.empty else None

    parts = [
        f"Demo Request contributes {summary['Total Deals']:,} total deals, {summary['Closed Deals']:,} closed deals, and a {summary['Win Rate']} win rate.",
    ]
    if top_country is not None and top_crm is not None:
        parts.append(
            f"The largest visible country and CRM concentrations are {top_country['Country']} "
            f"({top_country['Total Deals']:,} deals) and {top_crm['CRM']} ({top_crm['Total Deals']:,} deals)."
        )

    return " ".join(parts)


def render_demo_request_analysis_section(df: pd.DataFrame) -> None:
    tables = demo_request_breakdowns(df)
    demo_df = df[df["Source"].eq("Demo Request")].copy()
    demo_loss_categories = loss_category_summary(demo_df).head(5)

    st.markdown('<div class="section-title">Demo Request Performance Analysis</div>', unsafe_allow_html=True)
    st.dataframe(tables["summary"], hide_index=True, use_container_width=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("##### Top Loss Categories")
        st.dataframe(format_loss_category_summary(demo_loss_categories), hide_index=True, use_container_width=True)
    with col2:
        st.markdown("##### Top Countries")
        st.dataframe(tables["country"].head(5), hide_index=True, use_container_width=True)
    with col3:
        st.markdown("##### Top CRM Categories")
        st.dataframe(tables["crm"].head(5), hide_index=True, use_container_width=True)

    top_loss = demo_loss_categories.iloc[0] if not demo_loss_categories.empty else None
    top_loss_text = (
        f"The largest normalized Demo Request loss category is {top_loss['Loss Category']} "
        f"with {int(top_loss['Lost Deals']):,} lost deals ({top_loss['% of All Closed Lost Deals']:.1%} of Demo Request losses)."
        if top_loss is not None
        else "There are no Closed Lost Demo Request deals in the dataset."
    )
    finding_card(
        "Finding: Demo Request performance is a source-level conversion issue, not just a chart outlier.",
        "We reviewed Demo Request deals by closed volume, normalized loss category, country mix, and CRM mix.",
        f"{demo_request_interpretation(tables)} {top_loss_text}",
        "Demo Request should be reviewed as a lead-quality and follow-up workflow problem. The next operating review should inspect representative lost Demo Request deals from the top loss category and the highest-volume country/CRM combinations.",
    )


def normalize_loss_category(reason: object) -> str:
    text = normalize_text(reason).casefold()
    if not text:
        return "Other"
    if "competitor" in text:
        return "Competitive Pressure"
    if any(keyword in text for keyword in ["budget", "price", "profitable", "expensive", "cost"]):
        return "Budget / Price"
    if any(keyword in text for keyword in ["communicat", "no response", "lost contact", "not responding", "no answer"]):
        return "Lost Contact / Engagement"
    if any(keyword in text for keyword in ["no need", "not relevant", "refused", "no responsible", "not interested", "decided that it was not profitable"]):
        return "No Need / No Fit"
    if any(keyword in text for keyword in ["telecom", "technical", "integration", "infrastructure", "limitations", "numbers", "voice traffic"]):
        return "Technical / Infrastructure"
    return "Other"


def loss_category_summary(df: pd.DataFrame) -> pd.DataFrame:
    lost_df = df[df["Is Lost"]].copy()
    total_lost = len(lost_df)
    if lost_df.empty:
        return pd.DataFrame(columns=["Loss Category", "Lost Deals", "% of All Closed Lost Deals"])

    lost_df["Loss Category"] = lost_df["Loss reason description"].map(normalize_loss_category)
    summary = (
        lost_df.groupby("Loss Category", dropna=False)
        .size()
        .reset_index(name="Lost Deals")
        .sort_values("Lost Deals", ascending=False)
    )
    summary["% of All Closed Lost Deals"] = summary["Lost Deals"] / total_lost
    return summary


def format_loss_category_summary(summary: pd.DataFrame) -> pd.DataFrame:
    formatted = summary.copy()
    formatted["% of All Closed Lost Deals"] = formatted["% of All Closed Lost Deals"].map(lambda value: f"{value:.1%}")
    return formatted


def plot_loss_categories(summary: pd.DataFrame) -> go.Figure:
    if summary.empty:
        fig = go.Figure()
        fig.update_layout(
            title="Why Deals Are Lost",
            annotations=[
                {
                    "text": "No Closed Lost deals for the selected filters.",
                    "xref": "paper",
                    "yref": "paper",
                    "x": 0.5,
                    "y": 0.5,
                    "showarrow": False,
                    "font": {"size": 14, "color": COLORS["muted"]},
                    "align": "center",
                }
            ],
            xaxis={"visible": False},
            yaxis={"visible": False},
        )
        return style_chart(fig)

    chart_df = summary.copy()
    chart_df["Share label"] = chart_df["% of All Closed Lost Deals"].map(lambda value: f"{value:.1%}")
    fig = px.bar(
        chart_df,
        x="Lost Deals",
        y="Loss Category",
        orientation="h",
        text="Lost Deals",
        color="Lost Deals",
        color_continuous_scale=[[0, "#1E3A8A"], [0.55, "#38BDF8"], [1, "#F43F5E"]],
        hover_data={
            "Lost Deals": True,
            "% of All Closed Lost Deals": ":.1%",
            "Share label": False,
            "Loss Category": False,
        },
        labels={"Lost Deals": "Closed Lost deals", "Loss Category": ""},
        title="Why Deals Are Lost",
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"})
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
    return style_chart(fig)


def loss_category_interpretation(summary: pd.DataFrame) -> str:
    if summary.empty:
        return "There are no Closed Lost deals in the dataset, so loss themes cannot be evaluated."

    top = summary.iloc[0]
    top_category = top["Loss Category"]
    top_share = top["% of All Closed Lost Deals"]
    second = summary.iloc[1] if len(summary) > 1 else None

    second_text = (
        f" The next largest category is {second['Loss Category']} at {second['% of All Closed Lost Deals']:.1%}."
        if second is not None
        else ""
    )

    return (
        f"The largest classified driver of Closed Lost opportunities is {top_category}, representing {top_share:.1%} of lost deals."
        f"{second_text} The highest-count categories define the first cleanup priorities for sales enablement, competitive positioning, and CRM loss-reason hygiene."
    )


def render_loss_reason_analysis_section(df: pd.DataFrame) -> None:
    summary = loss_category_summary(df)

    st.markdown('<div class="section-title">Why Deals Are Lost</div>', unsafe_allow_html=True)
    other_row = summary[summary["Loss Category"].eq("Other")]
    if not other_row.empty and float(other_row.iloc[0]["% of All Closed Lost Deals"]) > 0.2:
        st.warning("Loss-reason classification quality is limited because a large share of lost deals are grouped into 'Other'. Consider standardizing loss-reason taxonomy.")

    col1, col2 = st.columns([0.95, 1.05])
    with col1:
        st.dataframe(format_loss_category_summary(summary), hide_index=True, use_container_width=True)
    with col2:
        st.plotly_chart(plot_loss_categories(summary), use_container_width=True)

    top = summary.iloc[0] if not summary.empty else None
    second = summary.iloc[1] if len(summary) > 1 else None
    evidence = loss_category_interpretation(summary)
    if top is not None and second is not None:
        conclusion = (
            f"{top['Loss Category']} is the primary loss theme and {second['Loss Category']} is the secondary theme. "
            "The business response should prioritize the largest classified theme first while improving taxonomy quality where Other remains high."
        )
    elif top is not None:
        conclusion = (
            f"{top['Loss Category']} is the primary classified loss theme. "
            "The business response should start with representative deal review inside that category."
        )
    else:
        conclusion = "Loss-category conclusions are unavailable because there are no Closed Lost records."

    finding_card(
        "Finding: Closed Lost reasons point to a focused set of operating priorities.",
        "We normalized detailed loss reasons into executive categories and ranked them by lost-deal count.",
        evidence,
        conclusion,
    )


def budget_price_loss_df(df: pd.DataFrame) -> pd.DataFrame:
    lost_df = df[df["Is Lost"]].copy()
    lost_df["Loss Category"] = lost_df["Loss reason description"].map(normalize_loss_category)
    return lost_df[lost_df["Loss Category"].eq("Budget / Price")].copy()


def budget_price_breakdown_table(df: pd.DataFrame, dimension: str, label: str) -> pd.DataFrame:
    budget_loss_df = budget_price_loss_df(df)
    total_budget_losses = len(budget_loss_df)
    if budget_loss_df.empty:
        return pd.DataFrame(columns=[label, "Lost deals", "% of Budget/Price losses"])

    if dimension == "Sales Rep bucket":
        budget_loss_df["Sales Rep bucket"] = budget_loss_df["Number of sales reps"].map(sales_rep_bucket)

    summary = (
        budget_loss_df.groupby(dimension, dropna=False)
        .size()
        .reset_index(name="Lost deals")
        .rename(columns={dimension: label})
        .sort_values("Lost deals", ascending=False)
    )
    summary["% of Budget/Price losses"] = summary["Lost deals"].map(lambda value: f"{(value / total_budget_losses):.1%}")
    return summary


def budget_price_loss_rate_by_budget(df: pd.DataFrame) -> pd.DataFrame:
    lost_df = df[df["Is Lost"]].copy()
    if lost_df.empty:
        return pd.DataFrame(columns=["PPC budget segment", "Budget/Price losses", "All Closed Lost deals", "Budget/Price loss rate"])

    lost_df["Loss Category"] = lost_df["Loss reason description"].map(normalize_loss_category)
    summary = (
        lost_df.groupby("PPC budget USD", dropna=False)
        .agg(
            **{
                "Budget/Price losses": ("Loss Category", lambda values: int(values.eq("Budget / Price").sum())),
                "All Closed Lost deals": ("Loss Category", "size"),
            }
        )
        .reset_index()
        .rename(columns={"PPC budget USD": "PPC budget segment"})
    )
    summary["Budget/Price loss rate"] = summary["Budget/Price losses"] / summary["All Closed Lost deals"].replace(0, pd.NA)
    summary["Budget/Price loss rate"] = summary["Budget/Price loss rate"].fillna(0)
    summary = summary.sort_values("PPC budget segment", key=lambda series: series.map(ppc_budget_sort_key))
    summary["Budget/Price loss rate"] = summary["Budget/Price loss rate"].map(lambda value: f"{value:.1%}")
    return summary


def budget_objection_interpretation(df: pd.DataFrame, budget_breakdown: pd.DataFrame, rate_table: pd.DataFrame) -> str:
    budget_loss_df = budget_price_loss_df(df)
    if budget_loss_df.empty:
        return "There are no Budget / Price Closed Lost deals in the current filtered view."

    top_budget = budget_breakdown.iloc[0] if not budget_breakdown.empty else None
    rate_numeric = rate_table.copy()
    if not rate_numeric.empty:
        rate_numeric["Rate numeric"] = rate_numeric["Budget/Price loss rate"].str.rstrip("%").astype(float) / 100
        highest_rate = rate_numeric.sort_values(["Rate numeric", "All Closed Lost deals"], ascending=[False, False]).iloc[0]
    else:
        highest_rate = None

    parts = []
    if top_budget is not None:
        parts.append(
            f"Budget/Price losses are most concentrated in the {top_budget['PPC budget segment']} PPC budget segment ({top_budget['Lost deals']:,} losses, {top_budget['% of Budget/Price losses']} of budget-related losses)."
        )
    if highest_rate is not None:
        parts.append(
            f"The highest Budget/Price loss rate appears in {highest_rate['PPC budget segment']} at {highest_rate['Budget/Price loss rate']}, which may indicate where pricing sensitivity is strongest."
        )
    parts.append(
        "If budget objections are concentrated in lower-budget segments, this may suggest a customer-fit or packaging issue. If they appear across budget tiers, it could potentially indicate broader pricing, value communication, or competitive-positioning friction. This requires further investigation before making causal claims."
    )
    return " ".join(parts)


def render_budget_objection_analysis_section(df: pd.DataFrame) -> None:
    budget_breakdown = budget_price_breakdown_table(df, "PPC budget USD", "PPC budget segment")
    crm_breakdown = budget_price_breakdown_table(df, "Client CRM", "CRM")
    country_breakdown = budget_price_breakdown_table(df, "Client country", "Country")
    sales_team_breakdown = budget_price_breakdown_table(df, "Sales Rep bucket", "Sales rep bucket")
    rate_table = budget_price_loss_rate_by_budget(df)

    st.markdown('<div class="section-title">Budget Objection Analysis</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### PPC Budget Segment")
        st.dataframe(budget_breakdown, hide_index=True, use_container_width=True)
    with col2:
        st.markdown("##### CRM Category")
        st.dataframe(crm_breakdown, hide_index=True, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### Country")
        st.dataframe(country_breakdown, hide_index=True, use_container_width=True)
    with col2:
        st.markdown("##### Sales Team Size")
        st.dataframe(sales_team_breakdown, hide_index=True, use_container_width=True)

    st.markdown("##### Budget/Price Loss Rate by PPC Budget Segment")
    st.dataframe(rate_table, hide_index=True, use_container_width=True)

    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">Interpretation</div>
            <div class="insight-body">{budget_objection_interpretation(df, budget_breakdown, rate_table)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def loss_category_rate_by_crm_adoption(df: pd.DataFrame, category: str, loss_column_name: str, rate_column_name: str) -> pd.DataFrame:
    lost_df = df[df["Is Lost"]].copy()
    if lost_df.empty:
        return pd.DataFrame(columns=["Segment", "Total Closed Lost Deals", loss_column_name, rate_column_name])

    lost_df["CRM Adoption Segment"] = lost_df["Client CRM"].where(
        lost_df["Client CRM"].eq("No CRM"),
        "CRM Users",
    )
    lost_df["Loss Category"] = lost_df["Loss reason description"].map(normalize_loss_category)
    summary = (
        lost_df.groupby("CRM Adoption Segment", dropna=False)
        .agg(
            **{
                "Total Closed Lost Deals": ("Stage", "size"),
                loss_column_name: ("Loss Category", lambda values: int(values.eq(category).sum())),
            }
        )
        .reindex(["No CRM", "CRM Users"])
        .reset_index()
        .rename(columns={"CRM Adoption Segment": "Segment"})
    )
    summary[rate_column_name] = summary[loss_column_name] / summary["Total Closed Lost Deals"].replace(0, pd.NA)
    summary[rate_column_name] = summary[rate_column_name].fillna(0)
    return summary


def format_loss_rate_table(summary: pd.DataFrame, rate_column_name: str) -> pd.DataFrame:
    formatted = summary.copy()
    formatted[rate_column_name] = formatted[rate_column_name].map(lambda value: f"{value:.1%}")
    return formatted


def no_crm_pricing_validation_interpretation(budget_table: pd.DataFrame, competitive_table: pd.DataFrame) -> str:
    budget_lookup = budget_table.set_index("Segment")
    competitive_lookup = competitive_table.set_index("Segment")

    no_crm_budget_rate = budget_lookup.loc["No CRM", "Budget/Price Loss Rate"]
    crm_budget_rate = budget_lookup.loc["CRM Users", "Budget/Price Loss Rate"]
    no_crm_competitive_rate = competitive_lookup.loc["No CRM", "Competitive Loss Rate"]
    crm_competitive_rate = competitive_lookup.loc["CRM Users", "Competitive Loss Rate"]

    budget_text = (
        "No CRM businesses are more likely to cite Budget/Price concerns among Closed Lost deals, which may indicate value-perception or packaging friction."
        if no_crm_budget_rate > crm_budget_rate
        else "No CRM businesses are not more likely to cite Budget/Price concerns in the current filtered view."
    )
    competitive_text = (
        "CRM Users are more likely to be lost to Competitive Pressure, which suggests competitor comparisons may matter more for CRM-adopting companies."
        if crm_competitive_rate > no_crm_competitive_rate
        else "CRM Users are not more likely to be lost to Competitive Pressure in the current filtered view."
    )
    value_text = (
        "Taken together, this potentially suggests that the lower No CRM win rate may be driven more by perceived value, sales-process maturity, or fit than by budget size alone, but this requires further investigation."
        if no_crm_budget_rate > crm_budget_rate
        else "The evidence does not clearly isolate value perception as the main driver of the No CRM win-rate gap, and requires further investigation."
    )

    return (
        f"{budget_text} {competitive_text} {value_text} "
        f"Budget/Price loss rates are {no_crm_budget_rate:.1%} for No CRM vs {crm_budget_rate:.1%} for CRM Users; "
        f"Competitive Pressure loss rates are {no_crm_competitive_rate:.1%} vs {crm_competitive_rate:.1%}."
    )


def render_no_crm_pricing_validation_section(df: pd.DataFrame) -> None:
    budget_table = loss_category_rate_by_crm_adoption(
        df,
        "Budget / Price",
        "Budget/Price Losses",
        "Budget/Price Loss Rate",
    )
    competitive_table = loss_category_rate_by_crm_adoption(
        df,
        "Competitive Pressure",
        "Competitive Pressure Losses",
        "Competitive Loss Rate",
    )

    st.markdown('<div class="section-title">Is the Pricing Problem Actually a No CRM Problem?</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### Budget / Price Loss Rate")
        st.dataframe(format_loss_rate_table(budget_table, "Budget/Price Loss Rate"), hide_index=True, use_container_width=True)
    with col2:
        st.markdown("##### Competitive Pressure Loss Rate")
        st.dataframe(format_loss_rate_table(competitive_table, "Competitive Loss Rate"), hide_index=True, use_container_width=True)

    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">Interpretation</div>
            <div class="insight-body">{no_crm_pricing_validation_interpretation(budget_table, competitive_table)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_recommendations_section() -> None:
    st.markdown('<div class="section-title">Recommendations</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="recommendation-card">
            <div class="insight-body">
                <strong>1. Review positioning and messaging for businesses without CRM adoption.</strong><br>
                The No CRM segment is large and converts below CRM Users, so messaging around sales process organization, lead visibility, and operational structure is worth testing.<br><br>
                <strong>2. Investigate competitive-loss patterns and competitor differentiation.</strong><br>
                Competitive Pressure is a meaningful loss category, especially for CRM Users, and should be reviewed against competitor positioning and differentiation claims.<br><br>
                <strong>3. Standardize loss-reason taxonomy to reduce "Other" classifications and improve root-cause visibility.</strong><br>
                A large Other category limits diagnostic quality and should be reduced through cleaner CRM picklists or sales-process hygiene.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_overview(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-title">Executive Snapshot</div>', unsafe_allow_html=True)
    draw_metric_row(df)
    st.write("")
    render_context_strip(df)
    st.write("")

    left, right = st.columns([1.05, 0.95])
    with left:
        st.plotly_chart(plot_stage_distribution(df), use_container_width=True)
    with right:
        st.plotly_chart(plot_distribution(df, "Source", "Deal Count by Source"), use_container_width=True)


def has_small_sample(chart_data: pd.DataFrame, *, min_closed_deals: int) -> bool:
    if chart_data.empty:
        return False
    return bool((chart_data["closed_deals"] < min_closed_deals).any())


def warning_note(message: str) -> None:
    st.markdown(f'<div class="warning-note">{message}</div>', unsafe_allow_html=True)


def render_analytics(df: pd.DataFrame, filter_context: FilterContext) -> None:
    st.markdown('<div class="section-title">Conversion Analytics</div>', unsafe_allow_html=True)
    warning_note(
        "<strong>Methodology Note</strong><br>"
        "To estimate the reliability of win-rate comparisons, an approximate 95% margin of error was calculated using:<br><br>"
        "ME ≈ 1.96 × √(0.25 / n)<br><br>"
        "where ME = margin of error and n = number of closed deals. "
        "The value 0.25 represents the maximum possible variance for a conversion rate and therefore provides a conservative estimate of uncertainty. "
        "This approximation was used to evaluate how much statistical noise could exist within small segments and to justify minimum closed-deal thresholds in ranking views. "
        "Segments with very small sample sizes can produce extreme win rates (0% or 100%) that are driven by random variation rather than actual performance differences."
    )

    left, right = st.columns(2)
    is_country_drill_down = bool(filter_context.countries)
    is_crm_drill_down = bool(filter_context.crms)
    country_mode = "drill-down" if is_country_drill_down else ("breakdown" if is_crm_drill_down else "global")
    crm_mode = "drill-down" if is_crm_drill_down else ("breakdown" if is_country_drill_down else "global")
    country_chart_data = get_country_chart_data(df, filter_context.countries, filter_context.crms)
    crm_chart_data = get_crm_chart_data(df, filter_context.crms, filter_context.countries)

    with left:
        st.plotly_chart(
            plot_segment_win_rate(
                country_chart_data,
                "Win Rate by Country",
                global_mode=country_mode == "global",
            ),
            use_container_width=True,
        )
        if is_country_drill_down:
            if has_small_sample(country_chart_data, min_closed_deals=20):
                warning_note("Small sample size: selected country has fewer than 20 closed deals.")
            st.caption("Country threshold is not applied because a country drill-down is selected.")
        elif country_mode == "breakdown":
            if has_small_sample(country_chart_data, min_closed_deals=20):
                warning_note("Small sample size: one or more country segments have fewer than 20 closed deals.")
            st.caption("Country threshold is not applied because this chart is showing a breakdown within the selected CRM.")
        else:
            st.caption("Countries with fewer than 20 closed deals are excluded because small samples create unstable win-rate estimates.")

    with right:
        st.plotly_chart(
            plot_segment_win_rate(
                crm_chart_data,
                "Win Rate by CRM",
                global_mode=crm_mode == "global",
            ),
            use_container_width=True,
        )
        if is_crm_drill_down:
            if has_small_sample(crm_chart_data, min_closed_deals=10):
                warning_note("Small sample size: selected CRM has fewer than 10 closed deals.")
            st.caption("CRM threshold is not applied because a CRM drill-down is selected.")
        elif crm_mode == "breakdown":
            if has_small_sample(crm_chart_data, min_closed_deals=10):
                warning_note("Small sample size: one or more CRM segments have fewer than 10 closed deals.")
            st.caption("CRM threshold is not applied because this chart is showing a breakdown within the selected country.")
        else:
            st.caption("CRM categories with fewer than 10 closed deals are excluded because small samples create unstable win-rate estimates.")

    st.plotly_chart(plot_ppc_budget_win_rate(df), use_container_width=True)
    st.plotly_chart(plot_crm_ppc_heatmap(df, min_closed_deals=5), use_container_width=True)
    st.caption("Cells with fewer than 5 closed deals are hidden to reduce statistical noise.")


def render_charts(df: pd.DataFrame, filter_context: FilterContext | None = None) -> None:
    # Kept as a compatibility wrapper; it does not change dashboard logic.
    render_overview(df)
    render_analytics(df, filter_context or FilterContext(countries=[], crms=[]))


def build_quality_summary(df: pd.DataFrame) -> pd.DataFrame:
    checks = quality_checks(df)
    return pd.DataFrame(
        {
            "Check": [check.name for check in checks],
            "Issue count": [int(check.mask.sum()) for check in checks],
        }
    )


def visible_quality_summary(df: pd.DataFrame) -> pd.DataFrame:
    summary = build_quality_summary(df)
    return summary[summary["Issue count"] > 0].copy()


def build_quality_detail(df: pd.DataFrame) -> pd.DataFrame:
    detail_frames = []
    detail_columns = [
        "AQL date",
        "Source",
        "Client CRM",
        "Client country",
        "PPC budget USD",
        "Stage",
        "Loss reason description",
        "Closing Date",
    ]

    for check in quality_checks(df):
        rows = df.loc[check.mask, detail_columns].copy()
        if not rows.empty:
            rows.insert(0, "Issue category", check.name)
            detail_frames.append(rows)

    if not detail_frames:
        return pd.DataFrame(columns=["Issue category", *detail_columns])

    return pd.concat(detail_frames, ignore_index=True)


def render_quality_section(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-title">Data Quality Controls</div>', unsafe_allow_html=True)

    checks = quality_checks(df)
    full_summary = build_quality_summary(df)
    summary = visible_quality_summary(df)
    total_issues = int(summary["Issue count"].sum()) if not summary.empty else 0

    warning_note(
        "<strong>Methodology Note</strong><br>"
        "Five validation rules were evaluated. Three returned active issues and two returned zero issues, "
        "so only actionable issue categories are shown below."
    )

    validation_summary = full_summary.copy()
    validation_summary["Issue count"] = validation_summary["Issue count"].map(
        lambda count: f"{int(count):,} issue" if int(count) == 1 else f"{int(count):,} issues"
    )

    col1, col2 = st.columns([0.85, 2.15])
    with col1:
        st.markdown(
            f"""
            <div class="quality-card">
                <div class="insight-label">Total Issues Found</div>
                <div class="kpi-value">{total_issues:,}</div>
                <div class="insight-body">Across all validation checks in the current filter context.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col2:
        st.dataframe(validation_summary, hide_index=True, use_container_width=True)

    if summary.empty:
        st.success("No actionable data quality issues found in the current filter context.")
        return

    top_issue = summary.sort_values("Issue count", ascending=False).iloc[0]
    warning_note(f"<strong>Methodology Note</strong><br>Highest-impact cleanup area: {top_issue['Check']} ({int(top_issue['Issue count']):,} rows).")

    st.write("")
    selected_check = st.selectbox(
        "Review issue category",
        options=summary["Check"].tolist(),
    )
    selected_mask = next(check.mask for check in checks if check.name == selected_check)
    issue_rows = df.loc[
        selected_mask,
        [
            "AQL date",
            "Source",
            "Client CRM",
            "Client country",
            "PPC budget USD",
            "Stage",
            "Loss reason description",
            "Closing Date",
        ],
    ].copy()

    st.caption(f"{len(issue_rows):,} rows flagged for the selected category.")
    st.dataframe(issue_rows, hide_index=True, use_container_width=True)

    with st.expander("All detailed issues", expanded=False):
        st.dataframe(build_quality_detail(df), hide_index=True, use_container_width=True)


def insight_card(label: str, value: str, body: str, *, icon: str = "●", variant: str = "important") -> None:
    st.markdown(
        f"""
        <div class="insight-card {variant}">
            <div class="insight-label"><span>{icon}</span>{label}</div>
            <div class="insight-value">{value}</div>
            <div class="insight-body">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def executive_summary_card(title: str, metric: str, body: str, *, icon: str, variant: str) -> None:
    st.markdown(
        f"""
        <div class="exec-summary-card {variant}">
            <div class="exec-title">{icon} {title}</div>
            <div class="exec-metric">{metric}</div>
            <div class="exec-body">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_executive_summary_cards() -> None:
    st.markdown('<div class="section-title">Executive Summary</div>', unsafe_allow_html=True)
    cards = [
        ("Strongest acquisition channel", "Partner referrals", "Consistently strong performance across multiple sales-team segments.", "✅", "positive"),
        ("Weakest major source", "Demo Request", "High-intent source with only 15.5% win rate.", "⚠", "warning"),
        ("Largest measurable loss driver", "Budget / Price", "28.1% of lost deals are tied to budget or price objections.", "💰", "risk"),
        ("Highest-performing segment", "CRM Users", "CRM Users convert at 36.2% vs 29.3% for No CRM.", "🏆", "positive"),
        ("Main bottleneck", "Negotiations", "Negotiation-stage friction should be reviewed alongside loss reasons.", "📊", "important"),
        ("Funnel review priority", "Inbound Call", "Inbound Call handling deserves deeper review together with Demo Request.", "⚠", "warning"),
    ]
    for row_start in range(0, len(cards), 3):
        columns = st.columns(3)
        for column, (title, metric, body, icon, variant) in zip(columns, cards[row_start:row_start + 3]):
            with column:
                executive_summary_card(title, metric, body, icon=icon, variant=variant)


def render_insights_summary_cards() -> None:
    rows = [
        [
            (
                "CRM adoption is associated with higher conversion rates",
                "36.2% vs 29.3%",
                "CRM Users convert at 36.2%, compared with 29.3% for No CRM. CRM adoption is best read as a maturity signal, not a standalone cause.",
                "✅",
                "positive",
            ),
            (
                "Partner referrals are the strongest acquisition channel",
                "Strongest source",
                "Partner referrals show consistently strong performance across multiple sales-team segments, making this channel a priority for source-quality review and scaling.",
                "✅",
                "positive",
            ),
            (
                "Demo Request is underperforming",
                "15.5% win rate",
                "Demo Request is a high-intent source, but conversion is weak. The issue is likely in source quality, qualification, routing, or follow-up workflow.",
                "⚠",
                "warning",
            ),
        ],
        [
            (
                "Budget / Price and Competitive Pressure are the largest measurable loss drivers",
                "28.1% and 17.2%",
                "Budget / Price accounts for 28.1% of classified losses, while Competitive Pressure accounts for 17.2%. These are the clearest measurable loss themes.",
                "●",
                "risk",
            ),
            (
                "Loss-reason reporting quality is limited",
                "Other = 35.3%",
                '"Other" represents 35.3% of all lost deals, limiting root-cause visibility and weakening management interpretation of loss patterns.',
                "●",
                "risk",
            ),
            (
                "Pipeline is concentrated in Ukraine, Poland, and Kazakhstan",
                "Top country cluster",
                "Pipeline volume is concentrated in Ukraine, Poland, and Kazakhstan. Country mix should be considered when interpreting source, CRM, and win-rate patterns.",
                "●",
                "important",
            ),
        ],
    ]

    for row in rows:
        columns = st.columns(3)
        for column, (headline, metric, explanation, icon, variant) in zip(columns, row):
            with column:
                insight_card(headline, metric, explanation, icon=icon, variant=variant)


def investigation_card(
    title: str,
    finding: str,
    evidence: str,
    hypothesis: str,
    recommendation: str,
    *,
    variant: str = "important",
) -> None:
    st.markdown(
        f"""
        <div class="finding-card {variant}">
            <div class="finding-title">🔎 {title}</div>
            <div class="investigation-grid">
                <div class="investigation-step {variant}">
                    <div class="step-label">Finding</div>
                    <div class="step-body">{finding}</div>
                </div>
                <div class="investigation-step evidence">
                    <div class="step-label">Evidence</div>
                    <div class="step-body">{evidence}</div>
                </div>
                <div class="investigation-step hypothesis">
                    <div class="step-label">💡 Hypothesis</div>
                    <div class="step-body">{hypothesis}</div>
                </div>
                <div class="investigation-step recommendation">
                    <div class="step-label">Recommendation</div>
                    <div class="step-body">{recommendation}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def report_investigation_card(title: str, finding: str, evidence: str, hypothesis: str, recommendation: str) -> None:
    st.markdown(
        f"""
        <div class="report-card">
            <div class="report-title">{title}</div>
            <div class="report-section">
                <div class="report-label">Finding</div>
                <div class="report-body">{finding}</div>
            </div>
            <div class="report-section">
                <div class="report-label">Evidence</div>
                <div class="report-body">{evidence}</div>
            </div>
            <div class="report-section">
                <div class="report-label">Hypothesis</div>
                <div class="report-body">{hypothesis}</div>
            </div>
            <div class="report-section">
                <div class="report-label">Recommendation</div>
                <div class="report-body">{recommendation}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def evidence_snapshot(items: list[tuple[str, str]]) -> None:
    rows_html = "".join(
        f"""
        <div class="evidence-snapshot-row">
            <div class="evidence-snapshot-label">{label}</div>
            <div class="evidence-snapshot-value">{value}</div>
        </div>
        """
        for label, value in items
    )
    st.markdown(
        f"""
        <div class="evidence-snapshot">
            <div class="evidence-snapshot-title">Evidence Snapshot</div>
            {rows_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_investigation_sections() -> None:
    st.markdown('<div class="section-title">INVESTIGATION FINDINGS</div>', unsafe_allow_html=True)
    investigations = [
        (
            "CRM ADOPTION VS BUSINESS MATURITY",
            "Businesses using CRM systems convert significantly better than businesses without CRM adoption (36.2% vs 29.3%).",
            "The conversion gap remains visible even after reviewing PPC budget distribution, sales-team size, and loss-reason patterns.<br><br>While CRM users operate with somewhat higher average PPC budgets ($628 vs $448), pricing objections occur at nearly identical rates (28.2% vs 28.0%). This suggests that budget differences alone do not explain the performance gap.",
            "The evidence suggests that CRM adoption should not be interpreted as the direct cause of higher conversion rates.<br><br>Instead, CRM usage appears to be a signal of broader operational maturity. Companies that invest in CRM systems are also more likely to have structured sales processes, better pipeline visibility, clearer lead ownership, stronger follow-up discipline, and more consistent performance management.",
            "For businesses currently operating without CRM systems, the conversation should not focus exclusively on technology adoption.<br><br>The stronger value proposition is operational visibility: improving process control, accountability, pipeline transparency, and reducing missed opportunities through better sales management practices.",
            [
                ("CRM Users win rate", "36.2%"),
                ("No CRM win rate", "29.3%"),
                ("Average PPC budget", "$628 vs $448"),
                ("Budget / Price objections", "28.2% vs 28.0%"),
            ],
        ),
        (
            "DEMO REQUEST PERFORMANCE",
            "Demo Request appears to be one of the weakest-performing acquisition sources despite representing a high-intent buyer action.",
            "The source generated 77 opportunities but converted at only 15.5%.<br><br>Given that requesting a demo is typically considered a strong buying signal, conversion performance at this level is notably weaker than expected. A substantial share of these opportunities ultimately ended in Closed Lost status rather than progressing through the funnel successfully.",
            "The issue is unlikely to be demand volume.<br><br>The more probable explanation is friction somewhere between lead capture and deal progression. Potential causes include qualification criteria, routing logic, ownership assignment, response speed, follow-up execution, or broader lead-quality issues.<br><br>In other words, the problem may not be attracting interest, but successfully converting that interest into revenue.",
            "Demo Request should be reviewed as a dedicated funnel rather than being grouped together with other acquisition sources.<br><br>The next investigation should map the full journey from form submission through qualification, ownership assignment, sales engagement, and final outcome to identify where the largest conversion drop-offs occur.",
            [
                ("Demo Request opportunities", "77"),
                ("Demo Request win rate", "15.5%"),
                ("Expected signal", "High intent"),
                ("Observed issue", "Weak conversion"),
            ],
        ),
        (
            "WHY DEALS ARE LOST",
            "The largest measurable drivers of lost opportunities are Budget / Price concerns and Competitive Pressure.",
            "Budget / Price accounts for 28.1% of categorized losses, while Competitive Pressure accounts for 17.2%.<br><br>Together, these categories represent nearly half of all classified lost opportunities and are therefore the strongest actionable themes currently visible in the dataset.",
            "These results should not be interpreted as evidence that pricing alone is the primary issue.<br><br>In many sales environments, \"price\" is often a proxy for perceived value. Prospects may understand the cost of the solution without fully understanding the business value it delivers.<br><br>Similarly, competitive losses may indicate gaps in differentiation, positioning, or objection handling rather than purely stronger competitor offerings.",
            "Further investigation should focus on value communication, competitive positioning, pricing conversations, and objection-handling practices.<br><br>A more structured competitor-tracking process would also improve visibility into recurring competitive threats and messaging weaknesses.",
            [
                ("Budget / Price losses", "28.1%"),
                ("Competitive Pressure losses", "17.2%"),
                ("Combined visible loss themes", "45.3%"),
                ("Primary business signal", "Value and differentiation risk"),
            ],
        ),
        (
            "LOSS REASON DATA QUALITY",
            "The current loss-reason taxonomy limits the quality of root-cause analysis.",
            "The category \"Other\" represents 35.3% of all lost opportunities, making it the largest loss category in the dataset.<br><br>As a result, a substantial portion of lost deals cannot be meaningfully analyzed or linked to specific operational issues.",
            "The size of the \"Other\" category suggests that the existing classification framework does not provide sufficient granularity for consistent reporting.<br><br>Important loss patterns may currently be hidden inside a catch-all bucket, reducing management visibility and making improvement efforts less targeted.",
            "Reduce the use of generic loss reasons and introduce more structured classification standards.<br><br>Improving loss-reason quality would significantly increase the reliability and actionability of future RevOps investigations.",
            [
                ("Other category share", "35.3%"),
                ("Relative size", "Largest loss category"),
                ("Diagnostic impact", "Root causes are hidden"),
            ],
        ),
    ]

    for title, finding, evidence, hypothesis, recommendation, snapshot_items in investigations:
        report_investigation_card(
            title,
            finding,
            evidence,
            hypothesis,
            recommendation,
        )
        evidence_snapshot(snapshot_items)

    st.markdown('<div class="section-title">FINAL REVOPS CONCLUSION</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="report-card">
            <div class="report-body">
                This investigation suggests that operational maturity is a stronger predictor of conversion performance than budget size alone.
                <br><br>
                Businesses using CRM systems consistently outperform businesses without CRM adoption, while pricing objections occur at nearly identical rates across both groups. This indicates that conversion outcomes are more likely influenced by process quality, pipeline visibility, accountability, and sales discipline than by budget constraints.
                <br><br>
                At the acquisition level, Demo Request underperforms despite being a high-intent source, indicating potential issues in qualification, routing, ownership, or follow-up execution.
                <br><br>
                The largest measurable loss drivers are Budget / Price and Competitive Pressure. However, loss analysis is currently constrained by the large volume of opportunities classified as "Other", limiting visibility into true root causes.
                <br><br>
                Based on the evidence available, the highest-priority improvement opportunities are:
                <ol>
                    <li>Improve loss-reason classification quality.</li>
                    <li>Investigate Demo Request funnel performance.</li>
                    <li>Strengthen messaging around operational visibility and CRM-driven sales maturity.</li>
                    <li>Improve competitive differentiation and value communication during the sales process.</li>
                </ol>
                The findings presented here should be interpreted as directional business evidence rather than proof of causation. However, the consistency of the observed patterns is sufficient to justify further operational review and targeted investigation.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_insights(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-title">RevOps Insights Summary</div>', unsafe_allow_html=True)
    render_back_to_dashboard_button("insights")
    st.caption("Insights are based on the full normalized dataset and are not recalculated by sidebar filters.")
    render_executive_summary_cards()
    render_insights_summary_cards()
    render_investigation_sections()


def render_back_to_dashboard_button(key_suffix: str) -> None:
    if st.button("← Back to Dashboard", key=f"back_to_dashboard_{key_suffix}"):
        st.session_state["static_page"] = None
        st.rerun()


def render_tool_card(tool: str, usage: str, variant: str = "important") -> None:
    insight_card(tool, "Used for", usage, icon="AI", variant=variant)


def render_ai_usage_page() -> None:
    st.markdown('<div class="section-title">AI Usage Report</div>', unsafe_allow_html=True)
    render_back_to_dashboard_button("ai_usage")
    st.caption("This page is static and documents how AI tools supported dashboard development and business analysis.")

    st.markdown(
        """
        <div class="finding-card positive">
            <div class="finding-title">Purpose</div>
            <div class="step-body">
                AI tools were used to accelerate dashboard development, data exploration, visualization design, and iterative refinement of analytical outputs.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Tools Used</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="finding-card important">
            <div class="finding-title">Development and analysis tools</div>
            <div class="step-body">
                <ol>
                    <li><strong>ChatGPT</strong> — Used for analytical brainstorming and investigation design. Helped formulate business hypotheses and validate analytical logic. Assisted with interpretation of findings and recommendation drafting.</li>
                    <li><strong>Claude Code (Codex)</strong> — Used for dashboard implementation and iterative development. Generated Streamlit components, layouts, charts, styling, and page structures. Assisted with debugging, refactoring, and UI improvements.</li>
                    <li><strong>Streamlit</strong> — Used as the application framework for dashboard delivery.</li>
                    <li><strong>Pandas</strong> — Used for data transformation, normalization, aggregation, validation checks, and KPI calculations.</li>
                    <li><strong>Plotly</strong> — Used for interactive charts and visualizations.</li>
                </ol>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Prompts That Were Most Effective</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="finding-card important">
            <div class="finding-title">Business-question-oriented prompts</div>
            <div class="step-body">
                The most effective prompts were business-question-oriented rather than visualization-oriented.
                <br><br>
                Examples:
                <ul>
                    <li>"Compare CRM and No CRM companies while controlling for budget effects."</li>
                    <li>"Group loss reasons into broader business categories and identify the largest measurable loss drivers."</li>
                    <li>"Determine whether CRM adoption is acting as a maturity signal or a direct conversion driver."</li>
                    <li>"Identify lead sources that underperform despite high intent."</li>
                    <li>"Design visualizations that support investigation rather than simply displaying metrics."</li>
                </ul>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">What Was Generated by AI</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="finding-card positive">
            <div class="finding-title">AI-generated assets</div>
            <div class="step-body">
                AI generated:
                <ul>
                    <li>Initial Streamlit application structure.</li>
                    <li>Dashboard layouts and navigation.</li>
                    <li>Data-quality validation logic.</li>
                    <li>Chart implementations and styling suggestions.</li>
                    <li>Investigation framework drafts.</li>
                    <li>Iterative UI refinements.</li>
                    <li>Supporting documentation drafts.</li>
                </ul>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">What Was Done Manually</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="finding-card important">
            <div class="finding-title">Human-led analysis and decisions</div>
            <div class="step-body">
                The following work was performed manually:
                <ul>
                    <li>Defining the business questions.</li>
                    <li>Designing investigation paths.</li>
                    <li>Deciding which metrics mattered.</li>
                    <li>Evaluating hypotheses.</li>
                    <li>Distinguishing correlation from causation.</li>
                    <li>Prioritizing findings.</li>
                    <li>Writing final business conclusions.</li>
                    <li>Deciding which insights should be surfaced in the dashboard.</li>
                </ul>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Challenges Encountered</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="finding-card warning">
            <div class="finding-title">Translating business questions into dashboard logic</div>
            <div class="step-body">
                The primary challenge was not coding.
                <br><br>
                The most difficult part was translating business questions into analytical workflows and dashboard components that supported investigation rather than simply reporting metrics.
                <br><br>
                Multiple technically correct visualizations were discarded because they did not answer the actual business question.
                <br><br>
                Several iterations were required to align dashboard outputs with the investigation logic and decision-making process.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Final Reflection</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="recommendation-card">
            <div class="finding-title">Final reflection</div>
            <div class="step-body">
                AI significantly accelerated development speed and implementation.
                <br><br>
                However, interpretation, prioritization, hypothesis testing, and final business conclusions remained human-driven.
                <br><br>
                The highest-value contribution during the project was deciding which questions should be asked of the data before deciding which charts should be built.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    inject_theme()
    st.title(APP_TITLE)
    st.markdown(
        '<div class="dashboard-subtitle">Executive view of pipeline quality, conversion performance, and revenue operations signals.</div>',
        unsafe_allow_html=True,
    )

    st.sidebar.markdown("### Revenue Ops")
    uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx", "xls"])
    uploaded_bytes = uploaded_file.getvalue() if uploaded_file else None
    st.session_state.setdefault("static_page", None)

    try:
        df = load_data(uploaded_bytes)
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    st.sidebar.divider()
    if st.sidebar.button("Insights", use_container_width=True):
        st.session_state["static_page"] = "insights"
    if st.sidebar.button("AI Usage Report", use_container_width=True):
        st.session_state["static_page"] = "ai_usage"

    filtered, filter_context = apply_filters(df)

    if st.session_state["static_page"] == "insights":
        render_insights(df)
        return
    if st.session_state["static_page"] == "ai_usage":
        render_ai_usage_page()
        return

    if filtered.empty:
        st.warning("No deals match the selected filters.")
        st.stop()

    overview_tab, analytics_tab, quality_tab = st.tabs(["Overview", "Analytics", "Data Quality"])

    with overview_tab:
        render_overview(filtered)

    with analytics_tab:
        render_analytics(filtered, filter_context)

    with quality_tab:
        render_quality_section(filtered)


if __name__ == "__main__":
    main()
