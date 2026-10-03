# DATAVIS v5 ULTIMATE — data visualizer
# to run -   streamlit run app.py

import io
import json
import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    from langchain_anthropic import ChatAnthropic
    from langchain_core.messages import HumanMessage
    AI_SUPPORT = True
except ImportError:
    AI_SUPPORT = False

try:
    import openpyxl  # noqa: F401 — needed by pandas for .xlsx engine + report export
    EXCEL_EXPORT_SUPPORT = True
except ImportError:
    EXCEL_EXPORT_SUPPORT = False


# ============================================================
# CONSTANTS
# ============================================================

LARGE_DATASET_THRESHOLD = 15   # rows above this trigger the stacked, full-width chart layout
HUGE_DATASET_THRESHOLD = 300   # rows above this trigger sampling/pagination for tables

PALETTES = {
    "Vivid":        px.colors.qualitative.Vivid,
    "Bold":         px.colors.qualitative.Bold,
    "Pastel":       px.colors.qualitative.Pastel,
    "Prism":        px.colors.qualitative.Prism,
    "Safe":         px.colors.qualitative.Safe,
    "Sunset (seq)": px.colors.sequential.Sunset,
    "Teal (seq)":   px.colors.sequential.Teal,
    "Plasma (seq)": px.colors.sequential.Plasma,
}

CONTINUOUS_SCALES = {
    "Viridis": "Viridis", "Plasma": "Plasma", "Turbo": "Turbo",
    "Sunset": "Sunset", "Teal": "Teal", "Bluered": "Bluered", "Rainbow": "Rainbow",
}

CHART_TYPES = [
    "Bar", "Line", "Area", "Scatter", "Pie", "Box", "Violin", "Histogram", "Treemap"
]

THEME_PRESETS = {
    "Midnight Violet": {"bg1": "#0e0e17", "bg2": "#161625", "accent2": "#18C6C6"},
    "Deep Ocean":       {"bg1": "#061019", "bg2": "#0d2436", "accent2": "#38f5b0"},
    "Charcoal Rose":    {"bg1": "#161116", "bg2": "#241a24", "accent2": "#ff6ec7"},
    "Pure Light":       {"bg1": "#f7f7fb", "bg2": "#ffffff", "accent2": "#7C4DFF"},
    "Warm Paper":       {"bg1": "#fbf6ee", "bg2": "#fffdf8", "accent2": "#e08e45"},
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def anonymize_names(df, name_col):
    """Replaces every value in name_col with a generic placeholder ('Entity 1', 'Entity 2', ...)
    so real names never have to leave the app. Returns (anonymized_df, real_to_placeholder_map)."""
    anon_df = df.copy()
    unique_names = anon_df[name_col].astype(str).unique().tolist()
    mapping = {name: f"Entity {i + 1}" for i, name in enumerate(unique_names)}
    anon_df[name_col] = anon_df[name_col].astype(str).map(mapping)
    return anon_df, mapping


def get_ai_summary(df, api_key, name_col=None, include_raw_rows=False):
    """Sends dataset stats (and, only if explicitly opted into, row-level data) to Claude
    and asks for a plain-language analysis.

    Privacy: row-level data is never sent unless include_raw_rows=True, and even then, if a
    name_col is provided, names are replaced with anonymous placeholders ('Entity 1', 'Entity 2', ...)
    before anything leaves the app. The response is translated back to real names afterward so it's
    still readable locally.
    """
    if not AI_SUPPORT:
        return None, "langchain_anthropic is not installed in this environment."
    try:
        llm = ChatAnthropic(
            model="claude-sonnet-4-5-20250929",
            anthropic_api_key=api_key,
            max_tokens=1024
        )
        summary_stats = df.describe().to_string()
        columns_info = ", ".join(df.columns)

        mapping = {}
        raw_data_section = ""
        if include_raw_rows:
            if name_col:
                anon_df, mapping = anonymize_names(df, name_col)
                raw_data_section = f"\nRaw data (names replaced with anonymous placeholders):\n{anon_df.to_string()}\n"
            else:
                raw_data_section = f"\nRaw data:\n{df.to_string()}\n"

        prompt = f"""
You are a data analyst. You will be given a dataset with these columns: {columns_info}

Summary statistics:
{summary_stats}
{raw_data_section}
First, briefly identify what kind of data this appears to be (e.g. academic performance, business/sales, health/fitness, or another domain) based on the column names and values.

Then, based on that context, provide:
1. Which rows (entities) perform best and which perform worst overall, using the most relevant numeric columns as the basis for "performance"
2. Specific columns/metrics where the weaker entities fall short
3. Practical, actionable advice for improving the weaker entities' outcomes

{"Refer to entities using whatever placeholder labels appear in the raw data (e.g. 'Entity 3') since real names were not shared with you." if include_raw_rows and name_col else "No row-level data was shared with you, so keep the analysis at the level of columns/metrics rather than naming individual entities."}
Keep your analysis grounded strictly in the data provided — do not invent figures or assume information not present in the table.
"""
        response = llm.invoke([HumanMessage(content=prompt)])
        text = response.content
        # Translate anonymous placeholders back to real names for local display only —
        # the substitution happens after the API call, so Claude itself never saw the real names.
        if mapping:
            reverse_map = {v: k for k, v in mapping.items()}
            for placeholder, real_name in reverse_map.items():
                text = text.replace(placeholder, real_name)
        return text, None
    except Exception as e:
        return None, str(e)


def extract_pdf_tables(file):
    """Extracts every table found in a PDF using pdfplumber. Returns a list of
    (dataframe, page_number, table_number_on_page) tuples. First row of each
    extracted table is used as the header."""
    file.seek(0)
    results = []
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for t_num, table in enumerate(tables, start=1):
                if not table or len(table) < 2:
                    continue
                header, *rows = table
                header = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(header)]
                tdf = pd.DataFrame(rows, columns=header)
                # try to coerce numeric-looking columns
                for c in tdf.columns:
                    coerced = pd.to_numeric(tdf[c].astype(str).str.replace(",", ""), errors="coerce")
                    if coerced.notna().sum() >= len(tdf) * 0.6:
                        tdf[c] = coerced
                results.append((tdf, page_num, t_num))
    return results


def load_file(file):
    """Reads an uploaded file (CSV, Excel, or PDF) into one or more DataFrames.
    Returns a list of (label, dataframe) pairs, since a PDF can contain several
    tables while CSV/Excel always yields exactly one."""
    file.seek(0)
    name = file.name.lower()
    if name.endswith('.csv'):
        return [(file.name, pd.read_csv(file))]
    elif name.endswith(('.xlsx', '.xls')):
        return [(file.name, pd.read_excel(file))]
    elif name.endswith('.pdf'):
        if not PDF_SUPPORT:
            st.error("PDF support requires the `pdfplumber` package, which isn't installed here.")
            return []
        tables = extract_pdf_tables(file)
        if not tables:
            st.warning(f"No tables could be detected in **{file.name}**.")
            return []
        return [
            (f"{file.name} (p{page}-t{tnum})", tdf)
            for tdf, page, tnum in tables
        ]
    else:
        st.error(f"Unsupported file type: {file.name}")
        return []


def growth_pct(first_score, last_score):
    """Returns % growth from first_score to last_score, or None if not computable."""
    if pd.isna(first_score) or pd.isna(last_score) or first_score == 0:
        return None
    return round(((last_score - first_score) / first_score) * 100, 2)


def subject_scale_changed(long_df, subject, first_semester, last_semester, tolerance=0.10):
    """Checks whether a subject's apparent maximum marks shifted between two semesters
    (e.g. exam went from out of 50 to out of 100). Uses the observed max score within each
    semester as a proxy for 'marks out of'. Returns True if the two maxima differ by more than
    `tolerance` (10% by default) — in which case a raw % growth number is not meaningful,
    since part of the change is just a different scale, not actual improvement."""
    first_max = long_df.loc[
        (long_df['Subject'] == subject) & (long_df['__Semester__'] == first_semester), 'Score'
    ].max()
    last_max = long_df.loc[
        (long_df['Subject'] == subject) & (long_df['__Semester__'] == last_semester), 'Score'
    ].max()
    if pd.isna(first_max) or pd.isna(last_max) or first_max == 0 or last_max == 0:
        return False
    ratio = last_max / first_max
    return abs(ratio - 1) > tolerance


def find_student_row(df, name_col, search_name):
    """Case-insensitive, whitespace-tolerant name lookup. Returns (row_or_None, match_count)."""
    matches = df[df[name_col].astype(str).str.strip().str.lower() == search_name.strip().lower()]
    if matches.empty:
        return None, 0
    return matches.iloc[0], len(matches)


def bar_chart_height(n_categories, per_category_px=28, min_height=420, max_height=1400):
    """Scales bar chart height with the number of categories so long name lists stay readable."""
    return int(min(max_height, max(min_height, n_categories * per_category_px)))


# ------------------------------------------------------------
# NEW UTILITIES — v3/v4
# ------------------------------------------------------------

def global_search(df, query):
    """Case-insensitive substring search across every column of df. Returns the
    matching subset. Empty query returns df unchanged."""
    if not query:
        return df
    q = query.strip().lower()
    mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(q, na=False)).any(axis=1)
    return df[mask]


def detect_outliers_iqr(df, numeric_cols, k=1.5):
    """Flags outlier rows per numeric column using the IQR rule. Returns a summary
    dataframe (column, lower/upper bound, outlier count) and a combined boolean
    mask of rows that are an outlier on at least one column."""
    rows = []
    combined_mask = pd.Series(False, index=df.index)
    for col in numeric_cols:
        series = pd.to_numeric(df[col], errors="coerce")
        q1, q3 = series.quantile(0.25), series.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0 or pd.isna(iqr):
            continue
        lower, upper = q1 - k * iqr, q3 + k * iqr
        col_mask = (series < lower) | (series > upper)
        combined_mask |= col_mask.fillna(False)
        rows.append({
            "Column": col, "Lower Bound": round(lower, 2), "Upper Bound": round(upper, 2),
            "Outlier Count": int(col_mask.sum())
        })
    return pd.DataFrame(rows), combined_mask


def clean_dataframe(df, drop_dupes, fillna_strategy, fillna_cols, drop_empty_cols, strip_whitespace):
    """Applies a set of light-touch cleaning operations chosen in the UI and returns
    (cleaned_df, list_of_change_descriptions) so the user sees exactly what changed."""
    log = []
    out = df.copy()

    if strip_whitespace:
        obj_cols = out.select_dtypes(include="object").columns
        for c in obj_cols:
            out[c] = out[c].astype(str).str.strip()
        if len(obj_cols):
            log.append(f"Trimmed whitespace on {len(obj_cols)} text column(s).")

    if drop_empty_cols:
        empty_cols = [c for c in out.columns if out[c].isna().all()]
        if empty_cols:
            out = out.drop(columns=empty_cols)
            log.append(f"Dropped {len(empty_cols)} fully-empty column(s): {', '.join(empty_cols)}")

    if drop_dupes:
        before = len(out)
        out = out.drop_duplicates()
        removed = before - len(out)
        if removed:
            log.append(f"Removed {removed} duplicate row(s).")

    if fillna_strategy != "None" and fillna_cols:
        for c in fillna_cols:
            if fillna_strategy == "Mean" and pd.api.types.is_numeric_dtype(out[c]):
                out[c] = out[c].fillna(out[c].mean())
            elif fillna_strategy == "Median" and pd.api.types.is_numeric_dtype(out[c]):
                out[c] = out[c].fillna(out[c].median())
            elif fillna_strategy == "Zero":
                out[c] = out[c].fillna(0)
        log.append(f"Filled missing values in {len(fillna_cols)} column(s) using '{fillna_strategy}'.")

    if not log:
        log.append("No changes applied — data is unchanged.")
    return out, log


_ID_LIKE_NAME_PATTERNS = ("id", "roll", "no.", "number", "code", "index", "rank", "year", "phone", "zip", "pincode")


def _looks_like_identifier_column(df, col):
    """Heuristic check for columns that are numeric but not really a 'metric' — roll numbers,
    IDs, ranks, phone numbers, etc. These shouldn't be averaged/compared against real metrics
    like marks, since 'highest average Roll No' is a technically-true but meaningless insight."""
    name_lower = str(col).strip().lower()
    if any(pat in name_lower for pat in _ID_LIKE_NAME_PATTERNS):
        return True
    series = df[col].dropna()
    if series.empty:
        return False
    # near-unique integer-looking column (e.g. roll numbers, serial IDs) — every value distinct
    if series.nunique() >= max(1, int(len(series) * 0.98)) and (series % 1 == 0).all():
        return True
    return False


def auto_insights(df, numeric_cols, name_col=None):
    """Generates a short list of plain-language, non-AI bullet insights purely from
    the data itself: best/worst columns, spread, and skew — a free fallback to the
    paid AI Insights tab.

    Only columns that look like genuine, comparable metrics (not IDs/roll numbers/ranks) are
    used for "highest/lowest average" style claims, since those are only meaningful when the
    columns being compared measure the same kind of thing.
    """
    bullets = []
    if not numeric_cols:
        return ["No numeric columns available to summarize."]

    metric_cols = [c for c in numeric_cols if not _looks_like_identifier_column(df, c)]
    excluded_cols = [c for c in numeric_cols if c not in metric_cols]

    if not metric_cols:
        return ["No numeric columns look like comparable metrics (the rest look like IDs/roll numbers/ranks)."]

    means = df[metric_cols].mean(numeric_only=True)
    stds = df[metric_cols].std(numeric_only=True)

    top_col = means.idxmax()
    bullets.append(f"**{top_col}** has the highest average ({means[top_col]:.2f}) among comparable metric columns.")

    if len(means) > 1:
        low_col = means.idxmin()
        bullets.append(f"**{low_col}** has the lowest average ({means[low_col]:.2f}) — worth a closer look.")

    most_variable = stds.idxmax() if not stds.empty else None
    if most_variable:
        bullets.append(f"**{most_variable}** shows the most spread (std dev {stds[most_variable]:.2f}), "
                        f"meaning entities differ widely on this metric.")

    if name_col:
        for col in metric_cols[:1]:
            top_row = df.loc[df[col].idxmax()]
            bot_row = df.loc[df[col].idxmin()]
            bullets.append(f"On **{col}**, **{top_row[name_col]}** leads ({top_row[col]}) while "
                            f"**{bot_row[name_col]}** trails ({bot_row[col]}).")

    n_rows = len(df)
    bullets.append(f"Dataset contains **{n_rows}** row(s) across **{len(metric_cols)}** comparable numeric metric(s).")
    if excluded_cols:
        bullets.append(
            f"ℹ️ Excluded from the comparisons above (look like IDs/roll numbers/ranks, not performance metrics): "
            f"{', '.join(f'**{c}**' for c in excluded_cols)}."
        )
    bullets.append(
        "⚠️ Note: even among the metric columns above, averages are only meaningful to compare "
        "*to each other* if the columns share the same scale (e.g. both out of 100). A metric out "
        "of 50 will look artificially 'lower' next to one out of 100 even if performance is identical."
    )
    return bullets


def build_excel_report(named_dfs, primary_label, df, numeric_cols):
    """Bundles raw data (per file), a stats summary, and a correlation matrix into
    a single downloadable multi-sheet Excel workbook."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for label, tdf in named_dfs:
            sheet = str(label)[:31].replace("/", "_").replace("\\", "_") or "Sheet"
            tdf.to_excel(writer, sheet_name=sheet, index=False)
        if numeric_cols:
            df[numeric_cols].describe().T.to_excel(writer, sheet_name="Summary Stats")
            if len(numeric_cols) >= 2:
                df[numeric_cols].corr(numeric_only=True).to_excel(writer, sheet_name="Correlation")
    buffer.seek(0)
    return buffer


def build_radar_figure(df, radar_cols, row_indices, label_col, colors):
    """Builds a Plotly radar (polar) figure for the given rows and metric columns."""
    fig_radar = go.Figure()
    for i, idx in enumerate(row_indices):
        r_values = df.loc[idx, radar_cols].tolist()
        r_values += [r_values[0]]  # close the loop
        categories = radar_cols + [radar_cols[0]]
        fig_radar.add_trace(go.Scatterpolar(
            r=r_values, theta=categories, fill='toself',
            name=str(df.loc[idx, label_col]) if label_col else f"Row {idx}",
            line=dict(color=colors[i % len(colors)])
        ))
    fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), showlegend=True)
    return fig_radar


def check_file_compatibility(named_dfs):
    """Compares columns across a list of (label, df) pairs. Returns a dict with
    'all_match' (bool), 'common_cols' (list shared by every file), and a per-file
    column report — used to decide whether files can be auto-combined or whether
    the user needs to pick which shared rows/columns to use."""
    col_sets = [set(df.columns) for _, df in named_dfs]
    common = set.intersection(*col_sets) if col_sets else set()
    all_match = all(cs == col_sets[0] for cs in col_sets)
    report = [
        {"file": label, "columns": len(df.columns), "rows": len(df)}
        for label, df in named_dfs
    ]
    # preserve original column order for the common set, using the first file as reference
    ordered_common = [c for c in named_dfs[0][1].columns if c in common] if named_dfs else []
    return {"all_match": all_match, "common_cols": ordered_common, "report": report}


def apply_plotly_theme(fig, dark_mode):
    """Applies a consistent dark/light template to a figure without touching its data/colors."""
    fig.update_layout(template="plotly_dark" if dark_mode else "plotly_white")
    return fig


def render_flexible_chart(df, chart_type, x_col, y_col, color_col, colors, cont_scale,
                           horizontal, sort_desc, dark_mode, title, height=None):
    """One entry point for every 'customizable' chart on the Explore tab — swaps chart
    type, orientation, sort order, and color scheme without duplicating plotting code."""
    plot_df = df.copy()
    if sort_desc and y_col in plot_df.columns and pd.api.types.is_numeric_dtype(plot_df[y_col]):
        plot_df = plot_df.sort_values(y_col, ascending=False)

    color_arg = color_col if color_col and color_col != "(none)" else None
    is_numeric_color = color_arg is not None and pd.api.types.is_numeric_dtype(plot_df[color_arg])
    color_kwargs = {"color_continuous_scale": cont_scale} if is_numeric_color else {"color_discrete_sequence": colors}

    xa, ya = (y_col, x_col) if horizontal and chart_type in ("Bar", "Box", "Violin") else (x_col, y_col)

    if chart_type == "Bar":
        fig = px.bar(plot_df, x=xa, y=ya, color=color_arg, title=title,
                     orientation="h" if horizontal else "v", **color_kwargs)
    elif chart_type == "Line":
        fig = px.line(plot_df, x=x_col, y=y_col, color=color_arg, markers=True, title=title, **color_kwargs)
    elif chart_type == "Area":
        fig = px.area(plot_df, x=x_col, y=y_col, color=color_arg, title=title, **color_kwargs)
    elif chart_type == "Scatter":
        fig = px.scatter(plot_df, x=x_col, y=y_col, color=color_arg, title=title, size_max=14, **color_kwargs)
    elif chart_type == "Pie":
        fig = px.pie(plot_df, names=x_col, values=y_col, title=title, color_discrete_sequence=colors)
    elif chart_type == "Box":
        fig = px.box(plot_df, x=xa, y=ya, color=color_arg, title=title, **color_kwargs)
    elif chart_type == "Violin":
        fig = px.violin(plot_df, x=xa, y=ya, color=color_arg, box=True, title=title, **color_kwargs)
    elif chart_type == "Histogram":
        fig = px.histogram(plot_df, x=x_col, color=color_arg, title=title, **color_kwargs)
    elif chart_type == "Treemap":
        fig = px.treemap(plot_df, path=[x_col], values=y_col, title=title, color_discrete_sequence=colors)
    else:
        fig = px.bar(plot_df, x=xa, y=ya, title=title, color_discrete_sequence=colors)

    if height:
        fig.update_layout(height=height)
    return apply_plotly_theme(fig, dark_mode)


# ============================================================
# PAGE SETUP + UI THEME
# ============================================================

st.set_page_config(page_title="DATAVIS v5 ULTIMATE", page_icon="📊", layout="wide")

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = True
if "theme_preset" not in st.session_state:
    st.session_state.theme_preset = "Midnight Violet"
if "compact_mode" not in st.session_state:
    st.session_state.compact_mode = False

with st.sidebar:
    st.header("🎨 Appearance")
    dark_mode = st.toggle("Dark mode", value=st.session_state.dark_mode)
    st.session_state.dark_mode = dark_mode
    theme_preset = st.selectbox("Theme preset", list(THEME_PRESETS.keys()),
                                 index=list(THEME_PRESETS.keys()).index(st.session_state.theme_preset))
    st.session_state.theme_preset = theme_preset
    compact_mode = st.toggle("Compact layout", value=st.session_state.compact_mode,
                              help="Tighter spacing and smaller headers — useful on laptop screens.")
    st.session_state.compact_mode = compact_mode
    palette_name = st.selectbox("Color palette (categories)", list(PALETTES.keys()), index=0)
    cont_scale_name = st.selectbox("Color scale (numeric gradients)", list(CONTINUOUS_SCALES.keys()), index=0)
    COLOR_SEQUENCE = PALETTES[palette_name]
    CONT_SCALE = CONTINUOUS_SCALES[cont_scale_name]
    accent = COLOR_SEQUENCE[0]

preset = THEME_PRESETS[theme_preset]
if dark_mode:
    bg1, bg2 = preset["bg1"], preset["bg2"]
    if preset["bg1"].startswith("#f") or preset["bg1"].startswith("#fb"):
        bg1, bg2 = "#0e0e17", "#161625"  # guard against picking a light preset while dark mode is on
    card_bg, text_col, muted = "rgba(255,255,255,0.04)", "#f0eefc", "#a39fc0"
else:
    bg1, bg2, card_bg, text_col, muted = "#f7f7fb", "#ffffff", "rgba(0,0,0,0.02)", "#1b1b2b", "#5c5876"
accent2 = preset["accent2"]
compact_pad = "6px" if compact_mode else "10px"

st.markdown(f"""
<style>
.stApp {{
    background: linear-gradient(160deg, {bg1}, {bg2});
    color: {text_col};
}}

h1 {{
    background: linear-gradient(90deg, {COLOR_SEQUENCE[0 % len(COLOR_SEQUENCE)]}, {COLOR_SEQUENCE[1 % len(COLOR_SEQUENCE)]}, {COLOR_SEQUENCE[2 % len(COLOR_SEQUENCE)]});
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-weight: 800;
}}

h2, h3 {{
    border-left: 4px solid {accent};
    padding-left: 10px;
    transition: border-color 0.25s ease-in-out;
    color: {text_col};
}}
h2:hover, h3:hover {{ border-left-color: {COLOR_SEQUENCE[1 % len(COLOR_SEQUENCE)]}; }}

div[data-testid="stPlotlyChart"] {{
    border-radius: 14px;
    padding: 8px;
    background: {card_bg};
    transition: box-shadow 0.25s ease-in-out, transform 0.2s ease-in-out;
    border: 1px solid rgba(128,128,128,0.12);
}}
div[data-testid="stPlotlyChart"]:hover {{
    box-shadow: 0 6px 22px rgba(124, 77, 255, 0.25);
    transform: translateY(-2px);
}}

[data-testid="stAppViewContainer"] {{ animation: fadeIn 0.6s ease-in; }}
@keyframes fadeIn {{ from {{ opacity: 0; }} to {{ opacity: 1; }} }}

div[data-testid="stAlert"] {{ border-radius: 10px; }}

section[data-testid="stSidebar"] {{
    border-right: 2px solid rgba(128,128,128,0.2);
}}

button[data-baseweb="tab"] {{ font-weight: 600; }}

/* metric-style cards used in the compatibility report */
.dv-card {{
    background: {card_bg};
    border: 1px solid rgba(128,128,128,0.15);
    border-radius: 12px;
    padding: {compact_pad} 16px;
    margin-bottom: 8px;
    transition: transform 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}}
.dv-card:hover {{ transform: translateY(-2px); box-shadow: 0 6px 18px rgba(0,0,0,0.18); }}

/* KPI stat row shown at the top once data is loaded */
.dv-kpi {{
    background: linear-gradient(135deg, {card_bg}, rgba(128,128,128,0.05));
    border: 1px solid rgba(128,128,128,0.15);
    border-left: 4px solid {accent2};
    border-radius: 12px;
    padding: 14px 16px;
    text-align: left;
}}
.dv-kpi .dv-kpi-label {{ font-size: 12px; color: {muted}; text-transform: uppercase; letter-spacing: 0.05em; }}
.dv-kpi .dv-kpi-value {{ font-size: 24px; font-weight: 800; color: {text_col}; }}

/* insight bullets in Dashboard tab */
.dv-insight {{
    background: {card_bg};
    border-left: 3px solid {accent2};
    border-radius: 8px;
    padding: 10px 14px;
    margin-bottom: 6px;
    color: {text_col};
    font-size: 14px;
}}

.dv-tag {{
    display: inline-block;
    background: rgba(128,128,128,0.15);
    border-radius: 999px;
    padding: 2px 10px;
    font-size: 11px;
    margin-right: 6px;
    color: {muted};
}}

.dv-byline {{
    display: inline-block;
    background: linear-gradient(90deg, {accent2}22, transparent);
    border: 1px solid {accent2}55;
    border-radius: 999px;
    padding: 3px 12px;
    font-size: 12px;
    color: {text_col};
    margin-top: 4px;
}}

.dv-about-hero {{
    background: linear-gradient(135deg, {card_bg}, {accent2}11);
    border: 1px solid rgba(128,128,128,0.15);
    border-radius: 16px;
    padding: 22px 26px;
    margin-bottom: 16px;
}}
.dv-about-hero h2 {{ margin: 0 0 4px 0; color: {text_col}; }}
.dv-about-hero .handle {{ color: {accent2}; font-weight: 700; }}

.dv-feature-pill {{
    display: inline-block;
    background: {card_bg};
    border: 1px solid rgba(128,128,128,0.18);
    border-radius: 10px;
    padding: 8px 12px;
    margin: 4px 6px 4px 0;
    font-size: 13px;
    color: {text_col};
}}
</style>
""", unsafe_allow_html=True)

st.title('DATAVIS v5 ULTIMATE 📊')
st.markdown(
    "<span class='dv-byline'>🥷 Coded by <b>Arun</b> · a.k.a <b>KIRITO_SERFORT</b></span>",
    unsafe_allow_html=True
)
st.caption(
    'Data analysis & visualization suite for students, academic work, and office use.\n\n'
    'PDF table extraction · multi-file compatibility checking · data cleaning · auto insights · '
    'Excel report export · fully customizable charts for large-scale, multi-dataset analysis. '
    'See the **ℹ️ About** tab for full details.'
)
st.divider()

# ============================================================
# SIDEBAR: DATA + SETTINGS
# ============================================================

with st.sidebar:
    st.header('📂 Data')
    file_types = ["csv", "xlsx", "xls"] + (["pdf"] if PDF_SUPPORT else [])
    uploaded_files = st.file_uploader(
        "Upload one or more files (CSV / Excel"
        + (" / PDF with tables" if PDF_SUPPORT else "")
        + "). Multiple files enable growth & compatibility tools.",
        type=file_types,
        accept_multiple_files=True
    )
    if not PDF_SUPPORT:
        st.caption("⚠️ PDF support unavailable — install `pdfplumber` to enable it.")

    st.header("📈 Growth Projection Settings")
    pv = st.number_input('Present value or score', min_value=0.0, value=1000.0, step=100.0)
    rate = st.number_input('Growth Rate %', min_value=-100.0, value=5.0, step=0.5) / 100
    periods = st.slider('Number of periods to project', min_value=1, max_value=50, value=10)

    st.header('🤖 AI Summary')
    api_key = st.text_input('Enter your Claude API key', type='password')

# ============================================================
# LOAD ALL FILES (flatten: PDFs can yield multiple tables per file)
# ============================================================

named_dfs = []          # list of (label, dataframe)
if uploaded_files:
    for f in uploaded_files:
        named_dfs.extend(load_file(f))

# ============================================================
# MAIN BODY: TABS
# ============================================================

(tab_dash, tab_charts, tab_stats, tab_personal, tab_compare, tab_growth,
 tab_raw, tab_clean, tab_proj, tab_export, tab_ai, tab_about) = st.tabs(
    ["🏠 Dashboard", "📊 Explore & Customize", "🧮 Stats & Correlation", "🎯 Personal Analysis",
     "⚔️ Head-to-Head", "📅 Multi-File Growth", "📁 Raw Data", "🧹 Data Cleaning",
     "🔮 Projection", "📤 Export Report", "🤖 AI Insights", "ℹ️ About"]
)

# -------------------- TAB: ABOUT (data-independent) --------------------
with tab_about:
    st.markdown(
        """
        <div class='dv-about-hero'>
            <h2>DATAVIS v5 ULTIMATE 📊</h2>
            <p>Coded by <span class='handle'>Arun</span> — online handle
            <span class='handle'>KIRITO_SERFORT</span></p>
            <p style='margin-top:10px;'>A one-file Streamlit data analysis & visualization suite built
            for students, academic project work, and everyday office reporting — upload a spreadsheet
            or PDF and get charts, statistics, cleaning tools, and an exportable report without writing
            any code.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### 👤 Creator")
        st.markdown(
            "<div class='dv-card'>"
            "<b>Name:</b> Arun<br>"
            "<b>Online handle:</b> KIRITO_SERFORT<br>"
            "<b>Project:</b> DATAVIS — personal data visualization suite<br>"
            "<b>Built with:</b> Python · Streamlit · Plotly · Pandas"
            "</div>",
            unsafe_allow_html=True
        )
        st.markdown("#### 🕓 Version History")
        st.markdown(
            "<div class='dv-card'>"
            "<b>v5 ULTIMATE</b> — About tab, refreshed UI styling, rebrand<br>"
            "<b>v3 ULTIMATE</b> — Dashboard, Data Cleaning, Export Report, searchable raw data, "
            "theme presets<br>"
            "<b>v2.5</b> — PDF table extraction, multi-file compatibility checking, customizable charts"
            "</div>",
            unsafe_allow_html=True
        )

    with col_b:
        st.markdown("#### 🧩 What it's for")
        st.markdown(
            "<div class='dv-card'>Built to take the busywork out of turning raw spreadsheets or "
            "scanned/PDF tables into readable charts and summaries — for coursework, lab reports, "
            "society/club records, small office datasets, or comparing performance across terms and "
            "semesters, without needing Excel formulas or a separate BI tool.</div>",
            unsafe_allow_html=True
        )
        st.markdown("#### ✨ Feature Highlights")
        for feat in [
            "📁 CSV / Excel / PDF table upload, single or multi-file",
            "🔗 Multi-file column compatibility check",
            "📊 9 chart types + radar + 3D scatter",
            "🧮 Correlation heatmap & distribution explorer",
            "🎯 Personal & Head-to-Head comparisons",
            "📅 Multi-file growth tracking",
            "🧹 Cleaning tools + IQR outlier detection",
            "📤 One-click multi-sheet Excel report",
            "🤖 Optional AI-powered insights",
            "🎨 5 theme presets + dark/light + compact mode",
        ]:
            st.markdown(f"<span class='dv-feature-pill'>{feat}</span>", unsafe_allow_html=True)

    st.divider()
    st.caption(
        "DATAVIS is an independent personal project and is not affiliated with any institution. "
        "Feedback and feature requests are welcome — this suite is actively evolving version to version."
    )

# -------------------- TAB: PROJECTION (data-independent) --------------------
with tab_proj:
    st.subheader('Compound Growth Projection')
    years = list(range(0, periods + 1))
    proj_values = [pv * (1 + rate) ** n for n in years]
    projection_df = pd.DataFrame({'Period': years, 'Projected Value': proj_values})
    c1, c2 = st.columns([1, 2])
    with c1:
        st.dataframe(projection_df, use_container_width=True, height=420)
    with c2:
        fig_growth = px.line(projection_df, x='Period', y='Projected Value', markers=True,
                              title=f"Projection at {rate*100:.1f}% per Period",
                              color_discrete_sequence=COLOR_SEQUENCE)
        fig_growth = apply_plotly_theme(fig_growth, dark_mode)
        st.plotly_chart(fig_growth, use_container_width=True)

if named_dfs:
    labels = [lbl for lbl, _ in named_dfs]
    with st.sidebar:
        st.header('🗂️ Active Dataset')
        primary_label = st.selectbox(
            "Dataset to use for single-file tools (Explore, Stats, Personal, Head-to-Head)",
            labels, index=0
        )
    df = dict(named_dfs)[primary_label].copy()

    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    non_numeric_cols = df.select_dtypes(exclude=np.number).columns.tolist()

    with st.sidebar:
        if numeric_cols:
            st.header('🔍 Filter Data')
            filter_col = st.selectbox('Filter rows by column', numeric_cols)
            fc = pd.to_numeric(df[filter_col], errors="coerce")
            min_val, max_val = float(fc.min()), float(fc.max())
            if min_val == max_val:
                st.caption(f"All rows share the same {filter_col} value ({min_val}) — nothing to filter.")
                selected_range = (min_val, max_val)
            else:
                selected_range = st.slider(f"Range for {filter_col}", min_val, max_val, (min_val, max_val))
            df = df[(fc >= selected_range[0]) & (fc <= selected_range[1])]

        st.download_button(
            "⬇️ Download filtered data (CSV)",
            df.to_csv(index=False).encode("utf-8"),
            file_name=f"filtered_{primary_label.replace('/', '_')}.csv",
            mime="text/csv"
        )

    use_stacked_layout = len(df) > LARGE_DATASET_THRESHOLD

    # -------------------- TAB: DASHBOARD (overview + KPIs + auto insights) --------------------
    with tab_dash:
        st.subheader(f"Overview: {primary_label}")

        k1, k2, k3, k4 = st.columns(4)
        kpi_defs = [
            ("Rows", f"{len(df):,}"),
            ("Columns", f"{len(df.columns)}"),
            ("Numeric metrics", f"{len(numeric_cols)}"),
            ("Missing cells", f"{int(df.isna().sum().sum()):,}"),
        ]
        for col, (label, value) in zip([k1, k2, k3, k4], kpi_defs):
            col.markdown(
                f"<div class='dv-kpi'><div class='dv-kpi-label'>{label}</div>"
                f"<div class='dv-kpi-value'>{value}</div></div>",
                unsafe_allow_html=True
            )

        st.markdown("")
        st.markdown("#### 🧠 Auto Insights")
        st.caption("Generated instantly from the data itself — no API key needed. "
                   "For deeper narrative analysis, see the AI Insights tab.")
        insight_name_col = non_numeric_cols[0] if non_numeric_cols else None
        for bullet in auto_insights(df, numeric_cols, insight_name_col):
            st.markdown(f"<div class='dv-insight'>💡 {bullet}</div>", unsafe_allow_html=True)

        if len(named_dfs) > 1:
            st.markdown("")
            st.markdown("#### 🗂️ Files Loaded")
            tag_html = "".join(f"<span class='dv-tag'>{lbl} · {len(tdf)}×{len(tdf.columns)}</span>"
                                for lbl, tdf in named_dfs)
            st.markdown(tag_html, unsafe_allow_html=True)

        if numeric_cols:
            st.markdown("")
            st.markdown("#### ⚡ Quick Chart")
            quick_metric = st.selectbox("Metric to preview", numeric_cols, key="dash_quick_metric")
            quick_label_col = non_numeric_cols[0] if non_numeric_cols else None
            if quick_label_col:
                fig_quick = px.bar(df, x=quick_label_col, y=quick_metric, color_discrete_sequence=COLOR_SEQUENCE,
                                    title=f"{quick_metric} by {quick_label_col}")
            else:
                fig_quick = px.histogram(df, x=quick_metric, color_discrete_sequence=COLOR_SEQUENCE,
                                          title=f"Distribution of {quick_metric}")
            fig_quick = apply_plotly_theme(fig_quick, dark_mode)
            fig_quick.update_layout(height=380)
            st.plotly_chart(fig_quick, use_container_width=True)

    # -------------------- TAB: RAW DATA (per file, readable, labeled, searchable) --------------------
    with tab_raw:
        st.subheader("Raw Data by File")
        st.caption("Each uploaded file (or PDF table) is shown separately below, labeled by source.")
        raw_search = st.text_input("🔎 Search across all files (filters every table below)", key="raw_search")
        for label, tdf in named_dfs:
            shown_df = global_search(tdf, raw_search)
            match_note = "" if not raw_search else f"  ·  {len(shown_df)} match(es)"
            with st.expander(f"📄 {label}  ·  {len(tdf)} rows × {len(tdf.columns)} cols{match_note}",
                              expanded=(len(named_dfs) == 1 or bool(raw_search))):
                if raw_search and shown_df.empty:
                    st.caption("No matches in this file.")
                else:
                    st.dataframe(shown_df, use_container_width=True,
                                 height=min(420, 60 + 35 * min(len(shown_df), 10)))
                    st.download_button(
                        f"⬇️ Download this table (CSV)",
                        shown_df.to_csv(index=False).encode("utf-8"),
                        file_name=f"{str(label).replace('/', '_').replace(' ', '_')}.csv",
                        mime="text/csv",
                        key=f"dl_raw_{label}"
                    )

    # -------------------- TAB: EXPLORE & CUSTOMIZE --------------------
    with tab_charts:
        st.subheader(f"Exploring: {primary_label}  ({len(df)} rows)")
        if use_stacked_layout:
            st.info(f"📊 {len(df)} rows detected — using a full-width layout and taller charts for readability.")

        st.markdown("#### 🛠️ Chart Builder")
        cc1, cc2, cc3, cc4 = st.columns(4)
        with cc1:
            chart_type = st.selectbox("Chart type", CHART_TYPES, index=0)
        with cc2:
            x_col = st.selectbox('X-axis / categories', df.columns, key="explore_x")
        with cc3:
            y_options = numeric_cols if numeric_cols else df.columns.tolist()
            y_col = st.selectbox('Y-axis / values', y_options, key="explore_y")
        with cc4:
            color_col = st.selectbox("Color by", ["(none)"] + df.columns.tolist(), key="explore_color")

        cc5, cc6, cc7 = st.columns(3)
        with cc5:
            horizontal = st.checkbox("Horizontal orientation", value=use_stacked_layout)
        with cc6:
            sort_desc = st.checkbox("Sort by value (desc)", value=False)
        with cc7:
            custom_height = st.slider("Chart height (px)", 300, 1400,
                                       bar_chart_height(df[x_col].nunique()) if x_col in df.columns else 500)

        fig = render_flexible_chart(
            df, chart_type, x_col, y_col, color_col, COLOR_SEQUENCE, CONT_SCALE,
            horizontal, sort_desc, dark_mode, title=f"{y_col} by {x_col}", height=custom_height
        )
        st.plotly_chart(fig, use_container_width=True)
        st.download_button("⬇️ Download this chart (HTML)", fig.to_html(), file_name="chart.html", mime="text/html")

        st.divider()

        # ---------- RADAR CHART ----------
        st.subheader('Radar (Hex) Chart — compare rows across metrics')
        rc1, rc2 = st.columns([1, 2])
        with rc1:
            if non_numeric_cols:
                label_col = st.selectbox("Row label column", non_numeric_cols, key="radar_label_col")
            else:
                label_col = None
                st.info("No text column found for labels — rows will be labeled by number instead.")
            radar_cols = st.multiselect(
                "Metrics to compare (spokes)", numeric_cols,
                default=numeric_cols[:5] if len(numeric_cols) >= 5 else numeric_cols
            )
        with rc2:
            row_indices = st.multiselect(
                "Rows (entities) to plot", df.index.tolist(), default=df.index[:2].tolist()
            )
            if radar_cols and row_indices:
                fig_radar = build_radar_figure(df, radar_cols, row_indices, label_col, COLOR_SEQUENCE)
                fig_radar.update_layout(height=520)
                fig_radar = apply_plotly_theme(fig_radar, dark_mode)
                st.plotly_chart(fig_radar, use_container_width=True)
            else:
                st.info("Select at least one metric and one row to see the radar chart.")

        st.divider()

        # ---------- 3D SCATTER EXPLORER ----------
        st.subheader("3D Explorer")
        st.caption("Plot every row in 3D space across three metrics at once — reveals clusters "
                   "and outliers that 2D charts can hide. Great for large-scale, multi-metric datasets.")

        if len(numeric_cols) >= 3:
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                x3d = st.selectbox("X axis", numeric_cols, index=0, key="x3d")
            with c2:
                y3d = st.selectbox("Y axis", numeric_cols, index=1, key="y3d")
            with c3:
                z3d = st.selectbox("Z axis", numeric_cols, index=2, key="z3d")
            with c4:
                color3d = st.selectbox("Color by", ["(none)"] + df.columns.tolist(), key="color3d")

            hover_name = label_col if 'label_col' in dir() and label_col else None
            is_num_color = color3d != "(none)" and pd.api.types.is_numeric_dtype(df[color3d])
            fig_3d = px.scatter_3d(
                df, x=x3d, y=y3d, z=z3d,
                color=None if color3d == "(none)" else color3d,
                hover_name=hover_name,
                title=f"{x3d} vs {y3d} vs {z3d}",
                opacity=0.85,
                color_continuous_scale=CONT_SCALE if is_num_color else None,
                color_discrete_sequence=COLOR_SEQUENCE if not is_num_color else None,
            )
            fig_3d.update_traces(marker=dict(size=6, line=dict(width=0.5, color='white')))
            fig_3d.update_layout(scene=dict(xaxis_title=x3d, yaxis_title=y3d, zaxis_title=z3d), height=650)
            fig_3d = apply_plotly_theme(fig_3d, dark_mode)
            st.plotly_chart(fig_3d, use_container_width=True)
        else:
            st.info("Need at least 3 numeric columns for the 3D explorer.")

    # -------------------- TAB: STATS & CORRELATION --------------------
    with tab_stats:
        st.subheader(f"Statistical Summary: {primary_label}")
        if numeric_cols:
            st.dataframe(df[numeric_cols].describe().T, use_container_width=True)
        else:
            st.info("No numeric columns to summarize.")

        st.markdown("#### Missing Values")
        missing = df.isna().sum()
        missing = missing[missing > 0]
        if missing.empty:
            st.success("No missing values detected in this dataset.")
        else:
            fig_missing = px.bar(
                x=missing.index, y=missing.values, title="Missing Values per Column",
                labels={"x": "Column", "y": "Missing Count"}, color_discrete_sequence=COLOR_SEQUENCE
            )
            fig_missing = apply_plotly_theme(fig_missing, dark_mode)
            st.plotly_chart(fig_missing, use_container_width=True)

        st.markdown("#### Correlation Heatmap")
        st.caption("Shows how strongly numeric columns move together — useful for spotting redundant "
                   "metrics or hidden relationships across large datasets.")
        if len(numeric_cols) >= 2:
            corr = df[numeric_cols].corr(numeric_only=True)
            fig_corr = px.imshow(
                corr, text_auto=".2f", aspect="auto",
                color_continuous_scale=CONT_SCALE, title="Correlation Matrix"
            )
            fig_corr.update_layout(height=max(420, 40 * len(numeric_cols)))
            fig_corr = apply_plotly_theme(fig_corr, dark_mode)
            st.plotly_chart(fig_corr, use_container_width=True)
        else:
            st.info("Need at least 2 numeric columns for a correlation heatmap.")

        st.markdown("#### Distribution Explorer")
        if numeric_cols:
            dist_col = st.selectbox("Column to inspect", numeric_cols, key="dist_col")
            fig_dist = px.histogram(
                df, x=dist_col, marginal="box", title=f"Distribution of {dist_col}",
                color_discrete_sequence=COLOR_SEQUENCE
            )
            fig_dist = apply_plotly_theme(fig_dist, dark_mode)
            st.plotly_chart(fig_dist, use_container_width=True)

    # -------------------- TAB: PERSONAL ANALYSIS --------------------
    with tab_personal:
        st.subheader("Personal Analysis")

        if non_numeric_cols:
            name_col = st.selectbox("Which column holds names?", non_numeric_cols, key="personal_name_col")
            search_name = st.text_input("Enter a name to analyze", key="personal_search")
        else:
            name_col, search_name = None, None
            st.info("No text column found to identify entities by name.")

        if name_col and search_name:
            student_row, match_count = find_student_row(df, name_col, search_name)

            if student_row is None:
                st.warning(f"No entry named '{search_name}' found in this data.")
            else:
                st.success(f"Found: {student_row[name_col]}" + (f"  ·  {match_count} matching rows" if match_count > 1 else ""))

                comparison_rows = []
                for subject in numeric_cols:
                    score = student_row[subject]
                    higher = df[df[subject] > score]
                    comparison_rows.append({
                        "Subject": subject,
                        "Score": score,
                        "Above": len(higher),
                        "Avg of Those Above": round(higher[subject].mean(), 2) if not higher.empty else None,
                        "Gap to Topper": round(df[subject].max() - score, 2)
                    })

                comparison_df = pd.DataFrame(comparison_rows)
                st.dataframe(comparison_df, use_container_width=True)

                chart_data = comparison_df.melt(
                    id_vars="Subject", value_vars=["Score", "Avg of Those Above"],
                    var_name="Metric", value_name="Value"
                )
                fig_personal = px.bar(
                    chart_data, x="Subject", y="Value", color="Metric", barmode="group",
                    title=f"{student_row[name_col]}'s Scores vs. Those Above Them",
                    color_discrete_sequence=COLOR_SEQUENCE
                )
                fig_personal = apply_plotly_theme(fig_personal, dark_mode)
                st.plotly_chart(fig_personal, use_container_width=True)

    # -------------------- TAB: HEAD-TO-HEAD --------------------
    with tab_compare:
        st.subheader("Head-to-Head Comparison")

        if non_numeric_cols:
            vs_name_col = st.selectbox("Name column", non_numeric_cols, key="vs_name_col")
            entity_options = df[vs_name_col].astype(str).tolist()

            colA, colB = st.columns(2)
            with colA:
                person_a = st.selectbox("First entity", entity_options, key="vs_a")
            with colB:
                default_b_index = 1 if len(entity_options) > 1 else 0
                person_b = st.selectbox("Second entity", entity_options, index=default_b_index, key="vs_b")

            if person_a and person_b and person_a != person_b:
                row_a, count_a = find_student_row(df, vs_name_col, person_a)
                row_b, count_b = find_student_row(df, vs_name_col, person_b)

                if count_a > 1:
                    st.warning(
                        f"⚠️ **{count_a} rows** match the name '{person_a}' — showing the first one found. "
                        "If there are two different people with this name, double-check which row is intended."
                    )
                if count_b > 1:
                    st.warning(
                        f"⚠️ **{count_b} rows** match the name '{person_b}' — showing the first one found. "
                        "If there are two different people with this name, double-check which row is intended."
                    )

                vs_cols = st.multiselect(
                    "Metrics to compare", numeric_cols,
                    default=numeric_cols[:6] if len(numeric_cols) >= 6 else numeric_cols,
                    key="vs_cols"
                )

                if vs_cols:
                    fig_vs = go.Figure()
                    for i, (row, label) in enumerate([(row_a, person_a), (row_b, person_b)]):
                        r_values = row[vs_cols].tolist()
                        r_values += [r_values[0]]
                        categories = vs_cols + [vs_cols[0]]
                        fig_vs.add_trace(go.Scatterpolar(
                            r=r_values, theta=categories, fill='toself', name=label,
                            line=dict(color=COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)])
                        ))
                    fig_vs.update_layout(polar=dict(radialaxis=dict(visible=True)), showlegend=True,
                                          title=f"{person_a} vs {person_b}")
                    fig_vs = apply_plotly_theme(fig_vs, dark_mode)
                    st.plotly_chart(fig_vs, use_container_width=True)

                    diff_rows = []
                    for subject in vs_cols:
                        a_score, b_score = row_a[subject], row_b[subject]
                        diff_rows.append({
                            "Subject": subject,
                            person_a: a_score,
                            person_b: b_score,
                            "Difference": round(a_score - b_score, 2),
                            "Leader": person_a if a_score > b_score else (person_b if b_score > a_score else "Tie")
                        })
                    diff_df = pd.DataFrame(diff_rows)
                    st.dataframe(diff_df, use_container_width=True)

                    a_wins = (diff_df["Leader"] == person_a).sum()
                    b_wins = (diff_df["Leader"] == person_b).sum()
                    st.info(f"**{person_a}** leads in {a_wins} metric(s) · **{person_b}** leads in {b_wins} metric(s)")
            else:
                st.info("Pick two different entities to compare.")
        else:
            st.info("No text column found to identify entities by name.")

    # -------------------- TAB: MULTI-FILE GROWTH --------------------
    with tab_growth:
        if len(named_dfs) < 2:
            st.info("Upload 2 or more files (or a multi-table PDF, e.g. one per semester) in the sidebar to unlock this section.")
        else:
            st.subheader("File Compatibility Check")
            compat = check_file_compatibility(named_dfs)

            cols = st.columns(len(named_dfs)) if len(named_dfs) <= 4 else [st]
            for i, r in enumerate(compat["report"]):
                target = cols[i] if len(named_dfs) <= 4 else st
                target.markdown(
                    f"<div class='dv-card'><b>{r['file']}</b><br>{r['rows']} rows · {r['columns']} columns</div>",
                    unsafe_allow_html=True
                )

            if compat["all_match"]:
                st.success("✅ All files share identical columns — they can be combined automatically.")
                use_cols = compat["common_cols"]
            else:
                st.warning(
                    "⚠️ These files don't share the exact same columns (different headers or column counts), "
                    "so they can't be auto-combined. Pick the columns that are common and meaningful across "
                    "all of them below."
                )
                use_cols = st.multiselect(
                    "Columns to use for combined analysis (shared across every file)",
                    compat["common_cols"], default=compat["common_cols"]
                )

            if not use_cols:
                st.info("Select at least one shared column to continue.")
            else:
                semester_dfs = []
                for i, (label, tdf) in enumerate(named_dfs):
                    sem_df = tdf[use_cols].copy()
                    sem_df['__Semester__'] = f"{i + 1}. {label}"
                    semester_dfs.append(sem_df)

                combined_df = pd.concat(semester_dfs, ignore_index=True)
                name_cols_multi = [c for c in combined_df.select_dtypes(exclude=np.number).columns if c != '__Semester__']

                if not name_cols_multi:
                    st.info("No text column found among the shared columns to identify entities by name across files.")
                else:
                    st.divider()
                    name_col_multi = st.selectbox(
                        "Which shared column holds entity/student names?", name_cols_multi, key="multi_name_col"
                    )
                    numeric_cols_multi = combined_df.select_dtypes(include=np.number).columns.tolist()

                    if not numeric_cols_multi:
                        st.info("No shared numeric columns to track growth on.")
                    else:
                        long_df = combined_df.melt(
                            id_vars=['__Semester__', name_col_multi], value_vars=numeric_cols_multi,
                            var_name='Subject', value_name='Score'
                        ).rename(columns={name_col_multi: 'Name'})

                        semesters_sorted = sorted(combined_df['__Semester__'].unique())
                        first_semester, last_semester = semesters_sorted[0], semesters_sorted[-1]

                        st.markdown("#### Individual Growth Trend")
                        search_name_multi = st.text_input(
                            "Name to see growth across files", key="multi_search_name"
                        )

                        student_long = pd.DataFrame()
                        if search_name_multi:
                            student_long = long_df[
                                long_df['Name'].astype(str).str.strip().str.lower() == search_name_multi.strip().lower()
                            ]
                            if student_long.empty:
                                st.warning(f"No entry named '{search_name_multi}' found across the uploaded files.")
                            else:
                                fig_trend = px.line(
                                    student_long, x='__Semester__', y='Score', color='Subject', markers=True,
                                    title=f"{search_name_multi}'s Growth Across Files",
                                    color_discrete_sequence=COLOR_SEQUENCE
                                )
                                fig_trend = apply_plotly_theme(fig_trend, dark_mode)
                                st.plotly_chart(fig_trend, use_container_width=True)

                        st.markdown("#### Top 10 Growth (Overview)")
                        avg_per_student_sem = long_df.groupby(['Name', '__Semester__'])['Score'].mean().reset_index()
                        latest_scores = avg_per_student_sem[avg_per_student_sem['__Semester__'] == last_semester]
                        top10_names = latest_scores.sort_values('Score', ascending=False).head(10)['Name'].tolist()
                        top10_long = long_df[long_df['Name'].isin(top10_names)]

                        scale_changed_subjects = [
                            s for s in numeric_cols_multi
                            if subject_scale_changed(long_df, s, first_semester, last_semester)
                        ]

                        growth_rows = []
                        for student in top10_names:
                            for subject in numeric_cols_multi:
                                fv = top10_long[(top10_long['Name'] == student) & (top10_long['Subject'] == subject) &
                                                 (top10_long['__Semester__'] == first_semester)]['Score']
                                lv = top10_long[(top10_long['Name'] == student) & (top10_long['Subject'] == subject) &
                                                 (top10_long['__Semester__'] == last_semester)]['Score']
                                scale_changed = subject in scale_changed_subjects
                                g = (growth_pct(fv.values[0], lv.values[0])
                                     if not fv.empty and not lv.empty and not scale_changed else None)
                                growth_rows.append({
                                    'Name': student, 'Subject': subject, 'Growth %': g,
                                    'Scale Changed?': 'Yes — max marks differ between semesters' if scale_changed else 'No'
                                })

                        growth_df = pd.DataFrame(growth_rows)
                        if scale_changed_subjects:
                            st.warning(
                                f"⚠️ **{', '.join(scale_changed_subjects)}** appear to have a different maximum "
                                f"score in {first_semester} vs. {last_semester} (e.g. out of 50 vs. out of 100). "
                                "A raw % growth number would be misleading there — it's shown as blank for those "
                                "subjects instead of a wrong number."
                            )
                        st.dataframe(growth_df, use_container_width=True)

                        avg_growth_by_subject = growth_df.groupby('Subject')['Growth %'].mean().reset_index()
                        fig_top10 = px.bar(avg_growth_by_subject, x='Subject', y='Growth %',
                                            title="Top 10's Average Growth Rate by Subject",
                                            color='Subject', color_discrete_sequence=COLOR_SEQUENCE)
                        fig_top10 = apply_plotly_theme(fig_top10, dark_mode)
                        st.plotly_chart(fig_top10, use_container_width=True)

                        if search_name_multi and not student_long.empty:
                            st.markdown("#### Searched Entity vs. Top 10 Growth")
                            searched_rows = []
                            for subject in numeric_cols_multi:
                                fs = student_long[(student_long['Subject'] == subject) &
                                                   (student_long['__Semester__'] == first_semester)]['Score']
                                ls = student_long[(student_long['Subject'] == subject) &
                                                   (student_long['__Semester__'] == last_semester)]['Score']
                                scale_changed = subject_scale_changed(long_df, subject, first_semester, last_semester)
                                g = (growth_pct(fs.values[0], ls.values[0])
                                     if not fs.empty and not ls.empty and not scale_changed else None)
                                searched_rows.append({'Subject': subject, 'Growth %': g, 'Who': search_name_multi})

                            searched_growth_df = pd.DataFrame(searched_rows)
                            top10_avg_labeled = avg_growth_by_subject.copy()
                            top10_avg_labeled['Who'] = 'Top 10 Average'
                            compare_df = pd.concat(
                                [searched_growth_df, top10_avg_labeled[['Subject', 'Growth %', 'Who']]], ignore_index=True
                            )
                            fig_compare = px.bar(compare_df, x='Subject', y='Growth %', color='Who', barmode='group',
                                                  title=f"{search_name_multi} vs. Top 10 Average Growth Rate",
                                                  color_discrete_sequence=COLOR_SEQUENCE)
                            fig_compare = apply_plotly_theme(fig_compare, dark_mode)
                            st.plotly_chart(fig_compare, use_container_width=True)

    # -------------------- TAB: DATA CLEANING --------------------
    with tab_clean:
        st.subheader(f"Data Cleaning: {primary_label}")
        st.caption("Light-touch cleanup tools. Nothing here overwrites your uploaded file — "
                   "download the cleaned result when you're happy with it.")

        oc1, oc2 = st.columns(2)
        with oc1:
            drop_dupes = st.checkbox("Remove duplicate rows", value=False)
            strip_whitespace = st.checkbox("Trim whitespace in text columns", value=True)
            drop_empty_cols = st.checkbox("Drop fully-empty columns", value=False)
        with oc2:
            fillna_strategy = st.selectbox("Fill missing values using…",
                                            ["None", "Mean", "Median", "Zero"])
            st.caption(
                "⚠️ There's no 'Forward fill' option here on purpose: for records like exam marks, "
                "a blank cell usually means the person was absent, not that they scored whatever the "
                "row above them scored. Filling it with the previous row's value would silently "
                "fabricate a mark for them. Mean/Median/Zero don't have that specific failure mode, "
                "but even they can distort things — if a student was genuinely absent, consider leaving "
                "that cell blank and excluding them from that column's calculations instead of filling it."
            )
            fillna_cols = []
            if fillna_strategy != "None":
                fillna_cols = st.multiselect("Columns to fill", df.columns.tolist(),
                                              default=[c for c in numeric_cols if df[c].isna().any()])

        cleaned_df, change_log = clean_dataframe(df, drop_dupes, fillna_strategy, fillna_cols,
                                                  drop_empty_cols, strip_whitespace)

        st.markdown("#### Change Log")
        for entry in change_log:
            st.markdown(f"<div class='dv-insight'>✅ {entry}</div>", unsafe_allow_html=True)

        st.markdown("#### Outlier Detection (IQR method)")
        st.caption("Flags values far outside the normal range for each numeric column — often the "
                   "fastest way to spot data-entry errors.")
        if numeric_cols:
            outlier_summary, outlier_mask = detect_outliers_iqr(cleaned_df, numeric_cols)
            if outlier_summary.empty:
                st.info("No column had enough spread to flag outliers.")
            else:
                st.dataframe(outlier_summary, use_container_width=True)
                n_flagged = int(outlier_mask.sum())
                if n_flagged:
                    with st.expander(f"🚩 {n_flagged} row(s) flagged as an outlier on at least one column"):
                        st.dataframe(cleaned_df[outlier_mask], use_container_width=True)
                else:
                    st.success("No individual rows flagged as outliers.")
        else:
            st.info("No numeric columns to check for outliers.")

        st.markdown("#### Preview & Download")
        st.dataframe(cleaned_df, use_container_width=True, height=min(420, 60 + 35 * min(len(cleaned_df), 10)))
        st.download_button(
            "⬇️ Download cleaned data (CSV)",
            cleaned_df.to_csv(index=False).encode("utf-8"),
            file_name=f"cleaned_{str(primary_label).replace('/', '_')}.csv",
            mime="text/csv"
        )

    # -------------------- TAB: EXPORT REPORT --------------------
    with tab_export:
        st.subheader("Export a Combined Report")
        st.caption("Bundles every uploaded file's raw data, plus summary statistics and the correlation "
                   "matrix for the active dataset, into one Excel workbook — handy for sharing or archiving.")

        if not EXCEL_EXPORT_SUPPORT:
            st.warning("Excel export requires the `openpyxl` package, which isn't installed here.")
        else:
            st.markdown(
                f"<div class='dv-card'>Workbook will include:<br>"
                f"• {len(named_dfs)} raw data sheet(s)<br>"
                f"• 1 summary statistics sheet (active dataset: <b>{primary_label}</b>)<br>"
                f"• 1 correlation matrix sheet (if ≥2 numeric columns)</div>",
                unsafe_allow_html=True
            )
            report_buffer = build_excel_report(named_dfs, primary_label, df, numeric_cols)
            st.download_button(
                "⬇️ Download Excel report (.xlsx)",
                report_buffer,
                file_name=f"DATAVIS-V5_report_{datetime.date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.divider()
        st.markdown("#### Export current chart")
        st.caption("Grab the latest chart built on the Explore & Customize tab as a standalone HTML file "
                   "from the download button on that tab, or export any Plotly chart as PNG using the "
                   "camera icon in its top-right toolbar.")

    # -------------------- TAB: AI INSIGHTS --------------------
    with tab_ai:
        st.subheader(f"AI-Generated Insights: {primary_label}")
        if not AI_SUPPORT:
            st.info("AI insights require `langchain_anthropic`, which isn't installed in this environment.")
        elif api_key:
            st.caption(
                "By default, only column names and summary statistics (means, quartiles, etc.) are sent "
                "to Claude — never the row-level data itself, so no individual's name or score leaves the app."
            )
            include_raw = st.checkbox(
                "Also include row-level data, so the AI can call out specific rows/entities by name "
                "(names will be anonymized as 'Entity 1', 'Entity 2', etc. before sending, and translated "
                "back for display here — but this does still send every row's numeric values to Anthropic's API)",
                value=False
            )
            if include_raw:
                st.warning(
                    "⚠️ This will send every row of this dataset's data (with names replaced by placeholders) "
                    "to Anthropic's API. If this data is about other people (e.g. classmates), make sure "
                    "you're comfortable sharing their scores this way before continuing."
                )
            if st.button('Generate AI Summary'):
                with st.spinner('Analyzing data...'):
                    ai_name_col = non_numeric_cols[0] if (include_raw and non_numeric_cols) else None
                    summary, error = get_ai_summary(df, api_key, name_col=ai_name_col, include_raw_rows=include_raw)
                if error:
                    st.error(f"Couldn't generate summary: {error}")
                else:
                    st.markdown(summary)
        else:
            st.info('Enter your Claude API key in the sidebar to generate an AI summary.')

else:
    with tab_dash:
        st.info('Upload a CSV, Excel, or PDF file in the sidebar to see your dashboard.')
    with tab_charts:
        st.info('Upload a CSV, Excel, or PDF file in the sidebar to get started.')
    with tab_raw:
        st.info('Upload a file to see its raw data here.')
    with tab_clean:
        st.info('Upload a file to use the data cleaning tools.')
    with tab_export:
        st.info('Upload a file to build an exportable report.')
