# DATAVIS v5 ULTIMATE — Data Visualizer & Student Marks Analysis Portal
# Coded by Arun a.k.a Arun
# to run -   streamlit run app.py

import io
import json
import datetime
import hashlib
import os
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image

import auth_db
import supabase_db

# Initialize local SQLite database
auth_db.init_db()

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "datavis_logo.jpg")

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
# CONSTANTS & CONFIG
# ============================================================

LARGE_DATASET_THRESHOLD = 15   # rows above this trigger stacked layout
HUGE_DATASET_THRESHOLD = 300   # rows above this trigger table pagination

# Payment & Pricing Constants (Indian Rupee ₹)
UPI_ID = "arun@okaxis"  # Default GPay UPI ID (Configurable in settings)

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
    """Replaces every value in name_col with a generic placeholder ('Entity 1', 'Entity 2', ...)."""
    anon_df = df.copy()
    if name_col and name_col in anon_df.columns:
        unique_names = anon_df[name_col].astype(str).unique().tolist()
        mapping = {name: f"Entity {i + 1}" for i, name in enumerate(unique_names)}
        anon_df[name_col] = anon_df[name_col].astype(str).map(mapping)
        return anon_df, mapping
    return anon_df, {}


def get_ai_summary(df, api_key, name_col=None, include_raw_rows=False):
    """Sends dataset stats to Claude and asks for a plain-language analysis."""
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
            if name_col and name_col in df.columns:
                anon_df, mapping = anonymize_names(df, name_col)
                raw_data_section = f"\nRaw data (names replaced with anonymous placeholders):\n{anon_df.to_string()}\n"
            else:
                raw_data_section = f"\nRaw data:\n{df.to_string()}\n"

        prompt = f"""
You are an expert data analyst. Analyze this dataset:
Columns: {columns_info}

Summary statistics:
{summary_stats}
{raw_data_section}

Provide:
1. Key findings and highest/lowest performing entities/metrics.
2. Areas of weakness or subjects where scores trail.
3. Actionable recommendations for improvement.
"""
        response = llm.invoke([HumanMessage(content=prompt)])
        text = response.content
        if mapping:
            reverse_map = {v: k for k, v in mapping.items()}
            for placeholder, real_name in reverse_map.items():
                text = text.replace(placeholder, real_name)
        return text, None
    except Exception as e:
        return None, str(e)


def extract_pdf_tables(file):
    """Extracts tables found in a PDF using pdfplumber."""
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
                for c in tdf.columns:
                    coerced = pd.to_numeric(tdf[c].astype(str).str.replace(",", ""), errors="coerce")
                    if coerced.notna().sum() >= len(tdf) * 0.6:
                        tdf[c] = coerced
                results.append((tdf, page_num, t_num))
    return results


def load_file(file):
    """Reads uploaded CSV, Excel, or PDF into DataFrames."""
    file.seek(0)
    name = file.name.lower()
    if name.endswith('.csv'):
        return [(file.name, pd.read_csv(file))]
    elif name.endswith(('.xlsx', '.xls')):
        return [(file.name, pd.read_excel(file))]
    elif name.endswith('.pdf'):
        if not PDF_SUPPORT:
            st.error("PDF support requires the `pdfplumber` package.")
            return []
        tables = extract_pdf_tables(file)
        if not tables:
            st.warning(f"No tables detected in **{file.name}**.")
            return []
        return [(f"{file.name} (p{page}-t{tnum})", tdf) for tdf, page, tnum in tables]
    else:
        st.error(f"Unsupported file type: {file.name}")
        return []


def generate_sample_marks_data():
    """Generates a sample student marks dataframe for 1-click testing."""
    np.random.seed(42)
    names = ["Aarav Sharma", "Ananya Verma", "Rohan Gupta", "Priya Patel", "Vikram Singh",
             "Neha Reddy", "Aditya Kumar", "Sneha Iyer", "Karan Malhotra", "Riya Sen"]
    data = {
        "Roll No": [101 + i for i in range(len(names))],
        "Student Name": names,
        "Mathematics": [92, 85, 78, 65, 95, 88, 72, 98, 54, 81],
        "Physics": [88, 90, 82, 70, 91, 84, 68, 96, 60, 79],
        "Chemistry": [90, 84, 75, 68, 89, 86, 74, 94, 58, 83],
        "English": [94, 88, 80, 75, 92, 90, 81, 95, 71, 87],
        "Computer Science": [96, 92, 85, 78, 98, 91, 79, 99, 65, 89]
    }
    return pd.DataFrame(data)


def compute_student_grades(df, numeric_cols, name_col=None):
    """Computes Total Marks, Average Score %, Grade (A+, A, B, C, D, Fail), and Rank."""
    out_df = df.copy()
    valid_metrics = [c for c in numeric_cols if not _looks_like_identifier_column(df, c)]
    if not valid_metrics:
        return out_df

    out_df["Total Marks"] = out_df[valid_metrics].sum(axis=1)
    out_df["Average %"] = out_df[valid_metrics].mean(axis=1).round(2)

    def get_grade(avg):
        if avg >= 90: return "A+ (Outstanding)"
        elif avg >= 80: return "A (Excellent)"
        elif avg >= 70: return "B (Good)"
        elif avg >= 60: return "C (Average)"
        elif avg >= 50: return "D (Pass)"
        else: return "F (Fail)"

    out_df["Grade"] = out_df["Average %"].apply(get_grade)
    out_df["Class Rank"] = out_df["Average %"].rank(ascending=False, method="min").astype(int)
    out_df = out_df.sort_values("Class Rank")
    return out_df


def growth_pct(first_score, last_score):
    if pd.isna(first_score) or pd.isna(last_score) or first_score == 0:
        return None
    return round(((last_score - first_score) / first_score) * 100, 2)


def subject_scale_changed(long_df, subject, first_semester, last_semester, tolerance=0.10):
    first_max = long_df.loc[
        (long_df['Subject'] == subject) & (long_df['__Semester__'] == first_semester), 'Score'
    ].max()
    last_max = long_df.loc[
        (long_df['Subject'] == subject) & (long_df['__Semester__'] == last_semester), 'Score'
    ].max()
    if pd.isna(first_max) or pd.isna(last_max) or first_max == 0 or last_max == 0:
        return False
    return abs((last_max / first_max) - 1) > tolerance


def find_student_row(df, name_col, search_name):
    if not search_name or not name_col or name_col not in df.columns:
        return None, 0
    matches = df[df[name_col].astype(str).str.strip().str.lower() == search_name.strip().lower()]
    if matches.empty:
        return None, 0
    return matches.iloc[0], len(matches)


def bar_chart_height(n_categories, per_category_px=28, min_height=420, max_height=1400):
    return int(min(max_height, max(min_height, n_categories * per_category_px)))


def global_search(df, query):
    if not query:
        return df
    q = query.strip().lower()
    mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(q, na=False)).any(axis=1)
    return df[mask]


def detect_outliers_iqr(df, numeric_cols, k=1.5):
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
            log.append(f"Dropped {len(empty_cols)} fully-empty column(s).")

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
    name_lower = str(col).strip().lower()
    if any(pat in name_lower for pat in _ID_LIKE_NAME_PATTERNS):
        return True
    series = df[col].dropna()
    if series.empty:
        return False
    # Heuristic: all unique integers look like IDs — BUT if values are in a
    # typical marks/percentage range (0–150) they are almost certainly scores,
    # not identifiers (e.g. "Mathematics": [92, 85, 78, …]).
    if pd.api.types.is_numeric_dtype(series):
        col_min, col_max = series.min(), series.max()
        if col_min >= 0 and col_max <= 150:
            return False   # Score/percentage range — definitely not an ID
    if series.nunique() >= max(1, int(len(series) * 0.98)) and (series % 1 == 0).all():
        return True
    return False


def auto_insights(df, numeric_cols, name_col=None):
    bullets = []
    if not numeric_cols:
        return ["No numeric columns available to summarize."]

    metric_cols = [c for c in numeric_cols if not _looks_like_identifier_column(df, c)]
    excluded_cols = [c for c in numeric_cols if c not in metric_cols]

    if not metric_cols:
        return ["No numeric columns look like performance metrics (others look like IDs/roll numbers)."]

    means = df[metric_cols].mean(numeric_only=True)
    stds = df[metric_cols].std(numeric_only=True)

    if not means.empty:
        top_col = means.idxmax()
        bullets.append(f"**{top_col}** has the highest class average ({means[top_col]:.2f}).")
        if len(means) > 1:
            low_col = means.idxmin()
            bullets.append(f"**{low_col}** has the lowest class average ({means[low_col]:.2f}) — needs focus.")

    if not stds.empty:
        most_var = stds.idxmax()
        bullets.append(f"**{most_var}** shows the highest score spread (std dev {stds[most_var]:.2f}).")

    if name_col and name_col in df.columns:
        for col in metric_cols[:1]:
            top_row = df.loc[df[col].idxmax()]
            bot_row = df.loc[df[col].idxmin()]
            bullets.append(f"On **{col}**, **{top_row[name_col]}** leads ({top_row[col]}) while **{bot_row[name_col]}** trails ({bot_row[col]}).")

    bullets.append(f"Dataset contains **{len(df)}** student/entity row(s) across **{len(metric_cols)}** metric column(s).")
    return bullets


def build_excel_report(named_dfs, primary_label, df, numeric_cols):
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
    fig_radar = go.Figure()
    for i, idx in enumerate(row_indices):
        if idx in df.index:
            r_values = df.loc[idx, radar_cols].tolist()
            r_values += [r_values[0]]
            categories = radar_cols + [radar_cols[0]]
            label_text = str(df.loc[idx, label_col]) if (label_col and label_col in df.columns) else f"Row {idx}"
            fig_radar.add_trace(go.Scatterpolar(
                r=r_values, theta=categories, fill='toself', name=label_text,
                line=dict(color=colors[i % len(colors)])
            ))
    fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), showlegend=True)
    return fig_radar


def check_file_compatibility(named_dfs):
    col_sets = [set(df.columns) for _, df in named_dfs]
    common = set.intersection(*col_sets) if col_sets else set()
    all_match = all(cs == col_sets[0] for cs in col_sets)
    report = [{"file": label, "columns": len(df.columns), "rows": len(df)} for label, df in named_dfs]
    ordered_common = [c for c in named_dfs[0][1].columns if c in common] if named_dfs else []
    return {"all_match": all_match, "common_cols": ordered_common, "report": report}


def apply_plotly_theme(fig, dark_mode):
    fig.update_layout(template="plotly_dark" if dark_mode else "plotly_white")
    return fig


def render_flexible_chart(df, chart_type, x_col, y_col, color_col, colors, cont_scale,
                           horizontal, sort_desc, dark_mode, title, height=None):
    plot_df = df.copy()
    if sort_desc and y_col in plot_df.columns and pd.api.types.is_numeric_dtype(plot_df[y_col]):
        plot_df = plot_df.sort_values(y_col, ascending=False)

    color_arg = color_col if color_col and color_col != "(none)" else None
    is_numeric_color = color_arg is not None and pd.api.types.is_numeric_dtype(plot_df[color_arg])
    color_kwargs = {"color_continuous_scale": cont_scale} if is_numeric_color else {"color_discrete_sequence": colors}

    xa, ya = (y_col, x_col) if horizontal and chart_type in ("Bar", "Box", "Violin") else (x_col, y_col)

    if chart_type == "Bar":
        fig = px.bar(plot_df, x=xa, y=ya, color=color_arg, title=title, orientation="h" if horizontal else "v", **color_kwargs)
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
# PAGE SETUP + SESSION STATE
# ============================================================

st.set_page_config(page_title="DATAVIS v5 ULTIMATE", page_icon="📊", layout="wide")

if "user" not in st.session_state:
    st.session_state.user = None
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = True
if "theme_preset" not in st.session_state:
    st.session_state.theme_preset = "Midnight Violet"
if "compact_mode" not in st.session_state:
    st.session_state.compact_mode = False
if "charged_session_id" not in st.session_state:
    st.session_state.charged_session_id = None
if "sample_data_loaded" not in st.session_state:
    st.session_state.sample_data_loaded = False


# ============================================================
# SIDEBAR: BRANDING, APPEARANCE & USER INFO
# ============================================================

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width='stretch')

    st.markdown("<h2 style='text-align:center; margin-top:-10px;'>DATAVIS v5</h2>", unsafe_allow_html=True)
    st.markdown(
        "<div style='text-align:center; font-size:12px; color:#18C6C6; margin-bottom:15px; font-weight:600;'>"
        "✨ Created by Arun a.k.a Arun"
        "</div>",
        unsafe_allow_html=True
    )

    st.header("🎨 Appearance & Settings")
    dark_mode = st.toggle("Dark mode", value=st.session_state.dark_mode)
    st.session_state.dark_mode = dark_mode
    theme_preset = st.selectbox("Theme preset", list(THEME_PRESETS.keys()),
                                 index=list(THEME_PRESETS.keys()).index(st.session_state.theme_preset))
    st.session_state.theme_preset = theme_preset
    compact_mode = st.toggle("Compact layout", value=st.session_state.compact_mode)
    st.session_state.compact_mode = compact_mode

    palette_name = st.selectbox("Color palette (categories)", list(PALETTES.keys()), index=0)
    cont_scale_name = st.selectbox("Color scale (numeric gradients)", list(CONTINUOUS_SCALES.keys()), index=0)
    COLOR_SEQUENCE = PALETTES[palette_name]
    CONT_SCALE = CONTINUOUS_SCALES[cont_scale_name]
    accent = COLOR_SEQUENCE[0]

    st.divider()

    # User Auth Status Panel in Sidebar
    if st.session_state.user:
        user_info = auth_db.get_user_by_id(st.session_state.user['id'])
        if user_info:
            st.session_state.user = user_info
        else:
            st.session_state.user = None

    if st.session_state.user:
        u = st.session_state.user
        st.markdown(f"### 👤 Account Profile")
        st.markdown(f"**{u['name']}**  \n`<span style='color:#a39fc0;'>{u['email']}</span>`", unsafe_allow_html=True)

        turns = u.get('turns_remaining', 0)
        is_paid = u.get('is_paid_tier', 0)

        # Do NOT explicitly reveal turn counter to user unless exhausted!
        if turns <= 0:
            st.error("🚨 **0 Free Analyses Left** (Top up for ₹20)")

        if st.button("🚪 Sign Out", key="sidebar_logout"):
            st.session_state.user = None
            st.session_state.charged_session_id = None
            st.rerun()

preset = THEME_PRESETS[theme_preset]
if dark_mode:
    bg1, bg2 = preset["bg1"], preset["bg2"]
    if preset["bg1"].startswith("#f") or preset["bg1"].startswith("#fb"):
        bg1, bg2 = "#0e0e17", "#161625"
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

.dv-card {{
    background: {card_bg};
    border: 1px solid rgba(128,128,128,0.15);
    border-radius: 14px;
    padding: {compact_pad} 18px;
    margin-bottom: 10px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.15);
}}

.dv-kpi {{
    background: linear-gradient(135deg, {card_bg}, rgba(128,128,128,0.05));
    border: 1px solid rgba(128,128,128,0.15);
    border-left: 4px solid {accent2};
    border-radius: 12px;
    padding: 14px 16px;
}}
.dv-kpi .dv-kpi-label {{ font-size: 12px; color: {muted}; text-transform: uppercase; letter-spacing:0.05em; }}
.dv-kpi .dv-kpi-value {{ font-size: 24px; font-weight: 800; color: {text_col}; }}

.dv-insight {{
    background: {card_bg};
    border-left: 3px solid {accent2};
    border-radius: 8px;
    padding: 10px 14px;
    margin-bottom: 6px;
    color: {text_col};
}}

.dv-auth-hero {{
    background: linear-gradient(135deg, {card_bg}, {accent2}22);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 18px;
    padding: 35px;
    margin-bottom: 25px;
    text-align: center;
}}

.dv-about-hero {{
    background: linear-gradient(135deg, {card_bg}, {accent2}18);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 18px;
    padding: 32px 36px;
    margin-bottom: 22px;
    text-align: center;
    box-shadow: 0 8px 24px rgba(0,0,0,0.18);
}}

.dv-tag {{
    display: inline-block;
    background: {accent2}28;
    border: 1px solid {accent2}60;
    color: {accent2};
    border-radius: 20px;
    padding: 4px 12px;
    font-size: 13px;
    font-weight: 600;
    margin: 3px 4px;
}}

.dv-medal-gold {{
    background: linear-gradient(135deg, #ffd700, #ffbf00, #e6a900);
    border-radius: 14px; padding: 16px 18px; text-align: center;
    box-shadow: 0 4px 18px rgba(255,215,0,0.25); border: 1px solid #ffd700;
}}
.dv-medal-silver {{
    background: linear-gradient(135deg, #c0c0c0, #a8a8a8, #909090);
    border-radius: 14px; padding: 16px 18px; text-align: center;
    box-shadow: 0 4px 14px rgba(192,192,192,0.20); border: 1px solid #c0c0c0;
}}
.dv-medal-bronze {{
    background: linear-gradient(135deg, #cd7f32, #b87333, #a0652a);
    border-radius: 14px; padding: 16px 18px; text-align: center;
    box-shadow: 0 4px 14px rgba(205,127,50,0.20); border: 1px solid #cd7f32;
}}
.dv-medal-name {{ font-size: 17px; font-weight: 800; color: #1a1a1a; margin: 6px 0 2px; }}
.dv-medal-score {{ font-size: 13px; color: #2a2a2a; font-weight: 600; }}
.dv-medal-grade {{ font-size: 12px; color: #333; margin-top: 2px; }}

.dv-badge {{
    display: inline-block; padding: 3px 10px; border-radius: 20px;
    font-size: 12px; font-weight: 700; letter-spacing: 0.03em;
}}
.dv-badge-aplus  {{ background: #38f5b022; color: #38f5b0; border: 1px solid #38f5b060; }}
.dv-badge-a      {{ background: #7C4DFF22; color: #7C4DFF; border: 1px solid #7C4DFF60; }}
.dv-badge-b      {{ background: #18C6C622; color: #18C6C6; border: 1px solid #18C6C660; }}
.dv-badge-c      {{ background: #FFB90022; color: #FFB900; border: 1px solid #FFB90060; }}
.dv-badge-d      {{ background: #FF6EC722; color: #FF6EC7; border: 1px solid #FF6EC760; }}
.dv-badge-f      {{ background: #FF444422; color: #FF4444; border: 1px solid #FF444460; }}
</style>
""", unsafe_allow_html=True)


# ============================================================
# LANDING & LOGIN / SIGN UP VIEW (If not logged in)
# ============================================================

if not st.session_state.user:
    col_l1, col_l2 = st.columns([1, 4])
    with col_l1:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=130)
    with col_l2:
        st.title("DATAVIS v5 ULTIMATE 📊")
        st.markdown(
            "<span style='font-size:16px; color:#18C6C6; font-weight:600;'>"
            "Created by <b>Arun a.k.a Arun</b></span>",
            unsafe_allow_html=True
        )

    st.markdown("""
    <div class='dv-auth-hero'>
        <h2 style='border:none; margin:0;'>Smart Data & Student Performance Visualizer</h2>
        <p style='font-size: 16px; margin-top: 10px; color:#a39fc0;'>
            Turn student marks, spreadsheets, and academic records into interactive charts, grade analytics, and reports instantly.
        </p>
    </div>
    """, unsafe_allow_html=True)

    auth_tab1, auth_tab2 = st.tabs(["🔑 Sign In", "📝 Create Free Account"])

    with auth_tab1:
        st.subheader("Welcome Back!")
        with st.form("login_form"):
            login_email = st.text_input("Email Address")
            login_password = st.text_input("Password", type="password")
            submit_login = st.form_submit_button("Sign In to Portal")

            if submit_login:
                success, msg, user_dict = auth_db.authenticate_user(login_email, login_password)
                if success:
                    st.session_state.user = user_dict
                    supabase_db.sync_user_to_supabase(user_dict)
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    with auth_tab2:
        st.subheader("Create an Account")
        with st.form("signup_form"):
            reg_name = st.text_input("Full Name")
            reg_email = st.text_input("Email Address")
            reg_password = st.text_input("Password (min. 6 chars)", type="password")
            reg_confirm = st.text_input("Confirm Password", type="password")
            submit_reg = st.form_submit_button("Register Account")

            if submit_reg:
                if reg_password != reg_confirm:
                    st.error("Passwords do not match!")
                else:
                    success, msg, user_dict = auth_db.register_user(reg_name, reg_email, reg_password)
                    if success:
                        st.session_state.user = user_dict
                        supabase_db.sync_user_to_supabase(user_dict)
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    st.divider()
    st.markdown("### 🌟 Key Capabilities")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("<div class='dv-card'><h4>📊 Student Marks Visualizer</h4>Grade distributions, radar graphs, 3D scatter plots for exam scores.</div>", unsafe_allow_html=True)
    with c2:
        st.markdown("<div class='dv-card'><h4>⚔️ Head-to-Head Comparison</h4>Compare any two students side-by-side with radar charts and leader scorecards.</div>", unsafe_allow_html=True)
    with c3:
        st.markdown("<div class='dv-card'><h4>🏆 Leaderboard & Grade Analytics</h4>Automatic A+ to F grade calculations, class rank, and top scorer badges.</div>", unsafe_allow_html=True)

    st.stop()


# ============================================================
# LOGGED IN PORTAL HEADER & NAVBAR
# ============================================================

u = st.session_state.user or {
    "id": 0,
    "name": "Guest",
    "email": "guest@example.com",
    "turns_remaining": 0,
    "total_analyses": 0,
    "is_paid_tier": 0,
    "role": "user"
}
turns = u.get('turns_remaining', 0)
is_paid = u.get('is_paid_tier', 0)

col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title('DATAVIS v5 ULTIMATE 📊')
    st.markdown(
        f"<span style='color:#18C6C6; font-size:14px; font-weight:600;'>"
        f"👋 Welcome <b>{u['name']}</b> ({u['email']}) · <i>Created by Arun a.k.a Arun</i></span>",
        unsafe_allow_html=True
    )
with col_h2:
    if turns <= 0:
        st.error("🚨 **0 Analyses Left** (Top up ₹20)")

st.caption(
    'Data analysis & visualization suite for student marks, academic performance, and office datasets.'
)
st.divider()


# ============================================================
# SIDEBAR: DATA UPLOAD & SETTINGS
# ============================================================

with st.sidebar:
    st.header('📂 Data Upload')

    if st.button("🧪 Load Sample Student Marks Dataset", help="Load a pre-made student mark sheet to test all features instantly!"):
        st.session_state.sample_data_loaded = True
        st.toast("Loaded sample student marks dataset!", icon="✅")

    file_types = ["csv", "xlsx", "xls"] + (["pdf"] if PDF_SUPPORT else [])
    uploaded_files = st.file_uploader(
        "Upload your dataset (CSV / Excel"
        + (" / PDF" if PDF_SUPPORT else "")
        + ")",
        type=file_types,
        accept_multiple_files=True
    )

    st.header("📈 Growth Projection Settings")
    pv = st.number_input('Present value or score', min_value=0.0, value=1000.0, step=100.0)
    rate = st.number_input('Growth Rate %', min_value=-100.0, value=5.0, step=0.5) / 100
    periods = st.slider('Number of periods to project', min_value=1, max_value=50, value=10)

    st.header('🤖 AI Summary')
    api_key = st.text_input('Enter your Claude API key', type='password')


# ============================================================
# TURN & FILE UPLOAD LIMITS GUARD (HIDDEN LIMITS UNTIL EXHAUSTED)
# ============================================================

named_dfs = []

if st.session_state.sample_data_loaded and not uploaded_files:
    named_dfs = [("Sample_Student_Marks.csv", generate_sample_marks_data())]

if uploaded_files:
    n_uploaded = len(uploaded_files)

    # 1. FILE UPLOAD LIMIT CHECKS
    if not is_paid and n_uploaded > 3:
        st.error(f"🚨 **Upload Limit Exceeded (Free Tier)**")
        st.warning(
            f"Free Tier allows a maximum of **3 file uploads per analysis**. You uploaded **{n_uploaded} files**.\n\n"
            "Please remove the extra files or upgrade to Paid Tier (**₹20 for 2 analyses**) to upload up to 6 files per analysis!"
        )
        uploaded_files = None
    elif is_paid and n_uploaded > 6:
        extra_files = n_uploaded - 6
        extra_fee = extra_files * 5.0
        st.info(
            f"ℹ️ **Multi-File Upload Notice**: You uploaded **{n_uploaded} files** ({extra_files} extra files beyond the 6 included).\n"
            f"Extra file upload charge: **₹{extra_fee:.0f}** (₹5 per extra file beyond 6)."
        )

if uploaded_files:
    session_signature = hashlib.md5("".join([f.name + str(f.size) for f in uploaded_files]).encode()).hexdigest()

    # Check remaining analysis turns balance
    if st.session_state.charged_session_id != session_signature:
        if turns <= 0:
            st.error("🚨 **Free Analysis Limit Exhausted!**")
            st.warning(
                "You have completed your **3 free dataset analyses**!\n\n"
                "To continue analyzing datasets, comparing student marks, and exporting reports, "
                "please top up your account balance (**₹20 for 2 analyses** via GPay / UPI)."
            )
            st.info("💡 Go to the **💳 Account & Billing** tab at the top to complete GPay UPI payment.")
            uploaded_files = None
        else:
            dataset_names = ", ".join([f.name for f in uploaded_files])
            extra_fee = max(0, (len(uploaded_files) - 6) * 5.0) if is_paid else 0.0
            success, remaining, msg = auth_db.consume_turn(
                u['id'], "Data Analysis & Visualization", dataset_names, len(uploaded_files), extra_fee
            )
            if success:
                st.session_state.charged_session_id = session_signature
                st.session_state.user['turns_remaining'] = remaining
                turns = remaining

if uploaded_files:
    for f in uploaded_files:
        named_dfs.extend(load_file(f))


# ============================================================
# MAIN BODY: TABS
# ============================================================

(tab_dash, tab_grades, tab_charts, tab_stats, tab_personal, tab_compare, tab_growth,
 tab_raw, tab_clean, tab_proj, tab_export, tab_ai, tab_account, tab_about) = st.tabs(
    ["🏠 Dashboard", "🏆 Leaderboard & Grades", "📊 Explore & Customize", "🧮 Stats & Correlation",
     "🎯 Personal Analysis", "⚔️ Head-to-Head", "📅 Multi-File Growth", "📁 Raw Data",
     "🧹 Data Cleaning", "🔮 Projection", "📤 Export Report", "🤖 AI Insights", "💳 Account & Billing", "ℹ️ About"]
)


# -------------------- TAB: ACCOUNT & BILLING (GPAY / UPI ₹) --------------------
with tab_account:
    st.subheader("💳 Account & Billing (GPay / UPI Payment)")

    u_fresh = auth_db.get_user_by_id(u['id'])
    if u_fresh:
        st.session_state.user = u_fresh
        u = u_fresh

    ac1, ac2, ac3 = st.columns(3)
    ac1.markdown(f"<div class='dv-kpi'><div class='dv-kpi-label'>User Profile</div><div class='dv-kpi-value' style='font-size:18px;'>{u['name']}</div><small>{u['email']}</small></div>", unsafe_allow_html=True)
    ac2.markdown(f"<div class='dv-kpi'><div class='dv-kpi-label'>Analyses Balance</div><div class='dv-kpi-value'>⚡ {u['turns_remaining']} Left</div><small>{'Paid Tier (6 Uploads)' if u.get('is_paid_tier') else 'Free Tier (3 Uploads)'}</small></div>", unsafe_allow_html=True)
    ac3.markdown(f"<div class='dv-kpi'><div class='dv-kpi-label'>Total Executed</div><div class='dv-kpi-value'>📊 {u['total_analyses']}</div></div>", unsafe_allow_html=True)

    st.divider()

    st.markdown("### 📱 Top Up Analysis Balance (GPay / PhonePe / UPI)")
    st.caption("Pay easily using GPay, PhonePe, Paytm, or any UPI app.")

    p1, p2, p3 = st.columns(3)
    turns_to_buy = 0
    amount_inr = 0.0

    with p1:
        st.markdown("<div class='dv-card'><h4>⚡ 2 Analyses</h4><h2 style='color:#18C6C6;'>₹20</h2><p>Includes up to 6 file uploads per analysis.</p></div>", unsafe_allow_html=True)
        if st.button("Buy 2 Analyses (₹20)", key="buy_2"):
            turns_to_buy = 2
            amount_inr = 20.0

    with p2:
        st.markdown("<div class='dv-card'><h4>⚡ 5 Analyses</h4><h2 style='color:#7C4DFF;'>₹50</h2><p>Popular pack for semester records.</p></div>", unsafe_allow_html=True)
        if st.button("Buy 5 Analyses (₹50)", key="buy_5"):
            turns_to_buy = 5
            amount_inr = 50.0

    with p3:
        st.markdown("<div class='dv-card'><h4>⚡ 12 Analyses</h4><h2 style='color:#FF6EC7;'>₹100</h2><p>Best value pack for academic work.</p></div>", unsafe_allow_html=True)
        if st.button("Buy 12 Analyses (₹100)", key="buy_12"):
            turns_to_buy = 12
            amount_inr = 100.0

    if turns_to_buy > 0:
        st.markdown("---")
        st.subheader(f"GPay / UPI Payment: ₹{amount_inr:.0f} for {turns_to_buy} Analyses")

        pay_c1, pay_c2 = st.columns([1, 1])

        with pay_c1:
            st.markdown(f"""
            <div class='dv-card'>
                <h4>📲 How to Pay via GPay / UPI:</h4>
                <ol>
                    <li>Open <b>Google Pay</b>, <b>PhonePe</b>, <b>Paytm</b>, or any UPI App.</li>
                    <li>Pay <b>₹{amount_inr:.0f}</b> to UPI ID: <b style='color:#18C6C6; font-size:18px;'>{UPI_ID}</b></li>
                    <li>Copy the <b>12-digit UPI Ref No / UTR</b> from your GPay payment receipt.</li>
                    <li>Enter the UTR number on the right to activate your turns instantly!</li>
                </ol>
            </div>
            """, unsafe_allow_html=True)

        with pay_c2:
            st.markdown("#### 🔑 Confirm Payment & Enter UTR Number")
            with st.form("upi_confirm_form"):
                utr_num = st.text_input("Enter 12-digit UTR / UPI Ref Number", placeholder="e.g. 427891023456")
                submit_pay = st.form_submit_button("Verify & Credit Analyses", type="primary")

                if submit_pay:
                    if not utr_num or len(utr_num.strip()) < 4:
                        st.error("Please enter a valid UPI UTR / Reference number.")
                    else:
                        success, new_bal, msg = auth_db.add_credits_upi(u['id'], turns_to_buy, amount_inr, utr_num.strip())
                        if success:
                            st.success(msg)
                            st.session_state.user['turns_remaining'] = new_bal
                            st.session_state.user['is_paid_tier'] = 1
                            supabase_db.sync_user_to_supabase(st.session_state.user)
                            st.rerun()
                        else:
                            st.error(msg)

    st.divider()

    st.markdown("### 📋 Activity & Payment History")
    analyses_history, txn_history = auth_db.get_user_history(u['id'])

    h_tab1, h_tab2 = st.tabs(["📊 Analysis History", "💳 UPI Payment Transactions"])

    with h_tab1:
        if analyses_history:
            st.dataframe(pd.DataFrame(analyses_history), width='stretch')
        else:
            st.info("No analyses run yet.")

    with h_tab2:
        if txn_history:
            st.dataframe(pd.DataFrame(txn_history), width='stretch')
        else:
            st.info("No payment transactions recorded yet.")


# -------------------- TAB: ABOUT (data-independent) --------------------
with tab_about:
    st.markdown(
        """
        <div class='dv-about-hero'>
            <h2>DATAVIS v5 ULTIMATE 📊</h2>
            <p>Coded by <span style='color:#18C6C6; font-weight:700;'>Arun a.k.a Arun</span></p>
            <p style='margin-top:10px;'>A complete data analysis & student performance visualization portal
            with User Authentication, GPay / UPI payments, and Supabase integration.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### 👤 Creator Information")
        st.markdown(
            "<div class='dv-card'>"
            "<b>Name:</b> Arun<br>"
            "<b>Credit:</b> Created by Arun a.k.a Arun<br>"
            "<b>Project:</b> DATAVIS — data visualization suite<br>"
            "<b>Built with:</b> Python · Streamlit · SQLite / Supabase · Plotly · Pandas"
            "</div>",
            unsafe_allow_html=True
        )

    with col_b:
        st.markdown("#### 💳 Pricing & Limits")
        st.markdown(
            "<div class='dv-card'>"
            "<b>Free Tier:</b> 3 Free Analyses (Max 3 files per upload)<br>"
            "<b>Paid Tier (₹20 for 2 turns):</b> Max 6 files per upload<br>"
            "<b>Extra Uploads:</b> ₹5 per extra file beyond 6<br>"
            "<b>Payment Method:</b> GPay / PhonePe / Paytm / Any UPI ID"
            "</div>",
            unsafe_allow_html=True
        )


# -------------------- TAB: PROJECTION (data-independent) --------------------
with tab_proj:
    st.subheader('Compound Growth Projection')
    years = list(range(0, periods + 1))
    proj_values = [pv * (1 + rate) ** n for n in years]
    projection_df = pd.DataFrame({'Period': years, 'Projected Value': proj_values})
    c1, c2 = st.columns([1, 2])
    with c1:
        st.dataframe(projection_df, width='stretch', height=420)
    with c2:
        fig_growth = px.line(projection_df, x='Period', y='Projected Value', markers=True,
                              title=f"Projection at {rate*100:.1f}% per Period",
                              color_discrete_sequence=COLOR_SEQUENCE)
        fig_growth = apply_plotly_theme(fig_growth, dark_mode)
        st.plotly_chart(fig_growth, width='stretch')


# -------------------- ANALYSIS TABS (FOR LOADED DATASETS) --------------------
if named_dfs:
    labels = [lbl for lbl, _ in named_dfs]
    with st.sidebar:
        st.header('🗂️ Active Dataset')
        primary_label = st.selectbox(
            "Dataset to use for single-file tools",
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

    # Compute grades if metrics exist
    graded_df = compute_student_grades(df, numeric_cols, non_numeric_cols[0] if non_numeric_cols else None)

    # -------------------- TAB: DASHBOARD --------------------
    with tab_dash:
        st.subheader(f"Overview: {primary_label}")

        k1, k2, k3, k4 = st.columns(4)
        kpi_defs = [
            ("Rows / Students", f"{len(df):,}"),
            ("Columns", f"{len(df.columns)}"),
            ("Numeric Metrics", f"{len(numeric_cols)}"),
            ("Missing Cells", f"{int(df.isna().sum().sum()):,}"),
        ]
        for col, (label, value) in zip([k1, k2, k3, k4], kpi_defs):
            col.markdown(
                f"<div class='dv-kpi'><div class='dv-kpi-label'>{label}</div>"
                f"<div class='dv-kpi-value'>{value}</div></div>",
                unsafe_allow_html=True
            )

        if "Grade" in graded_df.columns:
            st.markdown("")
            st.markdown("#### 🏆 Performance Highlights")
            g1, g2, g3 = st.columns(3)

            top_row = graded_df.iloc[0]
            name_col_val = non_numeric_cols[0] if non_numeric_cols else "Student"
            top_name = top_row[name_col_val] if name_col_val in top_row else "Top Performer"

            pass_count = (graded_df["Average %"] >= 50).sum()
            pass_rate = round((pass_count / len(graded_df)) * 100, 1)

            g1.markdown(f"<div class='dv-kpi'><div class='dv-kpi-label'>👑 Top Rank 1</div><div class='dv-kpi-value' style='font-size:20px; color:#38f5b0;'>{top_name}</div><small>Avg: {top_row['Average %']}%</small></div>", unsafe_allow_html=True)
            g2.markdown(f"<div class='dv-kpi'><div class='dv-kpi-label'>📈 Class Pass Rate</div><div class='dv-kpi-value'>{pass_rate}%</div><small>{pass_count}/{len(graded_df)} Students Passed</small></div>", unsafe_allow_html=True)
            g3.markdown(f"<div class='dv-kpi'><div class='dv-kpi-label'>📊 Class Average</div><div class='dv-kpi-value'>{graded_df['Average %'].mean():.1f}%</div><small>Overall Subjects</small></div>", unsafe_allow_html=True)

        st.markdown("")
        st.markdown("#### 🧠 Auto Insights")
        insight_name_col = non_numeric_cols[0] if non_numeric_cols else None
        for bullet in auto_insights(df, numeric_cols, insight_name_col):
            st.markdown(f"<div class='dv-insight'>💡 {bullet}</div>", unsafe_allow_html=True)

        if len(named_dfs) > 1:
            st.markdown("")
            st.markdown("#### 🗂️ Files Loaded")
            tag_html = "".join(f"<span class='dv-tag'>{lbl} · {len(tdf)}×{len(tdf.columns)}</span>" for lbl, tdf in named_dfs)
            st.markdown(tag_html, unsafe_allow_html=True)

        if numeric_cols:
            st.markdown("")
            st.markdown("#### ⚡ Quick Chart Preview")
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
            st.plotly_chart(fig_quick, width='stretch')

    # -------------------- TAB: LEADERBOARD & GRADES --------------------
    with tab_grades:
        st.subheader(f"🏆 Class Leaderboard & Grade Analytics ({primary_label})")

        if "Grade" in graded_df.columns:
            name_col_lb = non_numeric_cols[0] if non_numeric_cols else None

            # ── PODIUM: TOP 3 MEDALS ────────────────────────────────────────
            sorted_lb = graded_df.sort_values("Class Rank").reset_index(drop=True)
            medal_emojis = ["🥇", "🥈", "🥉"]
            medal_css   = ["dv-medal-gold", "dv-medal-silver", "dv-medal-bronze"]

            if len(sorted_lb) >= 3:
                st.markdown("#### 🏅 Top 3 Podium")
                pod1, pod2, pod3 = st.columns(3)
                for podium_col, idx, medal_emoji, medal_class in zip(
                    [pod1, pod2, pod3], [0, 1, 2], medal_emojis, medal_css
                ):
                    row = sorted_lb.iloc[idx]
                    nm = str(row[name_col_lb]) if name_col_lb else f"Rank {idx+1}"
                    avg = f"{row['Average %']:.1f}%"
                    grade_raw = str(row["Grade"])
                    grade_badge_map = {
                        "A+": "dv-badge-aplus", "A ": "dv-badge-a", "B ": "dv-badge-b",
                        "C ": "dv-badge-c", "D ": "dv-badge-d", "F ": "dv-badge-f"
                    }
                    badge_cls = next((v for k, v in grade_badge_map.items() if grade_raw.startswith(k.strip())), "dv-badge-b")
                    podium_col.markdown(
                        f"<div class='{medal_class}'>"
                        f"<div style='font-size:36px;'>{medal_emoji}</div>"
                        f"<div class='dv-medal-name'>{nm}</div>"
                        f"<div class='dv-medal-score'>{avg}</div>"
                        f"<span class='dv-badge {badge_cls}'>{grade_raw}</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                st.markdown("")

            # ── MAIN RANKINGS TABLE + GRADE CHART ──────────────────────────
            col_l1, col_l2 = st.columns([2, 1])

            with col_l1:
                st.markdown("#### 📋 Full Class Rankings")
                # Colour-code Grade column with badges
                display_df = graded_df.copy()
                st.dataframe(display_df, width='stretch', height=450)

            with col_l2:
                st.markdown("#### 🍰 Grade Distribution")
                grade_counts = graded_df["Grade"].value_counts().reset_index()
                grade_counts.columns = ["Grade", "Count"]
                fig_grade_pie = px.pie(grade_counts, names="Grade", values="Count", title="Grade Share",
                                       color_discrete_sequence=COLOR_SEQUENCE)
                fig_grade_pie = apply_plotly_theme(fig_grade_pie, dark_mode)
                st.plotly_chart(fig_grade_pie, width='stretch')

            # ── SUBJECT TOPPERS ─────────────────────────────────────────────
            metric_cols_lb = [c for c in numeric_cols if not _looks_like_identifier_column(graded_df, c)
                               and c not in ("Total Marks", "Average %", "Class Rank")]
            if metric_cols_lb and name_col_lb:
                st.markdown("#### 🌟 Subject-Wise Toppers")
                topper_cols = st.columns(min(3, len(metric_cols_lb)))
                for i, subj in enumerate(metric_cols_lb):
                    top_idx = graded_df[subj].idxmax()
                    top_row = graded_df.loc[top_idx]
                    top_name_val = str(top_row[name_col_lb])
                    top_score = top_row[subj]
                    col_ix = i % 3
                    topper_cols[col_ix].markdown(
                        f"<div class='dv-subject-topper'>"
                        f"<div><b style='color:{accent2};'>{subj}</b><br>"
                        f"<span style='font-size:15px; font-weight:700;'>{top_name_val}</span></div>"
                        f"<div style='text-align:right;'><span style='font-size:22px; font-weight:800;'>{top_score}</span><br>"
                        f"<small style='color:{muted};'>Top Score</small></div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

            # ── GRADE DISTRIBUTION BAR ──────────────────────────────────────
            st.markdown("#### 📊 Grade Band Bar Chart")
            grade_order = ["A+ (Outstanding)", "A (Excellent)", "B (Good)", "C (Average)", "D (Pass)", "F (Fail)"]
            grade_bar_df = graded_df["Grade"].value_counts().reindex(grade_order).dropna().reset_index()
            grade_bar_df.columns = ["Grade", "Count"]
            fig_grade_bar = px.bar(
                grade_bar_df, x="Grade", y="Count",
                color="Grade", color_discrete_sequence=COLOR_SEQUENCE,
                title="Students per Grade Band", text="Count"
            )
            fig_grade_bar.update_traces(textposition="outside")
            fig_grade_bar = apply_plotly_theme(fig_grade_bar, dark_mode)
            st.plotly_chart(fig_grade_bar, width='stretch')

        else:
            st.info("No numeric metric columns detected to calculate grades.")

    # -------------------- TAB: RAW DATA --------------------
    with tab_raw:
        st.subheader("Raw Data by File")
        raw_search = st.text_input("🔎 Search across all files", key="raw_search")
        for label, tdf in named_dfs:
            shown_df = global_search(tdf, raw_search)
            match_note = "" if not raw_search else f"  ·  {len(shown_df)} match(es)"
            with st.expander(f"📄 {label}  ·  {len(tdf)} rows × {len(tdf.columns)} cols{match_note}",
                              expanded=(len(named_dfs) == 1 or bool(raw_search))):
                if raw_search and shown_df.empty:
                    st.caption("No matches in this file.")
                else:
                    st.dataframe(shown_df, width='stretch',
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
        st.plotly_chart(fig, width='stretch')
        st.download_button("⬇️ Download this chart (HTML)", fig.to_html(), file_name="chart.html", mime="text/html")

        st.divider()

        # Radar Chart
        st.subheader('Radar (Hex) Chart — compare rows across metrics')
        rc1, rc2 = st.columns([1, 2])
        with rc1:
            label_col = non_numeric_cols[0] if non_numeric_cols else None
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
                st.plotly_chart(fig_radar, width='stretch')

        st.divider()

        # 3D Explorer
        st.subheader("3D Explorer")
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

            fig_3d = px.scatter_3d(
                df, x=x3d, y=y3d, z=z3d,
                color=None if color3d == "(none)" else color3d,
                title=f"{x3d} vs {y3d} vs {z3d}",
                opacity=0.85,
                color_continuous_scale=CONT_SCALE if color3d != "(none)" and pd.api.types.is_numeric_dtype(df[color3d]) else None,
                color_discrete_sequence=COLOR_SEQUENCE if color3d == "(none)" or not pd.api.types.is_numeric_dtype(df[color3d]) else None,
            )
            fig_3d.update_layout(scene=dict(xaxis_title=x3d, yaxis_title=y3d, zaxis_title=z3d), height=650)
            fig_3d = apply_plotly_theme(fig_3d, dark_mode)
            st.plotly_chart(fig_3d, width='stretch')

    # -------------------- TAB: STATS & CORRELATION --------------------
    with tab_stats:
        st.subheader(f"Statistical Summary & Subject Weakness Analysis: {primary_label}")
        if numeric_cols:
            st.dataframe(df[numeric_cols].describe().T, width='stretch')

        st.markdown("#### Correlation Heatmap")
        if len(numeric_cols) >= 2:
            corr = df[numeric_cols].corr(numeric_only=True)
            fig_corr = px.imshow(
                corr, text_auto=".2f", aspect="auto",
                color_continuous_scale=CONT_SCALE, title="Correlation Matrix"
            )
            fig_corr = apply_plotly_theme(fig_corr, dark_mode)
            st.plotly_chart(fig_corr, width='stretch')

        st.markdown("#### Distribution Explorer")
        if numeric_cols:
            dist_col = st.selectbox("Column to inspect", numeric_cols, key="dist_col")
            fig_dist = px.histogram(
                df, x=dist_col, marginal="box", title=f"Distribution of {dist_col}",
                color_discrete_sequence=COLOR_SEQUENCE
            )
            fig_dist = apply_plotly_theme(fig_dist, dark_mode)
            st.plotly_chart(fig_dist, width='stretch')

    # -------------------- TAB: PERSONAL ANALYSIS --------------------
    with tab_personal:
        st.subheader("Personal Analysis (Student Score Breakdown)")
        if non_numeric_cols:
            name_col = st.selectbox("Which column holds names/entities?", non_numeric_cols, key="personal_name_col")
            search_name = st.text_input("Enter a student name to analyze", key="personal_search")
        else:
            name_col, search_name = None, None

        if name_col and search_name:
            student_row, match_count = find_student_row(df, name_col, search_name)
            if student_row is not None:
                st.success(f"Found: {student_row[name_col]}")
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
                st.dataframe(comparison_df, width='stretch')

                chart_data = comparison_df.melt(
                    id_vars="Subject", value_vars=["Score", "Avg of Those Above"],
                    var_name="Metric", value_name="Value"
                )
                fig_personal = px.bar(
                    chart_data, x="Subject", y="Value", color="Metric", barmode="group",
                    title=f"{student_row[name_col]}'s Scores vs. Avg of Higher Performers",
                    color_discrete_sequence=COLOR_SEQUENCE
                )
                fig_personal = apply_plotly_theme(fig_personal, dark_mode)
                st.plotly_chart(fig_personal, width='stretch')

    # -------------------- TAB: HEAD-TO-HEAD COMPARISON --------------------
    with tab_compare:
        st.subheader("Head-to-Head Comparison")
        if non_numeric_cols:
            vs_name_col = st.selectbox("Name column", non_numeric_cols, key="vs_name_col")
            entity_options = df[vs_name_col].astype(str).tolist()

            colA, colB = st.columns(2)
            with colA:
                person_a = st.selectbox("First Student / Entity", entity_options, key="vs_a")
            with colB:
                default_b = 1 if len(entity_options) > 1 else 0
                person_b = st.selectbox("Second Student / Entity", entity_options, index=default_b, key="vs_b")

            if person_a and person_b and person_a != person_b:
                row_a, _ = find_student_row(df, vs_name_col, person_a)
                row_b, _ = find_student_row(df, vs_name_col, person_b)

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
                    st.plotly_chart(fig_vs, width='stretch')

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
                    st.dataframe(diff_df, width='stretch')

    # -------------------- TAB: MULTI-FILE GROWTH --------------------
    with tab_growth:
        if len(named_dfs) < 2:
            st.info("Upload 2 or more files (or a multi-table PDF) to track multi-file growth.")
        else:
            st.subheader("File Compatibility Check")
            compat = check_file_compatibility(named_dfs)
            if compat["all_match"]:
                st.success("✅ All files share identical columns.")
                use_cols = compat["common_cols"]
            else:
                use_cols = st.multiselect(
                    "Shared columns to use", compat["common_cols"], default=compat["common_cols"]
                )

            if use_cols:
                semester_dfs = []
                for i, (label, tdf) in enumerate(named_dfs):
                    sem_df = tdf[use_cols].copy()
                    sem_df['__Semester__'] = f"{i + 1}. {label}"
                    semester_dfs.append(sem_df)

                combined_df = pd.concat(semester_dfs, ignore_index=True)
                name_cols_multi = [c for c in combined_df.select_dtypes(exclude=np.number).columns if c != '__Semester__']

                if name_cols_multi:
                    name_col_multi = st.selectbox("Shared Name Column", name_cols_multi, key="multi_name_col")
                    numeric_cols_multi = combined_df.select_dtypes(include=np.number).columns.tolist()

                    if numeric_cols_multi:
                        long_df = combined_df.melt(
                            id_vars=['__Semester__', name_col_multi], value_vars=numeric_cols_multi,
                            var_name='Subject', value_name='Score'
                        ).rename(columns={name_col_multi: 'Name'})

                        search_name_multi = st.text_input("Student Name for Growth Trend", key="multi_search_name")
                        if search_name_multi:
                            student_long = long_df[long_df['Name'].astype(str).str.strip().str.lower() == search_name_multi.strip().lower()]
                            if not student_long.empty:
                                fig_trend = px.line(
                                    student_long, x='__Semester__', y='Score', color='Subject', markers=True,
                                    title=f"{search_name_multi}'s Multi-Semester Growth",
                                    color_discrete_sequence=COLOR_SEQUENCE
                                )
                                fig_trend = apply_plotly_theme(fig_trend, dark_mode)
                                st.plotly_chart(fig_trend, width='stretch')

    # -------------------- TAB: DATA CLEANING --------------------
    with tab_clean:
        st.subheader(f"Data Cleaning: {primary_label}")
        oc1, oc2 = st.columns(2)
        with oc1:
            drop_dupes = st.checkbox("Remove duplicate rows", value=False)
            strip_whitespace = st.checkbox("Trim whitespace in text columns", value=True)
            drop_empty_cols = st.checkbox("Drop fully-empty columns", value=False)
        with oc2:
            fillna_strategy = st.selectbox("Fill missing values using…", ["None", "Mean", "Median", "Zero"])
            fillna_cols = []
            if fillna_strategy != "None":
                fillna_cols = st.multiselect("Columns to fill", df.columns.tolist(), default=[c for c in numeric_cols if df[c].isna().any()])

        cleaned_df, change_log = clean_dataframe(df, drop_dupes, fillna_strategy, fillna_cols, drop_empty_cols, strip_whitespace)
        for entry in change_log:
            st.markdown(f"<div class='dv-insight'>✅ {entry}</div>", unsafe_allow_html=True)

        st.dataframe(cleaned_df, width='stretch')
        st.download_button("⬇️ Download cleaned CSV", cleaned_df.to_csv(index=False).encode("utf-8"), file_name="cleaned_data.csv")

    # -------------------- TAB: EXPORT REPORT --------------------
    with tab_export:
        st.subheader("Export Combined Multi-Sheet Excel Report")
        if EXCEL_EXPORT_SUPPORT:
            report_buffer = build_excel_report(named_dfs, primary_label, df, numeric_cols)
            st.download_button(
                "⬇️ Download Excel report (.xlsx)",
                report_buffer,
                file_name=f"DATAVIS-V5_report_{datetime.date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # -------------------- TAB: AI INSIGHTS --------------------
    with tab_ai:
        st.subheader(f"AI-Generated Insights: {primary_label}")
        if api_key:
            include_raw = st.checkbox("Include row-level data for entity callouts", value=False)
            if st.button('Generate AI Summary'):
                with st.spinner('Analyzing data...'):
                    summary, error = get_ai_summary(df, api_key, name_col=non_numeric_cols[0] if (include_raw and non_numeric_cols) else None, include_raw_rows=include_raw)
                if error:
                    st.error(f"Couldn't generate summary: {error}")
                else:
                    st.markdown(summary)
        else:
            st.info('Enter your Claude API key in the sidebar to generate an AI summary.')

else:
    with tab_dash:
        st.info('Upload a CSV, Excel, or PDF file in the sidebar (or click "Load Sample Student Marks Dataset") to get started.')
    with tab_grades:
        st.info('Upload a dataset to see the Leaderboard and Grade breakdown.')
    with tab_charts:
        st.info('Upload a file to build custom charts.')
