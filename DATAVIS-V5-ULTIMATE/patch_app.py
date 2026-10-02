"""Patch app.py: expand CSS block with premium styles and enhance leaderboard tab."""
import re

with open("app.py", "r", encoding="utf-8") as f:
    content = f.read()

# ----------------------------------------------------------------
# 1. EXPAND CSS (add missing classes after .dv-auth-hero block)
# ----------------------------------------------------------------
OLD_CSS_END = """.dv-auth-hero {{
    background: linear-gradient(135deg, {card_bg}, {accent2}22);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 18px;
    padding: 35px;
    margin-bottom: 25px;
    text-align: center;
}}
</style>
\"\"\", unsafe_allow_html=True)"""

NEW_CSS_END = """.dv-auth-hero {{
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
\"\"\", unsafe_allow_html=True)"""

if OLD_CSS_END in content:
    content = content.replace(OLD_CSS_END, NEW_CSS_END, 1)
    print("✅ CSS expanded with premium styles")
else:
    print("❌ CSS block not found — skipping CSS patch")

# ----------------------------------------------------------------
# 2. ENHANCE LEADERBOARD TAB — add podium medals + subject toppers
# ----------------------------------------------------------------
OLD_LEADERBOARD = '''    # -------------------- TAB: LEADERBOARD & GRADES --------------------
    with tab_grades:
        st.subheader(f"🏆 Class Leaderboard & Grade Analytics ({primary_label})")

        if "Grade" in graded_df.columns:
            col_l1, col_l2 = st.columns([2, 1])

            with col_l1:
                st.markdown("#### 🥇 Class Rankings & Grade Breakdown")
                st.dataframe(graded_df, width=\'stretch\', height=450)

            with col_l2:
                st.markdown("#### 🍰 Grade Distribution")
                grade_counts = graded_df["Grade"].value_counts().reset_index()
                grade_counts.columns = ["Grade", "Count"]

                fig_grade_pie = px.pie(grade_counts, names="Grade", values="Count", title="Grade Share",
                                       color_discrete_sequence=COLOR_SEQUENCE)
                fig_grade_pie = apply_plotly_theme(fig_grade_pie, dark_mode)
                st.plotly_chart(fig_grade_pie, width=\'stretch\')
        else:
            st.info("No numeric metric columns detected to calculate grades.")'''

NEW_LEADERBOARD = '''    # -------------------- TAB: LEADERBOARD & GRADES --------------------
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
                    avg = f"{row[\'Average %\']:.1f}%"
                    grade_raw = str(row["Grade"])
                    grade_badge_map = {
                        "A+": "dv-badge-aplus", "A ": "dv-badge-a", "B ": "dv-badge-b",
                        "C ": "dv-badge-c", "D ": "dv-badge-d", "F ": "dv-badge-f"
                    }
                    badge_cls = next((v for k, v in grade_badge_map.items() if grade_raw.startswith(k.strip())), "dv-badge-b")
                    podium_col.markdown(
                        f"<div class=\'{medal_class}\'>"
                        f"<div style=\'font-size:36px;\'>{medal_emoji}</div>"
                        f"<div class=\'dv-medal-name\'>{nm}</div>"
                        f"<div class=\'dv-medal-score\'>{avg}</div>"
                        f"<span class=\'dv-badge {badge_cls}\'>{grade_raw}</span>"
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
                st.dataframe(display_df, width=\'stretch\', height=450)

            with col_l2:
                st.markdown("#### 🍰 Grade Distribution")
                grade_counts = graded_df["Grade"].value_counts().reset_index()
                grade_counts.columns = ["Grade", "Count"]
                fig_grade_pie = px.pie(grade_counts, names="Grade", values="Count", title="Grade Share",
                                       color_discrete_sequence=COLOR_SEQUENCE)
                fig_grade_pie = apply_plotly_theme(fig_grade_pie, dark_mode)
                st.plotly_chart(fig_grade_pie, width=\'stretch\')

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
                        f"<div class=\'dv-subject-topper\'>"
                        f"<div><b style=\'color:{accent2};\'>{subj}</b><br>"
                        f"<span style=\'font-size:15px; font-weight:700;\'>{top_name_val}</span></div>"
                        f"<div style=\'text-align:right;\'><span style=\'font-size:22px; font-weight:800;\'>{top_score}</span><br>"
                        f"<small style=\'color:{muted};\'>Top Score</small></div>"
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
            st.plotly_chart(fig_grade_bar, width=\'stretch\')

        else:
            st.info("No numeric metric columns detected to calculate grades.")'''

if OLD_LEADERBOARD in content:
    content = content.replace(OLD_LEADERBOARD, NEW_LEADERBOARD, 1)
    print("✅ Leaderboard tab enhanced with podium medals + subject toppers")
else:
    print("❌ Leaderboard block not found — skipping leaderboard patch")

with open("app.py", "w", encoding="utf-8") as f:
    f.write(content)

print("Patch complete!")
