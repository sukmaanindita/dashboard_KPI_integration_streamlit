from __future__ import annotations

from io import BytesIO
import sys
from pathlib import Path
import calendar
import re

import pandas as pd
import plotly.express as px
import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from job_report.exporter import build_excel_report
from job_report.history import (
    compare_snapshots,
    group_active_trend,
    history_trend,
    job_change_details,
    list_periods,
    reset_manual_metric,
    save_manual_metrics,
    save_snapshot,
)
from job_report.processor import load_job_report


st.set_page_config(page_title="RM Report Job", page_icon="📊", layout="wide")
st.title("⭐ Integration Team Job Reporting Dashboard ⭐")
st.caption("Monthly Report Job Release Management - Integration Team")

default_files = sorted((ROOT / "data" / "source").glob("*.xlsx"))
uploaded = st.sidebar.file_uploader("Upload workbook", type=["xlsx"])
source = BytesIO(uploaded.getvalue()) if uploaded else (default_files[-1] if default_files else None)
source_name = uploaded.name if uploaded else (source.name if isinstance(source, Path) else "")

if source is None:
    st.info("Upload file Excel report untuk mulai menganalisis data.")
    st.stop()

try:
    report = load_job_report(source)
except Exception as exc:
    st.error(f"Workbook tidak dapat diproses: {exc}")
    st.stop()

summary = report.summary.copy()

filename_match = re.search(r"(?<!\d)(0[1-9]|1[0-2])\.", source_name)
default_month = int(filename_match.group(1)) if filename_match else 1
year_match = re.search(r"\b(20\d{2})\b", source_name)
default_year = int(year_match.group(1)) if year_match else 2026
month_options = list(range(1, 13))
selected_month = st.sidebar.selectbox(
    "Bulan report",
    month_options,
    index=default_month - 1,
    format_func=lambda month: f"{month:02d} - {calendar.month_name[month]}",
)
selected_year = st.sidebar.number_input(
    "Tahun report", min_value=2000, max_value=2100, value=default_year, step=1
)
reporting_period = f"{int(selected_year):04d}-{selected_month:02d}"
report_filename = (
    f"{selected_month:02d}. Report All Job ICT "
    f"{calendar.month_name[selected_month]} {int(selected_year)}_web.xlsx"
)
dynamic_report_filename = report_filename.replace("_web.xlsx", "_web_dynamic.xlsx")
history_dir = ROOT / "data" / "history"

st.sidebar.caption(f"Nama file: {report_filename}")
if st.sidebar.button("Simpan/Update Historical Snapshot", use_container_width=True):
    saved_path = save_snapshot(summary, reporting_period, history_dir)
    st.sidebar.success(f"Snapshot {reporting_period} tersimpan: {saved_path.name}")
periods = st.sidebar.multiselect("Period", sorted(summary["period"].unique()), default=sorted(summary["period"].unique()))
sources = st.sidebar.multiselect("Source", sorted(summary["source"].unique()), default=sorted(summary["source"].unique()))
priorities = st.sidebar.multiselect("Priority", sorted(summary["priority_name"].unique()), default=sorted(summary["priority_name"].unique()))
groups = st.sidebar.multiselect("Group job", sorted(summary["group_job"].unique()), default=sorted(summary["group_job"].unique()))

filtered = summary[
    summary["period"].isin(periods)
    & summary["source"].isin(sources)
    & summary["priority_name"].isin(priorities)
    & summary["group_job"].isin(groups)
]

saved_periods = list_periods(history_dir)
top_active = len(summary)
top_error = int(
    (summary["retry_lt_3h"].fillna(0) + summary["retry_gt_3h"].fillna(0)).gt(0).sum()
)
top_new = 0
top_disabled = 0
if reporting_period in saved_periods:
    top_trend = history_trend(history_dir)
    top_row = top_trend[top_trend["period"].eq(reporting_period)].iloc[0]
    top_active = int(top_row["active_jobs"])
    top_error = int(top_row["error_jobs"])
    top_new = int(top_row["new_jobs"])
    top_disabled = int(top_row["disabled_jobs"])

cols = st.columns(4)
cols[0].metric("Job Aktif", f"{top_active:,}")
cols[1].metric("Job Error", f"{top_error:,}")
cols[2].metric("Job Baru", f"{top_new:,}")
cols[3].metric("Job Disable", f"{top_disabled:,}")
if reporting_period not in saved_periods:
    st.caption(
        "Simpan historical snapshot untuk menghitung Job Baru dan Job Disable "
        "serta menerapkan adjustment manual pada KPI."
    )

tab_overview, tab_history, tab_detail, tab_quality = st.tabs(
    ["Overview", "Historical Analysis", "Job Detail", "Data Quality"]
)
with tab_overview:
    group_trend = group_active_trend(
        history_dir,
        int(selected_year),
        selected_month,
    )
    overview_period_title = (
        f"Periode January - {calendar.month_name[selected_month]} {int(selected_year)}"
    )
    st.subheader("All Job Active")
    st.caption(overview_period_title)
    if group_trend.empty:
        st.info(
            "Belum ada snapshot historical pada rentang ini. Simpan snapshot setiap bulan "
            "untuk menampilkan grafik job aktif per group."
        )
    else:
        group_trend = group_trend.copy()
        group_trend["month"] = group_trend["period"].map(
            lambda value: calendar.month_abbr[int(value[-2:])].upper()
        )
        month_order = [
            calendar.month_abbr[month].upper()
            for month in range(1, selected_month + 1)
            if f"{int(selected_year):04d}-{month:02d}" in set(group_trend["period"])
        ]
        group_order = sorted(group_trend["group_job"].astype(str).unique())
        active_table = (
            group_trend.pivot_table(
                index="month",
                columns="group_job",
                values="active_jobs",
                aggfunc="sum",
                fill_value=0,
            )
            .reindex(index=month_order, columns=group_order, fill_value=0)
        )
        active_table.index.name = "Bulan"

        trend_scope = st.selectbox(
            "Tampilkan tren",
            ["ALL JOB", *group_order],
            key="overview_active_job_scope",
        )
        if trend_scope == "ALL JOB":
            trend_values = active_table.sum(axis=1)
            chart_title = "All Job Active"
        else:
            trend_values = active_table[trend_scope]
            chart_title = f"Active Job — {trend_scope}"

        active_trend = trend_values.rename("active_jobs").reset_index()
        active_trend["change"] = active_trend["active_jobs"].diff()
        active_trend["point_label"] = active_trend.apply(
            lambda row: (
                f"{int(row['active_jobs'])}"
                if pd.isna(row["change"])
                else f"{int(row['active_jobs'])} ({int(row['change']):+d})"
            ),
            axis=1,
        )
        minimum = float(active_trend["active_jobs"].min())
        maximum = float(active_trend["active_jobs"].max())
        padding = max((maximum - minimum) * 0.35, 1.0)

        active_group_chart = px.line(
            active_trend,
            x="Bulan",
            y="active_jobs",
            markers=True,
            text="point_label",
            labels={"active_jobs": "Jumlah Job Aktif"},
            title=f"{chart_title}<br><sup>{overview_period_title}</sup>",
        )
        active_group_chart.update_traces(
            line={"width": 4, "color": "#2F75B5"},
            marker={"size": 11, "color": "#2F75B5"},
            textposition="top center",
            fill="tozeroy",
            fillcolor="rgba(47, 117, 181, 0.10)",
        )
        active_group_chart.update_layout(
            height=480,
            xaxis={"type": "category"},
            yaxis={"range": [max(0, minimum - padding), maximum + padding]},
            showlegend=False,
            title={"x": 0.5, "xanchor": "center"},
        )
        st.plotly_chart(active_group_chart, use_container_width=True)

        st.caption("All Job Active Detail per Group")
        st.dataframe(active_table, use_container_width=True)

        with st.expander("Detail Perubahan Job (+ / -)", expanded=False):
            if not saved_periods:
                st.info(
                    "Belum ada snapshot historical yang bisa dibandingkan."
                )
            else:
                change_periods = sorted(saved_periods, reverse=False)
                start_period = st.selectbox(
                    "Periode awal",
                    change_periods,
                    format_func=lambda value: pd.Period(value).strftime("%B %Y"),
                    key="overview_change_start_period",
                )
                end_period = st.selectbox(
                    "Periode akhir",
                    change_periods,
                    format_func=lambda value: pd.Period(value).strftime("%B %Y"),
                    key="overview_change_end_period",
                )
                if end_period < start_period:
                    start_period, end_period = end_period, start_period

                comparison = compare_snapshots(
                    history_dir,
                    end_period,
                    previous_period=start_period,
                )
                st.caption(f"Perbandingan {start_period} terhadap {end_period}")

                summary_frames = []
                for change_name, data in [
                    ("Job Baru", comparison.new_jobs),
                    ("Job Disable", comparison.disabled_jobs),
                    ("Job Error", comparison.error_jobs),
                ]:
                    if data.empty:
                        continue
                    counts = (
                        data.groupby(["group_job"], dropna=False)
                        .size()
                        .reset_index(name="job_count")
                    )
                    counts["change_label"] = change_name
                    summary_frames.append(counts)
                if summary_frames:
                    change_summary = pd.concat(summary_frames, ignore_index=True)
                    change_chart = px.bar(
                        change_summary,
                        x="group_job",
                        y="job_count",
                        color="change_label",
                        barmode="group",
                        text_auto=True,
                        color_discrete_map={
                            "Job Baru": "#70AD47",
                            "Job Disable": "#C00000",
                            "Job Error": "#FFC000",
                        },
                        labels={
                            "group_job": "Group Job",
                            "job_count": "Jumlah Job",
                            "change_label": "Perubahan",
                        },
                        title="Perubahan Job per Group",
                    )
                    change_chart.update_layout(
                        xaxis_tickangle=-45,
                        yaxis={"dtick": 1, "rangemode": "tozero"},
                        legend_title_text="Perubahan",
                    )
                    st.plotly_chart(change_chart, use_container_width=True)

                    change_table = (
                        change_summary.pivot_table(
                            index="group_job",
                            columns="change_label",
                            values="job_count",
                            aggfunc="sum",
                            fill_value=0,
                        )
                        .reindex(columns=["Job Baru", "Job Disable", "Job Error"], fill_value=0)
                        .reset_index()
                    )
                    change_table["Net Change"] = (
                        change_table["Job Baru"] - change_table["Job Disable"]
                    )
                    change_table = change_table.sort_values(
                        ["Job Baru", "Job Disable", "Job Error"], ascending=False
                    )
                    st.dataframe(change_table, use_container_width=True, hide_index=True)
                else:
                    st.info("Tidak ada perubahan yang dapat ditampilkan untuk rentang yang dipilih.")

                st.markdown("**Detail Nama Job**")
                if trend_scope != "ALL JOB":
                    st.caption(f"Detail difilter untuk group: {trend_scope}")

                change_columns = [
                    "job_name", "source", "group_job", "priority_name"
                ]
                if "period_type" in comparison.new_jobs.columns:
                    change_columns.append("period_type")
                new_changes = comparison.new_jobs.copy()
                disabled_changes = comparison.disabled_jobs.copy()
                error_changes = comparison.error_jobs.copy()

                if trend_scope != "ALL JOB":
                    new_changes = new_changes[new_changes["group_job"].eq(trend_scope)]
                    disabled_changes = disabled_changes[disabled_changes["group_job"].eq(trend_scope)]
                    error_changes = error_changes[error_changes["group_job"].eq(trend_scope)]

                new_col, disabled_col, error_col = st.columns(3)
                with new_col:
                    st.markdown(f"**+ Job Baru ({len(new_changes)})**")
                    st.dataframe(
                        new_changes[[col for col in change_columns if col in new_changes.columns]],
                        use_container_width=True,
                        hide_index=True,
                    )
                with disabled_col:
                    st.markdown(f"**− Job Disable ({len(disabled_changes)})**")
                    st.dataframe(
                        disabled_changes[[col for col in change_columns if col in disabled_changes.columns]],
                        use_container_width=True,
                        hide_index=True,
                    )
                with error_col:
                    st.markdown(f"**⚠ Job Error ({len(error_changes)})**")
                    st.dataframe(
                        error_changes[[col for col in change_columns if col in error_changes.columns]],
                        use_container_width=True,
                        hide_index=True,
                    )

    st.subheader("Current Report Success by Group and Priority")
    grouped = (
        filtered.groupby(["group_job", "priority_name"], as_index=False)
        .agg(
            job_count=("job_name", "count"),
            periodic_job=("periodic_job", "sum"),
            success=("success", "sum"),
            retry_lt_3h=("retry_lt_3h", "sum"),
            retry_gt_3h=("retry_gt_3h", "sum"),
            success_pct=("success_pct", "mean"),
        )
    )
    if not grouped.empty:
        grouped["success_pct"] = grouped["success_pct"].round(2)
        chart = px.bar(
            grouped,
            x="group_job",
            y="success_pct",
            color="priority_name",
            barmode="group",
            labels={"group_job": "Group Job", "success_pct": "Success (%)", "priority_name": "Priority"},
        )
        chart.update_layout(yaxis_range=[0, 100], xaxis_tickangle=-35)
        st.plotly_chart(chart, use_container_width=True)
        st.dataframe(grouped.sort_values("success_pct"), use_container_width=True, hide_index=True)
    else:
        st.info("Tidak ada data untuk kombinasi filter ini.")

with tab_detail:
    st.dataframe(filtered, use_container_width=True, hide_index=True)

with tab_quality:
    st.metric("Validation Issues", len(report.validation))
    if report.validation.empty:
        st.success("Tidak ditemukan masalah validasi.")
    else:
        st.dataframe(report.validation, use_container_width=True, hide_index=True)

with tab_history:
    periods = list_periods(history_dir)
    if not periods:
        st.info("Belum ada snapshot. Pilih bulan dan klik 'Simpan/Update Historical Snapshot'.")
    else:
        default_index = periods.index(reporting_period) if reporting_period in periods else len(periods) - 1
        analysis_period = st.selectbox(
            "Bulan yang dianalisis",
            periods,
            index=default_index,
            format_func=lambda value: pd.Period(value).strftime("%B %Y"),
        )
        comparison = compare_snapshots(history_dir, analysis_period)
        previous_label = comparison.previous_period or "belum tersedia"
        trend = history_trend(history_dir)
        selected_metric = trend[trend["period"].eq(analysis_period)].iloc[0]
        previous_metrics = trend[trend["period"].lt(analysis_period)].tail(1)
        previous_error_count = (
            int(previous_metrics.iloc[0]["error_jobs"])
            if not previous_metrics.empty
            else 0
        )
        error_delta = int(selected_metric["error_jobs"]) - previous_error_count

        with st.expander("Manual Historical Adjustment", expanded=False):
            st.caption(
                "Nilai tersimpan akan digunakan oleh KPI dan bar chart. "
                "Tabel detail tetap menggunakan hasil perbandingan snapshot."
            )
            with st.form(f"manual_metrics_{analysis_period}"):
                adjustment_cols = st.columns(4)
                manual_active = adjustment_cols[0].number_input(
                    "Job Aktif",
                    min_value=0,
                    value=int(selected_metric["active_jobs"]),
                    step=1,
                )
                manual_error = adjustment_cols[1].number_input(
                    "Job Error",
                    min_value=0,
                    value=int(selected_metric["error_jobs"]),
                    step=1,
                )
                manual_disabled = adjustment_cols[2].number_input(
                    "Job Disable",
                    min_value=0,
                    value=int(selected_metric["disabled_jobs"]),
                    step=1,
                )
                manual_new = adjustment_cols[3].number_input(
                    "Job Baru",
                    min_value=0,
                    value=int(selected_metric["new_jobs"]),
                    step=1,
                )
                action_cols = st.columns(2)
                save_adjustment = action_cols[0].form_submit_button(
                    "Simpan Adjustment", use_container_width=True
                )
                reset_error = action_cols[1].form_submit_button(
                    "Reset Job Error", use_container_width=True
                )
            if save_adjustment:
                save_manual_metrics(
                    history_dir,
                    analysis_period,
                    active_jobs=manual_active,
                    error_jobs=manual_error,
                    disabled_jobs=manual_disabled,
                    new_jobs=manual_new,
                )
                st.success(f"Adjustment {analysis_period} berhasil disimpan.")
                st.rerun()
            if reset_error:
                reset_manual_metric(history_dir, analysis_period, "error_jobs")
                st.success(
                    f"Job Error {analysis_period} dikembalikan ke hasil kalkulasi snapshot."
                )
                st.rerun()

        history_cols = st.columns(4)
        history_cols[0].metric("Job Aktif", int(selected_metric["active_jobs"]))
        history_cols[1].metric("Job Baru", int(selected_metric["new_jobs"]))
        history_cols[2].metric("Job Disable", int(selected_metric["disabled_jobs"]))
        history_cols[3].metric(
            "Error Job",
            int(selected_metric["error_jobs"]),
            error_delta,
            delta_color="inverse",
        )
        st.caption(f"Perbandingan {analysis_period} terhadap {previous_label}")

        trend = trend[trend["period"].le(analysis_period)].tail(3)
        if not trend.empty:
            comparison_chart_data = trend.melt(
                id_vars="period",
                value_vars=["error_jobs", "active_jobs", "disabled_jobs", "new_jobs"],
                var_name="metric",
                value_name="job_count",
            )
            metric_names = {
                "error_jobs": "Job Error",
                "active_jobs": "Job Aktif",
                "disabled_jobs": "Job Disable",
                "new_jobs": "Job Baru",
            }
            comparison_chart_data["metric"] = comparison_chart_data["metric"].map(metric_names)
            trend_chart = px.bar(
                comparison_chart_data,
                x="period",
                y="job_count",
                color="metric",
                barmode="group",
                text_auto=True,
                labels={"job_count": "Jumlah Job", "period": "Bulan", "metric": "Variable"},
                title="Three-Month Job Comparison",
            )
            trend_chart.update_layout(xaxis={"type": "category"})
            st.plotly_chart(trend_chart, use_container_width=True)

        new_tab, disabled_tab, error_tab = st.tabs(
            ["Job Baru", "Job Disable", "Error Job"]
        )
        detail_columns = [
            "job_name", "source", "period", "group_job", "priority_name",
            "periodic_job", "success_pct", "retry_lt_3h", "retry_gt_3h",
        ]
        with new_tab:
            st.dataframe(
                comparison.new_jobs[[c for c in detail_columns if c in comparison.new_jobs]],
                use_container_width=True,
                hide_index=True,
            )
        with disabled_tab:
            st.dataframe(
                comparison.disabled_jobs[[c for c in detail_columns if c in comparison.disabled_jobs]],
                use_container_width=True,
                hide_index=True,
            )
        with error_tab:
            st.dataframe(
                comparison.error_jobs[[c for c in detail_columns if c in comparison.error_jobs]],
                use_container_width=True,
                hide_index=True,
            )

st.sidebar.download_button(
    "Download Excel Report — Statis",
    data=build_excel_report(report, source),
    file_name=report_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.sidebar.download_button(
    "Download Excel Report — Dinamis",
    data=build_excel_report(report, source, dynamic=True),
    file_name=dynamic_report_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
