from pathlib import Path
import sys

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from job_report.history import (
    compare_snapshots,
    group_active_trend,
    history_trend,
    job_change_details,
    list_periods,
    load_manual_metrics,
    reset_manual_metric,
    save_manual_metrics,
    save_snapshot,
)


def sample_summary(job_names, retries=None):
    retries = retries or {}
    rows = []
    for name in job_names:
        retry = retries.get(name, 0)
        rows.append(
            {
                "job_name": name,
                "source": "SSIS",
                "period": "DAILY",
                "group_job": "TEST GROUP",
                "priority": "H",
                "priority_name": "HIGH",
                "periodic_job": 31,
                "success": 31 - retry,
                "retry_lt_3h": retry,
                "retry_gt_3h": 0,
                "success_pct": (31 - retry) / 31 * 100,
                "retry_lt_3h_pct": retry / 31 * 100,
                "retry_gt_3h_pct": 0,
                "completed_pct": 100,
            }
        )
    return pd.DataFrame(rows)


def test_history_detects_new_disabled_and_error_jobs(tmp_path):
    save_snapshot(sample_summary(["JOB_A", "JOB_B"]), "2026-06", tmp_path)
    save_snapshot(
        sample_summary(["JOB_B", "JOB_C"], retries={"JOB_B": 1}),
        "2026-07",
        tmp_path,
    )

    comparison = compare_snapshots(tmp_path, "2026-07")
    assert comparison.previous_period == "2026-06"
    assert set(comparison.new_jobs["job_name"]) == {"JOB_C"}
    assert set(comparison.disabled_jobs["job_name"]) == {"JOB_A"}
    assert set(comparison.error_jobs["job_name"]) == {"JOB_B"}


def test_history_periods_and_trend_are_chronological(tmp_path):
    save_snapshot(sample_summary(["JOB_A"]), "2026-07", tmp_path)
    save_snapshot(sample_summary(["JOB_A", "JOB_B"]), "2026-08", tmp_path)

    assert list_periods(tmp_path) == ["2026-07", "2026-08"]
    trend = history_trend(tmp_path)
    assert trend["period"].tolist() == ["2026-07", "2026-08"]
    assert trend["active_jobs"].tolist() == [1, 2]
    assert trend["new_jobs"].tolist() == [0, 1]
    assert trend["disabled_jobs"].tolist() == [0, 0]


def test_group_active_trend_starts_in_january_and_stops_at_selected_month(tmp_path):
    save_snapshot(sample_summary(["JOB_A"]), "2025-12", tmp_path)
    save_snapshot(sample_summary(["JOB_A", "JOB_B"]), "2026-01", tmp_path)
    save_snapshot(sample_summary(["JOB_A"]), "2026-03", tmp_path)
    save_snapshot(sample_summary(["JOB_A"]), "2026-08", tmp_path)

    trend = group_active_trend(tmp_path, 2026, 3)
    assert trend["period"].unique().tolist() == ["2026-01", "2026-03"]
    assert trend.groupby("period")["active_jobs"].sum().to_dict() == {
        "2026-01": 2,
        "2026-03": 1,
    }


def test_manual_metrics_override_historical_chart_values(tmp_path):
    save_snapshot(sample_summary(["JOB_A", "JOB_B"]), "2026-07", tmp_path)
    save_manual_metrics(
        tmp_path,
        "2026-07",
        active_jobs=350,
        error_jobs=12,
        disabled_jobs=4,
        new_jobs=8,
    )

    stored = load_manual_metrics(tmp_path).iloc[0]
    assert stored["error_jobs"] == 12
    trend = history_trend(tmp_path).iloc[0]
    assert trend["active_jobs"] == 350
    assert trend["error_jobs"] == 12
    assert trend["disabled_jobs"] == 4
    assert trend["new_jobs"] == 8

    reset_manual_metric(tmp_path, "2026-07", "error_jobs")
    reset_trend = history_trend(tmp_path).iloc[0]
    assert reset_trend["error_jobs"] == 0
    assert reset_trend["active_jobs"] == 350


def test_job_change_details_lists_added_and_disabled_names(tmp_path):
    save_snapshot(sample_summary(["JOB_A", "JOB_B"]), "2026-06", tmp_path)
    save_snapshot(sample_summary(["JOB_B", "JOB_C"]), "2026-07", tmp_path)

    changes = job_change_details(tmp_path, 2026, 7)
    added = changes[changes["change"].eq("+ Job Baru")]
    disabled = changes[changes["change"].eq("- Job Disable")]
    assert set(added["job_name"]) == {"JOB_C"}
    assert set(disabled["job_name"]) == {"JOB_A"}
    assert set(changes["previous_period"]) == {"2026-06"}
