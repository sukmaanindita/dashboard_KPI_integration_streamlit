from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re

import pandas as pd

from .processor import canonical_group


PERIOD_PATTERN = re.compile(r"^(\d{4})-(0[1-9]|1[0-2])$")
MANUAL_METRICS_FILE = "manual_history_metrics.csv"
METRIC_COLUMNS = ["active_jobs", "error_jobs", "disabled_jobs", "new_jobs"]


@dataclass(frozen=True)
class HistoryComparison:
    current_period: str
    previous_period: str | None
    active_jobs: pd.DataFrame
    new_jobs: pd.DataFrame
    disabled_jobs: pd.DataFrame
    error_jobs: pd.DataFrame
    previous_error_jobs: pd.DataFrame


def validate_period(period: str) -> str:
    if not PERIOD_PATTERN.match(period):
        raise ValueError("Period history harus menggunakan format YYYY-MM.")
    return period


def history_path(history_dir: str | Path, period: str) -> Path:
    validate_period(period)
    return Path(history_dir) / f"job_history_{period}.csv"


def save_snapshot(summary: pd.DataFrame, period: str, history_dir: str | Path) -> Path:
    path = history_path(history_dir, period)
    path.parent.mkdir(parents=True, exist_ok=True)
    snapshot = summary.copy()
    snapshot.insert(0, "reporting_period", period)
    snapshot.insert(1, "saved_at", datetime.now().isoformat(timespec="seconds"))
    snapshot.to_csv(path, index=False)
    return path


def load_snapshot(history_dir: str | Path, period: str) -> pd.DataFrame:
    path = history_path(history_dir, period)
    if not path.exists():
        raise FileNotFoundError(f"Snapshot {period} belum tersimpan.")
    data = pd.read_csv(path)
    if "group_job" in data:
        data["group_job"] = data["group_job"].map(canonical_group)
    return data


def list_periods(history_dir: str | Path) -> list[str]:
    directory = Path(history_dir)
    if not directory.exists():
        return []
    periods = []
    for path in directory.glob("job_history_????-??.csv"):
        period = path.stem.removeprefix("job_history_")
        if PERIOD_PATTERN.match(period):
            periods.append(period)
    return sorted(set(periods))


def load_manual_metrics(history_dir: str | Path) -> pd.DataFrame:
    path = Path(history_dir) / MANUAL_METRICS_FILE
    if not path.exists():
        return pd.DataFrame(columns=["period", *METRIC_COLUMNS])
    data = pd.read_csv(path, dtype={"period": str})
    for column in METRIC_COLUMNS:
        if column not in data:
            data[column] = pd.NA
        data[column] = pd.to_numeric(data[column], errors="coerce")
    return data[["period", *METRIC_COLUMNS]].drop_duplicates("period", keep="last")


def save_manual_metrics(
    history_dir: str | Path,
    period: str,
    **metrics: int,
) -> Path:
    validate_period(period)
    unknown = set(metrics) - set(METRIC_COLUMNS)
    if unknown:
        raise ValueError(f"Metric manual tidak dikenal: {', '.join(sorted(unknown))}")
    values = {column: max(0, int(metrics[column])) for column in METRIC_COLUMNS}
    path = Path(history_dir) / MANUAL_METRICS_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    data = load_manual_metrics(history_dir)
    data = data[data["period"].ne(period)]
    data = pd.concat([data, pd.DataFrame([{"period": period, **values}])], ignore_index=True)
    data.sort_values("period").to_csv(path, index=False)
    return path


def reset_manual_metric(
    history_dir: str | Path,
    period: str,
    metric: str,
) -> Path:
    validate_period(period)
    if metric not in METRIC_COLUMNS:
        raise ValueError(f"Metric manual tidak dikenal: {metric}")
    path = Path(history_dir) / MANUAL_METRICS_FILE
    data = load_manual_metrics(history_dir)
    if period in set(data["period"]):
        data.loc[data["period"].eq(period), metric] = pd.NA
        data.sort_values("period").to_csv(path, index=False)
    return path


def _job_key(frame: pd.DataFrame) -> pd.Series:
    # Source is included because identical job names may exist on different engines.
    return frame["source"].fillna("").astype(str) + "||" + frame["job_name"].fillna("").astype(str)


def _errors(frame: pd.DataFrame) -> pd.DataFrame:
    retry_lt = pd.to_numeric(frame["retry_lt_3h"], errors="coerce").fillna(0)
    retry_gt = pd.to_numeric(frame["retry_gt_3h"], errors="coerce").fillna(0)
    return frame[(retry_lt + retry_gt).gt(0)].copy()


def compare_snapshots(
    history_dir: str | Path,
    current_period: str,
    previous_period: str | None = None,
) -> HistoryComparison:
    periods = list_periods(history_dir)
    current = load_snapshot(history_dir, current_period)
    if previous_period is None:
        earlier = [period for period in periods if period < current_period]
        previous_period = earlier[-1] if earlier else None

    previous = (
        load_snapshot(history_dir, previous_period)
        if previous_period
        else current.iloc[0:0].copy()
    )
    current_keys = set(_job_key(current))
    previous_keys = set(_job_key(previous))
    new_jobs = current[_job_key(current).isin(current_keys - previous_keys)].copy()
    disabled_jobs = previous[_job_key(previous).isin(previous_keys - current_keys)].copy()
    return HistoryComparison(
        current_period=current_period,
        previous_period=previous_period,
        active_jobs=current,
        new_jobs=new_jobs,
        disabled_jobs=disabled_jobs,
        error_jobs=_errors(current),
        previous_error_jobs=_errors(previous),
    )


def history_trend(history_dir: str | Path) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    periods = list_periods(history_dir)
    previous_data: pd.DataFrame | None = None
    for period in periods:
        data = load_snapshot(history_dir, period)
        periodic = pd.to_numeric(data["periodic_job"], errors="coerce").fillna(0).sum()
        success = pd.to_numeric(data["success"], errors="coerce").fillna(0).sum()
        if previous_data is None:
            new_count = 0
            disabled_count = 0
        else:
            current_keys = set(_job_key(data))
            previous_keys = set(_job_key(previous_data))
            new_count = len(current_keys - previous_keys)
            disabled_count = len(previous_keys - current_keys)
        rows.append(
            {
                "period": period,
                "active_jobs": len(data),
                "error_jobs": len(_errors(data)),
                "new_jobs": new_count,
                "disabled_jobs": disabled_count,
                "success_rate": round(success / periodic * 100, 2) if periodic else 0,
            }
        )
        previous_data = data
    trend = pd.DataFrame(rows)
    if trend.empty:
        return trend
    manual = load_manual_metrics(history_dir).set_index("period")
    for row_index, period in trend["period"].items():
        if period not in manual.index:
            continue
        for column in METRIC_COLUMNS:
            value = manual.loc[period, column]
            if pd.notna(value):
                trend.loc[row_index, column] = int(value)
    return trend


def group_active_trend(
    history_dir: str | Path,
    year: int,
    through_month: int,
) -> pd.DataFrame:
    """Count active jobs per group for saved snapshots from January onward."""
    end_period = f"{year:04d}-{through_month:02d}"
    rows: list[dict[str, object]] = []
    for period in list_periods(history_dir):
        if not period.startswith(f"{year:04d}-") or period > end_period:
            continue
        data = load_snapshot(history_dir, period)
        counts = data.groupby("group_job", dropna=False).size()
        for group, active_jobs in counts.items():
            rows.append(
                {
                    "period": period,
                    "group_job": group or "TANPA GROUP",
                    "active_jobs": int(active_jobs),
                }
            )
    return pd.DataFrame(rows)


def job_change_details(
    history_dir: str | Path,
    year: int,
    through_month: int,
) -> pd.DataFrame:
    """List added and disabled jobs between consecutive saved snapshots."""
    end_period = f"{year:04d}-{through_month:02d}"
    periods = [
        period
        for period in list_periods(history_dir)
        if period.startswith(f"{year:04d}-") and period <= end_period
    ]
    columns = [
        "period", "previous_period", "change", "job_name", "source",
        "group_job", "period_type", "priority_name",
    ]
    rows: list[dict[str, object]] = []
    for previous_period, current_period in zip(periods, periods[1:]):
        previous = load_snapshot(history_dir, previous_period)
        current = load_snapshot(history_dir, current_period)
        previous_keys = set(_job_key(previous))
        current_keys = set(_job_key(current))
        added = current[_job_key(current).isin(current_keys - previous_keys)]
        disabled = previous[_job_key(previous).isin(previous_keys - current_keys)]
        for change, data in (("+ Job Baru", added), ("- Job Disable", disabled)):
            for _, item in data.iterrows():
                rows.append(
                    {
                        "period": current_period,
                        "previous_period": previous_period,
                        "change": change,
                        "job_name": item.get("job_name"),
                        "source": item.get("source"),
                        "group_job": item.get("group_job"),
                        "period_type": item.get("period"),
                        "priority_name": item.get("priority_name"),
                    }
                )
    return pd.DataFrame(rows, columns=columns)
