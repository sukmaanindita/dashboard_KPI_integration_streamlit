from pathlib import Path
from io import BytesIO
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from job_report.processor import canonical_group, load_job_report
from job_report.exporter import build_excel_report
from openpyxl import load_workbook


SOURCE = ROOT / "data" / "source" / "07. Report All JOB ICT July 2026.xlsx"


@pytest.fixture(scope="module")
def report():
    return load_job_report(SOURCE)


def test_loads_all_jobs(report):
    assert len(report.summary) == 342
    assert set(report.summary["period"]) == {"DAILY", "MONTHLY"}


def test_status_totals_equal_periodic_job(report):
    calculated = report.summary[["success", "retry_lt_3h", "retry_gt_3h"]].sum(axis=1)
    assert calculated.equals(report.summary["periodic_job"])


def test_percentages_are_bounded(report):
    columns = ["success_pct", "retry_lt_3h_pct", "retry_gt_3h_pct", "completed_pct"]
    assert report.summary[columns].ge(0).all().all()
    assert report.summary[columns].le(100).all().all()


def test_source_names_are_normalized(report):
    assert "SSIs" not in set(report.summary["source"])


def test_group_ampersand_is_standardized_to_dan():
    assert canonical_group("Asuransi & Fiducia") == "ASURANSI DAN FIDUCIA"
    assert canonical_group("GROUP  A&B") == "GROUP A DAN B"


def test_group_pivot_uses_average_of_job_percentages(report):
    pivot = report.pivot("DAILY")
    row = pivot[
        pivot["group_job"].eq("ASURANSI DAN FIDUCIA")
        & pivot["priority"].eq("H")
    ].iloc[0]
    assert row["job_count"] == 11
    assert row["success_pct"] == pytest.approx(99.71)


def test_export_matches_requested_three_sheet_layout(report):
    output = build_excel_report(report, SOURCE)
    workbook = load_workbook(BytesIO(output), data_only=True, read_only=False)
    assert workbook.sheetnames == [
        "Daily", "Monthly", "Sum Job", "Result Detail JOB", "MSG_ERROR"
    ]
    detail = workbook["Result Detail JOB"]
    assert len(detail.merged_cells.ranges) == 0
    assert not detail.column_dimensions["B"].hidden
    assert not detail.column_dimensions["C"].hidden
    assert detail.auto_filter.ref is None
    assert len(detail.auto_filter.filterColumn) == 0
    assert detail.auto_filter.sortState is None
    assert not detail.sheet_properties.filterMode
    assert detail.freeze_panes is None
    assert len(detail.sheet_view.selection) == 1
    assert detail.sheet_view.selection[0].sqref == "A1"
    assert detail.max_row == 344

    sum_job = workbook["Sum Job"]
    assert len(sum_job.merged_cells.ranges) == 0
    assert sum_job.cell(3, 3).value == "JOB SUMMARY - AGING - DAILY"
    assert sum_job.cell(4, 3).value == "NAMA JOB"
    assert sum_job.cell(3, 17).value == "AGING - DAILY"
    assert sum_job.cell(4, 17).value is None
    assert sum_job.cell(5, 17).value == "DAILY"
    assert sum_job.cell(5, 18).value == pytest.approx(100.0)
    assert sum_job.max_column == 20

    asuransi_title_row = next(
        row for row in range(1, sum_job.max_row + 1)
        if sum_job.cell(row, 3).value
        == "JOB SUMMARY - ASURANSI DAN FIDUCIA - DAILY"
    )
    asuransi_jobs = []
    row = asuransi_title_row + 2
    while sum_job.cell(row, 3).value:
        asuransi_jobs.append(sum_job.cell(row, 3).value)
        row += 1
    assert asuransi_jobs == sorted(asuransi_jobs, key=str.upper)

    msg_error = workbook["MSG_ERROR"]
    assert [msg_error.cell(1, column).value for column in range(1, 7)] == [
        "No", "JOB_NAME", "GROUP", "ERROR_DATE", "ERROR_MSG", "KETERANGAN"
    ]


def test_export_daily_average_matches_excel_pivot(report):
    output = build_excel_report(report, SOURCE)
    workbook = load_workbook(BytesIO(output), data_only=True, read_only=True)
    sheet = workbook["Daily"]
    row = next(
        row for row in range(4, sheet.max_row + 1)
        if sheet.cell(row, 1).value == "ASURANSI DAN FIDUCIA"
    )
    assert sheet.cell(row, 5).value == pytest.approx(99.71)


def test_dynamic_export_contains_recalculating_excel_formulas(report):
    output = build_excel_report(report, SOURCE, dynamic=True)
    workbook = load_workbook(BytesIO(output), data_only=False, read_only=True)
    detail = workbook["Result Detail JOB"]
    assert detail["CU3"].value == "=SUM(F3:CT3)"
    assert detail["CY3"].value == "=IFERROR(CV3/CU3*100,0)"
    assert workbook["Daily"]["E5"].data_type == "f"
    assert workbook["Sum Job"]["G5"].data_type == "f"
    assert workbook.calculation.calcMode == "auto"
