from __future__ import annotations

from copy import copy
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import Selection

from .processor import JobReport, canonical_group


DETAIL_SHEET = "Result Detail JOB"
OUTPUT_SHEETS = ["Daily", "Monthly", "Sum Job", DETAIL_SHEET, "MSG_ERROR"]
PERCENT_HEADERS = [
    "Success Persentage",
    "Retry Success Rate < 3 Jam\nPersentage",
    "Retry Success Rate > 3 Jam\nPersentage",
]


def _template_bytes(source: str | Path | BinaryIO | BytesIO) -> BytesIO | str | Path:
    if isinstance(source, (str, Path)):
        return source
    source.seek(0)
    return BytesIO(source.read())


def _copy_cell_style(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    target.number_format = source.number_format
    target.alignment = copy(source.alignment)


def _update_detail_metrics(report: JobReport, sheet, dynamic: bool = False) -> None:
    by_job = report.summary.set_index("job_name")
    for row in range(3, sheet.max_row + 1):
        job_name = sheet.cell(row, 1).value
        if not job_name or job_name not in by_job.index:
            continue
        item = by_job.loc[job_name]
        sheet.cell(row, 4).value = item["group_job"]
        if dynamic:
            values = [
                f"=SUM(F{row}:CT{row})",
                f'=SUMIF(F$2:CT$2,"Success",F{row}:CT{row})',
                f'=SUMIF(F$2:CT$2,"Retry Success Rate < 3 Jam",F{row}:CT{row})',
                f'=SUMIF(F$2:CT$2,"Retry Success Rate > 3 Jam",F{row}:CT{row})',
                f"=IFERROR(CV{row}/CU{row}*100,0)",
                f"=IFERROR(CW{row}/CU{row}*100,0)",
                f"=IFERROR(CX{row}/CU{row}*100,0)",
            ]
        else:
            values = [
                item["periodic_job"],
                item["success"],
                item["retry_lt_3h"],
                item["retry_gt_3h"],
                item["success_pct"],
                item["retry_lt_3h_pct"],
                item["retry_gt_3h_pct"],
            ]
        for column, value in zip(range(99, 106), values):
            sheet.cell(row, column).value = value if dynamic else float(value)


def _remove_merged_cells_for_copy(sheet) -> None:
    """Replace merged headers with normal cells so Excel ranges stay contiguous."""
    for merged_range in list(sheet.merged_cells.ranges):
        min_col, min_row, max_col, max_row = merged_range.bounds
        anchor = sheet.cell(min_row, min_col)
        value = anchor.value
        anchor_style = copy(anchor._style)
        number_format = anchor.number_format
        fill = copy(anchor.fill)
        font = copy(anchor.font)
        border = copy(anchor.border)
        protection = copy(anchor.protection)
        sheet.unmerge_cells(str(merged_range))

        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                cell = sheet.cell(row, column)
                cell._style = copy(anchor_style)
                cell.number_format = number_format
                cell.fill = copy(fill)
                cell.font = copy(font)
                cell.border = copy(border)
                cell.protection = copy(protection)
                cell.alignment = Alignment(
                    horizontal="centerContinuous" if max_col > min_col else "center",
                    vertical="center",
                    wrap_text=True,
                )
                if row != min_row or column != min_col:
                    cell.value = None


def _reset_detail_selection_state(sheet, detail_last_row: int) -> None:
    """Remove source filter/hidden state that makes Excel copy visible areas separately."""
    for dimension in sheet.column_dimensions.values():
        dimension.hidden = False
    for dimension in sheet.row_dimensions.values():
        dimension.hidden = False

    # Remove the source AutoFilter object completely. Even with empty criteria,
    # Excel may retain filtered-mode metadata and treat pasted cells as areas.
    sheet.auto_filter.ref = None
    sheet.auto_filter.filterColumn = []
    sheet.auto_filter.sortState = None
    sheet.sheet_properties.filterMode = False

    # The source freezes at F3. A copy such as C207:I210 crosses the left and
    # right panes and Excel can treat it as multiple areas. Reset to one pane
    # and one active selection so every rectangular range stays contiguous.
    sheet.freeze_panes = None
    view = sheet.sheet_view
    view.pane = None
    view.selection = [Selection(activeCell="A1", sqref="A1")]


def _replace_with_clean_detail_sheet(workbook, source_sheet):
    """Strip legacy worksheet metadata while retaining values and appearance."""
    source_index = workbook.worksheets.index(source_sheet)
    clean = workbook.create_sheet("__clean_result_detail__", source_index)

    for row in source_sheet.iter_rows():
        for source_cell in row:
            if source_cell.value is None and not source_cell.has_style:
                continue
            target = clean.cell(source_cell.row, source_cell.column, source_cell.value)
            if source_cell.has_style:
                target._style = copy(source_cell._style)
            target.number_format = source_cell.number_format
            target.alignment = copy(source_cell.alignment)
            target.protection = copy(source_cell.protection)

    for key, dimension in source_sheet.column_dimensions.items():
        clean_dimension = clean.column_dimensions[key]
        clean_dimension.width = dimension.width
        clean_dimension.bestFit = dimension.bestFit
        clean_dimension.hidden = False
    for index, dimension in source_sheet.row_dimensions.items():
        clean_dimension = clean.row_dimensions[index]
        clean_dimension.height = dimension.height
        clean_dimension.hidden = False

    clean.sheet_view.showGridLines = source_sheet.sheet_view.showGridLines
    clean.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]
    clean.freeze_panes = None
    clean.auto_filter.ref = None
    clean.sheet_properties.filterMode = False

    workbook.remove(source_sheet)
    clean.title = DETAIL_SHEET
    return clean


def _averageifs_formula(
    metric_column: str,
    period: str,
    group: str,
    detail_last_row: int,
    priority: str | None = None,
) -> str:
    safe_group = group.replace('"', '""')
    criteria = (
        f"'Result Detail JOB'!$C$3:$C${detail_last_row},\"{period}\","
        f"'Result Detail JOB'!$D$3:$D${detail_last_row},\"{safe_group}\""
    )
    if priority:
        criteria += (
            f",'Result Detail JOB'!$E$3:$E${detail_last_row},\"{priority}\""
        )
    return (
        f"=IFERROR(AVERAGEIFS('Result Detail JOB'!${metric_column}$3:"
        f"${metric_column}${detail_last_row},{criteria}),0)"
    )


def _update_period_sheet(
    report: JobReport,
    sheet,
    period: str,
    dynamic: bool = False,
) -> None:
    pivot = report.pivot(period)
    lookup = pivot.set_index(["group_job", "priority"])
    columns = {
        "success_pct": {"H": 5, "M": 6, "L": 7},
        "retry_lt_3h_pct": {"H": 11, "M": 12, "L": 13},
        "retry_gt_3h_pct": {"H": 17, "M": 18, "L": 19},
    }
    for row in range(4, sheet.max_row + 1):
        group = canonical_group(sheet.cell(row, 1).value)
        if not group:
            continue
        sheet.cell(row, 1).value = group
        for metric, priority_columns in columns.items():
            for priority, column in priority_columns.items():
                key = (group, priority)
                if dynamic:
                    metric_column = {
                        "success_pct": "CY",
                        "retry_lt_3h_pct": "CZ",
                        "retry_gt_3h_pct": "DA",
                    }[metric]
                    sheet.cell(row, column).value = _averageifs_formula(
                        metric_column,
                        period,
                        group,
                        len(report.summary) + 2,
                        priority,
                    )
                else:
                    sheet.cell(row, column).value = (
                        float(lookup.loc[key, metric]) if key in lookup.index else None
                    )
        sheet.cell(row, 22).value = f"=K{row}+Q{row}+E{row}"
        sheet.cell(row, 23).value = f"=L{row}+R{row}+F{row}"
        sheet.cell(row, 24).value = f"=M{row}+S{row}+G{row}"


def _write_summary_blocks(
    report: JobReport,
    sheet,
    *,
    start_row: int | None = None,
    section_heading: bool = True,
    alphabetical_jobs: bool = False,
    include_priority_table: bool = True,
    dynamic: bool = False,
) -> None:
    start_row = start_row if start_row is not None else sheet.max_row + 4
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_fill = PatternFill("solid", fgColor="FFF200")
    header_fill = PatternFill("solid", fgColor="FFFFFF")
    metric_fill = PatternFill("solid", fgColor="FFF2CC")
    pivot_fill = PatternFill("solid", fgColor="D9E2F3")
    pivot_period_fill = PatternFill("solid", fgColor="FFF200")
    pivot_total_fill = PatternFill("solid", fgColor="D9E2F3")
    bold = Font(bold=True)
    percent_format = "0.00"

    if section_heading:
        sheet.cell(start_row - 2, 3, "JOB SUMMARY PER GROUP & PIVOT ALL")
        sheet.cell(start_row - 2, 3).font = Font(bold=True, size=14, color="17365D")

    # Keep Daily and Monthly blocks for the same group next to each other,
    # matching the layout used in the source workbook.
    for group in sorted(report.summary["group_job"].unique()):
        group_data = report.summary[report.summary["group_job"].eq(group)]
        for period in ("DAILY", "MONTHLY"):
            data = group_data[group_data["period"].eq(period)]
            data = data.sort_values(
                ["job_name"] if alphabetical_jobs else ["priority", "job_name"],
                key=lambda values: values.astype(str).str.upper(),
            )
            if data.empty:
                continue
            title_row = start_row
            header_row = start_row + 1
            table_titles = (
                (3, 9, f"JOB SUMMARY - {group} - {period}"),
                (17, 20, f"{group} - {period}"),
                (23, 26, f"AVERAGE PER PRIORITY - {group} - {period}"),
            )
            for first_column, last_column, title in table_titles:
                title_cell = sheet.cell(title_row, first_column, title)
                title_cell.fill = title_fill
                title_cell.font = bold
                title_cell.border = border
                title_cell.alignment = Alignment(horizontal="centerContinuous")
                for column in range(first_column + 1, last_column + 1):
                    cell = sheet.cell(title_row, column)
                    cell.value = None
                    cell.fill = title_fill
                    cell.border = border
                    cell.alignment = Alignment(horizontal="centerContinuous")
            headers = ["NAMA JOB", "PERIOD", "GROUP BARU", "KRITERIA", *PERCENT_HEADERS]
            for column, value in enumerate(headers, 3):
                cell = sheet.cell(header_row, column, value)
                cell.fill = header_fill
                cell.font = bold
                cell.border = border
                cell.alignment = Alignment(
                    wrap_text=True, vertical="center", horizontal="center"
                )

            for offset, (_, item) in enumerate(data.iterrows(), 2):
                row = start_row + offset
                values = [
                    item["job_name"], item["period"], item["group_job"], item["priority"],
                    item["success_pct"], item["retry_lt_3h_pct"], item["retry_gt_3h_pct"],
                ]
                for column, value in enumerate(values, 3):
                    cell = sheet.cell(row, column, value)
                    cell.border = border
                    if column >= 7:
                        cell.fill = metric_fill
                        cell.number_format = percent_format
                if dynamic:
                    for column, metric_column in zip(range(7, 10), ("CY", "CZ", "DA")):
                        sheet.cell(row, column).value = (
                            f"=IFERROR(INDEX('Result Detail JOB'!${metric_column}$3:"
                            f"${metric_column}${len(report.summary) + 2},MATCH(C{row},"
                            f"'Result Detail JOB'!$A$3:$A${len(report.summary) + 2},0)),0)"
                        )

            pivot = report.pivot(period)
            pivot = pivot[pivot["group_job"].eq(group)].copy()
            priority_order = {"H": 0, "M": 1, "L": 2}
            pivot["_priority_order"] = pivot["priority"].map(priority_order).fillna(99)
            pivot = pivot.sort_values("_priority_order").drop(columns="_priority_order")
            pivot_headers = [
                None,
                "Average of Success Persentage",
                "Average of Retry Success Rate < 3 Jam",
                "Average of Retry Success Rate > 3 Jam",
            ]
            for column, value in enumerate(pivot_headers, 17):
                cell = sheet.cell(header_row, column, value)
                cell.fill = pivot_fill
                cell.font = bold
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="center")

            aggregate = {
                "success_pct": data["success_pct"].mean(),
                "retry_lt_3h_pct": data["retry_lt_3h_pct"].mean(),
                "retry_gt_3h_pct": data["retry_gt_3h_pct"].mean(),
            }
            pivot_rows = [(period, aggregate, 0), (group, aggregate, 1)]
            pivot_rows.extend((item["priority"], item, 2) for _, item in pivot.iterrows())
            if len(pivot):
                pivot_rows.append(("Grand Total", aggregate, 0))
            for offset, (label, item, indent) in enumerate(pivot_rows, 2):
                row = start_row + offset
                sheet.cell(row, 17, label)
                sheet.cell(row, 17).alignment = Alignment(indent=indent)
                if item is not None:
                    for column, metric in zip(
                        range(18, 21),
                        ("success_pct", "retry_lt_3h_pct", "retry_gt_3h_pct"),
                    ):
                        if dynamic:
                            metric_column = {
                                "success_pct": "CY",
                                "retry_lt_3h_pct": "CZ",
                                "retry_gt_3h_pct": "DA",
                            }[metric]
                            priority = label if label in {"H", "M", "L"} else None
                            sheet.cell(row, column).value = _averageifs_formula(
                                metric_column,
                                period,
                                group,
                                len(report.summary) + 2,
                                priority,
                            )
                        else:
                            sheet.cell(row, column, float(item[metric]))
                        sheet.cell(row, column).number_format = percent_format
                for column in range(17, 21):
                    sheet.cell(row, column).border = border
                if label == period:
                    for column in range(17, 21):
                        sheet.cell(row, column).fill = pivot_period_fill
                        sheet.cell(row, column).font = bold
                if label == "Grand Total":
                    for column in range(17, 21):
                        sheet.cell(row, column).font = bold
                        sheet.cell(row, column).fill = pivot_total_fill

            # A second, flat pivot makes the H/L/M comparison immediately
            # visible without expanding the hierarchy-style Pivot All table.
            priority_headers = [
                "PRIORITY",
                "Average Success Persentage",
                "Average Retry Success < 3 Jam",
                "Average Retry Success > 3 Jam",
            ]
            for column, value in enumerate(priority_headers, 23):
                cell = sheet.cell(header_row, column, value)
                cell.fill = pivot_fill
                cell.font = bold
                cell.border = border
                cell.alignment = Alignment(
                    wrap_text=True, vertical="center", horizontal="center"
                )

            priority_lookup = pivot.set_index("priority")
            for offset, priority in enumerate(("H", "M", "L"), 2):
                row = start_row + offset
                label_cell = sheet.cell(row, 23, priority)
                label_cell.font = bold
                label_cell.border = border
                label_cell.alignment = Alignment(horizontal="center")
                for column, metric in zip(
                    range(24, 27),
                    ("success_pct", "retry_lt_3h_pct", "retry_gt_3h_pct"),
                ):
                    cell = sheet.cell(row, column)
                    if priority in priority_lookup.index:
                        cell.value = float(priority_lookup.loc[priority, metric])
                        cell.number_format = percent_format
                    cell.border = border
                    cell.fill = metric_fill

            total_row = start_row + 5
            sheet.cell(total_row, 23, "Grand Total")
            for column, metric in zip(
                range(24, 27),
                ("success_pct", "retry_lt_3h_pct", "retry_gt_3h_pct"),
            ):
                sheet.cell(total_row, column, float(aggregate[metric]))
                sheet.cell(total_row, column).number_format = percent_format
            for column in range(23, 27):
                sheet.cell(total_row, column).border = border
                sheet.cell(total_row, column).fill = pivot_total_fill
                sheet.cell(total_row, column).font = bold

            block_height = max(len(data) + 2, len(pivot_rows) + 2)
            start_row += block_height + 3

    if not include_priority_table:
        # The flat Average per Priority table occupies W:Z. Pivot All already
        # contains H-M-L, so remove the duplicate table from Sum Job.
        sheet.delete_cols(23, 4)

    widths = {
        3: 62, 4: 12, 5: 34, 6: 12, 7: 22, 8: 28, 9: 28,
        17: 34, 18: 24, 19: 30, 20: 30,
    }
    if include_priority_table:
        widths.update({23: 13, 24: 24, 25: 28, 26: 28})
    for column, width in widths.items():
        sheet.column_dimensions[get_column_letter(column)].width = width


def build_excel_report(
    report: JobReport,
    source: str | Path | BinaryIO | BytesIO,
    dynamic: bool = False,
) -> bytes:
    """Create a three-sheet report while retaining the source workbook layout."""
    workbook = load_workbook(_template_bytes(source), data_only=False)
    missing = [name for name in OUTPUT_SHEETS if name not in workbook.sheetnames]
    if missing:
        raise ValueError(f"Sheet template tidak ditemukan: {', '.join(missing)}")

    for sheet in list(workbook.worksheets):
        if sheet.title not in OUTPUT_SHEETS:
            workbook.remove(sheet)
    # Reorder exactly as requested, regardless of the source order.
    workbook._sheets = [workbook[name] for name in OUTPUT_SHEETS]

    detail = workbook[DETAIL_SHEET]
    detail_last_row = detail.max_row
    _update_detail_metrics(report, detail, dynamic=dynamic)
    _remove_merged_cells_for_copy(detail)
    _reset_detail_selection_state(detail, detail_last_row)
    detail = _replace_with_clean_detail_sheet(workbook, detail)
    _update_period_sheet(report, workbook["Daily"], "DAILY", dynamic=dynamic)
    _update_period_sheet(report, workbook["Monthly"], "MONTHLY", dynamic=dynamic)

    # Rebuild Sum Job so it contains current Python results rather than stale
    # PivotTable cache values from the uploaded workbook.
    old_sum_job = workbook["Sum Job"]
    sum_index = workbook.worksheets.index(old_sum_job)
    workbook.remove(old_sum_job)
    sum_job = workbook.create_sheet("Sum Job", sum_index)
    _write_summary_blocks(
        report,
        sum_job,
        start_row=3,
        section_heading=False,
        alphabetical_jobs=True,
        include_priority_table=False,
        dynamic=dynamic,
    )

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    if dynamic:
        workbook.calculation.calcMode = "auto"
        workbook.calculation.calcOnSave = True
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()
