"""Generate conflict-free school timetables for Classes 1-12, divisions A-D.

Input:  requirements.csv
Output: school_timetable.xlsx

CSV columns:
    class,division,subject,teacher,periods_per_week

Each CSV row describes one subject/teacher allocation for one division. If a
subject is split across teachers, give each allocation its own row.
"""

from __future__ import annotations

import csv
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from ortools.sat.python import cp_model
except ImportError:
    cp_model = None


# --------------------------- SCHOOL SETTINGS ---------------------------
GRADES = range(1, 13)
DIVISIONS = ("A", "B", "C", "D")
DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
PERIODS_PER_DAY = 7
INPUT_CSV = Path("requirements.csv")
OUTPUT_XLSX = Path("school_timetable.xlsx")
SOLVER_TIME_LIMIT_SECONDS = 180

CSV_COLUMNS = ("class", "division", "subject", "teacher", "periods_per_week")

# Workbook styling
NAVY = "183153"
BLUE = "DCE8F5"
PALE = "F4F7FB"
WHITE = "FFFFFF"
GREY = "E9EDF2"
BORDER_COLOR = "D3DAE3"


def section_name(grade: int, division: str) -> str:
    return f"Class {grade}-{division}"


def all_sections() -> list[str]:
    return [section_name(grade, division) for grade in GRADES for division in DIVISIONS]


def create_input_template(path: Path) -> None:
    """Create a CSV with one starter row for every division."""
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_COLUMNS)
        for grade in GRADES:
            for division in DIVISIONS:
                writer.writerow([grade, division, "", "", ""])


def read_requirements(path: Path) -> list[dict[str, Any]]:
    """Read and validate lesson requirements from CSV."""
    requirements: list[dict[str, Any]] = []

    with path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("The CSV is empty. Keep the required header row.")

        # Permit harmless spaces/case differences in column names.
        header_map = {str(name).strip().lower(): name for name in reader.fieldnames}
        missing_columns = [col for col in CSV_COLUMNS if col not in header_map]
        if missing_columns:
            raise ValueError(
                "CSV is missing these columns: " + ", ".join(missing_columns)
                + f". Required columns: {', '.join(CSV_COLUMNS)}"
            )

        for row_number, row in enumerate(reader, start=2):
            values = {
                key: str(row.get(header_map[key], "") or "").strip()
                for key in CSV_COLUMNS
            }

            # The pre-created template has placeholder rows with no lesson yet.
            if not values["subject"] and not values["teacher"] and not values["periods_per_week"]:
                continue

            if not values["class"] or not values["division"]:
                raise ValueError(f"Row {row_number}: class and division are required.")
            if not values["subject"] or not values["teacher"] or not values["periods_per_week"]:
                raise ValueError(
                    f"Row {row_number}: subject, teacher, and periods_per_week must all be filled."
                )

            try:
                grade = int(values["class"])
            except ValueError as exc:
                raise ValueError(f"Row {row_number}: class must be a number from 1 to 12.") from exc

            division = values["division"].upper()
            if grade not in GRADES:
                raise ValueError(f"Row {row_number}: class must be from 1 to 12.")
            if division not in DIVISIONS:
                raise ValueError(f"Row {row_number}: division must be A, B, C, or D.")

            try:
                periods = int(values["periods_per_week"])
            except ValueError as exc:
                raise ValueError(f"Row {row_number}: periods_per_week must be a positive integer.") from exc
            if periods <= 0:
                raise ValueError(f"Row {row_number}: periods_per_week must be greater than zero.")

            requirements.append({
                "grade": grade,
                "division": division,
                "section": section_name(grade, division),
                "subject": values["subject"],
                "teacher": values["teacher"],
                "periods": periods,
                "source_row": row_number,
            })

    if not requirements:
        raise ValueError(
            "No lesson requirements were entered. Fill the subject, teacher, and "
            "periods_per_week columns in requirements.csv. Add one row per subject/teacher assignment."
        )

    sections_with_lessons = {item["section"] for item in requirements}
    missing_sections = sorted(set(all_sections()) - sections_with_lessons)
    if missing_sections:
        preview = ", ".join(missing_sections[:12])
        suffix = " ..." if len(missing_sections) > 12 else ""
        raise ValueError(
            f"No lesson requirements were provided for {len(missing_sections)} division(s): "
            f"{preview}{suffix}. Each of the 48 divisions must have at least one subject row."
        )

    return requirements


def expand_lessons(requirements: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Turn weekly counts into individual lesson instances for the solver."""
    lessons: list[dict[str, Any]] = []
    for item in requirements:
        for occurrence in range(1, item["periods"] + 1):
            lessons.append({
                "section": item["section"],
                "grade": item["grade"],
                "division": item["division"],
                "subject": item["subject"],
                "teacher": item["teacher"],
                "occurrence": occurrence,
            })
    return lessons


def validate_capacity(lessons: list[dict[str, Any]], slots_per_week: int) -> None:
    """Catch simple overloads before invoking the solver."""
    section_load: dict[str, int] = defaultdict(int)
    teacher_load: dict[str, int] = defaultdict(int)
    for lesson in lessons:
        section_load[lesson["section"]] += 1
        teacher_load[lesson["teacher"]] += 1

    too_full_sections = [f"{name}: {load}" for name, load in section_load.items() if load > slots_per_week]
    if too_full_sections:
        raise ValueError(
            f"A division cannot have more than {slots_per_week} lessons per week. "
            + "; ".join(too_full_sections)
        )

    too_full_teachers = [f"{name}: {load}" for name, load in teacher_load.items() if load > slots_per_week]
    if too_full_teachers:
        raise ValueError(
            f"A teacher cannot teach more than {slots_per_week} periods per week without clashes. "
            + "; ".join(too_full_teachers)
            + ". Assign another teacher or increase the number of available periods."
        )


def solve_timetable(lessons: list[dict[str, Any]]) -> list[int]:
    """Return the selected slot index for every lesson, with no resource clashes."""
    if cp_model is None:
        raise RuntimeError(
            "OR-Tools is not installed. Run: pip install -r requirements.txt"
        )

    slots_per_day = PERIODS_PER_DAY
    slot_count = len(DAYS) * slots_per_day
    model = cp_model.CpModel()

    # Each lesson must be assigned to exactly one day/period.
    assignment = [
        [model.NewBoolVar(f"lesson_{i}_slot_{slot}") for slot in range(slot_count)]
        for i in range(len(lessons))
    ]
    for i in range(len(lessons)):
        model.AddExactlyOne(assignment[i])

    lessons_by_section: dict[str, list[int]] = defaultdict(list)
    lessons_by_teacher: dict[str, list[int]] = defaultdict(list)
    for lesson_id, lesson in enumerate(lessons):
        lessons_by_section[lesson["section"]].append(lesson_id)
        lessons_by_teacher[lesson["teacher"]].append(lesson_id)

    # A division can have at most one lesson in a period.
    for section, lesson_ids in lessons_by_section.items():
        for slot in range(slot_count):
            model.AddAtMostOne(assignment[i][slot] for i in lesson_ids)

    # A teacher can teach at most one division in a period.
    for teacher, lesson_ids in lessons_by_teacher.items():
        for slot in range(slot_count):
            model.AddAtMostOne(assignment[i][slot] for i in lesson_ids)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT_SECONDS
    solver.parameters.num_search_workers = 8
    solver.parameters.random_seed = 17

    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        if status == cp_model.INFEASIBLE:
            detail = (
                "No clash-free timetable satisfies these requirements. Check teacher workloads, "
                "weekly periods, and whether one teacher has been assigned to too many divisions."
            )
        elif status == cp_model.UNKNOWN:
            detail = "The solver did not find a timetable within the time limit. Try increasing the time limit."
        else:
            detail = f"Solver status: {solver.StatusName(status)}."
        raise RuntimeError(detail)

    selected_slots: list[int] = []
    for lesson_id in range(len(lessons)):
        chosen = [slot for slot in range(slot_count) if solver.Value(assignment[lesson_id][slot])]
        if len(chosen) != 1:
            raise RuntimeError(f"Internal error: lesson {lesson_id} has {len(chosen)} assigned slots.")
        selected_slots.append(chosen[0])

    return selected_slots


def safe_sheet_name(name: str, used_names: set[str]) -> str:
    """Return a legal, unique Excel worksheet name."""
    cleaned = re.sub(r"[\[\]:*?/\\]", "-", name).strip().strip("'") or "Sheet"
    cleaned = cleaned[:31]
    candidate = cleaned
    suffix_num = 2
    while candidate.lower() in used_names:
        suffix = f" ({suffix_num})"
        candidate = cleaned[:31 - len(suffix)] + suffix
        suffix_num += 1
    used_names.add(candidate.lower())
    return candidate


def style_timetable_sheet(ws, title: str, subtitle: str, grid: list[list[str | None]]) -> None:
    """Write and format a Monday-to-Friday timetable grid."""
    period_count = PERIODS_PER_DAY
    final_col = period_count + 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=final_col)
    ws.cell(1, 1, title)
    ws.cell(1, 1).font = Font(name="Aptos Display", size=16, bold=True, color=WHITE)
    ws.cell(1, 1).fill = PatternFill("solid", fgColor=NAVY)
    ws.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=final_col)
    ws.cell(2, 1, subtitle)
    ws.cell(2, 1).font = Font(name="Aptos", size=10, italic=True, color="526174")
    ws.cell(2, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    thin = Side(style="thin", color=BORDER_COLOR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header = ws.cell(3, 1, "Day")
    header.fill = PatternFill("solid", fgColor=BLUE)
    header.font = Font(bold=True, color=NAVY)
    header.alignment = Alignment(horizontal="center", vertical="center")
    header.border = border

    for period in range(period_count):
        cell = ws.cell(3, period + 2, f"Period {period + 1}")
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(bold=True, color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[3].height = 26

    for day_index, day in enumerate(DAYS):
        row_num = day_index + 4
        day_cell = ws.cell(row_num, 1, day)
        day_cell.font = Font(bold=True, color=NAVY)
        day_cell.fill = PatternFill("solid", fgColor=PALE)
        day_cell.alignment = Alignment(horizontal="left", vertical="center")
        day_cell.border = border
        for period_index in range(period_count):
            text = grid[day_index][period_index] or ""
            cell = ws.cell(row_num, period_index + 2, text)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            if not text:
                cell.fill = PatternFill("solid", fgColor=GREY)
        ws.row_dimensions[row_num].height = 57

    ws.column_dimensions["A"].width = 15
    for col in range(2, final_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 19
    ws.freeze_panes = "B4"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{get_column_letter(final_col)}{len(DAYS) + 3}"


def create_workbook(
    requirements: list[dict[str, Any]],
    lessons: list[dict[str, Any]],
    selected_slots: list[int],
    output_path: Path,
) -> None:
    """Write one worksheet per division and one worksheet per teacher."""
    section_grid: dict[str, list[list[str | None]]] = {
        section: [[None for _ in range(PERIODS_PER_DAY)] for _ in DAYS]
        for section in all_sections()
    }
    teacher_names = sorted({lesson["teacher"] for lesson in lessons}, key=str.casefold)
    teacher_grid: dict[str, list[list[str | None]]] = {
        teacher: [[None for _ in range(PERIODS_PER_DAY)] for _ in DAYS]
        for teacher in teacher_names
    }

    for lesson, slot in zip(lessons, selected_slots):
        day_index, period_index = divmod(slot, PERIODS_PER_DAY)
        section = lesson["section"]
        teacher = lesson["teacher"]
        class_value = f"{lesson['subject']}\n{teacher}"
        teacher_value = f"{section}\n{lesson['subject']}"

        # Defensive verification, independent of the solver constraints.
        if section_grid[section][day_index][period_index] is not None:
            raise RuntimeError(f"Detected a class clash at {section}, {DAYS[day_index]}, Period {period_index + 1}.")
        if teacher_grid[teacher][day_index][period_index] is not None:
            raise RuntimeError(
                f"Detected a teacher clash for {teacher}, {DAYS[day_index]}, Period {period_index + 1}."
            )

        section_grid[section][day_index][period_index] = class_value
        teacher_grid[teacher][day_index][period_index] = teacher_value

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.sheet_view.showGridLines = False

    summary.merge_cells("A1:D1")
    summary["A1"] = "School Timetable Generator"
    summary["A1"].font = Font(name="Aptos Display", size=18, bold=True, color=WHITE)
    summary["A1"].fill = PatternFill("solid", fgColor=NAVY)
    summary["A1"].alignment = Alignment(horizontal="left", vertical="center")
    summary.row_dimensions[1].height = 34

    summary["A3"] = "Setting"
    summary["B3"] = "Value"
    for cell in (summary["A3"], summary["B3"]):
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(bold=True, color=NAVY)

    summary_rows = [
        ("Grades", "1-12"),
        ("Divisions per grade", "A, B, C, D"),
        ("Division timetables", len(all_sections())),
        ("School days", len(DAYS)),
        ("Periods per day", PERIODS_PER_DAY),
        ("Total period slots per division per week", len(DAYS) * PERIODS_PER_DAY),
        ("Scheduled lessons per week", len(lessons)),
        ("Unique teachers", len(teacher_names)),
        ("Clash checks", "No class or teacher is scheduled in two places in the same period."),
    ]
    for row_num, (label, value) in enumerate(summary_rows, start=4):
        summary.cell(row_num, 1, label)
        summary.cell(row_num, 2, value)
        summary.cell(row_num, 1).font = Font(bold=True)
        summary.cell(row_num, 1).fill = PatternFill("solid", fgColor=PALE)

    summary["A15"] = "Input format"
    summary["A15"].font = Font(bold=True, color=NAVY, size=12)
    summary["A16"] = "One row per subject/teacher assignment. periods_per_week is the number of lessons for that row."
    summary.merge_cells("A16:D16")
    summary["A17"] = "Change DAYS and PERIODS_PER_DAY near the top of the Python script to match your school."
    summary.merge_cells("A17:D17")
    summary["A18"] = "Grey cells are free periods. A class cell shows Subject / Teacher; a teacher cell shows Class / Subject."
    summary.merge_cells("A18:D18")
    for row in (16, 17, 18):
        summary.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
        summary.row_dimensions[row].height = 30
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 70
    summary.column_dimensions["C"].width = 18
    summary.column_dimensions["D"].width = 18

    # Include the input requirements for auditability.
    data_ws = wb.create_sheet("Requirements")
    data_ws.append(list(CSV_COLUMNS))
    for item in requirements:
        data_ws.append([
            item["grade"], item["division"], item["subject"], item["teacher"], item["periods"]
        ])
    data_ws.freeze_panes = "A2"
    data_ws.auto_filter.ref = data_ws.dimensions
    data_ws.sheet_view.showGridLines = False
    for cell in data_ws[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center")
    widths = [12, 12, 24, 26, 20]
    for col, width in enumerate(widths, start=1):
        data_ws.column_dimensions[get_column_letter(col)].width = width

    used_names = {"summary", "requirements"}
    # One class timetable per division (48 total).
    for section in all_sections():
        ws = wb.create_sheet(safe_sheet_name(section, used_names))
        style_timetable_sheet(ws, section, "Class timetable | subject followed by assigned teacher", section_grid[section])

    # One timetable per teacher.
    for teacher in teacher_names:
        ws = wb.create_sheet(safe_sheet_name(f"Teacher - {teacher}", used_names))
        style_timetable_sheet(ws, teacher, "Teacher timetable | division followed by subject", teacher_grid[teacher])

    wb.save(output_path)


def main() -> int:
    if not INPUT_CSV.exists():
        create_input_template(INPUT_CSV)
        print(f"Created {INPUT_CSV.resolve()} with starter rows for all 48 divisions.")
        print("Fill in the subject, teacher, and periods_per_week columns, then run this script again.")
        return 0

    try:
        requirements = read_requirements(INPUT_CSV)
        lessons = expand_lessons(requirements)
        validate_capacity(lessons, len(DAYS) * PERIODS_PER_DAY)
        selected_slots = solve_timetable(lessons)
        create_workbook(requirements, lessons, selected_slots, OUTPUT_XLSX)
    except (ValueError, RuntimeError, OSError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Success: created {OUTPUT_XLSX.resolve()}")
    print(f"Division timetables: {len(all_sections())}")
    print(f"Teacher timetables: {len({lesson['teacher'] for lesson in lessons})}")
    print(f"Lessons scheduled: {len(lessons)} per week")
    print("Verified: no class or teacher is double-booked in any period.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
