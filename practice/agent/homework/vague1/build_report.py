#!/usr/bin/env python3
"""Строит Excel-отчёт по задачам из todo_db.json.

Данные берутся только из БД todo (apps/todo_db.json).
Результат: homework/vague1/report.xlsx
"""
import json
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DB_PATH = Path(__file__).resolve().parents[4] / "apps" / "todo_db.json"
OUTPUT_PATH = Path(__file__).resolve().parent / "report.xlsx"
REPORT_DATE = date(2026, 11, 10)
PRIORITIES = ("low", "normal", "high")


def load_tasks(path):
    """Загрузить задачи из JSON-БД todo."""
    raw = json.loads(path.read_text(encoding="utf-8"))
    tasks = raw["payload"]["tasks"]
    tasks = sorted(tasks, key=lambda t: t["id"])
    return tasks


def build_summary(tasks):
    total = len(tasks)
    done = sum(1 for t in tasks if t["done"])
    not_done = total - done
    percent = (done / total) * 100 if total else 0.0
    overdue = sum(
        1 for t in tasks
        if not t["done"] and t["due"] is not None
        and date.fromisoformat(t["due"]) < REPORT_DATE
    )
    priority_counts = {p: sum(1 for t in tasks if t["priority"] == p) for p in PRIORITIES}
    return {
        "total": total,
        "done": done,
        "not_done": not_done,
        "percent": percent,
        "overdue": overdue,
        "priority_counts": priority_counts,
    }


def fmt_date(v):
    return date.fromisoformat(v) if v else None


def write_tasks_sheet(ws, tasks):
    headers = ["id", "название", "priority", "due", "done"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")
    for row, t in enumerate(tasks, start=2):
        due = fmt_date(t.get("due"))
        ws.cell(row=row, column=1, value=t["id"])
        ws.cell(row=row, column=2, value=t["title"])
        ws.cell(row=row, column=3, value=t["priority"])
        if due is not None:
            cell = ws.cell(row=row, column=4, value=due)
            cell.number_format = "YYYY-MM-DD"
        else:
            ws.cell(row=row, column=4, value="")
        ws.cell(row=row, column=5, value=t["done"])
    widths = {"A": 8, "B": 48, "C": 10, "D": 12, "E": 8}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def write_summary_sheet(ws, summary):
    pc = summary["priority_counts"]
    rows = [
        ("Всего", summary["total"]),
        ("Выполнено", summary["done"]),
        ("Не выполнено", summary["not_done"]),
        ("Процент выполненных, %", summary["percent"]),
        ("Просрочено", summary["overdue"]),
        ("Low", pc["low"]),
        ("Normal", pc["normal"]),
        ("High", pc["high"]),
    ]
    ws.cell(row=1, column=1, value="Показатель").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=1, column=2, value="Значение").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor="4472C4")
    ws.cell(row=1, column=2).fill = PatternFill("solid", fgColor="4472C4")
    for row, (label, value) in enumerate(rows, start=2):
        ws.cell(row=row, column=1, value=label)
        cell = ws.cell(row=row, column=2, value=value)
        if label == "Процент выполненных, %":
            cell.number_format = '0.0"%"'
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14


def main():
    tasks = load_tasks(DB_PATH)
    summary = build_summary(tasks)

    wb = Workbook()
    ws = wb.active
    ws.title = "Задачи"
    write_tasks_sheet(ws, tasks)

    ws2 = wb.create_sheet("Сводка")
    write_summary_sheet(ws2, summary)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"saved: {OUTPUT_PATH}")
    print(f"tasks={summary['total']} done={summary['done']} not_done={summary['not_done']} "
          f"percent={summary['percent']:.2f}% overdue={summary['overdue']} "
          f"low={summary['priority_counts']['low']} "
          f"normal={summary['priority_counts']['normal']} "
          f"high={summary['priority_counts']['high']}")


if __name__ == "__main__":
    main()