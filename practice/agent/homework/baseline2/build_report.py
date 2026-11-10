import json
import os
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_PATH = os.path.join(SCRIPT_DIR, "tasks_raw.jsonl")
OUT_PATH = os.path.join(SCRIPT_DIR, "report.xlsx")
REPORT_DATE = date(2026, 11, 10)
PRIORITIES = ("low", "normal", "high")


def load_tasks(path):
    tasks = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                tasks.append(json.loads(line))
    tasks.sort(key=lambda t: t["id"])
    return tasks


def build_summary(tasks):
    total = len(tasks)
    done = sum(1 for t in tasks if t["done"])
    not_done = total - done
    percent = (done / total) * 100 if total else 0.0
    overdue = [
        t for t in tasks
        if not t["done"] and t["due"] is not None and date.fromisoformat(t["due"]) < REPORT_DATE
    ]
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
    return v.strftime("%Y-%m-%d") if isinstance(v, date) else v


def write_tasks_sheet(ws, tasks):
    headers = ["id", "название", "priority", "due", "done"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")
    for row, t in enumerate(tasks, start=2):
        due = date.fromisoformat(t["due"]) if t["due"] else None
        ws.cell(row=row, column=1, value=t["id"])
        ws.cell(row=row, column=2, value=t["title"])
        ws.cell(row=row, column=3, value=t["priority"])
        ws.cell(row=row, column=4, value=due)
        if isinstance(due, date):
            ws.cell(row=row, column=4).number_format = "YYYY-MM-DD"
        ws.cell(row=row, column=5, value=t["done"])
    widths = {"id": 8, "название": 45, "priority": 10, "due": 12, "done": 8}
    for col, w in enumerate(widths.values(), start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"


def write_summary_sheet(ws, summary):
    pc = summary["priority_counts"]
    rows = [
        ("Всего задач", summary["total"]),
        ("Выполнено", summary["done"]),
        ("Не выполнено", summary["not_done"]),
        ("Процент выполненных", summary["percent"]),
        ("Просрочено", len(summary["overdue"])),
        ("Low", pc["low"]),
        ("Normal", pc["normal"]),
        ("High", pc["high"]),
    ]
    ws.cell(row=1, column=1, value="Параметр").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=1, column=2, value="Значение").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor="4472C4")
    ws.cell(row=1, column=2).fill = PatternFill("solid", fgColor="4472C4")
    for row, (label, value) in enumerate(rows, start=2):
        ws.cell(row=row, column=1, value=label)
        cell = ws.cell(row=row, column=2, value=value)
        if label == "Процент выполненных":
            cell.number_format = '0.0"%"'
    ws.cell(row=len(rows) + 2, column=1,
            value=f"Дата отчёта: {REPORT_DATE.isoformat()}; просрочено = невыполненные с due раньше даты")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 14


def main():
    tasks = load_tasks(RAW_PATH)
    summary = build_summary(tasks)

    wb = Workbook()
    ws = wb.active
    ws.title = "Задачи"
    write_tasks_sheet(ws, tasks)

    ws2 = wb.create_sheet("Сводка")
    write_summary_sheet(ws2, summary)

    wb.save(OUT_PATH)
    pc = summary["priority_counts"]
    print(f"Сохранено: {OUT_PATH}")
    print(f"Задач всего: {summary['total']}; выполнено: {summary['done']}; "
          f"не выполнено: {summary['not_done']}; %: {summary['percent']:.1f}; "
          f"просрочено: {len(summary['overdue'])}")
    print(f"Приоритеты: low={pc['low']} normal={pc['normal']} high={pc['high']}")


if __name__ == "__main__":
    main()