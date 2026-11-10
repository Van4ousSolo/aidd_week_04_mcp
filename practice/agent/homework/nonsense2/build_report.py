#!/usr/bin/env python3
"""Генерирует Excel-отчёт по задачам TODO-API в homework/nonsense2/report.xlsx."""
import json
import os
import sys
import urllib.request
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

API_URL = os.environ.get("TODO_API_URL", "http://127.0.0.1:5001").rstrip("/")
OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "report.xlsx")
EXIT_OK, EXIT_ERR = 0, 1

HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(color="FFFFFF", bold=True)
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="B0B0B0")


def fetch_tasks():
    req = urllib.request.Request(API_URL + "/api/tasks")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.load(resp)


def count_by(iterable, key):
    out = {}
    for t in iterable:
        out[t[key]] = out.get(t[key], 0) + 1
    return dict(sorted(out.items()))


def build_summary(ws, tasks, today):
    """Заполняет лист сводки статистикой по задачам."""
    done = [t for t in tasks if t["done"]]
    pending = [t for t in tasks if not t["done"]]
    overdue = [t for t in pending if t["due"] and t["due"] < today]

    done_by_p = count_by(done, "priority")
    pending_by_p = count_by(pending, "priority")
    all_by_p = count_by(tasks, "priority")

    def row(label, a="", b="", c="", d=""):
        ws.append([label, a, b, c, d])

    ws.append(["Отчёт по задачам TODO", "", "", "", ""])
    ws.append(["Сформирован:", date.today().strftime("%Y-%m-%d %H:%M:%S"), "", "", ""])
    ws.append([])
    row("Общее количество задач", total := len(tasks), "", "", "")
    row("Выполнено", len(done), "", "", "")
    row("Невыполнено", len(pending), "", "", "")
    row("Просрочено (невыполненные со сроком раньше сегодня)", len(overdue), "", "", "")
    ws.append([])
    row("Сводка по приоритетам", "", "", "", "")
    row("Приоритет", "Всего", "Выполнено", "Невыполнено", "")
    for p in ("high", "normal", "low"):
        row(f"  {p}", all_by_p.get(p, 0), done_by_p.get(p, 0), pending_by_p.get(p, 0), "")
    ws.append([])
    row("Сводка по состоянию", "", "", "", "")
    row("Состояние", "Всего", "high", "normal", "low")
    row("Выполнено", len(done),
        done_by_p.get("high", 0), done_by_p.get("normal", 0), done_by_p.get("low", 0))
    row("Невыполнено", len(pending),
        pending_by_p.get("high", 0), pending_by_p.get("normal", 0), pending_by_p.get("low", 0))

    # ---- форматирование ----
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left")
    ws.cell(row=2, column=1).font = Font(italic=True)

    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if label in ("Приоритет", "Состояние"):
            for c in ws[r]:
                c.font = Font(bold=True)
                c.border = Border(bottom=THIN)
        elif label in ("Выполнено", "Невыполнено", "Просрочено (невыполненные со сроком раньше сегодня)"):
            fill = OK_FILL if label == "Выполнено" else (BAD_FILL if label.startswith("Просрочено") else WARN_FILL)
            for c in ws[r]:
                c.fill = fill

    ws.column_dimensions["A"].width = 46
    for col in "BCDE":
        ws.column_dimensions[col].width = 14
    ws.freeze_panes = "B2"


def build_tasks_sheet(ws, tasks):
    """Заполняет лист задач всеми записями."""
    headers = ["id", "title", "priority", "due", "done", "created_at"]
    ws.append(headers)
    for t in tasks:
        ws.append([t["id"], t["title"], t["priority"],
                   t["due"], "да" if t["done"] else "нет",
                   t["created_at"]])

    for c in ws[1]:
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        c.border = Border(bottom=THIN)
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 20


def main():
    try:
        tasks = fetch_tasks()
    except Exception as e:
        print(f"Не удалось получить задачи из {API_URL}/api/tasks: {e}", file=sys.stderr)
        return EXIT_ERR

    today = date.today().isoformat()
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Итого"
    ws_tasks = wb.create_sheet("Задачи")

    build_summary(ws_sum, tasks, today)
    build_tasks_sheet(ws_tasks, tasks)

    try:
        wb.save(OUT_PATH)
    except Exception as e:
        print(f"Не удалось сохранить {OUT_PATH}: {e}", file=sys.stderr)
        return EXIT_ERR

    pending = sum(1 for t in tasks if not t["done"])
    overdue = sum(1 for t in tasks if not t["done"] and t["due"] and t["due"] < today)
    print(f"Отчёт сохранён: {OUT_PATH}")
    print(f"Всего: {len(tasks)} | выполнено: {len(tasks)-pending} | "
          f"невыполнено: {pending} | просрочено: {overdue}")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())