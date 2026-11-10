#!/usr/bin/env python3
"""Строит Excel-отчёт по задачам MCP todo из вывода list_tasks.

Данные берутся только из сохранённого вывода list_tasks (не из БД и не по сети).
Результат: reports/run3/report.xlsx
"""
import json
import re
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Пути и константы
# ---------------------------------------------------------------------------
PERSISTED_INPUT = (
    Path.home() / ".qwen/tmp/385c1a20454bfaffec34ff1208718d56fa366bbd4e0d524dba0ab3410d8dccc1/tool-results/chatcmpl-tool-87732f686cebce39.txt"
)
REPORT_DIR = Path(__file__).resolve().parent
OUTPUT_XLSX = REPORT_DIR / "report.xlsx"
REPORT_DATE = date(2026, 11, 10)

# Счётчики, возвращённые сервером через count_only (для перекрёстной проверки).
EXPECTED = {"total": 364, "done": 118, "overdue_undone": 45,
            "high": 125, "low": 124, "normal": 115}

PRIO_LABEL = {"low": "Низкий", "normal": "Обычный", "high": "Высокий"}
STATUS_LABEL = {True: "Выполнено", False: "Невыполнено"}

HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14)
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
SECTION_FONT = Font(bold=True, size=11)

# ---------------------------------------------------------------------------
# Разбор сохранённого вывода list_tasks (конкатенированные JSON-объекты,
# разделённые пробельными символами).
# ---------------------------------------------------------------------------
def parse_records(text: str) -> list[dict]:
    dec = json.JSONDecoder()
    idx = 0
    n = len(text)
    records = []
    while idx < n:
        while idx < n and text[idx] in " \t\r\n":
            idx += 1
        if idx >= n:
            break
        obj, end = dec.raw_decode(text, idx)
        records.append(obj)
        idx = end
    return records


# ---------------------------------------------------------------------------
# Сводные вычисления
# ---------------------------------------------------------------------------
def build_summary(records):
    total = len(records)
    done = sum(1 for r in records if r["done"])
    undone = total - done

    by_prio = {}
    for r in records:
        by_prio.setdefault(r["priority"], []).append(r)

    # Просроченные невыполненные: невыполненные со сроком раньше сегодняшнего дня.
    overdue = sorted(
        r["id"] for r in records
        if not r["done"] and r["due"] is not None
        and date.fromisoformat(r["due"]) < REPORT_DATE
    )

    checks = {
        "total": total, "done": done, "overdue_undone": len(overdue),
        "high": len(by_prio.get("high", [])),
        "low": len(by_prio.get("low", [])),
        "normal": len(by_prio.get("normal", [])),
    }
    return by_prio, done, undone, overdue, checks


# ---------------------------------------------------------------------------
# Лист «Задачи»
# ---------------------------------------------------------------------------
def write_tasks_sheet(wb, records):
    ws = wb.create_sheet("Задачи")
    headers = ["ID", "Название", "Приоритет", "Срок", "Статус"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    r = 2
    for rec in sorted(records, key=lambda x: x["id"]):
        due = date.fromisoformat(rec["due"]) if rec["due"] else None
        ws.cell(row=r, column=1, value=rec["id"])
        ws.cell(row=r, column=2, value=rec["title"])
        ws.cell(row=r, column=3, value=PRIO_LABEL[rec["priority"]])
        if due is not None:
            cell = ws.cell(row=r, column=4, value=due)
            cell.number_format = "YYYY-MM-DD"
        ws.cell(row=r, column=5, value=STATUS_LABEL[rec["done"]])
        r += 1

    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).border = BORDER
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{r - 1}"

    widths = {"A": 7, "B": 52, "C": 12, "D": 12, "E": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["B"].width = 52
    return r - 1


# ---------------------------------------------------------------------------
# Лист «Сводка»
# ---------------------------------------------------------------------------
def write_summary_sheet(wb, by_prio, done, undone, overdue, checks):
    ws = wb.create_sheet("Сводка")
    order = ["low", "normal", "high"]

    ws["A1"] = f"Сводка по задачам (на {REPORT_DATE.isoformat()})"
    ws["A1"].font = TITLE_FONT

    # --- Таблица по приоритетам и статусам ---
    ws["A3"] = "Количество задач по приоритетам и статусам"
    ws["A3"].font = SECTION_FONT
    headers = ["Приоритет", "Выполнено", "Невыполнено", "Всего"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")
    for i, prio in enumerate(order):
        row = 5 + i
        rows = by_prio.get(prio, [])
        d = sum(1 for x in rows if x["done"])
        ws.cell(row=row, column=1, value=PRIO_LABEL[prio])
        ws.cell(row=row, column=2, value=d)
        ws.cell(row=row, column=3, value=len(rows) - d)
        ws.cell(row=row, column=4, value=len(rows))
    total_row = 5 + len(order)
    ws.cell(row=total_row, column=1, value="Итого")
    ws.cell(row=total_row, column=2, value=done)
    ws.cell(row=total_row, column=3, value=undone)
    ws.cell(row=total_row, column=4, value=done + undone)
    for c in range(1, 5):
        cell = ws.cell(row=total_row, column=c)
        cell.font = Font(bold=True)
        cell.border = BORDER
    for row in range(4, total_row + 1):
        for c in range(1, 5):
            ws.cell(row=row, column=c).border = BORDER
    for c, w in zip("ABCD", (18, 14, 14, 12)):
        ws.column_dimensions[c].width = w

    # --- Ключевые показатели ---
    ws["A10"] = "Ключевые показатели"
    ws["A10"].font = SECTION_FONT
    metrics = [
        ("Всего задач", checks["total"]),
        ("Выполнено", done),
        ("Невыполнено", undone),
        ("Процент выполненных", done / (done + undone) if done + undone else 0),
        ("Просроченных невыполненных", len(overdue)),
    ]
    for i, (label, value) in enumerate(metrics):
        row = 11 + i
        ws.cell(row=row, column=1, value=label)
        cell = ws.cell(row=row, column=2, value=value)
        if "Процент" in label:
            cell.number_format = "0.0%"
        ws.cell(row=row, column=1).font = Font(bold=True)
    ws.column_dimensions["B"].width = 26

    return ws


# ---------------------------------------------------------------------------
def main():
    if not PERSISTED_INPUT.is_file():
        raise SystemExit(f"Не найден файл с данными: {PERSISTED_INPUT}")

    text = PERSISTED_INPUT.read_text(encoding="utf-8")
    records = parse_records(text)

    by_prio, done, undone, overdue, checks = build_summary(records)

    # Перекрёстная проверка с счётчиками сервера.
    print("=== Проверка счётчиков (сервер vs данные) ===")
    for k in EXPECTED:
        flag = "OK" if checks[k] == EXPECTED[k] else "MISMATCH"
        print(f"  {k}: данные={checks[k]} сервер={EXPECTED[k]} {flag}")
    assert checks == EXPECTED, (checks, EXPECTED)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_tasks_sheet(wb, records)
    write_summary_sheet(wb, by_prio, done, undone, overdue, checks)

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)
    print(f"=== Готово ===")
    print(f"  Файл: {OUTPUT_XLSX}")
    print(f"  Задач всего: {checks['total']} | выполнено: {done} | невыполнено: {undone}")
    print(f"  Просроченных невыполненных: {len(overdue)}")
    print(f"  По приоритетам: high={checks['high']} normal={checks['normal']} low={checks['low']}")
    print(f"  По статусам: high={by_prio.get('high', [None])[0] and sum(1 for x in by_prio.get('high',[]) if x['done'])} "
          f"normal={sum(1 for x in by_prio.get('normal',[]) if x['done'])} "
          f"low={sum(1 for x in by_prio.get('low',[]) if x['done'])}")


if __name__ == "__main__":
    main()