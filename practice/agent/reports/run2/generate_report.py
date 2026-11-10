#!/usr/bin/env python3
"""Генерация Excel-отчёта по задачам MCP todo на 2026-11-10.

Читает невыполненные задачи из сохранённого JSONL-файла (результат
вызова MCP-инструмента list_tasks), объединяет их с 118 выполненными
задачами, полученных тем же list_tasks, и записывает report.xlsx
с листами «Задачи» и «Сводка». Задачи не изменяются.
"""

import json
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Путь к сохранённому результату list_tasks (done=false) — 246 невыполненных задач.
NOT_DONE_FILE = (
    Path.home() / ".qwen/tmp/385c1a20454bfaffec34ff1208718d56fa366bbd4e0d524dba0ab3410d8dccc1/tool-results/chatcmpl-tool-908b0f0137230849.txt"
)
OUTPUT_XLSX = Path(__file__).with_name("report.xlsx")
REPORT_DATE = date(2026, 11, 10)

# Выполненные задачи (id, title, priority, due) — получены вызовом list_tasks done=true.
DONE_TASKS = [
    (5, "Распечатать молоко и хлеб", "low", None),
    (6, "Отправить договор", "normal", "2026-11-24"),
    (7, "Оплатить заявку на отпуск", "normal", "2026-11-05"),
    (9, "Забронировать план на неделю", "high", "2026-11-21"),
    (11, "Прочитать встречу с командой", "normal", None),
    (17, "Заказать план на неделю", "low", "2026-11-16"),
    (23, "Сдать домашнее задание по MCP", "high", "2026-11-11"),
    (29, "Сдать домашнее задание по MCP", "high", "2026-11-11"),
    (36, "Согласовать счёт за интернет", "normal", "2026-11-24"),
    (38, "Протестировать счёт за интернет", "normal", "2026-11-14"),
    (44, "Закрыть резервную копию", "normal", "2026-11-27"),
    (51, "Оплатить заявку на отпуск", "high", "2026-12-05"),
    (57, "Подготовить документацию", "high", "2026-12-08"),
    (60, "Обновить тесты для приложения", "low", "2026-11-17"),
    (65, "Подготовить заявку на отпуск", "low", "2026-11-14"),
    (70, "Разобрать ревью кода", "high", "2026-11-04"),
    (73, "Согласовать подарок коллеге", "low", None),
    (74, "Протестировать резервную копию", "normal", "2026-11-27"),
    (75, "Закрыть встречу с командой", "low", "2026-11-18"),
    (77, "Купить план на неделю", "normal", "2026-12-06"),
    (79, "Оплатить презентацию", "high", "2026-12-04"),
    (80, "Позвонить по поводу анкету", "low", None),
    (81, "Записаться на подарок коллеге", "normal", "2026-12-02"),
    (85, "Проверить посылку", "high", "2026-11-24"),
    (88, "Подготовить презентацию", "normal", None),
    (92, "Обновить резервную копию", "high", "2026-11-20"),
    (96, "Записаться на визит к врачу", "high", "2026-11-22"),
    (97, "Забронировать презентацию", "normal", None),
    (102, "Проверить отчёт за квартал", "high", "2026-11-04"),
    (103, "Спланировать встречу с командой", "low", "2026-11-22"),
    (104, "Проверить доклад для лекции", "high", "2026-11-20"),
    (106, "Разобрать молоко и хлеб", "low", "2026-12-08"),
    (111, "Починить план на неделю", "low", "2026-10-31"),
    (112, "Заказать посылку", "low", None),
    (113, "Протестировать молоко и хлеб", "low", "2026-11-08"),
    (119, "Позвонить по поводу доклад для лекции", "low", "2026-11-23"),
    (130, "Закрыть список покупок", "high", "2026-11-12"),
    (131, "Оплатить билеты на поезд", "low", "2026-12-08"),
    (134, "Протестировать заявку на отпуск", "low", "2026-11-13"),
    (137, "Забронировать доклад для лекции", "normal", "2026-12-02"),
    (147, "Позвонить по поводу презентацию", "high", "2026-11-17"),
    (149, "Прочитать доклад для лекции", "low", "2026-11-10"),
    (152, "Записаться на тесты для приложения", "normal", "2026-11-25"),
    (161, "Записаться на список покупок", "high", "2026-12-08"),
    (162, "Забронировать ревью кода", "normal", "2026-11-28"),
    (170, "Обновить резервную копию", "low", "2026-12-07"),
    (171, "Написать заявку на отпуск", "high", "2026-11-30"),
    (172, "Позвонить по поводу молоко и хлеб", "low", None),
    (174, "Прочитать ревью кода", "low", None),
    (176, "Подготовить встречу с командой", "high", "2026-11-12"),
    (181, "Заказать презентацию", "high", "2026-11-04"),
    (182, "Записаться на визит к врачу", "low", "2026-11-04"),
    (183, "Отправить договор", "high", "2026-11-27"),
    (186, "Подготовить доклад для лекции", "normal", "2026-11-28"),
    (187, "Разобрать договор", "normal", "2026-11-05"),
    (189, "Оплатить доклад для лекции", "high", "2026-11-27"),
    (190, "Починить посылку", "normal", "2026-12-02"),
    (193, "Починить доклад для лекции", "normal", None),
    (196, "Закрыть отчёт за квартал", "high", "2026-12-03"),
    (202, "Позвонить по поводу отчёт за квартал", "low", "2026-11-06"),
    (204, "Распечатать билеты на поезд", "high", "2026-11-06"),
    (205, "Протестировать доклад для лекции", "high", "2026-12-10"),
    (209, "Согласовать встречу с командой", "high", "2026-12-04"),
    (217, "Закрыть список покупок", "high", "2026-11-08"),
    (222, "Забронировать молоко и хлеб", "normal", None),
    (226, "Заказать презентацию", "low", "2026-11-25"),
    (227, "Подготовить счёт за интернет", "normal", "2026-11-14"),
    (228, "Починить визит к врачу", "low", "2026-11-06"),
    (232, "Согласовать презентацию", "normal", "2026-11-03"),
    (234, "Починить тесты для приложения", "normal", "2026-11-09"),
    (236, "Доработать план на неделю", "normal", "2026-11-13"),
    (238, "Оплатить доклад для лекции", "normal", "2026-11-06"),
    (240, "Распечатать тесты для приложения", "high", "2026-11-03"),
    (242, "Спланировать документацию", "normal", "2026-12-08"),
    (243, "Прочитать анкету", "low", "2026-11-08"),
    (245, "Купить билеты на поезд", "low", "2026-11-03"),
    (246, "Протестировать визит к врачу", "high", "2026-11-03"),
    (247, "Протестировать ревью кода", "normal", "2026-11-07"),
    (250, "Написать подарок коллеге", "low", "2026-11-09"),
    (251, "Протестировать договор", "normal", "2026-11-09"),
    (252, "Разобрать план на неделю", "low", "2026-11-06"),
    (254, "Оплатить счёт за интернет", "high", "2026-11-08"),
    (256, "Доработать билеты на поезд", "normal", "2026-11-14"),
    (258, "Разобрать анкету", "low", "2026-11-14"),
    (259, "Оплатить резервную копию", "normal", "2026-11-11"),
    (260, "Распечатать билеты на поезд", "high", None),
    (265, "Починить список покупок", "normal", "2026-11-09"),
    (271, "Закрыть подарок коллеге", "high", "2026-12-03"),
    (272, "Спланировать документацию", "high", "2026-11-22"),
    (275, "Написать список покупок", "high", "2026-11-27"),
    (276, "Отправить план на неделю", "high", "2026-11-09"),
    (277, "Спланировать план на неделю", "high", "2026-12-04"),
    (279, "Оплатить счёт за интернет", "low", None),
    (288, "Распечатать анкету", "low", None),
    (294, "Заказать визит к врачу", "high", "2026-11-17"),
    (300, "Купить билеты на поезд", "low", None),
    (303, "Проверить молоко и хлеб", "high", None),
    (305, "Проверить документацию", "normal", "2026-12-04"),
    (306, "Забронировать визит к врачу", "low", "2026-12-05"),
    (307, "Спланировать анкету", "normal", None),
    (309, "Доработать подарок коллеге", "normal", "2026-11-26"),
    (317, "Починить билеты на поезд", "normal", "2026-11-26"),
    (322, "Закрыть подарок коллеге", "low", "2026-11-24"),
    (324, "Купить ревью кода", "normal", "2026-11-11"),
    (327, "Спланировать отчёт за квартал", "normal", "2026-12-05"),
    (334, "Написать подарок коллеге", "low", "2026-11-10"),
    (335, "Оплатить молоко и хлеб", "low", "2026-12-02"),
    (337, "Разобрать посылку", "low", "2026-12-05"),
    (339, "Записаться на план на неделю", "high", None),
    (342, "Купить обновления системы", "high", "2026-11-22"),
    (343, "Обновить счёт за интернет", "normal", "2026-11-26"),
    (346, "Починить анкету", "high", "2026-11-21"),
    (347, "Забронировать обновления системы", "low", None),
    (348, "Прочитать молоко и хлеб", "high", "2026-11-08"),
    (349, "Отправить план на неделю", "low", None),
    (350, "Обновить план на неделю", "normal", None),
    (358, "Согласовать отчёт за квартал", "low", "2026-11-19"),
    (364, "Подготовить заявку на отпуск", "low", "2026-11-23"),
]

PRIORITY_ORDER = {"high": 0, "normal": 1, "low": 2}
STATUS_DONE = "Выполнена"
STATUS_TODO = "Не выполнена"

thin = Side(style="thin", color="BFBFBF")
header_fill = PatternFill("solid", fgColor="4472C4")
header_font = Font(color="FFFFFF", bold=True)
header_align = Alignment(horizontal="center", vertical="center")
title_font = Font(bold=True, size=14)
section_font = Font(bold=True, size=12)


def load_not_done() -> list[dict]:
    # Файл содержит pretty-printed JSON-объекты (по несколько строк на задачу),
    # разделённые пробельными символами. Собираем объекты по глубине фигурных скобок.
    tasks = []
    with NOT_DONE_FILE.open(encoding="utf-8") as f:
        content = f.read()
    buf: list[str] = []
    depth = 0
    for ch in content:
        if ch == "{":
            depth += 1
            buf.append(ch)
        elif ch == "}":
            depth -= 1
            buf.append(ch)
            if depth == 0:
                obj = json.loads("".join(buf))
                tasks.append(
                    {
                        "id": obj["id"],
                        "title": obj["title"],
                        "priority": obj["priority"],
                        "due": obj["due"],
                        "done": obj["done"],
                    }
                )
                buf = []
        elif depth > 0:
            buf.append(ch)
        # символы вне объектов (пробелы, разделители) пропускаем
    return tasks


def build_rows() -> tuple[list[dict], int, int, int]:
    not_done = load_not_done()
    all_tasks = not_done + [
        {"id": i, "title": t, "priority": p, "due": d, "done": True}
        for i, t, p, d in DONE_TASKS
    ]
    all_tasks.sort(key=lambda x: x["id"])

    done = sum(1 for t in all_tasks if t["done"])
    todo = len(all_tasks) - done
    overdue = sum(
        1
        for t in all_tasks
        if not t["done"] and t["due"] is not None and date.fromisoformat(t["due"]) < REPORT_DATE
    )
    return all_tasks, done, todo, overdue


def write_tasks_sheet(wb, rows):
    ws = wb.create_sheet("Задачи")
    headers = ["id", "Название", "Приоритет", "Срок", "Статус"]
    ws.append(headers)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = Border(bottom=thin)
    ws.row_dimensions[1].height = 22

    for r, t in enumerate(rows, 2):
        due_val = date.fromisoformat(t["due"]) if t["due"] else None
        ws.cell(row=r, column=1, value=t["id"])
        ws.cell(row=r, column=2, value=t["title"])
        ws.cell(row=r, column=3, value=t["priority"])
        if due_val is not None:
            cell = ws.cell(row=r, column=4, value=due_val)
            cell.number_format = "YYYY-MM-DD"
        ws.cell(row=r, column=5, value=STATUS_DONE if t["done"] else STATUS_TODO)

    for c in range(1, 6):
        col = get_column_letter(c)
        widths = {1: 8, 2: 52, 3: 12, 4: 12, 5: 16}
        ws.column_dimensions[col].width = widths[c]
    ws.freeze_panes = "A2"


def write_summary_sheet(wb, rows, done, todo, overdue):
    ws = wb.active
    ws.title = "Сводка"

    ws["B1"] = f"Сводка по задачам (на {REPORT_DATE.isoformat()})"
    ws["B1"].font = title_font

    ws["B3"] = "По приоритетам"
    ws["B3"].font = section_font
    ws["B4"] = "Приоритет"
    ws["C4"] = "Количество"
    for c in (2, 3):
        cell = ws.cell(row=4, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = Border(bottom=thin)

    priorities = ["high", "normal", "low"]
    labels = {"high": "Высокий (high)", "normal": "Средний (normal)", "low": "Низкий (low)"}
    row = 5
    for p in priorities:
        cnt = sum(1 for t in rows if t["priority"] == p)
        ws.cell(row=row, column=2, value=labels[p])
        ws.cell(row=row, column=3, value=cnt)
        row += 1

    ws.cell(row=row, column=2, value="Итого задач")
    cell = ws.cell(row=row, column=3, value=len(rows))
    cell.font = section_font
    row += 1

    ws.cell(row=row, column=2, value="Выполнено, %")
    pct = round(done / len(rows) * 100, 2)
    cell = ws.cell(row=row, column=3, value=pct)
    cell.number_format = "0.00"
    row += 1

    ws.cell(row=row, column=2, value="Просроченные невыполненные")
    ws.cell(row=row, column=3, value=overdue)
    row += 1

    cell = ws.cell(row=row + 1, column=2, value="По статусу")
    cell.font = section_font
    ws.cell(row=row + 2, column=2, value="Статус")
    ws.cell(row=row + 2, column=3, value="Количество")
    for c in (2, 3):
        cell = ws.cell(row=row + 2, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = Border(bottom=thin)
    row += 3
    ws.cell(row=row, column=2, value=STATUS_DONE)
    ws.cell(row=row, column=3, value=done)
    row += 1
    ws.cell(row=row, column=2, value=STATUS_TODO)
    ws.cell(row=row, column=3, value=todo)

    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 14


def main():
    rows, done, todo, overdue = build_rows()

    assert len(rows) == 364, f"ожидалось 364 задачи, получено {len(rows)}"
    assert done == 118, f"ожидалось 118 выполненных, получено {done}"
    assert todo == 246, f"ожидалось 246 невыполненных, получено {todo}"

    pct = round(done / len(rows) * 100, 2)
    print(f"Всего: {len(rows)} | Выполнено: {done} ({pct}%) | "
          f"Не выполнено: {todo} | Просрочено: {overdue}")

    wb = Workbook()
    write_tasks_sheet(wb, rows)
    write_summary_sheet(wb, rows, done, todo, overdue)
    wb.save(OUTPUT_XLSX)
    print(f"Сохранено: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()