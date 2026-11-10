#!/usr/bin/env python3
"""Генерация Excel-отчёта по задачам TODO-приложения."""

import json
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

# Пути (скрипт в homework/vague2/, БД в apps/ — 5 уровней до корня проекта)
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
DB_PATH = PROJECT_ROOT / "apps" / "todo_db.json"
REPORT_PATH = Path(__file__).parent / "report.xlsx"

PRIORITIES = ("low", "normal", "high")
PRIORITY_RU = {"low": "Низкий", "normal": "Средний", "high": "Высокий"}
DONE_RU = {True: "Выполнено", False: "В работе"}


def load_tasks(db_path: Path) -> list[dict]:
    """Загрузить задачи из JSON-файла БД (payload.tasks)."""
    doc = json.loads(db_path.read_text(encoding="utf-8"))
    return doc["payload"]["tasks"]


def analyze(tasks: list[dict]) -> dict:
    """Сводная статистика."""
    total = len(tasks)
    done = sum(1 for t in tasks if t["done"])
    pending = total - done
    today = date.today().isoformat()

    overdue = [
        t for t in tasks
        if not t["done"] and t["due"] and t["due"] < today
    ]
    by_priority = {p: {"total": 0, "done": 0} for p in PRIORITIES}
    for t in tasks:
        p = t["priority"]
        by_priority[p]["total"] += 1
        if t["done"]:
            by_priority[p]["done"] += 1

    by_due = {"no_due": 0, "upcoming": 0, "overdue": len(overdue)}
    for t in tasks:
        if t["done"]:
            continue
        if not t["due"]:
            by_due["no_due"] += 1
        elif t["due"] >= today:
            by_due["upcoming"] += 1

    return {
        "total": total,
        "done": done,
        "pending": pending,
        "overdue": len(overdue),
        "overdue_tasks": overdue,
        "by_priority": by_priority,
        "by_due": by_due,
    }


def create_summary_sheet(wb: Workbook, stats: dict) -> None:
    """Лист с сводкой."""
    ws = wb.create_sheet("Сводка")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18

    title_font = Font(bold=True, size=18, color="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    stat_fill = PatternFill("solid", fgColor="DDEBF7")

    # Заголовок
    ws["A1"] = "Отчёт по задачам TODO"
    ws["A1"].font = title_font
    ws["A2"] = f"Сгенерирован: {date.today().strftime('%d.%m.%Y')}"
    ws["A2"].font = Font(italic=True, color="666666")

    # Ключевые показатели
    rows = [
        ("Всего задач", stats["total"], "🎯"),
        ("Выполнено", stats["done"], "✅"),
        ("В работе", stats["pending"], "🔄"),
        ("Просрочено", stats["overdue"], "⚠️"),
    ]
    start_row = 4
    for i, (label, value, icon) in enumerate(rows):
        r = start_row + i
        ws.cell(row=r, column=1, value=f"{icon} {label}")
        ws.cell(row=r, column=2, value=value)
        for col in (1, 2):
            cell = ws.cell(row=r, column=col)
            cell.fill = stat_fill
            cell.alignment = Alignment(horizontal="left" if col == 1 else "center")
        ws.cell(row=r, column=2).font = Font(bold=True)

    # Приоритеты
    p_start = start_row + len(rows) + 2
    ws.cell(row=p_start, column=1, value="Приоритеты").font = header_font
    ws.cell(row=p_start, column=1).fill = header_fill
    ws.cell(row=p_start, column=2).fill = header_fill

    p_headers = ["Приоритет", "Всего", "Выполнено", "Осталось"]
    for col, h in enumerate(p_headers, start=1):
        c = ws.cell(row=p_start + 1, column=col, value=h)
        c.font = header_font
        c.fill = header_fill

    for i, p in enumerate(PRIORITIES):
        r = p_start + 2 + i
        d = stats["by_priority"][p]
        ws.cell(row=r, column=1, value=PRIORITY_RU[p])
        ws.cell(row=r, column=2, value=d["total"])
        ws.cell(row=r, column=3, value=d["done"])
        ws.cell(row=r, column=4, value=d["total"] - d["done"])
        for col in range(1, 5):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="center" if col > 1 else "left")

    # Сроки
    d_start = p_start + 3 + len(PRIORITIES) + 2
    ws.cell(row=d_start, column=1, value="Сроки").font = header_font
    ws.cell(row=d_start, column=1).fill = header_fill
    ws.cell(row=d_start, column=2).fill = header_fill

    d_headers = ["Категория", "Кол-во"]
    for col, h in enumerate(d_headers, start=1):
        c = ws.cell(row=d_start + 1, column=col, value=h)
        c.font = header_font
        c.fill = header_fill

    due_labels = {
        "no_due": "Без срока",
        "upcoming": "Скоро (не просрочены)",
        "overdue": "Просроченные",
    }
    for i, (key, label) in enumerate(due_labels.items()):
        r = d_start + 2 + i
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=stats["by_due"][key])

    # Комментарий
    comment_row = d_start + 2 + 3 + 1
    ws.cell(row=comment_row, column=1,
            value="⚠️ Просроченные задачи требуют внимания — срок истёк, но задача не выполнена.")


def create_tasks_sheet(wb: Workbook, tasks: list[dict]) -> None:
    """Лист со всеми задачами."""
    ws = wb.create_sheet("Все задачи")
    ws.sheet_view.showGridLines = False

    headers = ["ID", "Название", "Статус", "Приоритет", "Срок", "Создан"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")

    for i, t in enumerate(tasks, start=2):
        ws.cell(row=i, column=1, value=t["id"])
        ws.cell(row=i, column=2, value=t["title"])
        ws.cell(row=i, column=3, value=DONE_RU[t["done"]])
        ws.cell(row=i, column=4, value=PRIORITY_RU[t["priority"]])
        ws.cell(row=i, column=5, value=t["due"] or "—")
        ws.cell(row=i, column=6, value=t["created_at"][:10])

        if t["done"]:
            for col in range(1, 7):
                ws.cell(row=i, column=col).font = Font(color="2E7D32")
        elif t["due"] and t["due"] < date.today().isoformat():
            for col in range(1, 7):
                ws.cell(row=i, column=col).font = Font(color="C62828")

    # Ширина колонок
    widths = [6, 50, 14, 14, 12, 12]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # Заморозка заголовка
    ws.freeze_panes = "A2"


def create_overdue_sheet(wb: Workbook, overdue: list[dict]) -> None:
    """Лист с просроченными задачами."""
    ws = wb.create_sheet("Просроченные")
    ws.sheet_view.showGridLines = False

    if not overdue:
        ws["A1"] = "Просроченных задач нет! 🎉"
        ws["A1"].font = Font(bold=True, size=14, color="2E7D32")
        return

    headers = ["ID", "Название", "Приоритет", "Срок", "Дней просрочки"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="C62828")
        c.alignment = Alignment(horizontal="center")

    today = date.today()
    for i, t in enumerate(overdue, start=2):
        due_date = date.fromisoformat(t["due"])
        days_over = (today - due_date).days
        ws.cell(row=i, column=1, value=t["id"])
        ws.cell(row=i, column=2, value=t["title"])
        ws.cell(row=i, column=3, value=PRIORITY_RU[t["priority"]])
        ws.cell(row=i, column=4, value=t["due"])
        ws.cell(row=i, column=5, value=days_over)

        for col in range(1, 6):
            cell = ws.cell(row=i, column=col)
            cell.font = Font(color="C62828", bold=(col == 5))

    widths = [6, 50, 14, 12, 14]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"


def create_charts(wb: Workbook, stats: dict) -> None:
    """Добавить диаграммы на отдельный лист."""
    ws = wb.create_sheet("Диаграммы")

    # Bar: приоритеты
    bar = BarChart()
    bar.title = "Задачи по приоритету"
    bar.y_axis.title = "Кол-во задач"
    bar.x_axis.title = "Приоритет"

    ws.cell(row=1, column=1, value="Приоритет")
    ws.cell(row=1, column=2, value="Всего")
    ws.cell(row=1, column=3, value="Выполнено")

    rows = []
    for p in PRIORITIES:
        d = stats["by_priority"][p]
        rows.append((PRIORITY_RU[p], d["total"], d["done"]))
    for i, (label, total, done) in enumerate(rows):
        r = 2 + i
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=total)
        ws.cell(row=r, column=3, value=done)

    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=1 + len(rows))
    cats = Reference(ws, min_col=1, min_row=2, max_row=1 + len(rows))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 10
    bar.width = 20
    ws.add_chart(bar, "F2")

    # Pie: статус выполнения
    pie = PieChart()
    pie.title = "Доля выполненных задач"

    pie_row = 2 + len(rows) + 1
    ws.cell(row=pie_row, column=1, value="Выполнено")
    ws.cell(row=pie_row, column=2, value=stats["done"])
    ws.cell(row=pie_row + 1, column=1, value="В работе")
    ws.cell(row=pie_row + 1, column=2, value=stats["pending"])

    pie_data = Reference(ws, min_col=2, min_row=pie_row, max_row=pie_row + 1)
    pie_cats = Reference(ws, min_col=1, min_row=pie_row, max_row=pie_row + 1)
    pie.add_data(pie_data, titles_from_data=False)
    pie.set_categories(pie_cats)
    pie.height = 10
    pie.width = 20
    ws.add_chart(pie, "F24")


def main():
    tasks = load_tasks(DB_PATH)
    stats = analyze(tasks)

    wb = Workbook()
    # Удалить дефолтный лист
    wb.remove(wb.active)

    create_summary_sheet(wb, stats)
    create_tasks_sheet(wb, tasks)
    create_overdue_sheet(wb, stats["overdue_tasks"])
    create_charts(wb, stats)

    wb.save(REPORT_PATH)
    print(f"✅ Отчёт сохранён: {REPORT_PATH}")
    print(f"   Всего задач: {stats['total']}, выполнено: {stats['done']}, "
          f"просрочено: {stats['overdue']}")


if __name__ == "__main__":
    main()