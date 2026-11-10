#!/usr/bin/env python3
"""Генератор аналитического Excel-отчёта по задачам TODO."""
import argparse
import json
import urllib.request
from collections import Counter
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Цветовая палитра оформления
COLOR_PRIMARY = "1B365D"       # Глубокий темно-синий для главных заголовков
COLOR_SECONDARY = "4A6572"     # Приглушенный сине-серый для подзаголовков
COLOR_ACCENT = "E8EEF5"        # Светлый акцент для шапок таблиц
COLOR_ZEBRA = "F7F9FB"         # Чередование строк
COLOR_BORDER = "D1D5DB"        # Нейтральная граница
COLOR_ALERT = "C0392B"         # Акцент для просрочки

FONT_TITLE = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="1B365D")
FONT_BOLD = Font(name="Calibri", size=10, bold=True)
FONT_REGULAR = Font(name="Calibri", size=10)
FONT_MUTED = Font(name="Calibri", size=10, color="6B7280")
FONT_OVERDUE = Font(name="Calibri", size=10, color=COLOR_ALERT, bold=True)

FILL_PRIMARY = PatternFill(start_color=COLOR_PRIMARY, end_color=COLOR_PRIMARY, fill_type="solid")
FILL_SECONDARY = PatternFill(start_color=COLOR_SECONDARY, end_color=COLOR_SECONDARY, fill_type="solid")
FILL_HEADER = PatternFill(start_color=COLOR_ACCENT, end_color=COLOR_ACCENT, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid")

THIN_BORDER_SIDE = Side(border_style="thin", color=COLOR_BORDER)
THICK_BOTTOM_SIDE = Side(border_style="medium", color=COLOR_PRIMARY)
DOUBLE_BOTTOM_SIDE = Side(border_style="double", color=COLOR_PRIMARY)

BORDER_CELL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
BORDER_HEADER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THICK_BOTTOM_SIDE)
BORDER_TOTAL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=DOUBLE_BOTTOM_SIDE)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

PRIORITY_RU = {
    "high": "Высокий",
    "normal": "Обычный",
    "low": "Низкий",
}


def validate_tasks(tasks: List[Dict[str, Any]]) -> None:
    """Валидирует структуру и типы полей списка задач."""
    seen_ids = set()
    for task in tasks:
        if not isinstance(task, dict):
            raise ValueError("Элемент списка задач должен быть словарем")
        for key in ("id", "title", "priority", "due", "done"):
            if key not in task:
                raise ValueError(f"В задаче ID={task.get('id')} отсутствует обязательное поле '{key}'")
        if task["id"] in seen_ids:
            raise ValueError(f"Обнаружен дубликат ID задачи: {task['id']}")
        seen_ids.add(task["id"])

        if task["priority"] not in ("low", "normal", "high"):
            raise ValueError(f"Недопустимый приоритет '{task['priority']}' в задаче {task['id']}")
        if not isinstance(task["done"], bool):
            raise ValueError(f"Поле 'done' должно быть булевым в задаче {task['id']}")

        due = task.get("due")
        if due is not None:
            try:
                parsed_due = date.fromisoformat(due)
                if parsed_due.isoformat() != due:
                    raise ValueError
            except (ValueError, TypeError):
                raise ValueError(f"Срок '{due}' задачи {task['id']} не соответствует формату YYYY-MM-DD")


def build_report_data(tasks: List[Dict[str, Any]], as_of: date) -> Dict[str, Any]:
    """Рассчитывает агрегированные статистические метрики."""
    total = len(tasks)
    completed = sum(1 for t in tasks if t["done"])
    in_progress = total - completed
    completion_rate = round((completed / total * 100), 2) if total > 0 else 0.0

    as_of_str = as_of.isoformat()
    overdue_count = 0
    upcoming_7d_count = 0
    with_deadline_count = 0
    without_deadline_count = 0

    seven_days_later = (as_of + timedelta(days=7)).isoformat()

    priority_stats = {
        p: {"total": 0, "done": 0, "in_progress": 0, "overdue": 0}
        for p in ("high", "normal", "low")
    }

    processed_rows = []
    for t in sorted(tasks, key=lambda x: x["id"]):
        p = t["priority"]
        priority_stats[p]["total"] += 1
        is_done = t["done"]
        due_val = t.get("due")

        is_overdue = False
        days_diff = None

        if due_val is not None:
            with_deadline_count += 1
            days_diff = (date.fromisoformat(due_val) - as_of).days
            if not is_done and due_val < as_of_str:
                is_overdue = True
                overdue_count += 1
                priority_stats[p]["overdue"] += 1
            if not is_done and as_of_str <= due_val <= seven_days_later:
                upcoming_7d_count += 1
        else:
            without_deadline_count += 1

        if is_done:
            priority_stats[p]["done"] += 1
        else:
            priority_stats[p]["in_progress"] += 1

        processed_rows.append({
            "id": t["id"],
            "title": str(t["title"]),
            "priority_key": p,
            "priority_label": PRIORITY_RU.get(p, p),
            "due": due_val,
            "done": is_done,
            "status_label": "Выполнено" if is_done else "В работе",
            "is_overdue": is_overdue,
            "overdue_label": "Да" if is_overdue else ("Нет" if due_val else "—"),
            "days_diff": days_diff,
        })

    overdue_active_rate = round((overdue_count / in_progress * 100), 2) if in_progress > 0 else 0.0

    return {
        "as_of": as_of_str,
        "total": total,
        "completed": completed,
        "in_progress": in_progress,
        "completion_rate": completion_rate,
        "overdue": overdue_count,
        "overdue_active_rate": overdue_active_rate,
        "with_deadline": with_deadline_count,
        "without_deadline": without_deadline_count,
        "upcoming_7d": upcoming_7d_count,
        "priorities": priority_stats,
        "rows": processed_rows,
    }


def create_report(tasks: List[Dict[str, Any]], as_of: date, output_path: str | Path) -> Dict[str, Any]:
    """Создает Excel-файл отчёта с листами 'Реестр задач' и 'Сводная аналитика'."""
    validate_tasks(tasks)
    data = build_report_data(tasks, as_of)

    wb = Workbook()

    # -------------------------------------------------------------
    # Лист 1: Реестр задач
    # -------------------------------------------------------------
    ws_tasks = wb.active
    ws_tasks.title = "Реестр задач"
    ws_tasks.views.sheetView[0].showGridLines = True

    headers = [
        "ID",
        "Название задачи",
        "Приоритет",
        "Срок сдачи",
        "Статус",
        "Просрочена",
        "Дней до срока",
    ]
    ws_tasks.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws_tasks.cell(row=1, column=col_idx)
        cell.font = FONT_TITLE
        cell.fill = FILL_PRIMARY
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER

    ws_tasks.row_dimensions[1].height = 24

    for row_idx, r in enumerate(data["rows"], start=2):
        row_fill = FILL_ZEBRA if row_idx % 2 == 1 else None

        days_text = "—"
        if r["days_diff"] is not None:
            days_text = f"+{r['days_diff']}" if r["days_diff"] > 0 else str(r["days_diff"])

        row_values = [
            r["id"],
            r["title"],
            r["priority_label"],
            r["due"] if r["due"] else "Без срока",
            r["status_label"],
            r["overdue_label"],
            days_text,
        ]
        ws_tasks.append(row_values)

        ws_tasks.row_dimensions[row_idx].height = 20

        # Стилизация и экранирование ячеек строки
        for col_idx in range(1, len(headers) + 1):
            cell = ws_tasks.cell(row=row_idx, column=col_idx)
            cell.border = BORDER_CELL
            if row_fill:
                cell.fill = row_fill

            if col_idx == 1:
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_REGULAR
            elif col_idx == 2:
                cell.alignment = ALIGN_LEFT
                cell.font = FONT_REGULAR
                # Защита от Formula Injection: принудительно задаем строковый тип
                cell.data_type = "s"
            elif col_idx in (3, 4, 5, 7):
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_REGULAR
            elif col_idx == 6:
                cell.alignment = ALIGN_CENTER
                cell.font = FONT_OVERDUE if r["is_overdue"] else FONT_REGULAR

    ws_tasks.freeze_panes = "A2"
    ws_tasks.auto_filter.ref = ws_tasks.dimensions

    col_widths = {
        "A": 10,
        "B": 50,
        "C": 16,
        "D": 16,
        "E": 16,
        "F": 14,
        "G": 16,
    }
    for col_letter, width in col_widths.items():
        ws_tasks.column_dimensions[col_letter].width = width

    # -------------------------------------------------------------
    # Лист 2: Сводная аналитика
    # -------------------------------------------------------------
    ws_summary = wb.create_sheet(title="Сводная аналитика")
    ws_summary.views.sheetView[0].showGridLines = True

    # Заголовок раздела 1
    ws_summary.cell(row=2, column=1, value="КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ ЭФФЕКТИВНОСТИ (KPI)").font = FONT_SECTION
    ws_summary.cell(row=2, column=2, value="")
    for col in (1, 2):
        ws_summary.cell(row=2, column=col).fill = FILL_PRIMARY
        ws_summary.cell(row=2, column=col).border = BORDER_HEADER
    ws_summary.row_dimensions[2].height = 22

    kpi_rows = [
        ("Дата формирования отчёта", data["as_of"]),
        ("Всего задач в системе", data["total"]),
        ("Выполнено задач (Done)", data["completed"]),
        ("В работе (Active / Pending)", data["in_progress"]),
        ("Процент выполнения (%)", f"{data['completion_rate']}%"),
        ("Просрочено невыполненных задач", data["overdue"]),
        ("Доля просрочки среди активных (%)", f"{data['overdue_active_rate']}%"),
    ]

    curr_row = 3
    for label, val in kpi_rows:
        c1 = ws_summary.cell(row=curr_row, column=1, value=label)
        c2 = ws_summary.cell(row=curr_row, column=2, value=val)
        c1.font = FONT_REGULAR
        c1.alignment = ALIGN_LEFT
        c1.border = BORDER_CELL
        c2.font = FONT_BOLD if "Процент" in label or "Просрочено" in label else FONT_REGULAR
        c2.alignment = ALIGN_RIGHT
        c2.border = BORDER_CELL
        ws_summary.row_dimensions[curr_row].height = 19
        curr_row += 1

    # Заголовок раздела 2: Матрица приоритетов
    curr_row += 1
    ws_summary.cell(row=curr_row, column=1, value="МАТРИЦА РАСПРЕДЕЛЕНИЯ: ПРИОРИТЕТ × СТАТУС").font = FONT_SECTION
    for c in range(1, 6):
        ws_summary.cell(row=curr_row, column=c).fill = FILL_PRIMARY
        ws_summary.cell(row=curr_row, column=c).border = BORDER_HEADER
    ws_summary.row_dimensions[curr_row].height = 22

    curr_row += 1
    matrix_headers = ["Приоритет", "Всего", "Выполнено", "В работе", "Из них просрочено"]
    for c_idx, h in enumerate(matrix_headers, start=1):
        cell = ws_summary.cell(row=curr_row, column=c_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER
    ws_summary.row_dimensions[curr_row].height = 20

    curr_row += 1
    for p_key, p_name in (("high", "Высокий (high)"), ("normal", "Обычный (normal)"), ("low", "Низкий (low)")):
        p_data = data["priorities"][p_key]
        ws_summary.cell(row=curr_row, column=1, value=p_name).alignment = ALIGN_LEFT
        ws_summary.cell(row=curr_row, column=2, value=p_data["total"]).alignment = ALIGN_RIGHT
        ws_summary.cell(row=curr_row, column=3, value=p_data["done"]).alignment = ALIGN_RIGHT
        ws_summary.cell(row=curr_row, column=4, value=p_data["in_progress"]).alignment = ALIGN_RIGHT
        c_over = ws_summary.cell(row=curr_row, column=5, value=p_data["overdue"])
        c_over.alignment = ALIGN_RIGHT
        if p_data["overdue"] > 0:
            c_over.font = FONT_OVERDUE

        for c_idx in range(1, 6):
            cell = ws_summary.cell(row=curr_row, column=c_idx)
            cell.border = BORDER_CELL
            if not cell.font or cell.font == FONT_REGULAR:
                cell.font = FONT_REGULAR
        ws_summary.row_dimensions[curr_row].height = 19
        curr_row += 1

    # Итоговая строка матрицы
    ws_summary.cell(row=curr_row, column=1, value="ИТОГО").alignment = ALIGN_LEFT
    ws_summary.cell(row=curr_row, column=2, value=data["total"]).alignment = ALIGN_RIGHT
    ws_summary.cell(row=curr_row, column=3, value=data["completed"]).alignment = ALIGN_RIGHT
    ws_summary.cell(row=curr_row, column=4, value=data["in_progress"]).alignment = ALIGN_RIGHT
    ws_summary.cell(row=curr_row, column=5, value=data["overdue"]).alignment = ALIGN_RIGHT
    for c_idx in range(1, 6):
        cell = ws_summary.cell(row=curr_row, column=c_idx)
        cell.font = FONT_BOLD
        cell.fill = FILL_ZEBRA
        cell.border = BORDER_TOTAL
    ws_summary.row_dimensions[curr_row].height = 21

    # Заголовок раздела 3: Аналитика сроков
    curr_row += 2
    ws_summary.cell(row=curr_row, column=1, value="СТРУКТУРА ДЕДЛАЙНОВ").font = FONT_SECTION
    for c in (1, 2):
        ws_summary.cell(row=curr_row, column=c).fill = FILL_PRIMARY
        ws_summary.cell(row=curr_row, column=c).border = BORDER_HEADER
    ws_summary.row_dimensions[curr_row].height = 22

    deadline_rows = [
        ("Задач с установленным сроком", data["with_deadline"]),
        ("Задач без срока выполнения", data["without_deadline"]),
        ("Просрочено (требуют срочного внимания)", data["overdue"]),
        ("Дедлайн в течение 7 дней (активные)", data["upcoming_7d"]),
    ]

    curr_row += 1
    for label, val in deadline_rows:
        c1 = ws_summary.cell(row=curr_row, column=1, value=label)
        c2 = ws_summary.cell(row=curr_row, column=2, value=val)
        c1.font = FONT_REGULAR
        c1.alignment = ALIGN_LEFT
        c1.border = BORDER_CELL
        c2.font = FONT_OVERDUE if "Просрочено" in label and val > 0 else FONT_REGULAR
        c2.alignment = ALIGN_RIGHT
        c2.border = BORDER_CELL
        ws_summary.row_dimensions[curr_row].height = 19
        curr_row += 1

    summary_widths = {"A": 36, "B": 18, "C": 14, "D": 14, "E": 20}
    for col_letter, width in summary_widths.items():
        ws_summary.column_dimensions[col_letter].width = width

    # Сохранение файла с защитой от перезаписи
    target_path = Path(output_path).resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    if target_path.exists():
        raise FileExistsError(f"Файл отчёта уже существует: {target_path}. Перезапись запрещена.")

    with target_path.open("xb") as f:
        wb.save(f)

    return {
        "as_of": data["as_of"],
        "total": data["total"],
        "completed": data["completed"],
        "in_progress": data["in_progress"],
        "completion_rate": data["completion_rate"],
        "overdue": data["overdue"],
        "overdue_active_rate": data["overdue_active_rate"],
        "output": str(target_path),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", default="http://127.0.0.1:5001", help="URL сервиса TODO API")
    parser.add_argument("--date", type=date.fromisoformat, default=date.today(), help="Дата отчёта (YYYY-MM-DD)")
    parser.add_argument("--output", required=True, help="Путь к результирующему файлу .xlsx")
    args = parser.parse_args()

    api_endpoint = args.url.rstrip("/") + "/api/tasks"
    try:
        with urllib.request.urlopen(api_endpoint, timeout=15) as resp:
            tasks = json.load(resp)
    except Exception as exc:
        raise RuntimeError(f"Не удалось получить задачи по адресу {api_endpoint}: {exc}") from exc

    if not isinstance(tasks, list):
        raise ValueError("API должен возвращать список задач в формате JSON")

    summary = create_report(tasks, args.date, args.output)
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    main()
