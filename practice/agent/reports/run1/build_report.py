import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

RESULTS_FILE = (
    Path.home()
    / ".qwen/tmp/385c1a20454bfaffec34ff1208718d56fa366bbd4e0d524dba0ab3410d8dccc1/tool-results/chatcmpl-tool-88300ef0284c3a20.txt"
)
REPORT_FILE = Path("reports/run1/report.xlsx")
AS_OF = "2026-11-10"
PRIORITIES = ("high", "normal", "low")
STATUSES = ("выполнена", "не выполнена")

bold = Font(bold=True)


def parse_tasks(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8-sig")
    tasks = []
    decoder = json.JSONDecoder()
    idx = 0
    n = len(text)
    while idx < n:
        while idx < n and text[idx].isspace():
            idx += 1
        if idx >= n:
            break
        obj, end = decoder.raw_decode(text, idx)
        tasks.append(obj)
        idx = end
    return [t for t in tasks if isinstance(t, dict)]


def main() -> None:
    tasks = parse_tasks(RESULTS_FILE)
    assert tasks, f"не удалось разобрать задачи из {RESULTS_FILE}"

    by_priority: dict[str, int] = {}
    by_status: dict[str, int] = {}
    done = 0
    overdue = 0
    unexpected: set[str] = set()

    for t in tasks:
        p = t.get("priority")
        by_priority[p] = by_priority.get(p, 0) + 1
        if p not in PRIORITIES:
            unexpected.add(p)
        s = "выполнена" if t.get("done") else "не выполнена"
        by_status[s] = by_status.get(s, 0) + 1
        if t.get("done"):
            done += 1
            continue
        due = t.get("due")
        if due and due < AS_OF:
            overdue += 1

    total = len(tasks)
    pct = round(done / total * 100, 1) if total else 0.0

    wb = Workbook()

    # Лист «Задачи»
    ws = wb.active
    ws.title = "Задачи"
    for col, header in enumerate(("id", "название", "приоритет", "срок", "статус"), 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = bold
    ws.freeze_panes = "A2"
    row = 2
    for t in sorted(tasks, key=lambda x: x.get("id", 0)):
        ws.cell(row=row, column=1, value=t.get("id"))
        ws.cell(row=row, column=2, value=t.get("title"))
        ws.cell(row=row, column=3, value=t.get("priority"))
        ws.cell(row=row, column=4, value=t.get("due") or "")
        ws.cell(row=row, column=5, value="выполнена" if t.get("done") else "не выполнена")
        row += 1
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 16

    # Лист «Сводка»
    sw = wb.create_sheet("Сводка")
    sw["A1"] = f"Сводка по задачам (на дату {AS_OF})"
    sw["A1"].font = bold
    sw["A3"] = "Итоги"
    sw["A3"].font = bold
    summary = [
        ("Всего задач", total),
        ("Выполнено", done),
        ("Не выполнено", total - done),
        ("Процент выполненных, %", pct),
        ("Просрочено (невыполненные, срок раньше 2026-11-10)", overdue),
    ]
    for r, (label, value) in enumerate(summary, 4):
        sw.cell(row=r, column=1, value=label)
        cell = sw.cell(row=r, column=2, value=value)
        if label.startswith("Процент"):
            cell.number_format = "0.0"
    sw.column_dimensions["A"].width = 52
    sw.column_dimensions["B"].width = 12

    r = 10
    sw.cell(row=r, column=1, value="По приоритетам").font = bold
    r += 1
    sw.cell(row=r, column=1, value="Приоритет").font = bold
    sw.cell(row=r, column=2, value="Количество").font = bold
    r += 1
    for p in PRIORITIES:
        sw.cell(row=r, column=1, value=p)
        sw.cell(row=r, column=2, value=by_priority.get(p, 0))
        r += 1

    r += 1
    sw.cell(row=r, column=1, value="По статусам").font = bold
    r += 1
    sw.cell(row=r, column=1, value="Статус").font = bold
    sw.cell(row=r, column=2, value="Количество").font = bold
    r += 1
    for s in STATUSES:
        sw.cell(row=r, column=1, value=s)
        sw.cell(row=r, column=2, value=by_status.get(s, 0))
        r += 1

    wb.save(REPORT_FILE)

    print(f"всего задач: {total}")
    print(f"выполнено: {done} ({pct}%)")
    print(f"не выполнено: {total - done}")
    print(f"просрочено (срок раньше {AS_OF}): {overdue}")
    print("по приоритетам:", by_priority)
    print("по статусам:", by_status)
    if unexpected:
        print("неожиданные приоритеты:", sorted(unexpected))
    print(f"отчёт сохранён: {REPORT_FILE.resolve()}")


if __name__ == "__main__":
    main()