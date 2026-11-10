import json
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook

MCP_OUTPUT = Path(sys.argv[1]) if len(sys.argv) > 1 else None
OUTPUT_PATH = Path(__file__).resolve().parent / "report.xlsx"
TODAY = date(2026, 11, 10)


def parse_glued_json(text: str):
    """Разобрать файл, в котором JSON-объекты могут идти подряд (NDJSON / липкий JSON)."""
    decoder = json.JSONDecoder()
    text = text.strip()
    if text.startswith("---"):
        text = text[text.index("\n") + 1:].strip()
    objs, i, n = [], 0, len(text)
    while i < n:
        while i < n and text[i] in " \t\n\r":
            i += 1
        if i >= n:
            break
        obj, end = decoder.raw_decode(text, i)
        objs.append(obj)
        i = end
    return objs


def extract_task(obj: dict):
    due = obj.get("due")
    return {
        "id": int(obj["id"]),
        "title": str(obj["title"]),
        "priority": str(obj.get("priority") or "").strip().lower(),
        "due": due.strftime("%Y-%m-%d") if isinstance(due, date) else (str(due) if due else None),
        "done": bool(obj["done"]),
    }


def main():
    if MCP_OUTPUT is None:
        raise SystemExit("usage: build_report.py <mcp list_tasks output>")

    tasks = sorted((extract_task(r) for r in parse_glued_json(MCP_OUTPUT.read_text(encoding="utf-8"))),
                   key=lambda t: t["id"])

    unknown_priority = sorted({t["priority"] for t in tasks}
                              - {"low", "normal", "high"})
    if unknown_priority:
        print("warning: незнакомые priority:", unknown_priority)

    total = len(tasks)
    done_count = sum(1 for t in tasks if t["done"])
    not_done_count = total - done_count
    percent = round(done_count / total * 100, 2) if total else 0.0
    overdue = sum(
        1 for t in tasks
        if not t["done"] and t["due"] is not None and date.fromisoformat(t["due"]) < TODAY
    )
    counts = {"low": 0, "normal": 0, "high": 0}
    for t in tasks:
        if t["priority"] in counts:
            counts[t["priority"]] += 1

    wb = Workbook()

    ws = wb.active
    ws.title = "Задачи"
    ws.append(["id", "название", "priority", "due", "done"])
    for t in tasks:
        ws.append([t["id"], t["title"], t["priority"], t["due"] or "", t["done"]])
    ws.freeze_panes = "A2"
    for col, width in {"A": 8, "B": 48, "C": 10, "D": 12, "E": 8}.items():
        ws.column_dimensions[col].width = width

    ws2 = wb.create_sheet("Сводка")
    ws2.append(["Показатель", "Значение"])
    for label, value in [
        ("Всего", total),
        ("Выполнено", done_count),
        ("Не выполнено", not_done_count),
        ("Процент выполненных, %", percent),
        ("Просрочено", overdue),
        ("Low", counts["low"]),
        ("Normal", counts["normal"]),
        ("High", counts["high"]),
    ]:
        ws2.append([label, value])
    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 14

    wb.save(OUTPUT_PATH)
    print(f"saved: {OUTPUT_PATH}")
    print(f"tasks={total} done={done_count} not_done={not_done_count} "
          f"percent={percent}% overdue={overdue} low={counts['low']} "
          f"normal={counts['normal']} high={counts['high']}")


if __name__ == "__main__":
    main()