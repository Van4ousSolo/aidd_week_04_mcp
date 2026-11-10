from pathlib import Path
from datetime import date
from openpyxl import load_workbook

PATH = Path(__file__).resolve().parent / "report.xlsx"
wb = load_workbook(PATH)

assert wb.sheetnames == ["Задачи", "Сводка"], wb.sheetnames

ws = wb["Задачи"]
assert [c.value for c in ws[1]] == ["id", "название", "priority", "due", "done"], [c.value for c in ws[1]]

rows = list(ws.iter_rows(min_row=2, values_only=True))
tasks = [
    {"id": r[0], "title": r[1], "priority": r[2], "due": r[3], "done": bool(r[4])}
    for r in rows
]
assert len(tasks) == 464, len(tasks)
assert len({t["id"] for t in tasks}) == 464

today = date(2026, 11, 10)
expected = {
    "Всего": len(tasks),
    "Выполнено": sum(1 for t in tasks if t["done"]),
    "Не выполнено": sum(1 for t in tasks if not t["done"]),
    "Просрочено": sum(
        1 for t in tasks
        if not t["done"]
        and t["due"] is not None
        and date.fromisoformat(str(t["due"])) < today
    ),
    "Low": 0,
    "Normal": 0,
    "High": 0,
}
for t in tasks:
    expected[t["priority"].capitalize()] += 1

ws2 = wb["Сводка"]
pairs = {r[0]: r[1] for r in ws2.iter_rows(min_row=2, values_only=True)}
for label, value in expected.items():
    pairs[label] = float(pairs[label])

ok = True
for label, expected_value in expected.items():
    got = pairs[label]
    match = abs(float(got) - float(expected_value)) < 1e-9
    ok &= match
    print(f"{label}: сводка={got} пересчёт={expected_value} {'OK' if match else 'MISMATCH'}")

print("процент:", pairs["Процент выполненных, %"], "ожидается", round(expected["Выполнено"] / expected["Всего"] * 100, 2))
print("первые 3 задачи:", tasks[:3])
print("последние 3 задачи:", tasks[-3:])
print("ВСЁ OK" if ok else "ОШИБКА РАСЧЁТА")