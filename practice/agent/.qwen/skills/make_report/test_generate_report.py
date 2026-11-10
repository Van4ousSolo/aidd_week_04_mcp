#!/usr/bin/env python3
"""Модульные тесты для проверки генератора аналитических отчётов make_report."""
import importlib.util
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).parent
spec = importlib.util.spec_from_file_location("generate_report", HERE / "generate_report.py")
report_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(report_module)


class TestMakeReport(unittest.TestCase):
    def setUp(self):
        self.sample_tasks = [
            {"id": 1, "title": "Задача 1: завершена", "priority": "high", "due": "2026-11-01", "done": True},
            {"id": 2, "title": "=CMD()!A1", "priority": "high", "due": "2026-11-05", "done": False},  # Просрочена, формула
            {"id": 3, "title": "Задача 3: в работе", "priority": "normal", "due": "2026-11-12", "done": False},  # В течение 7 дней
            {"id": 4, "title": "Задача 4: без срока", "priority": "low", "due": None, "done": False},
            {"id": 5, "title": "Задача 5: дальний срок", "priority": "low", "due": "2026-12-01", "done": True},
        ]
        self.as_of = date(2026, 11, 10)

    def test_metrics_and_structure(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_file = Path(tmp_dir) / "test_report.xlsx"
            result = report_module.create_report(self.sample_tasks, self.as_of, out_file)

            # Проверка возвращаемого словаря
            self.assertEqual(result["total"], 5)
            self.assertEqual(result["completed"], 2)
            self.assertEqual(result["in_progress"], 3)
            self.assertEqual(result["completion_rate"], 40.0)
            self.assertEqual(result["overdue"], 1)
            self.assertAlmostEqual(result["overdue_active_rate"], 33.33, places=2)

            # Проверка созданного файла Excel
            wb = load_workbook(out_file, data_only=False)
            self.assertEqual(wb.sheetnames, ["Реестр задач", "Сводная аналитика"])

            ws_tasks = wb["Реестр задач"]
            # Заголовки
            headers = [cell.value for cell in ws_tasks[1]]
            self.assertIn("ID", headers)
            self.assertIn("Название задачи", headers)
            self.assertIn("Просрочена", headers)

            # Проверка защиты от формульных инъекций
            row2_title_cell = ws_tasks["B3"]  # строка для id=2
            self.assertEqual(row2_title_cell.data_type, "s")
            self.assertEqual(row2_title_cell.value, "=CMD()!A1")

            # Проверка содержимого листа сводки
            ws_summary = wb["Сводная аналитика"]
            summary_values = {row[0]: row[1] for row in ws_summary.iter_rows(values_only=True) if row[0]}
            self.assertEqual(summary_values.get("Всего задач в системе"), 5)
            self.assertEqual(summary_values.get("Выполнено задач (Done)"), 2)
            self.assertEqual(summary_values.get("В работе (Active / Pending)"), 3)
            self.assertEqual(summary_values.get("Процент выполнения (%)"), "40.0%")
            self.assertEqual(summary_values.get("Просрочено невыполненных задач"), 1)

    def test_protection_against_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_file = Path(tmp_dir) / "locked.xlsx"
            report_module.create_report(self.sample_tasks, self.as_of, out_file)
            self.assertTrue(out_file.exists())

            # Повторное создание по тому же пути должно выбросить FileExistsError
            with self.assertRaises(FileExistsError):
                report_module.create_report(self.sample_tasks, self.as_of, out_file)

    def test_empty_tasks(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_file = Path(tmp_dir) / "empty.xlsx"
            result = report_module.create_report([], self.as_of, out_file)
            self.assertEqual(result["total"], 0)
            self.assertEqual(result["completed"], 0)
            self.assertEqual(result["completion_rate"], 0.0)
            self.assertEqual(result["overdue"], 0)

            wb = load_workbook(out_file)
            self.assertEqual(wb.sheetnames, ["Реестр задач", "Сводная аналитика"])

    def test_validation_errors(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_dir = Path(tmp_dir)

            # Недопустимый приоритет
            bad_priority = [{"id": 1, "title": "T", "priority": "urgent", "due": None, "done": False}]
            with self.assertRaises(ValueError):
                report_module.create_report(bad_priority, self.as_of, out_dir / "err1.xlsx")

            # Дубликат ID
            dup_id = [
                {"id": 10, "title": "T1", "priority": "low", "due": None, "done": False},
                {"id": 10, "title": "T2", "priority": "high", "due": None, "done": True},
            ]
            with self.assertRaises(ValueError):
                report_module.create_report(dup_id, self.as_of, out_dir / "err2.xlsx")

            # Некорректный формат даты
            bad_date = [{"id": 1, "title": "T", "priority": "normal", "due": "10-11-2026", "done": False}]
            with self.assertRaises(ValueError):
                report_module.create_report(bad_date, self.as_of, out_dir / "err3.xlsx")


if __name__ == "__main__":
    unittest.main()
