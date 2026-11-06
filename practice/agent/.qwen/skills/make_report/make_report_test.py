#!/usr/bin/env python3
"""Тесты make_report (python -m unittest make_report_test -v)."""
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path

import openpyxl

import make_report

TODAY = date(2026, 11, 6)


def task(id_: int, title: str, done: bool, priority: str, due: str | None,
         created_at: str | None = None) -> dict:
    return {"id": id_, "title": title, "done": done, "priority": priority,
            "due": due,
            "created_at": created_at or "2026-11-06T10:00:00"}


def due(days: int) -> str:
    return (TODAY + timedelta(days=days)).isoformat()


class TestBuildReport(unittest.TestCase):
    def test_kpi(self):
        tasks = [
            task(1, "Выполнена с просроченным дедлайном", True, "high", due(-10)),
            task(2, "Не выполнена, просрочена", False, "high", due(-10)),
            task(3, "Не выполнена, скоро", False, "normal", due(3)),
            task(4, "Не выполнена, в срок", False, "low", due(20)),
            task(5, "Не выполнена, без дедлайна", False, "normal", None),
            task(6, "Выполнена, без дедлайна", True, "low", None),
        ]
        kpi = make_report.build_report(tasks, today=TODAY)["kpi"]
        self.assertEqual(kpi["Всего задач"], 6)
        self.assertEqual(kpi["Выполнено"], 2)
        self.assertEqual(kpi["Не выполнено"], 4)
        self.assertAlmostEqual(kpi["% выполнения"], 33.3)
        self.assertEqual(kpi["Просрочено"], 1)            # только НЕ выполненные
        self.assertEqual(kpi["Скоро (≤ 7 дней)"], 1)
        self.assertEqual(kpi["Без дедлайна"], 1)
        self.assertEqual(kpi["Высокий приоритет, не выполнено"], 1)

    def test_priority_breakdown(self):
        tasks = [
            task(1, "A", True, "high", None), task(2, "B", False, "high", None),
            task(3, "C", True, "low", None), task(4, "D", False, "low", None),
            task(5, "E", True, "normal", None),
        ]
        report = make_report.build_report(tasks, today=TODAY)
        self.assertEqual(report["prio_rows"], [
            ("Высокий", 2, 1, 1, 50.0),
            ("Обычный", 1, 1, 0, 100.0),
            ("Низкий", 2, 1, 1, 50.0),
        ])
        self.assertEqual(report["prio_totals"], {"total": 5, "done": 3,
                                                 "pend": 2, "pct": 60.0})

    def test_rows_sorted(self):
        tasks = [
            task(4, "D", False, "low", due(10)),
            task(1, "A", True, "high", due(-5)),
            task(3, "C", False, "normal", None),
            task(2, "B", False, "high", due(-5)),
        ]
        rows = make_report.build_report(tasks, today=TODAY)["rows"]
        self.assertEqual([r[0] for r in rows], [4, 3, 2, 1])
        self.assertFalse(any(r[2] == "Выполнена" for r in rows[:3]))

    def test_rows_cell_values(self):
        rows = make_report.build_report(
            [task(1, "Задача", False, "high", due(3))], today=TODAY)["rows"]
        self.assertEqual(rows[0], [1, "Задача", "Не выполнена", "Высокий", 3,
                                   date(2026, 11, 9), 3, "Скоро",
                                   date(2026, 11, 6), 0])

    def test_rows_no_due_is_empty_cell(self):
        row = make_report.build_report(
            [task(1, "Без дедлайна", False, "low", None)], today=TODAY)["rows"][0]
        self.assertIsNone(row[5])
        self.assertIsNone(row[6])
        self.assertEqual(row[7], "Нет дедлайна")

    def test_rows_done_deadline_state(self):
        row = make_report.build_report(
            [task(1, "X", True, "high", due(-5))], today=TODAY)["rows"][0]
        self.assertEqual(row[2], "Выполнена")
        self.assertEqual(row[7], "Выполнена")
        self.assertIsNone(row[6])

    def test_groups_cover_all(self):
        tasks = [
            task(1, "A", False, "high", due(-10)),
            task(2, "B", True, "high", due(-10)),
            task(3, "C", False, "normal", due(5)),
            task(4, "D", False, "low", due(30)),
            task(5, "E", False, "normal", None),
        ]
        report = make_report.build_report(tasks, today=TODAY)
        groups = report["groups"]
        self.assertEqual(groups["Просрочено"], {"Всего": 2, "Выполнено": 1,
                                                "Не выполнено": 1})
        self.assertEqual(groups["Скоро (≤ 7 дней)"], {"Всего": 1, "Выполнено": 0,
                                                       "Не выполнено": 1})
        self.assertEqual(groups["В срок"], {"Всего": 1, "Выполнено": 0,
                                           "Не выполнено": 1})
        self.assertEqual(groups["Нет дедлайна"], {"Всего": 1, "Выполнено": 0,
                                                   "Не выполнено": 1})
        self.assertEqual(report["group_totals"], {"total": 5, "done": 1,
                                                  "pend": 4})

    def test_file_structure_and_consistency(self):
        tasks = [
            task(1, "A", False, "high", due(-10)),
            task(2, "B", True, "high", due(-10)),
            task(3, "C", False, "normal", due(5)),
            task(4, "D", False, "low", due(30)),
            task(5, "E", False, "normal", None),
        ]
        report = make_report.build_report(tasks, today=TODAY)
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "report.xlsx"
            make_report.write_report(report, path)
            wb = openpyxl.load_workbook(path)
            self.assertEqual(wb.sheetnames, ["Сводка", "Задачи", "Дедлайны"])
            sw, tw, dw = wb["Сводка"], wb["Задачи"], wb["Дедлайны"]

            self.assertEqual(sw["B5"].value, 5)
            self.assertEqual(sw["B6"].value, 1)
            self.assertEqual(sw["B7"].value, 4)
            self.assertAlmostEqual(sw["B8"].value, 20.0)
            self.assertEqual(sw["B9"].value, 1)
            self.assertEqual(sw["B10"].value, 1)
            self.assertEqual(sw["B11"].value, 1)
            self.assertEqual(sw["B12"].value, 1)
            self.assertEqual(sw["B18"].value, 5)
            self.assertEqual(sw["C18"].value, 1)
            self.assertEqual(dw["B6"].value, 5)
            self.assertEqual(dw["C6"].value, 1)
            self.assertEqual(dw["D6"].value, 4)

            states = {r[0]: r[7] for r in tw.iter_rows(min_row=2, values_only=True)}
            self.assertEqual(states[4], "В срок")
            self.assertEqual(states[3], "Скоро")
            self.assertEqual(states[5], "Нет дедлайна")
            self.assertEqual(states[1], "Просрочена")
            self.assertEqual(states[2], "Выполнена")
            self.assertEqual(
                [c.value for c in tw[1]],
                ["ID", "Задача", "Статус", "Приоритет", "Ранг", "Дедлайн",
                 "Осталось дней", "Статус дедлайна", "Создана", "Возраст, дн."])

            self.assertEqual(tw["F2"].value.date(), date(2026, 12, 6))
            self.assertIsNone(tw["F4"].value)            # задача 5 без дедлайна — пустая ячейка
            self.assertEqual(tw["F6"].value.date(), date(2026, 10, 27))
            self.assertEqual(tw["F2"].number_format, "dd.mm.yyyy")


if __name__ == "__main__":
    unittest.main()