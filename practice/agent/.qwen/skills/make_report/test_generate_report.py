from datetime import date, timedelta
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from generate_report import create_report, deadline_status, summarize


class ReportTest(unittest.TestCase):
    def setUp(self):
        self.today = date(2026, 11, 7)
        self.tasks = [
            {"id": 1, "title": "Просроченная", "done": False, "priority": "high",
             "due": "2026-11-06", "created_at": "2026-10-01"},
            {"id": 2, "title": "На сегодня", "done": False, "priority": "normal",
             "due": "2026-11-07", "created_at": "2026-10-02"},
            {"id": 3, "title": "Выполненная", "done": True, "priority": "high",
             "due": "2026-11-01", "created_at": "2026-10-03"},
            {"id": 4, "title": "Без срока", "done": False, "priority": "low",
             "due": None, "created_at": "2026-10-04"},
            {"id": 5, "title": "Будущая", "done": False, "priority": "normal",
             "due": "2026-11-08", "created_at": "2026-10-05"},
        ]

    def test_summary_counts_statuses_and_priority(self):
        stats = summarize(self.tasks, self.today)
        self.assertEqual(stats["total"], 5)
        self.assertEqual(stats["completed"], 1)
        self.assertEqual(stats["open"], 4)
        self.assertEqual(stats["completion_rate"], 0.2)
        self.assertEqual(stats["overdue_open"], 1)
        self.assertEqual(stats["due_today_open"], 1)
        self.assertEqual(stats["without_due_open"], 1)
        self.assertEqual(stats["priorities"]["high"], {"total": 2, "completed": 1, "open": 1})

    def test_deadline_status_treats_completed_task_as_completed(self):
        self.assertEqual(deadline_status(self.tasks[2], self.today), "Выполнена")

    def test_empty_report_has_zero_completion_rate(self):
        self.assertEqual(summarize([], self.today)["completion_rate"], 0)

    def test_report_contains_required_sheets_and_metrics(self):
        with tempfile.TemporaryDirectory() as directory:
            path = create_report(self.tasks, Path(directory) / "report.xlsx", self.today)
            workbook = load_workbook(path, data_only=True)
            self.assertEqual(workbook.sheetnames, ["Сводка", "Задачи", "Методика"])
            summary = workbook["Сводка"]
            self.assertEqual(summary["B3"].value, 5)
            self.assertEqual(summary["B4"].value, 1)
            self.assertEqual(summary["B5"].value, 4)
            self.assertEqual(summary["B6"].value, 0.2)
            self.assertEqual(workbook["Задачи"].max_row, 6)
            self.assertEqual(workbook["Задачи"]["F2"].value, "Просрочена")
            self.assertEqual(workbook["Методика"].max_row, 10)

    def test_deadline_boundary_includes_today_and_excludes_tomorrow(self):
        self.assertEqual(deadline_status(self.tasks[1], self.today), "Срок сегодня")
        self.assertEqual(deadline_status(self.tasks[4], self.today), "В сроке")
        self.assertEqual(self.today + timedelta(days=1), date.fromisoformat(self.tasks[4]["due"]))


if __name__ == "__main__":
    unittest.main()
