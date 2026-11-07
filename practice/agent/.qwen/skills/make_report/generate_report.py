#!/usr/bin/env python3
"""Fetch TODO tasks and create a deterministic multi-sheet Excel report."""
from __future__ import annotations

import argparse
import json
import sys
import urllib.error
import urllib.request
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

PRIORITIES = ("high", "normal", "low")
PRIORITY_LABELS = {"high": "Высокий", "normal": "Обычный", "low": "Низкий"}


def fetch_tasks(api_url: str, timeout: float = 10) -> list[dict[str, Any]]:
    url = f"{api_url.rstrip('/')}/api/tasks"
    request = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            tasks = json.load(response)
    except urllib.error.HTTPError as error:
        raise RuntimeError(f"TODO API вернул HTTP {error.code} для {url}") from error
    except (urllib.error.URLError, TimeoutError) as error:
        reason = getattr(error, "reason", error)
        raise RuntimeError(f"Не удалось получить задачи из {url}: {reason}") from error
    except json.JSONDecodeError as error:
        raise RuntimeError(f"TODO API вернул некорректный JSON: {error}") from error
    if not isinstance(tasks, list) or any(not isinstance(task, dict) for task in tasks):
        raise RuntimeError("Ответ TODO API должен быть JSON-списком объектов задач")
    return tasks


def deadline_status(task: dict[str, Any], as_of: date) -> str:
    if task["done"]:
        return "Выполнена"
    if not task["due"]:
        return "Без срока"
    due = date.fromisoformat(task["due"])
    if due < as_of:
        return "Просрочена"
    if due == as_of:
        return "Срок сегодня"
    return "В сроке"


def summarize(tasks: list[dict[str, Any]], as_of: date) -> dict[str, Any]:
    completed = sum(bool(task["done"]) for task in tasks)
    open_tasks = len(tasks) - completed
    statuses = Counter(deadline_status(task, as_of) for task in tasks)
    priorities = {
        priority: {
            "total": sum(task["priority"] == priority for task in tasks),
            "completed": sum(task["priority"] == priority and task["done"] for task in tasks),
            "open": sum(task["priority"] == priority and not task["done"] for task in tasks),
        }
        for priority in PRIORITIES
    }
    return {
        "total": len(tasks),
        "completed": completed,
        "open": open_tasks,
        "completion_rate": completed / len(tasks) if tasks else 0,
        "overdue_open": statuses["Просрочена"],
        "due_today_open": statuses["Срок сегодня"],
        "without_due_open": sum(
            not task["done"] and not task["due"] for task in tasks
        ),
        "priorities": priorities,
    }


def _style_header(sheet, row: int = 1) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[row]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center")


def _add_table(sheet, name: str) -> None:
    if sheet.max_row < 2:
        return
    table = Table(displayName=name, ref=sheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def create_report(
    tasks: list[dict[str, Any]], output: str | Path, as_of: date | None = None
) -> Path:
    as_of = as_of or date.today()
    stats = summarize(tasks, as_of)
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Сводка"
    summary.append(["Показатель", "Значение"])
    summary_rows = [
        ("Дата отчёта", as_of.isoformat()),
        ("Всего задач", stats["total"]),
        ("Выполнено", stats["completed"]),
        ("Не выполнено", stats["open"]),
        ("Доля выполненных", stats["completion_rate"]),
        ("Просрочено (открытые)", stats["overdue_open"]),
        ("Срок сегодня (открытые)", stats["due_today_open"]),
        ("Без срока (открытые)", stats["without_due_open"]),
    ]
    for row in summary_rows:
        summary.append(row)
    summary["B6"].number_format = "0.0%"
    summary.append([])
    summary.append(["Приоритет", "Всего", "Выполнено", "Не выполнено"])
    for priority in PRIORITIES:
        values = stats["priorities"][priority]
        summary.append([
            PRIORITY_LABELS[priority], values["total"], values["completed"], values["open"]
        ])
    _style_header(summary, 1)
    _style_header(summary, 11)
    summary.freeze_panes = "A2"
    summary.column_dimensions["A"].width = 31
    for column in "BCD":
        summary.column_dimensions[column].width = 18

    details = workbook.create_sheet("Задачи")
    headers = [
        "ID", "Название", "Статус", "Приоритет", "Срок",
        "Статус срока", "Создана",
    ]
    details.append(headers)
    for task in sorted(tasks, key=lambda item: item["id"]):
        details.append([
            task["id"],
            task["title"],
            "Выполнена" if task["done"] else "Не выполнена",
            PRIORITY_LABELS.get(task["priority"], task["priority"]),
            task["due"] or "",
            deadline_status(task, as_of),
            task.get("created_at", ""),
        ])
    _style_header(details)
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions
    for column, width in {"A": 10, "B": 48, "C": 18, "D": 16, "E": 14, "F": 22, "G": 22}.items():
        details.column_dimensions[column].width = width
    _add_table(details, "TodoTasks")

    methodology = workbook.create_sheet("Методика")
    methodology.append(["Показатель", "Правило расчёта"])
    methodology_rows = [
        ("Всего задач", "Количество записей, полученных из TODO API."),
        ("Выполнено / Не выполнено", "Поле done равно true / false."),
        ("Доля выполненных", "Выполнено / всего; для пустого списка равно 0."),
        ("Просрочено", "Не выполнена и due раньше даты отчёта."),
        ("Срок сегодня", "Не выполнена и due равен дате отчёта."),
        ("Без срока", "Не выполнена и due отсутствует."),
        ("Статус срока", "Выполненная; без срока; просрочена; срок сегодня; иначе в сроке."),
        ("Сводка по приоритету", "Количество задач каждого приоритета, отдельно выполненные и открытые."),
        ("Дата отчёта", "Локальная календарная дата запуска скрипта; можно зафиксировать в тесте."),
    ]
    for row in methodology_rows:
        methodology.append(row)
    _style_header(methodology)
    methodology.freeze_panes = "A2"
    methodology.column_dimensions["A"].width = 29
    methodology.column_dimensions["B"].width = 92
    for row in methodology.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    _add_table(methodology, "ReportMethodology")

    workbook.save(output_path)
    return output_path


def default_output_path() -> Path:
    agent_dir = Path(__file__).resolve().parents[3]
    return agent_dir / "tasks_report.xlsx"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--api-url", default="http://127.0.0.1:5000")
    parser.add_argument("--output", type=Path, default=default_output_path())
    args = parser.parse_args(argv)
    try:
        tasks = fetch_tasks(args.api_url)
        output = create_report(tasks, args.output)
    except (RuntimeError, KeyError, TypeError, ValueError, OSError) as error:
        print(f"Ошибка создания отчёта: {error}", file=sys.stderr)
        return 1
    print(f"Отчёт сохранён: {output} ({len(tasks)} задач)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
