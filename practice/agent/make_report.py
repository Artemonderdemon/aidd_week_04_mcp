#!/usr/bin/env python3
"""Создать улучшенный отчёт по задачам в Excel с несколькими листами анализа."""
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
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
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


def analyze(tasks: list[dict[str, Any]], as_of: date) -> dict[str, Any]:
    completed = sum(bool(task["done"]) for task in tasks)
    open_tasks = len(tasks) - completed
    statuses = Counter(deadline_status(task, as_of) for task in tasks)

    priorities = {}
    for priority in PRIORITIES:
        priorities[priority] = {
            "total": sum(1 for t in tasks if t.get("priority") == priority),
            "completed": sum(1 for t in tasks if t.get("priority") == priority and t["done"]),
            "open": sum(1 for t in tasks if t.get("priority") == priority and not t["done"]),
        }

    overdue_tasks = [
        t for t in tasks if not t["done"] and t.get("due") and date.fromisoformat(t["due"]) < as_of
    ]
    due_today_tasks = [
        t for t in tasks
        if not t["done"] and t.get("due") and date.fromisoformat(t["due"]) == as_of
    ]
    no_due_tasks = [t for t in tasks if not t["done"] and not t.get("due")]

    return {
        "total": len(tasks),
        "completed": completed,
        "open": open_tasks,
        "completion_rate": completed / len(tasks) if tasks else 0,
        "overdue_open": statuses["Просрочена"],
        "due_today_open": statuses["Срок сегодня"],
        "without_due_open": statuses["Без срока"],
        "priorities": priorities,
        "overdue_tasks": overdue_tasks,
        "due_today_tasks": due_today_tasks,
        "no_due_tasks": no_due_tasks,
        "completion_by_priority": {
            priority: round(
                priorities[priority]["completed"] / priorities[priority]["total"] * 100, 1
            ) if priorities[priority]["total"] else 0
            for priority in PRIORITIES
        },
    }


def _style_header(sheet, row: int = 1, start_col: int = 1, end_col: int = None):
    """Применить стиль заголовка к строке."""
    fill = PatternFill("solid", fgColor="1F4E78")
    thin = Side(style="thin", color="FFFFFF")
    for cell in sheet[row][start_col - 1:end_col] if end_col else sheet[row]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", horizontal="center")
        cell.border = Border(bottom=thin)


def _add_table(sheet, name: str):
    if sheet.max_row < 2:
        return
    # Определяем начало таблицы: если строка 1 содержит None, это заголовок-название
    start_row = 1
    for col in range(2, sheet.max_column + 1):
        if sheet.cell(row=1, column=col).value is None:
            start_row = 2
            break
    # Находим последнюю строку с непустой первой колонкой
    max_row = 0
    for row in range(start_row, sheet.max_row + 1):
        val = sheet.cell(row=row, column=1).value
        if val not in (None, ""):
            max_row = row
    if max_row < start_row + 1:
        return
    table = Table(
        displayName=name,
        ref=f"A{start_row}:{sheet.cell(row=max_row, column=sheet.max_column).coordinate}"
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def _format_summary_sheet(summary, stats, as_of):
    """Заполнить лист Сводка."""
    # Заголовок (без merge — просто в A1, остальные пустые строки не создаём)
    summary["A1"] = "ОТЧЁТ ПО ЗАДАЧАМ"
    summary["A1"].font = Font(bold=True, size=14, color="1F4E78")

    # Основные показатели
    summary.append([])
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

    # Форматирование процентов
    summary["B6"].number_format = "0.0%"

    # Подсветка ключевых метрик
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    yellow_fill = PatternFill("solid", fgColor="FFEB9C")
    green_fill = PatternFill("solid", fgColor="C6EFCE")

    # Просроченные - жёлтые если > 0
    if stats["overdue_open"] > 0:
        for cell in summary["A7:B7"]:
            cell[0].fill = yellow_fill
            cell[1].fill = yellow_fill

    # Доля выполненных - цветовая шкала
    rate_cell = summary["B5"]
    if stats["completion_rate"] < 0.3:
        rate_cell.fill = red_fill
    elif stats["completion_rate"] < 0.5:
        rate_cell.fill = yellow_fill
    else:
        rate_cell.fill = green_fill

    summary.append([])
    summary.append(["Приоритет", "Всего", "Выполнено", "Не выполнено", "Доля выполненных"])
    for priority in PRIORITIES:
        values = stats["priorities"][priority]
        rate = stats["completion_by_priority"][priority]
        summary.append([
            PRIORITY_LABELS[priority], values["total"], values["completed"],
            values["open"], f"{rate}%"
        ])

    # Форматирование строк приоритетов
    for i, priority in enumerate(PRIORITIES):
        row_num = 12 + i
        rate_cell = summary.cell(row=row_num, column=5)
        rate_val = stats["completion_by_priority"][priority]
        if rate_val < 30:
            rate_cell.fill = red_fill
        elif rate_val < 50:
            rate_cell.fill = yellow_fill
        else:
            rate_cell.fill = green_fill

    _style_header(summary, 10)
    summary.freeze_panes = "A2"
    summary.column_dimensions["A"].width = 31
    summary.column_dimensions["B"].width = 18
    summary.column_dimensions["C"].width = 14
    summary.column_dimensions["D"].width = 16
    summary.column_dimensions["E"].width = 20

    # Заголовок таблицы приоритетов
    priority_header_cell = summary.cell(row=10, column=1)
    priority_header_cell.font = Font(bold=True, color="FFFFFF")
    priority_header_cell.fill = PatternFill("solid", fgColor="2E75B6")


def _format_tasks_sheet(details, tasks, stats, as_of):
    """Заполнить лист Задачи с дополнительными колонками."""
    headers = [
        "ID", "Название", "Статус", "Приоритет", "Срок",
        "Статус срока", "Создана",
    ]
    details.append(headers)

    # Сортировка: сначала просроченные, потом по сроку, потом по ID
    def sort_key(task):
        status = deadline_status(task, as_of)
        status_order = {"Просрочена": 0, "Срок сегодня": 1, "В сроке": 2, "Без срока": 3, "Выполнена": 4}
        return (status_order.get(status, 99), task.get("due") or "9999-99-99", task["id"])

    sorted_tasks = sorted(tasks, key=sort_key)

    for task in sorted_tasks:
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

    # Условное форматирование
    # Красная подсветка просроченных
    red_font = Font(color="9C0006", bold=True)
    red_fill = PatternFill("solid", fgColor="FFC7CE")

    for row_idx, task in enumerate(sorted_tasks, start=2):
        status_col = details.cell(row=row_idx, column=6)
        done_col = details.cell(row=row_idx, column=3)
        prio_col = details.cell(row=row_idx, column=4)

        status_text = deadline_status(task, as_of)
        if status_text == "Просрочена":
            for col in range(1, 8):
                details.cell(row=row_idx, column=col).fill = red_fill
            details.cell(row=row_idx, column=3).font = red_font
        elif status_text == "Срок сегодня":
            yellow_fill = PatternFill("solid", fgColor="FFEB9C")
            for col in range(1, 8):
                details.cell(row=row_idx, column=col).fill = yellow_fill
        elif task["done"]:
            green_fill = PatternFill("solid", fgColor="C6EFCE")
            for col in range(1, 8):
                details.cell(row=row_idx, column=col).fill = green_fill

        # Приоритет - цветные метки
        prio_text = PRIORITY_LABELS.get(task["priority"], task["priority"])
        if prio_text == "Высокий":
            prio_col.font = Font(bold=True, color="9C0006")

    # Ширина колонок
    for column, width in {"A": 10, "B": 48, "C": 18, "D": 16, "E": 14, "F": 22, "G": 22}.items():
        details.column_dimensions[column].width = width

    _add_table(details, "TodoTasks")


def _format_overdue_sheet(overdue_sheet, overdue_tasks):
    """Создать лист с просроченными задачами."""
    overdue_sheet["A1"] = "ПРОСРОЧЕННЫЕ ЗАДАЧИ"
    overdue_sheet["A1"].font = Font(bold=True, size=14, color="9C0006")
    overdue_sheet["B1"] = ""
    overdue_sheet.append(["ID", "Название", "Приоритет", "Срок", "Создана", "Дней просрочки"])

    from datetime import date
    today = date.today()

    for task in sorted(overdue_tasks, key=lambda x: x.get("due") or ""):
        due_date = date.fromisoformat(task["due"])
        days_overdue = (today - due_date).days
        overdue_sheet.append([
            task["id"],
            task["title"],
            PRIORITY_LABELS.get(task["priority"], task["priority"]),
            task["due"],
            task.get("created_at", ""),
            days_overdue,
        ])

    _style_header(overdue_sheet, 3)
    overdue_sheet.freeze_panes = "A4"

    # Красная подсветка
    red_font = Font(color="9C0006", bold=True)
    red_fill = PatternFill("solid", fgColor="FFC7CE")

    for row_idx in range(4, overdue_sheet.max_row + 1):
        for col in range(1, 7):
            cell = overdue_sheet.cell(row=row_idx, column=col)
            if row_idx > 3:
                cell.fill = red_fill
            if col == 1 or col == 3:
                cell.font = red_font

    overdue_sheet.column_dimensions["A"].width = 10
    overdue_sheet.column_dimensions["B"].width = 48
    overdue_sheet.column_dimensions["C"].width = 16
    overdue_sheet.column_dimensions["D"].width = 14
    overdue_sheet.column_dimensions["E"].width = 22
    overdue_sheet.column_dimensions["F"].width = 16

    _add_table(overdue_sheet, "OverdueTasks")


def _format_due_today_sheet(due_sheet, due_tasks):
    """Создать лист с задачами со сроком сегодня."""
    due_sheet["A1"] = "ЗАДАЧИ СО СРОКОМ СЕГОДНЯ"
    due_sheet["A1"].font = Font(bold=True, size=14, color="9C6500")
    due_sheet["B1"] = ""
    due_sheet.append(["ID", "Название", "Приоритет", "Срок", "Создана"])

    for task in sorted(due_tasks, key=lambda x: x["id"]):
        due_sheet.append([
            task["id"],
            task["title"],
            PRIORITY_LABELS.get(task["priority"], task["priority"]),
            task["due"],
            task.get("created_at", ""),
        ])

    _style_header(due_sheet, 3)
    due_sheet.freeze_panes = "A4"

    yellow_fill = PatternFill("solid", fgColor="FFEB9C")
    for row_idx in range(4, due_sheet.max_row + 1):
        for col in range(1, 6):
            due_sheet.cell(row=row_idx, column=col).fill = yellow_fill

    due_sheet.column_dimensions["A"].width = 10
    due_sheet.column_dimensions["B"].width = 48
    due_sheet.column_dimensions["C"].width = 16
    due_sheet.column_dimensions["D"].width = 14
    due_sheet.column_dimensions["E"].width = 22

    _add_table(due_sheet, "DueTodayTasks")


def _format_no_due_sheet(no_due_sheet, no_due_tasks):
    """Создать лист с задачами без срока."""
    no_due_sheet["A1"] = "ЗАДАЧИ БЕЗ СРОКА"
    no_due_sheet["A1"].font = Font(bold=True, size=14, color="000000")
    no_due_sheet["B1"] = ""
    no_due_sheet.append(["ID", "Название", "Приоритет", "Создана"])

    for task in sorted(no_due_tasks, key=lambda x: x["id"]):
        no_due_sheet.append([
            task["id"],
            task["title"],
            PRIORITY_LABELS.get(task["priority"], task["priority"]),
            task.get("created_at", ""),
        ])

    _style_header(no_due_sheet, 3)
    no_due_sheet.freeze_panes = "A4"

    gray_fill = PatternFill("solid", fgColor="D9D9D9")
    for row_idx in range(4, no_due_sheet.max_row + 1):
        for col in range(1, 5):
            no_due_sheet.cell(row=row_idx, column=col).fill = gray_fill

    no_due_sheet.column_dimensions["A"].width = 10
    no_due_sheet.column_dimensions["B"].width = 48
    no_due_sheet.column_dimensions["C"].width = 16
    no_due_sheet.column_dimensions["D"].width = 22

    _add_table(no_due_sheet, "NoDueTasks")


def _format_methodology_sheet(methodology):
    """Заполнить лист Методика."""
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
        ("Порядок сортировки задач", "По статусу срока (просроченные первыми), затем по дате срока, затем по ID."),
        ("Дней просрочки", "Разница между датой отчёта и сроком выполнения в днях."),
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


def create_report(
    tasks: list[dict[str, Any]], output: str | Path, as_of: date | None = None
) -> Path:
    as_of = as_of or date.today()
    stats = analyze(tasks, as_of)
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()

    # Лист 1: Сводка
    summary = workbook.active
    summary.title = "Сводка"
    _format_summary_sheet(summary, stats, as_of)

    # Лист 2: Задачи (все задачи)
    details = workbook.create_sheet("Задачи")
    _format_tasks_sheet(details, tasks, stats, as_of)

    # Лист 3: Просроченные
    if stats["overdue_tasks"]:
        overdue_sheet = workbook.create_sheet("Просроченные")
        _format_overdue_sheet(overdue_sheet, stats["overdue_tasks"])

    # Лист 4: Срок сегодня
    if stats["due_today_tasks"]:
        due_sheet = workbook.create_sheet("Срок сегодня")
        _format_due_today_sheet(due_sheet, stats["due_today_tasks"])

    # Лист 5: Без срока
    if stats["no_due_tasks"]:
        no_due_sheet = workbook.create_sheet("Без срока")
        _format_no_due_sheet(no_due_sheet, stats["no_due_tasks"])

    # Лист 6: Методика
    methodology = workbook.create_sheet("Методика")
    _format_methodology_sheet(methodology)

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