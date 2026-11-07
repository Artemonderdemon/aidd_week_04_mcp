#!/usr/bin/env python3
"""Small independent baseline report generator, intentionally not a Qwen run."""
import argparse
import json
import sys
import urllib.error
import urllib.request
from datetime import date
from pathlib import Path

from openpyxl import Workbook


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--api-url", default="http://127.0.0.1:5000")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "report_without_skill.xlsx",
    )
    args = parser.parse_args()
    url = f"{args.api_url.rstrip('/')}/api/tasks"
    try:
        with urllib.request.urlopen(url, timeout=10) as response:
            tasks = json.load(response)
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as error:
        print(f"Не удалось получить задачи: {error}", file=sys.stderr)
        return 1
    if not isinstance(tasks, list):
        print("TODO API должен вернуть список задач", file=sys.stderr)
        return 1

    today = date.today()
    completed = sum(bool(task["done"]) for task in tasks)
    open_tasks = [task for task in tasks if not task["done"]]
    overdue = [task for task in open_tasks if task["due"] and task["due"] < today.isoformat()]

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.append(["Metric", "Value"])
    summary.append(["Report date", today.isoformat()])
    summary.append(["Total tasks", len(tasks)])
    summary.append(["Completed", completed])
    summary.append(["Open", len(open_tasks)])
    summary.append(["Completion rate", completed / len(tasks) if tasks else 0])
    summary.append(["Overdue open", len(overdue)])
    summary.append([])
    summary.append(["Priority", "Total", "Completed", "Open"])
    for priority in ("high", "normal", "low"):
        matching = [task for task in tasks if task["priority"] == priority]
        summary.append([
            priority,
            len(matching),
            sum(bool(task["done"]) for task in matching),
            sum(not task["done"] for task in matching),
        ])

    details = workbook.create_sheet("Tasks")
    details.append(["ID", "Title", "Done", "Priority", "Due", "Deadline status"])
    for task in sorted(tasks, key=lambda item: item["id"]):
        if task["done"]:
            deadline = "Completed"
        elif not task["due"]:
            deadline = "No deadline"
        elif task["due"] < today.isoformat():
            deadline = "Overdue"
        elif task["due"] == today.isoformat():
            deadline = "Due today"
        else:
            deadline = "Upcoming"
        details.append([
            task["id"], task["title"], bool(task["done"]), task["priority"],
            task["due"] or "", deadline,
        ])
    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output)
    print(f"Baseline report saved: {args.output} ({len(tasks)} tasks)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
