#!/usr/bin/env python3
"""Генерация отчёта по задачам (HOMEWORK, раздел 1) в .xlsx.

Читает apps/todo_db.json, считает метрики на дату генерации и пишет
report_ДД.ММ.ГГГГ.xlsx с листами «Сводка», «Задачи», «Дедлайны»
по раскладке из report_format.txt. Значения записываются статикой —
прогоны в один день дают одинаковые файлы.
"""
from __future__ import annotations

import json
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Привязка к корню проекта: ищем каталог apps (данные задач).
# Отчёт кладём рядом с report_format.txt (practice/agent).
BASE = Path(__file__).resolve().parent
while not (BASE / "apps").is_dir() and BASE.parent != BASE:
    BASE = BASE.parent
DB_PATH = BASE / "apps" / "todo_db.json"
OUT_PATH = BASE / "practice" / "agent" / f"report_{date.today().strftime('%d.%m.%Y')}.xlsx"

PRI = {"low": "Низкий", "normal": "Обычный", "high": "Высокий"}
RANK = {"low": 1, "normal": 2, "high": 3}
GROUPS = ["Просрочено", "Скоро (≤ 7 дней)", "В срок", "Нет дедлайна"]

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE_FMT = "dd.mm.yyyy"

WIDTHS = {
    "Сводка": {"A": 34, "B": 13},
    "Задачи": {"A": 5, "B": 42, "C": 14, "D": 12, "E": 6,
               "F": 11, "G": 12, "H": 14, "I": 11, "J": 10},
    "Дедлайны": {"A": 18, "B": 9, "C": 12, "D": 14},
}


def pct(done: int, total: int) -> float:
    return round(done / total * 100, 1) if total else 0.0


def load_tasks() -> list[dict]:
    doc = json.loads(DB_PATH.read_text(encoding="utf-8"))
    return doc["payload"]["tasks"]


def due_of(t: dict) -> date | None:
    return date.fromisoformat(t["due"]) if t.get("due") else None


def build_report(tasks: list[dict], today: date | None = None) -> dict:
    """Считает все метрики отчёта (см. раскладку в report_format.txt)."""
    today = today or date.today()

    total = len(tasks)
    done_n = sum(1 for t in tasks if t["done"])
    pend = [t for t in tasks if not t["done"]]

    overdue = sum(1 for t in pend if (d := due_of(t)) and d < today)
    soon = sum(1 for t in pend
               if (d := due_of(t)) and today <= d <= today + timedelta(days=7))
    no_due = sum(1 for t in pend if t["due"] is None)
    high_pend = sum(1 for t in pend if t["priority"] == "high")

    kpi = {
        "Всего задач": total,
        "Выполнено": done_n,
        "Не выполнено": total - done_n,
        "% выполнения": pct(done_n, total),
        "Просрочено": overdue,
        "Скоро (≤ 7 дней)": soon,
        "Без дедлайна": no_due,
        "Высокий приоритет, не выполнено": high_pend,
    }

    prio_rows: list[tuple[str, int, int, int, float]] = []
    prio_totals = Counter()
    for p in ("high", "normal", "low"):
        group = [t for t in tasks if t["priority"] == p]
        dn = sum(1 for t in group if t["done"])
        prio_rows.append((PRI[p], len(group), dn, len(group) - dn, pct(dn, len(group))))
        prio_totals["total"] += len(group)
        prio_totals["done"] += dn
        prio_totals["pend"] += len(group) - dn
    prio_totals["pct"] = pct(prio_totals["done"], prio_totals["total"])

    groups = {g: {"Всего": 0, "Выполнено": 0, "Не выполнено": 0} for g in GROUPS}
    for t in tasks:
        d = due_of(t)
        if d is None:
            g = "Нет дедлайна"
        elif d < today:
            g = "Просрочено"
        elif d <= today + timedelta(days=7):
            g = "Скоро (≤ 7 дней)"
        else:
            g = "В срок"
        groups[g]["Всего"] += 1
        groups[g]["Выполнено" if t["done"] else "Не выполнено"] += 1

    group_totals = Counter()
    for g in GROUPS:
        c = groups[g]
        group_totals["total"] += c["Всего"]
        group_totals["done"] += c["Выполнено"]
        group_totals["pend"] += c["Не выполнено"]

    rows = []
    for t in tasks:
        d = due_of(t)
        if t["done"]:
            state, remain = "Выполнена", None
        else:
            remain = None if d is None else (d - today).days
            if d is None:
                state = "Нет дедлайна"
            elif d < today:
                state = "Просрочена"
            elif d <= today + timedelta(days=7):
                state = "Скоро"
            else:
                state = "В срок"
        created = date.fromisoformat(t["created_at"][:10])
        rows.append([t["id"], t["title"],
                     "Выполнена" if t["done"] else "Не выполнена",
                     PRI[t["priority"]], RANK[t["priority"]], d, remain,
                     state, created, (today - created).days])
    rows.sort(key=lambda r: (r[2] != "Не выполнена", r[4], r[5] is None,
                             r[5] or date.min, r[0]))

    return {"today": today, "kpi": kpi, "prio_rows": prio_rows,
            "prio_totals": prio_totals, "groups": groups,
            "group_totals": group_totals, "rows": rows}


def _header(cell) -> None:
    cell.font = Font(bold=True)
    cell.fill = HEADER_FILL
    cell.border = BORDER
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_summary(ws, report: dict) -> None:
    kpi = report["kpi"]
    order = ["Всего задач", "Выполнено", "Не выполнено", "% выполнения",
             "Просрочено", "Скоро (≤ 7 дней)", "Без дедлайна",
             "Высокий приоритет, не выполнено"]

    ws["A1"] = "ОТЧЁТ ПО ЗАДАЧАМ"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Дата формирования: {report['today'].strftime('%d.%m.%Y')}"
    ws["A2"].font = Font(bold=True)

    ws["A4"] = "Метрика"
    ws["B4"] = "Значение"
    _header(ws["A4"])
    _header(ws["B4"])
    for i, key in enumerate(order):
        r = 5 + i
        ws.cell(row=r, column=1, value=key).border = BORDER
        ws.cell(row=r, column=2, value=kpi[key]).border = BORDER
    for r in range(5, 13):
        ws.cell(row=r, column=2).number_format = "0"
    ws["B8"].number_format = "0,0"

    ws["A14"] = "Приоритет"
    ws["B14"] = "Всего"
    ws["C14"] = "Выполнено"
    ws["D14"] = "Не выполнено"
    ws["E14"] = "% выполнения"
    for c in range(1, 6):
        _header(ws.cell(row=14, column=c))
    for i, (label, total, done, pend, p) in enumerate(report["prio_rows"]):
        r = 15 + i
        for c, v in zip(range(1, 6), (label, total, done, pend, p)):
            ws.cell(row=r, column=c, value=v)
            ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=5).number_format = "0,0"
    r = 18
    ws.cell(row=r, column=1, value="ИТОГО")
    for c, v in zip(range(2, 6), (report["prio_totals"]["total"],
                                 report["prio_totals"]["done"],
                                 report["prio_totals"]["pend"],
                                 report["prio_totals"]["pct"])):
        ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=c).font = Font(bold=True)
        ws.cell(row=r, column=c).border = BORDER


def write_tasks(ws, report: dict) -> None:
    headers = ["ID", "Задача", "Статус", "Приоритет", "Ранг", "Дедлайн",
               "Осталось дней", "Статус дедлайна", "Создана", "Возраст, дн."]
    for c, h in enumerate(headers, start=1):
        _header(ws.cell(row=1, column=c, value=h))
    for i, row in enumerate(report["rows"]):
        r = 2 + i
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if c in (6, 9):
                cell.number_format = DATE_FMT
    ws.freeze_panes = "A2"


def write_deadlines(ws, report: dict) -> None:
    ws["A1"] = "Группа дедлайна"
    ws["B1"] = "Всего"
    ws["C1"] = "Выполнено"
    ws["D1"] = "Не выполнено"
    for c in range(1, 5):
        _header(ws.cell(row=1, column=c))
    for i, g in enumerate(GROUPS):
        r = 2 + i
        ws.cell(row=r, column=1, value=g)
        grp = report["groups"][g]
        for c, v in zip(range(2, 5), (grp["Всего"], grp["Выполнено"],
                                     grp["Не выполнено"])):
            ws.cell(row=r, column=c, value=v)
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER
    r = 6
    ws.cell(row=r, column=1, value="ИТОГО")
    tot = report["group_totals"]
    for c, v in zip(range(2, 5), (tot["total"], tot["done"], tot["pend"])):
        ws.cell(row=r, column=c, value=v)
    for c in range(1, 5):
        ws.cell(row=r, column=c).font = Font(bold=True)
        ws.cell(row=r, column=c).border = BORDER


def write_report(report: dict, path: Path) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, fn in (("Сводка", write_summary), ("Задачи", write_tasks),
                     ("Дедлайны", write_deadlines)):
        ws = wb.create_sheet(name)
        fn(ws, report)
    for name, widths in WIDTHS.items():
        for col, w in widths.items():
            wb[name].column_dimensions[col].width = w
    wb.save(path)


def main() -> None:
    report = build_report(load_tasks())
    write_report(report, OUT_PATH)
    print(f"Отчёт сохранён: {OUT_PATH}")


if __name__ == "__main__":
    main()