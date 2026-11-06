#!/usr/bin/env python3
"""Генерация отчёта по задачам (HOMEWORK 1.1) в .xlsx.

Читает apps/todo_db.json, считает метрики, пишет три листа:
«Сводка», «Задачи», «Дедлайны». Значения записываются статикой на
дату генерации (детерминированно между прогонами в один день).
"""
from __future__ import annotations

import json
import re
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DB_PATH = BASE / "apps" / "todo_db.json"
PRI = {"low": "Низкий", "normal": "Обычный", "high": "Высокий"}
RANK = {"low": 1, "normal": 2, "high": 3}
GROUPS = ("Просрочено", "Скоро (<= 7 дней)", "В срок", "Нет дедлайна")
GROUP_LABELS = {
    "Просрочено": "Просрочено",
    "Скоро (<= 7 дней)": "Скоро (≤ 7 дней)",
    "В срок": "В срок",
    "Нет дедлайна": "Нет дедлайна",
}

TODAY = date.today()
TODAY_STR = TODAY.strftime("%d.%m.%Y")
TODAY_ISO = TODAY.isoformat()


def load_tasks() -> list[dict]:
    doc = json.loads(DB_PATH.read_text(encoding="utf-8"))
    return doc["payload"]["tasks"]


def due_date(t: dict) -> date | None:
    return date.fromisoformat(t["due"]) if t.get("due") else None


def deadline_state(t: dict) -> tuple[str, int | None]:
    """Статус дедлайна и осталось дней для одной задачи."""
    d = due_date(t)
    if t["done"]:
        return "Выполнена", None
    if d is None:
        return "Нет дедлайна", None
    diff = (d - TODAY).days
    if diff < 0:
        return "Просрочена", diff
    if diff <= 7:
        return "Скоро", diff
    return "В срок", diff


def build_rows(tasks: list[dict]) -> list[list]:
    rows = []
    for t in tasks:
        d = due_date(t)
        state, remain = deadline_state(t)
        created = date.fromisoformat(t["created_at"][:10])
        rows.append([
            t["id"],
            t["title"],
            "Выполнена" if t["done"] else "Не выполнена",
            PRI[t["priority"]],
            RANK[t["priority"]],
            d,
            remain,
            state,
            created,
            (TODAY - created).days,
        ])
    rows.sort(key=lambda r: (
        r[2] != "Не выполнена",       # не выполненные первыми
        r[4],                          # ранг: высокий -> низкий
        r[5] is None,                  # пустой дедлайн — в конец
        r[5] or date.min,
        r[0],                          # id
    ))
    return rows


def build_report(tasks: list[dict]) -> dict[str, object]:
    total = len(tasks)
    done = [t for t in tasks if t["done"]]
    pend = [t for t in tasks if not t["done"]]
    done_n = len(done)

    overdue = [t for t in pend if (d := due_date(t)) and d < TODAY]
    soon = [t for t in pend if (d := due_date(t)) and TODAY <= d <= TODAY + timedelta(days=7)]
    no_due = [t for t in pend if t["due"] is None]
    high_pend = [t for t in pend if t["priority"] == "high"]

    kpi = {
        "Всего задач": total,
        "Выполнено": done_n,
        "Не выполнено": total - done_n,
        "% выполнения": round(done_n / total * 100, 1),
        "Просрочено": len(overdue),
        "Скоро (<= 7 дней)": len(soon),
        "Без дедлайна": len(no_due),
        "Высокий приоритет, не выполнено": len(high_pend),
    }

    prio_rows = []
    totals = Counter()
    for p in ("high", "normal", "low"):
        group = [t for t in tasks if t["priority"] == p]
        dn = sum(1 for t in group if t["done"])
        prio_rows.append({
            "label": PRI[p], "total": len(group), "done": dn,
            "pend": len(group) - dn,
            "pct": round(dn / len(group) * 100, 1) if group else 0,
        })
        totals["total"] += len(group)
        totals["done"] += dn
        totals["pend"] += len(group) - dn

    # Группы дедлайна (любой статус выполнения)
    groups = {g: Counter() for g in GROUPS}
    for t in tasks:
        d = due_date(t)
        if d is None:
            g = "Нет дедлайна"
        elif d < TODAY:
            g = "Просрочено"
        elif d <= TODAY + timedelta(days=7):
            g = "Скоро (<= 7 дней)"
        else:
            g = "В срок"
        groups[g]["Всего"] += 1
        groups[g]["Выполнено" if t["done"] else "Не выполнено"] += 1

    return {
        "kpi": kpi,
        "prio_rows": prio_rows,
        "prio_totals": totals,
        "groups": groups,
        "rows": build_rows(tasks),
    }


# ----------------------------- оформление -----------------------------

THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
TITLE_FONT = Font(bold=True, size=14)
HEADER_FONT = Font(bold=True)
TOTAL_FONT = Font(bold=True)
DATE_FORMAT = "dd.mm.yyyy"


def style_header(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.border = BORDER
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_title(cell):
    cell.font = TITLE_FONT


# ----------------------------- запись ------------------------------

def write_summary(ws, report):
    kpi = report["kpi"]
    order = [
        "Всего задач", "Выполнено", "Не выполнено", "% выполнения",
        "Просрочено", "Скоро (<= 7 дней)", "Без дедлайна",
        "Высокий приоритет, не выполнено",
    ]
    ws["A1"] = "ОТЧЁТ ПО ЗАДАЧАМ"
    style_title(ws["A1"])
    ws["A2"] = f"Дата формирования: {TODAY_STR}"
    ws["A2"].font = Font(bold=True)

    ws["A4"] = "Метрика"
    ws["B4"] = "Значение"
    style_header(ws["A4"])
    style_header(ws["B4"])

    for i, key in enumerate(order):
        r = 5 + i
        ws.cell(row=r, column=1, value=key)
        ws.cell(row=r, column=2, value=kpi[key])
    ws["B8"].number_format = "0,0"
    for r in range(5, 13):
        ws.cell(row=r, column=2).number_format = "0"
        ws.cell(row=r, column=1).border = BORDER
        ws.cell(row=r, column=2).border = BORDER

    # Таблица по приоритетам
    ws["A14"] = "Приоритет"
    ws["B14"] = "Всего"
    ws["C14"] = "Выполнено"
    ws["D14"] = "Не выполнено"
    ws["E14"] = "% выполнения"
    for c in range(1, 6):
        style_header(ws.cell(row=14, column=c))

    pr = report["prio_rows"]
    for i, row in enumerate(pr):
        r = 15 + i
        ws.cell(row=r, column=1, value=row["label"])
        ws.cell(row=r, column=2, value=row["total"])
        ws.cell(row=r, column=3, value=row["done"])
        ws.cell(row=r, column=4, value=row["pend"])
        ws.cell(row=r, column=5, value=row["pct"])
        ws.cell(row=r, column=5).number_format = "0,0"
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BORDER

    r = 18
    ws.cell(row=r, column=1, value="ИТОГО")
    ws.cell(row=r, column=2, value=report["prio_totals"]["total"])
    ws.cell(row=r, column=3, value=report["prio_totals"]["done"])
    ws.cell(row=r, column=4, value=report["prio_totals"]["pend"])
    ws.cell(row=r, column=5, value=round(
        report["prio_totals"]["done"] / report["prio_totals"]["total"] * 100, 1))
    ws.cell(row=r, column=5).number_format = "0,0"
    for c in range(1, 6):
        cell = ws.cell(row=r, column=c)
        cell.font = TOTAL_FONT
        cell.border = BORDER


def write_tasks(ws, report):
    headers = ["ID", "Задача", "Статус", "Приоритет", "Ранг", "Дедлайн",
               "Осталось дней", "Статус дедлайна", "Создана", "Возраст, дн."]
    for c, h in enumerate(headers, start=1):
        style_header(ws.cell(row=1, column=c))

    for i, row in enumerate(report["rows"]):
        r = 2 + i
        ws.cell(row=r, column=1, value=row[0])
        ws.cell(row=r, column=2, value=row[1])
        ws.cell(row=r, column=3, value=row[2])
        ws.cell(row=r, column=4, value=row[3])
        ws.cell(row=r, column=5, value=row[4])
        if row[5] is not None:
            c = ws.cell(row=r, column=6, value=row[5])
            c.number_format = DATE_FORMAT
        ws.cell(row=r, column=7, value=row[6])
        ws.cell(row=r, column=8, value=row[7])
        if row[8] is not None:
            c = ws.cell(row=r, column=9, value=row[8])
            c.number_format = DATE_FORMAT
        ws.cell(row=r, column=10, value=row[9])
        for col in range(1, 11):
            ws.cell(row=r, column=col).border = BORDER

    ws.freeze_panes = "A2"


def write_deadlines(ws, report):
    ws["A1"] = "Группа дедлайна"
    ws["B1"] = "Всего"
    ws["C1"] = "Выполнено"
    ws["D1"] = "Не выполнено"
    for c in range(1, 5):
        style_header(ws.cell(row=1, column=c))

    for i, g in enumerate(GROUPS):
        r = 2 + i
        ws.cell(row=r, column=1, value=GROUP_LABELS[g])
        c = report["groups"][g]
        ws.cell(row=r, column=2, value=c["Всего"])
        ws.cell(row=r, column=3, value=c["Выполнено"])
        ws.cell(row=r, column=4, value=c["Не выполнено"])
        for col in range(1, 5):
            ws.cell(row=r, column=col).border = BORDER

    r = 6
    ws["A6"] = "ИТОГО"
    tot = Counter()
    for g in GROUPS:
        c = report["groups"][g]
        tot["Всего"] += c["Всего"]
        tot["Выполнено"] += c["Выполнено"]
        tot["Не выполнено"] += c["Не выполнено"]
    ws.cell(row=r, column=2, value=tot["Всего"])
    ws.cell(row=r, column=3, value=tot["Выполнено"])
    ws.cell(row=r, column=4, value=tot["Не выполнено"])
    for col in range(1, 5):
        cell = ws.cell(row=r, column=col)
        cell.font = TOTAL_FONT
        cell.border = BORDER


def apply_column_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def main():
    tasks = load_tasks()
    report = build_report(tasks)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sw = wb.create_sheet("Сводка")
    tw = wb.create_sheet("Задачи")
    dw = wb.create_sheet("Дедлайны")

    write_summary(sw, report)
    write_tasks(tw, report)
    write_deadlines(dw, report)

    apply_column_widths(sw, {"A": 34, "B": 12, "C": 3})
    apply_column_widths(tw, {"A": 5, "B": 42, "C": 14, "D": 12, "E": 6,
                             "F": 11, "G": 12, "H": 14, "I": 11, "J": 10})
    apply_column_widths(dw, {"A": 18, "B": 9, "C": 12, "D": 14})

    out = BASE / f"report_{TODAY_STR}.xlsx"
    wb.save(out)
    print(f"Отчёт сохранён: {out}")


if __name__ == "__main__":
    main()