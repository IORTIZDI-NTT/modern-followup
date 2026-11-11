from __future__ import annotations

import argparse
import calendar
import math
import re
import shutil
import subprocess
import tempfile
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape
from openpyxl import load_workbook


CLIENT_RATE = 28.36
MONTHS_ES = (
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
)
WEEKDAY_COLUMNS = {0: "L", 1: "M", 2: "X", 3: "J", 4: "V"}
ABSENCE_LABELS = {
    "V": "Vacaciones",
    "LA": "Ausencia / licencia",
    "F": "Festivo",
    "FI": "Festivo",
}


@dataclass
class Quality:
    warnings: list[str]
    sources: list[dict[str, Any]]
    unmatched_actual: list[str]
    unmatched_schedule: list[str]


def normalize_name(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    normalized = " ".join(re.sub(r"[^a-zA-Z0-9 ]+", " ", text).lower().split())
    aliases = {"juanjo": "juan jose", "manu": "manuel"}
    normalized = " ".join(aliases.get(token, token) for token in normalized.split())
    return normalized


def number(value: Any, default: float = 0.0) -> float:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return default
    if isinstance(value, str):
        cleaned = value.strip().replace("€", "").replace(" ", "")
        if not cleaned or cleaned in {"-", "0xf", "#VALUE!", "#N/A"}:
            return default
        if "," in cleaned and "." not in cleaned:
            cleaned = cleaned.replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def excel_date(value: Any) -> datetime | None:
    serial = number(value, default=-1)
    if not 30_000 <= serial <= 60_000:
        return None
    return datetime(1899, 12, 30) + timedelta(days=serial)


def is_target_month(value: Any, target: datetime) -> bool:
    date = excel_date(value)
    return bool(date and (date.year, date.month) == (target.year, target.month))


def period_label(target: datetime) -> str:
    return f"{MONTHS_ES[target.month - 1].capitalize()} {target.year}"


def find_one(data_dir: Path, patterns: Iterable[str], label: str) -> Path:
    matches: list[Path] = []
    for pattern in patterns:
        matches.extend(data_dir.glob(pattern))
    matches = sorted(set(matches), key=lambda path: path.stat().st_mtime, reverse=True)
    if not matches:
        raise FileNotFoundError(f"No se encontró {label} en {data_dir}")
    return matches[0]


def style_workbook(path: Path) -> Path:
    """Return an XLSX copy used only to inspect formatting from an XLSB."""
    cache_dir = path.parent.parent / "processed"
    cache_dir.mkdir(parents=True, exist_ok=True)
    cached = cache_dir / f"{path.stem}.styles.xlsx"
    if cached.exists() and cached.stat().st_mtime >= path.stat().st_mtime:
        return cached
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("Se necesita LibreOffice/soffice para identificar las filas rojas inactivas de Equipo")
    with tempfile.TemporaryDirectory(prefix="modern-geco-") as temporary:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "xlsx", "--outdir", temporary, str(path.resolve())],
            check=True,
            capture_output=True,
            text=True,
            timeout=180,
        )
        converted = Path(temporary) / f"{path.stem}.xlsx"
        if not converted.exists():
            raise RuntimeError("LibreOffice no generó la copia XLSX necesaria para leer estilos")
        shutil.move(converted, cached)
    return cached


def inactive_team_names(path: Path, first_row: int, last_row: int) -> list[str]:
    workbook = load_workbook(style_workbook(path), read_only=True, data_only=False)
    sheet = workbook["Equipo"]
    inactive: list[str] = []
    for row_number in range(first_row, last_row + 1):
        cell = sheet.cell(row_number, 4)
        color = cell.font.color
        rgb = str(color.rgb).upper() if color and color.type == "rgb" and color.rgb else ""
        if rgb.endswith("C00000") and cell.value:
            inactive.append(str(cell.value).strip())
    workbook.close()
    return inactive


def best_match(source: str, candidates: Iterable[str]) -> tuple[str | None, float]:
    source_norm = normalize_name(source)
    if not source_norm:
        return None, 0.0
    scored: list[tuple[float, str]] = []
    source_tokens = set(source_norm.split())
    for candidate in candidates:
        candidate_norm = normalize_name(candidate)
        candidate_tokens = set(candidate_norm.split())
        if source_norm == candidate_norm:
            score = 1.0
        elif source_tokens and (source_tokens <= candidate_tokens or candidate_tokens <= source_tokens):
            score = 0.96
        elif source_tokens and all(
            any(SequenceMatcher(None, source_token, candidate_token).ratio() >= 0.88 for candidate_token in candidate_tokens)
            for source_token in source_tokens
        ):
            score = 0.94
        else:
            score = SequenceMatcher(None, source_norm, candidate_norm).ratio()
        scored.append((score, candidate))
    scored.sort(reverse=True)
    if not scored or scored[0][0] < 0.72:
        return None, scored[0][0] if scored else 0.0
    if len(scored) > 1 and scored[0][0] - scored[1][0] < 0.03 and scored[0][0] < 0.95:
        return None, scored[0][0]
    return scored[0][1], scored[0][0]


def load_forecast(path: Path, target: datetime) -> dict[str, Any]:
    frame = pd.read_excel(path, sheet_name="Forecast", header=None, engine="pyxlsb")
    date_row = frame.iloc[2]
    previous_columns = [column for column in range(0, 10) if excel_date(date_row.iloc[column])]
    current_columns = [column for column in range(10, 18) if excel_date(date_row.iloc[column])]
    adjustment_columns = [column for column in range(18, 25) if excel_date(date_row.iloc[column])]
    difference_columns = [column for column in range(25, len(date_row)) if excel_date(date_row.iloc[column])]
    current_position = next(
        (position for position, column in enumerate(current_columns) if is_target_month(date_row.iloc[column], target)),
        None,
    )
    if current_position is None:
        raise ValueError(f"Forecast actual no contiene {period_label(target)}")
    blocks = {
        "previous": next((column for column in previous_columns if is_target_month(date_row.iloc[column], target)), None),
        "current": current_columns[current_position],
        # En la GECO las fechas de este bloque están desplazadas un mes. La
        # correspondencia válida es posicional respecto a FORECAST ACTUAL.
        "adjustment": adjustment_columns[current_position] if current_position < len(adjustment_columns) else None,
        "difference": difference_columns[current_position] if current_position < len(difference_columns) else None,
    }
    if blocks["previous"] is None or blocks["current"] is None:
        raise ValueError(f"Forecast no contiene los bloques anterior y actual para {period_label(target)}")

    rows: list[dict[str, Any]] = []
    current_project = ""
    for _, row in frame.iterrows():
        if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() == "Total":
            current_project = "Total"
        elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():
            current_project = str(row.iloc[0]).strip()
        elif pd.notna(row.iloc[1]) and str(row.iloc[1]).strip():
            # Las secciones auxiliares posteriores (por ejemplo TOTAL AS) no
            # deben heredar la etiqueta del total contractual.
            current_project = ""
        metric = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if metric != "Net Revenue" or not current_project:
            continue
        item = {"project": current_project}
        for key, column in blocks.items():
            item[key] = number(row.iloc[column]) if column is not None else 0.0
        item["without_adjustment"] = item["current"] - item["adjustment"]
        if current_project == "Total":
            totals = item
        else:
            rows.append(item)
    if "totals" not in locals():
        raise ValueError("No se encontró la fila Total / Net Revenue en Forecast")
    totals["change_pct"] = (
        (totals["current"] - totals["previous"]) / totals["previous"]
        if totals["previous"] else 0.0
    )
    total_row = next(
        index for index, row in frame.iterrows()
        if str(row.iloc[1]).strip() == "Total" and str(row.iloc[2]).strip() == "Net Revenue"
    )
    total_metrics: dict[str, float] = {}
    for row_index in range(total_row, min(total_row + 6, len(frame))):
        metric = str(frame.iat[row_index, 2]).strip()
        if metric in {"Net Revenue", "Total Costs", "Contract Margin"}:
            total_metrics[metric] = number(frame.iat[row_index, blocks["current"]])
    current_revenue = total_metrics.get("Net Revenue", totals["current"])
    current_costs = total_metrics.get("Total Costs", 0.0)
    current_margin = total_metrics.get("Contract Margin", current_revenue - current_costs)
    adjustment_revenue = totals["adjustment"]
    adjustment_costs = 0.0
    if blocks["adjustment"] is not None:
        cost_row = next(
            (row_index for row_index in range(total_row, min(total_row + 6, len(frame)))
             if str(frame.iat[row_index, 2]).strip() == "Total Costs"),
            None,
        )
        if cost_row is not None:
            adjustment_costs = number(frame.iat[cost_row, blocks["adjustment"]])
    financials = {
        "revenue": {"with_adjustment": current_revenue, "without_adjustment": current_revenue - adjustment_revenue, "adjustment": adjustment_revenue},
        "costs": {"with_adjustment": current_costs, "without_adjustment": current_costs - adjustment_costs, "adjustment": adjustment_costs},
        "margin": {
            "with_adjustment": current_margin,
            "without_adjustment": current_margin - (adjustment_revenue - adjustment_costs),
            "adjustment": adjustment_revenue - adjustment_costs,
        },
    }
    financials["margin"]["with_pct"] = current_margin / current_revenue if current_revenue else 0.0
    financials["margin"]["without_pct"] = (
        financials["margin"]["without_adjustment"] / financials["revenue"]["without_adjustment"]
        if financials["revenue"]["without_adjustment"] else 0.0
    )
    return {"totals": totals, "projects": rows, "financials": financials}


def month_column(frame: pd.DataFrame, header_row: int, target: datetime, sheet: str) -> int:
    matches = [column for column, value in enumerate(frame.iloc[header_row]) if is_target_month(value, target)]
    if not matches:
        raise ValueError(f"{sheet} no contiene {period_label(target)}")
    return matches[0]


def load_team(path: Path, target: datetime) -> dict[str, Any]:
    frame = pd.read_excel(path, sheet_name="Equipo", header=None, engine="pyxlsb")
    start = month_column(frame, 1, target, "Equipo")
    fields = ["planned_hours", "internal_cost", "cost_with_overhead", "objective_hours",
              "gva_hours", "hours_difference", "forecast_hours", "increase", "forecast_difference", "return"]
    table_end = next(
        index for index, row in frame.iterrows()
        if str(row.iloc[2]).strip() == "TOTAL" and str(row.iloc[3]).strip() == "Total horas"
    )
    inactive_names = inactive_team_names(path, first_row=4, last_row=table_end)
    inactive_normalized = {normalize_name(name) for name in inactive_names}
    people: list[dict[str, Any]] = []
    inactive_people: list[dict[str, Any]] = []
    for _, row in frame.iloc[3:table_end].iterrows():
        name = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        team = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        unit = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not name or not team or name.lower().startswith(("total", "ajustes")):
            continue
        person = {"name": name, "team": team, "task": team, "unit": unit}
        for offset, field in enumerate(fields):
            person[field] = number(row.iloc[start + offset])
        if normalize_name(name) in inactive_normalized:
            inactive_people.append(person)
        else:
            people.append(person)
    summaries: dict[str, list[dict[str, Any]]] = {}
    for title, key in (("H. POR EQUIPOS", "by_team"), ("H. POR UNIDAD", "by_unit")):
        header_index = next(
            (index for index, row in frame.iterrows() if str(row.iloc[3]).strip() == title),
            None,
        )
        summary_rows: list[dict[str, Any]] = []
        if header_index is not None:
            for row_index in range(header_index + 1, len(frame)):
                label = str(frame.iat[row_index, 3]).strip() if pd.notna(frame.iat[row_index, 3]) else ""
                if not label:
                    if summary_rows:
                        break
                    continue
                summary_rows.append({
                    "label": label,
                    "planned_hours": number(frame.iat[row_index, start]),
                    "objective_hours": number(frame.iat[row_index, start + 3]),
                    "gva_hours": number(frame.iat[row_index, start + 4]),
                    "forecast_hours": number(frame.iat[row_index, start + 6]),
                })
                if label == "TOTAL":
                    break
        summaries[key] = summary_rows
    return {"people": people, "inactive_people": inactive_people, "inactive_names": inactive_names, "summaries": summaries}


def load_income(path: Path, target: datetime) -> list[dict[str, Any]]:
    frame = pd.read_excel(path, sheet_name="Ingresos", header=None, engine="pyxlsb")
    start = month_column(frame, 1, target, "Ingresos")
    fields = ["income_hours", "income_cost", "income_cost_overhead", "income_objective_hours",
              "efficient_hours", "planned_income", "contract_margin_pct"]
    rows: list[dict[str, Any]] = []
    for _, row in frame.iloc[3:].iterrows():
        name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not name or name.lower().startswith(("total", "ajustes")):
            continue
        item = {"name": name, "role": str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""}
        for offset, field in enumerate(fields):
            item[field] = number(row.iloc[start + offset])
        rows.append(item)
    return rows


def load_schedule(path: Path, target: datetime) -> list[dict[str, Any]]:
    wanted = f"{MONTHS_ES[target.month - 1]} {target.year}".upper()
    book = pd.ExcelFile(path, engine="openpyxl")
    sheet = next((name for name in book.sheet_names if name.strip().upper() == wanted), None)
    if sheet is None:
        raise ValueError(f"Horario laboral no contiene la hoja {wanted}")
    raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
    header_row = next(
        (index for index, row in raw.iterrows() if str(row.iloc[0]).strip() == "Equipo" and str(row.iloc[1]).strip() == "Usuario"),
        None,
    )
    if header_row is None:
        raise ValueError(f"No se encontró la cabecera Equipo/Usuario en {sheet}")
    weekday_counts = {weekday: 0 for weekday in range(5)}
    _, days = calendar.monthrange(target.year, target.month)
    for day in range(1, days + 1):
        weekday = datetime(target.year, target.month, day).weekday()
        if weekday < 5:
            weekday_counts[weekday] += 1
    rows: list[dict[str, Any]] = []
    summary_prefixes = ("datos ", "declarado ", "deuda ")
    for _, row in raw.iloc[header_row + 1:].iterrows():
        name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if not name or name.casefold().startswith(summary_prefixes):
            continue
        hours = {WEEKDAY_COLUMNS[index]: number(row.iloc[index + 2]) for index in range(5)}
        capacity = sum(hours[WEEKDAY_COLUMNS[index]] * count for index, count in weekday_counts.items())
        rows.append({
            "name": name,
            "schedule_team": str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else "",
            "weekday_hours": hours,
            "capacity_hours": capacity,
        })
    return rows


def load_incurred(path: Path, target: datetime) -> tuple[list[dict[str, Any]], list[dict[str, Any]], set[str]]:
    frame = pd.read_html(path, decimal=",", thousands=None, keep_default_na=False)[0]
    total_column = next(column for column in frame.columns if str(column).strip().lower() == "total validado")
    records: dict[str, dict[str, Any]] = {}
    absences: list[dict[str, Any]] = []
    codes: set[str] = set()
    day_columns = [column for column in frame.columns if re.fullmatch(r"\d{2}", str(column))]
    for _, row in frame.iterrows():
        name = str(row.iloc[0]).strip()
        category = str(row.iloc[1]).strip().upper()
        if not name:
            continue
        record = records.setdefault(name, {
            "name": name,
            "client_hours": 0.0,
            "ntt_hours": 0.0,
            "client_daily": {},
            "ntt_daily": {},
            "daily_codes": {},
        })
        total = number(row[total_column])
        if category == "CLIENTE":
            record["client_hours"] += total
        elif category == "NTT":
            record["ntt_hours"] += total
        for day_column in day_columns:
            raw_value = str(row[day_column]).strip().upper()
            if not raw_value or raw_value == "-":
                continue
            day = int(str(day_column))
            parsed_value = number(raw_value, default=float("nan"))
            if not math.isnan(parsed_value):
                if category == "CLIENTE":
                    record["client_daily"][day] = number(record["client_daily"].get(day)) + parsed_value
                elif category == "NTT":
                    record["ntt_daily"][day] = number(record["ntt_daily"].get(day)) + parsed_value
                continue
            try:
                date = datetime(target.year, target.month, day)
            except ValueError:
                continue
            record["daily_codes"][day] = raw_value
            key = (normalize_name(name), date.date(), raw_value)
            if any(event["key"] == key for event in absences):
                continue
            codes.add(raw_value)
            absences.append({
                "key": key,
                "name": name,
                "date": date,
                "date_label": date.strftime("%d/%m/%Y"),
                "code": raw_value,
                "label": ABSENCE_LABELS.get(raw_value, "Ausencia (código sin clasificar)"),
            })
    result = []
    for record in records.values():
        record["total_recorded_hours"] = record["client_hours"] + record["ntt_hours"]
        # CLIENTE representa las horas que se facturarán al cliente este mes.
        record["incurred_hours"] = record["client_hours"]
        result.append(record)
    absences.sort(key=lambda item: (item["date"], item["name"]))
    return result, absences, codes


def merge_people(
    team: list[dict[str, Any]],
    income: list[dict[str, Any]],
    schedule: list[dict[str, Any]],
    incurred: list[dict[str, Any]],
    inactive_names: list[str],
    quality: Quality,
) -> list[dict[str, Any]]:
    names = [person["name"] for person in team]

    def index_rows(rows: list[dict[str, Any]], unmatched: list[str] | None = None) -> dict[str, dict[str, Any]]:
        index: dict[str, dict[str, Any]] = {}
        for row in rows:
            match, _ = best_match(row["name"], names)
            if match:
                if match in index and "incurred_hours" in row:
                    for field in ("client_hours", "ntt_hours", "incurred_hours", "total_recorded_hours"):
                        index[match][field] = index[match].get(field, 0.0) + row.get(field, 0.0)
                else:
                    index[match] = row
            elif unmatched is not None:
                inactive_match, _ = best_match(row["name"], inactive_names)
                if not inactive_match:
                    unmatched.append(row["name"])
        return index

    income_index = index_rows(income)
    schedule_index = index_rows(schedule, quality.unmatched_schedule)
    incurred_index = index_rows(incurred, quality.unmatched_actual)
    merged: list[dict[str, Any]] = []
    for person in team:
        name = person["name"]
        item = dict(person)
        item.update(income_index.get(name, {}))
        item["name"] = name
        schedule_data = schedule_index.get(name, {})
        item.update({key: value for key, value in schedule_data.items() if key != "name"})
        item["team"] = str(schedule_data.get("schedule_team") or "SIN EQUIPO")
        item.update({key: value for key, value in incurred_index.get(name, {}).items() if key != "name"})
        for field in ("capacity_hours", "client_hours", "ntt_hours", "incurred_hours", "efficient_hours", "planned_income"):
            item[field] = number(item.get(field))
        item["pending_hours"] = max(item["planned_hours"] - item["incurred_hours"], 0.0)
        item["progress_pct"] = item["incurred_hours"] / item["planned_hours"] if item["planned_hours"] else 0.0
        efficiency = item["efficient_hours"] / item.get("income_hours", 0.0) if item.get("income_hours", 0.0) else 0.0
        item["efficiency"] = efficiency
        item["estimated_incurred_income"] = item["client_hours"] * CLIENT_RATE
        item["capacity_remaining"] = max(item["capacity_hours"] - item["incurred_hours"], 0.0)
        merged.append(item)
    merged.sort(key=lambda item: (-item["incurred_hours"], item["name"]))
    return merged


def build_calendar(people: list[dict[str, Any]], target: datetime) -> dict[str, Any]:
    weekday_labels = ("L", "M", "X", "J", "V", "S", "D")
    _, day_count = calendar.monthrange(target.year, target.month)
    days = []
    daily_totals = {day: {"actual": 0.0, "planned": 0.0} for day in range(1, day_count + 1)}
    for day in range(1, day_count + 1):
        date = datetime(target.year, target.month, day)
        days.append({"day": day, "weekday": weekday_labels[date.weekday()], "weekend": date.weekday() >= 5})
    rows: list[dict[str, Any]] = []
    for person in sorted(people, key=lambda item: item["name"]):
        cells = []
        schedule_hours = person.get("weekday_hours", {})
        client_daily = person.get("client_daily", {})
        ntt_daily = person.get("ntt_daily", {})
        daily_codes = person.get("daily_codes", {})
        for day_info in days:
            day = day_info["day"]
            date = datetime(target.year, target.month, day)
            weekday_key = WEEKDAY_COLUMNS.get(date.weekday())
            planned = number(schedule_hours.get(weekday_key)) if weekday_key else 0.0
            actual = number(client_daily.get(day))
            ntt = number(ntt_daily.get(day))
            code = str(daily_codes.get(day, ""))
            if person.get("is_dummy") and actual > 0:
                status = "dummy"
            elif code:
                status = "absence"
            elif day_info["weekend"] and actual == 0:
                status = "weekend"
            elif planned == 0 and actual > 0:
                status = "over"
            elif planned > 0 and actual == 0:
                status = "missing"
            elif actual > planned + 0.01:
                status = "over"
            elif actual < planned - 0.01:
                status = "under"
            else:
                status = "match"
            cells.append({
                "day": day,
                "actual": actual,
                "planned": planned,
                "ntt": ntt,
                "code": code,
                "status": status,
                "tooltip": f"CLIENTE {actual:.2f} h · Horario {planned:.2f} h · NTT {ntt:.2f} h" + (f" · {code}" if code else ""),
            })
            daily_totals[day]["actual"] += actual
            daily_totals[day]["planned"] += planned
        rows.append({
            "name": person["name"],
            "team": person["team"],
            "unit": person["unit"],
            "efficiency": person["efficiency"],
            "client_hours": person["client_hours"],
            "capacity_hours": person["capacity_hours"],
            "difference": person["client_hours"] - person["capacity_hours"],
            "is_dummy": bool(person.get("is_dummy")),
            "cells": cells,
        })
    return {"days": days, "rows": rows, "daily_totals": daily_totals}


def build_context(data_dir: Path, target: datetime) -> dict[str, Any]:
    geco = find_one(data_dir, ("*.xlsb", "*.XLSB"), "la GECO (.xlsb)")
    schedule_file = find_one(data_dir, ("*Horario*.xlsx", "*horario*.xlsx"), "el horario laboral (.xlsx)")
    incurred_file = find_one(data_dir, (f"{target:%Y%m}-Incurrido*.html", f"*{target:%Y%m}*Incurrido*.html"), "el incurrido mensual (.html)")
    quality = Quality(warnings=[], sources=[], unmatched_actual=[], unmatched_schedule=[])
    for label, path in (("GECO", geco), ("Horario laboral", schedule_file), ("Incurrido", incurred_file)):
        quality.sources.append({
            "label": label,
            "file": path.name,
            "size_mb": path.stat().st_size / 1_048_576,
            "modified": datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
        })
    forecast = load_forecast(geco, target)
    team_data = load_team(geco, target)
    income = load_income(geco, target)
    schedule = load_schedule(schedule_file, target)
    dummy_schedule_names = [row["name"] for row in schedule if row["schedule_team"].strip().upper() == "DUMMY"]
    dummy_people: list[dict[str, Any]] = []
    remaining_inactive_names: list[str] = []
    for person in team_data["inactive_people"]:
        dummy_match, _ = best_match(person["name"], dummy_schedule_names)
        if dummy_match:
            person["is_dummy"] = True
            dummy_people.append(person)
        else:
            remaining_inactive_names.append(person["name"])
    team = [{**person, "is_dummy": False} for person in team_data["people"]] + dummy_people
    incurred, absences, absence_codes = load_incurred(incurred_file, target)
    people = merge_people(team, income, schedule, incurred, remaining_inactive_names, quality)
    missing_team_names = [person["name"] for person in people if person["team"] == "SIN EQUIPO"]
    if missing_team_names:
        quality.warnings.append(
            f"{len(missing_team_names)} perfiles activos de GECO no tienen equipo en el horario laboral."
        )
    if quality.unmatched_actual:
        quality.warnings.append(f"{len(quality.unmatched_actual)} nombres del incurrido no enlazan automáticamente con Equipo.")
    if quality.unmatched_schedule:
        quality.warnings.append(f"{len(quality.unmatched_schedule)} nombres del horario no enlazan automáticamente con Equipo.")
    unknown_codes = sorted(code for code in absence_codes if code not in ABSENCE_LABELS)
    if unknown_codes:
        quality.warnings.append("Códigos de ausencia pendientes de clasificar: " + ", ".join(unknown_codes))
    totals = forecast["totals"]
    raw_client_hours = sum(record["client_hours"] for record in incurred)
    active_client_hours = sum(person["incurred_hours"] for person in people)
    kpis = {
        "forecast_previous": totals["previous"],
        "forecast_current": totals["current"],
        "forecast_current_client_hours": totals["current"] / CLIENT_RATE if CLIENT_RATE else 0.0,
        "forecast_adjustment": totals["adjustment"],
        "adjustment_client_hours": totals["adjustment"] / CLIENT_RATE if CLIENT_RATE else 0.0,
        "forecast_change": totals["difference"],
        "planned_hours": sum(person["planned_hours"] for person in people),
        "incurred_hours": raw_client_hours,
        "active_incurred_hours": active_client_hours,
        "outside_active_hours": raw_client_hours - active_client_hours,
        "pending_hours": sum(person["pending_hours"] for person in people),
        "capacity_hours": sum(person["capacity_hours"] for person in people),
        "planned_income": sum(person["planned_income"] for person in people),
        "estimated_incurred_income": raw_client_hours * CLIENT_RATE,
        "people_with_plan": sum(1 for person in people if person["planned_hours"] > 0),
        "absence_events": len(absences),
    }
    kpis["hours_progress_pct"] = active_client_hours / kpis["planned_hours"] if kpis["planned_hours"] else 0.0
    efficiency_people = [person for person in people if not person.get("is_dummy")]
    billable_hours = sum(person["client_hours"] for person in efficiency_people)
    weighted_efficiency = (
        sum(person["client_hours"] * person["efficiency"] for person in efficiency_people) / billable_hours
        if billable_hours else 0.0
    )
    kpis["weighted_efficiency"] = weighted_efficiency
    kpis["adjustment_team_hours"] = (
        kpis["adjustment_client_hours"] / weighted_efficiency if weighted_efficiency else 0.0
    )
    kpis["dummy_hours"] = sum(person["client_hours"] for person in people if person.get("is_dummy"))
    kpis["employee_client_hours"] = raw_client_hours - kpis["dummy_hours"]
    task_groups: dict[str, dict[str, float | str]] = {}
    for person in people:
        if person["client_hours"] <= 0 or person.get("is_dummy"):
            continue
        task = str(person.get("role") or person.get("team") or "Sin tarea")
        group = task_groups.setdefault(task, {"task": task, "client_hours": 0.0, "effective_hours": 0.0})
        group["client_hours"] = number(group["client_hours"]) + person["client_hours"]
        group["effective_hours"] = number(group["effective_hours"]) + person["client_hours"] * person["efficiency"]
    task_efficiencies: list[dict[str, Any]] = []
    for group in task_groups.values():
        client_hours = number(group["client_hours"])
        effective_hours = number(group["effective_hours"])
        task_efficiencies.append({
            **group,
            "efficiency": effective_hours / client_hours if client_hours else 0.0,
            "estimated_income": effective_hours * CLIENT_RATE,
        })
    task_efficiencies.sort(key=lambda item: number(item["client_hours"]), reverse=True)
    active_people = [person for person in people if not person.get("is_dummy")]
    dummy_profiles = [person for person in people if person.get("is_dummy")]
    team_totals: dict[str, dict[str, Any]] = {}
    for person in people:
        team_name = person["team"]
        group = team_totals.setdefault(team_name, {
            "team": team_name, "profiles": 0, "planned_hours": 0.0,
            "client_hours": 0.0, "ntt_hours": 0.0, "planned_income": 0.0,
        })
        group["profiles"] += 1
        for field in ("planned_hours", "client_hours", "ntt_hours", "planned_income"):
            group[field] += number(person.get(field))
    computed_team_totals = sorted(team_totals.values(), key=lambda item: (-number(item["client_hours"]), item["team"]))
    return {
        "period": target.strftime("%Y-%m"),
        "period_label": period_label(target),
        "generated_at": datetime.now(ZoneInfo("Europe/Madrid")).strftime("%d/%m/%Y %H:%M"),
        "client_rate": CLIENT_RATE,
        "forecast": forecast,
        "people": active_people,
        "calculation_people": people,
        "dummy_profiles": dummy_profiles,
        "computed_team_totals": computed_team_totals,
        "team_summaries": team_data["summaries"],
        "inactive_count": len(remaining_inactive_names),
        "task_efficiencies": task_efficiencies,
        "calendar": build_calendar(people, target),
        "absences": absences,
        "absence_codes": sorted(absence_codes),
        "quality": quality,
        "kpis": kpis,
    }


def render_site(context: dict[str, Any], output_dir: Path) -> None:
    project_root = Path(__file__).resolve().parents[1]
    templates_dir = project_root / "templates"
    env = Environment(
        loader=FileSystemLoader(templates_dir),
        autoescape=select_autoescape(("html", "xml")),
        trim_blocks=True,
        lstrip_blocks=True,
    )
    env.filters["money"] = lambda value: f"{number(value):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    env.filters["hours"] = lambda value: f"{number(value):,.1f} h".replace(",", "X").replace(".", ",").replace("X", ".")
    env.filters["percent"] = lambda value: f"{number(value) * 100:.1f}%".replace(".", ",")
    env.filters["decimal"] = lambda value: f"{number(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    output_dir.mkdir(parents=True, exist_ok=True)
    assets = output_dir / "assets"
    assets.mkdir(exist_ok=True)
    shutil.copy2(project_root / "static" / "style.css", assets / "style.css")
    pages = {
        "index.html": "dashboard.html",
        "economia.html": "economia.html",
        "equipo.html": "equipo.html",
        "calendario.html": "calendario.html",
        "calidad.html": "calidad.html",
    }
    for destination, template_name in pages.items():
        html = env.get_template(template_name).render(**context, current_page=destination)
        (output_dir / destination).write_text(html, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    default_month = datetime.now(ZoneInfo("Europe/Madrid")).strftime("%Y-%m")
    parser = argparse.ArgumentParser(description="Genera el informe HTML mensual de Modernización Lote 4")
    parser.add_argument("--data-dir", type=Path, default=Path("data/input"))
    parser.add_argument("--output-dir", type=Path, default=Path("site"))
    parser.add_argument("--month", default=default_month, help="Mes YYYY-MM; por defecto, el actual")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    target = datetime.strptime(args.month, "%Y-%m")
    context = build_context(args.data_dir, target)
    render_site(context, args.output_dir)
    print(f"Informe generado: {(args.output_dir / 'index.html').resolve()}")
    print(f"Periodo: {context['period_label']}; personas: {len(context['people'])}; alertas: {len(context['quality'].warnings)}")


if __name__ == "__main__":
    main()
