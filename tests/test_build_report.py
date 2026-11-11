from datetime import datetime
from pathlib import Path

import pandas as pd

from src.build_report import best_match, excel_date, load_forecast, load_incurred, normalize_name, number, period_label


def test_excel_date_for_november_2026():
    assert excel_date(46327) == datetime(2026, 11, 1)


def test_normalize_name_removes_accents():
    assert normalize_name("Adrián Juguera") == "adrian juguera"


def test_best_match_accepts_abbreviated_name():
    match, score = best_match("Adrián Juguera", ["Adrián Juguera Aquilino", "Alberto Moliner Castelló"])
    assert match == "Adrián Juguera Aquilino"
    assert score >= 0.95


def test_best_match_understands_common_name_abbreviations():
    match, _ = best_match("Juanjo Payá", ["Juan Jose Paya Soler", "Jose Juan Jiménez Rodríguez"])
    assert match == "Juan Jose Paya Soler"


def test_number_parses_spanish_decimal():
    assert number("9,33") == 9.33


def test_spanish_period_label():
    assert period_label(datetime(2026, 11, 1)) == "Noviembre 2026"


def test_forecast_adjustment_is_aligned_by_position(monkeypatch):
    frame = pd.DataFrame(index=range(40), columns=range(31))
    frame.iat[2, 3] = 46327   # noviembre, forecast anterior
    frame.iat[2, 11] = 46327  # noviembre, forecast actual
    frame.iat[2, 18] = 46296  # octubre mal rotulado en Ajustes
    frame.iat[2, 19] = 46327
    frame.iat[2, 25] = 46327  # primera posición del bloque Diferencia
    frame.iat[3, 0] = "PROYECTO"
    frame.iat[3, 1] = "EXT-1"
    frame.iat[3, 2] = "Net Revenue"
    frame.iat[3, 3] = 100
    frame.iat[3, 11] = 200
    frame.iat[3, 18] = 36
    frame.iat[3, 19] = 28
    frame.iat[3, 25] = 100
    frame.iat[33, 1] = "Total"
    frame.iat[33, 2] = "Net Revenue"
    frame.iat[33, 3] = 100
    frame.iat[33, 11] = 200
    frame.iat[33, 18] = 36
    frame.iat[33, 19] = 28
    frame.iat[33, 25] = 100
    monkeypatch.setattr(pd, "read_excel", lambda *args, **kwargs: frame)

    result = load_forecast(Path("GECO.xlsb"), datetime(2026, 11, 1))

    assert result["totals"]["adjustment"] == 36
    assert result["totals"]["without_adjustment"] == 164
    assert result["financials"]["revenue"]["with_adjustment"] == 200


def test_incurred_keeps_client_daily_hours(tmp_path):
    report = tmp_path / "202611-Incurrido.html"
    report.write_text(
        "<table><tr><th>Usuario</th><th>Jira</th><th>01</th><th>02</th><th>Total validado</th></tr>"
        "<tr><td>Ana</td><td>CLIENTE</td><td>8</td><td>9,33</td><td>17,33</td></tr>"
        "<tr><td>Ana</td><td>NTT</td><td>1</td><td>V</td><td>1</td></tr></table>",
        encoding="utf-8",
    )

    records, events, codes = load_incurred(report, datetime(2026, 11, 1))

    assert records[0]["client_hours"] == 17.33
    assert records[0]["client_daily"] == {1: 8.0, 2: 9.33}
    assert records[0]["ntt_daily"] == {1: 1.0}
    assert events[0]["code"] == "V"
    assert codes == {"V"}
