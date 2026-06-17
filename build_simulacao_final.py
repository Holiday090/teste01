from __future__ import annotations

import argparse
import re
import unicodedata
from collections import OrderedDict
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter


DEFAULT_TEMPLATE = "ficheiro template 2.xlsx"
DEFAULT_SIMULATION = "S24 - Simulação PVP S24-2026.xlsm"
DEFAULT_COMPARAVEL = "20260609-Relatorio comparavel_.xlsx"
DEFAULT_TOTAL_MEAS = "TOTAL - meas a 09-06-2026.XLSX"
DEFAULT_OUTPUT = "simulacao-final.xlsx"
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
FORMULA_TEMPLATE_ROW = 4
FORMULA_COLUMNS = (25, 26, 27, 28)  # Y, Z, AA, AB

COMPETITOR_TO_FINAL_COLUMN = {
    "CONTINENTE": 19,  # S
    "LIDL": 20,  # T
    "PINGO-DOCE": 21,  # U
}
DADOS_HEADER_ROW = 2


def build_header_map(row: tuple[Any, ...]) -> dict[str, int]:
    return {as_text(value).upper(): index for index, value in enumerate(row) if as_text(value)}


def required_column(headers: dict[str, int], name: str) -> int:
    key = name.strip().upper()
    if key not in headers:
        raise ValueError(f"Header not found: {name}")
    return headers[key]


def normalize_filename(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value.casefold())
    return "".join(char for char in normalized if not unicodedata.combining(char))


def is_excel_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS


def latest_matching_file(path: Path, required_terms: tuple[str, ...], label: str) -> Path:
    normalized_terms = tuple(normalize_filename(term) for term in required_terms)
    candidates = [
        candidate
        for candidate in path.iterdir()
        if is_excel_file(candidate)
        and all(term in normalize_filename(candidate.name) for term in normalized_terms)
    ]
    if not candidates:
        raise FileNotFoundError(f"Ficheiro não encontrado: {label}")
    return max(candidates, key=lambda candidate: (candidate.stat().st_mtime_ns, candidate.name))


def find_latest_comparavel(path: Path) -> Path:
    return latest_matching_file(path, ("relatorio", "comparavel"), "relatório comparável")


def find_latest_total_meas(path: Path) -> Path:
    return latest_matching_file(path, ("total", "meas"), "TOTAL MEAS")


def resolve_input_path(root: Path, value: str | None, latest_file) -> Path:
    if value is None:
        return latest_file(root)
    path = Path(value)
    if not path.is_absolute():
        path = root / path
    if not path.exists():
        raise FileNotFoundError(path)
    return path


def as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def item_key(value: Any) -> str:
    text = as_text(value)
    return text.lstrip("0") or text


def ean_key(value: Any) -> str:
    text = as_text(value)
    return text.lstrip("0") or text


def excluded_article(value: Any) -> bool:
    description = as_text(value).upper()
    return description.startswith(("SUB.", "PAL"))


def first_non_empty(current: Any, candidate: Any) -> Any:
    if current not in (None, ""):
        return current
    return candidate if candidate is not None else ""


def parse_date(value: Any) -> datetime | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return datetime(value.year, value.month, value.day)
    text = as_text(value)
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def format_slash_date(value: Any) -> str:
    parsed = parse_date(value)
    if parsed is None:
        return as_text(value)
    return parsed.strftime("%d/%m/%Y")


def format_dash_date(value: Any) -> str:
    parsed = parse_date(value)
    if parsed is None:
        return as_text(value)
    return parsed.strftime("%d-%m-%Y")


def load_simulation(simulation_path: Path) -> tuple[list[dict[str, Any]], datetime | None]:
    wb = load_workbook(simulation_path, data_only=True, read_only=True, keep_vba=True)
    dados_ws = wb["Dados"]
    shopping_date = parse_date(dados_ws["G1"].value) or parse_date(dados_ws["H1"].value)
    headers = build_header_map(next(dados_ws.iter_rows(min_row=DADOS_HEADER_ROW, max_row=DADOS_HEADER_ROW, values_only=True)))
    cols = {
        "itm8": required_column(headers, "ITM8"),
        "ean": required_column(headers, "EAN"),
        "description": required_column(headers, "Descrição"),
        "brand": required_column(headers, "Marca"),
        "pvp_competitor": required_column(headers, "PVP Concorrente"),
        "pvp_current": required_column(headers, "PVP Cadencier Actual"),
        "pvp_future": required_column(headers, "PVP Cadencier Futuro"),
        "competitor": required_column(headers, "Insignia Concorrente"),
        "psycho": required_column(headers, "Psyco"),
        "tipo": required_column(headers, "Tipo Produto"),
        "argus": required_column(headers, "Argus"),
        "aval": required_column(headers, "Aval"),
        "amont": required_column(headers, "Amont"),
        "st": required_column(headers, "Estatuto"),
        "promo_perm": required_column(headers, "Promo Permanente"),
        "edlp": required_column(headers, "EDLP"),
    }

    # Rebuild the TD Psyco pivot without the visual filters saved in Excel.
    # Final R must match TD Psyco L (PVP Cadencier Actual), and final S:T:U
    # must match TD Psyco N:O:P (CONTINENTE, LIDL, PINGO-DOCE). Mercadona is
    # excluded because TD Psyco does not expose it in the final Shopping columns.
    records: OrderedDict[tuple[Any, ...], dict[str, Any]] = OrderedDict()
    price_totals: dict[tuple[Any, ...], dict[str, list[float]]] = {}
    for row in dados_ws.iter_rows(min_row=DADOS_HEADER_ROW + 1, values_only=True):
        itm8 = as_text(row[cols["itm8"]])
        ean = as_text(row[cols["ean"]])
        if not itm8 and not ean:
            continue
        if excluded_article(row[cols["description"]]):
            continue

        competitor = as_text(row[cols["competitor"]]).upper()
        if competitor not in COMPETITOR_TO_FINAL_COLUMN:
            continue

        key = (
            row[cols["amont"]],
            row[cols["aval"]],
            itm8,
            ean,
            row[cols["description"]],
            row[cols["brand"]],
            row[cols["tipo"]],
            row[cols["psycho"]],
            row[cols["argus"]],
            row[cols["promo_perm"]],
            row[cols["edlp"]],
            row[cols["pvp_current"]],
            row[cols["pvp_future"]],
        )
        if key not in records:
            records[key] = {
                "amont": row[cols["amont"]],
                "aval": row[cols["aval"]],
                "itm8": itm8,
                "description": row[cols["description"]],
                "brand": row[cols["brand"]],
                "ean": ean,
                "tipo": row[cols["tipo"]],
                "psycho": row[cols["psycho"]],
                "promo_perm": row[cols["promo_perm"]],
                "edlp": row[cols["edlp"]],
                "argus": row[cols["argus"]],
                "st": row[cols["st"]],  # Estatuto always comes from Dados.
                "pvp_cadencier": row[cols["pvp_current"]],
                "prices": {},
            }
            price_totals[key] = {}

        price = as_number(row[cols["pvp_competitor"]])
        if price is not None:
            price_totals[key].setdefault(competitor, []).append(price)

    for key, totals in price_totals.items():
        records[key]["prices"] = {
            competitor: sum(values) / len(values)
            for competitor, values in totals.items()
            if values
        }

    wb.close()

    sorted_records = sorted(
        records.values(),
        key=lambda record: (
            as_text(record["amont"])[1:3],
            as_text(record["amont"])[4:7],
            as_text(record["aval"])[4:6],
            as_text(record["aval"])[7:9],
            as_text(record["brand"]),
            as_number(record["pvp_cadencier"]) if as_number(record["pvp_cadencier"]) is not None else 999999999,
        ),
    )
    return sorted_records, shopping_date


def load_comparavel(comparavel_path: Path) -> dict[str, dict[str, Any]]:
    wb = load_workbook(comparavel_path, data_only=True, read_only=True)
    ws = wb["Produtos"]

    promos: dict[str, dict[str, Any]] = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        ean = ean_key(row[0])
        if not ean:
            continue
        values = {
            "CONTINENTE": row[17] if row[17] is not None else "",
            "LIDL": row[25] if row[25] is not None else "",
            "PINGO-DOCE": row[21] if row[21] is not None else "",
        }
        current = promos.setdefault(ean, {"CONTINENTE": "", "LIDL": "", "PINGO-DOCE": ""})
        for competitor, value in values.items():
            if value not in (None, ""):
                current[competitor] = value

    wb.close()
    return promos


def header_index(headers: tuple[Any, ...], name: str) -> int:
    normalized_name = name.strip().upper()
    for index, value in enumerate(headers):
        if as_text(value).upper() == normalized_name:
            return index
    raise ValueError(f"Header not found: {name}")


def comment_header_index(headers: tuple[Any, ...]) -> int:
    for index, value in enumerate(headers):
        label = as_text(value).upper()
        if "COMENT" in label and "SUIVI" in label:
            return index
    raise ValueError("Header not found: Comentarios (face ao suivi)")


def history_header_index(headers: tuple[Any, ...]) -> int:
    return header_index(headers, "HISTORICO")


def load_total_meas(total_meas_path: Path) -> dict[str, dict[str, Any]]:
    wb = load_workbook(total_meas_path, data_only=True, read_only=True)
    ws = wb["sql_query3"]

    rows = ws.iter_rows(min_row=1, values_only=True)
    headers = next(rows)
    uvc_col = header_index(headers, "UVC")
    ean_col = header_index(headers, "EAN")
    in_mea_col = header_index(headers, "IN_MEA")
    pvp_col = header_index(headers, "PVP")

    by_uvc_ean: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        uvc = as_text(row[uvc_col])
        ean = as_text(row[ean_col])
        date = parse_date(row[in_mea_col])
        if not uvc or not ean or date is None:
            continue

        key = (uvc, ean)
        current = by_uvc_ean.get(key)
        if current is None or date < current["sort_date"]:
            by_uvc_ean[key] = {
                "sort_date": date,
                "date": date,
                "pvp": row[pvp_col] if row[pvp_col] is not None else "",
            }

    # The final template has EAN but not UVC. We first respect UVC+EAN to find the
    # first MEAS record per UVC, then collapse to the earliest occurrence per EAN.
    by_ean: dict[str, dict[str, Any]] = {}
    for (_uvc, ean), value in by_uvc_ean.items():
        current = by_ean.get(ean)
        if current is None or value["sort_date"] < current["sort_date"]:
            by_ean[ean] = value

    wb.close()
    return by_ean


def extract_week_number(filename: str) -> int | None:
    match = re.search(r"(?i)(?:^|[^A-Z0-9])S(\d{1,2})(?:[^0-9]|$)", filename)
    return int(match.group(1)) if match else None


def find_previous_week(path: Path, simulation_path: Path) -> Path | None:
    current_week = extract_week_number(simulation_path.name)
    if current_week is None or current_week <= 1:
        return None

    previous_week_label = f"S{current_week - 1}"
    candidates = sorted(
        candidate
        for candidate in path.glob(f"*{previous_week_label}*.xls*")
        if candidate.resolve() != simulation_path.resolve()
    )
    return candidates[0] if candidates else None


def load_previous_comments(previous_path: Path | None) -> dict[str, dict[str, Any]]:
    if previous_path is None or not previous_path.exists():
        return {"itm8": {}, "ean": {}}

    wb = load_workbook(previous_path, data_only=True, read_only=True)
    ws = wb["Folha1"]

    header_row = None
    headers: tuple[Any, ...] | None = None
    for row_number, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        labels = [as_text(value).upper() for value in row]
        if "ITM8" in labels and any("COMENT" in label and "SUIVI" in label for label in labels):
            header_row = row_number
            headers = row
            break

    if header_row is None or headers is None:
        wb.close()
        raise ValueError(f"Could not find ITM8/comments headers in {previous_path}")

    itm8_col = next(index for index, value in enumerate(headers) if as_text(value).upper() == "ITM8")
    ean_col = next((index for index, value in enumerate(headers) if as_text(value).upper() == "EAN"), None)
    comments_col = comment_header_index(headers)

    comments: dict[str, dict[str, Any]] = {"itm8": {}, "ean": {}}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        comment = row[comments_col] if row[comments_col] is not None else ""
        itm8 = item_key(row[itm8_col])
        if itm8 and itm8 not in comments["itm8"]:
            comments["itm8"][itm8] = comment
        if ean_col is not None:
            ean = ean_key(row[ean_col])
            if ean and ean not in comments["ean"]:
                comments["ean"][ean] = comment

    wb.close()
    return comments


def previous_comment_for(record: dict[str, Any], comments: dict[str, dict[str, Any]]) -> Any:
    by_itm8 = comments.get("itm8", {})
    by_ean = comments.get("ean", {})
    itm8 = item_key(record["itm8"])
    ean = ean_key(record["ean"])
    return by_itm8.get(itm8, by_ean.get(ean, ""))


def copy_cell_style(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.font:
        target.font = copy(source.font)
    if source.fill:
        target.fill = copy(source.fill)
    if source.border:
        target.border = copy(source.border)
    if source.alignment:
        target.alignment = copy(source.alignment)
    if source.protection:
        target.protection = copy(source.protection)


def prepare_template(template_path: Path):
    wb = load_workbook(template_path)
    ws = wb["Folha1"]

    ws["R3"] = "PVP Cadencier"

    return wb, ws


def capture_formula_templates(ws, columns: tuple[int, ...]) -> dict[int, Any]:
    return {col: ws.cell(FORMULA_TEMPLATE_ROW, col).value for col in columns}


def translated_formula(formula: Any, col: int, row_number: int) -> Any:
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    origin = f"{get_column_letter(col)}{FORMULA_TEMPLATE_ROW}"
    target = f"{get_column_letter(col)}{row_number}"
    return Translator(formula, origin=origin).translate_formula(target)


def clear_data_area(ws, max_rows: int) -> None:
    last_row = max(ws.max_row, max_rows)
    for row in ws.iter_rows(min_row=4, max_row=last_row, max_col=35):
        for cell in row:
            cell.value = None


def as_number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = as_text(value).replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def numeric_values(values: list[Any]) -> list[float]:
    return [number for value in values if (number := as_number(value)) is not None]


def excel_mode(values: list[float]) -> float | None:
    counts: dict[float, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    repeated = [value for value in values if counts[value] > 1]
    return repeated[0] if repeated else None


def numbers_equal(left: Any, right: Any) -> bool:
    left_number = as_number(left)
    right_number = as_number(right)
    if left_number is None or right_number is None:
        return left == right
    return abs(left_number - right_number) < 0.0000001


def calculate_condition_price(record: dict[str, Any], shopping_values: list[Any], promo_values: list[Any]) -> float | None:
    shopping_numbers = numeric_values(shopping_values)
    promo_numbers = numeric_values(promo_values)
    if not shopping_numbers and not promo_numbers:
        return None

    promo_perm_or_edlp = as_text(record["promo_perm"]).upper() == "O" or as_text(record["edlp"]).upper() == "O"
    if promo_perm_or_edlp:
        return min(promo_numbers) if promo_numbers else min(shopping_numbers)

    if as_text(record["psycho"]).upper() == "O":
        return min(shopping_numbers) if shopping_numbers else None

    if as_text(record["tipo"]).upper() == "E":
        mode_value = excel_mode(shopping_numbers)
        return mode_value if mode_value is not None else min(shopping_numbers)

    return sum(shopping_numbers) / len(shopping_numbers) if shopping_numbers else None


def fill_row(
    ws,
    row_number: int,
    record: dict[str, Any],
    promos: dict[str, Any],
    meas: dict[str, Any],
    previous_comment: Any,
    group_map: dict[str, Any],
    history_col: int,
    formula_templates: dict[int, Any],
) -> None:
    base_values = {
        1: record["amont"],
        2: record["aval"],
        8: record["itm8"],
        9: record["description"],
        10: record["brand"],
        11: record["ean"],
        12: record["tipo"],
        13: record["psycho"],
        14: record["promo_perm"],
        15: record["edlp"],
        16: record["argus"],
        17: record["st"],
        18: record["pvp_cadencier"],
    }
    for col, value in base_values.items():
        ws.cell(row_number, col).value = value

    prices = record["prices"]
    ws.cell(row_number, 19).value = prices.get("CONTINENTE")
    ws.cell(row_number, 20).value = prices.get("LIDL")
    ws.cell(row_number, 21).value = prices.get("PINGO-DOCE")

    ws.cell(row_number, 22).value = promos.get("CONTINENTE", "")
    ws.cell(row_number, 23).value = promos.get("LIDL", "")
    ws.cell(row_number, 24).value = promos.get("PINGO-DOCE", "")

    campaign_date_cell = ws.cell(row_number, 29)
    campaign_date_cell.value = meas.get("date", "")
    if campaign_date_cell.value not in (None, ""):
        campaign_date_cell.number_format = "dd-mm-yyyy"
    ws.cell(row_number, 30).value = meas.get("pvp", "")
    ws.cell(row_number, history_col).value = previous_comment if previous_comment is not None else ""

    amont = as_text(record["amont"])
    aval = as_text(record["aval"])
    grp = amont[1:3] if len(amont) >= 3 else ""
    ws.cell(row_number, 3).value = grp
    ws.cell(row_number, 4).value = group_map.get(grp, "")
    ws.cell(row_number, 5).value = amont[4:7] if len(amont) >= 7 else ""
    ws.cell(row_number, 6).value = aval[4:6] if len(aval) >= 6 else ""
    ws.cell(row_number, 7).value = aval[7:9] if len(aval) >= 9 else ""

    for col in FORMULA_COLUMNS:
        ws.cell(row_number, col).value = translated_formula(formula_templates.get(col), col, row_number)

    ws.cell(row_number, 31).value = f'=IF(AC{row_number}="","",IF(AC{row_number}>=(TODAY()+15),"não","sim"))'


def apply_row_styles(ws, rows_count: int) -> None:
    source_cells = [ws.cell(4, col) for col in range(1, 36)]
    for row_number in range(4, rows_count + 4):
        for col in range(1, 36):
            copy_cell_style(source_cells[col - 1], ws.cell(row_number, col))


def build_workbook(
    template_path: Path,
    simulation_path: Path,
    comparavel_path: Path,
    total_meas_path: Path,
    previous_week_path: Path | None,
    output_path: Path,
) -> None:
    records, shopping_date = load_simulation(simulation_path)
    comparavel = load_comparavel(comparavel_path)
    total_meas = load_total_meas(total_meas_path)
    previous_comments = load_previous_comments(previous_week_path)

    wb, ws = prepare_template(template_path)
    formula_templates = capture_formula_templates(ws, FORMULA_COLUMNS)
    group_map = {
        as_text(row[0]): row[1]
        for row in wb["Folha2"].iter_rows(min_row=2, values_only=True)
        if as_text(row[0])
    }
    clear_data_area(ws, len(records) + 3)
    apply_row_styles(ws, len(records))
    headers = tuple(ws.cell(3, col).value for col in range(1, ws.max_column + 1))
    history_col = history_header_index(headers) + 1

    date_label = format_slash_date(shopping_date)
    if date_label:
        ws["S2"] = f"Shopping {date_label}"
        ws["V2"] = f"PROMO {date_label}"

    for offset, record in enumerate(records, start=4):
        ean = ean_key(record["ean"])
        fill_row(
            ws,
            offset,
            record,
            comparavel.get(ean, {}),
            total_meas.get(ean, {}),
            previous_comment_for(record, previous_comments),
            group_map,
            history_col,
            formula_templates,
        )

    first_extra_row = len(records) + 4
    if ws.max_row >= first_extra_row:
        ws.delete_rows(first_extra_row, ws.max_row - first_extra_row + 1)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera a simulação final a partir do template e ficheiros de origem.")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--simulation", default=DEFAULT_SIMULATION)
    parser.add_argument("--comparavel", default=None, help="Opcional: por defeito usa o relatório comparável mais recente.")
    parser.add_argument("--total-meas", default=None, help="Opcional: por defeito usa o TOTAL MEAS mais recente.")
    parser.add_argument("--previous-week", default=None, help="Ficheiro da semana anterior para copiar Comentarios (face ao suivi).")
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    root = Path.cwd()
    simulation_path = root / args.simulation
    comparavel_path = resolve_input_path(root, args.comparavel, find_latest_comparavel)
    total_meas_path = resolve_input_path(root, args.total_meas, find_latest_total_meas)
    previous_week = Path(args.previous_week) if args.previous_week else find_previous_week(root, simulation_path)
    build_workbook(
        root / args.template,
        simulation_path,
        comparavel_path,
        total_meas_path,
        previous_week,
        root / args.output,
    )
    print(f"Ficheiro criado: {args.output}")
    if previous_week is None:
        print("Aviso: ficheiro da semana anterior não encontrado; comentários face ao suivi ficaram vazios.")


if __name__ == "__main__":
    main()
