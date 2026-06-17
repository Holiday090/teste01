from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    from build_simulacao_final import (
        DEFAULT_SIMULATION,
        DEFAULT_TEMPLATE,
        build_workbook,
        find_latest_comparavel,
        find_latest_total_meas,
        find_previous_week,
        resolve_input_path,
    )
except ModuleNotFoundError as error:
    if error.name == "openpyxl":
        print("Falta instalar a dependencia openpyxl.")
        print("Corre primeiro: python -m pip install -r requirements.txt")
        sys.exit(1)
    raise


DEFAULT_PC_OUTPUT = "simulacao-final-novo.xlsx"


def existing_file(root: Path, filename: str, label: str) -> Path:
    path = root / filename
    if not path.exists():
        raise FileNotFoundError(f"Ficheiro em falta ({label}): {path}")
    return path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Cria o ficheiro Excel final a partir do template e ficheiros de origem.",
        epilog=(
            "Exemplo: python gerar_ficheiro_final_pc.py\n"
            "Antes de correr, instala dependencias com: python -m pip install -r requirements.txt\n"
            "Sem argumentos, usa o relatorio comparavel e o TOTAL MEAS mais recentes da pasta.\n"
            "Os comentarios da semana anterior sao copiados para a coluna HISTORICO."
        ),
    )
    parser.add_argument("--folder", default=".", help="Pasta onde estao os ficheiros Excel e scripts.")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--simulation", default=DEFAULT_SIMULATION)
    parser.add_argument("--comparavel", default=None, help="Opcional: por defeito usa o relatorio comparavel mais recente.")
    parser.add_argument("--total-meas", default=None, help="Opcional: por defeito usa o TOTAL MEAS mais recente.")
    parser.add_argument("--previous-week", default=None, help="Opcional: ficheiro da semana anterior.")
    parser.add_argument("--output", default=DEFAULT_PC_OUTPUT, help="Nome do ficheiro final a criar.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    root = Path(args.folder).expanduser().resolve()

    template_path = existing_file(root, args.template, "template")
    simulation_path = existing_file(root, args.simulation, "Simulacao PVP")
    comparavel_path = resolve_input_path(root, args.comparavel, find_latest_comparavel)
    total_meas_path = resolve_input_path(root, args.total_meas, find_latest_total_meas)
    output_path = root / args.output

    if args.previous_week:
        previous_week_path = existing_file(root, args.previous_week, "semana anterior")
    else:
        previous_week_path = find_previous_week(root, simulation_path)

    build_workbook(
        template_path,
        simulation_path,
        comparavel_path,
        total_meas_path,
        previous_week_path,
        output_path,
    )

    print(f"Ficheiro criado: {output_path}")
    if previous_week_path is None:
        print("Aviso: ficheiro da semana anterior nao encontrado; comentarios face ao suivi ficaram vazios.")
    else:
        print(f"Comentarios copiados a partir de: {previous_week_path}")


if __name__ == "__main__":
    main()
