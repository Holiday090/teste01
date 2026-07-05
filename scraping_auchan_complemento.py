"""
Complemento do scraping parcial Auchan — completa Scraping_Auchan.xlsx.

SETUP:
    pip install playwright pandas openpyxl
    python3 -m playwright install chromium

EXECUÇÃO:
    PYTHONUNBUFFERED=1 python3 scraping_auchan_complemento.py
    PYTHONUNBUFFERED=1 python3 scraping_auchan_complemento.py --test

NOTAS:
    - Mantém o ficheiro original Scraping_Auchan.xlsx (2224 produtos das 4 categorias).
    - Adiciona as secções de Alimentação em falta (Produtos Lácteos, Mercearia, etc.).
    - NÃO inclui Congelados, Charcutaria nem Queijaria (já no ficheiro original).
    - Deduplica por URL — produtos já no Excel não são reprocessados.
    - Progresso separado: scraping_auchan_complemento_progress.json
    - Commit/push Git a cada 100 produtos (progresso + Excel).
    - Script completo do site: scraping_auchan_completo.py → Scraping_Auchan_Completo.xlsx
"""

from __future__ import annotations

import argparse
import json
import subprocess
import time
from pathlib import Path

import pandas as pd
import scraping_auchan as auchan
from playwright.sync_api import sync_playwright

OUTPUT_FILE = "Scraping_Auchan.xlsx"
PROGRESS_FILE = "scraping_auchan_complemento_progress.json"
REPO_ROOT = Path(__file__).resolve().parent
GIT_PUSH_BACKOFF_SECONDS = (4, 8, 16, 32)
_git_checkpoint_enabled = True

# Categorias já extraídas pelo scraping_auchan.py original — não repetir.
ORIGINAL_COMPLETED_URLS: tuple[str, ...] = (
    "/pt/produtos-frescos/charcutaria/",
    "/pt/produtos-frescos/queijaria/",
    "/pt/alimentacao/congelados/",
)

# Secções de Alimentação em falta no ficheiro original (sem Congelados).
COMPLEMENT_LISTING_URLS: tuple[str, ...] = (
    "/pt/alimentacao/",
    "/pt/alimentacao/produtos-lacteos/",
    "/pt/alimentacao/produtos-lacteos/bebidas-de-cafe-refrigeradas/",
    "/pt/alimentacao/produtos-lacteos/bebidas-vegetais/",
    "/pt/alimentacao/produtos-lacteos/gelatinas-e-sobremesas/",
    "/pt/alimentacao/produtos-lacteos/iogurtes/",
    "/pt/alimentacao/produtos-lacteos/leites/",
    "/pt/alimentacao/produtos-lacteos/manteiga-cremes-e-margarina/",
    "/pt/alimentacao/produtos-lacteos/natas-bechamel-e-chantilly/",
    "/pt/alimentacao/produtos-lacteos/ovos/",
    "/pt/alimentacao/produtos-lacteos/proteina/",
    "/pt/alimentacao/produtos-lacteos/queijaria/",
    "/pt/alimentacao/produtos-lacteos/sem-lactose/",
    "/pt/alimentacao/mercearia/",
    "/pt/alimentacao/mercearia/acucar-e-adocante/",
    "/pt/alimentacao/mercearia/arroz-e-massa/",
    "/pt/alimentacao/mercearia/azeite-oleo-e-vinagre/",
    "/pt/alimentacao/mercearia/batatas-fritas-e-aperitivos-snacks/",
    "/pt/alimentacao/mercearia/bolachas-e-bolos/",
    "/pt/alimentacao/mercearia/cafe-cha-e-infusao/",
    "/pt/alimentacao/mercearia/cereais-e-barras/",
    "/pt/alimentacao/mercearia/chocolates-e-achocolatados/",
    "/pt/alimentacao/mercearia/conservas/",
    "/pt/alimentacao/mercearia/cremes-compotas-e-mel/",
    "/pt/alimentacao/mercearia/farinha/",
    "/pt/alimentacao/mercearia/leite-condensado-e-preparado-para-bolos/",
    "/pt/alimentacao/mercearia/maionese-ketchup-mostarda-e-especialidades/",
    "/pt/alimentacao/mercearia/pastilhas-rebucados-e-chupas/",
    "/pt/alimentacao/mercearia/polpas-caldos-e-temperos/",
    "/pt/alimentacao/mercearia/refeicoes-sopas-e-pure/",
    "/pt/alimentacao/mercearia/sal-ervas-e-temperos/",
    "/pt/alimentacao/mercearia/tostas-e-pao-embalado/",
    "/pt/alimentacao/sabores-do-mundo/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-africana/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-america-do-sul/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-america-norte/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-asiatica/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-brasileira/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-europa-do-leste/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-europeia/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-indiana/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-medio-oriente/",
    "/pt/alimentacao/sabores-do-mundo/cozinha-mexicana/",
    "/pt/alimentacao/sabores-do-mundo/halal/",
    "/pt/alimentacao/future-taste/",
    "/pt/alimentacao/future-taste/algas/",
    "/pt/alimentacao/future-taste/economia-circular/",
    "/pt/alimentacao/future-taste/funcionais/",
    "/pt/alimentacao/future-taste/insectos/",
    "/pt/alimentacao/future-taste/refeicoes/",
    "/pt/alimentacao/future-taste/snacks/",
    "/pt/alimentacao/future-taste/veggie/",
)

auchan.PROGRESS_FILE = PROGRESS_FILE


def log(message: str) -> None:
    auchan.log(message)


def listing_url_to_display_name(listing_url: str) -> str:
    segments = [segment for segment in listing_url.strip("/").split("/") if segment and segment != "pt"]
    labels = [segment.replace("-", " ").title() for segment in segments]
    if labels and labels[0].lower() == "alimentacao":
        labels = labels[1:]
    if not labels:
        return "Alimentação"
    return " > ".join(labels)


def build_complement_targets() -> tuple[auchan.CategoryTarget, ...]:
    targets: list[auchan.CategoryTarget] = []
    for listing_url in COMPLEMENT_LISTING_URLS:
        targets.append(
            auchan.CategoryTarget(
                display_name=listing_url_to_display_name(listing_url),
                menu_path=("Alimentação",),
                listing_url=listing_url,
            )
        )
    return tuple(targets)


def category_key(listing_url: str) -> str:
    path = listing_url.rstrip("/") + "/"
    if not path.startswith("/"):
        path = "/" + path
    return path


def empty_progress(output_file: str = OUTPUT_FILE) -> dict:
    return {
        "output_file": output_file,
        "completed_categories": list(ORIGINAL_COMPLETED_URLS),
        "records": [],
        "processed_urls": [],
        "in_progress": None,
    }


def load_progress(output_file: str = OUTPUT_FILE) -> dict:
    progress_path = Path(PROGRESS_FILE)
    if not progress_path.exists():
        return empty_progress(output_file)

    with progress_path.open("r", encoding="utf-8") as progress_file:
        progress = json.load(progress_file)

    defaults = empty_progress(output_file)
    for key, default_value in defaults.items():
        if key not in progress:
            progress[key] = default_value

    completed = set(progress.get("completed_categories", []))
    completed.update(ORIGINAL_COMPLETED_URLS)
    progress["completed_categories"] = sorted(completed)
    return progress


def save_progress(progress: dict) -> None:
    with Path(PROGRESS_FILE).open("w", encoding="utf-8") as progress_file:
        json.dump(progress, progress_file, ensure_ascii=False, indent=2)


def commit_checkpoint_to_repo(product_count: int) -> bool:
    if not _git_checkpoint_enabled:
        return False

    files_to_add: list[str] = []
    if (REPO_ROOT / PROGRESS_FILE).exists():
        files_to_add.append(PROGRESS_FILE)
    if (REPO_ROOT / OUTPUT_FILE).exists():
        files_to_add.append(OUTPUT_FILE)

    if not files_to_add:
        return False

    try:
        subprocess.run(
            ["git", "add", "--", *files_to_add],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=True,
        )
        commit_result = subprocess.run(
            ["git", "commit", "-m", f"Checkpoint Auchan complemento: {product_count} produtos"],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        if commit_result.returncode != 0:
            combined = f"{commit_result.stdout}\n{commit_result.stderr}".lower()
            if "nothing to commit" in combined:
                return False
            log(f"[Git] git commit falhou: {commit_result.stderr.strip()}")
            return False

        branch_name = subprocess.check_output(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"],
            cwd=REPO_ROOT,
            text=True,
        ).strip()

        for attempt, wait_seconds in enumerate(GIT_PUSH_BACKOFF_SECONDS, start=1):
            push_result = subprocess.run(
                ["git", "push", "-u", "origin", branch_name],
                cwd=REPO_ROOT,
                capture_output=True,
                text=True,
                check=False,
            )
            if push_result.returncode == 0:
                log(
                    f"[Git] Complemento enviado para origin/{branch_name} "
                    f"({product_count} produtos totais no ficheiro)."
                )
                return True
            if attempt == len(GIT_PUSH_BACKOFF_SECONDS):
                log(f"[Git] push falhou: {push_result.stderr.strip()}")
                return False
            time.sleep(wait_seconds)
    except Exception as error:
        log(f"[Git] Erro: {error}")
        return False
    return False


def save_progress_with_git_if_needed(progress: dict) -> None:
    save_progress(progress)
    if not _git_checkpoint_enabled:
        return
    in_progress = progress.get("in_progress") or {}
    if in_progress.get("phase") != "collecting":
        return
    partial_count = len(in_progress.get("partial_products", []))
    if partial_count > 0 and partial_count % auchan.CHECKPOINT_EVERY == 0:
        commit_checkpoint_to_repo(len(progress.get("records", [])))


auchan.save_progress = save_progress_with_git_if_needed


def maybe_checkpoint(
    records: list[dict[str, str]],
    output_file: str,
    progress: dict,
    *,
    use_git: bool = True,
) -> None:
    if not records or len(records) % auchan.CHECKPOINT_EVERY != 0:
        return

    auchan.export_to_excel(records, output_file)
    save_progress(progress)
    log(
        f"[Checkpoint] {len(records)} produtos em '{output_file}' "
        f"(progresso: '{PROGRESS_FILE}')."
    )
    if use_git:
        commit_checkpoint_to_repo(len(records))


def seed_progress_from_excel(progress: dict, output_file: str) -> None:
    """Carrega os 2224 produtos existentes do Excel para não duplicar trabalho."""
    if progress.get("records"):
        return

    excel_path = Path(output_file)
    if not excel_path.exists():
        log(f"Excel '{output_file}' não encontrado — a iniciar complemento do zero.")
        return

    dataframe = pd.read_excel(excel_path)
    if dataframe.empty:
        return

    missing_columns = [column for column in auchan.EXCEL_COLUMNS if column not in dataframe.columns]
    if missing_columns:
        raise SystemExit(f"Colunas em falta no Excel existente: {missing_columns}")

    records = dataframe[auchan.EXCEL_COLUMNS].astype(str).to_dict(orient="records")
    processed_urls = [
        url for url in dataframe["URL"].dropna().astype(str).tolist() if url and url != "N/A"
    ]

    progress["records"] = records
    progress["processed_urls"] = sorted(set(processed_urls))
    save_progress(progress)
    log(
        f"Excel existente carregado — {len(records)} produtos, "
        f"{len(processed_urls)} URLs únicas. A adicionar secções em falta."
    )


def process_category(
    page,
    target: auchan.CategoryTarget,
    progress: dict,
    output_file: str,
    *,
    max_products: int | None = None,
    max_pages: int | None = None,
    use_progress: bool = True,
    use_git: bool = True,
) -> None:
    listing_key = category_key(target.listing_url)
    category_name = target.display_name

    if use_progress and listing_key in progress.get("completed_categories", []):
        log(f"[Categoria: {category_name}] Já concluída — a saltar.")
        return

    log(f"\n{'=' * 70}\n[Categoria: {category_name}]\nURL: {target.listing_url}\n{'=' * 70}")

    progress_state = progress if use_progress else None
    products = auchan.collect_all_category_products(
        page,
        target,
        navigate_via_menu_flag=False,
        max_products=max_products,
        max_pages=max_pages,
        progress=progress_state,
    )

    if not products:
        if use_progress:
            completed = list(progress.get("completed_categories", []))
            if listing_key not in completed:
                completed.append(listing_key)
            progress["completed_categories"] = completed
            progress["in_progress"] = None
            save_progress(progress)
            if use_git:
                commit_checkpoint_to_repo(len(progress.get("records", [])))
        return

    records: list[dict[str, str]] = progress.setdefault("records", [])
    processed_urls = set(progress.get("processed_urls", [])) if use_progress else set()
    pending_products = [product for product in products if product.get("url") not in processed_urls]
    skipped = len(products) - len(pending_products)

    if skipped:
        log(f"[Categoria: {category_name}] {skipped} produto(s) já no Excel — a saltar.")

    log(f"[Categoria: {category_name}] Fase 2 — {len(pending_products)} produtos pendentes.")
    total = len(pending_products)

    for index, product in enumerate(pending_products, start=1):
        product_url = product.get("url", "")
        auchan.log_product_progress(
            category_name,
            index,
            total,
            status="processing",
            product_name=str(product.get("name", "")).strip(),
        )
        try:
            product_data = auchan.extract_product_details(page, product, category_name)
            records.append(product_data)
            if use_progress and product_url:
                processed_urls.add(product_url)
                progress["records"] = records
                progress["processed_urls"] = sorted(processed_urls)
                save_progress(progress)
                maybe_checkpoint(records, output_file, progress, use_git=use_git)
            auchan.log_product_progress(
                category_name,
                index,
                total,
                status="done",
                product_name=product_data["Descrição do Produto"],
                details=f"{product_data['Caminho Categorias']} | EAN: {product_data['EAN / Referência']}",
            )
        except Exception as error:
            auchan.log_product_progress(category_name, index, total, status="error", details=str(error))
            records.append(auchan.empty_product_record(product, category_name))
            if use_progress and product_url:
                processed_urls.add(product_url)
                progress["records"] = records
                progress["processed_urls"] = sorted(processed_urls)
                save_progress(progress)
                maybe_checkpoint(records, output_file, progress, use_git=use_git)

        if index < total:
            auchan.human_delay(0.8, 1.8)

    if use_progress:
        completed = list(progress.get("completed_categories", []))
        if listing_key not in completed:
            completed.append(listing_key)
        progress["completed_categories"] = completed
        progress["in_progress"] = None
        save_progress(progress)
        auchan.export_to_excel(records, output_file)
        if use_git:
            commit_checkpoint_to_repo(len(records))

    log(f"[Categoria: {category_name}] Concluída — {total} novos produtos processados.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Complementa Scraping_Auchan.xlsx com secções de Alimentação em falta."
    )
    parser.add_argument("--test", action="store_true", help="2 categorias, 2 produtos cada.")
    parser.add_argument("--max-categories", type=int, help="Limita categorias a processar.")
    parser.add_argument("--max-pages", type=int, help="Limita páginas por categoria.")
    parser.add_argument("--max-products", type=int, help="Limita produtos por categoria.")
    parser.add_argument("--reset-progress", action="store_true", help="Apaga progresso do complemento.")
    parser.add_argument("--no-git", action="store_true", help="Desativa commit/push automático.")
    parser.add_argument("--output", default=OUTPUT_FILE, help=f"Excel de saída (predefinido: {OUTPUT_FILE}).")
    return parser.parse_args()


def main() -> None:
    global _git_checkpoint_enabled

    args = parse_args()
    use_progress = not args.test and not args.max_pages and not args.max_products and not args.max_categories
    use_git = use_progress and not args.no_git
    _git_checkpoint_enabled = use_git

    max_products = args.max_products if args.max_products else (2 if args.test else None)
    max_pages = args.max_pages
    max_categories = args.max_categories if args.max_categories else (2 if args.test else None)

    if args.reset_progress and Path(PROGRESS_FILE).exists():
        Path(PROGRESS_FILE).unlink()
        log(f"Progresso do complemento removido: {PROGRESS_FILE}")

    progress = load_progress(args.output) if use_progress else empty_progress(args.output)
    progress["output_file"] = args.output
    seed_progress_from_excel(progress, args.output)

    targets = build_complement_targets()
    completed = set(progress.get("completed_categories", []))
    pending = [target for target in targets if category_key(target.listing_url) not in completed]

    if max_categories:
        pending = pending[:max_categories]

    log(f"Complemento Scraping_Auchan.xlsx — {len(pending)} categorias pendentes (de {len(targets)}).")
    log(f"Produtos actuais no ficheiro: {len(progress.get('records', []))}")
    if use_git:
        log(f"Git: commit/push a cada {auchan.CHECKPOINT_EVERY} produtos.")

    if not pending:
        log("Todas as categorias de complemento já foram processadas.")
        return

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(
            locale="pt-PT",
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()

        try:
            for target in pending:
                process_category(
                    page,
                    target,
                    progress,
                    args.output,
                    max_products=max_products,
                    max_pages=max_pages,
                    use_progress=use_progress,
                    use_git=use_git,
                )
                if not args.test:
                    auchan.human_delay(1.0, 2.0)

            records = progress.get("records", [])
            if records:
                auchan.export_to_excel(records, args.output)
                if use_git:
                    commit_checkpoint_to_repo(len(records))
            log(f"\nConcluído. Total no ficheiro: {len(records)} produtos em '{args.output}'.")
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
