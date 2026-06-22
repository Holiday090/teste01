"""
Robô completo — extração de todas as categorias Action (do início ao fim).

SETUP (executar no terminal antes da primeira utilização):
    pip install playwright pandas openpyxl
    python3 -m playwright install

EXECUÇÃO:
    PYTHONUNBUFFERED=1 python3 scraping_action_completo.py

NOTAS:
    - Percorre automaticamente as 15 categorias principais do menu Produtos.
    - Em cada categoria: recolhe todas as páginas e visita cada produto.
    - Quando termina uma categoria, avança para a seguinte até concluir todas.
    - Guarda progresso intermédio para retomar em caso de interrupção.
    - Exporta o resultado final para Scraping_Action_Completo.xlsx
"""

from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scraping_action_teste import (
    BASE_URL,
    accept_cookies_if_visible,
    collect_all_category_product_urls,
    discover_main_categories,
    export_to_excel,
    extract_product_data,
    human_delay,
    log,
)

OUTPUT_FILE = "Scraping_Action_Completo.xlsx"
PROGRESS_FILE = "scraping_action_completo_progress.json"
CHECKPOINT_EVERY = 25


def empty_progress() -> dict:
    """Estado inicial do ficheiro de progresso."""
    return {
        "categories": [],
        "completed_categories": [],
        "records": [],
        "processed_urls": [],
        "in_progress": None,
    }


def load_progress() -> dict:
    """Carrega o progresso guardado, se existir."""
    progress_path = Path(PROGRESS_FILE)
    if not progress_path.exists():
        return empty_progress()

    with progress_path.open("r", encoding="utf-8") as progress_file:
        progress = json.load(progress_file)

    for key, default_value in empty_progress().items():
        if key not in progress:
            progress[key] = default_value

    return progress


def save_progress(progress: dict) -> None:
    """Guarda o progresso atual em JSON."""
    with Path(PROGRESS_FILE).open("w", encoding="utf-8") as progress_file:
        json.dump(progress, progress_file, ensure_ascii=False, indent=2)


def ensure_categories(page, progress: dict) -> list[dict[str, str]]:
    """Garante que a lista de categorias está disponível no progresso."""
    if progress.get("categories"):
        return list(progress["categories"])

    log("A descobrir categorias principais no menu Produtos...")
    page.goto(BASE_URL, wait_until="domcontentloaded", timeout=90_000)
    accept_cookies_if_visible(page)
    categories = discover_main_categories(page)
    progress["categories"] = categories

    log(f"Categorias encontradas ({len(categories)}):")
    for index, category in enumerate(categories, start=1):
        log(f"  {index}. {category['name']} — {category['listing_url']}")

    save_progress(progress)
    return categories


def get_category_product_urls(page, category: dict[str, str], progress: dict) -> list[str]:
    """Obtém as URLs de produto de uma categoria, reutilizando progresso se existir."""
    in_progress = progress.get("in_progress") or {}
    if (
        in_progress.get("category") == category["name"]
        and in_progress.get("product_urls")
    ):
        product_urls = list(in_progress["product_urls"])
        log(
            f"[Categoria: {category['name']}] Reutilizando "
            f"{len(product_urls)} URLs já recolhidas."
        )
        return product_urls

    product_urls = collect_all_category_product_urls(
        page,
        category["name"],
        category["listing_url"],
        navigate_via_menu=True,
    )
    progress["in_progress"] = {
        "category": category["name"],
        "product_urls": product_urls,
    }
    save_progress(progress)
    return product_urls


def process_category(
    page,
    category: dict[str, str],
    progress: dict,
    records: list[dict[str, str]],
    processed_urls: set[str],
) -> None:
    """Processa uma categoria completa: URLs + produtos individuais."""
    category_name = category["name"]
    category_index = progress["categories"].index(category) + 1
    total_categories = len(progress["categories"])

    log(
        f"\n{'=' * 70}\n"
        f"[Categoria {category_index}/{total_categories}: {category_name}] "
        f"A iniciar extração completa.\n"
        f"{'=' * 70}"
    )

    product_urls = get_category_product_urls(page, category, progress)
    if not product_urls:
        log(f"[Categoria: {category_name}] Nenhum produto encontrado. A avançar.")
        progress["completed_categories"].append(category_name)
        progress["in_progress"] = None
        save_progress(progress)
        return

    log(f"\n[Categoria: {category_name}] Fase 2 — Processamento individual dos produtos.")
    log(f"[Categoria: {category_name}] Produtos a processar: {len(product_urls)}")

    in_progress = progress.get("in_progress") or {}
    category_processed = set(in_progress.get("category_processed_urls", []))
    pending_urls = [url for url in product_urls if url not in category_processed]

    log(f"[Categoria: {category_name}] Produtos pendentes: {len(pending_urls)}")

    for product_url in pending_urls:
        overall_index = len(category_processed) + 1
        log(
            f"[Categoria: {category_name}] A processar produto "
            f"{overall_index}/{len(product_urls)}..."
        )

        try:
            product_data = extract_product_data(page, product_url, category_name)
            records.append(product_data)
            processed_urls.add(product_url)
            category_processed.add(product_url)
            progress["records"] = records
            progress["processed_urls"] = sorted(processed_urls)
            progress["in_progress"] = {
                "category": category_name,
                "product_urls": product_urls,
                "category_processed_urls": sorted(category_processed),
            }

            log(
                f"[Categoria: {category_name}] Produto {overall_index}/{len(product_urls)} concluído "
                f"| Sub-categoria: {product_data['Sub-categoria'] or '—'} "
                f"| {product_data['Descrição / Nome do artigo'][:60]} "
                f"| Regular: {product_data['Preço Regular']} "
                f"| Promo: {product_data['Preço Promocional'] or '—'}"
            )
        except Exception as error:
            log(
                f"[Categoria: {category_name}] Erro no produto "
                f"{overall_index}/{len(product_urls)}: {error}"
            )

        if overall_index < len(product_urls):
            human_delay(2, 5)

        if overall_index % CHECKPOINT_EVERY == 0 or overall_index == len(product_urls):
            export_to_excel(records, OUTPUT_FILE)
            save_progress(progress)
            log(
                f"[Categoria: {category_name}] Checkpoint — "
                f"{len(records)} registos totais guardados em '{OUTPUT_FILE}'."
            )

    progress["completed_categories"].append(category_name)
    progress["in_progress"] = None
    save_progress(progress)
    export_to_excel(records, OUTPUT_FILE)

    log(
        f"\n[Categoria: {category_name}] Concluída — "
        f"{len(product_urls)} produtos processados."
    )


def main() -> None:
    progress = load_progress()
    records: list[dict[str, str]] = list(progress.get("records", []))
    processed_urls: set[str] = set(progress.get("processed_urls", []))
    completed_categories: set[str] = set(progress.get("completed_categories", []))

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(
            locale="pt-PT",
            viewport={"width": 1440, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()

        try:
            categories = ensure_categories(page, progress)
            pending_categories = [
                category
                for category in categories
                if category["name"] not in completed_categories
            ]

            if completed_categories:
                log(
                    f"\nRetomando execução. Categorias já concluídas: "
                    f"{', '.join(progress['completed_categories'])}"
                )

            if not pending_categories:
                log("\nTodas as categorias já foram processadas.")
                export_to_excel(records, OUTPUT_FILE)
                log(
                    f"Ficheiro final disponível em '{OUTPUT_FILE}' "
                    f"com {len(records)} produtos."
                )
                return

            log(
                f"\nCategorias pendentes ({len(pending_categories)}): "
                f"{', '.join(category['name'] for category in pending_categories)}"
            )

            for category in pending_categories:
                process_category(
                    page,
                    category,
                    progress,
                    records,
                    processed_urls,
                )

            export_to_excel(records, OUTPUT_FILE)
            save_progress(progress)
            log(
                f"\n{'=' * 70}\n"
                f"EXTRAÇÃO COMPLETA CONCLUÍDA\n"
                f"Categorias processadas: {len(progress['completed_categories'])}/{len(categories)}\n"
                f"Total de produtos exportados: {len(records)}\n"
                f"Ficheiro: {OUTPUT_FILE}\n"
                f"{'=' * 70}"
            )

        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
