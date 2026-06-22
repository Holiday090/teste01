"""
Robô de teste 2 — extração completa da categoria Casa (todas as páginas).

SETUP (executar no terminal antes da primeira utilização):
    pip install playwright pandas openpyxl
    python3 -m playwright install

EXECUÇÃO:
    PYTHONUNBUFFERED=1 python3 scraping_action_teste2.py

NOTAS:
    - Recolhe URLs de todas as páginas da listagem Casa (~758 artigos).
    - Visita cada produto individualmente com delays aleatórios (2–5 s).
    - Guarda progresso intermédio para retomar em caso de interrupção.
    - Exporta o resultado final para Scraping_Action_Teste2.xlsx
"""

from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scraping_action_teste import (
    CATEGORY_NAME,
    collect_all_category_product_urls,
    export_to_excel,
    extract_product_data,
    human_delay,
    log,
)

CASA_LISTING_URL = "https://www.action.com/pt-pt/c/casa/"
OUTPUT_FILE = "Scraping_Action_Teste2.xlsx"
PROGRESS_FILE = "scraping_action_teste2_progress.json"
CHECKPOINT_EVERY = 25


def load_progress() -> dict:
    """Carrega o progresso guardado, se existir."""
    progress_path = Path(PROGRESS_FILE)
    if not progress_path.exists():
        return {"product_urls": [], "records": [], "processed_urls": []}

    with progress_path.open("r", encoding="utf-8") as progress_file:
        return json.load(progress_file)


def save_progress(product_urls: list[str], records: list[dict[str, str]], processed_urls: list[str]) -> None:
    """Guarda o progresso atual em JSON."""
    progress_payload = {
        "product_urls": product_urls,
        "records": records,
        "processed_urls": processed_urls,
    }
    with Path(PROGRESS_FILE).open("w", encoding="utf-8") as progress_file:
        json.dump(progress_payload, progress_file, ensure_ascii=False, indent=2)


def main() -> None:
    progress = load_progress()
    records: list[dict[str, str]] = list(progress.get("records", []))
    processed_urls: set[str] = set(progress.get("processed_urls", []))

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
            if progress.get("product_urls"):
                product_urls = list(progress["product_urls"])
                log(
                    f"[Categoria: {CATEGORY_NAME}] Retomando execução com "
                    f"{len(product_urls)} URLs já recolhidas."
                )
            else:
                product_urls = collect_all_category_product_urls(
                    page,
                    CATEGORY_NAME,
                    CASA_LISTING_URL,
                    navigate_via_menu=True,
                )
                save_progress(product_urls, records, sorted(processed_urls))

            log(f"\n[Categoria: {CATEGORY_NAME}] Fase 2 — Processamento individual dos produtos.")
            log(f"[Categoria: {CATEGORY_NAME}] Produtos a processar: {len(product_urls)}")
            if not product_urls:
                log("Nenhum produto encontrado. O ficheiro Excel não será gerado.")
                return

            pending_urls = [url for url in product_urls if url not in processed_urls]
            log(f"[Categoria: {CATEGORY_NAME}] Produtos pendentes: {len(pending_urls)}")

            for index, product_url in enumerate(pending_urls, start=1):
                overall_index = len(processed_urls) + 1
                log(
                    f"[Categoria: {CATEGORY_NAME}] A processar produto "
                    f"{overall_index}/{len(product_urls)}..."
                )

                try:
                    product_data = extract_product_data(page, product_url, CATEGORY_NAME)
                    records.append(product_data)
                    processed_urls.add(product_url)
                    log(
                        f"[Categoria: {CATEGORY_NAME}] Produto {overall_index}/{len(product_urls)} concluído "
                        f"| Sub-categoria: {product_data['Sub-categoria'] or '—'} "
                        f"| {product_data['Descrição / Nome do artigo'][:60]} "
                        f"| Regular: {product_data['Preço Regular']} "
                        f"| Promo: {product_data['Preço Promocional'] or '—'}"
                    )
                except Exception as error:
                    log(
                        f"[Categoria: {CATEGORY_NAME}] Erro no produto "
                        f"{overall_index}/{len(product_urls)}: {error}"
                    )

                if overall_index < len(product_urls):
                    human_delay(2, 5)

                if overall_index % CHECKPOINT_EVERY == 0 or overall_index == len(product_urls):
                    export_to_excel(records, OUTPUT_FILE)
                    save_progress(product_urls, records, sorted(processed_urls))
                    log(
                        f"[Categoria: {CATEGORY_NAME}] Checkpoint — "
                        f"{len(records)} registos guardados em '{OUTPUT_FILE}'."
                    )

            export_to_excel(records, OUTPUT_FILE)
            save_progress(product_urls, records, sorted(processed_urls))
            log(
                f"\n[Categoria: {CATEGORY_NAME}] Concluído. "
                f"{len(records)} produtos exportados para '{OUTPUT_FILE}'."
            )

        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
