"""
Robô de teste 2 — extração completa da categoria Casa (todas as páginas).

SETUP (executar no terminal antes da primeira utilização):
    pip install playwright pandas openpyxl
    python3 -m playwright install

EXECUÇÃO:
    python3 scraping_action_teste2.py

NOTAS:
    - Recolhe URLs de todas as páginas da listagem Casa (~758 artigos).
    - Visita cada produto individualmente com delays aleatórios (2–5 s).
    - Guarda progresso intermédio para retomar em caso de interrupção.
    - Exporta o resultado final para Scraping_Action_Teste2.xlsx
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from urllib.parse import urljoin

from playwright.sync_api import Page, sync_playwright

from scraping_action_teste import (
    BASE_URL,
    CATEGORY_NAME,
    PRODUCT_LINK_PATTERN,
    accept_cookies_if_visible,
    export_to_excel,
    extract_product_data,
    human_delay,
    log,
    open_casa_listing_page,
)

CASA_LISTING_URL = "https://www.action.com/pt-pt/c/casa/"
OUTPUT_FILE = "Scraping_Action_Teste2.xlsx"
PROGRESS_FILE = "scraping_action_teste2_progress.json"
CHECKPOINT_EVERY = 25


def listing_url_for_page(page_number: int) -> str:
    """Constrói o URL da listagem Casa para uma página específica."""
    if page_number <= 1:
        return CASA_LISTING_URL
    return f"{CASA_LISTING_URL}?page={page_number}#product-grid"


def get_listing_metadata(page: Page) -> tuple[int, int]:
    """Obtém o total de resultados e o número da última página."""
    page.wait_for_timeout(1000)

    metadata = page.evaluate(
        """() => {
            const textNodes = [...document.querySelectorAll('h1, h2, p, span, div')]
                .map((element) => element.innerText.trim())
                .filter(Boolean);

            let totalResults = 0;
            for (const text of textNodes) {
                const match = text.match(/(\\d+)\\s+resultados?/i);
                if (match) {
                    totalResults = Number(match[1]);
                    break;
                }
            }

            const pageNumbers = [...document.querySelectorAll('a[href*="page="]')]
                .map((anchor) => {
                    const match = anchor.getAttribute('href')?.match(/page=(\\d+)/);
                    return match ? Number(match[1]) : 0;
                })
                .filter((value) => value > 0);

            const lastPage = pageNumbers.length ? Math.max(...pageNumbers) : 1;
            return { totalResults, lastPage };
        }"""
    )

    total_results = int(metadata.get("totalResults", 0))
    last_page = int(metadata.get("lastPage", 1))
    return total_results, last_page


def collect_urls_from_listing_page(page: Page) -> list[str]:
    """Recolhe URLs únicas dos produtos visíveis na página de listagem atual."""
    page.wait_for_selector('a[href*="/pt-pt/p/"]', timeout=30_000)

    raw_hrefs = page.locator('a[href*="/pt-pt/p/"]').evaluate_all(
        "elements => elements.map(el => el.getAttribute('href')).filter(Boolean)"
    )

    product_urls: list[str] = []
    seen: set[str] = set()

    for href in raw_hrefs:
        if not PRODUCT_LINK_PATTERN.search(href):
            continue

        absolute_url = urljoin(BASE_URL, href)
        if absolute_url in seen:
            continue

        seen.add(absolute_url)
        product_urls.append(absolute_url)

    return product_urls


def collect_all_casa_product_urls(page: Page) -> list[str]:
    """Percorre todas as páginas da categoria Casa e devolve URLs únicas."""
    log(f"\n[Categoria: {CATEGORY_NAME}] Fase 1 — Recolha de URLs de todas as páginas.")
    open_casa_listing_page(page)
    total_results, last_page = get_listing_metadata(page)

    log(f"[Categoria: {CATEGORY_NAME}] Total de resultados: {total_results}")
    log(f"[Categoria: {CATEGORY_NAME}] Páginas a percorrer: {last_page}")

    all_product_urls: list[str] = []
    seen_urls: set[str] = set()

    for page_number in range(1, last_page + 1):
        if page_number > 1:
            log(f"[Categoria: {CATEGORY_NAME}] A abrir página {page_number}/{last_page}...")
            page.goto(
                listing_url_for_page(page_number),
                wait_until="domcontentloaded",
                timeout=90_000,
            )
            page.wait_for_timeout(1200)
            human_delay(1.5, 3.0)

        page_urls = collect_urls_from_listing_page(page)
        new_urls = [url for url in page_urls if url not in seen_urls]
        seen_urls.update(new_urls)
        all_product_urls.extend(new_urls)

        log(
            f"[Categoria: {CATEGORY_NAME}] Página {page_number}/{last_page} concluída "
            f"— {len(new_urls)} URLs novas | Total acumulado: {len(all_product_urls)}/{total_results}"
        )

    log(
        f"\n[Categoria: {CATEGORY_NAME}] Fase 1 concluída — "
        f"{len(all_product_urls)} URLs únicas recolhidas."
    )
    return all_product_urls


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
                product_urls = collect_all_casa_product_urls(page)
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
                    product_data = extract_product_data(page, product_url)
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
