"""
Robô de teste para extração de produtos da Action (categoria Casa - Página 1).

SETUP (executar no terminal antes da primeira utilização):
    pip install playwright pandas openpyxl
    python3 -m playwright install

EXECUÇÃO:
    python3 scraping_action_teste.py
"""

from __future__ import annotations

import random
import re
import time
from urllib.parse import urljoin

import pandas as pd
from playwright.sync_api import Page, sync_playwright

BASE_URL = "https://www.action.com/pt-pt/"
CATEGORY_NAME = "Casa"
OUTPUT_FILE = "Scraping_Action_Teste.xlsx"
PRODUCT_LINK_PATTERN = re.compile(r"/pt-pt/p/\d+/")


def log(message: str) -> None:
    """Imprime mensagens de progresso no terminal em tempo real."""
    print(message, flush=True)


def human_delay(min_seconds: float = 2.0, max_seconds: float = 5.0) -> None:
    """Pausa aleatória para simular comportamento humano."""
    time.sleep(random.uniform(min_seconds, max_seconds))


def accept_cookies_if_visible(page: Page) -> None:
    """Aceita o banner de cookies quando este aparece."""
    cookie_selectors = [
        "#onetrust-accept-btn-handler",
        'button:has-text("Aceitar todos")',
        'button:has-text("Aceitar")',
    ]
    for selector in cookie_selectors:
        button = page.locator(selector).first
        if button.count() and button.is_visible():
            button.click()
            page.wait_for_timeout(800)
            print("Cookies aceites.", flush=True)
            return


def open_casa_listing_page(page: Page) -> None:
    """Navega até à listagem 'Tudo de Casa' através do menu principal."""
    log(f"A aceder a {BASE_URL}")
    page.goto(BASE_URL, wait_until="domcontentloaded", timeout=90_000)
    human_delay(2, 4)

    accept_cookies_if_visible(page)

    produtos_button = page.get_by_role("button", name=re.compile(r"^Produtos$", re.I)).first
    if not produtos_button.count():
        raise RuntimeError('Botão "Produtos" não encontrado no menu principal.')

    produtos_button.click()
    page.wait_for_timeout(1000)
    log(f'[Categoria: {CATEGORY_NAME}] Menu "Produtos" aberto.')

    casa_link = page.get_by_role("link", name=re.compile(r"^Casa$", re.I)).first
    if not casa_link.count():
        raise RuntimeError('Categoria "Casa" não encontrada no submenu.')

    casa_link.hover()
    page.wait_for_timeout(1200)
    log(f'[Categoria: {CATEGORY_NAME}] Hover sobre a categoria principal.')

    tudo_de_casa = page.get_by_role("link", name=re.compile(r"Tudo de Casa", re.I)).first
    if not tudo_de_casa.count():
        raise RuntimeError('Opção "Tudo de Casa" não encontrada no menu lateral.')

    tudo_de_casa.click()
    page.wait_for_load_state("domcontentloaded", timeout=90_000)
    page.wait_for_timeout(1500)
    log(f'[Categoria: {CATEGORY_NAME}] Listagem aberta: {page.url}')


def collect_product_urls_from_page_1(page: Page) -> list[str]:
    """Recolhe URLs únicas dos produtos visíveis na primeira página."""
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


def format_price(whole: str, fractional: str) -> str:
    """Combina parte inteira e centimos num preço legível."""
    whole = (whole or "").strip()
    fractional = (fractional or "").strip()

    if whole and fractional:
        return f"{whole},{fractional}"
    if whole:
        return whole
    return ""


def extract_json_ld_product(page: Page) -> dict:
    """Extrai o bloco Product do JSON-LD da página."""
    return page.evaluate(
        """() => {
            const scripts = [...document.querySelectorAll('script[type="application/ld+json"]')];
            for (const script of scripts) {
                try {
                    const data = JSON.parse(script.textContent);
                    if (data && data['@type'] === 'Product') {
                        return data;
                    }
                } catch (error) {
                    continue;
                }
            }
            return {};
        }"""
    )


def extract_json_ld_breadcrumb(page: Page) -> dict:
    """Extrai o bloco BreadcrumbList do JSON-LD da página."""
    return page.evaluate(
        """() => {
            const scripts = [...document.querySelectorAll('script[type="application/ld+json"]')];
            for (const script of scripts) {
                try {
                    const data = JSON.parse(script.textContent);
                    if (data && data['@type'] === 'BreadcrumbList') {
                        return data;
                    }
                } catch (error) {
                    continue;
                }
            }
            return {};
        }"""
    )


def extract_subcategory(page: Page, category_name: str) -> str:
    """Obtém a sub-categoria a partir do breadcrumb (2.º nível após a categoria principal)."""
    breadcrumb_ld = extract_json_ld_breadcrumb(page)
    items = breadcrumb_ld.get("itemListElement", [])

    for item in items:
        if not isinstance(item, dict):
            continue

        position = item.get("position")
        entry = item.get("item", {})
        if position == 2 and isinstance(entry, dict):
            subcategory = str(entry.get("name", "")).strip()
            if subcategory:
                return subcategory

    breadcrumb_links = page.locator('nav[aria-label="Breadcrumbs"] a')
    link_count = breadcrumb_links.count()
    if link_count >= 2:
        first_link = breadcrumb_links.nth(0).inner_text().strip()
        second_link = breadcrumb_links.nth(1).inner_text().strip()
        if first_link.lower() == category_name.lower() and second_link:
            return second_link

    return ""


def extract_brand(product_ld: dict, product_name: str) -> str:
    """Obtém a marca a partir do JSON-LD ou, em alternativa, do nome do produto."""
    brand = product_ld.get("brand")
    if isinstance(brand, dict):
        brand_name = str(brand.get("name", "")).strip()
        if brand_name:
            return brand_name

    if not product_name:
        return ""

    known_brands = [
        "Grundig",
        "Endless Scent",
        "Zenova",
        "Heinz",
        "Dove",
        "Gillette",
        "Energizer",
        "Panasonic",
        "Listerine",
        "Slazenger",
        "Bestway",
        "Nor-Tec",
    ]
    for known_brand in known_brands:
        if known_brand.lower() in product_name.lower():
            return known_brand

    return ""


def extract_product_data(page: Page, product_url: str) -> dict[str, str]:
    """Extrai os campos solicitados na página individual do produto."""
    page.goto(product_url, wait_until="domcontentloaded", timeout=90_000)
    page.wait_for_selector("main h1", timeout=30_000)
    page.wait_for_timeout(800)

    accept_cookies_if_visible(page)

    product_name = page.locator("main h1").first.inner_text().strip()
    subtitle = page.locator("main h1 + p").first
    subtitle_text = subtitle.inner_text().strip() if subtitle.count() else ""

    if subtitle_text:
        description = f"{product_name} | {subtitle_text}"
    else:
        description = product_name

    price_container = page.locator('main [data-testid="product-card-price"]').first
    whole = price_container.locator('[data-testid="product-card-price-whole"]').first.inner_text().strip()
    fractional = price_container.locator('[data-testid="product-card-price-fractional"]').first.inner_text().strip()
    current_price = format_price(whole, fractional)

    original_locator = price_container.locator('[data-testid="product-card-price-original-amount"]').first
    original_price = original_locator.inner_text().strip() if original_locator.count() else ""

    if original_price:
        regular_price = original_price
        promotional_price = current_price
    else:
        regular_price = current_price
        promotional_price = ""

    product_ld = extract_json_ld_product(page)
    brand = extract_brand(product_ld, product_name)
    subcategory = extract_subcategory(page, CATEGORY_NAME)

    return {
        "Categoria Principal": CATEGORY_NAME,
        "Sub-categoria": subcategory,
        "Marca": brand,
        "Descrição / Nome do artigo": description,
        "Preço Regular": regular_price,
        "Preço Promocional": promotional_price,
        "URL": product_url,
    }


def export_to_excel(records: list[dict[str, str]], output_file: str) -> None:
    """Exporta os registos recolhidos para Excel."""
    columns = [
        "Categoria Principal",
        "Sub-categoria",
        "Marca",
        "Descrição / Nome do artigo",
        "Preço Regular",
        "Preço Promocional",
        "URL",
    ]
    dataframe = pd.DataFrame(records, columns=columns)
    dataframe.to_excel(output_file, index=False, engine="openpyxl")


def main() -> None:
    records: list[dict[str, str]] = []

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
            log(f"\n[Categoria: {CATEGORY_NAME}] A iniciar recolha de URLs (apenas Página 1).")
            open_casa_listing_page(page)
            product_urls = collect_product_urls_from_page_1(page)

            log(
                f"[Categoria: {CATEGORY_NAME}] Página 1 concluída — "
                f"{len(product_urls)} produtos encontrados."
            )
            if not product_urls:
                log("Nenhum produto encontrado. O ficheiro Excel não será gerado.")
                return

            log(f"\n[Categoria: {CATEGORY_NAME}] A iniciar processamento individual dos produtos.")
            for index, product_url in enumerate(product_urls, start=1):
                log(f"[Categoria: {CATEGORY_NAME}] A processar produto {index}/{len(product_urls)}...")
                try:
                    product_data = extract_product_data(page, product_url)
                    records.append(product_data)
                    log(
                        f"[Categoria: {CATEGORY_NAME}] Produto {index}/{len(product_urls)} concluído "
                        f"| Sub-categoria: {product_data['Sub-categoria'] or '—'} "
                        f"| {product_data['Descrição / Nome do artigo'][:60]} "
                        f"| Regular: {product_data['Preço Regular']} "
                        f"| Promo: {product_data['Preço Promocional'] or '—'}"
                    )
                except Exception as error:
                    log(f"[Categoria: {CATEGORY_NAME}] Erro no produto {index}/{len(product_urls)}: {error}")

                if index < len(product_urls):
                    human_delay(2, 5)

            export_to_excel(records, OUTPUT_FILE)
            log(
                f"\n[Categoria: {CATEGORY_NAME}] Concluído. "
                f"{len(records)} produtos exportados para '{OUTPUT_FILE}'."
            )

        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
