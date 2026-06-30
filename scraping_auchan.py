"""
Robô de extração de produtos do Auchan Portugal.

SETUP (executar no terminal antes da primeira utilização):
    pip install playwright pandas openpyxl
    python3 -m playwright install chromium

EXECUÇÃO:
    PYTHONUNBUFFERED=1 python3 scraping_auchan.py
    PYTHONUNBUFFERED=1 python3 scraping_auchan.py --test   # amostra reduzida para validação
    PYTHONUNBUFFERED=1 python3 scraping_auchan.py --only frescos --max-pages 2
    PYTHONUNBUFFERED=1 python3 scraping_auchan.py --only frescos --max-pages 2 --max-products 10 --random

PROGRESSO NO TERMINAL:
    O script imprime o estado da extração em tempo real (use PYTHONUNBUFFERED=1).
    Durante a execução verá mensagens como:

        [Categoria: Charcutaria] Fase 1 — Recolha de produtos (máx. 2 páginas).
        [Categoria: Charcutaria] Página 1/2 (offset 0) — 48 produtos únicos recolhidos.
        [Categoria: Charcutaria] Fase 2 — Extração detalhada de produtos.
        Categoria atual: Charcutaria | A processar produto 8/20... | Bacon Extra Cubos Auchan 2x75g
        Categoria atual: Charcutaria | Produto 8/20 concluído | Produtos Frescos -> ... | EAN: ... | Preço: ...

    Em caso de erro num produto específico, a linha correspondente é mostrada e o script
    continua para o produto seguinte.

NOTAS:
    - Hierarquia de categorias interpretada como:
      Categoria -> Sub Categoria -> Família -> Sub Família
      Exemplo: Alimentação -> Congelados -> Peixe -> Bacalhau
    - Navega pelo menu principal e extrai produtos das categorias definidas.
    - Percorre todas as páginas de cada subcategoria (botão "Ver mais produtos" ou paginação por URL).
    - Visita cada produto para obter breadcrumb completo, EAN e preços detalhados.
    - Exporta o resultado para Scraping_Auchan.xlsx
    - Grava checkpoint Excel a cada 100 produtos e retoma automaticamente após interrupção.
    - Repete pedidos com falha durante até 3 minutos antes de desistir.
"""

from __future__ import annotations

import argparse
import json
import random
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, TypeVar
from urllib.parse import urljoin, urlparse, urlunparse

import pandas as pd
from playwright.sync_api import Page, sync_playwright

BASE_URL = "https://www.auchan.pt"
OUTPUT_FILE = "Scraping_Auchan.xlsx"
PROGRESS_FILE = "scraping_auchan_progress.json"
CHECKPOINT_EVERY = 100
RETRY_TIMEOUT_SECONDS = 180
PAGE_SIZE = 48
PRODUCT_ID_PATTERN = re.compile(r"/(\d+)\.html")
QUANTITY_PATTERN = re.compile(
    r"(\d+(?:[.,]\d+)?\s*(?:x\s*\d+(?:[.,]\d+)?\s*)?(?:g|gr|kg|ml|cl|l|lt|un)(?:\s*\([^)]*\))?)",
    re.IGNORECASE,
)

CATEGORY_LEVELS: tuple[str, ...] = (
    "Categoria",
    "Sub Categoria",
    "Família",
    "Sub Família",
)

EXCEL_COLUMNS = [
    *CATEGORY_LEVELS,
    "Caminho Categorias",
    "Descrição do Produto",
    "EAN / Referência",
    "Preço Regular",
    "Preço Promocional",
    "Preço ao Quilo/Litro",
    "Quantidade da Embalagem",
    "URL",
    "Categoria Alvo",
]


@dataclass(frozen=True)
class CategoryTarget:
    """Definição de uma categoria a extrair."""

    display_name: str
    menu_path: tuple[str, ...]
    listing_url: str
    listing_fallback_names: tuple[str, ...] = ()


CATEGORY_TARGETS: tuple[CategoryTarget, ...] = (
    CategoryTarget(
        display_name="Charcutaria",
        menu_path=("Produtos Frescos", "Charcutaria"),
        listing_url="/pt/produtos-frescos/charcutaria/",
    ),
    CategoryTarget(
        display_name="Queijaria",
        menu_path=("Produtos Frescos", "Queijaria"),
        listing_url="/pt/produtos-frescos/queijaria/",
    ),
    CategoryTarget(
        display_name="Lácteos",
        menu_path=("Alimentação", "Produtos Lácteos"),
        listing_url="/pt/alimentacao/produtos-lacteos/",
        listing_fallback_names=("Produtos Lácteos", "Lácteos"),
    ),
    CategoryTarget(
        display_name="Congelados",
        menu_path=("Alimentação", "Congelados"),
        listing_url="/pt/alimentacao/congelados/",
        listing_fallback_names=("Ver todos", "Congelados"),
    ),
)

MENU_ITEM_IDS: dict[str, str] = {
    "Produtos Frescos": "menu-category-produtos-frescos",
    "Frescos": "menu-category-produtos-frescos",
    "Alimentação": "menu-category-alimentacao-",
    "Charcutaria": "menu-subcategory-charcutaria",
    "Queijaria": "menu-subcategory-queijaria",
    "Congelados": "menu-subcategory-congelados",
    "Produtos Lácteos": "menu-subcategory-produtos-lacteos",
}


def log(message: str) -> None:
    """Imprime mensagens de progresso no terminal em tempo real."""
    print(message, flush=True)


def log_product_progress(
    category: str,
    current: int,
    total: int,
    *,
    status: str = "processing",
    product_name: str = "",
    details: str = "",
) -> None:
    """Regista o progresso produto a produto no terminal (uma linha por evento)."""
    if status == "processing":
        preview = f" | {product_name[:70]}" if product_name else ""
        log(f"Categoria atual: {category} | A processar produto {current}/{total}...{preview}")
        return

    if status == "done":
        extra = f" | {details}" if details else ""
        log(f"Categoria atual: {category} | Produto {current}/{total} concluído{extra}")
        return

    if status == "error":
        log(f"Categoria atual: {category} | Erro no produto {current}/{total} | {details}")


def human_delay(min_seconds: float = 1.0, max_seconds: float = 2.5) -> None:
    """Pausa aleatória para simular comportamento humano."""
    time.sleep(random.uniform(min_seconds, max_seconds))


T = TypeVar("T")


def retry_for_duration(
    operation_name: str,
    operation: Callable[[], T],
    *,
    max_duration_seconds: float = RETRY_TIMEOUT_SECONDS,
    initial_delay_seconds: float = 5.0,
    max_delay_seconds: float = 30.0,
) -> T:
    """Repete uma operação durante até max_duration_seconds quando falha."""
    started_at = time.monotonic()
    attempt = 0
    delay_seconds = initial_delay_seconds
    last_error: Exception | None = None

    while time.monotonic() - started_at < max_duration_seconds:
        attempt += 1
        try:
            return operation()
        except Exception as error:
            last_error = error if isinstance(error, Exception) else Exception(str(error))
            elapsed_seconds = time.monotonic() - started_at
            remaining_seconds = max_duration_seconds - elapsed_seconds
            if remaining_seconds <= 0:
                break
            wait_seconds = min(delay_seconds, remaining_seconds)
            log(
                f"[Retry] {operation_name} falhou (tentativa {attempt}): {last_error}. "
                f"Nova tentativa em {wait_seconds:.0f}s..."
            )
            time.sleep(wait_seconds)
            delay_seconds = min(delay_seconds * 1.5, max_delay_seconds)

    if last_error is not None:
        raise last_error
    raise RuntimeError(f"{operation_name} falhou após {max_duration_seconds:.0f}s.")


def goto_with_retry(
    page: Page,
    url: str,
    *,
    operation_name: str = "",
    wait_until: str = "domcontentloaded",
    timeout: int = 90_000,
    settle_ms: int = 900,
) -> None:
    """Navega para um URL com retry automático durante 3 minutos."""
    label = operation_name or url

    def _navigate() -> None:
        page.goto(url, wait_until=wait_until, timeout=timeout)
        if settle_ms > 0:
            page.wait_for_timeout(settle_ms)

    retry_for_duration(label, _navigate)


def empty_progress(output_file: str = OUTPUT_FILE) -> dict:
    """Estado inicial do ficheiro de progresso."""
    return {
        "output_file": output_file,
        "completed_categories": [],
        "records": [],
        "processed_urls": [],
        "in_progress": None,
    }


def load_progress(output_file: str = OUTPUT_FILE) -> dict:
    """Carrega o progresso guardado, se existir."""
    progress_path = Path(PROGRESS_FILE)
    if not progress_path.exists():
        return empty_progress(output_file)

    with progress_path.open("r", encoding="utf-8") as progress_file:
        progress = json.load(progress_file)

    defaults = empty_progress(output_file)
    for key, default_value in defaults.items():
        if key not in progress:
            progress[key] = default_value

    if progress.get("output_file") != output_file:
        log(
            f"Aviso: ficheiro de progresso refere '{progress.get('output_file')}', "
            f"mas o output atual é '{output_file}'."
        )
    return progress


def save_progress(progress: dict) -> None:
    """Guarda o progresso atual em JSON."""
    with Path(PROGRESS_FILE).open("w", encoding="utf-8") as progress_file:
        json.dump(progress, progress_file, ensure_ascii=False, indent=2)


def maybe_checkpoint(records: list[dict[str, str]], output_file: str, progress: dict) -> None:
    """Guarda Excel parcial e progresso a cada CHECKPOINT_EVERY produtos."""
    if not records or len(records) % CHECKPOINT_EVERY != 0:
        return

    export_to_excel(records, output_file)
    save_progress(progress)
    log(
        f"[Checkpoint] {len(records)} produtos guardados em '{output_file}' "
        f"e progresso atualizado em '{PROGRESS_FILE}'."
    )


def na(value: str | None) -> str:
    """Normaliza valores em falta para 'N/A'."""
    if value is None:
        return "N/A"
    cleaned = str(value).strip()
    return cleaned if cleaned else "N/A"


def absolute_url(path_or_url: str) -> str:
    """Converte caminhos relativos em URLs absolutas."""
    return urljoin(BASE_URL, path_or_url)


def normalize_product_url(url: str) -> str:
    """Remove parâmetros de query e normaliza a URL do produto."""
    parsed = urlparse(absolute_url(url))
    return urlunparse((parsed.scheme, parsed.netloc, parsed.path, "", "", ""))


def extract_product_id(url: str) -> str:
    """Extrai o ID numérico do produto a partir do URL."""
    match = PRODUCT_ID_PATTERN.search(url)
    return match.group(1) if match else ""


def accept_cookies_if_visible(page: Page) -> None:
    """Aceita o banner de cookies quando este aparece."""
    cookie_selectors = [
        "#onetrust-accept-btn-handler",
        'button:has-text("Aceitar todos")',
        'button:has-text("Aceitar")',
    ]
    for selector in cookie_selectors:
        button = page.locator(selector).first
        try:
            if button.count() and button.is_visible():
                button.click()
                page.wait_for_timeout(800)
                log("Cookies aceites.")
                return
        except Exception:
            continue


def dismiss_blocking_modals(page: Page) -> None:
    """Fecha pop-ups que possam bloquear a interação com a página."""
    modal_selectors = [
        'button:has-text("Entrega em Casa")',
        'button:has-text("Recolha na Loja")',
        'button:has-text("Continuar")',
        'button:has-text("Fechar")',
        '[aria-label="Close"]',
    ]
    for selector in modal_selectors:
        button = page.locator(selector).first
        try:
            if button.count() and button.is_visible():
                button.click()
                page.wait_for_timeout(500)
        except Exception:
            continue


def click_menu_item_by_id(page: Page, element_id: str) -> bool:
    """Clica num item de menu pelo ID, usando JavaScript se necessário."""
    clicked = page.evaluate(
        """(itemId) => {
            const element = document.getElementById(itemId);
            if (!element) {
                return false;
            }
            element.scrollIntoView({ block: "center", inline: "center" });
            element.dispatchEvent(new MouseEvent("mouseenter", { bubbles: true }));
            element.click();
            return true;
        }""",
        element_id,
    )
    if clicked:
        page.wait_for_timeout(900)
    return bool(clicked)


def click_menu_link_by_text(page: Page, label: str) -> bool:
    """Clica num link de menu visível que corresponda ao texto indicado."""
    pattern = re.compile(rf"^{re.escape(label)}$", re.IGNORECASE)
    link = page.get_by_role("link", name=pattern).first
    try:
        if link.count() and link.is_visible():
            link.click()
            page.wait_for_timeout(900)
            return True
    except Exception:
        pass

    clicked = page.evaluate(
        """(labelText) => {
            const normalized = labelText.trim().toLowerCase();
            const links = [...document.querySelectorAll("a, button, span.dropdown-link")];
            for (const element of links) {
                const text = (element.innerText || element.textContent || "").trim();
                if (text.toLowerCase() === normalized && element.offsetParent !== null) {
                    element.scrollIntoView({ block: "center", inline: "center" });
                    element.click();
                    return true;
                }
            }
            return false;
        }""",
        label,
    )
    if clicked:
        page.wait_for_timeout(900)
    return bool(clicked)


def navigate_via_menu(page: Page, category: CategoryTarget) -> str:
    """Navega até à listagem da categoria através do menu principal."""
    log(f"A aceder a {BASE_URL}")
    goto_with_retry(page, BASE_URL, operation_name="página inicial Auchan", settle_ms=1500)
    accept_cookies_if_visible(page)
    dismiss_blocking_modals(page)

    for step in category.menu_path:
        if step in MENU_ITEM_IDS and click_menu_item_by_id(page, MENU_ITEM_IDS[step]):
            log(f'[Menu] Secção "{step}" aberta.')
            continue

        if click_menu_link_by_text(page, step):
            log(f'[Menu] Link "{step}" selecionado.')
            continue

        if step in category.listing_fallback_names and click_menu_link_by_text(page, step):
            log(f'[Menu] Fallback "{step}" selecionado.')
            continue

        log(f'[Menu] Item "{step}" não encontrado no menu; a usar URL direto como fallback.')

    page.wait_for_load_state("domcontentloaded", timeout=90_000)
    page.wait_for_timeout(1200)

    current_url = page.url
    expected_fragment = category.listing_url.split("?")[0]
    if expected_fragment not in current_url:
        log(f"[Menu] A abrir listagem por URL: {category.listing_url}")
        goto_with_retry(
            page,
            absolute_url(category.listing_url),
            operation_name=f"listagem {category.display_name}",
            settle_ms=1200,
        )

    accept_cookies_if_visible(page)
    dismiss_blocking_modals(page)
    log(f'[Categoria: {category.display_name}] Listagem aberta: {page.url}')
    return page.url


def get_listing_metadata(page: Page) -> tuple[int, int]:
    """Obtém o total de resultados e o intervalo visível na listagem."""
    metadata = page.evaluate(
        """() => {
            const match = document.body.innerText.match(
                /(\\d+)\\s*-\\s*(\\d+)\\s*de\\s*([\\d.,]+)\\s*resultados/i
            );
            if (!match) {
                return { totalResults: 0, from: 0, to: 0 };
            }

            const normalizeNumber = (value) => Number(value.replace(/\\./g, "").replace(",", ""));
            return {
                from: normalizeNumber(match[1]),
                to: normalizeNumber(match[2]),
                totalResults: normalizeNumber(match[3]),
            };
        }"""
    )
    total_results = int(metadata.get("totalResults", 0))
    visible_to = int(metadata.get("to", 0))
    return total_results, visible_to


def listing_url_for_offset(listing_url: str, start: int) -> str:
    """Constrói o URL da listagem para um deslocamento específico."""
    if start <= 0:
        return absolute_url(listing_url)
    separator = "&" if "?" in listing_url else "?"
    return absolute_url(f"{listing_url}{separator}start={start}&sz={PAGE_SIZE}")


def extract_products_from_listing(page: Page) -> list[dict[str, str]]:
    """Extrai dados básicos dos produtos visíveis na listagem atual."""
    return page.evaluate(
        """() => {
            const tiles = [...document.querySelectorAll(".auc-product-tile[data-pid]")];
            return tiles.map((tile) => {
                const pid = tile.getAttribute("data-pid") || "";
                const name = tile.querySelector(".auc-product-tile__name")?.innerText.trim() || "";
                const unitPrice = tile.querySelector(".auc-measures--price-per-unit")?.innerText.trim() || "";
                const link = tile.querySelector('a[href$=".html"]')?.getAttribute("href") || "";
                const listPrice = tile.querySelector(".strike-through.value, .auc-price__stricked .value")?.innerText.trim() || "";
                const salesPrice = tile.querySelector(".sales .value")?.innerText.trim() || "";
                return { pid, name, unitPrice, link, listPrice, salesPrice };
            }).filter((item) => item.pid && item.link);
        }"""
    )


def click_load_more_if_visible(page: Page) -> bool:
    """Clica no botão 'Ver mais produtos' se estiver visível."""
    button = page.locator(
        ".auc-js-show-more-next-button, button:has-text('Ver mais produtos'), button:has-text('Carregar mais')"
    ).first
    try:
        if button.count() and button.is_visible():
            button.scroll_into_view_if_needed()
            button.click()
            page.wait_for_timeout(2500)
            return True
    except Exception:
        return False
    return False


def collect_all_category_products(
    page: Page,
    category: CategoryTarget,
    *,
    navigate_via_menu_flag: bool = True,
    max_products: int | None = None,
    max_pages: int | None = None,
    random_sample: bool = False,
    progress: dict | None = None,
) -> list[dict[str, str]]:
    """Percorre todas as páginas de uma categoria e devolve produtos únicos."""
    pages_label = f" (máx. {max_pages} páginas)" if max_pages else ""
    log(f"\n[Categoria: {category.display_name}] Fase 1 — Recolha de produtos{pages_label}.")

    in_progress = (progress or {}).get("in_progress") or {}
    if (
        progress is not None
        and in_progress.get("category") == category.display_name
        and in_progress.get("phase") == "extracting"
        and in_progress.get("products")
    ):
        products = list(in_progress["products"])
        log(
            f"[Categoria: {category.display_name}] Retomando listagem já recolhida "
            f"({len(products)} produtos)."
        )
        if max_products:
            products = products[:max_products]
        return products

    collected: dict[str, dict[str, str]] = {}
    start_offset = 0
    if (
        progress is not None
        and in_progress.get("category") == category.display_name
        and in_progress.get("phase") == "collecting"
    ):
        start_offset = int(in_progress.get("collection_offset", 0))
        for product in in_progress.get("partial_products", []):
            if product.get("pid"):
                collected[str(product["pid"])] = product
        log(
            f"[Categoria: {category.display_name}] Retomando recolha a partir do offset "
            f"{start_offset} ({len(collected)} produtos parciais)."
        )

    if start_offset == 0:
        if navigate_via_menu_flag:
            navigate_via_menu(page, category)
        else:
            goto_with_retry(
                page,
                absolute_url(category.listing_url),
                operation_name=f"listagem {category.display_name}",
                settle_ms=1500,
            )
            accept_cookies_if_visible(page)
            dismiss_blocking_modals(page)

    total_results, _ = get_listing_metadata(page)
    if total_results:
        log(f"[Categoria: {category.display_name}] Total estimado: {total_results} produtos.")
    else:
        log(f"[Categoria: {category.display_name}] Total estimado indisponível; a paginar até esgotar resultados.")

    stagnant_rounds = 0
    pages_processed = 0

    while True:
        if start_offset > 0 or "start=" in page.url:
            listing_url = listing_url_for_offset(category.listing_url, start_offset)
            goto_with_retry(
                page,
                listing_url,
                operation_name=f"paginação {category.display_name} offset {start_offset}",
                settle_ms=1500,
            )
            accept_cookies_if_visible(page)
            dismiss_blocking_modals(page)

        previous_count = len(collected)

        while click_load_more_if_visible(page):
            page_products = extract_products_from_listing(page)
            for product in page_products:
                product_url = normalize_product_url(product["link"])
                collected[product["pid"]] = {
                    "pid": product["pid"],
                    "name": product.get("name", ""),
                    "unit_price": product.get("unitPrice", ""),
                    "list_price": product.get("listPrice", ""),
                    "sales_price": product.get("salesPrice", ""),
                    "url": product_url,
                }
            if max_products and len(collected) >= max_products:
                break

        page_products = extract_products_from_listing(page)
        for product in page_products:
            product_url = normalize_product_url(product["link"])
            collected[product["pid"]] = {
                "pid": product["pid"],
                "name": product.get("name", ""),
                "unit_price": product.get("unitPrice", ""),
                "list_price": product.get("listPrice", ""),
                "sales_price": product.get("salesPrice", ""),
                "url": product_url,
            }

        pages_processed += 1
        log(
            f"[Categoria: {category.display_name}] Página {pages_processed}"
            f"{f'/{max_pages}' if max_pages else ''} (offset {start_offset}) — "
            f"{len(collected)} produtos únicos recolhidos."
        )

        if progress is not None:
            progress["in_progress"] = {
                "category": category.display_name,
                "phase": "collecting",
                "collection_offset": start_offset,
                "partial_products": list(collected.values()),
            }
            save_progress(progress)

        if max_products and len(collected) >= max_products:
            break
        if max_pages and pages_processed >= max_pages:
            break

        if len(collected) == previous_count:
            stagnant_rounds += 1
        else:
            stagnant_rounds = 0

        _, visible_to = get_listing_metadata(page)
        if total_results and len(collected) >= total_results:
            break
        if total_results and visible_to >= total_results:
            break
        if stagnant_rounds >= 2:
            break

        start_offset += PAGE_SIZE
        if start_offset > 20_000:
            break

    products = list(collected.values())
    if random_sample:
        random.shuffle(products)
        log(f"[Categoria: {category.display_name}] Amostra aleatória aplicada sobre {len(products)} produtos.")
    if max_products:
        products = products[:max_products]
        if random_sample:
            log(f"[Categoria: {category.display_name}] Selecionados {len(products)} produtos aleatórios.")

    if progress is not None:
        progress["in_progress"] = {
            "category": category.display_name,
            "phase": "extracting",
            "products": products,
        }
        save_progress(progress)

    log(
        f"[Categoria: {category.display_name}] Fase 1 concluída — "
        f"{len(products)} produtos únicos recolhidos."
    )
    return products


def extract_json_ld(page: Page) -> tuple[dict, dict]:
    """Extrai os blocos Product e BreadcrumbList do JSON-LD."""
    data = page.evaluate(
        """() => {
            const scripts = [...document.querySelectorAll('script[type="application/ld+json"]')];
            let product = {};
            let breadcrumb = {};
            for (const script of scripts) {
                try {
                    const parsed = JSON.parse(script.textContent);
                    if (parsed && parsed["@type"] === "Product") {
                        product = parsed;
                    }
                    if (parsed && parsed["@type"] === "BreadcrumbList") {
                        breadcrumb = parsed;
                    }
                } catch (error) {
                    continue;
                }
            }
            return { product, breadcrumb };
        }"""
    )
    return data.get("product", {}), data.get("breadcrumb", {})


def normalize_label(value: str) -> str:
    """Normaliza texto para comparação (sem acentos/case/espços extra)."""
    normalized = value.strip().lower()
    replacements = {
        "á": "a", "à": "a", "â": "a", "ã": "a",
        "é": "e", "ê": "e",
        "í": "i",
        "ó": "o", "ô": "o", "õ": "o",
        "ú": "u", "ç": "c",
    }
    for source, target in replacements.items():
        normalized = normalized.replace(source, target)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized


def labels_match(left: str, right: str) -> bool:
    """Verifica se dois rótulos representam o mesmo produto/categoria."""
    left_norm = normalize_label(left)
    right_norm = normalize_label(right)
    if not left_norm or not right_norm:
        return False
    return left_norm == right_norm or left_norm in right_norm or right_norm in left_norm


def breadcrumb_names_from_json_ld(breadcrumb_ld: dict) -> list[str]:
    """Extrai a lista ordenada de nomes do breadcrumb."""
    items = breadcrumb_ld.get("itemListElement", [])
    names: list[str] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        entry = item.get("item", {})
        if isinstance(entry, dict):
            name = str(entry.get("name", "")).strip()
        else:
            name = str(item.get("name", "")).strip()
        if name:
            names.append(name)
    return names


def breadcrumb_from_json_ld(breadcrumb_ld: dict) -> str:
    """Constrói o breadcrumb completo a partir do JSON-LD."""
    return " -> ".join(breadcrumb_names_from_json_ld(breadcrumb_ld))


def parse_category_hierarchy(breadcrumb_names: list[str], product_name: str) -> dict[str, str]:
    """
    Interpreta o breadcrumb do Auchan na taxonomia de 4 níveis:
    Categoria -> Sub Categoria -> Família -> Sub Família.

    O último elemento do breadcrumb é ignorado quando corresponde ao nome
    do produto (página atual), pois não faz parte da hierarquia comercial.
    """
    levels = [name.strip() for name in breadcrumb_names if name.strip()]

    if levels and product_name and labels_match(levels[-1], product_name):
        levels = levels[:-1]

    hierarchy = {
        level_name: na(levels[index] if index < len(levels) else None)
        for index, level_name in enumerate(CATEGORY_LEVELS)
    }
    hierarchy["Caminho Categorias"] = " -> ".join(levels) if levels else "N/A"
    return hierarchy


def empty_product_record(product: dict[str, str], category_name: str) -> dict[str, str]:
    """Registo mínimo quando a extração detalhada falha."""
    record = {level_name: "N/A" for level_name in CATEGORY_LEVELS}
    record.update(
        {
            "Caminho Categorias": "N/A",
            "Descrição do Produto": na(product.get("name")),
            "EAN / Referência": na(product.get("pid")),
            "Preço Regular": na(product.get("sales_price") or product.get("list_price")),
            "Preço Promocional": "N/A",
            "Preço ao Quilo/Litro": na(product.get("unit_price")),
            "Quantidade da Embalagem": "N/A",
            "URL": product.get("url", "N/A"),
            "Categoria Alvo": category_name,
        }
    )
    return record


def extract_quantity(name: str, page_text: str) -> str:
    """Extrai a quantidade da embalagem a partir do nome ou da página."""
    for source in (name, page_text):
        match = QUANTITY_PATTERN.search(source)
        if match:
            return match.group(1).strip()

    capacity_match = re.search(
        r"(\d+(?:[.,]\d+)?\s*(?:KG|G|GR|ML|CL|L|LT|UN))",
        page_text,
        re.IGNORECASE,
    )
    if capacity_match:
        return capacity_match.group(1).strip()
    return ""


def extract_product_details(
    page: Page,
    product: dict[str, str],
    category_name: str,
) -> dict[str, str]:
    """Visita a página individual do produto e extrai os campos solicitados."""
    product_url = product["url"]

    def _load_product_page() -> None:
        goto_with_retry(
            page,
            product_url,
            operation_name=f"produto {product.get('pid', product_url)}",
            settle_ms=900,
        )
        accept_cookies_if_visible(page)
        dismiss_blocking_modals(page)

    retry_for_duration(f"extração {product_url}", _load_product_page)

    product_ld, breadcrumb_ld = extract_json_ld(page)

    description = ""
    try:
        description = page.locator("main h1, h1").first.inner_text(timeout=10_000).strip()
    except Exception:
        description = str(product.get("name", "")).strip()

    if not description:
        description = str(product_ld.get("name", "")).strip()

    breadcrumb_names = breadcrumb_names_from_json_ld(breadcrumb_ld)
    if not breadcrumb_names:
        try:
            breadcrumb_names = [
                text.strip()
                for text in page.locator(
                    '.breadcrumb a, .auc-breadcrumb a, [class*="breadcrumb"] a'
                ).all_inner_texts()
                if text.strip()
            ]
        except Exception:
            breadcrumb_names = []

    category_hierarchy = parse_category_hierarchy(breadcrumb_names, description)

    ean = str(product_ld.get("gtin13") or product_ld.get("gtin") or "").strip()
    if not ean:
        try:
            page_text = page.locator("body").inner_text(timeout=5_000)
            ean_match = re.search(r"EAN\s*[:\.]?\s*(\d{8,14})", page_text, re.IGNORECASE)
            if ean_match:
                ean = ean_match.group(1)
        except Exception:
            ean = ""

    reference = ean or str(product_ld.get("sku") or product.get("pid") or extract_product_id(product_url)).strip()

    list_price = ""
    sales_price = ""
    price_data = page.evaluate(
        """() => {
            const listNode = document.querySelector(
                ".strike-through.value, .auc-price__stricked .value, s .value"
            );
            const salesNode = document.querySelector(".sales .value, .auc-price__sales .value");
            const formatFromContent = (node) => {
                if (!node) {
                    return "";
                }
                const content = node.getAttribute("content");
                if (content) {
                    const numeric = Number(content);
                    if (!Number.isNaN(numeric)) {
                        return numeric.toFixed(2).replace(".", ",");
                    }
                }
                return (node.innerText || "").replace(/\\s+/g, " ").trim();
            };
            return {
                listPrice: formatFromContent(listNode),
                salesPrice: formatFromContent(salesNode),
            };
        }"""
    )
    list_price = str(price_data.get("listPrice", "")).strip()
    sales_price = str(price_data.get("salesPrice", "")).strip()

    if not list_price:
        list_price = str(product.get("list_price", "")).strip()
    if not sales_price:
        sales_price = str(product.get("sales_price", "")).strip()

    list_price = re.sub(r"(?i)^price reduced from\s*", "", list_price).strip()
    sales_price = re.sub(r"(?i)\s*to\s*$", "", sales_price).strip()

    if list_price and sales_price:
        regular_price = list_price
        promotional_price = sales_price
    elif sales_price:
        regular_price = sales_price
        promotional_price = ""
    else:
        regular_price = str(product.get("sales_price") or product.get("list_price") or "").strip()
        promotional_price = ""

    unit_price = ""
    try:
        unit_price = page.locator(".auc-measures--price-per-unit").first.inner_text(timeout=3_000).strip()
    except Exception:
        unit_price = str(product.get("unit_price", "")).strip()

    quantity = ""
    try:
        page_text = page.locator("body").inner_text(timeout=5_000)
        quantity = extract_quantity(description, page_text)
    except Exception:
        quantity = extract_quantity(description, "")

    return {
        **category_hierarchy,
        "Descrição do Produto": na(description),
        "EAN / Referência": na(reference),
        "Preço Regular": na(regular_price),
        "Preço Promocional": na(promotional_price) if promotional_price else "N/A",
        "Preço ao Quilo/Litro": na(unit_price),
        "Quantidade da Embalagem": na(quantity),
        "URL": product_url,
        "Categoria Alvo": category_name,
    }


def export_to_excel(records: list[dict[str, str]], output_file: str) -> None:
    """Exporta os registos recolhidos para Excel com colunas formatadas."""
    dataframe = pd.DataFrame(records, columns=EXCEL_COLUMNS)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Produtos")
        worksheet = writer.sheets["Produtos"]
        column_widths = {
            "A": 22,
            "B": 22,
            "C": 28,
            "D": 28,
            "E": 55,
            "F": 45,
            "G": 18,
            "H": 14,
            "I": 16,
            "J": 18,
            "K": 22,
            "L": 60,
            "M": 18,
        }
        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width


def process_category(
    page: Page,
    category: CategoryTarget,
    progress: dict,
    output_file: str,
    *,
    max_products: int | None = None,
    max_pages: int | None = None,
    random_sample: bool = False,
    use_progress: bool = True,
) -> None:
    """Processa uma categoria completa: listagem + detalhe de cada produto."""
    category_name = category.display_name
    if use_progress and category_name in progress.get("completed_categories", []):
        log(f"[Categoria: {category_name}] Já concluída anteriormente — a saltar.")
        return

    progress_state = progress if use_progress else None
    products = collect_all_category_products(
        page,
        category,
        navigate_via_menu_flag=True,
        max_products=max_products,
        max_pages=max_pages,
        random_sample=random_sample,
        progress=progress_state,
    )
    if not products:
        log(f"[Categoria: {category_name}] Nenhum produto encontrado.")
        if use_progress:
            completed = list(progress.get("completed_categories", []))
            if category_name not in completed:
                completed.append(category_name)
            progress["completed_categories"] = completed
            progress["in_progress"] = None
            save_progress(progress)
        return

    records: list[dict[str, str]] = progress.setdefault("records", [])
    processed_urls = set(progress.get("processed_urls", [])) if use_progress else set()

    pending_products = [product for product in products if product.get("url") not in processed_urls]
    if use_progress and len(pending_products) < len(products):
        log(
            f"[Categoria: {category_name}] Retomando extração — "
            f"{len(pending_products)} produtos pendentes de {len(products)}."
        )

    log(f"\n[Categoria: {category_name}] Fase 2 — Extração detalhada de produtos.")
    total = len(pending_products)

    for index, product in enumerate(pending_products, start=1):
        product_name = str(product.get("name", "")).strip()
        product_url = product.get("url", "")

        log_product_progress(
            category_name,
            index,
            total,
            status="processing",
            product_name=product_name,
        )
        try:
            product_data = extract_product_details(page, product, category_name)
            records.append(product_data)
            if use_progress and product_url:
                processed_urls.add(product_url)
                progress["records"] = records
                progress["processed_urls"] = sorted(processed_urls)
                save_progress(progress)
                maybe_checkpoint(records, output_file, progress)

            price_info = product_data["Preço Regular"]
            if product_data.get("Preço Promocional") not in ("N/A", ""):
                price_info = f"{product_data['Preço Regular']} (promo: {product_data['Preço Promocional']})"
            log_product_progress(
                category_name,
                index,
                total,
                status="done",
                product_name=product_data["Descrição do Produto"],
                details=(
                    f"{product_data['Caminho Categorias']} "
                    f"| EAN: {product_data['EAN / Referência']} "
                    f"| Preço: {price_info}"
                ),
            )
        except Exception as error:
            log_product_progress(
                category_name,
                index,
                total,
                status="error",
                details=str(error),
            )
            records.append(empty_product_record(product, category_name))
            if use_progress and product_url:
                processed_urls.add(product_url)
                progress["records"] = records
                progress["processed_urls"] = sorted(processed_urls)
                save_progress(progress)
                maybe_checkpoint(records, output_file, progress)

        if index < total:
            human_delay(0.8, 1.8)

    if use_progress:
        completed = list(progress.get("completed_categories", []))
        if category_name not in completed:
            completed.append(category_name)
        progress["completed_categories"] = completed
        progress["records"] = records
        progress["processed_urls"] = sorted(processed_urls)
        progress["in_progress"] = None
        save_progress(progress)
        export_to_excel(records, output_file)

    log(f"[Categoria: {category_name}] Concluída — {total} produtos processados nesta execução.")


FRESCOS_CATEGORIES = frozenset({"Charcutaria", "Queijaria"})


def resolve_categories(only: str | None) -> tuple[CategoryTarget, ...]:
    """Filtra as categorias a processar."""
    if not only:
        return CATEGORY_TARGETS

    normalized = only.strip().lower()
    if normalized == "frescos":
        return tuple(category for category in CATEGORY_TARGETS if category.display_name in FRESCOS_CATEGORIES)

    selected = {name.strip().lower() for name in only.split(",") if name.strip()}
    filtered = tuple(
        category for category in CATEGORY_TARGETS if category.display_name.lower() in selected
    )
    if not filtered:
        available = ", ".join(category.display_name for category in CATEGORY_TARGETS)
        raise SystemExit(f"Categorias inválidas em --only. Opções: frescos, {available}")
    return filtered


def parse_args() -> argparse.Namespace:
    """Interpreta argumentos de linha de comandos."""
    parser = argparse.ArgumentParser(description="Scraping de produtos Auchan Portugal.")
    parser.add_argument(
        "--test",
        action="store_true",
        help="Executa apenas uma amostra reduzida (2 produtos por categoria) para validação.",
    )
    parser.add_argument(
        "--only",
        help='Processa apenas categorias específicas (ex.: "frescos" ou "Charcutaria,Queijaria").',
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        help="Limita o número de páginas de listagem a recolher por categoria.",
    )
    parser.add_argument(
        "--max-products",
        type=int,
        help="Limita o número de produtos a processar por categoria.",
    )
    parser.add_argument(
        "--random",
        action="store_true",
        help="Seleciona aleatoriamente os produtos quando --max-products está definido.",
    )
    parser.add_argument(
        "--output",
        default=OUTPUT_FILE,
        help=f"Ficheiro Excel de saída (predefinido: {OUTPUT_FILE}).",
    )
    parser.add_argument(
        "--reset-progress",
        action="store_true",
        help="Apaga o ficheiro de progresso e reinicia a extração do zero.",
    )
    return parser.parse_args()


def should_use_progress(args: argparse.Namespace) -> bool:
    """Determina se a retoma automática deve estar ativa."""
    if args.reset_progress:
        return False
    return not args.test and not args.max_pages and not args.max_products and not args.random


def main() -> None:
    args = parse_args()
    max_products = args.max_products if args.max_products else (2 if args.test else None)
    max_pages = args.max_pages
    random_sample = args.random
    categories = resolve_categories(args.only)
    use_progress = should_use_progress(args)

    if args.reset_progress and Path(PROGRESS_FILE).exists():
        Path(PROGRESS_FILE).unlink()
        log(f"Progresso anterior removido: {PROGRESS_FILE}")

    progress = load_progress(args.output) if use_progress else empty_progress(args.output)
    progress["output_file"] = args.output

    if use_progress and progress.get("records"):
        log(
            f"Retoma automática ativa — {len(progress['records'])} produtos já guardados, "
            f"{len(progress.get('completed_categories', []))} categorias concluídas."
        )
        if progress.get("completed_categories"):
            log(
                "Categorias concluídas: "
                + ", ".join(progress["completed_categories"])
            )

    if args.test and not args.max_products:
        log("Modo teste ativo — apenas 2 produtos por categoria serão processados.")
    if max_products:
        log(f"Limite de produtos: {max_products} por categoria.")
    if random_sample:
        log("Amostragem aleatória ativa.")
    if max_pages:
        log(f"Limite de paginação: {max_pages} página(s) por categoria.")
    if use_progress:
        log(
            f"Checkpoints: Excel parcial a cada {CHECKPOINT_EVERY} produtos "
            f"e retry automático durante {RETRY_TIMEOUT_SECONDS // 60} minutos."
        )
    if args.only:
        log(
            "Categorias selecionadas: "
            + ", ".join(category.display_name for category in categories)
        )

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
            for category in categories:
                log(f"\n{'=' * 70}\n[Categoria: {category.display_name}] A iniciar extração.\n{'=' * 70}")
                process_category(
                    page,
                    category,
                    progress,
                    args.output,
                    max_products=max_products,
                    max_pages=max_pages,
                    random_sample=random_sample,
                    use_progress=use_progress,
                )

                if not args.test and not max_pages:
                    human_delay(1.5, 3.0)

            records = progress.get("records", [])
            if not records:
                log("Nenhum produto recolhido. O ficheiro Excel não será gerado.")
                return

            export_to_excel(records, args.output)
            if use_progress:
                save_progress(progress)
            log(f"\nConcluído. {len(records)} produtos exportados para '{args.output}'.")
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
