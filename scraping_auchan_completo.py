"""
Robô completo — extração de todos os produtos Auchan Portugal (todas as categorias).

SETUP (executar no terminal antes da primeira utilização):
    pip install playwright pandas openpyxl
    python3 -m playwright install chromium

EXECUÇÃO:
    PYTHONUNBUFFERED=1 python3 scraping_auchan_completo.py
    PYTHONUNBUFFERED=1 python3 scraping_auchan_completo.py --test
    PYTHONUNBUFFERED=1 python3 scraping_auchan_completo.py --discover-only
    PYTHONUNBUFFERED=1 python3 scraping_auchan_completo.py --max-categories 5 --max-products 3

NOTAS:
    - Descobre automaticamente todas as categorias de listagem no menu principal.
    - Reforça a descoberta com subcategorias encontradas nas páginas de secção.
    - Percorre cada categoria com paginação completa e visita cada produto.
    - Deduplica produtos globalmente (o mesmo artigo só é extraído uma vez).
    - Grava checkpoint Excel a cada 100 produtos e retoma após interrupção.
    - Exporta o resultado final para Scraping_Auchan_Completo.xlsx
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from urllib.parse import urlparse

import scraping_auchan as auchan
from playwright.sync_api import Page, sync_playwright

OUTPUT_FILE = "Scraping_Auchan_Completo.xlsx"
PROGRESS_FILE = "scraping_auchan_completo_progress.json"

EXCLUDE_MENU_IDS = frozenset(
    {
        "menu-category-folhetos",
        "menu-category-Promoções exclusivas auchan",
    }
)
EXCLUDE_PATH_FRAGMENTS = (
    "/folhetos",
    "/promocoes",
    "/campanhas",
    "/clube-auchan",
    "/search",
    "/pesquisa",
)

auchan.PROGRESS_FILE = PROGRESS_FILE


def log(message: str) -> None:
    auchan.log(message)


def empty_progress(output_file: str = OUTPUT_FILE) -> dict:
    return {
        "output_file": output_file,
        "categories": [],
        "completed_categories": [],
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

    return progress


def save_progress(progress: dict) -> None:
    with Path(PROGRESS_FILE).open("w", encoding="utf-8") as progress_file:
        json.dump(progress, progress_file, ensure_ascii=False, indent=2)


def maybe_checkpoint(records: list[dict[str, str]], output_file: str, progress: dict) -> None:
    if not records or len(records) % auchan.CHECKPOINT_EVERY != 0:
        return

    auchan.export_to_excel(records, output_file)
    save_progress(progress)
    log(
        f"[Checkpoint] {len(records)} produtos guardados em '{output_file}' "
        f"e progresso atualizado em '{PROGRESS_FILE}'."
    )


def normalize_listing_path(url: str) -> str:
    parsed = urlparse(auchan.absolute_url(url))
    path = parsed.path.rstrip("/")
    if not path.endswith("/"):
        path = f"{path}/"
    return path


def is_product_listing_url(url: str) -> bool:
    parsed = urlparse(auchan.absolute_url(url))
    path = parsed.path.rstrip("/")
    if not path.startswith("/pt/"):
        return False
    if path.endswith(".html"):
        return False

    lower_path = path.lower()
    if any(fragment in lower_path for fragment in EXCLUDE_PATH_FRAGMENTS):
        return False

    segments = [segment for segment in path.split("/") if segment]
    return len(segments) >= 2


def path_to_display_name(path: str) -> str:
    segments = [segment for segment in path.strip("/").split("/") if segment and segment != "pt"]
    labels = [segment.replace("-", " ").title() for segment in segments]
    return " > ".join(labels) if labels else path


def category_key(category: dict[str, str]) -> str:
    """Identificador único da categoria (URL de listagem)."""
    return normalize_listing_path(category["listing_url"])


def category_dict_to_target(category: dict[str, str]) -> auchan.CategoryTarget:
    listing_path = category["listing_url"]
    if not listing_path.startswith("/"):
        listing_path = urlparse(listing_path).path
    if not listing_path.endswith("/"):
        listing_path = f"{listing_path}/"

    menu_path = tuple(part.strip() for part in category.get("menu_path", "").split(" > ") if part.strip())
    if not menu_path and category.get("top_category"):
        menu_path = (category["top_category"],)

    return auchan.CategoryTarget(
        display_name=category["name"],
        menu_path=menu_path,
        listing_url=listing_path,
    )


def discover_menu_listing_urls(page: Page) -> dict[str, dict[str, str]]:
    """Descobre URLs de listagem a partir dos painéis do menu principal."""
    top_categories = page.evaluate(
        """() => [...document.querySelectorAll('[id^="menu-category-"]')]
            .map((element) => ({
                id: element.id,
                text: (element.innerText || element.textContent || '').trim().split('\\n')[0],
            }))
            .filter((item) => item.text)"""
    )

    discovered: dict[str, dict[str, str]] = {}

    for top_category in top_categories:
        if top_category["id"] in EXCLUDE_MENU_IDS:
            continue

        page.evaluate(
            """(menuId) => {
                document.querySelectorAll('[id^="menu-category-"]').forEach((element) => {
                    element.dispatchEvent(new MouseEvent('mouseleave', { bubbles: true }));
                });
                const menuElement = document.getElementById(menuId);
                if (!menuElement) {
                    return;
                }
                menuElement.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
                menuElement.click();
            }""",
            top_category["id"],
        )
        page.wait_for_timeout(700)

        links = page.evaluate(
            """(menuId) => {
                const menuElement = document.getElementById(menuId);
                if (!menuElement) {
                    return [];
                }
                const panel = menuElement
                    .closest('li, .nav-item, .dropdown')
                    ?.querySelector('.dropdown-menu, .submenu, .menu-dropdown');
                const scope = panel || menuElement.parentElement?.parentElement || document;
                const seen = new Set();
                const results = [];
                for (const anchor of scope.querySelectorAll('a[href]')) {
                    const href = anchor.href;
                    if (!href || seen.has(href)) {
                        continue;
                    }
                    seen.add(href);
                    const text = (anchor.innerText || anchor.textContent || '')
                        .trim()
                        .replace(/\\s+/g, ' ');
                    if (!text) {
                        continue;
                    }
                    results.push({ text, href });
                }
                return results;
            }""",
            top_category["id"],
        )

        for link in links:
            if not is_product_listing_url(link["href"]):
                continue

            listing_path = normalize_listing_path(link["href"])
            if listing_path in discovered:
                continue

            link_name = re.sub(r"\s+", " ", link["text"]).strip()
            display_name = path_to_display_name(listing_path)
            discovered[listing_path] = {
                "name": display_name,
                "listing_url": listing_path,
                "top_category": top_category["text"],
                "menu_path": top_category["text"],
                "source": "menu",
                "link_label": link_name,
            }

    return discovered


def discover_child_listing_urls(page: Page, parent_path: str) -> dict[str, dict[str, str]]:
    """Encontra subcategorias cujo URL é filho do caminho indicado."""
    normalized_parent = parent_path.rstrip("/")
    child_links = page.evaluate(
        """(parentPath) => {
            const results = [];
            const seen = new Set();
            for (const anchor of document.querySelectorAll('a[href]')) {
                const href = anchor.href;
                if (!href || seen.has(href)) {
                    continue;
                }
                let path = '';
                try {
                    path = new URL(href).pathname.replace(/\\/$/, '');
                } catch (error) {
                    continue;
                }
                if (!path.startsWith(parentPath + '/') || path === parentPath) {
                    continue;
                }
                if (path.endsWith('.html')) {
                    continue;
                }
                seen.add(href);
                const text = (anchor.innerText || anchor.textContent || '')
                    .trim()
                    .replace(/\\s+/g, ' ');
                if (!text || text.length > 80) {
                    continue;
                }
                results.push({ text, href: path + '/', path });
            }
            return results;
        }""",
        normalized_parent,
    )

    discovered: dict[str, dict[str, str]] = {}
    for link in child_links:
        listing_path = normalize_listing_path(link["href"])
        if not is_product_listing_url(listing_path):
            continue

        link_name = link["text"].strip()
        display_name = path_to_display_name(listing_path)
        discovered[listing_path] = {
            "name": display_name,
            "listing_url": listing_path,
            "top_category": path_to_display_name(listing_path).split(" > ")[0],
            "menu_path": path_to_display_name(listing_path),
            "source": "child",
            "link_label": link_name,
        }

    return discovered


def section_root_paths(listing_paths: list[str]) -> list[str]:
    roots = set()
    for listing_path in listing_paths:
        segments = [segment for segment in listing_path.strip("/").split("/") if segment and segment != "pt"]
        if len(segments) >= 2:
            roots.add(f"/pt/{segments[0]}/")
    return sorted(roots)


def discover_all_categories(page: Page) -> list[dict[str, str]]:
    """Descobre todas as categorias de listagem do site."""
    log("A descobrir categorias no menu principal...")
    auchan.goto_with_retry(page, auchan.BASE_URL, operation_name="página inicial Auchan", settle_ms=2000)
    auchan.accept_cookies_if_visible(page)
    auchan.dismiss_blocking_modals(page)

    discovered = discover_menu_listing_urls(page)
    log(f"  {len(discovered)} categorias encontradas no menu.")

    log("A reforçar descoberta com subcategorias nas páginas de secção...")
    roots = section_root_paths(list(discovered.keys()))
    for index, root_path in enumerate(roots, start=1):
        log(f"  Secção {index}/{len(roots)}: {root_path}")
        auchan.goto_with_retry(
            page,
            auchan.absolute_url(root_path),
            operation_name=f"secção {root_path}",
            settle_ms=2500,
        )
        auchan.accept_cookies_if_visible(page)
        auchan.dismiss_blocking_modals(page)
        page.wait_for_timeout(1200)

        child_categories = discover_child_listing_urls(page, root_path.rstrip("/"))
        added = 0
        for listing_path, category in child_categories.items():
            if listing_path in discovered:
                continue
            discovered[listing_path] = category
            added += 1
        if added:
            log(f"    +{added} subcategorias adicionais.")

    categories = sorted(
        discovered.values(),
        key=lambda item: item["listing_url"],
    )
    log(f"Total de categorias de listagem: {len(categories)}")
    return categories


def ensure_categories(page: Page, progress: dict) -> list[dict[str, str]]:
    if progress.get("categories"):
        categories = list(progress["categories"])
        log(f"Categorias carregadas do progresso ({len(categories)}).")
        return categories

    categories = discover_all_categories(page)
    progress["categories"] = categories
    save_progress(progress)

    log("Primeiras categorias:")
    for index, category in enumerate(categories[:10], start=1):
        log(f"  {index}. {category['name']} — {category['listing_url']}")
    if len(categories) > 10:
        log(f"  ... e mais {len(categories) - 10} categorias.")
    return categories


def process_category(
    page: Page,
    category: dict[str, str],
    progress: dict,
    output_file: str,
    *,
    max_products: int | None = None,
    max_pages: int | None = None,
    random_sample: bool = False,
    use_progress: bool = True,
) -> None:
    """Processa uma categoria completa com deduplicação global de produtos."""
    target = category_dict_to_target(category)
    category_name = target.display_name
    listing_key = category_key(category)

    if use_progress and listing_key in progress.get("completed_categories", []):
        log(f"[Categoria: {category_name}] Já concluída anteriormente — a saltar.")
        return

    category_index = progress["categories"].index(category) + 1
    total_categories = len(progress["categories"])
    log(
        f"\n{'=' * 70}\n"
        f"[Categoria {category_index}/{total_categories}: {category_name}]\n"
        f"URL: {target.listing_url}\n"
        f"{'=' * 70}"
    )

    progress_state = progress if use_progress else None
    products = auchan.collect_all_category_products(
        page,
        target,
        navigate_via_menu_flag=False,
        max_products=max_products,
        max_pages=max_pages,
        random_sample=random_sample,
        progress=progress_state,
    )

    if not products:
        log(f"[Categoria: {category_name}] Nenhum produto encontrado.")
        if use_progress:
            completed = list(progress.get("completed_categories", []))
            if listing_key not in completed:
                completed.append(listing_key)
            progress["completed_categories"] = completed
            progress["in_progress"] = None
            save_progress(progress)
        return

    records: list[dict[str, str]] = progress.setdefault("records", [])
    processed_urls = set(progress.get("processed_urls", [])) if use_progress else set()

    pending_products = [product for product in products if product.get("url") not in processed_urls]
    skipped = len(products) - len(pending_products)
    if skipped:
        log(
            f"[Categoria: {category_name}] {skipped} produto(s) já extraído(s) noutra categoria — a saltar."
        )

    if use_progress and len(pending_products) < len(products) and not skipped:
        log(
            f"[Categoria: {category_name}] Retomando extração — "
            f"{len(pending_products)} produtos pendentes de {len(products)}."
        )

    log(f"\n[Categoria: {category_name}] Fase 2 — Extração detalhada de produtos.")
    total = len(pending_products)

    for index, product in enumerate(pending_products, start=1):
        product_name = str(product.get("name", "")).strip()
        product_url = product.get("url", "")

        auchan.log_product_progress(
            category_name,
            index,
            total,
            status="processing",
            product_name=product_name,
        )
        try:
            product_data = auchan.extract_product_details(page, product, category_name)
            records.append(product_data)
            if use_progress and product_url:
                processed_urls.add(product_url)
                progress["records"] = records
                progress["processed_urls"] = sorted(processed_urls)
                save_progress(progress)
                maybe_checkpoint(records, output_file, progress)

            price_info = product_data["Preço Regular"]
            if product_data.get("Preço Promocional") not in ("N/A", ""):
                price_info = (
                    f"{product_data['Preço Regular']} "
                    f"(promo: {product_data['Preço Promocional']})"
                )
            auchan.log_product_progress(
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
            auchan.log_product_progress(
                category_name,
                index,
                total,
                status="error",
                details=str(error),
            )
            records.append(auchan.empty_product_record(product, category_name))
            if use_progress and product_url:
                processed_urls.add(product_url)
                progress["records"] = records
                progress["processed_urls"] = sorted(processed_urls)
                save_progress(progress)
                maybe_checkpoint(records, output_file, progress)

        if index < total:
            auchan.human_delay(0.8, 1.8)

    if use_progress:
        completed = list(progress.get("completed_categories", []))
        if listing_key not in completed:
            completed.append(listing_key)
        progress["completed_categories"] = completed
        progress["records"] = records
        progress["processed_urls"] = sorted(processed_urls)
        progress["in_progress"] = None
        save_progress(progress)
        auchan.export_to_excel(records, output_file)

    log(
        f"[Categoria: {category_name}] Concluída — "
        f"{total} produtos processados nesta execução."
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scraping completo de produtos Auchan Portugal (todas as categorias)."
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="Executa amostra reduzida (3 categorias, 2 produtos cada).",
    )
    parser.add_argument(
        "--discover-only",
        action="store_true",
        help="Apenas descobre e lista categorias, sem extrair produtos.",
    )
    parser.add_argument(
        "--max-categories",
        type=int,
        help="Limita o número de categorias a processar.",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        help="Limita o número de páginas de listagem por categoria.",
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
    if args.reset_progress:
        return False
    return (
        not args.test
        and not args.discover_only
        and not args.max_pages
        and not args.max_products
        and not args.max_categories
        and not args.random
    )


def main() -> None:
    args = parse_args()
    max_products = args.max_products if args.max_products else (2 if args.test else None)
    max_pages = args.max_pages
    max_categories = args.max_categories if args.max_categories else (3 if args.test else None)
    random_sample = args.random
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

    if args.test and not args.max_products:
        log("Modo teste ativo — 3 categorias, 2 produtos cada.")
    if max_categories:
        log(f"Limite de categorias: {max_categories}.")
    if max_products:
        log(f"Limite de produtos: {max_products} por categoria.")
    if max_pages:
        log(f"Limite de paginação: {max_pages} página(s) por categoria.")
    if use_progress:
        log(
            f"Checkpoints: Excel parcial a cada {auchan.CHECKPOINT_EVERY} produtos "
            f"e retry automático durante {auchan.RETRY_TIMEOUT_SECONDS // 60} minutos."
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
            categories = ensure_categories(page, progress)

            if args.discover_only:
                log(f"\nDescoberta concluída — {len(categories)} categorias:")
                for index, category in enumerate(categories, start=1):
                    log(
                        f"  {index:4d}. {category['name']} "
                        f"({category['listing_url']}) [{category.get('source', 'menu')}]"
                    )
                save_progress(progress)
                return

            completed = set(progress.get("completed_categories", []))
            pending_categories = [
                category for category in categories if category_key(category) not in completed
            ]

            if max_categories:
                pending_categories = pending_categories[:max_categories]

            if completed and pending_categories:
                log(
                    f"\nRetomando execução. Categorias já concluídas: "
                    f"{len(completed)}/{len(categories)}"
                )

            if not pending_categories:
                records = progress.get("records", [])
                if records:
                    auchan.export_to_excel(records, args.output)
                log(
                    f"\nTodas as categorias já foram processadas. "
                    f"{len(records)} produtos em '{args.output}'."
                )
                return

            log(
                f"\nCategorias pendentes: {len(pending_categories)} "
                f"(de {len(categories)} no total)."
            )

            for category in pending_categories:
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
                    auchan.human_delay(1.0, 2.0)

            records = progress.get("records", [])
            if not records:
                log("Nenhum produto recolhido. O ficheiro Excel não será gerado.")
                return

            auchan.export_to_excel(records, args.output)
            if use_progress:
                save_progress(progress)

            log(
                f"\n{'=' * 70}\n"
                f"EXTRAÇÃO COMPLETA CONCLUÍDA\n"
                f"Categorias processadas: {len(progress.get('completed_categories', []))}/"
                f"{len(categories)}\n"
                f"Total de produtos únicos exportados: {len(records)}\n"
                f"Ficheiro: {args.output}\n"
                f"{'=' * 70}"
            )
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
