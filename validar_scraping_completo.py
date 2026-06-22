"""Validação rápida do scraper completo (não executa o scraping total)."""
from playwright.sync_api import sync_playwright

from scraping_action_teste import (
    BASE_URL,
    accept_cookies_if_visible,
    collect_product_urls_from_page_1,
    discover_main_categories,
    export_to_excel,
    extract_product_data,
    open_category_listing_page,
)
from scraping_action_completo import OUTPUT_FILE, empty_progress, load_progress, save_progress


def main() -> None:
    print("=== Validação scraping_action_completo.py ===\n")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page(locale="pt-PT")

        page.goto(BASE_URL, wait_until="networkidle", timeout=90_000)
        accept_cookies_if_visible(page)

        categories = discover_main_categories(page)
        assert len(categories) == 15, f"Esperadas 15 categorias, obtidas {len(categories)}"
        print(f"OK — {len(categories)} categorias detetadas")

        first = categories[0]
        open_category_listing_page(page, first["name"])
        urls = collect_product_urls_from_page_1(page)
        assert urls, "Nenhuma URL encontrada na página 1"
        print(f"OK — {len(urls)} URLs na página 1 de '{first['name']}'")

        product = extract_product_data(page, urls[0], first["name"])
        required = {
            "Categoria Principal",
            "Sub-categoria",
            "Marca",
            "Descrição / Nome do artigo",
            "Preço Regular",
            "Preço Promocional",
            "URL",
        }
        assert required.issubset(product.keys()), f"Campos em falta: {required - set(product.keys())}"
        assert product["Categoria Principal"] == first["name"]
        print(f"OK — produto extraído: {product['Descrição / Nome do artigo'][:50]}")
        print(f"     Sub-categoria: {product['Sub-categoria'] or '—'}")

        test_file = "Scraping_Action_Validacao.xlsx"
        export_to_excel([product], test_file)
        print(f"OK — Excel de teste criado: {test_file}")

        progress = empty_progress()
        progress["categories"] = categories
        save_progress(progress)
        loaded = load_progress()
        assert len(loaded["categories"]) == 15
        print("OK — ficheiro de progresso lido/escrito corretamente")

        browser.close()

    print("\n=== TUDO OK — podes executar: ===")
    print("PYTHONUNBUFFERED=1 python3 scraping_action_completo.py")


if __name__ == "__main__":
    main()
