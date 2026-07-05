"""
Sincronização semanal do Scraping_Action_Completo.xlsx com o site Action.

SETUP:
    pip install playwright pandas openpyxl
    python3 -m playwright install chromium

EXECUÇÃO:
    PYTHONUNBUFFERED=1 python3 scraping_action_sincronizar.py
    PYTHONUNBUFFERED=1 python3 scraping_action_sincronizar.py --test
    PYTHONUNBUFFERED=1 python3 scraping_action_sincronizar.py --no-git

FUNCIONALIDADES:
    - Copia o Excel base com a data de execução no nome do ficheiro.
    - Varre o site (15 categorias) e compara com o ficheiro existente.
    - Produtos novos: adicionados com letra AZUL.
    - Produtos removidos do site: letra VERMELHA (mantidos no ficheiro).
    - Preços (e restantes campos) atualizados quando diferentes.
    - Progresso artigo-a-artigo no terminal.
    - Grava, faz commit e push a cada 100 artigos processados.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from playwright.sync_api import sync_playwright

from scraping_action_teste import (
    BASE_URL,
    accept_cookies_if_visible,
    collect_all_category_product_urls,
    discover_main_categories,
    extract_product_data,
    human_delay,
    log,
)

SOURCE_FILE = "Scraping_Action_Completo.xlsx"
PROGRESS_FILE = "scraping_action_sincronizar_progress.json"
REPO_ROOT = Path(__file__).resolve().parent
CHECKPOINT_EVERY = 100
GIT_PUSH_BACKOFF_SECONDS = (4, 8, 16, 32)

COLUMNS = [
    "Categoria Principal",
    "Sub-categoria",
    "Marca",
    "Descrição / Nome do artigo",
    "Preço Regular",
    "Preço Promocional",
    "URL",
]

FONT_NOVO = Font(color="0000FF")
FONT_REMOVIDO = Font(color="FF0000")
FONT_NORMAL = Font(color="000000")

_git_checkpoint_enabled = True


def normalize_url(url: str) -> str:
    """Normaliza URLs para comparação consistente."""
    text = str(url or "").strip()
    if not text:
        return ""
    if text.endswith("/"):
        return text
    return text + "/"


def as_text(value) -> str:
    """Converte valores do Excel para texto limpo."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def output_filename_for_today() -> str:
    """Gera o nome do ficheiro de saída com a data de execução."""
    return f"Scraping_Action_Atualizado_{datetime.now().strftime('%Y%m%d')}.xlsx"


def empty_progress(output_file: str) -> dict:
    return {
        "output_file": output_file,
        "source_file": SOURCE_FILE,
        "execution_date": datetime.now().strftime("%Y-%m-%d"),
        "site_scan_done": False,
        "site_urls": [],
        "url_to_category": {},
        "categories": [],
        "removed_marked": False,
        "records": [],
        "row_states": {},
        "checked_update_urls": [],
        "added_new_urls": [],
        "articles_processed": 0,
    }


def load_progress(output_file: str) -> dict:
    progress_path = REPO_ROOT / PROGRESS_FILE
    if not progress_path.exists():
        return empty_progress(output_file)

    with progress_path.open("r", encoding="utf-8") as progress_file:
        progress = json.load(progress_file)

    defaults = empty_progress(output_file)
    for key, default_value in defaults.items():
        if key not in progress:
            progress[key] = default_value

    if progress.get("output_file") != output_file:
        return empty_progress(output_file)

    return progress


def save_progress(progress: dict) -> None:
    with (REPO_ROOT / PROGRESS_FILE).open("w", encoding="utf-8") as progress_file:
        json.dump(progress, progress_file, ensure_ascii=False, indent=2)


def load_records_from_excel(path: Path) -> list[dict[str, str]]:
    """Carrega registos do Excel base."""
    dataframe = pd.read_excel(path)
    records: list[dict[str, str]] = []

    for _, row in dataframe.iterrows():
        record = {column: as_text(row.get(column, "")) for column in COLUMNS}
        if record["URL"]:
            records.append(record)

    return records


def copy_source_workbook(source_file: str, output_file: str) -> None:
    """Copia o ficheiro Excel base para o ficheiro datado."""
    source_path = REPO_ROOT / source_file
    output_path = REPO_ROOT / output_file

    if not source_path.exists():
        raise FileNotFoundError(f"Ficheiro base não encontrado: {source_path}")

    shutil.copy2(source_path, output_path)
    log(f"Cópia criada: {output_file}")


def prices_changed(old: dict[str, str], new: dict[str, str]) -> bool:
    """Verifica se os preços mudaram."""
    return (
        as_text(old.get("Preço Regular")) != as_text(new.get("Preço Regular"))
        or as_text(old.get("Preço Promocional")) != as_text(new.get("Preço Promocional"))
    )


def records_differ(old: dict[str, str], new: dict[str, str]) -> bool:
    """Verifica se algum campo relevante mudou."""
    for column in COLUMNS:
        if as_text(old.get(column)) != as_text(new.get(column)):
            return True
    return False


def save_styled_excel(records: list[dict[str, str]], row_states: dict[str, str], output_file: str) -> None:
    """Grava o Excel aplicando cores: azul=novo, vermelho=removido."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Produtos"
    worksheet.append(COLUMNS)

    for record in records:
        row_values = [as_text(record.get(column, "")) for column in COLUMNS]
        worksheet.append(row_values)

        url = normalize_url(record.get("URL", ""))
        state = row_states.get(url, "activo")
        font = FONT_NOVO if state == "novo" else FONT_REMOVIDO if state == "removido" else FONT_NORMAL

        for cell in worksheet[worksheet.max_row]:
            cell.font = font

    workbook.save(REPO_ROOT / output_file)


def commit_checkpoint_to_repo(output_file: str, product_count: int) -> bool:
    """Commit e push do Excel + progresso como cópia de segurança."""
    if not _git_checkpoint_enabled:
        return False

    files_to_add: list[str] = []
    if (REPO_ROOT / PROGRESS_FILE).exists():
        files_to_add.append(PROGRESS_FILE)
    if (REPO_ROOT / output_file).exists():
        files_to_add.append(output_file)

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
            [
                "git",
                "commit",
                "-m",
                f"Checkpoint Action sync: {product_count} artigos em {output_file}",
            ],
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
                    f"[Git] Backup enviado para origin/{branch_name} "
                    f"({product_count} artigos processados)."
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


def maybe_checkpoint(progress: dict, output_file: str) -> None:
    """Grava Excel e opcionalmente faz commit/push a cada 100 artigos."""
    save_styled_excel(progress["records"], progress["row_states"], output_file)
    save_progress(progress)

    processed = progress["articles_processed"]
    if processed > 0 and processed % CHECKPOINT_EVERY == 0:
        log(f"[Checkpoint] {processed} artigos processados — ficheiro guardado.")
        commit_checkpoint_to_repo(output_file, processed)


def scan_site_products(page, progress: dict, *, test_mode: bool = False) -> set[str]:
    """Varre todas as categorias e devolve URLs atuais no site."""
    log("\n=== FASE 1 — Varredura do site (todas as categorias) ===")

    if progress.get("categories"):
        categories = progress["categories"]
        log(f"Reutilizando lista de {len(categories)} categorias.")
    else:
        log("A descobrir categorias principais no menu Produtos...")
        categories = discover_main_categories(page)
        progress["categories"] = categories
        save_progress(progress)
        log(f"Categorias encontradas ({len(categories)}):")
        for index, category in enumerate(categories, start=1):
            log(f"  {index}. {category['name']} — {category['listing_url']}")

    site_urls: set[str] = set()
    url_to_category: dict[str, str] = dict(progress.get("url_to_category", {}))

    for index, category in enumerate(categories, start=1):
        if test_mode and index > 1:
            break
        category_name = category["name"]
        log(f"\n[Varredura {index}/{len(categories)}] Categoria: {category_name}")
        category_urls = collect_all_category_product_urls(
            page,
            category_name,
            category["listing_url"],
            navigate_via_menu=True,
        )
        for url in category_urls:
            normalized = normalize_url(url)
            site_urls.add(normalized)
            url_to_category[normalized] = category_name

        log(
            f"[Varredura] {category_name} concluída — "
            f"{len(category_urls)} URLs | Total no site: {len(site_urls)}"
        )

    progress["site_urls"] = sorted(site_urls)
    progress["url_to_category"] = url_to_category
    progress["site_scan_done"] = True
    save_progress(progress)

    log(f"\n=== Varredura concluída — {len(site_urls)} produtos no site ===")
    return site_urls


def mark_removed_products(progress: dict, site_urls: set[str]) -> int:
    """Marca produtos que já não existem no site com cor vermelha."""
    if progress.get("removed_marked"):
        return 0

    existing_urls = {normalize_url(record["URL"]) for record in progress["records"]}
    removed_urls = existing_urls - site_urls
    row_states = progress["row_states"]

    for url in removed_urls:
        row_states[url] = "removido"

    progress["row_states"] = row_states
    progress["removed_marked"] = True
    save_progress(progress)

    log(f"\n=== {len(removed_urls)} produtos marcados como REMOVIDOS (vermelho) ===")
    return len(removed_urls)


def initialize_workbook(source_file: str, output_file: str, progress: dict) -> None:
    """Prepara a cópia datada e carrega os registos iniciais."""
    output_path = REPO_ROOT / output_file

    if not output_path.exists():
        copy_source_workbook(source_file, output_file)

    if not progress.get("records"):
        records = load_records_from_excel(REPO_ROOT / source_file)
        row_states = {normalize_url(record["URL"]): "activo" for record in records}
        progress["records"] = records
        progress["row_states"] = row_states
        save_styled_excel(records, row_states, output_file)
        save_progress(progress)
        log(f"Registos carregados do ficheiro base: {len(records)}")


def update_existing_product(
    page,
    url: str,
    records: list[dict[str, str]],
    row_states: dict[str, str],
    url_to_category: dict[str, str],
) -> tuple[bool, str]:
    """Atualiza um produto existente se os dados do site forem diferentes."""
    normalized = normalize_url(url)
    record_index = next(
        (index for index, record in enumerate(records) if normalize_url(record["URL"]) == normalized),
        None,
    )
    if record_index is None:
        return False, "nao_encontrado"

    old_record = records[record_index]
    category_name = url_to_category.get(normalized) or old_record.get("Categoria Principal", "Casa")

    new_record = extract_product_data(page, url, category_name)
    changed = records_differ(old_record, new_record)

    if changed:
        records[record_index] = new_record
        row_states[normalized] = "activo"
        if prices_changed(old_record, new_record):
            detail = (
                f"preço actualizado | Regular: {old_record.get('Preço Regular')} → "
                f"{new_record.get('Preço Regular')} | Promo: "
                f"{old_record.get('Preço Promocional') or '—'} → "
                f"{new_record.get('Preço Promocional') or '—'}"
            )
        else:
            detail = "dados actualizados"
        return True, detail

    return False, "sem_alteracao"


def add_new_product(
    page,
    url: str,
    records: list[dict[str, str]],
    row_states: dict[str, str],
    url_to_category: dict[str, str],
) -> dict[str, str]:
    """Adiciona um produto novo com estado 'novo' (azul)."""
    normalized = normalize_url(url)
    category_name = url_to_category.get(normalized, "Casa")
    new_record = extract_product_data(page, url, category_name)
    records.append(new_record)
    row_states[normalized] = "novo"
    return new_record


def main() -> None:
    parser = argparse.ArgumentParser(description="Sincroniza Scraping_Action_Completo.xlsx com o site Action.")
    parser.add_argument("--source", default=SOURCE_FILE, help="Excel base de referência.")
    parser.add_argument("--test", action="store_true", help="Modo teste: 1 categoria + 5 artigos.")
    parser.add_argument("--no-git", action="store_true", help="Desactiva commit/push automático.")
    args = parser.parse_args()

    global _git_checkpoint_enabled
    _git_checkpoint_enabled = not args.no_git

    output_file = output_filename_for_today()
    progress = load_progress(output_file)

    log(f"=== Sincronização Action — {progress['execution_date']} ===")
    log(f"Ficheiro base: {args.source}")
    log(f"Ficheiro de saída: {output_file}")

    initialize_workbook(args.source, output_file, progress)

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
            if progress.get("site_scan_done"):
                site_urls = {normalize_url(url) for url in progress["site_urls"]}
                log(f"Retomando — varredura já concluída ({len(site_urls)} URLs no site).")
            else:
                page.goto(BASE_URL, wait_until="networkidle", timeout=90_000)
                accept_cookies_if_visible(page)
                site_urls = scan_site_products(page, progress, test_mode=args.test)

            if not progress.get("removed_marked"):
                mark_removed_products(progress, site_urls)

            existing_urls = {normalize_url(record["URL"]) for record in progress["records"]}
            urls_still_on_site = existing_urls & site_urls
            new_urls = sorted(site_urls - existing_urls)

            checked_updates = set(progress.get("checked_update_urls", []))
            added_new = set(progress.get("added_new_urls", []))

            pending_updates = sorted(urls_still_on_site - checked_updates)
            pending_new = sorted(set(new_urls) - added_new)

            if args.test:
                pending_updates = pending_updates[:3]
                pending_new = pending_new[:2]

            total_tasks = len(pending_updates) + len(pending_new)
            log(f"\n=== FASE 2 — Processamento artigo-a-artigo ({total_tasks} pendentes) ===")
            log(f"  Actualizar existentes: {len(pending_updates)}")
            log(f"  Adicionar novos: {len(pending_new)}")
            log(f"  Removidos (já marcados): {len(existing_urls - site_urls)}")

            records = progress["records"]
            row_states = progress["row_states"]
            url_to_category = progress.get("url_to_category", {})
            task_index = 0

            for url in pending_updates:
                task_index += 1
                log(f"\n[Artigo {task_index}/{total_tasks}] A verificar existente: {url}")
                try:
                    changed, detail = update_existing_product(
                        page,
                        url,
                        records,
                        row_states,
                        url_to_category,
                    )
                    checked_updates.add(normalize_url(url))
                    progress["checked_update_urls"] = sorted(checked_updates)
                    progress["articles_processed"] += 1

                    record = next(
                        r for r in records if normalize_url(r["URL"]) == normalize_url(url)
                    )
                    log(
                        f"[Artigo {task_index}/{total_tasks}] "
                        f"{'Actualizado' if changed else 'Sem alteração'} "
                        f"| {record['Descrição / Nome do artigo'][:55]} "
                        f"| {detail}"
                    )
                except Exception as error:
                    log(f"[Artigo {task_index}/{total_tasks}] Erro: {error}")

                maybe_checkpoint(progress, output_file)
                if task_index < total_tasks:
                    human_delay(2, 5)

            for url in pending_new:
                task_index += 1
                log(f"\n[Artigo {task_index}/{total_tasks}] A adicionar NOVO (azul): {url}")
                try:
                    new_record = add_new_product(
                        page,
                        url,
                        records,
                        row_states,
                        url_to_category,
                    )
                    added_new.add(normalize_url(url))
                    progress["added_new_urls"] = sorted(added_new)
                    progress["articles_processed"] += 1
                    log(
                        f"[Artigo {task_index}/{total_tasks}] NOVO adicionado "
                        f"| {new_record['Descrição / Nome do artigo'][:55]} "
                        f"| Regular: {new_record['Preço Regular']} "
                        f"| Promo: {new_record['Preço Promocional'] or '—'}"
                    )
                except Exception as error:
                    log(f"[Artigo {task_index}/{total_tasks}] Erro: {error}")

                maybe_checkpoint(progress, output_file)
                if task_index < total_tasks:
                    human_delay(2, 5)

            progress["records"] = records
            progress["row_states"] = row_states
            save_styled_excel(records, row_states, output_file)
            save_progress(progress)
            commit_checkpoint_to_repo(output_file, progress["articles_processed"])

            novos = sum(1 for state in row_states.values() if state == "novo")
            removidos = sum(1 for state in row_states.values() if state == "removido")

            log(
                f"\n{'=' * 70}\n"
                f"SINCRONIZAÇÃO CONCLUÍDA\n"
                f"Ficheiro: {output_file}\n"
                f"Total de linhas: {len(records)}\n"
                f"Novos (azul): {novos}\n"
                f"Removidos (vermelho): {removidos}\n"
                f"Artigos processados nesta execução: {progress['articles_processed']}\n"
                f"{'=' * 70}"
            )

        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
