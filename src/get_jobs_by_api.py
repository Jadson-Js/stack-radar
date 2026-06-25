#!/usr/bin/env python3
"""
get_jobs_by_api.py
==================
Única responsabilidade: chamar a API do Adzuna e salvar o resultado bruto
em raw_api_responses.json.

Não classifica, não processa, não gera Excel, não gera CSV.
"""

import os
import json
import time
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# Configuração e Credenciais
# ==========================================
APP_ID  = os.getenv('APP_ID')
APP_KEY = os.getenv('APP_KEY')

DAYS_AGO              = 7
MAX_WORKERS           = 2
RETRY_ATTEMPTS        = 3
RETRY_DELAY           = 5
JOB_LIMIT             = 100      # máximo de vagas únicas a coletar
MAX_PAGES_PER_TERM    = 150
MAX_CONSECUTIVE_FAILURES   = 3
MAX_CONSECUTIVE_SATURATION = 2

OUTPUT_FILE = "raw_api_responses.json"

COUNTRIES: dict[str, dict] = {
    'br': {
        'name': 'Brazil',
        'terms': [
            'desenvolvedor', 'developer', 'programador', 'engenheiro de software',
            'software engineer', 'web developer', 'analista de sistemas',
            'pessoa desenvolvedora', 'dev', 'desenvolvedor fullstack',
            'desenvolvedor backend', 'desenvolvedor frontend', 'desenvolvedor mobile',
            'full stack developer', 'backend developer', 'frontend developer',
            'arquiteto de software', 'software architect', 'tech lead',
            'líder técnico', 'analista de TI', 'analista desenvolvedor',
        ],
    },
    'us': {
        'name': 'USA',
        'terms': [
            'developer', 'software engineer', 'web developer', 'application developer',
            'computer programmer', 'backend developer', 'backend engineer',
            'frontend developer', 'frontend engineer', 'full stack developer',
            'full stack engineer', 'mobile developer', 'ios developer',
            'android developer', 'software architect', 'solutions architect',
            'tech lead', 'engineering manager',
        ],
    },
}


# ==========================================
# Funções de Busca
# ==========================================
def _fetch_page(country: str, term: str, page: int) -> tuple[list[dict] | None, dict | None]:
    """Busca uma página de resultados da API e retorna (resultados, json_bruto)."""
    url = f"https://api.adzuna.com/v1/api/jobs/{country}/search/{page}"
    params = {
        'app_id': APP_ID, 'app_key': APP_KEY, 'what': term,
        'max_days_old': DAYS_AGO, 'results_per_page': 50,
    }
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            resp = requests.get(url, params=params, timeout=15)
            if resp.status_code in (429, 503):
                wait = RETRY_DELAY * attempt
                print(f"  ! [{country.upper()}] p{page:>3} — {resp.status_code}, aguardando {wait}s")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            raw_json = resp.json()
            results = raw_json.get('results', [])
            print(f"  * [{country.upper()}] '{term[:20]:<20}' p{page:>3} — {len(results)} vagas")
            return results, raw_json
        except Exception as exc:
            print(f"  x [{country.upper()}] p{page:>3} — erro: {exc}")
            if attempt < RETRY_ATTEMPTS:
                time.sleep(RETRY_DELAY)
    return None, None


def collect_country(country: str, terms: list[str], global_seen: set) -> list[dict]:
    """Coleta todos os resultados brutos de um país e retorna lista de JSONs de página."""
    all_raw_pages = []

    for term in terms:
        if len(global_seen) >= JOB_LIMIT:
            print(f"\n  [!] Limite global atingido. Parando coleta.")
            break

        print(f"\n  > Buscando '{term}' ...")
        page = 1
        consecutive_failures    = 0
        consecutive_saturation  = 0

        while page <= MAX_PAGES_PER_TERM:
            if len(global_seen) >= JOB_LIMIT:
                break

            page_batch   = list(range(page, page + MAX_WORKERS))
            batch_results: dict[int, list[dict] | None] = {}

            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
                futures = {pool.submit(_fetch_page, country, term, p): p for p in page_batch}
                for fut in as_completed(futures):
                    res, raw_json = fut.result()
                    batch_results[futures[fut]] = res
                    if raw_json:
                        # Anota o contexto da requisição no próprio objeto
                        raw_json['_meta'] = {
                            'country_code': country,
                            'country_name': COUNTRIES[country]['name'],
                            'search_term':  term,
                            'page':         futures[fut],
                        }
                        all_raw_pages.append(raw_json)

            stop_term = False
            for p in page_batch:
                results = batch_results.get(p)

                if not results:
                    consecutive_failures += 1
                    if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                        print(f"  - '{term}' — {consecutive_failures} falhas consecutivas. Encerrando termo.")
                        stop_term = True
                        break
                    continue
                else:
                    consecutive_failures = 0

                new_on_page = 0
                for job in results:
                    jid = job.get('id') or job.get('redirect_url', '')
                    if jid and jid not in global_seen:
                        global_seen.add(jid)
                        new_on_page += 1

                if new_on_page == 0:
                    consecutive_saturation += 1
                    if consecutive_saturation >= MAX_CONSECUTIVE_SATURATION:
                        print(f"  - '{term}' — loop de duplicatas detectado. Encerrando termo.")
                        stop_term = True
                        break
                else:
                    consecutive_saturation = 0

                if len(global_seen) >= JOB_LIMIT:
                    stop_term = True
                    break

            if stop_term:
                break
            page += MAX_WORKERS

        if page > MAX_PAGES_PER_TERM:
            print(f"  - '{term}' atingiu o limite de segurança ({MAX_PAGES_PER_TERM} páginas).")

        print(f"  -> Vagas únicas acumuladas: {len(global_seen):,}")

    return all_raw_pages


# ==========================================
# Main
# ==========================================
def main():
    print("=" * 60)
    print("GET JOBS BY API — Coleta bruta do Adzuna")
    print(f"Limite global: {JOB_LIMIT:,} vagas únicas")
    print("=" * 60)

    global_seen: set = set()
    all_raw_pages: list[dict] = []

    for country, cfg in COUNTRIES.items():
        if len(global_seen) >= JOB_LIMIT:
            print(f"\n[!] Limite atingido. Pulando {cfg['name']}.")
            break

        print(f"\n{'─'*60}")
        print(f"País: {cfg['name']} ({country.upper()})")
        print(f"Vagas globais: {len(global_seen):,} / {JOB_LIMIT:,}")
        print(f"{'─'*60}")

        pages = collect_country(country, cfg['terms'], global_seen)
        all_raw_pages.extend(pages)

    # Salva tudo num JSON único
    output = {"requests": all_raw_pages}
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    total_results = sum(len(p.get('results', [])) for p in all_raw_pages)
    print(f"\n{'='*60}")
    print(f"✅ Coleta concluída!")
    print(f"   Páginas coletadas:  {len(all_raw_pages):,}")
    print(f"   Hits totais brutos: {total_results:,}")
    print(f"   Vagas únicas vistas:{len(global_seen):,}")
    print(f"   Arquivo salvo:      {OUTPUT_FILE}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
