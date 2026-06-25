#!/usr/bin/env python3
"""
process_api_response.py
========================
Única responsabilidade: ler o raw_api_responses.json, aplicar todas as
regras de negócio (classificação de nível, stacks, modalidade) e gerar:

  - raw_jobs_processed.csv  → registro de cada vaga com todos os dados extraídos
  - data.json               → payload final para o dashboard
"""

import json
import re
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# ==========================================
# Arquivos de entrada / saída
# ==========================================
RAW_JSON_PATH = "../raw_api_responses.json"
CSV_PATH      = "../raw_jobs_processed.csv"
JSON_PATH     = "../data.json"

TOP_N   = 15   # itens nos rankings de tecnologia
TOP_CAT = 10   # itens por categoria nos rankings de nível

LEVELS = ["General", "Senior", "Mid-level", "Junior", "Internship"]

COUNTRIES: dict[str, str] = {
    "Brazil": "Brazil",
    "USA":    "USA",
}

# ==========================================
# 1. Classificação de Nível (Senioridade)
# ==========================================
_SENIOR_RE = re.compile(
    r'\bs[eê]nior\b|\bsr\.?\s|\bstaff\s+\w*\s*engineer|\blead\s+\w*\s*engineer|'
    r'\bprincipal\s+\w*\s*(?:engineer|developer)|\barchitect\b|\btech\s+lead\b|'
    r'\bengineering\s+manager|\bsoftware\s+engineer\s+(?:ii|iii|iv|v)\b|'
    r'\bsenior\s+manager|\bdirector\s+of\s+engineering|\bvp\s+of\s+engineering|'
    r'\bmanager,?\s+software|\bstaff\s+(?:developer|engineer)|\bstaff\b.*\bengineer\b',
    re.IGNORECASE
)
_MID_RE = re.compile(
    r'\bpleno\b|\bmid[\s\-]?level\b',
    re.IGNORECASE
)
_JUNIOR_RE = re.compile(
    r'\bj[uú]nior\b|\bjr\.?\s|\bentry[\s\-]?level\b|\biniciante\b|'
    r'\b0[\s\-]?[aà][\s\-]?2\s*anos?|\bearly[\s\-]?career\b|'
    r'\bassociate\s+(?:engineer|developer)\b',
    re.IGNORECASE
)
_INTERNSHIP_RE = re.compile(
    r'\best[aá]gi[oá]rio?\b|\best[aá]gio\b|\binternship\b|\bintern\b|'
    r'\btrainee\b|\baprendiz\b|\bco[\s\-]?op\b|\bjovem\s+aprendiz\b',
    re.IGNORECASE
)

def classify_level(text: str) -> str:
    """Senior tem prioridade máxima para evitar falsos positivos de 'intern' em 'internal'."""
    if _SENIOR_RE.search(text):    return 'Senior'
    if _MID_RE.search(text):       return 'Mid-level'
    if _JUNIOR_RE.search(text):    return 'Junior'
    if _INTERNSHIP_RE.search(text):return 'Internship'
    return 'General'


# ==========================================
# 2. Classificação de Modalidade
# ==========================================

# HÍBRIDO primeiro — vagas híbridas frequentemente mencionam "remote" também,
# então a detecção de híbrido deve ter prioridade sobre a de remoto.
_HYBRID_RE = re.compile(
    r'h[íi]brido|h[íi]brida|hybrid|modelo\s+flex(?:ível)?|'
    r'flex(?:ible)?\s+work|work\s+flex|partially\s+remote|'
    r'semi[\s\-]?remote|partially\s+on[\s\-]?site',
    re.IGNORECASE
)

# REMOTO — apenas quando não for híbrido
_REMOTE_RE = re.compile(
    r'\bremot[oa]?\b|home[\s\-]?office|\bremote\b|distributed\s+team|'
    r'work(?:ing)?\s+from\s+home|\bwfh\b|tele[\s\-]?trabalho|teletrabajo|'
    r'100%\s+remot|fully\s+remote|full[\s\-]?remote|anywhere\b|'
    r'trabalho\s+remoto|vaga\s+remota',
    re.IGNORECASE
)

def classify_modality(text: str, location: str = '') -> str:
    """
    Classifica modalidade da vaga.
    Prioridade: Hybrid > Remote > On-site
    Híbrido tem prioridade porque vagas híbridas frequentemente
    mencionam a palavra 'remote' no texto (ex: "modelo híbrido com opção remota").
    """
    # location == 'Brasil' (sem cidade) geralmente indica vaga remota nacional
    location_is_country_only = location.strip().lower() in ('brasil', 'brazil', 'united states', 'usa')

    if _HYBRID_RE.search(text):
        return 'Hybrid'
    if _REMOTE_RE.search(text) or 'remote' in location.lower() or location_is_country_only:
        return 'Remote'
    return 'On-site'


# ==========================================
# 3. Stacks e Categorias
# ==========================================
CATEGORIES = {
    'Languages': [
        'JavaScript', 'Python', 'Java', 'C#', 'TypeScript', 'PHP', 'Ruby',
        'Golang', 'Go', 'Rust', 'Kotlin', 'Swift', 'C++', 'C', 'SQL', 'Dart',
        'Scala', 'R', 'COBOL', 'Perl', 'Elixir', 'Haskell', 'Lua', 'Shell',
        'Bash', 'Shell Script', 'PowerShell', 'Groovy', 'Clojure', 'F#',
        'Objective-C', 'Assembly', 'MATLAB', 'Julia',
    ],
    'Web Frameworks & Libs': [
        'React', 'Angular', 'Vue.js', 'Vue', 'Next.js', 'Nuxt.js', 'Nuxt',
        'Svelte', 'SvelteKit', 'Gatsby', 'Remix', 'Astro', 'Solid.js',
        'jQuery', 'Bootstrap', 'Tailwind', 'Tailwind CSS', 'Material UI',
        'Chakra UI', 'Ant Design', 'shadcn', 'Storybook',
    ],
    'Backend Frameworks & Libs': [
        'Node.js', 'Express', 'NestJS', 'Fastify', 'Koa',
        'Spring Boot', 'Spring', 'Spring MVC', 'Spring Security', 'Spring Cloud',
        'Django', 'Flask', 'FastAPI', 'Celery', 'SQLAlchemy',
        'Laravel', 'Symfony', 'CodeIgniter', 'Lumen',
        '.NET', 'ASP.NET', 'Entity Framework', 'Blazor',
        'Rails', 'Ruby on Rails', 'Sinatra',
        'Gin', 'Fiber', 'Echo', 'Actix', 'Rocket',
        'Ktor', 'Micronaut', 'Quarkus', 'Phoenix', 'Elixir Phoenix',
        'gRPC', 'GraphQL', 'REST', 'RESTful',
    ],
    'Mobile': [
        'Flutter', 'React Native', 'Ionic', 'Xamarin', 'MAUI',
        'Android', 'iOS', 'Swift', 'SwiftUI', 'Jetpack Compose',
        'Expo', 'Capacitor', 'Cordova',
    ],
    'Data & AI': [
        'TensorFlow', 'PyTorch', 'Keras', 'scikit-learn', 'Pandas',
        'NumPy', 'Spark', 'Apache Spark', 'Kafka', 'Apache Kafka',
        'Airflow', 'Apache Airflow', 'dbt', 'Hadoop', 'Hive',
        'Power BI', 'Tableau', 'Looker', 'Metabase', 'Superset',
        'MLflow', 'Langchain', 'LangChain', 'OpenAI', 'LLM', 'RAG',
        'Databricks', 'Snowflake', 'BigQuery', 'Redshift',
        'Machine Learning', 'Deep Learning', 'NLP', 'Computer Vision',
        'Jupyter', 'Matplotlib', 'Seaborn', 'Plotly',
    ],
    'Databases': [
        'PostgreSQL', 'MySQL', 'MongoDB', 'Redis', 'Oracle', 'SQL Server',
        'MariaDB', 'SQLite', 'Cassandra', 'DynamoDB', 'Elasticsearch',
        'Neo4j', 'CouchDB', 'InfluxDB', 'TimescaleDB', 'Supabase',
        'Firebase', 'PlanetScale', 'Neon', 'Cockroachdb',
    ],
    'DevOps & Cloud': [
        'Docker', 'Kubernetes', 'Helm', 'Istio', 'ArgoCD',
        'AWS', 'Azure', 'GCP', 'Google Cloud', 'DigitalOcean', 'Heroku', 'Vercel',
        'CI/CD', 'GitHub Actions', 'GitLab CI', 'Jenkins', 'CircleCI', 'Travis CI',
        'Terraform', 'Ansible', 'Puppet', 'Chef', 'Pulumi',
        'Linux', 'Unix', 'Nginx', 'Apache', 'Caddy',
        'Prometheus', 'Grafana', 'Datadog', 'Sentry', 'New Relic', 'Splunk',
        'CloudFormation', 'CDK', 'SAM', 'Lambda', 'ECS', 'EKS', 'S3', 'RDS', 'EC2',
        'Azure DevOps', 'Azure Kubernetes', 'AKS', 'GKE',
    ],
    'Tools & Practices': [
        'Git', 'GitHub', 'GitLab', 'Bitbucket', 'Jira', 'Confluence', 'Trello', 'Notion',
        'Scrum', 'Agile', 'Kanban', 'SAFe', 'XP', 'TDD', 'BDD', 'DDD', 'SOLID',
        'Clean Architecture', 'Clean Code', 'Microservices', 'Microsserviços',
        'Monorepo', 'Serverless', 'OpenAPI', 'Swagger', 'Postman', 'Insomnia',
        'WebSockets', 'WebSocket', 'MQTT', 'RabbitMQ', 'ActiveMQ', 'SQS',
        'OAuth', 'JWT', 'OpenID', 'SAML', 'SSO',
        'OWASP', 'Cybersecurity', 'Segurança', 'Penetration Testing',
    ],
}

# Mapa base: stack individual → categoria
_STACK_TO_CAT_RAW: dict[str, str] = {
    stack: cat for cat, stacks in CATEGORIES.items() for stack in stacks
}

# Aliases de família: stacks individuais que devem ser contadas juntas
STACK_ALIASES: dict[str, str] = {
    'JavaScript': 'JavaScript/TypeScript',
    'TypeScript': 'JavaScript/TypeScript',
    'C':          'C/C++',
    'C++':        'C/C++',
}

# Mapa canônico: inclui tanto stacks individuais quanto nomes de família
# Os nomes de família herdam a categoria da primeira entrada encontrada
STACK_TO_CAT: dict[str, str] = dict(_STACK_TO_CAT_RAW)
for _individual, _canonical in STACK_ALIASES.items():
    if _canonical not in STACK_TO_CAT:
        STACK_TO_CAT[_canonical] = _STACK_TO_CAT_RAW.get(_individual, 'Other')

def _compile_regex(keyword: str) -> re.Pattern:
    escaped = re.escape(keyword)
    if ' ' in keyword:
        escaped = r'\s+'.join(re.escape(p) for p in keyword.split())
    return re.compile(rf'(?i)(?<![a-z0-9]){escaped}(?![a-z0-9])')

# Compila regex apenas para stacks individuais (não para nomes de família)
REGEX_MAP: dict[str, re.Pattern] = {
    stack: _compile_regex(stack)
    for stack in _STACK_TO_CAT_RAW
}

def extract_stacks(text: str) -> list[str]:
    """Retorna lista de stacks canônicas encontradas no texto."""
    found = [STACK_ALIASES.get(s, s) for s, rx in REGEX_MAP.items() if rx.search(text)]
    return sorted(set(found))


# ==========================================
# 4. Helpers numéricos
# ==========================================
def safe_int(v) -> int:
    try: return int(float(v))
    except: return 0

def safe_float(v) -> float:
    try: return round(float(v), 1)
    except: return 0.0


# ==========================================
# 5. Leitura do raw_api_responses.json
# ==========================================
def load_raw_jobs(path: str) -> list[dict]:
    """
    Lê o raw_api_responses.json e retorna uma lista de vagas únicas (deduplicadas por ID).
    Cada item já inclui o campo _meta com country_name e search_term.
    """
    with open(path, encoding='utf-8') as f:
        data = json.load(f)

    seen_ids: set = set()
    unique_jobs: list[dict] = []

    for page in data.get('requests', []):
        meta         = page.get('_meta', {})
        country_name = meta.get('country_name')
        search_term  = meta.get('search_term', '')

        for job in page.get('results', []):
            jid = job.get('id') or job.get('redirect_url', '')
            if not jid or jid in seen_ids:
                continue
            
            # Retro-compatibilidade: inferir country do redirect_url se meta faltar
            if not country_name:
                url = job.get('redirect_url', '').lower()
                if 'br.' in url or '.br' in url:
                    country_name = 'Brazil'
                else:
                    country_name = 'USA'

            seen_ids.add(jid)
            job['_country_name'] = country_name
            job['_search_term']  = search_term
            unique_jobs.append(job)

    return unique_jobs


# ==========================================
# 6. Processamento das vagas
# ==========================================
def process_jobs(unique_jobs: list[dict]) -> tuple[pd.DataFrame, dict]:
    """
    Recebe a lista de vagas únicas e retorna:
      - df_processed: DataFrame com uma linha por vaga (para o CSV)
      - distribution_by_country: {country: {level: count}}
    """
    rows: list[dict] = []
    distribution_by_country: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))

    for job in unique_jobs:
        jid          = job.get('id') or job.get('redirect_url', '')
        title        = job.get('title', '') or ''
        desc         = job.get('description', '') or ''
        country_name = job.get('_country_name', 'Unknown')
        search_term  = job.get('_search_term', '')
        created      = job.get('created', '')
        location     = job.get('location', {}).get('display_name', '')

        text     = f"{title} {desc}"
        level    = classify_level(text)
        modality = classify_modality(text, location)
        stacks   = extract_stacks(text)

        distribution_by_country[country_name][level] += 1

        rows.append({
            'ID':              jid,
            'Search Term':     search_term,
            'Job Title':       title,
            'Country':         country_name,
            'Level':           level,
            'Modality':        modality,
            'Mentioned Stacks': ', '.join(stacks),
            'Creation Date':   created,
            'Description':     desc,
        })

    df = pd.DataFrame(rows)
    return df, dict(distribution_by_country)


# ==========================================
# 7. Agregação para o data.json
#    (mesma lógica do script.py original)
# ==========================================
def aggregate_for_dashboard(
    df_processed: pd.DataFrame,
    distribution_by_country: dict
) -> dict:
    """Gera o payload completo para o data.json."""

    # Expande stacks por vaga para montar o DataFrame de menções por nível
    mention_rows: list[dict] = []
    for _, row in df_processed.iterrows():
        stacks = [s.strip() for s in str(row['Mentioned Stacks']).split(',') if s.strip()]
        for stack in stacks:
            cat = STACK_TO_CAT.get(stack, 'Other')
            mention_rows.append({
                'Country':    row['Country'],
                'Level':      row['Level'],
                'Modality':   row['Modality'],
                'Category':   cat,
                'Technology': stack,
            })

    if not mention_rows:
        # Se não há menções, retorna estrutura vazia
        empty = {l: [] for l in LEVELS}
        return {
            "kpi": {"totalBR": 0, "totalUSA": 0, "totalGlobal": 0},
            "topGlobal": [], "topBR": [], "topUSA": [],
            "byLevel": empty, "byLevelBR": empty, "byLevelUSA": empty,
            "levelsJobsBR": {}, "levelsJobsUSA": {},
            "modalityBR": {}, "modalityUSA": {},
            "categoryBR": {}, "categoryUSA": {},
            "topRemote": [],
            "byCategory": {}, "byCategoryBR": {}, "byCategoryUSA": {},
        }

    mdf = pd.DataFrame(mention_rows)

    # Agrega contagens de modalidade por stack
    agg_rows: list[dict] = []
    for (country, level, cat, tech), grp in mdf.groupby(['Country', 'Level', 'Category', 'Technology']):
        total   = len(grp)
        remote  = (grp['Modality'] == 'Remote').sum()
        hybrid  = (grp['Modality'] == 'Hybrid').sum()
        onsite  = (grp['Modality'] == 'On-site').sum()
        agg_rows.append({
            'Country':    country,
            'Level':      level,
            'Category':   cat,
            'Technology': tech,
            'Total':      total,
            'Remote':     remote,
            'Hybrid':     hybrid,
            'On-site':    onsite,
            '% Remote':   round(remote / total * 100, 1) if total else 0.0,
        })

    all_df  = pd.DataFrame(agg_rows)
    br_df   = all_df[all_df['Country'] == 'Brazil'].copy()
    usa_df  = all_df[all_df['Country'] == 'USA'].copy()

    # ── Modalidade (baseada no nível General) ──
    def modal_pais(df: pd.DataFrame) -> dict:
        sub = df[df['Level'] == 'General']
        if sub.empty:
            return {'remote': 0, 'hybrid': 0, 'onsite': 0}
        return {
            'remote': safe_int(sub['Remote'].sum()),
            'hybrid': safe_int(sub['Hybrid'].sum()),
            'onsite': safe_int(sub['On-site'].sum()),
        }

    modal_br  = modal_pais(br_df)
    modal_usa = modal_pais(usa_df)
    total_modal_br  = sum(modal_br.values())  or 1
    total_modal_usa = sum(modal_usa.values()) or 1
    pct_remote_br   = round(modal_br['remote']  / total_modal_br  * 100, 1)
    pct_remote_usa  = round(modal_usa['remote'] / total_modal_usa * 100, 1)

    # ── Rankings globais de tecnologia ──
    def top_list(df: pd.DataFrame, n: int) -> list[dict]:
        if df.empty: return []
        out = df.groupby('Technology')['Total'].sum().sort_values(ascending=False).head(n)
        return [{'t': t, 'v': safe_int(v)} for t, v in out.items()]

    gen_all = all_df[all_df['Level'] == 'General']
    gen_br  = br_df[br_df['Level']   == 'General']
    gen_usa = usa_df[usa_df['Level'] == 'General']

    top_global_list = top_list(gen_all, TOP_N)
    top_br_list     = top_list(gen_br,  TOP_N)
    top_usa_list    = top_list(gen_usa, TOP_N)

    stack_global   = top_global_list[0]['t'] if top_global_list else '-'
    mencoes_global = top_global_list[0]['v'] if top_global_list else 0
    stack_br       = top_br_list[0]['t']     if top_br_list     else '-'
    mencoes_br     = top_br_list[0]['v']     if top_br_list     else 0
    stack_usa      = top_usa_list[0]['t']    if top_usa_list    else '-'
    mencoes_usa    = top_usa_list[0]['v']    if top_usa_list    else 0

    # ── Stacks mais remotas ──
    rem_df = gen_all.groupby('Technology').agg(Total=('Total', 'sum'), Remote=('Remote', 'sum'))
    rem_df = rem_df[rem_df['Total'] >= 5].copy()
    rem_df['pct'] = (rem_df['Remote'] / rem_df['Total'] * 100).round(1)
    top_remote = [
        {'t': t, 'pct': row['pct'], 'n': safe_int(row['Total'])}
        for t, row in rem_df.sort_values('pct', ascending=False).head(14).iterrows()
    ]

    # ── Rankings por nível ──
    def por_nivel_dict(df: pd.DataFrame) -> dict:
        result = {}
        for level in LEVELS:
            sub = df[df['Level'] == level]
            result[level] = top_list(sub, TOP_CAT)
        return result

    by_level     = por_nivel_dict(all_df)
    by_level_br  = por_nivel_dict(br_df)
    by_level_usa = por_nivel_dict(usa_df)

    # ── Distribuição de vagas por nível por país ──
    niveis_vagas_br  = dict(distribution_by_country.get('Brazil', {}))
    niveis_vagas_usa = dict(distribution_by_country.get('USA',    {}))

    real_total_br  = sum(niveis_vagas_br.values())  or 1
    real_total_usa = sum(niveis_vagas_usa.values()) or 1

    def pct_nivel(niveis: dict, nivel: str, total: int) -> float:
        return round(niveis.get(nivel, 0) / total * 100, 1) if total else 0.0

    pct_senior_br   = pct_nivel(niveis_vagas_br,  'Senior',     real_total_br)
    pct_estagio_br  = pct_nivel(niveis_vagas_br,  'Internship', real_total_br)
    pct_senior_usa  = pct_nivel(niveis_vagas_usa, 'Senior',     real_total_usa)
    pct_estagio_usa = pct_nivel(niveis_vagas_usa, 'Internship', real_total_usa)

    # ── Categorias ──
    def cats_dict(df: pd.DataFrame) -> dict:
        sub = df[(df['Level'] == 'General') & (~df['Category'].isin(['nan', 'None', '']))]
        return {k: safe_int(v) for k, v in sub.groupby('Category')['Total'].sum().to_dict().items()}

    cats_br  = cats_dict(br_df)
    cats_usa = cats_dict(usa_df)

    def por_cat_dict(df: pd.DataFrame) -> dict:
        df_clean = df[~df['Category'].isin(['nan', 'None', ''])]
        result = {}
        for cat in df_clean['Category'].dropna().unique():
            sub = df_clean[df_clean['Category'] == cat]
            result[cat] = top_list(sub, TOP_CAT)
        return result

    by_cat     = por_cat_dict(all_df)
    by_cat_br  = por_cat_dict(br_df)
    by_cat_usa = por_cat_dict(usa_df)

    return {
        'kpi': {
            'totalBR':          real_total_br,
            'totalUSA':         real_total_usa,
            'totalGlobal':      real_total_br + real_total_usa,
            'pctRemoteBR':      pct_remote_br,
            'pctRemoteUSA':     pct_remote_usa,
            'stackGlobal':      stack_global,
            'mentionsGlobal':   mencoes_global,
            'stackBR':          stack_br,
            'mentionsBR':       mencoes_br,
            'stackUSA':         stack_usa,
            'mentionsUSA':      mencoes_usa,
            'pctSeniorBR':      pct_senior_br,
            'pctInternshipBR':  pct_estagio_br,
            'pctSeniorUSA':     pct_senior_usa,
            'pctInternshipUSA': pct_estagio_usa,
        },
        'topGlobal':     top_global_list,
        'topBR':         top_br_list,
        'topUSA':        top_usa_list,
        'byLevel':       by_level,
        'byLevelBR':     by_level_br,
        'byLevelUSA':    by_level_usa,
        'levelsJobsBR':  niveis_vagas_br,
        'levelsJobsUSA': niveis_vagas_usa,
        'modalityBR':    modal_br,
        'modalityUSA':   modal_usa,
        'categoryBR':    cats_br,
        'categoryUSA':   cats_usa,
        'topRemote':     top_remote,
        'byCategory':    by_cat,
        'byCategoryBR':  by_cat_br,
        'byCategoryUSA': by_cat_usa,
    }


# ==========================================
# 8. Main
# ==========================================
def main():
    print("=" * 60)
    print("PROCESS API RESPONSE — Processamento e geração de dados")
    print("=" * 60)

    # 1. Carrega vagas únicas do JSON bruto
    print(f"\n[1] Lendo {RAW_JSON_PATH} ...")
    unique_jobs = load_raw_jobs(RAW_JSON_PATH)
    print(f"    Vagas únicas encontradas: {len(unique_jobs):,}")

    # 2. Processa cada vaga (level, modality, stacks)
    print("\n[2] Classificando vagas (nível, modalidade, stacks) ...")
    df_processed, distribution_by_country = process_jobs(unique_jobs)

    for country, dist in distribution_by_country.items():
        total = sum(dist.values()) or 1
        print(f"\n    [{country}]")
        for level, qty in sorted(dist.items(), key=lambda x: -x[1]):
            print(f"      {level:<20} {qty:>7,} ({qty/total*100:.1f}%)")

    # 3. Salva CSV com todas as vagas processadas
    print(f"\n[3] Salvando {CSV_PATH} ...")
    df_processed.to_csv(CSV_PATH, index=False, encoding='utf-8-sig')
    print(f"    ✅ {len(df_processed):,} vagas salvas.")

    # 4. Agrega e gera o data.json
    print(f"\n[4] Gerando {JSON_PATH} ...")
    dashboard_data = aggregate_for_dashboard(df_processed, distribution_by_country)

    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(dashboard_data, f, ensure_ascii=False, indent=2)

    kpi = dashboard_data['kpi']
    print(f"    ✅ {JSON_PATH} gerado com sucesso!")
    print(f"       BR:  {kpi['totalBR']:,} vagas | % Sênior: {kpi['pctSeniorBR']}% | % Estágio: {kpi['pctInternshipBR']}%")
    print(f"       USA: {kpi['totalUSA']:,} vagas | % Sênior: {kpi['pctSeniorUSA']}% | % Estágio: {kpi['pctInternshipUSA']}%")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
