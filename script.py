import os
import requests
import pandas as pd
import re
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from dotenv import load_dotenv

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# ==========================================
# 1. Configurações e Credenciais
# ==========================================
APP_ID  = os.getenv('APP_ID')
APP_KEY = os.getenv('APP_KEY')

DIAS_ATRAS        = 365
PAGINAS_POR_BUSCA = 100
MAX_WORKERS       = 2
RETRY_ATTEMPTS    = 3
RETRY_DELAY       = 5

PAISES: dict[str, dict] = {
    'br': {
        'nome': 'Brasil',
        'termos': [
            'desenvolvedor',
            'developer',
            'programador',
            'engenheiro de software',
            'software engineer',
            'web developer',
            'analista de sistemas',
            'pessoa desenvolvedora',
            'desenvolvedor fullstack',
            'desenvolvedor backend',
            'desenvolvedor frontend',
            'desenvolvedor mobile',
            'arquiteto de software',
        ],
    },
    'us': {
        'nome': 'EUA',
        'termos': [
            'software engineer',
            'software developer',
            'web developer',
            'backend developer',
            'frontend developer',
            'full stack developer',
            'mobile developer',
            'systems analyst',
            'cloud engineer',
            'devops engineer',
            'data engineer',
            'machine learning engineer',
            'platform engineer',
        ],
    },
}

# ==========================================
# 2. Classificação de nível
# ==========================================
_SENIOR_RE = re.compile(
    r's[eê]nior|sr\.?\s|pleno|mid[\s\-]?level|staff\s+engineer|lead\s+engineer|'
    r'principal\s+engineer|architect|tech\s+lead|engineering\s+manager',
    re.IGNORECASE
)
_JUNIOR_RE = re.compile(
    r'j[uú]nior|jr\.?\s|entry[\s\-]?level|iniciante|0[\s\-]?[aà][\s\-]?2\s*anos?|'
    r'early[\s\-]?career|associate\s+engineer|associate\s+developer',
    re.IGNORECASE
)
_ESTAGIO_RE = re.compile(
    r'est[aá]gi[oá]rio?|est[aá]gio|intern(ship)?|trainee|aprendiz|co[\s\-]?op',
    re.IGNORECASE
)

TODOS_OS_NIVEIS = ('Junior', 'Estágio/Intern', 'Pleno', 'Sênior', 'Geral')

def classificar_nivel(texto: str) -> str:
    if _ESTAGIO_RE.search(texto):
        return 'Estágio/Intern'
    if _JUNIOR_RE.search(texto):
        return 'Junior'
    if _SENIOR_RE.search(texto):
        # Diferencia pleno de sênior
        if re.search(r'pleno|mid[\s\-]?level', texto, re.IGNORECASE):
            return 'Pleno'
        return 'Sênior'
    return 'Geral'

# ==========================================
# 3. Dicionário de Categorias (EXPANDIDO)
# ==========================================
CATEGORIAS = {
    'Linguagens': [
        'JavaScript', 'Python', 'Java', 'C#', 'TypeScript', 'PHP', 'Ruby',
        'Golang', 'Go', 'Rust', 'Kotlin', 'Swift', 'C++', 'C', 'SQL', 'Dart',
        'Scala', 'R', 'COBOL', 'Perl', 'Elixir', 'Haskell', 'Lua', 'Shell',
        'Bash', 'PowerShell', 'Groovy', 'Clojure', 'F#', 'Objective-C',
        'Assembly', 'MATLAB', 'Julia',
    ],
    'Frameworks & Libs Web': [
        'React', 'Angular', 'Vue.js', 'Vue', 'Next.js', 'Nuxt.js', 'Nuxt',
        'Svelte', 'SvelteKit', 'Gatsby', 'Remix', 'Astro', 'Solid.js',
        'jQuery', 'Bootstrap', 'Tailwind', 'Tailwind CSS', 'Material UI',
        'Chakra UI', 'Ant Design', 'shadcn', 'Storybook',
    ],
    'Frameworks & Libs Backend': [
        'Node.js', 'Express', 'NestJS', 'Fastify', 'Koa',
        'Spring Boot', 'Spring', 'Spring MVC', 'Spring Security', 'Spring Cloud',
        'Django', 'Flask', 'FastAPI', 'Celery', 'SQLAlchemy',
        'Laravel', 'Symfony', 'CodeIgniter', 'Lumen',
        '.NET', 'ASP.NET', 'Entity Framework', 'Blazor',
        'Rails', 'Ruby on Rails', 'Sinatra',
        'Gin', 'Fiber', 'Echo',
        'Actix', 'Rocket',
        'Ktor', 'Micronaut', 'Quarkus',
        'Phoenix', 'Elixir Phoenix',
        'gRPC', 'GraphQL', 'REST', 'RESTful',
    ],
    'Mobile': [
        'Flutter', 'React Native', 'Ionic', 'Xamarin', 'MAUI',
        'Android', 'iOS', 'Swift', 'SwiftUI', 'Jetpack Compose',
        'Expo', 'Capacitor', 'Cordova',
    ],
    'Dados & IA': [
        'TensorFlow', 'PyTorch', 'Keras', 'scikit-learn', 'Pandas',
        'NumPy', 'Spark', 'Apache Spark', 'Kafka', 'Apache Kafka',
        'Airflow', 'Apache Airflow', 'dbt', 'Hadoop', 'Hive',
        'Power BI', 'Tableau', 'Looker', 'Metabase', 'Superset',
        'MLflow', 'Langchain', 'LangChain', 'OpenAI', 'LLM', 'RAG',
        'Databricks', 'Snowflake', 'BigQuery', 'Redshift',
        'Machine Learning', 'Deep Learning', 'NLP', 'Computer Vision',
        'Jupyter', 'Matplotlib', 'Seaborn', 'Plotly',
    ],
    'Bancos de Dados': [
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
        'CloudFormation', 'CDK', 'SAM',
        'Lambda', 'ECS', 'EKS', 'S3', 'RDS', 'EC2',
        'Azure DevOps', 'Azure Kubernetes', 'AKS', 'GKE',
    ],
    'Ferramentas & Práticas': [
        'Git', 'GitHub', 'GitLab', 'Bitbucket',
        'Jira', 'Confluence', 'Trello', 'Notion',
        'Scrum', 'Agile', 'Kanban', 'SAFe', 'XP',
        'TDD', 'BDD', 'DDD', 'SOLID', 'Clean Architecture', 'Clean Code',
        'Microservices', 'Microsserviços', 'Monorepo', 'Serverless',
        'OpenAPI', 'Swagger', 'Postman', 'Insomnia',
        'WebSockets', 'WebSocket', 'MQTT', 'RabbitMQ', 'ActiveMQ', 'SQS',
        'OAuth', 'JWT', 'OpenID', 'SAML', 'SSO',
        'OWASP', 'Cybersecurity', 'Segurança', 'Penetration Testing',
        'Linux', 'Bash', 'Shell Script',
    ],
}

STACK_PARA_CAT = {
    stack: cat
    for cat, stacks in CATEGORIAS.items()
    for stack in stacks
}

# ==========================================
# 4. Regex compiladas
# ==========================================
def _compilar_regex(keyword: str) -> re.Pattern:
    escaped = re.escape(keyword)
    # Para termos com espaços (ex: "Spring Boot"), busca flexível
    if ' ' in keyword:
        escaped = r'\s+'.join(re.escape(p) for p in keyword.split())
    return re.compile(rf'(?i)(?<![a-z0-9]){escaped}(?![a-z0-9])')

REGEX_MAP: dict[str, re.Pattern] = {
    stack: _compilar_regex(stack) for stack in STACK_PARA_CAT
}

_REMOTO_RE = re.compile(r'remoto|home[\s\-]?office|remote|distributed|anywhere', re.IGNORECASE)
_HIBRIDO_RE = re.compile(r'híbrido|híbrida|hybrid', re.IGNORECASE)

# ==========================================
# 5. Busca com retry e paralelismo
# ==========================================
def _buscar_pagina(country: str, termo: str, page: int) -> list[dict]:
    url = f"https://api.adzuna.com/v1/api/jobs/{country}/search/{page}"
    params = {
        'app_id': APP_ID,
        'app_key': APP_KEY,
        'what': termo,
        'max_days_old': DIAS_ATRAS,
        'results_per_page': 50,
    }
    for tentativa in range(1, RETRY_ATTEMPTS + 1):
        try:
            resp = requests.get(url, params=params, timeout=15)
            if resp.status_code in (429, 503):
                wait = RETRY_DELAY * tentativa
                print(f"  ⚠ [{country.upper()}] '{termo}' p{page:>2} — {resp.status_code}, aguardando {wait}s (tentativa {tentativa}/{RETRY_ATTEMPTS})")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            resultados = resp.json().get('results', [])
            print(f"  ✓ [{country.upper()}] '{termo[:20]:<20}' p{page:>2} — {len(resultados)} vagas")
            return resultados
        except requests.exceptions.HTTPError as exc:
            print(f"  ✗ [{country.upper()}] p{page:>2} — HTTP: {exc}")
            break
        except Exception as exc:
            print(f"  ✗ [{country.upper()}] p{page:>2} — erro: {exc}")
            if tentativa < RETRY_ATTEMPTS:
                time.sleep(RETRY_DELAY)
    return []


def obter_vagas_pais(country: str, termos: list[str]) -> dict[str, dict]:
    vagas: dict[str, dict] = {}
    for termo in termos:
        print(f"\n  🔍 Buscando '{termo}' ({PAGINAS_POR_BUSCA} págs)...")
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = {
                pool.submit(_buscar_pagina, country, termo, p): p
                for p in range(1, PAGINAS_POR_BUSCA + 1)
            }
            for fut in as_completed(futures):
                for vaga in fut.result():
                    vid = vaga.get('id') or vaga.get('redirect_url', '')
                    if vid and vid not in vagas:
                        vagas[vid] = vaga
    print(f"\n  → {len(vagas)} vagas únicas para {country.upper()}")
    return vagas

# ==========================================
# 6. Processamento
# ==========================================
def _modalidade(texto: str, loc: str) -> str:
    if _REMOTO_RE.search(texto) or 'remote' in loc:
        return 'Remoto'
    if _HIBRIDO_RE.search(texto):
        return 'Híbrido'
    return 'Presencial'


def processar_pais(country: str, vagas: dict[str, dict]) -> tuple[list[dict], dict[str, int]]:
    # Estrutura: nivel -> stack -> modalidade -> count
    contagem: dict[str, dict] = {
        nivel: defaultdict(lambda: {'Total': 0, 'Remoto': 0, 'Híbrido': 0, 'Presencial': 0})
        for nivel in TODOS_OS_NIVEIS
    }
    dist: dict[str, int] = defaultdict(int)

    for vaga in vagas.values():
        titulo = vaga.get('title', '')
        descricao = vaga.get('description', '') or ''
        texto = f"{titulo} {descricao}"
        nivel = classificar_nivel(texto)
        dist[nivel] += 1

        loc = vaga.get('location', {}).get('display_name', '').lower()
        modalidade = _modalidade(texto, loc)

        for stack, rx in REGEX_MAP.items():
            if rx.search(texto):
                contagem[nivel][stack]['Total'] += 1
                contagem[nivel][stack][modalidade] += 1

    dados: list[dict] = []
    for nivel in TODOS_OS_NIVEIS:
        for stack, info in contagem[nivel].items():
            if info['Total'] > 0:
                dados.append({
                    'País':          PAISES[country]['nome'],
                    'Nível':         nivel,
                    'Categoria':     STACK_PARA_CAT[stack],
                    'Tecnologia':    stack,
                    'Total':         info['Total'],
                    'Remoto':        info['Remoto'],
                    'Híbrido':       info['Híbrido'],
                    'Presencial':    info['Presencial'],
                    '% Remoto':      round(info['Remoto'] / info['Total'] * 100, 1) if info['Total'] else 0,
                })

    return dados, dist

# ==========================================
# 7. Relatório no terminal
# ==========================================
def imprimir_relatorio(df: pd.DataFrame, nome_pais: str) -> None:
    df_pais = df[df['País'] == nome_pais]
    print(f"\n\n{'#'*60}")
    print(f"#  🌎  MERCADO: {nome_pais.upper()}")
    print(f"{'#'*60}")
    for nivel in TODOS_OS_NIVEIS:
        df_nivel = df_pais[df_pais['Nível'] == nivel]
        if df_nivel.empty:
            continue
        print(f"\n{'='*60}")
        print(f"📊  {nivel.upper()}")
        print(f"{'='*60}")
        for cat in CATEGORIAS:
            df_cat = (
                df_nivel[df_nivel['Categoria'] == cat]
                .sort_values('Total', ascending=False)
                .head(10)
            )
            if not df_cat.empty:
                print(f"\n--- TOP: {cat.upper()} ---")
                print(df_cat[['Tecnologia', 'Total', 'Remoto', 'Híbrido', 'Presencial', '% Remoto']].to_string(index=False))

# ==========================================
# 8. Excel rico com openpyxl
# ==========================================

# Paleta de cores
COR_HEADER      = 'FF1F3864'  # azul escuro
COR_HEADER_FONT = 'FFFFFFFF'  # branco
COR_NIVEL = {
    'Sênior':       'FF2E75B6',
    'Pleno':        'FF2E86C1',
    'Junior':       'FF1E8449',
    'Estágio/Intern': 'FF8E44AD',
    'Geral':        'FF555555',
}
COR_ALT_ROW = 'FFF2F2F2'

THIN = Side(style='thin', color='FFCCCCCC')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _header_style(cell, bg='FF1F3864', fg='FFFFFFFF', bold=True, size=10):
    cell.font = Font(name='Arial', bold=bold, color=fg, size=size)
    cell.fill = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER


def _data_style(cell, alt=False, bold=False, align='center'):
    cell.font = Font(name='Arial', size=9, bold=bold)
    if alt:
        cell.fill = PatternFill('solid', start_color=COR_ALT_ROW)
    cell.alignment = Alignment(horizontal=align, vertical='center')
    cell.border = BORDER


def _auto_width(ws, extra=4):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + extra, 50)


def _criar_aba_pais(wb: Workbook, df_pais: pd.DataFrame, nome_pais: str):
    ws = wb.create_sheet(nome_pais)
    ws.freeze_panes = 'A3'

    # Título
    ws.merge_cells('A1:I1')
    titulo = ws['A1']
    titulo.value = f'📊 Análise de Mercado — {nome_pais} (últimos {DIAS_ATRAS} dias)'
    titulo.font = Font(name='Arial', bold=True, size=13, color='FFFFFFFF')
    titulo.fill = PatternFill('solid', start_color='FF1F3864')
    titulo.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    # Cabeçalho
    headers = ['Nível', 'Categoria', 'Tecnologia', 'Total', 'Remoto', 'Híbrido', 'Presencial', '% Remoto', 'Rank por Categoria']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        _header_style(c)
    ws.row_dimensions[2].height = 22

    # Dados ordenados
    df_sorted = df_pais.sort_values(['Nível', 'Categoria', 'Total'], ascending=[True, True, False])

    row = 3
    for _, r in df_sorted.iterrows():
        alt = (row % 2 == 0)
        nivel_cor = COR_NIVEL.get(r['Nível'], 'FF555555')
        values = [r['Nível'], r['Categoria'], r['Tecnologia'],
                  r['Total'], r['Remoto'], r['Híbrido'], r['Presencial'], r['% Remoto'], '']
        for col, val in enumerate(values, 1):
            c = ws.cell(row=row, column=col, value=val)
            _data_style(c, alt=alt, align='left' if col in (1,2,3) else 'center')
            if col == 1:  # Cor por nível
                c.font = Font(name='Arial', size=9, bold=True, color=nivel_cor)
            if col == 8:
                c.number_format = '0.0"%"'
        row += 1

    # Rank por categoria (coluna I) — fórmula RANK dentro do grupo é complexo; usamos rank via pandas
    # Adicionamos rank como dado pré-calculado
    row = 3
    for _, r in df_sorted.iterrows():
        # Rank dentro do mesmo nível+categoria
        subset = df_sorted[(df_sorted['Nível'] == r['Nível']) & (df_sorted['Categoria'] == r['Categoria'])]
        rank = list(subset['Tecnologia']).index(r['Tecnologia']) + 1
        c = ws.cell(row=row, column=9, value=f"#{rank}")
        _data_style(c, alt=(row % 2 == 0))
        row += 1

    # Heatmap na coluna Total
    if row > 3:
        ws.conditional_formatting.add(
            f'D3:D{row-1}',
            ColorScaleRule(
                start_type='min', start_color='FFFFFFFF',
                end_type='max',   end_color='FF1F3864'
            )
        )

    _auto_width(ws)
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 22


def _criar_aba_todos(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet('Todos os Dados', 0)
    ws.freeze_panes = 'A3'

    ws.merge_cells('A1:J1')
    t = ws['A1']
    t.value = f'🌎 Análise Completa — Brasil + EUA | Todos os Níveis e Tecnologias'
    t.font = Font(name='Arial', bold=True, size=13, color='FFFFFFFF')
    t.fill = PatternFill('solid', start_color='FF1F3864')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    headers = ['País', 'Nível', 'Categoria', 'Tecnologia', 'Total', 'Remoto', 'Híbrido', 'Presencial', '% Remoto']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        _header_style(c)
    ws.row_dimensions[2].height = 22

    df_sorted = df.sort_values(['País', 'Nível', 'Categoria', 'Total'], ascending=[True, True, True, False])
    for row, (_, r) in enumerate(df_sorted.iterrows(), start=3):
        alt = (row % 2 == 0)
        for col, val in enumerate([r['País'], r['Nível'], r['Categoria'], r['Tecnologia'],
                                    r['Total'], r['Remoto'], r['Híbrido'], r['Presencial'], r['% Remoto']], 1):
            c = ws.cell(row=row, column=col, value=val)
            _data_style(c, alt=alt, align='left' if col <= 4 else 'center')
            if col == 2:
                c.font = Font(name='Arial', size=9, bold=True, color=COR_NIVEL.get(str(val), 'FF555555'))
            if col == 9:
                c.number_format = '0.0"%"'

    if row > 3:
        ws.conditional_formatting.add(f'E3:E{row}', ColorScaleRule(
            start_type='min', start_color='FFFFFFFF',
            end_type='max', end_color='FF1F3864'
        ))

    _auto_width(ws)
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 24
    ws.column_dimensions['D'].width = 22


def _criar_aba_comparativo(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet('Comparativo BR vs EUA')
    ws.freeze_panes = 'A3'

    ws.merge_cells('A1:H1')
    t = ws['A1']
    t.value = '🆚 Comparativo BR vs EUA — Demanda por Tecnologia (todos os níveis)'
    t.font = Font(name='Arial', bold=True, size=13, color='FFFFFFFF')
    t.fill = PatternFill('solid', start_color='FF1F3864')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    pivot = (
        df.groupby(['Categoria', 'Tecnologia', 'País'])['Total']
        .sum()
        .unstack('País')
        .fillna(0)
        .astype(int)
        .reset_index()
    )
    paises = [c for c in pivot.columns if c not in ('Categoria', 'Tecnologia')]
    pivot['Total Geral'] = pivot[paises].sum(axis=1)
    pivot = pivot.sort_values(['Categoria', 'Total Geral'], ascending=[True, False])

    headers = ['Categoria', 'Tecnologia'] + paises + ['Total Geral']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        _header_style(c)
    ws.row_dimensions[2].height = 22

    for row, (_, r) in enumerate(pivot.iterrows(), start=3):
        alt = (row % 2 == 0)
        vals = [r['Categoria'], r['Tecnologia']] + [r[p] for p in paises] + [r['Total Geral']]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col, value=val)
            _data_style(c, alt=alt, align='left' if col <= 2 else 'center')
            if col == len(vals):
                c.font = Font(name='Arial', size=9, bold=True)

    if row > 3:
        n_paises = len(paises)
        ws.conditional_formatting.add(f'C3:C{row}', ColorScaleRule(
            start_type='min', start_color='FFFFFFFF', end_type='max', end_color='FF1A5276'
        ))
        if n_paises > 1:
            ws.conditional_formatting.add(f'D3:D{row}', ColorScaleRule(
                start_type='min', start_color='FFFFFFFF', end_type='max', end_color='FF1A5276'
            ))

    _auto_width(ws)
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 22


def _criar_aba_resumo(wb: Workbook, df: pd.DataFrame, dist_por_pais: dict):
    ws = wb.create_sheet('Resumo Executivo', 0)

    ws.merge_cells('A1:F1')
    t = ws['A1']
    t.value = '📈 Resumo Executivo — Panorama de Mercado Tech'
    t.font = Font(name='Arial', bold=True, size=14, color='FFFFFFFF')
    t.fill = PatternFill('solid', start_color='FF1F3864')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 32

    row = 3
    for country, cfg in PAISES.items():
        nome = cfg['nome']
        dist = dist_por_pais.get(nome, {})
        total = sum(dist.values())

        # Título do país
        ws.merge_cells(f'A{row}:F{row}')
        c = ws.cell(row=row, column=1, value=f'🌎 {nome} — {total} vagas coletadas')
        c.font = Font(name='Arial', bold=True, size=11, color='FFFFFFFF')
        c.fill = PatternFill('solid', start_color='FF2E75B6')
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[row].height = 22
        row += 1

        # Distribuição por nível
        for nivel in ['Sênior', 'Pleno', 'Geral', 'Junior', 'Estágio/Intern']:
            qtd = dist.get(nivel, 0)
            pct = qtd / total * 100 if total else 0
            alt = (row % 2 == 0)
            for col, val in enumerate([nivel, qtd, f'{pct:.1f}%'], 1):
                c = ws.cell(row=row, column=col, value=val)
                _data_style(c, alt=alt, align='left' if col == 1 else 'center')
                if col == 1:
                    c.font = Font(name='Arial', size=9, bold=True, color=COR_NIVEL.get(nivel, 'FF333333'))
            row += 1
        row += 1

    # Top 10 geral por país
    row += 1
    ws.merge_cells(f'A{row}:F{row}')
    c = ws.cell(row=row, column=1, value='🏆 Top 10 Tecnologias por País (todos os níveis)')
    c.font = Font(name='Arial', bold=True, size=11, color='FFFFFFFF')
    c.fill = PatternFill('solid', start_color='FF1F3864')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 22
    row += 1

    col_offset = 1
    for cfg in PAISES.values():
        nome = cfg['nome']
        top10 = (
            df[df['País'] == nome]
            .groupby('Tecnologia')['Total']
            .sum()
            .nlargest(10)
            .reset_index()
        )
        ws.cell(row=row, column=col_offset, value=nome).font = Font(name='Arial', bold=True, size=10)
        ws.cell(row=row, column=col_offset+1, value='Vagas').font = Font(name='Arial', bold=True, size=10)
        for i, (_, r) in enumerate(top10.iterrows(), start=1):
            ws.cell(row=row+i, column=col_offset, value=r['Tecnologia'])
            ws.cell(row=row+i, column=col_offset+1, value=r['Total'])
        col_offset += 3

    _auto_width(ws)


def exportar_excel(df: pd.DataFrame, nome_arquivo: str, dist_por_pais: dict) -> None:
    wb = Workbook()
    # Remove aba padrão
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    _criar_aba_resumo(wb, df, dist_por_pais)
    _criar_aba_todos(wb, df)

    for country, cfg in PAISES.items():
        nome_pais = cfg['nome']
        df_pais = df[df['País'] == nome_pais]
        if not df_pais.empty:
            _criar_aba_pais(wb, df_pais, nome_pais)

    _criar_aba_comparativo(wb, df)

    wb.save(nome_arquivo)
    print(f"\n✅  Arquivo exportado: {nome_arquivo}")
    abas = ' | '.join(wb.sheetnames)
    print(f"   Abas: {abas}")

# ==========================================
# 9. Main
# ==========================================
def analisar_texto_vagas() -> None:
    print("=" * 60)
    print("🚀  EXTRAÇÃO AVANÇADA DE VAGAS — BR + EUA  🚀")
    print("=" * 60)

    todos_dados: list[dict] = []
    dist_por_pais: dict[str, dict] = {}

    for country, cfg in PAISES.items():
        print(f"\n\n{'─'*60}")
        print(f"🌎  Coletando vagas: {cfg['nome']} ({country.upper()})")
        print(f"{'─'*60}")

        vagas = obter_vagas_pais(country, cfg['termos'])
        dados, dist = processar_pais(country, vagas)
        todos_dados.extend(dados)
        dist_por_pais[cfg['nome']] = dist

        total = sum(dist.values())
        print(f"\n📊 Distribuição por nível — {cfg['nome']}:")
        for nivel, qtd in sorted(dist.items(), key=lambda x: -x[1]):
            pct = qtd / total * 100 if total else 0
            print(f"   {nivel:<20} {qtd:>4} vagas ({pct:.1f}%)")

    df = pd.DataFrame(todos_dados)
    if df.empty:
        print("\nNenhum dado encontrado.")
        return

    for cfg in PAISES.values():
        imprimir_relatorio(df, cfg['nome'])

    exportar_excel(df, 'analise_mercado_br_eua.xlsx', dist_por_pais)


if __name__ == "__main__":
    analisar_texto_vagas()