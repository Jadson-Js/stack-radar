#!/usr/bin/env python3
"""
generate_data.py
Lê o market_analysis_br_usa.xlsx e gera data.json
para o dashboard index.html consumir dinamicamente.
"""
import json
import re
import pandas as pd

XLSX_PATH = "market_analysis_br_usa.xlsx"
OUTPUT_PATH = "data.json"
TOP_N = 15        # número de itens nos rankings
TOP_CAT = 10      # itens por categoria
LEVELS = ["General", "Senior", "Mid-level", "Junior", "Internship"]

# ── helpers ────────────────────────────────────────────────────────────────
def read_sheet(xl: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    """Lê a aba definindo a linha 1 (segunda linha) como o cabeçalho real."""
    # header=1 diz para o pandas usar a segunda linha do Excel (índice 1) como colunas,
    # ignorando automaticamente a primeira linha mesclada de título.
    df = xl.parse(sheet, header=1)
    
    # Remove colunas fantasmas geradas por células vazias no Excel (ex: Unnamed: 9)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed:', na=True)]
    df.columns = [str(c).strip() for c in df.columns]
    
    # Remover linhas totalmente nulas
    df = df.dropna(how="all")
    return df


def safe_int(v):
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return 0


def safe_float(v):
    try:
        return round(float(v), 1)
    except (ValueError, TypeError):
        return 0.0


def top_list(df: pd.DataFrame, col_tech: str, col_val: str, n: int) -> list[dict]:
    """Retorna lista [{t, v}, ...] ordenada por col_val descendente."""
    if df.empty:
        return []
    out = (
        df.groupby(col_tech)[col_val]
        .sum()
        .sort_values(ascending=False)
        .head(n)
    )
    return [{"t": t, "v": safe_int(v)} for t, v in out.items()]


# ── main ───────────────────────────────────────────────────────────────────
def main():
    xl = pd.ExcelFile(XLSX_PATH)

    # === Aba "Todos os Dados" / "All Data" ===
    sheet_name = "All Data" if "All Data" in xl.sheet_names else xl.sheet_names[0]
    all_df = read_sheet(xl, sheet_name)
    
    # Dicionário de normalização de colunas (Mantendo tudo em Inglês)
    col_map = {
        "Country": "Country", "País": "Country",
        "Level": "Level", "Nível": "Level",
        "Category": "Category", "Categoria": "Category",
        "Technology": "Technology", "Tecnologia": "Technology",
        "Total": "Total",
        "Remote": "Remote", "Remoto": "Remote",
        "Hybrid": "Hybrid", "Híbrido": "Hybrid",
        "On-site": "On-site", "Presencial": "On-site",
        "% Remote": "% Remote", "% Remoto": "% Remote"
    }
    all_df = all_df.rename(columns=col_map)
    
    # Sanitização contra espaços invisíveis que quebram as Categorias e Níveis
    for col in ["Country", "Level", "Category", "Technology"]:
        if col in all_df.columns:
            all_df[col] = all_df[col].astype(str).str.strip()

    for col in ["Total", "Remote", "Hybrid", "On-site"]:
        if col in all_df.columns:
            all_df[col] = all_df[col].apply(safe_int)
            
    if "% Remote" in all_df.columns:
        all_df["% Remote"] = all_df["% Remote"].apply(safe_float)

    # Normalizar valores internos das células de português para inglês
    country_map = {"Brasil": "Brazil", "Brazil": "Brazil", "USA": "USA", "EUA": "USA"}
    level_map = {
        "General": "General", "Senior": "Senior", "Mid-level": "Mid-level",
        "Junior": "Junior", "Internship": "Internship",
        "Geral": "General", "Sênior": "Senior", "Pleno": "Mid-level",
        "Estágio/Intern": "Internship", "Estágio": "Internship"
    }
    all_df["Country"] = all_df["Country"].map(lambda x: country_map.get(x, x))
    all_df["Level"] = all_df["Level"].map(lambda x: level_map.get(x, x))

    br_df  = all_df[all_df["Country"] == "Brazil"].copy()
    usa_df = all_df[all_df["Country"] == "USA"].copy()

    # ── KPIs de Modalidade (Filtra estritamente por Nível Geral/General para evitar double-count) ──
    def modal_pais(df: pd.DataFrame) -> dict:
        sub_general = df[df["Level"] == "General"]
        if sub_general.empty:
            return {"remote": 0, "hybrid": 0, "onsite": 0}
        return {
            "remote":     safe_int(sub_general["Remote"].sum()),
            "hybrid":    safe_int(sub_general["Hybrid"].sum()),
            "onsite": safe_int(sub_general["On-site"].sum()),
        }

    modal_br  = modal_pais(br_df)
    modal_usa = modal_pais(usa_df)

    total_modal_br  = sum(modal_br.values())  or 1
    total_modal_usa = sum(modal_usa.values()) or 1
    pct_remote_br   = round(modal_br["remote"]  / total_modal_br  * 100, 1)
    pct_remote_usa  = round(modal_usa["remote"] / total_modal_usa * 100, 1)

    # ── Rankings Globais ──
    general_all = all_df[all_df["Level"] == "General"].groupby("Technology")["Total"].sum()
    top_global = general_all.sort_values(ascending=False).head(TOP_N)
    top_global_list = [{"t": t, "v": safe_int(v)} for t, v in top_global.items()]

    top_br_list  = top_list(br_df[br_df["Level"] == "General"],  "Technology", "Total", TOP_N)
    top_usa_list = top_list(usa_df[usa_df["Level"] == "General"], "Technology", "Total", TOP_N)

    # #1 Stacks e Menções
    stack_global   = top_global_list[0]["t"] if top_global_list else "-"
    mencoes_global = top_global_list[0]["v"] if top_global_list else 0
    stack_br       = top_br_list[0]["t"]     if top_br_list     else "-"
    mencoes_br     = top_br_list[0]["v"]     if top_br_list     else 0
    stack_usa      = top_usa_list[0]["t"]    if top_usa_list    else "-"
    mencoes_usa    = top_usa_list[0]["v"]    if top_usa_list    else 0

    # Stack mais remota (Mínimo de 5 ocorrências)
    rem_df = all_df[all_df["Level"] == "General"].groupby("Technology").agg(
        Total=("Total", "sum"), Remote=("Remote", "sum")
    )
    rem_df = rem_df[rem_df["Total"] >= 5].copy()
    rem_df["pct"] = (rem_df["Remote"] / rem_df["Total"] * 100).round(1)
    top_remote_df = rem_df.sort_values("pct", ascending=False).head(14)
    top_remote = [
        {"t": t, "pct": row["pct"], "n": safe_int(row["Total"])}
        for t, row in top_remote_df.iterrows()
    ]

    # ── Rankings Por Nível ──
    def por_nivel_dict(df: pd.DataFrame) -> dict:
        result = {}
        for level in LEVELS:
            sub = df[df["Level"] == level]
            result[level] = top_list(sub, "Technology", "Total", TOP_CAT)
        return result

    by_level     = por_nivel_dict(all_df)
    by_level_br  = por_nivel_dict(br_df)
    by_level_usa = por_nivel_dict(usa_df)

    # ── Distribuição de Vagas por Nível (Resumo Executivo) ──
    def parse_resumo_niveis(xl: pd.ExcelFile) -> tuple[dict, dict]:
        raw = xl.parse("Executive Summary", header=None)
        niveis_br, niveis_usa = {}, {}
        pais_atual = None
        
        exec_level_map = {
            "General": "General", "Senior": "Senior", "Mid-level": "Mid-level",
            "Junior": "Junior", "Internship": "Internship",
            "Geral": "General", "Sênior": "Senior", "Pleno": "Mid-level",
            "Estágio/Intern": "Internship", "Estágio": "Internship"
        }
        
        for _, row in raw.iterrows():
            cell0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            cell1 = row.iloc[1] if pd.notna(row.iloc[1]) else None

            if "brazil" in cell0.lower() or "brasil" in cell0.lower():
                pais_atual = "br"
                continue
            elif "usa" in cell0.lower() or "eua" in cell0.lower():
                pais_atual = "usa"
                continue

            mapped_n = exec_level_map.get(cell0)
            if mapped_n is not None and cell1 is not None:
                qtd = safe_int(cell1)
                if pais_atual == "br":
                    niveis_br[mapped_n] = qtd
                elif pais_atual == "usa":
                    niveis_usa[mapped_n] = qtd
        return niveis_br, niveis_usa

    niveis_vagas_br, niveis_vagas_usa = parse_resumo_niveis(xl)

    # ── Agrupamento de Categorias Limpo ──
    br_df_clean = br_df[~br_df["Category"].isin(["nan", "None", ""])]
    usa_df_clean = usa_df[~usa_df["Category"].isin(["nan", "None", ""])]
    all_df_clean = all_df[~all_df["Category"].isin(["nan", "None", ""])]

    cats_br  = br_df_clean[br_df_clean["Level"] == "General"].groupby("Category")["Total"].sum().to_dict()
    cats_usa = usa_df_clean[usa_df_clean["Level"] == "General"].groupby("Category")["Total"].sum().to_dict()
    cats_br  = {k: safe_int(v) for k, v in cats_br.items()}
    cats_usa = {k: safe_int(v) for k, v in cats_usa.items()}

    # ── Rankings Por Categoria ──
    def por_cat_dict(df: pd.DataFrame) -> dict:
        cats = df["Category"].dropna().unique()
        result = {}
        for cat in cats:
            sub = df[df["Category"] == cat]
            result[cat] = top_list(sub, "Technology", "Total", TOP_CAT)
        return result

    by_cat     = por_cat_dict(all_df_clean)
    by_cat_br  = por_cat_dict(br_df_clean)
    by_cat_usa = por_cat_dict(usa_df_clean)

    # ── Total de Vagas Reais do Mercado ──
    def parse_total_vagas(xl: pd.ExcelFile) -> tuple[int, int]:
        raw = xl.parse("Executive Summary", header=None)
        total_br, total_usa = 0, 0
        for _, row in raw.iterrows():
            cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            m = re.search(r"([\d.,]+)\s*(?:vagas|collected jobs)", cell, re.IGNORECASE)
            if m:
                n = safe_int(m.group(1).replace(".", "").replace(",", ""))
                if "brazil" in cell.lower() or "brasil" in cell.lower():
                    total_br = n
                elif "usa" in cell.lower() or "eua" in cell.lower():
                    total_usa = n
        return total_br, total_usa

    real_total_br, real_total_usa = parse_total_vagas(xl)
    
    # Fallback de segurança para evitar divisão por zero
    if real_total_br == 0: 
        real_total_br = safe_int(br_df[br_df["Level"] == "General"]["Total"].max()) or 1
    if real_total_usa == 0: 
        real_total_usa = safe_int(usa_df[usa_df["Level"] == "General"]["Total"].max()) or 1

    # ── Percentual de Sênior / Estágio ──
    def pct_nivel(niveis_dict: dict, nivel: str, total: int) -> float:
        return round(niveis_dict.get(nivel, 0) / total * 100, 1) if total else 0.0

    pct_senior_br   = pct_nivel(niveis_vagas_br,  "Senior",      real_total_br)
    pct_estagio_br  = pct_nivel(niveis_vagas_br,  "Internship",  real_total_br)
    pct_senior_usa  = pct_nivel(niveis_vagas_usa, "Senior",      real_total_usa)
    pct_estagio_usa = pct_nivel(niveis_vagas_usa, "Internship",  real_total_usa)

    # ── Monta o Payload do JSON final ──
    data = {
        "kpi": {
            "totalBR":         real_total_br,
            "totalUSA":        real_total_usa,
            "totalGlobal":     real_total_br + real_total_usa,
            "pctRemoteBR":     pct_remote_br,
            "pctRemoteUSA":    pct_remote_usa,
            "stackGlobal":     stack_global,
            "mentionsGlobal":   mencoes_global,
            "stackBR":         stack_br,
            "mentionsBR":       mencoes_br,
            "stackUSA":        stack_usa,
            "mentionsUSA":      mencoes_usa,
            "pctSeniorBR":     pct_senior_br,
            "pctInternshipBR":    pct_estagio_br,
            "pctSeniorUSA":    pct_senior_usa,
            "pctInternshipUSA":   pct_estagio_usa,
        },
        "topGlobal": top_global_list,
        "topBR":     top_br_list,
        "topUSA":    top_usa_list,
        "byLevel":    by_level,
        "byLevelBR":  by_level_br,
        "byLevelUSA": by_level_usa,
        "levelsJobsBR":  niveis_vagas_br,
        "levelsJobsUSA": niveis_vagas_usa,
        "modalityBR":  modal_br,
        "modalityUSA": modal_usa,
        "categoryBR":  cats_br,
        "categoryUSA": cats_usa,
        "topRemote": top_remote,
        "byCategory":    by_cat,
        "byCategoryBR":  by_cat_br,
        "byCategoryUSA": by_cat_usa,
    }

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅ {OUTPUT_PATH} gerado com sucesso!")
    print(f"   Categorias ativas mapeadas no JSON -> BR: {len(cats_br)} | USA: {len(cats_usa)}")
    print(f"   Total Real das Vagas -> BR: {real_total_br:,} | USA: {real_total_usa:,}")

if __name__ == "__main__":
    main()