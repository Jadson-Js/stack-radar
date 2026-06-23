#!/usr/bin/env python3
"""
generate_data.py
Lê o analise_mercado_br_eua_v1.xlsx e gera data.json
para o dashboard index.html consumir dinamicamente.
"""
import json
import re
from collections import defaultdict
import pandas as pd

XLSX_PATH = "analise_mercado_br_eua_v1.xlsx"
OUTPUT_PATH = "data.json"
TOP_N = 15        # número de itens nos rankings
TOP_CAT = 10      # itens por categoria
NIVEIS = ["Geral", "Sênior", "Pleno", "Junior", "Estágio/Intern"]

# ── helpers ────────────────────────────────────────────────────────────────
def read_sheet(xl: pd.ExcelFile, sheet: str, skip_rows: int = 1) -> pd.DataFrame:
    """Lê uma aba pulando a primeira linha (título mesclado) e renomeia colunas."""
    df = xl.parse(sheet, skiprows=skip_rows)
    # A primeira linha real de dados vira o header
    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
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

    # === Aba "Todos os Dados" ===
    all_df = read_sheet(xl, "Todos os Dados")
    all_df.columns = ["País", "Nível", "Categoria", "Tecnologia",
                      "Total", "Remoto", "Híbrido", "Presencial", "% Remoto"]
    for col in ["Total", "Remoto", "Híbrido", "Presencial"]:
        all_df[col] = all_df[col].apply(safe_int)
    all_df["% Remoto"] = all_df["% Remoto"].apply(safe_float)

    br_df  = all_df[all_df["País"] == "Brasil"].copy()
    eua_df = all_df[all_df["País"] == "EUA"].copy()

    # ── KPIs ──────────────────────────────────────────────────────────────
    total_br  = safe_int(br_df["Total"].sum())
    total_eua = safe_int(eua_df["Total"].sum())

    # Distribuição de modalidade: soma por país (evita double-count por nível)
    # Usamos a aba Geral para modalidade para não duplicar
    def modal_pais(df: pd.DataFrame) -> dict:
        g = df.groupby("Nível")[["Remoto", "Híbrido", "Presencial"]].sum()
        # Soma de todos os níveis gera duplicatas por tecnologia contada em
        # vários níveis; usamos nível "Geral" como proxy de modalidade total.
        geral = g.loc["Geral"] if "Geral" in g.index else g.sum()
        return {
            "remoto":     safe_int(geral["Remoto"]),
            "hibrido":    safe_int(geral["Híbrido"]),
            "presencial": safe_int(geral["Presencial"]),
        }

    modal_br  = modal_pais(br_df)
    modal_eua = modal_pais(eua_df)

    total_modal_br  = sum(modal_br.values())  or 1
    total_modal_eua = sum(modal_eua.values()) or 1
    pct_remoto_br   = round(modal_br["remoto"]  / total_modal_br  * 100, 1)
    pct_remoto_eua  = round(modal_eua["remoto"] / total_modal_eua * 100, 1)

    # ── Top global (BR + EUA combinados, nível Geral para não duplicar) ──
    geral_all = all_df[all_df["Nível"] == "Geral"].groupby("Tecnologia")["Total"].sum()
    top_global = (
        geral_all.sort_values(ascending=False).head(TOP_N)
    )
    top_global_list = [{"t": t, "v": safe_int(v)} for t, v in top_global.items()]

    top_br_list  = top_list(br_df[br_df["Nível"] == "Geral"],  "Tecnologia", "Total", TOP_N)
    top_eua_list = top_list(eua_df[eua_df["Nível"] == "Geral"], "Tecnologia", "Total", TOP_N)

    # #1 stacks
    stack_global = top_global_list[0]["t"] if top_global_list else "-"
    stack_br     = top_br_list[0]["t"]     if top_br_list     else "-"
    stack_eua    = top_eua_list[0]["t"]    if top_eua_list    else "-"

    mencoes_global = top_global_list[0]["v"] if top_global_list else 0
    mencoes_br     = top_br_list[0]["v"]     if top_br_list     else 0
    mencoes_eua    = top_eua_list[0]["v"]    if top_eua_list    else 0

    # Stack mais remota (global, nível Geral, mínimo 5 ocorrências)
    rem_df = all_df[all_df["Nível"] == "Geral"].groupby("Tecnologia").agg(
        Total=("Total", "sum"), Remoto=("Remoto", "sum")
    )
    rem_df = rem_df[rem_df["Total"] >= 5].copy()
    rem_df["pct"] = (rem_df["Remoto"] / rem_df["Total"] * 100).round(1)
    top_remote_df = rem_df.sort_values("pct", ascending=False).head(14)
    top_remote = [
        {"t": t, "pct": row["pct"], "n": safe_int(row["Total"])}
        for t, row in top_remote_df.iterrows()
    ]


    # ── Por nível (global, BR, EUA) ──────────────────────────────────────
    def por_nivel_dict(df: pd.DataFrame) -> dict:
        result = {}
        for nivel in NIVEIS:
            sub = df[df["Nível"] == nivel]
            result[nivel] = top_list(sub, "Tecnologia", "Total", TOP_CAT)
        return result

    por_nivel     = por_nivel_dict(all_df)
    por_nivel_br  = por_nivel_dict(br_df)
    por_nivel_eua = por_nivel_dict(eua_df)

    # ── Distribuição de vagas por nível ──────────────────────────────────
    def niveis_vagas(df: pd.DataFrame) -> dict:
        """
        Estimamos o total de vagas por nível usando a aba Resumo Executivo,
        mas como é difícil de parsear, usamos a coluna Total da tecnologia
        mais frequente por nível como proxy.
        Em vez disso, somamos Total por tecnologia única no nível Geral de cada nível.
        """
        result = {}
        for nivel in NIVEIS:
            sub = df[df["Nível"] == nivel]
            # Total de menções (soma de todas as tecnologias — não é único por vaga)
            # Melhor: contar tecnologias distintas x média... mas sem a coluna de vagas reais
            # Usamos o máximo individual de uma tecnologia como proxy de volume
            result[nivel] = safe_int(sub["Total"].max()) if not sub.empty else 0
        return result

    # Lê do Resumo Executivo os valores reais de vagas por nível
    resumo_df = xl.parse("Resumo Executivo", skiprows=1)
    resumo_df.columns = [str(c).strip() for c in resumo_df.columns]

    def parse_resumo_niveis(xl: pd.ExcelFile) -> tuple[dict, dict]:
        """Extrai distribuição de vagas por nível do Resumo Executivo."""
        raw = xl.parse("Resumo Executivo", header=None)
        niveis_br, niveis_eua = {}, {}
        pais_atual = None
        for _, row in raw.iterrows():
            cell0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            cell1 = row.iloc[1] if pd.notna(row.iloc[1]) else None

            if "Brasil" in cell0:
                pais_atual = "br"
            elif "EUA" in cell0:
                pais_atual = "eua"

            for n in NIVEIS:
                if cell0 == n and cell1 is not None:
                    qtd = safe_int(cell1)
                    if pais_atual == "br":
                        niveis_br[n] = qtd
                    elif pais_atual == "eua":
                        niveis_eua[n] = qtd
        return niveis_br, niveis_eua

    niveis_vagas_br, niveis_vagas_eua = parse_resumo_niveis(xl)

    # ── Categorias ────────────────────────────────────────────────────────
    cats_br  = br_df[br_df["Nível"]  == "Geral"].groupby("Categoria")["Total"].sum().to_dict()
    cats_eua = eua_df[eua_df["Nível"] == "Geral"].groupby("Categoria")["Total"].sum().to_dict()
    cats_br  = {k: safe_int(v) for k, v in cats_br.items()}
    cats_eua = {k: safe_int(v) for k, v in cats_eua.items()}

    # ── Por categoria (global, BR, EUA) ──────────────────────────────────
    def por_cat_dict(df: pd.DataFrame) -> dict:
        cats = df["Categoria"].dropna().unique()
        result = {}
        for cat in cats:
            sub = df[df["Categoria"] == cat]
            result[cat] = top_list(sub, "Tecnologia", "Total", TOP_CAT)
        return result

    por_cat     = por_cat_dict(all_df)
    por_cat_br  = por_cat_dict(br_df)
    por_cat_eua = por_cat_dict(eua_df)



    # ── Total de vagas reais (Resumo Executivo) ──────────────────────────
    def parse_total_vagas(xl: pd.ExcelFile) -> tuple[int, int]:
        raw = xl.parse("Resumo Executivo", header=None)
        total_br, total_eua = 0, 0
        for _, row in raw.iterrows():
            cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            m = re.search(r"(\d[\d.]+)\s*vagas", cell)
            if m:
                n = safe_int(m.group(1).replace(".", ""))
                if "Brasil" in cell:
                    total_br = n
                elif "EUA" in cell:
                    total_eua = n
        return total_br, total_eua

    real_total_br, real_total_eua = parse_total_vagas(xl)
    if real_total_br == 0: real_total_br = total_br
    if real_total_eua == 0: real_total_eua = total_eua

    # ── Percentual Sênior / Estágio para KPI sub ──────────────────────────
    def pct_nivel(niveis_dict: dict, nivel: str, total: int) -> float:
        return round(niveis_dict.get(nivel, 0) / total * 100, 1) if total else 0.0

    pct_senior_br    = pct_nivel(niveis_vagas_br,  "Sênior",        real_total_br)
    pct_estagio_br   = pct_nivel(niveis_vagas_br,  "Estágio/Intern", real_total_br)
    pct_senior_eua   = pct_nivel(niveis_vagas_eua, "Sênior",        real_total_eua)
    pct_estagio_eua  = pct_nivel(niveis_vagas_eua, "Estágio/Intern", real_total_eua)

    # ── Monta JSON final ──────────────────────────────────────────────────
    data = {
        # KPIs
        "kpi": {
            "totalBR":         real_total_br,
            "totalEUA":        real_total_eua,
            "totalGeral":      real_total_br + real_total_eua,
            "pctRemotoBR":     pct_remoto_br,
            "pctRemotoEUA":    pct_remoto_eua,
            "stackGlobal":     stack_global,
            "mencoesGlobal":   mencoes_global,
            "stackBR":         stack_br,
            "mencoesBR":       mencoes_br,
            "stackEUA":        stack_eua,
            "mencoesEUA":      mencoes_eua,

            "pctSeniorBR":     pct_senior_br,
            "pctEstagioBR":    pct_estagio_br,
            "pctSeniorEUA":    pct_senior_eua,
            "pctEstagioEUA":   pct_estagio_eua,
        },
        # Rankings globais
        "topGlobal": top_global_list,
        "topBR":     top_br_list,
        "topEUA":    top_eua_list,
        # Por nível
        "porNivel":    por_nivel,
        "porNivelBR":  por_nivel_br,
        "porNivelEUA": por_nivel_eua,
        # Distribuição de vagas por nível
        "niveisVagasBR":  niveis_vagas_br,
        "niveisVagasEUA": niveis_vagas_eua,
        # Modalidade
        "modalBR":  modal_br,
        "modalEUA": modal_eua,
        # Categorias
        "catBR":  cats_br,
        "catEUA": cats_eua,
        # Top remoto
        "topRemote": top_remote,

        # Por categoria
        "porCategoria":    por_cat,
        "porCategoriaBR":  por_cat_br,
        "porCategoriaEUA": por_cat_eua,
    }

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅  {OUTPUT_PATH} gerado com sucesso!")
    print(f"   Total BR: {real_total_br:,} vagas | EUA: {real_total_eua:,} vagas")
    print(f"   #1 Global: {stack_global} ({mencoes_global} menções)")
    print(f"   #1 BR: {stack_br} | #1 EUA: {stack_eua}")



if __name__ == "__main__":
    main()
