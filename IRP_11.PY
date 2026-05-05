import warnings
warnings.filterwarnings("ignore")

import pandas as pd
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# CONFIGURAÇÃO
# ============================================================

PASTA_DADOS = Path(r"C:\Users\vitorL\OneDrive - Fundação Instituto de Administração\Área de Trabalho\Produção de O&G\Questão 1")
ARQUIVO_SAIDA = PASTA_DADOS / "Indicadores_IRP1.xlsx"

FATOR_BARRIL = 6.289941

BACIAS_PMCRP = ["Campos", "Santos", "Espírito Santo"]
ESTADOS_PMCRP = ["Espírito Santo", "Rio de Janeiro", "São Paulo", "Santa Catarina"]

# Todos os anos de 2010 até 2026, inclusive, com Mar + Terra
ANOS_MAR_TERRA = list(range(2010, 2027))
ANOS_FINAIS = ANOS_MAR_TERRA

# ============================================================
# FUNÇÕES
# ============================================================

def limpar_nome_coluna(col):
    c = str(col).strip().replace("\ufeff", "").replace("[", "").replace("]", "")
    mapa = {
        "Ano": "Ano",
        "Mês/Ano": "Mês/Ano",
        "Mes/Ano": "Mês/Ano",
        "M�s/Ano": "Mês/Ano",
        "Estado": "Estado",
        "Bacia": "Bacia",
        "Campo": "Campo",
        "Poço": "Poço",
        "Po�o": "Poço",
        "Ambiente": "Ambiente",
        "Instalação": "Instalação",
        "Instala��o": "Instalação",
        "Produção de Óleo (m³)": "Produção de Óleo (m³)",
        "Produ��o de �leo (m�)": "Produção de Óleo (m³)",
        "Produção de Condensado (m³)": "Produção de Condensado (m³)",
        "Produ��o de Condensado (m�)": "Produção de Condensado (m³)",
        "Produção de Gás Associado (Mm³)": "Produção de Gás Associado (Mm³)",
        "Produ��o de G�s Associado (Mm�)": "Produção de Gás Associado (Mm³)",
        "Produção de Gás Não Associado (Mm³)": "Produção de Gás Não Associado (Mm³)",
        "Produ��o de G�s N�o Associado (Mm�)": "Produção de Gás Não Associado (Mm³)",
    }
    return mapa.get(c, c)

def normalizar_colunas(df):
    df.columns = [limpar_nome_coluna(c) for c in df.columns]
    return df

def parse_numero(valor):
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip()

    if s in ["", "nan", "None", "-", " "]:
        return 0.0

    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return 0.0

def normalizar_ambiente(valor):
    s = str(valor).strip().lower()

    if s in ["mar", "maritimo", "marítimo", "offshore"]:
        return "Mar"
    if s in ["terra", "terrestre", "onshore"]:
        return "Terra"

    if "mar" in s:
        return "Mar"
    if "terra" in s:
        return "Terra"

    return str(valor).strip()

def ler_excel(arquivo):
    try:
        xls = pd.ExcelFile(arquivo)
        dfs = []

        for aba in xls.sheet_names:
            try:
                df = pd.read_excel(arquivo, sheet_name=aba)
                df = normalizar_colunas(df)

                if "Ano" not in df.columns and "Mês/Ano" in df.columns:
                    df["Ano"] = df["Mês/Ano"].astype(str).str[-4:]

                if "Ano" in df.columns:
                    df["Arquivo_Origem"] = arquivo.name
                    df["Aba_Origem"] = aba
                    dfs.append(df)

            except Exception:
                continue

        if dfs:
            return pd.concat(dfs, ignore_index=True)

        return None

    except Exception as e:
        print(f"Erro ao ler {arquivo.name}: {e}")
        return None

def formatar_aba(writer, nome_aba):
    ws = writer.sheets[nome_aba]

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cab = str(ws.cell(1, cell.column).value)
            if isinstance(cell.value, (int, float)):
                if "Proporção" in cab or "pelo Total" in cab:
                    # Mantém percentual com alta precisão visual
                    cell.number_format = "0.0000000000%"
                else:
                    # Mantém alta precisão visual sem alterar valor real
                    cell.number_format = "#,##0.0000000000"

# ============================================================
# ETAPA 1 - LER ARQUIVOS EXCEL
# ============================================================

print("=" * 80)
print("ETAPA 1 - LENDO ARQUIVOS EXCEL")
print("=" * 80)

arquivos = sorted(
    list(PASTA_DADOS.rglob("*.xlsx")) +
    list(PASTA_DADOS.rglob("*.xls"))
)

print(f"Arquivos Excel encontrados: {len(arquivos)}")
for a in arquivos[:20]:
    print(" -", a.name)
if len(arquivos) > 20:
    print(f"... e mais {len(arquivos) - 20} arquivos")

bases = []

for i, arquivo in enumerate(arquivos, start=1):
    print(f"\n[{i}/{len(arquivos)}] {arquivo.name}")

    df_temp = ler_excel(arquivo)

    if df_temp is None:
        print(" -> ignorado")
        continue

    print(f" -> linhas lidas: {len(df_temp)}")
    print(f" -> colunas: {list(df_temp.columns[:12])}")
    bases.append(df_temp)

if not bases:
    raise ValueError("Nenhum Excel válido foi lido.")

df_total = pd.concat(bases, ignore_index=True)
print(f"\nTotal de linhas consolidadas: {len(df_total)}")

# ============================================================
# ETAPA 2 - TRATAMENTO DA BASE
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 2 - TRATAMENTO DA BASE")
print("=" * 80)

colunas_necessarias = [
    "Ano",
    "Estado",
    "Bacia",
    "Ambiente",
    "Produção de Óleo (m³)",
    "Produção de Condensado (m³)",
    "Produção de Gás Associado (Mm³)",
    "Produção de Gás Não Associado (Mm³)"
]

for col in colunas_necessarias:
    if col not in df_total.columns:
        df_total[col] = "" if col in ["Estado", "Bacia", "Ambiente"] else 0

for col in [
    "Produção de Óleo (m³)",
    "Produção de Condensado (m³)",
    "Produção de Gás Associado (Mm³)",
    "Produção de Gás Não Associado (Mm³)"
]:
    print(f"Tratando coluna: {col}")
    df_total[col] = df_total[col].apply(parse_numero)

df_total["Ano"] = pd.to_numeric(df_total["Ano"], errors="coerce")
df_total["Estado"] = df_total["Estado"].astype(str).str.strip()
df_total["Bacia"] = df_total["Bacia"].astype(str).str.strip()
df_total["Ambiente"] = df_total["Ambiente"].apply(normalizar_ambiente)

print("\nValores únicos de Ambiente antes do recorte:")
print(sorted(df_total["Ambiente"].dropna().unique().tolist())[:30])

# ============================================================
# ETAPA 3 - APLICANDO RECORTE
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 3 - APLICANDO RECORTE")
print("=" * 80)

filtro_mar_terra = (
    df_total["Ano"].isin(ANOS_MAR_TERRA)
) & (
    df_total["Ambiente"].isin(["Mar", "Terra"])
)

df_total = df_total[filtro_mar_terra].copy()

print("Anos finais:", ANOS_FINAIS)

print("\nContagem por ano após recorte:")
print(df_total["Ano"].value_counts().sort_index())

print("\nContagem por ano e ambiente após recorte:")
print(df_total.groupby(["Ano", "Ambiente"]).size())

# ============================================================
# ETAPA 4 - AUDITORIA POR ANO E AMBIENTE
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 4 - AUDITORIA POR ANO E AMBIENTE")
print("=" * 80)

auditoria_ano_ambiente = (
    df_total.groupby(["Ano", "Ambiente"], as_index=False)
    .agg({
        "Produção de Óleo (m³)": "sum",
        "Produção de Condensado (m³)": "sum",
        "Produção de Gás Associado (Mm³)": "sum",
        "Produção de Gás Não Associado (Mm³)": "sum"
    })
    .sort_values(["Ano", "Ambiente"])
)

print(auditoria_ano_ambiente)

# ============================================================
# ETAPA 5 - AMOSTRAGEM DOS CÁLCULOS
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 5 - AMOSTRAGEM DOS CÁLCULOS")
print("=" * 80)

df_total["Óleo + Condensado (m³)"] = (
    df_total["Produção de Óleo (m³)"] + df_total["Produção de Condensado (m³)"]
)
df_total["Petróleo em Barril"] = df_total["Óleo + Condensado (m³)"] * FATOR_BARRIL

df_total["Gás Total (Mm³)"] = (
    df_total["Produção de Gás Associado (Mm³)"] + df_total["Produção de Gás Não Associado (Mm³)"]
)
df_total["Gás Equivalente (m³ petróleo)"] = df_total["Gás Total (Mm³)"] / 1000
df_total["Gás Natural em Barril"] = df_total["Gás Equivalente (m³ petróleo)"] * FATOR_BARRIL

df_total["Produção Total em Barril"] = (
    df_total["Petróleo em Barril"] + df_total["Gás Natural em Barril"]
)

print(df_total[[
    "Ano",
    "Ambiente",
    "Produção de Óleo (m³)",
    "Produção de Condensado (m³)",
    "Óleo + Condensado (m³)",
    "Petróleo em Barril",
    "Produção de Gás Associado (Mm³)",
    "Produção de Gás Não Associado (Mm³)",
    "Gás Total (Mm³)",
    "Gás Natural em Barril",
    "Produção Total em Barril"
]].head(20))

# ============================================================
# ETAPA 6 - IRP 1.1
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 6 - IRP 1.1")
print("=" * 80)

base_anos = pd.DataFrame({"Ano": ANOS_FINAIS})

irp11 = (
    df_total.groupby("Ano", as_index=False)
    .agg({
        "Produção de Óleo (m³)": "sum",
        "Produção de Condensado (m³)": "sum",
        "Óleo + Condensado (m³)": "sum",
        "Petróleo em Barril": "sum",
        "Produção de Gás Associado (Mm³)": "sum",
        "Produção de Gás Não Associado (Mm³)": "sum",
        "Gás Total (Mm³)": "sum",
        "Gás Natural em Barril": "sum",
        "Produção Total em Barril": "sum"
    })
    .sort_values("Ano")
)

irp11 = base_anos.merge(irp11, on="Ano", how="left").fillna(0)

print("IRP 1.1:")
print(irp11)

# ============================================================
# ETAPA 7 - IRP 1.2
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 7 - IRP 1.2")
print("=" * 80)

df_bacias_pmcrp = df_total[df_total["Bacia"].isin(BACIAS_PMCRP)].copy()

print("Linhas do recorte de bacias:", len(df_bacias_pmcrp))
print("Bacias encontradas:", sorted(df_bacias_pmcrp["Bacia"].dropna().unique().tolist()))

irp12 = (
    df_bacias_pmcrp.groupby(["Ano", "Bacia"])["Produção Total em Barril"]
    .sum()
    .unstack(fill_value=0)
    .reset_index()
)

for b in BACIAS_PMCRP:
    if b not in irp12.columns:
        irp12[b] = 0.0

irp12 = base_anos.merge(irp12, on="Ano", how="left").fillna(0)
irp12 = irp12.merge(irp11[["Ano", "Produção Total em Barril"]], on="Ano", how="left")

irp12 = irp12.rename(columns={
    "Produção Total em Barril": "Total Geral",
    "Campos": "Produção Bacia de Campos",
    "Santos": "Produção Bacia de Santos",
    "Espírito Santo": "Produção Bacia do Espírito Santo"
})

irp12["Proporção Campos"] = irp12["Produção Bacia de Campos"] / irp12["Total Geral"]
irp12["Proporção Santos"] = irp12["Produção Bacia de Santos"] / irp12["Total Geral"]
irp12["Proporção Espírito Santo"] = irp12["Produção Bacia do Espírito Santo"] / irp12["Total Geral"]

print("IRP 1.2:")
print(irp12)

# ============================================================
# ETAPA 8 - IRP 1.3
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 8 - IRP 1.3")
print("=" * 80)

irp13_todos_estados = (
    df_total.groupby(["Ano", "Estado"], as_index=False)
    .agg({
        "Petróleo em Barril": "sum",
        "Gás Natural em Barril": "sum",
        "Produção Total em Barril": "sum"
    })
    .sort_values(["Ano", "Estado"])
)

irp13_todos_estados = irp13_todos_estados.rename(columns={
    "Produção Total em Barril": "Somatório das Produções"
})

df_estados_pmcrp = irp13_todos_estados[irp13_todos_estados["Estado"].isin(ESTADOS_PMCRP)].copy()

pmcrp_total = (
    df_estados_pmcrp.groupby("Ano", as_index=False)["Somatório das Produções"]
    .sum()
    .rename(columns={"Somatório das Produções": "Somatório PMCRP em Barril"})
)

prop_rj = (
    df_estados_pmcrp[df_estados_pmcrp["Estado"] == "Rio de Janeiro"][["Ano", "Somatório das Produções"]]
    .rename(columns={"Somatório das Produções": "RJ"})
)
prop_sp = (
    df_estados_pmcrp[df_estados_pmcrp["Estado"] == "São Paulo"][["Ano", "Somatório das Produções"]]
    .rename(columns={"Somatório das Produções": "SP"})
)
prop_es = (
    df_estados_pmcrp[df_estados_pmcrp["Estado"] == "Espírito Santo"][["Ano", "Somatório das Produções"]]
    .rename(columns={"Somatório das Produções": "ES"})
)
prop_sc = (
    df_estados_pmcrp[df_estados_pmcrp["Estado"] == "Santa Catarina"][["Ano", "Somatório das Produções"]]
    .rename(columns={"Somatório das Produções": "SC"})
)

irp13_resumo = base_anos.copy()
irp13_resumo = irp13_resumo.merge(pmcrp_total, on="Ano", how="left")
irp13_resumo = irp13_resumo.merge(irp11[["Ano", "Produção Total em Barril"]], on="Ano", how="left")
irp13_resumo = irp13_resumo.merge(prop_rj, on="Ano", how="left")
irp13_resumo = irp13_resumo.merge(prop_sp, on="Ano", how="left")
irp13_resumo = irp13_resumo.merge(prop_es, on="Ano", how="left")
irp13_resumo = irp13_resumo.merge(prop_sc, on="Ano", how="left")
irp13_resumo = irp13_resumo.fillna(0)

irp13_resumo["Produção do Rio de Janeiro pelo Total"] = irp13_resumo["RJ"] / irp13_resumo["Produção Total em Barril"]
irp13_resumo["Produção de São Paulo pelo Total"] = irp13_resumo["SP"] / irp13_resumo["Produção Total em Barril"]
irp13_resumo["Produção do Espírito Santo pelo Total"] = irp13_resumo["ES"] / irp13_resumo["Produção Total em Barril"]
irp13_resumo["Produção de Santa Catarina pelo Total"] = irp13_resumo["SC"] / irp13_resumo["Produção Total em Barril"]
irp13_resumo["Soma PMCRP pelo Total"] = irp13_resumo["Somatório PMCRP em Barril"] / irp13_resumo["Produção Total em Barril"]

print("IRP 1.3 - Estados:")
print(df_estados_pmcrp.head(20))

print("\nIRP 1.3 - Resumo:")
print(irp13_resumo)

# ============================================================
# ETAPA 9 - EXPORTAÇÃO
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 9 - EXPORTAÇÃO")
print("=" * 80)

with pd.ExcelWriter(ARQUIVO_SAIDA, engine="openpyxl") as writer:
    auditoria_ano_ambiente.to_excel(writer, sheet_name="Auditoria_Ano_Ambiente", index=False)
    irp11.to_excel(writer, sheet_name="IRP1.1", index=False)
    irp12.to_excel(writer, sheet_name="IRP1.2", index=False)
    df_estados_pmcrp.to_excel(writer, sheet_name="IRP1.3_Estados", index=False)
    irp13_resumo.to_excel(writer, sheet_name="IRP1.3_Resumo", index=False)

    for aba in [
        "Auditoria_Ano_Ambiente",
        "IRP1.1",
        "IRP1.2",
        "IRP1.3_Estados",
        "IRP1.3_Resumo"
    ]:
        formatar_aba(writer, aba)
        print(f" -> formatada: {aba}")

print("\nArquivo gerado com sucesso:")
print(ARQUIVO_SAIDA)