import warnings
warnings.filterwarnings("ignore")

import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# CONFIGURAÇÃO
# ============================================================

PASTA_DADOS = Path(
    r"C:\Users\vitorL\OneDrive - Fundação Instituto de Administração\Área de Trabalho\Produção de O&G\Questão 1"
)

ARQUIVO_SAIDA = PASTA_DADOS / "Indicadores_IRP1_nova_metodologia.xlsx"

# Recortes espaciais da Questão 1
BACIAS_PMCRP = ["Campos", "Santos", "Espírito Santo"]

# Nova nota metodológica: SC, PR, SP, RJ e ES
ESTADOS_PMCRP = [
    "Santa Catarina",
    "São Paulo",
    "Rio de Janeiro",
    "Espírito Santo",
]

# Série histórica
ANO_INICIAL = 2010
ANO_FINAL = 2026
ANOS_FINAIS = list(range(ANO_INICIAL, ANO_FINAL + 1))

# Se, futuramente, o IRP1.2 precisar considerar apenas produção marítima,
# alterar para True. Pela nota atual, o padrão mantém Mar + Terra.
IRP12_APENAS_MAR = False


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def remover_acentos(texto):
    """Remove acentos para facilitar padronizações e comparações."""
    if pd.isna(texto):
        return ""
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return texto


def chave_texto(texto):
    """Cria uma chave normalizada para comparação textual."""
    return remover_acentos(texto).strip().lower()


def limpar_nome_coluna(col):
    """Padroniza nomes de colunas esperados nas bases da ANP."""
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
    """Converte valores numéricos com vírgula decimal e separadores brasileiros."""
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip()

    if s in ["", "nan", "None", "-", " "]:
        return 0.0

    # Ex.: 1.234.567,89 -> 1234567.89
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    # Ex.: 123,45 -> 123.45
    elif "," in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except Exception:
        return 0.0


def normalizar_ambiente(valor):
    s = chave_texto(valor)

    if s in ["mar", "maritimo", "offshore"]:
        return "Mar"
    if s in ["terra", "terrestre", "onshore"]:
        return "Terra"

    if "mar" in s:
        return "Mar"
    if "terra" in s:
        return "Terra"

    return str(valor).strip()


def normalizar_estado(valor):
    s = chave_texto(valor)

    mapa = {
        "espirito santo": "Espírito Santo",
        "rio de janeiro": "Rio de Janeiro",
        "sao paulo": "São Paulo",
        "santa catarina": "Santa Catarina",
        "parana": "Paraná",
    }

    return mapa.get(s, str(valor).strip())


def normalizar_bacia(valor):
    s = chave_texto(valor)

    mapa = {
        "campos": "Campos",
        "santos": "Santos",
        "espirito santo": "Espírito Santo",
    }

    return mapa.get(s, str(valor).strip())


def extrair_ano(df):
    """Garante que a coluna Ano exista, usando Mês/Ano quando necessário."""
    if "Ano" not in df.columns and "Mês/Ano" in df.columns:
        df["Ano"] = df["Mês/Ano"].astype(str).str[-4:]

    return df


def ler_excel(arquivo):
    """Lê todas as abas de um arquivo Excel e concatena as que possuem Ano."""
    try:
        xls = pd.ExcelFile(arquivo)
        dfs = []

        for aba in xls.sheet_names:
            try:
                df = pd.read_excel(arquivo, sheet_name=aba)
                df = normalizar_colunas(df)
                df = extrair_ano(df)

                if "Ano" in df.columns:
                    df["Arquivo_Origem"] = arquivo.name
                    df["Aba_Origem"] = aba
                    dfs.append(df)

            except Exception as e:
                print(f"   -> Aba ignorada: {aba} | Motivo: {e}")
                continue

        if dfs:
            return pd.concat(dfs, ignore_index=True)

        return None

    except Exception as e:
        print(f"Erro ao ler {arquivo.name}: {e}")
        return None


def dividir_seguro(numerador, denominador):
    """Evita divisão por zero."""
    return numerador.div(denominador.replace({0: pd.NA})).fillna(0)


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
        ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cab = str(ws.cell(1, cell.column).value)
            if isinstance(cell.value, (int, float)):
                if "Proporção" in cab or "pelo Total" in cab:
                    cell.number_format = "0.0000000000%"
                else:
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

# Evita reler o próprio arquivo de saída, caso ele já exista na pasta.
arquivos = [a for a in arquivos if a.name != ARQUIVO_SAIDA.name]

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
    "Produção de Gás Não Associado (Mm³)",
]

for col in colunas_necessarias:
    if col not in df_total.columns:
        df_total[col] = "" if col in ["Estado", "Bacia", "Ambiente"] else 0

colunas_producao = [
    "Produção de Óleo (m³)",
    "Produção de Condensado (m³)",
    "Produção de Gás Associado (Mm³)",
    "Produção de Gás Não Associado (Mm³)",
]

for col in colunas_producao:
    print(f"Tratando coluna: {col}")
    df_total[col] = df_total[col].apply(parse_numero)

df_total["Ano"] = pd.to_numeric(df_total["Ano"], errors="coerce")
df_total["Estado"] = df_total["Estado"].apply(normalizar_estado)
df_total["Bacia"] = df_total["Bacia"].apply(normalizar_bacia)
df_total["Ambiente"] = df_total["Ambiente"].apply(normalizar_ambiente)

# Variáveis centrais da nova metodologia
df_total["Petróleo Total (m³)"] = (
    df_total["Produção de Óleo (m³)"] +
    df_total["Produção de Condensado (m³)"]
)

df_total["Gás Total (Mm³)"] = (
    df_total["Produção de Gás Associado (Mm³)"] +
    df_total["Produção de Gás Não Associado (Mm³)"]
)

print("\nValores únicos de Ambiente antes do recorte:")
print(sorted(df_total["Ambiente"].dropna().unique().tolist())[:30])


# ============================================================
# ETAPA 3 - APLICANDO RECORTE TEMPORAL E DE AMBIENTE
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 3 - APLICANDO RECORTE")
print("=" * 80)

filtro_mar_terra = (
    df_total["Ano"].isin(ANOS_FINAIS)
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
        "Petróleo Total (m³)": "sum",
        "Produção de Gás Associado (Mm³)": "sum",
        "Produção de Gás Não Associado (Mm³)": "sum",
        "Gás Total (Mm³)": "sum",
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

colunas_amostra = [
    "Ano",
    "Estado",
    "Bacia",
    "Ambiente",
    "Produção de Óleo (m³)",
    "Produção de Condensado (m³)",
    "Petróleo Total (m³)",
    "Produção de Gás Associado (Mm³)",
    "Produção de Gás Não Associado (Mm³)",
    "Gás Total (Mm³)",
]

print(df_total[colunas_amostra].head(20))


# ============================================================
# ETAPA 6 - IRP 1.1
# Produção nacional total em unidades originais
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
        "Petróleo Total (m³)": "sum",
        "Produção de Gás Associado (Mm³)": "sum",
        "Produção de Gás Não Associado (Mm³)": "sum",
        "Gás Total (Mm³)": "sum",
    })
    .sort_values("Ano")
)

irp11 = base_anos.merge(irp11, on="Ano", how="left").fillna(0)

print("IRP 1.1:")
print(irp11)


# ============================================================
# ETAPA 7 - IRP 1.2
# Proporção das bacias selecionadas, separada por petróleo e gás
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 7 - IRP 1.2")
print("=" * 80)

df_bacias_pmcrp = df_total[df_total["Bacia"].isin(BACIAS_PMCRP)].copy()

if IRP12_APENAS_MAR:
    df_bacias_pmcrp = df_bacias_pmcrp[df_bacias_pmcrp["Ambiente"] == "Mar"].copy()

print("Linhas do recorte de bacias:", len(df_bacias_pmcrp))
print("Bacias encontradas:", sorted(df_bacias_pmcrp["Bacia"].dropna().unique().tolist()))

irp12_bacias = (
    df_bacias_pmcrp.groupby(["Ano", "Bacia"], as_index=False)
    .agg({
        "Petróleo Total (m³)": "sum",
        "Gás Total (Mm³)": "sum",
    })
    .sort_values(["Ano", "Bacia"])
)

irp12_bacias = irp12_bacias.merge(
    irp11[["Ano", "Petróleo Total (m³)", "Gás Total (Mm³)"]].rename(columns={
        "Petróleo Total (m³)": "IRP1.1 Petróleo Nacional (m³)",
        "Gás Total (Mm³)": "IRP1.1 Gás Nacional (Mm³)",
    }),
    on="Ano",
    how="left",
)

irp12_bacias["Proporção Petróleo pelo Total Nacional"] = dividir_seguro(
    irp12_bacias["Petróleo Total (m³)"],
    irp12_bacias["IRP1.1 Petróleo Nacional (m³)"],
)

irp12_bacias["Proporção Gás pelo Total Nacional"] = dividir_seguro(
    irp12_bacias["Gás Total (Mm³)"],
    irp12_bacias["IRP1.1 Gás Nacional (Mm³)"],
)

irp12_resumo = (
    df_bacias_pmcrp.groupby("Ano", as_index=False)
    .agg({
        "Petróleo Total (m³)": "sum",
        "Gás Total (Mm³)": "sum",
    })
    .rename(columns={
        "Petróleo Total (m³)": "Petróleo Bacias PMCRP (m³)",
        "Gás Total (Mm³)": "Gás Bacias PMCRP (Mm³)",
    })
)

irp12_resumo = base_anos.merge(irp12_resumo, on="Ano", how="left").fillna(0)

irp12_resumo = irp12_resumo.merge(
    irp11[["Ano", "Petróleo Total (m³)", "Gás Total (Mm³)"]].rename(columns={
        "Petróleo Total (m³)": "IRP1.1 Petróleo Nacional (m³)",
        "Gás Total (Mm³)": "IRP1.1 Gás Nacional (Mm³)",
    }),
    on="Ano",
    how="left",
)

irp12_resumo["IRP1.2 Petróleo - Bacias pelo Total Nacional"] = dividir_seguro(
    irp12_resumo["Petróleo Bacias PMCRP (m³)"],
    irp12_resumo["IRP1.1 Petróleo Nacional (m³)"],
)

irp12_resumo["IRP1.2 Gás - Bacias pelo Total Nacional"] = dividir_seguro(
    irp12_resumo["Gás Bacias PMCRP (Mm³)"],
    irp12_resumo["IRP1.1 Gás Nacional (Mm³)"],
)

print("IRP 1.2 - Bacias:")
print(irp12_bacias.head(20))

print("\nIRP 1.2 - Resumo:")
print(irp12_resumo)


# ============================================================
# ETAPA 8 - IRP 1.3
# Proporção dos estados selecionados, separada por petróleo e gás
# ============================================================

print("\n" + "=" * 80)
print("ETAPA 8 - IRP 1.3")
print("=" * 80)

irp13_todos_estados = (
    df_total.groupby(["Ano", "Estado"], as_index=False)
    .agg({
        "Petróleo Total (m³)": "sum",
        "Gás Total (Mm³)": "sum",
    })
    .sort_values(["Ano", "Estado"])
)

irp13_todos_estados = irp13_todos_estados.merge(
    irp11[["Ano", "Petróleo Total (m³)", "Gás Total (Mm³)"]].rename(columns={
        "Petróleo Total (m³)": "IRP1.1 Petróleo Nacional (m³)",
        "Gás Total (Mm³)": "IRP1.1 Gás Nacional (Mm³)",
    }),
    on="Ano",
    how="left",
)

irp13_todos_estados["Proporção Petróleo pelo Total Nacional"] = dividir_seguro(
    irp13_todos_estados["Petróleo Total (m³)"],
    irp13_todos_estados["IRP1.1 Petróleo Nacional (m³)"],
)

irp13_todos_estados["Proporção Gás pelo Total Nacional"] = dividir_seguro(
    irp13_todos_estados["Gás Total (Mm³)"],
    irp13_todos_estados["IRP1.1 Gás Nacional (Mm³)"],
)

df_estados_pmcrp = irp13_todos_estados[
    irp13_todos_estados["Estado"].isin(ESTADOS_PMCRP)
].copy()

irp13_resumo = (
    df_estados_pmcrp.groupby("Ano", as_index=False)
    .agg({
        "Petróleo Total (m³)": "sum",
        "Gás Total (Mm³)": "sum",
    })
    .rename(columns={
        "Petróleo Total (m³)": "Petróleo Estados PMCRP (m³)",
        "Gás Total (Mm³)": "Gás Estados PMCRP (Mm³)",
    })
)

irp13_resumo = base_anos.merge(irp13_resumo, on="Ano", how="left").fillna(0)

irp13_resumo = irp13_resumo.merge(
    irp11[["Ano", "Petróleo Total (m³)", "Gás Total (Mm³)"]].rename(columns={
        "Petróleo Total (m³)": "IRP1.1 Petróleo Nacional (m³)",
        "Gás Total (Mm³)": "IRP1.1 Gás Nacional (Mm³)",
    }),
    on="Ano",
    how="left",
)

irp13_resumo["IRP1.3 Petróleo - Estados pelo Total Nacional"] = dividir_seguro(
    irp13_resumo["Petróleo Estados PMCRP (m³)"],
    irp13_resumo["IRP1.1 Petróleo Nacional (m³)"],
)

irp13_resumo["IRP1.3 Gás - Estados pelo Total Nacional"] = dividir_seguro(
    irp13_resumo["Gás Estados PMCRP (Mm³)"],
    irp13_resumo["IRP1.1 Gás Nacional (Mm³)"],
)

print("IRP 1.3 - Estados PMCRP:")
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
    auditoria_ano_ambiente.to_excel(
        writer,
        sheet_name="Auditoria_Ano_Ambiente",
        index=False,
    )

    irp11.to_excel(
        writer,
        sheet_name="IRP1.1",
        index=False,
    )

    irp12_bacias.to_excel(
        writer,
        sheet_name="IRP1.2_Bacias",
        index=False,
    )

    irp12_resumo.to_excel(
        writer,
        sheet_name="IRP1.2_Resumo",
        index=False,
    )

    df_estados_pmcrp.to_excel(
        writer,
        sheet_name="IRP1.3_Estados",
        index=False,
    )

    irp13_resumo.to_excel(
        writer,
        sheet_name="IRP1.3_Resumo",
        index=False,
    )

    # Opcional: amostra da base tratada para auditoria
    df_total[colunas_amostra + ["Arquivo_Origem", "Aba_Origem"]].head(5000).to_excel(
        writer,
        sheet_name="Amostra_Base_Tratada",
        index=False,
    )

    for aba in [
        "Auditoria_Ano_Ambiente",
        "IRP1.1",
        "IRP1.2_Bacias",
        "IRP1.2_Resumo",
        "IRP1.3_Estados",
        "IRP1.3_Resumo",
        "Amostra_Base_Tratada",
    ]:
        formatar_aba(writer, aba)
        print(f" -> formatada: {aba}")

print("\nArquivo gerado com sucesso:")
print(ARQUIVO_SAIDA)
