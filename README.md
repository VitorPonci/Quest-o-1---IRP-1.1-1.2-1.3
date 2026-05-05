# Indicadores IRP 1 — Produção de Petróleo e Gás Natural

Este repositório contém um script em Python para consolidar bases de produção de petróleo e gás natural e calcular os indicadores **IRP 1.1, IRP 1.2 e IRP 1.3**, utilizados na análise da produção nacional e da participação das bacias e estados abrangidos pelo PMCRP.

O código lê arquivos Excel de uma pasta local, trata inconsistências de nomes de colunas e formatos numéricos, aplica recortes por ano, ambiente, bacia e estado, calcula volumes convertidos para barris e exporta os resultados consolidados para uma planilha Excel final.

---

## Objetivo

O objetivo do script é automatizar o cálculo dos indicadores de produção de petróleo e gás natural a partir de bases anuais/mensais da ANP ou de bases previamente organizadas em Excel.

A rotina permite:

- Consolidar múltiplos arquivos Excel em uma única base;
- Padronizar nomes de colunas com problemas de codificação;
- Tratar números em formato brasileiro;
- Separar e somar produção de óleo, condensado e gás natural;
- Converter volumes para barris;
- Calcular indicadores agregados por ano, bacia e estado;
- Exportar os resultados em um arquivo Excel formatado.

---

## Indicadores calculados

### IRP 1.1 — Produção total anual

Calcula, por ano, a produção total de:

- Óleo;
- Condensado;
- Óleo + condensado;
- Petróleo em barris;
- Gás associado;
- Gás não associado;
- Gás total;
- Gás natural convertido em barris;
- Produção total em barris.

---

### IRP 1.2 — Participação das bacias selecionadas

Calcula a produção total em barris das principais bacias do recorte PMCRP:

- Bacia de Campos;
- Bacia de Santos;
- Bacia do Espírito Santo.

Também calcula a proporção de cada bacia em relação à produção total geral.

---

### IRP 1.3 — Participação dos estados do PMCRP

Calcula a produção total em barris dos estados abrangidos pelo programa:

- Espírito Santo;
- Rio de Janeiro;
- São Paulo;
- Santa Catarina.

Também calcula a participação de cada estado e do conjunto PMCRP em relação ao total geral.

---

## Estrutura esperada dos dados

O script espera arquivos Excel (`.xlsx` ou `.xls`) contendo, preferencialmente, as seguintes colunas:

| Coluna | Descrição |
|---|---|
| Ano | Ano de referência da produção |
| Mês/Ano | Mês e ano de referência, quando disponível |
| Estado | Unidade da federação |
| Bacia | Bacia sedimentar |
| Campo | Nome do campo |
| Poço | Identificação do poço |
| Ambiente | Mar ou Terra |
| Instalação | Instalação associada à produção |
| Produção de Óleo (m³) | Volume de óleo produzido |
| Produção de Condensado (m³) | Volume de condensado produzido |
| Produção de Gás Associado (Mm³) | Volume de gás associado |
| Produção de Gás Não Associado (Mm³) | Volume de gás não associado |

O código também trata automaticamente algumas variações de nomes com problemas de acentuação ou codificação.

---

## Conversões utilizadas

O script utiliza o seguinte fator de conversão:

```python
FATOR_BARRIL = 6.289941

Petróleo em Barril = (Produção de Óleo + Produção de Condensado) × FATOR_BARRIL

Gás Total (Mm³) = Gás Associado + Gás Não Associado
Gás Equivalente (m³ petróleo) = Gás Total (Mm³) / 1000
Gás Natural em Barril = Gás Equivalente × FATOR_BARRIL

Produção Total em Barril = Petróleo em Barril + Gás Natural em Barril

ANOS_MAR_TERRA = list(range(2010, 2027))

