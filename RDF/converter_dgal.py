#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Conversor reproduzível DGAL (2017-2019) -> RDF/Turtle.

Grupo 7 | UC Web Semântica | tarefa 6.2: observações e catálogo DCAT-AP.

Entradas: 12 CSV/XLSX da DGAL na pasta dados/,
          mapeamento-vocabulario-g7.csv, fontes_dgal.csv e
          Ligacoes-Wikidata.csv na raiz da pasta RDF/.
Saídas:   catálogo RDF (derivados), SKOS territorial e períodos,
          DSDs, 12 ficheiros com observações e manifesto de conversão.

Metadados DCAT-AP do perfil do Grupo 7: catálogo, quatro séries, 12 datasets,
12 distribuições, temas, cobertura temporal/espacial, proveniência e licença.
Inclui 308 municípios ligados aos 18 distritos e 2 regiões autónomas por
skos:broader e as correspondências Wikidata fornecidas pelo grupo.
As URLs de origem são fornecidas em fontes_dgal.csv. Os URLs de acesso ao RDF
apontam para os ficheiros públicos no GitHub, sem caminhos locais file://.
Os nomes de domínio em w3id.org são propostos e ainda não estão registados.

Publicação GitHub: por omissão os 12 ficheiros RDF são descritos como
          https://raw.githubusercontent.com/Mota005/WB/<RAMO>/RDF/rdf/<ficheiro>.ttl
          O ramo é detetado no Git local, se possível; na ausência de informação,
          usa-se «main» e a escolha é assinalada no terminal.
Opcional: --github-ramo RAMO  (se o ramo público for diferente)
Opcional: --url-distribuicoes URL  (para publicar noutra plataforma)
Opcional: --fontes fontes_dgal.csv (utilizado automaticamente se estiver
          ao lado do script). Tem campos serie;ano;url_dataset_original;
          url_download_original e apenas URLs oficialmente confirmadas.

IMPORTANTE: executar «git push» depois de atualizar os ficheiros no repositório.
As URLs do GitHub são construídas a partir do endereço e do ramo configurados;
a execução local não confirma, por si só, que os ficheiros estão online.

Requisitos: Python >=3.10; pip install rdflib
Limites: SHACL e enriquecimento estatístico INE são etapas separadas.
Execução:   python converter_dgal.py
            python converter_dgal.py --dados dados --saida rdf
"""

from __future__ import annotations

import argparse
import csv
import subprocess
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from urllib.parse import quote, urlparse
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from datetime import date
from pathlib import Path

from rdflib import Graph, Literal, Namespace, URIRef
from rdflib.namespace import DCTERMS as DCT, FOAF, RDF, RDFS, SKOS, XSD

DEFAULT_BASE = "https://w3id.org/ws-megsi-g7/"  # PROPOSTA (a registar)
GITHUB_REPO_PADRAO = "Mota005/WB"  # repositório público indicado pelo utilizador
GITHUB_SUBPASTA = "RDF/rdf"
QB = Namespace("http://purl.org/linked-data/cube#")
DCAT = Namespace("http://www.w3.org/ns/dcat#")
SDMX_DIM = Namespace("http://purl.org/linked-data/sdmx/2009/dimension#")
SDMX_ATT = Namespace("http://purl.org/linked-data/sdmx/2009/attribute#")
EU_FILE = Namespace("http://publications.europa.eu/resource/authority/file-type/")
EU_LICENSE = Namespace("http://publications.europa.eu/resource/authority/licence/")
EU_THEME = Namespace("http://publications.europa.eu/resource/authority/data-theme/")
EU_COUNTRY = Namespace("http://publications.europa.eu/resource/authority/country/")
EU_LANG = Namespace("http://publications.europa.eu/resource/authority/language/")
EU_FREQUENCY = Namespace("http://publications.europa.eu/resource/authority/frequency/")
EU_ACCESS = Namespace("http://publications.europa.eu/resource/authority/access-right/")
EU_DISTR_STATUS = Namespace("http://publications.europa.eu/resource/authority/distribution-status/")
PROV = Namespace("http://www.w3.org/ns/prov#")
ADMS = Namespace("http://www.w3.org/ns/adms#")
VCARD = Namespace("http://www.w3.org/2006/vcard/ns#")
LICENCA_CC_BY = EU_LICENSE.CC_BY_4_0
TIPO_PUBLICADOR = Namespace("http://purl.org/adms/publishertype/")
# Primeira publicação RDF no GitHub (confirmada no projeto em 22/11/2026).
DATA_PRIMEIRA_PUBLICACAO = date(2026, 11, 22)

# Descrições e palavras-chave criadas para o perfil do Grupo 7.
TEMAS = (EU_THEME.ECON, EU_THEME.GOVE)
METADADOS_SERIES = {
    "Receitas": {
        "descricao": "Receitas municipais de 308 municípios portugueses, incluindo impostos, transferências, financiamento e outras receitas; os trimestres apresentam valores acumulados.",
        "descricao_en": "Municipal revenues for Portugal's 308 municipalities, including taxes, transfers, funding and other revenue; quarterly values are cumulative.",
        "keywords": ("receitas municipais", "impostos municipais", "transferências", "finanças locais"),
        "keywords_en": ("municipal revenue", "municipal taxes", "transfers", "local government finance"),
    },
    "Despesas": {
        "descricao": "Despesas municipais de 308 municípios portugueses, incluindo despesas correntes, investimento, pessoal e transferências; os trimestres apresentam valores acumulados.",
        "descricao_en": "Municipal expenditure for Portugal's 308 municipalities, including current expenses, investment, staff and transfers; quarterly values are cumulative.",
        "keywords": ("despesas municipais", "investimento", "despesas com pessoal", "finanças locais"),
        "keywords_en": ("municipal expenditure", "investment", "personnel costs", "local government finance"),
    },
    "Recursos Humanos": {
        "descricao": "Indicadores de recursos humanos dos 308 municípios portugueses, incluindo trabalhadores, dirigentes, categorias e tipos de vínculo; 2017 contém apenas o período anual.",
        "descricao_en": "Human resources indicators for Portugal's 308 municipalities, including staff, management categories and employment contracts; 2017 includes annual observations only.",
        "keywords": ("recursos humanos", "trabalhadores municipais", "vínculos laborais", "dirigentes"),
        "keywords_en": ("human resources", "municipal employees", "employment contracts", "managers"),
    },
    "Qualidade da Gestão": {
        "descricao": "Indicadores de gestão pública e sustentabilidade financeira dos 308 municípios portugueses, incluindo dívida, execução orçamental e prazos de pagamento.",
        "descricao_en": "Public management and financial sustainability indicators for Portugal's 308 municipalities, including debt, budget execution and payment times.",
        "keywords": ("gestão pública", "endividamento", "execução orçamental", "sustentabilidade financeira"),
        "keywords_en": ("public management", "indebtedness", "budget execution", "financial sustainability"),
    },
}

# [CORREÇÃO 4] Títulos em inglês, unidades e descrições por conjunto de dados
# (a descrição da série mencionava 2017 em todos os anos de Recursos Humanos).
TITULOS_EN = {"Receitas": "Revenue", "Despesas": "Expenditure",
              "Recursos Humanos": "Human Resources", "Qualidade da Gestão": "Management Quality"}
DESCRICAO_DATASET = {
    "Receitas": ("Receitas municipais dos 308 municípios portugueses: receitas totais, transferências do Orçamento do Estado, "
                 "financiamento da UE, impostos municipais (IMI, IMT, IUC, derrama, taxas e multas), venda de bens e serviços "
                 "e outras receitas. Valores em euros; os trimestres são acumulados desde o início do ano.",
                 "Municipal revenue of Portugal's 308 municipalities: total revenue, State Budget transfers, EU funding, "
                 "municipal taxes (IMI, IMT, IUC, corporate surcharge, fees and fines), sales of goods and services and other "
                 "revenue. Values in euros; quarters are cumulative from the start of the year."),
    "Despesas": ("Despesas municipais dos 308 municípios portugueses: despesa total, despesa corrente (pessoal, aquisição de bens, "
                 "juros, transferências correntes), investimento (terrenos e habitações, edifícios, infraestruturas básicas, "
                 "acessibilidades, outros) e transferências de capital. Valores em euros; os trimestres são acumulados.",
                 "Municipal expenditure of Portugal's 308 municipalities: total expenditure, current expenditure (staff, goods, "
                 "interest, current transfers), investment (land and housing, buildings, basic infrastructure, accessibility, "
                 "other) and capital transfers. Values in euros; quarters are cumulative."),
    "Recursos Humanos": ("Recursos humanos dos 308 municípios portugueses: total de trabalhadores, por tipologia (dirigentes, "
                         "técnicos superiores, outros) e por vínculo (comissão de serviço, com e sem vínculo permanente). "
                         "Número de pessoas.",
                         "Human resources of Portugal's 308 municipalities: total staff, by category (managers, senior "
                         "technical staff, other) and by employment link (service commission, permanent and non-permanent). "
                         "Number of people."),
    "Qualidade da Gestão": ("Indicadores de gestão e sustentabilidade financeira dos 308 municípios portugueses: dívida total, "
                            "receita e despesa totais, receita efetiva, saldo orçamental e pagamentos em atraso (euros); graus "
                            "de endividamento, de execução da receita e da despesa e de independência financeira (%); prazo "
                            "médio de pagamento (dias).",
                            "Management and financial sustainability indicators of Portugal's 308 municipalities: total debt, "
                            "total revenue and expenditure, effective revenue, budget balance and overdue payments (euros); "
                            "debt, revenue and expenditure execution and financial independence ratios (%); average payment "
                            "period (days)."),
}
NOTA_ANO = {
    ("Receitas", 2017): ("Sem valores anuais (A): o valor de fim de ano é o de T4.",
                         "No annual (A) values: the year-end value is T4."),
    ("Despesas", 2017): ("Sem valores anuais (A): o valor de fim de ano é o de T4.",
                         "No annual (A) values: the year-end value is T4."),
    ("Recursos Humanos", 2017): ("Só período anual; inclui também representantes eleitos e cargos políticos.",
                                 "Annual period only; also includes elected representatives and political positions."),
}

# [CORREÇÃO 1 do catálogo] Conjuntos de dados ORIGINAIS da DGAL, descritos no catálogo
# (Secção 3.3 / Tabela 6 do relatório). Metadados recolhidos pelo grupo no portal
# (Tabela 3 do relatório): títulos com acentos corrigidos, datas de criação/modificação
# e palavras-chave corrigidas («tranferencias» corrigido, «migrado» retirado).
# Datas não indicadas no relatório ficam em falta (None) em vez de inventadas.
TITULO_ORIGINAL = {"Receitas": "Receitas Municipais {ano}", "Despesas": "Despesas Municipais {ano}",
                   "Recursos Humanos": "Recursos Humanos {ano}",
                   "Qualidade da Gestão": "Qualidade da Gestão Pública {ano}"}
TITULO_ORIGINAL_EN = {"Receitas": "Municipal Revenue {ano}", "Despesas": "Municipal Expenditure {ano}",
                      "Recursos Humanos": "Municipal Human Resources {ano}",
                      "Qualidade da Gestão": "Public Management Quality {ano}"}
DATAS_ORIGINAL = {  # (dct:issued, dct:modified) no portal
    ("Receitas", 2017): (date(2018, 5, 4), date(2018, 5, 14)),
    ("Despesas", 2017): (date(2018, 5, 4), date(2018, 5, 14)),
    ("Recursos Humanos", 2017): (date(2018, 5, 4), date(2018, 5, 14)),
    ("Receitas", 2018): (date(2020, 12, 28), None), ("Receitas", 2019): (date(2020, 12, 28), None),
    ("Despesas", 2018): (date(2020, 12, 28), None), ("Despesas", 2019): (date(2020, 12, 28), None),
    ("Qualidade da Gestão", 2017): (date(2021, 10, 29), None),
    ("Qualidade da Gestão", 2018): (date(2021, 10, 29), None),
    ("Qualidade da Gestão", 2019): (date(2021, 10, 29), None),
    ("Recursos Humanos", 2018): (date(2021, 10, 29), date(2021, 12, 13)),
    ("Recursos Humanos", 2019): (date(2021, 10, 29), date(2021, 12, 13)),
}
KEYWORDS_ORIGINAL = {
    "Receitas": (("derrama", "financiamento", "IMI", "IMT", "IRC", "IUC", "receita", "transferências"),
                 ("municipal surcharge", "funding", "property tax", "property transfer tax", "corporate income tax",
                  "vehicle tax", "revenue", "transfers")),
    "Despesas": (("despesa", "encargos", "investimentos", "juros"),
                 ("expenditure", "charges", "investment", "interest")),
    "Qualidade da Gestão": (("assistência", "dívida", "endividamento", "execução", "independência financeira",
                             "pagamentos em atraso"),
                            ("financial assistance", "debt", "indebtedness", "execution", "financial independence",
                             "overdue payments")),
    "Recursos Humanos": (("cargos", "comissão de serviço", "dirigentes", "eleitos", "funcionários", "técnicos", "vínculo"),
                         ("positions", "service commission", "managers", "elected representatives", "employees",
                          "technical staff", "employment link")),
}
IANA = "http://www.iana.org/assignments/media-types/"
MEDIA_TYPE = {"csv": URIRef(IANA + "text/csv"),
              "xlsx": URIRef(IANA + "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
              "ttl": URIRef(IANA + "text/turtle")}
# Contacto oficial da DGAL: preencher com um URL ou e-mail confirmado (ex.: "mailto:...").
# Enquanto for None, os conjuntos de dados originais não têm dcat:contactPoint.
CONTACTO_DGAL = None
PAGINA_PROJETO = "https://github.com/Mota005/WB"

# [CORREÇÃO 5] Nomes de municípios: homónimos desambiguados com o nome usado pela DGAL
# («Lagoa (Algarve)», «Calheta (Madeira)», ...) e nome oficial quando o CSV usa outro.
# O nome do CSV de ligações fica em skos:altLabel.
NOMES_OFICIAIS = {"4302": "Praia da Vitória"}


# Ficheiros originais identificados no projeto. 2017: latin-1;
# 2018-2019: CP850 (testado nos nomes municipais com acentos).
INPUTS = [
    ("Receitas", 2017, "ReceitasMunicipais-2017.csv", "iso-8859-1"),
    ("Receitas", 2018, "ReceitasMunicipais-2018.csv", "cp850"),
    ("Receitas", 2019, "ReceitasMunicipais-2019.csv", "cp850"),
    ("Despesas", 2017, "DespesasMunicipais-2017.csv", "iso-8859-1"),
    ("Despesas", 2018, "DespesasMunicipais-2018.csv", "cp850"),
    ("Despesas", 2019, "DespesasMunicipais-2019.csv", "cp850"),
    ("Recursos Humanos", 2017, "RecursosHumanos-2017.xlsx", None),
    ("Recursos Humanos", 2018, "RecursosHumanos-2018.xlsx", None),
    ("Recursos Humanos", 2019, "RecursosHumanos-2019.xlsx", None),
    ("Qualidade da Gestão", 2017, "QualidadeGestaoPublica-2017.xlsx", None),
    ("Qualidade da Gestão", 2018, "QualidadeGestaoPublica-2018.xlsx", None),
    ("Qualidade da Gestão", 2019, "QualidadeGestaoPublica-2019.xlsx", None),
]
SERIES = {
    "Receitas": "receitas-municipais",
    "Despesas": "despesas-municipais",
    "Recursos Humanos": "recursos-humanos",
    "Qualidade da Gestão": "qualidade-gestao",
}
PERIODOS_VALIDOS = {"A", "T1", "T2", "T3", "T4"}

# [CORREÇÃO 3] Definição de cada período (skos:definition), para que o grafo
# diga explicitamente que os trimestres são acumulados nas séries de fluxos.
FIM_TRIMESTRE = {"T1": "31 de março", "T2": "30 de junho", "T3": "30 de setembro", "T4": "31 de dezembro"}


def definicao_periodo(ano: int, periodo: str) -> str:
    if periodo == "A":
        return (f"Valor anual de {ano} publicado pela DGAL. Em 2017 não existe nas séries "
                "Receitas e Despesas; nesses casos o valor de fim de ano é o de T4.")
    return (f"Período de 1 de janeiro a {FIM_TRIMESTRE[periodo]} de {ano}. Nas receitas e despesas "
            "o valor é acumulado desde o início do ano; na dívida, nos rácios e nos recursos humanos "
            "refere-se à situação no fim do período.")


# [CORREÇÃO 1] Em Recursos Humanos 2017 a coluna totalrh inclui os representantes
# eleitos (totalrh = dirigentes + técnicos superiores + outros + eleitos);
# em 2018 e 2019 não inclui. Para g7:trabalhadores ter o mesmo significado nos
# três anos, em 2017 o valor original passa a g7:trabalhadoresInclEleitos e
# g7:trabalhadores é calculado sem os eleitos. A identidade é verificada linha a linha.
def ajustar_rh_2017(meas: list, g7: Namespace, contexto: str) -> list:
    valores = dict(meas)
    total = valores.get(g7.trabalhadores)
    eleitos = valores.get(g7.representantesEleitos)
    partes = [valores.get(g7.dirigentes), valores.get(g7.tecnicosSuperiores), valores.get(g7.outrosTrabalhadores)]
    if total is None or eleitos is None or any(p is None for p in partes):
        raise ValueError(f"{contexto}: RH 2017 sem os campos necessários para separar os eleitos")
    soma_partes = sum(int(str(p)) for p in partes)
    if int(str(total)) != soma_partes + int(str(eleitos)):
        raise ValueError(f"{contexto}: totalrh ({total}) ≠ dirigentes + técnicos + outros + eleitos "
                         f"({soma_partes} + {eleitos}); rever a regra de RH 2017")
    valores[g7.trabalhadoresInclEleitos] = total
    valores[g7.trabalhadores] = Literal(soma_partes, datatype=XSD.integer)
    return list(valores.items())


def verificacoes_coerencia(valores: dict, g7: Namespace, destino: Path) -> list[dict]:
    """[VALIDAÇÃO] Controlos de coerência entre períodos e entre séries.

    Não altera dados: regista o resultado em verificacoes_coerencia.csv para o relatório.
    valores[(serie, ano, periodo, codigo)] = {propriedade: Decimal}
    """
    linhas = []

    def registar(controlo, serie, ano, ok, total, nota):
        linhas.append({"controlo": controlo, "serie": serie, "ano": ano,
                       "casos_ok": ok, "casos_verificados": total, "nota": nota})

    municipios = sorted({k[3] for k in valores})
    # 1) Valor anual (A) vs 4.º trimestre (T4) nas séries de fluxos.
    for serie, prop in (("Receitas", g7.receitasTotais), ("Despesas", g7.despesaTotal)):
        for ano in (2017, 2018, 2019):
            pares = [(valores.get((serie, ano, "A", m), {}).get(prop), valores.get((serie, ano, "T4", m), {}).get(prop))
                     for m in municipios]
            pares = [(a, t) for a, t in pares if t is not None]
            com_a = [(a, t) for a, t in pares if a is not None]
            if not com_a:
                registar("A igual a T4", serie, ano, 0, 0,
                         "Sem período A no ficheiro; usar T4 como valor de fim de ano")
            else:
                registar("A igual a T4", serie, ano, sum(a == t for a, t in com_a), len(com_a),
                         "Onde A ≠ T4, as consultas devem usar sempre o mesmo período em todos os anos")
    # 2) Trimestres acumulados: T1 ≤ T2 ≤ T3 ≤ T4.
    for serie, prop in (("Receitas", g7.receitasTotais), ("Despesas", g7.despesaTotal)):
        for ano in (2017, 2018, 2019):
            ok = tot = 0
            for m in municipios:
                seq = [valores.get((serie, ano, p, m), {}).get(prop) for p in ("T1", "T2", "T3", "T4")]
                if None in seq:
                    continue
                tot += 1
                ok += seq == sorted(seq)
            registar("T1 ≤ T2 ≤ T3 ≤ T4", serie, ano, ok, tot, "Confirma que os trimestres são acumulados")
    # 3) Totais da Qualidade da Gestão vs séries Receitas e Despesas.
    for serie, prop, prop_qg in (("Receitas", g7.receitasTotais, g7.receitaTotalQG),
                                 ("Despesas", g7.despesaTotal, g7.despesaTotalQG)):
        for ano in (2017, 2018, 2019):
            ok = tot = 0
            for (s, a, p, m), v in valores.items():
                if s != "Qualidade da Gestão" or a != ano or prop_qg not in v:
                    continue
                outro = valores.get((serie, ano, p, m), {}).get(prop)
                if outro is None:
                    continue
                tot += 1
                ok += abs(outro - v[prop_qg]) <= Decimal("0.01")
            registar(f"{serie} igual a Qualidade da Gestão", serie, ano, ok, tot,
                     "Diferenças justificam propriedades separadas (g7:receitaTotalQG / g7:despesaTotalQG)")
    # 4) Recursos Humanos: total = dirigentes + técnicos + outros (sem eleitos) em todos os anos.
    for ano in (2017, 2018, 2019):
        ok = tot = 0
        for (s, a, p, m), v in valores.items():
            if s != "Recursos Humanos" or a != ano:
                continue
            tot += 1
            ok += v.get(g7.trabalhadores) == (v.get(g7.dirigentes, 0) + v.get(g7.tecnicosSuperiores, 0)
                                              + v.get(g7.outrosTrabalhadores, 0))
        registar("trabalhadores = dirigentes + técnicos + outros", "Recursos Humanos", ano, ok, tot,
                 "g7:trabalhadores comparável entre anos (2017 sem eleitos)")
    with destino.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(linhas[0]), delimiter=";")
        w.writeheader()
        w.writerows(linhas)
    return linhas


def norm(coluna: str) -> str:
    """Aceita CODIGO_PERIODO / codigoperiodo / CODIGOPERIODO."""
    t = unicodedata.normalize("NFKD", str(coluna or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", t.upper())


def normaliza_codigo(raw) -> str:
    s = str(raw or "").strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if not re.fullmatch(r"\d{1,4}", s):
        raise ValueError(f"Código INE inesperado: {raw!r}")
    return s.zfill(4)


def valor_tipado(raw, tipo: str):
    """Retorna Literal numérico ou None; nunca converte valores omissos em zero."""
    texto = str(raw if raw is not None else "").strip()
    if texto in ("", "-", "—", "N/A", "NA", "NULL", "null"):
        return None
    try:
        if tipo == "xsd:integer":
            d = Decimal(texto)
            if not d.is_finite() or d != d.to_integral_value():
                raise ValueError(f"Valor não inteiro: {texto!r}")
            return Literal(int(d), datatype=XSD.integer)
        if tipo == "xsd:decimal":
            d = Decimal(texto)
            if not d.is_finite():
                raise ValueError(f"Valor não finito: {texto!r}")
            return Literal(d, datatype=XSD.decimal)
    except (InvalidOperation, ValueError) as e:
        raise ValueError(f"Valor numérico inesperado {texto!r} para {tipo}: {e}") from e
    raise ValueError(f"Tipo não suportado no mapeamento: {tipo!r}")


def ler_csv(caminho: Path, codificacao: str):
    with caminho.open("r", encoding=codificacao, newline="") as entrada:
        leitor = csv.DictReader(entrada, delimiter=";")
        if not leitor.fieldnames:
            raise ValueError(f"CSV sem cabeçalho: {caminho}")
        for linha in leitor:
            yield {norm(k): v for k, v in linha.items() if k is not None}


def ler_xlsx(caminho: Path):
    """Lê a primeira folha de um XLSX em streaming, só com biblioteca padrão.

    Compatível com células do tipo shared string, inline string, número e booleano.
    Não recalcula fórmulas; usa o valor em cache, quando existir.
    """
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(caminho) as pacote:
        shared = []
        if "xl/sharedStrings.xml" in pacote.namelist():
            tree = ET.fromstring(pacote.read("xl/sharedStrings.xml"))
            for si in tree.findall(f"{ns}si"):
                shared.append("".join((el.text or "") for el in si.iter(f"{ns}t")))
        nome_folha = "xl/worksheets/sheet1.xml"
        if nome_folha not in pacote.namelist():
            raise ValueError(f"Não foi encontrada a primeira folha em {caminho}")
        header = None
        with pacote.open(nome_folha) as folha:
            for evento, elemento in ET.iterparse(folha, events=("end",)):
                if elemento.tag != f"{ns}row":
                    continue
                celulas = {}
                for cel in elemento.findall(f"{ns}c"):
                    ref = cel.get("r", "")
                    coluna = re.match(r"[A-Z]+", ref)
                    if not coluna:
                        continue
                    # Referências A, Z, AA -> índices 0, 25, 26
                    i = 0
                    for letra in coluna.group():
                        i = 26 * i + ord(letra) - ord("A") + 1
                    i -= 1
                    v = cel.find(f"{ns}v")
                    val = v.text if v is not None else None
                    if cel.get("t") == "s" and val is not None:
                        val = shared[int(val)]
                    elif cel.get("t") == "inlineStr":
                        val = "".join(el.text or "" for el in cel.iter(f"{ns}t"))
                    celulas[i] = val
                if header is None:
                    header = {i: norm(v) for i, v in celulas.items() if v is not None}
                else:
                    yield {head: celulas.get(i) for i, head in header.items()}
                elemento.clear()


def ler_mapeamento(caminho: Path):
    if not caminho.is_file():
        raise FileNotFoundError(f"Falta a tabela de mapeamento: {caminho}")
    maps = defaultdict(dict)
    with caminho.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f, delimiter=";")
        for registo in r:
            serie = registo["Série"].strip()
            nome = norm(registo["Coluna da fonte"])
            medida = registo["Propriedade RDF proposta"].strip()
            tipo = registo["Tipo"].strip()
            if serie not in SERIES or not medida.startswith("g7:"):
                # A série "Enriquecimento INE" pertence à tarefa 5, não à 6.2 da DGAL.
                if serie == "Enriquecimento INE":
                    continue
                raise ValueError(f"Linha de mapeamento desconhecida: {registo}")
            if nome in maps[serie] and maps[serie][nome] != (medida, tipo):
                raise ValueError(f"Mapeamento contraditório: {serie} / {nome}")
            maps[serie][nome] = (medida, tipo)
    if set(maps) != set(SERIES):
        raise ValueError(f"Mapeamento incompleto: {set(SERIES) - set(maps)}")
    return maps


def grafo(base: str):
    g = Graph()
    prefixos = {
        "ws": base,
        "g7": base + "vocab#",
        "terr": base + "territorio/",
        "per": base + "periodo/",
        "obs": base + "obs/",
        "wsds": base + "dataset/",
        "dcat": str(DCAT),
        "qb": str(QB),
        "sdmx-dimension": str(SDMX_DIM),
        "sdmx-attribute": str(SDMX_ATT),
        "dct": str(DCT),
        "skos": str(SKOS),
        "xsd": str(XSD),
        "prov": str(PROV),
        "adms": str(ADMS),
        "vcard": str(VCARD),
        "foaf": str(FOAF),
    }
    for p, iri in prefixos.items():
        g.bind(p, Namespace(iri))
    return g


def guardar(g: Graph, caminho: Path):
    g.serialize(destination=str(caminho), format="turtle", encoding="utf-8")




def url_http_formato_valido(valor: str, origem: str) -> URIRef | None:
    """Valida apenas o formato HTTP(S); não testa acesso pela Internet."""
    valor = str(valor or "").strip()
    if not valor:
        return None
    u = urlparse(valor)
    if u.scheme not in ("http", "https") or not u.netloc:
        raise ValueError(f"{origem}: URL inválida, usar http(s): {valor!r}")
    return URIRef(valor)


def obter_ramo_github(pasta: Path, ramo_explicito: str | None) -> tuple[str, str]:
    """Descobre o ramo de referência do GitHub sem exigir acesso à Internet.

    Usa preferencialmente a referência local origin/HEAD; senão, o ramo atual.
    Não considera esta deteção uma prova de que o ramo já foi publicado.
    """
    if ramo_explicito:
        return ramo_explicito.strip(), "indicado em --github-ramo"

    comandos = [
        (["symbolic-ref", "--quiet", "--short", "refs/remotes/origin/HEAD"], "referência origin/HEAD"),
        (["branch", "--show-current"], "ramo local atual"),
    ]
    for argumentos, origem in comandos:
        try:
            resultado = subprocess.run(
                ["git", "-C", str(pasta), *argumentos],
                capture_output=True, text=True, timeout=5, check=False,
            )
        except (OSError, subprocess.TimeoutExpired):
            continue
        ramo = resultado.stdout.strip()
        if resultado.returncode == 0 and ramo:
            return ramo.removeprefix("origin/"), origem

    return "main", "valor por omissão (a confirmar no GitHub)"


def url_raw_github(repositorio: str, ramo: str) -> str:
    """Constrói a pasta pública para os .ttl, sem divulgar caminhos locais."""
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repositorio):
        raise ValueError("--github-repo tem de ter o formato UTILIZADOR/REPOSITORIO")
    if not ramo or ramo.startswith("-") or any(c in ramo for c in " \t\r\n\\?#"):
        raise ValueError(f"--github-ramo inválido: {ramo!r}")
    return (
        f"https://raw.githubusercontent.com/{repositorio}/"
        f"{quote(ramo, safe='/-._')}/{GITHUB_SUBPASTA}/"
    )


def ler_fontes_csv(caminho: Path | None) -> dict:
    """URLs reais das páginas DGAL, preenchidas manualmente a partir do portal."""
    if caminho is None:
        return {}
    if not caminho.is_file():
        raise FileNotFoundError(f"CSV de fontes não encontrado: {caminho}")
    fontes = {}
    with caminho.open("r", encoding="utf-8-sig", newline="") as f:
        leitor = csv.DictReader(f, delimiter=";")
        necessarias = {"serie", "ano", "url_dataset_original", "url_download_original"}
        if not necessarias.issubset(leitor.fieldnames or []):
            raise ValueError(f"O CSV --fontes precisa destas colunas: {sorted(necessarias)}")
        for linha in leitor:
            serie = (linha.get("serie") or "").strip()
            ano = int(linha.get("ano") or 0)
            if serie not in SERIES or ano not in (2017, 2018, 2019):
                raise ValueError(f"Série/ano inválidos no CSV de fontes: {serie!r} {ano}")
            chave = (serie, ano)
            if chave in fontes:
                raise ValueError(f"Fonte repetida em --fontes: {chave}")
            fontes[chave] = {
                "dataset": url_http_formato_valido(linha["url_dataset_original"], f"{chave} dataset"),
                "download": url_http_formato_valido(linha["url_download_original"], f"{chave} download"),
            }
    return fontes


def ler_ligacoes_territoriais(caminho: Path) -> dict[str, dict[str, str]]:
    """Correspondência validável de 308 códigos INE com distrito/RA e Wikidata.

    Fonte: Ligacoes-Wikidata.csv disponibilizado pelo grupo. Não deduz distritos
    a partir de nomes ou de heurísticas. Falta de correspondência é erro explícito.
    """
    if not caminho.is_file():
        raise FileNotFoundError(
            f"Falta {caminho.name} junto ao conversor. "
            "É necessário para preencher skos:broader nos 308 municípios."
        )
    ligacoes: dict[str, dict[str, str]] = {}
    distritos: dict[str, tuple[str, str]] = {}
    with caminho.open("r", encoding="utf-8-sig", newline="") as f:
        leitor = csv.DictReader(f, delimiter=";")
        obrigatorios = {
            "CODIGOINE", "MUNICIPIO", "MUNICIPIO_WIKIDATA",
            "DISTRITO_RA_COD", "DISTRITO_RA", "DISTRITO_RA_WIKIDATA",
        }
        if not obrigatorios.issubset(leitor.fieldnames or []):
            raise ValueError(f"Colunas em falta em {caminho.name}: {obrigatorios - set(leitor.fieldnames or [])}")
        for linha in leitor:
            codigo = normaliza_codigo(linha["CODIGOINE"])
            if codigo in ligacoes:
                raise ValueError(f"Código INE repetido em {caminho.name}: {codigo}")
            distrito_codigo = linha["DISTRITO_RA_COD"].strip()
            distrito_nome = linha["DISTRITO_RA"].strip()
            if not re.fullmatch(r"(?:\d{2}|RAA|RAM)", distrito_codigo) or not distrito_nome:
                raise ValueError(f"Distrito/RA inválido para {codigo}: {distrito_codigo!r}")
            ligacoes[codigo] = {
                "municipio": linha["MUNICIPIO"].strip(),
                "wikidata": linha["MUNICIPIO_WIKIDATA"].strip(),
                "distrito_codigo": distrito_codigo,
                "distrito_nome": distrito_nome,
                "distrito_wikidata": linha["DISTRITO_RA_WIKIDATA"].strip(),
            }
            origem = distritos.setdefault(distrito_codigo, (distrito_nome, linha["DISTRITO_RA_WIKIDATA"].strip()))
            if origem != (distrito_nome, linha["DISTRITO_RA_WIKIDATA"].strip()):
                raise ValueError(f"Distrito/RA contraditório: {distrito_codigo}")
            for campo in ("wikidata", "distrito_wikidata"):
                if ligacoes[codigo][campo]:
                    url_http_formato_valido(ligacoes[codigo][campo], f"{codigo} {campo}")
    if len(ligacoes) != 308 or len(distritos) != 20:
        raise ValueError(f"Ligacoes-Wikidata.csv incompleto: {len(ligacoes)} municípios, {len(distritos)} distritos/RA; esperados 308 e 20.")
    return ligacoes


def descrever_original(catalogo: Graph, ws: Namespace, cat, dgal, serie: str, ano: int, slug: str,
                       ficheiro: str, original: URIRef, download: URIRef | None, intervalo) -> None:
    """Descreve no catálogo o conjunto de dados original da DGAL e a distribuição usada."""
    catalogo.add((cat, DCAT.dataset, original))
    catalogo.add((original, RDF.type, DCAT.Dataset))
    catalogo.add((original, DCT.title, Literal(TITULO_ORIGINAL[serie].format(ano=ano), lang="pt")))
    catalogo.add((original, DCT.title, Literal(TITULO_ORIGINAL_EN[serie].format(ano=ano), lang="en")))
    desc_pt, desc_en = DESCRICAO_DATASET[serie]
    nota_pt, nota_en = NOTA_ANO.get((serie, ano), ("", ""))
    catalogo.add((original, DCT.description, Literal(
        f"Conjunto de dados original publicado pela DGAL. {desc_pt} Ano de referência: {ano}. {nota_pt}".strip(), lang="pt")))
    catalogo.add((original, DCT.description, Literal(
        f"Original dataset published by DGAL. {desc_en} Reference year: {ano}. {nota_en}".strip(), lang="en")))
    catalogo.add((original, DCT.publisher, dgal))
    catalogo.add((original, DCT.rightsHolder, dgal))
    catalogo.add((original, DCT.identifier, Literal(str(original).rstrip("/").rsplit("/", 1)[1])))
    catalogo.add((original, DCAT.landingPage, original))
    catalogo.add((original, DCT.license, LICENCA_CC_BY))
    catalogo.add((original, DCT.accessRights, EU_ACCESS.PUBLIC))
    catalogo.add((original, DCT.language, EU_LANG.POR))
    catalogo.add((original, DCT.spatial, EU_COUNTRY.PRT))
    catalogo.add((original, DCT.temporal, intervalo))
    catalogo.add((original, DCT.accrualPeriodicity, EU_FREQUENCY.NEVER))
    for tema in TEMAS:
        catalogo.add((original, DCAT.theme, tema))
    emitido, modificado = DATAS_ORIGINAL[(serie, ano)]
    catalogo.add((original, DCT.issued, Literal(emitido.isoformat(), datatype=XSD.date)))
    if modificado:
        catalogo.add((original, DCT.modified, Literal(modificado.isoformat(), datatype=XSD.date)))
    kw_pt, kw_en = KEYWORDS_ORIGINAL[serie]
    for palavra in kw_pt:
        catalogo.add((original, DCAT.keyword, Literal(palavra, lang="pt")))
    for palavra in kw_en:
        catalogo.add((original, DCAT.keyword, Literal(palavra, lang="en")))
    if CONTACTO_DGAL:
        contacto = ws["contacto/dgal"]
        catalogo.add((contacto, RDF.type, VCARD.Organization))
        catalogo.add((contacto, VCARD.fn, Literal("Direção-Geral das Autarquias Locais", lang="pt")))
        catalogo.add((contacto, VCARD.hasEmail if CONTACTO_DGAL.startswith("mailto:") else VCARD.hasURL,
                      URIRef(CONTACTO_DGAL)))
        catalogo.add((original, DCAT.contactPoint, contacto))
    # Distribuição: o ficheiro efetivamente descarregado e convertido pelo grupo.
    formato = "csv" if ficheiro.lower().endswith(".csv") else "xlsx"
    dist = ws[f"distribuicao/dgal-{slug}-{ano}-{formato}"]
    catalogo.add((original, DCAT.distribution, dist))
    catalogo.add((dist, RDF.type, DCAT.Distribution))
    catalogo.add((dist, DCT.title, Literal(f"Ficheiro DGAL {ficheiro}", lang="pt")))
    catalogo.add((dist, DCT.title, Literal(f"DGAL file {ficheiro}", lang="en")))
    catalogo.add((dist, DCT.identifier, Literal(ficheiro)))
    catalogo.add((dist, DCT.format, EU_FILE.CSV if formato == "csv" else EU_FILE.XLSX))
    catalogo.add((dist, DCAT.mediaType, MEDIA_TYPE[formato]))
    catalogo.add((dist, DCT.license, LICENCA_CC_BY))
    # accessURL: página do conjunto de dados no portal (permitido pelo DCAT-AP);
    # downloadURL só quando o URL direto do ficheiro estiver confirmado em fontes_dgal.csv.
    catalogo.add((dist, DCAT.accessURL, original))
    if download:
        catalogo.add((dist, DCAT.downloadURL, download))


def enriquecer_catalogo(
    catalogo: Graph, base: str, pasta_dados: Path, pasta_saida: Path,
    manifesto: list[dict], url_publica: str | None, fontes: dict,
) -> None:
    """Metadados do perfil do Grupo 7, sem alterar nenhuma observação RDF."""
    ws = Namespace(base)
    grupo = ws["agente/grupo7"]
    dgal = ws["agente/dgal"]
    cat = ws.catalogo
    catalogo.add((dgal, RDF.type, FOAF.Organization))
    catalogo.add((dgal, FOAF.name, Literal("Direção-Geral das Autarquias Locais", lang="pt")))
    catalogo.add((dgal, DCT.type, TIPO_PUBLICADOR.NationalAuthority))
    catalogo.add((grupo, DCT.type, TIPO_PUBLICADOR["Academia-ScientificOrganisation"]))
    catalogo.add((grupo, FOAF.homepage, URIRef("https://github.com/Mota005/WB")))

    # Contacto por URL pública do projeto (não se inventa um endereço de e-mail).
    contacto = ws["contacto/grupo7"]
    catalogo.add((contacto, RDF.type, VCARD.Organization))
    catalogo.add((contacto, VCARD.fn, Literal("Grupo 7 — Web Semântica", lang="pt")))
    catalogo.add((contacto, VCARD.hasURL, URIRef("https://github.com/Mota005/WB")))

    data_modificacao = Literal(date.today().isoformat(), datatype=XSD.date)
    data_publicacao = Literal(DATA_PRIMEIRA_PUBLICACAO.isoformat(), datatype=XSD.date)
    catalogo.add((cat, DCT.issued, data_publicacao))
    catalogo.add((cat, DCT.modified, data_modificacao))
    catalogo.add((cat, FOAF.homepage, URIRef("https://github.com/Mota005/WB")))
    catalogo.add((cat, DCT.language, EU_LANG.POR))
    catalogo.add((cat, DCAT.themeTaxonomy, URIRef("http://publications.europa.eu/resource/authority/data-theme")))
    catalogo.add((cat, DCT.license, LICENCA_CC_BY))
    # Descrição definitiva do catálogo, sem notas internas de rascunho.
    catalogo.remove((cat, DCT.description, None))
    catalogo.add((cat, DCT.description, Literal(
        "Catálogo de 12 conjuntos de dados em RDF/Turtle, derivados dos ficheiros "
        "CSV e XLSX da Direção-Geral das Autarquias Locais (DGAL), referentes a "
        "2017, 2018 e 2019. Organiza os conjuntos de dados em quatro séries "
        "temáticas e descreve as respetivas distribuições, licenças, cobertura "
        "temporal e espacial e proveniência. Descreve também os 12 conjuntos de dados "
        "originais da DGAL no Portal Nacional de Dados Abertos.", lang="pt")))
    catalogo.add((cat, DCT.title, Literal("RDF catalogue of DGAL municipal indicators (2017–2019)", lang="en")))
    catalogo.add((cat, DCT.description, Literal(
        "Catalogue of 12 RDF/Turtle datasets derived from the CSV and XLSX files of the Portuguese "
        "Directorate-General for Local Authorities (DGAL) for 2017, 2018 and 2019, organised in four "
        "thematic series, together with the 12 original DGAL datasets in the Portuguese open data portal.",
        lang="en")))

    # Cobertura temporal das séries (recomendada no perfil): 2017-01-01 a 2019-12-31.
    intervalo_series = ws["intervalo/2017-2019"]
    catalogo.add((intervalo_series, RDF.type, DCT.PeriodOfTime))
    catalogo.add((intervalo_series, DCAT.startDate, Literal("2017-01-01", datatype=XSD.date)))
    catalogo.add((intervalo_series, DCAT.endDate, Literal("2019-12-31", datatype=XSD.date)))

    # Cada série é um recurso distinto com frequência anual e duas categorias EU.
    for serie, slug in SERIES.items():
        serie_uri = ws[f"dataset/serie-{slug}"]
        meta = METADADOS_SERIES[serie]
        catalogo.add((serie_uri, DCT.description, Literal(meta["descricao"], lang="pt")))
        catalogo.add((serie_uri, DCT.description, Literal(meta["descricao_en"], lang="en")))
        catalogo.add((serie_uri, DCT.publisher, grupo))
        catalogo.add((serie_uri, DCAT.contactPoint, contacto))
        catalogo.add((serie_uri, DCT.accrualPeriodicity, EU_FREQUENCY.ANNUAL))
        catalogo.add((serie_uri, DCT.spatial, EU_COUNTRY.PRT))
        catalogo.add((serie_uri, DCT.title, Literal(TITULOS_EN[serie], lang="en")))
        catalogo.add((serie_uri, DCT.temporal, intervalo_series))
        for tema in TEMAS:
            catalogo.add((serie_uri, DCAT.theme, tema))

    intervalos = {}
    for item in manifesto:
        serie = item["serie"]
        ano = int(item["ano"])
        slug = SERIES[serie]
        nome = item["ficheiro_origem"]
        meta = METADADOS_SERIES[serie]
        ds = ws[f"dataset/rdf-{slug}-{ano}"]
        dist = ws[f"distribuicao/rdf-{slug}-{ano}-turtle"]
        dsd_slug = "recursos-humanos-2017" if serie == "Recursos Humanos" and ano == 2017 else slug

        # Um período temporal explícito por ano: 1 Jan a 31 Dez.
        if ano not in intervalos:
            interv = ws[f"intervalo/{ano}"]
            catalogo.add((interv, RDF.type, DCT.PeriodOfTime))
            catalogo.add((interv, DCAT.startDate, Literal(f"{ano}-01-01", datatype=XSD.date)))
            catalogo.add((interv, DCAT.endDate, Literal(f"{ano}-12-31", datatype=XSD.date)))
            intervalos[ano] = interv
        catalogo.add((ds, DCT.temporal, intervalos[ano]))
        catalogo.add((ds, DCT.spatial, EU_COUNTRY.PRT))
        catalogo.add((ds, DCT.language, EU_LANG.POR))
        catalogo.add((ds, DCT.issued, data_publicacao))
        catalogo.add((ds, DCT.modified, data_modificacao))
        catalogo.add((ds, DCAT.contactPoint, contacto))
        catalogo.add((ds, DCT.accessRights, EU_ACCESS.PUBLIC))
        catalogo.add((ds, DCT.license, LICENCA_CC_BY))
        catalogo.add((ds, DCT.rightsHolder, dgal))
        catalogo.add((ds, DCT.conformsTo, ws[f"dsd/{dsd_slug}"]))
        catalogo.add((ds, DCT.accrualPeriodicity, EU_FREQUENCY.NEVER))
        catalogo.add((ds, DCAT.temporalResolution,
                      Literal("P1Y" if serie == "Recursos Humanos" and ano == 2017 else "P3M", datatype=XSD.duration)))
        desc_pt, desc_en = DESCRICAO_DATASET[serie]
        nota_pt, nota_en = NOTA_ANO.get((serie, ano), ("", ""))
        catalogo.add((ds, DCT.description, Literal(
            f"{desc_pt} Ano de referência: {ano}. {nota_pt}".strip(), lang="pt")))
        catalogo.add((ds, DCT.description, Literal(
            f"{desc_en} Reference year: {ano}. {nota_en}".strip(), lang="en")))
        catalogo.add((ds, DCT.title, Literal(f"{TITULOS_EN[serie]} {ano} — data converted to RDF", lang="en")))
        for tema in TEMAS:
            catalogo.add((ds, DCAT.theme, tema))
        for palavra in meta["keywords"]:
            catalogo.add((ds, DCAT.keyword, Literal(palavra, lang="pt")))
        for palavra in meta["keywords_en"]:
            catalogo.add((ds, DCAT.keyword, Literal(palavra, lang="en")))
        # [CORREÇÃO 1-3] Notas específicas de cada série e ano (antes o texto era repetido).
        partes_nota = [
            f"Convertido a partir de {nome}; {item['linhas_sem_medidas']} linhas sem medidas não originaram "
            "observações; valores omissos não foram substituídos por zero; valores numéricos tipados em RDF."
        ]
        if serie in ("Receitas", "Despesas"):
            partes_nota.append("Os valores trimestrais (T1 a T4) são acumulados desde o início do ano.")
            if ano == 2017:
                partes_nota.append("O ficheiro não tem valores anuais (A); o valor de fim de ano é o de T4.")
        if serie == "Qualidade da Gestão":
            partes_nota.append("A receita total e a despesa total usam g7:receitaTotalQG e g7:despesaTotalQG, "
                               "distintas das medidas das séries Receitas e Despesas.")
        if serie == "Recursos Humanos" and ano == 2017:
            partes_nota.append("Só tem o período anual (A). A coluna totalrh inclui os representantes eleitos e foi "
                               "convertida em g7:trabalhadoresInclEleitos; g7:trabalhadores é calculado sem os eleitos "
                               "(dirigentes + técnicos superiores + outros), comparável com 2018 e 2019. Os indicadores "
                               "de eleitos e de cargos políticos têm estrutura estatística própria.")
        catalogo.add((ds, ADMS.versionNotes, Literal(" ".join(partes_nota), lang="pt")))

        # [CORREÇÃO 1 do catálogo] O conjunto de dados ORIGINAL da DGAL é identificado pelo
        # endereço da sua página no portal (fontes_dgal.csv) e descrito no catálogo, com a
        # distribuição CSV/XLSX usada. O derivado aponta para ele (dct:source, prov:wasDerivedFrom).
        fonte = fontes.get((serie, ano), {})
        original = fonte.get("dataset")
        if original is None:
            raise ValueError(f"Falta url_dataset_original em fontes_dgal.csv para {serie} {ano}")
        catalogo.add((ds, DCT.source, original))
        catalogo.add((ds, PROV.wasDerivedFrom, original))
        catalogo.add((ds, DCAT.landingPage, URIRef(PAGINA_PROJETO)))
        descrever_original(catalogo, ws, cat, dgal, serie, ano, slug, nome, original,
                           fonte.get("download"), intervalos[ano])

        # Distribuições: NÃO publicar caminhos file:// privados no catálogo.
        # Os URLs públicos apontam para os Turtle no GitHub (ou para o endereço
        # definido em --url-distribuicoes). Nunca são gerados caminhos file://.
        # Nota: os ficheiros têm de existir realmente nesse endereço após o push.
        ficheiro_rdf = pasta_saida / item["ficheiro_rdf"]
        if not ficheiro_rdf.is_file():
            raise FileNotFoundError(f"Não foi gerado o Turtle da distribuição: {ficheiro_rdf}")
        if url_publica:
            url = URIRef(url_publica.rstrip("/") + "/" + quote(ficheiro_rdf.name))
            catalogo.add((dist, DCAT.accessURL, url))
            catalogo.add((dist, DCAT.downloadURL, url))
        catalogo.add((dist, DCT.license, LICENCA_CC_BY))
        catalogo.add((dist, DCT.issued, data_publicacao))
        catalogo.add((dist, DCT.title, Literal(f"RDF observations {TITULOS_EN[serie]} {ano} (Turtle)", lang="en")))
        catalogo.add((dist, DCT.description, Literal(
            f"RDF/Turtle distribution of the {TITULOS_EN[serie]} series, year {ano}, with {item['observacoes_rdf']} observations.",
            lang="en")))
        catalogo.add((dist, ADMS.status, EU_DISTR_STATUS.COMPLETED))
        catalogo.add((dist, DCAT.byteSize, Literal(ficheiro_rdf.stat().st_size, datatype=XSD.nonNegativeInteger)))
        catalogo.add((dist, DCT.description,
                      Literal(f"Distribuição RDF/Turtle da série {serie}, ano {ano}, com {item['observacoes_rdf']} observações.", lang="pt")))


def relatorio_catalogo(catalogo: Graph, base: str) -> str:
    ws = Namespace(base)
    ds = sum(1 for _ in catalogo.subjects(RDF.type, DCAT.Dataset))
    dist = sum(1 for _ in catalogo.subjects(RDF.type, DCAT.Distribution))
    serie = sum(1 for _ in catalogo.subjects(RDF.type, DCAT.DatasetSeries))
    urls = sum(1 for _ in catalogo.triples((None, DCAT.accessURL, None)))
    return f"Catálogo DCAT-AP: {serie} séries; {ds} datasets; {dist} distribuições; {urls} accessURLs"


def main():
    parser = argparse.ArgumentParser(description="CSV/XLSX DGAL 2017–2019 para RDF/Turtle")
    pasta = Path(__file__).resolve().parent
    parser.add_argument("--dados", type=Path, default=pasta / "dados")
    parser.add_argument("--saida", type=Path, default=pasta / "rdf")
    parser.add_argument("--mapeamento", type=Path, default=pasta / "mapeamento-vocabulario-g7.csv")
    parser.add_argument("--base", default=DEFAULT_BASE)
    parser.add_argument("--url-distribuicoes", default=None,
                        help="URL pública alternativa onde os 12 .ttl estão publicados (sobrepõe o GitHub)")
    parser.add_argument("--github-repo", default=GITHUB_REPO_PADRAO,
                        help=f"Repositório GitHub público UTILIZADOR/REPO (padrão: {GITHUB_REPO_PADRAO})")
    parser.add_argument("--github-ramo", default=None,
                        help="Ramo publicado (quando omitido tenta detetar origin/HEAD ou o ramo local)")
    parser.add_argument("--fontes", type=Path, default=None,
                        help="CSV opcional com URLs originais; se omitido, usa fontes_dgal.csv junto ao script")
    parser.add_argument("--exigir-publicacao", action="store_true",
                        help="Interrompe se faltar URL publica dos 12 Turtle ou das 12 fontes DGAL")
    args = parser.parse_args()
    base = args.base.rstrip("/") + "/"
    dados = args.dados.resolve()
    saida = args.saida.resolve()
    saida.mkdir(parents=True, exist_ok=True)
    if args.url_distribuicoes:
        url_publica = str(url_http_formato_valido(args.url_distribuicoes, "--url-distribuicoes"))
        origem_url = "endereço fornecido em --url-distribuicoes"
    else:
        ramo, origem_url = obter_ramo_github(pasta.parent, args.github_ramo)
        url_publica = url_raw_github(args.github_repo, ramo)
    print(f"Distribuições RDF (URL base): {url_publica}")
    print(f"Referência GitHub: {origem_url}" if not args.url_distribuicoes else f"Origem: {origem_url}")
    ficheiro_fontes = args.fontes.resolve() if args.fontes else (pasta / "fontes_dgal.csv")
    fontes = ler_fontes_csv(ficheiro_fontes if ficheiro_fontes.exists() else None)
    if args.fontes and not args.fontes.resolve().is_file():
        raise FileNotFoundError(f"Nao foi encontrado o CSV de fontes: {args.fontes}")
    origens_confirmadas = sum(1 for f in fontes.values() if f.get("dataset"))
    if args.exigir_publicacao and (not url_publica or origens_confirmadas != len(INPUTS)):
        raise ValueError(
            "Para a publicacao final faltam URLs publicas dos Turtle "
            "ou paginas oficiais em fontes_dgal.csv "
            f"({origens_confirmadas}/{len(INPUTS)} fontes identificadas)."
        )
    mapa = ler_mapeamento(args.mapeamento.resolve())
    ligacoes = ler_ligacoes_territoriais(pasta / "Ligacoes-Wikidata.csv")
    faltam = [fn for _, _, fn, _ in INPUTS if not (dados / fn).is_file()]
    if faltam:
        raise FileNotFoundError("Faltam os seguintes ficheiros na pasta dados/: " + ", ".join(faltam))

    G7 = Namespace(base + "vocab#")
    TERR = Namespace(base + "territorio/")
    PER = Namespace(base + "periodo/")
    WS = Namespace(base)
    territ = grafo(base)
    periodos = grafo(base)
    dsd_g = grafo(base)
    catalogo = grafo(base)

    esquema_territorios = URIRef(base + "esquema/municipios")
    esquema_periodos = URIRef(base + "esquema/periodos")
    for g, esquema, rotulo in [
        (territ, esquema_territorios, "Municípios portugueses (códigos INE)"),
        (periodos, esquema_periodos, "Períodos da DGAL (2017–2019)"),
    ]:
        g.add((esquema, RDF.type, SKOS.ConceptScheme))
        g.add((esquema, DCT.title, Literal(rotulo, lang="pt")))

    municipio_nome = {}
    for ano in (2017, 2018, 2019):
        for p in sorted(PERIODOS_VALIDOS):
            uri = PER[str(ano) if p == "A" else f"{ano}-{p}"]
            periodos.add((uri, RDF.type, SKOS.Concept))
            periodos.add((uri, SKOS.inScheme, esquema_periodos))
            periodos.add((uri, SKOS.notation, Literal(f"{ano}-{p}")))
            periodos.add((uri, SKOS.prefLabel, Literal(f"{ano} — {p}", lang="pt")))
            periodos.add((uri, SKOS.definition, Literal(definicao_periodo(ano, p), lang="pt")))

    agente = WS["agente/grupo7"]
    catalogo.add((agente, RDF.type, FOAF.Organization))
    catalogo.add((agente, FOAF.name, Literal("Grupo 7 — Web Semântica, MEGSI", lang="pt")))
    cat = WS.catalogo
    catalogo.add((cat, RDF.type, DCAT.Catalog))
    catalogo.add((cat, DCT.title, Literal("Catálogo RDF de indicadores municipais da DGAL (2017–2019)", lang="pt")))
    catalogo.add((cat, DCT.publisher, agente))

    for serie, slug in SERIES.items():
        si = WS[f"dataset/serie-{slug}"]
        catalogo.add((si, RDF.type, DCAT.DatasetSeries))
        catalogo.add((si, DCT.title, Literal(serie, lang="pt")))

    manifest = []
    valores_coerencia = {}
    all_obs = set()
    all_mes = defaultdict(set)
    for serie, ano, ficheiro, codificacao in INPUTS:
        serie_slug = SERIES[serie]
        caminho = dados / ficheiro
        rows = ler_csv(caminho, codificacao) if codificacao else ler_xlsx(caminho)
        g = grafo(base)
        colunas_encontradas = set()
        medidas_encontradas = set()
        total_linhas = 0
        vazias = 0
        observacoes = 0
        periodo_contagens = defaultdict(int)
        dataset_uri = WS[f"dataset/rdf-{serie_slug}-{ano}"]
        serie_uri = WS[f"dataset/serie-{serie_slug}"]
        for linha_num, row in enumerate(rows, start=2):
            total_linhas += 1
            if total_linhas == 1:
                if "CODIGOINE" not in row or "CODIGOPERIODO" not in row:
                    raise ValueError(f"{ficheiro}: faltam CODIGOINE ou CODIGO_PERIODO")
            codigo = normaliza_codigo(row.get("CODIGOINE"))
            periodo = str(row.get("CODIGOPERIODO") or "").strip().upper()
            if periodo not in PERIODOS_VALIDOS:
                raise ValueError(f"{ficheiro}, linha {linha_num}: período inválido {periodo!r}")
            meas = []
            for coluna, (propriedade, tipo) in mapa[serie].items():
                if coluna not in row:
                    continue
                colunas_encontradas.add(coluna)
                try:
                    lit = valor_tipado(row.get(coluna), tipo)
                except ValueError as e:
                    raise ValueError(f"{ficheiro}, linha {linha_num}, coluna {coluna}: {e}") from e
                if lit is not None:
                    meas.append((URIRef(base + "vocab#" + propriedade.split(":", 1)[1]), lit))
                    medidas_encontradas.add(propriedade)
            if not meas:
                vazias += 1
                continue
            if serie == "Recursos Humanos" and ano == 2017:
                meas = ajustar_rh_2017(meas, G7, f"{ficheiro}, linha {linha_num}")
                medidas_encontradas.add("g7:trabalhadoresInclEleitos")
            valores_coerencia[(serie, ano, periodo, normaliza_codigo(row.get("CODIGOINE")))] = {
                prop: Decimal(str(lit)) for prop, lit in meas}
            obs_uri = WS[f"obs/{serie_slug}/{ano}/{periodo}/{codigo}"]
            if obs_uri in all_obs:
                raise ValueError(f"URI de observação duplicado: {obs_uri} em {ficheiro}, linha {linha_num}")
            all_obs.add(obs_uri)
            observacoes += 1
            periodo_contagens[periodo] += 1
            m_uri = TERR[codigo]
            nome = str(row.get("DSCAUTARQUIA") or row.get("DSCAUTARQUIA") or "").strip()
            if nome:
                municipio_nome.setdefault(codigo, nome.title())
            g.add((obs_uri, RDF.type, QB.Observation))
            g.add((obs_uri, QB.dataSet, dataset_uri))
            g.add((obs_uri, SDMX_DIM.refArea, m_uri))
            g.add((obs_uri, G7.ano, Literal(str(ano), datatype=XSD.gYear)))
            g.add((obs_uri, G7.periodo, PER[str(ano) if periodo == "A" else f"{ano}-{periodo}"]))
            for prop, literal in meas:
                g.add((obs_uri, prop, literal))

        if not observacoes:
            raise ValueError(f"Não foram produzidas observações para {ficheiro}")
        all_mes[serie].update(medidas_encontradas)
        guardar(g, saida / f"obs-{serie_slug}-{ano}.ttl")
        ds = dataset_uri
        catalogo.add((cat, DCAT.dataset, ds))
        catalogo.add((ds, RDF.type, DCAT.Dataset))
        catalogo.add((ds, RDF.type, QB.DataSet))
        catalogo.add((ds, DCAT.inSeries, serie_uri))
        catalogo.add((ds, DCT.title, Literal(f"{serie} {ano} — dados convertidos para RDF", lang="pt")))
        catalogo.add((ds, DCT.publisher, agente))
        catalogo.add((ds, DCT.identifier, Literal(f"rdf-{serie_slug}-{ano}")))
        # A fonte exata no dados.gov.pt é deixada em aberto para não inventar o URI.
        distr = WS[f"distribuicao/rdf-{serie_slug}-{ano}-turtle"]
        catalogo.add((ds, DCAT.distribution, distr))
        catalogo.add((distr, RDF.type, DCAT.Distribution))
        catalogo.add((distr, DCT.title, Literal(f"Observações RDF {serie} {ano} (Turtle)", lang="pt")))
        catalogo.add((distr, DCT.format, EU_FILE.RDF_TURTLE))
        catalogo.add((distr, DCAT.mediaType, MEDIA_TYPE["ttl"]))
        # RH 2017 tem duas medidas ausentes de RH 2018 e RH 2019.
        nome_dsd = (
            "recursos-humanos-2017"
            if serie == "Recursos Humanos" and ano == 2017
            else serie_slug
        )
        catalogo.add((ds, QB.structure, WS[f"dsd/{nome_dsd}"]))
        manifest.append({
            "serie": serie, "ano": ano, "ficheiro_origem": ficheiro,
            "formato": caminho.suffix.lower().lstrip("."), "codificacao": codificacao or "XLSX",
            "linhas_lidas": total_linhas, "linhas_sem_medidas": vazias,
            "observacoes_rdf": observacoes, "triplos_observacoes": len(g),
            "periodos": ", ".join(f"{k}:{periodo_contagens[k]}" for k in sorted(periodo_contagens)),
            "colunas_mapeadas_encontradas": len(colunas_encontradas),
            "medidas_com_valores": len(medidas_encontradas),
            "campos_propostos_ausentes": ", ".join(sorted(set(mapa[serie]) - colunas_encontradas)),
            "ficheiro_rdf": f"obs-{serie_slug}-{ano}.ttl",
        })
        print(f"{serie:21s} {ano}: {observacoes:4d} obs; {len(g):6d} triplos; "
              f"{vazias:3d} linhas sem valores")

    if set(municipio_nome) != set(ligacoes):
        raise ValueError(
            "Os municípios das observações não coincidem com Ligacoes-Wikidata.csv. "
            f"Em falta: {sorted(set(municipio_nome) - set(ligacoes))}; "
            f"não utilizados: {sorted(set(ligacoes) - set(municipio_nome))}"
        )
    esquema_distritos = URIRef(base + "esquema/distritos-regioes")
    territ.add((esquema_distritos, RDF.type, SKOS.ConceptScheme))
    territ.add((esquema_distritos, DCT.title, Literal("Distritos e regiões autónomas de Portugal", lang="pt")))
    contagem_nomes = defaultdict(int)
    for linha in ligacoes.values():
        contagem_nomes[linha["municipio"]] += 1
    for codigo in sorted(municipio_nome):
        linha = ligacoes[codigo]
        uri = TERR[codigo]
        territorio_pai = TERR[f"distrito-ra/{linha['distrito_codigo']}"]
        territ.add((uri, RDF.type, SKOS.Concept))
        territ.add((uri, SKOS.inScheme, esquema_territorios))
        territ.add((uri, SKOS.notation, Literal(codigo)))
        # [CORREÇÃO 5] prefLabel único: homónimos com o nome DGAL («Lagoa (Algarve)»).
        rotulo = NOMES_OFICIAIS.get(codigo, linha["municipio"])
        if contagem_nomes[linha["municipio"]] > 1:
            rotulo = municipio_nome[codigo]
            if "(" not in rotulo:
                raise ValueError(f"Município homónimo sem desambiguação na fonte DGAL: {codigo} {rotulo}")
        territ.add((uri, SKOS.prefLabel, Literal(rotulo, lang="pt")))
        if rotulo != linha["municipio"]:
            territ.add((uri, SKOS.altLabel, Literal(linha["municipio"], lang="pt")))
        territ.add((uri, SKOS.broader, territorio_pai))
        if linha["wikidata"]:
            territ.add((uri, SKOS.exactMatch, URIRef(linha["wikidata"])))
        territ.add((territorio_pai, RDF.type, SKOS.Concept))
        territ.add((territorio_pai, SKOS.inScheme, esquema_distritos))
        territ.add((territorio_pai, SKOS.prefLabel, Literal(linha["distrito_nome"], lang="pt")))
        territ.add((territorio_pai, SKOS.notation, Literal(linha["distrito_codigo"])))
        if linha["distrito_wikidata"]:
            territ.add((territorio_pai, SKOS.exactMatch, URIRef(linha["distrito_wikidata"])))
    rotulos = defaultdict(list)  # só no esquema dos municípios (distritos podem repetir nomes de capitais)
    for s, p, o in territ.triples((None, SKOS.prefLabel, None)):
        if (s, SKOS.inScheme, esquema_territorios) in set(territ.triples((s, SKOS.inScheme, None))):
            rotulos[str(o)].append(s)
    repetidos = {r: v for r, v in rotulos.items() if len(v) > 1}
    if repetidos:
        raise ValueError(f"skos:prefLabel repetido no território: {sorted(repetidos)}")
    guardar(territ, saida / "territorio.ttl")
    guardar(periodos, saida / "periodos.ttl")

    # Uma DSD por série, exceto Recursos Humanos:
    # - RH 2017: 9 medidas, incluindo cargos políticos e representantes eleitos;
    # - RH 2018-2019: 7 medidas comuns.
    # As medidas não são substituídas por zeros quando não existem na fonte.
    exclusivas_rh_2017 = {"g7:cargosPoliticos", "g7:representantesEleitos", "g7:trabalhadoresInclEleitos"}

    for serie, medidas in all_mes.items():
        slug = SERIES[serie]
        if serie == "Recursos Humanos":
            estruturas = [
                (f"{slug}-2017", medidas),
                (slug, medidas - exclusivas_rh_2017),
            ]
        else:
            estruturas = [(slug, medidas)]

        for nome_dsd, medidas_dsd in estruturas:
            dsd = WS[f"dsd/{nome_dsd}"]
            dsd_g.add((dsd, RDF.type, QB.DataStructureDefinition))

            for dim in (SDMX_DIM.refArea, G7.ano, G7.periodo):
                cod = str(dim).split("#")[-1] or str(dim).split("/")[-1]
                comp = WS[f"componente/{nome_dsd}/dim-{cod}"]
                dsd_g.add((dsd, QB.component, comp))
                dsd_g.add((comp, QB.dimension, dim))

            # Atributo de unidade: o valor está em cada medida do vocabulário
            # (sdmx-attribute:unitMeasure), por isso o atributo liga-se à MeasureProperty.
            comp_unidade = WS[f"componente/{nome_dsd}/atributo-unitMeasure"]
            dsd_g.add((dsd, QB.component, comp_unidade))
            dsd_g.add((comp_unidade, QB.attribute, SDMX_ATT.unitMeasure))
            dsd_g.add((comp_unidade, QB.componentAttachment, QB.MeasureProperty))
            for medida in sorted(medidas_dsd):
                prop = URIRef(base + "vocab#" + medida.split(":", 1)[1])
                comp = WS[f"componente/{nome_dsd}/medida-{medida.split(':', 1)[1]}"]
                dsd_g.add((dsd, QB.component, comp))
                dsd_g.add((comp, QB.measure, prop))
    guardar(dsd_g, saida / "estruturas-dsd.ttl")
    enriquecer_catalogo(catalogo, base, dados, saida, manifest, url_publica, fontes)
    guardar(catalogo, saida / "catalogo-derivados.ttl")
    # Vocabulário definido noutra tarefa do grupo: validar, mas NÃO duplicar.
    # Manter o ficheiro original em RDF/vocabulario-g7_proposta.ttl.
    # Se existir uma cópia gerada anteriormente em RDF/rdf/, remover apenas
    # quando for byte a byte idêntica à original.
    voc_ficheiro = pasta / "vocabulario-g7_proposta.ttl"
    copia_antiga = saida / "vocabulario-g7_proposta.ttl"
    if voc_ficheiro.is_file():
        voc_g = Graph()
        voc_g.parse(voc_ficheiro, format="turtle")
        if base != DEFAULT_BASE:
            raise ValueError("O URI base difere do vocabulário do grupo; alinhar os namespaces.")
        if copia_antiga.is_file():
            if voc_ficheiro.read_bytes() == copia_antiga.read_bytes():
                copia_antiga.unlink()
                print("Vocabulário: cópia idêntica removida de rdf/; mantém-se o original na raiz RDF/.")
            else:
                print("AVISO: existem duas versões diferentes do vocabulário; não foi apagado nenhum ficheiro.")
    elif copia_antiga.is_file():
        voc_g = Graph()
        voc_g.parse(copia_antiga, format="turtle")
        print("Vocabulário encontrado apenas em rdf/; mantido sem alterações.")
    else:
        voc_g = None
        print("AVISO: vocabulário do grupo não encontrado; importar quando estiver disponível.")
    # [VALIDAÇÃO] Todas as medidas usadas nas observações têm de estar declaradas no vocabulário.
    if voc_g is not None:
        declaradas = {str(s) for s in voc_g.subjects(RDF.type, QB.MeasureProperty)}
        usadas = {base + "vocab#" + m.split(":", 1)[1] for ms in all_mes.values() for m in ms}
        em_falta = sorted(u.split("#", 1)[1] for u in usadas - declaradas)
        if em_falta:
            raise ValueError("Medidas usadas mas não declaradas como qb:MeasureProperty no vocabulário: "
                             + ", ".join(em_falta))
        print(f"Vocabulário: as {len(usadas)} medidas usadas estão declaradas.")
    # [VALIDAÇÃO] Controlos de coerência entre períodos e entre séries.
    print("\nControlos de coerência (verificacoes_coerencia.csv):")
    for c in verificacoes_coerencia(valores_coerencia, G7, saida / "verificacoes_coerencia.csv"):
        print(f"  {c['controlo']:47s} {c['serie']:20s} {c['ano']}: {c['casos_ok']}/{c['casos_verificados']}")
    destino_manifesto = saida / "manifesto_conversao.csv"
    with destino_manifesto.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(manifest[0]), delimiter=";")
        writer.writeheader()
        writer.writerows(manifest)
    print(f"\nTOTAL: {sum(m['observacoes_rdf'] for m in manifest)} observações "
          f"em {len(manifest)} datasets; {len(municipio_nome)} municípios identificados")
    print(f"Saída: {saida}")
    print("Território: 308 municípios com skos:broader; 18 distritos + 2 regiões autónomas; ligações Wikidata do CSV.")
    print(relatorio_catalogo(catalogo, base))
    print("IMPORTANTE: confirmar que os 12 URLs abrem ficheiros .ttl e fazer git push se alteraste os dados.")
    print("Exemplo: " + url_publica.rstrip("/") + "/obs-receitas-municipais-2019.ttl")
    confirmadas = sum(1 for f in fontes.values() if f.get("dataset") or f.get("download"))
    if confirmadas < 12:
        print(f"AVISO: {12 - confirmadas} fontes DGAL ainda sem URLs oficiais confirmadas. ")
        print("Preencher as restantes URLs oficiais em fontes_dgal.csv e voltar a executar.")
    print("Seguinte: validar com SHACL -> python validar_shacl.py (requer pip install pyshacl).")
    print("Pendente: registar w3id.org; preencher url_download_original em fontes_dgal.csv e CONTACTO_DGAL, se confirmados.")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, FileNotFoundError, zipfile.BadZipFile) as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        sys.exit(1)
