#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Conversor reproduzível DGAL (2017-2019) -> RDF/Turtle.

Grupo 7 | UC Web Semântica | tarefa 6.2: observações e catálogo DCAT-AP.

Entradas: 12 CSV/XLSX da DGAL na pasta dados/ e
          mapeamento-vocabulario-g7.csv na raiz do projeto.
Saídas:   catálogo RDF (derivados), SKOS territorial e períodos,
          DSDs, 12 ficheiros com observações e manifesto de conversão.

Metadados DCAT-AP do perfil do Grupo 7: catálogo, quatro séries, 12 datasets,
12 distribuições, temas, cobertura temporal/espacial, proveniência e licença.
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
Execução:   python converter_dgal.py
            python converter_dgal.py --dados dados --saida rdf
"""

from __future__ import annotations

import argparse
import csv
import shutil
import subprocess
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from urllib.parse import quote, urlparse
from collections import defaultdict
from decimal import Decimal, InvalidOperation
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

# Descrições e palavras-chave criadas para o perfil do Grupo 7.
TEMAS = (EU_THEME.ECON, EU_THEME.GOVE)
METADADOS_SERIES = {
    "Receitas": {
        "descricao": "Receitas municipais de 308 municípios portugueses, incluindo impostos, transferências, financiamento e outras receitas; os trimestres apresentam valores acumulados.",
        "descricao_en": "Municipal revenues for Portugal's 308 municipalities, including taxes, transfers, funding and other revenue; quarterly values are cumulative.",
        "keywords": ("receitas municipais", "impostos municipais", "transferências", "finanças locais"),
    },
    "Despesas": {
        "descricao": "Despesas municipais de 308 municípios portugueses, incluindo despesas correntes, investimento, pessoal e transferências; os trimestres apresentam valores acumulados.",
        "descricao_en": "Municipal expenditure for Portugal's 308 municipalities, including current expenses, investment, staff and transfers; quarterly values are cumulative.",
        "keywords": ("despesas municipais", "investimento", "despesas com pessoal", "finanças locais"),
    },
    "Recursos Humanos": {
        "descricao": "Indicadores de recursos humanos dos 308 municípios portugueses, incluindo trabalhadores, dirigentes, categorias e tipos de vínculo; 2017 contém apenas o período anual.",
        "descricao_en": "Human resources indicators for Portugal's 308 municipalities, including staff, management categories and employment contracts; 2017 includes annual observations only.",
        "keywords": ("recursos humanos", "trabalhadores municipais", "vínculos laborais", "dirigentes"),
    },
    "Qualidade da Gestão": {
        "descricao": "Indicadores de gestão pública e sustentabilidade financeira dos 308 municípios portugueses, incluindo dívida, execução orçamental e prazos de pagamento.",
        "descricao_en": "Public management and financial sustainability indicators for Portugal's 308 municipalities, including debt, budget execution and payment times.",
        "keywords": ("gestão pública", "endividamento", "execução orçamental", "sustentabilidade financeira"),
    },
}


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
    catalogo.add((cat, DCT.language, EU_LANG.POR))
    catalogo.add((cat, DCAT.themeTaxonomy, URIRef("http://publications.europa.eu/resource/authority/data-theme")))
    catalogo.add((cat, DCT.license, LICENCA_CC_BY))
    # Substituir o texto de rascunho que afirmava nao haver qualquer fonte.
    catalogo.remove((cat, DCT.description, None))
    catalogo.add((cat, DCT.description, Literal(
        "Catálogo de 12 conjuntos de dados RDF derivados de CSV/XLSX da DGAL "
        "(2017-2019), com quatro series tematicas, metadados de distribuicao, "
        "licencas e proveniencia. URLs oficiais incluidas apenas quando confirmadas.", lang="pt")))

    # Cada série é um recurso distinto com frequência anual e duas categorias EU.
    for serie, slug in SERIES.items():
        serie_uri = ws[f"dataset/serie-{slug}"]
        meta = METADADOS_SERIES[serie]
        catalogo.add((serie_uri, DCT.description, Literal(meta["descricao"], lang="pt")))
        catalogo.add((serie_uri, DCT.description, Literal(meta["descricao_en"], lang="en")))
        catalogo.add((serie_uri, DCT.publisher, grupo))
        catalogo.add((serie_uri, DCT.accrualPeriodicity, EU_FREQUENCY.ANNUAL))
        catalogo.add((serie_uri, DCT.spatial, EU_COUNTRY.PRT))
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
        catalogo.add((ds, DCT.accessRights, EU_ACCESS.PUBLIC))
        catalogo.add((ds, DCT.license, LICENCA_CC_BY))
        catalogo.add((ds, DCT.rightsHolder, dgal))
        catalogo.add((ds, DCT.conformsTo, ws[f"dsd/{dsd_slug}"]))
        catalogo.add((ds, DCT.accrualPeriodicity, EU_FREQUENCY.NEVER))
        catalogo.add((ds, DCAT.temporalResolution,
                      Literal("P1Y" if serie == "Recursos Humanos" and ano == 2017 else "P3M", datatype=XSD.duration)))
        catalogo.add((ds, DCT.description,
                      Literal(meta["descricao_en"] + f" Reference year: {ano}.", lang="en")))
        for tema in TEMAS:
            catalogo.add((ds, DCAT.theme, tema))
        for palavra in meta["keywords"]:
            catalogo.add((ds, DCAT.keyword, Literal(palavra, lang="pt")))
        nota = (f"Convertido a partir de {nome}; {item['linhas_sem_medidas']} linhas sem medidas não originaram "
                "observações; valores omissos não foram substituídos por zero; valores numéricos tipados em RDF. "
                "Os dados trimestrais são acumulados. Os dois indicadores adicionais de Recursos Humanos "
                "de 2017 foram separados em estrutura estatística própria." if serie == "Recursos Humanos" else
                f"Convertido a partir de {nome}; {item['linhas_sem_medidas']} linhas sem medidas não originaram "
                "observações; valores omissos não foram substituídos por zero; valores numéricos tipados em RDF. "
                "Os dados trimestrais são acumulados.")
        catalogo.add((ds, ADMS.versionNotes, Literal(nota, lang="pt")))

        # Proveniência: identificar a cópia local sem divulgar caminhos do computador.
        # Quando houver URL oficial, essa passa a ser a origem e a landing page.
        fonte = fontes.get((serie, ano), {})
        fonte_local = ws[f"fonte/dgal/{slug}/{ano}"]
        catalogo.add((fonte_local, RDF.type, FOAF.Document))
        catalogo.add((fonte_local, DCT.identifier, Literal(nome)))
        catalogo.add((fonte_local, DCT.title,
                      Literal(f"Ficheiro DGAL original: {nome}", lang="pt")))
        catalogo.add((fonte_local, DCT.format,
                      EU_FILE.CSV if nome.lower().endswith(".csv") else EU_FILE.XLSX))
        origem = fonte.get("dataset") or fonte.get("download") or fonte_local
        catalogo.add((ds, DCT.source, origem))
        catalogo.add((ds, PROV.wasDerivedFrom, origem))
        if origem != fonte_local:
            catalogo.add((fonte_local, DCT.source, origem))
        if fonte.get("dataset"):
            catalogo.add((ds, DCAT.landingPage, fonte["dataset"]))

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

    agente = WS["agente/grupo7"]
    catalogo.add((agente, RDF.type, FOAF.Organization))
    catalogo.add((agente, FOAF.name, Literal("Grupo 7 — Web Semântica, MEGSI", lang="pt")))
    cat = WS.catalogo
    catalogo.add((cat, RDF.type, DCAT.Catalog))
    catalogo.add((cat, DCT.title, Literal("Catálogo RDF de indicadores municipais da DGAL (2017–2019)", lang="pt")))
    catalogo.add((cat, DCT.publisher, agente))
    catalogo.add((cat, DCT.description, Literal(
        "Metadados dos 12 conjuntos RDF derivados dos ficheiros CSV/XLSX da DGAL. "
        "A ligação às páginas de origem no dados.gov.pt deve ser acrescentada "
        "quando os respetivos URIs forem verificados.", lang="pt")))

    for serie, slug in SERIES.items():
        si = WS[f"dataset/serie-{slug}"]
        catalogo.add((si, RDF.type, DCAT.DatasetSeries))
        catalogo.add((si, DCT.title, Literal(serie, lang="pt")))

    manifest = []
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
        catalogo.add((ds, DCT.description, Literal(
            f"Conversão do ficheiro DGAL {ficheiro}; {observacoes} observações não vazias. "
            "Códigos de município e período preservados. Fonte disponível no Portal Nacional de Dados Abertos.", lang="pt")))
        # A fonte exata no dados.gov.pt é deixada em aberto para não inventar o URI.
        distr = WS[f"distribuicao/rdf-{serie_slug}-{ano}-turtle"]
        catalogo.add((ds, DCAT.distribution, distr))
        catalogo.add((distr, RDF.type, DCAT.Distribution))
        catalogo.add((distr, DCT.title, Literal(f"Observações RDF {serie} {ano} (Turtle)", lang="pt")))
        catalogo.add((distr, DCT.format, EU_FILE.RDF_TURTLE))
        catalogo.add((distr, DCAT.mediaType, URIRef("https://www.iana.org/assignments/media-types/text/turtle")))
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

    for codigo, nome in municipio_nome.items():
        uri = TERR[codigo]
        territ.add((uri, RDF.type, SKOS.Concept))
        territ.add((uri, SKOS.inScheme, esquema_territorios))
        territ.add((uri, SKOS.notation, Literal(codigo)))
        territ.add((uri, SKOS.prefLabel, Literal(nome, lang="pt")))
    guardar(territ, saida / "territorio.ttl")
    guardar(periodos, saida / "periodos.ttl")

    # Uma DSD por série, exceto Recursos Humanos:
    # - RH 2017: 9 medidas, incluindo cargos políticos e representantes eleitos;
    # - RH 2018-2019: 7 medidas comuns.
    # As medidas não são substituídas por zeros quando não existem na fonte.
    exclusivas_rh_2017 = {"g7:cargosPoliticos", "g7:representantesEleitos"}

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

            for medida in sorted(medidas_dsd):
                prop = URIRef(base + "vocab#" + medida.split(":", 1)[1])
                comp = WS[f"componente/{nome_dsd}/medida-{medida.split(':', 1)[1]}"]
                dsd_g.add((dsd, QB.component, comp))
                dsd_g.add((comp, QB.measure, prop))
    guardar(dsd_g, saida / "estruturas-dsd.ttl")
    enriquecer_catalogo(catalogo, base, dados, saida, manifest, url_publica, fontes)
    guardar(catalogo, saida / "catalogo-derivados.ttl")
    # O vocabulário da secção 4.2 é um artefacto separado, mas deve ser
    # carregado no mesmo repositório RDF que as observações.
    voc_ficheiro = pasta / "vocabulario-g7_proposta.ttl"
    if voc_ficheiro.is_file():
        voc_g = Graph()
        voc_g.parse(voc_ficheiro, format="turtle")  # valida sintaxe antes de copiar
        if base != DEFAULT_BASE:
            raise ValueError("O URI base difere do vocabulário proposto. Alinhar primeiro os namespaces.")
        destino_vocab = saida / voc_ficheiro.name
        if voc_ficheiro.resolve() != destino_vocab.resolve():
            shutil.copyfile(voc_ficheiro, destino_vocab)
    else:
        print("AVISO: não foi encontrado vocabulario-g7_proposta.ttl junto ao script. "
              "O RDF gerado usa propriedades g7: que precisam das definições do vocabulário.")
    destino_manifesto = saida / "manifesto_conversao.csv"
    with destino_manifesto.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(manifest[0]), delimiter=";")
        writer.writeheader()
        writer.writerows(manifest)
    print(f"\nTOTAL: {sum(m['observacoes_rdf'] for m in manifest)} observações "
          f"em {len(manifest)} datasets; {len(municipio_nome)} municípios identificados")
    print(f"Saída: {saida}")
    print(relatorio_catalogo(catalogo, base))
    print("IMPORTANTE: confirmar que os 12 URLs abrem ficheiros .ttl e fazer git push se alteraste os dados.")
    print("Exemplo: " + url_publica.rstrip("/") + "/obs-receitas-municipais-2019.ttl")
    confirmadas = sum(1 for f in fontes.values() if f.get("dataset") or f.get("download"))
    if confirmadas < 12:
        print(f"AVISO: {12 - confirmadas} fontes DGAL ainda sem URLs oficiais confirmadas. ")
        print("Preencher as restantes URLs oficiais em fontes_dgal.csv e voltar a executar.")
    print("Pendente: verificar URLs de publicacao, registar w3id.org e executar validacao formal SHACL na tarefa 6.3.")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, FileNotFoundError, zipfile.BadZipFile) as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        sys.exit(1)
