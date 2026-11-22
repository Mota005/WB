#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Conversor reproduzível DGAL (2017-2019) -> RDF/Turtle.

Grupo 7 | UC Web Semântica | tarefa 6.2: DSD separada para RH 2017.

Entradas: 12 CSV/XLSX da DGAL na pasta dados/ e
          mapeamento-vocabulario-g7.csv na raiz do projeto.
Saídas:   catálogo RDF (derivados), SKOS territorial e períodos,
          DSDs, 12 ficheiros com observações e manifesto de conversão.

Não inventa URLs das páginas originais nem dados em falta. A publicação
resolúvel do namespace w3id.org continua pendente de registo.

Requisitos: Python >=3.10; pip install rdflib
Execução:   python converter_dgal.py
            python converter_dgal.py --dados dados --saida rdf
"""

from __future__ import annotations

import argparse
import csv
import shutil
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

from rdflib import Graph, Literal, Namespace, URIRef
from rdflib.namespace import DCTERMS as DCT, FOAF, RDF, RDFS, SKOS, XSD

DEFAULT_BASE = "https://w3id.org/ws-megsi-g7/"  # PROPOSTA (a registar)
QB = Namespace("http://purl.org/linked-data/cube#")
DCAT = Namespace("http://www.w3.org/ns/dcat#")
SDMX_DIM = Namespace("http://purl.org/linked-data/sdmx/2009/dimension#")
SDMX_ATT = Namespace("http://purl.org/linked-data/sdmx/2009/attribute#")
EU_FILE = Namespace("http://publications.europa.eu/resource/authority/file-type/")
EU_LICENSE = Namespace("http://publications.europa.eu/resource/authority/licence/")

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
    }
    for p, iri in prefixos.items():
        g.bind(p, Namespace(iri))
    return g


def guardar(g: Graph, caminho: Path):
    g.serialize(destination=str(caminho), format="turtle", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="CSV/XLSX DGAL 2017–2019 para RDF/Turtle")
    pasta = Path(__file__).resolve().parent
    parser.add_argument("--dados", type=Path, default=pasta / "dados")
    parser.add_argument("--saida", type=Path, default=pasta / "rdf")
    parser.add_argument("--mapeamento", type=Path, default=pasta / "mapeamento-vocabulario-g7.csv")
    parser.add_argument("--base", default=DEFAULT_BASE)
    args = parser.parse_args()
    base = args.base.rstrip("/") + "/"
    dados = args.dados.resolve()
    saida = args.saida.resolve()
    saida.mkdir(parents=True, exist_ok=True)
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
    print("Nota: adicionar vocabulário-g7, URIs reais das fontes e validação SHACL nas fases seguintes.")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, FileNotFoundError, zipfile.BadZipFile) as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        sys.exit(1)
