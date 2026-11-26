#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validação SHACL do RDF do Grupo 7 com as regras do perfil (Tabela 7 do relatório).

Requisitos: pip install pyshacl   (instala também o rdflib)
Execução:   python validar_shacl.py
Lê:         rdf/*.ttl (dados e catálogo gerados pelo converter_dgal.py),
            vocabulario-g7_proposta.ttl (necessário para identificar as medidas)
            e shacl/perfil-g7-shapes.ttl.
Escreve:    rdf/relatorio_shacl.txt e termina com código 1 se houver violações.
As regras marcadas com sh:Warning (recomendadas) não fazem falhar a validação.
"""
import sys
import time
from pathlib import Path

from pyshacl import validate
from rdflib import Graph, Namespace

SH = Namespace("http://www.w3.org/ns/shacl#")
PASTA = Path(__file__).resolve().parent


def main() -> int:
    dados = Graph()
    ficheiros = sorted((PASTA / "rdf").glob("*.ttl")) + [PASTA / "vocabulario-g7_proposta.ttl"]
    for f in ficheiros:
        dados.parse(f, format="turtle")
    shapes = Graph().parse(PASTA / "shacl" / "perfil-g7-shapes.ttl", format="turtle")
    print(f"Dados: {len(ficheiros)} ficheiros, {len(dados)} triplos. A validar (pode demorar alguns minutos)...")
    inicio = time.time()
    conforme, relatorio, texto = validate(
        dados, shacl_graph=shapes, inference="none", advanced=True, allow_warnings=True)
    violacoes = len(list(relatorio.subjects(SH.resultSeverity, SH.Violation)))
    avisos = len(list(relatorio.subjects(SH.resultSeverity, SH.Warning)))
    (PASTA / "rdf" / "relatorio_shacl.txt").write_text(texto, encoding="utf-8")
    print(f"Conforme: {conforme} | violações: {violacoes} | avisos: {avisos} | {time.time() - inicio:.0f} s")
    print("Relatório completo em rdf/relatorio_shacl.txt")
    return 0 if violacoes == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
