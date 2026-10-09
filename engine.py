"""
Módulo 3 — motor de demanda ótima (tradução de CalcularCusto / CalcularCustoDemanda).

Não replica a planilha (abas, VBA, células). Extrai só a matemática e aplica
em cada UC Grupo A, em janela rolante de até 12 meses civis.

Tradução VBA → NumPy
---------------------
O VBA testa cada kW candidato em um For e, por candidato, soma 12 meses.
Aqui o grid vira um vetor (K,) e as medidas um vetor (M,); o custo mensal é
uma matriz (K, M) via broadcast. O ótimo é argmin da soma no eixo M.
Isso é o mesmo laço de CalcularCusto, sem iterar kW em Python.

Janela: para a competência (ano, mês) = t, entram os meses civis de t−11 … t
(inclusive). Ex.: em 04/2025 usa 05/2024 … 04/2025. Meses sem dado no histórico
ficam de fora; ``n_meses_janela`` é a quantidade de meses encontrados (1–12).

Execução (pasta pai Modulo1_API):

    pip install -r modulo3/requirements.txt
    python -m modulo3.engine
"""

from __future__ import annotations  # permite anotações de tipo com classes ainda não definidas

import argparse  # interpreta argumentos da linha de comando (CLI)
import logging  # registra mensagens de log no processamento
import sqlite3  # acesso ao banco neoenergia.sqlite
import sys  # saída em stderr e código de saída do programa
from pathlib import Path  # caminhos de arquivo de forma portável
from typing import Any  # tipo genérico para dicionários de resultado

import numpy as np  # vetores/matrizes e operações numéricas (núcleo da otimização)
import pandas as pd  # DataFrames para histórico e gravação no SQLite

from .tarifas import (  # tarifas Cosern extraídas da planilha v3.2
    COSERN_PLANILHA_V32,
    TarifaDemanda,
    TarifaEnergia,
    TariffManager,
)

logger = logging.getLogger(__name__)  # logger nomeado com o módulo atual

_MODULO3 = Path(__file__).resolve().parent  # pasta modulo3/ (onde está este arquivo)
_BANCO_PADRAO = _MODULO3.parent / "modulo2" / "data" / "neoenergia.sqlite"  # SQLite padrão
_PISO_KW = 30.0  # menor kW candidato na grade (regra VBA/planilha)

# Consulta: histórico Grupo A só em meses com demanda informada no PDF (evita API = contrato).
_SQL_HISTORICO = """
SELECT
    v.codigo_uc,
    v.ano,
    v.mes,
    v.demanda_medida_p,
    v.demanda_medida_fp,
    v.consumo_ativa_p,
    v.consumo_ativa_fp
FROM v_consumo_a_efetivo v
JOIN uc u ON u.codigo_uc = v.codigo_uc
INNER JOIN consumo_pdf_mensal p
    ON p.codigo_uc = v.codigo_uc
   AND p.ano = v.ano
   AND p.mes = v.mes
WHERE u.classe_tarifaria = 'A'
  AND (
        p.demanda_faturada_fp IS NOT NULL
     OR p.demanda_faturada_p IS NOT NULL
     OR p.demanda_medida_fp IS NOT NULL
     OR p.demanda_medida_p IS NOT NULL
  )
ORDER BY v.codigo_uc, v.ano, v.mes
"""

# DDL da tabela onde os resultados do estudo são persistidos.
_DDL_ESTUDO = """
CREATE TABLE IF NOT EXISTS estudo_demanda_otima (
    codigo_uc              TEXT NOT NULL,
    ano                    INTEGER NOT NULL,
    mes                    INTEGER NOT NULL,
    n_meses_janela         INTEGER NOT NULL,
    demanda_otima_verde    REAL,
    custo_demanda_verde    REAL,
    custo_energia_verde    REAL,
    custo_total_verde      REAL,
    demanda_otima_azul_p   REAL,
    demanda_otima_azul_fp  REAL,
    custo_demanda_azul_p   REAL,
    custo_demanda_azul_fp  REAL,
    custo_energia_azul     REAL,
    custo_total_azul       REAL,
    modalidade_recomendada TEXT NOT NULL,
    PRIMARY KEY (codigo_uc, ano, mes)
);
"""


# ---------------------------------------------------------------------------
# Núcleo: CalcularCustoDemanda (VBA, ramos idênticos)
# ---------------------------------------------------------------------------


def custo_demanda_mes(
    contratada: np.ndarray,
    medida: np.ndarray,
    tarifa: TarifaDemanda,
) -> np.ndarray:
    """Custo de demanda por (candidato, mês). Shape (K, M).

    VBA CalcularCustoDemanda:
      If medida > contratada * 1.05:
          medida * TD + (medida - contratada) * TD_ultr
      ElseIf medida < contratada:
          medida * TD + (contratada - medida) * TD_lim
      Else:   ' igual ou até +5%
          medida * TD
    """
    # Coluna (K,1): cada linha é um kW contratado candidato.
    c = np.asarray(contratada, dtype=np.float64).reshape(-1, 1)
    # Linha (1,M): cada coluna é a demanda medida/faturada de um mês da janela.
    m = np.asarray(medida, dtype=np.float64).reshape(1, -1)
    # Caso base VBA: medida na faixa [contratada, contratada×1,05] → só TUSD normal × medida.
    custo = m * tarifa.normal
    # Onde medida < contratada: acrescenta custo liminar (energia ociosa) no deficitário.
    custo = np.where(
        m < c,
        m * tarifa.normal + (c - m) * tarifa.liminar,
        custo,
    )
    # Onde medida > contratada×1,05: acrescenta ultrapassagem no excedente.
    custo = np.where(
        m > (c * 1.05),
        m * tarifa.normal + (m - c) * tarifa.ultrapassagem,
        custo,
    )
    return custo  # matriz (K, M) com custo de demanda em R$ por candidato e por mês


def grade_candidatos(medidas: np.ndarray) -> np.ndarray:
    """Inteiros de Int(max(30, min)) até Fix(max(30, max)) + 1, inclusive.

    VBA: se min ou max < 30, eleva a 30; intervalo
    Int(menor) .. Fix(maior)+1  (o +2 no UBound gera esse kW extra).
    """
    medidas = np.asarray(medidas, dtype=np.float64)  # garante array NumPy float
    medidas = medidas[np.isfinite(medidas)]  # remove NaN/Inf da série de medidas
    if medidas.size == 0:  # sem medidas válidas não há grade
        return np.array([], dtype=np.int64)  # vetor vazio de candidatos
    menor = max(float(np.min(medidas)), _PISO_KW)  # piso 30 kW e mínimo da janela
    maior = max(float(np.max(medidas)), _PISO_KW)  # piso também no máximo (VBA)
    lo = int(menor)  # Int(): trunca para inteiro (kW mínimo candidato)
    hi = int(maior) + 1  # Fix(maior)+1 no VBA (kW máximo candidato)
    if hi < lo:  # intervalo inválido
        return np.array([], dtype=np.int64)
    return np.arange(lo, hi + 1, dtype=np.int64)  # todos os kW inteiros candidatos, inclusive hi


def otimizar_demanda(medidas: np.ndarray, tarifa: TarifaDemanda) -> tuple[int, float]:
    """Um kW inteiro que minimiza a soma de CalcularCustoDemanda na janela."""
    medidas = np.asarray(medidas, dtype=np.float64)  # série de demanda (um valor por mês)
    candidatos = grade_candidatos(medidas)  # kW inteiros a testar
    if candidatos.size == 0 or not np.isfinite(medidas).any():  # não há o que otimizar
        return 0, float("nan")  # sinaliza ausência de resultado
    # Soma custos de demanda na janela (eixo M) para cada candidato kW (eixo K).
    totais = custo_demanda_mes(candidatos, medidas, tarifa).sum(axis=1)
    idx = int(np.argmin(totais))  # índice do candidato com menor custo total
    return int(candidatos[idx]), float(totais[idx])  # kW ótimo e custo de demanda na janela


def custo_energia(
    consumo_p: np.ndarray,
    consumo_fp: np.ndarray,
    tarifa: TarifaEnergia,
) -> float:
    """TUSD+TE ponta e FP; reativo propositalmente ignorado (como CalcularCusto)."""
    p = np.nan_to_num(np.asarray(consumo_p, dtype=np.float64), nan=0.0)  # kWh ponta por mês
    fp = np.nan_to_num(np.asarray(consumo_fp, dtype=np.float64), nan=0.0)  # kWh FP por mês
    # Soma na janela: consumo × (TUSD+TE) ponta e fora de ponta.
    return float(np.sum(p * tarifa.ponta + fp * tarifa.fora_ponta))


def estudar_janela(
    ponta: np.ndarray,
    fp: np.ndarray,
    cons_p: np.ndarray,
    cons_fp: np.ndarray,
    tarifas: TariffManager,
) -> dict[str, Any]:
    """Verde (max P/FP) + Azul P + Azul FP; totais com energia; flag de modalidade."""
    ponta = np.asarray(ponta, dtype=np.float64)  # demanda ponta em cada mês da janela
    fp = np.asarray(fp, dtype=np.float64)  # demanda FP em cada mês da janela
    verde_med = np.fmax(ponta, fp)  # Verde: um único kW por mês = max(ponta, FP)

    # Otimização de demanda contratada (só custo de demanda) para cada modalidade.
    kw_v, custo_dv = otimizar_demanda(verde_med, tarifas.verde_demanda)
    kw_p, custo_dp = otimizar_demanda(ponta, tarifas.azul_ponta_demanda)
    kw_fp, custo_dfp = otimizar_demanda(fp, tarifas.azul_fp_demanda)

    # Custo de energia na janela (independe do kW contratado escolhido).
    e_verde = custo_energia(cons_p, cons_fp, tarifas.verde_energia)
    e_azul = custo_energia(cons_p, cons_fp, tarifas.azul_energia)
    total_v = custo_dv + e_verde  # custo total simulado na modalidade Verde
    total_a = custo_dfp + custo_dp + e_azul  # custo total simulado na modalidade Azul
    # Excel: IF(custoVerde < custoAzul, "VERDE", "AZUL") — empate vai para Azul.
    modalidade = "VERDE" if total_v < total_a else "AZUL"

    return {
        "n_meses_janela": int(len(ponta)),  # quantos meses entraram nesta janela
        "demanda_otima_verde": kw_v,
        "custo_demanda_verde": custo_dv,
        "custo_energia_verde": e_verde,
        "custo_total_verde": total_v,
        "demanda_otima_azul_p": kw_p,
        "demanda_otima_azul_fp": kw_fp,
        "custo_demanda_azul_p": custo_dp,
        "custo_demanda_azul_fp": custo_dfp,
        "custo_energia_azul": e_azul,
        "custo_total_azul": total_a,
        "modalidade_recomendada": modalidade,
    }


def _competencias_janela_civil(ano: int, mes: int, tamanho: int = 12) -> list[tuple[int, int]]:
    """Doze meses civis, do mais antigo ao mais recente, terminando em (ano, mes)."""
    y, m = int(ano), int(mes)  # competência final da janela (mês t)
    itens: list[tuple[int, int]] = []  # lista temporária, do mais recente ao mais antigo
    # Repete tamanho vezes: adiciona (y,m) e retrocede um mês civil.
    for _ in range(tamanho):
        itens.append((y, m))  # inclui este ano/mês na janela
        m -= 1  # mês anterior
        if m < 1:  # virou janeiro → dezembro do ano anterior
            m = 12
            y -= 1
    itens.reverse()  # ordena do mais antigo ao mais recente (cronológico)
    return itens


def estudar_historico_uc(bloco: pd.DataFrame, tarifas: TariffManager) -> list[dict[str, Any]]:
    """Uma linha de estudo por competência; janela civil de 12 meses até o mês t."""
    # Ordena o histórico da UC por tempo (estável).
    bloco = bloco.sort_values(["ano", "mes"], kind="mergesort").reset_index(drop=True)
    linhas: list[dict[str, Any]] = []  # acumula um dict de estudo por competência
    # Para cada mês que existe no histórico, calcula o estudo usando a janela civil até esse mês.
    for _, atual in bloco.iterrows():
        chaves = _competencias_janela_civil(int(atual["ano"]), int(atual["mes"]))
        janela = pd.DataFrame(chaves, columns=["ano", "mes"])  # grade dos 12 meses civis
        # Mantém só linhas do histórico que caem nesses (ano, mes); meses sem dado somem.
        w = bloco.merge(janela, on=["ano", "mes"], how="inner").sort_values(
            ["ano", "mes"],
            kind="mergesort",
        )
        if w.empty:  # competência sem nenhum mês na janela (raro)
            continue  # pula esta competência
        # Aplica o núcleo de otimização nos arrays da janela filtrada.
        estudo = estudar_janela(
            w["demanda_medida_p"].to_numpy(),
            w["demanda_medida_fp"].to_numpy(),
            w["consumo_ativa_p"].to_numpy(),
            w["consumo_ativa_fp"].to_numpy(),
            tarifas,
        )
        estudo["codigo_uc"] = str(atual["codigo_uc"])  # identifica a UC
        estudo["ano"] = int(atual["ano"])  # competência estudada (ano)
        estudo["mes"] = int(atual["mes"])  # competência estudada (mês)
        linhas.append(estudo)  # guarda resultado desta competência
    return linhas


# ---------------------------------------------------------------------------
# SQLite
# ---------------------------------------------------------------------------


def conectar(banco: Path) -> sqlite3.Connection:
    if not banco.is_file():  # arquivo do banco precisa existir
        raise FileNotFoundError(
            f"SQLite não encontrado: {banco}. Rode python -m modulo2.tratamento antes."
        )
    con = sqlite3.connect(str(banco))  # abre conexão SQLite
    con.execute("PRAGMA foreign_keys = ON")  # habilita FKs
    con.execute("PRAGMA journal_mode = WAL")  # modo WAL para concorrência
    try:
        from modulo2 import db as m2db  # import tardio para migrar schema/views

        m2db.migrar_schema(con)  # garante views/tabelas atualizadas (ex.: v_consumo_a_efetivo)
        con.commit()  # persiste migrações
    except Exception:
        logger.debug("Migracao modulo2 nao aplicada (view pode estar desatualizada).")
    return con


def garantir_tabela(con: sqlite3.Connection) -> None:
    con.executescript(_DDL_ESTUDO)  # cria estudo_demanda_otima se não existir


def ler_historico(con: sqlite3.Connection, codigo_uc: str | None = None) -> pd.DataFrame:
    sql = _SQL_HISTORICO  # query base (todas as UCs Grupo A com PDF)
    params: tuple[Any, ...] = ()  # parâmetros SQL (vazio = todas)
    if codigo_uc:  # filtro opcional para uma UC
        sql = _SQL_HISTORICO.replace(
            "WHERE u.classe_tarifaria = 'A'",
            "WHERE u.classe_tarifaria = 'A' AND v.codigo_uc = ?",
        )
        params = (codigo_uc,)  # bind do código da UC
    df = pd.read_sql_query(sql, con, params=params)  # carrega histórico em DataFrame
    # Converte colunas numéricas (SQLite pode devolver texto).
    for col in (
        "demanda_medida_p",
        "demanda_medida_fp",
        "consumo_ativa_p",
        "consumo_ativa_fp",
    ):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    # Exige demanda ponta e FP para entrar no estudo (regra do motor).
    df = df.dropna(subset=["demanda_medida_p", "demanda_medida_fp"])
    return df


def gravar_estudo(con: sqlite3.Connection, resultado: pd.DataFrame, *, substituir: bool) -> None:
    garantir_tabela(con)  # DDL idempotente
    if substituir:  # modo batch: apaga tabela inteira antes
        con.execute("DELETE FROM estudo_demanda_otima")
        con.commit()
    if resultado.empty:  # nada a gravar
        return
    resultado.to_sql("estudo_demanda_otima", con, if_exists="append", index=False)


def processar(
    banco: Path | str = _BANCO_PADRAO,
    *,
    codigo_uc: str | None = None,
    tarifas: TariffManager = COSERN_PLANILHA_V32,
) -> pd.DataFrame:
    banco_path = Path(banco)  # normaliza caminho do SQLite
    con = conectar(banco_path)  # conexão com migração
    try:
        hist = ler_historico(con, codigo_uc)  # meses elegíveis ao estudo
        if hist.empty:  # sem dados de demanda
            logger.warning("Nenhum mês Grupo A com demanda medida.")
            gravar_estudo(con, hist, substituir=codigo_uc is None)  # limpa tudo se batch vazio
            return hist  # DataFrame vazio

        partes: list[dict[str, Any]] = []  # resultados de todas as UCs
        # Agrupa por UC e roda estudo mês a mês para cada uma.
        for _uc, bloco in hist.groupby("codigo_uc", sort=False):
            partes.extend(estudar_historico_uc(bloco, tarifas))
        out = pd.DataFrame(partes)  # tabular para SQLite/API
        colunas = [  # ordem fixa das colunas gravadas
            "codigo_uc",
            "ano",
            "mes",
            "n_meses_janela",
            "demanda_otima_verde",
            "custo_demanda_verde",
            "custo_energia_verde",
            "custo_total_verde",
            "demanda_otima_azul_p",
            "demanda_otima_azul_fp",
            "custo_demanda_azul_p",
            "custo_demanda_azul_fp",
            "custo_energia_azul",
            "custo_total_azul",
            "modalidade_recomendada",
        ]
        out = out[colunas]  # descarta colunas extras se houver
        if codigo_uc:  # recálculo de uma UC: remove linhas antigas só dessa UC
            con.execute(
                "DELETE FROM estudo_demanda_otima WHERE codigo_uc = ?",
                (codigo_uc,),
            )
            con.commit()
            gravar_estudo(con, out, substituir=False)  # append sem apagar outras UCs
        else:
            gravar_estudo(con, out, substituir=True)  # batch: substitui tabela inteira
        logger.info(
            "estudo_demanda_otima: %s linha(s) em %s",
            len(out),
            banco_path,
        )
        return out  # DataFrame com o estudo calculado
    finally:
        con.close()  # sempre fecha a conexão


def main() -> int:
    logging.basicConfig(  # configura log no terminal quando rodado como script
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )
    parser = argparse.ArgumentParser(
        description="Módulo 3 — demanda ótima Verde/Azul sobre o SQLite do Módulo 2."
    )
    parser.add_argument("--banco", type=Path, default=_BANCO_PADRAO)  # caminho do SQLite
    parser.add_argument("--uc", dest="codigo_uc", default=None, help="Processa só esta UC.")
    args = parser.parse_args()  # lê argv
    try:
        out = processar(args.banco, codigo_uc=args.codigo_uc)  # executa pipeline
    except FileNotFoundError as exc:
        print(exc, file=sys.stderr)  # banco ausente
        return 1  # código de erro
    print()
    print("Estudo concluído.")
    print(f"  linhas: {len(out)}")
    if not out.empty:
        print(f"  UCs: {out['codigo_uc'].nunique()}")
        print(out["modalidade_recomendada"].value_counts().to_string())
    return 0  # sucesso


if __name__ == "__main__":  # entrada: python -m modulo3.engine
    raise SystemExit(main())  # repassa código de saída ao SO
