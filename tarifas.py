"""Tarifas Cosern Grupo A extraídas da planilha DEMANDA ÓTIMA_V3.2 (aba Variáveis).

Não replica o bug do VBA que lia B2 (Verde) no lugar de B9 (Azul FP).
Nesta planilha B2 = B9, mas os campos ficam separados para o próximo ciclo tarifário.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class TarifaDemanda:
    """TUSD demanda (R$/kW): normal, liminar (ociosa) e ultrapassagem."""

    normal: float
    liminar: float
    ultrapassagem: float


@dataclass(frozen=True, slots=True)
class TarifaEnergia:
    """TUSD + TE de energia (R$/kWh), ponta e fora ponta."""

    tusd_ponta: float
    te_ponta: float
    tusd_fp: float
    te_fp: float

    @property
    def ponta(self) -> float:
        return self.tusd_ponta + self.te_ponta

    @property
    def fora_ponta(self) -> float:
        return self.tusd_fp + self.te_fp


@dataclass(frozen=True, slots=True)
class TariffManager:
    verde_demanda: TarifaDemanda
    azul_ponta_demanda: TarifaDemanda
    azul_fp_demanda: TarifaDemanda
    verde_energia: TarifaEnergia
    azul_energia: TarifaEnergia
    ciclo: str = "planilha_v3.2"


# Valores literais da planilha (não arredondar). Ultrapassagem = 2 * TD.
COSERN_PLANILHA_V32 = TariffManager(
    ciclo="Cosern — DEMANDA ÓTIMA_V3.2 Variáveis",
    verde_demanda=TarifaDemanda(
        normal=40.247502849999996,
        liminar=33.00295234,
        ultrapassagem=80.495005699999993,
    ),
    azul_ponta_demanda=TarifaDemanda(
        normal=96.079661759999993,
        liminar=78.785322640000004,
        ultrapassagem=192.15932351999999,
    ),
    azul_fp_demanda=TarifaDemanda(
        normal=40.247502849999996,
        liminar=33.00295234,
        ultrapassagem=80.495005699999993,
    ),
    verde_energia=TarifaEnergia(
        tusd_ponta=2.4298690399999998,
        te_ponta=0.68563485000000002,
        tusd_fp=9.3893690000000002e-2,
        te_fp=0.43042967999999998,
    ),
    azul_energia=TarifaEnergia(
        tusd_ponta=9.3893690000000002e-2,
        te_ponta=0.68563485000000002,
        tusd_fp=9.3893690000000002e-2,
        te_fp=0.43042967999999998,
    ),
)
