"""Confere o núcleo contra os totais da aba RESULTADOS da planilha v3.2.

    python -m modulo3.validar_planilha
"""
from __future__ import annotations

import numpy as np

from .engine import estudar_janela
from .tarifas import COSERN_PLANILHA_V32

P = np.array([30, 30, 30, 30, 35, 35, 31, 31, 31, 31, 31, 31], dtype=float)
FP = P.copy()
CONS_P = np.array(
    [
        13097.13,
        14199.32,
        13317.44,
        14082.39,
        14786.67,
        13531.83,
        11912.17,
        13724.55,
        13107.72,
        12956.39,
        14617.7,
        12632.38,
    ]
)
CONS_FP = np.array(
    [
        148031.1,
        149121,
        144276.3,
        160952.4,
        150469.2,
        145826.1,
        126132.3,
        129798.9,
        130372.2,
        131002.2,
        135891,
        143035.2,
    ]
)

ESPERADO = {
    "demanda_otima_verde": 31,
    "custo_demanda_verde": 15909.032926560001,
    "demanda_otima_azul_p": 31,
    "custo_demanda_azul_p": 37978.368700479987,
    "demanda_otima_azul_fp": 31,
    "custo_demanda_azul_fp": 15909.032926560001,
    "custo_energia_verde": 1393284.559209157,
    "custo_energia_azul": 1014936.6998234155,
    "custo_total_verde": 1409193.5921357169,
    "custo_total_azul": 1068824.1014504554,
    "modalidade_recomendada": "AZUL",
}


def main() -> None:
    r = estudar_janela(P, FP, CONS_P, CONS_FP, COSERN_PLANILHA_V32)
    ok = True
    for k, exp in ESPERADO.items():
        obt = r[k]
        bate = obt == exp if isinstance(exp, str) else abs(float(obt) - exp) < 1e-4
        if not bate:
            ok = False
        print(f"  {'OK' if bate else 'FALHOU'} {k}: obtido={obt} esperado={exp}")
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    raise SystemExit(main())
