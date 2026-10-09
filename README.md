# Módulo 3 — motor de demanda ótima

Traduz `CalcularCusto` / `CalcularCustoDemanda` da planilha legado para Pandas/NumPy.
**Não** lê JSON do Módulo 1 e **não** sobe HTTP — só o SQLite do Módulo 2.

Histórico de consumo e demanda: view **`v_consumo_a_efetivo`** (PDF com prioridade). **O estudo de demanda ótima usa somente competências com demanda na fatura PDF** — meses só da API (medida frequentemente igual ao contrato) não entram na janela.

Só **Grupo A** (`classe_tarifaria = 'A'`). B optante fica de fora.

## O que calcula

Para cada UC e cada mês `t`, janela dos **12 meses civis** terminando em `t` (ex.: em 04/2025 entram 05/2024 … 04/2025). Meses sem histórico na base são omitidos; `n_meses_janela` indica quantos entraram (1–12).

1. Busca o kW inteiro de menor **custo de demanda** (fórmula VBA, sem reativo).
2. Soma TUSD+TE de energia (ponta e FP) para obter o **custo total** Verde e Azul.
3. `modalidade_recomendada`: `VERDE` se custo Verde &lt; Azul; empate vai para `AZUL`.

Verde usa `max(demanda P, demanda FP)` no mês. Azul otimiza P e FP à parte.

Grid de kW: de `Int(max(30, min da janela))` até `Fix(max da janela)+1` (como o VBA).

Tarifas iniciais: números da aba Variáveis da planilha v3.2 (`modulo3/tarifas.py`).
Azul FP lê o bloco Azul FP, **não** a célula Verde.

## Como rodar

```bash
cd Modulo1_API
pip install -r modulo3/requirements.txt
python -m modulo3.engine
python -m modulo3.engine --uc <CODIGO_DA_UC>
```

Grava `estudo_demanda_otima` no mesmo arquivo `modulo2/data/neoenergia.sqlite`.
Carga completa **substitui** a tabela. Depois de `python -m modulo2.tratamento`, rode o motor de novo.

Conferência contra a aba RESULTADOS da planilha v3.2:

```bash
python -m modulo3.validar_planilha
```

A API do Módulo 2 dispara o mesmo motor no POST `/api/unidades/{id}/estudo` (rode o uvicorn **na pasta pai**).

## Tabela `estudo_demanda_otima`

| Coluna | Significado |
|---|---|
| `codigo_uc`, `ano`, `mes` | Competência estudada (mês `t`) |
| `n_meses_janela` | 1–12 |
| `demanda_otima_verde` | kW único Verde |
| `custo_demanda_verde` / `custo_energia_verde` / `custo_total_verde` | R$ na janela |
| `demanda_otima_azul_p` / `demanda_otima_azul_fp` | kW Azul |
| `custo_demanda_azul_p` / `custo_demanda_azul_fp` / `custo_energia_azul` / `custo_total_azul` | R$ |
| `modalidade_recomendada` | `VERDE` ou `AZUL` |
