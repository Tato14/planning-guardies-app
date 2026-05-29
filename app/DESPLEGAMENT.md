# Desplegament — Opcions

| Opció | Cost | Temps | Per a qui |
|---|---|---|---|
| A. Streamlit Cloud (gratuït) | 0 € | 30 min | Volem gratuït i fàcil; repo públic OK |
| B. Servidor propi Linux | Cost servidor | 2-3 h | Volem allotjar-ho internament |
| C. Docker (Azure/AWS) | Cost cloud | 1-2 h | Cluster ja existent |
| D. Local en un PC | 0 € | 15 min | Només la tècnica el necessita |

Per a la majoria de casos: opció A. Vegeu `HANDOFF_VSCODE.md` per a la guia pas a pas.

## Ús des de línia de comandes

```bash
# Generar planning
python planning_generator.py <entrada.xlsx> <plantilla.xlsx> <sortida.xlsx>

# Validar un planning generat
python validator.py <entrada.xlsx> <planning.xlsx>

# Detectar últim de la roda del mes anterior
python rotation_tracker.py <planning_mes_anterior.xlsx> <entrada_mes_nou.xlsx>
```

## Dependències

`pip install -r requirements.txt`

Només cal `streamlit` i `openpyxl`.

