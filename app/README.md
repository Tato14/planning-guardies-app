# Codi font de l'aplicació

| Fitxer | Descripció |
|---|---|
| `app.py` | Interfície Streamlit principal |
| `planning_generator.py` | Algorisme central (importable + CLI) |
| `validator.py` | Validador independent (CLI) |
| `rotation_tracker.py` | Detector de l'últim de la roda (CLI) |
| `requirements.txt` | Dependències Python |
| `Dockerfile` | Per desplegar en contenidor |
| `DESPLEGAMENT.md` | Opcions de desplegament |

## Execució local

```bash
pip install -r requirements.txt
streamlit run app.py
```

