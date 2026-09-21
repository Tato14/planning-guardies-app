"""Wrapper a l'arrel perquè Streamlit Cloud trobi l'app sense configuració.

Important: Streamlit Cloud reutilitza el procés de Python entre execucions i,
després d'un desplegament, pot conservar a sys.modules la versió ANTIGA dels
mòduls locals. Si el nou app.py importa un nom que el mòdul antic no tenia,
salta un ImportError encara que el codi del repositori sigui correcte.
Per evitar-ho, purguem els mòduls propis abans d'executar l'app: així sempre
es carreguen del disc.
"""
import sys
from pathlib import Path

APP_DIR = Path(__file__).parent / 'app'
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

# Descarrega les versions en memòria dels mòduls propis (no de les llibreries).
for _m in ('planning_generator', 'validator', 'rotation_tracker'):
    sys.modules.pop(_m, None)

exec(open(APP_DIR / 'app.py', encoding='utf-8').read())
