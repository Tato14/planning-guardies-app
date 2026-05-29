"""Wrapper a l'arrel perquè Streamlit Cloud trobi l'app sense configuració."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'app'))
exec(open(Path(__file__).parent / 'app' / 'app.py').read())
