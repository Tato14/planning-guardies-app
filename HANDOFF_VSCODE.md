# Desplegament a Streamlit Cloud des de VSCode

Aquesta guia t'acompanya des de zero fins a tenir l'app accessible per URL pública. Temps estimat: 45-60 minuts.

## Prerequisits

1. **VSCode** + **Git** instal·lats al PC.
2. **Compte de GitHub** (https://github.com/signup).
3. **Compte de Streamlit Community Cloud** (es crea durant el procés amb GitHub).

## Pas 1 — Obrir el paquet a VSCode

`File → Open Folder…` i selecciona aquesta carpeta.

## Pas 2 — Configurar Git

A `Terminal → New Terminal`, executa:

```bash
git config --global user.name "El teu nom"
git config --global user.email "tu@correu.com"
```

## Pas 3 — Iniciar sessió GitHub a VSCode

Panell d'identitat (icona de persona a baix) → "Sign in with GitHub". Autoritza al navegador.

## Pas 4 — Init repo + commit + push

1. Panel `Source Control` (icona de branques) → clica `Initialize Repository`.
2. Escriu missatge de commit (p. ex. "Versió inicial").
3. Clica el botó ✓.
4. Clica `Publish Branch` → nom suggerit `planning-guardies-app` → **PÚBLIC** (Streamlit Community Cloud gratuït necessita repo públic; el codi no conté dades sensibles).

## Pas 5 — Desplegar a Streamlit Cloud

1. https://share.streamlit.io → login amb GitHub.
2. `New app` → repo `planning-guardies-app`, branca `main`.
3. **Main file path**: `streamlit_app.py`.
4. Advanced → Python version `3.11`.
5. Deploy.

Triga 2-4 minuts. Quan acabi, tindràs una URL pública.

## Pas 6 — Restringir accés (opcional)

Settings → Sharing → Restricted by email → afegir el teu domini corporatiu.

## Pas 7 — Validar

Obre l'URL. Puja `EXEMPLE_Entrada_Mes_Demo.xlsx` i `03_Plantilla_Planning.xlsx`. Genera el planning i compara amb `05_Benchmark_Exemple.xlsx`. Si coincideixen, el desplegament és OK.

## Pas 8 — Anar a producció

Quan estiguis llest per fer-ho servir amb dades reals, no cal canviar res al codi. La tècnica només ha d'omplir un fitxer d'entrada amb les seves dades reals i pujar-lo a la mateixa URL.

## Per a canvis futurs

Edita el codi a VSCode → commit + push → Streamlit Cloud redesplega automàticament.

## Resolució de problemes

- **Error "ModuleNotFoundError"**: verifica `requirements.txt` a l'arrel.
- **App no carrega**: revisa logs a Streamlit Cloud Console.
- **Push falla**: re-autentica GitHub a VSCode.

