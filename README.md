# Generador de planning de guàrdies — Paquet v2

Sistema automatitzat per generar el planning mensual de guàrdies seguint les regles definides en una plantilla pujada per l'usuari. **No conté noms ni dades identificatives al codi.**

## Estructura del paquet

| Fitxer | Per a què |
|---|---|
| `01_Regles_Procés.docx` | Manual abstracte del procés (com funciona l'algorisme) |
| `02_Plantilla_Entrada.xlsx` | Plantilla buida amb les pestanyes Configuració, Radiòlegs i Vacances |
| `03_Plantilla_Planning.xlsx` | Plantilla buida del planning de sortida |
| `04_Prompt_Copilot.md` | Fallback per utilitzar amb Copilot (no recomanat com a principal) |
| `05_Benchmark_Exemple.xlsx` | Planning generat amb dades fictícies (per a testing) |
| `06_Proces_Altes_Baixes.md` | Procediment per altes i baixes de professionals |
| `EXEMPLE_Entrada_Mes_Demo.xlsx` | Entrada d'exemple amb noms ficticis (per testing) |
| `HANDOFF_VSCODE.md` | Guia per a desplegar a Streamlit Cloud des de VSCode |
| `app/` | Codi font de l'aplicació (Streamlit + scripts CLI) |
| `streamlit_app.py` | Wrapper a l'arrel per Streamlit Cloud |
| `requirements.txt` | Dependències Python |

## Com es fa servir

1. **Una vegada (setup inicial)**: omple la pestanya "Configuració" i "Radiòlegs" de `02_Plantilla_Entrada.xlsx` amb la teva plantilla real de professionals i els seus rols. Aquest fitxer és reutilitzable.
2. **Cada mes**: 
   - Còpia el fitxer 02 amb el nom del mes (p. ex. `Entrada_Juny_2026.xlsx`).
   - Revisa la **columna E "Actiu aquest mes"** de la pestanya Radiòlegs (vegeu més avall).
   - Empla la pestanya "Vacances" amb V/B/C/G/X de cada professional per al mes.
   - Indica al capçal el mes/any, primer radiòleg de la roda i festius.
   - Puja el fitxer a l'app web (o passa-ho per l'script CLI).
   - Descarrega el planning generat.

## Columna E — "Actiu aquest mes" (Radiòlegs)

Serveix per treure algú de la roda **només aquest mes**, sense haver d'esborrar ni tornar a afegir la seva fila:

- `Sí` (o buit) — participa normalment.
- `No` — queda fora de la roda, dels nou-incorporats i dels perfils especials. Si té un torn fix assignat a la pestanya Configuració, el planning hi escriurà `[SUBSTITUIR — Nom: NO actiu aquest mes]`.

El mes següent n'hi ha prou amb tornar-ho a posar a `Sí`. La plantilla és reutilitzable i la llista mestra de professionals no es toca mai.

## Codis de la pestanya Vacances

| Codi | Significat | Efecte sobre l'algorisme |
|---|---|---|
| `V` | Vacances | Bloqueja el dia **+ 2 dies de marge de reincorporació** |
| `B` | Baixa | Idèntic a `V` |
| `C` | Congrés | Bloqueja només el dia |
| `G` | Guàrdia en una altra institució | Bloqueja el dia, l'anterior i el posterior |
| `X` | Guàrdia ja assignada | Informatiu; no afecta l'algorisme |
| (buit) | Disponible | — |

### Marge de reincorporació

Quan una absència `V` o `B` acaba el dia **D**, el professional torna a entrar a la roda el dia **D+3**: es bloquegen D+1 i D+2 perquè tingui temps de reincorporar-se. **No perd el torn** — es manté al davant de la cua fins que estigui disponible, igual que abans.

El valor és un únic paràmetre a `app/planning_generator.py`:

```python
DIES_REINCORPORACIO = 3   # 1 = comportament antic (disponible l'endemà)
```

`C` i `G` no generen marge: `C` bloqueja només el dia i `G` ja bloqueja el dia anterior i el posterior.

## Privacitat

Tot el codi és genèric. Els noms reals dels professionals viuen exclusivament al fitxer d'entrada que la tècnica puja. Si l'app es desplega a un repositori públic (com Streamlit Community Cloud), no s'exposa cap dada identificativa.

**Regla dura: cap fitxer amb noms o correus reals no entra mai al repositori.** El `.gitignore` bloqueja `Plantilla_TD_IDI_REAL*.xlsx`, els plannings generats i els llistats de correus. Aquests fitxers viuen només al OneDrive de l'IDI i es pugen a l'app manualment cada mes. Abans de qualsevol `git push`, comproveu-ho:

```bash
git diff --stat origin/main..HEAD     # cap fitxer amb dades reals
git ls-files | grep -i "REAL\|Correus"  # ha de sortir buit
```

Els únics fitxers de dades versionats són `EXEMPLE_Entrada_Mes_Demo.xlsx` i `05_Benchmark_Exemple.xlsx`, amb noms ficticis (Alpha, Beta, Gamma…).

## Diferències respecte v1

La versió 1 tenia els noms hard-coded al codi i a les plantilles. La v2 és config-driven: el codi no sap qui són els professionals fins que llegeix el fitxer pujat. Permet aplicar el sistema a qualsevol organització amb la mateixa estructura de torns sense modificar codi.

## Validació

Vegeu `05_Benchmark_Exemple.xlsx` per a un planning generat amb dades fictícies. Per validar que el desplegament funciona, puja `EXEMPLE_Entrada_Mes_Demo.xlsx` a l'app i compara amb el benchmark.

