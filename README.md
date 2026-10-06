# Generador de planning de guàrdies — Paquet v3 (setmanes completes)

Sistema automatitzat per generar el planning mensual de guàrdies seguint les regles definides en una plantilla pujada per l'usuari. **No conté noms ni dades identificatives al codi.**

## El període: setmanes completes

Cada planning mensual està format per **setmanes completes, de dilluns a diumenge**. Una setmana pertany al mes en què cau el seu **dilluns**:

| Mes | Període | Setmanes |
|---|---|---|
| Novembre 2026 | **dg 01/11** → dg 06/12/2026 | 5 + dia pont (transició des del mes natural) |
| Desembre 2026 | dl 07/12/2026 → dg 03/01/2027 | 4 |
| Gener 2027 | dl 04/01 → dg 31/01/2027 | 4 |
| Març 2027 | dl 01/03 → dg 04/04/2027 | 5 |

Així cada setmana sencera és dins d'un sol fitxer i els enviaments setmanals no queden partits entre dos plannings. Cap planning inclou dies del mes anterior; pot acabar els primers dies del mes següent.

Detalls que convé saber:

- **Novembre 2026 és el mes de transició.** L'octubre es va fer per mes natural i acaba dissabte 31/10, així que el de novembre comença diumenge 1/11 (dia pont).
- **El fix del 1r dimecres** segueix sent el 1r dimecres del **mes natural**. Per exemple, el dc 02/12/2026 cau dins del planning de novembre.
- **Columnes grises de context.** La pestanya Vacances comença amb els 2 últims dies del període anterior. S'hi marca V o B només si l'absència ja venia d'abans; així el marge de reincorporació també s'aplica al començament del període.

## Estructura del paquet

| Fitxer | Per a què |
|---|---|
| `01_Regles_Procés.docx` | Manual abstracte del procés (com funciona l'algorisme) |
| `02_Plantilla_Entrada.xlsx` | Plantilla mestra buida (Configuració, Radiòlegs i Vacances en format setmanal) |
| `03_Plantilla_Planning.xlsx` | Plantilla genèrica del planning de sortida (l'app hi crea les pestanyes setmanals que calguin) |
| `Plantilles mensuals/` | Entrada i planning buit de cada mes, de nov 2026 a des 2027 (**local, no va a git: conté noms reals**) |
| `app/period_templates.py` | Genera les plantilles de qualsevol mes (també ho fa l'app) |
| `04_Prompt_Copilot.md` | Fallback per utilitzar amb Copilot (no recomanat com a principal) |
| `05_Benchmark_Exemple.xlsx` | Planning generat amb dades fictícies (per a testing) |
| `06_Proces_Altes_Baixes.md` | Procediment per altes i baixes de professionals |
| `EXEMPLE_Entrada_Mes_Demo.xlsx` | Entrada d'exemple amb noms ficticis (per testing) |
| `HANDOFF_VSCODE.md` | Guia per a desplegar a Streamlit Cloud des de VSCode |
| `app/` | Codi font de l'aplicació (Streamlit + scripts CLI) |
| `streamlit_app.py` | Wrapper a l'arrel per Streamlit Cloud |
| `requirements.txt` | Dependències Python |

## Com es fa servir

1. **Cada mes**, agafa l'entrada del mes. Pot venir de dos llocs:
   - la que l'app et deixa descarregar en generar el mes anterior (**recomanat**), que ja porta els radiòlegs i l'Actiu al dia i el primer de la roda calculat;
   - o la de la carpeta `Plantilles mensuals/`, generada amb la llista de professionals de novembre 2026.
2. Revisa la **columna E "Actiu aquest mes"** de la pestanya Radiòlegs (vegeu més avall).
3. A la pestanya Vacances, marca V/B/C/G/X de cada professional. Els festius oficials de Catalunya ja hi són marcats amb F; afegeix-ne o treu-ne si cal.
4. Comprova el primer radiòleg de la roda (fila 1).
5. Puja el fitxer a l'app juntament amb el planning buit (el del mes o el genèric) i genera.
6. Descarrega el planning (`Planning_AAAA-MM.xlsx`) i l'entrada del mes següent.

Per crear l'entrada d'un mes qualsevol, l'app té l'opció *Preparar la plantilla d'entrada d'un altre mes*. També es pot fer per línia d'ordres:

```bash
python3 app/period_templates.py Entrada_mestra.xlsx 03_Plantilla_Planning.xlsx "Plantilles mensuals" 2028-01 2028-12
```

Els festius de 2026 i 2027 són els oficials del DOGC. Per als anys que encara no tenen calendari publicat es marquen els 14 festius habituals, i la pestanya Instruccions ho avisa.

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

Els únics fitxers de dades versionats són `EXEMPLE_Entrada_Mes_Demo.xlsx` i `05_Benchmark_Exemple.xlsx`, amb noms ficticis (Alpha, Beta, Gamma…). La carpeta `Plantilles mensuals/` també queda fora de git.

## Diferències respecte v1

La versió 1 tenia els noms hard-coded al codi i a les plantilles. La v2 és config-driven: el codi no sap qui són els professionals fins que llegeix el fitxer pujat. Permet aplicar el sistema a qualsevol organització amb la mateixa estructura de torns sense modificar codi.

## Validació

Vegeu `05_Benchmark_Exemple.xlsx` per a un planning generat amb dades fictícies (juny 2026 per setmanes: dl 01/06 → dg 05/07). Per validar que el desplegament funciona, puja `EXEMPLE_Entrada_Mes_Demo.xlsx` a l'app i compara amb el benchmark.

Els fitxers d'entrada en format antic (columnes 1-31 del mes natural) es continuen acceptant, però el planning resultant no va per setmanes completes. L'app ho avisa.

## Compatibilitat amb els fluxos de Microsoft 365

L'Office Script no canvia. Busca les pestanyes que contenen «Setmana», llegeix la data de la fila 3 i treu el mes de la primera guàrdia, que ara és sempre dins del mes del planning. Els dies fora del període surten en gris amb text entre parèntesis, que l'script ignora.

Els fluxos 3 i 4 continuen llegint un únic fitxer (`Actual/planning.xlsx`). Al voltant del canvi de període, la setmana que anuncien pot pertànyer al planning següent. Mentre no s'adaptin perquè triïn el fitxer de cada setmana, cal vigilar-ho a mà.

