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
   - Empla la pestanya "Vacances" amb V/C/G/X de cada professional per al mes.
   - Indica al capçal el mes/any, primer radiòleg de la roda i festius.
   - Puja el fitxer a l'app web (o passa-ho per l'script CLI).
   - Descarrega el planning generat.

## Privacitat

Tot el codi és genèric. Els noms reals dels professionals viuen exclusivament al fitxer d'entrada que la tècnica puja. Si l'app es desplega a un repositori públic (com Streamlit Community Cloud), no s'exposa cap dada identificativa.

## Diferències respecte v1

La versió 1 tenia els noms hard-coded al codi i a les plantilles. La v2 és config-driven: el codi no sap qui són els professionals fins que llegeix el fitxer pujat. Permet aplicar el sistema a qualsevol organització amb la mateixa estructura de torns sense modificar codi.

## Validació

Vegeu `05_Benchmark_Exemple.xlsx` per a un planning generat amb dades fictícies. Per validar que el desplegament funciona, puja `EXEMPLE_Entrada_Mes_Demo.xlsx` a l'app i compara amb el benchmark.

