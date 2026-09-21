# Prompt per a Microsoft Copilot — Fallback

Aquest prompt és per usar Copilot com a alternativa a l'app Streamlit. **Recomanació**: feu servir l'app, no Copilot, per fiabilitat. Aquest fitxer queda com a contingència.

---

## Instruccions per a Copilot

Has de generar el planning mensual de guàrdies seguint les regles documentades. Et passo tres documents:

1. `01_Regles_Procés.docx` — Manual abstracte del procés. Llegeix-lo primer.
2. `<Entrada_Mes>.xlsx` — Fitxer d'entrada del mes amb pestanyes Configuració, Radiòlegs i Vacances.
3. `03_Plantilla_Planning.xlsx` — Plantilla buida del planning.

Has d'omplir la plantilla del planning seguint el procés:

### Pas 1 — Llegir el context
Llegeix el document de regles per entendre l'algorisme. Identifica els codis V/B/C/G/X i els rols (rotador, fix-només, fix-i-rota, sun-nit-only, weekend-day-only, nou-incorporat).

### Pas 2 — Llegir les dades
Obre el fitxer d'entrada:
- Pestanya "Vacances", capçal: mes/any (B1), primer radiòleg (D1), festius (F1).
- Pestanya "Configuració": qui cobreix cada torn fix i quins són els perfils especials.
- Pestanya "Radiòlegs": llista mestra amb rol (col. B) i **Actiu aquest mes (col. E)** de cadascú.
- Pestanya "Vacances", files 4 en endavant: V/B/C/G/X per dia/professional.

### Pas 2 bis — Filtrar els NO actius
Tot professional amb `No` a la columna E de la pestanya Radiòlegs **no fa cap guàrdia aquest mes**: exclou-lo de la roda, dels nou-incorporats i dels perfils especials. Si té un torn fix a Configuració, escriu-hi `[SUBSTITUIR — Nom: NO actiu aquest mes]`. Una cel·la buida equival a `Sí`.

### Pas 2 ter — Calcular el marge de reincorporació
Per a cada professional, si una absència `V` o `B` acaba el dia D (és a dir, D té V o B i D+1 no), **bloqueja també els dies D+1 i D+2**. Torna a estar disponible el dia D+3. `C` i `G` no generen marge.

### Pas 3 — Pre-omplir els fixos
Per a cada dia laborable no festiu, copia el radiòleg fix indicat a Configuració al torn corresponent del planning.

### Pas 4 — Generar la roda
- Crea la llista de rotadors (rotador + fix-i-rota + sun-nit-only + weekend-day-only), **excloent-ne els NO actius**.
- Ordena alfabèticament ignorant accents.
- Rota la cua perquè comenci pel primer radiòleg indicat a D1.

### Pas 5 — Assignar torns
Per a cada torn no fix, en ordre cronològic:
- Si és diumenge nit i hi ha pendent un sun-nit-only, assigna-l'hi.
- Si és Sat/Sun dia i hi ha pendent un weekend-day-only, assigna-l'hi.
- Altrament, pren el primer de la cua. Si és sun-nit-only o weekend-day-only, defèr-lo. Si pot treballar aquell dia, assigna'l. Si no, prova el següent.

### Pas 6 — Validar
- Cap doble assignació per dia.
- Cap V/B/C/G violat, ni cap dia de marge de reincorporació.
- Cap professional amb Actiu = "No" al planning.
- Sun-nit-only només a Dg nit.
- Weekend-day-only només a Ds/Dg dia.
- Distribució equitativa entre rotadors purs.

### Pas 7 — Entregar
Retorna l'Excel omplert i un informe breu (≤200 paraules) amb distribució de guàrdies i incongruències detectades.

