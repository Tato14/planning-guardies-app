"""
Generador de planning de guàrdies — versió config-driven per SETMANES COMPLETES.

Sense noms hard-coded al codi. Tota la informació identificativa
(noms, rols, fixos) prové de la plantilla pujada per l'usuari.

PERÍODE DEL PLANNING
  Cada planning mensual està format per setmanes completes (dilluns → diumenge).
  Una setmana pertany al mes en què cau el seu DILLUNS:
      Novembre 2026 = dl 02/11/2026 → dg 06/12/2026 (5 setmanes)
      Desembre 2026 = dl 07/12/2026 → dg 03/01/2027 (4 setmanes)
  Així cada setmana sencera és dins d'un sol planning i els enviaments
  setmanals no queden partits entre dos fitxers.

Estructura esperada del fitxer d'entrada (xlsx):
  - Pestanya "Configuració": rols fixos + perfils especials
  - Pestanya "Radiòlegs": llista mestra amb rol (col. B) i Actiu Sí/No (col. E)
  - Pestanya "Vacances": capçalera del període + una columna per data,
    fila "Festiu" (F) i codis V/B/C/G/X per professional i dia.

Es manté la compatibilitat amb el format antic (columnes 1..31 del mes natural).
Vegeu period_templates.py per generar les plantilles de cada mes.
"""
from __future__ import annotations

import calendar
import datetime
import re
import unicodedata
from collections import defaultdict, deque
from dataclasses import dataclass, field

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
# Constants
# ============================================================

DOW_MAP = {
    'dilluns': 0, 'dimarts': 1, 'dimecres': 2, 'dijous': 3,
    'divendres': 4, 'dissabte': 5, 'diumenge': 6,
    'lunes': 0, 'martes': 1, 'miércoles': 2, 'jueves': 3,
    'viernes': 4, 'sábado': 5, 'domingo': 6,
}

MESOS_CA = ['gener', 'febrer', 'març', 'abril', 'maig', 'juny', 'juliol',
            'agost', 'setembre', 'octubre', 'novembre', 'desembre']
MESOS_CA_ABREV = ['gen', 'febr', 'març', 'abr', 'maig', 'juny', 'jul',
                  'ag', 'set', 'oct', 'nov', 'des']
DIES_CA = ['dl', 'dt', 'dc', 'dj', 'dv', 'ds', 'dg']

ROLS_VALIDS = ('rotador', 'fix-només', 'fix-nomes', 'fix-i-rota', 'sun-nit-only',
               'weekend-day-only', 'nou-incorporat', 'nou incorporat')

# ------------------------------------------------------------
# Codis de la pestanya "Vacances"
# ------------------------------------------------------------
# Codis que l'usuari pot escriure a les cel·les:
CODIS_FULL = ('V', 'B', 'C', 'G', 'X')
#   V = Vacances       -> bloqueja el dia
#   B = Baixa          -> bloqueja el dia (mateix comportament que V)
#   C = Congrés        -> bloqueja el dia
#   G = Guàrdia externa-> bloqueja el dia, l'anterior i el posterior
#   X = Informatiu (guàrdia ja assignada); no afecta l'algorisme

# Codis derivats, calculats pel generador (mai escrits al full per l'usuari):
CODI_INACTIU = 'I'          # Radiòlegs!E = "No" -> no fa guàrdies aquest mes
CODI_REINCORPORACIO = 'R'   # dies de marge després d'una absència V/B

# Absències llargues que generen marge de reincorporació.
CODIS_AMB_REINCORPORACIO = ('V', 'B')

# Decalatge de reincorporació, en dies.
#   1 = comportament antic (disponible l'endemà de tornar)
#   3 = disponible al cap de 3 dies; per tant es bloquegen els 2 dies previs
DIES_REINCORPORACIO = 3

# Tots els codis que impedeixen assignar una guàrdia aquell dia.
CODIS_BLOQUEJANTS = ('V', 'B', 'C', 'G', CODI_INACTIU, CODI_REINCORPORACIO)

MOTIU_TEXT = {
    'V': 'vacances',
    'B': 'baixa',
    'C': 'congrés',
    'G': 'guàrdia externa',
    CODI_INACTIU: 'NO actiu aquest mes (Radiòlegs col. E = No)',
    CODI_REINCORPORACIO: f'marge de reincorporació ({DIES_REINCORPORACIO} dies)',
}

SHIFT_ROWS = {'8-16': (4, 9), '16-20': (11, 16), '20-08': (18, 23)}


# ============================================================
# Helpers
# ============================================================

def fold(s: str) -> str:
    if not s:
        return ''
    return unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().lower()


def _es_actiu(val) -> bool:
    """Interpreta la columna E (Actiu) de la pestanya Radiòlegs.

    Buit = actiu (retrocompatibilitat amb plantilles antigues sense la columna).
    Només 'No' / 'N' / 'FALSE' / '0' desactiven el professional.
    """
    if val is None:
        return True
    if isinstance(val, bool):
        return val
    s = fold(str(val)).strip()
    if s == '':
        return True
    return s not in ('no', 'n', 'fals', 'false', '0', 'nao', 'non')


def _es_fila_llegenda(name: str) -> bool:
    """Files de llegenda de la pestanya Radiòlegs que no són professionals."""
    f = fold(name).strip().rstrip(':').strip()
    return f.startswith('llegenda') or f in {fold(r) for r in ROLS_VALIDS}


def _parse_month(s) -> tuple[int, int]:
    if isinstance(s, (datetime.datetime, datetime.date)):
        return s.year, s.month
    s = str(s).strip()
    months = {
        'gener': 1, 'febrer': 2, 'març': 3, 'abril': 4, 'maig': 5, 'juny': 6,
        'juliol': 7, 'agost': 8, 'setembre': 9, 'octubre': 10, 'novembre': 11, 'desembre': 12,
        'enero': 1, 'febrero': 2, 'marzo': 3, 'mayo': 5, 'junio': 6,
        'julio': 7, 'septiembre': 9, 'noviembre': 11, 'diciembre': 12,
        'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
        'july': 7, 'august': 8, 'september': 9, 'october': 10, 'december': 12,
    }
    sl = s.lower()
    year = month = None
    m = re.search(r'(20\d{2})', s)
    if m:
        year = int(m.group(1))
    for k, v in months.items():
        if k in sl:
            month = v
            break
    if month is None:
        m = re.search(r'\b(0?[1-9]|1[0-2])[/\-](20\d{2})\b', s) or re.search(r'\b(20\d{2})-(0?[1-9]|1[0-2])\b', s)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            month, year = (a, b) if a <= 12 else (b, a)
    if year is None or month is None:
        raise ValueError(f"No s'ha pogut interpretar el mes: '{s}'")
    return year, month


def _parse_festius(s) -> list:
    if not s:
        return []
    parts = [p.strip() for p in str(s).replace(';', ',').split(',') if p.strip()]
    return [int(p) for p in parts if p.isdigit()]


def _as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        m = re.match(r'^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$', v)
        if m:
            try:
                return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
    return None


# ============================================================
# Períodes de setmanes completes
# ============================================================

def first_monday(year: int, month: int) -> datetime.date:
    d = datetime.date(year, month, 1)
    return d + datetime.timedelta(days=(7 - d.weekday()) % 7)


def next_month(year: int, month: int) -> tuple[int, int]:
    return (year + 1, 1) if month == 12 else (year, month + 1)


def prev_month(year: int, month: int) -> tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def period_for_month(year: int, month: int) -> tuple[datetime.date, datetime.date]:
    """Període del planning d'un mes: del 1r dilluns del mes al diumenge
    anterior al 1r dilluns del mes següent."""
    start = first_monday(year, month)
    ny, nm = next_month(year, month)
    end = first_monday(ny, nm) - datetime.timedelta(days=1)
    return start, end


def period_month_of(d: datetime.date) -> tuple[int, int]:
    """Mes al qual pertany la setmana que conté la data d (el mes del seu dilluns)."""
    monday = d - datetime.timedelta(days=d.weekday())
    return monday.year, monday.month


def month_label(year: int, month: int) -> str:
    return f"{MESOS_CA[month - 1].capitalize()} {year}"


def fmt_dia(d: datetime.date, with_year: bool = False) -> str:
    s = f"{DIES_CA[d.weekday()]} {d.day:02d}/{d.month:02d}"
    return f"{s}/{d.year}" if with_year else s


def week_sheet_name(monday: datetime.date) -> str:
    """Nom de pestanya: 'Setmana 49 (30 nov-06 des)'. Sempre conté 'Setmana'
    (l'Office Script de M365 busca aquesta paraula)."""
    sunday = monday + datetime.timedelta(days=6)
    iso = monday.isocalendar()[1]
    a = f"{monday.day:02d} {MESOS_CA_ABREV[monday.month - 1]}"
    b = f"{sunday.day:02d} {MESOS_CA_ABREV[sunday.month - 1]}"
    return f"Setmana {iso:02d} ({a}-{b})"


# ============================================================
# Data structures
# ============================================================

@dataclass
class Config:
    # Fix slots: (dow 0-6, shift '8-16'|'16-20', role 'N'|'B'|'N i B') -> name
    fix_slots: dict = field(default_factory=dict)
    # First-Wed-of-month special slot (16-20)
    first_wed_radiologist: str = ''
    # Lists per role
    rotators: list = field(default_factory=list)       # everyone in rotation
    fix_only: list = field(default_factory=list)       # excluded from rotation
    fix_and_rota: list = field(default_factory=list)   # both
    sun_nit_only: list = field(default_factory=list)
    weekend_day_only: list = field(default_factory=list)
    nou_incorporats: list = field(default_factory=list)  # cobreixen dimecres tarda (excepte 1r Dc del mes)
    inactius: list = field(default_factory=list)       # Radiòlegs!E = "No": fora del planning aquest mes


@dataclass
class Meta:
    """Període planificat. Internament els dies es numeren 1..N (índex dins del
    període); els índexs <= 0 són dies de context del període anterior."""
    year: int                                       # mes "etiqueta" del planning
    month: int
    festius: list = field(default_factory=list)     # índexs (1..N)
    start_radiologist: str = ''
    dates: list = field(default_factory=list)       # dates planificades, consecutives
    load_warnings: list = field(default_factory=list)
    input_format: str = 'setmanes'                  # 'setmanes' o 'mes' (format antic)

    def __post_init__(self):
        if not self.dates:   # compatibilitat: mes natural
            n = calendar.monthrange(self.year, self.month)[1]
            self.dates = [datetime.date(self.year, self.month, d) for d in range(1, n + 1)]

    # --- mida i extrems ---
    @property
    def n_days(self) -> int:
        return len(self.dates)

    @property
    def days_in_month(self) -> int:   # nom antic, es manté per compatibilitat
        return self.n_days

    @property
    def start(self) -> datetime.date:
        return self.dates[0]

    @property
    def end(self) -> datetime.date:
        return self.dates[-1]

    @property
    def label(self) -> str:
        return month_label(self.year, self.month)

    @property
    def code(self) -> str:
        return f"{self.year}-{self.month:02d}"

    # --- conversió índex <-> data ---
    def date_of(self, idx: int) -> datetime.date:
        return self.start + datetime.timedelta(days=idx - 1)

    def dow(self, idx: int) -> int:
        return self.date_of(idx).weekday()

    def index_of(self, d):
        d = _as_date(d)
        if d is None:
            return None
        idx = (d - self.start).days + 1
        return idx if 1 <= idx <= self.n_days else None

    def fmt(self, idx: int, with_year: bool = False) -> str:
        return fmt_dia(self.date_of(idx), with_year)

    def is_first_wed(self, idx: int) -> bool:
        """1r dimecres del MES NATURAL (la regla del fix no canvia amb les setmanes)."""
        d = self.date_of(idx)
        return d.weekday() == 2 and d.day <= 7

    # --- setmanes ---
    def week_mondays(self) -> list:
        m = self.start - datetime.timedelta(days=self.start.weekday())
        out = []
        while m <= self.end:
            out.append(m)
            m += datetime.timedelta(days=7)
        return out

    @property
    def n_weeks(self) -> int:
        return len(self.week_mondays())

    @property
    def festiu_dates(self) -> list:
        return [self.date_of(i) for i in sorted(self.festius)]

    def festius_text(self) -> str:
        return ', '.join(f"{d.day:02d}/{d.month:02d}" for d in self.festiu_dates)

    def set_festius_from_text(self, text: str) -> list:
        """Accepta 'dd/mm', 'dd/mm/aaaa' o 'dd' (si no és ambigu). Retorna avisos."""
        avisos, out = [], set()
        for tok in re.split(r'[,;\s]+', text or ''):
            tok = tok.strip()
            if not tok:
                continue
            m = re.match(r'^(\d{1,2})(?:/(\d{1,2}))?(?:/(\d{2,4}))?$', tok)
            if not m:
                avisos.append(f"Festiu no reconegut: «{tok}»")
                continue
            dd = int(m.group(1))
            mm = int(m.group(2)) if m.group(2) else None
            yy = int(m.group(3)) if m.group(3) else None
            if yy is not None and yy < 100:
                yy += 2000
            cands = [i for i in range(1, self.n_days + 1)
                     if self.date_of(i).day == dd
                     and (mm is None or self.date_of(i).month == mm)
                     and (yy is None or self.date_of(i).year == yy)]
            if len(cands) == 1:
                out.add(cands[0])
            elif not cands:
                avisos.append(f"El festiu «{tok}» no cau dins del període ({fmt_dia(self.start)} → {fmt_dia(self.end)})")
            else:
                avisos.append(f"El festiu «{tok}» és ambigu dins del període; escriu-lo com dd/mm")
        self.festius = sorted(out)
        return avisos


# ============================================================
# Loaders
# ============================================================

def _find_label_row(ws, prefix: str, max_row: int = 15):
    pf = fold(prefix)
    for r in range(1, min(ws.max_row, max_row) + 1):
        a = ws.cell(row=r, column=1).value
        if isinstance(a, str) and fold(a).strip().startswith(pf):
            return r
    return None


def _value_right_of(ws, row: int, prefix: str, max_col: int = 60):
    """Valor de la primera cel·la no buida a la dreta de l'etiqueta indicada."""
    pf = fold(prefix)
    last = min(ws.max_column, max_col)
    for c in range(1, last + 1):
        v = ws.cell(row=row, column=c).value
        if isinstance(v, str) and fold(v).strip().startswith(pf):
            for c2 in range(c + 1, last + 1):
                v2 = ws.cell(row=row, column=c2).value
                if v2 not in (None, ''):
                    return v2
            return None
    return None


def _majority_month(dates: list) -> tuple[int, int]:
    cnt = defaultdict(int)
    for d in dates:
        cnt[(d.year, d.month)] += 1
    return max(cnt.items(), key=lambda kv: kv[1])[0]


def _meta_setmanes(ws_v, header_row: int, date_cols: dict, warnings: list):
    """Format nou: capçalera amb dates reals (setmanes completes)."""
    raw_label = _value_right_of(ws_v, 1, 'mes del planning') or ws_v['B1'].value
    raw_start_rad = _value_right_of(ws_v, 1, 'primer radi') or ''
    raw_ini = _value_right_of(ws_v, 2, 'periode planificat')
    raw_fi = _value_right_of(ws_v, 2, 'fins al')

    all_dates = sorted(set(date_cols.values()))
    start = _as_date(raw_ini) or all_dates[0]
    end = _as_date(raw_fi) or all_dates[-1]
    if end < start:
        raise ValueError(f"El període està invertit: {fmt_dia(start, True)} → {fmt_dia(end, True)}")
    planned = [start + datetime.timedelta(days=i) for i in range((end - start).days + 1)]
    present = set(all_dates)
    missing = [d for d in planned if d not in present]
    if missing:
        raise ValueError("A la capçalera de la pestanya Vacances hi falten dates del període: "
                         + ', '.join(fmt_dia(d) for d in missing[:10]))

    try:
        year, month = _parse_month(raw_label) if raw_label else _majority_month(planned)
    except ValueError:
        year, month = _majority_month(planned)
        warnings.append(f"No s'ha entès el mes «{raw_label}»; s'ha pres {month_label(year, month)}.")

    festius = []
    fest_row = _find_label_row(ws_v, 'festiu')
    if fest_row:
        for col, d in date_cols.items():
            if start <= d <= end:
                v = ws_v.cell(row=fest_row, column=col).value
                if v is not None and str(v).strip():
                    festius.append((d - start).days + 1)

    col_index = {}
    for col, d in date_cols.items():
        idx = (d - start).days + 1
        if idx <= len(planned):          # idx <= 0: dies de context (període anterior)
            col_index[col] = idx

    if start.weekday() != 0:
        warnings.append(f"El període comença en {fmt_dia(start, True)}, no en dilluns: "
                        "s'entén com a dia pont amb el planning anterior.")
    if end.weekday() != 6:
        warnings.append(f"El període acaba en {fmt_dia(end, True)}, no en diumenge.")

    meta = Meta(year=year, month=month, festius=sorted(set(festius)),
                start_radiologist=str(raw_start_rad).strip(), dates=planned,
                input_format='setmanes')
    return meta, col_index


def _meta_mes(ws_v, int_cols: dict, warnings: list):
    """Format antic: B1 mes, D1 primer radiòleg, F1 festius, columnes 1..31."""
    raw_month = ws_v['B1'].value or ''
    raw_start = ws_v['D1'].value or ''
    raw_festius = ws_v['F1'].value or ''
    year, month = _parse_month(raw_month)
    n = calendar.monthrange(year, month)[1]
    festius = [d for d in _parse_festius(raw_festius) if 1 <= d <= n]
    if int_cols:
        col_index = {col: day for col, day in int_cols.items() if day <= n}
    else:
        col_index = {1 + d: d for d in range(1, n + 1)}
    warnings.append("Fitxer en format antic (mes natural, columnes 1-31): funciona, però el planning "
                    "NO anirà per setmanes completes. Feu servir les plantilles mensuals noves.")
    meta = Meta(year=year, month=month, festius=sorted(set(festius)),
                start_radiologist=str(raw_start).strip(), input_format='mes')
    return meta, col_index


def load_input(path) -> tuple[dict, Config, Meta]:
    """Read the input xlsx and return (constraints, config, meta)."""
    wb = openpyxl.load_workbook(path, data_only=True)

    # ----- Read Radiòlegs sheet -----
    if 'Radiòlegs' not in wb.sheetnames:
        raise ValueError("Falta la pestanya 'Radiòlegs' al fitxer d'entrada.")
    ws_r = wb['Radiòlegs']
    radiologists = {}   # name -> rol (tothom, actius i no actius)
    inactius = []       # noms amb Actiu = "No" aquest mes
    for row in range(4, ws_r.max_row + 1):
        name = ws_r.cell(row=row, column=1).value
        rol = ws_r.cell(row=row, column=2).value
        actiu = ws_r.cell(row=row, column=5).value   # columna E
        if not name or not isinstance(name, str):
            continue
        name = name.strip()
        if not name or _es_fila_llegenda(name):
            continue
        if rol:
            rol = str(rol).strip().lower()
        else:
            rol = 'rotador'  # default
        radiologists[name] = rol
        if not _es_actiu(actiu):
            inactius.append(name)

    # ----- Read Configuració sheet -----
    if 'Configuració' not in wb.sheetnames:
        raise ValueError("Falta la pestanya 'Configuració' al fitxer d'entrada.")
    ws_c = wb['Configuració']

    config = Config()

    # Walk down rows looking for fix slot rows (have a dow name in col A)
    for row in range(4, ws_c.max_row + 1):
        col_a = ws_c.cell(row=row, column=1).value
        col_b = ws_c.cell(row=row, column=2).value  # franja
        col_c = ws_c.cell(row=row, column=3).value  # rol
        col_d = ws_c.cell(row=row, column=4).value  # name

        if not col_a or not isinstance(col_a, str):
            continue
        a = col_a.strip().lower()
        b = (str(col_b).strip().lower() if col_b else '')
        c = (str(col_c).strip() if col_c else '')
        name = (str(col_d).strip() if col_d else '')

        if a in DOW_MAP and b in ('8-16', '16-20') and c and name:
            config.fix_slots[(DOW_MAP[a], b, c)] = name
        elif 'dimecres-1r' in a or '1r-dimecres' in a or 'primer dimecres' in a:
            if name:
                config.first_wed_radiologist = name
        elif a == 'sun-nit-only' and name:
            config.sun_nit_only.append(name)
        elif a == 'weekend-day-only' and name:
            config.weekend_day_only.append(name)

    # Els professionals marcats com a NO actius queden fora de tots els grups:
    # ni roda, ni nou-incorporats, ni perfils especials. No cal esborrar-los
    # de la plantilla: n'hi ha prou amb posar "No" a la columna E.
    config.inactius = list(inactius)
    inactius_set = set(inactius)
    # També cal treure'ls dels perfils especials llegits de la pestanya Configuració
    config.sun_nit_only = [n for n in config.sun_nit_only if n not in inactius_set]
    config.weekend_day_only = [n for n in config.weekend_day_only if n not in inactius_set]

    # Categorize radiologists by rol
    for name, rol in radiologists.items():
        if name in inactius_set:
            continue
        if rol == 'rotador':
            config.rotators.append(name)
        elif rol == 'fix-només' or rol == 'fix-nomes':
            config.fix_only.append(name)
        elif rol == 'fix-i-rota':
            config.fix_and_rota.append(name)
            config.rotators.append(name)  # rotates too
        elif rol == 'sun-nit-only':
            if name not in config.sun_nit_only:
                config.sun_nit_only.append(name)
            config.rotators.append(name)
        elif rol == 'weekend-day-only':
            if name not in config.weekend_day_only:
                config.weekend_day_only.append(name)
            config.rotators.append(name)
        elif rol == 'nou-incorporat' or rol == 'nou incorporat':
            # New professionals: only cover Wed afternoons (except 1st Wed of month)
            # NOT part of general rotation until they graduate to 'rotador'.
            config.nou_incorporats.append(name)

    # ----- Read Vacances sheet -----
    if 'Vacances' not in wb.sheetnames:
        raise ValueError("Falta la pestanya 'Vacances' al fitxer d'entrada.")
    ws_v = wb['Vacances']
    warnings = []

    header_row = _find_label_row(ws_v, 'radiol') or 3     # 'Radiòleg' -> 'radioleg'
    date_cols, int_cols = {}, {}
    for col in range(2, ws_v.max_column + 1):
        v = ws_v.cell(row=header_row, column=col).value
        d = _as_date(v) if not isinstance(v, str) else None
        if d is not None:
            date_cols[col] = d
        elif isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and 1 <= int(v) <= 31:
            int_cols[col] = int(v)

    if date_cols:
        meta, col_index = _meta_setmanes(ws_v, header_row, date_cols, warnings)
    else:
        meta, col_index = _meta_mes(ws_v, int_cols, warnings)

    first_name_row = header_row + 1
    fest_row = _find_label_row(ws_v, 'festiu')
    if fest_row and fest_row >= first_name_row:
        first_name_row = fest_row + 1

    constraints = {}
    seen = set()
    for row in range(first_name_row, ws_v.max_row + 1):
        raw = ws_v.cell(row=row, column=1).value
        name = raw.strip() if isinstance(raw, str) else ''
        cons = {}
        for col, idx in col_index.items():
            val = ws_v.cell(row=row, column=col).value
            if val and isinstance(val, str):
                v = val.strip().upper()
                if v in CODIS_FULL:
                    cons[idx] = v
        if not name:
            if cons:
                warnings.append(f"Vacances, fila {row}: hi ha marques però no hi ha nom (s'ignoren).")
            continue
        if name not in radiologists:
            if cons:
                warnings.append(f"Vacances, fila {row}: «{name}» no és a la pestanya Radiòlegs; "
                                "les seves marques s'ignoren.")
            continue
        seen.add(name)
        constraints.setdefault(name, {}).update(cons)

    # Assegura que tothom té entrada, també els que no surten a la pestanya Vacances
    for name in radiologists:
        constraints.setdefault(name, {})

    sense_fila = sorted((n for n in radiologists if n not in seen and n not in inactius_set), key=fold)
    if sense_fila:
        warnings.append("Sense fila a la pestanya Vacances (no se'ls poden marcar absències): "
                        + ', '.join(sense_fila))

    _apply_derived_codes(constraints, inactius_set, meta.n_days)
    meta.load_warnings = warnings
    return constraints, config, meta


def _apply_derived_codes(constraints: dict, inactius: set, n_days: int) -> None:
    """Afegeix els codis derivats (I i R) al diccionari de constriccions, in place.

    - I : tot el període bloquejat per als professionals amb Actiu = "No".
    - R : marge de reincorporació després del darrer dia d'una absència V/B.
          Amb DIES_REINCORPORACIO = 3, si l'absència acaba el dia D el
          professional torna a estar disponible el dia D+3, i per tant es
          bloquegen D+1 i D+2. Les absències que acaben als dies de context
          (període anterior, índexs <= 0) també generen marge dins del període.
    """
    for name in inactius:
        cons = constraints.setdefault(name, {})
        for day in range(1, n_days + 1):
            cons[day] = CODI_INACTIU

    extra_blocked = max(0, DIES_REINCORPORACIO - 1)
    if extra_blocked == 0:
        return

    for name, cons in constraints.items():
        if name in inactius or not cons:
            continue
        lo = min(min(cons), 1)
        # Dies en què s'acaba un bloc d'absència llarga (V/B)
        fins = [d for d in range(lo, n_days + 1)
                if cons.get(d) in CODIS_AMB_REINCORPORACIO
                and cons.get(d + 1) not in CODIS_AMB_REINCORPORACIO]
        for d in fins:
            for offset in range(1, extra_blocked + 1):
                nd = d + offset
                if nd > n_days:
                    break
                if nd < 1:
                    continue
                if nd not in cons:          # no trepitgem V/B/C/G/X ja existents
                    cons[nd] = CODI_REINCORPORACIO


# ============================================================
# Algorithm
# ============================================================

def _can_work(name: str, day: int, constraints: dict) -> bool:
    cons = constraints.get(name, {})
    if cons.get(day) in CODIS_BLOQUEJANTS:
        return False
    if cons.get(day - 1) == 'G':      # inclou la G d'un dia de context
        return False
    if cons.get(day + 1) == 'G':
        return False
    return True


def _motiu_bloqueig(name: str, day: int, constraints: dict) -> str:
    """Text llegible del motiu pel qual algú no pot fer guàrdia un dia."""
    cons = constraints.get(name, {})
    code = cons.get(day)
    if code in CODIS_BLOQUEJANTS:
        return MOTIU_TEXT.get(code, code)
    if cons.get(day - 1) == 'G' or cons.get(day + 1) == 'G':
        return 'guàrdia externa adjacent'
    return 'no disponible'


def _build_pool(config: Config, start_name: str) -> tuple[list, str]:
    pool = sorted(set(config.rotators), key=fold)
    if not start_name or not pool:
        return pool, ''
    sf = fold(start_name).strip()
    for i, n in enumerate(pool):
        if sf in fold(n) or fold(n).startswith(sf):
            return pool[i:] + pool[:i], ''
    # No és a la roda (inactiu, fix-només o mal escrit): seguim l'ordre alfabètic
    for i, n in enumerate(pool):
        if fold(n) > sf:
            return (pool[i:] + pool[:i],
                    f"«{start_name}» no és a la roda d'aquest mes; es comença pel següent "
                    f"en ordre alfabètic: {pool[i]}.")
    return pool, f"«{start_name}» no és a la roda d'aquest mes; es comença per {pool[0]}."


def _gen_slots(meta: Meta) -> list:
    slots = []
    for day in range(1, meta.n_days + 1):
        dow = meta.dow(day)
        is_fest = day in meta.festius
        if is_fest and dow <= 2:
            slots += [(day, dow, 'dia (8-20)', 'NB'),
                      (day, dow, 'nit (20-08)', 'NB')]
        elif is_fest and dow in (3, 4):
            slots += [(day, dow, 'dia (8-20)', 'G1'),
                      (day, dow, 'dia (8-20)', 'G2'),
                      (day, dow, 'nit (20-08)', 'G1'),
                      (day, dow, 'nit (20-08)', 'G2')]
        elif dow in (0, 1):
            slots.append((day, dow, 'nit (20-08)', 'NB'))
        elif dow == 2:
            if not meta.is_first_wed(day):
                slots.append((day, dow, '16-20', 'NB'))
            slots.append((day, dow, 'nit (20-08)', 'NB'))
        elif dow in (3, 4):
            slots += [(day, dow, '16-20', 'G1'),
                      (day, dow, '16-20', 'G2'),
                      (day, dow, 'nit (20-08)', 'G1'),
                      (day, dow, 'nit (20-08)', 'G2')]
        elif dow == 5:
            slots += [(day, dow, 'dia (8-20)', 'G1'),
                      (day, dow, 'dia (8-20)', 'G2'),
                      (day, dow, 'nit (20-08)', 'G1'),
                      (day, dow, 'nit (20-08)', 'G2')]
        elif dow == 6:
            slots += [(day, dow, 'dia (8-20)', 'NB'),
                      (day, dow, 'nit (20-08)', 'NB')]
    return slots


def generate_planning(constraints: dict, config: Config, meta: Meta) -> tuple[dict, list, list]:
    warnings = []
    pool, avis_pool = _build_pool(config, meta.start_radiologist)
    if avis_pool:
        warnings.append(avis_pool)
    if not pool:
        raise ValueError("Pool de rotació buit. Revisa la configuració.")

    slots = _gen_slots(meta)
    assignments = {}
    queue = deque(pool)
    # Separate rotation queue for nou-incorporats (Wed 16-20 slots only)
    nou_queue = deque(sorted(set(config.nou_incorporats), key=fold))
    assigned_today = defaultdict(set)
    deferred_sun_nit = 0      # torns pendents de perfils sun-nit-only
    deferred_weekend = 0      # torns pendents de perfils weekend-day-only

    # Pre-populate assigned_today with fixed-slot occupants
    # so that fix-i-rota professionals don't get a rotation slot on a day they cover a fix.
    for day in range(1, meta.n_days + 1):
        dow = meta.dow(day)
        if day in meta.festius:
            continue
        # Regular fix slots
        for (d_, lbl, role), name in config.fix_slots.items():
            if d_ != dow:
                continue
            if not _can_work(name, day, constraints):
                continue
            assigned_today[day].add(name)
        # First Wednesday special slot
        if dow == 2 and meta.is_first_wed(day) and config.first_wed_radiologist:
            if _can_work(config.first_wed_radiologist, day, constraints):
                assigned_today[day].add(config.first_wed_radiologist)

    def is_sun_nit(s): return s[1] == 6 and 'nit' in s[2]
    def is_weekend_dia(s): return s[1] in (5, 6) and 'dia' in s[2]
    def is_wed_afternoon(s): return s[1] == 2 and s[2] == '16-20'

    def is_sun_nit_only(n): return n in config.sun_nit_only
    def is_weekend_day_only(n): return n in config.weekend_day_only

    for slot in slots:
        day = slot[0]
        # Torns pendents de perfils sun-nit-only: al primer Dg nit lliure
        if is_sun_nit(slot) and deferred_sun_nit > 0:
            for cand_sn in config.sun_nit_only:
                if _can_work(cand_sn, day, constraints) and cand_sn not in assigned_today[day]:
                    assignments[slot] = cand_sn
                    assigned_today[day].add(cand_sn)
                    deferred_sun_nit -= 1
                    break
            if slot in assignments:
                continue
        # Torns pendents de perfils weekend-day-only: al primer Ds/Dg dia lliure
        if is_weekend_dia(slot) and deferred_weekend > 0:
            for cand_wd in config.weekend_day_only:
                if _can_work(cand_wd, day, constraints) and cand_wd not in assigned_today[day]:
                    assignments[slot] = cand_wd
                    assigned_today[day].add(cand_wd)
                    deferred_weekend -= 1
                    break
            if slot in assignments:
                continue

        # Wed afternoon (16-20, excluding 1st Wed which is a fix): try nou-incorporats first
        if is_wed_afternoon(slot) and nou_queue:
            chosen_nou = None
            tries_nou = 0
            n_nou = len(nou_queue)
            temp_skip_nou = []
            while nou_queue and tries_nou < n_nou:
                cand_nou = nou_queue.popleft()
                if _can_work(cand_nou, day, constraints) and cand_nou not in assigned_today[day]:
                    chosen_nou = cand_nou
                    break
                else:
                    temp_skip_nou.append(cand_nou)
                tries_nou += 1
            # Restore skipped nou-incorporats at the front (they didn't lose their turn)
            for n_ in reversed(temp_skip_nou):
                nou_queue.appendleft(n_)
            if chosen_nou:
                assignments[slot] = chosen_nou
                assigned_today[day].add(chosen_nou)
                nou_queue.append(chosen_nou)  # move to back of nou queue
                continue
            # If no nou-incorporat available, fall through to normal rotation

        # Normal rotation
        chosen = None
        tries = 0
        nq = len(queue)
        temp_skip = []
        while queue and tries < nq:
            cand = queue.popleft()
            if is_sun_nit_only(cand):
                deferred_sun_nit += 1
                queue.append(cand)
                tries += 1
                continue
            if is_weekend_day_only(cand):
                deferred_weekend += 1
                queue.append(cand)
                tries += 1
                continue
            if _can_work(cand, day, constraints) and cand not in assigned_today[day]:
                chosen = cand
                break
            else:
                temp_skip.append(cand)
            tries += 1
        for n in reversed(temp_skip):
            queue.appendleft(n)
        if chosen is None:
            warnings.append(f"No s'ha trobat ningú disponible per al torn {meta.fmt(day)} {slot[2]} {slot[3]}")
            continue
        assignments[slot] = chosen
        assigned_today[day].add(chosen)
        queue.append(chosen)

    if deferred_sun_nit > 0:
        warnings.append(f"{deferred_sun_nit} torn(s) de Sun-nit-only pendents")
    if deferred_weekend > 0:
        warnings.append(f"{deferred_weekend} torn(s) de Weekend-day-only pendents")

    return assignments, list(queue), warnings


def last_in_rotation(queue_final: list) -> str:
    return queue_final[-1] if queue_final else ''


def last_rotator(assignments: dict, config: Config) -> str:
    """Últim professional de la roda general que ha fet guàrdia al període."""
    excl = (set(config.fix_only) | set(config.sun_nit_only) |
            set(config.weekend_day_only) | set(config.nou_incorporats))
    last = ''
    for slot in sorted(assignments, key=lambda s: s[0]):   # sorted és estable: conserva l'ordre del dia
        name = assignments[slot]
        if name not in excl:
            last = name
    return last


def next_start_radiologist(assignments: dict, config: Config) -> str:
    """Primer de la roda del període següent: el següent alfabètic de l'últim rotador."""
    last = last_rotator(assignments, config)
    pool = sorted(set(config.rotators), key=fold)
    if not last or not pool:
        return ''
    if last in pool:
        return pool[(pool.index(last) + 1) % len(pool)]
    lf = fold(last)
    for n in pool:
        if fold(n) > lf:
            return n
    return pool[0]


# ============================================================
# Writer
# ============================================================

def prepare_planning_sheets(wb, meta: Meta) -> list:
    """Deixa el llibre amb una pestanya per setmana del període (dl→dg), datada i
    buida. Afegeix o treu pestanyes 'Setmana' segons calgui. Retorna les pestanyes
    en ordre, alineades amb meta.week_mondays()."""
    mondays = meta.week_mondays()
    sheets = [s for s in wb.sheetnames if 'setmana' in s.lower()]
    if not sheets:
        raise ValueError("La plantilla del planning no té cap pestanya «Setmana».")
    while len(sheets) < len(mondays):
        new = wb.copy_worksheet(wb[sheets[-1]])
        sheets.append(new.title)
    for sn in sheets[len(mondays):]:
        wb.remove(wb[sn])
    sheets = sheets[:len(mondays)]
    for i, sn in enumerate(sheets):            # noms temporals per evitar col·lisions
        wb[sn].title = f"tmp_setmana_{i}"

    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=False)
    out_font = Font(bold=True, size=10, color='808080')
    out_fill = PatternFill('solid', fgColor='D9D9D9')
    out_body = PatternFill('solid', fgColor='F2F2F2')
    fest_font = Font(bold=True, size=10, color='C0392B')
    fest_fill = PatternFill('solid', fgColor='FFF3CD')
    no_fill = PatternFill(fill_type=None)

    result = []
    for i, monday in enumerate(mondays):
        ws = wb[f"tmp_setmana_{i}"]
        ws.title = week_sheet_name(monday)
        sunday = monday + datetime.timedelta(days=6)
        ws['A1'] = f"Setmana {monday.isocalendar()[1]:02d} · {fmt_dia(monday)} – {fmt_dia(sunday, True)}"
        ws['A1'].font = Font(bold=True, size=12)
        ws['A1'].alignment = left
        ws['A2'] = (f"Planning de {meta.label}: {fmt_dia(meta.start, True)} → "
                    f"{fmt_dia(meta.end, True)} ({meta.n_weeks} setmanes)")
        ws['A2'].font = Font(italic=True, size=9, color='595959')
        ws['A2'].alignment = left
        for r in range(3, 24):
            for c in range(2, 9):
                cell = ws.cell(row=r, column=c)
                cell.value = None
                if r >= 4:
                    cell.fill = no_fill
        for dow in range(7):
            d = monday + datetime.timedelta(days=dow)
            col = 2 + dow
            hdr = ws.cell(row=3, column=col)
            hdr.value = d
            hdr.number_format = '[$-403]ddd dd/mm/yyyy'
            hdr.alignment = center
            idx = meta.index_of(d)
            if idx is None:
                hdr.font = out_font
                hdr.fill = out_fill
                ws.cell(row=4, column=col).value = '(fora del període — vegeu el planning anterior)' \
                    if d < meta.start else '(fora del període — vegeu el planning següent)'
                ws.cell(row=4, column=col).font = Font(italic=True, size=9, color='808080')
                ws.cell(row=4, column=col).alignment = Alignment(wrap_text=True, vertical='top')
                for r in range(4, 24):
                    ws.cell(row=r, column=col).fill = out_body
            elif idx in meta.festius:
                hdr.font = fest_font
                hdr.fill = fest_fill
        result.append(ws)
    return result


def write_planning(assignments: dict, config: Config, meta: Meta, constraints: dict,
                   template_path, output_path):
    wb = openpyxl.load_workbook(template_path)
    week_sheets = prepare_planning_sheets(wb, meta)
    mondays = meta.week_mondays()

    bold = Font(bold=True, size=10)
    small = Font(size=10)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    fill_fix = PatternFill('solid', fgColor='E8F1FB')
    fill_rot = PatternFill('solid', fgColor='F4FBF0')
    fill_fest = PatternFill('solid', fgColor='FFF3CD')
    fill_sub = PatternFill('solid', fgColor='FBEFD8')

    def fixed_for(day, dow, label):
        if day in meta.festius:
            return None
        if label == '16-20' and dow == 2 and meta.is_first_wed(day):
            if config.first_wed_radiologist:
                return [('N i B', f"{config.first_wed_radiologist} (fix 1r dimecres)")]
        result = []
        for role in ('N', 'B', 'N i B'):
            name = config.fix_slots.get((dow, label, role))
            if name:
                # Apply substitution if person can't work
                if not _can_work(name, day, constraints):
                    reason = _motiu_bloqueig(name, day, constraints)
                    name = f"[SUBSTITUIR — {name}: {reason}]"
                result.append((role, name))
        return result if result else None

    def put(ws, col, label, lines, fill):
        start, end = SHIFT_ROWS[label]
        r = start
        for line in lines:
            if r > end:
                break
            c = ws.cell(row=r, column=col)
            c.value = line
            c.alignment = left
            c.font = bold if any(line.startswith(p) for p in ('N:', 'B:', 'N i B', 'Grup', 'Festiu', '(', 'Dissabte', 'Diumenge', 'FESTIU', '[')) else small
            c.fill = fill
            r += 1

    for day in range(1, meta.n_days + 1):
        d = meta.date_of(day)
        dow = d.weekday()
        ws = week_sheets[(d - mondays[0]).days // 7]
        col = 2 + dow
        is_fest = day in meta.festius

        # 8-16
        if is_fest and dow <= 2:
            nb = assignments.get((day, dow, 'dia (8-20)', 'NB'))
            put(ws, col, '8-16', ['FESTIU — Torn dia 8-20:', f'N i B: {nb}' if nb else '—'], fill_fest)
        elif is_fest and dow in (3, 4):
            g1 = assignments.get((day, dow, 'dia (8-20)', 'G1'))
            g2 = assignments.get((day, dow, 'dia (8-20)', 'G2'))
            put(ws, col, '8-16', ['FESTIU — Torn dia 8-20:', f'Grup 1: {g1}', f'Grup 2: {g2}'], fill_fest)
        elif dow == 5:
            g1 = assignments.get((day, dow, 'dia (8-20)', 'G1'))
            g2 = assignments.get((day, dow, 'dia (8-20)', 'G2'))
            put(ws, col, '8-16', ['Dissabte — Torn dia 8-20:', f'Grup 1: {g1}', f'Grup 2: {g2}'], fill_rot)
        elif dow == 6:
            nb = assignments.get((day, dow, 'dia (8-20)', 'NB'))
            put(ws, col, '8-16', ['Diumenge — Torn dia 8-20:', f'N i B: {nb}'], fill_rot)
        else:
            fx = fixed_for(day, dow, '8-16')
            if fx:
                put(ws, col, '8-16', [f'{r}: {n}' for r, n in fx], fill_fix)

        # 16-20
        if is_fest:
            put(ws, col, '16-20', ['(inclòs torn dia)'], fill_fest)
        elif dow in (5, 6):
            put(ws, col, '16-20', ['(inclòs torn dia)'], fill_rot)
        else:
            fx = fixed_for(day, dow, '16-20')
            if fx:
                has_sub = any('SUBSTITUIR' in n for _, n in fx)
                fill = fill_sub if has_sub else (fill_fest if (dow == 2 and meta.is_first_wed(day)) else fill_fix)
                put(ws, col, '16-20', [f'{r}: {n}' for r, n in fx], fill)
            elif dow == 2:
                nb = assignments.get((day, dow, '16-20', 'NB'))
                put(ws, col, '16-20', [f'N i B: {nb}'], fill_rot)
            elif dow in (3, 4):
                g1 = assignments.get((day, dow, '16-20', 'G1'))
                g2 = assignments.get((day, dow, '16-20', 'G2'))
                put(ws, col, '16-20', [f'Grup 1: {g1}', f'Grup 2: {g2}'], fill_rot)

        # 20-08
        if is_fest and dow <= 2:
            nb = assignments.get((day, dow, 'nit (20-08)', 'NB'))
            put(ws, col, '20-08', ['FESTIU — Torn nit 20-08:', f'N i B: {nb}'], fill_fest)
        elif is_fest and dow in (3, 4):
            g1 = assignments.get((day, dow, 'nit (20-08)', 'G1'))
            g2 = assignments.get((day, dow, 'nit (20-08)', 'G2'))
            put(ws, col, '20-08', ['FESTIU — Torn nit 20-08:', f'Grup 1: {g1}', f'Grup 2: {g2}'], fill_fest)
        elif dow in (0, 1, 2):
            nb = assignments.get((day, dow, 'nit (20-08)', 'NB'))
            put(ws, col, '20-08', [f'N i B: {nb}'], fill_rot)
        elif dow in (3, 4, 5):
            g1 = assignments.get((day, dow, 'nit (20-08)', 'G1'))
            g2 = assignments.get((day, dow, 'nit (20-08)', 'G2'))
            put(ws, col, '20-08', [f'Grup 1: {g1}', f'Grup 2: {g2}'], fill_rot)
        elif dow == 6:
            nb = assignments.get((day, dow, 'nit (20-08)', 'NB'))
            put(ws, col, '20-08', [f'N i B: {nb}'], fill_rot)

    for ws in week_sheets:
        ws.column_dimensions['A'].width = 14
        for c in range(2, 9):
            ws.column_dimensions[get_column_letter(c)].width = 32

    wb.save(output_path)
    return output_path


def validate_planning(assignments: dict, config: Config, constraints: dict, meta: Meta) -> list:
    errors = []
    by_day = defaultdict(list)
    for slot, name in assignments.items():
        by_day[slot[0]].append(name)
    for day, names in by_day.items():
        if len(names) != len(set(names)):
            errors.append(f"{meta.fmt(day)}: doble assignació")
    for slot, name in assignments.items():
        if not _can_work(name, slot[0], constraints):
            motiu = _motiu_bloqueig(name, slot[0], constraints)
            errors.append(f"{meta.fmt(slot[0])} {slot[2]} {slot[3]}: {name} en conflicte ({motiu})")
    # Cap professional marcat com a NO actiu pot aparèixer a la roda
    inactius = set(config.inactius)
    for slot, name in assignments.items():
        if name in inactius:
            errors.append(f"{meta.fmt(slot[0])} {slot[2]}: {name} està marcat com a NO actiu aquest mes")
    for slot, name in assignments.items():
        if name in config.sun_nit_only and not (slot[1] == 6 and 'nit' in slot[2]):
            errors.append(f"{meta.fmt(slot[0])} {slot[2]}: {name} (Sun-nit-only) fora del seu torn")
    for slot, name in assignments.items():
        if name in config.weekend_day_only and not (slot[1] in (5, 6) and 'dia' in slot[2]):
            errors.append(f"{meta.fmt(slot[0])} {slot[2]}: {name} (Weekend-day-only) fora del seu torn")
    # Nou-incorporats han d'aparèixer només a dimecres 16-20 (excepte 1r Dc del mes)
    for slot, name in assignments.items():
        if name in config.nou_incorporats:
            if not (slot[1] == 2 and slot[2] == '16-20'):
                errors.append(f"{meta.fmt(slot[0])} {slot[2]}: {name} (Nou-incorporat) fora del seu torn Dc 16-20")
            elif meta.is_first_wed(slot[0]):
                errors.append(f"{meta.fmt(slot[0])} {slot[2]}: {name} (Nou-incorporat) NO pot cobrir el 1r dimecres del mes")
    # Equity (excluding fix and special)
    counts = defaultdict(int)
    exclude = (set(config.fix_only) | set(config.sun_nit_only) | set(config.weekend_day_only) |
               set(config.fix_and_rota) | set(config.nou_incorporats))
    for n in assignments.values():
        if n not in exclude:
            counts[n] += 1
    if counts:
        mx, mn = max(counts.values()), min(counts.values())
        if mx - mn > 1:
            errors.append(f"Distribució desigual entre rotadors purs: màx {mx}, mín {mn}")
    return errors


# ============================================================
# CLI
# ============================================================
if __name__ == '__main__':
    import argparse, sys
    p = argparse.ArgumentParser()
    p.add_argument('input', help='Fitxer Excel d\'entrada (Configuració + Radiòlegs + Vacances)')
    p.add_argument('template', help='Plantilla Excel buida del planning')
    p.add_argument('output', help='Camí de sortida del planning omplert')
    args = p.parse_args()

    cons, config, meta = load_input(args.input)
    print(f"Planning: {meta.label} · {fmt_dia(meta.start, True)} → {fmt_dia(meta.end, True)} "
          f"({meta.n_days} dies, {meta.n_weeks} setmanes)")
    print(f"Festius: {meta.festius_text() or 'cap'}")
    print(f"Rotadors: {len(config.rotators)}, Fixos només: {len(config.fix_only)}")
    print(f"Sun-nit-only: {config.sun_nit_only}")
    print(f"Weekend-day-only: {config.weekend_day_only}")
    print(f"Nou-incorporats (dimecres 16-20): {config.nou_incorporats}")
    if config.inactius:
        print(f"NO actius aquest mes ({len(config.inactius)}): {', '.join(config.inactius)}")
    for w in meta.load_warnings:
        print(f"  · {w}")

    assignments, queue, warn = generate_planning(cons, config, meta)
    write_planning(assignments, config, meta, cons, args.template, args.output)
    errs = validate_planning(assignments, config, cons, meta)

    print(f"\nAssignacions: {len(assignments)}")
    print(f"Últim de la roda: {last_rotator(assignments, config)}")
    print(f"Primer del període següent: {next_start_radiologist(assignments, config)}")
    if warn:
        print("\nAvisos:")
        for w in warn: print(f"  - {w}")
    if errs:
        print("\nErrors:")
        for e in errs: print(f"  - {e}")
        sys.exit(1)
    print("\nValidació OK ✓")
