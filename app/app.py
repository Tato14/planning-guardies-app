"""Streamlit UI per al generador de planning de guàrdies (config-driven, setmanes completes)."""
import hashlib
import io
import tempfile
from collections import Counter, defaultdict

import streamlit as st

import planning_generator as pg
import period_templates as ptpl
from planning_generator import (
    load_input, generate_planning, write_planning, validate_planning,
)

# Es llegeix del mòdul, no s'importa pel nom: si Streamlit Cloud servís una
# versió antiga del mòdul en memòria, l'app segueix arrencant en comptes de
# petar amb un ImportError.
DIES_REINCORPORACIO = getattr(pg, 'DIES_REINCORPORACIO', 3)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Generador de planning de guàrdies", page_icon="🩺", layout="wide")

st.title("🩺 Generador de planning de guàrdies")
st.caption("Genera el planning de cada mes per **setmanes completes** (dilluns → diumenge) seguint les "
           "regles del fitxer d'entrada (Configuració + Radiòlegs + Vacances).")

with st.sidebar:
    st.header("Com fer servir")
    st.markdown("""
1. **Puja el fitxer d'entrada del mes** (plantilla mensual, amb les vacances posades).
2. **Revisa el període** i la configuració detectada.
3. **Puja la plantilla buida** del planning (la del mes o la genèrica: s'adapta sola).
4. **Clica "Generar planning"**.
5. **Descarrega** el planning i, si vols, l'entrada del mes següent ja preparada.
""")
    st.divider()
    st.markdown("**Com va el període**")
    st.markdown("Cada planning té setmanes senceres. Una setmana pertany al mes del seu **dilluns**: "
                "el de novembre de 2026 va del dl 02/11 al dg 06/12.")
    st.divider()
    st.markdown("Plantilles de cada mes: carpeta `Plantilles mensuals` del paquet.")


def _bytes_wb(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _period_text(meta) -> str:
    return f"{pg.fmt_dia(meta.start, True)} → {pg.fmt_dia(meta.end, True)}"


# ------------------------------------------------------------
# Pas 1 — Fitxer d'entrada
# ------------------------------------------------------------
st.header("Pas 1 — Fitxer d'entrada")
in_file = st.file_uploader(
    "Puja el fitxer Excel d'entrada del mes (Configuració + Radiòlegs + Vacances)",
    type=["xlsx"], key="in_file"
)

constraints = config = meta = None
in_bytes = None
if in_file:
    in_bytes = in_file.getvalue()
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as f:
            f.write(in_bytes)
            in_path = f.name
        constraints, config, meta = load_input(in_path)
        st.success(f"Fitxer carregat: **{meta.label}**, del {_period_text(meta)} "
                   f"({meta.n_weeks} setmanes, {meta.n_days} dies). "
                   f"{len(constraints)} radiòlegs a la plantilla, {len(config.rotators)} a la roda.")
        for w in meta.load_warnings:
            if w.startswith('Fitxer en format antic') or 'dia pont' in w:
                st.info(w)
            else:
                st.warning(w)
        if config.inactius:
            st.warning(
                f"**{len(config.inactius)} professional(s) marcats com a NO actius** aquest mes "
                f"(columna E de la pestanya Radiòlegs): {', '.join(sorted(config.inactius))}. "
                "Queden fora de la roda sense haver-los d'esborrar de la plantilla."
            )
    except Exception as e:
        st.error(f"Error llegint el fitxer: {e}")
        import traceback
        st.code(traceback.format_exc())
        constraints = config = meta = None

# ------------------------------------------------------------
# Pas 2 — Configuració detectada
# ------------------------------------------------------------
if constraints and meta and config:
    st.header("Pas 2 — Període i configuració")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Mes del planning", meta.label)
    with c2:
        st.metric("Període", f"{meta.start:%d/%m} → {meta.end:%d/%m/%Y}")
    with c3:
        st.metric("Setmanes", meta.n_weeks)
    with c4:
        st.metric("Festius", meta.festius_text() or "Cap")

    e1, e2 = st.columns(2)
    with e1:
        new_start = st.text_input("Primer radiòleg de la roda (editable)", value=meta.start_radiologist or "")
        meta.start_radiologist = new_start
    with e2:
        new_festius = st.text_input("Festius dins del període (dd/mm, separats per comes; editable)",
                                    value=meta.festius_text())
        for a in meta.set_festius_from_text(new_festius):
            st.warning(a)

    with st.expander("📋 Veure rols i fixos detectats"):
        st.subheader("Torns fixos")
        DOW_NAMES = ['Dilluns', 'Dimarts', 'Dimecres', 'Dijous', 'Divendres', 'Dissabte', 'Diumenge']
        fix_rows = []
        for (dow, shift, role), name in sorted(config.fix_slots.items()):
            fix_rows.append({"Dia": DOW_NAMES[dow], "Franja": shift, "Rol": role, "Radiòleg": name})
        if config.first_wed_radiologist:
            fix_rows.append({"Dia": "1r dimecres del mes", "Franja": "16-20", "Rol": "N i B",
                             "Radiòleg": config.first_wed_radiologist})
        if fix_rows:
            st.dataframe(fix_rows, hide_index=True, width='stretch')
        else:
            st.info("Cap torn fix definit.")

        st.subheader("Casos especials")
        st.write(f"**Sun-nit-only**: {', '.join(config.sun_nit_only) if config.sun_nit_only else 'cap'}")
        st.write(f"**Weekend-day-only**: {', '.join(config.weekend_day_only) if config.weekend_day_only else 'cap'}")
        st.write(f"**Fix-i-rota**: {', '.join(config.fix_and_rota) if config.fix_and_rota else 'cap'}")
        st.write(f"**Fix-només**: {', '.join(config.fix_only) if config.fix_only else 'cap'}")
        st.write(f"**Nou-incorporat** (cobreix dimecres 16-20): {', '.join(config.nou_incorporats) if config.nou_incorporats else 'cap'}")
        st.write(f"**NO actius aquest mes** (col. E = No): {', '.join(sorted(config.inactius)) if config.inactius else 'cap'}")

        st.subheader("Codis de la pestanya Vacances")
        st.markdown(f"""
| Codi | Significat | Efecte |
|---|---|---|
| `V` | Vacances | Bloqueja el dia + {DIES_REINCORPORACIO - 1} dies de marge de reincorporació |
| `B` | Baixa | Bloqueja el dia + {DIES_REINCORPORACIO - 1} dies de marge de reincorporació |
| `C` | Congrés | Bloqueja només el dia |
| `G` | Guàrdia externa | Bloqueja el dia, l'anterior i el posterior |
| `X` | Guàrdia ja assignada | Informatiu, no afecta l'algorisme |
""")

    with st.expander("📅 Preparar la plantilla d'entrada d'un altre mes"):
        st.caption("Crea l'entrada d'un mes qualsevol amb els mateixos radiòlegs, rols, Actiu i "
                   "configuració que el fitxer que has pujat, i la pestanya Vacances nova "
                   "(dates del període i festius marcats).")
        opcions, (y, m) = [], (meta.year, meta.month)
        for _ in range(18):
            y, m = pg.next_month(y, m)
            opcions.append((y, m))
        tria = st.selectbox("Mes", opcions, format_func=lambda ym: pg.month_label(*ym), key="altre_mes")
        try:
            data_altre = _bytes_wb(ptpl.build_input_template(in_bytes, tria[0], tria[1]))
            s_, e_ = pg.period_for_month(*tria)
            st.caption(f"Període: {pg.fmt_dia(s_, True)} → {pg.fmt_dia(e_, True)}")
            st.download_button("📥 Descarregar l'entrada", data=data_altre,
                               file_name=f"Entrada_{tria[0]}-{tria[1]:02d}.xlsx", mime=XLSX_MIME,
                               key="dl_altre")
        except Exception as e:
            st.error(f"No s'ha pogut preparar la plantilla: {e}")

# ------------------------------------------------------------
# Pas 3 i 4 — Plantilla i generació
# ------------------------------------------------------------
if constraints and meta and config:
    st.header("Pas 3 — Plantilla del planning")
    tmpl_file = st.file_uploader("Puja la plantilla buida del planning (la del mes o la genèrica)",
                                 type=["xlsx"], key="tmpl")

    if tmpl_file:
        st.success("Plantilla carregada. Es crearan tantes pestanyes setmanals com calgui.")
        tmpl_bytes = tmpl_file.getvalue()
        run_key = hashlib.sha1(in_bytes + tmpl_bytes + meta.start_radiologist.encode()
                               + meta.festius_text().encode()).hexdigest()

        st.header("Pas 4 — Generar planning")
        if st.button("🚀 Generar planning", type="primary", width='stretch'):
            with st.spinner("Calculant..."):
                try:
                    assignments, queue_final, warnings = generate_planning(constraints, config, meta)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as f:
                        f.write(tmpl_bytes)
                        tmp_tmpl = f.name
                    out_path = tmp_tmpl.replace('.xlsx', '_planning.xlsx')
                    write_planning(assignments, config, meta, constraints, tmp_tmpl, out_path)
                    errors = validate_planning(assignments, config, constraints, meta)
                    with open(out_path, 'rb') as f:
                        planning_bytes = f.read()

                    seg = pg.next_start_radiologist(assignments, config)
                    ny, nm = pg.next_month(meta.year, meta.month)
                    try:
                        next_bytes = _bytes_wb(ptpl.build_input_template(in_bytes, ny, nm, start_radiologist=seg))
                        next_err = ''
                    except Exception as e:
                        next_bytes, next_err = None, str(e)

                    by_day = defaultdict(list)
                    for slot, name in assignments.items():
                        by_day[slot[0]].append(f"{slot[2]} {slot[3]}: {name}")

                    st.session_state['resultat'] = {
                        'key': run_key,
                        'n': len(assignments),
                        'errors': errors,
                        'warnings': warnings,
                        'dist': Counter(Counter(assignments.values()).values()),
                        'last': pg.last_rotator(assignments, config),
                        'next': seg,
                        'next_ym': (ny, nm),
                        'planning': planning_bytes,
                        'next_input': next_bytes,
                        'next_err': next_err,
                        'by_day': [(meta.fmt(d), lines) for d, lines in sorted(by_day.items())],
                    }
                except Exception as e:
                    st.session_state.pop('resultat', None)
                    st.error(f"Error: {e}")
                    import traceback
                    st.code(traceback.format_exc())

        res = st.session_state.get('resultat')
        if res and res['key'] == run_key:
            st.divider()
            st.subheader(f"📊 Resultats — {meta.label} ({_period_text(meta)})")
            if not res['errors']:
                st.success(f"✅ Validació OK — {res['n']} assignacions, sense conflictes.")
            else:
                st.error(f"❌ {len(res['errors'])} problemes:")
                for e in res['errors']:
                    st.markdown(f"- {e}")
            if res['warnings']:
                with st.expander("⚠️ Avisos"):
                    for w in res['warnings']:
                        st.markdown(f"- {w}")

            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("Total guàrdies", res['n'])
            with m2:
                st.metric("Persones amb 1", res['dist'].get(1, 0))
            with m3:
                st.metric("Persones amb 2+", sum(v for k, v in res['dist'].items() if k >= 2))

            st.info(f"🔄 **Últim de la roda:** {res['last'] or '—'} · "
                    f"**Primer del mes següent:** {res['next'] or '—'}")

            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    label=f"📥 Descarregar el planning ({meta.label})",
                    data=res['planning'],
                    file_name=f"Planning_{meta.code}.xlsx",
                    mime=XLSX_MIME, width='stretch', key="dl_planning",
                )
            with d2:
                ny, nm = res['next_ym']
                if res['next_input']:
                    st.download_button(
                        label=f"📥 Entrada de {pg.month_label(ny, nm)} ja preparada",
                        data=res['next_input'],
                        file_name=f"Entrada_{ny}-{nm:02d}.xlsx",
                        mime=XLSX_MIME, width='stretch', key="dl_next",
                        help="Mateixos radiòlegs i configuració, primer de la roda posat i festius marcats. "
                             "Només cal omplir-hi les vacances.",
                    )
                else:
                    st.warning(f"No s'ha pogut preparar l'entrada del mes següent: {res['next_err']}")

            with st.expander("📅 Veure totes les assignacions"):
                for label, lines in res['by_day']:
                    st.markdown(f"**{label}**")
                    for line in lines:
                        st.text(f"  {line}")

st.divider()
st.caption("Algorisme v3.0 — setmanes completes, config-driven, sense dades identificatives al codi.")
