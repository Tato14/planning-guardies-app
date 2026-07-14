"""Streamlit UI per al generador de planning de guàrdies (config-driven)."""
import datetime, tempfile
from collections import Counter, defaultdict
import streamlit as st
from planning_generator import (
    load_input, generate_planning, write_planning,
    validate_planning, last_in_rotation
)

st.set_page_config(page_title="Generador de planning de guàrdies", page_icon="🩺", layout="wide")

st.title("🩺 Generador de planning de guàrdies")
st.caption("Genera el planning mensual seguint les regles definides al fitxer d'entrada (Configuració + Radiòlegs + Vacances).")

with st.sidebar:
    st.header("Com fer servir")
    st.markdown("""
1. **Puja el fitxer d'entrada** (.xlsx amb les pestanyes Configuració, Radiòlegs i Vacances).
2. **Revisa la configuració** detectada.
3. **Puja la plantilla buida** del planning de sortida.
4. **Clica "Generar planning"**.
5. **Descarrega** el resultat i revisa-ho.
""")
    st.divider()
    st.markdown("**Plantilles necessàries:**")
    st.markdown("- `02_Plantilla_Entrada.xlsx` (omplerta amb els teus radiòlegs i vacances del mes)")
    st.markdown("- `03_Plantilla_Planning.xlsx` (la plantilla buida del planning)")
    st.markdown("Vegeu el README del paquet per a detalls.")

# Step 1: Upload input
st.header("Pas 1 — Fitxer d'entrada")
in_file = st.file_uploader(
    "Puja el fitxer Excel d'entrada del mes (Configuració + Radiòlegs + Vacances)",
    type=["xlsx"], key="in_file"
)

constraints = config = meta = None
if in_file:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as f:
            f.write(in_file.getvalue())
            in_path = f.name
        constraints, config, meta = load_input(in_path)
        st.success(f"Fitxer carregat. {len(constraints)} radiòlegs amb constriccions, {len(config.rotators)} a la roda.")
    except Exception as e:
        st.error(f"Error llegint el fitxer: {e}")
        import traceback
        st.code(traceback.format_exc())

# Step 2: Show config
if constraints and meta and config:
    st.header("Pas 2 — Configuració detectada")
    col1, col2, col3 = st.columns(3)
    month_names = {1: 'gener', 2: 'febrer', 3: 'març', 4: 'abril', 5: 'maig', 6: 'juny',
                   7: 'juliol', 8: 'agost', 9: 'setembre', 10: 'octubre', 11: 'novembre', 12: 'desembre'}
    with col1:
        st.metric("Mes/Any", f"{month_names[meta.month].capitalize()} {meta.year}")
    with col2:
        st.metric("Primer de la roda", meta.start_radiologist or "—")
    with col3:
        st.metric("Festius", ", ".join(str(d) for d in meta.festius) or "Cap")
    
    edit_col1, edit_col2, edit_col3 = st.columns(3)
    with edit_col1:
        new_month_str = st.text_input("Mes/Any (editable)", value=f"{month_names[meta.month].capitalize()} {meta.year}")
    with edit_col2:
        new_start = st.text_input("Primer radiòleg (editable)", value=meta.start_radiologist or "")
        meta.start_radiologist = new_start
    with edit_col3:
        new_festius = st.text_input("Festius (editable)", value=",".join(str(d) for d in meta.festius))
        try:
            meta.festius = [int(x.strip()) for x in new_festius.split(',') if x.strip()]
        except ValueError: pass

    with st.expander("📋 Veure rols i fixos detectats"):
        st.subheader("Torns fixos")
        DOW_NAMES = ['Dilluns','Dimarts','Dimecres','Dijous','Divendres','Dissabte','Diumenge']
        fix_rows = []
        for (dow, shift, role), name in sorted(config.fix_slots.items()):
            fix_rows.append({"Dia": DOW_NAMES[dow], "Franja": shift, "Rol": role, "Radiòleg": name})
        if config.first_wed_radiologist:
            fix_rows.append({"Dia": "1r Dimecres mes", "Franja": "16-20", "Rol": "N i B", "Radiòleg": config.first_wed_radiologist})
        if fix_rows: st.dataframe(fix_rows, hide_index=True, width='stretch')
        else: st.info("Cap torn fix definit.")
        
        st.subheader("Casos especials")
        st.write(f"**Sun-nit-only**: {', '.join(config.sun_nit_only) if config.sun_nit_only else 'cap'}")
        st.write(f"**Weekend-day-only**: {', '.join(config.weekend_day_only) if config.weekend_day_only else 'cap'}")
        st.write(f"**Fix-i-rota**: {', '.join(config.fix_and_rota) if config.fix_and_rota else 'cap'}")
        st.write(f"**Fix-només**: {', '.join(config.fix_only) if config.fix_only else 'cap'}")
        st.write(f"**Nou-incorporat** (cobreix dimecres 16-20): {', '.join(config.nou_incorporats) if config.nou_incorporats else 'cap'}")

# Step 3: Upload template
if constraints and meta and config:
    st.header("Pas 3 — Plantilla del planning")
    tmpl_file = st.file_uploader("Puja la plantilla buida del planning", type=["xlsx"], key="tmpl")
    
    if tmpl_file:
        st.success("Plantilla carregada.")
        
        st.header("Pas 4 — Generar planning")
        if st.button("🚀 Generar planning", type="primary", width='stretch'):
            with st.spinner("Calculant..."):
                try:
                    assignments, queue_final, warnings = generate_planning(constraints, config, meta)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as f:
                        tmp_tmpl = f.name
                    with open(tmp_tmpl, 'wb') as f:
                        f.write(tmpl_file.getvalue())
                    out_path = tmp_tmpl.replace('.xlsx', '_planning.xlsx')
                    write_planning(assignments, config, meta, constraints, tmp_tmpl, out_path)
                    errors = validate_planning(assignments, config, constraints, meta)
                    
                    st.divider()
                    st.subheader("📊 Resultats")
                    
                    if not errors:
                        st.success(f"✅ Validació OK — {len(assignments)} assignacions, sense conflictes.")
                    else:
                        st.error(f"❌ {len(errors)} problemes:")
                        for e in errors: st.markdown(f"- {e}")
                    
                    if warnings:
                        with st.expander("⚠️ Avisos"):
                            for w in warnings: st.markdown(f"- {w}")
                    
                    c1, c2, c3 = st.columns(3)
                    cnt = Counter(assignments.values())
                    dist = Counter(cnt.values())
                    with c1: st.metric("Total guàrdies", len(assignments))
                    with c2: st.metric("Persones amb 1", dist.get(1, 0))
                    with c3: st.metric("Persones amb 2+", sum(v for k, v in dist.items() if k >= 2))
                    
                    last = last_in_rotation(queue_final)
                    st.info(f"🔄 **Últim de la roda aquest mes:** {last}")
                    
                    with open(out_path, 'rb') as f:
                        st.download_button(
                            label="📥 Descarregar planning",
                            data=f.read(),
                            file_name=f"Planning_{meta.year}_{meta.month:02d}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch',
                        )
                    
                    with st.expander("📅 Veure totes les assignacions"):
                        by_day = defaultdict(list)
                        for slot, name in assignments.items():
                            by_day[slot[0]].append(f"{slot[2]} {slot[3]}: {name}")
                        DOW = ['Dl','Dt','Dc','Dj','Dv','Ds','Dg']
                        for day in sorted(by_day.keys()):
                            d = datetime.date(meta.year, meta.month, day)
                            st.markdown(f"**{day:02d}/{meta.month:02d} ({DOW[d.weekday()]})**")
                            for line in by_day[day]: st.text(f"  {line}")
                except Exception as e:
                    st.error(f"Error: {e}")
                    import traceback; st.code(traceback.format_exc())

st.divider()
st.caption("Algorisme v2.0 — config-driven, sense dades identificatives al codi.")
                           