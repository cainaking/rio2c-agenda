"""
Web app para a equipe atualizar a agenda da Sala Transforma.
Rodar com: streamlit run app.py
"""
import streamlit as st
import os
import io
import zipfile
from datetime import datetime
import engine

st.set_page_config(page_title="Sala Transforma — Rio2C 2026", page_icon="🎬", layout="wide")

# ============== HEADER ==============
st.markdown("""
<style>
.big-title { font-size: 32px; font-weight: 700; color: #1F4E79; margin-bottom: 0; }
.subtitle { font-size: 14px; color: #666; margin-top: 0; }
.stButton>button { background-color: #1F4E79; color: white; border: none; padding: 10px 20px; font-weight: 600; }
.stButton>button:hover { background-color: #2E75B6; color: white; }
.success-box { padding: 15px; border-radius: 8px; background: #E2EFDA; border-left: 4px solid #548235; }
.warning-box { padding: 15px; border-radius: 8px; background: #FCE4D6; border-left: 4px solid #C65911; }
</style>
<p class="big-title">🎬 Sala Transforma — Rio2C 2026</p>
<p class="subtitle">Painel de atualização da agenda de matchmaking</p>
<hr/>
""", unsafe_allow_html=True)

# ============== HELPERS ==============
@st.cache_data(ttl=10)
def get_companies():
    return engine.get_all_companies()

@st.cache_data(ttl=10)
def get_people():
    return engine.get_all_people()

def reload_caches():
    get_companies.clear()
    get_people.clear()

def show_overrides_summary():
    o = engine.load_overrides()
    cols = st.columns(5)
    cols[0].metric("Bloqueios", len(o['bloqueios_horario']))
    cols[1].metric("Matches extras", len(o['matches_extras']))
    cols[2].metric("Matches cancelados", len(o['matches_cancelados']))
    cols[3].metric("Correções de empresa", len(o['correcoes_empresa']))
    cols[4].metric("Participantes novos", len(o['participantes_novos']))

# ============== SIDEBAR — REGERAR ==============
with st.sidebar:
    st.markdown("### 🔄 Regerar Agenda")
    st.caption("Aplica todas as atualizações pendentes e gera os arquivos novos.")
    if st.button("🚀 Regerar Tudo Agora", type="primary", use_container_width=True):
        with st.spinner("Recalculando agendamento... isso leva ~30 segundos."):
            try:
                result = engine.run_full()
                st.session_state['last_result'] = result
                st.session_state['last_run'] = datetime.now().strftime('%H:%M:%S')
                st.success(f"✅ Agendamento concluído! {result['scheduled']}/{result['total_matches']} matches agendados.")
                if result['unscheduled']:
                    st.warning(f"⚠️ {len(result['unscheduled'])} match(es) não couberam — veja a aba Resultado.")
            except Exception as e:
                st.error(f"Erro: {e}")
                import traceback; st.code(traceback.format_exc())

    if 'last_run' in st.session_state:
        st.caption(f"Última execução: {st.session_state['last_run']}")

    st.markdown("---")
    st.markdown("### 📥 Baixar Arquivos")
    if 'last_result' in st.session_state:
        result = st.session_state['last_result']

        # Download agenda geral
        with open(result['output_general'], 'rb') as f:
            st.download_button("📋 Agenda da Equipe (limpa)", f.read(),
                file_name=os.path.basename(result['output_general']),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        # Download agenda detalhada (com estatísticas, carga, etc.)
        with open(result['output_detalhado'], 'rb') as f:
            st.download_button("📊 Agenda Detalhada (com stats/carga)", f.read(),
                file_name=os.path.basename(result['output_detalhado']),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        # Download zip de todas as individuais
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fname in os.listdir(result['output_dir']):
                fpath = os.path.join(result['output_dir'], fname)
                if os.path.isfile(fpath):
                    zf.write(fpath, fname)
        st.download_button("📦 Todas as Agendas (ZIP)", zip_buf.getvalue(),
            file_name=f"agendas_individuais_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip", use_container_width=True)
    else:
        st.info("Aperte 'Regerar' para gerar os arquivos.")

    st.markdown("---")
    st.markdown("### 📊 Status atual")
    show_overrides_summary()

# ============== TABS PRINCIPAIS ==============
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🚫 Bloquear Horário",
    "➕ Adicionar Match",
    "❌ Cancelar Match",
    "🔧 Corrigir Empresa",
    "👤 Novo Participante",
    "📋 Resultado",
])

# ============== TAB 1: BLOQUEAR HORÁRIO ==============
with tab1:
    st.subheader("🚫 Bloquear horário de uma empresa")
    st.caption("Use quando alguém avisar que não pode mais em determinado horário. "
              "O sistema vai realocar a reunião automaticamente quando regerar.")

    companies = get_companies()
    with st.form("form_bloquear"):
        c1, c2 = st.columns(2)
        empresa = c1.selectbox("Empresa", options=companies, key="bloq_empresa")
        pessoa = c2.text_input("Pessoa específica (opcional)", help="Deixe vazio para bloquear todos da empresa")
        c3, c4, c5 = st.columns(3)
        dia = c3.selectbox("Dia", options=engine.DAYS)
        slot = c4.selectbox("Horário", options=engine.SLOTS)
        motivo = c5.text_input("Motivo (opcional)")
        if st.form_submit_button("✅ Bloquear", type="primary"):
            o = engine.load_overrides()
            o['bloqueios_horario'].append({
                'empresa': empresa, 'pessoa': pessoa, 'dia': dia,
                'slot': slot, 'motivo': motivo,
            })
            engine.save_overrides(o)
            reload_caches()
            st.success(f"✅ Bloqueio adicionado: {empresa} — {dia} {slot}. Aperte 'Regerar' na barra lateral.")

    # Lista de bloqueios atuais
    o = engine.load_overrides()
    if o['bloqueios_horario']:
        st.markdown("---")
        st.markdown("##### Bloqueios atuais")
        for i, b in enumerate(o['bloqueios_horario']):
            cols = st.columns([5, 1])
            txt = f"**{b['empresa']}**"
            if b.get('pessoa'): txt += f" ({b['pessoa']})"
            txt += f" — {b['dia']} {b['slot']}"
            if b.get('motivo'): txt += f"  _• {b['motivo']}_"
            cols[0].markdown(txt)
            if cols[1].button("🗑️", key=f"del_bloq_{i}"):
                o['bloqueios_horario'].pop(i)
                engine.save_overrides(o); st.rerun()

# ============== TAB 2: ADICIONAR MATCH ==============
with tab2:
    st.subheader("➕ Adicionar novo match")
    st.caption("Use quando aprovarem uma reunião nova entre duas empresas.")

    companies = get_companies()
    with st.form("form_match_novo"):
        c1, c2 = st.columns(2)
        emp_a = c1.selectbox("Empresa A", options=companies, key="ma_a")
        emp_b = c2.selectbox("Empresa B", options=companies, key="ma_b")
        c3, c4 = st.columns(2)
        tier = c3.selectbox("Tier (prioridade)", options=['Tier 1', 'Tier 2', 'Tier 3'], index=2)
        score = c4.number_input("Score", min_value=0, max_value=200, value=50)
        if st.form_submit_button("✅ Adicionar match", type="primary"):
            if emp_a == emp_b:
                st.error("Empresa A e B não podem ser a mesma.")
            else:
                o = engine.load_overrides()
                o['matches_extras'].append({
                    'empresa_a': emp_a, 'empresa_b': emp_b,
                    'reps_a': [], 'reps_b': [],
                    'tier': tier, 'score': score,
                })
                engine.save_overrides(o)
                st.success(f"✅ Match adicionado: {emp_a} × {emp_b}. Aperte 'Regerar' na barra lateral.")

    o = engine.load_overrides()
    if o['matches_extras']:
        st.markdown("---")
        st.markdown("##### Matches adicionados manualmente")
        for i, m in enumerate(o['matches_extras']):
            cols = st.columns([5, 1])
            cols[0].markdown(f"**{m['empresa_a']}** × **{m['empresa_b']}** ({m.get('tier', 'Tier 3')})")
            if cols[1].button("🗑️", key=f"del_extra_{i}"):
                o['matches_extras'].pop(i)
                engine.save_overrides(o); st.rerun()

# ============== TAB 3: CANCELAR MATCH ==============
with tab3:
    st.subheader("❌ Cancelar um match aprovado")
    st.caption("Use quando uma reunião for cancelada — ela não será mais agendada.")

    companies = get_companies()
    with st.form("form_cancelar"):
        c1, c2 = st.columns(2)
        emp_a = c1.selectbox("Empresa A", options=companies, key="cancel_a")
        emp_b = c2.selectbox("Empresa B", options=companies, key="cancel_b")
        if st.form_submit_button("✅ Cancelar match", type="primary"):
            o = engine.load_overrides()
            o['matches_cancelados'].append({'empresa_a': emp_a, 'empresa_b': emp_b})
            engine.save_overrides(o)
            st.success(f"✅ Match cancelado: {emp_a} × {emp_b}. Aperte 'Regerar' na barra lateral.")

    o = engine.load_overrides()
    if o['matches_cancelados']:
        st.markdown("---")
        st.markdown("##### Matches cancelados")
        for i, m in enumerate(o['matches_cancelados']):
            cols = st.columns([5, 1])
            cols[0].markdown(f"**{m['empresa_a']}** × **{m['empresa_b']}**")
            if cols[1].button("↩️ Reativar", key=f"del_cancel_{i}"):
                o['matches_cancelados'].pop(i)
                engine.save_overrides(o); st.rerun()

# ============== TAB 4: CORRIGIR EMPRESA ==============
with tab4:
    st.subheader("🔧 Corrigir empresa de uma pessoa")
    st.caption("Use quando uma pessoa estiver listada na empresa errada (ex: Tiago Campany na Globoplay quando ele é da Globo Internacional).")

    people = get_people()
    companies = get_companies()
    with st.form("form_correcao"):
        nome_options = sorted(set(p[0] for p in people))
        nome = st.selectbox("Nome da pessoa", options=nome_options)
        c1, c2 = st.columns(2)
        emp_atual = c1.selectbox("Empresa atual (errada)", options=companies)
        emp_nova = c2.text_input("Empresa correta", placeholder="Ex: Globo Internacional")
        if st.form_submit_button("✅ Corrigir", type="primary"):
            if not emp_nova.strip():
                st.error("Digite a empresa correta.")
            else:
                o = engine.load_overrides()
                o['correcoes_empresa'].append({
                    'nome': nome.lower(), 'empresa_antiga': emp_atual, 'empresa_nova': emp_nova.strip(),
                })
                engine.save_overrides(o); reload_caches()
                st.success(f"✅ Correção aplicada: {nome} → {emp_nova}. Aperte 'Regerar' na barra lateral.")

    o = engine.load_overrides()
    if o['correcoes_empresa']:
        st.markdown("---")
        st.markdown("##### Correções aplicadas")
        for i, c in enumerate(o['correcoes_empresa']):
            cols = st.columns([5, 1])
            cols[0].markdown(f"**{c['nome'].title()}**: {c['empresa_antiga']} → **{c['empresa_nova']}**")
            if cols[1].button("🗑️", key=f"del_corr_{i}"):
                o['correcoes_empresa'].pop(i)
                engine.save_overrides(o); reload_caches(); st.rerun()

# ============== TAB 5: NOVO PARTICIPANTE ==============
with tab5:
    st.subheader("👤 Adicionar participante novo")
    st.caption("Use quando alguém não estiver na planilha original.")

    with st.form("form_novo_p"):
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome completo")
        empresa = c2.text_input("Empresa")
        c3, c4 = st.columns(2)
        cargo = c3.text_input("Cargo")
        email = c4.text_input("E-mail")
        obs = st.text_input("Observação (opcional)")

        st.markdown("**Marque os horários em que a pessoa NÃO pode:**")
        indisp_slots = []
        for day in engine.DAYS:
            cols = st.columns(len(engine.SLOTS) + 1)
            cols[0].markdown(f"**{day}**")
            for j, slot in enumerate(engine.SLOTS):
                if cols[j+1].checkbox(slot, key=f"{day}_{slot}"):
                    indisp_slots.append([day, slot])

        if st.form_submit_button("✅ Adicionar participante", type="primary"):
            if not nome.strip() or not empresa.strip():
                st.error("Nome e Empresa são obrigatórios.")
            else:
                o = engine.load_overrides()
                o['participantes_novos'].append({
                    'nome': nome.strip(), 'empresa': empresa.strip(),
                    'cargo': cargo.strip(), 'email': email.strip(), 'obs': obs.strip(),
                    'indisp_slots': indisp_slots,
                })
                engine.save_overrides(o); reload_caches()
                st.success(f"✅ Participante adicionado: {nome} ({empresa}). Aperte 'Regerar' na barra lateral.")

    o = engine.load_overrides()
    if o['participantes_novos']:
        st.markdown("---")
        st.markdown("##### Participantes adicionados")
        for i, p in enumerate(o['participantes_novos']):
            cols = st.columns([5, 1])
            txt = f"**{p['nome']}** ({p['empresa']})"
            if p.get('cargo'): txt += f" — _{p['cargo']}_"
            if p.get('indisp_slots'): txt += f"  • {len(p['indisp_slots'])} slot(s) bloqueado(s)"
            cols[0].markdown(txt)
            if cols[1].button("🗑️", key=f"del_p_{i}"):
                o['participantes_novos'].pop(i)
                engine.save_overrides(o); reload_caches(); st.rerun()

# ============== TAB 6: RESULTADO ==============
with tab6:
    st.subheader("📋 Resultado da última execução")
    if 'last_result' not in st.session_state:
        st.info("👈 Aperte 'Regerar' na barra lateral para gerar a agenda.")
    else:
        r = st.session_state['last_result']

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Matches totais", r['total_matches'])
        c2.metric("✅ Agendados", r['scheduled'])
        c3.metric("❌ Não agendados", len(r['unscheduled']))
        c4.metric("Empresas", r['companies_with_meetings'])

        if r['unscheduled']:
            st.markdown("### ⚠️ Matches não agendados")
            for m in r['unscheduled']:
                st.markdown(f"- **{m['company_a']}** × **{m['company_b']}** "
                           f"(Tier: {m['tier']}, Score: {m['score']})")
            st.caption("Estes matches não couberam por falta de horários compatíveis. "
                      "Considere remover bloqueios ou checar disponibilidade dos envolvidos.")

        # Visão por dia
        st.markdown("### 📅 Distribuição por dia")
        for day in engine.DAYS:
            with st.expander(f"**{day}** — {sum(1 for m in r['scheduled_matches'] if m['day'] == day)} reuniões"):
                for slot in engine.SLOTS:
                    meetings = sorted(r['schedule'].get((day, slot), []), key=lambda x: x['company_a'].lower())
                    if meetings:
                        st.markdown(f"**{slot}** ({len(meetings)}/{engine.TABLES_PER_SLOT} mesas)")
                        for idx, m in enumerate(meetings, 1):
                            st.markdown(f"&nbsp;&nbsp;&nbsp;Mesa {idx}: **{m['company_a']}** × **{m['company_b']}**", unsafe_allow_html=True)
