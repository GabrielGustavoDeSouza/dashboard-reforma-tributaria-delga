"""
Dashboard Reforma Tributaria - Grupo Delga
Le arquivos MS Project (.mpp) diretamente, calcula KPIs e permite editar
% concluido, exportando um .xml (MSPDI) que abre nativamente no MS Project.
"""
import io
import os
import re
import json
import hashlib
import tempfile
import datetime as dt

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# ---------------------------------------------------------------------------
# CONFIG / BRANDING
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Reforma Tributária · Grupo Delga", layout="wide", page_icon="📊")

NAVY_DARK = "#0E0E20"
NAVY = "#000099"
BLUE = "#0C00FF"
BLUE_MID = "#9FAFFF"
BLUE_LIGHT = "#CCD7FF"
BG_CARD = "#F4F6FF"
MUTED = "#5C5C8A"

st.markdown(f"""
<style>
.block-container {{ padding-top: 1.5rem; }}
.kpi-card {{
    border-radius: 10px; padding: 1rem 1.2rem; height: 100%;
}}
.kpi-label {{ font-size: 0.72rem; letter-spacing: 1px; font-weight: 700; opacity: 0.85; }}
.kpi-value {{ font-size: 2rem; font-weight: 700; margin-top: 0.2rem; }}
.etapa-badge {{
    display:inline-block; border-radius:6px; padding:2px 8px; font-size:0.7rem;
    font-weight:700; color:white; background:{BLUE};
}}
.pending-tag {{
    display:inline-block; border-radius:6px; padding:2px 8px; font-size:0.72rem;
    font-style:italic; color:{MUTED}; background:#E7EAF7;
}}
</style>
""", unsafe_allow_html=True)

st.markdown(f"<h1 style='color:{NAVY_DARK};'>Reforma Tributária <span style='color:{BLUE};'>· Grupo Delga</span></h1>", unsafe_allow_html=True)
st.caption("Dashboard gerado a partir do arquivo MS Project (.mpp) — leitura direta, sem etapa manual de exportação.")

# ---------------------------------------------------------------------------
# ACESSO POR SENHA
# ---------------------------------------------------------------------------
APP_PASSWORD = "@DelgaRef2030"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown(f"<h2 style='color:{NAVY_DARK};'>Acesso restrito</h2>", unsafe_allow_html=True)
    st.caption("Digite a senha de acesso ao dashboard do Comitê da Reforma Tributária.")
    with st.form("login_form"):
        pwd = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar", type="primary")
    if entrar:
        if pwd == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()

# ---------------------------------------------------------------------------
# JVM / MPXJ INIT (uma vez por processo)
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def init_mpxj():
    import mpxj
    import glob
    import jpype
    if not jpype.isJVMStarted():
        jars = glob.glob(mpxj.mpxj_dir + "/*.jar")
        mpxj.startJVM(classpath=jars)
    from org.mpxj.reader import UniversalProjectReader
    from org.mpxj.writer import UniversalProjectWriter, FileFormat
    return UniversalProjectReader, UniversalProjectWriter, FileFormat


def read_project(file_bytes: bytes, suffix: str):
    UniversalProjectReader, _, _ = init_mpxj()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    reader = UniversalProjectReader()
    project = reader.read(tmp_path)
    return project


@st.cache_resource(show_spinner=False)
def get_project_cached(file_hash: str, _file_bytes: bytes, suffix: str):
    """Cache compartilhado por TODAS as sessoes/usuarios (nao e por-sessao),
    chaveado pelo hash do arquivo. Assim quem abrir o link nao precisa
    reprocessar nem reenviar nada."""
    return read_project(_file_bytes, suffix)


# ---------------------------------------------------------------------------
# PERSISTENCIA EM DISCO - o ultimo arquivo enviado fica fixo pra qualquer
# pessoa que abrir o link, sem precisar reenviar. Enquanto o app do Streamlit
# Cloud estiver "acordado" isso e permanente; se ele hibernar por inatividade
# longa ou for reiniciado, o disco volta ao zero e e preciso reenviar 1x.
# ---------------------------------------------------------------------------
DATA_DIR = "data"
META_PATH = os.path.join(DATA_DIR, "meta.json")
os.makedirs(DATA_DIR, exist_ok=True)


def persisted_file_path(suffix: str) -> str:
    return os.path.join(DATA_DIR, f"current_project{suffix}")


def save_persisted(file_bytes: bytes, suffix: str, original_name: str):
    # remove versao antiga com outra extensao, se existir
    for ext in (".mpp", ".xml"):
        p = persisted_file_path(ext)
        if os.path.exists(p):
            os.remove(p)
    with open(persisted_file_path(suffix), "wb") as f:
        f.write(file_bytes)
    with open(META_PATH, "w") as f:
        json.dump({
            "original_name": original_name,
            "suffix": suffix,
            "uploaded_at": dt.datetime.now().isoformat(timespec="minutes"),
        }, f)


def load_persisted():
    for ext in (".mpp", ".xml"):
        p = persisted_file_path(ext)
        if os.path.exists(p):
            with open(p, "rb") as f:
                data = f.read()
            meta = {}
            if os.path.exists(META_PATH):
                with open(META_PATH) as f:
                    meta = json.load(f)
            return data, ext, meta
    return None, None, None


def java_date_to_py(jdate):
    if jdate is None:
        return None
    # LocalDateTime tem getYear/getMonthValue/... no MPXJ (java.time)
    try:
        return dt.datetime(jdate.getYear(), jdate.getMonthValue(), jdate.getDayOfMonth(),
                            jdate.getHour(), jdate.getMinute())
    except Exception:
        return None


ETAPA_RE = re.compile(r"^\s*(\d+)\.\s*(.+)$")


def extract_tasks(project) -> pd.DataFrame:
    tasks = project.getTasks()
    rows = []
    current_etapa_num = None
    current_etapa_title = None

    for i in range(tasks.size()):
        t = tasks.get(i)
        if t is None:
            continue
        name = str(t.getName() or "").strip()
        level = int(t.getOutlineLevel()) if t.getOutlineLevel() is not None else 0
        is_summary = bool(t.getSummary())
        is_milestone = bool(t.getMilestone())
        start = java_date_to_py(t.getStart())
        finish = java_date_to_py(t.getFinish())
        pct = float(t.getPercentageComplete()) if t.getPercentageComplete() is not None else 0.0

        resources = []
        try:
            assigns = t.getResourceAssignments()
            for j in range(assigns.size()):
                r = assigns.get(j).getResource()
                if r is not None:
                    resources.append(str(r.getName()))
        except Exception:
            pass

        # detecta nova etapa de topo (outline level 2, nome comeca com "N. ")
        if level == 2 and is_summary:
            m = ETAPA_RE.match(name)
            if m:
                current_etapa_num = int(m.group(1))
                current_etapa_title = m.group(2).strip()

        rows.append({
            "id": int(t.getID()) if t.getID() is not None else i,
            "unique_id": int(t.getUniqueID()) if t.getUniqueID() is not None else i,
            "name": name,
            "level": level,
            "is_summary": is_summary,
            "is_milestone": is_milestone,
            "start": start,
            "finish": finish,
            "pct": pct,
            "resource": ", ".join(resources),
            "etapa_num": current_etapa_num,
            "etapa_title": current_etapa_title,
        })

    df = pd.DataFrame(rows)
    # remove as duas linhas de raiz (projeto + wrapper "Projeto: Reforma Tributaria")
    df = df[df["etapa_num"].notna()].copy()
    df["etapa_num"] = df["etapa_num"].astype(int)
    return df


def compute_planned_pct(row, today):
    if row["start"] is None or row["finish"] is None:
        return 0.0
    if today <= row["start"]:
        return 0.0
    if today >= row["finish"]:
        return 100.0
    total = (row["finish"] - row["start"]).total_seconds()
    elapsed = (today - row["start"]).total_seconds()
    if total <= 0:
        return 100.0
    return max(0.0, min(100.0, elapsed / total * 100.0))


TOLERANCIA_PP = 2.0  # pontos percentuais de tolerancia pro farol


def classify_status(diff: float):
    """diff = real_pct - planned_pct. Retorna (rotulo, cor)."""
    if diff < -TOLERANCIA_PP:
        return "atrasado", "#D64545"
    if diff > TOLERANCIA_PP:
        return "adiantado", BLUE
    return "em andamento", "#1B8A5A"


def build_leaves(df: pd.DataFrame, today: dt.datetime) -> pd.DataFrame:
    leaves = df[~df["is_summary"]].copy()
    leaves["planned_pct"] = leaves.apply(lambda r: compute_planned_pct(r, today), axis=1)
    leaves["diff_pp"] = leaves["pct"] - leaves["planned_pct"]
    return leaves


def build_etapa_summary(leaves: pd.DataFrame):
    etapas = []
    for num, g in leaves.groupby("etapa_num"):
        title = g["etapa_title"].iloc[0]
        unique_starts = g["start"].dropna().apply(lambda d: d.date()).nunique()
        tem_cronograma = unique_starts > 1
        real_pct = g["pct"].mean() if len(g) else 0.0
        planned_pct = g["planned_pct"].mean() if len(g) else 0.0
        diff_pp = real_pct - planned_pct
        status_label, status_color = classify_status(diff_pp)
        etapas.append({
            "etapa_num": num,
            "etapa_title": title,
            "tem_cronograma": tem_cronograma,
            "real_pct": real_pct,
            "planned_pct": planned_pct,
            "diff_pp": diff_pp,
            "status_label": status_label,
            "status_color": status_color,
            "n_tasks": len(g),
            "n_done": int((g["pct"] >= 100).sum()),
        })
    return pd.DataFrame(etapas).sort_values("etapa_num")


# ---------------------------------------------------------------------------
# UI - UPLOAD (com persistencia: o arquivo enviado fica valendo pra todo
# mundo que abrir o link, ate alguem enviar um novo)
# ---------------------------------------------------------------------------
persisted_bytes, persisted_suffix, persisted_meta = load_persisted()

with st.sidebar:
    st.markdown("### Arquivo do projeto")
    if persisted_meta:
        st.success(f"Arquivo atual: **{persisted_meta['original_name']}**\n\nEnviado em {persisted_meta['uploaded_at'].replace('T', ' ')}")
    uploaded = st.file_uploader(
        "Enviar novo .mpp (ou .xml exportado do Project)" if persisted_meta else "Envie o .mpp (ou .xml exportado do Project)",
        type=["mpp", "xml"],
    )
    st.caption("O arquivo enviado aqui fica salvo para qualquer pessoa que abrir este link — não precisa reenviar toda vez, só quando quiser atualizar os dados.")
    today_override = st.date_input("Data de referência (\"hoje\")", value=dt.date.today())
    st.divider()
    st.caption("Etapas sem variação de data entre as tarefas são tratadas como **sem cronograma definido** e não entram nos KPIs — assim que você lançar datas reais, elas passam a contar automaticamente.")

if uploaded is not None:
    # novo envio: persiste em disco e passa a valer pra todo mundo
    suffix = ".mpp" if uploaded.name.lower().endswith(".mpp") else ".xml"
    file_bytes = uploaded.getvalue()
    save_persisted(file_bytes, suffix, uploaded.name)
    original_name = uploaded.name
elif persisted_bytes is not None:
    # ninguem enviou nada nesta sessao: usa o que ja esta salvo
    file_bytes = persisted_bytes
    suffix = persisted_suffix
    original_name = persisted_meta["original_name"]
else:
    st.info("Envie o arquivo .mpp na barra lateral para gerar o dashboard.")
    st.stop()

file_hash = hashlib.md5(file_bytes).hexdigest()
with st.spinner("Lendo o arquivo do MS Project..."):
    project = get_project_cached(file_hash, file_bytes, suffix)

df = extract_tasks(project)
today_dt = dt.datetime.combine(today_override, dt.time(12, 0))
leaves_df = build_leaves(df, today_dt)
etapa_df = build_etapa_summary(leaves_df)

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
tracked = etapa_df[(etapa_df["etapa_num"] != 0) & (etapa_df["tem_cronograma"])]
progresso_total = tracked["real_pct"].mean() if len(tracked) else 0.0
planned_total = tracked["planned_pct"].mean() if len(tracked) else 0.0
# SPI aqui = variacao em pontos percentuais (Real - Previsto), nao a razao classica do EVM.
# 0 = em dia; negativo = atrasado; positivo = adiantado frente ao cronograma.
spi = progresso_total - planned_total

# marco critico: entre as etapas 1 a 6 (0 = so referencia legal, 7 e 8 ainda
# sem cronograma e nao entram), prioriza a etapa 5 (frente critica sinalizada
# pela diretoria) e, dentro dela, a data mais proxima; se a etapa 5 nao tiver
# pendencia com data, cai para a proxima pendencia mais proxima entre 1 e 6.
tracked_set = set(etapa_df[etapa_df["tem_cronograma"]]["etapa_num"])
candidatos = df[
    (df["etapa_num"].between(1, 6))
    & (df["etapa_num"].isin(tracked_set))
    & (~df["is_summary"])
    & (df["pct"] < 100)
    & (df["start"].notna())
].copy()
candidatos["prioridade"] = candidatos["etapa_num"].apply(lambda n: 0 if n == 5 else 1)
candidatos = candidatos.sort_values(["prioridade", "start"])

marco_txt = "—"
marco_etapa_txt = ""
if len(candidatos):
    m0 = candidatos.iloc[0]
    marco_txt = f"{m0['name'].strip()} — {m0['start'].strftime('%d/%m/%Y')}"
    marco_etapa_txt = f"etapa {int(m0['etapa_num'])}"

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f"""<div class="kpi-card" style="background:{NAVY_DARK};color:white;">
        <div class="kpi-label" style="color:{BLUE_LIGHT};">PROGRESSO TOTAL DO PROJETO</div>
        <div class="kpi-value">{progresso_total:.0f}%</div>
        <div style="font-size:0.72rem;opacity:0.75;">exclui etapas sem cronograma definido</div>
        </div>""", unsafe_allow_html=True)
with c2:
    if planned_total < 5:
        spi_txt = "—"
        spi_note = "poucas tarefas com previsão vencida ainda para medir"
        spi_color = MUTED
        farol_dot = ""
    else:
        sinal = "+" if spi > 0 else ""
        spi_txt = f"{sinal}{spi:.0f}%"
        spi_note = "real − previsto pela data · negativo = atrasado, positivo = adiantado"
        spi_label, spi_dot_color = classify_status(spi)
        spi_color = BLUE if spi < 0 else NAVY_DARK
        farol_dot = f'''<span title="{spi_label}" style="display:inline-block;width:10px;height:10px;
            border-radius:50%;background:{spi_dot_color};margin-right:8px;vertical-align:middle;"></span>'''
    st.markdown(f"""<div class="kpi-card" style="background:{BG_CARD};border:1px solid {BLUE_MID};">
        <div class="kpi-label" style="color:{MUTED};">{farol_dot}SPI (VARIAÇÃO DE PRAZO)</div>
        <div class="kpi-value" style="color:{spi_color};">{spi_txt}</div>
        <div style="font-size:0.72rem;color:{MUTED};">{spi_note}</div>
        </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="kpi-card" style="background:{BLUE_LIGHT};border:1px solid {BLUE};">
        <div class="kpi-label" style="color:{NAVY};">MARCO CRÍTICO{' · ' + marco_etapa_txt if marco_etapa_txt else ''}</div>
        <div class="kpi-value" style="font-size:1.15rem;color:{NAVY_DARK};">{marco_txt}</div>
        </div>""", unsafe_allow_html=True)

st.write("")

# ---------------------------------------------------------------------------
# ATIVIDADES EM ATRASO
# Estritamente: tarefas cuja data de TERMINO ja passou e que nao estao 100%
# concluidas. O farol de +-2pp e so pro indicador geral (SPI / status por
# etapa) - aqui e "venceu o prazo e nao terminou", ponto, sem tolerancia de
# ritmo intermediario.
# ---------------------------------------------------------------------------
tracked_set_top = set(etapa_df[etapa_df["tem_cronograma"]]["etapa_num"])
atrasadas = leaves_df[
    (leaves_df["etapa_num"] != 0)
    & (leaves_df["etapa_num"].isin(tracked_set_top))
    & (leaves_df["finish"].notna())
    & (leaves_df["finish"] <= today_dt)
    & (leaves_df["pct"] < 100)
].copy()
atrasadas["dias_atraso"] = (today_dt - atrasadas["finish"]).dt.days
atrasadas = atrasadas.sort_values("dias_atraso", ascending=False)

if len(atrasadas):
    st.subheader(f"🔴 Atividades em atraso ({len(atrasadas)})")
    st.caption("Prazo de término já vencido e ainda não concluídas — não é ritmo, é prazo estourado.")
    rows_html = []
    for _, r in atrasadas.head(15).iterrows():
        real = min(max(r["pct"], 0), 100)
        planned = min(max(r["planned_pct"], 0), 100)
        venc_txt = r["finish"].strftime("%d/%m/%Y") if pd.notna(r["finish"]) else "—"
        dias = int(r["dias_atraso"])
        rows_html.append(f'''
        <div style="display:grid;grid-template-columns:1.1fr 0.6fr 3fr 3.4fr 1.3fr;gap:14px;align-items:center;
            padding:8px 0;border-bottom:1px solid #EEF0FA;">
            <div style="font-size:0.82rem;color:#D64545;">venceu {venc_txt}<br><span style="font-size:0.72rem;">({dias}d atrás)</span></div>
            <div><span class="etapa-badge">{int(r["etapa_num"])}</span></div>
            <div style="font-size:0.88rem;color:{NAVY_DARK};">{r["name"].strip()}</div>
            <div style="position:relative;background:#F3D8D8;border-radius:8px;height:14px;">
                <div style="position:absolute;left:0;top:0;background:#D64545;height:14px;border-radius:8px;width:{real:.1f}%;"></div>
            </div>
            <div style="text-align:right;font-size:0.82rem;">
                <span style="font-weight:700;color:#D64545;">{r["pct"]:.0f}%</span>
                <span style="color:{MUTED};"> concluído</span>
            </div>
        </div>''')
    st.markdown("".join(rows_html), unsafe_allow_html=True)
    st.caption("Ordenado pelas mais atrasadas primeiro (dias corridos desde o vencimento).")
else:
    st.success(f"Nenhuma atividade fora da tolerância de ±{TOLERANCIA_PP:.0f}pp em relação ao previsto.")

st.write("")

# ---------------------------------------------------------------------------
# STATUS POR ETAPA
# ---------------------------------------------------------------------------
st.subheader("Status por etapa")

etapa_rows_html = []
for _, row in etapa_df.iterrows():
    n = int(row["etapa_num"])
    title = row["etapa_title"]
    is_critical_stage = "CRITICO" in title.upper() or "CRÍTICO" in title.upper()
    crit_dot = f'<span style="color:{BLUE};margin-left:6px;">●</span>' if is_critical_stage else ""

    if n == 0:
        bar_html = f'''<div style="background:#E7EAF7;border-radius:8px;height:14px;width:100%;position:relative;">
            <div style="background:{MUTED};opacity:0.35;height:14px;border-radius:8px;width:100%;"></div>
        </div>'''
        pct_html = f'<span style="color:{MUTED};font-size:0.85rem;">referência</span>'
        farol_html = ""
    elif not row["tem_cronograma"]:
        bar_html = f'''<div style="background:#E7EAF7;border-radius:8px;height:14px;width:100%;"></div>'''
        pct_html = f'<span class="pending-tag">aguardando cronograma</span>'
        farol_html = ""
    else:
        pct = min(row["real_pct"], 100)
        bar_color = BLUE if is_critical_stage else BLUE_MID
        bar_html = f'''<div style="background:#E7EAF7;border-radius:8px;height:14px;width:100%;">
            <div style="background:{bar_color};height:14px;border-radius:8px;width:{pct:.1f}%;"></div>
        </div>'''
        pct_html = f'<span style="font-weight:700;color:{NAVY_DARK};">{row["real_pct"]:.0f}%</span>'
        farol_html = f'''<span title="{row['status_label']}" style="display:inline-block;width:9px;height:9px;
            border-radius:50%;background:{row['status_color']};margin-right:8px;"></span>'''

    etapa_rows_html.append(f'''
    <div style="display:grid;grid-template-columns:2.6fr 5fr 1fr;gap:16px;align-items:center;padding:9px 0;border-bottom:1px solid #EEF0FA;">
        <div style="font-size:0.92rem;color:{NAVY_DARK};">{farol_html}<b>{n}.</b> {title}{crit_dot}</div>
        <div>{bar_html}</div>
        <div style="text-align:right;">{pct_html}</div>
    </div>''')

st.markdown(f'<div style="margin-top:4px;">{"".join(etapa_rows_html)}</div>', unsafe_allow_html=True)
st.caption(f"🟢 em andamento (±{TOLERANCIA_PP:.0f}pp)  ·  🔴 atrasado  ·  🔵 adiantado")

st.write("")

# ---------------------------------------------------------------------------
# ATIVIDADES RELEVANTES
# ---------------------------------------------------------------------------
st.subheader("Atividades relevantes — próximas e em andamento")

leaves = df[(~df["is_summary"]) & (df["etapa_num"] != 0)].copy()
tracked_etapas = set(etapa_df[etapa_df["tem_cronograma"]]["etapa_num"])
leaves = leaves[leaves["etapa_num"].isin(tracked_etapas)]
leaves = leaves[leaves["pct"] < 100]
leaves = leaves.sort_values("start").head(10)

if len(leaves):
    header_html = f'''<div style="display:grid;grid-template-columns:1.1fr 0.7fr 3.4fr 1.4fr 1fr;
        background:{NAVY_DARK};color:white;border-radius:8px 8px 0 0;padding:9px 14px;font-size:0.78rem;font-weight:700;letter-spacing:0.3px;">
        <div>DATA</div><div>ETAPA</div><div>ATIVIDADE</div><div>RESPONSÁVEL</div><div style="text-align:right;">% CONCLUÍDO</div>
    </div>'''
    body_rows = []
    for i, (_, r) in enumerate(leaves.iterrows()):
        bg = BG_CARD if i % 2 == 0 else "white"
        data_txt = r["start"].strftime("%d/%m/%Y") if pd.notna(r["start"]) else "a definir"
        body_rows.append(f'''<div style="display:grid;grid-template-columns:1.1fr 0.7fr 3.4fr 1.4fr 1fr;
            background:{bg};padding:9px 14px;font-size:0.86rem;border-bottom:1px solid #EEF0FA;align-items:center;">
            <div style="color:{MUTED};">{data_txt}</div>
            <div><span class="etapa-badge">{int(r["etapa_num"])}</span></div>
            <div style="color:{NAVY_DARK};">{r["name"].strip()}</div>
            <div style="color:{MUTED};">{r["resource"] or "—"}</div>
            <div style="text-align:right;font-weight:700;color:{NAVY_DARK};">{r["pct"]:.0f}%</div>
        </div>''')
    st.markdown(header_html + "".join(body_rows) + '<div style="border-radius:0 0 8px 8px;overflow:hidden;"></div>', unsafe_allow_html=True)
else:
    st.caption("Nenhuma atividade com cronograma definido no momento.")

st.write("")

# ---------------------------------------------------------------------------
# EDICAO DE % CONCLUIDO + EXPORTACAO PARA MS PROJECT
# ---------------------------------------------------------------------------
st.subheader("Atualizar % concluído")
st.caption("Edite a coluna **% concluído** abaixo e gere um arquivo atualizado para reabrir no MS Project.")

editable = df[~df["is_summary"]][["unique_id", "id", "etapa_num", "name", "start", "pct"]].copy()
editable["name"] = editable["name"].str.strip()
editable = editable.sort_values(["etapa_num", "id"])
editable_display = editable.rename(columns={
    "etapa_num": "Etapa", "name": "Atividade", "start": "Início", "pct": "% concluído"
})

edited = st.data_editor(
    editable_display[["Etapa", "Atividade", "Início", "% concluído"]],
    hide_index=True,
    use_container_width=True,
    disabled=["Etapa", "Atividade", "Início"],
    column_config={
        "% concluído": st.column_config.NumberColumn(min_value=0, max_value=100, step=5)
    },
    height=400,
)

col_a, col_b = st.columns([1, 3])
with col_a:
    gerar = st.button("Gerar arquivo atualizado (.xml)", type="primary")

if gerar:
    _, UniversalProjectWriter, FileFormat = init_mpxj()
    import jpype

    # importante: usa uma copia independente do projeto pra gerar o .xml,
    # em vez do objeto "project" (que e compartilhado com todas as sessoes
    # que estao vendo o dashboard ao vivo). Assim, editar aqui pra exportar
    # nao muda o que outras pessoas veem ate voce de fato reenviar o arquivo
    # atualizado na barra lateral.
    edit_project = read_project(file_bytes, suffix)

    changed = 0
    for idx, (uid, new_pct) in enumerate(zip(editable["unique_id"], edited["% concluído"])):
        old_pct = editable["pct"].iloc[idx]
        if new_pct != old_pct:
            task = edit_project.getTaskByUniqueID(jpype.java.lang.Integer(int(uid)))
            if task is not None:
                task.setPercentageComplete(jpype.java.lang.Double(float(new_pct)))
                changed += 1

    out_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xml").name
    writer = UniversalProjectWriter(FileFormat.MSPDI)
    writer.write(edit_project, out_path)

    with open(out_path, "rb") as f:
        xml_bytes = f.read()

    st.success(f"{changed} tarefa(s) atualizada(s). Baixe o arquivo abaixo, confira no MS Project e, quando quiser que o dashboard de todo mundo reflita essa versão, reenvie esse mesmo .xml na barra lateral.")
    st.download_button(
        "Baixar .xml atualizado",
        data=xml_bytes,
        file_name=f"Reforma_Tributaria_atualizado_{dt.date.today().isoformat()}.xml",
        mime="application/xml",
    )
