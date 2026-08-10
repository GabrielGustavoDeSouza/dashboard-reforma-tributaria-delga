"""
Dashboard Reforma Tributaria - Grupo Delga
Le arquivos MS Project (.mpp) diretamente, calcula KPIs e permite editar
% concluido, exportando um .xml (MSPDI) que abre nativamente no MS Project.
"""
import io
import re
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


def build_etapa_summary(df: pd.DataFrame, today: dt.datetime):
    leaves = df[~df["is_summary"]].copy()
    leaves["planned_pct"] = leaves.apply(lambda r: compute_planned_pct(r, today), axis=1)

    etapas = []
    for num, g in leaves.groupby("etapa_num"):
        title = g["etapa_title"].iloc[0]
        unique_starts = g["start"].dropna().apply(lambda d: d.date()).nunique()
        tem_cronograma = unique_starts > 1
        real_pct = g["pct"].mean() if len(g) else 0.0
        planned_pct = g["planned_pct"].mean() if len(g) else 0.0
        etapas.append({
            "etapa_num": num,
            "etapa_title": title,
            "tem_cronograma": tem_cronograma,
            "real_pct": real_pct,
            "planned_pct": planned_pct,
            "n_tasks": len(g),
            "n_done": int((g["pct"] >= 100).sum()),
        })
    return pd.DataFrame(etapas).sort_values("etapa_num")


# ---------------------------------------------------------------------------
# UI - UPLOAD
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Arquivo do projeto")
    uploaded = st.file_uploader("Envie o .mpp (ou .xml exportado do Project)", type=["mpp", "xml"])
    st.caption("O arquivo é lido diretamente — nenhuma conversão manual necessária.")
    today_override = st.date_input("Data de referência (\"hoje\")", value=dt.date.today())
    st.divider()
    st.caption("Etapas sem variação de data entre as tarefas são tratadas como **sem cronograma definido** e não entram nos KPIs — assim que você lançar datas reais, elas passam a contar automaticamente.")

if uploaded is None:
    st.info("Envie o arquivo .mpp na barra lateral para gerar o dashboard.")
    st.stop()

suffix = ".mpp" if uploaded.name.lower().endswith(".mpp") else ".xml"
file_bytes = uploaded.getvalue()

if ("project_obj" not in st.session_state) or (st.session_state.get("last_file_name") != uploaded.name) or (st.session_state.get("last_file_size") != len(file_bytes)):
    with st.spinner("Lendo o arquivo do MS Project..."):
        project = read_project(file_bytes, suffix)
        st.session_state["project_obj"] = project
        st.session_state["last_file_name"] = uploaded.name
        st.session_state["last_file_size"] = len(file_bytes)

project = st.session_state["project_obj"]
df = extract_tasks(project)
today_dt = dt.datetime.combine(today_override, dt.time(12, 0))
etapa_df = build_etapa_summary(df, today_dt)

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
tracked = etapa_df[(etapa_df["etapa_num"] != 0) & (etapa_df["tem_cronograma"])]
progresso_total = tracked["real_pct"].mean() if len(tracked) else 0.0
planned_total = tracked["planned_pct"].mean() if len(tracked) else 0.0
spi = (progresso_total / planned_total * 100) if planned_total > 0 else None

# marco critico: proxima entrega da etapa "0. Marcos Legais" ainda incompleta,
# por data de inicio (nao depende mais da flag Milestone, pois o usuario pode
# ter dado duracao real a essas tarefas dentro do Project)
marcos = df[(df["etapa_num"] == 0) & (~df["is_summary"]) & (df["pct"] < 100) & (df["start"].notna())]
marcos = marcos.sort_values("start")
marco_txt = "—"
if len(marcos):
    m0 = marcos.iloc[0]
    marco_txt = f"{m0['name'].strip()} — {m0['start'].strftime('%d/%m/%Y')}"

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f"""<div class="kpi-card" style="background:{NAVY_DARK};color:white;">
        <div class="kpi-label" style="color:{BLUE_LIGHT};">PROGRESSO TOTAL DO PROJETO</div>
        <div class="kpi-value">{progresso_total:.0f}%</div>
        <div style="font-size:0.72rem;opacity:0.75;">exclui etapas sem cronograma definido</div>
        </div>""", unsafe_allow_html=True)
with c2:
    spi_txt = "—" if spi is None else f"{spi:.0f}%"
    spi_color = BLUE if (spi is not None and spi < 90) else NAVY_DARK
    st.markdown(f"""<div class="kpi-card" style="background:{BG_CARD};border:1px solid {BLUE_MID};">
        <div class="kpi-label" style="color:{MUTED};">SPI (ÍNDICE DE PRAZO)</div>
        <div class="kpi-value" style="color:{spi_color};">{spi_txt}</div>
        <div style="font-size:0.72rem;color:{MUTED};">real ÷ previsto pela data</div>
        </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="kpi-card" style="background:{BLUE_LIGHT};border:1px solid {BLUE};">
        <div class="kpi-label" style="color:{NAVY};">MARCO CRÍTICO</div>
        <div class="kpi-value" style="font-size:1.15rem;color:{NAVY_DARK};">{marco_txt}</div>
        </div>""", unsafe_allow_html=True)

st.write("")

# ---------------------------------------------------------------------------
# STATUS POR ETAPA
# ---------------------------------------------------------------------------
st.subheader("Status por etapa")

for _, row in etapa_df.iterrows():
    n = int(row["etapa_num"])
    label = f"**{n}. {row['etapa_title']}**"
    cols = st.columns([3, 6, 1.3])
    with cols[0]:
        crit = " 🔴" if "CRITICO" in row["etapa_title"].upper() or "CRÍTICO" in row["etapa_title"].upper() else ""
        st.markdown(label + crit)
    with cols[1]:
        if n == 0:
            st.progress(0, text="referência (cronograma legal)")
        elif not row["tem_cronograma"]:
            st.markdown(f'<span class="pending-tag">aguardando cronograma</span>', unsafe_allow_html=True)
        else:
            st.progress(min(int(row["real_pct"]), 100))
    with cols[2]:
        if n != 0 and row["tem_cronograma"]:
            st.markdown(f"**{row['real_pct']:.0f}%**")
        else:
            st.markdown("—")

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
    show = leaves[["start", "etapa_num", "name", "resource", "pct"]].copy()
    show.columns = ["Data", "Etapa", "Atividade", "Responsável", "% concluído"]
    show["Data"] = show["Data"].apply(lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "a definir")
    show["Atividade"] = show["Atividade"].str.strip()
    st.dataframe(show, hide_index=True, use_container_width=True)
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

    changed = 0
    for idx, (uid, new_pct) in enumerate(zip(editable["unique_id"], edited["% concluído"])):
        old_pct = editable["pct"].iloc[idx]
        if new_pct != old_pct:
            task = project.getTaskByUniqueID(jpype.java.lang.Integer(int(uid)))
            if task is not None:
                task.setPercentageComplete(jpype.java.lang.Double(float(new_pct)))
                changed += 1

    out_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xml").name
    writer = UniversalProjectWriter(FileFormat.MSPDI)
    writer.write(project, out_path)

    with open(out_path, "rb") as f:
        xml_bytes = f.read()

    st.success(f"{changed} tarefa(s) atualizada(s). Baixe o arquivo abaixo e abra direto no MS Project (Arquivo → Abrir).")
    st.download_button(
        "Baixar .xml atualizado",
        data=xml_bytes,
        file_name=f"Reforma_Tributaria_atualizado_{dt.date.today().isoformat()}.xml",
        mime="application/xml",
    )
