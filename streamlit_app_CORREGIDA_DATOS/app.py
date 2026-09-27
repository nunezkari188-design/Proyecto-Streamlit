import streamlit as st
import pandas as pd
import os
from analisis import *

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE LA PÁGINA
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="La Ciencia Sí es Cosa de Chicas",
    layout="wide",
    page_icon="🧬"
)

# ─────────────────────────────────────────────
# ESTILOS CSS
# ─────────────────────────────────────────────

st.markdown("""
<style>

/* Fondo general */
.stApp {
    background-color: #020617;
    color: white;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #020617, #071126);
    border-right: 1px solid #1e293b;
}

/* Textos sidebar */
section[data-testid="stSidebar"] * {
    color: white;
}

/* Tarjetas */
.metric-card {
    background: linear-gradient(135deg, #0f172a, #111827);
    border: 1px solid #1e3a8a;
    border-radius: 18px;
    padding: 25px;
    text-align: center;
    box-shadow: 0px 0px 20px rgba(30,58,138,0.15);
}

/* Títulos */
.main-title {
    text-align:center;
    color:white;
    font-size:64px;
    font-weight:800;
    margin-bottom:10px;
}

.subtitle {
    text-align:center;
    color:#94a3b8;
    font-size:20px;
}

/* Contenedor principal */
.hero {
    background: linear-gradient(135deg, #0f172a, #111827);
    padding: 40px;
    border-radius: 25px;
    border: 1px solid #1e3a8a;
    margin-bottom: 30px;
}

/* Secciones */
.section-title {
    color:#e2e8f0;
    font-size:34px;
    font-weight:700;
    margin-top:40px;
    margin-bottom:20px;
}

.section-sub {
    color:#94a3b8;
    font-size:15px;
    margin-top:-14px;
    margin-bottom:22px;
    line-height:1.6;
}

/* Tarjetas estadísticas */
.info-box {
    background-color:#111827;
    padding:20px;
    border-radius:15px;
    border:1px solid #1e293b;
    margin-bottom:20px;
}

/* Tarjetas cualitativas */
.qual-card{
    background: linear-gradient(135deg,#0f172a,#111827);
    border:1px solid #1e3a8a;
    border-radius:20px;
    padding:25px;
    margin-bottom:25px;
    box-shadow:0 0 20px rgba(30,58,138,.15);
}

.qual-title{
    color:#60a5fa;
    font-size:28px;
    font-weight:700;
    margin-bottom:15px;
}

.qual-text{
    color:#e2e8f0;
    font-size:18px;
    line-height:1.9;
}

/* ───────────────────────────────────────────
   TABLAS BONITAS (reemplazo de st.dataframe)
   ─────────────────────────────────────────── */

.tabla-wrap {
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid #1e3a5f;
    margin-bottom: 22px;
    box-shadow: 0 0 22px rgba(30,58,138,0.12);
}

.tabla-bonita {
    width: 100%;
    border-collapse: collapse;
    font-size: 14.5px;
    background: #0c1424;
}

.tabla-bonita thead tr {
    background: linear-gradient(135deg, #1e3a8a, #312e81);
}

.tabla-bonita thead th {
    color: #f1f5f9 !important;
    font-weight: 700;
    text-align: center;
    padding: 13px 14px;
    border-bottom: 2px solid #1e3a8a;
    white-space: nowrap;
}

.tabla-bonita thead th.izq {
    text-align: left;
}

.tabla-bonita tbody td {
    padding: 11px 14px;
    text-align: center;
    color: #e2e8f0;
    border-bottom: 1px solid #1e293b;
}

.tabla-bonita tbody td.izq {
    text-align: left;
    font-weight: 600;
    color: #f1f5f9;
}

.tabla-bonita tbody tr:nth-child(even) {
    background: #0f1b30;
}

.tabla-bonita tbody tr:hover {
    background: #16243f;
    transition: background 0.15s ease;
}

.tabla-bonita tbody tr:last-child td {
    border-bottom: none;
}

/* Badges para significancia / magnitud dentro de tablas */
.badge {
    display: inline-block;
    padding: 3px 11px;
    border-radius: 999px;
    font-size: 12.5px;
    font-weight: 700;
}

.badge-sig    { background: rgba(52,211,153,0.18); color: #34d399; }
.badge-nosig  { background: rgba(148,163,184,0.15); color: #94a3b8; }
.badge-up     { background: rgba(52,211,153,0.18);  color: #34d399; }
.badge-down   { background: rgba(248,113,113,0.18); color: #f87171; }
.badge-flat   { background: rgba(251,191,36,0.18);  color: #fbbf24; }

/* Estado de datos en sidebar */
.status-card {
    border-radius: 14px;
    padding: 14px 16px;
    margin-bottom: 6px;
    font-size: 14px;
    font-weight: 600;
    border: 1px solid;
}

.status-ok {
    background: rgba(52,211,153,0.12);
    border-color: rgba(52,211,153,0.4);
    color: #34d399;
}

.status-warn {
    background: rgba(251,191,36,0.12);
    border-color: rgba(251,191,36,0.4);
    color: #fbbf24;
}

.status-bad {
    background: rgba(248,113,113,0.12);
    border-color: rgba(248,113,113,0.4);
    color: #f87171;
}

/* Caja de ANOVA */
.anova-card {
    background: linear-gradient(135deg, #111827, #0c1424);
    border: 1px solid #312e81;
    border-radius: 18px;
    padding: 22px 26px;
    margin-bottom: 18px;
}

.anova-f {
    font-size: 38px;
    font-weight: 800;
    color: #c084fc;
}

.anova-label {
    color: #94a3b8;
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

/* Scroll bonito */
::-webkit-scrollbar {
    width: 10px;
}

::-webkit-scrollbar-track {
    background: #020617;
}

::-webkit-scrollbar-thumb {
    background: #1e3a8a;
    border-radius: 10px;
}

</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPERS DE RENDERIZADO
# ─────────────────────────────────────────────

def badge_sig(es_significativo, txt_si="Sig.", txt_no="No sig."):
    if es_significativo:
        return f'<span class="badge badge-sig">✅ {txt_si}</span>'
    return f'<span class="badge badge-nosig">{txt_no}</span>'


def badge_delta(valor, umbral=0.05):
    """Badge de color según signo/magnitud de un delta."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    try:
        v = float(valor)
    except (TypeError, ValueError):
        return str(valor)

    flecha = "▲" if v > 0 else ("▼" if v < 0 else "■")
    clase = "badge-up" if v > umbral else ("badge-down" if v < -umbral else "badge-flat")
    return f'<span class="badge {clase}">{flecha} {v:+.3f}</span>'


def fmt_celda(valor):
    """Formatea un valor para mostrarlo dentro de una celda HTML."""
    if valor is None:
        return "—"
    if isinstance(valor, float):
        if pd.isna(valor):
            return "—"
        return f"{valor:.3f}"
    return str(valor)


def render_tabla_html(df, columnas_izq=None, columnas_badge_sig=None,
                       columnas_badge_delta=None, decimales=3):
    """
    Renderiza un DataFrame como tabla HTML con estilo oscuro elegante.

    columnas_izq: columnas que deben alinearse a la izquierda (texto, p.ej. nombre de ítem)
    columnas_badge_sig: columnas tipo "Sig.*" que deben mostrarse como badge ✅/—
    columnas_badge_delta: columnas tipo "Δ" que deben mostrarse como badge ▲/▼ coloreado
    """
    columnas_izq = columnas_izq or []
    columnas_badge_sig = columnas_badge_sig or []
    columnas_badge_delta = columnas_badge_delta or []

    html = ['<div class="tabla-wrap"><table class="tabla-bonita"><thead><tr>']

    for col in df.columns:
        clase = "izq" if col in columnas_izq else ""
        html.append(f'<th class="{clase}">{col}</th>')
    html.append("</tr></thead><tbody>")

    for _, fila in df.iterrows():
        html.append("<tr>")
        for col in df.columns:
            valor = fila[col]
            clase = "izq" if col in columnas_izq else ""

            if col in columnas_badge_delta:
                contenido = badge_delta(valor)
            elif col in columnas_badge_sig:
                es_sig = (valor == "✅") or (valor is True)
                contenido = badge_sig(es_sig)
            elif isinstance(valor, float):
                contenido = fmt_celda(round(valor, decimales))
            else:
                contenido = valor if (valor is not None and str(valor) != "nan") else "—"

            html.append(f'<td class="{clase}">{contenido}</td>')
        html.append("</tr>")

    html.append("</tbody></table></div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def detectar_estado_datos(ruta_archivo):
    """
    Inspecciona las hojas del Excel sin cargarlo completo, para saber
    si tiene Pre, Post, ambos o ninguno. No lanza excepciones: si el
    archivo no existe o está corrupto, devuelve 'falta'.
    """
    if not os.path.exists(ruta_archivo):
        return "falta"
    try:
        hojas = [h.lower().strip() for h in pd.ExcelFile(ruta_archivo).sheet_names]
        tiene_pre  = any("pre" in h for h in hojas)
        tiene_post = any("post" in h for h in hojas)
        if tiene_pre and tiene_post:
            return "completo"
        elif tiene_pre:
            return "solo_pre"
        else:
            return "falta"
    except Exception:
        return "falta"


# ─────────────────────────────────────────────
# ARCHIVOS
# ─────────────────────────────────────────────

COLEGIOS = {
    "Colegio Eucarístico": os.path.join(DATA_DIR, "colegio_eucaristico.xlsx"),
    "Colegio Alemán": os.path.join(DATA_DIR, "colegio_aleman.xlsx"),
    "Colegio Sonia Ahumada": os.path.join(DATA_DIR, "colegio_sonia_ahumada.xlsx"),
    "Colegio Enseñanza": os.path.join(DATA_DIR, "colegio_ensenanza.xlsx")
}

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

with st.sidebar:

    st.markdown("""
    <h1 style='font-size:40px;'>🧬 Panel de Control</h1>
    """, unsafe_allow_html=True)

    st.markdown("### Selecciona una vista")

    OPCION_COMPARACION = "📊 Comparación entre colegios (ANOVA)"

    colegio = st.selectbox(
        "",
        list(COLEGIOS.keys()) + [OPCION_COMPARACION]
    )

    st.markdown("---")

    st.markdown("## Estado de datos")

    # Estado dinámico: se revisa cada colegio individualmente,
    # en vez de asumir que solo un colegio tiene datos completos.
    for nombre_col, ruta_col in COLEGIOS.items():
        estado = detectar_estado_datos(ruta_col)
        es_actual = "👉 " if nombre_col == colegio else ""

        if estado == "completo":
            st.markdown(
                f'<div class="status-card status-ok">{es_actual}{nombre_col}<br>'
                f'✅ Pre-test y Post-test completos</div>',
                unsafe_allow_html=True
            )
        elif estado == "solo_pre":
            st.markdown(
                f'<div class="status-card status-warn">{es_actual}{nombre_col}<br>'
                f'⚠️ Solo hay Pre-test disponible</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                f'<div class="status-card status-bad">{es_actual}{nombre_col}<br>'
                f'❌ Archivo no encontrado</div>',
                unsafe_allow_html=True
            )

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────

st.markdown("""
<div class="hero">

<h1 class="main-title">
🧬 La Ciencia Sí es Cosa de Chicas
</h1>

<p class="subtitle">
Dashboard interactivo de análisis STEM · Universidad del Norte
</p>

</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# VISTA: COMPARACIÓN ENTRE COLEGIOS (ANOVA one-way)
# ─────────────────────────────────────────────
# Esta vista es independiente de cualquier colegio puntual, por eso vive
# en su propia opción del sidebar en vez de repetirse dentro de cada colegio.

if colegio == OPCION_COMPARACION:

    st.markdown("""
    <div class="hero">

    <h1 class="main-title" style="font-size:44px;">
    🏫 Comparación entre colegios
    </h1>

    <p class="subtitle">
    Análisis ANOVA one-way · Pre-test · Las cuatro instituciones
    </p>

    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <p class="section-sub">
    El ANOVA de un factor responde: ¿existen diferencias estadísticamente
    significativas entre los cuatro colegios en el pre-test? Cuando el ANOVA
    es significativo, se complementa con comparaciones por pares (post-hoc,
    t-test de Welch) para identificar específicamente entre qué colegios
    está la diferencia.
    </p>
    """, unsafe_allow_html=True)

    with st.spinner("Cargando datos de todos los colegios para la comparación..."):
        datos_pre_colegios, errores_carga = cargar_pre_todos_colegios(COLEGIOS)

    if errores_carga:
        nombres_con_error = ", ".join(errores_carga.keys())
        st.info(f"No se pudieron cargar para esta comparación: {nombres_con_error}")

    if len(datos_pre_colegios) >= 2:

        for nombre_escala, config in ESCALAS.items():

            cols = config["columnas"]
            resultado_anova = anova_entre_colegios(datos_pre_colegios, cols)

            if "error" in resultado_anova:
                continue

            st.markdown(f"#### 📐 {nombre_escala}")

            c1, c2 = st.columns([1, 2])

            with c1:
                st.markdown(f"""
                <div class="anova-card">
                    <div class="anova-label">Estadístico F</div>
                    <div class="anova-f">{resultado_anova['F']}</div>
                    <div style="margin-top:10px;">
                        <span class="anova-label">p-valor</span>
                        <div style="font-size:20px; font-weight:700; color:#e2e8f0;">
                            {resultado_anova['p_valor']}
                        </div>
                    </div>
                    <div style="margin-top:10px;">
                        {badge_sig(resultado_anova['significativo'],
                                   "Diferencias significativas",
                                   "Sin diferencias significativas")}
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with c2:
                df_medias = pd.DataFrame({
                    "Colegio": list(resultado_anova["medias"].keys()),
                    "Media": list(resultado_anova["medias"].values()),
                    "n": [resultado_anova["ns"][c] for c in resultado_anova["medias"].keys()],
                })
                render_tabla_html(df_medias, columnas_izq=["Colegio"])

            # Post-hoc solo si el ANOVA fue significativo
            if resultado_anova["significativo"]:
                df_posthoc = posthoc_pares_colegios(resultado_anova["grupos"])
                df_sig = df_posthoc[df_posthoc["Sig.*"] == "✅"] if len(df_posthoc) > 0 else df_posthoc

                if len(df_sig) > 0:
                    st.markdown("##### 🔍 Pares de colegios con diferencia significativa (post-hoc)")
                    render_tabla_html(
                        df_sig,
                        columnas_izq=["Colegio A", "Colegio B"],
                        columnas_badge_sig=["Sig.*"],
                        columnas_badge_delta=["Δ (B-A)"],
                    )
                else:
                    st.caption(
                        "El ANOVA general fue significativo, pero ninguna "
                        "comparación par a par alcanzó significancia individual "
                        "(esto puede ocurrir cuando el efecto se distribuye entre "
                        "varios pares en lugar de concentrarse en uno solo)."
                    )

            st.markdown("<br>", unsafe_allow_html=True)

    else:
        st.warning(
            "Se necesitan al menos 2 colegios con datos de Pre-test cargados "
            "para realizar la comparación ANOVA."
        )

    st.markdown("""
    <br><br>
    <hr style="border:1px solid #1e293b;">
    <p style="text-align:center; color:#64748b; font-size:14px;">
    Dashboard desarrollado para investigación STEM · Universidad del Norte
    </p>
    """, unsafe_allow_html=True)

    st.stop()

# ─────────────────────────────────────────────
# CARGAR DATOS
# ─────────────────────────────────────────────

archivo = COLEGIOS[colegio]

if not os.path.exists(archivo):
    st.error(f"No se encontró el archivo: {archivo}")
    st.stop()

try:
    pre, post = cargar_excel(archivo)

except Exception as e:

    st.error("⚠️ Este colegio todavía no tiene datos completos.")
    st.info(str(e))

    st.stop()

# ─────────────────────────────────────────────
# MÉTRICAS PRINCIPALES
# ─────────────────────────────────────────────

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <h1 style="color:#60a5fa;">{max(len(pre), len(post))}</h1>
        <p>Participantes</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="metric-card">
        <h1 style="color:#c084fc;">4</h1>
        <p>Escalas STEM</p>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="metric-card">
        <h1 style="color:#34d399;">Pre/Post</h1>
        <p>Comparación</p>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="metric-card">
        <h1 style="color:#f97316;">Dashboard</h1>
        <p>Interactivo</p>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# PERFIL SOCIODEMOGRÁFICO
# ─────────────────────────────────────────────

st.markdown("""
<div class="section-title">
📊 Perfil Sociodemográfico
</div>
""", unsafe_allow_html=True)

sd = estadisticos_sociodem(pre, post)

edad = sd["edad"]

st.markdown(f"""
<div class="info-box">

<h3 style="color:#60a5fa;">
👩 Edad promedio
</h3>

<p style="font-size:30px; color:white;">
{edad['media']} ± {edad['de']} años
</p>

<p style="color:#94a3b8;">
Rango {edad['min']}–{edad['max']}
</p>

</div>
""", unsafe_allow_html=True)

if "tipo_colegio" in sd:
    st.markdown("#### 🏫 Tipo de colegio")
    df_tc = sd["tipo_colegio"].reset_index().rename(columns={"index": "Tipo_colegio"})
    render_tabla_html(df_tc, columnas_izq=["Tipo_colegio"])

if "padres_STEM" in sd:
    st.markdown("#### 👨‍🔬 Padres trabajan en STEM")
    df_ps = sd["padres_STEM"].reset_index().rename(columns={"index": "Padres_trabajan_STEM"})
    render_tabla_html(df_ps, columnas_izq=["Padres_trabajan_STEM"])

if "conoce_mujer_STEM" in sd:
    st.markdown("#### 👩‍🔬 Conoce mujer STEM")
    df_cm = sd["conoce_mujer_STEM"].reset_index().rename(columns={"index": "Conoce_mujer_STEM"})
    render_tabla_html(df_cm, columnas_izq=["Conoce_mujer_STEM"])

# ─────────────────────────────────────────────
# RESULTADOS DE ESCALAS
# ─────────────────────────────────────────────

st.markdown("""
<div class="section-title">
🧪 Resultados de Escalas STEM
</div>
""", unsafe_allow_html=True)

st.markdown("""
<p class="section-sub">
Comparación Pre-test vs. Post-test mediante t-test de Welch para muestras
independientes (los instrumentos no cuentan con identificador individual,
por lo que no se asume emparejamiento participante a participante).
</p>
""", unsafe_allow_html=True)

resultados_escalas = {}

for nombre_escala, config in ESCALAS.items():

    cols = config["columnas"]

    res = test_t_pareado(pre, post, cols)

    if "error" in res:
        continue

    resultados_escalas[nombre_escala] = res

    st.markdown(f"""
    <div class="info-box">
    <h2 style="color:#c084fc;">
    📐 {nombre_escala}
    </h2>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric(
            "PRE",
            f"{res['media_pre']:.3f}"
        )

    with c2:
        st.metric(
            "POST",
            f"{res['media_post']:.3f}",
            delta=f"{res['delta']:.3f}"
        )

    with c3:
        st.metric(
            "p-valor",
            f"{res['p_valor']}"
        )

    with c4:
        st.metric(
            "Cohen's d",
            f"{res['d_cohen']}"
        )

    st.markdown(
        f'<p style="font-size:16px; color:#e2e8f0; margin-bottom:18px;">'
        f'Magnitud del efecto: <strong>{res["magnitud"]}</strong> &nbsp; '
        f'{badge_sig(res["significativo"])}</p>',
        unsafe_allow_html=True
    )

    # Tabla ítem a ítem, ahora con tabla bonita en vez de solo gráfica
    df_items = test_item_a_item(pre, post, cols, config["etiquetas"])
    if len(df_items) > 0:
        st.markdown("##### Detalle por ítem")
        render_tabla_html(
            df_items,
            columnas_izq=["Ítem"],
            columnas_badge_sig=["Sig.*"],
            columnas_badge_delta=["Δ"],
        )

    fig1 = grafica_items(
        df_items,
        nombre_escala,
        colegio,
        config["color"]
    )

    st.pyplot(fig1)

# ─────────────────────────────────────────────
# APOYO PERCIBIDO (variables binarias)
# ─────────────────────────────────────────────

df_apoyo = analizar_binarias(pre, post, APOYO_COLS)

if len(df_apoyo) > 0:
    st.markdown("""
    <div class="section-title">
    🤝 Apoyo percibido
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <p class="section-sub">
    Comparación de proporciones (test z para dos muestras independientes).
    </p>
    """, unsafe_allow_html=True)

    render_tabla_html(
        df_apoyo,
        columnas_izq=["Variable"],
        columnas_badge_sig=["Sig.*"],
        columnas_badge_delta=["Δ pp"],
    )

# ─────────────────────────────────────────────
# RESUMEN GENERAL
# ─────────────────────────────────────────────

st.markdown("""
<div class="section-title">
📋 Tabla Resumen General
</div>
""", unsafe_allow_html=True)

filas = []

for nombre_escala, res in resultados_escalas.items():

    filas.append({
        "Escala": nombre_escala,
        "PRE media": res["media_pre"],
        "POST media": res["media_post"],
        "Δ": res["delta"],
        "p-valor": res["p_valor"],
        "Cohen's d": res["d_cohen"],
        "Magnitud": res["magnitud"],
        "Sig.*": "✅" if res["significativo"] else ""
    })

df_resumen = pd.DataFrame(filas)

render_tabla_html(
    df_resumen,
    columnas_izq=["Escala"],
    columnas_badge_sig=["Sig.*"],
    columnas_badge_delta=["Δ"],
)

# ─────────────────────────────────────────────
# GRÁFICA RADAR
# ─────────────────────────────────────────────

st.markdown("""
<div class="section-title">
🕸️ Comparación General PRE vs POST
</div>
""", unsafe_allow_html=True)

fig_radar = grafica_radar(
    resultados_escalas,
    colegio
)

st.pyplot(fig_radar)

# ─────────────────────────────────────────────
# HALLAZGOS CUALITATIVOS
# ─────────────────────────────────────────────

st.markdown("""
<div class="section-title">
📝 Hallazgos Cualitativos
</div>
""", unsafe_allow_html=True)

hallazgos = interpretacion_cualitativa(colegio)

for titulo, texto in hallazgos.items():

    with st.container(border=True):

        st.markdown(
            f"### 🔬 {titulo}"
        )

        st.write(texto)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────

st.markdown("""
<br><br>

<hr style="border:1px solid #1e293b;">

<p style="
text-align:center;
color:#64748b;
font-size:14px;
">
Dashboard desarrollado para investigación STEM · Universidad del Norte
</p>
""", unsafe_allow_html=True)
