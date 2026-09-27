# =============================================================================
#  PROGRAMA DE ANÁLISIS: "La Ciencia sí es Cosa de Chicas"
#  Universidad del Norte · Barranquilla, Colombia
# =============================================================================
#  Este programa carga los Excel de cada colegio, limpia los datos,
#  y genera un análisis completo pre-test vs post-test.
#
#  CÓMO USARLO
#  -----------
#  1. Pon todos tus Excel en la misma carpeta que este archivo.
#  2. Ajusta el diccionario COLEGIOS (línea ~80) con los nombres exactos
#     de tus archivos.
#  3. Ejecuta:  python analisis_ciencia_chicas.py
#  4. Verás los resultados en consola y se guardarán en una carpeta
#     "resultados/" con tablas CSV y gráficas PNG por cada colegio.
# =============================================================================
import streamlit as st
import os
import warnings
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from scipy import stats
from wordcloud import WordCloud
from collections import Counter
import re
import random

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
#  1. CONFIGURACIÓN: COLEGIOS Y ARCHIVOS EXCEL
#     Agrega aquí tus 4 colegios. La clave es el nombre que quieres mostrar
#     y el valor es el nombre exacto del archivo Excel.
# ─────────────────────────────────────────────────────────────────────────────

COLEGIOS = {
    "Colegio Eucarístico": os.path.join(DATA_DIR, "colegio_eucaristico.xlsx"),
    # "Colegio 2": "Datos_Colegio_2.xlsx",
    # "Colegio 3": "Datos_Colegio_3.xlsx",
    # "Colegio 4": "Datos_Colegio_4.xlsx",
}

# ─────────────────────────────────────────────────────────────────────────────
#  2. DEFINICIÓN DE ESCALAS E ÍTEMS
#     Cada escala tiene: nombre de columnas en Excel, etiqueta legible y
#     el enunciado exacto que aparecía en el instrumento (según hoja Codigos).
# ─────────────────────────────────────────────────────────────────────────────

ESCALAS = {
    "Habilidades STEM": {
        "columnas": ["Hab_resolver_mat", "Hab_usar_lab", "Hab_entender",
                     "Hab_usar_tec", "Hab_explicar"],
        "etiquetas": {
            "Hab_resolver_mat": "Resolver problemas matemáticos",
            "Hab_usar_lab":     "Usar equipos de laboratorio",
            "Hab_entender":     "Entender conceptos científicos",
            "Hab_usar_tec":     "Usar tecnología para aprender",
            "Hab_explicar":     "Explicar fenómenos científicos",
        },
        "color": "#c084fc",
        "tipo":  "likert_1_5",
    },
    "Interés en materias STEM": {
        "columnas": ["Bio", "Qui", "Fis", "Mat", "Tec", "Ing"],
        "etiquetas": {
            "Bio": "Biología", "Qui": "Química",  "Fis": "Física",
            "Mat": "Matemáticas", "Tec": "Tecnología", "Ing": "Ingeniería",
        },
        "color": "#f97316",
        "tipo":  "likert_1_5",
    },
    "Normas de género": {
        "columnas": ["Norma1", "Norma2", "Norma3", "Norma4", "Norma5"],
        "etiquetas": {
            "Norma1": "Niñas igual de buenas en matemáticas",
            "Norma2": "Ciencia es para hombres y mujeres",
            "Norma3": "Cómoda siendo única niña en ciencia",
            "Norma4": "Importante más mujeres científicas",
            "Norma5": "Hombres NO entienden mejor la tecnología",
        },
        "color": "#34d399",
        "tipo":  "likert_1_5",
    },
    "Expectativas futuras STEM": {
        "columnas": ["Futuro1", "Futuro2", "Futuro3", "Futuro4", "Futuro5"],
        "etiquetas": {
            "Futuro1": "Estudiar carrera STEM",
            "Futuro2": "Ser parte de investigación científica",
            "Futuro3": "Trabajar en lab o empresa de tecnología",
            "Futuro4": "Enseñar ciencia a otras personas",
            "Futuro5": "Crear solución a un problema con ciencia",
        },
        "color": "#60a5fa",
        "tipo":  "likert_1_5",
    },
}

# Columnas binarias (0/1): apoyo percibido y experiencias previas STEM
APOYO_COLS = {
    "Apoyo1": "Familia apoya interés STEM",
    "Apoyo2": "Profesores animan a seguir STEM",
    "Apoyo3": "Amigas comparten interés STEM",
    "Apoyo4": "Tiene referente mujer científica cercana",
    "Apoyo5": "Se siente capaz de trabajar en ciencia",
}

EXP_COLS = {
    "Exp1": "Ha participado en actividad científica",
    "Exp2": "Ha visitado un laboratorio",
    "Exp3": "Ha leído sobre ciencia por gusto",
    "Exp4": "Ha hablado con una científica",
    "Exp5": "Ha hecho proyecto de ciencia",
}

# ─────────────────────────────────────────────────────────────────────────────
#  3. FUNCIONES DE CARGA Y LIMPIEZA
# ─────────────────────────────────────────────────────────────────────────────

def cargar_excel(ruta_archivo):

    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")

    # Leer nombres de hojas disponibles
    excel_file = pd.ExcelFile(ruta_archivo)
    hojas = excel_file.sheet_names

    pre = None
    post = None

    # Buscar hoja PRE
    for hoja in hojas:
        nombre = hoja.lower().strip()

        if "pre" in nombre:
            pre = pd.read_excel(ruta_archivo, sheet_name=hoja, header=1)
            break

    # Buscar hoja POST
    for hoja in hojas:
        nombre = hoja.lower().strip()

        if "post" in nombre:
            post = pd.read_excel(ruta_archivo, sheet_name=hoja, header=1)
            break

    # Si no existe PRE
    if pre is None:
        raise ValueError(
            f"El archivo '{os.path.basename(ruta_archivo)}' no tiene hoja PRE."
        )

    # Si no existe POST
    if post is None:

        raise ValueError(
        "Este colegio solo tiene Pre-test cargado."
        )

    # Limpiar filas vacías
    if "ID" in pre.columns:
        pre = pre.dropna(subset=["ID"]).reset_index(drop=True)

    if "ID" in post.columns:
        post = post.dropna(subset=["ID"]).reset_index(drop=True)

    # Convertir columnas numéricas
    todas_num = (
        list(col for esc in ESCALAS.values() for col in esc["columnas"]) +
        list(APOYO_COLS.keys()) +
        list(EXP_COLS.keys()) +
        [
            "Edad",
            "Grado",
            "Tipo_colegio",
            "Padres_trabajan_STEM",
            "Internet_casa",
            "Dispositivo_estudio",
            "Conoce_mujer_STEM",
            "Participacion_programa_STEM"
        ]
    )

    for df in [pre, post]:

        for col in todas_num:

            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

    return pre, post

def cargar_asistencia(ruta_archivo):
    """
    Lee la hoja Asistencia del Excel.
    Los encabezados están en la fila 4 (header=3 en pandas).
    """
    asistencia = pd.read_excel(ruta_archivo, sheet_name="Asistencia", header=3)
    return asistencia

# ─────────────────────────────────────────────────────────────────────────────
#  4. FUNCIONES DE ANÁLISIS ESTADÍSTICO
# ─────────────────────────────────────────────────────────────────────────────

def promedio_escala(df, columnas):
    """
    Calcula el promedio de cada participante en una escala Likert.
    Ignora NaN de forma individual (una sola pregunta sin respuesta
    no invalida a toda la participante).
    """
    return df[columnas].mean(axis=1)

def test_t_independiente(pre, post, columnas):
    """
    T-test de Welch para muestras independientes.
    Corresponde al caso en que pre y post no tienen identificadores
    individuales y no se pueden emparejar participante a participante.
    Cohen's d se calcula con la desviación estándar combinada (pooled).
    """
    v_pre  = promedio_escala(pre,  columnas).dropna()
    v_post = promedio_escala(post, columnas).dropna()

    n_pre  = len(v_pre)
    n_post = len(v_post)

    if n_pre < 3 or n_post < 3:
        return {"error": "Muy pocos datos"}

    t_stat, p_valor = stats.ttest_ind(v_pre, v_post, equal_var=False,
                                      nan_policy="omit")

    media_pre  = v_pre.mean()
    media_post = v_post.mean()
    de_pre     = v_pre.std(ddof=1)
    de_post    = v_post.std(ddof=1)

    media_diff = media_post - media_pre

    # Cohen's d con SD pooled
    sd_pool = np.sqrt(((n_pre - 1) * de_pre**2 + (n_post - 1) * de_post**2)
                      / (n_pre + n_post - 2))
    d_cohen = media_diff / sd_pool if sd_pool > 0 else 0

    abs_d = abs(d_cohen)
    if abs_d < 0.2:
        magnitud = "sin efecto práctico"
    elif abs_d < 0.5:
        magnitud = "efecto pequeño"
    elif abs_d < 0.8:
        magnitud = "efecto medio"
    else:
        magnitud = "efecto grande"

    # IC 95% de la diferencia de medias (Welch)
    se_diff = np.sqrt(de_pre**2 / n_pre + de_post**2 / n_post)
    # Grados de libertad Welch-Satterthwaite
    df_welch = (de_pre**2/n_pre + de_post**2/n_post)**2 / (
        (de_pre**2/n_pre)**2 / (n_pre - 1) +
        (de_post**2/n_post)**2 / (n_post - 1)
    )
    t_crit = stats.t.ppf(0.975, df=df_welch)
    ic_inf = media_diff - t_crit * se_diff
    ic_sup = media_diff + t_crit * se_diff

    return {
        "n_pre":  n_pre,
        "n_post": n_post,
        "n":      min(n_pre, n_post),   # compatibilidad con imprimir_resultado_escala

        "media_pre":  round(media_pre,  3),
        "de_pre":     round(de_pre,     3),
        "media_post": round(media_post, 3),
        "de_post":    round(de_post,    3),

        "delta":      round(media_diff, 3),

        "t_stat":     round(t_stat,  3),
        "p_valor":    round(p_valor, 4),

        "significativo": p_valor < 0.05,

        "d_cohen":  round(d_cohen, 3),
        "magnitud": magnitud,

        "ic_95": (round(ic_inf, 3), round(ic_sup, 3)),
    }

# Alias para mantener compatibilidad con el resto del código
def test_t_pareado(pre, post, columnas):
    return test_t_independiente(pre, post, columnas)

# ─────────────────────────────────────────────────────────────────────────────
#  ANOVA ONE-WAY ENTRE COLEGIOS + POST-HOC (Welch por pares)
# ─────────────────────────────────────────────────────────────────────────────

def anova_entre_colegios(datos_por_colegio, columnas):
    """
    Compara los promedios de una escala entre varios colegios usando
    ANOVA one-way (stats.f_oneway). Recibe un diccionario
    {nombre_colegio: DataFrame} (normalmente los pre-test de cada colegio)
    y la lista de columnas que conforman la escala.

    Devuelve un dict con F, p-valor, si es significativo, y la media de
    cada colegio para poder mostrarla en una tabla.
    """
    grupos = {}
    for nombre, df in datos_por_colegio.items():
        v = promedio_escala(df, columnas).dropna()
        if len(v) >= 3:
            grupos[nombre] = v

    if len(grupos) < 2:
        return {"error": "Se necesitan al menos 2 colegios con datos suficientes"}

    F, p_valor = stats.f_oneway(*grupos.values())

    medias = {nombre: round(v.mean(), 3) for nombre, v in grupos.items()}
    ns     = {nombre: len(v) for nombre, v in grupos.items()}

    return {
        "F": round(F, 3),
        "p_valor": round(p_valor, 4),
        "significativo": p_valor < 0.05,
        "medias": medias,
        "ns": ns,
        "grupos": grupos,   # se reutiliza en el post-hoc para no recalcular
    }


def posthoc_pares_colegios(grupos):
    """
    Toma el diccionario `grupos` devuelto por anova_entre_colegios
    (nombre_colegio -> Serie de promedios) y compara cada par de
    colegios con el t-test de Welch (muestras independientes).

    Devuelve un DataFrame con una fila por par de colegios, incluyendo
    solo los que tengan datos suficientes. Se recomienda mostrar primero
    los pares con Sig.* = ✅.
    """
    from itertools import combinations

    filas = []
    nombres = list(grupos.keys())

    for c1, c2 in combinations(nombres, 2):
        v1, v2 = grupos[c1], grupos[c2]
        n1, n2 = len(v1), len(v2)

        if n1 < 3 or n2 < 3:
            continue

        t_stat, p_valor = stats.ttest_ind(v1, v2, equal_var=False)

        m1, m2 = v1.mean(), v2.mean()
        s1, s2 = v1.std(ddof=1), v2.std(ddof=1)
        sd_pool = np.sqrt(((n1-1)*s1**2 + (n2-1)*s2**2) / (n1+n2-2))
        d_cohen = (m2 - m1) / sd_pool if sd_pool > 0 else 0

        filas.append({
            "Colegio A": c1,
            "M_A": round(m1, 3),
            "Colegio B": c2,
            "M_B": round(m2, 3),
            "Δ (B-A)": round(m2 - m1, 3),
            "t": round(t_stat, 3),
            "p-valor": round(p_valor, 4),
            "d Cohen": round(d_cohen, 3),
            "Sig.*": "✅" if p_valor < 0.05 else "",
        })

    df = pd.DataFrame(filas)
    if len(df) > 0:
        df = df.sort_values("p-valor").reset_index(drop=True)
    return df


def cargar_pre_todos_colegios(colegios_dict):
    """
    Carga la hoja Pre de cada colegio definido en `colegios_dict`
    ({nombre: ruta_archivo}) y devuelve {nombre: DataFrame}.
    Colegios cuyo archivo falle al cargar se omiten silenciosamente
    (se reportan aparte si se necesita).
    """
    resultado = {}
    errores = {}
    for nombre, ruta in colegios_dict.items():
        try:
            pre, _ = cargar_excel(ruta)
            resultado[nombre] = pre
        except Exception as e:
            errores[nombre] = str(e)
    return resultado, errores


def test_item_a_item(pre, post, columnas, etiquetas):

    resultados = []

    for col in columnas:

        if col not in pre.columns or col not in post.columns:
            continue

        v_pre  = pre[col].dropna()
        v_post = post[col].dropna()

        n_pre  = len(v_pre)
        n_post = len(v_post)
        n      = min(n_pre, n_post)

        if n_pre >= 3 and n_post >= 3:

            t_stat, p_valor = stats.ttest_ind(
                v_pre,
                v_post,
                equal_var=False
            )

        else:

            t_stat = None
            p_valor = None

        resultados.append({
            "Ítem": etiquetas.get(col, col),

            "Columna": col,

            "PRE media":
                round(v_pre.mean(), 3)
                if n_pre > 0 else np.nan,

            "POST media":
                round(v_post.mean(), 3)
                if n_post > 0 else np.nan,

            "Δ":
                round(
                    v_post.mean()
                    - v_pre.mean(),
                    3
                )
                if n_pre > 0 and n_post > 0 else np.nan,

            "t":
                round(t_stat, 3)
                if t_stat is not None else "—",

            "p-valor":
                round(p_valor, 4)
                if p_valor is not None else "—",

            "Sig.*":
                "✅"
                if (
                    p_valor is not None
                    and p_valor < 0.05
                )
                else "",

            "n_pre": n_pre,
            "n_post": n_post,
        })

    return pd.DataFrame(resultados)
def analizar_binarias(pre, post, columnas_dict):
    """
    Para columnas binarias (0/1), calcula el porcentaje que marcó 'sí'
    en pre y post, y aplica un test z de comparación de proporciones
    para dos muestras independientes (no se asume emparejamiento,
    ya que los instrumentos no incluyen identificador individual).
    """
    resultados = []
    for col, etiqueta in columnas_dict.items():
        if col not in pre.columns or col not in post.columns:
            continue

        vp = pre[col].dropna().astype(int).values
        vq = post[col].dropna().astype(int).values

        n_pre, n_post = len(vp), len(vq)

        if n_pre == 0 or n_post == 0:
            continue

        pct_pre  = round(vp.mean() * 100, 1)
        pct_post = round(vq.mean() * 100, 1)

        x_pre, x_post = vp.sum(), vq.sum()
        p_pool = (x_pre + x_post) / (n_pre + n_post)

        if 0 < p_pool < 1:
            se = np.sqrt(p_pool * (1 - p_pool) * (1/n_pre + 1/n_post))
            z = (vq.mean() - vp.mean()) / se if se > 0 else 0
            p_val = round(2 * (1 - stats.norm.cdf(abs(z))), 4)
            sig   = "✅" if p_val < 0.05 else ""
        else:
            z = None
            p_val = None
            sig   = ""

        resultados.append({
            "Variable":    etiqueta,
            "Columna":     col,
            "% PRE sí":    pct_pre,
            "% POST sí":   pct_post,
            "Δ pp":        round(pct_post - pct_pre, 1),
            "n PRE":       n_pre,
            "n POST":      n_post,
            "z":           round(z, 3) if z is not None else "—",
            "p-valor":     p_val if p_val is not None else "—",
            "Sig.*":       sig,
        })

    return pd.DataFrame(resultados)

def estadisticos_sociodem(pre, post):
    """
    Calcula estadísticos descriptivos del perfil sociodemográfico
    usando el post-test.
    """

    stats_dict = {}

    # Edad
    if "Edad" in post.columns:

        edad = post["Edad"].dropna()

        if len(edad) > 0:

            stats_dict["edad"] = {
                "media": round(edad.mean(), 2),
                "de": round(edad.std(ddof=1), 2),
                "min": int(edad.min()),
                "max": int(edad.max()),
            }

        else:

            stats_dict["edad"] = {
                "media": np.nan,
                "de": np.nan,
                "min": np.nan,
                "max": np.nan,
            }

    else:

        stats_dict["edad"] = {
            "media": np.nan,
            "de": np.nan,
            "min": np.nan,
            "max": np.nan,
        }

    # Distribuciones categóricas
    for col, nombre in [
        ("Tipo_colegio", "tipo_colegio"),
        ("Padres_trabajan_STEM", "padres_STEM"),
        ("Internet_casa", "internet"),
        ("Dispositivo_estudio", "dispositivo"),
        ("Participacion_programa_STEM", "prog_STEM_previo"),
        ("Conoce_mujer_STEM", "conoce_mujer_STEM"),
    ]:

        if col in post.columns:

            counts = (
                post[col]
                .dropna()
                .astype(int)
                .value_counts()
                .sort_index()
            )

            pcts = (
                counts / counts.sum() * 100
            ).round(1)

            stats_dict[nombre] = pd.DataFrame({
                "n": counts,
                "%": pcts
            })

    # Convivencia
    vive_cols = {
        "Vive_mama": "Mamá",
        "Vive_papa": "Papá",
        "Vive_abuelos": "Abuelos",
        "Vive_hermanos": "Hermanos",
        "Vive_otra": "Otro familiar",
    }

    vive_data = {}

    for col, etiq in vive_cols.items():

        if col in post.columns:

            n = post[col].fillna(0).astype(int).sum()

            vive_data[etiq] = {
                "n": n,
                "%": round(n / len(post) * 100, 1)
            }

    if len(vive_data) > 0:
        stats_dict["convivencia"] = pd.DataFrame(vive_data).T

    return stats_dict

# ─────────────────────────────────────────────────────────────────────────────
#  5. FUNCIONES DE VISUALIZACIÓN
# ─────────────────────────────────────────────────────────────────────────────

ESTILO = {
    "fondo":      "#0d1117",
    "fondo2":     "#111827",
    "borde":      "#1e3a5f",
    "texto":      "#e0e6f0",
    "texto2":     "#90a4ae",
    "gris_pre":   "#475569",
}

def aplicar_estilo(fig, ax_list):
    """
    Aplica el estilo oscuro a una figura y lista de ejes.
    """
    fig.patch.set_facecolor(ESTILO["fondo"])
    for ax in ax_list:
        ax.set_facecolor(ESTILO["fondo2"])
        for spine in ax.spines.values():
            spine.set_color(ESTILO["borde"])
        ax.tick_params(colors=ESTILO["texto2"])
        ax.xaxis.label.set_color(ESTILO["texto2"])
        ax.yaxis.label.set_color(ESTILO["texto2"])
        ax.title.set_color(ESTILO["texto"])

def grafica_resumen_escalas(resultados_escalas, nombre_colegio):
    """
    Gráfico de barras agrupadas: media PRE vs POST para cada escala.
    Marca con asterisco las diferencias significativas.
    """
    nombres = list(resultados_escalas.keys())
    medias_pre  = [r["media_pre"]  for r in resultados_escalas.values()]
    medias_post = [r["media_post"] for r in resultados_escalas.values()]
    sig_flags   = [r["significativo"] for r in resultados_escalas.values()]

    x = np.arange(len(nombres))
    ancho = 0.35

    fig, ax = plt.subplots(figsize=(10, 5), facecolor=ESTILO["fondo"])
    aplicar_estilo(fig, [ax])

    barras_pre  = ax.bar(x - ancho/2, medias_pre,  ancho, label="Pre-test",
                         color=ESTILO["gris_pre"], edgecolor="none")
    barras_post = ax.bar(x + ancho/2, medias_post, ancho, label="Post-test",
                         color="#c084fc", edgecolor="none")

    # Añadir valor encima de cada barra
    for bar, val, sig in zip(barras_post, medias_post, sig_flags):
        label = f"{val:.2f}" + (" *" if sig else "")
        ax.text(bar.get_x() + bar.get_width()/2, val + 0.05,
                label, ha="center", va="bottom",
                fontsize=8.5, color=ESTILO["texto"])

    for bar, val in zip(barras_pre, medias_pre):
        ax.text(bar.get_x() + bar.get_width()/2, val + 0.05,
                f"{val:.2f}", ha="center", va="bottom",
                fontsize=8.5, color=ESTILO["texto2"])

    ax.set_xticks(x)
    ax.set_xticklabels(nombres, rotation=15, ha="right", fontsize=9,
                       color=ESTILO["texto"])
    ax.set_ylim(0, 5.8)
    ax.set_ylabel("Promedio (escala 1–5)", fontsize=9)
    ax.set_title(f"Comparación Pre vs Post · {nombre_colegio}", fontsize=11,
                 color=ESTILO["texto"], pad=12)
    ax.yaxis.grid(True, color=ESTILO["borde"], linewidth=0.6, linestyle="--")
    ax.set_axisbelow(True)

    ax.legend(facecolor=ESTILO["fondo"], edgecolor=ESTILO["borde"],
              labelcolor=ESTILO["texto"], fontsize=9)

    ax.text(0.99, 0.98, "* p < 0.05", transform=ax.transAxes,
            ha="right", va="top", fontsize=8, color="#4ade80")

    fig.tight_layout()
    return fig

def grafica_items(df_items, nombre_escala, nombre_colegio, color):
    """
    Gráfico horizontal de barras para cada ítem de una escala.
    Muestra PRE y POST lado a lado.
    """
    etiquetas = df_items["Ítem"].tolist()
    vals_pre  = df_items["PRE media"].tolist()
    vals_post = df_items["POST media"].tolist()
    sig_flags = [s == "✅" for s in df_items["Sig.*"].tolist()]

    y = np.arange(len(etiquetas))
    alto = max(3.5, len(etiquetas) * 0.55)

    fig, ax = plt.subplots(figsize=(9, alto), facecolor=ESTILO["fondo"])
    aplicar_estilo(fig, [ax])

    ax.barh(y + 0.18, vals_pre,  0.34, label="Pre-test",
            color=ESTILO["gris_pre"], edgecolor="none")
    ax.barh(y - 0.18, vals_post, 0.34, label="Post-test",
            color=color, edgecolor="none")

    # Asterisco en ítems significativos
    for i, (val, sig) in enumerate(zip(vals_post, sig_flags)):
        if sig:
            ax.text(val + 0.06, y[i] - 0.18, "*", color="#4ade80",
                    fontsize=12, va="center")

    ax.set_yticks(y)
    ax.set_yticklabels(etiquetas, fontsize=9, color=ESTILO["texto"])
    ax.set_xlim(0, 5.8)
    ax.set_xlabel("Promedio (1–5)", fontsize=9)
    ax.set_title(f"{nombre_escala} · {nombre_colegio}", fontsize=10,
                 color=ESTILO["texto"], pad=10)
    ax.xaxis.grid(True, color=ESTILO["borde"], linewidth=0.5, linestyle="--")
    ax.set_axisbelow(True)
    ax.legend(facecolor=ESTILO["fondo"], edgecolor=ESTILO["borde"],
              labelcolor=ESTILO["texto"], fontsize=8)

    fig.tight_layout()
    return fig

def grafica_cambios_individuales(pre, post, columnas, nombre_escala,
                                  nombre_colegio, color):
    """
    Barras de cambio individual (post - pre) por participante.
    Verde = mejoró, Rojo = bajó.
    """
    v_pre  = promedio_escala(pre,  columnas)
    v_post = promedio_escala(post, columnas)
    diffs  = (v_post - v_pre).dropna().values

    colores = ["#4ade80" if d > 0 else "#f87171" if d < 0 else "#94a3b8"
               for d in diffs]

    fig, ax = plt.subplots(figsize=(9, 3.5), facecolor=ESTILO["fondo"])
    aplicar_estilo(fig, [ax])

    ax.bar(range(len(diffs)), diffs, color=colores, edgecolor="none")
    ax.axhline(0, color=ESTILO["texto2"], linewidth=1, linestyle="--")

    ax.set_xlabel("Participante (n° de registro)", fontsize=9)
    ax.set_ylabel("Δ (post − pre)", fontsize=9)
    ax.set_title(f"Cambio individual · {nombre_escala} · {nombre_colegio}",
                 fontsize=10, color=ESTILO["texto"], pad=10)
    ax.yaxis.grid(True, color=ESTILO["borde"], linewidth=0.5, linestyle="--")
    ax.set_axisbelow(True)
    ax.set_xticks(range(len(diffs)))
    ax.set_xticklabels([str(i+1) for i in range(len(diffs))],
                       fontsize=8, color=ESTILO["texto2"])

    leg = [mpatches.Patch(color="#4ade80", label="Subió"),
           mpatches.Patch(color="#f87171", label="Bajó"),
           mpatches.Patch(color="#94a3b8", label="Sin cambio")]
    ax.legend(handles=leg, facecolor=ESTILO["fondo"],
              edgecolor=ESTILO["borde"], labelcolor=ESTILO["texto"], fontsize=8)

    fig.tight_layout()
    return fig

def grafica_binarias(df_bin, titulo, nombre_colegio, color, sufijo):
    """
    Barras agrupadas para variables binarias (apoyo / experiencias).
    Muestra % que marcó 'sí' en pre y post.
    """
    etiquetas = df_bin["Variable"].tolist()
    pct_pre   = df_bin["% PRE sí"].tolist()
    pct_post  = df_bin["% POST sí"].tolist()

    x = np.arange(len(etiquetas))
    fig, ax = plt.subplots(figsize=(10, 4.5), facecolor=ESTILO["fondo"])
    aplicar_estilo(fig, [ax])

    ax.bar(x - 0.2, pct_pre,  0.38, label="Pre-test",
           color=ESTILO["gris_pre"], edgecolor="none")
    ax.bar(x + 0.2, pct_post, 0.38, label="Post-test",
           color=color, edgecolor="none")

    ax.set_xticks(x)
    ax.set_xticklabels(etiquetas, rotation=20, ha="right", fontsize=8.5,
                       color=ESTILO["texto"])
    ax.set_ylim(0, 115)
    ax.set_ylabel("% que marcó 'sí'", fontsize=9)
    ax.set_title(f"{titulo} · {nombre_colegio}", fontsize=10,
                 color=ESTILO["texto"], pad=10)
    ax.yaxis.grid(True, color=ESTILO["borde"], linewidth=0.5, linestyle="--")
    ax.set_axisbelow(True)
    ax.legend(facecolor=ESTILO["fondo"], edgecolor=ESTILO["borde"],
              labelcolor=ESTILO["texto"], fontsize=8)

    fig.tight_layout()
    return fig

def grafica_radar(resultados_escalas, nombre_colegio):
    """
    Gráfico de radar (araña) comparando PRE vs POST en las 4 escalas.
    """
    nombres = list(resultados_escalas.keys())
    N = len(nombres)
    medias_pre  = [r["media_pre"]  for r in resultados_escalas.values()]
    medias_post = [r["media_post"] for r in resultados_escalas.values()]

    angulos = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    # Cerrar el polígono repitiendo el primer punto
    medias_pre  = medias_pre  + medias_pre[:1]
    medias_post = medias_post + medias_post[:1]
    angulos     = angulos + angulos[:1]

    fig, ax = plt.subplots(figsize=(6, 5), subplot_kw=dict(polar=True),
                           facecolor=ESTILO["fondo"])
    ax.set_facecolor(ESTILO["fondo2"])

    ax.plot(angulos, medias_pre,  color=ESTILO["gris_pre"],
            linewidth=2, linestyle="--", label="Pre-test")
    ax.fill(angulos, medias_pre, color=ESTILO["gris_pre"], alpha=0.15)

    ax.plot(angulos, medias_post, color="#c084fc",
            linewidth=2, label="Post-test")
    ax.fill(angulos, medias_post, color="#c084fc", alpha=0.2)

    ax.set_xticks(angulos[:-1])
    ax.set_xticklabels(nombres, size=9, color=ESTILO["texto"])
    ax.set_ylim(1, 5)
    ax.set_yticks([2, 3, 4, 5])
    ax.set_yticklabels(["2","3","4","5"], size=7, color=ESTILO["texto2"])
    ax.grid(color=ESTILO["borde"], linewidth=0.7)
    ax.spines["polar"].set_color(ESTILO["borde"])

    ax.legend(loc="upper right", bbox_to_anchor=(1.35, 1.1),
              facecolor=ESTILO["fondo"], edgecolor=ESTILO["borde"],
              labelcolor=ESTILO["texto"], fontsize=9)
    ax.set_title(f"Radar Pre vs Post · {nombre_colegio}", size=10,
                 color=ESTILO["texto"], pad=20)

    fig.tight_layout()
    return fig

# ─────────────────────────────────────────────────────────────────────────────
#  6. FUNCIÓN DE REPORTE EN CONSOLA
# ─────────────────────────────────────────────────────────────────────────────

def separador(char="─", ancho=70):
    print(char * ancho)

def imprimir_resultado_escala(nombre, res):

    sig = "✅ SIGNIFICATIVO" if res["significativo"] else "❌ No significativo"

    delta_signo = (
        f"+{res['delta']:.3f}"
        if res['delta'] >= 0
        else f"{res['delta']:.3f}"
    )

    st.markdown(f"## 📐 {nombre}")

    col1, col2 = st.columns(2)

    with col1:
        st.metric(
            "PRE",
            f"{res['media_pre']:.3f}",
            help=f"DE = {res['de_pre']:.3f}"
        )

    with col2:
        st.metric(
            "POST",
            f"{res['media_post']:.3f}",
            delta=delta_signo,
            help=f"DE = {res['de_post']:.3f}"
        )

    st.write(f"t({res['n']-1}) = {res['t_stat']}")
    st.write(f"p = {res['p_valor']} — {sig}")
    st.write(f"Cohen's d = {res['d_cohen']} ({res['magnitud']})")
    st.write(f"IC 95%: [{res['ic_95'][0]}, {res['ic_95'][1]}]")

    st.divider()
# ─────────────────────────────────────────────────────────────────────────────
#  7. FUNCIÓN PRINCIPAL: ANALIZAR UN COLEGIO
# ─────────────────────────────────────────────────────────────────────────────

def analizar_colegio(nombre_colegio, ruta_archivo, carpeta_out):
    """
    Carga los datos de un colegio y ejecuta el análisis completo:
    - Estadísticos descriptivos sociodemográficos
    - T-test pareado por cada escala Likert
    - Análisis ítem a ítem
    - Análisis de variables binarias (apoyo y experiencias)
    - Genera gráficas y guarda CSVs de resultados
    """
    separador("═")
    print(f"  ANALIZANDO: {nombre_colegio}")
    separador("═")

    # ── Crear carpeta de salida para este colegio ──
    # No necesitamos crear la carpeta si no guardamos archivos gráficos,
    # pero sí para CSVs y TXTs si se mantienen.
    nombre_limpio = nombre_colegio.replace(" ", "_").replace("ó","o").replace("é","e").replace("í","i")
    carpeta_colegio = os.path.join(carpeta_out, nombre_limpio)
    os.makedirs(carpeta_colegio, exist_ok=True)

    # ── Cargar datos ──
    print(f"\n  Cargando datos desde: {ruta_archivo}")
    pre, post = cargar_excel(ruta_archivo)
    n = len(post)
    print(f"  Participantes: {n} (post-test) / {len(pre)} (pre-test)")

    # ── Sociodemográfico ──
    separador()
    st.markdown("""
<h2 style="
    color:#e2e8f0;
    margin-top:40px;
    margin-bottom:20px;
">
📊 Perfil Sociodemográfico
</h2>
""", unsafe_allow_html=True)
    separador()
    sd = estadisticos_sociodem(pre, post)
    edad = sd["edad"]
    st.write(f"  Edad: {edad['media']} ± {edad['de']} años  (rango {edad['min']}–{edad['max']})")
    if "tipo_colegio" in sd:
        st.write(f"\n  Tipo de colegio:")
        st.write(sd["tipo_colegio"].to_string())
    if "padres_STEM" in sd:
        st.write(f"\n  Padres trabajan en STEM (1=Sí, 2=No, 3=No sé):")
        st.write(sd["padres_STEM"].to_string())
    if "conoce_mujer_STEM" in sd:
        st.write(f"\n  Conoce una mujer en STEM (1=Sí, 2=No):")
        st.write(sd["conoce_mujer_STEM"].to_string())

    # ── Análisis escalas Likert ──
    separador()
    st.write("  ANÁLISIS PRE-POST POR ESCALA (t-test pareado)")
    separador()

    resultados_escalas = {}
    tablas_items = {}

    for nombre_escala, config in ESCALAS.items():
        cols   = config["columnas"]
        etiq   = config["etiquetas"]
        color  = config["color"]

        # T-test de la escala completa
        res = test_t_pareado(pre, post, cols)
        if "error" in res:
            st.write(f"\n  ⚠️  {nombre_escala}: {res['error']}")
            continue

        resultados_escalas[nombre_escala] = res
        imprimir_resultado_escala(nombre_escala, res)

        # Ítem a ítem
        df_items = test_item_a_item(pre, post, cols, etiq)
        tablas_items[nombre_escala] = df_items

        # Guardar CSV de ítems
        ruta_csv = os.path.join(carpeta_colegio,
                                f"items_{nombre_escala.replace(' ','_')}.csv")
        df_items.to_csv(ruta_csv, index=False, encoding="utf-8-sig")

        # Gráficas de ítems y cambios individuales
        grafica_items(df_items, nombre_escala, nombre_colegio, color)
        grafica_cambios_individuales(pre, post, cols, nombre_escala,
                                     nombre_colegio, color)

    # ── Gráficas de resumen ──
    if resultados_escalas:
        grafica_resumen_escalas(resultados_escalas, nombre_colegio)
        grafica_radar(resultados_escalas, nombre_colegio)

    # ── Variables binarias: Apoyo y Experiencias ──
    separador()
    st.write("  APOYO PERCIBIDO (variables binarias 0/1)")
    separador()
    df_apoyo = analizar_binarias(pre, post, APOYO_COLS)
    st.write(df_apoyo.to_string(index=False))
    df_apoyo.to_csv(
        os.path.join(carpeta_colegio, "apoyo_percibido.csv"),
        index=False, encoding="utf-8-sig"
    )
    grafica_binarias(df_apoyo, "Apoyo percibido", nombre_colegio,
                     "#c084fc", "apoyo")

    separador()
    st.write   ("  EXPERIENCIAS STEM PREVIAS (variables binarias 0/1)")
    separador()
    df_exp = analizar_binarias(pre, post, EXP_COLS)
    st.write(df_exp.to_string(index=False))
    df_exp.to_csv(
        os.path.join(carpeta_colegio, "experiencias_STEM.csv"),
        index=False, encoding="utf-8-sig"
    )
    grafica_binarias(df_exp, "Experiencias STEM", nombre_colegio,
                     "#f97316", "experiencias")

    # ── Tabla resumen general ──
    separador()
    st.write("  TABLA RESUMEN GENERAL")
    separador()
    filas_resumen = []
    for nombre_escala, res in resultados_escalas.items():
        delta_str = f"+{res['delta']:.3f}" if res['delta'] >= 0 else f"{res['delta']:.3f}"
        filas_resumen.append({
            "Escala":       nombre_escala,
            "PRE media":    res["media_pre"],
            "PRE DE":       res["de_pre"],
            "POST media":   res["media_post"],
            "POST DE":      res["de_post"],
            "Δ":            res["delta"],
            "t":            res["t_stat"],
            "p-valor":      res["p_valor"],
            "Sig. (p<.05)": "Sí" if res["significativo"] else "No",
            "Cohen's d":    res["d_cohen"],
            "Magnitud":     res["magnitud"],
            "IC 95% inf":   res["ic_95"][0],
            "IC 95% sup":   res["ic_95"][1],
            "n":            res["n"],
        })
    df_resumen = pd.DataFrame(filas_resumen)
    st.write(df_resumen[["Escala","PRE media","POST media","Δ","p-valor",
                       "Sig. (p<.05)","Cohen's d","Magnitud"]].to_string(index=False))

    ruta_resumen = os.path.join(carpeta_colegio, "resumen_general.csv")
    df_resumen.to_csv(ruta_resumen, index=False, encoding="utf-8-sig")
    st.write(f"\n  ✅  Tabla resumen guardada en: {ruta_resumen}")

    # ── Respuestas abiertas ──
    separador()
    st.write("  RESPUESTAS ABIERTAS")
    separador()
    for col, titulo in [
        ("Futuro_imaginado_STEM", "Futuro imaginado STEM (PRE)"),
        ("Futuro_imaginado_STEM", "Futuro imaginado STEM (POST)"),
        ("Sobre los talleres",    "Lo que más gustó de los talleres"),
        ("Para mejorar",          "Qué mejorarían"),
    ]:
        df_fuente = pre if "PRE" in titulo else post
        if col in df_fuente.columns:
            respuestas = df_fuente[col].dropna().reset_index(drop=True)
            ruta_txt = os.path.join(
                carpeta_colegio,
                f"respuestas_{titulo[:30].replace(' ','_')}.txt"
            )
            with open(ruta_txt, "w", encoding="utf-8") as f:
                f.write(f"=== {titulo} · {nombre_colegio} ===\n\n")
                for i, resp in enumerate(respuestas, 1):
                    f.write(f"[{i:02d}] {resp}\n\n")
            st.write(f"  Respuestas '{titulo}' guardadas en: {ruta_txt}")

    return df_resumen, resultados_escalas

# ─────────────────────────────────────────────────────────────────────────────
#  8. COMPARACIÓN ENTRE COLEGIOS (cuando hay más de uno)
# ─────────────────────────────────────────────────────────────────────────────

def comparar_colegios(todos_resultados, carpeta_out):
    """
    Cuando hay más de un colegio, genera una gráfica comparativa
    del delta (cambio post-pre) en cada escala.
    """
    if len(todos_resultados) < 2:
        return

    separador("═")
    st.write("  COMPARACIÓN ENTRE COLEGIOS")
    separador("═")

    nombres_colegios = list(todos_resultados.keys())
    nombres_escalas  = list(list(todos_resultados.values())[0].keys())

    x = np.arange(len(nombres_escalas))
    ancho = 0.8 / len(nombres_colegios)
    colores = ["#c084fc", "#60a5fa", "#34d399", "#f97316"]

    fig, ax = plt.subplots(figsize=(11, 5), facecolor=ESTILO["fondo"])
    aplicar_estilo(fig, [ax])

    for i, (nombre_col, res_col) in enumerate(todos_resultados.items()):
        deltas = [res_col[esc]["delta"] for esc in nombres_escalas]
        offset = (i - len(nombres_colegios)/2 + 0.5) * ancho
        ax.bar(x + offset, deltas, ancho, label=nombre_col,
               color=colores[i % len(colores)], edgecolor="none")

    ax.axhline(0, color=ESTILO["texto2"], linewidth=1, linestyle="--")
    ax.set_xticks(x)
    ax.set_xticklabels(nombres_escalas, rotation=15, ha="right",
                       fontsize=9, color=ESTILO["texto"])
    ax.set_ylabel("Δ (post − pre)", fontsize=9)
    ax.set_title("Comparación del cambio (Δ) entre colegios",
                 fontsize=11, color=ESTILO["texto"], pad=12)
    ax.yaxis.grid(True, color=ESTILO["borde"], linewidth=0.5, linestyle="--")
    ax.set_axisbelow(True)
    ax.legend(facecolor=ESTILO["fondo"], edgecolor=ESTILO["borde"],
              labelcolor=ESTILO["texto"], fontsize=9)

    fig.tight_layout()
    st.pyplot(fig) # Display the plot
    plt.close(fig) # Close the figure to free up memory

# ─────────────────────────────────────────────────────────────────────────────
#  9. PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # ── Carpeta donde están los Excel y donde se guardarán los resultados ──
    # Por defecto: la misma carpeta donde está este script.
    # Puedes cambiarla si tus archivos están en otro lado:
    #   CARPETA_DATOS = "C:/Users/tu_usuario/Desktop/datos_proyecto"
    # En Colab, puedes usar os.getcwd() para el directorio actual, o montar Google Drive:
    # from google.colab import drive
    # drive.mount('/content/drive')
    # CARPETA_DATOS = '/content/drive/MyDrive/ruta/a/tus/archivos'
    CARPETA_DATOS = os.getcwd()
    CARPETA_OUT   = os.path.join(CARPETA_DATOS, "resultados")
    os.makedirs(CARPETA_OUT, exist_ok=True) # Keep this to save CSVs and TXTs

    separador("═", 70)
    print("  LA CIENCIA SÍ ES COSA DE CHICAS — ANÁLISIS DE DATOS")
    print("  Universidad del Norte · Barranquilla, Colombia")
    separador("═", 70)
    print(f"  Carpeta de datos:     {CARPETA_DATOS}")
    print(f"  Carpeta de salida:    {CARPETA_OUT}")
    print(f"  Colegios a analizar:  {list(COLEGIOS.keys())}")

    todos_resultados = {}

    for nombre_colegio, nombre_archivo in COLEGIOS.items():
        # Si el archivo está en una ruta absoluta de Windows, necesitarás ajustarlo
        # para el entorno de Colab (por ejemplo, si lo subiste a /content/ o Drive)
        ruta = os.path.join(CARPETA_DATOS, nombre_archivo)
        try:
            _, resultados_escalas = analizar_colegio(
                nombre_colegio, ruta, CARPETA_OUT
            )
            todos_resultados[nombre_colegio] = resultados_escalas
        except FileNotFoundError as e:
            print(f"\n  ⚠️  ARCHIVO NO ENCONTRADO: {e}")
            print(f"     Verifica que '{nombre_archivo}' esté en: {CARPETA_DATOS}")
        except Exception as e:
            print(f"\n  ❌  Error inesperado en {nombre_colegio}: {e}")
            raise  # En desarrollo, muestra el traceback completo

    # Comparación entre colegios si hay más de uno
    comparar_colegios(todos_resultados, CARPETA_OUT)

    separador("═", 70)
    st.write("\n  ✅  ANÁLISIS COMPLETADO")
    print(f"     Todos los resultados están en: {CARPETA_OUT}")
    separador("═", 70)

# ─────────────────────────────────────────────
# INTERPRETACIÓN CUALITATIVA
# ─────────────────────────────────────────────

def interpretacion_cualitativa(colegio):

    interpretaciones = {

        "Colegio Eucarístico": {

            "🧬 Aspiraciones STEM":
            """
            Las respuestas muestran un cambio interesante hacia profesiones vinculadas con medicina,
            biología, medio ambiente, investigación e ingeniería biomédica.

            Aunque algunas estudiantes aún manifiestan incertidumbre respecto a su futuro profesional,
            la mayoría logra establecer alguna conexión entre sus intereses y el ámbito científico.
            Un elemento particularmente significativo es la aparición de discursos relacionados con
            el impacto social de la ciencia, especialmente en temas de salud y bienestar. Algunas
            estudiantes mencionan investigar enfermedades o contribuir a resolver problemas que afectan
            a las mujeres.
            """,

            "📚 Percepción de los talleres":
            """
            Los temas más valorados fueron programación, ciencia de datos, trabajo en equipo,
            visitas académicas y actividades recreativas.

            Las respuestas muestran una combinación equilibrada entre la valoración del contenido
            académico y la experiencia social de aprendizaje.
            """,

            "💡 Oportunidades de mejora":
            """
            Las principales sugerencias fueron aumentar el tiempo de las actividades, realizar más
            visitas universitarias, incorporar más actividades tecnológicas y desarrollar talleres
            más extensos.

            La repetición de estas solicitudes indica que las estudiantes percibieron las actividades
            como insuficientes en duración para profundizar en los temas.
            """,

            "📈 Interpretación general":
            """
            En este colegio se observa una apropiación importante de la relación entre ciencia y sociedad.
            Las estudiantes no solo hablan de carreras STEM, sino también de investigación, salud y
            contribución social, lo cual refleja una comprensión más amplia del papel de la ciencia.
            """
        },

        "Colegio Alemán": {

            "🧬 Aspiraciones STEM":
            """
            Las respuestas muestran un grupo con alta diversidad vocacional. No todas las estudiantes
            se proyectan directamente en carreras STEM, pero incluso entre quienes manifiestan dudas
            aparecen referencias claras a biología marina, genética, química, geología, investigación
            científica y trabajo en laboratorio.

            Varias estudiantes expresan que antes no contemplaban estas áreas y ahora pueden imaginarse
            dentro de contextos científicos concretos.
            """,

            "📚 Percepción de los talleres":
            """
            Los elementos más valorados fueron ciencia de datos, probabilidad y estadística, algoritmos,
            trabajo colaborativo y la libertad para tomar decisiones durante las actividades.

            Las respuestas sugieren que las estudiantes apreciaron especialmente aquellas experiencias
            que les permitieron participar activamente en la construcción del conocimiento.
            """,

            "💡 Oportunidades de mejora":
            """
            Las principales recomendaciones fueron realizar más actividades, aumentar la duración y
            hacer las sesiones más dinámicas.

            Varias estudiantes manifestaron no tener sugerencias de mejora, lo cual suele interpretarse
            como un alto nivel de satisfacción con la experiencia.
            """,

            "📈 Interpretación general":
            """
            El programa parece haber contribuido a que las estudiantes relacionen la ciencia con
            experiencias concretas y cercanas. STEM deja de percibirse como un conjunto abstracto de
            disciplinas y comienza a asociarse con profesiones e investigaciones reales.
            """
        },

        "Colegio Enseñanza": {

            "🧬 Aspiraciones STEM":
            """
            Este colegio presenta una de las evidencias cualitativas más fuertes de orientación vocacional
            STEM. Las estudiantes mencionan con frecuencia carreras como astrofísica, química, medicina,
            bacteriología, ingeniería biomédica, ingeniería matemática, sistemas y tecnología.

            Las respuestas son detalladas y muestran una visualización concreta de trayectorias
            académicas y profesionales futuras.
            """,

            "📚 Percepción de los talleres":
            """
            Los aspectos más valorados fueron ciencia de datos, análisis de información, laboratorios,
            la visita al museo y las aplicaciones reales de las matemáticas.

            Varias estudiantes destacan que descubrieron usos prácticos de conceptos que anteriormente
            consideraban abstractos.
            """,

            "💡 Oportunidades de mejora":
            """
            Las recomendaciones se concentraron en aumentar el tiempo de las actividades, incorporar
            más física, más biología y realizar más experimentos.

            Resulta llamativo que las sugerencias busquen ampliar el programa y no modificar su estructura
            principal.
            """,

            "📈 Interpretación general":
            """
            Las respuestas sugieren que el programa no solo fortaleció el interés por STEM, sino que
            ayudó a las estudiantes a visualizar trayectorias académicas específicas. La capacidad de
            nombrar carreras concretas constituye un indicador importante de ampliación de horizontes
            vocacionales.
            """
        },

        "Colegio Sonia Ahumada": {

            "🧬 Aspiraciones STEM":
            """
            Las respuestas evidencian un impacto importante en la percepción de las carreras científicas.
            Las estudiantes mencionan ingeniería, medicina, tecnología, arqueología, neurociencia,
            psicología y ciencias cognitivas.

            Incluso quienes indican no querer una carrera científica reconocen la importancia de la
            ciencia y la tecnología en la vida cotidiana.
            """,

            "📚 Percepción de los talleres":
            """
            Los aspectos más destacados fueron ciencia de datos, computación, matemáticas aplicadas,
            las visitas al museo y las metodologías de enseñanza utilizadas.

            Muchas respuestas resaltan que las actividades lograron mostrar una versión más atractiva
            y accesible de la ciencia.
            """,

            "💡 Oportunidades de mejora":
            """
            La mayoría de las estudiantes manifestó satisfacción con el programa. Las pocas sugerencias
            incluyeron una mayor interacción con estudiantes de otros colegios, más encuentros y más
            actividades colaborativas.

            Esto indica que el interés se orienta a ampliar la experiencia más que a modificar sus
            contenidos.
            """,

            "📈 Interpretación general":
            """
            Las respuestas muestran una transformación importante en la percepción de STEM como un
            campo cercano, accesible y relevante. Las estudiantes enfatizan repetidamente la motivación
            y la inspiración recibidas durante el programa.
            """
        }

    }

    return interpretaciones.get(
        colegio,
        {"Información": "No existe interpretación cualitativa disponible."}
    )