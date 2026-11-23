import hmac
import math
import os
import re
import time
from datetime import timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

try:
    from google import genai
    from google.genai import types as genai_types
except ImportError:
    genai = None
    genai_types = None

# Configuración de página
st.set_page_config(page_title="Monitor de Rendimiento de Procesos", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def get_config_value(name, default=None):
    value = os.getenv(name)
    if value:
        return value
    try:
        return st.secrets.get(name, default)
    except FileNotFoundError:
        return default


app_password = get_config_value("APP_PASSWORD")
if not isinstance(app_password, str) or len(app_password) < 16:
    st.error("Acceso bloqueado: configura APP_PASSWORD con al menos 16 caracteres en .env o en los secretos de la app.")
    st.stop()

if not st.session_state.get("acceso_autorizado", False):
    st.title("Monitor de Rendimiento de Procesos")
    with st.form("formulario_inicio_sesion"):
        password_input = st.text_input("Contraseña de acceso", type="password")
        iniciar_sesion = st.form_submit_button("Iniciar sesión")
    if iniciar_sesion:
        if hmac.compare_digest(password_input.encode("utf-8"), app_password.encode("utf-8")):
            st.session_state["acceso_autorizado"] = True
            st.rerun()
        st.error("Contraseña incorrecta.")
    st.stop()

if st.sidebar.button("Cerrar sesión", key="cerrar_sesion"):
    st.session_state["acceso_autorizado"] = False
    st.rerun()

# Aplicar reinicio diferido de filtros ANTES de instanciar cualquier widget
if st.session_state.pop("_pendiente_reset_filtros", False):
    for _k in (
        "filtro_mes_directo",
        "filtro_meses",
        "filtro_grupos",
        "filtro_horarios",
        "toggle_rango_fechas",
        "selector_franja_tab1",
        "ultimo_clic_tab1",
        "selector_franja_tab3",
        "ultimo_clic_tab3",
        "aud_feriados_sel",
    ):
        st.session_state.pop(_k, None)
    st.session_state["toggle_todos_meses"] = True

# Estilo sobrio para filtros y etiquetas en la sidebar
st.markdown(
    """
    <style>
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {
            background-color: #262730;
            border: 1px solid #4a4d5e;
            color: #f0f2f6;
        }
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] svg {
            fill: #c7ccd6;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

ORDEN_MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

COLUMNAS_REQUERIDAS = [
    "Rango Horario",
    "JOBNAME",
    "Hora Inicio",
    "Hora Finalización",
    "Tiempo de Demora",
    "Grupo",
    "Veces de Repetición",
]

PATRONES_POR_CATEGORIA = {
    "Diario": [
        "Diario (Todos los días, sin excepciones)",
        "Diario (No corre domingos)",
        "Diario (No corre sábados)",
        "Diario (No corre sábados ni domingos)",
        "Diario (No corre feriados)",
        "Diario (No corre domingos, No corre feriados)",
        "Diario (No corre sábados ni domingos, No corre feriados)",
        "Personalizado...",
    ],
    "Mensual": [
        "Mensual (Cierre/Inicio: días 30, 31, fin de mes o 1)",
        "Mensual (Otro día del mes)",
        "Personalizado...",
    ],
    "Semanal / Otros": [
        "Semanal / Quincenal",
        "Esporádico / A demanda",
        "Personalizado...",
    ],
}


def clave_orden_mes(nombre_mes):
    partes = str(nombre_mes).strip().split()
    if len(partes) >= 2:
        mes_num = ORDEN_MESES.get(partes[0].lower(), 99)
        try:
            anio = int(partes[-1])
        except ValueError:
            anio = 9999
        return (anio, mes_num, str(nombre_mes))
    return (9999, 99, str(nombre_mes))


def clave_orden_horario(rango):
    nums = re.findall(r"\d+", str(rango))
    if len(nums) >= 2:
        return (int(nums[0]), int(nums[1]), str(rango))
    if len(nums) == 1:
        return (int(nums[0]), 0, str(rango))
    return (99, 99, str(rango))


def extraer_mes_de_archivo(nombre_archivo):
    """Extrae el número de mes (1-12) si el archivo sigue el patrón procesos_<Mes>_por_dia.xlsx."""
    stem = Path(str(nombre_archivo)).stem
    prefijo = "procesos_"
    sufijo = "_por_dia"
    if stem.lower().startswith(prefijo) and stem.lower().endswith(sufijo):
        mes_txt = stem[len(prefijo):-len(sufijo)].strip().lower()
        return ORDEN_MESES.get(mes_txt)
    return None


def formato_hhmmss(horas_dec):
    if pd.isna(horas_dec):
        return "00:00:00"
    total_seg = max(0, int(round(float(horas_dec) * 3600)))
    h, rem = divmod(total_seg, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def formato_seg_hhmmss(segundos):
    if pd.isna(segundos):
        return "00:00:00"
    val = float(segundos)
    signo = "-" if val < 0 else ""
    total_seg = int(round(abs(val)))
    h, rem = divmod(total_seg, 3600)
    m, s = divmod(rem, 60)
    return f"{signo}{h:02d}:{m:02d}:{s:02d}"


def etiqueta_ym(ym_str):
    try:
        anio, mes = ym_str.split("-")
        return f"{MESES_ES[int(mes)]} {anio}"
    except Exception:
        return str(ym_str)


def aplicar_eje_tiempo_hhmmss(fig, max_horas, eje="y"):
    if max_horas is None or pd.isna(max_horas) or float(max_horas) <= 0:
        return fig
    max_h = float(max_horas)
    pasos = [
        1 / 60, 2 / 60, 5 / 60, 10 / 60, 15 / 60, 30 / 60,
        1, 2, 4, 5, 10, 20, 25, 50, 100, 200, 250, 500, 1000, 2000, 5000, 10000,
    ]
    paso = pasos[-1]
    for p in pasos:
        if max_h / p <= 7:
            paso = p
            break
    n_ticks = int(math.ceil(max_h / paso)) + 1
    tickvals = [round(i * paso, 6) for i in range(n_ticks)]
    ticktext = [formato_hhmmss(v) for v in tickvals]
    cfg = dict(tickmode="array", tickvals=tickvals, ticktext=ticktext)
    if eje == "y":
        fig.update_yaxes(**cfg)
    else:
        fig.update_xaxes(**cfg)
    return fig


def aplicar_eje_entero(fig, eje="y"):
    cfg = dict(tickformat="d", separatethousands=False)
    if eje == "y":
        fig.update_yaxes(**cfg)
    else:
        fig.update_xaxes(**cfg)
    return fig


@st.cache_resource
def get_engine():
    db_url = get_config_value("SUPABASE_URL")
    if not db_url:
        raise ValueError("Falta SUPABASE_URL en el archivo .env o en los secretos de la app.")
    return create_engine(db_url)


engine = get_engine()


# --- GESTIÓN DE CLASIFICACIONES MANUALES Y PROCESOS EN SUPABASE ---
def asegurar_tabla_clasificacion_manual():
    q = text("""
        CREATE TABLE IF NOT EXISTS public.procesos_clasificacion_manual (
            jobname TEXT PRIMARY KEY,
            categoria_manual TEXT NOT NULL,
            patron_manual TEXT,
            notas TEXT,
            actualizado_en TIMESTAMPTZ DEFAULT NOW()
        )
    """)
    with engine.begin() as conn:
        conn.execute(q)


@st.cache_data(ttl=600)
def cargar_clasificaciones_manuales():
    try:
        asegurar_tabla_clasificacion_manual()
        q = text("""
            SELECT jobname, categoria_manual, patron_manual, notas
            FROM public.procesos_clasificacion_manual
        """)
        with engine.connect() as conn:
            rows = conn.execute(q).fetchall()
        return {
            str(r.jobname): {
                "categoria_manual": str(r.categoria_manual),
                "patron_manual": str(r.patron_manual or ""),
                "notas": str(r.notas or ""),
            }
            for r in rows
        }
    except Exception:
        return {}


def guardar_clasificacion_manual(jobname, categoria_manual, patron_manual, notas=""):
    asegurar_tabla_clasificacion_manual()
    q = text("""
        INSERT INTO public.procesos_clasificacion_manual (
            jobname, categoria_manual, patron_manual, notas, actualizado_en
        )
        VALUES (:jobname, :categoria, :patron, :notas, NOW())
        ON CONFLICT (jobname) DO UPDATE SET
            categoria_manual = EXCLUDED.categoria_manual,
            patron_manual = EXCLUDED.patron_manual,
            notas = EXCLUDED.notas,
            actualizado_en = NOW()
    """)
    with engine.begin() as conn:
        conn.execute(
            q,
            {
                "jobname": str(jobname).strip(),
                "categoria": str(categoria_manual).strip(),
                "patron": str(patron_manual).strip(),
                "notas": str(notas or "").strip(),
            },
        )


def eliminar_clasificacion_manual(jobnames):
    if not jobnames:
        return 0
    asegurar_tabla_clasificacion_manual()
    q = text("DELETE FROM public.procesos_clasificacion_manual WHERE jobname = ANY(:jobs)")
    with engine.begin() as conn:
        res = conn.execute(q, {"jobs": list(jobnames)})
        return res.rowcount or 0


@st.cache_data(ttl=600)
def cargar_resumen_jobs_gestion():
    q = text("""
        SELECT jobname,
               STRING_AGG(DISTINCT grupo, ', ' ORDER BY grupo) AS grupo,
               COUNT(*) AS registros,
               COUNT(DISTINCT fecha) AS dias_ejecutados,
               COUNT(DISTINCT TO_CHAR(fecha, 'YYYY-MM')) AS meses_activos,
               SUM(tiempo_demora_seg) AS seg_totales
        FROM public.ejecuciones_procesos
        WHERE jobname IS NOT NULL
        GROUP BY jobname
        ORDER BY jobname
    """)
    with engine.connect() as conn:
        return pd.read_sql(q, conn)


def editar_metadatos_proceso_supabase(jobname_actual, nuevo_jobname, nuevo_grupo):
    job_old = str(jobname_actual).strip()
    job_new = str(nuevo_jobname).strip()
    grp_new = str(nuevo_grupo).strip()

    with engine.begin() as conn:
        if grp_new:
            res = conn.execute(
                text("""
                    UPDATE public.ejecuciones_procesos
                    SET jobname = :job_new, grupo = :grp_new
                    WHERE jobname = :job_old
                """),
                {"job_new": job_new, "grp_new": grp_new, "job_old": job_old},
            )
        else:
            res = conn.execute(
                text("""
                    UPDATE public.ejecuciones_procesos
                    SET jobname = :job_new
                    WHERE jobname = :job_old
                """),
                {"job_new": job_new, "job_old": job_old},
            )

        if job_old != job_new:
            try:
                existe_nuevo = conn.execute(
                    text("SELECT 1 FROM public.procesos_clasificacion_manual WHERE jobname = :jn"),
                    {"jn": job_new},
                ).fetchone()
                if existe_nuevo:
                    conn.execute(
                        text("DELETE FROM public.procesos_clasificacion_manual WHERE jobname = :jo"),
                        {"jo": job_old},
                    )
                else:
                    conn.execute(
                        text("""
                            UPDATE public.procesos_clasificacion_manual
                            SET jobname = :jn, actualizado_en = NOW()
                            WHERE jobname = :jo
                        """),
                        {"jn": job_new, "jo": job_old},
                    )
            except Exception:
                pass

        return res.rowcount or 0


def eliminar_procesos_supabase(jobnames, meses_a_borrar=None):
    if not jobnames:
        return 0
    lista_jobs = list(jobnames)

    with engine.begin() as conn:
        if meses_a_borrar:
            res = conn.execute(
                text("""
                    DELETE FROM public.ejecuciones_procesos
                    WHERE jobname = ANY(:jobs) AND mes = ANY(:meses)
                """),
                {"jobs": lista_jobs, "meses": list(meses_a_borrar)},
            )
        else:
            res = conn.execute(
                text("DELETE FROM public.ejecuciones_procesos WHERE jobname = ANY(:jobs)"),
                {"jobs": lista_jobs},
            )
            try:
                conn.execute(
                    text("DELETE FROM public.procesos_clasificacion_manual WHERE jobname = ANY(:jobs)"),
                    {"jobs": lista_jobs},
                )
            except Exception:
                pass
        return res.rowcount or 0


# --- HERRAMIENTAS SQL DE SOLO LECTURA PARA EL ANALISTA IA (GEMINI) ---
def _resolver_ym_desde_texto(texto_mes: str):
    """Convierte '2026-02', 'Febrero 2026' o 'febrero' al formato YYYY-MM existente en BD."""
    limpio = str(texto_mes or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}", limpio):
        return limpio

    partes = limpio.lower().split()
    mes_num = None
    anio = None
    for p in partes:
        if p in ORDEN_MESES:
            mes_num = ORDEN_MESES[p]
        elif re.fullmatch(r"\d{4}", p):
            anio = int(p)

    if mes_num is not None and anio is not None:
        return f"{anio:04d}-{mes_num:02d}"

    if mes_num is not None:
        with engine.connect() as conn:
            row = conn.execute(
                text("""
                    SELECT TO_CHAR(fecha, 'YYYY-MM') AS ym
                    FROM public.ejecuciones_procesos
                    WHERE EXTRACT(MONTH FROM fecha)::int = :m
                    ORDER BY fecha DESC
                    LIMIT 1
                """),
                {"m": mes_num},
            ).fetchone()
        if row:
            return str(row[0])
    return None


def ia_obtener_resumen_mensual() -> dict:
    """Devuelve el resumen mensual completo de la base de datos: días cargados, procesos, jobs únicos y tiempo acumulado."""
    q = text("""
        SELECT TO_CHAR(fecha, 'YYYY-MM') AS ym,
               COUNT(DISTINCT fecha) AS dias_cargados,
               COUNT(*) AS total_ejecuciones,
               COUNT(DISTINCT jobname) AS jobs_unicos,
               SUM(tiempo_demora_seg) AS seg_totales
        FROM public.ejecuciones_procesos
        WHERE fecha IS NOT NULL
        GROUP BY TO_CHAR(fecha, 'YYYY-MM')
        ORDER BY ym
    """)
    with engine.connect() as conn:
        rows = conn.execute(q).fetchall()

    meses = []
    for r in rows:
        seg = float(r.seg_totales or 0)
        dias = max(1, int(r.dias_cargados or 1))
        meses.append({
            "ym": str(r.ym),
            "mes": etiqueta_ym(str(r.ym)),
            "dias_cargados": int(r.dias_cargados),
            "total_ejecuciones": int(r.total_ejecuciones),
            "jobs_unicos": int(r.jobs_unicos),
            "tiempo_acum_hhmmss": formato_seg_hhmmss(seg),
            "promedio_diario_hhmmss": formato_seg_hhmmss(seg / dias),
        })
    return {"resumen_mensual": meses}


def ia_comparar_fechas(fecha_a: str, fecha_b: str) -> dict:
    """Compara dos fechas exactas (YYYY-MM-DD): totales, top 15 jobs con mayor variación de tiempo y diferencias por franja horaria."""
    fa = pd.to_datetime(str(fecha_a).strip(), errors="coerce")
    fb = pd.to_datetime(str(fecha_b).strip(), errors="coerce")
    if pd.isna(fa) or pd.isna(fb):
        return {"error": "Usa el formato de fecha YYYY-MM-DD (ej. 2026-10-15)."}
    fa_d, fb_d = fa.date(), fb.date()

    q_tot = text("""
        SELECT fecha, COUNT(*) AS ejecuciones, COUNT(DISTINCT jobname) AS jobs_unicos,
               SUM(tiempo_demora_seg) AS seg_totales
        FROM public.ejecuciones_procesos
        WHERE fecha IN (:fa, :fb)
        GROUP BY fecha
        ORDER BY fecha
    """)

    q_jobs = text("""
        WITH a AS (
            SELECT jobname, grupo, COUNT(*) AS ejec_a, SUM(tiempo_demora_seg) AS seg_a
            FROM public.ejecuciones_procesos WHERE fecha = :fa GROUP BY jobname, grupo
        ),
        b AS (
            SELECT jobname, grupo, COUNT(*) AS ejec_b, SUM(tiempo_demora_seg) AS seg_b
            FROM public.ejecuciones_procesos WHERE fecha = :fb GROUP BY jobname, grupo
        )
        SELECT COALESCE(a.jobname, b.jobname) AS jobname,
               COALESCE(a.grupo, b.grupo) AS grupo,
               COALESCE(a.ejec_a, 0) AS ejec_a,
               COALESCE(b.ejec_b, 0) AS ejec_b,
               COALESCE(a.seg_a, 0) AS seg_a,
               COALESCE(b.seg_b, 0) AS seg_b,
               (COALESCE(b.seg_b, 0) - COALESCE(a.seg_a, 0)) AS delta_seg
        FROM a FULL OUTER JOIN b ON a.jobname = b.jobname AND a.grupo = b.grupo
        ORDER BY ABS(COALESCE(b.seg_b, 0) - COALESCE(a.seg_a, 0)) DESC
        LIMIT 15
    """)

    q_hora = text("""
        WITH a AS (
            SELECT rango_horario, COUNT(*) AS ejec_a, SUM(tiempo_demora_seg) AS seg_a
            FROM public.ejecuciones_procesos WHERE fecha = :fa GROUP BY rango_horario
        ),
        b AS (
            SELECT rango_horario, COUNT(*) AS ejec_b, SUM(tiempo_demora_seg) AS seg_b
            FROM public.ejecuciones_procesos WHERE fecha = :fb GROUP BY rango_horario
        )
        SELECT COALESCE(a.rango_horario, b.rango_horario) AS rango_horario,
               COALESCE(a.ejec_a, 0) AS ejec_a,
               COALESCE(b.ejec_b, 0) AS ejec_b,
               (COALESCE(b.seg_b, 0) - COALESCE(a.seg_a, 0)) AS delta_seg
        FROM a FULL OUTER JOIN b ON a.rango_horario = b.rango_horario
        ORDER BY ABS(COALESCE(b.seg_b, 0) - COALESCE(a.seg_a, 0)) DESC
        LIMIT 8
    """)

    with engine.connect() as conn:
        tot_rows = conn.execute(q_tot, {"fa": fa_d, "fb": fb_d}).fetchall()
        job_rows = conn.execute(q_jobs, {"fa": fa_d, "fb": fb_d}).fetchall()
        hora_rows = conn.execute(q_hora, {"fa": fa_d, "fb": fb_d}).fetchall()

    return {
        "fecha_a": str(fa_d),
        "fecha_b": str(fb_d),
        "totales": [
            {
                "fecha": str(r.fecha),
                "ejecuciones": int(r.ejecuciones),
                "jobs_unicos": int(r.jobs_unicos),
                "tiempo_acum_hhmmss": formato_seg_hhmmss(r.seg_totales),
            }
            for r in tot_rows
        ],
        "top_15_jobs_con_mayor_diferencia": [
            {
                "jobname": str(r.jobname),
                "grupo": str(r.grupo),
                "ejecuciones_fecha_a": int(r.ejec_a),
                "ejecuciones_fecha_b": int(r.ejec_b),
                "tiempo_fecha_a": formato_seg_hhmmss(r.seg_a),
                "tiempo_fecha_b": formato_seg_hhmmss(r.seg_b),
                "diferencia_b_menos_a": formato_seg_hhmmss(r.delta_seg),
            }
            for r in job_rows
        ],
        "top_franjas_con_mayor_variacion": [
            {
                "rango_horario": str(r.rango_horario),
                "ejecuciones_fecha_a": int(r.ejec_a),
                "ejecuciones_fecha_b": int(r.ejec_b),
                "diferencia_tiempo_b_menos_a": formato_seg_hhmmss(r.delta_seg),
            }
            for r in hora_rows
        ],
    }


def ia_comparar_meses(mes_a: str, mes_b: str) -> dict:
    """Compara dos meses (ej. 'Enero 2026' y 'Febrero 2026'): totales, jobs con mayor variación y jobs exclusivos."""
    ym_a = _resolver_ym_desde_texto(mes_a)
    ym_b = _resolver_ym_desde_texto(mes_b)
    if not ym_a or not ym_b:
        return {"error": f"No se pudieron interpretar los meses '{mes_a}' y '{mes_b}'."}

    q_tot = text("""
        SELECT TO_CHAR(fecha, 'YYYY-MM') AS ym,
               COUNT(DISTINCT fecha) AS dias_cargados,
               COUNT(*) AS ejecuciones,
               COUNT(DISTINCT jobname) AS jobs_unicos,
               SUM(tiempo_demora_seg) AS seg_totales
        FROM public.ejecuciones_procesos
        WHERE TO_CHAR(fecha, 'YYYY-MM') IN (:yma, :ymb)
        GROUP BY TO_CHAR(fecha, 'YYYY-MM')
        ORDER BY ym
    """)

    q_jobs = text("""
        WITH a AS (
            SELECT jobname, STRING_AGG(DISTINCT grupo, ',') AS grupo,
                   COUNT(DISTINCT fecha) AS dias_a, COUNT(*) AS ejec_a, SUM(tiempo_demora_seg) AS seg_a
            FROM public.ejecuciones_procesos
            WHERE TO_CHAR(fecha, 'YYYY-MM') = :yma GROUP BY jobname
        ),
        b AS (
            SELECT jobname, STRING_AGG(DISTINCT grupo, ',') AS grupo,
                   COUNT(DISTINCT fecha) AS dias_b, COUNT(*) AS ejec_b, SUM(tiempo_demora_seg) AS seg_b
            FROM public.ejecuciones_procesos
            WHERE TO_CHAR(fecha, 'YYYY-MM') = :ymb GROUP BY jobname
        )
        SELECT COALESCE(a.jobname, b.jobname) AS jobname,
               COALESCE(a.grupo, b.grupo) AS grupo,
               COALESCE(a.dias_a, 0) AS dias_a,
               COALESCE(b.dias_b, 0) AS dias_b,
               COALESCE(a.ejec_a, 0) AS ejec_a,
               COALESCE(b.ejec_b, 0) AS ejec_b,
               COALESCE(a.seg_a, 0) AS seg_a,
               COALESCE(b.seg_b, 0) AS seg_b,
               (COALESCE(b.seg_b, 0) - COALESCE(a.seg_a, 0)) AS delta_seg
        FROM a FULL OUTER JOIN b ON a.jobname = b.jobname
        ORDER BY ABS(COALESCE(b.seg_b, 0) - COALESCE(a.seg_a, 0)) DESC
        LIMIT 15
    """)

    q_exclusivos = text("""
        WITH a AS (
            SELECT DISTINCT jobname FROM public.ejecuciones_procesos WHERE TO_CHAR(fecha, 'YYYY-MM') = :yma
        ),
        b AS (
            SELECT DISTINCT jobname FROM public.ejecuciones_procesos WHERE TO_CHAR(fecha, 'YYYY-MM') = :ymb
        )
        SELECT
            (SELECT COUNT(*) FROM a LEFT JOIN b ON a.jobname = b.jobname WHERE b.jobname IS NULL) AS solo_en_a_count,
            (SELECT COUNT(*) FROM b LEFT JOIN a ON b.jobname = a.jobname WHERE a.jobname IS NULL) AS solo_en_b_count,
            (SELECT ARRAY_AGG(a.jobname ORDER BY a.jobname) FROM (
                SELECT a.jobname FROM a LEFT JOIN b ON a.jobname = b.jobname WHERE b.jobname IS NULL LIMIT 25
            ) a) AS muestra_solo_en_a,
            (SELECT ARRAY_AGG(b.jobname ORDER BY b.jobname) FROM (
                SELECT b.jobname FROM b LEFT JOIN a ON b.jobname = a.jobname WHERE a.jobname IS NULL LIMIT 25
            ) b) AS muestra_solo_en_b
    """)

    with engine.connect() as conn:
        tot_rows = conn.execute(q_tot, {"yma": ym_a, "ymb": ym_b}).fetchall()
        job_rows = conn.execute(q_jobs, {"yma": ym_a, "ymb": ym_b}).fetchall()
        exc_row = conn.execute(q_exclusivos, {"yma": ym_a, "ymb": ym_b}).fetchone()

    return {
        "mes_a": etiqueta_ym(ym_a),
        "mes_b": etiqueta_ym(ym_b),
        "totales": [
            {
                "mes": etiqueta_ym(str(r.ym)),
                "dias_cargados": int(r.dias_cargados),
                "ejecuciones": int(r.ejecuciones),
                "jobs_unicos": int(r.jobs_unicos),
                "tiempo_acum_hhmmss": formato_seg_hhmmss(r.seg_totales),
            }
            for r in tot_rows
        ],
        "jobs_exclusivos": {
            "cantidad_solo_en_mes_a": int(exc_row.solo_en_a_count or 0) if exc_row else 0,
            "cantidad_solo_en_mes_b": int(exc_row.solo_en_b_count or 0) if exc_row else 0,
            "muestra_solo_en_mes_a": list(exc_row.muestra_solo_en_a or []) if exc_row else [],
            "muestra_solo_en_mes_b": list(exc_row.muestra_solo_en_b or []) if exc_row else [],
        },
        "top_15_jobs_mayor_variacion_tiempo": [
            {
                "jobname": str(r.jobname),
                "grupo": str(r.grupo),
                "dias_mes_a": int(r.dias_a),
                "dias_mes_b": int(r.dias_b),
                "tiempo_mes_a": formato_seg_hhmmss(r.seg_a),
                "tiempo_mes_b": formato_seg_hhmmss(r.seg_b),
                "diferencia_b_menos_a": formato_seg_hhmmss(r.delta_seg),
            }
            for r in job_rows
        ],
    }


def ia_detalle_de_fecha(fecha: str) -> dict:
    """Devuelve el detalle operativo de una fecha (YYYY-MM-DD): totales, top 15 jobs más pesados y franjas más cargadas."""
    f = pd.to_datetime(str(fecha).strip(), errors="coerce")
    if pd.isna(f):
        return {"error": "Formato inválido. Usa YYYY-MM-DD."}
    fd = f.date()

    q_tot = text("""
        SELECT COUNT(*) AS ejecuciones, COUNT(DISTINCT jobname) AS jobs_unicos,
               SUM(tiempo_demora_seg) AS seg_totales
        FROM public.ejecuciones_procesos WHERE fecha = :f
    """)
    q_top = text("""
        SELECT jobname, grupo, COUNT(*) AS ejecuciones, SUM(tiempo_demora_seg) AS seg_totales
        FROM public.ejecuciones_procesos WHERE fecha = :f
        GROUP BY jobname, grupo ORDER BY seg_totales DESC LIMIT 15
    """)
    q_franja = text("""
        SELECT rango_horario, COUNT(*) AS ejecuciones, SUM(tiempo_demora_seg) AS seg_totales
        FROM public.ejecuciones_procesos WHERE fecha = :f
        GROUP BY rango_horario ORDER BY seg_totales DESC LIMIT 8
    """)

    with engine.connect() as conn:
        tot = conn.execute(q_tot, {"f": fd}).fetchone()
        tops = conn.execute(q_top, {"f": fd}).fetchall()
        franjas = conn.execute(q_franja, {"f": fd}).fetchall()

    return {
        "fecha": str(fd),
        "ejecuciones": int(tot.ejecuciones or 0) if tot else 0,
        "jobs_unicos": int(tot.jobs_unicos or 0) if tot else 0,
        "tiempo_acum_hhmmss": formato_seg_hhmmss(tot.seg_totales or 0) if tot else "00:00:00",
        "top_15_jobs": [
            {
                "jobname": str(r.jobname),
                "grupo": str(r.grupo),
                "ejecuciones": int(r.ejecuciones),
                "tiempo_acum_hhmmss": formato_seg_hhmmss(r.seg_totales),
            }
            for r in tops
        ],
        "top_franjas": [
            {
                "rango_horario": str(r.rango_horario),
                "ejecuciones": int(r.ejecuciones),
                "tiempo_acum_hhmmss": formato_seg_hhmmss(r.seg_totales),
            }
            for r in franjas
        ],
    }


def ia_analizar_job_o_grupo(termino: str) -> dict:
    """Busca un JOBNAME exacto/parcial o un Grupo y devuelve su evolución mensual, primera/última fecha, días ejecutados y tiempos."""
    t = str(termino or "").strip()
    if not t:
        return {"error": "Indica un nombre de JOBNAME o Grupo."}

    q = text("""
        SELECT TO_CHAR(fecha, 'YYYY-MM') AS ym,
               COUNT(DISTINCT fecha) AS dias_ejecutados,
               COUNT(*) AS ejecuciones,
               SUM(tiempo_demora_seg) AS seg_totales,
               AVG(tiempo_demora_seg) AS seg_prom,
               MAX(tiempo_demora_seg) AS seg_max,
               MIN(fecha) AS primera_fecha_mes,
               MAX(fecha) AS ultima_fecha_mes
        FROM public.ejecuciones_procesos
        WHERE UPPER(jobname) = UPPER(:exacto) OR UPPER(grupo) = UPPER(:exacto)
        GROUP BY TO_CHAR(fecha, 'YYYY-MM')
        ORDER BY ym
    """)

    with engine.connect() as conn:
        rows = conn.execute(q, {"exacto": t}).fetchall()
        if not rows:
            q_like = text("""
                SELECT jobname, grupo, COUNT(DISTINCT fecha) AS dias,
                       MIN(fecha) AS primera_fecha, MAX(fecha) AS ultima_fecha,
                       SUM(tiempo_demora_seg) AS seg_totales
                FROM public.ejecuciones_procesos
                WHERE jobname ILIKE :pat OR grupo ILIKE :pat
                GROUP BY jobname, grupo
                ORDER BY seg_totales DESC
                LIMIT 15
            """)
            similares = conn.execute(q_like, {"pat": f"%{t}%"}).fetchall()
            return {
                "termino": t,
                "coincidencias_parciales": [
                    {
                        "jobname": str(r.jobname),
                        "grupo": str(r.grupo),
                        "dias_ejecutados": int(r.dias),
                        "primera_fecha": str(r.primera_fecha),
                        "ultima_fecha": str(r.ultima_fecha),
                        "tiempo_acum_hhmmss": formato_seg_hhmmss(r.seg_totales),
                    }
                    for r in similares
                ],
            }

    return {
        "termino": t,
        "evolucion_mensual": [
            {
                "mes": etiqueta_ym(str(r.ym)),
                "dias_ejecutados": int(r.dias_ejecutados),
                "ejecuciones": int(r.ejecuciones),
                "tiempo_acum_hhmmss": formato_seg_hhmmss(r.seg_totales),
                "promedio_hhmmss": formato_seg_hhmmss(r.seg_prom),
                "maximo_hhmmss": formato_seg_hhmmss(r.seg_max),
                "rango_fechas": f"{r.primera_fecha_mes} a {r.ultima_fecha_mes}",
            }
            for r in rows
        ],
    }


def ia_revisar_ausencias_mensuales(mes_objetivo: str) -> dict:
    """Identifica qué procesos de tipo mensual (corren <= 2.5 días/mes en el histórico) NO tuvieron ejecución en el mes indicado."""
    ym = _resolver_ym_desde_texto(mes_objetivo)
    if not ym:
        return {"error": f"No se reconoció el mes '{mes_objetivo}'."}

    q = text("""
        WITH por_dia AS (
            SELECT fecha, jobname, STRING_AGG(DISTINCT grupo, ',') AS grupo
            FROM public.ejecuciones_procesos
            WHERE fecha IS NOT NULL AND jobname IS NOT NULL
            GROUP BY fecha, jobname
        ),
        resumen AS (
            SELECT jobname,
                   STRING_AGG(DISTINCT grupo, ',') AS grupo,
                   COUNT(*) AS dias_tot,
                   COUNT(DISTINCT TO_CHAR(fecha, 'YYYY-MM')) AS meses_act,
                   MIN(TO_CHAR(fecha, 'YYYY-MM')) AS ym_ini,
                   MAX(TO_CHAR(fecha, 'YYYY-MM')) AS ym_fin,
                   BOOL_OR(TO_CHAR(fecha, 'YYYY-MM') = :ym) AS corrio_en_mes,
                   ARRAY_AGG(DISTINCT TO_CHAR(fecha, 'YYYY-MM') ORDER BY TO_CHAR(fecha, 'YYYY-MM')) AS meses_lista
            FROM por_dia
            GROUP BY jobname
        )
        SELECT jobname, grupo, dias_tot, meses_act, ym_ini, ym_fin, meses_lista
        FROM resumen
        WHERE meses_act >= 2
          AND (dias_tot::float / GREATEST(meses_act, 1)) <= 2.5
          AND ym_ini <= :ym
          AND ym_fin >= :ym
          AND NOT corrio_en_mes
        ORDER BY meses_act DESC, jobname
        LIMIT 40
    """)
    with engine.connect() as conn:
        rows = conn.execute(q, {"ym": ym}).fetchall()

    return {
        "mes_analizado": etiqueta_ym(ym),
        "cantidad_mensuales_vigentes_sin_corrida_ese_mes": len(rows),
        "jobs_ausentes_en_ese_mes": [
            {
                "jobname": str(r.jobname),
                "grupo": str(r.grupo),
                "meses_activos_totales": int(r.meses_act),
                "vigente_desde": etiqueta_ym(str(r.ym_ini)),
                "ultima_aparicion": etiqueta_ym(str(r.ym_fin)),
            }
            for r in rows
        ],
    }


def _descubrir_modelos_fallback(client, excluir):
    """Busca modelos Gemini Flash alternativos habilitados para la API key actual."""
    candidatos = []
    try:
        for m in client.models.list():
            nombre = (getattr(m, "name", "") or "").replace("models/", "")
            acciones = getattr(m, "supported_actions", None) or []
            if acciones and "generateContent" not in acciones:
                continue
            low = nombre.lower()
            if (
                nombre
                and nombre not in excluir
                and "gemini" in low
                and "flash" in low
                and not any(x in low for x in ("embedding", "image", "tts", "audio", "veo"))
            ):
                candidatos.append(nombre)
    except Exception:
        pass
    return candidatos[:3]


def responder_pregunta_ia(pregunta: str, contexto_actual: dict) -> str:
    if genai is None or genai_types is None:
        return (
            "Falta instalar la librería oficial de Google GenAI. Ejecuta en tu terminal:\n\n"
            "`python -m pip install google-genai`"
        )

    api_key = get_config_value("GEMINI_API_KEY") or get_config_value("GOOGLE_API_KEY")
    if not api_key:
        return (
            "Falta configurar `GEMINI_API_KEY` (o `GOOGLE_API_KEY`) en `.env` o en los secretos "
            "de la app."
        )

    modelo_principal = get_config_value("GEMINI_MODEL", "gemini-3.8-flash").strip()
    extras_env = [
        m.strip()
        for m in get_config_value("GEMINI_FALLBACK_MODELS", "").split(",")
        if m.strip() and m.strip() != modelo_principal
    ]

    client = genai.Client(api_key=api_key)

    prompt_usuario = (
        f"Contexto actual visible en el dashboard:\n{contexto_actual}\n\n"
        f"Pregunta del usuario:\n{pregunta}"
    )

    config_ia = genai_types.GenerateContentConfig(
        system_instruction=(
            "Eres un analista senior de Operaciones de TI que audita procesos batch en Supabase. "
            "Responde en español, de forma directa, ejecutiva y basada EXCLUSIVAMENTE en los datos "
            "obtenidos mediante tus herramientas SQL y el contexto del dashboard. "
            "Usa SIEMPRE formato hh:mm:ss para los tiempos acumulados y enteros sin comas para conteos. "
            "Cuando expliques diferencias entre días o meses, cita los JOBNAME, grupos y franjas horarias "
            "con mayor impacto real. Si los datos no alcanzan para confirmar la causa raíz, dilo claramente."
        ),
        tools=[
            ia_obtener_resumen_mensual,
            ia_comparar_fechas,
            ia_comparar_meses,
            ia_detalle_de_fecha,
            ia_analizar_job_o_grupo,
            ia_revisar_ausencias_mensuales,
        ],
        temperature=0.2,
    )

    modelos_a_probar = [modelo_principal] + extras_env
    descubiertos = False
    ultimo_error = None
    idx = 0

    while idx < len(modelos_a_probar):
        modelo_actual = modelos_a_probar[idx]
        idx += 1

        for intento in range(3):
            try:
                resp = client.models.generate_content(
                    model=modelo_actual,
                    contents=prompt_usuario,
                    config=config_ia,
                )
                texto = resp.text or "No se recibió respuesta del modelo."
                if modelo_actual != modelo_principal:
                    return f"*(Respondido con modelo alternativo `{modelo_actual}` por saturación temporal de `{modelo_principal}`)*\n\n{texto}"
                return texto
            except Exception as e:
                ultimo_error = e
                err_txt = str(e).upper()
                es_temporal = any(
                    tok in err_txt
                    for tok in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "HIGH DEMAND", "OVERLOADED")
                )
                es_no_disponible = "404" in err_txt or "NOT_FOUND" in err_txt

                if es_temporal and intento < 2:
                    time.sleep(2 * (intento + 1))
                    continue

                if (es_temporal or es_no_disponible) and not descubiertos:
                    descubiertos = True
                    for alt in _descubrir_modelos_fallback(client, set(modelos_a_probar)):
                        modelos_a_probar.append(alt)
                break

    return (
        "Los servidores de Gemini están saturados en este momento (`503 UNAVAILABLE`). "
        f"Intenta nuevamente en unos segundos.\n\nDetalle técnico: `{ultimo_error}`"
    )


# --- UTILIDADES DE PARSEO, CARGA Y LIMPIEZA EN BD ---
def parse_time_to_seconds(t_val):
    if pd.isna(t_val):
        return 0.0
    if isinstance(t_val, (timedelta, pd.Timedelta)):
        return float(t_val.total_seconds())
    if hasattr(t_val, "hour"):
        return float(t_val.hour * 3600 + t_val.minute * 60 + t_val.second)
    parts = str(t_val).strip().split(":")
    if len(parts) == 3:
        try:
            return float(int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2]))
        except ValueError:
            return 0.0
    return 0.0


def parse_fecha_hoja(nombre_hoja):
    limpio = str(nombre_hoja).strip()
    for fmt in ("%d-%m-%y", "%d-%m-%Y", "%Y-%m-%d"):
        dt = pd.to_datetime(limpio, format=fmt, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    return None


def nombre_de_fuente_excel(archivo):
    if isinstance(archivo, (str, Path)):
        return Path(archivo).name
    return getattr(archivo, "name", "archivo.xlsx")


def procesar_excels_subidos(archivos_subidos, fecha_manual=None, solo_mes_del_archivo=True):
    """Procesa archivos subidos o rutas Path con deduplicación por fecha y validación de mes."""
    archivos_ordenados = sorted(
        list(archivos_subidos),
        key=lambda a: (
            extraer_mes_de_archivo(nombre_de_fuente_excel(a)) or 99,
            nombre_de_fuente_excel(a).lower(),
        ),
    )

    hojas_por_fecha = {}
    hojas_omitidas = []
    avisos_lote = []

    for archivo in archivos_ordenados:
        nombre_arch = nombre_de_fuente_excel(archivo)
        mes_esperado = extraer_mes_de_archivo(nombre_arch) if solo_mes_del_archivo else None

        if hasattr(archivo, "seek"):
            archivo.seek(0)
        xls = pd.ExcelFile(archivo)
        es_hoja_unica = len(xls.sheet_names) == 1

        for sheet in xls.sheet_names:
            fecha_hoja = parse_fecha_hoja(sheet)
            if fecha_hoja is None:
                if es_hoja_unica and fecha_manual is not None:
                    fecha_hoja = fecha_manual
                else:
                    hojas_omitidas.append(f"{nombre_arch} [{sheet}] (nombre de hoja no es una fecha DD-MM-YY)")
                    continue

            if mes_esperado is not None and fecha_hoja.month != mes_esperado:
                hojas_omitidas.append(
                    f"{nombre_arch} [{sheet}] ({fecha_hoja}: pertenece a {MESES_ES[fecha_hoja.month]}, "
                    f"omitida porque el archivo es de {MESES_ES[mes_esperado]})"
                )
                continue

            df = pd.read_excel(xls, sheet_name=sheet)
            df.columns = [str(c).strip() for c in df.columns]

            faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
            if faltantes:
                hojas_omitidas.append(f"{nombre_arch} [{sheet}] (faltan columnas: {', '.join(faltantes)})")
                continue

            nombre_mes_real = f"{MESES_ES[fecha_hoja.month]} {fecha_hoja.year}"

            df_clean = pd.DataFrame({
                "fecha": fecha_hoja,
                "mes": nombre_mes_real,
                "rango_horario": df["Rango Horario"].astype(str).str.strip(),
                "jobname": df["JOBNAME"].astype(str).str.strip(),
                "hora_inicio": df["Hora Inicio"].astype(str),
                "hora_fin": df["Hora Finalización"].astype(str),
                "tiempo_demora_seg": df["Tiempo de Demora"].apply(parse_time_to_seconds),
                "grupo": df["Grupo"].astype(str).str.strip(),
                "veces_repeticion": pd.to_numeric(
                    df["Veces de Repetición"], errors="coerce"
                ).fillna(1).astype(int),
            })

            df_clean = df_clean.dropna(subset=["fecha"])
            if df_clean.empty:
                continue

            if fecha_hoja in hojas_por_fecha:
                avisos_lote.append(
                    f"Fecha repetida en el lote ({fecha_hoja}): se conservó la versión de '{nombre_arch} [{sheet}]'."
                )
            hojas_por_fecha[fecha_hoja] = df_clean

    if not hojas_por_fecha:
        return pd.DataFrame(), hojas_omitidas, avisos_lote

    fechas_ord = sorted(hojas_por_fecha.keys())
    full_df = pd.concat([hojas_por_fecha[f] for f in fechas_ord], ignore_index=True)
    return full_df, hojas_omitidas, avisos_lote


def subir_datos_a_supabase(df_nuevo, reemplazar_existentes=True):
    fechas_nuevas = sorted(df_nuevo["fecha"].dropna().unique().tolist())
    reemplazados = 0

    with engine.begin() as conn:
        if reemplazar_existentes and fechas_nuevas:
            res = conn.execute(
                text("DELETE FROM public.ejecuciones_procesos WHERE fecha = ANY(:fechas)"),
                {"fechas": fechas_nuevas},
            )
            reemplazados = res.rowcount or 0

        df_nuevo.to_sql(
            name="ejecuciones_procesos",
            con=conn,
            schema="public",
            if_exists="append",
            index=False,
            chunksize=1000,
            method="multi",
        )

    return len(df_nuevo), len(fechas_nuevas), reemplazados


def eliminar_dias_supabase(fechas_a_borrar):
    with engine.begin() as conn:
        res = conn.execute(
            text("DELETE FROM public.ejecuciones_procesos WHERE fecha = ANY(:fechas)"),
            {"fechas": fechas_a_borrar},
        )
        return res.rowcount


@st.cache_data(ttl=600)
def detectar_inconsistencias_meses():
    """Detecta fechas con etiqueta de mes cruzada (ej. 01-10 en Setiembre) o variante 'Setiembre'."""
    q = text("""
        SELECT fecha,
               mes,
               COUNT(*) AS registros
        FROM public.ejecuciones_procesos
        WHERE fecha IS NOT NULL
        GROUP BY fecha, mes
        ORDER BY fecha, mes
    """)
    with engine.connect() as conn:
        df = pd.read_sql(q, conn)

    if df.empty:
        return pd.DataFrame()

    df["fecha"] = pd.to_datetime(df["fecha"]).dt.date
    df["mes_correcto"] = df["fecha"].map(lambda f: f"{MESES_ES[f.month]} {f.year}")

    def mes_num_de_texto(txt):
        partes = str(txt).strip().split()
        if not partes:
            return None
        return ORDEN_MESES.get(partes[0].lower())

    df["mes_num_etiqueta"] = df["mes"].map(mes_num_de_texto)
    df["es_mes_cruzado"] = df["mes_num_etiqueta"] != df["fecha"].map(lambda f: f.month)

    fechas_con_mes_real = set(df[~df["es_mes_cruzado"]]["fecha"].tolist())

    inconsistencias = df[df["mes"] != df["mes_correcto"]].copy()
    if inconsistencias.empty:
        return pd.DataFrame()

    def accion_sugerida(row):
        if row["es_mes_cruzado"] and row["fecha"] in fechas_con_mes_real:
            return "Eliminar (duplicado cruzado)"
        if row["es_mes_cruzado"]:
            return "Corregir mes (fecha única)"
        return "Normalizar nombre de mes"

    inconsistencias["accion"] = inconsistencias.apply(accion_sugerida, axis=1)
    return inconsistencias[["fecha", "mes", "mes_correcto", "registros", "accion"]].reset_index(drop=True)


def limpiar_y_normalizar_meses_supabase():
    """Elimina duplicados de mes cruzado y normaliza las etiquetas de mes según la fecha real."""
    case_mes_num = """
        CASE LOWER(SPLIT_PART(COALESCE({col}, ''), ' ', 1))
            WHEN 'enero' THEN 1 WHEN 'febrero' THEN 2 WHEN 'marzo' THEN 3
            WHEN 'abril' THEN 4 WHEN 'mayo' THEN 5 WHEN 'junio' THEN 6
            WHEN 'julio' THEN 7 WHEN 'agosto' THEN 8 WHEN 'septiembre' THEN 9
            WHEN 'setiembre' THEN 9 WHEN 'octubre' THEN 10 WHEN 'noviembre' THEN 11
            WHEN 'diciembre' THEN 12 ELSE 0
        END
    """

    expr_mes_canonico = """
        CASE EXTRACT(MONTH FROM fecha)::int
            WHEN 1 THEN 'Enero' WHEN 2 THEN 'Febrero' WHEN 3 THEN 'Marzo'
            WHEN 4 THEN 'Abril' WHEN 5 THEN 'Mayo' WHEN 6 THEN 'Junio'
            WHEN 7 THEN 'Julio' WHEN 8 THEN 'Agosto' WHEN 9 THEN 'Septiembre'
            WHEN 10 THEN 'Octubre' WHEN 11 THEN 'Noviembre' WHEN 12 THEN 'Diciembre'
        END || ' ' || EXTRACT(YEAR FROM fecha)::int::text
    """

    q_delete_cruzados = text(f"""
        DELETE FROM public.ejecuciones_procesos e
        WHERE EXTRACT(MONTH FROM e.fecha)::int <> ({case_mes_num.format(col='e.mes')})
          AND EXISTS (
              SELECT 1
              FROM public.ejecuciones_procesos ok
              WHERE ok.fecha = e.fecha
                AND EXTRACT(MONTH FROM ok.fecha)::int = ({case_mes_num.format(col='ok.mes')})
          )
    """)

    q_update_canonico = text(f"""
        UPDATE public.ejecuciones_procesos
        SET mes = ({expr_mes_canonico})
        WHERE fecha IS NOT NULL
          AND mes IS DISTINCT FROM ({expr_mes_canonico})
    """)

    with engine.begin() as conn:
        res_del = conn.execute(q_delete_cruzados)
        res_upd = conn.execute(q_update_canonico)

    return (res_del.rowcount or 0), (res_upd.rowcount or 0)


# --- CONSULTAS CACHEADAS ---
@st.cache_data(ttl=600)
def cargar_filtros():
    with engine.connect() as conn:
        meses_raw = [
            r[0]
            for r in conn.execute(
                text("SELECT DISTINCT mes FROM public.ejecuciones_procesos WHERE mes IS NOT NULL")
            ).fetchall()
        ]
        grupos = [
            r[0]
            for r in conn.execute(
                text("SELECT DISTINCT grupo FROM public.ejecuciones_procesos WHERE grupo IS NOT NULL ORDER BY grupo")
            ).fetchall()
        ]
        horarios_raw = [
            r[0]
            for r in conn.execute(
                text("SELECT DISTINCT rango_horario FROM public.ejecuciones_procesos WHERE rango_horario IS NOT NULL")
            ).fetchall()
        ]
    meses = sorted(meses_raw, key=clave_orden_mes)
    horarios = sorted(horarios_raw, key=clave_orden_horario)
    return meses, grupos, horarios


@st.cache_data(ttl=600)
def cargar_resumen_fechas():
    q = text("""
        SELECT fecha,
               mes,
               COUNT(*) AS registros,
               SUM(tiempo_demora_seg) / 3600.0 AS horas_totales
        FROM public.ejecuciones_procesos
        GROUP BY fecha, mes
        ORDER BY fecha DESC
    """)
    with engine.connect() as conn:
        df = pd.read_sql(q, conn)
    if not df.empty:
        df["fecha"] = pd.to_datetime(df["fecha"]).dt.date
    return df


@st.cache_data(ttl=600)
def consultar_datos_agregados(meses_sel, grupos_sel, horarios_sel, fecha_inicio=None, fecha_fin=None):
    where_clauses = ["1=1"]
    params = {}

    if meses_sel:
        where_clauses.append("mes = ANY(:meses)")
        params["meses"] = list(meses_sel)
    if grupos_sel:
        where_clauses.append("grupo = ANY(:grupos)")
        params["grupos"] = list(grupos_sel)
    if horarios_sel:
        where_clauses.append("rango_horario = ANY(:horarios)")
        params["horarios"] = list(horarios_sel)
    if fecha_inicio is not None:
        where_clauses.append("fecha >= :fecha_inicio")
        params["fecha_inicio"] = fecha_inicio
    if fecha_fin is not None:
        where_clauses.append("fecha <= :fecha_fin")
        params["fecha_fin"] = fecha_fin

    where_sql = " AND ".join(where_clauses)

    q_diario = text(f"""
        SELECT fecha, mes,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales,
               COUNT(*) AS total_ejecuciones,
               AVG(tiempo_demora_seg)/60.0 AS promedio_min
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY fecha, mes
        ORDER BY fecha
    """)

    q_jobs = text(f"""
        SELECT jobname, grupo,
               SUM(tiempo_demora_seg)/3600.0 AS horas_acumuladas,
               COUNT(*) AS ejecuciones,
               SUM(veces_repeticion) AS repeticiones_totales,
               AVG(tiempo_demora_seg)/60.0 AS promedio_min,
               MAX(tiempo_demora_seg)/60.0 AS max_min,
               STDDEV(tiempo_demora_seg)/60.0 AS desviacion_min
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY jobname, grupo
        ORDER BY horas_acumuladas DESC
        LIMIT 25
    """)

    q_horario = text(f"""
        SELECT rango_horario,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales,
               COUNT(*) AS cantidad_procesos,
               AVG(tiempo_demora_seg)/60.0 AS promedio_min
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY rango_horario
    """)

    with engine.connect() as conn:
        df_diario = pd.read_sql(q_diario, conn, params=params)
        df_jobs = pd.read_sql(q_jobs, conn, params=params)
        df_horario = pd.read_sql(q_horario, conn, params=params)

    if not df_diario.empty:
        df_diario["fecha"] = pd.to_datetime(df_diario["fecha"]).dt.date

    if not df_horario.empty:
        df_horario = df_horario.sort_values(
            by="rango_horario", key=lambda col: col.map(clave_orden_horario)
        ).reset_index(drop=True)

    return df_diario, df_jobs, df_horario


@st.cache_data(ttl=600)
def consultar_detalle_procesos(
    meses_sel,
    grupos_sel,
    horarios_sel,
    fecha_inicio=None,
    fecha_fin=None,
    franja_especifica=None,
    busqueda_job="",
    limite=5000,
):
    where_clauses = ["1=1"]
    params = {}

    if meses_sel:
        where_clauses.append("mes = ANY(:meses)")
        params["meses"] = list(meses_sel)
    if grupos_sel:
        where_clauses.append("grupo = ANY(:grupos)")
        params["grupos"] = list(grupos_sel)
    if horarios_sel:
        where_clauses.append("rango_horario = ANY(:horarios)")
        params["horarios"] = list(horarios_sel)
    if fecha_inicio is not None:
        where_clauses.append("fecha >= :fecha_inicio")
        params["fecha_inicio"] = fecha_inicio
    if fecha_fin is not None:
        where_clauses.append("fecha <= :fecha_fin")
        params["fecha_fin"] = fecha_fin
    if franja_especifica and franja_especifica != "Todo el día":
        where_clauses.append("rango_horario = :franja_especifica")
        params["franja_especifica"] = franja_especifica
    if busqueda_job and busqueda_job.strip():
        where_clauses.append("jobname ILIKE :busqueda_job")
        params["busqueda_job"] = f"%{busqueda_job.strip()}%"

    where_sql = " AND ".join(where_clauses)

    q_count = text(f"SELECT COUNT(*) FROM public.ejecuciones_procesos WHERE {where_sql}")
    q_detalle = text(f"""
        SELECT fecha,
               rango_horario,
               jobname,
               grupo,
               hora_inicio,
               hora_fin,
               tiempo_demora_seg,
               veces_repeticion
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        ORDER BY fecha DESC, tiempo_demora_seg DESC, hora_inicio ASC
        LIMIT :limite
    """)

    params_detalle = {**params, "limite": limite}

    with engine.connect() as conn:
        total = conn.execute(q_count, params).scalar() or 0
        df_det = pd.read_sql(q_detalle, conn, params=params_detalle)

    return df_det, int(total)


@st.cache_data(ttl=600)
def consultar_comparativa_diaria(fechas_sel, grupos_sel, horarios_sel):
    where_clauses = ["fecha = ANY(:fechas)"]
    params = {"fechas": list(fechas_sel)}

    if grupos_sel:
        where_clauses.append("grupo = ANY(:grupos)")
        params["grupos"] = list(grupos_sel)
    if horarios_sel:
        where_clauses.append("rango_horario = ANY(:horarios)")
        params["horarios"] = list(horarios_sel)

    where_sql = " AND ".join(where_clauses)

    q_resumen = text(f"""
        SELECT fecha,
               COUNT(*) AS cantidad_procesos,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales,
               AVG(tiempo_demora_seg)/60.0 AS promedio_min
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY fecha
        ORDER BY fecha
    """)

    q_hora = text(f"""
        SELECT fecha,
               rango_horario,
               COUNT(*) AS cantidad_procesos,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY fecha, rango_horario
    """)

    q_grupo = text(f"""
        SELECT fecha,
               grupo,
               COUNT(*) AS cantidad_procesos,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY fecha, grupo
        ORDER BY grupo, fecha
    """)

    with engine.connect() as conn:
        df_res = pd.read_sql(q_resumen, conn, params=params)
        df_hora = pd.read_sql(q_hora, conn, params=params)
        df_grp = pd.read_sql(q_grupo, conn, params=params)

    for df_x in (df_res, df_hora, df_grp):
        if not df_x.empty:
            df_x["fecha"] = pd.to_datetime(df_x["fecha"]).dt.date
            df_x["fecha_str"] = df_x["fecha"].astype(str)

    if not df_hora.empty:
        df_hora = df_hora.sort_values(
            by=["rango_horario", "fecha"],
            key=lambda col: col.map(clave_orden_horario) if col.name == "rango_horario" else col,
        ).reset_index(drop=True)

    return df_res, df_hora, df_grp


@st.cache_data(ttl=600)
def consultar_comparativa_mensual(meses_sel, grupos_sel, horarios_sel):
    where_clauses = ["mes = ANY(:meses)"]
    params = {"meses": list(meses_sel)}

    if grupos_sel:
        where_clauses.append("grupo = ANY(:grupos)")
        params["grupos"] = list(grupos_sel)
    if horarios_sel:
        where_clauses.append("rango_horario = ANY(:horarios)")
        params["horarios"] = list(horarios_sel)

    where_sql = " AND ".join(where_clauses)

    q_resumen = text(f"""
        SELECT mes,
               COUNT(DISTINCT fecha) AS dias_cargados,
               COUNT(*) AS cantidad_procesos,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY mes
    """)

    q_dia_mes = text(f"""
        SELECT mes,
               EXTRACT(DAY FROM fecha)::int AS dia_mes,
               COUNT(*) AS cantidad_procesos,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY mes, EXTRACT(DAY FROM fecha)
        ORDER BY dia_mes
    """)

    q_hora = text(f"""
        SELECT mes,
               rango_horario,
               COUNT(*) AS cantidad_procesos,
               SUM(tiempo_demora_seg)/3600.0 AS horas_totales
        FROM public.ejecuciones_procesos
        WHERE {where_sql}
        GROUP BY mes, rango_horario
    """)

    with engine.connect() as conn:
        df_res = pd.read_sql(q_resumen, conn, params=params)
        df_dia = pd.read_sql(q_dia_mes, conn, params=params)
        df_hora = pd.read_sql(q_hora, conn, params=params)

    if not df_res.empty:
        df_res = df_res.sort_values(by="mes", key=lambda c: c.map(clave_orden_mes)).reset_index(drop=True)
        df_res["promedio_diario_hrs"] = df_res["horas_totales"] / df_res["dias_cargados"].clip(lower=1)

    if not df_hora.empty:
        df_hora = df_hora.sort_values(
            by="rango_horario", key=lambda c: c.map(clave_orden_horario)
        ).reset_index(drop=True)

    return df_res, df_dia, df_hora


@st.cache_data(ttl=600)
def consultar_perfil_cadencia_jobs(grupos_sel_tuple, feriados_tuple, jobname_unico=None):
    where_clauses = ["jobname IS NOT NULL", "fecha IS NOT NULL"]
    params = {}

    if grupos_sel_tuple:
        where_clauses.append("grupo = ANY(:grupos)")
        params["grupos"] = list(grupos_sel_tuple)
    if jobname_unico:
        where_clauses.append("jobname = :jobname_unico")
        params["jobname_unico"] = str(jobname_unico)

    where_sql = " AND ".join(where_clauses)

    if feriados_tuple:
        cond_feriado = "fecha = ANY(CAST(:feriados AS date[]))"
        params["feriados"] = list(feriados_tuple)
    else:
        cond_feriado = "FALSE"

    q = text(f"""
        WITH por_dia AS (
            SELECT fecha,
                   jobname,
                   STRING_AGG(DISTINCT grupo, ', ' ORDER BY grupo) AS grupo_dia,
                   SUM(tiempo_demora_seg) AS seg_dia,
                   COUNT(*) AS ejec_dia
            FROM public.ejecuciones_procesos
            WHERE {where_sql}
            GROUP BY fecha, jobname
        )
        SELECT jobname,
               STRING_AGG(DISTINCT grupo_dia, ', ' ORDER BY grupo_dia) AS grupo,
               MIN(fecha) AS primera_fecha,
               MAX(fecha) AS ultima_fecha,
               COUNT(*) AS dias_ejecutados,
               SUM(ejec_dia) AS total_ejecuciones,
               SUM(seg_dia) AS seg_totales,
               COUNT(DISTINCT TO_CHAR(fecha, 'YYYY-MM')) AS meses_con_ejecucion,
               COUNT(*) FILTER (
                   WHERE EXTRACT(ISODOW FROM fecha) BETWEEN 1 AND 5 AND NOT ({cond_feriado})
               ) AS ejec_habiles,
               COUNT(*) FILTER (WHERE EXTRACT(ISODOW FROM fecha) BETWEEN 1 AND 5) AS ejec_lun_vie,
               COUNT(*) FILTER (WHERE EXTRACT(ISODOW FROM fecha) = 6) AS ejec_sabados,
               COUNT(*) FILTER (WHERE EXTRACT(ISODOW FROM fecha) = 7) AS ejec_domingos,
               COUNT(*) FILTER (WHERE {cond_feriado}) AS ejec_feriados,
               COUNT(*) FILTER (
                   WHERE EXTRACT(DAY FROM fecha) IN (1, 30, 31)
                      OR fecha = (DATE_TRUNC('month', fecha) + INTERVAL '1 month - 1 day')::date
               ) AS ejec_cierre_inicio,
               ARRAY_AGG(DISTINCT TO_CHAR(fecha, 'YYYY-MM') ORDER BY TO_CHAR(fecha, 'YYYY-MM')) AS meses_activos_ym
        FROM por_dia
        GROUP BY jobname
        ORDER BY dias_ejecutados DESC, jobname
    """)

    with engine.connect() as conn:
        df = pd.read_sql(q, conn, params=params)

    if not df.empty:
        df["primera_fecha"] = pd.to_datetime(df["primera_fecha"]).dt.date
        df["ultima_fecha"] = pd.to_datetime(df["ultima_fecha"]).dt.date

    return df


def detectar_posibles_feriados(df_fechas_disp):
    if df_fechas_disp.empty:
        return []
    df_c = (
        df_fechas_disp.groupby("fecha", as_index=False)["registros"]
        .sum()
        .sort_values("fecha")
    )
    df_c["dow"] = pd.to_datetime(df_c["fecha"]).dt.dayofweek
    df_lv = df_c[df_c["dow"] < 5]
    if len(df_lv) < 5:
        return []
    mediana_lv = df_lv["registros"].median()
    umbral = mediana_lv * 0.75
    return sorted(df_lv[df_lv["registros"] < umbral]["fecha"].tolist())


def clasificar_cadencia_y_estabilidad(df_perfil, df_fechas_disp, feriados_lista, overrides_manuales=None):
    if df_perfil.empty or df_fechas_disp.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    overrides_manuales = overrides_manuales or {}
    feriados_set = set(feriados_lista)

    fechas_unicas = sorted(df_fechas_disp["fecha"].dropna().unique().tolist())
    df_cal = pd.DataFrame({"fecha": fechas_unicas})
    dt_s = pd.to_datetime(df_cal["fecha"])
    df_cal["ym"] = dt_s.dt.strftime("%Y-%m")
    df_cal["dow"] = dt_s.dt.dayofweek
    df_cal["es_lun_vie"] = df_cal["dow"] < 5
    df_cal["es_sab"] = df_cal["dow"] == 5
    df_cal["es_dom"] = df_cal["dow"] == 6
    df_cal["es_feriado"] = df_cal["fecha"].isin(feriados_set)
    df_cal["es_habil"] = df_cal["es_lun_vie"] & (~df_cal["es_feriado"])
    df_cal["es_cierre_inicio"] = dt_s.dt.day.isin([1, 30, 31]) | dt_s.dt.is_month_end

    meses_cal = sorted(df_cal["ym"].unique().tolist())
    primer_ym = meses_cal[0]
    meses_con_cierre = sorted(df_cal[df_cal["es_cierre_inicio"]]["ym"].unique().tolist())
    ultimo_ym_cierre = meses_con_cierre[-1] if meses_con_cierre else meses_cal[-1]
    penultimo_ym_cierre = meses_con_cierre[-2] if len(meses_con_cierre) >= 2 else ultimo_ym_cierre

    idx_fecha = {f: i for i, f in enumerate(fechas_unicas)}
    idx_ym = {m: i for i, m in enumerate(meses_cal)}
    n_cal = len(fechas_unicas)

    def prefijo(arr):
        ps = [0] * (len(arr) + 1)
        for i, val in enumerate(arr):
            ps[i + 1] = ps[i] + int(val)
        return ps

    ps_hab = prefijo(df_cal["es_habil"].tolist())
    ps_lv = prefijo(df_cal["es_lun_vie"].tolist())
    ps_sab = prefijo(df_cal["es_sab"].tolist())
    ps_dom = prefijo(df_cal["es_dom"].tolist())
    ps_fer = prefijo(df_cal["es_feriado"].tolist())

    resultados = []
    for row in df_perfil.itertuples(index=False):
        i_ini = idx_fecha.get(row.primera_fecha, 0)
        i_fin = idx_fecha.get(row.ultima_fecha, n_cal - 1)

        cal_total = max(1, i_fin - i_ini + 1)
        cal_hab = ps_hab[i_fin + 1] - ps_hab[i_ini]
        cal_lv = ps_lv[i_fin + 1] - ps_lv[i_ini]
        cal_sab = ps_sab[i_fin + 1] - ps_sab[i_ini]
        cal_dom = ps_dom[i_fin + 1] - ps_dom[i_ini]
        cal_fer = ps_fer[i_fin + 1] - ps_fer[i_ini]
        cal_antes = i_ini
        cal_despues = (n_cal - 1) - i_fin

        dias_ej = int(row.dias_ejecutados)
        meses_ej = max(1, int(row.meses_con_ejecucion))
        prom_dias_mes = dias_ej / meses_ej
        tasa_total = dias_ej / cal_total
        tasa_hab = (int(row.ejec_habiles) / cal_hab) if cal_hab > 0 else 0.0
        tasa_lv = (int(row.ejec_lun_vie) / cal_lv) if cal_lv > 0 else 0.0
        tasa_sab = (int(row.ejec_sabados) / cal_sab) if cal_sab > 0 else None
        tasa_dom = (int(row.ejec_domingos) / cal_dom) if cal_dom > 0 else None
        tasa_fer = (int(row.ejec_feriados) / cal_fer) if cal_fer > 0 else None
        ratio_cierre = int(row.ejec_cierre_inicio) / max(1, dias_ej)

        raw_meses = row.meses_activos_ym
        if isinstance(raw_meses, (list, tuple)):
            meses_act = list(raw_meses)
        else:
            meses_act = [m.strip() for m in str(raw_meses).strip("{}").split(",") if m.strip()]

        ym_inicio = meses_act[0] if meses_act else row.primera_fecha.strftime("%Y-%m")
        ym_fin = meses_act[-1] if meses_act else row.ultima_fecha.strftime("%Y-%m")

        # 1. Clasificación automática
        es_mensual = (
            (meses_ej >= 2 and prom_dias_mes <= 2.5 and tasa_hab < 0.35)
            or (cal_total >= 15 and dias_ej <= 2 and ratio_cierre >= 0.5)
        )
        es_diario = (
            not es_mensual
            and (dias_ej >= 3 or (cal_total <= 4 and tasa_total >= 0.75))
            and (tasa_total >= 0.60 or tasa_hab >= 0.65 or tasa_lv >= 0.65)
        )

        if es_mensual:
            categoria_auto = "Mensual"
            if ratio_cierre >= 0.60:
                patron_auto = "Mensual (Cierre/Inicio: días 30, 31, fin de mes o 1)"
            else:
                patron_auto = "Mensual (Otro día del mes)"
            excepciones_txt = "Solo cierre/inicio de mes" if ratio_cierre >= 0.60 else "Otro día fijo del mes"
        elif es_diario:
            categoria_auto = "Diario"
            omite_sab = (cal_sab >= 2) and (tasa_sab is not None and tasa_sab < 0.25)
            omite_dom = (cal_dom >= 2) and (tasa_dom is not None and tasa_dom < 0.25)
            omite_fer = (cal_fer >= 1) and (tasa_fer is not None and tasa_fer < 0.40)

            excepciones = []
            if omite_sab and omite_dom:
                excepciones.append("No corre sábados ni domingos")
            elif omite_sab:
                excepciones.append("No corre sábados")
            elif omite_dom:
                excepciones.append("No corre domingos")
            if omite_fer:
                excepciones.append("No corre feriados")

            if not excepciones:
                patron_auto = "Diario (Todos los días, sin excepciones)"
                excepciones_txt = "Ninguna (corre todos los días)"
            else:
                patron_auto = f"Diario ({', '.join(excepciones)})"
                excepciones_txt = ", ".join(excepciones)
        else:
            categoria_auto = "Semanal / Otros"
            if dias_ej >= 3 and 2.5 < prom_dias_mes <= 8:
                patron_auto = "Semanal / Quincenal"
                excepciones_txt = "Frecuencia intermedia"
            else:
                patron_auto = "Esporádico / A demanda"
                excepciones_txt = "Ejecución eventual"

        # 2. Aplicar clasificación manual si existe para este JOBNAME
        ov = overrides_manuales.get(str(row.jobname))
        if ov:
            categoria = ov.get("categoria_manual") or categoria_auto
            patron = ov.get("patron_manual") or PATRONES_POR_CATEGORIA.get(categoria, [patron_auto])[0]
            origen_clas = "Manual"
            if categoria == "Diario":
                if "sin excepciones" in patron.lower():
                    excepciones_txt = "Ninguna (corre todos los días)"
                else:
                    excepciones_txt = patron
            elif categoria == "Mensual":
                excepciones_txt = "Solo cierre/inicio de mes" if "Cierre/Inicio" in patron else patron
            else:
                excepciones_txt = patron
        else:
            categoria = categoria_auto
            patron = patron_auto
            origen_clas = "Automática"

        i_ym_ini = idx_ym.get(ym_inicio, 0)
        i_ym_fin = idx_ym.get(ym_fin, len(meses_cal) - 1)
        meses_en_ventana = meses_cal[i_ym_ini : i_ym_fin + 1]
        meses_faltantes_list = [etiqueta_ym(m) for m in meses_en_ventana if m not in meses_act]
        meses_faltantes_txt = ", ".join(meses_faltantes_list) if meses_faltantes_list else "Ninguno"

        es_observacion = False
        if categoria == "Diario":
            es_alta = cal_antes >= 7
            es_baja = cal_despues >= 7
        elif categoria == "Mensual":
            es_alta = ym_inicio > primer_ym
            if len(meses_con_cierre) >= 3:
                es_baja = ym_fin < penultimo_ym_cierre
                es_observacion = (not es_baja) and (ym_fin < ultimo_ym_cierre)
            else:
                es_baja = ym_fin < ultimo_ym_cierre
        else:
            es_alta = False
            es_baja = False

        if es_alta and es_baja:
            estado = "Temporal (Alta y Baja)"
        elif es_alta:
            estado = "Agregado (Alta)"
        elif es_baja:
            estado = "Retirado (Baja)"
        elif es_observacion:
            estado = "En observación (sin corrida último mes)"
        elif meses_faltantes_list:
            estado = "Estable (con meses sin corrida)"
        else:
            estado = "Estable (todos los meses)"

        resultados.append({
            "jobname": row.jobname,
            "grupo": row.grupo,
            "categoria": categoria,
            "categoria_auto": categoria_auto,
            "patron": patron,
            "patron_auto": patron_auto,
            "origen_clas": origen_clas,
            "excepciones": excepciones_txt,
            "estado": estado,
            "es_alta": es_alta,
            "es_baja": es_baja,
            "primera_fecha": row.primera_fecha,
            "ultima_fecha": row.ultima_fecha,
            "ym_inicio": ym_inicio,
            "ym_fin": ym_fin,
            "dias_ejecutados": dias_ej,
            "meses_con_ejecucion": meses_ej,
            "meses_faltantes": meses_faltantes_txt,
            "total_ejecuciones": int(row.total_ejecuciones),
            "tiempo_acum_hhmmss": formato_seg_hhmmss(row.seg_totales),
            "prom_dia_hhmmss": formato_seg_hhmmss(float(row.seg_totales) / max(1, dias_ej)),
            "meses_activos_ym": meses_act,
        })

    df_clas = pd.DataFrame(resultados)

    filas_estabilidad = []
    for idx_m, ym in enumerate(meses_cal):
        ym_prev = meses_cal[idx_m - 1] if idx_m > 0 else None

        df_d = df_clas[df_clas["categoria"] == "Diario"]
        df_m = df_clas[df_clas["categoria"] == "Mensual"]

        diarios_ejec = int(df_d["meses_activos_ym"].apply(lambda lst: ym in lst).sum())
        diarios_vig = int(((df_d["ym_inicio"] <= ym) & ((~df_d["es_baja"]) | (df_d["ym_fin"] >= ym))).sum())

        mensuales_ejec = int(df_m["meses_activos_ym"].apply(lambda lst: ym in lst).sum())
        mensuales_vig = int(((df_m["ym_inicio"] <= ym) & ((~df_m["es_baja"]) | (df_m["ym_fin"] >= ym))).sum())
        mensuales_sin_corrida = max(0, mensuales_vig - mensuales_ejec)

        altas_d = int((df_d["es_alta"] & (df_d["ym_inicio"] == ym)).sum())
        bajas_d = int((df_d["es_baja"] & (df_d["ym_fin"] == ym_prev)).sum()) if ym_prev else 0
        altas_m = int((df_m["es_alta"] & (df_m["ym_inicio"] == ym)).sum())
        bajas_m = int((df_m["es_baja"] & (df_m["ym_fin"] == ym_prev)).sum()) if ym_prev else 0

        filas_estabilidad.append({
            "ym": ym,
            "Mes": etiqueta_ym(ym),
            "Diarios Vigentes": diarios_vig,
            "Diarios Ejecutados": diarios_ejec,
            "Altas Diarias (+)": altas_d,
            "Bajas Diarias (-)": bajas_d,
            "Mensuales Vigentes": mensuales_vig,
            "Mensuales Ejecutados": mensuales_ejec,
            "Mensuales Sin Corrida": mensuales_sin_corrida,
            "Altas Mensuales (+)": altas_m,
            "Bajas Mensuales (-)": bajas_m,
        })

    df_estab = pd.DataFrame(filas_estabilidad)

    df_cambios = df_clas[
        df_clas["categoria"].isin(["Diario", "Mensual"])
        & (df_clas["es_alta"] | df_clas["es_baja"])
    ].copy()
    if not df_cambios.empty:
        df_cambios = df_cambios.sort_values(["categoria", "primera_fecha", "jobname"])

    return df_clas, df_estab, df_cambios


# --- COMPONENTES VISUALES ---
def graficar_barras_horario(df_horario, metrica):
    orden = df_horario["rango_horario"].tolist()
    df_plot = df_horario.copy()
    df_plot["tiempo_hhmmss"] = df_plot["horas_totales"].apply(formato_hhmmss)
    df_plot["proc_str"] = df_plot["cantidad_procesos"].astype(int).astype(str)
    df_plot["prom_str"] = df_plot["promedio_min"].map(lambda x: f"{x:.2f}")

    if metrica == "Cantidad de Procesos":
        y_col = "cantidad_procesos"
        y_label = "Procesos Ejecutados"
        text_col = "proc_str"
    else:
        y_col = "horas_totales"
        y_label = "Tiempo Acumulado (hh:mm:ss)"
        text_col = "tiempo_hhmmss"

    fig = px.bar(
        df_plot,
        x="rango_horario",
        y=y_col,
        text=text_col,
        custom_data=["proc_str", "tiempo_hhmmss", "prom_str"],
        category_orders={"rango_horario": orden},
        labels={"rango_horario": "Rango Horario"},
    )
    fig.update_traces(
        textposition="outside",
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Procesos Ejecutados: %{customdata[0]}<br>"
            "Tiempo Acumulado: %{customdata[1]}<br>"
            "Promedio (min/proc): %{customdata[2]}<extra></extra>"
        ),
    )
    fig.update_layout(
        xaxis_type="category",
        yaxis_title=y_label,
        xaxis_tickangle=-45,
        clickmode="event+select",
    )
    if metrica == "Cantidad de Procesos":
        aplicar_eje_entero(fig, eje="y")
    else:
        aplicar_eje_tiempo_hhmmss(fig, df_plot["horas_totales"].max(), eje="y")
    return fig


def preparar_tabla_horario(df_horario):
    df_t = df_horario.copy()
    df_t["tiempo_hhmmss"] = df_t["horas_totales"].apply(formato_hhmmss)
    df_t = df_t[["rango_horario", "cantidad_procesos", "tiempo_hhmmss", "promedio_min"]].rename(
        columns={
            "rango_horario": "Rango Horario",
            "cantidad_procesos": "Procesos Ejecutados",
            "tiempo_hhmmss": "Tiempo Acumulado (hh:mm:ss)",
            "promedio_min": "Promedio (min/proc)",
        }
    )
    return df_t.style.format({
        "Procesos Ejecutados": "{:.0f}",
        "Promedio (min/proc)": "{:.2f}",
    })


def extraer_franjas_seleccionadas(evento, franjas_validas):
    if not evento:
        return []
    selection = evento.get("selection") if isinstance(evento, dict) else getattr(evento, "selection", None)
    if not selection:
        return []
    points = selection.get("points", []) if isinstance(selection, dict) else getattr(selection, "points", [])
    franjas = []
    for p in points:
        x_val = p.get("x") if isinstance(p, dict) else getattr(p, "x", None)
        if x_val in franjas_validas and x_val not in franjas:
            franjas.append(x_val)
    return franjas


def render_vista_horaria_interactiva(
    df_horario,
    prefijo,
    titulo_grafico,
    titulo_resumen,
    meses_sel,
    grupos_sel,
    horarios_sel,
    fecha_inicio,
    fecha_fin,
):
    st.subheader(titulo_grafico)
    metrica = st.radio(
        "Mostrar en barras:",
        options=["Cantidad de Procesos", "Tiempo Acumulado (hh:mm:ss)"],
        horizontal=True,
        key=f"metrica_{prefijo}",
    )

    key_sel = f"selector_franja_{prefijo}"
    key_clic = f"ultimo_clic_{prefijo}"
    key_ver = f"chart_ver_{prefijo}"
    key_busq = f"busq_job_{prefijo}"

    ver_actual = st.session_state.get(key_ver, 0)
    fig = graficar_barras_horario(df_horario, metrica)
    evento = st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"chart_{prefijo}_{ver_actual}",
        on_select="rerun",
        selection_mode="points",
    )
    st.caption(
        "💡 Haz clic en cualquier barra del gráfico para ver abajo los procesos de esa franja. "
        "'Tiempo Acumulado (hh:mm:ss)' suma la duración individual de los procesos ejecutados en paralelo."
    )

    franjas_disp_vista = df_horario["rango_horario"].tolist()
    opciones_franja = ["Todo el día"] + franjas_disp_vista

    franjas_clic = extraer_franjas_seleccionadas(evento, franjas_disp_vista)
    if franjas_clic != st.session_state.get(key_clic, []):
        st.session_state[key_clic] = franjas_clic
        st.session_state[key_sel] = franjas_clic[-1] if franjas_clic else "Todo el día"

    if st.session_state.get(key_sel) not in opciones_franja:
        st.session_state[key_sel] = "Todo el día"

    if titulo_resumen:
        st.subheader(titulo_resumen)
    st.dataframe(preparar_tabla_horario(df_horario), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("📋 Lista de Procesos Ejecutados")

    def seleccionar_todo_el_dia():
        st.session_state[key_sel] = "Todo el día"
        st.session_state[key_clic] = []
        st.session_state[key_ver] = st.session_state.get(key_ver, 0) + 1

    col_sel, col_btn, col_busq = st.columns([2, 1, 2], vertical_alignment="bottom")
    with col_sel:
        franja_elegida = st.selectbox(
            "Filtrar lista por franja horaria:",
            options=opciones_franja,
            key=key_sel,
        )
    with col_btn:
        st.button(
            "📋 Todo el día",
            key=f"btn_todo_{prefijo}",
            on_click=seleccionar_todo_el_dia,
            use_container_width=True,
            help="Muestra todos los procesos del periodo filtrado y limpia la barra seleccionada.",
        )
    with col_busq:
        busqueda_job = st.text_input(
            "Buscar por JOBNAME (opcional):",
            placeholder="Ej. SIUCD021",
            key=key_busq,
        )

    df_det, total_coincidencias = consultar_detalle_procesos(
        meses_sel=meses_sel,
        grupos_sel=grupos_sel,
        horarios_sel=horarios_sel,
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
        franja_especifica=franja_elegida,
        busqueda_job=busqueda_job,
        limite=5000,
    )

    etiqueta_franja = "todo el día" if franja_elegida == "Todo el día" else f"la franja **{franja_elegida}**"
    if total_coincidencias == 0:
        st.info(f"No se encontraron procesos para {etiqueta_franja} con los filtros actuales.")
    else:
        if total_coincidencias > len(df_det):
            st.caption(
                f"Mostrando los **{len(df_det)}** procesos más largos de **{total_coincidencias}** "
                f"encontrados en {etiqueta_franja}. Filtra por fecha, franja o JOBNAME para acotar."
            )
        else:
            st.caption(f"Mostrando **{total_coincidencias}** proceso(s) en {etiqueta_franja} (ordenados por mayor duración).")

        df_p = df_det.copy()
        df_p["Duración (hh:mm:ss)"] = df_p["tiempo_demora_seg"].apply(formato_seg_hhmmss)
        df_p["Duración (min)"] = df_p["tiempo_demora_seg"] / 60.0
        df_p = df_p[
            [
                "fecha",
                "rango_horario",
                "jobname",
                "grupo",
                "hora_inicio",
                "hora_fin",
                "Duración (hh:mm:ss)",
                "Duración (min)",
                "veces_repeticion",
            ]
        ].rename(
            columns={
                "fecha": "Fecha",
                "rango_horario": "Rango Horario",
                "jobname": "Proceso (JOBNAME)",
                "grupo": "Grupo",
                "hora_inicio": "Hora Inicio",
                "hora_fin": "Hora Fin",
                "veces_repeticion": "Repeticiones",
            }
        )

        st.dataframe(
            df_p.style.format({
                "Duración (min)": "{:.2f}",
                "Repeticiones": "{:.0f}",
            }),
            use_container_width=True,
            hide_index=True,
        )


def resetear_filtros():
    st.session_state["filtro_mes_directo"] = "Todos los meses"
    st.session_state["toggle_todos_meses"] = True
    st.session_state["filtro_meses"] = []
    st.session_state["toggle_rango_fechas"] = False
    st.session_state["filtro_grupos"] = []
    st.session_state["filtro_horarios"] = []
    for p in ("tab1", "tab3"):
        st.session_state[f"selector_franja_{p}"] = "Todo el día"
        st.session_state[f"ultimo_clic_{p}"] = []
        st.session_state[f"busq_job_{p}"] = ""
        st.session_state[f"chart_ver_{p}"] = st.session_state.get(f"chart_ver_{p}", 0) + 1


def refrescar_app(mensaje_exito):
    """Limpia caché y programa el reset de filtros para el inicio del próximo rerun (evita conflictos de widgets)."""
    st.cache_data.clear()
    st.session_state["_pendiente_reset_filtros"] = True
    st.session_state["notificacion_exito"] = mensaje_exito
    st.rerun()


def _on_enviar_pregunta_ia():
    texto = st.session_state.get("ia_input_pregunta", "").strip()
    if texto:
        st.session_state["ia_pregunta_pendiente"] = texto
        st.session_state["ia_input_pregunta"] = ""


# --- INTERFAZ DE USUARIO ---
st.title("📊 Monitor de Rendimiento y Cuellos de Botella de Procesos")
st.markdown("Análisis histórico diario, mensual, comparativas y auditoría de cadencia operativa.")

if "notificacion_exito" in st.session_state:
    st.success(st.session_state.pop("notificacion_exito"))

try:
    meses_disp, grupos_disp, horarios_disp = cargar_filtros()
    df_fechas_disp = cargar_resumen_fechas()
    df_inconsistencias = detectar_inconsistencias_meses()
    overrides_manuales = cargar_clasificaciones_manuales()
except Exception as e:
    st.error(f"Error conectando a Supabase. Verifica tu archivo .env: {e}")
    st.stop()

# Limpiar valores en sesión que ya no existan tras un borrado o limpieza (antes de instanciar widgets)
if (
    "filtro_mes_directo" in st.session_state
    and st.session_state["filtro_mes_directo"] not in (["Todos los meses"] + meses_disp)
):
    st.session_state["filtro_mes_directo"] = "Todos los meses"
if "filtro_meses" in st.session_state:
    st.session_state["filtro_meses"] = [m for m in st.session_state["filtro_meses"] if m in meses_disp]
if "filtro_grupos" in st.session_state:
    st.session_state["filtro_grupos"] = [g for g in st.session_state["filtro_grupos"] if g in grupos_disp]
if "filtro_horarios" in st.session_state:
    st.session_state["filtro_horarios"] = [h for h in st.session_state["filtro_horarios"] if h in horarios_disp]

if not df_inconsistencias.empty:
    n_cruzados = int((df_inconsistencias["accion"] == "Eliminar (duplicado cruzado)").sum())
    if n_cruzados > 0:
        st.warning(
            f"Se detectaron **{n_cruzados}** fecha(s) duplicadas por mes cruzado en Supabase "
            "(ej. día 01 cargado bajo el mes anterior). Puedes limpiarlas en un clic desde **⚙️ Gestión de Datos**."
        )

# --- SIDEBAR DE FILTROS GLOBALES ---
st.sidebar.header("🔍 Filtros de Consulta")

with st.sidebar.container(border=True):
    st.markdown("**Periodo**")

    mes_directo = st.selectbox(
        "Mes directo:",
        options=["Todos los meses"] + meses_disp,
        key="filtro_mes_directo",
    )

    if mes_directo != "Todos los meses":
        meses_sel = [mes_directo]
        todos_los_meses = False
    else:
        todos_los_meses = st.toggle("Todos los meses", value=True, key="toggle_todos_meses")
        if todos_los_meses:
            meses_sel = meses_disp
        else:
            meses_sel = st.multiselect(
                "Seleccionar Mes(es):",
                options=meses_disp,
                placeholder="Elige uno o más meses...",
                key="filtro_meses",
            )

    fecha_inicio = None
    fecha_fin = None
    if not df_fechas_disp.empty and meses_sel:
        fechas_periodo = df_fechas_disp[df_fechas_disp["mes"].isin(meses_sel)]["fecha"]
        if not fechas_periodo.empty:
            min_f = fechas_periodo.min()
            max_f = fechas_periodo.max()
            usar_rango = st.toggle("Filtrar por rango de fechas", value=False, key="toggle_rango_fechas")
            if usar_rango:
                rango_sel = st.date_input(
                    "Rango de fechas:",
                    value=(min_f, max_f),
                    min_value=min_f,
                    max_value=max_f,
                )
                if isinstance(rango_sel, tuple) and len(rango_sel) == 2:
                    fecha_inicio, fecha_fin = rango_sel
                elif isinstance(rango_sel, tuple) and len(rango_sel) == 1:
                    fecha_inicio = fecha_fin = rango_sel[0]
                else:
                    fecha_inicio = fecha_fin = rango_sel

with st.sidebar.container(border=True):
    st.markdown("**Segmentación**")
    grupos_sel = st.multiselect(
        "Grupo:",
        options=grupos_disp,
        placeholder="Todos los grupos",
        key="filtro_grupos",
    )
    horarios_sel = st.multiselect(
        "Franja horaria:",
        options=horarios_disp,
        placeholder="Todas las franjas",
        key="filtro_horarios",
    )

st.sidebar.button("🔄 Restablecer filtros", on_click=resetear_filtros, use_container_width=True)

if mes_directo != "Todos los meses":
    resumen_meses = mes_directo
elif todos_los_meses:
    resumen_meses = "Todos los meses"
else:
    resumen_meses = f"{len(meses_sel)} mes(es)"
resumen_grupos = "Todos" if not grupos_sel else f"{len(grupos_sel)} seleccionado(s)"
resumen_horas = "Todas" if not horarios_sel else f"{len(horarios_sel)} seleccionada(s)"
st.sidebar.caption(f"**Activo:** {resumen_meses} · Grupos: {resumen_grupos} · Franjas: {resumen_horas}")

# Consultar datos aplicando TODOS los filtros globales
if meses_disp and meses_sel:
    df_diario, df_jobs, df_horario = consultar_datos_agregados(
        meses_sel=meses_sel,
        grupos_sel=grupos_sel,
        horarios_sel=horarios_sel,
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )
else:
    df_diario = pd.DataFrame()
    df_jobs = pd.DataFrame()
    df_horario = pd.DataFrame()

if not df_diario.empty:
    df_por_fecha = (
        df_diario.groupby("fecha", as_index=False)
        .agg(
            horas_totales=("horas_totales", "sum"),
            total_ejecuciones=("total_ejecuciones", "sum"),
        )
        .sort_values("fecha")
    )
else:
    df_por_fecha = pd.DataFrame()

es_un_solo_dia = len(df_por_fecha) == 1

# --- KPIs PRINCIPALES ---
if not df_por_fecha.empty:
    total_horas = df_por_fecha["horas_totales"].sum()
    total_proc = int(df_por_fecha["total_ejecuciones"].sum())
    promedio_diario = df_por_fecha["horas_totales"].mean()
    total_dias = len(df_por_fecha)
    dia_mas_pesado = df_por_fecha.loc[df_por_fecha["horas_totales"].idxmax()]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Tiempo Total Acumulado", formato_hhmmss(total_horas))
    col2.metric("Total Ejecuciones", f"{total_proc}")

    if es_un_solo_dia and not df_horario.empty:
        franja_pico = df_horario.loc[df_horario["cantidad_procesos"].idxmax()]
        col3.metric(
            "Franja con Más Procesos",
            f"{franja_pico['rango_horario']}",
            f"{int(franja_pico['cantidad_procesos'])} procesos",
            delta_color="off",
        )
        franja_pesada = df_horario.loc[df_horario["horas_totales"].idxmax()]
        col4.metric(
            "Franja Más Pesada",
            f"{franja_pesada['rango_horario']}",
            formato_hhmmss(franja_pesada["horas_totales"]),
            delta_color="off",
        )
    else:
        col3.metric(
            "Promedio Diario",
            formato_hhmmss(promedio_diario),
            f"{total_dias} días",
            delta_color="off",
        )
        col4.metric(
            "Día Más Pesado",
            f"{dia_mas_pesado['fecha']}",
            formato_hhmmss(dia_mas_pesado["horas_totales"]),
            delta_color="off",
        )
else:
    st.info("No hay datos para los filtros seleccionados. Ajusta los filtros en la barra lateral o carga nuevos días en **⚙️ Gestión de Datos**.")

# --- CUADRO DE CONSULTA IA (SIN BOTÓN: SOLO ESCRIBIR Y PRESIONAR ENTER) ---
with st.container(border=True):
    st.text_input(
        "🤖 ¿Tiene alguna pregunta?",
        placeholder="Escribe tu pregunta sobre los datos o diferencias y presiona Enter (ej. ¿Por qué febrero tuvo menos mensuales que enero?)",
        key="ia_input_pregunta",
        on_change=_on_enviar_pregunta_ia,
    )

    pregunta_pendiente = st.session_state.pop("ia_pregunta_pendiente", None)
    if pregunta_pendiente:
        contexto_ia = {
            "meses_seleccionados": mes_directo if mes_directo != "Todos los meses" else ("Todos" if todos_los_meses else meses_sel),
            "rango_fechas_activo": f"{fecha_inicio} a {fecha_fin}" if fecha_inicio else "Completo",
            "grupos_filtrados": grupos_sel if grupos_sel else "Todos",
            "franjas_filtradas": horarios_sel if horarios_sel else "Todas",
            "dias_visibles": int(len(df_por_fecha)),
            "tiempo_total_visible": formato_hhmmss(df_por_fecha["horas_totales"].sum()) if not df_por_fecha.empty else "00:00:00",
            "ejecuciones_visibles": int(df_por_fecha["total_ejecuciones"].sum()) if not df_por_fecha.empty else 0,
        }
        with st.spinner("Consultando Supabase y analizando con Gemini..."):
            resp_ia = responder_pregunta_ia(pregunta_pendiente, contexto_ia)
        st.session_state["ia_ultima_consulta"] = {
            "pregunta": pregunta_pendiente,
            "respuesta": resp_ia,
        }

    if "ia_ultima_consulta" in st.session_state:
        ult_ia = st.session_state["ia_ultima_consulta"]
        with st.expander(f"💬 Respuesta a: {ult_ia['pregunta']}", expanded=True):
            st.markdown(ult_ia["respuesta"])

# --- PESTAÑAS DE ANÁLISIS, COMPARATIVA, AUDITORÍA Y GESTIÓN ---
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📅 Evolución Diaria y Mensual",
    "🔥 Top Procesos Críticos",
    "⏰ Concurrencia Horaria",
    "📈 Impacto y Variabilidad",
    "⚖️ Comparativa Diaria y Mensual",
    "🧭 Análisis Diarios vs Mensuales",
    "⚙️ Gestión de Datos",
])

# === TAB 1: EVOLUCIÓN DIARIA Y MENSUAL ===
with tab1:
    if df_diario.empty:
        st.warning("Sin datos disponibles para los filtros actuales.")
    elif es_un_solo_dia:
        fecha_sel = df_por_fecha["fecha"].iloc[0]
        render_vista_horaria_interactiva(
            df_horario=df_horario,
            prefijo="tab1",
            titulo_grafico=f"Actividad por Rango Horario — {fecha_sel}",
            titulo_resumen=f"Detalle Horario del {fecha_sel}",
            meses_sel=meses_sel,
            grupos_sel=grupos_sel,
            horarios_sel=horarios_sel,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
        )
    else:
        st.subheader("Carga Operativa por Día (Tiempo Acumulado hh:mm:ss)")
        df_d = df_diario.copy()
        df_d["tiempo_hhmmss"] = df_d["horas_totales"].apply(formato_hhmmss)
        df_d["proc_str"] = df_d["total_ejecuciones"].astype(int).astype(str)
        df_d["prom_str"] = df_d["promedio_min"].map(lambda x: f"{x:.2f}")

        fig_dia = px.bar(
            df_d,
            x="fecha",
            y="horas_totales",
            color="mes",
            custom_data=["mes", "tiempo_hhmmss", "proc_str", "prom_str"],
            category_orders={"mes": meses_disp},
            labels={"fecha": "Fecha", "horas_totales": "Tiempo Acumulado (hh:mm:ss)", "mes": "Mes"},
        )
        fig_dia.update_traces(
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Mes: %{customdata[0]}<br>"
                "Tiempo Acumulado: %{customdata[1]}<br>"
                "Procesos Ejecutados: %{customdata[2]}<br>"
                "Promedio (min/proc): %{customdata[3]}<extra></extra>"
            )
        )
        aplicar_eje_tiempo_hhmmss(fig_dia, df_por_fecha["horas_totales"].max(), eje="y")
        st.plotly_chart(fig_dia, use_container_width=True, key="chart_tab1_diario")

        st.subheader("Comparativa Mensual")
        df_mensual = df_diario.groupby("mes", as_index=False).agg(
            Dias=("fecha", "nunique"),
            horas_sum=("horas_totales", "sum"),
            Ejecuciones=("total_ejecuciones", "sum"),
            prom_dia_hrs=("horas_totales", "mean"),
        )
        df_mensual = df_mensual.sort_values(by="mes", key=lambda col: col.map(clave_orden_mes))
        df_mensual["Tiempo Acumulado (hh:mm:ss)"] = df_mensual["horas_sum"].apply(formato_hhmmss)
        df_mensual["Promedio Diario (hh:mm:ss)"] = df_mensual["prom_dia_hrs"].apply(formato_hhmmss)

        df_mensual_vista = df_mensual[
            ["mes", "Dias", "Ejecuciones", "Tiempo Acumulado (hh:mm:ss)", "Promedio Diario (hh:mm:ss)"]
        ].rename(columns={"mes": "Mes", "Dias": "Días"})

        st.dataframe(
            df_mensual_vista.style.format({"Días": "{:.0f}", "Ejecuciones": "{:.0f}"}),
            use_container_width=True,
            hide_index=True,
        )

# === TAB 2: TOP PROCESOS CRÍTICOS ===
with tab2:
    if df_jobs.empty:
        st.warning("Sin datos disponibles para los filtros actuales.")
    else:
        st.subheader("Top 25 Jobs con Mayor Tiempo Acumulado")
        df_jobs_plot = df_jobs.copy()
        df_jobs_plot["tiempo_hhmmss"] = df_jobs_plot["horas_acumuladas"].apply(formato_hhmmss)
        df_jobs_plot["ejec_str"] = df_jobs_plot["ejecuciones"].astype(int).astype(str)
        df_jobs_plot["prom_str"] = df_jobs_plot["promedio_min"].map(lambda x: f"{x:.2f}")
        df_jobs_plot["max_str"] = df_jobs_plot["max_min"].map(lambda x: f"{x:.2f}")

        df_top15 = df_jobs_plot.head(15)
        fig_jobs = px.bar(
            df_top15,
            x="horas_acumuladas",
            y="jobname",
            orientation="h",
            color="grupo",
            text="tiempo_hhmmss",
            custom_data=["grupo", "tiempo_hhmmss", "ejec_str", "prom_str", "max_str"],
            labels={"horas_acumuladas": "Tiempo Acumulado (hh:mm:ss)", "jobname": "Proceso (JOBNAME)", "grupo": "Grupo"},
        )
        fig_jobs.update_traces(
            textposition="outside",
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Grupo: %{customdata[0]}<br>"
                "Tiempo Acumulado: %{customdata[1]}<br>"
                "Ejecuciones: %{customdata[2]}<br>"
                "Promedio (min): %{customdata[3]}<br>"
                "Máximo (min): %{customdata[4]}<extra></extra>"
            ),
        )
        fig_jobs.update_layout(yaxis={"categoryorder": "total ascending"})
        aplicar_eje_tiempo_hhmmss(fig_jobs, df_top15["horas_acumuladas"].max(), eje="x")
        st.plotly_chart(fig_jobs, use_container_width=True, key="chart_tab2_jobs")

        df_jobs_tabla = df_jobs_plot[
            [
                "jobname",
                "grupo",
                "tiempo_hhmmss",
                "ejecuciones",
                "repeticiones_totales",
                "promedio_min",
                "max_min",
                "desviacion_min",
            ]
        ].rename(
            columns={
                "jobname": "Proceso (JOBNAME)",
                "grupo": "Grupo",
                "tiempo_hhmmss": "Tiempo Acumulado (hh:mm:ss)",
                "ejecuciones": "Ejecuciones",
                "repeticiones_totales": "Repeticiones Totales",
                "promedio_min": "Promedio (min)",
                "max_min": "Máximo (min)",
                "desviacion_min": "Desv. Estándar (min)",
            }
        )

        st.dataframe(
            df_jobs_tabla.style.format(
                {
                    "Ejecuciones": "{:.0f}",
                    "Repeticiones Totales": "{:.0f}",
                    "Promedio (min)": "{:.2f}",
                    "Máximo (min)": "{:.2f}",
                    "Desv. Estándar (min)": "{:.2f}",
                },
                na_rep="—",
            ),
            use_container_width=True,
            hide_index=True,
        )
        st.caption(
            "Nota: 'Desv. Estándar (min)' muestra '—' cuando el job tuvo solo 1 ejecución en el periodo filtrado "
            "(se requieren al menos 2 ejecuciones para medir variabilidad)."
        )

# === TAB 3: CONCURRENCIA HORARIA ===
with tab3:
    if df_horario.empty:
        st.warning("Sin datos disponibles para los filtros actuales.")
    else:
        render_vista_horaria_interactiva(
            df_horario=df_horario,
            prefijo="tab3",
            titulo_grafico="Distribución de Carga por Rango Horario (Detección de Cuellos de Botella)",
            titulo_resumen="Resumen por Rango Horario",
            meses_sel=meses_sel,
            grupos_sel=grupos_sel,
            horarios_sel=horarios_sel,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
        )

# === TAB 4: IMPACTO Y VARIABILIDAD ===
with tab4:
    if df_por_fecha.empty:
        st.warning("Sin datos disponibles para los filtros actuales.")
    elif es_un_solo_dia and not df_horario.empty:
        fecha_sel = df_por_fecha["fecha"].iloc[0]
        st.subheader(f"Tendencia Horaria: Escenario Actual — {fecha_sel}")
        df_h4 = df_horario.copy()
        df_h4["tiempo_hhmmss"] = df_h4["horas_totales"].apply(formato_hhmmss)
        df_h4["proc_str"] = df_h4["cantidad_procesos"].astype(int).astype(str)
        df_h4["prom_str"] = df_h4["promedio_min"].map(lambda x: f"{x:.2f}")

        fig_actual = px.line(
            df_h4,
            x="rango_horario",
            y="horas_totales",
            markers=True,
            custom_data=["tiempo_hhmmss", "proc_str", "prom_str"],
            category_orders={"rango_horario": df_h4["rango_horario"].tolist()},
            labels={"rango_horario": "Rango Horario", "horas_totales": "Tiempo Acumulado (hh:mm:ss)"},
        )
        fig_actual.update_traces(
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Tiempo Acumulado: %{customdata[0]}<br>"
                "Procesos Ejecutados: %{customdata[1]}<br>"
                "Promedio (min/proc): %{customdata[2]}<extra></extra>"
            )
        )
        fig_actual.update_layout(xaxis_type="category", xaxis_tickangle=-45)
        promedio_hora = df_h4["horas_totales"].mean()
        fig_actual.add_hline(
            y=promedio_hora,
            line_dash="dash",
            line_color="#9aa4b2",
            annotation_text=f"Promedio por franja: {formato_hhmmss(promedio_hora)}",
            annotation_position="top right",
        )
        aplicar_eje_tiempo_hhmmss(fig_actual, df_h4["horas_totales"].max(), eje="y")
        st.plotly_chart(fig_actual, use_container_width=True, key="chart_tab4_horario")

        st.info(
            "**Nota de análisis:** Los procesos con alta **desviación estándar (`Desv. Estándar (min)`)** en la pestaña "
            "'Top Procesos Críticos' son los más sensibles a picos de concurrencia o congestión operativa."
        )
    else:
        st.subheader("Tendencia Operativa: Escenario Actual")
        df_esc = df_por_fecha.copy()
        df_esc["tiempo_hhmmss"] = df_esc["horas_totales"].apply(formato_hhmmss)
        df_esc["proc_str"] = df_esc["total_ejecuciones"].astype(int).astype(str)

        fig_actual = px.line(
            df_esc,
            x="fecha",
            y="horas_totales",
            custom_data=["tiempo_hhmmss", "proc_str"],
            labels={"fecha": "Fecha", "horas_totales": "Tiempo Acumulado (hh:mm:ss)"},
        )
        fig_actual.update_traces(
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Tiempo Acumulado: %{customdata[0]}<br>"
                "Procesos Ejecutados: %{customdata[1]}<extra></extra>"
            )
        )
        promedio_actual = df_esc["horas_totales"].mean()
        fig_actual.add_hline(
            y=promedio_actual,
            line_dash="dash",
            line_color="#9aa4b2",
            annotation_text=f"Promedio diario: {formato_hhmmss(promedio_actual)}",
            annotation_position="top right",
        )
        aplicar_eje_tiempo_hhmmss(fig_actual, df_esc["horas_totales"].max(), eje="y")
        st.plotly_chart(fig_actual, use_container_width=True, key="chart_tab4_diario")

        st.info(
            "**Nota de análisis:** Los procesos con alta **desviación estándar (`Desv. Estándar (min)`)** en la pestaña "
            "'Top Procesos Críticos' son los más sensibles a picos de concurrencia o congestión operativa."
        )

# === TAB 5: COMPARATIVA DIARIA Y MENSUAL ===
with tab5:
    st.subheader("⚖️ Comparativa de Gráficos de Barras (Diaria y Mensual)")
    st.caption(
        "Compara múltiples días o meses lado a lado. Respeta los filtros de Grupo y Franja Horaria de la barra lateral "
        "y permite elegir libremente cualquier fecha o mes disponible en Supabase."
    )

    if df_fechas_disp.empty:
        st.warning("No hay datos cargados para comparar.")
    else:
        col_m1, col_m2 = st.columns([2, 2])
        with col_m1:
            modo_comp = st.radio(
                "Nivel de comparativa:",
                options=["📅 Comparativa Diaria (entre días)", "🗓️ Comparativa Mensual (entre meses)"],
                horizontal=True,
                key="comp_modo",
            )
        with col_m2:
            metrica_comp = st.radio(
                "Métrica a comparar:",
                options=["Cantidad de Procesos", "Tiempo Acumulado (hh:mm:ss)"],
                horizontal=True,
                key="comp_metrica",
            )

        if modo_comp.startswith("📅"):
            fechas_todas = sorted(df_fechas_disp["fecha"].dropna().unique().tolist(), reverse=True)
            default_dias = fechas_todas[:2] if len(fechas_todas) >= 2 else fechas_todas

            dias_comp = st.multiselect(
                "Selecciona 2 o más días para comparar:",
                options=fechas_todas,
                default=default_dias,
                placeholder="Elige los días a comparar...",
                key="comp_dias_sel",
            )

            if len(dias_comp) < 2:
                st.info("Selecciona al menos **2 días** para ver la comparativa lado a lado.")
            else:
                df_c_res, df_c_hora, df_c_grp = consultar_comparativa_diaria(dias_comp, grupos_sel, horarios_sel)

                if df_c_res.empty:
                    st.warning("No hay registros para los días y filtros de segmentación seleccionados.")
                else:
                    df_res_t = df_c_res.copy()
                    df_res_t["Tiempo Acumulado (hh:mm:ss)"] = df_res_t["horas_totales"].apply(formato_hhmmss)
                    st.markdown("**Resumen de los días seleccionados**")
                    st.dataframe(
                        df_res_t[["fecha_str", "cantidad_procesos", "Tiempo Acumulado (hh:mm:ss)", "promedio_min"]]
                        .rename(columns={
                            "fecha_str": "Fecha",
                            "cantidad_procesos": "Procesos Ejecutados",
                            "promedio_min": "Promedio (min/proc)",
                        })
                        .style.format({"Procesos Ejecutados": "{:.0f}", "Promedio (min/proc)": "{:.2f}"}),
                        use_container_width=True,
                        hide_index=True,
                    )

                    st.subheader("Comparativa por Rango Horario")
                    df_ch = df_c_hora.copy()
                    df_ch["tiempo_hhmmss"] = df_ch["horas_totales"].apply(formato_hhmmss)
                    df_ch["proc_str"] = df_ch["cantidad_procesos"].astype(int).astype(str)
                    orden_horas_c = sorted(df_ch["rango_horario"].unique().tolist(), key=clave_orden_horario)
                    orden_fechas_c = sorted(df_ch["fecha_str"].unique().tolist())

                    y_col = "cantidad_procesos" if metrica_comp == "Cantidad de Procesos" else "horas_totales"
                    y_lab = "Procesos Ejecutados" if metrica_comp == "Cantidad de Procesos" else "Tiempo Acumulado (hh:mm:ss)"

                    fig_ch = px.bar(
                        df_ch,
                        x="rango_horario",
                        y=y_col,
                        color="fecha_str",
                        barmode="group",
                        custom_data=["fecha_str", "proc_str", "tiempo_hhmmss"],
                        category_orders={"rango_horario": orden_horas_c, "fecha_str": orden_fechas_c},
                        labels={"rango_horario": "Rango Horario", "fecha_str": "Fecha"},
                    )
                    fig_ch.update_traces(
                        hovertemplate=(
                            "<b>%{x}</b><br>"
                            "Fecha: %{customdata[0]}<br>"
                            "Procesos Ejecutados: %{customdata[1]}<br>"
                            "Tiempo Acumulado: %{customdata[2]}<extra></extra>"
                        )
                    )
                    fig_ch.update_layout(xaxis_type="category", yaxis_title=y_lab, xaxis_tickangle=-45)
                    if metrica_comp == "Cantidad de Procesos":
                        aplicar_eje_entero(fig_ch, eje="y")
                    else:
                        aplicar_eje_tiempo_hhmmss(fig_ch, df_ch["horas_totales"].max(), eje="y")
                    st.plotly_chart(fig_ch, use_container_width=True, key="chart_comp_dia_hora")

                    if not df_c_grp.empty:
                        st.subheader("Comparativa por Grupo Operativo")
                        df_cg = df_c_grp.copy()
                        df_cg["tiempo_hhmmss"] = df_cg["horas_totales"].apply(formato_hhmmss)
                        df_cg["proc_str"] = df_cg["cantidad_procesos"].astype(int).astype(str)

                        fig_cg = px.bar(
                            df_cg,
                            x="grupo",
                            y=y_col,
                            color="fecha_str",
                            barmode="group",
                            custom_data=["fecha_str", "proc_str", "tiempo_hhmmss"],
                            category_orders={"fecha_str": orden_fechas_c},
                            labels={"grupo": "Grupo", "fecha_str": "Fecha"},
                        )
                        fig_cg.update_traces(
                            hovertemplate=(
                                "<b>%{x}</b><br>"
                                "Fecha: %{customdata[0]}<br>"
                                "Procesos Ejecutados: %{customdata[1]}<br>"
                                "Tiempo Acumulado: %{customdata[2]}<extra></extra>"
                            )
                        )
                        fig_cg.update_layout(xaxis_type="category", yaxis_title=y_lab, xaxis_tickangle=-35)
                        if metrica_comp == "Cantidad de Procesos":
                            aplicar_eje_entero(fig_cg, eje="y")
                        else:
                            aplicar_eje_tiempo_hhmmss(fig_cg, df_cg["horas_totales"].max(), eje="y")
                        st.plotly_chart(fig_cg, use_container_width=True, key="chart_comp_dia_grupo")

        else:
            default_meses = meses_disp[-3:] if len(meses_disp) > 3 else meses_disp
            meses_comp = st.multiselect(
                "Selecciona 2 o más meses para comparar:",
                options=meses_disp,
                default=default_meses,
                placeholder="Elige los meses a comparar...",
                key="comp_meses_sel",
            )

            if len(meses_comp) < 2:
                st.info("Selecciona al menos **2 meses** para ver la comparativa lado a lado.")
            else:
                meses_comp_ord = sorted(meses_comp, key=clave_orden_mes)
                df_m_res, df_m_dia, df_m_hora = consultar_comparativa_mensual(meses_comp_ord, grupos_sel, horarios_sel)

                if df_m_res.empty:
                    st.warning("No hay registros para los meses y filtros de segmentación seleccionados.")
                else:
                    y_col = "cantidad_procesos" if metrica_comp == "Cantidad de Procesos" else "horas_totales"
                    y_lab = "Procesos Ejecutados" if metrica_comp == "Cantidad de Procesos" else "Tiempo Acumulado (hh:mm:ss)"

                    st.subheader("Comparativa Total por Mes")
                    df_mr = df_m_res.copy()
                    df_mr["tiempo_hhmmss"] = df_mr["horas_totales"].apply(formato_hhmmss)
                    df_mr["prom_dia_hhmmss"] = df_mr["promedio_diario_hrs"].apply(formato_hhmmss)
                    df_mr["proc_str"] = df_mr["cantidad_procesos"].astype(int).astype(str)
                    df_mr["dias_str"] = df_mr["dias_cargados"].astype(int).astype(str)
                    text_m = "proc_str" if metrica_comp == "Cantidad de Procesos" else "tiempo_hhmmss"

                    fig_mt = px.bar(
                        df_mr,
                        x="mes",
                        y=y_col,
                        color="mes",
                        text=text_m,
                        custom_data=["dias_str", "proc_str", "tiempo_hhmmss", "prom_dia_hhmmss"],
                        category_orders={"mes": meses_comp_ord},
                        labels={"mes": "Mes"},
                    )
                    fig_mt.update_traces(
                        textposition="outside",
                        hovertemplate=(
                            "<b>%{x}</b><br>"
                            "Días Cargados: %{customdata[0]}<br>"
                            "Procesos Ejecutados: %{customdata[1]}<br>"
                            "Tiempo Acumulado: %{customdata[2]}<br>"
                            "Promedio Diario: %{customdata[3]}<extra></extra>"
                        ),
                    )
                    fig_mt.update_layout(xaxis_type="category", yaxis_title=y_lab, showlegend=False)
                    if metrica_comp == "Cantidad de Procesos":
                        aplicar_eje_entero(fig_mt, eje="y")
                    else:
                        aplicar_eje_tiempo_hhmmss(fig_mt, df_mr["horas_totales"].max(), eje="y")
                    st.plotly_chart(fig_mt, use_container_width=True, key="chart_comp_mes_total")

                    st.dataframe(
                        df_mr[["mes", "dias_cargados", "cantidad_procesos", "tiempo_hhmmss", "prom_dia_hhmmss"]]
                        .rename(columns={
                            "mes": "Mes",
                            "dias_cargados": "Días Cargados",
                            "cantidad_procesos": "Procesos Ejecutados",
                            "tiempo_hhmmss": "Tiempo Acumulado (hh:mm:ss)",
                            "prom_dia_hhmmss": "Promedio Diario (hh:mm:ss)",
                        })
                        .style.format({"Días Cargados": "{:.0f}", "Procesos Ejecutados": "{:.0f}"}),
                        use_container_width=True,
                        hide_index=True,
                    )

                    st.subheader("Comparativa por Día del Mes (01 al 31)")
                    df_md = df_m_dia.copy()
                    df_md["dia_label"] = df_md["dia_mes"].map(lambda d: f"{int(d):02d}")
                    df_md["tiempo_hhmmss"] = df_md["horas_totales"].apply(formato_hhmmss)
                    df_md["proc_str"] = df_md["cantidad_procesos"].astype(int).astype(str)
                    orden_dias_mes = [f"{d:02d}" for d in range(1, 32)]

                    fig_md = px.bar(
                        df_md,
                        x="dia_label",
                        y=y_col,
                        color="mes",
                        barmode="group",
                        custom_data=["mes", "proc_str", "tiempo_hhmmss"],
                        category_orders={"dia_label": orden_dias_mes, "mes": meses_comp_ord},
                        labels={"dia_label": "Día del Mes", "mes": "Mes"},
                    )
                    fig_md.update_traces(
                        hovertemplate=(
                            "<b>Día %{x}</b><br>"
                            "Mes: %{customdata[0]}<br>"
                            "Procesos Ejecutados: %{customdata[1]}<br>"
                            "Tiempo Acumulado: %{customdata[2]}<extra></extra>"
                        )
                    )
                    fig_md.update_layout(xaxis_type="category", yaxis_title=y_lab)
                    if metrica_comp == "Cantidad de Procesos":
                        aplicar_eje_entero(fig_md, eje="y")
                    else:
                        aplicar_eje_tiempo_hhmmss(fig_md, df_md["horas_totales"].max(), eje="y")
                    st.plotly_chart(fig_md, use_container_width=True, key="chart_comp_mes_dia")

                    st.subheader("Comparativa por Rango Horario entre Meses")
                    df_mh = df_m_hora.copy()
                    df_mh["tiempo_hhmmss"] = df_mh["horas_totales"].apply(formato_hhmmss)
                    df_mh["proc_str"] = df_mh["cantidad_procesos"].astype(int).astype(str)
                    orden_horas_m = sorted(df_mh["rango_horario"].unique().tolist(), key=clave_orden_horario)

                    fig_mh = px.bar(
                        df_mh,
                        x="rango_horario",
                        y=y_col,
                        color="mes",
                        barmode="group",
                        custom_data=["mes", "proc_str", "tiempo_hhmmss"],
                        category_orders={"rango_horario": orden_horas_m, "mes": meses_comp_ord},
                        labels={"rango_horario": "Rango Horario", "mes": "Mes"},
                    )
                    fig_mh.update_traces(
                        hovertemplate=(
                            "<b>%{x}</b><br>"
                            "Mes: %{customdata[0]}<br>"
                            "Procesos Ejecutados: %{customdata[1]}<br>"
                            "Tiempo Acumulado: %{customdata[2]}<extra></extra>"
                        )
                    )
                    fig_mh.update_layout(xaxis_type="category", yaxis_title=y_lab, xaxis_tickangle=-45)
                    if metrica_comp == "Cantidad de Procesos":
                        aplicar_eje_entero(fig_mh, eje="y")
                    else:
                        aplicar_eje_tiempo_hhmmss(fig_mh, df_mh["horas_totales"].max(), eje="y")
                    st.plotly_chart(fig_mh, use_container_width=True, key="chart_comp_mes_hora")

# === TAB 6: ANÁLISIS DE PROCESOS DIARIOS VS MENSUALES ===
feriados_auto_global = detectar_posibles_feriados(df_fechas_disp)

with tab6:
    st.subheader("🧭 Auditoría de Cadencia: Procesos Diarios vs. Mensuales")
    st.caption(
        "Analiza el histórico completo cargado en Supabase para identificar procesos mensuales (días 30, 31, fin de mes o 1), "
        "procesos diarios con sus excepciones (sábados, domingos o feriados) y detectar altas o bajas reales en el catálogo. "
        "Respeta las clasificaciones manuales guardadas desde ⚙️ Gestión de Datos."
    )

    if df_fechas_disp.empty:
        st.warning("No hay datos históricos cargados para realizar el análisis de cadencia.")
    else:
        fechas_hist_ordenadas = sorted(df_fechas_disp["fecha"].dropna().unique().tolist())

        col_f1, col_f2 = st.columns([2, 3])
        with col_f1:
            grupos_aud = st.multiselect(
                "Filtrar auditoría por Grupo (opcional):",
                options=grupos_disp,
                default=grupos_sel,
                placeholder="Todos los grupos",
                key="aud_grupos_sel",
            )
        with col_f2:
            if "aud_feriados_sel" not in st.session_state:
                st.session_state["aud_feriados_sel"] = feriados_auto_global
            else:
                st.session_state["aud_feriados_sel"] = [
                    f for f in st.session_state["aud_feriados_sel"] if f in fechas_hist_ordenadas
                ]

            feriados_sel = st.multiselect(
                "Feriados considerados (detectados por caída de volumen en días laborables; puedes editar):",
                options=fechas_hist_ordenadas,
                key="aud_feriados_sel",
                placeholder="Sin feriados seleccionados",
            )

        with st.spinner("Analizando patrones de ejecución en todo el histórico..."):
            df_perfil = consultar_perfil_cadencia_jobs(tuple(grupos_aud), tuple(sorted(feriados_sel)))
            df_clas, df_estab, df_cambios = clasificar_cadencia_y_estabilidad(
                df_perfil,
                df_fechas_disp,
                feriados_sel,
                overrides_manuales=overrides_manuales,
            )

        if df_clas.empty:
            st.warning("No se encontraron procesos con los filtros seleccionados.")
        else:
            n_total_jobs = len(df_clas)
            df_diarios_cat = df_clas[df_clas["categoria"] == "Diario"]
            df_mensuales_cat = df_clas[df_clas["categoria"] == "Mensual"]
            df_otros_cat = df_clas[df_clas["categoria"] == "Semanal / Otros"]

            n_diarios = len(df_diarios_cat)
            n_diarios_exc = int((df_diarios_cat["excepciones"] != "Ninguna (corre todos los días)").sum())
            n_mensuales = len(df_mensuales_cat)
            n_mensuales_cierre = int(
                df_mensuales_cat["patron"].str.contains("Cierre/Inicio", na=False).sum()
            )
            n_altas = int(df_cambios["es_alta"].sum()) if not df_cambios.empty else 0
            n_bajas = int(df_cambios["es_baja"].sum()) if not df_cambios.empty else 0

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Jobs Únicos Analizados", f"{n_total_jobs}")
            k2.metric("Procesos Diarios", f"{n_diarios}", f"{n_diarios_exc} con excepciones", delta_color="off")
            k3.metric("Procesos Mensuales", f"{n_mensuales}", f"{n_mensuales_cierre} en días 30/31/1", delta_color="off")
            k4.metric("Semanal / Otros", f"{len(df_otros_cat)}")
            k5.metric("Cambios Confirmados", f"{n_altas + n_bajas}", f"+{n_altas} altas / -{n_bajas} bajas", delta_color="off")

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.markdown("**Distribución por Tipo de Proceso**")
                df_cat_cnt = df_clas.groupby("categoria", as_index=False).agg(cantidad=("jobname", "count"))
                df_cat_cnt["cant_str"] = df_cat_cnt["cantidad"].astype(int).astype(str)
                fig_cat = px.bar(
                    df_cat_cnt,
                    x="categoria",
                    y="cantidad",
                    color="categoria",
                    text="cant_str",
                    custom_data=["cant_str"],
                    labels={"categoria": "Categoría", "cantidad": "Cantidad de Procesos"},
                )
                fig_cat.update_traces(
                    textposition="outside",
                    hovertemplate="<b>%{x}</b><br>Procesos: %{customdata[0]}<extra></extra>",
                )
                fig_cat.update_layout(showlegend=False, xaxis_type="category")
                aplicar_eje_entero(fig_cat, eje="y")
                st.plotly_chart(fig_cat, use_container_width=True, key="chart_aud_cat")

            with col_g2:
                st.markdown("**Desglose de Excepciones en Procesos Diarios y Mensuales**")
                df_pat_cnt = (
                    df_clas[df_clas["categoria"].isin(["Diario", "Mensual"])]
                    .groupby(["patron", "categoria"], as_index=False)
                    .agg(cantidad=("jobname", "count"))
                    .sort_values("cantidad", ascending=True)
                )
                if not df_pat_cnt.empty:
                    df_pat_cnt["cant_str"] = df_pat_cnt["cantidad"].astype(int).astype(str)
                    fig_exc = px.bar(
                        df_pat_cnt,
                        x="cantidad",
                        y="patron",
                        color="categoria",
                        orientation="h",
                        text="cant_str",
                        custom_data=["categoria", "cant_str"],
                        labels={"cantidad": "Cantidad de Procesos", "patron": "Patrón / Excepción", "categoria": "Categoría"},
                    )
                    fig_exc.update_traces(
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Categoría: %{customdata[0]}<br>Procesos: %{customdata[1]}<extra></extra>",
                    )
                    aplicar_eje_entero(fig_exc, eje="x")
                    st.plotly_chart(fig_exc, use_container_width=True, key="chart_aud_exc")

            st.markdown("---")
            st.subheader("📈 Estabilidad Mensual del Catálogo (Vigentes vs. Ejecutados)")
            st.caption(
                "**Vigentes** refleja el catálogo activo (cambia solo por Altas o Bajas confirmadas). "
                "**Ejecutados** muestra cuántos corrieron ese mes; la diferencia ('Mensuales Sin Corrida') corresponde "
                "a procesos vigentes que saltaron ese mes o cuyo cierre cayó en el mes contiguo."
            )

            df_estab_long = df_estab.melt(
                id_vars=["Mes"],
                value_vars=["Diarios Vigentes", "Mensuales Vigentes", "Mensuales Ejecutados"],
                var_name="Tipo",
                value_name="Cantidad",
            )
            df_estab_long["cant_str"] = df_estab_long["Cantidad"].astype(int).astype(str)
            fig_est = px.bar(
                df_estab_long,
                x="Mes",
                y="Cantidad",
                color="Tipo",
                barmode="group",
                text="cant_str",
                custom_data=["Tipo", "cant_str"],
                category_orders={"Mes": df_estab["Mes"].tolist()},
                labels={"Mes": "Mes", "Cantidad": "Cantidad de Procesos"},
            )
            fig_est.update_traces(
                textposition="outside",
                hovertemplate="<b>%{x}</b><br>%{customdata[0]}: %{customdata[1]}<extra></extra>",
            )
            fig_est.update_layout(xaxis_type="category")
            aplicar_eje_entero(fig_est, eje="y")
            st.plotly_chart(fig_est, use_container_width=True, key="chart_aud_estabilidad")

            st.dataframe(
                df_estab[
                    [
                        "Mes",
                        "Diarios Vigentes",
                        "Diarios Ejecutados",
                        "Altas Diarias (+)",
                        "Bajas Diarias (-)",
                        "Mensuales Vigentes",
                        "Mensuales Ejecutados",
                        "Mensuales Sin Corrida",
                        "Altas Mensuales (+)",
                        "Bajas Mensuales (-)",
                    ]
                ].style.format({
                    "Diarios Vigentes": "{:.0f}",
                    "Diarios Ejecutados": "{:.0f}",
                    "Altas Diarias (+)": "{:.0f}",
                    "Bajas Diarias (-)": "{:.0f}",
                    "Mensuales Vigentes": "{:.0f}",
                    "Mensuales Ejecutados": "{:.0f}",
                    "Mensuales Sin Corrida": "{:.0f}",
                    "Altas Mensuales (+)": "{:.0f}",
                    "Bajas Mensuales (-)": "{:.0f}",
                }),
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("---")
            st.subheader("🔄 Registro de Cambios en el Catálogo (Altas y Bajas)")
            if df_cambios.empty:
                st.success("No se detectaron altas ni bajas confirmadas: el catálogo de procesos diarios y mensuales se mantuvo estable.")
            else:
                df_camb_vista = df_cambios[
                    [
                        "estado",
                        "categoria",
                        "origen_clas",
                        "jobname",
                        "grupo",
                        "primera_fecha",
                        "ultima_fecha",
                        "dias_ejecutados",
                        "patron",
                    ]
                ].rename(columns={
                    "estado": "Movimiento",
                    "categoria": "Categoría",
                    "origen_clas": "Origen",
                    "jobname": "Proceso (JOBNAME)",
                    "grupo": "Grupo",
                    "primera_fecha": "Primera Ejecución (Alta)",
                    "ultima_fecha": "Última Ejecución",
                    "dias_ejecutados": "Días Ejecutados",
                    "patron": "Patrón / Excepciones",
                })
                st.dataframe(
                    df_camb_vista.style.format({"Días Ejecutados": "{:.0f}"}),
                    use_container_width=True,
                    hide_index=True,
                )

            st.markdown("---")
            st.subheader("📋 Catálogo Completo Clasificado")
            c_f1, c_f2, c_f3 = st.columns([2, 2, 2])
            with c_f1:
                filtro_cat_aud = st.multiselect(
                    "Filtrar por Categoría:",
                    options=["Diario", "Mensual", "Semanal / Otros"],
                    default=["Diario", "Mensual"],
                    key="aud_filtro_cat",
                )
            with c_f2:
                estados_disp = sorted(df_clas["estado"].unique().tolist())
                filtro_est_aud = st.multiselect(
                    "Filtrar por Estado:",
                    options=estados_disp,
                    default=[],
                    placeholder="Todos los estados",
                    key="aud_filtro_estado",
                )
            with c_f3:
                busq_aud = st.text_input(
                    "Buscar JOBNAME o excepción:",
                    placeholder="Ej. DPAHD10, domingos...",
                    key="aud_busq_job",
                )

            df_cat_vista = df_clas.copy()
            if filtro_cat_aud:
                df_cat_vista = df_cat_vista[df_cat_vista["categoria"].isin(filtro_cat_aud)]
            if filtro_est_aud:
                df_cat_vista = df_cat_vista[df_cat_vista["estado"].isin(filtro_est_aud)]
            if busq_aud and busq_aud.strip():
                q_txt = busq_aud.strip().lower()
                df_cat_vista = df_cat_vista[
                    df_cat_vista["jobname"].str.lower().str.contains(q_txt, na=False)
                    | df_cat_vista["patron"].str.lower().str.contains(q_txt, na=False)
                    | df_cat_vista["excepciones"].str.lower().str.contains(q_txt, na=False)
                ]

            st.caption(
                f"Mostrando **{len(df_cat_vista)}** proceso(s). "
                "¿Algún proceso quedó mal clasificado (ej. aparece Mensual pero es Diario)? "
                "Cámbialo o elimínalo en **⚙️ Gestión de Datos → Gestión de Procesos**."
            )
            st.dataframe(
                df_cat_vista[
                    [
                        "jobname",
                        "grupo",
                        "categoria",
                        "origen_clas",
                        "patron",
                        "estado",
                        "primera_fecha",
                        "ultima_fecha",
                        "dias_ejecutados",
                        "meses_con_ejecucion",
                        "meses_faltantes",
                        "tiempo_acum_hhmmss",
                        "prom_dia_hhmmss",
                    ]
                ]
                .rename(columns={
                    "jobname": "Proceso (JOBNAME)",
                    "grupo": "Grupo",
                    "categoria": "Categoría",
                    "origen_clas": "Origen",
                    "patron": "Patrón / Excepciones",
                    "estado": "Vigencia",
                    "primera_fecha": "Primera Fecha",
                    "ultima_fecha": "Última Fecha",
                    "dias_ejecutados": "Días Ejecutados",
                    "meses_con_ejecucion": "Meses Activos",
                    "meses_faltantes": "Meses Faltantes",
                    "tiempo_acum_hhmmss": "Tiempo Acumulado (hh:mm:ss)",
                    "prom_dia_hhmmss": "Promedio por Día (hh:mm:ss)",
                })
                .style.format({
                    "Días Ejecutados": "{:.0f}",
                    "Meses Activos": "{:.0f}",
                }),
                use_container_width=True,
                hide_index=True,
            )

# === TAB 7: GESTIÓN DE DATOS (CARGA, BORRADO DE DÍAS Y GESTIÓN DE PROCESOS) ===
with tab7:
    if not df_inconsistencias.empty:
        with st.container(border=True):
            st.subheader("🧹 Limpieza de Duplicados Cruzados y Normalización de Meses")
            st.caption(
                "Se detectaron registros cuya columna `mes` no coincide con la fecha real de la hoja "
                "(por ejemplo, un día 01 cargado dentro del Excel del mes anterior o escrito como 'Setiembre')."
            )
            st.dataframe(
                df_inconsistencias.rename(columns={
                    "fecha": "Fecha",
                    "mes": "Mes Actual en BD",
                    "mes_correcto": "Mes Correcto",
                    "registros": "Registros",
                    "accion": "Acción Automática",
                }).style.format({"Registros": "{:.0f}"}),
                use_container_width=True,
                hide_index=True,
            )

            if st.button(
                "🧹 Limpiar duplicados cruzados y normalizar meses",
                type="primary",
                key="btn_limpiar_cruzados",
            ):
                with st.spinner("Limpiando duplicados y normalizando etiquetas en Supabase..."):
                    borrados, actualizados = limpiar_y_normalizar_meses_supabase()
                refrescar_app(
                    f"Limpieza completada: {borrados} registros duplicados eliminados y "
                    f"{actualizados} registros normalizados al mes calendario real."
                )

    col_subir, col_borrar = st.columns(2, gap="large")

    # --- SECCIÓN 1: SUBIR NUEVOS DÍAS / MESES ---
    with col_subir:
        st.subheader("📤 Cargar archivos Excel a Supabase")
        st.caption(
            "Sube uno o varios `.xlsx` diarios o mensuales (`procesos_<Mes>_por_dia.xlsx`). "
            "Calcula el mes real desde cada hoja (`DD-MM-YY`), deduplica fechas en el lote y evita cruces de mes."
        )

        archivos_subidos = st.file_uploader(
            "Seleccionar archivo(s) Excel",
            type=["xlsx"],
            accept_multiple_files=True,
            key="uploader_excel",
        )

        solo_mes_archivo = st.checkbox(
            "Ignorar hojas fuera del mes del archivo (evita cargar el día 01 del mes siguiente si viene pegado en el Excel)",
            value=True,
            key="chk_solo_mes_archivo",
        )

        usar_fecha_manual = st.checkbox(
            "Asignar fecha manualmente (solo si subes un Excel de 1 sola hoja cuyo nombre no es fecha)",
            value=False,
            key="chk_fecha_manual",
        )
        fecha_manual = st.date_input("Fecha para la hoja:") if usar_fecha_manual else None

        reemplazar_existentes = st.checkbox(
            "Reemplazar si alguna de las fechas ya existe en Supabase (evita duplicados)",
            value=True,
            key="chk_reemplazar",
        )

        if archivos_subidos:
            try:
                df_preview, omitidas, avisos_lote = procesar_excels_subidos(
                    archivos_subidos,
                    fecha_manual=fecha_manual,
                    solo_mes_del_archivo=solo_mes_archivo,
                )

                if avisos_lote:
                    with st.expander(f"ℹ️ Fechas deduplicadas en el lote ({len(avisos_lote)})"):
                        for av in avisos_lote:
                            st.write(f"- {av}")

                if omitidas:
                    with st.expander(f"⚠️ Hojas omitidas ({len(omitidas)})"):
                        for item in omitidas:
                            st.write(f"- {item}")

                if df_preview.empty:
                    st.error("No se encontraron hojas válidas para cargar.")
                else:
                    resumen_nuevo = (
                        df_preview.groupby(["fecha", "mes"], as_index=False)
                        .agg(
                            registros=("jobname", "count"),
                            seg_totales=("tiempo_demora_seg", "sum"),
                        )
                        .sort_values("fecha")
                    )
                    resumen_nuevo["Tiempo Acumulado (hh:mm:ss)"] = resumen_nuevo["seg_totales"].apply(formato_seg_hhmmss)

                    st.write(f"**Vista previa:** {len(resumen_nuevo)} día(s) detectado(s) | {len(df_preview)} registros")
                    st.dataframe(
                        resumen_nuevo[["fecha", "mes", "registros", "Tiempo Acumulado (hh:mm:ss)"]]
                        .rename(columns={"fecha": "Fecha", "mes": "Mes", "registros": "Registros"})
                        .style.format({"Registros": "{:.0f}"}),
                        use_container_width=True,
                        hide_index=True,
                    )

                    if st.button("🚀 Cargar días a Supabase", type="primary", use_container_width=True, key="btn_cargar_supa"):
                        with st.spinner("Subiendo registros a Supabase..."):
                            n_reg, n_dias, n_reemp = subir_datos_a_supabase(
                                df_preview, reemplazar_existentes=reemplazar_existentes
                            )
                        msg_reemp = f" ({n_reemp} registros previos reemplazados)" if n_reemp else ""
                        refrescar_app(
                            f"Se cargaron {n_reg} registros correspondientes a {n_dias} día(s){msg_reemp}. Gráficos actualizados."
                        )
            except Exception as e:
                st.error(f"Error procesando los archivos Excel: {e}")

        carpeta_local = BASE_DIR / "data"
        archivos_locales = (
            sorted(
                carpeta_local.glob("procesos_*_por_dia.xlsx"),
                key=lambda p: (extraer_mes_de_archivo(p.name) or 99, p.name.lower()),
            )
            if carpeta_local.is_dir()
            else []
        )

        with st.expander(f"📁 Importar desde carpeta local `data/` ({len(archivos_locales)} archivo(s))"):
            if not carpeta_local.is_dir():
                st.caption(f"No se encontró la carpeta `{carpeta_local}`.")
            elif not archivos_locales:
                st.caption(f"No hay archivos `procesos_*_por_dia.xlsx` en `{carpeta_local}`.")
            else:
                sel_locales = st.multiselect(
                    "Selecciona archivos de la carpeta `data/` para cargar o recargar:",
                    options=archivos_locales,
                    format_func=lambda p: p.name,
                    default=[],
                    placeholder="Elige uno o más archivos locales...",
                    key="ms_archivos_locales",
                )
                if sel_locales:
                    if st.button(
                        "📥 Procesar y cargar seleccionados desde `data/`",
                        type="primary",
                        use_container_width=True,
                        key="btn_cargar_locales",
                    ):
                        try:
                            with st.spinner("Leyendo archivos de `data/` y subiendo a Supabase..."):
                                df_loc, omit_loc, _ = procesar_excels_subidos(
                                    sel_locales,
                                    solo_mes_del_archivo=solo_mes_archivo,
                                )
                                if df_loc.empty:
                                    st.error("No se encontraron hojas válidas en los archivos seleccionados.")
                                else:
                                    n_reg, n_dias, n_reemp = subir_datos_a_supabase(
                                        df_loc, reemplazar_existentes=reemplazar_existentes
                                    )
                                    msg_reemp = f" ({n_reemp} previos reemplazados)" if n_reemp else ""
                                    msg_omit = f" · {len(omit_loc)} hoja(s) fuera de mes omitidas" if omit_loc else ""
                                    refrescar_app(
                                        f"Carga desde `data/` completada: {n_reg} registros de {n_dias} día(s){msg_reemp}{msg_omit}."
                                    )
                        except Exception as e:
                            st.error(f"Error cargando desde carpeta local: {e}")

    # --- SECCIÓN 2: BORRAR DÍAS EXISTENTES ---
    with col_borrar:
        st.subheader("🗑️ Eliminar días cargados")
        st.caption("Selecciona una o varias fechas registradas en Supabase para eliminarlas y recalcular los gráficos.")

        if df_fechas_disp.empty:
            st.info("No hay fechas almacenadas actualmente en la base de datos.")
        else:
            mes_filtro_borrar = st.selectbox(
                "Filtrar lista por mes (opcional):",
                options=["Todos"] + meses_disp,
                key="sel_mes_borrar",
            )

            df_fechas_vista = (
                df_fechas_disp.copy()
                if mes_filtro_borrar == "Todos"
                else df_fechas_disp[df_fechas_disp["mes"] == mes_filtro_borrar].copy()
            )
            df_fechas_vista["Tiempo Acumulado (hh:mm:ss)"] = df_fechas_vista["horas_totales"].apply(formato_hhmmss)

            mapa_etiquetas = {
                row["fecha"]: f"{row['fecha']} | {row['mes']} ({int(row['registros'])} reg. / {row['Tiempo Acumulado (hh:mm:ss)']})"
                for _, row in df_fechas_vista.iterrows()
            }

            fechas_sel_borrar = st.multiselect(
                "Seleccionar día(s) a eliminar:",
                options=df_fechas_vista["fecha"].tolist(),
                format_func=lambda f: mapa_etiquetas[f],
                placeholder="Elige uno o más días...",
                key="ms_fechas_borrar",
            )

            if fechas_sel_borrar:
                df_sel = df_fechas_vista[df_fechas_vista["fecha"].isin(fechas_sel_borrar)]
                total_reg_borrar = int(df_sel["registros"].sum())
                st.warning(
                    f"Vas a eliminar **{len(fechas_sel_borrar)} día(s)** ({total_reg_borrar} registros en total). "
                    "Esta acción no se puede deshacer."
                )

                confirmar_borrado = st.checkbox(
                    "Confirmo que deseo eliminar permanentemente los días seleccionados",
                    key="chk_confirmar_borrar",
                )

                if st.button(
                    "🗑️ Eliminar días seleccionados",
                    type="primary",
                    disabled=not confirmar_borrado,
                    use_container_width=True,
                    key="btn_borrar_supa",
                ):
                    with st.spinner("Eliminando registros en Supabase..."):
                        eliminados = eliminar_dias_supabase(fechas_sel_borrar)
                    refrescar_app(
                        f"Se eliminaron {eliminados} registros de {len(fechas_sel_borrar)} día(s). Gráficos actualizados."
                    )

            with st.expander(f"📋 Ver fechas cargadas ({len(df_fechas_vista)} días)"):
                st.dataframe(
                    df_fechas_vista[["fecha", "mes", "registros", "Tiempo Acumulado (hh:mm:ss)"]]
                    .rename(columns={"fecha": "Fecha", "mes": "Mes", "registros": "Registros"})
                    .style.format({"Registros": "{:.0f}"}),
                    use_container_width=True,
                    hide_index=True,
                )

    # --- SECCIÓN 3 (PARTE INFERIOR): GESTIÓN DE PROCESOS (MODIFICAR / ELIMINAR) ---
    st.markdown("---")
    st.subheader("🛠️ Gestión de Procesos (Modificar Clasificación / Editar / Eliminar)")
    st.caption(
        "Corrige la clasificación de un proceso (por ejemplo, cambiar `DPAHD10` de Mensual a Diario), "
        "edita su grupo o nombre en Supabase, o elimina por completo procesos que ya no deban aparecer en el análisis."
    )

    df_jobs_gest = cargar_resumen_jobs_gestion()
    if df_jobs_gest.empty:
        st.info("No hay procesos registrados en la base de datos.")
    else:
        col_mod_proc, col_del_proc = st.columns(2, gap="large")

        # --- SUBSECCIÓN 3A: MODIFICAR PROCESO ---
        with col_mod_proc:
            with st.container(border=True):
                st.markdown("### ✏️ Modificar un proceso")

                busq_mod = st.text_input(
                    "Buscar JOBNAME (ej. DPAHD10):",
                    placeholder="Escribe parte del nombre para filtrar...",
                    key="txt_busq_mod_job",
                )

                df_jobs_mod_opts = df_jobs_gest
                if busq_mod and busq_mod.strip():
                    df_jobs_mod_opts = df_jobs_gest[
                        df_jobs_gest["jobname"].str.contains(busq_mod.strip(), case=False, na=False)
                    ]

                lista_jobs_mod = df_jobs_mod_opts["jobname"].tolist()
                if not lista_jobs_mod:
                    st.warning("No se encontró ningún proceso con ese texto.")
                else:
                    job_sel_mod = st.selectbox(
                        "Seleccionar proceso (JOBNAME):",
                        options=lista_jobs_mod,
                        key="sel_job_mod",
                    )

                    fila_job = df_jobs_gest[df_jobs_gest["jobname"] == job_sel_mod].iloc[0]
                    ov_actual = overrides_manuales.get(job_sel_mod)

                    df_perf_uno = consultar_perfil_cadencia_jobs(
                        (), tuple(sorted(feriados_auto_global)), jobname_unico=job_sel_mod
                    )
                    df_clas_uno, _, _ = clasificar_cadencia_y_estabilidad(
                        df_perf_uno,
                        df_fechas_disp,
                        feriados_auto_global,
                        overrides_manuales=overrides_manuales,
                    )

                    if not df_clas_uno.empty:
                        info_c = df_clas_uno.iloc[0]
                        cat_actual = info_c["categoria"]
                        cat_auto = info_c["categoria_auto"]
                        patron_actual = info_c["patron"]
                        origen_actual = info_c["origen_clas"]
                    else:
                        cat_actual = ov_actual["categoria_manual"] if ov_actual else "Semanal / Otros"
                        cat_auto = "Semanal / Otros"
                        patron_actual = ov_actual["patron_manual"] if ov_actual else ""
                        origen_actual = "Manual" if ov_actual else "Automática"

                    st.info(
                        f"**{job_sel_mod}** · Grupo: `{fila_job['grupo']}`\n\n"
                        f"Registros: **{int(fila_job['registros'])}** · "
                        f"Días ejecutados: **{int(fila_job['dias_ejecutados'])}** · "
                        f"Meses activos: **{int(fila_job['meses_activos'])}** · "
                        f"Tiempo acum.: **{formato_seg_hhmmss(fila_job['seg_totales'])}**\n\n"
                        f"Clasificación vigente: **{cat_actual}** ({origen_actual}) — *{patron_actual}*"
                        + (f"  \n*(Detección automática original: {cat_auto})*" if origen_actual == "Manual" else "")
                    )

                    categorias_opts = ["Diario", "Mensual", "Semanal / Otros"]
                    idx_cat = categorias_opts.index(cat_actual) if cat_actual in categorias_opts else 0

                    nueva_cat = st.selectbox(
                        "Nueva categoría:",
                        options=categorias_opts,
                        index=idx_cat,
                        key="sel_nueva_cat_job",
                    )

                    patrones_opts = PATRONES_POR_CATEGORIA[nueva_cat]
                    idx_pat = patrones_opts.index(patron_actual) if patron_actual in patrones_opts else 0
                    patron_elegido = st.selectbox(
                        "Patrón / Excepciones:",
                        options=patrones_opts,
                        index=idx_pat,
                        key="sel_nuevo_patron_job",
                    )

                    patron_final = patron_elegido
                    if "Personalizado" in patron_elegido:
                        patron_final = st.text_input(
                            "Describe el patrón personalizado:",
                            value=patron_actual if patron_actual not in patrones_opts else "",
                            placeholder="Ej. Diario (Solo lunes, miércoles y viernes)",
                            key="txt_patron_custom_job",
                        ).strip() or patron_elegido

                    nota_mod = st.text_input(
                        "Nota / Motivo del cambio (opcional):",
                        value=ov_actual.get("notas", "") if ov_actual else "",
                        placeholder="Ej. Confirmado con Operaciones TI que corre diario",
                        key="txt_nota_mod_job",
                    )

                    c_b1, c_b2 = st.columns(2)
                    with c_b1:
                        if st.button(
                            "💾 Guardar clasificación",
                            type="primary",
                            use_container_width=True,
                            key="btn_guardar_clas_job",
                        ):
                            guardar_clasificacion_manual(job_sel_mod, nueva_cat, patron_final, nota_mod)
                            refrescar_app(
                                f"Clasificación de '{job_sel_mod}' actualizada a '{nueva_cat}' ({patron_final})."
                            )
                    with c_b2:
                        if st.button(
                            "↩️ Restaurar automática",
                            disabled=ov_actual is None,
                            use_container_width=True,
                            key="btn_restaurar_clas_job",
                        ):
                            eliminar_clasificacion_manual([job_sel_mod])
                            refrescar_app(
                                f"Se restauró la clasificación automática para '{job_sel_mod}'."
                            )

                    with st.expander("🔧 Editar nombre (JOBNAME) o Grupo en Supabase"):
                        st.caption("Modifica directamente los registros de este proceso en `public.ejecuciones_procesos`.")
                        nuevo_nombre_job = st.text_input(
                            "Nuevo nombre (JOBNAME):",
                            value=job_sel_mod,
                            key="txt_edit_jobname",
                        )
                        grupo_actual_str = str(fila_job["grupo"]).split(",")[0].strip()
                        nuevo_grupo_job = st.text_input(
                            "Nuevo Grupo:",
                            value=grupo_actual_str,
                            key="txt_edit_grupo_job",
                        )
                        if st.button(
                            "✏️ Aplicar cambio en BD",
                            use_container_width=True,
                            key="btn_aplicar_edit_bd",
                        ):
                            if not nuevo_nombre_job.strip():
                                st.error("El nombre del proceso no puede quedar vacío.")
                            else:
                                with st.spinner("Actualizando registros en Supabase..."):
                                    n_upd = editar_metadatos_proceso_supabase(
                                        job_sel_mod, nuevo_nombre_job, nuevo_grupo_job
                                    )
                                refrescar_app(
                                    f"Se actualizaron {n_upd} registros del proceso '{job_sel_mod}' en Supabase."
                                )

        # --- SUBSECCIÓN 3B: ELIMINAR PROCESOS ---
        with col_del_proc:
            with st.container(border=True):
                st.markdown("### 🗑️ Eliminar procesos")
                st.caption(
                    "Elimina uno o varios procesos por completo de Supabase (o solo en meses específicos) "
                    "para recalcular KPIs, gráficos y auditoría."
                )

                col_fd1, col_fd2 = st.columns(2)
                with col_fd1:
                    filtro_grp_del = st.selectbox(
                        "Filtrar por Grupo (opcional):",
                        options=["Todos"] + grupos_disp,
                        key="sel_grp_del_job",
                    )
                with col_fd2:
                    busq_del = st.text_input(
                        "Buscar JOBNAME:",
                        placeholder="Ej. DPAHD10",
                        key="txt_busq_del_job",
                    )

                df_jobs_del_vista = df_jobs_gest.copy()
                if filtro_grp_del != "Todos":
                    df_jobs_del_vista = df_jobs_del_vista[
                        df_jobs_del_vista["grupo"].str.contains(re.escape(filtro_grp_del), na=False)
                    ]
                if busq_del and busq_del.strip():
                    df_jobs_del_vista = df_jobs_del_vista[
                        df_jobs_del_vista["jobname"].str.contains(busq_del.strip(), case=False, na=False)
                    ]

                mapa_etiq_jobs = {
                    row["jobname"]: (
                        f"{row['jobname']} | {row['grupo']} "
                        f"({int(row['registros'])} reg. / {int(row['dias_ejecutados'])} días / "
                        f"{formato_seg_hhmmss(row['seg_totales'])})"
                    )
                    for _, row in df_jobs_del_vista.iterrows()
                }

                jobs_sel_eliminar = st.multiselect(
                    "Seleccionar proceso(s) a eliminar:",
                    options=df_jobs_del_vista["jobname"].tolist(),
                    format_func=lambda j: mapa_etiq_jobs[j],
                    placeholder="Elige uno o más JOBNAME...",
                    key="ms_jobs_eliminar",
                )

                alcance_del = st.radio(
                    "Alcance de la eliminación:",
                    options=["Todo el histórico (eliminar por completo)", "Solo en meses específicos"],
                    horizontal=True,
                    key="radio_alcance_del_job",
                )

                meses_del_job = None
                if alcance_del == "Solo en meses específicos":
                    meses_del_job = st.multiselect(
                        "Selecciona el/los mes(es) donde eliminar esos procesos:",
                        options=meses_disp,
                        placeholder="Elige uno o más meses...",
                        key="ms_meses_del_job",
                    )

                if jobs_sel_eliminar:
                    df_sel_j = df_jobs_gest[df_jobs_gest["jobname"].isin(jobs_sel_eliminar)]
                    total_reg_j = int(df_sel_j["registros"].sum())

                    if alcance_del.startswith("Todo"):
                        st.warning(
                            f"Vas a eliminar **por completo {len(jobs_sel_eliminar)} proceso(s)** "
                            f"({total_reg_j} registros históricos en total). Esta acción no se puede deshacer."
                        )
                        puede_borrar = True
                    else:
                        if meses_del_job:
                            st.warning(
                                f"Vas a eliminar **{len(jobs_sel_eliminar)} proceso(s)** únicamente en: "
                                f"**{', '.join(meses_del_job)}**. Esta acción no se puede deshacer."
                            )
                            puede_borrar = True
                        else:
                            st.info("Selecciona al menos un mes para continuar.")
                            puede_borrar = False

                    confirmar_del_jobs = st.checkbox(
                        "Confirmo que deseo eliminar permanentemente los procesos seleccionados",
                        key="chk_confirmar_borrar_jobs",
                    )

                    if st.button(
                        "🗑️ Eliminar proceso(s) seleccionado(s)",
                        type="primary",
                        disabled=(not confirmar_del_jobs or not puede_borrar),
                        use_container_width=True,
                        key="btn_borrar_jobs_supa",
                    ):
                        with st.spinner("Eliminando procesos en Supabase..."):
                            n_elim = eliminar_procesos_supabase(
                                jobs_sel_eliminar,
                                meses_a_borrar=meses_del_job if not alcance_del.startswith("Todo") else None,
                            )
                        refrescar_app(
                            f"Se eliminaron {n_elim} registros de {len(jobs_sel_eliminar)} proceso(s). "
                            "Gráficos y auditoría actualizados."
                        )

        # --- LISTADO DE CLASIFICACIONES MANUALES ACTIVAS ---
        if overrides_manuales:
            with st.expander(f"📋 Ver procesos con clasificación manual ({len(overrides_manuales)})"):
                df_ov = pd.DataFrame([
                    {
                        "Proceso (JOBNAME)": j,
                        "Categoría Manual": v.get("categoria_manual", ""),
                        "Patrón / Excepciones": v.get("patron_manual", ""),
                        "Notas": v.get("notas", ""),
                    }
                    for j, v in sorted(overrides_manuales.items())
                ])
                st.dataframe(df_ov, use_container_width=True, hide_index=True)

                ov_sel_restaurar = st.multiselect(
                    "Restaurar clasificación automática para:",
                    options=sorted(overrides_manuales.keys()),
                    placeholder="Selecciona uno o más procesos...",
                    key="ms_restaurar_ov",
                )
                if ov_sel_restaurar:
                    if st.button(
                        "↩️ Restaurar seleccionados a detección automática",
                        key="btn_restaurar_ov_lote",
                    ):
                        eliminar_clasificacion_manual(ov_sel_restaurar)
                        refrescar_app(
                            f"Se restauró la detección automática para {len(ov_sel_restaurar)} proceso(s)."
                        )