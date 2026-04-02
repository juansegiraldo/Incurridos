import io
import re
import unicodedata
import pandas as pd
import streamlit as st


def leer_datos(archivo, header=0):
    if hasattr(archivo, "seek"):
        archivo.seek(0)
    return pd.read_excel(archivo, header=header)


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def texto_equivalente(valor):
    txt = normalizar_texto(valor).upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    return txt


def clave_columna(valor):
    txt = texto_equivalente(valor)
    txt = re.sub(r"[^A-Z0-9]+", "", txt)
    return txt


def tiene_alguna_columna(df, candidatos):
    cols = {clave_columna(c) for c in df.columns}
    return any(clave_columna(c) in cols for c in candidatos)


def leer_datos_auto_header(archivo, columnas_esperadas=None):
    intentos = [0, 1]
    mejor_df = None
    mejor_score = -1
    columnas_esperadas = columnas_esperadas or []

    for header in intentos:
        try:
            df = leer_datos(archivo, header=header)
            if not columnas_esperadas:
                return df
            score = 0
            for grupo in columnas_esperadas:
                if tiene_alguna_columna(df, grupo):
                    score += 1
            if score > mejor_score:
                mejor_score = score
                mejor_df = df
            if score == len(columnas_esperadas):
                return df
        except Exception:
            continue
    if mejor_df is not None:
        return mejor_df
    return leer_datos(archivo, header=0)


def resolver_columna(df, candidatos, label):
    columnas = {clave_columna(col): col for col in df.columns}
    for candidato in candidatos:
        real = columnas.get(clave_columna(candidato))
        if real is not None:
            return real
    st.error(
        f"No encontre la columna para '{label}'. Columnas disponibles: {', '.join([str(c) for c in df.columns])}"
    )
    st.stop()


def resolver_columna_opcional(df, candidatos):
    columnas = {clave_columna(col): col for col in df.columns}
    for candidato in candidatos:
        real = columnas.get(clave_columna(candidato))
        if real is not None:
            return real
    return None


def to_numeric_safe(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0.0)


def preparar_factores(df_filtrado, col_persona, fx_cop_default, fx_gbp_default):
    personas = (
        df_filtrado[col_persona]
        .fillna("SIN_NOMBRE")
        .astype(str)
        .str.strip()
        .replace("", "SIN_NOMBRE")
        .unique()
    )
    personas = sorted(personas)
    base = pd.DataFrame(
        {"Persona": personas, "Excluir": False, "Factor": 1.0, "Moneda": "EUR", "Tasa a EUR": 1.0}
    )
    editor = st.data_editor(
        base,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "Persona": st.column_config.TextColumn("Persona", disabled=True),
            "Excluir": st.column_config.CheckboxColumn("Excluir"),
            "Factor": st.column_config.NumberColumn("Factor", min_value=0.0, max_value=1.0, step=0.05),
            "Moneda": st.column_config.SelectboxColumn("Moneda", options=["EUR", "COP", "GBP"]),
            "Tasa a EUR": st.column_config.NumberColumn("Tasa a EUR", min_value=0.000001, step=0.0001),
        },
    )
    editor["FactorFinal"] = editor.apply(
        lambda row: 0.0 if bool(row["Excluir"]) else float(row["Factor"]),
        axis=1,
    )
    return {
        fila["Persona"]: {
            "factor": float(fila["FactorFinal"]),
            "fx": (
                fx_cop_default
                if fila["Moneda"] == "COP" and float(fila["Tasa a EUR"]) == 1.0
                else (
                    fx_gbp_default
                    if fila["Moneda"] == "GBP" and float(fila["Tasa a EUR"]) == 1.0
                    else (float(fila["Tasa a EUR"]) if float(fila["Tasa a EUR"]) > 0 else 1.0)
                )
            ),
            "moneda": fila["Moneda"],
        }
        for _, fila in editor.iterrows()
    }


def aplicar_ajustes(df_filtrado, col_persona, col_jornadas, col_coste, factores):
    df_adj = df_filtrado.copy()
    df_adj[col_persona] = (
        df_adj[col_persona].fillna("SIN_NOMBRE").astype(str).str.strip().replace("", "SIN_NOMBRE")
    )
    df_adj["FactorPersona"] = df_adj[col_persona].map(
        lambda p: factores.get(p, {}).get("factor", 1.0)
    )
    df_adj["FxPersona"] = df_adj[col_persona].map(lambda p: factores.get(p, {}).get("fx", 1.0))
    df_adj["JornadasAdj"] = to_numeric_safe(df_adj[col_jornadas]) * df_adj["FactorPersona"]
    df_adj["CosteAdj"] = (
        to_numeric_safe(df_adj[col_coste]) * df_adj["FxPersona"] * df_adj["FactorPersona"]
    )
    return df_adj


def tabla_final_por_proyecto(df_adj, col_project_name):
    final = (
        df_adj.groupby(col_project_name, dropna=False)[["CosteAdj", "JornadasAdj"]]
        .sum()
        .reset_index()
        .rename(
            columns={
                col_project_name: "PROJECT NAME",
                "CosteAdj": "YTD costs",
                "JornadasAdj": "YTD dates",
            }
        )
        .sort_values("PROJECT NAME")
    )
    final["YTD costs"] = final["YTD costs"].round(2)
    final["YTD dates"] = final["YTD dates"].round(2)
    return final


def descargar_excel(df, nombre_archivo):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="YTD")
    salida.seek(0)
    st.download_button(
        "Descargar Excel salida",
        data=salida,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


st.set_page_config(page_title="Resumen Incurridos", layout="wide")
st.title("Resumen de Incurridos para Controlling")

st.sidebar.header("Archivos")
archivo_incurridos = st.sidebar.file_uploader(
    "Incurridos (obligatorio)", type=["xlsx", "xls"], key="incurridos"
)
archivo_controlling = st.sidebar.file_uploader(
    "Controlling (opcional)", type=["xlsx", "xls"], key="controlling"
)

if archivo_incurridos is None:
    st.info("Carga primero el archivo de Incurridos para empezar.")
    st.stop()

try:
    df_incurridos = leer_datos_auto_header(
        archivo_incurridos,
        columnas_esperadas=[["Proyecto", "PROJECT NAME"], ["Nombre Completo"], ["Jornadas"], ["CosteEUR", "Coste EUR"]],
    )
except Exception as exc:
    st.error(f"No pude leer Incurridos: {exc}")
    st.stop()

col_project_name = resolver_columna(df_incurridos, ["PROJECT NAME", "Proyecto"], "PROJECT NAME")
col_persona = resolver_columna(df_incurridos, ["Nombre Completo", "Empleado", "Persona"], "Persona")
col_jornadas = resolver_columna(df_incurridos, ["Jornadas"], "Jornadas")
col_coste = resolver_columna(df_incurridos, ["CosteEUR", "Coste EUR", "Coste"], "CosteEUR")
col_fecha = resolver_columna_opcional(df_incurridos, ["FechaImputacion", "Fecha Imputacion", "Fecha", "Date"])

df_work = df_incurridos.copy()
df_work[col_project_name] = df_work[col_project_name].apply(normalizar_texto)

st.success(f"Incurridos cargado. Registros: {len(df_work)}")

st.subheader("Filtros")
jp_default = "JUAN SEBASTIAN GIRALDO"

df_jp = df_work.copy()
control_projects_jp = None

if archivo_controlling is not None:
    try:
        df_controlling = leer_datos_auto_header(
            archivo_controlling,
            columnas_esperadas=[["PROJECT NAME", "Proyecto"], ["JP Responsable", "JP"]],
        )
        col_ctrl_project = resolver_columna(df_controlling, ["PROJECT NAME", "Proyecto"], "PROJECT NAME (Controlling)")
        col_ctrl_jp = resolver_columna(df_controlling, ["JP Responsable", "JP"], "JP Responsable (Controlling)")
        df_controlling[col_ctrl_project] = df_controlling[col_ctrl_project].apply(normalizar_texto)
        df_controlling[col_ctrl_jp] = df_controlling[col_ctrl_jp].apply(normalizar_texto)

        jps_ctrl = sorted([x for x in df_controlling[col_ctrl_jp].dropna().unique() if str(x).strip() != ""])
        if jps_ctrl:
            jp_match = next((x for x in jps_ctrl if texto_equivalente(x) == texto_equivalente(jp_default)), None)
            jp_seleccionado = st.selectbox(
                "JP Responsable (desde controlling)",
                options=jps_ctrl,
                index=jps_ctrl.index(jp_match) if jp_match in jps_ctrl else 0,
            )
            control_projects_jp = set(
                df_controlling[df_controlling[col_ctrl_jp] == jp_seleccionado][col_ctrl_project]
                .dropna()
                .astype(str)
                .str.strip()
            )
            df_jp = df_work[df_work[col_project_name].isin(control_projects_jp)].copy()
            st.caption(f"Proyectos habilitados por JP en controlling: {len(control_projects_jp)}")
        else:
            st.warning("El controlling no trae valores de JP Responsable. Se mostraran todos los proyectos de Incurridos.")
    except Exception as exc:
        st.error(f"No pude procesar controlling para filtro JP: {exc}")
        st.stop()
else:
    st.info("Sin controlling: se muestran proyectos desde Incurridos sin filtro por JP Responsable.")

proyectos = sorted([x for x in df_jp[col_project_name].dropna().unique() if str(x).strip() != ""])

col_a, col_b = st.columns([1, 4])
with col_a:
    seleccionar_todos = st.checkbox("Seleccionar todos", value=True)
with col_b:
    proyectos_sel = st.multiselect(
        "PROJECT NAME (multiseleccion con busqueda)",
        options=proyectos,
        default=proyectos if seleccionar_todos else [],
    )

if not proyectos_sel:
    st.warning("Selecciona al menos un proyecto.")
    st.stop()

df_filtrado = df_jp[df_jp[col_project_name].isin(proyectos_sel)].copy()
st.caption(f"Registros filtrados: {len(df_filtrado)}")

st.subheader("Ajustes por persona")
st.caption("Puedes excluir personas completas o aplicar factor fraccional (0 a 1).")
fx_col1, fx_col2 = st.columns(2)
with fx_col1:
    fx_cop_default = st.number_input("Tasa COP a EUR (default)", min_value=0.0, value=0.00022, step=0.00001, format="%.6f")
with fx_col2:
    fx_gbp_default = st.number_input("Tasa GBP a EUR (default)", min_value=0.0, value=1.17, step=0.01, format="%.4f")
st.caption("Si en una persona dejas 'Tasa a EUR' en 1.0 y moneda COP/GBP, se aplican estos defaults.")
factores = preparar_factores(df_filtrado, col_persona, fx_cop_default, fx_gbp_default)
df_ajustado = aplicar_ajustes(df_filtrado, col_persona, col_jornadas, col_coste, factores)

st.subheader("Resumen rapido")
c1, c2 = st.columns(2)
with c1:
    st.metric("YTD costs total", f"{df_ajustado['CosteAdj'].sum():,.2f}")
with c2:
    st.metric("YTD dates total", f"{df_ajustado['JornadasAdj'].sum():,.2f}")

st.subheader("Validacion de consistencia")
coste_base = to_numeric_safe(df_filtrado[col_coste]).sum()
jornadas_base = to_numeric_safe(df_filtrado[col_jornadas]).sum()
coste_ajustado = df_ajustado["CosteAdj"].sum()
jornadas_ajustadas = df_ajustado["JornadasAdj"].sum()
st.dataframe(
    pd.DataFrame(
        {
            "Metrica": ["Coste", "Jornadas"],
            "Base": [round(coste_base, 2), round(jornadas_base, 2)],
            "Ajustado": [round(coste_ajustado, 2), round(jornadas_ajustadas, 2)],
            "Diferencia": [round(coste_ajustado - coste_base, 2), round(jornadas_ajustadas - jornadas_base, 2)],
        }
    ),
    use_container_width=True,
)

st.subheader("Salida final para controlling")
tabla_final = tabla_final_por_proyecto(df_ajustado, col_project_name)
st.dataframe(tabla_final, use_container_width=True)
st.download_button(
    "Descargar CSV salida",
    data=tabla_final.to_csv(index=False).encode("utf-8"),
    file_name="ytd_para_controlling.csv",
    mime="text/csv",
)
descargar_excel(tabla_final, "ytd_para_controlling.xlsx")

if archivo_controlling is not None:
    try:
        df_controlling = leer_datos_auto_header(
            archivo_controlling,
            columnas_esperadas=[["PROJECT NAME", "Proyecto"]],
        )
        col_ctrl_project = resolver_columna(df_controlling, ["PROJECT NAME", "Proyecto"], "PROJECT NAME (Controlling)")
        control_projects = set(df_controlling[col_ctrl_project].apply(normalizar_texto))
        salida_projects = set(tabla_final["PROJECT NAME"].apply(normalizar_texto))
        sin_cruce = sorted([x for x in salida_projects if x and x not in control_projects])

        st.subheader("Validacion de cruce con controlling")
        if sin_cruce:
            st.warning("Hay proyectos en salida que no existen en controlling.")
            st.dataframe(pd.DataFrame({"PROJECT NAME sin cruce": sin_cruce}), use_container_width=True)
        else:
            st.success("Todos los PROJECT NAME de la salida existen en controlling.")
    except Exception as exc:
        st.error(f"No pude validar el archivo de controlling: {exc}")

st.subheader("Analisis adicional")
tab_detalle, tab_hist = st.tabs(["Detalle por proyecto", "Histogramas de jornadas por dia"])

with tab_detalle:
    proyecto_detalle = st.selectbox(
        "Proyecto para ver detalle",
        options=proyectos_sel,
        index=0,
        key="detalle_proyecto",
    )
    df_detalle = df_ajustado[df_ajustado[col_project_name] == proyecto_detalle].copy()
    cols_detalle = [col_project_name, col_persona]
    if col_fecha is not None:
        cols_detalle.append(col_fecha)
    cols_detalle.extend([col_jornadas, "JornadasAdj", col_coste, "CosteAdj", "FactorPersona", "FxPersona"])
    cols_detalle = [c for c in cols_detalle if c in df_detalle.columns]
    st.dataframe(df_detalle[cols_detalle], use_container_width=True)

    st.markdown("**Resumen por persona (proyecto seleccionado)**")
    resumen_persona = (
        df_detalle.groupby(col_persona, dropna=False)
        .agg(
            JornadasBase=(col_jornadas, lambda s: to_numeric_safe(s).sum()),
            JornadasAdj=("JornadasAdj", "sum"),
            CosteBase=(col_coste, lambda s: to_numeric_safe(s).sum()),
            CosteAdj=("CosteAdj", "sum"),
        )
        .reset_index()
        .sort_values("CosteAdj", ascending=False)
    )
    for c in ["JornadasBase", "JornadasAdj", "CosteBase", "CosteAdj"]:
        resumen_persona[c] = resumen_persona[c].round(2)
    st.dataframe(resumen_persona, use_container_width=True)

with tab_hist:
    if col_fecha is None:
        st.info("No encontre columna de fecha en Incurridos para construir histogramas por dia.")
    else:
        proyectos_hist = st.multiselect(
            "Proyectos para histograma diario",
            options=proyectos_sel,
            default=proyectos_sel[:3] if len(proyectos_sel) > 3 else proyectos_sel,
            key="hist_proyectos",
        )
        if not proyectos_hist:
            st.warning("Selecciona al menos un proyecto para ver histogramas.")
        else:
            for proyecto in proyectos_hist:
                df_hist = df_ajustado[df_ajustado[col_project_name] == proyecto].copy()
                df_hist["FechaDia"] = pd.to_datetime(df_hist[col_fecha], errors="coerce").dt.date
                serie = (
                    df_hist.dropna(subset=["FechaDia"])
                    .groupby("FechaDia")["JornadasAdj"]
                    .sum()
                    .reset_index()
                    .sort_values("FechaDia")
                )
                st.markdown(f"**{proyecto}**")
                if serie.empty:
                    st.caption("Sin fechas validas para este proyecto.")
                else:
                    st.bar_chart(serie.set_index("FechaDia")["JornadasAdj"])