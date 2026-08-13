import io
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Catastro de Vehículos 2027",
    page_icon="🚗",
    layout="wide",
)

ARCHIVO_DEFAULT = "catastro_vehiculos_transformado.xlsx"
HOJA_DEFAULT = "Catastro_final"
CRITERIOS = [
    "Criterio antigüedad >= 8 años",
    "Criterio priorización alta",
    "Criterio gasto mantención > $5.000.000",
    "Criterio kilometraje > 100.000",
]


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().lower()


def cumple_criterio(valor):
    texto = normalizar_texto(valor)
    return texto in ["cumple", "si cumple", "sí cumple"]


@st.cache_data(show_spinner=False)
def leer_excel(archivo):
    """Lee la hoja Catastro_final. Si no existe, usa la primera hoja disponible."""
    xls = pd.ExcelFile(archivo, engine="openpyxl")
    hoja = HOJA_DEFAULT if HOJA_DEFAULT in xls.sheet_names else xls.sheet_names[0]
    data = pd.read_excel(archivo, sheet_name=hoja, engine="openpyxl")
    data.columns = [str(c).strip() for c in data.columns]
    return data, hoja


def preparar_datos(df):
    df = df.copy()

    for col in ["Kilometraje acumulado", "Gasto en mantención acumulada", "Año Vehículo"]:
        if col in df.columns:
            df[col + " num"] = pd.to_numeric(df[col], errors="coerce")

    if "Año Vehículo num" in df.columns:
        df["Antigüedad 2027"] = 2027 - df["Año Vehículo num"]

    for criterio in CRITERIOS:
        if criterio not in df.columns:
            df[criterio] = "no cumple"

    df["Cantidad criterios cumplidos"] = df[CRITERIOS].apply(
        lambda fila: sum(cumple_criterio(v) for v in fila), axis=1
    )
    df["Cumple 3 o más criterios"] = df["Cantidad criterios cumplidos"].apply(
        lambda x: "cumple" if x >= 3 else "no cumple"
    )
    df["Cumple antigüedad y priorización alta"] = df.apply(
        lambda r: "cumple"
        if cumple_criterio(r.get("Criterio antigüedad >= 8 años"))
        and cumple_criterio(r.get("Criterio priorización alta"))
        else "no cumple",
        axis=1,
    )
    return df


def aplicar_multiselect(df, columna, etiqueta):
    if columna not in df.columns:
        return df
    opciones = sorted([str(x) for x in df[columna].dropna().unique()])
    seleccion = st.sidebar.multiselect(etiqueta, opciones)
    if seleccion:
        return df[df[columna].astype(str).isin(seleccion)]
    return df


def dataframe_a_excel(df):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reporte_filtrado")
    salida.seek(0)
    return salida


def generar_minuta(df, filtros_aplicados):
    total = len(df)
    servicios = df["Servicio"].nunique() if "Servicio" in df.columns else 0
    altas = 0
    if "Priorización" in df.columns:
        altas = df["Priorización"].astype(str).str.lower().str.contains("alta", na=False).sum()

    tres_mas = df["Cumple 3 o más criterios"].eq("cumple").sum() if "Cumple 3 o más criterios" in df.columns else 0
    gasto_total = df.get("Gasto en mantención acumulada num", pd.Series(dtype=float)).sum(skipna=True)
    km_promedio = df.get("Kilometraje acumulado num", pd.Series(dtype=float)).mean(skipna=True)

    lineas = []
    lineas.append("MINUTA CATÁSTRO DE VEHÍCULOS")
    lineas.append(f"Fecha de emisión: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    lineas.append("")
    lineas.append("1. Filtros aplicados")
    lineas.extend([f"- {f}" for f in filtros_aplicados] if filtros_aplicados else ["- Sin filtros aplicados"])
    lineas.append("")
    lineas.append("2. Resumen general")
    lineas.append(f"- Total de vehículos considerados: {total}")
    lineas.append(f"- Servicios considerados: {servicios}")
    lineas.append(f"- Vehículos con priorización alta: {altas}")
    lineas.append(f"- Vehículos que cumplen 3 o más criterios: {tres_mas}")
    lineas.append(f"- Gasto total en mantención acumulada: ${gasto_total:,.0f}".replace(",", "."))
    if pd.notna(km_promedio):
        lineas.append(f"- Kilometraje promedio: {km_promedio:,.0f} km".replace(",", "."))
    else:
        lineas.append("- Kilometraje promedio: No disponible")
    lineas.append("")
    lineas.append("3. Vehículos que cumplen 3 o más criterios")

    cols_base = [c for c in ["Servicio", "I.R.N.V.M.", "Tipo vehículo", "Año Vehículo", "Priorización", "Cantidad criterios cumplidos"] if c in df.columns]
    candidatos = df[df["Cantidad criterios cumplidos"] >= 3].copy() if "Cantidad criterios cumplidos" in df.columns else pd.DataFrame()
    if candidatos.empty:
        lineas.append("- No se identifican vehículos que cumplan 3 o más criterios en el universo filtrado.")
    else:
        for _, row in candidatos[cols_base].head(50).iterrows():
            detalle = " | ".join([f"{c}: {row.get(c, '')}" for c in cols_base])
            lineas.append(f"- {detalle}")
        if len(candidatos) > 50:
            lineas.append(f"- Se omiten {len(candidatos) - 50} registros adicionales por extensión de la minuta.")

    lineas.append("")
    lineas.append("4. Observaciones")
    if "Observaciones calidad de datos" in df.columns:
        obs = df["Observaciones calidad de datos"].dropna().astype(str)
        obs = obs[~obs.str.lower().eq("sin observaciones")]
        if obs.empty:
            lineas.append("- No se observan alertas relevantes de calidad de datos en los registros filtrados.")
        else:
            for item in obs.value_counts().head(10).items():
                lineas.append(f"- {item[0]}: {item[1]} registro(s)")
    else:
        lineas.append("- La base no contiene columna de observaciones de calidad de datos.")

    return "\n".join(lineas)


st.title("🚗 Catastro de Vehículos 2027")
st.caption("Aplicación para análisis, filtros, gráficos y generación de minuta del catastro de vehículos.")

with st.sidebar:
    st.header("Carga de datos")
    archivo_subido = st.file_uploader("Cargar archivo Excel", type=["xlsx"])

try:
    archivo = archivo_subido if archivo_subido is not None else ARCHIVO_DEFAULT
    df_original, hoja_usada = leer_excel(archivo)
except Exception as e:
    st.error("No fue posible cargar el archivo Excel. Verifique que el archivo esté en la misma carpeta de app.py o cárguelo desde la barra lateral.")
    st.exception(e)
    st.stop()

st.info(f"Hoja utilizada: {hoja_usada} | Registros cargados: {len(df_original):,}".replace(",", "."))
df = preparar_datos(df_original)

st.sidebar.header("Filtros")
df_filtrado = df.copy()
filtros_aplicados = []

for columna, etiqueta in [
    ("Servicio", "Servicio"),
    ("Tipo vehículo", "Tipo de vehículo"),
    ("Estado", "Estado"),
    ("Priorización", "Priorización"),
    ("Criterio antigüedad >= 8 años", "Criterio antigüedad"),
    ("Criterio priorización alta", "Criterio priorización alta"),
    ("Criterio gasto mantención > $5.000.000", "Criterio gasto mantención"),
    ("Criterio kilometraje > 100.000", "Criterio kilometraje"),
]:
    antes = len(df_filtrado)
    df_filtrado = aplicar_multiselect(df_filtrado, columna, etiqueta)
    if len(df_filtrado) != antes:
        filtros_aplicados.append(f"{etiqueta}: filtro aplicado")

st.subheader("Indicadores principales")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Vehículos", f"{len(df_filtrado):,}".replace(",", "."))
col2.metric("Servicios", df_filtrado["Servicio"].nunique() if "Servicio" in df_filtrado.columns else 0)
altas = df_filtrado["Priorización"].astype(str).str.lower().str.contains("alta", na=False).sum() if "Priorización" in df_filtrado.columns else 0
col3.metric("Priorización alta", f"{altas:,}".replace(",", "."))
tres = df_filtrado["Cumple 3 o más criterios"].eq("cumple").sum()
col4.metric("3 o más criterios", f"{tres:,}".replace(",", "."))
gasto_total = df_filtrado.get("Gasto en mantención acumulada num", pd.Series(dtype=float)).sum(skipna=True)
col5.metric("Gasto mantención", "$" + f"{gasto_total:,.0f}".replace(",", "."))

st.subheader("Gráficos")
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Vehículos por Servicio",
    "Mayor gasto",
    "Antigüedad flota",
    "3 o más criterios",
    "Antigüedad + alta",
])

with tab1:
    if "Servicio" in df_filtrado.columns:
        graf = df_filtrado.groupby("Servicio").size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        st.plotly_chart(px.bar(graf, x="Servicio", y="Cantidad", title="Cantidad de vehículos por Servicio"), use_container_width=True)

with tab2:
    if "Gasto en mantención acumulada num" in df_filtrado.columns:
        cols = [c for c in ["Servicio", "I.R.N.V.M.", "Tipo vehículo", "Gasto en mantención acumulada num"] if c in df_filtrado.columns]
        top = df_filtrado[cols].dropna(subset=["Gasto en mantención acumulada num"]).sort_values("Gasto en mantención acumulada num", ascending=False).head(20)
        st.plotly_chart(px.bar(top, x="I.R.N.V.M.", y="Gasto en mantención acumulada num", color="Servicio", title="Top 20 vehículos con mayor gasto en mantención"), use_container_width=True)

with tab3:
    if "Antigüedad 2027" in df_filtrado.columns:
        antig = df_filtrado.dropna(subset=["Antigüedad 2027"]).copy()
        st.plotly_chart(px.histogram(antig, x="Antigüedad 2027", nbins=20, title="Distribución de antigüedad de la flota al año 2027"), use_container_width=True)

with tab4:
    graf = df_filtrado.groupby("Cumple 3 o más criterios").size().reset_index(name="Cantidad")
    st.plotly_chart(px.bar(graf, x="Cumple 3 o más criterios", y="Cantidad", title="Vehículos que cumplen con 3 o más criterios"), use_container_width=True)

with tab5:
    graf = df_filtrado.groupby("Cumple antigüedad y priorización alta").size().reset_index(name="Cantidad")
    st.plotly_chart(px.bar(graf, x="Cumple antigüedad y priorización alta", y="Cantidad", title="Vehículos que cumplen antigüedad y priorización alta"), use_container_width=True)

st.subheader("Tabla de resultados")
st.dataframe(df_filtrado, use_container_width=True, height=420)

st.subheader("Descargas")
col_a, col_b = st.columns(2)
with col_a:
    st.download_button(
        label="Descargar Excel filtrado",
        data=dataframe_a_excel(df_filtrado),
        file_name="reporte_catastro_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_b:
    minuta = generar_minuta(df_filtrado, filtros_aplicados)
    st.download_button(
        label="Descargar minuta TXT",
        data=minuta.encode("utf-8"),
        file_name="minuta_catastro_vehiculos.txt",
        mime="text/plain",
    )

with st.expander("Ver minuta generada"):
    st.text(minuta)
