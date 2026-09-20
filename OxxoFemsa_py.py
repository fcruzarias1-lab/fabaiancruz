import glob
import io
import os

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Tablero Maestro Barcel",
    layout="wide",
    page_icon="",
)

# -----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# -----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_CARPETA = os.path.join(BASE_DIR, "OxxoFemsa_txt")
ARCHIVO_MAESTRO = os.path.join(BASE_DIR, "BaseOxxoFemsa.xlsx")

ORDEN_ESTATUS = [
    "Activo Saludable",
    "Riesgo de Quiebre (<= 3d)",
    "Alerta: Planograma",
    "Quiebre: Venta sin Stock (OOS)",
    "Sobreinventario (> 25d)",
    "Inactivo / Sin Presencia",
]

COLORES_ESTATUS = {
    "Activo Saludable": "#2b8a3e",
    "Riesgo de Quiebre (<= 3d)": "#f59f00",
    "Alerta: Planograma": "#fab005",
    "Quiebre: Venta sin Stock (OOS)": "#c92a2a",
    "Sobreinventario (> 25d)": "#d9480f",
    "Inactivo / Sin Presencia": "#868e96",
}

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    .kpi-box {
        background-color: #ffffff;
        padding: 12px 14px;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.07);
        text-align: left;
        border-top: 4px solid #eaeaea;
        min-height: 104px;
    }
    .kpi-box h5 {
        margin: 0;
        color: #666;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
    }
    .kpi-box p {
        margin: 3px 0;
        color: #0056b3;
        font-size: 9px;
        font-weight: 700;
    }
    .kpi-box h2 {
        margin: 0;
        font-size: 24px;
        font-weight: 800;
        color: #111;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("️ Tablero In Stock Barcel | OXXO")
st.markdown("### Inteligencia de Datos Diaria")


# -----------------------------------------------------------------------------
# FUNCIONES
# -----------------------------------------------------------------------------
def normalizar_clave(serie):
    """Normaliza claves para evitar fallas de cruce por espacios o decimales."""
    return (
        serie.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


@st.cache_data(show_spinner=False)
def cargar_base_tiendas(ruta_archivo):
    if not os.path.exists(ruta_archivo):
        return None

    try:
        libro = pd.ExcelFile(ruta_archivo, engine="openpyxl")
        df_m = None
        hoja_maestra = None
        for hoja in libro.sheet_names:
            encabezados = pd.read_excel(
                ruta_archivo, sheet_name=hoja, nrows=0, engine="openpyxl"
            )
            columnas = [str(c).strip() for c in encabezados.columns]
            if any(c.upper() == "CR_TIENDA" for c in columnas):
                hoja_maestra = hoja
                df_m = pd.read_excel(
                    ruta_archivo, sheet_name=hoja, dtype=str, engine="openpyxl"
                )
                break

        if df_m is None:
            st.sidebar.error(
                "La base maestra no contiene una hoja con la columna CR_TIENDA."
            )
            return None

        df_m.columns = [str(c).strip() for c in df_m.columns]
        if df_m.empty:
            st.sidebar.error(f"La hoja maestra '{hoja_maestra}' está vacía.")
            return None

        col_t = next(c for c in df_m.columns if c.upper() == "CR_TIENDA")
        df_m = df_m.rename(columns={col_t: "CR_TIENDA"})
        df_m["CR_TIENDA"] = normalizar_clave(df_m["CR_TIENDA"])

        object_cols = df_m.select_dtypes(include="object").columns
        df_m[object_cols] = df_m[object_cols].apply(
            lambda col: col.fillna("").astype(str).str.strip()
        )

        df_m = df_m.drop_duplicates(subset=["CR_TIENDA"], keep="first")
        return df_m
    except Exception as e:
        st.sidebar.error(f"Error al leer el Excel maestro: {e}")
        return None


@st.cache_data(show_spinner=False)
def cargar_datos_diarios(ruta, df_m):
    try:
        df = pd.read_csv(ruta, sep="|", encoding="utf-8", low_memory=False)
    except UnicodeDecodeError:
        df = pd.read_csv(ruta, sep="|", encoding="latin-1", low_memory=False)

    df.columns = [str(c).strip() for c in df.columns]

    if df.empty or len(df.columns) == 0:
        return pd.DataFrame()

    candidatos_tienda = [
        c for c in df.columns if "TIENDA" in c.upper() and "CR" in c.upper()
    ]
    if not candidatos_tienda:
        st.error(
            "El TXT no contiene la columna CR_TIENDA. "
            f"Columnas detectadas: {', '.join(df.columns[:12])}"
        )
        return pd.DataFrame()
    df = df.rename(columns={candidatos_tienda[0]: "CR_TIENDA"})
    df["CR_TIENDA"] = normalizar_clave(df["CR_TIENDA"])

    columnas_numericas = [
        "STOCK_ON_HAND",
        "UNIDADES_VENDIDAS",
        "UNIDADES_VENDIDAS_4W",
        "PROMEDIO_SEMANAL_4W",
        "DI",
    ]
    for col in columnas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        else:
            df[col] = 0.0

    if "PLANOGRAMADO" in df.columns:
        valores_si = {"SI", "SÍ", "1", "TRUE", "VERDADERO", "YES"}
        df["PLANOGRAMADO"] = (
            df["PLANOGRAMADO"]
            .fillna("NO")
            .astype(str)
            .str.strip()
            .str.upper()
            .map(lambda valor: "SI" if valor in valores_si else "NO")
        )
    else:
        df["PLANOGRAMADO"] = "NO"

    if "ITEM_DESC" in df.columns:
        df["ITEM_DESC"] = (
            df["ITEM_DESC"]
            .fillna("DESCONOCIDO")
            .astype(str)
            .str.strip()
        )
    else:
        df["ITEM_DESC"] = "DESCONOCIDO"

    condiciones = [
        (df["STOCK_ON_HAND"] > 0) & (df["DI"] > 25),
        (df["STOCK_ON_HAND"] > 0) & (df["DI"] <= 3),
        (df["STOCK_ON_HAND"] > 0),
        (df["STOCK_ON_HAND"] <= 0) & (df["UNIDADES_VENDIDAS"] > 0),
        (df["STOCK_ON_HAND"] <= 0) & (df["PLANOGRAMADO"] == "SI"),
    ]
    resultados = [
        "Sobreinventario (> 25d)",
        "Riesgo de Quiebre (<= 3d)",
        "Activo Saludable",
        "Quiebre: Venta sin Stock (OOS)",
        "Alerta: Planograma",
    ]
    df["ESTATUS_OPERATIVO"] = np.select(
        condiciones,
        resultados,
        default="Inactivo / Sin Presencia",
    )

    mapa_acciones = {
        "Activo Saludable": "Mantener nivel y monitorear",
        "Riesgo de Quiebre (<= 3d)": "Priorizar resurtido preventivo",
        "Alerta: Planograma": "Validar exhibición, alta y surtido",
        "Quiebre: Venta sin Stock (OOS)": "Resurtir con prioridad alta",
        "Sobreinventario (> 25d)": "Revisar pedido y redistribución",
        "Inactivo / Sin Presencia": "Validar presencia y catálogo",
    }
    df["ACCION_RECOMENDADA"] = df["ESTATUS_OPERATIVO"].map(mapa_acciones)

    if df_m is not None:
        columnas_maestro = [
            c for c in df_m.columns if c == "CR_TIENDA" or c not in df.columns
        ]
        df = pd.merge(
            df,
            df_m[columnas_maestro],
            on="CR_TIENDA",
            how="left",
            validate="many_to_one",
        )
        columnas_texto = df.select_dtypes(include="object").columns
        df[columnas_texto] = df[columnas_texto].fillna("No Mapeado en Excel")

    return df


@st.cache_data(show_spinner=False)
def generar_excel(df_exportar):
    salida = io.BytesIO()
    try:
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            df_exportar.to_excel(writer, index=False, sheet_name="Detalle")

            resumen = (
                df_exportar.groupby("ESTATUS_OPERATIVO", dropna=False)
                .agg(
                    REGISTROS=("CR_TIENDA", "size"),
                    TIENDAS=("CR_TIENDA", "nunique"),
                    INVENTARIO=("STOCK_ON_HAND", "sum"),
                )
                .reset_index()
            )
            resumen.to_excel(writer, index=False, sheet_name="Resumen")

            if "PLAZA" in df_exportar.columns:
                ranking_export = construir_ranking_plaza(df_exportar)
                if not ranking_export.empty:
                    ranking_export.to_excel(
                        writer, index=False, sheet_name="Ranking"
                    )

        salida.seek(0)
        return salida.getvalue()
    except Exception as e:
        st.error(f"Error exportando Excel: {e}")
        return None


def construir_ranking_plaza(df_base):
    if "PLAZA" not in df_base.columns or df_base.empty:
        return pd.DataFrame()

    ranking = (
        df_base.groupby("PLAZA", dropna=False)
        .agg(
            REGISTROS=("CR_TIENDA", "size"),
            TIENDAS=("CR_TIENDA", "nunique"),
            SALUDABLES=(
                "ESTATUS_OPERATIVO",
                lambda s: (s == "Activo Saludable").sum(),
            ),
            OOS=(
                "ESTATUS_OPERATIVO",
                lambda s: (s == "Quiebre: Venta sin Stock (OOS)").sum(),
            ),
            INVENTARIO=("STOCK_ON_HAND", "sum"),
        )
        .reset_index()
    )
    ranking["% IN STOCK"] = np.where(
        ranking["REGISTROS"] > 0,
        ranking["SALUDABLES"] / ranking["REGISTROS"] * 100,
        0,
    ).round(1)
    ranking["% OOS"] = np.where(
        ranking["REGISTROS"] > 0,
        ranking["OOS"] / ranking["REGISTROS"] * 100,
        0,
    ).round(1)
    return ranking.sort_values(
        ["% IN STOCK", "% OOS"], ascending=[False, True]
    ).reset_index(drop=True)


def tarjeta_kpi(contenedor, titulo, subtitulo, valor, color):
    contenedor.markdown(
        f"""
        <div class="kpi-box" style="border-top-color:{color};">
            <h5>{titulo}</h5>
            <p>{subtitulo}</p>
            <h2>{valor}</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# ARCHIVOS Y CARGA
# -----------------------------------------------------------------------------
os.makedirs(RUTA_CARPETA, exist_ok=True)
archivos = sorted(
    glob.glob(os.path.join(RUTA_CARPETA, "*.txt")),
    key=os.path.getmtime,
    reverse=True,
)

if not archivos:
    st.info(
        f"📂 Coloca tus archivos .txt dentro de: `{os.path.abspath(RUTA_CARPETA)}`"
    )
    st.stop()

archivo_sel = st.sidebar.selectbox(
    "🗓️ HISTORIAL DIARIO:",
    archivos,
    format_func=os.path.basename,
)

if st.sidebar.button("🔄 Actualizar archivos y caché", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

df_maestro = cargar_base_tiendas(ARCHIVO_MAESTRO)
if df_maestro is not None:
    st.sidebar.success("Base interna vinculada.")
else:
    st.sidebar.warning("⚠️ Excel maestro no encontrado o no válido.")

df_completo = cargar_datos_diarios(archivo_sel, df_maestro)
if df_completo.empty:
    st.error("El archivo seleccionado no contiene registros utilizables.")
    st.stop()


# -----------------------------------------------------------------------------
# FILTROS
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("🎛️ FILTROS APLICADOS")
df_filtrado = df_completo.copy()

filtros_sidebar = [
    ("GERENCIA", "Gerencia"),
    ("CEVES", "Ceves"),
    ("REGION", "Región OXXO"),
    ("PLAZA", "Plaza OXXO"),
]

for numero, (col, label) in enumerate(filtros_sidebar, start=1):
    if col in df_filtrado.columns:
        opciones = sorted(
            [
                str(x)
                for x in df_filtrado[col].dropna().unique()
                if str(x) not in {"", "N/A", "No Mapeado en Excel"}
            ]
        )

        seleccion = st.sidebar.multiselect(
            f"{numero}. {label}:",
            options=opciones,
            default=opciones,
        )

        if seleccion:
            df_filtrado = df_filtrado[
                df_filtrado[col].astype(str).isin(seleccion)
            ]
        else:
            df_filtrado = df_filtrado.iloc[0:0]

productos = sorted(
    {
        str(p)
        for p in df_filtrado["ITEM_DESC"].dropna().unique()
        if str(p) not in {"", "N/A"}
    }
)
sel_productos = st.sidebar.multiselect(
    "3. Seleccionar productos:",
    options=productos,
    default=productos,
)
if sel_productos:
    df_filtrado = df_filtrado[df_filtrado["ITEM_DESC"].isin(sel_productos)]
else:
    df_filtrado = df_filtrado.iloc[0:0]

estatus_presentes = [
    estatus
    for estatus in ORDEN_ESTATUS
    if estatus in df_filtrado["ESTATUS_OPERATIVO"].unique()
]
sel_estatus = st.sidebar.multiselect(
    "4. Estatus de inventario:",
    options=estatus_presentes,
    default=estatus_presentes,
)
if sel_estatus:
    df_filtrado = df_filtrado[
        df_filtrado["ESTATUS_OPERATIVO"].isin(sel_estatus)
    ]
else:
    df_filtrado = df_filtrado.iloc[0:0]

st.sidebar.caption(f"Archivo: {os.path.basename(archivo_sel)}")
st.sidebar.caption(f"Registros visibles: {len(df_filtrado):,}")


# -----------------------------------------------------------------------------
# CONTENIDO PRINCIPAL
# -----------------------------------------------------------------------------
if df_filtrado.empty:
    st.warning("No hay registros para los filtros seleccionados.")
    st.stop()

registros = len(df_filtrado)
registros_saludables = (
    df_filtrado["ESTATUS_OPERATIVO"] == "Activo Saludable"
).sum()
registros_oos = (
    df_filtrado["ESTATUS_OPERATIVO"] == "Quiebre: Venta sin Stock (OOS)"
).sum()

in_stock_pct = registros_saludables / registros * 100 if registros else 0
oos_pct = registros_oos / registros * 100 if registros else 0
pog_si = df_filtrado.loc[
    df_filtrado["PLANOGRAMADO"] == "SI", "CR_TIENDA"
].nunique()
activas = df_filtrado["CR_TIENDA"].nunique()
soh_total = df_filtrado["STOCK_ON_HAND"].sum()
promedio_4w = df_filtrado["PROMEDIO_SEMANAL_4W"].sum()

k1, k2, k3, k4, k5, k6 = st.columns(6)
tarjeta_kpi(k1, "IN STOCK", "REGISTROS SALUDABLES", f"{in_stock_pct:.1f}%", "#2b8a3e")
tarjeta_kpi(k2, "OOS", "VENTA SIN STOCK", f"{oos_pct:.1f}%", "#c92a2a")
tarjeta_kpi(k3, "PLANOGRAMADAS", "TIENDAS POG: SI", f"{pog_si:,}", "#7950f2")
tarjeta_kpi(k4, "TIENDAS ACTIVAS", "TOTAL VISTA", f"{activas:,}", "#6f42c1")
tarjeta_kpi(k5, "INVENTARIO", "PIEZAS SOH", f"{int(soh_total):,}", "#0056b3")
tarjeta_kpi(k6, "PROM. SEMANAL 4W", "DESPLAZAMIENTO", f"{int(promedio_4w):,}", "#e67e22")

st.markdown("---")

col_exportar, col_contexto = st.columns([1, 3])
with col_exportar:
    excel_data = generar_excel(df_filtrado)
    if excel_data:
        st.download_button(
            label="📥 Exportar Reporte Excel",
            data=excel_data,
            file_name=f"Reporte_InStock_Barcel_{os.path.splitext(os.path.basename(archivo_sel))[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.error("Error al generar el archivo de exportación.")
with col_contexto:
    st.caption(
        "El Excel descargable incluye el detalle filtrado, el resumen general por estatus y el ranking ejecutivo por plaza."
    )

# -----------------------------------------------------------------------------
# PESTAÑAS
# -----------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "🌋 Matriz detallada",
        "📦 Auditoría por producto",
        "🌀 Rendimiento SOH/4W",
        "🔥 Salud operativa",
        "🏆 Ranking de plazas",
        "🎯 Acciones y quiebres",
    ]
)

with tab1:
    st.subheader("🌋 Matriz operativa de cierre detallado")
    columnas_detalle = [
        c
        for c in [
            "CR_TIENDA",
            "CEVES ID",
            "CEVES",
            "DIVISION",
            "GERENCIA",
            "CLIENTE NOMBRE",
            "CODIGO GB",
            "REGION",
            "PLAZA",
            "RUTA_ID",
            "ITEM_DESC",
            "STOCK_ON_HAND",
            "UNIDADES_VENDIDAS",
            "UNIDADES_VENDIDAS_4W",
            "PROMEDIO_SEMANAL_4W",
            "PLANOGRAMADO",
            "DI",
            "ESTATUS_OPERATIVO",
            "ACCION_RECOMENDADA",
        ]
        if c in df_filtrado.columns
    ]

    orden_detalle = [c for c in ["CR_TIENDA", "ITEM_DESC"] if c in columnas_detalle]
    tabla_detalle = df_filtrado[columnas_detalle]
    if orden_detalle:
        tabla_detalle = tabla_detalle.sort_values(orden_detalle)

    st.dataframe(
        tabla_detalle,
        use_container_width=True,
        height=500,
        hide_index=True,
        column_config={
            "DI": st.column_config.NumberColumn("DI", format="%.1f"),
            "STOCK_ON_HAND": st.column_config.NumberColumn("SOH", format="%.0f"),
            "PROMEDIO_SEMANAL_4W": st.column_config.NumberColumn(
                "Prom. semanal 4W", format="%.1f"
            ),
        },
    )

with tab2:
    c_sku1, c_sku2 = st.columns([5, 5])

    with c_sku1:
        st.subheader("📊 Planogramación por SKU")
        df_pog_sku = (
            df_filtrado.groupby(["ITEM_DESC", "PLANOGRAMADO"])["CR_TIENDA"]
            .nunique()
            .unstack(fill_value=0)
            .reset_index()
        )
        for col in ["SI", "NO"]:
            if col not in df_pog_sku.columns:
                df_pog_sku[col] = 0

        df_pog_sku["ACTIVAS"] = df_pog_sku["SI"] + df_pog_sku["NO"]
        df_pog_sku["% POG"] = np.where(
            df_pog_sku["ACTIVAS"] > 0,
            df_pog_sku["SI"] / df_pog_sku["ACTIVAS"] * 100,
            0,
        ).round(1)
        df_pog_sku = df_pog_sku.rename(
            columns={
                "ITEM_DESC": "PRODUCTO",
                "SI": "POG (SI)",
                "NO": "NO POG (NO)",
            }
        ).sort_values("ACTIVAS", ascending=False)

        st.dataframe(
            df_pog_sku,
            use_container_width=True,
            height=440,
            hide_index=True,
            column_config={
                "% POG": st.column_config.ProgressColumn(
                    "% POG",
                    min_value=0,
                    max_value=100,
                    format="%.1f%%",
                )
            },
        )

    with c_sku2:
        st.subheader("🚨 Top 10 alertas de surtido por SKU")
        estados_alerta = [
            "Alerta: Planograma",
            "Quiebre: Venta sin Stock (OOS)",
            "Riesgo de Quiebre (<= 3d)",
        ]
        df_alertas = df_filtrado[
            df_filtrado["ESTATUS_OPERATIVO"].isin(estados_alerta)
        ]

        if not df_alertas.empty:
            resumen_alertas = (
                df_alertas.groupby(["ITEM_DESC", "ESTATUS_OPERATIVO"])["CR_TIENDA"]
                .nunique()
                .reset_index(name="TIENDAS_AFECTADAS")
            )
            top_skus = (
                resumen_alertas.groupby("ITEM_DESC")["TIENDAS_AFECTADAS"]
                .sum()
                .nlargest(10)
                .index
            )
            resumen_alertas = resumen_alertas[
                resumen_alertas["ITEM_DESC"].isin(top_skus)
            ]

            fig_alertas = px.bar(
                resumen_alertas,
                x="TIENDAS_AFECTADAS",
                y="ITEM_DESC",
                color="ESTATUS_OPERATIVO",
                orientation="h",
                barmode="stack",
                color_discrete_map=COLORES_ESTATUS,
                labels={"TIENDAS_AFECTADAS": "Tiendas afectadas", "ITEM_DESC": ""},
            )
            fig_alertas.update_layout(
                yaxis={"categoryorder": "total ascending"},
                legend=dict(orientation="h", y=-0.25),
                margin=dict(l=10, r=10, t=20, b=80),
            )
            st.plotly_chart(fig_alertas, use_container_width=True)
        else:
            st.success("🎉 Sin alertas de surtido para los filtros actuales.")

with tab3:
    g1, g2 = st.columns([6, 4])

    with g1:
        st.subheader("📈 Venta promedio vs inventario")
        columnas_hover = [
            c
            for c in ["CR_TIENDA", "ITEM_DESC", "PLAZA", "DI", "PLANOGRAMADO"]
            if c in df_filtrado.columns
        ]
        fig_scatter = px.scatter(
            df_filtrado,
            x="PROMEDIO_SEMANAL_4W",
            y="STOCK_ON_HAND",
            color="ESTATUS_OPERATIVO",
            hover_data=columnas_hover,
            color_discrete_map=COLORES_ESTATUS,
            opacity=0.72,
            labels={
                "PROMEDIO_SEMANAL_4W": "Promedio semanal 4W",
                "STOCK_ON_HAND": "Inventario SOH",
                "ESTATUS_OPERATIVO": "Estatus",
            },
        )
        fig_scatter.update_traces(marker=dict(size=9, line=dict(width=0.4, color="white")))
        fig_scatter.update_layout(legend=dict(orientation="h", y=-0.25))
        st.plotly_chart(fig_scatter, use_container_width=True)

    with g2:
        st.subheader("🏆 Top 10 tiendas con sobreinventario")
        df_sobre = df_filtrado[
            df_filtrado["ESTATUS_OPERATIVO"] == "Sobreinventario (> 25d)"
        ]
        if not df_sobre.empty:
            top_tiendas = (
                df_sobre.groupby("CR_TIENDA")["STOCK_ON_HAND"]
                .sum()
                .nlargest(10)
                .reset_index()
            )
            fig_top = px.bar(
                top_tiendas,
                x="STOCK_ON_HAND",
                y="CR_TIENDA",
                orientation="h",
                color_discrete_sequence=["#d9480f"],
                labels={"STOCK_ON_HAND": "Inventario SOH", "CR_TIENDA": "Tienda"},
            )
            fig_top.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig_top, use_container_width=True)
        else:
            st.success("Sin tiendas clasificadas con sobreinventario.")

with tab4:
    p1, p2 = st.columns(2)

    with p1:
        st.subheader("⭕ Salud del stock")
        salud = (
            df_filtrado["ESTATUS_OPERATIVO"]
            .value_counts()
            .rename_axis("ESTATUS_OPERATIVO")
            .reset_index(name="REGISTROS")
        )
        fig_donut = px.pie(
            salud,
            values="REGISTROS",
            names="ESTATUS_OPERATIVO",
            hole=0.55,
            color="ESTATUS_OPERATIVO",
            color_discrete_map=COLORES_ESTATUS,
        )
        fig_donut.update_traces(textposition="inside", textinfo="percent")
        fig_donut.update_layout(legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig_donut, use_container_width=True)

    with p2:
        st.subheader("▲ Planogramación por plaza")
        if "PLAZA" in df_filtrado.columns:
            df_plaza_pog = (
                df_filtrado.groupby(["PLAZA", "PLANOGRAMADO"])["CR_TIENDA"]
                .nunique()
                .reset_index(name="TIENDAS")
            )
            fig_plaza = px.bar(
                df_plaza_pog,
                x="PLAZA",
                y="TIENDAS",
                color="PLANOGRAMADO",
                barmode="group",
                color_discrete_map={"SI": "#2b8a3e", "NO": "#c92a2a"},
                labels={"PLAZA": "Plaza", "TIENDAS": "Tiendas", "PLANOGRAMADO": "POG"},
            )
            fig_plaza.update_layout(xaxis_tickangle=-35)
            st.plotly_chart(fig_plaza, use_container_width=True)
        else:
            st.info("La columna PLAZA no está disponible en los datos.")

with tab5:
    tipo_ranking = st.radio(
        "Ver ranking por:",
        ["PLAZA", "CEVES"],
        horizontal=True
    )

    if tipo_ranking == "PLAZA":
        ranking_plaza = construir_ranking_plaza(df_filtrado)
        columna_ranking = "PLAZA"
        titulo = "🏆 Ranking ejecutivo por plaza"
    else:
        if "CEVES" in df_filtrado.columns:
            ranking_plaza = (
                df_filtrado.groupby("CEVES", dropna=False)
                .agg(
                    REGISTROS=("CR_TIENDA", "size"),
                    TIENDAS=("CR_TIENDA", "nunique"),
                    SALUDABLES=(
                        "ESTATUS_OPERATIVO",
                        lambda s: (s == "Activo Saludable").sum(),
                    ),
                    OOS=(
                        "ESTATUS_OPERATIVO",
                        lambda s: (s == "Quiebre: Venta sin Stock (OOS)").sum(),
                    ),
                    INVENTARIO=("STOCK_ON_HAND", "sum"),
                )
                .reset_index()
            )
            ranking_plaza["% IN STOCK"] = np.where(
                ranking_plaza["REGISTROS"] > 0,
                ranking_plaza["SALUDABLES"] / ranking_plaza["REGISTROS"] * 100,
                0,
            ).round(1)
            ranking_plaza["% OOS"] = np.where(
                ranking_plaza["REGISTROS"] > 0,
                ranking_plaza["OOS"] / ranking_plaza["REGISTROS"] * 100,
                0,
            ).round(1)
            ranking_plaza = ranking_plaza.sort_values(
                ["% IN STOCK", "% OOS"], ascending=[False, True]
            ).reset_index(drop=True)
        else:
            ranking_plaza = pd.DataFrame()

        columna_ranking = "CEVES"
        titulo = "🏭 Ranking ejecutivo por CEVES"

    st.subheader(titulo)
    if not ranking_plaza.empty:
        ranking_mostrar = ranking_plaza.copy()
        ranking_mostrar.insert(0, "POSICIÓN", range(1, len(ranking_mostrar) + 1))
        st.dataframe(
            ranking_mostrar,
            use_container_width=True,
            hide_index=True,
            height=470,
            column_config={
                "% IN STOCK": st.column_config.ProgressColumn(
                    "% IN STOCK", min_value=0, max_value=100, format="%.1f%%"
                ),
                "% OOS": st.column_config.NumberColumn("% OOS", format="%.1f%%"),
                "INVENTARIO": st.column_config.NumberColumn("INVENTARIO", format="%.0f"),
            },
        )

        fig_ranking = px.bar(
            ranking_plaza.sort_values("% IN STOCK", ascending=True),
            x="% IN STOCK",
            y=columna_ranking,
            orientation="h",
            color="% IN STOCK",
            color_continuous_scale="RdYlGn",
            range_color=[0, 100],
            text="% IN STOCK",
            labels={"% IN STOCK": "In Stock (%)", columna_ranking: ""},
        )
        fig_ranking.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig_ranking.update_layout(coloraxis_showscale=False, xaxis_range=[0, 105])
        st.plotly_chart(fig_ranking, use_container_width=True)
    else:
        st.info(f"La columna {columna_ranking} no está disponible para construir el ranking.")

with tab6:
    a1, a2 = st.columns([6, 4])

    with a1:
        st.subheader("🎯 Acciones comerciales recomendadas")
        columnas_accion = [
            c
            for c in [
                "GERENCIA",
                "CEVES",
                "CODIGO GB",
                "RUTA_ID",
                "CR_TIENDA",
                "CLIENTE NOMBRE",
                "REGION",
                "PLAZA",
                "ITEM_DESC",
                "STOCK_ON_HAND",
                "UNIDADES_VENDIDAS",
                "PROMEDIO_SEMANAL_4W",
                "DI",
                "ESTATUS_OPERATIVO",
                "ACCION_RECOMENDADA",
            ]
            if c in df_filtrado.columns
        ]
        prioridades = {
            "Quiebre: Venta sin Stock (OOS)": 1,
            "Riesgo de Quiebre (<= 3d)": 2,
            "Alerta: Planograma": 3,
            "Sobreinventario (> 25d)": 4,
            "Inactivo / Sin Presencia": 5,
            "Activo Saludable": 6,
        }
        acciones = df_filtrado[columnas_accion].copy()
        acciones["PRIORIDAD"] = acciones["ESTATUS_OPERATIVO"].map(prioridades)
        acciones = acciones.sort_values(
            ["PRIORIDAD", "UNIDADES_VENDIDAS", "PROMEDIO_SEMANAL_4W"],
            ascending=[True, False, False],
        ).drop(columns="PRIORIDAD")

        st.dataframe(
            acciones,
            use_container_width=True,
            hide_index=True,
            height=500,
        )

    with a2:
        st.subheader("🚨 Quiebres prioritarios")
        quiebres = df_filtrado[
            df_filtrado["ESTATUS_OPERATIVO"]
            == "Quiebre: Venta sin Stock (OOS)"
        ].copy()

        if not quiebres.empty:
            top_quiebres = (
                quiebres.groupby("ITEM_DESC")
                .agg(
                    TIENDAS_AFECTADAS=("CR_TIENDA", "nunique"),
                    UNIDADES_VENDIDAS=("UNIDADES_VENDIDAS", "sum"),
                    PROMEDIO_4W=("PROMEDIO_SEMANAL_4W", "sum"),
                )
                .reset_index()
                .sort_values(
                    ["UNIDADES_VENDIDAS", "PROMEDIO_4W"],
                    ascending=False,
                )
                .head(10)
            )
            st.dataframe(
                top_quiebres,
                use_container_width=True,
                hide_index=True,
            )

            fig_quiebres = px.bar(
                top_quiebres.sort_values("TIENDAS_AFECTADAS"),
                x="TIENDAS_AFECTADAS",
                y="ITEM_DESC",
                orientation="h",
                color_discrete_sequence=["#c92a2a"],
                labels={"TIENDAS_AFECTADAS": "Tiendas", "ITEM_DESC": ""},
            )
            st.plotly_chart(fig_quiebres, use_container_width=True)
        else:
            st.success("🎉 No hay quiebres OOS para los filtros actuales.")