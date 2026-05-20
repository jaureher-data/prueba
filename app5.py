import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from datetime import date
from dateutil.relativedelta import relativedelta


from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
from io import BytesIO
import hashlib

import time

from st_aggrid import AgGrid
from st_aggrid import GridOptionsBuilder

from app3 import cargar_data








# ============================================================
# APP: Sistema inteligente de inventarios
# Autor: Portafolio Streamlit
# Descripción: Dashboard para analizar inventario, rotación,
# clasificación ABC, cobertura, punto de reorden y alertas.
# ============================================================

st.set_page_config(
    page_title="Sistema Inteligente de Inventarios",
    page_icon="📦",
    layout="wide"
)



# ------------------------------------------------------------
# AUTENTICACIÓN
# ------------------------------------------------------------
# Usuarios de demostración.
# En un proyecto real, estas credenciales deben ir en una base de datos
# o en .streamlit/secrets.toml, nunca directamente en el código.
USERS = {
    "admin": {
        "name": "Administrador",
        "password_hash": hashlib.sha256("admin123".encode()).hexdigest(),
        "role": "admin"
    },
    "cliente": {
        "name": "Cliente Demo",
        "password_hash": hashlib.sha256("cliente123".encode()).hexdigest(),
        "role": "cliente"
    }
}


def check_password(username, password):
    user = USERS.get(username)
    if not user:
        return False
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    return password_hash == user["password_hash"]


def login_screen():
    st.markdown('<div class="main-title">🔐 Acceso al Sistema</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitle">Ingrese sus credenciales para acceder al Sistema Inteligente de Inventarios.</div>',
        unsafe_allow_html=True
    )

    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Ingresar")

        if submitted:
            if check_password(username, password):
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.name = USERS[username]["name"]
                st.session_state.role = USERS[username]["role"]
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")

    st.info("Usuarios demo: admin / admin123  |  cliente / cliente123")


def logout():
    for key in ["logged_in", "username", "name", "role"]:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()


# ------------------------------------------------------------
# ESTILOS
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    .main-title {
        font-size: 34px;
        font-weight: 800;
        color: #1f2937;
    }
    .subtitle {
        font-size: 17px;
        color: #4b5563;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #f9fafb;
        border-radius: 14px;
        padding: 18px;
        border: 1px solid #e5e7eb;
    }
    .warning-box {
        background-color: #fff7ed;
        border-left: 5px solid #f97316;
        padding: 14px;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ------------------------------------------------------------
# CONTROL DE SESIÓN
# ------------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    login_screen()
    st.stop()

# ------------------------------------------------------------
# FUNCIONES DE APOYO
# ------------------------------------------------------------
@st.cache_data
def load_csv(file):
    return pd.read_csv(file)


def normalize_columns(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )
    return df


def find_column(df, possible_names):
    for col in df.columns:
        if col in possible_names:
            return col
    return None


def calculate_abc(df, value_col):
    df = df.sort_values(value_col, ascending=False).copy()
    total = df[value_col].sum()
    if total == 0:
        df["participacion_%"] = 0
        df["participacion_acumulada_%"] = 0
        df["abc"] = "C"
        return df

    df["participacion_%"] = df[value_col] / total * 100
    df["participacion_acumulada_%"] = df["participacion_%"].cumsum()

    conditions = [
        df["participacion_acumulada_%"] <= 80,
        df["participacion_acumulada_%"] <= 95
    ]
    choices = ["A", "B"]
    df["abc"] = np.select(conditions, choices, default="C")
    return df


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Inventario Analizado")
    return output.getvalue()


# ------------------------------------------------------------
# ENCABEZADO
# ------------------------------------------------------------
header_col1, header_col2 = st.columns([4, 1])

with header_col1:
    st.markdown('<div class="main-title">📦 Sistema Inteligente de Inventarios</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitle">App de analítica para optimizar inventario, detectar riesgos de quiebre, calcular cobertura, rotación, clasificación ABC y punto de reorden.</div>',
        unsafe_allow_html=True
    )

with header_col2:
    st.write(f"👤 {st.session_state.name}")
    st.caption(f"Rol: {st.session_state.role}")
    if st.button("Cerrar sesión"):
        logout()

# ------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------
st.sidebar.title("📌 Menú")
pagina = st.sidebar.radio(
    "Seleccione una página:",
    [
        "📊 Dashboard",
        "📦 Inventario",
        "🚨 Alertas",
        "📈 Forecasting"
    ]
)


st.sidebar.title("⚙️ Configuración")
st.sidebar.success(f"Sesión activa: {st.session_state.name}")
st.sidebar.caption(f"Rol: {st.session_state.role}")
st.sidebar.markdown("Carga un archivo CSV con información de inventario y ventas.")

if st.session_state.role == "admin":
    uploaded_file = st.sidebar.file_uploader("Subir archivo CSV", type=["csv"])
else:
    uploaded_file = None
    st.sidebar.info("El usuario cliente visualiza datos demo. Solo el administrador puede cargar archivos.")

lead_time_days = st.sidebar.number_input(
    "Lead time en días",
    min_value=1,
    max_value=120,
    value=5,
    step=1
)

safety_stock_days = st.sidebar.number_input(
    "Stock de seguridad en días",
    min_value=0,
    max_value=120,
    value=7,
    step=1
)

st.sidebar.markdown("---")
st.sidebar.markdown("### Columnas esperadas")
st.sidebar.markdown(
    """
    El archivo debe incluir columnas similares a:
    - producto
    - categoria
    - stock
    - ventas
    - precio
    - costo
    """
)

# ------------------------------------------------------------
# DATOS DE EJEMPLO
# ------------------------------------------------------------
# ============================================================
#Credencial de google
# ============================================================

# ============================================================
#Cargar los datos una sóla vez 
# ============================================================

#@st.cache_data
#def cargar_datos():
    #exec(open("app3V2.py").read())
#    return cargar_data()

if "df" not in st.session_state:
    st.session_state.df,st.session_state.dfventas,st.session_state.dfcompra = cargar_data()

df = st.session_state.df
dfventas=st.session_state.dfventas
dfcompra=st.session_state.dfcompra


SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES
)
service = build("drive", "v3", credentials=creds)



#st.write("Nombres de las columnas:", df.columns.tolist())



#if uploaded_file is None:
#    st.info("Sube un archivo CSV o usa los datos de ejemplo para explorar la app.")
#
#    sample_data = pd.DataFrame({
#        "producto": [
#            "Perfume A", "Perfume B", "Perfume C", "Perfume D", "Perfume E",
#            "Perfume F", "Perfume G", "Perfume H", "Perfume I", "Perfume J"
#        ],
#        "categoria": [
#            "Premium", "Premium", "Media", "Media", "Económica",
#            "Premium", "Media", "Económica", "Media", "Premium"
#        ],
#        "stock": [120, 40, 300, 25, 500, 12, 180, 700, 65, 18],
#        "ventas": [80, 70, 40, 55, 20, 35, 45, 15, 60, 50],
#        "precio": [250000, 220000, 120000, 135000, 70000, 300000, 110000, 60000, 150000, 280000],
#        "costo": [150000, 130000, 75000, 85000, 40000, 190000, 70000, 35000, 95000, 170000]
#    })
#    df = sample_data.copy()
#else:
#    df = load_csv(uploaded_file)

# ------------------------------------------------------------
# PROCESAMIENTO
# ------------------------------------------------------------
df = normalize_columns(df)

product_col = find_column(df, ["sku","producto", "product", "referencia", "item"])
category_col = find_column(df, ["categoria", "category", "linea", "tipo"])
stock_col = find_column(df, ["stock", "inventario", "existencias", "cantidad"])
sales_col = find_column(df, ["ventas", "sales", "unidades_vendidas", "demanda"])
price_col = find_column(df, ["precio", "price", "valor_unitario", "venta_unitaria"])
cost_col = find_column(df, ["costo", "cost", "costo_unitario"])

required = {
    "producto": product_col,
    "stock": stock_col,
    "ventas": sales_col,
    "precio": price_col
}

missing = [name for name, col in required.items() if col is None]

if missing:
    st.error(f"Faltan columnas necesarias: {', '.join(missing)}")
    st.stop()

if category_col is None:
    df["categoria"] = "Sin categoría"
    category_col = "categoria"

if cost_col is None:
    df["costo"] = 0
    cost_col = "costo"

# Convertir columnas numéricas
for col in [stock_col, sales_col, price_col, cost_col]:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# Métricas base
df["valor_inventario"] = df[stock_col] * df[cost_col]
df["valor_ventas"] = df[sales_col] * df[price_col]
df["margen_unitario"] = df[price_col] - df[cost_col]
df["margen_total"] = df["margen_unitario"] * df[sales_col]

# Supuesto: ventas son mensuales
# Venta diaria promedio = ventas mensuales / 30
df["venta_diaria_promedio"] = df[sales_col] / 30

df["cobertura_dias"] = np.where(
    df["venta_diaria_promedio"] > 0,
    df[stock_col] / df["venta_diaria_promedio"],
    np.inf
)

df["cobertura_meses"] = df["cobertura_dias"] / 30

df["punto_reorden"] = (
    df["venta_diaria_promedio"] * lead_time_days
) + (
    df["venta_diaria_promedio"] * safety_stock_days
)

df["stock_seguridad"] = df["venta_diaria_promedio"] * safety_stock_days

df["rotacion_mensual"] = np.where(
    df[stock_col] > 0,
    df[sales_col] / df[stock_col],
    0
)

df["estado_inventario"] = np.select(
    [
        df[stock_col] <= df["punto_reorden"],
        df["cobertura_meses"] > 6,
        df["rotacion_mensual"] < 0.1
    ],
    [
        "Riesgo de quiebre",
        "Sobrestock",
        "Baja rotación"
    ],
    default="Saludable"
)

# ABC con base en valor de ventas
df = calculate_abc(df, "valor_ventas")

# ------------------------------------------------------------
# FILTROS Pagina 1
# ------------------------------------------------------------
st.markdown("## 🔎 Filtros")
col_f1, col_f2= st.columns([2,9])

with col_f1:
    selected_status = st.multiselect(
        "Estado del inventario",
        options=sorted(df["inventory_status"].unique()),
        default=sorted(df["inventory_status"].unique())
    )
    selected_abc = st.multiselect(
        "Clasificación ABC",
        options=["A", "B", "C"],
        default=["A", "B", "C"]
    )


filtered = df[
    #(df[category_col].isin(selected_categories)) &
    (df["abc_class"].isin(selected_abc)) &
    (df["inventory_status"].isin(selected_status))
].copy()    

with col_f2:
    gb = GridOptionsBuilder.from_dataframe(filtered)

    #gb.configure_selection("single")
    gb.configure_selection("multiple")

    gridOptions = gb.build()

    response = AgGrid(
        filtered,
        gridOptions=gridOptions
    )


#Esta parte del codigo se utiliza para resetear los botones


if "ejecBoton_01" not in st.session_state:
    st.session_state.ejecBoton_01 = False
if "ejecBoton_02" not in st.session_state:
    st.session_state.ejecBoton_02 = False


selected = response["selected_rows"]

if selected is None or len(selected)==0:
    st.write("Ninguna fila seleccionada")
else:
    col1a, col2a, col3a = st.columns([1.5,7,1.5])
    with col1a:
        st.write("Columna 1")
        #if st.button("Eliminar de",disabled=st.session_state.ejecBoton_01):
        if st.button("Eliminar de"):
            st.write("Presion  1")
            st.session_state.ejecBoton_01 = True
            

            if st.session_state.ejecBoton_01:
                
                FILE_ID = "1YB68NRBQK0Diqhj3B1vxoEtsIUxtdm07"
                # Simulación de proceso
                request = service.files().get_media(fileId=FILE_ID)
                file = io.BytesIO()
                downloader = MediaIoBaseDownload(file, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                file.seek(0)
                dfeliminados = pd.read_csv(file)

                st.dataframe(dfeliminados)
                




                # Resetear el botón
                st.session_state.ejecBoton_01 = False
                st.success("Proceso 1 finalizado")

                # Recargar interfaz
                st.rerun()

    with col2a:
        st.write(selected)
    with col3a:
        st.write("Columna 2")
        if st.button("comprar"):
            st.write("Presion  2")
            st.session_state.ejecBoton_02 = True
            if st.session_state.ejecBoton_01:
                st.info("Ejecutando proceso 2")

                # Simulación de proceso
                time.sleep(3)

                st.success("Proceso 2 finalizado")

                # Resetear el botón
                st.session_state.ejecBoton_02 = False

                # Recargar interfaz
                st.rerun()
    
    
# ============================================================
# PÁGINA 1: DASHBOARD EJECUTIVO
# ============================================================

if pagina == "📊 Dashboard":
    st.title("📊 Dashboard ejecutivo")
    st.write("Resumen general del estado del inventario y desempeño comercial.")

    # KPIs principales
    col1, col2, col3, col4 = st.columns(4)

    if selected is not None:
        #st.write("Se debe seleccionar fila")
        skus = selected["sku"].tolist()
        filtered = filtered[filtered["sku"].isin(skus)]

    with col1:
        st.metric(
            "📦 Productos",
            f"{len(filtered):,.0f}"
        )

    with col2:
        st.metric(
            "💰 Valor inventario",
            f"${filtered['valor_inventario'].sum():,.0f}"
        )

    with col3:
        st.metric(
            "📈 Ventas mensuales",
            f"${filtered['valor_ventas'].sum():,.0f}"
        )

    with col4:
        productos_riesgo = (
            filtered["estado_inventario"] == "Riesgo de quiebre"
        ).sum()

        st.metric(
            "🚨 Riesgo de quiebre",
            f"{productos_riesgo:,.0f}"
        )

    st.markdown("---")


    # Gráficos principales 1
    col_g1, col_g2 = st.columns(2)



    if selected is not None:
        
        with col_g1:
            ventas_summary = dfventas[dfventas["SKU"].isin(skus)]

            fig_ventas = px.bar(
                ventas_summary,
                x="Mes",
                y="total_quantity",
                color="SKU",
                title="📊 Valor de ventas",
                text_auto=True
            )

            st.plotly_chart(fig_ventas, use_container_width=True)

        with col_g2:
            compras_summary = dfcompra[dfcompra["SKU"].isin(skus)]

            fig_compras = px.bar(
                compras_summary,
                x="Mes",
                y="total_quantity",
                color="SKU",
                title="📊 Valor de compras",
                text_auto=True
            )

            st.plotly_chart(fig_compras, use_container_width=True)

        st.markdown("---")



    # Gráficos principales 2
    col_g1, col_g2 = st.columns(2)

    with col_g1:
        abc_summary = (
            filtered
            .groupby("abc", as_index=False)["valor_ventas"]
            .sum()
        )

        fig_abc = px.bar(
            abc_summary,
            x="abc",
            y="valor_ventas",
            title="📊 Valor de ventas por clasificación ABC",
            text_auto=True
        )

        st.plotly_chart(fig_abc, use_container_width=True)

    with col_g2:
        status_summary = (
            filtered
            .groupby("estado_inventario", as_index=False)
            .size()
            .rename(columns={"size": "productos"})
        )

        fig_status = px.pie(
            status_summary,
            names="estado_inventario",
            values="productos",
            title="🚦 Estado general del inventario 2"
        )

        st.plotly_chart(fig_status, use_container_width=True)

    st.markdown("---")

    # Top productos
    st.subheader("🏆 Top 10 productos por valor de ventas")

    top_sales = (
        filtered
        .sort_values("valor_ventas", ascending=False)
        .head(10)
    )

    fig_top_sales = px.bar(
        top_sales,
        x="valor_ventas",
        y=product_col,
        orientation="h",
        title="Productos con mayor valor de ventas",
        text_auto=True
    )

    fig_top_sales.update_layout(
        yaxis={"categoryorder": "total ascending"}
    )

    st.plotly_chart(fig_top_sales, use_container_width=True)



# ============================================================
# PÁGINA 2: INVENTARIO
# ============================================================

elif pagina == "📦 Inventario":




    st.title("📦 Gestión de inventario")
    st.write("Detalle de productos, stock, ventas, rotación, cobertura y clasificación ABC.")

    # KPIs de inventario
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "📦 Stock total",
            f"{filtered[stock_col].sum():,.0f}"
        )

    with col2:
        st.metric(
            "💰 Valor inventario",
            f"${filtered['valor_inventario'].sum():,.0f}"
        )

    with col3:
        cobertura_promedio = (
            filtered["cobertura_meses"]
            .replace([np.inf, -np.inf], np.nan)
            .mean()
        )

        st.metric(
            "⏳ Cobertura promedio",
            f"{cobertura_promedio:.2f} meses"
        )

    with col4:
        rotacion_promedio = filtered["rotacion_mensual"].mean()

        st.metric(
            "🔁 Rotación promedio",
            f"{rotacion_promedio:.2f}"
        )

    st.markdown("---")

    # Filtros internos
#    filtro_categoria=selected_categories
#    filtro_abc=selected_abc
#    filtro_estado=selected_status
    

    inventario_filtrado = filtered[
        (filtered["abc"].isin(selected_abc)) &
        (filtered["estado_inventario"].isin(selected_status)) &
        (filtered[category_col].isin(selected_categories))
    ].copy()

    st.markdown("---")

    # Tabla principal
    st.subheader("📄 Tabla detallada de inventario")

    columnas_inventario = [
        product_col,
        category_col,
        stock_col,
        sales_col,
        "valor_inventario",
        "valor_ventas",
        "rotacion_mensual",
        "cobertura_dias",
        "cobertura_meses",
        "abc",
        "estado_inventario"
    ]

    st.dataframe(
        inventario_filtrado[columnas_inventario]
        .replace([np.inf, -np.inf], np.nan),
        use_container_width=True
    )

    st.markdown("---")

    # Visualización: cobertura vs rotación
    st.subheader("📊 Análisis de rotación y cobertura")

    fig_rotacion = px.scatter(
        inventario_filtrado.replace([np.inf, -np.inf], np.nan),
        x="rotacion_mensual",
        y="cobertura_meses",
        size="valor_ventas",
        color="estado_inventario",
        hover_name=product_col,
        title="Rotación mensual vs cobertura de inventario"
    )

    st.plotly_chart(fig_rotacion, use_container_width=True)

    st.markdown("---")

    # Ranking de productos con mayor inventario
    st.subheader("🏷️ Top 10 productos con mayor valor de inventario")

    top_inventory = (
        inventario_filtrado
        .sort_values("valor_inventario", ascending=False)
        .head(10)
    )

    fig_inventory = px.bar(
        top_inventory,
        x="valor_inventario",
        y=product_col,
        orientation="h",
        title="Productos con mayor capital inmovilizado",
        text_auto=True
    )

    fig_inventory.update_layout(
        yaxis={"categoryorder": "total ascending"}
    )

    st.plotly_chart(fig_inventory, use_container_width=True)


# ============================================================
# PÁGINA 3: ALERTAS
# ============================================================

elif pagina == "🚨 Alertas":

    st.title("🚨 Alertas de inventario")
    st.write("Identificación de productos con riesgo de quiebre, sobrestock o baja rotación.")

    # Filtrar solo productos con alertas
    alertas = filtered[
        filtered["estado_inventario"].isin(
            ["Riesgo de quiebre", "Sobrestock", "Baja rotación"]
        )
    ].copy()

    # KPIs de alertas
    col1, col2, col3 = st.columns(3)

    with col1:
        riesgo_quiebre = (
            alertas["estado_inventario"] == "Riesgo de quiebre"
        ).sum()

        st.metric(
            "🔴 Riesgo de quiebre",
            f"{riesgo_quiebre:,.0f}"
        )

    with col2:
        sobrestock = (
            alertas["estado_inventario"] == "Sobrestock"
        ).sum()

        st.metric(
            "🟡 Sobrestock",
            f"{sobrestock:,.0f}"
        )

    with col3:
        baja_rotacion = (
            alertas["estado_inventario"] == "Baja rotación"
        ).sum()

        st.metric(
            "📉 Baja rotación",
            f"{baja_rotacion:,.0f}"
        )

    st.markdown("---")

    if alertas.empty:
        st.success("🟢 No se detectan alertas críticas con los filtros actuales.")

    else:
        st.subheader("📋 Productos con alertas")

        columnas_alertas = [
            product_col,
            category_col,
            stock_col,
            sales_col,
            "abc",
            "rotacion_mensual",
            "cobertura_meses",
            "punto_reorden",
            "estado_inventario"
        ]

        st.dataframe(
            alertas[columnas_alertas]
            .replace([np.inf, -np.inf], np.nan)
            .sort_values("estado_inventario"),
            use_container_width=True
        )

        st.markdown("---")

        # Gráfico de alertas por tipo
        st.subheader("📊 Distribución de alertas")

        resumen_alertas = (
            alertas
            .groupby("estado_inventario", as_index=False)
            .size()
            .rename(columns={"size": "productos"})
        )

        fig_alertas = px.bar(
            resumen_alertas,
            x="estado_inventario",
            y="productos",
            title="Cantidad de productos por tipo de alerta",
            text_auto=True
        )

        st.plotly_chart(fig_alertas, use_container_width=True)

        st.markdown("---")

        # Prioridad de reposición
        st.subheader("🔁 Prioridad de reposición")

        reposicion = alertas[
            alertas["estado_inventario"] == "Riesgo de quiebre"
        ].copy()

        if reposicion.empty:
            st.info("No hay productos actualmente en riesgo de quiebre.")
        else:
            reposicion["cantidad_sugerida_reponer"] = (
                reposicion["punto_reorden"] - reposicion[stock_col]
            ).clip(lower=0)

            reposicion = reposicion.sort_values(
                "cantidad_sugerida_reponer",
                ascending=False
            )

            columnas_reposicion = [
                product_col,
                stock_col,
                sales_col,
                "punto_reorden",
                "cantidad_sugerida_reponer",
                "cobertura_meses"
            ]

            st.dataframe(
                reposicion[columnas_reposicion]
                .replace([np.inf, -np.inf], np.nan),
                use_container_width=True
            )

            fig_reposicion = px.bar(
                reposicion.head(10),
                x="cantidad_sugerida_reponer",
                y=product_col,
                orientation="h",
                title="Top productos con mayor necesidad de reposición",
                text_auto=True
            )

            fig_reposicion.update_layout(
                yaxis={"categoryorder": "total ascending"}
            )

            st.plotly_chart(fig_reposicion, use_container_width=True)

# ============================================================
# PÁGINA 5: FORECASTING
# ============================================================

elif pagina == "📈 Forecasting":

    st.title("📈 Forecasting de ventas")
    st.write("Predicción simple de ventas para apoyar decisiones de compra y reposición.")

    st.markdown("---")

    # Validar columnas necesarias
    columnas_necesarias = [product_col, sales_col]

    for col in columnas_necesarias:
        if col not in filtered.columns:
            st.error(f"Falta la columna necesaria: {col}")
            st.stop()

    # Parámetros
    st.subheader("⚙️ Parámetros de predicción")

    col1, col2 = st.columns(2)

    with col1:
        meses_proyeccion = st.slider(
            "Meses a proyectar",
            min_value=1,
            max_value=6,
            value=2
        )

    with col2:
        crecimiento_estimado = st.slider(
            "Crecimiento mensual estimado (%)",
            min_value=-50,
            max_value=100,
            value=10
        )

    st.info(
        "Esta página usa un modelo simple basado en crecimiento mensual estimado. "
        "Es ideal para una primera versión del portafolio."
    )

    st.markdown("---")

    # Base de forecasting
    forecast = filtered.copy()

    forecast["ventas_actuales"] = forecast[sales_col]

    for mes in range(1, meses_proyeccion + 1):
        forecast[f"forecast_mes_{mes}"] = (
            forecast["ventas_actuales"] *
            ((1 + crecimiento_estimado / 100) ** mes)
        )

    forecast["forecast_total"] = forecast[
        [f"forecast_mes_{mes}" for mes in range(1, meses_proyeccion + 1)]
    ].sum(axis=1)

    forecast["compra_sugerida"] = (
        forecast["forecast_total"] + forecast["stock_seguridad"] - forecast[stock_col]
    ).clip(lower=0)

    st.subheader("📊 Resumen ejecutivo de forecast")

    kpi1, kpi2, kpi3 = st.columns(3)

    with kpi1:
        st.metric(
            "📦 Ventas actuales",
            f"{forecast['ventas_actuales'].sum():,.0f}"
        )

    with kpi2:
        st.metric(
            "📈 Forecast total",
            f"{forecast['forecast_total'].sum():,.0f}"
        )

    with kpi3:
        st.metric(
            "🛒 Compra sugerida",
            f"{forecast['compra_sugerida'].sum():,.0f}"
        )

    st.markdown("---")

    # Tabla forecast
    st.subheader("📄 Forecast por producto")

    columnas_forecast = [
        product_col,
        category_col,
        stock_col,
        sales_col,
        "ventas_actuales"
    ] + [f"forecast_mes_{mes}" for mes in range(1, meses_proyeccion + 1)] + [
        "forecast_total",
        "stock_seguridad",
        "compra_sugerida",
        "abc",
        "estado_inventario"
    ]

    st.dataframe(
        forecast[columnas_forecast]
        .replace([np.inf, -np.inf], np.nan),
        use_container_width=True
    )

    st.markdown("---")

    # Gráfico top compra sugerida
    st.subheader("🛒 Top productos con mayor compra sugerida")

    top_compra = (
        forecast
        .sort_values("compra_sugerida", ascending=False)
        .head(10)
    )

    fig_compra = px.bar(
        top_compra,
        x="compra_sugerida",
        y=product_col,
        orientation="h",
        title="Top 10 productos por compra sugerida",
        text_auto=True
    )

    fig_compra.update_layout(
        yaxis={"categoryorder": "total ascending"}
    )

    st.plotly_chart(fig_compra, use_container_width=True)

    st.markdown("---")

    # Gráfico forecast por clasificación ABC
    st.subheader("🅰️ Forecast por clasificación ABC")

    forecast_abc = (
        forecast
        .groupby("abc", as_index=False)["forecast_total"]
        .sum()
    )

    fig_forecast_abc = px.bar(
        forecast_abc,
        x="abc",
        y="forecast_total",
        title="Forecast total por clasificación ABC",
        text_auto=True
    )

    st.plotly_chart(fig_forecast_abc, use_container_width=True)