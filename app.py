import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import unicodedata
import re
import os
import io
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Mapa Comercial", layout="wide", initial_sidebar_state="expanded")

# --- TÍTULO, ESTILOS CSS E ÍCONOS PROFESIONALES ---
st.markdown("""
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    
    <style>
    .block-container {
        max-width: 98% !important;
        padding-top: 2rem !important; 
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    
    .titulo-empresarial {
        font-family: 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
        font-size: 2.8rem;
        font-weight: 700;
        letter-spacing: -0.5px;
        padding-bottom: 10px;
        margin-bottom: 30px;
        border-bottom: 3px solid #1abc9c;
    }

    /* --- TARJETAS DE KPIs --- */
    [data-testid="stMetric"] {
        background-color: #1f2937;
        border-radius: 10px;
        padding: 15px 20px;
        border: 1px solid #374151;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        text-align: center;
    }
    [data-testid="stMetricLabel"] {
        justify-content: center !important; 
        align-items: center !important;
        display: flex !important;
        width: 100% !important;
    }
    [data-testid="stMetricLabel"] > div {
        color: #1abc9c !important; 
        font-size: 1.05rem !important;
        font-weight: 700 !important;
    }
    [data-testid="stMetricValue"] {
        justify-content: center !important; 
        display: flex !important;
        width: 100% !important;
    }
    [data-testid="stMetricValue"] > div {
        color: #f9fafb !important; 
        font-size: 2.2rem !important;
    }
    
    /* --- TARJETAS DE GRÁFICOS (Contenedores) --- */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #1f2937 !important;
        border-radius: 10px !important;
        border: 1px solid #374151 !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3) !important;
        padding: 1rem 1.5rem !important;
    }
    
    .chart-title {
        margin-top: 0 !important;
        margin-bottom: 1rem !important;
        color: #f9fafb;
        font-size: 1.4rem;
        font-weight: 600;
        font-family: 'Segoe UI', sans-serif;
    }
    
    .icon-header {
        vertical-align: text-bottom; 
        font-size: 1.7rem; 
        color: #1abc9c; 
        margin-right: 8px;
    }
    </style>
    <h1 class="titulo-empresarial">Mapa Comercial</h1>
""", unsafe_allow_html=True)

def norm(v):
    if pd.isna(v): return ""
    t = unicodedata.normalize("NFKD", str(v).strip())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).upper()

# --- FUNCIONES DE FORMATO DE NÚMEROS ---
def formato_corto(num, es_moneda=False):
    if num >= 1_000_000_000:
        val = f"{num / 1_000_000_000:.1f} B"
    elif num >= 1_000_000:
        val = f"{num / 1_000_000:.1f} M"
    elif num >= 1_000:
        val = f"{num / 1_000:.1f} K"
    else:
        val = f"{num:,.0f}".replace(",", ".")
    return f"${val}" if es_moneda else val

def formato_completo(num, es_moneda=False):
    val = f"{num:,.0f}".replace(",", ".")
    return f"${val}" if es_moneda else val

# --- CONEXIÓN A SQL Y EXTRACCIÓN DE DATOS ---
# Usamos caché para no saturar la base de datos si cambiamos de pestaña rápido
@st.cache_data(ttl=3600) # Se refresca cada hora o si cambian las fechas
def extraer_datos_sql(fecha_inicio, fecha_fin):
    # La conexión usa los datos de .streamlit/secrets.toml
    # Requiere instalar: pip install pyodbc sqlalchemy
    conn = st.connection("sqlserver", type="sql")
    
    # Adaptamos tu lógica de DatosLimpios para agrupar por Cliente y Mes
    # Asumo que en Vista_Ventas_origen_v2 tienes columnas como:
    # Nombre_Cliente, Vendedor, Proveedor (o similar). Ajusta los nombres si es necesario.
    query = f"""
        WITH DatosLimpios AS (
            SELECT 
                Nombre_Cliente, -- AJUSTAR: Nombre de la columna del cliente en tu vista
                Vendedor_Factura, -- AJUSTAR: Nombre del vendedor
                Proveedor,       -- AJUSTAR: Nombre de la marca/proveedor
                fecha,
                Importe_sin_iva,
                CASE 
                    WHEN Importe_sin_iva < 0 AND cantidad > 0 THEN cantidad * -1
                    ELSE cantidad 
                END AS cantidad_neta
            FROM 
                [VS_REPORTING].[dbo].[Vista_Ventas_origen_v2]
            WHERE 
                fecha >= '{fecha_inicio}' AND fecha <= '{fecha_fin}'
        )
        SELECT 
            Nombre_Cliente,
            Vendedor_Factura,
            Proveedor,
            FORMAT(fecha, 'yyyy-MM') AS Mes_Agrupado, -- Formato AAAA-MM para la línea de tiempo
            SUM(cantidad_neta) AS Cant,
            SUM(Importe_sin_iva) AS [Total S/IVA]
        FROM 
            DatosLimpios
        GROUP BY 
            Nombre_Cliente,
            Vendedor_Factura,
            Proveedor,
            FORMAT(fecha, 'yyyy-MM')
    """
    
    df = conn.query(query)
    # Creamos la Key normalizada para cruzar con el Excel
    if not df.empty:
        df["Cliente_Key"] = df["Nombre_Cliente"].map(norm)
    return df

# --- CARGA DEL MAESTRO DE COORDENADAS (EXCEL FIJO) ---
@st.cache_data
def cargar_clientes_geolocalizados():
    archivo_local = 'Clientes_Geolocalizados.xlsx'
    if not os.path.exists(archivo_local):
        st.error(f"Falta el archivo maestro: {archivo_local}")
        return pd.DataFrame()
    try:
        clients = pd.read_excel(archivo_local, sheet_name="Clientes")
        clients["Cliente_Key"] = clients["Nombre_Cliente"].map(norm)
        return clients
    except Exception as e:
        st.error(f"Error leyendo coordenadas: {e}")
        return pd.DataFrame()

# --- DESCARGA OPTIMIZADA (EN CACHÉ) ---
@st.cache_data(show_spinner=False)
def convertir_excel(df):
    output = io.BytesIO()
    df_export = df.drop(columns=["Cliente_Key", "Latitud", "Longitud"], errors="ignore")
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Datos Filtrados')
    return output.getvalue()


# --- BARRA LATERAL (FILTRO DE FECHAS Y ACCIONES) ---
with st.sidebar:
    st.markdown("### Rango de Análisis")
    
    # Fechas por defecto: mes actual
    hoy = datetime.date.today()
    primer_dia_mes = hoy.replace(day=1)
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_inicio = st.date_input("Desde", primer_dia_mes)
    with col_f2:
        f_fin = st.date_input("Hasta", hoy)
        
    st.markdown("---")
    # Botón explícito para forzar la consulta al SQL
    buscar_sql = st.button("Consultar Base de Datos", type="primary", use_container_width=True)


# --- FLUJO PRINCIPAL ---
# Solo intentamos procesar si el usuario le dio al botón de buscar o si ya hay datos en memoria
if buscar_sql or 'datos_cargados' in st.session_state:
    st.session_state['datos_cargados'] = True
    
    with st.spinner('Extrayendo información desde SQL Server...'):
        data = extraer_datos_sql(f_inicio, f_fin)
        clients = cargar_clientes_geolocalizados()

    if not data.empty and not clients.empty:
        
        clients_unique = clients.drop_duplicates(subset=["Cliente_Key"])
        detail = data.merge(clients_unique[["Cliente_Key", "Vendedor", "Direccion", "Localidad", "Provincia", "Latitud", "Longitud"]], 
                            on="Cliente_Key", how="left")

        # --- LIMPIEZA DE TEXTOS PARA LOS FILTROS ---
        detail["Provincia"] = detail["Provincia"].astype(str).str.title().str.strip()
        detail["Localidad"] = detail["Localidad"].astype(str).str.title().str.strip()
        detail["Provincia"] = detail["Provincia"].replace("Nan", "")
        detail["Localidad"] = detail["Localidad"].replace("Nan", "")

        def opts(s): return sorted(x for x in s.dropna().astype(str).str.strip().unique() if x and x.lower() != 'nan')

        # --- FILTROS SUPERIORES ---
        c1, c2, c3, c4, c5 = st.columns(5)
        # Ahora usamos 'Mes_Agrupado' que viene del SQL
        with c1: months = st.multiselect("MES", opts(detail["Mes_Agrupado"]))
        with c2: providers = st.multiselect("PROVEEDOR / MARCA", opts(detail["Proveedor"]))
        with c3: sellers = st.multiselect("VENDEDOR", opts(detail["Vendedor_Factura"]))
        with c4: provinces = st.multiselect("PROVINCIA", opts(detail["Provincia"]))
        with c5: locations = st.multiselect("LOCALIDAD", opts(detail["Localidad"]))
        
        search_query = st.text_input("BUSCAR CLIENTE (Nombre)", "")

        filtered = detail.copy()
        filtros = [
            ("Mes_Agrupado", months), ("Proveedor", providers), ("Vendedor_Factura", sellers),
            ("Provincia", provinces), ("Localidad", locations)
        ]
        for col, sel in filtros:
            if sel: filtered = filtered[filtered[col].astype(str).isin(sel)]
                
        if search_query:
            filtered = filtered[filtered["Nombre_Cliente"].str.contains(search_query, case=False, na=False)]

        with st.sidebar:
            st.markdown("### Exportar")
            st.download_button(
                label="Descargar Búsqueda (Excel)",
                data=convertir_excel(filtered),
                file_name="Reporte_Clientes_Filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        st.markdown("---")

        # --- KPIs ---
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        
        total_facturacion = filtered["Total S/IVA"].sum()
        total_unidades = filtered["Cant"].sum()
        clientes_activos = filtered.loc[filtered["Total S/IVA"] > 0, "Cliente_Key"].nunique()
        ticket_promedio = total_facturacion / clientes_activos if clientes_activos else 0
        total_marcas = filtered["Proveedor"].nunique()

        kpi1.metric("FACTURACIÓN", formato_corto(total_facturacion, True), help=f"Valor exacto: {formato_completo(total_facturacion, True)}")
        kpi2.metric("UNIDADES", formato_corto(total_unidades, False), help=f"Valor exacto: {formato_completo(total_unidades, False)}")
        kpi3.metric("CLIENTES ACTIVOS", f"{clientes_activos}")
        kpi4.metric("TICKET PROMEDIO", formato_corto(ticket_promedio, True), help=f"Valor exacto: {formato_completo(ticket_promedio, True)}")
        kpi5.metric("MARCAS", f"{total_marcas}")

        st.markdown("---")

        summary_map = filtered.groupby(["Cliente_Key", "Nombre_Cliente", "Vendedor_Factura", "Latitud", "Longitud"], dropna=False, as_index=False).agg(
            Facturacion=("Total S/IVA", "sum"), Unidades=("Cant", "sum")
        )
        mapped = summary_map.dropna(subset=["Latitud", "Longitud"]).copy()

        # --- MAPA OPTIMIZADO (Rendimiento WebGL) ---
        if not mapped.empty:
            with st.container(border=True):
                vista_mapa = st.radio(
                    "VISTA DEL MAPA", 
                    ["Mapa de Calor", "Marcadores", "Combinado"], 
                    horizontal=True,
                    label_visibility="collapsed"
                )
                
                center_lat = -38.4161
                center_lon = -63.6167
                zoom_level = 3.8
                
                cap = max(float(mapped["Facturacion"].quantile(.98)), 1) 
                mapped["Peso"] = mapped["Facturacion"].clip(0, cap)
                mapped["Facturacion_Formateada"] = mapped["Facturacion"].apply(lambda x: formato_corto(x, True))
                
                if vista_mapa == "Mapa de Calor":
                    heat = px.density_map(
                        mapped, lat="Latitud", lon="Longitud", z="Peso", radius=22, 
                        center={"lat": center_lat, "lon": center_lon}, zoom=zoom_level, map_style="carto-positron", 
                        hover_name="Nombre_Cliente", 
                        hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Peso": False},
                        labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"},
                        height=550, color_continuous_scale="Turbo"
                    )
                    heat.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(heat, use_container_width=True, config={'scrollZoom': False})
                    
                elif vista_mapa == "Marcadores":
                    points = px.scatter_map(
                        mapped, lat="Latitud", lon="Longitud", color="Facturacion", 
                        hover_name="Nombre_Cliente", 
                        hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Facturacion": False},
                        labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"},
                        center={"lat": center_lat, "lon": center_lon}, zoom=zoom_level, map_style="carto-positron", height=550,
                        color_continuous_scale="Turbo"
                    )
                    points.update_traces(marker=dict(size=7)) 
                    points.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(points, use_container_width=True, config={'scrollZoom': False})
                    
                else:
                    combined = px.density_map(
                        mapped, lat="Latitud", lon="Longitud", z="Peso", radius=22, 
                        center={"lat": center_lat, "lon": center_lon}, zoom=zoom_level, map_style="carto-positron", 
                        hover_name="Nombre_Cliente", 
                        hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Peso": False},
                        labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"},
                        height=550, color_continuous_scale="Turbo"
                    )
                    
                    puntos_extra = px.scatter_map(
                        mapped, lat="Latitud", lon="Longitud", 
                        hover_name="Nombre_Cliente", 
                        hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False},
                        labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"}
                    )
                    
                    capa_puntos = puntos_extra.data[0]
                    capa_puntos.marker.color = '#ffffff' 
                    capa_puntos.marker.size = 4 
                    capa_puntos.marker.opacity = 0.95
                    
                    combined.add_trace(capa_puntos)
                    combined.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(combined, use_container_width=True, config={'scrollZoom': False})

        # --- GRÁFICOS INFERIORES ---
        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">timeline</i> Evolución Mensual</h3>', unsafe_allow_html=True)
            evolucion = filtered.groupby("Mes_Agrupado", as_index=False)["Total S/IVA"].sum()
            evolucion["Fact_Tooltip"] = evolucion["Total S/IVA"].apply(lambda x: formato_completo(x, True))
            
            # Ordenamos por la fecha AAAA-MM
            evolucion = evolucion.sort_values(by="Mes_Agrupado")
            
            fig_evo = px.area(evolucion, x="Mes_Agrupado", y="Total S/IVA", markers=True, custom_data=["Fact_Tooltip"])
            fig_evo.update_traces(
                line_color='#1abc9c', fill='tozeroy', fillcolor='rgba(26, 188, 156, 0.2)',
                hovertemplate='<b>Mes:</b> %{x}<br><b>Facturación:</b> %{customdata[0]}<extra></extra>'
            )
            fig_evo.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis_title=None, yaxis_title=None,
                yaxis=dict(showgrid=True, gridcolor='#374151'), xaxis=dict(showgrid=False),
                height=350, margin=dict(l=0, r=0, t=10, b=0)
            )
            st.plotly_chart(fig_evo, use_container_width=True)

        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">domain</i> Top 10 Proveedores</h3>', unsafe_allow_html=True)
            top_prov = filtered.groupby("Proveedor", as_index=False)["Total S/IVA"].sum().nlargest(10, "Total S/IVA")
            top_prov["Fact_Tooltip"] = top_prov["Total S/IVA"].apply(lambda x: formato_completo(x, True))
            
            fig_prov = px.bar(top_prov, x="Total S/IVA", y="Proveedor", orientation='h', 
                              color_discrete_sequence=['#16a085'], text_auto='.3s', custom_data=["Fact_Tooltip"])
            fig_prov.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(visible=False), yaxis_title=None,
                yaxis={'categoryorder':'total ascending'},
                height=400, margin=dict(l=0, r=0, t=10, b=0)
            )
            fig_prov.update_traces(
                textfont_size=13, textangle=0, textposition="outside", cliponaxis=False,
                hovertemplate='<b>Proveedor:</b> %{y}<br><b>Facturación:</b> %{customdata[0]}<extra></extra>'
            )
            st.plotly_chart(fig_prov, use_container_width=True)

        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">groups</i> Top 10 Clientes</h3>', unsafe_allow_html=True)
            top_clientes = summary_map.nlargest(10, "Facturacion").copy()
            top_clientes["Fact_Tooltip"] = top_clientes["Facturacion"].apply(lambda x: formato_completo(x, True))
            
            fig_cli = px.bar(top_clientes, x="Facturacion", y="Nombre_Cliente", orientation='h', 
                             color_discrete_sequence=['#e67e22'], text_auto='.3s', custom_data=["Fact_Tooltip"])
            fig_cli.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(visible=False), yaxis_title=None,
                yaxis={'categoryorder':'total ascending'},
                height=400, margin=dict(l=0, r=0, t=10, b=0)
            )
            fig_cli.update_traces(
                textfont_size=13, textangle=0, textposition="outside", cliponaxis=False,
                hovertemplate='<b>Cliente:</b> %{y}<br><b>Facturación:</b> %{customdata[0]}<extra></extra>'
            )
            st.plotly_chart(fig_cli, use_container_width=True)

        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">leaderboard</i> Ranking 10 Vendedores</h3>', unsafe_allow_html=True)
            top_vend = filtered.groupby("Vendedor_Factura", as_index=False)["Total S/IVA"].sum().nlargest(10, "Total S/IVA")
            top_vend["Fact_Tooltip"] = top_vend["Total S/IVA"].apply(lambda x: formato_completo(x, True))
            
            fig_vend = px.bar(top_vend, x="Total S/IVA", y="Vendedor_Factura", orientation='h', 
                              color_discrete_sequence=['#34495e'], text_auto='.3s', custom_data=["Fact_Tooltip"])
            fig_vend.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(visible=False), yaxis_title=None,
                yaxis={'categoryorder':'total ascending'},
                height=400, margin=dict(l=0, r=0, t=10, b=0)
            )
            fig_vend.update_traces(
                textfont_size=13, textangle=0, textposition="outside", cliponaxis=False,
                hovertemplate='<b>Vendedor:</b> %{y}<br><b>Facturación:</b> %{customdata[0]}<extra></extra>'
            )
            st.plotly_chart(fig_vend, use_container_width=True)

    else:
        st.warning("No hay datos para mostrar en el rango de fechas seleccionado.")
else:
    # Mensaje inicial si no se ha buscado
    st.info("👋 ¡Bienvenido! Selecciona el rango de fechas en el menú lateral y haz clic en 'Consultar Base de Datos' para comenzar.")
