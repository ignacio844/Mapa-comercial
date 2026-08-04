import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import unicodedata
import re
import io
import os
import pymysql

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Mapa Comercial", layout="wide", initial_sidebar_state="expanded")

# --- TÍTULO Y ESTILOS CSS ---
st.markdown("""
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
    .block-container { max-width: 98% !important; padding-top: 2rem !important; padding-left: 1rem !important; padding-right: 1rem !important; }
    .titulo-empresarial { font-family: 'Segoe UI', Roboto, sans-serif; font-size: 2.8rem; font-weight: 700; letter-spacing: -0.5px; padding-bottom: 10px; margin-bottom: 30px; border-bottom: 3px solid #1abc9c; }
    [data-testid="stMetric"] { background-color: #1f2937; border-radius: 10px; padding: 15px 20px; border: 1px solid #374151; box-shadow: 0 4px 6px rgba(0,0,0,0.3); text-align: center; }
    [data-testid="stMetricLabel"] { justify-content: center !important; display: flex !important; width: 100% !important; }
    [data-testid="stMetricLabel"] > div { color: #1abc9c !important; font-size: 1.05rem !important; font-weight: 700 !important; }
    [data-testid="stMetricValue"] { justify-content: center !important; display: flex !important; width: 100% !important; }
    [data-testid="stMetricValue"] > div { color: #f9fafb !important; font-size: 2.2rem !important; }
    div[data-testid="stVerticalBlockBorderWrapper"] { background-color: #1f2937 !important; border-radius: 10px !important; border: 1px solid #374151 !important; box-shadow: 0 4px 6px rgba(0,0,0,0.3) !important; padding: 1rem 1.5rem !important; }
    .chart-title { margin-top: 0 !important; margin-bottom: 1rem !important; color: #f9fafb; font-size: 1.4rem; font-weight: 600; font-family: 'Segoe UI', sans-serif; }
    .icon-header { vertical-align: text-bottom; font-size: 1.7rem; color: #1abc9c; margin-right: 8px; }
    </style>
    <h1 class="titulo-empresarial">Mapa Comercial</h1>
""", unsafe_allow_html=True)

# --- FUNCIONES DE FORMATO ---
def formato_corto(num, es_moneda=False):
    if num >= 1_000_000_000: val = f"{num / 1_000_000_000:.1f} B"
    elif num >= 1_000_000: val = f"{num / 1_000_000:.1f} M"
    elif num >= 1_000: val = f"{num / 1_000:.1f} K"
    else: val = f"{num:,.0f}".replace(",", ".")
    return f"${val}" if es_moneda else val

def formato_completo(num, es_moneda=False):
    val = f"{num:,.0f}".replace(",", ".")
    return f"${val}" if es_moneda else val

# --- CARGA DE DATOS (CONEXIÓN NATIVA CON PAGINACIÓN) ---
@st.cache_data(ttl=43200)
def cargar_datos():
    print("1. Iniciando conexión nativa a Aiven...")
    try:
        db = st.secrets["mysql"]
        
        # Conectamos con un contexto de seguridad más permisivo por si hay microcortes
        conexion = pymysql.connect(
            host=db['host'],
            user=db['username'],
            password=db['password'],
            database=db['database'],
            port=db['port'],
            ssl={'ssl_cert_reqs': 'CERT_NONE'}, 
            connect_timeout=20,
            read_timeout=60,
            cursorclass=pymysql.cursors.DictCursor
        )
        
        print("2. Descargando tabla por bloques (Paginación)...")
        todas_las_filas = []
        offset = 0
        limit = 300  # Pedimos de a 300 clientes para que sea ultra rápido y no sature
        
        with conexion.cursor() as cursor:
            while True:
                # Le pedimos a MySQL un pedacito específico de la tabla
                cursor.execute(f"SELECT * FROM clientes_geolocalizados LIMIT {limit} OFFSET {offset}")
                bloque = cursor.fetchall()
                
                if not bloque:
                    break # Si el bloque viene vacío, terminamos de descargar
                    
                todas_las_filas.extend(bloque)
                print(f"   -> Descargados {len(todas_las_filas)} clientes...")
                offset += limit
                
        print("3. Convirtiendo a formato Pandas...")
        df = pd.DataFrame(todas_las_filas)
        
        # Convertimos las columnas numéricas
        if not df.empty:
            for col in df.columns:
                if df[col].dtype == 'object' and col not in ["Nombre_Cliente", "Vendedor", "Direccion", "Numero", "Localidad", "Provincia", "cuit_cl", "Cliente_Key"]:
                    try:
                        df[col] = pd.to_numeric(df[col])
                    except:
                        pass
        
        conexion.close()
        
        print(f"4. ¡Éxito! Se descargaron {len(df)} filas en total.")
        return df

    except Exception as e:
        print(f"ERROR DETECTADO: {e}")
        st.error(f"Error de conexión: {e}")
        return pd.DataFrame()

# --- DESCARGA OPTIMIZADA ---
@st.cache_data(show_spinner=False)
def convertir_excel(df):
    output = io.BytesIO()
    df_export = df.drop(columns=["Latitud", "Longitud"], errors="ignore")
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Datos Filtrados')
    return output.getvalue()

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### Acciones")
    st.success("🟢 Conectado a Aiven MySQL")
    if st.button("🔄 Actualizar Datos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")

# --- FLUJO PRINCIPAL ---
detail = cargar_datos()

if not detail.empty:
    detail["Provincia"] = detail["Provincia"].astype(str).str.title().str.strip().replace("Nan", "")
    detail["Localidad"] = detail["Localidad"].astype(str).str.title().str.strip().replace("Nan", "")

    def opts(s): return sorted(x for x in s.dropna().astype(str).str.strip().unique() if x and x.lower() != 'nan')

    # --- FILTROS ---
    c1, c2, c3, c4 = st.columns(4)
    with c1: sellers = st.multiselect("VENDEDOR", opts(detail["Vendedor"]))
    with c2: provinces = st.multiselect("PROVINCIA", opts(detail["Provincia"]))
    with c3: locations = st.multiselect("LOCALIDAD", opts(detail["Localidad"]))
    with c4: search_query = st.text_input("BUSCAR CLIENTE (Nombre)", "")

    filtered = detail.copy()
    if sellers: filtered = filtered[filtered["Vendedor"].isin(sellers)]
    if provinces: filtered = filtered[filtered["Provincia"].isin(provinces)]
    if locations: filtered = filtered[filtered["Localidad"].isin(locations)]
    if search_query: filtered = filtered[filtered["Nombre_Cliente"].str.contains(search_query, case=False, na=False)]

    with st.sidebar:
        st.markdown("### Exportar")
        st.download_button(label="Descargar Búsqueda (Excel)", data=convertir_excel(filtered), file_name="Reporte_Clientes_Filtrados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown("---")

    # --- LÓGICA DE COLUMNAS SEGURA ---
    # Obtenemos solo las columnas numéricas para evitar sumar textos
    cols_numericas = filtered.select_dtypes(include=['number']).columns
    cols_excluir_num = {"TOTAL 2026", "Latitud", "Longitud", "Cliente_Key", "cuit_cl"} 
    cols_proveedores = [col for col in cols_numericas if col not in cols_excluir_num]

    # --- KPIs ---
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    total_facturacion = filtered["TOTAL 2026"].sum() if "TOTAL 2026" in filtered.columns else 0
    clientes_activos = (filtered["TOTAL 2026"] > 0).sum() if "TOTAL 2026" in filtered.columns else len(filtered)
    ticket_promedio = total_facturacion / clientes_activos if clientes_activos else 0
    
    total_marcas = sum(1 for prov in cols_proveedores if filtered[prov].sum() > 0)
    geolocalizados = filtered["Latitud"].notna().sum()

    kpi1.metric("FACTURACIÓN", formato_corto(total_facturacion, True), help=f"Valor exacto: {formato_completo(total_facturacion, True)}")
    kpi2.metric("CLIENTES ACTIVOS", f"{clientes_activos:,.0f}".replace(",", "."))
    kpi3.metric("TICKET PROMEDIO", formato_corto(ticket_promedio, True), help=f"Valor exacto: {formato_completo(ticket_promedio, True)}")
    kpi4.metric("MARCAS ACTIVAS", f"{total_marcas}")
    kpi5.metric("GEOLOCALIZADOS", f"{geolocalizados} / {len(filtered)}")
    st.markdown("---")

    # --- PREPARACIÓN DEL MAPA ---
    mapped = filtered.dropna(subset=["Latitud", "Longitud"]).copy()
    if "TOTAL 2026" in mapped.columns:
        mapped["Facturacion"] = mapped["TOTAL 2026"]
    if "Vendedor" in mapped.columns:
        mapped["Vendedor_Factura"] = mapped["Vendedor"]

    if not mapped.empty:
        with st.container(border=True):
            vista_mapa = st.radio("VISTA DEL MAPA", ["Mapa de Calor", "Marcadores", "Combinado"], horizontal=True, label_visibility="collapsed")
            center_lat, center_lon, zoom_level = -38.4161, -63.6167, 3.8
            cap = max(float(mapped["Facturacion"].quantile(.98)), 1) 
            mapped["Peso"] = mapped["Facturacion"].clip(0, cap)
            mapped["Facturacion_Formateada"] = mapped["Facturacion"].apply(lambda x: formato_corto(x, True))
            
            if vista_mapa == "Mapa de Calor":
                heat = px.density_map(mapped, lat="Latitud", lon="Longitud", z="Peso", radius=22, center={"lat": center_lat, "lon": center_lon}, zoom=zoom_level, map_style="carto-positron", hover_name="Nombre_Cliente", hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Peso": False}, labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"}, height=550, color_continuous_scale="Turbo")
                heat.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(heat, use_container_width=True, config={'scrollZoom': False})
            elif vista_mapa == "Marcadores":
                points = px.scatter_map(mapped, lat="Latitud", lon="Longitud", color="Facturacion", hover_name="Nombre_Cliente", hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Facturacion": False}, labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"}, center={"lat": center_lat, "lon": center_lon}, zoom=zoom_level, map_style="carto-positron", height=550, color_continuous_scale="Turbo")
                points.update_traces(marker=dict(size=7)) 
                points.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(points, use_container_width=True, config={'scrollZoom': False})
            else:
                combined = px.density_map(mapped, lat="Latitud", lon="Longitud", z="Peso", radius=22, center={"lat": center_lat, "lon": center_lon}, zoom=zoom_level, map_style="carto-positron", hover_name="Nombre_Cliente", hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Peso": False}, labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"}, height=550, color_continuous_scale="Turbo")
                puntos_extra = px.scatter_map(mapped, lat="Latitud", lon="Longitud", hover_name="Nombre_Cliente", hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False}, labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"})
                capa_puntos = puntos_extra.data[0]
                capa_puntos.marker.color = '#ffffff'; capa_puntos.marker.size = 4; capa_puntos.marker.opacity = 0.95
                combined.add_trace(capa_puntos)
                combined.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(combined, use_container_width=True, config={'scrollZoom': False})

        # --- GRÁFICOS INFERIORES ---
        col_graf_1, col_graf_2 = st.columns(2)
        
        with col_graf_1:
            with st.container(border=True):
                st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">groups</i> Top 10 Clientes</h3>', unsafe_allow_html=True)
                top_clientes = mapped.nlargest(10, "Facturacion").copy()
                top_clientes["Fact_Tooltip"] = top_clientes["Facturacion"].apply(lambda x: formato_completo(x, True))
                fig_cli = px.bar(top_clientes, x="Facturacion", y="Nombre_Cliente", orientation='h', color_discrete_sequence=['#e67e22'], text_auto='.3s', custom_data=["Fact_Tooltip"])
                fig_cli.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", xaxis=dict(visible=False), yaxis_title=None, yaxis={'categoryorder':'total ascending'}, height=400, margin=dict(l=0, r=0, t=10, b=0))
                fig_cli.update_traces(textfont_size=13, textangle=0, textposition="outside", cliponaxis=False, hovertemplate='<b>Cliente:</b> %{y}<br><b>Facturación:</b> %{customdata[0]}<extra></extra>')
                st.plotly_chart(fig_cli, use_container_width=True)

        with col_graf_2:
            with st.container(border=True):
                st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">leaderboard</i> Top 10 Vendedores</h3>', unsafe_allow_html=True)
                top_vend = filtered.groupby("Vendedor", as_index=False)["TOTAL 2026"].sum().nlargest(10, "TOTAL 2026")
                top_vend["Fact_Tooltip"] = top_vend["TOTAL 2026"].apply(lambda x: formato_completo(x, True))
                fig_vend = px.bar(top_vend, x="TOTAL 2026", y="Vendedor", orientation='h', color_discrete_sequence=['#34495e'], text_auto='.3s', custom_data=["Fact_Tooltip"])
                fig_vend.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", xaxis=dict(visible=False), yaxis_title=None, yaxis={'categoryorder':'total ascending'}, height=400, margin=dict(l=0, r=0, t=10, b=0))
                fig_vend.update_traces(textfont_size=13, textangle=0, textposition="outside", cliponaxis=False, hovertemplate='<b>Vendedor:</b> %{y}<br><b>Facturación:</b> %{customdata[0]}<extra></extra>')
                st.plotly_chart(fig_vend, use_container_width=True)
                
        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">domain</i> Top 10 Proveedores (Marcas)</h3>', unsafe_allow_html=True)
            # Calculamos las ventas por proveedor usando las columnas numéricas seguras
            ventas_prov = [{"Proveedor": prov, "Facturacion": filtered[prov].sum()} for prov in cols_proveedores]
            top_prov = pd.DataFrame(ventas_prov).nlargest(10, "Facturacion")
            
            if not top_prov.empty:
                top_prov["Fact_Tooltip"] = top_prov["Facturacion"].apply(lambda x: formato_completo(x, True))
                fig_prov = px.bar(top_prov, x="Facturacion", y="Proveedor", orientation='h', color_discrete_sequence=['#16a085'], text_auto='.3s', custom_data=["Fact_Tooltip"])
                fig_prov.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", xaxis=dict(visible=False), yaxis_title=None, yaxis={'categoryorder':'total ascending'}, height=400, margin=dict(l=0, r=0, t=10, b=0))
                fig_prov.update_traces(textfont_size=13, textangle=0, textposition="outside", cliponaxis=False, hovertemplate='<b>Proveedor:</b> %{y}<br><b>Facturación:</b> %{customdata[0]}<extra></extra>')
                st.plotly_chart(fig_prov, use_container_width=True)
            else:
                st.info("No hay datos numéricos de proveedores para mostrar.")

else:
    st.info("Haz clic en 'Actualizar Datos' en el menú de la izquierda para comenzar.")
