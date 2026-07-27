import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import unicodedata
import re
import os
import io
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Mapa Comercial", page_icon="📊", layout="wide")

# --- TÍTULO, ESTILOS CSS E ÍCONOS PROFESIONALES ---
st.markdown("""
    <!-- Importamos los íconos de Google Material -->
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
    
    /* Títulos de los gráficos dentro de las tarjetas */
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

# --- CARGA DE DATOS ---
@st.cache_data
def cargar_datos():
    archivo_local = 'Clientes_Geolocalizados.xlsx'
    if not os.path.exists(archivo_local):
        return pd.DataFrame(), pd.DataFrame()
    try:
        xls = pd.ExcelFile(archivo_local)
        if 'Data' not in xls.sheet_names or 'Clientes' not in xls.sheet_names:
            os.remove(archivo_local) 
            return pd.DataFrame(), pd.DataFrame()
            
        data = pd.read_excel(xls, "Data")
        clients = pd.read_excel(xls, "Clientes")
        
        data["Cliente_Key"] = data["Nombre_Cliente"].map(norm)
        clients["Cliente_Key"] = clients["Nombre_Cliente"].map(norm)
        
        data["Cant"] = pd.to_numeric(data["Cant"], errors="coerce").fillna(0)
        data["Total S/IVA"] = pd.to_numeric(data["Total S/IVA"], errors="coerce").fillna(0)
        
        return data, clients
    except Exception as e:
        if os.path.exists(archivo_local):
            os.remove(archivo_local)
        return pd.DataFrame(), pd.DataFrame()

# --- ACTUALIZAR DESDE DRIVE ---
FILE_ID = '1Xe1iull8fs7xUZbajYSMjEhzbxet2fui'

def actualizar_desde_drive():
    with st.spinner('Conectando a Google Drive y procesando...'):
        try:
            SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
            if os.path.exists('credenciales.json'):
                creds = Credentials.from_service_account_file('credenciales.json', scopes=SCOPES)
            else:
                creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
            
            servicio = build('drive', 'v3', credentials=creds)
            request = servicio.files().get_media(fileId=FILE_ID)
            
            fh = io.FileIO('Data_Descargada_Temp.xlsx', 'wb')
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            fh.close()

            xls = pd.ExcelFile('Data_Descargada_Temp.xlsx')
            data = pd.read_excel(xls, 'Data')
            clients = pd.read_excel(xls, 'Clientes')
            
            coordenadas_provincias = {
                'BUENOS AIRES': (-36.0, -60.0, 1.5), 'CAPITAL FEDERAL': (-34.60, -58.38, 0.05),
                'CORDOBA': (-31.5, -64.0, 1.0), 'SANTA FE': (-30.5, -61.0, 1.2),
                'MENDOZA': (-34.5, -68.5, 0.8), 'TUCUMAN': (-27.0, -65.5, 0.3),
                'ENTRE RIOS': (-32.0, -59.0, 0.8), 'SALTA': (-24.5, -65.0, 0.7),
                'MISIONES': (-27.0, -54.5, 0.4), 'CHACO': (-26.5, -60.5, 0.8),
                'CORRIENTES': (-28.5, -58.5, 0.8), 'NEUQUEN': (-38.5, -70.0, 0.8),
                'RIO NEGRO': (-40.5, -67.0, 1.0), 'SAN JUAN': (-31.0, -69.0, 0.6),
                'SANTIAGO DEL ESTERO': (-27.5, -63.0, 0.8), 'SAN LUIS': (-33.5, -66.0, 0.6),
                'LA PAMPA': (-37.0, -65.0, 1.0), 'SANTA CRUZ': (-49.0, -70.0, 1.5),
                'CHUBUT': (-44.0, -69.0, 1.2), 'JUJUY': (-23.0, -66.0, 0.4),
                'TIERRA DEL FUEGO': (-54.0, -67.0, 0.3), 'CATAMARCA': (-27.5, -67.0, 0.6),
                'FORMOSA': (-24.5, -60.0, 0.6), 'LA RIOJA': (-29.5, -67.0, 0.6),
            }
            
            coordenadas_ciudades = {
                'MAR DEL PLATA': (-38.002, -57.556), 'BAHIA BLANCA': (-38.719, -62.272),
                'ROSARIO': (-32.946, -60.639), 'CORDOBA': (-31.420, -64.188),
                'MENDOZA': (-32.890, -68.827), 'TUCUMAN': (-26.819, -65.222),
                'LA PLATA': (-34.920, -57.953), 'SANTA FE': (-31.633, -60.700),
                'SALTA': (-24.782, -65.411), 'SAN JUAN': (-31.537, -68.536),
                'RESISTENCIA': (-27.451, -58.986), 'NEUQUEN': (-38.951, -68.059),
                'TANDIL': (-37.321, -59.133), 'POSADAS': (-27.362, -55.896),
                'PARANA': (-31.731, -60.531), 'FORMOSA': (-26.184, -58.173),
                'SAN LUIS': (-33.301, -66.337), 'LA RIOJA': (-29.413, -66.855),
                'RIO CUARTO': (-33.123, -64.349), 'CONCORDIA': (-31.393, -58.016),
                'CORRIENTES': (-27.469, -58.834), 'SAN MIGUEL DE TUCUMAN': (-26.819, -65.222),
                'QUILMES': (-34.729, -58.267), 'MORON': (-34.653, -58.619),
                'SAN ISIDRO': (-34.471, -58.526), 'PILAR': (-34.458, -58.914)
            }

            clients['Prov_Limpia'] = clients['Provincia'].astype(str).str.upper().str.strip()
            clients['Localidad_Limpia'] = clients['Localidad'].astype(str).str.upper().str.strip()
            
            reemplazos = {
                'CÓRDOBA': 'CORDOBA', 'ENTRE RÍOS': 'ENTRE RIOS', 'TUCUMÁN': 'TUCUMAN',
                'NEUQUÉN': 'NEUQUEN', 'RÍO NEGRO': 'RIO NEGRO', 'SANTE FE': 'SANTA FE',
                'NAN': 'BUENOS AIRES', 
            }
            for mal, bien in reemplazos.items():
                clients['Prov_Limpia'] = clients['Prov_Limpia'].str.replace(mal, bien)

            np.random.seed(42) 
            lats, lons = [], []
            
            for prov, loc in zip(clients['Prov_Limpia'], clients['Localidad_Limpia']):
                if loc in coordenadas_ciudades:
                    lat = np.random.normal(coordenadas_ciudades[loc][0], 0.015)
                    lon = np.random.normal(coordenadas_ciudades[loc][1], 0.015)
                else:
                    lat_center, lon_center, std = coordenadas_provincias.get(prov, coordenadas_provincias['BUENOS AIRES'])
                    std = std * 0.6 
                    lat = np.random.normal(lat_center, std)
                    lon = np.random.normal(lon_center, std)
                    
                    if prov == 'BUENOS AIRES':
                        if lat < -35.5 and lon > -56.5: lon = -56.5 - abs(np.random.normal(0, 0.1))
                        if lat < -37.5 and lon > -57.5: lon = -57.5 - abs(np.random.normal(0, 0.1))
                        if lat < -38.5 and lon > -59.0: lon = -59.0 - abs(np.random.normal(0, 0.1))

                lats.append(lat)
                lons.append(lon)

            clients['Latitud'] = lats
            clients['Longitud'] = lons

            with pd.ExcelWriter('Clientes_Geolocalizados.xlsx') as writer:
                data.to_excel(writer, sheet_name='Data', index=False)
                clients.to_excel(writer, sheet_name='Clientes', index=False)
                
            if os.path.exists('Data_Descargada_Temp.xlsx'):
                os.remove('Data_Descargada_Temp.xlsx')
                
            cargar_datos.clear() 
            st.success("¡Base de datos actualizada con éxito!")
            
        except Exception as e:
            st.error(f"Ocurrió un error al actualizar: {e}")

# --- BARRA LATERAL ---
with st.sidebar:
    if st.button("Actualizar Datos desde Drive", type="primary", use_container_width=True):
        actualizar_desde_drive()
    st.markdown("---")

# --- FLUJO PRINCIPAL ---
data, clients = cargar_datos()

if not data.empty and not clients.empty:
    
    clients_unique = clients.drop_duplicates(subset=["Cliente_Key"])
    detail = data.merge(clients_unique[["Cliente_Key", "Vendedor", "Direccion", "Localidad", "Provincia", "Latitud", "Longitud"]], 
                        on="Cliente_Key", how="left")

    def opts(s): return sorted(x for x in s.dropna().astype(str).str.strip().unique() if x)

    # --- FILTROS SUPERIORES ---
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: months = st.multiselect("MES", opts(detail["Mes"]))
    with c2: providers = st.multiselect("PROVEEDOR / MARCA", opts(detail["Proveedor"]))
    with c3: sellers = st.multiselect("VENDEDOR", opts(detail["Vendedor_Factura"]))
    with c4: provinces = st.multiselect("PROVINCIA", opts(detail["Provincia"]))
    with c5: locations = st.multiselect("LOCALIDAD", opts(detail["Localidad"]))
    
    search_query = st.text_input("BUSCAR CLIENTE (Nombre)", "")

    filtered = detail.copy()
    filtros = [
        ("Mes", months), ("Proveedor", providers), ("Vendedor_Factura", sellers),
        ("Provincia", provinces), ("Localidad", locations)
    ]
    for col, sel in filtros:
        if sel: filtered = filtered[filtered[col].astype(str).isin(sel)]
            
    if search_query:
        filtered = filtered[filtered["Nombre_Cliente"].str.contains(search_query, case=False, na=False)]

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

    summary_map = filtered.groupby(["Cliente_Key", "Nombre_Cliente", "Latitud", "Longitud"], dropna=False, as_index=False).agg(
        Facturacion=("Total S/IVA", "sum"), Unidades=("Cant", "sum")
    )
    mapped = summary_map.dropna(subset=["Latitud", "Longitud"]).copy()

# 1. Agregamos "Vendedor_Factura" a la agrupación
    summary_map = filtered.groupby(["Cliente_Key", "Nombre_Cliente", "Vendedor_Factura", "Latitud", "Longitud"], dropna=False, as_index=False).agg(
        Facturacion=("Total S/IVA", "sum"), Unidades=("Cant", "sum")
    )
    mapped = summary_map.dropna(subset=["Latitud", "Longitud"]).copy()

# --- MAPA (EN TARJETA) ---
    if not mapped.empty:
        with st.container(border=True):
            # 1. Agregamos la tercera pestaña "Combinado"
            tabs = st.tabs(["Mapa de Calor", "Marcadores", "Combinado"])
            
            center = {"lat": float(mapped["Latitud"].median()), "lon": float(mapped["Longitud"].median())}
            cap = max(float(mapped["Facturacion"].quantile(.98)), 1) 
            mapped["Peso"] = mapped["Facturacion"].clip(0, cap)
            mapped["Tamaño_Marcador"] = mapped["Facturacion"].clip(lower=0)
            
            # Formateamos la facturación para que se vea bonita en el tooltip del mapa
            mapped["Facturacion_Formateada"] = mapped["Facturacion"].apply(lambda x: formato_corto(x, True))
            
            with tabs[0]:
                heat = px.density_map(
                    mapped, lat="Latitud", lon="Longitud", z="Peso", radius=22, 
                    center=center, zoom=3.2, map_style="carto-positron", 
                    hover_name="Nombre_Cliente", 
                    hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Peso": False},
                    labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"},
                    height=550, color_continuous_scale="Turbo"
                )
                heat.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(heat, use_container_width=True)
                
            with tabs[1]:
                points = px.scatter_map(
                    mapped, lat="Latitud", lon="Longitud", size="Tamaño_Marcador", color="Facturacion", 
                    hover_name="Nombre_Cliente", 
                    hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Facturacion": False, "Tamaño_Marcador": False},
                    labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"},
                    center=center, zoom=3.2, map_style="carto-positron", height=550
                )
                points.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(points, use_container_width=True)
                
            # 2. Nueva pestaña: Calor + Marcadores
            with tabs[2]:
                # Base: Mapa de calor
                combined = px.density_map(
                    mapped, lat="Latitud", lon="Longitud", z="Peso", radius=22, 
                    center=center, zoom=3.2, map_style="carto-positron", 
                    hover_name="Nombre_Cliente", 
                    hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Peso": False},
                    labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"},
                    height=550, color_continuous_scale="Turbo"
                )
                # Capa extra: Puntos de dispersión (grises oscuros para contrastar)
                puntos_extra = px.scatter_map(
                    mapped, lat="Latitud", lon="Longitud", size="Tamaño_Marcador", 
                    color_discrete_sequence=["#1f2937"], 
                    hover_name="Nombre_Cliente", 
                    hover_data={"Facturacion_Formateada": True, "Vendedor_Factura": True, "Latitud": False, "Longitud": False, "Tamaño_Marcador": False},
                    labels={"Facturacion_Formateada": "Facturación", "Vendedor_Factura": "Vendedor"}
                )
                # Inyectamos los puntos al mapa base
                capa_puntos = puntos_extra.data[0]
                capa_puntos.marker.opacity = 0.8
                combined.add_trace(capa_puntos)
                
                combined.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(combined, use_container_width=True)

        # --- GRÁFICOS INFERIORES (EN TARJETAS) ---
        
        # 1. Evolución mensual
        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">timeline</i> Evolución Mensual</h3>', unsafe_allow_html=True)
            evolucion = filtered.groupby("Mes", as_index=False)["Total S/IVA"].sum()
            orden_meses = ["Enero", "01-Enero", "Febrero", "02-Febrero", "Marzo", "03-Marzo", 
                           "Abril", "04-Abril", "Mayo", "05-Mayo", "Junio", "06-Junio", 
                           "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            
            fig_evo = px.area(evolucion, x="Mes", y="Total S/IVA", markers=True, category_orders={"Mes": orden_meses})
            fig_evo.update_traces(line_color='#1abc9c', fill='tozeroy', fillcolor='rgba(26, 188, 156, 0.2)')
            fig_evo.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis_title=None, yaxis_title=None,
                yaxis=dict(showgrid=True, gridcolor='#374151'), xaxis=dict(showgrid=False),
                height=350, margin=dict(l=0, r=0, t=10, b=0)
            )
            st.plotly_chart(fig_evo, use_container_width=True)

        # 2. Top Proveedores
        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">domain</i> Top 10 Proveedores</h3>', unsafe_allow_html=True)
            top_prov = filtered.groupby("Proveedor", as_index=False)["Total S/IVA"].sum().nlargest(10, "Total S/IVA")
            fig_prov = px.bar(top_prov, x="Total S/IVA", y="Proveedor", orientation='h', 
                              color_discrete_sequence=['#16a085'], text_auto='.3s')
            fig_prov.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(visible=False), yaxis_title=None,
                yaxis={'categoryorder':'total ascending'},
                height=400, margin=dict(l=0, r=0, t=10, b=0)
            )
            fig_prov.update_traces(textfont_size=13, textangle=0, textposition="outside", cliponaxis=False)
            st.plotly_chart(fig_prov, use_container_width=True)

        # 3. Top Clientes
        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">groups</i> Top 10 Clientes</h3>', unsafe_allow_html=True)
            top_clientes = summary_map.nlargest(10, "Facturacion")
            fig_cli = px.bar(top_clientes, x="Facturacion", y="Nombre_Cliente", orientation='h', 
                             color_discrete_sequence=['#e67e22'], text_auto='.3s')
            fig_cli.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(visible=False), yaxis_title=None,
                yaxis={'categoryorder':'total ascending'},
                height=400, margin=dict(l=0, r=0, t=10, b=0)
            )
            fig_cli.update_traces(textfont_size=13, textangle=0, textposition="outside", cliponaxis=False)
            st.plotly_chart(fig_cli, use_container_width=True)

        # 4. Ranking Vendedores
        with st.container(border=True):
            st.markdown('<h3 class="chart-title"><i class="material-icons icon-header">leaderboard</i> Ranking 10 Vendedores</h3>', unsafe_allow_html=True)
            top_vend = filtered.groupby("Vendedor_Factura", as_index=False)["Total S/IVA"].sum().nlargest(10, "Total S/IVA")
            fig_vend = px.bar(top_vend, x="Total S/IVA", y="Vendedor_Factura", orientation='h', 
                              color_discrete_sequence=['#34495e'], text_auto='.3s')
            fig_vend.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(visible=False), yaxis_title=None,
                yaxis={'categoryorder':'total ascending'},
                height=400, margin=dict(l=0, r=0, t=10, b=0)
            )
            fig_vend.update_traces(textfont_size=13, textangle=0, textposition="outside", cliponaxis=False)
            st.plotly_chart(fig_vend, use_container_width=True)

    else:
        st.warning("No hay datos para mostrar con los filtros seleccionados.")
else:
    st.info("Haz clic en el botón 'Actualizar Datos desde Drive' en el menú de la izquierda para comenzar.")
