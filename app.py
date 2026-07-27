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
st.set_page_config(page_title="Mapa Comercial", page_icon="🗺️", layout="wide")
st.title("🗺️ Mapa Comercial - Ventas y Territorio")

def norm(v):
    if pd.isna(v): return ""
    t = unicodedata.normalize("NFKD", str(v).strip())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).upper()

# --- CARGA DE DATOS BLINDADA ---
@st.cache_data
def cargar_datos():
    archivo_local = 'Clientes_Geolocalizados.xlsx'
    
    # Si no existe, devolvemos DataFrames vacíos
    if not os.path.exists(archivo_local):
        return pd.DataFrame(), pd.DataFrame()
    
    try:
        xls = pd.ExcelFile(archivo_local)
        # Verificamos que tenga las dos pestañas necesarias
        if 'Data' not in xls.sheet_names or 'Clientes' not in xls.sheet_names:
            os.remove(archivo_local) # Borramos el archivo corrupto/viejo
            return pd.DataFrame(), pd.DataFrame()
            
        data = pd.read_excel(xls, "Data")
        clients = pd.read_excel(xls, "Clientes")
        
        data["Cliente_Key"] = data["Nombre_Cliente"].map(norm)
        clients["Cliente_Key"] = clients["Nombre_Cliente"].map(norm)
        
        data["Cant"] = pd.to_numeric(data["Cant"], errors="coerce").fillna(0)
        data["Total S/IVA"] = pd.to_numeric(data["Total S/IVA"], errors="coerce").fillna(0)
        
        return data, clients
    except Exception as e:
        # Si hay cualquier otro error al leer, borramos para forzar actualización
        if os.path.exists(archivo_local):
            os.remove(archivo_local)
        return pd.DataFrame(), pd.DataFrame()

# --- ACTUALIZAR DESDE DRIVE ---
# ¡VERIFICA QUE ESTE ID SEA EL DE "Data Ene-Jun26.xlsx"!
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

            clients['Prov_Limpia'] = clients['Provincia'].astype(str).str.upper().str.strip()
            reemplazos = {
                'CÓRDOBA': 'CORDOBA', 'ENTRE RÍOS': 'ENTRE RIOS', 'TUCUMÁN': 'TUCUMAN',
                'NEUQUÉN': 'NEUQUEN', 'RÍO NEGRO': 'RIO NEGRO', 'SANTE FE': 'SANTA FE',
                'NAN': 'BUENOS AIRES', 
            }
            for mal, bien in reemplazos.items():
                clients['Prov_Limpia'] = clients['Prov_Limpia'].str.replace(mal, bien)

            np.random.seed(42) 
            lats, lons = [], []
            for prov in clients['Prov_Limpia']:
                lat, lon, std = coordenadas_provincias.get(prov, coordenadas_provincias['BUENOS AIRES'])
                lats.append(np.random.normal(lat, std))
                lons.append(np.random.normal(lon, std))

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
    if st.button("🔄 Actualizar Datos desde Drive", type="primary", use_container_width=True):
        actualizar_desde_drive()
    st.markdown("---")

# --- FLUJO PRINCIPAL ---
data, clients = cargar_datos()

if not data.empty and not clients.empty:
    
    clients_unique = clients.drop_duplicates(subset=["Cliente_Key"])
    detail = data.merge(clients_unique[["Cliente_Key", "Vendedor", "Direccion", "Localidad", "Provincia", "Latitud", "Longitud"]], 
                        on="Cliente_Key", how="left")

    def money(v): return (f"${v:,.0f}").replace(",",".")
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

    kpi1.metric("FACTURACIÓN", money(total_facturacion))
    kpi2.metric("UNIDADES", f"{total_unidades:,.0f}".replace(",", "."))
    kpi3.metric("CLIENTES ACTIVOS", clientes_activos)
    kpi4.metric("TICKET PROMEDIO", money(ticket_promedio))
    kpi5.metric("MARCAS", total_marcas)

    st.markdown("---")

    summary_map = filtered.groupby(["Cliente_Key", "Nombre_Cliente", "Latitud", "Longitud"], dropna=False, as_index=False).agg(
        Facturacion=("Total S/IVA", "sum"), Unidades=("Cant", "sum")
    )
    mapped = summary_map.dropna(subset=["Latitud", "Longitud"]).copy()

    # --- MAPA Y GRÁFICOS ---
   # --- 4. EL MAPA Y GRÁFICOS ---
    if not mapped.empty:
        tabs = st.tabs(["🔥 Mapa de Calor", "📍 Marcadores"])
        
        center = {"lat": float(mapped["Latitud"].median()), "lon": float(mapped["Longitud"].median())}
        cap = max(float(mapped["Facturacion"].quantile(.98)), 1) 
        mapped["Peso"] = mapped["Facturacion"].clip(0, cap)
        
        # NUEVA LÍNEA: Evitamos tamaños negativos para los círculos
        mapped["Tamaño_Marcador"] = mapped["Facturacion"].clip(lower=0)
        
        with tabs[0]:
            heat = px.density_map(mapped, lat="Latitud", lon="Longitud", z="Peso", radius=22, 
                                  center=center, zoom=3.2, map_style="carto-positron", 
                                  hover_name="Nombre_Cliente", height=500, color_continuous_scale="Turbo")
            heat.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False)
            st.plotly_chart(heat, use_container_width=True)
            
        with tabs[1]:
            # CAMBIAMOS size="Facturacion" por size="Tamaño_Marcador"
            points = px.scatter_map(mapped, lat="Latitud", lon="Longitud", size="Tamaño_Marcador", color="Facturacion", 
                                    hover_name="Nombre_Cliente", center=center, zoom=3.2, map_style="carto-positron", height=500)
            points.update_layout(margin=dict(l=0, r=0, t=0, b=0), coloraxis_showscale=False)
            st.plotly_chart(points, use_container_width=True)

        st.markdown("---")

        row1_col1, row1_col2 = st.columns(2)
        with row1_col1:
            st.subheader("Evolución mensual")
            evolucion = filtered.groupby("Mes", as_index=False)["Total S/IVA"].sum()
            orden_meses = ["Enero", "01-Enero", "Febrero", "02-Febrero", "Marzo", "03-Marzo", 
                           "Abril", "04-Abril", "Mayo", "05-Mayo", "Junio", "06-Junio", 
                           "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            fig_evo = px.area(evolucion, x="Mes", y="Total S/IVA", markers=True, category_orders={"Mes": orden_meses})
            fig_evo.update_traces(line_color='#1abc9c', fill='tozeroy', fillcolor='rgba(26, 188, 156, 0.2)')
            st.plotly_chart(fig_evo, use_container_width=True)

        with row1_col2:
            st.subheader("Top Proveedores")
            top_prov = filtered.groupby("Proveedor", as_index=False)["Total S/IVA"].sum().nlargest(10, "Total S/IVA")
            fig_prov = px.bar(top_prov, x="Total S/IVA", y="Proveedor", orientation='h', color_discrete_sequence=['#16a085'])
            fig_prov.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_prov, use_container_width=True)

        row2_col1, row2_col2 = st.columns(2)
        with row2_col1:
            st.subheader("Top Clientes")
            top_clientes = summary_map.nlargest(10, "Facturacion")
            fig_cli = px.bar(top_clientes, x="Facturacion", y="Nombre_Cliente", orientation='h', color_discrete_sequence=['#e67e22'])
            fig_cli.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_cli, use_container_width=True)

        with row2_col2:
            st.subheader("Ranking Vendedores")
            top_vend = filtered.groupby("Vendedor_Factura", as_index=False)["Total S/IVA"].sum().nlargest(10, "Total S/IVA")
            fig_vend = px.bar(top_vend, x="Total S/IVA", y="Vendedor_Factura", orientation='h', color_discrete_sequence=['#34495e'])
            fig_vend.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_vend, use_container_width=True)
    else:
        st.warning("No hay datos para mostrar con los filtros seleccionados.")
else:
    st.info("👋 ¡Hola! Haz clic en el botón azul 'Actualizar Datos desde Drive' en el menú de la izquierda para comenzar.")
