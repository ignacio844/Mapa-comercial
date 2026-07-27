import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import os
import io
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Mapa Comercial", page_icon="🗺️", layout="wide")
st.title("🗺️ Mapa Comercial - Ventas y Territorio")

# --- 1. CARGA DE DATOS (DEBE IR ARRIBA PARA QUE EL RESTO LO CONOZCA) ---
@st.cache_data
def cargar_datos():
    if not os.path.exists('Clientes_Geolocalizados.xlsx'):
        return pd.DataFrame()
    df = pd.read_excel('Clientes_Geolocalizados.xlsx')
    df = df.dropna(subset=['Latitud', 'Longitud'])
    df['TOTAL 2026'] = pd.to_numeric(df['TOTAL 2026'], errors='coerce').fillna(0)
    return df

# --- 2. FUNCIÓN PARA ACTUALIZAR DESDE DRIVE ---
# Recuerda volver a pegar tu FILE_ID real aquí
FILE_ID = '1Xe1iull8fs7xUZbajYSMjEhzbxet2fui' 

def actualizar_desde_drive():
    with st.spinner('Conectando a Google Drive y descargando datos...'):
        try:
            # 1. Conexión y descarga
            SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
            # Si el archivo existe localmente, lo usa. Si no, usa los secretos de la nube.
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
            
            # Cerramos el archivo para evitar el WinError 32
            fh.close()

            # 2. Procesamiento y Geolocalización simulada
            df_temp = pd.read_excel('Data_Descargada_Temp.xlsx', sheet_name='Clientes')
            
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

            df_temp['Prov_Limpia'] = df_temp['Provincia'].astype(str).str.upper().str.strip()
            reemplazos = {
                'CÓRDOBA': 'CORDOBA', 'ENTRE RÍOS': 'ENTRE RIOS', 'TUCUMÁN': 'TUCUMAN',
                'NEUQUÉN': 'NEUQUEN', 'RÍO NEGRO': 'RIO NEGRO', 'SANTE FE': 'SANTA FE',
                'NAN': 'BUENOS AIRES', 
            }
            for mal, bien in reemplazos.items():
                df_temp['Prov_Limpia'] = df_temp['Prov_Limpia'].str.replace(mal, bien)

            np.random.seed(42) 
            lats, lons = [], []
            for prov in df_temp['Prov_Limpia']:
                lat, lon, std = coordenadas_provincias.get(prov, coordenadas_provincias['BUENOS AIRES'])
                lats.append(np.random.normal(lat, std))
                lons.append(np.random.normal(lon, std))

            df_temp['Latitud'] = lats
            df_temp['Longitud'] = lons

            # 3. Guardar y limpiar
            df_temp.to_excel('Clientes_Geolocalizados.xlsx', index=False)
            if os.path.exists('Data_Descargada_Temp.xlsx'):
                os.remove('Data_Descargada_Temp.xlsx')
                
            # Limpiar caché ahora funcionará sin problemas
            cargar_datos.clear() 
            st.success("¡Base de datos actualizada con éxito!")
            
        except Exception as e:
            st.error(f"Ocurrió un error al actualizar: {e}")

# --- 3. INTERFAZ: BARRA LATERAL Y FILTROS ---
st.sidebar.header("Opciones")
if st.sidebar.button("🔄 Actualizar Datos desde Drive"):
    actualizar_desde_drive()

st.sidebar.markdown("---")
st.sidebar.header("Filtros")

# Llamamos a los datos
df = cargar_datos()

if not df.empty:
    provincias = df['Provincia'].dropna().unique().tolist()
    prov_sel = st.sidebar.multiselect("Provincia", provincias, default=provincias)

    vendedores = df['Vendedor'].dropna().unique().tolist()
    vend_sel = st.sidebar.multiselect("Vendedor", vendedores, default=vendedores)

    df_filtrado = df[(df['Provincia'].isin(prov_sel)) & (df['Vendedor'].isin(vend_sel))]

    # --- 4. KPIs ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Clientes Activos", f"{len(df_filtrado)}")
    col2.metric("Facturación Proyectada", f"${df_filtrado['TOTAL 2026'].sum():,.0f}")
    ticket_promedio = (df_filtrado['TOTAL 2026'].sum() / len(df_filtrado)) if len(df_filtrado) > 0 else 0
    col3.metric("Ticket Promedio", f"${ticket_promedio:,.0f}")

    st.markdown("---")

    # --- 5. MAPA ---
    if not df_filtrado.empty:
        max_peso = df_filtrado['TOTAL 2026'].max() if df_filtrado['TOTAL 2026'].max() > 0 else 1
        df_filtrado_mapa = df_filtrado.copy()
        df_filtrado_mapa['Peso_Scaled'] = df_filtrado_mapa['TOTAL 2026'] / max_peso

        fig_mapa = go.Figure()
        fig_mapa.add_trace(go.Densitymapbox(
            lat=df_filtrado_mapa['Latitud'], lon=df_filtrado_mapa['Longitud'], z=df_filtrado_mapa['Peso_Scaled'],
            radius=15, colorscale='Hot', opacity=0.6, showscale=False
        ))
        fig_mapa.add_trace(go.Scattermapbox(
            lat=df_filtrado_mapa['Latitud'], lon=df_filtrado_mapa['Longitud'], mode='markers',
            marker=go.scattermapbox.Marker(size=6, color='teal', opacity=0.8),
            text=df_filtrado_mapa['Nombre_Cliente'], hoverinfo='text'
        ))
        fig_mapa.update_layout(
            mapbox_style="carto-positron", mapbox_center_lon=-64, mapbox_center_lat=-38,
            mapbox_zoom=3.5, margin={"r":0,"t":0,"l":0,"b":0}, height=500
        )
        st.plotly_chart(fig_mapa, use_container_width=True)

        st.markdown("---")
        
        # --- 6. GRÁFICOS INFERIORES ---
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.subheader("Top 10 Clientes")
            top_clientes = df_filtrado.nlargest(10, 'TOTAL 2026')
            fig_clientes = px.bar(top_clientes, x='TOTAL 2026', y='Nombre_Cliente', orientation='h', color_discrete_sequence=['#e67e22'])
            fig_clientes.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_clientes, use_container_width=True)
            
        with col_chart2:
            st.subheader("Ranking Vendedores")
            ranking_vend = df_filtrado.groupby('Vendedor')['TOTAL 2026'].sum().reset_index().nlargest(10, 'TOTAL 2026')
            fig_vend = px.bar(ranking_vend, x='TOTAL 2026', y='Vendedor', orientation='h', color_discrete_sequence=['#2980b9'])
            fig_vend.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_vend, use_container_width=True)
    else:
        st.warning("No hay datos para mostrar con los filtros seleccionados.")
else:
    st.info("No hay datos disponibles. Haz clic en 'Actualizar Datos desde Drive' para comenzar.")