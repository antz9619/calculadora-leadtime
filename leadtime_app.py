import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
from collections import Counter
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as XLImage
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import PatternFill
from openpyxl.chart.label import DataLabelList
import numpy as np
import unicodedata

# --- FERIADOS 2025 ---
feriados_2025 = [
    "2025-01-01", "2025-03-03", "2025-03-24", "2025-04-02",
    "2025-04-17", "2025-04-18", "2025-05-01", "2025-05-25",
    "2025-06-20", "2025-07-09", "2025-12-08", "2025-12-25"
]
feriados_set = set(pd.to_datetime(feriados_2025).date)

def es_feriado(fecha):
    return fecha in feriados_set

def es_dia_habil(fecha):
    if fecha.weekday() >= 5:  # 5=Sab, 6=Dom
        return False
    if es_feriado(fecha):
        return False
    return True

def calcular_dias_habiles(fecha_inicio, fecha_fin):
    if pd.isna(fecha_inicio) or pd.isna(fecha_fin):
        return None
    if fecha_inicio > fecha_fin:
        return 0
    fecha_inicio = fecha_inicio.date()
    fecha_fin = fecha_fin.date()
    dias = 0
    current = fecha_inicio
    while current <= fecha_fin:
        if es_dia_habil(current):
            dias += 1
        current += timedelta(days=1)
    return dias

# --- LISTA DE LOCALIDADES AMBA ---
amba_localidades = [
    "CIUDAD AUTONOMA BUENOS AIRES",
    "AVELLANEDA",
    "LANUS",
    "LOMAS DE ZAMORA",
    "LA MATANZA",
    "MORON",
    "SAN MARTIN",
    "VICENTE LOPEZ",
    "SAN ISIDRO",
    "TRES DE FEBRERO",
    "MORENO",
    "HURLINGHAM",
    "ITUZAINGO",
    "BERAZATEGUI",
    "FLORENCIO VARELA",
    "QUILMES",
    "ALMIRANTE BROWN",
    "ESTEBAN ECHEVERRIA",
    "EZEIZA",
    "SAN FERNANDO",
    "TIGRE",
    "SAN MIGUEL",
    "MALVINAS ARGENTINAS",
    "JOSE C. PAZ",
    "PILAR",
    "ESCOBAR",
    "MERLO",
    "MARCOS PAZ",
    "GENERAL RODRIGUEZ",
    "PRESIDENTE PERON",
    "CAÑUELAS",
    "SAN VICENTE",
    "BRANDSEN",
    "BERISSO",
    "ENSENADA",
    "LA PLATA",
    "MUNRO",
    "SAAVEDRA",
    "FLORES",
    "ALMAGRO",
    "VILLA URQUIZA",
    "COLEGIALES",
    "PALERMO",
    "RECOLETA",
    "BELGRANO",
    "NUÑEZ",
    "CABALLITO",
    "BOEDO",
    "SAN TELMO",
    "CONSTITUCION",
    "RETIRO",
    "SAN CRISTOBAL",
    "BALVANERA",
    "MONTSERRAT"
]

# --- EXCEPCIONES: localidades que NO son AMBA aunque tengan nombre similar ---
excepciones_amba = [
    "SAN MARTIN, SANTA FE",
    "SAN MARTIN, MENDOZA",
    "SAN MARTIN, SAN JUAN",
    "SAN MARTIN, CORRIENTES",
    "SAN MARTIN, ENTRE RIOS",
    "VILLA LIB. GENERAL SAN MARTIN",
    "GENERAL SAN MARTIN",
    "SAN MARTIN DE LOS ANDES",
    "SAN MARTIN DE LA VEGA",
    "TANDIL, BUENOS AIRES",
    "MAR DEL PLATA, BUENOS AIRES",
    "BAHIA BLANCA, BUENOS AIRES",
    "NECOCHEA, BUENOS AIRES",
    "OLAVARRIA, BUENOS AIRES",
    "AZUL, BUENOS AIRES"
]

def determinar_zona(localidad_destino):
    if pd.isna(localidad_destino):
        return "INTERIOR"
    
    # Normalizar: mayúsculas, sin espacios extra, sin tildes
    localidad_destino = str(localidad_destino).upper().strip()
    localidad_destino = ''.join(
        c for c in unicodedata.normalize('NFD', localidad_destino)
        if unicodedata.category(c) != 'Mn'
    )
    
    # 1. Verificar excepciones primero
    for excepcion in excepciones_amba:
        if excepcion in localidad_destino:
            return "INTERIOR"
    
    # 2. Coincidencia exacta
    for localidad_amba in amba_localidades:
        if localidad_amba == localidad_destino:
            return "AMBA"
    
    # 3. Coincidencia parcial segura: empieza con nombre de AMBA + coma o espacio
    for localidad_amba in amba_localidades:
        if localidad_destino.startswith(localidad_amba):
            resto = localidad_destino[len(localidad_amba):].strip()
            if resto.startswith(",") or resto.startswith(" ") or len(resto) == 0:
                return "AMBA"
    
    # 4. Palabras clave seguras
    palabras_clave_amba = ["CAPITAL FEDERAL", "C.A.B.A.", "CABA", "CIUDAD AUTONOMA"]
    for palabra in palabras_clave_amba:
        if palabra in localidad_destino:
            return "AMBA"
    
    return "INTERIOR"

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Calculadora de Lead Time", layout="wide")

st.title("📊 Calculadora de Lead Time")
st.markdown("Sube tu reporte diario y obtén estadísticas + PPT listo para presentar.")

uploaded_file = st.file_uploader("📂 Sube tu archivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Leer Excel
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Prueba")
    except:
        # Intentar con la primera hoja si no encuentra "Prueba"
        df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Renombrar columnas si es necesario
    if 'Localidad destino' in df.columns:
        df['Loc'] = df['Localidad destino']
    
    # Convertir columnas de fecha
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df['Fecha último estado'] = pd.to_datetime(df['Fecha último estado'], errors='coerce')
    
    # Determinar ZONA (AMBA o INTERIOR)
    df['ZONA'] = df['Loc'].apply(determinar_zona)
    
    # Determinar días prometidos según ZONA, pero con excepción para Delivery Hero Riders
    def determinar_dias_prometidos(row):
        # Caso especial: DELIVERY HERO E-COMMERCE S.A. + RIDERS
        if row.get('Cliente', '') == "DELIVERY HERO E-COMMERCE S.A." and row.get('Subcuenta', '') == "RIDERS":
            return 3  # Siempre 3 días, sin importar zona
        else:
            # Comportamiento normal
            return 3 if row['ZONA'] == "AMBA" else 5

    df['Días Prometidos'] = df.apply(determinar_dias_prometidos, axis=1)
    
    # --- CORRECCIÓN: CÁLCULO DE LEAD TIME PARA PENDIENTES ---
    def calcular_lead_time(row):
        estado = str(row['Estado']).lower()
        ed = str(row.get('ED', '')).upper() if 'ED' in df.columns else 'SI'
        
        # Determinar si el pedido está entregado
        entregado = (
            (ed == "NO" and "esperando retiro" in estado) or 
            "entregada" in estado
        )
        
        if entregado:
            # Para pedidos ENTREGADOS: calcular desde creación hasta último estado
            lead_time = calcular_dias_habiles(row['Fecha'], row['Fecha último estado'])
        else:
            # Para pedidos PENDIENTES: calcular desde creación hasta HOY
            lead_time = calcular_dias_habiles(row['Fecha'], datetime.now())
        
        # Aplicar día de gracia para Delivery Hero Riders
        if row.get('Cliente', '') == "DELIVERY HERO E-COMMERCE S.A." and row.get('Subcuenta', '') == "RIDERS":
            if pd.notna(lead_time) and lead_time > 0:
                lead_time = max(0, lead_time - 1)  # No permitir valores negativos
        
        return lead_time

    df['Lead Time'] = df.apply(calcular_lead_time, axis=1)
    
    # Columna para identificar si se aplicó día de gracia
    df['Día de Gracia Aplicado'] = df.apply(
        lambda row: "Sí" if row.get('Cliente', '') == "DELIVERY HERO E-COMMERCE S.A." and row.get('Subcuenta', '') == "RIDERS" else "No",
        axis=1
    )
    
    # --- CORRECCIÓN: CÁLCULO DE CUMPLIMIENTO PARA PENDIENTES ---
    def determinar_cumplimiento(row):
        estado = str(row['Estado']).lower()
        ed = str(row.get('ED', '')).upper() if 'ED' in df.columns else 'SI'
        condicion_venta = str(row.get('Condición de venta', '')).upper() if 'Condición de venta' in df.columns else ''
        
        # Si ED es "NO" y estado es "esperando retiro"
        if ed == "NO" and "esperando retiro" in estado:
            if pd.notna(row['Lead Time']) and row['Lead Time'] <= row['Días Prometidos']:
                # Si condición de venta es PD, marcar como entregada pero con nota
                if condicion_venta == "PD":
                    return "Entregada - En Tiempo (PD: Pago Pendiente)"
                else:
                    return "Entregada - En Tiempo"
            else:
                if condicion_venta == "PD":
                    return "Entregada - Fuera de Tiempo (PD: Pago Pendiente)"
                else:
                    return "Entregada - Fuera de Tiempo"
        
        # Para ED "SI" o cualquier otro caso, considerar "Entregada" como entregado
        elif "entregada" in estado:
            if pd.notna(row['Lead Time']) and row['Lead Time'] <= row['Días Prometidos']:
                return "Entregada - En Tiempo"
            else:
                return "Entregada - Fuera de Tiempo"
        
        else:
            # Para pendientes: usar el lead time calculado (que ahora es hasta hoy)
            if pd.notna(row['Lead Time']):
                if row['Lead Time'] < row['Días Prometidos']:
                    return "Pendiente - En Tiempo"
                elif row['Lead Time'] == row['Días Prometidos']:
                    return "Pendiente - Último Día"
                else:
                    return "Pendiente - Fuera de Tiempo"
            else:
                return "Pendiente - Sin datos"
    
    df['Cumplimiento'] = df.apply(determinar_cumplimiento, axis=1)
    
    # Calcular días restantes para pendientes en tiempo (CORREGIDO)
    def calcular_dias_restantes(row):
        cumplimiento = str(row['Cumplimiento'])
        
        if "Pendiente" in cumplimiento and "Fuera" not in cumplimiento and "Sin datos" not in cumplimiento:
            restantes = row['Días Prometidos'] - row['Lead Time']
            return f"{int(restantes)} días restantes" if restantes > 0 else "Vence hoy"
        return ""
    
    df['Días Restantes'] = df.apply(calcular_dias_restantes, axis=1)
    
    # --- ALERTA DE DEVOLUCIÓN ---
    # Para pedidos con ED="NO" y estado "Esperando retiro" con más de 15 días hábiles desde la fecha último estado
    def alerta_devolucion(row):
        estado = str(row['Estado']).lower()
        ed = str(row.get('ED', '')).upper() if 'ED' in df.columns else 'SI'
        fecha_ultimo_estado = row['Fecha último estado']
        
        if ed == "NO" and "esperando retiro" in estado and pd.notna(fecha_ultimo_estado):
            dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, datetime.now())
            if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 15:
                return "Sugerir devolución"
        return ""
    
    df['Alerta Devolución'] = df.apply(alerta_devolucion, axis=1)
    
    # --- ALERTA DE REDESPACHO ---
    def alerta_redespacho(row):
        estado = str(row['Estado']).lower()
        fecha_ultimo_estado = row['Fecha último estado']
        
        if "redespacho" in estado and pd.notna(fecha_ultimo_estado):
            dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, datetime.now())
            if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 2:  # 2 días hábiles = 48 horas
                return "Redespacho demorado"
        return ""
    
    df['Alerta Redespacho'] = df.apply(alerta_redespacho, axis=1)
    
    # --- ALERTA PENDIENTE FUERA DE TIEMPO ---
    # Para pedidos con estado "Pendiente - Fuera de Tiempo"
    def alerta_pendiente_fuera_tiempo(row):
        cumplimiento = str(row['Cumplimiento'])
        
        if cumplimiento == "Pendiente - Fuera de Tiempo":
            return "Fuera de tiempo crítico"
        return ""
    
    df['Alerta Pendiente Fuera Tiempo'] = df.apply(alerta_pendiente_fuera_tiempo, axis=1)
    
    # --- ALERTA DE PAGO PENDIENTE ---
    # Para pedidos con Condición de venta = "PD" y estado "Esperando retiro" con más de 5 días hábiles desde la fecha último estado
    def alerta_pago_pendiente(row):
        estado = str(row['Estado']).lower()
        condicion_venta = str(row.get('Condición de venta', '')).upper() if 'Condición de venta' in df.columns else ''
        fecha_ultimo_estado = row['Fecha último estado']
        
        if condicion_venta == "PD" and "esperando retiro" in estado and pd.notna(fecha_ultimo_estado):
            dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, datetime.now())
            if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 5:
                return "Pago pendiente demorado"
        return ""
    
    df['Alerta Pago Pendiente'] = df.apply(alerta_pago_pendiente, axis=1)
    
    # --- FILTROS ---
    st.sidebar.header("🔍 Filtros")

    # Filtro por Cliente (con verificación)
    if 'Cliente' in df.columns:
        clientes = sorted(df['Cliente'].dropna().unique())
        cliente_seleccionado = st.sidebar.selectbox("Cliente", ["Todos"] + clientes)
    else:
        st.error("❌ La columna 'Cliente' no existe en el archivo. Verifica el nombre de las columnas.")
        st.stop()

    # Filtro por Subcuenta (con verificación)
    if 'Subcuenta' in df.columns:
        subcuentas = sorted(df['Subcuenta'].dropna().unique())
        subcuenta_seleccionada = st.sidebar.selectbox("Subcuenta", ["Todas"] + subcuentas)
    else:
        st.error("❌ La columna 'Subcuenta' no existe en el archivo. Verifica el nombre de las columnas.")
        st.stop()

    # Filtro por Agencia origen (con verificación)
    if 'Agencia origen' in df.columns:
        agencias_origen = sorted(df['Agencia origen'].dropna().unique())
        agencia_origen_seleccionada = st.sidebar.selectbox("Agencia origen", ["Todas"] + agencias_origen)
    else:
        st.warning("⚠️ La columna 'Agencia origen' no existe en el archivo. Se omitirá este filtro.")
        agencia_origen_seleccionada = "Todas"

    # Filtro por Agencia destino (con verificación)
    if 'Agencia destino' in df.columns:
        agencias = sorted(df['Agencia destino'].dropna().unique())
        agencia_seleccionada = st.sidebar.selectbox("Agencia destino", ["Todas"] + agencias)
    else:
        st.error("❌ La columna 'Agencia destino' no existe en el archivo. Verifica el nombre de las columnas.")
        st.stop()

    # Filtro por ED (con verificación)
    if 'ED' in df.columns:
        ed_opciones = sorted(df['ED'].dropna().unique())
        ed_seleccionada = st.sidebar.selectbox("Entrega a Domicilio (ED)", ["Todas"] + ed_opciones)
    else:
        st.warning("⚠️ La columna 'ED' no existe en el archivo. Se omitirá este filtro.")
        ed_seleccionada = "Todas"

    # Filtro por Condición de venta (con verificación)
    if 'Condición de venta' in df.columns:
        condiciones_venta = sorted(df['Condición de venta'].dropna().unique())
        condicion_venta_seleccionada = st.sidebar.selectbox("Condición de venta", ["Todas"] + condiciones_venta)
    else:
        st.warning("⚠️ La columna 'Condición de venta' no existe en el archivo. Se omitirá este filtro.")
        condicion_venta_seleccionada = "Todas"

    # Aplicar filtros
    if cliente_seleccionado != "Todos":
        df = df[df['Cliente'] == cliente_seleccionado]

    if subcuenta_seleccionada != "Todas":
        df = df[df['Subcuenta'] == subcuenta_seleccionada]

    if 'Agencia origen' in df.columns and agencia_origen_seleccionada != "Todas":
        df = df[df['Agencia origen'] == agencia_origen_seleccionada]

    if agencia_seleccionada != "Todas":
        df = df[df['Agencia destino'] == agencia_seleccionada]

    if 'ED' in df.columns and 'ed_seleccionada' in locals() and ed_seleccionada != "Todas":
        df = df[df['ED'] == ed_seleccionada]

    if 'Condición de venta' in df.columns and condicion_venta_seleccionada != "Todas":
        df = df[df['Condición de venta'] == condicion_venta_seleccionada]
    
    # --- ESTADÍSTICAS ---
    st.header("📈 Estadísticas")
    
    total_pedidos = df.shape[0]
    entregados = df[df['Cumplimiento'].str.startswith("Entregada")].shape[0]
    pendientes = total_pedidos - entregados
    
    # Clasificación detallada
    en_tiempo = df[df['Cumplimiento'] == "Entregada - En Tiempo"].shape[0]
    en_tiempo_pd = df[df['Cumplimiento'] == "Entregada - En Tiempo (PD: Pago Pendiente)"].shape[0]
    fuera_tiempo = df[df['Cumplimiento'] == "Entregada - Fuera de Tiempo"].shape[0]
    fuera_tiempo_pd = df[df['Cumplimiento'] == "Entregada - Fuera de Tiempo (PD: Pago Pendiente)"].shape[0]
    pendiente_en_tiempo = df[df['Cumplimiento'] == "Pendiente - En Tiempo"].shape[0]
    pendiente_fuera_tiempo = df[df['Cumplimiento'] == "Pendiente - Fuera de Tiempo"].shape[0]
    pendiente_ultimo_dia = df[df['Cumplimiento'] == "Pendiente - Último Día"].shape[0]
    pendiente_sin_datos = df[df['Cumplimiento'] == "Pendiente - Sin datos"].shape[0]
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Total Pedidos", total_pedidos)
    with col2:
        st.metric("✅ Entregados", entregados)
    with col3:
        st.metric("⏳ Pendientes", pendientes)
    with col4:
        if entregados > 0:
            porcentaje = ((en_tiempo + en_tiempo_pd) / entregados) * 100
            st.metric("🎯 % Cumplimiento Entregados", f"{porcentaje:.1f}%")
        else:
            st.metric("🎯 % Cumplimiento", "0%")
    
    # Gráfico de torta - Cumplimiento general
    cumplimiento_labels = [
        "Entregada - En Tiempo", 
        "Entregada - En Tiempo (PD)",
        "Entregada - Fuera de Tiempo", 
        "Entregada - Fuera de Tiempo (PD)",
        "Pendiente - En Tiempo", 
        "Pendiente - Último Día",
        "Pendiente - Fuera de Tiempo",
        "Pendiente - Sin datos"
    ]
    
    cumplimiento_values = [
        en_tiempo,
        en_tiempo_pd,
        fuera_tiempo,
        fuera_tiempo_pd,
        pendiente_en_tiempo, 
        pendiente_ultimo_dia,
        pendiente_fuera_tiempo,
        pendiente_sin_datos
    ]
    
    # Colores en orden correcto
    colores = ["#28a745", "#2ecc71", "#dc3545", "#e74c3c", "#ffc107", "#fd7e14", "#6c757d", "#17a2b8"]
    
    fig1 = px.pie(
        names=cumplimiento_labels,
        values=cumplimiento_values,
        title="Distribución de Cumplimiento General",
        color=cumplimiento_labels,
        color_discrete_map={
            "Entregada - En Tiempo": "#28a745",
            "Entregada - En Tiempo (PD)": "#2ecc71",
            "Entregada - Fuera de Tiempo": "#dc3545",
            "Entregada - Fuera de Tiempo (PD)": "#e74c3c",
            "Pendiente - En Tiempo": "#ffc107",
            "Pendiente - Último Día": "#fd7e14",
            "Pendiente - Fuera de Tiempo": "#6c757d",
            "Pendiente - Sin datos": "#17a2b8"
        },
        hole=0.4
    )
    fig1.update_traces(textinfo='percent+value', textposition='inside')
    st.plotly_chart(fig1, use_container_width=True)
    
    # Gráfico por Localidad (Top 10 con más fuera de tiempo) - BARRAS HORIZONTALES
    fuera_tiempo_df = df[df['Cumplimiento'].str.contains("Fuera", na=False)]
    if not fuera_tiempo_df.empty:
        top_localidades = fuera_tiempo_df['Loc'].value_counts().head(10)
        
        # Crear gráfico de barras horizontales
        fig2 = px.bar(
            y=top_localidades.index,
            x=top_localidades.values,
            labels={'x': 'Pedidos Fuera de Tiempo', 'y': 'Localidad'},
            title="Top 10 Localidades con Más Pedidos Fuera de Tiempo",
            color_discrete_sequence=["#dc3545"],
            orientation='h'  # Barras horizontales
        )
        fig2.update_traces(texttemplate='%{x}', textposition='outside')
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("No hay pedidos fuera de tiempo para mostrar.")
    
    # Gráfico por Producto - BARRAS HORIZONTALES
    if 'Producto' in df.columns:
        servicio_stats = df.groupby('Producto')['Cumplimiento'].value_counts().unstack(fill_value=0)
        
        # Asegurar que todas las categorías estén presentes
        for label in cumplimiento_labels:
            if label not in servicio_stats.columns:
                servicio_stats[label] = 0
        
        # Reordenar columnas según el orden deseado
        servicio_stats = servicio_stats[cumplimiento_labels]
        
        # Calcular porcentajes por servicio
        servicio_totales = servicio_stats.sum(axis=1)
        servicio_porcentajes = servicio_stats.div(servicio_totales, axis=0) * 100
        
        # Crear texto para las barras (valor y porcentaje)
        servicio_texto = servicio_stats.copy().astype(str)
        for col in servicio_stats.columns:
            servicio_texto[col] = servicio_stats[col].astype(str) + " (" + servicio_porcentajes[col].round(1).astype(str) + "%)"
        
        # Crear gráfico de barras horizontales apiladas
        fig3 = go.Figure()
        
        for i, categoria in enumerate(cumplimiento_labels):
            fig3.add_trace(go.Bar(
                name=categoria,
                y=servicio_stats.index,  # Productos en el eje Y
                x=servicio_stats[categoria],  # Cantidad en el eje X
                text=servicio_texto[categoria],
                textposition='auto',
                marker_color=colores[i],
                orientation='h'  # Orientación horizontal
            ))
        
        fig3.update_layout(
            title="Cumplimiento por Producto",
            barmode='stack',
            yaxis_title="Producto",
            xaxis_title="Cantidad de Pedidos",
            height=600  # Altura fija para mejor visualización
        )
        
        st.plotly_chart(fig3, use_container_width=True)
    
    # --- FUNCION AUXILIAR PARA GENERAR EXCEL ---
    def generar_excel_desde_df(df, nombre_hoja="Datos"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
        output.seek(0)
        return output

    # --- ALERTAS DE DEVOLUCIÓN ---
    alertas_devolucion = df[df['Alerta Devolución'] == "Sugerir devolución"]
    if not alertas_devolucion.empty:
        st.header("🚨 Alertas de Devolución")
        st.write("Los siguientes pedidos están en estado 'Esperando retiro' por más de 15 días hábiles. Se sugiere devolución al remitente.")
        
        columnas_alerta = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Fecha último estado', 'Alerta Devolución']
        df_alerta = alertas_devolucion[columnas_alerta]
        
        st.dataframe(df_alerta)
        
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Devolución")
        st.download_button(
            label="📥 Descargar Alertas Devolución (Excel)",
            data=excel_data,
            file_name="Alertas_Devolucion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # --- ALERTAS DE REDESPACHO ---
    alertas_redespacho = df[df['Alerta Redespacho'] == "Redespacho demorado"]
    if not alertas_redespacho.empty:
        st.header("🚨 Alertas de Redespacho Demorado")
        st.write("Los siguientes pedidos están en estado 'Redespacho' por más de 48 horas hábiles.")
        
        columnas_alerta = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Fecha último estado', 'Alerta Redespacho']
        df_alerta = alertas_redespacho[columnas_alerta]
        
        st.dataframe(df_alerta)
        
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Redespacho")
        st.download_button(
            label="📥 Descargar Alertas Redespacho (Excel)",
            data=excel_data,
            file_name="Alertas_Redespacho.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # --- ALERTAS DE PENDIENTE FUERA DE TIEMPO ---
    alertas_pendiente_fuera_tiempo = df[df['Alerta Pendiente Fuera Tiempo'] == "Fuera de tiempo crítico"]
    if not alertas_pendiente_fuera_tiempo.empty:
        st.header("🚨 Alertas de Pendiente Fuera de Tiempo")
        st.write("Los siguientes pedidos están pendientes y fuera del tiempo de entrega prometido.")
        
        columnas_alerta = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Fecha último estado', 'Días Prometidos', 'Lead Time', 'Alerta Pendiente Fuera Tiempo']
        df_alerta = alertas_pendiente_fuera_tiempo[columnas_alerta]
        
        st.dataframe(df_alerta)
        
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Pendiente Fuera Tiempo")
        st.download_button(
            label="📥 Descargar Alertas Pendiente Fuera Tiempo (Excel)",
            data=excel_data,
            file_name="Alertas_Pendiente_Fuera_Tiempo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- ALERTAS DE PAGO PENDIENTE ---
    alertas_pago_pendiente = df[df['Alerta Pago Pendiente'] == "Pago pendiente demorado"]
    if not alertas_pago_pendiente.empty:
        st.header("🚨 Alertas de Pago Pendiente Demorado")
        st.write("Los siguientes pedidos están en estado 'Esperando retiro' con condición de venta PD por más de 5 días hábiles. Se sugiere gestionar el cobro.")
        
        columnas_alerta = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Fecha último estado', 'Condición de venta', 'Alerta Pago Pendiente']
        df_alerta = alertas_pago_pendiente[columnas_alerta]
        
        st.dataframe(df_alerta)
        
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Pago Pendiente")
        st.download_button(
            label="📥 Descargar Alertas Pago Pendiente (Excel)",
            data=excel_data,
            file_name="Alertas_Pago_Pendiente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- DESCARGA COMBINADA DE TODAS LAS ALERTAS ---
    st.header("📥 Descarga Combinada de Todas las Alertas")

    todas_alertas = df[
        (df['Alerta Devolución'] == "Sugerir devolución") |
        (df['Alerta Redespacho'] == "Redespacho demorado") |
        (df['Alerta Pendiente Fuera Tiempo'] == "Fuera de tiempo crítico") |
        (df['Alerta Pago Pendiente'] == "Pago pendiente demorado")
    ]

    if not todas_alertas.empty:
        columnas_todas = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Estado', 'Fecha último estado', 
                          'Alerta Devolución', 'Alerta Redespacho', 'Alerta Pendiente Fuera Tiempo', 'Alerta Pago Pendiente']
        df_todas = todas_alertas[columnas_todas]
        
        st.dataframe(df_todas)
        
        excel_todas = generar_excel_desde_df(df_todas, "Todas las Alertas")
        st.download_button(
            label="📥 Descargar Todas las Alertas (Excel)",
            data=excel_todas,
            file_name="Todas_Alertas_Combinadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("✅ No hay alertas activas en este momento.")

    # --- DESCARGAS ---
    st.header("📥 Descargas Generales")
    
    # Preparar Excel con gráficos
    output_excel = io.BytesIO()

    # Crear datos para el gráfico de estadísticas
    stats_data = {
        "Métrica": [
            "Total Pedidos", "Entregados", "Pendientes",
            "Entregada - En Tiempo", "Entregada - En Tiempo (PD)",
            "Entregada - Fuera de Tiempo", "Entregada - Fuera de Tiempo (PD)",
            "Pendiente - En Tiempo", "Pendiente - Último Día",
            "Pendiente - Fuera de Tiempo", "Pendiente - Sin datos",
            "% Cumplimiento (solo entregados)"
        ],
        "Valor": [
            total_pedidos, entregados, pendientes,
            en_tiempo, en_tiempo_pd,
            fuera_tiempo, fuera_tiempo_pd,
            pendiente_en_tiempo, pendiente_ultimo_dia,
            pendiente_fuera_tiempo, pendiente_sin_datos,
            f"{((en_tiempo + en_tiempo_pd)/entregados*100):.2f}%" if entregados > 0 else "0%"
        ]
    }
    stats_df = pd.DataFrame(stats_data)
    
    # Guardar en Excel
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Base", index=False)
        
        # Hoja de estadísticas
        stats_df.to_excel(writer, sheet_name="Estadísticas", index=False)
        
        # Obtener la hoja de trabajo
        workbook = writer.book
        worksheet = writer.sheets["Estadísticas"]
        
        # Crear datos para el gráfico de torta
        pie_data = [
            ["Categoría", "Cantidad"],
            ["Entregada - En Tiempo", en_tiempo],
            ["Entregada - En Tiempo (PD)", en_tiempo_pd],
            ["Entregada - Fuera de Tiempo", fuera_tiempo],
            ["Entregada - Fuera de Tiempo (PD)", fuera_tiempo_pd],
            ["Pendiente - En Tiempo", pendiente_en_tiempo],
            ["Pendiente - Último Día", pendiente_ultimo_dia],
            ["Pendiente - Fuera de Tiempo", pendiente_fuera_tiempo],
            ["Pendiente - Sin datos", pendiente_sin_datos]
        ]
        
        # Escribir datos para el gráfico de torta
        for i, row in enumerate(pie_data, start=15):
            for j, value in enumerate(row, start=6):
                worksheet.cell(row=i, column=j, value=value)
        
        # Crear gráfico de torta
        pie_chart = PieChart()
        pie_chart.title = "Distribución de Cumplimiento"
        
        # Referencias a los datos
        labels = Reference(worksheet, min_col=6, min_row=16, max_row=24)
        data = Reference(worksheet, min_col=7, min_row=15, max_row=24)
        
        # Añadir datos al gráfico
        pie_chart.add_data(data, titles_from_data=True)
        pie_chart.set_categories(labels)
        
        # Estilo del gráfico
        pie_chart.style = 10  # Estilo predefinido
        
        # Añadir etiquetas de datos
        pie_chart.dataLabels = DataLabelList()
        pie_chart.dataLabels.showPercent = True
        pie_chart.dataLabels.showVal = True
        pie_chart.dataLabels.showCatName = True
        
        # Colores personalizados
        colors = ['28a745', '2ecc71', 'dc3545', 'e74c3c', 'ffc107', 'fd7e14', '6c757d', '17a2b8']
        for i, point in enumerate(pie_chart.series[0].data_points):
            point.graphicalProperties.solidFill = colors[i]
        
        # Añadir gráfico a la hoja
        worksheet.add_chart(pie_chart, "D15")

    output_excel.seek(0)
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        st.download_button(
            label="📥 Descargar Excel Actualizado (Completo)",
            data=output_excel,
            file_name="Reporte_LeadTime_Actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # --- GENERAR POWERPOINT ---
    def crear_pptx():
        prs = Presentation()
        
        # Slide 1: Título
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = "Reporte de Cumplimiento de Entregas"
        subtitle.text = "Lead Time - PedidosYa\nGenerado automáticamente"
        
        # Slide 2: Resumen Ejecutivo
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Resumen Ejecutivo"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Métricas Clave:"
        p.font.bold = True
        p.font.size = Pt(20)
        
        metrics = [
            f"• Total de pedidos: {total_pedidos}",
            f"• Entregados: {entregados} ({(entregados/total_pedidos*100):.1f}%)",
            f"• Pendientes: {pendientes} ({(pendientes/total_pedidos*100):.1f}%)",
            f"• Entregada - En Tiempo: {en_tiempo}",
            f"• Entregada - En Tiempo (PD): {en_tiempo_pd}",
            f"• Entregada - Fuera de Tiempo: {fuera_tiempo}",
            f"• Entregada - Fuera de Tiempo (PD): {fuera_tiempo_pd}",
            f"• Pendiente - En Tiempo: {pendiente_en_tiempo}",
            f"• Pendiente - Último Día: {pendiente_ultimo_dia}",
            f"• Pendiente - Fuera de Tiempo: {pendiente_fuera_tiempo}",
            f"• Pendiente - Sin datos: {pendiente_sin_datos}"
        ]
        
        if entregados > 0:
            cumplidos = en_tiempo + en_tiempo_pd
            metrics.append(f"• % Cumplimiento (solo entregados): {(cumplidos/entregados*100):.2f}%")
        
        for metric in metrics:
            p = tf.add_paragraph()
            p.text = metric
            p.font.size = Pt(16)
            if "Entregada - En Tiempo" in metric and "(PD)" not in metric:
                p.font.color.rgb = RGBColor(40, 167, 69)
            elif "Entregada - En Tiempo (PD)" in metric:
                p.font.color.rgb = RGBColor(46, 204, 113)
            elif "Entregada - Fuera de Tiempo" in metric and "(PD)" not in metric:
                p.font.color.rgb = RGBColor(220, 53, 69)
            elif "Entregada - Fuera de Tiempo (PD)" in metric:
                p.font.color.rgb = RGBColor(231, 76, 60)
            elif "Pendiente - En Tiempo" in metric:
                p.font.color.rgb = RGBColor(255, 193, 7)
            elif "Pendiente - Último Día" in metric:
                p.font.color.rgb = RGBColor(253, 126, 20)
            elif "Pendiente - Fuera de Tiempo" in metric:
                p.font.color.rgb = RGBColor(108, 117, 125)
            elif "Pendiente - Sin datos" in metric:
                p.font.color.rgb = RGBColor(23, 162, 184)
        
        # Slide 3: Gráfico de Cumplimiento
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Distribución de Cumplimiento"
        
        img_buffer = io.BytesIO()
        fig1.write_image(img_buffer, format="png", width=800, height=500, engine="kaleido")
        img_buffer.seek(0)
        left = Inches(0.5)
        top = Inches(1.5)
        slide.shapes.add_picture(img_buffer, left, top, width=Inches(9))
        
        # Slide 4: Top Localidades (si existe)
        if not fuera_tiempo_df.empty:
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            title = slide.shapes.title
            title.text = "Top 10 Localidades con Más Fuera de Tiempo"
            
            img_buffer2 = io.BytesIO()
            fig2.write_image(img_buffer2, format="png", width=800, height=500, engine="kaleido")
            img_buffer2.seek(0)
            left = Inches(0.5)
            top = Inches(1.5)
            slide.shapes.add_picture(img_buffer2, left, top, width=Inches(9))
        
        # Slide 5: Por Producto
        if 'Producto' in df.columns:
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            title = slide.shapes.title
            title.text = "Cumplimiento por Producto"
            
            fig3_pptx = go.Figure()
            
            for i, categoria in enumerate(cumplimiento_labels):
                fig3_pptx.add_trace(go.Bar(
                    name=categoria,
                    y=servicio_stats.index,
                    x=servicio_stats[categoria],
                    text=servicio_texto[categoria],
                    textposition='auto',
                    marker_color=colores[i],
                    orientation='h',
                    marker_line=dict(width=1, color='black')
                ))
            
            fig3_pptx.update_layout(
                title="Cumplimiento por Producto",
                barmode='stack',
                yaxis_title="Producto",
                xaxis_title="Cantidad de Pedidos",
                paper_bgcolor='white',
                plot_bgcolor='white',
                height=600
            )
            
            img_buffer3 = io.BytesIO()
            fig3_pptx.write_image(img_buffer3, format="png", width=800, height=600, engine="kaleido")
            img_buffer3.seek(0)
            left = Inches(0.5)
            top = Inches(1.5)
            slide.shapes.add_picture(img_buffer3, left, top, width=Inches(9), height=Inches(6))
        
        # Slide 6: Recomendaciones
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Recomendaciones Estratégicas"
        
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = "Acciones Recomendadas:"
        p.font.bold = True
        p.font.size = Pt(20)
        
        recomendaciones = [
            "• Monitorear localidades con alto índice de fuera de tiempo",
            "• Optimizar rutas en zonas con mayor volumen de pendientes",
            "• Coordinar con transportistas en áreas con bajo cumplimiento",
            "• Implementar alertas proactivas para pedidos próximos a vencer"
        ]
        
        for rec in recomendaciones:
            p = tf.add_paragraph()
            p.text = rec
            p.font.size = Pt(16)
        
        pptx_buffer = io.BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        return pptx_buffer
    
    with col_btn2:
        if st.button("📊 Generar y Descargar PowerPoint"):
            pptx_data = crear_pptx()
            st.download_button(
                label="⬇️ Descargar Presentación PPTX",
                data=pptx_data,
                file_name="Reporte_LeadTime_Presentacion.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    
    # --- VISTA PREVIA DE DATOS ---
    st.header("🔍 Vista Previa de Datos (primeras 10 filas)")

    columnas_mostrar = [
        'Cliente', 'Subcuenta', 'Agencia origen', 'Agencia destino', 'Condición de venta',
        'Fecha', 'Fecha último estado', 'Estado', 'ED', 'ZONA', 'Loc', 'Producto',
        'Lead Time', 'Días Prometidos', 'Día de Gracia Aplicado',
        'Cumplimiento', 'Días Restantes',
        'Alerta Devolución', 'Alerta Redespacho', 'Alerta Pendiente Fuera Tiempo', 'Alerta Pago Pendiente'
    ]

    df_vista_previa = df[columnas_mostrar].head(10)
    st.dataframe(df_vista_previa)

    # Botón para descargar vista previa completa en Excel
    excel_vista = generar_excel_desde_df(df[columnas_mostrar], "Vista Previa Completa")
    st.download_button(
        label="📥 Descargar Vista Previa Completa (Excel)",
        data=excel_vista,
        file_name="Vista_Previa_Datos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- PRUEBA RÁPIDA EN SIDEBAR ---
    st.sidebar.markdown("### 🧪 Prueba de Clasificación")
    prueba_localidad = st.sidebar.text_input("Ingresa localidad para probar:", "VICENTE LOPEZ, BUENOS AIRES")
    if prueba_localidad:
        zona = determinar_zona(prueba_localidad)
        st.sidebar.success(f"Clasificación: **{zona}**")

else:
    st.info("👆 Por favor, sube un archivo Excel para comenzar.")
    st.markdown("""
    **Instrucciones:**
    1. Haz clic en "Browse files".
    2. Selecciona tu archivo Excel.
    3. ¡Listo! La app calculará automáticamente y mostrará gráficos y botones de descarga.
    """)