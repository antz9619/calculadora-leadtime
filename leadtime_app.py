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
import pytz  # Para manejo de zonas horarias

# --- CONFIGURACIÓN DE ZONA HORARIA ---
# Definir la zona horaria de Argentina
ZONA_HORARIA_ARGENTINA = pytz.timezone('America/Argentina/Buenos_Aires')

def obtener_fecha_actual_argentina():
    """Obtiene la fecha actual en la zona horaria de Argentina"""
    return datetime.now(ZONA_HORARIA_ARGENTINA)

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
    
    # Convertir ambos a timezone-naive para comparar
    if hasattr(fecha_inicio, 'tz') and fecha_inicio.tz is not None:
        fecha_inicio = fecha_inicio.replace(tzinfo=None)
    if hasattr(fecha_fin, 'tz') and fecha_fin.tz is not None:
        fecha_fin = fecha_fin.replace(tzinfo=None)
    
    # Extraer solo la parte de fecha
    fecha_inicio = fecha_inicio.date() if hasattr(fecha_inicio, 'date') else fecha_inicio
    fecha_fin = fecha_fin.date() if hasattr(fecha_fin, 'date') else fecha_fin
    
    if fecha_inicio > fecha_fin:
        return 0
    
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

st.title("📊 Calculadora de Lead Time - Indicadores Mejorados")
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
    
    # --- CÁLCULO DE LEAD TIME CORREGIDO ---
    def calcular_lead_time(row):
        try:
            estado = str(row['Estado']).lower()
            ed = str(row.get('ED', '')).upper() if 'ED' in df.columns else 'SI'
            
            # Determinar si el pedido está entregado
            entregado = (
                (ed == "NO" and "esperando retiro" in estado) or 
                "entregada" in estado
            )
            
            # Usar la fecha actual de Argentina (sin timezone para compatibilidad)
            fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
            
            if entregado:
                # Para pedidos ENTREGADOS: calcular desde creación hasta último estado
                lead_time = calcular_dias_habiles(row['Fecha'], row['Fecha último estado'])
            else:
                # Para pedidos PENDIENTES: calcular desde creación hasta HOY
                lead_time = calcular_dias_habiles(row['Fecha'], fecha_actual_argentina)
            
            # Aplicar día de gracia para Delivery Hero Riders
            if row.get('Cliente', '') == "DELIVERY HERO E-COMMERCE S.A." and row.get('Subcuenta', '') == "RIDERS":
                if pd.notna(lead_time) and lead_time > 0:
                    lead_time = max(0, lead_time - 1)
            
            return lead_time
        
        except Exception as e:
            return None

    df['Lead Time'] = df.apply(calcular_lead_time, axis=1)
    
    # Columna para identificar si se aplicó día de gracia
    df['Día de Gracia Aplicado'] = df.apply(
        lambda row: "Sí" if row.get('Cliente', '') == "DELIVERY HERO E-COMMERCE S.A." and row.get('Subcuenta', '') == "RIDERS" else "No",
        axis=1
    )
    
    # --- CÁLCULO DE CUMPLIMIENTO MEJORADO (CON VISITAS Y ACCIONES) ---
    def determinar_cumplimiento_mejorado(row):
        estado = str(row['Estado']).lower()
        ed = str(row.get('ED', '')).upper() if 'ED' in df.columns else 'SI'
        condicion_venta = str(row.get('Condición de venta', '')).upper() if 'Condición de venta' in df.columns else ''
        visitas = row.get('Visitas', 0) if 'Visitas' in df.columns else 0
        
        # PRIMERO: Verificar si es una devolución (estado cerrado)
        if "devolución informada" in estado or "devolucion informada" in estado:
            return "Devuelto"
        
        # Si ED es "NO" y estado es "esperando retiro"
        if ed == "NO" and "esperando retiro" in estado:
            if pd.notna(row['Lead Time']) and row['Lead Time'] <= row['Días Prometidos']:
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
            # --- NUEVA LÓGICA: PEDIDOS PENDIENTES CON VISITAS ---
            # Calcular días desde la última visita hasta hoy
            fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
            dias_desde_ultima_visita = calcular_dias_habiles(row['Fecha último estado'], fecha_actual_argentina) if pd.notna(row['Fecha último estado']) else None
            
            # Verificar si tuvo al menos una visita dentro del tiempo prometido
            lead_time_hasta_visita = calcular_dias_habiles(row['Fecha'], row['Fecha último estado']) if pd.notna(row['Fecha último estado']) else None
            
            # Estados que indican una visita
            estados_visita = [
                "visita a domicilio", "reprogramada", "domicilio incompleto", 
                "domicilio incorrecto", "ausente", "rechazado"
            ]
            
            es_estado_visita = any(estado_visita in estado for estado_visita in estados_visita)
            
            if es_estado_visita and visitas > 0 and pd.notna(lead_time_hasta_visita):
                if lead_time_hasta_visita <= row['Días Prometidos']:
                    # Tuvo visita en tiempo, pero requiere acción según el motivo
                    if "domicilio incompleto" in estado:
                        return "Pendiente - Visita en Tiempo (Datos Incompletos)"
                    elif "domicilio incorrecto" in estado:
                        return "Pendiente - Visita en Tiempo (Domicilio Incorrecto)"
                    elif "ausente" in estado:
                        return "Pendiente - Visita en Tiempo (Cliente Ausente)"
                    elif "rechazado" in estado:
                        return "Pendiente - Visita en Tiempo (Cliente Rechazó)"
                    else:
                        return "Pendiente - Visita en Tiempo"
                else:
                    # Visita fuera de tiempo
                    return "Pendiente - Visita Fuera de Tiempo"
            
            # Para pendientes sin visita específica
            if pd.notna(row['Lead Time']):
                if row['Lead Time'] < row['Días Prometidos']:
                    return "Pendiente - En Tiempo"
                elif row['Lead Time'] == row['Días Prometidos']:
                    return "Pendiente - Último Día"
                else:
                    return "Pendiente - Fuera de Tiempo"
            else:
                return "Pendiente - Fuera de Tiempo"
    
    df['Cumplimiento'] = df.apply(determinar_cumplimiento_mejorado, axis=1)
    
    # Calcular días restantes para pendientes en tiempo
    def calcular_dias_restantes(row):
        cumplimiento = str(row['Cumplimiento'])
        
        if "Pendiente" in cumplimiento and "Fuera" not in cumplimiento and "Visita" not in cumplimiento:
            if pd.notna(row['Lead Time']):
                restantes = row['Días Prometidos'] - row['Lead Time']
                return f"{int(restantes)} días restantes" if restantes > 0 else "Vence hoy"
        return ""
    
    df['Días Restantes'] = df.apply(calcular_dias_restantes, axis=1)
    
    # --- NUEVAS ALERTAS MEJORADAS ---
    
    # Alerta para visitas en tiempo pero sin seguimiento
    def alerta_visita_sin_seguimiento(row):
        try:
            estado = str(row['Estado']).lower()
            cumplimiento = str(row['Cumplimiento'])
            fecha_ultimo_estado = row['Fecha último estado']
            
            # Verificar si es un pedido con visita en tiempo que requiere acción
            if "Visita en Tiempo" in cumplimiento and pd.notna(fecha_ultimo_estado):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_visita = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                
                if dias_desde_visita is not None and dias_desde_visita >= 3:
                    if "Datos Incompletos" in cumplimiento:
                        return "Solicitar datos completos"
                    elif "Domicilio Incorrecto" in cumplimiento:
                        return "Verificar domicilio"
                    elif "Cliente Ausente" in cumplimiento:
                        return "Coordinar nueva visita"
                    elif "Cliente Rechazó" in cumplimiento:
                        return "Sugerir devolución"
                    else:
                        return "Requiere seguimiento"
            return ""
        except Exception as e:
            return ""

    df['Alerta Seguimiento Visita'] = df.apply(alerta_visita_sin_seguimiento, axis=1)
    
    # Alerta para pedidos con múltiples visitas sin resultado
    def alerta_visitas_multiples(row):
        try:
            visitas = row.get('Visitas', 0) if 'Visitas' in df.columns else 0
            estado = str(row['Estado']).lower()
            
            if visitas >= 2 and ("ausente" in estado or "rechazado" in estado):
                return "Múltiples visitas sin éxito - Evaluar devolución"
            return ""
        except Exception as e:
            return ""

    df['Alerta Visitas Múltiples'] = df.apply(alerta_visitas_multiples, axis=1)
    
    # --- NUEVA ALERTA: UNA SOLA VISITA SIN SEGUIMIENTO EN 5 DÍAS ---
    def alerta_una_visita_sin_seguimiento(row):
        try:
            visitas = row.get('Visitas', 0) if 'Visitas' in df.columns else 0
            estado = str(row['Estado']).lower()
            fecha_ultimo_estado = row['Fecha último estado']
            cumplimiento = str(row['Cumplimiento'])
            
            # Estados que indican una visita realizada
            estados_visita = [
                "visita a domicilio", "reprogramada", "domicilio incompleto", 
                "domicilio incorrecto", "ausente", "rechazado"
            ]
            
            es_estado_visita = any(estado_visita in estado for estado_visita in estados_visita)
            
            # Verificar condiciones para la alerta
            if (visitas == 1 and 
                es_estado_visita and 
                pd.notna(fecha_ultimo_estado) and
                "Visita" in cumplimiento and  # Solo para pedidos con visita
                "Devuelto" not in cumplimiento and  # Excluir devueltos
                "Entregada" not in cumplimiento):  # Excluir entregados
                
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_visita = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                
                if dias_desde_visita is not None and dias_desde_visita >= 5:
                    return f"1 visita hace {dias_desde_visita} días hábiles - Sin seguimiento"
            
            return ""
        except Exception as e:
            return ""

    df['Alerta Una Visita Sin Seguimiento'] = df.apply(alerta_una_visita_sin_seguimiento, axis=1)
    
    # --- ALERTAS EXISTENTES (MANTENIDAS) ---
    
    def alerta_devolucion(row):
        try:
            estado = str(row['Estado']).lower()
            ed = str(row.get('ED', '')).upper() if 'ED' in df.columns else 'SI'
            fecha_ultimo_estado = row['Fecha último estado']
            
            if ed == "NO" and "esperando retiro" in estado and pd.notna(fecha_ultimo_estado):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 15:
                    return "Sugerir devolución"
            return ""
        except Exception as e:
            return ""

    df['Alerta Devolución'] = df.apply(alerta_devolucion, axis=1)
    
    def alerta_redespacho(row):
        try:
            estado = str(row['Estado']).lower()
            fecha_ultimo_estado = row['Fecha último estado']
            
            if "redespacho" in estado and pd.notna(fecha_ultimo_estado):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 2:
                    return "Redespacho demorado"
            return ""
        except Exception as e:
            return ""

    df['Alerta Redespacho'] = df.apply(alerta_redespacho, axis=1)
    
    def alerta_pendiente_fuera_tiempo(row):
        cumplimiento = str(row['Cumplimiento'])
        
        if cumplimiento == "Pendiente - Fuera de Tiempo":
            return "Fuera de tiempo crítico"
        return ""
    
    df['Alerta Pendiente Fuera Tiempo'] = df.apply(alerta_pendiente_fuera_tiempo, axis=1)
    
    def alerta_pago_pendiente(row):
        try:
            estado = str(row['Estado']).lower()
            condicion_venta = str(row.get('Condición de venta', '')).upper() if 'Condición de venta' in df.columns else ''
            fecha_ultimo_estado = row['Fecha último estado']
            
            if condicion_venta == "PD" and "esperando retiro" in estado and pd.notna(fecha_ultimo_estado):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 5:
                    return "Pago pendiente demorado"
            return ""
        except Exception as e:
            return ""

    df['Alerta Pago Pendiente'] = df.apply(alerta_pago_pendiente, axis=1)
    
    # --- FILTROS ---
    st.sidebar.header("🔍 Filtros")

    # Filtro por Cliente
    if 'Cliente' in df.columns:
        clientes = sorted(df['Cliente'].dropna().unique())
        cliente_seleccionado = st.sidebar.selectbox("Cliente", ["Todos"] + clientes)
    else:
        st.error("❌ La columna 'Cliente' no existe en el archivo.")
        st.stop()

    # Filtro por Subcuenta
    if 'Subcuenta' in df.columns:
        subcuentas = sorted(df['Subcuenta'].dropna().unique())
        subcuenta_seleccionada = st.sidebar.selectbox("Subcuenta", ["Todas"] + subcuentas)
    else:
        st.error("❌ La columna 'Subcuenta' no existe en el archivo.")
        st.stop()

    # Filtro por Agencia origen
    if 'Agencia origen' in df.columns:
        agencias_origen = sorted(df['Agencia origen'].dropna().unique())
        agencia_origen_seleccionada = st.sidebar.selectbox("Agencia origen", ["Todas"] + agencias_origen)
    else:
        st.warning("⚠️ La columna 'Agencia origen' no existe. Se omitirá este filtro.")
        agencia_origen_seleccionada = "Todas"

    # Filtro por Agencia destino
    if 'Agencia destino' in df.columns:
        agencias = sorted(df['Agencia destino'].dropna().unique())
        agencia_seleccionada = st.sidebar.selectbox("Agencia destino", ["Todas"] + agencias)
    else:
        st.error("❌ La columna 'Agencia destino' no existe en el archivo.")
        st.stop()

    # Filtro por ED
    if 'ED' in df.columns:
        ed_opciones = sorted(df['ED'].dropna().unique())
        ed_seleccionada = st.sidebar.selectbox("Entrega a Domicilio (ED)", ["Todas"] + ed_opciones)
    else:
        st.warning("⚠️ La columna 'ED' no existe. Se omitirá este filtro.")
        ed_seleccionada = "Todas"

    # Filtro por Condición de venta
    if 'Condición de venta' in df.columns:
        condiciones_venta = sorted(df['Condición de venta'].dropna().unique())
        condicion_venta_seleccionada = st.sidebar.selectbox("Condición de venta", ["Todas"] + condiciones_venta)
    else:
        st.warning("⚠️ La columna 'Condición de venta' no existe. Se omitirá este filtro.")
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

    if 'ED' in df.columns and ed_seleccionada != "Todas":
        df = df[df['ED'] == ed_seleccionada]

    if 'Condición de venta' in df.columns and condicion_venta_seleccionada != "Todas":
        df = df[df['Condición de venta'] == condicion_venta_seleccionada]
    
    # --- ESTADÍSTICAS MEJORADAS ---
    st.header("📊 Indicadores de Cumplimiento Mejorados")
    
    total_pedidos = df.shape[0]
    entregados = df[df['Cumplimiento'].str.startswith("Entregada")].shape[0]
    devueltos = df[df['Cumplimiento'] == "Devuelto"].shape[0]
    pendientes_reales = total_pedidos - entregados - devueltos
    
    # Nuevas categorías para visitas
    visita_en_tiempo = df[df['Cumplimiento'].str.contains("Visita en Tiempo", na=False)].shape[0]
    visita_fuera_tiempo = df[df['Cumplimiento'] == "Pendiente - Visita Fuera de Tiempo"].shape[0]
    
    # Clasificación detallada
    en_tiempo = df[df['Cumplimiento'] == "Entregada - En Tiempo"].shape[0]
    en_tiempo_pd = df[df['Cumplimiento'] == "Entregada - En Tiempo (PD: Pago Pendiente)"].shape[0]
    fuera_tiempo = df[df['Cumplimiento'] == "Entregada - Fuera de Tiempo"].shape[0]
    fuera_tiempo_pd = df[df['Cumplimiento'] == "Entregada - Fuera de Tiempo (PD: Pago Pendiente)"].shape[0]
    devuelto_count = devueltos
    pendiente_en_tiempo = df[df['Cumplimiento'] == "Pendiente - En Tiempo"].shape[0]
    pendiente_fuera_tiempo = df[df['Cumplimiento'] == "Pendiente - Fuera de Tiempo"].shape[0]
    pendiente_ultimo_dia = df[df['Cumplimiento'] == "Pendiente - Último Día"].shape[0]
    
    # --- NUEVOS INDICADORES MEJORADOS ---
    
    # 1. Cumplimiento tradicional (solo entregados)
    cumplimiento_tradicional = ((en_tiempo + en_tiempo_pd) / entregados * 100) if entregados > 0 else 0
    
    # 2. Cumplimiento de gestión (entregados + visitas en tiempo)
    pedidos_gestionados = entregados + visita_en_tiempo
    cumplimiento_gestion = (pedidos_gestionados / total_pedidos * 100) if total_pedidos > 0 else 0
    
    # 3. Efectividad de visitas
    total_visitas = visita_en_tiempo + visita_fuera_tiempo
    efectividad_visitas = (visita_en_tiempo / total_visitas * 100) if total_visitas > 0 else 0
    
    # 4. Tasa de resolución (entregados / total gestionado)
    tasa_resolucion = (entregados / pedidos_gestionados * 100) if pedidos_gestionados > 0 else 0
    
    # Métricas principales en una sola línea
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Total Pedidos", total_pedidos)
    with col2:
        st.metric("✅ Entregados", entregados, f"{(entregados/total_pedidos*100):.1f}%")
    with col3:
        st.metric("🔄 Devueltos", devueltos, f"{(devueltos/total_pedidos*100):.1f}%")
    with col4:
        st.metric("⏳ Pendientes", pendientes_reales, f"{(pendientes_reales/total_pedidos*100):.1f}%")
    
    # Segunda línea de métricas
    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.metric("🎯 Cumplimiento Tradicional", f"{cumplimiento_tradicional:.1f}%")
    with col6:
        st.metric("🚀 Cumplimiento Gestión", f"{cumplimiento_gestion:.1f}%")
    with col7:
        st.metric("📋 Visitas en Tiempo", visita_en_tiempo, f"{(visita_en_tiempo/total_pedidos*100):.1f}%")
    with col8:
        st.metric("📊 Efectividad Visitas", f"{efectividad_visitas:.1f}%")
    
    # --- TABLA DE RESUMEN MEJORADA ---
    st.header("📈 Detalle de Estados")
    
    resumen_data = {
        "Categoría": [
            "TOTAL PEDIDOS",
            "ENTREGADOS",
            " - En Tiempo",
            " - En Tiempo (PD)",
            " - Fuera de Tiempo", 
            " - Fuera de Tiempo (PD)",
            "DEVUELTOS",
            "PENDIENTES CON VISITA",
            " - Visita en Tiempo",
            " - Visita Fuera de Tiempo",
            "PENDIENTES SIN VISITA",
            " - En Tiempo",
            " - Último Día",
            " - Fuera de Tiempo"
        ],
        "Cantidad": [
            total_pedidos,
            entregados,
            en_tiempo,
            en_tiempo_pd,
            fuera_tiempo,
            fuera_tiempo_pd,
            devuelto_count,
            visita_en_tiempo + visita_fuera_tiempo,
            visita_en_tiempo,
            visita_fuera_tiempo,
            pendiente_en_tiempo + pendiente_ultimo_dia + pendiente_fuera_tiempo,
            pendiente_en_tiempo,
            pendiente_ultimo_dia,
            pendiente_fuera_tiempo
        ],
        "Porcentaje": [
            "100%",
            f"{(entregados/total_pedidos*100):.1f}%",
            f"{(en_tiempo/total_pedidos*100):.1f}%",
            f"{(en_tiempo_pd/total_pedidos*100):.1f}%",
            f"{(fuera_tiempo/total_pedidos*100):.1f}%",
            f"{(fuera_tiempo_pd/total_pedidos*100):.1f}%",
            f"{(devuelto_count/total_pedidos*100):.1f}%",
            f"{((visita_en_tiempo + visita_fuera_tiempo)/total_pedidos*100):.1f}%",
            f"{(visita_en_tiempo/total_pedidos*100):.1f}%",
            f"{(visita_fuera_tiempo/total_pedidos*100):.1f}%",
            f"{((pendiente_en_tiempo + pendiente_ultimo_dia + pendiente_fuera_tiempo)/total_pedidos*100):.1f}%",
            f"{(pendiente_en_tiempo/total_pedidos*100):.1f}%",
            f"{(pendiente_ultimo_dia/total_pedidos*100):.1f}%",
            f"{(pendiente_fuera_tiempo/total_pedidos*100):.1f}%"
        ]
    }
    
    resumen_df = pd.DataFrame(resumen_data)
    st.dataframe(resumen_df, use_container_width=True)
    
    # --- GRÁFICO DE CUMPLIMIENTO MEJORADO ---
    categorias_mejoradas = [
        "Entregada - En Tiempo", 
        "Entregada - En Tiempo (PD)",
        "Entregada - Fuera de Tiempo", 
        "Entregada - Fuera de Tiempo (PD)",
        "Devuelto",
        "Pendiente - Visita en Tiempo",
        "Pendiente - Visita Fuera de Tiempo",
        "Pendiente - En Tiempo", 
        "Pendiente - Último Día",
        "Pendiente - Fuera de Tiempo"
    ]
    
    valores_mejorados = [
        en_tiempo,
        en_tiempo_pd,
        fuera_tiempo,
        fuera_tiempo_pd,
        devuelto_count,
        visita_en_tiempo,
        visita_fuera_tiempo,
        pendiente_en_tiempo, 
        pendiente_ultimo_dia,
        pendiente_fuera_tiempo
    ]
    
    fig1 = px.pie(
        names=categorias_mejoradas,
        values=valores_mejorados,
        title="Distribución de Cumplimiento Mejorado (Incluyendo Visitas)",
        color=categorias_mejoradas,
        color_discrete_map={
            "Entregada - En Tiempo": "#28a745",
            "Entregada - En Tiempo (PD)": "#2ecc71",
            "Entregada - Fuera de Tiempo": "#dc3545",
            "Entregada - Fuera de Tiempo (PD)": "#e74c3c",
            "Devuelto": "#9b59b6",
            "Pendiente - Visita en Tiempo": "#3498db",
            "Pendiente - Visita Fuera de Tiempo": "#e67e22",
            "Pendiente - En Tiempo": "#ffc107",
            "Pendiente - Último Día": "#fd7e14",
            "Pendiente - Fuera de Tiempo": "#6c757d"
        },
        hole=0.4
    )
    fig1.update_traces(textinfo='percent+value', textposition='inside')
    st.plotly_chart(fig1, use_container_width=True)
    
    # --- NUEVAS ALERTAS MEJORADAS EN LA INTERFAZ ---
    
    # Alertas de seguimiento de visitas
    alertas_seguimiento = df[df['Alerta Seguimiento Visita'] != ""]
    if not alertas_seguimiento.empty:
        st.header("🔔 Alertas de Seguimiento de Visitas")
        st.write("Pedidos con visita en tiempo que requieren acción:")
        
        columnas_alerta = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Estado', 'Visitas', 
                          'Fecha último estado', 'Cumplimiento', 'Alerta Seguimiento Visita']
        df_alerta = alertas_seguimiento[columnas_alerta]
        
        st.dataframe(df_alerta)
        
        # Función auxiliar para generar Excel
        def generar_excel_desde_df(df, nombre_hoja="Datos"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            output.seek(0)
            return output
        
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Seguimiento")
        st.download_button(
            label="📥 Descargar Alertas de Seguimiento (Excel)",
            data=excel_data,
            file_name="Alertas_Seguimiento_Visitas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Alertas de visitas múltiples
    alertas_multiples = df[df['Alerta Visitas Múltiples'] != ""]
    if not alertas_multiples.empty:
        st.header("🔄 Alertas de Visitas Múltiples")
        st.write("Pedidos con múltiples visitas sin éxito:")
        
        columnas_alerta = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Estado', 'Visitas', 
                          'Fecha último estado', 'Alerta Visitas Múltiples']
        df_alerta = alertas_multiples[columnas_alerta]
        
        st.dataframe(df_alerta)
        
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Visitas Múltiples")
        st.download_button(
            label="📥 Descargar Alertas Visitas Múltiples (Excel)",
            data=excel_data,
            file_name="Alertas_Visitas_Multiples.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # --- NUEVA ALERTA: UNA VISITA SIN SEGUIMIENTO ---
    alertas_una_visita = df[df['Alerta Una Visita Sin Seguimiento'] != ""]
    if not alertas_una_visita.empty:
        st.header("⏰ Alertas de Una Visita Sin Seguimiento")
        st.write("Pedidos con solo una visita que no han tenido seguimiento en 5+ días hábiles:")
        
        columnas_alerta = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Estado', 'Visitas', 
                          'Fecha último estado', 'Cumplimiento', 'Alerta Una Visita Sin Seguimiento']
        df_alerta = alertas_una_visita[columnas_alerta]
        
        st.dataframe(df_alerta)
        
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Una Visita Sin Seguimiento")
        st.download_button(
            label="📥 Descargar Alertas Una Visita Sin Seguimiento (Excel)",
            data=excel_data,
            file_name="Alertas_Una_Visita_Sin_Seguimiento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Alertas de devolución existentes
    alertas_devolucion = df[df['Alerta Devolución'] == "Sugerir devolución"]
    if not alertas_devolucion.empty:
        st.header("🚨 Alertas de Devolución")
        st.write("Los siguientes pedidos están en estado 'Esperando retiro' por más de 15 días hábiles:")
        
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
    
    # Alertas de redespacho
    alertas_redespacho = df[df['Alerta Redespacho'] == "Redespacho demorado"]
    if not alertas_redespacho.empty:
        st.header("🚨 Alertas de Redespacho Demorado")
        st.write("Los siguientes pedidos están en estado 'Redespacho' por más de 48 horas hábiles:")
        
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
    
    # Alertas de pendiente fuera de tiempo
    alertas_pendiente_fuera_tiempo = df[df['Alerta Pendiente Fuera Tiempo'] == "Fuera de tiempo crítico"]
    if not alertas_pendiente_fuera_tiempo.empty:
        st.header("🚨 Alertas de Pendiente Fuera de Tiempo")
        st.write("Los siguientes pedidos están pendientes y fuera del tiempo de entrega prometido:")
        
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

    # Alertas de pago pendiente
    alertas_pago_pendiente = df[df['Alerta Pago Pendiente'] == "Pago pendiente demorado"]
    if not alertas_pago_pendiente.empty:
        st.header("🚨 Alertas de Pago Pendiente Demorado")
        st.write("Los siguientes pedidos están en estado 'Esperando retiro' con condición de venta PD por más de 5 días hábiles:")
        
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
        (df['Alerta Seguimiento Visita'] != "") |
        (df['Alerta Visitas Múltiples'] != "") |
        (df['Alerta Una Visita Sin Seguimiento'] != "") |
        (df['Alerta Devolución'] == "Sugerir devolución") |
        (df['Alerta Redespacho'] == "Redespacho demorado") |
        (df['Alerta Pendiente Fuera Tiempo'] == "Fuera de tiempo crítico") |
        (df['Alerta Pago Pendiente'] == "Pago pendiente demorado")
    ]

    if not todas_alertas.empty:
        columnas_todas = ['Guia', 'Cliente', 'Destinatario', 'Loc', 'Estado', 'Visitas', 'Fecha último estado', 
                          'Cumplimiento', 'Alerta Seguimiento Visita', 'Alerta Visitas Múltiples', 
                          'Alerta Una Visita Sin Seguimiento', 'Alerta Devolución', 'Alerta Redespacho', 
                          'Alerta Pendiente Fuera Tiempo', 'Alerta Pago Pendiente']
        
        # Filtrar columnas que existen en el DataFrame
        columnas_existentes = [col for col in columnas_todas if col in df.columns]
        df_todas = todas_alertas[columnas_existentes]
        
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

    # --- DESCARGAS GENERALES ---
    st.header("📥 Descargas Generales")
    
    # Preparar Excel con gráficos
    output_excel = io.BytesIO()

    # Crear datos para el gráfico de estadísticas
    stats_data = {
        "Métrica": [
            "Total Pedidos", "Entregados", "Devueltos", "Pendientes Reales",
            "Entregada - En Tiempo", "Entregada - En Tiempo (PD)",
            "Entregada - Fuera de Tiempo", "Entregada - Fuera de Tiempo (PD)",
            "Devuelto",
            "Pendiente - Visita en Tiempo", "Pendiente - Visita Fuera de Tiempo",
            "Pendiente - En Tiempo", "Pendiente - Último Día",
            "Pendiente - Fuera de Tiempo",
            "% Cumplimiento Tradicional", "% Cumplimiento Gestión"
        ],
        "Valor": [
            total_pedidos, entregados, devueltos, pendientes_reales,
            en_tiempo, en_tiempo_pd,
            fuera_tiempo, fuera_tiempo_pd,
            devuelto_count,
            visita_en_tiempo, visita_fuera_tiempo,
            pendiente_en_tiempo, pendiente_ultimo_dia,
            pendiente_fuera_tiempo,
            f"{cumplimiento_tradicional:.2f}%", f"{cumplimiento_gestion:.2f}%"
        ]
    }
    
    if len(stats_data["Métrica"]) == len(stats_data["Valor"]):
        stats_df = pd.DataFrame(stats_data)
    else:
        st.error("❌ Error: Las listas de estadísticas tienen longitudes diferentes")
        stats_df = pd.DataFrame({"Métrica": ["Error en estadísticas"], "Valor": ["Contactar al administrador"]})
    
    # Guardar en Excel
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Base", index=False)
        stats_df.to_excel(writer, sheet_name="Estadísticas", index=False)

    output_excel.seek(0)
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        st.download_button(
            label="📥 Descargar Excel Actualizado (Completo)",
            data=output_excel,
            file_name="Reporte_LeadTime_Actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # --- GENERAR POWERPOINT (SIMPLIFICADO) ---
    def crear_pptx():
        prs = Presentation()
        
        # Slide 1: Título
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = "Reporte de Cumplimiento de Entregas"
        subtitle.text = "Lead Time - Indicadores Mejorados\nGenerado automáticamente"
        
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
            f"• Devueltos: {devueltos} ({(devueltos/total_pedidos*100):.1f}%)",
            f"• Cumplimiento Tradicional: {cumplimiento_tradicional:.1f}%",
            f"• Cumplimiento Gestión: {cumplimiento_gestion:.1f}%",
            f"• Visitas en Tiempo: {visita_en_tiempo}",
            f"• Alertas Activas: {len(todas_alertas) if 'todas_alertas' in locals() else 0}"
        ]
        
        for metric in metrics:
            p = tf.add_paragraph()
            p.text = metric
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
        'Fecha', 'Fecha último estado', 'Estado', 'Visitas', 'ED', 'ZONA', 'Loc', 'Producto',
        'Lead Time', 'Días Prometidos', 'Día de Gracia Aplicado',
        'Cumplimiento', 'Días Restantes',
        'Alerta Seguimiento Visita', 'Alerta Visitas Múltiples', 'Alerta Una Visita Sin Seguimiento',
        'Alerta Devolución', 'Alerta Redespacho', 'Alerta Pendiente Fuera Tiempo', 'Alerta Pago Pendiente'
    ]

    # Mostrar solo las columnas que existen en el DataFrame
    columnas_existentes = [col for col in columnas_mostrar if col in df.columns]
    df_vista_previa = df[columnas_existentes].head(10)
    st.dataframe(df_vista_previa)

    # Botón para descargar vista previa completa en Excel
    excel_vista = generar_excel_desde_df(df[columnas_existentes], "Vista Previa Completa")
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