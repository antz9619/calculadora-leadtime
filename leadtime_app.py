import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
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
import re  # Para expresiones regulares en detección robusta

# --- CONFIGURACIÓN DE ZONA HORARIA ---
# Definir la zona horaria de Argentina
ZONA_HORARIA_ARGENTINA = pytz.timezone('America/Argentina/Buenos_Aires')

def obtener_fecha_actual_argentina():
    """Obtiene la fecha actual en la zona horaria de Argentina"""
    return datetime.now(ZONA_HORARIA_ARGENTINA)

# --- FUNCIÓN PARA GENERAR EXCEL (MOVIDA AL INICIO) ---
def generar_excel_desde_df(df, nombre_hoja="Datos"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)
    output.seek(0)
    return output

# --- SISTEMA MEJORADO DE FERIADOS Y PUENTES ---

def es_dia_festivo(fecha=None):
    """Verifica si la fecha es un día festivo configurado"""
    if fecha is None:
        fecha = date.today()
    
    # Convertir a date si es datetime
    if isinstance(fecha, datetime):
        fecha = fecha.date()
    
    festivos = [
        (1, 1),   # Año Nuevo
        (3, 24),  # Día Nacional de la Memoria
        (4, 2),   # Día del Veterano
        (5, 1),   # Día del Trabajo
        (5, 25),  # Día de la Revolución de Mayo
        (6, 17),  # Paso a la Inmortalidad del Gral. Martín Güemes
        (6, 20),  # Día de la Bandera
        (7, 9),   # Día de la Independencia
        (10, 12), # Día de la Raza
        (11, 20), # Día de la Soberanía Nacional
        (12, 8),  # Inmaculada Concepción
        (12, 25), # Navidad
        # Agregar más según necesidad
    ]
    return (fecha.month, fecha.day) in festivos

def es_feriado_puente(fecha):
    """
    Detecta feriados puente con mensaje descriptivo.
    Devuelve (bool, str) donde str es el motivo detallado
    """
    # Convertir a date si es datetime
    if isinstance(fecha, datetime):
        fecha = fecha.date()
    
    # Viernes antes de fin de semana festivo
    if fecha.weekday() == 4:  # Viernes
        sabado = fecha + timedelta(days=1)
        domingo = fecha + timedelta(days=2)
        
        if es_dia_festivo(sabado) and es_dia_festivo(domingo):
            return True, f"Viernes puente (festivo {sabado.strftime('%d/%m')} y {domingo.strftime('%d/%m')})"
        elif es_dia_festivo(sabado):
            return True, f"Viernes puente (festivo {sabado.strftime('%d/%m')})"
        elif es_dia_festivo(domingo):
            return True, f"Viernes puente (festivo {domingo.strftime('%d/%m')})"
    
    return False, ""

def es_dia_laborable(fecha):
    """Determina si una fecha es laborable (no fin de semana, no feriado, no puente)"""
    # Convertir a date si es datetime
    if isinstance(fecha, datetime):
        fecha = fecha.date()
    
    # Verificar fin de semana
    if fecha.weekday() >= 5:  # Sábado (5) o Domingo (6)
        return False
    
    # Verificar feriados
    if es_dia_festivo(fecha):
        return False
    
    # Verificar puentes
    if es_feriado_puente(fecha)[0]:  # Usamos el booleano de retorno
        return False
        
    return True

# REEMPLAZAR las funciones existentes
def es_dia_habil(fecha):
    """Determina si un día es hábil (versión mejorada con puentes)"""
    return es_dia_laborable(fecha)

def es_feriado(fecha):
    """Determina si un día es feriado (versión mejorada)"""
    return es_dia_festivo(fecha) or es_feriado_puente(fecha)[0]

# --- DICCIONARIO DE SEMANAS REALES (CALENDARIO) ---
def obtener_semana_calendario(fecha):
    """
    Calcula la semana del año según calendario (lunes a domingo)
    usando el estándar ISO 8601
    """
    if pd.isna(fecha):
        return None
    try:
        # Asegurarse de que es datetime
        if isinstance(fecha, str):
            fecha = pd.to_datetime(fecha, errors='coerce')
        # Calcular semana ISO (lunes como primer día de la semana)
        semana = fecha.isocalendar()[1]
        return semana
    except:
        return None

# --- FUNCIÓN CALCULAR DÍAS HÁBILES ACTUALIZADA (CON PUENTES) ---
def calcular_dias_habiles(fecha_inicio, fecha_fin):
    """
    Calcula días hábiles entre dos fechas CONSIDERANDO FERIADOS Y PUENTES
    """
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
    current = fecha_inicio + timedelta(days=1)  # Empezar desde el día siguiente
    
    while current <= fecha_fin:
        if es_dia_habil(current):  # ← ¡AHORA SÍ usa la lógica completa con puentes!
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

    # --- MEJORADO: Rellenar clientes vacíos con "EVENTUAL" de forma más robusta ---
    if 'Cliente' in df.columns:
        # Rellenar NaN y strings vacíos
        df['Cliente'] = df['Cliente'].fillna("EVENTUAL")
        df['Cliente'] = df['Cliente'].apply(lambda x: "EVENTUAL" if str(x).strip() == "" else x)
        df['Cliente'] = df['Cliente'].astype(str)
    else:
        st.error("❌ La columna 'Cliente' no existe en el archivo.")
        st.stop()

    # Renombrar columnas si es necesario
    if 'Localidad destino' in df.columns:
        df['Loc'] = df['Localidad destino']

    # Convertir columnas de fecha
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df['Fecha último estado'] = pd.to_datetime(df['Fecha último estado'], errors='coerce')

    # Aplicar la función para crear la columna de semana calendario
    df['Semana Calendario'] = df['Fecha'].apply(obtener_semana_calendario)

    # --- NUEVO: EXCLUIR LA AGENCIA DE DESTINO DE PAQUETERÍA INTERNA ---
    # Filtrar para excluir la agencia (6100) Administracion IPE
    if 'Agencia destino' in df.columns:
        df_original_count = df.shape[0]
        df = df[df['Agencia destino'] != "(6100) Administracion IPE"]
        excluded_count = df_original_count - df.shape[0]
        if excluded_count > 0:
            st.info(f"ℹ️ Se excluyeron {excluded_count} guías con destino a '(6100) Administracion IPE' (paquetería interna).")

    # Determinar ZONA (AMBA o INTERIOR)
    df['ZONA'] = df['Loc'].apply(determinar_zona)

    # Determinar días prometidos según ZONA, pero con excepción para Delivery Hero Riders
    def determinar_dias_prometidos_robusta(row):
        """
        Versión más robusta para determinar días prometidos
        """
        try:
            # Normalizar y limpiar los valores
            cliente = str(row.get('Cliente', '')).strip().upper()
            subcuenta = str(row.get('Subcuenta', '')).strip().upper()
            zona = str(row.get('ZONA', '')).strip()
            
            # Caso RIDERS (más flexible en la comparación)
            if ("DELIVERY HERO" in cliente and "RIDERS" in subcuenta):
                return 3
            
            # Caso normal
            if zona == "AMBA":
                return 3
            else:
                return 5
                
        except Exception as e:
            # Si hay error, retornar valor por defecto
            return 5

    # Prueba temporal con la versión robusta
    df['Días Prometidos'] = df.apply(determinar_dias_prometidos_robusta, axis=1)

    # --- CÁLCULO DE LEAD TIME CORREGIDO (CON PUENTES) ---
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
            
            return lead_time
        except Exception as e:
            return None

    df['Lead Time'] = df.apply(calcular_lead_time, axis=1)


    # --- CÁLCULO DE CUMPLIMIENTO MEJORADO (CON VISITAS Y ACCIONES) ---
    def determinar_cumplimiento_mejorado(row):
        estado = str(row['Estado']).lower()
        ed = str(row.get('ED', '')).upper() if 'ED' in df.columns else 'SI'
        condicion_venta = str(row.get('Condición de venta', '')).upper() if 'Condición de venta' in df.columns else ''
        visitas = row.get('Visitas', 0) if 'Visitas' in df.columns else 0

        # --- NUEVO: Verificar si es CANCELADA ---
        if "cancelada" in estado:
            return "Cancelada"
        
        # --- MEJORADO: Identificar devoluciones de EVENTUAL por el nombre del destinatario ---
        if row.get('Cliente', '') == "EVENTUAL":
            destinatario = ""
            if 'Destinatario' in row.index:
                destinatario_value = row['Destinatario']
                if pd.notna(destinatario_value) and str(destinatario_value).strip() != "":
                    destinatario = str(destinatario_value).lower().strip()
            
            palabras_devolucion = ["devolucion", "devolucion md", "devolucion p-ya", "dev. pedidos ya/", 
                                "devoluciones", "devo", "devol", "devolución", "devoluciónes", 
                                "devol pedido ya", "dev a origen"]
            
            if destinatario and any(palabra in destinatario for palabra in palabras_devolucion):
                return "Devuelto"

        # --- ACTUALIZADO: Verificar si es una devolución (estado cerrado) ---
        if "devolución informada" in estado or "devolucion informada" in estado or "devuelta" in estado:
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
            # --- CORRECCIÓN CLAVE: USAR SIEMPRE EL LEAD TIME CONTRA HOY PARA EVALUAR VENCIMIENTO ---
            if pd.notna(row['Lead Time']):
                if row['Lead Time'] < row['Días Prometidos']:
                    base_estado = "Pendiente - En Tiempo"
                elif row['Lead Time'] == row['Días Prometidos']:
                    base_estado = "Pendiente - Último Día"
                else:
                    base_estado = "Pendiente - Fuera de Tiempo"

                # --- AÑADIR DETALLE DE VISITA (solo si aplica) ---
                estados_visita = [
                    "visita a domicilio", "reprogramada", "domicilio incompleto", 
                    "domicilio incorrecto", "ausente", "rechazado"
                ]
                es_estado_visita = any(e in estado for e in estados_visita)

                if es_estado_visita and visitas > 0:
                    if "domicilio incompleto" in estado:
                        return base_estado + " (Datos Incompletos)"
                    elif "domicilio incorrecto" in estado:
                        return base_estado + " (Domicilio Incorrecto)"
                    elif "ausente" in estado:
                        return base_estado + " (Cliente Ausente)"
                    elif "rechazado" in estado:
                        return base_estado + " (Cliente Rechazó)"
                    else:
                        return base_estado + " (Visita Realizada)"
                else:
                    return base_estado
            else:
                return "Pendiente - Fuera de Tiempo"

    df['Cumplimiento'] = df.apply(determinar_cumplimiento_mejorado, axis=1)

    # Calcular días restantes para pendientes en tiempo
    def calcular_dias_restantes(row):
        cumplimiento = str(row['Cumplimiento'])
        # Incluir todos los pendientes, incluso con visita, EXCEPTO:
        # - Entregados, Devueltos, Cancelados
        # - Pendientes ya FUERA de tiempo (para esos mostramos "Vencido")
        if ("Pendiente" in cumplimiento and 
            "Entregada" not in cumplimiento and 
            "Devuelto" not in cumplimiento and 
            "Cancelada" not in cumplimiento):
            
            if pd.notna(row['Lead Time']):
                restantes = row['Días Prometidos'] - row['Lead Time']
                if restantes > 1:
                    return f"{int(restantes)} días restantes"
                elif restantes == 1:
                    return "Vence mañana"
                elif restantes == 0:
                    return "Vence hoy"
                else:
                    return "Vencido"
        return ""
    df['Días Restantes'] = df.apply(calcular_dias_restantes, axis=1)            

    # --- ALERTA DE EN TRÁNSITO DEMORADO (ACTUALIZADA CON PUENTES) ---
    def alerta_en_transito_demorado(row):
        try:
            estado = str(row['Estado']).lower()
            fecha_ultimo_estado = row['Fecha último estado']
            if "en tránsito" in estado and pd.notna(fecha_ultimo_estado):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 2:
                    return "En tránsito demorado (>48hs)"
            return ""
        except Exception as e:
            return ""

    df['Alerta En Tránsito Demorado'] = df.apply(alerta_en_transito_demorado, axis=1)

    # --- NUEVA ALERTA: ESTADO "CREADA" DEMORADO MÁS DE 24 HORAS (ACTUALIZADA CON PUENTES) ---
    def alerta_creada_demorada(row):
        try:
            estado = str(row['Estado']).lower()
            fecha_ultimo_estado = row['Fecha último estado']
            fecha_creacion = row['Fecha']
            
            # Verificar si el estado es "Creada" y tenemos fechas válidas
            if "creada" in estado and pd.notna(fecha_ultimo_estado) and pd.notna(fecha_creacion):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                
                # Calcular diferencia en horas (no días hábiles, sino horas reales)
                diferencia_horas = (fecha_actual_argentina - fecha_ultimo_estado).total_seconds() / 3600
                
                # Si lleva más de 24 horas en estado "Creada"
                if diferencia_horas >= 24:
                    # Calcular también días hábiles para información adicional
                    dias_habiles_creada = calcular_dias_habiles(fecha_creacion, fecha_actual_argentina)
                    return f"Creada demorada ({diferencia_horas:.1f} horas, {dias_habiles_creada} días hábiles)"
                
                # Opcional: Alerta preventiva entre 12-24 horas
                elif diferencia_horas >= 12:
                    dias_habiles_creada = calcular_dias_habiles(fecha_creacion, fecha_actual_argentina)
                    return f"Creada próxima a vencer ({diferencia_horas:.1f} horas)"
                    
            return ""
        except Exception as e:
            return ""

    df['Alerta Creada Demorada'] = df.apply(alerta_creada_demorada, axis=1)

    # --- NUEVAS ALERTAS MEJORADAS (ACTUALIZADAS CON PUENTES) ---
    # Alerta para pedidos con MÚLTIPLES visitas sin seguimiento (2+ visitas, 3+ días)
    def alerta_seguimiento_visitas(row):
        try:
            estado = str(row['Estado']).lower()
            cumplimiento = str(row['Cumplimiento'])
            fecha_ultimo_estado = row['Fecha último estado']
            visitas = row.get('Visitas', 0) if 'Visitas' in df.columns else 0

            # Solo para pedidos con 2 o más visitas
            if visitas >= 2 and "Visita" in cumplimiento and pd.notna(fecha_ultimo_estado):
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

    df['Alerta Seguimiento Visitas'] = df.apply(alerta_seguimiento_visitas, axis=1)

    # --- NUEVA ALERTA: UNA SOLA VISITA SIN SEGUIMIENTO EN 5 DÍAS (ACTUALIZADA CON PUENTES) ---
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

            # Verificar condiciones para la alerta (exactamente 1 visita)
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

    # --- ALERTAS EXISTENTES (MANTENIDAS Y ACTUALIZADAS CON PUENTES) ---
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
            if "redespachada" in estado and pd.notna(fecha_ultimo_estado):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 2:
                    return "Redespacho demorado"
            return ""
        except Exception as e:
            return ""

    df['Alerta Redespacho'] = df.apply(alerta_redespacho, axis=1)

    def alerta_en_transito_demorado(row):
        try:
            estado = str(row['Estado']).lower()
            fecha_ultimo_estado = row['Fecha último estado']
            if "en tránsito" in estado and pd.notna(fecha_ultimo_estado):
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                dias_desde_ultimo_estado = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                if dias_desde_ultimo_estado is not None and dias_desde_ultimo_estado >= 5:
                    return "En tránsito demorado"
            return ""
        except Exception as e:
            return ""
        
    df['Alerta En Tránsito Demorado'] = df.apply(alerta_en_transito_demorado, axis=1)

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

    # --- NUEVA ALERTA: REPROGRAMADA CON 0 VISITAS (ACTUALIZADA CON PUENTES) ---
    def alerta_reprogramada_sin_visitas(row):
        try:
            estado = str(row['Estado']).lower()
            visitas = row.get('Visitas', 0) if 'Visitas' in df.columns else 0
            # Verificar si el estado contiene "reprogramada" y tiene 0 visitas
            if "reprogramada" in estado and visitas == 0:
                fecha_actual_argentina = obtener_fecha_actual_argentina().replace(tzinfo=None)
                fecha_ultimo_estado = row['Fecha último estado']

                if pd.notna(fecha_ultimo_estado):
                    dias_desde_reprogramacion = calcular_dias_habiles(fecha_ultimo_estado, fecha_actual_argentina)
                    if dias_desde_reprogramacion is not None and dias_desde_reprogramacion >= 2:
                        return f"Reprogramada sin visita ({dias_desde_reprogramacion} días hábiles)"
                    else:
                        return "Reprogramada sin visita"
                else:
                    return "Reprogramada sin visita"
            return ""
        except Exception as e:
            return ""

    df['Alerta Reprogramada Sin Visitas'] = df.apply(alerta_reprogramada_sin_visitas, axis=1)

    # --- NUEVA ALERTA GENERAL: VENCIMIENTO MAÑANA PARA TODOS LOS CLIENTES (CORREGIDA CON PUENTES) ---
    def alerta_vencimiento_mañana(row):
        """
        Genera una alerta si:
        - El pedido está pendiente
        - Y vence mañana (Lead Time == Días Prometidos - 1) -> "Vence mañana"
        - O ya está vencido (Lead Time >= Días Prometidos) -> "Ya vencido"
        - Considera la excepción de RIDERS (siempre 3 días sin importar zona)
        - CORRECCIÓN: Verifica si "mañana" es día hábil
        """
        try:
            estado = str(row.get('Estado', '')).lower()
            cumplimiento = str(row.get('Cumplimiento', ''))
            lead_time = row.get('Lead Time')
            cliente = str(row.get('Cliente', '')).strip()
            subcuenta = str(row.get('Subcuenta', '')).strip()
            zona = row.get('ZONA', '')
            fecha_creacion = row.get('Fecha')
            
            # Solo aplicar a pedidos pendientes (no entregados, no cancelados, no devueltos)
            if ("entregada" in estado or 
                "cancelada" in estado or 
                "devuelto" in cumplimiento.lower()):
                return ""

            # Calcular días prometidos CORRECTAMENTE (igual que en determinar_dias_prometidos)
            if cliente == "DELIVERY HERO E-COMMERCE S.A." and subcuenta == "RIDERS":
                dias_prometidos_correcto = 3  # Excepción RIDERS: siempre 3 días
            else:
                dias_prometidos_correcto = 3 if zona == "AMBA" else 5  # Comportamiento normal

            # Verificar condiciones de vencimiento
            if (pd.notna(lead_time) and isinstance(lead_time, (int, float)) and
                pd.notna(fecha_creacion)):
                
                # Ya vencido
                if lead_time >= dias_prometidos_correcto:
                    return "Ya vencido"
                
                # Vence mañana - PERO VERIFICAR SI MAÑANA ES DÍA HÁBIL
                elif lead_time == dias_prometidos_correcto - 1:
                    fecha_actual = obtener_fecha_actual_argentina().date()
                    fecha_manana = fecha_actual + timedelta(days=1)
                    
                    # Verificar si "mañana" es día hábil
                    if es_dia_habil(fecha_manana):
                        return "Vence mañana"
                    else:
                        # Si mañana NO es hábil, buscar el próximo día hábil
                        proximo_dia_habil = fecha_manana
                        while not es_dia_habil(proximo_dia_habil):
                            proximo_dia_habil += timedelta(days=1)
                        
                        # Calcular días hábiles hasta el próximo día hábil
                        dias_hasta_proximo_habil = calcular_dias_habiles(fecha_actual, proximo_dia_habil)
                        
                        if dias_hasta_proximo_habil == 1:
                            return f"Vence {proximo_dia_habil.strftime('%d/%m')}"
                        else:
                            return f"Vence en {dias_hasta_proximo_habil} días"
                    
            return ""
        except Exception as e:
            return ""

    df['Alerta Vencimiento Mañana'] = df.apply(alerta_vencimiento_mañana, axis=1)

    # --- ASIGNAR PRIORIDAD A LAS ALERTAS (ACTUALIZADA CON NUEVA ALERTA) ---
    def asignar_prioridad(row):
        if row['Alerta Vencimiento Mañana'] == "Ya vencido":
            return "ALTA - Ya Vencido"
        elif row['Alerta Pendiente Fuera Tiempo'] == "Fuera de tiempo crítico":
            return "ALTA - Fuera de Tiempo"
        elif row['Alerta Devolución'] == "Sugerir devolución":
            return "ALTA - Devolución Demorada"
        elif row['Alerta Redespacho'] == "Redespacho demorado":
            return "ALTA - Redespacho"
        elif row['Alerta En Tránsito Demorado'] != "":
            return "ALTA - Fuera de Tiempo"
        elif row['Alerta Reprogramada Sin Visitas'] != "":
            return "ALTA - Reprogramada Sin Visita"
        elif row['Alerta Creada Demorada'] != "" and "demorada" in row['Alerta Creada Demorada'].lower():
            return "ALTA - Creada Demorada"
        elif row['Alerta Vencimiento Mañana'] == "Vence mañana":
            return "ALTA - Vence Mañana"
        elif row['Alerta Seguimiento Visitas'] != "":
            return "MEDIA - Seguimiento Visitas"
        elif row['Alerta Una Visita Sin Seguimiento'] != "":
            return "MEDIA - 1 Visita Sin Seg."
        elif row['Alerta Creada Demorada'] != "" and "próxima a vencer" in row['Alerta Creada Demorada']:
            return "MEDIA - Creada Próxima a Vencer"
        elif row['Alerta Pago Pendiente'] == "Pago pendiente demorado":
            return "BAJA - Pago Pendiente"
        else:
            return ""

    df['Prioridad Alerta'] = df.apply(asignar_prioridad, axis=1)

    # Ordenar el DataFrame por prioridad para que las alertas altas aparezcan primero
    prioridad_orden = {
        "ALTA - Ya Vencido": 1,  # Nueva prioridad (más alta)
        "ALTA - Fuera de Tiempo": 2, 
        "ALTA - Devolución Demorada": 3, 
        "ALTA - Redespacho": 4,
        "ALTA - Reprogramada Sin Visita": 5,
        "ALTA - Creada Demorada": 6,
        "ALTA - Vence Mañana": 7,
        "MEDIA - Seguimiento Visitas": 8, 
        "MEDIA - 1 Visita Sin Seg.": 9,
        "MEDIA - Creada Próxima a Vencer": 10,
        "BAJA - Pago Pendiente": 11
    }
    df['Orden Prioridad'] = df['Prioridad Alerta'].map(prioridad_orden).fillna(999)
    df = df.sort_values('Orden Prioridad').reset_index(drop=True)

    # --- FILTROS MEJORADOS CON DEPENDENCIA ---
    st.sidebar.header("🔍 Filtros")

    # Filtro por Cliente
    if 'Cliente' in df.columns:
        clientes = sorted(df['Cliente'].dropna().unique())
        cliente_seleccionado = st.sidebar.selectbox("Cliente", ["Todos"] + clientes)
    else:
        st.error("❌ La columna 'Cliente' no existe en el archivo.")
        st.stop()

    # Crear DataFrame temporal para filtros dependientes
    df_filtrado = df.copy()

    # Aplicar filtro de cliente primero (si se seleccionó)
    if cliente_seleccionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Cliente'] == cliente_seleccionado]

    # Filtro por Subcuenta (DEPENDE del cliente seleccionado)
    if 'Subcuenta' in df_filtrado.columns:
        # Usar df_filtrado en lugar de df
        subcuentas = sorted(df_filtrado['Subcuenta'].dropna().unique())
        subcuenta_seleccionada = st.sidebar.selectbox("Subcuenta", ["Todas"] + subcuentas)
    else:
        st.error("❌ La columna 'Subcuenta' no existe en el archivo.")
        st.stop()

    # Filtro por Agencia origen (también dependiente del cliente)
    if 'Agencia origen' in df_filtrado.columns:
        agencias_origen = sorted(df_filtrado['Agencia origen'].dropna().unique())
        agencia_origen_seleccionada = st.sidebar.selectbox("Agencia origen", ["Todas"] + agencias_origen)
    else:
        st.warning("⚠️ La columna 'Agencia origen' no existe. Se omitirá este filtro.")
        agencia_origen_seleccionada = "Todas"

    # Filtro por Agencia destino (dependiente del cliente)
    if 'Agencia destino' in df_filtrado.columns:
        agencias = sorted(df_filtrado['Agencia destino'].dropna().unique())
        agencia_seleccionada = st.sidebar.selectbox("Agencia destino", ["Todas"] + agencias)
    else:
        st.error("❌ La columna 'Agencia destino' no existe en el archivo.")
        st.stop()

    # Filtro por ZONA (dependiente de los filtros anteriores)
    if 'ZONA' in df_filtrado.columns:
        zonas = sorted(df_filtrado['ZONA'].dropna().unique())
        zona_seleccionada = st.sidebar.selectbox("Zona", ["Todas"] + zonas)
    else:
        st.warning("⚠️ La columna 'ZONA' no existe. Se omitirá este filtro.")
        zona_seleccionada = "Todas"        

    # Filtro por ED (dependiente de los filtros anteriores)
    if 'ED' in df_filtrado.columns:
        ed_opciones = sorted(df_filtrado['ED'].dropna().unique())
        ed_seleccionada = st.sidebar.selectbox("Entrega a Domicilio (ED)", ["Todas"] + ed_opciones)
    else:
        st.warning("⚠️ La columna 'ED' no existe. Se omitirá este filtro.")
        ed_seleccionada = "Todas"

    # Filtro por Condición de venta (dependiente de los filtros anteriores)
    if 'Condición de venta' in df_filtrado.columns:
        condiciones_venta = sorted(df_filtrado['Condición de venta'].dropna().unique())
        condicion_venta_seleccionada = st.sidebar.selectbox("Condición de venta", ["Todas"] + condiciones_venta)
    else:
        st.warning("⚠️ La columna 'Condición de venta' no existe. Se omitirá este filtro.")
        condicion_venta_seleccionada = "Todas"

    # Ahora aplicar TODOS los filtros al DataFrame original
    df_final = df.copy()

    if cliente_seleccionado != "Todos":
        df_final = df_final[df_final['Cliente'] == cliente_seleccionado]
    if subcuenta_seleccionada != "Todas":
        df_final = df_final[df_final['Subcuenta'] == subcuenta_seleccionada]
    if 'Agencia origen' in df_final.columns and agencia_origen_seleccionada != "Todas":
        df_final = df_final[df_final['Agencia origen'] == agencia_origen_seleccionada]
    if agencia_seleccionada != "Todas":
        df_final = df_final[df_final['Agencia destino'] == agencia_seleccionada]
    if 'ZONA' in df_final.columns and zona_seleccionada != "Todas":
        df_final = df_final[df_final['ZONA'] == zona_seleccionada]    
    if 'ED' in df_final.columns and ed_seleccionada != "Todas":
        df_final = df_final[df_final['ED'] == ed_seleccionada]
    if 'Condición de venta' in df_final.columns and condicion_venta_seleccionada != "Todas":
        df_final = df_final[df_final['Condición de venta'] == condicion_venta_seleccionada]

    # Reemplazar el DataFrame original con el filtrado
    df = df_final

    # --- NUEVA SECCIÓN: PORCENTAJE DE CUMPLIMIENTO POR SEMANA CON ALERTAS DE VARIACIÓN ---
    st.header("📈 Porcentaje de Cumplimiento por Semana con Alertas de Variación")

    # CORREGIDO: Calcular el porcentaje de cumplimiento por semana
    def calcular_cumplimiento_semana(grupo):
        # EXCLUIR CANCELADAS del total
        total_pedidos_semana = grupo[grupo['Cumplimiento'] != "Cancelada"].shape[0]
        if total_pedidos_semana == 0:
            return 0
        
        # Contar SOLO entregas en tiempo (no incluir visitas en tiempo)
        cumplidos_semana = grupo[
            grupo['Cumplimiento'].isin([
                "Entregada - En Tiempo", 
                "Entregada - En Tiempo (PD: Pago Pendiente)"
            ])
        ].shape[0]
        
        return (cumplidos_semana / total_pedidos_semana * 100)

    # Agrupar por semana y calcular el porcentaje de cumplimiento
    df_semana = df[df['Cumplimiento'] != "Cancelada"].groupby('Semana Calendario').apply(
        calcular_cumplimiento_semana
    ).reset_index(name='Porcentaje Cumplimiento')

    # Ordenar por semana para calcular variaciones
    df_semana = df_semana.sort_values('Semana Calendario').reset_index(drop=True)

    # Calcular variación respecto a la semana anterior
    df_semana['Variación vs Semana Anterior'] = df_semana['Porcentaje Cumplimiento'].diff()

    # Calcular variación porcentual
    df_semana['Variación Porcentual'] = (df_semana['Porcentaje Cumplimiento'].pct_change() * 100).round(2)

    # Formatear el porcentaje
    df_semana['Porcentaje Cumplimiento'] = df_semana['Porcentaje Cumplimiento'].round(2)

    # --- FUNCIÓN PARA GENERAR ALERTAS DE VARIACIÓN ---
    def generar_alerta_variacion(row):
        variacion = row['Variación vs Semana Anterior']
        variacion_porcentual = row['Variación Porcentual']
        
        if pd.isna(variacion):
            return "🔵 Semana de referencia"
        elif variacion > 5:  # Mejora significativa (>5 puntos)
            return f"🟢 Excelente! +{variacion:.1f}pts (+{variacion_porcentual:.1f}%)"
        elif variacion > 2:  # Mejora moderada
            return f"🟡 Mejoró +{variacion:.1f}pts (+{variacion_porcentual:.1f}%)"
        elif variacion >= -2:  # Estable
            return f"⚪ Estable {variacion:+.1f}pts"
        elif variacion > -5:  # Caída moderada
            return f"🟠 Alerta! {variacion:.1f}pts ({variacion_porcentual:+.1f}%)"
        else:  # Caída significativa
            return f"🔴 CRÍTICO! {variacion:.1f}pts ({variacion_porcentual:+.1f}%)"

    df_semana['Alerta Variación'] = df_semana.apply(generar_alerta_variacion, axis=1)

    # Mostrar la tabla de semanas con alertas
    st.subheader("Tabla de Cumplimiento por Semana con Alertas de Variación")
    st.dataframe(df_semana[['Semana Calendario', 'Porcentaje Cumplimiento', 'Alerta Variación']], 
                use_container_width=True)

    # --- GRÁFICO MEJORADO CON ANOTACIONES DE VARIACIÓN ---
    if len(df_semana) > 1:
        fig_semana = px.line(
            df_semana,
            x='Semana Calendario',
            y='Porcentaje Cumplimiento',
            title='Evolución del Porcentaje de Cumplimiento por Semana con Alertas de Variación',
            markers=True,
            line_shape='linear'
        )
        
        # Agregar anotaciones para las variaciones significativas
        for i, row in df_semana.iterrows():
            if i > 0:  # No aplicar a la primera semana
                variacion = row['Variación vs Semana Anterior']
                if abs(variacion) >= 2:  # Solo anotar variaciones significativas
                    color = 'green' if variacion > 0 else 'red'
                    fig_semana.add_annotation(
                        x=row['Semana Calendario'],
                        y=row['Porcentaje Cumplimiento'],
                        text=f"{variacion:+.1f}pts",
                        showarrow=True,
                        arrowhead=2,
                        arrowsize=1,
                        arrowwidth=2,
                        arrowcolor=color,
                        bgcolor=color,
                        bordercolor=color,
                        font=dict(color='white', size=10)
                    )
        
        # Mejorar formato del gráfico
        fig_semana.update_layout(
            xaxis_title='Semana Calendario',
            yaxis_title='Porcentaje de Cumplimiento (%)',
            yaxis=dict(range=[0, 100]),
            hovermode='x unified'
        )
        
        # Agregar línea de referencia
        fig_semana.add_hline(
            y=80, 
            line_dash="dash", 
            line_color="red",
            annotation_text="Objetivo 80%"
        )
        
        st.plotly_chart(fig_semana, use_container_width=True)

    # --- RESUMEN DE TENDENCIAS ---
    st.subheader("📊 Resumen de Tendencias por Semana")

    if len(df_semana) > 1:
        # Calcular métricas de tendencia
        ultima_semana = df_semana.iloc[-1]
        penultima_semana = df_semana.iloc[-2] if len(df_semana) > 1 else None
        
        mejora_semanas = (df_semana['Variación vs Semana Anterior'] > 0).sum()
        total_comparables = len(df_semana) - 1  # Excluir la primera semana
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if penultima_semana is not None:
                variacion_actual = ultima_semana['Variación vs Semana Anterior']
                if not pd.isna(variacion_actual):
                    if variacion_actual > 0:
                        st.success(f"📈 Semana {ultima_semana['Semana Calendario']}: **+{variacion_actual:.1f}pts** vs semana anterior")
                    else:
                        st.error(f"📉 Semana {ultima_semana['Semana Calendario']}: **{variacion_actual:.1f}pts** vs semana anterior")
        
        with col2:
            if total_comparables > 0:
                tasa_mejora = (mejora_semanas / total_comparables) * 100
                st.metric("📊 Tasa de Mejora Semanal", f"{tasa_mejora:.1f}%")
        
        with col3:
            if len(df_semana) >= 4:
                # Tendencia últimas 4 semanas
                ultimas_4 = df_semana.tail(4)
                tendencia = ultimas_4['Porcentaje Cumplimiento'].mean()
                st.metric("📅 Promedio Últimas 4 Semanas", f"{tendencia:.1f}%")

    # --- NOTIFICACIONES PUSH PARA VARIACIONES CRÍTICAS ---
    st.header("🔔 Notificaciones de Variación en Tiempo Real")

    if len(df_semana) > 1:
        ultima_semana = df_semana.iloc[-1]
        variacion_actual = ultima_semana['Variación vs Semana Anterior']
        
        if not pd.isna(variacion_actual):
            if variacion_actual < -10:
                st.error(f"""
                🚨 **ALERTA CRÍTICA** 
                **Caída drástica en la última semana:** {variacion_actual:.1f} puntos
                **Recomendación:** Revisar procesos urgentemente
                """)
            elif variacion_actual < -5:
                st.warning(f"""
                ⚠️ **ALERTA IMPORTANTE**
                **Caída significativa en la última semana:** {variacion_actual:.1f} puntos
                **Recomendación:** Analizar causas y tomar acciones
                """)
            elif variacion_actual > 10:
                st.success(f"""
                🎉 **LOGRO DESTACADO**
                **Mejora excepcional en la última semana:** +{variacion_actual:.1f} puntos
                **Recomendación:** Replicar buenas prácticas
                """)
            elif variacion_actual > 5:
                st.info(f"""
                👍 **BUEN DESEMPEÑO**
                **Mejora significativa en la última semana:** +{variacion_actual:.1f} puntos
                **Recomendación:** Mantener tendencia positiva
                """)

    # --- NUEVA SECCIÓN: PORCENTAJE DE CUMPLIMIENTO POR SEMANA Y ZONA CON ALERTAS DE VARIACIÓN ---
    st.header("📈 Porcentaje de Cumplimiento por Semana y Zona")

    # CORREGIDO: Calcular el porcentaje de cumplimiento por semana y zona
    def calcular_cumplimiento_semana_zona(grupo):
        # EXCLUIR CANCELADAS del total
        total_pedidos_semana = grupo[grupo['Cumplimiento'] != "Cancelada"].shape[0]
        if total_pedidos_semana == 0:
            return 0
        
        # Contar SOLO entregas en tiempo (no incluir visitas en tiempo)
        cumplidos_semana = grupo[
            grupo['Cumplimiento'].isin([
                "Entregada - En Tiempo", 
                "Entregada - En Tiempo (PD: Pago Pendiente)"
            ])
        ].shape[0]
        
        return (cumplidos_semana / total_pedidos_semana * 100)

    # Agrupar por semana y zona para calcular el porcentaje de cumplimiento
    df_semana_zona = df[df['Cumplimiento'] != "Cancelada"].groupby(['Semana Calendario', 'ZONA']).apply(
        calcular_cumplimiento_semana_zona
    ).reset_index(name='Porcentaje Cumplimiento')

    # También calcular el total por semana
    df_semana_total = df[df['Cumplimiento'] != "Cancelada"].groupby('Semana Calendario').apply(
        calcular_cumplimiento_semana_zona
    ).reset_index(name='Porcentaje Cumplimiento')
    df_semana_total['ZONA'] = 'TOTAL'

    # Combinar ambos DataFrames
    df_semana_completo = pd.concat([df_semana_zona, df_semana_total], ignore_index=True)

    # Ordenar por semana y zona para calcular variaciones
    df_semana_completo = df_semana_completo.sort_values(['Semana Calendario', 'ZONA']).reset_index(drop=True)

    # Calcular variación respecto a la semana anterior por zona
    df_semana_completo['Variación vs Semana Anterior'] = df_semana_completo.groupby('ZONA')['Porcentaje Cumplimiento'].diff()

    # Calcular variación porcentual por zona
    df_semana_completo['Variación Porcentual'] = (df_semana_completo.groupby('ZONA')['Porcentaje Cumplimiento'].pct_change() * 100).round(2)

    # Formatear el porcentaje
    df_semana_completo['Porcentaje Cumplimiento'] = df_semana_completo['Porcentaje Cumplimiento'].round(2)

    # --- FUNCIÓN PARA GENERAR ALERTAS DE VARIACIÓN POR ZONA ---
    def generar_alerta_variacion_zona(row):
        variacion = row['Variación vs Semana Anterior']
        variacion_porcentual = row['Variación Porcentual']
        
        if pd.isna(variacion):
            return "🔵 Semana de referencia"
        elif variacion > 5:  # Mejora significativa (>5 puntos)
            return f"🟢 Excelente! +{variacion:.1f}pts (+{variacion_porcentual:.1f}%)"
        elif variacion > 2:  # Mejora moderada
            return f"🟡 Mejoró +{variacion:.1f}pts (+{variacion_porcentual:.1f}%)"
        elif variacion >= -2:  # Estable
            return f"⚪ Estable {variacion:+.1f}pts"
        elif variacion > -5:  # Caída moderada
            return f"🟠 Alerta! {variacion:.1f}pts ({variacion_porcentual:+.1f}%)"
        else:  # Caída significativa
            return f"🔴 CRÍTICO! {variacion:.1f}pts ({variacion_porcentual:+.1f}%)"

    df_semana_completo['Alerta Variación'] = df_semana_completo.apply(generar_alerta_variacion_zona, axis=1)

    # Mostrar la tabla de semanas con alertas por zona
    st.subheader("Tabla de Cumplimiento por Semana y Zona con Alertas de Variación")

    # Pivotar la tabla para mejor visualización
    df_pivot = df_semana_completo.pivot_table(
        index='Semana Calendario', 
        columns='ZONA', 
        values=['Porcentaje Cumplimiento', 'Alerta Variación'],
        aggfunc='first'
    )

    # Reorganizar las columnas para mejor presentación
    df_display = pd.DataFrame()
    for semana in df_pivot.index:
        row_data = {'Semana Calendario': semana}
        
        for zona in ['AMBA', 'INTERIOR', 'TOTAL']:
            if zona in df_pivot['Porcentaje Cumplimiento'].columns:
                row_data[f'{zona} - % Cumplimiento'] = df_pivot['Porcentaje Cumplimiento'][zona][semana]
                row_data[f'{zona} - Alerta'] = df_pivot['Alerta Variación'][zona][semana]
        
        df_display = pd.concat([df_display, pd.DataFrame([row_data])], ignore_index=True)

    # Ordenar por semana
    df_display = df_display.sort_values('Semana Calendario').reset_index(drop=True)

    # Formatear porcentajes
    for col in df_display.columns:
        if '% Cumplimiento' in col:
            df_display[col] = df_display[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A")

    st.dataframe(df_display, use_container_width=True)

    # --- GRÁFICO MEJORADO CON LÍNEAS POR ZONA Y ANOTACIONES ---
    if len(df_semana_completo) > 1:
        fig_semana_zona = px.line(
            df_semana_completo,
            x='Semana Calendario',
            y='Porcentaje Cumplimiento',
            color='ZONA',
            title='Evolución del Porcentaje de Cumplimiento por Semana y Zona',
            markers=True,
            line_shape='linear',
            color_discrete_map={
                'AMBA': '#28a745',
                'INTERIOR': '#007bff', 
                'TOTAL': '#ff6b00'
            }
        )
        
        # Agregar anotaciones para variaciones significativas por zona
        for zona in ['AMBA', 'INTERIOR', 'TOTAL']:
            df_zona = df_semana_completo[df_semana_completo['ZONA'] == zona].sort_values('Semana Calendario')
            for i, row in df_zona.iterrows():
                if i > 0:  # No aplicar a la primera semana de cada zona
                    variacion = row['Variación vs Semana Anterior']
                    if not pd.isna(variacion) and abs(variacion) >= 2:  # Solo anotar variaciones significativas
                        color = 'green' if variacion > 0 else 'red'
                        # Posicionar las anotaciones para evitar superposición
                        y_offset = 0
                        if zona == 'AMBA':
                            y_offset = 3
                        elif zona == 'INTERIOR':
                            y_offset = -3
                        # Para TOTAL, no aplicar offset o aplicar uno diferente
                        
                        fig_semana_zona.add_annotation(
                            x=row['Semana Calendario'],
                            y=row['Porcentaje Cumplimiento'] + y_offset,
                            text=f"{variacion:+.1f}",
                            showarrow=True,
                            arrowhead=2,
                            arrowsize=1,
                            arrowwidth=2,
                            arrowcolor=color,
                            bgcolor=color,
                            bordercolor=color,
                            font=dict(color='white', size=8),
                            yshift=10 if variacion > 0 else -10
                        )
        
        # Mejorar formato del gráfico
        fig_semana_zona.update_layout(
            xaxis_title='Semana Calendario',
            yaxis_title='Porcentaje de Cumplimiento (%)',
            yaxis=dict(range=[0, 100]),
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        # Agregar línea de referencia
        fig_semana_zona.add_hline(
            y=80, 
            line_dash="dash", 
            line_color="red",
            annotation_text="Objetivo 80%"
        )
        
        st.plotly_chart(fig_semana_zona, use_container_width=True)

    # --- RESUMEN DE TENDENCIAS POR ZONA ---
    st.subheader("📊 Resumen de Tendencias por Semana y Zona")

    if len(df_semana_completo) > 1:
        # Crear columnas para cada zona
        zonas = ['AMBA', 'INTERIOR', 'TOTAL']
        cols = st.columns(3)
        
        for idx, zona in enumerate(zonas):
            with cols[idx]:
                st.subheader(f"Zona {zona}")
                
                # Filtrar datos de la zona
                df_zona = df_semana_completo[df_semana_completo['ZONA'] == zona].sort_values('Semana Calendario')
                
                if len(df_zona) > 1:
                    ultima_semana = df_zona.iloc[-1]
                    penultima_semana = df_zona.iloc[-2] if len(df_zona) > 1 else None
                    
                    if penultima_semana is not None:
                        variacion_actual = ultima_semana['Variación vs Semana Anterior']
                        if not pd.isna(variacion_actual):
                            if variacion_actual > 0:
                                st.success(f"📈 **+{variacion_actual:.1f}pts** vs semana anterior")
                            else:
                                st.error(f"📉 **{variacion_actual:.1f}pts** vs semana anterior")
                    
                    # Mostrar métricas clave
                    st.metric(
                        f"Última Semana {ultima_semana['Semana Calendario']}",
                        f"{ultima_semana['Porcentaje Cumplimiento']:.1f}%"
                    )
                    
                    if len(df_zona) >= 4:
                        ultimas_4 = df_zona.tail(4)
                        tendencia = ultimas_4['Porcentaje Cumplimiento'].mean()
                        st.metric("Promedio Últimas 4 Semanas", f"{tendencia:.1f}%")

    # --- ANÁLISIS COMPARATIVO ENTRE ZONAS ---
    st.header("📊 Análisis Comparativo entre Zonas")

    if len(df_semana_completo) > 1:
        # Calcular diferencia AMBA vs INTERIOR por semana
        df_amba = df_semana_completo[df_semana_completo['ZONA'] == 'AMBA'][['Semana Calendario', 'Porcentaje Cumplimiento']]
        df_interior = df_semana_completo[df_semana_completo['ZONA'] == 'INTERIOR'][['Semana Calendario', 'Porcentaje Cumplimiento']]
        
        df_comparativo = pd.merge(df_amba, df_interior, on='Semana Calendario', suffixes=('_AMBA', '_INTERIOR'))
        df_comparativo['Diferencia (AMBA - INTERIOR)'] = df_comparativo['Porcentaje Cumplimiento_AMBA'] - df_comparativo['Porcentaje Cumplimiento_INTERIOR']
        
        # Gráfico de diferencias
        fig_diferencias = px.bar(
            df_comparativo,
            x='Semana Calendario',
            y='Diferencia (AMBA - INTERIOR)',
            title='Diferencia de Cumplimiento: AMBA vs INTERIOR',
            color='Diferencia (AMBA - INTERIOR)',
            color_continuous_scale='RdYlGn',
            color_continuous_midpoint=0
        )
        
        fig_diferencias.update_layout(
            xaxis_title='Semana Calendario',
            yaxis_title='Diferencia de Cumplimiento (%)',
            hovermode='x unified'
        )
        
        # Agregar línea en cero
        fig_diferencias.add_hline(y=0, line_dash="solid", line_color="black")
        
        st.plotly_chart(fig_diferencias, use_container_width=True)
        
        # Resumen de diferencias
        st.subheader("🔍 Resumen de Diferencias AMBA vs INTERIOR")
        
        if len(df_comparativo) > 0:
            ultima_diferencia = df_comparativo.iloc[-1]['Diferencia (AMBA - INTERIOR)']
            promedio_diferencia = df_comparativo['Diferencia (AMBA - INTERIOR)'].mean()
            
            col1, col2 = st.columns(2)
            
            with col1:
                if ultima_diferencia > 0:
                    st.success(f"**Última semana:** AMBA +{ultima_diferencia:.1f}pts sobre INTERIOR")
                else:
                    st.error(f"**Última semana:** INTERIOR +{abs(ultima_diferencia):.1f}pts sobre AMBA")
            
            with col2:
                if promedio_diferencia > 0:
                    st.info(f"**Promedio histórico:** AMBA +{promedio_diferencia:.1f}pts sobre INTERIOR")
                else:
                    st.warning(f"**Promedio histórico:** INTERIOR +{abs(promedio_diferencia):.1f}pts sobre AMBA")

    # --- NOTIFICACIONES PUSH PARA VARIACIONES CRÍTICAS POR ZONA ---
    st.header("🔔 Notificaciones de Variación en Tiempo Real por Zona")

    if len(df_semana_completo) > 1:
        for zona in ['AMBA', 'INTERIOR', 'TOTAL']:
            df_zona = df_semana_completo[df_semana_completo['ZONA'] == zona].sort_values('Semana Calendario')
            
            if len(df_zona) > 1:
                ultima_semana = df_zona.iloc[-1]
                variacion_actual = ultima_semana['Variación vs Semana Anterior']
                
                if not pd.isna(variacion_actual):
                    if variacion_actual < -10:
                        st.error(f"""
                        🚨 **ALERTA CRÍTICA - {zona}** 
                        **Caída drástica en la última semana:** {variacion_actual:.1f} puntos
                        **Recomendación:** Revisar procesos urgentemente en zona {zona}
                        """)
                    elif variacion_actual < -5:
                        st.warning(f"""
                        ⚠️ **ALERTA IMPORTANTE - {zona}**
                        **Caída significativa en la última semana:** {variacion_actual:.1f} puntos
                        **Recomendación:** Analizar causas y tomar acciones en zona {zona}
                        """)
                    elif variacion_actual > 10:
                        st.success(f"""
                        🎉 **LOGRO DESTACADO - {zona}**
                        **Mejora excepcional en la última semana:** +{variacion_actual:.1f} puntos
                        **Recomendación:** Replicar buenas prácticas de zona {zona}
                        """)
                    elif variacion_actual > 5:
                        st.info(f"""
                        👍 **BUEN DESEMPEÑO - {zona}**
                        **Mejora significativa en la última semana:** +{variacion_actual:.1f} puntos
                        **Recomendación:** Mantener tendencia positiva en zona {zona}
                        """)

    # --- AGREGAR ALERTAS AL DATAFRAME PRINCIPAL ---
    # Crear diccionarios de mapeo
    mapeo_semana = df_semana.set_index('Semana Calendario')['Porcentaje Cumplimiento'].to_dict()
    mapeo_alerta = df_semana.set_index('Semana Calendario')['Alerta Variación'].to_dict()
    mapeo_variacion = df_semana.set_index('Semana Calendario')['Variación vs Semana Anterior'].to_dict()

    # Aplicar mapeos al DataFrame principal
    df['Porcentaje Cumplimiento Semana'] = df['Semana Calendario'].map(mapeo_semana)
    df['Alerta Variación Semana'] = df['Semana Calendario'].map(mapeo_alerta)
    df['Variación vs Semana Anterior'] = df['Semana Calendario'].map(mapeo_variacion)

    # Mover las columnas al lado de "Semana Calendario"
    columnas = df.columns.tolist()
    pos_semana = columnas.index('Semana Calendario')

    # Insertar las nuevas columnas después de la semana
    nuevas_columnas = ['Porcentaje Cumplimiento Semana', 'Alerta Variación Semana', 'Variación vs Semana Anterior']
    for i, col in enumerate(nuevas_columnas):
        columnas.insert(pos_semana + 1 + i, col)
        columnas.remove(col)

    # Reordenar el DataFrame
    df = df[columnas]

    # Formatear columnas
    df['Porcentaje Cumplimiento Semana'] = df['Porcentaje Cumplimiento Semana'].apply(
        lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A"
    )
    df['Variación vs Semana Anterior'] = df['Variación vs Semana Anterior'].apply(
        lambda x: f"{x:+.1f} pts" if pd.notna(x) else "N/A"
    )

    # --- ALERTAS CRÍTICAS EN EL SIDEBAR ---
    st.sidebar.header("🚨 Alertas Críticas de Variación")

    if len(df_semana) > 1:
        # Buscar variaciones críticas (caídas > 5 puntos)
        alertas_criticas = df_semana[
            (df_semana['Variación vs Semana Anterior'] < -5) & 
            (pd.notna(df_semana['Variación vs Semana Anterior']))
        ]
        
        if not alertas_criticas.empty:
            st.sidebar.error("### 📉 Caídas Significativas")
            for _, alerta in alertas_criticas.iterrows():
                st.sidebar.write(f"**Semana {alerta['Semana Calendario']}**: {alerta['Variación vs Semana Anterior']:.1f}pts")
        
        # Buscar mejoras significativas
        mejoras_significativas = df_semana[
            (df_semana['Variación vs Semana Anterior'] > 5) & 
            (pd.notna(df_semana['Variación vs Semana Anterior']))
        ]
        
        if not mejoras_significativas.empty:
            st.sidebar.success("### 📈 Mejoras Significativas")
            for _, mejora in mejoras_significativas.iterrows():
                st.sidebar.write(f"**Semana {mejora['Semana Calendario']}**: +{mejora['Variación vs Semana Anterior']:.1f}pts")

    # --- ESTADÍSTICAS MEJORADAS ---
    st.header("📊 Indicadores/Alertas")

    # El total de pedidos ahora EXCLUYE las canceladas
    total_pedidos = df[df['Cumplimiento'] != "Cancelada"].shape[0]
    entregados = df[df['Cumplimiento'].str.startswith("Entregada")].shape[0]
    devueltos = df[df['Cumplimiento'] == "Devuelto"].shape[0]
    canceladas = df[df['Cumplimiento'] == "Cancelada"].shape[0]
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

    # --- CORREGIDO: NUEVOS INDICADORES MEJORADOS ---
    # 1. SLA Principal (según tu consulta): Entregas en tiempo / Total pedidos (excl. canceladas)
    sla_principal = ((en_tiempo + en_tiempo_pd) / total_pedidos * 100) if total_pedidos > 0 else 0

    # 2. Cumplimiento tradicional (solo entregados en tiempo vs total entregados)
    cumplimiento_tradicional = ((en_tiempo + en_tiempo_pd) / entregados * 100) if entregados > 0 else 0

    # 3. Cumplimiento de gestión (entregados + visitas en tiempo)
    cumplimiento_gestion = ((en_tiempo + en_tiempo_pd + visita_en_tiempo) / total_pedidos * 100) if total_pedidos > 0 else 0

    # 4. Efectividad de visitas
    total_visitas = visita_en_tiempo + visita_fuera_tiempo
    efectividad_visitas = (visita_en_tiempo / total_visitas * 100) if total_visitas > 0 else 0

    # 5. Tasa de resolución (entregados / total gestionado)
    tasa_resolucion = (entregados / (entregados + visita_en_tiempo + visita_fuera_tiempo) * 100) if (entregados + visita_en_tiempo + visita_fuera_tiempo) > 0 else 0

    # --- NUEVOS KPIs DE EFICIENCIA ---
    # Primer Intento de Entrega (First Attempt Delivery Rate - FADR)
    primer_intento_entrega = df[
        (df['Cumplimiento'].str.startswith("Entregada")) &
        (df.get('Visitas', 0) <= 1)
    ].shape[0]

    fadr = (primer_intento_entrega / entregados * 100) if entregados > 0 else 0

    # Pedidos por Visita (Solo para entregados, para no distorsionar)
    total_visitas_entregados = df[df['Cumplimiento'].str.startswith("Entregada")]['Visitas'].sum()
    pedidos_con_visita = df[(df['Cumplimiento'].str.startswith("Entregada")) & (df.get('Visitas', 0) >= 1)].shape[0]
    pedidos_por_visita = (pedidos_con_visita / total_visitas_entregados) if total_visitas_entregados > 0 else 0

    # --- KPI MEJORADO: TASA DE RECHAZO/AUSENCIA - VERSIÓN ROBUSTA CON REGEX ---
    def es_rechazo_ausente_regex(estado):
        """Detecta rechazo/ausencia usando expresiones regulares para mayor precisión"""
        if pd.isna(estado):
            return False
        
        estado_str = str(estado).strip()
        
        # Patrones regex para detectar motivos
        patrones = [
            r'\[Motivo:\s*(Rechazado|Ausente)',  # [Motivo: Rechazado] o [Motivo: Ausente]
            r'\[Motivo:\s*.*(rechaz|ausent)',    # Cualquier variación
            r'(rechazado|ausente).*\[Motivo:',   # Formato inverso
            r'cliente\s+(rechazó|no aceptó|ausente|no se presentó)',
            r'motivo.*rechaz|motivo.*ausent'     # Otras variaciones
        ]
        
        for patron in patrones:
            if re.search(patron, estado_str, re.IGNORECASE):
                return True
        
        return False

    # Aplicar la versión con regex
    rechazos_ausentes = df[
        df['Estado'].apply(es_rechazo_ausente_regex) &
        (df['Visitas'] > 0)
    ].shape[0]

    total_con_visita = df[df['Visitas'] > 0].shape[0]

    tasa_rechazo_ausencia = (rechazos_ausentes / total_con_visita * 100) if total_con_visita > 0 else 0

    # --- NUEVAS MÉTRICAS PARA ALERTAS DE "CREADA" ---
    alertas_creada_criticas = df[df['Alerta Creada Demorada'].str.contains("demorada", na=False)].shape[0]
    alertas_creada_preventivas = df[df['Alerta Creada Demorada'].str.contains("próxima a vencer", na=False)].shape[0]

    # --- PRESENTACIÓN DE MÉTRICAS EN 5 COLUMNAS (ACTUALIZADAS) ---

    col1, col2, col3, col4, col5 = st.columns(5)

    # Columna 1: Volumen
    with col1:
        st.metric("📦 Total Pedidos (Excl. Canceladas)", total_pedidos)
        st.metric("🎯 SLA Principal", f"{sla_principal:.1f}%")

    # Columna 2: Entregas y Real
    with col2:
        st.metric("✅ Entregados", entregados, f"{(entregados/total_pedidos*100):.1f}%")
        st.metric("📊 Cumplimiento Entregas", f"{cumplimiento_tradicional:.1f}%")

    # Columna 3: Devueltos y Visitas
    with col3:
        st.metric("🔄 Devueltos", devueltos, f"{(devueltos/total_pedidos*100):.1f}%")
        st.metric("📋 Cumplimiento Gestión", f"{cumplimiento_gestion:.1f}%")

    # Columna 4: Pendientes y Rechazos
    with col4:
        st.metric("⏳ Pendientes", pendientes_reales, f"{(pendientes_reales/total_pedidos*100):.1f}%")
        st.metric("🚫 Tasa Rechazo/Ausencia", f"{tasa_rechazo_ausencia:.1f}%")
        
    # Columna 5: Canceladas y Alertas Creada
    with col5:
        st.metric("❌ Canceladas", canceladas, f"{(canceladas/(total_pedidos + canceladas)*100):.1f}%")
        st.metric("🚨 Creadas Demoradas (>24h)", alertas_creada_criticas)

    # --- DEBUG: Mostrar ejemplos de rechazo/ausencia detectados ---
    if st.sidebar.checkbox("🔍 Mostrar pedidos con rechazo/ausencia detectados"):
        ejemplos_rechazo = df[df['Estado'].apply(es_rechazo_ausente_regex) & (df['Visitas'] > 0)]
        if not ejemplos_rechazo.empty:
            st.sidebar.write(f"📋 Ejemplos detectados ({len(ejemplos_rechazo)}):")
            st.sidebar.dataframe(ejemplos_rechazo[['Guia', 'Estado', 'Visitas']].head(5))
        else:
            st.sidebar.info("No se detectaron pedidos con rechazo/ausencia")

    # --- TABLA DE RESUMEN MEJORADA (ACTUALIZADA) ---
    st.header("📈 Detalle de Estados")
    resumen_data = {
        "Categoría": [
            "TOTAL PEDIDOS (Excl. Canceladas)",
            "ENTREGADOS",
            " - En Tiempo",
            " - En Tiempo (PD)",
            " - Fuera de Tiempo", 
            " - Fuera de Tiempo (PD)",
            "DEVUELTOS",
            "CANCELADAS",
            "PENDIENTES CON VISITA",
            " - Visita en Tiempo",
            " - Visita Fuera de Tiempo",
            "PENDIENTES SIN VISITA",
            " - En Tiempo",
            " - Último Día",
            " - Fuera de Tiempo",
            "SLA PRINCIPAL (En Tiempo/Total)",
            "TASA RECHAZO/AUSENCIA"
        ],
        "Cantidad": [
            total_pedidos,
            entregados,
            en_tiempo,
            en_tiempo_pd,
            fuera_tiempo,
            fuera_tiempo_pd,
            devuelto_count,
            canceladas,
            visita_en_tiempo + visita_fuera_tiempo,
            visita_en_tiempo,
            visita_fuera_tiempo,
            pendiente_en_tiempo + pendiente_ultimo_dia + pendiente_fuera_tiempo,
            pendiente_en_tiempo,
            pendiente_ultimo_dia,
            pendiente_fuera_tiempo,
            "",  # Vacío para cantidad
            rechazos_ausentes
        ],
        "Porcentaje": [
            "100%",
            f"{(entregados/total_pedidos*100):.1f}%",
            f"{(en_tiempo/total_pedidos*100):.1f}%",
            f"{(en_tiempo_pd/total_pedidos*100):.1f}%",
            f"{(fuera_tiempo/total_pedidos*100):.1f}%",
            f"{(fuera_tiempo_pd/total_pedidos*100):.1f}%",
            f"{(devuelto_count/total_pedidos*100):.1f}%",
            f"{(canceladas/(total_pedidos + canceladas)*100):.1f}%",  # Sobre total incluyendo canceladas
            f"{((visita_en_tiempo + visita_fuera_tiempo)/total_pedidos*100):.1f}%",
            f"{(visita_en_tiempo/total_pedidos*100):.1f}%",
            f"{(visita_fuera_tiempo/total_pedidos*100):.1f}%",
            f"{((pendiente_en_tiempo + pendiente_ultimo_dia + pendiente_fuera_tiempo)/total_pedidos*100):.1f}%",
            f"{(pendiente_en_tiempo/total_pedidos*100):.1f}%",
            f"{(pendiente_ultimo_dia/total_pedidos*100):.1f}%",
            f"{(pendiente_fuera_tiempo/total_pedidos*100):.1f}%",
            f"{sla_principal:.1f}%",
            f"{tasa_rechazo_ausencia:.1f}%"
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
        "Cancelada",
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
        canceladas,
        visita_en_tiempo,
        visita_fuera_tiempo,
        pendiente_en_tiempo, 
        pendiente_ultimo_dia,
        pendiente_fuera_tiempo
    ]
    fig1 = px.pie(
        names=categorias_mejoradas,
        values=valores_mejorados,
        title="Distribución de Cumplimiento Mejorado (Incluyendo Visitas y Canceladas)",
        color=categorias_mejoradas,
        color_discrete_map={
            "Entregada - En Tiempo": "#28a745",
            "Entregada - En Tiempo (PD)": "#2ecc71",
            "Entregada - Fuera de Tiempo": "#dc3545",
            "Entregada - Fuera de Tiempo (PD)": "#e74c3c",
            "Devuelto": "#9b59b6",
            "Cancelada": "#95a5a6",
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

    # --- GRÁFICO COMPARATIVO DE INDICADORES (ACTUALIZADO) ---
    st.header("📈 Comparativa de Indicadores de Cumplimiento")
    indicadores_comparativa = {
        "Indicador": ["SLA Principal", "Cumplimiento Entregas", "Cumplimiento Gestión", "Tasa Rechazo/Ausencia"],
        "Porcentaje": [sla_principal, cumplimiento_tradicional, cumplimiento_gestion, tasa_rechazo_ausencia],
        "Descripción": [
            "Entregas en tiempo / Total pedidos",
            "Entregas en tiempo / Total entregados", 
            "Gestión total (entregas + visitas en tiempo)",
            "Rechazos y ausencias / Total con visita"
        ]
    }
    df_comparativa = pd.DataFrame(indicadores_comparativa)
    fig2 = px.bar(
        df_comparativa,
        x="Indicador",
        y="Porcentaje",
        color="Indicador",
        text="Porcentaje",
        title="Comparativa de Diferentes Indicadores de Cumplimiento",
        color_discrete_map={
            "SLA Principal": "#28a745",
            "Cumplimiento Entregas": "#ffc107",
            "Cumplimiento Gestión": "#007bff",
            "Tasa Rechazo/Ausencia": "#dc3545"
        }
    )
    fig2.update_traces(texttemplate='%{y:.1f}%', textposition='outside')
    fig2.update_layout(showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

    # --- GRÁFICO DE CUMPLIMIENTO REAL POR CLIENTE ---
    st.header("📈 Cumplimiento Real por Cliente")

    # Función auxiliar para calcular entregas en tiempo
    def calcular_entregas_en_tiempo(grupo):
        return grupo[grupo['Cumplimiento'].isin(["Entregada - En Tiempo", "Entregada - En Tiempo (PD: Pago Pendiente)"])].shape[0]

    # Función auxiliar para calcular visitas en tiempo
    def calcular_visitas_en_tiempo(grupo):
        return grupo[grupo['Cumplimiento'].str.contains("Visita en Tiempo", na=False)].shape[0]

    # Agrupar y aplicar funciones
    df_cliente = df[df['Cumplimiento'] != "Cancelada"].groupby('Cliente').agg(
        Total_Pedidos=('Guia', 'count'),
        Entregas_En_Tiempo=('Cumplimiento', lambda x: calcular_entregas_en_tiempo(x.to_frame().assign(Cumplimiento=x))),
        Visitas_En_Tiempo=('Cumplimiento', lambda x: calcular_visitas_en_tiempo(x.to_frame().assign(Cumplimiento=x)))
    ).reset_index()

    # --- CORRECCIÓN: Asegurar que las columnas son numéricas ---
    df_cliente['Total_Pedidos'] = pd.to_numeric(df_cliente['Total_Pedidos'], errors='coerce').fillna(0)
    df_cliente['Entregas_En_Tiempo'] = pd.to_numeric(df_cliente['Entregas_En_Tiempo'], errors='coerce').fillna(0)
    df_cliente['Visitas_En_Tiempo'] = pd.to_numeric(df_cliente['Visitas_En_Tiempo'], errors='coerce').fillna(0)

    # Ahora sí, calcular el cumplimiento real
    df_cliente['Cumplimiento_Real'] = ((df_cliente['Entregas_En_Tiempo'] + df_cliente['Visitas_En_Tiempo']) / df_cliente['Total_Pedidos'].replace(0, 1) * 100).round(2)

    # Evitar valores infinitos o NaN
    df_cliente['Cumplimiento_Real'] = df_cliente['Cumplimiento_Real'].replace([np.inf, -np.inf], 0).fillna(0)

    # Filtrar clientes con al menos 5 pedidos para evitar ruido
    df_cliente = df_cliente[df_cliente['Total_Pedidos'] >= 5]

    fig_cliente = px.bar(
        df_cliente.sort_values('Cumplimiento_Real', ascending=True),
        x='Cumplimiento_Real',
        y='Cliente',
        orientation='h',
        text='Cumplimiento_Real',
        title='Cumplimiento Real por Cliente (Mín. 5 pedidos)',
        color='Cumplimiento_Real',
        color_continuous_scale='RdYlGn' # Rojo a Verde
    )
    fig_cliente.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig_cliente.update_layout(yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig_cliente, use_container_width=True)

    # --- GRÁFICO DE CUMPLIMIENTO REAL POR ZONA ---
    st.header("🗺️ Cumplimiento Real por Zona (AMBA vs INTERIOR)")

    # Reutilizamos las mismas funciones auxiliares
    df_zona = df[df['Cumplimiento'] != "Cancelada"].groupby('ZONA').agg(
        Total_Pedidos=('Guia', 'count'),
        Entregas_En_Tiempo=('Cumplimiento', lambda x: calcular_entregas_en_tiempo(x.to_frame().assign(Cumplimiento=x))),
        Visitas_En_Tiempo=('Cumplimiento', lambda x: calcular_visitas_en_tiempo(x.to_frame().assign(Cumplimiento=x)))
    ).reset_index()

    # --- CORRECCIÓN: Asegurar que las columnas son numéricas ---
    df_zona['Total_Pedidos'] = pd.to_numeric(df_zona['Total_Pedidos'], errors='coerce').fillna(0)
    df_zona['Entregas_En_Tiempo'] = pd.to_numeric(df_zona['Entregas_En_Tiempo'], errors='coerce').fillna(0)
    df_zona['Visitas_En_Tiempo'] = pd.to_numeric(df_zona['Visitas_En_Tiempo'], errors='coerce').fillna(0)

    df_zona['Cumplimiento_Real'] = ((df_zona['Entregas_En_Tiempo'] + df_zona['Visitas_En_Tiempo']) / df_zona['Total_Pedidos'].replace(0, 1) * 100).round(2)
    df_zona['Cumplimiento_Real'] = df_zona['Cumplimiento_Real'].replace([np.inf, -np.inf], 0).fillna(0)

    fig_zona = px.bar(
        df_zona,
        x='ZONA',
        y='Cumplimiento_Real',
        text='Cumplimiento_Real',
        title='Comparativa de Cumplimiento Real por Zona',
        color='ZONA',
        color_discrete_map={'AMBA': '#28a745', 'INTERIOR': '#007bff'}
    )
    fig_zona.update_traces(texttemplate='%{y:.1f}%', textposition='outside')
    st.plotly_chart(fig_zona, use_container_width=True)

    # --- NUEVO GRÁFICO: TOP 5 AGENCIAS CON MÁS CANCELACIONES ---
    if canceladas > 0:
        st.header("📉 Top 5 Agencias Origen con Más Cancelaciones")
        top_agencias_cancel = df[df['Cumplimiento'] == "Cancelada"]['Agencia origen'].value_counts().head(5)
        if not top_agencias_cancel.empty:
            df_top_ag = top_agencias_cancel.reset_index()
            df_top_ag.columns = ['Agencia Origen', 'Cantidad']

        fig_cancel = px.bar(
            top_agencias_cancel,
            x=top_agencias_cancel.values,
            y=top_agencias_cancel.index,
            orientation='h',
            text=top_agencias_cancel.values,
            title="Top 5 Agencias Origen con Más Cancelaciones",
            color=top_agencias_cancel.values,
            color_continuous_scale='Blues'
        )
        fig_cancel.update_traces(texttemplate='%{text}', textposition='outside')
        fig_cancel.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_cancel, use_container_width=True)
    else:
        st.info("✅ No hay cancelaciones para mostrar.")   

    # --- NUEVO GRÁFICO: TOP 5 LOCALIDADES CON MÁS PEDIDOS FUERA DE TIEMPO ---
    st.header("⏳ Top 5 Localidades (Loc) con Más Pedidos Fuera de Tiempo")
    top_loc_fuera_tiempo = df[df['Cumplimiento'] == "Pendiente - Fuera de Tiempo"]['Loc'].value_counts().head(5)

    if not top_loc_fuera_tiempo.empty:
        # Convertir la Serie en un DataFrame para evitar errores
        df_top_loc = top_loc_fuera_tiempo.reset_index()
        df_top_loc.columns = ['Localidad', 'Cantidad']

        fig_loc = px.bar(
            top_loc_fuera_tiempo,
            x=top_loc_fuera_tiempo.values,
            y=top_loc_fuera_tiempo.index,
            orientation='h',
            text=top_loc_fuera_tiempo.values,
            title="Top 5 Localidades con Más Pedidos Fuera de Tiempo",
            color=top_loc_fuera_tiempo.values,
            color_continuous_scale='Reds'
        )
        fig_loc.update_traces(texttemplate='%{text}', textposition='outside')
        fig_loc.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_loc, use_container_width=True)
    else:
        st.info("✅ No hay pedidos 'Fuera de Tiempo' para mostrar.")

    # --- NUEVA SECCIÓN: ALERTAS DE ESTADO "CREADA" DEMORADO ---
    alertas_creada_demorada = df[df['Alerta Creada Demorada'] != ""]
    if not alertas_creada_demorada.empty:
        st.header("🚨 Alertas de Estado 'Creada' Demorado")
        st.write("Los siguientes pedidos están en estado 'Creada' por más de 24 horas:")
        
        columnas_alerta = [
            'Guia','Importe total', 'Cliente', 'Destinatario', 'Loc', 'ZONA', 
            'Fecha', 'Fecha último estado', 'Estado', 
            'Alerta Creada Demorada', 'Prioridad Alerta'
        ]
        
        # Filtrar columnas existentes
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_creada_demorada[columnas_existentes]
        
        # Mostrar tabla
        st.dataframe(df_alerta)
        
        # Botón de descarga
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Creada Demorada")
        st.download_button(
            label="📥 Descargar Alertas de Estado 'Creada' Demorado (Excel)",
            data=excel_data,
            file_name="Alertas_Creada_Demorada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- NUEVAS ALERTAS MEJORADAS EN LA INTERFAZ ---
    # Alertas de seguimiento de visitas (2+ visitas)
    alertas_seguimiento = df[df['Alerta Seguimiento Visitas'] != ""]
    if not alertas_seguimiento.empty:
        st.header("🔄 Alertas de Seguimiento de Visitas (2+ Visitas)")
        st.write("Pedidos con múltiples visitas que requieren acción:")
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Estado', 'Visitas', 
                          'Fecha último estado', 'Cumplimiento', 'Alerta Seguimiento Visitas', 'Prioridad Alerta']
        # Filtrar columnas existentes
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_seguimiento[columnas_existentes]
        st.dataframe(df_alerta)
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Seguimiento Visitas")
        st.download_button(
            label="📥 Descargar Alertas de Seguimiento de Visitas (Excel)",
            data=excel_data,
            file_name="Alertas_Seguimiento_Visitas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- NUEVA ALERTA: UNA SOLA VISITA SIN SEGUIMIENTO ---
    alertas_una_visita = df[df['Alerta Una Visita Sin Seguimiento'] != ""]
    if not alertas_una_visita.empty:
        st.header("⏰ Alertas de Una Visita Sin Seguimiento")
        st.write("Pedidos con solo una visita que no han tenido seguimiento en 5+ días hábiles:")
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Estado', 'Visitas', 
                          'Fecha último estado', 'Cumplimiento', 'Alerta Una Visita Sin Seguimiento', 'Prioridad Alerta']
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_una_visita[columnas_existentes]
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
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Fecha último estado', 'Estado'
                           , 'Alerta Devolución', 'Prioridad Alerta']
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_devolucion[columnas_existentes]
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
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Fecha último estado', 'Estado', 'Alerta Redespacho', 'Prioridad Alerta']
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_redespacho[columnas_existentes]
        st.dataframe(df_alerta)
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Redespacho")
        st.download_button(
            label="📥 Descargar Alertas Redespacho (Excel)",
            data=excel_data,
            file_name="Alertas_Redespacho.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- NUEVA ALERTA: REPROGRAMADA SIN VISITAS ---
    alertas_reprogramada_sin_visitas = df[df['Alerta Reprogramada Sin Visitas'] != ""]
    if not alertas_reprogramada_sin_visitas.empty:
        st.header("🚨 Alertas de Reprogramada Sin Visitas")
        st.write("Pedidos en estado 'Reprogramada' que no tienen visitas registradas:")
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 
        'Estado', 'Visitas', 'Fecha último estado', 'Cumplimiento', 
        'Alerta Reprogramada Sin Visitas', 'Prioridad Alerta']
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_reprogramada_sin_visitas[columnas_existentes]
        st.dataframe(df_alerta)
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Reprogramada Sin Visitas")
        st.download_button(
            label="📥 Descargar Alertas Reprogramada Sin Visitas (Excel)",
            data=excel_data,
            file_name="Alertas_Reprogramada_Sin_Visitas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )    

    # Alertas de En Tránsito Demorado
    alertas_en_transito = df[df['Alerta En Tránsito Demorado'] != ""]
    if not alertas_en_transito.empty:
        st.header("🚨 Alertas de En Tránsito Demorado")
        st.write("Los siguientes pedidos están en estado 'En tránsito' por más de 48 horas hábiles:")
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Fecha último estado', 'Estado', 'Alerta En Tránsito Demorado', 'Prioridad Alerta']
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_en_transito[columnas_existentes]
        st.dataframe(df_alerta)
        excel_data = generar_excel_desde_df(df_alerta, "Alertas En Tránsito Demorado")
        st.download_button(
            label="📥 Descargar Alertas En Tránsito Demorado (Excel)",
            data=excel_data,
            file_name="Alertas_En_Transito_Demorado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )     

    # Alertas de pendiente fuera de tiempo
    alertas_pendiente_fuera_tiempo = df[df['Alerta Pendiente Fuera Tiempo'] == "Fuera de tiempo crítico"]
    if not alertas_pendiente_fuera_tiempo.empty:
        st.header("🚨 Alertas de Pendiente Fuera de Tiempo")
        st.write("Los siguientes pedidos están pendientes y fuera del tiempo de entrega prometido:")
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Fecha último estado', 'Estado', 'Días Prometidos', 'Lead Time', 'Alerta Pendiente Fuera Tiempo', 'Prioridad Alerta']
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_pendiente_fuera_tiempo[columnas_existentes]
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
        columnas_alerta = ['Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Fecha último estado', 'Estado', 'Condición de venta', 'Alerta Pago Pendiente', 'Prioridad Alerta']
        columnas_existentes = [col for col in columnas_alerta if col in df.columns]
        df_alerta = alertas_pago_pendiente[columnas_existentes]
        st.dataframe(df_alerta)
        excel_data = generar_excel_desde_df(df_alerta, "Alertas Pago Pendiente")
        st.download_button(
            label="📥 Descargar Alertas Pago Pendiente (Excel)",
            data=excel_data,
            file_name="Alertas_Pago_Pendiente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- NUEVA SECCIÓN: ALERTAS DE VENCIMIENTO MAÑANA Y YA VENCIDOS ---
    alertas_vencimiento_mañana = df[df['Alerta Vencimiento Mañana'].isin(["Vence mañana", "Ya vencido"])]
    if not alertas_vencimiento_mañana.empty:
        st.header("🚨 Alertas de Vencimiento")
        st.write("Pedidos que **vence mañana** o que **ya están vencidos**:")
        
        # Mostrar estadísticas rápidas
        vence_mañana = len(alertas_vencimiento_mañana[alertas_vencimiento_mañana['Alerta Vencimiento Mañana'] == "Vence mañana"])
        ya_vencido = len(alertas_vencimiento_mañana[alertas_vencimiento_mañana['Alerta Vencimiento Mañana'] == "Ya vencido"])
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("📅 Vencen Mañana", vence_mañana)
        with col2:
            st.metric("⏰ Ya Vencidos", ya_vencido)
        
        columnas_alerta = [
            'Guia', 'Importe total', 'Cliente', 'Subcuenta', 'Destinatario', 'Tel Destinatario',
            'Loc', 'ZONA', 'Fecha', 'Fecha último estado', 'Estado', 
            'Días Prometidos', 'Lead Time', 'Cumplimiento', 'Días Restantes',
            'Alerta Vencimiento Mañana', 'Prioridad Alerta'
        ]
        columnas_existentes = [col for col in columnas_alerta if col in alertas_vencimiento_mañana.columns]
        df_alerta = alertas_vencimiento_mañana[columnas_existentes]
        st.dataframe(df_alerta)

        excel_data = generar_excel_desde_df(df_alerta, "Alertas Vencimiento")
        st.download_button(
            label="📥 Descargar Alertas de Vencimiento (Excel)",
            data=excel_data,
            file_name="Alertas_Vencimiento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- DESCARGA COMBINADA DE TODAS LAS ALERTAS (ACTUALIZADA) ---
    st.header("📥 Descarga Combinada de Todas las Alertas")
    todas_alertas = df[
        (df['Alerta Seguimiento Visitas'] != "") |
        (df['Alerta Una Visita Sin Seguimiento'] != "") |
        (df['Alerta Devolución'] == "Sugerir devolución") |
        (df['Alerta Redespacho'] == "Redespacho demorado") |
        (df['Alerta Pendiente Fuera Tiempo'] == "Fuera de tiempo crítico") |
        (df['Alerta Pago Pendiente'] == "Pago pendiente demorado") |
        (df['Alerta En Tránsito Demorado'] != "") |
        (df['Alerta Creada Demorada'] != "") |
        (df['Alerta Vencimiento Mañana'].isin(["Vence mañana", "Ya vencido"]))
    ]
    if not todas_alertas.empty:
        columnas_todas = [
            'Guia','Importe total', 'Cliente', 'Destinatario', 'Tel Destinatario', 'Loc', 'ZONA', 'Visitas', 'Fecha último estado', 'Días Prometidos', 'Lead Time', 'Estado', 'Cumplimiento', 'Prioridad Alerta',
            'Alerta Seguimiento Visitas', 'Alerta Una Visita Sin Seguimiento', 
            'Alerta Devolución', 'Alerta Redespacho', 
            'Alerta Pendiente Fuera Tiempo', 'Alerta Pago Pendiente',
            'Alerta Vencimiento Mañana',
            'Alerta En Tránsito Demorado', 'Alerta Creada Demorada'
        ]
        
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

    # --- ACTUALIZAR LA VISTA PREVIA ---
    st.header("🔍 Vista Previa de Datos con Alertas de Variación")

    columnas_mostrar = [
        'Cliente', 'Subcuenta', 'Agencia origen', 'Agencia destino', 'Condición de venta',
        'Fecha', 'Semana Calendario', 'Porcentaje Cumplimiento Semana', 
        'Alerta Variación Semana', 'Variación vs Semana Anterior',  # Nuevas columnas
        'Fecha último estado', 'Estado', 'Visitas', 'ED', 'Loc', 'ZONA', 'Producto',
        'Lead Time', 'Días Prometidos',
        'Cumplimiento', 'Días Restantes', 'Prioridad Alerta',
        'Alerta Seguimiento Visitas', 'Alerta Una Visita Sin Seguimiento',
        'Alerta Devolución', 'Alerta Redespacho', 'Alerta Pendiente Fuera Tiempo', 
        'Alerta Pago Pendiente', 'Alerta En Tránsito Demorado', 'Alerta Creada Demorada',
        'Alerta Vencimiento Mañana'  # Nueva columna
    ]

    columnas_existentes = [col for col in columnas_mostrar if col in df.columns]
    df_vista_previa = df[columnas_existentes].head(10)
    st.dataframe(df_vista_previa)

    # --- DESCARGAS GENERALES ACTUALIZADAS ---
    st.header("📥 Descargas Generales Actualizadas")
    
    # Preparar datos para el Excel de estadísticas
    alertas_vencimiento_count = len(df[df['Alerta Vencimiento Mañana'].isin(["Vence mañana", "Ya vencido"])])
    vence_mañana_count = len(df[df['Alerta Vencimiento Mañana'] == "Vence mañana"])
    ya_vencido_count = len(df[df['Alerta Vencimiento Mañana'] == "Ya vencido"])

    stats_data = {
        "Métrica": [
            "Total Pedidos", "Entregados", "Devueltos", "Canceladas", "Pendientes Reales",
            "Entregada - En Tiempo", "Entregada - En Tiempo (PD)",
            "Entregada - Fuera de Tiempo", "Entregada - Fuera de Tiempo (PD)",
            "Devuelto", "Cancelada",
            "Pendiente - Visita en Tiempo", "Pendiente - Visita Fuera de Tiempo",
            "Pendiente - En Tiempo", "Pendiente - Último Día",
            "Pendiente - Fuera de Tiempo",
            "SLA Principal (%)", "Cumplimiento Entregas (%)", "Cumplimiento Gestión (%)",
            "FADR (%)", "Pedidos por Visita", "Tasa Rechazo/Ausencia (%)",
            "Alertas Creada Demoradas", "Alertas Creada Próximas a Vencer",
            "Alertas Vencimiento Total", "Alertas Vencen Mañana", "Alertas Ya Vencidos"
        ],
        "Valor": [
            total_pedidos, entregados, devueltos, canceladas, pendientes_reales,
            en_tiempo, en_tiempo_pd,
            fuera_tiempo, fuera_tiempo_pd,
            devuelto_count, canceladas,
            visita_en_tiempo, visita_fuera_tiempo,
            pendiente_en_tiempo, pendiente_ultimo_dia,
            pendiente_fuera_tiempo,
            f"{sla_principal:.2f}%", f"{cumplimiento_tradicional:.2f}%", f"{cumplimiento_gestion:.2f}%",
            f"{fadr:.2f}%", f"{pedidos_por_visita:.2f}", f"{tasa_rechazo_ausencia:.2f}%",
            alertas_creada_criticas, alertas_creada_preventivas,
            alertas_vencimiento_count, vence_mañana_count, ya_vencido_count
        ]
    }
    
    if len(stats_data["Métrica"]) == len(stats_data["Valor"]):
        stats_df = pd.DataFrame(stats_data)
    else:
        st.error("❌ Error: Las listas de estadísticas tienen longitudes diferentes")
        stats_df = pd.DataFrame({"Métrica": ["Error en estadísticas"], "Valor": ["Contactar al administrador"]})

    # Guardar en Excel actualizado
    output_excel_actualizado = io.BytesIO()
    with pd.ExcelWriter(output_excel_actualizado, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Base Completa", index=False)
        df_semana.to_excel(writer, sheet_name="Cumplimiento por Semana", index=False)
        stats_df.to_excel(writer, sheet_name="Estadísticas", index=False)
    output_excel_actualizado.seek(0)

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        st.download_button(
            label="📥 Descargar Excel Actualizado (Completo)",
            data=output_excel_actualizado,
            file_name="Reporte_LeadTime_Con_Porcentaje_Semana.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # También actualizar la descarga de la vista previa
    excel_vista_actualizada = generar_excel_desde_df(df[columnas_existentes], "Vista Previa Completa")
    with col_btn2:
        st.download_button(
            label="📥 Descargar Vista Previa Actualizada",
            data=excel_vista_actualizada,
            file_name="Vista_Previa_Actualizada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- GENERAR POWERPOINT (ACTUALIZADO) ---
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
        
        # Calcular métricas de vencimiento para el PowerPoint
        alertas_vencimiento_count = len(df[df['Alerta Vencimiento Mañana'].isin(["Vence mañana", "Ya vencido"])])
        vence_mañana_count = len(df[df['Alerta Vencimiento Mañana'] == "Vence mañana"])
        ya_vencido_count = len(df[df['Alerta Vencimiento Mañana'] == "Ya vencido"])
        
        metrics = [
            f"• Total de pedidos (Excl. Canceladas): {total_pedidos}",
            f"• Entregados: {entregados} ({(entregados/total_pedidos*100):.1f}%)",
            f"• Devueltos: {devueltos} ({(devueltos/total_pedidos*100):.1f}%)",
            f"• Canceladas: {canceladas}",
            f"• SLA Principal: {sla_principal:.1f}%",
            f"• Cumplimiento Entregas: {cumplimiento_tradicional:.1f}%",
            f"• Cumplimiento Gestión: {cumplimiento_gestion:.1f}%",
            f"• FADR (1er Intento): {fadr:.1f}%",
            f"• Tasa Rechazo/Ausencia: {tasa_rechazo_ausencia:.1f}%",
            f"• Visitas en Tiempo: {visita_en_tiempo}",
            f"• Alertas Activas: {len(todas_alertas)}",
            f"• Alertas Creada Demoradas: {alertas_creada_criticas}",
            f"• Alertas Vencimiento: {alertas_vencimiento_count}",
            f"  - Vencen mañana: {vence_mañana_count}",
            f"  - Ya vencidos: {ya_vencido_count}"
        ]
        for metric in metrics:
            p = tf.add_paragraph()
            p.text = metric
            p.font.size = Pt(16)

        # Slide 3: Cumplimiento por Semana
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Cumplimiento por Semana"
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = "Evolución Semanal:"
        p.font.bold = True
        p.font.size = Pt(18)
        
        # Agregar las últimas 4 semanas
        if len(df_semana) > 0:
            ultimas_semanas = df_semana.tail(4)
            for _, semana in ultimas_semanas.iterrows():
                p = tf.add_paragraph()
                p.text = f"Semana {semana['Semana Calendario']}: {semana['Porcentaje Cumplimiento']:.1f}% - {semana['Alerta Variación']}"
                p.font.size = Pt(14)

        pptx_buffer = io.BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        return pptx_buffer

    if st.button("📊 Generar y Descargar PowerPoint Actualizado"):
        pptx_data = crear_pptx()
        st.download_button(
            label="⬇️ Descargar Presentación PPTX Actualizada",
            data=pptx_data,
            file_name="Reporte_LeadTime_Presentacion_Actualizada.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

else:
    st.info("👆 Por favor, sube un archivo Excel para comenzar.")
    st.markdown("""
    **Instrucciones:**
    1. Haz clic en "Browse files".
    2. Selecciona tu archivo Excel.
    3. ¡Listo! La app calculará automáticamente y mostrará gráficos y botones de descarga.
    """)