import streamlit as st
import pandas as pd
from docx import Document
from datetime import datetime
import io

# 1. Configuración de la página
st.set_page_config(page_title="Generador de Oficios", page_icon="🚗")
st.title("🚗 Generador Automático de Oficios de Comisión")
st.write("Llena los siguientes datos para descargar tu oficio de forma inmediata.")

# 2. Cargar la base de datos (con caché para que la web sea muy rápida)
@st.cache_data
def cargar_datos():
    return pd.read_excel('Información del empleado.xlsx')

try:
    df = cargar_datos()
except Exception as e:
    st.error("❌ No se pudo cargar la base de datos. Verifica los archivos.")
    st.stop()

# 3. Interfaz de usuario: Los campos que llenarán los empleados
with st.form("formulario_oficio"):
    num_empleado = st.number_input("1. Número de Empleado:", min_value=1, step=1, format="%d")
    placa_input = st.text_input("2. Placas de la unidad que ocuparás (Ej. HM4036G):").strip().upper()
    lugar_motivo = st.text_area("3. ¿Hacia dónde te diriges y en qué fecha? (Ej. Secretaría de Finanzas, 12 de septiembre):")
    
    # Botón para enviar el formulario
    generar = st.form_submit_button("Generar Oficio")

# 4. Lógica que se ejecuta al presionar el botón
if generar:
    if not num_empleado or not placa_input or not lugar_motivo:
        st.warning("⚠️ Por favor, llena todos los campos antes de generar el oficio.")
    else:
        # Buscar empleado
        empleado_data = df[df['No. empleado'] == num_empleado]
        # Buscar vehículo
        vehiculo_data = df[df['placa'].astype(str).str.upper() == placa_input]
        
        if empleado_data.empty:
            st.error(f"⚠️ No se encontró el número de empleado {num_empleado}.")
        elif vehiculo_data.empty:
            st.error(f"⚠️ No se encontró la placa '{placa_input}'.")
        else:
            # Extraer datos si todo es correcto
            datos_emp = empleado_data.iloc[0]
            datos_veh = vehiculo_data.iloc[0]
            
            # Fecha y hora actual
            ahora = datetime.now()
            fecha_actual = ahora.strftime("%d/%m/%Y") 
            hora_actual = ahora.strftime("%H:%M")
            
            modelo = str(datos_veh['Modelo'])
            if modelo.endswith('.0'):
                modelo = modelo[:-2]
                
            reemplazos = {
                '[Fecha]': fecha_actual,
                '[Nombre]': str(datos_emp['Nombre']),
                '[Adscripción]': str(datos_emp['Adscripción']),
                '[Puesto]': str(datos_emp['Puesto']),
                '[Nombramiento]': str(datos_emp['Nombramiento']),
                '[RFC]': str(datos_emp['RFC']),
                '[Unidad]': str(datos_veh['Unidad']),
                '[Modelo]': modelo,
                '[Placa]': str(datos_veh['placa']), 
                '[Horario en tiempo real]': hora_actual,
                '[Lugar al que asistirá y Fecha en dia y mes]': lugar_motivo
            }
            
            try:
                # Modificar el Word
                doc = Document('OFICIO COMISIÓN 2026_2.docx')
                for parrafo in doc.paragraphs:
                    for etiqueta, valor_real in reemplazos.items():
                        if etiqueta in parrafo.text:
                            parrafo.text = parrafo.text.replace(etiqueta, valor_real)
                
                # Guardar el documento en la memoria de la web
                bio = io.BytesIO()
                doc.save(bio)
                nombre_archivo_salida = f"Oficio_{datos_emp['Nombre'].replace(' ', '_')}_{placa_input}.docx"
                
                st.success(f"✅ ¡Hola {datos_emp['Nombre']}! Tu oficio está listo.")
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar Documento Word",
                    data=bio.getvalue(),
                    file_name=nombre_archivo_salida,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"❌ Ocurrió un error al procesar la plantilla: {e}")
