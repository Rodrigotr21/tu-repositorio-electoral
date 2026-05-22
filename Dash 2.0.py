import streamlit as st
import pandas as pd
import os
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración de la página
st.set_page_config(
    page_title="Gestión de Formatos Electorales | GOECOR/JEL",
    page_icon="🗳️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos personalizados para mejorar la apariencia (Colores Institucionales)
st.markdown("""
    <style>
    .main-title {
        color: #1F4E79;
        font-size: 32px;
        font-weight: bold;
        margin-bottom: 5px;
    }
    .subtitle {
        color: #595959;
        font-size: 16px;
        margin-bottom: 25px;
    }
    .card {
        background-color: #F8FAFC;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1F4E79;
        margin-bottom: 20px;
    }
    .signature-box {
        background-color: #F1F5F9;
        padding: 15px;
        border-radius: 5px;
        border: 1px dashed #94A3B8;
        font-weight: bold;
        color: #334155;
    }
    </style>
""", unsafe_allow_html=True)

# Base de datos estructurada con el contenido real y exacto de los archivos de la GOECOR
FORMATOS = {
    "FM02": {
        "codigo": "FM02-GOECOR/JEL",
        "nombre": "Lista de Chequeo de Acondicionamiento y Señalización del Local de Votación",
        "proposito": "Verificar el estado del acondicionamiento físico, accesos y la señalización del local de votación, asegurando el estricto cumplimiento de las normativas vigentes estipuladas por la GOECOR.",
        "firmas": "Coordinador de Local de Votación (CLV)",
        "tipo_descarga": "Excel (Checklist)",
        "columnas": ["N°", "Aspecto a Evaluar / Elemento de Señalización", "Cumple (Sí)", "No Cumple (No)", "No Aplica (N/A)", "Observaciones / Detalles"],
        "datos_ejemplo": [
            ["1", "Señalización externa e interna de accesos, ingreso y salida clara", "", "", "", ""],
            ["2", "Rampas y vías de acceso libres de obstáculos para personas con discapacidad", "", "", "", ""],
            ["3", "Acondicionamiento de aulas y ubicación de mesas según croquis de la directiva", "", "", "", ""],
            ["4", "Publicación visible de la lista de electores en el exterior del local", "", "", "", ""],
            ["5", "Iluminación adecuada y operativa en pasadizos, escaleras y aulas de sufragio", "", "", "", ""]
        ]
    },
    "FM03": {
        "codigo": "FM03-GOECOR/JEL",
        "nombre": "Control de Distribución del Material Electoral durante la Jornada Electoral",
        "proposito": "Controlar el reparto detallado del material electoral y refrigerios desde el centro de acopio local hacia las mesas de sufragio, registrando la hora exacta y el personal (Titular/Suplente/Fila) con el que se instaló cada mesa.",
        "firmas": "No requiere firma (Documento operativo de control interno)",
        "tipo_descarga": "Excel (Control Operativo)",
        "columnas": [
            "N°", "Nombre del CM o CM STAE", "N° Mesa de Sufragio", "Hora Distribución (HH:MM)", 
            "Caja Material (*)", "Caja Refrigerio (*)", "Presidente (Titular)", "Secretario (Titular)", 
            "3er Miembro (Titular)", "1er Suplente", "2do Suplente", "Electores de la Fila (*)"
        ],
        "datos_ejemplo": [
            ["1", "Pérez Gómez Juan Carlos", "012345", "06:15", "✓", "✓", "✓", "✓", "✓", "", "", ""],
            ["2", "Rojas Mendoza María Fe", "012346", "06:22", "✓", "✓", "✓", "✓", "", "✓", "", ""],
            ["3", "Torres Castro Carlos Alberto", "012347", "06:30", "✓", "✓", "", "", "", "", "", "✓"]
        ]
    },
    "FM05": {
        "codigo": "FM05-GOECOR/JEL",
        "nombre": "Reporte de la Jornada Electoral (SIDE en Físico)",
        "proposito": "Registrar de forma física las variables críticas de la jornada electoral para su transmisión al Centro de Soporte (Instalación de mesas, fuerzas del orden, conformación y motivos de no instalación).",
        "firmas": "Sin firmas (Reporte informativo técnico para transmisión)",
        "tipo_descarga": "Excel (Reporte Consolidado)",
        "columnas": ["Sección / Numeral", "Variable Electoral / Descripción del Reporte", "Valor / Registro Requerido", "Instrucciones de Llenado"],
        "datos_ejemplo": [
            ["1", "CANTIDAD DE MESAS INSTALADAS", "______", "Indicar el total de mesas operativas en el local."],
            ["2", "NÚMERO Y HORA DE INSTALACIÓN DE 1ra MESA DE SUFRAGIO", "Mesa: _________ / Hora: ____:____", "Reportar inmediatamente al Centro de Soporte."],
            ["3", "PRESENCIA DE EFECTIVOS DE LA PNP Y/O FFAA EN EL LOCAL", "PNP: [  ]   FFAA: [  ]   AMBOS: [  ]", "Indicar la cantidad total de efectivos desplegados."],
            ["4", "MESAS INSTALADAS PARA PERSONAS CON DISCAPACIDAD", "______", "Cantidad de mesas especiales habilitadas."],
            ["5", "CONFORMACIÓN DE MESAS (CANTIDAD DE PERSONAS)", "Titulares: ____ Suplentes: ____ Fila: ____", "Especificar el origen de los miembros instalados."],
            ["6", "MESAS NO INSTALADAS POR AUSENCIA DE MIEMBROS DE MESA", "______", "Cantidad de mesas cerradas por inasistencia."],
            ["6", "MESAS NO INSTALADAS POR FALTA DE MATERIAL CRÍTICO", "______", "Registrar obligatoriamente el número de mesa afectado."],
            ["7", "PRESENCIA DE ACTORES ELECTORALES EN EL LOCAL DE VOTACIÓN", "ONPE: [ ] JEE: [ ] FISCALÍA: [ ] PERSONEROS: [ ]", "Marcar los actores presentes durante la jornada."]
        ]
    },
    "FM07": {
        "codigo": "FM07-GOECOR/JEL",
        "nombre": "Control de Documentos y Materiales Electorales en el Centro de Acopio",
        "proposito": "Gestionar y validar rigurosamente la recepción y repliegue de las actas y sobres electorales de seguridad devueltos por las mesas de sufragio al Centro de Acopio al cierre de la votación.",
        "firmas": "Coordinador de Local de Votación (CLV)",
        "tipo_descarga": "Excel (Validación de Actas)",
        "columnas": [
            "N°", "N° de Mesa de Sufragio", "Sobre Plomo (√)", "Sobre Verde (√)", 
            "Sobre Celeste (√)", "Sobre Rojo (√)", "Caja Restos Electorales (√)", 
            "Sobre Anaranjado (√)", "Sobre Cédulas No Impugnadas (√)"
        ],
        "datos_ejemplo": [
            ["1", "012345", "√", "√", "√", "√", "√", "√", "√"],
            ["2", "012346", "√", "√", "√", "√", "√", "√", "√"],
            ["3", "012347", "√", "√", "√", "√", "√", "√", "√"]
        ]
    },
    "FM11": {
        "codigo": "FM11-GOECOR/JEL",
        "nombre": "Acta de Recepción y Devolución del Local de Votación",
        "proposito": "Documentar formalmente el estado de la infraestructura civil y aulas cedidas por la institución educativa antes (Recepción) y después (Devolución) de la jornada electoral.",
        "firmas": "Coordinador de Local de Votación y Responsable Designado de la Institución Propietaria",
        "tipo_descarga": "Word (Acta Formal / Acta de Compromiso)",
        "columnas": ["Condiciones de Recepción/Devolución de Ambientes", "Especificación de Aula(s) / Ubicación", "Estado en Recepción (PRE)", "Estado en Devolución (POST)"],
        "datos_ejemplo": [
            ["Ambientes Cedidos (Aulas / Patios)", "Aulas del 101 al 115, Patio Central", "Recibido Conforme", "Devuelto Conforme"],
            ["Ventanas rotas / Vidrios dañados", "Aula 104", "01 vidrio rajado al ingreso", "Permanece en igual condición"],
            ["Puertas rotas / Cerraduras defectuosas", "Ninguna", "Conforme", "Devuelto Conforme"],
            ["Baños malogrados / Sanitarios inoperativos", "SS.HH. de Varones - Pabellón B", "01 fluxómetro atorado", "Reparado / Operativo"],
            ["Otros (Pizarras, Mobiliario escolar, etc.)", "Aulas de primer piso", "Mobiliario completo", "Devuelto sin novedades"]
        ]
    },
    "FM12": {
        "codigo": "FM12-GOECOR/JEL",
        "nombre": "Hoja de Ruta para el Traslado de Sobres Plomos por Entregas a la Sede ODPE",
        "proposito": "Garantizar la trazabilidad absoluta en el traslado físico de las actas electorales (sobres plomos) desde el local de votación hacia la sede principal de la ODPE por cada lote de entrega.",
        "firmas": "Responsable del Traslado de Sobres, Coordinador de Local de Votación y Encargado de Recepción ODPE",
        "tipo_descarga": "Excel / Word (Hoja de Ruta)",
        "columnas": ["N°", "N° de la Mesa de Sufragio", "Elección / Consulta 1 (✓)", "Elección / Consulta 2 (✓)", "Elección / Consulta 3 (✓)", "Firma del Responsable del Traslado"],
        "datos_ejemplo": [
            ["1", "012345", "✓", "", "", "___________________________"],
            ["2", "012346", "✓", "✓", "", "___________________________"],
            ["3", "012347", "✓", "✓", "✓", "___________________________"]
        ]
    },
    "FM13": {
        "codigo": "FM13-GOECOR/JEL",
        "nombre": "Hoja de Ruta para el Traslado de Documentos Electorales desde el Local hacia la ODPE",
        "proposito": "Monitorear el repliegue seguro y completo de los documentos electorales restantes y sobres con cédulas de sufragio utilizadas hacia la ODPE al término del escrutinio.",
        "firmas": "Responsable del Traslado de Sobres, Coordinador de Local de Votación (CLV/RLV) y Encargado de Recepción en la ODPE",
        "tipo_descarga": "Excel (Trazabilidad)",
        "columnas": ["N°", "N° Mesa de Sufragio", "Sobre con Cédulas de Sufragio (✓)", "Elección / Consulta 1 (✓)", "Elección / Consulta 2 (✓)", "Elección / Consulta 3 (✓)"],
        "datos_ejemplo": [
            ["1", "012345", "✓", "✓", "✓", "✓"],
            ["2", "012346", "✓", "✓", "✓", "✓"],
            ["3", "012347", "✓", "✓", "✓", "✓"]
        ]
    },
    "FM15": {
        "codigo": "FM15-GOECOR/JEL",
        "nombre": "Cargo de Entrega del Material de Reserva al CM / CM STAE",
        "proposito": "Registrar de manera transparente cualquier asignación excepcional de material electoral de contingencia o reserva (ánforas, cédulas, tintas) entregado a los coordinadores de mesa.",
        "firmas": "Coordinador de Local de Votación (CLV) y Coordinador de Mesa (CM / CM STAE) receptor",
        "tipo_descarga": "Word (Cargo de Entrega)",
        "columnas": ["Descripción del Material de Reserva", "Cantidad de Material Entregado", "Número de Mesa de Sufragio", "Nombres y Apellidos del CM / CM STAE", "Firma del CM / CM STAE"],
        "datos_ejemplo": [
            ["Tinta Indelebre de contingencia", "1 unidad", "012345", "Pedro Alcántara Gómez", "[Firma Física]"],
            ["Cédulas de Sufragio Adicionales", "10 unidades", "012349", "Lucía Fernández Flores", "[Firma Física]"],
            ["Tampón para huella dactilar", "1 unidad", "012352", "Marcos Ruiz Espinoza", "[Firma Física]"]
        ]
    },
    "FM16": {
        "codigo": "FM16-GOECOR/JEL",
        "nombre": "Cargo de Recepción de Sobres Verdes, Celestes, Rojos y Anaranjados en Custodia",
        "proposito": "Contabilizar y custodiar en forma matricial los sobres de seguridad de colores correspondientes a las diferentes elecciones o consultas frente a la cantidad de ánforas recolectadas.",
        "firmas": "Responsable de Local de Votación (RLV), Responsable del Traslado de Sobres y Encargado de Custodia ODPE",
        "tipo_descarga": "Excel (Matriz de Control)",
        "columnas": ["Tipo de Sobre", "Elección / Consulta 1 (Ánforas)", "Elección / Consulta 2 (Ánforas)", "Elección / Consulta 3 (Ánforas)", "Total General Custodiado"],
        "datos_ejemplo": [
            ["SOBRES ROJOS", "15", "0", "0", "15 Ánforas"],
            ["SOBRES ANARANJADOS", "0", "15", "0", "15 Ánforas"],
            ["SOBRES VERDES", "15", "15", "0", "30 Ánforas"],
            ["SOBRES CELESTES", "0", "0", "15", "15 Ánforas"]
        ]
    }
}

def crear_excel_formato(id_formato, info):
    """Genera un archivo Excel altamente formateado y profesional basándose en los lineamientos visuales"""
    wb = Workbook()
    ws = wb.active
    ws.title = id_formato
    
    # Asegurar líneas de cuadrícula visibles
    ws.views.sheetView[0].showGridLines = True
    
    # Paleta de colores institucionales
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    
    # Fuentes
    font_title = Font(name="Arial", size=14, bold=True, color="1F4E79")
    font_subtitle = Font(name="Arial", size=10, italic=True, color="595959")
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=10)
    font_bold = Font(name="Arial", size=10, bold=True)
    
    # Bordes
    thin_side = Side(border_style="thin", color="D9D9D9")
    thick_bottom = Side(border_style="medium", color="1F4E79")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # Alineación
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    
    # 1. Título e Identificación Oficial del Formato
    ws["A1"] = info["codigo"]
    ws["A1"].font = font_title
    ws.merge_cells("A1:E1")
    
    ws["A2"] = info["nombre"]
    ws["A2"].font = Font(name="Arial", size=12, bold=True, color="333333")
    ws.merge_cells("A2:E2")
    
    ws["A3"] = f"Propósito: {info['proposito']}"
    ws["A3"].font = font_subtitle
    ws["A3"].alignment = align_left
    ws.merge_cells("A3:E3")
    
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 28
    
    start_row = 5
    
    # 2. Encabezados de la Tabla de Datos
    for col_num, header_text in enumerate(info["columnas"], start=1):
        cell = ws.cell(row=start_row, column=col_num, value=header_text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = align_center if any(x in header_text.lower() for x in ["n°", "hora", "✓", "√", "sí", "no"]) else align_left
        cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom)
    ws.row_dimensions[start_row].height = 26
    
    # 3. Datos de la tabla con Zebra Striping
    current_row = start_row + 1
    for i, row_data in enumerate(info["datos_ejemplo"]):
        for col_num, val in enumerate(row_data, start=1):
            cell = ws.cell(row=current_row, column=col_num, value=val)
            cell.font = font_body
            cell.border = border_cell
            
            if col_num == 1 or col_num == 2:
                cell.alignment = align_left
            else:
                cell.alignment = align_center if len(str(val)) < 15 else align_left
                
            if i % 2 == 1:
                cell.fill = zebra_fill
                
        ws.row_dimensions[current_row].height = 20
        current_row += 1
        
    # Agregar filas vacías adicionales para llenado manual
    for _ in range(5):
        for col_num in range(1, len(info["columnas"]) + 1):
            cell = ws.cell(row=current_row, column=col_num)
            cell.border = border_cell
            if (_ + len(info["datos_ejemplo"])) % 2 == 1:
                cell.fill = zebra_fill
        ws.row_dimensions[current_row].height = 20
        current_row += 1
        
    # 4. Sección de Firmas
    current_row += 2
    ws.cell(row=current_row, column=1, value="CONTROL DE VALIDACIÓN Y FIRMAS REGISTRADAS:").font = font_bold
    current_row += 1
    
    firmas_req = info["firmas"]
    ws.cell(row=current_row, column=1, value=f"Responsables Obligatorios: {firmas_req}").font = font_subtitle
    current_row += 2
    
    if "No requiere" not in firmas_req and "Sin firmas" not in firmas_req:
        lista_firmas = [f.strip() for f in firmas_req.replace(" y ", ",").split(",")]
        for idx, f_name in enumerate(lista_firmas):
            col_pos = (idx * 2) + 1
            if col_pos <= len(info["columnas"]):
                ws.cell(row=current_row, column=col_pos, value="___________________________").alignment = align_center
                ws.cell(row=current_row+1, column=col_pos, value=f_name).font = Font(name="Arial", size=9, bold=True)
                ws.cell(row=current_row+1, column=col_pos).alignment = align_center
                ws.cell(row=current_row+2, column=col_pos, value="DNI / Firma Digital / Huella").font = Font(name="Arial", size=8, italic=True)
                ws.cell(row=current_row+2, column=col_pos).alignment = align_center
    
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 2, 3]:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- INTERFAZ COMPLETA DE STREAMLIT ---

st.markdown('<div class="main-title">🗳️ Sistema de Control y Formatos Electorales</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Oficina General de Coordinación Electoral (GOECOR) - JEL.</div>', unsafe_allow_html=True)

st.sidebar.header("Panel de Navegación")
formato_seleccionado = st.sidebar.selectbox(
    "Seleccione el Formato Oficial:",
    options=list(FORMATOS.keys()),
    format_func=lambda x: f"{x} - {FORMATOS[x]['nombre'][:40]}..."
)

info_f = FORMATOS[formato_seleccionado]

col_izq, col_der = st.columns([2, 3])

with col_izq:
    st.markdown("### 📋 Especificaciones Técnicas")
    st.markdown(f"""
    <div class="card" style="color: #000000;">
        <h4><b style="color: #1F4E79;">Código Oficial:</b> {info_f['codigo']}</h4>
        <p><b>Denominación:</b> {info_f['nombre']}</p>
        <p><b>Propósito General:</b> {info_f['proposito']}</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### ✍️ Protocolo de Firmas")
    st.markdown(f'<div class="signature-box">🔒 {info_f["firmas"]}</div>', unsafe_allow_html=True)
    
    st.markdown("### 📥 Descarga de Plantilla Oficial")
    st.write(f"De acuerdo al flujo del manual, se sugiere el uso de formato **{info_f['tipo_descarga']}**.")
    
    # 1. Identificar carpeta
    carpeta_actual = os.path.dirname(os.path.abspath(__file__))
    
    # 2. Deducir el nombre del archivo (CORREGIDO según las extensiones reales de la carpeta)
    formatos_word = ["FM02", "FM05", "FM11", "FM15", "FM16"]
    extension = ".docx" if formato_seleccionado in formatos_word else ".xlsx"
    nombre_archivo_fisico = f"{formato_seleccionado}{extension}"

    # 3. Construir la ruta completa al archivo
    ruta_archivo = os.path.join(carpeta_actual, nombre_archivo_fisico)
    
    # 4. Enlace de descarga de archivo físico u opción B
    if os.path.exists(ruta_archivo):
        with open(ruta_archivo, "rb") as file:
            st.download_button(
                label=f"💾 Descargar {info_f['codigo']} (Archivo Original)",
                data=file,
                file_name=nombre_archivo_fisico,
                mime="application/octet-stream",
                key=f"dl_real_{formato_seleccionado}"
            )
        st.success("✅ Archivo físico enlazado correctamente.")
    else:
        st.warning(f"⚠️ No se encontró el archivo '{nombre_archivo_fisico}' en el servidor.")
        st.info("💡 Ejecutando generador dinámico de respaldo en tiempo real.")
        
        excel_data = crear_excel_formato(formato_seleccionado, info_f)
        st.download_button(
            label=f"⚙️ Generar y Descargar {info_f['codigo']} (.xlsx)",
            data=excel_data,
            file_name=f"{info_f['codigo'].replace('/', '_')}_Generado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_gen_{formato_seleccionado}"
        )

with col_der:
    st.markdown(f"### 🔍 Vista Previa Estructurada de Datos")
    st.write("Estructura de la matriz de datos fiel al documento original:")
    
    # Renderizado directo de la estructura tal cual el archivo real
    df_preview = pd.DataFrame(info_f["datos_ejemplo"], columns=info_f["columnas"])
    st.dataframe(df_preview, width='stretch', hide_index=True)
    
    st.markdown("#### 📊 Métricas Operativas Asociadas al Formato")
    m1, m2 = st.columns(2)
    if formato_seleccionado in ["FM02", "FM11"]:
        m1.metric("Momento de Aplicación", "Pre y Post Jornada")
        m2.metric("Nivel de Criticidad", "Alto (Infraestructura)")
    elif formato_seleccionado in ["FM03", "FM05"]:
        m1.metric("Momento de Aplicación", "Mañana (Instalación)")
        m2.metric("Nivel de Criticidad", "Crítico (SIDE)")
    elif formato_seleccionado in ["FM07", "FM12", "FM13", "FM16"]:
        m1.metric("Momento de Aplicación", "Tarde / Cierre")
        m2.metric("Nivel de Criticidad", "Máximo (Actas)")
    else:
        m1.metric("Momento de Aplicación", "Durante el Día D")
        m2.metric("Nivel de Criticidad", "Medio")

st.markdown("---")
st.markdown("### 📑 Resumen General del Catálogo GOECOR/JEL")
with st.expander("Ver lista resumida de responsabilidades de firmas"):
    tabla_resumen = []
    for k, v in FORMATOS.items():
        tabla_resumen.append({"Código": v["codigo"], "Nombre": v["nombre"], "Quién Firma": v["firmas"]})
    st.table(pd.DataFrame(tabla_resumen))