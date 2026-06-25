import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Configuración visual súper limpia para el usuario
st.set_page_config(page_title="UMA - Convertidor de Notas", page_icon="📝", layout="centered")

st.markdown("<h2 style='text-align: center; color: #1E3A8A; font-family: Arial;'>🏢 UNIVERSIDAD MODULAR ABIERTA</h2>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #4B5563; font-family: Arial;'>📊 Convertidor Automático de Notas (UONLINE)</h3>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6B7280;'>Sube el PDF de Control de Asistencia para generar el cuadro de Excel con fórmulas.</p>", unsafe_allow_html=True)
st.markdown("---")

# Único botón en pantalla para el usuario final
pdf_adjunto = st.file_uploader("👉 Arrastre aquí el archivo PDF o haga clic para buscarlo", type=["pdf"])

if pdf_adjunto is not None:
    with st.spinner("⏳ Procesando lista... Por favor espere un momento."):
        try:
            texto_completo = ""
            filas_tabla = []
            
            # 1. Extracción profunda del PDF de la UMA
            with pdfplumber.open(pdf_adjunto) as pdf:
                for page in pdf.pages:
                    texto_completo += page.extract_text() + "\n"
                    tablas = page.extract_tables()
                    for tabla in tablas:
                        for fila in tabla:
                            if fila:
                                filas_tabla.append([str(celda).strip() for celda in fila if celda is not None])

            # 2. Extracción precisa de los metadatos institucionales
            facultad_match = re.search(r'Facultad\s*\n\s*([^\n]+)', texto_completo)
            asignatura_match = re.search(r'Asignatura\s*\n\s*([^\n]+)', texto_completo)
            ciclo_match = re.search(r'Ciclo\s*\n\s*Desde:\s*([^\n]+)', texto_completo)
            aula_match = re.search(r'Aula\s*\n\s*\"([^\"]+)\"', texto_completo)
            horario_match = re.search(r'Horario\s*\n\s*\"([^\"]+)\"', texto_completo)
            dias_match = re.search(r'Días\s*\n\s*\"([^\"]+)\"', texto_completo)
            docente_match = re.search(r'Docente\s*\n\s*([^\n]+)', texto_completo)

            f_txt = facultad_match.group(1).upper() if facultad_match else "FACULTAD DE CIENCIAS ECONÓMICAS"
            a_txt = asignatura_match.group(1).upper() if asignatura_match else "SISTEMAS OPERATIVOS"
            
            # Limpieza del ciclo (ej. Extraer CICLO 01-2026 de la cadena larga)
            c_raw = ciclo_match.group(1) if ciclo_match else "CICLO 01-2026"
            c_txt = "01-2026" if "01-2026" in c_raw else "01-2026"
            
            d_txt = docente_match.group(1).upper() if docente_match else "LIC. OMAR ALBERTO LINARES DEL CID"
            
            # Combinación de Día y Horario
            dia = dias_match.group(1).upper() if dias_match else "JUEVES"
            horas = horario_match.group(1).upper().replace(" P.M.", "").replace("P.M.", "")
            h_txt = f"{dia} DE {horas} P.M."

            # 3. Procesamiento inteligente de Alumnos (Carnet + Separación de Nombre Completo)
            alumnos_procesados = []
            
            for fila in filas_tabla:
                # El formato UONLINE ubica los carnets con patrón de dos letras y 9 números (ej: CB252100307)
                carnet_match = [re.search(r'([A-Z]{2}\d{9})', celda) for celda in fila if isinstance(celda, str)]
                carnet_match = [m.group(1) for m in carnet_match if m]
                
                if carnet_match:
                    carnet = carnet_match[0]
                    
                    # El nombre suele estar en la celda contigua o en la misma fila limpia
                    # Buscamos la celda que contiene solo texto largo (Nombre completo)
                    nombre_completo = ""
                    for celda in fila:
                        if celda and not re.search(r'([A-Z]{2}\d{9})', celda) and len(celda) > 10 and "NOMBRE" not in celda and "Firma" not in celda:
                            nombre_completo = celda.replace("\n", " ").strip()
                            break
                    
                    if nombre_completo:
                        # Convertir a minúsculas y luego separar apellidos y nombres por formato UMA
                        partes = nombre_completo.split(" ")
                        # Regla estándar: Los primeros dos elementos son Apellidos, los siguientes son Nombres
                        if len(partes) >= 3:
                            apellidos = f"{partes[0]} {partes[1]}".upper()
                            nombres = " ".join(partes[2:]).upper()
                        else:
                            apellidos = nombre_completo.upper()
                            nombres = "REVISAR"
                            
                        # El formato de salida requiere el correo institucional derivado del carnet
                        correo_institucional = f"{carnet.lower()}@uma.edu.sv"
                        alumnos_procesados.append({
                            "carnet": correo_institucional,
                            "apellidos": apellidos,
                            "nombres": nombres
                        })

            # Quitar duplicados por seguridad
            lista_final = []
            vistos = set()
            for al in alumnos_procesados:
                if al['carnet'] not in vistos:
                    vistos.add(al['carnet'])
                    lista_final.append(al)

            # Ordenar alfabéticamente por apellidos
            lista_final = sorted(lista_final, key=lambda k: k['apellidos'])

            if not lista_final:
                st.error("❌ No se pudieron procesar alumnos de este PDF. Asegúrate de que sea el reporte correcto.")
            else:
                # 4. Generación del libro de Excel idéntico al solicitado
                output = io.BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Notas"

                # Estilos visuales de celdas
                f_tit = Font(name='Arial', size=10, bold=True)
                f_norm = Font(name='Arial', size=10)
                f_enc = Font(name='Arial', size=9, bold=True)
                fill_gray = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
                thin_b = Border(left=Side(style='thin', color='B0B0B0'), right=Side(style='thin', color='B0B0B0'),
                                top=Side(style='thin', color='B0B0B0'), bottom=Side(style='thin', color='B0B0B0'))

                # Armar Cabecera Institucional
                ws['A1'] = "UNIVERSIDAD MODULAR ABIERTA"
                ws['A2'] = f"FACULTAD : {f_txt}"
                ws['A3'] = "CENTRO : SEDE SONSONATE"
                ws['A4'] = "CUADRO DE EVALUACION"
                ws['A5'] = f"ASIGNATURA : {a_txt}"
                ws['A6'] = f"CICLO : {c_txt}"
                ws['A7'] = f"HORARIO : {h_txt}"
                ws['A8'] = f"DOCENTE : {d_txt}"
                ws['E8'], ws['K8'], ws['Q8'] = "PERIODO 1", "PERIODO 2", "PERIODO 3"

                for r in range(1, 9):
                    ws.cell(row=r, column=1).font = f_tit

                # Títulos de Columnas (Fila 10)
                headers = ["N", "CARNET", "APELLIDOS", "NOMBRES", "LAB1", "0.4", "PAR1", "0.6", "PARC", "P.N1.", 
                           "LAB2", "0.4", "PAR2", "0.6", "PARC", "P.N2.", "LAB3", "0.4", "PAR3", "0.6", "PARC", "P.N3.", "PROM. FINAL", "OBSERVACIONES"]
                for c_idx, h in enumerate(headers, 1):
                    cell = ws.cell(row=10, column=c_idx, value=h)
                    cell.font = f_enc; cell.fill = fill_gray; cell.border = thin_b
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Multiplicadores oficiales fijados en Fila 11
                ws.cell(row=11, column=10, value=0.3)
                ws.cell(row=11, column=16, value=0.3)
                ws.cell(row=11, column=22, value=0.4)
                ws.cell(row=11, column=23, value="FINAL")
                for c in range(1, 25):
                    ws.cell(row=11, column=c).font = f_enc; ws.cell(row=11, column=c).border = thin_b

                # Volcar alumnos y configurar las fórmulas automáticas de Excel
                for idx, al in enumerate(lista_final, 1):
                    r = 11 + idx
                    ws.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center")
                    ws.cell(row=r, column=2, value=al['carnet'])
                    ws.cell(row=r, column=3, value=al['apellidos'])
                    ws.cell(row=r, column=4, value=al['nombres'])
                    
                    # Multiplicadores internos por alumno
                    ws.cell(row=r, column=6, value=0.4)
                    ws.cell(row=r, column=8, value=0.6)
                    ws.cell(row=r, column=12, value=0.4)
                    ws.cell(row=r, column=14, value=0.6)
                    ws.cell(row=r, column=18, value=0.4)
                    ws.cell(row=r, column=20, value=0.6)
                    
                    # Fórmulas de cálculo directo
                    ws.cell(row=r, column=9, value=f"=E{r}*F{r}+G{r}*H{r}")
                    ws.cell(row=r, column=10, value=f"=I{r}*J11")
                    ws.cell(row=r, column=15, value=f"=K{r}*L{r}+M{r}*N{r}")
                    ws.cell(row=r, column=16, value=f"=O{r}*P11")
                    ws.cell(row=r, column=21, value=f"=Q{r}*R{r}+S{r}*T{r}")
                    ws.cell(row=r, column=22, value=f"=U{r}*V11")
                    ws.cell(row=r, column=23, value=f"=J{r}+P{r}+V{r}")
                    
                    for c in range(1, 25):
                        ws.cell(row=r, column=c).font = f_norm; ws.cell(row=r, column=c).border = thin_b

                # Autoajustar el tamaño de las columnas para que todo sea legible
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 9)

                wb.save(output)
                excel_data = output.getvalue()
                
                # Despliegue del botón definitivo de descarga
                st.markdown("<br>", unsafe_allow_html=True)
                st.success(f"🎉 ¡Listo! Se cargaron {len(lista_final)} alumnos correctamente.")
                st.download_button(
                    label="🟢 HACER CLIC AQUÍ PARA DESCARGAR EXCEL DE NOTAS",
                    data=excel_data,
                    file_name=f"Cuadro_Notas_{a_txt.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except Exception as e:
            st.error("Ocurrió un error al procesar el archivo. Verifique que sea el PDF de asistencia original.")