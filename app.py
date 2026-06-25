import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Configuración visual muy limpia
st.set_page_config(page_title="UMA - Convertidor de Notas", page_icon="📊")

st.write("### 🏢 UNIVERSIDAD MODULAR ABIERTA")
st.title("📊 Convertidor Automático de PDF a Excel")
st.write("Sube el archivo PDF del listado para generar el cuadro de notas oficial en Excel.")

# 1. Botón único para subir el archivo
pdf_adjunto = st.file_uploader("👉 Haz clic aquí para buscar y subir el archivo PDF", type=["pdf"])

if pdf_adjunto is not None:
    with st.spinner("Procesando archivo... Por favor espera un momento."):
        
        # --- Extracción de datos interna ---
        texto_completo = ""
        alumnos = []
        with pdfplumber.open(pdf_adjunto) as pdf:
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"
        
        # Buscar datos de cabecera automáticamente
        facultad = re.search(r'(?i)FACULTAD\s*:\s*([^\n]+)', texto_completo)
        asignatura = re.search(r'(?i)ASIGNATURA\s*:\s*([^\n]+)', texto_completo)
        ciclo = re.search(r'(?i)CICLO\s*:\s*([^\n]+)', texto_completo)
        docente = re.search(r'(?i)DOCENTE\s*:\s*([^\n]+)', texto_completo)
        horario = re.search(r'(?i)HORARIO\s*:\s*([^\n]+)', texto_completo)
        
        f_txt = facultad.group(1).strip().upper() if facultad else "FACULTAD DE CIENCIAS ECONÓMICAS"
        a_txt = asignatura.group(1).strip().upper() if asignatura else "SISTEMAS OPERATIVOS"
        c_txt = ciclo.group(1).strip() if ciclo else "01-2026"
        d_txt = docente.group(1).strip().upper() if docente else "LIC. OMAR ALBERTO LINARES DEL CID"
        h_txt = horario.group(1).strip().upper() if horario else "JUEVES DE 1:00 A 4:40 P.M."

        # Extraer alumnos por correos @uma.edu.sv
        lineas = texto_completo.split('\n')
        for linea in lineas:
            match_correo = re.search(r'([a-zA-Z0-9._%+-]+@uma\.edu\.sv)', linea)
            if match_correo:
                correo = match_correo.group(1)
                texto_sin_correo = linea.replace(correo, "").strip()
                partes = [p.strip() for p in re.split(r',\s*|\s{2,}', texto_sin_correo) if p]
                apellidos = partes[0].upper() if len(partes) > 0 else "REVISAR"
                nombres = partes[1].upper() if len(partes) > 1 else "REVISAR"
                alumnos.append({"carnet": correo, "apellidos": apellidos, "nombres": nombres})

        # Quitar duplicados de alumnos
        alumnos_unicos = []
        vistos = set()
        for a in alumnos:
            if a['carnet'] not in vistos:
                vistos.add(a['carnet'])
                alumnos_unicos.append(a)

        if not alumnos_unicos:
            st.error("❌ No se encontraron alumnos en este PDF. Asegúrate de que sea el archivo correcto.")
        else:
            # --- Creación del Excel con Fórmulas ---
            output = io.BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = "Notas"
            
            # Estilos visuales idénticos a la plantilla
            f_tit = Font(name='Arial', size=10, bold=True)
            f_norm = Font(name='Arial', size=10)
            f_enc = Font(name='Arial', size=9, bold=True)
            fill_gray = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
            thin_b = Border(left=Side(style='thin', color='B0B0B0'), right=Side(style='thin', color='B0B0B0'),
                            top=Side(style='thin', color='B0B0B0'), bottom=Side(style='thin', color='B0B0B0'))

            # Cabeceras del Excel institucional
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

            # Encabezados de columnas (Fila 10 y 11)
            headers = ["N", "CARNET", "APELLIDOS", "NOMBRES", "LAB1", "0.4", "PAR1", "0.6", "PARC", "P.N1.", 
                       "LAB2", "0.4", "PAR2", "0.6", "PARC", "P.N2.", "LAB3", "0.4", "PAR3", "0.6", "PARC", "P.N3.", "PROM. FINAL", "OBSERVACIONES"]
            for c_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=10, column=c_idx, value=h)
                cell.font = f_enc; cell.fill = fill_gray; cell.border = thin_b
                cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.cell(row=11, column=10, value=0.3)
            ws.cell(row=11, column=16, value=0.3)
            ws.cell(row=11, column=22, value=0.4)
            ws.cell(row=11, column=23, value="FINAL")
            for c in range(1, 25):
                ws.cell(row=11, column=c).font = f_enc; ws.cell(row=11, column=c).border = thin_b

            # Agregar alumnos y escribir las fórmulas automáticas
            for idx, al in enumerate(alumnos_unicos, 1):
                r = 11 + idx
                ws.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center")
                ws.cell(row=r, column=2, value=al['carnet'])
                ws.cell(row=r, column=3, value=al['apellidos'])
                ws.cell(row=r, column=4, value=al['nombres'])
                
                # Multiplicadores fijos por celda
                ws.cell(row=r, column=6, value=0.4)
                ws.cell(row=r, column=8, value=0.6)
                ws.cell(row=r, column=12, value=0.4)
                ws.cell(row=r, column=14, value=0.6)
                ws.cell(row=r, column=18, value=0.4)
                ws.cell(row=r, column=20, value=0.6)
                
                # Fórmulas matemáticas automáticas de Excel
                ws.cell(row=r, column=9, value=f"=E{r}*F{r}+G{r}*H{r}")   # Nota parcial 1
                ws.cell(row=r, column=10, value=f"=I{r}*J11")            # Ponderación 30% p1
                ws.cell(row=r, column=15, value=f"=K{r}*L{r}+M{r}*N{r}") # Nota parcial 2
                ws.cell(row=r, column=16, value=f"=O{r}*P11")            # Ponderación 30% p2
                ws.cell(row=r, column=21, value=f"=Q{r}*R{r}+S{r}*T{r}") # Nota parcial 3
                ws.cell(row=r, column=22, value=f"=U{r}*V11")            # Ponderación 40% p3
                ws.cell(row=r, column=23, value=f"=J{r}+P{r}+V{r}")      # Promedio final completo
                
                for c in range(1, 25):
                    ws.cell(row=r, column=c).font = f_norm; ws.cell(row=r, column=c).border = thin_b

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 8)

            wb.save(output)
            excel_data = output.getvalue()
            
            # --- INTERFAZ FINAL PARA EL USUARIO ---
            st.success("🎉 ¡Tu archivo de Excel está listo!")
            
            # Botón enorme y llamativo de descarga
            st.download_button(
                label="🟢 HACER CLIC AQUÍ PARA DESCARGAR EL EXCEL DE NOTAS",
                data=excel_data,
                file_name=f"Cuadro_Notas_{a_txt.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )