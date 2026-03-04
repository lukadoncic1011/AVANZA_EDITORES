import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, IntVar, Label, Radiobutton, Button 
import json
import os
import sys
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Flowable
from reportlab.lib import colors
from reportlab.lib.units import inch
import numpy as np 
from PIL import Image, ImageTk
import subprocess

#reportes en pdf individuales y generales
#boton en AVANZA_EDITORES


def resource_path(relative_path):
    """
    Obtiene la ruta absoluta a un recurso, funciona para desarrollo y para
    PyInstaller al generar el ejecutable.
    """
    try:
        # PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Si no está en PyInstaller, usa la ruta base del script (donde se ejecuta)
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- COLORES PARA ESTÉTICA (Copiados de AVANZA_EDITORES_3.0.py) ---
COLOR_FONDO_CIELO = "#E0FFFF" # Azul Cielo
COLOR_AZUL_REY = "#0D1096"   # Azul Rey
COLOR_TEXTO_PRINCIPAL = "#004D99" # Azul Oscuro
COLOR_NARANJA = "#D87F4B" # NARANJA para encabezados de tabla
# ------------------------------------------

# --- Lógica de ReportLab para la generación de PDFs ---

# FUNCIÓN DE UTILIDAD: Nivel de Desempeño
def obtener_nivel(puntaje_100, area):
    """
    Calcula el nivel de desempeño a partir de un puntaje en escala 0-100.
    """
    # Lógica especial para Inglés (niveles B2, B1, A2, A1)
    if area.lower() == "inglés": 
        # Umbrales ejemplo (asumiendo que puntaje_100 es de 0 a 100):
        if puntaje_100 >= 85: return "B2"
        elif puntaje_100 >= 70: return "B1"
        elif puntaje_100 >= 55: return "A2"
        else: return "A1"
            
    # Lógica para otras áreas
    if puntaje_100 >= 85: return "Satisfactorio"
    elif puntaje_100 >= 70: return "Aceptable"
    elif puntaje_100 >= 50: return "Mínimo"
    else: return "Insuficiente"

# FUNCIÓN DE UTILIDAD: Percentil por Área
def obtener_percentil_area(area, puntaje_100):
    """
    Simula el percentil de rendimiento por área (usando puntaje_100 como percentil).
    """
    return puntaje_100

# FUNCIÓN DE UTILIDAD: Generar Barra de Percentil con Texto
def generar_percentil_barra(canvas, x, y, percentil, ancho=2.5 * inch, alto=0.2 * inch):
    """
    Dibuja una barra de percentil simple con marcas y el valor numérico.
    """
    percentil = max(0, min(100, percentil))
    line_y = y + alto / 2
    canvas.setStrokeColor(colors.lightgrey)
    canvas.setLineWidth(1)
    canvas.line(x, line_y, x + ancho, line_y)
    
    canvas.setStrokeColor(colors.black)
    canvas.setLineWidth(0.5)
    canvas.setFont("Helvetica", 7)
    
    for i in range(11):
        mark_value = i * 10
        mark_x = x + ancho * (mark_value / 100.0)
        mark_height = 5 if i % 5 == 0 else 3 
        canvas.line(mark_x, line_y - mark_height / 2, mark_x, line_y + mark_height / 2)
        canvas.drawCentredString(mark_x, line_y - mark_height / 2 - 8, str(mark_value))
        
    indicator_x = x + ancho * (percentil / 100.0)
    
    canvas.setStrokeColor(colors.red)
    canvas.setLineWidth(1)
    canvas.line(indicator_x, line_y, indicator_x, line_y + 10)
    
    canvas.setFillColor(colors.black)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawCentredString(indicator_x, line_y + 12, f"{int(percentil)}")

# --- NUEVA FUNCIÓN DE UTILIDAD ---
def contar_estudiantes_por_nivel(resultados, areas):
    """
    Cuenta cuántos estudiantes obtuvieron cada nivel de desempeño para cada área.
    """
    conteo_por_area = {}
    
    # Definir los niveles que esperamos
    niveles_base = ["Satisfactorio", "Aceptable", "Mínimo", "Insuficiente"]
    niveles_ingles = ["B2", "B1", "A2", "A1"]
    
    for area in areas:
        niveles_a_contar = niveles_ingles if area.lower() == 'inglés' else niveles_base
        
        # Inicializar el conteo para esta área con cero en todos los niveles
        conteo_por_area[area] = {nivel: 0 for nivel in niveles_a_contar}
        
        for estudiante in resultados:
            puntaje_100 = int(estudiante.get(f"Puntaje_{area}", 0))
            nivel = obtener_nivel(puntaje_100, area)
            
            # Incrementar el contador si el nivel es uno de los esperados
            if nivel in conteo_por_area[area]:
                 conteo_por_area[area][nivel] += 1

    return conteo_por_area


# NUEVA FUNCIÓN DE UTILIDAD PARA DIBUJAR LAS BARRAS EN EL PDF INDIVIDUAL
def dibujar_grafico_barras_comparacion(canvas, x_start_pdf, y_start_pdf, estudiante, promedios_grupo, areas, width=7 * inch, height=3 * inch):
    """
    Dibuja el gráfico de barras comparando el puntaje del estudiante vs. el promedio del grupo.
    (La lógica de esta función se mantiene sin cambios, ya que se le pasa la lista de 'areas' filtrada o completa).
    """
    c = canvas
    
    data_grafico = {}
    for area in areas:
        puntaje_estudiante = int(estudiante.get(f"Puntaje_{area}", 0))
        puntaje_grupo = int(promedios_grupo.get(area, 0)) 
        
        data_grafico[area] = {
            'Estudiante': puntaje_estudiante,
            'Grupo': puntaje_grupo
        }

    x_margin = 0.5 * inch
    y_margin = 0.5 * inch
    plot_width = width - 2 * x_margin
    plot_height = height - y_margin - 0.7 * inch 
    x_plot_start = x_start_pdf + x_margin
    y_plot_start = y_start_pdf + y_margin

    areas_list = list(data_grafico.keys())
    num_areas = len(areas_list)
    
    if num_areas == 0:
        return

    group_width = plot_width / num_areas
    bar_width = group_width * 0.25 
    bar_spacing = group_width * 0.05
    
    # 1. Título y Leyenda
    c.setFont("Courier-Bold", 12)
    c.drawCentredString(x_start_pdf + width / 2, y_start_pdf + height - 0.2 * inch, "COMPARACIÓN DE PUNTAJES POR ÁREA")

    legend_y = y_plot_start + plot_height + 0.15 * inch
    c.setFont("Courier", 10)
    
    c.setFillColor(colors.blue)
    c.rect(x_start_pdf + 0.5 * inch, legend_y, 0.15 * inch, 0.15 * inch, fill=1)
    c.setFillColor(colors.black)
    c.drawString(x_start_pdf + 0.5 * inch + 0.2 * inch, legend_y + 2, "Estudiante")
    
    c.setFillColor(colors.orange)
    c.rect(x_start_pdf + 2.5 * inch, legend_y, 0.15 * inch, 0.15 * inch, fill=1)
    c.setFillColor(colors.black)
    c.drawString(x_start_pdf + 2.5 * inch + 0.2 * inch, legend_y + 2, "Grupo")


    # 2. Dibujar Eje Y (Escala 0 a 100)
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(x_plot_start, y_plot_start, x_plot_start, y_plot_start + plot_height)
    
    c.setFont("Courier", 8)
    for i in range(0, 101, 20):
        y_pos = y_plot_start + plot_height * (i / 100.0)
        c.line(x_plot_start - 3, y_pos, x_plot_start, y_pos) 
        c.drawString(x_plot_start - 20, y_pos - 3, str(i))

    # 3. Dibujar las Barras y Etiquetas
    for i, area in enumerate(areas_list):
        puntaje_estudiante = data_grafico[area]['Estudiante']
        puntaje_grupo = data_grafico[area]['Grupo']

        center_x = x_plot_start + group_width * (i + 0.5)
        
        # Barra Estudiante (Azul)
        bar_x_est = center_x - bar_width - bar_spacing
        bar_h_est = plot_height * (puntaje_estudiante / 100.0)
        c.setFillColor(colors.blue)
        c.rect(bar_x_est, y_plot_start, bar_width, bar_h_est, fill=1)
        
        # Etiqueta Estudiante
        c.setFillColor(colors.black)
        c.setFont("Courier-Bold", 8)
        c.drawCentredString(bar_x_est + bar_width / 2, y_plot_start + bar_h_est + 2, str(puntaje_estudiante))
        
        # Barra Grupo (Naranja)
        bar_x_grp = center_x + bar_spacing
        bar_h_grp = plot_height * (puntaje_grupo / 100.0)
        c.setFillColor(colors.orange)
        c.rect(bar_x_grp, y_plot_start, bar_width, bar_h_grp, fill=1)
        
        # Etiqueta Grupo
        c.setFillColor(colors.black)
        c.setFont("Courier-Bold", 8)
        c.drawCentredString(bar_x_grp + bar_width / 2, y_plot_start + bar_h_grp + 2, str(puntaje_grupo))
        
        # Etiqueta del Área (Eje X)
        c.setFont("Courier-Bold", 10)
        c.drawCentredString(center_x, y_plot_start - 20, area.replace('_', ' ').upper())
        
        # Línea de la base del gráfico (Eje X)
        c.line(x_plot_start, y_plot_start, x_plot_start + plot_width, y_plot_start)


# FUNCIÓN PRINCIPAL MODIFICADA: Ahora acepta el parámetro `escala`
def generar_pdf_individual(estudiante, metadata, c, escala):
    """Genera el contenido de una página (reporte individual) en el PDF con el formato deseado."""
    
    nombre = estudiante["Nombre"]
    id_estudiante = estudiante["ID_Estudiante"]
    puntaje_global = estudiante["Puntaje_Global"]
    percentil_global = estudiante["Percentil_Global"]
    puesto = estudiante["Puesto"]

    colegio = metadata["colegio"]
    grado = metadata["grado"]
    examen = metadata["examen"]
    ciudad = metadata["ciudad"]
    fecha = metadata["fecha"]
    areas_principales = metadata["areas_principales_calificadas"]
    promedios_grupo = metadata["promedios_grupo"]
    
    ancho, alto = letter
    
    
    # -----------------------------------------------------------------------
    # 1. MARCO GRANDE DE LA PÁGINA 
    # -----------------------------------------------------------------------
    margin = 0.5 * inch 
    c.setStrokeColor(colors.blue)  
    c.setLineWidth(2)
    c.setDash(5, 5) 
    c.rect(margin, margin, ancho - 2 * margin, alto - 2 * margin)
    c.setDash(1, 0)
    c.setLineWidth(1) 
    
    # -----------------------------------------------------------------------
    # 0. LOGO (AÑADIDO)
    # -----------------------------------------------------------------------
    
    banner_file = "imagenes/logo.png" 
    banner_width = 3.6 * inch
    banner_height = 1 * inch
    x_banner_start = ancho - 5 * inch - banner_width
    y_banner_start = alto - 0.6 * inch - banner_height 
    
    try:
        c.drawImage(
            banner_file, 
            x_banner_start, 
            y_banner_start, 
            width=banner_width, 
            height=banner_height,
            preserveAspectRatio=True
        )
        
       
    except Exception as e:
        c.setFillColor(colors.red)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(ancho / 2, y_banner_start + banner_height / 2, "ERROR: No se encontró 'banner_examen.png'")
    

# -----------------------------------------------------------------------
    # 0. BANNER (AÑADIDO)
    # -----------------------------------------------------------------------
    
    banner_file = "imagenes/avanza_editores.png" 
    banner_width = 5 * inch
    banner_height = 2 * inch
    x_banner_start = ancho - 0.7 * inch - banner_width
    y_banner_start = alto - 0 * inch - banner_height 
    
    try:
        c.drawImage(
            banner_file, 
            x_banner_start, 
            y_banner_start, 
            width=banner_width, 
            height=banner_height,
            preserveAspectRatio=True
        )
        
       
    except Exception as e:
        c.setFillColor(colors.red)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(ancho / 2, y_banner_start + banner_height / 2, "ERROR: No se encontró 'banner_examen.png'")
    

    # 1. Cabecera (Datos Personales y Puesto) - Lado IZQUIERDO
    x_start_izq = inch
    y_start_izq = alto - 120
    ancho_izq = 3.7 * inch
    alto_izq = 147 
    
    c.setFillColor(colors.white)
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    c.rect(x_start_izq, y_start_izq - alto_izq, ancho_izq, alto_izq, fill=1)

    text_y = y_start_izq - 20 

    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 12)
    c.drawString(x_start_izq + 0.1 * inch, text_y, "Nombres y apellidos")
    c.setFont("Courier", 12)
    c.drawString(x_start_izq + 0.1 * inch, text_y - 15, nombre.upper())
    
    campos_izq = [
        ("Fecha", fecha.upper()),
        ("Municipio", ciudad.replace('_', ' ').title()),
        ("Institución", colegio.upper()),
        ("Grado", grado),
    ]
    
    current_y_campo = text_y - 35 
    c.setLineWidth(1)
    
    for label, value in campos_izq:
        c.setFont("Courier-Bold", 12)
        c.drawString(x_start_izq + 0.1 * inch, current_y_campo, f"{label}:")
        
        c.setFont("Courier", 10)
        c.drawString(x_start_izq + 1.5 * inch, current_y_campo, value)
        c.line(x_start_izq + 1.5 * inch, current_y_campo - 2, x_start_izq + ancho_izq - 0.1 * inch, current_y_campo - 2)
        
        current_y_campo -= 13

    puesto_y = y_start_izq - alto_izq + 0.2 * inch
    
    c.setFillColor(colors.lightblue)
    c.rect(x_start_izq, puesto_y, ancho_izq * 0.7, 0.4 * inch, fill=1)
    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 10)
    c.drawString(x_start_izq + 0.1 * inch, puesto_y + 0.15 * inch, "Puesto que ocupó en el grupo")
    
    c.setFillColor(colors.white)
    c.rect(x_start_izq + ancho_izq * 0.7, puesto_y, ancho_izq * 0.3, 0.4 * inch, fill=1)
    c.setFillColor(colors.black)
    c.drawCentredString(x_start_izq + ancho_izq * 0.85, puesto_y + 0.15 * inch, str(puesto))
    
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    c.line(x_start_izq, puesto_y, x_start_izq + ancho_izq, puesto_y) 
    c.line(x_start_izq + ancho_izq * 0.7, puesto_y, x_start_izq + ancho_izq * 0.7, puesto_y + 0.4 * inch) 


    # 2. Cabecera (Puntaje Global y Percentil) - Lado DERECHO
    x_start_der = x_start_izq + ancho_izq + 0.1 * inch 
    y_start_der = alto - 120
    ancho_der = 3.0 * inch
    alto_der = 147 
    
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    c.rect(x_start_der, y_start_der - alto_der, ancho_der, alto_der, fill=0)

    puntaje_y = y_start_der - 30
    
    c.setFillColor(colors.lightblue)
    c.rect(x_start_der, puntaje_y, ancho_der, 30, fill=1)
    
    c.setFillColor(colors.black)
    c.setFont("Courier", 11)
    c.drawCentredString(x_start_der + ancho_der / 2, puntaje_y + 20, f"Examen {examen}")
    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 11)
    c.drawCentredString(x_start_der + ancho_der / 2, puntaje_y + 7, "PUNTAJE GLOBAL")
    
    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 30)
    
    puntaje_escalado = int(puntaje_global)
    text_width_puntaje = c.stringWidth(str(puntaje_escalado), "Courier-Bold", 30)
    
    center_x = x_start_der + ancho_der / 2
    pos_puntaje_x = center_x - 5 
    c.drawString(pos_puntaje_x - text_width_puntaje, puntaje_y - 30, str(puntaje_escalado))
    
    c.setFillColor(colors.orangered)
    c.drawString(pos_puntaje_x, puntaje_y - 30, f" /{escala}")

    percentil_y = puntaje_y - 60
    
    c.setFillColor(colors.lightblue)
    c.rect(x_start_der, percentil_y, ancho_der, 20, fill=1)
    
    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 11)
    c.drawCentredString(x_start_der + ancho_der / 2, percentil_y + 5, "¿EN QUÉ PERCENTIL ME ENCUENTRO?")
    
    c.setFillColor(colors.black)
    c.setFont("Courier", 7)
    c.drawCentredString(x_start_der + ancho_der / 2, percentil_y - 15, "con respecto a los estudiantes del grupo")
    
    barra_x = x_start_der + (ancho_der - 2.5 * inch) / 2 
    barra_y = percentil_y - 50 
    generar_percentil_barra(c, barra_x, barra_y, percentil_global, ancho=2.5 * inch)
    
    current_y_end_header = y_start_izq - alto_izq - 20 

    # 3. Tabla de Puntajes (Área, Puntaje, Nivel, GRÁFICA)
    c.setFont("Courier-Bold", 14)
    c.drawCentredString(ancho / 2, current_y_end_header, "RESULTADOS POR ÁREA")
    
    start_y = current_y_end_header - 30
    row_height = 0.5 * inch 
    
    col_width_tabla = 6.8 * inch 
    col_width_area = 1.9 * inch
    col_width_puntaje = 0.6 * inch
    col_width_nivel = 1.2 * inch
    col_width_grafica = 3.3 * inch 
    
    col_area_x_start = inch 
    col_puntaje_x_start = col_area_x_start + col_width_area 
    col_nivel_x_start = col_puntaje_x_start + col_width_puntaje 
    col_grafica_x_start = col_nivel_x_start + col_width_nivel 
    
    c.setFillColor(colors.skyblue)
    c.rect(inch, start_y, col_width_tabla, 0.25 * inch, fill=1) 
    
    c.setFillColor(colors.black) 
    c.setFont("Courier-Bold", 10)
    
    c.drawString(col_area_x_start + 0.1 * inch, start_y + 0.08 * inch, "ÁREA")
    c.drawCentredString(col_puntaje_x_start + col_width_puntaje/2, start_y + 0.08 * inch, "%") 
    c.drawCentredString(col_nivel_x_start + col_width_nivel/2, start_y + 0.08 * inch, "NIVEL") 
    c.drawCentredString(col_grafica_x_start + col_width_grafica/2, start_y + 0.08 * inch, "GRÁFICA") 
    
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(col_puntaje_x_start, start_y, col_puntaje_x_start, start_y + 0.25 * inch)
    c.line(col_nivel_x_start, start_y, col_nivel_x_start, start_y + 0.25 * inch)
    c.line(col_grafica_x_start, start_y, col_grafica_x_start, start_y + 0.25 * inch)
    
    c.rect(inch, start_y, col_width_tabla, 0.25 * inch, fill=0)
    
    current_y = start_y
    
    for i, area in enumerate(areas_principales):
        current_y -= row_height 
        
        puntaje_estudiante = estudiante.get(f"Puntaje_{area}", 0) 
        
        if isinstance(puntaje_estudiante, (int, float)):
            puntaje_entero_100 = int(puntaje_estudiante) 
            puntaje_100 = puntaje_entero_100
            puntaje_formateado = str(puntaje_entero_100)
            nivel_desempeño = obtener_nivel(puntaje_100, area)
            percentil_area = puntaje_100 
        else:
            puntaje_100 = 0
            puntaje_formateado = "N/A"
            nivel_desempeño = "N/A"
            percentil_area = 0
            
        
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        c.line(inch, current_y, inch + col_width_tabla, current_y)

        c.setLineWidth(0.5)
        c.setDash(2, 2) 
        
        c.line(col_puntaje_x_start, current_y, col_puntaje_x_start, current_y + row_height)
        c.line(col_nivel_x_start, current_y, col_nivel_x_start, current_y + row_height)
        c.line(col_grafica_x_start, current_y, col_grafica_x_start, current_y + row_height)
        
        c.setDash(1, 0)
        
        c.setLineWidth(1)
        c.line(inch + col_width_tabla, current_y, inch + col_width_tabla, current_y + row_height)
        c.line(inch, current_y, inch, current_y + row_height)
        
        c.setFillColor(colors.black)
        c.setFont("Courier-Bold", 10)
        area_display = area.replace('_', ' ')
        c.drawString(col_area_x_start + 0.1 * inch, current_y + row_height / 2 + 5, area_display.upper())

        c.setFont("Courier-Bold", 12)
        c.drawCentredString(col_puntaje_x_start + col_width_puntaje/2, current_y + row_height / 2 + 5, puntaje_formateado)

        c.setFont("Courier-Bold", 10) 
        c.drawCentredString(col_nivel_x_start + col_width_nivel/2, current_y + row_height / 2 + 5, nivel_desempeño.upper())
        
        if percentil_area > 0 or area.lower() == "inglés": 
            barra_area_x = col_grafica_x_start + 0.2 * inch
            barra_area_y = current_y + row_height / 2 - 10
            generar_percentil_barra(c, barra_area_x, barra_area_y, percentil_area, ancho=2.5 * inch)
            
    c.setLineWidth(1)
    c.setDash(1, 0) 
    c.line(inch, current_y, inch + col_width_tabla, current_y)
    
    
    # -----------------------------------------------------------------------
    # 4. GRÁFICO DE BARRAS COMPARATIVO (TODAS LAS ÁREAS)
    # -----------------------------------------------------------------------
    
    grafico_y_start = current_y - 0.2 * inch
    grafico_x_start = inch
    
    dibujar_grafico_barras_comparacion(
        c, 
        grafico_x_start, 
        grafico_y_start - 2.5 * inch, 
        estudiante, 
        promedios_grupo, 
        areas_principales, 
        width=6.5 * inch, 
        height=2.5 * inch
    )


# --- CLASE PARA DIBUJAR LOS GRÁFICOS DE BARRAS EN EL REPORTE GENERAL (Flowable) ---
class ComparacionBarras(Flowable):
    """Clase personalizada para dibujar los gráficos de barras en ReportLab (usada en el reporte general)."""
    def __init__(self, data, width=7 * inch, height=3 * inch):
        Flowable.__init__(self)
        self.data = data 
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        areas = list(self.data.keys())
        
        x_margin = 0.5 * inch
        y_margin = 0.5 * inch
        
        plot_width = self.width - 2 * x_margin
        plot_height = self.height - y_margin - 0.7 * inch
        
        x_start = x_margin
        y_start = y_margin
        
        areas_list = list(self.data.keys())
        num_areas = len(areas_list)
        
        if num_areas == 0:
            return

        group_width = plot_width / num_areas
        bar_width = group_width * 0.25 
        bar_spacing = group_width * 0.05 
        
        # 1. Título y Leyenda
        legend_y = y_start + plot_height + 0.15 * inch
        c.setFont("Helvetica", 10)
        
        c.setFillColor(colors.blue)
        c.rect(x_start, legend_y, 0.15 * inch, 0.15 * inch, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_start + 0.2 * inch, legend_y + 2, "Estudiante (Ejemplo)")
        
        c.setFillColor(colors.orange)
        c.rect(x_start + 2.5 * inch, legend_y, 0.15 * inch, 0.15 * inch, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_start + 2.5 * inch + 0.2 * inch, legend_y + 2, "Grupo (Promedio)")

        # 2. Dibujar Eje Y (Escala 0 a 100)
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        c.line(x_start, y_start, x_start, y_start + plot_height)
        
        c.setFont("Helvetica", 8)
        for i in range(0, 101, 10):
            y_pos = y_start + plot_height * (i / 100.0)
            c.line(x_start - 3, y_pos, x_start, y_pos)
            c.drawString(x_start - 20, y_pos - 3, str(i))
            c.setStrokeColor(colors.lightgrey)
            c.line(x_start, y_pos, x_start + plot_width, y_pos) 

        # 3. Dibujar las Barras y Etiquetas
        for i, area in enumerate(areas_list):
            puntaje_estudiante = int(self.data[area]['Estudiante'])
            puntaje_grupo = int(self.data[area]['Grupo'])

            center_x = x_start + group_width * (i + 0.5)
            
            
            # Barra Grupo (Naranja)
            bar_x_grp = center_x + bar_spacing
            bar_h_grp = plot_height * (puntaje_grupo / 100.0)
            c.setFillColor(colors.orange)
            c.rect(bar_x_grp, y_start, bar_width, bar_h_grp, fill=1, stroke=0)
            
            # Etiqueta Grupo
            c.setFillColor(colors.black)
            c.setFont("Courier-Bold", 8)
            c.drawCentredString(bar_x_grp + bar_width / 2, y_start + bar_h_grp + 2, str(puntaje_grupo))
            
            # Etiqueta del Área (Eje X)
            c.setFont("Courier-Bold", 8)
            c.drawCentredString(center_x, y_start - 20, area.replace('_', ' ').upper())
            
        # Línea de la base del gráfico (Eje X)
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        c.line(x_start, y_start, x_start + plot_width, y_start)

# --- CLASE FLOWABLE PARA GRÁFICO DE DISTRIBUCIÓN POR NIVEL (VERSIÓN FINAL) ---
class DistribucionNivelesBarra(Flowable):
    """
    Dibuja un diagrama de barras por área donde cada barra representa la cantidad
    de estudiantes en un nivel específico. Usa un solo color (Azul Rey) y es
    visualmete consistente con ComparacionBarras.
    """
    def __init__(self, area_name, data, width=7.0 * inch, height=3.0 * inch):
        Flowable.__init__(self)
        self.area_name = area_name
        self.data = data  # {nivel: conteo, ...}
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        
        # Parámetros de Dibujo (Similares a ComparacionBarras)
        x_margin = 0.5 * inch
        y_margin = 0.5 * inch
        
        plot_width = self.width - 2 * x_margin
        plot_height = self.height - y_margin - 0.7 * inch
        
        x_start = x_margin
        y_start = y_margin
        
        total_estudiantes = sum(self.data.values())
        
        if total_estudiantes == 0:
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x_start, y_start, f"No hay datos para el área {self.area_name.upper()}")
            return
            
        # 1. Título y Color de Barra
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(self.width / 2, self.height - 0.2 * inch, f"DISTRIBUCIÓN DE ESTUDIANTES EN: {self.area_name.replace('_', ' ').upper()}")
        
        BAR_COLOR = colors.navy # Azul Rey / Navy
        
        # 2. Definir Orden de los Niveles (Eje X)
        if self.area_name.lower() == 'inglés':
            niveles_ordenados = ["B2", "B1", "A2", "A1"]
        else:
            niveles_ordenados = ["Satisfactorio", "Aceptable", "Mínimo", "Insuficiente"]
        
        # 3. Determinar el Máximo Conteo para la Escala Y
        max_conteo_nivel = max(self.data.values()) if self.data.values() else 1
        
        # Ajuste al múltiplo de 5 más cercano por encima
        if max_conteo_nivel < 5:
            max_y_scale = 5
        elif max_conteo_nivel <= 10:
            max_y_scale = 10
        else:
            max_y_scale = int(np.ceil(max_conteo_nivel / 5.0)) * 5 
        
        num_niveles = len(niveles_ordenados)
        group_width = plot_width / num_niveles
        bar_width = group_width * 0.5 # Ancho de cada barra de nivel
        
        # 4. Dibujar Eje Y (Escala 0 al Máximo Conteo)
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        c.line(x_start, y_start, x_start, y_start + plot_height)
        
        # Establecer pasos para el eje Y
        step = max(1, int(max_y_scale / 5)) if max_y_scale > 10 else 1
        step = max(step, 1) 
            
        c.setFont("Helvetica", 8)
        
        for i in range(0, max_y_scale + step, step):
            if i > max_y_scale: continue
            y_pos = y_start + plot_height * (i / max_y_scale)
            c.setStrokeColor(colors.black)
            c.line(x_start - 3, y_pos, x_start, y_pos) 
            c.drawString(x_start - 20, y_pos - 3, str(i))
            
            # Líneas de la cuadrícula
            c.setStrokeColor(colors.lightgrey)
            c.setLineWidth(0.5)
            c.line(x_start, y_pos, x_start + plot_width, y_pos) 

        # 5. Dibujar las Barras y Etiquetas (Niveles en el Eje X)
        for i, nivel in enumerate(niveles_ordenados):
            conteo = self.data.get(nivel, 0)
            
            center_x = x_start + group_width * (i + 0.5)
            
            # Calcular posición y altura de la barra
            bar_x = center_x - bar_width / 2
            bar_h = plot_height * (conteo / max_y_scale)
            
            # Dibujar la Barra (Azul Rey)
            c.setFillColor(BAR_COLOR)
            c.rect(bar_x, y_start, bar_width, bar_h, fill=1, stroke=0) 
            
            # Etiqueta de Conteo (Encima de la barra)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(bar_x + bar_width / 2, y_start + bar_h + 2, str(conteo))
            
            # Etiqueta del Nivel (Eje X)
            c.setFont("Helvetica-Bold", 8) # Tamaño 8 para mejorar la visibilidad
            
            # Lógica de rotación para nombres largos (ej. Satisfactorio, Insuficiente)
            if len(nivel) > 20: 
                c.saveState()
                # Transladar al centro de la barra, ligeramente abajo del Eje X
                c.translate(center_x, y_start - 10) 
                c.rotate(45) # Rotar 45 grados
                c.drawCentredString(0, 0, nivel.upper())
                c.restoreState()
            else:
                c.drawCentredString(center_x, y_start - 20, nivel.upper())
        
        # 6. Línea de la base del gráfico (Eje X) y Total
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        c.line(x_start, y_start, x_start + plot_width, y_start) 
        
        c.setFont("Helvetica", 9)
        c.drawCentredString(self.width / 2, y_start - 45, f"Total Estudiantes Evaluados: {total_estudiantes}")

# --- Lógica de la Aplicación Tkinter ---

class ReporteApp:
    def __init__(self, master):
        self.master = master
        master.title("Generador de Reportes de Examen")
        
        self.archivo_json = None
        self.escala_global = 500 

        self.label_estado = tk.Label(master, text="Esperando cargar archivo JSON...", fg="red")
        self.label_estado.pack(pady=10)

        self.btn_cargar = tk.Button(master, text="Cargar Archivo JSON", command=self.cargar_json)
        self.btn_cargar.pack(pady=5)

        self.btn_individual = tk.Button(master, text="Generar Reportes Individuales (PDF)", 
                                        command=self.generar_individuales, state=tk.DISABLED)
        self.btn_individual.pack(pady=5)

        self.btn_general = tk.Button(master, text="Generar Reporte General (PDF)", 
                                     command=self.generar_general, state=tk.DISABLED)
        self.btn_general.pack(pady=5)
        
        self.btn_salir = tk.Button(master, text="Salir", command=master.quit)
        self.btn_salir.pack(pady=10)
        
    def preguntar_escala(self):
        self.wait_var = IntVar(self.master)
        top = Toplevel(self.master)
        top.title("Seleccionar Escala Global")
        top.transient(self.master) 
        top.grab_set() 

        Label(top, text="Por favor, seleccione la escala máxima del puntaje global:").pack(padx=10, pady=10)
        
        self.escala_var = IntVar(top, value=self.escala_global)

        def set_escala():
            self.escala_global = self.escala_var.get()
            top.destroy()
            self.wait_var.set(1) 
            
        Radiobutton(top, text="Escala sobre 500", variable=self.escala_var, value=500).pack(anchor='w', padx=20)
        Radiobutton(top, text="Escala sobre 100", variable=self.escala_var, value=100).pack(anchor='w', padx=20)
        
        Button(top, text="Confirmar Escala", command=set_escala).pack(pady=15)
        
        self.master.wait_variable(self.wait_var)
        
        if self.escala_global not in [100, 500]:
            self.escala_global = 500

    def calcular_promedios(self, resultados, areas):
        promedios = {}
        for area in areas:
            key = f"Puntaje_{area}"
            puntajes = [e.get(key, 0) for e in resultados if isinstance(e.get(key), (int, float))]
            if puntajes:
                promedios[area] = np.mean(puntajes)
            else:
                promedios[area] = 0
        return promedios

    def cargar_json(self):
        filepath = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        if not filepath:
            return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
            self.archivo_json = filepath
            
            self.preguntar_escala()
            
            self.data["resultados_estudiantes"].sort(key=lambda x: x.get("Puesto", float('inf')))
            
            areas = self.data["metadata"]["areas_principales_calificadas"]
            promedios = self.calcular_promedios(self.data["resultados_estudiantes"], areas)
            self.data["metadata"]["promedios_grupo"] = promedios
            
            filename = os.path.basename(filepath)
            self.label_estado.config(text=f"Archivo cargado: {filename} (Escala: {self.escala_global})", fg="green")
            self.btn_individual.config(state=tk.NORMAL)
            self.btn_general.config(state=tk.NORMAL)
            messagebox.showinfo("Éxito", f"Archivo '{filename}' cargado correctamente.\nEscala Global: {self.escala_global}.")

        except Exception as e:
            messagebox.showerror("Error de Carga", f"No se pudo cargar el archivo JSON: {e}")
            self.archivo_json = None
            self.label_estado.config(text="Error al cargar el archivo JSON.", fg="red")
            self.btn_individual.config(state=tk.DISABLED)
            self.btn_general.config(state=tk.DISABLED)

    def generar_individuales(self):
        if not self.archivo_json:
            messagebox.showerror("Error", "Primero debe cargar un archivo JSON.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile="Reportes_Individuales_Examen.pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_path:
            return

        try:
            c = canvas.Canvas(output_path, pagesize=letter)
            metadata = self.data["metadata"]
            escala = self.escala_global

            for estudiante in self.data["resultados_estudiantes"]:
                generar_pdf_individual(estudiante, metadata, c, escala)
                c.showPage()

            c.save() 
            
            messagebox.showinfo("Éxito", f"Reporte individual generado exitosamente en:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error de Generación", f"Ocurrió un error al generar el PDF: {e}")

    def generar_general(self):
        """Genera un PDF con la tabla general, promedios, gráfico Max/Min y gráficos de distribución."""
        if not self.archivo_json:
            messagebox.showerror("Error", "Primero debe cargar un archivo JSON.")
            return
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile="Reporte_General_Examen.pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_path:
            return

        try:
            doc = SimpleDocTemplate(output_path, pagesize=letter)
            styles = getSampleStyleSheet()
            flowables = []
            
            metadata = self.data["metadata"]
            resultados = self.data["resultados_estudiantes"] 
            areas = metadata["areas_principales_calificadas"]
            promedios = metadata.get("promedios_grupo", self.calcular_promedios(resultados, areas))
            
            # --- TABLA DE METADATA (COLEGIO, GRADO, ETC.) ---
            TABLE_CONTENT_WIDTH = 6.5 * inch 
            col_width = TABLE_CONTENT_WIDTH / 2
            
            data_cabecera = [
                [Paragraph(f"<b>REPORTE GENERAL</b>", styles['Heading3']), ''],
                [f"Colegio: {metadata['colegio']}", f"Grado: {metadata['grado']}"],
                [f"Examen: {metadata['examen']}", f"Fecha: {metadata['fecha']}"]
            ]

            tabla_cabecera = Table(data_cabecera, colWidths=[col_width, col_width])

            tabla_cabecera.setStyle(TableStyle([
                ('SPAN', (0, 0), (1, 0)),
                ('ALIGN', (0, 0), (1, 0), 'CENTER'),
                ('VALIGN', (0, 0), (1, 0), 'MIDDLE'),
                ('BACKGROUND', (0, 0), (1, 0), colors.skyblue),
                ('LINEBELOW', (0, 0), (1, 0), 1, colors.black),
                ('BOTTOMPADDING', (0, 0), (1, 0), 8),
                ('ALIGN', (0, 1), (1, 2), 'LEFT'),
                ('VALIGN', (0, 1), (1, 2), 'TOP'),
                ('LEFTPADDING', (0, 1), (1, 2), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey) 
            ]))
            
            flowables.append(tabla_cabecera)
            flowables.append(Spacer(1, 0.5 * inch))

            # --- TABLA GENERAL DE ESTUDIANTES ---
            headers = ["#", "Nombre"] + areas + ["Global"]
            table_data = [headers]

            for e in resultados:
                row = [e["Puesto"], e["Nombre"]]
                for area in areas:
                    row.append(str(int(e.get(f"Puntaje_{area}", 0))))
                row.append(str(int(e['Puntaje_Global'])))
                table_data.append(row)

            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.skyblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ])

            col_widths = [0.2*inch, 2.4*inch]
            col_area_width = (8.5*inch - 0.5*inch - 2.1*inch - 1.8*inch - 0.7*inch) / len(areas)
            col_widths.extend([col_area_width] * len(areas))
            col_widths.extend([0.5*inch]) # Ajustado para solo 'Global'

            tabla = Table(table_data, colWidths=col_widths)
            tabla.setStyle(table_style)
            flowables.append(tabla)
            flowables.append(Spacer(1, 0.5 * inch))

            # --- TABLA DE PROMEDIOS DEL GRUPO ---
            
            promedio_headers = ["Área"]
            promedio_values = ["Puntaje Promedio"]
            
            puntajes_globales_grupo = [e["Puntaje_Global"] for e in resultados]
            promedio_global = np.mean(puntajes_globales_grupo) if puntajes_globales_grupo else 0
            
            data_grafico = {} 
            
            for area in areas:
                promedio_headers.append(area)
                promedio_value = promedios.get(area, 0)
                promedio_values.append(f"{promedio_value:.2f}") 
                
                # Para el Flowable, usamos el primer estudiante como ejemplo de comparación
                puntaje_estudiante = int(resultados[0].get(f"Puntaje_{area}", 0)) if resultados else 0
                data_grafico[area] = {
                    'Estudiante': puntaje_estudiante,
                    'Grupo': int(promedio_value) 
                }
            
            promedio_headers.append("Global")
            promedio_values.append(f"{promedio_global:.2f}")

            promedio_data = [promedio_headers, promedio_values]
            
            promedio_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.skyblue),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lavender),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
            ])
            
            num_cols = len(promedio_headers)
            col_width_promedio = (8.5 * inch - 2 * inch) / num_cols
            promedio_widths = [col_width_promedio] * num_cols
            promedio_widths[0] = 1.5 * inch
            
            tabla_promedios = Table(promedio_data, colWidths=promedio_widths)
            tabla_promedios.setStyle(promedio_style)
            flowables.append(tabla_promedios)
            flowables.append(Spacer(1, 0.5 * inch))

            # -------------------------------------------------------------------
            # --- MODIFICACIÓN: GRÁFICO DE BARRAS MAX/MIN (Promedio Grupo) ---
            # -------------------------------------------------------------------
            
            puntajes_grupo_referencia = {area: data['Grupo'] for area, data in data_grafico.items()}
            areas_para_grafico_destacado = []
            data_grafico_destacado = {}

            if puntajes_grupo_referencia:
                area_max = max(puntajes_grupo_referencia, key=puntajes_grupo_referencia.get)
                area_min = min(puntajes_grupo_referencia, key=puntajes_grupo_referencia.get)
                
                areas_para_grafico_destacado.append(area_max)
                if area_max != area_min:
                    areas_para_grafico_destacado.append(area_min)

                for area in areas_para_grafico_destacado:
                    data_grafico_destacado[area] = data_grafico[area]
            
            styles = getSampleStyleSheet()
            flowables.append(Paragraph("<b>🎯 ANÁLISIS DESTACADO (Promedio Máx y Mín del Grupo)</b>", styles['h3']))
            flowables.append(Spacer(1, 0.1 * inch))

            if data_grafico_destacado:
                grafico_destacado = ComparacionBarras(
                    data_grafico_destacado, 
                    width=7.0 * inch, 
                    height=3.0 * inch
                )
                flowables.append(grafico_destacado)
                flowables.append(Spacer(1, 0.5 * inch))
            
            # -------------------------------------------------------------------
            # --- NUEVA SECCIÓN: DISTRIBUCIÓN DE ESTUDIANTES POR NIVEL ---
            # -------------------------------------------------------------------
            
            
            conteo_por_area = contar_estudiantes_por_nivel(resultados, areas)
            
            for area in areas:
                data_distribucion = conteo_por_area[area]
                
                grafico_distribucion = DistribucionNivelesBarra(
                    area, 
                    data_distribucion, 
                    width=7.1 * inch, 
                    height=2.5 * inch 
                )
                
                flowables.append(grafico_distribucion)
                flowables.append(Spacer(1, 0.2 * inch)) 

            # -------------------------------------------------------------------

            doc.build(flowables)
            
            messagebox.showinfo("Éxito", f"Reporte general generado exitosamente en:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error de Generación", f"Ocurrió un error al generar el PDF: {e}")


# --- Ejecución de la Aplicación ---
if __name__ == "__main__":
    try:
        import numpy as np
    except ImportError:
        # En un entorno real, la instalación debe hacerse antes. Aquí solo se informa.
        messagebox.showerror("Error de Dependencia", "La librería NumPy es necesaria. Instálala con 'pip install numpy'")
        exit()
        
    root = tk.Tk()
    app = ReporteApp(root)
    root.mainloop()