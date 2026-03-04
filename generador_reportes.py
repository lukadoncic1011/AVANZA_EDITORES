import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, IntVar, Label, Radiobutton, Button 
import json
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Flowable
from reportlab.lib import colors
from reportlab.lib.units import inch
import numpy as np 
from PIL import Image
import subprocess

#reportes en pdf individuales y generales
#boton en AVANZA_EDITORES


# --- Lógica de ReportLab para la generación de PDFs ---

# FUNCIÓN DE UTILIDAD: Nivel de Desempeño
def obtener_nivel(puntaje_100, area):
    """
    Calcula el nivel de desempeño a partir de un puntaje en escala 0-100.
    """
    # Lógica especial para Inglés (niveles B2, B1, A2, A1)
    if area.lower() == "inglés": 
        # Umbrales ejemplo (asumiendo que puntaje_100 es de 0 a 100):
        if puntaje_100 >= 76: return "B2"
        elif puntaje_100 >= 51: return "B1"
        elif puntaje_100 >= 26: return "A2"
        else: return "A1"
            
    # Lógica para otras áreas
    if puntaje_100 >= 76: return "Satisfactorio"
    elif puntaje_100 >= 51: return "Aceptable"
    elif puntaje_100 >= 26: return "Mínimo"
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
    
    #canvas.setFillColor(colors.black)
    #canvas.setFont("Helvetica-Bold", 9)
    #canvas.drawCentredString(indicator_x, line_y + 12, f"{int(percentil)}")

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
def dibujar_grafico_barras_comparacion(c, x, y, estudiante, promedios_grupo, areas, width=6.5*inch, height=3.8*inch):
    """
    Versión maximizada: Eje Y de 10 en 10, etiquetas de 4 letras y mayor altura visual.
    """
    padding = 40
    chart_width = width - 2 * padding
    # Aumentamos chart_height para que las barras tengan más recorrido vertical
    chart_height = height - 35 
    
    # 1. DIBUJAR EJES
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.2)
    # Eje Y vertical (ajustado inicio en y+35)
    c.line(x + padding, y + 35, x + padding, y + 35 + chart_height) 
    # Eje X horizontal
    c.line(x + padding, y + 35, x + padding + chart_width, y + 35) 

    # 2. ESCALA DE 10 EN 10 CON LÍNEAS DE GUÍA
    c.setFont("Helvetica", 7)
    # Cambiamos el step del range a 10
    for i in range(0, 101, 10):
        y_pos = y + 35 + (i / 100) * chart_height
        
        # Dibujamos líneas horizontales tenues para facilitar la lectura
        c.setStrokeColor(colors.lightgrey)
        c.setLineWidth(0.5)
        c.line(x + padding, y_pos, x + padding + chart_width, y_pos)
        
        # Números de la escala
        c.setFillColor(colors.black)
        c.drawRightString(x + padding - 5, y_pos - 2, str(i))

    if not areas:
        return

    n_areas = len(areas)
    # Ajuste de grosores para que se vea balanceado con la nueva altura
    bar_width = (chart_width / n_areas) * 0.38 
    gap = (chart_width / n_areas) * 0.05
    
    # 3. DIBUJAR BARRAS CON ALTURA MAXIMIZADA
    for i, area_original in enumerate(areas):
        etiqueta_visual = area_original[:4].upper()
        
        p_est = int(estudiante.get(f"Puntaje_{area_original}", 0))
        p_grp = int(promedios_grupo.get(area_original, 0))
        
        # El cálculo (puntaje / 100) * chart_height ahora usa el nuevo alto total
        h_est = (p_est / 100) * chart_height
        h_grp = (p_grp / 100) * chart_height

        x_base = x + padding + (i * (chart_width / n_areas)) + gap
        
        # BARRA ESTUDIANTE (Azul Real)
        c.setFillColor(colors.royalblue)
        c.rect(x_base, y + 35, bar_width, h_est, fill=1, stroke=0)
        
        # Valor sobre barra Estudiante
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x_base + bar_width/2, y + 38 + h_est, str(p_est))
        
        # BARRA GRUPO (Naranja Vibrante)
        c.setFillColor(colors.HexColor("#DD712A"))
        c.rect(x_base + bar_width, y + 35, bar_width, h_grp, fill=1, stroke=0)
        
        # Valor sobre barra Grupo
        c.drawCentredString(x_base + bar_width + bar_width/2, y + 38 + h_grp, str(p_grp))
        
        # ETIQUETAS DE MATERIA (4 LETRAS)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x_base + bar_width, y + 18, etiqueta_visual)
        
    # 4. LEYENDA EN LA BASE
    leyenda_y = y + 2
    c.setFillColor(colors.royalblue)
    c.rect(x + padding, leyenda_y, 7, 7, fill=1)
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8)
    c.drawString(x + padding + 12, leyenda_y + 1, "Estudiante")
    
    c.setFillColor(colors.orangered)
    c.rect(x + padding + 100, leyenda_y, 7, 7, fill=1)
    c.setFillColor(colors.black)
    c.drawString(x + padding + 112, leyenda_y + 1, "Promedio Grupo")


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
    
    # --- ORDEN ESPECÍFICO Y EXCEPCIÓN DE CIENCIAS ---
    # Definimos el orden exacto solicitado omitiendo "Ciencias"
    orden_deseado = [
        "Lectura crítica",
        "Matemáticas",
        "Biología",
        "Química",
        "Física",
        "C. Sociales",
        "Inglés"
    ]
    
    # Filtramos la lista basándonos en la disponibilidad en metadata
    areas_principales = [area for area in orden_deseado if area in metadata["areas_principales_calificadas"]]
    
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
    # 0. LOGOS (LOGO Y AVANZA EDITORES)
    # -----------------------------------------------------------------------
    try:
        # Se elimina la línea de "logo.png" (el de la izquierda)
        # Se ajusta la posición X para que "avanza_editores.png" quede en el centro
        ancho_logo = 6.8 * inch
        posicion_x_centro = (ancho - ancho_logo) / 2.5
        
        c.drawImage("imagenes/REPORTE_INDIVIDUAL.png", 
                    posicion_x_centro, 
                    alto - 1.8 * inch, 
                    width=ancho_logo, 
                    height=1.5 * inch, 
                    preserveAspectRatio=True)
    except:
        pass

    # -----------------------------------------------------------------------
    # 1. CABECERA IZQUIERDA (DATOS INSTITUCIÓN Y PERSONALES)
    # -----------------------------------------------------------------------
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
    c.setFont("Courier", 11)
    c.drawString(x_start_izq + 0.1 * inch, text_y - 15, nombre.upper())
    
    # Ajuste de fuente para asegurar que la Institución quepa correctamente
    campos_izq = [
        ("Fecha", fecha.upper()),
        ("Municipio", ciudad.replace('_', ' ').title()),
        ("Institución", colegio.upper()),
        ("Grado", grado),
    ]
    
    current_y_campo = text_y - 35 
    for label, value in campos_izq:
        c.setFont("Courier-Bold", 11)
        c.drawString(x_start_izq + 0.1 * inch, current_y_campo, f"{label}:")
        c.setFont("Courier", 9) # Fuente ajustada para nombres largos
        c.drawString(x_start_izq + 1.3 * inch, current_y_campo, str(value))
        c.setLineWidth(0.5)
        c.line(x_start_izq + 1.3 * inch, current_y_campo - 2, x_start_izq + ancho_izq - 0.1 * inch, current_y_campo - 2)
        current_y_campo -= 14

    # RECUADRO DEL PUESTO (Asegurando alineación y visibilidad)
    puesto_y = y_start_izq - alto_izq + 0.2 * inch
    c.setFillColor(colors.lightblue)
    c.rect(x_start_izq, puesto_y, ancho_izq * 0.7, 0.4 * inch, fill=1)
    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 10)
    c.drawString(x_start_izq + 0.1 * inch, puesto_y + 0.15 * inch, "Puesto que ocupó en el grupo")
    
    c.setFillColor(colors.white)
    c.rect(x_start_izq + ancho_izq * 0.7, puesto_y, ancho_izq * 0.3, 0.4 * inch, fill=1)
    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 14)
    c.drawCentredString(x_start_izq + ancho_izq * 0.85, puesto_y + 0.12 * inch, str(puesto))
    
    c.setLineWidth(2)
    c.line(x_start_izq + ancho_izq * 0.7, puesto_y, x_start_izq + ancho_izq * 0.7, puesto_y + 0.4 * inch)

    # -----------------------------------------------------------------------
    # 2. CABECERA DERECHA (PUNTAJE GLOBAL Y PERCENTIL)
    # -----------------------------------------------------------------------
    # Definición de dimensiones y posición del recuadro derecho
    x_start_der = x_start_izq + ancho_izq + 0.1 * inch 
    y_start_der = alto - 120
    ancho_der = 3.0 * inch
    alto_der = 147 
    
    # Dibujo del marco exterior derecho
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    c.rect(x_start_der, y_start_der - alto_der, ancho_der, alto_der, fill=0)

    # --- SECCIÓN SUPERIOR: NOMBRE DEL EXAMEN Y TÍTULO ---
    puntaje_y = y_start_der - 30
    c.setFillColor(colors.lightblue)
    c.rect(x_start_der, puntaje_y, ancho_der, 30, fill=1)
    
    c.setFillColor(colors.black)
    c.setFont("Courier", 10)
    # Aquí aparece el nombre del examen que faltaba
    c.drawCentredString(x_start_der + ancho_der / 2, puntaje_y + 20, f"Examen {examen}")
    
    c.setFont("Courier-Bold", 11)
    c.drawCentredString(x_start_der + ancho_der / 2, puntaje_y + 7, "PUNTAJE GLOBAL")
    
    # --- SECCIÓN MEDIA: VALOR DEL PUNTAJE GLOBAL ---
    c.setFont("Courier-Bold", 30)
    puntaje_texto = str(int(puntaje_global))
    # Ajuste para centrar el puntaje con respecto a la escala (ej: 350 / 500)
    c.drawCentredString(x_start_der + ancho_der / 2 - 15, puntaje_y - 30, puntaje_texto)
    
    c.setFillColor(colors.orangered)
    c.setFont("Courier-Bold", 15) # Escala un poco más pequeña que el puntaje
    c.drawString(x_start_der + ancho_der / 2 + 10, puntaje_y - 30, f" /{escala}")

    # --- SECCIÓN INFERIOR: PERCENTIL ---
    percentil_y = puntaje_y - 60
    c.setFillColor(colors.lightblue)
    c.rect(x_start_der, percentil_y, ancho_der, 20, fill=1)
    
    c.setFillColor(colors.black)
    c.setFont("Courier-Bold", 10)
    c.drawCentredString(x_start_der + ancho_der / 2, percentil_y + 5, "¿EN QUÉ PERCENTIL ME ENCUENTRO?")
    
    # Texto explicativo debajo del título de percentil
    c.setFont("Courier", 7)
    c.drawCentredString(x_start_der + ancho_der / 2, percentil_y - 12, "con respecto a los estudiantes del grupo")
    
    # Dibujo de la barra de percentil gráfica
    barra_x = x_start_der + (ancho_der - 2.5 * inch) / 2 
    generar_percentil_barra(c, barra_x, percentil_y - 50, percentil_global, ancho=2.5 * inch)
    
    
    # -----------------------------------------------------------------------
    # 3. TABLA DE RESULTADOS POR ÁREA (MÁS COMPACTA)
    # -----------------------------------------------------------------------
    current_y_end_header = y_start_izq - alto_izq - 15 
    c.setFont("Helvetica-Bold", 12) # Fuente de título un poco más pequeña
    c.drawCentredString(ancho / 2, current_y_end_header, "RESULTADOS POR ÁREA")
    
    start_y = current_y_end_header - 25
    row_height = 0.35 * inch # REDUCIDO: Antes era 0.45, esto ahorra mucho espacio
    col_width_tabla = 6.8 * inch 
    col_area_x_start = inch 
    col_puntaje_x_start = col_area_x_start + 1.9 * inch 
    col_nivel_x_start = col_puntaje_x_start + 0.6 * inch 
    col_grafica_x_start = col_nivel_x_start + 1.2 * inch 
    
    # Cabecera de tabla
    c.setFillColor(colors.skyblue)
    c.rect(inch, start_y, col_width_tabla, 0.20 * inch, fill=1)
    c.setFillColor(colors.black) 
    c.setFont("Helvetica-Bold", 9) # Fuente de cabecera reducida
    c.drawString(col_area_x_start + 0.1 * inch, start_y + 0.05 * inch, "ÁREA")
    c.drawCentredString(col_puntaje_x_start + 0.3 * inch, start_y + 0.05 * inch, "%")
    c.drawCentredString(col_nivel_x_start + 0.6 * inch, start_y + 0.05 * inch, "NIVEL")
    
    current_y = start_y
    for area in areas_principales:
        current_y -= row_height 
        puntaje_val = estudiante.get(f"Puntaje_{area}", 0) 
        puntaje_int = int(puntaje_val)
        nivel = obtener_nivel(puntaje_int, area)
            
        c.setStrokeColor(colors.black)
        c.rect(inch, current_y, col_width_tabla, row_height, fill=0)
        
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 8) # REDUCIDO: Fuente de datos a 8 para que sea compacto
        c.drawString(col_area_x_start + 0.1 * inch, current_y + row_height / 2 - 2, area.upper())
        c.drawCentredString(col_puntaje_x_start + 0.3 * inch, current_y + row_height / 2 - 2, str(puntaje_int))
        c.drawCentredString(col_nivel_x_start + 0.6 * inch, current_y + row_height / 2 - 2, nivel.upper())
        
        # Barra de percentil más delgada para la tabla compacta
        generar_percentil_barra(c, col_grafica_x_start + 0.2 * inch, current_y + 8, puntaje_int, ancho=2.5 * inch, alto=0.15 * inch)
    
    # -----------------------------------------------------------------------
    # 4. GRÁFICO DE BARRAS COMPARATIVO
    # -----------------------------------------------------------------------
    grafico_y_start = current_y - 0.15 * inch
    dibujar_grafico_barras_comparacion(
        c, inch, grafico_y_start - 2.4 * inch, # Altura de gráfico ligeramente reducida
        estudiante, promedios_grupo, areas_principales, 
        width=6.5 * inch, height=2.3 * inch
    )

    # -----------------------------------------------------------------------
    # 5. IMÁGENES AL PIE DE PÁGINA (REPETIDAS ABAJO)
    # -----------------------------------------------------------------------
    try:
        # Usamos los mismos valores que configuraste arriba para mantener la simetría
        ancho_logo_inferior = 6.5 * inch
        # Misma fórmula de posición que usaste arriba
        posicion_x_inferior = (ancho - ancho_logo_inferior) / 2.5
        
        # Dibujamos el logo en la parte inferior (margin + 0.2 inch para que no toque el borde)
        c.drawImage("imagenes/pie_pagina.png", 
                    posicion_x_inferior, 
                    margin + 0.2 * inch, 
                    width=ancho_logo_inferior, 
                    height=1.0 * inch, # Un poco más bajo para que no quite espacio al gráfico
                    preserveAspectRatio=True)
    except:
        pass
    
    
    
    
    

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
        
class ComparacionBarras2(Flowable):
    """Gráfico para Cantidades: TOTAL (Azul) vs SOBRE 261 (Naranja)"""
    def __init__(self, data, width=7 * inch, height=3 * inch):
        Flowable.__init__(self)
        self.data = data 
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        areas_list = list(self.data.keys())
        x_start, y_start = 0.5 * inch, 0.5 * inch
        plot_width, plot_height = self.width - 1*inch, self.height - 1.2*inch
        
        # --- LEYENDA ---
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.royalblue)
        c.rect(x_start, self.height - 0.3*inch, 8, 8, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_start + 12, self.height - 0.28*inch, "Total Estudiantes")
        
        c.setFillColor(colors.orange)
        c.rect(x_start + 100, self.height - 0.3*inch, 8, 8, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_start + 112, self.height - 0.28*inch, "Cumplen Meta")

        # Eje Y automático
        max_val = max([v['Grupo'] for v in self.data.values()]) if self.data else 10
        escala_y = int(np.ceil(max_val / 10.0)) * 10
        
        c.setFont("Helvetica", 7)
        for i in range(0, escala_y + 1, max(1, escala_y // 5)):
            y_pos = y_start + plot_height * (i / escala_y)
            c.setStrokeColor(colors.lightgrey)
            c.line(x_start, y_pos, x_start + plot_width, y_pos)
            c.drawRightString(x_start - 5, y_pos - 2, str(i))

        # Dibujo de Barras
        for i, label in enumerate(areas_list):
            val = int(self.data[label]['Grupo'])
            center_x = x_start + (plot_width/len(areas_list))*(i+0.5)
            bar_h = plot_height * (val/escala_y)
            
            # Color: Azul para TOTAL, Naranja para el resto
            c.setFillColor(colors.royalblue if "TOTAL" in label.upper() else colors.orange)
            c.rect(center_x - 15, y_start, 30, bar_h, fill=1, stroke=0)
            
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(center_x, y_start + bar_h + 3, str(val))
            c.drawCentredString(center_x, y_start - 12, label)
            
class ComparacionBarras3(Flowable):
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

class ComparacionGlobal500(Flowable):
    """Gráfico para Puntajes: META (Azul) vs GRUPO (Naranja)"""
    def __init__(self, data, width=7 * inch, height=3 * inch):
        Flowable.__init__(self)
        self.data = data 
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        x_start, y_start = 0.5 * inch, 0.5 * inch
        plot_width, plot_height = self.width - 1*inch, self.height - 1.2*inch
        
        # --- LEYENDA ---
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.royalblue)
        c.rect(x_start, self.height - 0.3*inch, 8, 8, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_start + 12, self.height - 0.28*inch, "Meta (Referente)")
        
        c.setFillColor(colors.orange)
        c.rect(x_start + 100, self.height - 0.3*inch, 8, 8, fill=1)
        c.setFillColor(colors.black)
        c.drawString(x_start + 112, self.height - 0.28*inch, "Promedio Grupo")

        # Eje Y 0-500
        c.setFont("Helvetica", 7)
        for i in range(0, 501, 100):
            y_pos = y_start + plot_height * (i / 500.0)
            c.setStrokeColor(colors.lightgrey)
            c.line(x_start, y_pos, x_start + plot_width, y_pos)
            c.drawRightString(x_start - 5, y_pos - 2, str(i))

        areas = list(self.data.keys())
        for i, label in enumerate(areas):
            val = int(self.data[label]['Grupo'])
            center_x = x_start + (plot_width/len(areas))*(i+0.5)
            bar_h = plot_height * (val/500.0)
            
            # Color: Azul para META, Naranja para GRUPO
            c.setFillColor(colors.royalblue if "META" in label.upper() else colors.orange)
            c.rect(center_x - 20, y_start, 40, bar_h, fill=1, stroke=0)
            
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(center_x, y_start + bar_h + 3, str(val))
            c.drawCentredString(center_x, y_start - 12, label)

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

def background_estetica(canvas, doc):
    canvas.saveState()
    ancho, alto = letter
    
    # -----------------------------------------------------------------------
    # 1. ENCABEZADO (Logo de arriba)
    # -----------------------------------------------------------------------
    try:
        # Tus coordenadas y medidas exactas
        ancho_logo_sup = 6.8 * inch
        posicion_x = (ancho - ancho_logo_sup) / 2.5
        
        canvas.drawImage("imagenes/avanza_editores.png", 
                         posicion_x, 
                         alto - 1.8 * inch, 
                         width=ancho_logo_sup, 
                         height=1.5 * inch, 
                         preserveAspectRatio=True)
    except:
        pass

    
        
    canvas.restoreState()

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
            
            # --- NUEVA LÓGICA DE PUESTOS ÚNICOS ---
            # 1. Ordenamos por Puntaje_Global de mayor a menor
            self.data["resultados_estudiantes"].sort(
                key=lambda x: x.get("Puntaje_Global", 0), 
                reverse=True
            )
            
            # 2. Asignamos un nuevo puesto secuencial del 1 al N
            # Esto rompe los empates por orden de aparición (aleatorio/arbitrario)
            for i, estudiante in enumerate(self.data["resultados_estudiantes"], start=1):
                estudiante["Puesto"] = i
            # ----------------------------------------
            
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
        if not self.archivo_json:
            messagebox.showerror("Error", "Primero debe cargar un archivo JSON.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile="Reporte_General_Detallado.pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_path:
            return

        try:
            from reportlab.platypus import PageBreak
            metadata = self.data["metadata"]
            resultados = self.data["resultados_estudiantes"]
            
            # --- 1. CONFIGURACIÓN DE MATERIAS ---
            orden_materias = ["Lectura crítica", "Matemáticas", "Biología", "Química", "Física", "C. Sociales", "Inglés"]
            areas_visibles = [a for a in orden_materias if a in metadata["areas_principales_calificadas"]]

            doc = SimpleDocTemplate(output_path, pagesize=letter, 
                                    leftMargin=0.5*inch, rightMargin=0.5*inch, 
                                    topMargin=1.8*inch, bottomMargin=0.8*inch)
            flowables = []
            styles = getSampleStyleSheet()

            # --- 2. TÍTULO Y CABECERA (ESTILO IMAGEN) ---
            estilo_titulo = styles['Heading1']
            estilo_titulo.alignment = 1 # Centrado
            estilo_titulo.textColor = colors.navy
            flowables.append(Paragraph("INFORME GENERAL DE RESULTADOS", estilo_titulo))
            flowables.append(Spacer(1, 0.1 * inch))

            estilo_label = styles['Normal']
            estilo_label.fontName = "Helvetica-Bold"
            estilo_label.fontSize = 9

            # Definición de datos incluyendo el GRADO
            datos_cabecera = [
                [
                    Paragraph(f"INSTITUCIÓN: <font face='Helvetica'>{metadata.get('colegio', '').upper()}</font>", estilo_label),
                    Paragraph(f"MUNICIPIO: <font face='Helvetica'>{metadata.get('ciudad', '').replace('_', ' ').upper()}</font>", estilo_label)
                ],
                [
                    Paragraph(f"EVALUACIÓN: <font face='Helvetica'>{metadata.get('examen', '').upper()}</font>", estilo_label),
                    Paragraph(f"GRADO: <font face='Helvetica'>{metadata.get('grado', 'N/A').upper()}</font>", estilo_label)
                ],
                [
                    Paragraph(f"FECHA: <font face='Helvetica'>{metadata.get('fecha', '').upper()}</font>", estilo_label),
                    "" # Celda vacía para completar la fila
                ]
            ]

            
            flowables.append(Spacer(1, 0.3 * inch))

            tabla_header = Table(datos_cabecera, colWidths=[4.0*inch, 3.2*inch])
            tabla_header.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0)]))
            flowables.append(tabla_header)
            flowables.append(Spacer(1, 0.2 * inch))

            # --- 3. CÁLCULOS PARA LA TABLA Y ANÁLISIS ---
            header = ["Pto", "Nombre Estudiante"] + [a[:4].upper() for a in areas_visibles] + ["GLO"]
            tabla_data = [header]

            sumas_areas = {area: 0 for area in areas_visibles}
            suma_global = 0
            conteo = len(resultados)

            for est in resultados:
                fila = [str(est.get("Puesto", "-")), est.get("Nombre", "N/A")[:25].upper()]
                for area in areas_visibles:
                    val = est.get(f"Puntaje_{area}", 0)
                    fila.append(str(int(val)))
                    sumas_areas[area] += val
                p_glo = int(est.get("Puntaje_Global", 0))
                suma_global += p_glo
                # Creamos el párrafo con etiquetas <b>
                celda_glo_bold = Paragraph(f"<b>{p_glo}</b>", estilo_label)
                fila.append(celda_glo_bold)
                
                tabla_data.append(fila)

            # Fila de Promedios
            promedios = {area: sumas_areas[area]/conteo for area in areas_visibles}
            fila_promedio = ["-", "PROMEDIO GRUPO"]
            for area in areas_visibles:
                fila_promedio.append(f"{promedios[area]:.1f}")
            
            # --- CAMBIO AQUÍ: Promedio Global en Negrita ---
            prom_glo_final = f"<b>{suma_global/conteo:.1f}</b>"
            fila_promedio.append(Paragraph(prom_glo_final, estilo_label))
            tabla_data.append(fila_promedio)

            t_res = Table(tabla_data, colWidths=[0.4*inch, 2.3*inch] + [0.45*inch]*(len(areas_visibles)+1), repeatRows=1)
            t_res.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.skyblue),
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                # --- NUEVAS LÍNEAS DE ESTILO ---
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),   # Encabezados bold
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'), # Fila promedio bold
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),             # Centrado vertical
                # -------------------------------
                ('ALIGN', (2,0), (-1,-1), 'CENTER')
            ]))
            flowables.append(t_res)
            
            #promedio graficado
            flowables.append(PageBreak())
            
            estilo_sub = styles['Heading2']
            estilo_sub.alignment = 1 # Centrado
            flowables.append(Paragraph("PROMEDIOS GENERALES POR ÁREA", estilo_sub))
            #flowables.append(Spacer(1, 0.2 * inch))

            # Preparamos los datos para TODAS las áreas visibles
            # 'Estudiante' se envía en 0 para que solo se grafique el promedio del grupo
            data_todas_areas = {
                #area: {'Estudiante': 0, 'Grupo': promedios[area]}
                area[:4].upper(): {'Estudiante': 0, 'Grupo': promedios[area]} 
                for area in areas_visibles
            }

            # Insertamos el gráfico con todas las materias
            flowables.append(ComparacionBarras3(data_todas_areas, width=7.2*inch, height=3.5*inch))
            
            flowables.append(Spacer(1, 0.2 * inch))
            #texto_resumen = "El gráfico anterior muestra el desempeño promedio del grupo en cada una de las dimensiones evaluadas."
            #flowables.append(Paragraph(texto_resumen, styles['Normal']))
            
            # (Aquí puede continuar tu lógica de Extremos: MÁX y MÍN si deseas conservarla)
            #flowables.append(Spacer(1, 0.4 * inch))

            # --- 4. NUEVA SECCIÓN: COMPARATIVA EXTREMOS ---
            #flowables.append(PageBreak())
            
            # Identificar materia más alta y más baja
            materia_top = max(promedios, key=promedios.get)
            materia_low = min(promedios, key=promedios.get)

            estilo_sub = styles['Heading2']
            estilo_sub.alignment = 1 # Centrado
            flowables.append(Paragraph("ANÁLISIS COMPARATIVO DE RENDIMIENTO", estilo_sub))
            flowables.append(Spacer(1, 0.2 * inch))

            # Gráfico de comparación de promedios extremos
            data_extremos = {
                f"MÁX: {materia_top}": {'Estudiante': 0, 'Grupo': promedios[materia_top]},
                f"MÍN: {materia_low}": {'Estudiante': 0, 'Grupo': promedios[materia_low]}
            }
            # Reutilizamos la lógica de ComparacionBarras pero solo con promedios de grupo
            flowables.append(ComparacionBarras3(data_extremos, width=7*inch, height=2.5*inch))
            
            # Texto explicativo
            texto_analisis = f"""Este gráfico compara el área con mejor desempeño <b>({materia_top.upper()}: {promedios[materia_top]:.1f})</b> 
            frente al área con mayores retos <b>({materia_low.upper()}: {promedios[materia_low]:.1f})</b>. 
            La brecha detectada es de <b>{promedios[materia_top] - promedios[materia_low]:.1f}</b> puntos."""
            flowables.append(Paragraph(texto_analisis, styles['Normal']))

            # --- 5. PROMEDIO GLOBAL ---
            flowables.append(PageBreak())
            
            flowables.append(Paragraph("ANÁLISIS DE REFERENCIA NACIONAL (261)", estilo_sub))
            flowables.append(Spacer(1, 0.4 * inch))

            # Gráfico 1: Estudiantes
            est_sobre = len([e for e in resultados if e.get("Puntaje_Global", 0) >= 261])
            data_meta_est = {
                "TOTAL": {'Estudiante': 0, 'Grupo': len(resultados)},
                "SOBRE 261": {'Estudiante': 0, 'Grupo': est_sobre}
            }
            flowables.append(Paragraph("Alumnos que superan la media nacional", estilo_label))
            flowables.append(ComparacionBarras2(data_meta_est, width=7.2*inch, height=3.0*inch))
            
            flowables.append(Spacer(1, 0.6 * inch))

            # Gráfico 2: Puntaje Global (Escala 500)
            prom_glo = suma_global / len(resultados)
            data_meta_glo = {
                "META ICFES": {'Estudiante': 0, 'Grupo': 261},
                "PROM. GRUPO": {'Estudiante': 0, 'Grupo': prom_glo}
            }
            flowables.append(Paragraph("Promedio nacional y del grupo", estilo_label))
            flowables.append(ComparacionGlobal500(data_meta_glo, width=7.2*inch, height=3.2*inch))
            
            # 6. GRAFICO POR MATERIA
            flowables.append(Spacer(1, 0.4 * inch))
            flowables.append(Paragraph("DISTRIBUCIÓN POR NIVELES DE DESEMPEÑO", estilo_sub))
            flowables.append(Spacer(1, 0.2 * inch))
            
            conteo_por_area = contar_estudiantes_por_nivel(resultados, areas_visibles)
            for area in areas_visibles:
                flowables.append(DistribucionNivelesBarra(area, conteo_por_area[area], width=7*inch, height=2.5*inch))
                flowables.append(Spacer(1, 0.3 * inch))

            doc.build(flowables, onFirstPage=background_estetica, onLaterPages=background_estetica)
            messagebox.showinfo("Éxito", "Reporte General generado con análisis comparativo.")

        except Exception as e:
            messagebox.showerror("Error", f"Error crítico: {e}")


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