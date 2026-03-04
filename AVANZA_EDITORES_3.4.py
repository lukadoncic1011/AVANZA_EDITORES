import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk 
from PIL import Image, ImageTk
import json
import os
import sys
import subprocess
# from pandas import ExcelWriter # Ya no se usa ExcelWriter

#RANGO DEL BANCO DE PREGUNTA
LEC = 25
MAT = 50
BIO = 60
QUI = 70
FIS = 80
SOC = 100
ING = 120

PROMEDI = 7


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


# --- COLORES PARA ESTÉTICA (MANTENIDOS) ---
COLOR_FONDO_CIELO = "#E0FFFF" # Azul Cielo
COLOR_AZUL_REY = "#0D1096"   # Azul Rey
COLOR_TEXTO_PRINCIPAL = "#004D99" # Azul Oscuro
COLOR_NARANJA = "#D87F4B" # NARANJA para encabezados de tabla
# ------------------------------------------
def ejecutar_generador_reportes():
    """Ejecuta el archivo generador_reportes.exe en la subcarpeta especificada."""
    try:
        # 1. Definir la ruta relativa hacia la subcarpeta y el ejecutable
        # Usamos resource_path para que funcione tanto en script como en el ejecutable final
        ruta_ejecutable = resource_path(os.path.join("generador_reportes", "generador_reportes.exe"))
        
        # 2. Verificar si el archivo existe antes de intentar abrirlo
        if os.path.exists(ruta_ejecutable):
            # Ejecutamos el .exe directamente
            subprocess.Popen([ruta_ejecutable])
        else:
            messagebox.showerror("Archivo no encontrado", 
                                 f"No se encontró el archivo en:\n{ruta_ejecutable}")
            
    except Exception as e:
        # Muestra un mensaje de error si no se puede ejecutar
        messagebox.showerror("Error de Ejecución", f"No se pudo iniciar el generador: {e}")


# --- CONFIGURACIÓN DE LAS MATERIAS (Ahora solo lista de potenciales) ---
# Se mantiene la lista de todas las materias posibles, pero los rangos serán dinámicos
MATERIAS_POTENCIALES = {
        'Lectura crítica': (1, LEC),
        'Matemáticas': (LEC + 1, MAT),
        'Biología': (MAT + 1, BIO),
        'Química': (BIO + 1, QUI),
        'Física': (QUI + 1, FIS),
        'C. Sociales': (FIS + 1, SOC),
        'Inglés': (SOC + 1, ING)
    }


# Las 5 áreas principales que se sumarán al puntaje global (Máx 500)
AREAS_PRINCIPALES = [
        'Lectura crítica', 
        'Matemáticas', 
        'Biología', 
        'Química', 
        'Física', 
        'C. Sociales',
        'Inglés'
    ]

# --- Función de utilidad para cálculos (fuera de las clases) ---
# --- MODIFICACIÓN: ESCALA GLOBAL - FUNCIÓN PREPARAR DATOS ---
def preparar_datos_estadisticos(df_calificaciones, rangos_calificados, escala_maxima): 
    """
    Calcula el puntaje global (máx 500, o escala_maxima si se pasa), el ranking, 
    consolida Ciencias Naturales y escala el puntaje global basado en el número 
    de áreas principales calificadas (1 a 5).
    
    NOTA: Asume que las columnas 'Max_Aciertos_Materia' han sido creadas en AppLogica.calificar
    """
    
    df = df_calificaciones.copy()
    
    # Lista para llevar el registro de las áreas principales que realmente se calificaron
    materias_principal_calificadas = [] 

    # --- 1. LÓGICA DE CIENCIAS NATURALES (Consolidación) ---
    
    # Calcular el máximo de aciertos REAL de Ciencias Naturales (Suma de los Max_Aciertos de sus sub-áreas)
    MAX_CIENCIAS_NATURALES_REAL = 0
    hay_sub_ciencias_calificadas = False
    for sub_area in ['Biología', 'Química', 'Física']:
        max_col = f'Max_Aciertos_{sub_area}'
        if max_col in df.columns and df[max_col].max() > 0:
            # Usamos el valor máximo de la columna Max_Aciertos_{sub_area}
            MAX_CIENCIAS_NATURALES_REAL += df[max_col].max() 
            hay_sub_ciencias_calificadas = True
            
    # Sumar los aciertos individuales en 'Aciertos_Ciencias_naturales'
    aciertos_cn = 0
    if 'Aciertos_Biología' in df.columns: aciertos_cn += df['Aciertos_Biología']
    if 'Aciertos_Química' in df.columns: aciertos_cn += df['Aciertos_Química']
    if 'Aciertos_Física' in df.columns: aciertos_cn += df['Aciertos_Física']
    df['Aciertos_Ciencias'] = aciertos_cn
        
    if MAX_CIENCIAS_NATURALES_REAL > 0:
         # Si se calificó alguna sub-área, la principal 'Ciencias_naturales' participa
         materias_principal_calificadas.append('Ciencias')
    
    # 2. Normalización a Puntaje Individual (Máx 100)
    
    for mat in AREAS_PRINCIPALES:
        col_aciertos = f'Aciertos_{mat}'
        col_puntaje = f'Puntaje_{mat}'
        max_col = f'Max_Aciertos_{mat}'
        
        # 2a. Determinar puntos máximos de aciertos para la materia
        if mat == 'Ciencias':
            max_aciertos = MAX_CIENCIAS_NATURALES_REAL
            if max_aciertos > 0:
                 df[max_col] = max_aciertos # Crear columna de max aciertos para CN

        elif mat in rangos_calificados:
            max_aciertos = df[max_col].max() if max_col in df.columns else 0
            if max_aciertos > 0:
                 materias_principal_calificadas.append(mat)
        else:
            max_aciertos = 0
            
        # 2b. Normalización (Aciertos / Max_Aciertos) * 100
        if col_aciertos in df.columns and max_aciertos > 0:
            df[col_puntaje] = (df[col_aciertos] / max_aciertos) * 100
            # Clip para asegurar que no se pase de 100 por posibles errores de redondeo
            df[col_puntaje] = df[col_puntaje].clip(upper=100) 
        else:
             df[col_puntaje] = 0
        
        # 2d. Asegurar que las 5 áreas principales existan (con 0 si no fueron calificadas)
        if col_puntaje not in df.columns:
             df[col_puntaje] = 0

    
    # 3. Calcular el PUNTAJE GLOBAL (Escalado dinámicamente)
    
    # Columnas de puntaje que se usarán para la suma RAW (siempre las 5 áreas principales)
    final_puntaje_cols = [f'Puntaje_{mat}' for mat in AREAS_PRINCIPALES]
    
    # Calcular el Puntaje Global RAW (Suma de los puntajes de 0 a 100 de las 5 áreas)
    df['Puntaje_Global_Raw'] = df[final_puntaje_cols].sum(axis=1)

    # N: Número de áreas principales que realmente contribuyeron al puntaje
    # Usamos set() para asegurar unicidad
    #N = len(list(set(materias_principal_calificadas))) 
    N = PROMEDI

    if N > 0:
        # Puntos máximos posibles RAW (N * 100)
        MAX_PUNTOS_POSIBLES_RAW = N * 100
        
        # Fórmula de Escalado: Puntaje_Global = (Raw Score / Max_Raw_Score) * escala_maxima
        df['Puntaje_Global'] = (df['Puntaje_Global_Raw'] / MAX_PUNTOS_POSIBLES_RAW) * escala_maxima
        
        # Asegurar que el puntaje no exceda escala_maxima
        df['Puntaje_Global'] = df['Puntaje_Global'].clip(upper=escala_maxima)
        
    else:
        df['Puntaje_Global'] = 0

    # 4. Cálculos Estadísticos Finales
    df = df.drop(columns=['Puntaje_Global_Raw'])
    # Eliminar las columnas Max_Aciertos_* que solo fueron temporales (excepto para CN)
    cols_to_drop = [col for col in df.columns if col.startswith('Max_Aciertos_')]
    df = df.drop(columns=cols_to_drop, errors='ignore')
    
    df['Puesto'] = df['Puntaje_Global'].rank(ascending=False, method='min').astype(int)
    total_estudiantes = len(df)
    df['Percentil_Global'] = ((total_estudiantes - df['Puesto']) / total_estudiantes) * 100
    df['Percentil_Global'] = df['Percentil_Global'].apply(lambda x: min(99, max(1, round(x)))) 
    
    # 5. Redondeo del Puntaje Global a 2 decimales
    df['Puntaje_Global'] = df['Puntaje_Global'].round(2)
    
    return df, list(set(materias_principal_calificadas))


# ------------------------------------
# --- CLASE 1: VENTANA PRINCIPAL (MENÚ) ---
class CalificadorApp:
    
    
    
    def __init__(self, master):
        
        
        self.master = master
        master.title("Menú Principal del Calificador y Reporte")
        master.geometry("720x720") 
        master.config(bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        
        # --- ESTADO CENTRALIZADO ---
        self.sello_img = None # <--- NUEVA REFERENCIA DE IMAGEN 
        self.df_respuestas = None         
        self.respuestas_correctas = None  
        self.df_calificaciones = None     
        self.metadata = {}                
        self._temp_df_resultados = None   
        self._temp_metadata = None  
        self.banner_img = None 
             
        # ---------------------------

        self._create_widgets()

    def _create_widgets(self):
        """Crea los widgets en la ventana principal."""
        tk.Label(self.master, text="Selecciona la acción a realizar:", font=("Arial", 14, "bold"), fg=COLOR_TEXTO_PRINCIPAL, bg=COLOR_FONDO_CIELO).pack(pady=10) # ✅ FONDO AZUL CIELO
        
        # === PRIMERA IMAGEN: SELLO DE CALIDAD (Parte Superior) ===
        SELLO_PATH = resource_path(os.path.join("imagenes", "logo.png"))
        
        
        TARGET_WIDTH_SELLO = 100 # Ancho deseado más pequeño

        try:
            pil_image_sello = Image.open(SELLO_PATH)
            original_width, original_height = pil_image_sello.size
            new_height_sello = int(original_height * (TARGET_WIDTH_SELLO / original_width))
            resized_image_sello = pil_image_sello.resize((TARGET_WIDTH_SELLO, new_height_sello), Image.Resampling.LANCZOS)
            
            self.sello_img = ImageTk.PhotoImage(resized_image_sello) # <--- Guardar Referencia
            
            sello_label = tk.Label(self.master, 
                                    image=self.sello_img, 
                                    bg=COLOR_FONDO_CIELO)
            
            # Usar pack para colocar el sello en la parte superior (antes del título)
            sello_label.pack(side=tk.TOP, pady=5) 
            
        except Exception as e:
            print(f"ADVERTENCIA: No se pudo cargar el sello ('sello.png'): {e}")
     
        # === SEGUNDA IMAGEN: BANNER (Parte Inferior) ===
        BANNER_PATH = resource_path(os.path.join("imagenes", "avanza_editores.png"))
        TARGET_WIDTH = 500 

        try:
            pil_image = Image.open(BANNER_PATH)
            original_width, original_height = pil_image.size
            if original_width > TARGET_WIDTH:
                new_height = int(original_height * (TARGET_WIDTH / original_width))
            else:
                 new_height = original_height 
                 TARGET_WIDTH = original_width 
                 
            resized_image = pil_image.resize((TARGET_WIDTH, new_height), Image.Resampling.LANCZOS)
            self.banner_img = ImageTk.PhotoImage(resized_image) # <--- Guardar Referencia
            
            banner_label = tk.Label(self.master, 
                                    image=self.banner_img, 
                                    bg=COLOR_FONDO_CIELO)
            
            # Usar pack para colocar el banner en la parte más inferior
            banner_label.pack(side=tk.BOTTOM, pady=10) 
            
        except Exception as e:
            print(f"ADVERTENCIA: No se pudo cargar el archivo 'banner.png': {e}")
        # =======================================================
        
        

        # Asegurar que el tamaño inicial de la ventana es suficiente para el contenido
        self.master.geometry("720x450")
        
        frame_botones = tk.Frame(self.master, bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        frame_botones.pack(pady=10)
        # Botón 1: INICIAR CALIFICACIÓN (Abre AppLogica)
        self.iniciar_boton = tk.Button(frame_botones, 
                                       text="1. INICIAR CALIFICACIÓN (Ingresar Datos y Generar Resultados)", 
                                       command=self.abrir_ventana_calificador,
                                       bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.iniciar_boton.pack(side=tk.TOP, padx=15, ipadx=10, ipady=5, pady=5) 

        # Botón 2: IMPORTAR ACIERTOS (Modificado)
        self.reporte_boton = tk.Button(frame_botones, 
                                       text="2. IMPORTAR ACIERTOS Y EXPORTAR REPORTES (Excel)", 
                                       command=self.abrir_ventana_importar, 
                                       bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.reporte_boton.pack(side=tk.TOP, padx=15, ipadx=10, ipady=5, pady=5) 

        # -----------------------------------------------------------------
        # --- NUEVO BOTÓN PARA GENERAR REPORTES (generador_reportes.exe) ---
        # -----------------------------------------------------------------
        self.boton_reportes = tk.Button(
            frame_botones, # El padre es frame_botones
            text="3. ABRIR GENERADOR DE REPORTES (generador_reportes.exe)", 
            command=ejecutar_generador_reportes, # Llama a la función definida anteriormente
            font=('Arial', 10, 'bold'),
            bg=COLOR_NARANJA, 
            fg='white',
            relief=tk.FLAT
        )
        # ESTA LÍNEA ES CLAVE: Coloca el botón en la ventana
        self.boton_reportes.pack(side=tk.TOP, padx=15, ipadx=10, ipady=5, pady=5)

        
        
        tk.Label(self.master, text="Nota: La carga de resultados y datos del examen ahora es un solo archivo JSON.", fg="gray", bg=COLOR_FONDO_CIELO).pack(pady=5) # ✅ FONDO AZUL CIELO


    # --- NUEVA FUNCIÓN PARA ABRIR LA VENTANA DE IMPORTACIÓN ---
    def abrir_ventana_importar(self):
        """Abre la ventana para cargar el Excel de aciertos directos."""
        ventana_importar = tk.Toplevel(self.master)
        ventana_importar.title("Paso 2: Importación de Aciertos")
        ventana_importar.config(bg=COLOR_FONDO_CIELO)
        AppImportarAciertos(ventana_importar, self, self.metadata)
        
        
    def abrir_ventana_calificador(self):
        
        
        
        
        """Abre la ventana para cargar archivos y calificar."""
        ventana_calificador = tk.Toplevel(self.master)
        ventana_calificador.title("Paso 1: Configuración y Calificación")
        ventana_calificador.config(bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        AppLogica(ventana_calificador, self, self.metadata)

    # --- MODIFICADO: Flujo de Reporte Individual (Carga de 1 archivo JSON) ---
    def iniciar_reporte_individual(self):
        """
        Inicia el flujo de reporte en un solo paso: pide el JSON consolidado.
        """
        ruta = filedialog.askopenfilename(
            defaultextension=".json", 
            filetypes=[("Archivos JSON", "*.json")], 
            title="Seleccionar Archivo de Resultados Consolidados (JSON)"
        )
        if ruta:
            try:
                with open(ruta, 'r', encoding='utf-8') as f:
                    consolidated_data = json.load(f)
                
                # 1. Extraer Metadata
                metadata = consolidated_data.get('metadata', {})
                
                # 2. Extraer Resultados y convertir a DataFrame
                resultados_list = consolidated_data.get('resultados_estudiantes', [])
                df_resultados = pd.DataFrame(resultados_list)
                
                # Chequeo de datos mínimos
                if df_resultados.empty or not all(k in metadata for k in ["colegio", "examen", "grado"]):
                     messagebox.showerror("Error de Archivo", "El archivo JSON está incompleto o malformado (faltan resultados o datos del examen).")
                     return
                
                # Carga completa, abrir ventana
                self._open_report_window(metadata, df_resultados)
                
            except json.JSONDecodeError:
                messagebox.showerror("Error de Carga", f"No se pudo cargar: El archivo no es un JSON válido.")
            except Exception as e:
                messagebox.showerror("Error de Carga", f"No se pudo cargar el archivo: {e}")

    def _open_report_window(self, metadata, df_resultados):
        """Función interna para abrir la ventana de reportes con la metadata y el DF cargados."""
        ventana_resultados = tk.Toplevel(self.master)
        ventana_resultados.title("Resultados Detallados por Estudiante y Área")
        ventana_resultados.config(bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        
        #ReportesApp(ventana_resultados, df_resultados, metadata)


# ---------------------------------------------
# --- CLASE 2: LÓGICA DE CARGA Y CALIFICACIÓN ---
# ---------------------------------------------
class AppLogica:
    def __init__(self, master, main_app, initial_metadata):
        self.master = master 
        self.main_app = main_app 
        
        self.df_respuestas = self.main_app.df_respuestas
        self.respuestas_correctas = self.main_app.respuestas_correctas
        self.metadata_vars = {} 
        self.initial_metadata = initial_metadata
        
        # --- NUEVOS ESTADOS PARA LA CALIFICACIÓN DINÁMICA ---
        self.rangos_calificar = {} # {Materia: {inicio, fin, multiplicador, extra}}
        self.materia_vars = {}     # Variables para los rangos de entrada
        # --- MODIFICACIÓN: ESCALA GLOBAL - Variable de Control ---
        self.escala_var = tk.IntVar(value=500) 
        # ----------------------------------------------------
        
        self.ruta_banco = tk.StringVar(value="Banco de Respuestas (No Cargado)")
        self.ruta_respuestas = tk.StringVar(value="Respuestas de Estudiantes (No Cargado)")
        if self.main_app.respuestas_correctas is not None:
             self.ruta_banco.set("Clave en memoria")
        if self.main_app.df_respuestas is not None:
             self.ruta_respuestas.set("Respuestas en memoria")
             
        # Inicialización del estado aquí para asegurar que exista (SOLUCIÓN al Error 1)
        self.estado = None 
             
        self.crear_widgets()

    def crear_widgets(self):
        
        BANNER_PATH = resource_path(os.path.join("imagenes", "avanza_editores.png"))
        TARGET_WIDTH = 450

        try:
            pil_image = Image.open(BANNER_PATH)
            original_width, original_height = pil_image.size
            if original_width > TARGET_WIDTH:
                new_height = int(original_height * (TARGET_WIDTH / original_width))
            else:
                 new_height = original_height 
                 TARGET_WIDTH = original_width 
                 
            resized_image = pil_image.resize((TARGET_WIDTH, new_height), Image.Resampling.LANCZOS)
            self.banner_img = ImageTk.PhotoImage(resized_image) # <--- Guardar Referencia
            
            banner_label = tk.Label(self.master, 
                                    image=self.banner_img, 
                                    bg=COLOR_FONDO_CIELO)
            
            # Usar pack para colocar el banner en la parte más inferior
            banner_label.pack(side=tk.BOTTOM, pady=10) 
            
        except Exception as e:
            print(f"ADVERTENCIA: No se pudo cargar el archivo 'banner.png': {e}")
        # =======================================================
        
        
        
        # --- SECCIÓN 1: Metadata Input ---
        frame_metadata = tk.LabelFrame(self.master, text="Datos del Examen", padx=10, pady=5, bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        frame_metadata.pack(padx=10, pady=5, fill="x")
        
        campos = [
            ("Colegio:", "colegio"), ("Nombre del Examen:", "examen"),
            ("Grado:", "grado"), ("Fecha de Aplicación (YYYY-MM-DD):", "fecha"),
            ("Ciudad:", "ciudad")
        ]
        
        for i, (label_text, var_name) in enumerate(campos):
            initial_value = self.initial_metadata.get(var_name, "") # Prellenar si existe
            tk.Label(frame_metadata, text=label_text, bg=COLOR_FONDO_CIELO).grid(row=i, column=0, padx=5, pady=2, sticky="w") # ✅ FONDO AZUL CIELO
            var = tk.StringVar(value=initial_value)
            entry = tk.Entry(frame_metadata, textvariable=var, width=30)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky="ew")
            self.metadata_vars[var_name] = var
            
        frame_metadata.grid_columnconfigure(1, weight=1)

        # --- SECCIÓN 2: Configuración de Materias y Rangos (NUEVO) ---
        frame_rangos = tk.LabelFrame(self.master, text="Paso 1: Configuración Dinámica de Materias y Rangos", padx=10, pady=5, bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        frame_rangos.pack(padx=10, pady=5, fill="x")
        
        # --- MODIFICACIÓN GUI: AÑADIR MULTIPLICADOR Y PUNTOS EXTRA ---
        # Actualización de encabezados para 6 columnas
        tk.Label(frame_rangos, text="Materia", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=0, padx=5, pady=2)
        tk.Label(frame_rangos, text="Pregunta Inicial", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=1, padx=5, pady=2)
        tk.Label(frame_rangos, text="Pregunta Final", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=2, padx=5, pady=2)
        tk.Label(frame_rangos, text="Multiplicador", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=3, padx=5, pady=2) # NUEVO
        tk.Label(frame_rangos, text="Puntos Extra", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=4, padx=5, pady=2) # NUEVO
        tk.Label(frame_rangos, text="¿Calificar?", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=5, padx=5, pady=2) # MOVIDO
        
        # Usar MATERIAS_POTENCIALES para generar las filas dinámicas
        row_num = 1
        for mat, (inicio_default, fin_default) in MATERIAS_POTENCIALES.items():
            
            # Nombre de la materia (Col 0)
            tk.Label(frame_rangos, text=mat, bg=COLOR_FONDO_CIELO).grid(row=row_num, column=0, padx=5, pady=2, sticky="w")

            # Entrada para Pregunta Inicial (Col 1)
            var_inicio = tk.StringVar(value=str(inicio_default))
            entry_inicio = tk.Entry(frame_rangos, textvariable=var_inicio, width=8)
            entry_inicio.grid(row=row_num, column=1, padx=5, pady=2)

            # Entrada para Pregunta Final (Col 2)
            var_fin = tk.StringVar(value=str(fin_default))
            entry_fin = tk.Entry(frame_rangos, textvariable=var_fin, width=8)
            entry_fin.grid(row=row_num, column=2, padx=5, pady=2)
            
            # Multiplicador (Col 3) - Spinbox de 1 a 4
            var_mult = tk.StringVar(value="1")
            spin_mult = tk.Spinbox(frame_rangos, from_=1, to_=4, textvariable=var_mult, width=5)
            spin_mult.grid(row=row_num, column=3, padx=5, pady=2)
            
            # Puntos Extra (Col 4) - Entrada numérica con valor por defecto 0
            var_extra = tk.StringVar(value="0")
            entry_extra = tk.Entry(frame_rangos, textvariable=var_extra, width=7)
            entry_extra.grid(row=row_num, column=4, padx=5, pady=2)


            # Checkbox para seleccionar (Col 5)
            default_checked = mat in ['Lectura crítica', 'Matemáticas', 'Biología', 'C. Sociales', 'Inglés']
            var_check = tk.BooleanVar(value=default_checked) 
            check_box = tk.Checkbutton(frame_rangos, variable=var_check, bg=COLOR_FONDO_CIELO) 
            check_box.grid(row=row_num, column=5, padx=5, pady=2)

            # Guardar referencias para el proceso de calificación
            self.materia_vars[mat] = {
                'inicio': var_inicio,
                'fin': var_fin,
                'multiplicador': var_mult, # NUEVO
                'extra': var_extra,        # NUEVO
                'calificar': var_check
            }

            row_num += 1
        # --- FIN MODIFICACIÓN GUI ---
            
        # --- SECCIÓN 3: Carga de Archivos ---
        frame_archivos = tk.LabelFrame(self.master, text="Paso 2: Carga de Archivos", padx=10, pady=10, bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        frame_archivos.pack(padx=10, pady=5, fill="x")

        tk.Button(frame_archivos, text="1. Cargar Banco de Respuestas (Clave)", command=self.cargar_banco).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Label(frame_archivos, textvariable=self.ruta_banco, fg="darkgreen", bg=COLOR_FONDO_CIELO).grid(row=0, column=1, padx=5, pady=5, sticky="w") # ✅ FONDO AZUL CIELO
        tk.Button(frame_archivos, text="2. Cargar Respuestas Estudiantes", command=self.cargar_respuestas_estudiantes).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        tk.Label(frame_archivos, textvariable=self.ruta_respuestas, fg="darkgreen", bg=COLOR_FONDO_CIELO).grid(row=1, column=1, padx=5, pady=5, sticky="w") # ✅ FONDO AZUL CIELO

        # --- SECCIÓN 4: Acciones ---
        frame_acciones = tk.LabelFrame(self.master, text="Paso 3: Calificación y Exportación de Resultados", padx=10, pady=10, bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        frame_acciones.pack(padx=10, pady=5, fill="x")

        # --- MODIFICACIÓN: ESCALA GLOBAL - Botones de Radio en la GUI ---
        escala_frame = tk.Frame(frame_acciones, bg=COLOR_FONDO_CIELO)
        escala_frame.grid(row=0, column=0, columnspan=2, pady=5, padx=10, sticky="w") 
        
        tk.Label(escala_frame, text="Escala Máxima Global:", 
                 font=("Arial", 10, "bold"), bg=COLOR_FONDO_CIELO).pack(side=tk.LEFT, padx=5)

        tk.Radiobutton(escala_frame, text="100", variable=self.escala_var, value=100, 
                       font=("Arial", 10), bg=COLOR_FONDO_CIELO).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(escala_frame, text="500", variable=self.escala_var, value=500, 
                       font=("Arial", 10), bg=COLOR_FONDO_CIELO).pack(side=tk.LEFT)
        # --- FIN MODIFICACIÓN ESCALA GLOBAL ---

        # El botón de calificar ahora está en la fila 1
        tk.Button(frame_acciones, text="3. CALIFICAR Y GUARDAR EN MEMORIA", command=self.calificar, bg="green", fg="white", font=("Arial", 12, "bold")).grid(row=1, column=0, columnspan=2, padx=10, pady=10, ipadx=20)
        
        # Botón Único: EXPORTAR A JSON CONSOLIDADO (Ahora en la fila 2)
        tk.Button(frame_acciones, 
                  text="4. EXPORTAR RESULTADOS COMPLETOS (JSON)", 
                  command=self.exportar_resultados_json, 
                  bg="#800080", fg="white", font=("Arial", 10, "bold")).grid(row=2, column=0, columnspan=2, padx=10, pady=10, ipadx=10, sticky="ew")
        
        # Etiqueta de estado (Ahora en la fila 3)
        self.estado = tk.Label(frame_acciones, text="Esperando carga de archivos...", fg="red", bg=COLOR_FONDO_CIELO) # ✅ FONDO AZUL CIELO
        self.estado.grid(row=3, column=0, columnspan=2, pady=5) 

        
        self.master.geometry("750x900") 
    
    def cargar_banco(self):
        
        # --- DEFINICIÓN DE LA RUTA FIJA ---
        # Usamos barras inclinadas (/) para compatibilidad con Windows en Python
        DIRECTORIO_FIJO = "calificar"
        # -----------------------------------
        
        ruta = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
        if ruta:
            try:
                # SOLUCIÓN 2: Manejo de errores más específico para el problema de la hoja.
                df_banco = pd.read_excel(ruta, sheet_name='RespuestasCorrectas')
                self.respuestas_correctas = df_banco.iloc[0].str.upper().astype(str)
                self.main_app.respuestas_correctas = self.respuestas_correctas 
                self.ruta_banco.set(ruta.split('/')[-1])
                self.estado.config(text=f"Banco cargado: {self.ruta_banco.get()}", fg="black")
            except ValueError:
                messagebox.showerror("Error de Carga", "No se pudo encontrar la hoja 'RespuestasCorrectas' en el archivo Excel. Por favor, verifique el nombre de la hoja.")
                self.respuestas_correctas = None
            except Exception as e:
                messagebox.showerror("Error de Carga", f"No se pudo cargar el Banco de Respuestas: {e}")
                self.respuestas_correctas = None

    def cargar_respuestas_estudiantes(self):
        # --- DEFINICIÓN DE LA RUTA FIJA ---
        # Usamos barras inclinadas (/) para compatibilidad con Windows en Python
        DIRECTORIO_FIJO = "calificar"
        # -----------------------------------
        
        ruta = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
        if ruta:
            try:
                # SOLUCIÓN 2: Manejo de errores más específico.
                self.df_respuestas = pd.read_excel(ruta, sheet_name='Respuestas')
                # La conversión a mayúsculas y str se aplica a un rango amplio de preguntas (1 a 120)
                for i in range(1, 121):
                    col = f'Pregunta_{i}'
                    if col in self.df_respuestas.columns:
                        self.df_respuestas[col] = self.df_respuestas[col].astype(str).str.upper()

                self.main_app.df_respuestas = self.df_respuestas 
                self.ruta_respuestas.set(ruta.split('/')[-1])
                self.estado.config(text=f"Respuestas cargadas: {self.ruta_respuestas.get()}", fg="black")
            except ValueError:
                messagebox.showerror("Error de Carga", "No se pudo encontrar la hoja 'Respuestas' en el archivo Excel. Por favor, verifique el nombre de la hoja.")
                self.df_respuestas = None
            except Exception as e:
                messagebox.showerror("Error de Carga", f"No se pudo cargar el archivo de respuestas: {e}")
                self.df_respuestas = None
    
    def _collect_metadata(self):
        """Recoge la metadata de los campos de entrada."""
        metadata = {name: var.get() for name, var in self.metadata_vars.items()}
        if not all(metadata.values()):
            return None
        return metadata
        
    # --- MODIFICACIÓN: FUNCIÓN DE RECOLECCIÓN DE RANGOS ---
    def _collect_and_validate_ranges(self):
        """
        Recoge y valida los rangos de preguntas, el multiplicador (1-4) 
        y los puntos extra (>=0) seleccionados por el usuario.
        """
        rangos = {}
        
        for mat, vars in self.materia_vars.items():
            if vars['calificar'].get():
                try:
                    inicio = int(vars['inicio'].get())
                    fin = int(vars['fin'].get())
                    
                    multiplicador = int(vars['multiplicador'].get()) # NUEVO
                    puntos_extra = int(vars['extra'].get())          # NUEVO
                    
                    if inicio <= 0 or fin < inicio:
                        messagebox.showerror("Error de Rango", f"El rango de {mat} es inválido: {inicio}-{fin}. Debe ser mayor a 0 y Final >= Inicial.")
                        return None
                    
                    # Validar Multiplicador (1-4)
                    if not 1 <= multiplicador <= 4:
                        messagebox.showerror("Error de Multiplicador", f"El multiplicador de {mat} debe estar entre 1 y 4.")
                        return None
                        
                    # Validar Puntos Extra (>= 0)
                    if puntos_extra < 0:
                        messagebox.showerror("Error de Puntos Extra", f"Los puntos extra para {mat} no pueden ser negativos.")
                        return None
                        
                    # Guardar todos los parámetros
                    rangos[mat] = {
                        'inicio': inicio,
                        'fin': fin,
                        'multiplicador': multiplicador,
                        'extra': puntos_extra
                    }
                    
                except ValueError:
                    messagebox.showerror("Error de Entrada", f"Las entradas de {mat} (Preguntas, Multiplicador o Puntos Extra) deben ser números enteros.")
                    return None
                    
        if not rangos:
             messagebox.showwarning("Selección Vacía", "Debe seleccionar al menos una materia para calificar y definir su rango.")
             return None
             
        self.rangos_calificar = rangos
        return rangos
    # --- FIN MODIFICACIÓN RECOLECCIÓN DE RANGOS ---

    # --- MODIFICACIÓN: FUNCIÓN CALIFICAR ---
    def calificar(self):
        # 0. Obtener rangos seleccionados y validar (nueva estructura)
        rangos_a_calificar = self._collect_and_validate_ranges()
        if rangos_a_calificar is None:
             return

        if self.df_respuestas is None or self.respuestas_correctas is None:
            messagebox.showwarning("Faltan Archivos", "Por favor, carga ambos archivos de Excel antes de calificar (Paso 2).")
            return

        metadata = self._collect_metadata()
        if metadata is None:
             messagebox.showwarning("Metadata Incompleta", "Por favor, complete todos los campos de 'Datos del Examen'.")
             return
             
        try:
            df_calificaciones = self.df_respuestas[['ID_Estudiante', 'Nombre']].copy()
        except KeyError as e:
            messagebox.showerror("Error de Columna", f"El archivo de respuestas debe tener las columnas 'ID_Estudiante' y 'Nombre'. Error: {e}")
            return

        # Obtener la escala seleccionada
        escala_seleccionada = self.escala_var.get()
        
        # Calificación basada en los rangos seleccionados
        for materia, config in rangos_a_calificar.items():
            inicio = config['inicio']
            fin = config['fin']
            multiplicador = config['multiplicador']
            puntos_extra = config['extra']
            
            columnas_materia = [f'Pregunta_{i}' for i in range(inicio, fin + 1)]
            
            correctas_materia_serie = self.respuestas_correctas.reindex(columnas_materia).dropna()
            
            # Máximo de preguntas posibles (Aciertos base)
            max_aciertos_base = len(correctas_materia_serie) 
            
            if not correctas_materia_serie.empty:
                df_respuestas_materia_final = self.df_respuestas[correctas_materia_serie.index]
                
                # 1. Calcular aciertos base (0 or 1)
                df_aciertos_base = (df_respuestas_materia_final.eq(correctas_materia_serie, axis=1))
                aciertos_raw = df_aciertos_base.sum(axis=1) # Suma de aciertos (0 a Max_Aciertos_Base)
                
                # 2. Aplicar Multiplicador
                aciertos_multiplicados = aciertos_raw * multiplicador
                
                # 3. Sumar Puntos Extra
                aciertos_ajustados = aciertos_multiplicados + puntos_extra
                
                # 4. Enforzar Límite: Max Aciertos = Max Preguntas * Multiplicador
                max_aciertos_posibles_ajustado = max_aciertos_base * multiplicador
                
                # Si el puntaje ajustado excede el límite (Max Q * Multiplier), se recorta al límite.
                df_calificaciones[f'Aciertos_{materia}'] = aciertos_ajustados.clip(upper=max_aciertos_posibles_ajustado)
                
                # Guardar el máximo de aciertos REAL para la normalización en preparar_datos_estadisticos
                df_calificaciones[f'Max_Aciertos_{materia}'] = max_aciertos_posibles_ajustado
                
            else:
                 df_calificaciones[f'Aciertos_{materia}'] = 0
                 df_calificaciones[f'Max_Aciertos_{materia}'] = 0 # 0 si no se calificó

        columnas_aciertos = [col for col in df_calificaciones.columns if col.startswith('Aciertos_')]
        df_calificaciones['Aciertos_Total'] = df_calificaciones[columnas_aciertos].sum(axis=1)

        # Usar los rangos calificados para la normalización, cálculo global y obtener las áreas calificadas
        df_calificaciones, areas_principales_calificadas = preparar_datos_estadisticos(df_calificaciones, rangos_a_calificar, escala_seleccionada)
        
        # --- NUEVO: 3. Calcular Promedio del Grupo por Materia (0-100) ---
        promedios_grupo = {}
        for area in areas_principales_calificadas:
            col_puntaje = f'Puntaje_{area}'
            if col_puntaje in df_calificaciones.columns:
                promedios_grupo[area] = df_calificaciones[col_puntaje].mean()
        # -----------------------------------------------------------------

        # 1. Actualizar el estado centralizado con el DF calificado
        self.main_app.df_calificaciones = df_calificaciones 
        
        # 2. Guardar la metadata en el estado centralizado, incluyendo los datos de calibración
        self.main_app.metadata = metadata 
        self.main_app.metadata['rangos_calibracion_detallados'] = rangos_a_calificar
        self.main_app.metadata['areas_principales_calificadas'] = areas_principales_calificadas
        self.main_app.metadata['promedios_grupo'] = promedios_grupo 
        self.main_app.metadata['Puntaje_Global_Max'] = escala_seleccionada # Guardar la escala para el reporte


        self.estado.config(text="Calificación COMPLETA. Resultados y Datos del Examen guardados en memoria.", fg="blue")
        messagebox.showinfo("Proceso Completo", "La calificación de los exámenes se ha completado con éxito y los resultados están listos para exportar.")
        
    # --- FIN MODIFICACIÓN CALIFICAR ---
    
    # --- MODIFICADO: exportar_resultados_json (Anteriormente exportar_metadata_json) ---
    def exportar_resultados_json(self):
        
        # --- DEFINICIÓN DE LA RUTA FIJA ---
        DIRECTORIO_FIJO = resource_path(os.path.join("resultados_excel"))
        
        df_final = self.main_app.df_calificaciones
        metadata = self.main_app.metadata
        
        if df_final is None or df_final.empty:
            messagebox.showwarning("Faltan Datos", "Primero debes calificar los exámenes.")
            return

        # 1. PEDIR AL USUARIO EL NOMBRE DEL ARCHIVO
        ruta_completa_seleccionada = filedialog.asksaveasfilename(
            defaultextension=".json", 
            filetypes=[("Archivos JSON", "*.json")], 
            initialdir=DIRECTORIO_FIJO,
            initialfile="Resultados_Consolidados.json",
            title="Guardar resultados"
        )
        
        if ruta_completa_seleccionada:
            nombre_base = os.path.splitext(os.path.basename(ruta_completa_seleccionada))[0]

            try:
                os.makedirs(DIRECTORIO_FIJO, exist_ok=True)
            except OSError as e:
                messagebox.showerror("Error", f"No se pudo crear el directorio: {e}")
                return
            
            ruta_json = os.path.join(DIRECTORIO_FIJO, nombre_base + ".json")
            ruta_excel = os.path.join(DIRECTORIO_FIJO, nombre_base + ".xlsx")
            
            # --- ORDENAMIENTO Y CÁLCULO DE PROMEDIO DE CIENCIAS ---
            try:
                df_export = df_final.copy()

                # Lista de materias en el orden solicitado
                materias_orden = [
                    'Lectura crítica', 
                    'Matemáticas', 
                    'C. Sociales', 
                    'Biología', 
                    'Química', 
                    'Física', 
                    'Inglés'
                ]

                # CÁLCULO DEL PROMEDIO DE CIENCIAS (Biología, Química, Física)
                cols_ciencias = ['Puntaje_Biología', 'Puntaje_Química', 'Puntaje_Física']
                # Verificamos que las tres existan para promediar
                existentes = [c for c in cols_ciencias if c in df_export.columns]
                if existentes:
                    df_export['Puntaje_Ciencias'] = df_export[existentes].mean(axis=1).round(2)
                
                # Definimos el orden de las columnas finales para el JSON
                cols_id = ['ID_Estudiante', 'Nombre', 'Puesto', 'Puntaje_Global', 'Percentil_Global']
                
                # Construimos la lista de puntajes siguiendo el orden + Ciencias al final
                cols_puntajes = [f'Puntaje_{m}' for m in materias_orden if f'Puntaje_{m}' in df_export.columns]
                if 'Puntaje_Ciencias' in df_export.columns:
                    cols_puntajes.append('Puntaje_Ciencias')
                
                # Construimos la lista de aciertos
                cols_aciertos = [f'Aciertos_{m}' for m in materias_orden if f'Aciertos_{m}' in df_export.columns]

                # Unimos todo en el orden jerárquico final
                orden_final_columnas = cols_id + cols_puntajes + cols_aciertos
                
                # Agregamos cualquier columna extra que no esté en nuestra lista
                extras = [c for c in df_export.columns if c not in orden_final_columnas]
                df_export = df_export[orden_final_columnas + extras]

                # Redondeo final de seguridad
                for col in df_export.columns:
                    if any(key in col for key in ['Puntaje_', 'Percentil_']):
                        df_export[col] = pd.to_numeric(df_export[col], errors='coerce').round(2)
                        
            except Exception as e:
                messagebox.showerror("Error de Procesamiento", f"Fallo al organizar materias: {e}")
                return

            # --- EXPORTACIÓN JSON ---
            try:
                resultados_list = df_export.to_dict('records')
                consolidated_data = {
                    "metadata": metadata,
                    "resultados_estudiantes": resultados_list
                }
                with open(ruta_json, 'w', encoding='utf-8') as f:
                    json.dump(consolidated_data, f, indent=4, ensure_ascii=False) 
            except Exception as e:
                messagebox.showerror("Error JSON", f"No se pudo guardar: {e}")
                return
            
            # --- EXPORTACIÓN EXCEL ---
            try:
                df_export.to_excel(ruta_excel, index=False, sheet_name='Resultados_Consolidados')
            except Exception as e:
                messagebox.showwarning("Aviso", "JSON guardado, pero Excel falló.")
                
            self.estado.config(text="Exportación exitosa. Ciencias promediado al final.", fg="#800080")
            messagebox.showinfo("Éxito", "Archivos creados con el orden y promedio de ciencias.")
            
# -------------------------------------------------------------
# --- CLASE 3: LÓGICA DE IMPORTACIÓN DE ACIERTOS Y REPORTES ---
# -------------------------------------------------------------
class AppImportarAciertos:
    def __init__(self, master, main_app, initial_metadata):
        self.master = master 
        self.main_app = main_app 
        
        self.df_aciertos_importados = None
        self.metadata_vars = {} 
        self.initial_metadata = initial_metadata
        
        self.rangos_calificar = {} 
        self.materia_vars = {}     
        self.escala_var = tk.IntVar(value=500) 
        
        self.ruta_excel = tk.StringVar(value="Excel de Aciertos (No Cargado)")
        self.estado = None 
             
        self.crear_widgets()

    def crear_widgets(self):
        # --- SECCIÓN 1: Metadata Input ---
        frame_metadata = tk.LabelFrame(self.master, text="Datos del Examen", padx=10, pady=5, bg=COLOR_FONDO_CIELO) 
        frame_metadata.pack(padx=10, pady=5, fill="x")
        
        campos = [
            ("Colegio:", "colegio"), ("Nombre del Examen:", "examen"),
            ("Grado:", "grado"), ("Fecha de Aplicación (YYYY-MM-DD):", "fecha"),
            ("Ciudad:", "ciudad")
        ]
        
        for i, (label_text, var_name) in enumerate(campos):
            initial_value = self.initial_metadata.get(var_name, "") 
            tk.Label(frame_metadata, text=label_text, bg=COLOR_FONDO_CIELO).grid(row=i, column=0, padx=5, pady=2, sticky="w") 
            var = tk.StringVar(value=initial_value)
            entry = tk.Entry(frame_metadata, textvariable=var, width=30)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky="ew")
            self.metadata_vars[var_name] = var
            
        frame_metadata.grid_columnconfigure(1, weight=1)

        # --- SECCIÓN 2: Configuración de Materias (Para calcular máximos) ---
        frame_rangos = tk.LabelFrame(self.master, text="Configuración de Materias (Necesario para calcular %)", padx=10, pady=5, bg=COLOR_FONDO_CIELO)
        frame_rangos.pack(padx=10, pady=5, fill="x")
        
        tk.Label(frame_rangos, text="Materia", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=0, padx=5, pady=2)
        tk.Label(frame_rangos, text="Preg. Inicial", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=1, padx=5, pady=2)
        tk.Label(frame_rangos, text="Preg. Final", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=2, padx=5, pady=2)
        tk.Label(frame_rangos, text="Multiplicador", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=3, padx=5, pady=2)
        tk.Label(frame_rangos, text="¿Evaluar?", font=("Arial", 9, "bold"), bg=COLOR_FONDO_CIELO).grid(row=0, column=4, padx=5, pady=2)
        
        row_num = 1
        for mat, (inicio_default, fin_default) in MATERIAS_POTENCIALES.items():
            tk.Label(frame_rangos, text=mat, bg=COLOR_FONDO_CIELO).grid(row=row_num, column=0, padx=5, pady=2, sticky="w")

            var_inicio = tk.StringVar(value=str(inicio_default))
            tk.Entry(frame_rangos, textvariable=var_inicio, width=8).grid(row=row_num, column=1, padx=5, pady=2)

            var_fin = tk.StringVar(value=str(fin_default))
            tk.Entry(frame_rangos, textvariable=var_fin, width=8).grid(row=row_num, column=2, padx=5, pady=2)
            
            var_mult = tk.StringVar(value="1")
            tk.Spinbox(frame_rangos, from_=1, to_=4, textvariable=var_mult, width=5).grid(row=row_num, column=3, padx=5, pady=2)
            
            # Puntos extra se omiten en importación directa para simplificar, a menos que el excel no los incluya.
            var_extra = tk.StringVar(value="0") 

            default_checked = mat in ['Lectura crítica', 'Matemáticas', 'Biología', 'C. Sociales', 'Inglés']
            var_check = tk.BooleanVar(value=default_checked) 
            tk.Checkbutton(frame_rangos, variable=var_check, bg=COLOR_FONDO_CIELO).grid(row=row_num, column=4, padx=5, pady=2)

            self.materia_vars[mat] = {'inicio': var_inicio, 'fin': var_fin, 'multiplicador': var_mult, 'extra': var_extra, 'calificar': var_check}
            row_num += 1
            
        # --- SECCIÓN 3: Carga de Archivos ---
        frame_archivos = tk.LabelFrame(self.master, text="Carga de Archivos", padx=10, pady=10, bg=COLOR_FONDO_CIELO) 
        frame_archivos.pack(padx=10, pady=5, fill="x")

        tk.Button(frame_archivos, text="1. Cargar Excel de Aciertos", command=self.cargar_excel_aciertos).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Label(frame_archivos, textvariable=self.ruta_excel, fg="darkgreen", bg=COLOR_FONDO_CIELO).grid(row=0, column=1, padx=5, pady=5, sticky="w") 

        # --- SECCIÓN 4: Acciones ---
        frame_acciones = tk.LabelFrame(self.master, text="Procesamiento y Exportación", padx=10, pady=10, bg=COLOR_FONDO_CIELO)
        frame_acciones.pack(padx=10, pady=5, fill="x")

        escala_frame = tk.Frame(frame_acciones, bg=COLOR_FONDO_CIELO)
        escala_frame.grid(row=0, column=0, columnspan=2, pady=5, padx=10, sticky="w") 
        tk.Label(escala_frame, text="Escala Máxima Global:", font=("Arial", 10, "bold"), bg=COLOR_FONDO_CIELO).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(escala_frame, text="100", variable=self.escala_var, value=100, bg=COLOR_FONDO_CIELO).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(escala_frame, text="500", variable=self.escala_var, value=500, bg=COLOR_FONDO_CIELO).pack(side=tk.LEFT)

        tk.Button(frame_acciones, text="2. PROCESAR Y EXPORTAR RESULTADOS (JSON/EXCEL)", command=self.procesar_y_exportar, bg="#800080", fg="white", font=("Arial", 12, "bold")).grid(row=1, column=0, columnspan=2, padx=10, pady=10, ipadx=20)
        
        self.estado = tk.Label(frame_acciones, text="Esperando carga de Excel...", fg="red", bg=COLOR_FONDO_CIELO)
        self.estado.grid(row=2, column=0, columnspan=2, pady=5) 
        
        self.master.geometry("750x700")

    def cargar_excel_aciertos(self):
        ruta = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
        if ruta:
            try:
                self.df_aciertos_importados = pd.read_excel(ruta)
                # Validar columnas básicas
                if 'ID_Estudiante' not in self.df_aciertos_importados.columns or 'Nombre' not in self.df_aciertos_importados.columns:
                    messagebox.showerror("Error", "El Excel debe contener las columnas 'ID_Estudiante' y 'Nombre'.")
                    self.df_aciertos_importados = None
                    return
                    
                self.ruta_excel.set(ruta.split('/')[-1])
                self.estado.config(text=f"Excel cargado: {self.ruta_excel.get()}", fg="black")
            except Exception as e:
                messagebox.showerror("Error de Carga", f"No se pudo cargar el archivo Excel: {e}")
                self.df_aciertos_importados = None

    def _collect_metadata(self):
        metadata = {name: var.get() for name, var in self.metadata_vars.items()}
        if not all(metadata.values()): return None
        return metadata

    def _collect_and_validate_ranges(self):
        rangos = {}
        for mat, vars in self.materia_vars.items():
            if vars['calificar'].get():
                try:
                    inicio = int(vars['inicio'].get())
                    fin = int(vars['fin'].get())
                    multiplicador = int(vars['multiplicador'].get())
                    
                    if inicio <= 0 or fin < inicio: return None
                    rangos[mat] = {'inicio': inicio, 'fin': fin, 'multiplicador': multiplicador, 'extra': 0}
                except ValueError: return None
        if not rangos: return None
        return rangos

    def procesar_y_exportar(self):
        rangos_a_calificar = self._collect_and_validate_ranges()
        if rangos_a_calificar is None:
             messagebox.showerror("Error", "Revisa la configuración de materias y rangos.")
             return

        if self.df_aciertos_importados is None:
            messagebox.showwarning("Faltan Archivos", "Por favor, carga el Excel de aciertos.")
            return

        metadata = self._collect_metadata()
        if metadata is None:
             messagebox.showwarning("Metadata Incompleta", "Completa los 'Datos del Examen'.")
             return

        df_procesado = self.df_aciertos_importados.copy()
        escala_seleccionada = self.escala_var.get()

        # Generar las columnas Max_Aciertos para que la función de estadística haga su trabajo
        for materia, config in rangos_a_calificar.items():
            inicio = config['inicio']
            fin = config['fin']
            multiplicador = config['multiplicador']
            
            cantidad_preguntas = (fin - inicio) + 1
            max_aciertos_posibles = cantidad_preguntas * multiplicador
            
            col_aciertos = f'Aciertos_{materia}'
            
            if col_aciertos in df_procesado.columns:
                df_procesado[col_aciertos] = df_procesado[col_aciertos].clip(upper=max_aciertos_posibles)
                df_procesado[f'Max_Aciertos_{materia}'] = max_aciertos_posibles
            else:
                df_procesado[col_aciertos] = 0
                df_procesado[f'Max_Aciertos_{materia}'] = 0

        # Pasar por el mismo motor estadístico que la calificación normal
        df_calificaciones, areas_principales_calificadas = preparar_datos_estadisticos(df_procesado, rangos_a_calificar, escala_seleccionada)
        
        promedios_grupo = {}
        for area in areas_principales_calificadas:
            col_puntaje = f'Puntaje_{area}'
            if col_puntaje in df_calificaciones.columns:
                promedios_grupo[area] = df_calificaciones[col_puntaje].mean()

        # Guardar en memoria de la App Principal
        self.main_app.df_calificaciones = df_calificaciones 
        self.main_app.metadata = metadata 
        self.main_app.metadata['rangos_calibracion_detallados'] = rangos_a_calificar
        self.main_app.metadata['areas_principales_calificadas'] = areas_principales_calificadas
        self.main_app.metadata['promedios_grupo'] = promedios_grupo 
        self.main_app.metadata['Puntaje_Global_Max'] = escala_seleccionada

        self.estado.config(text="Procesamiento COMPLETO. Iniciando exportación...", fg="blue")
        
        # Llamar a la misma función de exportación que ya tienes configurada
        self._exportar_resultados()

    def _exportar_resultados(self):
        # Esta función es idéntica a tu exportar_resultados_json original
        DIRECTORIO_FIJO = resource_path(os.path.join("resultados_excel"))
        df_final = self.main_app.df_calificaciones
        metadata = self.main_app.metadata

        ruta_completa_seleccionada = filedialog.asksaveasfilename(
            defaultextension=".json", 
            filetypes=[("Archivos JSON", "*.json")], 
            initialdir=DIRECTORIO_FIJO,
            initialfile="Resultados_Consolidados_Importados.json",
            title="Guardar resultados"
        )
        
        if ruta_completa_seleccionada:
            nombre_base = os.path.splitext(os.path.basename(ruta_completa_seleccionada))[0]
            os.makedirs(DIRECTORIO_FIJO, exist_ok=True)
            ruta_json = os.path.join(DIRECTORIO_FIJO, nombre_base + ".json")
            ruta_excel = os.path.join(DIRECTORIO_FIJO, nombre_base + ".xlsx")
            
            try:
                df_export = df_final.copy()
                materias_orden = ['Lectura crítica', 'Matemáticas', 'C. Sociales', 'Biología', 'Química', 'Física', 'Inglés']

                cols_ciencias = ['Puntaje_Biología', 'Puntaje_Química', 'Puntaje_Física']
                existentes = [c for c in cols_ciencias if c in df_export.columns]
                if existentes:
                    df_export['Puntaje_Ciencias'] = df_export[existentes].mean(axis=1).round(2)
                
                cols_id = ['ID_Estudiante', 'Nombre', 'Puesto', 'Puntaje_Global', 'Percentil_Global']
                cols_puntajes = [f'Puntaje_{m}' for m in materias_orden if f'Puntaje_{m}' in df_export.columns]
                if 'Puntaje_Ciencias' in df_export.columns: cols_puntajes.append('Puntaje_Ciencias')
                cols_aciertos = [f'Aciertos_{m}' for m in materias_orden if f'Aciertos_{m}' in df_export.columns]

                orden_final_columnas = cols_id + cols_puntajes + cols_aciertos
                extras = [c for c in df_export.columns if c not in orden_final_columnas]
                df_export = df_export[orden_final_columnas + extras]

                for col in df_export.columns:
                    if any(key in col for key in ['Puntaje_', 'Percentil_']):
                        df_export[col] = pd.to_numeric(df_export[col], errors='coerce').round(2)
                        
            except Exception as e:
                messagebox.showerror("Error", f"Fallo al organizar materias: {e}")
                return

            try:
                resultados_list = df_export.to_dict('records')
                consolidated_data = {"metadata": metadata, "resultados_estudiantes": resultados_list}
                with open(ruta_json, 'w', encoding='utf-8') as f:
                    json.dump(consolidated_data, f, indent=4, ensure_ascii=False) 
            except Exception as e:
                messagebox.showerror("Error JSON", f"No se pudo guardar: {e}")
                return
            
            try:
                df_export.to_excel(ruta_excel, index=False, sheet_name='Resultados_Importados')
            except Exception as e:
                pass
                
            self.estado.config(text="Exportación exitosa.", fg="#800080")
            messagebox.showinfo("Éxito", "Archivos JSON y Excel creados exitosamente.")





# ---------------------------------------------
# --- FUNCIÓN PRINCIPAL DE EJECUCIÓN ---
# ---------------------------------------------
def main():
    root = tk.Tk()
    root.title("Calificación de Respuestas")
    
    # =======================================================
    # === CÓDIGOS PARA AGREGAR EL LOGO A TODAS LAS VENTANAS ===
    # La ruta usa barras / para evitar problemas de escape en Windows
    LOGO_PATH = resource_path(os.path.join("imagenes", "logo.png"))

    try:
        # 1. Cargar la imagen (PNG). Tkinter usa PhotoImage para íconos.
        logo_img = tk.PhotoImage(file=LOGO_PATH)
        
        # 2. Aplicar el ícono a la ventana principal Y a todas las subventanas (por el 'True')
        root.iconphoto(True, logo_img)
        
    except tk.TclError:
        # Manejo de error si el archivo no se encuentra o no es un formato válido
        print(f"ADVERTENCIA: No se pudo cargar el logo desde la ruta: {LOGO_PATH}. Verifica que exista y sea PNG/GIF.")
    # =======================================================

    aplicacion = CalificadorApp(root) 
    root.mainloop()

if __name__ == "__main__":
    main()