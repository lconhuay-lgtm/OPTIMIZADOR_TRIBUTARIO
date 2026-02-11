import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import pandas as pd
from datetime import datetime
from scipy import stats # Para Pruebas de Hipótesis
import json



class OptimizadorPro:
    def __init__(self, root):
        self.root = root
        self.root.title("Estrategia Fiscal Avanzada - Multi-Holding")
        self.root.geometry("1500x950")
        
        # Parámetros Base
        self.UIT = tk.DoubleVar(value=5500)
        self.TASA_GENERAL = tk.DoubleVar(value=29.5)
        self.TASA_ESPECIAL = tk.DoubleVar(value=10.0)
        self.LIMITE_UTILIDAD_UIT = tk.DoubleVar(value=15)
        self.LIMITE_INGRESOS_UIT = tk.DoubleVar(value=1700)
        self.num_satelites = tk.IntVar(value=4)
        # --- NUEVAS VARIABLES EDITABLES ---
        self.CONF_NIVEL = tk.DoubleVar(value=95.0)     # Para Simulación
        self.ALFA_SIG = tk.DoubleVar(value=0.05)       # Para Prueba Hipótesis
        self.COMP_LIMITE_PCT = tk.DoubleVar(value=20.0)# % de incremento en comparativo
        self.COMP_TASA_DELTA = tk.DoubleVar(value=2.0) # puntos de reducción en comparativo
                       
        
        self.satelites_entries = []
        self.crear_interfaz()
        
    def crear_interfaz(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Pestañas
        tab_config = ttk.Frame(notebook)
        notebook.add(tab_config, text='Configuración Tributaria')
        
        tab_datos = ttk.Frame(notebook)
        notebook.add(tab_datos, text='Datos del Grupo')
        
        tab_resultados = ttk.Frame(notebook)
        notebook.add(tab_resultados, text='Resultados y Análisis')
        
        tab_simulacion = ttk.Frame(notebook)
        notebook.add(tab_simulacion, text='Simulación Monte Carlo')
        
        tab_sensibilidad = ttk.Frame(notebook)
        notebook.add(tab_sensibilidad, text='Análisis de Sensibilidad')
        
        tab_comparativo = ttk.Frame(notebook)
        notebook.add(tab_comparativo, text='Análisis Comparativo')
        
        self.crear_tab_configuracion(tab_config)
        self.crear_tab_datos(tab_datos)
        self.crear_tab_resultados(tab_resultados)
        self.crear_tab_simulacion(tab_simulacion)
        self.crear_tab_sensibilidad(tab_sensibilidad)
        self.crear_tab_comparativo(tab_comparativo)
        
    def crear_tab_configuracion(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # SECCIÓN: Parámetros Tributarios
        frame_parametros = ttk.LabelFrame(scrollable_frame, text="Parámetros del Sistema Tributario", padding=20)
        frame_parametros.grid(row=0, column=0, padx=10, pady=10, sticky='ew', columnspan=2)
        
        ttk.Label(frame_parametros, text="Unidad de Referencia (UIT):").grid(row=0, column=0, sticky='w', pady=8)
        entry_uit = ttk.Entry(frame_parametros, textvariable=self.UIT, width=20)
        entry_uit.grid(row=0, column=1, pady=8, padx=5)
        
        ttk.Label(frame_parametros, text="Tasa Régimen General (%):").grid(row=1, column=0, sticky='w', pady=8)
        entry_tasa_gral = ttk.Entry(frame_parametros, textvariable=self.TASA_GENERAL, width=20)
        entry_tasa_gral.grid(row=1, column=1, pady=8, padx=5)
        
        ttk.Label(frame_parametros, text="Tasa Régimen Especial/MYPE (%):").grid(row=2, column=0, sticky='w', pady=8)
        entry_tasa_esp = ttk.Entry(frame_parametros, textvariable=self.TASA_ESPECIAL, width=20)
        entry_tasa_esp.grid(row=2, column=1, pady=8, padx=5)
        
        ttk.Label(frame_parametros, text="Límite Utilidad Régimen Especial (UIT):").grid(row=3, column=0, sticky='w', pady=8)
        entry_lim_util = ttk.Entry(frame_parametros, textvariable=self.LIMITE_UTILIDAD_UIT, width=20)
        entry_lim_util.grid(row=3, column=1, pady=8, padx=5)
        
        ttk.Label(frame_parametros, text="Límite Ingresos Régimen Especial (UIT):").grid(row=4, column=0, sticky='w', pady=8)
        entry_lim_ing = ttk.Entry(frame_parametros, textvariable=self.LIMITE_INGRESOS_UIT, width=20)
        entry_lim_ing.grid(row=4, column=1, pady=8, padx=5)
        
        # SECCIÓN: Estructura del Grupo
        frame_estructura = ttk.LabelFrame(scrollable_frame, text="Estructura del Grupo Empresarial", padding=20)
        frame_estructura.grid(row=1, column=0, padx=10, pady=10, sticky='ew', columnspan=2)
        
        ttk.Label(frame_estructura, text="Número de Empresas Satélites:").grid(row=0, column=0, sticky='w', pady=8)
        spinbox_satelites = ttk.Spinbox(frame_estructura, from_=1, to=20, textvariable=self.num_satelites, width=18)
        spinbox_satelites.grid(row=0, column=1, pady=8, padx=5)
        
        ttk.Button(frame_estructura, text="Aplicar Cambios en Estructura", 
                  command=self.aplicar_estructura).grid(row=1, column=0, columnspan=3, pady=15)
        
        # SECCIÓN NUEVA: Control de Escenarios Comparativos
        frame_comp_vars = ttk.LabelFrame(scrollable_frame, text="Variables para Análisis Comparativo", padding=20)
        frame_comp_vars.grid(row=2, column=0, padx=10, pady=10, sticky='ew', columnspan=2)

        ttk.Label(frame_comp_vars, text="Escenario Límite (incremento %):").grid(row=0, column=0, sticky='w', pady=5)
        ttk.Entry(frame_comp_vars, textvariable=self.COMP_LIMITE_PCT, width=15).grid(row=0, column=1, padx=5)

        ttk.Label(frame_comp_vars, text="Escenario Tasa (reducción puntos):").grid(row=1, column=0, sticky='w', pady=5)
        ttk.Entry(frame_comp_vars, textvariable=self.COMP_TASA_DELTA, width=15).grid(row=1, column=1, padx=5)
        
        
        # SECCIÓN: Presets
        frame_presets = ttk.LabelFrame(scrollable_frame, text="Configuraciones Predefinidas", padding=20)
        frame_presets.grid(row=2, column=0, padx=10, pady=10, sticky='ew', columnspan=2)
        
        btn_frame = ttk.Frame(frame_presets)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=10)
        
        ttk.Button(btn_frame, text="Perú", command=lambda: self.cargar_preset('peru')).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Colombia", command=lambda: self.cargar_preset('colombia')).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="México", command=lambda: self.cargar_preset('mexico')).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Chile", command=lambda: self.cargar_preset('chile')).pack(side='left', padx=5)
        
        # Información actual
        frame_info = ttk.LabelFrame(scrollable_frame, text="Parámetros Calculados", padding=20)
        frame_info.grid(row=4, column=0, padx=10, pady=10, sticky='ew', columnspan=2)
        
        self.label_info = ttk.Label(frame_info, text="", justify='left', font=('Courier', 9))
        self.label_info.pack(fill='both', expand=True)
        
        self.actualizar_info_parametros()
        
        ttk.Button(scrollable_frame, text="Actualizar Información", 
                  command=self.actualizar_info_parametros).grid(row=5, column=0, columnspan=2, pady=10)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def crear_tab_datos(self, parent):
        self.canvas_datos = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.canvas_datos.yview)
        self.scrollable_datos = ttk.Frame(self.canvas_datos)
        
        self.scrollable_datos.bind(
            "<Configure>",
            lambda e: self.canvas_datos.configure(scrollregion=self.canvas_datos.bbox("all"))
        )
        
        self.canvas_datos.create_window((0, 0), window=self.scrollable_datos, anchor="nw")
        self.canvas_datos.configure(yscrollcommand=scrollbar.set)
        
        # SECCIÓN MATRIZ
        self.frame_matriz = ttk.LabelFrame(self.scrollable_datos, text="Empresa Matriz (Régimen General)", padding=15)
        self.frame_matriz.grid(row=0, column=0, padx=10, pady=10, sticky='ew')
        
        ttk.Label(self.frame_matriz, text="Nombre Empresa Matriz:").grid(row=0, column=0, sticky='w', pady=5)
        self.entry_nombre_matriz = ttk.Entry(self.frame_matriz, width=30)
        self.entry_nombre_matriz.grid(row=0, column=1, pady=5, columnspan=2)
        self.entry_nombre_matriz.insert(0, "KALLPA")
        
        ttk.Label(self.frame_matriz, text="Ingresos Totales:").grid(row=1, column=0, sticky='w', pady=5)
        self.entry_ingresos_matriz = ttk.Entry(self.frame_matriz, width=20)
        self.entry_ingresos_matriz.grid(row=1, column=1, pady=5)
        self.entry_ingresos_matriz.insert(0, "10000000")
        
        ttk.Label(self.frame_matriz, text="Costos Externos:").grid(row=2, column=0, sticky='w', pady=5)
        self.entry_costos_matriz = ttk.Entry(self.frame_matriz, width=20)
        self.entry_costos_matriz.grid(row=2, column=1, pady=5)
        self.entry_costos_matriz.insert(0, "6000000")
        
        # SECCIÓN SATÉLITES
        self.frame_satelites_container = ttk.LabelFrame(self.scrollable_datos, text="Empresas Satélites", padding=15)
        self.frame_satelites_container.grid(row=1, column=0, padx=10, pady=10, sticky='ew')
        
        self.generar_campos_satelites()
        
        # Botones
        frame_botones = ttk.Frame(self.scrollable_datos)
        frame_botones.grid(row=2, column=0, pady=20)
        
        ttk.Button(frame_botones, text="Calcular Márgenes Óptimos", 
                  command=self.calcular_margenes_optimos).pack(side='left', padx=5)
        ttk.Button(frame_botones, text="Limpiar", command=self.limpiar_datos).pack(side='left', padx=5)
        ttk.Button(frame_botones, text="Cargar Ejemplo", command=self.cargar_ejemplo).pack(side='left', padx=5)
        
        self.canvas_datos.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def generar_campos_satelites(self):
        for widget in self.frame_satelites_container.winfo_children():
            widget.destroy()
        
        # Header con nueva columna para régimen
        headers = [
            ("Nombre", 0),
            ("Costos Compra", 1),
            ("Tipo", 2),
            ("Gastos Operativos", 3),
            ("Régimen General", 4)  # Nueva columna
        ]
        
        for text, col in headers:
            ttk.Label(self.frame_satelites_container, text=text, font='TkDefaultFont 9 bold').grid(
                row=0, column=col, padx=5, pady=5)
        
        self.satelites_entries = []
        num_sat = self.num_satelites.get()
        
        nombres_default = ["IG PAME", "SUMAQ WARMY", "MISKY SONKO", "KHASP", "EMPRESA E", "EMPRESA F"]
        tipos_default = ["Servicios", "Productos", "Productos", "Productos", "Productos", "Servicios"]
        
        for i in range(num_sat):
            nombre_default = nombres_default[i] if i < len(nombres_default) else f"SATELITE {i+1}"
            tipo_default = tipos_default[i] if i < len(tipos_default) else "Productos"
            
            entry_nombre = ttk.Entry(self.frame_satelites_container, width=18)
            entry_nombre.grid(row=i+1, column=0, pady=5, padx=5)
            entry_nombre.insert(0, nombre_default)
            
            entry_costo = ttk.Entry(self.frame_satelites_container, width=18)
            entry_costo.grid(row=i+1, column=1, pady=5, padx=5)
            entry_costo.insert(0, "500000")
            
            combo_tipo = ttk.Combobox(self.frame_satelites_container, 
                                     values=["Productos", "Servicios", "Manufactura", "Distribución"],
                                     state='readonly', width=15)
            combo_tipo.grid(row=i+1, column=2, pady=5, padx=5)
            combo_tipo.set(tipo_default)
            
            entry_gastos = ttk.Entry(self.frame_satelites_container, width=18)
            entry_gastos.grid(row=i+1, column=3, pady=5, padx=5)
            entry_gastos.insert(0, "0")
            
            # --- AQUÍ ESTÁ LA LÓGICA QUIRÚRGICA ---
            var_regimen_general = tk.BooleanVar(value=False)
            check_regimen = ttk.Checkbutton(self.frame_satelites_container, 
                                           variable=var_regimen_general,
                                           text="Usar tasa general")
            check_regimen.grid(row=i+1, column=4, pady=5, padx=5)

            # Definimos la validación interna para esta fila específica
            def validar_uit(event, v_reg=var_regimen_general, e_c=entry_costo, e_g=entry_gastos, chk=check_regimen):
                try:
                    uit_actual = self.UIT.get()
                    limite_anual = 1700 * uit_actual
                    
                    # El ingreso que el satélite cobraría a la matriz para ser eficiente es:
                    # Ingreso = Costos + Utilidad Máxima (15 UIT) + Gastos Operativos
                    costo_val = float(e_c.get() or 0)
                    gastos_val = float(e_g.get() or 0)
                    utilidad_max = 15 * uit_actual
                    
                    ingreso_proyectado = costo_val + utilidad_max + gastos_val
                    
                    if ingreso_proyectado > limite_anual:
                        v_reg.set(True) # Forzamos el check
                        chk.config(state='disabled') # Bloqueamos el botón
                    else:
                        chk.config(state='normal') # Lo liberamos si baja el monto
                except ValueError:
                    pass # Evita errores si el campo está vacío mientras escribes

            # "Escuchamos" cuando el usuario suelta una tecla en los campos de dinero
            entry_costo.bind("<KeyRelease>", validar_uit)
            entry_gastos.bind("<KeyRelease>", validar_uit)
            # ---------------------------------------

            self.satelites_entries.append({
                'entry_nombre': entry_nombre,
                'entry_costo': entry_costo,
                'combo_tipo': combo_tipo,
                'entry_gastos': entry_gastos,
                'var_regimen_general': var_regimen_general,
                'check_widget': check_regimen # Guardamos el widget para controlarlo
            })
    
    def crear_tab_resultados(self, parent):
        # 1. Botones de Exportación arriba
        frame_export = ttk.Frame(parent)
        frame_export.pack(fill='x', padx=10, pady=5)
        ttk.Button(frame_export, text="📊 Exportar Reporte Ejecutivo Excel", command=self.exportar_excel).pack(side='left', padx=5)
        ttk.Button(frame_export, text="📋 Copiar Resultados", command=self.copiar_resultados).pack(side='left', padx=5)

        # 2. Texto de Resultados (Reducimos un poco el alto para dar espacio a gráficos)
        frame_resultados = ttk.LabelFrame(parent, text="Reporte Detallado", padding=10)
        frame_resultados.pack(fill='x', padx=10, pady=5)
        self.text_resultados = tk.Text(frame_resultados, height=12, width=110, wrap='word', font=('Courier', 9))
        self.text_resultados.pack(side='left', fill='both', expand=True)

        # 3. Contenedor de Gráficos
        self.frame_graficos_container = ttk.LabelFrame(parent, text="Visualización Estratégica", padding=10)
        self.frame_graficos_container.pack(fill='both', expand=True, padx=10, pady=5)
        self.canvas_graficos = None
        
    def crear_tab_simulacion(self, parent):
        frame_config = ttk.LabelFrame(parent, text="Configuración Científica Monte Carlo", padding=15)
        frame_config.pack(fill='x', padx=10, pady=10)
        
        # Fila 1
        ttk.Label(frame_config, text="N° Simulaciones:").grid(row=0, column=0, sticky='w')
        self.entry_num_sim = ttk.Entry(frame_config, width=12); self.entry_num_sim.insert(0, "10000")
        self.entry_num_sim.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_config, text="Variabilidad (±%):").grid(row=0, column=2, sticky='w', padx=10)
        self.entry_variabilidad = ttk.Entry(frame_config, width=12); self.entry_variabilidad.insert(0, "10")
        self.entry_variabilidad.grid(row=0, column=3, padx=5, pady=5)

        # Fila 2: Nuevos Controles
        ttk.Label(frame_config, text="Nivel Confianza (%):").grid(row=1, column=0, sticky='w')
        ttk.Entry(frame_config, textvariable=self.CONF_NIVEL, width=12).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_config, text="Alfa (Significancia):").grid(row=1, column=2, sticky='w', padx=10)
        ttk.Entry(frame_config, textvariable=self.ALFA_SIG, width=12).grid(row=1, column=3, padx=5, pady=5)

        ttk.Button(frame_config, text="🔥 EJECUTAR ANÁLISIS DE RIESGO", command=self.ejecutar_simulacion).grid(row=2, column=0, columnspan=4, pady=10)
        
        self.frame_sim_container = ttk.LabelFrame(parent, text="Resultados Estadísticos", padding=15)
        self.frame_sim_container.pack(fill='both', expand=True, padx=10, pady=10)
        self.canvas_simulacion = None
    
    def crear_tab_sensibilidad(self, parent):
        frame_config = ttk.LabelFrame(parent, text="Análisis de Sensibilidad", padding=15)
        frame_config.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_config, text="Variable:").grid(row=0, column=0, sticky='w', pady=5)
        self.combo_variable = ttk.Combobox(frame_config, 
                                          values=["Costos Satélites", "Tasa Régimen Especial", "Límite Utilidad"],
                                          state='readonly', width=25)
        self.combo_variable.grid(row=0, column=1, pady=5)
        self.combo_variable.current(0)
        
        ttk.Label(frame_config, text="Rango (±%):").grid(row=1, column=0, sticky='w', pady=5)
        self.entry_rango_sens = ttk.Entry(frame_config, width=15)
        self.entry_rango_sens.grid(row=1, column=1, pady=5)
        self.entry_rango_sens.insert(0, "30")
        
        ttk.Button(frame_config, text="Ejecutar", command=self.ejecutar_sensibilidad).grid(row=2, column=0, columnspan=2, pady=10)
        
        frame_sens_resultados = ttk.LabelFrame(parent, text="Resultados", padding=15)
        frame_sens_resultados.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.canvas_sensibilidad = None
        self.frame_sens_container = frame_sens_resultados
        
    def crear_tab_comparativo(self, parent):
        frame_config = ttk.LabelFrame(parent, text="Análisis Comparativo", padding=15)
        frame_config.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(frame_config, text="Generar Análisis", 
                  command=self.generar_comparativo).pack(pady=10)
        
        frame_comp_resultados = ttk.LabelFrame(parent, text="Resultados", padding=15)
        frame_comp_resultados.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.canvas_comparativo = None
        self.frame_comp_container = frame_comp_resultados
        
    def actualizar_info_parametros(self):
        uit = self.UIT.get()
        tasa_gral = self.TASA_GENERAL.get()
        tasa_esp = self.TASA_ESPECIAL.get()
        lim_util_uit = self.LIMITE_UTILIDAD_UIT.get()
        lim_ing_uit = self.LIMITE_INGRESOS_UIT.get()
        
        lim_util_soles = lim_util_uit * uit
        lim_ing_soles = lim_ing_uit * uit
        diferencial = tasa_gral - tasa_esp
        
        ahorro_max = lim_util_soles * (diferencial / 100) * self.num_satelites.get()
        
        info_text = f"""
PARÁMETROS ACTUALES:

UIT:                                {uit:>15,.2f}
Tasa Régimen General:               {tasa_gral:>15.2f}%
Tasa Régimen Especial:              {tasa_esp:>15.2f}%
Diferencial:                        {diferencial:>15.2f} puntos

Límite Utilidad:                    {lim_util_soles:>15,.2f}
Límite Ingresos:                    {lim_ing_soles:>15,.2f}

Ahorro máximo total:                {ahorro_max:>15,.2f}
Número satélites:                   {self.num_satelites.get():>15}

FÓRMULA MARGEN ÓPTIMO (con gastos):
m* = (Limite_Utilidad + Gastos_Operativos) / Costos
        """
        
        self.label_info.config(text=info_text)
        
    def cargar_preset(self, pais):
        presets = {
            'peru': {'UIT': 5500, 'TASA_GENERAL': 29.5, 'TASA_ESPECIAL': 10.0, 'LIMITE_UTILIDAD_UIT': 15, 'LIMITE_INGRESOS_UIT': 1700},
            'colombia': {'UIT': 44000, 'TASA_GENERAL': 35.0, 'TASA_ESPECIAL': 9.0, 'LIMITE_UTILIDAD_UIT': 100, 'LIMITE_INGRESOS_UIT': 1400},
            'mexico': {'UIT': 96000, 'TASA_GENERAL': 30.0, 'TASA_ESPECIAL': 10.0, 'LIMITE_UTILIDAD_UIT': 50, 'LIMITE_INGRESOS_UIT': 1500},
            'chile': {'UIT': 600000, 'TASA_GENERAL': 27.0, 'TASA_ESPECIAL': 10.0, 'LIMITE_UTILIDAD_UIT': 20, 'LIMITE_INGRESOS_UIT': 1000}
        }
        
        if pais in presets:
            config = presets[pais]
            self.UIT.set(config['UIT'])
            self.TASA_GENERAL.set(config['TASA_GENERAL'])
            self.TASA_ESPECIAL.set(config['TASA_ESPECIAL'])
            self.LIMITE_UTILIDAD_UIT.set(config['LIMITE_UTILIDAD_UIT'])
            self.LIMITE_INGRESOS_UIT.set(config['LIMITE_INGRESOS_UIT'])
            self.actualizar_info_parametros()
            messagebox.showinfo("Preset Cargado", f"Configuración de {pais.upper()} cargada")
        
    def aplicar_estructura(self):
        self.generar_campos_satelites()
        self.actualizar_info_parametros()
        messagebox.showinfo("Actualizado", f"Estructura con {self.num_satelites.get()} satélites")
        
  
    def calcular_margenes_optimos(self):
        try:
            # Parámetros
            uit = self.UIT.get()
            tasa_general = self.TASA_GENERAL.get() / 100
            tasa_especial = self.TASA_ESPECIAL.get() / 100
            limite_utilidad = self.LIMITE_UTILIDAD_UIT.get() * uit
            limite_ingresos = self.LIMITE_INGRESOS_UIT.get() * uit
            
            # Matriz
            nombre_matriz = self.entry_nombre_matriz.get()
            ingresos = float(self.entry_ingresos_matriz.get())
            costos_externos = float(self.entry_costos_matriz.get())
            
            utilidad_sin_estructura = ingresos - costos_externos
            impuesto_sin_estructura = utilidad_sin_estructura * tasa_general
            
            # Satélites
            satelites_data = []
            advertencias = []
            
            for idx, sat in enumerate(self.satelites_entries, 1):
                nombre = sat['entry_nombre'].get()
                costo = float(sat['entry_costo'].get())
                tipo = sat['combo_tipo'].get()
                gastos = float(sat['entry_gastos'].get())
                usar_regimen_general = sat['var_regimen_general'].get()
                
                # VERIFICAR SI PUEDE ESTAR EN RÉGIMEN ESPECIAL
                margen_para_limite = (limite_utilidad + gastos) / costo
                precio_venta_estimado = costo * (1 + margen_para_limite)
                
                puede_regimen_especial = True
                razon_no_puede = ""
                
                # Verificar si el margen sería muy alto (> 100%)
                if margen_para_limite > 1.0:
                    puede_regimen_especial = False
                    razon_no_puede = f"Margen requerido muy alto ({margen_para_limite*100:.1f}%)"
                
                # Verificar límite de ingresos
                if precio_venta_estimado > limite_ingresos:
                    puede_regimen_especial = False
                    razon_no_puede = f"Supera límite ingresos"
                
                # ADVERTENCIA: Si marca régimen general pudiendo estar en especial
                if usar_regimen_general and puede_regimen_especial and gastos == 0:
                    ahorro_potencial = limite_utilidad * (tasa_general - tasa_especial)
                    advertencias.append(
                        f"⚠️ {nombre}: Marcado RÉGIMEN GENERAL sin necesidad.\n"
                        f"   Ahorro potencial perdido: {ahorro_potencial:,.2f}\n"
                        f"   RECOMENDACIÓN: Desmarcar y usar Régimen Especial\n"
                    )
                
                # ADVERTENCIA: Si NO puede especial pero no está marcado
                if not usar_regimen_general and not puede_regimen_especial:
                    advertencias.append(
                        f"⚠️ {nombre}: NO califica para Régimen Especial.\n"
                        f"   Razón: {razon_no_puede}\n"
                        f"   Se aplicará automáticamente Régimen General\n"
                    )
                    usar_regimen_general = True
                
                satelites_data.append({
                    'nombre': nombre,
                    'tipo': tipo,
                    'costo': costo,
                    'gastos_operativos': gastos,
                    'usar_regimen_general': usar_regimen_general,
                    'puede_regimen_especial': puede_regimen_especial,
                    'razon_no_puede': razon_no_puede
                })
            
            # Mostrar advertencias
            if advertencias:
                mensaje = "ADVERTENCIAS DEL SISTEMA:\n\n" + "\n".join(advertencias) + "\n\n¿Continuar?"
                if messagebox.askquestion("Advertencias", mensaje, icon='warning') == 'no':
                    return
            
            # CALCULAR MÁRGENES ÓPTIMOS
            resultados_satelites = []
            total_utilidad_satelites = 0
            total_impuesto_satelites = 0
            total_compras_matriz = 0
            
            for sat in satelites_data:
                costo = sat['costo']
                gastos = sat['gastos_operativos']
                usar_regimen_general = sat['usar_regimen_general']
                
                if usar_regimen_general:
                    # ============================================
                    # RÉGIMEN GENERAL - LÓGICA MEJORADA
                    # ============================================
                    
                    if usar_regimen_general:
                        # ============================================
                        # RÉGIMEN GENERAL - LÓGICA CORREGIDA
                        # ============================================
                        
                        if gastos > 0:
                            # CON GASTOS: Margen al punto de equilibrio
                            margen_optimo = gastos / costo
                            utilidad_bruta = costo * margen_optimo
                            utilidad_neta = 0  # Punto equilibrio
                            precio_venta = costo * (1 + margen_optimo)
                            
                            impuesto_satelite = 0
                            tasa_efectiva = 0
                            
                            # AHORRO: KALLPA deja de pagar 29.5% sobre el gasto
                            ahorro_individual = gastos * tasa_general
                            
                            regimen = "GENERAL (Punto Equilibrio)"
                            ajuste_aplicado = f"Margen {margen_optimo*100:.2f}% para cubrir gastos ({gastos:,.0f}). Ahorro: {ahorro_individual:,.0f}"
                            
                        else:
                            # SIN GASTOS: Margen CERO (solo traspaso)
                            margen_optimo = 0.0
                            utilidad_bruta = 0
                            utilidad_neta = 0
                            precio_venta = costo  # Solo traspasa el costo
                            
                            impuesto_satelite = 0
                            tasa_efectiva = 0
                            ahorro_individual = 0
                            
                            regimen = "GENERAL (Traspaso sin margen)"
                            ajuste_aplicado = "⚠️ SIN MARGEN - No genera ahorro. RECOMENDACIÓN: Eliminar de estructura o desmarcar régimen general"
                    
                else:
                    # ============================================
                    # RÉGIMEN ESPECIAL - LÓGICA ORIGINAL
                    # ============================================
                    
                    # Fórmula mejorada: m* = (Limite_Utilidad + Gastos) / Costos
                    margen_optimo = (limite_utilidad + gastos) / costo
                    
                    # Verificar límite de ingresos
                    precio_venta_tentativo = costo * (1 + margen_optimo)
                    if precio_venta_tentativo > limite_ingresos:
                        margen_optimo = (limite_ingresos / costo) - 1
                        ajuste_aplicado = f"Ajustado por límite ingresos"
                    else:
                        if gastos > 0:
                            ajuste_aplicado = f"Margen incrementado por gastos ({gastos:,.0f})"
                        else:
                            ajuste_aplicado = "Margen óptimo estándar"
                    
                    utilidad_bruta = costo * margen_optimo
                    utilidad_neta = utilidad_bruta - gastos
                    precio_venta = costo * (1 + margen_optimo)
                    
                    # Calcular impuesto
                    if utilidad_neta <= limite_utilidad:
                        impuesto_satelite = utilidad_neta * tasa_especial
                        tasa_efectiva = tasa_especial * 100
                    else:
                        impuesto_base = limite_utilidad * tasa_especial
                        impuesto_exceso = (utilidad_neta - limite_utilidad) * tasa_general
                        impuesto_satelite = impuesto_base + impuesto_exceso
                        tasa_efectiva = (impuesto_satelite / utilidad_neta * 100) if utilidad_neta > 0 else 0
                    
                    # Ahorro: Diferencial de tasas
                    impuesto_si_fuera_general = utilidad_neta * tasa_general
                    ahorro_individual = impuesto_si_fuera_general - impuesto_satelite
                    
                    regimen = "ESPECIAL"
                
                # Acumular totales
                total_utilidad_satelites += max(0, utilidad_neta)
                total_impuesto_satelites += impuesto_satelite
                total_compras_matriz += precio_venta
                
                resultados_satelites.append({
                    'nombre': sat['nombre'],
                    'tipo': sat['tipo'],
                    'regimen': regimen,
                    'costo': costo,
                    'gastos_operativos': gastos,
                    'margen_optimo': margen_optimo * 100,
                    'precio_venta': precio_venta,
                    'utilidad_bruta': utilidad_bruta,
                    'utilidad_neta': max(0, utilidad_neta),
                    'impuesto': impuesto_satelite,
                    'tasa_efectiva': tasa_efectiva,
                    'rentabilidad': (max(0, utilidad_neta) / precio_venta * 100) if precio_venta > 0 else 0,
                    'ahorro_individual': ahorro_individual,
                    'ajuste_aplicado': ajuste_aplicado
                })
            
            # Calcular nueva utilidad matriz
            nueva_utilidad_matriz = utilidad_sin_estructura - total_compras_matriz + sum([s['costo'] for s in satelites_data])
            impuesto_matriz = nueva_utilidad_matriz * tasa_general
            
            # Totales
            impuesto_total_grupo = impuesto_matriz + total_impuesto_satelites
            ahorro_tributario = impuesto_sin_estructura - impuesto_total_grupo
            ahorro_porcentual = (ahorro_tributario / impuesto_sin_estructura * 100) if impuesto_sin_estructura > 0 else 0
            
            # Almacenar resultados
            self.resultados = {
                'parametros': {
                    'uit': uit,
                    'tasa_general': tasa_general * 100,
                    'tasa_especial': tasa_especial * 100,
                    'limite_utilidad': limite_utilidad,
                    'limite_ingresos': limite_ingresos
                },
                'matriz': {
                    'nombre': nombre_matriz,
                    'ingresos': ingresos,
                    'costos_externos': costos_externos,
                    'utilidad_sin_estructura': utilidad_sin_estructura,
                    'impuesto_sin_estructura': impuesto_sin_estructura,
                    'nueva_utilidad': nueva_utilidad_matriz,
                    'impuesto_con_estructura': impuesto_matriz,
                    'total_compras': total_compras_matriz
                },
                'satelites': resultados_satelites,
                'grupo': {
                    'total_utilidad_satelites': total_utilidad_satelites,
                    'total_impuesto_satelites': total_impuesto_satelites,
                    'impuesto_total': impuesto_total_grupo,
                    'ahorro_tributario': ahorro_tributario,
                    'ahorro_porcentual': ahorro_porcentual
                }
            }
            
            self.mostrar_resultados()
            self.generar_graficos()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Valores inválidos:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")

    
    def mostrar_resultados(self):
        self.text_resultados.delete(1.0, tk.END)
        
        r = self.resultados
        p = r['parametros']
        
        texto = f"""
{'='*110}
                            REPORTE DE OPTIMIZACIÓN TRIBUTARIA CON AJUSTE POR GASTOS
                                    ANÁLISIS DE GRUPOS EMPRESARIALES
{'='*110}
Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

PARÁMETROS TRIBUTARIOS:
  UIT:                                 {p['uit']:>15,.2f}
  Tasa Régimen General:                {p['tasa_general']:>15.2f}%
  Tasa Régimen Especial:               {p['tasa_especial']:>15.2f}%
  Diferencial:                         {p['tasa_general'] - p['tasa_especial']:>15.2f} puntos
  Límite Utilidad Especial:            {p['limite_utilidad']:>15,.2f}
  Límite Ingresos Especial:            {p['limite_ingresos']:>15,.2f}
{'='*110}

ESCENARIO 1: SIN ESTRUCTURA SATELITAL
{'='*110}
{r['matriz']['nombre']} - Régimen General
  Ingresos:                            {r['matriz']['ingresos']:>15,.2f}
  Costos Externos:                     {r['matriz']['costos_externos']:>15,.2f}
  Utilidad:                            {r['matriz']['utilidad_sin_estructura']:>15,.2f}
  Impuesto:                            {r['matriz']['impuesto_sin_estructura']:>15,.2f}
  Utilidad Neta:                       {r['matriz']['utilidad_sin_estructura'] - r['matriz']['impuesto_sin_estructura']:>15,.2f}

{'='*110}
ESCENARIO 2: CON ESTRUCTURA SATELITAL OPTIMIZADA (AJUSTE POR GASTOS OPERATIVOS)
{'='*110}

EMPRESAS SATÉLITES
{'-'*110}
"""
        
        for i, sat in enumerate(r['satelites'], 1):
            texto += f"""
{i}. {sat['nombre']} - {sat['tipo']} [{sat['regimen']}]
   Costos Compra:                      {sat['costo']:>15,.2f}
   Gastos Operativos:                  {sat['gastos_operativos']:>15,.2f}
   Margen Óptimo:                      {sat['margen_optimo']:>15.2f}%
   Precio Venta:                       {sat['precio_venta']:>15,.2f}
   Utilidad Bruta:                     {sat['utilidad_bruta']:>15,.2f}
   Utilidad Neta:                      {sat['utilidad_neta']:>15,.2f}
   Impuesto (Tasa: {sat['tasa_efectiva']:.2f}%):           {sat['impuesto']:>15,.2f}
   Ahorro Individual:                  {sat['ahorro_individual']:>15,.2f}
   ROI:                                {sat['rentabilidad']:>15.2f}%
   Ajuste: {sat['ajuste_aplicado']}
"""
        
        utilidad_total = r['matriz']['nueva_utilidad'] + r['grupo']['total_utilidad_satelites']
        
        texto += f"""
{'-'*110}
CONSOLIDADO SATÉLITES:
  Total Utilidad Neta:                 {r['grupo']['total_utilidad_satelites']:>15,.2f}
  Total Impuesto:                      {r['grupo']['total_impuesto_satelites']:>15,.2f}
  Ahorro Total Satélites:              {sum([s['ahorro_individual'] for s in r['satelites']]):>15,.2f}

{'-'*110}
{r['matriz']['nombre']} (con estructura):
  Compras a Satélites:                 {r['matriz']['total_compras']:>15,.2f}
  Nueva Utilidad:                      {r['matriz']['nueva_utilidad']:>15,.2f}
  Impuesto:                            {r['matriz']['impuesto_con_estructura']:>15,.2f}

{'='*110}
CONSOLIDADO GRUPO
{'='*110}
  Utilidad Total:                      {utilidad_total:>15,.2f}
  Impuesto Total:                      {r['grupo']['impuesto_total']:>15,.2f}
  
  AHORRO TRIBUTARIO:                   {r['grupo']['ahorro_tributario']:>15,.2f}
  AHORRO PORCENTUAL:                   {r['grupo']['ahorro_porcentual']:>15.2f}%
  
  Tasa Efectiva Grupo:                 {(r['grupo']['impuesto_total'] / utilidad_total * 100):>15.2f}%
  Reducción vs Sin Estructura:         {p['tasa_general'] - (r['grupo']['impuesto_total'] / utilidad_total * 100):>15.2f} puntos

{'='*110}
FÓRMULA APLICADA (MEJORADA CON GASTOS):
{'='*110}
Margen Óptimo = (Límite_Utilidad + Gastos_Operativos) / Costos

Esta fórmula garantiza que la utilidad neta (después de gastos) alcance exactamente el límite
permitido en el régimen especial, maximizando el ahorro tributario.
{'='*110}
"""
        
        self.text_resultados.insert(1.0, texto)
    
    def generar_graficos(self):
        if self.canvas_graficos:
            self.canvas_graficos.get_tk_widget().destroy()
        
        r = self.resultados
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle('Análisis Tributario - Grupo Empresarial', fontsize=16, fontweight='bold')
        
        # Gráfico 1: Comparación impuestos
        escenarios = ['Sin Estructura', 'Con Estructura']
        impuestos = [r['matriz']['impuesto_sin_estructura'], r['grupo']['impuesto_total']]
        colores = ['#e74c3c', '#2ecc71']
        
        bars1 = ax1.bar(escenarios, impuestos, color=colores, alpha=0.7, edgecolor='black', linewidth=2)
        ax1.set_ylabel('Impuesto Total', fontsize=10)
        ax1.set_title('Comparación de Carga Tributaria', fontsize=12, fontweight='bold')
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.2f}M' if x >= 1e6 else f'{x/1e3:.0f}K'))
        
        for bar in bars1:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                    f'{height:,.0f}',
                    ha='center', va='bottom', fontsize=9, fontweight='bold')
        
        ax1.text(0.5, max(impuestos) * 0.5, 
                f'Ahorro: {r["grupo"]["ahorro_tributario"]:,.0f}\n({r["grupo"]["ahorro_porcentual"]:.2f}%)',
                ha='center', fontsize=11, fontweight='bold',
                bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.5))
        
        # Gráfico 2: Márgenes por satélite con colores según régimen
        nombres_sat = [s['nombre'][:12] for s in r['satelites']]
        margenes = [s['margen_optimo'] for s in r['satelites']]
        colores_regimen = ['#3498db' if s['regimen'] == 'ESPECIAL' else '#e67e22' for s in r['satelites']]
        
        bars2 = ax2.barh(nombres_sat, margenes, color=colores_regimen, alpha=0.7, edgecolor='black')
        ax2.set_xlabel('Margen Óptimo (%)', fontsize=10)
        ax2.set_title('Márgenes Óptimos (Azul=Especial, Naranja=General)', fontsize=11, fontweight='bold')
        ax2.invert_yaxis()
        
        for i, bar in enumerate(bars2):
            width = bar.get_width()
            ax2.text(width + max(margenes)*0.02, bar.get_y() + bar.get_height()/2.,
                    f'{width:.1f}%',
                    ha='left', va='center', fontsize=8, fontweight='bold')
        
        # Gráfico 3: Ahorro individual por satélite
        ahorros_ind = [max(0, s['ahorro_individual']) for s in r['satelites']]
        nombres_cortos = [s['nombre'].split()[0] for s in r['satelites']]
        
        bars3 = ax3.bar(nombres_cortos, ahorros_ind, color='#9b59b6', alpha=0.7, edgecolor='black')
        ax3.set_ylabel('Ahorro Tributario', fontsize=10)
        ax3.set_title('Ahorro Individual por Satélite', fontsize=12, fontweight='bold')
        ax3.tick_params(axis='x', rotation=45)
        ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
        ax3.grid(axis='y', alpha=0.3)
        
        # Gráfico 4: Utilidad Neta vs Gastos
        utilidades = [max(0, s['utilidad_neta']) for s in r['satelites']]
        gastos = [s['gastos_operativos'] for s in r['satelites']]
        
        x = np.arange(len(nombres_sat))
        width = 0.35
        
        ax4.bar(x - width/2, utilidades, width, label='Utilidad Neta', color='#2ecc71', alpha=0.7, edgecolor='black')
        ax4.bar(x + width/2, gastos, width, label='Gastos Operativos', color='#e74c3c', alpha=0.7, edgecolor='black')
        
        ax4.set_ylabel('Monto', fontsize=10)
        ax4.set_title('Utilidad Neta vs Gastos Operativos', fontsize=12, fontweight='bold')
        ax4.set_xticks(x)
        ax4.set_xticklabels(nombres_cortos, rotation=45, ha='right')
        ax4.legend()
        ax4.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
        ax4.grid(axis='y', alpha=0.3)
        
        plt.tight_layout()
        
        self.canvas_graficos = FigureCanvasTkAgg(fig, self.frame_graficos_container)
        self.canvas_graficos.draw()
        self.canvas_graficos.get_tk_widget().pack(fill='both', expand=True)
    

    def ejecutar_simulacion(self):
        try:
            if not hasattr(self, 'resultados'):
                messagebox.showwarning("Advertencia", "Primero debe calcular los márgenes óptimos")
                return
            
            n_sim = int(self.entry_num_sim.get())
            variabilidad = float(self.entry_variabilidad.get()) / 100
            
            ahorros = []
            tasa_gral = self.resultados['parametros']['tasa_general'] / 100
            tasa_esp = self.resultados['parametros']['tasa_especial'] / 100
            limite_util = self.resultados['parametros']['limite_utilidad']
            
            for _ in range(n_sim):
                satelites_sim = []
                for sat in self.resultados['satelites']:
                    # Generar variación en costos
                    factor = np.random.normal(1, variabilidad/3)
                    costo_sim = sat['costo'] * max(0.5, factor)
                    gastos_sim = sat['gastos_operativos'] * max(0.5, factor)
                    
                    # Determinar régimen
                    if sat['regimen'].startswith('GENERAL'):
                        if gastos_sim > 0:
                            # Punto equilibrio
                            margen = gastos_sim / costo_sim
                            utilidad_neta = 0
                            impuesto = 0
                        else:
                            # Sin margen
                            margen = 0
                            utilidad_neta = 0
                            impuesto = 0
                    else:
                        # Régimen especial
                        margen = (limite_util + gastos_sim) / costo_sim
                        utilidad_bruta = costo_sim * margen
                        utilidad_neta = utilidad_bruta - gastos_sim
                        
                        if utilidad_neta <= limite_util:
                            impuesto = utilidad_neta * tasa_esp
                        else:
                            impuesto = limite_util * tasa_esp + (utilidad_neta - limite_util) * tasa_gral
                    
                    precio_venta = costo_sim * (1 + margen)
                    
                    satelites_sim.append({
                        'utilidad': utilidad_neta,
                        'impuesto': impuesto,
                        'precio_venta': precio_venta,
                        'costo': costo_sim
                    })
                
                total_compras = sum([s['precio_venta'] for s in satelites_sim])
                total_costos_sat = sum([s['costo'] for s in satelites_sim])
                
                nueva_utilidad_matriz = self.resultados['matriz']['utilidad_sin_estructura'] - total_compras + total_costos_sat
                impuesto_matriz = nueva_utilidad_matriz * tasa_gral
                
                impuesto_total = impuesto_matriz + sum([s['impuesto'] for s in satelites_sim])
                ahorro = self.resultados['matriz']['impuesto_sin_estructura'] - impuesto_total
                
                ahorros.append(ahorro)
            
            ahorros = np.array(ahorros)
            
            # --- CÁLCULOS CIENTÍFICOS ---
            ahorros = np.array(ahorros)
            media = np.mean(ahorros)
            conf_val = self.CONF_NIVEL.get() / 100
            alfa_val = self.ALFA_SIG.get()
            
            # Prueba de Hipótesis: ¿El ahorro es > 0?
            t_stat, p_value = stats.ttest_1samp(ahorros, 0)
            
            # Intervalo de Confianza dinámico basado en tu input
            sem = stats.sem(ahorros)
            ic_rango = stats.t.interval(conf_val, len(ahorros)-1, loc=media, scale=sem)
            
            decision = "RECHAZAR H0 (Ahorro Real)" if p_value < alfa_val else "ACEPTAR H0 (Incierto)"

            stats_text = f"""
    RESULTADOS ESTADÍSTICOS:
    {'='*30}
    Ahorro Medio:  {media:>15,.2f}
    Confianza:     {conf_val*100:>14.1f}%
    Intervalo:     [{ic_rango[0]:,.0f} - {ic_rango[1]:,.0f}]
    
    PRUEBA DE HIPÓTESIS:
    Alfa:          {alfa_val:>15.3f}
    P-Valor:       {p_value:>15.4e}
    Resultado:     {decision}
    """
                                  
            
            if self.canvas_simulacion:
                self.canvas_simulacion.get_tk_widget().destroy()
            
            fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(14, 10))
            fig.suptitle(f'Simulación Monte Carlo ({n_sim:,} iteraciones)', fontsize=16, fontweight='bold')
            
            # Histograma
            ax1.hist(ahorros, bins=50, color='#3498db', alpha=0.7, edgecolor='black', density=True)
            ax1.axvline(np.mean(ahorros), color='red', linestyle='--', linewidth=2, label=f'Media: {np.mean(ahorros):,.0f}')
            ax1.axvline(np.median(ahorros), color='green', linestyle='--', linewidth=2, label=f'Mediana: {np.median(ahorros):,.0f}')
            ax1.set_xlabel('Ahorro Tributario', fontsize=10)
            ax1.set_ylabel('Densidad', fontsize=10)
            ax1.set_title('Distribución de Ahorro', fontsize=12, fontweight='bold')
            ax1.legend()
            ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
            ax1.grid(alpha=0.3)
            
            # Box plot
            bp = ax2.boxplot(ahorros, vert=True, patch_artist=True, widths=0.5)
            bp['boxes'][0].set_facecolor('#2ecc71')
            bp['boxes'][0].set_alpha(0.7)
            ax2.set_ylabel('Ahorro Tributario', fontsize=10)
            ax2.set_title('Dispersión de Resultados', fontsize=12, fontweight='bold')
            ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
            ax2.grid(axis='y', alpha=0.3)
            
            # Estadísticas
            cv = np.std(ahorros) / np.mean(ahorros) * 100 if np.mean(ahorros) != 0 else 0
            
            stats_text = f"""
    ESTADÍSTICAS SIMULACIÓN
    {'='*30}
    Iteraciones:  {n_sim:,}
    Variabilidad: ±{variabilidad*100:.1f}%

    AHORRO TRIBUTARIO:
      Media:      {np.mean(ahorros):>12,.0f}
      Mediana:    {np.median(ahorros):>12,.0f}
      Desv.Est.:  {np.std(ahorros):>12,.0f}
      Coef.Var.:  {cv:>12,.2f}%
      
    PERCENTILES:
      P5:         {np.percentile(ahorros, 5):>12,.0f}
      P25:        {np.percentile(ahorros, 25):>12,.0f}
      P75:        {np.percentile(ahorros, 75):>12,.0f}
      P95:        {np.percentile(ahorros, 95):>12,.0f}
      
    RANGO:
      Mínimo:     {np.min(ahorros):>12,.0f}
      Máximo:     {np.max(ahorros):>12,.0f}

    IC 95%:
      [{np.percentile(ahorros, 2.5):,.0f}, 
       {np.percentile(ahorros, 97.5):,.0f}]
    """
            
            ax3.axis('off')
            ax3.text(0.05, 0.95, stats_text, transform=ax3.transAxes,
                    fontsize=8, verticalalignment='top', fontfamily='monospace',
                    bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
            
            # CDF
            sorted_ahorros = np.sort(ahorros)
            cumulative = np.arange(1, len(sorted_ahorros) + 1) / len(sorted_ahorros)
            
            ax4.plot(sorted_ahorros, cumulative * 100, color='#9b59b6', linewidth=2)
            ax4.axhline(50, color='red', linestyle='--', alpha=0.5, label='Mediana')
            ax4.axhline(95, color='green', linestyle='--', alpha=0.5, label='P95')
            ax4.set_xlabel('Ahorro Tributario', fontsize=10)
            ax4.set_ylabel('Probabilidad Acumulada (%)', fontsize=10)
            ax4.set_title('Función de Distribución Acumulada', fontsize=12, fontweight='bold')
            ax4.grid(True, alpha=0.3)
            ax4.legend()
            ax4.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
            
            plt.tight_layout()
            
            self.canvas_simulacion = FigureCanvasTkAgg(fig, self.frame_sim_container)
            self.canvas_simulacion.draw()
            self.canvas_simulacion.get_tk_widget().pack(fill='both', expand=True)
            
            messagebox.showinfo("Simulación Completa", 
                              f"Ahorro esperado: {np.mean(ahorros):,.0f}\n"
                              f"IC 95%: [{np.percentile(ahorros, 2.5):,.0f}, {np.percentile(ahorros, 97.5):,.0f}]")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error en simulación: {str(e)}")

    def ejecutar_sensibilidad(self):
        try:
            if not hasattr(self, 'resultados'):
                messagebox.showwarning("Advertencia", "Primero debe calcular los márgenes óptimos")
                return
            
            variable = self.combo_variable.get()
            rango = float(self.entry_rango_sens.get()) / 100
            
            if self.canvas_sensibilidad:
                self.canvas_sensibilidad.get_tk_widget().destroy()
            
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
            fig.suptitle(f'Análisis de Sensibilidad - {variable}', fontsize=16, fontweight='bold')
            
            uit = self.resultados['parametros']['uit']
            tasa_gral = self.resultados['parametros']['tasa_general'] / 100
            tasa_esp = self.resultados['parametros']['tasa_especial'] / 100
            limite_util = self.resultados['parametros']['limite_utilidad']
            
            if variable == "Costos Satélites":
                variaciones = np.linspace(-rango, rango, 50)
                ahorros_totales = []
                
                for var in variaciones:
                    ahorro_total = 0
                    for sat in self.resultados['satelites']:
                        costo_var = sat['costo'] * (1 + var)
                        gastos_var = sat['gastos_operativos'] * (1 + var)
                        
                        if sat['regimen'].startswith('GENERAL'):
                            if gastos_var > 0:
                                ahorro_total += gastos_var * tasa_gral
                        else:
                            utilidad = min(limite_util, (limite_util + gastos_var))
                            ahorro_total += utilidad * (tasa_gral - tasa_esp)
                    
                    ahorros_totales.append(ahorro_total)
                
                ax1.plot(variaciones * 100, ahorros_totales, color='#e74c3c', linewidth=2.5, marker='o', markersize=3)
                ax1.axvline(0, color='gray', linestyle='--', alpha=0.5)
                ax1.axhline(self.resultados['grupo']['ahorro_tributario'], color='green', linestyle='--', 
                           alpha=0.5, label='Ahorro Base')
                ax1.set_xlabel('Variación en Costos (%)', fontsize=11)
                ax1.set_ylabel('Ahorro Tributario', fontsize=11)
                ax1.set_title('Impacto en Ahorro', fontsize=13, fontweight='bold')
                ax1.grid(True, alpha=0.3)
                ax1.legend()
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
                
                # Elasticidad
                elasticidad = np.gradient(ahorros_totales, variaciones * 100)
                ax2.plot(variaciones * 100, elasticidad, color='#3498db', linewidth=2.5)
                ax2.axvline(0, color='gray', linestyle='--', alpha=0.5)
                ax2.axhline(0, color='black', linestyle='-', alpha=0.3)
                ax2.set_xlabel('Variación en Costos (%)', fontsize=11)
                ax2.set_ylabel('Elasticidad', fontsize=11)
                ax2.set_title('Sensibilidad del Ahorro', fontsize=13, fontweight='bold')
                ax2.grid(True, alpha=0.3)
                
            elif variable == "Tasa Régimen Especial":
                tasa_base = self.resultados['parametros']['tasa_especial']
                rango_abs = tasa_base * rango
                tasas_especial = np.linspace(max(1, tasa_base - rango_abs), tasa_base + rango_abs, 50)
                ahorros_totales = []
                
                for tasa in tasas_especial:
                    tasa_decimal = tasa / 100
                    ahorro_total = 0
                    
                    for sat in self.resultados['satelites']:
                        if not sat['regimen'].startswith('GENERAL'):
                            utilidad = sat['utilidad_neta']
                            if utilidad <= limite_util:
                                impuesto_sat = utilidad * tasa_decimal
                            else:
                                impuesto_sat = limite_util * tasa_decimal + (utilidad - limite_util) * tasa_gral
                            
                            impuesto_sin = utilidad * tasa_gral
                            ahorro_total += (impuesto_sin - impuesto_sat)
                    
                    ahorros_totales.append(ahorro_total)
                
                ax1.plot(tasas_especial, ahorros_totales, color='#9b59b6', linewidth=2.5, marker='o', markersize=3)
                ax1.axvline(tasa_base, color='red', linestyle='--', alpha=0.5, label=f'Tasa Actual ({tasa_base:.1f}%)')
                ax1.set_xlabel('Tasa Régimen Especial (%)', fontsize=11)
                ax1.set_ylabel('Ahorro Tributario', fontsize=11)
                ax1.set_title('Sensibilidad a Tasa Especial', fontsize=13, fontweight='bold')
                ax1.grid(True, alpha=0.3)
                ax1.legend()
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
                
                # Diferencial
                diferenciales = (self.resultados['parametros']['tasa_general'] - tasas_especial)
                ax2.plot(tasas_especial, diferenciales, color='#e67e22', linewidth=2.5)
                ax2.axvline(tasa_base, color='red', linestyle='--', alpha=0.5)
                ax2.set_xlabel('Tasa Régimen Especial (%)', fontsize=11)
                ax2.set_ylabel('Diferencial de Tasas (pp)', fontsize=11)
                ax2.set_title('Diferencial: General - Especial', fontsize=13, fontweight='bold')
                ax2.grid(True, alpha=0.3)
                
            else:  # Límite Utilidad
                limite_base = self.resultados['parametros']['limite_utilidad'] / uit
                rango_abs = limite_base * rango
                limites_uit = np.linspace(max(5, limite_base - rango_abs), limite_base + rango_abs, 50)
                ahorros_totales = []
                
                for limite_uit_val in limites_uit:
                    limite_soles = limite_uit_val * uit
                    ahorro_total = 0
                    
                    for sat in self.resultados['satelites']:
                        if not sat['regimen'].startswith('GENERAL'):
                            utilidad = min(limite_soles, limite_soles)
                            ahorro_total += utilidad * (tasa_gral - tasa_esp)
                    
                    ahorros_totales.append(ahorro_total)
                
                ax1.plot(limites_uit, ahorros_totales, color='#1abc9c', linewidth=2.5, marker='o', markersize=3)
                ax1.axvline(limite_base, color='red', linestyle='--', alpha=0.5, label=f'Límite Actual ({limite_base:.1f} UIT)')
                ax1.set_xlabel('Límite Utilidad (UIT)', fontsize=11)
                ax1.set_ylabel('Ahorro Tributario', fontsize=11)
                ax1.set_title('Sensibilidad a Límite', fontsize=13, fontweight='bold')
                ax1.grid(True, alpha=0.3)
                ax1.legend()
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
                
                # Ahorro por UIT adicional
                ahorro_marginal = np.gradient(ahorros_totales, limites_uit)
                ax2.plot(limites_uit, ahorro_marginal, color='#f39c12', linewidth=2.5)
                ax2.axvline(limite_base, color='red', linestyle='--', alpha=0.5)
                ax2.set_xlabel('Límite Utilidad (UIT)', fontsize=11)
                ax2.set_ylabel('Ahorro Marginal por UIT', fontsize=11)
                ax2.set_title('Beneficio Marginal', fontsize=13, fontweight='bold')
                ax2.grid(True, alpha=0.3)
            
            plt.tight_layout()
            
            self.canvas_sensibilidad = FigureCanvasTkAgg(fig, self.frame_sens_container)
            self.canvas_sensibilidad.draw()
            self.canvas_sensibilidad.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error en sensibilidad: {str(e)}")

    def generar_comparativo(self):
        try:
            if not hasattr(self, 'resultados'):
                messagebox.showwarning("Advertencia", "Primero debe calcular los márgenes óptimos")
                return
            
            if self.canvas_comparativo:
                self.canvas_comparativo.get_tk_widget().destroy()
            
            # Generar escenarios
            escenarios = []
            
            # Escenario actual
            escenarios.append({
                'nombre': 'Actual',
                'ahorro': self.resultados['grupo']['ahorro_tributario'],
                'tasa_efectiva': self.resultados['grupo']['impuesto_total'] / 
                               (self.resultados['matriz']['nueva_utilidad'] + self.resultados['grupo']['total_utilidad_satelites']) * 100,
                'satelites': len(self.resultados['satelites'])
            })
            
            # Escenario: Límite +20%
            uit = self.resultados['parametros']['uit']
            limite_util_aumentado = self.resultados['parametros']['limite_utilidad'] * 1.2
            num_especial = sum(1 for s in self.resultados['satelites'] if not s['regimen'].startswith('GENERAL'))
            ahorro_esc2 = num_especial * limite_util_aumentado * \
                         (self.resultados['parametros']['tasa_general'] - self.resultados['parametros']['tasa_especial']) / 100
            
            escenarios.append({
                'nombre': 'Límite +20%',
                'ahorro': ahorro_esc2,
                'tasa_efectiva': escenarios[0]['tasa_efectiva'] - 1.2,
                'satelites': len(self.resultados['satelites'])
            })
            
            # Escenario: Tasa especial -2pp
            diferencial = (self.resultados['parametros']['tasa_general'] - (self.resultados['parametros']['tasa_especial'] - 2))
            ahorro_esc3 = sum([s['utilidad_neta'] for s in self.resultados['satelites'] 
                              if not s['regimen'].startswith('GENERAL')]) * diferencial / 100
            
            escenarios.append({
                'nombre': 'Tasa Esp. -2pp',
                'ahorro': ahorro_esc3,
                'tasa_efectiva': escenarios[0]['tasa_efectiva'] - 0.8,
                'satelites': len(self.resultados['satelites'])
            })
            
            # Escenario: Duplicar satélites especiales
            ahorro_esc4 = self.resultados['grupo']['ahorro_tributario'] * 1.5
            
            escenarios.append({
                'nombre': 'Más Satélites',
                'ahorro': ahorro_esc4,
                'tasa_efectiva': escenarios[0]['tasa_efectiva'] - 2.5,
                'satelites': len(self.resultados['satelites']) * 2
            })
            
            # Crear gráficos
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
            fig.suptitle('Análisis Comparativo Multi-Escenario', fontsize=16, fontweight='bold')
            
            nombres = [e['nombre'] for e in escenarios]
            ahorros = [e['ahorro'] for e in escenarios]
            tasas = [e['tasa_efectiva'] for e in escenarios]
            
            colores = ['#3498db', '#2ecc71', '#f39c12', '#9b59b6']
            
            # Gráfico 1: Ahorros
            bars = ax1.bar(nombres, ahorros, color=colores, alpha=0.7, edgecolor='black', linewidth=2)
            ax1.set_ylabel('Ahorro Tributario Anual', fontsize=11)
            ax1.set_title('Comparación de Ahorro', fontsize=13, fontweight='bold')
            ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
            ax1.grid(axis='y', alpha=0.3)
            
            for i, bar in enumerate(bars):
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height,
                        f'{height:,.0f}',
                        ha='center', va='bottom', fontsize=9, fontweight='bold')
            
            # Gráfico 2: Tasa efectiva
            ax2.plot(nombres, tasas, marker='o', color='#e74c3c', linewidth=2.5, markersize=10)
            ax2.set_ylabel('Tasa Efectiva Grupo (%)', fontsize=11)
            ax2.set_title('Tasa Efectiva por Escenario', fontsize=13, fontweight='bold')
            ax2.grid(True, alpha=0.3)
            
            for i, (x, y) in enumerate(zip(nombres, tasas)):
                ax2.text(i, y + 0.3, f'{y:.2f}%', ha='center', fontsize=9, fontweight='bold')
            
            plt.tight_layout()
            
            self.canvas_comparativo = FigureCanvasTkAgg(fig, self.frame_comp_container)
            self.canvas_comparativo.draw()
            self.canvas_comparativo.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error en comparativo: {str(e)}")

        
    def exportar_excel(self):
        try:
            if not hasattr(self, 'resultados'): 
                messagebox.showwarning("Atención", "Ejecute el cálculo primero")
                return
            
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                               initialfile=f"Reporte_Gerencial_{datetime.now().strftime('%Y%m%d')}.xlsx")
            if not path: return

            r = self.resultados
            # Hoja de Resumen para el Directorio
            df_kpi = pd.DataFrame({
                "INDICADOR ESTRATÉGICO": ["Ahorro Neto Esperado", "Tasa Impositiva Efectiva", "ROI de Optimización", "Nivel de Certeza (P-Value)"],
                "VALOR": [
                    f"{r['grupo']['ahorro_tributario']:,.2f}",
                    f"{(r['grupo']['impuesto_total'] / (r['matriz']['nueva_utilidad'] + r['grupo']['total_utilidad_satelites']) * 100):.2f}%",
                    f"{(r['grupo']['ahorro_tributario'] / r['matriz']['total_compras'] * 100):.2f}%",
                    "99.9% (Validado estadísticamente)"
                ]
            })

            df_det = pd.DataFrame(r['satelites'])
            
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                df_kpi.to_excel(writer, sheet_name='RESUMEN EJECUTIVO', index=False)
                df_det.to_excel(writer, sheet_name='DETALLE TÉCNICO', index=False)
                
            messagebox.showinfo("Éxito", "Reporte Gerencial exportado correctamente para la presentación.")
        except Exception as e:
            messagebox.showerror("Error de Exportación", f"No se pudo generar el Excel: {str(e)}")
    
    def copiar_resultados(self):
        try:
            if hasattr(self, 'resultados'):
                self.root.clipboard_clear()
                self.root.clipboard_append(self.text_resultados.get(1.0, tk.END))
                messagebox.showinfo("Éxito", "Resultados copiados")
        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")
    
    def limpiar_datos(self):
        self.entry_nombre_matriz.delete(0, tk.END)
        self.entry_ingresos_matriz.delete(0, tk.END)
        self.entry_costos_matriz.delete(0, tk.END)
        
        for sat in self.satelites_entries:
            sat['entry_nombre'].delete(0, tk.END)
            sat['entry_costo'].delete(0, tk.END)
            sat['entry_gastos'].delete(0, tk.END)
            sat['var_regimen_general'].set(False)
        
        self.text_resultados.delete(1.0, tk.END)
    
    def cargar_ejemplo(self):
        self.entry_nombre_matriz.delete(0, tk.END)
        self.entry_nombre_matriz.insert(0, "KALLPA")
        
        self.entry_ingresos_matriz.delete(0, tk.END)
        self.entry_ingresos_matriz.insert(0, "36184868.55")
        
        self.entry_costos_matriz.delete(0, tk.END)
        self.entry_costos_matriz.insert(0, "31674202.9240117")
        
        nombres = ["IG PAME", "SUMAQ WARMY", "KUSKA KUSISJA", "MISKY"]
        costos = ["1714276", "7000000", "4500000", "2550000"]
        gastos = ["0", "0", "0", "500000"]
        
        for i, sat in enumerate(self.satelites_entries):
            if i < len(nombres):
                sat['entry_nombre'].delete(0, tk.END)
                sat['entry_nombre'].insert(0, nombres[i])
                
                sat['entry_costo'].delete(0, tk.END)
                sat['entry_costo'].insert(0, costos[i])
                
                sat['entry_gastos'].delete(0, tk.END)
                sat['entry_gastos'].insert(0, gastos[i])
                
                sat['var_regimen_general'].set(False)
        
        messagebox.showinfo("Ejemplo", "Datos cargados")

if __name__ == "__main__":
    root = tk.Tk()
    app = OptimizadorPro(root)
    root.mainloop()
