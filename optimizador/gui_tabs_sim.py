import tkinter as tk
from tkinter import ttk


class GuiSimTabsMixin:
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

        ttk.Button(frame_config, text="EJECUTAR ANALISIS DE RIESGO", command=self.ejecutar_simulacion).grid(row=2, column=0, columnspan=4, pady=10)

        # --- MARCO: INCERTIDUMBRE DE LA MATRIZ Y CORRELACION TERRITORIAL ---
        frame_matriz = ttk.LabelFrame(frame_config, text="Incertidumbre de la Matriz y Correlacion Territorial", padding=10)
        frame_matriz.grid(row=3, column=0, columnspan=4, pady=10, sticky='ew')
        frame_matriz.columnconfigure(1, weight=1)

        # Fila 0: Checkbutton costos matriz
        ttk.Checkbutton(frame_matriz, text="Incluir variabilidad en costos externos de la matriz",
                        variable=self.VAR_MATRIZ_COSTOS).grid(row=0, column=0, columnspan=2, sticky='w', pady=5)

        # Fila 1: Variabilidad costos matriz
        ttk.Label(frame_matriz, text="Variabilidad costos matriz (+/-%):"
                  ).grid(row=1, column=0, sticky='w', padx=20, pady=5)
        self.entry_pct_var_matriz_costos = ttk.Entry(frame_matriz, textvariable=self.PCT_VAR_MATRIZ_COSTOS, width=12)
        self.entry_pct_var_matriz_costos.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        self.entry_pct_var_matriz_costos.config(state='disabled')

        # Fila 2: Checkbutton ingresos matriz
        ttk.Checkbutton(frame_matriz, text="Incluir variabilidad en ingresos de la matriz",
                        variable=self.VAR_MATRIZ_INGRESOS).grid(row=2, column=0, columnspan=2, sticky='w', pady=5)

        # Fila 3: Variabilidad ingresos matriz
        ttk.Label(frame_matriz, text="Variabilidad ingresos matriz (+/-%):"
                  ).grid(row=3, column=0, sticky='w', padx=20, pady=5)
        self.entry_pct_var_matriz_ingresos = ttk.Entry(frame_matriz, textvariable=self.PCT_VAR_MATRIZ_INGRESOS, width=12)
        self.entry_pct_var_matriz_ingresos.grid(row=3, column=1, sticky='w', padx=5, pady=5)
        self.entry_pct_var_matriz_ingresos.config(state='disabled')

        # Fila 4: Distribucion de factores
        ttk.Label(frame_matriz, text="Distribucion de factores:").grid(row=4, column=0, sticky='w', pady=5)
        ttk.Combobox(frame_matriz, textvariable=self.DISTRIBUCION_FACTORES,
                     values=["Log-normal", "Normal"], state='readonly', width=15
                     ).grid(row=4, column=1, sticky='w', padx=5, pady=5)

        # Fila 5: Correlacion matriz-satelites
        ttk.Label(frame_matriz, text="Correlacion matriz-satelites (rho):"
                  ).grid(row=5, column=0, sticky='w', pady=5)
        ttk.Entry(frame_matriz, textvariable=self.CORRELACION_MATRIZ_SAT, width=12
                  ).grid(row=5, column=1, sticky='w', padx=5, pady=5)

        # Callbacks para habilitar/deshabilitar Entries
        def _toggle_var_costos(*args):
            state = 'normal' if self.VAR_MATRIZ_COSTOS.get() else 'disabled'
            self.entry_pct_var_matriz_costos.config(state=state)
        self.VAR_MATRIZ_COSTOS.trace_add('write', _toggle_var_costos)

        def _toggle_var_ingresos(*args):
            state = 'normal' if self.VAR_MATRIZ_INGRESOS.get() else 'disabled'
            self.entry_pct_var_matriz_ingresos.config(state=state)
        self.VAR_MATRIZ_INGRESOS.trace_add('write', _toggle_var_ingresos)

        self.frame_sim_container = ttk.LabelFrame(parent, text="Resultados Estadisticos", padding=15)
        self.frame_sim_container.pack(fill='both', expand=True, padx=10, pady=10)
        self.canvas_simulacion = None
    
    def crear_tab_sensibilidad(self, parent):
        frame_config = ttk.LabelFrame(parent, text="Análisis de Sensibilidad", padding=15)
        frame_config.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_config, text="Variable:").grid(row=0, column=0, sticky='w', pady=5)
        self.combo_variable = ttk.Combobox(frame_config,
                                          values=["Costos Satélites", "Tasa Régimen Especial",
                                                  "Límite Utilidad", "Número de Satélites"],
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
        
