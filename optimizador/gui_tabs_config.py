import tkinter as tk
from tkinter import ttk, messagebox


class GuiConfigMixin:
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

        ttk.Label(frame_parametros, text="Tasa Pago a Cuenta IR (%):").grid(row=5, column=0, sticky='w', pady=8)
        entry_tasa_pc = ttk.Entry(frame_parametros, textvariable=self.TASA_PAGO_CUENTA, width=20)
        entry_tasa_pc.grid(row=5, column=1, pady=8, padx=5)

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
        
        tasa_pc = self.TASA_PAGO_CUENTA.get()

        info_text = f"""
PARAMETROS ACTUALES:

UIT:                                {uit:>15,.2f}
Tasa Regimen General:               {tasa_gral:>15.2f}%
Tasa Regimen Especial:              {tasa_esp:>15.2f}%
Diferencial:                        {diferencial:>15.2f} puntos
Tasa Pago a Cuenta IR:              {tasa_pc:>15.2f}%

Limite Utilidad:                    {lim_util_soles:>15,.2f}
Limite Ingresos:                    {lim_ing_soles:>15,.2f}

Ahorro maximo total:                {ahorro_max:>15,.2f}
Numero satelites:                   {self.num_satelites.get():>15}

FORMULA MARGEN OPTIMO (con gastos):
m* = (Limite_Utilidad + Gastos_Operativos) / Costos

FORMULA MARGEN EQUILIBRIO (IR = Pago a Cuenta):
m_eq = (t_pc * C + t_esp * G) / (C * (t_esp - t_pc))

Margen Equilibrio sin gastos:       {(tasa_pc / (tasa_esp - tasa_pc) * 100):>15.2f}%
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
        
  
