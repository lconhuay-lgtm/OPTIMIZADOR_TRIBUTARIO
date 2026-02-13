import tkinter as tk
from tkinter import ttk


class GuiPagosMixin:
    def crear_tab_pagos_cuenta(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        frame_config_pc = ttk.LabelFrame(scrollable_frame,
                                         text="Analisis de Pagos a Cuenta y Equilibrio Fiscal (Teoria de Juegos)",
                                         padding=15)
        frame_config_pc.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        ttk.Label(frame_config_pc, text="Tasa Pago a Cuenta (%):").grid(row=0, column=0, sticky='w', pady=5)
        ttk.Entry(frame_config_pc, textvariable=self.TASA_PAGO_CUENTA, width=12).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_config_pc, text="Pasos Sensibilidad:").grid(row=0, column=2, sticky='w', padx=10, pady=5)
        self.entry_pasos_sens_pc = ttk.Entry(frame_config_pc, width=12)
        self.entry_pasos_sens_pc.insert(0, "15")
        self.entry_pasos_sens_pc.grid(row=0, column=3, padx=5, pady=5)

        ttk.Button(frame_config_pc, text="Ejecutar Analisis Pagos a Cuenta y Sensibilidad",
                   command=self.analizar_pagos_cuenta).grid(row=1, column=0, columnspan=2, pady=10)

        ttk.Label(frame_config_pc, text="Sim. Equilibrio (N):").grid(row=1, column=2, sticky='w', padx=10)
        self.entry_n_sim_eq = ttk.Entry(frame_config_pc, width=12)
        self.entry_n_sim_eq.insert(0, "5000")
        self.entry_n_sim_eq.grid(row=1, column=3, padx=5, pady=5)

        ttk.Button(frame_config_pc, text="Simular Equilibrio (Monte Carlo)",
                   command=self.simular_equilibrio_financiero).grid(row=2, column=0, columnspan=4, pady=10)

        # Texto de resultados
        frame_texto_pc = ttk.LabelFrame(scrollable_frame, text="Analisis Detallado por Satelite", padding=10)
        frame_texto_pc.grid(row=1, column=0, padx=10, pady=5, sticky='ew')
        self.text_pagos_cuenta = tk.Text(frame_texto_pc, height=18, width=140, wrap='none', font=('Courier', 8))
        scroll_h = ttk.Scrollbar(frame_texto_pc, orient='horizontal', command=self.text_pagos_cuenta.xview)
        scroll_v = ttk.Scrollbar(frame_texto_pc, orient='vertical', command=self.text_pagos_cuenta.yview)
        self.text_pagos_cuenta.configure(xscrollcommand=scroll_h.set, yscrollcommand=scroll_v.set)
        self.text_pagos_cuenta.grid(row=0, column=0, sticky='nsew')
        scroll_v.grid(row=0, column=1, sticky='ns')
        scroll_h.grid(row=1, column=0, sticky='ew')
        frame_texto_pc.columnconfigure(0, weight=1)
        frame_texto_pc.rowconfigure(0, weight=1)

        # Graficos
        self.frame_pagos_graficos = ttk.LabelFrame(scrollable_frame, text="Visualizacion Pagos a Cuenta vs IR", padding=10)
        self.frame_pagos_graficos.grid(row=2, column=0, padx=10, pady=5, sticky='ew')
        self.canvas_pagos_cuenta = None

        # Graficos simulacion equilibrio
        self.frame_eq_sim_graficos = ttk.LabelFrame(scrollable_frame,
                                                     text="Simulacion Monte Carlo - Equilibrio Financiero", padding=10)
        self.frame_eq_sim_graficos.grid(row=3, column=0, padx=10, pady=5, sticky='ew')
        self.canvas_equilibrio_sim = None

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

