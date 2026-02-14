import tkinter as tk
from tkinter import ttk


class GuiDatosMixin:
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
    
