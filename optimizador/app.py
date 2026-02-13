import tkinter as tk
from tkinter import ttk

from optimizador.gui_tabs_config import GuiConfigMixin
from optimizador.gui_tabs_datos import GuiDatosMixin
from optimizador.gui_tabs_sim import GuiSimTabsMixin
from optimizador.gui_tabs_pagos import GuiPagosMixin
from optimizador.calculo_margenes import CalculoMargenesMixin
from optimizador.calculo_equilibrio import CalculoEquilibrioMixin
from optimizador.analisis_pagos_cuenta import PagosCuentaMixin
from optimizador.analisis_pagos_global import PagosCuentaGlobalMixin
from optimizador.pagos_cuenta_display import PagosCuentaDisplayMixin
from optimizador.pagos_cuenta_graficos import PagosCuentaGraficosMixin
from optimizador.simulacion_factores import SimFactoresMixin
from optimizador.simulacion_monte_carlo import SimMCMixin
from optimizador.simulacion_mc_viz import SimMCVizMixin
from optimizador.simulacion_equilibrio import SimEquilibrioMixin
from optimizador.simulacion_eq_loop import SimEqLoopMixin
from optimizador.simulacion_eq_viz import SimEqVizMixin
from optimizador.sensibilidad import SensibilidadMixin
from optimizador.sensibilidad_nsats import SensNsatsMixin
from optimizador.comparativo import ComparativoMixin
from optimizador.resultados_display import ResultadosDisplayMixin
from optimizador.resultados_graficos import ResultadosGraficosMixin
from optimizador.excel_estilos import ExcelEstilosMixin
from optimizador.excel_export import ExcelExportMixin
from optimizador.excel_ws1 import ExcelWs1Mixin
from optimizador.excel_ws2_ws3 import ExcelWs2Ws3Mixin
from optimizador.excel_ws4 import ExcelWs4Mixin
from optimizador.excel_ws5_ws6 import ExcelWs5Ws6Mixin
from optimizador.excel_ws7 import ExcelWs7Mixin
from optimizador.excel_ws7_detalle import ExcelWs7DetalleMixin
from optimizador.excel_ws8 import ExcelWs8Mixin
from optimizador.excel_ws9_ws10 import ExcelWs9Ws10Mixin
from optimizador.utilidades import UtilidadesMixin


class OptimizadorPro(
    GuiConfigMixin, GuiDatosMixin, GuiSimTabsMixin, GuiPagosMixin,
    CalculoMargenesMixin, CalculoEquilibrioMixin,
    PagosCuentaMixin, PagosCuentaGlobalMixin,
    PagosCuentaDisplayMixin, PagosCuentaGraficosMixin,
    SimFactoresMixin, SimMCMixin, SimMCVizMixin,
    SimEquilibrioMixin, SimEqLoopMixin, SimEqVizMixin,
    SensibilidadMixin, SensNsatsMixin,
    ComparativoMixin,
    ResultadosDisplayMixin, ResultadosGraficosMixin,
    ExcelEstilosMixin, ExcelExportMixin,
    ExcelWs1Mixin, ExcelWs2Ws3Mixin, ExcelWs4Mixin,
    ExcelWs5Ws6Mixin, ExcelWs7Mixin, ExcelWs7DetalleMixin,
    ExcelWs8Mixin, ExcelWs9Ws10Mixin,
    UtilidadesMixin,
):
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
        self.TASA_PAGO_CUENTA = tk.DoubleVar(value=1.5)  # Tasa pago a cuenta IR (%)

        # --- PARAMETROS DE SIMULACION PARA MATRIZ Y CORRELACION TERRITORIAL ---
        self.VAR_MATRIZ_COSTOS = tk.BooleanVar(value=False)
        self.PCT_VAR_MATRIZ_COSTOS = tk.DoubleVar(value=0.0)
        self.VAR_MATRIZ_INGRESOS = tk.BooleanVar(value=False)
        self.PCT_VAR_MATRIZ_INGRESOS = tk.DoubleVar(value=0.0)
        self.DISTRIBUCION_FACTORES = tk.StringVar(value="Log-normal")
        self.CORRELACION_MATRIZ_SAT = tk.DoubleVar(value=0.0)
        # Widgets Entry que necesitan habilitarse/deshabilitarse
        self.entry_pct_var_matriz_costos = None
        self.entry_pct_var_matriz_ingresos = None

        
        self.satelites_entries = []
        # Almacenar figuras para exportación Excel
        self.fig_resultados = None
        self.fig_simulacion = None
        self.fig_sensibilidad = None
        self.fig_comparativo = None
        self.datos_simulacion = None  # Estadísticas Monte Carlo
        self.fig_pagos_cuenta = None
        self.datos_pagos_cuenta = None
        self.fig_equilibrio_sim = None
        self.datos_equilibrio_sim = None
        self.datos_sensibilidad_nsats = None
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

        tab_pagos_cuenta = ttk.Frame(notebook)
        notebook.add(tab_pagos_cuenta, text='Pagos a Cuenta y Equilibrio')

        self.crear_tab_configuracion(tab_config)
        self.crear_tab_datos(tab_datos)
        self.crear_tab_resultados(tab_resultados)
        self.crear_tab_simulacion(tab_simulacion)
        self.crear_tab_sensibilidad(tab_sensibilidad)
        self.crear_tab_comparativo(tab_comparativo)
        self.crear_tab_pagos_cuenta(tab_pagos_cuenta)
        
