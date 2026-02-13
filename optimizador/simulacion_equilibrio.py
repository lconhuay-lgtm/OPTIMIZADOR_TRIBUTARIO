import numpy as np
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class SimEquilibrioMixin:
    def simular_equilibrio_financiero(self):
        """
        SIMULACIÓN MONTE CARLO – DISTRIBUCIÓN DEL SALDO DE REGULARIZACIÓN (PAGO MARZO).
        Compara el escenario ÓPTIMO (máximo ahorro) vs EQUILIBRIO (saldo ~ 0).
        Incluye:
          - Bootstrap percentil para IC y p‑valor.
          - Gráficos: densidad del grupo + violines por empresa.
          - Reporte de percentiles y reducción de riesgo.
        """
        try:
            if not hasattr(self, 'resultados'):
                messagebox.showwarning("Advertencia", "Primero debe calcular los márgenes óptimos")
                return
    
            # --------------------------------------------------------------
            # 1. CONFIGURACIÓN DE LA SIMULACIÓN
            # --------------------------------------------------------------
            n_sim = int(self.entry_n_sim_eq.get())
            variabilidad = float(self.entry_variabilidad.get()) / 100 if hasattr(self, 'entry_variabilidad') else 0.10
    
            # Parámetros de incertidumbre de la matriz y correlación
            var_matriz_costos = self.VAR_MATRIZ_COSTOS.get()
            pct_var_costos = self.PCT_VAR_MATRIZ_COSTOS.get() / 100.0
            var_matriz_ingresos = self.VAR_MATRIZ_INGRESOS.get()
            pct_var_ingresos = self.PCT_VAR_MATRIZ_INGRESOS.get() / 100.0
            distribucion = self.DISTRIBUCION_FACTORES.get()
            rho = self.CORRELACION_MATRIZ_SAT.get()
            if rho < -1.0 or rho > 1.0:
                raise ValueError("La correlación (rho) debe estar entre -1 y 1.")
    
            # Parámetros tributarios
            tasa_gral = self.resultados['parametros']['tasa_general'] / 100
            tasa_esp = self.resultados['parametros']['tasa_especial'] / 100
            tasa_pc = self.TASA_PAGO_CUENTA.get() / 100
            limite_util = self.resultados['parametros']['limite_utilidad']
            limite_ing = self.resultados['parametros']['limite_ingresos']
    
            # Datos base de la matriz
            ingresos_matriz_base = self.resultados['matriz']['ingresos']
            costos_matriz_base = self.resultados['matriz']['costos_externos']
            n_sats = len(self.resultados['satelites'])
    
            # --------------------------------------------------------------

            # --- Bucle MC (delegado) ---
            historial_saldos = self._simular_eq_mc_iterations(
                n_sim, variabilidad, ingresos_matriz_base, costos_matriz_base,
                tasa_gral, tasa_esp, tasa_pc, limite_util, limite_ing,
                var_matriz_costos, pct_var_costos, var_matriz_ingresos, pct_var_ingresos,
                distribucion, rho, n_sats
            )

            # --------------------------------------------------------------
            # 4. ESTADÍSTICAS AVANZADAS (BOOTSTRAP)
            # --------------------------------------------------------------
            saldos_g_opt = np.array(historial_saldos['Grupo_Total']['opt'])
            saldos_g_eq  = np.array(historial_saldos['Grupo_Total']['eq'])
    
            # Bootstrap percentil para la media del saldo en equilibrio
            conf_val = self.CONF_NIVEL.get() / 100
            n_boot = 5000
            rng = np.random.default_rng(42)
    
            boot_eq = np.array([
                np.mean(rng.choice(saldos_g_eq, size=len(saldos_g_eq), replace=True))
                for _ in range(n_boot)
            ])
            alpha = 1 - conf_val
            ic_eq_low = np.percentile(boot_eq, 100 * alpha / 2)
            ic_eq_high = np.percentile(boot_eq, 100 * (1 - alpha / 2))
    
            # Bootstrap para la diferencia de medias (óptimo - equilibrio)
            boot_diff = np.array([
                np.mean(rng.choice(saldos_g_opt, size=len(saldos_g_opt), replace=True)) -
                np.mean(rng.choice(saldos_g_eq,   size=len(saldos_g_eq),   replace=True))
                for _ in range(n_boot)
            ])
            p_valor_diff = np.mean(boot_diff <= 0)   # unilateral: H1: media_opt > media_eq
            p_valor_diff = max(p_valor_diff, 1 / n_boot)
    
            # Estadísticas descriptivas
            stats_eq = {
                'media': np.mean(saldos_g_eq),
                'mediana': np.median(saldos_g_eq),
                'std': np.std(saldos_g_eq),
                'p5': np.percentile(saldos_g_eq, 5),
                'p25': np.percentile(saldos_g_eq, 25),
                'p75': np.percentile(saldos_g_eq, 75),
                'p95': np.percentile(saldos_g_eq, 95),
                'ic_low': ic_eq_low,
                'ic_high': ic_eq_high
            }
            stats_opt = {
                'media': np.mean(saldos_g_opt),
                'mediana': np.median(saldos_g_opt),
                'p95': np.percentile(saldos_g_opt, 95)
            }
    
            # --------------------------------------------------------------

            # --- Visualizacion (delegada) ---
            empresas = ['Matriz'] + [s['nombre'] for s in self.resultados['satelites']]
            fig = self._visualizar_equilibrio_sim(
                stats_eq, stats_opt, saldos_g_opt, saldos_g_eq,
                p_valor_diff, n_sim, historial_saldos, empresas
            )

            self.fig_equilibrio_sim = fig
    
            # --------------------------------------------------------------
            # 6. GUARDAR RESULTADOS PARA EXPORTACIÓN
            # --------------------------------------------------------------
            self.datos_equilibrio_sim = {
                'n_sim': n_sim,
                'variabilidad': variabilidad * 100,
                'config': {
                    'var_matriz_costos': var_matriz_costos,
                    'pct_var_costos': pct_var_costos * 100,
                    'var_matriz_ingresos': var_matriz_ingresos,
                    'pct_var_ingresos': pct_var_ingresos * 100,
                    'distribucion': distribucion,
                    'rho': rho
                },
                'grupo_eq': stats_eq,
                'grupo_opt': stats_opt,
                'p_valor_diferencia': p_valor_diff,
                'historial_saldos': historial_saldos,   # guardado para posible exportación
                'empresas': empresas
            }
    
            self.canvas_equilibrio_sim = FigureCanvasTkAgg(fig, self.frame_eq_sim_graficos)
            self.canvas_equilibrio_sim.draw()
            self.canvas_equilibrio_sim.get_tk_widget().pack(fill='both', expand=True)
    
            # --------------------------------------------------------------
            # 7. MENSAJE DE FINALIZACIÓN
            # --------------------------------------------------------------
            messagebox.showinfo(
                "Simulación de Equilibrio Financiero",
                f"Análisis completado.\n\n"
                f"Saldo medio en equilibrio: {stats_eq['media']:,.0f}\n"
                f"IC 95%: [{stats_eq['ic_low']:,.0f} , {stats_eq['ic_high']:,.0f}]\n"
                f"Reducción del P95: {100*(1 - stats_eq['p95']/stats_opt['p95']):.1f}%\n"
                f"p-valor (óptimo > equilibrio): {p_valor_diff:.4f}"
            )
    
        except Exception as e:
            messagebox.showerror("Error", f"Error en simulación de equilibrio:\n{str(e)}")
            import traceback
            traceback.print_exc()   # para depuración en consola

