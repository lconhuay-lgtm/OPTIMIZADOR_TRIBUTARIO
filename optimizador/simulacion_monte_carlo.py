import numpy as np
from scipy import stats
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class SimMCMixin:
    def ejecutar_simulacion(self):
        try:
            if not hasattr(self, 'resultados'):
                messagebox.showwarning("Advertencia", "Primero debe calcular los margenes optimos")
                return

            n_sim = int(self.entry_num_sim.get())
            variabilidad = float(self.entry_variabilidad.get()) / 100

            # Leer parametros de variabilidad de matriz y correlacion
            var_matriz_costos = self.VAR_MATRIZ_COSTOS.get()
            pct_var_costos = self.PCT_VAR_MATRIZ_COSTOS.get() / 100.0
            var_matriz_ingresos = self.VAR_MATRIZ_INGRESOS.get()
            pct_var_ingresos = self.PCT_VAR_MATRIZ_INGRESOS.get() / 100.0
            distribucion = self.DISTRIBUCION_FACTORES.get()
            rho = self.CORRELACION_MATRIZ_SAT.get()

            if rho < -1.0 or rho > 1.0:
                raise ValueError("La correlacion (rho) debe estar entre -1 y 1.")
            if pct_var_costos < 0 or pct_var_ingresos < 0:
                raise ValueError("Los porcentajes de variabilidad deben ser >= 0.")

            ahorros = []
            tasa_gral = self.resultados['parametros']['tasa_general'] / 100
            tasa_esp = self.resultados['parametros']['tasa_especial'] / 100
            limite_util = self.resultados['parametros']['limite_utilidad']

            # Valores base de la matriz
            ingresos_base = self.resultados['matriz']['ingresos']
            costos_ext_base = self.resultados['matriz']['costos_externos']
            util_sin_est_base = self.resultados['matriz']['utilidad_sin_estructura']
            imp_sin_est_base = self.resultados['matriz']['impuesto_sin_estructura']
            n_sats = len(self.resultados['satelites'])

            for _ in range(n_sim):
                # Generar factores correlacionados para esta iteracion
                f_costos, f_ingresos, factores_sat = self._generar_factores_simulacion(
                    n_sats, variabilidad, var_matriz_costos, pct_var_costos,
                    var_matriz_ingresos, pct_var_ingresos, distribucion, rho
                )

                # Aplicar variabilidad a la matriz
                ingresos_sim = ingresos_base * f_ingresos
                costos_ext_sim = costos_ext_base * f_costos
                util_sin_estructura_sim = ingresos_sim - costos_ext_sim
                imp_sin_estructura_sim = max(0, util_sin_estructura_sim) * tasa_gral

                satelites_sim = []
                for idx, sat in enumerate(self.resultados['satelites']):
                    factor = factores_sat[idx]
                    costo_sim = sat['costo'] * factor
                    gastos_sim = sat['gastos_operativos'] * factor

                    if sat['regimen'].startswith('GENERAL'):
                        if gastos_sim > 0:
                            margen = gastos_sim / costo_sim
                            utilidad_neta = 0
                            impuesto = 0
                        else:
                            margen = 0
                            utilidad_neta = 0
                            impuesto = 0
                    else:
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

                total_compras = sum(s['precio_venta'] for s in satelites_sim)
                total_costos_sat = sum(s['costo'] for s in satelites_sim)

                nueva_utilidad_matriz = util_sin_estructura_sim - total_compras + total_costos_sat
                impuesto_matriz = max(0, nueva_utilidad_matriz) * tasa_gral

                impuesto_total = impuesto_matriz + sum(s['impuesto'] for s in satelites_sim)
                ahorro = imp_sin_estructura_sim - impuesto_total

                ahorros.append(ahorro)

            ahorros = np.array(ahorros)

            # --- CALCULOS ESTADISTICOS AVANZADOS ---
            media = np.mean(ahorros)
            mediana = np.median(ahorros)
            std = np.std(ahorros, ddof=1)
            cv = (std / media * 100) if media != 0 else 0
            conf_val = self.CONF_NIVEL.get() / 100
            alfa_val = self.ALFA_SIG.get()

            # Bootstrap percentil para IC y p-valor no parametrico
            n_boot = 5000
            boot_medias = np.array([
                np.mean(np.random.choice(ahorros, size=len(ahorros), replace=True))
                for _ in range(n_boot)
            ])
            alpha_boot = 1 - conf_val
            ic_boot_lower = np.percentile(boot_medias, 100 * alpha_boot / 2)
            ic_boot_upper = np.percentile(boot_medias, 100 * (1 - alpha_boot / 2))

            # P-valor no parametrico: proporcion de bootstrap medias <= 0
            p_valor_boot = np.mean(boot_medias <= 0)
            p_valor_boot = max(p_valor_boot, 1 / n_boot)  # floor

            # T-test clasico (referencia)
            t_stat, p_value_t = stats.ttest_1samp(ahorros, 0)

            decision = "RECHAZAR H0 (Ahorro Significativo)" if p_valor_boot < alfa_val else "NO RECHAZAR H0 (Incierto)"

            # Ajuste Log-Normal
            ahorros_positivos = ahorros[ahorros > 0]
            ln_ajustado = False
            ln_shape = ln_loc = ln_scale = 0
            if len(ahorros_positivos) > 50:
                try:
                    ln_shape, ln_loc, ln_scale = stats.lognorm.fit(ahorros_positivos, floc=0)
                    ln_ajustado = True
                except Exception:
                    pass

            # --- Visualizacion (delegada) ---
            fig = self._visualizar_simulacion(
                ahorros, media, mediana, std, cv, conf_val, alfa_val,
                ic_boot_lower, ic_boot_upper, p_valor_boot, p_value_t,
                t_stat, decision, n_sim, variabilidad,
                ln_ajustado, ln_shape, ln_loc, ln_scale, ahorros_positivos,
                var_matriz_costos, pct_var_costos,
                var_matriz_ingresos, pct_var_ingresos,
                distribucion, rho
            )

            self.fig_simulacion = fig
            self.datos_simulacion = {
                'n_sim': n_sim,
                'variabilidad': variabilidad * 100,
                'media': media,
                'mediana': mediana,
                'std': std,
                'cv': cv,
                'min': float(np.min(ahorros)),
                'max': float(np.max(ahorros)),
                'p5': float(np.percentile(ahorros, 5)),
                'p25': float(np.percentile(ahorros, 25)),
                'p75': float(np.percentile(ahorros, 75)),
                'p95': float(np.percentile(ahorros, 95)),
                'ic_lower': ic_boot_lower,
                'ic_upper': ic_boot_upper,
                'ic_metodo': 'Bootstrap Percentil',
                't_stat': t_stat,
                'p_value': p_valor_boot,
                'p_value_t': p_value_t,
                'decision': decision,
                'conf_nivel': conf_val * 100,
                'alfa': alfa_val,
                'n_bootstrap': n_boot,
                'ln_ajustado': ln_ajustado,
                'ln_shape': ln_shape if ln_ajustado else None,
                'ln_loc': ln_loc if ln_ajustado else None,
                'ln_scale': ln_scale if ln_ajustado else None,
                # Parametros de variabilidad matriz y correlacion
                'var_matriz_costos': var_matriz_costos,
                'pct_var_costos': pct_var_costos * 100,
                'var_matriz_ingresos': var_matriz_ingresos,
                'pct_var_ingresos': pct_var_ingresos * 100,
                'distribucion': distribucion,
                'rho': rho,
            }

            self.canvas_simulacion = FigureCanvasTkAgg(fig, self.frame_sim_container)
            self.canvas_simulacion.draw()
            self.canvas_simulacion.get_tk_widget().pack(fill='both', expand=True)

            messagebox.showinfo("Simulacion Completa",
                                f"Ahorro esperado: {media:,.0f}\n"
                                f"IC {conf_val*100:.0f}% Bootstrap: [{ic_boot_lower:,.0f}, {ic_boot_upper:,.0f}]\n"
                                f"P-valor (bootstrap): {p_valor_boot:.4e}\n"
                                f"Decision: {decision}")

        except Exception as e:
            messagebox.showerror("Error", f"Error en simulacion: {str(e)}")

