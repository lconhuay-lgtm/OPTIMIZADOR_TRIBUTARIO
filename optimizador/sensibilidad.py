import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import messagebox


class SensibilidadMixin:
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
            limite_ing = self.resultados['parametros']['limite_ingresos']
            util_sin_estructura = self.resultados['matriz']['utilidad_sin_estructura']

            # Datos base de satélites para el helper
            sats_base = [
                {
                    'costo': s['costo'],
                    'gastos': s['gastos_operativos'],
                    'es_general': s['regimen'].startswith('GENERAL')
                }
                for s in self.resultados['satelites']
            ]
            ahorro_base = self.resultados['grupo']['ahorro_tributario']

            if variable == "Costos Satélites":
                variaciones = np.linspace(-rango, rango, 50)
                ahorros_totales = []

                for var in variaciones:
                    sats_var = [
                        {
                            'costo': s['costo'] * (1 + var),
                            'gastos': s['gastos'] * (1 + var),
                            'es_general': s['es_general']
                        }
                        for s in sats_base
                    ]
                    ahorro = self._calcular_ahorro_grupo(
                        sats_var, tasa_gral, tasa_esp,
                        limite_util, limite_ing, util_sin_estructura
                    )
                    ahorros_totales.append(ahorro)

                ax1.plot(variaciones * 100, ahorros_totales, color='#e74c3c', linewidth=2.5, marker='o', markersize=3)
                ax1.axvline(0, color='gray', linestyle='--', alpha=0.5)
                ax1.axhline(ahorro_base, color='green', linestyle='--',
                           alpha=0.5, label=f'Ahorro Base ({ahorro_base:,.0f})')
                ax1.set_xlabel('Variación en Costos (%)', fontsize=11)
                ax1.set_ylabel('Ahorro Tributario', fontsize=11)
                ax1.set_title('Impacto en Ahorro', fontsize=13, fontweight='bold')
                ax1.grid(True, alpha=0.3)
                ax1.legend()
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.2f}M' if x >= 1e6 else f'{x/1000:.0f}K'))

                # Elasticidad (variación del ahorro por punto porcentual de cambio en costos)
                ahorros_arr = np.array(ahorros_totales)
                elasticidad = np.gradient(ahorros_arr, variaciones * 100)
                ax2.plot(variaciones * 100, elasticidad, color='#3498db', linewidth=2.5)
                ax2.axvline(0, color='gray', linestyle='--', alpha=0.5)
                ax2.axhline(0, color='black', linestyle='-', alpha=0.3)
                ax2.set_xlabel('Variación en Costos (%)', fontsize=11)
                ax2.set_ylabel('dAhorro/dCosto (S/ por pp)', fontsize=11)
                ax2.set_title('Sensibilidad del Ahorro', fontsize=13, fontweight='bold')
                ax2.grid(True, alpha=0.3)
                ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.1f}K'))

            elif variable == "Tasa Régimen Especial":
                tasa_base = self.resultados['parametros']['tasa_especial']
                rango_abs = tasa_base * rango
                tasas_especial = np.linspace(max(1, tasa_base - rango_abs), tasa_base + rango_abs, 50)
                ahorros_totales = []

                for tasa in tasas_especial:
                    ahorro = self._calcular_ahorro_grupo(
                        sats_base, tasa_gral, tasa / 100,
                        limite_util, limite_ing, util_sin_estructura
                    )
                    ahorros_totales.append(ahorro)

                ax1.plot(tasas_especial, ahorros_totales, color='#9b59b6', linewidth=2.5, marker='o', markersize=3)
                ax1.axvline(tasa_base, color='red', linestyle='--', alpha=0.5, label=f'Tasa Actual ({tasa_base:.1f}%)')
                ax1.axhline(ahorro_base, color='green', linestyle='--',
                           alpha=0.5, label=f'Ahorro Base ({ahorro_base:,.0f})')
                ax1.set_xlabel('Tasa Régimen Especial (%)', fontsize=11)
                ax1.set_ylabel('Ahorro Tributario', fontsize=11)
                ax1.set_title('Sensibilidad a Tasa Especial', fontsize=13, fontweight='bold')
                ax1.grid(True, alpha=0.3)
                ax1.legend()
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.2f}M' if x >= 1e6 else f'{x/1000:.0f}K'))

                # Elasticidad: dAhorro/dTasa
                ahorros_arr = np.array(ahorros_totales)
                elasticidad = np.gradient(ahorros_arr, tasas_especial)
                ax2.plot(tasas_especial, elasticidad, color='#e67e22', linewidth=2.5)
                ax2.axvline(tasa_base, color='red', linestyle='--', alpha=0.5)
                ax2.axhline(0, color='black', linestyle='-', alpha=0.3)
                ax2.set_xlabel('Tasa Régimen Especial (%)', fontsize=11)
                ax2.set_ylabel('dAhorro/dTasa (S/ por pp)', fontsize=11)
                ax2.set_title('Elasticidad a Tasa Especial', fontsize=13, fontweight='bold')
                ax2.grid(True, alpha=0.3)
                ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))

            elif variable == "Límite Utilidad":
                limite_base = self.resultados['parametros']['limite_utilidad'] / uit
                rango_abs = limite_base * rango
                limites_uit = np.linspace(max(1, limite_base - rango_abs), limite_base + rango_abs, 50)
                ahorros_totales = []

                for limite_uit_val in limites_uit:
                    ahorro = self._calcular_ahorro_grupo(
                        sats_base, tasa_gral, tasa_esp,
                        limite_uit_val * uit, limite_ing, util_sin_estructura
                    )
                    ahorros_totales.append(ahorro)

                ax1.plot(limites_uit, ahorros_totales, color='#1abc9c', linewidth=2.5, marker='o', markersize=3)
                ax1.axvline(limite_base, color='red', linestyle='--', alpha=0.5, label=f'Limite Actual ({limite_base:.1f} UIT)')
                ax1.axhline(ahorro_base, color='green', linestyle='--',
                           alpha=0.5, label=f'Ahorro Base ({ahorro_base:,.0f})')
                ax1.set_xlabel('Limite Utilidad (UIT)', fontsize=11)
                ax1.set_ylabel('Ahorro Tributario', fontsize=11)
                ax1.set_title('Sensibilidad a Limite', fontsize=13, fontweight='bold')
                ax1.grid(True, alpha=0.3)
                ax1.legend()
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.2f}M' if x >= 1e6 else f'{x/1000:.0f}K'))

                # Ahorro marginal por UIT adicional
                ahorros_arr = np.array(ahorros_totales)
                ahorro_marginal = np.gradient(ahorros_arr, limites_uit)
                ax2.plot(limites_uit, ahorro_marginal, color='#f39c12', linewidth=2.5)
                ax2.axvline(limite_base, color='red', linestyle='--', alpha=0.5)
                ax2.axhline(0, color='black', linestyle='-', alpha=0.3)
                ax2.set_xlabel('Limite Utilidad (UIT)', fontsize=11)
                ax2.set_ylabel('Ahorro Marginal por UIT', fontsize=11)
                ax2.set_title('Beneficio Marginal', fontsize=13, fontweight='bold')
                ax2.grid(True, alpha=0.3)
                ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))

            else:  # Numero de Satelites
                self._sensibilidad_num_satelites(
                    ax1, ax2, sats_base, tasa_gral, tasa_esp,
                    limite_util, limite_ing, util_sin_estructura,
                    ahorro_base, rango)

            plt.tight_layout()

            self.fig_sensibilidad = fig  # Guardar para Excel

            self.canvas_sensibilidad = FigureCanvasTkAgg(fig, self.frame_sens_container)
            self.canvas_sensibilidad.draw()
            self.canvas_sensibilidad.get_tk_widget().pack(fill='both', expand=True)

        except Exception as e:
            messagebox.showerror("Error", f"Error en sensibilidad: {str(e)}")

