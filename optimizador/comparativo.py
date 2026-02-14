import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import messagebox


class ComparativoMixin:
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
                'tasa_efectiva': (self.resultados['grupo']['impuesto_total'] /
                               (self.resultados['matriz']['nueva_utilidad'] + self.resultados['grupo']['total_utilidad_satelites']) * 100)
                               if (self.resultados['matriz']['nueva_utilidad'] + self.resultados['grupo']['total_utilidad_satelites']) != 0 else 0,
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

            self.fig_comparativo = fig  # Guardar para Excel
            self.datos_comparativo = escenarios  # Guardar datos

            self.canvas_comparativo = FigureCanvasTkAgg(fig, self.frame_comp_container)
            self.canvas_comparativo.draw()
            self.canvas_comparativo.get_tk_widget().pack(fill='both', expand=True)

        except Exception as e:
            messagebox.showerror("Error", f"Error en comparativo: {str(e)}")

        
