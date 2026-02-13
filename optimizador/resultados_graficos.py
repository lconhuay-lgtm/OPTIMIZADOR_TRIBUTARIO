import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np


class ResultadosGraficosMixin:
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

        self.fig_resultados = fig  # Guardar para Excel

        self.canvas_graficos = FigureCanvasTkAgg(fig, self.frame_graficos_container)
        self.canvas_graficos.draw()
        self.canvas_graficos.get_tk_widget().pack(fill='both', expand=True)

