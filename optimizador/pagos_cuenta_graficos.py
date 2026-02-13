import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class PagosCuentaGraficosMixin:
    def _generar_graficos_pagos_cuenta(self):
        """Genera graficos priorizando el punto de equilibrio financiero."""
        if self.canvas_pagos_cuenta:
            self.canvas_pagos_cuenta.get_tk_widget().destroy()

        d = self.datos_pagos_cuenta
        g = d['global']
        n_sats = len(d['satelites'])

        # Layout: primera fila = 2 paneles globales, luego satelites
        n_cols = 2
        n_rows_sats = (n_sats + 1) // 2
        n_rows = 1 + n_rows_sats

        fig, axes = plt.subplots(n_rows, n_cols, figsize=(14, 5 * n_rows), squeeze=False)
        fig.suptitle('PUNTO DE EQUILIBRIO FINANCIERO - IR = Pagos a Cuenta (Saldo = 0)',
                     fontsize=14, fontweight='bold', color='#1F4E79')

        # --- PANEL GLOBAL 1: Comparativo 3 escenarios ---
        ax_g1 = axes[0][0]
        categorias = ['IR\nTotal', 'Pagos a\nCuenta', 'Saldo\nRegulariz.']
        vals_sin = [g['impuesto_sin_estructura'], g['pc_matriz_sin'], g['saldo_matriz_sin']]
        vals_eq = [g['ir_grupo_eq'], g['pc_grupo_eq'], g['saldo_grupo_eq']]
        vals_opt = [g['ir_grupo_opt'], g['pc_grupo_opt'], g['saldo_grupo_opt']]
        x = np.arange(len(categorias))
        width = 0.25
        bars1 = ax_g1.bar(x - width, vals_sin, width, label='Sin Estructura', color='#E74C3C', alpha=0.85, edgecolor='white')
        bars2 = ax_g1.bar(x, vals_eq, width, label='Equilibrio (IR=PaC)', color='#3498DB', alpha=0.85, edgecolor='white')
        bars3 = ax_g1.bar(x + width, vals_opt, width, label='Margen Optimo (ref)', color='#95A5A6', alpha=0.7, edgecolor='white')
        ax_g1.set_title('Impuesto Global: Sin Estructura vs Equilibrio vs Optimo', fontsize=10, fontweight='bold')
        ax_g1.set_xticks(x)
        ax_g1.set_xticklabels(categorias, fontsize=8)
        ax_g1.legend(fontsize=7)
        ax_g1.grid(True, alpha=0.3, axis='y')
        ax_g1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:,.0f}K'))
        for bar in bars1:
            h = bar.get_height()
            ax_g1.text(bar.get_x() + bar.get_width()/2., h, f'{h/1000:,.0f}K',
                      ha='center', va='bottom', fontsize=6, fontweight='bold', color='#C0392B')
        for bar in bars2:
            h = bar.get_height()
            ax_g1.text(bar.get_x() + bar.get_width()/2., h, f'{h/1000:,.0f}K',
                      ha='center', va='bottom', fontsize=6, fontweight='bold', color='#2980B9')
        ax_g1.annotate(f'Ahorro Equilibrio:\nS/ {g["ahorro_eq_total"]:,.0f} ({g["ahorro_eq_pct"]:.1f}%)',
                       xy=(0.98, 0.95), xycoords='axes fraction', ha='right', va='top',
                       fontsize=8, fontweight='bold', color='#2980B9',
                       bbox=dict(boxstyle='round,pad=0.4', facecolor='#E8F0FE', edgecolor='#2980B9', alpha=0.9))

        # --- PANEL GLOBAL 2: Eficiencia (3 escenarios) ---
        ax_g2 = axes[0][1]
        metricas = ['Tasa\nEfectiva %', 'Eficiencia\nPaC %']
        vals_sin_pct = [g['tasa_efectiva_sin'], g['eficiencia_sin']]
        vals_eq_pct = [g['tasa_efectiva_eq'], g['eficiencia_eq']]
        vals_opt_pct = [g['tasa_efectiva_opt'], g['eficiencia_opt']]
        x2 = np.arange(len(metricas))
        bars4 = ax_g2.bar(x2 - width, vals_sin_pct, width, label='Sin Estructura', color='#E74C3C', alpha=0.85, edgecolor='white')
        bars5 = ax_g2.bar(x2, vals_eq_pct, width, label='Equilibrio', color='#3498DB', alpha=0.85, edgecolor='white')
        bars6 = ax_g2.bar(x2 + width, vals_opt_pct, width, label='Margen Optimo (ref)', color='#95A5A6', alpha=0.7, edgecolor='white')
        ax_g2.set_title('Eficiencia Financiera: 3 Escenarios', fontsize=10, fontweight='bold')
        ax_g2.set_xticks(x2)
        ax_g2.set_xticklabels(metricas, fontsize=9)
        ax_g2.set_ylabel('%', fontsize=10)
        ax_g2.legend(fontsize=7)
        ax_g2.grid(True, alpha=0.3, axis='y')
        ax_g2.axhline(y=100, color='#3498DB', linestyle='--', alpha=0.5, linewidth=1.5)
        ax_g2.text(1.35, 100, 'Equilibrio\nperfecto', fontsize=7, color='#3498DB', alpha=0.7, va='center')
        for bar in bars4:
            h = bar.get_height()
            ax_g2.text(bar.get_x() + bar.get_width()/2., h, f'{h:.1f}%',
                      ha='center', va='bottom', fontsize=7, fontweight='bold', color='#C0392B')
        for bar in bars5:
            h = bar.get_height()
            ax_g2.text(bar.get_x() + bar.get_width()/2., h, f'{h:.1f}%',
                      ha='center', va='bottom', fontsize=7, fontweight='bold', color='#2980B9')

        # --- PANELES POR SATELITE: IR vs PaC con equilibrio destacado ---
        for idx, sat in enumerate(d['satelites']):
            row = 1 + idx // n_cols
            col = idx % n_cols
            ax = axes[row][col]

            sens = sat['sensibilidad']
            margenes = [s['margen'] for s in sens]
            irs = [s['ir'] for s in sens]
            pacs = [s['pago_cuenta'] for s in sens]

            ax.plot(margenes, irs, 'b-', linewidth=2, label='IR Anual', marker='o', markersize=3)
            ax.plot(margenes, pacs, 'r--', linewidth=2, label=f'Pago a Cuenta ({d["tasa_pago_cuenta"]:.1f}%)', marker='s', markersize=3)
            ax.fill_between(margenes, irs, pacs, where=[i < p for i, p in zip(irs, pacs)],
                           alpha=0.15, color='red', label='Saldo a Favor')
            ax.fill_between(margenes, irs, pacs, where=[i >= p for i, p in zip(irs, pacs)],
                           alpha=0.15, color='green', label='Saldo por Pagar')

            # EQUILIBRIO: linea principal, gruesa, destacada
            if sat['margen_equilibrio'] is not None:
                ax.axvline(sat['margen_equilibrio'], color='#3498DB', linestyle='-', linewidth=2.5,
                          label=f'EQUILIBRIO ({sat["margen_equilibrio"]:.2f}%)', zorder=5)
                # Punto de cruce
                ax.plot(sat['margen_equilibrio'], sat['ir_equilibrio'], 'D', color='#3498DB',
                       markersize=10, zorder=6, markeredgecolor='white', markeredgewidth=1.5)

            # Margen optimo: linea secundaria, punteada, mas discreta
            ax.axvline(sat['margen_optimo_regimen'], color='gray', linestyle=':', linewidth=1,
                      label=f'M. Optimo ref ({sat["margen_optimo_regimen"]:.2f}%)', alpha=0.7)

            ax.set_xlabel('Margen (%)', fontsize=9)
            ax.set_ylabel('Monto (S/)', fontsize=9)
            ax.set_title(f'{sat["nombre"]} - Punto de Equilibrio', fontsize=11, fontweight='bold')
            ax.legend(fontsize=6, loc='upper left')
            ax.grid(True, alpha=0.3)
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))

        # Ocultar subplots vacios
        for idx in range(n_sats, n_rows_sats * n_cols):
            row = 1 + idx // n_cols
            col = idx % n_cols
            axes[row][col].set_visible(False)

        plt.tight_layout()

        self.fig_pagos_cuenta = fig

        self.canvas_pagos_cuenta = FigureCanvasTkAgg(fig, self.frame_pagos_graficos)
        self.canvas_pagos_cuenta.draw()
        self.canvas_pagos_cuenta.get_tk_widget().pack(fill='both', expand=True)

