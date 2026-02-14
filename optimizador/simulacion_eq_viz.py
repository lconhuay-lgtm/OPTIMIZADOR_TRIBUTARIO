import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Patch


class SimEqVizMixin:
    def _visualizar_equilibrio_sim(self, stats_eq, stats_opt,
                                    saldos_g_opt, saldos_g_eq,
                                    p_valor_diff, n_sim,
                                    historial_saldos, empresas):
        """Graficos para simulacion de equilibrio financiero."""
        # 5. VISUALIZACIÓN AVANZADA (DENSIDAD + VIOLINES)
        # --------------------------------------------------------------
        if self.canvas_equilibrio_sim:
            self.canvas_equilibrio_sim.get_tk_widget().destroy()

        fig = plt.figure(figsize=(16, 14))
        gs = fig.add_gridspec(2, 2, height_ratios=[1, 1.5], hspace=0.30, wspace=0.30)

        # ----- PANEL SUPERIOR IZQUIERDO: DENSIDAD DEL GRUPO CONSOLIDADO -----
        ax1 = fig.add_subplot(gs[0, 0])

        # Estimacion de densidad kernel (KDE) para suavizar
        from scipy.stats import gaussian_kde
        kde_opt = gaussian_kde(saldos_g_opt)
        kde_eq  = gaussian_kde(saldos_g_eq)

        x_min = min(saldos_g_opt.min(), saldos_g_eq.min())
        x_max = max(saldos_g_opt.max(), saldos_g_eq.max())
        x = np.linspace(x_min, x_max, 300)

        ax1.fill_between(x, kde_opt(x), alpha=0.35, color='#E74C3C',
                         edgecolor='#C0392B', linewidth=1.5)
        ax1.fill_between(x, kde_eq(x),  alpha=0.55, color='#27AE60',
                         edgecolor='#1E8449', linewidth=1.5)

        # Lineas de referencia
        ax1.axvline(stats_eq['media'], color='#1E8449', linestyle='--', linewidth=2)
        ax1.axvline(stats_opt['media'], color='#943126', linestyle='--', linewidth=2)

        ax1.set_title(f'SALDO REGULARIZACION GRUPO ({n_sim:,} sim.)',
                      fontsize=10, fontweight='bold')
        ax1.set_xlabel('Saldo a pagar en Marzo (S/)', fontsize=8)
        ax1.yaxis.set_visible(False)
        ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}k'))
        ax1.tick_params(labelsize=8)

        # Leyenda compacta
        from matplotlib.patches import Patch
        leg_density = [
            Patch(facecolor='#E74C3C', alpha=0.5, edgecolor='#C0392B', label='Optimo'),
            Patch(facecolor='#27AE60', alpha=0.6, edgecolor='#1E8449', label='Equilibrio')
        ]
        ax1.legend(handles=leg_density, fontsize=7, loc='upper left')

        # ----- PANEL SUPERIOR DERECHO: ESTADISTICAS -----
        ax_stats = fig.add_subplot(gs[0, 1])
        ax_stats.axis('off')

        # Recuadro con estadisticas clave
        p95_ratio = stats_eq['p95'] / stats_opt['p95'] if stats_opt['p95'] != 0 else 1
        textstr = (
            f"{'EQUILIBRIO vs OPTIMO':^36}\n"
            f"{'='*36}\n\n"
            f"{'SALDO GRUPO (Equilibrio)':}\n"
            f"  Media:     {stats_eq['media']:>12,.0f}\n"
            f"  Mediana:   {stats_eq['mediana']:>12,.0f}\n"
            f"  P5:        {stats_eq['p5']:>12,.0f}\n"
            f"  P95:       {stats_eq['p95']:>12,.0f}\n"
            f"  IC 95%:    [{stats_eq['ic_low']:>10,.0f} ,\n"
            f"              {stats_eq['ic_high']:>10,.0f}]\n\n"
            f"{'SALDO GRUPO (Optimo)':}\n"
            f"  Media:     {stats_opt['media']:>12,.0f}\n"
            f"  P95:       {stats_opt['p95']:>12,.0f}\n\n"
            f"Reduccion P95: {100*(1 - p95_ratio):.1f}%\n"
            f"p-valor (Opt>Eq): {p_valor_diff:.4f}"
        )

        ax_stats.text(0.05, 0.95, textstr, transform=ax_stats.transAxes, fontsize=8,
                      verticalalignment='top', fontfamily='monospace',
                      bbox=dict(boxstyle='round', facecolor='lightyellow', alpha=0.9,
                                edgecolor='#1F4E79', linewidth=1.5))

        # ----- PANEL INFERIOR: BARRAS MEDIANAS POR EMPRESA (Opt vs Eq) -----
        ax2 = fig.add_subplot(gs[1, :])

        empresas = ['Matriz'] + [s['nombre'] for s in self.resultados['satelites']]
        mediana_opt = [np.median(historial_saldos[emp]['opt']) for emp in empresas]
        mediana_eq  = [np.median(historial_saldos[emp]['eq'])  for emp in empresas]
        p95_opt = [np.percentile(historial_saldos[emp]['opt'], 95) for emp in empresas]
        p5_opt  = [np.percentile(historial_saldos[emp]['opt'], 5)  for emp in empresas]
        p95_eq  = [np.percentile(historial_saldos[emp]['eq'], 95) for emp in empresas]
        p5_eq   = [np.percentile(historial_saldos[emp]['eq'], 5)  for emp in empresas]

        x_pos = np.arange(len(empresas))
        width = 0.35

        # Barras de mediana con barras de error P5-P95
        err_opt = [np.array(mediana_opt) - np.array(p5_opt),
                   np.array(p95_opt) - np.array(mediana_opt)]
        err_eq  = [np.array(mediana_eq) - np.array(p5_eq),
                   np.array(p95_eq) - np.array(mediana_eq)]

        bars_opt = ax2.bar(x_pos - width/2, mediana_opt, width, label='Optimo (max ahorro)',
                           color='#E74C3C', alpha=0.8, edgecolor='white',
                           yerr=err_opt, capsize=3, error_kw={'linewidth': 0.8, 'alpha': 0.5})
        bars_eq = ax2.bar(x_pos + width/2, mediana_eq, width, label='Equilibrio (saldo~0)',
                          color='#27AE60', alpha=0.8, edgecolor='white',
                          yerr=err_eq, capsize=3, error_kw={'linewidth': 0.8, 'alpha': 0.5})

        # Linea horizontal en cero
        ax2.axhline(0, color='black', linewidth=1, linestyle='-', alpha=0.5)

        # Etiquetas de valor sobre cada barra
        for i, (mo, me) in enumerate(zip(mediana_opt, mediana_eq)):
            if abs(mo) > 100:
                ax2.text(i - width/2, mo + (abs(mo) * 0.05 + 50) * (1 if mo >= 0 else -1),
                         f'{mo/1000:.1f}k', ha='center', fontsize=6.5, fontweight='bold', color='#943126')
            if abs(me) > 100:
                ax2.text(i + width/2, me + (abs(me) * 0.05 + 50) * (1 if me >= 0 else -1),
                         f'{me/1000:.1f}k', ha='center', fontsize=6.5, fontweight='bold', color='#1E8449')

        ax2.set_title('SALDO DE REGULARIZACION POR EMPRESA - Mediana con rango P5-P95',
                      fontsize=11, fontweight='bold')
        ax2.set_xticks(x_pos)
        ax2.set_xticklabels(empresas, fontsize=8, fontweight='bold', rotation=15 if len(empresas) > 4 else 0)
        ax2.set_ylabel('Saldo a Regularizar en Marzo (S/)', fontsize=9)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}k'))
        ax2.tick_params(labelsize=8)
        ax2.grid(axis='y', alpha=0.2, linestyle='--')
        ax2.legend(fontsize=8, loc='upper right')

        fig.suptitle(f'SIMULACION EQUILIBRIO FINANCIERO - {n_sim:,} escenarios Monte Carlo',
                     fontsize=13, fontweight='bold', y=0.99)
        self.fig_equilibrio_sim = fig

        return fig
