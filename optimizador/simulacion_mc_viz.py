import numpy as np
import matplotlib.pyplot as plt
from scipy import stats


class SimMCVizMixin:
    def _visualizar_simulacion(self, ahorros, media, mediana, std, cv, conf_val, alfa_val,
                                ic_boot_lower, ic_boot_upper, p_valor_boot, p_value_t,
                                t_stat, decision, n_sim, variabilidad,
                                ln_ajustado, ln_shape, ln_loc, ln_scale, ahorros_positivos,
                                var_matriz_costos, pct_var_costos,
                                var_matriz_ingresos, pct_var_ingresos,
                                distribucion, rho):
        """Graficos 2x2 para simulacion Monte Carlo."""
        n_boot = 5000  # para texto informativo
        # --- VISUALIZACION MEJORADA: 2x2 ---
        if self.canvas_simulacion:
            self.canvas_simulacion.get_tk_widget().destroy()

        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
        fig.suptitle(f'Simulacion Monte Carlo ({n_sim:,} iter.) - Analisis de Riesgo',
                     fontsize=13, fontweight='bold', y=0.98)

        # [1] Histograma + curva Log-Normal ajustada
        ax1.hist(ahorros, bins=60, color='#3498db', alpha=0.7, edgecolor='white',
                 density=True, label='Datos simulados')
        if ln_ajustado:
            x_ln = np.linspace(max(0.01, ahorros.min()), ahorros.max(), 300)
            pdf_ln = stats.lognorm.pdf(x_ln, ln_shape, ln_loc, ln_scale)
            ax1.plot(x_ln, pdf_ln, 'r-', linewidth=2,
                     label=f'Log-Normal (s={ln_shape:.3f})')
        ax1.axvline(media, color='orange', linestyle='--', linewidth=2,
                    label=f'Media: {media:,.0f}')
        ax1.axvline(mediana, color='green', linestyle=':', linewidth=2,
                    label=f'Mediana: {mediana:,.0f}')
        ax1.fill_betweenx([0, ax1.get_ylim()[1] if ax1.get_ylim()[1] > 0 else 0.001],
                          ic_boot_lower, ic_boot_upper, alpha=0.15, color='red',
                          label=f'IC {conf_val*100:.0f}% Boot')
        ax1.set_xlabel('Ahorro Tributario (S/)', fontsize=9)
        ax1.set_ylabel('Densidad', fontsize=9)
        ax1.set_title('Histograma + Ajuste Log-Normal', fontsize=10, fontweight='bold')
        ax1.legend(fontsize=6.5, loc='upper right')
        ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
        ax1.tick_params(labelsize=8)
        ax1.grid(alpha=0.3)

        # [2] Q-Q Plot (Log-Normal)
        if ln_ajustado and len(ahorros_positivos) > 10:
            ahorros_sorted = np.sort(ahorros_positivos)
            n_qq = len(ahorros_sorted)
            teoricos = stats.lognorm.ppf(
                (np.arange(1, n_qq + 1) - 0.5) / n_qq,
                ln_shape, ln_loc, ln_scale
            )
            ax2.scatter(teoricos, ahorros_sorted, s=4, alpha=0.4, color='#2E75B6')
            lim_min = min(teoricos.min(), ahorros_sorted.min())
            lim_max = max(teoricos.max(), ahorros_sorted.max())
            ax2.plot([lim_min, lim_max], [lim_min, lim_max], 'r--', linewidth=1.5,
                     label='Linea 45')
            ax2.set_xlabel('Cuantiles Teoricos (Log-Normal)', fontsize=9)
            ax2.set_ylabel('Cuantiles Observados', fontsize=9)
            ax2.set_title('Q-Q Plot Log-Normal', fontsize=10, fontweight='bold')
            ax2.legend(fontsize=7)
        else:
            sorted_ahorros = np.sort(ahorros)
            cumulative = np.arange(1, len(sorted_ahorros) + 1) / len(sorted_ahorros)
            ax2.plot(sorted_ahorros, cumulative * 100, color='#9b59b6', linewidth=2)
            ax2.axhline(50, color='red', linestyle='--', alpha=0.5, label='Mediana')
            ax2.axhline(95, color='green', linestyle='--', alpha=0.5, label='P95')
            ax2.set_xlabel('Ahorro Tributario', fontsize=9)
            ax2.set_ylabel('Prob. Acumulada (%)', fontsize=9)
            ax2.set_title('CDF (sin ajuste Log-Normal)', fontsize=10, fontweight='bold')
            ax2.legend(fontsize=7)
        ax2.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
        ax2.tick_params(labelsize=8)
        ax2.grid(alpha=0.3)

        # [3] Violin + Boxplot
        parts = ax3.violinplot(ahorros, positions=[1], showmeans=True,
                               showmedians=True, showextrema=False)
        for pc in parts['bodies']:
            pc.set_facecolor('#3498db')
            pc.set_alpha(0.4)
        parts['cmeans'].set_color('red')
        parts['cmedians'].set_color('green')
        bp = ax3.boxplot(ahorros, positions=[1], widths=0.15, patch_artist=True,
                         showfliers=False)
        bp['boxes'][0].set_facecolor('#2ecc71')
        bp['boxes'][0].set_alpha(0.6)
        ax3.set_ylabel('Ahorro Tributario (S/)', fontsize=9)
        ax3.set_title('Violin + Box Plot', fontsize=10, fontweight='bold')
        ax3.set_xticks([1])
        ax3.set_xticklabels(['Ahorro'])
        ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1000:.0f}K'))
        ax3.tick_params(labelsize=8)
        ax3.grid(axis='y', alpha=0.3)

        # [4] Panel de estadisticos
        mat_info = ""
        if var_matriz_costos:
            mat_info += f"\nVar.Costos Mat: +/-{pct_var_costos*100:.1f}%"
        if var_matriz_ingresos:
            mat_info += f"\nVar.Ing. Mat:   +/-{pct_var_ingresos*100:.1f}%"
        if var_matriz_costos and rho != 0:
            mat_info += f"\nCorrelacion:    rho={rho:.2f}"
        if mat_info:
            mat_info = f"\nMATRIZ:{mat_info}\nDistribucion:   {distribucion}"

        stats_text = f"""ESTADISTICAS SIMULACION
{'='*34}
Iteraciones:  {n_sim:,}
Variabilidad: +/-{variabilidad*100:.1f}%{mat_info}

AHORRO TRIBUTARIO:
  Media:      {media:>12,.0f}
  Mediana:    {mediana:>12,.0f}
  Desv.Est.:  {std:>12,.0f}
  Coef.Var.:  {cv:>12,.2f}%

PERCENTILES:
  P5:         {np.percentile(ahorros, 5):>12,.0f}
  P25:        {np.percentile(ahorros, 25):>12,.0f}
  P75:        {np.percentile(ahorros, 75):>12,.0f}
  P95:        {np.percentile(ahorros, 95):>12,.0f}

BOOTSTRAP IC {conf_val*100:.0f}%:
  [{ic_boot_lower:>12,.0f},
   {ic_boot_upper:>12,.0f}]

PRUEBA HIPOTESIS (H0: ahorro=0):
  P-valor Boot:  {p_valor_boot:.4e}
  t-stat ref:    {t_stat:.4f}
  P-valor t:     {p_value_t:.4e}
  Decision:      {decision}"""
        if ln_ajustado:
            stats_text += f"""

LOG-NORMAL FIT:
  Shape(s):  {ln_shape:.4f}
  Scale:     {ln_scale:>12,.0f}"""

        ax4.axis('off')
        ax4.text(0.05, 0.95, stats_text, transform=ax4.transAxes,
                 fontsize=7, verticalalignment='top', fontfamily='monospace',
                 bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

        plt.subplots_adjust(left=0.08, right=0.95, top=0.93, bottom=0.08,
                            hspace=0.35, wspace=0.30)

        return fig
