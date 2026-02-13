import numpy as np
import matplotlib.pyplot as plt


class SensNsatsMixin:
    def _sensibilidad_num_satelites(self, ax1, ax2, sats_base, tasa_gral, tasa_esp,
                                     limite_util, limite_ing, util_sin_estructura,
                                     ahorro_base, rango):
        """Sensibilidad al numero de satelites."""
        # Satelite de referencia: promedio de los satelites NO generales, operando al limite
        sats_especial = [s for s in sats_base if not s['es_general']]
        if sats_especial:
            costo_ref = np.mean([s['costo'] for s in sats_especial])
            gastos_ref = np.mean([s['gastos'] for s in sats_especial])
        else:
            costo_ref = np.mean([s['costo'] for s in sats_base])
            gastos_ref = np.mean([s['gastos'] for s in sats_base])

        n_actual = len(sats_base)
        n_max = max(n_actual * 3, 15)
        ns = list(range(1, n_max + 1))
        ahorros_n = []
        tasa_efectiva_n = []
        ir_grupo_n = []

        for n in ns:
            sats_n = [{'costo': costo_ref, 'gastos': gastos_ref, 'es_general': False}] * n
            ahorro_n = self._calcular_ahorro_grupo(
                sats_n, tasa_gral, tasa_esp,
                limite_util, limite_ing, util_sin_estructura
            )
            ahorros_n.append(ahorro_n)

            # Calcular tasa efectiva para cada N
            total_imp_sat = 0
            total_compras = 0
            total_costos_s = 0
            for s in sats_n:
                margen_s = (limite_util + s['gastos']) / s['costo']
                precio_s = s['costo'] * (1 + margen_s)
                if precio_s > limite_ing:
                    margen_s = (limite_ing / s['costo']) - 1
                    precio_s = s['costo'] * (1 + margen_s)
                util_s = max(0, s['costo'] * margen_s - s['gastos'])
                if util_s <= limite_util:
                    imp_s = util_s * tasa_esp
                else:
                    imp_s = limite_util * tasa_esp + (util_s - limite_util) * tasa_gral
                total_imp_sat += imp_s
                total_compras += precio_s
                total_costos_s += s['costo']

            util_matriz = util_sin_estructura - total_compras + total_costos_s
            imp_matriz = max(0, util_matriz) * tasa_gral
            imp_total = imp_matriz + total_imp_sat
            util_total = max(0, util_matriz) + sum(
                max(0, s['costo'] * ((limite_util + s['gastos']) / s['costo']) - s['gastos'])
                for s in sats_n
            )
            tasa_ef = (imp_total / util_total * 100) if util_total > 0 else 0
            tasa_efectiva_n.append(tasa_ef)
            ir_grupo_n.append(imp_total)

        # Guardar datos para exportacion Excel
        self.datos_sensibilidad_nsats = {
            'ns': ns,
            'ahorros': ahorros_n,
            'tasa_efectiva': tasa_efectiva_n,
            'ir_grupo': ir_grupo_n,
            'n_actual': n_actual,
            'costo_ref': costo_ref,
            'gastos_ref': gastos_ref,
            'ahorro_base': ahorro_base,
        }

        # Grafico 1: Ahorro vs N satelites
        ax1.plot(ns, ahorros_n, color='#8E44AD', linewidth=2.5, marker='o', markersize=4)
        ax1.axvline(n_actual, color='red', linestyle='--', linewidth=2,
                   label=f'N Actual ({n_actual})')
        ax1.axhline(ahorro_base, color='green', linestyle='--',
                   alpha=0.5, label=f'Ahorro Base ({ahorro_base:,.0f})')
        ax1.fill_between(ns, ahorros_n, alpha=0.1, color='#8E44AD')
        ax1.set_xlabel('Numero de Satelites', fontsize=11)
        ax1.set_ylabel('Ahorro Tributario (S/)', fontsize=11)
        ax1.set_title('Ahorro Fiscal vs N Satelites', fontsize=13, fontweight='bold')
        ax1.grid(True, alpha=0.3)
        ax1.legend(fontsize=9)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(
            lambda x, p: f'{x/1e6:.2f}M' if abs(x) >= 1e6 else f'{x/1000:.0f}K'))

        # Grafico 2: Tasa efectiva vs N satelites
        ax2.plot(ns, tasa_efectiva_n, color='#E67E22', linewidth=2.5, marker='s', markersize=3)
        ax2.axvline(n_actual, color='red', linestyle='--', linewidth=2,
                   label=f'N Actual ({n_actual})')
        tasa_sin_est = tasa_gral * 100
        ax2.axhline(tasa_sin_est, color='#E74C3C', linestyle=':',
                   alpha=0.7, label=f'Sin Estructura ({tasa_sin_est:.1f}%)')
        ax2.set_xlabel('Numero de Satelites', fontsize=11)
        ax2.set_ylabel('Tasa Efectiva Grupo (%)', fontsize=11)
        ax2.set_title('Tasa Efectiva vs N Satelites', fontsize=13, fontweight='bold')
        ax2.grid(True, alpha=0.3)
        ax2.legend(fontsize=9)

