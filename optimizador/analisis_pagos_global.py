import numpy as np
from tkinter import messagebox


class PagosCuentaGlobalMixin:
    def _analizar_global_pagos_cuenta(self, resultados_pc, tasa_pc, tasa_general):
        """Analisis global: sin estructura vs con estructura en punto de equilibrio."""
        # --- ANALISIS GLOBAL: SIN ESTRUCTURA vs CON ESTRUCTURA (enfoque equilibrio financiero) ---
        r = self.resultados
        impuesto_sin_estructura = r['matriz']['impuesto_sin_estructura']
        utilidad_sin_estructura = r['matriz']['utilidad_sin_estructura']

        # Datos base
        ingresos_matriz = r['matriz']['ingresos']
        total_costos_satelites = sum(rpc['costo'] for rpc in resultados_pc)

        # Pagos a cuenta SIN estructura
        pc_matriz_sin = tasa_pc * ingresos_matriz
        saldo_matriz_sin = impuesto_sin_estructura - pc_matriz_sin

        # === CONSOLIDADO EN PUNTO DE EQUILIBRIO (CORRECCION: recalcular impuesto matriz) ===
        total_compras_eq = sum(rpc['precio_equilibrio'] for rpc in resultados_pc)
        total_ir_eq_sats = sum(rpc['ir_equilibrio'] for rpc in resultados_pc)
        total_pc_eq_sats = sum(rpc['pago_cuenta_eq'] for rpc in resultados_pc)
        total_saldo_eq_sats = sum(rpc['saldo_equilibrio'] for rpc in resultados_pc)

        # Recalcular utilidad e impuesto de la matriz EN EQUILIBRIO
        # La matriz compra a los satelites a precio_equilibrio en lugar de a costo
        utilidad_matriz_eq = utilidad_sin_estructura - total_compras_eq + total_costos_satelites
        utilidad_matriz_eq = max(0, utilidad_matriz_eq)
        impuesto_matriz_eq = utilidad_matriz_eq * tasa_general
        pc_matriz_eq = tasa_pc * ingresos_matriz
        saldo_matriz_eq = impuesto_matriz_eq - pc_matriz_eq

        # Grupo en equilibrio
        ir_grupo_eq = impuesto_matriz_eq + total_ir_eq_sats
        pc_grupo_eq = pc_matriz_eq + total_pc_eq_sats
        saldo_grupo_eq = ir_grupo_eq - pc_grupo_eq
        eficiencia_eq = (ir_grupo_eq / pc_grupo_eq * 100) if pc_grupo_eq > 0 else 0
        utilidad_eq_sats = sum(rpc['utilidad_neta_eq'] for rpc in resultados_pc)
        utilidad_total_eq = utilidad_matriz_eq + utilidad_eq_sats
        tasa_efectiva_eq = (ir_grupo_eq / utilidad_total_eq * 100) if utilidad_total_eq > 0 else 0
        ahorro_eq_total = impuesto_sin_estructura - ir_grupo_eq
        ahorro_eq_pct = (ahorro_eq_total / impuesto_sin_estructura * 100) if impuesto_sin_estructura > 0 else 0

        # === CONSOLIDADO EN MARGEN OPTIMO (REFERENCIA - misma correccion) ===
        total_compras_opt = sum(rpc['precio_optimo'] for rpc in resultados_pc)
        total_ir_opt_sats = sum(rpc['ir_optimo'] for rpc in resultados_pc)
        total_pc_opt_sats = sum(rpc['pago_cuenta_optimo'] for rpc in resultados_pc)
        total_saldo_opt_sats = sum(rpc['saldo_optimo'] for rpc in resultados_pc)

        utilidad_matriz_opt = utilidad_sin_estructura - total_compras_opt + total_costos_satelites
        utilidad_matriz_opt = max(0, utilidad_matriz_opt)
        impuesto_matriz_opt = utilidad_matriz_opt * tasa_general
        pc_matriz_opt = tasa_pc * ingresos_matriz
        saldo_matriz_opt = impuesto_matriz_opt - pc_matriz_opt

        ir_grupo_opt = impuesto_matriz_opt + total_ir_opt_sats
        pc_grupo_opt = pc_matriz_opt + total_pc_opt_sats
        saldo_grupo_opt = ir_grupo_opt - pc_grupo_opt
        eficiencia_opt = (ir_grupo_opt / pc_grupo_opt * 100) if pc_grupo_opt > 0 else 0
        utilidad_opt_sats = r['grupo']['total_utilidad_satelites']
        utilidad_total_opt = utilidad_matriz_opt + utilidad_opt_sats
        tasa_efectiva_opt = (ir_grupo_opt / utilidad_total_opt * 100) if utilidad_total_opt > 0 else 0
        ahorro_opt_total = impuesto_sin_estructura - ir_grupo_opt
        ahorro_opt_pct = (ahorro_opt_total / impuesto_sin_estructura * 100) if impuesto_sin_estructura > 0 else 0

        eficiencia_sin = (impuesto_sin_estructura / pc_matriz_sin * 100) if pc_matriz_sin > 0 else 0
        tasa_efectiva_sin = (impuesto_sin_estructura / utilidad_sin_estructura * 100) if utilidad_sin_estructura > 0 else 0

        self.datos_pagos_cuenta = {
            'tasa_pago_cuenta': tasa_pc * 100,
            'satelites': resultados_pc,
            'global': {
                'nombre_matriz': r['matriz']['nombre'],
                'ingresos_matriz': r['matriz']['ingresos'],
                'impuesto_sin_estructura': impuesto_sin_estructura,
                'utilidad_sin_estructura': utilidad_sin_estructura,
                'pc_matriz_sin': pc_matriz_sin,
                'saldo_matriz_sin': saldo_matriz_sin,
                'tasa_efectiva_sin': tasa_efectiva_sin,
                'eficiencia_sin': eficiencia_sin,
                # Matriz en equilibrio
                'utilidad_matriz_eq': utilidad_matriz_eq,
                'impuesto_matriz_eq': impuesto_matriz_eq,
                'pc_matriz_eq': pc_matriz_eq,
                'saldo_matriz_eq': saldo_matriz_eq,
                # Equilibrio (principal)
                'ir_grupo_eq': ir_grupo_eq,
                'pc_grupo_eq': pc_grupo_eq,
                'saldo_grupo_eq': saldo_grupo_eq,
                'eficiencia_eq': eficiencia_eq,
                'tasa_efectiva_eq': tasa_efectiva_eq,
                'ahorro_eq_total': ahorro_eq_total,
                'ahorro_eq_pct': ahorro_eq_pct,
                'total_ir_eq_sats': total_ir_eq_sats,
                'total_pc_eq_sats': total_pc_eq_sats,
                'total_saldo_eq_sats': total_saldo_eq_sats,
                # Matriz en margen optimo
                'utilidad_matriz_opt': utilidad_matriz_opt,
                'impuesto_matriz_opt': impuesto_matriz_opt,
                'pc_matriz_opt': pc_matriz_opt,
                'saldo_matriz_opt': saldo_matriz_opt,
                # Margen optimo (referencia)
                'ir_grupo_opt': ir_grupo_opt,
                'pc_grupo_opt': pc_grupo_opt,
                'saldo_grupo_opt': saldo_grupo_opt,
                'eficiencia_opt': eficiencia_opt,
                'tasa_efectiva_opt': tasa_efectiva_opt,
                'ahorro_opt_total': ahorro_opt_total,
                'ahorro_opt_pct': ahorro_opt_pct,
                'total_ir_opt_sats': total_ir_opt_sats,
                'total_pc_opt_sats': total_pc_opt_sats,
                'total_saldo_opt_sats': total_saldo_opt_sats,
            },
        }

        self._mostrar_resultados_pagos_cuenta()
        self._generar_graficos_pagos_cuenta()

        n_eq = sum(1 for rpc in resultados_pc if rpc['margen_equilibrio'] is not None)
        messagebox.showinfo("Analisis de Equilibrio Financiero",
                            f"Punto de equilibrio calculado para {n_eq}/{len(resultados_pc)} satelites.\n"
                            f"Eficiencia PaC en Equilibrio: {eficiencia_eq:.1f}% (ideal=100%)\n"
                            f"Saldo Grupo en Equilibrio: {saldo_grupo_eq:,.2f}")
