import numpy as np


class SimEqLoopMixin:
    def _simular_eq_mc_iterations(self, n_sim, variabilidad,
                                   ingresos_base, costos_base,
                                   tasa_gral, tasa_esp, tasa_pc,
                                   limite_util, limite_ing,
                                   var_m_c, pct_c, var_m_i, pct_i,
                                   distribucion, rho, n_sats):
        """Ejecuta iteraciones Monte Carlo para equilibrio financiero."""
        ingresos_matriz_base = ingresos_base
        costos_matriz_base = costos_base
        var_matriz_costos = var_m_c
        pct_var_costos = pct_c
        var_matriz_ingresos = var_m_i
        pct_var_ingresos = pct_i
        # 2. ESTRUCTURA PARA ALMACENAR HISTORIAL DE SALDOS
        # --------------------------------------------------------------
        # Cada clave guarda dos listas: 'opt' (escenario óptimo) y 'eq' (equilibrio)
        historial_saldos = {
            'Matriz': {'opt': [], 'eq': []},
            'Grupo_Total': {'opt': [], 'eq': []}
        }
        for sat in self.resultados['satelites']:
            historial_saldos[sat['nombre']] = {'opt': [], 'eq': []}

        # --------------------------------------------------------------
        # 3. BUCLE PRINCIPAL DE MONTE CARLO
        # --------------------------------------------------------------
        for _ in range(n_sim):
            # A) Factores aleatorios correlacionados
            f_costos_m, f_ingresos_m, factores_sat = self._generar_factores_simulacion(
                n_sats, variabilidad,
                var_matriz_costos, pct_var_costos,
                var_matriz_ingresos, pct_var_ingresos,
                distribucion, rho
            )

            # B) Simular la matriz base
            ingresos_sim = ingresos_matriz_base * f_ingresos_m
            costos_m_sim = costos_matriz_base * f_costos_m
            util_base_sim = max(0, ingresos_sim - costos_m_sim)  # utilidad antes de compras a satélites
            pac_matriz = tasa_pc * ingresos_sim   # pago a cuenta fijo (depende de ingresos)

            # Acumuladores para el grupo
            saldo_grupo_opt = 0.0
            saldo_grupo_eq = 0.0
            compras_opt_total = 0.0
            compras_eq_total = 0.0
            costos_sats_total = 0.0

            # C) Simular cada satélite
            for idx, sat in enumerate(self.resultados['satelites']):
                factor = factores_sat[idx]
                costo_sim = sat['costo'] * factor
                gastos_sim = sat['gastos_operativos'] * factor
                es_gral = sat['regimen'].startswith('GENERAL')
                costos_sats_total += costo_sim

                # ========== ESCENARIO 1: MARGEN ÓPTIMO (maximiza ahorro) ==========
                if costo_sim > 0:
                    if es_gral:
                        # RÉGIMEN GENERAL: margen de equilibrio operativo (utilidad neta cero)
                        m_opt = (gastos_sim / costo_sim) if gastos_sim > 0 else 0.0
                    else:
                        # RÉGIMEN ESPECIAL: fórmula estándar (maximiza utilidad hasta el límite)
                        m_opt = (limite_util + gastos_sim) / costo_sim
                        # Ajuste por límite de ingresos
                        if costo_sim * (1 + m_opt) > limite_ing:
                            m_opt = (limite_ing / costo_sim) - 1
                else:
                    m_opt = 0.0

                precio_opt = costo_sim * (1 + m_opt)
                util_opt = max(0, precio_opt - costo_sim - gastos_sim)

                # Impuesto según régimen (óptimo)
                if es_gral:
                    ir_opt = util_opt * tasa_gral
                elif util_opt <= limite_util:
                    ir_opt = util_opt * tasa_esp
                else:
                    ir_opt = limite_util * tasa_esp + (util_opt - limite_util) * tasa_gral

                pac_opt = tasa_pc * precio_opt
                saldo_opt = ir_opt - pac_opt

                compras_opt_total += precio_opt
                saldo_grupo_opt += saldo_opt
                historial_saldos[sat['nombre']]['opt'].append(saldo_opt)

                # ========== ESCENARIO 2: MARGEN DE EQUILIBRIO (IR = PaC) ==========
                m_eq, reg_eq = self._calcular_margen_equilibrio_exacto(
                    costo_sim, gastos_sim, tasa_pc, tasa_esp, tasa_gral,
                    limite_util, limite_ing, es_gral
                )
                if m_eq is None:
                    m_eq = m_opt   # fallback al margen óptimo si no hay equilibrio factible

                precio_eq = costo_sim * (1 + m_eq)
                util_eq = max(0, precio_eq - costo_sim - gastos_sim)
                ir_eq = self._calcular_ir_equilibrio(
                    util_eq, tasa_esp, tasa_gral, limite_util, reg_eq or "General"
                )
                pac_eq = tasa_pc * precio_eq
                saldo_eq = ir_eq - pac_eq

                compras_eq_total += precio_eq
                saldo_grupo_eq += saldo_eq
                historial_saldos[sat['nombre']]['eq'].append(saldo_eq)

            # D) Calcular la matriz final para ambos escenarios
            # ---- Escenario Óptimo ----
            util_matriz_opt = max(0, util_base_sim - compras_opt_total + costos_sats_total)
            imp_matriz_opt = util_matriz_opt * tasa_gral
            saldo_matriz_opt = imp_matriz_opt - pac_matriz

            # ---- Escenario Equilibrio ----
            util_matriz_eq = max(0, util_base_sim - compras_eq_total + costos_sats_total)
            imp_matriz_eq = util_matriz_eq * tasa_gral
            saldo_matriz_eq = imp_matriz_eq - pac_matriz

            # Guardar historiales de matriz y grupo
            historial_saldos['Matriz']['opt'].append(saldo_matriz_opt)
            historial_saldos['Matriz']['eq'].append(saldo_matriz_eq)

            historial_saldos['Grupo_Total']['opt'].append(saldo_matriz_opt + saldo_grupo_opt)
            historial_saldos['Grupo_Total']['eq'].append(saldo_matriz_eq + saldo_grupo_eq)


        return historial_saldos
