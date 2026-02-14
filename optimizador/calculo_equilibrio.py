import numpy as np


class CalculoEquilibrioMixin:
    def _calcular_margen_equilibrio_exacto(self, costo, gastos, tasa_pc, tasa_especial,
                                             tasa_general, limite_utilidad, limite_ingresos, es_general):
        """Calcula el margen de equilibrio exacto donde IR = PaC (saldo = 0).

        Prueba 3 escenarios en orden:
        A) Regimen Especial Puro: Util * t_esp = Ingreso * t_pc
        B) Regimen MYPE Mixto: L*t_esp + (Util-L)*t_gral = Ingreso * t_pc
        C) Regimen General Puro: Util * t_gral = Ingreso * t_pc

        Returns: (margen, regimen_str) o (None, None) si no hay solucion factible.
        """
        if costo <= 0:
            return None, None

        # --- Escenario C: Regimen General Puro ---
        # Util * t_gral = Ingreso * t_pc
        # (C*m - G) * t_gral = C*(1+m) * t_pc
        # m = (C*t_pc + G*t_gral) / (C*(t_gral - t_pc))
        m_eq_gral = None
        if (tasa_general - tasa_pc) > 0:
            m_eq_gral = (tasa_pc * costo + tasa_general * gastos) / (costo * (tasa_general - tasa_pc))

        # Si esta forzado a General, solo Escenario C
        if es_general:
            if m_eq_gral is not None and m_eq_gral > 0:
                util_c = costo * m_eq_gral - gastos
                if util_c >= 0:
                    return m_eq_gral, "General"
            return None, None

        # --- Escenario A: Regimen Especial Puro ---
        # Util * t_esp = Ingreso * t_pc
        # (C*m - G) * t_esp = C*(1+m) * t_pc
        # m = (C*t_pc + G*t_esp) / (C*(t_esp - t_pc))
        if (tasa_especial - tasa_pc) > 0:
            m_eq_esp = (tasa_pc * costo + tasa_especial * gastos) / (costo * (tasa_especial - tasa_pc))
            if m_eq_esp > 0:
                util_a = costo * m_eq_esp - gastos
                precio_a = costo * (1 + m_eq_esp)
                if util_a >= 0 and util_a <= limite_utilidad and precio_a <= limite_ingresos:
                    return m_eq_esp, "Especial"

        # --- Escenario B: Regimen MYPE Mixto (LA CORRECCION CLAVE) ---
        # L*t_esp + (Util - L)*t_gral = Ingreso * t_pc
        # L*t_esp + (C*m - G - L)*t_gral = C*(1+m)*t_pc
        # C*m*t_gral - C*m*t_pc = C*t_pc + G*t_gral + L*t_gral - L*t_esp
        # m = [C*t_pc + G*t_gral + L*(t_gral - t_esp)] / [C*(t_gral - t_pc)]
        if (tasa_general - tasa_pc) > 0:
            m_eq_mixto = (tasa_pc * costo + tasa_general * gastos + limite_utilidad * (tasa_general - tasa_especial)) / (costo * (tasa_general - tasa_pc))
            if m_eq_mixto > 0:
                util_b = costo * m_eq_mixto - gastos
                precio_b = costo * (1 + m_eq_mixto)
                # Valido solo si utilidad > limite (si no, seria Escenario A)
                if util_b > limite_utilidad and precio_b <= limite_ingresos:
                    return m_eq_mixto, "Mixto"

        # --- Fallback a Escenario C si nada mas funciona ---
        if m_eq_gral is not None and m_eq_gral > 0:
            util_c = costo * m_eq_gral - gastos
            if util_c >= 0:
                return m_eq_gral, "General"

        return None, None

    def _calcular_ir_equilibrio(self, utilidad_neta, tasa_especial, tasa_general,
                                 limite_utilidad, regimen_equilibrio):
        """Calcula el IR en el punto de equilibrio segun el regimen determinado."""
        if regimen_equilibrio == "Especial":
            return utilidad_neta * tasa_especial
        elif regimen_equilibrio == "Mixto":
            return limite_utilidad * tasa_especial + (utilidad_neta - limite_utilidad) * tasa_general
        else:  # General
            return utilidad_neta * tasa_general

    

    def _calcular_ahorro_grupo(self, satelites_info, tasa_general, tasa_especial,
                               limite_utilidad, limite_ingresos, utilidad_sin_estructura):
        """Recalcula el ahorro tributario total del grupo con parámetros dados.

        Replica la lógica exacta de calcular_margenes_optimos() para garantizar
        consistencia entre el cálculo base y el análisis de sensibilidad.
        """
        total_impuesto_satelites = 0
        total_compras_matriz = 0
        total_costos_satelites = 0

        for sat in satelites_info:
            costo = sat['costo']
            gastos = sat['gastos']
            es_general = sat['es_general']

            if es_general:
                if gastos > 0:
                    margen = gastos / costo
                    impuesto = 0
                else:
                    margen = 0
                    impuesto = 0
            else:
                margen = (limite_utilidad + gastos) / costo
                precio_tentativo = costo * (1 + margen)

                if precio_tentativo > limite_ingresos:
                    margen = (limite_ingresos / costo) - 1

                if margen > 1.0:
                    # No califica, forzar general
                    if gastos > 0:
                        margen = gastos / costo
                    else:
                        margen = 0
                    impuesto = 0
                else:
                    utilidad_bruta = costo * margen
                    utilidad_neta = utilidad_bruta - gastos

                    if utilidad_neta <= limite_utilidad:
                        impuesto = utilidad_neta * tasa_especial
                    else:
                        impuesto = (limite_utilidad * tasa_especial +
                                    (utilidad_neta - limite_utilidad) * tasa_general)

            precio_venta = costo * (1 + margen)
            total_impuesto_satelites += impuesto
            total_compras_matriz += precio_venta
            total_costos_satelites += costo

        nueva_utilidad_matriz = utilidad_sin_estructura - total_compras_matriz + total_costos_satelites
        impuesto_matriz = nueva_utilidad_matriz * tasa_general

        impuesto_total_grupo = impuesto_matriz + total_impuesto_satelites
        impuesto_sin_estructura = utilidad_sin_estructura * tasa_general

        return impuesto_sin_estructura - impuesto_total_grupo

