import numpy as np
from tkinter import messagebox


class PagosCuentaMixin:
    def analizar_pagos_cuenta(self):
        """Analisis de punto de equilibrio financiero: donde IR = Pagos a Cuenta (saldo = 0, eficiencia = 100%).

        El margen de equilibrio es la metrica principal. El margen optimo (regimen) se muestra como referencia.
        Usa calcular_margen_equilibrio_exacto con 3 escenarios: Especial, Mixto MYPE, General.
        """
        try:
            if not hasattr(self, 'resultados'):
                messagebox.showwarning("Advertencia", "Primero debe calcular los margenes optimos")
                return

            tasa_pc = self.TASA_PAGO_CUENTA.get() / 100
            tasa_general = self.resultados['parametros']['tasa_general'] / 100
            tasa_especial = self.resultados['parametros']['tasa_especial'] / 100
            limite_utilidad = self.resultados['parametros']['limite_utilidad']
            limite_ingresos = self.resultados['parametros']['limite_ingresos']
            n_pasos = int(self.entry_pasos_sens_pc.get())

            resultados_pc = []

            for sat in self.resultados['satelites']:
                costo = sat['costo']
                gastos = sat['gastos_operativos']
                margen_actual = sat['margen_optimo'] / 100
                es_general = sat['regimen'].startswith('GENERAL')

                # === MARGEN DE EQUILIBRIO EXACTO (3 escenarios: A, B, C) ===
                margen_equilibrio, regimen_equilibrio = self._calcular_margen_equilibrio_exacto(
                    costo, gastos, tasa_pc, tasa_especial, tasa_general,
                    limite_utilidad, limite_ingresos, es_general
                )

                # === Metricas EN el punto de equilibrio ===
                if margen_equilibrio is not None:
                    precio_eq = costo * (1 + margen_equilibrio)
                    utilidad_neta_eq = max(0, costo * margen_equilibrio - gastos)
                    pago_cuenta_eq = tasa_pc * precio_eq
                    ir_eq = self._calcular_ir_equilibrio(
                        utilidad_neta_eq, tasa_especial, tasa_general,
                        limite_utilidad, regimen_equilibrio
                    )
                    saldo_eq = ir_eq - pago_cuenta_eq  # debe ser ~0
                    ahorro_eq = utilidad_neta_eq * tasa_general - ir_eq
                else:
                    precio_eq = 0
                    utilidad_neta_eq = 0
                    pago_cuenta_eq = 0
                    ir_eq = 0
                    saldo_eq = 0
                    ahorro_eq = 0

                # === Metricas EN el margen optimo (referencia) ===
                precio_opt = costo * (1 + margen_actual)
                pago_cuenta_opt = tasa_pc * precio_opt
                ir_opt = sat['impuesto']
                saldo_opt = ir_opt - pago_cuenta_opt
                eficiencia_opt = (ir_opt / pago_cuenta_opt * 100) if pago_cuenta_opt > 0 else 0

                # === Sensibilidad: variar margen centrado en el equilibrio ===
                centro = margen_equilibrio if margen_equilibrio is not None else margen_actual
                margen_min = max(0.001, centro * 0.2)
                margen_max_feasible = min(limite_ingresos / costo - 1, 1.0) if costo > 0 else 1.0
                margen_max = min(centro * 3.0, margen_max_feasible)
                if margen_max <= margen_min:
                    margen_max = margen_min + 0.10

                margenes_test = np.linspace(margen_min, margen_max, n_pasos)

                margenes_extras = [margen_actual]
                if margen_equilibrio is not None:
                    margenes_extras.append(margen_equilibrio)
                margenes_test = np.sort(np.unique(np.concatenate([margenes_test, margenes_extras])))

                sensibilidad = []
                for m in margenes_test:
                    precio = costo * (1 + m)
                    utilidad_neta = max(0, costo * m - gastos)

                    if es_general:
                        ir = utilidad_neta * tasa_general
                        regimen_m = "General"
                        ahorro_regimen = 0
                    elif utilidad_neta > limite_utilidad or precio > limite_ingresos:
                        if utilidad_neta > limite_utilidad:
                            ir = limite_utilidad * tasa_especial + (utilidad_neta - limite_utilidad) * tasa_general
                            regimen_m = "Mixto"
                        else:
                            ir = utilidad_neta * tasa_general
                            regimen_m = "General*"
                        ahorro_regimen = utilidad_neta * tasa_general - ir
                    else:
                        ir = utilidad_neta * tasa_especial
                        regimen_m = "Especial"
                        ahorro_regimen = utilidad_neta * (tasa_general - tasa_especial)

                    pago_cuenta = tasa_pc * precio
                    saldo = ir - pago_cuenta
                    efic = (ir / pago_cuenta * 100) if pago_cuenta > 0 else 0

                    sensibilidad.append({
                        'margen': m * 100,
                        'precio_venta': precio,
                        'utilidad_neta': utilidad_neta,
                        'regimen': regimen_m,
                        'ir': ir,
                        'pago_cuenta': pago_cuenta,
                        'saldo': saldo,
                        'ahorro_regimen': ahorro_regimen,
                        'eficiencia_pc': efic,
                        'es_optimo': abs(m - margen_actual) < 0.0001,
                        'es_equilibrio': margen_equilibrio is not None and abs(m - margen_equilibrio) < 0.0001,
                    })

                # Sensibilidad 2D: recalcular equilibrio exacto para cada tasa PaC
                tasas_pc_test = [0.5, 1.0, 1.5, 2.0, 2.5]
                sens_2d = []
                for t_pc in tasas_pc_test:
                    t_pc_dec = t_pc / 100
                    m_eq_2d, reg_2d = self._calcular_margen_equilibrio_exacto(
                        costo, gastos, t_pc_dec, tasa_especial, tasa_general,
                        limite_utilidad, limite_ingresos, es_general
                    )

                    if m_eq_2d is not None and m_eq_2d > 0:
                        precio_2d = costo * (1 + m_eq_2d)
                        util_2d = max(0, costo * m_eq_2d - gastos)
                        ir_2d = self._calcular_ir_equilibrio(
                            util_2d, tasa_especial, tasa_general, limite_utilidad, reg_2d
                        )
                        pc_2d = t_pc_dec * precio_2d
                        ahorro_2d = util_2d * tasa_general - ir_2d
                    else:
                        m_eq_2d = 0
                        precio_2d = 0
                        util_2d = 0
                        ir_2d = 0
                        pc_2d = 0
                        ahorro_2d = 0
                        reg_2d = "N/A"

                    sens_2d.append({
                        'tasa_pc': t_pc,
                        'margen_equilibrio': m_eq_2d * 100 if m_eq_2d else 0,
                        'precio_venta': precio_2d,
                        'utilidad_neta': util_2d,
                        'pago_cuenta': pc_2d,
                        'ir': ir_2d,
                        'saldo': ir_2d - pc_2d,
                        'ahorro_regimen': ahorro_2d,
                        'regimen': reg_2d,
                    })

                resultados_pc.append({
                    'nombre': sat['nombre'],
                    'costo': costo,
                    'gastos': gastos,
                    'es_general': es_general,
                    # Equilibrio (METRICA PRINCIPAL)
                    'margen_equilibrio': margen_equilibrio * 100 if margen_equilibrio is not None else None,
                    'regimen_equilibrio': regimen_equilibrio,
                    'precio_equilibrio': precio_eq,
                    'utilidad_neta_eq': utilidad_neta_eq,
                    'ir_equilibrio': ir_eq,
                    'pago_cuenta_eq': pago_cuenta_eq,
                    'saldo_equilibrio': saldo_eq,
                    'ahorro_equilibrio': ahorro_eq,
                    # Margen optimo (REFERENCIA)
                    'margen_optimo_regimen': margen_actual * 100,
                    'precio_optimo': precio_opt,
                    'ir_optimo': ir_opt,
                    'pago_cuenta_optimo': pago_cuenta_opt,
                    'saldo_optimo': saldo_opt,
                    'eficiencia_optimo': eficiencia_opt,
                    # Sensibilidad
                    'sensibilidad': sensibilidad,
                    'sensibilidad_2d': sens_2d,
                })

            self._analizar_global_pagos_cuenta(resultados_pc, tasa_pc, tasa_general)

        except Exception as e:
            messagebox.showerror("Error", f"Error en analisis de pagos a cuenta: {str(e)}")
