import tkinter as tk


class PagosCuentaDisplayMixin:
    def _mostrar_resultados_pagos_cuenta(self):
        """Muestra resultados priorizando el punto de equilibrio financiero (IR = PaC, saldo = 0)."""
        self.text_pagos_cuenta.delete(1.0, tk.END)
        d = self.datos_pagos_cuenta
        g = d['global']

        texto = f"""{'='*140}
  PUNTO DE EQUILIBRIO FINANCIERO - Donde IR = Pagos a Cuenta (Saldo = 0, Eficiencia = 100%)
  Tasa Pago a Cuenta: {d['tasa_pago_cuenta']:.1f}%  |  Grupo: {g['nombre_matriz']}
{'='*140}

  CONCEPTO: El punto de equilibrio financiero es el margen donde los pagos a cuenta adelantados
  al fisco ({d['tasa_pago_cuenta']:.1f}% de ingresos) igualan EXACTAMENTE al IR anual definitivo.
  En este punto: saldo de regularizacion = 0, eficiencia = 100%, flujo de caja optimizado.

  Formulas segun escenario:
  A) Especial Puro: m = (C*t_pc + G*t_esp) / (C*(t_esp - t_pc))       [si Util <= 15 UIT]
  B) Mixto MYPE:    m = (C*t_pc + G*t_gral + L*(t_gral-t_esp)) / (C*(t_gral-t_pc))  [si Util > 15 UIT e Ingreso <= 1700 UIT]
  C) General Puro:  m = (C*t_pc + G*t_gral) / (C*(t_gral - t_pc))     [si Ingreso > 1700 UIT]

{'='*140}
  I. IMPUESTO GLOBAL SIN ESTRUCTURA SATELITAL
{'='*140}
  Empresa:                        {g['nombre_matriz']}
  Ingresos Totales:               {g['ingresos_matriz']:>15,.2f}
  Utilidad Gravable:              {g['utilidad_sin_estructura']:>15,.2f}
  Tasa Aplicable:                 {'29.5% (Regimen General)':>30}
  Impuesto a la Renta:            {g['impuesto_sin_estructura']:>15,.2f}
  {'-'*70}
  Pagos a Cuenta ({d['tasa_pago_cuenta']:.1f}%):        {g['pc_matriz_sin']:>15,.2f}
  Saldo Regularizacion:           {g['saldo_matriz_sin']:>15,.2f}  {'(SALDO A FAVOR)' if g['saldo_matriz_sin'] < 0 else '(SALDO POR PAGAR)'}
  Eficiencia PaC:                 {g['eficiencia_sin']:>10.1f}%  (ideal=100%)
  Tasa Efectiva:                  {g['tasa_efectiva_sin']:>10.1f}%

{'='*140}
  II. IMPUESTO GLOBAL CON ESTRUCTURA EN PUNTO DE EQUILIBRIO (IR = PaC)
{'='*140}
  A) EMPRESA MATRIZ
  {'-'*70}
  Utilidad Matriz (ajustada):     {g['utilidad_matriz_eq']:>15,.2f}
  Impuesto Matriz:                {g['impuesto_matriz_eq']:>15,.2f}
  Pagos a Cuenta Matriz:          {g['pc_matriz_eq']:>15,.2f}
  Saldo Matriz:                   {g['saldo_matriz_eq']:>15,.2f}  {'(FAVOR)' if g['saldo_matriz_eq'] < 0 else '(PAGAR)'}

  B) SATELITES EN PUNTO DE EQUILIBRIO (saldo = 0 por satelite)
  {'-'*70}
  IR Satelites (en equilibrio):   {g['total_ir_eq_sats']:>15,.2f}
  PaC Satelites (en equilibrio):  {g['total_pc_eq_sats']:>15,.2f}
  Saldo Satelites:                {g['total_saldo_eq_sats']:>15,.2f}  (debe ser ~0)

  C) GRUPO CONSOLIDADO EN EQUILIBRIO
  {'-'*70}
  IR Total Grupo:                 {g['ir_grupo_eq']:>15,.2f}
  PaC Total Grupo:                {g['pc_grupo_eq']:>15,.2f}
  Saldo Grupo:                    {g['saldo_grupo_eq']:>15,.2f}  {'(~0 = EQUILIBRIO PERFECTO)' if abs(g['saldo_grupo_eq']) < 100 else '(FAVOR)' if g['saldo_grupo_eq'] < 0 else '(PAGAR)'}
  Eficiencia PaC:                 {g['eficiencia_eq']:>10.1f}%  (ideal=100%)
  Tasa Efectiva:                  {g['tasa_efectiva_eq']:>10.1f}%
  Ahorro vs Sin Estructura:       {g['ahorro_eq_total']:>15,.2f} ({g['ahorro_eq_pct']:.1f}%)

{'='*140}
  III. COMPARATIVO: SIN ESTRUCTURA vs EQUILIBRIO vs MARGEN OPTIMO
{'='*140}
  {'Concepto':<35} | {'SIN Estructura':>16} | {'EQUILIBRIO':>16} | {'Margen Optimo':>16} | {'Mejor Escenario':>16}
  {'-'*110}
  {'IR Total Grupo':<35} | {g['impuesto_sin_estructura']:>16,.0f} | {g['ir_grupo_eq']:>16,.0f} | {g['ir_grupo_opt']:>16,.0f} | {'Equilibrio' if g['ir_grupo_eq'] < g['ir_grupo_opt'] else 'M. Optimo':>16}
  {'PaC Total Grupo':<35} | {g['pc_matriz_sin']:>16,.0f} | {g['pc_grupo_eq']:>16,.0f} | {g['pc_grupo_opt']:>16,.0f} | {'-':>16}
  {'Saldo Regularizacion':<35} | {g['saldo_matriz_sin']:>16,.0f} | {g['saldo_grupo_eq']:>16,.0f} | {g['saldo_grupo_opt']:>16,.0f} | {'Equilibrio' if abs(g['saldo_grupo_eq']) < abs(g['saldo_grupo_opt']) else 'M. Optimo':>16}
  {'Eficiencia PaC (%)':<35} | {g['eficiencia_sin']:>15.1f}% | {g['eficiencia_eq']:>15.1f}% | {g['eficiencia_opt']:>15.1f}% | {'Equilibrio' if abs(g['eficiencia_eq'] - 100) < abs(g['eficiencia_opt'] - 100) else 'M. Optimo':>16}
  {'Tasa Efectiva (%)':<35} | {g['tasa_efectiva_sin']:>15.1f}% | {g['tasa_efectiva_eq']:>15.1f}% | {g['tasa_efectiva_opt']:>15.1f}% | {'Equilibrio' if g['tasa_efectiva_eq'] < g['tasa_efectiva_opt'] else 'M. Optimo':>16}
  {'Ahorro Tributario':<35} | {'-':>16} | {g['ahorro_eq_total']:>16,.0f} | {g['ahorro_opt_total']:>16,.0f} | {'Equilibrio' if g['ahorro_eq_total'] > g['ahorro_opt_total'] else 'M. Optimo':>16}
  {'Ahorro %':<35} | {'-':>16} | {g['ahorro_eq_pct']:>15.1f}% | {g['ahorro_opt_pct']:>15.1f}% | {'-':>16}
  {'='*110}

{'='*140}
  IV. DETALLE POR SATELITE - PUNTO DE EQUILIBRIO FINANCIERO
{'='*140}
"""

        for i, sat in enumerate(d['satelites'], 1):
            tiene_eq = sat['margen_equilibrio'] is not None
            eq_str = f"{sat['margen_equilibrio']:.4f}%" if tiene_eq else "N/A"
            texto += f"""
  {'-'*140}
  {i}. {sat['nombre']} | Costo: {sat['costo']:,.0f} | Gastos: {sat['gastos']:,.0f}
  {'-'*140}
  >>> PUNTO DE EQUILIBRIO (IR = PaC, Saldo = 0) <<<
  Margen Equilibrio:              {eq_str:>20}  {'   [Regimen: ' + sat['regimen_equilibrio'] + ']' if tiene_eq else '   [Sin equilibrio factible]'}
  Precio Venta en Equilibrio:     {sat['precio_equilibrio']:>15,.2f}
  Utilidad Neta en Equilibrio:    {sat['utilidad_neta_eq']:>15,.2f}
  IR en Equilibrio:               {sat['ir_equilibrio']:>15,.2f}
  PaC en Equilibrio:              {sat['pago_cuenta_eq']:>15,.2f}
  Saldo en Equilibrio:            {sat['saldo_equilibrio']:>15,.2f}  {'(~0 PERFECTO)' if abs(sat['saldo_equilibrio']) < 1 else ''}
  Ahorro vs Reg. General:         {sat['ahorro_equilibrio']:>15,.2f}

  --- Referencia: Margen Optimo (maximiza ahorro fiscal) ---
  Margen Optimo:                  {sat['margen_optimo_regimen']:>10.4f}%
  IR en Margen Optimo:            {sat['ir_optimo']:>15,.2f}
  PaC en Margen Optimo:           {sat['pago_cuenta_optimo']:>15,.2f}
  Saldo en Margen Optimo:         {sat['saldo_optimo']:>15,.2f}  {'(FAVOR)' if sat['saldo_optimo'] < 0 else '(PAGAR)'}
  Eficiencia PaC en Optimo:       {sat['eficiencia_optimo']:>10.1f}%
"""
            # Tabla de sensibilidad
            texto += f"""
  {'SENSIBILIDAD - Como varia el saldo al mover el margen':^140}
  El punto EQUILIBRIO tiene saldo=0 (eficiencia 100%). El punto OPTIMO maximiza el ahorro fiscal.
  {'-'*140}
  {'Margen%':>8} | {'Precio Venta':>14} | {'Util. Neta':>12} | {'Regimen':>10} | {'IR':>12} | {'Pago Cta':>12} | {'Saldo':>12} | {'Ahorro Reg':>12} | {'Efic%':>7} | {'Nota':>14}
  {'-'*140}
"""
            for s in sat['sensibilidad']:
                nota = ""
                if s['es_equilibrio']:
                    nota = "<<< EQUILIBRIO"
                elif s['es_optimo']:
                    nota = "(ref: optimo)"
                texto += f"  {s['margen']:>8.4f} | {s['precio_venta']:>14,.0f} | {s['utilidad_neta']:>12,.0f} | {s['regimen']:>10} | {s['ir']:>12,.0f} | {s['pago_cuenta']:>12,.0f} | {s['saldo']:>12,.0f} | {s['ahorro_regimen']:>12,.0f} | {s['eficiencia_pc']:>6.1f}% | {nota:>14}\n"

            # Tabla sensibilidad 2D
            texto += f"""
  {'SENSIBILIDAD - Punto de Equilibrio ante cambios en la Tasa de Pago a Cuenta':^140}
  Si SUNAT cambiara la tasa PaC, cual seria el nuevo margen de equilibrio para mantener saldo = 0.
  {'-'*120}
  {'Tasa PaC%':>10} | {'M. Equil%':>10} | {'Precio Eq.':>14} | {'Util. Neta':>12} | {'IR':>12} | {'PaC':>12} | {'Saldo':>12} | {'Ahorro Reg':>12} | {'Regimen':>10}
  {'-'*120}
"""
            for s2 in sat['sensibilidad_2d']:
                texto += f"  {s2['tasa_pc']:>10.1f} | {s2['margen_equilibrio']:>9.4f}% | {s2['precio_venta']:>14,.0f} | {s2['utilidad_neta']:>12,.0f} | {s2['ir']:>12,.0f} | {s2['pago_cuenta']:>12,.0f} | {s2['saldo']:>12,.0f} | {s2['ahorro_regimen']:>12,.0f} | {s2['regimen']:>10}\n"

        texto += f"""
{'='*140}
  V. INTERPRETACION ESTRATEGICA
{'='*140}

  PUNTO DE EQUILIBRIO FINANCIERO (concepto central):
  Cada satelite tiene un margen donde IR = PaC, lo que significa:
    - Saldo de regularizacion = 0 (no hay dinero inmovilizado ni falta liquidez)
    - Eficiencia del pago a cuenta = 100%
    - El flujo de caja es optimo: cada sol adelantado se consume exactamente

  EQUILIBRIO vs MARGEN OPTIMO:
  - El MARGEN OPTIMO maximiza el ahorro fiscal (diferencial de tasas).
  - El MARGEN DE EQUILIBRIO maximiza la eficiencia financiera (saldo = 0).
  - Si ambos coinciden: situacion ideal.
  - Si difieren: la gerencia debe decidir que priorizar:
      * Equilibrio: mejor flujo de caja, cero saldos pendientes
      * Optimo: mayor ahorro fiscal, pero posible saldo a favor/contra

  RESULTADO DEL GRUPO EN EQUILIBRIO:
  - IR Total: {g['ir_grupo_eq']:,.0f} | Ahorro: {g['ahorro_eq_total']:,.0f} ({g['ahorro_eq_pct']:.1f}%)
  - Eficiencia PaC: {g['eficiencia_eq']:.1f}% | Tasa Efectiva: {g['tasa_efectiva_eq']:.1f}%
{'='*140}
"""
        self.text_pagos_cuenta.insert(1.0, texto)

