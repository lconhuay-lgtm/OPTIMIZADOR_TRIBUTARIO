import tkinter as tk
from datetime import datetime


class ResultadosDisplayMixin:
    def mostrar_resultados(self):
        self.text_resultados.delete(1.0, tk.END)
        
        r = self.resultados
        p = r['parametros']
        
        texto = f"""
{'='*110}
                            REPORTE DE OPTIMIZACIÓN TRIBUTARIA CON AJUSTE POR GASTOS
                                    ANÁLISIS DE GRUPOS EMPRESARIALES
{'='*110}
Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

PARÁMETROS TRIBUTARIOS:
  UIT:                                 {p['uit']:>15,.2f}
  Tasa Régimen General:                {p['tasa_general']:>15.2f}%
  Tasa Régimen Especial:               {p['tasa_especial']:>15.2f}%
  Diferencial:                         {p['tasa_general'] - p['tasa_especial']:>15.2f} puntos
  Límite Utilidad Especial:            {p['limite_utilidad']:>15,.2f}
  Límite Ingresos Especial:            {p['limite_ingresos']:>15,.2f}
{'='*110}

ESCENARIO 1: SIN ESTRUCTURA SATELITAL
{'='*110}
{r['matriz']['nombre']} - Régimen General
  Ingresos:                            {r['matriz']['ingresos']:>15,.2f}
  Costos Externos:                     {r['matriz']['costos_externos']:>15,.2f}
  Utilidad:                            {r['matriz']['utilidad_sin_estructura']:>15,.2f}
  Impuesto:                            {r['matriz']['impuesto_sin_estructura']:>15,.2f}
  Utilidad Neta:                       {r['matriz']['utilidad_sin_estructura'] - r['matriz']['impuesto_sin_estructura']:>15,.2f}

{'='*110}
ESCENARIO 2: CON ESTRUCTURA SATELITAL OPTIMIZADA (AJUSTE POR GASTOS OPERATIVOS)
{'='*110}

EMPRESAS SATÉLITES
{'-'*110}
"""
        
        for i, sat in enumerate(r['satelites'], 1):
            texto += f"""
{i}. {sat['nombre']} - {sat['tipo']} [{sat['regimen']}]
   Costos Compra:                      {sat['costo']:>15,.2f}
   Gastos Operativos:                  {sat['gastos_operativos']:>15,.2f}
   Margen Óptimo:                      {sat['margen_optimo']:>15.2f}%
   Precio Venta:                       {sat['precio_venta']:>15,.2f}
   Utilidad Bruta:                     {sat['utilidad_bruta']:>15,.2f}
   Utilidad Neta:                      {sat['utilidad_neta']:>15,.2f}
   Impuesto (Tasa: {sat['tasa_efectiva']:.2f}%):           {sat['impuesto']:>15,.2f}
   Ahorro Individual:                  {sat['ahorro_individual']:>15,.2f}
   ROI:                                {sat['rentabilidad']:>15.2f}%
   Pago a Cuenta ({self.TASA_PAGO_CUENTA.get():.1f}%):            {sat['pago_a_cuenta']:>15,.2f}
   Saldo Regularizacion:               {sat['saldo_regularizacion']:>15,.2f}  {'(Favor)' if sat['saldo_regularizacion'] < 0 else '(Pagar)'}
   Margen Equilibrio (IR=PaC):         {f"{sat['margen_equilibrio_pc']:.4f}%" if sat['margen_equilibrio_pc'] is not None else "N/A":>15}
   Ajuste: {sat['ajuste_aplicado']}
"""
        
        utilidad_total = r['matriz']['nueva_utilidad'] + r['grupo']['total_utilidad_satelites']
        
        texto += f"""
{'-'*110}
CONSOLIDADO SATELITES:
  Total Utilidad Neta:                 {r['grupo']['total_utilidad_satelites']:>15,.2f}
  Total Impuesto:                      {r['grupo']['total_impuesto_satelites']:>15,.2f}
  Ahorro Total Satelites:              {sum([s['ahorro_individual'] for s in r['satelites']]):>15,.2f}
  Total Pagos a Cuenta:                {sum([s['pago_a_cuenta'] for s in r['satelites']]):>15,.2f}
  Total Saldo Regularizacion:          {sum([s['saldo_regularizacion'] for s in r['satelites']]):>15,.2f}

{'-'*110}
{r['matriz']['nombre']} (con estructura):
  Compras a Satélites:                 {r['matriz']['total_compras']:>15,.2f}
  Nueva Utilidad:                      {r['matriz']['nueva_utilidad']:>15,.2f}
  Impuesto:                            {r['matriz']['impuesto_con_estructura']:>15,.2f}

{'='*110}
CONSOLIDADO GRUPO
{'='*110}
  Utilidad Total:                      {utilidad_total:>15,.2f}
  Impuesto Total:                      {r['grupo']['impuesto_total']:>15,.2f}
  
  AHORRO TRIBUTARIO:                   {r['grupo']['ahorro_tributario']:>15,.2f}
  AHORRO PORCENTUAL:                   {r['grupo']['ahorro_porcentual']:>15.2f}%
  
  Tasa Efectiva Grupo:                 {(r['grupo']['impuesto_total'] / utilidad_total * 100):>15.2f}%
  Reducción vs Sin Estructura:         {p['tasa_general'] - (r['grupo']['impuesto_total'] / utilidad_total * 100):>15.2f} puntos

{'='*110}
FÓRMULA APLICADA (MEJORADA CON GASTOS):
{'='*110}
Margen Óptimo = (Límite_Utilidad + Gastos_Operativos) / Costos

Esta fórmula garantiza que la utilidad neta (después de gastos) alcance exactamente el límite
permitido en el régimen especial, maximizando el ahorro tributario.
{'='*110}
"""
        
        self.text_resultados.insert(1.0, texto)
    
