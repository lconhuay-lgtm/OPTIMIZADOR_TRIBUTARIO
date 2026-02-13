import numpy as np
from tkinter import messagebox


class CalculoMargenesMixin:
    def calcular_margenes_optimos(self):
        try:
            # Parámetros
            uit = self.UIT.get()
            tasa_general = self.TASA_GENERAL.get() / 100
            tasa_especial = self.TASA_ESPECIAL.get() / 100
            limite_utilidad = self.LIMITE_UTILIDAD_UIT.get() * uit
            limite_ingresos = self.LIMITE_INGRESOS_UIT.get() * uit
            
            # Matriz
            nombre_matriz = self.entry_nombre_matriz.get()
            ingresos = float(self.entry_ingresos_matriz.get())
            costos_externos = float(self.entry_costos_matriz.get())
            
            utilidad_sin_estructura = ingresos - costos_externos
            impuesto_sin_estructura = utilidad_sin_estructura * tasa_general
            
            # Satélites
            satelites_data = []
            advertencias = []
            
            for idx, sat in enumerate(self.satelites_entries, 1):
                nombre = sat['entry_nombre'].get()
                costo = float(sat['entry_costo'].get())
                tipo = sat['combo_tipo'].get()
                gastos = float(sat['entry_gastos'].get())
                usar_regimen_general = sat['var_regimen_general'].get()
                
                # VERIFICAR SI PUEDE ESTAR EN RÉGIMEN ESPECIAL
                margen_para_limite = (limite_utilidad + gastos) / costo
                precio_venta_estimado = costo * (1 + margen_para_limite)
                
                puede_regimen_especial = True
                razon_no_puede = ""
                
                # Verificar si el margen sería muy alto (> 100%)
                if margen_para_limite > 1.0:
                    puede_regimen_especial = False
                    razon_no_puede = f"Margen requerido muy alto ({margen_para_limite*100:.1f}%)"
                
                # Verificar límite de ingresos
                if precio_venta_estimado > limite_ingresos:
                    puede_regimen_especial = False
                    razon_no_puede = f"Supera límite ingresos"
                
                # ADVERTENCIA: Si marca régimen general pudiendo estar en especial
                if usar_regimen_general and puede_regimen_especial and gastos == 0:
                    ahorro_potencial = limite_utilidad * (tasa_general - tasa_especial)
                    advertencias.append(
                        f"⚠️ {nombre}: Marcado RÉGIMEN GENERAL sin necesidad.\n"
                        f"   Ahorro potencial perdido: {ahorro_potencial:,.2f}\n"
                        f"   RECOMENDACIÓN: Desmarcar y usar Régimen Especial\n"
                    )
                
                # ADVERTENCIA: Si NO puede especial pero no está marcado
                if not usar_regimen_general and not puede_regimen_especial:
                    advertencias.append(
                        f"⚠️ {nombre}: NO califica para Régimen Especial.\n"
                        f"   Razón: {razon_no_puede}\n"
                        f"   Se aplicará automáticamente Régimen General\n"
                    )
                    usar_regimen_general = True
                
                satelites_data.append({
                    'nombre': nombre,
                    'tipo': tipo,
                    'costo': costo,
                    'gastos_operativos': gastos,
                    'usar_regimen_general': usar_regimen_general,
                    'puede_regimen_especial': puede_regimen_especial,
                    'razon_no_puede': razon_no_puede
                })
            
            # Mostrar advertencias
            if advertencias:
                mensaje = "ADVERTENCIAS DEL SISTEMA:\n\n" + "\n".join(advertencias) + "\n\n¿Continuar?"
                if messagebox.askquestion("Advertencias", mensaje, icon='warning') == 'no':
                    return
            
            # CALCULAR MÁRGENES ÓPTIMOS
            resultados_satelites = []
            total_utilidad_satelites = 0
            total_impuesto_satelites = 0
            total_compras_matriz = 0
            
            for sat in satelites_data:
                costo = sat['costo']
                gastos = sat['gastos_operativos']
                usar_regimen_general = sat['usar_regimen_general']
                
                if usar_regimen_general:
                    # ============================================
                    # RÉGIMEN GENERAL - LÓGICA MEJORADA
                    # ============================================
                    
                    if usar_regimen_general:
                        # ============================================
                        # RÉGIMEN GENERAL - LÓGICA CORREGIDA
                        # ============================================
                        
                        if gastos > 0:
                            # CON GASTOS: Margen al punto de equilibrio
                            margen_optimo = gastos / costo
                            utilidad_bruta = costo * margen_optimo
                            utilidad_neta = 0  # Punto equilibrio
                            precio_venta = costo * (1 + margen_optimo)
                            
                            impuesto_satelite = 0
                            tasa_efectiva = 0
                            
                            # AHORRO: KALLPA deja de pagar 29.5% sobre el gasto
                            ahorro_individual = gastos * tasa_general
                            
                            regimen = "GENERAL (Punto Equilibrio)"
                            ajuste_aplicado = f"Margen {margen_optimo*100:.2f}% para cubrir gastos ({gastos:,.0f}). Ahorro: {ahorro_individual:,.0f}"
                            
                        else:
                            # SIN GASTOS: Margen CERO (solo traspaso)
                            margen_optimo = 0.0
                            utilidad_bruta = 0
                            utilidad_neta = 0
                            precio_venta = costo  # Solo traspasa el costo
                            
                            impuesto_satelite = 0
                            tasa_efectiva = 0
                            ahorro_individual = 0
                            
                            regimen = "GENERAL (Traspaso sin margen)"
                            ajuste_aplicado = "⚠️ SIN MARGEN - No genera ahorro. RECOMENDACIÓN: Eliminar de estructura o desmarcar régimen general"
                    
                else:
                    # ============================================
                    # RÉGIMEN ESPECIAL - LÓGICA ORIGINAL
                    # ============================================
                    
                    # Fórmula mejorada: m* = (Limite_Utilidad + Gastos) / Costos
                    margen_optimo = (limite_utilidad + gastos) / costo
                    
                    # Verificar límite de ingresos
                    precio_venta_tentativo = costo * (1 + margen_optimo)
                    if precio_venta_tentativo > limite_ingresos:
                        margen_optimo = (limite_ingresos / costo) - 1
                        ajuste_aplicado = f"Ajustado por límite ingresos"
                    else:
                        if gastos > 0:
                            ajuste_aplicado = f"Margen incrementado por gastos ({gastos:,.0f})"
                        else:
                            ajuste_aplicado = "Margen óptimo estándar"
                    
                    utilidad_bruta = costo * margen_optimo
                    utilidad_neta = utilidad_bruta - gastos
                    precio_venta = costo * (1 + margen_optimo)
                    
                    # Calcular impuesto
                    if utilidad_neta <= limite_utilidad:
                        impuesto_satelite = utilidad_neta * tasa_especial
                        tasa_efectiva = tasa_especial * 100
                    else:
                        impuesto_base = limite_utilidad * tasa_especial
                        impuesto_exceso = (utilidad_neta - limite_utilidad) * tasa_general
                        impuesto_satelite = impuesto_base + impuesto_exceso
                        tasa_efectiva = (impuesto_satelite / utilidad_neta * 100) if utilidad_neta > 0 else 0
                    
                    # Ahorro: Diferencial de tasas
                    impuesto_si_fuera_general = utilidad_neta * tasa_general
                    ahorro_individual = impuesto_si_fuera_general - impuesto_satelite
                    
                    regimen = "ESPECIAL"
                
                # Pagos a cuenta (1.5% del precio de venta)
                tasa_pc = self.TASA_PAGO_CUENTA.get() / 100
                pago_a_cuenta = tasa_pc * precio_venta
                saldo_regularizacion = impuesto_satelite - pago_a_cuenta

                # Margen de equilibrio: donde IR = pago a cuenta
                margen_equilibrio_pc = None
                if not usar_regimen_general and (tasa_especial - tasa_pc) > 0:
                    m_eq = (tasa_pc * costo + tasa_especial * gastos) / (costo * (tasa_especial - tasa_pc))
                    util_eq = costo * m_eq - gastos
                    if util_eq <= limite_utilidad and costo * (1 + m_eq) <= limite_ingresos:
                        margen_equilibrio_pc = m_eq * 100
                elif usar_regimen_general and (tasa_general - tasa_pc) > 0:
                    m_eq = (tasa_pc * costo + tasa_general * gastos) / (costo * (tasa_general - tasa_pc))
                    margen_equilibrio_pc = m_eq * 100

                # Acumular totales
                total_utilidad_satelites += max(0, utilidad_neta)
                total_impuesto_satelites += impuesto_satelite
                total_compras_matriz += precio_venta

                resultados_satelites.append({
                    'nombre': sat['nombre'],
                    'tipo': sat['tipo'],
                    'regimen': regimen,
                    'costo': costo,
                    'gastos_operativos': gastos,
                    'margen_optimo': margen_optimo * 100,
                    'precio_venta': precio_venta,
                    'utilidad_bruta': utilidad_bruta,
                    'utilidad_neta': max(0, utilidad_neta),
                    'impuesto': impuesto_satelite,
                    'tasa_efectiva': tasa_efectiva,
                    'rentabilidad': (max(0, utilidad_neta) / precio_venta * 100) if precio_venta > 0 else 0,
                    'ahorro_individual': ahorro_individual,
                    'ajuste_aplicado': ajuste_aplicado,
                    'pago_a_cuenta': pago_a_cuenta,
                    'saldo_regularizacion': saldo_regularizacion,
                    'margen_equilibrio_pc': margen_equilibrio_pc
                })
            
            # Calcular nueva utilidad matriz
            nueva_utilidad_matriz = utilidad_sin_estructura - total_compras_matriz + sum([s['costo'] for s in satelites_data])
            impuesto_matriz = nueva_utilidad_matriz * tasa_general
            
            # Totales
            impuesto_total_grupo = impuesto_matriz + total_impuesto_satelites
            ahorro_tributario = impuesto_sin_estructura - impuesto_total_grupo
            ahorro_porcentual = (ahorro_tributario / impuesto_sin_estructura * 100) if impuesto_sin_estructura > 0 else 0
            
            # Almacenar resultados
            self.resultados = {
                'parametros': {
                    'uit': uit,
                    'tasa_general': tasa_general * 100,
                    'tasa_especial': tasa_especial * 100,
                    'limite_utilidad': limite_utilidad,
                    'limite_ingresos': limite_ingresos
                },
                'matriz': {
                    'nombre': nombre_matriz,
                    'ingresos': ingresos,
                    'costos_externos': costos_externos,
                    'utilidad_sin_estructura': utilidad_sin_estructura,
                    'impuesto_sin_estructura': impuesto_sin_estructura,
                    'nueva_utilidad': nueva_utilidad_matriz,
                    'impuesto_con_estructura': impuesto_matriz,
                    'total_compras': total_compras_matriz
                },
                'satelites': resultados_satelites,
                'grupo': {
                    'total_utilidad_satelites': total_utilidad_satelites,
                    'total_impuesto_satelites': total_impuesto_satelites,
                    'impuesto_total': impuesto_total_grupo,
                    'ahorro_tributario': ahorro_tributario,
                    'ahorro_porcentual': ahorro_porcentual
                }
            }
            
            self.mostrar_resultados()
            self.generar_graficos()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Valores inválidos:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")

    
