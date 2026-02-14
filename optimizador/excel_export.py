import tempfile
from datetime import datetime
from tkinter import messagebox, filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment


class ExcelExportMixin:
    def exportar_excel(self):
        try:
            if not hasattr(self, 'resultados'):
                messagebox.showwarning("Atencion", "Ejecute el calculo primero")
                return

            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"Reporte_Gerencial_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            )
            if not path:
                return

            r = self.resultados
            utilidad_total = r['matriz']['nueva_utilidad'] + r['grupo']['total_utilidad_satelites']
            tasa_efectiva = (r['grupo']['impuesto_total'] / utilidad_total * 100) if utilidad_total > 0 else 0
            roi = (r['grupo']['ahorro_tributario'] / r['matriz']['total_compras'] * 100) if r['matriz']['total_compras'] > 0 else 0

            wb = Workbook()

            with tempfile.TemporaryDirectory() as tmpdir:
                img_resultados = self._guardar_figura_temp(self.fig_resultados, 'resultados', tmpdir)
                img_simulacion = self._guardar_figura_temp(self.fig_simulacion, 'simulacion', tmpdir)
                img_sensibilidad = self._guardar_figura_temp(self.fig_sensibilidad, 'sensibilidad', tmpdir)
                img_comparativo = self._guardar_figura_temp(self.fig_comparativo, 'comparativo', tmpdir)
                img_pagos_cuenta = self._guardar_figura_temp(self.fig_pagos_cuenta, 'pagos_cuenta', tmpdir)
                img_equilibrio_sim = self._guardar_figura_temp(self.fig_equilibrio_sim, 'equilibrio_sim', tmpdir)

                imgs = {
                    'resultados': img_resultados, 'simulacion': img_simulacion,
                    'sensibilidad': img_sensibilidad, 'comparativo': img_comparativo,
                    'pagos_cuenta': img_pagos_cuenta, 'equilibrio_sim': img_equilibrio_sim,
                }

                self._exportar_ws1(wb, r, tasa_efectiva, roi)
                self._exportar_ws2(wb, r)
                self._exportar_ws3(wb, imgs)
                self._exportar_ws4(wb, imgs)
                self._exportar_ws5(wb, imgs)
                self._exportar_ws6(wb, imgs)
                self._exportar_ws7(wb, r, imgs)
                self._exportar_ws8(wb, r)
                self._exportar_ws9(wb, r)
                self._exportar_ws10(wb, imgs)

                wb.save(path)

            messagebox.showinfo(
                "Exportacion Completa",
                f"Reporte Gerencial exportado exitosamente.\n\n"
                f"Hojas generadas:\n"
                f"  1. Resumen Ejecutivo (+ Equilibrio Financiero)\n"
                f"  2. Detalle Satelites (+ Metricas Equilibrio)\n"
                f"  3. Graficos Resultados{'  [OK]' if img_resultados else '  [Pendiente]'}\n"
                f"  4. Simulacion Monte Carlo{'  [OK]' if self.datos_simulacion else '  [Pendiente]'}\n"
                f"  5. Sensibilidad{'  [OK]' if img_sensibilidad else '  [Pendiente]'}\n"
                f"  6. Comparativo{'  [OK]' if img_comparativo else '  [Pendiente]'}\n"
                f"  7. Pagos a Cuenta y Equilibrio{'  [OK]' if self.datos_pagos_cuenta else '  [Pendiente]'}\n"
                f"  8. Sensibilidad Margenes\n"
                f"  9. Parametros\n"
                f" 10. Sim. Equilibrio (Monte Carlo){'  [OK]' if self.datos_equilibrio_sim else '  [Pendiente]'}\n\n"
                f"Archivo: {path}"
            )
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            messagebox.showerror("Error de Exportacion", f"No se pudo generar el Excel:\n{str(e)}\n\nDetalle:\n{tb[-500:]}")
