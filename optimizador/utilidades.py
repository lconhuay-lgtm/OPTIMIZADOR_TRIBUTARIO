import json
from tkinter import messagebox


class UtilidadesMixin:
    def copiar_resultados(self):
        try:
            if hasattr(self, 'resultados'):
                self.root.clipboard_clear()
                self.root.clipboard_append(self.text_resultados.get(1.0, tk.END))
                messagebox.showinfo("Éxito", "Resultados copiados")
        except Exception as e:
            messagebox.showerror("Error", f"Error: {str(e)}")
    
    def limpiar_datos(self):
        self.entry_nombre_matriz.delete(0, tk.END)
        self.entry_ingresos_matriz.delete(0, tk.END)
        self.entry_costos_matriz.delete(0, tk.END)
        
        for sat in self.satelites_entries:
            sat['entry_nombre'].delete(0, tk.END)
            sat['entry_costo'].delete(0, tk.END)
            sat['entry_gastos'].delete(0, tk.END)
            sat['var_regimen_general'].set(False)
        
        self.text_resultados.delete(1.0, tk.END)
    
    def cargar_ejemplo(self):
        self.entry_nombre_matriz.delete(0, tk.END)
        self.entry_nombre_matriz.insert(0, "KALLPA")
        
        self.entry_ingresos_matriz.delete(0, tk.END)
        self.entry_ingresos_matriz.insert(0, "36184868.55")
        
        self.entry_costos_matriz.delete(0, tk.END)
        self.entry_costos_matriz.insert(0, "31674202.9240117")
        
        nombres = ["IG PAME", "SUMAQ WARMY", "KUSKA KUSISJA", "MISKY"]
        costos = ["1714276", "7000000", "4500000", "2550000"]
        gastos = ["0", "0", "0", "500000"]
        
        for i, sat in enumerate(self.satelites_entries):
            if i < len(nombres):
                sat['entry_nombre'].delete(0, tk.END)
                sat['entry_nombre'].insert(0, nombres[i])
                
                sat['entry_costo'].delete(0, tk.END)
                sat['entry_costo'].insert(0, costos[i])
                
                sat['entry_gastos'].delete(0, tk.END)
                sat['entry_gastos'].insert(0, gastos[i])
                
                sat['var_regimen_general'].set(False)
        
        messagebox.showinfo("Ejemplo", "Datos cargados")

if __name__ == "__main__":
    root = tk.Tk()
    app = OptimizadorPro(root)
    root.mainloop()