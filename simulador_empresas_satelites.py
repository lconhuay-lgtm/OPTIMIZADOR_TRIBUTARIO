#!/usr/bin/env python3
"""
Optimizador Tributario Pro - Punto de entrada.

Arquitectura modular: toda la logica se encuentra en el paquete 'optimizador/'.
La clase OptimizadorPro se construye via herencia multiple de Mixins.
"""
import tkinter as tk
from optimizador import OptimizadorPro


if __name__ == "__main__":
    root = tk.Tk()
    app = OptimizadorPro(root)
    root.mainloop()
