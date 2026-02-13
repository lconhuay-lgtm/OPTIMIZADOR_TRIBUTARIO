import numpy as np


class SimFactoresMixin:
    def _generar_factores_simulacion(self, n_satelites, variabilidad_sat,
                                      var_matriz_costos, pct_var_costos,
                                      var_matriz_ingresos, pct_var_ingresos,
                                      distribucion, rho):
        """Genera factores aleatorios correlacionados para una iteracion de simulacion.

        Retorna: (factor_matriz_costos, factor_matriz_ingresos, lista_factores_satelites)
        - Correlacion rho se aplica entre Z_matriz (costos) y Z de cada satelite.
        - Ingresos de matriz se generan independientemente.
        - Si var_matriz_costos=False, correlacion se ignora.
        """
        def _factor_desde_z(z, cv, dist):
            """Convierte un valor Z~N(0,1) en un factor multiplicativo."""
            if cv <= 0:
                return 1.0
            if dist == "Log-normal":
                sigma_ln = np.sqrt(np.log(1 + cv ** 2))
                mu_ln = -0.5 * sigma_ln ** 2
                return float(np.exp(mu_ln + sigma_ln * z))
            else:  # Normal
                return max(0.0, 1.0 + cv * z)

        cv_sat = variabilidad_sat / 3.0  # CV efectivo de satelites
        cv_costos = pct_var_costos / 3.0
        cv_ingresos = pct_var_ingresos / 3.0

        # Factor costos matriz
        Z_matriz = None
        if var_matriz_costos:
            Z_matriz = np.random.randn()
            factor_costos = _factor_desde_z(Z_matriz, cv_costos, distribucion)
        else:
            factor_costos = 1.0

        # Factor ingresos matriz (independiente)
        if var_matriz_ingresos:
            Z_ing = np.random.randn()
            factor_ingresos = _factor_desde_z(Z_ing, cv_ingresos, distribucion)
        else:
            factor_ingresos = 1.0

        # Factores por satelite (correlacionados con Z_matriz si aplica)
        factores_sat = []
        for _ in range(n_satelites):
            if var_matriz_costos and Z_matriz is not None and rho != 0:
                eps = np.random.randn()
                Z_i = rho * Z_matriz + np.sqrt(1 - rho ** 2) * eps
            else:
                Z_i = np.random.randn()
            factor_sat = _factor_desde_z(Z_i, cv_sat, distribucion)
            factores_sat.append(factor_sat)

        return factor_costos, factor_ingresos, factores_sat

