# -*- coding: utf-8 -*-
"""
scripts/verify_models_stress_test.py — Empirical Challenger Verification Suite
Mathematical and Engineering Model Stress-Testing for TP N°1 (UTN FRSR)

Tests:
1. Criticality Matrix Recalculation Engine:
   - Exhaustive combinatorial generation of all 324 discrete (F, S, A, O, D) tuples.
   - Verification of min C (4, Class C) and max C (80, Class A).
   - Boundary condition testing (C=9 -> C, C=10 -> B, C=14 -> B, C=15 -> A).
   - Prevention of non-discrete values in web UI (index.html) and Python (app.py).
2. Standby Motor 55 kW Financial and Reliability Model:
   - Baseline scenario: DT=120h, Cost=$313,620 USD, MTBF=1,416.00h, MTTR=24.00h, Availability=98.3333%.
   - Standby scenario: F2 DT=4h (-38h), DT=82h, Direct Savings=$95,000.00 USD, MTTR=16.40h,
     Availability=98.8611%, Holding Cost=$400 USD, Net Annual Savings=$94,600 USD, Investment=$4,000 USD,
     Payback=15.37 days, ROI in [2,275%, 2,365%].
   - Verification that both index.html and app.py calculate these exact figures.
"""

import sys
import re
import math
import itertools
from pathlib import Path
import unittest

# Ensure UTF-8 output on Windows consoles
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

ROOT_DIR = Path(__file__).resolve().parent.parent
APP_PY_PATH = ROOT_DIR / "app.py"
INDEX_HTML_PATH = ROOT_DIR / "index.html"
REPORT_BUILDER_PATH = ROOT_DIR / "scripts" / "report_builder.py"


def clasificar_python(c, umbral_a=15, umbral_b=10):
    """Reference classification function matching app.py."""
    if c >= umbral_a:
        return ("Clase A - CRÍTICO",
                "Alta Disponibilidad: monitoreo continuo + overhaul programado")
    if c >= umbral_b:
        return ("Clase B - SEMICRÍTICO",
                "Sistemático / Predictivo: re-lubricación y mediciones por calendario/condición")
    return ("Clase C - NO CRÍTICO",
            "Condicional / TPM: inspección CIL por operador y atención a la falla")


def clasificar_js(c):
    """Reference classification function matching index.html."""
    if c >= 15:
        return ("Clase A — CRÍTICO", "Alta Disponibilidad (CBM + Overhaul)")
    elif c >= 10:
        return ("Clase B — SEMICRÍTICO", "Sistemático / Predictivo")
    else:
        return ("Clase C — NO CRÍTICO", "Modelo Condicional / TPM")


class TestCriticalityMatrixModel(unittest.TestCase):
    """Stress-testing of the Criticality Matrix model."""

    def setUp(self):
        self.F_discrete = [1, 2, 3, 4]
        self.S_discrete = [1, 2, 3]
        self.A_discrete = [1, 2, 3]
        self.O_discrete = [1, 3, 10]
        self.D_discrete = [1, 2, 4]

    def test_01_exhaustive_discrete_combinations(self):
        """Stress-test all 324 discrete combinations of (F, S, A, O, D)."""
        combinations = list(itertools.product(
            self.F_discrete,
            self.S_discrete,
            self.A_discrete,
            self.O_discrete,
            self.D_discrete
        ))
        self.assertEqual(len(combinations), 324, "Debe haber exactamente 324 combinaciones discretas (4*3*3*3*3)")

        c_values = []
        classes_py = {"Clase A - CRÍTICO": 0, "Clase B - SEMICRÍTICO": 0, "Clase C - NO CRÍTICO": 0}
        classes_js = {"Clase A — CRÍTICO": 0, "Clase B — SEMICRÍTICO": 0, "Clase C — NO CRÍTICO": 0}

        for f, s, a, o, d in combinations:
            ci = s + a + o + d
            c = f * ci
            c_values.append(c)

            # Python classification
            clase_py, _ = clasificar_python(c)
            classes_py[clase_py] += 1

            # JS classification
            clase_js, _ = clasificar_js(c)
            classes_js[clase_js] += 1

            # Consistency between Python and JS class
            if c >= 15:
                self.assertIn("Clase A", clase_py)
                self.assertIn("Clase A", clase_js)
            elif c >= 10:
                self.assertIn("Clase B", clase_py)
                self.assertIn("Clase B", clase_js)
            else:
                self.assertIn("Clase C", clase_py)
                self.assertIn("Clase C", clase_js)

        # Min and Max C
        min_c = min(c_values)
        max_c = max(c_values)

        self.assertEqual(min_c, 4, "El valor mínimo de C debe ser 4 (F=1, S=1, A=1, O=1, D=1)")
        self.assertEqual(max_c, 80, "El valor máximo de C debe ser 80 (F=4, S=3, A=3, O=10, D=4)")

        # Verify class counts
        self.assertEqual(sum(classes_py.values()), 324)
        self.assertEqual(sum(classes_js.values()), 324)

        print(f"\n[OK] Combinatoria exhaustiva: 324 combinaciones analizadas.")
        print(f"     Min C = {min_c} (Clase C), Max C = {max_c} (Clase A).")
        print(f"     Distribución: Clase A = {classes_py['Clase A - CRÍTICO']} ({classes_py['Clase A - CRÍTICO']/324*100:.1f}%), "
              f"Clase B = {classes_py['Clase B - SEMICRÍTICO']} ({classes_py['Clase B - SEMICRÍTICO']/324*100:.1f}%), "
              f"Clase C = {classes_py['Clase C - NO CRÍTICO']} ({classes_py['Clase C - NO CRÍTICO']/324*100:.1f}%).")

    def test_02_boundary_classifications(self):
        """Stress-test exact boundary conditions (C=9, C=10, C=14, C=15) and floating-point perturbations."""
        # Exact integer boundaries
        self.assertEqual(clasificar_python(9)[0], "Clase C - NO CRÍTICO")
        self.assertEqual(clasificar_python(10)[0], "Clase B - SEMICRÍTICO")
        self.assertEqual(clasificar_python(14)[0], "Clase B - SEMICRÍTICO")
        self.assertEqual(clasificar_python(15)[0], "Clase A - CRÍTICO")

        self.assertEqual(clasificar_js(9)[0], "Clase C — NO CRÍTICO")
        self.assertEqual(clasificar_js(10)[0], "Clase B — SEMICRÍTICO")
        self.assertEqual(clasificar_js(14)[0], "Clase B — SEMICRÍTICO")
        self.assertEqual(clasificar_js(15)[0], "Clase A — CRÍTICO")

        # Discrete tuples that yield exact boundaries
        combos_c10 = []
        combos_c14 = []
        combos_c15 = []

        for f, s, a, o, d in itertools.product(
            self.F_discrete, self.S_discrete, self.A_discrete, self.O_discrete, self.D_discrete
        ):
            c = f * (s + a + o + d)
            if c == 10:
                combos_c10.append((f, s, a, o, d))
            elif c == 14:
                combos_c14.append((f, s, a, o, d))
            elif c == 15:
                combos_c15.append((f, s, a, o, d))

        self.assertGreater(len(combos_c10), 0, "Deben existir combinaciones que den exactamente C=10")
        self.assertGreater(len(combos_c14), 0, "Deben existir combinaciones que den exactamente C=14")
        self.assertGreater(len(combos_c15), 0, "Deben existir combinaciones que den exactamente C=15")

        for combo in combos_c10:
            c = combo[0] * sum(combo[1:])
            self.assertEqual(c, 10)
            self.assertIn("Clase B", clasificar_python(c)[0])

        for combo in combos_c14:
            c = combo[0] * sum(combo[1:])
            self.assertEqual(c, 14)
            self.assertIn("Clase B", clasificar_python(c)[0])

        for combo in combos_c15:
            c = combo[0] * sum(combo[1:])
            self.assertEqual(c, 15)
            self.assertIn("Clase A", clasificar_python(c)[0])

        # Floating-point perturbations
        eps = 1e-6
        self.assertEqual(clasificar_python(10.0 - eps)[0], "Clase C - NO CRÍTICO")
        self.assertEqual(clasificar_python(10.0)[0], "Clase B - SEMICRÍTICO")
        self.assertEqual(clasificar_python(15.0 - eps)[0], "Clase B - SEMICRÍTICO")
        self.assertEqual(clasificar_python(15.0)[0], "Clase A - CRÍTICO")

        print(f"\n[OK] Condiciones de contorno C=9 (Clase C), C=10 (Clase B), C=14 (Clase B), C=15 (Clase A) verificadas.")
        print(f"     Tuplas con C=10: {len(combos_c10)} | Tuplas con C=14: {len(combos_c14)} | Tuplas con C=15: {len(combos_c15)}.")

    def test_03_non_discrete_prevention_in_codebases(self):
        """Verify that non-discrete inputs are strictly prevented in both index.html and app.py."""
        # 1. Inspect index.html
        self.assertTrue(INDEX_HTML_PATH.exists(), "index.html debe existir")
        with open(INDEX_HTML_PATH, "r", encoding="utf-8") as f:
            html_content = f.read()

        # Check that selects are used for F, S, A, O, D in renderCriticidadTable
        self.assertIn('actualizarValorCriticidad(${idx}, \'f\', this.value)', html_content)
        self.assertIn('actualizarValorCriticidad(${idx}, \'s\', this.value)', html_content)
        self.assertIn('actualizarValorCriticidad(${idx}, \'a\', this.value)', html_content)
        self.assertIn('actualizarValorCriticidad(${idx}, \'o\', this.value)', html_content)
        self.assertIn('actualizarValorCriticidad(${idx}, \'d\', this.value)', html_content)

        # Verify options in HTML
        # O must only have 1, 3, 10
        o_select_match = re.search(r'actualizarValorCriticidad\([^)]+\'o\'[^<]+<select[^>]*>(.*?)</select>', html_content, re.DOTALL)
        if not o_select_match:
            # Alternate search
            o_select_match = re.search(r'<!-- SELECTOR O.*?<select[^>]*>(.*?)</select>', html_content, re.DOTALL)
        self.assertIsNotNone(o_select_match, "Selector de O debe encontrarse en index.html")
        o_options = re.findall(r'value="(\d+)"', o_select_match.group(1))
        self.assertEqual(sorted([int(x) for x in o_options]), [1, 3, 10], "O en index.html debe ser estrictamente [1, 3, 10]")

        # D must only have 1, 2, 4
        d_select_match = re.search(r'<!-- SELECTOR D.*?<select[^>]*>(.*?)</select>', html_content, re.DOTALL)
        self.assertIsNotNone(d_select_match, "Selector de D debe encontrarse en index.html")
        d_options = re.findall(r'value="(\d+)"', d_select_match.group(1))
        self.assertEqual(sorted([int(x) for x in d_options]), [1, 2, 4], "D en index.html debe ser estrictamente [1, 2, 4]")

        # F must only have 1, 2, 3, 4
        f_select_match = re.search(r'<!-- SELECTOR F.*?<select[^>]*>(.*?)</select>', html_content, re.DOTALL)
        self.assertIsNotNone(f_select_match, "Selector de F debe encontrarse en index.html")
        f_options = re.findall(r'value="(\d+)"', f_select_match.group(1))
        self.assertEqual(sorted([int(x) for x in f_options]), [1, 2, 3, 4], "F en index.html debe ser estrictamente [1, 2, 3, 4]")

        # 2. Inspect app.py
        self.assertTrue(APP_PY_PATH.exists(), "app.py debe existir")
        with open(APP_PY_PATH, "r", encoding="utf-8") as f:
            py_content = f.read()

        # Check st.column_config.SelectboxColumn
        self.assertIn('"F": st.column_config.SelectboxColumn("F", options=[1, 2, 3, 4]', py_content)
        self.assertIn('"S": st.column_config.SelectboxColumn("S", options=[1, 2, 3]', py_content)
        self.assertIn('"A": st.column_config.SelectboxColumn("A", options=[1, 2, 3]', py_content)
        self.assertIn('"O": st.column_config.SelectboxColumn("O", options=[1, 3, 10]', py_content)
        self.assertIn('"D": st.column_config.SelectboxColumn("D", options=[1, 2, 4]', py_content)

        print("[OK] Prevención de valores no discretos: comprobada en index.html (<select>) y en app.py (SelectboxColumn).")


class TestStandbyMotorModel(unittest.TestCase):
    """Stress-testing of the Standby Motor 55 kW financial and reliability model."""

    def setUp(self):
        self.top = 7200.0
        self.costo_parada_usd = 2500.0
        self.costo_hh_usd = 25.0
        self.holding_cost = 400.0
        self.investment = 4000.0

        # Raw failure data
        self.raw_fallas = [
            {"evento": "F1", "ttp": 28.0, "repuestos": 1850.0, "hh": 48.0},
            {"evento": "F2", "ttp": 42.0, "repuestos": 3400.0, "hh": 24.0},
            {"evento": "F3", "ttp": 6.0,  "repuestos": 320.0,  "hh": 12.0},
            {"evento": "F4", "ttp": 36.0, "repuestos": 4200.0, "hh": 36.0},
            {"evento": "F5", "ttp": 8.0,  "repuestos": 450.0,  "hh": 16.0},
        ]

    def test_04_baseline_model(self):
        """Stress-test Baseline: downtime=120h, cost=$313,620, MTBF=1,416h, MTTR=24h, Availability=98.3333%."""
        total_ttp = sum(f["ttp"] for f in self.raw_fallas)
        self.assertAlmostEqual(total_ttp, 120.0, places=4, msg="Downtime base debe ser exactamente 120.0 h")

        total_repuestos = sum(f["repuestos"] for f in self.raw_fallas)
        total_hh = sum(f["hh"] for f in self.raw_fallas)
        costo_parada = total_ttp * self.costo_parada_usd
        costo_mo = total_hh * self.costo_hh_usd
        costo_total = costo_parada + total_repuestos + costo_mo

        self.assertAlmostEqual(costo_parada, 300000.0, places=2)
        self.assertAlmostEqual(total_repuestos, 10220.0, places=2)
        self.assertAlmostEqual(costo_mo, 3400.0, places=2)
        self.assertAlmostEqual(costo_total, 313620.0, places=2, msg="Costo total base debe ser $313,620 USD")

        n_fallas = len(self.raw_fallas)
        tfr = self.top - total_ttp
        self.assertAlmostEqual(tfr, 7080.0, places=2)

        mtbf = tfr / n_fallas
        mttr = total_ttp / n_fallas
        availability = (mtbf / (mtbf + mttr)) * 100

        self.assertAlmostEqual(mtbf, 1416.00, places=2, msg="MTBF base debe ser 1,416.00 h")
        self.assertAlmostEqual(mttr, 24.00, places=2, msg="MTTR base debe ser 24.00 h")
        self.assertAlmostEqual(availability, 98.333333, places=4, msg="Disponibilidad base debe ser 98.3333%")

        print(f"\n[OK] Modelo Caso Base verificado:")
        print(f"     TTP = {total_ttp:.1f} h, Costo Total = ${costo_total:,.2f} USD, MTBF = {mtbf:.2f} h, "
              f"MTTR = {mttr:.2f} h, Disponibilidad = {availability:.4f}%.")

    def test_05_standby_scenario_model(self):
        """Stress-test Standby: F2 DT=4h (-38h), DT=82h, direct savings=$95,000, MTTR=16.4h, Avail=98.8611%,
           holding=$400, net savings=$94,600, invest=$4,000, Payback=15.37 days, ROI in [2,275%, 2,365%]."""
        standby_fallas = [dict(f) for f in self.raw_fallas]
        for f in standby_fallas:
            if f["evento"] == "F2":
                f["ttp"] = 4.0

        f2_baseline_ttp = 42.0
        f2_standby_ttp = 4.0
        ahorro_horas_f2 = f2_baseline_ttp - f2_standby_ttp
        self.assertAlmostEqual(ahorro_horas_f2, 38.0, places=4, msg="Reducción de downtime F2 debe ser 38.0 h")

        direct_savings = ahorro_horas_f2 * self.costo_parada_usd
        self.assertAlmostEqual(direct_savings, 95000.0, places=2, msg="Ahorro directo de lucro cesante debe ser $95,000.00 USD")

        total_ttp_standby = sum(f["ttp"] for f in standby_fallas)
        self.assertAlmostEqual(total_ttp_standby, 82.0, places=4, msg="Downtime total con Standby debe ser 82.0 h")

        n_fallas = len(standby_fallas)
        tfr_standby = self.top - total_ttp_standby
        self.assertAlmostEqual(tfr_standby, 7118.0, places=2)

        mttr_standby = total_ttp_standby / n_fallas
        mtbf_standby = tfr_standby / n_fallas
        availability_standby = (tfr_standby / self.top) * 100

        self.assertAlmostEqual(mttr_standby, 16.40, places=2, msg="MTTR global con Standby debe ser 16.40 h")
        self.assertAlmostEqual(mtbf_standby, 1423.60, places=2, msg="MTBF global con Standby debe ser 1,423.60 h")
        self.assertAlmostEqual(availability_standby, 98.861111, places=4, msg="Disponibilidad con Standby debe ser 98.8611%")

        # Financial metrics
        net_annual_savings = direct_savings - self.holding_cost
        self.assertAlmostEqual(net_annual_savings, 94600.0, places=2, msg="Ahorro neto anual operacional debe ser $94,600 USD")

        first_year_net_benefit = direct_savings - self.investment - self.holding_cost
        self.assertAlmostEqual(first_year_net_benefit, 90600.0, places=2, msg="Beneficio neto Año 1 debe ser $90,600 USD")

        # Payback
        # Formula standard: Investment / (Annual Savings / 365 days)
        payback_days = self.investment / (direct_savings / 365.0)
        self.assertAlmostEqual(payback_days, 15.3684, places=2, msg="Payback debe ser aprox 15.37 días")
        payback_rounded = round(payback_days, 2)
        self.assertEqual(payback_rounded, 15.37)

        # Production hours equivalent for payback
        prod_hours_payback = self.investment / self.costo_parada_usd
        self.assertAlmostEqual(prod_hours_payback, 1.60, places=2, msg="Recupero en producción: 1.60 horas")

        # ROI definitions
        # 1. Capital ROI: (Direct Savings - Investment) / Investment * 100%
        roi_capital = ((direct_savings - self.investment) / self.investment) * 100
        self.assertAlmostEqual(roi_capital, 2275.0, places=2, msg="ROI sobre capital debe ser 2,275%")

        # 2. Operating ROI: Net Annual Savings / Investment * 100%
        roi_operating = (net_annual_savings / self.investment) * 100
        self.assertAlmostEqual(roi_operating, 2365.0, places=2, msg="ROI operacional debe ser 2,365%")

        # Verify ROI is within the requested range [2,275%, 2,365%]
        self.assertTrue(2275.0 <= roi_capital <= 2365.0)
        self.assertTrue(2275.0 <= roi_operating <= 2365.0)

        print(f"\n[OK] Modelo Escenario Standby verificado:")
        print(f"     F2 DT = {f2_standby_ttp:.1f} h (-{ahorro_horas_f2:.1f} h), TTP Total = {total_ttp_standby:.1f} h.")
        print(f"     Ahorro Directo = ${direct_savings:,.2f} USD, MTTR = {mttr_standby:.2f} h, Disp = {availability_standby:.4f}%.")
        print(f"     Costo Posesión = ${self.holding_cost:,.2f} USD, Ahorro Neto Anual = ${net_annual_savings:,.2f} USD.")
        print(f"     Inversión = ${self.investment:,.2f} USD, Beneficio Neto Año 1 = ${first_year_net_benefit:,.2f} USD.")
        print(f"     Payback = {payback_days:.2f} días ({prod_hours_payback:.2f} h de producción).")
        print(f"     ROI Capital = {roi_capital:.0f}%, ROI Operacional = {roi_operating:.0f}% (Rango: [2,275%, 2,365%]).")

    def test_06_code_implementation_consistency(self):
        """Verify that both index.html and app.py calculate and display these exact figures."""
        with open(INDEX_HTML_PATH, "r", encoding="utf-8") as f:
            html = f.read()

        with open(APP_PY_PATH, "r", encoding="utf-8") as f:
            py = f.read()

        # index.html checks
        self.assertIn('$95,000 USD', html, "index.html debe contener el ahorro de $95,000 USD")
        self.assertIn('2,275%', html, "index.html debe contener el ROI de 2,275%")
        self.assertIn('15.3', html, "index.html debe contener el Payback de 15.3 días")
        self.assertIn('f.ttp = 4.0', html, "index.html debe asignar F2.ttp = 4.0 al activar Standby")
        self.assertIn('STANDBY (82h)', html, "index.html debe reflejar 82h en tag Standby")
        self.assertIn('16.40', html, "index.html debe contener el MTTR de 16.40 h")
        self.assertIn('98.86', html, "index.html debe contener la disponibilidad de 98.86%")
        self.assertIn('ai.toFixed(2)', html, "index.html debe formatear la disponibilidad con 2 decimales")

        # app.py checks
        self.assertIn('ahorro_horas_f2 = 42.0 - mttr_standby', py)
        self.assertIn('ahorro_downtime_f2 = ahorro_horas_f2 * costo_par', py)
        self.assertIn('nuevo_ttp_global = 120.0 - ahorro_horas_f2', py)
        self.assertIn('nuevo_mttr_global = nuevo_ttp_global / 5.0', py)
        self.assertIn('nueva_disp_global = ((TOP_anio - nuevo_ttp_global) / TOP_anio) * 100', py)
        self.assertIn('roi_standby = ((ahorro_downtime_f2 - inv_motor) / inv_motor) * 100', py)
        self.assertIn('payback_dias = (inv_motor / (ahorro_downtime_f2 / 365.0))', py)
        self.assertIn('base.loc[base["evento"] == "F2", "ttp"] = 4.0', py)

        # report_builder.py checks
        if REPORT_BUILDER_PATH.exists():
            with open(REPORT_BUILDER_PATH, "r", encoding="utf-8") as f:
                rb = f.read()
            self.assertIn('$95,000 USD', rb)
            self.assertIn('16.40', rb)
            self.assertIn('2,275%', rb)
            self.assertIn('15.3', rb)
            self.assertIn('90,600', rb)

        print("[OK] Consistencia de código: verificada en index.html, app.py y report_builder.py.")


if __name__ == "__main__":
    print("=" * 80)
    print(" EMPIRICAL CHALLENGER: STRESS-TESTING MATHEMATICAL & ENGINEERING MODELS")
    print("=" * 80)
    unittest.main(verbosity=2)
