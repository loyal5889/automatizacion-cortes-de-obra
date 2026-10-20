import unittest
from types import SimpleNamespace as NS
from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from darca.tramos import desglose_lavado
from darca.modelos import D, ErrorDatos


def row(wbs, total, avance=0):
    return NS(wbs=wbs, total=D(total), ejecutada=D(avance),
              tramo=f"TRAMO {wbs.split('.')[1]}",
              actividad="Lavado de concreto en fachada flotante")


class TramosTest(unittest.TestCase):
    def setUp(self):
        self.budget = NS(items={11: NS(cantidad=D(192))})
        self.mapping = {"1.3.1": 11}
        self.old = NS(archivo="anterior.xlsx", registros=[row("1.3.1", 192)])
        self.split = NS(archivo="nuevo.xlsx", registros=[
            row("1.1.10", 64, 32), row("1.2.10", 64, 32), row("1.3.1", 64, 32)])

    def test_tres_tramos_suman_96_sin_cambiar_fuentes_o_mapeo_contractual(self):
        reports = [self.old, self.split, deepcopy(self.old)]
        findings = []
        mapping, totals, legacy = desglose_lavado(reports, self.budget, self.mapping, findings)
        self.assertEqual(sum(r.ejecutada for r in self.split.registros if mapping[r.wbs] == 11), D(96))
        self.assertEqual(sum(totals.values()), D(192))
        self.assertEqual(self.old.registros[0].total, D(192))
        self.assertEqual(self.mapping, {"1.3.1": 11})
        self.assertIn(self.old.archivo, legacy)

    def test_reportes_antiguos_solos_no_cambian(self):
        mapping, totals, legacy = desglose_lavado([self.old], self.budget, self.mapping, [])
        self.assertEqual(mapping, self.mapping)
        self.assertFalse(totals)
        self.assertFalse(legacy)

    def test_desglose_incompleto_se_rechaza(self):
        self.split.registros.pop()
        with self.assertRaises(ErrorDatos):
            desglose_lavado([self.old, self.split], self.budget, self.mapping, [])

    def test_total_contratado_no_se_triplica(self):
        self.split.registros[2].total = D(192)
        with self.assertRaises(ErrorDatos):
            desglose_lavado([self.old, self.split], self.budget, self.mapping, [])

    def test_avance_agregado_no_se_atribuye_arbitrariamente_al_tramo_tres(self):
        self.old.registros[0].ejecutada = D(20)
        with self.assertRaises(ErrorDatos):
            desglose_lavado([self.old, self.split], self.budget, self.mapping, [])

    def test_no_reasigna_wbs_de_otro_item(self):
        self.mapping["1.1.10"] = 10
        with self.assertRaises(ErrorDatos):
            desglose_lavado([self.split], self.budget, self.mapping, [])
