#!/usr/bin/env python3
r"""CORTES DE OBRA — DEMOSTRACIÓN PÚBLICA

1. Instale Python 3.11+ y las dependencias: python -m pip install -r requirements.txt
2. Prepare datos ficticios: python scripts/crear_demo.py
3. Ejecute python generar_corte.py o INICIAR_CORTE.bat en Windows.
4. Para la demo: inicio 05/01/2026, fin 17/01/2026, corte 1.
5. Ingrese porcentajes como 10 para 10%. Revise antes de registrar definitivamente.
6. El Excel aparece en output/. El histórico confirmado aparece en data/historico/.

El conjunto de demostración no contiene datos de clientes. Consulte README.md
para arquitectura, controles, alcance, ejemplos y limitaciones.
"""
import argparse
from datetime import datetime
import json
from pathlib import Path
import sys

from darca.calculo import preparar
from darca.excel import exportar
from darca.historico import registrar
from darca.modelos import ErrorDatos, Parametros, fecha, numero, porcentaje


def preguntar(label, parser, default=None):
    while True:
        raw = input(f"{label}" + (f" [{default}]" if default is not None else "") + ": ").strip()
        if not raw and default is not None:
            raw = str(default)
        try:
            return parser(raw)
        except (ValueError, ErrorDatos) as exc:
            print(f"Entrada inválida: {exc}")


def corte_numero(value):
    n = int(value)
    if n < 1:
        raise ValueError("Debe ser un entero positivo.")
    return n


def descuento(value):
    n = numero(value)
    if n < 0:
        raise ValueError("No puede ser negativo.")
    return n


def pesos(value):
    return "$" + f"{value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def main(argv=None):
    parser = argparse.ArgumentParser(description="Genera cortes de obra con validación y registro opcional.")
    parser.add_argument("--manual", action="store_true", help="Muestra el manual incluido en el código y termina.")
    parser.add_argument("--proyecto", type=Path, default=Path(__file__).resolve().parent,
                        help="Carpeta de proyecto con config/ y data/ (por defecto junto al script).")
    parser.add_argument("--numero", type=corte_numero)
    parser.add_argument("--inicio", type=fecha)
    parser.add_argument("--fin", type=fecha)
    parser.add_argument("--amortizacion", type=porcentaje)
    parser.add_argument("--retefuente", type=porcentaje)
    parser.add_argument("--garantia", type=porcentaje)
    parser.add_argument("--otros", type=descuento)
    parser.add_argument("--preliminar", action="store_true", help="No solicita ni permite registrar el histórico.")
    parser.add_argument("--confirmar-generacion", action="store_true",
                        help="Omite confirmación de generación únicamente con --preliminar.")
    args = parser.parse_args(argv)
    if args.manual:
        print(__doc__)
        return 0
    root = args.proyecto.resolve()
    try:
        if args.confirmar_generacion and not args.preliminar:
            raise ErrorDatos("--confirmar-generacion requiere --preliminar; nunca confirma el histórico.")
        print("Cortes de obra · introduzca 40 para 40%; números sin separadores de miles.")
        print("Manual: INICIAR_CORTE.bat -Manual (Windows) o python generar_corte.py --manual")

        print("Defina los datos del corte antes de exportar el Excel.")
        inicio = args.inicio if args.inicio is not None else preguntar("Fecha inicial del periodo (DD/MM/AAAA)", fecha)
        fin = args.fin if args.fin is not None else preguntar("Fecha final del periodo (DD/MM/AAAA)", fecha)
        numero_corte = args.numero if args.numero is not None else preguntar("Número de corte", corte_numero)
        anticipo = args.amortizacion if args.amortizacion is not None else preguntar("Porcentaje de anticipo a amortizar (%)", porcentaje)
        retefuente = args.retefuente if args.retefuente is not None else preguntar("Porcentaje de retención en la fuente (%)", porcentaje)
        garantia = args.garantia if args.garantia is not None else preguntar("Porcentaje de retención en garantía (%)", porcentaje)
        otros = args.otros if args.otros is not None else preguntar("Otros descuentos (pesos)", descuento, "0")
        params = Parametros(numero_corte, inicio, fin, anticipo, retefuente, garantia, otros)
        result = preparar(root, params)
        print(f"\n{result.config['proyecto']} · Corte {params.numero:02d} · {params.inicio:%d/%m/%Y} al {params.fin:%d/%m/%Y}")
        for label, nombre in (("amortizacion", "Anticipo a amortizar"), ("retefuente", "Retención en la fuente"), ("garantia", "Retención en garantía")):
            print(f"{nombre}: {getattr(params,label)*100:g}% — base: {'bruto menos IVA' if label == 'retefuente' else result.config['bases_deducciones'][label]}")
        print(f"Otros descuentos: {pesos(params.otros)}")
        print("AIU / IVA:", ", ".join(f"{k}: {v*100:g}%" for k, v in result.presupuesto.tasas.items()))
        for key in ("directos", "bruto", "neto"):
            print(f"{key.capitalize()}: {pesos(result.finanzas[key])}")
        print("Archivos incluidos:")
        for report in result.reportes:
            print(f"  {report.fecha}: {report.archivo}")
        for finding in result.hallazgos:
            print(f"[{finding.nivel}] {finding.codigo}: {finding.archivo} {finding.mensaje}")
        if result.criticos:
            print("Hay errores críticos: solo se puede exportar un borrador marcado como no apto para cobro.")
        if not args.confirmar_generacion and input("\n¿Generar Excel preliminar? [s/N]: ").strip().casefold() not in ("s", "si", "sí"):
            print("Generación cancelada. Histórico sin cambios.")
            return 0
        output = exportar(result, root / "output")
        print(f"\nExcel preliminar verificado: {output}")
        if args.preliminar or result.criticos:
            print("No se ha registrado el corte en el histórico.")
            return 2 if result.criticos else 0
        phrase = f"REGISTRAR {params.numero}"
        reply = input(f"\nRevise el Excel. Para registrar definitivamente escriba {phrase!r}; Enter conserva solo el borrador: ")
        if reply.strip() == phrase:
            definitive, control = registrar(root, result, exportar)
            print(f"Corte definitivo: {definitive}\nControl: {control}")
        else:
            print("Vista preliminar conservada. Histórico sin cambios.")
        return 0
    except (ErrorDatos, OSError) as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        print("No se registró ningún corte definitivo.", file=sys.stderr)
        try:
            folder = root / "output"
            folder.mkdir(parents=True, exist_ok=True)
            log = folder / f"ERROR_VALIDACION_{datetime.now():%Y%m%d_%H%M%S_%f}.json"
            log.write_text(json.dumps({"nivel": "ERROR", "mensaje": str(exc)}, indent=2, ensure_ascii=False), encoding="utf-8")
            print(f"Detalle de validación: {log}", file=sys.stderr)
        except OSError:
            pass
        return 2
    except (EOFError, KeyboardInterrupt):
        print("\nOperación cancelada; no se confirmó el histórico.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
