"""Compatibilidad del desglose de lavado de concreto de DARCA II.

El mapeo contractual conserva su identidad para verificar históricos existentes.
Los dos WBS nuevos son desgloses del mismo ítem, no ítems ni cantidades nuevas.
No se modifican los reportes ni sus huellas.
"""
from .modelos import D, ErrorDatos, Hallazgo, normalizar


def desglose_lavado(reportes, presupuesto, mapping, hallazgos):
    codigos = {"1.1.10", "1.2.10", "1.3.1"}
    nuevos = codigos - {"1.3.1"}
    if not any(r.wbs in nuevos for reporte in reportes for r in reporte.registros):
        return mapping, {}, set()
    item = mapping.get("1.3.1")
    if (item != 11 or presupuesto.items[item].cantidad != D(192)
            or any(w in mapping and mapping[w] != item for w in nuevos)):
        raise ErrorDatos("El desglose de lavado de concreto requiere el ítem 11 de 192 m² y WBS compatibles.")
    effective = {**mapping, **{w: item for w in nuevos}}
    legacy = set()
    for reporte in reportes:
        rows = {r.wbs: r for r in reporte.registros if r.wbs in codigos}
        if set(rows) == {"1.3.1"} and rows["1.3.1"].total == D(192):
            # No atribuir un avance agregado a un tramo sin evidencia.
            if rows["1.3.1"].ejecutada != 0:
                raise ErrorDatos(f"{reporte.archivo}: el lavado figura agregado en 192 m² con avance. "
                                 "Distribuya ese avance por tramo antes de combinarlo con reportes desglosados.")
            legacy.add(reporte.archivo)
        elif set(rows) == codigos and all(r.total == D(64) for r in rows.values()):
            for w, r in rows.items():
                if normalizar(r.tramo) != f"tramo {w.split('.')[1]}":
                    raise ErrorDatos(f"{reporte.archivo}: tramo incompatible para {w}.")
                if r.actividad and "lavado de concreto" not in normalizar(r.actividad):
                    raise ErrorDatos(f"{reporte.archivo}: actividad incompatible para {w}.")
        else:
            raise ErrorDatos(f"{reporte.archivo}: desglose de lavado incompleto o cantidades incompatibles. "
                             "Use los tres WBS 1.1.10, 1.2.10 y 1.3.1 con 64 m² cada uno.")
    hallazgos.append(Hallazgo("INFO", "LAVADO_TRES_TRAMOS",
        "Ítem 11: se suman 1.1.10, 1.2.10 y 1.3.1; contrato total 192 m² (64 por tramo). "
        "Reportes antiguos agregados sin avance se conservan sin duplicar cantidades.", item=item))
    return effective, {w: D(64) for w in codigos}, legacy
