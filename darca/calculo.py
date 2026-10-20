from dataclasses import asdict
from datetime import date

from .historico import leer_historico
from .lectura import leer_config, leer_presupuesto, leer_reporte
from .tramos import desglose_lavado
from .modelos import (D, ErrorDatos, Hallazgo, Resultado, fecha, firma, hash_archivo,
                      normalizar, numero)


def calcular_finanzas(directos, tasas, parametros, bases):
    f = {"directos": directos}
    for name in ("administracion", "imprevistos", "utilidad"):
        f[name] = directos * tasas[name]
    f["iva"] = f["utilidad"] * tasas["iva"]
    f["bruto"] = sum(f.values(), D(0))
    for name in ("amortizacion", "retefuente", "garantia"):
        base = f["bruto"] - f["iva"] if name == "retefuente" else f[bases[name]]
        f[name] = base * getattr(parametros, name)
    f["otros"] = parametros.otros
    f["neto"] = f["bruto"] - sum((f[k] for k in ("amortizacion", "retefuente", "garantia", "otros")), D(0))
    return f


def preparar(root, parametros):
    parametros.validar()
    config, mapping = leer_config(root)
    findings = []
    budget_files = [p for p in (root / "data/presupuesto").glob("*.xlsx") if not p.name.startswith("~$")]
    if len(budget_files) != 1:
        raise ErrorDatos("Coloque exactamente un archivo .xlsx de presupuesto en data/presupuesto/.")
    budget = leer_presupuesto(budget_files[0], config, findings)
    if set(mapping.values()) != set(budget.items):
        raise ErrorDatos("El mapeo debe cubrir todos los ítems contractuales y no asignar ítems inexistentes.")
    context = firma({"presupuesto": budget.huella, "mapeo": mapping, "config": config})
    history, snapshot = leer_historico(root / "data/historico")
    earlier = [r for r in history if r["numero"] < parametros.numero]
    prev = earlier[-1] if earlier else None
    if any(r["numero"] == parametros.numero for r in history):
        findings.append(Hallazgo("ERROR", "CORTE_EXISTENTE", "El número del corte ya existe: solo se permite vista preliminar; no se sobrescribe."))
    elif parametros.numero != len(history) + 1:
        findings.append(Hallazgo("ERROR", "SECUENCIA", f"El siguiente corte registrable es {len(history) + 1}."))
    if prev:
        if prev["contexto"] != context:
            raise ErrorDatos("El contrato, configuración o mapeo cambió respecto al histórico. Se requiere una migración revisada.")
        if set(prev["cantidades_acumuladas"]) != {str(i) for i in budget.items}:
            raise ErrorDatos("Los ítems del histórico no coinciden con el presupuesto.")
        if parametros.inicio <= fecha(prev["fin"]):
            findings.append(Hallazgo("ERROR", "PERIODO_SOLAPADO", "El periodo debe iniciar después del último corte confirmado."))

    files = sorted(p for p in (root / "data/reportes").glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise ErrorDatos("No hay reportes .xlsx en data/reportes/.")
    reports, hashes, occupied = [], {}, {}
    source_hashes = {p.relative_to(root).as_posix(): hash_archivo(p) for p in
                     [budget_files[0], root / "config/proyecto.json", root / "config/mapeo_wbs.json"]}
    for path in files:
        try:
            report = leer_reporte(path, config, findings)
        except ErrorDatos as exc:
            # Una lectura incompleta no puede producir importes aparentemente válidos.
            raise ErrorDatos(str(exc)) from exc
        source_hashes[path.relative_to(root).as_posix()] = report.sha256
        if report.fecha > parametros.fin:
            findings.append(Hallazgo("INFO", "REPORTE_POSTERIOR", f"Reporte del {report.fecha}: excluido por fecha final.", report.archivo))
            continue
        if report.huella in hashes:
            findings.append(Hallazgo("ERROR", "REPORTE_DUPLICADO", f"Mismos datos que {hashes[report.huella]}; se excluye la copia.", report.archivo))
            continue
        hashes[report.huella] = report.archivo
        for r in report.registros:
            key = (r.fecha, r.wbs)
            if key in occupied:
                findings.append(Hallazgo("ERROR", "SEMANA_DUPLICADA", f"WBS {r.wbs} y fecha {r.fecha} ya presentes en {occupied[key]}.", report.archivo, r.fila))
            occupied[key] = report.archivo
        reports.append(report)
    reports.sort(key=lambda r: (r.fecha, r.archivo))
    if not reports:
        raise ErrorDatos("No hay reportes incluidos hasta la fecha final indicada.")
    prior_reports = {r["huella"]: r for r in prev["reportes"]} if prev else {}
    included = {r.huella for r in reports}
    for digest, r in prior_reports.items():
        if digest not in included:
            findings.append(Hallazgo("ERROR", "REPORTE_HISTORICO_FALTANTE", f"Falta o fue modificado {r['archivo']}. Conserve todos los reportes ya cobrados."))
    for report in reports:
        if report.huella in prior_reports:
            findings.append(Hallazgo("INFO", "YA_INCLUIDO", "Incluido en un corte anterior: se conserva para el acumulado y se resta lo ya cobrado.", report.archivo))
        elif report.fecha < parametros.inicio:
            if not prev and parametros.numero > 1:
                findings.append(Hallazgo("INFO", "BASE_ANTERIOR", "Reporte anterior al periodo: solo integra la cantidad anterior; no se cobra en este corte.", report.archivo))
            else:
                findings.append(Hallazgo("ERROR", "REPORTE_FUERA_PERIODO", "Reporte nuevo anterior al inicio del periodo. Amplíe el periodo o revise la fecha.", report.archivo))
        if report.fecha > date.today():
            findings.append(Hallazgo("ADVERTENCIA", "FECHA_FUTURA", f"El reporte tiene fecha futura: {report.fecha}.", report.archivo))

    mapping, tramos_totales, reportes_agregados = desglose_lavado(reports, budget, mapping, findings)
    totals = {i: D(0) for i in budget.items}
    before_period = {i: D(0) for i in budget.items}
    grouped = {}
    for report in reports:
        for r in report.registros:
            total_tramo = tramos_totales.get(r.wbs, r.total)
            assigned = mapping.get(r.wbs)
            if assigned is None:
                findings.append(Hallazgo("ERROR", "WBS_SIN_MAPEO", f"WBS {r.wbs} sin ítem contractual.", r.archivo, r.fila))
            else:
                if r.unidad != budget.items[assigned].unidad:
                    findings.append(Hallazgo("ERROR", "UNIDAD", f"WBS {r.wbs}: unidad {r.unidad}; presupuesto {budget.items[assigned].unidad}.", r.archivo, r.fila, assigned))
                totals[assigned] += r.ejecutada
                if report.fecha < parametros.inicio:
                    before_period[assigned] += r.ejecutada
            if r.wbs not in grouped:
                grouped[r.wbs] = {"wbs": r.wbs, "actividad": r.actividad, "tramo": r.tramo,
                                  "unidad": r.unidad, "total": total_tramo, "item": assigned,
                                  "semanas": {}, "acumulada": D(0)}
            g = grouped[r.wbs]
            if g["total"] != total_tramo or g["unidad"] != r.unidad or normalizar(g["tramo"]) != normalizar(r.tramo):
                findings.append(Hallazgo("ERROR", "WBS_INCONSISTENTE", f"WBS {r.wbs}: cambió total, unidad o tramo entre reportes.", r.archivo, r.fila, assigned))
            if r.actividad:
                g["actividad"] = r.actividad
            g["semanas"][r.archivo] = r.ejecutada
            g["acumulada"] += r.ejecutada
    for wbs, g in grouped.items():
        if not g["actividad"] and g["item"]:
            g["actividad"] = budget.items[g["item"]].concepto
        if g["acumulada"] > g["total"]:
            findings.append(Hallazgo("ERROR", "EXCESO_TRAMO", f"WBS {wbs}: acumulado {g['acumulada']} superior al total {g['total']}.", item=g["item"]))
    for report in reports:
        missing = set(mapping) - {r.wbs for r in report.registros}
        if report.archivo in reportes_agregados:
            missing -= {"1.1.10", "1.2.10"}
        if missing:
            findings.append(Hallazgo("ERROR", "WBS_FALTANTES", f"Faltan WBS del mapeo: {', '.join(sorted(missing))}. Declare cero explícito si no hubo avance.", report.archivo))
    for i, item in budget.items.items():
        distributed = sum((g["total"] for g in grouped.values() if g["item"] == i), D(0))
        if distributed != item.cantidad:
            findings.append(Hallazgo("ERROR", "DISTRIBUCION", f"Ítem {i}: los totales por WBS suman {distributed}; contrato {item.cantidad}.", item=i))

    activities = []
    for i, item in sorted(budget.items.items()):
        # Incluso sin histórico, un borrador no debe cobrar semanas anteriores.
        # Reconstruir el anterior no equivale a confirmar que ya fue pagado.
        previous = before_period[i]
        if prev and numero(prev["cantidades_acumuladas"][str(i)]) != previous:
            findings.append(Hallazgo("ERROR", "ANTERIOR_NO_CONCILIA",
                f"Ítem {i}: el avance anterior al periodo no coincide con el histórico confirmado. Revise fechas y reportes.", item=i))
        actual = totals[i] - previous
        if actual < 0 or (prev and totals[i] < numero(prev["cantidades_acumuladas"][str(i)])):
            findings.append(Hallazgo("ERROR", "CORTE_NEGATIVO", f"Ítem {i}: acumulado inferior a lo ya cobrado.", item=i))
        if totals[i] > item.cantidad:
            findings.append(Hallazgo("ERROR", "EXCESO_CONTRATO", f"Ítem {i}: acumulado superior a cantidad contratada.", item=i))
        activities.append({"item": i, "anterior": previous, "actual": actual, "acumulada": totals[i],
                           "valor": actual * item.precio, "restante": item.cantidad - totals[i]})
    if not prev and parametros.numero > 1:
        findings.append(Hallazgo("ADVERTENCIA", "ANTERIOR_RECONSTRUIDO",
            "Cantidades anteriores reconstruidas con reportes previos al inicio. No hay corte anterior confirmado; registre los cortes pendientes en orden antes de emitir un definitivo."))
    finance = calcular_finanzas(sum((a["valor"] for a in activities), D(0)), budget.tasas, parametros, config["bases_deducciones"])
    accumulated_direct = sum((a["acumulada"] * budget.items[a["item"]].precio for a in activities), D(0))
    factor = 1 + budget.tasas["administracion"] + budget.tasas["imprevistos"] + budget.tasas["utilidad"] * (1 + budget.tasas["iva"])
    finance["ejecutado_acumulado"] = accumulated_direct * factor
    finance["pendiente"] = budget.total - finance["ejecutado_acumulado"]
    if finance["neto"] < 0:
        findings.append(Hallazgo("ERROR", "NETO_NEGATIVO", "Las deducciones exceden el valor bruto del corte."))
    findings.append(Hallazgo("INFO", "PRECISION", f"Cantidades a {config['decimales_cantidad']} decimales; precios y cálculos monetarios sin redondeo intermedio, presentación a dos decimales."))
    return Resultado(budget, parametros, config, mapping, reports,
                     sorted(grouped.values(), key=lambda g: tuple(map(int, g["wbs"].split('.')))),
                     activities, finance, findings, context, snapshot, source_hashes, earlier)
