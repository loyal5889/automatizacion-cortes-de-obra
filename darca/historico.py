"""Histórico confirmado, inmutable y con escritura atómica bajo bloqueo exclusivo."""
from dataclasses import asdict
from datetime import datetime
import json
import os
from pathlib import Path
import tempfile

from .modelos import D, ErrorDatos, fecha, firma, hash_archivo, numero, serializable


def leer_historico(folder):
    records = []
    for path in sorted(folder.glob("*.json")):
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
            checksum = doc.pop("integridad")
            if firma(doc) != checksum or doc["version"] != 1:
                raise ValueError("Integridad o versión inválida")
            if type(doc["numero"]) is not int or doc["numero"] < 1:
                raise ValueError("Número inválido")
            if path.name != f"corte_{doc['numero']:04d}.json":
                raise ValueError("El nombre del archivo no coincide con su número")
            if not doc["cantidades_acumuladas"] or any(numero(v) < 0 for v in doc["cantidades_acumuladas"].values()):
                raise ValueError("Cantidades históricas inválidas")
            if fecha(doc["inicio"]) > fecha(doc["fin"]):
                raise ValueError("Periodo inválido")
            for key in ("bruto", "neto", "deducciones", "contexto", "reportes", "excel", "excel_sha256"):
                if key not in doc:
                    raise ValueError(f"Campo faltante: {key}")
            records.append(doc)
        except (OSError, ValueError, KeyError, TypeError) as exc:
            raise ErrorDatos(f"Histórico inválido {path.name}: {exc}. No se asumirá corte anterior cero.") from exc
    records.sort(key=lambda x: x["numero"])
    if [r["numero"] for r in records] != list(range(1, len(records) + 1)):
        raise ErrorDatos("El histórico tiene números duplicados o cortes faltantes. Revise su respaldo.")
    for prev, current in zip(records, records[1:]):
        if fecha(current["inicio"]) <= fecha(prev["fin"]):
            raise ErrorDatos("El histórico contiene periodos superpuestos.")
        if set(prev["cantidades_acumuladas"]) != set(current["cantidades_acumuladas"]):
            raise ErrorDatos("El histórico tiene conjuntos de ítems incompatibles.")
        if any(numero(current["cantidades_acumuladas"][k]) < numero(v)
               for k, v in prev["cantidades_acumuladas"].items()):
            raise ErrorDatos("Las cantidades confirmadas disminuyen entre cortes del histórico.")
        if current["contexto"] != prev["contexto"]:
            raise ErrorDatos("El histórico mezcla contratos o mapeos distintos.")
    return records, firma(records)


def registrar(root, resultado, exportar):
    """Llamar únicamente después de la confirmación explícita de la consola."""
    if resultado.criticos:
        raise ErrorDatos("Hay errores críticos. No se permite registrar el corte.")
    folder = root / "data/historico"
    folder.mkdir(parents=True, exist_ok=True)
    lock = folder / ".lock"
    try:
        fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        raise ErrorDatos("Otro proceso está registrando un corte (data/historico/.lock).") from None
    final = None
    temp = None
    try:
        os.close(fd)
        records, snapshot = leer_historico(folder)
        if snapshot != resultado.firma_historico:
            raise ErrorDatos("El histórico cambió desde el cálculo. Genere nuevamente el corte.")
        p = resultado.parametros
        if p.numero != len(records) + 1:
            raise ErrorDatos("Número existente o fuera de secuencia. No se sobrescribe el histórico.")
        for relative, digest in resultado.archivos_origen.items():
            path = root / relative
            if not path.is_file() or hash_archivo(path) != digest:
                raise ErrorDatos(f"La fuente cambió desde el cálculo: {relative}. Vuelva a generar.")
        # La exportación definitiva y su control se crean bajo el mismo bloqueo.
        final = exportar(resultado, root / "output", definitivo=True)
        finance = resultado.finanzas
        doc = {
            "version": 1, "proyecto": resultado.config["proyecto"], "numero": p.numero,
            "fecha": datetime.now().astimezone().isoformat(), "inicio": p.inicio, "fin": p.fin,
            "cantidades_acumuladas": {str(a["item"]): a["acumulada"] for a in resultado.actividades},
            "bruto": finance["bruto"], "neto": finance["neto"],
            "deducciones": {k: finance[k] for k in ("amortizacion", "retefuente", "garantia", "otros")},
            "parametros": asdict(p), "bases_deducciones": {**resultado.config["bases_deducciones"], "retefuente": "bruto_menos_iva"},
            "reportes": [{"archivo": r.archivo, "huella": r.huella, "sha256": r.sha256,
                          "fecha": r.fecha} for r in resultado.reportes],
            "contexto": resultado.firma_contexto,
            "excel": final.relative_to(root).as_posix(), "excel_sha256": hash_archivo(final)
        }
        # Canonicalizar antes de firmar para verificar igual tras leer fechas y decimales.
        doc = json.loads(json.dumps(doc, default=serializable, ensure_ascii=False))
        doc["integridad"] = firma(doc)
        destination = folder / f"corte_{p.numero:04d}.json"
        if destination.exists():
            raise ErrorDatos("El número del corte ya existe.")
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=folder,
                                         suffix=".tmp", delete=False) as handle:
            temp = Path(handle.name)
            json.dump(doc, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp, destination)
        return final, destination
    except Exception:
        if final is not None and final.exists():
            final.unlink()
        raise
    finally:
        if temp is not None and temp.exists():
            temp.unlink()
        lock.unlink(missing_ok=True)
