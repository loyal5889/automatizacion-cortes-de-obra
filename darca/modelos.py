from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import hashlib
import json
import re
import unicodedata

D = Decimal
ZERO = D(0)


class ErrorDatos(ValueError):
    """Error de entrada que no puede sustituirse por cero."""


def texto(value):
    return " ".join(str("" if value is None else value).split())


def normalizar(value):
    value = unicodedata.normalize("NFKD", texto(value)).casefold()
    return "".join(c for c in value if not unicodedata.combining(c))


def unidad(value):
    key = normalizar(value).replace(" ", "").replace("²", "2")
    return {"m2": "m²", "mt2": "m²", "unidad": "und", "unidades": "und",
            "un": "und", "und": "und", "u": "und", "ml": "ml",
            "m.l.": "ml", "metrolineal": "ml"}.get(key, key)


def numero(value, contexto="valor"):
    if isinstance(value, bool) or value is None or texto(value) == "":
        raise ErrorDatos(f"{contexto}: falta un número válido (cero debe ser explícito).")
    try:
        if isinstance(value, str):
            s = value.strip().replace(" ", "")
            # Sin separador de miles: evita interpretar 1.234 como 1234 sin saberlo.
            if "," in s and "." in s:
                raise ValueError("Use números sin separador de miles")
            value = s.replace(",", ".")
        result = D(str(value))
        if not result.is_finite():
            raise ValueError("No finito")
        return result
    except (InvalidOperation, ValueError):
        raise ErrorDatos(f"{contexto}: {value!r} no es numérico. Use 1234,56 o 1234.56 sin miles.") from None


def porcentaje(value):
    result = numero(str(value).strip().removesuffix("%"), "porcentaje")
    if not ZERO <= result <= 100:
        raise ErrorDatos("El porcentaje debe estar entre 0 y 100; escriba 40 para 40%.")
    return result / 100


def cantidad(value, decimales=6):
    return numero(value).quantize(D(1).scaleb(-decimales), rounding=ROUND_HALF_UP)


def fecha(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto(value), fmt).date()
        except ValueError:
            pass
    raise ErrorDatos(f"Fecha inválida: {value!r}. Use DD/MM/AAAA o AAAA-MM-DD.")


def serializable(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, D):
        return str(value)
    raise TypeError(type(value).__name__)


def firma(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False,
                                    default=serializable).encode("utf-8")).hexdigest()


def hash_archivo(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


@dataclass
class Hallazgo:
    nivel: str
    codigo: str
    mensaje: str
    archivo: str = ""
    fila: int | None = None
    item: int | None = None


@dataclass
class Item:
    numero: int
    concepto: str
    descripcion: str
    unidad: str
    cantidad: D
    precio: D
    total: D
    fila: int


@dataclass
class Presupuesto:
    archivo: str
    items: dict[int, Item]
    tasas: dict[str, D]
    total: D
    directo: D
    huella: str


@dataclass
class Registro:
    fecha: date
    id: str
    wbs: str
    actividad: str
    tramo: str
    unidad: str
    total: D
    ejecutada: D
    observacion: str
    archivo: str
    fila: int


@dataclass
class Reporte:
    archivo: str
    registros: list[Registro]
    huella: str
    sha256: str

    @property
    def fecha(self):
        return self.registros[0].fecha


@dataclass
class Parametros:
    numero: int
    inicio: date
    fin: date
    amortizacion: D
    retefuente: D
    garantia: D
    otros: D = ZERO

    def validar(self):
        if type(self.numero) is not int or self.numero < 1:
            raise ErrorDatos("El número del corte debe ser un entero positivo.")
        if self.inicio > self.fin:
            raise ErrorDatos("La fecha inicial es posterior a la final.")
        for name in ("amortizacion", "retefuente", "garantia"):
            if not getattr(self, name).is_finite() or not ZERO <= getattr(self, name) <= 1:
                raise ErrorDatos(f"{name}: porcentaje fuera del rango 0–100%.")
        if not self.otros.is_finite() or self.otros < 0:
            raise ErrorDatos("Otros descuentos debe ser un número no negativo.")


@dataclass
class Resultado:
    presupuesto: Presupuesto
    parametros: Parametros
    config: dict
    mapeo: dict
    reportes: list[Reporte]
    consolidado: list[dict]
    actividades: list[dict]
    finanzas: dict[str, D]
    hallazgos: list[Hallazgo]
    firma_contexto: str
    firma_historico: str
    archivos_origen: dict[str, str] = field(default_factory=dict)
    cortes_anteriores: list[dict] = field(default_factory=list)

    @property
    def criticos(self):
        return any(h.nivel == "ERROR" for h in self.hallazgos)
