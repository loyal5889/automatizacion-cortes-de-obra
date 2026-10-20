<p align="center"><strong>INGENIERÍA · AUTOMATIZACIÓN · CONTROL DE OBRA</strong></p>

# Automatización de cortes de obra

**De reportes semanales a un corte trazable en Excel.** Un programa en Python que conecta cantidades ejecutadas, presupuesto contractual e históricos para preparar cortes de obra con fórmulas auditables.

![Presentación del proyecto](docs/images/portada.png)

**Python · openpyxl · Decimal · Excel · JSON · unittest**

[Consultar DARCA II](#caso-real-darca-ii) · [Ver los cortes definitivos](examples/) · [Decisiones técnicas](docs/ARQUITECTURA.md) · [Proyecto complementario: Curva S](https://github.com/loyal5889/dashboard-avance-programado-vs-real-curva-s)

## El problema

Preparar un corte exige relacionar varios reportes con el presupuesto, descontar cantidades anteriores y aplicar deducciones. Hacerlo manualmente facilita errores como cobrar dos veces un avance o dejar fuera un tramo de la misma actividad.

## La solución

El programa identifica actividades por WBS, consolida reportes y calcula únicamente la ejecución correspondiente al periodo solicitado. Produce un libro con tres hojas: **Corte**, **Consolidado reportes** y **Presupuesto**.

| Función | Qué permite revisar |
|---|---|
| Una columna por corte confirmado | Cantidad individual de cada periodo, actual y acumulada |
| Vínculo WBS → ítem contractual | Varios tramos sumados dentro del mismo ítem |
| Precio, saldo y ejecución | Cantidades y valores conectados mediante fórmulas |
| Anticipo, retefuente y garantía | Porcentajes ingresados por el usuario y bases visibles |
| Histórico JSON con firma | Detección de cambios en registros y reportes ya utilizados |
| Borrador y confirmación separados | Revisar el Excel antes de registrar un corte |

## Caso real: DARCA II

Este repositorio presenta los archivos reales disponibles al **20 de octubre de 2026**:

- [Cortes definitivos 1 y 2](examples/), con sus cantidades, valores y presupuesto.
- [Presupuesto y reportes semanales](fuentes_darca/).
- [Históricos confirmados](fuentes_darca/historico/).

Los Excel se publican sin modificar respecto a los originales. El corte 2 abarca del 20 de septiembre al 20 de octubre de 2026. Las fechas propias de cada reporte determinan su inclusión; el nombre semanal de un archivo no significa que todo su periodo haya sido ejecutado.

### Ejecutar una copia local

Requiere Python 3.11 o posterior. Después de clonar el repositorio:

```bash
python -m pip install -r requirements.txt
python scripts/preparar_darca.py
python generar_corte.py --manual
```

El preparador copia las fuentes reales a `data/` y los cortes definitivos a `output/`. Se detiene si esos destinos ya existen para evitar mezclar históricos. En Windows, `INICIAR_CORTE.bat` inicia el flujo interactivo. Los siguientes cortes deben usar el número y periodo que correspondan; generar un borrador no confirma un histórico.

## Flujo de información

```mermaid
flowchart LR
  P[Presupuesto contractual] --> C[Cálculo con Decimal]
  R[Reportes semanales / DATOS] --> W[Consolidación por WBS]
  W --> C
  H[Histórico confirmado] --> C
  C --> E[Excel preliminar]
  E --> V[Revisión del usuario]
  V --> J[Registro confirmado]
```

## Mi trabajo en este proyecto

- Traducción del proceso de cortes a reglas verificables de cantidades, periodos y deducciones.
- Implementación de lectura de Excel, consolidación, generación de fórmulas y registro histórico.
- Resolución de casos reales de estructura: división por tramos e históricos por corte.
- Preparación de un flujo local con lanzador para Windows y pruebas de regresión.

## Validación

```bash
python -m unittest discover -s tests -v
```

La edición pública incluye pruebas con fuentes sintéticas para tres cortes consecutivos, retefuente, protección frente a sobrescritura del ejemplo y compatibilidad del desglose por tramos. El código contrasta los resultados del Excel con el cálculo de Python antes de entregar el archivo.

## Alcance y límites

- La estructura de las fuentes debe respetar los encabezados y la hoja `DATOS`. El mapeo se configura en `config/`.
- La regla de desglose de lavado incluida está acotada a tres tramos de 64 m²; no representa una conversión general automática para cualquier contrato.
- La retefuente se calcula como `(bruto − IVA sobre utilidad) × porcentaje`; es una regla del flujo implementado, no asesoría tributaria.
- El histórico exige confirmación explícita. Los borradores no se consideran cortes registrados.
- La edición pública incluye archivos reales de DARCA II. Las fuentes sintéticas se conservan únicamente como apoyo de las pruebas automatizadas.

**Autor:** [loyal5889](https://github.com/loyal5889) · [Explorar el dashboard de avance](https://github.com/loyal5889/dashboard-avance-programado-vs-real-curva-s)
