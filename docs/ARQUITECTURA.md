# Arquitectura y decisiones

| Componente | Responsabilidad |
|---|---|
| `generar_corte.py` | Parámetros, resumen y confirmación del usuario |
| `darca/lectura.py` | Esquema de fuentes, tipos y lectura de fórmulas admitidas |
| `darca/calculo.py` | Consolidación, periodos, cantidades y finanzas |
| `darca/tramos.py` | Compatibilidad controlada de un desglose contractual |
| `darca/historico.py` | Firmas, secuencia y escritura atómica bajo bloqueo |
| `darca/excel.py` | Presentación, fórmulas, resultados guardados y verificación |

**Decimal.** Evita redondear precios intermedios; la presentación monetaria utiliza dos decimales.

**Histórico.** Guarda acumulados por ítem. La cantidad de un corte individual se obtiene restando acumulados consecutivos. La firma detecta alteraciones, pero no equivale a una firma digital de identidad ni a un sistema de permisos.

**Excel.** Las fórmulas permanecen visibles y tienen resultados guardados para lectores que no recalculan. El generador no ejecuta macros de las fuentes.

**Confirmación.** Generar un archivo es distinto de registrar una cuenta. La escritura definitiva verifica que el histórico y las fuentes no hayan cambiado desde el cálculo.

**Portabilidad.** Se publica el motor y un conjunto real de DARCA II en `fuentes_darca/`. `scripts/preparar_darca.py` prepara una copia local. Las pruebas usan datos sintéticos aislados para evitar alterar los históricos reales.
