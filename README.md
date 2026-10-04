# Proyectos y Macros en VBA

Macros en VBA y scripts en Python desarrollados para cálculo de redes y planificación en sistemas operativos.

---

### 1. `networking/`
Calculadora de subredes IPv4 implementada en VBA para Excel (`.xlsm`) y módulos de código fuente (`.bas`):
- **`Subred_FLSM_VLSM.xlsm`**: Libro de trabajo con hojas y datos de prueba preparados para cada cálculo.
- **`Modulo_Subred.bas`**: Desglose básico de subredes. A partir de una IP y un prefijo base/objetivo, genera la tabla con dirección de red, máscara, primera y última IP utilizable, dirección de broadcast y conteo de hosts útiles.
- **`Modulo_FLSM.bas`**: Máscara de subred de longitud fija (FLSM). Lee una lista de áreas con requerimientos de hosts, calcula la máscara uniforme requerida para el área mayor, genera el desglose binario a nivel de bits con delimitación visual y produce la tabla de asignación completa.
- **`Modulo_VLSM.bas`**: Máscara de subred de longitud variable (VLSM). Asigna bloques de tamaño adaptable por área a partir de un grupo de bloques disponibles, rastrea los bloques padre y genera el árbol binario de división con anotaciones de subred.

---

### 2. `os-scheduling/`
Algoritmos de planificación de CPU y reemplazo de páginas con implementaciones en VBA (visualización en Excel) y Python (salida en consola):

**VBA (`src/` y `.xlsm`):**
- **`Round_Robin.xlsm`**: Libro interactivo con datos de procesos y macros para visualización de planificación.
- **`Page_Replacement.xlsm`**: Libro interactivo para pruebas y renderizado de marcos en reemplazo de páginas.
- **`Round_Robin.bas`**: Planificador Round Robin. Lee la tabla de procesos (nombre, tiempo de llegada, ráfaga), solicita el quantum y genera diagramas de llegada, estado de colas, diagramas de Gantt (por proceso y secuencial), además de métricas de tiempo de espera (WT) y tiempo de retorno (CT) con sus promedios.
- **`Page_Replacement.bas`**: Simulador de reemplazo de páginas. Soporta NRU (con seguimiento de bits R y M), LRU (exacto) y FIFO (con reencolado). Renderiza la cuadrícula de estados de marcos con resaltado de fallos y reporta conteo de fallos, tasa de fallos y rendimiento.

**Python (`python/`):**
- **`round_robin.py`**: Simulación en consola de planificación Round Robin con procesos y quantum configurables.
- **`page_replacement.py`**: Simulador configurable para NRU, LRU (exacto y Aging con contadores de n-bits) y FIFO (cola simple, reencolado y segunda oportunidad). Soporta políticas de desempate, intervalos de reinicio del bit R y configuración de ticks de envejecimiento.

---

## Uso

### Módulos VBA
1. Abrir el libro `.xlsm` correspondiente (o crear uno nuevo).
2. En el Editor de VBA (`Alt + F11`), importar los archivos `.bas` desde **Archivo > Importar archivo**.
3. Configurar los datos en la hoja según lo requerido por la macro y ejecutarla.

### Scripts en Python
```bash
python os-scheduling/python/page_replacement.py
python os-scheduling/python/round_robin.py
```
Edita las variables de configuración al inicio de cada script para modificar los datos de prueba y parámetros de los algoritmos.
