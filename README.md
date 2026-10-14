# Licitaciones de obras · Plataforma de Contratación del Sector Público

Vigilancia diaria de licitaciones de **obras** en plazo publicadas en la
[Plataforma de Contratación del Sector Público](https://contrataciondelestado.es) (PLACSP),
filtradas por ámbito geográfico (por defecto, Comunitat Valenciana y Andalucía).

El script lee el feed ATOM oficial de la Plataforma y los ficheros mensuales de datos abiertos que publica
[Hacienda](https://www.hacienda.gob.es/es-ES/GobiernoAbierto/Datos%20Abiertos/Paginas/LicitacionesContratante.aspx).
Con ellos:

- mantiene una **base local sin duplicados**, siempre con la versión más reciente de cada licitación;
- marca las **nuevas** desde la ejecución anterior;
- registra los **cambios** de plazo, de importe y de estado (incluidas las anulaciones);
- genera un **Excel** con las obras **en plazo**, ordenadas por fecha límite, con las columnas ya tipadas
  (importes, fechas, días restantes, porcentajes, enlaces) para filtrar y tratar los datos;
- se puede **programar** para que se ejecute solo cada día laborable en Windows.

Requisitos: Python 3.8 o superior y el paquete `openpyxl` para el Excel. Los lanzadores `.bat` lo instalan
si falta. Sin `openpyxl`, el script genera un CSV en su lugar.

---

## Uso local (Windows)

| Fichero | Uso |
|---|---|
| `carga_inicial.bat` | **Una sola vez.** Descarga los ficheros mensuales de los 2 últimos meses y crea la base. |
| `programar_tarea.bat` | **Una sola vez.** Deja programada la ejecución automática (ver más abajo). |
| `ejecutar_licitaciones.bat` | Ejecución manual cuando se quiera. Actualiza la base y abre el Excel. |
| `quitar_tarea.bat` | Elimina la ejecución automática. |

Los lanzadores aplican el rango de importe de 500.000 € a 10.000.000 € sin IVA. Para cambiarlo, edita el `.bat`
(y la variable `$Argumentos` de `programar_tarea.ps1` para la ejecución automática).

### Ejecución automática en Windows

`programar_tarea.bat` registra en el Programador de tareas de Windows la tarea **Licitaciones PLACSP**
(no necesita permisos de administrador):

- De lunes a viernes a las **07:45**. Para otra hora: `programar_tarea.bat 08:30` (sustituye la anterior).
- Si el ordenador estaba apagado a esa hora, se ejecuta en cuanto se inicia sesión.
- Se ejecuta sin ventana (`pythonw`). Cada ejecución queda anotada en `registro.log`.
- Si hay licitaciones **nuevas o con cambios**, abre `licitaciones_en_plazo.xlsx` al terminar.
- Comprueba que Python funciona e instala `openpyxl` si falta.

La tarea solo se ejecuta con la sesión de Windows iniciada. Si el Excel está abierto en ese momento,
el resultado se guarda con la fecha y hora en el nombre (`licitaciones_en_plazo_2026-10-14_074512.xlsx`)
y se avisa en `registro.log`.

Desde una consola:

```
python licitaciones_obras.py                         # actualización diaria
python licitaciones_obras.py --meses 2               # carga inicial con ficheros mensuales
python licitaciones_obras.py --min 500000 --max 10000000
python licitaciones_obras.py --nuts ES521            # solo provincia de Alicante
python licitaciones_obras.py --nuts ES52,ES61,ES62   # añadir Región de Murcia
python licitaciones_obras.py --help                  # todas las opciones
```

| Opción | Por defecto | Descripción |
|---|---|---|
| `--nuts` | `ES52,ES61` | Prefijos NUTS del lugar de ejecución (ES52 C. Valenciana, ES61 Andalucía, ES521 Alicante…) |
| `--lugar` | — | Además, texto de provincia o municipio |
| `--paginas` | `30` | Máximo de páginas del feed. Se detiene antes si llega a lo leído el día anterior |
| `--meses` / `--zip AAAAMM` | — | Carga de ficheros mensuales (carga inicial) |
| `--min` / `--max` | — | Rango de presupuesto base sin IVA |
| `--incluir-mixtos` | no | Incluye contratos que no son de tipo obras pero tienen algún CPV 45 |
| `--base` | `base_licitaciones.json` | Ruta de la base JSON |
| `--salida` | `licitaciones_en_plazo.xlsx` | Fichero de resultados `.xlsx` o `.csv`. Se puede repetir para generar los dos |
| `--log` | — | Añade la salida a un fichero de registro |
| `--abrir` | `nunca` | Abre el resultado al terminar: `nunca`, `siempre` o `novedades` |

## Automatización con GitHub Actions

Hay dos flujos en `.github/workflows/`:

- **Pruebas** (`pruebas.yml`): se ejecuta en cada cambio, en Linux y Windows.
- **Actualización diaria** (`actualizacion_diaria.yml`): de lunes a viernes a las 07:15 (hora peninsular en verano).
  Lee la Plataforma, actualiza `datos/base_licitaciones.json`, `datos/licitaciones_en_plazo.csv` y
  `datos/licitaciones_en_plazo.xlsx`, y los guarda en el repositorio.
  El resumen de cada ejecución se ve en la pestaña **Actions**.

Para activarla:

1. En **Settings → Actions → General → Workflow permissions**, marca **Read and write permissions**.
2. Haz la carga inicial: **Actions → Actualización diaria → Run workflow**, con `meses = 2`.
3. A partir de ahí se ejecuta sola. El CSV del día está siempre en `datos/licitaciones_en_plazo.csv`.

Para cambiar el ámbito o el rango de importe, edita los argumentos del paso *Leer la Plataforma* del flujo.

## Resultado en Excel

`licitaciones_en_plazo.xlsx` tiene tres hojas. Las dos de datos son **tablas de Excel** (filtros y orden
en la cabecera, se pueden usar directamente en tablas dinámicas o Power Query).

**En plazo**: una fila por licitación abierta, ordenadas por fecha límite.

| Columna | Tipo y formato |
|---|---|
| Nueva | `SÍ` en verde si apareció desde la ejecución anterior |
| Fecha límite | Fecha y hora (`dd/mm/aaaa hh:mm`) |
| Días restantes | Fórmula: se recalcula con la fecha del día al abrir. Rojo ≤ 7 días, ámbar ≤ 14 |
| Presupuesto base sin IVA, Valor estimado | Número con formato `#.##0,00 €` |
| Comunidad, Provincia, Municipio | Texto (a partir del código NUTS) |
| Órgano de contratación, Objeto, Expediente, Procedimiento | Texto |
| Plazo de ejecución (meses) | Número. Los plazos en días se dividen entre 30; los de años, por 12 |
| Clasificación exigida | Texto (`C4-3; G6-2`). Solo si la Plataforma la informa |
| Juicio de valor | Porcentaje del peso de los criterios sujetos a juicio de valor |
| Criterios de adjudicación, Garantías, Revisión de precios | Texto. `(JV)` = juicio de valor |
| Nº lotes | Número entero |
| CPV | Texto |
| PCAP, PPT, Ficha PLACSP | Enlaces directos |
| Primera detección | Fecha en que entró en la base |
| Cambios | Historial de plazo, importe y estado (resaltado si tiene contenido) |

**Resumen**: fecha de actualización, fuentes leídas, totales y desglose por comunidad y provincia
(número, importe, nuevas y vencimientos en 7 días). Son fórmulas sobre la hoja *En plazo*.

**Histórico**: todas las licitaciones de obras de la base, en cualquier estado (en plazo, pendiente de
adjudicación, adjudicada, resuelta, anulada), sin filtro de importe. Sirve para analizar órganos, importes
y tipologías a lo largo del tiempo.

El libro se regenera en cada ejecución: lo que se anote en él se pierde. El seguimiento de las licitaciones
seleccionadas debe hacerse en el libro de preselección.

## Columnas del CSV

Con `--salida fichero.csv`. Separador `;`, coma decimal y UTF-8 con BOM, para que Excel en español lo abra directamente.

| Columna | Contenido |
|---|---|
| `nueva` | `SÍ` si apareció desde la ejecución anterior |
| `expediente`, `organo_contratacion`, `objeto` | Identificación |
| `presupuesto_base_sin_iva`, `valor_estimado` | Importes |
| `fecha_limite_presentacion`, `dias_restantes` | Plazo de presentación (días naturales desde hoy) |
| `comunidad`, `provincia`, `lugar_ejecucion` | Ubicación (a partir del código NUTS) |
| `procedimiento`, `plazo_ejecucion` | Procedimiento y duración del contrato |
| `clasificacion_exigida` | Grupo-subgrupo-categoría (p. ej. `C4-3`). Solo si la Plataforma la informa |
| `criterios_adjudicacion`, `pct_juicio_valor` | Criterios con su peso. `(JV)` = juicio de valor |
| `garantias`, `revision_precios`, `num_lotes` | Condiciones del contrato |
| `cpv` | Códigos CPV del contrato y de sus lotes |
| `url_pcap`, `url_ppt`, `url_licitacion` | Enlaces directos al pliego administrativo, al técnico y a la ficha |
| `primera_deteccion`, `cambios` | Historial de la licitación en la base |

## Limitaciones conocidas

- **Retraso de la fuente.** La Plataforma publica estos datos con unos días de retraso respecto a la publicación de la licitación.
- **Clasificación exigida.** Solo viene en una parte de las licitaciones; en el resto hay que leerla en el PCAP.
- **Calidad del dato de origen.** Algunos órganos cargan pesos incoherentes (por ejemplo, 100 en cada criterio) o el lugar de ejecución a nivel de comunidad.
- **Sistemas dinámicos de adquisición.** Aparecen en el listado con plazos de varios años; no son obras puntuales.
- **Procedimientos con normas internas** (entidades del sector público no AAPP, como Aena o Navantia): régimen distinto a la LCSP general.
- **Acceso desde GitHub.** No se ha verificado que la Plataforma acepte conexiones desde los servidores de GitHub. La primera ejecución manual lo confirma; si falla, se usa la ejecución local.

## Estructura

```
├── licitaciones_obras.py          Script principal
├── carga_inicial.bat              Carga inicial (Windows, una vez)
├── programar_tarea.bat / .ps1     Programa la ejecución automática (Windows, una vez)
├── quitar_tarea.bat               Elimina la ejecución automática
├── LEEME.txt                      Instrucciones de instalación y uso (Windows)
├── ejecutar_licitaciones.bat      Ejecución manual (Windows)
├── datos/                         Resultados versionados (los genera GitHub Actions)
├── tests/test_licitaciones.py     Pruebas con un feed sintético de estructura CODICE real
└── .github/workflows/             Pruebas y actualización diaria
```

Para ejecutar las pruebas en local: `python -m unittest discover -s tests -v`
