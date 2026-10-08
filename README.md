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
- genera un CSV con las obras **en plazo**, ordenadas por fecha límite, con los datos necesarios para la preselección.

Solo usa la biblioteca estándar de Python (3.8 o superior). No hay dependencias que instalar.

---

## Uso local (Windows)

| Fichero | Uso |
|---|---|
| `carga_inicial.bat` | **Una sola vez.** Descarga los ficheros mensuales de los 2 últimos meses y crea la base. |
| `ejecutar_licitaciones.bat` | **Cada día.** Actualiza la base y abre `licitaciones_en_plazo.csv` en Excel. |

Los dos lanzadores aplican el rango de importe de 500.000 € a 10.000.000 € sin IVA. Para cambiarlo, edita el `.bat`.

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
| `--min` / `--max` | — | Rango de presupuesto base sin IVA para el CSV |
| `--incluir-mixtos` | no | Incluye contratos que no son de tipo obras pero tienen algún CPV 45 |
| `--base` / `--salida` | raíz | Rutas de la base JSON y del CSV |

## Automatización con GitHub Actions

Hay dos flujos en `.github/workflows/`:

- **Pruebas** (`pruebas.yml`): se ejecuta en cada cambio, en Linux y Windows.
- **Actualización diaria** (`actualizacion_diaria.yml`): de lunes a viernes a las 07:15 (hora peninsular en verano).
  Lee la Plataforma, actualiza `datos/base_licitaciones.json` y `datos/licitaciones_en_plazo.csv`, y los guarda en el repositorio.
  El resumen de cada ejecución se ve en la pestaña **Actions**.

Para activarla:

1. En **Settings → Actions → General → Workflow permissions**, marca **Read and write permissions**.
2. Haz la carga inicial: **Actions → Actualización diaria → Run workflow**, con `meses = 2`.
3. A partir de ahí se ejecuta sola. El CSV del día está siempre en `datos/licitaciones_en_plazo.csv`.

Para cambiar el ámbito o el rango de importe, edita los argumentos del paso *Leer la Plataforma* del flujo.

## Columnas del CSV

Separador `;`, coma decimal y UTF-8 con BOM, para que Excel en español lo abra directamente.

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
├── ejecutar_licitaciones.bat      Lanzador diario (Windows)
├── carga_inicial.bat              Lanzador de carga inicial (Windows)
├── datos/                         Resultados versionados (los genera GitHub Actions)
├── tests/test_licitaciones.py     Pruebas con un feed sintético de estructura CODICE real
└── .github/workflows/             Pruebas y actualización diaria
```

Para ejecutar las pruebas en local: `python -m unittest discover -s tests -v`
