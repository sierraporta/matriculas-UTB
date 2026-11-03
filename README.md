# Monitoreo de matrículas de primera vez, UTB

Script que convierte los reportes BAN "Retención Levantada" en un documento de diapositivas (`.pptx`) y un resumen en Excel para dar seguimiento a la matrícula nueva por cohortes, programas y escuelas, frente a las metas de cada periodo.

## Aviso sobre los datos

**Los datos originales no se incluyen en este repositorio por razones éticas.** Los reportes BAN contienen información personal de estudiantes (nombres, documentos de identidad, fechas de nacimiento, contacto), y su difusión no está autorizada. Por eso `.gitignore` excluye todos los archivos `.xls` y `.csv` de los reportes, las salidas generadas y los notebooks originales, cuyas celdas podían mostrar filas de los datos.

El script solo escribe agregados (conteos por programa, escuela y periodo). Ningún dato personal sale de la memoria del proceso. Lo único que se versiona es `metas.csv`, que contiene metas institucionales por programa y no datos de personas.

Quien quiera reproducir el análisis necesita sus propios reportes, con acceso autorizado a ICEBERG.

## Qué produce

Con los reportes en la carpeta, `matriculas.py` genera en `Salida/`:

- `Matriculas_AAAAMMDD.pptx`: documento con portada (fecha de reporte y fecha de corte de los datos), resumen, seguimiento a igual día, proyección de cierre y su método, curvas acumuladas por escuela (dos programas por diapositiva), tablas comparativas por programa y por escuela, brecha frente a la meta, permanencia por cohorte y calidad de los datos.
- `Matriculas_AAAAMMDD.pdf`: copia en PDF del mismo documento, para compartir. Requiere LibreOffice instalado; si no está, el script lo avisa y sigue sin el PDF.
- `Matriculas_resumen.xlsx`: las tablas detrás de cada diapositiva (cumplimiento, tipos de estudiante, curvas diarias, prueba retrospectiva de la proyección, perfil de ingreso).
- `resumen.txt`: resumen corto en texto.

## Uso

```bash
pip install -r requirements.txt
python3 matriculas.py
```

Opciones:

| Opción | Efecto |
|---|---|
| `--sin-pdf` | No exporta el PDF. |
| `--fecha-reporte AAAA-MM-DD` | Fecha que aparece en la portada (por defecto, hoy). |
| `--por-lamina N` | Programas por diapositiva en cada escuela (por defecto 2). |
| `--anios N` | Cohortes a mostrar por semestre (por defecto todas). |
| `--bruto` | No descarta devoluciones ni duplicados (conteo antiguo). |
| `--cerrado` | Trata el último periodo como cerrado. |
| `--figuras` | Guarda además las figuras sueltas en `Salida/Figuras`. |
| `--datos`, `--salida`, `--metas` | Rutas alternativas. |

## Entradas esperadas

- Reportes con nombre `BAN - Reporte de Retención Levantada AAAAPP.xls` (TSV con encabezado, latin-1), donde `PP` es `10` o `20`. Si solo hay `.csv` (sin encabezado, 28 columnas en el orden del reporte), también se leen; si existen ambos, se usa el `.xls`.
- `metas.csv` con columnas `anio,sem,programa,escuela,meta`. Las líneas que empiezan por `#` se ignoran. Si falta la meta de un año, se usa la más reciente como referencia y la diapositiva lo indica.

## Definiciones

- **Matriculado nuevo**: `Tipo_estudiante == PRIMERA VEZ`, sin `Devolucion == Si`, con una fila por (Id, Programa) que conserva el último pago.
- **Origen del ciclo**: 1 de noviembre del año anterior para el periodo `10`, 1 de mayo para el `20`. Las curvas son conteos acumulados por día desde ese origen, suavizados con interpolación monótona (PCHIP) sobre puntos semanales.
- **Periodo en curso**: el último periodo se considera abierto si su último pago tiene menos de 30 días respecto a la fecha de reporte.
- **SPC %**: 200·(real − meta)/(real + meta), simétrico respecto a la meta.
- **Permanencia**: porcentaje de la cohorte cuyo Id reaparece con pago en periodos posteriores cerrados. No distingue deserción de grado, pausa o cambio de programa.

## Proyección de cierre

Para cada cohorte cerrada del mismo semestre se calcula qué fracción de su cierre ya estaba pagada el mismo día del ciclo. La cohorte actual se divide entre esa fracción, y cada cohorte de referencia aporta una estimación; el rango es mínimo y máximo, y el valor central la mediana. Se valida con una prueba retrospectiva (cada cohorte cerrada se proyecta con las demás) y la proyección solo se declara confiable desde el primer día en que el error máximo queda bajo 15%. Con pocas cohortes el rango es indicativo y no un intervalo de confianza, y el método supone que el calendario de pagos se repite. Se calcula solo para el total de la UTB.

## Estructura

```
matriculas.py      script principal
metas.csv          metas por programa y periodo
requirements.txt   dependencias
Salida/            resultados (no versionado)
```
