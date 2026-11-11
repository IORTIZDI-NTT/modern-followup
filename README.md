# Seguimiento Modernización Lote 4

Generador de un informe HTML mensual a partir de la GECO, el horario laboral del equipo y el informe de incurrido.

## Preparación

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Coloca en `data/input/`:

- una GECO en formato `.xlsb` con las hojas `Forecast`, `Equipo` e `Ingresos`;
- un Excel `Horario Laboral Equipo.xlsx` con una hoja por mes;
- el informe `YYYYMM-Incurrido-....html` del mes que se va a procesar.

Los datos reales se almacenan en este repositorio privado. Los `.xlsb` usan Git LFS para evitar que el historial normal de Git crezca innecesariamente.

## Generación

El mes actual de Europe/Madrid es el valor predeterminado:

```bash
.venv/bin/python -m src.build_report
```

También se puede indicar explícitamente:

```bash
.venv/bin/python -m src.build_report --month 2026-11
```

El sitio se genera en `site/index.html`. Incluye las páginas Resumen, Economía, Equipo, Calendario y Datos.

## Publicación

El workflow `.github/workflows/publish-report.yml` ejecuta las pruebas, genera el mes actual y publica `site/` mediante GitHub Pages cada vez que cambian los datos o el generador. También permite una ejecución manual indicando un mes `YYYY-MM`.

En la configuración del repositorio, GitHub Pages debe utilizar **GitHub Actions** como origen y conservar la visibilidad restringida de la organización. No se debe cambiar el sitio a visibilidad pública porque contiene información personal y económica.

Para distinguir las personas inactivas marcadas en rojo, el generador utiliza `soffice`/LibreOffice para crear una copia local de estilos en `data/processed/`. Esa copia es una caché y no se versiona.

## Criterios iniciales

- Horas facturables: `Total validado` de las filas CLIENTE. Las horas NTT se conservan aparte.
- El KPI global de horas facturables incluye todas las filas CLIENTE del HTML; la pestaña Equipo y el calendario detallado muestran únicamente personas activas y concilian la diferencia por separado.
- Horas pendientes: máximo entre horas planificadas en Equipo y horas CLIENTE facturables.
- Facturación estimada: horas CLIENTE × 28,36 €/h.
- Forecast actual: escenario final que ya contiene los ajustes de bolsa o reducciones.
- Forecast sin ajustes: forecast actual menos el ajuste correspondiente.
- El bloque de ajustes se alinea por posición con Forecast actual porque sus fechas internas están desplazadas.
- Solo se agregan personas de la primera tabla de Equipo; las filas rojas se excluyen y las tablas resumen se presentan aparte.
- Capacidad: horas L–V del horario multiplicadas por los días de cada semana del mes; todavía no descuenta un calendario oficial de festivos.
