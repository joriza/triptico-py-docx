# Kit docente: tríptico para un emprendimiento casero

Material completo para un **taller de 4 clases** (secundaria, ~16 años, encuentros de 60–80 min) cuyo producto final es un **tríptico impreso** que promociona un emprendimiento casero de gelatina. Incluye los 4 materiales didácticos (guía docente, cuaderno del alumno, modelo y rúbrica), el modelo editable en Word y un generador determinista de trípticos de ejemplo.

## Camino rápido

1. **Dictar el taller** con los 4 documentos impresos (sección Material didáctico).
2. **Editar el modelo** `triptico-dulce-casero.docx` en Word 2010: cada panel es una celda de una tabla 1×3; las líneas punteadas son guías de pliegue.
3. **Regenerar los trípticos de ejemplo** (opcional):

   ```cmd
   python generar-tripticos-docx.py [todos|gelatina|mermelada|gomitas|panes]
   ```

   Sin argumento regenera los 4. Cada corrida verifica determinismo y XML antes de escribir.

## Material didáctico

| Documento | Archivo | Público |
|-----------|---------|---------|
| Guía docente | `guia-docente-triptico-gelatina.md` / `.pdf` | Docente (impersonal formal) |
| Cuaderno del alumno | `cuaderno-del-alumno-triptico-gelatina.md` / `.pdf` | Alumno (voseo rioplatense) |
| Modelo resuelto | `modelo-triptico-gelatina.md` / `.pdf` | Referencia (los 6 paneles) |
| Rúbrica | `rubrica-triptico-gelatina.md` / `.pdf` | Evaluación (5 criterios × 4 niveles) |
| Modelo editable | `triptico-dulce-casero.docx` + `triptico-dulce-casero.md` / `.pdf` | Alumno (parte desde este archivo) |

## Trípticos de ejemplo por producto

Generados por `generar-tripticos-docx.py`, con datos ficticios etiquetados en el propio documento:

| Clave | Archivo | Emprendimiento |
|-------|---------|----------------|
| `gelatina` | `triptico-gelatina-la-perla.docx` | Gelatina presentada como postre |
| `mermelada` | `triptico-mermelada.docx` | Mermeladas de fruta de estación |
| `gomitas` | `triptico-gomitas.docx` | Gomitas con jugo de fruta real |
| `panes` | `triptico-panes.docx` | Pan de garbanzos y de almendras (un solo equipo) |

## Estructura del tríptico (hoja A4 apaisada, 2 caras)

| Cara | Panel 1 | Panel 2 | Panel 3 |
|------|---------|---------|---------|
| Externa | Solapa | Contratapa (contacto) | Portada |
| Interna | Interior 1 | Interior 2 | Interior 3 |

## Generador: reglas y códigos de salida

- **Determinista**: misma entrada → archivo byte a byte idéntico (fechas de zip fijas; doble corrida de verificación integrada).
- El contenido vive como datos dentro del propio script: hay **una única forma de regenerar**.
- Única dependencia no estándar: `defusedxml` (validación de XML).

| Código | Significado |
|--------|-------------|
| 0 | Todo bien |
| 2 | Argumento inválido |
| 3 | Datos internos inválidos |
| 4 | Determinismo roto |
| 5 | XML mal formado |
| 6 | Error de escritura |

## Regeneración de los PDF del kit

Los PDF del kit se generan con el **pipeline global de la máquina** (fuera de este repo): `generar-pdf` (Chrome headless, en `C:\Users\USER\herramientas\pdf`). Documentos del kit con defaults:

```powershell
generar-pdf -Doc guia-docente-triptico-gelatina
```

El tríptico de ejemplo lleva receta propia (apaisado, simétrico angosto, sin encabezado):

```powershell
generar-pdf -Doc triptico-dulce-casero -Orientacion apaisado -SinEncabezado -MargenLateral 1.2cm
```

**Requisitos:** Word 2010 o superior para editar los `.docx`; Python 3 + `defusedxml` para el generador; Chrome + pandoc para los PDF. Dos corridas con la misma entrada producen archivos **byte a byte idénticos**, tanto en PDF como en `.docx`.

## Estándares del proyecto

- Los PDF del kit se generaron con Chrome headless: márgenes espejados (lomo 1,5 cm / externo 1,0 cm), margen vertical 1,2 cm, encabezado con el nombre del documento y fecha ISO-8601 en el pie del lado del lomo.
- Los markdown **no usan LaTeX**: todo `$` es literal (precios, referencias).
- Convención de idioma: voseo en material del alumno; impersonal formal en material docente.

## Documentación interna

- `odd/tasks/triptico-gelatina-casera.md` — especificaciones (S1–S29), tareas (T1–T25) y log de decisiones.
- `doc/pdt-triptico.md` — cuaderno de pendientes y pedidos del usuario.
