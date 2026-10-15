# Versiones - A11y-Help (VBA Excel)

Historial de cambios del proyecto **A11y-Help**, asistente en Excel con macros (VBA) para estandarizar hallazgos y criterios de accesibilidad web basados en `axe-core`.

---

## Versión 0.2.0

> Fecha: 15-10-2026  
> Autor: Diego Javier Mena Amado

### Carga automática del JSON (Local / Nube), caché en memoria y búsqueda inteligente

- **Resolución automática del archivo `es.json` sin editar código VBA (`ObtenerCatalogoA11y`)**:
  1. Busca automáticamente `es.json` en la misma carpeta donde se encuentra el archivo Excel (`ThisWorkbook.Path`).
  2. Si no está en la carpeta local, lo descarga automáticamente desde GitHub (`https://raw.githubusercontent.com/ingelectronicadj/a11y-help-VBAexcel/main/es.json`).
  3. Si tampoco hay conexión a internet, abre un cuadro de diálogo (`Application.GetOpenFilename`) para elegir el archivo `.json` manualmente sin necesidad de entrar al editor de VBA ni depurar.
- **Caché en memoria (`gCatalogoCache`)**:
  - El catálogo JSON se procesa una sola vez por sesión y queda en memoria, haciendo que las aperturas posteriores del formulario sean instantáneas.
  - Se agregó la macro `RecargarCatalogoJSON` para limpiar la caché si se edita el archivo `es.json`.
- **Búsqueda inteligente multi-palabra y por ID técnico (`CoincideBusqueda`)**:
  - Ahora busca simultáneamente en el **ID de la regla (`clave`, ej. `image-alt`, `color-contrast`)**, en la **descripción (`description` / `pass` / `fail`)** y en el **texto corto de ayuda (`help`)**.
  - Soporta múltiples palabras clave separadas por espacio en cualquier orden.
  - Al abrir el formulario, se lista el catálogo completo desde el inicio y se preselecciona la primera coincidencia.
- **Mejoras de usabilidad en `UserForm1`**:
  - Soporte para **doble clic (`lstResultados_DblClick`)** sobre cualquier resultado para insertarlo directamente en la celda activa.
  - Eliminación del procedimiento vacío residual `optCorrecta_Click`.
- **Correcciones en `es.json`**:
  - Se añadieron los campos `description` y `help` faltantes en la regla `aria-unsupported-attr` para completar las **105 reglas** y **80 comprobaciones (`checks`)**.
- **Exportación estándar VBA y configuración**:
  - Se añadió `Modulos/AbrirFormulario.bas` con cabecera estándar `Attribute VB_Name` para importación directa en el editor de VBA.
  - Se incorporó `.gitignore` para excluir archivos temporales de bloqueo de Excel (`~$*.xlsm`).

---

## Versión 0.1.0

> Fecha: 23-06-2025  
> Autor: Diego Javier Mena Amado

- Versión inicial del asistente en Excel (`Test-a11y.xlsm`) con `UserForm` en VBA, integración con `JsonConverter.bas` y diccionario en español `es.json` para reglas (`description`) y comprobaciones (`pass` / `fail`).
