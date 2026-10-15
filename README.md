# A11y-Help: Asistente para Auditoría de Accesibilidad en Excel

![Version](https://img.shields.io/badge/version-0.2.0-0f766e.svg)
![VBA](https://img.shields.io/badge/Excel-VBA_%2F_Macros-217346.svg)
![axe-core](https://img.shields.io/badge/axe--core-105_Reglas_%7C_80_Checks_(ES)-D9381E.svg)
![License](https://img.shields.io/badge/License-GPL_v3-blue.svg)

## Descripción
**A11y-Help** es una herramienta en **Microsoft Excel con macros (VBA)** diseñada para agilizar y estandarizar la documentación de auditorías de accesibilidad digital. Permite buscar en tiempo real entre **105 reglas** y **80 comprobaciones (`checks`)** de `axe-core` en español (`es.json`) e insertar con un clic (o doble clic) el comportamiento esperado, el criterio de éxito (`pass`) o el fallo (`fail`) directamente en la celda activa.

## Uso rápido (sin necesidad de editar código VBA)
1. **Descargar o clonar** este repositorio (manteniendo `Test-a11y.xlsm` y `es.json` en la misma carpeta).
2. **Abrir** `Test-a11y.xlsm` en Excel y habilitar las macros si se solicita.
3. **Presionar el botón** en la hoja de cálculo para abrir el formulario de búsqueda (`AbrirFormulario`).
   - El asistente cargará automáticamente `es.json` desde la misma carpeta del archivo Excel.
   - Si mueves únicamente el archivo `.xlsm` a otra ubicación, el asistente descargará automáticamente el catálogo desde GitHub o te permitirá seleccionarlo mediante una ventana de diálogo.
4. **Seleccionar el tipo de texto a consultar**:
   - **Comportamiento esperado** (`rules.description` / `rules.help`)
   - **Correcto** (`checks.pass`)
   - **Incorrecto** (`checks.fail`)
5. **Buscar** escribiendo cualquier palabra clave en español o el ID técnico de la regla (por ejemplo: `color-contrast`, `image-alt`, `aria nombre`, `teclado`).
6. Hacer **doble clic** sobre el resultado (o pulsar **Insertar**) para escribir la descripción en la celda activa de Excel.

### Cómo actualizar los módulos en el Editor de VBA (Opcional)
Si deseas actualizar manualmente el código dentro de tu propio libro `.xlsm`:
1. En Excel, presiona `Alt + F11` para abrir el Editor de Visual Basic.
2. Importa `Modulos/JsonConverter.bas` y `Modulos/AbrirFormulario.bas` (*Archivo → Importar archivo...*).
3. Copia el contenido de [`userForm`](./userForm) dentro del código de `UserForm1`.

## Requisitos
- Microsoft **Excel habilitado para macros** (`.xlsm`).
- **Activar macros** en el Centro de Confianza de Excel:
  1. Selecciona *Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza*.
  2. En *Configuración de macros*, elige *Habilitar macros* y confirma con *Aceptar*.
- Referencia `Microsoft Scripting Runtime` (`Scripting.Dictionary`) habilitada en *Herramientas → Referencias* del editor VBA (ya configurada en `Test-a11y.xlsm`).

## Funcionalidades principales (v0.2.0)
- **Carga automática inteligente (Local → Nube → Selector manual):** Detecta `es.json` en la carpeta del libro (`ThisWorkbook.Path`), con respaldo automático desde GitHub Raw (`MSXML2.XMLHTTP`) o selector de archivo (`GetOpenFilename`).
- **Caché en memoria:** Parsea el JSON una sola vez por sesión para que abrir el buscador sea instantáneo.
- **Búsqueda multi-palabra y por ID de regla:** Filtra simultáneamente por ID de regla (`clave`), descripción (`description` / `pass` / `fail`) y ayuda corta (`help`).
- **Inserción ágil:** Inserta en la celda activa mediante el botón **Insertar** o con **doble clic** sobre cualquier elemento de la lista.
- **Catálogo completo:** Incluye las **105 reglas** y **80 checks** en español de `axe-core`.

Consulta el historial detallado de mejoras en [`CHANGELOG.md`](./CHANGELOG.md).

## Desarrollo
- **Lenguaje:** VBA (Visual Basic for Applications)
- **Interfaz:** `UserForm1` en Excel
- **Parser JSON:** `VBA-JSON` (`Modulos/JsonConverter.bas`)
- **Entrada de datos:** [`es.json`](./es.json) (local o remoto)

## Licencia

Este proyecto está licenciado bajo los términos de la **GNU GPL v3**. Si usas este módulo para Excel (total o parcialmente) en otros proyectos, por favor brinda atribución al autor original: **Diego Javier Mena Amado**.  
Puedes hacerlo mediante una mención en el README, comentarios en el código, o en la documentación del sistema o archivo que lo utilice.
