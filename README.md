# 📌 A11y-Help: Asistente para Auditoría de Accesibilidad en Excel

## 📖 Descripción
**A11y-Help** es una herramienta en **Excel con macros (VBA)** diseñada para ayudar en auditorías de accesibilidad digital. Permite buscar y autocompletar descripciones de reglas de accesibilidad a partir de un JSON, facilitando la documentación y el análisis en evaluaciones de accesibilidad web.

## 💻 Uso
1. **Descromprimir carpeta descargada**
1. **Abrir** el archivo Excel.
2. Presionar el botón para abrir el formulario de búsqueda.
3. Escribir cualquier palabra clave
4. Darle a la opcion depurar
5. ubicar al costado izquierdo "userForm1"
6. Doble clic sobre el formulario
7. Ubicar la línea jsonFile y cambiar la ruta por la de la carpeta descomprimida donde se encuentra el archivo json que tiene todas las reglas de accesibilidad.

## 🔧 Requisitos
- Microsoft **Excel habilitado para macros** (`.xlsm`)
- **Activar macros** en el Centro de Confianza de Excel
    - Para activar macros en el Centro de Confianza de Excel, sigue estos pasos: 
        - Selecciona "Archivo" -> "Opciones".
        - Selecciona "Centro de confianza" -> "Configuración del Centro de confianza".
        - En el lado izquierdo, selecciona "Configuración de macros".
        - Elige la opción: "Habilitar todas las macros".
        - Confirma los cambios haciendo clic en "Aceptar".
- **Activar pestaña programador** en Excel
    - Para activar la pestaña programador, sigue estos pasos: 
        - Selecciona "Archivo" -> "Opciones".
        - Selecciona "Personalizar cinta de opciones" 
        - En la lista de pestañas principales a la derecha, busca la casilla de verificación para "Programador" o "Desarrollador" y márcala.
        - Confirma los cambios haciendo clic en "Aceptar".
- Archivo **JSON** con reglas de accesibilidad


## 🚀 Funcionalidades
- 📂 **Carga automática de reglas desde JSON**
- 🔎 **Búsqueda en tiempo real** con autocompletado
- 📋 **Inserción de descripciones en celdas de Excel**
- 🎨 **Interfaz de usuario intuitiva con formulario en VBA**

## 🛠️ Desarrollo
- **Lenguaje:** VBA (Visual Basic for Applications)
- **Interfaz:** UserForm en Excel
- **Entrada de datos:** Archivo JSON

## 📌 Próximos pasos
- 🔄 Optimización del procesamiento del JSON
- Llamado del archivo JSON desde un lugar en la nube para evitar el ajuste del jsonFile en el código
- Añadir más opciones de búsqueda y autocompletado
- Mejorar la interfaz de usuario para una experiencia más intuitiva y agradable
- Agregar más funcionalidades y características para mejorar el rendimiento y la eficiencia del proceso de auditoría de accesibilidad en Excel.

## Licencia

Este proyecto está licenciado bajo los términos de la **GNU GPL v3**. Si usas este modulo para Excel (total o parcialmente) en otros proyectos, por favor brinda atribución al autor original: **Diego Javier Mena Amado**  
Puedes hacerlo mediante una mención en el README, comentarios en el código, o en la documentación del sistema o archivo que lo utilice.
