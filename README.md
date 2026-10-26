# Complemento de Excel: Morfo - Limpieza y Transformación de Texto (.xlam)

Complemento profesional nativo para **Microsoft Excel** diseñado para analistas de datos, contadores, administradores de bases de datos y usuarios que trabajan frecuentemente con datos exportados de ERPs, CRMs o páginas web.

---

## 🌟 Características Principales

### 1. Transformación de Formato de Texto
- **MAYÚSCULAS**: Convierte el texto seleccionado a mayúsculas completas.
- **minúsculas**: Convierte el texto seleccionado a minúsculas.
- **Nombre Propio**: Convierte cada palabra a formato título (capitalización inicial de cada palabra), respetando acentos y caracteres especiales en español.
- **Tipo Oración**: Convierte únicamente la primera letra de una palabra, oración o grupo de palabras a mayúscula y el resto a minúsculas.

### 2. Limpieza de Bases de Datos
- **Quitar Espacios Innecesarios**:
  - Elimina espacios sobrantes al inicio y al final (`Trim`).
  - Reduce secuencias de espacios dobles o múltiples a un único espacio.
  - Elimina **espacios duros / no separables de internet (ASCII 160 / `&nbsp;`)** que la fórmula `ESPACIOS()` de Excel no puede quitar.
- **Quitar Acentos / Tildes**:
  - Reemplaza vocales acentuadas (`á, é, í, ó, ú, ü, à, è, etc.`) por sus equivalentes sin tilde (`a, e, i, o, u`).
  - **Protección de la Ñ**: Por defecto conserva la letra `ñ / Ñ`, con opción en el menú avanzado de convertirla en `n / N` para sistemas legados.
- **Quitar Caracteres Raros / Control**:
  - Elimina códigos de control ASCII invisibles (códigos `0` al `31` y `127`).
  - Elimina caracteres invisibles Unicode (`Zero-Width Space`, BOM, etc.).
  - Reemplaza saltos de línea dentro de la celda por espacios para no romper la exportación en archivos CSV o bases de datos relacionales.
- **Limpieza Total (1-Clic)**: Aplica de forma combinada la eliminación de caracteres raros, tildes y normalización de espacios en una sola pasada.

### 3. Modos de Ejecución Flexibles (Fila Individual o Rangos)
- **Fila Individual o Celda**: Selecciona una celda o fila y aplica la acción directamente.
- **Rango de Filas**: Selecciona un bloque de filas (ej. `A2:D5000`) o columnas completas.
- **Rendimiento Ultrarrápido en Memoria**: Procesamiento mediante matrices (`Arrays`) que permite limpiar decenas de miles de filas en menos de 1 segundo.

---

## 🖥️ Integración con Excel

El complemento se integra en dos lugares clave de la interfaz de Excel:

1. **Cinta de Opciones (Ribbon)**:
   - Se crea una pestaña dedicada llamada **`Morfo`** ubicada por defecto **al final** de la cinta de opciones (para no alterar tus pestañas habituales).
   - Puedes cambiar su ubicación en cualquier momento con el archivo [Configurar_Posicion_Pestana.bat](file:///c:/Users/DELL/Documents/morfo_excel/Configurar_Posicion_Pestana.bat) o arrastrándola desde las opciones de Excel.
   - Botones con iconos intuitivos organizados en:
     - *Formato de Mayúsculas y Minúsculas*
     - *Limpieza de Bases de Datos*
     - *Personalización (Opciones Avanzadas)*

2. **Menú Contextual (Clic Derecho)**:
   - Al hacer clic derecho sobre cualquier celda, fila o rango, aparece el submenú:
     `🧹 Morfo: Limpieza y Texto` con acceso instantáneo a todas las opciones.

3. **Editor de Selección y Párrafos (Transformar fragmentos específicos)**:
   - Permite seleccionar con el ratón **cualquier palabra, frase o fragmento** dentro de un párrafo largo y transformarlo a MAYÚSCULAS, minúsculas, Nombre Propio, sin acentos, etc., sin alterar el resto de la celda.
   - Soporta el atajo **Shift + F3** dentro del editor para alternar mayúsculas y minúsculas (igual que en Microsoft Word).
   - Accesible desde el botón **Editor de Selección** en la cinta, desde el menú de clic derecho (`📝 Editar / Transformar Selección...`), o desde la ventana de Opciones Avanzadas.

4. **Ventana Flotante Interactiva (Opciones Avanzadas con Vista Previa)**:
   - Muestra el contenido de la celda activa y una **vista previa en tiempo real** del resultado antes de aplicar los cambios.
   - Permite activar o desactivar cada regla de limpieza de forma individual mediante casillas de verificación.

5. **Fórmulas Personalizadas (UDF)**:
   También puedes usar las funciones directamente en fórmulas de Excel:
   - `=UDF_MAYUSCULAS(A2)`
   - `=UDF_MINUSCULAS(A2)`
   - `=UDF_NOMBRE_PROPIO(A2)`
   - `=UDF_TIPO_ORACION(A2)`
   - `=UDF_LIMPIAR_ESPACIOS(A2)`
   - `=UDF_QUITAR_ACENTOS(A2; [conservarEnie])`
   - `=UDF_QUITAR_RAROS(A2; [quitarSaltos])`
   - `=UDF_LIMPIEZA_TOTAL(A2)`

---

## 📁 Estructura del Proyecto

```
Ansi_excel/
├── dist/
│   └── Morfo.xlam                  # Complemento compilado listo para usar
├── src/
│   ├── modArrayProcessor.bas       # Motor de procesamiento en memoria para filas y rangos
│   ├── modTextTransform.bas        # Lógica de mayúsculas, minúsculas y nombre propio
│   ├── modDataClean.bas            # Lógica de limpieza de espacios, acentos y raros
│   ├── modRibbonCallbacks.bas      # Callbacks para la Cinta de Opciones
│   ├── modContextMenuFallback.bas  # Menú contextual de clic derecho (CommandBars)
│   ├── modUserFormLauncher.bas     # Lanzador de la ventana flotante
│   ├── modSelfTest.bas             # Pruebas de integración
│   └── frmAdvancedClean.frm        # Formulario flotante con vista previa interactiva
├── customUI/
│   └── customUI14.xml              # Definición RibbonX (Cinta y Menú Contextual XML)
├── scripts/
│   ├── Build-Addin.ps1             # Compila el código fuente en el archivo .xlam
│   ├── Install-Addin.ps1           # Instala y activa el complemento en Excel
│   └── Create-Sample.ps1           # Genera un libro de prueba con datos sucios
├── tests/
│   └── Run-Tests.ps1               # Suite de verificación automatizada
├── Ejemplo_Datos_Prueba.xlsx       # Libro de prueba listo para abrir en Excel
└── README.md                       # Documentación
```

---

## 🚀 Instalación y Uso Rápido (1 Clic)

### Instalación por Doble Clic (Recomendada)
En la carpeta principal del proyecto encontrarás el instalador automático:
👉 **`Instalar_Complemento.bat`** (haz doble clic sobre él)

El instalador realiza todo el trabajo por ti de forma 100% automatizada:
1. Copia el archivo `.xlam` a la carpeta oficial de complementos de Excel (`%APPDATA%\Microsoft\AddIns`).
2. Registra y activa el complemento en la lista oficial de Add-ins de Excel para que funcione en **cualquier libro sin macros (`.xlsx`)**, archivos `.csv` o libros nuevos.
3. Configura las **Ubicaciones de Confianza (Trusted Locations)** en el Registro de Windows de Excel para que las macros nunca se bloqueen.
4. Desbloquea los archivos en Windows (`Unblock-File`).

### Cambiar la posición de la pestaña 'Morfo' en la cinta
Si deseas cambiar el lugar donde aparece la pestaña:
👉 **`Configurar_Posicion_Pestana.bat`** (haz doble clic para elegir: Al final, Después de Inicio, Datos, Revisar, etc.)

O también de forma nativa en Excel:
`Archivo > Opciones > Personalizar cinta de opciones` y mover la pestaña "Morfo" con las flechas Subir / Bajar.

### Personalizar Logos e Iconos de los Botones
Puedes personalizar el logo de cualquiera de las 9 funciones:
👉 **`Configurar_Logos.bat`** (haz doble clic para abrir el asistente interactivo)
- Te permite elegir qué imagen de `customUI\logos` asignar a cada función.
- Te permite abrir el explorador de Windows para buscar y cargar cualquier archivo `.png` de tu equipo.
- O simplemente guarda tus imágenes `.png` en `customUI\logos` con nombres como `mayus.png`, `minus.png`, `proper.png`, etc., y se detectarán solas.

### Desinstalación
Si alguna vez deseas retirar el complemento, simplemente haz doble clic en:
👉 **`Desinstalar_Complemento.bat`**

### ¿Cómo probarlo ahora mismo?
1. Abre el archivo de prueba incluido: [Ejemplo_Datos_Prueba.xlsx](file:///c:/Users/DELL/Documents/Ansi_excel/Ejemplo_Datos_Prueba.xlsx).
2. Verás la pestaña **"Morfo"** en la parte superior.
3. Selecciona una celda individual o un rango de filas (ej. `B5:B8`).
4. Haz clic en **Limpieza Total** o **Nombre Propio** desde la cinta de opciones, o haz **clic derecho** sobre las celdas seleccionadas y elige la opción que desees en el menú.
5. Para abrir la ventana interactiva con vista previa en vivo, haz clic en **Opciones Avanzadas**.
