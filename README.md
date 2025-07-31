# AutoDoc - Generador de Documentos Automático

## Descripción
AutoDoc es una aplicación que automatiza la generación de documentos Excel basándose en una plantilla y datos de entrada.

## Funcionalidades
- Crea automáticamente una carpeta "AutoDoc" en el escritorio
- Copia la plantilla Excel necesaria
- Solicita parámetros de configuración de forma interactiva
- Procesa múltiples hojas de Excel
- Genera documentos individuales para cada registro

## Instalación y Uso

### Opción 1: Ejecutar desde código fuente
1. Asegúrate de tener Python 3.7+ instalado
2. Instala las dependencias:
   ```
   pip install -r requirements.txt
   ```
3. Ejecuta el programa:
   ```
   python app.py
   ```

### Opción 2: Usar el ejecutable
1. Ejecuta `AutoDoc.exe` directamente
2. No requiere instalación de Python

## Generar el Ejecutable
Para crear el archivo .exe:
```
python build_exe.py
```

El ejecutable se creará en la carpeta `dist/`.

## Uso del Programa

1. **Al ejecutar el programa:**
   - Se creará automáticamente una carpeta "AutoDoc" en el escritorio
   - Se copiará la plantilla Excel necesaria

2. **Configuración requerida:**
   - Nombre del archivo Excel de entrada
   - Nombres de las hojas a procesar (separadas por comas)
   - Rango de filas para cada hoja (fila inicio y fin)

3. **Proceso:**
   - El programa procesará cada fila en el rango especificado
   - Generará un archivo Excel por cada registro válido
   - Los archivos se guardarán en subcarpetas por hoja

## Estructura de Datos Esperada
El archivo Excel de entrada debe tener las siguientes columnas:
- Columna A: Fecha
- Columna B: Modelo
- Columna D: Número de chasis
- Columna E: Número de motor

## Archivos Generados
- Los archivos se guardan en la carpeta "AutoDoc" del escritorio
- Se crea una subcarpeta por cada hoja procesada
- Cada archivo se nombra con el número de chasis correspondiente

## Requisitos del Sistema
- Windows 10 o superior
- Para el ejecutable: No requiere Python instalado
- Para el código fuente: Python 3.7+ con las dependencias listadas 