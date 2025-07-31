# INSTRUCCIONES PARA USAR AUTODOC.EXE

## 🚀 Cómo usar el programa AutoDoc

### Paso 1: Preparación
1. Copia el archivo `AutoDoc.exe` a cualquier carpeta de tu computadora
2. Asegúrate de que el archivo `Plantilla.xlsx` esté en la misma carpeta que `AutoDoc.exe`
3. Coloca tu archivo Excel de datos en la misma carpeta
4. **IMPORTANTE**: El programa creará automáticamente una carpeta "AutoDoc" en el escritorio (o en el directorio actual si no encuentra el escritorio)

### Paso 2: Ejecutar el programa
1. Haz doble clic en `AutoDoc.exe`
2. Se abrirá una ventana de terminal (consola)

### Paso 3: Configuración
El programa te pedirá la siguiente información:

#### 1. Nombre del archivo Excel
- Ingresa el nombre de tu archivo Excel (ej: `datos.xlsx`)
- Si no escribes la extensión `.xlsx`, el programa la agregará automáticamente

#### 2. Nombres de las hojas
- Escribe los nombres de las hojas que quieres procesar
- Si tienes múltiples hojas, sepáralas con comas
- Ejemplo: `Hoja1, Hoja2, PLANTA`

#### 3. Rango de filas
- Para cada hoja, te pedirá:
  - **Fila de inicio**: El número de la primera fila a procesar
  - **Fila de fin**: El número de la última fila a procesar
- Ejemplo: Si quieres procesar de la fila 10 a la 50, ingresa:
  - Fila de inicio: `10`
  - Fila de fin: `50`

### Paso 4: Procesamiento
- El programa creará automáticamente una carpeta llamada "AutoDoc" en tu escritorio (o en el directorio actual)
- Copiará la plantilla Excel y el archivo de datos a esa carpeta
- Procesará cada fila del rango especificado
- Generará un archivo Excel por cada registro válido
- Creará subcarpetas por cada hoja procesada

### Paso 5: Resultados
- Los archivos generados estarán en la carpeta "AutoDoc" del escritorio (o directorio actual)
- Se creará una subcarpeta por cada hoja procesada
- Cada archivo se nombrará con el número de chasis correspondiente
- Los archivos se organizarán en subcarpetas por nombre de hoja

## 📋 Estructura de datos requerida

Tu archivo Excel debe tener las siguientes columnas:
- **Columna A**: Fecha
- **Columna B**: Modelo
- **Columna D**: Número de chasis
- **Columna E**: Número de motor

## ⚠️ Notas importantes

1. **Archivos necesarios**: Asegúrate de tener tanto `AutoDoc.exe` como `Plantilla.xlsx` en la misma carpeta
2. **Formato de datos**: Los datos deben estar en las columnas correctas (A, B, D, E)
3. **Rangos válidos**: Las filas de inicio y fin deben ser números positivos, y la fila de inicio debe ser menor o igual a la fila de fin
4. **Nombres de hojas**: Los nombres deben coincidir exactamente con los que están en tu archivo Excel

## 🔧 Solución de problemas

### Error: "No se encontró Plantilla.xlsx"
- Asegúrate de que `Plantilla.xlsx` esté en la misma carpeta que `AutoDoc.exe`

### Error: "El archivo no se encontró"
- Verifica que el nombre del archivo Excel sea correcto
- Asegúrate de que el archivo esté en la misma carpeta

### Error: "La hoja no existe"
- Verifica que el nombre de la hoja coincida exactamente con el que está en tu archivo Excel
- Los nombres son sensibles a mayúsculas y minúsculas
- Usa el script `verificar_hojas.py` para ver las hojas disponibles

### Error: "No se pudo encontrar el escritorio"
- El programa usará automáticamente el directorio actual como alternativa
- Esto es normal y no afecta el funcionamiento

### El programa se cierra inmediatamente
- Ejecuta `AutoDoc.exe` desde una terminal para ver los mensajes de error
- Verifica que todos los archivos necesarios estén presentes

### Problemas de codificación
- El programa está optimizado para funcionar en Windows
- Si ves caracteres extraños, es normal y no afecta el funcionamiento

## 📞 Soporte

Si tienes problemas, verifica:
1. Que todos los archivos estén en la misma carpeta.
2. Que los nombres de archivos y hojas sean correctos.
3. Que los rangos de filas sean válidos.
4. Que la estructura de datos sea la correcta. 
5. Ultima opción contacta la desarrollador.