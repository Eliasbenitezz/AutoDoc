import openpyxl

input_filename = 'Hoja de entrada.xlsx' #Acá tenes que especificar el nombre del archivo o copiar la ruta entera

# --- Cargar el libro de trabajo de entrada ---
try:
    workbook_input = openpyxl.load_workbook(input_filename) 
    sheet_input = workbook_input.active 
    print(f"Archivo '{input_filename}' cargado correctamente.")
except FileNotFoundError:
    print(f"Error: El archivo '{input_filename}' no se encontró.")
    print("Asegúrate de que el archivo esté en el mismo directorio que el script o proporciona la ruta completa.")
    exit() # Salir del script si el archivo no existe

'''Procesamos los datos desde la segunda fila, teniendo en cuenta que la primera
fila es el encabezado que indentifica el tipo de dato de esa columna'''
for row_index, row in enumerate(sheet_input.iter_rows(min_row=2), start=2):
    ''''Obtenemos los datos de las celdas (entendiendo que el orden Nombre, Apellido, Curso, Nota es A, B, C, D)
    Usamos .value para obtener el contenido de cada celda'''
    nombre = row[0].value
    apellido = row[1].value
    curso = row[2].value
    nota = row[3].value

    # Verificar si tenemos nombre y apellido para crear el archivo
    if nombre and apellido:
        print(f"\nProcesando fila {row_index}: {nombre} {apellido}")

        #Cramos un libro unico para el alumno
        workbook_output = openpyxl.Workbook()
        sheet_output = workbook_output.active
        sheet_output.title = "Datos Alumno" # Acá podes cambiar el nombre de la hoja si queres

        # --- Escribir los datos en las celdas especificadas ---
        # Puedes cambiar las celdas de destino aquí según necesites
        # Ejemplo: Escribir etiquetas en la columna A y valores en la B
        sheet_output['B1'] = nombre

        sheet_output['B2'] = apellido

        sheet_output['B3'] = curso

        sheet_output['B4'] = nota

        '''Creamos un nombre de salida para el archivo 
        remplazamos los espacios en blanco con _ para mejor lectura'''
        safe_nombre = str(nombre).replace(' ', '_')
        safe_apellido = str(apellido).replace(' ', '_')
        output_filename = f"{safe_nombre}_{safe_apellido}.xlsx"

        output_filepath = output_filename

        # --- Guardar el nuevo libro de trabajo ---
        try:
            workbook_output.save(output_filepath)
            print(f"Archivo '{output_filepath}' creado y guardado para {nombre} {apellido}.")
        except Exception as e:
            print(f"Error al guardar el archivo '{output_filepath}': {e}")
    else:
        print(f"\nAdvertencia: Fila {row_index} omitida. Nombre o Apellido faltante.")
        # Opcionalmente, puedes decidir qué hacer si falta nombre/apellido
        # Podrías usar un ID, el índice de fila, o simplemente omitirla como ahora.

print("\n--- Proceso completado ---")