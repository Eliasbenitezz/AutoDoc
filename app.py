import openpyxl

input_filename = 'Hoja de entrada.xlsx' # Archivo de datos de entrada
plantilla_filename = 'Plantilla.xlsx'   # Archivo de plantilla base

# Cargamos el libro de entrada 
try:
    workbook_input = openpyxl.load_workbook(input_filename) 
    sheet_input = workbook_input.active 
    print(f"Archivo '{input_filename}' cargado correctamente.")
except FileNotFoundError:
    print(f"Error: El archivo '{input_filename}' no se encontró.")
    print("Asegúrate de que el archivo esté en el mismo directorio que el script o proporciona la ruta completa.")
    exit() # Salir del script si el archivo no existe

for row_index, row in enumerate(sheet_input.iter_rows(min_row=2), start=2):
    nombre = row[0].value
    apellido = row[1].value
    curso = row[2].value
    nota = row[3].value

    # Verificar si tenemos nombre y apellido para crear el archivo
    if nombre and apellido:
        print(f"\nProcesando fila {row_index}: {nombre} {apellido}")

        # Cargar la plantilla para cada alumno
        try:
            workbook_output = openpyxl.load_workbook(plantilla_filename)
            sheet_output = workbook_output.active
        except FileNotFoundError:
            print(f"Error: La plantilla '{plantilla_filename}' no se encontró.")
            exit()

        # Escribir los datos en las celdas especificadas
        sheet_output['B1'] = nombre
        sheet_output['B2'] = apellido
        sheet_output['B3'] = curso
        sheet_output['B4'] = nota

        # Crear nombre de salida para el archivo
        safe_nombre = str(nombre).replace(' ', '_')
        safe_apellido = str(apellido).replace(' ', '_')
        output_filename = f"{safe_nombre}_{safe_apellido}.xlsx"
        output_filepath = output_filename

        # Guardar el nuevo libro de trabajo
        try:
            workbook_output.save(output_filepath)
            print(f"Archivo '{output_filepath}' creado y guardado para {nombre} {apellido}.")
        except Exception as e:
            print(f"Error al guardar el archivo '{output_filepath}': {e}")
    else:
        print(f"\nAdvertencia: Fila {row_index} omitida. Nombre o Apellido faltante.")

print("\n--- Proceso completado ---")