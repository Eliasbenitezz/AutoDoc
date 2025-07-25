import warnings
warnings.filterwarnings("ignore", category=UserWarning)
import openpyxl
import random
import os

input_filename = 'cargaV2.xlsx' # Archivo de datos de entrada
sheet_input_name = "DIC24" # Pedir al usuario el nombre de la hoja de entrada
plantilla_filename = 'Plantilla.xlsx'   # Archivo de plantilla base

# Crear carpeta de salida con el nombre de la hoja
output_dir = sheet_input_name
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Cargamos el libro de entrada 
try:
    workbook_input = openpyxl.load_workbook(input_filename, data_only=True) 
    if sheet_input_name:
        if sheet_input_name in workbook_input.sheetnames:
            sheet_input = workbook_input[sheet_input_name]
            print(f"Hoja '{sheet_input_name}' seleccionada correctamente.")
        else:
            print(f"Error: La hoja '{sheet_input_name}' no existe en el archivo '{input_filename}'.")
            print(f"Las hojas disponibles son: {workbook_input.sheetnames}")
            exit()
    else:
        sheet_input = workbook_input.active 
        print(f"Hoja activa seleccionada por defecto: '{sheet_input.title}'")
    print(f"Archivo '{input_filename}' cargado correctamente.")
except FileNotFoundError:
    print(f"Error: El archivo '{input_filename}' no se encontró.")
    print("Asegúrate de que el archivo esté en el mismo directorio que el script o proporciona la ruta completa.")
    exit() # Salir del script si el archivo no existe

for row_index, row in enumerate(sheet_input.iter_rows(min_row=2), start=2):
    fecha = row[0].value
    modelo = row[1].value
    nrchasis = row[3].value 
    nrmotor = row[4].value 

    # Verificar si tenemos nombre y apellido para crear el archivo
    if modelo and nrchasis:
        print(f"\nProcesando fila {row_index}: {modelo} {nrchasis}")

        # Cargar la plantilla para cada alumno
        try:
            workbook_output = openpyxl.load_workbook(plantilla_filename)
            sheet_output = workbook_output.active
        except FileNotFoundError:
            print(f"Error: La plantilla '{plantilla_filename}' no se encontró.")
            exit()

        # Escribir los datos en las celdas especificadas
        sheet_output['B2'] = fecha
        sheet_output['U2'] = nrchasis 
        sheet_output['E3'] = nrchasis
        sheet_output['M5'] = nrmotor
        sheet_output['U4'] = modelo
        sheet_output['D22'] = random.randint(30, 40)
        sheet_output['D23'] = random.randint(90,120)
        sheet_output['D24'] = random.randint(5,10)
        sheet_output['D25'] = random.randint(5,10)
        sheet_output['D26'] = random.randint(0,3)
        sheet_output['D27'] = random.randint(0,3)
        sheet_output['D28'] = random.randint(80,100)
        sheet_output['N12'] = random.randint(70,90)

        # Crear nombre de salida para el archivo
        safe_nombre = str(nrchasis)
        output_filename = f"{nrchasis}.xlsx"
        output_filepath = os.path.join(output_dir, output_filename)

        # Guardar el nuevo libro de trabajo 
        try:
            workbook_output.save(output_filepath)
            print(f"Archivo '{output_filepath}' creado y guardado para {nrchasis}.")
        except Exception as e:
            print(f"Error al guardar el archivo '{output_filepath}': {e}")
    else:
        print(f"\nAdvertencia: Fila {row_index} omitida. Nrchasis faltante.")

print("\n--- Proceso completado ---")