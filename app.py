import warnings
warnings.filterwarnings("ignore", category=UserWarning)
import openpyxl
import random
import os

input_filename = 'cargaV2.xlsx' # Archivo de datos de entrada

sheet_input_name = ["PLANTA"] #Agrega las hojas que quieras procesar, entre comillas y separadas por comas

plantilla_filename = 'Plantilla.xlsx'  #Archivo de plantilla que se usará para crear los archivos de salida

try:
    workbook_input = openpyxl.load_workbook(input_filename, data_only=True) 
    print(f"Archivo '{input_filename}' cargado correctamente.")
except FileNotFoundError:
    print(f"Error: El archivo '{input_filename}' no se encontró.")
    print("Asegúrate de que el archivo esté en el mismo directorio que el script o proporciona la ruta completa.")
    exit()

for sheet_name in sheet_input_name:
    print(f"\n{'='*50}")
    print(f"Procesando hoja: {sheet_name}")
    print(f"{'='*50}")
    
    if sheet_name not in workbook_input.sheetnames:
        print(f"Error: La hoja '{sheet_name}' no existe en el archivo '{input_filename}'.")
        print(f"Saltando a la siguiente hoja...")
        continue
    
    sheet_input = workbook_input[sheet_name]
    print(f"Hoja '{sheet_name}' seleccionada correctamente.")
    
    output_dir = sheet_name
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Carpeta '{output_dir}' creada.")
    
    archivos_procesados = 0
    
    for row_index, row in enumerate(sheet_input.iter_rows(min_row=2), start=2):
        fecha = row[0].value
        modelo = row[1].value
        nrchasis = row[3].value 
        nrmotor = row[4].value 

        if modelo and nrchasis:
            print(f"Procesando fila {row_index}: {modelo} {nrchasis}")

            try:
                workbook_output = openpyxl.load_workbook(plantilla_filename)
                sheet_output = workbook_output.active
            except FileNotFoundError:
                print(f"Error: La plantilla '{plantilla_filename}' no se encontró.")
                exit()

            sheet_output['B2'] = fecha
            sheet_output['N2'] = nrchasis 
            sheet_output['E3'] = nrchasis
            sheet_output['E5'] = nrchasis
            sheet_output['L5'] = nrmotor
            sheet_output['S4'] = modelo
            sheet_output['D22'] = random.randint(30, 40)
            sheet_output['D23'] = random.randint(90,120)
            sheet_output['D24'] = random.randint(5,10)
            sheet_output['D25'] = random.randint(5,10)
            sheet_output['D26'] = random.randint(0,3)
            sheet_output['D27'] = random.randint(0,3)
            sheet_output['D28'] = random.randint(80,100)
            sheet_output['M12'] = random.randint(70,90) 
            sheet_output['G18'] = random.randint(250,260)
            sheet_output['J18'] = random.randint(180,200)
            sheet_output['N18'] = random.randint(100,140)
            sheet_output['Q18'] = random.randint(100,140) 

            safe_nombre = str(nrchasis)
            output_filename = f"{nrchasis}.xlsx"
            output_filepath = os.path.join(output_dir, output_filename)

            try:
                workbook_output.save(output_filepath)
                print(f"Archivo '{output_filepath}' creado y guardado para {nrchasis}.")
                archivos_procesados += 1
            except Exception as e:
                print(f"Error al guardar el archivo '{output_filepath}': {e}")
        else:
            print(f"Advertencia: Fila {row_index} omitida. Nrchasis faltante.")
    
    print(f"\nHoja '{sheet_name}' completada. Archivos procesados: {archivos_procesados}")

print(f"\n{'='*50}")
print("--- Proceso completado para todas las hojas ---")
print(f"{'='*50}") 