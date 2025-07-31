# Configuración opcional para AutoDoc
# Puedes modificar estos valores según tus necesidades

# Configuración por defecto (opcional)
DEFAULT_CONFIG = {
    'input_filename': '',  # Déjalo vacío para que se solicite
    'sheet_names': [],     # Déjalo vacío para que se solicite
    'row_ranges': {},      # Déjalo vacío para que se solicite
}

# Configuración de rangos de valores aleatorios
RANDOM_RANGES = {
    'D22': (30, 40),      # Rango para D22
    'D23': (90, 120),     # Rango para D23
    'D24': (5, 10),       # Rango para D24
    'D25': (5, 10),       # Rango para D25
    'D26': (0, 3),        # Rango para D26
    'D27': (0, 3),        # Rango para D27
    'D28': (80, 100),     # Rango para D28
    'M12': (70, 90),      # Rango para M12
    'G18': (250, 260),    # Rango para G18
    'J18': (180, 200),    # Rango para J18
    'N18': (100, 140),    # Rango para N18
    'Q18': (100, 140),    # Rango para Q18
}

# Configuración de columnas de datos
DATA_COLUMNS = {
    'fecha': 0,      # Columna A
    'modelo': 1,     # Columna B
    'nrchasis': 3,   # Columna D
    'nrmotor': 4,    # Columna E
}

# Configuración de celdas de salida
OUTPUT_CELLS = {
    'fecha': 'B2',
    'nrchasis_1': 'N2',
    'nrchasis_2': 'E3',
    'nrchasis_3': 'E5',
    'nrmotor': 'L5',
    'modelo': 'S4',
} 