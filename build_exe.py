import subprocess
import sys
import os

def build_exe():
    """Construye el ejecutable usando PyInstaller"""
    print("Iniciando construcción del ejecutable...")
    
    # Comando para PyInstaller
    cmd = [
        "pyinstaller",
        "--onefile",  # Crear un solo archivo ejecutable
        "--console",  # Incluir consola
        "--name=AutoDoc",  # Nombre del ejecutable
        "--add-data=Plantilla.xlsx;.",  # Incluir la plantilla
        "--exclude-module=tkinter",  # Excluir tkinter
        "--exclude-module=_tkinter",  # Excluir _tkinter
        "--exclude-module=tk",  # Excluir tk
        "--exclude-module=Tkinter",  # Excluir Tkinter
        "--icon=NONE",  # Sin icono por ahora
        "app.py"
    ]
    
    try:
        # Ejecutar PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("Ejecutable creado exitosamente!")
        print("El archivo 'AutoDoc.exe' se encuentra en la carpeta 'dist'")
        
    except subprocess.CalledProcessError as e:
        print(f"Error al crear el ejecutable: {e}")
        print(f"Error output: {e.stderr}")
        return False
    except FileNotFoundError:
        print("Error: PyInstaller no está instalado. Instalando...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
            print("PyInstaller instalado. Intentando crear el ejecutable nuevamente...")
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)
            print("Ejecutable creado exitosamente!")
            print("El archivo 'AutoDoc.exe' se encuentra en la carpeta 'dist'")
        except subprocess.CalledProcessError as e:
            print(f"Error al instalar PyInstaller: {e}")
            return False
    
    return True

if __name__ == "__main__":
    build_exe() 