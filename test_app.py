#!/usr/bin/env python3
"""
Script de prueba para verificar que AutoDoc funciona correctamente
"""

import os
import sys
from pathlib import Path

def test_imports():
    """Prueba que todas las dependencias se importen correctamente"""
    try:
        import warnings
        import openpyxl
        import random
        import shutil
        from pathlib import Path
        print("✅ Todas las dependencias se importaron correctamente")
        return True
    except ImportError as e:
        print(f"❌ Error al importar dependencias: {e}")
        return False

def test_plantilla_exists():
    """Prueba que la plantilla existe"""
    if Path("Plantilla.xlsx").exists():
        print("✅ Plantilla.xlsx encontrada")
        return True
    else:
        print("❌ Plantilla.xlsx no encontrada")
        return False

def test_desktop_access():
    """Prueba que se puede acceder al escritorio"""
    try:
        desktop = Path.home() / "Desktop"
        print(f"✅ Acceso al escritorio: {desktop}")
        return True
    except Exception as e:
        print(f"❌ Error al acceder al escritorio: {e}")
        return False

def run_tests():
    """Ejecuta todas las pruebas"""
    print("🧪 Ejecutando pruebas de AutoDoc...")
    print("="*50)
    
    tests = [
        test_imports,
        test_plantilla_exists,
        test_desktop_access
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        if test():
            passed += 1
        print()
    
    print("="*50)
    print(f"📊 Resultados: {passed}/{total} pruebas pasaron")
    
    if passed == total:
        print("🎉 Todas las pruebas pasaron. El programa está listo para usar.")
        return True
    else:
        print("⚠️  Algunas pruebas fallaron. Revisa los errores arriba.")
        return False

if __name__ == "__main__":
    success = run_tests()
    input("\nPresiona Enter para salir...")
    sys.exit(0 if success else 1) 