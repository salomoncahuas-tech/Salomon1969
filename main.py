#!/usr/bin/env python3
"""
IN Piura - Plan de Verificación de Campo
Restauración de Ecosistemas - Cuenca Alta del Río Piura, Perú.

Punto de entrada de la aplicación de escritorio.
Ejecutar: python main.py
"""

import sys
import os

# Asegurar que el directorio del script esté en el path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import AplicacionINPiura


def main():
    aplicacion = AplicacionINPiura()
    aplicacion.mainloop()


if __name__ == "__main__":
    main()
