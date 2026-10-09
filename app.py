"""Ponto de entrada — duplo clique abre o app, sem instalar."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.app import main

if __name__ == "__main__":
    main()
