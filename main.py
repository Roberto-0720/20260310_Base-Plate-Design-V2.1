"""
Base Plate Design - Main Entry Point
=====================================
Run this file to start the application:
    python main.py
"""

import tkinter as tk
import sys
import os

# Ensure the project root is on the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from baseplate_design.app import BasePlateApp


def main():
    root = tk.Tk()
    app = BasePlateApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
