"""
Main entry point for the Engineer Syndicate Lookup application
"""

import tkinter as tk
import sys
from pathlib import Path

# Add src directory to path
src_path = Path(__file__).parent
sys.path.insert(0, str(src_path))

from gui.app_window import AppWindow
from gui.styles import apply_theme


def main():
    """
    Launch the application
    """
    root = tk.Tk()
    
    # Apply theme
    apply_theme(root)
    
    # Create and run the application
    app = AppWindow(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Start the application
    root.mainloop()


if __name__ == "__main__":
    main()