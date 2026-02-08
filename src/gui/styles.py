"""
Styles and theme configuration for the GUI
"""

from tkinter import ttk


def apply_theme(root):
    """
    Apply custom theme and styles to the application
    
    :param root: The root Tk window
    """
    style = ttk.Style(root)
    
    # Use a modern theme as base
    available_themes = style.theme_names()
    if 'clam' in available_themes:
        style.theme_use('clam')
    elif 'alt' in available_themes:
        style.theme_use('alt')
    
    # Configure button styles
    style.configure(
        'TButton',
        padding=6,
        relief="flat",
        background="#0078D4",
        foreground="white"
    )
    
    style.map(
        'TButton',
        background=[('active', '#005A9E')],
        relief=[('pressed', 'flat'), ('!pressed', 'flat')]
    )
    
    # Configure entry styles
    style.configure(
        'TEntry',
        padding=5,
        relief="solid",
        borderwidth=1
    )
    
    # Configure label styles
    style.configure(
        'TLabel',
        background='#f0f0f0'
    )
    
    # Configure frame styles
    style.configure(
        'TFrame',
        background='#f0f0f0'
    )
    
    return style


# Color scheme
COLORS = {
    'primary': '#0078D4',
    'primary_dark': '#005A9E',
    'success': '#107C10',
    'error': '#E81123',
    'warning': '#FF8C00',
    'background': '#f0f0f0',
    'text': '#000000',
    'text_secondary': '#605E5C'
}


# Font configurations
FONTS = {
    'title': ('Segoe UI', 16, 'bold'),
    'heading': ('Segoe UI', 12, 'bold'),
    'normal': ('Segoe UI', 10),
    'small': ('Segoe UI', 9),
    'mono': ('Courier New', 9)
}