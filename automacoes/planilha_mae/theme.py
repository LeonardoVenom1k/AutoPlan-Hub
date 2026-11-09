"""Identidade visual do Consolidador de Planilha Mãe."""
from tkinter import ttk


COLORS = {
    'navy': '#17365D',
    'blue': '#2F75B5',
    'blue_hover': '#245E91',
    'background': '#EEF3F8',
    'surface': '#FFFFFF',
    'border': '#CAD6E2',
    'text': '#1F2937',
    'muted': '#64748B',
    'success': '#18864B',
    'warning': '#A16207',
    'danger': '#B42318',
    'log': '#152235',
}


def configure_styles(root):
    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except Exception:
        pass
    style.configure('TFrame', background=COLORS['background'])
    style.configure('TLabel', background=COLORS['background'], foreground=COLORS['text'], font=('Segoe UI', 10))
    style.configure('Card.TFrame', background=COLORS['surface'], relief='solid', borderwidth=1)
    style.configure('Card.TLabel', background=COLORS['surface'], foreground=COLORS['text'], font=('Segoe UI', 10))
    style.configure('CardTitle.TLabel', background=COLORS['surface'], foreground=COLORS['navy'], font=('Segoe UI', 12, 'bold'))
    style.configure('Muted.TLabel', background=COLORS['surface'], foreground=COLORS['muted'], font=('Segoe UI', 9))
    style.configure('TEntry', padding=7)
    style.configure('TCombobox', padding=6)
    style.configure('TButton', font=('Segoe UI', 10), padding=(11, 7))
    style.configure('Primary.TButton', background=COLORS['blue'], foreground='white', font=('Segoe UI', 11, 'bold'), padding=(18, 10))
    style.map('Primary.TButton', background=[('active', COLORS['blue_hover']), ('disabled', '#9CB6CF')])
    style.configure('Mode.TRadiobutton', background=COLORS['surface'], foreground=COLORS['text'], font=('Segoe UI', 10, 'bold'), padding=7)
    style.map('Mode.TRadiobutton', background=[('active', '#E8F1FA')])
    style.configure('Horizontal.TProgressbar', troughcolor='#DCE5EF', background=COLORS['blue'], thickness=12)
    return style

