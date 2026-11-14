"""Informações compartilhadas pelo AutoPlan Hub."""
from pathlib import Path
import tkinter as tk

APP_NAME = 'AutoPlan Hub'
APP_VERSION = '0.5.0'
APP_AUTHOR = 'Leonardo Aguiar'
REVIEW_NOTICE = (
    'Revise cuidadosamente a planilha gerada ou atualizada antes de entregar, '
    'enviar ou utilizar como versão oficial. Confira se não faltam dados e se '
    'nenhuma informação precisa ser corrigida.'
)
APP_ROOT = Path(__file__).resolve().parent
ICON_PNG = APP_ROOT / 'assets' / 'autoplan-hub-icon.png'
ICON_ICO = APP_ROOT / 'assets' / 'autoplan-hub.ico'
NOTIFICATION_SOUND = APP_ROOT / 'sounds' / '872716__heyheytheree__notify-sci-fi-ui-sfx-free.wav'


def play_notification_sound():
    """Reproduz o aviso de conclusão sem bloquear a interface."""
    try:
        import winsound
        if not NOTIFICATION_SOUND.is_file():
            return False
        winsound.PlaySound(
            str(NOTIFICATION_SOUND),
            winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT,
        )
        return True
    except Exception:
        return False


def center_window(window, width, height):
    """Centraliza uma janela na tela principal."""
    window.update_idletasks()
    x = max(0, (window.winfo_screenwidth() - width) // 2)
    y = max(0, (window.winfo_screenheight() - height) // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


def apply_window_icon(window):
    """Aplica o ícone na janela e na barra de tarefas do Windows."""
    try:
        image = tk.PhotoImage(file=str(ICON_PNG))
        window.iconphoto(True, image)
        window._autoplan_icon = image
    except Exception:
        pass
    try:
        window.iconbitmap(str(ICON_ICO))
    except Exception:
        pass

