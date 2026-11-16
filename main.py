# -*- coding: utf-8 -*-
"""Menu principal do AutoPlan Hub."""
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from app_info import APP_AUTHOR, APP_NAME, APP_VERSION, apply_window_icon, center_window
from automacoes.ambiente_trabalho.service import launch_environment, load_config

APP_DIR = Path(__file__).resolve().parent


def python_gui_executable():
    """Prefere pythonw no Windows para não abrir uma janela de terminal."""
    current = Path(sys.executable)
    candidate = current.with_name('pythonw.exe')
    return candidate if candidate.is_file() else current


class AutoPlanHub(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('AutoPlan Hub')
        self.minsize(820, 680)
        self.configure(bg='#EEF3F8')
        apply_window_icon(self)

        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('Hub.TButton', font=('Segoe UI', 12, 'bold'), padding=(18, 14))
        style.configure('Secondary.TButton', font=('Segoe UI', 10), padding=(12, 9))

        header = tk.Frame(self, bg='#17365D', padx=30, pady=24)
        header.pack(fill='x')
        tk.Label(header, text='AutoPlan Hub', bg='#17365D', fg='white',
                 font=('Segoe UI', 25, 'bold')).pack(anchor='w')
        tk.Label(header, text='Central de automações para rotinas operacionais',
                 bg='#17365D', fg='#D9EAF7', font=('Segoe UI', 11)).pack(anchor='w', pady=(3, 0))

        content = tk.Frame(self, bg='#EEF3F8', padx=30, pady=25)
        content.pack(fill='both', expand=True)
        tk.Label(content, text='Automações disponíveis', bg='#EEF3F8', fg='#1F2937',
                 font=('Segoe UI', 16, 'bold')).pack(anchor='w', pady=(0, 14))

        work_card = tk.Frame(content, bg='white', highlightbackground='#C9D5E3',
                             highlightthickness=1, padx=22, pady=18)
        work_card.pack(fill='x', anchor='n')
        tk.Label(work_card, text='Abrir ambiente de trabalho', bg='white', fg='#17365D',
                 font=('Segoe UI', 15, 'bold')).grid(row=0, column=0, sticky='w')
        tk.Label(work_card,
                 text='Abre os aplicativos, sites e pastas escolhidos e lembra suas preferências.',
                 bg='white', fg='#4B5563', font=('Segoe UI', 10),
                 wraplength=600, justify='left').grid(row=1, column=0, sticky='w', pady=(5, 12))
        work_actions = tk.Frame(work_card, bg='white')
        work_actions.grid(row=0, column=1, rowspan=2, padx=(25, 0))
        ttk.Button(work_actions, text='Abrir', style='Hub.TButton',
                   command=self.launch_work_environment).pack(side='left')
        ttk.Button(work_actions, text='Configurar', style='Secondary.TButton',
                   command=self.open_work_environment).pack(side='left', padx=(8, 0))
        work_card.columnconfigure(0, weight=1)

        card = tk.Frame(content, bg='white', highlightbackground='#C9D5E3',
                        highlightthickness=1, padx=22, pady=20)
        card.pack(fill='x', anchor='n', pady=(14, 0))
        tk.Label(card, text='Consolidador de Planilha Mãe', bg='white', fg='#17365D',
                 font=('Segoe UI', 15, 'bold')).grid(row=0, column=0, sticky='w')
        tk.Label(card,
                 text='Consolida planilhas diárias, reúne inconformidades e atualiza o Painel.',
                 bg='white', fg='#4B5563', font=('Segoe UI', 10),
                 wraplength=600, justify='left').grid(row=1, column=0, sticky='w', pady=(5, 14))
        ttk.Button(card, text='Abrir automação', style='Hub.TButton',
                   command=self.open_consolidator).grid(row=0, column=1, rowspan=2, padx=(25, 0))
        card.columnconfigure(0, weight=1)

        rank_card = tk.Frame(content, bg='white', highlightbackground='#C9D5E3',
                             highlightthickness=1, padx=22, pady=20)
        rank_card.pack(fill='x', anchor='n', pady=(14, 0))
        tk.Label(rank_card, text='Gerador de Planilha Rank', bg='white', fg='#17365D',
                 font=('Segoe UI', 15, 'bold')).grid(row=0, column=0, sticky='w')
        tk.Label(rank_card,
                 text='Gera rankings, ocorrências, análises por cliente e painéis de supervisão.',
                 bg='white', fg='#4B5563', font=('Segoe UI', 10),
                 wraplength=600, justify='left').grid(row=1, column=0, sticky='w', pady=(5, 14))
        ttk.Button(rank_card, text='Abrir automação', style='Hub.TButton',
                   command=self.open_rank).grid(row=0, column=1, rowspan=2, padx=(25, 0))
        rank_card.columnconfigure(0, weight=1)

        footer = tk.Frame(self, bg='#E3EAF2', padx=20, pady=9)
        footer.pack(fill='x', side='bottom')
        tk.Label(footer, text=f'{APP_NAME}  •  Desenvolvido por {APP_AUTHOR}  •  Versão {APP_VERSION}', bg='#E3EAF2', fg='#526273',
                 font=('Segoe UI', 9)).pack(side='left')
        ttk.Button(footer, text='Fechar', style='Secondary.TButton',
                   command=self.destroy).pack(side='right')
        center_window(self, 920, 760)

    def open_work_environment(self):
        module = APP_DIR / 'automacoes' / 'ambiente_trabalho' / 'ui.py'
        if not module.is_file():
            messagebox.showerror('Recurso não encontrado',
                                 'O módulo do Ambiente de Trabalho não foi encontrado.')
            return
        try:
            subprocess.Popen([str(python_gui_executable()), '-m', 'automacoes.ambiente_trabalho.ui'],
                             cwd=str(APP_DIR))
        except Exception as exc:
            messagebox.showerror('Não foi possível abrir', f'Falha ao abrir o ambiente de trabalho:\n{exc}')

    def launch_work_environment(self):
        try:
            successes, failures = launch_environment(load_config())
        except Exception as exc:
            messagebox.showerror('Não foi possível abrir', f'Falha ao abrir o ambiente de trabalho:\n{exc}')
            return
        if failures:
            messagebox.showwarning(
                'Ambiente aberto com avisos',
                'Itens abertos:\n- ' + ('\n- '.join(successes) if successes else 'Nenhum') +
                '\n\nNão foi possível abrir:\n- ' + '\n- '.join(failures),
            )

    def open_consolidator(self):
        module = APP_DIR / 'automacoes' / 'planilha_mae' / 'ui.py'
        if not module.is_file():
            messagebox.showerror('Automação não encontrada',
                                 'O módulo do Consolidador de Planilha Mãe não foi encontrado.')
            return
        try:
            subprocess.Popen([str(python_gui_executable()), '-m', 'automacoes.planilha_mae.ui'], cwd=str(APP_DIR))
        except Exception as exc:
            messagebox.showerror('Não foi possível abrir', f'Falha ao abrir a automação:\n{exc}')

    def open_rank(self):
        module = APP_DIR / 'automacoes' / 'planilha_rank' / 'ui.py'
        if not module.is_file():
            messagebox.showerror('Automação não encontrada',
                                 'O módulo do Gerador de Planilha Rank não foi encontrado.')
            return
        try:
            subprocess.Popen([str(python_gui_executable()), '-m', 'automacoes.planilha_rank.ui'],
                             cwd=str(APP_DIR))
        except Exception as exc:
            messagebox.showerror('Não foi possível abrir', f'Falha ao abrir a automação:\n{exc}')


if __name__ == '__main__':
    AutoPlanHub().mainloop()

