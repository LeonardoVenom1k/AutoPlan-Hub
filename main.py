# -*- coding: utf-8 -*-
"""Menu principal do AutoPlan Hub."""
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from app_info import APP_AUTHOR, APP_NAME, APP_VERSION, apply_window_icon, center_window
from automacoes.ambiente_trabalho.service import launch_environment, load_config
from atualizacao.service import fetch_release, is_newer, prepare_update

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
        self.minsize(820, 600)
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

        content_area = tk.Frame(self, bg='#EEF3F8')
        content_area.pack(fill='both', expand=True)
        canvas = tk.Canvas(content_area, bg='#EEF3F8', highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_area, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        content = tk.Frame(canvas, bg='#EEF3F8', padx=30, pady=18)
        content_window = canvas.create_window((0, 0), window=content, anchor='nw')
        content.bind(
            '<Configure>',
            lambda _event: canvas.configure(scrollregion=canvas.bbox('all')),
        )
        canvas.bind(
            '<Configure>',
            lambda event: canvas.itemconfigure(content_window, width=event.width),
        )

        def scroll_menu(event):
            canvas.yview_scroll(int(-event.delta / 120), 'units')

        canvas.bind('<Enter>', lambda _event: canvas.bind_all('<MouseWheel>', scroll_menu))
        canvas.bind('<Leave>', lambda _event: canvas.unbind_all('<MouseWheel>'))
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
        ttk.Button(work_actions, text='Abrir', style='Hub.TButton', width=13,
                   command=self.launch_work_environment).pack(side='left')
        ttk.Button(work_actions, text='Configurar', style='Hub.TButton', width=13,
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

        planned = tk.Frame(content, bg='#EEF3F8')
        planned.pack(fill='x', anchor='n', pady=(14, 0))
        planned.columnconfigure(0, weight=1)
        planned.columnconfigure(1, weight=1)

        location_card = tk.Frame(planned, bg='#F8FAFC', highlightbackground='#C9D5E3',
                                 highlightthickness=1, padx=18, pady=14)
        location_card.grid(row=0, column=0, sticky='nsew', padx=(0, 7))
        tk.Label(location_card, text='Auditoria de localização', bg='#F8FAFC', fg='#526273',
                 font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        tk.Label(location_card,
                 text='Auditar alterações de localização em atendimentos já finalizados.',
                 bg='#F8FAFC', fg='#64748B', font=('Segoe UI', 9),
                 wraplength=330, justify='left').pack(anchor='w', pady=(4, 10))
        ttk.Button(location_card, text='Em desenvolvimento', state='disabled').pack(anchor='w')

        cleanup_card = tk.Frame(planned, bg='#F8FAFC', highlightbackground='#C9D5E3',
                                highlightthickness=1, padx=18, pady=14)
        cleanup_card.grid(row=0, column=1, sticky='nsew', padx=(7, 0))
        tk.Label(cleanup_card, text='Limpeza de planilhas diárias', bg='#F8FAFC', fg='#526273',
                 font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        tk.Label(cleanup_card,
                 text='Limpar e preparar planilhas de Logística, Raízen ou SPTrans.',
                 bg='#F8FAFC', fg='#64748B', font=('Segoe UI', 9),
                 wraplength=330, justify='left').pack(anchor='w', pady=(4, 10))
        ttk.Button(cleanup_card, text='Em desenvolvimento', state='disabled').pack(anchor='w')

        collection_card = tk.Frame(planned, bg='#F8FAFC', highlightbackground='#C9D5E3',
                                   highlightthickness=1, padx=18, pady=14)
        collection_card.grid(row=1, column=0, columnspan=2, sticky='nsew', pady=(14, 0))
        tk.Label(collection_card, text='Cobrador automático', bg='#F8FAFC', fg='#526273',
                 font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        tk.Label(
            collection_card,
            text=('Programará cobranças de previsão de chegada, usando a empresa da planilha '
                  'e respeitando os estados que não devem receber mensagem.'),
            bg='#F8FAFC', fg='#64748B', font=('Segoe UI', 9),
            wraplength=720, justify='left',
        ).pack(anchor='w', pady=(4, 10))
        ttk.Button(collection_card, text='Em desenvolvimento', state='disabled').pack(anchor='w')

        footer = tk.Frame(self, bg='#E3EAF2', padx=20, pady=9)
        footer.pack(fill='x', side='bottom')
        tk.Label(footer, text=f'{APP_NAME}  •  Desenvolvido por {APP_AUTHOR}  •  Versão {APP_VERSION}', bg='#E3EAF2', fg='#526273',
                 font=('Segoe UI', 9)).pack(side='left')
        ttk.Button(footer, text='Fechar', style='Secondary.TButton',
                   command=self.destroy).pack(side='right')
        self.update_button = tk.Button(
            footer, text='Verificando atualização...', font=('Segoe UI', 8, 'bold'),
            bg='#D7E0EA', fg='#334155', activebackground='#C9D5E3',
            disabledforeground='#526273', relief='flat', padx=10, pady=5,
            cursor='arrow', state='disabled',
        )
        self.update_button.pack(side='right', padx=(0, 10))
        self.available_release = None
        self._update_check_running = False
        self._update_check_job = None
        center_window(self, 920, 700)
        self.after(1000, self.check_for_update)

    def check_for_update(self):
        if self._update_check_running:
            return
        self._update_check_running = True
        self.update_button.configure(
            text='Verificando...', state='disabled', bg='#D7E0EA', fg='#526273', cursor='arrow')
        threading.Thread(target=self._check_update_worker, daemon=True).start()

    def _schedule_next_update_check(self):
        self._update_check_running = False
        if self._update_check_job is not None:
            self.after_cancel(self._update_check_job)
        self._update_check_job = self.after(60 * 60 * 1000, self.check_for_update)

    def _check_update_worker(self):
        try:
            release = fetch_release()
            self.after(0, self._show_update_result, release, None)
        except Exception as exc:
            self.after(0, self._show_update_result, None, exc)

    def _show_update_result(self, release, error):
        self._schedule_next_update_check()
        if error:
            self.update_button.configure(
                text='Verificação indisponível', state='disabled', bg='#B7791F', fg='white',
                disabledforeground='white', cursor='arrow', command=lambda: None)
            return
        if is_newer(release):
            self.available_release = release
            self.update_button.configure(
                text=f'Baixar v{release.version}', state='normal', bg='#C62828', fg='white',
                activebackground='#A61B1B', cursor='hand2', command=self.download_update)
            return
        self.available_release = None
        self.update_button.configure(
            text=f'Versão {APP_VERSION} atual', state='disabled', bg='#2E7D32', fg='white',
            disabledforeground='white', cursor='arrow', command=lambda: None)

    def download_update(self):
        release = self.available_release
        if release is None:
            return
        if not messagebox.askyesno(
            'Atualizar AutoPlan Hub',
            f'Baixar e instalar o AutoPlan Hub {release.version}?\n\n'
            'O programa será fechado, o executável antigo será removido e a nova versão será aberta.',
        ):
            return
        self.update_button.configure(
            text='Baixando...', state='disabled', bg='#C62828', fg='white', cursor='arrow')
        threading.Thread(target=self._download_update_worker, args=(release,), daemon=True).start()

    def _download_update_worker(self, release):
        try:
            target = prepare_update(release)
            self.after(0, self._finish_update, target, None)
        except Exception as exc:
            self.after(0, self._finish_update, None, exc)

    def _finish_update(self, target, error):
        if error:
            self.update_button.configure(
                text=f'Baixar v{self.available_release.version}', state='normal',
                bg='#C62828', fg='white', cursor='hand2', command=self.download_update)
            messagebox.showerror('Falha na atualização', str(error))
            return
        messagebox.showinfo(
            'Atualização preparada',
            f'A nova versão será instalada em:\n{target}\n\nO AutoPlan Hub será reiniciado.',
        )
        self.destroy()

    def open_work_environment(self):
        module = APP_DIR / 'automacoes' / 'ambiente_trabalho' / 'ui.py'
        if not getattr(sys, 'frozen', False) and not module.is_file():
            messagebox.showerror('Recurso não encontrado',
                                 'O módulo do Ambiente de Trabalho não foi encontrado.')
            return
        try:
            self._launch_module('automacoes.ambiente_trabalho.ui', '--ambiente-trabalho')
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
        if not getattr(sys, 'frozen', False) and not module.is_file():
            messagebox.showerror('Automação não encontrada',
                                 'O módulo do Consolidador de Planilha Mãe não foi encontrado.')
            return
        try:
            self._launch_module('automacoes.planilha_mae.ui', '--planilha-mae')
        except Exception as exc:
            messagebox.showerror('Não foi possível abrir', f'Falha ao abrir a automação:\n{exc}')

    def open_rank(self):
        module = APP_DIR / 'automacoes' / 'planilha_rank' / 'ui.py'
        if not getattr(sys, 'frozen', False) and not module.is_file():
            messagebox.showerror('Automação não encontrada',
                                 'O módulo do Gerador de Planilha Rank não foi encontrado.')
            return
        try:
            self._launch_module('automacoes.planilha_rank.ui', '--planilha-rank')
        except Exception as exc:
            messagebox.showerror('Não foi possível abrir', f'Falha ao abrir a automação:\n{exc}')

    @staticmethod
    def _launch_module(module, frozen_argument):
        if getattr(sys, 'frozen', False):
            command = [sys.executable, frozen_argument]
        else:
            command = [str(python_gui_executable()), '-m', module]
        subprocess.Popen(command, cwd=str(APP_DIR))


def run():
    argument = sys.argv[1] if len(sys.argv) > 1 else ''
    if argument == '--ambiente-trabalho':
        from automacoes.ambiente_trabalho.ui import WorkEnvironmentWindow
        WorkEnvironmentWindow().mainloop()
    elif argument == '--planilha-mae':
        from automacoes.planilha_mae.ui import ConsolidatorWindow
        ConsolidatorWindow().mainloop()
    elif argument == '--planilha-rank':
        from automacoes.planilha_rank.ui import RankWindow
        RankWindow().mainloop()
    else:
        AutoPlanHub().mainloop()


if __name__ == '__main__':
    run()

