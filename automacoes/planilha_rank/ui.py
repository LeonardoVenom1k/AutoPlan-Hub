# -*- coding: utf-8 -*-
"""Interface gráfica do Gerador de Planilha Rank."""
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app_info import APP_AUTHOR, APP_NAME, APP_VERSION, apply_window_icon, center_window
from automacoes.planilha_mae.theme import COLORS, configure_styles
from . import engine
from .service import RankJobConfig, planned_output, process, validate


class RankWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('AutoPlan Hub - Gerador de Planilha Rank')
        self.minsize(900, 650)
        self.configure(bg=COLORS['background'])
        apply_window_icon(self)
        configure_styles(self)

        now = datetime.now()
        self.mother_file = tk.StringVar()
        self.previous_rank = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / 'Documents'))
        self.year = tk.StringVar(value=str(now.year))
        self.month = tk.StringVar(value=engine.MONTHS_PT[now.month])
        self.events = queue.Queue()
        self.worker = None

        self._build_header()
        self._build_content()
        self._build_footer()
        center_window(self, 1000, 735)
        self.after(120, self._poll_events)

    def _build_header(self):
        header = tk.Frame(self, bg=COLORS['navy'], padx=28, pady=19)
        header.pack(fill='x')
        tk.Label(header, text='Gerador de Planilha Rank', bg=COLORS['navy'], fg='white',
                 font=('Segoe UI', 20, 'bold')).pack(anchor='w')
        tk.Label(header,
                 text='Crie rankings, painéis de ocorrências e análises de supervisão a partir da Planilha Mãe.',
                 bg=COLORS['navy'], fg='#D9EAF7', font=('Segoe UI', 10)).pack(anchor='w', pady=(3, 0))

    def _card(self, parent, title):
        card = ttk.Frame(parent, style='Card.TFrame', padding=17)
        card.pack(fill='x', pady=(0, 12))
        ttk.Label(card, text=title, style='CardTitle.TLabel').grid(
            row=0, column=0, columnspan=3, sticky='w', pady=(0, 10))
        card.columnconfigure(1, weight=1)
        return card

    def _build_content(self):
        container = tk.Frame(self, bg=COLORS['background'])
        container.pack(fill='both', expand=True)
        canvas = tk.Canvas(container, bg=COLORS['background'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        shell = ttk.Frame(canvas, padding=(24, 18, 24, 8))
        window = canvas.create_window((0, 0), window=shell, anchor='nw')
        shell.bind('<Configure>', lambda event: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda event: canvas.itemconfigure(window, width=event.width))
        canvas.bind_all('<MouseWheel>', lambda event: canvas.yview_scroll(int(-event.delta / 120), 'units'))

        source = self._card(shell, '1. Arquivos de origem')
        ttk.Label(source, text='Planilha Mãe', style='Card.TLabel').grid(row=1, column=0, sticky='w')
        ttk.Entry(source, textvariable=self.mother_file).grid(
            row=2, column=0, columnspan=2, sticky='ew', padx=(0, 8), pady=(5, 12))
        ttk.Button(source, text='Selecionar arquivo', command=self._pick_mother).grid(row=2, column=2, pady=(5, 12))

        ttk.Label(source, text='Planilha Rank anterior ou modelo (opcional)',
                  style='Card.TLabel').grid(row=3, column=0, sticky='w')
        ttk.Label(source, text='Reaproveita o vínculo entre prestadores e supervisão.',
                  style='Muted.TLabel').grid(row=3, column=1, sticky='e', padx=8)
        ttk.Entry(source, textvariable=self.previous_rank).grid(
            row=4, column=0, columnspan=2, sticky='ew', padx=(0, 8), pady=(5, 0))
        ttk.Button(source, text='Selecionar Rank', command=self._pick_previous).grid(row=4, column=2, pady=(5, 0))

        settings = self._card(shell, '2. Período e destino')
        ttk.Label(settings, text='Mês', style='Card.TLabel').grid(row=1, column=0, sticky='w')
        ttk.Combobox(settings, textvariable=self.month, values=list(engine.MONTHS_PT.values()),
                     state='readonly', width=18).grid(row=2, column=0, sticky='w', padx=(0, 12), pady=(4, 0))
        ttk.Label(settings, text='Ano', style='Card.TLabel').grid(row=1, column=1, sticky='w')
        ttk.Entry(settings, textvariable=self.year, width=10).grid(row=2, column=1, sticky='w', pady=(4, 0))
        ttk.Button(settings, text='Detectar pela Planilha Mãe',
                   command=self._infer_period).grid(row=2, column=2, sticky='w', padx=12)
        ttk.Label(settings, text='Pasta de saída', style='Card.TLabel').grid(
            row=3, column=0, sticky='w', pady=(14, 0))
        ttk.Entry(settings, textvariable=self.output_dir).grid(
            row=4, column=0, columnspan=2, sticky='ew', padx=(0, 8), pady=(4, 0))
        ttk.Button(settings, text='Selecionar pasta', command=self._pick_output).grid(row=4, column=2, pady=(4, 0))

        activity = self._card(shell, 'Andamento')
        self.progress = ttk.Progressbar(activity, maximum=100)
        self.progress.grid(row=1, column=0, columnspan=3, sticky='ew')
        self.progress_label = ttk.Label(activity, text='Pronto para gerar', style='Muted.TLabel')
        self.progress_label.grid(row=2, column=0, columnspan=3, sticky='w', pady=(6, 7))
        self.log = tk.Text(activity, height=6, wrap='word', font=('Consolas', 9),
                           bg=COLORS['log'], fg='#E6EDF3', insertbackground='white',
                           relief='flat', padx=9, pady=7)
        self.log.tag_configure('ok', foreground='#7EE787')
        self.log.tag_configure('error', foreground='#FF8B82')
        self.log.tag_configure('info', foreground='#8CC8FF')
        self.log.grid(row=3, column=0, columnspan=3, sticky='nsew')

    def _build_footer(self):
        footer = tk.Frame(self, bg='#E2EAF2', padx=24, pady=10)
        footer.pack(fill='x', side='bottom')
        tk.Label(footer, text=f'{APP_NAME}  •  {APP_AUTHOR}  •  v{APP_VERSION}',
                 bg='#E2EAF2', fg=COLORS['muted'], font=('Segoe UI', 9)).pack(side='left')
        self.run_button = ttk.Button(footer, text='Gerar Planilha Rank',
                                     style='Primary.TButton', command=self._start)
        self.run_button.pack(side='right')
        ttk.Button(footer, text='Voltar ao menu', command=self._back_to_menu).pack(
            side='right', padx=(0, 10))

    def _pick_mother(self):
        path = filedialog.askopenfilename(
            title='Selecione a Planilha Mãe', filetypes=[('Excel', '*.xlsx *.xlsm')])
        if path:
            self.mother_file.set(path)
            self._infer_period(show_warning=False)

    def _pick_previous(self):
        path = filedialog.askopenfilename(
            title='Selecione a Planilha Rank anterior ou modelo',
            filetypes=[('Excel', '*.xlsx *.xlsm')])
        if path:
            self.previous_rank.set(path)

    def _pick_output(self):
        path = filedialog.askdirectory(title='Selecione a pasta de saída')
        if path:
            self.output_dir.set(path)

    def _infer_period(self, show_warning=True):
        path = self.mother_file.get().strip()
        if not Path(path).is_file():
            if show_warning:
                messagebox.showinfo('Planilha Mãe', 'Selecione a Planilha Mãe primeiro.')
            return
        try:
            year, month = engine.infer_period_from_mother(path)
            self.year.set(str(year))
            self.month.set(engine.MONTHS_PT[month])
        except Exception as exc:
            if show_warning:
                messagebox.showerror('Período não identificado', str(exc))

    def _config(self, overwrite=False):
        try:
            month = next(number for number, name in engine.MONTHS_PT.items()
                         if name == self.month.get())
            year = int(self.year.get())
        except Exception as exc:
            raise ValueError('Selecione um mês e ano válidos.') from exc
        return RankJobConfig(
            self.mother_file.get().strip(), self.output_dir.get().strip(),
            year, month, self.previous_rank.get().strip(), overwrite,
        )

    def _start(self):
        try:
            config = self._config()
            validate(config)
            output = planned_output(config)
            if output.exists():
                if not messagebox.askyesno(
                        'Arquivo já existe', f'{output.name} já existe. Deseja substituir?'):
                    return
                config = self._config(overwrite=True)
        except Exception as exc:
            messagebox.showerror('Confira os campos', str(exc))
            return
        self.log.delete('1.0', 'end')
        self.progress['value'] = 0
        self.run_button.state(['disabled'])
        self._log('Iniciando geração da Planilha Rank...', 'info')
        self.worker = threading.Thread(target=self._worker, args=(config,), daemon=True)
        self.worker.start()

    def _worker(self, config):
        try:
            result = process(
                config, lambda value, message: self.events.put(('progress', value, message)))
            self.events.put(('done', result))
        except Exception as exc:
            self.events.put(('error', exc))

    def _poll_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == 'progress':
                    _, value, message = event
                    self.progress['value'] = value
                    self.progress_label.config(text=f'{value}% — {message}')
                    self._log(message, 'info')
                elif event[0] == 'done':
                    result = event[1]
                    self.run_button.state(['!disabled'])
                    self._log(f'Concluído: {result.output}', 'ok')
                    messagebox.showinfo(
                        'Planilha Rank concluída',
                        f'Arquivo:\n{result.output}\n\nRelatório:\n{result.report}\n\n'
                        f'Visitas no ranking: {result.base_rank}\n'
                        f'Ocorrências: {result.occurrences}\n'
                        f'Não mapeadas: {result.unmapped}',
                    )
                elif event[0] == 'error':
                    exc = event[1]
                    self.run_button.state(['!disabled'])
                    self._log(f'ERRO: {exc}', 'error')
                    messagebox.showerror('Falha na geração', str(exc))
        except queue.Empty:
            pass
        self.after(120, self._poll_events)

    def _back_to_menu(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning(
                'Processamento em andamento',
                'Aguarde a geração terminar antes de voltar ao menu.')
            return
        self.destroy()

    def _log(self, message, tag):
        self.log.insert('end', message + '\n', tag)
        self.log.see('end')


def run():
    RankWindow().mainloop()


if __name__ == '__main__':
    run()

