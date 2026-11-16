"""Tela de configuração do ambiente de trabalho."""
import tkinter as tk
from tkinter import messagebox, ttk

from app_info import APP_NAME, apply_window_icon, center_window
from automacoes.ambiente_trabalho.service import (
    WorkEnvironmentConfig,
    launch_environment,
    load_config,
    save_config,
)


class WorkEnvironmentWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f'{APP_NAME} - Ambiente de trabalho')
        self.configure(bg='#EEF3F8')
        self.resizable(False, False)
        apply_window_icon(self)
        self.config_data = load_config()
        self.vars = {}
        self._build()
        center_window(self, 760, 700)

    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'), padding=(14, 10))
        style.configure('Secondary.TButton', font=('Segoe UI', 10), padding=(12, 9))

        header = tk.Frame(self, bg='#17365D', padx=28, pady=20)
        header.pack(fill='x')
        tk.Label(header, text='Abrir ambiente de trabalho', bg='#17365D', fg='white',
                 font=('Segoe UI', 21, 'bold')).pack(anchor='w')
        tk.Label(header, text='Escolha o que será aberto e salve para as próximas vezes.',
                 bg='#17365D', fg='#D9EAF7', font=('Segoe UI', 10)).pack(anchor='w', pady=(3, 0))

        body = tk.Frame(self, bg='#EEF3F8', padx=28, pady=20)
        body.pack(fill='both', expand=True)

        card = tk.Frame(body, bg='white', highlightbackground='#C9D5E3',
                        highlightthickness=1, padx=22, pady=18)
        card.pack(fill='x')
        tk.Label(card, text='Aplicativos de comunicação', bg='white', fg='#17365D',
                 font=('Segoe UI', 13, 'bold')).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 10))
        tk.Label(card, text='Abrir', bg='white', fg='#64748B', font=('Segoe UI', 9, 'bold')).grid(row=1, column=0)
        tk.Label(card, text='Ferramenta', bg='white', fg='#64748B', font=('Segoe UI', 9, 'bold')).grid(row=1, column=1, sticky='w')
        tk.Label(card, text='Onde abrir', bg='white', fg='#64748B', font=('Segoe UI', 9, 'bold')).grid(row=1, column=2, sticky='w')
        for row, (key, label) in enumerate((('whatsapp', 'WhatsApp'), ('outlook', 'Outlook'), ('teams', 'Teams')), start=2):
            enabled = tk.BooleanVar(value=getattr(self.config_data, f'{key}_enabled'))
            mode = tk.StringVar(value='Desktop' if getattr(self.config_data, f'{key}_mode') == 'desktop' else 'Web no Chrome')
            self.vars[f'{key}_enabled'] = enabled
            self.vars[f'{key}_mode'] = mode
            ttk.Checkbutton(card, variable=enabled).grid(row=row, column=0, padx=(2, 12), pady=6)
            tk.Label(card, text=label, bg='white', fg='#1F2937', font=('Segoe UI', 10)).grid(row=row, column=1, sticky='w', pady=6)
            ttk.Combobox(card, textvariable=mode, state='readonly', width=19,
                         values=('Web no Chrome', 'Desktop')).grid(row=row, column=2, sticky='w', padx=(25, 0), pady=6)
        card.columnconfigure(1, weight=1)

        extras = tk.Frame(body, bg='white', highlightbackground='#C9D5E3',
                          highlightthickness=1, padx=22, pady=18)
        extras.pack(fill='x', pady=(14, 0))
        tk.Label(extras, text='Outras ferramentas', bg='white', fg='#17365D',
                 font=('Segoe UI', 13, 'bold')).pack(anchor='w', pady=(0, 8))
        labels = (
            ('maxtrack', 'Maxtrack GO'), ('zendesk', 'Zendesk'),
            ('chatgpt', 'ChatGPT'), ('downloads', 'Pasta Downloads'),
            ('notepad', 'Bloco de Notas'),
        )
        grid = tk.Frame(extras, bg='white')
        grid.pack(fill='x')
        for index, (key, label) in enumerate(labels):
            var = tk.BooleanVar(value=getattr(self.config_data, f'{key}_enabled'))
            self.vars[f'{key}_enabled'] = var
            ttk.Checkbutton(grid, text=label, variable=var).grid(
                row=index // 2, column=index % 2, sticky='w', padx=(0, 55), pady=7)

        note = tk.Label(body,
                        text='As escolhas ficam salvas somente neste computador. O AutoPlan Hub não fecha o Chrome nem altera abas já abertas.',
                        bg='#EEF3F8', fg='#526273', font=('Segoe UI', 9), wraplength=680, justify='left')
        note.pack(anchor='w', pady=(14, 0))

        actions = tk.Frame(self, bg='#E3EAF2', padx=22, pady=12)
        actions.pack(fill='x', side='bottom')
        ttk.Button(actions, text='Voltar ao menu', style='Secondary.TButton', command=self.destroy).pack(side='left')
        ttk.Button(actions, text='Salvar preferências', style='Secondary.TButton', command=self._save).pack(side='right')
        ttk.Button(actions, text='Salvar e abrir agora', style='Primary.TButton', command=self._save_and_launch).pack(side='right', padx=(0, 10))

    def _config_from_form(self):
        return WorkEnvironmentConfig(
            whatsapp_enabled=self.vars['whatsapp_enabled'].get(),
            whatsapp_mode='desktop' if self.vars['whatsapp_mode'].get() == 'Desktop' else 'web',
            outlook_enabled=self.vars['outlook_enabled'].get(),
            outlook_mode='desktop' if self.vars['outlook_mode'].get() == 'Desktop' else 'web',
            teams_enabled=self.vars['teams_enabled'].get(),
            teams_mode='desktop' if self.vars['teams_mode'].get() == 'Desktop' else 'web',
            maxtrack_enabled=self.vars['maxtrack_enabled'].get(),
            zendesk_enabled=self.vars['zendesk_enabled'].get(),
            chatgpt_enabled=self.vars['chatgpt_enabled'].get(),
            downloads_enabled=self.vars['downloads_enabled'].get(),
            notepad_enabled=self.vars['notepad_enabled'].get(),
        )

    def _save(self, show_message=True):
        try:
            self.config_data = self._config_from_form()
            save_config(self.config_data)
            if show_message:
                messagebox.showinfo('Preferências salvas', 'Suas escolhas foram salvas neste computador.')
            return True
        except OSError as exc:
            messagebox.showerror('Não foi possível salvar', f'Falha ao salvar as preferências:\n{exc}')
            return False

    def _save_and_launch(self):
        if not self._save(show_message=False):
            return
        successes, failures = launch_environment(self.config_data)
        if failures:
            messagebox.showwarning(
                'Ambiente aberto com avisos',
                'Itens abertos:\n- ' + ('\n- '.join(successes) if successes else 'Nenhum') +
                '\n\nNão foi possível abrir:\n- ' + '\n- '.join(failures),
            )
        else:
            messagebox.showinfo('Ambiente de trabalho', 'As ferramentas selecionadas foram abertas.')


if __name__ == '__main__':
    WorkEnvironmentWindow().mainloop()

