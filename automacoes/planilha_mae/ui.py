# -*- coding: utf-8 -*-
"""Interface gráfica do Consolidador de Planilha Mãe."""
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app_info import (APP_AUTHOR, APP_NAME, APP_VERSION, REVIEW_NOTICE,
                      apply_window_icon, center_window, play_notification_sound)
from . import engine
from .service import JobConfig, planned_output, process
from .theme import COLORS, configure_styles


class ConsolidatorWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('AutoPlan Hub - Consolidador de Planilha Mãe')
        self.minsize(900, 700)
        self.configure(bg=COLORS['background'])
        apply_window_icon(self)
        configure_styles(self)

        now=datetime.now()
        self.mode=tk.StringVar(value='create')
        self.reference=tk.StringVar()
        self.final_file=tk.StringVar()
        self.output_dir=tk.StringVar(value=str(Path.home()/'Documents'))
        self.year=tk.StringVar(value=str(now.year))
        self.month=tk.StringVar(value=engine.MONTHS_PT[now.month])
        self.daily_files=[]
        self.events=queue.Queue()
        self.worker=None

        self._build_header()
        self._build_content()
        self._build_footer()
        self._refresh_mode()
        center_window(self,1000,780)
        self.after(120,self._poll_events)

    def _build_header(self):
        header=tk.Frame(self,bg=COLORS['navy'],padx=28,pady=19)
        header.pack(fill='x')
        tk.Label(header,text='Consolidador de Planilha Mãe',bg=COLORS['navy'],fg='white',
                 font=('Segoe UI',20,'bold')).pack(anchor='w')
        tk.Label(header,text='Consolide arquivos diários, inconformidades e o Painel em um único fluxo.',
                 bg=COLORS['navy'],fg='#D9EAF7',font=('Segoe UI',10)).pack(anchor='w',pady=(3,0))

    def _card(self,parent,title):
        card=ttk.Frame(parent,style='Card.TFrame',padding=17)
        card.pack(fill='x',pady=(0,12))
        ttk.Label(card,text=title,style='CardTitle.TLabel').grid(row=0,column=0,columnspan=3,sticky='w',pady=(0,10))
        card.columnconfigure(1,weight=1)
        return card

    def _build_content(self):
        container=tk.Frame(self,bg=COLORS['background'])
        container.pack(fill='both',expand=True)
        canvas=tk.Canvas(container,bg=COLORS['background'],highlightthickness=0)
        page_scroll=ttk.Scrollbar(container,orient='vertical',command=canvas.yview)
        canvas.configure(yscrollcommand=page_scroll.set)
        canvas.pack(side='left',fill='both',expand=True)
        page_scroll.pack(side='right',fill='y')
        shell=ttk.Frame(canvas,padding=(24,18,24,8))
        window=canvas.create_window((0,0),window=shell,anchor='nw')
        shell.bind('<Configure>',lambda event:canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',lambda event:canvas.itemconfigure(window,width=event.width))
        canvas.bind_all('<MouseWheel>',lambda event:canvas.yview_scroll(int(-event.delta/120),'units'))

        mode_card=self._card(shell,'1. Escolha o modo')
        ttk.Radiobutton(mode_card,text='Criar uma nova Planilha Mãe',variable=self.mode,value='create',
                        command=self._refresh_mode,style='Mode.TRadiobutton').grid(row=1,column=0,sticky='w',padx=(0,20))
        ttk.Radiobutton(mode_card,text='Atualizar uma Planilha Mãe existente',variable=self.mode,value='update',
                        command=self._refresh_mode,style='Mode.TRadiobutton').grid(row=1,column=1,sticky='w')

        self.reference_card=self._card(shell,'Planilha Mãe para atualizar')
        self.reference_entry=ttk.Entry(self.reference_card,textvariable=self.reference)
        self.reference_entry.grid(row=1,column=0,columnspan=2,sticky='ew',padx=(0,8))
        ttk.Button(self.reference_card,text='Selecionar arquivo',command=self._pick_reference).grid(row=1,column=2)

        source_card=self._card(shell,'2. Arquivos de origem')
        self.source_card=source_card
        ttk.Label(source_card,text='Planilhas diárias',style='Card.TLabel').grid(row=1,column=0,sticky='w')
        self.files_count=ttk.Label(source_card,text='Nenhum arquivo selecionado',style='Muted.TLabel')
        self.files_count.grid(row=1,column=1,sticky='e',padx=8)
        ttk.Button(source_card,text='Selecionar várias',command=self._pick_daily).grid(row=1,column=2)
        list_frame=ttk.Frame(source_card,style='Card.TFrame')
        list_frame.grid(row=2,column=0,columnspan=3,sticky='ew',pady=(7,13))
        self.files_list=tk.Listbox(list_frame,height=4,font=('Segoe UI',9),borderwidth=1,
                                   relief='solid',activestyle='none',exportselection=False)
        scrollbar=ttk.Scrollbar(list_frame,orient='vertical',command=self.files_list.yview)
        self.files_list.configure(yscrollcommand=scrollbar.set)
        self.files_list.pack(side='left',fill='both',expand=True)
        scrollbar.pack(side='right',fill='y')

        ttk.Label(source_card,text='Relatório final do mês (resultado e horários)',style='Card.TLabel').grid(row=3,column=0,sticky='w')
        ttk.Entry(source_card,textvariable=self.final_file).grid(row=4,column=0,columnspan=2,sticky='ew',padx=(0,8),pady=(5,0))
        ttk.Button(source_card,text='Selecionar relatório',command=self._pick_final).grid(row=4,column=2,pady=(5,0))

        settings=self._card(shell,'3. Período e destino')
        self.settings_card=settings
        ttk.Label(settings,text='Mês',style='Card.TLabel').grid(row=1,column=0,sticky='w')
        ttk.Combobox(settings,textvariable=self.month,values=list(engine.MONTHS_PT.values()),
                     state='readonly',width=17).grid(row=2,column=0,sticky='w',padx=(0,12),pady=(4,0))
        ttk.Label(settings,text='Ano',style='Card.TLabel').grid(row=1,column=1,sticky='w')
        ttk.Entry(settings,textvariable=self.year,width=10).grid(row=2,column=1,sticky='w',pady=(4,0))
        ttk.Button(settings,text='Inferir pelos arquivos',command=self._infer_dates).grid(row=2,column=2,sticky='w',padx=12)
        self.output_label=ttk.Label(settings,text='Pasta de saída',style='Card.TLabel')
        self.output_label.grid(row=3,column=0,sticky='w',pady=(14,0))
        self.output_entry=ttk.Entry(settings,textvariable=self.output_dir)
        self.output_entry.grid(row=4,column=0,columnspan=2,sticky='ew',padx=(0,8),pady=(4,0))
        self.output_button=ttk.Button(settings,text='Selecionar pasta',command=self._pick_output)
        self.output_button.grid(row=4,column=2,pady=(4,0))

        activity=self._card(shell,'Andamento')
        self.progress=ttk.Progressbar(activity,maximum=100)
        self.progress.grid(row=1,column=0,columnspan=3,sticky='ew')
        self.progress_label=ttk.Label(activity,text='Pronto para iniciar',style='Muted.TLabel')
        self.progress_label.grid(row=2,column=0,columnspan=3,sticky='w',pady=(6,7))
        self.log=tk.Text(activity,height=5,wrap='word',font=('Consolas',9),bg=COLORS['log'],
                         fg='#E6EDF3',insertbackground='white',relief='flat',padx=9,pady=7)
        self.log.tag_configure('ok',foreground='#7EE787')
        self.log.tag_configure('error',foreground='#FF8B82')
        self.log.tag_configure('info',foreground='#8CC8FF')
        self.log.grid(row=3,column=0,columnspan=3,sticky='nsew')

    def _build_footer(self):
        footer=tk.Frame(self,bg='#E2EAF2',padx=24,pady=10)
        footer.pack(fill='x',side='bottom')
        tk.Label(footer,text=f'{APP_NAME}  •  {APP_AUTHOR}  •  v{APP_VERSION}',bg='#E2EAF2',fg=COLORS['muted'],font=('Segoe UI',9)).pack(side='left')
        self.run_button=ttk.Button(footer,text='Executar processamento',style='Primary.TButton',command=self._start)
        self.run_button.pack(side='right')
        ttk.Button(footer,text='Voltar ao menu',command=self._back_to_menu).pack(side='right',padx=(0,10))

    def _back_to_menu(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning('Processamento em andamento',
                                   'Aguarde o processamento terminar antes de voltar ao menu.')
            return
        self.destroy()

    def _refresh_mode(self):
        if self.mode.get()=='update':
            self.reference_card.pack(fill='x',pady=(0,12),before=self.source_card)
            self.output_label.grid_remove(); self.output_entry.grid_remove(); self.output_button.grid_remove()
            self.after_idle(self.reference_entry.focus_set)
        else:
            self.reference_card.pack_forget()
            self.output_label.grid(); self.output_entry.grid(); self.output_button.grid()

    def _pick_reference(self):
        path=filedialog.askopenfilename(title='Selecione a Planilha Mãe',filetypes=[('Excel','*.xlsx *.xlsm')])
        if path: self.reference.set(path)

    def _pick_daily(self):
        paths=filedialog.askopenfilenames(title='Selecione as planilhas diárias',filetypes=[('Excel','*.xlsx *.xlsm')])
        if not paths: return
        self.daily_files=list(paths)
        self.files_list.delete(0,'end')
        for path in paths: self.files_list.insert('end',Path(path).name)
        self.files_count.config(text=f'{len(paths)} arquivo(s) selecionado(s)')
        self._infer_dates(show_warning=False)

    def _pick_final(self):
        path=filedialog.askopenfilename(title='Selecione o relatório final do mês',filetypes=[('Excel','*.xlsx *.xlsm')])
        if path: self.final_file.set(path)

    def _pick_output(self):
        path=filedialog.askdirectory(title='Selecione a pasta de saída')
        if path: self.output_dir.set(path)

    def _infer_dates(self,show_warning=True):
        if not self.daily_files:
            if show_warning: messagebox.showinfo('Arquivos diários','Selecione as planilhas diárias primeiro.')
            return
        year,month=engine.infer_period(self.daily_files)
        self.year.set(str(year)); self.month.set(engine.MONTHS_PT[month])

    def _config(self,overwrite=False):
        try:
            month=next(number for number,name in engine.MONTHS_PT.items() if name==self.month.get())
            year=int(self.year.get())
        except Exception as exc:
            raise ValueError('Selecione um mês e ano válidos.') from exc
        return JobConfig(self.mode.get(),tuple(self.daily_files),self.final_file.get().strip(),year,month,
                         self.output_dir.get().strip(),self.reference.get().strip(),overwrite)

    def _start(self):
        try:
            config=self._config()
            from .service import validate
            validate(config)
            output=planned_output(config)
            if config.mode=='create' and output.exists():
                if not messagebox.askyesno('Arquivo já existe',f'{output.name} já existe. Deseja substituir?'):
                    return
                config=self._config(overwrite=True)
        except Exception as exc:
            messagebox.showerror('Confira os campos',str(exc)); return
        self.log.delete('1.0','end'); self.progress['value']=0
        self.run_button.state(['disabled'])
        self._log('Iniciando processamento...','info')
        self.worker=threading.Thread(target=self._worker,args=(config,),daemon=True)
        self.worker.start()

    def _worker(self,config):
        try:
            result=process(config,lambda value,message:self.events.put(('progress',value,message)))
            self.events.put(('done',result))
        except Exception as exc:
            self.events.put(('error',exc))

    def _poll_events(self):
        try:
            while True:
                event=self.events.get_nowait()
                if event[0]=='progress':
                    _,value,message=event; self.progress['value']=value
                    self.progress_label.config(text=f'{value}% — {message}'); self._log(message,'info')
                elif event[0]=='done':
                    result=event[1]; self.run_button.state(['!disabled']); self._log(f'Concluído: {result.output}','ok')
                    play_notification_sound()
                    self._log(f'REVISE ANTES DE ENTREGAR: {REVIEW_NOTICE}','info')
                    backup=f'\n\nBackup: {result.backup}' if result.backup else ''
                    messagebox.showwarning(
                        'Processamento concluído — revisão necessária',
                        f'Arquivo:\n{result.output}\n\nRelatório:\n{result.report}{backup}\n\n'
                        f'ATENÇÃO:\n{REVIEW_NOTICE}',
                    )
                elif event[0]=='error':
                    exc=event[1]; self.run_button.state(['!disabled']); self._log(f'ERRO: {exc}','error')
                    messagebox.showerror('Falha no processamento',str(exc))
        except queue.Empty:
            pass
        self.after(120,self._poll_events)

    def _log(self,message,tag):
        self.log.insert('end',message+'\n',tag); self.log.see('end')


def run():
    ConsolidatorWindow().mainloop()


if __name__=='__main__':
    run()

