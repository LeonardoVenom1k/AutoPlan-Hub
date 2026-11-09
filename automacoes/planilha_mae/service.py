"""Orquestra os arquivos sem depender da interface gráfica."""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import shutil

from openpyxl import load_workbook

from . import engine


@dataclass(frozen=True)
class JobConfig:
    mode: str
    daily_files: tuple[str, ...]
    final_file: str
    year: int
    month: int
    output_dir: str = ''
    mother_file: str = ''
    overwrite: bool = False


@dataclass(frozen=True)
class JobResult:
    output: Path
    report: Path
    backup: Path | None
    added: int


def planned_output(config: JobConfig) -> Path:
    if config.mode == 'update':
        return Path(config.mother_file)
    return Path(config.output_dir) / engine.output_filename(config.year, config.month)


def validate(config: JobConfig):
    if not config.daily_files:
        raise ValueError('Selecione pelo menos uma planilha diária.')
    if not Path(config.final_file).is_file():
        raise ValueError('Selecione o relatório final do mês.')
    if not 2000 <= config.year <= 2100 or not 1 <= config.month <= 12:
        raise ValueError('Selecione um mês e ano válidos.')
    if config.mode == 'update':
        if not Path(config.mother_file).is_file():
            raise ValueError('Selecione a Planilha Mãe que será atualizada.')
    elif not Path(config.output_dir).is_dir():
        raise ValueError('Selecione uma pasta de saída válida.')


def process(config: JobConfig, progress=lambda value, message: None) -> JobResult:
    validate(config)
    output = planned_output(config)
    if config.mode == 'create' and output.exists() and not config.overwrite:
        raise FileExistsError(output)

    report=[]
    backup=None
    added=0
    progress(2, 'Preparando o processamento...')

    if config.mode == 'create':
        engine.create_new_mother(config.daily_files, config.final_file, str(output),
                                 config.year, config.month, report, progress)
    else:
        backup=output.with_name(f'{output.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{output.suffix}')
        shutil.copy2(output,backup)
        report.append(f'Backup criado antes da atualização: {backup.name}')
        progress(5, f'Backup criado: {backup.name}')
        workbook=load_workbook(output,keep_vba=output.suffix.lower()=='.xlsm')
        try:
            added=engine.merge_into_existing(workbook,config.daily_files,config.final_file,
                                             config.year,config.month,report,progress)
            workbook.save(output)
        finally:
            workbook.close()

    engine.create_native_pivot(output,report,progress)
    report_path=output.parent/(output.stem+'_relatorio.txt')
    with open(report_path,'w',encoding='utf-8') as file:
        file.write(f'Relatório: {datetime.now():%d/%m/%Y %H:%M:%S}\n')
        file.write(f'Arquivo: {output}\n')
        file.write(f'Período: {engine.MONTHS_PT[config.month]} {config.year}\n')
        file.write(f'Modo: {config.mode}\n\n')
        file.write('\n'.join(report) if report else 'Nenhum conflito ou aviso registrado.')
    progress(100, 'Processamento concluído.')
    return JobResult(output,report_path,backup,added)

