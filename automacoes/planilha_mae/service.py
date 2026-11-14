"""Orquestra os arquivos sem depender da interface gráfica."""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import shutil

from openpyxl import load_workbook

from app_info import REVIEW_NOTICE
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


def build_report_text(config, output, backup, added, events, generated_at=None):
    generated_at = generated_at or datetime.now()
    mode = 'Atualização de uma Planilha Mãe existente' if config.mode == 'update' else 'Criação de uma nova Planilha Mãe'
    daily_files = '\n'.join(f'- {Path(path).name}' for path in config.daily_files)
    event_text = '\n'.join(f'- {event}' for event in events) if events else '- Nenhum conflito ou aviso registrado.'
    backup_text = str(backup) if backup else 'Não se aplica.'
    added_text = str(added) if config.mode == 'update' else 'Arquivo novo criado.'
    return (
        'AUTOPLAN HUB — RELATÓRIO DE PROCESSAMENTO DA PLANILHA MÃE\n'
        '================================================================\n\n'
        '1. IDENTIFICAÇÃO\n'
        f'Data e hora: {generated_at:%d/%m/%Y %H:%M:%S}\n'
        f'Operação: {mode}\n'
        f'Período: {engine.MONTHS_PT[config.month]} {config.year}\n'
        f'Arquivo resultante: {output}\n'
        f'Backup: {backup_text}\n\n'
        '2. ARQUIVOS UTILIZADOS\n'
        f'Relatório final do mês: {Path(config.final_file).name}\n'
        f'Planilhas diárias selecionadas: {len(config.daily_files)}\n'
        f'{daily_files}\n\n'
        '3. RESULTADO RESUMIDO\n'
        f'Registros adicionados durante a atualização: {added_text}\n'
        'A aba Inconformidades e o Painel foram recalculados.\n'
        'Os filtros e o detalhamento por duplo clique foram preparados no Excel.\n\n'
        '4. OCORRÊNCIAS, ATUALIZAÇÕES E DECISÕES\n'
        f'{event_text}\n\n'
        '5. REVISÃO ANTES DA ENTREGA\n'
        f'{REVIEW_NOTICE}\n'
        '- Confira se todas as planilhas diárias foram selecionadas.\n'
        '- Confira os totais do Painel e as linhas de Inconformidades.\n'
        '- Revise conflitos, motivos não mapeados e registros ignorados.\n'
        '- Abra algumas contagens por duplo clique para validar os detalhes.\n'
    )


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
    report_path.write_text(
        build_report_text(config, output, backup, added, report), encoding='utf-8'
    )
    progress(100, 'Processamento concluído.')
    return JobResult(output,report_path,backup,added)

