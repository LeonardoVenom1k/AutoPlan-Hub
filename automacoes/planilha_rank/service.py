"""Orquestra a geração da Planilha Rank sem depender da interface."""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from . import engine


@dataclass(frozen=True)
class RankJobConfig:
    mother_file: str
    output_dir: str
    year: int
    month: int
    previous_rank: str = ''
    overwrite: bool = False


@dataclass(frozen=True)
class RankJobResult:
    output: Path
    report: Path
    base_rank: int
    occurrences: int
    vehicle: int
    unmapped: int
    mapped: int
    mapping_technicians: int
    mapping_providers: int


def planned_output(config):
    return Path(config.output_dir) / engine.output_filename(config.year, config.month)


def validate(config):
    if not Path(config.mother_file).is_file():
        raise ValueError('Selecione uma Planilha Mãe válida.')
    if config.previous_rank and not Path(config.previous_rank).is_file():
        raise ValueError('A Planilha Rank anterior selecionada não foi encontrada.')
    if not Path(config.output_dir).is_dir():
        raise ValueError('Selecione uma pasta de saída válida.')
    if not 2000 <= config.year <= 2100 or not 1 <= config.month <= 12:
        raise ValueError('Selecione um mês e ano válidos.')


def process(config, progress=lambda value, message: None):
    validate(config)
    output = planned_output(config)
    if output.exists() and not config.overwrite:
        raise FileExistsError(output)
    report = []
    progress(2, 'Preparando a Planilha Rank...')
    counts = engine.create_rank_workbook(
        config.mother_file, output, config.previous_rank, report, progress,
    )
    engine.create_native_pivots(output, report, progress)
    report_path = output.parent / f'{output.stem}_relatorio.txt'
    report_path.write_text(
        f'Relatório: {datetime.now():%d/%m/%Y %H:%M:%S}\n'
        f'Arquivo: {output}\n'
        f'Planilha Mãe: {Path(config.mother_file).name}\n'
        f'Planilha Rank usada no mapeamento: '
        f'{Path(config.previous_rank).name if config.previous_rank else "não selecionada"}\n'
        f'Período: {engine.MONTHS_PT[config.month]} {config.year}\n\n'
        + '\n'.join(report),
        encoding='utf-8',
    )
    progress(100, 'Planilha Rank concluída.')
    return RankJobResult(output, report_path, **counts)

