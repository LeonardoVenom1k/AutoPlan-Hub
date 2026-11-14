"""Orquestra a geração da Planilha Rank sem depender da interface."""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app_info import REVIEW_NOTICE
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


def build_report_text(config, output, counts, events, generated_at=None):
    generated_at = generated_at or datetime.now()
    previous = Path(config.previous_rank).name if config.previous_rank else 'Não selecionada.'
    event_text = '\n'.join(f'- {event}' for event in events) if events else '- Nenhum aviso registrado.'
    return (
        'AUTOPLAN HUB — RELATÓRIO DE PROCESSAMENTO DA PLANILHA RANK\n'
        '================================================================\n\n'
        '1. IDENTIFICAÇÃO\n'
        f'Data e hora: {generated_at:%d/%m/%Y %H:%M:%S}\n'
        f'Período: {engine.MONTHS_PT[config.month]} {config.year}\n'
        f'Arquivo resultante: {output}\n\n'
        '2. ARQUIVOS UTILIZADOS\n'
        f'Planilha Mãe: {Path(config.mother_file).name}\n'
        f'Planilha Rank anterior usada no mapeamento: {previous}\n\n'
        '3. RESULTADO RESUMIDO\n'
        f'Visitas consideradas no ranking: {counts["base_rank"]}\n'
        f'Ocorrências penalizadoras: {counts["occurrences"]}\n'
        f'Registros do Painel Veículo Indisponível: {counts["vehicle"]}\n'
        f'Técnicos carregados para o mapeamento: {counts["mapping_technicians"]}\n'
        f'Prestadores sem ambiguidade carregados: {counts["mapping_providers"]}\n'
        f'Ocorrências com supervisão mapeada: {counts["mapped"]}\n'
        f'Ocorrências sem supervisão mapeada: {counts["unmapped"]}\n\n'
        '4. DETALHAMENTO DO PROCESSAMENTO\n'
        f'{event_text}\n\n'
        '5. COMO O MAPEAMENTO FOI FEITO\n'
        'A supervisão é localizada primeiro pelo campo Técnico. O prestador é usado '
        'somente quando todos os seus registros conhecidos apontam para a mesma supervisão.\n'
        'No Painel Supervisão, a apresentação permanece em Supervisor e Prestador.\n\n'
        '6. REVISÃO ANTES DA ENTREGA\n'
        f'{REVIEW_NOTICE}\n'
        '- Confira se o Rank anterior correto foi usado no mapeamento.\n'
        '- Revise qualquer ocorrência marcada como Não mapeado.\n'
        '- Confira os totais dos rankings e dos painéis.\n'
        '- Expanda Supervisor e Prestador para validar os agrupamentos.\n'
        '- Abra algumas contagens por duplo clique para validar os detalhes.\n'
    )


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
        build_report_text(config, output, counts, report), encoding='utf-8'
    )
    progress(100, 'Planilha Rank concluída.')
    return RankJobResult(output, report_path, **counts)

