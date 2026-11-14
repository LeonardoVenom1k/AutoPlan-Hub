# -*- coding: utf-8 -*-
"""Regras de leitura da Planilha Mãe e geração da Planilha Rank."""
from __future__ import annotations

import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

from automacoes.planilha_mae import engine as mother_engine


HEADERS = mother_engine.STANDARD_HEADERS
SUPERVISION_HEADERS = HEADERS[:4] + ['Supervisão'] + HEADERS[4:]
MONTHS_PT = mother_engine.MONTHS_PT

PENALTY_CATEGORIES = (
    'Atrasado - Técnico informou',
    'Atrasado - Técnico não informou',
    'Atraso Penalizado',
    'No Show(Não Comparecimento do técnico)',
)

VEHICLE_CATEGORIES = (
    'Cancelada pelo cliente',
    'Improdutiva - Ausência da equipe do cliente',
    'Improdutiva - Divergência/ ausência do kit de instalação',
    'Improdutiva - Veículo disponível com falta de integridade',
    'Improdutiva - Veículo Indisponível',
    'Reagendamento - Cancelado pelo Cliente',
)

BLUE = '4472C4'
LIGHT_BLUE = 'D9EAF7'
INCONFORMITY_BLUE = '8EA9DB'
INCONFORMITY_LIGHT_BLUE = 'B8CCE4'
RED = 'FF0000'
BLACK = '000000'
WHITE = 'FFFFFF'
BORDER = '9FBAD0'
DATE_HEADERS = {'data inicio', 'data fim', 'data visita inicio', 'data visita fim'}


def _category_map(categories):
    return {mother_engine.plain(value): value for value in categories}


PENALTY_MAP = _category_map(PENALTY_CATEGORIES)
VEHICLE_MAP = _category_map(VEHICLE_CATEGORIES)


def canonical_penalty(value):
    return PENALTY_MAP.get(mother_engine.plain(value))


def canonical_vehicle(value):
    return VEHICLE_MAP.get(mother_engine.plain(value))


def is_completed(record):
    text = f"{record.get('etapa visita', '')} {record.get('status', '')}"
    return 'concluido' in mother_engine.plain(text)


def _record_from_row(headers, values, sheet_name):
    record = {'_sheet': sheet_name}
    for index, header in enumerate(headers):
        if not header or index >= len(values):
            continue
        value = values[index]
        if value not in (None, ''):
            record[header] = value
    identifier = record.get('id contrato')
    if identifier in (None, ''):
        return None
    record['_id'] = str(identifier).strip()
    return record


def _read_sheet(ws):
    header_row = None
    headers = None
    rows = ws.iter_rows(values_only=True)
    for number, values in enumerate(rows, start=1):
        normalized = [mother_engine.canonical_header(value) if value not in (None, '') else '' for value in values]
        if 'id contrato' in normalized:
            header_row = number
            headers = normalized
            break
        if number >= 25:
            break
    if header_row is None:
        return []
    records = []
    for values in rows:
        record = _record_from_row(headers, values, ws.title)
        if record:
            records.append(record)
    return records


def read_mother(path):
    """Lê abas diárias e Inconformidades sem alterar a Planilha Mãe."""
    workbook = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    daily_records = []
    inconformities = []
    dated_sheets = []
    try:
        for ws in workbook.worksheets:
            name = mother_engine.plain(ws.title)
            if name == 'inconformidades':
                inconformities.extend(_read_sheet(ws))
                continue
            date = mother_engine.date_from_filename(ws.title)
            if date:
                records = _read_sheet(ws)
                if records:
                    daily_records.extend(records)
                    dated_sheets.append(date)
    finally:
        workbook.close()
    if not dated_sheets and not inconformities:
        raise ValueError('A Planilha Mãe não contém abas diárias nem a aba Inconformidades reconhecível.')
    return daily_records, inconformities, dated_sheets


def infer_period_from_mother(path):
    daily, inconformities, dated_sheets = read_mother(path)
    dates = list(dated_sheets)
    for record in daily + inconformities:
        value = record.get('data inicio')
        if isinstance(value, datetime):
            dates.append(value)
    if not dates:
        now = datetime.now()
        return now.year, now.month
    return Counter((value.year, value.month) for value in dates).most_common(1)[0][0]


def load_supervision_mapping(previous_rank, report=None):
    """Reaproveita Prestador -> Supervisão de uma Planilha Rank anterior."""
    if not previous_rank:
        return {}
    workbook = load_workbook(previous_rank, data_only=True, read_only=True, keep_links=False)
    mapping = {}
    conflicts = set()
    try:
        sheet = next((ws for ws in workbook.worksheets
                      if mother_engine.plain(ws.title) == 'base supervisao'), None)
        if sheet is None:
            raise ValueError('A Planilha Rank anterior não contém a aba Base Supervisão.')
        for record in _read_sheet(sheet):
            provider = record.get('prestador de servico do tecnico')
            supervisor = record.get('supervisao')
            if provider in (None, '') or supervisor in (None, ''):
                continue
            key = mother_engine.plain(provider)
            old = mapping.get(key)
            if old and mother_engine.plain(old) != mother_engine.plain(supervisor):
                conflicts.add(str(provider))
                continue
            mapping[key] = str(supervisor).strip()
    finally:
        workbook.close()
    if report is not None:
        report.append(f'Mapeamento de supervisão reaproveitado: {len(mapping)} prestador(es).')
        if conflicts:
            report.append(f'AVISO: {len(conflicts)} prestador(es) tinham supervisões conflitantes; foi mantido o primeiro valor.')
    return mapping


def _overlay(base, update):
    result = dict(base)
    for key, value in update.items():
        if not key.startswith('_') and value not in (None, ''):
            result[key] = value
    return result


def build_base_rank(daily_records, inconformities):
    penalties = []
    for record in inconformities:
        category = canonical_penalty(record.get('observacoes'))
        if category:
            item = dict(record)
            item['_rank_category'] = category
            penalties.append(item)

    used_penalties = set()
    result = []
    for daily in daily_records:
        match = mother_engine.find_final_record(daily, penalties)
        if match is not None:
            record = _overlay(daily, match)
            category = match['_rank_category']
            used_penalties.add(id(match))
            record['etapa visita'] = category
            record['observacoes'] = category
            result.append(record)
            continue
        category = canonical_penalty(daily.get('observacoes'))
        if category:
            record = dict(daily)
            record['etapa visita'] = category
            record['observacoes'] = category
            result.append(record)
        elif is_completed(daily):
            result.append(dict(daily))

    for penalty in penalties:
        if id(penalty) in used_penalties:
            continue
        record = dict(penalty)
        category = penalty['_rank_category']
        record['etapa visita'] = category
        record['observacoes'] = category
        result.append(record)

    def sort_key(record):
        start = record.get('data inicio')
        return (start if isinstance(start, datetime) else datetime.max,
                str(record.get('_id', '')))
    result.sort(key=sort_key)
    return result


def build_vehicle_base(inconformities):
    result = []
    for record in inconformities:
        category = canonical_vehicle(record.get('observacoes'))
        if not category:
            continue
        item = dict(record)
        item['observacoes'] = category
        result.append(item)
    return result


def build_supervision_base(inconformities, mapping):
    result = []
    for record in inconformities:
        category = canonical_penalty(record.get('observacoes'))
        if not category:
            continue
        item = dict(record)
        item['observacoes'] = category
        provider = item.get('prestador de servico do tecnico')
        item['supervisao'] = mapping.get(mother_engine.plain(provider), 'Não mapeado')
        result.append(item)
    return result


def summarize_rank(base_rank):
    summary = defaultdict(lambda: Counter({'Total Geral': 0}))
    for record in base_rank:
        provider = str(record.get('prestador de servico do tecnico') or '(vazio)').strip()
        category = canonical_penalty(record.get('etapa visita') or record.get('observacoes'))
        summary[provider]['Total Geral'] += 1
        if category:
            summary[provider][category] += 1
    rows = []
    for provider, counts in summary.items():
        occurrences = sum(counts[category] for category in PENALTY_CATEGORIES)
        rows.append({
            'provider': provider,
            **{category: counts[category] for category in PENALTY_CATEGORIES},
            'occurrences': occurrences,
            'total': counts['Total Geral'],
        })
    best = sorted(
        (row for row in rows if row['occurrences'] == 0 and row['provider'] != '(vazio)'),
        key=lambda row: (-row['total'], mother_engine.plain(row['provider'])),
    )[:10]
    ugly = sorted(
        (row for row in rows if row['occurrences'] > 0 and row['provider'] != '(vazio)'),
        key=lambda row: (-row['occurrences'], -row['total'], mother_engine.plain(row['provider'])),
    )
    no_show = sorted(
        (row for row in ugly if row[PENALTY_CATEGORIES[3]] > 0),
        key=lambda row: (-row[PENALTY_CATEGORIES[3]], -row['occurrences'], -row['total'],
                         mother_engine.plain(row['provider'])),
    )
    return rows, best, ugly, no_show


def output_filename(year, month):
    return f'Planilha Rank - {MONTHS_PT[month]} {year} - Logística e EAB - Fechamento.xlsx'


def _value(record, header):
    return record.get(mother_engine.canonical_header(header))


def _style_data_sheet(ws, headers, records, table_name, *,
                      header_fill=BLUE, header_font=WHITE,
                      row_fills=None, table_style='TableStyleMedium2'):
    ws.append(headers)
    for record in records:
        ws.append([_value(record, header) for header in headers])
    header = ws[1]
    for cell in header:
        cell.fill = PatternFill('solid', fgColor=header_fill)
        cell.font = Font(name='Calibri', size=11, bold=True, color=header_font)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    thin = Side(style='thin', color=BORDER)
    for row_number, row in enumerate(
            ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)), start=2):
        for cell in row:
            cell.font = Font(name='Calibri', size=10)
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical='center')
            if row_fills:
                cell.fill = PatternFill('solid', fgColor=row_fills[(row_number - 2) % len(row_fills)])
    for index, header_name in enumerate(headers, start=1):
        key = mother_engine.canonical_header(header_name)
        if key in DATE_HEADERS:
            for cell in ws.iter_cols(min_col=index, max_col=index, min_row=2, max_row=ws.max_row):
                for item in cell:
                    item.number_format = 'dd/mm/yyyy hh:mm'
    widths = [14, 34, 34, 38, 17, 24, 18, 45, 38, 19, 19, 19, 19, 14, 15, 18]
    for index in range(1, len(headers) + 1):
        ws.column_dimensions[ws.cell(1, index).column_letter].width = widths[index - 1]
    ws.freeze_panes = 'A2'
    table_ref = f'A1:{ws.cell(ws.max_row, len(headers)).coordinate}'
    if records:
        # A Tabela já possui seus próprios filtros. Manter também um AutoFilter
        # sobre a mesma faixa gera um XLSX que algumas versões do Excel recusam.
        table = Table(displayName=table_name, ref=table_ref)
        if table_style:
            table.tableStyleInfo = TableStyleInfo(name=table_style, showRowStripes=True,
                                                  showFirstColumn=False, showLastColumn=False)
        ws.add_table(table)


def _write_simple_table(ws, start_row, headers, rows, table_name, widths=None,
                        header_fill=BLUE):
    for column, value in enumerate(headers, start=1):
        ws.cell(start_row, column, value)
    for offset, values in enumerate(rows, start=1):
        for column, value in enumerate(values, start=1):
            ws.cell(start_row + offset, column, value)
    end_row = start_row + max(1, len(rows))
    end_col = len(headers)
    for cell in ws[start_row]:
        if cell.column <= end_col:
            cell.fill = PatternFill('solid', fgColor=header_fill)
            cell.font = Font(name='Calibri', size=11, bold=True, color=WHITE)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[start_row].height = 32
    if rows:
        ref = f'A{start_row}:{ws.cell(end_row, end_col).coordinate}'
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(table)
    if widths:
        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(1, index).column_letter].width = width


def create_rank_workbook(mother_file, output, previous_rank='', report=None, progress=None):
    report = report if report is not None else []
    if progress:
        progress(8, 'Lendo a Planilha Mãe...')
    daily, inconformities, _ = read_mother(mother_file)
    mapping = load_supervision_mapping(previous_rank, report)
    base_rank = build_base_rank(daily, inconformities)
    vehicle_base = build_vehicle_base(inconformities)
    supervision_base = build_supervision_base(inconformities, mapping)
    _, best, ugly, no_show = summarize_rank(base_rank)

    if progress:
        progress(38, 'Montando bases e rankings...')
    workbook = Workbook()
    workbook.remove(workbook.active)
    order = [
        'Painel Inconformidades', 'Inconformidades', 'Base Rank', 'Painel Rank',
        'Rank The Best', 'Rank Ugly', 'Base Veículo Indisponível',
        'Painel Veículo Indisponível', 'Painel Supervisão', 'Base Supervisão',
    ]
    sheets = {name: workbook.create_sheet(name) for name in order}
    _style_data_sheet(
        sheets['Inconformidades'], HEADERS, inconformities, 'TabelaInconformidades',
        header_fill=WHITE, header_font=BLACK,
        row_fills=(INCONFORMITY_BLUE, INCONFORMITY_LIGHT_BLUE), table_style=None,
    )
    _style_data_sheet(sheets['Base Rank'], HEADERS, base_rank, 'TabelaBaseRank')
    _style_data_sheet(sheets['Base Veículo Indisponível'], HEADERS, vehicle_base, 'TabelaBaseVeiculo')
    _style_data_sheet(sheets['Base Supervisão'], SUPERVISION_HEADERS, supervision_base, 'TabelaBaseSupervisao')

    _write_simple_table(
        sheets['Rank The Best'], 1, ['Rótulos de Linha', 'Total Geral'],
        [(row['provider'], row['total']) for row in best], 'TabelaRankMelhores', [55, 16],
    )
    ugly_headers = ['Rótulos de Linha', *PENALTY_CATEGORIES, 'Ocorrências', 'Total Geral']
    ugly_rows = [
        (row['provider'], *(row[category] or None for category in PENALTY_CATEGORIES),
         row['occurrences'], row['total']) for row in ugly
    ]
    no_show_rows = [
        (row['provider'], row[PENALTY_CATEGORIES[3]], row['occurrences'], row['total'])
        for row in no_show
    ]
    _write_simple_table(
        sheets['Rank Ugly'], 1,
        [PENALTY_CATEGORIES[3], PENALTY_CATEGORIES[3] + '\u00a0',
         'Ocorrências', 'Total Geral'],
        no_show_rows, 'TabelaRankNoShow', [55, 36, 14, 14], header_fill=RED,
    )
    ugly_start = len(no_show_rows) + 9
    _write_simple_table(sheets['Rank Ugly'], ugly_start, ugly_headers, ugly_rows,
                        'TabelaRankOcorrencias', [55, 22, 26, 20, 36, 14, 14])
    for name in ('Painel Inconformidades', 'Painel Rank',
                 'Painel Veículo Indisponível', 'Painel Supervisão'):
        sheets[name].sheet_view.showGridLines = True
        sheets[name].column_dimensions['A'].width = 52
    for name in ('Base Rank', 'Painel Rank', 'Base Veículo Indisponível', 'Base Supervisão'):
        sheets[name].sheet_state = 'hidden'
    workbook.active = order.index('Rank The Best')
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    workbook.close()

    unmapped = sum(1 for row in supervision_base if row.get('supervisao') == 'Não mapeado')
    report.extend([
        f'Visitas consideradas no ranking: {len(base_rank)}.',
        f'Ocorrências consideradas: {len(supervision_base)}.',
        f'Registros de veículo indisponível/cliente: {len(vehicle_base)}.',
        f'Prestadores sem ocorrência no Top 10: {len(best)}.',
    ])
    if unmapped:
        report.append(f'AVISO: {unmapped} ocorrência(s) ficaram com Supervisão = Não mapeado.')
    if progress:
        progress(72, 'Arquivo-base do Rank criado.')
    return {
        'base_rank': len(base_rank), 'occurrences': len(supervision_base),
        'vehicle': len(vehicle_base), 'unmapped': unmapped,
    }


def create_native_pivots(path, report, progress=None):
    """Cria as quatro Tabelas Dinâmicas nativas no Microsoft Excel."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('Instale pywin32 para criar as Tabelas Dinâmicas nativas.') from exc
    if progress:
        progress(78, 'Criando Tabelas Dinâmicas nativas...')
    pythoncom.CoInitialize()
    excel = book = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook_path = Path(path).resolve()
        open_error = None
        for attempt in range(3):
            try:
                book = excel.Workbooks.Open(str(workbook_path), UpdateLinks=0, ReadOnly=False,
                                            IgnoreReadOnlyRecommended=True, Notify=False,
                                            AddToMru=False, Local=True)
                break
            except Exception as exc:
                open_error = exc
                if attempt < 2:
                    time.sleep(1.5)
        if book is None:
            raise RuntimeError(f'O Excel não conseguiu abrir o arquivo: {open_error}')

        xl_database = 1
        xl_row_field = 1
        xl_column_field = 2
        xl_count = -4112
        xl_descending = 2
        xl_sheet_hidden = 0
        specs = [
            ('Inconformidades', 'Painel Inconformidades', 'TabelaInconformidades',
             'PivotInconformidades', ['Observações'], [], 'Observações'),
            ('Base Rank', 'Painel Rank', 'TabelaBaseRank',
             'PivotRank', ['Prestador de Serviço do Técnico'], ['Etapa Visita'], 'Etapa Visita'),
            ('Base Veículo Indisponível', 'Painel Veículo Indisponível', 'TabelaBaseVeiculo',
             'PivotVeiculo', ['Cliente'], ['Observações'], 'Observações'),
            ('Base Supervisão', 'Painel Supervisão', 'TabelaBaseSupervisao',
             'PivotSupervisao', ['Supervisão', 'Prestador de Serviço do Técnico'],
             ['Observações'], 'Observações'),
        ]
        for source_name, panel_name, table_name, pivot_name, row_fields, column_fields, data_field in specs:
            source = book.Worksheets(source_name)
            panel = book.Worksheets(panel_name)
            panel.Visible = -1
            panel.Cells.Clear()
            if source.ListObjects.Count == 0 or source.UsedRange.Rows.Count < 2:
                panel.Range('A3').Value = 'Sem dados para o período.'
                report.append(f'AVISO: {panel_name} não recebeu Tabela Dinâmica por falta de dados.')
                continue
            table = source.ListObjects(1)
            table.Name = table_name
            cache = book.PivotCaches().Create(SourceType=xl_database, SourceData=table.Name)
            pivot = cache.CreatePivotTable(TableDestination=panel.Range('A3'), TableName=pivot_name)
            pivot.AddDataField(pivot.PivotFields(data_field), f'Contagem de {data_field}', xl_count)
            for position, field_name in enumerate(row_fields, start=1):
                field = pivot.PivotFields(field_name)
                field.Orientation = xl_row_field
                field.Position = position
            for position, field_name in enumerate(column_fields, start=1):
                field = pivot.PivotFields(field_name)
                field.Orientation = xl_column_field
                field.Position = position
            pivot.RowAxisLayout(0)
            pivot.TableStyle2 = 'PivotStyleLight16'
            pivot.ShowDrillIndicators = True
            pivot.EnableDrilldown = True
            pivot.DisplayFieldCaptions = True
            pivot.ShowTableStyleRowStripes = False
            pivot.RefreshTable()
            if panel_name == 'Painel Veículo Indisponível':
                pivot.PivotFields('Cliente').AutoSort(
                    xl_descending, f'Contagem de {data_field}'
                )
            panel.Columns.AutoFit()
            panel.Columns('A').ColumnWidth = min(max(panel.Columns('A').ColumnWidth, 24), 58)

        for name in ('Base Rank', 'Painel Rank', 'Base Veículo Indisponível', 'Base Supervisão'):
            book.Worksheets(name).Visible = xl_sheet_hidden
        book.Worksheets('Rank The Best').Activate()
        excel.ActiveWindow.DisplayGridlines = True
        book.Save()
        report.append('Tabelas Dinâmicas criadas com filtros e detalhamento por duplo clique.')
    except Exception as exc:
        raise RuntimeError(f'Não foi possível criar as Tabelas Dinâmicas do Rank: {exc}') from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

