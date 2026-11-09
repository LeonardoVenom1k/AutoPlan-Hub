# -*- coding: utf-8 -*-
"""Consolidador de Planilha Mãe do AutoPlan Hub.
Instale: python -m pip install openpyxl pywin32
Execute pelo menu: python main.py

Consolida planilhas diárias, identifica inconformidades, aplica cores por status
 e monta o Painel de contagem de observações conforme o modelo fornecido.
"""
import copy
import re
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from datetime import datetime
from collections import Counter
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

MONTHS_PT = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}
ID_NAMES = {'id cont','id contrato','id_cont','id'}
SPECIAL_SHEETS = {'painel','inconformidades'}
# Ordem e grafia do cabeçalho padrão observado na Planilha Mãe original.
STANDARD_HEADERS = ['Id contrato','Cliente','Técnico','Prestador de Serviço do Técnico','Tipo visita','Etapa Visita','Status','Observações','Informações','Data Inicio','Data Fim','Data Visita Inicio','Data Visita Fim','Placa','Frota']

# Cores aproximadas do arquivo original (RGB sem alpha).
FILL_COMPLETED = '92D050'   # verde: concluído
FILL_CANCELLED = 'FF0000'   # vermelho: cancelado
FILL_WAITING = 'F4B5F4'     # rosa: aguardando gestor / improdutiva em andamento
FILL_SCHEDULED = 'FCE4D6'   # bege: agendado
FILL_IN_PROGRESS = '5B9BD5' # azul: em andamento
FILL_INCONFORMITY_A = 'A9C2E8'
FILL_INCONFORMITY_B = 'B8CCE4'
PANEL_BLUE = 'D9EAF7'

OFFICIAL_CATEGORIES = [
    'Reagendamento - Lançado pelo supervisor',
    'Reagendamento - Cancelado pelo cliente',
    'Reagendamento - Sem tempo hábil',
    'Reagendamento - Veículo Indisponível',
    'Reagendamento - Técnico Cancelou',
    'Improdutiva - Veículo Indisponível',
    'Improdutiva - Veículo disponível com falta de integridade',
    'Improdutiva - Horário Excedido',
    'Improdutiva com necessidade de troca',
    'Improdutiva - Divergência/ ausência do kit de instalação',
    'Improdutiva - Ausência da equipe do cliente',
    'Improdutiva - Acesso não Liberado(Documentação)',
    'Improdutiva - Divergência do agendamento',
    'Cancelada pela Maxtrack',
    'Cancelada pelo cliente',
    'Não finalizado no dia',
    'Divergência de lançamento',
    'No Show(Não Comparecimento do técnico)',
    'Atraso Penalizado',
    'Atrasado - Técnico informou',
    'Atrasado - Técnico não informou',
]
DAILY_DELAY_CATEGORIES = {
    'Atrasado - Técnico informou',
    'Atrasado - Técnico não informou',
}


def norm(value):
    return re.sub(r'\s+', ' ', str(value or '').strip().lower())


def plain(value):
    """Normaliza acentos e também os caracteres quebrados dos relatórios exportados."""
    import unicodedata
    text = norm(value)
    fixes = {
        't�cnico':'tecnico', 'servi�o':'servico', 'observa��es':'observacoes',
        'informa��es':'informacoes', 'previs�o':'previsao', 'avalia��o':'avaliacao',
        'manh�':'manha', 've�culo':'veiculo', 'diverg�ncia':'divergencia',
        'aus�ncia':'ausencia', 'hor�rio':'horario', 'n�o':'nao', 'a��o':'acao',
        'lan�ado':'lancado', 'h�bil':'habil', 'instala��o':'instalacao',
        'manuten��o':'manutencao', 'desinstala��o':'desinstalacao',
        'conclu�do':'concluido',
    }
    for bad, good in fixes.items():
        text = text.replace(bad, good)
    text = ''.join(ch for ch in unicodedata.normalize('NFD', text)
                   if unicodedata.category(ch) != 'Mn')
    return re.sub(r'\s+', ' ', text).strip()


HEADER_ALIASES = {
    'id cont':'id contrato', 'id_cont':'id contrato', 'id':'id contrato',
    'tipo de visita':'tipo visita', 'etapa da visita':'etapa visita',
    'data inicio previsao':'data inicio', 'data fim previsao':'data fim',
    'identificador/placa':'placa', 'identificador placa':'placa',
    'motivo cancelamento da visita':'motivo cancelamento',
}


def canonical_header(value):
    key = plain(value)
    return HEADER_ALIASES.get(key, key)


def is_red_fill(fill):
    """Detecta preenchimento vermelho em formatos RGB ou tema/indexado."""
    color = getattr(fill, 'fgColor', None)
    if not color:
        return False
    rgb = getattr(color, 'rgb', None)
    if isinstance(rgb, str):
        rgb = rgb[-6:].upper()
        # Vermelhos comuns do Excel: vermelho puro, vermelho escuro e tons vivos.
        if rgb in {'FF0000', 'C00000', 'E60000', 'F00000', 'FF3333', 'FF6666', 'D00000'}:
            return True
        try:
            red, green, blue = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
            return red >= 180 and green <= 100 and blue <= 100 and red > green * 1.6
        except (ValueError, TypeError):
            pass
    # Índice 3 costuma representar vermelho no Excel legado.
    return getattr(color, 'type', None) == 'indexed' and getattr(color, 'indexed', None) == 3


def find_header_row(ws, limit=25):
    for r in range(1, min(ws.max_row, limit) + 1):
        if any(norm(ws.cell(r,c).value) in ID_NAMES for c in range(1, ws.max_column + 1)):
            return r
    return None


def visible_headers(ws, row):
    result = {}
    for c in range(1, ws.max_column + 1):
        letter = get_column_letter(c)
        dim = ws.column_dimensions.get(letter)
        if dim and dim.hidden:
            continue
        value = ws.cell(row,c).value
        if value is not None and str(value).strip():
            result[canonical_header(value)] = c
    return result


def find_id_col(headers):
    return next((c for h,c in headers.items() if h in ID_NAMES), None)


def date_from_filename(filename, default_year=None):
    """Aceita DD-MM, DD-MM-AA e DD-MM-AAAA."""
    m = re.search(r'(?<!\d)(\d{2})[-_/](\d{2})(?:[-_/](\d{2}|\d{4}))?(?!\d)', Path(filename).stem)
    if not m: return None
    year_text = m.group(3)
    if year_text:
        year = int(year_text)
        if year < 100: year += 2000
    else:
        year = default_year or datetime.now().year
    try: return datetime(year,int(m.group(2)),int(m.group(1)))
    except ValueError: return None


def infer_period(paths):
    explicit = [date_from_filename(p) for p in paths if re.search(r'\d{2}[-_/]\d{2}[-_/](?:\d{2}|\d{4})', Path(p).stem)]
    default_year = Counter(d.year for d in explicit if d).most_common(1)[0][0] if any(explicit) else datetime.now().year
    dates = [date_from_filename(p, default_year) for p in paths]
    dates = [d for d in dates if d]
    if not dates: return datetime.now().year, datetime.now().month
    counts = Counter((d.year,d.month) for d in dates)
    year,month = counts.most_common(1)[0][0]
    return year,month


def output_filename(year,month):
    return f'Planilha mãe - {MONTHS_PT[month]} {year} - Logísta e EAB.xlsx'


def read_daily(path):
    """Lê somente abas, linhas e colunas visíveis. Guarda todos os campos lidos
    para que campos extras (ex.: motivo de cancelamento) também ajudem na regra."""
    wb = load_workbook(path, data_only=True, read_only=False)
    found=[]
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != 'visible': continue
            hr=find_header_row(ws)
            if not hr: continue
            heads=visible_headers(ws,hr)
            idcol=find_id_col(heads)
            if not idcol: continue
            records=[]
            for r in range(hr+1,ws.max_row+1):
                if ws.row_dimensions[r].hidden: continue
                rid=ws.cell(r,idcol).value
                if rid is None or not str(rid).strip(): continue
                rec={'_id':str(rid).strip(),'_file':Path(path).name,'_sheet':ws.title}
                for h,c in heads.items():
                    cell=ws.cell(r,c)
                    v=cell.value
                    if v is not None and str(v).strip():
                        # Nas planilhas de origem, uma célula vermelha com "Agendado"
                        # indica que a visita foi cancelada. Corrige o valor antes de importar.
                        if 'agendad' in norm(v) and is_red_fill(cell.fill):
                            v='Cancelada'
                        rec[h]=v
                records.append(rec)
            if records: found.append((ws.title,hr,list(heads.items()),records))
    finally: wb.close()
    return found


def visit_key(rec):
    """Identifica uma visita sem juntar ocorrências diferentes do mesmo contrato."""
    def value(*names):
        for name in names:
            v = rec.get(name)
            if v not in (None, ''):
                if isinstance(v, datetime): return v.strftime('%Y-%m-%d %H:%M:%S')
                return plain(v)
        return ''
    return (
        value('_id', 'id contrato'),
        value('data inicio'),
        value('placa'),
        value('tipo visita'),
        value('tecnico'),
    )


def final_visit_key(rec):
    """Relaciona a visita diária com o relatório final.

    O técnico não faz parte desta chave porque o nome/equipe pode ser corrigido
    entre o agendamento e a finalização da visita.
    """
    def value(*names):
        for name in names:
            current = rec.get(name)
            if current not in (None, ''):
                if isinstance(current, datetime):
                    return current.strftime('%Y-%m-%d %H:%M:%S')
                return plain(current)
        return ''
    return (
        value('_id', 'id contrato'),
        value('data inicio'),
        value('placa'),
        value('tipo visita'),
    )


def format_arrival(value):
    """Formata o horário real no padrão usado na aba Inconformidades."""
    if isinstance(value, datetime):
        return f'Chegou às {value:%H}h{value:%M}'
    if hasattr(value, 'hour') and hasattr(value, 'minute'):
        return f'Chegou às {value.hour:02d}h{value.minute:02d}'
    if isinstance(value, str):
        match = re.search(
            r'(?<!\d)([01]?\d|2[0-3])\s*(?::|h|horas?\s+e)\s*([0-5]?\d)(?!\d)',
            plain(value),
        )
        if match:
            return f'Chegou às {int(match.group(1)):02d}h{int(match.group(2)):02d}'
    return None


def compose_information(existing=None, arrival=None, *reasons):
    """Padroniza o horário e reúne os motivos sem repetir informações."""
    parts=[]
    standardized_arrival = format_arrival(arrival) or format_arrival(existing)
    if standardized_arrival:
        parts.append(standardized_arrival)
    elif existing not in (None, ''):
        parts.append(str(existing).strip())
    for reason in reasons:
        if reason in (None, ''):
            continue
        text = str(reason).strip()
        if text and plain(text) not in {plain(part) for part in parts}:
            parts.append(text)
    return ' | '.join(parts) if parts else None


def official_category(value):
    """Converte os motivos do relatório baixado para as 21 categorias oficiais."""
    key = plain(value)
    aliases = {
        'cancelado pela maxtrack':'Cancelada pela Maxtrack',
        'cancelada pela maxtrack':'Cancelada pela Maxtrack',
        'cancelado pelo cliente':'Cancelada pelo cliente',
        'cancelada pelo cliente':'Cancelada pelo cliente',
        'veiculo indisponivel':'Improdutiva - Veículo Indisponível',
        'improdutiva - veiculo indisponivel':'Improdutiva - Veículo Indisponível',
        'veiculo disponivel com falta de integridade':'Improdutiva - Veículo disponível com falta de integridade',
        'improdutiva - veiculo disponivel com falta de integridade':'Improdutiva - Veículo disponível com falta de integridade',
        'horario excedido':'Improdutiva - Horário Excedido',
        'improdutiva - horario excedido':'Improdutiva - Horário Excedido',
        'improdutiva com necessidade de troca':'Improdutiva com necessidade de troca',
        'divergencia/ ausencia do kit de instalacao':'Improdutiva - Divergência/ ausência do kit de instalação',
        'improdutiva - divergencia/ ausencia do kit de instalacao':'Improdutiva - Divergência/ ausência do kit de instalação',
        'ausencia da equipe do cliente':'Improdutiva - Ausência da equipe do cliente',
        'improdutiva - ausencia da equipe do cliente':'Improdutiva - Ausência da equipe do cliente',
        'acesso nao liberado(documentacao)':'Improdutiva - Acesso não Liberado(Documentação)',
        'improdutiva - acesso nao liberado(documentacao)':'Improdutiva - Acesso não Liberado(Documentação)',
        'divergencia no agendamento':'Improdutiva - Divergência do agendamento',
        'improdutiva - divergencia do agendamento':'Improdutiva - Divergência do agendamento',
        'template errado':'Divergência de lançamento',
        'divergencia de lancamento':'Divergência de lançamento',
        'reagendamento - lancado pelo supervisor':'Reagendamento - Lançado pelo supervisor',
        'reagendamento - cancelado pelo cliente':'Reagendamento - Cancelado pelo cliente',
        'reagendamento - sem tempo habil':'Reagendamento - Sem tempo hábil',
        'reagendamento - veiculo indisponivel':'Reagendamento - Veículo Indisponível',
        'reagendamento - tecnico cancelou':'Reagendamento - Técnico Cancelou',
        'no show(nao comparecimento do tecnico)':'No Show(Não Comparecimento do técnico)',
        'atraso penalizado':'Atraso Penalizado',
        'atrasado - tecnico informou':'Atrasado - Técnico informou',
        'atrasado tecnico informou':'Atrasado - Técnico informou',
        'atrasado - tecnico nao informou':'Atrasado - Técnico não informou',
        'atrasado tecnico nao informou':'Atrasado - Técnico não informou',
        'nao finalizado no dia':'Não finalizado no dia',
    }
    return aliases.get(key)


def read_monthly_inconformities(path, year, month, report):
    return monthly_inconformities(read_monthly_records(path, year, month), report)


def monthly_inconformities(monthly_records, report):
    """Seleciona e classifica as inconformidades do relatório final já lido."""
    records=[]
    for rec in monthly_records:
        reason = rec.get('motivo cancelamento')
        category = official_category(reason)
        if not category and 'aguardando' in plain(rec.get('etapa visita')):
            category = 'Não finalizado no dia'
        if category:
            rec = rec.copy()
            rec['observacoes'] = category
            rec['informacoes'] = compose_information(
                rec.get('informacoes'), None, reason
            )
            records.append(rec)
        elif reason not in (None, ''):
            report.append(f'MOTIVO NÃO MAPEADO no relatório final: {reason!r} (ID {rec.get("_id", "")}).')
    return records


def read_monthly_records(path, year, month):
    """Lê todos os registros do relatório final dentro do período selecionado."""
    records=[]
    if not path: return records
    for _,_,_,rows in read_daily(path):
        for rec in rows:
            start = rec.get('data inicio')
            if isinstance(start, datetime) and (start.year, start.month) != (year, month):
                continue
            records.append(rec)
    return records


def is_inconformity(rec):
    """Critério acordado: qualquer Cancelada entra mesmo sem observações;
    também identifica atraso, improdutiva, reagendamento e divergência."""
    status=norm(rec.get('status',''))
    etapa=norm(rec.get('etapa visita',''))
    # Busca os motivos nos campos textuais da linha, inclusive campo extra da origem.
    all_text=' | '.join(norm(v) for k,v in rec.items() if not str(k).startswith('_') and v is not None)
    if 'cancelad' in status or 'cancelad' in etapa:
        return True
    keywords=('atrasad','improdutiva','reagendamento','divergência','divergencia')
    return any(word in all_text for word in keywords)


def apply_table_borders(ws, min_row=1, max_row=None, min_col=1, max_col=None):
    max_row = max_row or ws.max_row
    max_col = max_col or ws.max_column
    side = Side(style='thin', color='000000')
    border = Border(left=side, right=side, top=side, bottom=side)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = copy.copy(border)

def style_header(ws, header_row=1):
    for cell in ws[header_row]:
        cell.font=Font(bold=True,color='FFFFFF')
        cell.fill=PatternFill('solid',fgColor='4472C4')
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        cell.border=Border(bottom=Side(style='thin',color='1F1F1F'))
    ws.freeze_panes=f'A{header_row+1}'
    ws.auto_filter.ref=f'A{header_row}:{get_column_letter(max(1,ws.max_column))}{ws.max_row}'
    ws.row_dimensions[header_row].height=32
    apply_table_borders(ws, header_row, ws.max_row, 1, ws.max_column)


def fit_columns(ws, max_width=36):
    for c in range(1,ws.max_column+1):
        letter=get_column_letter(c)
        longest=0
        for r in range(1,min(ws.max_row,250)+1):
            v=ws.cell(r,c).value
            if v is not None: longest=max(longest,len(str(v)))
        ws.column_dimensions[letter].width=min(max_width,max(12,longest+2))


def order_workbook_sheets(wb):
    """Mantém dias em ordem cronológica e deixa Inconformidades/Painel no final."""
    def key(ws):
        m=re.fullmatch(r'(\d{2})-(\d{2})-(\d{4})', ws.title.strip())
        if m: return (0, int(m.group(3)), int(m.group(2)), int(m.group(1)))
        name=plain(ws.title)
        if name == 'inconformidades': return (1,0,0,0)
        if name == 'painel': return (2,0,0,0)
        return (3,0,0,0)
    wb._sheets.sort(key=key)


def create_native_pivot(path, report, progress=None):
    """Cria a Tabela Dinâmica nativa que permite filtro e duplo clique nos detalhes.

    openpyxl preserva Tabelas Dinâmicas existentes, mas não cria uma nova. Por isso
    esta etapa usa o Excel instalado no Windows através do pywin32.
    """
    try:
        import win32com.client
        import pythoncom
    except ImportError as exc:
        raise RuntimeError(
            'Para criar o Painel com filtro e detalhamento, instale: '
            'python -m pip install pywin32'
        ) from exc

    if progress: progress(97, 'Criando Tabela Dinâmica e habilitando detalhamento...')
    pythoncom.CoInitialize()
    excel=None; book=None
    try:
        excel=win32com.client.DispatchEx('Excel.Application')
        excel.Visible=False
        excel.DisplayAlerts=False
        workbook_path=Path(path).resolve()
        if not workbook_path.is_file() or workbook_path.stat().st_size == 0:
            raise FileNotFoundError(f'Arquivo de destino inválido: {workbook_path}')
        open_error=None
        for attempt in range(3):
            try:
                book=excel.Workbooks.Open(
                    str(workbook_path), UpdateLinks=0, ReadOnly=False,
                    IgnoreReadOnlyRecommended=True, Notify=False,
                    AddToMru=False, Local=True,
                )
                break
            except Exception as exc:
                open_error=exc
                if attempt < 2:
                    time.sleep(1.5)
        if book is None:
            raise RuntimeError(
                'O Excel não conseguiu abrir o arquivo após 3 tentativas. '
                'Feche a Planilha Mãe no Excel, aguarde a sincronização do OneDrive '
                f'e tente novamente. Detalhe: {open_error}'
            )
        source_ws=None
        for sheet_index in range(1, book.Worksheets.Count + 1):
            candidate=book.Worksheets(sheet_index)
            if plain(candidate.Name) == 'inconformidades':
                source_ws=candidate
                break
        if source_ws is None:
            raise RuntimeError('A aba Inconformidades não foi encontrada no arquivo.')
        if source_ws.Name != 'Inconformidades':
            source_ws.Name='Inconformidades'

        xl_up=-4162; xl_to_left=-4159; xl_src_range=1; xl_yes=1
        xl_database=1; xl_row_field=1; xl_count=-4112
        last_row=source_ws.Cells(source_ws.Rows.Count,1).End(xl_up).Row
        last_col=source_ws.Cells(1,source_ws.Columns.Count).End(xl_to_left).Column
        source_range=source_ws.Range(source_ws.Cells(1,1),source_ws.Cells(last_row,last_col))

        if source_ws.ListObjects.Count:
            table=source_ws.ListObjects(1)
            table.Resize(source_range)
            table.Name='TabelaInconformidades'
        else:
            table=source_ws.ListObjects.Add(xl_src_range,source_range,None,xl_yes)
            table.Name='TabelaInconformidades'
        table.TableStyle='TableStyleMedium2'

        # Recria somente o Painel; as demais abas e dados permanecem intactos.
        for sheet_index in range(book.Worksheets.Count, 0, -1):
            candidate=book.Worksheets(sheet_index)
            if plain(candidate.Name) == 'painel':
                candidate.Delete()
        panel=book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
        panel.Name='Painel'
        # Algumas versões do Excel ignoram a posição informada ao adicionar a aba.
        # Move explicitamente o Painel para imediatamente depois de Inconformidades.
        panel.Move(None,source_ws)
        panel=book.Worksheets('Painel')
        panel.Cells.Interior.Pattern=-4142  # sem preenchimento fora da Tabela Dinâmica

        cache=book.PivotCaches().Create(SourceType=xl_database,SourceData=table.Name)
        pivot=cache.CreatePivotTable(TableDestination=panel.Range('A3'),TableName='PainelInconformidades')
        # O mesmo campo atua como linha e como contagem. A contagem deve ser
        # adicionada primeiro; caso contrário algumas versões do Excel removem
        # o campo da área de linhas.
        pivot.AddDataField(pivot.PivotFields('Observações'),'Contagem de Observações',xl_count)
        observation=pivot.PivotFields('Observações')
        observation.Orientation=xl_row_field
        observation.Position=1
        pivot.RowAxisLayout(0)  # formato compacto: Rótulos de Linha
        pivot.TableStyle2='PivotStyleLight16'
        pivot.ShowDrillIndicators=True
        pivot.EnableDrilldown=True
        pivot.DisplayFieldCaptions=True
        pivot.ShowTableStyleRowStripes=False
        pivot.RefreshTable()

        panel.Columns('A').ColumnWidth=58
        panel.Columns('B').ColumnWidth=27
        panel.Activate()
        excel.ActiveWindow.DisplayGridlines=True
        book.Save()
        report.append('Painel criado como Tabela Dinâmica nativa; filtro e duplo clique habilitados.')
    except Exception as exc:
        raise RuntimeError(f'Não foi possível criar a Tabela Dinâmica no Excel: {exc}') from exc
    finally:
        if book is not None:
            try: book.Close(SaveChanges=True)
            except Exception: pass
        if excel is not None:
            try: excel.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()


def source_priority(rec):
    """Arquivos marcados como atualizado/tarde têm prioridade sobre os da manhã.
    O estado da visita também será comparado separadamente."""
    name = norm(rec.get('_file', ''))
    score = 10 if any(x in name for x in ('atualizado', 'atualizada', 'tarde', 'updated', 'afternoon')) else 0
    return score


def status_priority(rec):
    text = norm(rec.get('status', '')) + ' ' + norm(rec.get('etapa visita', ''))
    # Cancelamento explícito (incluindo vermelho convertido de Agendado) é terminal.
    if 'cancelad' in text: return 100
    if 'concluido' in text or 'concluído' in text: return 90
    if 'em andamento' in text: return 70
    if 'aguardando' in text: return 50
    if 'agendada' in text or 'agendado' in text: return 30
    return 10


def merge_record(target, incoming, report, day, rid):
    """Mescla versões do mesmo contrato. Dá preferência ao arquivo atualizado e
    ao estado mais avançado; campos não vazios são preservados e conflitos são relatados."""
    old_status, new_status = status_priority(target), status_priority(incoming)
    old_src, new_src = source_priority(target), source_priority(incoming)
    prefer_incoming = (new_status > old_status) or (new_status == old_status and new_src > old_src)
    # Se o registro recebido é mais recente/avançado, seus valores não vazios atualizam o alvo.
    for key, val in incoming.items():
        if key.startswith('_') or val in (None, ''): continue
        old = target.get(key)
        if old in (None, ''):
            target[key] = val
        elif str(old).strip() != str(val).strip():
            if prefer_incoming:
                report.append(f'ATUALIZAÇÃO dia {day:02d}, ID {rid}, campo {key}: {old!r} -> {val!r} ({incoming.get("_file", "")}).')
                target[key] = val
            else:
                report.append(f'CONFLITO dia {day:02d}, ID {rid}, campo {key}: mantido {old!r}, ignorado {val!r} ({incoming.get("_file", "")}).')
    # Guarda origem prioritária para que a próxima comparação use a mesma regra.
    if prefer_incoming:
        target['_file'] = incoming.get('_file', target.get('_file', ''))
    return target


def row_fill_for_record(rec):
    status = norm(rec.get('status', ''))
    etapa = norm(rec.get('etapa visita', ''))
    text = status + ' ' + etapa
    if 'cancelad' in text: color = FILL_CANCELLED
    elif 'concluido' in text or 'concluído' in text: color = FILL_COMPLETED
    elif 'em andamento' in text: color = FILL_IN_PROGRESS
    elif 'aguardando gestor' in text or 'aguardando' in text: color = FILL_WAITING
    elif 'agendad' in text: color = FILL_SCHEDULED
    else: color = None
    return PatternFill('solid', fgColor=color) if color else None


def apply_day_row_style(ws, row, rec):
    fill = row_fill_for_record(rec)
    if fill:
        for cell in ws[row]:
            cell.fill = copy.copy(fill)
            if fill.fgColor.rgb and fill.fgColor.rgb[-6:].upper() == FILL_CANCELLED:
                cell.font = copy.copy(cell.font)
                cell.font = Font(name=cell.font.name or 'Calibri', size=cell.font.sz or 11, bold=cell.font.bold, color='FFFFFF')
            else:
                cell.font = copy.copy(cell.font)
                cell.font = Font(name=cell.font.name or 'Calibri', size=cell.font.sz or 11, bold=cell.font.bold, color='000000')
            side = Side(style='thin', color='000000')
            cell.border = Border(left=side, right=side, top=side, bottom=side)

def merge_daily_records(paths, year, month, report):
    """Retorna dados agrupados por dia e mescla campos não vazios por Id."""
    by_day={}
    for path in paths:
        d=date_from_filename(path, year)
        if not d:
            report.append(f'DATA NÃO IDENTIFICADA: {Path(path).name}; ignorado.')
            continue
        if (d.year,d.month)!=(year,month):
            report.append(f'FORA DO MÊS: {Path(path).name}; ignorado.')
            continue
        candidates=read_daily(path)
        if not candidates:
            report.append(f'REGISTROS NÃO ENCONTRADOS: {Path(path).name}; confira cabeçalho/colunas visíveis.')
            continue
        day_records=by_day.setdefault(d.day,{})
        for _,_,_,records in candidates:
            for rec in records:
                rid=rec['_id']; key=visit_key(rec)
                if key not in day_records: day_records[key]=rec.copy()
                else:
                    merge_record(day_records[key], rec, report, d.day, rid)
    return by_day


def value_for_header(rec, header):
    return rec.get(canonical_header(header))


def write_day_sheet(ws, records):
    ws.append(STANDARD_HEADERS)
    style_header(ws)
    for rec in records:
        ws.append([value_for_header(rec,h) for h in STANDARD_HEADERS])
        apply_day_row_style(ws, ws.max_row, rec)
    apply_table_borders(ws)
    fit_columns(ws)


def rebuild_inconformities(wb, final_records=None, final_all_records=None,
                           report=None, daily_source_records=None):
    """Une o relatório mensal final aos dois tipos de atraso anotados no dia."""
    report = report if report is not None else []
    final_index = {final_visit_key(rec): rec for rec in final_all_records or []}
    daily_source_index = {
        final_visit_key(rec): rec for rec in daily_source_records or []
    }
    existing_name = next((n for n in wb.sheetnames if plain(n) == 'inconformidades'), None)
    if existing_name:
        ws=wb[existing_name]
        ws.title='Inconformidades'
        # Remove old rows while retaining sheet object where possible.
        if ws.max_row: ws.delete_rows(1,ws.max_row)
    else: ws=wb.create_sheet('Inconformidades')
    ws.append(STANDARD_HEADERS)
    style_header(ws)

    combined=[]; seen=set()
    for rec in final_records or []:
        category=official_category(rec.get('observacoes'))
        if not category: continue
        key=(visit_key(rec), plain(category))
        if key not in seen:
            seen.add(key); rec=rec.copy(); rec['observacoes']=category; combined.append(rec)

    for day_ws in wb.worksheets:
        if plain(day_ws.title) in SPECIAL_SHEETS: continue
        hr=find_header_row(day_ws)
        if not hr: continue
        heads=visible_headers(day_ws,hr)
        for r in range(hr+1,day_ws.max_row+1):
            values={h:day_ws.cell(r,c).value for h,c in heads.items()}
            if not any(v not in (None,'') for v in values.values()): continue
            category=official_category(values.get('observacoes'))
            if category not in DAILY_DELAY_CATEGORIES: continue
            values['_id']=str(values.get('id contrato','')).strip()
            source_rec = daily_source_index.get(final_visit_key(values), {})
            daily_reason = source_rec.get('motivo cancelamento')
            final_rec = final_index.get(final_visit_key(values))
            if final_rec:
                # Somente a cópia enviada para Inconformidades recebe o estado
                # final; a linha original da aba diária permanece intocada.
                if 'agendad' in plain(values.get('etapa visita')):
                    final_stage = final_rec.get('etapa visita') or final_rec.get('status')
                    if final_stage not in (None, ''):
                        values['etapa visita'] = final_stage
                if 'agendad' in plain(values.get('status')):
                    final_status = final_rec.get('status') or final_rec.get('etapa visita')
                    if final_status not in (None, ''):
                        values['status'] = final_status
                arrival_value = final_rec.get('data visita inicio')
                arrival = format_arrival(arrival_value)
                values['informacoes'] = compose_information(
                    values.get('informacoes'), arrival_value,
                    daily_reason, final_rec.get('motivo cancelamento')
                )
                if not arrival:
                    report.append(
                        f'HORÁRIO DE CHEGADA NÃO ENCONTRADO no relatório final '
                        f'(ID {values.get("_id", "")}, aba {day_ws.title}).'
                    )
            else:
                values['informacoes'] = compose_information(
                    values.get('informacoes'), None, daily_reason
                )
                report.append(
                    f'VISITA ATRASADA NÃO ENCONTRADA no relatório final '
                    f'(ID {values.get("_id", "")}, aba {day_ws.title}).'
                )
            key=(visit_key(values), plain(category))
            if key not in seen:
                seen.add(key); values['observacoes']=category; combined.append(values)

    for rec in combined:
        ws.append([value_for_header(rec,h) for h in STANDARD_HEADERS])
        row = ws.max_row
        fill = PatternFill('solid', fgColor=FILL_INCONFORMITY_A if row % 2 == 0 else FILL_INCONFORMITY_B)
        for cell in ws[row]:
            cell.fill = copy.copy(fill)
            cell.font = Font(name='Calibri', size=11, color='000000')
    apply_table_borders(ws)
    fit_columns(ws)


def canonical_text(value):
    return plain(value)


def rebuild_panel(wb):
    """Cria o painel no formato de resumo mostrado pelo usuário: rótulos e
    contagem das Observações na aba Inconformidades."""
    if 'Painel' in wb.sheetnames:
        ws=wb['Painel']
        if ws.max_row: ws.delete_rows(1,ws.max_row)
    else:
        ws=wb.create_sheet('Painel')

    categories = OFFICIAL_CATEGORIES
    counts=Counter()
    if 'Inconformidades' in wb.sheetnames:
        inc=wb['Inconformidades']
        hr=find_header_row(inc)
        if hr:
            heads=visible_headers(inc,hr)
            obs_col=heads.get(canonical_header('Observações'))
            if obs_col:
                for r in range(hr+1,inc.max_row+1):
                    value=inc.cell(r,obs_col).value
                    if value is not None and str(value).strip():
                        counts[canonical_text(value)] += 1

    # Reproduz a posição visual do painel original (título dos campos na linha 3).
    ws.sheet_view.showGridLines = True
    ws['A3']='Rótulos de Linha'
    ws['B3']='Contagem de Observações'
    for row, label in enumerate(categories, start=4):
        ws.cell(row,1).value=label
        ws.cell(row,2).value=counts.get(canonical_text(label),0)
    total_row=4+len(categories)
    ws.cell(total_row,1).value='Total Geral'
    ws.cell(total_row,2).value=sum(counts.get(canonical_text(label),0) for label in categories)

    header_fill=PatternFill('solid',fgColor='D9EAF7')
    total_fill=PatternFill('solid',fgColor='D9EAF7')
    # O Painel é a única aba sem bordas: visual limpo como no modelo enviado.
    # Remove bordas antigas também quando estamos atualizando uma Mãe existente.
    for row_cells in ws.iter_rows(min_row=1, max_row=max(ws.max_row, total_row), min_col=1, max_col=max(ws.max_column, 2)):
        for cell in row_cells:
            cell.border = Border()
    for cell in ws[3][:2]:
        cell.font=Font(bold=True, color='000000')
        cell.fill=header_fill
        cell.alignment=Alignment(vertical='center')
    for cell in ws[total_row][:2]:
        cell.font=Font(bold=True, color='000000')
        cell.fill=total_fill
    ws.column_dimensions['A'].width=58
    ws.column_dimensions['B'].width=27
    for r in range(4,total_row+1):
        ws.cell(r,2).alignment=Alignment(horizontal='right')
    # Sem bordas no Painel, inclusive nas linhas de dados e no Total Geral.
    ws.freeze_panes='A4'


def create_new_mother(daily_paths, final_path, output, year, month, report, progress=None):
    if progress: progress(10, 'Lendo e comparando planilhas diárias...')
    wb=Workbook(); wb.remove(wb.active)
    by_day=merge_daily_records(daily_paths,year,month,report)
    final_all_records=read_monthly_records(final_path,year,month)
    final_records=monthly_inconformities(final_all_records,report)
    days=sorted(by_day)
    for idx, day in enumerate(days, start=1):
        ws=wb.create_sheet(f'{day:02d}-{month:02d}-{year}')
        write_day_sheet(ws,by_day[day].values())
        if progress: progress(10 + int(55*idx/max(1,len(days))), f'Criando e formatando aba {ws.title}...')
    if not wb.sheetnames:
        ws=wb.create_sheet(f'01-{month:02d}-{year}')
        write_day_sheet(ws,[])
        report.append('Nenhuma planilha diária válida foi importada.')
    if progress: progress(72, 'Identificando e montando Inconformidades...')
    daily_source_records=[rec for records in by_day.values() for rec in records.values()]
    rebuild_inconformities(
        wb, final_records, final_all_records, report, daily_source_records
    )
    if progress: progress(86, 'Montando Painel e contagens...')
    rebuild_panel(wb)
    order_workbook_sheets(wb)
    if progress: progress(95, 'Salvando arquivo Excel...')
    wb.save(output); wb.close()


def merge_into_existing(wb, daily_paths, final_path, year, month, report, progress=None):
    if progress: progress(10, 'Lendo e comparando planilhas diárias...')
    by_day=merge_daily_records(daily_paths,year,month,report)
    final_all_records=read_monthly_records(final_path,year,month)
    final_records=monthly_inconformities(final_all_records,report)
    added=0
    day_items=list(by_day.items())
    for day_index, (day,records) in enumerate(day_items, start=1):
        title=f'{day:02d}-{month:02d}-{year}'
        ws=next((s for s in wb.worksheets if norm(s.title)==norm(title)),None)
        if ws is None:
            ws=wb.create_sheet(title)
            write_day_sheet(ws,records.values())
            added+=len(records)
            continue
        hr=find_header_row(ws)
        if not hr:
            report.append(f'ABA {ws.title} sem cabeçalho reconhecível; criada não foi alterada.')
            continue
        heads=visible_headers(ws,hr); idcol=find_id_col(heads)
        if not idcol:
            report.append(f'ABA {ws.title} sem coluna de ID; não alterada.')
            continue
        existing={}
        for r in range(hr+1,ws.max_row+1):
            v=ws.cell(r,idcol).value
            if v is not None and str(v).strip():
                row_rec={h:ws.cell(r,c).value for h,c in heads.items()}
                row_rec['_id']=str(v).strip()
                existing[visit_key(row_rec)]=r
        # Cabeçalho destino pode ter grafia/capitalização diferente.
        for key,rec in records.items():
            rid=rec['_id']
            if key in existing:
                tr=existing[key]
                existing_rec = {'_file': 'Planilha Mãe existente'}
                for h,c in heads.items():
                    existing_rec[h] = ws.cell(tr,c).value
                # Mãe existente é base; o registro atualizado pode atualizar status/etapa e campos alterados.
                merged = merge_record(existing_rec, rec, report, day, rid)
                for h,c in heads.items():
                    if h in merged and h not in ID_NAMES and merged[h] not in (None, ''):
                        ws.cell(tr,c).value = merged[h]
                apply_day_row_style(ws, tr, {norm(k):v for k,v in merged.items() if not k.startswith('_')})
            else:
                tr=ws.max_row+1
                for c in range(1,ws.max_column+1):
                    src=ws.cell(max(hr+1,tr-1),c); dst=ws.cell(tr,c)
                    if src.has_style: dst._style=copy.copy(src._style)
                    if src.number_format: dst.number_format=src.number_format
                for h,c in heads.items():
                    if h in rec: ws.cell(tr,c).value=rec[h]
                apply_day_row_style(ws, tr, {norm(k):v for k,v in rec.items() if not k.startswith('_')})
                existing[key]=tr; added+=1
        if progress: progress(10 + int(55*day_index/max(1,len(day_items))), f'Atualizando e formatando aba {title}...')
    if progress: progress(72, 'Identificando e montando Inconformidades...')
    # Aplica bordas em todas as células usadas, inclusive registros antigos da Mãe.
    for day_ws in wb.worksheets:
        if norm(day_ws.title) not in SPECIAL_SHEETS and find_header_row(day_ws):
            apply_table_borders(day_ws)
    daily_source_records=[rec for records in by_day.values() for rec in records.values()]
    rebuild_inconformities(
        wb, final_records, final_all_records, report, daily_source_records
    )
    if progress: progress(86, 'Montando Painel e contagens...')
    rebuild_panel(wb)
    order_workbook_sheets(wb)
    return added


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('AutoPlan Hub - Consolidador de Planilha Mãe')
        self.geometry('900x700')
        self.minsize(820, 620)
        self.configure(bg='#F3F6FA')
        style = ttk.Style(self)
        try: style.theme_use('clam')
        except tk.TclError: pass
        style.configure('TFrame', background='#F3F6FA')
        style.configure('TLabel', background='#F3F6FA', font=('Segoe UI', 10))
        style.configure('TLabelframe', background='#F3F6FA')
        style.configure('TLabelframe.Label', background='#F3F6FA', font=('Segoe UI', 10, 'bold'))
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('Accent.TButton', font=('Segoe UI', 11, 'bold'), padding=9)
        self.mode=tk.StringVar(value='create'); self.reference=tk.StringVar(); self.final_file=tk.StringVar(); self.output_dir=tk.StringVar(); self.year=tk.StringVar(); self.month=tk.StringVar()
        self.daily_files=[]
        ttk.Label(self,text='CONSOLIDADOR DE PLANILHA MÃE',font=('Segoe UI',18,'bold')).pack(pady=(14,3))
        ttk.Label(self,text='Crie uma Planilha Mãe ou atualize uma Planilha Mãe existente.',wraplength=800).pack(pady=(0,8))
        mode=ttk.LabelFrame(self,text='Modo de operação',padding=8); mode.pack(fill='x',padx=14,pady=4)
        ttk.Radiobutton(mode,text='Criar Planilha Mãe do zero',variable=self.mode,value='create',command=self.refresh_mode).grid(row=0,column=0,sticky='w',padx=5)
        ttk.Radiobutton(mode,text='Atualizar Planilha Mãe existente',variable=self.mode,value='update',command=self.refresh_mode).grid(row=0,column=1,sticky='w',padx=5)
        self.ref_frame=ttk.Frame(self,padding=(14,4))
        ttk.Label(self.ref_frame,text='Planilha Mãe para atualizar:').grid(row=0,column=0,sticky='w')
        self.reference_entry=ttk.Entry(self.ref_frame,textvariable=self.reference)
        self.reference_entry.grid(row=1,column=0,sticky='ew',pady=3)
        ttk.Button(self.ref_frame,text='Selecionar...',command=self.pick_reference).grid(row=1,column=1,padx=5)
        self.ref_frame.columnconfigure(0,weight=1)
        self.period_frame=ttk.Frame(self,padding=(14,4)); self.period_frame.pack(fill='x')
        ttk.Label(self.period_frame,text='Mês:').grid(row=0,column=0,sticky='w')
        self.month_combo=ttk.Combobox(self.period_frame,textvariable=self.month,values=list(MONTHS_PT.values()),state='readonly',width=15); self.month_combo.grid(row=0,column=1,padx=5)
        ttk.Label(self.period_frame,text='Ano:').grid(row=0,column=2,sticky='w')
        ttk.Entry(self.period_frame,textvariable=self.year,width=8).grid(row=0,column=3,padx=5)
        ttk.Button(self.period_frame,text='Inferir mês/ano',command=self.infer_dates).grid(row=0,column=4,padx=8)
        self.files_frame=ttk.Frame(self,padding=(14,4)); self.files_frame.pack(fill='x')
        ttk.Label(self.files_frame,text='Planilhas diárias (selecione várias):').grid(row=0,column=0,sticky='w')
        self.files_count=ttk.Label(self.files_frame,text='Nenhum arquivo selecionado')
        self.files_count.grid(row=0,column=1,sticky='e',padx=5)
        list_frame=ttk.Frame(self.files_frame)
        list_frame.grid(row=1,column=0,sticky='nsew',pady=3)
        self.files_list=tk.Listbox(list_frame,height=4,font=('Segoe UI',9),selectmode='extended',
                                   activestyle='none',exportselection=False)
        files_scroll=ttk.Scrollbar(list_frame,orient='vertical',command=self.files_list.yview)
        self.files_list.configure(yscrollcommand=files_scroll.set)
        self.files_list.pack(side='left',fill='both',expand=True)
        files_scroll.pack(side='right',fill='y')
        ttk.Button(self.files_frame,text='Selecionar planilhas...',command=self.pick_daily).grid(row=1,column=1,padx=5,sticky='n')
        ttk.Label(self.files_frame,text='Relatório final do mês (fonte das inconformidades):').grid(row=2,column=0,sticky='w',pady=(7,0))
        ttk.Entry(self.files_frame,textvariable=self.final_file).grid(row=3,column=0,sticky='ew',pady=3)
        ttk.Button(self.files_frame,text='Selecionar relatório...',command=self.pick_final).grid(row=3,column=1,padx=5)
        self.output_label=ttk.Label(self.files_frame,text='Pasta de saída:')
        self.output_label.grid(row=4,column=0,sticky='w',pady=(7,0))
        self.output_entry=ttk.Entry(self.files_frame,textvariable=self.output_dir)
        self.output_entry.grid(row=5,column=0,sticky='ew',pady=3)
        self.output_button=ttk.Button(self.files_frame,text='Selecionar pasta...',command=self.pick_output_dir)
        self.output_button.grid(row=5,column=1,padx=5)
        self.files_frame.columnconfigure(0,weight=1)
        self.progress=ttk.Progressbar(self,maximum=100); self.progress.pack(fill='x',padx=18,pady=5)
        log_frame=ttk.LabelFrame(self,text='Andamento do processamento',padding=6); log_frame.pack(fill='both',expand=True,padx=18,pady=6)
        self.log=tk.Text(log_frame,height=9,wrap='word',font=('Consolas',9),bg='#17212F',fg='#E6EDF3',insertbackground='white',relief='flat',padx=8,pady=6)
        self.log.tag_configure('ok',foreground='#7EE787'); self.log.tag_configure('warn',foreground='#F2CC60'); self.log.tag_configure('info',foreground='#79C0FF'); self.log.tag_configure('error',foreground='#FF7B72')
        self.log.pack(fill='both',expand=True)
        self.progress_label=ttk.Label(self,text='Pronto para iniciar')
        self.progress_label.pack(anchor='w',padx=20)
        ttk.Button(self,text='EXECUTAR PROCESSAMENTO',style='Accent.TButton',command=self.run).pack(pady=9)
        now=datetime.now(); self.month.set(MONTHS_PT[now.month]); self.year.set(str(now.year)); self.output_dir.set(str(Path.home()/'Documents'))
        self.refresh_mode()

    def refresh_mode(self):
        if self.mode.get()=='create':
            self.ref_frame.pack_forget()
            self.output_label.grid()
            self.output_entry.grid()
            self.output_button.grid()
        else:
            if not self.ref_frame.winfo_manager():
                self.ref_frame.pack(fill='x',before=self.period_frame)
            self.output_label.grid_remove()
            self.output_entry.grid_remove()
            self.output_button.grid_remove()
            self.after_idle(self.reference_entry.focus_set)

    def pick_reference(self):
        p=filedialog.askopenfilename(title='Selecione a Planilha Mãe',filetypes=[('Excel','*.xlsx *.xlsm')])
        if p: self.reference.set(p)

    def pick_daily(self):
        ps=filedialog.askopenfilenames(title='Selecione planilhas diárias',filetypes=[('Excel','*.xlsx *.xlsm')])
        if ps:
            self.daily_files=list(ps)
            self.files_list.delete(0,'end')
            for path in ps:
                self.files_list.insert('end',Path(path).name)
            self.files_count.config(text=f'{len(ps)} arquivo(s) selecionado(s)')
            self.infer_dates()

    def pick_final(self):
        p=filedialog.askopenfilename(title='Selecione o relatório final do mês',filetypes=[('Excel','*.xlsx *.xlsm')])
        if p: self.final_file.set(p)

    def infer_dates(self):
        if not self.daily_files:
            messagebox.showinfo('Selecionar arquivos','Selecione as planilhas diárias primeiro.'); return
        y,m=infer_period(self.daily_files); self.year.set(str(y)); self.month.set(MONTHS_PT[m])

    def pick_output_dir(self):
        p=filedialog.askdirectory(title='Selecione a pasta de saída')
        if p: self.output_dir.set(p)

    def say(self,msg,tag='info'):
        self.log.insert('end',msg+'\n',tag); self.log.see('end'); self.update_idletasks()

    def update_progress(self, value, message):
        self.progress['value']=value
        self.progress_label.config(text=f'{value}% — {message}')
        self.say(message, 'info')
        self.update_idletasks()

    def run(self):
        if not self.daily_files:
            messagebox.showerror('Faltam arquivos','Selecione pelo menos uma planilha diária.'); return
        final_path=Path(self.final_file.get().strip())
        if not final_path.is_file():
            messagebox.showerror('Relatório final','Selecione o relatório final do mês.'); return
        try:
            year=int(self.year.get()); month=next(k for k,v in MONTHS_PT.items() if v==self.month.get())
            if not 2000<=year<=2100: raise ValueError()
        except Exception:
            messagebox.showerror('Período inválido','Selecione mês e ano válidos.'); return
        if self.mode.get()=='create':
            folder=Path(self.output_dir.get().strip())
            if not folder.is_dir():
                messagebox.showerror('Pasta inválida','Selecione uma pasta de saída válida.'); return
            output=folder/output_filename(year,month)
            if output.exists() and not messagebox.askyesno('Arquivo já existe',f'{output.name} já existe. Deseja substituir?'): return
        else:
            output=Path(self.reference.get().strip())
            if not output.is_file():
                messagebox.showerror('Planilha Mãe','Selecione uma Planilha Mãe existente válida.'); return
            folder=output.parent
        report=[]; self.log.delete('1.0','end'); self.progress['value']=0; self.progress_label.config(text='Iniciando...')
        self.say('Iniciando processamento...', 'info')
        try:
            if self.mode.get()=='create':
                create_new_mother(self.daily_files,str(final_path),str(output),year,month,report,self.update_progress)
                self.say(f'Criada: {output}', 'ok')
            else:
                backup=output.with_name(f'{output.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{output.suffix}')
                shutil.copy2(output,backup)
                report.append(f'Backup criado antes da atualização: {backup.name}')
                self.say(f'Backup de segurança: {backup}', 'ok')
                wb=load_workbook(output,keep_vba=output.suffix.lower()=='.xlsm')
                added=merge_into_existing(wb,self.daily_files,str(final_path),year,month,report,self.update_progress)
                wb.save(output); wb.close()
                self.say(f'Planilha Mãe atualizada. Novos registros adicionados: {added}', 'ok')
            create_native_pivot(output,report,self.update_progress)
            rp=folder/(output.stem+'_relatorio.txt')
            with open(rp,'w',encoding='utf-8') as f:
                f.write(f'Relatório: {datetime.now():%d/%m/%Y %H:%M:%S}\nArquivo: {output}\nPeríodo: {self.month.get()} {year}\nModo: {self.mode.get()}\n\n')
                f.write('\n'.join(report) if report else 'Nenhum conflito ou aviso registrado.')
            self.progress['value']=100; self.progress_label.config(text='100% — Concluído'); self.say(f'Relatório: {rp}', 'ok')
            messagebox.showinfo('Concluído',f'Arquivo criado/atualizado:\n{output}\n\nRelatório:\n{rp}\n\nA aba Inconformidades e o Painel foram recalculados. Confira o resultado antes de usar como versão oficial.')
        except Exception as e:
            self.say(f'ERRO: {e}', 'error')
            messagebox.showerror('Erro',f'Falha no processamento:\n{e}')

if __name__=='__main__':
    App().mainloop()
