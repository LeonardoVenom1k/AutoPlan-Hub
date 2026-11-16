from pathlib import Path
from datetime import date
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / 'docs' / 'Documentacao_AutoPlan_Hub.docx'
ICON = ROOT / 'assets' / 'autoplan-hub-icon.png'

NAVY = '17365D'
BLUE = '2F75B5'
PALE_BLUE = 'D9EAF7'
LIGHT_BLUE = 'EEF5FB'
LIGHT_GRAY = 'F3F4F6'
MID_GRAY = 'D9D9D9'
TEXT = '1F2937'
MUTED = '64748B'
WHITE = 'FFFFFF'


def set_cell_fill(cell, color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tc_pr.append(shd)
    shd.set(qn('w:fill'), color)


def set_cell_margins(cell, top=110, start=120, bottom=110, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in('w:tcMar')
    if tc_mar is None:
        tc_mar = OxmlElement('w:tcMar')
        tc_pr.append(tc_mar)
    for name, value in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tc_mar.find(qn(f'w:{name}'))
        if node is None:
            node = OxmlElement(f'w:{name}')
            tc_mar.append(node)
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')


def set_table_borders(table, color=MID_GRAY, size='6'):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn('w:tblBorders'))
    if borders is None:
        borders = OxmlElement('w:tblBorders')
        tbl_pr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        tag = borders.find(qn(f'w:{edge}'))
        if tag is None:
            tag = OxmlElement(f'w:{edge}')
            borders.append(tag)
        tag.set(qn('w:val'), 'single')
        tag.set(qn('w:sz'), size)
        tag.set(qn('w:color'), color)


def repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement('w:tblHeader')
    tbl_header.set(qn('w:val'), 'true')
    tr_pr.append(tbl_header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement('w:cantSplit')
    tr_pr.append(cant_split)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run('Página ')
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor.from_string(MUTED)
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    paragraph._p.append(fld)


def add_field(paragraph, instruction):
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), instruction)
    paragraph._p.append(fld)


def style_table(table, widths=None, header=True):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    if header:
        repeat_table_header(table.rows[0])
    for r_idx, row in enumerate(table.rows):
        prevent_row_split(row)
        if r_idx == 0 and header:
            for cell in row.cells:
                set_cell_fill(cell, NAVY)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(0)
                    for run in p.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.size = Pt(9)
        else:
            if r_idx % 2 == 0:
                for cell in row.cells:
                    set_cell_fill(cell, LIGHT_BLUE)
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.05
                    for run in p.runs:
                        run.font.size = Pt(9)
        for cell in row.cells:
            set_cell_margins(cell)
    if widths:
        for row in table.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = width


def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.paragraph_format.keep_with_next = True
    return p


def add_body(doc, text, bold_lead=None):
    p = doc.add_paragraph()
    if bold_lead and text.startswith(bold_lead):
        p.add_run(bold_lead).bold = True
        p.add_run(text[len(bold_lead):])
    else:
        p.add_run(text)
    return p


def add_bullets(doc, items, level=0):
    for item in items:
        p = doc.add_paragraph(style='List Bullet' if level == 0 else 'List Bullet 2')
        p.add_run(item)


def add_numbered(doc, items):
    for number, item in enumerate(items, start=1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.65)
        p.paragraph_format.first_line_indent = Cm(-0.65)
        p.add_run(f'{number}.').bold = True
        p.add_run(f'  {item}')


def configure_document(doc):
    doc.settings.odd_and_even_pages_header_footer = False
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    # Apresentação inspirada na ABNT NBR 14724: A4, 3 cm à esquerda e
    # no topo, 2 cm à direita e na base.
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

    normal = doc.styles['Normal']
    normal.font.name = 'Arial'
    normal._element.rPr.rFonts.set(qn('w:ascii'), 'Arial')
    normal._element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial')
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor.from_string(TEXT)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    title = doc.styles['Title']
    title.font.name = 'Arial'
    title._element.rPr.rFonts.set(qn('w:ascii'), 'Arial')
    title._element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial')
    title.font.size = Pt(30)
    title.font.bold = True
    title.font.color.rgb = RGBColor(0, 0, 0)
    title_ppr = title.element.get_or_add_pPr()
    title_border = title_ppr.find(qn('w:pBdr'))
    if title_border is not None:
        title_ppr.remove(title_border)

    for name, size in [('Heading 1', 18), ('Heading 2', 13), ('Heading 3', 11)]:
        style = doc.styles[name]
        style.font.name = 'Arial'
        style._element.rPr.rFonts.set(qn('w:ascii'), 'Arial')
        style._element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial')
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(14 if name == 'Heading 1' else 9)
        style.paragraph_format.space_after = Pt(6)

    for name in ('TOC 1', 'TOC 2', 'TOC 3'):
        if name in doc.styles:
            style = doc.styles[name]
            style.font.name = 'Arial'
            style._element.rPr.rFonts.set(qn('w:ascii'), 'Arial')
            style._element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial')
            style.font.size = Pt(10)
            style.paragraph_format.line_spacing = 1.0
            style.paragraph_format.space_after = Pt(2)

    header = section.header
    hp = header.paragraphs[0]
    hp.text = 'AUTOPLAN HUB   |   DOCUMENTAÇÃO DO PROJETO'
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in hp.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(NAVY)

    footer = section.footer
    table = footer.add_table(rows=1, cols=2, width=Cm(16.6))
    table.autofit = False
    table.columns[0].width = Cm(12.5)
    table.columns[1].width = Cm(4.1)
    left = table.cell(0, 0).paragraphs[0]
    left.add_run('Leonardo Aguiar   |   AutoPlan Hub 0.6.1').font.size = Pt(8.5)
    left.runs[0].font.color.rgb = RGBColor.from_string(MUTED)
    add_page_number(table.cell(0, 1).paragraphs[0])
    set_table_borders(table, color='FFFFFF', size='0')

    # Capa limpa; cabeçalho e rodapé são usados a partir da segunda página.
    first_header = section.first_page_header
    first_header.paragraphs[0].text = ''
    first_footer = section.first_page_footer
    first_footer.paragraphs[0].text = ''


def build():
    doc = Document()
    configure_document(doc)
    props = doc.core_properties
    props.title = 'Documentação do projeto AutoPlan Hub'
    props.subject = 'Manual técnico operacional e histórico de versões'
    props.author = 'Leonardo Aguiar'
    props.keywords = 'AutoPlan Hub, automação, planilha mãe, documentação, versionamento'

    # Capa
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(35)
    p.add_run().add_picture(str(ICON), width=Inches(1.55))
    title = doc.add_paragraph(style='Title')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run('Documentação do projeto AutoPlan Hub')
    direct_ppr = title._p.get_or_add_pPr()
    direct_border = direct_ppr.find(qn('w:pBdr'))
    if direct_border is not None:
        direct_ppr.remove(direct_border)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('Manual técnico operacional e histórico de versões')
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor.from_string(MUTED)

    doc.add_paragraph()
    cover = doc.add_table(rows=4, cols=2)
    cover_data = [
        ('Versão do software', '0.6.1'),
        ('Versão do documento', '3.1'),
        ('Responsável', 'Leonardo Aguiar'),
        ('Data de emissão', '16 de novembro de 2026'),
    ]
    for row, values in zip(cover.rows, cover_data):
        row.cells[0].text, row.cells[1].text = values
        row.cells[0].paragraphs[0].runs[0].bold = True
    style_table(cover, [Cm(6), Cm(8)], header=False)
    for idx, row in enumerate(cover.rows):
        if idx % 2 == 0:
            for cell in row.cells:
                set_cell_fill(cell, LIGHT_BLUE)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    r = p.add_run('Código público e dados operacionais protegidos')
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(NAVY)
    doc.add_page_break()

    # Folha de rosto
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(25)
    r = p.add_run('LEONARDO AGUIAR')
    r.bold = True
    r.font.size = Pt(12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(80)
    r = p.add_run('AUTOPLAN HUB')
    r.bold = True
    r.font.size = Pt(18)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('Manual técnico operacional e de conformidade')
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(7)
    p.paragraph_format.space_before = Pt(70)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run('Documentação destinada à operação, manutenção, controle de versões e uso responsável do AutoPlan Hub em ambiente corporativo.')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(110)
    p.add_run('BRASIL\n2026')
    doc.add_page_break()

    add_heading(doc, 'Apresentação', 1)
    add_body(doc, 'Este documento registra a finalidade, a arquitetura, as regras de funcionamento, os procedimentos de uso e o histórico de versões do AutoPlan Hub. Ele serve como referência para operação, manutenção, evolução e publicação controlada do projeto em um repositório público.')
    add_body(doc, 'Repositório oficial: https://github.com/LeonardoVenom1k/AutoPlan-Hub')
    add_body(doc, 'A versão 0.6.1 separa as ações Abrir e Configurar no menu. Abrir usa imediatamente as preferências salvas. Configurar permite escolher aplicativos, sites e indicar que o Chrome já restaura as abas da sessão anterior, evitando a abertura repetida dos mesmos links.')

    doc.add_page_break()
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run('Sumário')
    r.bold = True
    r.font.name = 'Arial'
    r.font.size = Pt(18)
    toc = doc.add_paragraph()
    add_field(toc, 'TOC \\o "1-3" \\h \\z \\u')

    doc.add_page_break()
    add_heading(doc, '1 Visão geral do produto', 1)
    add_heading(doc, '1.1 Objetivo', 2)
    add_body(doc, 'O AutoPlan Hub centraliza automações operacionais em uma única interface. Cada automação é isolada em seu próprio módulo, o que reduz o impacto de mudanças, facilita testes e permite acrescentar novas funções sem transformar o projeto em um único arquivo difícil de manter.')
    add_heading(doc, '1.2 Escopo da versão 0.6.1', 2)
    add_bullets(doc, [
        'Menu principal com o Ambiente de Trabalho como primeira opção e as automações de planilhas logo abaixo.',
        'Preferências persistentes para abrir WhatsApp, Outlook e Teams no Desktop ou no Chrome.',
        'Botões separados para abrir imediatamente ou modificar a configuração.',
        'Opção para respeitar a restauração das abas da sessão anterior pelo Chrome.',
        'Abertura opcional de Maxtrack GO, Zendesk, ChatGPT, Downloads e Bloco de Notas.',
        'Identidade visual própria, ícone para Windows e informações de autor e versão.',
        'Consolidador de Planilha Mãe com modos de criação e atualização.',
        'Leitura de planilhas diárias e do relatório final mensal.',
        'Consolidação de inconformidades e atrasos registrados manualmente.',
        'Atualização do resultado final e do horário de chegada das visitas atrasadas somente em Inconformidades.',
        'Motivos das fontes final e diária registrados no campo Informações.',
        'Tratamento de espaços nos nomes das abas e novas tentativas de abertura no Excel.',
        'Comparação tolerante a ajustes de horário e troca de técnico ou prestador.',
        'Prioridade do motivo sobre a mensagem de chegada e datas no padrão DD/MM/AAAA HH:MM.',
        'Geração de Painel por Tabela Dinâmica com filtro e detalhamento por duplo clique.',
        'Backup automático antes da atualização de um arquivo existente.',
        'Processamento em segundo plano para manter a interface responsiva.',
        'Gerador de Planilha Rank em janela independente, alimentado pela Planilha Mãe.',
        'Rankings The Best, Ugly e No Show, painéis por cliente e supervisão e quatro Tabelas Dinâmicas nativas.',
        'Sincronização do resultado final em Etapa Visita e Status nas inconformidades, com prioridade para o ID do contrato.',
        'Prioridade para informações explicativas registradas nas planilhas diárias.',
        'Apresentação do Rank com painel de veículos ordenado, No Show destacado e inconformidades em faixas azuis.',
    ])
    add_heading(doc, '1.3 Público e ambiente de uso', 2)
    add_body(doc, 'O software foi projetado para uso local em Windows, com Microsoft Excel instalado. O código-fonte é mantido em repositório público, mas arquivos com dados reais, credenciais, relatórios e informações internas não devem ser versionados.')

    add_heading(doc, '2 Arquitetura e organização do código', 1)
    add_body(doc, 'A solução separa interface, orquestração e regras de Excel. Essa divisão permite alterar o visual sem interferir no processamento e modificar regras de consolidação sem reescrever o menu.')
    architecture = [
        ('main.py', 'Menu principal do AutoPlan Hub e abertura dos módulos.'),
        ('app_info.py', 'Nome, autor, versão, caminhos do ícone e centralização das janelas.'),
        ('automacoes/planilha_mae/ui.py', 'Tela do Consolidador, seleção de arquivos, progresso e mensagens.'),
        ('automacoes/planilha_mae/service.py', 'Validação, criação, atualização, backup e relatório de processamento.'),
        ('automacoes/planilha_mae/engine.py', 'Leitura, normalização, comparação, escrita do Excel e Tabela Dinâmica.'),
        ('automacoes/planilha_mae/theme.py', 'Cores e estilos visuais do módulo.'),
        ('automacoes/planilha_rank/ui.py', 'Tela do Gerador de Rank, seleção da origem, período e progresso.'),
        ('automacoes/planilha_rank/service.py', 'Validação, destino e relatório da geração do Rank.'),
        ('automacoes/planilha_rank/engine.py', 'Bases, classificações, planilhas formatadas e quatro Tabelas Dinâmicas.'),
        ('automacoes/ambiente_trabalho/ui.py', 'Tela para escolher aplicativos, sites, pasta e modo Desktop ou Web.'),
        ('automacoes/ambiente_trabalho/service.py', 'Grava preferências no perfil do Windows e abre os itens selecionados.'),
        ('assets/', 'Ícones PNG e ICO do AutoPlan Hub.'),
        ('requirements.txt', 'Dependências necessárias para executar o projeto.'),
        ('.gitignore', 'Bloqueio de planilhas, backups, relatórios e arquivos temporários no Git.'),
    ]
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = 'Arquivo ou pasta'
    table.rows[0].cells[1].text = 'Responsabilidade'
    for path, responsibility in architecture:
        cells = table.add_row().cells
        cells[0].text, cells[1].text = path, responsibility
    style_table(table, [Cm(6.2), Cm(9.3)])
    add_heading(doc, '2.1 Fluxo interno', 2)
    add_numbered(doc, [
        'A interface coleta os arquivos, o período e o modo de operação.',
        'O serviço valida os caminhos e prepara o destino ou o backup.',
        'O motor lê e normaliza os arquivos de origem.',
        'Os registros são comparados e consolidados por uma chave de visita.',
        'A Planilha Mãe, a aba Inconformidades e o Painel são gerados ou atualizados.',
        'O Excel cria a Tabela Dinâmica nativa e o relatório de processamento é salvo.',
    ])

    add_heading(doc, '3 Instalação e execução', 1)
    add_heading(doc, '3.1 Requisitos', 2)
    req = doc.add_table(rows=1, cols=2)
    req.rows[0].cells[0].text = 'Componente'
    req.rows[0].cells[1].text = 'Requisito'
    for a, b in [
        ('Sistema operacional', 'Windows'),
        ('Python', '3.12 ou versão compatível'),
        ('Microsoft Excel', 'Instalado e disponível para automação COM'),
        ('Pacotes Python', 'openpyxl e pywin32'),
    ]:
        cells = req.add_row().cells; cells[0].text=a; cells[1].text=b
    style_table(req, [Cm(5), Cm(10.5)])
    add_heading(doc, '3.2 Instalação das dependências', 2)
    p = doc.add_paragraph()
    r = p.add_run('python -m pip install -r requirements.txt')
    r.font.name = 'Consolas'; r.font.size = Pt(9.5)
    add_heading(doc, '3.3 Inicialização', 2)
    add_body(doc, 'O usuário pode abrir o arquivo Abrir AutoPlan Hub.bat ou executar python main.py. O menu permanece aberto quando uma automação é iniciada. O botão Voltar ao menu fecha somente o módulo atual.')

    add_heading(doc, '3.4 Ambiente de trabalho', 2)
    add_numbered(doc, [
        'Abrir a primeira opção do menu principal.',
        'Marcar WhatsApp, Outlook e Teams e escolher Desktop ou Web no Chrome.',
        'Marcar os demais sites, a pasta Downloads e o Bloco de Notas conforme a rotina.',
        'Usar Salvar preferências para registrar a seleção ou Salvar e abrir agora para registrar e iniciar os itens.',
        'Marcar O Chrome já restaura minhas abas quando o navegador estiver configurado para recuperar a sessão anterior.',
    ])
    add_body(doc, 'No menu principal, Abrir executa diretamente a configuração salva e Configurar abre a tela de escolhas. As preferências são gravadas em uma pasta de configuração do usuário do Windows. O iniciador não encerra o Chrome e não remove abas abertas. Quando a restauração de sessão estiver marcada, o Chrome é iniciado sem receber novamente a lista de links. Quando um aplicativo Desktop não estiver instalado ou registrado no Windows, a tela informa a falha sem impedir a abertura dos demais itens.')

    add_heading(doc, '4 Consolidador de Planilha Mãe', 1)
    add_heading(doc, '4.1 Modo Criar', 2)
    add_body(doc, 'Cria uma nova Planilha Mãe na pasta de saída selecionada. Se já existir um arquivo com o mesmo nome, a interface solicita confirmação antes da substituição.')
    add_heading(doc, '4.2 Modo Atualizar', 2)
    add_body(doc, 'Atualiza diretamente a Planilha Mãe escolhida. A pasta de saída não é exibida nesse modo. Antes da primeira alteração, o software cria um backup na mesma pasta, com data e hora no nome.')
    add_heading(doc, '4.3 Arquivos de entrada', 2)
    inputs = doc.add_table(rows=1, cols=3)
    for i, text in enumerate(['Entrada', 'Obrigatória', 'Finalidade']): inputs.rows[0].cells[i].text=text
    for values in [
        ('Planilhas diárias', 'Sim', 'Fornecem visitas, estados, horários e atrasos marcados pela equipe.'),
        ('Relatório final mais recente', 'Sim', 'Fonte de comparação das inconformidades, do resultado atualizado e do horário real de chegada, inclusive quando baixada alguns dias depois.'),
        ('Planilha Mãe existente', 'Somente ao atualizar', 'Arquivo que receberá novos registros e recálculo do Painel.'),
    ]:
        cells=inputs.add_row().cells
        for i,value in enumerate(values): cells[i].text=value
    style_table(inputs, [Cm(4.2), Cm(3.2), Cm(8.1)])
    add_heading(doc, '4.4 Resultado', 2)
    add_bullets(doc, [
        'Abas diárias em ordem cronológica.',
        'Aba Inconformidades após as abas diárias.',
        'Aba Painel imediatamente depois de Inconformidades.',
        'Tabela Dinâmica com filtro, contagem e detalhamento por duplo clique.',
        'Relatório de conflitos, atualizações, avisos e backup.',
    ])

    add_heading(doc, '5 Gerador de Planilha Rank', 1)
    add_heading(doc, '5.1 Entradas e resultado', 2)
    add_body(doc, 'A Planilha Mãe concluída é a entrada obrigatória. Opcionalmente, o usuário pode selecionar uma Planilha Rank anterior para reaproveitar a relação entre prestadores e supervisores. O resultado é salvo com o nome Planilha Rank - Mês Ano - Logística e EAB - Fechamento.xlsx na pasta escolhida.')
    add_heading(doc, '5.2 Estrutura gerada', 2)
    add_bullets(doc, [
        'Painel Inconformidades e sua base detalhada.',
        'Base Rank oculta e Painel Rank oculto, usados pelos cálculos do fechamento.',
        'Rank The Best com os dez prestadores de maior volume sem ocorrências penalizadoras.',
        'Rank Ugly com prestadores que possuem ocorrências e uma relação separada de No Show.',
        'Base Veículo Indisponível oculta e painel por cliente e categoria.',
        'Painel Supervisão e Base Supervisão oculta, com prestadores agrupados por responsável.',
    ])
    add_heading(doc, '5.3 Critérios dos rankings', 2)
    add_body(doc, 'A Base Rank combina as visitas concluídas com Atrasado - Técnico informou, Atrasado - Técnico não informou, Atraso Penalizado e No Show. The Best considera somente prestadores sem essas ocorrências e ordena pelo volume concluído. Ugly ordena prestadores com ocorrências pela quantidade total. No Show permanece em uma tabela própria. Espaços excedentes e variações simples são normalizados para impedir a separação indevida do mesmo prestador.')
    add_heading(doc, '5.4 Veículos e supervisão', 2)
    add_body(doc, 'O painel de veículos reúne cancelamentos pelo cliente, ausências da equipe, divergências ou ausência de kit, falta de integridade, veículo indisponível e reagendamento cancelado pelo cliente. O painel de supervisão usa as quatro categorias penalizadoras. Quando não existe Rank anterior ou surge um prestador novo, a supervisão fica como Não mapeado para revisão humana.')
    add_heading(doc, '5.5 Painéis dinâmicos', 2)
    add_body(doc, 'Os quatro painéis são Tabelas Dinâmicas nativas do Excel. Os filtros permanecem disponíveis e o duplo clique em uma contagem abre os registros que formam aquele total. No Painel Veículo Indisponível, os clientes são ordenados do maior Total Geral para o menor. Os painéis de supervisão e de veículo indisponível exibem grades completas. No Rank Ugly, o quadro de No Show fica abaixo da tabela principal com cabeçalho vermelho. A aba Inconformidades segue o mesmo padrão visual da Planilha Mãe. As bases técnicas ficam ocultas para preservar a apresentação sem impedir auditoria ou atualização.')
    add_body(doc, 'No Painel Supervisão, o vínculo é reaproveitado pelo campo Técnico. O prestador é usado somente como alternativa quando todos os seus registros conhecidos pertencem à mesma supervisão, evitando classificar incorretamente empresas que possuem técnicos de equipes diferentes. A Tabela Dinâmica apresenta apenas a hierarquia Supervisão e Prestador; o técnico é uma chave interna de associação e não aparece como terceiro nível do painel.')
    add_body(doc, 'A comparação de técnicos normaliza acentos, maiúsculas e espaços, considera o nome antes de complementos de equipe ou empresa e permite aproximação textual apenas quando existe um único candidato seguro. Se o registro não possuir prestador, o painel utiliza o nome do técnico como rótulo, evitando o agrupamento (vazio).')
    add_body(doc, 'Durante o processamento, o log mostra a leitura do Rank anterior e cada associação no formato Técnico para Supervisão, indicando se o vínculo veio do próprio técnico ou de um prestador sem ambiguidade. Também apresenta os totais carregados, mapeados e não mapeados. Quando nenhum Rank anterior é selecionado, a tela exibe um aviso explícito. O relatório final registra a mesma auditoria para facilitar a conferência posterior.')
    add_heading(doc, '5.6 Relatórios e revisão obrigatória', 2)
    add_body(doc, 'Os relatórios são organizados em seções de identificação, arquivos utilizados, resultado resumido, detalhamento do processamento, método de mapeamento quando aplicável e checklist de revisão. Ao finalizar qualquer geração ou atualização, o AutoPlan Hub também exibe um aviso para que o usuário revise dados ausentes, conflitos, totais, associações e detalhes antes de entregar, enviar ou considerar a planilha como versão oficial.')
    add_body(doc, 'A conclusão também reproduz um arquivo WAV da pasta sounds de forma assíncrona. Se o arquivo estiver ausente ou o sistema não conseguir reproduzi-lo, o processamento e a mensagem final continuam normalmente.')

    add_heading(doc, '6 Regras de dados e inconformidades', 1)
    add_heading(doc, '6.1 Identificação de datas e colunas', 2)
    add_body(doc, 'O motor aceita nomes de arquivos com datas nos formatos DD-MM, DD-MM-AA e DD-MM-AAAA. Os cabeçalhos são normalizados para absorver diferenças como Tipo de Visita e Tipo visita, Etapa da Visita e Etapa Visita, além de caracteres acentuados danificados por exportações.')
    add_heading(doc, '6.2 Chave de visita', 2)
    add_body(doc, 'O contrato não é usado isoladamente para eliminar duplicidades. Entre as planilhas diárias, a comparação procura o mesmo contrato no mesmo dia e usa placa, tipo de visita e proximidade do horário para escolher o registro correto. Ajustes de até quatro horas são aceitos; diferenças maiores permanecem como visitas distintas para não misturar turnos. Técnico e prestador podem mudar durante o atendimento sem criar uma nova visita.')
    add_heading(doc, '6.3 Categorias oficiais', 2)
    categories = [
        ('Reagendamento', 'Lançado pelo supervisor'), ('Reagendamento', 'Cancelado pelo cliente'),
        ('Reagendamento', 'Sem tempo hábil'), ('Reagendamento', 'Veículo Indisponível'),
        ('Reagendamento', 'Técnico Cancelou'), ('Improdutiva', 'Veículo Indisponível'),
        ('Improdutiva', 'Veículo disponível com falta de integridade'), ('Improdutiva', 'Horário Excedido'),
        ('Improdutiva', 'Com necessidade de troca'), ('Improdutiva', 'Divergência ou ausência do kit de instalação'),
        ('Improdutiva', 'Ausência da equipe do cliente'), ('Improdutiva', 'Acesso não Liberado Documentação'),
        ('Improdutiva', 'Divergência do agendamento'), ('Cancelamento', 'Cancelada pela Maxtrack'),
        ('Cancelamento', 'Cancelada pelo cliente'), ('Pendência', 'Não finalizado no dia'),
        ('Divergência', 'Divergência de lançamento'), ('Comparecimento', 'No Show Não Comparecimento do técnico'),
        ('Atraso', 'Atraso Penalizado'), ('Atraso diário', 'Técnico informou'),
        ('Atraso diário', 'Técnico não informou'),
    ]
    cat = doc.add_table(rows=1, cols=2)
    cat.rows[0].cells[0].text='Grupo'; cat.rows[0].cells[1].text='Categoria'
    for group, name in categories:
        cells=cat.add_row().cells; cells[0].text=group; cells[1].text=name
    style_table(cat, [Cm(4.2), Cm(11.3)])
    add_heading(doc, '6.4 Painel e detalhamento', 2)
    add_body(doc, 'O Painel é uma Tabela Dinâmica nativa do Excel, não uma tabela estática. O filtro é disponibilizado no campo Rótulos de Linha. O duplo clique em uma contagem cria uma nova aba de detalhes com os registros que compõem aquele total.')
    add_heading(doc, '6.5 Resultado e chegada das visitas atrasadas', 2)
    add_body(doc, 'Para as categorias Atrasado - Técnico informou e Atrasado - Técnico não informou, o motor procura primeiro o ID do contrato no relatório final. Quando o contrato aparece uma única vez, a correspondência é aceita mesmo após uma remarcação de horário. Se houver mais de uma visita do mesmo contrato, data, proximidade do horário, placa e tipo de visita evitam misturar atendimentos diferentes. O resultado operacional mais avançado é gravado igualmente em Etapa Visita e Status somente na cópia de Inconformidades. Uma explicação preenchida em Informações na planilha diária tem prioridade sobre o motivo final e sobre a chegada automática. Se o conteúdo for apenas uma chegada, ele é padronizado como Chegou às 12h35. A linha original da aba diária não é modificada.')

    add_heading(doc, '7 Segurança privacidade e recuperação', 1)
    add_heading(doc, '7.1 Proteção dos dados', 2)
    add_body(doc, 'O projeto deve processar os arquivos somente no computador autorizado. Planilhas, relatórios, backups e dados operacionais não devem ser enviados ao repositório. O arquivo .gitignore bloqueia extensões e pastas de dados mais comuns, mas o responsável deve conferir cada commit antes da publicação.')
    add_heading(doc, '7.2 Backup', 2)
    add_body(doc, 'No modo Atualizar, uma cópia é criada antes da gravação. O padrão do nome é nome_original_backup_AAAAMMDD_HHMMSS.extensão. O backup deve ser mantido até a validação do arquivo atualizado.')
    add_heading(doc, '7.3 Recuperação', 2)
    add_numbered(doc, [
        'Fechar a Planilha Mãe atualizada no Excel.',
        'Localizar o backup criado na mesma pasta.',
        'Comparar a data e a hora para selecionar a versão correta.',
        'Renomear ou copiar o backup para o nome oficial somente após confirmar o conteúdo.',
    ])

    add_heading(doc, '7.4 Proteção de dados pessoais e LGPD', 2)
    add_body(doc, 'As planilhas podem conter nomes, identificadores profissionais, horários, clientes e registros de atendimento associados a pessoas naturais. A organização responsável deve definir sua posição como controladora ou operadora, a finalidade e a base legal de cada tratamento, limitar o acesso, estabelecer prazo de retenção e atender os direitos dos titulares. O programa executa processamento local e não elimina essas obrigações.')
    add_body(doc, 'Em conformidade com os princípios de finalidade, necessidade, segurança, prevenção e responsabilização previstos na Lei nº 13.709 de 2018, recomenda-se usar somente os dados necessários, manter o repositório sem planilhas reais, restringir permissões das pastas, conservar backups pelo tempo estritamente necessário e registrar incidentes e correções.')
    add_heading(doc, '7.5 Segurança da informação e incidentes', 2)
    add_body(doc, 'A Lei Geral de Proteção de Dados exige medidas técnicas e administrativas aptas a proteger dados pessoais. A Resolução CD ANPD nº 2 de 2022 e o guia de segurança da ANPD reforçam controles proporcionais ao risco. Para este projeto, isso inclui atualização do Windows e do Excel, contas individuais, bloqueio de tela, antivírus, cópias de segurança protegidas, revisão dos arquivos de saída e comunicação interna imediata em caso de acesso indevido, perda, envio incorreto ou alteração não autorizada.')
    add_heading(doc, '7.6 Propriedade intelectual e licenças', 2)
    add_body(doc, 'O código-fonte é protegido conforme a Lei nº 9.609 de 1998, sem prejuízo da legislação de direitos autorais. A publicação, distribuição ou licenciamento deve respeitar a titularidade definida pelo autor e pela organização. Bibliotecas, ícone, som e demais componentes de terceiros devem manter suas licenças, créditos e condições de uso. A publicidade do código não autoriza o compartilhamento de dados operacionais, dados pessoais ou credenciais.')
    add_heading(doc, '7.7 Limites e responsabilidade de uso', 2)
    add_body(doc, 'O AutoPlan Hub auxilia a consolidação e a análise, mas não substitui a revisão humana nem constitui decisão automatizada final. Toda planilha deve ser conferida antes da entrega. Links externos são abertos nos serviços escolhidos pelo usuário e continuam sujeitos às políticas, autenticação e disponibilidade de cada fornecedor. Esta documentação é operacional e não substitui avaliação jurídica da organização.')

    add_heading(doc, '8 Testes e controle de qualidade', 1)
    tests = [
        ('Sintaxe Python', 'Compilação dos módulos sem erros.'),
        ('Interface', 'Abertura centralizada do menu e da automação; alternância entre Criar e Atualizar.'),
        ('Processamento em segundo plano', 'Interface permanece responsiva durante a execução.'),
        ('Tabela Dinâmica', 'Criação pelo Excel, filtro ativo, detalhamento habilitado e Painel após Inconformidades.'),
        ('Dados duplicados', 'Visitas repetidas são comparadas pela chave composta.'),
        ('Atrasos', 'Resultado e horário são atualizados em Inconformidades sem modificar a aba diária.'),
        ('Privacidade', 'Arquivos reais, backups e relatórios permanecem fora do Git.'),
        ('Planilha Rank', 'Dez abas na ordem prevista, bases técnicas ocultas, quatro painéis dinâmicos e rankings consistentes.'),
    ]
    table=doc.add_table(rows=1,cols=2)
    table.rows[0].cells[0].text='Verificação'; table.rows[0].cells[1].text='Critério de aceite'
    for name, criterion in tests:
        cells=table.add_row().cells; cells[0].text=name; cells[1].text=criterion
    style_table(table,[Cm(5.5),Cm(10)])
    add_heading(doc, '8.1 Lista de validação antes do uso', 2)
    add_bullets(doc, [
        'Confirmar o mês e o ano selecionados.',
        'Confirmar que todos os arquivos diários foram incluídos.',
        'Confirmar o relatório final mensal correto.',
        'Fechar o arquivo de destino no Excel antes do processamento.',
        'Revisar o relatório gerado e os totais do Painel.',
        'Manter o backup até a aprovação do resultado.',
    ])

    add_heading(doc, '9 Versionamento Git e publicação', 1)
    add_heading(doc, '9.1 Modelo de versão', 2)
    add_body(doc, 'O projeto usa versionamento semântico no formato MAJOR.MINOR.PATCH. MAJOR indica mudança incompatível, MINOR indica nova funcionalidade compatível e PATCH indica correção compatível. A versão atual é 0.6.1, que separa as ações do ambiente de trabalho e evita links duplicados quando o Chrome restaura a sessão anterior.')
    add_heading(doc, '9.2 Convenção de commits', 2)
    commits = [
        ('feat:', 'Nova funcionalidade.'), ('fix:', 'Correção de defeito.'),
        ('docs:', 'Documentação.'), ('chore:', 'Configuração, organização ou manutenção sem mudança funcional.'),
        ('refactor:', 'Reorganização interna sem alterar o comportamento esperado.'),
        ('test:', 'Criação ou alteração de testes.'), ('style:', 'Mudança visual ou formatação sem alterar regras.'),
    ]
    table=doc.add_table(rows=1,cols=2)
    table.rows[0].cells[0].text='Prefixo'; table.rows[0].cells[1].text='Uso'
    for prefix,use in commits:
        cells=table.add_row().cells; cells[0].text=prefix; cells[1].text=use
    style_table(table,[Cm(3.4),Cm(12.1)])
    add_heading(doc, '9.3 Histórico inicial de commits', 2)
    history = [
        ('dc16b21', 'chore: configura arquivos ignorados do projeto'),
        ('18a5489', 'feat: cria AutoPlan Hub e consolidador de planilha mãe'),
        ('7a65b91', 'docs: documenta instalação e estrutura do AutoPlan Hub'),
        ('6aafc91', 'style: atualiza identidade visual do AutoPlan Hub'),
        ('9ffaa43', 'docs: adiciona documentação técnica e histórico de versões'),
        ('c887035', 'feat: atualiza atrasos com dados do relatório final'),
        ('eeb4ad7', 'docs: documenta atualização das inconformidades'),
        ('47ce450', 'feat: cria motor da planilha rank'),
        ('0b90908', 'feat: adiciona gerador de rank ao menu'),
        ('5f5eedf', 'fix: corrige filtros e cabeçalhos do rank'),
    ]
    table=doc.add_table(rows=1,cols=2)
    table.rows[0].cells[0].text='Commit'; table.rows[0].cells[1].text='Descrição'
    for commit,description in history:
        cells=table.add_row().cells; cells[0].text=commit; cells[1].text=description
    style_table(table,[Cm(3.4),Cm(12.1)])
    add_heading(doc, '9.4 Publicação no GitHub', 2)
    add_body(doc, 'O repositório oficial público está disponível em https://github.com/LeonardoVenom1k/AutoPlan-Hub. Antes de cada envio, o comando git status deve ser revisado para confirmar que não existem planilhas, relatórios, backups, credenciais, dados pessoais ou arquivos temporários preparados para commit.')

    add_heading(doc, '10 Manutenção e evolução', 1)
    add_heading(doc, '10.1 Onde alterar cada tipo de comportamento', 2)
    maintenance = [
        ('Visual, campos e mensagens', 'ui.py e theme.py'),
        ('Versão, autor, nome e ícones', 'app_info.py'),
        ('Validação, backup e destino', 'service.py'),
        ('Cabeçalhos, categorias, comparação e Excel', 'engine.py'),
        ('Novo cartão no menu', 'main.py'),
        ('Nova automação', 'Novo pacote dentro de automacoes'),
    ]
    table=doc.add_table(rows=1,cols=2)
    table.rows[0].cells[0].text='Mudança'; table.rows[0].cells[1].text='Local principal'
    for change,location in maintenance:
        cells=table.add_row().cells; cells[0].text=change; cells[1].text=location
    style_table(table,[Cm(7.5),Cm(8)])
    add_heading(doc, '10.2 Processo recomendado para uma alteração', 2)
    add_numbered(doc, [
        'Criar uma branch com nome curto e descritivo.',
        'Alterar apenas o módulo responsável pelo comportamento.',
        'Executar testes com dados fictícios antes de usar dados reais.',
        'Validar a interface, a planilha gerada e o relatório.',
        'Atualizar este documento e o README quando a mudança afetar o uso ou a arquitetura.',
        'Atualizar APP_VERSION conforme o impacto e criar um commit convencional.',
    ])
    add_heading(doc, '10.3 Próximas evoluções sugeridas', 2)
    add_bullets(doc, [
        'Empacotamento em executável Windows com ícone próprio.',
        'Instalador e mecanismo de atualização controlada.',
        'Tela de configurações para categorias e mapeamentos.',
        'Tela de manutenção do cadastro de supervisão sem depender de uma planilha anterior.',
        'Novos módulos no menu do AutoPlan Hub.',
    ])

    add_heading(doc, '11 Histórico de versões', 1)
    versions = doc.add_table(rows=1, cols=5)
    for i, value in enumerate(['Software', 'Documento', 'Data', 'Responsável', 'Resumo']):
        versions.rows[0].cells[i].text=value
    for values in [
        ['0.1.0','1.0','09/11/2026','Leonardo Aguiar','Estrutura inicial do AutoPlan Hub, módulo Consolidador, identidade visual, Git e documentação.'],
        ['0.2.0','1.1','09/11/2026','Leonardo Aguiar','Resultado final e horário de chegada dos atrasos atualizados somente em Inconformidades.'],
        ['0.2.1','1.2','09/11/2026','Leonardo Aguiar','Padronização de horários, motivos em Informações e abertura mais robusta do Excel.'],
        ['0.2.2','1.3','09/11/2026','Leonardo Aguiar','Comparação por dia e horário próximo, motivo prioritário e datas padronizadas.'],
        ['0.3.0','1.4','14/11/2026','Leonardo Aguiar','Gerador de Planilha Rank, rankings, painéis dinâmicos e reaproveitamento da supervisão.'],
        ['0.3.1','1.5','14/11/2026','Leonardo Aguiar','Correspondência por contrato e sincronização de Etapa Visita e Status nas inconformidades.'],
        ['0.3.2','1.6','14/11/2026','Leonardo Aguiar','Prioridade para explicações registradas em Informações nas planilhas diárias.'],
        ['0.3.3','1.7','14/11/2026','Leonardo Aguiar','Painel de veículos ordenado, No Show destacado e inconformidades com faixas azuis.'],
        ['0.3.4','1.8','14/11/2026','Leonardo Aguiar','No Show reposicionado, grades nos painéis e supervisão identificada pelo técnico.'],
        ['0.3.5','1.9','14/11/2026','Leonardo Aguiar','Log detalha o mapeamento e o painel organiza Supervisão e Prestador.'],
        ['0.4.0','2.0','14/11/2026','Leonardo Aguiar','Relatórios completos e aviso obrigatório de revisão antes da entrega.'],
        ['0.4.1','2.1','14/11/2026','Leonardo Aguiar','Grade preta fina aplicada em toda a base de inconformidades do Rank.'],
        ['0.4.2','2.2','14/11/2026','Leonardo Aguiar','Associação de técnicos aprimorada e prestadores vazios preenchidos pelo técnico.'],
        ['0.5.0','2.3','14/11/2026','Leonardo Aguiar','Som de notificação reproduzido ao concluir as automações.'],
        ['0.5.1','2.4','14/11/2026','Leonardo Aguiar','Entrada renomeada para relatório final mais recente usado na comparação.'],
        ['0.6.0','3.0','16/11/2026','Leonardo Aguiar','Ambiente de trabalho configurável, preferências persistentes e documentação formal com referências legais.'],
        ['0.6.1','3.1','16/11/2026','Leonardo Aguiar','Botões Abrir e Configurar separados e opção para respeitar a restauração de abas do Chrome.'],
    ]:
        row=versions.add_row().cells
        for i,value in enumerate(values): row[i].text=value
    style_table(versions,[Cm(1.8),Cm(2.4),Cm(2.4),Cm(3.6),Cm(5.3)])
    add_heading(doc, '11.1 Regra de atualização deste documento', 2)
    add_body(doc, 'Toda versão funcional que alterar o comportamento, a interface, a arquitetura, as dependências ou as regras de dados deve gerar uma nova linha no histórico. Correções apenas textuais podem incrementar somente a versão do documento.')

    doc.add_page_break()
    add_heading(doc, '12 Referências', 1)
    references = [
        'ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR 14724: informação e documentação, trabalhos acadêmicos, apresentação. Rio de Janeiro: ABNT.',
        'ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR 6023: informação e documentação, referências, elaboração. Rio de Janeiro: ABNT.',
        'ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR 6024: informação e documentação, numeração progressiva das seções de um documento. Rio de Janeiro: ABNT.',
        'ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. ABNT NBR 6027: informação e documentação, sumário, apresentação. Rio de Janeiro: ABNT.',
        'BRASIL. Lei nº 9.609, de 19 de fevereiro de 1998. Dispõe sobre a proteção da propriedade intelectual de programa de computador. Disponível em: https://www.planalto.gov.br/ccivil_03/leis/l9609.htm. Acesso em: 16 nov. 2026.',
        'BRASIL. Lei nº 9.610, de 19 de fevereiro de 1998. Altera, atualiza e consolida a legislação sobre direitos autorais. Disponível em: https://www.planalto.gov.br/ccivil_03/leis/l9610.htm. Acesso em: 16 nov. 2026.',
        'BRASIL. Lei nº 12.965, de 23 de abril de 2014. Estabelece princípios, garantias, direitos e deveres para o uso da Internet no Brasil. Disponível em: https://www.planalto.gov.br/ccivil_03/_ato2011-2014/2014/lei/l12965.htm. Acesso em: 16 nov. 2026.',
        'BRASIL. Lei nº 13.709, de 14 de agosto de 2018. Lei Geral de Proteção de Dados Pessoais. Disponível em: https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/lei/l13709.htm. Acesso em: 16 nov. 2026.',
        'AUTORIDADE NACIONAL DE PROTEÇÃO DE DADOS. Guia orientativo sobre segurança da informação para agentes de tratamento de pequeno porte. Brasília: ANPD, 2021. Disponível em: https://www.gov.br/anpd/pt-br/centrais-de-conteudo/materiais-educativos-e-publicacoes/processo-guia-orientativo-sobre-seguranca-da-informacao-para-agentes-de-tratamento-de-pequeno-porte.pdf. Acesso em: 16 nov. 2026.',
        'AUTORIDADE NACIONAL DE PROTEÇÃO DE DADOS. Resolução CD ANPD nº 2, de 27 de janeiro de 2022. Disponível em: https://www.gov.br/anpd/pt-br/acesso-a-informacao/institucional/atos-normativos/regulamentacoes_anpd/resolucao-cd-anpd-no-2-de-27-de-janeiro-de-2022. Acesso em: 16 nov. 2026.',
    ]
    for reference in references:
        p = doc.add_paragraph(reference)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(12)

    doc.add_page_break()
    annex_heading = add_heading(doc, 'Anexo A Checklist de liberação', 1)
    annex_heading.paragraph_format.keep_with_next = False
    annex_heading.paragraph_format.space_after = Pt(14)
    checklist = [
        'Versão atualizada em app_info.py.', 'README e documentação revisados.',
        'Módulos compilados sem erro.', 'Teste com dados fictícios concluído.',
        'Tabela Dinâmica e detalhamento validados no Excel.', 'Backup e restauração verificados.',
        'git status sem dados operacionais.', 'Commit convencional criado.',
        'Repositório público revisado e sem dados operacionais, pessoais ou credenciais.',
    ]
    for item in checklist:
        p=doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_together = True
        p.add_run('□  ').font.size=Pt(12)
        p.add_run(item)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == '__main__':
    build()
