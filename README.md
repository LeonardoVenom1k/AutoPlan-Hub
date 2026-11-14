# AutoPlan Hub

![Ícone do AutoPlan Hub](assets/autoplan-hub-icon.png)

Central local de automações para rotinas operacionais.

**Versão:** 0.3.2
**Desenvolvido por:** Leonardo Aguiar

## Automações disponíveis

### Consolidador de Planilha Mãe

- Consolida planilhas diárias.
- Compara atualizações da manhã e da tarde.
- Importa o relatório final de inconformidades.
- Atualiza a aba de inconformidades.
- Nas inconformidades de atraso, substitui o estado Agendada pelo resultado do relatório final.
- Registra o horário real no formato `Chegou às 12h35`, sem alterar as abas diárias.
- Leva os motivos encontrados nas planilhas final e diárias para `Informações`.
- Quando existe motivo, mantém somente o motivo em `Informações`, sem o texto de chegada.
- Aceita ajuste de horário de até 4 horas e troca de técnico/prestador ao localizar a mesma visita, sem misturar atendimentos de turnos diferentes.
- Formata datas e horários como `DD/MM/AAAA HH:MM`.
- Aceita pequenas variações nos nomes das abas e tenta reabrir o arquivo antes de falhar.
- Cria um Painel com Tabela Dinâmica nativa do Excel.
- Permite abrir os detalhes das contagens com duplo clique.
- Cria backup antes de atualizar uma Planilha Mãe existente.
- Nas inconformidades, cruza o relatório final pelo ID do contrato mesmo quando o horário foi remarcado.
- Mantém `Etapa Visita` e `Status` com o mesmo resultado final, sem alterar as abas diárias.
- Preserva e prioriza explicações preenchidas em `Informações` nas planilhas diárias; mensagens que contêm somente a chegada continuam padronizadas.

### Gerador de Planilha Rank

- Usa uma Planilha Mãe pronta como fonte da análise mensal.
- Reproduz a estrutura da planilha de fechamento com dez abas e bases técnicas ocultas.
- Gera os rankings `The Best`, `Ugly` e a relação separada de `No Show`.
- Consolida os painéis de inconformidades, veículos indisponíveis e supervisão.
- Cria quatro Tabelas Dinâmicas nativas do Excel, com filtros e detalhamento por duplo clique.
- Permite selecionar uma Planilha Rank anterior para reaproveitar o cadastro de prestadores por supervisor.
- Indica prestadores novos como `Não mapeado`, evitando atribuições automáticas incorretas.
- Normaliza espaços e pequenas variações nos nomes para não dividir o mesmo prestador em duas linhas.

## Requisitos

- Windows.
- Python 3.12 ou versão compatível.
- Microsoft Excel instalado.

## Instalação

```powershell
python -m pip install -r requirements.txt
```

## Execução

```powershell
python main.py
```

Também é possível abrir pelo arquivo `Abrir AutoPlan Hub.bat`.

## Organização

```text
automacoes/
├── planilha_mae/
│   ├── engine.py   # regras de leitura, comparação e geração do Excel
│   ├── service.py  # criação, atualização, backup e relatório
│   ├── theme.py    # cores e estilos da interface
│   └── ui.py       # tela do consolidador
└── planilha_rank/
    ├── engine.py   # bases, rankings, formatação e Tabelas Dinâmicas
    ├── service.py  # validação, destino e relatório de processamento
    └── ui.py       # tela independente do gerador de Rank
```

## Privacidade

Planilhas, relatórios, backups e arquivos com dados operacionais são ignorados pelo Git. Não inclua dados reais no repositório.

