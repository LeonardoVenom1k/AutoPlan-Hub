# AutoPlan Hub

![Ícone do AutoPlan Hub](assets/autoplan-hub-icon.png)

Central local de automações para rotinas operacionais.

**Versão:** 0.5.1
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
- Ordena o Painel Veículo Indisponível pelo `Total Geral`, do maior para o menor.
- Exibe o quadro de `No Show` abaixo da tabela principal do `Rank Ugly`, com cabeçalho vermelho.
- Mantém a aba `Inconformidades` com o mesmo padrão visual usado na Planilha Mãe.
- Aplica grades visíveis nos painéis de supervisão e de veículo indisponível.
- Identifica a supervisão pelo campo `Técnico`, usando o prestador apenas quando seu vínculo não é ambíguo.
- Exibe no andamento e no relatório cada associação `Técnico → Supervisão`, sua fonte, os vínculos carregados e quantas ocorrências foram ou não mapeadas.
- Organiza o Painel Supervisão na hierarquia `Supervisão → Prestador`; o técnico é usado apenas para localizar o supervisor correto.
- Gera relatórios divididos em identificação, arquivos utilizados, resultados, detalhamento do processamento e checklist de revisão.
- Ao concluir qualquer geração ou atualização, avisa que a planilha deve ser revisada antes de ser entregue ou usada como versão oficial.
- Aplica uma grade preta fina em toda a aba `Inconformidades` da Planilha Rank, preservando as faixas azuis.
- Compara técnicos por nome completo, nome-base e aproximação segura quando houver um único candidato confiável.
- Quando o prestador estiver vazio no Painel Supervisão, usa o nome do técnico no lugar de `(vazio)`.
- Reproduz um som de notificação ao concluir as automações de Planilha Mãe e Planilha Rank.
- Identifica a entrada mensal como o relatório final mais recente usado para comparar resultados e horários, permitindo atualizações feitas aos poucos.

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

