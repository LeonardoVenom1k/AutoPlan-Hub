# AutoPlan Hub

![Ícone do AutoPlan Hub](assets/autoplan-hub-icon.png)

Central local de automações para rotinas operacionais.

**Versão:** 0.2.0  
**Desenvolvido por:** Leonardo Aguiar

## Automação disponível

### Consolidador de Planilha Mãe

- Consolida planilhas diárias.
- Compara atualizações da manhã e da tarde.
- Importa o relatório final de inconformidades.
- Atualiza a aba de inconformidades.
- Nas inconformidades de atraso, substitui o estado Agendada pelo resultado do relatório final.
- Registra o horário real no formato `Chegou às 12h35`, sem alterar as abas diárias.
- Cria um Painel com Tabela Dinâmica nativa do Excel.
- Permite abrir os detalhes das contagens com duplo clique.
- Cria backup antes de atualizar uma Planilha Mãe existente.

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
└── planilha_mae/
    ├── engine.py   # regras de leitura, comparação e geração do Excel
    ├── service.py  # criação, atualização, backup e relatório
    ├── theme.py    # cores e estilos da interface
    └── ui.py       # tela do consolidador
```

## Privacidade

Planilhas, relatórios, backups e arquivos com dados operacionais são ignorados pelo Git. Não inclua dados reais no repositório.

