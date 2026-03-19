# Dashboard Contábil - Gerador V3

Este projeto contém o script principal (`gerar_dashboard_v3.py`) responsável por carregar os dados contábeis em formato Excel, realizar o tratamento de agrupamentos (destinos, plataformas, grupos de pessoas), aplicar relatórios de custos de passagens e hospedagens (e extrair dados de transporte), e gerar um dashboard de métricas consolidado.

## Estrutura do Projeto

* `gerar_dashboard_v3.py`: Script principal do gerador de dashboard.
* `.gitignore`: Configuração para impedir envio de arquivos de dados, html e configs pessoais.
* `requirements.txt`: Dependências do projeto.
* `config.json`: (opcional) Configurações variadas do gerador.

## Pré-requisitos

1. Python 3.8+
2. Instalar dependências:
```bash
pip install -r requirements.txt
```

## Como Usar

Para rodar a ferramenta e gerar o relatório do Dashboard:

```bash
python gerar_dashboard_v3.py
```

Isto irá ler as planilhas de entrada (dados) e gerar um arquivo interativo consolidado em HTML como dashboard de saída que pode ser aberto no navegador de sua preferência.

> Os arquivos `.xlsx` com as bases completas não são incluídos no repositório de forma nativa por questões de confidencialidade de dados, por isso eles estão presentes no `.gitignore`.
