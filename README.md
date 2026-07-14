# Gerenciador-de-sentimentos


Aplicativo desktop feito em **Python + Flet** para registrar e acompanhar suas emoções ao longo do tempo.

## Funcionalidades

- **Registrar**: escolha uma emoção (com emoji), defina a intensidade (1 a 10) e escreva uma descrição livre. Data e horário são gravados automaticamente.
- **Histórico**: lista de todos os registros, com opção de excluir cada um.
- **Gráfico**: gráfico de barras mostrando a frequência de cada emoção, além de cartões com estatísticas rápidas.
- **Exportar**: gere um arquivo **Word (.docx)** ou **Excel (.xlsx)** com todos os registros formatados em tabela.

Os dados ficam salvos localmente em `emocoes.db` (SQLite), na mesma pasta do app.

## Como instalar e rodar

```bash
# 1. Crie um ambiente virtual (opcional, mas recomendado)
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Rode o aplicativo
python main.py
```

Uma janela desktop nativa será aberta automaticamente.

## Rodar no navegador (opcional)

Se preferir abrir como app web local, troque a última linha do `main.py`:

```python
ft.app(target=main, view=ft.AppView.WEB_BROWSER)
```

## Estrutura de arquivos

```
emotion_app/
├── main.py            # código do aplicativo
├── requirements.txt   # dependências
├── emocoes.db          # banco de dados (criado automaticamente no 1º uso)
└── README.md
```

## Personalização rápida

- **Cores**: edite as constantes `BG`, `SURFACE`, `PRIMARY`, `PRIMARY_2` no topo do `main.py`.
- **Lista de emoções**: edite a lista `EMOTIONS` (nome, emoji, cor em hexadecimal).
- **Tamanho da janela**: `page.window.width` / `page.window.height`.
