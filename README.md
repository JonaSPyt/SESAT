<div style="text-align: center;">
  <img style="margin: 10px;" width="420" height="461" src="https://github.com/user-attachments/assets/a6de189e-a45a-4ce8-99ec-4326e7b67633" />
  <img style="margin: 10px;" width="1363" height="687" src="https://github.com/user-attachments/assets/bf4c60f7-aaeb-4121-a0e6-4b9d584a2b23" />
</div>

# SESAT

Sistema desktop em Python (Tkinter) para controle de entrada e saida de equipamentos de TI.

## Funcionalidades
- Login de usuarios com perfis (Supervisor e Consultor).

- Cadastro, edicao e exclusao de registros de equipamentos.
- Consulta automatica de patrimonio via intranet TRE-CE.
- Importacao de planilhas `.xlsx` com prevencao de duplicados.
- Exportacao para `.xlsx` no formato operacional.
- Registro de logs de auditoria (login, cadastro, alteracoes, exclusoes, importacao).

## Tecnologias
- Python 3
- Tkinter
- SQLite
- `requests`
- `beautifulsoup4`
- `openpyxl`

## Estrutura principal
- `app.py`: interface grafica e fluxo principal do sistema.
- `database.py`: persistencia SQLite, autenticacao e logs.
- `consulta_api.py`: consulta de patrimonio na intranet.
- `importar.py`: importacao de planilha `.xlsx`.
- `exportar.py`: exportacao de planilha `.xlsx`.
- `config.example.ini`: modelo de configuracao local.

## Requisitos
1. Python 3 instalado.
2. Dependencias Python:

```bash
pip install requests beautifulsoup4 openpyxl
```

## Configuracao
1. Copie `config.example.ini` para `config.ini`.
2. Edite `config.ini`:

```ini
[database]
caminho_rede =
```

- Se `caminho_rede` ficar vazio, o banco sera criado localmente (`sesat.db`).
- Se preencher com caminho de rede, o banco sera compartilhado nesse diretorio.

## Como executar

```bash
python app.py
```

## Build (opcional, executavel)
Exemplo com PyInstaller:

```bash
pyinstaller SESAT.spec
```

## Seguranca antes de publicar no GitHub
Este repositorio foi preparado para nao subir arquivos sensiveis por padrao (`.gitignore`).

Checklist recomendado:
- Nao subir `config.ini` real com caminho interno da rede.
- Nao subir `sesat.db` com dados reais.
- Nao subir pastas `build/` e `dist/`.
- Trocar as senhas padrao apos o primeiro login no ambiente real.

## Publicacao no GitHub (passo a passo)
No diretorio do projeto:

```bash
git init
git add .
git commit -m "chore: prepara projeto SESAT para github"
git branch -M main
git remote add origin <URL_DO_SEU_REPOSITORIO>
git push -u origin main
```

## Observacoes
- A consulta da intranet depende de acesso a rede interna do TRE-CE.
- Em `consulta_api.py`, a requisicao usa `verify=False` por causa de certificado interno.
