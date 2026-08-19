# Automação de Inventário de Paleteiras

Automação desenvolvida em Python para processamento das bases de inventário de paleteiras, identificação de equipamentos pendentes e envio automático de cobranças por e-mail utilizando a **Gmail API**.

A aplicação cruza duas bases Excel:

- Base de leituras/inventários;
- Base de paleteiras em operação.

A partir desse cruzamento, a aplicação identifica quais equipamentos estão dentro ou fora do prazo de inventário e envia automaticamente um e-mail para os responsáveis de cada filial.

---

## 📋 Funcionalidades

A aplicação realiza as seguintes etapas:

1. Valida a existência das bases Excel.
2. Carrega a Base 1.
3. Carrega a Base 2.
4. Normaliza os QR Codes.
5. Cruza as informações das duas bases.
6. Identifica a última leitura de cada paleteira.
7. Calcula o prazo de inventário.
8. Classifica cada equipamento como:
   - `LIDO`
   - `PENDENTE`
   - `NUNCA LIDO`
9. Calcula o tempo desde a última leitura.
10. Calcula o valor financeiro das pendências.
11. Gera um relatório completo em Excel.
12. Gera um resumo por filial.
13. Gera o corpo HTML do e-mail.
14. Envia automaticamente os e-mails através da Gmail API.
15. Permite múltiplos destinatários por filial.
16. Permite configurar caminhos, destinatários e parâmetros através de `.env`.

---

# 🛠️ Tecnologias utilizadas

- Python
- Pandas
- OpenPyXL
- Gmail API
- Google OAuth 2.0
- python-dotenv
- Microsoft SharePoint/OneDrive sincronizado localmente
- Excel

---

# 📌 Requisitos

## Python

É necessário possuir Python instalado.

A documentação atual do Google para a Gmail API recomenda Python 3.10.7 ou superior para o quickstart.

Recomenda-se utilizar Python 3.10 ou superior.

Verifique a instalação:

```bash
python --version
```

ou:

```bash
py --version
```

Exemplo:

```text
Python 3.14.6
```

---

# 📁 Estrutura do projeto

A estrutura recomendada do projeto é:

```text
automacao-paleteiras/
│
├── main.py
│
├── .env
├── .env.example
├── .gitignore
├── requirements.txt
│
├── credentials.json
├── token.json
│
└── saida/
```

### Arquivos importantes

| Arquivo | Função |
|---|---|
| `main.py` | Código principal da aplicação |
| `.env` | Configurações locais e dados sensíveis |
| `.env.example` | Modelo do `.env` |
| `.gitignore` | Impede envio de arquivos sensíveis ao GitHub |
| `requirements.txt` | Dependências Python |
| `credentials.json` | Credenciais OAuth do Google |
| `token.json` | Token de autorização do Gmail |
| `saida/` | Relatórios gerados pela aplicação |

---

# 🚀 Instalação

## 1. Clonar o repositório

Depois que o projeto estiver publicado no GitHub:

```bash
git clone URL_DO_REPOSITORIO
```

Entre na pasta:

```bash
cd automacao-paleteiras
```

---

# 2. Criar ambiente virtual

No Windows:

```bash
python -m venv .venv
```

Ative o ambiente:

```bash
.venv\Scripts\activate
```

Se estiver utilizando PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Quando o ambiente estiver ativo, o terminal deverá apresentar algo semelhante a:

```text
(.venv) C:\...\automacao-paleteiras>
```

---

# 3. Atualizar o pip

```bash
python -m pip install --upgrade pip
```

---

# 4. Instalar as dependências

Se o projeto possuir `requirements.txt`:

```bash
pip install -r requirements.txt
```

As principais bibliotecas utilizadas são:

```text
pandas
openpyxl
python-dotenv
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
```

A instalação das bibliotecas oficiais do cliente Google pode ser feita via pip.

---

# 5. Instalação manual das dependências

Caso o `requirements.txt` não esteja disponível, execute:

```bash
pip install pandas
pip install openpyxl
pip install python-dotenv
pip install google-api-python-client
pip install google-auth-httplib2
pip install google-auth-oauthlib
```

Ou em um único comando:

```bash
pip install pandas openpyxl python-dotenv google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

O `python-dotenv` é utilizado para carregar as configurações do arquivo `.env` para as variáveis de ambiente da aplicação.

---

# 🔐 Configuração da Gmail API

A aplicação utiliza OAuth 2.0 para autorização da conta Gmail.

## 1. Criar um projeto no Google Cloud

Acesse o Google Cloud Console:

https://console.cloud.google.com/

Crie um novo projeto.

Exemplo de nome:

```text
Automação Inventário Paleteiras
```

---

# 2. Ativar a Gmail API

Dentro do projeto:

```text
APIs e serviços
    ↓
Biblioteca
    ↓
Gmail API
    ↓
Ativar
```

A Gmail API precisa estar habilitada no projeto do Google Cloud antes da aplicação realizar chamadas à API.

---

# 3. Configurar OAuth

No Google Cloud:

```text
Google Auth Platform
```

Configure:

```text
Branding
Audience
Data Access
```

Para uma aplicação interna de uma organização Google Workspace, o Google permite configurar o público como `Internal`, quando essa opção estiver disponível para o ambiente da organização.

---

# 4. Criar credencial OAuth

Acesse:

```text
Google Auth Platform
    ↓
Clients
    ↓
Create Client
```

Selecione:

```text
Application type:
Desktop app
```

Crie a credencial.

A documentação oficial do Google orienta criar um OAuth Client ID para aplicativo desktop e salvar o JSON baixado como `credentials.json`.

---

# 5. Baixar credentials.json

Baixe o arquivo JSON gerado pelo Google.

Renomeie para:

```text
credentials.json
```

Coloque na raiz do projeto:

```text
automacao-paleteiras/
│
├── main.py
├── credentials.json
├── .env
└── ...
```

---

# ⚠️ IMPORTANTE — credentials.json NÃO deve ir para o GitHub

O arquivo:

```text
credentials.json
```

deve estar no `.gitignore`.

Nunca faça commit desse arquivo em um repositório público.

---

# 🔑 Primeiro acesso ao Gmail

Na primeira execução da aplicação, ela verificará se existe:

```text
token.json
```

Caso não exista, será aberto o fluxo de autorização do Google.

Será necessário:

1. Fazer login na conta Gmail que enviará os e-mails.
2. Autorizar o acesso solicitado.
3. Concluir o processo de autorização.

Depois disso, a aplicação criará:

```text
token.json
```

O Google documenta que o token é armazenado localmente após a primeira autorização para evitar novo login em execuções posteriores.

---

# ⚠️ IMPORTANTE — token.json também não deve ir para o GitHub

O arquivo:

```text
token.json
```

também deve estar no `.gitignore`.

Não compartilhe esse arquivo.

---

# ⚙️ Configuração do .env

Crie na raiz do projeto:

```text
.env
```

Exemplo:

```env
# ============================================================
# GMAIL
# ============================================================

GMAIL_CREDENTIALS_FILE=credentials.json
GMAIL_TOKEN_FILE=token.json


# ============================================================
# SHAREPOINT / BASES
# ============================================================

SHAREPOINT_DIR=C:\CAMINHO\DO\SHAREPOINT

ARQUIVO_BASE_1=Gestão Manutenção de Frotas - Controle de Paleteiras.xlsx
ARQUIVO_BASE_2=Planilha de Paleteiras Disktrans.xlsx

ABA_BASE_2=Base


# ============================================================
# INVENTÁRIO
# ============================================================

PRAZO_DIAS=7


# ============================================================
# DESTINATÁRIOS
# ============================================================

CON_EMAIL_1=
CON_EMAIL_2=

VIX_EMAIL_1=
VIX_EMAIL_2=

SSA_EMAIL_1=
SSA_EMAIL_2=

MOC_EMAIL_1=
MOC_EMAIL_2=

AJU_EMAIL_1=
AJU_EMAIL_2=

REC_EMAIL_1=
REC_EMAIL_2=

FOR_EMAIL_1=
FOR_EMAIL_2=
```

---

# 📧 Configuração dos destinatários

Cada filial pode possuir até dois destinatários na configuração atual.

Exemplo:

```env
CON_EMAIL_1=responsavel1@empresa.com.br
CON_EMAIL_2=responsavel2@empresa.com.br
```

A aplicação transforma automaticamente isso em:

```python
"CON": [
    "responsavel1@empresa.com.br",
    "responsavel2@empresa.com.br",
]
```

---

# 🏢 Adicionando novas filiais

As filiais utilizadas pela aplicação são configuradas no código na função `carregar_destinatarios()`.

Exemplo:

```python
filiais = [
    "CON",
    "VIX",
    "SSA",
    "MOC",
    "AJU",
    "REC",
    "FOR",
]
```

Para uma nova filial, adicione o código:

```python
"NOV",
```

Depois configure no `.env`:

```env
NOV_EMAIL_1=responsavel1@patrus.com.br
NOV_EMAIL_2=responsavel2@patrus.com.br
```

---

# 📂 Configuração das bases

A aplicação utiliza duas planilhas.

## Base 1

Arquivo:

```text
Gestão Manutenção de Frotas - Controle de Paleteiras.xlsx
```

Essa base contém os registros de leitura/inventário.

São utilizadas as colunas:

```text
CODIGO
FILIAL
NOME
DATAHORA
```

---

## Base 2

Arquivo:

```text
Planilha de Paleteiras Disktrans.xlsx
```

A aplicação utiliza a aba:

```text
Base
```

As colunas necessárias são:

```text
FILIAL
QR Code
NR_DISK
MODELO
CUSTO modelo
```

---

# 📍 Caminho do SharePoint

Como as bases estão sincronizadas localmente pelo SharePoint/OneDrive, informe no `.env` o diretório onde os arquivos estão disponíveis.

Exemplo:

```env
SHAREPOINT_DIR=C:\Users\usuario\Empresa\Automacao\1.6 Alimentação BI
```

O código adicionará automaticamente os nomes dos arquivos configurados no `.env`.

---

# ⏱️ Prazo de inventário

O prazo padrão é:

```env
PRAZO_DIAS=7
```

Isso significa que uma paleteira permanece regular por 7 dias após sua última leitura.

Exemplo:

```text
Última leitura:
10/08/2026

Prazo:
7 dias

Vencimento:
17/08/2026
```

Após o vencimento, o equipamento passa a ser:

```text
PENDENTE
```

---

# 🔄 Regras de status

A aplicação utiliza três status.

## LIDO

A paleteira possui uma leitura válida e ainda está dentro do prazo.

```text
Última leitura + prazo > data atual
```

---

## PENDENTE

A paleteira possui uma leitura, porém o prazo já venceu.

```text
Última leitura + prazo <= data atual
```

---

## NUNCA LIDO

A paleteira existe na Base 2, mas não possui uma leitura correspondente na Base 1.

---

# 📊 Processamento

Ao executar:

```bash
python main.py
```

a aplicação realiza:

```text
Validar bases
      ↓
Carregar Base 1
      ↓
Carregar Base 2
      ↓
Normalizar QR Codes
      ↓
Cruzar as bases
      ↓
Calcular vencimentos
      ↓
Determinar status
      ↓
Gerar resumo
      ↓
Salvar Excel
      ↓
Identificar pendências
      ↓
Gerar HTML
      ↓
Enviar e-mails
```

---

# 📧 Envio de e-mails

A aplicação envia somente para filiais configuradas no `.env`.

Para cada filial:

1. Verifica as pendências.
2. Se não houver pendências, não envia e-mail.
3. Se houver pendências, gera o HTML.
4. Cria o assunto.
5. Envia para os destinatários configurados.

Exemplo de assunto:

```text
[PENDÊNCIAS] Inventário de Paleteiras - CON
```

---

# 🚫 Filial sem pendências

Se uma filial não possuir equipamentos:

```text
PENDENTE
```

ou:

```text
NUNCA LIDO
```

o sistema não envia e-mail para aquela filial.

Será exibido no terminal:

```text
Nenhuma pendência encontrada para CON. E-mail não será enviado.
```

---

# 📁 Arquivos gerados

Após a execução, a pasta `saida/` será criada.

```text
saida/
│
├── Resultado_Inventario.xlsx
└── Resumo_Filiais.xlsx
```

---

# 📊 Resultado_Inventario.xlsx

Contém o detalhamento dos equipamentos processados.

Principais informações:

```text
FILIAL
QR Code
NR_DISK
MODELO
CUSTO
CODIGO_LEITURA
NOME
DATAHORA
VENCIMENTO
STATUS
DIAS_SEM_LEITURA
HORAS_RESTANTES
DATA_ANALISE
```

---

# 📈 Resumo_Filiais.xlsx

Contém os indicadores consolidados por filial.

Inclui:

```text
TOTAL
LIDOS
PENDENTES
NUNCA_LIDOS
VALOR_PENDENTE
NAO_LIDOS
PERCENTUAL_LIDOS
PERCENTUAL_NAO_LIDOS
```

---

# 🧪 Testando a aplicação

Antes de automatizar a execução diária, execute manualmente:

```bash
python main.py
```

O terminal deverá apresentar algo semelhante a:

```text
============================================================
AUTOMAÇÃO DE INVENTÁRIO DE PALETEIRAS
============================================================

Validando bases...
Base 1: 19/08/2026 08:30:00
Base 2: 19/08/2026 08:31:00

Processando dados...

Processando e-mail da filial CON...
E-mail enviado para: usuario1@patrus.com.br, usuario2@patrus.com.br
ID da mensagem: xxxxxxxxxxxxxxxxx

Processando e-mail da filial VIX...
E-mail enviado para: usuario1@patrus.com.br, usuario2@patrus.com.br

Processamento concluído.
Total: 000
Lidos: 000
Pendentes: 000
Nunca lidos: 000
```

---

# 🔐 Segurança

Nunca envie para o GitHub:

```text
.env
credentials.json
token.json
```

Também não devem ser enviados:

```text
saida/
```

O `.gitignore` deve conter:

```gitignore
.env
credentials.json
token.json
saida/
```

---

# 📝 .env.example

O projeto deve possuir um arquivo:

```text
.env.example
```

Esse arquivo serve apenas como modelo.

Exemplo:

```env
GMAIL_CREDENTIALS_FILE=credentials.json
GMAIL_TOKEN_FILE=token.json

SHAREPOINT_DIR=C:\CAMINHO\DO\SHAREPOINT

ARQUIVO_BASE_1=Gestão Manutenção de Frotas - Controle de Paleteiras.xlsx
ARQUIVO_BASE_2=Planilha de Paleteiras Disktrans.xlsx

ABA_BASE_2=Base

PRAZO_DIAS=7

CON_EMAIL_1=
CON_EMAIL_2=

VIX_EMAIL_1=
VIX_EMAIL_2=

SSA_EMAIL_1=
SSA_EMAIL_2=
```

O usuário deve copiar:

```bash
.env.example
```

para:

```bash
.env
```

e preencher os dados reais.

---

# 🐙 GitHub

Antes de enviar o projeto:

```bash
git status
```

Verifique se estes arquivos **não aparecem para commit**:

```text
.env
credentials.json
token.json
saida/
```

Depois:

```bash
git add .
```

Confira:

```bash
git status
```

Se estiver tudo correto:

```bash
git commit -m "Estrutura inicial da automação de inventário"
```

Depois:

```bash
git push
```

---

# 🔍 Verificação antes do primeiro push

Execute:

```bash
git status
```

O repositório deve conter arquivos semelhantes a:

```text
main.py
.env.example
.gitignore
requirements.txt
README.md
```

E não deve conter:

```text
.env
credentials.json
token.json
saida/
```

---

# 🆘 Problemas comuns

## Python não encontrado

Erro:

```text
'python' não é reconhecido como um comando
```

Verifique:

```bash
python --version
```

ou:

```bash
py --version
```

Se necessário, instale o Python.

---

## Biblioteca não encontrada

Exemplo:

```text
ModuleNotFoundError: No module named 'pandas'
```

Execute:

```bash
pip install -r requirements.txt
```

---

## Gmail pede autorização novamente

Verifique se:

```text
credentials.json
```

está presente.

Se o token estiver inválido, pode ser necessário excluir:

```text
token.json
```

e executar novamente:

```bash
python main.py
```

Isso fará o fluxo OAuth novamente.

---

## Erro `Invalid To header`

Verifique os e-mails no `.env`.

Exemplo correto:

```env
CON_EMAIL_1=usuario1@patrus.com.br
CON_EMAIL_2=usuario2@patrus.com.br
```

Evite espaços ou caracteres extras.

---

## Base não encontrada

Erro:

```text
Base 1 não encontrado
```

Verifique:

```env
SHAREPOINT_DIR=
```

e confirme se os nomes configurados correspondem exatamente aos arquivos.

---

## Coluna não encontrada

Se aparecer:

```text
Colunas não encontradas na Base 1
```

ou:

```text
Colunas não encontradas na Base 2
```

verifique se os nomes das colunas nas planilhas foram alterados.

---

# 🧹 Reset completo da autorização Gmail

Se for necessário refazer completamente a autenticação:

1. Feche a aplicação.
2. Exclua:

```text
token.json
```

3. Execute:

```bash
python main.py
```

4. Faça novamente a autorização no Google.

Não é necessário excluir `credentials.json`.

---

# 🏗️ Desenvolvimento futuro

A aplicação está estruturada para receber futuras melhorias.

Possíveis evoluções:

- Execução automática diária.
- Agendamento pelo Windows Task Scheduler.
- Registro de logs.
- Controle de histórico dos e-mails enviados.
- Prevenção de cobranças duplicadas.
- Inclusão automática de novas filiais.
- Dashboard de acompanhamento.
- Monitoramento de falhas.
- Notificação de erro da automação.
- Separação das configurações em módulos.
- Execução em servidor.
- Integração futura com outras APIs.

---

# 🔁 Fluxo completo da solução

```text
                    SHAREPOINT
                        │
                        ▼
              ┌──────────────────┐
              │    Base 1 Excel  │
              │    Leituras      │
              └────────┬─────────┘
                       │
                       │
              ┌────────▼─────────┐
              │    Base 2 Excel  │
              │    Paleteiras    │
              └────────┬─────────┘
                       │
                       ▼
              ┌──────────────────┐
              │     Python       │
              │                  │
              │ Cruzamento       │
              │ QR Codes         │
              │ Status           │
              │ Prazos           │
              │ Valores          │
              └────────┬─────────┘
                       │
             ┌─────────┴──────────┐
             │                    │
             ▼                    ▼
      Resultado Excel       Resumo Excel
             │
             ▼
       Pendências por
           filial
             │
             ▼
       HTML personalizado
             │
             ▼
         Gmail API
             │
      ┌──────┴──────┐
      ▼             ▼
     CON           VIX
      │             │
      ▼             ▼
  Destinatários  Destinatários
```

---

# 👨‍💻 Autor

Projeto desenvolvido para automação do processo de inventário e acompanhamento de paleteiras.

**Automação - Patrus Transportes**

---

# 📄 Licença

Este projeto é de uso interno.

Não distribuir credenciais, tokens, dados das bases ou informações internas da empresa.