# Checklists de equipamentos e inventário de paleteiras

Automações Python que cruzam cadastros com históricos de registros, geram relatórios Excel e PDF e enviam cobranças por grupo via Gmail. O e-mail mantém as tabelas de pendências e inclui o resumo visual em PDF.

| Automação | Identificação do equipamento | O que analisa |
| --- | --- | --- |
| `checklist.py` | Placa | Último checklist, prazo, tipo realizado e código da filial |
| `paleteiras.py` | QR Code | Leitura de inventário, prazo e custo dos equipamentos pendentes |

Cada script funciona de forma independente e precisa apenas das suas duas bases. Ambos usam o mesmo `.env`, a mesma configuração do Gmail e os módulos de relatório em `automacao/`.

## Estrutura

```text
checklist.py                 # Execução da análise de checklists
paleteiras.py                # Execução do inventário de paleteiras
automacao/
    mensagem_email.py        # Montagem do e-mail e do anexo
    relatorios_pdf.py        # Dados, gráficos e diagramação dos PDFs
logo/logo.png               # Logo utilizada nos cabeçalhos
tests/                      # Testes com dados fictícios e envio simulado
.env.example                # Modelo público de configuração
requirements.txt            # Dependências de execução
requirements-dev.txt        # Dependências adicionais de desenvolvimento
```

Os arquivos `.env`, `credentials.json`, `token.json`, as bases em `dados/` e os relatórios em `saida/` são locais e ignorados pelo Git. As bases também podem ficar em uma pasta sincronizada externa ao projeto.

## Preparação

Use Python 3.11 ou superior. No PowerShell, na raiz do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Na primeira configuração, crie o `.env` sem sobrescrever uma configuração existente:

```powershell
if (-not (Test-Path -LiteralPath .env)) {
    Copy-Item -LiteralPath .env.example -Destination .env
}
```

Abra o `.env` em um editor e preencha:

- `SHAREPOINT_DIR`: caminho da pasta com as bases Excel. Prefira um caminho absoluto; no Windows, use `/` entre as pastas.
- `ARQUIVO_BASE_1`, `ARQUIVO_BASE_2` e `ABA_BASE_2`: bases e aba do inventário.
- `CHECKLIST_ARQUIVO_EQUIPAMENTOS` e `CHECKLIST_ARQUIVO_LEITURAS`: cadastro e histórico de checklists.
- `PRAZO_DIAS` e `CHECKLIST_PRAZO_DIAS`: prazos das duas automações.
- `<FILIAL>_EMAIL_1` até `<FILIAL>_EMAIL_9`: destinatários de cada grupo. Campos vazios não recebem mensagens.
- `GMAIL_CREDENTIALS_FILE` e `GMAIL_TOKEN_FILE`: caminhos dos arquivos OAuth. Caminhos relativos são resolvidos a partir da raiz do projeto.

Os dois scripts carregam exclusivamente o `.env`; variáveis já definidas no ambiente têm prioridade. O `.env.example` deve permanecer sem configurações pessoais.

Exemplo fictício dos caminhos e prazos no `.env`:

```dotenv
SHAREPOINT_DIR=C:/Dados/Inventario

# Paleteiras: Base 1 = leituras; Base 2 = cadastro dos equipamentos.
ARQUIVO_BASE_1="Gestão Manutenção de Frotas - Controle de Paleteiras.xlsx"
ARQUIVO_BASE_2="Planilha de Paleteiras Disktrans.xlsx"
ABA_BASE_2=Base
PRAZO_DIAS=7

# Checklists: cadastro e histórico de respostas.
CHECKLIST_ARQUIVO_EQUIPAMENTOS="Base Checklist.xlsx"
CHECKLIST_ARQUIVO_LEITURAS="Especifico Cliente - Relatorio BI - Checklist Veicular.xlsx"
CHECKLIST_PRAZO_DIAS=7

GMAIL_CREDENTIALS_FILE=credentials.json
GMAIL_TOKEN_FILE=token.json
```

`SHAREPOINT_DIR` aponta para arquivos disponíveis no computador; o código não baixa as planilhas do SharePoint. Se a pasta for sincronizada, aguarde a atualização e a disponibilidade local dos arquivos antes de executar.

Para enviar mensagens, disponibilize as credenciais OAuth do Gmail no caminho configurado. Na primeira autenticação, o fluxo solicita acesso à conta e grava o token localmente.

## Execução

Execute os comandos no PowerShell, dentro da pasta do projeto. Não é necessário ativar o ambiente virtual ao usar o caminho completo do Python abaixo.

Para analisar **somente checklists**:

```powershell
.\.venv\Scripts\python.exe checklist.py
```

Para analisar **somente paleteiras**:

```powershell
.\.venv\Scripts\python.exe paleteiras.py
```

**Executar os scripts gera os relatórios e envia e-mails aos grupos com problemas e destinatários configurados.** Para executar somente a geração de relatórios, deixe os destinatários vazios no `.env` e verifique se não há destinatários definidos nas variáveis do ambiente.

Uma rotina de utilização é:

1. Atualizar as duas bases da automação que será executada.
2. Conferir os nomes dos arquivos, o prazo e os destinatários no `.env`.
3. Fechar os relatórios Excel anteriores para permitir sua gravação.
4. Executar o script desejado e aguardar a mensagem de conclusão.
5. Abrir os arquivos em `saida/` e conferir os totais e as pendências.

Os arquivos de saída são sobrescritos a cada execução. O código não possui agendamento nem controle para impedir o reenvio da mesma cobrança: executar novamente pode enviar novos e-mails enquanto houver pendências.

## Inventário de paleteiras — `paleteiras.py`

### Bases necessárias

**Base 1 — leituras realizadas**, definida por `ARQUIVO_BASE_1`:

- O script lê a primeira aba e espera o cabeçalho na **linha 7 do Excel** (`header=6`).
- As colunas obrigatórias são:

| Coluna | Conteúdo |
| --- | --- |
| `CODIGO` | Código lido, usado para identificar o QR Code |
| `FILIAL` | Filial informada na leitura; obrigatória na entrada, mas não comparada com a filial do cadastro |
| `NOME` | Pessoa que realizou a leitura |
| `DATAHORA` | Data e hora da leitura |

Datas do Excel, datas em formato ISO e textos com dia primeiro são tratados pelo leitor. Registros sem QR válido ou sem data válida são descartados.

**Base 2 — cadastro de paleteiras**, definida por `ARQUIVO_BASE_2`:

- O script lê a aba definida em `ABA_BASE_2` — por padrão, `Base` — com o cabeçalho na **linha 1**.
- As colunas obrigatórias são:

| Coluna | Conteúdo |
| --- | --- |
| `FILIAL` | Filial responsável pelo equipamento e pelo recebimento da cobrança |
| `QR Code` | Identificação usada no cruzamento com as leituras |
| `NR_DISK` | Identificação patrimonial exibida no relatório |
| `MODELO` | Modelo da paleteira |
| `CUSTO modelo` | Custo cadastrado, usado na soma dos valores pendentes |

O cadastro define quais equipamentos serão analisados. Uma leitura sem equipamento correspondente no cadastro não entra no resultado. Linhas do cadastro sem QR válido são descartadas.

### Cruzamento e classificação

O código normaliza `CODIGO` e `QR Code`: utiliza a parte anterior ao primeiro hífen, remove o sufixo `.0`, mantém os dígitos e completa com zeros à esquerda até ter pelo menos cinco posições. Por exemplo, `123 - descrição` e `00123` correspondem ao QR `00123`.

**A Base 1 precisa conter apenas a leitura mais recente de cada QR normalizado, e a Base 2 deve ter um cadastro por QR.** Atualmente, `paleteiras.py` cruza as linhas sem ordenar e eliminar leituras anteriores. Se houver várias leituras do mesmo QR, o equipamento poderá aparecer repetido e aumentar os totais e custos. A consolidação das leituras deve ser feita na base antes da execução.

O vencimento é a data/hora da leitura acrescida de `PRAZO_DIAS`. O status é calculado em relação à data/hora do computador no momento da execução:

| Status | Condição | Entra na cobrança? |
| --- | --- | --- |
| `LIDO` | Há leitura e o vencimento ainda não foi atingido | Não |
| `PENDENTE` | Há leitura e o vencimento foi atingido | Sim |
| `NUNCA LIDO` | Não foi encontrada leitura válida na base fornecida | Sim |

Com prazo de 7 dias, uma leitura em 01/09 às 10h vence em 08/09 às 10h. A partir desse horário, passa a `PENDENTE`. `NUNCA LIDO` descreve a ausência de registro no arquivo analisado, não comprova que o equipamento jamais foi inventariado.

O resumo calcula `NAO_LIDOS = PENDENTES + NUNCA_LIDOS` e soma o custo desses dois grupos em `VALOR_PENDENTE`. Custos vazios ou que não podem ser convertidos são tratados como zero. O valor representa o custo cadastrado dos equipamentos pendentes, não uma perda confirmada.

A filial usada para agrupar e cobrar é a `FILIAL` da **Base 2**. Esta automação não verifica divergência entre a filial da leitura e a do cadastro, não valida tipo de checklist e não aplica a coluna `Ativo` do cadastro de checklists.

### Saídas e e-mail das paleteiras

| Arquivo em `saida/` | Conteúdo |
| --- | --- |
| `Resultado_Inventario.xlsx` | Equipamentos analisados, QR, patrimônio, modelo, custo, pessoa (`NOME`), leitura (`DATAHORA`), vencimento, status e dias sem leitura |
| `Resumo_Filiais.xlsx` | Totais por filial, lidos, pendentes, nunca lidos, percentuais e valor pendente |
| `Resumo_Visual_Paleteiras.pdf` | Gráficos, indicadores, resumo e detalhes dos equipamentos |

O e-mail da filial lista somente `PENDENTE` e `NUNCA LIDO`, com QR, patrimônio, modelo, última leitura, dias sem leitura, custo e status. A pessoa que realizou a leitura consta no Excel e nos detalhes do PDF; não aparece na tabela do corpo desse e-mail.

O anexo `Resumo_Paleteiras_<FILIAL>.pdf` inclui todos os equipamentos daquela filial, inclusive os lidos no prazo, para que os percentuais usem o total correto. Filiais sem pendências ou sem destinatários configurados não recebem mensagens.

## Checklists de equipamentos — `checklist.py`

### Bases necessárias

**Cadastro**, definido por `CHECKLIST_ARQUIVO_EQUIPAMENTOS`: primeira aba, cabeçalho na **linha 1**, com `Placa`, `Código Filial`, `Filial` e `Tipo de Equipamento`. A coluna `Ativo` é opcional.

**Histórico**, definido por `CHECKLIST_ARQUIVO_LEITURAS`: primeira aba, cabeçalho na **linha 6 do Excel** (`header=5`), com estas colunas obrigatórias:

| Coluna | Utilização |
| --- | --- |
| `PLACAVEICULO` | Placa para cruzamento com o cadastro |
| `NOMECHECKLIST` | Tipo de checklist realizado |
| `CODIGOFILIALRESPOSTA` | Código da filial informado na resposta |
| `DATARESPOSTA` | Data/hora usada para selecionar o último checklist e calcular o prazo |
| `ANOMES` | Ano/mês que auxilia na interpretação de datas textuais ambíguas |
| `RESPONDIDOPOR` | Responsável apresentado nos relatórios e no e-mail |
| `DATA` | Campo exigido na importação; o prazo usa `DATARESPOSTA` |
| `RESPOSTALOG` | Campo exigido e carregado do relatório de origem |

O histórico pode conter várias respostas por placa. O código descarta datas inválidas da seleção e mantém o registro de data/hora mais recente por placa, mesmo quando esse registro tem tipo ou filial incorretos.

### Regras de análise

- O cadastro usa `Placa`, `Código Filial`, `Filial` e `Tipo de Equipamento`. A coluna opcional `Ativo` permite excluir equipamentos da análise: `Não` desabilita a placa inteira, inclusive quando há várias linhas dela. Coluna ausente ou valor vazio mantém o equipamento ativo.
- A análise seleciona o último checklist por placa e verifica prazo, tipo e código de filial. O prazo vence ao completar o número de dias configurado.
- Para baterias, somente `Bateria` é aceito. Se o último registro for uma troca inicial, troca final ou outro tipo, o equipamento fica pendente mesmo com um registro recente.
- `STATUS_CHECKLIST` inclui a pendência de bateria; `STATUS_PRAZO` indica exclusivamente a situação do prazo.
- `MANUTENCAO` é um grupo de cobrança. A filial esperada é exibida como `MANUTENCAO`, e a validação usa o código individual cadastrado para cada equipamento.
- `BACKUP` e `FROTAS` têm mapeamentos de exibição em `FILIAIS_OPERACIONAIS`, atualmente para CON quando o código cadastrado é 2. Os grupos mantêm seus próprios destinatários.

Exemplo de manutenção: uma placa cadastrada como `MANUTENCAO` com código 7 precisa ter a resposta no código 7. Outra placa do mesmo grupo com código 20 precisa da resposta no código 20. Ambas são cobradas em `MANUTENCAO`.

Um equipamento pode estar no prazo e apresentar divergência de tipo ou filial. Por isso, o total `OK` do resumo de checklists não significa ausência de toda inconsistência. A lista de problemas inclui pendências **ou** divergências; no PDF, o indicador `Em dia` considera o cumprimento de prazo, tipo e filial. Um mesmo equipamento pode aparecer nas duas tabelas do e-mail quando possui pendência e inconsistência.

### Saídas e e-mail dos checklists

| Arquivo em `saida/` | Conteúdo |
| --- | --- |
| `Resultado_Checklist.xlsx` | Todos os equipamentos ativos analisados, último registro, responsável, prazo, status e inconsistências |
| `Problemas_Checklist.xlsx` | Somente equipamentos com pendência ou inconsistência |
| `Resumo_Checklist.xlsx` | Totais, status e contagens de divergências por grupo |
| `Resumo_Visual_Checklist.pdf` | Indicadores, gráficos, resumo e detalhes |

O e-mail mantém as tabelas de checklists pendentes e de inconsistências, incluindo o responsável. O anexo `Resumo_Checklist_<FILIAL>.pdf` reúne todos os equipamentos ativos do grupo, inclusive os regulares. Grupos sem problemas ou sem destinatários não recebem mensagens.

## Configuração e teste do envio de e-mail

Os destinatários são definidos por grupo, usando de 1 a 9 campos. Exemplo fictício:

```dotenv
CON_EMAIL_1=responsavel@example.com
CON_EMAIL_2=
MANUTENCAO_EMAIL_1=manutencao@example.com
```

As siglas precisam corresponder às aceitas em `carregar_destinatarios()` de cada script. Criar uma nova chave no `.env` não adiciona uma filial à lista do código. Há diferenças entre as listas atuais: por exemplo, checklist usa `MAH` e `SAO`, enquanto paleteiras usa `MHA` e `SÃO`; os grupos `BACKUP`, `FROTAS` e `MANUTENCAO` estão na lista do checklist.

Para testar um e-mail, deixe preenchida somente uma filial com seu endereço no `.env` e execute apenas a automação desejada. Esse grupo precisa ter alguma pendência ou inconsistência que gere cobrança. Os demais destinatários devem ficar vazios, inclusive nas variáveis do ambiente. Como a configuração é compartilhada, a mesma chave, como `CON_EMAIL_1`, serve às duas automações.

Na primeira tentativa de envio, o código abre o fluxo de autorização do Gmail se não houver token válido. Depois, reutiliza o `token.json` e tenta renová-lo quando necessário. O remetente é a conta autorizada nesse fluxo. Os relatórios são salvos antes do envio; uma falha de autenticação pode ocorrer mesmo depois de os arquivos terem sido gerados.

## Problemas comuns

| Situação | O que conferir |
| --- | --- |
| Base não encontrada | `SHAREPOINT_DIR`, nome do arquivo e disponibilidade local da planilha |
| Colunas não encontradas | Grafia dos cabeçalhos e linha esperada: leituras de paleteiras na 7, histórico de checklist na 6, cadastros na 1 |
| Aba não encontrada | `ABA_BASE_2` deve ser igual ao nome da aba do cadastro de paleteiras |
| Equipamento aparece como nunca lido/realizado | Identificador no cadastro e na base de registros, validade da data e abrangência do arquivo exportado |
| Paleteira repetida no resultado | QR normalizado duplicado no cadastro ou mais de uma leitura por QR na Base 1 |
| Nenhum e-mail enviado | Existência de problemas para o grupo, sigla aceita pelo script e destinatários preenchidos |
| Erro ao salvar Excel | Feche o relatório aberto e confira a permissão de escrita em `saida/` |
| Erro de importação de módulo | Instale as dependências no mesmo Python usado para executar e mantenha a pasta `automacao/` junto dos scripts |

## Desenvolvimento e testes

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Os testes simulam o envio de mensagens; não enviam e-mails reais. Alguns testes criam planilhas e PDFs temporários para verificar os arquivos produzidos.

## Arquivos para o GitHub

Versione o código, os testes, a logo, a documentação, as dependências e o `.env.example` vazio. O `.gitignore` exclui configurações locais, credenciais e tokens, bases Excel/CSV/Parquet, PDFs, relatórios, logs, caches e ambientes virtuais. Não use `git add -f` para incluir esses arquivos.

Antes de publicar, confira `git status --short` e o conteúdo das alterações. O `.gitignore` não remove arquivos já presentes em commits antigos.
