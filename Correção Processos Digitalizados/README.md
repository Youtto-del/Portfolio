# 🤖 Automação de Triagem, Consulta e Preparação de Processos

Projeto de automação desenvolvido em Python para realizar triagem de intimações, consultar processos em um sistema web, identificar processos originários e preparar automaticamente uma planilha para importação.

O projeto combina **RPA, Selenium, Pandas e Excel** para automatizar diferentes etapas de uma rotina de consulta, análise, cruzamento e preparação de dados processuais.

## 🎯 Objetivo

Automatizar um fluxo composto por três etapas principais:

1. Download e tratamento de relatórios;
2. Consulta automatizada de processos;
3. Preparação dos dados para importação.

O fluxo reduz a necessidade de realizar manualmente consultas individuais, comparar informações entre diferentes planilhas e preparar os resultados para uma etapa posterior de importação.

## ⚙️ Fluxo do Projeto

```text
Sistema Web
     ↓
Download do relatório
     ↓
Triagem das intimações
     ↓
Identificação dos processos
     ↓
Consulta automatizada
     ↓
Coleta dos processos originários
     ↓
Coleta dos status
     ↓
Coleta das datas
     ↓
Cruzamento com relatório de desdobramentos
     ↓
Filtragem dos casos relevantes
     ↓
Identificação das pastas
     ↓
Preparação da planilha de importação
     ↓
Arquivo final para Smart Import
```

## 📂 Estrutura

O projeto é dividido em dois componentes principais:

```text
📁 projeto/
│
├── 📄 automacao.py
├── 📄 prepara_import.py
│
├── 📁 SmartImports/
│
└── 📄 README.md
```

### `automacao.py`

Responsável pela automação principal:

* Acesso ao sistema;
* Autenticação;
* Download de relatório;
* Importação de planilhas;
* Triagem de intimações;
* Consulta de processos;
* Coleta de processos originários;
* Coleta de status;
* Coleta de datas;
* Cruzamento inicial dos dados;
* Geração da lista de notas;
* Chamada da etapa de preparação da importação.

### `prepara_import.py`

Responsável pela etapa final de tratamento:

* Filtragem dos processos;
* Identificação dos casos relevantes;
* Cruzamento com o relatório de desdobramentos;
* Identificação das pastas;
* Construção do DataFrame final;
* Preenchimento da planilha modelo;
* Geração do arquivo final para importação.

## 🔐 Autenticação

As credenciais utilizadas pelo sistema são carregadas externamente através de:

```text
credentials.json
```

O usuário e a senha não são armazenados diretamente no código-fonte.

O arquivo de credenciais deve permanecer fora do repositório público.

Recomenda-se adicioná-lo ao `.gitignore`:

```text
credentials.json
```

## 🌐 Automação Web

O projeto utiliza Selenium WebDriver para controlar o navegador Chrome.

A automação realiza:

* Acesso ao sistema;
* Login;
* Seleção de perfil;
* Navegação;
* Download de relatórios;
* Pesquisa de processos;
* Consulta de informações;
* Extração de dados;
* Navegação entre páginas e janelas.

## 📥 Download e Tratamento de Relatórios

O projeto realiza automaticamente o download de uma planilha disponibilizada pelo sistema.

Após o download, o arquivo é carregado utilizando Pandas e tem sua estrutura tratada para facilitar a análise.

O processo inclui:

* Leitura da planilha;
* Identificação do cabeçalho;
* Renomeação das colunas;
* Remoção de linhas auxiliares;
* Preparação dos dados para consulta.

## 🔎 Triagem de Intimações

As intimações são comparadas com uma base de desdobramentos.

Os processos são classificados entre:

```text
Cadastrados
Não cadastrados
```

Essa etapa permite determinar quais processos precisam seguir para a etapa de consulta.

## 🔍 Consulta de Processos

Os processos selecionados são consultados automaticamente no sistema web.

Para cada processo, o robô coleta:

* Número do processo;
* Primeiro processo originário;
* Segundo processo originário;
* Terceiro processo originário;
* Status dos processos relacionados;
* Data de distribuição.

Os resultados são organizados em um DataFrame.

## 📊 Geração da Lista de Notas

Os dados coletados são exportados para:

```text
Lista de notas.xlsx
```

O arquivo contém informações estruturadas sobre os processos consultados e seus relacionamentos.

Entre os dados armazenados estão:

| Campo               | Descrição                     |
| ------------------- | ----------------------------- |
| `Processo`          | Processo consultado           |
| `originario_1`      | Primeiro processo originário  |
| `Status 1`          | Status do primeiro originário |
| `originario_2`      | Segundo processo originário   |
| `Status 2`          | Status do segundo originário  |
| `originario_3`      | Terceiro processo originário  |
| `Status 3`          | Status do terceiro originário |
| `Data distribuição` | Data de distribuição          |

## 🧠 Filtragem dos Processos

O módulo `prepara_import.py` utiliza a `Lista de notas.xlsx` para identificar processos que atendem aos critérios de status definidos no fluxo.

São considerados os casos relacionados aos status:

```text
Migrado
Digitalizado
```

Os registros selecionados são armazenados em um novo DataFrame:

```text
Processo
Originario
Status
```

O resultado é salvo em:

```text
Resultado filtrado.xlsx
```

## 🗃️ Identificação das Pastas

Após a filtragem, os processos são cruzados novamente com o relatório de desdobramentos.

Quando o processo originário é localizado na base, o programa identifica a pasta de desdobramento correspondente.

São então organizadas informações como:

```text
Pasta
Pasta desdobramento
Número antigo
Processo
```

## 📑 Preparação da Importação

O programa utiliza uma planilha modelo como base:

```text
Modelo EPROC ATT.xlsx
```

Uma cópia do modelo é criada automaticamente dentro do diretório:

```text
SmartImports/
```

O nome do arquivo recebe a data da execução.

Exemplo:

```text
Correcao Digit EPROC ATT - 051026.xlsx
```

Os dados processados são então adicionados à planilha `Importacao`.

## 📤 Resultado Final

Ao final do processo, é gerada uma planilha estruturada para utilização em uma etapa posterior de importação.

O fluxo final pode ser representado por:

```text
Lista de notas.xlsx
        ↓
Filtragem por status
        ↓
Resultado filtrado.xlsx
        ↓
Cruzamento com desdobramentos
        ↓
Identificação das pastas
        ↓
Planilha modelo
        ↓
SmartImports/Correcao Digit EPROC ATT - [data].xlsx
```

## 🛠️ Tecnologias Utilizadas

* Python
* Selenium
* Pandas
* Excel
* OpenPyXL
* JSON
* EasyGUI
* Chrome WebDriver
* Web Automation
* RPA
* Data Processing

## 📦 Principais Bibliotecas

```text
pandas
selenium
webdriver-manager
easygui
openpyxl
```

## ▶️ Execução

Instale as dependências:

```bash
pip install selenium pandas easygui webdriver-manager openpyxl
```

Configure as credenciais externamente e disponibilize os arquivos necessários para o fluxo.

Execute o script principal:

```bash
python automacao.py
```

O processo realizará automaticamente as etapas de consulta e preparação dos dados.

## 🔐 Privacidade e Segurança

O projeto trabalha com informações processuais e relatórios que podem conter dados sensíveis.

Os arquivos utilizados durante a execução podem conter:

* Números de processos;
* Processos originários;
* Status;
* Datas;
* Informações de intimações;
* Estruturas internas de organização.

Por esse motivo, arquivos reais de entrada e saída não devem ser publicados em um repositório público.

Não publique:

* `credentials.json`;
* Relatórios reais;
* Planilhas de processos;
* Resultados reais;
* Informações de clientes;
* Dados internos do sistema.

Para demonstrações públicas, utilize dados fictícios ou anonimizados.

## 💡 Conceitos Demonstrados

* RPA (Robotic Process Automation);
* Automação de navegador;
* Selenium WebDriver;
* Web scraping;
* Processamento de dados;
* Pandas;
* OpenPyXL;
* Manipulação de Excel;
* Download automatizado;
* Tratamento e limpeza de dados;
* Cruzamento de bases;
* Filtragem de dados;
* Automação de tarefas repetitivas;
* Geração automatizada de relatórios;
* Preparação de dados para importação.

## 🚀 Possíveis Melhorias

* Substituição de `sleep()` por `WebDriverWait`;
* Remoção de caminhos locais específicos;
* Utilização de variáveis de ambiente para credenciais;
* Implementação de logs estruturados;
* Tratamento específico de erros de conexão;
* Validação dos arquivos baixados;
* Seletores Selenium mais robustos;
* Configuração externa das URLs;
* Validação dos dados antes da importação;
* Melhor tratamento de arquivos temporários;
* Separação entre configuração, automação e processamento;
* Correção do controle de execução do módulo `prepara_import`;
* Implementação de testes para as etapas de tratamento dos DataFrames.

## 📌 Observação

Os seletores, URLs e estruturas utilizadas neste projeto dependem do sistema web para o qual a automação foi desenvolvida.

Para utilização em outro ambiente, essas configurações e etapas de navegação precisam ser adaptadas à estrutura correspondente.
