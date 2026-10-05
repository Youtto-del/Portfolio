# 🤖 Automação de Coleta e Análise de Intimações

Automação desenvolvida em Python e Selenium para realizar consultas em um sistema web, coletar informações de processos e identificar eventos relacionados a prazos e intimações.

O programa recebe uma lista de processos a partir de uma planilha Excel, realiza o acesso ao sistema, consulta cada processo individualmente, coleta informações relevantes e consolida os resultados automaticamente em uma nova planilha.

O projeto demonstra a aplicação de **RPA (Robotic Process Automation)** para automatizar tarefas repetitivas de consulta, extração e organização de dados.

## 🎯 Objetivo

Automatizar um processo de consulta que envolve a análise individual de diversos processos e a coleta de informações relacionadas a intimações, eventos e partes processuais.

O robô automatiza etapas como:

* Acesso ao sistema;
* Autenticação;
* Pesquisa de processos;
* Consulta de informações processuais;
* Identificação de eventos;
* Coleta de dados;
* Organização das informações;
* Exportação dos resultados para Excel.

## ⚙️ Funcionamento

```text
intimacoes.xls
      ↓
Leitura dos processos
      ↓
Seleção do grau de jurisdição
      ↓
Acesso ao sistema
      ↓
Login automático
      ↓
Pesquisa dos processos
      ↓
Análise dos eventos
      ↓
Coleta das informações
      ↓
Organização dos dados
      ↓
Lista de notas.xlsx
```

## 🖥️ Seleção do Grau de Jurisdição

O programa utiliza uma interface gráfica simples para permitir a seleção entre:

* 1º Grau
* 2º Grau

A escolha determina o endereço do sistema utilizado durante a execução.

## 🔐 Gerenciamento de Credenciais

As credenciais de acesso são armazenadas separadamente em um arquivo JSON:

```text
credentials.json
```

O programa carrega as informações durante a execução:

```python
with open('credentials.json', 'r') as read_file:
    credenciais = json.load(read_file)
```

Essa abordagem evita armazenar usuário e senha diretamente no código-fonte.

O arquivo `credentials.json` não deve ser publicado em um repositório público.

Recomenda-se adicioná-lo ao `.gitignore`:

```text
credentials.json
```

## 📥 Entrada de Dados

Os processos são carregados a partir do arquivo:

```text
intimacoes.xls
```

O programa utiliza a planilha para obter a lista de processos que serão consultados.

## 🌐 Automação do Navegador

O projeto utiliza Selenium WebDriver para controlar o navegador Chrome.

Entre as operações automatizadas estão:

* Abertura do sistema;
* Login;
* Seleção do perfil;
* Pesquisa de processos;
* Navegação pelas informações processuais;
* Abertura de documentos;
* Coleta de dados;
* Navegação entre abas;
* Fechamento de páginas;
* Retorno à consulta principal.

## 🔎 Consulta dos Processos

Para cada processo encontrado na planilha, o robô:

1. Obtém o número do processo;
2. Realiza a pesquisa;
3. Consulta o processo;
4. Obtém o processo originário;
5. Coleta a data de distribuição;
6. Identifica as partes envolvidas;
7. Consulta informações relacionadas;
8. Analisa os eventos disponíveis;
9. Identifica eventos relacionados a prazos;
10. Coleta informações complementares;
11. Armazena os resultados.

## 📋 Análise de Eventos

O programa verifica diferentes tipos de eventos utilizando seletores CSS.

São analisados eventos relacionados a:

* Prazos aguardando abertura;
* Prazos em aberto;
* Eventos com diferentes estados visuais;
* Informações associadas aos eventos processuais.

Quando um evento relevante é identificado, o programa realiza a coleta das informações correspondentes.

## 📄 Coleta de Informações

Dependendo do evento encontrado, o programa pode coletar informações como:

* Número do processo;
* Processo originário;
* Data de distribuição;
* Nome da parte;
* CPF;
* Requerente;
* Requerido;
* Tipo da ação;
* Descrição do evento;
* Conteúdo relacionado à certidão.

Quando determinadas informações não estão disponíveis, o programa utiliza valores indicativos, como:

```text
Sem dados
```

ou:

```text
Sem número de precatório
```

## 📊 Resultado

Após concluir as consultas, os dados são organizados utilizando Pandas e exportados para:

```text
Lista de notas.xlsx
```

O arquivo contém informações como:

| Coluna              | Descrição                       |
| ------------------- | ------------------------------- |
| `Principal`         | Número principal do processo    |
| `Originário`        | Processo originário             |
| `Data distribuição` | Data de distribuição            |
| `Título`            | Título relacionado ao registro  |
| `Tipo`              | Tipo da ação                    |
| `Requerente`        | Parte requerente                |
| `CPF`               | CPF coletado no sistema         |
| `Requerido`         | Parte requerida                 |
| `Certidão`          | Informações coletadas do evento |

## 🛠️ Tecnologias Utilizadas

* Python
* Selenium
* Pandas
* Excel
* JSON
* EasyGUI
* Chrome WebDriver
* Web Automation
* RPA

### Bibliotecas

```text
selenium
pandas
xlrd
easygui
webdriver-manager
```

## 📂 Estrutura do Projeto

```text
📁 automacao-intimacoes/
│
├── 📄 automacao.py
├── 📄 credentials.example.json
├── 📄 intimacoes_exemplo.xls
├── 📄 README.md
└── 📄 .gitignore
```

## ▶️ Como Executar

### 1. Instalar as dependências

```bash
pip install selenium pandas xlrd easygui webdriver-manager
```

### 2. Configurar as credenciais

Crie um arquivo:

```text
credentials.json
```

seguindo a estrutura esperada pelo programa.

Não utilize credenciais reais em arquivos publicados no GitHub.

### 3. Preparar a planilha

Disponibilize o arquivo:

```text
intimacoes.xls
```

contendo os processos que serão consultados.

### 4. Executar

```bash
python automacao.py
```

Durante a execução, o programa poderá solicitar a resolução manual de um CAPTCHA caso ele seja apresentado pelo sistema.

Após a conclusão, será gerado:

```text
Lista de notas.xlsx
```

## 🔐 Privacidade e Segurança

Este projeto foi adaptado para apresentação em portfólio.

O processo automatizado pode trabalhar com informações pessoais e processuais, incluindo dados de identificação das partes.

Por esse motivo:

* Não publique arquivos Excel contendo dados reais;
* Não publique informações processuais reais;
* Não publique CPF ou outros dados pessoais;
* Não publique credenciais;
* Não publique URLs internas;
* Utilize dados fictícios ou anonimizados para demonstração.

## 💡 Conceitos Demonstrados

* Automação de processos (RPA);
* Automação de navegador;
* Selenium WebDriver;
* Web scraping;
* Extração estruturada de informações;
* Manipulação de dados com Pandas;
* Integração entre automação web e Excel;
* Leitura e geração de arquivos `.xls` e `.xlsx`;
* Gerenciamento externo de credenciais;
* Tratamento de exceções;
* Navegação entre múltiplas abas;
* Processamento de dados em lote;
* Automação de tarefas repetitivas.

## 🚀 Possíveis Melhorias

* Substituição de `sleep()` por `WebDriverWait`;
* Utilização de seletores mais robustos;
* Implementação de sistema de logs;
* Tratamento específico para erros de conexão;
* Limitação do número de tentativas;
* Utilização de variáveis de ambiente para credenciais;
* Configuração externa das URLs;
* Validação dos dados antes da exportação;
* Interface gráfica para configuração da execução;
* Geração de relatório de erros e processos não encontrados.

## 📌 Observação

O projeto foi desenvolvido para automatizar uma rotina específica de consulta e coleta de informações. Para utilização em outros sistemas, os seletores, URLs, campos e etapas de navegação precisam ser adaptados à estrutura da aplicação utilizada.
