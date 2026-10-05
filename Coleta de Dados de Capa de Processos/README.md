# 🤖 Automação de Consulta e Atualização de Processos

Automação desenvolvida em Python e Selenium para realizar consultas de processos em um sistema web, coletar informações relacionadas e consolidar os resultados automaticamente em uma planilha Excel.

O robô recebe uma lista de processos a partir de um arquivo Excel, realiza a autenticação no sistema, pesquisa cada processo individualmente, coleta informações adicionais e gera uma nova planilha com os dados obtidos.

## 🎯 Objetivo

Automatizar um processo de consulta e coleta de informações que normalmente exigiria a execução manual das seguintes etapas:

* Acessar o sistema;
* Realizar login;
* Pesquisar cada processo;
* Abrir os resultados;
* Consultar informações relacionadas;
* Coletar os dados;
* Organizar os resultados em uma planilha.

## ⚙️ Funcionamento

```text
Resultado.xlsx
      ↓
Leitura dos processos
      ↓
Abertura do navegador
      ↓
Login automático
      ↓
Pesquisa dos processos
      ↓
Acesso aos detalhes
      ↓
Coleta das informações
      ↓
Organização dos resultados
      ↓
PROCESSOS ATUALIZADOS.xlsx
```

## 🔐 Gerenciamento de Credenciais

As credenciais utilizadas para acesso ao sistema são armazenadas separadamente em um arquivo JSON.

```text
credentials.json
```

O programa carrega as informações durante a execução:

```python
with open('credentials.json', 'r') as read_file:
    credenciais = json.load(read_file)
```

O arquivo `credentials.json` não deve ser publicado em um repositório público.

Recomenda-se adicioná-lo ao `.gitignore`:

```text
credentials.json
```

## 📥 Entrada de Dados

Os processos são carregados a partir do arquivo:

```text
Resultado.xlsx
```

O programa utiliza as seguintes colunas:

| Coluna           | Finalidade                                     |
| ---------------- | ---------------------------------------------- |
| `sem_formatacao` | Número do processo utilizado na pesquisa       |
| `formatado`      | Número do processo utilizado para apresentação |

## 🌐 Automação do Navegador

O projeto utiliza Selenium WebDriver para controlar o navegador Chrome.

Entre as operações automatizadas estão:

* Abertura do sistema;
* Login;
* Navegação pelos menus;
* Preenchimento do campo de pesquisa;
* Consulta dos processos;
* Abertura dos resultados;
* Extração das informações;
* Retorno à tela de pesquisa.

## 🔎 Consulta dos Processos

Para cada processo presente na planilha, o robô:

1. Obtém o número do processo;
2. Insere o número no campo de pesquisa;
3. Executa a pesquisa;
4. Acessa o resultado encontrado;
5. Aguarda o carregamento dos detalhes;
6. Coleta as informações disponíveis;
7. Armazena os dados;
8. Retorna à tela de consulta;
9. Continua para o próximo processo.

## 🔄 Tratamento de Tentativas

O código possui uma lógica de repetição para situações em que o resultado da pesquisa não é carregado imediatamente.

Caso o resultado não seja localizado, o robô realiza uma nova tentativa de pesquisa.

Durante a execução, o programa informa no terminal a quantidade de tentativas realizadas.

## 📊 Informações Coletadas

Para cada processo consultado, são armazenadas:

* Número principal do processo;
* Número antigo;
* Número originário;
* Número do processo formatado.

Quando determinadas informações não estão disponíveis, o sistema registra valores indicativos, como:

```text
Sem numero antigo
```

ou:

```text
Sem dados do processo
```

## 📤 Resultado

Ao finalizar as consultas, os dados são organizados utilizando Pandas e exportados para:

```text
PROCESSOS ATUALIZADOS.xlsx
```

A planilha final contém:

| Coluna               | Descrição                           |
| -------------------- | ----------------------------------- |
| `Principal`          | Número principal do processo        |
| `Número antigo`      | Número antigo associado ao processo |
| `Originário`         | Número do processo originário       |
| `Processo Formatado` | Número formatado do processo        |

## 🛠️ Tecnologias Utilizadas

* Python
* Selenium
* Pandas
* Excel
* JSON
* Chrome WebDriver
* Web Automation
* RPA

### Bibliotecas

```text
selenium
pandas
webdriver-manager
openpyxl
```

## 📂 Estrutura do Projeto

```text
📁 automacao-processos/
│
├── 📄 automacao.py
├── 📄 credentials.example.json
├── 📄 Resultado_exemplo.xlsx
├── 📄 README.md
└── 📄 .gitignore
```

## ▶️ Como Executar

### 1. Instalar as dependências

```bash
pip install selenium pandas openpyxl webdriver-manager
```

### 2. Configurar as credenciais

Crie um arquivo:

```text
credentials.json
```

seguindo a estrutura esperada pelo programa.

### 3. Preparar a planilha

Disponibilize o arquivo:

```text
Resultado.xlsx
```

com as colunas necessárias para a execução da automação.

### 4. Executar

```bash
python automacao.py
```

O navegador será aberto automaticamente e o robô iniciará o processo de consulta.

Ao final, será gerado:

```text
PROCESSOS ATUALIZADOS.xlsx
```

## 🔐 Privacidade e Segurança

Este projeto foi adaptado para apresentação em portfólio.

Dados reais, informações pessoais, credenciais, URLs internas e demais informações confidenciais devem ser removidos ou substituídos antes da publicação do projeto.

## 💡 Conceitos Demonstrados

* Automação de processos (RPA)
* Automação de navegador
* Selenium WebDriver
* Extração de informações
* Manipulação de dados com Pandas
* Integração entre automação web e Excel
* Leitura e geração de arquivos `.xlsx`
* Gerenciamento externo de credenciais
* Tratamento de exceções
* Processamento de dados em lote
* Automação de tarefas repetitivas

## 🚀 Possíveis Melhorias

* Substituição de `sleep()` por `WebDriverWait`;
* Utilização de seletores mais robustos;
* Implementação de sistema de logs;
* Tratamento específico para erros de conexão;
* Limite de tentativas para evitar loops indefinidos;
* Configuração externa de URLs e parâmetros;
* Utilização de variáveis de ambiente para credenciais;
* Criação de interface gráfica;
* Geração de relatório final da execução.

## 📌 Observação

O projeto foi desenvolvido para automatizar uma rotina específica de consulta e coleta de informações. Para utilização em outros sistemas, os seletores, URLs e etapas de navegação precisam ser adaptados à estrutura da aplicação utilizada.
