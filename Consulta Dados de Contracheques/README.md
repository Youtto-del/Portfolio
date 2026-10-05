# 🤖 Automação de Coleta e Organização de Contracheques

Automação desenvolvida em Python e Selenium para realizar consultas de contracheques em um sistema web, coletar informações de diferentes períodos e organizar os resultados automaticamente em arquivos Excel e PDF.

O projeto combina automação de navegador, processamento de dados, geração de documentos e organização de arquivos para reduzir tarefas manuais e repetitivas relacionadas à consulta de informações.

## 🎯 Objetivo

Automatizar o processo de consulta e organização de contracheques para diferentes pessoas e períodos.

A aplicação permite:

* Acessar automaticamente o sistema;
* Realizar autenticação;
* Consultar diferentes competências;
* Identificar diferentes tipos de folha;
* Capturar informações apresentadas pelo sistema;
* Exportar os dados para Excel;
* Gerar versões em PDF;
* Consolidar os arquivos por período;
* Organizar os resultados em diretórios individuais.

## ⚙️ Funcionamento

```text
Dados.xlsx
    ↓
Leitura dos dados
    ↓
Login no sistema
    ↓
Consulta da competência
    ↓
Identificação da folha
    ↓
Coleta das informações
    ↓
Captura da página
    ↓
Geração do PDF
    ↓
Extração dos dados para Excel
    ↓
Consolidação dos arquivos
    ↓
Resultados organizados
```

## 📥 Entrada de Dados

O programa utiliza uma planilha Excel chamada:

```text
Dados.xlsx
```

A planilha fornece os parâmetros necessários para cada execução, incluindo:

* Nome;
* Mês inicial;
* Ano inicial;
* Mês final;
* Ano final;
* Login;
* Senha.

As credenciais são carregadas durante a execução e utilizadas para autenticação no sistema.

## 🔐 Autenticação

O acesso ao sistema é realizado automaticamente utilizando Selenium.

A função `acesso()` recebe o login e a senha e realiza:

1. Abertura do sistema;
2. Navegação até a área de consulta;
3. Preenchimento do usuário;
4. Preenchimento da senha;
5. Login;
6. Acesso à área de contracheques.

As credenciais não ficam gravadas diretamente no código-fonte.

## 🌐 Automação do Navegador

O projeto utiliza Selenium WebDriver para controlar o Google Chrome.

Entre as operações automatizadas estão:

* Navegação entre páginas;
* Preenchimento de formulários;
* Seleção de competências;
* Seleção do tipo de folha;
* Acesso a conteúdo dentro de `iframe`;
* Extração de informações;
* Captura da página;
* Navegação entre diferentes consultas.

## 📅 Consulta por Período

A aplicação calcula automaticamente a quantidade de meses que precisam ser consultados com base no período inicial e final informado.

O processo percorre as competências sequencialmente e atualiza automaticamente:

* Mês;
* Ano;
* Quantidade de consultas;
* Tipo de folha.

O código também possui tratamento específico para períodos que envolvem a folha do 13º salário.

## 📄 Tipos de Folha

O programa identifica a quantidade de opções disponíveis no campo de folha e determina automaticamente o tipo de consulta.

São tratados cenários como:

```text
Folha mensal
Folha 13
Folha mensal única
```

Isso permite que o processo se adapte à quantidade de opções disponibilizadas pelo sistema para cada competência.

## 📸 Captura da Página

Após a consulta, o programa utiliza o Chrome DevTools Protocol para realizar uma captura da página:

```python
navegador.execute_cdp_cmd(
    "Page.captureScreenshot",
    {"format": "png", "captureBeyondViewport": True}
)
```

A captura é inicialmente armazenada como PNG e posteriormente convertida para PDF.

## 📑 Geração de PDF

As capturas de tela são convertidas em documentos PDF utilizando ReportLab.

Cada período consultado pode gerar um PDF individual.

Posteriormente, os PDFs são combinados em um único documento por período utilizando `pypdf`.

## 📊 Extração para Excel

As informações apresentadas na página são extraídas utilizando Selenium e transformadas em um DataFrame do Pandas.

O resultado é salvo em arquivos Excel contendo:

```text
Descrição | Valor
```

Os arquivos individuais são posteriormente consolidados em uma única planilha.

## 📚 Consolidação dos Resultados

A função `unir_arquivos()` realiza duas etapas principais.

### Excel

Os arquivos individuais de uma determinada competência são reunidos em um único arquivo:

```text
Contracheques [ano].xlsx
```

Cada conjunto de informações é organizado em uma aba separada.

### PDF

Os PDFs correspondentes ao período também são reunidos em um único documento:

```text
Capturas [ano].pdf
```

Dessa forma, os resultados ficam centralizados e organizados.

## 📁 Organização dos Arquivos

Para cada pessoa, o programa cria automaticamente um diretório utilizando o nome tratado para evitar problemas com caracteres especiais:

```text
Contracheques_Nome
```

Os arquivos são armazenados dentro desse diretório.

Exemplo:

```text
Contracheques_Nome/
│
├── 2025_01_pdf_contracheque_mensal.pdf
├── 2025_01_xlsx_contracheque_mensal.xlsx
├── 2025_02_pdf_contracheque_mensal.pdf
├── 2025_02_xlsx_contracheque_mensal.xlsx
├── Contracheques 2025.xlsx
└── Capturas 2025.pdf
```

## 🛠️ Tecnologias Utilizadas

* Python
* Selenium
* Pandas
* Excel
* ReportLab
* PyPDF
* Pillow
* Chrome WebDriver
* EasyGUI
* Regular Expressions
* Base64
* Web Automation
* RPA

## 📦 Principais Bibliotecas

```text
pandas
selenium
Pillow
reportlab
pypdf
easygui
```

## ▶️ Como Executar

Instale as dependências necessárias:

```bash
pip install pandas selenium pillow reportlab pypdf easygui openpyxl
```

Prepare a planilha de entrada:

```text
Dados.xlsx
```

Configure os dados necessários para cada consulta e execute:

```bash
python automacao.py
```

O navegador será controlado automaticamente durante o processo e os resultados serão armazenados nos diretórios correspondentes.

## 🔐 Privacidade e Segurança

Este projeto trabalha com informações potencialmente sensíveis, incluindo dados de identificação, credenciais de acesso e informações financeiras relacionadas a contracheques.

Por esse motivo, dados reais não devem ser publicados no repositório.

Não publique:

* `Dados.xlsx` contendo dados reais;
* Contracheques em PDF;
* Planilhas de resultados;
* Credenciais;
* Capturas de tela reais;
* Informações financeiras;
* Dados pessoais.

Para demonstrações públicas, recomenda-se utilizar dados fictícios ou completamente anonimizados.

## 💡 Conceitos Demonstrados

* Robotic Process Automation (RPA);
* Automação de navegador;
* Web scraping;
* Selenium WebDriver;
* Manipulação de dados com Pandas;
* Automação de Excel;
* Geração de PDF;
* Conversão de imagens para PDF;
* Consolidação de documentos;
* Processamento de dados por período;
* Automação de tarefas repetitivas;
* Organização automatizada de arquivos;
* Manipulação de `iframe`;
* Automação utilizando Chrome DevTools Protocol.

## 🚀 Possíveis Melhorias

* Utilização de variáveis de ambiente para credenciais;
* Separação das configurações em arquivo externo;
* Substituição de `sleep()` por `WebDriverWait` em todos os pontos necessários;
* Implementação de logs estruturados;
* Tratamento específico de erros de autenticação;
* Validação dos dados de entrada;
* Sistema de recuperação após falhas;
* Interface gráfica para configuração da execução;
* Geração de relatório final;
* Melhor controle de arquivos temporários;
* Utilização de seletores mais robustos;
* Criação de configuração para diferentes ambientes.

## 📌 Observação

O projeto foi desenvolvido para automatizar uma rotina específica de consulta e organização de documentos.

Os seletores, URLs e etapas de navegação dependem da estrutura do sistema utilizado e podem precisar de adaptação caso a aplicação seja utilizada em outro ambiente.
