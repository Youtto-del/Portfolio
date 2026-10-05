# 📄 Automação de Petições com Python

Automação desenvolvida em Python para **gerar automaticamente documentos de petição a partir de dados armazenados em uma planilha Excel**.

O sistema lê os dados do arquivo `Resultado.xlsx`, identifica os registros que atendem a um determinado critério e utiliza um documento Word como modelo para gerar automaticamente uma petição individual para cada registro.

O projeto foi desenvolvido com foco na **redução de tarefas manuais, padronização de documentos e ganho de produtividade**.

---

## 🎯 Objetivo

Automatizar a criação de múltiplas petições que, de forma manual, exigiriam a abertura de um modelo Word e o preenchimento individual de informações como:

* Nome
* Comarca
* Juízo
* Número do processo
* Data

A automação permite gerar os documentos em lote a partir das informações disponíveis na planilha.

---

## ⚙️ Como funciona

O processo de execução pode ser resumido em:

```text
Resultado.xlsx
      ↓
Leitura dos dados com Pandas
      ↓
Filtragem dos registros
      ↓
Carregamento do modelo Word
      ↓
Substituição dos campos
      ↓
Geração das petições
      ↓
Pasta "Peticoes"
```

### 1. Leitura da planilha

O programa utiliza **Pandas** para carregar os dados do arquivo:

```text
Resultado.xlsx
```

A planilha deve conter informações utilizadas no preenchimento dos documentos, como:

* `Nome`
* `Juizo`
* `Comarca`
* `Processos`
* `Data do Pagamento`

---

### 2. Filtragem dos registros

O sistema verifica a coluna **Data do Pagamento** e seleciona os registros que possuem o valor:

```text
Sem dados
```

Somente esses registros são utilizados para gerar as petições.

---

### 3. Carregamento do modelo

Para cada registro selecionado, o programa abre o arquivo:

```text
modelo.docx
```

Esse arquivo funciona como um **template**, contendo marcadores que serão substituídos automaticamente.

---

### 4. Substituição das informações

Os marcadores presentes no documento são associados aos dados da planilha:

| Marcador | Informação         |
| -------- | ------------------ |
| `XXXX`   | Nome               |
| `YYYY`   | Comarca            |
| `WWWW`   | Data atual         |
| `QQQQ`   | Número do processo |
| `ZZZZ`   | Juízo              |

O programa substitui esses marcadores pelas informações correspondentes de cada registro.

---

### 5. Geração dos documentos

Após o preenchimento, cada petição é salva individualmente na pasta:

```text
Peticoes/
```

Os arquivos são nomeados seguindo o padrão:

```text
Petição - NOME.docx
```

---

## 🛠️ Tecnologias utilizadas

* **Python**
* **Pandas** — leitura e filtragem dos dados da planilha
* **python-docx** — manipulação de documentos Word
* **docxedit** — substituição de textos no documento
* **datetime** — obtenção e formatação da data
* **locale** — formatação da data em português
* **os** — criação do diretório de saída

---

## 📂 Estrutura esperada

```text
📁 projeto/
│
├── 📄 automacao.py
├── 📄 Resultado.xlsx
├── 📄 modelo.docx
│
└── 📁 Peticoes/
    ├── Petição - NOME 1.docx
    ├── Petição - NOME 2.docx
    └── ...
```

A pasta `Peticoes` é criada automaticamente durante a execução.

---

## ▶️ Como executar

### 1. Instalar as dependências

```bash
pip install pandas python-docx docxedit openpyxl
```

### 2. Preparar os arquivos

Coloque na mesma pasta do programa:

```text
automacao.py
Resultado.xlsx
modelo.docx
```

Certifique-se de que a planilha possui as colunas esperadas pelo programa.

### 3. Executar

```bash
python automacao.py
```

Após a execução, as petições geradas estarão disponíveis na pasta:

```text
Peticoes/
```

---

## 🔐 Privacidade e dados

Este projeto foi adaptado para fins de demonstração e portfólio.

**Dados reais, informações pessoais, números de processos e demais informações sensíveis devem ser removidos ou substituídos antes da publicação do projeto em um repositório público.**

O arquivo `Resultado.xlsx` utilizado em ambiente real não deve ser disponibilizado publicamente caso contenha informações pessoais ou confidenciais.

---

## 💡 Aplicação prática

A solução demonstra a utilização de Python para automatizar um processo baseado em:

**Dados estruturados → processamento → preenchimento de template → geração de documentos.**

Esse tipo de automação pode ser aplicado a diferentes cenários que exigem a criação repetitiva de documentos personalizados a partir de informações armazenadas em planilhas ou bancos de dados.

---

## 📌 Competências demonstradas

Este projeto demonstra conhecimentos em:

* Automação de processos
* Python
* Manipulação de arquivos Excel
* Processamento de dados com Pandas
* Automação de documentos
* Manipulação de arquivos `.docx`
* Criação de documentos em lote
* Tratamento e transformação de dados
* Organização de arquivos
