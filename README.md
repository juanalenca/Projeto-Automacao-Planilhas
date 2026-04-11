# 📊 Projeto Automação de Planilhas

> Transforme a extração e análise de dados contábeis em um processo rápido, seguro e totalmente automatizado.

O **Projeto Automação de Planilhas** é uma aplicação desktop desenvolvida em Python projetada para processar, extrair e consolidar dados de arquivos de texto referentes a relatórios de vantagens e descontos de folha de pagamento (ex: sistema Reciprev). A ferramenta elimina o trabalho manual e repetitivo através de extração automatizada via Expressões Regulares (Regex), alimentando simultaneamente planilhas Excel dinâmicas e um banco de dados relacional (SQLite).

---

## ✨ Principais Funcionalidades

- **Extração Inteligente de Dados:** Leitura e processamento de arquivos `.txt` brutos contendo relatórios financeiros de vantagens e descontos de pessoal.
- **Interface Gráfica Intuitiva:** Uma interface simples construída com `Tkinter`, permitindo seleção ágil de arquivos de texto e da planilha base.
- **Geração de Planilhas Dinâmicas:** 
  - Salva os dados processados em relatórios `.xlsx` individuais segmentados.
  - Alimenta de forma sequencial uma planilha base já existente, destacando novas inserções (amarelo) e organizando informações.
- **Relatório de Cálculo Final:** Criação e estruturação automática de uma aba hierárquica chamada `CÁLCULO FINAL FOLHA`, com agrupamentos por Órgão, Categoria, e Classificação de verbas.
- **Persistência de Dados e Histórico:** Armazenamento automático e relacional das verbas, códigos, órgãos e classificações no banco de dados SQLite (`contab_reciprev.db`).

---

## 🛠 Tecnologias Utilizadas

Este projeto foi construído utilizando o melhor do ecossistema Python para processamento de dados e interface:

*   **[Python 3.x](https://www.python.org/)** - Linguagem base do projeto.
*   **[Pandas](https://pandas.pydata.org/)** - Essencial para estruturação e exportação precisa de dados em lote para Excel.
*   **[OpenPyXL](https://openpyxl.readthedocs.io/)** - Utilizada para formatação, estilização (como cores e fontes) e injeção de dados em planilhas Excel já existentes.
*   **[Tkinter](https://docs.python.org/3/library/tkinter.html)** - Instanciamento da interface gráfica amigável para o usuário final.
*   **[SQLite3](https://docs.python.org/3/library/sqlite3.html)** - Banco de dados embutido e estruturado para persistência e classificação segura (script autônomo via `server.py`).
*   **Expressões Regulares (`re`)** - Motor de busca inteligente utilizado para localizar e quebrar os dados de tabelas irregulares contidas nos `.txt`.

---

## 🚀 Como Funciona a Estritetura

O projeto é dividido essencialmente em dois arquivos vitais:

1.  **`server.py`:**
    Script de inicialização do Banco de Dados. Responsável por executar instruções SQL e gerar o banco de dados e as relações normatizadas nas tabelas: `verbas`, `orgaos`, `classificacoes`, `categorias`, `elementos` e `codigo_orgaos`.
2.  **`planilhas.py`:**
    A aplicação em si (GUI). Engloba a lógica de upload dos arquivos, processamento e limpeza dos caracteres, injeção das métricas na base de dados (`contab_reciprev.db`) e geração/alteração das planilhas Excel.

---

## 💻 Instalação e Execução

### Pré-requisitos
Certifique-se de que o **Python 3.8+** esteja instalado em sua máquina e as bibliotecas a seguir estejam disponíveis em seu ambiente (utilizando pip):

```bash
pip install pandas openpyxl
```

### Configuração Inicial

1. **Clone este repositório:**
   ```bash
   git clone https://github.com/juanalenca/Projeto-Automacao-Planilhas.git
   cd Projeto-Automacao-Planilhas
   ```

2. **Gere o Banco de Dados local:**
   Antes de processar qualquer planilha, inicialize o banco de dados.
   ```bash
   python server.py
   ```
   > Você deverá ver no console: `Banco de dados 'contab_reciprev.db' criado com sucesso.`

3. **Inicie a Aplicação:**
   ```bash
   python planilhas.py
   ```

---

## 🖱️ Módulo de Uso

1. Ao abrir a janela da aplicação, clique em **"Anexar Arquivos"** e selecione um ou mais relatórios textuais (`.txt`).
2. Clique em **"Anexar Planilha Dinâmica"** e aponte para um arquivo base de Excel (que possua preferencialmente uma aba chamada `2020 DADOS`).
3. Com tudo selecionado, clique no grande botão **"Gerar Planilha Dinâmica"**.
4. O robô irá processar tudo autonomamente e avisará quando o fluxo terminar. Os resultados estarão em sua planilha selecionada e dentro do diretório gerador de planilhas.

---

*Projeto voltado à inteligência e automação de operações de rotina contábil, desenvolvido para otimização de tempo e redução substancial de erros operacionais.*
