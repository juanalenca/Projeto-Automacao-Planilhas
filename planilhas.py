"""


import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os
import pandas as pd
import sqlite3
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Definir variáveis globais
caminho_pasta = 'C:\\AutomacaoTabelasReciprev\\'
diretorio_planilhas = os.path.join(caminho_pasta, 'planilhas_geradas')
caminho_banco_dados = os.path.join(caminho_pasta, 'verbas2.db')
nome_aba_planilha_dinamica = '2020 DADOS'
dynamic_spreadsheet_path = ''

def processar_arquivo(caminho_arquivo, diretorio_planilhas, dynamic_spreadsheet_path):
    nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]

    # Carregar a planilha dinâmica
    workbook = load_workbook(dynamic_spreadsheet_path)
    worksheet = workbook[nome_aba_planilha_dinamica]

    with open(caminho_arquivo, 'r', encoding='utf-8') as file:
        conteudo = file.read()

    padrao_vantagens = r'(COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nCOD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR|\Z)'
    padrao_descontos = r'(COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nTOTAL\s+LIQUIDO|\Z)'

    def salvar_parte(texto, parte):
        caminho = os.path.join(diretorio_planilhas, f'{nome_arquivo}_parte{parte}.txt')
        with open(caminho, 'w', encoding='utf-8') as f:
            f.write(texto.strip())

    def dividir_tabela(texto):
        linhas = texto.strip().split('\n')
        metade_index = len(linhas[0]) // 2
        cabecalho1, cabecalho2 = linhas[0][:metade_index].strip(), linhas[0][metade_index:].strip()
        linha_separadora = '=' * len(cabecalho1)
        
        dados1, dados2 = [cabecalho1, linha_separadora], [cabecalho2, linha_separadora]
        for linha in linhas[2:]:
            if len(linha) > metade_index:
                parte1, parte2 = linha[:metade_index].strip(), linha[metade_index:].strip()
            else:
                parte1, parte2 = linha.strip(), ""
            dados1.append(parte1)
            dados2.append(parte2)
        
        tabela1 = '\n'.join(dados1)
        tabela2 = '\n'.join(dados2)
        return tabela1, tabela2

    def limpar_cod(cod):
        return re.sub(r'\D', '', cod).strip()

    def validar_numero(valor):
        try:
            valor = valor.replace('.', '').replace(',', '.')
            return float(valor)
        except ValueError:
            return 0.0

    def extrair_dados(texto, tipo, categoria, orgao):
        linhas = texto.strip().split('\n')[2:]
        dados = []
        for linha in linhas:
            campos = re.split(r'\s{2,}', linha)
            if len(campos) >= 4:
                cod = limpar_cod(campos[0])
                descricao = campos[1].strip()
                tot_func = validar_numero(campos[2])
                tot_valor = validar_numero(campos[3])
                dados.append((tipo, cod, descricao, tot_func, tot_valor, categoria, orgao))
        return dados

    def salvar_planilha(dados, nome_planilha):
        df = pd.DataFrame(dados, columns=['TIPO', 'COD', 'DESCRIÇÃO', 'TOT. FUNC.', 'TOT. VALOR', 'CATEGORIA', 'ORGÃO'])
        caminho_planilha = os.path.join(diretorio_planilhas, f'{nome_planilha}.xlsx')
        df.to_excel(caminho_planilha, index=False)

    def inserir_dados_planilha_dinamica(dados, worksheet):
        fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # Amarelo
        
        # Encontrar a última linha com dados
        ultima_linha = worksheet.max_row
        while not any(worksheet[ultima_linha][col].value for col in range(worksheet.max_column)):
            ultima_linha -= 1

        # Inserir dados logo abaixo da última linha com dados
        for dado in dados:
            worksheet.append(dado)
            ultima_linha += 1
            for cell in worksheet[ultima_linha]:
                cell.fill = fill


    def atualizar_banco_de_dados(dados):
        try:
            conexao = sqlite3.connect(caminho_banco_dados)
            cursor = conexao.cursor()
            cursor.executemany('INSERT INTO tabela (TIPO, COD, DESCRICAO, TOT_FUNC, TOT_VALOR, CATEGORIA, ORGAO) VALUES (?, ?, ?, ?, ?, ?, ?)', dados)
            conexao.commit()
        except sqlite3.Error as e:
            print(f"Erro ao inserir dados no banco de dados: {e}")
        finally:
            conexao.close()

    def extrair_categoria(conteudo):
        padrao_categoria = r'\b(APO|PEN|PPR)\b'
        match = re.search(padrao_categoria, conteudo)
        return match.group(0) if match else None

    def extrair_orgao(conteudo):
        padroes_orgao = [
            r'\bINATIVOS E PENSIONISTAS CAMARA\b',
            r'\bINATIVOS E PENSIONISTAS FUNDACAO CULTURA\b',
            r'\bINATIVOS E PENSIONISTAS GERALDAO\b',
            r'\bINATIVOS E PENSIONISTAS SETOR EDUCACIONA\b',
            r'\bINATIVOS E PENSIONS. SIST PREVIDENCIARIO\b',
            r'\bINATIVOS E PENSIONISTAS IASC\b'
        ]
        for padrao in padroes_orgao:
            match = re.search(padrao, conteudo)
            if match:
                return match.group(0)
        return None

    categoria = extrair_categoria(conteudo)
    orgao = extrair_orgao(conteudo)

    vantagens = re.findall(padrao_vantagens, conteudo)
    for i, vantagem in enumerate(vantagens):
        parte1, parte2 = dividir_tabela(vantagem)
        salvar_parte(parte1, f'vantagens_{i + 1}_parte1')
        salvar_parte(parte2, f'vantagens_{i + 1}_parte2')
        
        dados_parte1 = extrair_dados(parte1, 'P', categoria, orgao)
        dados_parte2 = extrair_dados(parte2, 'P', categoria, orgao)
        
        salvar_planilha(dados_parte1, f'vantagens_{i + 1}_parte1')
        salvar_planilha(dados_parte2, f'vantagens_{i + 1}_parte2')
        
        inserir_dados_planilha_dinamica(dados_parte1, worksheet)
        inserir_dados_planilha_dinamica(dados_parte2, worksheet)
        
        atualizar_banco_de_dados(dados_parte1)
        atualizar_banco_de_dados(dados_parte2)

    descontos = re.findall(padrao_descontos, conteudo)
    for i, desconto in enumerate(descontos):
        parte1, parte2 = dividir_tabela(desconto)
        salvar_parte(parte1, f'descontos_{i + 1}_parte1')
        salvar_parte(parte2, f'descontos_{i + 1}_parte2')
        
        dados_parte1 = extrair_dados(parte1, 'D', categoria, orgao)
        dados_parte2 = extrair_dados(parte2, 'D', categoria, orgao)
        
        salvar_planilha(dados_parte1, f'descontos_{i + 1}_parte1')
        salvar_planilha(dados_parte2, f'descontos_{i + 1}_parte2')
        
        inserir_dados_planilha_dinamica(dados_parte1, worksheet)
        inserir_dados_planilha_dinamica(dados_parte2, worksheet)
        
        atualizar_banco_de_dados(dados_parte1)
        atualizar_banco_de_dados(dados_parte2)

    workbook.save(dynamic_spreadsheet_path)

def anexar_arquivos():
    arquivos = filedialog.askopenfilenames(filetypes=[("Text Files", "*.txt")])
    for arquivo in arquivos:
        lista_arquivos.insert(tk.END, arquivo)

def anexar_planilha_dinamica():
    global dynamic_spreadsheet_path
    dynamic_spreadsheet_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if dynamic_spreadsheet_path:
        label_planilha_dinamica.config(text=os.path.basename(dynamic_spreadsheet_path))

def gerar_planilha_dinamica():
    if not dynamic_spreadsheet_path:
        messagebox.showwarning("Aviso", "Por favor, anexe uma planilha dinâmica.")
        return

    if lista_arquivos.size() == 0:
        messagebox.showwarning("Aviso", "Por favor, anexe arquivos de texto para processar.")
        return

    for idx in range(lista_arquivos.size()):
        caminho_arquivo = lista_arquivos.get(idx)
        processar_arquivo(caminho_arquivo, diretorio_planilhas, dynamic_spreadsheet_path)

    messagebox.showinfo("Sucesso", "A planilha dinâmica foi atualizada com sucesso.")
    lista_arquivos.delete(0, tk.END)
    label_planilha_dinamica.config(text="Nenhum arquivo anexado")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Processador de Arquivos")

frame_arquivos = tk.Frame(root)
frame_arquivos.pack(pady=10)

label_arquivos = tk.Label(frame_arquivos, text="Arquivos de texto:")
label_arquivos.pack(side=tk.LEFT)

btn_anexar_arquivos = tk.Button(frame_arquivos, text="Anexar Arquivos", command=anexar_arquivos)
btn_anexar_arquivos.pack(side=tk.LEFT, padx=5)

lista_arquivos = tk.Listbox(root, width=80, height=10)
lista_arquivos.pack(pady=10)

frame_planilha_dinamica = tk.Frame(root)
frame_planilha_dinamica.pack(pady=10)

label_planilha_dinamica = tk.Label(frame_planilha_dinamica, text="Nenhum arquivo anexado")
label_planilha_dinamica.pack(side=tk.LEFT)

btn_anexar_planilha_dinamica = tk.Button(frame_planilha_dinamica, text="Anexar Planilha Dinâmica", command=anexar_planilha_dinamica)
btn_anexar_planilha_dinamica.pack(side=tk.LEFT, padx=5)

btn_gerar_planilha_dinamica = tk.Button(root, text="Gerar Planilha Dinâmica", command=gerar_planilha_dinamica)
btn_gerar_planilha_dinamica.pack(pady=20)

root.mainloop()

"""





import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os
import pandas as pd
import sqlite3
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Definir variáveis globais
caminho_pasta = 'C:\\AutomacaoTabelasReciprev\\'
diretorio_planilhas = os.path.join(caminho_pasta, 'planilhas_geradas')
caminho_banco_dados = os.path.join(caminho_pasta, 'verbas2.db')
nome_aba_planilha_dinamica = '2020 DADOS'
dynamic_spreadsheet_path = ''
caminho_empresa_recifin = 'C:\\AutomacaoTabelasReciprev\\tabelas\\empresa 23 recifin.xlsx'  # Atualizar com o caminho correto

def processar_arquivo(caminho_arquivo, diretorio_planilhas, dynamic_spreadsheet_path):
    nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]

    try:
        # Carregar a planilha dinâmica
        workbook = load_workbook(dynamic_spreadsheet_path)
        worksheet = workbook[nome_aba_planilha_dinamica]
    except Exception as e:
        print(f"Erro ao carregar a planilha dinâmica: {e}")
        return

    with open(caminho_arquivo, 'r', encoding='utf-8') as file:
        conteudo = file.read()

    padrao_vantagens = r'(COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nCOD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR|\Z)'
    padrao_descontos = r'(COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nTOTAL\s+LIQUIDO|\Z)'

    def salvar_parte(texto, parte):
        caminho = os.path.join(diretorio_planilhas, f'{nome_arquivo}_parte{parte}.txt')
        with open(caminho, 'w', encoding='utf-8') as f:
            f.write(texto.strip())

    def dividir_tabela(texto):
        linhas = texto.strip().split('\n')
        metade_index = len(linhas[0]) // 2
        cabecalho1, cabecalho2 = linhas[0][:metade_index].strip(), linhas[0][metade_index:].strip()
        linha_separadora = '=' * len(cabecalho1)
        
        dados1, dados2 = [cabecalho1, linha_separadora], [cabecalho2, linha_separadora]
        for linha in linhas[2:]:
            if len(linha) > metade_index:
                parte1, parte2 = linha[:metade_index].strip(), linha[metade_index:].strip()
            else:
                parte1, parte2 = linha.strip(), ""
            dados1.append(parte1)
            dados2.append(parte2)
        
        tabela1 = '\n'.join(dados1)
        tabela2 = '\n'.join(dados2)
        return tabela1, tabela2

    def limpar_cod(cod):
        return re.sub(r'\D', '', cod).strip()

    def validar_numero(valor):
        try:
            valor = valor.replace('.', '').replace(',', '.')
            return float(valor)
        except ValueError:
            return 0.0

    def extrair_dados(texto, tipo, categoria, orgao):
        linhas = texto.strip().split('\n')[2:]
        dados = []
        for linha in linhas:
            campos = re.split(r'\s{2,}', linha)
            if len(campos) >= 4:
                cod = limpar_cod(campos[0])
                descricao = campos[1].strip()
                tot_func = validar_numero(campos[2])
                tot_valor = validar_numero(campos[3])
                dados.append((tipo, cod, descricao, tot_func, tot_valor, categoria, orgao))
        return dados

    def salvar_planilha(dados, nome_planilha):
        df = pd.DataFrame(dados, columns=['TIPO', 'COD', 'DESCRIÇÃO', 'TOT. FUNC.', 'TOT. VALOR', 'CATEGORIA', 'ORGÃO'])
        caminho_planilha = os.path.join(diretorio_planilhas, f'{nome_planilha}.xlsx')
        df.to_excel(caminho_planilha, index=False)

    def inserir_dados_planilha_dinamica(dados, worksheet):
        fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # Amarelo
        
        # Encontrar a última linha com dados
        ultima_linha = worksheet.max_row
        while not any(worksheet[ultima_linha][col].value for col in range(worksheet.max_column)):
            ultima_linha -= 1

        # Inserir dados logo abaixo da última linha com dados
        for dado in dados:
            worksheet.append(dado)
            ultima_linha += 1
            for cell in worksheet[ultima_linha]:
                cell.fill = fill

    def atualizar_banco_de_dados(dados):
        try:
            conexao = sqlite3.connect(caminho_banco_dados)
            cursor = conexao.cursor()
            cursor.executemany('INSERT INTO tabela (TIPO, COD, DESCRICAO, TOT_FUNC, TOT_VALOR, CATEGORIA, ORGAO) VALUES (?, ?, ?, ?, ?, ?, ?)', dados)
            conexao.commit()
        except sqlite3.Error as e:
            print(f"Erro ao inserir dados no banco de dados: {e}")
        finally:
            conexao.close()

    def extrair_categoria(conteudo):
        padrao_categoria = r'\b(APO|PEN|PPR)\b'
        match = re.search(padrao_categoria, conteudo)
        return match.group(0) if match else None

    def extrair_orgao(conteudo):
        padroes_orgao = [
            r'\bINATIVOS E PENSIONISTAS CAMARA\b',
            r'\bINATIVOS E PENSIONISTAS FUNDACAO CULTURA\b',
            r'\bINATIVOS E PENSIONISTAS GERALDAO\b',
            r'\bINATIVOS E PENSIONISTAS SETOR EDUCACIONA\b',
            r'\bINATIVOS E PENSIONS. SIST PREVIDENCIARIO\b',
            r'\bINATIVOS E PENSIONISTAS IASC\b'
        ]
        for padrao in padroes_orgao:
            match = re.search(padrao, conteudo)
            if match:
                return match.group(0)
        return None

    categoria = extrair_categoria(conteudo)
    orgao = extrair_orgao(conteudo)

    vantagens = re.findall(padrao_vantagens, conteudo)
    for i, vantagem in enumerate(vantagens):
        parte1, parte2 = dividir_tabela(vantagem)
        salvar_parte(parte1, f'vantagens_{i + 1}_parte1')
        salvar_parte(parte2, f'vantagens_{i + 1}_parte2')
        
        dados_parte1 = extrair_dados(parte1, 'P', categoria, orgao)
        dados_parte2 = extrair_dados(parte2, 'P', categoria, orgao)
        
        salvar_planilha(dados_parte1, f'vantagens_{i + 1}_parte1')
        salvar_planilha(dados_parte2, f'vantagens_{i + 1}_parte2')
        
        inserir_dados_planilha_dinamica(dados_parte1, worksheet)
        inserir_dados_planilha_dinamica(dados_parte2, worksheet)
        
        atualizar_banco_de_dados(dados_parte1)
        atualizar_banco_de_dados(dados_parte2)

    descontos = re.findall(padrao_descontos, conteudo)
    for i, desconto in enumerate(descontos):
        parte1, parte2 = dividir_tabela(desconto)
        salvar_parte(parte1, f'descontos_{i + 1}_parte1')
        salvar_parte(parte2, f'descontos_{i + 1}_parte2')
        
        dados_parte1 = extrair_dados(parte1, 'D', categoria, orgao)
        dados_parte2 = extrair_dados(parte2, 'D', categoria, orgao)
        
        salvar_planilha(dados_parte1, f'descontos_{i + 1}_parte1')
        salvar_planilha(dados_parte2, f'descontos_{i + 1}_parte2')
        
        inserir_dados_planilha_dinamica(dados_parte1, worksheet)
        inserir_dados_planilha_dinamica(dados_parte2, worksheet)
        
        atualizar_banco_de_dados(dados_parte1)
        atualizar_banco_de_dados(dados_parte2)

    nova_aba = workbook.copy_worksheet(worksheet)
    nova_aba.title = nome_arquivo
    workbook.save(dynamic_spreadsheet_path)
    print(f'Arquivo {nome_arquivo} processado com sucesso!')

def inserir_classificacoes(dados_classificacoes):
    try:
        conexao = sqlite3.connect(caminho_banco_dados)
        cursor = conexao.cursor()
        cursor.executemany('INSERT INTO classificacoes (codigo, nome) VALUES (?, ?)', dados_classificacoes)
        conexao.commit()
    except sqlite3.Error as e:
        print(f"Erro ao inserir classificações: {e}")
    finally:
        conexao.close()

# Carregar os dados do arquivo Excel "empresa 23 recifin.xlsx"
try:
    dados_empresa_recifin = pd.read_excel(caminho_empresa_recifin, sheet_name=0)
    dados_empresa_recifin_tuplas = [tuple(row) for row in dados_empresa_recifin.values]
    inserir_classificacoes(dados_empresa_recifin_tuplas)
except Exception as e:
    print(f"Erro ao carregar ou processar o arquivo empresa 23 recifin.xlsx: {e}")

# Código adicional para abrir o diretório e iniciar o processamento
root = tk.Tk()
root.withdraw()

arquivo_selecionado = filedialog.askopenfilename(title='Selecione o arquivo a ser processado')
if arquivo_selecionado:
    processar_arquivo(arquivo_selecionado, diretorio_planilhas, dynamic_spreadsheet_path)
    messagebox.showinfo('Sucesso', f'Arquivo {os.path.basename(arquivo_selecionado)} processado com sucesso!')
else:
    messagebox.showwarning('Atenção', 'Nenhum arquivo foi selecionado!')









import sqlite3
import pandas as pd

# Conectar ao banco de dados
conn = sqlite3.connect('verbas2.db')
cursor = conn.cursor()

# Criação da nova tabela para armazenar as informações detalhadas das verbas
cursor.execute('''
CREATE TABLE IF NOT EXISTS verbas_detalhadas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo TEXT,
    verba TEXT,
    descricao TEXT,
    quantidade REAL,
    valor REAL,
    categoria TEXT,
    orgao TEXT
)
''')

# Criação da tabela para armazenar os códigos de classificação
cursor.execute('''
CREATE TABLE IF NOT EXISTS classificacoes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo_classificacao TEXT UNIQUE,
    descricao TEXT
)
''')

# Função para popular a tabela classificacoes com dados do arquivo Recifin
def inserir_classificacoes(dados):
    for index, row in dados.iterrows():
        try:
            cursor.execute('''
            INSERT OR IGNORE INTO classificacoes (codigo_classificacao, descricao)
            VALUES (?, ?)
            ''', (row['CLASSIFICAÇÃO'], row['DESCR. VERBA']))
        except sqlite3.Error as e:
            print(f"Erro ao inserir classificação: {e}")

# Carregar os dados do arquivo Excel (Recifin)
caminho_recifin = 'C:\\AutomacaoTabelasReciprev\\tabelas\\empresa 23 recifin.xlsx'  # Atualize o caminho conforme necessário
dados_empresa_recifin = pd.read_excel(caminho_recifin)

# Inserir classificações no banco de dados
inserir_classificacoes(dados_empresa_recifin)

# Commit das operações
conn.commit()

# Fechar a conexão
conn.close()
