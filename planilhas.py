import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os
import sqlite3
import pandas as pd
from typing import List, Tuple, Optional, Any
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# Configurações de Diretórios (Dinâmicos)
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIRETORIO_PLANILHAS = os.path.join(BASE_DIR, 'arquivos-analise', 'planilhas_geradas')
CAMINHO_BANCO_DADOS = os.path.join(BASE_DIR, 'contab_reciprev.db')

# Garantir que o diretório de destino existe
os.makedirs(DIRETORIO_PLANILHAS, exist_ok=True)

NOME_ABA_PLANILHA_DINAMICA = '2020 DADOS'
NOME_ABA_CALCULO_FINAL = 'CÁLCULO FINAL FOLHA'

# Estado global da UI
dynamic_spreadsheet_path = ''


# ==========================================
# Funções Auxiliares de Processamento
# ==========================================

def salvar_parte(texto: str, nome_arquivo: str, parte: str) -> None:
    """Salva uma parte do texto em um arquivo .txt para análise ou registro."""
    caminho = os.path.join(DIRETORIO_PLANILHAS, f'{nome_arquivo}_parte{parte}.txt')
    with open(caminho, 'w', encoding='utf-8') as f:
        f.write(texto.strip())


def dividir_tabela(texto: str) -> Tuple[str, str]:
    """Divide a tabela em duas partes (colunas)."""
    linhas = texto.strip().split('\n')
    if not linhas:
        return "", ""
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

    return '\n'.join(dados1), '\n'.join(dados2)


def limpar_cod(cod: str) -> str:
    """Remove caracteres nâo numéricos do código."""
    return re.sub(r'\D', '', cod).strip()


def validar_numero(valor: str) -> float:
    """Converte string numérica formatada em número float."""
    try:
        if isinstance(valor, str):
            valor = valor.replace('.', '').replace(',', '.')
        return float(valor)
    except ValueError:
        return 0.0


def obter_classificacao(cursor: sqlite3.Cursor, verba: str) -> Optional[str]:
    """Busca a classificação de uma verba a partir do banco de dados SQLite."""
    cursor.execute("SELECT id FROM verbas WHERE verba = ?", (verba,))
    verba_id = cursor.fetchone()

    if verba_id:
        cursor.execute("SELECT classificacao FROM classificacoes WHERE verba_id = ?", (verba_id[0],))
        resultado = cursor.fetchone()
        return resultado[0] if resultado else None
    return None


def extrair_dados(texto: str, tipo: str, categoria: str, orgao: str) -> List[Tuple[Any, ...]]:
    """Extrai os dados das tabelas parciais e retorna listagem de tuplas populadas."""
    linhas = texto.strip().split('\n')[2:]  # Ignorar cabeçalhos
    dados = []

    with sqlite3.connect(CAMINHO_BANCO_DADOS) as conexao:
        cursor = conexao.cursor()
        for linha in linhas:
            if not linha.strip() or len(linha.strip()) < 20:
                continue

            campos = re.split(r'\s{2,}', linha.strip())
            if len(campos) >= 4:
                cod = limpar_cod(campos[0])
                descricao = campos[1].strip()
                tot_func = validar_numero(campos[2])
                tot_valor = validar_numero(campos[3])
                classificacao = obter_classificacao(cursor, cod)

                dado = (tipo, cod, descricao, tot_func, tot_valor, classificacao, categoria, orgao)
                # Adiciona somente se as 8 colunas estiverem presentes e estruturadas
                if len(dado) == 8:
                    dados.append(dado)
                else:
                    print(f"Aviso: Tupla com estrutura incorreta ignorada: {dado}")
            else:
                print(f"Aviso: Linha ignorada por falta de campos exigidos: {linha}")
    
    return dados


def salvar_planilha(dados: List[Tuple], nome_planilha: str) -> None:
    """Exporta os dados extraídos em formato XLSX para debug."""
    if not dados:
        return
    df = pd.DataFrame(dados, columns=['TIPO', 'COD', 'DESCRIÇÃO', 'TOT. FUNC.', 'TOT. VALOR', 'CLASSIFICAÇÃO', 'CATEGORIA', 'ORGÃO'])
    caminho_planilha = os.path.join(DIRETORIO_PLANILHAS, f'{nome_planilha}.xlsx')
    df.to_excel(caminho_planilha, index=False)


def inserir_dados_planilha_dinamica(dados: List[Tuple], worksheet) -> None:
    """Insere novos dados no final da planilha anexada, destacando em amarelo."""
    fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    ultima_linha = worksheet.max_row
    # Certificar-se que a linha está realmente vazia
    while ultima_linha > 0 and not any(cell.value for cell in worksheet[ultima_linha]):
        ultima_linha -= 1

    for dado in dados:
        worksheet.append(dado)
        ultima_linha += 1
        for cell in worksheet[ultima_linha]:
            cell.fill = fill


def atualizar_banco_de_dados(dados: List[Tuple]) -> None:
    """Insere em cascata novos registros não duplicados no banco de dados local SQLite."""
    try:
        with sqlite3.connect(CAMINHO_BANCO_DADOS) as conexao:
            cursor = conexao.cursor()

            for dado in dados:
                tipo, cod, descricao, tot_func, tot_valor, classificacao, categoria, orgao = dado

                cursor.execute('SELECT id FROM verbas WHERE verba = ?', (cod,))
                resultado = cursor.fetchone()

                if resultado:
                    print(f"Verba já mapeada no DB: {cod} - Pulando.")
                    continue

                cursor.execute('''
                    INSERT INTO verbas (verba, tipo, descricao, quantidade, total_valor)
                    VALUES (?, ?, ?, ?, ?)
                ''', (cod, tipo, descricao, tot_func, tot_valor))
                verba_id = cursor.lastrowid

                if classificacao:
                    cursor.execute('INSERT INTO classificacoes (verba_id, classificacao) VALUES (?, ?)', (verba_id, classificacao))
                if orgao:
                    cursor.execute('INSERT INTO orgaos (verba_id, orgao) VALUES (?, ?)', (verba_id, orgao))
                if categoria:
                    cursor.execute('INSERT INTO categorias (verba_id, categoria) VALUES (?, ?)', (verba_id, categoria))
                if cod:
                    cursor.execute('INSERT INTO codigo_orgaos (verba_id, codigo_orgao) VALUES (?, ?)', (verba_id, cod))

            conexao.commit()
    except sqlite3.Error as e:
        print(f"Erro SQLite: {e}")


def extrair_categoria(conteudo: str) -> str:
    """Busca categorias específicas como APO, PEN ou PPR na página textual."""
    match = re.search(r'\b(APO|PEN|PPR)\b', conteudo)
    return match.group(0) if match else "CATEGORIA NÃO ENCONTRADA"


def extrair_orgao(conteudo: str) -> str:
    """Mapeia o órgão de destino na linha orgânica textual."""
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
    return "ÓRGÃO NÃO ENCONTRADO"


def inserir_dados_calculo_final(dados: List[Tuple], worksheet_calculo_final) -> None:
    """Constrói a montagem visual dos cálculos agrupados e totalizados no Excel."""
    dados_agrupados = {}
    for dado in dados:
        if len(dado) >= 8:
            key = (dado[0], dado[6], dado[7])  # -> TIPO, CATEGORIA, ORGÃO
            dados_agrupados.setdefault(key, []).append(dado)

    # Estilos Visuais
    bold_font = Font(bold=True)
    alignment_center = Alignment(horizontal='center')
    alignment_left = Alignment(horizontal='left')
    alignment_right = Alignment(horizontal='right')

    linha_atual = getattr(inserir_dados_calculo_final, 'linha_atual', 1)

    for (tipo, categoria, orgao), items in dados_agrupados.items():
        worksheet_calculo_final[f'A{linha_atual}'] = f'{tipo} - {orgao}'
        worksheet_calculo_final[f'A{linha_atual}'].font = bold_font
        worksheet_calculo_final[f'A{linha_atual}'].alignment = alignment_left
        worksheet_calculo_final.merge_cells(f'A{linha_atual}:E{linha_atual}')
        linha_atual += 1

        classificacao_grupos = {}
        for item in items:
            classificacao_grupos.setdefault(item[5], []).append(item)

        for classificacao, itens_cl in classificacao_grupos.items():
            total_valor = sum(item[4] for item in itens_cl)
            
            worksheet_calculo_final[f'B{linha_atual}'] = classificacao
            worksheet_calculo_final[f'C{linha_atual}'] = f'R$ {total_valor:,.2f}'
            worksheet_calculo_final[f'B{linha_atual}'].font = bold_font
            worksheet_calculo_final[f'C{linha_atual}'].font = bold_font
            worksheet_calculo_final[f'B{linha_atual}'].alignment = alignment_left
            worksheet_calculo_final[f'C{linha_atual}'].alignment = alignment_left
            linha_atual += 1

            for item in itens_cl:
                cod, descricao, tot_valor = item[1], item[2], item[4]
                worksheet_calculo_final[f'C{linha_atual}'] = cod
                worksheet_calculo_final[f'D{linha_atual}'] = descricao
                worksheet_calculo_final[f'E{linha_atual}'] = f'R$ {tot_valor:,.2f}'
                worksheet_calculo_final[f'C{linha_atual}'].alignment = alignment_center
                worksheet_calculo_final[f'D{linha_atual}'].alignment = alignment_left
                worksheet_calculo_final[f'E{linha_atual}'].alignment = alignment_right
                linha_atual += 1
            linha_atual += 1

    for col in range(1, 6):
        worksheet_calculo_final.column_dimensions[get_column_letter(col)].width = 20
    
    # Armazena estado local na função ao invés de buscar a planilha toda vez
    inserir_dados_calculo_final.linha_atual = linha_atual


def processar_secao(secoes: List[str], nome_arquivo: str, tipo: str, categoria: str, orgao: str, 
                    worksheet, worksheet_calculo_final, prefixo: str) -> None:
    """Função utilitária (DRY) para processar blocos idênticos de Vantagens e Descontos."""
    for i, secao in enumerate(secoes):
        p1, p2 = dividir_tabela(secao)
        
        # Salva o texto isolado
        nome_p1, nome_p2 = f'{prefixo}_{i+1}_parte1', f'{prefixo}_{i+1}_parte2'
        salvar_parte(p1, nome_arquivo, nome_p1)
        salvar_parte(p2, nome_arquivo, nome_p2)

        # Extrai os dados em tuplas tipificadas
        dados_p1 = extrair_dados(p1, tipo, categoria, orgao)
        dados_p2 = extrair_dados(p2, tipo, categoria, orgao)

        # Trata e emite XLSX das frações
        salvar_planilha(dados_p1, f'{nome_arquivo}_{nome_p1}')
        salvar_planilha(dados_p2, f'{nome_arquivo}_{nome_p2}')

        # Atualizações dinâmicas na memória do Excel Base
        if dados_p1: 
            inserir_dados_planilha_dinamica(dados_p1, worksheet)
            inserir_dados_calculo_final(dados_p1, worksheet_calculo_final)
            atualizar_banco_de_dados(dados_p1)
        if dados_p2:
            inserir_dados_planilha_dinamica(dados_p2, worksheet)
            inserir_dados_calculo_final(dados_p2, worksheet_calculo_final)
            atualizar_banco_de_dados(dados_p2)


# ==========================================
# Núcleo de Processamento Principal
# ==========================================
def processar_arquivo(caminho_arquivo: str, dynamic_path: str) -> None:
    """Motor central para despachar relatórios .txt para o workbook Excel dinâmico."""
    nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]

    workbook = load_workbook(dynamic_path)
    if NOME_ABA_PLANILHA_DINAMICA not in workbook.sheetnames:
        workbook.create_sheet(NOME_ABA_PLANILHA_DINAMICA)
    worksheet = workbook[NOME_ABA_PLANILHA_DINAMICA]

    if NOME_ABA_CALCULO_FINAL in workbook.sheetnames:
        worksheet_calculo_final = workbook[NOME_ABA_CALCULO_FINAL]
    else:
        worksheet_calculo_final = workbook.create_sheet(NOME_ABA_CALCULO_FINAL)
        # Reseta o contador estrutural de agrupamento sempre que for arquivo novo totalitário
        if hasattr(inserir_dados_calculo_final, 'linha_atual'):
            del inserir_dados_calculo_final.linha_atual

    with open(caminho_arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()

    categoria = extrair_categoria(conteudo)
    orgao = extrair_orgao(conteudo)

    # Identificadores Regex para partições do documento
    padrao_v = r'(COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nCOD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR|\Z)'
    padrao_d = r'(COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nTOTAL\s+LIQUIDO|\Z)'

    vantagens = re.findall(padrao_v, conteudo)
    descontos = re.findall(padrao_d, conteudo)

    processar_secao(vantagens, nome_arquivo, 'P', categoria, orgao, worksheet, worksheet_calculo_final, 'vantagens')
    processar_secao(descontos, nome_arquivo, 'D', categoria, orgao, worksheet, worksheet_calculo_final, 'descontos')

    workbook.save(dynamic_path)


# ==========================================
# Interface Gráfica e Lógica (Tkinter)
# ==========================================
def anexar_arquivos() -> None:
    arquivos = filedialog.askopenfilenames(filetypes=[("Text Files", "*.txt")])
    for arquivo in arquivos:
        if arquivo not in lista_arquivos.get(0, tk.END):
            lista_arquivos.insert(tk.END, arquivo)

def anexar_planilha_dinamica() -> None:
    global dynamic_spreadsheet_path
    caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if caminho:
        dynamic_spreadsheet_path = caminho
        label_planilha_dinamica.config(text=os.path.basename(dynamic_spreadsheet_path))

def gerar_planilha_dinamica() -> None:
    if not dynamic_spreadsheet_path:
        messagebox.showwarning("Aviso", "Por favor, anexe uma planilha base do Excel (*.xlsx).")
        return

    if lista_arquivos.size() == 0:
        messagebox.showwarning("Aviso", "Por favor, anexe um ou mais arquivos de relatório (.txt).")
        return

    try:
        for idx in range(lista_arquivos.size()):
            caminho_arquivo = lista_arquivos.get(idx)
            processar_arquivo(caminho_arquivo, dynamic_spreadsheet_path)
            
        messagebox.showinfo("Sucesso", "Todas as informações calculadas e injetadas! A Planilha dinâmica foi atualizada com sucesso.")
        lista_arquivos.delete(0, tk.END)
        label_planilha_dinamica.config(text="Nenhum arquivo de destino anexado")
        
        global dynamic_spreadsheet_path
        dynamic_spreadsheet_path = ''
    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Falha ao rodar processamento:\n\nERRO TÉCNICO: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Extração Inteligente Reciprev")
    root.geometry("600x350")
    root.resizable(False, False)

    frame_arquivos = tk.Frame(root)
    frame_arquivos.pack(pady=15)
    
    tk.Label(frame_arquivos, text="📄 Relatórios (.txt):", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
    tk.Button(frame_arquivos, text="Anexar Arquivos TXT", command=anexar_arquivos, bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)

    lista_arquivos = tk.Listbox(root, width=80, height=9)
    lista_arquivos.pack(pady=5)

    frame_planilha_dinamica = tk.Frame(root)
    frame_planilha_dinamica.pack(pady=10)
    
    tk.Label(frame_planilha_dinamica, text="📊 Planilha Base (.xlsx):", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
    tk.Button(frame_planilha_dinamica, text="Selecionar Excel Dinâmico", command=anexar_planilha_dinamica, bg="#2196F3", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
    
    label_planilha_dinamica = tk.Label(root, text="Nenhum arquivo de destino anexado", fg="gray")
    label_planilha_dinamica.pack(pady=1)

    tk.Button(root, text="Processar e Sobrescrever Planilha Base", command=gerar_planilha_dinamica, bg="#f44336", fg="white", font=("Arial", 11, "bold")).pack(pady=15)

    root.mainloop()
