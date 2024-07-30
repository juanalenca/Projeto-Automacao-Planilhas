import re
import pandas as pd

def processar_arquivo(caminho_arquivo):
    with open(caminho_arquivo, 'r') as file:
        conteudo = file.read()
    
    # Regex para identificar e separar as tabelas de vantagens e descontos
    padrao_vantagens = r'(COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+V A N T A G E N S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nCOD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR|\Z)'
    padrao_descontos = r'(COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR\s+COD\s+D E S C O N T O S\s+TOT\. FUNC\.\s+TOT\. VALOR[\s\S]+?)(?=\nTOTAL\s+LIQUIDO|\Z)'

    def salvar_parte(texto, prefixo, numero, parte):
        with open(f'{prefixo}_{numero}_parte{parte}.txt', 'w') as f:
            f.write(texto.strip())

    def dividir_tabela(texto):
        linhas = texto.strip().split('\n')
        metade_index = len(linhas[0]) // 2
        cabecalho1, cabecalho2 = linhas[0][:metade_index].strip(), linhas[0][metade_index:].strip()
        linha_separadora = '=' * len(cabecalho1)
        
        dados1, dados2 = [cabecalho1, linha_separadora], [cabecalho2, linha_separadora]
        for linha in linhas[2:]:
            metade_index = len(linha) // 2
            parte1, parte2 = linha[:metade_index].strip(), linha[metade_index:].strip()
            dados1.append(parte1)
            dados2.append(parte2)
        
        tabela1 = '\n'.join(dados1)
        tabela2 = '\n'.join(dados2)
        return tabela1, tabela2

    def extrair_dados(tabela):
        linhas = tabela.split('\n')
        dados = []
        for linha in linhas[1:]:
            partes = re.split(r'\s{2,}', linha.strip())
            if len(partes) == 4:
                dados.append(partes)
        return dados

    def salvar_planilha(dados, prefixo):
        df = pd.DataFrame(dados, columns=['COD', 'DESCRIÇÃO', 'TOT. FUNC.', 'TOT. VALOR'])
        df.to_excel(f'{prefixo}.xlsx', index=False)

    # Encontrar e separar vantagens
    vantagens = re.findall(padrao_vantagens, conteudo)
    for i, vantagem in enumerate(vantagens):
        parte1, parte2 = dividir_tabela(vantagem)
        salvar_parte(parte1, 'vantagens', i + 1, 1)
        salvar_parte(parte2, 'vantagens', i + 1, 2)
        
        dados_parte1 = extrair_dados(parte1)
        dados_parte2 = extrair_dados(parte2)
        
        salvar_planilha(dados_parte1, f'vantagens_{i + 1}_parte1')
        salvar_planilha(dados_parte2, f'vantagens_{i + 1}_parte2')

    # Encontrar e separar descontos
    descontos = re.findall(padrao_descontos, conteudo)
    for i, desconto in enumerate(descontos):
        parte1, parte2 = dividir_tabela(desconto)
        salvar_parte(parte1, 'descontos', i + 1, 1)
        salvar_parte(parte2, 'descontos', i + 1, 2)
        
        dados_parte1 = extrair_dados(parte1)
        dados_parte2 = extrair_dados(parte2)
        
        salvar_planilha(dados_parte1, f'descontos_{i + 1}_parte1')
        salvar_planilha(dados_parte2, f'descontos_{i + 1}_parte2')

# Caminho para o arquivo de entrada
caminho_arquivo = 'C:\\AutomacaoTabelasReciprev\\824525.txt'

# Processar o arquivo
processar_arquivo(caminho_arquivo)

print("O processamento foi concluído. Verifique as planilhas geradas.")
