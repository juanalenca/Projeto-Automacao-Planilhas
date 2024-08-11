import sqlite3
import pandas as pd

# Caminhos dos arquivos de planilhas
file_path1 = 'C:\\AutomacaoReciprev\\tabelas\\dados-cassificacao.xlsx'
file_path2 = 'C:\\AutomacaoReciprev\\tabelas\\dados-cassificacao2.xlsx'

# Ler as planilhas
data1 = pd.read_excel(file_path1)
data2 = pd.read_excel(file_path2, engine='odf')

# Combinar os dados das duas planilhas
combined_data = pd.concat([data1, data2])

# Remover linhas duplicadas
combined_data = combined_data.drop_duplicates()

# Selecionar apenas as colunas relevantes
combined_data = combined_data[['CÓD. VERBA', 'TIPO VERBA', 'DESCR. VERBA', 'CLASSIFICAÇÃO', 'CATEGORIA']]

# Renomear as colunas para facilitar o mapeamento
combined_data.columns = ['cod_verba', 'tipo_verba', 'desc_verba', 'classificacao', 'categoria']

# Conectar ao banco de dados e criar/atualizar a tabela
conn = sqlite3.connect('database.db')
c = conn.cursor()

# Criar nova tabela com colunas adicionais (se necessário)
c.execute('''
CREATE TABLE IF NOT EXISTS verbas (
    cod_verba INTEGER PRIMARY KEY,
    tipo_verba TEXT,
    desc_verba TEXT,
    classificacao TEXT,
    categoria TEXT
)
''')

# Inserir os dados no banco de dados
for index, row in combined_data.iterrows():
    c.execute('''
    INSERT OR IGNORE INTO verbas (cod_verba, tipo_verba, desc_verba, classificacao, categoria)
    VALUES (?, ?, ?, ?, ?)
    ''', (row['cod_verba'], row['tipo_verba'], row['desc_verba'], row['classificacao'], row['categoria']))

# Salvar as alterações
conn.commit()

# Fechar a conexão
conn.close()
