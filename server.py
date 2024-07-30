import sqlite3

# Criação de conexão com o banco de dados (cria o arquivo database.db)
conn = sqlite3.connect('database.db')
c = conn.cursor()

# Criação de uma tabela para armazenar as informações das verbas
c.execute('''
CREATE TABLE IF NOT EXISTS verbas (
    cod_verba INTEGER PRIMARY KEY,
    tipo_verba TEXT,
    desc_verba TEXT
)
''')

# Lista com os dados da planilha
data = [
    (513, 'D', 'INSS - CONTRIBUICAO'),
    (516, 'D', 'IMPOSTO DE RENDA'),
    (755, 'D', 'CO-PARTICIPACAO SAUDE REC'),
    (777, 'D', 'SAUDE RECIFE'),
    (6, 'P', 'GRAT ATIVIDADE PREVID'),
    (42, 'P', 'CARGO EM COMISSAO'),
    (366, 'P', 'VALE REFEICAO'),
    (652, 'D', 'EMPREST. BANCO BRASIL'),
    (40, 'P', 'GRAT OPERADOR FOLHA PAGTO'),
    (183, 'P', 'ADIANTAMENTO 1/3 FERIAS'),
    (530, 'D', 'CONTR RECIFIN'),
    (7, 'P', 'GRAT ASSIS SAUDE SERVIDOR'),
    (687, 'D', 'EMP BRADESCO'),
    (133, 'P', 'INSALUBRIDADE'),
    (532, 'D', 'VALE TRANSPORTE'),
    (457, 'P', 'ABONO PERM EMC41 PREVIDEN'),
    (441, 'D', 'DEV. GRATIFICACAO INDEVID'),
    (418, 'D', 'PERNAMBUCRED - INTEG.MENS'),
    (126, 'P', 'GRAT MEM EQ APOI/COM CONT'),
    (149, 'P', 'REUNIAO'),
    (182, 'P', 'FER 1/3 ART 65 L 15127/88'),
    (483, 'D', 'DES AD 1/3 FERIAS'),
    (678, 'D', 'EMPRESTIMO CAIXA ECONOMIC'),
    (651, 'D', 'EMRPESTIMO BANCO REAL'),
    (148, 'P', 'CURSOS E TREINAMENTOS'),
    (137, 'P', 'GRAT AG CONTRT/PRESID COM'),
    (39, 'P', 'CARG COM ART5 L.19208/24')
]

# Inserção dos dados na tabela
for item in data:
    c.execute('''
    INSERT OR IGNORE INTO verbas (cod_verba, tipo_verba, desc_verba) VALUES (?, ?, ?)
    ''', item)

# Salvando (commit) as alterações
conn.commit()

# Fechando a conexão
conn.close()
