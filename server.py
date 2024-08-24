import sqlite3
import os

def create_tables():
    db_name = 'contab_reciprev.db'
    with sqlite3.connect(db_name) as conn:
        cursor = conn.cursor()

        # Criando a tabela 'verbas'
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS verbas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                verba TEXT NOT NULL UNIQUE,
                tipo TEXT NOT NULL,
                descricao TEXT NOT NULL,
                quantidade INTEGER NOT NULL DEFAULT 0,
                total_valor REAL NOT NULL DEFAULT 0.0
            )
        ''')

        # Criando a tabela 'orgaos'
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS orgaos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                verba_id INTEGER NOT NULL,
                orgao TEXT NOT NULL,
                FOREIGN KEY (verba_id) REFERENCES verbas(id) ON DELETE CASCADE
            )
        ''')

        # Criando a tabela 'classificacoes'
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS classificacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                verba_id INTEGER NOT NULL,
                classificacao TEXT NOT NULL,
                FOREIGN KEY (verba_id) REFERENCES verbas(id) ON DELETE CASCADE
            )
        ''')

        # Criando a tabela 'categorias'
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS categorias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                verba_id INTEGER NOT NULL,
                categoria TEXT NOT NULL,
                FOREIGN KEY (verba_id) REFERENCES verbas(id) ON DELETE CASCADE
            )
        ''')

        # Criando a tabela 'elementos'
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS elementos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                verba_id INTEGER NOT NULL,
                elemento TEXT NOT NULL,
                FOREIGN KEY (verba_id) REFERENCES verbas(id) ON DELETE CASCADE
            )
        ''')

        # Criando a tabela 'codigo_orgaos'
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS codigo_orgaos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                verba_id INTEGER NOT NULL,
                codigo_orgao TEXT NOT NULL,
                FOREIGN KEY (verba_id) REFERENCES verbas(id) ON DELETE CASCADE
            )
        ''')

        # Criando índices para melhorar a performance das consultas
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_verba_id_orgaos ON orgaos(verba_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_verba_id_classificacoes ON classificacoes(verba_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_verba_id_categorias ON categorias(verba_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_verba_id_elementos ON elementos(verba_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_verba_id_codigo_orgaos ON codigo_orgaos(verba_id)')

    # Verifica se o banco de dados foi criado
    if os.path.exists(db_name):
        print(f"Banco de dados '{db_name}' criado com sucesso.")
    else:
        print(f"Falha ao criar o banco de dados '{db_name}'.")

if __name__ == "__main__":
    create_tables()