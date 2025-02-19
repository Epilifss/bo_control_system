import pyodbc
import hashlib

def create_connection():
    try:
        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"  # Use o driver correto para o SQL Server
            "SERVER=tidelliserver;"                   # Nome ou IP do servidor
            "DATABASE=bo_system;"                     # Nome do banco de dados
            "UID=Totvs;"                              # Nome de usuário
            "PWD=totvs;"                              # Senha
        )
        return conn
    except pyodbc.Error as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        return None
    
def create_connection_mikonos():
    try:
        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"  # Use o driver correto para o SQL Server
            "SERVER=192.168.0.10;"                   # Nome ou IP do servidor
            "DATABASE=DADOSADV;"                     # Nome do banco de dados
            "UID=Totvs;"                              # Nome de usuário
            "PWD=totvs;"                              # Senha
        )
        return conn
    except pyodbc.Error as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        return None

def init_db():
    conn = create_connection()
    if conn is None:
        return
    
    try:
        cursor = conn.cursor()
        
        # Tabela de usuários
        cursor.execute('''
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='users' AND xtype='U')
            CREATE TABLE users (
                id INT IDENTITY(1,1) PRIMARY KEY,
                username VARCHAR(50) UNIQUE NOT NULL,
                password_hash VARCHAR(255) NOT NULL,
                module VARCHAR(50),
                is_admin BIT DEFAULT 0
            )
        ''')
        
        # Tabela de BOs
        cursor.execute('''
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='bo_records' AND xtype='U')
            CREATE TABLE bo_records (
                id INT IDENTITY(1,1) PRIMARY KEY,
                bo_number VARCHAR(20) UNIQUE NOT NULL,
                op INT,
                loja VARCHAR(100),
                nf_envio VARCHAR(50),
                nf_devolucao VARCHAR(50),
                op_venda VARCHAR(50),
                tipo_ocorrencia VARCHAR(100),
                motivo VARCHAR(100),
                frete VARCHAR(50),
                setor_responsavel VARCHAR(100),
                status VARCHAR(50),
                previsao_embarque DATE,
                descricao TEXT,
                created_at DATETIME DEFAULT GETDATE(),
                dt_embarque DATE,
                modulo VARCHAR(50),
                filial VARCHAR(4)
            )
        ''')
        
        # Tabela de sequência para números de BO
        # cursor.execute('''
        #     IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='bo_sequence' AND xtype='U')
        #     CREATE TABLE bo_sequence (
        #         year INT PRIMARY KEY,
        #         last_number INT DEFAULT 0
        #     )
        # ''')
        
        # Inserir usuário admin padrão se não existir
        admin_password = hashlib.sha256("4dmin123@4".encode()).hexdigest()
        cursor.execute('''
            IF NOT EXISTS (SELECT * FROM users WHERE username = 'TI')
            INSERT INTO users (username, password_hash, is_admin)
            VALUES (?, ?, 1)
        ''', ("TI", admin_password))
        
        conn.commit()
        print("Banco de dados inicializado com sucesso!")
    except pyodbc.Error as e:
        print(f"Erro ao inicializar o banco de dados: {e}")
    finally:
        if conn:
            cursor.close()
            conn.close()

if __name__ == "__main__":
    init_db()