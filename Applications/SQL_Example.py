import cx_Oracle

username = 'seu_usuario'
password = 'sua_senha'
dsn = 'seu_host:porta/seu_servico'

def run_query():
    try:
        with cx_Oracle.connect(username, password, dsn) as connection:
            print("Conexão estabelecida com sucesso!")
            with connection.cursor() as cursor:
                cursor.execute("SELECT * FROM sua_tabela")
                for row in cursor:
                    print(row)
    except cx_Oracle.DatabaseError as e:
        print(f"Ocorreu um erro ao conectar ao banco de dados: {e}")

if __name__ == "__main__":
    run_query()