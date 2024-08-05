import pandas as pd
import mariadb
import time
import warnings
import PySimpleGUI as sg

warnings.filterwarnings("ignore")

# Função para criar a tabela no banco de dados
def criar_tabela(cursor, colunas):
    # Cria a string SQL para definir as colunas
    colunas_sql = ", ".join([f"{col['nome']} {col['tipo']}" for col in colunas])
    try:
        # Executa o comando SQL para criar a tabela
        cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS DADOS (
            ID INT AUTO_INCREMENT PRIMARY KEY,
            {colunas_sql}
        )
        """)
        conexao.commit()  # Confirma as mudanças no banco de dados
        sg.popup("Tabela criada com sucesso")  # Mostra mensagem de sucesso
    except mariadb.Error as e:
        sg.popup(f"Erro ao criar a tabela: {e}")  # Mostra mensagem de erro

# Função para importar dados do Excel e inserir no banco de dados
def importar_dados(filepath, cursor, colunas):
    df = pd.read_excel(filepath)  # Lê o arquivo Excel
    required_columns = [col['nome'] for col in colunas]  # Lista de colunas necessárias
    for col in required_columns:
        if col not in df.columns:
            sg.popup(f"Coluna {col} não encontrada no DataFrame.")  # Mostra mensagem de erro se a coluna não for encontrada
            return

    try:
        inicio = time.time()  # Inicia o cronômetro
        for index, row in df.iterrows():
            valores = []
            for col in colunas:
                if col['tipo'].startswith('DATE'):
                    valor = pd.to_datetime(row[col['nome']], format='%d/%m/%Y').date()
                else:
                    valor = row[col['nome']]
                valores.append(valor)
            placeholders = ", ".join(["%s"] * len(valores))
            sql = f"INSERT INTO DADOS ({', '.join(required_columns)}) VALUES ({placeholders})"
            cursor.execute(sql, valores)
            conexao.commit()  # Confirma as mudanças no banco de dados
        final = time.time()  # Termina o cronômetro
        sg.popup("Dados inseridos com sucesso no SQL", f'Tempo de Processamento: {int(final - inicio)} segundos')
    except mariadb.Error as e:
        sg.popup(f"Erro ao inserir dados no banco de dados: {e}")

# Função para conectar ao banco de dados
def conectar_banco(host, user, password, database, port):
    try:
        conexao = mariadb.connect(
            host=host, 
            user=user,
            password=password, 
            database=database,
            port=port
        )
        return conexao
    except mariadb.Error as e:
        sg.popup(f"Erro ao conectar ao banco de dados: {e}")
        return None

# Layout da Interface Gráfica para login no servidor
layout_login = [
    [sg.Text('Host: '), sg.InputText('localhost', key='host')],
    [sg.Text('Usuário: '), sg.InputText('root', key='user')],
    [sg.Text('Senha: '), sg.InputText('', password_char='*', key='password')],
    [sg.Text('Banco de Dados: '), sg.InputText('mydatabase', key='database')],
    [sg.Text('Porta: '), sg.InputText('3303', key='port')],
    [sg.Button('Conectar')]
]

# Layout da Interface Gráfica para definir a estrutura da tabela
#quantidade de input = Quantidade de colunas (Cabeçalho)
layout_definicao = [
    [sg.Text('Definir Colunas da Tabela')],
    [sg.Text('Nome da Coluna'), sg.Input(key='nome_coluna')],
    [sg.Text('Tipo da Coluna'), sg.Combo(['VARCHAR(255)', 'DECIMAL(10, 2)', 'DATE'], key='tipo_coluna')],
    [sg.Button('Adicionar Coluna'), sg.Button('Criar Tabela')]
]

# Layout da Interface Gráfica para importar dados
layout_importacao = [
    [sg.Text('Selecione o arquivo Excel: '), sg.Input(), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
    [sg.Button('Importar')]
]

layout_principal = [
    [sg.Frame('Definição da Tabela', layout_definicao)],
    [sg.Frame('Importação de Dados', layout_importacao)]
]

# Criar a janela de login
window_login = sg.Window('Login no Servidor', layout_login)

# Loop para eventos da interface de login
while True:
    event, values = window_login.read()
    if event == sg.WINDOW_CLOSED:
        break
    if event == 'Conectar':
        host = values['host']
        user = values['user']
        password = values['password']
        database = values['database']
        port = int(values['port'])
        
        conexao = conectar_banco(host, user, password, database, port)
        if conexao:
            cursor = conexao.cursor()
            window_login.close()
            break

if conexao:
    # Criar a janela principal
    window_principal = sg.Window('Definir Tabela e Importar Dados', layout_principal)

    colunas = []

    # Loop para eventos da interface gráfica principal
    while True:
        event, values = window_principal.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == 'Adicionar Coluna':
            nome_coluna = values['nome_coluna']
            tipo_coluna = values['tipo_coluna']
            if nome_coluna and tipo_coluna:
                colunas.append({'nome': nome_coluna, 'tipo': tipo_coluna})
                sg.popup(f"Coluna {nome_coluna} ({tipo_coluna}) adicionada.")
            else:
                sg.popup("Por favor, defina tanto o nome quanto o tipo da coluna.")
        if event == 'Criar Tabela':
            if colunas:
                criar_tabela(cursor, colunas)
            else:
                sg.popup("Nenhuma coluna definida.")
        if event == 'Importar':
            filepath = values[0]
            if filepath:
                importar_dados(filepath, cursor, colunas)
            else:
                sg.popup("Por favor, selecione um arquivo Excel.")

    conexao.close()
    window_principal.close()
