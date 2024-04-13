import mysql.connector as mysql
import xlsxwriter
import xlsxwriter.exceptions
import logging
import datetime

# Configuração do logger para registrar exceções e informações
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s\n")

# Configuração do manipulador de arquivo para o logger
file_handler = logging.FileHandler("info.log")
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# Solicita o nome do banco de dados e da tabela ao usuário
mydatabase = input("Database: ")wdadwadwa
table_name = input("Table: ")
user = "root"
date = datetime.datetime.now()

# Função para conectar ao banco de dados MySQL e obter os dados da tabela
def connect_to_mysql(table_name):
    try:
        # Conexão ao banco de dados MySQL
        mysql_conn = mysql.connect(
            host="localhost",
            port=3306,
            user=user,
            password="ADMIN",
            database=mydatabase,
            charset="utf8"
        )

        if mysql_conn.is_connected():
            # Se conectado com sucesso, obtém informações sobre a versão do MySQL
            db_info = mysql_conn.get_server_info()
            logger.info("Connected to MySQL version " + db_info)
            print("Connected to MySQL version", db_info)

            cursor = mysql_conn.cursor()

            cursor.execute('SELECT * FROM ' + table_name)  # Executa a consulta para obter os dados da tabela
            header = [row[0] for row in cursor.description]  # Obtém os nomes das colunas da tabela
            rows = cursor.fetchall()  # Obtém todas as linhas de dados da tabela

            # Fecha a conexão com o banco de dados
            mysql_conn.close()

    except mysql.Error as e:
        logger.error("Failed to connect: " + str(e))
        print("Failed to connect:", e)
        raise

    return header, rows

# Função para exportar dados para um arquivo Excel
def export_to_excel():
    try:
        # Cria um novo arquivo Excel com base no nome do banco de dados, tabela e data atual
        excel_file_name = 'DB ' + mydatabase.capitalize() + ' - ' + 'Table ' + table_name.capitalize() + ' ' + date.strftime("%d-%m") + '.xlsx'
        logger.info("Creating Excel file: " + excel_file_name)
        print("Creating Excel file:", excel_file_name)

        workbook = xlsxwriter.Workbook(excel_file_name)
        worksheet = workbook.add_worksheet(table_name)

        # Define o formato das células do cabeçalho e do corpo
        header_cell_format = workbook.add_format({'bold': True, 'border': True, 'bg_color': 'green'})
        body_cell_format = workbook.add_format({'border': False})

        # Conecta-se ao MySQL para obter os dados da tabela
        header, rows = connect_to_mysql(table_name)

        row = 0
        column = 0

        # Escreve os nomes das colunas na primeira linha da planilha
        for column_name in header:
            worksheet.write(row, column, column_name, header_cell_format)
            column += 1

        row += 1

        # Escreve os dados das linhas na planilha
        for row_data in rows:
            column = 0
            for cell_data in row_data:
                if isinstance(cell_data, bytes):
                    worksheet.write(row, column, 'Image', body_cell_format)  # Tratamento especial para BLOBs
                else:
                    worksheet.write(row, column, cell_data, body_cell_format)
                column += 1

            row += 1

        # Fecha o arquivo Excel após a escrita dos dados
        logger.info(str(row) + ' rows written successfully to ' + excel_file_name)
        print(str(row) + ' rows written successfully to ' + excel_file_name)
        workbook.close()

    except xlsxwriter.exceptions.XlsxWriterException as e:
        logger.error("Failed to export to Excel: " + str(e))
        print("Failed to export to Excel:", e)
        raise

try:
    export_to_excel()

except Exception as e:
    logger.exception("An error occurred:")
    print("An error occurred:", e)
