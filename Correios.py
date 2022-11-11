import requests
from bs4 import BeautifulSoup
import pygsheets # conectar no Sheets
import time
import logging

path = ('credenciais.json')
gc = pygsheets.authorize(service_account_file=path)

# acessando e logando na planilha do Sheets
planilha = gc.open_by_key('id da planilha')
aba = planilha.worksheet_by_title('Conferencia')

i = 2

logging.basicConfig(filename='Consulta Correios.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s')

# informação para o usuario
logging.warning("Planilha conectada com sucesso")
print("Planilha conectada com sucesso")

# Calculando quantidade de pesquisas
celulas = aba.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix')
total_celulas = len(celulas)

total = str(total_celulas)

logging.warning("Total de consultas a serem realizadas: " + total)
print("Total de consultas a serem realizadas: " + total)

while i <= total_celulas:

    try:
        codigo1 = aba.get_value((i, 2))
        codigo = str(codigo1)
        
        # acessando site da API
        iniciourl = 'https://linketrack.com/track?codigo='
        urlfinal = '&utm_source=track'
        url = iniciourl + codigo + urlfinal
        r = requests.get(url)
         
        # usando Beautiful Soup para extrair dados do site da API
        soup = BeautifulSoup(r.text, 'html.parser')

        status = soup.find_all("span", class_='')
    
        logging.warning("Pesquisa realizada no site dos correios")
        print("Pesquisa realizada no site dos correios")
        
        # conversao e tratamento de dados
        statusmenos = str(status[1])

        status1 = statusmenos.replace("<span>", "")
        statusfinal = status1.replace("</span>", "")

        aba.update_value((i, 3), statusfinal)

        logging.warning("Planilha atualizada com sucesso")
        print("Planilha atualizada com sucesso")

        i = i + 1
        
        # tempo de espera para evitar erros com a API
        time.sleep(5)

    except:
        logging.warning("Ocorreu um erro inesperado. Tentando novamente em 5 segundos")
        print("Ocorreu um erro inesperado. Tentando novamente em 5 segundos")
        aba.update_value((i, 3), "Erro ao pesquisar")
        time.sleep(5)

logging.warning("Consultas finalizadas!")
