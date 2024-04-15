# Importando as bibliotecas necessárias
import requests
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account
from gs import GoogleSpreadsheet
from httplib2 import Http
from oauth2client import file, client, tools


# Definindo a chave da API e outras configurações
api_key = "apikey"
time_granularity = "Daily" # Daily, Weekly, Monthly
main_domain_only = False
# Domínios para os quais queremos obter dados
domains = ["mercadolivre.com.br", "americanas.com.br", "magazineluiza.com.br","casasbahia.com.br", "amazon.com.br"]


# Função para validar se a data está no formato correto
def dateFormat(date):
    split = date.split('-')
    if len(split) != 3:
        return False
    
    year, month, day = split
    
    if len(year) == 4 and len(month) == 2 and len(day) == 2:
        int(year)
        int(month)
        int(day)
        return True
    else:
        return False


# Função para comparar as datas e garantir que a inicial seja antes da final
def dateCompare(start_date, end_date):
    start_year, start_month, start_day = map(int, start_date.split('-'))
    end_year, end_month, end_day = map(int, end_date.split('-'))
    
    if start_year < end_year or (start_year == end_year and start_month < end_month) or (start_year == end_year and start_month == end_month and start_day <= end_day):
        return True
    else:
        return False


# Laço para obter datas válidas
while True:
    start_date = input("Digite a Data Inicial (YYYY-MM-DD): ")
    end_date = input("Digite a Data Final (YYYY-MM-DD): ")
    
    # Verifica se as datas são válidas e se o intervalo está correto
    if dateFormat(start_date) and dateFormat(end_date) and dateCompare(start_date, end_date):
        print("Datas válidas e no intervalo correto.")
        break
    else:
        print("Datas inválidas ou intervalo incorreto. Tente novamente.")

'''
# Definindo a chave de autenticação do Google Sheets
spreadsheet_id = 'idsheetskey'
sheets = GoogleSpreadsheet(spreadsheet_id)
'''

def searchTraffic():
    '''
    Extrai dados de tráfego para os domínios no período de tempo especificado.
    '''
    data_list = []
    
    for website in domains:
         # Monta a URL para obter os dados de tráfego
         url = "https://api.similarweb.com/v1/website/"+website+"/total-traffic-and-engagement/visits?api_key="+api_key+"&start_date="+start_date+"&end_date="+end_date+"&country=br&granularity="+time_granularity    
         
         headers = {"accept": "application/json"}
         response = requests.get(url, headers=headers)
         data = response.json()
        
         # Verifica se a resposta contém os dados de visitas
         if "visits" in data:
             for entry in data["visits"]:
                 # Adiciona os dados na lista
                 data_list.append({ "Date": entry["date"], "Domain": website, "Visits": entry["visits"]})
                    
    # Cria um DataFrame e o pivota
    df = pd.DataFrame(data_list)
    df_pivoted = df.pivot(index="Date", columns='Domain', values='Visits')
    
    # Salva os dados em um arquivo Excel
    df_pivoted.to_excel('visitas.xlsx', index=True)
    print("Dados salvos como visitas.xlsx")
'''
    # Criando uma instância da classe GoogleSpreadsheet
    gs = GoogleSpreadsheet(spreadsheet_id)

    # Adicionando uma nova coluna "date" para ser inserida no Google Sheets e no Excel
    df['date_for_sheets'] = df['Date']

    # Mantendo os domínios pivotados no DataFrame
    df_pivoted = df_pivoted.reset_index()

    # Enviando os dados para o Google Sheets
    range_name = 'Trafego!A1:F'  
    gs.writeLines(df_pivoted, 'Trafego', 0, 0, isDataFrame=True, header=True)
    print('Dados enviados para o Google Sheets com sucesso!')
'''
searchTraffic()


def searchChannel():
    '''
    Extrai dados de tráfego por channel para os domínios no período de tempo especificado.
    '''
    data_list = []

    for website in domains:
        # Monta a URL para obter os dados de tráfego
        url_channel = "https://api.similarweb.com/v1/website/"+website+"/traffic-sources/overview-share?api_key="+api_key+"&start_date="+start_date+"&end_date="+end_date+"&country=br&granularity="+time_granularity

        # Faz as requisições para obter os dados
        response_channel = requests.get(url_channel, headers={"accept": "application/json"})

        # Converte as respostas em formato JSON
        data_channel = response_channel.json()["visits"][website]

        # Percorre os sites e suas fontes de tráfego na resposta JSON
        for item in data_channel:
            source_type = item["source_type"]
            for visit in item["visits"]:
                date = visit["date"]
                organic = visit["organic"]
                paid = visit["paid"]
                channel_traffic = organic + paid  # Calcula o tráfego total

                if source_type == "Search":
                    # Adiciona os dados na lista separando organic e paid
                    data_list.append({"Date": date, "Domain": website, "Channel": "Organic Search", "Channel Traffic": organic, "Traffic Type": "Organic"})
                    data_list.append({"Date": date, "Domain": website, "Channel": "Paid Search", "Channel Traffic": paid, "Traffic Type": "Paid"})
                else:
                    # Adiciona os dados na lista para outros source_type
                    channel_traffic = organic + paid  # Calcula o tráfego total
                    data_list.append({"Date": date, "Domain": website, "Channel": source_type, "Channel Traffic": channel_traffic, "Traffic Type": "Other"})
        
    # Cria um DataFrame
    df = pd.DataFrame(data_list)

    # Agrupa os dados por Date, Domain, Channel e soma os valores de Channel Traffic
    df_grouped = df.groupby(['Date', 'Domain', 'Channel'])['Channel Traffic'].sum().reset_index()

    # Renomeia a coluna 'Traffic Type' para 'Channel Traffic'
    df_grouped = df_grouped.rename(columns={'Traffic Type': 'Channel Traffic'})

    # Salva os dados em um arquivo Excel
    df_grouped.to_excel('canais.xlsx', index=False)
    print("Dados salvos como canais.xlsx")
'''
    # Criando uma instância da classe GoogleSpreadsheet
    gs = GoogleSpreadsheet(spreadsheet_id)

    # Enviando os dados para o Google Sheets
    range_name = 'Canais!B2:F'  
    gs.writeLines(df_grouped, 'Canais', 1, 1, isDataFrame=True, header=False)
    print('Dados enviados para o Google Sheets com sucesso!')
'''
searchChannel()