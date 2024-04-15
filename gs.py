# -*- coding: utf-8 -*-
from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

import json
import requests
import pandas as pd
import auth as auth

def numberToLetter(number):
    '''
        Obtem a combinação de letras da colauna correspondente ao número passado como parâmetro
    
        params:
                number - número da coluna a ser convertido para combinação de letras (int)

        return:
                letter - combinação de letras correspondente ao número passado como parâmetro
    '''
    
    ## Definindo variável em que será armazenada a combinação de letras corrrespondente a coluna
    letter = "" 

    ## Enquanto o número for diferente de 0
    while(number >= 0):

        ## se o número for maior do que 26 (quantidade de letras do alfabeto)
        if(number >= 26):

            ## Normalizando o número para que haja letra que o represente
            auxNum = int(number/26)
        else:
            
            ## Se o número for menor que 26, pegando o resto de divisão e somando 1 pois não existe letra correpondente ao número 0
            auxNum = (number%26)+1
        
        ## Obtendo o caractere correspondente ao número ASCII obtido
        auxLet = chr(64+auxNum)

        ## Anexando a letra obtida a string de retorno
        letter = letter + auxLet

        ## Atualizando o valor do número
        number = number-26

    return letter


class GoogleSpreadsheet:
    def __init__(self, SPREADSHEET_ID):
        '''
            Classe usada para fazer conexão com o Google Sheets

            params:
                    SPREADSHEET_ID - ID Spreadsheet obtida diretammente na URL (string). Exemplo.: https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit#gid=0. OBS.: É preciso que o usuário autenticado tenha acesso a planilha
            
            OBS.: Se desejar criar uma nova spreadsheet, instanciar o objeto com SPREADSHEET_ID = '' e usar o método createSpreadsheet
        '''

        ## Definindo a autorização de manipulação a nível de Google Drive para poder manipular a planilha com todas as permissões de edição 
        self.SCOPE = 'https://www.googleapis.com/auth/drive'
        
        ## Setando o ID da Spreadsheet como atributo
        self.SPREADSHEET_ID = SPREADSHEET_ID

        ## Criando objeto para autenticação
        authInst = auth.Auth(self.SCOPE)

        ## Obtendo credenciais
        credentials = authInst.get_credentials()

        ## Loggando no serviço do Google Sheets com as credenciais obtidas
        self.service = build('sheets', 'v4', http=credentials.authorize(Http()))


    def createSpreadsheet(self, spreadsheetTitle=False):
        '''
            Cria uma nova spreadsheet nova na conta do usuário autenticado
        
            params:
                    spreadsheetTitle (opcional) - Título da planilha (string). Default: False = Planilha Sem Título

            returns:
                    Nenhu
        '''


        ## Se houver título, criar spreadsheet com ele
        if(spreadsheetTitle):
            body = {
                    "properties": {
                                        "title": spreadsheetTitle
                                    } 
                    }
        
        ## Se não criar planilha sem título            
        else:
            body = {}

        ## Fazendo requisição de criação com o parâmetro de título e salvando a URL obtida como retorno da requisição como atributo 
        self.SPREADSHEET_ID = self.service.spreadsheets().create(fields='spreadsheetId', body=body).execute()["spreadsheetId"]


    def createPage(self, sheetName=False):
        '''
            Cria uma nova aba na spreadsheet de Id guardado no objeto instanciado

            params:
                    sheetName (opcional) - Título da aba (string). Default: False = 'Página' + número da página
        
            returns:
                    Nenhum
        '''

        ## Se houver título, criar aba com ele
        if(sheetName):

            body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                "title": sheetName,
                                }
                            }
                        }
                    ]
                }
        
        ## Se não, criar aba sem título
        else:
            body={
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                            "title": "",
                            }
                        }
                    }
                ]
            }

        ## Fazendo requisição de criação 
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.SPREADSHEET_ID, body=body).execute()


    def readLines(self, rangeName): 
        '''
            Retorna os dados de um range específico

            params:
                    rangeName - Intervalo a se obter os dados (string). Formato: Página!Range. Exemplo: Página1!A:Z
        
            returns:
                    pd.DataFrame com as colunas do range lido (pd.DataFrame)
        '''

        ## Fazendo requisição de obtenção de valores
        result = self.service.spreadsheets().values().get(
                                                        spreadsheetId=self.SPREADSHEET_ID,
                                                        range=rangeName
                                                    ).execute()
        
        ## Transformando os valores obtidos em matriz
        values = result.get('values', []) 
        
        ## Abortando caso não haja dados
        if not values:
            print('No data found.')
            return 

        ## Transformando os dados obtidos em um pd.DataFrame
        spreadsheetData = pd.DataFrame(values, columns=values[0], dtype='str')
        
        ## Droppando a coluna de index 
        spreadsheetData = spreadsheetData.drop(spreadsheetData.index[0])
        
        ## Obtendo lista de colunas
        colunas = list(spreadsheetData.columns)
        
        ## Setando os tipos de cada coluna
        for coluna in colunas:
            try:
                float(spreadsheetData[coluna].iloc[0].replace(',',''))
                spreadsheetData[coluna] = spreadsheetData[coluna].apply(lambda x: int(x))
    
            except:
                try:

                    spreadsheetData[coluna] = spreadsheetData[coluna].apply(lambda x: float(x.replace(",",'.')))
                except:
                    pass

        ## Retornando o DataFrame
        return spreadsheetData

    def getLastRow(self, rangeName):
        '''
            Obtém o número da última linha de um range dado

            params:
                    rangeName - Intervalo a se obter os dados (string). Formato: Página!Range. Exemplo: Página1!A:Z

            returns:
                    Número da última linha (int)
        '''
        ## Obtendo os dados da planilha e armazenando numa matriz
        values = self.service.spreadsheets().values().get(spreadsheetId=self.SPREADSHEET_ID, range=rangeName).execute()

        x = pd.DataFrame(values['values'])

        last_row = (x[x.columns[0]].count())

        ## Retornando número de linhas menos 1 
        return (last_row)
    
    def getLastColumn(self, line, rangeName):
        '''
            Obtém o número da última coluna de uma determinada linha

            params:
                    line                - Número da linha que se que obter a útima coluna (int)
                    rangeName           - Intervalo a se obter os dados (string). Formato: Página!Range. Exemplo: Página1!A:Z

            returns:
                    Número da última coluna (int)
        '''



        ## Obtendo os dados da planilha e armazenando numa matriz
        values = self.service.spreadsheets().values().get(spreadsheetId=self.SPREADSHEET_ID, range=rangeName).execute()

        ## Retornando número de colunas menos 1 da linha especificada
        return(len(values['values'][line]))

    def clearValues(self, rangeName):
        '''
            Limpa os valores de um range determinado
        
            params:
                    rangeName - Intervalo a se obter os dados (string). Formato: Página!Range. Exemplo: Página1!A:Z
        
            returns:
                    Nenhum
        '''

        ## Fazendo requisição de limpeza de células
        self.service.spreadsheets().values().clear(spreadsheetId=self.SPREADSHEET_ID, range=rangeName,body={}).execute()
    

    def setFormat(self, sheetId, rangeRow, rangeColumn, numberFormat = ["TEXT", ""], horizontalAlignment = "LEFT", fontColor = [0,0,0], backgroundColor = [255,255,255], isBold = False, isItalic = False, fontFamily = 'Arial', fontSize = 12):
        '''
            Configura a formatação de um intervalo determinado

            params:
                    sheetId             - Id da Sheet obtida diretammente na URL (string). Exemplo.: https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit#gid=<sheetId>
                    rangeRow            - Lista contendo dois valores inteiros, simbolizando, respectivamente, as linhas de ínicio e fim do range (lista). Exemplo: [1,25] - range da linha 1 a linha 25
                    rangeColumn         - Lista contendo dois valores inteiros, simbolizando, respectivamente, as colunas de ínicio e fim do range (lista). Exemplo: [1,3] - range da coluna A a coluna C
                    numberFormat        - Lista de strings, simbolizando, respectivamente, o tipo de dados na coluna e a máscara de formato (lista). Exemplo: ["NUMBER", "#.###,##"]. Default = ["TEXT", ""] (Texto Plano)
                    horizontalAlignment - Posição do texto no intervalo determinado (string). Default: "LEFT"
                    fontColor           - Código RGB representando a cor da fonte do intervalo informado (lista). Default = [0,0,0] (Preto)
                    backgroundColor     - Código RGB representando a cor da célula do intervalo informado (lista). Default = [255,255,255] (Branco)
                    isBold              - Define de o texto do intervalo dado estará em negrito - Default = False
                    isItalic            - Define de o texto do intervalo dado estará em itálico - Default = False
                    fontFamily          - Define a família da fonte do intervalo dado. Default = 'Arial'
                    fontSize            - Define o tamanho da fonte do intervalo dado. Default = 12

            returns:
                    Nenhum

        '''
        
        ## Setando as linhas e as colunas de começo e fim do intervalo  
        startRowIndex = rangeRow[0]
        endRowIndex = rangeRow[1]
        startColumnIndex = rangeColumn[0]
        endColumnIndex = rangeColumn[1]
        
        ## Setando a formatação do texto da células do intervalo
        fieldType = numberFormat[0]
        pattern = numberFormat[1]

        ## Criando dicionário com os parâmetros passados para a função
        param ={ "requests": [{ "repeatCell": { "cell": { "userEnteredFormat": { "backgroundColor": { "blue": backgroundColor[2] / 255, "green": backgroundColor[1] / 255, "red": backgroundColor[0] / 255 }, "horizontalAlignment": horizontalAlignment, "textFormat": { "fontFamily": fontFamily, "fontSize": fontSize, "foregroundColor": { "blue": fontColor[2] / 255, "green": fontColor[1] / 255, "red": fontColor[0] / 255 }, "bold": isBold, "italic": isItalic }, "numberFormat": { "pattern": pattern, "type": fieldType } } }, "range": { "endColumnIndex": endColumnIndex, "endRowIndex": endRowIndex, "sheetId": sheetId, "startColumnIndex": startColumnIndex, "startRowIndex": startRowIndex }, "fields": "UserEnteredFormat" } }] }
        
        ## Fazendo a requisição com os parâmtros definidos
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.SPREADSHEET_ID, body=param).execute()

    def writeLines(self, lines, sheetName, startColumn, startRow, isDataFrame = True, header = True):
        '''
            Escreve dados na planlha do objeto instaciado

            params:
                    lines       - Conjunto de dados a ser inserido na tabela (matriz ou pd.DataFrame). OBS.: como padrão, espera-se um pd.DataFrame, mas também pode ser uma matriz, desde que o parâmetro isDataFrame seja False
                    sheetName   - Nome da aba que receberá os dados (string). Exemplo: 'Página1'
                    startColumn - Coluna de ínicio da inserção de dados (int). OBS.: Primeira coluna = 0
                    startRow    - Linha de ínicio da inserção de dados (int). OBS.: Primeira linha = 0
                    isDataFrame - Informa se os dados a serem inseridos estão em um pd.DataFrame. Default = True. OBS.: se for False, os dados devem estar em uma matriz
                    header      - Seleciona os cabeçalhos do pd.DataFrame como primeira linha de dados a serem inseridas no intervalo dado. Default = True. OBS.: também pode-se passar uma lista com nomes de colunas alternativos para substituir os nomes das colunas do pd.DataFrame

            returns:
                    Nenhum        
        '''

        ## Definindo coluna e linha de início da inserção dos dados
        startColumn = numberToLetter(startColumn)
        startRow = startRow + 1

        
        ## Adicionando o cabeçalho a matriz se o parâmetro header for verdadeiro
        if(header):
            try:
                header = list(lines.columns)
            except:
                pass
                
        
        ## Convertendo os dados em matriz se ele estiver em um objeto pd.DataFrame
        if(isDataFrame):
            lines = lines.values.tolist()
            if(header):
                lines.insert(0,header)
            
        ## Definindo o intervalo de inserção de dados
        rangeName = sheetName + "!" + str(startColumn) + str(startRow)
        
        ## Definindo parâmetros da requisição
        body = {
            'values': lines
        }

        ## Fazendo a requisição
        self.service.spreadsheets().values().update(spreadsheetId = self.SPREADSHEET_ID,valueInputOption = "USER_ENTERED",range=rangeName, body=body).execute()