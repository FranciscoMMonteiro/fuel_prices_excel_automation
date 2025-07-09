# IMPORTANDO BIBLIOTECAS
import sys
import os
user = os.getlogin()
import os
sys.path.append()
sys.path.append()
import pymongo
import mongo # type: ignore
import MongoDB # type: ignore
import pandas as pd
import datetime
import xlwings as xw
from dateutil.relativedelta import relativedelta
import numpy as np
import sys
import warnings
import ctypes
import time
import ast

# DEFININDO FUNÇÕES

def add_new_tickers_to_import(mongoBD_collection_tickers_to_import, list_tickers_check):
        db_ticker_list = []
        for doc in mongoBD_collection_tickers_to_import.find({'type':'energy'}):
            db_ticker_list.append(doc['ticker']['bbg'])
        missing_tickers = [ticker for ticker in list_tickers_check if ticker not in db_ticker_list]
        if missing_tickers:
            list_json_upload = []
            for ticker in missing_tickers:
                json_i = {}
                ticker_cuted = ticker.replace(' Index','')
                ticker_cuted = ticker_cuted.replace(' Comdty','')
                json_i['_id'] = ticker_cuted
                json_i['ticker'] = {'bbg':ticker}
                json_i['in_use'] = True
                json_i['type'] = 'energy'
                json_i['fields'] = 'bbg_energy'

                list_json_upload.append(json_i)
            adicionar_bd = int(input("Achamos tickers na planilha que não estão presentes no Banco de Dados.\n"+
                                  "Deseja adiciona-los?.\n"+
                                  '1 - Sim\n' +
                                  '2 - Não\n'+
                                  ""))
            if adicionar_bd == 1:
                mongo.bulk_update(mongoBD_collection_tickers_to_import,list_json_upload)
                print('Tickers adicionados ao banco de dados de tickers a serem puxados.\n' +
                       'Caso queria pegar os preços rodar rotina de preços da bloomberg.')
            else:
                print('Tickers não foram incluidos na rotina de preços a serem puxados.')


def bring_cmd_to_front():
    # Get the current process ID
    pid = os.getpid()

    # Get the window handle of the command prompt associated with the current process
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()

    if hwnd:
        # Bring the window to the foreground
        ctypes.windll.user32.SetForegroundWindow(hwnd)

def acha_quarter_atual(mes):
    if mes in [1,2,3]:
        return 1
    elif mes in [4,5,6]:
        return 2
    elif mes in [7,8,9]:
        return 3
    elif mes in [10,11,12]:
        return 4
    else:
        return None
    
def get_dataframe_price(ticker_list:list, mdb:pymongo.MongoClient,start_date:datetime.datetime=datetime.datetime(1,1,1), end_date:datetime.datetime = datetime.datetime(9999,12,31)):
    """
    Get prices from the MongoDB database based on the ticker passed and return it into a pandas dataframe.

    Parameters:
    ticker (str): A single ticker symbol (str) or a list of ticker symbols (list) to plot.
    MongoClient (MongoClient): The MongoDB client instance connected to the database.

    Returns:
    df: Dataframe with prices for each date of the ticker passed.
    """
    query = {'$and': [
        {'_id.ticker': {"$in": ticker_list}},
        {'_id.date': {"$gte": start_date, "$lte": end_date}}]}
    documents = mdb.client['etl']['bbg_raw.daily'].find(query)
    list_for_df = []
    found_tickers = set()
    for doc in documents:
        row = {'ticker':doc['_id']['ticker'],'date':doc['_id']['date'], 'PX_LAST':doc['PX_LAST']}
        list_for_df.append(row)
        found_tickers.add(doc['_id']['ticker'])
    df = pd.DataFrame(list_for_df)
    df['date'] = pd.to_datetime(df['date'])
    df = df.sort_values(by='date',ascending=True)

    not_found_tickers = [ticker for ticker in ticker_list if ticker not in found_tickers]
    return df , not_found_tickers


def atualiza_header(sheet,cell_beginning_header,quarters_list):
    sheet.range(cell_beginning_header,cell_beginning_header.end("right")).clear_contents()
    i = 0
    for quarter in quarters_list:
        sheet.range(cell_beginning_header).offset(0,i).value = quarter
        i = i + 1

    return None

def add_months(row):
    return row['date'] + relativedelta(months=int(row['delta_month']))


def main():

    print('Bem-vindo ao atualizador de preços de energia. Digite uma das opções abaixo.')

    # INICIANDO CONEXÃO COM O BD
    tipo_bd = "PROD"
    mdb = MongoDB.OurMongoClient(MongoDB.get_mongo_conn(environment=tipo_bd))

    # PEGANDO VARIAVEL DE OBJETO EXCEL
    folder = fr''
    name = "MS_Price_Deck_new.xlsm"
    file_path = os.path.join(folder, name)
    wb = xw.Book(file_path)
    sheet = wb.sheets("MarketPrices")



    # Pegando celulas de referencia para puxar tickers
    last_cell_spot = sheet.range(sheet.cells.last_cell.row,sheet.range('bbgticker').column).end('up')
    initial_cell_spot = sheet.range('bbgticker').offset(1,0)
    last_cell_spot = sheet.range(sheet.cells.last_cell.row,initial_cell_spot.column).end('up')

    initial_cell_foward = sheet.range('bbgtickerfoward').offset(1,0)
    last_cell_foward = last_cell_spot.offset(0,1)

    list_rows = []
    for cell in list(sheet.range(initial_cell_spot,last_cell_spot).rows):
        list_rows.append(cell.row)
    list_ticker_spot = list(sheet.range(initial_cell_spot,last_cell_spot).value)
    list_ticker_foward = list(sheet.range(initial_cell_foward,last_cell_foward).value)
    dict_rows = {list_rows[i]: [list_ticker_spot[i], list_ticker_foward[i]] for i in range(len(list_rows))}
    #list_ticker_spot.remove(None)
    df_rows = pd.DataFrame(dict_rows, index = ['bbgticker','bbgtickerfoward']).T.reset_index().rename(columns = {'index':'row'})
    df_rows['bbgtickerfoward_raiz'] = df_rows['bbgtickerfoward'].str.split().str[0].str[:-1]

    # Variaveis de tempo para puxar informações de tickers do BD
    data_atual = datetime.datetime.today()
    mes_atual = data_atual.month
    quarter_atual = acha_quarter_atual(mes_atual)
    quarter_atual_str = f'{quarter_atual}Q{str(data_atual.year)[2:]}'
    foward_36 = datetime.datetime.today() + relativedelta(months=36)


    

    # Cria opcao de puxar dados do dia atual ou de algum dia especifico
    loop_opcoes = True
    while loop_opcoes:
        opcao = int(input(
                "1- Atualizar para data de hoje.\n"
                "2- Atualizar para data especificada.\n"
                "\n"
                "Digite o numero da opção desejada:\n"
                ""
                ))
        loop_opcoes = False
        if opcao == 1:
            data_especificada = data_atual
        elif opcao == 2:
            data_especificada = input("Digite a data de interesse no formato yyyy-mm-dd (ou yyyy/mm/dd): ")
            if "/" in data_especificada:
                data_especificada = datetime.datetime.strptime(data_especificada, "%Y/%m/%d")
            else:
                data_especificada = datetime.datetime.strptime(data_especificada, "%Y-%m-%d")
            if data_especificada>datetime.datetime.today():
                loop_opcoes = True
                print('Não aceitamos datas futuras. Selecionar uma data passada.')
                print('')
        else:
            print('')
            loop_opcoes = True
    
    
    # Puxando tickers spot do BD para pegar preços historicos
    df_spot = pd.DataFrame()
    list_ticker_spot = list(set(list_ticker_spot))
    list_ticker_spot.remove(None)
    list_tickers_not_found = []
    #for ticker in tqdm(list_ticker_spot, desc = "Importando preços historicos do Banco de Dados") :
        #print(ticker)
    print()
    print("Importando preços historicos do Banco de Dados")
    print()

    try:
        #df_i = get_dataframe_price(ticker,mdb,datetime.datetime(2011,1,1),data_atual)
        #df_spot = pd.concat([df_spot,df_i])
        df_spot, tickers_not_found_spot = get_dataframe_price(list_ticker_spot,mdb,datetime.datetime(2011,1,1),data_especificada)
    except Exception as e:
        print(e)
        # list_tickers_not_found.append(ticker)

    if tickers_not_found_spot:
        print('Não achou no banco de dados os tickers de preços historicos:')
        print(tickers_not_found_spot)
        print()
    else:
        print("Todos os tickers de preços historicos foram encontrados.")
        print()


    # Retirando dias onde não há preços (feriados, fins de semanas etc)
    df_spot = df_spot.dropna(subset='PX_LAST')

    df_spot['year'] = df_spot['date'].dt.year
    df_spot['month'] = df_spot['date'].dt.month
    df_spot['quarter'] = df_spot['month'].apply(acha_quarter_atual)

    # PEGANDO PREÇOS FUTUROS

    list_ticker_foward_all_quarters = []
    list_ticker_foward = list(set(list_ticker_foward))
    list_ticker_foward.remove(None)
    for ticker in list_ticker_foward:
        for i in range(0,31):
            list_splited_ticker = ticker.split()
            list_splited_ticker[0] = list_splited_ticker[0][:-1] + str(i)
            ticker_adjusted = " ".join(list_splited_ticker)
            list_ticker_foward_all_quarters.append(ticker_adjusted)


    df_foward = pd.DataFrame()
    tickers_not_found = []
    #for ticker in tqdm(list_ticker_foward_all_quarters , desc = "Importando preços futuros do Banco de Dados"):
        #print(ticker)
    print("Importando preços futuros do Banco de Dados")
    print()
    try:
    #    df_i = get_dataframe_price(ticker,mdb,data_atual - datetime.timedelta(days=14),data_atual)
    #    df_foward = pd.concat([df_foward,df_i])
        df_foward, tickers_not_found_foward = get_dataframe_price(list_ticker_foward_all_quarters,mdb,data_especificada - datetime.timedelta(days=14),data_especificada)
    except Exception as e:
        #tickers_not_found.append(ticker)
        print(e)
        print()

    if tickers_not_found_foward:
        print(f'Não achou no banco de dados os tickers de preços futuros: {tickers_not_found_foward}')
        print()
    else:
        print("Todos os tickers de preços futuros foram encontrados.")
        print()

    # Checando data dos dados a serem inputados na planilha
    continuar = False
    bring_cmd_to_front()
    if df_spot['date'].max().strftime('%Y-%m') != data_especificada.strftime('%Y-%m'):
        print("Atualizar preços no banco de dados.")
        user_input = "n"
    else:
        user_input = input(f"Ultimos dados de preços historicos disponiveis são de {df_spot['date'].max().day}/{df_spot['date'].max().month}/{df_spot['date'].max().year}(dd/mm/aaaa).\n"
                       f"Ultimos dados de preços futuros disponiveis são de {df_foward['date'].max().day}/{df_foward['date'].max().month}/{df_foward['date'].max().year}(dd/mm/aaaa).\n"
                        "Caso queira dados mais recentes, checar a ingestão de dados da bloomberg para o banco de dados.\n"
                        "Deseja continuar? (s/n) :\n"+
                        "")
    
    while continuar == False:
        if user_input == "n":
            bring_cmd_to_front()
            input("Encerrando atualização. Pressione enter para sair.")
            sys.exit()
        elif user_input == "s":
            print("Prosseguindo com a atualização.")
            print()
            continuar = True
        else:
            bring_cmd_to_front()
            user_input = input("Não foi possivei idendificar o comando.\n"
                                f"Ultimos dados historicos disponiveis são de {df_spot['date'].max().day}/{df_spot['date'].max().month}/{df_spot['date'].max().year}(dd/mm/aaaa).\n"
                                f"Ultimos dados de preços futuros disponiveis são de {df_foward['date'].max().day}/{df_foward['date'].max().month}/{df_foward['date'].max().year}(dd/mm/aaaa).\n"
                                "Caso queira dados mais recentes, checar a ingestão de dados da bloomberg para o banco de dados.\n"
                                "Deseja continuar? (s/n) :")


    # Manegandos dataframes para deixar no formato ideal

    df_spot = df_spot.groupby(by = ['ticker','year','quarter','month'])['PX_LAST'].mean().reset_index()
    df_spot = df_spot.pivot_table(values= 'PX_LAST', index = 'ticker', columns = ['year','quarter','month'])

    df_foward = df_foward.dropna(subset='PX_LAST')
    df_foward = df_foward[~(df_foward['PX_LAST']=="#N/A Invalid Security")]

    df_foward_aux = df_foward.groupby(['ticker'])[['date']].max()
    df_foward_aux = df_foward_aux.rename(columns = {'date':'last_date_available'}).reset_index()

    df_foward = pd.merge(df_foward, df_foward_aux, on= "ticker")
    df_foward = df_foward[df_foward['date'] == df_foward['last_date_available']]
    df_foward = df_foward.drop(columns = 'last_date_available')

    df_foward['delta_month'] = df_foward['ticker'].apply(lambda x: x.split()[0][5:])

    df_foward['future_date'] = df_foward.apply(add_months, axis =1)
    df_foward['raiz_ticker'] = df_foward['ticker'].apply(lambda x: x.split()[0][:5])
    df_foward['future_month'] = df_foward['future_date'].dt.month
    df_foward['future_year'] = df_foward['future_date'].dt.year
    df_foward['future_quarter'] = df_foward['future_month'].apply(lambda x : acha_quarter_atual(x))
    df_foward['quarter_year_str'] = df_foward['future_quarter'].astype(str) + 'Q' +  df_foward['future_year'].astype(str).str[2:]

    df_foward = df_foward.pivot_table(values='PX_LAST', index = 'raiz_ticker' , columns = ['future_year','future_quarter','future_month'])


    df_spot.columns = df_spot.columns.values
    df_spot.reset_index(inplace=True)
    df_foward.columns =  df_foward.columns.values
    df_foward.reset_index(inplace=True)


    df_rows = pd.merge(df_rows,df_spot, left_on = 'bbgticker' , right_on='ticker', how = 'left',suffixes=('', '_to_drop'))
    df_rows = pd.merge(df_rows,df_foward, left_on = 'bbgtickerfoward_raiz' , right_on='raiz_ticker', how = 'left',suffixes=('', '_to_drop'))
    string_columns = df_rows.columns[df_rows.columns.map(type) == str]
    columns_to_keep = string_columns[~string_columns.str.endswith('_to_drop')].tolist()
    columns_to_keep.extend(df_rows.columns[df_rows.columns.map(type) != str])
    df_rows = df_rows[columns_to_keep].copy()

    index_columns = ['bbgticker','bbgtickerfoward','bbgtickerfoward_raiz','ticker','raiz_ticker','row']
    new_columns = []
    for col in df_rows.columns:
        if isinstance(col,str) and col not in index_columns:
            new_col = ast.literal_eval(col)
        else:
            new_col = col
        new_columns.append(new_col)
        
    df_rows.columns = new_columns

    df_rows.set_index(index_columns, inplace=True)

    df_rows.columns = pd.MultiIndex.from_tuples(df_rows.columns , names = ['year', 'quarter','month'])

    df_rows = df_rows.groupby(level = ['year','quarter'], axis = 1).mean()

    n_quarters = len(df_rows.columns)
    init_clean = sheet['bbgticker'].offset(1,2)
    final_clean = sheet.range(last_cell_spot.row,sheet['bbgticker'].offset(0,2+ n_quarters).column)

    # Guardando ultimo ouput em uma nova sheet
    data_precos_ult_att = sheet['price_data'].value
    data_precos_ult_att_str = data_precos_ult_att.strftime('%Y.%m.%d')
    #data_especificada_str = data_especificada.strftime("%d.%m.%Y")
    last_date = sheet['data'].value.strftime("%d.%m.%Y")
    if f"MarketPrices{data_precos_ult_att_str}" in wb.sheet_names:
        print(f"Substituindo aba de preços de {data_precos_ult_att_str} pelos preços que estavam em MarketPrices, antes de pegar novos preços.")
        wb.sheets(f"MarketPrices{data_precos_ult_att_str}").delete()
    new_sheet = wb.sheets.add(f"MarketPrices{data_precos_ult_att_str}")
    source_range = sheet.used_range
    new_sheet.range('A1').value = source_range.value
    source_range.api.Copy()
    new_sheet.range('A1').api.PasteSpecial(Paste = -4122)
    sheet.api.Move(Before = wb.sheets[0].api)


    # Limpando a sheet atual
    sheet.range(init_clean,final_clean).clear_contents()

    # Colando valores do dataframe
    init_clean.options(index=False,header = False).value = df_rows 

    # Criando lista com nomes de quarters a serem colados
    list_quarters = []
    for col in df_rows.columns:
        list_quarters.append(str(col[1]) + 'Q' + str(col[0])[2:] )

    # Atualizando os quarters
    atualiza_header(sheet,sheet['beginning'],list_quarters)
    # Atualizando a ultima data de atualização
    sheet['data'].value = data_atual

    # Atualizando a ultima data de atualização
    sheet['price_data'].value = data_especificada

    # Checando se todos os tickers da planilha estão na tabela do banco de
    # dados que contem os tickers a serem importados diariamente e caso nao esteja
    # já inclui da tabela.
    list_check = list_ticker_spot + list_ticker_foward_all_quarters
    list_check = list(set(list_check))
    list_check = [item for item in list_check if item is not None]
    add_new_tickers_to_import(mdb['gestao']['asset.metadata'],list_check)

    
if __name__ == "__main__":
    inicio = time.time()
    warnings.filterwarnings('ignore')
    main()
    fim = time.time()
    delta_tempo = tempo_decorrido = fim - inicio
    print(f"Codigo levou {delta_tempo:.1f} segundos para rodar.")
    print()
    input("Atualização concluída com successo. Pressione ENTER para sair.")





#####################################
# Coisas a incluir ainda
# Preços de coal:
# - https://www.eia.gov/opendata/browser/coal/shipments/plant-state-aggregates?frequency=quarterly&data=ash-content;heat-content;price;quantity;sulfur-content;&start=2011-01&sortColumn=period;&sortDirection=desc;
# - https://api.eia.gov/v2/coal/shipments/plant-state-aggregates/data/?frequency=quarterly&data[0]=ash-content&data[1]=heat-content&data[2]=price&data[3]=quantity&data[4]=sulfur-content&start=2011-Q1&sort[0][column]=period&sort[0][direction]=desc&offset=0&length=5000
# - X-Params: {
#    "frequency": "quarterly",
#    "data": [
#        "ash-content",
#        "heat-content",
#        "price",
#        "quantity",
#        "sulfur-content"
#   ],
#    "facets": {},
#    "start": "2011-Q1",
#    "end": null,
#    "sort": [
#       {
#            "column": "period",
#            "direction": "desc"
#        }
#    ],
#    "offset": 0,
#    "length": 5000
#}
