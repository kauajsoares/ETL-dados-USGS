import os
import io
import re
import json
import msal
import time
import zipfile
import requests
import urllib.parse
import pandas as pd
import sciencebasepy
from io import BytesIO
from pathlib import Path
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential


# Carrega as variáveis do arquivo .env para o ambiente da sessão
load_dotenv()


# CARREGA AS VARIÁVEIS A PARTIR DO ARQUIVO .ENV

TENANT_ID = os.getenv("TENANT_ID") 
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
USERNAME = os.getenv("username_microsoft")
PASSWORD = os.getenv("password_microsoft")
SHAREPOINT_URL = os.getenv("SHAREPOINT_URL_RAIZ")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE")
SHAREPOINT_PASTA = os.getenv("SHAREPOINT_PASTA")
SHAREPOINT_DOC = os.getenv("SHAREPOINT_DOC") 


def get_acesstoken():
    AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
    SCOPE = [f"{SHAREPOINT_URL}/.default"]

    # Usa ConfidentialClientApplication porque estamos usando um client secret
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY
        )

    # Adquire o token usando o fluxo de nome de usuário/senha (ROPC)
    result = app.acquire_token_by_username_password(
        username=USERNAME,
        password=PASSWORD,
        scopes=SCOPE
        )

    if "access_token" in result:
        access_token = result['access_token']
        print("\nToken de acesso gerado com sucesso!")
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json;odata=verbose'
        }
        # print("\nExemplo de cabeçalho (headers) para uma chamada de API:")
        # print(headers)
        return access_token
    else:
        print("\n❌ Erro ao obter o token:")
        print("Erro:", result.get("error"))
        print("Descrição:", result.get("error_description"))

token = get_acesstoken()

def upload_excel_para_sharepoint(access_token: str, sharepoint_url: str, sharepoint_site: str, pasta_destino: str, nome_arquivo: str, df: pd.DataFrame):

    if not all([access_token, sharepoint_url, sharepoint_site, pasta_destino, nome_arquivo]):
        print("Todos os parâmetros (token, URL, subpasta do site, pasta de destino, nome do arquivo) são obrigatórios.")
        return
    if df.empty:
        print("O DataFrame fornecido está vazio.")
        return

    print("\n" + "="*50)
    print("INICIANDO UPLOAD DE ARQUIVO PARA O SHAREPOINT")
    print("="*50)

    # Constrói a URL da API para o upload.
    api_url = (
        f"{SHAREPOINT_URL}{SHAREPOINT_SITE}/_api/web/"
        f"GetFolderByServerRelativeUrl('{SHAREPOINT_DOC}/{SHAREPOINT_PASTA}')/Files/"
        f"add(url='{nome_arquivo}',overwrite=true)")

    # Converte o DataFrame para um arquivo Excel em um buffer de bytes na memória.
    try:
        excel_buffer = io.BytesIO()
        df.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_buffer.seek(0)  # Volta ao início do buffer para a leitura
        file_content = excel_buffer.read()
    except Exception as e:
        print(f"❌ Erro ao converter o DataFrame para Excel: {e}")
        return

    # Cabeçalhos da requisição
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/octet-stream' # Indica que estamos enviando dados binários
    }

    print(f"▶️  Enviando arquivo '{nome_arquivo}' para a pasta '{pasta_destino}'...")
    print(f"▶️  URL da API: {api_url}")

    try:
        # Faz a requisição POST com o conteúdo do arquivo no corpo (data)
        response = requests.post(api_url, headers=headers, data=file_content)

        # Verifica a resposta
        if response.status_code == 200 or response.status_code == 201:
            print(f"✅ Arquivo '{nome_arquivo}' enviado com sucesso! Status Code: {response.status_code}")
            # Extrai a URL do arquivo recém-criado da resposta
            file_url = response.json().get('d', {}).get('ServerRelativeUrl')
            if file_url:
                print(f"   URL do arquivo: {sharepoint_url}{file_url}")
        else:
            print(f"❌ Falha no upload do arquivo. Status Code: {response.status_code}")
            print("   Resposta recebida do servidor:")
            try:
                print(json.dumps(response.json(), indent=2))
            except json.JSONDecodeError:
                print(response.text)

    except requests.exceptions.RequestException as e:
        print(f"❌ Ocorreu um erro de conexão ao tentar fazer o upload: {e}")

        # IDs dos releases ScienceBase
sciencebase_ids = {
    2022 : '6197ccbed34eb622f692ee1c', 
    2023 : '63b5f411d34e92aad3caa57f',
    2024 : '65a6e45fd34e5af967a46749', 
    2025 : '677eaf95d34e760b392c4970'  
}


# Lista fixa de arquivos CSV desejados
arquivos_desejados = [
    "mcs2022-alumi_world.csv", "mcs2022-cobal_world.csv", "mcs2022-coppe_world.csv",
    "mcs2022-tin_world.csv", "mcs2022-graph_world.csv", "mcs2022-lithi_world.csv",
    "mcs2022-manga_world.csv", "mcs2022-niobi_world.csv","mcs2022-nicke_world.csv", "mcs2022-raree_world.csv",
    "mcs2022-vanad_world.csv", "mcs2022-zinc_world.csv", "mcs2022-simet_world.csv",

    "mcs2023-alumi_world.csv", "mcs2023-cobal_world.csv", "mcs2023-coppe_world.csv",
    "mcs2023-tin_world.csv", "mcs2023-graph_world.csv", "mcs2023-lithi_world.csv",
    "mcs2023-manga_world.csv", "mcs2023-niobi_world.csv", "mcs2023-nicke_world.csv", "mcs2023-raree_world.csv",
    "mcs2023-vanad_world.csv", "mcs2023-zinc_world.csv", "mcs2023-simet_world.csv",

    "mcs2024-alumi_world.csv", "mcs2024-cobal_world.csv", "mcs2024-coppe_world.csv",
    "mcs2024-tin_world.csv", "mcs2024-graph_world.csv", "mcs2024-lithi_world.csv",
    "mcs2024-manga_world.csv", "mcs2024-niobi_world.csv", "mcs2024-nicke_world.csv", "mcs2024-raree_world.csv",
    "mcs2024-vanad_world.csv", "mcs2024-zinc_world.csv", "mcs2024-simet_world.csv",

    #Metais Nao Ferrosos(13 de cima mais os 5 de baixo Niobio nao entra em Metais Ñ Ferrosos)
    "mcs2022-chrom_world.csv","mcs2022-lead_world.csv","mcs2022-mgmet_world.csv", "mcs2022-molyb_world.csv", "mcs2022-timin_world.csv",
    "mcs2023-chrom_world.csv","mcs2023-lead_world.csv","mcs2023-mgmet_world.csv", "mcs2023-molyb_world.csv", "mcs2023-timin_world.csv",
    "mcs2024-chrom_world.csv","mcs2024-lead_world.csv","mcs2024-mgmet_world.csv", "mcs2024-molyb_world.csv", "mcs2024-timin_world.csv"

]

mapa_paises = {
    "argentina": "Argentina",
    "australia": "Austrália",
    "austria": "Áustria",
    "bahrain": "Bahrein",
    "bhutan": "Butão",
    "bhutan9": "Butão",
    "bolivia": "Bolívia",
    "brazil": "Brasil",
    "burma": "Mianmar",
    "burundi": "Burundi",
    "canada": "Canadá",
    "chile": "Chile",
    "china": "China",
    "congo (kinshasa)": "República Democrática do Congo",
    "cuba": "Cuba",
    "côte d’ivoire": "Costa do Marfim",
    "finland": "Finlândia",
    "france": "França",
    "gabon": "Gabão",
    "georgia": "Geórgia",
    "germany": "Alemanha",
    "ghana": "Gana",
    "greenland": "Groenlândia",
    "iceland": "Islândia",
    "india": "Índia",
    "india9": "Índia",
    "indonesia": "Indonésia",
    "iran": "Irã",
    "japan": "Japão",
    "kazakhstan": "Cazaquistão",
    "kazakhstan, concentrate": "Cazaquistão (concentrado)",
    "kenya": "Quênia",
    "korea, north": "Coreia do Norte",
    "korea, republic of": "Coreia do Sul",
    "laos": "Laos",
    "madagascar": "Madagáscar",
    "malaysia": "Malásia",
    "malaysia9": "Malásia",
    "mexico": "México",
    "morocco": "Marrocos",
    "mozambique": "Moçambique",
    "new caledonia": "Nova Caledônia",
    "new caledonia11": "Nova Caledônia",
    "new caledonia9": "Nova Caledônia",
    "new caledonia (overseas territory of france)": "Nova Caledônia",
    "nigeria": "Nigéria",
    "norway": "Noruega",
    "other countries": "Outros países",
    "other countries ": "Outros países",
    "other countries ": "Outros países",
    "other countries6": "Outros países",
    "papua new guinea": "Papua-Nova Guiné",
    "peru": "Peru",
    "philippines": "Filipinas",
    "poland": "Polônia",
    "portugal": "Portugal",
    "russia": "Rússia",
    "rwanda": "Ruanda",
    "sierra leone": "Serra Leoa",
    "south africa": "África do Sul",
    "spain": "Espanha",
    "sri lanka": "Sri Lanka",
    "sweden": "Suécia",
    "tajikistan": "Tajiquistão",
    "tanzania": "Tanzânia",
    "thailand": "Tailândia",
    "turkey": "Turquia",
    "ukraine": "Ucrânia",
    "ukraine9": "Ucrânia",
    "ukraine, concentrate": "Ucrânia (concentrado)",
    "united arab emirates": "Emirados Árabes Unidos",
    "united states": "Estados Unidos",
    "united states ": "Estados Unidos",
    "uzbekistan": "Uzbequistão",
    "vietnam": "Vietnã",
    "vietnam ": "Vietnã",
    "world total (rounded)": "Total mundial (arredondado)",
    "world total (rounded), excluding u.s. production": "Total mundial (arredondado, excluindo produção EUA)",
    "zambia": "Zâmbia",
    "zimbabwe": "Zimbábue",
    "null": 'null',
    "namibia" : "Namíbia",
    "united states": "Estados Unidos",  # com espaço não padrão
}

mapa_commodities = {
    "alumi": "Alumínio",
    "chrom": "Cromo",
    "cobal": "Cobalto",
    "coppe": "Cobre",
    "tin": "Estanho",
    "graph": "Grafite",
    "lead": "Chumbo",
    "lithi": "Lítio",
    "mgmet": "Magnésio",
    "manga": "Manganês",
    "molyb": "Molibdênio",
    "niobi": "Nióbio",
    "nicke": "Níquel",
    "simet": "Silício",
    "raree": "Terras Raras",
    "timin": "Titânio", 
    "vanad": "Vanádio",
    "zinc": "Zinco",

}
mapa_commodities_2025 = {
#Commodities nos arquivos de 2025
    "Aluminum": "Alumínio", 
    "Chromium" : "Cromo", 
    "Lead" : "Chumbo" ,
    "Cobalt": "Cobalto",
    "Copper ": "Cobre",
    "Tin": "Estanho",
    "Graphite": "Grafite",
    "Lithium ": "Lítio",
    "Magnesium metal" : "Magnésio",
    "Manganese": "Manganês", 
    "Molybdenum " : "Molibdênio",
    "Niobium": "Nióbio",
    "Nickel": "Níquel",
    "Silicon": "Silício",
    "Rare earths": "Terras Raras", 
    "Titanium Mineral Concentrates" : "Titânio",
    "Vanadium": "Vanádio",
    "Zinc": "Zinco"
}

commodities_validas_en = [ #Espaço a mais em Copper, Lithium e Molybdenum no arquivo original
    "Aluminum", "Lead", "Chromium", "Cobalt", "Copper ", "Tin", "Graphite", "Lithium ","Magnesium metal" 
    ,"Manganese", "Molybdenum ", "Niobium", "Nickel", "Silicon", "Rare earths", "Titanium Mineral Concentrates", "Vanadium", "Zinc"
] #lista para selecionar as commodities no arquivo de 2025 

#Caso a parte em aluminio
tipo_por_commodity = {
    "alumi": "smelter production",
}

def processar_zip(zip_bytes, arquivos_desejados, fonte_label,):
    dfs = []

    with zipfile.ZipFile(zip_bytes) as z:
        for file_name in z.namelist():
            nome_base = file_name.split("/")[-1].lower()
            if nome_base in arquivos_desejados:
                with z.open(file_name) as f:
                    df = pd.read_csv(f, sep=",", thousands=",", quotechar='"')

                    # Normalizar nomes de países
                    df['Country'] = df['Country'].str.lower().map(mapa_paises).fillna(df['Country'])

                    # Verificar coluna Type
                    tipos_unicos = df['Type'].dropna().unique()
                    commodity_abrev = re.search(r"mcs\d{4}-(.*?)_world\.csv", nome_base)
                    commodity_abrev = commodity_abrev.group(1) if commodity_abrev else 'unknown'

                    if len(tipos_unicos) > 1:
                        if commodity_abrev in tipo_por_commodity:
                            tipo_desejado = tipo_por_commodity[commodity_abrev]
                            df = df[df['Type'] == tipo_desejado]
                            print(f"Arquivo {file_name}: múltiplos tipos encontrados, filtrando pelo tipo '{tipo_desejado}'")
                        else:
                            print(f"Atenção: arquivo {file_name} tem múltiplos tipos e nenhum tipo definido no mapa: {tipos_unicos}")
                            
                    elif len(tipos_unicos) == 1:
                        pass  # apenas um tipo → mantém todas as linhas
                    else:
                        pass  # nenhum tipo → mantém todas as linhas

                    # Selecionar apenas colunas válidas (sem *_notes)
                    colunas_producao = [
                        col for col in df.columns 
                        if re.match(r'prod_(?:t|kt)_\d{4}$', col, re.IGNORECASE) and "_notes" not in col.lower()
                    ]

                    # Se não houver colunas reais, tenta estimadas no mesmo formato
                    if not colunas_producao:
                        colunas_producao = [
                            col for col in df.columns 
                            if re.match(r'prod_(?:t|kt)_est_\d{4}$', col, re.IGNORECASE) and "_notes" not in col.lower()
                        ]
                        if colunas_producao:
                            print(f"Atenção: usando produção ESTIMADA no arquivo {file_name}")
                        else:
                            print(f"Nenhuma coluna de produção encontrada em {file_name}, pulando...")
                            continue

                    # Limpeza e conversão
                    for col in colunas_producao:
                        df[col] = (
                            df[col]
                            .astype(str)
                            .str.replace(r"[^\d\.\-]", "", regex=True)
                        )
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                
                        # Converter kt -> toneladas
                        if '_kt_' in col.lower():
                            df[col] = df[col] * 1000

                    # Reorganizar dataframe
                    df = df[['Country'] + colunas_producao]
                    df_long = df.melt(id_vars=['Country'], var_name='Variavel', value_name='Valor')

                    # Extrair ano dinamicamente
                    df_long['Ano'] = df_long['Variavel'].str.extract(r'(\d{4})')
                    df_long = df_long.dropna(subset=['Ano'])
                    df_long['Ano'] = df_long['Ano'].astype(int)

                    #Filtrar pelos anos desejados
                    df_long = df_long[df_long['Ano'].isin([2020, 2021, 2022])]

                    # Substituir NaN por None (equivalente a NULL)
                    df_long['Valor'] = df_long['Valor'].where(pd.notna(df_long['Valor']), None)

                    #remove a linha apenas se pais E valor estiverem nulo
                    df_long = df_long.dropna(subset=['Country', 'Valor'], how='all')

                    # Identificar commodity pelo nome do arquivo
                    df_long['Commodity'] = mapa_commodities.get(commodity_abrev, commodity_abrev)

                    # Remover coluna Variavel, deixar limpo
                    df_long = df_long.drop(columns=['Variavel'])

                    dfs.append(df_long)


                        

    return dfs



sb = sciencebasepy.SbSession()
todos_dfs = []

for item_id in sciencebase_ids.values():
    item = sb.get_item(item_id)
    fonte_label = item['title'][:20].replace(" ", "_").lower()
    print(f"\nPROCESSANDO RELEASE: {item['title']}")

    if 'files' in item:
        for file in item['files']:
            file_name = file['name'].lower()
            if file_name.endswith('.zip') and 'world' in file_name:
                print(f"Baixando ZIP: {file_name}")
                response = requests.get(file['url'])
                response.raise_for_status()
                zip_bytes = BytesIO(response.content)

                dfs = processar_zip(zip_bytes, arquivos_desejados, fonte_label)
                # display(dfs)
                todos_dfs.extend(dfs)

# Concatenar todos os DataFrames
tabela_final = pd.concat(todos_dfs, ignore_index=True)

# Visualizar os dados
# display(tabela_final)

dfs_2025 = []

release_2025_id = sciencebase_ids.get(2025)
item_2025 = sb.get_item(release_2025_id)
print(f"\nPROCESSANDO RELEASE 2025: {item_2025['title']}")

if 'files' in item_2025:
    for file in item_2025['files']:
        file_name = file['name'].lower()
        if file_name.endswith('.zip') and 'world' in file_name:
            print(f"Baixando ZIP: {file_name}")
            response = requests.get(file['url'])
            response.raise_for_status()
            zip_bytes = BytesIO(response.content)

            with zipfile.ZipFile(zip_bytes) as z:
                for file_in_zip in z.namelist():
                    if file_in_zip.endswith(".csv"):
                       # print(f"Lendo arquivo: {file_in_zip}")
                        with z.open(file_in_zip) as f:
                            df = pd.read_csv(f)

                            if 'UNIT_MEAS' not in df.columns:
                                print("Coluna UNIT_MEAS não encontrada.")
                                continue

                            colunas_producao = [col for col in df.columns if col.startswith("PROD_")]

                            if not colunas_producao:
                                print("Nenhuma coluna de produção detectada.")
                                continue

                            df = df[["COUNTRY", "COMMODITY", "UNIT_MEAS"] + colunas_producao]

                            df_long = df.melt(
                                id_vars=["COUNTRY", "COMMODITY", "UNIT_MEAS"],
                                value_vars=colunas_producao,
                                var_name="Ano",
                                value_name="Valor"
                            )

                            df_long["Ano"] = df_long["Ano"].str.extract(r"(\d{4})")
                            df_long = df_long.dropna(subset=["Ano"])
                            df_long["Ano"] = df_long["Ano"].astype(int)

                            df_long = df_long[df_long["Ano"].isin([2023, 2024])]

                            df_long.rename(columns={
                                "COUNTRY": "Country",
                                "COMMODITY": "Commodity"
                            }, inplace=True)

                            df_long["Country"] = df_long["Country"].str.strip().str.lower().map(mapa_paises).fillna(df_long["Country"])
                            df_long = df_long[df_long["Commodity"].isin(commodities_validas_en)]
                            df_long["Commodity"] = df_long["Commodity"].map(mapa_commodities_2025)

                            # Normalizar e converter de thousand para metric tons
                            df_long["UNIT_MEAS"] = df_long["UNIT_MEAS"].str.strip().str.lower()

                            cond = df_long["UNIT_MEAS"] == "thousand metric tons"
                            df_long.loc[cond, "Valor"] = df_long.loc[cond, "Valor"] * 1000
                            df_long.loc[cond, "UNIT_MEAS"] = "metric tons"

                            dfs_2025.append(df_long)

# Exibir resultado
if dfs_2025:
    tabela_2025 = pd.concat(dfs_2025, ignore_index=True)
    # display(tabela_2025[["Country", "Commodity", "Ano", "Valor"]])
else:
    print("Nenhum dado foi processado para 2025.")

# Concatenar as duas tabelas (de releases antigos e 2025)
tabela_completa = pd.concat([tabela_final, tabela_2025], ignore_index=True)

#dispensa a coluna de unidade de medida pq ja foi tudo convertido 
if 'UNIT_MEAS' in tabela_completa.columns:
    tabela_completa = tabela_completa.drop(columns=['UNIT_MEAS'])


nome_arquivo = "TESTE2ProdUSGS20-24.xlsx"
# display(tabela_completa)

token = get_acesstoken()
if token:
    upload_excel_para_sharepoint(
        access_token=token,
        sharepoint_url=SHAREPOINT_URL,
        sharepoint_site=SHAREPOINT_SITE,
        pasta_destino=SHAREPOINT_PASTA,
        nome_arquivo=nome_arquivo,
        df=tabela_completa
    )
else: 
    print('Codigo de autenticação faltando')