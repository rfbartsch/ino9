#!/usr/bin/env python
# coding: utf-8

# In[1]:


#CONSOLIDADO
import pyautogui
import time
import os
import datetime
import pandas as pd
import openpyxl
import re
import shutil
import locale
import glob
import win32com.client
import email
import numpy as np
import keyboard
import pytz
import requests
import numpy
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import FORMULAE

from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


# Marque o tempo de início
start_time = time.time()
print("Início Custodia OnShore")


# Configurar locale para usar separador de milhar "."
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
tz = pytz.timezone('America/Sao_Paulo')

# Set up constants

PASTA_DOWNLOADS = r'C:\Users\dados.100486\Downloads'
PASTA_DIVERSIFICACAO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Tech Inove\Base Dados\Diversificacao'




PASTA_MI_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Mercado Internacional\Historico'
PASTA_RF = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Vencimento RF Connect'
PASTA_DIVERSIFICACAO_ATUAL = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao\Historico'
PASTA_VENC_RF_COMPLETO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Vencimento RF Connect\Completo'

arquivo_guia_fundos = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões\dGuiaConsol.xlsx'
arquivo_bd_fundos = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões\dBD.xlsx'


PASTA_ARQUIVOS_PADROES = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões'

PASTA_CAMPANHA_ATIVACAO = r'C:\Users\Douglas\XP Investimentos\Bruno Martins - Power BI\Base Dados\Positivador\Campanha Ativação'

PASTA_DIVERSIFICACAO_HISTORICO = r'C:\Users\Douglas\XP Investimentos\Bruno Martins - Power BI\Base Dados\Diversificacao\Historico Consolidado'


PAUSE_TIME = 2
SHORT_PAUSE_TIME = 1


# Set up PyAutoGUI failsafe
pyautogui.FAILSAFE = True

# Define helper functions
def mover_mouse(x, y):
    pyautogui.moveTo(x, y)

def clicar_pagina(x, y, pause=PAUSE_TIME):
    mover_mouse(x, y)
    time.sleep(pause)
    pyautogui.click()

def clicar_pagina_3x(x, y, pause=PAUSE_TIME):
    mover_mouse(x, y)
    time.sleep(pause)
    pyautogui.tripleClick()
    
def format_string_percentage(value):
    try:
        number = float(value)
        return f"{number:.2%}"
    except ValueError:
        return value

def extract_indexador(row):
    value = row['Indexador/Benchmark']
    if isinstance(value, str):
        if '% CDI' in value:
            return '% CDI'
        elif '% IPC-A' in value:
            return '% IPC-A'
        elif '% IGP-M' in value:
            return '% IGP-M'
        elif 'CDI +' in value:
            return 'CDI +'
        elif 'DOLAR PTAX +' in value:
            return 'DOLAR PTAX +'
        elif 'IGP-M +' in value:
            return 'IGP-M +'
        elif 'IPC-A +' in value:
            return 'IPC-A +'
        elif 'LFT +' in value:
            return 'LFT +'
        elif 'LFT' in value:
            return 'LFT'
    else:
        return 'Pré'
    
def extract_taxa(row):
    value = row['Indexador/Benchmark']
    if isinstance(value, str):
        if '% CDI' in value:
            return value.split()[0]
        elif '% IPC-A' in value:
            return value.split()[0]
        elif '% IGP-M' in value:
            return value.split()[0]
        elif 'CDI +' in value:
            return value.split('+')[1].strip()
        elif 'DOLAR PTAX +' in value:
            return value.split('DOLAR PTAX +')[1].strip()
        elif 'IGP-M +' in value:
            return value.split('+')[1].strip()
        elif 'IPC-A +' in value:
            return value.split('+')[1].strip()
        elif 'LFT +' in value:
            return value.split('+')[1].strip()
        elif 'LFT' in value:
            return '0%'
    else:
        return value
    
# definir uma função que aplica as transformações necessárias na coluna Cod_Ativo
def transforma_cod_ativo(cod_ativo):
    if cod_ativo.startswith('PRE DU'):
        # extrair o conteúdo que vem após "PRE DU"
        novo_cod_ativo = re.search(r'PRE DU (.*)', cod_ativo).group(1)
    elif cod_ativo.startswith('PRE'):
        # extrair o conteúdo que vem após "PRE"
        novo_cod_ativo = re.search(r'PRE (.*)', cod_ativo).group(1)
    elif cod_ativo.startswith('FLU'):
        # extrair o conteúdo que vem após "FLU U"
        novo_cod_ativo = re.search(r'FLU U (.*)|FLU (.*)', cod_ativo)
        novo_cod_ativo = novo_cod_ativo.group(1) if novo_cod_ativo.group(1) else novo_cod_ativo.group(2)
    else:
        # manter o valor original
        novo_cod_ativo = cod_ativo
    return novo_cod_ativo

def aplicar_formula(coluna):
    if coluna == "" or coluna == "Total":
        return ""
    elif coluna in ["Bond", "Unknown", "Equity", "Stock", "Cash", "USD", "REIT", "ETF", "EUR", "Alternative", "Balanced", "Semi-liquid"]:
        return df.loc[df['Top Level Owner'] == coluna, 'Valor Total'].values[0]
    else:
        try:
            pos_espaco = coluna.index(" ")
            return df.loc[df['Top Level Owner'] == coluna, 'Valor Total'].values[0]
        except ValueError:
            return coluna

        
        
#####


# Get the current date and time in the desired format
formatted_date = datetime.now().strftime("%Y%m%d_%H%M%S")
formatted_date2 = datetime.now().strftime("%Y%m")
time.sleep(5)


# Lista todos os arquivos na pasta e filtra aqueles que começam com "Diversificacao_"
arquivos_filtrados = [arquivo for arquivo in os.listdir(PASTA_DIVERSIFICACAO_ATUAL) if arquivo.startswith("Diversificacao_")]

# Encontra o arquivo mais recente usando a função max() com uma chave personalizada para comparar o tempo de modificação dos arquivos
filename = max(arquivos_filtrados, key=lambda x: os.path.getmtime(os.path.join(PASTA_DIVERSIFICACAO_ATUAL, x)))

date_str = filename.split("_")[1]
date_str = date_str.split("_")[0]  
date_formatted = datetime.strptime(date_str, "%Y%m%d")
date_formatted_str = date_formatted.strftime("%d/%m/%Y")
# Load the file into a dataframe
filepath = os.path.join(PASTA_DIVERSIFICACAO_ATUAL, filename)
df = pd.read_excel(filepath)

df = df.reset_index(drop=True)

df['Data_Ref'] = df['Data_Ref2']

# Exclui o produto Renda Fixa, exceto Produto Estruturado (COE)
mask = (df['Produto'] == 'Renda Fixa') & (df['Sub Produto'] != 'Produto Estruturado')

# remove as linhas que atendem a condição da máscara
df = df.drop(df[mask].index)

# Selecionar as últimas 6 letras da coluna "ATIVO" e atribuir à coluna "COD_ATIVO_FIP"
df['COD_ATIVO_FIP'] = df['Ativo'].str[-6:]

# Criar a coluna "COD_ATIVO_2" com a função desejada
df['COD_ATIVO_2'] = df.apply(lambda x: x['COD_ATIVO_FIP'] if (x['Produto'] == 'Fundos' and x['COD_ATIVO_FIP'].endswith(('11', '12'))) else '', axis=1)

# Criar a coluna "Novo_ativo" com a função desejada
df['Novo_ativo'] = df.apply(lambda x: x['COD_ATIVO_2'] if x['COD_ATIVO_2'] == x['COD_ATIVO_FIP'] else x['Ativo'], axis=1)

# Criar as colunas "COD_ATIVO" e "COE" com a função desejada
df['COD_ATIVO'] = df.apply(lambda x: x['Novo_ativo'].split(' ', 1)[1] if x['Produto'] == 'Renda Fixa' else '', axis=1)

# Substituir os valores na coluna "Produto"
df['Produto'] = df['Produto'].replace({'Renda Fixa': 'Alternativo'})

# Substituir os valores na coluna "Sub Produto"
df['Sub Produto'] = df['Sub Produto'].replace({'Produto Estruturado': 'COE'})

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Fundo Imobiliário - Cetipado' if (x['Produto'] == 'Fundos' and x['Novo_ativo'].endswith(('11', '12'))) else x['Sub Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Fundo Imobiliário - Listado' if (x['Sub Produto'] == 'Fundo Imobiliário') else x['Sub Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Produto'] = df.apply(lambda x: 'Fundo Imobiliário' if (x['Sub Produto'] == 'Fundo Imobiliário - Cetipado' and x['Novo_ativo'].endswith(('11', '12'))) else x['Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Produto'] = df.apply(lambda x: 'Fundo Imobiliário' if (x['Sub Produto'] == 'Fundo Imobiliário - Cetipado' or x['Sub Produto'] == 'Fundo Imobiliário - Listado') else x['Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Produto'] = df.apply(lambda x: 'Cripto' if (x['Produto'] == 'OTHER PRODUCTS') else x['Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Tesouro Direto' if (x['Sub Produto'] == 'NÃO ENCONTRADO') else x['Sub Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Tesouro Direto' if (x['Sub Produto'] == 'Título Público') else x['Sub Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Produto'] = df.apply(lambda x: 'Renda Fixa' if (x['Sub Produto'] == 'Tesouro Direto') else x['Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Ativo'] = df.apply(lambda x: 'NTN-B1' if (x['Ativo'] == 'NÃO ENCONTRADO') else x['Ativo'], axis=1)

# Remover as colunas "COD_ATIVO_FIP" e "COD_ATIVO_2"
df = df.drop(['COD_ATIVO_FIP', 'COD_ATIVO_2','Data_Ref2'], axis=1)

df['Onshore/Offshore'] = 'Onshore'

df['Origem'] = filename

df['Assessor'] = df['Assessor'].astype(str)
df['Assessor'] = df['Assessor'].replace('.0','')
df['Cliente'] = df['Cliente'].astype(str)
df['Cliente'] = df['Cliente'].replace('.0','')




# In[2]:


#### Mercado Internacional



PASTA_MI_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Mercado Internacional\Historico'

# Lista todos os arquivos na pasta e filtra aqueles que começam com "Diversificacao_"
arquivos_filtrados = [arquivo for arquivo in os.listdir(PASTA_MI_HISTORICO) if arquivo.startswith("Custodia_")]

# Encontra o arquivo mais recente usando a função max() com uma chave personalizada para comparar o tempo de modificação dos arquivos
filename = max(arquivos_filtrados, key=lambda x: os.path.getmtime(os.path.join(PASTA_MI_HISTORICO, x)))

date_str = filename.split("_")[3]
date_str = date_str.split("_")[0]  
date_formatted = datetime.strptime(date_str, "%Y%m%d")
date_formatted_str = date_formatted.strftime("%d/%m/%Y")
# Load the file into a dataframe
filepath = os.path.join(PASTA_MI_HISTORICO, filename)
df_mi = pd.read_excel(filepath)


df_mi = df_mi.rename(columns={'Date':'Data_Ref',
                        'Código Conta':'Cliente',
                        'Código Assessor':'Assessor',
                        'Descrição Grupo Tipo':'Produto',
                        'Código Ativo':'Novo_ativo',
                        'Descrição Ativo': 'Emissor',
                        'Quantidade do Ativo': 'Quantidade',
                        'Valor de Mercado ($)': 'NET'})



df_mi["Data_Ref"] = date_formatted_str

df_mi['Sub Produto'] = ''
df_mi['Ativo'] = ''
df_mi['Produto em Garantia'] = ''
df_mi['CNPJ Fundo'] = ''
df_mi['Data de Vencimento'] = ''
df_mi['COD_ATIVO'] = ''

# Ou criando uma Series com o mesmo valor para todas as linhas
df_mi['Onshore/Offshore'] = 'Offshore'

# Criar a condição para a coluna "Sub Produto"
df_mi['Produto'] = df_mi.apply(lambda x: 'Renda Fixa' if (x['Produto'] == 'FixedIncome' ) else x['Produto'], axis=1)
df_mi['Produto'] = df_mi.apply(lambda x: 'Somente Financeiro' if (x['Produto'] == 'Cash' ) else x['Produto'], axis=1)
df_mi['Produto'] = df_mi.apply(lambda x: 'Renda Variável' if (x['Produto'] == 'Equity' ) else x['Produto'], axis=1)
df_mi['Sub Produto'] = df_mi.apply(lambda x: 'T-Note' if (x['Emissor'] == 'UNITED STATES TREASURY NOTE' ) else x['Sub Produto'], axis=1)
df_mi['Sub Produto'] = df_mi.apply(lambda x: 'Saldo em Conta' if (x['Produto'] == 'Somente Financeiro' ) else x['Sub Produto'], axis=1)
df_mi['Sub Produto'] = df_mi.apply(lambda x: 'Bond' if (x['Produto'] == 'Renda Fixa' and x['Emissor'] != 'UNITED STATES TREASURY NOTE') else x['Sub Produto'], axis=1)
df_mi['Sub Produto'] = df_mi.apply(lambda x: 'Equity' if x['Produto'] == 'Renda Variável' else x['Sub Produto'], axis=1)
df_mi['COD_ATIVO'] = df_mi.apply(lambda x: '' if x['Produto'] != 'Renda Fixa' else x['Novo_ativo'], axis=1)
df_mi['Emissor'] = df_mi.apply(lambda x: '' if x['Emissor'] == 'Account Balance' else x['Emissor'], axis=1)
df_mi['Novo_ativo'] = df_mi.apply(lambda x: '' if x['Novo_ativo'] == 'CASH' else x['Novo_ativo'], axis=1)
df_mi['Origem'] = filename
df_mi = df_mi[['Assessor', 'Cliente', 'Produto', 'Sub Produto', 'Produto em Garantia', 'CNPJ Fundo', 'Ativo', 'Emissor', 'Data de Vencimento', 'Quantidade', 'NET', 'Data_Ref', 'Novo_ativo', 'COD_ATIVO','Onshore/Offshore','Origem']]
# Excluir o primeiro caractere 'A' das linhas da coluna 'Assessor'
df_mi['Assessor'] = df_mi['Assessor'].str.slice(1)
df_mi['Cliente'] = df_mi['Cliente'].astype(str)
df_mi['Cliente'] = df_mi['Cliente'].replace('.0','')

df = pd.concat([df, df_mi], ignore_index=True)





# In[3]:


## GUIA

arquivo_guia_fundos = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões\dGuiaConsol.xlsx'
arquivo_bd_fundos = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões\dBD.xlsx'

# Load the file into a new dataframe
df_guia_fundos_555 = pd.read_excel(os.path.join(arquivo_guia_fundos))

#df_guia_fundos_555 = pd.concat([df_guia_fundos_555, df_bd_fundos], ignore_index=True)
df_guia_fundos_555 = df_guia_fundos_555.dropna(subset=['Fundo'])

# preencher cada coluna vazia com o valor 0
df_guia_fundos_555['Taxa de Administração.1'].fillna(0, inplace=True)
df_guia_fundos_555['Taxa de Performance'].fillna(0, inplace=True)
df_guia_fundos_555['Taxa de Administração x Comissão'].fillna(0, inplace=True)
df_guia_fundos_555['Taxa de Performance x Comissão***'].fillna(0, inplace=True)
df_guia_fundos_555['Margem Total (Margem Taxa Adm + Margem Taxa Perf)'].fillna(0, inplace=True)

df_guia_fundos_555['CNPJ'] = df_guia_fundos_555['CNPJ'].astype(str)

df['CNPJ Fundo'] = df['CNPJ Fundo'].astype(str).apply(lambda x: x.split('.')[0])



# Merge the DataFrames by "Cliente" and "Conta XP" columns, keeping the values from "df"
df = pd.merge(df, df_guia_fundos_555[['Margem Total (Margem Taxa Adm + Margem Taxa Perf)',
                                      'Taxa de Performance x Comissão***',
                                      'Taxa de Administração x Comissão',
                                      'Taxa de Performance',
                                      'Taxa de Administração',
                                      'CVM',
                                      'Anbima',
                                      'Classificação XP',
                                      'CNPJ',
                                      'Gestor',
                                      'Administrador',
                                      'Custodiante',
                                      'Benchmark',
                                      'Liquidez total (Cotização + Liquidação) em número (aproximação para filtros)']],
              left_on='CNPJ Fundo', right_on='CNPJ', how='left')

df = df.drop(["CNPJ"], axis=1)


cols_to_replace = [
    'Taxa de Administração', 'Taxa de Performance', 'Taxa de Administração x Comissão',
    'Taxa de Performance x Comissão***', 'Margem Total (Margem Taxa Adm + Margem Taxa Perf)'
]

df[cols_to_replace] = df[cols_to_replace].replace('-', 0) 
    
    
cols_to_format = [
    'Taxa de Administração', 'Taxa de Performance', 'Taxa de Administração x Comissão',
    'Taxa de Performance x Comissão***', 'Margem Total (Margem Taxa Adm + Margem Taxa Perf)'
]

df[cols_to_format] = df[cols_to_format].applymap(format_string_percentage)

df['Classificação XP'] = df['Classificação XP'].fillna('-')

df_bd_class_inove_fundos = pd.read_excel(os.path.join(arquivo_bd_fundos),sheet_name='Ajuste Classe Para Inove')

df = pd.merge(df, df_bd_class_inove_fundos, left_on='Classificação XP', right_on='Classificacao XP', how='left')

df["TD"] = df["Ativo"].str.split().str[0]



#################### CLASS INOVE

# Aplica as regras definidas para criar a coluna "Class_Inove2"
df["Class_Inove2"] = np.where(df["Onshore/Offshore"] == "Offshore", "Internacional",
                        np.where(df["Produto"] == "Renda Variável", "Renda Variável",
                        np.where(df["Produto"] == "Somente Financeiro", "Liquidez",
                        np.where(df["Produto"] == "Alternativo", "Alternativo",
                        np.where(df["Produto"] == "Cripto", "Alternativo",
                        np.where(df["Produto"] == "Fundo Imobiliário", "Fundo Imobiliário",
                        np.where(df["Liquidez total (Cotização + Liquidação) em número (aproximação para filtros)"] == 0, "Liquidez",        
                        np.where((df["Sub Produto"] == "Tesouro Direto") & (df["TD"] == "LTN"), "Pré-Fixado",
                        np.where((df["Sub Produto"] == "Tesouro Direto") & (df["TD"] == "NTN-B"), "Inflação",
                        np.where((df["Sub Produto"] == "Tesouro Direto") & (df["TD"] == "NTN-B1"), "Inflação",
                        np.where((df["Sub Produto"] == "Tesouro Direto") & (df["TD"] == "NTN-F"), "Pré-Fixado",
                        np.where((df["Sub Produto"] == "Tesouro Direto") & (df["TD"] == "LFT"), "Liquidez", "-"))))))))))))


df['ID'] = df[['Assessor', 'Cliente', 'Produto', 'Sub Produto', 'Ativo', 'Emissor', 'Data de Vencimento']].apply(lambda x: '_'.join(x.astype(str)), axis=1)

df = df.drop_duplicates(subset=['ID'])

df.reset_index()

df["Class_Inove3"] = np.where(df["Class_Inove2"] == "-", df["Class_Inove"], df["Class_Inove2"])

# Encontrar o índice da linha onde a coluna 'NET_Original' contém a palavra 'NET'
linha_net = df.loc[df['NET'] == 'NET'].index

# Remover a linha com a palavra 'NET', se existir
if not linha_net.empty:
    df = df.drop(linha_net)
    
df['Quantidade'] = df['Quantidade'].astype(float)
df['NET'] = df['NET'].astype(float)
df['Dt_Aplicacao'] = ''
df['PU_ATUAL'] = np.where((df['NET'] == 0) | (df['Quantidade'] == 0), 0, df['NET'] / df['Quantidade'])
df['Carencia'] = ''
df['PERPETUO'] = "Não"
df['COD_CLIENTE_XPUS']=''

df = df.drop(["Produto em Garantia","Class_Inove"], axis=1)

df = df.rename(columns={'Assessor':'COD_ASSESSOR',
                        'Cliente':'COD_CLIENTE',
                        'Novo_Ativo':'Ativo',
                        'Data de Vencimento':'Vencimento',
                        'NET':'NET_Original',
                        'Quantidade':'Qtd',
                        'Class_Inove3':'Class_Inove',
                        'Benchmark':'Indexador/Benchmark',
                        'Liquidez total (Cotização + Liquidação) em número (aproximação para filtros)': 'Liquidez'})


df['COD_ASSESSOR'] = df['COD_ASSESSOR'].astype(str)

df['Taxa']=''
df['Indexador']=''


df = df[['Data_Ref',
         'Origem',
         'COD_CLIENTE', 
         'COD_ASSESSOR', 
         'Produto',
         'Sub Produto', 
         'Class_Inove', 
         'Ativo', 
         'COD_ATIVO', 
         'Emissor', 
         'Indexador/Benchmark', 
         'Dt_Aplicacao', 
         'Vencimento', 
         'Carencia', 
         'Qtd', 
         'PU_ATUAL', 
         'NET_Original', 
         'Liquidez',
         'CNPJ Fundo',
         'Margem Total (Margem Taxa Adm + Margem Taxa Perf)',
         'Taxa de Performance x Comissão***',
         'Taxa de Administração x Comissão',
         'Taxa de Performance',
         'Taxa de Administração',
         'CVM',
         'Anbima',
         'Classificação XP',
         'Gestor',
         'Administrador',
         'Custodiante',
         'Onshore/Offshore',
         'PERPETUO',
         'COD_CLIENTE_XPUS',
         'Taxa',
         'Indexador',
         
         
        
        ]]

df_diversificador=df



# In[4]:


PASTA_VENC_RF_COMPLETO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Vencimento RF Connect\Completo'
##############Venc RF Completo
# Lista todos os arquivos na pasta e filtra aqueles que começam com "Diversificacao_"
arquivos_filtrados = [arquivo for arquivo in os.listdir(PASTA_VENC_RF_COMPLETO) if arquivo.startswith("Venc_RF_Completo_")]

# Encontra o arquivo mais recente usando a função max() com uma chave personalizada para comparar o tempo de modificação dos arquivos
filename = max(arquivos_filtrados, key=lambda x: os.path.getmtime(os.path.join(PASTA_VENC_RF_COMPLETO, x)))

# Extrai o valor após o primeiro sublinhado

valor = filename.split('_')[3]
valor = valor.split('_')[0]
date_formatted = datetime.strptime(valor, "%Y%m%d")

filepath = os.path.join(PASTA_VENC_RF_COMPLETO, filename)
df = pd.read_excel(filepath)

# separar o conteúdo da coluna Ativo
df[['Ativo', 'Cod_Ativo']] = df['Ativo'].str.split(' ', 1, expand=True)

# aplicar a função transforma_cod_ativo à coluna Cod_Ativo usando apply()
df['Cod_Ativo'] = df['Cod_Ativo'].apply(transforma_cod_ativo)

df['Data_Ref'] = date_formatted

# remover a coluna financeiro
df = df.drop('Financeiro', axis=1)
df['Qtd'] = df['Qtd'].astype(str).str.replace(',', '.')

# substituir a vírgula por um ponto na coluna Qtd
#df['Qtd'] = df['Qtd'].str.replace(',', '.')
df['PU_Atual'] = df['PU_Atual'].str.replace(',', '.')

# converter as colunas Qtd e PU_Atual para float
df['Qtd'] = df['Qtd'].astype(float)
df['PU_Atual'] = df['PU_Atual'].astype(float)

df['Financeiro'] = df['Qtd']*df['PU_Atual']


# Criar a coluna "Carencia2" com a função desejada
df["Carencia2"] = np.where(df["Ativo"] == "LFT", df["Dt_Aplicacao"],
                           np.where(df["Carencia"].str.contains("/", na=False), df["Carencia"],
                                    df["Vencimento"] + ", " + df["Carencia"]))

df['Carencia2'].fillna("-", inplace=True)

df['Produto'] = "Renda Fixa"

# Criar a condição para a coluna "Sub Produto"
df['Produto'] = df.apply(lambda x: 'Fundos' 
                             if (x['Ativo'] == 'FND') 
                             else x['Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Fundos' 
                             if (x['Ativo'] == 'FND') 
                             else x['Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Ativo'] = df.apply(lambda x: 'Emissor' 
                             if (x['Produto'] == 'Fundos') 
                             else x['Ativo'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Emissor'] = df.apply(lambda x: '' 
                             if (x['Ativo'] == 'FIC FIM XP SPECIAL SITUATIONS I CAPITAL CREDITO PR') 
                             else x['Emissor'], axis=1)

df['Sub Produto'] = ""

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Crédito Privado' 
                             if (x['Ativo'] == 'CRI' 
                                 or x['Ativo'] == 'CRA'
                                 or x['Ativo'] == 'DEB') 
                             else x['Sub Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Emissão Bancária' 
                             if (x['Ativo'] == 'CDB' 
                                 or x['Ativo'] == 'LF'
                                 or x['Ativo'] == 'LC'
                                 or x['Ativo'] == 'LFSN'
                                 or x['Ativo'] == 'LIG'
                                 or x['Ativo'] == 'LCI'
                                 or x['Ativo'] == 'LCA') 
                             else x['Sub Produto'], axis=1)

# Criar a condição para a coluna "Sub Produto"
df['Sub Produto'] = df.apply(lambda x: 'Título Público' 
                             if (x['Ativo'] == 'LTN' 
                                 or x['Ativo'] == 'NTN-F'
                                 or x['Ativo'] == 'LFT'
                                 or x['Ativo'] == 'NTN-B'
                                 or x['Ativo'] == 'NTN-C') 
                             else x['Sub Produto'], axis=1)


df['Carencia2'] = pd.to_datetime(df['Carencia2'], errors='coerce')
df['Liquidez'] = (pd.to_datetime(df['Carencia2'], errors='coerce') - pd.Timestamp.now()).apply(lambda x: x.days if x.days >= 0 else 0)
df["Class_Inove"] = np.where(df["Financeiro"] == 0, "Default",
                       np.where(df["Ativo"].isin(["NTN-F", "LTN"]), "Pré-Fixado",
                       np.where(df["Ativo"].isin(["NTN-B1","NTN-B", "NTN-C"]), "Inflação",
                       np.where(df["Ativo"] == "LFT", "Liquidez",
                       np.where(df["Produto"] == "Fundos", "Alternativo",
                       np.where(df["Indexador"].str.contains("CDI"), "Pós-Fixado",
                       np.where(df["Indexador"].str.contains("IPCA|IPC-A|IGP-M|IGPM"), "Inflação",
                       "Pré-Fixado")))))))

df['Carencia2'] = pd.to_datetime(df['Carencia2'], format='%d/%m/%Y')
df = df.drop(["AC","Carencia"], axis=1)
df = df.rename(columns={'PU_Atual':'PU_ATUAL',
                        'Financeiro':'NET_Original',
                        'Cod_Ativo':'COD_ATIVO',
                        'Carencia2':'Carencia',
                        'Cliente':'COD_CLIENTE',
                        'Assessor': 'COD_ASSESSOR',
                        'Indexador': 'Indexador/Benchmark',
                        'Liquidez total (Cotização + Liquidação) em número (aproximação para filtros)': 'Liquidez'})

df['CNPJ Fundo'] = ''
df['Margem Total (Margem Taxa Adm + Margem Taxa Perf)'] = 0
df['Taxa de Performance x Comissão***'] = 0
df['Taxa de Administração x Comissão'] = 0
df['Taxa de Performance'] = 0
df['Taxa de Administração'] = 0
df['CVM'] = ''
df['Anbima'] = ''
df['Classificação XP'] = ''
df['Gestor'] = ''
df['Administrador'] = ''
df['Custodiante'] = ''


df['Origem'] = filename
df['Onshore/Offshore'] = 'Onshore'
df['PERPETUO'] = "Não"
df['COD_CLIENTE_XPUS'] = ''
df['COD_ASSESSOR'] = df['COD_ASSESSOR'].str.slice(start=1)


df['Vencimento'] = pd.to_datetime(df['Vencimento'], format='%d/%m/%Y')
df['Taxa'] = ''
df['Taxa'] = df['Indexador/Benchmark']
df['Taxa'] = df['Taxa'].str.replace(' %', '')
df['Indexador'] = ''
df['Taxa'] = df.apply(extract_taxa, axis=1)
df['Taxa'] = df.apply(lambda x: x['Indexador/Benchmark'] if (x['Produto'] == 'Renda Fixa' and x['Class_Inove'] == 'Pré-Fixado') else x['Taxa'], axis=1)
df['Taxa'] = df.apply(lambda row: row['Indexador/Benchmark'].split(' +')[1].strip() 
                      if 'DOLAR PTAX +' in str(row['Indexador/Benchmark']) 
                      else row['Taxa'], axis=1)
df['Taxa'] = df['Taxa'].str.replace(' %', '')
df['Taxa'] = df['Taxa'].str.replace('%', '')
df['Taxa'] = df['Taxa'].replace('', '0')
df['Taxa'] = df['Taxa'].astype(float)
df['Indexador'] = df.apply(extract_indexador, axis=1)
df['Indexador'] = df['Indexador'].fillna('Pré')


df = df[['Data_Ref',
         'Origem',
         'COD_CLIENTE', 
         'COD_ASSESSOR', 
         'Produto',
         'Sub Produto', 
         'Class_Inove', 
         'Ativo', 
         'COD_ATIVO', 
         'Emissor', 
         'Indexador/Benchmark', 
         'Dt_Aplicacao', 
         'Vencimento', 
         'Carencia', 
         'Qtd', 
         'PU_ATUAL', 
         'NET_Original', 
         'Liquidez',
         'CNPJ Fundo',
         'Margem Total (Margem Taxa Adm + Margem Taxa Perf)',
         'Taxa de Performance x Comissão***',
         'Taxa de Administração x Comissão',
         'Taxa de Performance',
         'Taxa de Administração',
         'CVM',
         'Anbima',
         'Classificação XP',
         'Gestor',
         'Administrador',
         'Custodiante',
         'Onshore/Offshore',
         'PERPETUO',
         'COD_CLIENTE_XPUS',
         'Taxa',
         'Indexador',
         
         
        
        ]]


df_rf = df



df = pd.concat([df_diversificador, df_rf], ignore_index=True)

# Criar a condição para a coluna "Sub Produto"
df['Vencimento'] = df.apply(lambda x: '' if (x['Produto'] == 'Cripto') else x['Vencimento'], axis=1)

# Encontrar o índice da linha onde a coluna 'NET_Original' contém a palavra 'NET'
linha_net = df.loc[df['NET_Original'] == 'NET'].index

# Remover a linha com a palavra 'NET', se existir
if not linha_net.empty:
    df = df.drop(linha_net)
    
df_rf_diversificador = df



# In[5]:


print(df)


# In[6]:


#### ENTRADA ADDEPAR
PASTA_ARQUIVOS_PADROES = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões'
PASTA_ADEPPAR_PORTFOLIO_HOLDING_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Adeppar\Portfolio Holding\Historico'
PASTA_ADEPPAR = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Adeppar'



# Lista todos os arquivos na pasta e filtra aqueles que começam com "Diversificacao_"
arquivos_filtrados = [arquivo for arquivo in os.listdir(PASTA_ADEPPAR_PORTFOLIO_HOLDING_HISTORICO) if arquivo.startswith("Diversificador_XPUS_Adeppar_Portfolio_Holding_")]

# Encontra o arquivo mais recente usando a função max() com uma chave personalizada para comparar o tempo de modificação dos arquivos
filename = max(arquivos_filtrados, key=lambda x: os.path.getmtime(os.path.join(PASTA_ADEPPAR_PORTFOLIO_HOLDING_HISTORICO, x)))



# Extrai o valor após o primeiro sublinhado

valor = filename.split('_')[5]
valor = valor.split('_')[0]
valor = valor.split('.xlsx')[0]
valor = valor.replace('-', '')

date_formatted = datetime.strptime(valor, "%Y%m%d")

filepath = os.path.join(PASTA_ADEPPAR_PORTFOLIO_HOLDING_HISTORICO, filename)
df = pd.read_excel(filepath, header = 2)

# salvar a primeira linha como cabeçalho de coluna
df.columns = df.iloc[0]

# remover a primeira linha
df = df.iloc[1:]

time.sleep(5)


df['Valor Total']=""
# Função para aplicar a fórmula na coluna 1
def aplicar_formula(coluna):
    if coluna == "" or coluna == "Total":
        return ""
    elif coluna in ["Bond", "Unknown", "Equity", "Stock", "Cash", "USD", "REIT", "ETF", "EUR", "Alternative", "Balanced", "Semi-liquid"]:
        return df.loc[df['Top Level Owner'] == coluna, 'Valor Total'].values[0]
    else:
        try:
            pos_espaco = coluna.index(" ")
            return df.loc[df['Top Level Owner'] == coluna, 'Valor Total'].values[0]
        except ValueError:
            return coluna

# Cria a coluna 1 com a fórmula aplicada
df['Coluna1'] = df['Top Level Owner'].apply(aplicar_formula)

# Cria a coluna 2 com a seguinte fórmula, se a coluna 1 for vazia, apresentar o último conteúdo da coluna 1, se não for vazia, apresentar a coluna 1
df['Coluna2'] = df['Coluna1'].replace('', method='ffill')

# Função para repetir o valor da coluna Top Level Owner, se for um dos tipos de ativos abaixo
def repetir_valor(row):
    if row['Top Level Owner'] in ['Equity', 'Alternative', 'Bond', 'Cash & Cash Equivalent', 'Structured Product', 'Structured Note', 'Balanced', 'Fixed Income']:
        return row['Top Level Owner']
    else:
        return ''

# Cria a coluna Class_Ativos com a fórmula aplicada
df['Class_Ativos_Previa'] = df.apply(repetir_valor, axis=1)
df['Class_Ativos'] = df['Class_Ativos_Previa'].replace('', method='ffill')

# Função para comparar os valores da coluna Top Level Owner, e repetir o valor se for um dos tipos de ativos abaixo
def comparar_valores(row):
    if row['Top Level Owner'] in ['Unknown','Stock', 'Mutual Fund', 'REIT', 'ETF', 'Bond ETF', 'Stock', 'Money Market Fund', 'Certificate of Deposity', 'Closed End Fund', 'Semi-liquid', 'Cash']:
        return row['Top Level Owner']
    else:
        return ''



# Cria a coluna Tipo_Ativos com a fórmula aplicada
df['Tipo_Ativos_Previa'] = df.apply(comparar_valores, axis=1)
df['Tipo_Ativos'] = df['Tipo_Ativos_Previa'].replace('', method='ffill')

df['Sim'] = "Sim"
df['Manter?2'] = "Não"

df['Nova_Coluna'] = df.apply(lambda row: row['Top Level Owner'] if row['Top Level Owner'] not in [row['Tipo_Ativos'], row['Class_Ativos'], row['Coluna2']] else '', axis=1)

df['CCY2'] = df.apply(lambda row: row['CCY'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['QUANTITY2'] = df.apply(lambda row: row['QUANTITY'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)

df['UNIT COST2'] = df.apply(lambda row: row['UNIT COST'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['TOTAL COST2'] = df.apply(lambda row: row['TOTAL COST'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['mkt price2'] = df.apply(lambda row: row['mkt price'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['mkt value2'] = df.apply(lambda row: row['mkt value'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['accrued interest2'] = df.apply(lambda row: row['accrued interest'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['P/L OPEN2'] = df.apply(lambda row: row['P/L OPEN'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['TWR (USD)2'] = df.apply(lambda row: row['TWR (USD)'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['%2'] = df.apply(lambda row: row['%'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)

df['Manter?2'] = df.apply(lambda row: row['Sim'] if all([row['Tipo_Ativos'] != '', row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else 'Não', axis=1)

# Excluir as linhas em que Manter? seja igual a Não
df = df[df['Manter?2'] == 'Sim']
df = df[df['Nova_Coluna'] != 'Sim']
df = df[df['Nova_Coluna'] != 'Total']
df['Nova_Coluna'] = df['Nova_Coluna'].replace('USD', 'Cash')


df['Date'] = pd.to_datetime(date_formatted, format='%Y-%m-%d').date()


df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d')


df = df.dropna(subset=['CCY'])

df.drop(columns=['Manter?2','Sim','Tipo_Ativos_Previa','Class_Ativos_Previa','Top Level Owner', 'CCY', 'QUANTITY', 'UNIT COST', 'TOTAL COST', 'mkt price', 'mkt value', 'accrued interest', 'P/L OPEN', 'TWR (USD)', '%', 'Valor Total', 'Coluna1'], inplace=True)
df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')


df = df.rename(columns={'Nova_Coluna':'Empresa','Coluna2': 'COD_CLIENTE_XP_US','CCY2': 'CCY', 'QUANTITY2': 'QUANTITY', 'UNIT COST2': 'UNIT COST', 'TOTAL COST2': 'TOTAL COST', 'mkt price2': 'mkt price', 'mkt value2': 'mkt value', 'accrued interest2': 'accrued interest', 'P/L OPEN2': 'P/L OPEN', 'TWR (USD)2': 'TWR (USD)', '%2': '%'})
# Define a ordem das colunas
nova_ordem = ['Date', 'COD_CLIENTE_XP_US', 'Class_Ativos','Tipo_Ativos','Empresa','CCY', 'QUANTITY', 'UNIT COST', 'TOTAL COST', 'mkt price', 'mkt value', 'accrued interest', 'P/L OPEN', 'TWR (USD)', '%']


df['Empresa2'] = df['Empresa']
df['Class_Ativos2'] = df['Class_Ativos']

#dividir a coluna ATIVO em duas colunas
df[['Empresa', 'VENCIMENTO']] = df['Empresa'].str.split('Due ', expand=True)

# preencher as linhas vazias da coluna VENCIMENTO com 'Mar 25, 2300'
df['VENCIMENTO'].fillna('Mar 25, 2100', inplace=True)


df[['MES', 'DIA', 'ANO']] = df['VENCIMENTO'].str.split(' ', expand=True)
df['DIA'] = df['DIA'].str.strip(',')

meses_dict = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
              'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}

df['MES'] = df['MES'].map(meses_dict)

df['VENCIMENTO'] = df.apply(lambda row: str(row['DIA']) + '-' + str(row['MES']) + '-' + str(row['ANO']), axis=1)

df['VENCIMENTO'] = pd.to_datetime(df['VENCIMENTO'], format='%d-%m-%Y') # converter para objetos datetime
df['VENCIMENTO'] = df['VENCIMENTO'].dt.strftime('%d-%m-%Y') # formatar a coluna como uma string no formato desejado

# criar uma nova coluna PERPETUO com base na coluna ATIVO
df['PERPETUO'] = np.where(df['Empresa'].str.contains('Perp'), 'Sim', 'Não')

# excluir a palavra "Perp" da coluna ATIVO
df['Empresa'] = df['Empresa'].str.replace('Perp', '')



#Separar taxa 1
# contar espaços e separar a partir do penúltimo espaço
df['num_espacos'] = df['Empresa'].apply(lambda x: x.count(' '))
df['taxa_separada'] = df['Empresa'].apply(lambda x: ' '.join(x.split(' ')[-2:]))
df.loc[df['Empresa'].str.contains('%'), 'taxa_separada'] = df['Empresa'].apply(lambda x: ' '.join(x.split(' ')[-3:-1]))
df['Empresa'] = df['Empresa'].apply(lambda x: ' '.join(x.split(' ')[:-3]) if '%' in x else x)
df['Empresa'] = df['Empresa'].str.replace(r'\d{2}/\d{2}/\d{2}', '', regex=True)

# Use expressões regulares e o método str.extract() para extrair o conteúdo desejado
df['taxa_separada2'] = df['Empresa'].str.extract(r'(\d+\.\d+|\d+\s\d+\/\d+|\d+)')

# excluir o conteúdo de taxa_separada2 de ATIVO
df['Empresa'] = df['Empresa'].str.replace(r'(\d+\.\d+|\d+\s\d+\/\d+|\d+)', '', regex=True)

# remover possíveis espaços duplicados resultantes do replace
df['Empresa'] = df['Empresa'].str.replace('  ', ' ')

df['taxa_separada'] = df['taxa_separada'].str.replace(r'\b\d+/\d+\b', '', regex=True)

df['taxa_separada'] = df['taxa_separada'].replace(' ', np.nan)
df['taxa_separada'] = df['taxa_separada'].fillna(df['taxa_separada2'])


df = df.rename(columns={'Class_Ativos':'Produto','taxa_separada': 'Indexador/Benchmark'})
# Define a ordem das colunas

df['Produto'] = df.apply(lambda x: 'Fundos' if (x['Class_Ativos2'] == 'Balanced') else x['Produto'], axis=1)
df['Produto'] = df.apply(lambda x: 'Renda Variável' if (x['Class_Ativos2'] == 'ETF') else x['Produto'], axis=1)
df['Produto'] = df.apply(lambda x: 'Renda Fixa' if (x['Class_Ativos2'] == 'Bond') else x['Produto'], axis=1)
df['Produto'] = df.apply(lambda x: 'Fundos' if (x['Tipo_Ativos'] == 'Money Market Fund') else x['Produto'], axis=1)
df['Produto'] = df.apply(lambda x: 'Somente Financeiro' if (x['Tipo_Ativos'] == 'Cash' and x['Produto'] == 'Cash & Cash Equivalent') else x['Produto'], axis=1)
df['Produto'] = df.apply(lambda x: 'Renda Variável' if (x['Class_Ativos2'] == 'Equity') else x['Produto'], axis=1)
df['Produto'] = df.apply(lambda x: 'Renda Fixa' if (x['Class_Ativos2'] == 'Fixed Income') else x['Produto'], axis=1)
df['Produto'] = df.apply(lambda x: 'Renda Fixa' if (x['Class_Ativos2'] == 'Structured Note') else x['Produto'], axis=1)
df['Tipo_Ativos'] = df.apply(lambda x: 'Bond' if (x['Tipo_Ativos'] == 'Unknown') else x['Tipo_Ativos'], axis=1)
df['Indexador/Benchmark'] = df.apply(lambda x: x['Indexador/Benchmark'] if (x['Class_Ativos2'] == 'Bond' or x['Class_Ativos2'] == 'Structured Note' or x['Tipo_Ativos'] == 'Bond') else '', axis=1)
df['Tipo_Ativos'] = df.apply(lambda x: 'T-Bill' if 'United States Treas Bills' in x['Empresa'] else ('T-Bond' if 'United States Treas BDS' in x['Empresa'] else x['Tipo_Ativos']), axis=1)
df['VENCIMENTO'] = df.apply(lambda x: '' if (x['Produto'] != 'Renda Fixa') else x['VENCIMENTO'], axis=1)
df['Emissor'] = df.apply(lambda x: '' if (x['Produto'] != 'Renda Fixa' and x['Tipo_Ativos'] != 'Mutual Fund'
                                         and x['Tipo_Ativos'] != 'Cash'and x['Tipo_Ativos'] != 'Closed End Fund'
                                         and x['Tipo_Ativos'] != 'Money Market Fund'
                                         ) else x['Empresa2'], axis=1)
df['Produto'] = df.apply(lambda x: 'Fundos' if (x['Tipo_Ativos'] == 'Mutual Fund'
                                         ) else x['Produto'], axis=1)
colunas_para_excluir = [
    'CCY', 'UNIT COST', 'TOTAL COST',
    'accrued interest', 'P/L OPEN', 'TWR (USD)', '%',
     'MES', 'DIA', 'ANO', 'num_espacos','Empresa','Class_Ativos2','taxa_separada2'
]

# Excluir as colunas
df = df.drop(colunas_para_excluir, axis=1)

df['Origem'] = filename
df['Class_Inove'] = "Internacional"
colunas_vazias = ['Dt_Aplicacao','Carencia','Liquidez','COD_ATIVO', 'CNPJ Fundo', 'Margem Total (Margem Taxa Adm + Margem Taxa Perf)', 'Taxa de Performance x Comissão***', 'Taxa de Administração x Comissão', 'Taxa de Performance', 'Taxa de Administração', 'CVM', 'Anbima', 'Classificação XP', 'Gestor', 'Administrador', 'Custodiante']
df['Onshore/Offshore'] = "Offshore"

for coluna in colunas_vazias:
    df[coluna] = ''


df = df.rename(columns={'Empresa2':'Ativo','Tipo_Ativos':'Sub Produto','Date':'Data_Ref','COD_CLIENTE_XP_US': 'COD_CLIENTE','VENCIMENTO': 'Vencimento',
                        'QUANTITY': 'Qtd', 'mkt price': 'PU_ATUAL', 'mkt value': 'NET_Original'})



# Identificar o arquivo com o prefixo 'dAssessores'
arquivo_xp_us = None
for arquivo in os.listdir(PASTA_ARQUIVOS_PADROES):
    if arquivo.startswith('Clientes XP US'):
        arquivo_xp_us = arquivo
        break

if arquivo_xp_us is not None:
    # Carregar o arquivo em um DataFrame
    caminho_arquivo = os.path.join(PASTA_ARQUIVOS_PADROES, arquivo_xp_us)
    df_clientes_xpus = pd.read_excel(caminho_arquivo)

    # Verificar o resultado
    
else:
    print("Nenhum arquivo encontrado com o prefixo 'Assessores'.")

df = pd.merge(df, df_clientes_xpus, left_on='COD_CLIENTE', right_on='COD_CLIENTE', how='left')

# Identificar o arquivo com o prefixo 'dAssessores'
arquivo_xp_us = None
for arquivo in os.listdir(PASTA_DOWNLOADS):
    if arquivo.startswith('Custodia_OffShore_XPUS_'):
        arquivo_xp_us = arquivo
        break

if arquivo_xp_us is not None:
    # Carregar o arquivo em um DataFrame
    caminho_arquivo = os.path.join(PASTA_DOWNLOADS, arquivo_xp_us)
    df_clientes_xpus = pd.read_excel(caminho_arquivo)

    # Verificar o resultado
    
else:
    print("Nenhum arquivo encontrado com o prefixo 'Assessores'.")

df = pd.merge(df, df_clientes_xpus, left_on='COD_CLIENTE', right_on='Código Cliente', how='left')

# Excluir o primeiro caractere 'A' das linhas da coluna 'Assessor'
df['Código Assessor'] = df['Código Assessor'].astype(str)
# Preencher células vazias com 'A-'
df['Código Assessor'] = df['Código Assessor'].replace('', 'A-')

# Preencher células vazias com 'A-'
df['Código Cliente'] = df['Código Cliente'].replace('', '-')


# Excluir o primeiro caractere 'A' das linhas da coluna 'Assessor'
df['Código Assessor'] = df['Código Assessor'].str.slice(1)

# Preencher linhas vazias de 'COD_ASSESSOR' com os valores de 'Assessor'
df['COD_ASSESSOR'] = df['COD_ASSESSOR'].fillna(df['Código Assessor'])

# Preencher linhas vazias de 'Assessor' com o texto "Conta XPUS - Assessor Nao Identificado"
df['COD_ASSESSOR'] = df['COD_ASSESSOR'].fillna("Conta XPUS - Assessor Nao Identificado")

colunas_para_excluir = [
    'Custódia','COD_CLIENTE_ON','PF/PJ','Código Cliente','Código Assessor'
]
df = df.drop(colunas_para_excluir, axis=1)
 
df['Taxa'] = ''
df['Taxa'] = df['Indexador/Benchmark']
df['Taxa'] = df['Taxa'].str.replace(' %', '')
df['Taxa'] = df['Taxa'].replace('', '0')
df['Taxa'] = df['Taxa'].replace('(usd) (offshore)', '0')
df['Taxa'] = df['Taxa'].replace('(usd) (off)', '0')

def convert_fraction_to_decimal(value):
    if isinstance(value, str) and ' ' in value and '/' in value:
        integer_part, fraction_part = value.split()
        numerator, denominator = map(int, fraction_part.split('/'))
        decimal_value = float(integer_part) + float(numerator) / float(denominator)
        return decimal_value
    return value

df['Taxa2'] = df['Taxa'].apply(convert_fraction_to_decimal)

df['Taxa'] = df['Taxa2']
# Remover as colunas "COD_ATIVO_FIP" e "COD_ATIVO_2"
df = df.drop(['Taxa2'], axis=1)

df['Taxa'] = df['Taxa'].astype(float)

df['Indexador'] = 'Pré'

df = df[['Data_Ref',
         'Origem',
         'COD_CLIENTE', 
         'COD_ASSESSOR', 
         'Produto',
         'Sub Produto', 
         'Class_Inove', 
         'Ativo', 
         'COD_ATIVO', 
         'Emissor', 
         'Indexador/Benchmark', 
         'Dt_Aplicacao', 
         'Vencimento', 
         'Carencia', 
         'Qtd', 
         'PU_ATUAL', 
         'NET_Original', 
         'Liquidez',
         'CNPJ Fundo',
         'Margem Total (Margem Taxa Adm + Margem Taxa Perf)',
         'Taxa de Performance x Comissão***',
         'Taxa de Administração x Comissão',
         'Taxa de Performance',
         'Taxa de Administração',
         'CVM',
         'Anbima',
         'Classificação XP',
         'Gestor',
         'Administrador',
         'Custodiante',
         'Onshore/Offshore',
         'PERPETUO',
         'COD_CLIENTE_XPUS',
         'Taxa',
         'Indexador',
        
         
        
        ]]


df.to_excel(os.path.join(PASTA_ADEPPAR, f'Diversificador_XPUS_Adeppar_Portfolio_Holding_Edit.xlsx'), index=False)  


df = pd.concat([df_rf_diversificador, df], ignore_index=True)

df3= df


# In[7]:


print(df)


# In[8]:


##############OffShore XPUS Fixed Income
PASTA_ADEPPAR_FIXED_INCOME_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Adeppar\Fixed Income\Historico'
# Encontra o arquivo mais recente que começa com 'diversificacao_'

# Lista todos os arquivos na pasta e filtra aqueles que começam com "Diversificacao_"
arquivos_filtrados = [arquivo for arquivo in os.listdir(PASTA_ADEPPAR_FIXED_INCOME_HISTORICO) if arquivo.startswith("Diversificador_XPUS_Adeppar_Fixed_Income_")]

# Encontra o arquivo mais recente usando a função max() com uma chave personalizada para comparar o tempo de modificação dos arquivos
filename = max(arquivos_filtrados, key=lambda x: os.path.getmtime(os.path.join(PASTA_ADEPPAR_FIXED_INCOME_HISTORICO, x)))



valor = filename.split('_')[5]
valor = valor.split('_')[0]
valor = valor.split('.xlsx')[0]
valor = valor.replace('-', '')
date_formatted = datetime.strptime(valor, "%Y%m%d")

# Read the saved file back into a dataframe

filepath = os.path.join(PASTA_ADEPPAR_FIXED_INCOME_HISTORICO, filename)
df = pd.read_excel(filepath, header = 3)
 
# salvar a primeira linha como cabeçalho de coluna
df.columns = df.iloc[0]

# remover a primeira linha
df = df.iloc[1:]


# Carregar o arquivo XLS em um dataframe do pandas
df['Valor Total']=""
# Função para aplicar a fórmula na coluna 1
def aplicar_formula(coluna):
    if coluna == "" or coluna == "Total":
        return ""
    elif coluna in ["Bond", "Unknown", "Structured Note", "Structured Product", "Certificate of Deposit"]:
        return df.loc[df['Holding Account'] == coluna, 'Valor Total'].values[0]
    else:
        try:
            pos_espaco = coluna.index(" ")
            return df.loc[df['Holding Account'] == coluna, 'Valor Total'].values[0]
        except ValueError:
            return coluna

# Cria a coluna 1 com a fórmula aplicada
df['Coluna1'] = df['Holding Account'].apply(aplicar_formula)

# Cria a coluna 2 com a seguinte fórmula, se a coluna 1 for vazia, apresentar o último conteúdo da coluna 1, se não for vazia, apresentar a coluna 1
df['Coluna2'] = df['Coluna1'].replace('', method='ffill')


# Função para repetir o valor da coluna Top Level Owner, se for um dos tipos de ativos abaixo
def repetir_valor(row):
    if row['Holding Account'] in ['Bond', 'Unknown', 'Structured Note', 'Structured Product', 'Certificate of Deposit']:
        return row['Holding Account']
    else:
        return ''

# Cria a coluna Class_Ativos com a fórmula aplicada
df['Class_Ativos_Previa'] = df.apply(repetir_valor, axis=1)
df['Class_Ativos'] = df['Class_Ativos_Previa'].replace('', method='ffill')

df['Sim'] = "Sim"
df['Manter?2'] = "Não"

df['Nova_Coluna'] = df.apply(lambda row: row['Holding Account'] if row['Holding Account'] not in [row['Class_Ativos'], row['Coluna2']] else '', axis=1)


df['Value2'] = df.apply(lambda row: row['Value'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['Cost Basis2'] = df.apply(lambda row: row['Cost Basis'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['Accrued Income2'] = df.apply(lambda row: row['Accrued Income'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['Coupon Rate2'] = df.apply(lambda row: row['Coupon Rate'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['Years to Maturity2'] = df.apply(lambda row: row['Years to Maturity'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['Yield to Worst2'] = df.apply(lambda row: row['Yield to Worst'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['Modified Duration2'] = df.apply(lambda row: row['Modified Duration'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)
df['% of Fixed Income2'] = df.apply(lambda row: row['% of Fixed Income'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else '', axis=1)


df['Manter?2'] = df.apply(lambda row: row['Sim'] if all([row['Class_Ativos'] != '', row['Coluna2'] != '', row['Nova_Coluna'] != '']) else 'Não', axis=1)

# Excluir as linhas em que Manter? seja igual a Não
df = df[df['Manter?2'] == 'Sim']
df = df[df['Nova_Coluna'] != 'Sim']
df = df[df['Nova_Coluna'] != 'Total']
df = df[df['Value2'] != 0]


df['Date'] = pd.to_datetime(date_formatted, format='%Y-%m-%d').date()

novo_nome = f"Diversificador_XPUS_Adeppar_Fixed_Income_{formatted_date}.xlsx"

df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d')



df.drop(columns=['Yield to Maturity',
                 'Valor Total',
                 'Holding Account',
                 'Value', 
                 'Class_Ativos_Previa', 
                 'Sim', 
                 'Manter?2', 
                 'Coluna1', 
                 'Cost Basis', 
                 'Accrued Income', 
                 'Coupon Rate', 
                 'Years to Maturity', 
                 'Yield to Worst', 
                 'Modified Duration', 
                 '% of Fixed Income'], inplace=True)



df = df.rename(columns={
    'Nova_Coluna':'ATIVO',
    'Coluna2': 'COD_CLIENTE_XPUS',
    'Value2': 'NET_ORIGINAL',
    'Cost Basis2': 'Cost Basis',
    'Accrued Income2': 'Accrued Income',
    'Coupon Rate2': 'Coupon Rate',
    'Class_Ativos': 'CLASSE_ATIVO',
    'Date':'Data_Ref',
    'Years to Maturity2': 'Years to Maturity',
    'Yield to Worst2': 'Yield to Worst',
    'Modified Duration2': 'Modified Duration',
    '% of Fixed Income2': '% of Fixed Income'
    
})


# Define a ordem das colunas
nova_ordem = ['Data_Ref',
              'COD_CLIENTE_XPUS',
              'ATIVO',
              'NET_ORIGINAL',
              'CLASSE_ATIVO'
              
]



#dividir a coluna ATIVO em duas colunas
df[['ATIVO', 'VENCIMENTO']] = df['ATIVO'].str.split('Due ', expand=True)

# preencher as linhas vazias da coluna VENCIMENTO com 'Mar 25, 2300'
df['VENCIMENTO'].fillna('Mar 25, 2100', inplace=True)


df[['MES', 'DIA', 'ANO']] = df['VENCIMENTO'].str.split(' ', expand=True)
df['DIA'] = df['DIA'].str.strip(',')

meses_dict = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
              'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}

df['MES'] = df['MES'].map(meses_dict)

df['VENCIMENTO'] = df.apply(lambda row: str(row['DIA']) + '-' + str(row['MES']) + '-' + str(row['ANO']), axis=1)

##########






df['VENCIMENTO'] = pd.to_datetime(df['VENCIMENTO'], format='%d-%m-%Y') # converter para objetos datetime
df['VENCIMENTO'] = df['VENCIMENTO'].dt.strftime('%d-%m-%Y') # formatar a coluna como uma string no formato desejado

# criar uma nova coluna PERPETUO com base na coluna ATIVO
df['PERPETUO'] = np.where(df['ATIVO'].str.contains('Perp'), 'Sim', 'Não')

# excluir a palavra "Perp" da coluna ATIVO
df['ATIVO'] = df['ATIVO'].str.replace('Perp', '')

#Separar taxa 1
# contar espaços e separar a partir do penúltimo espaço
df['num_espacos'] = df['ATIVO'].apply(lambda x: x.count(' '))
df['taxa_separada'] = df['ATIVO'].apply(lambda x: ' '.join(x.split(' ')[-2:]))
df.loc[df['ATIVO'].str.contains('%'), 'taxa_separada'] = df['ATIVO'].apply(lambda x: ' '.join(x.split(' ')[-3:-1]))
df['ATIVO'] = df['ATIVO'].apply(lambda x: ' '.join(x.split(' ')[:-3]) if '%' in x else x)
df['ATIVO'] = df['ATIVO'].str.replace(r'\d{2}/\d{2}/\d{2}', '', regex=True)

# Use expressões regulares e o método str.extract() para extrair o conteúdo desejado
df['taxa_separada2'] = df['ATIVO'].str.extract(r'(\d+\.\d+|\d+\s\d+\/\d+|\d+)')

# excluir o conteúdo de taxa_separada2 de ATIVO
df['ATIVO'] = df['ATIVO'].str.replace(r'(\d+\.\d+|\d+\s\d+\/\d+|\d+)', '', regex=True)

# remover possíveis espaços duplicados resultantes do replace
df['ATIVO'] = df['ATIVO'].str.replace('  ', ' ')

df['taxa_separada'] = df['taxa_separada'].str.replace(r'\b\d+/\d+\b', '', regex=True)



df['taxa_separada'] = df['taxa_separada'].replace(' ', np.nan)
df['taxa_separada'] = df['taxa_separada'].fillna(df['taxa_separada2'])


df = df.rename(columns={'Class_Ativos':'Produto','taxa_separada': 'Indexador/Benchmark'})
# Define a ordem das colunas

df['Taxa'] = df['Indexador/Benchmark']
df['Emissor'] = ''
df['Produto'] = 'Renda Fixa'
df['Sub Produto'] = df['CLASSE_ATIVO']
#df['Produto'] = df['ATIVO'].apply(lambda x: ' '.join(x.split(' ')[:-3]) if '%' in x else x)
df['Sub Produto'] = df.apply(lambda x: 'T-Bill' if 'United States Treas Bills' in x['ATIVO'] else ('T-Bond' if 'United States Treas BDS' in x['ATIVO'] else x['Sub Produto']), axis=1)
df['Emissor'] = df.apply(lambda x: '' if (x['Produto'] != 'Renda Fixa' and x['Sub Produto'] != 'Structured Note'
                                         and x['Sub Produto'] != 'Unknown'
                                         ) else x['Emissor'], axis=1)


df['Origem'] = filename
df['Class_Inove'] = "Internacional"
colunas_vazias = ['Dt_Aplicacao','Carencia','Liquidez','COD_ATIVO', 'CNPJ Fundo', 'Margem Total (Margem Taxa Adm + Margem Taxa Perf)', 'Taxa de Performance x Comissão***', 'Taxa de Administração x Comissão', 'Taxa de Performance', 'Taxa de Administração', 'CVM', 'Anbima', 'Classificação XP', 'Gestor', 'Administrador', 'Custodiante']
df['Onshore/Offshore'] = "Offshore"

for coluna in colunas_vazias:
    df[coluna] = ''





# Identificar o arquivo com o prefixo 'dAssessores'
arquivo_xp_us = None
for arquivo in os.listdir(PASTA_ARQUIVOS_PADROES):
    if arquivo.startswith('Clientes XP US'):
        arquivo_xp_us = arquivo
        break

if arquivo_xp_us is not None:
    # Carregar o arquivo em um DataFrame
    caminho_arquivo = os.path.join(PASTA_ARQUIVOS_PADROES, arquivo_xp_us)
    df_clientes_xpus = pd.read_excel(caminho_arquivo)

    # Verificar o resultado
    
else:
    print("Nenhum arquivo encontrado com o prefixo 'Assessores'.")

df = pd.merge(df, df_clientes_xpus, left_on='COD_CLIENTE_XPUS', right_on='COD_CLIENTE_XPUS', how='left')
#df['COD_CLIENTE'] = df.apply(lambda x: x['COD_CLIENTE_XPUS'] if (x['COD_CLIENTE_ON'] == '-') else x['COD_CLIENTE_ON'], axis=1)




# Identificar o arquivo com o prefixo 'dAssessores'
arquivo_xp_us = None
for arquivo in os.listdir(PASTA_DOWNLOADS):
    if arquivo.startswith('Custodia_OffShore_XPUS_'):
        arquivo_xp_us = arquivo
        break

if arquivo_xp_us is not None:
    # Carregar o arquivo em um DataFrame
    caminho_arquivo = os.path.join(PASTA_DOWNLOADS, arquivo_xp_us)
    df_clientes_xpus = pd.read_excel(caminho_arquivo)

    # Verificar o resultado
    
else:
    print("Nenhum arquivo encontrado com o prefixo 'Assessores'.")
    



df = pd.merge(df, df_clientes_xpus, left_on='COD_CLIENTE_XPUS', right_on='Código Cliente', how='left')
df['Código Assessor'] = df['Código Assessor'].astype(str)


# Preencher células vazias com 'A-'
df['Código Assessor'] = df['Código Assessor'].replace('', 'A-')

# Preencher células vazias com 'A-'
df['Código Cliente'] = df['Código Cliente'].replace('', '-')

# Excluir o primeiro caractere 'A' das linhas da coluna 'Assessor'
df['Código Assessor'] = df['Código Assessor'].str.slice(1)

# Preencher linhas vazias de 'COD_ASSESSOR' com os valores de 'Assessor'
df['COD_ASSESSOR'] = df['COD_ASSESSOR'].fillna(df['Código Assessor'])

# Preencher linhas vazias de 'Assessor' com o texto "Conta XPUS - Assessor Nao Identificado"
df['COD_ASSESSOR'] = df['COD_ASSESSOR'].fillna("Conta XPUS - Assessor Nao Identificado")

colunas_para_excluir = [
    'Custódia','COD_CLIENTE_ON','PF/PJ','Código Cliente','Código Assessor','Cost Basis', 'Accrued Income', 'Coupon Rate', 'Years to Maturity',
       'Yield to Worst', 'Modified Duration', '% of Fixed Income',
        'MES', 'DIA', 'ANO', 'num_espacos','taxa_separada2','COD_CLIENTE_ON'
]

# Excluir as colunas
df = df.drop(colunas_para_excluir, axis=1)



df = df.rename(columns={'VENCIMENTO': 'Vencimento','DATA_REF':'Data_Ref','ATIVO':'Ativo','NET_ORIGINAL':'NET_Original','COD_CLIENTE_XPUS_x': 'COD_CLIENTE_XPUS'})

df['Qtd'] = ''
df['PU_ATUAL'] = ''


df['Indexador'] = 'Pré'



df = df[['Data_Ref',
         'Origem',
         'COD_CLIENTE', 
         'COD_ASSESSOR', 
         'Produto',
         'Sub Produto', 
         'Class_Inove', 
         'Ativo', 
         'COD_ATIVO', 
         'Emissor', 
         'Indexador/Benchmark', 
         'Dt_Aplicacao', 
         'Vencimento', 
         'Carencia', 
         'Qtd', 
         'PU_ATUAL', 
         'NET_Original', 
         'Liquidez',
         'CNPJ Fundo',
         'Margem Total (Margem Taxa Adm + Margem Taxa Perf)',
         'Taxa de Performance x Comissão***',
         'Taxa de Administração x Comissão',
         'Taxa de Performance',
         'Taxa de Administração',
         'CVM',
         'Anbima',
         'Classificação XP',
         'Gestor',
         'Administrador',
         'Custodiante',
         'Onshore/Offshore',
         'PERPETUO',
         'COD_CLIENTE_XPUS',
         'Taxa',
         'Indexador',
       
         
         
        
        ]]
df['Vencimento'] = pd.to_datetime(df['Vencimento'])



df['Taxa'] = ''
df['Taxa'] = df['Indexador/Benchmark']
df['Taxa'] = df['Taxa'].str.replace(' %', '')

def convert_fraction_to_decimal(value):
    if isinstance(value, str) and ' ' in value and '/' in value:
        integer_part, fraction_part = value.split()
        numerator, denominator = map(int, fraction_part.split('/'))
        decimal_value = float(integer_part) + float(numerator) / float(denominator)
        return decimal_value
    return value

df['Taxa2'] = df['Taxa'].apply(convert_fraction_to_decimal)

df['Taxa'] = df['Taxa2']
# Remover as colunas "COD_ATIVO_FIP" e "COD_ATIVO_2"
df = df.drop(['Taxa2'], axis=1)
df['Taxa'] = df['Taxa'].replace('', '0')
df['Taxa'] = df['Taxa'].replace('(usd) (offshore)', '0')
df['Taxa'] = df['Taxa'].replace('(usd) (off)', '0')
df['Taxa'] = df['Taxa'].astype(float)



df.to_excel(os.path.join(PASTA_ADEPPAR, f'Diversificador_XPUS_Adeppar_Fixed_Income-Edit.xlsx'), index=False)

df = pd.concat([df3, df], ignore_index=True)






# In[9]:


df['Nova_Coluna'] = df['Ativo'].apply(lambda x: 'LFT' if 'LFT' in str(x) else '')


# Exclui o produto Renda Fixa, exceto Produto Estruturado (COE)
mask = (df['Sub Produto'] == 'Tesouro Direto') & (df['Nova_Coluna'] != 'LFT')


# remove as linhas que atendem a condição da máscara
df = df.drop(df[mask].index)



PASTA_MTM_AGIO_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Renda Fixa HUB\MTM Agio\Historico'

# Lista todos os arquivos na pasta e filtra aqueles que começam com "Diversificacao_"
arquivos_filtrados = [arquivo for arquivo in os.listdir(PASTA_MTM_AGIO_HISTORICO) if arquivo.startswith("Renda_Fixa_MTM_Ágio_")]

# Encontra o arquivo mais recente usando a função max() com uma chave personalizada para comparar o tempo de modificação dos arquivos
filename = max(arquivos_filtrados, key=lambda x: os.path.getmtime(os.path.join(PASTA_MTM_AGIO_HISTORICO, x)))


# Extrai o valor após o primeiro sublinhado

valor = filename.split('_')[4]
valor = valor.split('_')[0]
valor = valor.split('.xlsx')[0]
valor = valor.replace('-', '')

valor = valor[6:] + "/" + valor[4:6] + "/" + valor[:4]

valor2 = valor[:2] + "/" + valor[3:5] + "/" + valor[6:]





filepath = os.path.join(PASTA_MTM_AGIO_HISTORICO, filename)
df_agio = pd.read_excel(filepath)



df_agio['Data_Ref'] = valor2

df_agio = df_agio[df_agio['Classe'] == 'Tesouro Direto']

# Preenche vazios na coluna Ativo com NTN-B1
df_agio['Ativo'].fillna('NTN-B1', inplace=True)

# Na coluna Ativo, altera NTNB para NTN-B
df_agio['Ativo'] = df_agio['Ativo'].str.replace('NTNB', 'NTN-B')

# Na coluna Ativo, altera NTNF para NTN-F
df_agio['Ativo'] = df_agio['Ativo'].str.replace('NTNF', 'NTN-F')

# Na coluna Indexador, troca Prefixado por Pré
df_agio['Indexador'] = df_agio['Indexador'].str.replace('Prefixado', 'Pré')

# Cria coluna Produto com valor Renda Fixa
df_agio['Produto'] = 'Renda Fixa'

# Cria coluna Emissor com Tesouro Nacional
df_agio['Emissor'] = 'Tesouro Nacional'

# Cria coluna Emissor com Tesouro Nacional
df_agio['Onshore/Offshore'] = 'Onshore'

df_agio['Emissor'] = df_agio.apply(lambda x: 'TESOURO NACIONAL' if (x['Ativo'] == 'NTN-B1') else x['Emissor'], axis=1)
#df['Taxa'] = df.apply(lambda x: 0 if (x['Sub Produto'] == 'Tesouro Direto') else x['Taxa'], axis=1)

df_agio['Disponível'] = df_agio['Disponível'].astype(float)
df_agio['Valor Curva de Aplicação'] = df_agio['Valor Curva de Aplicação'].astype(float)
import numpy as np
df_agio['PU_ATUAL'] = np.where((df_agio['Valor Curva de Aplicação'] == 0) | (df_agio['Disponível'] == 0), 0, df_agio['Valor Curva de Aplicação'] / df_agio['Disponível'])



df_agio['Origem'] = filename

df_agio["Class_Inove"] = np.where(df_agio["Indexador"] == "Renda+", "Inflação",
                        np.where(df_agio["Indexador"] == "Pré", "Pré-Fixado",
                        np.where(df_agio["Indexador"] == "IPCA+", "Inflação","-"
                                )))



# Renomeia coluna Disponível para Qtd
df_agio.rename(columns={'Disponível': 'Qtd',
                        'Classe': 'Sub Produto',
                        'Data Aplicação':'Dt_Aplicacao',
                        'Código Assessor': 'COD_ASSESSOR',
                        'Codigo Cliente':'COD_CLIENTE',
                        'Ticker':'COD_ATIVO',
                        'Valor Curva de Aplicação':'NET_Original',
                        'Taxa Aplicação Média':'Taxa'
                       }, inplace=True)




columns_to_drop = [
    'PF/ PJ',
    'Perfil Suitability',
    'Tipo Ativo',
    'Valor Aplicado',
    'Financeiro Bruto de Venda',
    'Ágio Médio',
    'Alíquota IR Média (Se aplicável)',
    'Maior Alíquota IR (Se aplicável)',
    'Taxa Venda Atual',
    'Duration (Prazo Médio)',
    'Pu Anbima',
    'Taxa Anbima',
    'Financeiro Anbima',
    'Disponível Swap',
    'Data Atualização',
    'Escritório'
]

df_agio = df_agio.drop(columns=columns_to_drop)




df_agio['Vencimento'] = pd.to_datetime(df_agio['Vencimento'], format='%d/%m/%Y')
df_agio['Indexador/Benchmark']=''

df_agio['Carencia']=''
df_agio['Liquidez'] = (df_agio['Vencimento'] - pd.Timestamp.now()).dt.days


columns_to_add = ['CNPJ Fundo', 'Margem Total (Margem Taxa Adm + Margem Taxa Perf)', 'Taxa de Performance x Comissão***', 'Taxa de Administração x Comissão', 'Taxa de Performance', 'Taxa de Administração', 'CVM', 'Anbima', 'Classificação XP', 'Gestor', 'Administrador', 'Custodiante', 'Onshore/Offshore', 'PERPETUO', 'COD_CLIENTE_XPUS']

for col in columns_to_add:
    df_agio[col] = ''

df_agio['Onshore/Offshore']='Onshore'

df_agio = df_agio[['Data_Ref',
         'Origem',
         'COD_CLIENTE', 
         'COD_ASSESSOR', 
         'Produto',
         'Sub Produto', 
         'Class_Inove', 
         'Ativo', 
         'COD_ATIVO', 
         'Emissor', 
         'Indexador/Benchmark', 
         'Dt_Aplicacao', 
         'Vencimento', 
         'Carencia', 
         'Qtd', 
         'PU_ATUAL', 
         'NET_Original', 
         'Liquidez',
         'CNPJ Fundo',
         'Margem Total (Margem Taxa Adm + Margem Taxa Perf)',
         'Taxa de Performance x Comissão***',
         'Taxa de Administração x Comissão',
         'Taxa de Performance',
         'Taxa de Administração',
         'CVM',
         'Anbima',
         'Classificação XP',
         'Gestor',
         'Administrador',
         'Custodiante',
         'Onshore/Offshore',
         'PERPETUO',
         'COD_CLIENTE_XPUS',
         'Taxa',
         'Indexador',
        
                   
         
        
        ]]


#df['Taxa'] = df['Taxa'].astype(str).apply(lambda x: 0 + x if x.startswith('-') else x)

df_agio['Taxa'] = df_agio['Taxa'].replace('.', ',')
df_agio['Indexador'] = df_agio['Indexador'].replace('IPCA+', 'IPC-A +')

df.reset_index(drop=True, inplace=True)
df_agio.reset_index(drop=True, inplace=True)

df_concatenado = pd.concat([df, df_agio], ignore_index=True)


df = pd.concat([df, df_agio], ignore_index=True)
# Na coluna Ativo, altera NTNF para NTN-F


# In[10]:


## COtacao Dolar
cotacao2 = 4.90
####### BUSCA COTACAO DOLAR

df['NET_Original'] = df['NET_Original'].astype(float)
df['NET Modulo'] = np.abs(df['NET_Original'])
try:

    ##### COnverter valores

    url = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados?formato=json&dataInicial=01/01/2021'

    # envia a requisição para a API
    response = requests.get(url)

    # verifica se a requisição foi bem sucedida
    if response.status_code == 200:
        # converte a resposta em um DataFrame
        df_cotacoes = pd.DataFrame(response.json())

        # renomeia as colunas
        df_cotacoes.columns = ['datahora', 'cotacao']

        # converte a coluna 'data' para o tipo datetime
        df_cotacoes['datahora'] = pd.to_datetime(df_cotacoes['datahora'], dayfirst=True)

        # ordena o DataFrame pela data
        df_cotacoes = df_cotacoes.sort_values('datahora')
        df_cotacoes['datahora'] = df_cotacoes['datahora'].dt.strftime('%Y-%m-%d')

        # encontra o maior valor da coluna datahora
        ultima_data = df_cotacoes['datahora'].max()

        # cria um novo DataFrame com as datas faltantes
        df_faltantes = pd.DataFrame(pd.date_range(start=ultima_data, end=datetime.now(), freq='D'), columns=['datahora'])
        # imprime o DataFrame resultante

    else:
        print('Erro ao buscar cotações')

    # obtém a data atual
    data_atual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # verifica se a data do último registro é anterior a hoje
    if df_cotacoes['datahora'].max() < data_atual:
        # cria um novo DataFrame com as datas a serem preenchidas
        df_faltantes = pd.DataFrame(pd.date_range(start=df_cotacoes['datahora'].max(), end=datetime.today(), freq='D'), columns=['datahora'])

        # adiciona a coluna 'cotacao' com o último valor disponível
        ultimo_valor = df_cotacoes['cotacao'].iloc[-1]
        df_faltantes['cotacao'] = ultimo_valor

    else:
        # caso contrário, não há datas faltantes a serem preenchidas
        df_faltantes = pd.DataFrame(columns=['datahora', 'cotacao'])

    # ajusta o tipo da coluna 'datahora' para o tipo datetime
    df_faltantes['datahora'] = pd.to_datetime(df_faltantes['datahora'], format='%Y-%m-%d')

    # ordena o DataFrame pela data
    df_faltantes = df_faltantes.sort_values('datahora')

    # cria um novo DataFrame com as datas a serem preenchidas
    df_faltantes = pd.DataFrame(pd.date_range(start=df_cotacoes['datahora'].max(), end=datetime.today(), freq='D'), columns=['datahora'])

    # adiciona a coluna 'cotacao' com o último valor disponível
    ultimo_valor = df_cotacoes['cotacao'].iloc[-1]
    df_faltantes['cotacao'] = ultimo_valor


    # concatena os DataFrames
    df_cotacoes = pd.concat([df_cotacoes, df_faltantes], axis=0, ignore_index=True)


    df_cotacoes['datahora'] = pd.to_datetime(df_cotacoes['datahora'], format='%Y-%m-%d %H:%M:%S')

    df_cotacoes['cotacao'] = pd.to_numeric(df_cotacoes['cotacao'])

    #df['Data_Ref'] = pd.to_datetime(df['Data_Ref'], format='%Y/%m/%d')
    df_cotacoes['datahora'] = df_cotacoes['datahora'].dt.strftime('%d/%m/%Y')
    #df_cotacoes['datahora'] = pd.to_datetime(df_cotacoes['datahora'])
    df = pd.merge(df, df_cotacoes, left_on='Data_Ref', right_on='datahora', how='left')
    # Criar a coluna "COD_ATIVO_2" com a função desejada
    df['cotacao'] = df.apply(lambda x: 1 if (x['Onshore/Offshore'] == 'Onshore') else x['cotacao'], axis=1)
    # Remover as colunas "COD_ATIVO_FIP" e "COD_ATIVO_2"
    df = df.drop(['datahora'], axis=1)
    df = df.rename(columns={'cotacao':'Cotação'})

    df['NET Modulo Posicao'] = df['NET Modulo']*df['Cotação']
except:
     df['NET Modulo Posicao'] = df['NET Modulo']*cotacao2

df['ID'] = df[['COD_ASSESSOR', 'COD_CLIENTE', 'Produto', 'Sub Produto', 'Ativo', 'Emissor','COD_ATIVO','NET_Original','Indexador/Benchmark', 'Vencimento']].apply(lambda x: '_'.join(x.astype(str)), axis=1)

df = df.drop_duplicates(subset=['ID'])




# In[11]:


PASTA_SALDO_EM_CONTA = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Saldo em Conta'
PASTA_NET = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\NET'
# Adicionar Suitability(NET_)


novo_nome3 = f'NET.xlsx'
filepath = os.path.join(PASTA_NET, novo_nome3)
df_net = pd.read_excel(filepath)

df_net = df_net.drop(['Unnamed: 0'], axis=1)   
# Converter a coluna 'NET Modulo' para float
df_net['Conta XP'] = df_net['Conta XP'].astype(float)
#df['COD_CLIENTE'] = df['COD_CLIENTE'].str.strip()
df_net['Conta XP'] = df_net['Conta XP'].astype(str).apply(lambda x: x.split('.')[0])
df_net['Conta XP'] = df_net['Conta XP'].astype(str)


df = pd.merge(df, df_net[['Conta XP','Cliente','S', 'I','Posição']], left_on='COD_CLIENTE', right_on='Conta XP', how='left')



df = df.drop(['Conta XP'], axis=1)
df = df.rename(columns={'Cliente':'NOME_CLIENTE','Posição':'Custódia','S':'Suitability','I':'Tipo de Investidor'})


# Criar a condição para a coluna "Sub Produto"
df['Suitability'] = df.apply(lambda x: 'Agressivo' if (x['Suitability'] == 'A') else x['Suitability'], axis=1)
df['Suitability'] = df.apply(lambda x: 'Não Preenchido' if (x['Suitability'] == 'N') else x['Suitability'], axis=1)
df['Suitability'] = df.apply(lambda x: 'Desatualizado' if (x['Suitability'] == 'D') else x['Suitability'], axis=1)
df['Suitability'] = df.apply(lambda x: 'Moderado' if (x['Suitability'] == 'M') else x['Suitability'], axis=1)
df['Suitability'] = df.apply(lambda x: 'Conservador' if (x['Suitability'] == 'C') else x['Suitability'], axis=1)

df['Tipo de Investidor'] = df.apply(lambda x: 'Profissional' if (x['Tipo de Investidor'] == 'P') else x['Tipo de Investidor'], axis=1)
df['Tipo de Investidor'] = df.apply(lambda x: 'Qualificado' if (x['Tipo de Investidor'] == 'Q') else x['Tipo de Investidor'], axis=1)
df['Tipo de Investidor'] = df.apply(lambda x: 'Regular' if (x['Tipo de Investidor'] == 'R') else x['Tipo de Investidor'], axis=1)


novo_nome3 = f'saldo-em-conta.xlsx'
filepath = os.path.join(PASTA_SALDO_EM_CONTA, novo_nome3)
df_saldo = pd.read_excel(filepath, header=1)


df_saldo['Cliente'] = df_saldo['Cliente'].astype(str)
df = pd.merge(df, df_saldo[['Cliente','Assessor']], left_on='COD_CLIENTE', right_on='Cliente', how='left')
df = df.drop(['Cliente'], axis=1)






# In[12]:


arquivo_ativacao_consolidado =r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Ativacao\Ativacao Consolidado.xlsx'

PASTA_MTM_AGIO_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Renda Fixa HUB\MTM Agio\Historico'

arquivos = glob.glob(os.path.join(PASTA_MTM_AGIO_HISTORICO, 'Renda_Fixa_MTM_Ágio_*'))
arquivos.sort(key=os.path.getmtime)
filename = arquivos[0]

# Extrai o valor após o primeiro sublinhado
nome_arquivo = os.path.basename(filename)
valor = nome_arquivo.split('_')[4]
valor = valor.split('_')[0]
valor = valor.split('.xlsx')[0]
valor = valor.replace('-', '')

valor = valor[6:] + "/" + valor[4:6] + "/" + valor[:4]

valor2 = valor[:2] + "/" + valor[3:5] + "/" + valor[6:]

df_rf1 = pd.read_excel(filename)

df_rf1['Codigo Cliente'] = df_rf1['Codigo Cliente'].astype(str).apply(lambda x: x.split('.0')[0])
df['COD_CLIENTE'] = df['COD_CLIENTE'].astype(str).apply(lambda x: x.split('.0')[0])
df_rf1['Código Assessor'] = df_rf1['Código Assessor'].astype(str).apply(lambda x: x.split('.0')[0])
df_rf1['Disponível'] = df_rf1['Disponível'].astype(str).apply(lambda x: x.split('.0')[0])

import math


def round_up(value):
    if pd.isna(value):
        return value
    else:
        return math.ceil(value * 100) / 100

df_rf1['Disponível'] = df_rf1['Disponível'].astype(float).apply(round_up)

df_rf1['Disponível'] = df_rf1['Disponível'].astype(str).apply(lambda x: x.split('.0')[0])


#df['Qtd'] = df['Qtd'].astype(float).apply(round_up)

df['COD_ASSESSOR'] = df['COD_ASSESSOR'].astype(str).apply(lambda x: x.split('.0')[0])

df['Qtd'] = df['Qtd'].astype(str).apply(lambda x: x.split('.0')[0])

df_rf1['ID3'] = df_rf1['Codigo Cliente'].astype(str) + '_' + df_rf1['Ticker'].astype(str) + '_' + df_rf1['Código Assessor'].astype(str) + '_' + df_rf1['Disponível'].astype(str)

df['ID2'] = df['COD_CLIENTE'].astype(str) + '_' + df['COD_ATIVO'].astype(str) + '_' + df['COD_ASSESSOR'].astype(str) + '_' + df['Qtd'].astype(str)

teste = f'df.xlsx'     


rent_fundos = os.path.join(PASTA_DOWNLOADS, teste)


df.to_excel(rent_fundos, index=False)

teste = f'df_rf1.xlsx'     


rent_fundos = os.path.join(PASTA_DOWNLOADS, teste)


df_rf1.to_excel(rent_fundos, index=False)



# In[13]:


df = pd.merge(df, df_rf1,
              left_on='ID2', right_on='ID3', how='left')


columns_to_drop = [
    'Código Assessor',
    'Data Aplicação',
    'Codigo Cliente',
    'PF/ PJ',
    'Perfil Suitability',
    'Ticker',
    'Ativo_y',
    'Tipo Ativo',
    'Vencimento_y',
    'Classe',
    'Emissor_y',
    'Disponível',
    'Escritório',
    'ID3'
]
df = df.drop(columns=columns_to_drop)

df.rename(columns={'Ativo_x': 'Ativo',
                        'Vencimento_x': 'Vencimento',
                        'Indexador_x':'Indexador',
                        'Emissor_x':'Emissor',                   
                       }, inplace=True)



df_ativacao = pd.read_excel(arquivo_ativacao_consolidado)
df_ativacao['Cliente'] = df_ativacao['Cliente'].astype(str).apply(lambda x: x.split('.0')[0])
df = pd.merge(df, df_ativacao[['Cliente','MES_ATIVACAO','Ativacao 300k?']],
              left_on='COD_CLIENTE', right_on='Cliente', how='left')


def ajustar_ativo(row):
    if row['Produto'] == 'Fundo Imobiliário':
        return row['Ativo'][-6:]
    else:
        return row['Ativo']

# Aplicando a função ajustar_ativo a cada linha do DataFrame
df['Ativo'] = df.apply(ajustar_ativo, axis=1)

df['ID'] = df['COD_CLIENTE'].astype(str) + '_' + df['Ativo'].astype(str)


# In[14]:


df_net_total = df.groupby('COD_CLIENTE')['NET Modulo Posicao'].sum().reset_index()

df = pd.merge(df, df_net_total[['COD_CLIENTE',
                                      'NET Modulo Posicao',
                                      ]],
              left_on='COD_CLIENTE', right_on='COD_CLIENTE', how='left')




df = df.rename(columns={'NET Modulo Posicao_x':'NET Modulo Posicao','NET Modulo Posicao_y':'NET Total'
                        })




df['NET Modulo Posicao'] = df['NET Modulo Posicao'].astype(float)
df['NET Total'] = df['NET Total'].astype(float)

df['NET Liquidez'] = np.where(df['Class_Inove'] == 'Liquidez', df['NET Modulo Posicao'], 0)
df['%_Liquidez'] = np.where(df['Class_Inove'] == 'Liquidez', df['NET Liquidez'] / df['NET Total'] * 100,0)
df['%_Liquidez'] = df['%_Liquidez'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Inflação'] = np.where(df['Class_Inove'] == 'Inflação', df['NET Modulo Posicao'], 0)
df['%_Inflação'] = np.where(df['Class_Inove'] == 'Inflação', df['NET Inflação'] / df['NET Total'] * 100,0)
df['%_Inflação'] = df['%_Inflação'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Pós-Fixado'] = np.where(df['Class_Inove'] == 'Pós-Fixado', df['NET Modulo Posicao'], 0)
df['%_Pós-Fixado'] = np.where(df['Class_Inove'] == 'Pós-Fixado', df['NET Pós-Fixado'] / df['NET Total'] * 100,0)
df['%_Pós-Fixado'] = df['%_Pós-Fixado'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Pré-Fixado'] = np.where(df['Class_Inove'] == 'Pré-Fixado', df['NET Modulo Posicao'], 0)
df['%_Pré-Fixado'] = np.where(df['Class_Inove'] == 'Pré-Fixado', df['NET Pré-Fixado'] / df['NET Total'] * 100,0)
df['%_Pré-Fixado'] = df['%_Pré-Fixado'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Renda Variável'] = np.where(df['Class_Inove'] == 'Renda Variável', df['NET Modulo Posicao'], 0)
df['%_Renda Variável'] = np.where(df['Class_Inove'] == 'Renda Variável', df['NET Renda Variável'] / df['NET Total'] * 100,0)
df['%_Renda Variável'] = df['%_Renda Variável'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Fundo Imobiliário'] = np.where(df['Class_Inove'] == 'Fundo Imobiliário', df['NET Modulo Posicao'], 0)
df['%_Fundo Imobiliário'] = np.where(df['Class_Inove'] == 'Fundo Imobiliário', df['NET Fundo Imobiliário'] / df['NET Total'] * 100,0)
df['%_Fundo Imobiliário'] = df['%_Fundo Imobiliário'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Previdência'] = np.where(df['Produto'] == 'Previdência', df['NET Modulo Posicao'], 0)
df['%_Previdência'] = np.where(df['Produto'] == 'Previdência', df['NET Previdência'] / df['NET Total'] * 100,0)
df['%_Previdência'] = df['%_Previdência'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Alternativo'] = np.where(df['Class_Inove'] == 'Alternativo', df['NET Modulo Posicao'], 0)
df['%_Alternativo'] = np.where(df['Class_Inove'] == 'Alternativo', df['NET Alternativo'] / df['NET Total'] * 100,0)
df['%_Alternativo'] = df['%_Alternativo'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Internacional'] = np.where(df['Class_Inove'] == 'Internacional', df['NET Modulo Posicao'], 0)
df['%_Internacional'] = np.where(df['Class_Inove'] == 'Internacional', df['NET Internacional'] / df['NET Total'] * 100,0)
df['%_Internacional'] = df['%_Internacional'].apply(lambda x: '{:.2f}%'.format(x))

df['NET Multimercado'] = np.where(df['Class_Inove'] == 'Multimercado', df['NET Modulo Posicao'], 0)
df['%_Multimercado'] = np.where(df['Class_Inove'] == 'Multimercado', df['NET Multimercado'] / df['NET Total'] * 100,0)
df['%_Multimercado'] = df['%_Multimercado'].apply(lambda x: '{:.2f}%'.format(x))


condicao = (df['Class_Inove'] == 'Renda Variável') & (df['Produto'] == 'Fundos')

df['NET FIA RV'] = np.where(condicao, df['NET Modulo Posicao'], 0)
df['%_FIA RV'] = np.where(condicao, df['NET FIA RV'] / df['NET Total'] * 100, 0)
df['%_FIA RV'] = df['%_FIA RV'].apply(lambda x: '{:.2f}%'.format(x))

condicao = (df['Class_Inove'] == 'Renda Variável') & (df['Produto'] != 'Fundos')

df['NET RV (exc FIA)'] = np.where(condicao, df['NET Modulo Posicao'], 0)
df['%_RV exc FIA'] = np.where(condicao, df['NET RV (exc FIA)'] / df['NET Total'] * 100, 0)
df['%_RV exc FIA'] = df['%_RV exc FIA'].apply(lambda x: '{:.2f}%'.format(x))

condicao = (df['Produto'] == 'Renda Fixa') & (df['Sub Produto'] == 'Crédito Privado')

df['NET CP'] = np.where(condicao, df['NET Modulo Posicao'], 0)
df['%_CP'] = np.where(condicao, df['NET CP'] / df['NET Total'] * 100, 0)
df['%_CP'] = df['%_CP'].apply(lambda x: '{:.2f}%'.format(x))



### Ao INCLUIR ALGUMA PORCENTAGEM AQUI, LEMBRAR DE INSERIR A COLUNA NA LINHA 105 da CELULA ABAIXO e NA CELULA DO ENQUADRAMENTO
df['COD_CLIENTE'] = df['COD_CLIENTE'].astype(str).apply(lambda x: x.split('.')[0])

novo_nome3 = f'dAssessores.xls'
filepath = os.path.join(PASTA_ARQUIVOS_PADROES, novo_nome3)
df_assessores = pd.read_excel(filepath)

# Substituir 'nan%' por string vazia nas colunas específicas
cols = ['Margem Total (Margem Taxa Adm + Margem Taxa Perf)', 'Taxa de Performance x Comissão***', 'Taxa de Administração x Comissão', 'Taxa de Performance', 'Taxa de Administração']

for col in cols:
    df[col] = df[col].replace('nan%', '')
       
df['COD_ASSESSOR'] = df['COD_ASSESSOR'].astype(str).apply(lambda x: x.split('.0')[0])
    
df_assessores['CODIGO_INOVE_USUARIO'] = df_assessores['CODIGO_INOVE_USUARIO'].astype(str)
df = pd.merge(df, df_assessores[['CODIGO_INOVE_USUARIO','NOME_PIPEDRIVE','EQUIPE_COMERCIAL']], left_on='COD_ASSESSOR', right_on='CODIGO_INOVE_USUARIO', how='left')

df = df.drop(['CODIGO_INOVE_USUARIO'], axis=1)
df = df.rename(columns={'NOME_PIPEDRIVE':'NOME_ASSESSOR'})

cols_to_replace = ['%_Previdência',
                   '%_Alternativo',
                   '%_Internacional',
                   '%_Multimercado',
                   '%_Renda Variável',
                   '%_Fundo Imobiliário',
                   '%_Pré-Fixado',
                   '%_Pós-Fixado',
                   '%_Inflação',
                   '%_Liquidez',
                   '%_FIA RV',
                   '%_RV exc FIA',
                   '%_CP',
                   'Taxa',
                   'Margem Total (Margem Taxa Adm + Margem Taxa Perf)', 
                   'Taxa de Performance x Comissão***', 
                   'Taxa de Administração x Comissão', 
                   'Taxa de Administração']

df[cols_to_replace] = df[cols_to_replace].replace('.', ',') 

df['Produto'] = df.apply(lambda x: 'Alternativo' if (x['Produto'] == 'Alternative' or x['Produto'] == 'Cripto') else x['Produto'], axis=1)

df['COD_ASSESSOR'] = df['COD_ASSESSOR'].astype(str).apply(lambda x: x.split('.0')[0])

df['Produto'] = df.apply(lambda x: 'Fundos' if (x['Sub Produto'] == 'Mutual Fund') else x['Produto'], axis=1)



# In[15]:


### Adiciona informações de FII

PASTA_FII = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Fundo Imobiliário HUB'
PASTA_DIVERSIFICACAO_INOVE_CONSOL = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao Inove\Consolidado'
PASTA_DIVERSIFICACAO_INOVE_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao Inove\Consolidado\Historico'  
PASTA_DIVERSIFICACAO_INOVE_FECHAMENTO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao Inove\Consolidado\Fechamento'  

# Abra o arquivo em um DataFrame usando o caminho completo do arquivo
caminho_completo = os.path.join(PASTA_FII, f'FII_Visao_Geral.xlsx')
df_fii = pd.read_excel(caminho_completo)

# Criar uma nova coluna com os últimos 6 dígitos da coluna "Ativo"
df_fii['ativo_final'] = df_fii['Ativo'].str[-6:]


df_fii['ID'] = df_fii['Cliente'].astype(str) + '_' + df_fii['ativo_final'].astype(str)
# Remover as colunas "COD_ATIVO_FIP" e "COD_ATIVO_2"
df_fii = df_fii.drop(['Escritório', 'Assessor','Ativo'], axis=1)
df_fii = df_fii.rename(columns={'ativo_final':'Ativo',
                    'Nominal (R$)':'NET HUB FII',
                    'Qt Ativo':'Qtd Cotas',
                    })

df = pd.merge(df, df_fii,
              left_on='ID', right_on='ID', how='left')

df.rename(columns={'Ativo_x': 'Ativo',              
                       }, inplace=True)



# Remover as colunas "Ativo" e "Ativo2"
    #df = df.drop(['Ativo', 'Ativo2'], axis=1)
    
# Caminho para o ChromeDriver
path_to_chromedriver = '/path/to/chromedriver' 

# Criar uma nova instância do navegador Google Chrome
driver = webdriver.Chrome(executable_path=path_to_chromedriver)

# Ir para a página que queremos raspar
driver.get('https://www.fundsexplorer.com.br/ranking')

# Esperar um pouco para a página carregar
driver.implicitly_wait(10)

# Pegar o HTML da página e passar para o BeautifulSoup
soup = BeautifulSoup(driver.page_source, 'html.parser')

# Encontrar a tabela
table = soup.find('table')

# Raspar a tabela
df_dados_fii = pd.read_html(str(table),decimal=',',thousands='.')[0]

# Fechar o navegador
driver.quit()


df = pd.merge(df, df_dados_fii,
          left_on='Ativo', right_on='Fundos', how='left')

df = df.drop(['Fundos'], axis=1)

df["Ativo"] = df["Ativo"].astype(str)
df["Emissor"] = df["Emissor"].astype(str)
df["Indexador/Benchmark"] = df["Indexador/Benchmark"].astype(str)

# criar uma nova coluna com as colunas concatenadas
df["Ativo Resumido"] = df.apply(lambda row: row["Ativo"] + " - " + row["Emissor"] + " - " + row["Indexador/Benchmark"], axis=1)

arquivo_positivador = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Positivador\Positivador.xlsx'

df_positivador = pd.read_excel(arquivo_positivador)
df_positivador['Cliente'] = df_positivador['Cliente'].astype(str).apply(lambda x: x.split('.0')[0])
df = pd.merge(df, df_positivador,
              left_on='COD_CLIENTE', right_on='Cliente', how='left')


df = df.drop(['Indexador_y', 'Cliente_x','Cliente_y','Ativo_y','Assessor_y'], axis=1)
df = df.rename(columns={'Data_Ref_x':'Data_Ref',
                    'Data_Ref_y':'Data_Ref_Positivador',
                    'Assessor_x':'Assessor',
                        'MES_ATIVACAO':'Mes_Ativacao',
                        'Segmento_x':'Segmento FII',
                        'Segmento_y':'Segmento_atendimento'
                    })


df['%_Fundo Imobiliário2'] = df['%_Fundo Imobiliário']
df['%_Fundo Imobiliário2'] = df['%_Fundo Imobiliário2'].str.replace('%', '')
df['%_Fundo Imobiliário2'] = df['%_Fundo Imobiliário2'].astype(float)


df['%_Liquidez2'] = df['%_Liquidez']
df['%_Liquidez2'] = df['%_Liquidez2'].str.replace('%', '')
df['%_Liquidez2'] = df['%_Liquidez2'].astype(float)

df['%_Inflação2'] = df['%_Inflação']
df['%_Inflação2'] = df['%_Inflação2'].str.replace('%', '')
df['%_Inflação2'] = df['%_Inflação2'].astype(float)

df['%_Pós-Fixado2'] = df['%_Pós-Fixado']
df['%_Pós-Fixado2'] = df['%_Pós-Fixado2'].str.replace('%', '')
df['%_Pós-Fixado2'] = df['%_Pós-Fixado2'].astype(float)

df['%_Pré-Fixado2'] = df['%_Pré-Fixado']
df['%_Pré-Fixado2'] = df['%_Pré-Fixado2'].str.replace('%', '')
df['%_Pré-Fixado2'] = df['%_Pré-Fixado2'].astype(float)

df['%_Renda Variável2'] = df['%_Renda Variável']
df['%_Renda Variável2'] = df['%_Renda Variável2'].str.replace('%', '')
df['%_Renda Variável2'] = df['%_Renda Variável2'].astype(float)

df['%_Previdência2'] = df['%_Previdência']
df['%_Previdência2'] = df['%_Previdência2'].str.replace('%', '')
df['%_Previdência2'] = df['%_Previdência2'].astype(float)

df['%_Alternativo2'] = df['%_Alternativo']
df['%_Alternativo2'] = df['%_Alternativo2'].str.replace('%', '')
df['%_Alternativo2'] = df['%_Alternativo2'].astype(float)

df['%_Internacional2'] = df['%_Internacional']
df['%_Internacional2'] = df['%_Internacional2'].str.replace('%', '')
df['%_Internacional2'] = df['%_Internacional2'].astype(float)

df['%_Multimercado2'] = df['%_Multimercado']
df['%_Multimercado2'] = df['%_Multimercado2'].str.replace('%', '')
df['%_Multimercado2'] = df['%_Multimercado2'].astype(float)

df['%_Multimercado2'] = df['%_Multimercado']
df['%_Multimercado2'] = df['%_Multimercado2'].str.replace('%', '')
df['%_Multimercado2'] = df['%_Multimercado2'].astype(float)

df['%_FIA RV2'] = df['%_FIA RV']
df['%_FIA RV2'] = df['%_FIA RV2'].str.replace('%', '')
df['%_FIA RV2'] = df['%_FIA RV2'].astype(float)

df['%_RV exc FIA2'] = df['%_RV exc FIA']
df['%_RV exc FIA2'] = df['%_RV exc FIA2'].str.replace('%', '')
df['%_RV exc FIA2'] = df['%_RV exc FIA2'].astype(float)

df['%_CP2'] = df['%_CP']
df['%_CP2'] = df['%_CP2'].str.replace('%', '')
df['%_CP2'] = df['%_CP2'].astype(float)




# Agrupar por 'Cliente' e calcular a soma de '%_Fundo_Imobiliário'
df_grouped = df.groupby('COD_CLIENTE')['%_Fundo Imobiliário2',
                                   '%_Liquidez2',
                                   '%_Inflação2',
                                   '%_Pós-Fixado2',
                                   '%_Pré-Fixado2',
                                   '%_Renda Variável2',
                                   '%_Previdência2',
                                   '%_Alternativo2',
                                   '%_Internacional2',
                                   '%_Multimercado2',
                                   '%_FIA RV2',
                                    '%_RV exc FIA2',
                                      '%_CP2'].sum().reset_index()

# Renomear a coluna para evitar conflitos durante a operação de merge
df_grouped = df_grouped.rename(columns={'%_Fundo Imobiliário2': 'Total %_Fundo Imobiliário',
                                        '%_Liquidez2':'Total %_Liquidez',
                                        '%_Inflação2':'Total %_Inflação',
                                        '%_Pós-Fixado2':'Total %_Pós-Fixado',
                                        '%_Pré-Fixado2':'Total %_Pré-Fixado',
                                       '%_Renda Variável2':'Total %_Renda Variável',
                                       '%_Previdência2':'Total %_Previdência',
                                       '%_Alternativo2':'Total %_Alternativo',
                                       '%_Internacional2':'Total %_Internacional',
                                       '%_Multimercado2':'Total %_Multimercado',
                                       '%_FIA RV2':'Total %_FIA RV',
                                        '%_RV exc FIA2':'Total %_RV exc FIA',
                                        '%_CP2':'Total %_CP'
                                       
                                       
                                       
                                       
                                       
                                       
                                       })
df = pd.merge(df, df_grouped,
              left_on='COD_CLIENTE', right_on='COD_CLIENTE', how='left')





PASTA_NET = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\NET\Historico'

arquivos = glob.glob(os.path.join(PASTA_NET, 'NET_*'))
arquivos.sort(key=os.path.getmtime)
filename = arquivos[0]
df_net = pd.read_excel(filename)
df_net['Conta XP'] = df_net['Conta XP'].astype(str).apply(lambda x: x.split('.0')[0])
df = pd.merge(df, df_net,
              left_on='COD_CLIENTE', right_on='Conta XP', how='left')

novo_nome = f'Diversificacao_Inove_{formatted_date}.xlsx'
novo_nome2 = f'Diversificacao_Inove_{formatted_date2}.xlsx'
novo_nome3 = f'Diversificacao_Inove.xlsx'
# Esperar um pouco mais antes de renomear o arquivo
time.sleep(1)
path_to_file = os.path.join(PASTA_DIVERSIFICACAO_INOVE_CONSOL, f'Diversificacao_Inove.xlsx')
# Cria um ID único com todas as colunas
df['ID'] = df.apply(lambda row: '_'.join(row.values.astype(str)), axis=1)

# Remove as linhas duplicadas
df = df.drop_duplicates()
df.to_excel(path_to_file)
shutil.copy(os.path.join(PASTA_DIVERSIFICACAO_INOVE_CONSOL, novo_nome3), os.path.join(PASTA_DIVERSIFICACAO_INOVE_HISTORICO, novo_nome))
shutil.copy(os.path.join(PASTA_DIVERSIFICACAO_INOVE_CONSOL, novo_nome3), os.path.join(PASTA_DIVERSIFICACAO_INOVE_FECHAMENTO, novo_nome2))





# In[ ]:






# In[ ]:





# In[16]:


#CONSOLIDADO
import pyautogui
import time
import os
import datetime
import pandas as pd
import openpyxl
import re
import shutil
import locale
import glob
import win32com.client
import email
import numpy as np
import keyboard
import pytz
import requests
import numpy
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import FORMULAE

from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


# Marque o tempo de início
start_time = time.time()
print("Início Custodia OnShore")


# Configurar locale para usar separador de milhar "."
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
tz = pytz.timezone('America/Sao_Paulo')

# Set up constants

PASTA_DOWNLOADS = r'C:\Users\dados.100486\Downloads'
PASTA_DIVERSIFICACAO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao'
PASTA_MI_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Mercado Internacional\Historico'
PASTA_RF = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Vencimento RF Connect'
PASTA_DIVERSIFICACAO_ATUAL = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao\Historico'
VENC_RF_COMPLETO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Vencimento RF Connect\Completo'

arquivo_guia_fundos = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões\dGuiaConsol.xlsx'
arquivo_bd_fundos = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões\dBD.xlsx'


PASTA_ARQUIVOS_PADROES = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Arquivos Padrões'

PASTA_CAMPANHA_ATIVACAO = r'C:\Users\Douglas\XP Investimentos\Bruno Martins - Power BI\Base Dados\Positivador\Campanha Ativação'

PASTA_DIVERSIFICACAO_HISTORICO = r'C:\Users\Douglas\XP Investimentos\Bruno Martins - Power BI\Base Dados\Diversificacao\Historico Consolidado'


PAUSE_TIME = 2
SHORT_PAUSE_TIME = 1


# Set up PyAutoGUI failsafe
pyautogui.FAILSAFE = True

# Define helper functions
def mover_mouse(x, y):
    pyautogui.moveTo(x, y)

def clicar_pagina(x, y, pause=PAUSE_TIME):
    mover_mouse(x, y)
    time.sleep(pause)
    pyautogui.click()

def clicar_pagina_3x(x, y, pause=PAUSE_TIME):
    mover_mouse(x, y)
    time.sleep(pause)
    pyautogui.tripleClick()
    
def format_string_percentage(value):
    try:
        number = float(value)
        return f"{number:.2%}"
    except ValueError:
        return value

def extract_indexador(row):
    value = row['Indexador/Benchmark']
    if isinstance(value, str):
        if '% CDI' in value:
            return '% CDI'
        elif '% IPC-A' in value:
            return '% IPC-A'
        elif '% IGP-M' in value:
            return '% IGP-M'
        elif 'CDI +' in value:
            return 'CDI +'
        elif 'DOLAR PTAX +' in value:
            return 'DOLAR PTAX +'
        elif 'IGP-M +' in value:
            return 'IGP-M +'
        elif 'IPC-A +' in value:
            return 'IPC-A +'
        elif 'LFT +' in value:
            return 'LFT +'
        elif 'LFT' in value:
            return 'LFT'
    else:
        return 'Pré'
    
def extract_taxa(row):
    value = row['Indexador/Benchmark']
    if isinstance(value, str):
        if '% CDI' in value:
            return value.split()[0]
        elif '% IPC-A' in value:
            return value.split()[0]
        elif '% IGP-M' in value:
            return value.split()[0]
        elif 'CDI +' in value:
            return value.split('+')[1].strip()
        elif 'DOLAR PTAX +' in value:
            return value.split('DOLAR PTAX +')[1].strip()
        elif 'IGP-M +' in value:
            return value.split('+')[1].strip()
        elif 'IPC-A +' in value:
            return value.split('+')[1].strip()
        elif 'LFT +' in value:
            return value.split('+')[1].strip()
        elif 'LFT' in value:
            return '0%'
    else:
        return value
    
# definir uma função que aplica as transformações necessárias na coluna Cod_Ativo
def transforma_cod_ativo(cod_ativo):
    if cod_ativo.startswith('PRE DU'):
        # extrair o conteúdo que vem após "PRE DU"
        novo_cod_ativo = re.search(r'PRE DU (.*)', cod_ativo).group(1)
    elif cod_ativo.startswith('PRE'):
        # extrair o conteúdo que vem após "PRE"
        novo_cod_ativo = re.search(r'PRE (.*)', cod_ativo).group(1)
    elif cod_ativo.startswith('FLU'):
        # extrair o conteúdo que vem após "FLU U"
        novo_cod_ativo = re.search(r'FLU U (.*)|FLU (.*)', cod_ativo)
        novo_cod_ativo = novo_cod_ativo.group(1) if novo_cod_ativo.group(1) else novo_cod_ativo.group(2)
    else:
        # manter o valor original
        novo_cod_ativo = cod_ativo
    return novo_cod_ativo

def aplicar_formula(coluna):
    if coluna == "" or coluna == "Total":
        return ""
    elif coluna in ["Bond", "Unknown", "Equity", "Stock", "Cash", "USD", "REIT", "ETF", "EUR", "Alternative", "Balanced", "Semi-liquid"]:
        return df.loc[df['Top Level Owner'] == coluna, 'Valor Total'].values[0]
    else:
        try:
            pos_espaco = coluna.index(" ")
            return df.loc[df['Top Level Owner'] == coluna, 'Valor Total'].values[0]
        except ValueError:
            return coluna

        
        
#####


# Get the current date and time in the desired format
formatted_date = datetime.now().strftime("%Y%m%d_%H%M%S")
formatted_date2 = datetime.now().strftime("%Y%m")
time.sleep(5)



PASTA_DIVERSIFICACAO_INOVE_ENQ = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao Inove\Enquadramento'
PASTA_DIVERSIFICACAO_INOVE_ENQ_HISTORICO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao Inove\Enquadramento\Historico'  
PASTA_DIVERSIFICACAO_INOVE_ENQ_FECHAMENTO = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao Inove\Enquadramento\Fechamento'
PASTA_DIVERSIFICACAO_INOVE_CONSOL = r'C:\Users\dados.100486\OneDrive - XP Investimentos\Tech Inove\Base Dados\Diversificacao Inove\Consolidado'


path_to_file = os.path.join(PASTA_DIVERSIFICACAO_INOVE_CONSOL, f'Diversificacao_Inove.xlsx')
df= pd.read_excel(path_to_file)

df_enquadramento = df.drop(['Origem',
                            'Data_Ref',
                            'Produto',
                            'Sub Produto',
                            'Class_Inove', 
                            'Ativo', 'COD_ATIVO', 
                            'Emissor', 
                            'Indexador/Benchmark',
                            'Dt_Aplicacao', 
                            'Vencimento', 
                            'PU_ATUAL', 
                            'NET_Original', 
                            'Liquidez', 
                            'CNPJ Fundo',
                            'Margem Total (Margem Taxa Adm + Margem Taxa Perf)', 'Taxa de Performance x Comissão***', 'Taxa de Administração x Comissão', 'Taxa de Performance', 'Taxa de Administração', 'CVM', 'Anbima', 'Classificação XP', 'Gestor', 'Administrador', 'Custodiante'], axis=1)

df_enquadramento = df_enquadramento.groupby(['COD_CLIENTE']).agg(
    {'COD_ASSESSOR': 'first',
     'NET Total': 'first',
     'NOME_CLIENTE':'first',
     'Suitability':'first',
     'Tipo de Investidor':'first',
     'Aplicação Financeira Declarada':'first',
     'Sexo':'first',
     'Custódia':'first',
     'NOME_ASSESSOR':'first',
     'NET Liquidez': 'sum',
     'NET Inflação': 'sum', 
     'NET Pós-Fixado': 'sum', 
     'NET Pré-Fixado': 'sum', 
     'NET Renda Variável': 'sum',
     'NET RV (exc FIA)': 'sum',
     'NET FIA RV': 'sum',
     'NET Fundo Imobiliário': 'sum', 
     'NET Previdência': 'sum', 
     'NET Alternativo': 'sum',
     'NET Multimercado': 'sum', 
     'NET Internacional': 'sum'}).reset_index()


df_enquadramento['%_Liquidez'] = df_enquadramento['NET Liquidez'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Inflação'] = df_enquadramento['NET Inflação'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Pós-Fixado'] = df_enquadramento['NET Pós-Fixado'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Pré-Fixado'] = df_enquadramento['NET Pré-Fixado'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Renda Variável'] = df_enquadramento['NET Renda Variável'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_FIA RV'] = df_enquadramento['NET FIA RV'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_RV exc FIA'] = df_enquadramento['NET RV (exc FIA)'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Fundo Imobiliário'] = df_enquadramento['NET Fundo Imobiliário'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Previdência'] = df_enquadramento['NET Previdência'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Alternativo'] = df_enquadramento['NET Alternativo'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Internacional'] = df_enquadramento['NET Internacional'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Alternativo'] = df_enquadramento['NET Alternativo'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Internacional'] = df_enquadramento['NET Internacional'] / df_enquadramento['NET Total'] * 100
df_enquadramento['%_Multimercado'] = df_enquadramento['NET Multimercado'] / df_enquadramento['NET Total'] * 100

novo_nome = f'Enquadramento_Inove_{formatted_date}.xlsx'
novo_nome2 = f'Enquadramento_Inove_{formatted_date2}.xlsx'
novo_nome3 = f'Enquadramento_Inove.xlsx'
# Esperar um pouco mais antes de renomear o arquivo
time.sleep(1)
path_to_file = os.path.join(PASTA_DIVERSIFICACAO_INOVE_ENQ, f'Enquadramento_Inove.xlsx')
df_enquadramento.to_excel(path_to_file)
shutil.copy(os.path.join(PASTA_DIVERSIFICACAO_INOVE_ENQ, novo_nome3), os.path.join(PASTA_DIVERSIFICACAO_INOVE_ENQ_HISTORICO, novo_nome))
shutil.copy(os.path.join(PASTA_DIVERSIFICACAO_INOVE_ENQ, novo_nome3), os.path.join(PASTA_DIVERSIFICACAO_INOVE_ENQ_FECHAMENTO, novo_nome2))


print("Fim")
# Marque o tempo de término
end_time = time.time()

# Calcule e imprima a duração
duration = end_time - start_time
print(f"Tempo de execução: {duration:.2f} segundos")


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




