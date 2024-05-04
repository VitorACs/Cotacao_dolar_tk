#-*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkcalendar import DateEntry
import pandas as pd
import requests
from datetime import datetime
import numpy as np

#fazer uma lista direta para as cotações
requesicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = list(requesicao.json())

#funções 

def pegar_cotacao():
    moeda = lista2.get()
    data = lista3.get()
    ano = data[-4:]
    mes = data[3:5]
    dia = data[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    resposta1['text'] = f"A cotação da {moeda} no dia {data} foi de: R${valor_moeda}"

def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o arquivo excel')
    var_caminhoarquivo.set(caminho_arquivo)
    try:
        df = pd.read_excel(caminho_arquivo)
        texto = tk.Label(text='arquivo selecionado',bg='black',fg='white',justify='center',width=35,height=2)
        texto.grid(row=6,column=2)
    except:
        texto = tk.Label(text='arquivo não selecionado',bg='black',fg='white',justify='center',width=35,height=2)
        texto.grid(row=6,column=2)

def varias_cotacoes():
    try:
        # ler o dataframe de moedas
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        # pegar a data de inicio e data de fim das cotacoes
        data_inicial = lista6.get()
        data_final = lista7.get()
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?" \
                   f"start_date={ano_inicial}{mes_inicial}{dia_inicial}&" \
                   f"end_date={ano_final}{mes_final}{dia_final}"

            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                print(data)
                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
        df.to_excel("Teste.xlsx")
        resposta2['text'] = "Arquivo Atualizado com Sucesso"
    except:
        resposta2['text'] = "Selecione um arquivo Excel no Formato Correto"
    



# janela do programa
janela = tk.Tk()
janela.title('Cotação de Moedas')
#janela.geometry('400x500')
janela.configure(bg='black')

texto1 = tk.Label(text='Cotação para uma moeda',bg='black',fg='yellow',justify='center',width=35,height=1)
texto1.configure(font=('arial', 16,'bold'))
texto1.grid(row=0,column=0,columnspan=3,sticky='EW')

texto2 = tk.Label(text='Selecione a moeda que deseja consultar',bg='black',fg='White',height=2)
texto2.grid(row=1,column=0,columnspan=2)
lista2 = ttk.Combobox(janela,values=dicionario_moedas) # definir a lista 1 
lista2.grid(row=1,column=2,padx=10,pady=10,sticky='nsew')

texto3 = tk.Label(text='Selecione o dia que deseja consultar',bg='black',fg='White',height=2)
texto3.grid(row=2,column=0,columnspan=2)
lista3 = DateEntry(year=2023,locale='pt_br') 
lista3.grid(row=2,column=2,padx=10,pady=10,sticky='nsew')

resposta1 = tk.Label(text='',bg='black',fg='White',height=2)
resposta1.grid(row=3,column=0,columnspan=2)

botao1 = tk.Button(text='Pegar cotações',command=pegar_cotacao)
botao1.grid(row=3,column=2,sticky='ew',padx=10,pady=10)

texto4 = tk.Label(text='Cotação para mais de uma moeda',bg='black',fg='yellow',justify='center',width=35,height=4)
texto4.configure(font=('arial', 11,'bold'))
texto4.grid(row=4,column=0,columnspan=3,sticky='EW')

texto5 = tk.Label(text='Selecione o arquivo excel para cotação de várias moedas',bg='black',fg='White',height=3)
texto5.grid(row=5,column=0,columnspan=2)
botao2 = tk.Button(text='Pegar arquivo excel',command=selecionar_arquivo)
botao2.grid(row=5,column=2,sticky='ew',padx=10,pady=10)

var_caminhoarquivo = tk.StringVar()

texto6 = tk.Label(text='Data de inicio:',bg='black',fg='white',justify='center',width=30,height=2)
texto6.grid(row=7,column=0)
lista6 = DateEntry(year=2023,locale='pt_br') 
lista6.grid(row=7,column=1,padx=10,pady=10,sticky='nsew')

texto7 = tk.Label(text='Data Final:',bg='black',fg='white',justify='center',width=30,height=2)
texto7.grid(row=8,column=0)
lista7 = DateEntry(year=2023,locale='pt_br') 
lista7.grid(row=8,column=1,padx=10,pady=10,sticky='nsew')

botao3 = tk.Button(text='Atualizar cotações',command=varias_cotacoes)
botao3.grid(row=9,column=0,padx=10,pady=10,sticky='ew')

resposta2 = tk.Label(text='',bg='black',fg='white',justify='center',width=30,height=2)
resposta2.grid(row=9,column=1,columnspan=2,sticky='ew',padx=10,pady=10)


janela.mainloop()