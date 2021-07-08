from PySimpleGUI.PySimpleGUI import VerticalSeparator
import pandas as pd
import PySimpleGUI as sg
import json
import os
import pyautogui as py
import numpy as np
import time
import datetime
import numpy as np

time.time()
df = pd.read_excel('sec753-00.xlsx',engine="openpyxl")
df2 = pd.read_excel('sec451-00.xlsx',engine="openpyxl")
lista_df = [df,df2]
lista_df[0]['np'][0]
dic_df = {'01':df,
        '02':df2}
dic_df['01']['np']

key = ['01','02']
dic2 = dict(zip(key,lista_df))

dict_test = {}

dict_test.update(dic_df)

len(dic_df['01'])

for i in dic_df['01']:
    print(i)

df['np'][0]
df.loc[df.isnan()]
df.isnull().any()
type(df.loc[1,'desenho']) == str
np.isnan(df.loc[1,'desenho']) == True
type(df.loc[0,'desenho']) == np.float64

len(df.loc[0,'desenho'][df.loc[0,'desenho'].find('-')+1:])

type(df['desenho'][1])
type(df['np'][1])
df['np'][1]+'.'+ df['programa'][1]

py.position()
py.KEYBOARD_KEYS
py.LOG_SCREENSHOTS
py.getWindowsWithTitle('Google')[0].maximize()
py.getAllTitles()
py.getAllWindows()
py.getActiveWindow().minimize()
py.getActiveWindow().maximize()

import subprocess
def process_exists(process_name):
    progs = str(subprocess.check_output('tasklist'))
    if process_name in progs:
        return True
    else:
        return False
    
process_exists('Edge.exe')
process_exists('icad.exe')

start = datetime.datetime.now()
end = datetime.datetime.now()
diference_time = end - start
print(diference_time)
    
index = df[df['np'].isnull()].index
df = df.drop(index)
py.write(df['np'][1])



layout = [
            [sg.Text('Part Number', size=(20,0)),sg.Input(size=(15,0),key='partnumber')],
            [sg.Text('_'  * 100, size=(65, 1))],
            [sg.Text('Qual Tipo de Ordem?', justification='left',
            size=(20,0))], 
            [sg.Radio('Serviço','os',default=True, size=(15,0)),
            sg.Radio('Estoque','os', size=(15,0)),
            sg.Radio('Protótipo','os', size=(15,0))],
            [sg.Text('Número da Ordem', size=(20,0)), sg.Input(size=(15,0),key='numero')],
            [sg.Text('_'  * 100, size=(65, 1))],
            [sg.Text('Item da O.S.', size=(20,0)), 
            sg.Combo(values= list(np.arange(1,100)),key='item')],
            [sg.Submit(),sg.Cancel()]
            ]
        
        #Janela
janela = sg.Window('Controle de desenhos').layout(layout)
janela.element_list()
        
        #Extrair Informações
button, values = janela.Read()


inicio = 1 
pn = ('sec451-00', 't225-00', 'sec753-00')
item = 3
new_array = np.array([])
while inicio <= item:
    for i in np.arange(inicio,item+1):
        new_array = np.append(new_array,pn[i-1])
    print(inicio)
    inicio+=1


np.arange(1,3)

sg.Combo?
sg.Input?
sg.Output?
'OI'.lower()
# 