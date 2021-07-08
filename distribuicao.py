from PySimpleGUI.PySimpleGUI import VerticalSeparator
import pandas as pd
import PySimpleGUI as sg
import json
import os
import pyautogui as py
import numpy as np
import time
import subprocess
import datetime

class Distribuicao:
    
    def __init__(self):
        #layout
        layout = [
            [sg.Text('Part Number', size=(15,0)),sg.Input(size=(15,0),key='partnumber')],
            [sg.Button('Verificar Part Number')]
            ]
        
        #Janela
        janela = sg.Window('Distribuição desenhos').layout(layout)
        
        #Extrair Informações
        self.button, self.values = janela.Read()
        self.comprimento = 3 # Colocar o len(df) após testes

    def VerificarPath(self):
        '''
            Verifica de o arquivo existe.
        '''
        if os.path.exists(self.values['partnumber']+'.xlsx'):
            print('Arquivo existe no formato XLSX.')
            data = self.values['partnumber'].lower()+'.xlsx'
            print(data)
            self.df = pd.read_excel(data,engine="openpyxl")
            index = self.df[self.df['np'].isnull()].index
            self.df = self.df.drop(index)
            print(self.df.head())
            tela.Alerta_Inicio()
            self.path = True
            print('Path = {}'.format(self.path))

        elif os.path.exists(self.values['partnumber']+'.xls'):
            print('Arquivo existe no formato XLS.')
            data = self.values['partnumber'].lower()+'.xls'
            print(data)
            self.df = pd.read_excel(data,engine="openpyxl")
            index = self.df[self.df['np'].isnull()].index
            self.df = self.df.drop(index)
            print(self.df.head())
            tela.Alerta_Inicio()
            self.path = True
            print('Path = {}'.format(self.path))
        
        else:
            print('Arquivo não Existe.')
             #layout
            layout = [
                [sg.Text('Arquivo não existe.')],
                [sg.Button('Criar Lista com Desenhos')]
                ]
        
            #Janela
            janela = sg.Window('Listagem não existente').layout(layout)
        
            #Extrair Informações
            self.button, self.values = janela.Read()
            self.path = False

    def Alerta_Inicio(self):
        self.alerta = py.alert('Feche todos os desenhos, tanto no Autocad quanto no SolidEdge.\nO código irá rodar. Não mexa no mouse ou teclado.')

    def Iniciar(self):
        partnumber = self.values['partnumber']
        print(f'Part Number: {partnumber}')

    def CheckAba(self, index):
        self.index_aba = index
        if self.df['aba'][self.index_aba] == 's':
            print('Desenho tem Aba')
            self.aba = True
        else:
            print('Desenho não tem Aba')
            self.aba = False

    def CheckFormato(self, index):
        self.index_formato = index
        if self.df['programa'][self.index_formato] == 'dft':
            print('Desenho em Edge')
            self.formato = 'edge'
        else:
            print('Desenho em CAD')
            self.formato = 'cad'

    def CheckConjunto(self, index):
        self.index_conjunto = index
        if self.df['conjunto'][self.index_conjunto] == 's':
            print('É Desenho de Conjunto')
            self.conjunto = True
        else:
            print('Não é Desenho de Conjunto')
            self.conjunto = False  

    def CheckEdge(self):
        edge = 'Edge.exe'
        progs = str(subprocess.check_output('tasklist'))
        if edge in progs:
            print('Edge Aberto')
            self.edge = True
        else:
            print('Edge Fechado')
            self.edge = False   

    def CheckImpressora(self,index):
        self.index_impressora = index
        if self.df['impressora'][self.index_impressora] == 400:
            print('Imprimir na HP 400')
            self.impressora = 400
        elif self.df['impressora'][self.index_impressora] == 500:
            print('Imprimir na HP 500')
            self.impressora = 500

    def CheckCad(self):
        cad = 'icad.exe'
        auto_cad = 'acadlt.exe'
        progs = str(subprocess.check_output('tasklist'))
        if cad in progs:
            print('Cad Aberto')
            self.cad = True
        elif auto_cad in progs:
            print('AutoCad Aberto')
            self.cad = True
        else:
            print('Cad Fechado')
            self.cad = False        

    def Imprimir(self):
        py.PAUSE = 2.0
        if self.conjunto == True:
            if self.formato == 'edge':
                if self.impressora == 400:
                    py.hotkey('ctrl','p')
                    py.moveTo(941,371)
                    py.click()
                    py.moveTo(882,457)
                    py.click()
                    py.moveTo(992,521)
                    py.doubleClick()
                    py.write('3')
                    if self.aba == True:
                        py.moveTo(769,630)
                        py.click()
                        
                    elif self.aba == False:
                        py.moveTo(763,649)
                        py.click()

                    # Alterar os comandos abaixo
                    py.moveTo(1169,367)
                    # Imprimir
                    py.click()    
                    time.sleep(5)
                    py.moveTo(1882,148)
                    py.click()
                    time.sleep(1) 
                    py.write('n')
                    time.sleep(3)
                    py.moveTo(1900,924)
                    py.click()

                elif self.impressora == 500:
                    py.hotkey('ctrl','p')
                    py.moveTo(941,371)
                    py.click()
                    py.moveTo(867,498)
                    py.click()
                    py.moveTo(992,521)
                    py.doubleClick()
                    py.write('3')
                    if self.aba == True:
                        py.moveTo(769,630)
                        py.click()
                       
                    elif self.aba == False:
                        py.moveTo(763,649)
                        py.click()

                    # Alterar os comandos abaixo
                    py.moveTo(1169,367)
                    # Imprimir
                    py.click()    
                    time.sleep(20)    
                    py.moveTo(1882,148)
                    py.click()
                    time.sleep(5) 
                    py.write('n')
                    time.sleep(3)
                    py.moveTo(1900,924)
                    py.click()  

            elif self.formato == 'cad':
                if self.impressora == 400:
                    py.hotkey('ctrl','p')
                    py.moveTo(832,361)
                    py.click()
                    py.moveTo(769,471)
                    py.click()
                    py.moveTo(993,574)
                    py.doubleClick()
                    py.write('3')
                    py.moveTo(675,652)
                    py.click()
                    py.moveTo(656,709)
                    py.click()
                    # Alterar os comandos abaixo
                    py.moveTo(1102,794)
                    py.click()
                    time.sleep(5)
                    py.moveTo(1900,924)
                    py.click()    
                    time.sleep(10)
                    py.write('_wclose', interval=0.5)
                    time.sleep(3)
                    py.press('enter')
                    time.sleep(3)
                    py.write('n')
                    

                elif self.impressora == 500:
                    py.hotkey('ctrl','p')
                    py.moveTo(832,361)
                    py.click()
                    py.moveTo(792,495)
                    py.click()
                    py.press('enter')
                    py.moveTo(993,574)
                    py.doubleClick()
                    py.write('3')
                    py.moveTo(675,652)
                    py.click()
                    py.moveTo(656,709)
                    py.click()
                    # Alterar os comandos abaixo
                    py.moveTo(1102,794)
                    py.click()
                    time.sleep(5)
                    py.moveTo(1900,924)
                    py.click()      
                    time.sleep(10)
                    py.write('_wclose', interval=0.5)
                    time.sleep(3)
                    py.press('enter')
                    time.sleep(3)
                    py.write('n')
                    time.sleep(3)
                    py.moveTo(1900,924)
                    py.click()
                
                
        elif self.conjunto == False:
            if self.formato == 'edge':
                if self.impressora == 400:
                    py.hotkey('ctrl','p')
                    py.moveTo(941,371)
                    py.click()
                    py.moveTo(882,457)
                    py.click()
                    py.moveTo(992,521)
                    py.doubleClick()
                    py.write('1')
                    if self.aba == True:
                        py.moveTo(769,630)
                        py.click()
                       
                    elif self.aba == False:
                        py.moveTo(763,649)
                        py.click()

                    # Alterar os comandos abaixo
                    py.moveTo(1169,367)
                    # Imprimir
                    py.click()    
                    time.sleep(5)
                    py.moveTo(1882,148)
                    py.click() 
                    time.sleep(1) 
                    py.write('n')
                    time.sleep(3)
                    py.moveTo(1900,924)
                    py.click()

                elif self.impressora == 500:
                    py.hotkey('ctrl','p')
                    py.moveTo(941,371)
                    py.click()
                    py.moveTo(867,498)
                    py.click()
                    py.moveTo(992,521)
                    py.doubleClick()
                    py.write('1')
                    if self.aba == True:
                        py.moveTo(770,662)
                        py.click()
                        
                    elif self.aba == False:
                        py.moveTo(763,649)
                        py.click()

                    # Alterar os comandos abaixo
                    py.moveTo(1169,367)
                    # Imprimir
                    py.click()    
                    time.sleep(20)
                    py.moveTo(1882,148)
                    py.click() 
                    time.sleep(5) 
                    py.write('n')
                    time.sleep(3)
                    py.moveTo(1900,924)
                    py.click()      
            
            elif self.formato == 'cad':
                if self.impressora == 400:
                    py.hotkey('ctrl','p')
                    py.moveTo(832,361)
                    py.click()
                    py.moveTo(769,471)
                    py.click()
                    py.moveTo(993,574)
                    py.doubleClick()
                    py.write('1')
                    py.moveTo(675,652)
                    py.click()
                    py.moveTo(656,709)
                    py.click()
                    # Alterar os comandos abaixo
                    py.moveTo(1102,794)
                    # Imprimir
                    py.click()
                    time.sleep(5)
                    py.moveTo(1900,924)
                    py.click()      
                    time.sleep(10)
                    py.write('_wclose', interval=0.5)
                    time.sleep(3)
                    py.press('enter')
                    time.sleep(3)
                    py.write('n')
                    time.sleep(3)
                    py.moveTo(1900,924)
                    py.click()
    
                elif self.impressora == 500:
                    py.hotkey('ctrl','p')
                    py.moveTo(832,361)
                    py.click()
                    py.moveTo(792,495)
                    py.click()
                    py.press('enter')
                    py.moveTo(993,574)
                    py.doubleClick()
                    py.write('1')  
                    py.moveTo(675,652)
                    py.click()
                    py.moveTo(656,709)
                    py.click()
                    # Alterar os comandos abaixo
                    py.moveTo(1102,794)
                    # Imprimir
                    py.click()  
                    time.sleep(5)
                    py.moveTo(1900,924)
                    py.click()    
                    time.sleep(10)
                    py.write('_wclose', interval=0.5)
                    time.sleep(3)
                    py.press('enter')
                    time.sleep(3)
                    py.write('n')
                    time.sleep(3)
                    py.moveTo(1900,924)
                    py.click()
                                 
    def ProcuraDesenho(self):
        if self.path == True:
            for i in np.arange(0,len(self.df)):
                py.press('winleft')
                time.sleep(5)
                py.write('google desktop')
                time.sleep(5)
                py.press('enter')
                time.sleep(5)
                if type(self.df.loc[i,'desenho']) == float or type(self.df.loc[i,'desenho']) == np.float64:
                    print('Valor é NA') 
                    print(self.df['np'][i]+'.'+ self.df['programa'][i])
                    desenho = self.df['np'][i]
                    py.write(desenho+'.'+ self.df['programa'][i],interval=0.5)
                    time.sleep(10)
                    py.moveTo(837,620)
                    time.sleep(7)
                    # Verifica se o programa Cad está aberto
                    tela.CheckCad()
                    # Verifica se o programa Solid Edge está aberto
                    tela.CheckEdge()
                    if self.edge == False or self.cad == False:
                        py.press('enter')
                        time.sleep(25)
                        py.getActiveWindow().maximize()
                        py.moveTo(1000,1000)
                        time.sleep(5)
                        # Verifica se o desenho é de conjunto
                        tela.CheckConjunto(i)
                        # Verifica se o desenho é de conjunto
                        tela.CheckAba(i)
                        # Verifica se o desenho irá ser impresso na plotter ou não
                        tela.CheckImpressora(i)
                        # Verifica se o programa é cad ou edge
                        tela.CheckFormato(i)  
                        # Processo de Impressão
                        tela.Imprimir()  

                    elif self.edge == True or self.cad == True:
                        py.press('enter')
                        time.sleep(10)
                        py.getActiveWindow().maximize()
                        py.moveTo(1000,1000)
                        time.sleep(7)
                        tela.CheckConjunto(i)
                        # Verifica se o desenho é de conjunto
                        tela.CheckAba(i)
                        # Verifica se o desenho irá ser impresso na plotter ou não
                        tela.CheckImpressora(i)
                        # Verifica se o programa é cad ou edge
                        tela.CheckFormato(i)  
                        # Processo de Impressão
                        tela.Imprimir() 
                        
                elif type(self.df.loc[i,'desenho']) == str:
                    print('Valor não é NA')
                    desenho = self.df.loc[i,'desenho'][self.df.loc[i,'desenho'].find('-')+1:]
                    py.write(desenho+'.'+ self.df['programa'][i], interval=0.5)
                    time.sleep(10)
                    py.moveTo(837,620)
                    time.sleep(7)
                    # Verifica se o programa Cad está aberto
                    tela.CheckCad()
                    # Verifica se o programa Solid Edge está aberto
                    tela.CheckEdge()
                    if self.edge == False or self.cad == False:
                        py.press('enter')
                        time.sleep(20)
                        py.getActiveWindow().maximize()
                        py.moveTo(1000,1000)
                        time.sleep(5)
                         # Verifica se o desenho é de conjunto
                        tela.CheckConjunto(i)
                        # Verifica se o desenho é de conjunto
                        tela.CheckAba(i)
                        # Verifica se o desenho irá ser impresso na plotter ou não
                        tela.CheckImpressora(i)
                        # Verifica se o programa é cad ou edge
                        tela.CheckFormato(i)  
                        # Processo de Impressão
                        tela.Imprimir() 

                    elif self.edge == True or self.cad == True:
                        py.press('enter')
                        time.sleep(10)
                        py.getActiveWindow().maximize()
                        py.moveTo(1000,1000)
                        time.sleep(7)
                         # Verifica se o desenho é de conjunto
                        tela.CheckConjunto(i)
                        # Verifica se o desenho é de conjunto
                        tela.CheckAba(i)
                        # Verifica se o desenho irá ser impresso na plotter ou não
                        tela.CheckImpressora(i)
                        # Verifica se o programa é cad ou edge
                        tela.CheckFormato(i)  
                        # Processo de Impressão
                        tela.Imprimir() 

    def TimeStart (self):
        self.start = datetime.datetime.now()

    def TimeEnd (self):
        self.end = datetime.datetime.now()

    def FimPrograma (self):
        self.fim = self.end - self.start
        #layout
        layout = [
                [sg.Text('FIM DO PROCESSO')],
                [sg.Text('TEMPO DE EXECUÇÃO: {}'.format(self.fim))],
                [sg.Button('FECHAR BOT')]
                ]
        
        #Janela
        janela = sg.Window('FIM PROCESSO').layout(layout)
        
        #Extrair Informações
        self.button, self.values = janela.Read()    

# Instanciando a Classe
tela = Distribuicao()

# Tempo Início
tela.TimeStart()

# Iniciar o Programa
tela.Iniciar()

# Verificar se o arquivo existe
tela.VerificarPath()

# Colocando para rodar o programa
tela.ProcuraDesenho()

# Tempo Fim
tela.TimeEnd()

# Aviso Fim do Programa
tela.FimPrograma()

