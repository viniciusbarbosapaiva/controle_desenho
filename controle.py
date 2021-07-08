from platform import win32_ver
from PySimpleGUI.PySimpleGUI import Submit, VerticalSeparator
import pandas as pd
import PySimpleGUI as sg
import json
import os
import pyautogui as py
import numpy as np
import time
import subprocess
import datetime
import traceback

class Controle:
    py.PAUSE = 3
    def __init__(self,data):
        self.usuario = data['nome']
        self.senha = data['senha']
        self.surco = data['site']
        self.path = True
        print('Nome do Usuario: {}'.format(self.usuario))
        print('Senha: {}'.format(self.usuario))

    def TelaInicial (self): 
        self.process = True   
        layout = [
            [sg.Text('_'  * 100, size=(65, 1))],
            [sg.Text('Qual Tipo de Ordem?', justification='left',
            size=(20,0))], 
            [sg.Radio('Serviço','os',default=True, size=(15,0)),
            sg.Radio('Estoque','os', size=(15,0)),
            sg.Radio('Protótipo','os', size=(15,0))],
            [sg.Text('Número da Ordem', size=(20,0)), sg.Input(size=(15,0),key='numero')],
            [sg.Text('_'  * 100, size=(65, 1))],
            [sg.Text('Quantos Itens Têm a Ordem?', size=(20,0)), 
            sg.Combo(values= list(np.arange(1,100)),default_value=1,key='item')],
            [sg.Button('Submit'),sg.Button('Cancel')],
            ]
        

        #Janela
        janela = sg.Window('Controle de desenhos').layout(layout)
        self.button, self.values = janela.read()
        print(self.button,self.values)
    
        if self.button == 'Cancel':
            self.process == False
            janela.close()
             
        elif self.button == 'Submit':
            self.process == True
            janela.close()
            bot.Confirmacao()
       
    def Confirmacao(self):
        if self.process == True:
            layout = [
                [sg.Text('Os Valores Abaixo Estão Corretos?', size=(20,0))],
                [sg.Text(f'Ordem de Serviço: {self.values[0]}', size=(20,0))],
                [sg.Text(f'Ordem de Estoque: {self.values[1]}', size=(20,0))],
                [sg.Text(f'Ordem de Protótipo: {self.values[2]}', size=(20,0))],
                [sg.Text('Número da Ordem: {}'.format(self.values['numero']), size=(20,0))],
                [sg.Text('Quantos Itens: {}'.format(self.values['item']), size=(20,0))],
                [sg.Button('Sim'),sg.Button('Não')],]
            janela = sg.Window('Confirmação').layout(layout)
            button,values = janela.read()
            print(button,values)
            
            if button == 'Sim':
                janela.close()
                bot.PartNumberList()      
        
            if button == 'Não':
                janela.close()
                bot.TelaInicial()
                    
    def Iniciar(self):
        partnumber = self.values['partnumber']
        print(f'Part Number: {partnumber}')

    def PartNumberList(self):
        self.lista_np = np.array([])
        for i in np.arange(1,self.values['item']+1):
            layout = [
                [sg.Text('Part Number (Item: {})'.format(i), size=(20,0)),
                sg.Input(size=(20,0),key='partnumber')],
                [sg.Button('Submit'), sg.Button('Cancel')]
                ]

            janela = sg.Window('Lista de Part Number').layout(layout)
            button,values = janela.read()
            print(button,values)
            if button == (None,'Cancel'):
                self.path = False
                self.process = False
                break
            else: 
                self.lista_np = np.append(self.lista_np,values['partnumber'].lower())
            #janela.close()
        print(self.lista_np)

    def VerificarPath(self):
        '''
            Verifica de o arquivo existe.
        '''
        self.dict_df = {}
        for i in np.arange(1,self.values['item']+1):
            if os.path.exists(self.lista_np[i-1]+'.xlsx') and self.process == True:
                print('Arquivo existe no formato XLSX.')
                data = self.lista_np[i-1]+'.xlsx'
                print(data)
                self.df = pd.read_excel(data,engine="openpyxl")
                index = self.df[self.df['np'].isnull()].index
                self.df = self.df.drop(index)
                print(self.df.head())

                dict_df = {self.lista_np[i-1]:self.df}
                self.dict_df.update(dict_df)

                self.path = True
                print('Path = {}'.format(self.path))

            elif os.path.exists(self.lista_np[i-1]+'.xls') and self.process == True:
                print('Arquivo existe no formato XLS.')
                data = self.lista_np[i-1]+'.xls'
                print(data)
                self.df = pd.read_excel(data,engine="openpyxl")
                index = self.df[self.df['np'].isnull()].index
                self.df = self.df.drop(index)
                print(self.df.head())

                dict_df = {self.lista_np[i-1]:self.df}
                self.dict_df.update(dict_df)

                self.path = True
                print('Path = {}'.format(self.path))
        
            else:
                print('Arquivo não Existe.')
             #  layout
                layout = [
                    [sg.Text('Arquivo não existe.')],
                    [sg.Button('Criar Lista com Desenhos')]
                    ]
        
                #Janela
                janela = sg.Window('Listagem não existente').layout(layout)
        
                #Extrair Informações
                self.button, self.values = janela.Read()
                self.path = False
        
        print(self.dict_df)

    def OpenGoogle(self):

        if self.path == True:
            bot.Alerta_Inicio()
            py.press('winleft')
            py.write('google chrome')
            py.press('enter')
            py.getActiveWindow().maximize()
            try:
                py.write(self.surco)
                py.press('enter')
                print('Site encontrado.')
                self.process == True
                print('Process = {}'.format(self.process))
                bot.Login()
                bot.IrParaDistribuicao()
                bot.Fabricacao()
                bot.FabricacaoDistribuica()
                bot.Inspecao()
                bot.InspecaoMontagemDistribuica()
                bot.Montagem()
                bot.InspecaoMontagemDistribuica()
                bot.CloseWindown()
            except IOError:
                print('Não é possível conectar ao site expecificado.')
                self.process = False
                print('Process = {}'.format(self.process))
                pass
        
    def Login(self):
        if self.process == True:
            py.moveTo(1710,720)
            py.doubleClick()
            py.press('backspace')
            py.write(self.usuario)

            py.press('tab')
            py.press('backspace')
            py.write(self.senha)

            py.press('tab')
            py.press('enter')

    def IrParaDistribuicao(self):
        py.moveTo(798,140)
        py.moveTo(788,221)
        py.click()

    def Fabricacao(self):
        py.moveTo(562,1025) # Mover para "Inclusão de Distribuição"
        py.click() # Clica para "Inclusão de Distribuição"
        py.press('enter') # Confirma data
        py.moveTo(812,276) # Mover para campo "Finalidade"
        py.click() # Clica no campo 'Finalidade'
        py.moveTo(806,345) # Mover para "Fabricação"
        py.click() # Clica em "Fabricação"
        py.press('tab') # Passa para o próximo campo
        py.write(self.usuario) # Escreve o "Entregue por"
        py.press('tab') # Passa para o próximo campo
        py.press('tab') # Passa para o próximo campo
        py.moveTo(844,364) # Mover para o campo "OS"
        py.click() # Clica no campo "OS"
        if self.values[0] == True: # Seleciona "Serviço"
            py.moveTo(789,418)
        elif self.values[1] == True: # Seleciona "Inspeção"
            py.moveTo(819,432)
        elif self.values[2] == True: # Seleciona "Protótipo"
            py.moveTo(800,443)
        py.click() # Seleciona uma das opções acima
        py.press('tab') # Passa para o próximo campo
        py.write(self.values['numero']) # Escreve o item 
        py.press('tab') # Passa para o próximo campo
        py.press('tab') # Passa para o próximo campo
        py.press('enter') # Confirma inclusão da distribuição
        time.sleep(10)

    def FabricacaoDistribuica(self):
        py.moveTo(904,994)
        py.click()
        for i in np.arange(1,self.values['item']+1):
            for x in np.arange(0,len(self.dict_df[self.lista_np[i-1]])):
                if type(self.dict_df[self.lista_np[i-1]].loc[x,'desenho']) == float or type(self.dict_df[self.lista_np[i-1]].loc[x,'desenho']) == np.float64:
                    print('Valor é NA')
                    desenho = self.dict_df[self.lista_np[i-1]]['np'][x]

                elif type(self.dict_df[self.lista_np[i-1]].loc[x,'desenho']) == str:
                    print('Valor Não é NA')
                    desenho = self.dict_df[self.lista_np[i-1]]['desenho'][x] 
                if x == 0 and i == 1:
                    py.moveTo(860,1025)
                    py.click()
                else: 
                    py.moveTo(611,1022)
                    py.click()
                py.write(desenho, interval=0.5)
                py.press('enter')
                py.write('{}'.format(i))
                py.press('enter')
                py.press('enter')
                py.press('enter')
                py.press('enter')
                py.press('enter')
        py.press('esc') # Voltar tela de distribuição
        py.moveTo(1029,996) # Mover para "Impressão"
        py.click() # Clica na "Impressão"
        py.press('enter') # Confirma
        #py.moveTo(964,583) # Mover para o PDF
        #py.click() # Clica no PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('enter') # Clica para download
        time.sleep(10)
        py.moveTo(124,1010) # Mover para o arquivo baixado
        py.click() # Clica no arquivo baixado
        py.moveTo(1842,128) # Mover para o ícone de imprimir
        py.click() # Clica no ícone de imprimrir
        py.moveTo(1473,150) # Mover para a seleção de impressora
        py.click() # Clica na seleção de impressora
        py.moveTo(1469,180) # Mover para a impressora desejada
        py.click() # Clica na impressora desejada
        py.moveTo(1421,264) # Mover para o campo de "Cópias"
        py.doubleClick() # Seleciona o campo de "Cópias"
        py.write('2') # Altera a quantidade de cópias para dois
        py.moveTo(1470,893) # Mover para imprimir
        py.click() # Imprimir
        py.moveTo(471,18) # Mover para fechar aba
        py.click() # Fechar aba
        py.moveTo(959,1021) # Mover para voltar para a tela de "Distribuição de desenhos"
        py.click() # Clica para voltar para a tela de "Distribuição de desenhos"

    def Inspecao(self):
        py.moveTo(562,1025) # Mover para "Inclusão de Distribuição"
        py.click() # Clica para "Inclusão de Distribuição"
        py.press('enter') # Confirma data
        py.moveTo(812,276) # Mover para campo "Finalidade"
        py.click() # Clica no campo 'Finalidade'
        py.moveTo(799,363) # Mover para "Inspeção"
        py.click() # Clica em "Inspeção"
        py.press('tab') # Passa para o próximo campo
        py.write(self.usuario) # Escreve o "Entregue por"
        py.press('tab') # Passa para o próximo campo "Solicitante"
        py.press('tab') # Passa para o próximo campo "Observação"
        py.press('tab') # Passa para o botão "Gravar"
        py.press('enter') # Confirma inclusão da distribuição
        time.sleep(10)

    def Montagem(self):
        py.moveTo(562,1025) # Mover para "Inclusão de Distribuição"
        py.click() # Clica para "Inclusão de Distribuição"
        py.press('enter') # Confirma data
        py.moveTo(812,276) # Mover para campo "Finalidade"
        py.click() # Clica no campo 'Finalidade'
        py.moveTo(819,375) # Mover para "Montagem"
        py.click() # Clica em "Montagem"
        py.press('tab') # Passa para o próximo campo
        py.write(self.usuario) # Escreve o "Entregue por"
        py.press('tab') # Passa para o próximo campo
        py.press('tab') # Passa para o próximo campo
        py.moveTo(844,364) # Mover para o campo "OS"
        py.click() # Clica no campo "OS"
        if self.values[0] == True: # Seleciona "Serviço"
            py.moveTo(789,418)
        elif self.values[1] == True: # Seleciona "Inspeção"
            py.moveTo(819,432)
        elif self.values[2] == True: # Seleciona "Protótipo"
            py.moveTo(800,443)
        py.click() # Seleciona uma das opções acima
        py.press('tab') # Passa para o próximo campo
        py.write(self.values['numero']) # Escreve o item 
        py.press('tab') # Passa para o próximo campo
        py.press('tab') # Passa para o próximo campo
        py.press('enter') # Confirma inclusão da distribuição
        time.sleep(10)

    def InspecaoMontagemDistribuica(self):
        py.moveTo(904,994)
        py.click()
        for i in np.arange(1,self.values['item']+1):
            for x in np.arange(0,len(self.dict_df[self.lista_np[i-1]])):
                if self.dict_df[self.lista_np[i-1]].loc[x,'conjunto'] == 's':
                    if type(self.dict_df[self.lista_np[i-1]].loc[x,'desenho']) == float or type(self.dict_df[self.lista_np[i-1]].loc[x,'desenho']) == np.float64:
                        print('Valor é NA')
                        desenho = self.dict_df[self.lista_np[i-1]]['np'][x]

                    elif type(self.dict_df[self.lista_np[i-1]].loc[x,'desenho']) == str :
                        print('Valor Não é NA')
                        desenho = self.dict_df[self.lista_np[i-1]]['desenho'][x] 
                    if x == 0 and i == 1:
                        py.moveTo(860,1025)
                        py.click()
                    else: 
                        py.moveTo(611,1022)
                        py.click()
                    py.write(desenho, interval=0.5)
                    py.press('enter')
                    py.write('{}'.format(i))
                    py.press('enter')
                    py.press('enter')
                    py.press('enter')
                    py.press('enter')
                    py.press('enter')
        py.press('esc') # Voltar tela de distribuição
        py.moveTo(1029,996) # Mover para "Impressão"
        py.click() # Clica na "Impressão"
        py.press('enter') # Confirma
        #py.moveTo(964,583) # Mover para o PDF
        #py.click() # Clica no PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('tab') # Mover para o PDF
        py.press('enter') # Clica para download
        time.sleep(10)
        py.moveTo(124,1010) # Mover para o arquivo baixado
        py.click() # Clica no arquivo baixado
        py.moveTo(1842,128) # Mover para o ícone de imprimir
        py.click() # Clica no ícone de imprimrir
        py.moveTo(1473,150) # Mover para a seleção de impressora
        py.click() # Clica na seleção de impressora
        py.moveTo(1469,180) # Mover para a impressora desejada
        py.click() # Clica na impressora desejada
        py.moveTo(1421,264) # Mover para o campo de "Cópias"
        py.doubleClick() # Seleciona o campo de "Cópias"
        py.write('1') # Altera a quantidade de cópias para dois
        py.moveTo(1470,893) # Mover para imprimir
        py.click() # Imprimir
        py.moveTo(471,18) # Mover para fechar aba
        py.click() # Fechar aba
        py.moveTo(959,1021) # Mover para voltar para a tela de "Distribuição de desenhos"
        py.click() # Clica para voltar para a tela de "Distribuição de desenhos"

    def Alerta_Inicio(self):
        self.alerta = py.alert('O BOT IRÁ INICIAR.\nNão mexa no mouse ou teclado.')

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

    def CloseWindown (self):
        py.hotkey('alt','f4')

if __name__ == '__main__':
    with open('config.json') as config_file:
        data = json.load(config_file)
    
    bot = Controle(data)
        
    # Tempo Início
    bot.TimeStart()

    # Tela Inicial
    bot.TelaInicial()

    # Verificar se o arquivo existe
    bot.VerificarPath()   

    # Abrindo o Google
    bot.OpenGoogle()

    # Tempo Fim
    bot.TimeEnd()

    # Aviso Fim do Programa
    bot.FimPrograma()



  

