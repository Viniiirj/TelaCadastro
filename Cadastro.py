import PySimpleGUI as sg
import pandas as pd
import openpyxl

class CadastroFerramenta:
    def __init__(self):
        sg.theme('DarkTeal9')
        self.dados()
        self.layout()
        self.event()

    def dados(self):
        EXCEL_FILE = 'Dados.xlsx'
        self.EXCEL_FILE = EXCEL_FILE
        self.df = pd.read_excel(self.EXCEL_FILE)
    def layout(self):
        self.layout = [
    [sg.Text('Central de Ferramentaria')],
    [sg.Text('Cód. Ferramenta:', size=(13,1)), sg.InputText(key='Cod. Ferramenta', size=(5,1)),
     sg.Text('Part Number:', size=(10,1)), sg.InputText(key='Part Number', size=(5,1))],
    [sg.Text('Fabricante:', size=(13,1)), sg.InputText(key='Fabricante', size=(15,1))],
    [sg.Text('Voltagem de uso:', size=(13,1)), sg.Combo(['110v', '220v'], key='Voltagem')],
    [sg.Text('Tamanho:', size=(13,1)), sg.InputText(key='Tamanho', size=(5,1)), sg.Text('Unidade de Medida:', size=(14,1)), sg.Combo(['Polegadas','Centimetros'],key='Unidade de Medida', size=(10,1))],
    [sg.Text('Tipo de Ferramenta:', size=(15,1)), sg.InputText(key='Tipo de Ferramenta', size=(15,1))],
    [sg.Text('Material da Ferramenta:', size=(17,1)), sg.InputText(key='Material da Ferramenta', size=(15,1))],
    [sg.Text('Descrição da Ferramenta:', size=(19,1)), sg.InputText(key='Descrição da Ferramenta', size=(30,1))],


    [sg.Submit('Cadastrar'), sg.Button('Limpar'), sg.Exit('Sair')]
]
        self.janela = sg.Window('Cadastro de ferramentas', self.layout)


    def clear_input(self):

        for key in self.values:
            self.janela[key]('')
        return None
    def event(self):


        while True:
            self.event, self.values = self.janela.read()
            if self.event == sg.WIN_CLOSED or self.event == 'Sair':
                break
            if self.event == 'Limpar':
                self.clear_input()
            if self.event == 'Cadastrar':
                self.df = self.df.append(self.values, ignore_index=True)
                self.df.to_excel(self.EXCEL_FILE, index=False)
                sg.popup('Ferramenta cadastrada!')
                self.clear_input()
        self.janela.close()

CadastroFerramenta()