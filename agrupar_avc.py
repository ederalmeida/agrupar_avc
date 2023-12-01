import pandas as pd
import os
import sys
import PySimpleGUI as sg
import csv

def exibir_janela_inicial():
    sg.theme('Reddit')

    cabecalho = [[sg.Text('Agrupador de AVCs', size=(20,1), justification='center', font=("Helvetica", 20))],
                 [sg.Text('_'  * 60, size=(45, 1))]]
    
    linha1 = [[sg.Text('')],
              [sg.Text('Selecione a pasta onde estão salvos os AVCs', size=(50, 1))],
              [sg.InputText('', key='-PASTA-', size=(40, 1)), sg.FolderBrowse('procurar')]]
    
    linha2 = [[sg.Text('')],
              [sg.Text('Selecione o arquivo com as informações das Concessões', size=(50, 1))],
              [sg.InputText('', key='-INF_CONC-', size=(40, 1)), sg.FileBrowse('procurar')]]
    
    linha3 = [[sg.Text('')],
              [sg.Text('Nome do arquivo agrupador (não precisa da extensão .xlsx)', size=(50, 1))],
              [sg.InputText('', key='-NOME_ARQUIVO-', size=(40, 1))]]
    
    linha4 = [[sg.Text('')],
              [sg.Button('EXECUTAR', key='-AGRUPAR-', enable_events=True)]]
    
    layout = [cabecalho,
              linha1,
              linha2,
              linha3,
              linha4]

    janela = sg.Window('Agrupador de AVCs', layout, default_element_size=(40, 1), element_justification='left', grab_anywhere=False) 

    while True:
        event, values = janela.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            janela.close()
            sys.exit()

        if event == '-CADASTRAR-':
            janela_cadastrar()

        if event == '-AGRUPAR-':
            if values['-PASTA-'] == '':
                sg.popup('Favor indicar a pasta onde estão os AVCs', title='Erro')
            elif values['-NOME_ARQUIVO-'] == '':
                sg.popup('Favor nome para o arquivo que será criado', title='Erro')
            else:
                janela.close()
                sg.popup('O processo de agrupamento irá iniciar agora. Ao final irá aparecer uma janela indicado a conclusão', title='Aviso')
                dados_concessoes = importar_cadastro_concessoes(values['-INF_CONC-'])
                dados_para_excel = armazenar_dados_avc(obter_relacao_xls(values['-PASTA-']), dados_concessoes)
                exportar_excel(dados_para_excel, os.path.join(values['-PASTA-'], values['-NOME_ARQUIVO-']))
                sg.popup('Agrupamento de AVCs concluído')

# construir janela para cadastrar concessões, material e centro de lucro

# Importar informações sobre materia e centros de lucro das concessões
def importar_cadastro_concessoes(caminho_arquivo):

    #path_atual = os.path.abspath(os.path.dirname(__file__))
    dados_concessao = {}

    with open(caminho_arquivo, encoding="utf8") as arquivo:
 
        tabela = csv.reader(arquivo, delimiter=';')

        for linha in tabela:
            dados_concessao[linha[0]] = [linha[1], linha[2]]

    arquivo.close()
    
    return dados_concessao
 

# Funçao para obter lista de arquivos xls no diretorio indicado pelo usuário
def obter_relacao_xls(caminho):
    # lista que irá receber todos os caminhos para os xmls na pasta que foi indicada
    enderecos_arquivos_xls = []

    # Lendo a pasta indicada
    listar_objetos_diretorio = os.listdir(caminho)
    
    # Percorrendo todos os objetos encontrados na pasta
    for objeto in listar_objetos_diretorio:
        
        # se o final do arquivo for .xml adiciona na lista
        if objeto[-4:].lower() == '.xls':
            enderecos_arquivos_xls.append([objeto[4:8], os.path.join(caminho, objeto)])

    # Se a lista não for vazia, retorna a lista
    if len(enderecos_arquivos_xls) != 0:
        return enderecos_arquivos_xls

    # Se for, informa
    #TODO melhorar para que não feche o robô. De preferência, voltar para a tela de interação que chamou a função
    else:
        sg.popup('Não existem arquivos .XLS no diretório indicado a serem importados', title='ERRO!')
        sys.exit()

def verfica_valor(valor):
    virgula = ','
    if virgula not in valor:
        tamanho_parte_inteira = len(valor) - 2
        valor_ajustado = valor[0:tamanho_parte_inteira] + ',' + valor[-2::1] 
    else:
        valor_ajustado = valor

    return valor_ajustado

def armazenar_dados_avc(relacao_xls, dados_concessoes):    
    dados_avcs = []

    for avc in relacao_xls:
        material_e_centro_lucro = dados_concessoes.get(avc[0])

        dados = pd.read_html(avc[1])
        
        for i in range(0, len(dados[0])):
            if str(dados[0][0][i]) == 'nan':
                continue
            elif str(dados[0][1][i])[0:7] == 'Empresa':
                continue
            elif dados[0][0][i] == 'Total Geral':
                continue
            else:
                primeiro_vencimento = verfica_valor(dados[0][9][i])
                segundo_vencimento = verfica_valor(dados[0][10][i])
                terceiro_vencimento = verfica_valor(dados[0][11][i])
                dados_avcs.append([dados[0][0][i], dados[0][1][i], dados[0][2][i], primeiro_vencimento,
                                    segundo_vencimento, terceiro_vencimento, material_e_centro_lucro[0], material_e_centro_lucro[1]])
            
    return dados_avcs

def exportar_excel(dados_para_excel, nome_arquivo):
    df = pd.DataFrame(dados_para_excel, columns=['ONS', 'USUARIAS', 'CNPJ', '1o. Vencimento', '2o. Vencimento', '3o. Vencimento', 'Material', 'Centro Lucro'])
    df.to_excel(nome_arquivo + '.xlsx', index=False)


while True:
    exibir_janela_inicial()
