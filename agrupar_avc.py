import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import sys
import csv

def procurar_pasta():
    pasta = filedialog.askdirectory()
    if pasta:
        entry_pasta.delete(0, tk.END)
        entry_pasta.insert(0, pasta)

def procurar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if arquivo:
        entry_arquivo.delete(0, tk.END)
        entry_arquivo.insert(0, arquivo)

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
        if objeto[-4:].lower() in('.xls', 'xlsx'):
            enderecos_arquivos_xls.append([objeto[4:8], os.path.join(caminho, objeto)])

    # Se a lista não for vazia, retorna a lista
    if len(enderecos_arquivos_xls) != 0:
        return enderecos_arquivos_xls

    # Se for, informa
    #TODO melhorar para que não feche o robô. De preferência, voltar para a tela de interação que chamou a função
    else:
        messagebox.showwarning("Erro!", 'Não existem arquivos .XLS no diretório indicado a serem importados')
        sys.exit()

def verfica_valor(valor):
    virgula = ','
    if virgula not in valor:
        tamanho_parte_inteira = len(valor) - 2
        valor_ajustado = valor[0:tamanho_parte_inteira] + ',' + valor[-2::1] 
    else:
        valor_ajustado = valor

    return valor_ajustado

def armazenar_dados_avc(relacao_xls, dados_concessoes, atribuicao):    
    dados_avcs = {}
    dados_clientes = []

    for avc in relacao_xls:
        material_e_centro_lucro = dados_concessoes.get(avc[0])

        dados = pd.read_html(avc[1])
        
        for i in range(0, len(dados[0])):
            if str(dados[0][1][i]) == 'Usuárias':
                dt_vencimento_1 = str(dados[0][9][i])[-10:]
                dt_vencimento_2 = str(dados[0][10][i])[-10:]
                dt_vencimento_3 = str(dados[0][11][i])[-10:]
            elif str(dados[0][0][i]) == 'nan':
                continue
            elif str(dados[0][1][i])[0:7] == 'Empresa':
                continue
            elif dados[0][0][i] == 'Total Geral':
                continue
            else:
                primeiro_vencimento = verfica_valor(dados[0][9][i])
                segundo_vencimento = verfica_valor(dados[0][10][i])
                terceiro_vencimento = verfica_valor(dados[0][11][i])
                dados_clientes.append([dados[0][0][i], dados[0][1][i], dados[0][2][i], primeiro_vencimento,
                                    segundo_vencimento, terceiro_vencimento, material_e_centro_lucro[0],
                                    material_e_centro_lucro[1], atribuicao])
                
    dados_avcs = {
        'dt_vencimento_1': dt_vencimento_1,
        'dt_vencimento_2': dt_vencimento_2,
        'dt_vencimento_3': dt_vencimento_3,
        'dados_clientes': dados_clientes
        }
            
    return dados_avcs

def exportar_excel(dados_para_excel, nome_arquivo):
    df = pd.DataFrame(dados_para_excel.get('dados_clientes'), columns=['ONS', 'Usuarias', 'CNPJ',
                                                                       dados_para_excel.get('dt_vencimento_1'),
                                                                       dados_para_excel.get('dt_vencimento_2'),
                                                                       dados_para_excel.get('dt_vencimento_3'),
                                                                       'Material','Centro Lucro', 'Atribuicao'])
    df = df.sort_values(by=['ONS'])
    df.to_excel(nome_arquivo + '.xlsx', index=False)

def executar():
    pasta = entry_pasta.get()
    arquivo = entry_arquivo.get()
    atribuicao = entry_atribuicao.get()
    nome_arquivo = entry_nome_arquivo.get()
    
    if not (pasta and arquivo and nome_arquivo):
        messagebox.showwarning("Campos incompletos", "Por favor, preencha todos os campos.")
    else:
        # Aqui você pode adicionar o código para a ação do botão "Executar"
        messagebox.showinfo("Informação", "O processo de agrupamento irá iniciar agora. Ao final irá aparecer uma janela indicado a conclusão")
        dados_concessoes = importar_cadastro_concessoes(arquivo)
        dados_para_excel = armazenar_dados_avc(obter_relacao_xls(pasta), dados_concessoes, atribuicao)
        exportar_excel(dados_para_excel, os.path.join(pasta, nome_arquivo))
        messagebox.showinfo("Informação", "Processamento concluído com sucesso!")

# Criação da janela principal
root = tk.Tk()
root.title("Agrupar AVCs")

# Label e campo para selecionar a pasta
label_pasta = tk.Label(root, text="Selecione a pasta onde estão salvos os AVCs")
label_pasta.grid(row=0, column=0, padx=10, pady=5, sticky='w')
entry_pasta = tk.Entry(root, width=50)
entry_pasta.grid(row=0, column=1, padx=10, pady=5)
button_pasta = tk.Button(root, text="Procurar Pasta    ", command=procurar_pasta)
button_pasta.grid(row=0, column=2, padx=10, pady=5)

# Label e campo para selecionar o arquivo CSV
label_arquivo = tk.Label(root, text="Selecione o arquivo com as informações das Concessões")
label_arquivo.grid(row=1, column=0, padx=10, pady=5, sticky='w')
entry_arquivo = tk.Entry(root, width=50)
entry_arquivo.grid(row=1, column=1, padx=10, pady=5)
button_arquivo = tk.Button(root, text="Procurar Arquivo", command=procurar_arquivo)
button_arquivo.grid(row=1, column=2, padx=10, pady=5)

# Label e campo para inserir texto de atribuição
label_atribuicao = tk.Label(root, text="Informar texto para o campo ATRIBUIÇÃO")
label_atribuicao.grid(row=2, column=0, padx=10, pady=5, sticky='w')
entry_atribuicao = tk.Entry(root, width=50)
entry_atribuicao.grid(row=2, column=1, padx=10, pady=5)

# Label e campo para nome do arquivo agrupador
label_nome_arquivo = tk.Label(root, text="Nome do arquivo agrupador (não precisa da extensão .xlsx)")
label_nome_arquivo.grid(row=3, column=0, padx=10, pady=5, sticky='w')
entry_nome_arquivo = tk.Entry(root, width=50)
entry_nome_arquivo.grid(row=3, column=1, padx=10, pady=5)

# Botão para executar a ação
button_executar = tk.Button(root, text="       Executar       ", command=executar)
button_executar.grid(row=4, column=2, padx=10, pady=10)

# Iniciar o loop principal da interface gráfica
root.mainloop()


