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

def verifica_valor(valor):
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
        
        # Para ambos .xlsx e .xls, use pd.read_excel()
        if str(avc[1]).lower().endswith(('.xlsx', '.xls')):
            try:
                # Para .xls, tente engine 'xlrd' ou 'openpyxl'
                if str(avc[1]).lower().endswith('.xls'):
                    try:
                        dados = pd.read_excel(avc[1], engine='xlrd')
                    except:
                        # Se 'xlrd' não funcionar, tente sem engine
                        dados = pd.read_excel(avc[1])
                else:
                    dados = pd.read_excel(avc[1])
                df = dados
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler arquivo {avc[0]}: {str(e)}")
                # Tente instalar: pip install xlrd
                continue
        else:
            messagebox.showinfo("Formato não suportado", f"Formato não suportado: {avc[0]}")
            continue
        
        if df is None or df.empty:
            continue
            
        # Inicializa as variáveis de data
        dt_vencimento_1 = None
        dt_vencimento_2 = None
        dt_vencimento_3 = None
        
        # Itera pelas linhas do DataFrame
        for i in range(len(df)):
            # Para acessar valores em um DataFrame, use .iloc ou .loc
            # Verifica se a coluna 1 (segunda coluna) tem o valor 'Usuárias'
            # Primeiro, verifique se há colunas suficientes
            #if df.shape[1] < 12:
            #    continue
                
            if str(df.iloc[i, 1]) == 'Usuárias':
                dt_vencimento_1 = str(df.iloc[i, 3])[-10:] if len(str(df.iloc[i, 3])) >= 10 else ''
                dt_vencimento_2 = str(df.iloc[i, 4])[-10:] if len(str(df.iloc[i, 4])) >= 10 else ''
                dt_vencimento_3 = str(df.iloc[i, 5])[-10:] if len(str(df.iloc[i, 5])) >= 10 else ''
            elif pd.isna(df.iloc[i, 0]):  # Verifica se é NaN
                continue
            elif str(df.iloc[i, 1])[0:7] == 'Empresa':
                continue
            elif str(df.iloc[i, 0]) == 'Total Geral':
                continue
            elif str(df.iloc[i, 1])[0:10] == 'Relatório':
                continue
            elif pd.isna(df.iloc[i, 4]):  # Verifica se é NaN
                continue
            else:
                #primeiro_vencimento = verifica_valor(df.iloc[i, 3])
                #segundo_vencimento = verifica_valor(df.iloc[i, 4])
                #terceiro_vencimento = verifica_valor(df.iloc[i, 5])
                primeiro_vencimento = df.iloc[i, 3]
                segundo_vencimento = df.iloc[i, 4]
                terceiro_vencimento = df.iloc[i, 5]
                dados_clientes.append([
                    df.iloc[i, 0],  # Coluna 0
                    df.iloc[i, 1],  # Coluna 1
                    df.iloc[i, 2],  # Coluna 2
                    primeiro_vencimento,
                    segundo_vencimento,
                    terceiro_vencimento,
                    material_e_centro_lucro[0] if material_e_centro_lucro else None,
                    material_e_centro_lucro[1] if material_e_centro_lucro else None,
                    atribuicao
                ])
                
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


