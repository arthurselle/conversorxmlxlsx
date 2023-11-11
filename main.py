import xmltodict
import os
import pandas as pd
from datetime import datetime 

#lendo os dados e armazendando-os em variaveis e validando campos 
def getInfos(nome_arquivo, valores):
    #pega cada arquivo xml, abre e aramazena na variavel arquivo_xml
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:
        #transforma o xml em um "dicionário" que o python consegue ler
        dic_xml = xmltodict.parse(arquivo_xml)
        #pegando os dados da nf e armazenando em variaveis       
        try:
            if "NFe" in dic_xml: 
                infos_nfe   = dic_xml["NFe"]['infNFe']
            else:
                infos_nfe = dic_xml["nfeProc"]["NFe"]["infNFe"]
            
            numeroNota  = infos_nfe["ide"]["nNF"]
            emissorCnpj = infos_nfe["emit"]['CNPJ']
            nomeCliente = infos_nfe["dest"]['xNome'] 
            endCliente  = infos_nfe["dest"]['enderDest']
            if "vol" in infos_nfe["transp"]:
                pesoBruto   = infos_nfe["transp"]["vol"]["pesoB"]
            else:
                pesoBruto = "peso não informado";  
            #adicinando os valores em uma array para incrementar na tabela  
            valores.append([numeroNota, emissorCnpj, nomeCliente, endCliente, pesoBruto])
        except Exception as e:
            print(f"Ocorreu(ram) o(s) seguinte(s) erro(s): {e}")

#pega os arquivos xml de dentro do diretório    
listaArquivos = os.listdir("nfs")

#cria as colunas p/ adicionar no xml 
colunas = ["numeroNota", "emissorCnpj", "nomeCliente", "endCliente", "pesoBruto"]
valores = []

#passa por cada arquivo para capturar os dados e armazena-los
for arquivo in listaArquivos:
    getInfos(arquivo, valores)
    #Move os arquivos já lidos para outra pastsa
    os.rename(f'nfs/{arquivo}', f'XML_lidos/{arquivo}')

tabela = pd.DataFrame(columns=colunas, data=valores)

#Nome da pasta para os arquivos exportados
pasta_export = "xlsx_exportados"

#Cria a pasta se não existir
if not os.path.exists(pasta_export):
    os.makedirs(pasta_export)

#Nome do arquivo com a data atual
nome_arquivo = f"Notas_Importadas_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
caminho_arquivo = os.path.join(pasta_export, nome_arquivo)

#Salva a tabela no arquivo XLSX
tabela.to_excel(caminho_arquivo, index=False)