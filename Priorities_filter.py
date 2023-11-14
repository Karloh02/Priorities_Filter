import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt 
from datetime import datetime
import win32com.client as wincl
import time
from AppOpener import close
import os, os.path
import os, shutil
import os
import ezdxf
import math

#Le o arquivo excel que contem todas as informacoes necessarias
pasta_excel = pd.read_excel(r"C:\Users\U57534\OneDrive - Bühler\Desktop\PROJETOS\MFT_Implementation\TA_COPY.xlsx")

#Transforma em arrays todas as informacoes de interesse
APCA_codigo_excel = (pasta_excel["Codigo SAP"].tolist())
Chiffre_excel = (pasta_excel["Chifra"].tolist())

#Arrays das chifras
Text_Files_ = []

#Testes
#print(Chiffre_excel)

#Funcao que irá fazer o loop para salvar os APCA's de todas as máquinas de uma certa chifra. 
def Separa_APCA(array_strings_chifra, APCA, Chiffre, Text_Files_):
    array_export = []
    for i in range(len(array_strings_chifra)):
        var_string = array_strings_chifra[i]
        for k in range(len(Chiffre)):
            var_append = APCA[k]
            if str(Chiffre[k]) == var_string:
                variavel = APCA[k]
                variavel = variavel[0:14]
                array_export.append(variavel)
 
    name = array_strings_chifra[0] + ".txt"
    Text_Files_.append(name)
    new_array = np.array(array_export)
    new_array = list(dict.fromkeys(new_array))
    np.savetxt(name, new_array, delimiter=" ", fmt="%s")
    return(Text_Files_)

#salva a data correta para usar no SAP
#datas[0] é a data de hoje escrita no padrão SAP

datas = [0]
def acerta_data(datas):
    data_agora = str(datetime.now())
    data_SAP = ""
    data_SAP = data_agora[8:10] + "." + data_agora[5:7] + "." + data_agora[0:4]

    datas[0] = data_SAP
    return(datas)
acerta_data(datas)

#Organiza todos as chifras de interesse em arquivos de texto

#Separa_APCA(["DRHK"], APCA_codigo_excel, Chiffre_excel, Text_Files_) Já foi feito
#Separa_APCA(["TAS", "LAAA", "LAAB", "LAAC"], APCA_codigo_excel, Chiffre_excel, Text_Files_) Já foi feito
#Separa_APCA(["LADB"], APCA_codigo_excel,Chiffre_excel, Text_Files_)
#Separa_APCA(["AHSG"], APCA_codigo_excel,Chiffre_excel, Text_Files_)
#Separa_APCA(["AHAS"], APCA_codigo_excel,Chiffre_excel, Text_Files_)
#Separa_APCA(["MZAL"], APCA_codigo_excel,Chiffre_excel, Text_Files_)
#print(Text_Files_)

#x = np.loadtxt(Text_Files_[0],dtype="str", delimiter=" ", usecols=(0), unpack=True)
#print(x)

Local_Salva_Arquivos = []

#funcao para importar todos os arquios do SAP
def importando_arquivos(datas, APCA_utilizados, Local_Salva_Arquivos):
    onde_salvo = []

    #carrega o texto com todos os APCA's
    carregando_texto = np.loadtxt(APCA_utilizados, dtype="str", delimiter=" ", usecols=(0), unpack = True)
    #carregando_texto = list(dict.fromkeys(carregando_texto))
    print(carregando_texto)
    #cria o nome da chiffre a ser usada depois para salvar os aqruivos
    #nome_chiffre = APCA_utilizados.removesuffix(".txt")
    #abre as lista com todos os arquivos da Chiffre
    
    for i in range(len(carregando_texto)):
        nome_chiffre = APCA_utilizados.removesuffix(".txt")
        Cod_APCA = str(carregando_texto[i])
        Diretorio = [r"C:\Users\U57534\MFT_Temp"]
        new_name = nome_chiffre + str(i) + ".xlsx"  #Texto
        Data_Sap = datas
        Centro = 2103
        Aplicacao = "PP01"

        #Cria a planilha com os dados

        Planilha_Flutuante = str(Diretorio[0]) + "/" + "Planilha_Flutuante.xlsm"
        workbook = xlsxwriter.Workbook(Planilha_Flutuante)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Local")
        worksheet.write(1, 0, "Centro")
        worksheet.write(2, 0, "Comando")
        worksheet.write(3, 0, "Código")
        worksheet.write(4, 0, "Texto")
        worksheet.write(5, 0, "Data")
        worksheet.write(0, 1, Diretorio[0])
        worksheet.write(1, 1, Centro)
        worksheet.write(2, 1, Aplicacao)
        worksheet.write(3, 1, Cod_APCA)
        worksheet.write(4, 1, new_name)
        worksheet.write(5, 1, Data_Sap)

        #adiciona o codigo VBA na planilha
        
        workbook.add_vba_project(r"C:\Users\U57534\OneDrive - Bühler\Desktop\PROJETOS\MFT_Implementation\Integrando_SAP\xl\vbaProject.bin")
        workbook.close()

        #Tempo para permitir que todos os arquivos fiquem salvos
        
        time.sleep(3)

        excel_macro = wincl.DispatchEx("Excel.application")
        excel_path = os.path.expanduser(Planilha_Flutuante)
        workbook2 = excel_macro.Workbooks.Open(Filename = excel_path, ReadOnly = 1)
        excel_macro.Application.Run("SAPDownloadAttachment")
        excel_macro.Application.Quit()

        time.sleep(5)
        close("Excel")
        print(new_name)
        os.remove(Planilha_Flutuante)
        time.sleep(3)
    
        print("Starting...")
        text_dir = str(Diretorio[0]) + "/" + new_name
        print(text_dir)
        onde_salvo.append(text_dir)
    texto_arquivo = nome_chiffre + "_Codigos.txt"
    print(onde_salvo)
    np.savetxt(texto_arquivo, onde_salvo, delimiter=" ", fmt="%s")
    
    return()
#criando o excel como todos os dados da DRHK

for j in range(len(Text_Files_)):
    importando_arquivos(datas[0], Text_Files_[j], Local_Salva_Arquivos)  #Posição 1 da lista de chiffras

#trecho feito para ler todas as planilhas de uma determinada Chiffre + Cria nova planilha com todas as informações Todos_Codigos_"Chiffre".xlsx

def cria_planilhas_todos_codigos(texto, local):
    codigos_load = np.loadtxt(local, dtype="str", delimiter=" ", usecols=(0), unpack=True)

    #Agora precisamos ler as planilhas salvas e ir colocando todas as informações nas respectivas colunas 
    Nivel_explosao = []     #OK
    Numero_item = []        #OK
    Categoria_item = []     #OK Numero Do item
    Texto_breve = []        #OK
    Unidade = []            #OK
    Quantidade = []         #OK
    Tipo_suprimento = []    #OK

    numero = 0
    for q in range(len(codigos_load)):
        #Lê a planilha
        planilha_do_codigo = pd.read_excel(codigos_load[q])

        #Lê as informações da planilha
        Nivel_explosao_individual = (planilha_do_codigo["Nível explosão"].tolist())
        Numero_item_individual = (planilha_do_codigo["Nº item"].tolist())
        Categoria_item_individual = (planilha_do_codigo["Nº componente"].tolist())
        Texto_breve_individual = (planilha_do_codigo["Texto breve objeto"].tolist())
        Unidade_individual = (planilha_do_codigo["Unid.medida básica"].tolist())
        Quantidade_individual = (planilha_do_codigo["Qtd.componente"].tolist())
        Tipo_suprimento_individual = (planilha_do_codigo["Tipo de suprimento especial (mestre de m"].tolist())
        numero += (len(Nivel_explosao_individual))

        #Salva as informaçõews da planilha no respectivo array. 

        for t in range(len(Nivel_explosao_individual)):
            Nivel_explosao.append(Nivel_explosao_individual[t])
            Numero_item.append(Numero_item_individual[t])
            Categoria_item.append(Categoria_item_individual[t])
            Texto_breve.append(Texto_breve_individual[t])
            Unidade.append(Unidade_individual[t])
            Quantidade.append(Quantidade_individual[t])
            Tipo_suprimento.append(Tipo_suprimento_individual[t])
            
        #print(len(Nivel_explosao))
    #criando nova planilha 
    Diretorio = [r"C:\Users\U57534\MFT_Temp"]
    Planilha_dados = str(Diretorio[0]) + "/" + "Todos_Codigos_" + texto + ".xlsx"
    workbook_planilha_codigos = xlsxwriter.Workbook(Planilha_dados)
    worksheet_planilha_codigos = workbook_planilha_codigos.add_worksheet(texto)

    #                               y  x
    worksheet_planilha_codigos.write(0, 0, "Nível Explosão")
    worksheet_planilha_codigos.write(0, 1, "Número item")
    worksheet_planilha_codigos.write(0, 2, "Categoria do item")
    worksheet_planilha_codigos.write(0, 3, "Número do componente")
    worksheet_planilha_codigos.write(0, 4, "Texto Breve")
    worksheet_planilha_codigos.write(0, 5, "Unidade")
    worksheet_planilha_codigos.write(0, 6, "Quantidade")
    worksheet_planilha_codigos.write(0, 7, "Tipo de Suprimento")

    #Deixa a planilha bonitinha S2
    worksheet_planilha_codigos.set_column("A:A", 14)
    worksheet_planilha_codigos.set_column("B:B", 12)
    worksheet_planilha_codigos.set_column("C:C", 16)
    worksheet_planilha_codigos.set_column("D:D", 23)
    worksheet_planilha_codigos.set_column("E:E", 45)
    worksheet_planilha_codigos.set_column("F:F",  8)
    worksheet_planilha_codigos.set_column("G:G", 11)
    worksheet_planilha_codigos.set_column("H:H", 18)

    #Depois leremos essas planilhas e teremos os dados de área/perímetro/tempo de corte/numero de dobras
    for q in range(len(Nivel_explosao)):
        worksheet_planilha_codigos.write(q+1, 0, str(Nivel_explosao[q]))
        worksheet_planilha_codigos.write(q+1, 1, str(Numero_item[q]))
        worksheet_planilha_codigos.write(q+1, 2, str(Categoria_item[q]))
        worksheet_planilha_codigos.write(q+1, 3, str(Numero_item[q]))
        worksheet_planilha_codigos.write(q+1, 4, str(Texto_breve[q]))
        worksheet_planilha_codigos.write(q+1, 5, str(Unidade[q]))
        worksheet_planilha_codigos.write(q+1, 6, str(Quantidade[q]))
        worksheet_planilha_codigos.write(q+1, 7, str(Tipo_suprimento[q]))
    print(len(Numero_item))
    workbook_planilha_codigos.close()
    
    #Salva o array com todos os itens dentro 

    return()
Locais_texto_Planilhas = [r"C:\Users\U57534\AHAS_Codigos.txt", r"C:\Users\U57534\AHSG_Codigos.txt", r"C:\Users\U57534\DRHK_Codigos.txt", r"C:\Users\U57534\LADB_Codigos.txt",r"C:\Users\U57534\MZAL_Codigos.txt", r"C:\Users\U57534\TAS_Codigos.txt"]
#cria_planilhas_todos_codigos("AHAS", Locais_texto_Planilhas[0])
#cria_planilhas_todos_codigos("AHSG", Locais_texto_Planilhas[1])
#cria_planilhas_todos_codigos("DRHK", Locais_texto_Planilhas[2])
#cria_planilhas_todos_codigos("LADB", Locais_texto_Planilhas[3])
#cria_planilhas_todos_codigos("MZAL", Locais_texto_Planilhas[4])
#cria_planilhas_todos_codigos("TAS" , Locais_texto_Planilhas[5])

#planilhas com todos os dados de todos os APCA's salva. 

#Cálculo das áreas e dos perímetros de cada um dos componentes status 50

#setando cada código de chapa com espessura e material. 

#Espessura e códigos conferidos.
codigos_chapas = ["UOR -11057-173", "UOR -11057-471", "UOR -11057-476", "UOR -11057-481", "UOR -11057-491", "UOR -11057-495", "UOR -11057-029", "UOR -11057-151", "UOR -11057-101", "UOR -11057-159", "UOR -11057-155", "UOR -11057-163", "UOR -11057-028", "UOR -11057-180", "UOR -11057-176", "UOR -11057-171","UOR -11057-170", "UOR -11057-072",
                  "UOR -11000-232", "UOR -11000-007", "UOR -11000-009", "UOR -11000-234","UOR -11000-238","UOR -11000-246", "UOR -11000-250","UOR -11000-254", "UOR -11000-256", "UOR -11000-258", "UOR -11000-262", "UOR -11000-264", "UOR -11000-266","UOR -11000-270","UOR -11000-274", "UOR -11000-278", "UOR -11000-282", "UOR -11000-286", "UOR -11000-288", "UOR -11000-290", "UOR -11000-292", "UOR -11000-295","UOR -11000-297", "UOR -11000-430", "UOR -11000-434", "UOR -11000-436", "UOR -11000-438", "UOR -11000-632", "UOR -11000-634", "UOR -11000-636", "UOR -11000-638", "UOR -11000-646", "UOR -11000-878", "UOR -11000-964", "UOR -11000-966", "UOR -11000-013", "UOR -11000-015", "UOR -11000-019", "UOR -11000-021", "UOR -11000-024", "UOR -11000-027", "UOR -11000-230", "UOR -11000-017", "UOR -11000-006", "UOR -11000-008", "UOR -11000-010", "UOR -11000-012", "UOR -11000-034", "UOR -11000-035", "UOR -11000-036", "UOR -11000-037", "UOR -11000-038", "UOR -11000-063", "UOR -11000-064", "UOR -11000-242"]
espessura_chapa = [2, 5, 6, 5, 5, 5, 12, 1, 2.5, 1.5, 1, 2, 10, 8, 6, 5, 4, 3, 8, 1, 1, 10, 12, 20,22, 25, 3, 6, 12, 15, 20, 50, 6, 12, 20, 75, 6, 12, 20, 101, 12, 6, 10, 11, 12, 8, 10, 11, 12, 20, 12, 25, 50, 1, 1.5, 3, 3, 4, 5, 6, 2, 1, 1.5, 2, 2.5, 8, 10, 4, 6, 3, 12, 5, 15]
material_chapa = ["Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Inox", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carnono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono", "Carbono"]

#irá ler a planilha com todos os códigos dos componentes e irá realizar o cálculo para todos os materiais =>
#Lê os códigos de materiais 
#Se existir em [codigos chapas] 
#Ler o código acima; procurar pelo dxf - calcular perímetro - calcular área. 
#fazer isso para todas os códigos. 


def searchfiles(extension, folder, start):
    "Create a txt file with all the file of a type"
    a = []
    with open(extension[1:] + "file.txt", "w", encoding="utf-8") as filewrite:
        for r, d, f in os.walk(folder):
            for file in f:
                if file.endswith(extension) and file.startswith(start):
                    filewrite.write(f"{r + file}\n")
                    dir =str(folder + "/" + file)
                    a.append("a")
    if len(a)>0:
        dir = dir
    else:
        dir = "a"
    return(dir)

def adicionando_dados_na_planilha(nome, diretorio_planilha, codigos_chapas, espessura_chapas, material_chapa, diretorio_procura):
    nome_arquivo_planilha = r"C:\Users\U57534\MFT_Temp" + "/" + nome + "_Com_dados.xlsx"

    leitura = pd.read_excel(diretorio_planilha)

    nivel_explosao = (leitura["Nível Explosão"].tolist())
    numero_item = (leitura["Número item"].tolist())
    codigo_item = (leitura["Categoria do item"].tolist())
    texto_breve = (leitura["Texto Breve"].tolist())
    unidade = (leitura["Unidade"].tolist())
    quantidade = (leitura["Quantidade"].tolist())
    tipo_de_suprimento = (leitura["Tipo de Suprimento"].tolist())
    
    #criar planilha Chiffre_Tempos -> Colocando as informações de tempo de laser, perimetro area

    workbook = xlsxwriter.Workbook(nome_arquivo_planilha)
    worksheet = workbook.add_worksheet(nome + "Com parâmetros")
    worksheet.set_column("A:A", 17) #nivel_explosao
    worksheet.set_column("B:B", 12) #numero_item
    worksheet.set_column("C:C", 16) #codigo_item
    worksheet.set_column("D:D", 20) #massa do item
    worksheet.set_column("E:E", 20) #tipo de suprimento
    worksheet.set_column("F:F", 20) #área do componente 
    worksheet.set_column("G:G", 20) #perimetro do componente
    worksheet.set_column("H:H", 20) #tempo de corte
    worksheet.set_column("I:I", 20) #numero de dobras

    worksheet.write(0, 0, "Nível Explosão")
    worksheet.write(0, 1, "Número do item")
    worksheet.write(0, 2, "Código do item")
    worksheet.write(0, 3, "Massa do item")
    worksheet.write(0, 4, "Tipo de suprimento")
    worksheet.write(0, 5, "Área do componente")
    worksheet.write(0, 6, "Perímetro do componente")
    worksheet.write(0, 7, "Tempo de corte")
    worksheet.write(0, 8, "Número de dobras")
    
    nivel_explosao_fabricados = []
    numero_item_fabricados = []
    codigo_item_fabricado = []
    massa_componente_fabricado = []
    tipo_de_suprimento_fabricado = []
    area_componente_fabricado = []
    perimetro_total_fabricado = []
    tempo_de_corte_fabricado = []
    dobras_fabricado = []



    for i in range(len(nivel_explosao)):
        if codigo_item[i] in codigos_chapas:
            posicao = codigos_chapas.index(codigo_item[i])
            material = material_chapa[posicao]
            espessura = espessura_chapa[posicao]

            if material == "Inox":
                massa_esp = 8000
            else:
                massa_esp = 7850
            
            massa_componente = quantidade[i]
            codigo_componente = codigo_item[i - 1]

            area_componente = (massa_componente/(massa_esp*(espessura/1000)))*2
            #para calcular o perímetro precisamos encontrar o arquivo dxf e colocar na função para calcular o tempo de corte
            #procurar os arquivos em Desenhos_Windchill que comecem com (codigo do item e eterminem com .dxf)
            #salvar todas as ocorrencias 
            #Caso len ocorrencias >0 utilizar a primeira ocorrencia do componente
            #Fazer os calculos

            ocorrencias = searchfiles(".dxf", diretorio_procura, codigo_componente)
            
            if ocorrencias != "a" and os.path.isfile(ocorrencias): 
                arquivo = ocorrencias
                #leitor de dxf para o arquivo
                espessuras_chapas_corte=["1","1.5","2","2.5","3","4","5","6","8","10","12","15"]
                tempo_corte_carb=[0.35,0.37,0.40,0.42,0.45,0.51,0.58,0.66,0.85,1.1,1.42,2.08]
                tempo_furo_carb=[0.0002,0.0005,0.0006,0.0007,0.0008,0.0023,0.0035,0.0043,0.0087,0.0286,0.0624,0.298]
                tempo_corte_inox=[0.25,0.29,0.33,0.39,0.45,0.6,0.81,1.08,1.96,3.54,6.4]
                tempo_furo_inox=[0.00072,0.00087,0.00122,0.00149,0.00176,0.0026,0.0039,0.0052,0.00823,0.013,0.0273]

                for file in os.listdir():
                    longitud_total = 0
                    file = arquivo
                    dwg = ezdxf.readfile(file)
                    msp = dwg.modelspace()
                    num = 0
                    material_novo = material
                    espessura_nova = espessura
                    dobras = 0
                    #print(material_novo, espessura_nova)
                    for e in (msp):

                        if e.dxf.layer == "IV_BEND_DOWN" or e.dxf.layer == "IV_BEND":
                            dobras += 1
                        
                        if e.dxf.layer == "OUTER" or e.dxf.layer == "IV_ARC_CENTERS" or e.dxf.layer == "IV_INTERIOR_PROFILES" or e.dxf.layer == "IV_TOOL_CENTER" or e.dxf.layer == "IV_TOOL_CENTER_DOWN" or e.dxf.layer == "IV_FEATURE_PROFILES" or e.dxf.layer == "AM_KONT":
                            if e.dxf.layer == "IV_INTERIOR_PROFILES":
                                num += 1
                            
                            if e.dxftype() == "LINE":
                                dl = math.sqrt((e.dxf.start[0]-e.dxf.end[0])**2 + (e.dxf.start[1]- e.dxf.end[1])**2) 
                                if dl < 0: 
                                    dl = -dl
                                else:
                                    dl=dl
                                longitud_total = longitud_total + dl
                            elif e.dxftype() == "ARC":
                                raio=round(e.dxf.radius)
                                start=round(e.dxf.start_angle)
                                end=round(e.dxf.end_angle)
                                Ang=start-end

                                if end == 0:
                                    end=360
                                    Ang=start-end
                                    if Ang > 0:
                                        Ang=Ang
                                    else:
                                        Ang=-Ang
                                da= 2*math.pi*raio*Ang/360
                                if da < 0:
                                    da=-da
                                else:
                                    da=da
                                longitud_total=longitud_total+da
                            
                            elif e.dxftype() == "SPLINE":
                                puntos = e.get_control_points()
                                for i in range(len(puntos)-1):
                                    ds = math.sqrt((puntos[i][0]-puntos[i+1][0])**2 + (puntos[i][1]- puntos[i+1][1])**2) 
                                    longitud_total = longitud_total + ds

                longitud_total = round(longitud_total, 2)
                if material_novo == "Carbono":
                    variavel = espessuras_chapas_corte.index(str(espessura_nova))
                    tempo_de_corte = longitud_total/1000*(tempo_corte_carb[variavel]) + tempo_furo_carb[variavel]*(num + 1)
                    #print(tempo_de_corte, "tentando")
                elif material_novo == "Inox":
                    variavel = espessuras_chapas_corte.index(str(espessura_nova))
                    tempo_de_corte = longitud_total/1000*(tempo_corte_inox[variavel]) + tempo_furo_inox[variavel]*(num)
                else:
                    print("ERRO")

            else:
                dobras = "Erro"
                longitud_total = "Erro" 
                tempo_de_corte = "Erro"

            nivel_explosao_fabricados.append(str(nivel_explosao[i-1]))
            numero_item_fabricados.append(str(numero_item[i - 1]))
            codigo_item_fabricado.append(str(codigo_item[i - 1]))
            massa_componente_fabricado.append(str(massa_componente))
            tipo_de_suprimento_fabricado.append(str(tipo_de_suprimento[i - 1]))
            area_componente_fabricado.append(str(round(area_componente, 2)))
            perimetro_total_fabricado.append(str(longitud_total))
            tempo_de_corte_fabricado.append(str(tempo_de_corte))
            dobras_fabricado.append(str(dobras))
            print("Calculado com sucesso",longitud_total, i)
    for u in range(len(nivel_explosao_fabricados)):
        worksheet.write(u + 1, 0, nivel_explosao_fabricados[u])     #nivel explosao
        worksheet.write(u + 1, 1, numero_item_fabricados[u])     #Numero do item  
        worksheet.write(u + 1, 2, codigo_item_fabricado[u])     #Codigo do item
        worksheet.write(u + 1, 3, massa_componente_fabricado[u])     #Massa do item
        worksheet.write(u + 1, 4, tipo_de_suprimento_fabricado[u])     #Tipo de suprimento
        worksheet.write(u + 1, 5, area_componente_fabricado[u])     #Área do componnete
        worksheet.write(u + 1, 6, perimetro_total_fabricado[u])     #perimetro do componente
        worksheet.write(u + 1, 7, tempo_de_corte_fabricado[u])     #Tempo de corte
        worksheet.write(u + 1, 8, dobras_fabricado[u])     #Numero de dobras


    workbook.close()
    return()

adicionando_dados_na_planilha("DRHK", r"C:\Users\U57534\MFT_Temp\Todos_Codigos_DRHK.xlsx", codigos_chapas, espessura_chapa, material_chapa, r"\\ctbn33\AVOR\__Desenhos_Windchill\DRHK\DRHK-10153-001 - 16.08.2023")
#adicionando_dados_na_planilha("MZAL", r"C:\Users\U57534\MFT_Temp\Todos_Codigos_MZAL.xlsx", codigos_chapas, espessura_chapa, material_chapa, r"\\ctbn33\AVOR\__Desenhos_Windchill\MZAL")
adicionando_dados_na_planilha("LADB", r"C:\Users\U57534\MFT_Temp\Todos_Codigos_LADB.xlsx", codigos_chapas, espessura_chapa, material_chapa, r"\\ctbn33\AVOR\__Desenhos_Windchill\LADB")
adicionando_dados_na_planilha("AHSG", r"C:\Users\U57534\MFT_Temp\Todos_Codigos_AHSG.xlsx", codigos_chapas, espessura_chapa, material_chapa, r"\\ctbn33\AVOR\__Desenhos_Windchill\AHSG")
adicionando_dados_na_planilha("AHAS", r"C:\Users\U57534\MFT_Temp\Todos_Codigos_AHAS.xlsx", codigos_chapas, espessura_chapa, material_chapa, r"\\ctbn33\AVOR\__Desenhos_Windchill\AHAS")
#adicionando_dados_na_planilha("TAS", r"C:\Users\U57534\MFT_Temp\Todos_Codigos_TAS.xlsx", codigos_chapas, espessura_chapa, material_chapa, r"\\ctbn33\AVOR\__Desenhos_Windchill")'