#importando bibliotecas

import pandas as pd
import sys
import os
import time
from datetime import datetime, timedelta, date
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from tqdm import trange
from getpass4 import getpass
import threading

#função switch que recebe o estado e a cidade para retornar a regional

def switch(estado, cidade):
    if (estado == "DF") or (estado == "GO") or (estado == "MS") or (estado == "MT"):
        resultado = "CO"
        return resultado
    elif estado == "MG":
        resultado = "MG"
        return resultado
    elif (estado == "AL") or (estado == "BA") or (estado == "CE") or (estado == "PA") or (estado == "PB") or (estado == "PE") or (estado == "PI") or (estado == "RN") or (estado == "SE"):
        resultado = "NORDESTE"
        return resultado
    elif (estado == "MA") or (estado == "AM") or (estado == "AP"):
        resultado = "NORTE"
        return resultado
    elif (estado == "RJ") or (estado == "ES"):
        resultado = "SUDESTE"
        return resultado
    elif (estado == "PR") or (estado == "SC") or (estado == "RS"):
        resultado = "SUL"
        return resultado
    elif (estado == "SP"):
        if (cidade == "SAO PAULO") or (cidade == "ARUJA") or (cidade == "BIRITIBA MIRIM") or (cidade == "CUBATAO") or (cidade == "DIADEMA") or (cidade == "FERRAZ DE VASCONCELOS") or (cidade == "GUARAREMA") or (cidade == "GUARUJA") or (cidade == "GUARULHOS") or (cidade == "ITANHAEM") or (cidade == "ITAQUAQUECETUBA") or (cidade == "MAUA") or (cidade == "MOGI DAS CRUZES") or (cidade == "POA") or (cidade == "PRAIA GRANDE") or (cidade == "REGISTRO") or (cidade == "RIBEIRAO PIRES") or (cidade == "SANTO ANDRE") or (cidade == "SANTOS") or (cidade == "SAO BERNARDO DO CAMPO") or (cidade == "SAO CAETANO DO SUL") or (cidade == "SAO VICENTE") or (cidade == "SUZANO"):
            resultado = "SP CAPITAL"
            return resultado
        else:
            resultado = "SP INTERIOR"
            return resultado
    else:
        resultado = ""
        return resultado
        
        
#função para calcular o aging

def ClassificarAging(aging):
    
    if aging >= 90:
        return "> 90"
    elif aging >= 60:
        return "> 60"
    elif aging >= 45:
        return "> 45"
    elif aging >= 30:
        return "> 30"
    elif aging >= 25:
        return "> 25"
    elif aging >= 20:
        return "> 20"
    elif aging >= 15:
        return "> 15"
    elif aging >= 7:
        return "> 7"
    elif aging >= 5:
        return "> 5"
    elif aging >= 4:
        return "4"
    elif aging >= 3:
        return "3"
    elif aging >= 2:
        return "2"
    elif aging >= 1:
        return "1"
    else:
        return "0"


usuario = ""
senha = ""

def login():

    options = Options()
    options.add_argument('--no-sandbox')
    options.add_argument('--headless')
    options.add_argument("window-size=1200x1200")
    navegador_login = webdriver.Chrome(options=options)
    
    validacao_login = 0
    
    def validar_usuario():
    
        usuario = input("Digite o login WFM:")
        senha = getpass("Digite a senha WFM: ")    
        
        #entrando na página do wfm e realizando login na ferramenta

        navegador_login.get("url do sistema (ocultado por segurança de dados )")
        
        time.sleep(5)

        navegador_login.find_element('xpath','//*[@id="loginForm:username"]').send_keys(usuario)
        navegador_login.find_element('xpath','//*[@id="loginForm:password"]').send_keys(senha)
        navegador_login.find_element('xpath','//*[@id="loginForm:j_idt15"]').click()
        
        time.sleep(5)
        
        try:
            errouSenha = navegador_login.find_element('xpath', '//*[@id="loginForm:j_idt16"]/div[1]/span')
        except NoSuchElementException:
            errouSenha = ""
            
            
        if errouSenha != "":
            print("A senha não foi validada no sistema. Tente Novamente!")
            validar_usuario()
        else:
            print("Usuário autenticado com sucesso!")
            navegador_login.close()
            validacao_login = 1
    
    if validacao_login == 0:
        validar_usuario()

#chamar função de login

login()

#importar a planilha excel

tabela = pd.read_excel("base_robo.xlsx")

valordividido = int(len(tabela)/4)

tabela_robo1 = tabela.loc[:valordividido,:]
tabela_robo2 = tabela.loc[valordividido+1:(valordividido*2),:]
tabela_robo3 = tabela.loc[(valordividido*2)+1:(valordividido*3),:]
tabela_robo4 = tabela.loc[(valordividido*3)+1:,:]

tabela_robo1 = tabela_robo1.reset_index(drop=True)
tabela_robo2 = tabela_robo2.reset_index(drop=True)
tabela_robo3 = tabela_robo3.reset_index(drop=True)
tabela_robo4 = tabela_robo4.reset_index(drop=True)

def executarRobo(tabela, num_robo):

    #iniciando o chrome driver
    
    options = Options()
    options.add_argument('--no-sandbox')
    options.add_argument('--headless')
    options.add_argument("window-size=1200x1200")
    navegador = webdriver.Chrome(options=options)
    
    #entrando na página do wfm e realizando login na ferramenta

    navegador.get("link do sistema")


    navegador.find_element('xpath','//*[@id="loginForm:username"]').send_keys(usuario)
    navegador.find_element('xpath','//*[@id="loginForm:password"]').send_keys(senha)
    navegador.find_element('xpath','//*[@id="loginForm:j_idt15"]').click();

    #sleep para aguardar ferramenta fazer login 

    time.sleep(2)

    #variavel contador que irá contar a quantidade de iterações para chamar a função salvar excel após 100 iterações

    contador = 1


    #inicio do laço for sobre o total de linhas encontradas na planilha
    #trange é a função para criar o barra de progresso

    for item in trange(len(tabela)):

        # receber o número da ordem e converter em string
        texto = tabela.loc[item, "ID"]
        ordem = str(texto)
        
        #acessar a página do wfm com o número da ordem    
        navegador.get("link do sistema"+ordem)
            
        #pegar as informações necessárias 
        
        try:
            status_element = WebDriverWait(navegador, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="j_idt21:0:val_status"]'))
            )
        except TimeoutException:
            continue
        except NoSuchElementException:
            continue
            
        status = navegador.find_element('xpath','//*[@id="j_idt21:0:val_status"]').text
        motivoStatus = navegador.find_element('xpath','//*[@id="j_idt21:0:val_status_reason"]').text
        motivoCancelamento = navegador.find_element('xpath','//*[@id="j_idt21:0:val_cancel_reason"]').text
        segmento = navegador.find_element('xpath','//*[@id="val_segment"]').text
        redeAcesso = navegador.find_element('xpath','//*[@id="j_idt21:0:val_rede"]').text
        estado = navegador.find_element('xpath','//*[@id="j_idt21:0:val_state"]').text
        cidade = navegador.find_element('xpath','//*[@id="j_idt21:0:val_city"]').text
        regional = switch(estado, cidade)
        documento = navegador.find_element('xpath','//*[@id="val_documento"]').text
        produtosOrdem = navegador.find_element('xpath', '//*[@id="j_idt21:0:val_woi_product"]').text
        
        #realizando um try/except pois não são em todas as ordens que aparece o campo motivo de pendencia, então criei uma exceção 
        
        try:
            motivoDaPendencia = navegador.find_element('xpath','//*[@id="j_idt21:0:issuesTable_data"]/tr/td[4]').text
        except NoSuchElementException:
            motivoDaPendencia = ""
            
        
        #buscando a informação de serviço para realizar a validação de ordem MOTF
        
        servicos = navegador.find_elements('xpath','//*[@id="j_idt21:0:j_idt159_data"]/tr')
        
        #variaveis para validar MOTF
        #logica MOTF -> ordem precisa ter uma BL Siebel em desconexão e BL Next para Adicionar
        
        validaMotf = ""
        validaBlSiebel = 0
        validaBlNext = 0
        
        
        #laço for para verificar todas as linhas de serviços para identificar se possui as condições de motf
        
        for t in range(1, len(servicos)+1):
        
            colunaServicos2 = navegador.find_elements('xpath','//*[@id="j_idt21:0:j_idt159_data"]/tr['+str(t)+']/td[2]')
            colunaServicos3 = navegador.find_elements('xpath','//*[@id="j_idt21:0:j_idt159_data"]/tr['+str(t)+']/td[6]')
            
            
            
            for s in range(len(colunaServicos2)):
                
                
                produtoServico = colunaServicos2[s].text
                statusServico = colunaServicos3[s].text
                configOrdem = navegador.find_elements('xpath','//*[@id="j_idt21:0:val_specificationAcronym"]')
                
                for a in range(len(configOrdem)):
                    
                    configOrdemTexto = configOrdem[a].text
                    
                
                arrConfigOrdem = produtoServico.split(" ")
                    
                
                    
                if (arrConfigOrdem[0] == "Power") and (len(arrConfigOrdem) == 3) and (statusServico == "Desconectar"):
                        validaBlSiebel = 1
                    
                if (len(arrConfigOrdem) == 2) and (arrConfigOrdem[1] == "Mbps"):
                        validaBlNext = 1

                if (validaBlSiebel == 1) and (validaBlNext == 1) and (configOrdemTexto == "IN_L"):
                        validaMotf = "Sim"
        


        #definindo as variaveis que serão modificadas conforme a motiventação da ordem
        
        pdRetencao = 0
        pdAgendamento = 0
        pdTecnica = 0
        agendada = 0
        pdEnriquecimento = 0
        cancelada = 0
        execucao = 0
        statusRetencao = 0
        dataDaCriacao = ""
        dataEntradaRetencao = ""
        dataEnriquecimento = ""
        dataRetencao = ""
        dataSaidaEnriquecimento = ""
        validaEntradaEnri = 0
        validaSaidaEnri = 0
        validaCancelamento = 0
        dataCancelamento = ""
        validaEntrada = 0
        
        #buscando as linhas de movimentação da ordem
        
        rows = navegador.find_elements('xpath','//*[@id="j_idt21:0:j_idt203_data"]/tr')
        
        #laço for para contar todas as movimentações que possuem na ordem
        
        for r in range(len(rows)+1,0,-1):
            
            coluna = navegador.find_elements('xpath','//*[@id="j_idt21:0:j_idt203_data"]/tr['+str(r)+']/td[3]')
            coluna2 = navegador.find_elements('xpath','//*[@id="j_idt21:0:j_idt203_data"]/tr['+str(r)+']/td[4]')
            coluna3 = navegador.find_elements('xpath','//*[@id="j_idt21:0:j_idt203_data"]/tr['+str(r)+']/td[2]')   
                   

            for x in range(len(coluna)):
                
                
                if r == len(rows):
                    dataDaCriacao = coluna3[x].text
                
                if coluna[x].text == "Pendente":
                    if coluna2[x].text == "Agendamento":
                        pdAgendamento = pdAgendamento + 1
                            
                        if (validaSaidaEnri == 0) and (validaEntradaEnri == 1):
                            dataSaidaEnriquecimento = coluna3[x].text
                            validaSaidaEnri = 1
                            
                    if coluna2[x].text == "Retencao":
                            
                        if (validaSaidaEnri == 0) and (validaEntradaEnri == 1):
                            dataSaidaEnriquecimento = coluna3[x].text
                            validaSaidaEnri = 1
                            
                        pdRetencao = pdRetencao + 1
                        
                        dataRetencao = coluna3[x].text
                        
                        if validaEntrada == 0:
                            dataEntradaRetencao = coluna3[x].text
                            validaEntrada = 1
                            
                    if coluna2[x].text == "Tecnica":
                        pdTecnica = pdTecnica + 1
                            
                        if (validaSaidaEnri == 0) and (validaEntradaEnri == 1):
                            dataSaidaEnriquecimento = coluna3[x].text
                            validaSaidaEnri = 1
                                
                    if coluna2[x].text == "Enriquecimento":
                        if validaEntrada == 0:
                            pdEnriquecimento = pdEnriquecimento + 1
                            
                        if (validaSaidaEnri == 0) and (validaEntradaEnri == 1):
                            dataSaidaEnriquecimento = coluna3[x].text
                            validaSaidaEnri = 1
                            
                        if validaEntradaEnri == 0:
                            dataEnriquecimento = coluna3[x].text
                            validaEntradaEnri = 1
                            
                if coluna[x].text == "Agendada":
                    agendada = agendada + 1
                    if (validaSaidaEnri == 0) and (validaEntradaEnri == 1):
                        dataSaidaEnriquecimento = coluna3[x].text
                        validaSaidaEnri = 1
                if coluna[x].text == "Execucao":
                    execucao = execucao + 1
                    if (validaSaidaEnri == 0) and (validaEntradaEnri == 1):
                        dataSaidaEnriquecimento = coluna3[x].text
                        validaSaidaEnri = 1
                if coluna[x].text == "Cancelada":
                    
                    if validaCancelamento == 0:
                        dataCancelamento = coluna3[x].text
                        validaCancelamento = 1
                    
                    if validaEntrada == 0:
                        cancelada = cancelada + 1
                        
                        if (validaSaidaEnri == 0) and (validaEntradaEnri == 1):
                            dataSaidaEnriquecimento = coluna3[x].text
                            validaSaidaEnri = 1
                    
                
        
        
        #criando uma váriavel de retorno com todas as informações separadas com ;
        
        """
        consolidadoResposta = documento+";"+segmento+";"+regional+";"+redeAcesso+";"+status+";"+motivoStatus+";"+motivoCancelamento+";"+str(agendada)+";"+
        str(pdTecnica)+";"+str(pdAgendamento)+";"+str(execucao)+";"+str(dataEntradaRetencao)+";"+str(dataDaCriacao)+";"+str(dataEnriquecimento)+";"+
        str(dataSaidaEnriquecimento)+";"+str(dataCancelamento)+";"+motivoDaPendencia+";"+validaMotf
        """
        
        tabela.loc[item, "DOCUMENTO"] = documento
        tabela.loc[item, "SEGMENTO"] = segmento
        tabela.loc[item, "REGIONAL"] = regional
        tabela.loc[item, "REDE_ACESSO"] = redeAcesso
        tabela.loc[item, "STATUS"] = status
        tabela.loc[item, "MOTIVO_STATUS"] = motivoStatus
        tabela.loc[item, "MOTIVO_CANCELAMENTO"] = motivoCancelamento
        tabela.loc[item, "AGENDADA"] = agendada
        tabela.loc[item, "PD_ENRIQUECIMENTO"] = pdEnriquecimento
        tabela.loc[item, "PD_RETENCAO"] = pdRetencao
        tabela.loc[item, "PD_TECNICA"] = pdTecnica
        tabela.loc[item, "PD_AGENDAMENTO"] = pdAgendamento
        tabela.loc[item, "EXECUCAO"] = execucao
        tabela.loc[item, "DATA_CRIACAO"] = dataDaCriacao
        tabela.loc[item, "DATA_ENTRADA_RETENCAO"] = dataEntradaRetencao
        tabela.loc[item, "DATA_ULTIMA_RETENCAO"] = dataRetencao
        tabela.loc[item, "DATA_CANCELAMENTO"] = dataCancelamento
        tabela.loc[item, "DATA_ENRIQUECIMENTO"] = dataEnriquecimento
        tabela.loc[item, "DATA_SAIDA_ENRIQUECIMENTO"] = dataSaidaEnriquecimento
        tabela.loc[item, "MOTIVO_PD_TECNICO"] = motivoDaPendencia
        tabela.loc[item, "VALIDA_MOTF"] = validaMotf  
        tabela.loc[item, "PRODUTOS"] = produtosOrdem
        
        
        #funções para calcular aging e classificar
        
        hoje = date.today();

        if dataDaCriacao != "":
            date_dataDaCriacao = datetime.strptime(dataDaCriacao, "%d/%m/%Y %H:%M:%S").date()
            
        if dataEntradaRetencao != "":
            date_dataEntradaRetencao = datetime.strptime(dataEntradaRetencao, "%d/%m/%Y %H:%M:%S").date()
            
        if dataCancelamento != "":
            date_dataCancelamento = datetime.strptime(dataCancelamento, "%d/%m/%Y %H:%M:%S").date()
            
        if dataEnriquecimento != "":
            date_dataEnriquecimento = datetime.strptime(dataEnriquecimento, "%d/%m/%Y %H:%M:%S").date()
        
        if dataSaidaEnriquecimento != "":
            date_dataSaidaEnriquecimento = datetime.strptime(dataSaidaEnriquecimento, "%d/%m/%Y %H:%M:%S").date()
            
        if dataRetencao != "":
            date_dataRetencao = datetime.strptime(dataRetencao, "%d/%m/%Y %H:%M:%S").date()

            
        
        if (dataEntradaRetencao != "") and (dataDaCriacao != ""):
            agingEntradaRetencao = int((date_dataEntradaRetencao - date_dataDaCriacao) / timedelta(days=1))
        else:
            agingEntradaRetencao = 0
            
        if (dataCancelamento != "") and (dataDaCriacao != ""):
            agingCancelamento = int((date_dataCancelamento - date_dataDaCriacao) / timedelta(days=1))
        else:
            agingCancelamento = 0
            
        if (dataSaidaEnriquecimento != "") and (dataEnriquecimento != ""):
            agingEnriquecimento = int((date_dataSaidaEnriquecimento - date_dataEnriquecimento) / timedelta(days=1))
        else:
            agingEnriquecimento = 0
            
        if (dataRetencao != ""):
            agingPVA = int((hoje - date_dataRetencao) / timedelta(days=1))
        else:
            agingPVA = 0
            
        
        agingClass_EntradaRetencao = ClassificarAging(agingEntradaRetencao)
        agingClass_Cancelamento = ClassificarAging(agingCancelamento)
        agingClass_Enriquecimento = ClassificarAging(agingEnriquecimento)
        agingClass_PVA = ClassificarAging(agingPVA)
        
        tabela.loc[item, "AGING_ENTRADA_RETENCAO"] = agingEntradaRetencao
        tabela.loc[item, "AGING_CANCELAMENTO"] = agingCancelamento
        tabela.loc[item, "AGING_ENRIQUECIMENTO"] = agingEnriquecimento
        tabela.loc[item, "AGING_CLASS_ENTRADA_RETENCAO"] = agingClass_EntradaRetencao
        tabela.loc[item, "AGING_CLASS_CANCELAMENTO"] = agingClass_Cancelamento
        tabela.loc[item, "AGING_CLASS_ENRIQUECIMENTO"] = agingClass_Enriquecimento
        tabela.loc[item, "AGING_PVA"] = agingClass_PVA
            
        
        #identifico a coluna do excel que irei colocar a resposta
        """    
        tabela.loc[item, "RETORNO"] = consolidadoResposta
        """
        #if para saber se o contador chegou a 100 para salvar o excel
        
        if contador==10:
            tabela.to_excel(f"base_robo_atualizada{num_robo}.xlsx", index=False)
            contador = 1
        else:
            contador += 1


    navegador.close()
    tabela.to_excel(f"base_robo_atualizada{num_robo}.xlsx", index=False)


robo1 = threading.Thread(target=executarRobo, args=(tabela_robo1,1,))
robo1.start()

time.sleep(10)

robo2 = threading.Thread(target=executarRobo, args=(tabela_robo2,2,))
robo2.start()

time.sleep(10)

robo3 = threading.Thread(target=executarRobo, args=(tabela_robo3,3,))
robo3.start()

time.sleep(10)

robo4 = threading.Thread(target=executarRobo, args=(tabela_robo4,4,))
robo4.start()


robo1.join()
robo2.join()
robo3.join()
robo4.join()

tabela = pd.read_excel("base_robo_atualizada1.xlsx")

for x in range(2, 5):

    num = str(x)
	
    nova_tabela = pd.read_excel(f"base_robo_atualizada{num}.xlsx")
        
    tabela = pd.concat([tabela, nova_tabela])
    

tabela.to_excel("base_robo_atualizada_final.xlsx", index=False)

os.remove("base_robo_atualizada1.xlsx")
os.remove("base_robo_atualizada2.xlsx")
os.remove("base_robo_atualizada3.xlsx")
os.remove("base_robo_atualizada4.xlsx")