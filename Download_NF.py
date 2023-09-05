
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import win32gui
import pyautogui
import pyperclip
import os
import time
import pandas as pd
from os.path import join
from pywinauto import Desktop
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


pasta = r"C:\Users\usuario\documents\pasta"                # pasta do Relatorio de NF´s
localdestino = r"C:\Users\usuario\documents\pasta\pasta"   # Diretorio p/ armazenar downloads

link = 'https://palmasto.webiss.com.br/externo/nfse/visualizar/NUM_CNPJ'
janela = Tk()

janela.withdraw()
janela.attributes("-topmost", True)

# Exiba a caixa para selecionar o arquivo
Arquivo_Excel = askopenfilename(title="Selecionar Arquivo oom a Lista de NF-e", initialdir=pasta, filetypes=[('Arquivo Excel', '*.xlsx')])

janela.destroy()

excel_file = pd.ExcelFile(Arquivo_Excel)

ultima_aba = excel_file.sheet_names[-1]
print("Última aba:", ultima_aba)
tabela = pd.read_excel(excel_file, sheet_name=ultima_aba)

total_linhas = tabela.shape[0]
print("Total de Linhas:", total_linhas)


edge_options = webdriver.EdgeOptions()
edge_options.use_chromium = True




edge_options.set_capability("ms:edgeOptions", {"args": ["--multiple-automatic-download-confirmation", "false"]})
prefs = {
    'download.default_directory' : localdestino,
    'profile.default_content_setting_values.automatic_downloads': 1,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
}
edge_options.add_experimental_option('prefs', prefs)

driver = webdriver.Edge(options=edge_options)



linha = 0
for i in range(linha, total_linhas):

    NF = tabela.iloc[i, 11]  # Coluna da NF ex'2023000000123456'
    CV = tabela.iloc[i, 10]  # Coluna do Cód. de Verificação ex: 'ABC-DEF'
    RPS = tabela.iloc[i, 4]  # Opcional 
    link_nf = join(str(link), str(CV), str(NF))
    arquivo_validate = 'NotaEletronica_'+str(NF)+'_'+str(CV)+'.pdf' 
  
    if arquivo_validate in os.listdir(localdestino):
        continue
    else:
        driver.get(link_nf)  # Abre o link da NF no Microsoft Edg
        try:
            try:
                conteudo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "nota-eletronica-html")))
            except:
                driver.refresh()
                conteudo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "nota-eletronica-html")))
            conteudo = conteudo.text
            Numero_NF = driver.find_element(By.ID, "Numero")
            CodigoDeVerificacao = driver.find_element(By.ID, "CodigoDeVerificacao")

            # Obtenha o valor do elemento
            Numero_NF = Numero_NF.get_attribute("value")
            CodigoDeVerificacao = CodigoDeVerificacao.get_attribute("value")
            
            arquivo_NFE = os.path.join(localdestino,'NotaEletronica_'+str(Numero_NF)+'_'+str(CodigoDeVerificacao)+'.pdf')
  
            if os.path.isfile(arquivo_NFE):
                breakpoint()
           
            # RPS = conteudo.split("RPS número ")[1].split(" Série")[0]
            # PAGADOR = conteudo.split("Nome/Razão Social")[1].split("\nCPF/CNPJ")[0]
            # nome_arquivo = os.path.join(localdestino,'RPS - '+str(RPS)+' - '+PAGADOR+'.pdf')
            # nome_arquivo = nome_arquivo.replace('\n','')
            # ds_arquivo = str(RPS)+' - '+PAGADOR+'.pdf'
            # ds_arquivo = ds_arquivo.replace('\n','')
            time.sleep(1)
            pyautogui.press('tab', 7)
            pyautogui.press('enter')
            while not os.path.isfile(arquivo_NFE):
                os.path.isfile(arquivo_NFE)

        except:
            driver.refresh()
            conteudo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "numero-documento")))
            print(NF, "Não foi transmitida")
            continue



# Verificação geral
for i in range(total_linhas):
    NF = tabela.iloc[i, 11]  # Coluna L
    CV = tabela.iloc[i, 10]  # Coluna K
    RPS = tabela.iloc[i, 4]  # Coluna E   Opcional
    
    arquivo_validate = 'NotaEletronica_'+str(NF)+'_'+str(CV)+'.pdf'
    if arquivo_validate not in os.listdir(localdestino):
        print(f"A nota fiscal {NF} não foi encontrada na pasta")




