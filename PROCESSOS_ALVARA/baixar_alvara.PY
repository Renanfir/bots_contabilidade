import time
import pyautogui as bot
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import xlrd

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get("https://google.com")
driver.maximize_window()

wdw = WebDriverWait(driver, 30)
visibilidade = EC.visibility_of_element_located

#Escreve o link no url
url = f"https://www.blumenau.sc.gov.br/cidadao/Default.aspx"
driver.get(url)


caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\bots_Contabilidade\COD_NOMES_ALVARA.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('conteudo')


for row in range(sheet.nrows):

    cod = sheet.cell_value(row, 0)
    cod = int(cod)
    cod = str(cod)
    nome = sheet.cell_value(row, 1)
    

    pessoaJuridica = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_divMenuPortalDinamico"]/dd[17]/a'))).click()

    campoCmc = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_painelPesquisaCpfCmc_tbCmc_I"]'))).click()
    campoCmc = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_painelPesquisaCpfCmc_tbCmc_I"]'))).send_keys(cod)
    entra = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_painelPesquisaCpfCmc_btnPesquisa_CD"]/span'))).click()
    seleciona = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_GridView1"]/tbody/tr[2]/td[1]/a'))).click()

    bolinha11 = f'//*[@id="ctl00_ContentBody_callbackPanelGeral_gvLancamentos_DXSelBtn0_D"]'
    
    bolinha12 = wdw.until(visibilidade((By.XPATH, bolinha11))).click()
    bolinha2 = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_callbackPanelGeral_gvParcelas_DXSelBtn0_D"]'))).click()

    entra = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_callbackPanelGeral_btnImprimirBoleto_CD"]/span'))).click()
    sai = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_lnkLogoutPessoa"]'))).click()
time.sleep(10)