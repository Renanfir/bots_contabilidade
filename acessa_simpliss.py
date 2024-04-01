import time
import pyautogui as bot
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import xlrd




#Entra no google
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.get("https://google.com")
driver.maximize_window()

wdw = WebDriverWait(driver, 15)
visibilidade = EC.visibility_of_element_located

#Escreve o link no url
url = f"https://nfse.blumenau.sc.gov.br/contrib/Account/Login"
driver.get(url)

#Preenche o CPF
campo_cpf = driver.find_element(By.ID, 'Login_UserName')
campo_cpf.send_keys('30647741000120')

#Preenche a senha EXCLUIR A SENHA SE POSTAR NO GITHUB!!!!!!!
campo_senha = driver.find_element(By.ID, "Login_Password")
campo_senha.send_keys('xxxxxxxxx')

#Clica no botão enter
acessar = driver.find_element(By.XPATH, '//*[@id="conteudo-principal"]/form/input[6]').click()

#Seleciona contador
caixa_selecao = driver.find_element(By.XPATH, '//*[@id="TipoAcesso_Regra"]').click()
bot.press('down')
bot.press('enter')

#Clica no botão enter
entrar = driver.find_element(By.XPATH, '//*[@id="conteudo-principal"]/form/div[2]/div/button').click()

time.sleep(1.5)

#Seleciona ISSQN no topo da página
nfse_element = driver.find_element(By.CLASS_NAME, "sistemaAtual")
ActionChains(driver).move_to_element(nfse_element).perform()

laranjinha = driver.find_element(By.CLASS_NAME, "col-md-2")
laranjinha.click()

# boneco = driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div/div[2]/div[1]/a[1]').click()
boneco = wdw.until(visibilidade((By.XPATH, '//*[@id="wrap"]/div[2]/div/div[2]/div[1]/a[1]'))).click()


#Muda numero de empresas na tela de 10 para 100
caixa_selecao = wdw.until(visibilidade((By.XPATH, '//*[@id="DataTable_length"]/label/select')))
selecao = Select(caixa_selecao)
selecao.select_by_value("100")


#Seleciona ordem de nome
razao_social = driver.find_element(By.XPATH, '//*[@id="DataTable"]/thead/tr/th[5]')
razao_social.click()
razao_social.click()


#Abre o arquivo excel com os cnpjts do simpliss
caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\bots_Contabilidade\CNPJS_SIMPLISS.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('fazendo')


for row in range(sheet.nrows):
    cnpj = sheet.cell_value(row, 0)
    xpath = f'//img[@title="Selecionar" and contains(@onclick, "{cnpj}")]'
    button = wdw.until(visibilidade((By.XPATH, xpath))).click()
    
    handles = driver.window_handles
    driver.switch_to.window(handles[-1])

    link_url = f"https://nfse.blumenau.sc.gov.br/contrib/app/issqn/relatorio_livros?ck={cnpj}"
    driver.execute_script("window.open(arguments[0], '_blank');", link_url)

    bot.sleep(6)

    div_container = driver.find_element(By.XPATH ,'//*[@id="ctl00_ContentPlaceHolder1_panel_Cadastro"]/div[12]/div[2]')
    data1 = div_container.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txt_Data_Inicio"]').send_keys("01022024")
    
    
    
    # bot.sleep(5)

    # bot.click(809, 587)
    # bot.sleep(0.1)
    # bot.click(846, 624)
    # bot.sleep(2)
    # bot.press('TAB')
    # bot.press('TAB')
    # bot.sleep(2)
    # bot.write("01022024")
    # bot.sleep(2)
    # bot.press('TAB')
    # bot.sleep(2)
    # bot.write("29022024")
    # bot.sleep(2)
    # bot.click(719, 708)
    # bot.click(908, 946)
    
        
    bot.sleep(900)   

    
