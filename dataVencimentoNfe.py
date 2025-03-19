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
url = f"https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g="
driver.get(url)
# captcha = wdw.until(visibilidade((By.XPATH, '//*[@id="checkbox"]'))).is_displayed()
# while(captcha):
#     bot.sleep(7)
#     captcha = wdw.until(visibilidade((By.XPATH, '//*[@id="checkbox"]'))).is_displayed()
# bot.click(713, 535)
bot.sleep(9999)