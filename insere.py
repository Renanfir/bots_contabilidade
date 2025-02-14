import pyautogui as bot
import xlrd

bot.click(1806, 18)
bot.sleep(1)

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\SIMPLISS WEB.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')            
for row in range(sheet.nrows):
    cnpj = sheet.cell_value(row, 0)
    bot.click(950, 749)
    bot.sleep(0.65)
    bot.click(634, 414)
    bot.write(cnpj)
    bot.sleep(0.65)
    bot.click(867, 490)
    bot.sleep(0.65)
    bot.click(954, 745)
    bot.sleep(0.65)