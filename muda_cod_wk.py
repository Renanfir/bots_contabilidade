import pyautogui as bot
import xlrd

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\bots_Contabilidade\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('LISTA SISTEMAS WK')

bot.click(1802, 15)
for row in range(sheet.nrows):
    nome = sheet.cell_value(row, 0)
    bot.click(17, 55)
    bot.sleep(0.3)
    bot.write(nome)
    bot.press('enter')
    bot.sleep(0.2)
    bot.write('aureo')
    bot.press('TAB')
    bot.write('aureore')
    bot.press('enter')
    bot.sleep(1)
    bot.click(140, 32)
    bot.sleep(0.1)
    bot.click(186, 98)
    bot.sleep(4.5)
    bot.click(617, 95)
    bot.click(347, 219)
    bot.doubleClick(158, 392)
    bot.write('SC/2024/323416')
    bot.press('TAB')
    bot.write('200624')
    bot.click(1795, 999)    
    bot.hotkey('alt','f4')
    bot.sleep(0.2)
    bot.click(40, 55)
