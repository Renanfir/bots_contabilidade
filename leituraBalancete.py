import pdfplumber
import pyautogui as bot

bot.click(1806, 16)
bot.sleep(1.5)

path = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\TESTE BALANCETE - BALANCO AUTOMATICO\CATARINENSE.pdf'

listaContas = [
    "6303", "6305", "1179", "3212", "6306", "6304", "2113", "6511", "6497", "6474", "6468", 
    "6563", "6451", "6481", "6570", "6942", "7485", "1542", "6513", "6586", "2983", "6882", 
    "6706", "1234", "1235", "50055", "7008", "7114", "1000", "65123","58271","1101"
    "886", "6831", "6669", "66111", "57997", "7048", "7108"
]


with pdfplumber.open(path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            for line in text.split("\n"):
                if not line[0] == "*" and line[0].isdigit():
                    if line.split()[0] in listaContas:
                    # if line[-1] == "D":
                        codigo = line.split()[0]
                        valor = line.split()[-1].rstrip("D").replace(".","")
                        bot.write(codigo)
                        bot.press('enter')
                        bot.sleep(0.15)
                        bot.write(valor)
                        bot.press('enter')
                        bot.sleep(0.15)
                        bot.press('enter')
                        bot.sleep(0.15)
                        bot.press('enter')
                        bot.sleep(0.15)
                        bot.press('enter')
                        bot.sleep(0.15)
                        bot.press('enter')
                        bot.sleep(0.15)
                        bot.press('enter')
                




