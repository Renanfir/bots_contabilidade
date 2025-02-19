import pyautogui as bot
altura = 74
bot.click(1803,15) #Sai da tela
bot.sleep(1.2)
for i in range(16):
    
    
    bot.click(1107, altura)
    bot.sleep(0.15)
    bot.click(852, 363)
    bot.sleep(0.2)
    bot.press("tab")
    bot.press("tab")
    bot.write("99")
    bot.press("tab")
    bot.write("9")
    bot.click(1692,946)
    bot.click(1789,990)
    bot.sleep(0.3)
    altura += 16
