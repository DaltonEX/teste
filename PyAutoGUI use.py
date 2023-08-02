from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
import time

browser = webdriver.Edge()

browser.get('https://deere.sharepoint.com/sites/itnw/Lists/Portaria_DataBase/AllItems.aspx')

browser.find_element(By.XPATH, '//*[@id="appRoot"]/div[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[1]/button/span').click()
time.sleep(1)
pyautogui.typewrite('538901\tJoao Pedro\tFreitas')

pyautogui.press('tab',2)
time.sleep(1)
pyautogui.press('enter')

browser.close()