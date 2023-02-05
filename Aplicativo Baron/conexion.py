from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time


website = 'https://1826.segurlike.es/SegElevia1826/LoginManager.mvc/LogOn?ReturnUrl=%2fsegelevia1826%2fPageRender.mvc%2fseg.Agenda.AgendaDataPanel'
#website = 'https://signin.intra.42.fr/users/sign_in'
#website = 'https://www.google.com/'
path = "//Users/paula/Downloads/chromedriver"
s = Service(path)
driver = webdriver.Chrome(service = s)
driver.get(website)
driver.maximize_window()

# defino las credenciales de entrada (esto se incluira en la BBDD)
User = "PAULA_R"
Password = "Likeahouseonfire22!"

driver.find_element(By.XPATH, '//input[@name="username"]').send_keys(User)
driver.find_element(By.XPATH, '//input[@name="password"]').send_keys(Password)
driver.find_element(By.XPATH, '//button[@type="submit"]').click()
time.sleep(10)
result = driver.find_element(By.XPATH, '//li[@data-ec-menu = "Informes"]').click()
time.sleep(50)
#print(result)
driver.quit()
#website = 'https://www.google.com/'
#driver.get(website)
#driver.maximize_window()
#driver.find_element(By.XPATH, '//input[@name = "q"]').send_keys("fotos de gatos")
#time.sleep(10)
#driver.find_element(By.XPATH, '//input[@value = "Buscar con Google"]').click
#time.sleep(10)