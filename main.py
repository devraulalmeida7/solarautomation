from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

chrome_options = webdriver.ChromeOptions()
chrome_options.debugger_address = "127.0.0.1:9222"

driver = webdriver.Chrome(options=chrome_options)






# button = driver.find_element(By.XPATH, "//button[contains(text(), 'Optimizer 1.1.6 (13E59216)')]")
button = driver.find_element(By.CLASS_NAME, "kqYmEA")
# button2 = driver.find_element(By.XPATH, "(//div[@class='kqYmEA'])[3]")
button.click()


# espera = WebDriverWait(driver, 10)  # Tempo máximo de espera (10s)
# botao = espera.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Optimizer 1.1.6 (13E59216)')]")))
# botao.click()


# print("Informação extraida:", elemento)

driver.quit()