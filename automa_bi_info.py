from selenium import webdriver  # Importa el módulo webdriver para interactuar con navegadores web
from selenium.webdriver.common.by import By  # Proporciona métodos para localizar elementos en una página
from selenium.webdriver.common.keys import Keys  # Permite el uso de teclas del teclado como Enter, Esc, etc.

from selenium.webdriver.support import expected_conditions as EC  # Contiene condiciones esperadas para usar con WebDriverWait
from selenium.webdriver.support.ui import WebDriverWait  # Permite esperar la presencia de ciertos elementos o condiciones

from selenium.webdriver.chrome.service import Service as ChromeService  # Servicio para gestionar la instancia de Chrome
from webdriver_manager.chrome import ChromeDriverManager  # Administra la descarga/actualización automática del driver de Chrome
from selenium.webdriver.chrome.options import Options  # Permite configurar opciones para el navegador Chrome

from selenium.webdriver.common.action_chains import ActionChains

import time  # Proporciona funciones relacionadas con el tiempo, como sleep
import os
import csv

from dotenv import load_dotenv

def configurar_navegador(ruta_descarga):
    chrome_options = Options()
    prefs = {"download.default_directory": ruta_descarga}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

    return driver

def iniciar_sesion(driver,user,passw):
    driver.get("https://app.powerbi.com/groups/d696634a-57af-413c-bea8-4f6b3e9f2933/dashboards/b5ece1fe-bdbd-4de6-ac4d-54eeea1039bd?experience=power-bi") #poner la url justo donde esta el panel
    driver.maximize_window()
    time.sleep(6)

    if "Power" in str(driver.title):
        # Ingreso de nombre de usuario
        userName=esperar_y_encontrar_elemento(driver,locator=(By.XPATH, '//*[@id="email"]'))
        userName.click()
        userName.send_keys(user)
        time.sleep(10)

        # Click en botón siguiente
        button=esperar_y_encontrar_elemento(driver,locator=(By.XPATH, '//*[@id="submitBtn"]'))
        button.click()
        time.sleep(4)

        # Ingreso de contraseña
        password=esperar_y_encontrar_elemento(driver,locator=(By.XPATH,'//*[@id="i0118"]' ))
        password.click()
        password.send_keys(passw)

        # Click en botón ingresar
        button=esperar_y_encontrar_elemento(driver,locator=(By.XPATH,'//*[@id="idSIButton9"]' ))
        button.click()

        time.sleep(4)
        
        # Click en botón ingresar
        button=esperar_y_encontrar_elemento(driver,locator=(By.XPATH,'//*[@id="idBtn_Back"]' ))
        button.click()

        time.sleep(10)


    else:
        # Ingreso de nombre de usuario
        userName=esperar_y_encontrar_elemento(driver,locator=(By.XPATH, '//*[@id="i0116"]'))
        userName.click()
        userName.send_keys(user)


        # Click en botón siguiente
        button=esperar_y_encontrar_elemento(driver,locator=(By.XPATH, '//*[@id="idSIButton9"]'))
        button.click()
        time.sleep(4)

        # Ingreso de contraseña
        password=esperar_y_encontrar_elemento(driver,locator=(By.XPATH,'//*[@id="i0118"]' ))
        password.click()
        password.send_keys(passw)

        # Click en botón ingresar
        button=esperar_y_encontrar_elemento(driver,locator=(By.XPATH,'//*[@id="idSIButton9"]' ))
        button.click()

        time.sleep(4)
        
        # Click en botón ingresar
        button=esperar_y_encontrar_elemento(driver,locator=(By.XPATH,'//*[@id="idBtn_Back"]' ))
        button.click()

        time.sleep(10)


#si no le esta exportando,muchas veces es porque se hizo algún cambio en la tabla en el power bi desktop y se volvió a publicar al pro eso puede generar error
#lo que debes de hacer es ir al panel en el pro y clik derecho fuera de la tabla y selecionar inspeccionar ya ahí buscas como se llama el botón de los tres punticos que esta arriba a la derecha de la tabla 
# y tambien el botón de  la opción exportar al encontrarlos copias los nombres de los botones en "copy full xpath" y el nombre de los tres punticos es el path
#que vas a pegar en button_options y el botón de exportar lo pegas en el button_export_csv que esta en la función de abajo (paso_descargar). 






def paso_descargar(driver):
    button_options=esperar_y_encontrar_elemento(driver,locator=(By.XPATH, '/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/dashboard-container-new/div/div/div[2]/div/dashboard-new/dashboard-base/div/pbi-grid/ul/li[2]/tile-new/tile-content/section/div/div[1]/tileoptions-context-menu/button'))
    actions = ActionChains(driver)
    actions.move_to_element(button_options).click().perform()

    time.sleep(5)

    button_export_csv = driver.find_element(By.XPATH,'/html/body/div[2]/div[3]/div/div/div/div/div[6]/button')
    actions.move_to_element(button_export_csv).click().perform()

  
    time.sleep(10)

 


def esperar_y_encontrar_elemento(driver, locator, timeout=20):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))
    return driver.find_element(*locator)

def main():
    driver=configurar_navegador(r"C:\Users\prac.planeacionfi\OneDrive - Prebel S.A BIC\Escritorio\Practi-SOFÍA\bot_powerBI_CMI")

    iniciar_sesion(driver,"planeacion.financierabi@prebel.com.co","Prebelvidas123.-")
    
    paso_descargar(driver)


                
main()


# DESCARGAR ARCHIVO DE EXCEL CON REFERENCIAS DIARIAS DESDE POWER BI PRO
# from selenium.webdriver.common.action_chains import ActionChains

# edge_options = Options()
# edge_options.use_chromium = True  # Utilizar el modo Chromium

# driver_path = r'C:\Users\kevin.primera\Downloads\edgedriver_win64\msedgedriver.exe'
# service = Service(driver_path)

# driver = webdriver.Edge(service=service, options=edge_options)
# driver.get('https://app.powerbi.com/groups/me/dashboards/c6ba9019-3062-409e-ac35-5f25eb009861?experience=power-bi')

# driver.maximize_window()

# time.sleep(60)

# # Hacer clic en el botón de opciones
# button_options = driver.find_element(By.XPATH,'/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-extension-page-outlet/div[2]/dashboard-container-new/div/div/div[2]/div/dashboard-new/dashboard-base/div/pbi-grid/ul/li[2]/tile-new/tile-content/section/div/div[1]/tileoptions-context-menu/button/mat-icon')
# actions = ActionChains(driver)
# actions.move_to_element(button_options).click().perform()

# time.sleep(10)

# # Hacer clic en el botón de exportar como CSV
# button_export_csv = driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/div/div/div/div/div[6]/button')
# actions.move_to_element(button_export_csv).click().perform()

# time.sleep(10