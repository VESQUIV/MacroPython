        
import time

while True:   
    try:

        print('0.- EJECUTANDO ATEXIT')
        #-------------------------------------------SALIDA TXT----------------------------------------------------------
        import atexit
        import sys
        

        def guardar_salida():
            # Guardar la salida estándar y la salida de error en archivos
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            

            # Cerrar los archivos
            salida_archivo.close()
            
        # Registrar la función para ser llamada al finalizar el script
        atexit.register(guardar_salida)

        # Guardar la salida estándar y la salida de error originales
        original_stdout = sys.stdout
        original_stderr = sys.stderr

        # Abrir archivos para guardar la salida estándar y la salida de error
        salida_archivo = open('Salida.csv', 'w')


        # Redirigir la salida estándar y la salida de error a los archivos
        sys.stdout = salida_archivo
        sys.stderr = salida_archivo
        print('TXT iniciado con exito')

       # Redirigir la salida estándar y la salida de error a los archivos
        sys.stdout = salida_archivo
        sys.stderr = salida_archivo
        print('TXT iniciado con exito')






        #-----------------------------------SMART_REPORT--------------------------------------------------------------------------------------------------------

        print('1.- EJECUTANDO DESCARGA DE REPORTE SMART')
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from datetime import datetime, timedelta
        from dotenv import load_dotenv 
        from pathlib import Path
        import pyautogui
        import os
        import time
        import win32com.client
        print('Librerías Cargadas Con Exito')

        FILE_ORIGEN1 = Path('C:/Users/vesquiv3/Downloads/rptCallResposenDurationDetails.xlsx')
        DOWNLOADS = Path('C:/Users/vesquiv3/Downloads')

        # Verificar si la carpeta existe
        if os.path.exists(DOWNLOADS):
            # Recorrer todos los archivos en la carpeta
            for archivo in os.listdir(DOWNLOADS):
                # Obtener la ruta completa del archivo
                ruta_archivo = os.path.join(DOWNLOADS, FILE_ORIGEN1)
                # Verificar si es un archivo
                if os.path.isfile(FILE_ORIGEN1):
                    # Borrar el archivo
                    os.remove(FILE_ORIGEN1)
            print("Archivos eliminados con éxito.")
        else:
            print("La carpeta no existe.")
        
        print('Cargando Chrome')

        #CARGAR CHROME
        load_dotenv()
        driver = webdriver.Chrome()
        driver.maximize_window()
        print('Drivers Cargados Con Exito') 

        # Abre una página web en el navegador
        print('Ingresando a SMART')
        URL_SMART = os.getenv('URL_SMART')
        driver.get(URL_SMART)
        #time.sleep(3)
        print("URL SMART Abierto")

        print('Ingresando Credenciales')
        USERNAME_SMART = os.getenv('USERNAME_SMART')
        PASSWORD_SMART = os.getenv('PASSWORD_SMART')
        

        #bloque agregado
        #comprueba que el componente existe antes de ingresar credenciales
        
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID,'username')))
        finally:
            print('Elemento cargado')
        #Fin del bloque agregado


        Ingresar_username = driver.find_element(By.ID,'username')
        Ingresar_username.send_keys(USERNAME_SMART)
        #time.sleep(1)
        ingresar_password = driver.find_element(By.ID,'password')
        ingresar_password.send_keys(PASSWORD_SMART)
        #time.sleep(1)

        ingresar_password.send_keys(Keys.ENTER)
        # time.sleep(6)
        print("Se han ingresado las credenciales")

        # Abre una página web en el navegador
        print('Ingresando a Change Domain')
        URL_CD = os.getenv('URL_CD')
        driver.get(URL_CD)
        time.sleep(2)
        print("URL CD Abierto")

        print('Interacción de Teclado')
        boton_dominio = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/change-domain/div/div[2]/domain-search/div/sm-ui-search/div/sm-ui-button[1]/button/span')
        boton_dominio.click()
        time.sleep(1)

        pyautogui.hotkey('down')
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)

        boton_update = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/change-domain/div/div[4]/sm-ui-button[1]/button')
        boton_update.click()
        time.sleep(1)

        # Abre una página web en el navegador
        print('Ingresando a CRD')
        URL_CRD = os.getenv('URL_CRD')
        driver.get(URL_CRD)
        print('URL CRD Abierto')
        time.sleep(3)

        boton_search = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/smartreport/div[2]/fieldset/div[3]/p-radiobutton/div/div[2]/span')
        boton_search.click()
        # time.sleep(1)

        boton_d2 = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/smartreport/div[3]/fieldset/div/div[1]/domain-search/div/sm-ui-search/div/sm-ui-button[1]/button/span')
        boton_d2.click()
        # time.sleep(1)
        print("Dominio Ingresado")

        print('Interacción de teclado')
        pyautogui.hotkey('down')
        pyautogui.hotkey('down')
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)

        #fecha_ayer = (datetime.now() - timedelta(days=1)).strftime('%m/%d/%Y')

        #eliminar_fecha = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/smartreport/div[4]/fieldset/div[1]/div[1]/p-calendar[1]/span/input')
        #eliminar_fecha.send_keys(Keys.CONTROL,'a')
        #eliminar_fecha.send_keys(Keys.DELETE)
        #time.sleep(2)
        #print("Fecha Eliminada")

        #ingresar_fecha = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/smartreport/div[4]/fieldset/div[1]/div[1]/p-calendar[1]/span/input')
        #ingresar_fecha.send_keys(fecha_ayer)
        #time.sleep(2)
        #print("Fecha ingresada")

        #eliminar_hora = driver.find_element(By.XPATH,'//*[@id="time"]/span/input')
        #eliminar_hora.click()
        #eliminar_hora.send_keys(Keys.CONTROL,'a')
        #eliminar_hora.send_keys(Keys.DELETE)
        #time.sleep(2)
        #print("Hora Eliminada")

        #HORA_SMART = os.getenv('HORA_SMART')
        #pyautogui.typewrite(HORA_SMART)
        #time.sleep(2)
        #print("Hora Ingresada")

        click_page = driver.find_element(By.ID,'pageWrapper')
        click_page.click()
        # time.sleep(1)

        boton_FromStatus = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/smartreport/div[4]/fieldset/div[2]/div[1]/sm-ui-dropdown/p-dropdown/div/div[3]/span')
        boton_FromStatus.click()
        time.sleep(1)
        teclas = ActionChains(driver)
        teclas.send_keys("r").perform()
        time.sleep(1)
        teclas.send_keys(Keys.ENTER)
        time.sleep(1)

        boton_ToStatus = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/smartreport/div[4]/fieldset/div[2]/div[3]/sm-ui-dropdown/p-dropdown/div/div[3]')
        boton_ToStatus.click()
        time.sleep(1)
        teclas = ActionChains(driver)
        teclas.send_keys("m").perform()
        time.sleep(1)
        click_page = driver.find_element(By.ID,'pageWrapper')
        click_page.click()
        time.sleep(1)

        boton_GenerateReport = driver.find_element(By.XPATH,'//*[@id="pageWrapper"]/div/div/sm-ui-route/div[12]/smartreport/div[5]/sm-ui-button/button')
        boton_GenerateReport.click()
        time.sleep(5)
        print("Reporte Generado y Guardado con Exito")

        # Cerrar el navegador
        driver.quit()

        #--------------------------------------------COPY_SMART-----------------------------------------------------------------------------------------
        # print('2.- EJECUTANDO COPY SMART')
        # from pathlib import Path
        # import shutil
        # import time
        

        # FILE_ORIGEN = Path('C:/Users/vesquiv3/Downloads/rptCallResposenDurationDetails.xlsx')
        # FILE_DESTINO = Path('C:/Users/vesquiv3/Desktop/SMART_Report.xlsx')

        # shutil.copy2(FILE_ORIGEN,FILE_DESTINO)
        # # time.sleep(3)
        # print('Se ha copiado archivo a DESKTOP')

        print('2.- EJECUTANDO COPY SMART')
        from pathlib import Path
        import shutil
        import time
        from datetime import datetime
        import os



        FILE_ORIGEN = Path('C:/Users/vesquiv3/Downloads/rptCallResposenDurationDetails.xlsx')
        FILE_DESTINO = Path('C:/Users/vesquiv3/Desktop/SMART_Report.xlsx')


        FILE_RESPALDO = Path('C:/Users/vesquiv3/azureford/MPL Power BI Dashboard - Documents/MacroPython/Respaldo')
        FILE_MACROPY = Path('C:/Users/vesquiv3/azureford/MPL Power BI Dashboard - Documents/MacroPython')


        FORMATO = ".xlsx"


        # FECHA_HOY = datetime.now().strftime('%m-%d-%Y-%H.%M')
        # NOMBRE_HOY = os.path.join(FECHA_HOY + FORMATO)
        # ARCHIVO_NUEVO = os.path.join(FILE_RESPALDO , NOMBRE_HOY)


        shutil.copy2(FILE_ORIGEN,FILE_DESTINO)
        time.sleep(2)
        print('Se ha copiado archivo a DESKTOP')


        shutil.copy2(FILE_DESTINO, FILE_MACROPY)
        time.sleep(1)
        print('Se ha copiado archivo a Carpeta Ford')


        # shutil.copy2(FILE_DESTINO, ARCHIVO_NUEVO )
        # time.sleep(1)
        # print('Se ha creado Respaldo')

        #--------------------------------------------MACRO-----------------------------------------------------------------------------------------------
        print('3.- EJECUTANDO MACRO')
        import win32com.client
        from pathlib import Path
        import openpyxl 

        excel = win32com.client.Dispatch('Excel.Application')

        libro = excel.Workbooks.Open(Path(r'C:/Users/vesquiv3/Desktop/WMS_PS_Actualizado2_12_08.xlsm'))

        excel.Application.Run('Macro2')

        libro.Save()
        libro.Close()

        excel.Quit()
        print('MACRO EJECUTADA')
        # time.sleep(2)
        #gc.collect()
        #print('MEMORIA RAM LIBERADA')

        
        # CSV_FILE = 'C:/Users/VESQUIV3/Desktop/filetmp1.csv'
        # LIBRO_CSV = openpyxl.load_workbook(CSV_FILE)

        # hoja_trabajo = LIBRO_CSV.active

        # valor_celda_A2 = hoja_trabajo['A2'].value 

        # if valor_celda_A2 == "":
        #     print('TEMPLATE VACIO')
        #     continue      
        #------------------------------------------------SAP-----------------------------------------------------------------------------------------------

        print('3.- EJECUTANDO SAP')
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from datetime import datetime, timedelta
        from dotenv import load_dotenv
        import os
        import time

        load_dotenv()
        driver = webdriver.Chrome()
        driver.maximize_window()

        # Abre una página web en el navegador PRODUCTIVO
        URL_SAP = os.getenv('URL_SAP')
        driver.get(URL_SAP)
        # time.sleep(5)
        print("Open SAP")

        # Abre una página web en el navegador PRUEBA
        # PRUEBAURL_SAP = os.getenv('PRUEBAURL_SAP')
        # driver.get(PRUEBAURL_SAP)
        # time.sleep(5)
        # print("Open SAP")

        #Credenciales Productivo
        USERNAME_SAP = os.getenv('USERNAME_SAP')
        PASSWORD_SAP = os.getenv('PASSWORD_SAP')
        SAP_username = driver.find_element(By.ID,'USERNAME_FIELD-inner')
        SAP_username.send_keys(USERNAME_SAP)
        # time.sleep(1)
        SAP_password = driver.find_element(By.ID,'PASSWORD_FIELD-inner')
        SAP_password.send_keys(PASSWORD_SAP)
        # time.sleep(1)

        #Credenciales PRUEBA
        # PRUEBAUSERNAME_SAP = os.getenv('PRUEBAUSERNAME_SAP')
        # PRUEBAPASSWORD_SAP = os.getenv('PRUEBAPASSWORD_SAP')
        # SAP_username = driver.find_element(By.ID,'USERNAME_FIELD-inner')
        # SAP_username.send_keys(PRUEBAUSERNAME_SAP)
        # time.sleep(1) 
        # SAP_password = driver.find_element(By.ID,'PASSWORD_FIELD-inner')
        # SAP_password.send_keys(PRUEBAPASSWORD_SAP)
        # time.sleep(1)

        SAP_password.send_keys(Keys.ENTER)
        time.sleep(5)
        print("Se han ingreado las credenciales")

        # Abre una página web en el navegador PRODUCTIVO
        URL_ODC = os.getenv('URL_ODC')
        time.sleep(5)
        driver.get(URL_ODC)
        time.sleep(5)
        print("Open Delivery Creation")

        # Abre una página web en el navegador PRUEBA
        # PRUEBAURL_ODC = os.getenv('PRUEBAURL_ODC')
        # driver.get(PRUEBAURL_ODC)
        # time.sleep(4)
        # print("Open Delivery Creation")

        import pyautogui
        #Warehouse
        time.sleep(4)
        print('Interacción de teclado')
        pyautogui.hotkey('shift','f5')
        time.sleep(5)

        for i in range(5):
            pyautogui.hotkey('down')
        time.sleep(5)

        pyautogui.hotkey('shift','f10')
        time.sleep(4)

        for i in range(5):
            pyautogui.hotkey('down')
        time.sleep(4)
        #SET VALUE    
        pyautogui.hotkey('enter')
        time.sleep(3)

        for i in range(11):
            pyautogui.hotkey('tab')
        time.sleep(6)
        #FILE NAME
        pyautogui.press('f4')
        time.sleep(4)

        for i in range(4):
            pyautogui.hotkey('tab')
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(3)

        for i in range(6):
            pyautogui.hotkey('tab')    
        time.sleep(3)
        pyautogui.typewrite('filetmp1.csv')
        time.sleep(3)
        pyautogui.press('enter')        
        time.sleep(3)
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('up')
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(4)
        pyautogui.press('f8')
        time.sleep(6)
        pyautogui.press('f8')
        #time.sleep(4)
        #pyautogui.hotkey('tab')
        #pyautogui.hotkey('tab')
        #pyautogui.sleep(2)
        #pyautogui.press('enter')  
        # pyautogui.hotkey('ctrl', 'shift', 'j')
        # pyautogui.sleep(2)
        # pyautogui.typewrite('document.querySelector("#msgarea [data-toolbaritem-id="M0:50::btn[8]-r"]").click();')
        # pyautogui.sleep(1)
        # pyautogui.press('enter')

        time.sleep(15)
        print('REPORTE ENVIADO')

        driver.quit()
        print('Chrome Cerrado')

        
        

    #---------------------------------------EXCEPT------------------------------------------------------------------------------

    except:    
        print('EJECUNTANDO MACRO ERROR DE ENVIO')
    #     excel = win32com.client.Dispatch('Excel.Application')

    #     libro = excel.Workbooks.Open(Path(r'C:/Users/vesquiv3/Desktop/MacroCorreos.xlsm'))

    #     excel.Application.Run('CorreoFileNotSent')

    #     libro.Save()
    #     libro.Close()

    #     excel.Quit()

    #     print('Reporte NO ENVIADO')

    # time.sleep(10)
