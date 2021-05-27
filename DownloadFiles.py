from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as exc_condts
from selenium.webdriver.common.by import By
#from webdriver_manager.chrome import ChromeDriverManager
import time, os, csv, math
from zipfile import  ZipFile as Zip
from openpyxl import load_workbook 

class reporte():

    def __init__(self):
        #Establecemos las siguietnes opciones para evitar la ventana de chrome que indica que nuestra conexión no es privada o no es segura
        self.options = webdriver.ChromeOptions()
        self.options.add_argument('--ignore-ssl-errors=yes')#ignorar errores de ssl
        self.options.add_argument('--ignore-certificate-errors')#ignorar errores de certificado
        #self.options.headless = True#Para que se ejecute en background

        #Luego instanciamos nuestro webdriver y le pasamos las opciones para evitar la ventana que indica que la conexión no es segura
        self.driver = webdriver.Chrome(executable_path='C:/driver_Chrome/chromedriver.exe',options=self.options)
        #self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=self.options)

    def login_I2KV5(self):

        self.url_cpu = 'https://www.xxxxxxx.com'
        
        self.driver.minimize_window()#Para que inicie minimizado el explorador
        self.driver.get(self.url_cpu)

        #Esperamos hasta que el boton de login este localizable en la pagina
        WebDriverWait(self.driver, 10).until(exc_condts.presence_of_element_located((By.ID, 'btn_submit')))

        #Luego ingresamos el user y el password
        self.login = self.driver.find_element_by_id('txf_username')
        self.login.send_keys('xxxxx')

        self.password = self.driver.find_element_by_id('txf_imtinfo')
        self.password.send_keys('xxxxx')

        self.btn_login = self.driver.find_element_by_id('btn_submit')
        self.btn_login.click()
        
        #Esperamos hasta que el elemento indicado este localizable y con esto saber que la pagina ha cargado con exito para continuar
        #el elemento por ID es la opcion que dice settings de la pagina
        WebDriverWait(self.driver, 10).until(exc_condts.presence_of_element_located((By.ID, 'menu.com.iemp.system')))


    def descarga_reporte(self):

        self.lista_urls = [
            'https://xxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxxxx&from=topnMonitorView',
            'https://xxxxxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxxx&from=topnMonitorView',
            'https://xxxxxxxxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxx&from=topnMonitorView',
            'https://xxxxxxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxx&from=topnMonitorView',
            'https://xxxxxxxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxx&from=topnMonitorView',
            'https://xxxxxxxxxxxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxx&from=topnMonitorView',
            'https://xxxxxxxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxxxxx&from=topnMonitorView',
            'https://xxxxxxxxxxxxxxx/pm/themes/default/pm/app/history_pm_alone.html?curMenuId=com.iemp.app.pm.historypm&monitorPortletID=test&mapKey=xxxxx&from=topnMonitorView',
            
        ]

        for url in self.lista_urls:
            self.driver.get(url)
            #Esperamos a que este visible el elemento donde esta la opcion del historico por semana y damos click
            WebDriverWait(self.driver, 10).until(exc_condts.presence_of_element_located((By.ID, 'timerange_6'))).click()
            
            #Esperamos un tiempo para que este habilitado el boton de export que nos permite descargar el archivo
            time.sleep(1)
            
            #Ubicamos el boton de export y damos click para descargar el archivo
            self.btn_export = self.driver.find_element_by_id('btnExportData')
            self.btn_export.click()


    def logout(self):

        self.logout = self.driver.find_element_by_id('login_logoutIcon')
        self.logout.click()
        time.sleep(0.5)

        self.boton_ok_salir = self.driver.find_element_by_id('fw_btn_ok')
        self.boton_ok_salir.click()
        self.driver.close()

    #Las siguientes lineas de codigo corresponden a descomprimir los archivos y colocarlos en el directorio 
    # que nos interesa para procesarlos y preparar el reporte para enviarlor
    def descomprime_data(self):

        self.ruta_archivo = 'C:/Users/FRANK BOJORQUEZ/Downloads/'
        self.ruta_destino = 'D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/Archivos_vpn/'

        self.lista_archivos_zip = os.listdir(self.ruta_archivo)

        self.archivos_zip = [self.archivo_zip for self.archivo_zip in self.lista_archivos_zip if '.zip' in self.archivo_zip]

        self.n = 0
        for self.archivo_zip in self.archivos_zip:   
            self.archivo = Zip(self.ruta_archivo+self.archivo_zip)
            self.archivo.extractall(self.ruta_destino)
            self.archivo.close()
            self.archivo_csv = os.listdir(self.ruta_destino)[self.n]
            os.rename(self.ruta_destino+self.archivo_csv,self.ruta_destino+f'Archivo_{self.n}_'+self.archivo_csv)#Renombramos el archivo en la carepta destino
            os.remove(self.ruta_archivo+self.archivo_zip)
            self.n +=1

#---- creamos nuestro objeto de tipo reporte y llamamos a los metodos de la clase reporte--------------------------

generador_reporte = reporte() #Creacion de nuestro objeto de tipo reporte

generador_reporte.login_I2KV5()#-----llamada a la funcion que hace el login en la pagina de Iportal
generador_reporte.descarga_reporte()#---Llamada a la funcion que descaraga los archivos que contienen los datos de cpu
generador_reporte.logout()#----------llamada a la funcion que hace el logout del Iportal
generador_reporte.descomprime_data()#Llamada a la fucnion que descomprime los archivos y los coloca en la ruta: D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/Archivos_vpn/
#--------------------------------------------------------------------------------------------------------------------

#_______________________________________________________________________________________________________________________

#Las siguientes funciones se utilizan para obtener el promedio del CPU y agregar el valor en 
#el archivo a enviar
#_______________________________________________________________________________________________________________________


# ----------------------- Funcion que obtiene el promedio del CPU de cada Tarjeta -----------------------------------------------
def cpu(archivo):
    ruta_archivos_excell = 'D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/Archivos_vpn/'

    valores_cpu = []

    with open(ruta_archivos_excell+archivo) as f:
        lector = csv.reader(f)
        for fila in lector:
            if fila[1].isnumeric():
                valores_cpu.append(int(fila[1]))

    promedio_tarjeta = sum(valores_cpu) / len(valores_cpu) 
    return promedio_tarjeta

#-------------------------------------------------------------------------------------------------------------------

def elimina_archivos_csv():
    
    ruta_archivos_csv = 'D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/Archivos_vpn/'
    archivos_csv = os.listdir(ruta_archivos_csv)

    for archivo in archivos_csv:
        os.remove(ruta_archivos_csv+archivo)

#--------------funcion que Modifica el nombre del archivo a enviar con la fecha actual------------------------------------------
def cambiar_nombre_archivo():
    
    fecha = time.strftime("%d-%m-%y")
    ruta_archivo_final = 'D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/ArchivoFinal_a_enviar/'

    nombre_original = os.listdir(ruta_archivo_final)[0]
    nombre_modificado = f'Plantilla_ Disponibilidad CORE_IT-{fecha}.xlsx'
    
    #renombramos el archivo con la fecha actual
    os.rename(ruta_archivo_final+nombre_original,ruta_archivo_final+nombre_modificado)

    return nombre_modificado, ruta_archivo_final
#--------------------------------------------------------------------------------------------------------------------

#-------------- funcion que Inserta el valor del CPU en el reporte a enviar -------------------------------------------------
def inserta_dato_cpu(nombre_archivo, path, promedio_cpu_vpn):

    reporte = load_workbook(path+nombre_archivo)
    hoja_plataformas = reporte.get_sheet_by_name('PLATAFORMAS')
    hoja_plataformas['D5'] = promedio_cpu_vpn
    reporte.save(path+nombre_archivo)
#---------------------------------------------------------------------------------------------------------------------

#----- Llamamos a la fucion cpu la cual nos devuelve el promedio por tarjeta y el resultado lo vamos sumando -------------------------------
lista_archivos_csv = os.listdir('D:/A_PYTHON/ProgramasPython/Control_NodosCA/Reporte_CPU_VPN-PPS/Archivos_vpn/')
suma_valores_cpu = 0
for archivo in lista_archivos_csv:
    suma_valores_cpu += cpu(archivo)

#redondeamos el valor y lo dividimos en la cantidad de tarjetas
promedio_cpu_vpn = str(math.ceil(suma_valores_cpu/8))+'%' 

#---------------------------------------------------------------------------------------------------------------------

#Llamamos a la funcion que borra los archivos CSV los cuales contienen la informacion de CP de cada tarjeta

elimina_archivos_csv()

#------------------------------------------------------------------------------------------------------------------------

#Llamamos a la funcion cambiar_nombre_archivo la cual nos devuelve el nuevo nombre del archivo a 
#enviar por correo y la ruta del archivo

nuevo_nombre, path_file = cambiar_nombre_archivo()

#--------------------------------------------------------------------------------------------------------------------

#------Llamamos a la funcion que inserta el valor de CPU en el archivo a enviar por correo------------------------------------------------------------------------

inserta_dato_cpu(nuevo_nombre, path_file, promedio_cpu_vpn)

#--------------------------------------------------------------------------------------------------------------------
#El acrhivo que contiene el subject del correo con la fecha correspondiente

def modificar_subject_txt():

    FECHA = time.strftime("%d/%b/%Y")

    PAHT_FILE = r'D:\A_PYTHON\ProgramasPython\Control_NodosCA\Reporte_CPU_VPN-PPS\ArchivoFinal_a_enviar\Subjet_Y_Cuerpo_Del_Correo.txt'
    CUERPO_CORREO = '\n\nBuenos días,\n\nSe adjunta reporte semanal de CPU para el OSS/VPN/PPS.\n\nSaludos.'
    SUBJECT = f'Carga de Cpu ::: {FECHA}'

    with open(PAHT_FILE, 'w') as file:
        file.write(SUBJECT)
        file.write(CUERPO_CORREO)


modificar_subject_txt()

print('Fin del programa...')

#--------------------------------------------------------------------------------------------------------------------
