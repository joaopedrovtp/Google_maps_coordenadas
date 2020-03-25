'''

Scrip criado para busca de coordenadas geográficas no googlemaps
Planilha excel utilizada como fonte de dados - verificar arquivo auxiliares

'chromedriver.exe' deve estar na pasta Documentos

'''

import time  # gerencia o tempo do script
import os  # lida com o sistema operacional da máquina
from selenium import webdriver  # importa o Selenium para manipulação Web
from tkinter import Tk  # GUI para selecionar arquivo que eu desejo
from tkinter.filedialog import askopenfilename
from tkinter import messagebox as tkMessageBox
from openpyxl import load_workbook, workbook  # biblioteca que lida com o excel
import re

# escolha o arquivo em excel que você deseja trabalhar
root = Tk()  # Inicia uma GUI externa
excel_file = askopenfilename()  # Abre uma busca do arquivo que você deseja importar
root.destroy()

usuario_os = os.getlogin()

# declarando as variáveis a serem utilizadas dentro da automação Web
chromedriver = "/Users/" + usuario_os + "/Documents/chromedriver"  # local onde está o seu arquivo chromedriver
capabilities = {'chromeOptions': {'useAutomationExtension': False,
                                  'args': ['--disable-extensions']}}
driver = webdriver.Chrome(chromedriver, desired_capabilities=capabilities)
driver.implicitly_wait(30)

# iniciando a busca WEB
driver.maximize_window()  # maximiza a janela do chrome
driver.get("https://www.google.com/maps")  # acessa o googlemaps
time.sleep(5)

# importando o arquivo excel a ser utilizado
book = load_workbook(excel_file)  # abre o arquivo excel que será utilizado para cadastro
sheet = book["Coordenadas"]  # seleciona a sheet chamada "Coordenadas"
i = 2  # aqui indica começará da segunda linha do excel, ou seja, pulará o cabeçalho
for r in sheet.rows:

    endereco = sheet[i][1]
    munic_UF = sheet[i][3]

    if str(type(endereco.value)) == "<class 'NoneType'>":
       break

    endereco_completo = endereco.value + " " + munic_UF.value

    #preenche com o endereço completo e aperta o botão buscar
    driver.find_element_by_id("searchboxinput").send_keys(endereco_completo)
    driver.find_element_by_id("searchbox-searchbutton").click()

    # Aguarda carregar a URL e coleta os dados das coordenadas geográficas
    time.sleep(5)
    url = driver.current_url
    latlong = re.search('@(.+?)17z', url)
    if latlong:
        latlong = latlong.group(1).rsplit(",", 2)
        lat = latlong[0]
        long = latlong[1]
    sheet[i][6].value = lat.replace(".",",")
    sheet[i][7].value = long.replace(".",",")

    # Caso não localize o endereço, serã informado na célula do excel
    if lat == "":
        sheet[i][6].value= "Nao foi possivel identificar"

    # Limpa os campos para nova busca de coordenadas
    driver.find_element_by_id("searchboxinput").clear()
    lat=""
    long=""

    i += 1

# salva o excel na área de trabalho
caminho_arquivo = os.path.join("C:\ ".strip(), "Users", usuario_os, "Desktop", "Resultado_final.xlsx")
book.save(caminho_arquivo)

# Avisa sobre a finalização do robô e encerra o script
window = Tk()
window.wm_withdraw()
tkMessageBox.showinfo(title="Aviso", message="Script finalizado! Arquivo salvo na área de trabalho!")
window.destroy()
driver.close()
