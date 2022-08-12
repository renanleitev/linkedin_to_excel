import contextlib
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys # noqa
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import pyperclip as pyper
import pyautogui as pg
from time import sleep


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), chrome_options=options) # noqa


class Driver():
    """ Para impedir de fechar automaticamente """
    def iniciar(self):
        global driver

    """ Para abrir qualquer página """
    def abrir(self, site):
        driver.get(site)

    """ Para logar no sistema """
    def login(self):
        try:
            entrar = driver.find_element(By.XPATH, '/html/body/nav/div/a[2]')
            entrar.click()
            email = driver.find_element(By.XPATH, '//*[@id="username"]')
            email.click()
            """ ATENÇÃO: INSIRA O SEU "E-MAIL" (ENTRE ASPAS)
            NO CAMPO ENTRE PARÊNTESES = send_keys('e-mail')"""
            email.send_keys('seu_email@gmail.com')
            sleep(1)
            senha = driver.find_element(By.XPATH, '//*[@id="password"]')
            senha.click()
            """ ATENÇÃO: INSIRA A SUA "SENHA" (ENTRE ASPAS)
            NO CAMPO ENTRE PARÊNTESES = send_keys('senha') """
            senha.send_keys('sua_senha')
            sleep(1)
            btn_entrar = driver.find_element(By.XPATH, '//*[@id="organic-div"]/form/div[3]/button') # noqa
            btn_entrar.click()
            sleep(3)
            perfil = driver.find_element(By.XPATH, '/html/body/div[5]/div[3]/div/div/div[2]/div/div/div/div[1]/div[1]/a/div[1]/img') # noqa
            perfil.click()
            sleep(2)
            """ ATENÇÃO: INSIRA A QUANTIDADE DE CERTIFICADOS (SEM ASPAS)
            NO CAMPO NUMERO_DE_CERTIFICADOS"""
            certificado = driver.find_element(By.LINK_TEXT, 'Exibir todas as NUMERO_DE_CERTIFICADOS licenças e certificados') # noqa
            certificado.click()
            SCROLL_PAUSE_TIME = 1
            # Get scroll height
            last_height = driver.execute_script("return document.body.scrollHeight") # noqa
            while True:
                # Scroll down to bottom
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") # noqa
                # Wait to load page
                sleep(SCROLL_PAUSE_TIME)
                # Get new scroll height and compare with last scroll height
                new_height = driver.execute_script("return document.body.scrollHeight") # noqa
                if new_height == last_height:
                    break
                last_height = new_height
            sleep(0.5)
            pg.hotkey('ctrl', 'a')
            sleep(0.5)
            pg.hotkey('ctrl', 'c')
            driver.close()
            texto = pyper.paste()
        except Exception as e:
            print('Erro ao logar:', e)
        try:
            with open('testeselenium.txt', 'a', encoding='utf-8') as file:
                file.write(texto)
        except Exception as e:
            print('Erro:', e)


def organiza():
    lista = []
    with open('testeselenium.txt', 'r', encoding='utf-8') as file:
        for palavra in file:
            nova_palavra = palavra.strip()
            if not nova_palavra.isspace() and nova_palavra != '' and 'Exibir' not in nova_palavra and 'Logo' not in nova_palavra: # noqa
                lista.append(nova_palavra)
    lista_nova = lista[lista.index('Licenças e certificados')+1:lista.index('Editar perfil público e URL')] # noqa
    lista_final = []
    for c in lista_nova:
        x = len(c)
        final = slice(0, x // 2)
        lista_final.append(c[final])
    lista_temp = []
    lista_org = []
    lista_iter = iter(lista_final)
    with contextlib.suppress(Exception):
        for c in lista_final:
            if len(lista_temp) == 3:
                iterador = next(lista_iter)
                if 'Código' not in iterador:
                    lista_temp.append('Sem código')
                    lista_org.append(lista_temp)
                    lista_temp = [iterador]
                else:
                    lista_temp.append(iterador)
                    lista_temp = []
            else:
                lista_temp.append(next(lista_iter))
    lista_iter = iter(lista_final)
    lista_temp = []
    with contextlib.suppress(Exception):
        while lista_iter is not None:
            iterador = next(lista_iter)
            lista_temp.append(iterador)
            if 'Código' in iterador:
                lista_org.append(lista_temp[lista_temp.index(iterador)-3:lista_temp.index(iterador)+1]) # noqa
    book = Workbook()
    sheet = book.active
    sheet["A1"] = "Certificado"
    sheet["B1"] = "Instituição"
    sheet["C1"] = "Data de Emissão"
    sheet["D1"] = "Código da Credencial"
    conta = 2
    with contextlib.suppress(Exception):
        for x in lista_org:
            sheet[f"A{conta}"] = x[0]
            sheet[f"B{conta}"] = x[1]
            sheet[f"C{conta}"] = x[2]
            sheet[f"D{conta}"] = x[3]
            conta += 1
    book.save("lista_certificados.xlsx")
    print('Excel gerado com sucesso.')


if __name__ == "__main__":
    nav = Driver()
    nav.abrir("https://br.linkedin.com/")
    nav.login()
    organiza()
