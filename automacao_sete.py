### Importações para ler o arquivo excel



### Importação para automatizar a web
from selenium import webdriver # navegador
from selenium.webdriver.common.by import By # achar os elementos
from selenium.webdriver.common.keys import Keys # para digitar os elementos
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import pandas as pd
import openpyxl
import pyperclip
import time
import datetime



# Abrir o navegador
navegador = webdriver.Chrome()
espera = WebDriverWait(navegador, 10)

# Acessando o site
navegador.get("https://www.fnde.gov.br/sete/src/renderer/login-view.html")
# Colocando o navegador em tela cheia
navegador.maximize_window()
time.sleep(1)
# Acessando a pagina 

campo_usuario = navegador.find_element("id", "loginemail")
campo_usuario.send_keys("educacao@novaponte.mg.gov.br")

time.sleep(1)
campo_senha = navegador.find_element("id", "loginpassword")
campo_senha.send_keys("educa2025")
time.sleep(1)
botao_entrar = navegador.find_element("id", "loginsubmit")
#espera.until(EC.element_to_be_clickable(botao_entrar))
botao_entrar.click()    
time.sleep(2)


# selecionar campo alunos
campo_alunos = navegador.find_element(By.XPATH, '//*[@id="sidebar"]/div/ul/li[7]/a/p')
time.sleep(1)
campo_alunos.click()
time.sleep(1)

#selecionar cadastrar
cadastrar = navegador.find_element(By.XPATH, '//*[@id="alunoMenu"]/ul/li[1]/a')
time.sleep(1)
cadastrar.click()
time.sleep(1)
pyautogui.click(821,520, duration=1)
time.sleep(1)

# 0 - ler o arquivo excel()
workbook = openpyxl.load_workbook('teodoro.xlsx')
sheet_cadastro = workbook['Plan1']
nome_do_arquivo = "teodoro.xlsx"

# 1 - loopar arquivo excel
df = pd.read_excel(nome_do_arquivo)
for linha in sheet_cadastro.iter_rows(min_row=2):
    nome = linha[0].value
    pyperclip.copy(nome)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    cpf = linha[7].value
    time.sleep(1)
    pyperclip.copy(cpf)
    pyautogui.hotkey("tab")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    dataNascimento = linha[1].value
    dataNascimento = dataNascimento.strftime("%d/%m/%Y")
    pyperclip.copy(dataNascimento)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    nomeResponsavel = linha[8].value
    pyperclip.copy(nomeResponsavel)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.hotkey('tab')
    time.sleep(1)

    grauParentesco = linha[9].value
    pyperclip.copy(grauParentesco)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(10)




    sexo = linha[2].value
    cor = linha[3].value
    localizacao1 = linha[4].value
    nivelEnsino = linha[5].value
    turnoEnsino = linha[6].value
    endereco = linha[8].value
    time.sleep()    
    
#for index, row in df.iterrows():


#   1.1 - preeencher os dados liBeatriz Vitoria Ribeiro Bragados para cada linha do navegador
#campo_nome = navegador.find_element("id", "regnome")
#campo_senha.send_keys(nome)


time.sleep(10)
# clicar na localização do aluno
#localizacao = navegador.find_element(By.XPATH, '//*[@id="tab-dados-pessoais"]/div/div[3]/div/div/div[2]/div[2]/label')
#time.sleep(1)
#localizacao.click()
#time.sleep(1)
# Dar scroll
#pyautogui.click(1911,996, duration=1)

# Avancar a pagina
#pyautogui.click(1625,821, duration=1)
#time.sleep(3)
