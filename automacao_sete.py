# Importações necessárias (somente Selenium e bibliotecas relacionadas)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Abrir o navegador
navegador = webdriver.Chrome()
espera = WebDriverWait(navegador, 10)  # Espera até 10 segundos por elementos

# Acessando o site
navegador.get("https://www.fnde.gov.br/sete/src/renderer/login-view.html")
navegador.maximize_window()
time.sleep(1)  # Pequeno delay para garantir que a página carregue

# Login
campo_usuario = espera.until(EC.presence_of_element_located((By.ID, "loginemail")))
campo_usuario.send_keys("educacao@novaponte.mg.gov.br")

campo_senha = espera.until(EC.presence_of_element_located((By.ID, "loginpassword")))
campo_senha.send_keys("educa2025")

botao_entrar = espera.until(EC.element_to_be_clickable((By.ID, "loginsubmit")))
botao_entrar.click()
time.sleep(2)  # Espera para a página carregar após o login

# Navegar para o menu de alunos
campo_alunos = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidebar"]/div/ul/li[7]/a/p')))
time.sleep(2)
campo_alunos.click()
time.sleep(1)

# Selecionar "Cadastrar"
cadastrar = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="alunoMenu"]/ul/li[1]/a')))
time.sleep(1)
cadastrar.click()
time.sleep(1)

# Clique inicial para abrir o formulário (usando Selenium exclusivamente)
# Substitua '//*[@id="seu_xpath_aqui"' pelo XPath real do elemento que você clicava com PyAutoGUI
try:
    formulario_inicial = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="seu_xpath_aqui"')))
    formulario_inicial.click()
    time.sleep(2)  # Pequeno delay após o clique
except Exception as e:
    print(f"Erro ao clicar no elemento inicial: {str(e)}")
    # O script continuará, mas você pode adicionar lógica aqui se necessário (ex: parar ou relogar)

# Ler o arquivo Excel usando pandas
nome_do_arquivo = "teodoro.xlsx"
df = pd.read_excel(nome_do_arquivo)  # Lê o arquivo como DataFrame

# Loop pelas linhas do DataFrame
for index, row in df.iterrows():
    try:
        # Preencher os campos
        # 1. Nome
        campo_nome = espera.until(EC.presence_of_element_located((By.ID, "regnome")))  # ID do campo nome
        campo_nome.clear()
        campo_nome.send_keys(str(row[0]))  # Nome (coluna 0)

        # 2. CPF
        campo_cpf = espera.until(EC.presence_of_element_located((By.ID, "cpf_id")))  # Substitua pelo ID real do CPF
        campo_cpf.clear()
        campo_cpf.send_keys(str(row[7]))  # CPF (coluna 7)

        # 3. Data de Nascimento
        campo_data_nascimento = espera.until(EC.presence_of_element_located((By.ID, "data_nascimento_id")))  # Substitua pelo ID real
        campo_data_nascimento.clear()
        data_nascimento_formatada = row[1].strftime("%d/%m/%Y")  # Formata a data (coluna 1)
        campo_data_nascimento.send_keys(data_nascimento_formatada)

        # 4. Nome do Responsável
        campo_nome_responsavel = espera.until(EC.presence_of_element_located((By.ID, "nome_responsavel_id")))  # Substitua pelo ID real
        campo_nome_responsavel.clear()
        campo_nome_responsavel.send_keys(str(row[8]))  # Nome do Responsável (coluna 8)

        # 5. Grau de Parentesco
        campo_grau_parentesco = espera.until(EC.presence_of_element_located((By.ID, "grau_parentesco_id")))  # Substitua pelo ID real
        campo_grau_parentesco.clear()
        campo_grau_parentesco.send_keys(str(row[9]))  # Grau de Parentesco (coluna 9)

        # 6. Sexo (assumindo que é um radio button ou dropdown)
        sexo = str(row[2]).upper()  # Sexo (coluna 2), ex: 'M' ou 'F'
        if sexo == 'M':
            campo_sexo_masculino = espera.until(EC.element_to_be_clickable((By.ID, "sexo_masculino_id")))  # Substitua pelo ID real
            campo_sexo_masculino.click()
        elif sexo == 'F':
            campo_sexo_feminino = espera.until(EC.element_to_be_clickable((By.ID, "sexo_feminino_id")))  # Substitua pelo ID real
            campo_sexo_feminino.click()

        # 7. Cor/Raça (assumindo dropdown ou seleção)
        cor = str(row[3])  # Cor (coluna 3)
        campo_cor = espera.until(EC.presence_of_element_located((By.ID, "cor_id")))  # Substitua pelo ID real
        campo_cor.clear()
        campo_cor.send_keys(cor)  # Ou use .click() se for uma seleção

        # 8. Localização (assumindo radio button ou dropdown)
        localizacao = str(row[4])  # Localização (coluna 4)
        campo_localizacao = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tab-dados-pessoais"]/div/div[3]/div/div/div[2]/div[2]/label')))  # Como no seu código
        campo_localizacao.click()  # Clique na localização (ajuste se necessário)

        # 9. Nível de Ensino (assumindo dropdown)
        nivel_ensino = str(row[5])  # Nível de Ensino (coluna 5)
        campo_nivel_ensino = espera.until(EC.presence_of_element_located((By.ID, "nivel_ensino_id")))
        campo_nivel_ensino.clear()
        campo_nivel_ensino.send_keys(nivel_ensino)

        # 10. Turno de Ensino (assumindo dropdown)
        turno_ensino = str(row[6])  # Turno de Ensino (coluna 6)
        campo_turno_ensino = espera.until(EC.presence_of_element_located((By.ID, "turno_ensino_id")))
        campo_turno_ensino.clear()
        campo_turno_ensino.send_keys(turno_ensino)

        # Avançar/Salvar o formulário
        botao_avancar = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="seu_xpath_botao_avancar"]')))  # Substitua pelo XPath real
        botao_avancar.click()
        time.sleep(2)  # Pequeno delay para a página processar

        print(f"Registro para {row[0]} enviado com sucesso!")

    except Exception as e:
        print(f"Erro ao processar a linha {index}: {str(e)}")
        continue  # Pula para a próxima linha se der erro

# Fechar o navegador após o loop
navegador.quit()
print("Processo concluído!")

