from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from sys import exit

import pandas as pd
import time
import urllib

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get("https://web.whatsapp.com/")

def ler_excel(arquivo):
    print("buscando o arquivo excel...")
    time.sleep(2)
    try:
        return pd.read_excel(arquivo)
    except:
        print("Não foi possível ler o arquivo!")
        print("")
        exit()

def esperar_carregar_whatsaspp():
    while len(driver.find_elements_by_id("side")) < 1:
        time.sleep(1)


print("Aguardando leitura do QR Code...")
esperar_carregar_whatsaspp()

contatos = ler_excel("./contatos.xlsx")
print("Carregando contatos...")
mensagem = contatos["mensagem"][0]
time.sleep(2)



print("Enviando mensagens...")

print(str(len(contatos["telefone"])) + " contato(s) contrado(s)")
enviados = 0
total_erros = 0
contatos_erros = []


for i, contato in enumerate(contatos["telefone"]):
    try:
        nome = contatos.loc[i, "nome"]
        tel = contatos.loc[i, "telefone"].replace(" ", "")
        texto = urllib.parse.quote(mensagem.replace('{nome}', nome)) 
        link = f'https://web.whatsapp.com/send?phone={tel}&text={texto}'

        driver.get(link)
        
        esperar_carregar_whatsaspp()
        driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]').send_keys(Keys.ENTER)

        time.sleep(3)

        enviados += 1

    except:
        total_erros += 1
        contatos_erros.append([nome, tel, mensagem])

if(enviados > 0):
    print(f'{enviados} mensagem(ens) enviada(as)')

if(total_erros > 0):
    print(f"Erro ao enviar mensagem para {str(total_erros)} contato(s)")
    print("Gerando arquivo de erro...")

    df_erro_contatos = pd.DataFrame(contatos_erros, columns=['nome', 'telefone', 'mensagem'])

    df_erro_contatos.to_excel('erros.xlsx', sheet_name='Erros', header=True, index=False)

else:
    print("Nenhum erro ao enviar as mensagens!")

time.sleep(5)
print("Programa finalizado!")

